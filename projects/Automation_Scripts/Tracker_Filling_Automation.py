"""
Populate Master Tracker (Pre-Joining) from the Conventional Utilization
Report + individual candidate "Verification" files.

WORKFLOW
--------
1. Read "NEW UPLOAD" sheet of Conventional_Utilization_Report.xlsx -
   these are the candidates to add.
2. For each candidate:
   a. Skip if their ID already exists anywhere in the Master Tracker
      (Pre-Joining / PE & CVC / Fresher / H-List) - print a warning,
      don't create a duplicate row.
   b. Find their individual "<Name>_<ID>.xlsx" file and parse its
      "Verification" sheet into named blocks (EDUCATION, EMPLOYMENT
      CURRENT, EMPLOYMENT 2-6, CRIMINAL, GDC).
   c. For each check type, ONLY use a block's data if that check is
      listed in "Checks to be Initiate" AND the block's own
      "APPLICANT ID" cell matches this candidate's ID. If the ID
      doesn't match (data mixed up between candidates) or the block
      is missing, that check's columns are set to "NA" and a warning
      is printed instead of writing wrong data.
   d. Build the new Pre-Joining row and append it below the sheet's
      existing data - existing rows/sheets/formatting are untouched.
3. Write the result to a NEW file (not overwriting the original), so
   you can review before replacing the live tracker.

Run this from the same folder as the three input files.
"""

import argparse
import glob
import logging
import os
import sys
from datetime import datetime

import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side

# ===========================
# CONFIG - filenames (override via CLI args, see --help)
# ===========================

RUN_TS = datetime.now().strftime("%Y%m%d_%H%M%S")

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument("--input-folder", default="New_Cases",
                     help="Folder containing the Conventional Utilization Report, "
                          "the Master Tracker, and the individual candidate files")
parser.add_argument("--report-file", default=None,
                     help="Path to the Conventional Utilization Report "
                          "(default: <input-folder>/Conventional Utilization Report.xlsx)")
parser.add_argument("--master-tracker-file", default=None,
                     help="Path to the live Master Tracker, read-only, never modified "
                          "(default: <input-folder>/Infy_Conventional-Master Tracker.xlsx)")
parser.add_argument("--output-file", default=None,
                     help="Output path. Default: timestamped, in the current folder.")
parser.add_argument("--log-file", default=None,
                     help="Log file path. Default: timestamped, in the current folder.")
parser.add_argument("--repair-incomplete", action="store_true",
                     help="If an ID already exists in the Master Tracker but its check-detail "
                          "columns are completely blank (an old incomplete row), backfill that "
                          "same row from the reference file instead of skipping it as a duplicate. "
                          "Rows that already have any check data filled in are never touched.")
args = parser.parse_args()

INPUT_FOLDER = args.input_folder
REPORT_FILE = args.report_file or os.path.join(INPUT_FOLDER, "Conventional Utilization Report.xlsx")
MASTER_TRACKER_FILE = args.master_tracker_file or os.path.join(INPUT_FOLDER, "Infy_Conventional-Master Tracker.xlsx")
OUTPUT_FILE = args.output_file or f"Infy_Conventional-Master Tracker_Updated_{RUN_TS}.xlsx"
LOG_FILE = args.log_file or f"tracker_run_{RUN_TS}.log"

MASTER_TRACKER_SHEET = "Pre-Joining"

# Which sheets to search for EXISTING ids (duplicate check)
MT_ID_CHECK_SHEETS = ["Pre-Joining", "PE & CVC", "Fresher", "H-List"]

# Columns the NEW UPLOAD sheet must have - checked up front so a typo'd
# / renamed column fails fast with a clear message instead of crashing
# mid-batch on some arbitrary candidate.
REQUIRED_NEW_UPLOAD_COLUMNS = [
    "applicantId", "candidateName", "conventionalCheckNames",
    "candidateCode", "contactNumber", "emailId", "createdOn",
]

# ===========================
# LOGGING - console + timestamped file, so the run has a permanent audit trail
# ===========================

logger = logging.getLogger("tracker_fill")
logger.setLevel(logging.INFO)
_fmt = logging.Formatter("%(asctime)s | %(levelname)-7s | %(message)s", "%Y-%m-%d %H:%M:%S")

_console = logging.StreamHandler(sys.stdout)
_console.setFormatter(_fmt)
logger.addHandler(_console)

_file_handler = logging.FileHandler(LOG_FILE, encoding="utf-8")
_file_handler.setFormatter(_fmt)
logger.addHandler(_file_handler)

logger.info("Run started. Log file: %s", LOG_FILE)

# ===========================
# INPUT FILE CHECKS - fail fast with a clear message, not a raw traceback
# ===========================

_missing = [p for p in (REPORT_FILE, MASTER_TRACKER_FILE) if not os.path.isfile(p)]
if _missing:
    for p in _missing:
        logger.error("Input file not found: %s", os.path.abspath(p))
    logger.error("Aborting - fix the path(s) above (or pass --input-folder / --report-file / --master-tracker-file) and re-run.")
    sys.exit(1)

# Fail fast if the Master Tracker doesn't have the sheet we write to -
# otherwise this would only surface AFTER processing every candidate.
_mt_sheet_names = openpyxl.load_workbook(MASTER_TRACKER_FILE, read_only=True).sheetnames
if MASTER_TRACKER_SHEET not in _mt_sheet_names:
    logger.error(
        "Master Tracker has no '%s' sheet. Found sheets: %s",
        MASTER_TRACKER_SHEET, _mt_sheet_names,
    )
    logger.error("Aborting before processing any candidate.")
    sys.exit(1)

# ===========================
# HELPERS
# ===========================


def clean_id(value):
    """
    Strips whitespace INCLUDING non-breaking spaces (\\xa0), which the
    Conventional Utilization Report's applicantId column has trailing
    on some rows. Also drops a trailing ".0" from float-read IDs.
    """
    text = str(value).replace("\xa0", " ").strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text


def find_candidate_file(candidate_name, candidate_id, folder="."):
    """
    Finds the candidate's reference xlsx file, supporting two layouts:
      1. FLAT: "<Name>_<ID>.xlsx" sitting directly in `folder`.
      2. PER-CANDIDATE SUBFOLDER: a folder named "<Name>_<ID>" (or
         just containing the ID) holding the xlsx alongside other
         files (e.g. a zipped proof-documents attachment) - the xlsx
         inside that subfolder is picked, preferring one whose own
         filename also contains the ID.
    Falls back to any .xlsx directly under `folder` whose name
    contains the ID, in case naming doesn't line up exactly.
    Excel's temporary lock files ("~$...") are always skipped.
    """
    name_variants = [
        candidate_name.strip().replace(" ", "_"),
        candidate_name.strip(),
    ]

    def is_real_xlsx(path):
        return not os.path.basename(path).startswith("~$")

    # 1. Flat layout - file directly in folder
    for variant in name_variants:
        candidate_path = os.path.join(folder, f"{variant}_{candidate_id}.xlsx")
        if os.path.exists(candidate_path) and is_real_xlsx(candidate_path):
            return candidate_path

    # 2. Per-candidate subfolder layout
    subfolder_candidates = [
        os.path.join(folder, f"{variant}_{candidate_id}")
        for variant in name_variants
    ]
    subfolder_candidates += glob.glob(os.path.join(folder, f"*{candidate_id}*"))

    for sub in subfolder_candidates:
        if not os.path.isdir(sub):
            continue

        xlsx_files = [
            f for f in glob.glob(os.path.join(sub, "*.xlsx")) if is_real_xlsx(f)
        ]
        if not xlsx_files:
            continue

        # Prefer a file whose own name also contains the ID (in case
        # the subfolder holds more than one xlsx).
        for f in xlsx_files:
            if candidate_id in os.path.basename(f):
                return f

        return xlsx_files[0]

    # 3. Last resort - any .xlsx directly under folder containing the ID
    matches = [
        f for f in glob.glob(os.path.join(folder, f"*{candidate_id}*.xlsx"))
        if is_real_xlsx(f)
    ]
    if matches:
        return matches[0]

    return None


def parse_verification_blocks(filepath):
    """
    The "Verification" sheet is laid out as repeated blocks:
        <BLOCK TITLE>                (col A only, rest blank)
        <header row>                 (column names)
        <data row(s)>
        <blank row>
    Returns {BLOCK_TITLE_UPPER: DataFrame}. Robust to blocks appearing
    in any order / any subset being present for a given candidate.
    """

    raw = pd.read_excel(filepath, sheet_name="Verification", header=None)

    blocks = {}
    i = 0
    n = len(raw)

    while i < n:
        row = raw.iloc[i]
        first_cell = row.iloc[0]
        rest_blank = row.iloc[1:].isna().all()

        if pd.notna(first_cell) and rest_blank:
            title = str(first_cell).strip().upper()
            header_idx = i + 1

            if header_idx >= n:
                break

            headers = raw.iloc[header_idx].tolist()
            data_rows = []
            j = header_idx + 1

            while j < n:
                r = raw.iloc[j]
                if r.isna().all():
                    break
                # A new block title row also ends this block's data.
                if pd.notna(r.iloc[0]) and r.iloc[1:].isna().all():
                    break
                data_rows.append(r.tolist())
                j += 1

            blocks[title] = pd.DataFrame(data_rows, columns=headers)
            i = j
        else:
            i += 1

    return blocks


def get_block_row(blocks, title_options, target_id):
    """
    Returns the first data row (as a dict) of whichever block title in
    title_options exists AND whose own "APPLICANT ID" matches
    target_id. Returns None if no such valid block is found (missing,
    or ID mismatch - both treated the same: "don't trust this data").
    """

    for title in title_options:
        block = blocks.get(title.upper())
        if block is None or len(block) == 0:
            continue

        block_id = None
        for col in block.columns:
            if str(col).strip().upper() == "APPLICANT ID":
                block_id = clean_id(block.iloc[0][col])
                break

        if block_id == target_id:
            return block.iloc[0].to_dict()

    return None


def block_val(block_row, col_name, default="NA"):
    if block_row is None:
        return default
    val = block_row.get(col_name)
    if pd.isna(val):
        return default
    return val


def clean_text(value, default="NA"):
    """Cleans a pass-through value from the report: strips whitespace
    (including non-breaking spaces), and substitutes `default` for
    anything missing/NaN."""
    if pd.isna(value):
        return default
    text = str(value).replace("\xa0", " ").strip()
    return text if text else default


def parse_checks_list(value):
    if pd.isna(value):
        return []
    return [t.strip().upper() for t in str(value).split(",") if t.strip()]


# Master Tracker wants common display names, not the raw report codes.
# CURR_EMP and every EMP entry both collapse to "Employment".
CHECKS_DISPLAY_MAP = {
    "EDU": "Education",
    "CURR_EMP": "Employment",
    "EMP": "Employment",
    "CRIM": "CRC",
    "CRC": "CRC",
    "GDC": "GDC",
}


def format_checks_display(checks_list):
    """
    Maps raw check codes (EDU, CURR_EMP, EMP, CRIM, GDC) to the Master
    Tracker's common display names, de-duplicated, first-seen order,
    joined the "natural language" way - commas between items, and an
    "&" (no comma before it) between the last two:
        1 item:  "Education"
        2 items: "Education & Employment"
        3+:      "Education, Employment, CRC & GDC"
    """
    seen = []
    for code in checks_list:
        display = CHECKS_DISPLAY_MAP.get(code, code)
        if display not in seen:
            seen.append(display)

    if not seen:
        return "NA"
    if len(seen) == 1:
        return seen[0]
    return ", ".join(seen[:-1]) + " & " + seen[-1]


def format_date_ddmmmyy(value):
    """
    Normalizes any parseable date/datetime (including ones with a time
    component) to "DD-MMM-YY", e.g. 05-Sep-26. Returns "NA" if the
    value is missing or unparseable.
    NOTE: dayfirst=True is assumed for ambiguous numeric dates
    (e.g. 03/04/26 -> 3 Apr 2026). Verify this matches the report's
    actual convention on real data.
    """
    if value is None:
        return "NA"
    if isinstance(value, str) and value.strip().upper() in ("", "NA", "NAN", "PENDING"):
        return value.strip() if value.strip().upper() == "PENDING" else "NA"
    try:
        if pd.isna(value):
            return "NA"
    except (TypeError, ValueError):
        pass
    dt = pd.to_datetime(value, dayfirst=True, errors="coerce")
    if pd.isna(dt):
        return "NA"
    return dt.strftime("%d-%b-%y")


def formatted_block_date(block_row, col_name):
    """Like block_val, but runs the result through format_date_ddmmmyy."""
    raw = block_val(block_row, col_name, default=None)
    if raw is None:
        return "NA"
    return format_date_ddmmmyy(raw)


# Styling applied to every cell we paste into the Master Tracker.
CELL_FONT = Font(name="Calibri", size=11)
CELL_ALIGN = Alignment(horizontal="center", vertical="center")
_THIN = Side(style="thin")
CELL_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


def style_cell(cell):
    cell.font = CELL_FONT
    cell.alignment = CELL_ALIGN
    cell.border = CELL_BORDER


# ===========================
# LOAD EXISTING MASTER TRACKER IDS (for duplicate check)
# ===========================

existing_ids = set()

for sheet in MT_ID_CHECK_SHEETS:
    try:
        df = pd.read_excel(MASTER_TRACKER_FILE, sheet_name=sheet, keep_default_na=False)
    except Exception:
        continue

    for col in ["C_ID", "Emp_No"]:
        if col in df.columns:
            existing_ids.update(
                clean_id(v) for v in df[col] if str(v).strip() not in ("", "NA", "nan")
            )

logger.info("Loaded %d existing IDs from Master Tracker for duplicate check.", len(existing_ids))

# ===========================
# LOAD NEW UPLOAD CANDIDATES
# ===========================

new_upload = pd.read_excel(REPORT_FILE, sheet_name="NEW UPLOAD")

_missing_cols = [c for c in REQUIRED_NEW_UPLOAD_COLUMNS if c not in new_upload.columns]
if _missing_cols:
    logger.error(
        "NEW UPLOAD sheet is missing required column(s): %s", _missing_cols
    )
    logger.error("Found columns: %s", list(new_upload.columns))
    logger.error("Aborting before processing any candidate - fix the report's columns and re-run.")
    sys.exit(1)

logger.info("NEW UPLOAD sheet schema OK - %d candidate row(s) found.", len(new_upload))

new_rows = []
skipped_duplicates = []
skipped_no_file = []
skip_notes = []

skipped_errors = []

# ===========================
# LOAD WORKBOOK FOR WRITING (loaded here, before the loop, so repair
# mode can check existing rows while processing each candidate)
# ===========================

wb = openpyxl.load_workbook(MASTER_TRACKER_FILE)
ws = wb[MASTER_TRACKER_SHEET]

header_cells = ws[1]

header_map = {}
seen_counts = {}

for cell in header_cells:
    if cell.value is None:
        continue

    name = str(cell.value).strip()

    if name in seen_counts:
        seen_counts[name] += 1
        # Replicate pandas' own duplicate-column naming convention
        # (Vendor, Vendor.1, Vendor.2, ...) since the raw sheet has
        # several columns literally named "Vendor".
        mapped_name = f"{name}.{seen_counts[name]}"
    else:
        seen_counts[name] = 0
        mapped_name = name

    header_map[mapped_name] = cell.column

# ws.max_row can overcount if trailing rows carry formatting/merged
# cells but no real data (Excel keeps them "in use") - that phantom
# row is what was causing a blank row to appear above the pasted
# data. Instead, walk up from the bottom on the ID column and find
# the last row that actually has a value. Also build an id -> row
# lookup here, needed by --repair-incomplete below.
id_col_for_scan = header_map.get("C_ID", 1)
last_data_row = 1
id_to_row = {}
for r in range(2, ws.max_row + 1):
    cell_val = ws.cell(row=r, column=id_col_for_scan).value
    if cell_val not in (None, ""):
        last_data_row = r
        id_to_row[clean_id(cell_val)] = r
next_row = last_data_row + 1

next_sno = 1
sno_col = header_map.get("S.No")
if sno_col:
    existing_snos = [
        ws.cell(row=r, column=sno_col).value
        for r in range(2, last_data_row + 1)
    ]
    existing_snos = [v for v in existing_snos if isinstance(v, (int, float))]
    if existing_snos:
        next_sno = int(max(existing_snos)) + 1

# Anchor columns used to decide if an existing row was left completely
# blank by a previous (buggy/incomplete) run - repair mode only ever
# touches a row when EVERY one of these is empty, so partially-worked
# rows are never overwritten.
_INCOMPLETE_ANCHOR_COLUMNS = ["Ed_SDate", "CE1_Sdate", "eC_CRC_SDate", "eP_CRC_SDate", "GDB_SDate"]


def is_row_incomplete(row_num):
    checked_any = False
    for col_name in _INCOMPLETE_ANCHOR_COLUMNS:
        col_idx = header_map.get(col_name)
        if col_idx is None:
            continue
        checked_any = True
        val = ws.cell(row=row_num, column=col_idx).value
        if val not in (None, "", "NA"):
            return False
    return checked_any


repaired_rows = []  # list of (row_num, row_dict) for --repair-incomplete

for row_num, cand in new_upload.iterrows():
    raw_id_for_log = cand.get("applicantId", "<missing>")
    raw_name_for_log = cand.get("candidateName", "<missing>")
    try:

        target_id = clean_id(cand["applicantId"])
        candidate_name = str(cand["candidateName"]).strip()

        repair_row_num = None
        if target_id in existing_ids:
            if args.repair_incomplete and target_id in id_to_row and is_row_incomplete(id_to_row[target_id]):
                repair_row_num = id_to_row[target_id]
                logger.info(
                    "INCOMPLETE ROW FOUND: ID %s (%s) is in Master Tracker but has no check "
                    "data filled - will backfill row %d instead of skipping.",
                    target_id, candidate_name, repair_row_num,
                )
            else:
                skipped_duplicates.append(target_id)
                logger.info("DUPLICATE: ID %s (%s) already in Master Tracker - skipped.", target_id, candidate_name)
                continue

        checks_list = parse_checks_list(cand.get("conventionalCheckNames"))
        edu_needed = "EDU" in checks_list
        # Report data seen so far uses "CRC" as the code for this check,
        # but the display-mapping was originally written for "CRIM" - so
        # both are accepted here to avoid a silent skip either way.
        crim_needed = "CRIM" in checks_list or "CRC" in checks_list
        gdc_needed = "GDC" in checks_list
        emp_count = checks_list.count("EMP")

        # Some candidates get an explicit "CURR_EMP" code for the current
        # employer. Others only ever get a run of "EMP" codes with no
        # "CURR_EMP" at all - in that case the FIRST "EMP" represents the
        # current employer (its data lives in the "EMPLOYMENT 1" block,
        # already an accepted CE1 title below) and only the REMAINING EMP
        # count should map to the numbered E2-E6 blocks. Without this,
        # CE1 silently never gets filled (curr_emp_needed stayed False)
        # AND the numbered blocks end up shifted by one, both at once.
        has_curr_emp_code = "CURR_EMP" in checks_list
        curr_emp_needed = has_curr_emp_code or emp_count >= 1
        emp_count_for_numbered = emp_count if has_curr_emp_code else max(emp_count - 1, 0)

        candidate_file = find_candidate_file(candidate_name, target_id, folder=INPUT_FOLDER)

        if candidate_file is None:
            skipped_no_file.append({"ApplicantId": target_id, "CandidateName": candidate_name})
            logger.warning(
                "REFERENCE DOC NOT THERE: %s_%s - not added to Master Tracker "
                "(no individual reference file found in %s).",
                candidate_name, target_id, INPUT_FOLDER,
            )
            continue

        blocks = parse_verification_blocks(candidate_file)

        row = {}

        # ---- Basic identity fields ----
        row["C_ID"] = target_id
        row["DV_Ref ID"] = clean_text(cand.get("candidateCode"))
        row["Emp_No"] = "NA"
        row["C_Name"] = candidate_name
        row["Mob "] = cand.get("contactNumber") if pd.notna(cand.get("contactNumber")) else "NA"
        row["Email id"] = clean_text(cand.get("emailId"))
        row["Package Name"] = "Pre-Joining"
        row["DOJ"] = "NA"
        row["Location"] = "NA"
        row["Checks to be Initiate"] = format_checks_display(checks_list)
        row["Case Rcvd Date "] = format_date_ddmmmyy(cand.get("createdOn"))

        row["Edu_SPOC"] = "NA"
        row["Edu-Insuff Raised Date"] = "NA"
        row["Edu-Insuff"] = "NA"
        row["Edu-Insuff Cleared Date"] = "NA"
        row["Ed_SDate"] = "Pending"
        row["Education Type"] = "HIGHEST"

        # ---- EDUCATION ----
        edu_row = get_block_row(
            blocks, ["EDUCATION CHECK HIGHEST EDUCATION"], target_id
        ) if edu_needed else None

        if edu_needed and edu_row is None:
            skip_notes.append(f"{target_id} ({candidate_name}): EDU needed but block missing/ID-mismatch - set NA")

        if edu_row is not None:
            row["Qualification Attained"] = block_val(edu_row, "QUALIFICATION ATTAINED")
            row["Ed_College Name"] = block_val(edu_row, "NAME OF THE EDUCATIONAL INSTITUTION")
            row["Ed_YOP"] = block_val(edu_row, "YEAR OF PASSING")
            row["Ed_INT_SRC 1"] = "Pending"
            row["Ed_INT_SRC 2"] = "Pending"
            row["Ed_V_Source"] = "Pending"
            row["V_Desgn"] = "Pending"
            row["Ed_Result"] = "Pending"
            row["Ed_Edate"] = "Pending"
            row["Vendor"] = "Pending"
        else:
            row["Qualification Attained"] = "NA"
            row["Ed_College Name"] = "NA"
            row["Ed_YOP"] = "NA"
            row["Ed_INT_SRC 1"] = "NA"
            row["Ed_INT_SRC 2"] = "NA"
            row["Ed_V_Source"] = "NA"
            row["V_Desgn"] = "NA"
            row["Ed_Result"] = "NA"
            row["Ed_Edate"] = "NA"
            row["Vendor"] = "NA"

        row["Ed_Univ"] = "NA"
        row["\nEd_Regn"] = "NA"

        # ---- EMPLOYMENT (current) ----
        row["CE_Status"] = "YES" if curr_emp_needed else "NO"

        ce1_row = get_block_row(
            blocks,
            ["EMPLOYMENT CHECK CURRENT EMPLOYER", "EMPLOYMENT 1"],
            target_id,
        ) if curr_emp_needed else None

        if curr_emp_needed and ce1_row is None:
            skip_notes.append(f"{target_id} ({candidate_name}): CURR_EMP needed but block missing/ID-mismatch - set NA")

        if ce1_row is not None:
            row["CE1_Name"] = block_val(ce1_row, "EMPLOYERS NAME")
            row["CE1_DOJ"] = formatted_block_date(ce1_row, "DOJ")
            row["CE1_LWD"] = formatted_block_date(ce1_row, "DOE")
            row["CE1_C_Desgn"] = "NA"
            row["CE1_Sdate"] = "Pending"
            row["CE1_Source"] = "Pending"
            row["CE1_V_Desgn"] = "Pending"
            row["CE1_Result"] = "Pending"
            row["CE1_Edate"] = "Pending"
        else:
            row["CE1_Name"] = "NA"
            row["CE1_DOJ"] = "NA"
            row["CE1_LWD"] = "NA"
            row["CE1_C_Desgn"] = "NA"
            row["CE1_Sdate"] = "NA"
            row["CE1_Source"] = "NA"
            row["CE1_V_Desgn"] = "NA"
            row["CE1_Result"] = "NA"
            row["CE1_Edate"] = "NA"

        # ---- Previous employments E2-E6 ----
        # NOTE: no sample "EMPLOYMENT 2/3/4/5" block was available to
        # confirm its exact column names - this assumes the SAME column
        # names as "EMPLOYMENT CHECK CURRENT EMPLOYER". Verify against a
        # real multi-employment candidate file and adjust
        # EMP_BLOCK_TITLE_OPTIONS / column names below if needed.
        for idx in range(2, 7):
            needed_i = emp_count >= (idx - 1)
            title_options = [f"EMPLOYMENT {idx}", f"EMPLOYMENT CHECK {idx}"]

            emp_row = get_block_row(blocks, title_options, target_id) if needed_i else None

            if needed_i and emp_row is None:
                skip_notes.append(
                    f"{target_id} ({candidate_name}): EMPLOYMENT {idx} needed but block missing/ID-mismatch - set NA"
                )

            prefix = f"E{idx}"

            if emp_row is not None:
                row[f"{prefix}_Name"] = block_val(emp_row, "EMPLOYERS NAME")
                row[f"{prefix}_DOJ"] = formatted_block_date(emp_row, "DOJ")
                row[f"{prefix}_LWD"] = formatted_block_date(emp_row, "DOE")
                row[f"{prefix}_C_Desgn"] = "NA"
                row[f"{prefix}_Sdate"] = "Pending"
                row[f"{prefix}_Source"] = "Pending"
                row[f"{prefix}_V_Desgn"] = "Pending"
                row[f"{prefix}_Result"] = "Pending"
                row[f"{prefix}_EDate"] = "Pending"
            else:
                row[f"{prefix}_Name"] = "NA"
                row[f"{prefix}_DOJ"] = "NA"
                row[f"{prefix}_LWD"] = "NA"
                row[f"{prefix}_C_Desgn"] = "NA"
                row[f"{prefix}_Sdate"] = "NA"
                row[f"{prefix}_Source"] = "NA"
                row[f"{prefix}_V_Desgn"] = "NA"
                row[f"{prefix}_Result"] = "NA"
                row[f"{prefix}_EDate"] = "NA"

        # ---- CRIMINAL (permanent address only - current address always NA) ----
        row["eC_CRC_SDate"] = "NA"
        row["eCurrent_CRC_Add"] = "NA"
        row["eC_CRC_Result"] = "NA"
        row["eC_CRC_EDate"] = "NA"
        row["Vendor.1"] = "NA"

        crim_row = get_block_row(
            blocks, ["CRIMINAL CHECK PERMANENT ADDRESS"], target_id
        ) if crim_needed else None

        if crim_needed and crim_row is None:
            skip_notes.append(f"{target_id} ({candidate_name}): CRIM needed but block missing/ID-mismatch - set NA")

        if crim_row is not None:
            row["eP_CRC_SDate"] = "Pending"
            row["ePermanent_CRC_Add"] = block_val(crim_row, "ADDRESS")
            row["eP_CRC_Result"] = "Pending"
            row["eP_CRC_EDate"] = "Pending"
            row["Vendor.2"] = "Pending"
        else:
            row["eP_CRC_SDate"] = "NA"
            row["ePermanent_CRC_Add"] = "NA"
            row["eP_CRC_Result"] = "NA"
            row["eP_CRC_EDate"] = "NA"
            row["Vendor.2"] = "NA"

        # ---- GDC ----
        gdc_row = get_block_row(
            blocks, ["GLOBAL DATABASE CHECK"], target_id
        ) if gdc_needed else None

        if gdc_needed and gdc_row is None:
            skip_notes.append(f"{target_id} ({candidate_name}): GDC needed but block missing/ID-mismatch - set NA")

        if gdc_row is not None:
            row["DOB"] = formatted_block_date(gdc_row, "DATEOFBIRTH")
            row["Father name"] = block_val(gdc_row, "FATHERNAME")
            row["Gender"] = block_val(gdc_row, "GENDER")
            row["GDB_SDate"] = "Pending"
            row["GDB_Add"] = block_val(gdc_row, "ADDRESS")
            row["GDB_Result"] = "Pending"
            row["GDB_EDate"] = "Pending"
            row["Vendor.3"] = "Pending"
        else:
            row["DOB"] = "NA"
            row["Father name"] = "NA"
            row["Gender"] = "NA"
            row["GDB_SDate"] = "NA"
            row["GDB_Add"] = "NA"
            row["GDB_Result"] = "NA"
            row["GDB_EDate"] = "NA"
            row["Vendor.3"] = "NA"

        row["ReINT_GDB_SDate"] = "NA"
        row["ReINT_GDB_Result"] = "NA"
        row["ReINT_GDB_EDate"] = "NA"
        row["Vendor.4"] = "NA"

        if repair_row_num:
            repaired_rows.append((repair_row_num, target_id, row))
            logger.info("REPAIRED: ID %s (%s) - row %d backfilled with check data.", target_id, candidate_name, repair_row_num)
        else:
            new_rows.append(row)
            existing_ids.add(target_id)  # guard against dupes within this same run

    except Exception as e:
        skipped_errors.append((raw_id_for_log, raw_name_for_log, str(e)))
        logger.error(
            "SKIPPED due to unexpected error - row %s (applicantId=%s, candidateName=%s): %s",
            row_num, raw_id_for_log, raw_name_for_log, e,
        )
        continue

# ===========================
# APPEND TO PRE-JOINING SHEET (preserving everything else)
# ===========================


for repair_row_num, _repair_target_id, row_dict in repaired_rows:
    for col_name, value in row_dict.items():
        if col_name.strip() == "S.No":
            continue  # never touch an existing row's S.No
        col_idx = header_map.get(col_name.strip())
        if col_idx is None:
            logger.warning("Column '%s' not found in Pre-Joining headers - value skipped.", col_name)
            continue
        cell = ws.cell(row=repair_row_num, column=col_idx, value=value)
        style_cell(cell)

for row_dict in new_rows:
    for col_name, value in row_dict.items():
        col_idx = header_map.get(col_name.strip())
        if col_idx is None:
            # Column name in our mapping didn't match the sheet's actual
            # header text exactly (e.g. trailing space differences) -
            # flag it instead of silently dropping data.
            logger.warning("Column '%s' not found in Pre-Joining headers - value skipped.", col_name)
            continue
        cell = ws.cell(row=next_row, column=col_idx, value=value)
        style_cell(cell)

    if sno_col:
        sno_cell = ws.cell(row=next_row, column=sno_col, value=next_sno)
        style_cell(sno_cell)
        next_sno += 1

    next_row += 1

wb.save(OUTPUT_FILE)

# A dedicated, shareable list of "Reference doc not there" cases - so
# whoever chases missing docs doesn't have to dig through the log file.
MISSING_REF_REPORT = None
if skipped_no_file:
    MISSING_REF_REPORT = f"Missing_Reference_Docs_{RUN_TS}.xlsx"
    missing_df = pd.DataFrame(skipped_no_file)
    missing_df["Status"] = "Reference doc not there"
    missing_df.to_excel(MISSING_REF_REPORT, index=False)

# ===========================
# SUMMARY
# ===========================

logger.info("=" * 60)
logger.info("Added %d new candidate(s) to Pre-Joining.", len(new_rows))
logger.info("Repaired %d incomplete row(s) (backfilled in place): %s", len(repaired_rows), [rid for _, rid, _ in repaired_rows] if repaired_rows else [])
logger.info("Skipped %d duplicate(s): %s", len(skipped_duplicates), skipped_duplicates)
logger.info(
    "%d candidate(s) SKIPPED - Reference doc not there: %s",
    len(skipped_no_file),
    [d["ApplicantId"] for d in skipped_no_file],
)
if MISSING_REF_REPORT:
    logger.info("List of missing-reference-doc candidates saved to: %s", MISSING_REF_REPORT)
if skip_notes:
    logger.info("%d check(s) set to NA due to missing/mismatched block:", len(skip_notes))
    for note in skip_notes:
        logger.info("  - %s", note)
if skipped_errors:
    logger.error(
        "%d candidate(s) SKIPPED due to unexpected errors (fix these and re-run just for them):",
        len(skipped_errors),
    )
    for raw_id, raw_name, err in skipped_errors:
        logger.error("  - applicantId=%s, candidateName=%s: %s", raw_id, raw_name, err)
logger.info("Saved to: %s", OUTPUT_FILE)
logger.info("Full run log saved to: %s", LOG_FILE)
logger.info("Review this file, then replace the live Master Tracker yourself once verified.")
