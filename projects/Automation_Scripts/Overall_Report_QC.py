import os
import re
import fitz
import pandas as pd
from pypdf import PdfReader

TRACKER_FILE = r"Infy_Conventional-Master Tracker.xlsx"
REPORTS_FOLDER = r"Reports"


# -----------------------------
# PDF TEXT EXTRACTION
# -----------------------------
def extract_pdf_text(pdf_path):

    text = ""
    page_errors = []

    EXPECTED_FOOTER = " ".join(
        (
            "Registered Office: Digiverifier Solutions Private Limited. "
            "No. #3/1, Langford Road, Shantinagar, Richmond Town, "
            "Bangalore - 560025"
        )
        .upper()
        .replace("–", "-")
        .replace(".", "")
        .replace("#", "")
        .split()
    )

    try:

        reader = PdfReader(pdf_path)

        for page_no, page in enumerate(reader.pages, start=1):

            page_text = page.extract_text()

            # Blank Page Check

            clean_text = page_text.upper()

            # Remove footer
            clean_text = re.sub(
                r"REGISTERED OFFICE.*?560025",
                "",
                clean_text,
                flags=re.I
            )

            # Remove page number
            clean_text = re.sub(
                r"PAGE\s+\d+(?:\s+OF\s+\d+)?",
                "",
                clean_text,
                flags=re.I
            )

            if len(clean_text.strip()) < 30:

                page_errors.append(
                    f"Blank Page Found --> Page : {page_no}"
                )

            else:

                normalized_page = " ".join(
                    page_text.upper()
                    .replace("–", "-")
                    .replace(".", "")
                    .replace("#", "")
                    .split()
                )

                # HEADER CHECK (PAGE 1 ONLY)
                if page_no == 1:

                    if "INFOSYS LIMITED" not in normalized_page:

                        page_errors.append(
                            "Header Missing --> Page : 1"
                        )

                    if (
                        "IL BANGALORE STP PARK I" not in normalized_page
                        or
                        "ELECTRONICS CITY" not in normalized_page
                        or
                        "HOSUR ROAD" not in normalized_page
                        or
                        "560100" not in normalized_page
                    ):

                        page_errors.append(
                            "Header Address Missing --> Page : 1"
                        )


                # FOOTER CHECK (EVERY PAGE)

                if (
                    "REGISTERED OFFICE" not in normalized_page
                    or
                    "LANGFORD ROAD" not in normalized_page
                    or
                    "560025" not in normalized_page
                ):

                    page_errors.append(
                        f"Footer Missing --> Page : {page_no}"
                    )


                # ORG NAME CHECK (EVERY PAGE)

                if (
                    "DIGIVERIFIER SOLUTIONS PRIVATE LIMITED"
                    not in normalized_page
                ):

                    page_errors.append(
                        f"Org Name Wrong --> Page : {page_no}"
                    )


                # PAGE NUMBER CHECK (EVERY PAGE)
               
                page_match = re.search(
                    r"PAGE\s+(\d+)(?:\s+OF\s+\d+)?",
                    normalized_page,
                    re.I
                )

                if (
                    not page_match
                    or
                    int(page_match.group(1)) != page_no
                ):

                    page_errors.append(
                        f"Pg No is Invalid --> Page : {page_no}"
                    )

            if page_text:
                text += page_text + "\n"

    except Exception as e:

        print(f"Unable to read {pdf_path}")
        print(e)

    return text, page_errors

# -----------------------------
# PDF DATA EXTRACTION
# -----------------------------
def extract_pdf_details(pdf_text):

    data = {
        "id": "",
        "id_type": "",
        "report_type": "",
        "status": "",
        "report_date": ""
    }

    # CAN ID (10 digits)
    m = re.search(r"Candidate ID.*?(\d{10})", pdf_text, re.I)

    if m:
        data["id"] = m.group(1)
        data["id_type"] = "CAN"

    # EMP ID (7 digits)
    m = re.search(r"Emp ID.*?(\d{7})", pdf_text, re.I)

    if m:
        data["id"] = m.group(1)
        data["id_type"] = "EMP"

    # Report Type
    m = re.search(
        r"Report Type.*?(FINAL|SUPPLEMENTARY|STOP\s*CHECK)",
        pdf_text,
        re.I
    )

    if m:
        data["report_type"] = (
            m.group(1)
            .replace("  ", " ")
            .strip()
            .upper()
        )

    # Statusr"Verification Status\s+([A-Z]+)",
    m = re.search(
        r"Verification Status\s+([A-Z ]+)",
        pdf_text,
        re.I
    )

    if m:
        data["status"] = m.group(1).upper()

    from datetime import datetime

    report_dates = re.findall(
        r"(?:Report Date).*?(\d{2}-[A-Za-z]{3}-\d{4})",
        pdf_text,
        re.I
    )

    if report_dates:

        latest_date = max(
            report_dates,
            key=lambda x: datetime.strptime(
                x,
                "%d-%b-%Y"
            )
        )

        data["report_date"] = latest_date
    
    return data

def validate_color(pdf_path, expected_status):

    # Each status can render as more than one close RGB shade depending
    # on how the PDF was generated (font/rasterizer rounding) - e.g.
    # GREEN #008000 (32768) vs #007F00 (32512), one G value apart.
    # Every shade listed here is treated as a valid match, not just a
    # friendlier label on a mismatch.
    EXPECTED_COLOR_SHADES = {
        "GREEN": {32768, 32512},
        "ORANGE": {16753920, 16753664},
        "RED": {16711680, 15597568},
        "MAROON": {12582912, 8388608},
        "STOP CHECK": {0},
    }

    COLOR_NAMES = {
        32768: "GREEN",
        32512: "GREEN",
        16753920: "ORANGE",
        16753664: "ORANGE",
        16711680: "RED",
        15597568: "RED",
        8388608: "MAROON",
        12582912: "MAROON",
        0: "BLACK"
    }

    doc = fitz.open(pdf_path)

    found = False

    for page in doc:

        data = page.get_text("dict")

        for block in data["blocks"]:

            if "lines" not in block:
                continue

            for line in block["lines"]:

                for span in line["spans"]:

                    text = span["text"].strip().upper()

                    expected_status = expected_status.strip().upper()

                    if text == expected_status:

                        found = True 

                        actual_color = span["color"]

                        expected_shades = EXPECTED_COLOR_SHADES[
                            expected_status
                        ]

                        if actual_color not in expected_shades:

                            found_color = COLOR_NAMES.get(
                                actual_color,
                                f"UNKNOWN({actual_color})"
                            )

                            # DEBUG: prints the raw color integer so real
                            # shade values seen in production PDFs can be
                            # added to COLOR_NAMES/EXPECTED_COLOR_SHADES
                            # above instead of guessing.
                            print(
                                f"[DEBUG COLOR] file={os.path.basename(pdf_path)} "
                                f"page={page.number + 1} text={text} "
                                f"expected_status={expected_status} "
                                f"expected_shades={sorted(expected_shades)} "
                                f"actual_color_int={actual_color}"
                            )

                            return (
                                f"COLOR MISMATCH | "
                                f"PAGE={page.number + 1} "
                                f"TEXT={text} "
                                f"Expected={expected_status} "
                                f"Found={found_color}"
                            )

    if not found:

        return (
            f"STATUS TEXT NOT FOUND | "
            f"{expected_status}"
        )

    return None

def validate_filename_status(file_name, pdf_status):

    file_upper = file_name.upper()

    statuses = [
        "GREEN",
        "ORANGE",
        "RED",
        "MAROON",
        "STOPCHECK",
        "STOP CHECK"
    ]

    file_status = None

    for status in statuses:

        if status in file_upper:

            file_status = status.replace(
                "STOPCHECK",
                "STOP CHECK"
            )

            break

    if not file_status:

        return (
            "STATUS NOT FOUND IN FILE NAME"
        )

    pdf_status = pdf_status.strip().upper()

    if file_status != pdf_status:

        return (
            f"FILENAME STATUS MISMATCH "
            f"(PDF STATUS={pdf_status}, "
            f"FILE NAME={file_status})"
        )

    return None

from datetime import datetime

def validate_filename_id(file_name, tracker):

    date_fields = {
        "FR-CAN ID_Date": "CAN",
        "SR-CAN ID_Date": "CAN",
        "SR2-CAN ID_Date": "CAN",
        "IR_Date": "CAN",

        "FR-EMP ID_Date": "EMP",
        "SR-EMP ID_Date": "EMP",
        "SR2-EMP ID_Date": "EMP",
        "FR_Date": "EMP",
        "SR_Date": "EMP"
    }

    latest_date = None
    latest_type = None

    for col, id_type in date_fields.items():

        if col not in tracker.index:
            continue

        value = tracker[col]

        if pd.isna(value):
            continue

        try:

            dt = pd.to_datetime(value)

            if latest_date is None or dt > latest_date:

                latest_date = dt
                latest_type = id_type

        except:
            pass

    if latest_type is None:
        return None

    file_upper = file_name.upper()

    if latest_type == "CAN":

        c_id = str(tracker.get("C_ID", "")).strip()

        if c_id and c_id != "nan":

            if c_id not in file_name:

                return (
                    f"FILE NAME ID MISMATCH "
                    f"(Expected C_ID={c_id})"
                )

    elif latest_type == "EMP":

        emp_id = str(tracker.get("Emp_No", "")).strip()

        if (
            emp_id
            and emp_id.lower() != "nan"
            and emp_id.upper() != "NA"
        ):

            if emp_id not in file_name:

                return (
                    f"FILE NAME ID MISMATCH "
                    f"(Expected EMP ID={emp_id})"
                )

        else:

            c_id = str(tracker.get("C_ID", "")).strip()

            if c_id and c_id != "nan":

                if c_id not in file_name:

                    return (
                        f"FILE NAME ID MISMATCH "
                        f"(Expected C_ID={c_id})"
                    )

            else:

                candidate_name = str(
                    tracker.get("C_Name", "")
                ).strip()

                if candidate_name.upper() not in file_upper:

                    return (
                        f"FILE NAME MISMATCH "
                        f"(Expected Name={candidate_name})"
                    )

    return None

def validate_es_priority(pdf_path, pdf_status):

    pdf_status = pdf_status.strip().upper()

    priority_rank = {
        "GREEN": 1,
        "ORANGE": 2,
        "RED": 3,
        "MAROON": 4,
        "STOP CHECK": 5
    }

    colors_found = set()

    doc = fitz.open(pdf_path)

    for page in doc:

        if page.number + 1 == 1:
            continue

        data = page.get_text("dict")

        for block in data["blocks"]:

            if "lines" not in block:
                continue

            for line in block["lines"]:

                for span in line["spans"]:

                    text = span["text"].strip().upper()

                    if text in [
                        "GREEN",
                        "ORANGE",
                        "RED",
                        "MAROON"
                    ]:
                        
                        colors_found.add(text)

    # "STOP CHECK" is 2 words and can be split across spans, so it's
    # detected via a full-page text scan (as before), then folded into
    # the same colors_found set/priority ranking as everything else.
    stop_check_found = False

    for page in doc:

        if page.number + 1 == 1:
            continue

        text = page.get_text().upper()

        if "STOP CHECK" in text:

            stop_check_found = True
            break

    if stop_check_found:
        colors_found.add("STOP CHECK")

    # STOP CHECK is the highest priority status. If it's genuinely
    # present in the ES, that alone makes the report valid - any
    # RED/ORANGE/MAROON also found is simply the constituent check
    # result that triggered the stop, not an error.
    if pdf_status == "STOP CHECK":

        if not stop_check_found:

            return (
                "INVALID ES COLOR "
                "(Verification Status=STOP CHECK, "
                "STOP CHECK not found in ES)"
            )

        return None

    if not colors_found:
        return None

    highest_color = max(
        colors_found,
        key=lambda x: priority_rank[x]
    )

    if highest_color != pdf_status:

        found_text = ", ".join(
            sorted(colors_found, key=lambda x: priority_rank[x])
        )

        return (
            f"INVALID ES COLOR "
            f"(Verification Status={pdf_status}, "
            f"Found in ES={found_text})"
        )

    return None


# -----------------------------
# TRACKER COLUMN MAPPING
# -----------------------------
def get_date_column(folder_name):

    folder_name = folder_name.upper()

    if "CID FINAL" in folder_name:
        return "FR-CAN ID_Date"

    elif "CID SUPPLEMENTARY" in folder_name:
        return "SR-CAN ID_Date"

    elif "EMP ID FINAL" in folder_name:
        return "FR-EMP ID_Date"

    elif "EMP SUPPLEMENTARY" in folder_name:
        return "SR-EMP ID_Date"

    return None

def get_result_column(folder_name):

    folder_name = folder_name.upper()

    if "CID FINAL" in folder_name:
        return "FR-CAN ID_Result"

    elif "CID SUPPLEMENTARY" in folder_name:
        return "SR-CAN ID_Result"

    elif "EMP ID FINAL" in folder_name:
        return "FR-EMP ID_Result"

    elif "EMP SUPPLEMENTARY" in folder_name:
        return "SR-EMP ID_Result"

    elif "STOP CHECK" in folder_name:
        return "O_Result"

    return None


def get_latest_tracker_date(tracker):
    """
    STOP CHECK reports can happen at any stage (CAN or EMP, final or
    supplementary) - there is no single fixed tracker column to compare
    against. Instead, take the LATEST date across all 4 stage columns
    (mirrors the MAX logic already used on the PDF side in
    extract_pdf_details), so whichever stage the stop check actually
    happened at, the comparison still lines up.
    """

    stage_columns = [
        "FR-CAN ID_Date",
        "SR-CAN ID_Date",
        "FR-EMP ID_Date",
        "SR-EMP ID_Date",
    ]

    latest_dt = None

    for col in stage_columns:

        if col not in tracker.index:
            continue

        dt = pd.to_datetime(tracker[col], errors="coerce")

        if pd.notna(dt):

            if latest_dt is None or dt > latest_dt:
                latest_dt = dt

    if latest_dt is None:
        return ""

    return latest_dt.strftime("%d-%b-%Y")

# -----------------------------
# LOAD TRACKER
# -----------------------------
all_sheets = pd.read_excel(
    TRACKER_FILE,
    sheet_name=None
)

frames = []

for sheet_name, df in all_sheets.items():

    if "C_ID" in df.columns:
        frames.append(df)

tracker_df = pd.concat(
    frames,
    ignore_index=True
)

# =====================================
# CLEANUP IDS
# =====================================
tracker_df["C_ID"] = (
    tracker_df["C_ID"]
    .astype(str)
    .str.replace(".0", "", regex=False)
    .str.strip()
)

tracker_df["Emp_No"] = (
    tracker_df["Emp_No"]
    .astype(str)
    .str.replace(".0", "", regex=False)
    .str.strip()
)

total_reports = 0
failed_reports = 0

# -----------------------------
# SCAN REPORTS
# -----------------------------
for root, dirs, files in os.walk(REPORTS_FOLDER):

    for file in files:

        if not file.lower().endswith(".pdf"):
            continue

        total_reports += 1

        pdf_path = os.path.join(root, file)

        pdf_text, page_errors = extract_pdf_text(pdf_path)

        pdf = extract_pdf_details(pdf_text)

        errors = []

        errors.extend(page_errors)
                
        # -------------------
        # FIND TRACKER ROW
        # -------------------

        tracker_row = pd.DataFrame()

        if pdf["id_type"] == "CAN":
            tracker_row = tracker_df[
                tracker_df["C_ID"] == pdf["id"]
            ]

        elif pdf["id_type"] == "EMP":
            tracker_row = tracker_df[
                tracker_df["Emp_No"] == pdf["id"]
            ]

        if tracker_row.empty:
            print("\n" + "=" * 80)
            print(file)
            print("ID NOT FOUND IN TRACKER")
            continue
            
        tracker = tracker_row.iloc[-1]

        # -------------------
        # REPORT TYPE
        # -------------------

        tracker_type = str(
            tracker["Report Type"]
        ).strip().upper()

        if pdf["report_type"]:

            if tracker_type != pdf["report_type"]:

                errors.append(
                    f"REPORT TYPE MISMATCH "
                    f"(Tracker={tracker_type}, "
                    f"PDF={pdf['report_type']})"
                )

        # -------------------
        # STATUS
        # -------------------

        result_column = get_result_column(root)

        if result_column and result_column in tracker.index:

            tracker_status = str(
                tracker[result_column]
            ).strip().upper()

            pdf_status = pdf["status"].strip().upper()

            if tracker_status != pdf_status:

                errors.append(
                    f"STATUS MISMATCH "
                    f"(Tracker={tracker_status}, "
                    f"PDF={pdf_status})"
                )

            color_error = validate_color(
                pdf_path,
                pdf["status"]
            )

            if color_error:

                errors.append(color_error)

            # FILENAME STATUS VALIDATION

            filename_error = validate_filename_status(
                file,
                pdf["status"]
            )

            if filename_error:

                errors.append(
                    filename_error
                )

            # FILE NAME ID VALIDATION
            file_id_error = validate_filename_id(
                file,
                tracker
            )

            if file_id_error:

                errors.append(
                    file_id_error
                )

            # ES PRIORITY VALIDATION
            es_error = validate_es_priority(
                pdf_path,
                pdf["status"]
            )

            if es_error:

                errors.append(es_error)

        # -------------------
        # DATE
        # -------------------

        date_column = get_date_column(root)

        if date_column and date_column in tracker.index:

            tracker_date_value = pd.to_datetime(
                tracker[date_column],
                errors="coerce"
            )

            if pd.notna(tracker_date_value):

                tracker_date = tracker_date_value.strftime(
                    "%d-%b-%Y"
                )

            else:

                tracker_date = ""

            pdf_date = pdf["report_date"]

            if tracker_date and pdf_date:

                if tracker_date != pdf_date:

                    errors.append(
                        f"DATE MISMATCH "
                        f"(Tracker={tracker_date}, "
                        f"PDF={pdf_date})"
                    )

        elif "STOP CHECK" in root.upper():

            # No single fixed column applies to STOP CHECK - compare
            # against whichever stage column in the tracker is latest.
            tracker_date = get_latest_tracker_date(tracker)
            pdf_date = pdf["report_date"]

            if tracker_date and pdf_date:

                if tracker_date != pdf_date:

                    errors.append(
                        f"DATE MISMATCH "
                        f"(Tracker={tracker_date}, "
                        f"PDF={pdf_date})"
                    )

        # -------------------
        # PRINT RESULT
        # -------------------

        if errors:

            failed_reports += 1

            print("\n" + "=" * 80)
            print("FILE :", file)
            print("STATUS : FAIL ❌")

            for err in errors:
                print(err)

        else:

            pass

passed_reports = total_reports - failed_reports

print("\n" + "=" * 80)
print("👍 OVERALL REPORT VALIDATION COMPLETED")
print(f"📄 Total Reports Checked : {total_reports}")
print(f"✅ Passed Reports        : {passed_reports}")
print(f"❌ Failed Reports        : {failed_reports}")
print("🛡️ All QC Validation Rules Executed")
print("✅ Process Finished Successfully")
print("=" * 80)