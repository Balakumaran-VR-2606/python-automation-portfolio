"""Refresh four Infosys progress workbooks from the conventional master tracker."""

from __future__ import annotations

import math
import os
import queue
import re
import shutil
import threading
import uuid
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk


MASTER_NAME = "Infy_Conventional-Master Tracker.xlsx"
DATE_FORMAT = "%d-%b-%Y"
GDC_PATTERN = "Lateral Pre-Employment GDC * - Digiverifier.xlsx"
EXPORT_RESULTS = {"GREEN", "RED", "ORANGE", "MAROON", "STOPCHECK", "STOP CHECK"}
XL_PASTE_FORMATS = -4122
XL_OPENXML_WORKBOOK = 51
INVALID_NEGATIVE_HEADERS = {
    "fr-can id_tat",
    "fr-emp id_tat",
    "ir-tat",
    "fr-tat",
    "ce1-tat",
    "ce1_due_date",
}

# All filtering is by the actual "Package Name" column. In these workbooks,
# "Report Type" contains FINAL/SUPPLEMENTARY values rather than package names.
TARGETS = [
    {
        "label": "Lateral Pre Joining",
        "sheet": "Pre-Joining",
        "packages": {"pre-joining", "pre-joining lateral"},
        "pattern": "Lateral Pre Joining * - Digiverifier.xlsx",
        "filename": "Lateral Pre Joining {date} - Digiverifier.xlsx",
        "formulas": [
            ("N", "=WORKDAY.INTL(RC[-1],13,1,'H-List'!C1:C1)", "M"),
            ("N", "=WORKDAY.INTL(RC[-1],10,1,'H-List'!C1:C1)", "M"),
            ("O", "=IF(RC[8]=\"\",NETWORKDAYS.INTL(RC[-2],TODAY(),1,'H-List'!C1:C1),RC[8])", None),
            ("Z", "=NETWORKDAYS.INTL(RC[-13],RC[-2],1,'H-List'!C1:C1)", None),
            ("AI", "=NETWORKDAYS.INTL(RC[36],RC[-2],1,'H-List'!C1:C1)", None),
            ("BM", "=NETWORKDAYS.INTL(RC[-15],RC[-3],1,'H-List'!C1:C1)", None),
            ("BT", "=IF(RC[-1]=\"\",\"\",WORKDAY.INTL(RC[-1],13,1,'H-List'!C1:C1))", "BS"),
            ("BT", "=IF(RC[-1]=\"\",\"\",WORKDAY.INTL(RC[-1],10,1,'H-List'!C1:C1))", "BS"),
            ("CI", "=NETWORKDAYS.INTL(RC[-9],RC[-1],1,'H-List'!C1:C1)", None),
        ],
    },
    {
        "label": "Lateral Post Employment",
        "sheet": "PE & CVC",
        "packages": {"post employment"},
        "pattern": "Lateral Post Employment * - Digiverifier.xlsx",
        "filename": "Lateral Post Employment {date} - Digiverifier.xlsx",
        "formulas": [
            ("N", "=WORKDAY.INTL(RC[-1],13,1,'H-List'!C1:C1)", "M"),
            ("N", "=WORKDAY.INTL(RC[-1],10,1,'H-List'!C1:C1)", "M"),
            ("O", "=IF(RC[8]=\"\",NETWORKDAYS.INTL(RC[-2],TODAY(),1,'H-List'!C1:C1),RC[8])", None),
            ("Z", "=NETWORKDAYS.INTL(RC[-13],RC[-2],1,'H-List'!C1:C1)", None),
            ("AC", "=NETWORKDAYS.INTL(RC[-16],RC[-2],1,'H-List'!C1:C1)", None),
            ("BF", "=NETWORKDAYS.INTL(RC[-14],RC[-3],1,'H-List'!C1:C1)", None),
            ("BM", "=IF(RC[-1]=\"\",\"\",WORKDAY.INTL(RC[-1],13,1,'H-List'!C1:C1))", "BL"),
            ("BM", "=IF(RC[-1]=\"\",\"\",WORKDAY.INTL(RC[-1],10,1,'H-List'!C1:C1))", "BL"),
            ("CA", "=NETWORKDAYS.INTL(RC[-9],RC[-1],1,'H-List'!C1:C1)", None),
        ],
    },
    {
        "label": "CVC Lateral",
        "sheet": "PE & CVC",
        "packages": {"cvc lateral", "cvc fresher"},
        "pattern": "CVC Lateral * - Digiverifier.xlsx",
        "filename": "CVC Lateral {date} - Digiverifier.xlsx",
        "formulas": [
            ("N", "=WORKDAY.INTL(RC[-1],13,1,'H-List'!C1:C1)", "M"),
            ("N", "=WORKDAY.INTL(RC[-1],10,1,'H-List'!C1:C1)", "M"),
            ("O", "=IF(RC[8]=\"\",NETWORKDAYS.INTL(RC[-2],TODAY(),1,'H-List'!C1:C1),RC[8])", None),
            ("Z", "=NETWORKDAYS.INTL(RC[-13],RC[-2],1,'H-List'!C1:C1)", None),
            ("AC", "=NETWORKDAYS.INTL(RC[-16],RC[-2],1,'H-List'!C1:C1)", None),
            ("BF", "=NETWORKDAYS.INTL(RC[-14],RC[-3],1,'H-List'!C1:C1)", None),
            ("BM", "=IF(RC[-1]=\"\",\"\",WORKDAY.INTL(RC[-1],13,1,'H-List'!C1:C1))", "BL"),
            ("BM", "=IF(RC[-1]=\"\",\"\",WORKDAY.INTL(RC[-1],10,1,'H-List'!C1:C1))", "BL"),
            ("CA", "=NETWORKDAYS.INTL(RC[-9],RC[-1],1,'H-List'!C1:C1)", None),
        ],
    },
    {
        "label": "Fresher",
        "sheet": "Fresher",
        "packages": {"fresher", "fresher(cvc)"},
        "pattern": "Fresher * Digiverifier.xlsx",
        "filename": "Fresher {date} Digiverifier.xlsx",
        "formulas": [
            ("N", "=WORKDAY.INTL(RC[-1],10,1,'H-List'!C1:C1)", None),
            ("O", "=IF(RC[8]=\"\",NETWORKDAYS.INTL(RC[-2],TODAY(),1,'H-List'!C1:C1),RC[8])", None),
            ("Z", "=NETWORKDAYS.INTL(RC[-13],RC[-2],1,'H-List'!C1:C1)", None),
            ("AI", "=NETWORKDAYS.INTL(RC[10],RC[-2],1,'H-List'!C1:C1)", None),
            ("BM", "=NETWORKDAYS.INTL(RC[-15],RC[-4],1,'H-List'!C1:C1)", None),
        ],
    },
]


def normalize(value: object) -> str:
    if value is None:
        return ""
    return " ".join(str(value).replace("\xa0", " ").strip().split()).casefold()


def is_date_header(value: object) -> bool:
    header = normalize(value).replace("_", " ").replace("-", " ")
    tokens = set(re.findall(r"[a-z]+", header))
    if tokens.intersection({"sla", "tat", "count"}):
        return False
    return bool(tokens.intersection({"date", "sdate", "edate", "dob"}))


def as_column(values: object, expected_rows: int) -> list[object]:
    if expected_rows <= 0:
        return []
    if expected_rows == 1:
        if isinstance(values, tuple) and len(values) == 1 and isinstance(values[0], tuple):
            return [values[0][0]]
        return [values]
    return [row[0] for row in values]


def clean_cell(value: object) -> object:
    if value is None:
        return None
    if isinstance(value, str):
        stripped = value.strip()
        if not stripped or stripped.upper() in {"#VALUE!", "#N/A", "#REF!", "#DIV/0!", "#NAME?", "#NUM!", "#NULL!"}:
            return None
        if stripped in {"-33041", "33041"}:
            return None
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        # Excel CVErr values are returned by COM as integer Variant values.
        if value in (-33041, 33041) or (
            isinstance(value, int) and -2147483648 <= value <= -2147480000
        ):
            return None
        if isinstance(value, float) and not math.isfinite(value):
            return None
    return value


def excel_serial(day: datetime) -> float:
    return (day - datetime(1899, 12, 30)).total_seconds() / 86400


def row_bounds(rows: list[int]) -> list[tuple[int, int]]:
    if not rows:
        return []
    rows = sorted(rows)
    blocks = []
    start = previous = rows[0]
    for row in rows[1:]:
        if row != previous + 1:
            blocks.append((start, previous))
            start = row
        previous = row
    blocks.append((start, previous))
    return blocks


def column_index(name: str) -> int:
    result = 0
    for char in name.upper():
        result = result * 26 + (ord(char) - ord("A") + 1)
    return result


def get_headers(ws) -> tuple[list[object], dict[str, int]]:
    max_col = ws.UsedRange.Columns.Count
    raw = ws.Range(ws.Cells(1, 1), ws.Cells(1, max_col)).Value2
    header_row = list(raw[0]) if isinstance(raw, tuple) else [raw]
    return header_row, {normalize(value): i + 1 for i, value in enumerate(header_row) if normalize(value)}


def last_data_row(ws, id_col: int) -> int:
    return ws.Cells(ws.Rows.Count, id_col).End(-4162).Row  # xlUp


def formula_blocks(ws, column: str, formula: str, rows: list[int]) -> None:
    col = column_index(column)
    for first, last in row_bounds(rows):
        ws.Range(ws.Cells(first, col), ws.Cells(last, col)).FormulaR1C1 = formula


def apply_tracker_formulas(ws, specs: list[tuple[str, str, str | None]], log, holiday_range: str) -> tuple[list[int], dict[str, int]]:
    _, headers = get_headers(ws)
    for required in ("c_id", "case_sdate", "package name", "status", "o_date", "o_result"):
        if required not in headers:
            raise ValueError(f"{ws.Name}: required column '{required}' was not found.")

    last_row = last_data_row(ws, headers["c_id"])
    if last_row < 2:
        return [], headers
    ids = as_column(ws.Range(ws.Cells(2, headers["c_id"]), ws.Cells(last_row, headers["c_id"])).Value2, last_row - 1)
    data_rows = [index + 2 for index, value in enumerate(ids) if value is not None and str(value).strip()]
    data_row_set = set(data_rows)
    if not data_rows:
        return [], headers

    formula_columns = sorted({column_index(col) for col, _, _ in specs})
    minimum_col, maximum_col = min(formula_columns), max(formula_columns)
    for col_number in formula_columns:
        if col_number > ws.Columns.Count:
            raise ValueError(f"{ws.Name}: formula column index is outside the worksheet.")

    cutoff = excel_serial(datetime(2025, 8, 28))
    for column, formula, date_column in specs:
        formula = formula.replace("'H-List'!C1:C1", holiday_range)
        if date_column:
            date_col = column_index(date_column)
            dates = as_column(ws.Range(ws.Cells(2, date_col), ws.Cells(last_row, date_col)).Value2, last_row - 1)
            historic_rows, current_rows = [], []
            for index, row in enumerate(range(2, last_row + 1)):
                if row not in data_row_set:
                    continue
                raw_date = dates[index]
                try:
                    is_historic = raw_date is not None and float(raw_date) <= cutoff
                except (TypeError, ValueError):
                    is_historic = False
                (historic_rows if is_historic else current_rows).append(row)
            if ",13,1," in formula:
                formula_blocks(ws, column, formula, historic_rows)
            else:
                formula_blocks(ws, column, formula, current_rows)
        else:
            formula_blocks(ws, column, formula, data_rows)

    log(f"  {ws.Name}: applied {len(specs)} formula rules across {len(data_rows):,} populated rows.")
    return data_rows, headers


def freeze_formula_results(ws, specs: list[tuple[str, str, str | None]], rows: list[int]) -> None:
    for column in {spec[0] for spec in specs}:
        col = column_index(column)
        for first, last in row_bounds(rows):
            value_range = ws.Range(ws.Cells(first, col), ws.Cells(last, col))
            value_range.Value2 = value_range.Value2


def find_template(folder: Path, output_path: Path, pattern: str) -> Path:
    if output_path.is_file():
        return output_path
    candidates = [p for p in folder.glob(pattern) if p.is_file() and not p.name.startswith("~$")]
    if not candidates:
        raise FileNotFoundError(f"No workbook found matching '{pattern}' in {folder}.")
    return max(candidates, key=lambda p: p.stat().st_mtime)


def find_latest_gdc_file(folder: Path) -> Path:
    candidates = [p for p in folder.glob(GDC_PATTERN) if p.is_file() and not p.name.startswith("~$")]
    if not candidates:
        raise FileNotFoundError(f"No GDC progress workbook found matching '{GDC_PATTERN}' in {folder}.")

    def date_and_modified(path: Path) -> tuple[datetime, float]:
        match = re.search(r"(\d{2}-[A-Za-z]{3}-\d{4})", path.stem)
        try:
            file_date = datetime.strptime(match.group(1), DATE_FORMAT) if match else datetime.min
        except ValueError:
            file_date = datetime.min
        return file_date, path.stat().st_mtime

    return max(candidates, key=date_and_modified)


def is_dated_progress_file(path: Path) -> bool:
    match = re.search(r"(\d{2}-[A-Za-z]{3}-\d{4})", path.stem)
    if not match:
        return False
    try:
        datetime.strptime(match.group(1), DATE_FORMAT)
        return True
    except ValueError:
        return False


def backup_files(files: list[Path], folder: Path, log) -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    backup_dir = folder / "Progress_Automation_Backups" / stamp
    backup_dir.mkdir(parents=True, exist_ok=True)
    copied = set()
    for path in files:
        resolved = str(path.resolve()).casefold()
        if resolved in copied or not path.exists():
            continue
        copied.add(resolved)
        shutil.copy2(path, backup_dir / path.name)
    log(f"Backup copies saved in: {backup_dir}")
    return backup_dir


def copy_progress_values(excel, master_rows: tuple, master_headers: dict[str, int], source_ws, spec: dict, template_copy: Path, destination: Path, log) -> int:
    package_col = master_headers["package name"]
    id_col = master_headers["c_id"]
    status_col = master_headers["status"]
    o_date_col = master_headers["o_date"]
    o_result_col = master_headers["o_result"]
    selected_packages = spec["packages"]
    selected_rows = []

    for record in master_rows:
        case_id = record[id_col - 1] if id_col - 1 < len(record) else None
        if case_id is None or not str(case_id).strip():
            continue
        package = normalize(record[package_col - 1] if package_col - 1 < len(record) else None)
        if package in selected_packages:
            selected_rows.append(record)

    if not selected_rows:
        raise ValueError(f"No source rows matched {spec['label']} package(s): {', '.join(sorted(selected_packages))}.")

    target_wb = excel.Workbooks.Open(str(template_copy), UpdateLinks=0, ReadOnly=False, AddToMru=False)
    if target_wb.ReadOnly:
        target_wb.Close(SaveChanges=False)
        raise PermissionError(f"Progress workbook is open read-only: {template_copy.name}. Close it and try again.")
    try:
        target_ws = None
        for sheet in target_wb.Worksheets:
            if normalize(sheet.Name) == normalize(spec["sheet"]):
                target_ws = sheet
                break
        if target_ws is None:
            target_ws = target_wb.Worksheets(1)

        target_headers, target_map = get_headers(target_ws)
        if "c_id" not in target_map:
            raise ValueError(f"{template_copy.name}: target header 'C_ID' was not found.")
        source_header_row, _ = get_headers(source_ws)
        source_map = {normalize(value): i + 1 for i, value in enumerate(source_header_row) if normalize(value)}
        missing = [str(value).strip() for value in target_headers if normalize(value) and normalize(value) not in source_map]
        if missing:
            raise ValueError(f"{template_copy.name}: these target fields are missing in the tracker: {', '.join(missing)}")

        # Carry export status into the prepared progress output only when the
        # tracker status was already updated from an O_Date + colour result.
        data_matrix = []
        for record in selected_rows:
            output_row = []
            for header in target_headers:
                normalized = normalize(header)
                if not normalized:
                    output_row.append(None)
                    continue
                value = record[source_map[normalized] - 1] if source_map[normalized] - 1 < len(record) else None
                if normalized == "s.no":
                    value = len(data_matrix) + 1
                elif normalized == "ed_univ" and (value is None or not str(value).strip()):
                    value = "-"
                elif normalized == "ce1_due_date":
                    r_date_idx = source_map.get("ce1_r_date")
                    r_date = record[r_date_idx - 1] if r_date_idx and r_date_idx - 1 < len(record) else None
                    if r_date is None or not str(r_date).strip():
                        value = None
                if is_date_header(header) and isinstance(value, (int, float)) and not isinstance(value, bool):
                    if value < 367 or value > 2958465:
                        value = None
                if normalized in INVALID_NEGATIVE_HEADERS and isinstance(value, (int, float)) and value < 0:
                    value = None
                output_row.append(clean_cell(value))
            data_matrix.append(tuple(output_row))

        used = target_ws.UsedRange
        last_used_row = max(1, used.Row + used.Rows.Count - 1)
        target_last_col = max(len(target_headers), used.Column + used.Columns.Count - 1)
        if last_used_row >= 2:
            target_ws.Range(target_ws.Cells(2, 1), target_ws.Cells(last_used_row, target_last_col)).ClearContents()

        if len(data_matrix) > max(0, last_used_row - 1) and last_used_row >= 2:
            target_ws.Rows(2).Copy()
            target_ws.Range(
                target_ws.Cells(last_used_row + 1, 1),
                target_ws.Cells(len(data_matrix) + 1, target_last_col),
            ).PasteSpecial(Paste=XL_PASTE_FORMATS)
            excel.CutCopyMode = False

        data_end_col = max(
            (index for index, header in enumerate(target_headers, start=1) if normalize(header)),
            default=len(target_headers),
        )
        clear_through_row = max(last_used_row, len(data_matrix) + 1)
        if clear_through_row >= 2 and target_last_col >= 1:
            old_body = target_ws.Range(
                target_ws.Cells(2, 1),
                target_ws.Cells(clear_through_row, target_last_col),
            )
            old_body.Borders.LineStyle = 0  # xlLineStyleNone; redraw only the populated table below.

        if data_matrix:
            end_row = len(data_matrix) + 1
            output_values = target_ws.Range(target_ws.Cells(2, 1), target_ws.Cells(end_row, len(target_headers)))
            output_values.Value2 = tuple(data_matrix)
            output_range = target_ws.Range(target_ws.Cells(2, 1), target_ws.Cells(end_row, data_end_col))
            output_range.Font.Name = "Calibri"
            output_range.Font.Size = 11
            output_range.WrapText = True
            output_range.HorizontalAlignment = -4108  # xlCenter
            output_range.VerticalAlignment = -4108  # xlCenter
            output_range.Borders.LineStyle = 1  # xlContinuous
            output_range.Borders.Weight = 2  # xlThin
            for column, header in enumerate(target_headers, start=1):
                if column <= data_end_col:
                    column_range = target_ws.Range(
                        target_ws.Cells(2, column),
                        target_ws.Cells(end_row, column),
                    )
                    if is_date_header(header):
                        column_range.NumberFormat = "dd-mmm-yy"
                    elif set(re.findall(r"[a-z]+", normalize(header))).intersection({"sla", "tat", "count"}):
                        column_range.NumberFormat = "General"
            target_ws.Rows(f"2:{end_row}").AutoFit()
        target_wb.Save()
    finally:
        target_wb.Close(SaveChanges=False)

    log(f"  {spec['label']}: {len(data_matrix):,} tracker rows prepared -> {destination.name}")
    return len(data_matrix)


def run_automation(folder: Path, date_text: str, log) -> list[Path]:
    try:
        run_date = datetime.strptime(date_text, DATE_FORMAT)
    except ValueError as exc:
        raise ValueError("Enter the report date as DD-MMM-YYYY, for example 05-Oct-2026.") from exc
    date_label = run_date.strftime(DATE_FORMAT)
    master_path = folder / MASTER_NAME
    if not master_path.is_file():
        raise FileNotFoundError(f"Master Tracker not found: {master_path}")
    gdc_destination = folder / f"Lateral Pre-Employment GDC {date_label} - Digiverifier.xlsx"
    gdc_source = find_latest_gdc_file(folder)

    resolved_targets = []
    for spec in TARGETS:
        destination = folder / spec["filename"].format(date=date_label)
        template = find_template(folder, destination, spec["pattern"])
        resolved_targets.append((spec, template, destination))

    progress_patterns = [spec["pattern"] for spec in TARGETS] + [GDC_PATTERN]
    old_progress_files = {
        path
        for pattern in progress_patterns
        for path in folder.glob(pattern)
        if path.is_file() and not path.name.startswith("~$") and is_dated_progress_file(path)
    }

    backup_inputs = [master_path]
    for _, template, destination in resolved_targets:
        backup_inputs.extend([template, destination])
    backup_inputs.extend([gdc_source, gdc_destination])
    backup_inputs.extend(sorted(old_progress_files, key=lambda path: path.name.casefold()))
    backup_dir = backup_files(backup_inputs, folder, log)

    temp_master = folder / f".Progress_Automation_{uuid.uuid4().hex}_master.xlsx"
    shutil.copy2(master_path, temp_master)
    temp_targets: list[tuple[Path, Path]] = []
    for _, template, destination in resolved_targets:
        temp_copy = folder / f".Progress_Automation_{uuid.uuid4().hex}_target.xlsx"
        shutil.copy2(template, temp_copy)
        temp_targets.append((temp_copy, destination))
    temp_gdc = folder / f".Progress_Automation_{uuid.uuid4().hex}_gdc.xlsx"
    shutil.copy2(gdc_source, temp_gdc)

    excel = None
    master_wb = None
    ready_to_commit = False
    try:
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False

        log("Opening a working copy of the master tracker...")
        master_wb = excel.Workbooks.Open(str(temp_master), UpdateLinks=0, ReadOnly=False, AddToMru=False)
        if master_wb.ReadOnly:
            raise PermissionError("The working copy of the Master Tracker opened read-only.")

        source_sheets: dict[str, object] = {}
        sheet_data: dict[str, tuple] = {}
        status_updates = 0
        for spec in TARGETS:
            if spec["sheet"] not in source_sheets:
                ws = master_wb.Worksheets(spec["sheet"])
                holiday_ws = master_wb.Worksheets("H-List")
                holiday_last_row = max(1, last_data_row(holiday_ws, 1))
                holiday_range = f"'H-List'!R1C1:R{holiday_last_row}C1"
                data_rows, headers = apply_tracker_formulas(ws, spec["formulas"], log, holiday_range)
                source_sheets[spec["sheet"]] = (ws, headers, data_rows)

        log("Recalculating tracker formulas in Excel...")
        excel.CalculateFullRebuild()

        # Freeze only the requested formula columns after Excel has calculated them.
        for spec in TARGETS:
            ws, _, data_rows = source_sheets[spec["sheet"]]
            freeze_formula_results(ws, spec["formulas"], data_rows)

        # Update Pending -> EXPORT only when the overall report date and a
        # completed colour result are both present in the Master Tracker.
        for sheet_name, (ws, headers, data_rows) in source_sheets.items():
            if not data_rows:
                continue
            last_row = max(data_rows)
            last_col = ws.UsedRange.Columns.Count
            raw = ws.Range(ws.Cells(1, 1), ws.Cells(last_row, last_col)).Value2
            matrix = [list(row) for row in raw]
            package_idx = headers["package name"] - 1
            status_idx = headers["status"] - 1
            date_idx = headers["o_date"] - 1
            result_idx = headers["o_result"] - 1
            id_idx = headers["c_id"] - 1
            allowed_packages = set().union(*(spec["packages"] for spec in TARGETS if spec["sheet"] == sheet_name))
            sheet_updates = 0
            for row_number in data_rows:
                record = matrix[row_number - 1]
                if normalize(record[package_idx]) not in allowed_packages:
                    continue
                status = normalize(record[status_idx])
                result = " ".join(str(record[result_idx] or "").strip().upper().replace("\xa0", " ").split())
                report_date = record[date_idx]
                if isinstance(report_date, (int, float)):
                    has_report_date = report_date > 0
                else:
                    has_report_date = report_date is not None and str(report_date).strip() != ""
                if status == "pending" and has_report_date and result in EXPORT_RESULTS:
                    record[status_idx] = "EXPORT"
                    ws.Cells(row_number, headers["status"]).Value2 = "EXPORT"
                    sheet_updates += 1
            status_updates += sheet_updates
            sheet_data[sheet_name] = tuple(tuple(row) for row in matrix[1:])
            log(f"  {sheet_name}: marked {sheet_updates:,} pending report(s) EXPORT from O_Date/O_Result.")

        log(f"Status updates applied: {status_updates:,}")
        master_wb.Save()

        counts = []
        for (spec, _, destination), (temp_copy, _) in zip(resolved_targets, temp_targets):
            source_ws, headers, _ = source_sheets[spec["sheet"]]
            master_rows = sheet_data[spec["sheet"]]
            count = copy_progress_values(
                excel,
                master_rows,
                headers,
                source_ws,
                spec,
                temp_copy,
                destination,
                log,
            )
            counts.append(count)

        master_wb.Close(SaveChanges=True)
        master_wb = None
        ready_to_commit = True
        log("All four progress workbooks are prepared. Updating the dated output files and Master Tracker...")

    finally:
        if master_wb is not None:
            try:
                master_wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.ScreenUpdating = True
                excel.DisplayAlerts = True
                excel.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    if not ready_to_commit:
        for temporary in [temp_master, temp_gdc, *(item[0] for item in temp_targets)]:
            try:
                if temporary.exists():
                    temporary.unlink()
            except OSError:
                pass
        raise RuntimeError("The Excel preparation step did not complete; originals were not replaced.")

    commit_order = [*temp_targets, (temp_master, master_path), (temp_gdc, gdc_destination)]
    committed: list[Path] = []
    try:
        for temporary, destination in commit_order:
            os.replace(temporary, destination)
            committed.append(destination)
    except Exception:
        backup_by_name = {p.name.casefold(): p for p in backup_dir.iterdir() if p.is_file()}
        for destination in reversed(committed):
            backup = backup_by_name.get(destination.name.casefold())
            if backup and backup.exists():
                shutil.copy2(backup, destination)
            else:
                try:
                    destination.unlink()
                except OSError:
                    pass
        raise
    finally:
        for temporary, _ in commit_order:
            try:
                if temporary.exists():
                    temporary.unlink()
            except OSError:
                pass

    preserved_outputs = {
        destination.resolve()
        for _, _, destination in resolved_targets
    } | {gdc_destination.resolve()}
    for previous_file in sorted(old_progress_files, key=lambda path: path.name.casefold()):
        if previous_file.resolve() in preserved_outputs or not previous_file.exists():
            continue
        try:
            previous_file.unlink()
            log(f"Removed old progress file after backup: {previous_file.name}")
        except OSError as exc:
            log(f"Could not remove old progress file {previous_file.name}: {exc}")

    outputs = [destination for _, destination in temp_targets]
    outputs.append(gdc_destination)
    log(f"GDC workbook copied: {gdc_source.name} -> {gdc_destination.name}; source workbook left unchanged.")
    log(f"Backup copies: {backup_dir}")
    return outputs


class ProgressSheetApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Progress Sheet Automation")
        self.geometry("940x740")
        self.minsize(820, 640)
        self.configure(bg="#eef3f8")
        self.events: queue.Queue[tuple] = queue.Queue()
        self.running = False
        self.script_dir = Path(__file__).resolve().parent

        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("App.TFrame", background="#eef3f8")
        style.configure("Panel.TFrame", background="#ffffff")
        style.configure("Title.TLabel", background="#12365a", foreground="#ffffff", font=("Segoe UI", 20, "bold"))
        style.configure("Subtitle.TLabel", background="#12365a", foreground="#d8e7f5", font=("Segoe UI", 10))
        style.configure("Section.TLabel", background="#ffffff", foreground="#183b5b", font=("Segoe UI", 11, "bold"))
        style.configure("Hint.TLabel", background="#ffffff", foreground="#60758a", font=("Segoe UI", 9))
        style.configure("Primary.TButton", background="#1769aa", foreground="#ffffff", font=("Segoe UI", 10, "bold"), padding=(16, 9), borderwidth=0)
        style.map("Primary.TButton", background=[("disabled", "#a8bdd0"), ("active", "#12558a")])
        style.configure("Secondary.TButton", background="#e5edf5", foreground="#24435f", font=("Segoe UI", 9, "bold"), padding=(10, 7), borderwidth=0)
        style.map("Secondary.TButton", background=[("active", "#d5e2ef")])

        header = tk.Frame(self, bg="#12365a", padx=23, pady=17)
        header.pack(fill="x")
        ttk.Label(header, text="PROGRESS SHEET AUTOMATION", style="Title.TLabel").pack(anchor="w")
        ttk.Label(header, text="Refresh package progress sheets from the Master Tracker.", style="Subtitle.TLabel").pack(anchor="w", pady=(3, 0))

        outer = ttk.Frame(self, style="App.TFrame", padding=(17, 14, 17, 12))
        outer.pack(fill="both", expand=True)

        date_panel = ttk.Frame(outer, style="Panel.TFrame", padding=14)
        date_panel.pack(fill="x", pady=(0, 10))
        ttk.Label(date_panel, text="Progress sheet date", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(date_panel, text="Enter the date to use in all four output filenames.", style="Hint.TLabel").grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.date_var = tk.StringVar(value=datetime.now().strftime(DATE_FORMAT))
        ttk.Entry(date_panel, textvariable=self.date_var, width=20, font=("Segoe UI", 11)).grid(row=0, column=1, rowspan=2, sticky="e", padx=(20, 3), ipady=4)
        date_panel.columnconfigure(0, weight=1)

        mapping = ttk.Frame(outer, style="Panel.TFrame", padding=14)
        mapping.pack(fill="x", pady=(0, 10))
        ttk.Label(mapping, text="Package mapping", style="Section.TLabel").pack(anchor="w", pady=(0, 7))
        mappings = (
            "Pre-Joining / Pre-Joining Lateral  →  Lateral Pre Joining",
            "Post Employment  →  Lateral Post Employment",
            "CVC Lateral / CVC Fresher  →  CVC Lateral",
            "Fresher / Fresher(CVC)  →  Fresher",
        )
        for value in mappings:
            ttk.Label(mapping, text=value, style="Hint.TLabel").pack(anchor="w", pady=2)
        ttk.Label(mapping, text="GDC: latest workbook is copied to the selected date; old dated progress files are backed up then removed.", style="Hint.TLabel").pack(anchor="w", pady=(7, 0))

        button_bar = ttk.Frame(outer, style="App.TFrame")
        button_bar.pack(fill="x", pady=(0, 8))
        self.status = tk.StringVar(value="Ready")
        ttk.Label(button_bar, textvariable=self.status, background="#eef3f8", foreground="#526b83", font=("Segoe UI", 9, "bold")).pack(side="left")
        self.start_button = ttk.Button(button_bar, text="Start Progress Sheet Update", style="Primary.TButton", command=self.start)
        self.start_button.pack(side="right")
        ttk.Button(button_bar, text="Clear Log", style="Secondary.TButton", command=self.clear_log).pack(side="right", padx=(0, 8))

        log_panel = ttk.Frame(outer, style="Panel.TFrame", padding=12)
        log_panel.pack(fill="both", expand=True)
        ttk.Label(log_panel, text="Run log", style="Section.TLabel").pack(anchor="w", pady=(0, 6))
        log_body = ttk.Frame(log_panel, style="Panel.TFrame")
        log_body.pack(fill="both", expand=True)
        self.log = tk.Text(log_body, wrap="word", font=("Segoe UI", 9), bg="#fbfdff", fg="#27384a", insertbackground="#1769aa", relief="flat", padx=10, pady=9)
        self.log.tag_configure("heading", foreground="#174d79", font=("Segoe UI", 9, "bold"))
        self.log.tag_configure("success", foreground="#147544")
        self.log.tag_configure("error", foreground="#b8322a")
        scrollbar = ttk.Scrollbar(log_body, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=scrollbar.set)
        self.log.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y", padx=(5, 0))
        ttk.Label(outer, text=f"Working folder: {self.script_dir}", style="Hint.TLabel").pack(fill="x", pady=(8, 0))
        self.after(100, self.poll)

    def append(self, text: str) -> None:
        lowered = text.lower()
        tag = "error" if "error" in lowered or "failed" in lowered else "success" if "completed" in lowered else "heading" if "opening" in lowered or "backup" in lowered else None
        if tag:
            self.log.insert("end", text + "\n", tag)
        else:
            self.log.insert("end", text + "\n")
        self.log.see("end")

    def clear_log(self) -> None:
        if not self.running:
            self.log.delete("1.0", "end")

    def start(self) -> None:
        if self.running:
            return
        date_text = self.date_var.get().strip()
        try:
            datetime.strptime(date_text, DATE_FORMAT)
        except ValueError:
            messagebox.showerror("Invalid date", "Enter the date as DD-MMM-YYYY, for example 05-Oct-2026.", parent=self)
            return
        outputs = [self.script_dir / spec["filename"].format(date=date_text) for spec in TARGETS]
        gdc_output = self.script_dir / f"Lateral Pre-Employment GDC {date_text} - Digiverifier.xlsx"
        summary = "Save and close the Master Tracker and progress workbooks before continuing.\n\nThis run will:\n\n• Apply the requested formulas in the Master Tracker and freeze their results as values.\n• Update Pending to EXPORT when O_Date and an allowed O_Result colour are present.\n• Replace data from row 2 in these four dated progress files:\n\n" + "\n".join(f"  {path.name}" for path in outputs) + f"\n\n• Copy the latest GDC progress workbook to:\n  {gdc_output.name}\n  Its contents will stay unchanged.\n• Back up all old dated progress workbooks, then remove those old copies from Downloads.\n• Keep the Master Tracker in place.\n\nContinue?"
        if not messagebox.askyesno("Confirm progress sheet update", summary, parent=self):
            return
        self.running = True
        self.start_button.configure(state="disabled")
        self.status.set("Preparing workbooks...")
        self.append("Starting Progress Sheet Automation")
        thread = threading.Thread(target=self.worker, args=(date_text,), daemon=True)
        thread.start()

    def worker(self, date_text: str) -> None:
        def log(message: str) -> None:
            self.events.put(("log", message))
        try:
            outputs = run_automation(self.script_dir, date_text, log)
            self.events.put(("done", outputs))
        except Exception as exc:
            self.events.put(("error", str(exc)))

    def poll(self) -> None:
        try:
            while True:
                kind, value = self.events.get_nowait()
                if kind == "log":
                    self.append(value)
                elif kind == "done":
                    self.running = False
                    self.start_button.configure(state="normal")
                    self.status.set("Completed")
                    self.append("Completed successfully.")
                    messagebox.showinfo("Progress sheets updated", "Updated files:\n\n" + "\n".join(path.name for path in value), parent=self)
                elif kind == "error":
                    self.running = False
                    self.start_button.configure(state="normal")
                    self.status.set("Failed")
                    self.append("Error: " + value)
                    messagebox.showerror("Update failed", value, parent=self)
        except queue.Empty:
            pass
        self.after(100, self.poll)


if __name__ == "__main__":
    ProgressSheetApp().mainloop()
