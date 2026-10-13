"""Apply and freeze the Progress_Sheet_Automation formulas in the Master Tracker only."""

from __future__ import annotations

import shutil
import time
from datetime import datetime
from pathlib import Path


MASTER_PATTERN = "Infy_Conventional-Master Tracker*.xlsx"
HISTORIC_CUTOFF = (datetime(2025, 8, 28) - datetime(1899, 12, 30)).days
XL_UP = -4162
XL_PASTE_VALUES = -4163


# Formula rules copied from Progress_Sheet_Automation.pyw. These are R1C1
# formulas; the holiday-list reference is extended to the last used H-List row.
SHEET_FORMULAS = {
    "Pre-Joining": [
        ("N", "=WORKDAY.INTL(RC[-1],13,1,'H-List'!C1:C1)", "M"),
        ("N", "=WORKDAY.INTL(RC[-1],10,1,'H-List'!C1:C1)", "M"),
        ("O", '=IF(RC[8]="",NETWORKDAYS.INTL(RC[-2],TODAY(),1,\'H-List\'!C1:C1),RC[8])', None),
        ("Z", "=NETWORKDAYS.INTL(RC[-13],RC[-2],1,'H-List'!C1:C1)", None),
        ("AI", "=NETWORKDAYS.INTL(RC[36],RC[-2],1,'H-List'!C1:C1)", None),
        ("BM", "=NETWORKDAYS.INTL(RC[-15],RC[-3],1,'H-List'!C1:C1)", None),
        ("BT", '=IF(OR(RC[-1]="",UPPER(TRIM(RC[-1]))="NA",UPPER(TRIM(RC[-1]))="N/A"),"",WORKDAY.INTL(RC[-1],13,1,\'H-List\'!C1:C1))', "BS"),
        ("BT", '=IF(OR(RC[-1]="",UPPER(TRIM(RC[-1]))="NA",UPPER(TRIM(RC[-1]))="N/A"),"",WORKDAY.INTL(RC[-1],10,1,\'H-List\'!C1:C1))', "BS"),
        ("CI", "=NETWORKDAYS.INTL(RC[-9],RC[-1],1,'H-List'!C1:C1)", None),
    ],
    "PE & CVC": [
        ("N", "=WORKDAY.INTL(RC[-1],13,1,'H-List'!C1:C1)", "M"),
        ("N", "=WORKDAY.INTL(RC[-1],10,1,'H-List'!C1:C1)", "M"),
        ("O", '=IF(RC[8]="",NETWORKDAYS.INTL(RC[-2],TODAY(),1,\'H-List\'!C1:C1),RC[8])', None),
        ("Z", "=NETWORKDAYS.INTL(RC[-13],RC[-2],1,'H-List'!C1:C1)", None),
        ("AC", "=NETWORKDAYS.INTL(RC[-16],RC[-2],1,'H-List'!C1:C1)", None),
        ("BF", "=NETWORKDAYS.INTL(RC[-14],RC[-3],1,'H-List'!C1:C1)", None),
        ("BM", '=IF(OR(RC[-1]="",UPPER(TRIM(RC[-1]))="NA",UPPER(TRIM(RC[-1]))="N/A"),"",WORKDAY.INTL(RC[-1],13,1,\'H-List\'!C1:C1))', "BL"),
        ("BM", '=IF(OR(RC[-1]="",UPPER(TRIM(RC[-1]))="NA",UPPER(TRIM(RC[-1]))="N/A"),"",WORKDAY.INTL(RC[-1],10,1,\'H-List\'!C1:C1))', "BL"),
        ("CA", "=NETWORKDAYS.INTL(RC[-9],RC[-1],1,'H-List'!C1:C1)", None),
    ],
    "Fresher": [
        ("N", "=WORKDAY.INTL(RC[-1],10,1,'H-List'!C1:C1)", None),
        ("O", '=IF(RC[8]="",NETWORKDAYS.INTL(RC[-2],TODAY(),1,\'H-List\'!C1:C1),RC[8])', None),
        ("Z", "=NETWORKDAYS.INTL(RC[-13],RC[-2],1,'H-List'!C1:C1)", None),
        ("AI", "=NETWORKDAYS.INTL(RC[10],RC[-2],1,'H-List'!C1:C1)", None),
        ("BM", "=NETWORKDAYS.INTL(RC[-15],RC[-4],1,'H-List'!C1:C1)", None),
    ],
}


def normalize(value: object) -> str:
    if value is None:
        return ""
    return " ".join(str(value).replace("\xa0", " ").strip().split()).casefold()


def column_index(name: str) -> int:
    result = 0
    for char in name.upper():
        result = result * 26 + ord(char) - ord("A") + 1
    return result


def as_column(values: object, expected_rows: int) -> list[object]:
    if expected_rows <= 0:
        return []
    if expected_rows == 1:
        if isinstance(values, tuple) and len(values) == 1 and isinstance(values[0], tuple):
            return [values[0][0]]
        return [values]
    return [row[0] for row in values]


def row_blocks(rows: list[int]) -> list[tuple[int, int]]:
    if not rows:
        return []
    rows = sorted(rows)
    blocks = []
    first = previous = rows[0]
    for row in rows[1:]:
        if row != previous + 1:
            blocks.append((first, previous))
            first = row
        previous = row
    blocks.append((first, previous))
    return blocks


def get_headers(ws) -> dict[str, int]:
    last_col = ws.UsedRange.Columns.Count
    values = ws.Range(ws.Cells(1, 1), ws.Cells(1, last_col)).Value
    headers = list(values[0]) if isinstance(values, tuple) else [values]
    return {normalize(value): index + 1 for index, value in enumerate(headers) if normalize(value)}


def formula_blocks(ws, column: str, formula: str, rows: list[int]) -> None:
    col = column_index(column)
    for first, last in row_blocks(rows):
        ws.Range(ws.Cells(first, col), ws.Cells(last, col)).FormulaR1C1 = formula


def apply_formulas(ws, specs: list[tuple[str, str, str | None]], holiday_range: str) -> list[int]:
    headers = get_headers(ws)
    if "c_id" not in headers:
        raise ValueError(f"'{ws.Name}' sheet has no C_ID header in row 1.")

    last_row = ws.Cells(ws.Rows.Count, headers["c_id"]).End(XL_UP).Row
    if last_row < 2:
        print(f"  {ws.Name}: no data rows; skipped.")
        return []

    ids = as_column(
        ws.Range(ws.Cells(2, headers["c_id"]), ws.Cells(last_row, headers["c_id"])).Value2,
        last_row - 1,
    )
    data_rows = [row + 2 for row, value in enumerate(ids) if value is not None and str(value).strip()]
    if not data_rows:
        print(f"  {ws.Name}: no populated C_ID rows; skipped.")
        return []
    data_row_set = set(data_rows)

    for column, original_formula, date_column in specs:
        formula = original_formula.replace("'H-List'!C1:C1", holiday_range)
        if date_column:
            date_col = column_index(date_column)
            dates = as_column(
                ws.Range(ws.Cells(2, date_col), ws.Cells(last_row, date_col)).Value2,
                last_row - 1,
            )
            historic_rows, current_rows, blank_ce1_rows = [], [], []
            for offset, row in enumerate(range(2, last_row + 1)):
                if row not in data_row_set:
                    continue
                raw_date = dates[offset]
                if column in {"BT", "BM"} and (
                    raw_date is None
                    or not str(raw_date).strip()
                    or str(raw_date).strip().upper() in {"NA", "N/A"}
                ):
                    blank_ce1_rows.append(row)
                    continue
                try:
                    historic = raw_date is not None and float(raw_date) <= HISTORIC_CUTOFF
                except (TypeError, ValueError):
                    historic = False
                (historic_rows if historic else current_rows).append(row)
            if blank_ce1_rows:
                col_index = column_index(column)
                for first, last in row_blocks(blank_ce1_rows):
                    ws.Range(ws.Cells(first, col_index), ws.Cells(last, col_index)).ClearContents()
            formula_blocks(ws, column, formula, historic_rows if ",13,1," in formula else current_rows)
        else:
            formula_blocks(ws, column, formula, data_rows)

    print(f"  {ws.Name}: applied {len(specs)} formula rules to {len(data_rows):,} populated rows.")
    return data_rows


def freeze_results(ws, specs: list[tuple[str, str, str | None]], rows: list[int]) -> None:
    for column in {rule[0] for rule in specs}:
        col = column_index(column)
        for first, last in row_blocks(rows):
            value_range = ws.Range(ws.Cells(first, col), ws.Cells(last, col))
            # Let Excel freeze the formulas itself. Assigning Value2 back through
            # COM can turn CVErr(#VALUE!) into its raw integer HRESULT instead
            # of retaining the visible Excel error value.
            value_range.Copy()
            value_range.PasteSpecial(Paste=XL_PASTE_VALUES)
            ws.Application.CutCopyMode = False


def find_master(folder: Path) -> Path:
    matches = [
        path for path in folder.glob(MASTER_PATTERN)
        if path.is_file() and not path.name.startswith("~$")
        and "before_formulas_" not in path.stem.casefold()
    ]
    if not matches:
        raise FileNotFoundError(f"No Master Tracker workbook found in {folder}.")
    return max(matches, key=lambda path: path.stat().st_mtime)


def main() -> None:
    folder = Path(__file__).resolve().parent
    master_path = find_master(folder)
    print(f"Master Tracker: {master_path.name}")
    print("This will update formula columns in the tracker only; progress files are untouched.")
    print("Close the Master Tracker in Excel before continuing.")
    answer = input("Apply formulas and save their results as values? (Y/N): ").strip().upper()
    if answer != "Y":
        print("Cancelled; no workbook changes made.")
        return

    backup_folder = folder / "Tracker_Formula_Backups"
    backup_folder.mkdir(exist_ok=True)
    backup_path = backup_folder / (
        f"{master_path.stem}_before_formulas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )
    shutil.copy2(master_path, backup_path)
    print(f"Backup created: {backup_path}")

    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        workbook = excel.Workbooks.Open(
            str(master_path), UpdateLinks=0, ReadOnly=False, AddToMru=False
        )
        if workbook.ReadOnly:
            raise PermissionError("Master Tracker opened read-only. Close it in Excel and try again.")

        holiday_ws = workbook.Worksheets("H-List")
        holiday_last_row = max(1, holiday_ws.Cells(holiday_ws.Rows.Count, 1).End(XL_UP).Row)
        holiday_range = f"'H-List'!R1C1:R{holiday_last_row}C1"

        sheet_rows = {}
        for sheet_name, specs in SHEET_FORMULAS.items():
            ws = workbook.Worksheets(sheet_name)
            sheet_rows[sheet_name] = apply_formulas(ws, specs, holiday_range)

        print("Recalculating formulas in Excel...")
        excel.CalculateFullRebuild()
        for sheet_name, specs in SHEET_FORMULAS.items():
            freeze_results(workbook.Worksheets(sheet_name), specs, sheet_rows[sheet_name])

        workbook.Save()
        print("Done. Formula results have been saved as normal values in the Master Tracker.")
        print("No progress workbook was opened or changed.")
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        print(f"Failed: {error}")
    finally:
        input("Press Enter to close...")
