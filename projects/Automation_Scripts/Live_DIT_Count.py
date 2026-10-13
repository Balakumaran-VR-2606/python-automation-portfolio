import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import (
    get_column_letter,
    column_index_from_string
)
from datetime import datetime, date

# =====================================================
# FILES
# =====================================================

source_file = r"C:\Users\Balakumaran VR\Downloads\Infy_Conventional-Master Tracker.xlsx"
target_file = r"C:\Users\Balakumaran VR\Downloads\Case Validation.xlsx"

month_input = input(
    "\nEnter Month-Year (MMM-YYYY) : "
).strip()

target_month = datetime.strptime(
    month_input,
    "%b-%Y"
)

dashboard_input = input(
    "\nEnter Dashboard Date (DD-MMM-YYYY) : "
).strip()

dashboard_date = datetime.strptime(
    dashboard_input,
    "%d-%b-%Y"
).date()


es_count = 0
ce_count = 0
es_daywise = {}
ce_daywise = {}

dashboard = {

    "CID_FINAL": {
    "EDU": 0,
    "CRC": 0,
    "GDC": 0,
    "REINT_GDC": 0,
    "CE1": 0,
    "CE2": 0,
    "CE3": 0,
    "CE4": 0,
    "CE5": 0,
    "CE6": 0
},

    "EMP_FINAL": {
    "EDU": 0,
    "CRC": 0,
    "GDC": 0,
    "REINT_GDC": 0,
    "CE1": 0,
    "CE2": 0,
    "CE3": 0,
    "CE4": 0,
    "CE5": 0,
    "CE6": 0
},

    "CID_SUPPLEMENTARY": {
        "EDU": 0,
        "CRC": 0,
        "GDC": 0,
        "REINT_GDC": 0,

        "CE1": 0,
        "CE2": 0,
        "CE3": 0,
        "CE4": 0,
        "CE5": 0,
        "CE6": 0
    },

    "EMP_SUPPLEMENTARY": {
        "EDU": 0,
        "CRC": 0,
        "GDC": 0,
        "REINT_GDC": 0,

        "CE1": 0,
        "CE2": 0,
        "CE3": 0,
        "CE4": 0,
        "CE5": 0,
        "CE6": 0
    },

    "EMP_SUPPLEMENTARY2": {
        "EDU": 0,
        "CRC": 0,
        "GDC": 0,
        "REINT_GDC": 0,
        "CE1": 0,
        "CE2": 0,
        "CE3": 0,
        "CE4": 0,
        "CE5": 0,
        "CE6": 0
    },

    "PECVC_IR": {
        "EDU": 0,
        "CRC": 0,
        "GDC": 0,
        "CE2": 0,
        "CE3": 0,
        "CE4": 0,
        "CE5": 0,
        "CE6": 0
    },

    "PECVC_FINAL": {
        "EDU": 0,
        "CE1": 0,
        "CE2": 0,
        "CE3": 0,
        "CE4": 0,
        "CE5": 0,
        "CRC": 0,
        "GDC": 0
    },

    "PECVC_SUPP": {
        "CE1": 0
    },

    "PECVC_SUPP2": {
        "CE1": 0
    },

    "PECVC_SUPP3": {
        "CE1": 0
    },

    "FRESHER_CID": {
        "EDU": 0,
        "CRC": 0,
        "GDC": 0
    },

    "FRESHER_EMP": {
        "EDU": 0,
        "CRC": 0,
        "GDC": 0
    }
}

accepted_colors = {
    "GREEN",
    "RED",
    "ORANGE",
    "MAROON"
}

def is_valid_color(value):

    return (
        str(value).strip().upper()
        in accepted_colors
    )

def increment_count(
    section,
    check_name,
    value
):

    if is_valid_color(value):
        dashboard[section][check_name] += 1


processed_checks = set()
processed_checks_supp2 = set()
processed_checks_supp = set()

def increment_unique_count(
    section,
    check_name,
    result,
    unique_key
):
    if unique_key in processed_checks:
        return
    
    increment_count(
        section,
        check_name,
        result
    )

    processed_checks.add(unique_key)

def is_valid_color(value):

    return (
        str(value).strip().upper()
        in accepted_colors
    )

def increment_unique_count_supp2(
    section,
    check_name,
    result,
    unique_key
):
    if unique_key in processed_checks_supp2:
        return

    increment_count(
        section,
        check_name,
        result
    )

    processed_checks_supp2.add(unique_key)

def increment_unique_count_supp(
    section,
    check_name,
    result,
    unique_key
):
    if unique_key in processed_checks_supp:
        return

    increment_count(
        section,
        check_name,
        result
    )

    processed_checks_supp.add(unique_key)

# ES COUNT

for sheet in ["Pre-Joining", "PE & CVC", "Fresher"]:

    df = pd.read_excel(
        source_file,
        sheet_name=sheet,
        engine="openpyxl"
    )

    for value in df.iloc[:, 11]:   # L Column

        try:

            dt = pd.to_datetime(value)

            if (
                dt.month == target_month.month
                and dt.year == target_month.year
            ):
                es_count += 1

                day_key = dt.strftime("%d-%b-%Y")

                es_daywise[day_key] = (
                    es_daywise.get(day_key, 0) + 1
                )

        except:
            pass


# CE COUNT - PJ (BS)

df_pj = pd.read_excel(
    source_file,
    sheet_name="Pre-Joining",
    engine="openpyxl"
)

for value in df_pj.iloc[:, 70]:

    try:

        dt = pd.to_datetime(value)

        if (
            dt.month == target_month.month
            and dt.year == target_month.year
        ):
            ce_count += 1

            day_key = dt.strftime("%d-%b-%Y")

            ce_daywise[day_key] = (
                ce_daywise.get(day_key, 0) + 1
            )

    except:
        pass


# CE COUNT - PE & CVC (BL)

df_pe = pd.read_excel(
    source_file,
    sheet_name="PE & CVC",
    engine="openpyxl"
)

for value in df_pe.iloc[:, 63]:

    try:

        dt = pd.to_datetime(value)

        if (
            dt.month == target_month.month
            and dt.year == target_month.year
        ):
            ce_count += 1

            day_key = dt.strftime("%d-%b-%Y")

            ce_daywise[day_key] = (
                ce_daywise.get(day_key, 0) + 1
            )

    except:
        pass

# CE COUNT - FRESHER (AS)

df_fresher = pd.read_excel(
    source_file,
    sheet_name="Fresher",
    engine="openpyxl"
)

for value in df_fresher.iloc[:, 44]:   # AS

    try:

        dt = pd.to_datetime(value)

        if (
            dt.month == target_month.month
            and dt.year == target_month.year
        ):
            ce_count += 1

            day_key = dt.strftime("%d-%b-%Y")

            ce_daywise[day_key] = (
                ce_daywise.get(day_key, 0) + 1
            )

    except:
        pass

# =====================================================
# PRODUCTIVITY DASHBOARD - PRE JOINING
# =====================================================

df_pj_dashboard = pd.read_excel(
    source_file,
    sheet_name="Pre-Joining",
    engine="openpyxl"
)

df_fresher_dashboard = pd.read_excel(
    source_file,
    sheet_name="Fresher",
    engine="openpyxl"
)

df_pecvc_dashboard = pd.read_excel(
    source_file,
    sheet_name="PE & CVC",
    engine="openpyxl"
)

df_pecvc_dashboard.columns = (
    df_pecvc_dashboard.columns.str.strip()
)

pj_previous_checks = set()

for _, row in df_pj_dashboard.iterrows():

    fr_can = pd.to_datetime(
        row.get("FR-CAN ID_Date"),
        errors="coerce"
    )
    
    if (
        pd.notna(fr_can)
        and fr_can.date() < dashboard_date
        and str(row.get("Report Type", "")).strip().upper() == "FINAL"
    ):

        # EDU
        if pd.notna(row.get("Ed_Result")):

            pj_previous_checks.add((
                str(row.get("C_ID")),
                "EDU",
                str(row.get("Ed_SDate")),
                str(row.get("Ed_Edate")),
                str(row.get("Ed_Result"))
            ))

        # CRC
        if pd.notna(row.get("eP_CRC_Result")):

            pj_previous_checks.add((
                str(row.get("C_ID")),
                "CRC",
                str(row.get("eP_CRC_SDate")),
                str(row.get("eP_CRC_EDate")),
                str(row.get("eP_CRC_Result"))
            ))

        # GDC
        if pd.notna(row.get("GDB_Result")):

            pj_previous_checks.add((
                str(row.get("C_ID")),
                "GDC",
                str(row.get("GDB_SDate")),
                str(row.get("GDB_EDate")),
                str(row.get("GDB_Result"))
            ))

        # CE1

        ce_status = str(
            row.get("CE_Status", "")
        ).strip().upper()
        
        if pd.notna(row.get("CE1_Result")):

            pj_previous_checks.add((
                str(row.get("C_ID")),
                "CE1",
                str(row.get("CE1_Sdate")),
                str(row.get("CE1_Edate")),
                str(row.get("CE1_Result"))
            ))

        
        # CE2
        pj_previous_checks.add((
            str(row.get("C_ID")),
            "CE2",
            str(row.get("E2_Sdate")),
            str(row.get("E2_EDate")),
            str(row.get("E2_Result"))
        ))

        # CE3
        pj_previous_checks.add((
            str(row.get("C_ID")),
            "CE3",
            str(row.get("E3_Sdate")),
            str(row.get("E3_EDate")),
            str(row.get("E3_Result"))
        ))

        # CE4
        pj_previous_checks.add((
            str(row.get("C_ID")),
            "CE4",
            str(row.get("E4_Sdate")),
            str(row.get("E4_EDate")),
            str(row.get("E4_Result"))
        ))

        # CE5
        pj_previous_checks.add((
            str(row.get("C_ID")),
            "CE5",
            str(row.get("E5_Sdate")),
            str(row.get("E5_EDate")),
            str(row.get("E5_Result"))
        ))

        # CE6
        pj_previous_checks.add((
            str(row.get("C_ID")),
            "CE6",
            str(row.get("E6_Sdate")),
            str(row.get("E6_EDate")),
            str(row.get("E6_Result"))
        ))

for _, row in df_pecvc_dashboard.iterrows():

    fr_date = pd.to_datetime(
        row.get("FR_Date"),
        errors="coerce"
    )

    if (
        pd.notna(fr_date)
        and fr_date.date() < dashboard_date
    ):

        gdc_key = (
            str(row.get("Emp_No") or row.get("C_ID")),
            "GDC",
            str(row.get("GDB_SDate")),
            str(row.get("GDB_EDate")),
            str(row.get("GDB_Result"))
        )

        processed_checks.add(gdc_key)

        edu_key = (
            str(row.get("Emp_No") or row.get("C_ID")),
            "EDU",
            str(row.get("Ed_SDate")),
            str(row.get("Ed_Edate")),
            str(row.get("Ed_Result"))
        )

        processed_checks.add(edu_key)

        crc_key = (
            str(row.get("Emp_No") or row.get("C_ID")),
            "CRC",
            str(row.get("eP_CRC_SDate")),
            str(row.get("eP_CRC_EDate")),
            str(row.get("eP_CRC_Result"))
        )

        processed_checks.add(crc_key)

        ce1_key = (
            str(row.get("Emp_No") or row.get("C_ID")),
            "CE1",
            str(row.get("CE1_Sdate")),
            str(row.get("CE1_Edate")),
            str(row.get("CE1_Result"))
        )
        processed_checks.add(ce1_key)

        ce2_key = (
           str(row.get("Emp_No") or row.get("C_ID")),
           "CE2",
           str(row.get("E2_Sdate")),
           str(row.get("E2_EDate")),
           str(row.get("E2_Result"))
        )
        processed_checks.add(ce2_key)

        ce3_key = (
            str(row.get("Emp_No") or row.get("C_ID")),
            "CE3",
            str(row.get("E3_Sdate")),
            str(row.get("E3_EDate")),
            str(row.get("E3_Result"))
        )
        processed_checks.add(ce3_key)

        ce4_key = (
            str(row.get("Emp_No") or row.get("C_ID")),
            "CE4",
            str(row.get("E4_Sdate")),
            str(row.get("E4_EDate")),
            str(row.get("E4_Result"))
        )
        processed_checks.add(ce4_key)

        ce5_key = (
            str(row.get("Emp_No") or row.get("C_ID")),
            "CE5",
            str(row.get("E5_Sdate")),
            str(row.get("E5_EDate")),
            str(row.get("E5_Result"))
        )
        processed_checks.add(ce5_key)

for _, row in df_pj_dashboard.iterrows():

    fr_can = pd.to_datetime(
        row.get("FR-CAN ID_Date"),
        errors="coerce"
    )

    fr_emp = pd.to_datetime(
        row.get("FR-EMP ID_Date"),
        errors="coerce"
    )

    sr_can = pd.to_datetime(
        row.get("SR-CAN ID_Date"),
        errors="coerce"
    )

    sr2_can = pd.to_datetime(
        row.get("SR2-CAN ID_Date"),
        errors="coerce"
    )

    sr3_can = pd.to_datetime(
        row.get("SR3-CAN ID_Date"),
        errors="coerce"
    )

    sr_emp = pd.to_datetime(
        row.get("SR-EMP ID_Date"),
        errors="coerce"
    )

    sr2_emp = pd.to_datetime(
        row.get("SR2-EMP ID_Date"),
        errors="coerce"
    )

    sr3_emp = pd.to_datetime(
        row.get("SR3-EMP ID_Date"),
        errors="coerce"
    )

    if (
        pd.notna(sr2_emp)
        and sr2_emp.date() == dashboard_date
    ):
        print(
            "CHECK",
            row.get("C_ID"),
            "FR_EMP=",
            fr_emp,
            "SR2_EMP=",
            sr2_emp
        )

    fr_can_selected = (
        pd.notna(fr_can)
        and fr_can.date() == dashboard_date
    )

    fr_emp_selected = (
        pd.notna(fr_emp)
        and fr_emp.date() == dashboard_date
    )
      
    # =========================
    # CID FINAL
    # =========================

    if (
        pd.notna(fr_can)
        and fr_can.date() == dashboard_date
    ):

        edu_key = (
            str(row.get("C_ID")),
            "EDU",
            str(row.get("Ed_SDate")),
            str(row.get("Ed_Edate")),
            str(row.get("Ed_Result"))
        )

        if edu_key not in pj_previous_checks:
            increment_count(
                "CID_FINAL",
                "EDU",
                row.get("Ed_Result")
            )

        crc_key = (
            str(row.get("C_ID")),
            "CRC",
            str(row.get("eP_CRC_SDate")),
            str(row.get("eP_CRC_EDate")),
            str(row.get("eP_CRC_Result"))
        )

        if crc_key not in pj_previous_checks:
            increment_count(
                "CID_FINAL",
                "CRC",
                row.get("eP_CRC_Result")
            )

        gdc_key = (
            str(row.get("C_ID")),
            "GDC",
            str(row.get("GDB_SDate")),
            str(row.get("GDB_EDate")),
            str(row.get("GDB_Result"))
        )

        if gdc_key not in pj_previous_checks:
            increment_count(
                "CID_FINAL",
                "GDC",
                row.get("GDB_Result")
            )

                # REINT_GDC
        reint_key = (
            str(row.get("C_ID")),
            "REINT_GDC",
            str(row.get("ReINT_GDB_SDate")),
            str(row.get("ReINT_GDB_EDate")),
            str(row.get("ReINT_GDB_Result"))
        )

        if reint_key not in pj_previous_checks:
            increment_count(
                "CID_FINAL",
                "REINT_GDC",
                row.get("ReINT_GDB_Result")
            )

        # CE1
        ce_status = str(
            row.get("CE_Status", "")
        ).strip().upper()

        if ce_status == "NO":

            ce1_key = (
                str(row.get("C_ID")),
                "CE1",
                str(row.get("CE1_Sdate")),
                str(row.get("CE1_Edate")),
                str(row.get("CE1_Result"))
            )
            
            if ce1_key not in pj_previous_checks:
                increment_count(
                    "CID_FINAL",
                    "CE1",
                    row.get("CE1_Result")
                )

        # CE2
        ce2_key = (
            str(row.get("C_ID")),
            "CE2",
            str(row.get("E2_Sdate")),
            str(row.get("E2_EDate")),
            str(row.get("E2_Result"))
        )

        if ce2_key not in pj_previous_checks:
            increment_count(
                "CID_FINAL",
                "CE2",
                row.get("E2_Result")
            )

        # CE3
        ce3_key = (
            str(row.get("C_ID")),
            "CE3",
            str(row.get("E3_Sdate")),
            str(row.get("E3_EDate")),
            str(row.get("E3_Result"))
        )

        if ce3_key not in pj_previous_checks:
            increment_count(
                "CID_FINAL",
                "CE3",
                row.get("E3_Result")
            )

        # CE4
        ce4_key = (
            str(row.get("C_ID")),
            "CE4",
            str(row.get("E4_Sdate")),
            str(row.get("E4_EDate")),
            str(row.get("E4_Result"))
        )

        if ce4_key not in pj_previous_checks:
            increment_count(
                "CID_FINAL",
                "CE4",
                row.get("E4_Result")
            )

        # CE5
        ce5_key = (
            str(row.get("C_ID")),
            "CE5",
            str(row.get("E5_Sdate")),
            str(row.get("E5_EDate")),
            str(row.get("E5_Result"))
        )

        if ce5_key not in pj_previous_checks:
            increment_count(
                "CID_FINAL",
                "CE5",
                row.get("E5_Result")
            )

        # CE6
        ce6_key = (
            str(row.get("C_ID")),
            "CE6",
            str(row.get("E6_Sdate")),
            str(row.get("E6_EDate")),
            str(row.get("E6_Result"))
        )

        if ce6_key not in pj_previous_checks:
            increment_count(
                "CID_FINAL",
                "CE6",
                row.get("E6_Result")
            )

    # OLD FR-CAN + CURRENT FR-EMP

    if (
        fr_emp_selected
        and pd.notna(fr_can)
        and fr_can.date() != dashboard_date
    ):

        # REINT_GDC

        reint_key = (
            str(row.get("C_ID")),
            "REINT_GDC",
            str(row.get("ReINT_GDB_SDate")),
            str(row.get("ReINT_GDB_EDate")),
            str(row.get("ReINT_GDB_Result"))
        )

        if reint_key not in pj_previous_checks:

            increment_count(
                "EMP_FINAL",
                "REINT_GDC",
                row.get("ReINT_GDB_Result")
            )

        ce_status = str(
            row.get("CE_Status", "")
        ).strip().upper()

        if ce_status in ["YES", "POB"]:

            ce1_key = (
                str(row.get("C_ID")),
                "CE1",
                str(row.get("CE1_Sdate")),
                str(row.get("CE1_Edate")),
                str(row.get("CE1_Result"))
            )

            if ce1_key not in processed_checks:

                increment_count(
                    "EMP_FINAL",
                    "CE1",
                    row.get("CE1_Result")
                )

                processed_checks.add(ce1_key)

    # FR-CAN BLANK + FR-EMP CURRENT

    if (
        fr_emp_selected
        and pd.isna(fr_can)
    ):

        edu_key = (
            str(row.get("C_ID")),
            "EDU",
            str(row.get("Ed_SDate")),
            str(row.get("Ed_Edate")),
            str(row.get("Ed_Result"))
        )

        if edu_key not in pj_previous_checks:
            increment_count(
                "EMP_FINAL",
                "EDU",
                row.get("Ed_Result")
            )

        crc_key = (
            str(row.get("C_ID")),
            "CRC",
            str(row.get("eP_CRC_SDate")),
            str(row.get("eP_CRC_EDate")),
            str(row.get("eP_CRC_Result"))
        )

        if crc_key not in pj_previous_checks:
            increment_count(
                "EMP_FINAL",
                "CRC",
                row.get("eP_CRC_Result")
            )

        gdc_key = (
            str(row.get("C_ID")),
            "GDC",
            str(row.get("GDB_SDate")),
            str(row.get("GDB_EDate")),
            str(row.get("GDB_Result"))
        )

        if gdc_key not in pj_previous_checks:
            increment_count(
                "EMP_FINAL",
                "GDC",
                row.get("GDB_Result")
            )

        reint_key = (
            str(row.get("C_ID")),
            "REINT_GDC",
            str(row.get("ReINT_GDB_SDate")),
            str(row.get("ReINT_GDB_EDate")),
            str(row.get("ReINT_GDB_Result"))
        )

        if reint_key not in pj_previous_checks:
            increment_count(
                "EMP_FINAL",
                "REINT_GDC",
                row.get("ReINT_GDB_Result")
            )

        ce1_key = (
            str(row.get("C_ID")),
            "CE1",
            str(row.get("CE1_Sdate")),
            str(row.get("CE1_Edate")),
            str(row.get("CE1_Result"))
        )
        
        if ce1_key not in pj_previous_checks:
            increment_count(
                "EMP_FINAL",
                "CE1",
                row.get("CE1_Result")
            )

        ce2_key = (
            str(row.get("C_ID")),
            "CE2",
            str(row.get("E2_Sdate")),
            str(row.get("E2_EDate")),
            str(row.get("E2_Result"))
        )

        if ce2_key not in pj_previous_checks:
            increment_count(
                "EMP_FINAL",
                "CE2",
                row.get("E2_Result")
            )

        ce3_key = (
            str(row.get("C_ID")),
            "CE3",
            str(row.get("E3_Sdate")),
            str(row.get("E3_EDate")),
            str(row.get("E3_Result"))
        )

        if ce3_key not in pj_previous_checks:
            increment_count(
                "EMP_FINAL",
                "CE3",
                row.get("E3_Result")
            )

        ce4_key = (
            str(row.get("C_ID")),
            "CE4",
            str(row.get("E4_Sdate")),
            str(row.get("E4_EDate")),
            str(row.get("E4_Result"))
        )

        if ce4_key not in pj_previous_checks:
            increment_count(
                "EMP_FINAL",
                "CE4",
                row.get("E4_Result")
            )

        ce5_key = (
            str(row.get("C_ID")),
            "CE5",
            str(row.get("E5_Sdate")),
            str(row.get("E5_EDate")),
            str(row.get("E5_Result"))
        )

        if ce5_key not in pj_previous_checks:
            increment_count(
                "EMP_FINAL",
                "CE5",
                row.get("E5_Result")
            )

        ce6_key = (
            str(row.get("C_ID")),
            "CE6",
            str(row.get("E6_Sdate")),
            str(row.get("E6_EDate")),
            str(row.get("E6_Result"))
        )

        if ce6_key not in pj_previous_checks:
            increment_count(
                "EMP_FINAL",
                "CE6",
                row.get("E6_Result")
            )

    # CID SUPPLEMENTARY + EMP SUPPLEMENTARY
    # FR-CAN BLANK
    # SR-CAN / SR2-CAN / SR3-CAN = Dashboard Date

    if (
        pd.isna(fr_can)
        and (
            (
                pd.notna(sr_can)
                and sr_can.date() == dashboard_date
            )
            or
            (
                pd.notna(sr2_can)
                and sr2_can.date() == dashboard_date
            )
            or
            (
                pd.notna(sr3_can)
                and sr3_can.date() == dashboard_date
            )
        )
    ):

        # -----------------------
        # CID SUPPLEMENTARY
        # -----------------------

        # EDU
        edu_key = (
            str(row.get("C_ID")),
            "EDU",
            str(row.get("Ed_SDate")),
            str(row.get("Ed_Edate")),
            str(row.get("Ed_Result"))
        )

        if edu_key not in pj_previous_checks:
            increment_count(
                "CID_SUPPLEMENTARY",
                "EDU",
                row.get("Ed_Result")
            )

        # CRC
        crc_key = (
            str(row.get("C_ID")),
            "CRC",
            str(row.get("eP_CRC_SDate")),
            str(row.get("eP_CRC_EDate")),
            str(row.get("eP_CRC_Result"))
        )

        if crc_key not in pj_previous_checks:
            increment_count(
                "CID_SUPPLEMENTARY",
                "CRC",
                row.get("eP_CRC_Result")
            )

        # GDC
        gdc_key = (
            str(row.get("C_ID")),
            "GDC",
            str(row.get("GDB_SDate")),
            str(row.get("GDB_EDate")),
            str(row.get("GDB_Result"))
        )

        if gdc_key not in pj_previous_checks:
            increment_count(
                "CID_SUPPLEMENTARY",
                "GDC",
                row.get("GDB_Result")
            )

        # REINT_GDC
        reint_key = (
            str(row.get("C_ID")),
            "REINT_GDC",
            str(row.get("ReINT_GDB_SDate")),
            str(row.get("ReINT_GDB_EDate")),
            str(row.get("ReINT_GDB_Result"))
        )

        if reint_key not in pj_previous_checks:
            increment_count(
                "CID_SUPPLEMENTARY",
                "REINT_GDC",
                row.get("ReINT_GDB_Result")
            )

        # CE1
        ce1_key = (
            str(row.get("C_ID")),
            "CE1",
            str(row.get("CE1_Sdate")),
            str(row.get("CE1_Edate")),
            str(row.get("CE1_Result"))
        )

        if ce1_key not in pj_previous_checks:
            increment_count(
                "CID_SUPPLEMENTARY",
                "CE1",
                row.get("CE1_Result")
            )

        # CE2
        ce2_key = (
            str(row.get("C_ID")),
            "CE2",
            str(row.get("E2_Sdate")),
            str(row.get("E2_EDate")),
            str(row.get("E2_Result"))
        )

        if ce2_key not in pj_previous_checks:
            increment_count(
                "CID_SUPPLEMENTARY",
                "CE2",
                row.get("E2_Result")
            )

        # CE3
        ce3_key = (
            str(row.get("C_ID")),
            "CE3",
            str(row.get("E3_Sdate")),
            str(row.get("E3_EDate")),
            str(row.get("E3_Result"))
        )

        if ce3_key not in pj_previous_checks:
            increment_count(
                "CID_SUPPLEMENTARY",
                "CE3",
                row.get("E3_Result")
            )

        # CE4
        ce4_key = (
            str(row.get("C_ID")),
            "CE4",
            str(row.get("E4_Sdate")),
            str(row.get("E4_EDate")),
            str(row.get("E4_Result"))
        )

        if ce4_key not in pj_previous_checks:
            increment_count(
                "CID_SUPPLEMENTARY",
                "CE4",
                row.get("E4_Result")
            )

        # CE5
        ce5_key = (
            str(row.get("C_ID")),
            "CE5",
            str(row.get("E5_Sdate")),
            str(row.get("E5_EDate")),
            str(row.get("E5_Result"))
        )

        if ce5_key not in pj_previous_checks:
            increment_count(
                "CID_SUPPLEMENTARY",
                "CE5",
                row.get("E5_Result")
            )

        # CE6
        ce6_key = (
            str(row.get("C_ID")),
            "CE6",
            str(row.get("E6_Sdate")),
            str(row.get("E6_EDate")),
            str(row.get("E6_Result"))
        )

        if ce6_key not in pj_previous_checks:
            increment_count(
                "CID_SUPPLEMENTARY",
                "CE6",
                row.get("E6_Result")
            )

    # EMP SUPPLEMENTARY
    # FR-EMP BLANK
    # SR-EMP / SR2-EMP / SR3-EMP = Dashboard Date

    if (
        pd.isna(fr_emp)
        and (
            (
                pd.notna(sr_emp)
                and sr_emp.date() == dashboard_date
            )
            or
            (
                pd.notna(sr2_emp)
                and sr2_emp.date() == dashboard_date
            )
            or
            (
                pd.notna(sr3_emp)
                and sr3_emp.date() == dashboard_date
            )
        )
    ):
            
            edu_key = (
                str(row.get("C_ID")),
                "EDU",
                str(row.get("Ed_SDate")),
                str(row.get("Ed_Edate")),
                str(row.get("Ed_Result"))
            )

            if edu_key not in pj_previous_checks:
                increment_unique_count_supp(
                    "EMP_SUPPLEMENTARY",
                    "EDU",
                    row.get("Ed_Result"),
                    edu_key
                )

            crc_key = (
                str(row.get("C_ID")),
                "CRC",
                str(row.get("eP_CRC_SDate")),
                str(row.get("eP_CRC_EDate")),
                str(row.get("eP_CRC_Result"))
            )

            if crc_key not in pj_previous_checks:
                increment_unique_count_supp(
                    "EMP_SUPPLEMENTARY",
                    "CRC",
                    row.get("eP_CRC_Result"),
                    crc_key
                )

            gdc_key = (
                str(row.get("C_ID")),
                "GDC",
                str(row.get("GDB_SDate")),
                str(row.get("GDB_EDate")),
                str(row.get("GDB_Result"))
            )

            if gdc_key not in pj_previous_checks:
                increment_unique_count_supp(
                    "EMP_SUPPLEMENTARY",
                    "GDC",
                    row.get("GDB_Result"),
                    gdc_key
                )

            reint_key = (
                str(row.get("C_ID")),
                "REINT_GDC",
                str(row.get("ReINT_GDB_SDate")),
                str(row.get("ReINT_GDB_EDate")),
                str(row.get("ReINT_GDB_Result"))
            )

            if reint_key not in pj_previous_checks:
                increment_unique_count_supp(
                    "EMP_SUPPLEMENTARY",
                    "REINT_GDC",
                    row.get("ReINT_GDB_Result"),
                    reint_key
                )

            ce1_key = (
                str(row.get("C_ID")),
                "CE1",
                str(row.get("CE1_Sdate")),
                str(row.get("CE1_Edate")),
                str(row.get("CE1_Result"))
            )

            print(
                "CE1 KEY",
                ce1_key,
                "EXISTS=",
                ce1_key in pj_previous_checks
            )

            if ce1_key not in pj_previous_checks:
                increment_unique_count_supp(
                    "EMP_SUPPLEMENTARY",
                    "CE1",
                    row.get("CE1_Result"),
                    ce1_key
                )

            ce2_key = (
                str(row.get("C_ID")),
                "CE2",
                str(row.get("E2_Sdate")),
                str(row.get("E2_EDate")),
                str(row.get("E2_Result"))
            )

            if ce2_key not in pj_previous_checks:
                increment_unique_count_supp(
                    "EMP_SUPPLEMENTARY",
                    "CE2",
                    row.get("E2_Result"),
                    ce2_key
                )

            ce3_key = (
                str(row.get("C_ID")),
                "CE3",
                str(row.get("E3_Sdate")),
                str(row.get("E3_EDate")),
                str(row.get("E3_Result"))
            )

            if ce3_key not in pj_previous_checks:
                increment_unique_count_supp(
                    "EMP_SUPPLEMENTARY",
                    "CE3",
                    row.get("E3_Result"),
                    ce3_key
                )

            ce4_key = (
                str(row.get("C_ID")),
                "CE4",
                str(row.get("E4_Sdate")),
                str(row.get("E4_EDate")),
                str(row.get("E4_Result"))
            )

            if ce4_key not in pj_previous_checks:
                increment_unique_count_supp(
                    "EMP_SUPPLEMENTARY",
                    "CE4",
                    row.get("E4_Result"),
                    ce4_key
                )

            ce5_key = (
                str(row.get("C_ID")),
                "CE5",
                str(row.get("E5_Sdate")),
                str(row.get("E5_EDate")),
                str(row.get("E5_Result"))
            )

            if ce5_key not in pj_previous_checks:
                increment_unique_count_supp(
                    "EMP_SUPPLEMENTARY",
                    "CE5",
                    row.get("E5_Result"),
                    ce5_key
                )

            ce6_key = (
                str(row.get("C_ID")),
                "CE6",
                str(row.get("E6_Sdate")),
                str(row.get("E6_EDate")),
                str(row.get("E6_Result"))
            )

            if ce6_key not in pj_previous_checks:
                increment_unique_count_supp(
                    "EMP_SUPPLEMENTARY",
                    "CE6",
                    row.get("E6_Result"),
                    ce6_key
                )

            # EMP SUPPLEMENTARY -2

            if (
                pd.notna(sr2_emp)
                and sr2_emp.date() == dashboard_date
            ):

                print(
                    "SUPP2 HIT",
                    row.get("C_ID")
                )

                increment_unique_count_supp2(
                    "EMP_SUPPLEMENTARY2",
                    "EDU",
                    row.get("Ed_Result"),
                    edu_key
                )

                increment_unique_count_supp2(
                    "EMP_SUPPLEMENTARY2",
                    "CRC",
                    row.get("eP_CRC_Result"),
                    crc_key
                )

                increment_unique_count_supp2(
                    "EMP_SUPPLEMENTARY2",
                    "GDC",
                    row.get("GDB_Result"),
                    gdc_key
                )

                increment_unique_count_supp2(
                    "EMP_SUPPLEMENTARY2",
                    "REINT_GDC",
                    row.get("ReINT_GDB_Result"),
                    reint_key
                )

                increment_unique_count_supp2(
                    "EMP_SUPPLEMENTARY2",
                    "CE1",
                    row.get("CE1_Result"),
                    ce1_key
                )

                increment_unique_count_supp2(
                    "EMP_SUPPLEMENTARY2",
                    "CE2",
                    row.get("E2_Result"),
                    ce2_key
                )

                increment_unique_count_supp2(
                    "EMP_SUPPLEMENTARY2",
                    "CE3",
                    row.get("E3_Result"),
                    ce3_key
                )

                increment_unique_count_supp2(
                    "EMP_SUPPLEMENTARY2",
                    "CE4",
                    row.get("E4_Result"),
                    ce4_key
                )

                increment_unique_count_supp2(
                    "EMP_SUPPLEMENTARY2",
                    "CE5",
                    row.get("E5_Result"),
                    ce5_key
                )

                increment_unique_count_supp2(
                    "EMP_SUPPLEMENTARY2",
                    "CE6",
                    row.get("E6_Result"),
                    ce6_key
                )    

# FRESHER LOGIC

for _, row in df_fresher_dashboard.iterrows():

    fr_can = pd.to_datetime(
        row.get("FR-CAN ID_Date"),
        errors="coerce"
    )

    fr_emp = pd.to_datetime(
        row.get("FR-EMP ID_Date"),
        errors="coerce"
    )

    sr_can = pd.to_datetime(
        row.get("SR-CAN ID_Date"),
        errors="coerce"
    )

    sr2_can = pd.to_datetime(
        row.get("SR2-CAN ID_Date"),
        errors="coerce"
    )

    sr3_can = pd.to_datetime(
        row.get("SR3-CAN ID_Date"),
        errors="coerce"
    )

    sr_emp = pd.to_datetime(
        row.get("SR-EMP ID_Date"),
        errors="coerce"
    )

    sr2_emp = pd.to_datetime(
        row.get("SR2-EMP ID_Date"),
        errors="coerce"
    )

    if pd.notna(sr2_emp):
        print(
            "SR2 EMP FOUND",
            row.get("C_ID"),
            sr2_emp
        )

    sr3_emp = pd.to_datetime(
        row.get("SR3-EMP ID_Date"),
        errors="coerce"
    )

    if (
        pd.notna(fr_can)
        and fr_can.date() == dashboard_date
    ):
        
        increment_count(
            "FRESHER_CID",
            "EDU",
            row.get("Ed_Result")
        )

        increment_count(
            "FRESHER_CID",
            "CRC",
            row.get("eP_CRC_Result")
        )

        increment_count(
            "FRESHER_CID",
            "GDC",
            row.get("GDB_Result")
        )

    if (
        pd.notna(fr_emp)
        and fr_emp.date() == dashboard_date
        and pd.isna(fr_can)
    ):
        
        increment_count(
            "FRESHER_EMP",
            "EDU",
            row.get("Ed_Result")
        )

        increment_count(
            "FRESHER_EMP",
            "CRC",
            row.get("eP_CRC_Result")
        )

        increment_count(
            "FRESHER_EMP",
            "GDC",
            row.get("GDB_Result")
        )

# CID SUPPLEMENTARY

    if (
        pd.isna(fr_can)
        and (
            (
                pd.notna(sr_can)
                and sr_can.date() == dashboard_date
            )
            or
            (
                pd.notna(sr2_can)
                and sr2_can.date() == dashboard_date
            )
            or
            (
                pd.notna(sr3_can)
                and sr3_can.date() == dashboard_date
            )
        )
    ):

        increment_count(
            "CID_SUPPLEMENTARY",
            "EDU",
            row.get("Ed_Result")
        )

        increment_count(
            "CID_SUPPLEMENTARY",
            "CRC",
            row.get("eP_CRC_Result")
        )

        increment_count(
            "CID_SUPPLEMENTARY",
            "GDC",
            row.get("GDB_Result")
        )

        increment_count(
            "CID_SUPPLEMENTARY",
            "REINT_GDC",
            row.get("ReINT_GDB_Result")
        )

        increment_count(
            "EMP_SUPPLEMENTARY",
            "CE1",
            row.get("CE1_Result")
        )

        increment_count(
            "EMP_SUPPLEMENTARY",
            "CE2",
            row.get("E2_Result")
        )

        increment_count(
            "EMP_SUPPLEMENTARY",
            "CE3",
            row.get("E3_Result")
        )

        increment_count(
            "EMP_SUPPLEMENTARY",
            "CE4",
            row.get("E4_Result")
        )

        increment_count(
            "EMP_SUPPLEMENTARY",
            "CE5",
            row.get("E5_Result")
        )

        increment_count(
            "EMP_SUPPLEMENTARY",
            "CE6",
            row.get("E6_Result")
        )


# PE & CVC LOGIC

for _, row in df_pecvc_dashboard.iterrows():

    ir_date = pd.to_datetime(
        row.get("IR_Date"),
        errors="coerce"
    )

    fr_date = pd.to_datetime(
        row.get("FR_Date"),
        errors="coerce"
    )

    ir_selected = (
        pd.notna(ir_date)
        and ir_date.date() == dashboard_date
    )

    fr_selected = (
        pd.notna(fr_date)
        and fr_date.date() == dashboard_date
    )

    # CASE 1
    # IR_Date = Dashboard Date
    # FR_Date = Blank
    # Full Count (Except CE1)

    if ir_selected:

        increment_count(
            "PECVC_IR",
            "EDU",
            row.get("Ed_Result")
        )

        increment_count(
            "PECVC_IR",
            "CRC",
            row.get("eP_CRC_Result")
        )

        increment_count(
            "PECVC_IR",
            "GDC",
            row.get("GDB_Result")
        )

        increment_count(
            "PECVC_IR",
            "CE2",
            row.get("E2_Result")
        )

        increment_count(
            "PECVC_IR",
            "CE3",
            row.get("E3_Result")
        )

        increment_count(
            "PECVC_IR",
            "CE4",
            row.get("E4_Result")
        )

        increment_count(
            "PECVC_IR",
            "CE5",
            row.get("E5_Result")
        )

        increment_count(
            "PECVC_IR",
            "CE6",
            row.get("E6_Result")
        )

    # CASE 2
    # IR Blank + FR Current
    # Full Count

    if (
        fr_selected
        and pd.isna(ir_date)
    ):

        unique_key = (
            str(row.get("Emp_No") or row.get("C_ID")),
            "EDU",
            str(row.get("Ed_SDate")),
            str(row.get("Ed_Edate")),
            str(row.get("Ed_Result"))
        )

        increment_unique_count(
            "PECVC_FINAL",
            "EDU",
            row.get("Ed_Result"),
            unique_key
        )

        crc_key = (
            str(row.get("Emp_No") or row.get("C_ID")),
            "CRC",
            str(row.get("eP_CRC_SDate")),
            str(row.get("eP_CRC_EDate")),
            str(row.get("eP_CRC_Result"))
        )

        increment_unique_count(
            "PECVC_FINAL",
            "CRC",
            row.get("eP_CRC_Result"),
            crc_key
        )

        unique_key = (
            str(row.get("Emp_No") or row.get("C_ID")),
            "GDC",
            str(row.get("GDB_SDate")),
            str(row.get("GDB_EDate")),
            str(row.get("GDB_Result"))
        )
        
        increment_unique_count(
            "PECVC_FINAL",
            "GDC",
            row.get("GDB_Result"),
            unique_key
        )

        ce1_key = (
            str(row.get("Emp_No") or row.get("C_ID")),
            "CE1",
            str(row.get("CE1_Sdate")),
            str(row.get("CE1_Edate")),
            str(row.get("CE1_Result"))
        )

        increment_unique_count(
            "PECVC_FINAL",
            "CE1",
            row.get("CE1_Result"),
            ce1_key
        )

        ce2_key = (
            str(row.get("Emp_No") or row.get("C_ID")),
            "CE2",
            str(row.get("E2_Sdate")),
            str(row.get("E2_EDate")),
            str(row.get("E2_Result"))
        )

        increment_unique_count(
            "PECVC_FINAL",
            "CE2",
            row.get("E2_Result"),
            ce2_key
        )

        ce3_key = (
            str(row.get("Emp_No") or row.get("C_ID")),
            "CE3",
            str(row.get("E3_Sdate")),
            str(row.get("E3_EDate")),
            str(row.get("E3_Result"))
        )

        increment_unique_count(
            "PECVC_FINAL",
            "CE3",
            row.get("E3_Result"),
            ce3_key
        )

        ce4_key = (
            str(row.get("Emp_No") or row.get("C_ID")),
            "CE4",
            str(row.get("E4_Sdate")),
            str(row.get("E4_EDate")),
            str(row.get("E4_Result"))
        )

        increment_unique_count(
            "PECVC_FINAL",
            "CE4",
            row.get("E4_Result"),
            ce4_key
        )

        ce5_key = (
            str(row.get("Emp_No") or row.get("C_ID")),
            "CE5",
            str(row.get("E5_Sdate")),
            str(row.get("E5_EDate")),
            str(row.get("E5_Result"))
        )

        increment_unique_count(
            "PECVC_FINAL",
            "CE5",
            row.get("E5_Result"),
            ce5_key
        )

        ce6_key = (
            str(row.get("Emp_No") or row.get("C_ID")),
            "CE6",
            str(row.get("E6_Sdate")),
            str(row.get("E6_EDate")),
            str(row.get("E6_Result"))
        )

        increment_unique_count(
            "PECVC_FINAL",
            "CE6",
            row.get("E6_Result"),
            ce6_key
        )

    # CASE 3
    # IR Old Date + FR Current Date
    # CE1 Only

    if (
        fr_selected
        and pd.notna(ir_date)
        and ir_date.date() != dashboard_date
    ):

        increment_count(
            "PECVC_FINAL",
            "CE1",
            row.get("CE1_Result")
        )

# =====================================================
# COPY DATA
# =====================================================

sheet_mapping = {
    "Pre-Joining": ("PJ", "A:FC"),
    "PE & CVC": ("PE & CVC", "A:EI"),
    "Fresher": ("Fresher", "A:CE")
}

with pd.ExcelWriter(
    target_file,
    engine="openpyxl",
    mode="w"
) as writer:

    for source_sheet, (target_sheet, cols) in sheet_mapping.items():

        df = pd.read_excel(
            source_file,
            sheet_name=source_sheet,
            usecols=cols,
            engine="openpyxl",
            keep_default_na=False
        )

        df = df.fillna("")

        df.to_excel(
            writer,
            sheet_name=target_sheet,
            index=False
        )

check_display_names = {
    "CE1": "CE1",
    "CE2": "E2",
    "CE3": "E3",
    "CE4": "E4",
    "CE5": "E5",
    "CE6": "E6"
}
print("✅ Data Copied")
print("\n" + "=" * 60)
print("PRODUCTIVITY DASHBOARD")
print("=" * 60)

print("\nCID FINAL")
for k, v in dashboard["CID_FINAL"].items():
    print(f"{check_display_names.get(k, k):<12} : {v}")

print("\nEMP FINAL")
for k, v in dashboard["EMP_FINAL"].items():
    print(f"{check_display_names.get(k, k):<12} : {v}")

print("\nCID SUPPLEMENTARY")
for k, v in dashboard["CID_SUPPLEMENTARY"].items():
    print(f"{check_display_names.get(k, k):<12} : {v}")

print("\nEMP SUPPLEMENTARY")
for k, v in dashboard["EMP_SUPPLEMENTARY"].items():
    print(f"{check_display_names.get(k, k):<12} : {v}")

print("\nEMP SUPPLEMENTARY -2")
for k, v in dashboard["EMP_SUPPLEMENTARY2"].items():
    print(f"{check_display_names.get(k, k):<12} : {v}")

print("\nPE & CVC IR")
for k, v in dashboard["PECVC_IR"].items():
    print(f"{check_display_names.get(k, k):<12} : {v}")

print("\nPE & CVC FINAL")
for k, v in dashboard["PECVC_FINAL"].items():
    print(f"{check_display_names.get(k, k):<12} : {v}")

print("\nFRESHER CID")
for k, v in dashboard["FRESHER_CID"].items():
    print(f"{check_display_names.get(k, k):<12} : {v}")

print("\nFRESHER EMP")
for k, v in dashboard["FRESHER_EMP"].items():
    print(f"{check_display_names.get(k, k):<12} : {v}")

# =====================================================
# CONSOLIDATED SUMMARY
# =====================================================

all_checks = [
    "EDU",
    "CRC",
    "GDC",
    "REINT_GDC",
    "CE1",
    "CE2",
    "CE3",
    "CE4",
    "CE5",
    "CE6"
]

can_id_summary = {}
emp_id_summary = {}

for check in all_checks:

    can_id_summary[check] = (
        dashboard["CID_FINAL"].get(check, 0)
        + dashboard["CID_SUPPLEMENTARY"].get(check, 0)
        + dashboard["PECVC_IR"].get(check, 0)
        + dashboard["FRESHER_CID"].get(check, 0)
    )

    emp_id_summary[check] = (
        dashboard["EMP_FINAL"].get(check, 0)
        + dashboard["EMP_SUPPLEMENTARY"].get(check, 0)
        + dashboard["EMP_SUPPLEMENTARY2"].get(check, 0)
        + dashboard["PECVC_FINAL"].get(check, 0)
        + dashboard["FRESHER_EMP"].get(check, 0)
    )

print("\n" + "=" * 70)
print("CAN_ID REPORTS COUNT")
print("=" * 70)

print(
    f"EDU:{can_id_summary['EDU']} | "
    f"CRC:{can_id_summary['CRC']} | "
    f"GDC:{can_id_summary['GDC']} | "
    f"R_GDC:{can_id_summary['REINT_GDC']}"
)

print(
    f"CE1:{can_id_summary['CE1']} | "
    f"E2:{can_id_summary['CE2']} | "
    f"E3:{can_id_summary['CE3']} | "
    f"E4:{can_id_summary['CE4']} | "
    f"E5:{can_id_summary['CE5']} | "
    f"E6:{can_id_summary['CE6']}"
)

print("=" * 70)

print("\n" + "=" * 70)
print("EMP_ID REPORTS COUNT")
print("=" * 70)

print(
    f"EDU:{emp_id_summary['EDU']} | "
    f"CRC:{emp_id_summary['CRC']} | "
    f"GDC:{emp_id_summary['GDC']} | "
    f"R_GDC:{emp_id_summary['REINT_GDC']}"
)

print(
    f"CE1:{emp_id_summary['CE1']} | "
    f"E2:{emp_id_summary['CE2']} | "
    f"E3:{emp_id_summary['CE3']} | "
    f"E4:{emp_id_summary['CE4']} | "
    f"E5:{emp_id_summary['CE5']} | "
    f"E6:{emp_id_summary['CE6']}"
)

print("=" * 70)

print("\n" + "=" * 70)
print("ES / CE SUMMARY")
print("=" * 70)

print(f"\n✅ ES Count : {es_count}")

for day, count in sorted(
    es_daywise.items(),
    key=lambda x: datetime.strptime(
        x[0],
        "%d-%b-%Y"
    )
):
    print(f"   {day} : {count}")

print()

print(f"\n✅ CE Count : {ce_count}")

for day, count in sorted(
    ce_daywise.items(),
    key=lambda x: datetime.strptime(
        x[0],
        "%d-%b-%Y"
    )
):
    print(f"   {day} : {count}")

print("=" * 70)

# =====================================================
# OPEN WORKBOOK
# =====================================================

wb = load_workbook(target_file)

# =====================================================
# DELETE COLUMNS
# =====================================================

def delete_columns(ws, ranges):

    for rng in ranges:

        start_col, end_col = rng.split(":")

        start_idx = column_index_from_string(start_col)
        end_idx = column_index_from_string(end_col)

        ws.delete_cols(
            start_idx,
            end_idx - start_idx + 1
        )

# PJ
delete_columns(wb["PJ"], [
    "FA:FD",
    "EW:EY",
    "ES:EU",
    "EK:EQ",
    "EB:EI",
    "DS:DZ",
    "DJ:DQ",
    "DA:DH",
    "CR:CY",
    "CH:CP",
    "BZ:CF",
    "BJ:BX",
    "AP:BH",
    "AI:AI",
    "Z:Z",
    "N:V",
    "L:L",
    "I:J",
    "F:G",
    "C:C"
])

# PE & CVC
delete_columns(wb["PE & CVC"], [
    "EF:EI",
    "EB:ED",
    "DT:DZ",
    "DK:DR",
    "DB:DI",
    "CS:CZ",
    "CJ:CQ",
    "BZ:CH",
    "BR:BX",
    "BC:BP",
    "AJ:BA",
    "AC:AC",
    "Z:Z",
    "N:V",
    "L:L",
    "I:J",
    "F:G",
    "C:C"
])

# Fresher
delete_columns(wb["Fresher"], [
    "CC:CF",
    "BY:CA",
    "BQ:BW",
    "BJ:BO",
    "AP:BH",
    "AI:AI",
    "Z:Z",
    "N:V",
    "L:L",
    "I:J",
    "F:G",
    "C:C"
])

print("✅ Columns Deleted")

# =====================================================
# DATE FORMAT
# =====================================================

sheets = ["PJ", "PE & CVC", "Fresher"]

for sheet_name in sheets:

    ws = wb[sheet_name]

    for row in ws.iter_rows():

        for cell in row:

            if isinstance(
                cell.value,
                (datetime, date)
            ):

                cell.number_format = "dd-mmm-yy"

print("✅ Date Format Applied")

# =====================================================
# FORMATTING
# =====================================================

thin_border = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

for sheet_name in sheets:

    ws = wb[sheet_name]

    ws.freeze_panes = "A2"

    for cell in ws[1]:

        cell.font = Font(bold=True)

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    for row in ws.iter_rows(
        min_row=1,
        max_row=ws.max_row,
        min_col=1,
        max_col=ws.max_column
    ):

        for cell in row:

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )
            
            cell.border = thin_border

    for column in ws.columns:

        max_length = 0

        column_letter = get_column_letter(
            column[0].column
        )

        for cell in column:

            try:
                if cell.value is not None:

                    max_length = max(
                        max_length,
                        len(str(cell.value))
                    )

            except:
                pass

        ws.column_dimensions[
            column_letter
        ].width = min(
            max_length + 2,
            50
        )

print("✅ Formatting Applied")

# =====================================================
# SAVE
# =====================================================

wb.save(target_file)

print("\n✅ Case Validation Preparation Completed")