import os
import shutil
import win32com.client

# =====================================
# DATE
# =====================================

report_date = input(
    "Enter Date (DD-MMM-YYYY): "
).strip()

# =====================================
# CLOSE EXCEL FILES
# =====================================

try:

    excel = win32com.client.Dispatch(
        "Excel.Application"
    )

    target_files = [
        "Daily Count DIT.xlsx",
        "Infosys Conventional-AM-Invoice Sheet.xlsx",
        "Infy_Conventional-Master Tracker.xlsx"
    ]

    for wb in list(excel.Workbooks):

        try:

            file_name = os.path.basename(
                wb.FullName
            )

            print(f"OPEN : {file_name}")

            if file_name in target_files:

                wb.Close(
                    SaveChanges=True
                )

                print(
                    f"✅ Closed : {file_name}"
                )

        except Exception as e:

            print(
                f"⚠️ Error Closing Workbook : {e}"
            )

except Exception as e:

    print(
        f"⚠️ Excel Not Detected : {e}"
    )

# =====================================
# DELETE REPORTS FOLDER
# =====================================

reports_folder = os.path.join(
    os.path.expanduser("~"),
    "Downloads",
    "Reports"
)

if os.path.exists(reports_folder):

    shutil.rmtree(reports_folder)

    print(
        "✅ Reports Folder Deleted"
    )

else:

    print(
        "⚠️ Reports Folder Not Found"
    )

# =====================================
# RENAME LOCAL MASTER TRACKER
# =====================================

old_tracker = (
    r"C:\Users\Balakumaran VR\Downloads"
    r"\Infy_Conventional-Master Tracker.xlsx"
)

new_tracker = (
    rf"C:\Users\Balakumaran VR\Downloads"
    rf"\Master Tracker {report_date}.xlsx"
)

if os.path.exists(old_tracker):

    if os.path.exists(new_tracker):

        os.remove(new_tracker)

    os.rename(
        old_tracker,
        new_tracker
    )

    print(
        f"✅ Renamed : Master Tracker {report_date}.xlsx"
    )

else:

    print(
        "❌ Infy_Conventional-Master Tracker.xlsx not found"
    )

# =====================================
# RENAME REPORTS ZIP
# =====================================

downloads = os.path.join(
    os.path.expanduser("~"),
    "Downloads"
)

for file in os.listdir(downloads):

    if (
        file.upper().startswith("REPORTS")
        and
        file.lower().endswith(".zip")
    ):

        old_zip = os.path.join(
            downloads,
            file
        )

        new_zip = os.path.join(
            downloads,
            f"Reports {report_date}.zip"
        )

        if os.path.exists(new_zip):
            os.remove(new_zip)

        os.rename(
            old_zip,
            new_zip
        )

        print(
            f"✅ Renamed : Reports {report_date}.zip"
        )

        break

print("\n🚀 CLEANUP COMPLETED")
