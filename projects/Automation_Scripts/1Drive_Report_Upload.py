import os
import shutil
from datetime import datetime
from colorama import init, Fore, Style

init(autoreset=True)

# =====================================================
# FIND REPORTS FOLDER IN DOWNLOADS
# =====================================================

downloads = os.path.join(
    os.path.expanduser("~"),
    "Downloads"
)

REPORTS_ROOT = None

for item in os.listdir(downloads):

    full_path = os.path.join(downloads, item)

    if (
        item.lower() == "reports"
        and os.path.isdir(full_path)
    ):
        REPORTS_ROOT = full_path
        break

if not REPORTS_ROOT:
    print("\nExtracted Reports folder not found.")
    exit()

# =====================================================
# ONEDRIVE ROOT
# =====================================================

ONEDRIVE_ROOT = r"C:\Users\Balakumaran VR\OneDrive - digiverifier solutions pvt ltd\CROSSBOW Global Marketplace Solutions - Infosys_Conventional\Conv-Reports\Infy-Conv- QC Reports"

print("\n")
print(Fore.CYAN + "=" * 60)
print(Fore.YELLOW + "      DIGIVERIFIER | INFOSYS QC REPORT UPLOADER")
print(Fore.CYAN + "=" * 60)

print(
    Fore.GREEN +
    f"\nReports Folder Found : {REPORTS_ROOT}"
)

# =====================================================
# REPORT TYPE MAPPING
# =====================================================

REPORT_MAPPING = {
    "INTERIM": "INTERIM",
    "CID FINAL": "PRE FINAL",
    "CID SUPPLEMENTARY": "PRE FINAL",
    "CID SUPPLEMENTARY-2": "PRE FINAL",
    "CID SUPPLEMENTARY-3": "PRE FINAL",
    "EMP FINAL": "FINAL",
    "EMP SUPPLEMENTARY": "Supplementary Reports",
    "EMP SUPPLEMENTARY-2": "Supplementary 2 Reports",
    "EMP SUPPLEMENTARY-3": "Supplementary 3 Reports",
    "STOP CHECK": "STOP CHECK"
}

FOLDER_OPTIONS = {
    "1": "INTERIM",
    "2": "FINAL",
    "3": "PRE FINAL",
    "4": "STOP CHECK",
    "5": "Supplementary Reports",
    "6": "Supplementary 2 Reports",
    "7": "Supplementary 3 Reports"
}

# =====================================================
# USER INPUT
# =====================================================

print(Fore.CYAN + "\nSelect folders to create/upload\n")

for k, v in FOLDER_OPTIONS.items():
    print(Fore.YELLOW + f"{k}. {v}")

selection = input(
    "\nEnter Numbers (comma separated): "
).strip()

selected_folders = set()

for item in selection.split(","):
    item = item.strip()

    if item in FOLDER_OPTIONS:
        selected_folders.add(
            FOLDER_OPTIONS[item]
        )

report_date = input(
    "\nEnter Report Date (DD-MMM-YYYY): "
).strip()

start_time = datetime.now()

dt = datetime.strptime(
    report_date,
    "%d-%b-%Y"
)

year_folder = dt.strftime("%Y")
month_folder = dt.strftime("%b-%Y")

# =====================================================
# CREATE FOLDER STRUCTURE
# =====================================================

print("\n")
print(Fore.CYAN + "=" * 60)
print(Fore.YELLOW + "CREATING ONEDRIVE STRUCTURE")
print(Fore.CYAN + "=" * 60)

for target_folder in selected_folders:

    final_path = os.path.join(
        ONEDRIVE_ROOT,
        target_folder,
        year_folder,
        month_folder,
        report_date
    )

    os.makedirs(final_path, exist_ok=True)

    print(
        Fore.GREEN +
        f"[OK] Created : {final_path}"
    )

# =====================================================
# COPY REPORTS
# =====================================================

total_uploaded = 0

print("\n")
print(Fore.CYAN + "=" * 60)
print(Fore.YELLOW + "UPLOADING REPORTS")
print(Fore.CYAN + "=" * 60)

for root, dirs, files in os.walk(REPORTS_ROOT):

    folder_name = os.path.basename(root).upper()

    if folder_name not in REPORT_MAPPING:
        continue

    target_folder = REPORT_MAPPING[
        folder_name
    ]

    print(
        Fore.MAGENTA +
        f"{folder_name} ---> {target_folder}"
    )

    if target_folder not in selected_folders:
        continue

    destination_folder = os.path.join(
        ONEDRIVE_ROOT,
        target_folder,
        year_folder,
        month_folder,
        report_date
    )

    for file in files:

        if not file.lower().endswith(".pdf"):
            continue

        source_file = os.path.join(
            root,
            file
        )

        destination_file = os.path.join(
            destination_folder,
            file
        )

        try:

            source_long = (
                "\\\\?\\" +
                os.path.abspath(source_file)
            )

            destination_long = (
                "\\\\?\\" +
                os.path.abspath(destination_file)
            )

            shutil.copy2(
                source_long,
                destination_long
            )

            total_uploaded += 1

            print(
                Fore.YELLOW +
                f"Uploaded : {file}"
            )

        except Exception as e:

            print(
                Fore.RED +
                f"FAILED : {file}"
            )

            print(
                Fore.RED +
                f"PATH LENGTH : {len(destination_file)}"
            )

            print(
                Fore.RED +
                str(e)
            )  

# =====================================================
# SUMMARY
# =====================================================

end_time = datetime.now()

duration = end_time - start_time

print("\n")
print(Fore.CYAN + "=" * 60)
print(Fore.YELLOW + "UPLOAD SUMMARY")
print(Fore.CYAN + "=" * 60)

print(
    Fore.GREEN +
    f"TOTAL FILES UPLOADED : {total_uploaded}"
)

print(
    Fore.WHITE +
    f"START TIME : {start_time.strftime('%I:%M:%S %p')}"
)

print(
    Fore.WHITE +
    f"END TIME   : {end_time.strftime('%I:%M:%S %p')}"
)

print(
    Fore.YELLOW +
    f"DURATION   : {duration}"
)

print(
    Fore.GREEN +
    "STATUS : SUCCESS ✅"
)

print(Fore.CYAN + "=" * 60)

input("\nPress Enter To Exit...")