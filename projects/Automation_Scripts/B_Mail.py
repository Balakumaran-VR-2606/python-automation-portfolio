import os
import glob
import win32com.client

downloads = r"C:\Users\Balakumaran VR\Downloads"


def latest_master_tracker():
    candidates = []
    for pattern in ("Master Tracker*.xlsx", "Infy_Conventional-Master Tracker.xlsx"):
        candidates.extend(glob.glob(os.path.join(downloads, pattern)))
    candidates = [path for path in candidates if os.path.isfile(path)]
    if not candidates:
        raise FileNotFoundError(
            "No Master Tracker workbook found in Downloads. Expected a dated "
            "'Master Tracker*.xlsx' or 'Infy_Conventional-Master Tracker.xlsx'."
        )
    return max(candidates, key=os.path.getmtime)

invoice_folder = (
    r"C:\Users\Balakumaran VR\OneDrive - digiverifier solutions pvt ltd"
    r"\Shortcuts\Client Account Managers - Documents\General"
)

# Master Tracker
master_tracker = latest_master_tracker()

# Invoice Sheet
invoice_sheet = max(
    glob.glob(
        os.path.join(
            invoice_folder,
            "Infosys Conventional-AM-Invoice Sheet*"
        )
    ),
    key=os.path.getmtime
)

outlook = win32com.client.Dispatch(
    "Outlook.Application"
)

mail = outlook.CreateItem(0)

mail.To = "bhargavi.babu@digiverifier.com"

mail.Subject = "MT & DIT"

mail.HTMLBody = """
<div style="font-family:Calibri;font-size:11pt;">
Hi Bhargavi,
<br><br>
FYR.
</div>
"""

attached_items = []
skipped_items = []

# Master Tracker
try:

    mail.Attachments.Add(master_tracker)

    attached_items.append("MT")

except Exception as e:

    skipped_items.append("MT")

    print(f"⚠️ MT Not Attached : {e}")

# Invoice Sheet
try:

    mail.Attachments.Add(invoice_sheet)

    attached_items.append("DIT")

except Exception as e:

    skipped_items.append("DIT")

    print(f"⚠️ DIT Not Attached : {e}")

print("\n========================")

if attached_items:

    print(
        "✅ ATTACHED :",
        ", ".join(attached_items)
    )

if skipped_items:

    print(
        "⚠️ SKIPPED :",
        ", ".join(skipped_items)
    )

print("========================\n")

mail.Display()

choice = input(
    "\nSend Mail? (Y/N): "
).strip().upper()

if choice == "Y":

    mail.Send()

    print("✅ Mail Sent Successfully")

else:

    print("✅ Draft Opened - Not Sent")
