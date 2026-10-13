import os
import win32com.client

master_tracker = r"C:\Users\Balakumaran VR\OneDrive - digiverifier solutions pvt ltd\Shortcuts\Infosys_Conventional - Documents\Infosys_Conventional\Infy_Conventional-Master Tracker.xlsx"

outlook = win32com.client.Dispatch(
    "Outlook.Application"
)

mail = outlook.CreateItem(0)

mail.To = "bhargavi.babu@digiverifier.com"

mail.Subject = "Master Tracker"

mail.HTMLBody = """
<div style="font-family:Calibri;font-size:11pt;">
Hi Bhargavi,
<br><br>
FYR.
</div>
"""

mail.Attachments.Add(master_tracker)

mail.Display()

choice = input(
    "\nSend Mail? (Y/N): "
).strip().upper()

if choice == "Y":

    mail.Send()

    print("✅ Mail Sent Successfully")

else:

    print("✅ Draft Opened - Not Sent")