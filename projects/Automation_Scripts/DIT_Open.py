import os
import glob

folder = r"C:\Users\Balakumaran VR\OneDrive - digiverifier solutions pvt ltd\Shortcuts\Client Account Managers - Documents\General"

files = glob.glob(
    os.path.join(
        folder,
        "Infosys Conventional-AM-Invoice Sheet*"
    )
)

if files:
    os.startfile(files[0])
    print("✅ Invoice Sheet Opened")
else:
    print("❌ Invoice Sheet not found")

# Daily Count DIT

dit_file = r"C:\Users\Balakumaran VR\Downloads\Daily Count DIT.xlsx"

if os.path.exists(dit_file):
    os.startfile(dit_file)
    print("✅ Daily Count DIT Opened")
else:
    print("❌ Daily Count DIT not found")