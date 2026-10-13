import os

files_to_open = [
    r"C:\Users\Balakumaran VR\OneDrive - digiverifier solutions pvt ltd\Shortcuts\Client Account Managers - Documents\General\SLN(Conventional)-AM-Invoice Sheet.xlsx",
    r"C:\Users\Balakumaran VR\OneDrive - digiverifier solutions pvt ltd\Shortcuts\Client Account Managers - Documents\General\Spruce(Conventional)-AM-Invoice Sheet.xlsx",
    r"C:\Users\Balakumaran VR\OneDrive - digiverifier solutions pvt ltd\Shortcuts\Client Account Managers - Documents\General\SRS Infoway(Conventional)-AM-Invoice Sheet.xlsx",
    r"C:\Users\Balakumaran VR\OneDrive - digiverifier solutions pvt ltd\Shortcuts\Client Account Managers - Documents\General\TekisHub(Conventional)-AM-Invoice Sheet.xlsx",
    r"C:\Users\Balakumaran VR\OneDrive - digiverifier solutions pvt ltd\Shortcuts\Client Account Managers - Documents\General\First Meridian(Conventional)-AM-Invoice Sheet.xlsx"
]

for file in files_to_open:

    if os.path.exists(file):
        os.startfile(file)
        print(f"✅ Opened : {os.path.basename(file)}")

    else:
        print(f"❌ File Not Found : {file}")
