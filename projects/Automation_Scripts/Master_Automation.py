import subprocess
import sys
import winsound
import win32com.client

speaker = win32com.client.Dispatch("SAPI.SpVoice")

voices = speaker.GetVoices()

speaker.Voice = voices.Item(1)

print("\nDATE INPUT FORMAT: DD-MMM-YYYY (example: 06-Oct-2026)")

scripts = [
    "Live_DIT_Count.py",
    "Report_Count.py",
    "DIT_Open.py",
    "Overall_Report_QC.py",
    "Report+RDT+MT_Validation+MailAutomation.py",
    "1Drive_Report_Upload.py",
    "Datasheet.py",
    "RDT.py",
    "PS.py",
    "Cleanup.py",
    "DS+Live_DIT.py",
    "B_Mail.py",
    "TS.py"
]

skip_scripts = set()

print("\nSelect Automations To Skip\n")

print("1. Live DIT Count")
print("2. Report Count")
print("3. DIT Open")
print("4. Overall Report QC")
print("5. Validation & Mail Automation")
print("6. OneDrive Upload")
print("7. Datasheet")
print("8. RDT")
print("9. Progress Sheet")
print("10. Cleanup")
print("11. DS + Live DIT Validation")
print("12. B Mail")
print("13. Teams Automation")

skip_input = input(
    "\nEnter numbers to skip (comma separated) or press Enter: "
).strip()

skip_map = {
    "1": "Live_DIT_Count.py",
    "2": "Report_Count.py",
    "3": "DIT_Open.py",
    "4": "Overall_Report_QC.py",
    "5": "Report+RDT+MT_Validation+MailAutomation.py",
    "6": "1Drive_Report_Upload.py",
    "7": "Datasheet.py",
    "8": "RDT.py",
    "9": "PS.py",
    "10": "Cleanup.py",
    "11": "DS+Live_DIT.py",
    "12": "B_Mail.py",
    "13": "TS.py"
}

if skip_input:

    for item in skip_input.split(","):

        item = item.strip()

        if item in skip_map:

            skip_scripts.add(
                skip_map[item]
            )

failed_scripts = []

for script in scripts:

    if script in skip_scripts:

        print("\n" + "=" * 70)
        print(f"⏭️ SKIPPED : {script}")
        print("=" * 70)

        continue

    print("\n" + "=" * 70)
    print(f"🚀 STARTING : {script}")
    print("=" * 70)

    try:
        result = subprocess.run(
            [sys.executable, script]
        )

        if script == "Live_DIT_Count.py":
            print("✅ LIVE DIT COUNT COMPLETED")
            speaker.Speak(
                "Bala, Live D I T Count taken"
            )

        elif script == "Report_Count.py":
            print("✅ REPORT COUNT COMPLETED")
            speaker.Speak(
                "Bala, Report Count completed da"
            )

        elif script == "DIT_Open.py":
            print("✅ DIT INVOICE SHEETS OPENED")
            speaker.Speak(
                "Bala, Invoice sheets opened"
            )

        elif script == "Overall_Report_QC.py":
            print("✅ OVERALL REPORT QC COMPLETED")
            speaker.Speak(
                "Bala, Overall Report Q C Done"
            )

        elif script == "Report+RDT+MT_Validation+MailAutomation.py":
            print("✅ VALIDATION COMPLETED")
            speaker.Speak(
                "Bala, Validation is completed"
            )

        elif script == "1Drive_Report_Upload.py":
            print("✅ ONEDRIVE UPLOAD COMPLETED")
            speaker.Speak(
                "Bala, Reports are uploaded in onedrive"
            )

        elif script == "Datasheet.py":
            print("✅ DATASHEET UPDATED")
            speaker.Speak(
                "Bala, Datasheet done da"
            )
        
        elif script == "RDT.py":
            print("✅ RDT MAIL COMPLETED")
            speaker.Speak(
                "Bala, Report Dispatch Tracker mail sent da"
            )
        
        elif script == "PS.py":
            print("✅ PROGRESS SHEET MAIL COMPLETED")
            speaker.Speak(
                "Bala, Progress Sheet mail sent da"
            )

        elif script == "Cleanup.py":
            print("✅ CLEANUP COMPLETED")
            speaker.Speak(
                "Bala, Cleanup completed da"
            )

        elif script == "DS+Live_DIT.py":
            print("✅ DATASHEET vs INVOICE VALIDATION COMPLETED")
            speaker.Speak(
                "Bala, Datasheet and Invoice validation completed da"
            )
        
        elif script == "B_Mail.py":
            print("✅ B MAIL COMPLETED")
            speaker.Speak(
                "Bala, Bhargavi mail sent da"
            )

        elif script == "TS.py":
            print("✅ TEAMS AUTOMATION COMPLETED")
            speaker.Speak(
                "Bala, Teams automation Done da"
            )

        if result.returncode != 0:

            print(f"❌ FAILED : {script}")

            speaker.Speak(
                f"Bala, {script} failed. Can we Rerun?"
            )

            while result.returncode != 0:

                rerun = input(
                    f"\n🔄 {script} failed. Rerun now? (Y/N): "
                ).strip().upper()

                if rerun != "Y":

                    failed_scripts.append(script)

                    print(
                        f"\n⏭️ Skipped Rerun : {script}"
                    )

                    break

                print(f"\n🔄 RERUNNING : {script}")

                result = subprocess.run(
                    [sys.executable, script]
                )

                if result.returncode == 0:

                    print(
                        f"✅ RERUN SUCCESS : {script}"
                    )

    except Exception as e:
        print(f"❌ ERROR : {script}")
        print(e)
        failed_scripts.append(script)

print("\n" + "=" * 70)
print("🎯 INFOSYS QC AUTOMATION - SUMMARY")
print("=" * 70)

if failed_scripts:
    print("\n❌ FAILED SCRIPTS")

    for s in failed_scripts:
        print(f"   • {s}")

    winsound.Beep(400, 1000)

else:
    print("\n✅ ALL SCRIPTS EXECUTED SUCCESSFULLY")
    print("📊 LIVE DIT COUNT")
    print("📈 REPORT COUNT")
    print("📂 DIT INVOICE SHEETS")
    print("🔍 OVERALL REPORT QC")
    print("📧 VALIDATION & MAIL AUTOMATION")
    print("☁️ ONEDRIVE UPLOAD")
    print("📑 DATASHEET UPDATED") 
    print("📨 RDT MAIL")
    print("📋 PROGRESS SHEET MAIL")
    print("🧹 CLEANUP COMPLETED")
    print("✅ DATASHEET vs INVOICE VALIDATION") 
    print("📬 B MAIL")
    print("💬 TEAMS AUTOMATION")

    if not skip_scripts:

        speaker.Speak(
            "Bala. All automation successfully done da. Be Happy."
        )

    else:

        speaker.Speak(
            "Bala. Selected automation successfully done da."
        )

    if skip_scripts:

        print("\n⏭️ SKIPPED SCRIPTS")

        for s in skip_scripts:
            print(f"   • {s}")

print("\n🏁 PROCESS COMPLETED")
print("=" * 70)