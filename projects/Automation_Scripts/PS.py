import os
import tempfile
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# =====================================================
# GET PROGRESS SHEET MAIL FROM CLARIFICATIONS
# =====================================================

source_mail = None

for store in namespace.Folders:

    try:

        if "balakumaran.vr@digiverifier.com" in store.Name.lower():
            
            clarifications = store.Folders["Clarifications"]

            latest_time = None

            for item in clarifications.Items:

                try:
                    
                    if "PROGRESS SHEET" in item.Subject.upper():
                        
                        if (
                            latest_time is None
                            or item.ReceivedTime > latest_time
                        ):

                            latest_time = item.ReceivedTime
                            source_mail = item

                except Exception as e:
                    pass
   
            break

    except:
        pass

if source_mail is None:

    print("Progress Sheet mail not found")

    exit()

# =====================================================
# SAVE REQUIRED ATTACHMENTS
# =====================================================

temp_folder = tempfile.mkdtemp()

saved_files = []

for attachment in source_mail.Attachments:

    name = attachment.FileName

    if "Lateral Pre-Employment GDC" in name:
        continue

    allowed = [
        "Fresher",
        "CVC Lateral",
        "Lateral Post Employment",
        "Lateral Pre Joining"
    ]

    if any(x in name for x in allowed):

        path = os.path.join(
            temp_folder,
            name
        )

        attachment.SaveAsFile(path)

        saved_files.append(path)

print(f"Attachments Collected : {len(saved_files)}")

# =====================================================
# FIND LATEST SENT THREAD
# =====================================================

sent_folder = None

for store in namespace.Folders:

    try:

        if (
            store.Name.lower()
            ==
            "infosys.bgvdoc@digiverifier.com"
        ):

            sent_folder = store.Folders["Sent Items"]

            break

    except:
        pass

if sent_folder is None:

    print(
        "Infosys mailbox not found"
    )

    exit()

latest_mail = None

for item in sent_folder.Items:

    try:

        if (
            "RE: PROGRESS SHEET FORMAT"
            in item.Subject.upper()
        ):

            if (
                latest_mail is None
                or item.SentOn > latest_mail.SentOn
            ):

                latest_mail = item

    except:
        pass

if latest_mail is None:

    print(
        "Progress Sheet thread not found"
    )

    exit()

# =====================================================
# REPLY ALL
# =====================================================

mail = latest_mail.ReplyAll()

mail.Display()

body = mail.HTMLBody

mail.HTMLBody = (
    '<span style="font-family:Calibri;font-size:11pt;">'
    'Hi Team,<br><br>'
    'Attached here with the Progress sheets for all the packages.'
    '</span><br><br>'
    + body
)

for file_path in saved_files:

    mail.Attachments.Add(file_path)

print(
    "\nDraft Created Successfully"
)

choice = input(
    "\nSend Mail? (Y/N): "
).strip().upper()

if choice == "Y":

    mail.Send()

    print(
        "\nProgress Sheet Mail Sent Successfully"
    )

else:

    try:
        mail.Close(1)
    except:
        pass

    print(
        "\nMail Cancelled"
    )