import os
import win32com.client

downloads = r"C:\Users\Balakumaran VR\Downloads"

rdt_files = [
    os.path.join(downloads, f)
    for f in os.listdir(downloads)
    if f.startswith("Report Dispatch Tracker")
    and f.endswith(".xlsx")
]

latest_file = max(
    rdt_files,
    key=os.path.getmtime
)

outlook = win32com.client.Dispatch(
    "Outlook.Application"
)

namespace = outlook.GetNamespace("MAPI")

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

latest_mail = None

for item in sent_folder.Items:

    try:

        if (
            "RE: REPORT DISPATCH TRACKER- CONSOLIDATED"
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
        "Required sent mail not found"
    )

else:

    mail = latest_mail.ReplyAll()

    mail.Display()

    body = mail.HTMLBody

    mail.HTMLBody = (
        '<span style="font-family:Calibri;font-size:11pt;">'
        'Hi Team,<br><br>'
        'Please find the attached Report dispatch tracker.'
        '</span><br><br>'
        + body
    )

    mail.Attachments.Add(latest_file)
      
    print(
        "Draft Created Successfully"
    )

    choice = input(
        "\nSend Mail? (Y/N): "
    ).strip().upper()

    if choice == "Y":

        mail.Send()

        print(
            "Mail Sent Successfully"
        )

    else:

        print(
            "Mail Cancelled"
        )