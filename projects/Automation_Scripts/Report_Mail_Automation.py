"""Create Outlook report emails from PDFs in Downloads\\Reports."""

from __future__ import annotations

import html
import sys
import uuid
from datetime import datetime
from pathlib import Path

import win32com.client


DOWNLOADS = Path(r"C:\Users\Balakumaran VR\Downloads")
REPORTS_FOLDER = DOWNLOADS / "Reports"
SENDER_EMAIL = "infosys.bgvdoc@digiverifier.com"
LATERAL_RECIPIENT = "lateralbgc@infosys.com"
FRESHER_RECIPIENT = "freshersbgc@infosys.com"
DATE_FORMAT = "%d-%b-%Y"


def prompt_report_date() -> str:
    while True:
        date_text = input("Subject date (DD-MMM-YYYY, e.g. 07-Oct-2026): ").strip()
        try:
            return datetime.strptime(date_text, DATE_FORMAT).strftime(DATE_FORMAT)
        except ValueError:
            print("Invalid date. Enter it as DD-MMM-YYYY, for example 07-Oct-2026.")


def get_recipient_and_subject(pdf_path: Path, report_date: str) -> tuple[str, str]:
    report_name = pdf_path.stem
    folder_path = str(pdf_path.parent).casefold()

    if "fresher" in folder_path:
        return FRESHER_RECIPIENT, f"{report_name} - {report_date}"
    if "cvc" in folder_path:
        return LATERAL_RECIPIENT, f"{report_name}_CVC - {report_date}"
    return LATERAL_RECIPIENT, f"{report_name} - {report_date}"


def build_html_body(note: str) -> str:
    note_text = (note or "").replace("\xa0", " ").replace("\u200b", "").replace("\ufeff", "").strip()
    body = (
        '<p style="margin:0 0 12pt 0">Dear Team,</p>'
        '<p style="margin:0 0 12pt 0">Please find the enclosed verified report '
        'with reference to the subject above.</p>'
    )
    if note_text:
        note_html = html.escape(note_text).replace("\r\n", "\n").replace("\n", "<br>")
        body += (
            '<p style="margin:0 0 12pt 0"><strong>Note : '
            f"{note_html}</strong>"
            "</p>"
        )
    body += '<p style="margin:0">Regards,<br>BGV Team - DigiVerifier</p>'
    return (
        '<html><body><div style="font-family:Calibri, sans-serif; '
        'font-size:11pt">'
        f"{body}"
        "</div></body></html>"
    )


def main() -> int:
    if not REPORTS_FOLDER.is_dir():
        print(f"Reports folder not found: {REPORTS_FOLDER}")
        return 1

    pdf_files = sorted(
        (path for path in REPORTS_FOLDER.rglob("*") if path.is_file() and path.suffix.casefold() == ".pdf"),
        key=lambda path: (str(path.parent).casefold(), path.name.casefold()),
    )
    if not pdf_files:
        print(f"No PDF reports found under: {REPORTS_FOLDER}")
        return 1

    report_date = prompt_report_date()
    note = input("Note to include in every email (leave blank for no Note): ").strip()

    print(f"\nFound {len(pdf_files)} PDF report(s). Subject date: {report_date}")
    if input("Create Outlook drafts for these reports? (Y/N): ").strip().upper() != "Y":
        print("Cancelled. No drafts were created.")
        return 0

    outlook = win32com.client.Dispatch("Outlook.Application")
    sender_account = None
    for account in outlook.Session.Accounts:
        try:
            if str(account.SmtpAddress).casefold() == SENDER_EMAIL.casefold():
                sender_account = account
                break
        except Exception:
            continue
    if sender_account is None:
        print(f"Outlook sender account not found: {SENDER_EMAIL}")
        return 1

    try:
        # olFolderDrafts = 16. Save each draft in the selected account's own store.
        sender_drafts = sender_account.DeliveryStore.GetDefaultFolder(16)
    except Exception as exc:
        print(f"Could not open the {SENDER_EMAIL} Drafts folder: {exc}")
        return 1
    print(f"Drafts will be saved to: {sender_drafts.FolderPath}")

    body = build_html_body(note)
    cleanup_tag = "ReportMailCleanup_" + uuid.uuid4().hex
    created_mails = []
    for pdf_path in pdf_files:
        recipient, subject = get_recipient_and_subject(pdf_path, report_date)
        try:
            mail = sender_drafts.Items.Add("IPM.Note")
            mail.SendUsingAccount = sender_account
            mail.To = recipient
            mail.Subject = subject
            mail.HTMLBody = body
            mail.Attachments.Add(str(pdf_path))
            mail.BillingInformation = cleanup_tag
            mail.Save()
            created_mails.append(mail)
            print(f"Draft Created : {len(created_mails)} | {subject} | {pdf_path.name}")
        except Exception as exc:
            print(f"Draft Failed : {pdf_path.name} | {exc}")

    print(f"\nTotal drafts created: {len(created_mails)} / {len(pdf_files)}")
    if not created_mails:
        return 1

    if input("Send all created drafts now? (Y/N): ").strip().upper() == "Y":
        sent_count = 0
        for mail in created_mails:
            try:
                mail.BillingInformation = ""
                mail.Save()
                mail.Send()
                sent_count += 1
            except Exception as exc:
                print(f"Send Failed : {mail.Subject} | {exc}")
        print(f"Successfully sent: {sent_count} / {len(created_mails)}")
    else:
        print("Send declined. Permanently deleting the drafts created in this run...")
        deleted_count = 0
        try:
            deleted_folder = sender_account.DeliveryStore.GetDefaultFolder(3)  # olFolderDeletedItems
            for mail in created_mails:
                try:
                    mail.Delete()  # Move from Drafts to Deleted Items.
                except Exception as exc:
                    print(f"Draft removal failed : {mail.Subject} | {exc}")

            tagged_items = deleted_folder.Items.Restrict(
                f"[BillingInformation] = '{cleanup_tag}'"
            )
            tagged_total = tagged_items.Count
            for item_index in range(tagged_total, 0, -1):
                try:
                    tagged_items.Item(item_index).Delete()  # Remove from Deleted Items too.
                    deleted_count += 1
                except Exception as exc:
                    print(f"Permanent deletion failed for a tagged draft: {exc}")
                    break
            print(f"Permanently deleted from Outlook folders: {deleted_count} / {len(created_mails)}")
        except Exception as exc:
            print(f"Could not complete permanent draft cleanup: {exc}")
            print("Check the Infosys BGVDoc Deleted Items folder for this run's drafts.")

    return 0


if __name__ == "__main__":
    sys.exit(main())
