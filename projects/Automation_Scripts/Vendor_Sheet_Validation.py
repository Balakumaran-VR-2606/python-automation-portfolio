import os
import json
import pandas as pd
import re
from datetime import datetime
from openpyxl.styles import (
    Font,
    Alignment,
    Border,
    Side,
    PatternFill
)

from openpyxl.utils import get_column_letter
import win32com.client

def get_outlook_folders():
    """
    Returns a list of folders to search. Outlook's "Focused Inbox"
    feature splits mail between the "Inbox" folder and an "Other"/
    "Others" folder based on sender - so a mail can land outside
    plain Inbox. We search both.
    """

    outlook = win32com.client.Dispatch(
        "Outlook.Application"
    )

    namespace = outlook.GetNamespace(
        "MAPI"
    )

    root = namespace.Folders[
        "Infosys.bgvdoc@digiverifier.com"
    ]

    folders = []

    for name in ["Inbox", "Others", "Other", "Sent Items", "Sent"]:
        try:
            folders.append(root.Folders[name])
        except:
            pass

    return folders


def get_sender_smtp(mail):
    """
    mail.SenderEmailAddress returns an Exchange Legacy DN
    (e.g. "/O=EXCHANGELABS/OU=..." ) instead of a plain SMTP
    address for internal senders. This resolves the real SMTP
    address in all cases, both internal (EX) and external (SMTP).
    """

    try:
        if mail.SenderEmailType == "EX":
            return str(
                mail.Sender.GetExchangeUser().PrimarySmtpAddress
            ).lower()
        else:
            return str(mail.SenderEmailAddress).lower()
    except:
        pass

    try:
        # Fallback via MAPI property (works even when GetExchangeUser fails,
        # e.g. for a mail-enabled user not resolvable in the address book).
        PR_SMTP_ADDRESS = (
            "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
        )
        return str(
            mail.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
        ).lower()
    except:
        pass

    try:
        return str(mail.SenderEmailAddress).lower()
    except:
        return ""

VENDOR_REPORT_SENDERS = {
    "DHI": {"verifier4@dhiverification.com", "sachinb@dhiverification.com"},
    "IPANGS": {"verify3@ipangs.com", "verify5@ipangs.com"},
    "KRYA": {
        "jennifer.baktha@kryasolutions.com",
        "cam1@kryacheck.com",
        "naveenkumar.bapurao@kryacheck.com",
        "anupriya.srinivas@kryasolutions.com",
    },
}

KRYA_CONTACT_EMAILS = VENDOR_REPORT_SENDERS["KRYA"]
MAIL_INDEX_CACHE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Vendor_Mail_Index_Cache.json")


def get_recipient_emails(mail):
    """Extract SMTP addresses from Outlook To/Cc fields and Recipients."""
    addresses = set(re.findall(
        r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}",
        f"{getattr(mail, 'To', '')};{getattr(mail, 'CC', '')}"
    ))
    try:
        for recipient in mail.Recipients:
            try:
                address_entry = recipient.AddressEntry
                if address_entry.Type == "EX":
                    exchange_user = address_entry.GetExchangeUser()
                    smtp = exchange_user.PrimarySmtpAddress if exchange_user else ""
                else:
                    smtp = recipient.Address
                if smtp and "@" in str(smtp):
                    addresses.add(str(smtp).lower())
            except Exception:
                pass
    except Exception:
        pass
    return {address.lower() for address in addresses}


def get_email_datetime(value):
    return pd.to_datetime(value, errors="coerce").to_pydatetime()


def get_new_content(body):
    """
    Outlook reply/forward bodies include the ENTIRE quoted thread below
    the new content. If a reply mentions "stopcheck" for a DIFFERENT
    case, but the quoted thread underneath still contains an OLDER
    case's ID, a naive "does the body contain both the ID and
    'stopcheck'" check wrongly links them together. This returns only
    the text ABOVE the first quote/reply separator, so matching only
    looks at what THIS mail's sender actually wrote.
    """

    earliest = len(body)

    quote_patterns = [
        re.compile(r"_{10,}"),  # Outlook's horizontal line before a quote
        re.compile(r"-{3,}\s*original message\s*-{3,}", re.IGNORECASE),
        re.compile(
            r"^from:\s*.+\r?\n\s*sent:\s*.+",
            re.IGNORECASE | re.MULTILINE
        ),
    ]

    for pattern in quote_patterns:
        match = pattern.search(body)
        if match and match.start() < earliest:
            earliest = match.start()

    return body[:earliest]


def get_attachment_ids(mail):
    """
    Some vendors (e.g. GDC reports) put the check ID in the ATTACHMENT
    FILE NAME rather than in the mail body. This scans attachment names
    for ID-shaped numbers so those still get picked up.
    """

    ids = set()

    try:
        for attachment in mail.Attachments:
            try:
                ids.update(
                    re.findall(r"\d{6,12}", str(attachment.FileName))
                )
            except:
                continue
    except:
        pass

    return ids


def extract_stopcheck_ids(content):
    """
    Returns the set of IDs in `content` that are genuinely associated
    with a "stop check" instruction - NOT just every ID that happens to
    appear anywhere in the same mail.

    - If the mail mentions only ONE id total, any "stop check" mention
      anywhere in the mail applies to it (covers the common format
      where the instruction and the ID are a couple of lines apart,
      e.g. "Kindly STOP Check Below Case" then the ID on its own line).
    - If the mail mentions MULTIPLE ids (e.g. a status table with one
      row per candidate), only the ids on a LINE that itself contains
      "stop check" are counted - otherwise one row's "Already Give
      STOP Check" remark would wrongly get linked to every other
      unrelated id listed in the same table/mail.
    """

    all_ids = set(re.findall(r"\d{6,12}", content))

    if len(all_ids) <= 1:
        content_clean = content.replace(" ", "").replace("-", "")
        if "stopcheck" in content_clean:
            return all_ids
        return set()

    matched = set()

    for line in content.splitlines():
        line_clean = line.replace(" ", "").replace("-", "")
        if "stopcheck" in line_clean:
            matched.update(re.findall(r"\d{6,12}", line))

    return matched


def get_filtered_folder_items(folder):
    """Ask Outlook for only this run's vendor and stop-check candidates.

    Restrict is applied by Outlook before Python enumerates messages, so a
    single-vendor run does not walk every item in the mailbox. KRYA mail can
    be either from one of the vendor contacts or addressed to them (To/Cc).
    """
    base_items = folder.Items
    predicates = []
    for address in sorted(VENDOR_REPORT_SENDERS[VENDOR]):
        predicates.append(
            f'"urn:schemas:httpmail:fromemail" = \'{address}\''
        )

    if VENDOR == "KRYA":
        for address in sorted(KRYA_CONTACT_EMAILS):
            escaped = address.replace("'", "''")
            predicates.extend([
                f'"urn:schemas:httpmail:displayto" LIKE \'%{escaped}%\'',
                f'"urn:schemas:httpmail:displaycc" LIKE \'%{escaped}%\'',
            ])

    # These are the shared stop-check mailboxes used by the existing matching
    # logic. Restrict them by sender too; do not enumerate unrelated mail.
    stopcheck_senders = {
        "saravanan.shanmugam@digiverifier.com",
        "daniel.durairaja@digiverifier.com",
        "jaiganesh.velayutham@digiverifier.com",
        "rakesh.kumaravel@digiverifier.com",
    }
    for address in sorted(stopcheck_senders):
        predicates.append(
            f'"urn:schemas:httpmail:fromemail" = \'{address}\''
        )

    dasl_filter = "@SQL=" + " OR ".join(f"({term})" for term in predicates)
    try:
        return base_items.Restrict(dasl_filter)
    except Exception as exc:
        raise RuntimeError(
            f"Outlook could not apply the selected-vendor mail filter in "
            f"'{folder.Name}'. No full-mailbox scan was started. "
            f"Filter error: {exc}"
        ) from exc


def build_mail_indexes():
    """Build or reuse mail indexes for the currently selected vendor only."""
    refresh = os.environ.get("VENDOR_VALIDATION_REFRESH_MAIL_INDEX", "0") == "1"
    vendor_cache = MAIL_INDEX_CACHE[:-5] + f"_{VENDOR}.json"
    cache_candidates = [vendor_cache, MAIL_INDEX_CACHE]
    if not refresh:
        for cache_path in cache_candidates:
            if not os.path.isfile(cache_path):
                continue
            try:
                with open(cache_path, "r", encoding="utf-8") as cache_file:
                    cached = json.load(cache_file)
                if cached.get("version") == 2 and cached.get("vendor") == VENDOR:
                    stopcheck_indexes = {
                        check: {key: get_email_datetime(value) for key, value in rows.items()}
                        for check, rows in cached["stopcheck_indexes"].items()
                    }
                    vendor_index = {
                        check: {key: get_email_datetime(value) for key, value in rows.items()}
                        for check, rows in cached["vendor_index"].items()
                    }
                    print(f"Reused {VENDOR} mail cache; other vendors were not searched.")
                    return stopcheck_indexes, {VENDOR: vendor_index}
                if cached.get("version") == 1 and VENDOR in cached.get("vendor_indexes", {}):
                    stopcheck_indexes = {
                        check: {key: get_email_datetime(value) for key, value in rows.items()}
                        for check, rows in cached["stopcheck_indexes"].items()
                    }
                    vendor_index = {
                        check: {key: get_email_datetime(value) for key, value in rows.items()}
                        for check, rows in cached["vendor_indexes"][VENDOR].items()
                    }
                    print(f"Reused {VENDOR} from the existing mail cache; other vendors were not searched.")
                    return stopcheck_indexes, {VENDOR: vendor_index}
            except Exception as exc:
                print(f"Mail-index cache could not be read ({exc}); rebuilding it.")

    stopcheck_indexes = {"EDU": {}, "CRC": {}, "GDC": {}}
    vendor_index = {"EDU": {}, "CRC": {}, "GDC": {}}
    vendor_indexes = {VENDOR: vendor_index}
    edu_stopcheck_senders = {
        "saravanan.shanmugam@digiverifier.com",
        "daniel.durairaja@digiverifier.com",
    }
    crc_stopcheck_senders = {
        "jaiganesh.velayutham@digiverifier.com",
        "rakesh.kumaravel@digiverifier.com",
    }
    gdc_stopcheck_senders = {"rakesh.kumaravel@digiverifier.com"}

    folders = get_outlook_folders()
    restricted_folders = []
    total = 0
    for folder in folders:
        items = get_filtered_folder_items(folder)
        restricted_folders.append((folder, items))
        total += items.Count
    print(
        f"Searching only {VENDOR} mail and matching stop-check mail: "
        f"{total} Outlook candidates after Outlook filtering..."
    )
    scanned = 0
    seen_items = set()

    for folder, filtered_items in restricted_folders:
        for mail in filtered_items:
            scanned += 1
            if scanned % 2000 == 0:
                print(f"  ...scanned {scanned}/{total}")
            try:
                entry_id = str(mail.EntryID)
                if entry_id in seen_items:
                    continue
                seen_items.add(entry_id)
                sender = get_sender_smtp(mail)
                subject = str(mail.Subject).lower()
                received = getattr(mail, "ReceivedTime", None) or getattr(mail, "SentOn", None)
                if received is None:
                    continue
            except Exception:
                continue

            is_edu_subject = (
                "case initiation || education verification" in subject
                or "education verification reports" in subject
                or "education" in subject
                or re.search(r"\bedu\b", subject) is not None
                or "verification report" in subject
            )
            is_crc_subject = "case initiation request || e court" in subject
            is_gdc_subject = "case initiation request || gdc" in subject
            stopcheck_ids = None
            if VENDOR == "KRYA":
                krya_addresses = {sender} | get_recipient_emails(mail)
                mail_vendor = "KRYA" if krya_addresses.intersection(KRYA_CONTACT_EMAILS) else None
            else:
                mail_vendor = VENDOR if sender in VENDOR_REPORT_SENDERS[VENDOR] else None
            stopcheck_candidate = sender in (
                edu_stopcheck_senders | crc_stopcheck_senders | gdc_stopcheck_senders
            )
            if not mail_vendor and not stopcheck_candidate:
                continue

            try:
                raw_body = str(mail.Body)
                body = raw_body.lower()
            except Exception:
                continue
            new_content = get_new_content(raw_body).lower()
            new_content_clean = new_content.replace(" ", "").replace("-", "")
            vendor_ids = None

            for check_type, stop_senders, subject_match in (
                ("EDU", edu_stopcheck_senders, is_edu_subject),
                ("CRC", crc_stopcheck_senders, is_crc_subject),
                ("GDC", gdc_stopcheck_senders, is_gdc_subject),
            ):
                if sender in stop_senders and subject_match and "stopcheck" in new_content_clean:
                    if stopcheck_ids is None:
                        stopcheck_ids = extract_stopcheck_ids(new_content)
                    for case_id in stopcheck_ids:
                        previous = stopcheck_indexes[check_type].get(case_id)
                        if previous is None or received > previous:
                            stopcheck_indexes[check_type][case_id] = received

            if mail_vendor and (is_edu_subject or mail_vendor == "KRYA"):
                vendor_ids = set(re.findall(r"\d{6,12}", body)) | get_attachment_ids(mail)
                if mail_vendor == "KRYA":
                    if is_gdc_subject or "gdc" in subject or "global database" in subject:
                        report_check = "GDC"
                    elif is_crc_subject or "e court" in subject or "criminal" in subject or re.search(r"\bcrc\b", subject):
                        report_check = "CRC"
                    else:
                        report_check = "EDU"
                else:
                    report_check = "EDU"
                for case_id in vendor_ids:
                    previous = vendor_indexes[mail_vendor][report_check].get(case_id)
                    if previous is None or received > previous:
                        vendor_indexes[mail_vendor][report_check][case_id] = received
            if mail_vendor in {"DHI", "IPANGS"} and is_crc_subject:
                vendor_ids = set(re.findall(r"\d{6,12}", body)) | get_attachment_ids(mail)
                for case_id in vendor_ids:
                    previous = vendor_indexes[mail_vendor]["CRC"].get(case_id)
                    if previous is None or received > previous:
                        vendor_indexes[mail_vendor]["CRC"][case_id] = received
            if mail_vendor in {"DHI", "IPANGS"} and is_gdc_subject:
                vendor_ids = set(re.findall(r"\d{6,12}", body)) | get_attachment_ids(mail)
                for case_id in vendor_ids:
                    previous = vendor_indexes[mail_vendor]["GDC"].get(case_id)
                    if previous is None or received > previous:
                        vendor_indexes[mail_vendor]["GDC"][case_id] = received

    def serialize_indexes(indexes):
        return {
            key: {
                subkey: {
                    item: pd.to_datetime(value).isoformat()
                    for item, value in rows.items()
                }
                for subkey, rows in nested.items()
            }
            for key, nested in indexes.items()
        }

    cache_payload = {
        "version": 2,
        "vendor": VENDOR,
        "stopcheck_indexes": serialize_indexes(stopcheck_indexes),
        "vendor_index": serialize_indexes({VENDOR: vendor_index})[VENDOR],
    }
    try:
        with open(vendor_cache, "w", encoding="utf-8") as cache_file:
            json.dump(cache_payload, cache_file, ensure_ascii=False)
        print(f"Saved {VENDOR} mail cache: {vendor_cache}")
    except Exception as exc:
        print(f"Could not save mail-index cache ({exc}); this run can continue.")

    print(
        f"{VENDOR} mail index built - "
        f"EDU {len(vendor_index['EDU'])}, CRC {len(vendor_index['CRC'])}, "
        f"GDC {len(vendor_index['GDC'])}"
    )
    return stopcheck_indexes, vendor_indexes


def find_stopcheck_mail(
    check_id,
    check_type
):
    idx = STOPCHECK_INDEX.get(check_type, {})
    return idx.get(
        str(check_id).strip().replace(" ", "").replace("-", "")
    )


def find_vendor_report_mail(
    check_id,
    check_type="EDU"
):
    idx = VENDOR_INDEX.get(check_type, {})
    return idx.get(
        str(check_id).strip().replace(" ", "").replace("-", "")
    )

def format_datetime(dt_value):

    if dt_value is None:
        return "NA"

    try:
        return pd.to_datetime(
            dt_value,
            utc=True
        ).strftime("%d-%b-%Y %H:%M")

    except:
        return "NA"


def lookup_mail_display(id_value, check_type):
    """
    Looks up stopcheck / vendor-report mail for this ID + check type and
    returns both as display-ready strings (format_datetime already
    applied, "NA" when nothing found). Used everywhere a remark row is
    generated, so a real timestamp shows up whenever the mail exists -
    instead of leaving these columns blank/"NA" without ever actually
    checking.
    """

    stopcheck_time = find_stopcheck_mail(id_value, check_type)
    vendor_report_time = find_vendor_report_mail(id_value, check_type)

    return (
        format_datetime(stopcheck_time),
        format_datetime(vendor_report_time)
    )

from datetime import timedelta

VALID_RESULTS = {
    "GREEN",
    "RED",
    "ORANGE",
    "MAROON",
    "STOP CHECK",
    "STOPCHECK"
}

MONTH = (
    os.environ.get("VENDOR_VALIDATION_MONTH", "").strip()
    or input("Enter Month (Ex: Aug-2026): ").strip()
)

VENDOR = (
    os.environ.get("VENDOR_VALIDATION_VENDOR", "").strip().upper()
    or input("Enter Vendor (DHI/KRYA/IPANGS): ").strip().upper()
)

try:
    selected_month = datetime.strptime(MONTH, "%b-%Y")
except ValueError as exc:
    raise ValueError(
        f"Invalid month '{MONTH}'. Enter month as MMM-YYYY, for example Aug-2026."
    ) from exc

if VENDOR not in {"DHI", "KRYA", "IPANGS"}:
    raise ValueError(
        f"Invalid vendor '{VENDOR}'. Choose DHI, KRYA, or IPANGS."
    )

# The Master Tracker may use vendor aliases in its Vendor columns. Treat
# those as matches too, so rows are not wrongly flagged as vendor mismatches.
if VENDOR == "IPANGS":
    VENDOR_MATCH_VALUES = {"IPANGS", "IPANG"}
elif VENDOR == "KRYA":
    VENDOR_MATCH_VALUES = {"KRYA", "KRYA CHECK"}
else:
    VENDOR_MATCH_VALUES = {VENDOR}


def extract_email(value):
    """
    The IPANGS file's SPOC column has "Display Name <email>" format
    (e.g. "Haripriya V <haripriya.v@digiverifier.com>"), not a plain
    email address. This pulls out just the email part for matching.
    """

    match = re.search(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", str(value))

    if match:
        return match.group(0).lower()

    return str(value).strip().lower()


DEFAULT_SPOC_EMAILS = {
    "EDU": {
        "daniel.durairaja@digiverifier.com",
        "saravanan.shanmugam@digiverifier.com"
    },
    "CRC": {
        "jaiganesh.velayutham@digiverifier.com",
        "rakesh.kumaravel@digiverifier.com"
    },
    "CRIMINAL": {
        "jaiganesh.velayutham@digiverifier.com",
        "rakesh.kumaravel@digiverifier.com"
    },
    "GDC": {
        "rakesh.kumaravel@digiverifier.com"
    }
}


def parse_email_list(raw_value):
    return {
        email.lower()
        for email in re.findall(
            r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}",
            str(raw_value)
        )
    }


def selected_spoc_emails(check_name):
    if VENDOR == "KRYA":
        return set()
    env_name = f"VENDOR_VALIDATION_SPOC_{check_name}"
    raw_value = os.environ.get(env_name)
    if raw_value is None:
        defaults = sorted(DEFAULT_SPOC_EMAILS[check_name])
        raw_value = input(
            f"SPOC email(s) for {check_name} "
            f"(comma-separated; Enter keeps default {', '.join(defaults)}): "
        ).strip()
        if not raw_value:
            return set(DEFAULT_SPOC_EMAILS[check_name])
    emails = parse_email_list(raw_value)
    if not emails:
        raise ValueError(
            f"No valid email addresses entered for {check_name}. "
            "Use one or more comma-separated email addresses."
        )
    return emails


SPOC_EMAILS = {
    check_name: selected_spoc_emails(check_name)
    for check_name in DEFAULT_SPOC_EMAILS
}


def normalize_spoc_name(value):
    return re.sub(r"[^a-z]", "", str(value).lower())


def selected_spoc_names(check_name):
    env_name = f"VENDOR_VALIDATION_KRYA_SPOC_{check_name}"
    if env_name in os.environ:
        raw_value = os.environ[env_name]
    else:
        base_env_name = f"VENDOR_VALIDATION_SPOC_{check_name}"
        if base_env_name in os.environ:
            raw_value = os.environ[base_env_name]
        else:
            raw_value = input(
                f"Krya SPOC name(s) for {check_name} (comma-separated): "
            )
    return {
        normalize_spoc_name(name)
        for name in re.split(r"[,;\n]+", raw_value)
        if normalize_spoc_name(name)
    }


SPOC_NAMES = {
    check_name: selected_spoc_names(check_name) if VENDOR == "KRYA" else set()
    for check_name in DEFAULT_SPOC_EMAILS
}

for check_name, emails in SPOC_EMAILS.items():
    if VENDOR == "KRYA":
        print(f"SPOC name filter - {check_name}: {', '.join(sorted(SPOC_NAMES[check_name])) or '(none)'}")
    else:
        print(f"SPOC filter - {check_name}: {', '.join(sorted(emails)) or '(none)'}")

# Build/reuse the shared index before opening vendor spreadsheets, so a
# missing source workbook does not discard a completed Outlook scan.
STOPCHECK_INDEX, ALL_VENDOR_INDEXES = build_mail_indexes()
VENDOR_INDEX = ALL_VENDOR_INDEXES.get(VENDOR, {"EDU": {}, "CRC": {}, "GDC": {}})


def find_column_like(df, name_prefix):
    """
    Finds a column whose name starts with name_prefix (case-insensitive,
    whitespace-tolerant). Needed because reading a sheet with repeated
    headers across side-by-side blocks (e.g. the IPANGS file's Education /
    GDC / CRC blocks all having a "URN / Ref No." column) makes pandas
    rename duplicates with ".1", ".2" suffixes - so an exact-name lookup
    would only ever find the first block's column.
    """

    prefix = name_prefix.strip().lower()

    for col in df.columns:
        if str(col).strip().lower().startswith(prefix):
            return col

    raise KeyError(
        f"No column starting with '{name_prefix}' found. "
        f"Available columns: {list(df.columns)}"
    )


def load_ipangs_source_data():
    """
    For VENDOR == 'IPANGS', the SPOC-initiated ID lists come from
    "Digiverifier Data Validation <Month> - <Year>.xlsx" (sheet "Sheet")
    instead of the Crossbow file. Three column blocks on that one sheet:
      A:R  - Education Check   (SPOC + "URN / Ref No.")
      Z:AK - Global Database Check / GDC (SPOC + "URN / Ref No.")
      AL:AW- E-Court Check / CRC (SPOC + "URN / Ref No.")

    Returns (edu_df, crc_df, criminal_df, global_df) with the SAME
    column names the rest of the script already expects from the
    Crossbow file ("Check id" / "Ref. No." / "Client Ref No" /
    "Check ID"), so nothing downstream needs to change.
    """

    ipangs_month_str = (
        selected_month.strftime("%B") + " - " + str(selected_month.year)
    )

    ipangs_file = f"Digiverifier Data Validation  {ipangs_month_str}.xlsx"

    print(f"IPANGS vendor selected - reading source file: {ipangs_file}")

    raw = pd.read_excel(ipangs_file, sheet_name="Sheet", header=2)

    # Column blocks by position (A=0). A:R = 18 cols, Z:AK = 12 cols,
    # AL:AW = 12 cols.
    edu_block = raw.iloc[:, 0:18].copy()
    gdc_block = raw.iloc[:, 25:37].copy()
    crc_block = raw.iloc[:, 37:49].copy()

    spoc_col_edu = find_column_like(
        edu_block, "Case Initiated by"
    )
    id_col_edu = find_column_like(
        edu_block, "URN / Ref No"
    )

    spoc_col_gdc = find_column_like(
        gdc_block, "Case Initiated by"
    )
    id_col_gdc = find_column_like(
        gdc_block, "URN / Ref No"
    )

    spoc_col_crc = find_column_like(
        crc_block, "Case Initiated by"
    )
    id_col_crc = find_column_like(
        crc_block, "URN / Ref No"
    )

    # ---- Education ----
    edu_block = edu_block[
        edu_block[spoc_col_edu].apply(extract_email).isin(SPOC_EMAILS["EDU"])
    ]

    edu_df = pd.DataFrame({
        "Check id": edu_block[id_col_edu],
        "Spoc Name": edu_block[spoc_col_edu]
    })

    # ---- GDC (Global) ----
    gdc_block = gdc_block[
        gdc_block[spoc_col_gdc].apply(extract_email).isin(SPOC_EMAILS["GDC"])
    ]

    global_df = pd.DataFrame({
        "Check ID": gdc_block[id_col_gdc],
        "Spoc Name": gdc_block[spoc_col_gdc]
    })

    # ---- CRC (E-Court) ----
    crc_block = crc_block[
        crc_block[spoc_col_crc].apply(extract_email).isin(SPOC_EMAILS["CRC"])
    ]

    crc_df = pd.DataFrame({
        "Ref. No.": crc_block[id_col_crc],
        "Spoc Name": crc_block[spoc_col_crc]
    })

    # No separate "Criminal" block in the IPANGS file - keep an empty
    # dataframe with the expected column so downstream code (which
    # concatenates crc_df + criminal_df) still works unchanged.
    criminal_df = pd.DataFrame({
        "Client Ref No": pd.Series(dtype=object),
        "Spoc Name": pd.Series(dtype=object)
    })

    return edu_df, crc_df, criminal_df, global_df


def load_krya_source_data():
    """Read Krya Applicant IDs for rows whose Components Detail includes EDU."""
    month_label = selected_month.strftime("%b %Y")
    expected_name = f"DIGIVERIFIER SOLUTIONS PRIVATE LIMITED {month_label}.xlsx"
    krya_file = expected_name
    if not os.path.isfile(krya_file):
        possible = [
            name for name in os.listdir(".")
            if name.lower().startswith("digiverifier solutions private limited")
            and name.lower().endswith(".xlsx")
            and selected_month.strftime("%b").lower() in name.lower()
            and str(selected_month.year) in name
        ]
        if len(possible) == 1:
            krya_file = possible[0]
        else:
            raise FileNotFoundError(
                f"Krya datasheet for {month_label} was not found. "
                f"Expected: {expected_name}"
            )

    print(f"KRYA source file: {krya_file}")
    raw = pd.read_excel(krya_file, header=None, dtype=object)
    required = {"applicantid", "spocname", "componentsdetail"}
    header_row = None
    for row_index in range(min(len(raw), 35)):
        normalized = {
            re.sub(r"[^a-z0-9]", "", str(value).lower())
            for value in raw.iloc[row_index].tolist()
        }
        if required.issubset(normalized):
            header_row = row_index
            break
    if header_row is None:
        raise ValueError(
            "Krya datasheet headers Applicant ID, SPOC Name, and Components Detail "
            "were not found together in the first 35 rows."
        )

    df = pd.read_excel(krya_file, header=header_row, dtype=object)
    col_map = {
        re.sub(r"[^a-z0-9]", "", str(column).lower()): column
        for column in df.columns
    }
    applicant_col = col_map["applicantid"]
    spoc_col = col_map["spocname"]
    components_col = col_map["componentsdetail"]

    spoc_names = df[spoc_col].fillna("").astype(str).apply(
        normalize_spoc_name
    )
    components = df[components_col].fillna("").astype(str)

    component_patterns = {
        "EDU": r"\b(?:EDU|EDUCATION)\s*:\s*\d+\b",
        "CRC": r"\b(?:CRC|E[\s-]*COURT)\s*:\s*\d+\b",
        "CRIMINAL": r"\b(?:CRIMINAL|CRIM)\s*:\s*\d+\b",
        "GDC": r"\b(?:GDC|GLOBAL\s+DATABASE)\s*:\s*\d+\b",
    }

    def selected_rows(check_name):
        names = SPOC_NAMES[check_name]
        if not names:
            return df.iloc[0:0].copy()
        name_mask = spoc_names.apply(
            lambda value: bool(value) and any(
                len(name) >= 3 and (name in value or value in name)
                for name in names
            )
        )
        component_mask = components.str.contains(
            component_patterns[check_name], case=False, regex=True
        )
        selected = df[name_mask & component_mask].copy()
        selected[applicant_col] = selected[applicant_col].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
        return selected[selected[applicant_col].str.match(r"^\d+$", na=False)].drop_duplicates(applicant_col)

    edu_selected = selected_rows("EDU")
    crc_selected = selected_rows("CRC")
    criminal_selected = selected_rows("CRIMINAL")
    gdc_selected = selected_rows("GDC")
    print(
        "KRYA source IDs selected by SPOC name + component: "
        f"EDU {len(edu_selected)}, CRC {len(crc_selected)}, "
        f"Criminal {len(criminal_selected)}, GDC {len(gdc_selected)}"
    )

    edu_df = pd.DataFrame({"Check id": edu_selected[applicant_col], "Spoc Name": edu_selected[spoc_col]})
    crc_df = pd.DataFrame({"Ref. No.": crc_selected[applicant_col], "Spoc Name": crc_selected[spoc_col]})
    criminal_df = pd.DataFrame({"Client Ref No": criminal_selected[applicant_col], "Spoc Name": criminal_selected[spoc_col]})
    global_df = pd.DataFrame({"Check ID": gdc_selected[applicant_col], "Spoc Name": gdc_selected[spoc_col]})
    return edu_df, crc_df, criminal_df, global_df


# ===========================
# LOAD FILES
# ===========================

mt_file = r"Infy_Conventional-Master Tracker.xlsx"

if VENDOR == "IPANGS":

    edu_df, crc_df, criminal_df, global_df = load_ipangs_source_data()

elif VENDOR == "KRYA":

    edu_df, crc_df, criminal_df, global_df = load_krya_source_data()

else:

    crossbow_file = f"Crossbow - {selected_month.strftime('%b %Y')}.xlsx"
    if not os.path.isfile(crossbow_file):
        raise FileNotFoundError(
            f"Crossbow workbook for {MONTH} was not found: {crossbow_file}"
        )
    print(f"Crossbow source file: {crossbow_file}")

    edu_df = pd.read_excel(crossbow_file, sheet_name="Edu")

    edu_df = edu_df[
        edu_df["Spoc Name"].apply(extract_email).isin(SPOC_EMAILS["EDU"])
    ]

    crc_df = pd.read_excel(crossbow_file, sheet_name="CRC")

    crc_df = crc_df[
        crc_df["Spoc Name"].apply(extract_email).isin(SPOC_EMAILS["CRC"])
    ]

    criminal_df = pd.read_excel(
        crossbow_file,
        sheet_name="Criminal"
    )

    criminal_df = criminal_df[
        criminal_df["Spoc Name"].apply(extract_email).isin(SPOC_EMAILS["CRIMINAL"])
    ]

    global_df = pd.read_excel(
        crossbow_file,
        sheet_name="Global"
    )

    global_df = global_df[
        global_df["Spoc Name"].apply(extract_email).isin(SPOC_EMAILS["GDC"])
    ]

crc_final_df = pd.concat(
    [crc_df, criminal_df],
    ignore_index=True
)

pre_df = pd.read_excel(
    mt_file,
    sheet_name="Pre-Joining",
    keep_default_na=False
)

pe_df = pd.read_excel(
    mt_file,
    sheet_name="PE & CVC",
    keep_default_na=False
)

fresh_df = pd.read_excel(
    mt_file,
    sheet_name="Fresher",
    keep_default_na=False
)

mt_df = pd.concat([pre_df, pe_df, fresh_df], ignore_index=True)

CHECKS_TO_INITIATE_COLUMN_EXISTS = "Checks to be Initiate" in mt_df.columns

if not CHECKS_TO_INITIATE_COLUMN_EXISTS:
    print(
        "WARNING: Column 'Checks to be Initiate' not found in the "
        "Master Tracker sheets - CRC 'wrongly initiated' checks will "
        "be SKIPPED entirely (not guessed). Check the exact column "
        "spelling if you expected this check to run."
    )

# =================================
# EDU - MT FILTER
# =================================

edu_mt_df = mt_df[
    mt_df["Ed_Result"]
    .astype(str)
    .str.upper()
    .isin(VALID_RESULTS)
].copy()

edu_mt_df["Ed_Edate"] = pd.to_datetime(
    edu_mt_df["Ed_Edate"],
    errors="coerce"
)

edu_mt_df = edu_mt_df[
    (edu_mt_df["Ed_Edate"].dt.month == selected_month.month)
    &
    (edu_mt_df["Ed_Edate"].dt.year == selected_month.year)
]

edu_mt_df = edu_mt_df[
    edu_mt_df["Vendor"]
    .astype(str)
    .str.strip()
    .str.upper()
    .isin(VENDOR_MATCH_VALUES)
]

# ===========================
# CRC - PRE JOINING
# ===========================

pre_crc = pre_df[
    pre_df["eP_CRC_Result"]
    .astype(str)
    .str.upper()
    .isin(VALID_RESULTS)
].copy()

pre_crc["eP_CRC_EDate"] = pd.to_datetime(
    pre_crc["eP_CRC_EDate"],
    errors="coerce"
)

pre_crc = pre_crc[
    (pre_crc["eP_CRC_EDate"].dt.month == selected_month.month)
    &
    (pre_crc["eP_CRC_EDate"].dt.year == selected_month.year)
]

pre_crc = pre_crc[
    pre_crc["Vendor.2"]
    .astype(str)
    .str.strip()
    .str.upper()
    .isin(VENDOR_MATCH_VALUES)
]

# ===========================
# CRC - PE & CVC
# ===========================

pe_crc = pe_df[
    pe_df["eP_CRC_Result"]
    .astype(str)
    .str.upper()
    .isin(VALID_RESULTS)
].copy()

pe_crc["eP_CRC_EDate"] = pd.to_datetime(
    pe_crc["eP_CRC_EDate"],
    errors="coerce"
)

pe_crc = pe_crc[
    (pe_crc["eP_CRC_EDate"].dt.month == selected_month.month)
    &
    (pe_crc["eP_CRC_EDate"].dt.year == selected_month.year)
]

pe_crc = pe_crc[
    pe_crc["Vendor.2"]
    .astype(str)
    .str.strip()
    .str.upper()
    .isin(VENDOR_MATCH_VALUES)
]

# ===========================
# CRC - FRESHER
# ===========================

fresh_crc = fresh_df[
    fresh_df["eP_CRC_Result"]
    .astype(str)
    .str.upper()
    .isin(VALID_RESULTS)
].copy()

fresh_crc["eP_CRC_EDate"] = pd.to_datetime(
    fresh_crc["eP_CRC_EDate"],
    errors="coerce"
)

fresh_crc = fresh_crc[
    (fresh_crc["eP_CRC_EDate"].dt.month == selected_month.month)
    &
    (fresh_crc["eP_CRC_EDate"].dt.year == selected_month.year)
]

fresh_crc = fresh_crc[
    fresh_crc["Vendor.1"]
    .astype(str)
    .str.strip()
    .str.upper()
    .isin(VENDOR_MATCH_VALUES)
]

# FINAL CRC MT

crc_mt_df = pd.concat(
    [pre_crc, pe_crc, fresh_crc],
    ignore_index=True
)

def gdc_month_mask(df, month, year):
    """
    A GDC row counts as belonging to the selected month if EITHER
    GDB_EDate OR ReINT_GDB_EDate falls in that month - some
    re-initiated cases only have their date recorded in the
    ReINT_GDB_EDate column, not the original GDB_EDate.
    """

    primary = pd.to_datetime(df["GDB_EDate"], errors="coerce")

    if "ReINT_GDB_EDate" in df.columns:
        fallback = pd.to_datetime(df["ReINT_GDB_EDate"], errors="coerce")
    else:
        fallback = pd.Series(pd.NaT, index=df.index)

    primary_match = (primary.dt.month == month) & (primary.dt.year == year)
    fallback_match = (fallback.dt.month == month) & (fallback.dt.year == year)

    return primary_match | fallback_match


# =================================
# GDC - PRE JOINING
# =================================

pre_gdc = pre_df[
    pre_df["GDB_Result"]
    .astype(str)
    .str.upper()
    .isin(VALID_RESULTS)
].copy()

pre_gdc["GDB_EDate"] = pd.to_datetime(
    pre_gdc["GDB_EDate"],
    errors="coerce"
)

pre_gdc = pre_gdc[
    gdc_month_mask(pre_gdc, selected_month.month, selected_month.year)
]

pre_gdc = pre_gdc[
    pre_gdc["Vendor.3"]
    .astype(str)
    .str.strip()
    .str.upper()
    .isin(VENDOR_MATCH_VALUES)
]

# =================================
# GDC - PE & CVC
# =================================

pe_gdc = pe_df[
    pe_df["GDB_Result"]
    .astype(str)
    .str.upper()
    .isin(VALID_RESULTS)
].copy()

pe_gdc["GDB_EDate"] = pd.to_datetime(
    pe_gdc["GDB_EDate"],
    errors="coerce"
)

pe_gdc = pe_gdc[
    gdc_month_mask(pe_gdc, selected_month.month, selected_month.year)
]

pe_gdc = pe_gdc[
    pe_gdc["Vendor.3"]
    .astype(str)
    .str.strip()
    .str.upper()
    .isin(VENDOR_MATCH_VALUES)
]

# =================================
# GDC - FRESHER
# =================================

fresh_gdc = fresh_df[
    fresh_df["GDB_Result"]
    .astype(str)
    .str.upper()
    .isin(VALID_RESULTS)
].copy()

fresh_gdc["GDB_EDate"] = pd.to_datetime(
    fresh_gdc["GDB_EDate"],
    errors="coerce"
)

fresh_gdc = fresh_gdc[
    gdc_month_mask(fresh_gdc, selected_month.month, selected_month.year)
]

fresh_gdc = fresh_gdc[
    fresh_gdc["Vendor.2"]
    .astype(str)
    .str.strip()
    .str.upper()
    .isin(VENDOR_MATCH_VALUES)
]

# =================================
# FINAL GDC
# =================================

gdc_mt_df = pd.concat(
    [pre_gdc, pe_gdc, fresh_gdc],
    ignore_index=True
)

# ===========================
# OUTPUT HOLDERS
# ===========================

remarks_data = []

matched_ids = 0
vendor_mistakes = 0
spoc_mistakes = 0
missing_mt = 0

# ===========================
# VALIDATION FUNCTION
# ===========================

def check_mt_id(id_value, check_type):

    global matched_ids
    global vendor_mistakes
    global spoc_mistakes
    global missing_mt
    
    # =========================
    # CHECK CONCERN MT SHEET
    # =========================

    if check_type == "EDU":

        mt_case = edu_mt_df[
            (
                edu_mt_df["C_ID"].astype(str)
                == str(id_value)
            )
            |
            (
                edu_mt_df["Emp_No"].astype(str)
                == str(id_value)
            )
        ]

    elif check_type == "CRC":

        mt_case = crc_mt_df[
            (
                crc_mt_df["C_ID"].astype(str)
                == str(id_value)
            )
            |
            (
                crc_mt_df["Emp_No"].astype(str)
                == str(id_value)
            )
        ]

    elif check_type == "GDC":

        mt_case = gdc_mt_df[
            (
                gdc_mt_df["C_ID"].astype(str)
                == str(id_value)
            )
            |
            (
                gdc_mt_df["Emp_No"].astype(str)
                == str(id_value)
            )

        ]

    else:
        return

    # =========================
    # CRC WRONGLY INITIATED CHECK
    # =========================
    # If the candidate's Master Tracker "Checks to be Initiate" field
    # does NOT list CRC, then this CRC/Criminal case should never have
    # been opened - it was wrongly initiated by the SPOC.

    if check_type == "CRC":

        candidate_row = mt_df[
            (
                mt_df["C_ID"].astype(str)
                == str(id_value)
            )
            |
            (
                mt_df["Emp_No"].astype(str)
                == str(id_value)
            )
        ]

        if len(candidate_row) > 0 and CHECKS_TO_INITIATE_COLUMN_EXISTS:

            checks_to_initiate = str(
                candidate_row.iloc[0].get(
                    "Checks to be Initiate", ""
                )
            ).strip()

            if "crc" not in checks_to_initiate.lower():

                stopcheck_time = find_stopcheck_mail(
                    id_value,
                    "CRC"
                )

                vendor_report_time = find_vendor_report_mail(
                    id_value,
                    "CRC"
                )

                if stopcheck_time is None:

                    remarks_data.append({
                        "ID": id_value,
                        "Check": "CRC",
                        "Stopcheck DateTime": "NA",

                        "Vendor Report DateTime":
                        format_datetime(vendor_report_time),

                        "Remarks":
                        (
                            "Case wrongly initiated by SPOC and "
                            "stopcheck also not given"
                        ),
                        "Mistake On": "SPOC"
                    })

                    spoc_mistakes += 1

                    return

                remarks_data.append({
                    "ID": id_value,
                    "Check": "CRC",

                    "Stopcheck DateTime":
                    format_datetime(stopcheck_time),

                    "Vendor Report DateTime":
                    format_datetime(vendor_report_time),

                    "Remarks":
                    "Case wrongly initiated by SPOC",
                    "Mistake On": "SPOC"
                })

                spoc_mistakes += 1

                return

    # =========================
    # MATCHED
    # =========================

    if len(mt_case) > 0:

        row = mt_case.iloc[0]

        if check_type == "EDU":

            source_value = str(
                row.get("Ed_V_Source", "")
            ).strip()

            vendor_value = str(
                row.get("Vendor", "")
            ).strip()

            # Treat "NA" and "STOP CHECK" the same way - both mean the
            # SPOC is expected to have sent a stopcheck mail for this ID.
            if vendor_value.upper() in ["NA", "STOP CHECK", "STOPCHECK"]:

                stopcheck_time = find_stopcheck_mail(
                    id_value,
                    "EDU"
                )

                vendor_report_time = find_vendor_report_mail(
                    id_value,
                    "EDU"
                )

                if stopcheck_time is None:

                    remarks_data.append({
                        "ID": id_value,
                        "Check": "EDU",
                        "Stopcheck DateTime": "NA",

                        "Vendor Report DateTime":
                        format_datetime(vendor_report_time),

                        "Remarks":
                        (
                            f"Vendor is {vendor_value} | "
                            f"Source = {source_value} | "
                            f"Stopcheck not provided"
                        ),
                        "Mistake On": "SPOC"
                    })

                    spoc_mistakes += 1

                    return

                if vendor_report_time is not None:

                    remarks_data.append({
                        "ID": id_value,
                        "Check": "EDU",

                        "Stopcheck DateTime":
                        format_datetime(stopcheck_time),

                        "Vendor Report DateTime":
                        format_datetime(vendor_report_time),

                        "Remarks":
                        (
                            f"Vendor is {vendor_value} | "
                            f"Source = {source_value} | "
                            f"Stopcheck provided but vendor submitted report"
                        ),

                        "Mistake On": "Vendor"
                    })

                    vendor_mistakes += 1

                    return

                remarks_data.append({
                    "ID": id_value,
                    "Check": "EDU",

                    "Stopcheck DateTime":
                    format_datetime(stopcheck_time),

                    "Vendor Report DateTime":
                    "NA",

                    "Remarks":
                    (
                        f"Vendor is {vendor_value} | "
                        f"Source = {source_value} | "
                        f"Stopcheck provided"
                    ),

                    "Mistake On": "NA"
                })

                return

        stopcheck_display, vendor_display = lookup_mail_display(
            id_value, check_type
        )

        remarks_data.append({
            "ID": id_value,
            "Check": check_type,
            "Stopcheck DateTime": stopcheck_display,
            "Vendor Report DateTime": vendor_display,
            "Remarks": "Verified",
            "Mistake On": "NA"
        })

        matched_ids += 1
        return

    # =========================
    # FULL TRACKER SEARCH
    # =========================

    full_case = mt_df[
        (
            mt_df["C_ID"].astype(str)
            == str(id_value)
        )
        |
        (
            mt_df["Emp_No"].astype(str)
            == str(id_value)
        )
    ]
    
    if len(full_case) == 0:

        stopcheck_display, vendor_display = lookup_mail_display(
            id_value, check_type
        )

        remarks_data.append({
            "ID": id_value,
            "Check": check_type,
            "Stopcheck DateTime": stopcheck_display,
            "Vendor Report DateTime": vendor_display,
            "Remarks": "ID not available in Master Tracker",
            "Mistake On": "Vendor"
        })

        vendor_mistakes += 1
        missing_mt += 1

        return

    row = full_case.iloc[0]

    # =========================
    # MONTH MISMATCH - EDU
    # =========================

    if check_type == "EDU":

        full_month = pd.to_datetime(
            row.get("Ed_Edate", ""),
            errors="coerce"
        )

        if pd.notna(full_month):

            if (
                full_month.month != selected_month.month
                or
                full_month.year != selected_month.year
            ):

                stopcheck_display, vendor_display = lookup_mail_display(
                    id_value, "EDU"
                )

                remarks_data.append({
                    "ID": id_value,
                    "Check": "EDU",
                    "Stopcheck DateTime": stopcheck_display,
                    "Vendor Report DateTime": vendor_display,
                    "Remarks":
                    f"ID available in MT but belongs to {full_month.strftime('%b-%Y')}",
                    "Mistake On": "Vendor"
                })

                vendor_mistakes += 1
                return

    # =========================
    # EDU SPECIAL LOGIC
    # =========================

    if check_type == "EDU":

        source_value = str(
            row.get("Ed_V_Source", "")
        ).strip()

        vendor_value = str(
            row.get("Vendor", "")
        ).strip()

        # Treat "NA" and "STOP CHECK" the same way - both mean the SPOC
        # is expected to have sent a stopcheck mail for this ID. Handles
        # cases like tracker Vendor = "STOP CHECK" (was previously only
        # reported as a generic "vendor mismatch").
        if vendor_value.upper() in ["NA", "STOP CHECK", "STOPCHECK"]:

            stopcheck_time = find_stopcheck_mail(
                id_value,
                "EDU"
            )

            vendor_report_time = find_vendor_report_mail(
                id_value,
                "EDU"
            )
            
            # Stopcheck NOT provided

            if stopcheck_time is None:

                remarks_data.append({
                    "ID": id_value,
                    "Check": "EDU",
                    "Stopcheck DateTime": "NA",

                    "Vendor Report DateTime":
                    format_datetime(vendor_report_time),

                    "Remarks":
                    (
                        f"Vendor is {vendor_value} | "
                        f"Source = {source_value} | "
                        f"Stopcheck not provided"
                    ),
                    "Mistake On": "SPOC"
                })

                spoc_mistakes += 1

                return

            # Stopcheck provided but vendor submitted report

            if vendor_report_time is not None:

                remarks_data.append({
                    "ID": id_value,
                    "Check": "EDU",

                    "Stopcheck DateTime":
                    format_datetime(stopcheck_time),

                    "Vendor Report DateTime":
                    format_datetime(vendor_report_time),

                    "Remarks":
                    (
                        f"Vendor is {vendor_value} | "
                        f"Source = {source_value} | "
                        f"Stopcheck provided but vendor submitted report"
                    ),

                    "Mistake On": "Vendor"
                })

                vendor_mistakes += 1

                return

            # Stopcheck provided and no report

            remarks_data.append({
                "ID": id_value,
                "Check": "EDU",

                "Stopcheck DateTime":
                format_datetime(stopcheck_time),

                "Vendor Report DateTime":
                "NA",

                "Remarks":
                (
                    f"Vendor is {vendor_value} | "
                    f"Source = {source_value} | "
                    f"Stopcheck provided"
                ),

                "Mistake On": "NA"
            })

            return

        stopcheck_display, vendor_display = lookup_mail_display(
            id_value, "EDU"
        )

        if vendor_value.strip().upper() in VENDOR_MATCH_VALUES:

            remarks_data.append({
                "ID": id_value,
                "Check": "EDU",
                "Stopcheck DateTime": stopcheck_display,
                "Vendor Report DateTime": vendor_display,
                "Remarks": "Verified",
                "Mistake On": "NA"
            })

            matched_ids += 1

            return

        remarks_data.append({
            "ID": id_value,
            "Check": "EDU",
            "Stopcheck DateTime": stopcheck_display,
            "Vendor Report DateTime": vendor_display,
            "Remarks":
            f"Found in MT but vendor mismatch ({vendor_value})",
            "Mistake On": "Vendor"
        })

        vendor_mistakes += 1

        return

    # =========================
    # CRC MONTH MISMATCH
    # =========================

    if check_type == "CRC":

        full_month = pd.to_datetime(
            row.get("eP_CRC_EDate", ""),
            errors="coerce"
        )

        if pd.notna(full_month):

            if (
                full_month.month != selected_month.month
                or
                full_month.year != selected_month.year
            ):

                stopcheck_display, vendor_display = lookup_mail_display(
                    id_value, "CRC"
                )

                remarks_data.append({
                    "ID": id_value,
                    "Check": "CRC",
                    "Stopcheck DateTime": stopcheck_display,
                    "Vendor Report DateTime": vendor_display,
                    "Remarks":
                    f"ID available in MT but belongs to {full_month.strftime('%b-%Y')}",
                    "Mistake On": "Vendor"
                })

                vendor_mistakes += 1
                return

    # =========================
    # GDC MONTH MISMATCH
    # =========================

    if check_type == "GDC":

        full_month = pd.to_datetime(
            row.get("GDB_EDate", ""),
            errors="coerce"
        )

        full_month_reint = pd.to_datetime(
            row.get("ReINT_GDB_EDate", ""),
            errors="coerce"
        )

        primary_ok = (
            pd.notna(full_month)
            and full_month.month == selected_month.month
            and full_month.year == selected_month.year
        )

        reint_ok = (
            pd.notna(full_month_reint)
            and full_month_reint.month == selected_month.month
            and full_month_reint.year == selected_month.year
        )

        # Only a genuine month mismatch if NEITHER date falls in the
        # selected month - some re-initiated cases only have their
        # date recorded under ReINT_GDB_EDate, not GDB_EDate.
        if (pd.notna(full_month) or pd.notna(full_month_reint)) and not (
            primary_ok or reint_ok
        ):

            display_month = full_month if pd.notna(full_month) else full_month_reint

            stopcheck_display, vendor_display = lookup_mail_display(
                id_value, "GDC"
            )

            remarks_data.append({
                "ID": id_value,
                "Check": "GDC",
                "Stopcheck DateTime": stopcheck_display,
                "Vendor Report DateTime": vendor_display,
                "Remarks":
                f"ID available in MT but belongs to {display_month.strftime('%b-%Y')}",
                "Mistake On": "Vendor"
            })

            vendor_mistakes += 1
            return

    # =========================
    # CRC
    # =========================

    if check_type == "CRC":

        vendor_value = str(
            row.get("Vendor.2", "")
        ).strip()

        stopcheck_display, vendor_display = lookup_mail_display(
            id_value, "CRC"
        )

        if vendor_value.strip().upper() in VENDOR_MATCH_VALUES:

            remarks_data.append({
                "ID": id_value,
                "Check": "CRC",
                "Stopcheck DateTime": stopcheck_display,
                "Vendor Report DateTime": vendor_display,
                "Remarks": "Verified",
                "Mistake On": "NA"
            })

            matched_ids += 1

            return

        remarks_data.append({
            "ID": id_value,
            "Check": "CRC",
            "Stopcheck DateTime": stopcheck_display,
            "Vendor Report DateTime": vendor_display,
            "Remarks":
            f"ID available in MT under Vendor {vendor_value}",
            "Mistake On": "Vendor"
        })

        vendor_mistakes += 1
        return

    # =========================
    # GDC
    # =========================

    if check_type == "GDC":

        vendor_value = str(
            row.get("Vendor.3", "")
        ).strip()

        stopcheck_display, vendor_display = lookup_mail_display(
            id_value, "GDC"
        )

        if vendor_value.strip().upper() in VENDOR_MATCH_VALUES:

            remarks_data.append({
                "ID": id_value,
                "Check": "GDC",
                "Stopcheck DateTime": stopcheck_display,
                "Vendor Report DateTime": vendor_display,
                "Remarks": "Verified",
                "Mistake On": "NA"
            })

            matched_ids += 1

            return

        remarks_data.append({
            "ID": id_value,
            "Check": "GDC",
            "Stopcheck DateTime": stopcheck_display,
            "Vendor Report DateTime": vendor_display,
            "Remarks":
            f"ID available in MT under Vendor {vendor_value}",
            "Mistake On": "Vendor"
        })

        vendor_mistakes += 1
        return

# ===========================
# BUILD MAIL INDEX (ONE-TIME SCAN)
# ===========================

# ===========================
# EDUCATION
# ===========================

for _, row in edu_df.iterrows():

    check_id = str(
        row["Check id"]
    ).strip()

    check_id = check_id.replace(
        ".0",
        ""
    )
   
    check_mt_id(check_id, "EDU")

# CRC

for _, row in crc_df.iterrows():

    check_id = str(
        row["Ref. No."]
    ).strip()

    check_id = check_id.replace(
        ".0",
        ""
    )

    check_mt_id(
        check_id,
        "CRC"
    )

# Criminal

for _, row in criminal_df.iterrows():

    check_id = str(
        row["Client Ref No"]
    ).strip()

    check_id = check_id.replace(
        ".0",
        ""
    )

    check_mt_id(
        check_id,
        "CRC"
    )

# ===========================
# GDC
# ===========================

for _, row in global_df.iterrows():
    check_id = str(
        row["Check ID"]
    ).strip()

    check_id = check_id.replace(
        ".0",
        ""
    )
    check_mt_id(check_id, "GDC")

# ===========================
# REMARKS SHEET
# ===========================

remarks_df = pd.DataFrame(
    remarks_data,
    columns=[
        "ID",
        "Check",
        "Stopcheck DateTime",
        "Vendor Report DateTime",
        "Remarks",
        "Mistake On"
    ]
)

if not remarks_df.empty:
    remarks_df.insert(
        0,
        "S.No",
        range(1, len(remarks_df) + 1)
    )

# ===========================
# SUMMARY
# ===========================

summary_df = pd.DataFrame({
    "Metric": [
        "Total Edu IDs Validated",
        "Total CRC IDs Validated",
        "Total GDC IDs Validated",
        "Matched IDs",
        "Missing in MT",
        "Vendor Mistakes",
        "SPOC Mistakes"
    ],
    "Value": [
        len(edu_df),
        len(crc_final_df),
        len(global_df),
        matched_ids,
        missing_mt,
        vendor_mistakes,
        spoc_mistakes
    ]
})

summary_df.insert(
    0,
    "S.No",
    range(1, len(summary_df) + 1)
)

edu_export = pd.DataFrame({
    "Sl.No": range(1, len(edu_df) + 1),
    "Check ID": edu_df["Check id"]
})

crc_export = pd.DataFrame({
    "Sl.No": range(1, len(crc_final_df) + 1),
    "Check ID": pd.concat([
        crc_df["Ref. No."],
        criminal_df["Client Ref No"]
    ], ignore_index=True)
})

gdc_export = pd.DataFrame({
    "Sl.No": range(1, len(global_df) + 1),
    "Check ID": global_df["Check ID"]
})

edu_mt_df = edu_mt_df[
    [
        "C_ID",
        "Emp_No",
        "C_Name",
        "Ed_V_Source",
        "Ed_Result",
        "Ed_Edate",
        "Vendor",
        "Closed By"
    ]
]

edu_mt_df.insert(
    0,
    "S.No",
    range(1, len(edu_mt_df) + 1)
)

edu_mt_df["Ed_Edate"] = pd.to_datetime(
    edu_mt_df["Ed_Edate"],
    errors="coerce"
).dt.strftime("%d-%b-%Y")

crc_mt_df = crc_mt_df[
    [
        "C_ID",
        "Emp_No",
        "C_Name",
        "eP_CRC_Result",
        "eP_CRC_EDate",
        "Vendor.2"
    ]
]

crc_mt_df.insert(
    0,
    "S.No",
    range(1, len(crc_mt_df) + 1)
)

crc_mt_df["eP_CRC_EDate"] = pd.to_datetime(
    crc_mt_df["eP_CRC_EDate"],
    errors="coerce"
).dt.strftime("%d-%b-%Y")

crc_mt_df.rename(
    columns={
        "Vendor.2": "Vendor"
    },
    inplace=True
)

for col in [
    "ReINT_GDB_Result",
    "ReINT_GDB_EDate",
    "Vendor.4"
]:
    if col not in gdc_mt_df.columns:
        gdc_mt_df[col] = ""

gdc_mt_df = gdc_mt_df[
    [
        "C_ID",
        "Emp_No",
        "C_Name",
        "GDB_Result",
        "GDB_EDate",
        "Vendor.3",
        "ReINT_GDB_Result",
        "ReINT_GDB_EDate",
        "Vendor.4"
    ]
]

gdc_mt_df.insert(
    0,
    "S.No",
    range(1, len(gdc_mt_df) + 1)
)

gdc_mt_df["GDB_EDate"] = pd.to_datetime(
    gdc_mt_df["GDB_EDate"],
    errors="coerce"
).dt.strftime("%d-%b-%Y")

gdc_mt_df["ReINT_GDB_EDate"] = pd.to_datetime(
    gdc_mt_df["ReINT_GDB_EDate"],
    errors="coerce"
).dt.strftime("%d-%b-%Y")

gdc_mt_df.rename(
    columns={
        "Vendor.3": "GDB Vendor",
        "Vendor.4": "ReINT GDB Vendor"
    },
    inplace=True
)

# ===========================
# EXPORT
# ===========================

output_file = os.environ.get("VENDOR_VALIDATION_OUTPUT", "Validation_Output.xlsx")
with pd.ExcelWriter(output_file,
                    engine="openpyxl") as writer:

    edu_export.to_excel(writer,
                    sheet_name="Edu - DS",
                    index=False)

    crc_export.to_excel(writer,
                    sheet_name="CRC - DS",
                    index=False)

    gdc_export.to_excel(writer,
                    sheet_name="GDC - DS",
                    index=False)
    
    edu_mt_df.to_excel(
        writer,
        sheet_name="Edu - MT",
        index=False
    )

    crc_mt_df.to_excel(
        writer,
        sheet_name="CRC - MT",
        index=False
    )

    gdc_mt_df.to_excel(
        writer,
        sheet_name="GDC - MT",
        index=False
    )

    remarks_df.to_excel(writer,
                        sheet_name="O_Remarks",
                        index=False)

    summary_df.to_excel(writer,
                        sheet_name="Validation Summary",
                        index=False)

    # =================================
    # FORMAT ALL SHEETS
    # =================================

    wb = writer.book

    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    header_fill = PatternFill(
        start_color="D9EAF7",
        end_color="D9EAF7",
        fill_type="solid"
    )

    for ws in wb.worksheets:

        ws.freeze_panes = "A2"

        # Header Formatting
        for cell in ws[1]:

            cell.font = Font(bold=True)

            cell.fill = header_fill

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

            cell.border = thin_border
            
        # All Cell Formatting
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )
                cell.border = thin_border

        # AutoFit
        for col in ws.columns:

            max_length = 0

            column = get_column_letter(
                col[0].column
            )

            for cell in col:

                try:
                    if cell.value is not None:
                        max_length = max(
                            max_length,
                            len(str(cell.value))
                        )
                except:
                    pass

            ws.column_dimensions[
                column
            ].width = max_length + 3

        # Filter
        ws.auto_filter.ref = ws.dimensions

print(f"Validation completed successfully for {VENDOR}: {output_file}")

