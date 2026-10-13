import os
import re
import glob
import html
import pyperclip
import pyautogui
import time
import win32clipboard

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


def copy_teams_message_with_bold_headings(text):
    """Put plain text and rich HTML on the clipboard for Teams paste."""
    bold_headings = {"LATERAL", "FRESHER", "CVC"}
    html_lines = []
    for line in text.splitlines():
        escaped_line = html.escape(line)
        if line.strip().upper() in bold_headings:
            escaped_line = f"<strong>{escaped_line}</strong>"
        html_lines.append(escaped_line)

    fragment = "<br>".join(html_lines)
    html_prefix = (
        "<!DOCTYPE html><html><body><!--StartFragment-->"
    )
    html_suffix = "<!--EndFragment--></body></html>"
    html_document = html_prefix + fragment + html_suffix

    # CF_HTML offsets are byte offsets from the start of the clipboard data.
    header_template = (
        "Version:0.9\r\n"
        "StartHTML:{start_html:010d}\r\n"
        "EndHTML:{end_html:010d}\r\n"
        "StartFragment:{start_fragment:010d}\r\n"
        "EndFragment:{end_fragment:010d}\r\n"
    )
    header_placeholder = header_template.format(
        start_html=0, end_html=0, start_fragment=0, end_fragment=0
    ).encode("ascii")
    prefix_bytes = html_prefix.encode("utf-8")
    fragment_bytes = fragment.encode("utf-8")
    suffix_bytes = html_suffix.encode("utf-8")
    start_html = len(header_placeholder)
    start_fragment = start_html + len(prefix_bytes)
    end_fragment = start_fragment + len(fragment_bytes)
    end_html = end_fragment + len(suffix_bytes)
    header = header_template.format(
        start_html=start_html,
        end_html=end_html,
        start_fragment=start_fragment,
        end_fragment=end_fragment,
    ).encode("ascii")
    html_clipboard_data = header + html_document.encode("utf-8")

    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, text)
        html_format = win32clipboard.RegisterClipboardFormat("HTML Format")
        win32clipboard.SetClipboardData(html_format, html_clipboard_data)
    finally:
        win32clipboard.CloseClipboard()

# =====================================================
# COPY RDT MESSAGE FIRST
# =====================================================

input(
    "Copy Consolidated RDT message from Chat A and press Enter..."
)

copied_text = pyperclip.paste()

# =====================================================
# EXTRACT RDT DETAILS
# =====================================================

date_match = re.search(
    r"Consolidated RDT-\s*(\d{2}-[A-Za-z]{3}-\d{4})",
    copied_text
)

total_match = re.search(
    r"Total Reports\s*-\s*(\d+)",
    copied_text
)

if not date_match:
    raise Exception("RDT Date not found")

if not total_match:
    raise Exception("Total Reports not found")

rdt_date = date_match.group(1)
total_reports = total_match.group(1)

if "Lateral" not in copied_text:
    raise Exception("Lateral section not found")

rdt_content = copied_text.split(
    "Lateral",
    1
)[1]

rdt_content = "LATERAL\n" + rdt_content

# Teams chat-la varra <br> text remove
rdt_content = rdt_content.replace(
    "<br>",
    "\n"
)

# Heading formatting
rdt_content = rdt_content.replace(
    "\nCVC",
    "\n\nCVC"
)

rdt_content = rdt_content.replace(
    "\nFresher",
    "\n\nFRESHER"
)

# =====================================================
# FILES FOR CHAT 1
# =====================================================

master_tracker_1 = latest_master_tracker()

datasheet = max(
    glob.glob(
        os.path.join(
            downloads,
            "Infosys_Conventional-Datasheet*"
        )
    ),
    key=os.path.getmtime
)

dit_count = max(
    glob.glob(
        os.path.join(
            downloads,
            "*DIT*"
        )
    ),
    key=os.path.getmtime
)

# =====================================================
# CHAT 1
# =====================================================

chat_1 = r"https://teams.microsoft.com/l/chat/19:dc6f8d91-afb8-499f-9b44-8450c52e9bdf_fcb507e9-08da-4e67-8e6f-d9003fcdadfe@unq.gbl.spaces/conversations?context=%7B%22contextType%22%3A%22chat%22%7D"

message_1 = f"""Hi,

Attached ({rdt_date}) Daily wise DIT count, Datasheet & MT FYR.
"""

os.startfile(chat_1)

time.sleep(8)

pyautogui.click(1636, 957)

time.sleep(2)

pyautogui.click(1614, 909)

time.sleep(4)

pyautogui.hotkey("alt", "n")

time.sleep(1)

files_text = (
    f'"{master_tracker_1}" '
    f'"{datasheet}" '
    f'"{dit_count}"'
)

pyperclip.copy(files_text)

pyautogui.hotkey("ctrl", "v")

time.sleep(1)

pyautogui.press("enter")

time.sleep(10)

pyautogui.click(677, 643)

time.sleep(1)

pyperclip.copy(message_1)
pyautogui.hotkey("ctrl", "v")

time.sleep(10)

# SEND CHAT 1
pyautogui.click(1765, 959)

# =====================================================
# FILES FOR CHAT 2
# =====================================================

master_tracker_2 = latest_master_tracker()

dispatch_tracker = max(
    glob.glob(
        os.path.join(
            downloads,
            "Report Dispatch Tracker*"
        )
    ),
    key=os.path.getmtime
)

reports_zip = max(
    glob.glob(
        os.path.join(
            downloads,
            "Reports*"
        )
    ),
    key=os.path.getmtime
)

# =====================================================
# CHAT 2
# =====================================================

chat_2 = r"https://teams.microsoft.com/l/chat/19:75f7d7944ee344e384faec25b5cc0c17@thread.v2/conversations?context=%7B%22contextType%22%3A%22chat%22%7D"

message_2 = f"""Hi Team!!

Attached {rdt_date} Reports (Total - {total_reports}) & Dispatch and Master Tracker FYR.

{rdt_content}
"""

time.sleep(5)

os.startfile(chat_2)

time.sleep(8)

pyautogui.click(1638, 959)

time.sleep(2)

pyautogui.click(1625, 909)

time.sleep(4)

pyautogui.hotkey("alt", "n")

time.sleep(1)

files_text = (
    f'"{master_tracker_2}" '
    f'"{dispatch_tracker}" '
    f'"{reports_zip}"'
)

pyperclip.copy(files_text)

pyautogui.hotkey("ctrl", "v")

time.sleep(1)

pyautogui.press("enter")

time.sleep(10)

pyautogui.click(679, 646)

time.sleep(1)

copy_teams_message_with_bold_headings(message_2)
pyautogui.hotkey("ctrl", "v")

time.sleep(10)

# SEND CHAT 2
pyautogui.click(1766, 957)

print("✅ CHAT 1 COMPLETED")
print("✅ CHAT 2 COMPLETED")
print("🚀 ALL TEAMS AUTOMATIONS COMPLETED")
