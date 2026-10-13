import pyperclip
import re

input(
    "Copy DIT Summary and press Enter..."
)

text = pyperclip.paste()

cid_total = 0
eid_total = 0

for line in text.splitlines():

    line = line.strip()

    if not line:
        continue

    # CID
    if "CID" in line.upper():

        match = re.search(
            r"(\d+)$",
            line
        )

        if match:
            cid_total += int(match.group(1))

    # EMP + STOP CHECK
    elif (
        "EMP" in line.upper()
        or "STOP CHECK" in line.upper()
    ):

        match = re.search(
            r"(\d+)$",
            line
        )

        if match:
            eid_total += int(match.group(1))

print("\n====================")
print(f"CID Report Count : {cid_total}")
print(f"EID Report Count : {eid_total}")
print("====================")