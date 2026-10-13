"""Check fictional report rows for basic completeness and valid values."""
from __future__ import annotations

import csv
from datetime import date
from pathlib import Path

HERE = Path(__file__).resolve().parent
SAMPLES = HERE / "sample_data"
OUTPUT = HERE / "output" / "report_qc_issues.csv"
VALID_STATUSES = {"completed", "in progress", "on hold"}


def main() -> None:
    with (SAMPLES / "reports.csv").open(newline="", encoding="utf-8-sig") as stream:
        rows = list(csv.DictReader(stream))
    issues: list[dict[str, str]] = []
    for row in rows:
        reasons: list[str] = []
        if not row.get("applicant_id", "").strip():
            reasons.append("Missing applicant ID")
        if row.get("status", "").strip().casefold() not in VALID_STATUSES:
            reasons.append("Unsupported or missing status")
        try:
            date.fromisoformat(row.get("report_date", "").strip())
        except ValueError:
            reasons.append("Invalid report date (expected YYYY-MM-DD)")
        if reasons:
            issues.append({"applicant_id": row.get("applicant_id", ""), "issues": "; ".join(reasons)})
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    with OUTPUT.open("w", newline="", encoding="utf-8") as stream:
        writer = csv.DictWriter(stream, fieldnames=["applicant_id", "issues"])
        writer.writeheader()
        writer.writerows(issues)
    print(f"Checked {len(rows)} report rows; found {len(issues)} row(s) for review: {OUTPUT}")


if __name__ == "__main__":
    main()
