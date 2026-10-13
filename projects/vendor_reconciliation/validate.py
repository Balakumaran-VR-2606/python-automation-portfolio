"""Compare fictional vendor source records with a master tracker."""
from __future__ import annotations

import csv
from pathlib import Path

HERE = Path(__file__).resolve().parent
SAMPLES = HERE / "sample_data"
OUTPUT = HERE / "output" / "validation_results.csv"


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open(newline="", encoding="utf-8-sig") as stream:
        return list(csv.DictReader(stream))


def key(row: dict[str, str]) -> tuple[str, str, str]:
    return tuple(row.get(field, "").strip().casefold() for field in ("applicant_id", "vendor", "check_type"))


def main() -> None:
    source = read_csv(SAMPLES / "source_candidates.csv")
    tracker = read_csv(SAMPLES / "master_tracker.csv")
    tracker_keys = {key(row) for row in tracker}
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    with OUTPUT.open("w", newline="", encoding="utf-8") as stream:
        fields = ["applicant_id", "vendor", "check_type", "result"]
        writer = csv.DictWriter(stream, fieldnames=fields)
        writer.writeheader()
        for row in source:
            result = "MATCHED" if key(row) in tracker_keys else "REVIEW - NOT FOUND"
            writer.writerow({field: row.get(field, "").strip() for field in fields[:-1]} | {"result": result})
    print(f"Reviewed {len(source)} source rows. Results: {OUTPUT}")


if __name__ == "__main__":
    main()
