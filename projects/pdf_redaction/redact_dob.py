"""Redact supported Date of Birth text values from searchable PDFs."""
from __future__ import annotations

import argparse
import re
from pathlib import Path

import fitz

DATE_PATTERNS = (
    re.compile(r"Date of Birth\s+\d{1,2}\s+[A-Za-z]{3}\s+\d{4}", re.IGNORECASE),
    re.compile(r"Date of Birth\s+\d{1,2}[-/]\d{1,2}[-/]\d{4}", re.IGNORECASE),
)


def redact_pdf(source: Path, destination: Path) -> int:
    redactions = 0
    with fitz.open(source) as document:
        for page in document:
            text = page.get_text("text") or ""
            for pattern in DATE_PATTERNS:
                for match in pattern.finditer(text):
                    for rectangle in page.search_for(match.group(0)):
                        page.add_redact_annot(rectangle, fill=(1, 1, 1))
                        redactions += 1
            page.apply_redactions()
        destination.parent.mkdir(parents=True, exist_ok=True)
        document.save(destination)
    return redactions


def main() -> int:
    parser = argparse.ArgumentParser(description="Redact supported DOB text from searchable PDF files.")
    parser.add_argument("--input-folder", type=Path, required=True)
    parser.add_argument("--output-folder", type=Path, required=True)
    args = parser.parse_args()
    if not args.input_folder.is_dir():
        parser.error("input folder does not exist")
    args.output_folder.mkdir(parents=True, exist_ok=True)
    files = sorted(args.input_folder.glob("*.pdf"))
    if not files:
        print("No PDF files found.")
        return 0
    for source in files:
        destination = args.output_folder / source.name
        count = redact_pdf(source, destination)
        print(f"Processed {source.name}: {count} redaction(s)")
    print("Note: image-only scanned PDFs need OCR before text-based DOB matching.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
