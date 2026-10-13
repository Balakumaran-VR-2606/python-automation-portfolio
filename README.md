# Balakumaran VR - Python Automation Portfolio

Built by Balakumaran VR. LinkedIn: https://www.linkedin.com/in/balakumaran-vr-367a93279.

Selected examples of Python automation work focused on reducing repetitive operations, validating data, and improving document workflows. The examples in this repository are generalized portfolio demonstrations and use fictional sample data. They do not contain employer, client, vendor, candidate, or mailbox data.

## Projects

### Unified Operations Automation Hub
A lightweight command-line orchestrator that runs selected workflow tasks in sequence, records each task's result, continues through the run, and summarizes failures.

### Privacy-Focused PDF Date-of-Birth Redaction
A PyMuPDF utility that finds supported date-of-birth labels in searchable PDF text, applies PDF redactions, and writes separate sanitized copies. Scanned image-only PDFs require OCR before text-based matching.

### Vendor Validation and Tracker Reconciliation
A reusable CSV demonstration that compares applicant identifiers, vendor names, and check types across a source list and a master tracker, then creates an exception-ready output file.

### Background Verification Report Quality Control
A small data-quality demonstration that checks report records for missing identifiers, unsupported statuses, and invalid dates and writes a review file.

## Other workflow experience

Additional internal tools developed for tracker refreshes, progress reporting, operational summaries, report distribution, email drafting, and case document retrieval. Production-specific integrations and business rules are intentionally not included in this public-safe sample repository.

## Getting started

1. Install Python 3.10 or later.
2. Install the PDF dependency only if you want to use the redaction utility:

   ```powershell
   python -m pip install -r requirements.txt
   ```

3. Run the demonstrations from the repository root:

   ```powershell
   python automation_hub/automation_hub.py
   python projects/vendor_reconciliation/validate.py
   python projects/report_qc/check_reports.py
   python projects/pdf_redaction/redact_dob.py --input-folder path\to\input --output-folder path\to\sanitized
   ```

The two data demonstrations use the fictional CSV files under their `sample_data` folders. Generated files go into ignored `output` folders.

## Safety and data handling

Use only files you are authorized to process. Do not add credentials, mailbox exports, candidate records, employer/client documents, or production spreadsheets to this repository. Keep personal configuration in environment variables or local files excluded by `.gitignore`.
