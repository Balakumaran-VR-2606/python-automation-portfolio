"""Run selected portfolio demonstrations in a predictable sequence."""
from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TASKS = {
    "vendor-validation": ROOT / "projects" / "vendor_reconciliation" / "validate.py",
    "report-quality": ROOT / "projects" / "report_qc" / "check_reports.py",
}


def run(tasks: list[str]) -> int:
    results: list[tuple[str, int]] = []
    for name in tasks:
        script = TASKS[name]
        print(f"\n--- Starting: {name} ---", flush=True)
        completed = subprocess.run([sys.executable, str(script)], cwd=ROOT, check=False)
        results.append((name, completed.returncode))
        state = "completed" if completed.returncode == 0 else f"failed ({completed.returncode})"
        print(f"--- {name}: {state} ---", flush=True)
    print("\nRun summary")
    for name, code in results:
        print(f"{'PASS' if code == 0 else 'FAIL'}  {name}")
    return 0 if all(code == 0 for _, code in results) else 1


def main() -> int:
    parser = argparse.ArgumentParser(description="Run portfolio workflow demonstrations.")
    parser.add_argument("tasks", nargs="*", choices=sorted(TASKS), help="Tasks to run; defaults to all.")
    args = parser.parse_args()
    return run(args.tasks or list(TASKS))


if __name__ == "__main__":
    raise SystemExit(main())
