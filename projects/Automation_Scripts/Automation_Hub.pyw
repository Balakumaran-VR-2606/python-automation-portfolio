"""Tkinter launcher for the Infosys QC automation scripts.

Place this file beside Master_Automation.py and the automation scripts, then
run it with ``python Automation_Hub.py``. Child scripts run sequentially in
their existing order. Their prompts and output appear in the log panel.
"""

from __future__ import annotations

import os
import json
import queue
import subprocess
import sys
import threading
from datetime import datetime
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk


AUTOMATIONS = [
    ("Live DIT Count", "Live_DIT_Count.py"),
    ("Report Count", "Report_Count.py"),
    ("DIT Open", "DIT_Open.py"),
    ("Overall Report QC", "Overall_Report_QC.py"),
    ("Validation & Mail Automation", "Report+RDT+MT_Validation+MailAutomation.py"),
    ("OneDrive Upload", "1Drive_Report_Upload.py"),
    ("Datasheet", "Datasheet.py"),
    ("RDT", "RDT.py"),
    ("Progress Sheet", "PS.py"),
    ("Cleanup", "Cleanup.py"),
    ("DS + Live DIT Validation", "DS+Live_DIT.py"),
    ("B Mail", "B_Mail.py"),
    ("Teams Automation", "TS.py"),
]

# kind: entry, choice, or multiline. Answers are matched to each child's
# input(prompt) text so conditional prompts do not shift later answers.
INPUT_SPECS = {
    "Live_DIT_Count.py": [
        ("month", "Month-Year (MMM-YYYY)", "entry", ""),
        ("dashboard_date", "Dashboard Date (DD-MMM-YYYY)", "entry", ""),
    ],
    "Report_Count.py": [
        ("dit_summary", "Paste DIT Summary from clipboard", "multiline", ""),
    ],
    "Report+RDT+MT_Validation+MailAutomation.py": [
        ("validation_mode", "Validation dates: 1 Today, 2 Yesterday, 3 Custom", "choice", "1"),
        ("validation_dates", "Custom date(s), DD-MMM-YYYY, comma-separated (Option 3)", "entry", ""),
        ("mail_subject_date", "Mail subject date (DD-MMM-YYYY)", "entry", ""),
        ("mail_note", "Optional mail Note (leave blank to omit from body)", "multiline", ""),
    ],
    "1Drive_Report_Upload.py": [
        ("drive_folders", "Choose report folders", "multichoice", ""),
        ("drive_date", "Report Date (DD-MMM-YYYY)", "entry", ""),
    ],
    "Datasheet.py": [
        ("datasheet_date", "Dashboard Date (DD-MMM-YYYY)", "entry", ""),
    ],
    "Cleanup.py": [("cleanup_date", "Cleanup Date (DD-MMM-YYYY)", "entry", "")],
    "TS.py": [
        ("teams_message", "Paste Consolidated RDT message from Chat A", "multiline", ""),
    ],
}

DRIVE_FOLDER_OPTIONS = [
    ("1", "INTERIM"),
    ("2", "FINAL"),
    ("3", "PRE FINAL"),
    ("4", "STOP CHECK"),
    ("5", "Supplementary Reports"),
    ("6", "Supplementary 2 Reports"),
    ("7", "Supplementary 3 Reports"),
]


RUNNER = r'''import builtins, json, os, runpy, sys
script_path = os.path.abspath(sys.argv[1])
sys.argv[:] = [script_path]
answers = json.loads(os.environ.pop("AUTOMATION_HUB_ANSWERS", "{}"))
clipboard_text = os.environ.pop("AUTOMATION_HUB_CLIPBOARD", None)
original_input = builtins.input
def gui_input(prompt=""):
    if prompt:
        sys.stdout.write(prompt)
        sys.stdout.flush()
    lowered = prompt.lower()
    if clipboard_text is not None and (
        "copy dit summary" in lowered
        or "copy consolidated rdt message" in lowered
    ):
        sys.stdout.write(" [clipboard text provided in setup]\n")
        sys.stdout.flush()
        return ""
    for fragment, answer in answers.items():
        if fragment.lower() in lowered:
            if "press enter" not in lowered:
                sys.stdout.write(" [provided in setup]\n")
                sys.stdout.flush()
            return answer
    return original_input()
builtins.input = gui_input
previous_clipboard = None
clipboard_changed = False
if clipboard_text is not None:
    import pyperclip
    previous_clipboard = pyperclip.paste()
    pyperclip.copy(clipboard_text)
    clipboard_changed = True
try:
    runpy.run_path(script_path, run_name="__main__")
finally:
    if clipboard_changed:
        import pyperclip
        pyperclip.copy(previous_clipboard)
'''


class AutomationHub(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Infosys QC Automation Hub")
        self.geometry("1080x760")
        self.minsize(850, 620)

        self.script_dir = Path(__file__).resolve().parent
        self.events: queue.Queue[tuple] = queue.Queue()
        self.process: subprocess.Popen | None = None
        self.process_lock = threading.Lock()
        self.running = False
        self.closing = False
        self.failed_scripts: list[tuple[str, str]] = []
        self.status_vars: dict[str, tk.StringVar] = {}
        self.status_labels: dict[str, ttk.Label] = {}
        self.selected_vars: dict[str, tk.BooleanVar] = {}
        self._last_output_tail = ""

        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(60, self._drain_events)

    def _build_ui(self) -> None:
        self.title("Infosys QC Automation Hub")
        self.geometry("1180x900")
        self.minsize(980, 740)
        self.configure(bg="#eef3f8")

        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("App.TFrame", background="#eef3f8")
        style.configure("Panel.TFrame", background="#ffffff")
        style.configure("Title.TLabel", background="#12365a", foreground="#ffffff", font=("Segoe UI", 21, "bold"))
        style.configure("Subtitle.TLabel", background="#12365a", foreground="#d8e7f5", font=("Segoe UI", 10))
        style.configure("Section.TLabel", background="#ffffff", foreground="#183b5b", font=("Segoe UI", 12, "bold"))
        style.configure("Muted.TLabel", background="#ffffff", foreground="#64748b", font=("Segoe UI", 9))
        style.configure("Count.TLabel", background="#ffffff", foreground="#2563a6", font=("Segoe UI", 9, "bold"))
        style.configure("Task.TCheckbutton", background="#ffffff", foreground="#203247", font=("Segoe UI", 10))
        style.map("Task.TCheckbutton", background=[("active", "#f4f8fc")])
        style.configure("Ready.TLabel", background="#eaf1f8", foreground="#42617c", padding=(8, 2), font=("Segoe UI", 8, "bold"))
        style.configure("Queued.TLabel", background="#e8efff", foreground="#3857a6", padding=(8, 2), font=("Segoe UI", 8, "bold"))
        style.configure("Running.TLabel", background="#fff2d8", foreground="#9a5b00", padding=(8, 2), font=("Segoe UI", 8, "bold"))
        style.configure("Completed.TLabel", background="#e1f5ea", foreground="#197346", padding=(8, 2), font=("Segoe UI", 8, "bold"))
        style.configure("Failed.TLabel", background="#fde8e7", foreground="#ad3028", padding=(8, 2), font=("Segoe UI", 8, "bold"))
        style.configure("Stopped.TLabel", background="#eef0f3", foreground="#596579", padding=(8, 2), font=("Segoe UI", 8, "bold"))
        style.configure("Primary.TButton", background="#1769aa", foreground="#ffffff", font=("Segoe UI", 10, "bold"), padding=(16, 9), borderwidth=0)
        style.map("Primary.TButton", background=[("disabled", "#a8bdd0"), ("active", "#12558a")], foreground=[("disabled", "#eef3f8")])
        style.configure("Secondary.TButton", background="#e5edf5", foreground="#24435f", font=("Segoe UI", 9, "bold"), padding=(10, 7), borderwidth=0)
        style.map("Secondary.TButton", background=[("active", "#d5e2ef")])
        style.configure("Danger.TButton", background="#fde8e7", foreground="#ad3028", font=("Segoe UI", 9, "bold"), padding=(10, 7), borderwidth=0)
        style.map("Danger.TButton", background=[("active", "#f8d4d2")])
        style.configure("Input.TButton", background="#1769aa", foreground="#ffffff", font=("Segoe UI", 9, "bold"), padding=(14, 7), borderwidth=0)
        style.map("Input.TButton", background=[("disabled", "#a8bdd0"), ("active", "#12558a")])

        header = tk.Frame(self, bg="#12365a", padx=24, pady=17)
        header.pack(fill="x")
        title_area = tk.Frame(header, bg="#12365a")
        title_area.pack(side="left", fill="x", expand=True)
        ttk.Label(title_area, text="INFOSYS QC AUTOMATION", style="Title.TLabel").pack(anchor="w")
        ttk.Label(title_area, text="One place to prepare, run, and monitor your QC workflows.", style="Subtitle.TLabel").pack(anchor="w", pady=(3, 0))
        self.header_status = tk.StringVar(value="●  READY")
        tk.Label(header, textvariable=self.header_status, bg="#1d4d78", fg="#eaf4ff", font=("Segoe UI", 9, "bold"), padx=12, pady=7).pack(side="right", anchor="center")

        outer = ttk.Frame(self, style="App.TFrame", padding=(18, 15, 18, 12))
        outer.pack(fill="both", expand=True)

        task_panel = ttk.Frame(outer, style="Panel.TFrame", padding=(14, 11, 14, 13))
        task_panel.pack(fill="x", pady=(0, 11))
        task_heading = ttk.Frame(task_panel, style="Panel.TFrame")
        task_heading.pack(fill="x", pady=(0, 7))
        ttk.Label(task_heading, text="Choose automations", style="Section.TLabel").pack(side="left")
        self.selection_count = tk.StringVar(value="13 of 13 selected")
        ttk.Label(task_heading, textvariable=self.selection_count, style="Count.TLabel").pack(side="right")

        task_grid = ttk.Frame(task_panel, style="Panel.TFrame")
        task_grid.pack(fill="x")
        for column in range(2):
            task_grid.columnconfigure(column, weight=1, uniform="task-column")
        for index, (label, filename) in enumerate(AUTOMATIONS):
            self.selected_vars[filename] = tk.BooleanVar(value=True)
            self.status_vars[filename] = tk.StringVar(value="Ready")
            self.selected_vars[filename].trace_add("write", lambda *_: self._update_selection_count())
            row, col = divmod(index, 2)
            card = ttk.Frame(task_grid, style="Panel.TFrame", padding=(7, 3))
            card.grid(row=row, column=col, sticky="ew", padx=(0, 10) if col == 0 else (10, 0), pady=2)
            card.columnconfigure(0, weight=1)
            ttk.Checkbutton(card, text=label, variable=self.selected_vars[filename], style="Task.TCheckbutton").grid(row=0, column=0, sticky="w")
            status_label = ttk.Label(card, textvariable=self.status_vars[filename], style="Ready.TLabel")
            status_label.grid(row=0, column=1, sticky="e", padx=(8, 0))
            self.status_labels[filename] = status_label

        controls_panel = ttk.Frame(outer, style="Panel.TFrame", padding=(11, 9))
        controls_panel.pack(fill="x", pady=(0, 11))
        self.run_button = ttk.Button(controls_panel, text="▶  Run Selected (13)", style="Primary.TButton", command=self._run_selected)
        self.run_button.pack(side="left", padx=(0, 8))
        self.rerun_button = ttk.Button(controls_panel, text="↻  Rerun Failed", style="Secondary.TButton", command=self._rerun_failed, state="disabled")
        self.rerun_button.pack(side="left", padx=(0, 8))
        self.stop_button = ttk.Button(controls_panel, text="■  Stop Current", style="Danger.TButton", command=self._stop_current, state="disabled")
        self.stop_button.pack(side="left", padx=(0, 14))
        ttk.Button(controls_panel, text="Select All", style="Secondary.TButton", command=lambda: self._select_all(True)).pack(side="left", padx=(0, 6))
        ttk.Button(controls_panel, text="Clear All", style="Secondary.TButton", command=lambda: self._select_all(False)).pack(side="left")
        ttk.Button(controls_panel, text="Clear Log", style="Secondary.TButton", command=self._clear_log).pack(side="right")

        log_panel = ttk.Frame(outer, style="Panel.TFrame", padding=(12, 9, 10, 10))
        log_panel.pack(fill="both", expand=True)
        log_heading = ttk.Frame(log_panel, style="Panel.TFrame")
        log_heading.pack(fill="x", pady=(0, 6))
        ttk.Label(log_heading, text="Run activity", style="Section.TLabel").pack(side="left")
        ttk.Label(log_heading, text="Script output and prompts appear here", style="Muted.TLabel").pack(side="right")
        log_body = ttk.Frame(log_panel, style="Panel.TFrame")
        log_body.pack(fill="both", expand=True)
        self.log = tk.Text(log_body, wrap="word", state="disabled", font=("Segoe UI", 9), height=10, background="#fbfdff", foreground="#27384a", insertbackground="#1769aa", selectbackground="#cfe3f7", relief="flat", padx=10, pady=9, spacing1=1, spacing3=1)
        self.log.tag_configure("heading", foreground="#174d79", font=("Segoe UI", 9, "bold"))
        self.log.tag_configure("success", foreground="#147544")
        self.log.tag_configure("error", foreground="#b8322a")
        self.log.tag_configure("prompt", foreground="#725000", font=("Segoe UI", 9, "bold"))
        scrollbar = ttk.Scrollbar(log_body, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=scrollbar.set)
        self.log.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y", padx=(5, 0))

        input_panel = ttk.Frame(outer, style="Panel.TFrame", padding=(12, 9))
        input_panel.pack(fill="x", pady=(11, 0))
        ttk.Label(input_panel, text="Waiting script input", style="Section.TLabel").pack(side="left", padx=(0, 12))
        self.input_var = tk.StringVar()
        self.input_entry = ttk.Entry(input_panel, textvariable=self.input_var, show="")
        self.input_entry.pack(side="left", fill="x", expand=True, padx=(0, 9), ipady=5)
        self.input_entry.bind("<Return>", self._send_input)
        self.send_button = ttk.Button(input_panel, text="Send input", style="Input.TButton", command=self._send_input, state="disabled")
        self.send_button.pack(side="left")

        self.footer_var = tk.StringVar(value=f"Scripts folder: {self.script_dir}")
        ttk.Label(outer, textvariable=self.footer_var, style="Muted.TLabel", anchor="w").pack(fill="x", pady=(8, 0))
        self._update_selection_count()

    def _update_selection_count(self) -> None:
        total = len(self.selected_vars)
        selected = sum(1 for var in self.selected_vars.values() if var.get())
        self.selection_count.set(f"{selected} of {total} selected")
        self.run_button.configure(text=f"▶  Run Selected ({selected})")

    def _set_task_status(self, filename: str, status: str) -> None:
        self.status_vars[filename].set(status)
        style_name = status if status in {"Ready", "Queued", "Running", "Completed", "Failed", "Stopped"} else "Ready"
        self.status_labels[filename].configure(style=f"{style_name}.TLabel")

    def _append_log(self, text: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", text)
        self.log.see("end")
        self.log.configure(state="disabled")
        self._last_output_tail = (self._last_output_tail + text)[-500:]
        # getpass falls back to stdin when a script is run without a console.
        # Hide typed text in the GUI entry whenever the visible prompt is a password prompt.
        tail = self._last_output_tail.lower().rsplit("\n", 1)[-1]
        self.input_entry.configure(show="•" if "password" in tail else "")

    def _clear_log(self) -> None:
        if self.running:
            return
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        self._last_output_tail = ""

    def _select_all(self, value: bool) -> None:
        if self.running:
            return
        for var in self.selected_vars.values():
            var.set(value)

    def _set_running_ui(self, running: bool) -> None:
        self.running = running
        self.header_status.set("●  RUNNING" if running else "●  READY")
        self.run_button.configure(state="disabled" if running else "normal")
        self.rerun_button.configure(
            state="disabled" if running or not self.failed_scripts else "normal"
        )
        self.stop_button.configure(state="normal" if running else "disabled")
        self.send_button.configure(state="normal" if running else "disabled")
        if not running:
            self.input_entry.configure(show="")

    def _run_selected(self) -> None:
        if self.running:
            return
        selected = [
            (label, filename)
            for label, filename in AUTOMATIONS
            if self.selected_vars[filename].get()
        ]
        if not selected:
            messagebox.showinfo("No automations selected", "Select at least one automation to run.")
            return
        if any(filename in INPUT_SPECS for _, filename in selected):
            self._open_input_form(selected)
        else:
            self._start_run(selected, {}, {})

    def _open_input_form(self, scripts: list[tuple[str, str]]) -> None:
        window = tk.Toplevel(self)
        window.title("Automation Inputs")
        window.geometry("780x680")
        window.transient(self)
        window.grab_set()

        ttk.Label(
            window,
            text="Enter validation dates, mail subject date (DD-MMM-YYYY), optional Note, folder choices, and clipboard text before starting. Draft and send questions will appear when each mail script runs.",
            wraplength=740,
            font=("Segoe UI", 10, "bold"),
        ).pack(fill="x", padx=14, pady=(12, 8))

        canvas = tk.Canvas(window, highlightthickness=0)
        scrollbar = ttk.Scrollbar(window, orient="vertical", command=canvas.yview)
        content = ttk.Frame(canvas, padding=8)
        content.bind(
            "<Configure>",
            lambda _event: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas_window = canvas.create_window((0, 0), window=content, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind(
            "<Configure>",
            lambda event: canvas.itemconfigure(canvas_window, width=event.width),
        )
        canvas.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=4)
        scrollbar.pack(side="right", fill="y", padx=(0, 10), pady=4)

        fields: dict[tuple[str, str], object] = {}
        for label, filename in scripts:
            specs = INPUT_SPECS.get(filename, [])
            if not specs:
                continue
            group = ttk.LabelFrame(content, text=label)
            group.pack(fill="x", padx=4, pady=5)
            for key, caption, kind, default in specs:
                line = ttk.Frame(group)
                line.pack(fill="x", padx=8, pady=4)
                ttk.Label(line, text=caption, width=34, wraplength=260, anchor="w").pack(
                    side="left", padx=(0, 8)
                )
                if kind == "multichoice":
                    widget = tk.Listbox(
                        line,
                        selectmode=tk.MULTIPLE,
                        exportselection=False,
                        height=len(DRIVE_FOLDER_OPTIONS),
                        width=42,
                    )
                    for number, report_folder in DRIVE_FOLDER_OPTIONS:
                        widget.insert("end", f"{number}  —  {report_folder}")
                elif kind == "choice":
                    widget = ttk.Combobox(
                        line,
                        values=("Y", "N") if key != "validation_mode" else ("1", "2", "3"),
                        state="readonly",
                        width=15,
                    )
                    widget.set(default)
                elif kind == "multiline":
                    widget = tk.Text(line, height=4, width=45, wrap="word")
                    if default:
                        widget.insert("1.0", default)
                else:
                    widget = ttk.Entry(line, width=48)
                    if default:
                        widget.insert(0, default)
                widget.pack(side="left", fill="x", expand=True)
                fields[(filename, key)] = widget

        button_row = ttk.Frame(window)
        button_row.pack(fill="x", padx=12, pady=10)

        def read_field(filename: str, key: str) -> str:
            widget = fields[(filename, key)]
            if isinstance(widget, tk.Text):
                return widget.get("1.0", "end-1c").strip()
            if isinstance(widget, tk.Listbox):
                selected = widget.curselection()
                return ",".join(DRIVE_FOLDER_OPTIONS[index][0] for index in selected)
            return widget.get().strip()

        def begin() -> None:
            answers_by_script: dict[str, dict[str, str]] = {}
            clipboard_by_script: dict[str, str] = {}
            missing: list[str] = []

            def put(filename: str, prompt_part: str, value: str) -> None:
                answers_by_script.setdefault(filename, {})[prompt_part] = value

            for _, filename in scripts:
                if filename not in INPUT_SPECS:
                    continue
                values = {key: read_field(filename, key) for key, *_ in INPUT_SPECS[filename]}

                if filename == "Live_DIT_Count.py":
                    if not values["month"]:
                        missing.append("Live DIT Count: Month-Year")
                    if not values["dashboard_date"]:
                        missing.append("Live DIT Count: Dashboard Date")
                    put(filename, "Enter Month-Year", values["month"])
                    put(filename, "Enter Dashboard Date", values["dashboard_date"])
                elif filename == "Report_Count.py":
                    if not values["dit_summary"]:
                        missing.append("Report Count: DIT Summary text")
                    clipboard_by_script[filename] = values["dit_summary"]
                    put(filename, "Copy DIT Summary", "")
                elif filename == "Report+RDT+MT_Validation+MailAutomation.py":
                    mode = values["validation_mode"]
                    put(filename, "Choose Option", mode)
                    if mode == "3":
                        if not values["validation_dates"]:
                            missing.append("Validation & Mail: Custom date(s), since Option 3 is selected")
                        put(filename, "Enter Date(s)", values["validation_dates"])
                    subject_date = values["mail_subject_date"]
                    if not subject_date:
                        missing.append("Validation & Mail: mail subject date (DD-MMM-YYYY)")
                    else:
                        try:
                            datetime.strptime(subject_date, "%d-%b-%Y")
                        except ValueError:
                            missing.append("Validation & Mail: subject date must be DD-MMM-YYYY, for example 07-Oct-2026")
                    put(filename, "Subject date", subject_date)
                    put(filename, "Note to include in every email", values["mail_note"])
                    put(filename, "Press Enter to Exit", "")
                elif filename == "1Drive_Report_Upload.py":
                    if not values["drive_folders"]:
                        missing.append("OneDrive Upload: choose at least one report folder")
                    if not values["drive_date"]:
                        missing.append("OneDrive Upload: report date")
                    put(filename, "Enter Numbers", values["drive_folders"])
                    put(filename, "Enter Report Date", values["drive_date"])
                    put(filename, "Press Enter To Exit", "")
                elif filename == "Datasheet.py":
                    if not values["datasheet_date"]:
                        missing.append("Datasheet: Dashboard Date")
                    put(filename, "Enter Dashboard Date", values["datasheet_date"])
                elif filename == "Cleanup.py":
                    if not values["cleanup_date"]:
                        missing.append("Cleanup: report date")
                    put(filename, "Enter Date", values["cleanup_date"])
                elif filename == "TS.py":
                    if not values["teams_message"]:
                        missing.append("Teams Automation: Consolidated RDT message")
                    clipboard_by_script[filename] = values["teams_message"]
                    put(filename, "Copy Consolidated RDT message", "")

            if missing:
                messagebox.showerror("Inputs needed", "Please complete:\n\n" + "\n".join(missing), parent=window)
                return
            window.grab_release()
            window.destroy()
            self._start_run(scripts, answers_by_script, clipboard_by_script)

        ttk.Button(button_row, text="Cancel", command=window.destroy).pack(side="right", padx=(6, 0))
        ttk.Button(button_row, text="Start Automations", command=begin).pack(side="right")

    def _rerun_failed(self) -> None:
        if self.running or not self.failed_scripts:
            return
        failed = self.failed_scripts.copy()
        if any(filename in INPUT_SPECS for _, filename in failed):
            self._open_input_form(failed)
        else:
            self._start_run(failed, {}, {})

    def _start_run(
        self,
        scripts: list[tuple[str, str]],
        answers_by_script: dict[str, dict[str, str]],
        clipboard_by_script: dict[str, str],
    ) -> None:
        missing = [name for _, name in scripts if not (self.script_dir / name).is_file()]
        if missing:
            messagebox.showerror(
                "Automation files missing",
                "These scripts must be beside Automation_Hub.py:\n\n" + "\n".join(missing),
            )
            return
        if not (self.script_dir / "Master_Automation.py").exists():
            messagebox.showwarning(
                "Master script not found",
                "Place Automation_Hub.py in the same folder as Master_Automation.py and the automation scripts.",
            )
        self.failed_scripts = []
        self.rerun_button.configure(state="disabled")
        self._set_running_ui(True)
        for _, filename in scripts:
            self.status_vars[filename].set("Queued")
        threading.Thread(
            target=self._worker,
            args=(scripts, answers_by_script, clipboard_by_script),
            daemon=True,
        ).start()

    def _worker(
        self,
        scripts: list[tuple[str, str]],
        answers_by_script: dict[str, dict[str, str]],
        clipboard_by_script: dict[str, str],
    ) -> None:
        completed = 0
        failures: list[tuple[str, str]] = []
        self.events.put(("log", "\n" + "=" * 68 + "\nRUN STARTED\n" + "=" * 68 + "\n"))
        for label, filename in scripts:
            if self.closing:
                break
            script_path = self.script_dir / filename
            self.events.put(("status", filename, "Running"))
            self.events.put(("log", f"\n{'=' * 68}\nSTARTING: {label} ({filename})\n{'=' * 68}\n"))
            try:
                env = os.environ.copy()
                env.setdefault("PYTHONIOENCODING", "utf-8")
                env["AUTOMATION_HUB_ANSWERS"] = json.dumps(
                    answers_by_script.get(filename, {}), ensure_ascii=False
                )
                if filename in clipboard_by_script:
                    env["AUTOMATION_HUB_CLIPBOARD"] = clipboard_by_script[filename]
                python_executable = sys.executable
                if os.name == "nt" and Path(python_executable).name.lower() == "pythonw.exe":
                    console_python = Path(python_executable).with_name("python.exe")
                    if console_python.is_file():
                        python_executable = str(console_python)
                process = subprocess.Popen(
                    [python_executable, "-u", "-c", RUNNER, str(script_path)],
                    cwd=str(self.script_dir),
                    stdin=subprocess.PIPE,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    bufsize=0,
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                    env=env,
                )
                with self.process_lock:
                    self.process = process

                assert process.stdout is not None
                while True:
                    char = process.stdout.read(1)
                    if char == "":
                        break
                    self.events.put(("log", char))
                return_code = process.wait()
            except Exception as exc:
                return_code = -1
                self.events.put(("log", f"\nLauncher error: {exc}\n"))
            finally:
                with self.process_lock:
                    self.process = None

            if return_code == 0:
                completed += 1
                self.events.put(("status", filename, "Completed"))
                self.events.put(("log", f"\nCOMPLETED: {label}\n"))
            else:
                failures.append((label, filename))
                self.events.put(("status", filename, "Failed" if return_code != -15 else "Stopped"))
                self.events.put(("log", f"\nFAILED: {label} (exit code {return_code})\n"))
                if return_code == -15:
                    break

        self.events.put(("finished", completed, failures))

    def _send_input(self, _event=None) -> str:
        value = self.input_var.get()
        with self.process_lock:
            process = self.process
            try:
                if process is None or process.stdin is None or process.poll() is not None:
                    self._append_log("\nNo active script is waiting for input.\n")
                    return "break"
                process.stdin.write(value + "\n")
                process.stdin.flush()
            except (BrokenPipeError, OSError) as exc:
                self._append_log(f"\nCould not send input: {exc}\n")
                return "break"
        self.input_var.set("")
        return "break"

    def _stop_current(self) -> None:
        with self.process_lock:
            process = self.process
            if process and process.poll() is None:
                process.terminate()
                self._append_log("\nStop requested for current automation.\n")

    def _drain_events(self) -> None:
        try:
            while True:
                event = self.events.get_nowait()
                kind = event[0]
                if kind == "log":
                    self._append_log(event[1])
                elif kind == "status":
                    _, filename, status = event
                    self._set_task_status(filename, status)
                elif kind == "finished":
                    _, completed, failures = event
                    self.failed_scripts = failures
                    self._set_running_ui(False)
                    self.footer_var.set(
                        f"Finished: {completed} completed, {len(failures)} failed. Folder: {self.script_dir}"
                    )
                    if failures:
                        self._append_log("\nFAILED AUTOMATIONS:\n")
                        for label, filename in failures:
                            self._append_log(f"  • {label} ({filename})\n")
                        failed_names = "\n".join(f"• {label}" for label, _ in failures)
                        if messagebox.askyesno(
                            "Rerun failed automations?",
                            "These automations failed:\n\n"
                            f"{failed_names}\n\n"
                            "Do you want to rerun all failed automations now?",
                            parent=self,
                        ):
                            self._rerun_failed()
                    else:
                        self._append_log("\nAll selected automations completed.\n")
        except queue.Empty:
            pass
        if not self.closing:
            self.after(60, self._drain_events)

    def _on_close(self) -> None:
        if self.running and not messagebox.askyesno(
            "Automation running", "Stop the current script and close the application?"
        ):
            return
        self.closing = True
        self._stop_current()
        self.destroy()


if __name__ == "__main__":
    app = AutomationHub()
    app.mainloop()

