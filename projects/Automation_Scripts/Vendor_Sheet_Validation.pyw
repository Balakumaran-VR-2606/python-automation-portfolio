from __future__ import annotations

import os
import queue
import re
import shutil
import subprocess
import sys
import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk


SCRIPT_DIR = Path(__file__).resolve().parent
SCRIPT_PATH = SCRIPT_DIR / "Vendor_Sheet_Validation.py"
OUTPUT_PREFIX = "Validation_Output_"
DEFAULT_SPOCS = {
    "EDU": "daniel.durairaja@digiverifier.com, saravanan.shanmugam@digiverifier.com",
    "CRC": "jaiganesh.velayutham@digiverifier.com, rakesh.kumaravel@digiverifier.com",
    "CRIMINAL": "jaiganesh.velayutham@digiverifier.com, rakesh.kumaravel@digiverifier.com",
    "GDC": "rakesh.kumaravel@digiverifier.com",
}
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")


class VendorValidationApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Vendor Sheet Validation")
        self.geometry("940x800")
        self.minsize(800, 680)
        self.configure(bg="#eef3f8")
        self.events: queue.Queue[tuple] = queue.Queue()
        self.running = False
        self.worker: threading.Thread | None = None
        self.vars: dict[str, tk.StringVar] = {}

        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except tk.TclError:
            style.theme_use("clam")
        style.configure("Title.TLabel", font=("Segoe UI", 19, "bold"), foreground="#12355b")
        style.configure("Hint.TLabel", font=("Segoe UI", 9), foreground="#40566f")
        style.configure("Run.TButton", font=("Segoe UI", 10, "bold"), padding=(16, 8))
        style.configure("TLabelframe.Label", font=("Segoe UI", 10, "bold"), foreground="#12355b")

        outer = ttk.Frame(self, padding=16)
        outer.pack(fill="both", expand=True)
        ttk.Label(outer, text="VENDOR SHEET VALIDATION", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            outer,
            text="Choose one vendor, any two vendors, or all three; set SPOC filters for the selected checks.",
            style="Hint.TLabel",
            wraplength=880,
        ).pack(anchor="w", pady=(4, 12))

        basic = ttk.LabelFrame(outer, text="Validation details", padding=10)
        basic.pack(fill="x", pady=(0, 10))
        self.vars["month"] = tk.StringVar(value=datetime.now().strftime("%b-%Y"))
        self.vars["vendor_mode"] = tk.StringVar(value="Single vendor")
        self.vars["vendor1"] = tk.StringVar(value="DHI")
        self.vars["vendor2"] = tk.StringVar(value="IPANGS")
        self.only_krya_mode = False
        ttk.Label(basic, text="Month (MMM-YYYY)").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=5)
        ttk.Entry(basic, textvariable=self.vars["month"], width=22).grid(row=0, column=1, sticky="w", pady=5)
        ttk.Label(basic, text="Run vendors").grid(row=0, column=2, sticky="w", padx=(26, 10), pady=5)
        mode_box = ttk.Combobox(basic, textvariable=self.vars["vendor_mode"], values=("Single vendor", "Choose two vendors", "All three vendors"), state="readonly", width=21)
        mode_box.grid(row=0, column=3, sticky="w", pady=5)
        ttk.Label(basic, text="Vendor 1").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=5)
        self.vendor1_box = ttk.Combobox(basic, textvariable=self.vars["vendor1"], values=("DHI", "IPANGS", "KRYA"), state="readonly", width=20)
        self.vendor1_box.grid(row=1, column=1, sticky="w", pady=5)
        ttk.Label(basic, text="Vendor 2").grid(row=1, column=2, sticky="w", padx=(26, 10), pady=5)
        self.vendor2_box = ttk.Combobox(basic, textvariable=self.vars["vendor2"], values=("DHI", "IPANGS", "KRYA"), state="readonly", width=21)
        self.vendor2_box.grid(row=1, column=3, sticky="w", pady=5)
        mode_box.bind("<<ComboboxSelected>>", self.update_vendor_controls)
        self.vendor1_box.bind("<<ComboboxSelected>>", self.update_vendor_controls)
        self.vendor2_box.bind("<<ComboboxSelected>>", self.update_vendor_controls)
        self.refresh_mail_index = tk.BooleanVar(value=False)
        ttk.Checkbutton(basic, text="Refresh selected vendor's Outlook mail cache", variable=self.refresh_mail_index).grid(row=2, column=0, columnspan=3, sticky="w", pady=(4, 0))

        spoc_frame = ttk.LabelFrame(outer, text="SPOC filters used to select source IDs", padding=10)
        spoc_frame.pack(fill="x", pady=(0, 10))
        labels = {
            "EDU": "Education",
            "CRC": "CRC / E-Court",
            "CRIMINAL": "Criminal",
            "GDC": "Global / GDC",
        }
        self.spoc_label_widgets = {}
        self.spoc_entry_widgets = {}
        for row, check in enumerate(("EDU", "CRC", "CRIMINAL", "GDC")):
            self.vars[check] = tk.StringVar(value=DEFAULT_SPOCS[check])
            label = ttk.Label(spoc_frame, text=labels[check], width=22)
            label.grid(row=row, column=0, sticky="w", padx=(0, 10), pady=4)
            entry = ttk.Entry(spoc_frame, textvariable=self.vars[check])
            entry.grid(row=row, column=1, sticky="ew", pady=4)
            self.spoc_label_widgets[check] = label
            self.spoc_entry_widgets[check] = entry
        self.krya_name_vars = {}
        krya_names_frame = ttk.LabelFrame(spoc_frame, text="Krya SPOC names (when running Krya with another vendor)", padding=6)
        self.krya_names_frame = krya_names_frame
        for row, check in enumerate(("EDU", "CRC", "CRIMINAL", "GDC")):
            self.krya_name_vars[check] = tk.StringVar(value="")
            name_label = ttk.Label(krya_names_frame, text=labels[check], width=22)
            name_label.grid(row=row, column=0, sticky="w", padx=(0, 10), pady=3)
            ttk.Entry(krya_names_frame, textvariable=self.krya_name_vars[check]).grid(row=row, column=1, sticky="ew", pady=3)
        krya_names_frame.columnconfigure(1, weight=1)
        krya_names_frame.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(5, 0))
        spoc_frame.columnconfigure(1, weight=1)
        self.spoc_hint = ttk.Label(
            spoc_frame,
            text="Rows are filtered using these SPOC emails. Krya report mails match any listed contact in From, To, or Cc.",
            style="Hint.TLabel",
            wraplength=860,
        )
        self.spoc_hint.grid(row=5, column=0, columnspan=2, sticky="w", pady=(7, 0))
        self.update_vendor_controls()

        controls = ttk.Frame(outer)
        controls.pack(fill="x", pady=(10, 0))
        self.status = tk.StringVar(value="Ready")
        ttk.Label(controls, textvariable=self.status, style="Hint.TLabel").pack(side="left")
        self.open_button = ttk.Button(controls, text="Open Validation Outputs", command=self.open_output, state="disabled")
        self.open_button.pack(side="right", padx=(8, 0))
        ttk.Button(controls, text="Clear Log", command=self.clear_log).pack(side="right", padx=(8, 0))
        self.start_button = tk.Button(controls, text="Start Validation", font=("Segoe UI", 10, "bold"), bg="#1769aa", fg="white", activebackground="#12558a", activeforeground="white", relief="flat", padx=18, pady=8, cursor="hand2", command=self.start_validation)
        self.start_button.pack(side="right")
        log_frame = ttk.LabelFrame(outer, text="Validation log", padding=8)
        log_frame.pack(fill="both", expand=True)
        self.log = tk.Text(log_frame, wrap="word", font=("Segoe UI", 10), bg="#ffffff", fg="#243447", insertbackground="#12355b", selectbackground="#cfe4ff", relief="flat", padx=12, pady=10, spacing1=2, spacing3=2)
        self.log.tag_configure("success", foreground="#16804a")
        self.log.tag_configure("error", foreground="#c0392b")
        self.log.tag_configure("heading", foreground="#12355b", font=("Segoe UI", 10, "bold"))
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=scrollbar.set)
        self.log.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")


        self.after(100, self.poll_events)

    def clear_log(self) -> None:
        if not self.running:
            self.log.delete("1.0", "end")

    def update_vendor_controls(self, _event=None) -> None:
        mode = self.vars["vendor_mode"].get()
        self.vendor1_box.configure(state="readonly" if mode != "All three vendors" else "disabled")
        self.vendor2_box.configure(state="readonly" if mode == "Choose two vendors" else "disabled")
        if mode == "All three vendors":
            vendors = {"DHI", "IPANGS", "KRYA"}
        elif mode == "Choose two vendors":
            vendors = {self.vars["vendor1"].get(), self.vars["vendor2"].get()}
        else:
            vendors = {self.vars["vendor1"].get()}
        only_krya = vendors == {"KRYA"}
        if only_krya != self.only_krya_mode:
            for check in ("EDU", "CRC", "CRIMINAL", "GDC"):
                current_value = self.vars[check].get().strip()
                if only_krya and "@" in current_value:
                    self.vars[check].set(self.krya_name_vars[check].get().strip())
                elif not only_krya and current_value and "@" not in current_value:
                    self.krya_name_vars[check].set(current_value)
                    self.vars[check].set(DEFAULT_SPOCS[check])
                elif not only_krya and not current_value:
                    self.vars[check].set(DEFAULT_SPOCS[check])
            self.only_krya_mode = only_krya
        if "KRYA" in vendors and not only_krya:
            self.krya_names_frame.grid()
        else:
            self.krya_names_frame.grid_remove()
        if only_krya:
            self.spoc_hint.configure(
                text="For Krya, enter SPOC name(s) in Education, CRC / E-Court, Criminal, and Global / GDC. Each check matches SPOC Name and its component in the datasheet."
            )
        elif "KRYA" in vendors:
            self.spoc_hint.configure(
                text="Top fields use emails for DHI/IPANGS. The separate Krya section accepts SPOC names. Krya mail search matches listed contacts in From, To, or Cc."
            )
        else:
            self.spoc_hint.configure(
                text="Rows are filtered using these SPOC emails. Use commas for multiple emails."
            )

    def append_log(self, text: str) -> None:
        lowered = text.lower()
        tag = "error" if ("traceback" in lowered or "failed" in lowered or "error:" in lowered) else "success" if ("completed successfully" in lowered or "validation passed" in lowered) else "heading" if ("starting vendor sheet validation" in lowered or "====" in text) else None
        if tag:
            self.log.insert("end", text, tag)
        else:
            self.log.insert("end", text)
        self.log.see("end")

    def validate_inputs(self) -> dict[str, str] | None:
        month = self.vars["month"].get().strip()
        try:
            datetime.strptime(month, "%b-%Y")
        except ValueError:
            messagebox.showerror("Invalid month", "Enter month as MMM-YYYY, for example Oct-2026.", parent=self)
            return None
        mode = self.vars["vendor_mode"].get()
        if mode == "All three vendors":
            vendors = ["DHI", "IPANGS", "KRYA"]
        elif mode == "Choose two vendors":
            vendors = [self.vars["vendor1"].get().strip().upper(), self.vars["vendor2"].get().strip().upper()]
            if vendors[0] == vendors[1]:
                messagebox.showerror("Choose two vendors", "Select two different vendors.", parent=self)
                return None
        else:
            vendors = [self.vars["vendor1"].get().strip().upper()]

        result = {"VENDOR_VALIDATION_MONTH": month, "VENDOR_VALIDATION_VENDORS": ",".join(vendors)}
        if self.refresh_mail_index.get():
            result["VENDOR_VALIDATION_REFRESH_MAIL_INDEX"] = "1"
        only_krya = vendors == ["KRYA"]
        if "KRYA" in vendors and not only_krya:
            if not any(self.krya_name_vars[check].get().strip() for check in ("EDU", "CRC", "CRIMINAL", "GDC")):
                messagebox.showerror("Krya SPOC names needed", "Enter one or more Krya SPOC names in the Krya section.", parent=self)
                return None
            for check in ("EDU", "CRC", "CRIMINAL", "GDC"):
                names = [item.strip() for item in re.split(r"[,;\n]+", self.krya_name_vars[check].get()) if item.strip()]
                result[f"VENDOR_VALIDATION_KRYA_SPOC_{check}"] = ",".join(names)
        if only_krya and not any(self.vars[check].get().strip() for check in ("EDU", "CRC", "CRIMINAL", "GDC")):
            messagebox.showerror("Krya SPOC names needed", "Enter one or more names in the Education, CRC / E-Court, Criminal, or Global / GDC field.", parent=self)
            return None
        for check in ("EDU", "CRC", "CRIMINAL", "GDC"):
            raw = self.vars[check].get().strip()
            if only_krya:
                result[f"VENDOR_VALIDATION_SPOC_{check}"] = raw
                continue
            parts = [p.strip() for p in re.split(r"[,;\n]+", raw) if p.strip()]
            if not parts:
                messagebox.showerror("SPOC email needed", f"Enter at least one SPOC email for {check}.", parent=self)
                return None
            invalid = [p for p in parts if not EMAIL_RE.fullmatch(p)]
            if invalid:
                messagebox.showerror("Invalid email", f"Check {check} email(s):\n" + "\n".join(invalid), parent=self)
                return None
            result[f"VENDOR_VALIDATION_SPOC_{check}"] = ",".join(p.lower() for p in parts)
        return result

    def python_executable(self) -> str:
        current = Path(sys.executable)
        if current.name.lower() == "pythonw.exe":
            console_python = current.with_name("python.exe")
            if console_python.exists():
                return str(console_python)
        found = shutil.which("python.exe") or shutil.which("python")
        return found or sys.executable

    def start_validation(self) -> None:
        if self.running:
            return
        if not SCRIPT_PATH.exists():
            messagebox.showerror("Script not found", f"Keep Vendor_Sheet_Validation.py beside this windowed app.\n\n{SCRIPT_PATH}", parent=self)
            return
        env_values = self.validate_inputs()
        if env_values is None:
            return

        child_env = os.environ.copy()
        child_env.update(env_values)
        child_env["PYTHONUTF8"] = "1"
        self.open_button.configure(state="disabled")
        self.start_button.configure(state="disabled")
        self.running = True
        self.status.set("Running validation…")
        self.append_log("\n" + "=" * 72 + "\nStarting Vendor Sheet Validation\n" + "=" * 72 + "\n")
        self.worker = threading.Thread(target=self.run_child, args=(child_env,), daemon=True)
        self.worker.start()

    def run_child(self, child_env: dict[str, str]) -> None:
        flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        try:
            vendors = [item.strip().upper() for item in child_env.pop("VENDOR_VALIDATION_VENDORS").split(",") if item.strip()]
            failed_vendors = []
            for index, vendor in enumerate(vendors):
                vendor_env = child_env.copy()
                vendor_env["VENDOR_VALIDATION_VENDOR"] = vendor
                month = vendor_env["VENDOR_VALIDATION_MONTH"]
                vendor_env["VENDOR_VALIDATION_OUTPUT"] = str(SCRIPT_DIR / f"{OUTPUT_PREFIX}{vendor}_{month}.xlsx")
                self.events.put(("line", f"\n--- Starting {vendor} validation ---\n"))
                process = subprocess.Popen(
                    [self.python_executable(), "-u", str(SCRIPT_PATH)],
                    cwd=str(SCRIPT_DIR), env=vendor_env,
                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    text=True, encoding="utf-8", errors="replace", bufsize=1,
                    creationflags=flags,
                )
                assert process.stdout is not None
                for line in process.stdout:
                    self.events.put(("line", line))
                code = process.wait()
                if code != 0:
                    failed_vendors.append(vendor)
            if failed_vendors:
                self.events.put(("line", "\nFailed vendors: " + ", ".join(failed_vendors) + "\n"))
                self.events.put(("done", 1))
            else:
                self.events.put(("done", 0))
        except Exception as exc:
            self.events.put(("error", str(exc)))

    def poll_events(self) -> None:
        try:
            while True:
                kind, value = self.events.get_nowait()
                if kind == "line":
                    self.append_log(value)
                elif kind == "done":
                    self.running = False
                    self.start_button.configure(state="normal")
                    if value == 0:
                        self.status.set("Completed successfully")
                        self.append_log("\nCompleted successfully.\n")
                        if list(SCRIPT_DIR.glob(f"{OUTPUT_PREFIX}*.xlsx")):
                            self.open_button.configure(state="normal")
                    else:
                        self.status.set(f"Failed (exit code {value})")
                        self.append_log(f"\nFailed with exit code {value}. See the log above.\n")
                        if list(SCRIPT_DIR.glob(f"{OUTPUT_PREFIX}*.xlsx")):
                            self.open_button.configure(state="normal")
                elif kind == "error":
                    self.running = False
                    self.start_button.configure(state="normal")
                    self.status.set("Could not start validation")
                    self.append_log(f"\nLauncher error: {value}\n")
        except queue.Empty:
            pass
        self.after(100, self.poll_events)

    def open_output(self) -> None:
        outputs = list(SCRIPT_DIR.glob(f"{OUTPUT_PREFIX}*.xlsx"))
        if outputs:
            try:
                os.startfile(str(SCRIPT_DIR))
            except Exception as exc:
                messagebox.showerror("Could not open outputs", str(exc), parent=self)
        else:
            messagebox.showinfo("Output missing", "Validation output files have not been created yet.", parent=self)


if __name__ == "__main__":
    VendorValidationApp().mainloop()
