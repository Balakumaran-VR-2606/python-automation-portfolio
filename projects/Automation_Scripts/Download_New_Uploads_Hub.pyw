"""Windowed interface for Download_New_Uploads.py.

Keep this launcher beside Download_New_Uploads.py and start it by double-clicking.
The downloader's browser workflow remains in the original script.
"""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk


RUNNER = r'''import builtins, os, runpy, sys
script_path = os.path.abspath(sys.argv[1])
destination = os.environ.pop("DOWNLOAD_HUB_NEW_CASES", "")
sys.argv[:] = [script_path]
if destination:
    sys.argv.extend(["--new-cases-folder", destination])
applicant_ids = os.environ.pop("DOWNLOAD_HUB_APPLICANT_IDS", "")
original_input = builtins.input
ids_supplied = False
def window_input(prompt=""):
    global ids_supplied
    if prompt:
        sys.stdout.write(prompt)
        sys.stdout.flush()
    if "applicant ids:" in prompt.lower():
        ids_supplied = True
        sys.stdout.write(" [IDs from window]\n")
        sys.stdout.flush()
        return applicant_ids
    if ids_supplied and not prompt.strip():
        ids_supplied = False
        sys.stdout.write(" [ID list complete]\n")
        sys.stdout.flush()
        return ""
    return original_input()
builtins.input = window_input
runpy.run_path(script_path, run_name="__main__")
'''


class DownloadHub(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Digiverifier | New Upload Downloader")
        self.geometry("960x840")
        self.minsize(820, 720)
        self.configure(background="#eef2f8")

        self.script_dir = Path(__file__).resolve().parent
        self.script_path = self.script_dir / "Download_New_Uploads.py"
        self.default_destination = Path.home() / "Downloads" / "New_Cases"
        self.events: queue.Queue[tuple] = queue.Queue()
        self.process: subprocess.Popen | None = None
        self.process_lock = threading.Lock()
        self.running = False
        self.closing = False
        self.progress_count = 0
        self.total_ids = 0

        self.username_var = tk.StringVar()
        self.password_var = tk.StringVar()
        self.destination_var = tk.StringVar(value=str(self.default_destination))
        self.status_var = tk.StringVar(value="Ready when you are")
        self.progress_var = tk.StringVar(value="0 / 0 cases")
        self.show_password_var = tk.BooleanVar(value=False)

        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(60, self._drain_events)

    def _build_ui(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TFrame", background="#eef2f8")
        style.configure("Card.TFrame", background="#ffffff")
        style.configure("TLabel", background="#eef2f8", foreground="#25324a", font=("Segoe UI", 10))
        style.configure("Card.TLabel", background="#ffffff", foreground="#25324a", font=("Segoe UI", 10))
        style.configure("Card.TCheckbutton", background="#ffffff", foreground="#25324a", font=("Segoe UI", 9))
        style.map("Card.TCheckbutton", background=[("active", "#ffffff")])
        style.configure("Title.TLabel", background="#172554", foreground="#ffffff", font=("Segoe UI", 22, "bold"))
        style.configure("Subtitle.TLabel", background="#172554", foreground="#c7d2fe", font=("Segoe UI", 10))
        style.configure("Section.TLabel", background="#ffffff", foreground="#172554", font=("Segoe UI", 12, "bold"))
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), foreground="#ffffff", background="#2563eb", padding=(16, 10))
        style.map("Accent.TButton", background=[("active", "#1d4ed8"), ("disabled", "#9ca3af")])
        style.configure("Success.TButton", font=("Segoe UI", 10, "bold"), foreground="#ffffff", background="#059669", padding=(16, 10))
        style.map("Success.TButton", background=[("active", "#047857"), ("disabled", "#9ca3af")])
        style.configure("Danger.TButton", font=("Segoe UI", 10, "bold"), foreground="#ffffff", background="#e11d48", padding=(16, 10))
        style.map("Danger.TButton", background=[("active", "#be123c"), ("disabled", "#9ca3af")])
        style.configure("Soft.TButton", font=("Segoe UI", 9, "bold"), padding=(10, 7))

        header = ttk.Frame(self, style="Card.TFrame")
        header.pack(fill="x", padx=18, pady=(18, 12))
        hero = tk.Frame(header, bg="#172554", padx=22, pady=18)
        hero.pack(fill="x")
        ttk.Label(hero, text="NEW UPLOADS  ·  CASE DOWNLOADER", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            hero,
            text="Sign in once, paste Applicant IDs, and let the downloader file each case for you.",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(5, 0))

        form = ttk.Frame(self, style="Card.TFrame", padding=18)
        form.pack(fill="x", padx=18, pady=6)
        ttk.Label(form, text="Portal sign-in", style="Section.TLabel").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 12))

        ttk.Label(form, text="Username / email", style="Card.TLabel").grid(row=1, column=0, sticky="w", padx=(0, 12), pady=5)
        self.username_entry = ttk.Entry(form, textvariable=self.username_var, width=45)
        self.username_entry.grid(row=1, column=1, sticky="ew", pady=5)

        ttk.Label(form, text="Password", style="Card.TLabel").grid(row=2, column=0, sticky="w", padx=(0, 12), pady=5)
        self.password_entry = ttk.Entry(form, textvariable=self.password_var, width=45, show="•")
        self.password_entry.grid(row=2, column=1, sticky="ew", pady=5)
        ttk.Checkbutton(
            form,
            text="Show",
            variable=self.show_password_var,
            command=self._toggle_password,
            style="Card.TCheckbutton",
        ).grid(row=2, column=2, sticky="w", padx=8)

        ttk.Label(form, text="Save extracted cases to", style="Card.TLabel").grid(row=3, column=0, sticky="w", padx=(0, 12), pady=(10, 5))
        ttk.Entry(form, textvariable=self.destination_var).grid(row=3, column=1, sticky="ew", pady=(10, 5))
        self.browse_button = ttk.Button(form, text="Browse…", style="Soft.TButton", command=self._browse_destination)
        self.browse_button.grid(row=3, column=2, padx=(8, 0), pady=(10, 5))
        form.columnconfigure(1, weight=1)

        id_card = ttk.Frame(self, style="Card.TFrame", padding=18)
        id_card.pack(fill="both", expand=True, padx=18, pady=6)
        top = ttk.Frame(id_card, style="Card.TFrame")
        top.pack(fill="x")
        ttk.Label(top, text="Applicant IDs", style="Section.TLabel").pack(side="left")
        ttk.Label(top, text="Paste one per row or comma-separated", style="Card.TLabel").pack(side="right")
        self.ids_text = tk.Text(
            id_card,
            height=8,
            wrap="word",
            font=("Consolas", 11),
            background="#f8fafc",
            foreground="#172554",
            insertbackground="#2563eb",
            relief="flat",
            padx=12,
            pady=10,
            undo=True,
        )
        self.ids_text.pack(fill="both", expand=True, pady=(10, 10))
        self.ids_text.insert("1.0", "Copy the Applicant ID cells from Excel and paste here…")
        self.ids_text.bind("<FocusIn>", self._clear_id_placeholder)

        destination_note = ttk.Frame(id_card, style="Card.TFrame")
        destination_note.pack(fill="x")
        ttk.Label(destination_note, text="📁  Destination:", style="Card.TLabel").pack(side="left")
        ttk.Label(destination_note, textvariable=self.destination_var, style="Card.TLabel").pack(side="left", padx=(5, 0))

        controls = ttk.Frame(self)
        controls.pack(fill="x", padx=18, pady=(8, 6))
        self.start_button = ttk.Button(controls, text="⬇  START DOWNLOAD", style="Success.TButton", command=self._start)
        self.start_button.pack(side="left")
        self.stop_button = ttk.Button(controls, text="■  Stop", style="Danger.TButton", command=self._stop, state="disabled")
        self.stop_button.pack(side="left", padx=8)
        self.clear_button = ttk.Button(controls, text="✕  Clear IDs", style="Soft.TButton", command=self._clear_ids)
        self.clear_button.pack(side="left")
        self.open_button = ttk.Button(controls, text="↗  Open New_Cases", style="Soft.TButton", command=self._open_destination)
        self.open_button.pack(side="right")

        status = ttk.Frame(self, style="Card.TFrame", padding=(14, 10))
        status.pack(fill="x", padx=18, pady=6)
        self.status_dot = tk.Label(status, text="●", fg="#f59e0b", bg="#ffffff", font=("Segoe UI", 12))
        self.status_dot.pack(side="left")
        ttk.Label(status, textvariable=self.status_var, style="Card.TLabel").pack(side="left", padx=(7, 0))
        ttk.Label(status, textvariable=self.progress_var, style="Card.TLabel", font=("Segoe UI", 9, "bold")).pack(side="right")

        prompt_frame = ttk.Frame(self, style="Card.TFrame", padding=(12, 8))
        prompt_frame.pack(fill="x", padx=18, pady=(2, 6))
        ttk.Label(
            prompt_frame,
            text="Browser prompt (OTP / confirmation)",
            style="Card.TLabel",
        ).pack(side="left", padx=(0, 10))
        self.runtime_input_var = tk.StringVar()
        self.runtime_input_entry = ttk.Entry(prompt_frame, textvariable=self.runtime_input_var, state="disabled")
        self.runtime_input_entry.pack(side="left", fill="x", expand=True)
        self.runtime_input_entry.bind("<Return>", self._send_runtime_input)
        self.send_input_button = ttk.Button(
            prompt_frame,
            text="↵  Send",
            style="Accent.TButton",
            command=self._send_runtime_input,
            state="disabled",
        )
        self.send_input_button.pack(side="left", padx=(8, 0))

        log_frame = ttk.LabelFrame(self, text="Activity log", padding=8)
        log_frame.pack(fill="both", expand=True, padx=18, pady=(4, 16))
        self.log = tk.Text(log_frame, height=9, wrap="word", state="disabled", font=("Consolas", 9), background="#0f172a", foreground="#dbeafe", insertbackground="#ffffff")
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=log_scroll.set)
        self.log.pack(side="left", fill="both", expand=True)
        log_scroll.pack(side="right", fill="y")

    def _toggle_password(self) -> None:
        self.password_entry.configure(show="" if self.show_password_var.get() else "•")

    def _clear_id_placeholder(self, _event=None) -> None:
        if self.ids_text.get("1.0", "end-1c").strip() == "Copy the Applicant ID cells from Excel and paste here…":
            self.ids_text.delete("1.0", "end")
            self.ids_text.unbind("<FocusIn>")

    def _clear_ids(self) -> None:
        self.ids_text.delete("1.0", "end")
        self.ids_text.focus_set()

    def _browse_destination(self) -> None:
        folder = filedialog.askdirectory(initialdir=self.destination_var.get() or str(Path.home() / "Downloads"))
        if folder:
            self.destination_var.set(folder)

    def _open_destination(self) -> None:
        path = Path(self.destination_var.get()).expanduser()
        path.mkdir(parents=True, exist_ok=True)
        os.startfile(path)

    def _start(self) -> None:
        if self.running:
            return
        if not self.script_path.is_file():
            messagebox.showerror(
                "Downloader script missing",
                f"Keep this windowed app beside the downloader script:\n\n{self.script_path}",
                parent=self,
            )
            return

        username = self.username_var.get().strip()
        password = self.password_var.get()
        raw_ids = self.ids_text.get("1.0", "end-1c").strip()
        if raw_ids == "Copy the Applicant ID cells from Excel and paste here…":
            raw_ids = ""
        destination = self.destination_var.get().strip()
        if not username or not password:
            messagebox.showerror("Sign-in details needed", "Enter your portal username and password first.", parent=self)
            return
        if not raw_ids:
            messagebox.showerror("Applicant IDs needed", "Paste at least one Applicant ID.", parent=self)
            self.ids_text.focus_set()
            return
        if not destination:
            messagebox.showerror("Destination needed", "Choose where extracted case folders should be saved.", parent=self)
            return

        self.total_ids = sum(1 for item in raw_ids.replace(";", ",").replace("\t", ",").replace("\n", ",").split(",") if item.strip())
        self.progress_count = 0
        self.progress_var.set(f"0 / {self.total_ids} cases")
        self.status_var.set("Starting browser automation…")
        self.status_dot.configure(fg="#2563eb")
        self._set_running(True)
        threading.Thread(
            target=self._worker,
            args=(username, password, raw_ids, destination),
            daemon=True,
        ).start()

    def _worker(self, username: str, password: str, applicant_ids: str, destination: str) -> None:
        self.events.put(("log", "\n" + "=" * 58 + "\nDIGIVERIFIER NEW UPLOAD DOWNLOAD STARTED\n" + "=" * 58 + "\n"))
        try:
            env = os.environ.copy()
            env["PORTAL_USERNAME"] = username
            env["PORTAL_PASSWORD"] = password
            env["DOWNLOAD_HUB_APPLICANT_IDS"] = applicant_ids
            env["DOWNLOAD_HUB_NEW_CASES"] = destination
            env.setdefault("PYTHONIOENCODING", "utf-8")
            python_executable = sys.executable
            if os.name == "nt" and Path(python_executable).name.lower() == "pythonw.exe":
                console_python = Path(python_executable).with_name("python.exe")
                if console_python.is_file():
                    python_executable = str(console_python)
            process = subprocess.Popen(
                [python_executable, "-u", "-c", RUNNER, str(self.script_path)],
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
            carry = ""
            while True:
                char = process.stdout.read(1)
                if char == "":
                    break
                self.events.put(("log", char))
                carry = (carry + char)[-400:]
                if char == "\n":
                    self._track_progress(carry)
            return_code = process.wait()
            self.events.put(("finished", return_code))
        except Exception as exc:
            self.events.put(("log", f"\nLauncher error: {exc}\n"))
            self.events.put(("finished", -1))
        finally:
            with self.process_lock:
                self.process = None

    def _track_progress(self, line: str) -> None:
        lower_line = line.lower()
        if "automatic chart click did not expose a recognizable search field" in lower_line:
            self.events.put((
                "status",
                "Waiting for New Upload in Chrome — click its legend, then press Send here.",
            ))
            self.events.put(("attention",))
        elif "inline new upload search field still not detected" in lower_line:
            self.events.put(("status", "Could not open the New Upload case table; check the Activity log."))
        elif "--- searching applicant id" in lower_line:
            self.progress_count = min(self.progress_count + 1, self.total_ids)
            self.events.put(("progress", self.progress_count, self.total_ids))

    def _set_running(self, running: bool) -> None:
        self.running = running
        self.start_button.configure(state="disabled" if running else "normal")
        self.stop_button.configure(state="normal" if running else "disabled")
        self.clear_button.configure(state="disabled" if running else "normal")
        self.browse_button.configure(state="disabled" if running else "normal")
        self.runtime_input_entry.configure(state="normal" if running else "disabled")
        self.send_input_button.configure(state="normal" if running else "disabled")

    def _send_runtime_input(self, _event=None) -> str:
        value = self.runtime_input_var.get()
        with self.process_lock:
            process = self.process
            if not process or process.poll() is not None or process.stdin is None:
                self.status_var.set("No active browser prompt")
                return "break"
            try:
                process.stdin.write(value + "\n")
                process.stdin.flush()
            except (BrokenPipeError, OSError) as exc:
                self._append_log(f"\nCould not send browser response: {exc}\n")
                return "break"
        self.runtime_input_var.set("")
        self.status_var.set("Response sent to downloader")
        self.runtime_input_entry.focus_set()
        return "break"

    def _stop(self) -> None:
        with self.process_lock:
            process = self.process
            if process and process.poll() is None:
                process.terminate()
                self.status_var.set("Stopping…")

    def _append_log(self, text: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", text)
        self.log.see("end")
        self.log.configure(state="disabled")

    def _drain_events(self) -> None:
        try:
            while True:
                event = self.events.get_nowait()
                if event[0] == "log":
                    self._append_log(event[1])
                elif event[0] == "progress":
                    _, count, total = event
                    self.progress_var.set(f"{count} / {total} cases")
                    self.status_var.set(f"Processing case {count} of {total}…")
                elif event[0] == "status":
                    self.status_var.set(event[1])
                elif event[0] == "attention":
                    self.status_dot.configure(fg="#d97706")
                elif event[0] == "finished":
                    return_code = event[1]
                    self._set_running(False)
                    self.password_var.set("")
                    if return_code == 0:
                        self.status_var.set("Run finished — check the log for case results")
                        self.status_dot.configure(fg="#059669")
                    else:
                        self.status_var.set("Run stopped or ended with an error")
                        self.status_dot.configure(fg="#e11d48")
                        messagebox.showerror(
                            "Download run ended",
                            "The downloader stopped with an error. Check the Activity log and Tracker Filling-Summary.log for details.",
                            parent=self,
                        )
        except queue.Empty:
            pass
        if not self.closing:
            self.after(60, self._drain_events)

    def _on_close(self) -> None:
        if self.running and not messagebox.askyesno("Download in progress", "Stop the browser automation and close this app?", parent=self):
            return
        self.closing = True
        self._stop()
        self.destroy()


if __name__ == "__main__":
    app = DownloadHub()
    app.mainloop()






