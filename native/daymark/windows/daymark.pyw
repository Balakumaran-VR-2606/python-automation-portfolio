import json
import os
import subprocess
import sys
import threading
import time
import uuid
import winsound
from datetime import datetime, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk

try:
    import pystray
    from PIL import Image, ImageDraw
    TRAY_AVAILABLE = True
except ImportError:
    # The task manager remains usable without optional notification-area packages.
    pystray = None
    Image = None
    ImageDraw = None
    TRAY_AVAILABLE = False

APP_NAME = "Daymark"
APP_DIR = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / APP_NAME
DATA_FILE = APP_DIR / "reminders.json"
BG = "#0d1522"
CARD = "#141f30"
FIELD = "#0f1928"
LINE = "#2a3950"
TEXT = "#f2f2ed"
MUTED = "#a1adbd"
MINT = "#9ce5c3"
BLUE = "#1a293d"
TASK_PREFIX = "Daymark-"
APP = None


def read_tasks():
    try:
        with DATA_FILE.open("r", encoding="utf-8") as stream:
            data = json.load(stream)
        return data if isinstance(data, list) else []
    except (OSError, json.JSONDecodeError):
        return []


def write_tasks(tasks):
    APP_DIR.mkdir(parents=True, exist_ok=True)
    temp = DATA_FILE.with_suffix(".tmp")
    with temp.open("w", encoding="utf-8") as stream:
        json.dump(tasks, stream, ensure_ascii=False, indent=2)
    temp.replace(DATA_FILE)


def task_name(task_id):
    return TASK_PREFIX + task_id


def scheduled_action():
    if getattr(sys, "frozen", False):
        return [sys.executable]
    pythonw = Path(sys.executable).with_name("pythonw.exe")
    interpreter = str(pythonw if pythonw.exists() else sys.executable)
    return [interpreter, str(Path(__file__).resolve())]


def schedule_task(task):
    if task.get("completed"):
        cancel_task(task["id"])
        return
    when = max(int(task.get("next_at", 0)), int(time.time() * 1000) + 60_000)
    due = datetime.fromtimestamp(when / 1000)
    target = subprocess.list2cmdline(scheduled_action() + ["--ring", task["id"]])
    command = [
        "schtasks.exe", "/Create", "/F", "/SC", "ONCE",
        # schtasks parses /SD using the current Windows locale's date order.
        "/SD", due.strftime("%x"), "/ST", due.strftime("%H:%M"),
        "/TN", task_name(task["id"]), "/TR", target, "/IT",
    ]
    subprocess.run(command, capture_output=True, text=True, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0), check=False)


def cancel_task(task_id):
    subprocess.run(
        ["schtasks.exe", "/Delete", "/F", "/TN", task_name(task_id)],
        capture_output=True, text=True, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0), check=False,
    )


def schedule_all():
    for task in read_tasks():
        if not task.get("completed"):
            schedule_task(task)


def find_task(task_id):
    return next((task for task in read_tasks() if task.get("id") == task_id), None)


def advance_task(task_id):
    tasks = read_tasks()
    now = int(time.time() * 1000)
    for task in tasks:
        if task.get("id") != task_id or task.get("completed"):
            continue
        interval = max(60_000, int(task.get("repeat_ms", 3_600_000)))
        missed = max(1, ((now - int(task.get("next_at", now))) // interval) + 1)
        task["next_at"] = int(task.get("next_at", now)) + missed * interval
        write_tasks(tasks)
        schedule_task(task)
        return task
    return None


def start_scheduled_ring(task_id):
    task = find_task(task_id)
    if not task or task.get("completed"):
        return
    advance_task(task_id)
    duration = max(1, min(600, int(task.get("sound_seconds", 20))))
    stop = threading.Event()
    root = tk.Tk()
    root.title("Daymark reminder")
    root.configure(bg=BG)
    root.geometry("370x210")
    root.resizable(False, False)
    root.attributes("-topmost", True)
    body = tk.Frame(root, bg=CARD, padx=22, pady=18)
    body.pack(fill="both", expand=True, padx=12, pady=12)
    tk.Label(body, text="REMINDER", bg=CARD, fg=MINT, font=("Segoe UI", 9, "bold")).pack(anchor="w")
    tk.Label(body, text=task.get("title", "Task"), bg=CARD, fg=TEXT, font=("Segoe UI", 16, "bold"), wraplength=320).pack(anchor="w", pady=(8, 4))
    status = tk.Label(body, text=f"Beeping for up to {duration} seconds", bg=CARD, fg=MUTED, font=("Segoe UI", 10))
    status.pack(anchor="w")

    def complete_now():
        tasks = read_tasks()
        for item in tasks:
            if item.get("id") == task_id:
                item["completed"] = True
        write_tasks(tasks)
        cancel_task(task_id)
        stop.set()
        root.destroy()

    tk.Button(body, text="Mark done", command=complete_now, bg=MINT, fg="#10251e", relief="flat", font=("Segoe UI", 10, "bold"), padx=12, pady=7).pack(anchor="e", pady=(15, 0))
    root.protocol("WM_DELETE_WINDOW", root.destroy)

    def beep_loop():
        deadline = time.monotonic() + duration
        while time.monotonic() < deadline and not stop.is_set():
            current = find_task(task_id)
            if not current or current.get("completed"):
                break
            try:
                winsound.Beep(740, 260)
            except RuntimeError:
                winsound.MessageBeep(winsound.MB_ICONASTERISK)
            stop.wait(0.65)
        if not stop.is_set():
            stop.set()
            try:
                root.after(0, root.destroy)
            except tk.TclError:
                pass

    threading.Thread(target=beep_loop, daemon=True).start()
    root.mainloop()


class DaymarkApp:
    def __init__(self, background=False):
        self.root = tk.Tk()
        self.root.title("Daymark — Task Reminders")
        self.root.geometry("780x750")
        self.root.minsize(620, 600)
        self.root.configure(bg=BG)
        self.root.protocol("WM_DELETE_WINDOW", self.hide_to_tray)
        self.icon = None
        self.background = background
        self.title_var = tk.StringVar()
        self.repeat_var = tk.StringVar(value="2")
        self.unit_var = tk.StringVar(value="hours")
        self.sound_var = tk.StringVar(value="20")
        self.start_var = tk.StringVar(value=(datetime.now() + timedelta(minutes=2)).strftime("%Y-%m-%d %H:%M"))
        self.show_completed = False
        self._build()
        self._start_tray()
        schedule_all()
        self.root.after(15000, self._refresh_periodically)
        if background and self.icon is not None:
            self.root.withdraw()

    def _build(self):
        shell = tk.Frame(self.root, bg=BG, padx=22, pady=18)
        shell.pack(fill="both", expand=True)
        brand = tk.Frame(shell, bg=BG)
        brand.pack(fill="x")
        tk.Label(brand, text="d", bg=MINT, fg="#10251e", font=("Segoe UI", 16, "bold"), width=3, pady=4).pack(side="left")
        tk.Label(brand, text="  daymark", bg=BG, fg=TEXT, font=("Segoe UI", 17, "bold")).pack(side="left")
        tk.Label(brand, text="Saved on this laptop", bg=BG, fg=MUTED, font=("Segoe UI", 9)).pack(side="right")

        tk.Label(shell, text="MAKE ROOM FOR WHAT MATTERS", bg=BG, fg=MINT, font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(23, 8))
        tk.Label(shell, text="Keep the things you mean to do.", bg=BG, fg=TEXT, font=("Segoe UI", 25, "bold")).pack(anchor="w")
        tk.Label(shell, text="Reminders run in the background until you mark a task done.", bg=BG, fg=MUTED, font=("Segoe UI", 10)).pack(anchor="w", pady=(5, 14))

        form = tk.Frame(shell, bg=CARD, padx=17, pady=15, highlightbackground=LINE, highlightthickness=1)
        form.pack(fill="x")
        tk.Label(form, text="A NEW REMINDER", bg=CARD, fg=MINT, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(form, text="What needs doing?", bg=CARD, fg=TEXT, font=("Segoe UI", 15, "bold")).pack(anchor="w", pady=(4, 12))
        self._label(form, "Task name")
        self._entry(form, self.title_var, "e.g. Take evening medicine").pack(fill="x", pady=(4, 10))
        row = tk.Frame(form, bg=CARD)
        row.pack(fill="x")
        col1 = tk.Frame(row, bg=CARD); col1.pack(side="left", fill="x", expand=True, padx=(0, 8))
        col2 = tk.Frame(row, bg=CARD); col2.pack(side="left", fill="x", expand=True, padx=(8, 0))
        self._label(col1, "Start time (YYYY-MM-DD HH:MM)")
        self._entry(col1, self.start_var, "2026-11-12 20:00").pack(fill="x", pady=(4, 8))
        self._label(col2, "Repeat every")
        repeat = tk.Frame(col2, bg=CARD); repeat.pack(fill="x", pady=(4, 8))
        self._entry(repeat, self.repeat_var, "2", width=8).pack(side="left", fill="x", expand=True)
        ttk.Combobox(repeat, textvariable=self.unit_var, values=("minutes", "hours", "days"), state="readonly", width=10).pack(side="left", padx=(7, 0), ipady=7)
        self._label(form, "Sound duration (seconds, 1–600)")
        self._entry(form, self.sound_var, "20").pack(fill="x", pady=(4, 10))
        actions = tk.Frame(form, bg=CARD); actions.pack(fill="x")
        tk.Label(actions, text="Repeats until marked done", bg=CARD, fg=MUTED, font=("Segoe UI", 9)).pack(side="left", pady=9)
        tk.Button(actions, text="＋  Add reminder", command=self.add_task, bg=MINT, fg="#10251e", relief="flat", font=("Segoe UI", 10, "bold"), padx=15, pady=9).pack(side="right")

        permission = tk.Frame(shell, bg=CARD, padx=13, pady=10, highlightbackground=LINE, highlightthickness=1)
        permission.pack(fill="x", pady=(10, 15))
        tk.Label(permission, text="Windows Task Scheduler will ring at the saved time, even if this window is hidden. Keep the laptop awake and signed in.", bg=CARD, fg=MUTED, font=("Segoe UI", 9), wraplength=700, justify="left").pack(anchor="w")
        tk.Checkbutton(permission, text="Start Daymark when I sign in to Windows", command=self.toggle_startup,
                       bg=CARD, fg=TEXT, selectcolor=FIELD, activebackground=CARD, activeforeground=TEXT,
                       variable=self._startup_var(), font=("Segoe UI", 9)).pack(anchor="w", pady=(4, 0))

        head = tk.Frame(shell, bg=BG); head.pack(fill="x", pady=(1, 4))
        tk.Label(head, text="REMINDERS", bg=BG, fg=MINT, font=("Segoe UI", 9, "bold")).pack(side="left")
        self.filter_button = tk.Button(head, text="Completed", command=self.toggle_filter, bg=BLUE, fg=TEXT, relief="flat", font=("Segoe UI", 9), padx=10, pady=5)
        self.filter_button.pack(side="right")
        self.list_frame = tk.Frame(shell, bg=BG)
        self.list_frame.pack(fill="both", expand=True)
        self.refresh()

    def _startup_var(self):
        import tkinter
        self.startup_var = tkinter.BooleanVar(value=startup_enabled())
        return self.startup_var

    def _label(self, parent, value):
        tk.Label(parent, text=value, bg=CARD, fg=TEXT, font=("Segoe UI", 9, "bold")).pack(anchor="w")

    def _entry(self, parent, variable, hint, width=None):
        entry = tk.Entry(parent, textvariable=variable, bg=FIELD, fg=TEXT, insertbackground=TEXT,
                         relief="flat", highlightbackground=LINE, highlightcolor=MINT, highlightthickness=1,
                         font=("Segoe UI", 10), width=width)
        entry.insert(0, "") if False else None
        return entry

    def _start_tray(self):
        if not TRAY_AVAILABLE:
            return
        image = Image.new("RGB", (64, 64), (156, 229, 195))
        draw = ImageDraw.Draw(image); draw.rounded_rectangle((3, 3, 61, 61), radius=14, fill=(156, 229, 195))
        draw.text((18, 11), "d", fill=(16, 37, 30))
        menu = pystray.Menu(pystray.MenuItem("Open Daymark", self._tray_open, default=True), pystray.MenuItem("Exit", self._tray_exit))
        self.icon = pystray.Icon("Daymark", image, "Daymark Reminders", menu)
        self.icon.run_detached()

    def _tray_open(self, icon=None, item=None): self.root.after(0, self.show_window)
    def _tray_exit(self, icon=None, item=None): self.root.after(0, self.exit_app)
    def show_window(self): self.root.deiconify(); self.root.lift(); self.root.focus_force(); self.refresh()
    def hide_to_tray(self):
        if self.icon is not None:
            self.root.withdraw()
        else:
            self.root.destroy()
    def exit_app(self):
        if self.icon: self.icon.stop()
        self.root.destroy()

    def toggle_startup(self):
        enabled = bool(self.startup_var.get())
        set_startup(enabled)
        self.toast("Starts with Windows." if enabled else "Windows startup disabled.")

    def add_task(self):
        title = self.title_var.get().strip()
        try:
            start_at = datetime.strptime(self.start_var.get().strip(), "%Y-%m-%d %H:%M")
            repeat_value = int(self.repeat_var.get().strip())
            sound_seconds = int(self.sound_var.get().strip())
        except ValueError:
            self.toast("Check the date, repeat interval, and sound duration."); return
        if not title:
            self.toast("Enter a task name."); return
        if start_at <= datetime.now():
            self.toast("Choose a start time in the future."); return
        if repeat_value < 1 or repeat_value > 999 or sound_seconds < 1 or sound_seconds > 600:
            self.toast("Repeat must be 1–999; sound must be 1–600 seconds."); return
        unit = self.unit_var.get()
        factor = {"minutes": 60_000, "hours": 3_600_000, "days": 86_400_000}[unit]
        task = {"id": uuid.uuid4().hex, "title": title, "next_at": int(start_at.timestamp() * 1000),
                "repeat_ms": repeat_value * factor, "repeat_value": repeat_value, "repeat_unit": unit,
                "sound_seconds": sound_seconds, "completed": False}
        tasks = read_tasks(); tasks.insert(0, task); write_tasks(tasks); schedule_task(task)
        self.title_var.set(""); self.refresh()
        self.toast("Reminder saved on this laptop.")

    def toggle_filter(self):
        self.show_completed = not self.show_completed
        self.refresh()

    def refresh(self):
        if not hasattr(self, "list_frame"):
            return
        for child in self.list_frame.winfo_children(): child.destroy()
        tasks = read_tasks()
        visible = [t for t in tasks if bool(t.get("completed")) == self.show_completed]
        visible.sort(key=lambda t: int(t.get("next_at", 0)))
        self.filter_button.configure(text="In progress" if self.show_completed else "Completed")
        if not visible:
            tk.Label(self.list_frame, text="No completed reminders." if self.show_completed else "No reminders yet.", bg=BG, fg=MUTED, font=("Segoe UI", 10)).pack(anchor="w", pady=12)
        for task in visible:
            card = tk.Frame(self.list_frame, bg=CARD, padx=13, pady=11, highlightbackground=LINE, highlightthickness=1)
            card.pack(fill="x", pady=5)
            tk.Label(card, text=task.get("title", "Reminder"), bg=CARD, fg=TEXT, font=("Segoe UI", 11, "bold"), anchor="w").pack(fill="x")
            due = datetime.fromtimestamp(int(task.get("next_at", 0)) / 1000).strftime("%a, %d %b · %I:%M %p")
            detail = "Completed" if task.get("completed") else f"Next: {due}  ·  Every {task.get('repeat_value')} {task.get('repeat_unit')}  ·  {task.get('sound_seconds')}s sound"
            tk.Label(card, text=detail, bg=CARD, fg=MUTED, font=("Segoe UI", 9), anchor="w").pack(fill="x", pady=(3, 7))
            row = tk.Frame(card, bg=CARD); row.pack(fill="x")
            if not task.get("completed"):
                tk.Button(row, text="Mark done", command=lambda t=task: self.complete(t["id"]), bg=MINT, fg="#10251e", relief="flat", font=("Segoe UI", 9, "bold"), padx=10, pady=5).pack(side="right")
            tk.Button(row, text="Delete", command=lambda t=task: self.delete(t["id"]), bg=BLUE, fg=TEXT, relief="flat", font=("Segoe UI", 9), padx=10, pady=5).pack(side="right", padx=(0, 6))

    def _refresh_periodically(self):
        if self.root.winfo_exists():
            self.refresh()
            self.root.after(15000, self._refresh_periodically)

    def complete(self, task_id):
        tasks = read_tasks()
        for task in tasks:
            if task["id"] == task_id: task["completed"] = True
        write_tasks(tasks); cancel_task(task_id); self.refresh()

    def delete(self, task_id):
        write_tasks([task for task in read_tasks() if task.get("id") != task_id])
        cancel_task(task_id); self.refresh()

    def toast(self, message):
        messagebox.showinfo("Daymark", message, parent=self.root)

    def run(self): self.root.mainloop()


def startup_enabled():
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run") as key:
            winreg.QueryValueEx(key, "Daymark")
        return True
    except OSError:
        return False


def set_startup(enabled):
    import winreg
    key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
    if enabled:
        if getattr(sys, "frozen", False):
            command = subprocess.list2cmdline([sys.executable, "--background"])
        else:
            pythonw = Path(sys.executable).with_name("pythonw.exe")
            interpreter = str(pythonw if pythonw.exists() else sys.executable)
            command = subprocess.list2cmdline([interpreter, str(Path(__file__).resolve()), "--background"])
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
            winreg.SetValueEx(key, "Daymark", 0, winreg.REG_SZ, command)
    else:
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
                winreg.DeleteValue(key, "Daymark")
        except OSError:
            pass


def main():
    APP_DIR.mkdir(parents=True, exist_ok=True)
    if "--ring" in sys.argv:
        index = sys.argv.index("--ring")
        if index + 1 < len(sys.argv): start_scheduled_ring(sys.argv[index + 1])
        return
    background = "--background" in sys.argv
    global APP
    APP = DaymarkApp(background=background)
    APP.run()


if __name__ == "__main__":
    main()
