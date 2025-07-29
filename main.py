import sys
import os
import time
import json
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
import win32gui
import win32process
import psutil
import pystray
from PIL import Image
import shutil
import ctypes
from pathlib import Path

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

DATA_FILE = resource_path('data.json')
SETTINGS_FILE = resource_path('settings.json')
LAST_USED_FILE = resource_path('last_used.json')
ICON_PATH = resource_path('trayicon.ico')
TRACK_INTERVAL = 1

def get_active_app():
    try:
        hwnd = win32gui.GetForegroundWindow()
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        process = psutil.Process(pid)
        return process.name()
    except Exception:
        return None

def load_json(path):
    if os.path.exists(path):
        with open(path, 'r') as f:
            return json.load(f)
    return {}

def save_json(path, data):
    with open(path, 'w') as f:
        json.dump(data, f)

def is_startup_enabled(user=True):
    try:
        import winreg
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        root = winreg.HKEY_CURRENT_USER if user else winreg.HKEY_LOCAL_MACHINE
        with winreg.OpenKey(root, key_path, 0, winreg.KEY_READ) as key:
            try:
                val = winreg.QueryValueEx(key, "AllSeeingEye")
                return True
            except FileNotFoundError:
                return False
    except Exception:
        return False

def set_startup(enable, user=True):
    try:
        import winreg
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        root = winreg.HKEY_CURRENT_USER if user else winreg.HKEY_LOCAL_MACHINE
        with winreg.OpenKey(root, key_path, 0, winreg.KEY_WRITE) as key:
            if enable:
                exe_path = sys.executable
                winreg.SetValueEx(key, "AllSeeingEye", 0, winreg.REG_SZ, f'"{exe_path}"')
            else:
                try:
                    winreg.DeleteValue(key, "AllSeeingEye")
                except FileNotFoundError:
                    pass
    except Exception as e:
        messagebox.showerror("Startup Setting Error", f"Error setting startup: {e}")

class AppTracker:
    def __init__(self):
        self.running = True
        self.data = load_json(DATA_FILE)
        self.settings = load_json(SETTINGS_FILE)
        self.recent_apps = []
        self.current_app = None
        self.start_time = time.time()
        self.lock = threading.Lock()
        self.last_used = load_json(LAST_USED_FILE)
        self.whitelist = set(self.settings.get("whitelist", []))
        self.idle_threshold = 300  # 5 minutes

    def start(self):
        threading.Thread(target=self.track_loop, daemon=True).start()

    def track_loop(self):
        class LASTINPUTINFO(ctypes.Structure):
            _fields_ = [("cbSize", ctypes.c_uint), ("dwTime", ctypes.c_uint)]

        def get_idle_duration():
            lii = LASTINPUTINFO()
            lii.cbSize = ctypes.sizeof(LASTINPUTINFO)
            ctypes.windll.user32.GetLastInputInfo(ctypes.byref(lii))
            millis = ctypes.windll.kernel32.GetTickCount() - lii.dwTime
            return millis / 1000.0

        while self.running:
            idle_time = get_idle_duration()
            now = time.time()
            with self.lock:
                if idle_time >= self.idle_threshold:
                    if self.current_app != "Idle Time":
                        if self.current_app:
                            elapsed = now - self.start_time
                            self.data[self.current_app] = self.data.get(self.current_app, 0) + elapsed
                            self.last_used[self.current_app] = now
                            if self.current_app not in self.recent_apps:
                                self.recent_apps.insert(0, self.current_app)
                                self.recent_apps = self.recent_apps[:3]
                        self.current_app = "Idle Time"
                        self.start_time = now
                else:
                    app = get_active_app()
                    if app and (app in self.whitelist or len(self.whitelist) == 0):
                        if app != self.current_app:
                            if self.current_app:
                                elapsed = now - self.start_time
                                self.data[self.current_app] = self.data.get(self.current_app, 0) + elapsed
                                self.last_used[self.current_app] = now
                                if self.current_app not in self.recent_apps:
                                    self.recent_apps.insert(0, self.current_app)
                                    self.recent_apps = self.recent_apps[:3]
                            self.current_app = app
                            self.start_time = now
            save_json(DATA_FILE, self.data)
            save_json(LAST_USED_FILE, self.last_used)
            save_json(SETTINGS_FILE, self.settings)
            time.sleep(TRACK_INTERVAL)

    def stop(self):
        self.running = False
        now = time.time()
        with self.lock:
            if self.current_app:
                elapsed = now - self.start_time
                self.data[self.current_app] = self.data.get(self.current_app, 0) + elapsed
                self.last_used[self.current_app] = now
                if self.current_app not in self.recent_apps:
                    self.recent_apps.insert(0, self.current_app)
                    self.recent_apps = self.recent_apps[:3]
        save_json(DATA_FILE, self.data)
        save_json(LAST_USED_FILE, self.last_used)
        save_json(SETTINGS_FILE, self.settings)

class TrackerGUI:
    def __init__(self, tracker):
        self.tracker = tracker
        self.root = tk.Tk()
        self.root.title("AllSeeingEye")
        self.root.geometry("400x400")

        # Load icon safely
        if os.path.exists(ICON_PATH):
            try:
                self.root.iconbitmap(ICON_PATH)
            except Exception:
                pass

        # Recent apps display
        self.recent_labels = [tk.Label(self.root, text="") for _ in range(3)]
        for lbl in self.recent_labels:
            lbl.pack(pady=2)

        self.export_btn = tk.Button(self.root, text="Export Data", command=self.export_data)
        self.export_btn.pack(pady=10)

        # Whitelist UI
        self.whitelist_frame = tk.Frame(self.root)
        self.whitelist_frame.pack(pady=10, fill='both', expand=True)

        tk.Label(self.whitelist_frame, text="Whitelisted Apps (empty = all):").pack()

        self.whitelist_listbox = tk.Listbox(self.whitelist_frame, height=6)
        self.whitelist_listbox.pack(side='left', fill='both', expand=True)

        scrollbar = tk.Scrollbar(self.whitelist_frame)
        scrollbar.pack(side='right', fill='y')

        self.whitelist_listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.whitelist_listbox.yview)

        self.btn_frame = tk.Frame(self.root)
        self.btn_frame.pack()

        self.add_btn = tk.Button(self.btn_frame, text="Add EXE to whitelist", command=self.add_whitelist)
        self.add_btn.grid(row=0, column=0, padx=5, pady=5)

        self.remove_btn = tk.Button(self.btn_frame, text="Remove Selected", command=self.remove_whitelist)
        self.remove_btn.grid(row=0, column=1, padx=5, pady=5)

        # Dark mode toggle
        self.var_dark_mode = tk.IntVar(value=self.tracker.settings.get("dark_mode", 0))
        self.dark_mode_chk = tk.Checkbutton(self.root, text="Dark Mode", variable=self.var_dark_mode, command=self.toggle_dark_mode)
        self.dark_mode_chk.pack(pady=5)

        # Startup checkboxes
        self.var_start_minimized = tk.IntVar(value=self.tracker.settings.get("start_minimized", 0))
        self.start_min_chk = tk.Checkbutton(self.root, text="Start Minimized to Tray", variable=self.var_start_minimized, command=self.save_settings)
        self.start_min_chk.pack(pady=5)

        self.var_startup_user = tk.IntVar(value=1 if is_startup_enabled(True) else 0)
        self.chk_startup_user = tk.Checkbutton(self.root, text="Run at Startup (Current User)", variable=self.var_startup_user, command=self.toggle_startup_user)
        self.chk_startup_user.pack(pady=5)

        self.var_startup_all = tk.IntVar(value=1 if is_startup_enabled(False) else 0)
        self.chk_startup_all = tk.Checkbutton(self.root, text="Run at Startup (All Users - Requires Admin)", variable=self.var_startup_all, command=self.toggle_startup_all)
        self.chk_startup_all.pack(pady=5)

        self.tray_icon = None
        self.is_tray_active = False

        self.update_whitelist_listbox()
        self.apply_dark_mode(self.var_dark_mode.get())
        self.update_ui()

        self.root.protocol("WM_DELETE_WINDOW", self.minimize_to_tray)

        if self.var_start_minimized.get() == 1:
            self.minimize_to_tray()

    def update_ui(self):
        self.tracker.lock.acquire()
        recent = self.tracker.recent_apps
        data = self.tracker.data
        last_used = self.tracker.last_used
        self.tracker.lock.release()

        for i in range(3):
            if i < len(recent):
                app = recent[i]
                seconds = int(data.get(app, 0))
                time_str = time.strftime('%H:%M:%S', time.gmtime(seconds))
                last_used_ts = last_used.get(app)
                last_used_str = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(last_used_ts)) if last_used_ts else "Never"
                self.recent_labels[i].config(text=f"{app} - {time_str} (Last used: {last_used_str})")
            else:
                self.recent_labels[i].config(text="")
        self.root.after(1000, self.update_ui)

    def export_data(self):
        filepath = filedialog.asksaveasfilename(defaultextension=".txt",
                                                filetypes=[("Text Files", "*.txt")])
        if filepath:
            with open(filepath, 'w') as f:
                self.tracker.lock.acquire()
                for app, seconds in sorted(self.tracker.data.items(), key=lambda x: -x[1]):
                    time_str = time.strftime('%H:%M:%S', time.gmtime(int(seconds)))
                    last_used_ts = self.tracker.last_used.get(app)
                    last_used_str = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(last_used_ts)) if last_used_ts else "Never"
                    f.write(f"{app}: {time_str} (Last used: {last_used_str})\n")
                self.tracker.lock.release()

    def add_whitelist(self):
        file_path = filedialog.askopenfilename(title="Select EXE", filetypes=[("Executables", "*.exe")])
        if file_path:
            exe_name = os.path.basename(file_path)
            self.tracker.whitelist.add(exe_name)
            self.update_whitelist_listbox()
            self.save_settings()

    def remove_whitelist(self):
        selected = self.whitelist_listbox.curselection()
        if selected:
            exe_name = self.whitelist_listbox.get(selected[0])
            if exe_name in self.tracker.whitelist:
                self.tracker.whitelist.remove(exe_name)
                self.update_whitelist_listbox()
                self.save_settings()

    def update_whitelist_listbox(self):
        self.whitelist_listbox.delete(0, tk.END)
        for exe in sorted(self.tracker.whitelist):
            self.whitelist_listbox.insert(tk.END, exe)

    def save_settings(self):
        self.tracker.settings["dark_mode"] = self.var_dark_mode.get()
        self.tracker.settings["start_minimized"] = self.var_start_minimized.get()
        self.tracker.settings["whitelist"] = list(self.tracker.whitelist)
        save_json(SETTINGS_FILE, self.tracker.settings)

    def toggle_dark_mode(self):
        enabled = self.var_dark_mode.get()
        self.apply_dark_mode(enabled)
        self.save_settings()

    def apply_dark_mode(self, enabled):
        bg = "#222222" if enabled else "#f0f0f0"
        fg = "#dddddd" if enabled else "#000000"
        btn_bg = "#333333" if enabled else "#f0f0f0"
        btn_fg = fg
        select_color = "#555555" if enabled else "#f0f0f0"

        self.root.configure(bg=bg)
        for widget in self.root.winfo_children():
            cls = widget.__class__.__name__
            if cls == 'Frame':
                widget.configure(bg=bg)
                for child in widget.winfo_children():
                    try:
                        child.configure(bg=bg, fg=fg)
                        if isinstance(child, tk.Button):
                            child.configure(bg=btn_bg, fg=btn_fg, borderwidth=2, relief='raised', highlightthickness=1)
                        if isinstance(child, tk.Checkbutton):
                            child.configure(bg=bg, fg=fg, selectcolor=select_color)
                        if isinstance(child, tk.Listbox):
                            child.configure(bg=bg, fg=fg, borderwidth=2, highlightthickness=1)
                    except:
                        pass
            else:
                try:
                    widget.configure(bg=bg, fg=fg)
                    if isinstance(widget, tk.Button):
                        widget.configure(bg=btn_bg, fg=btn_fg, borderwidth=2, relief='raised', highlightthickness=1)
                    if isinstance(widget, tk.Checkbutton):
                        widget.configure(bg=bg, fg=fg, selectcolor=select_color)
                    if isinstance(widget, tk.Listbox):
                        widget.configure(bg=bg, fg=fg, borderwidth=2, highlightthickness=1)
                except:
                    pass

    def toggle_startup_user(self):
        enable = bool(self.var_startup_user.get())
        set_startup(enable, user=True)

    def toggle_startup_all(self):
        enable = bool(self.var_startup_all.get())
        set_startup(enable, user=False)

    def get_tray_title(self):
        current_app = self.tracker.current_app or "No active window"
        return f"Tracking: {current_app}"

    def minimize_to_tray(self):
        self.root.withdraw()
        if not self.is_tray_active:
            if not os.path.exists(ICON_PATH):
                return
            image = Image.open(ICON_PATH).resize((64, 64), Image.LANCZOS)
            menu = pystray.Menu(
                pystray.MenuItem("Restore", self.restore_from_tray),
                pystray.MenuItem("Exit", self.exit_app)
            )
            self.tray_icon = pystray.Icon("AllSeeingEye", image, self.get_tray_title(), menu)
            self.tray_icon.run_detached()
            self.is_tray_active = True

    def restore_from_tray(self, icon=None, item=None):
        self.root.deiconify()
        if self.tray_icon:
            self.tray_icon.stop()
            self.tray_icon = None
            self.is_tray_active = False

    def exit_app(self, icon=None, item=None):
        self.tracker.stop()
        if self.tray_icon:
            self.tray_icon.stop()
        self.root.destroy()

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    tracker = AppTracker()
    tracker.start()
    gui = TrackerGUI(tracker)
    gui.run()
