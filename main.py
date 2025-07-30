import sys
import os
import time
import json
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
import win32gui
import win32process
import win32event
import win32api
import winerror
import psutil
import pystray
from PIL import Image, ImageDraw
import ctypes
from pathlib import Path

def get_tray_image():
    if ICON_PATH.exists():
        try:
            return Image.open(str(ICON_PATH)).resize((64, 64), Image.LANCZOS)
        except:
            pass
    img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rectangle([8, 8, 56, 56], outline=(0, 0, 0, 255))
    return img

def get_appdata_folder():
    if sys.platform == "win32":
        base = os.getenv("LOCALAPPDATA")
    else:
        base = str(Path.home())
    folder = Path(base) / "AllSeeingEye"
    folder.mkdir(parents=True, exist_ok=True)
    return folder

APPDATA_DIR = get_appdata_folder()
DATA_FILE = APPDATA_DIR / "data.json"
DATA_BAK = APPDATA_DIR / "data.json.bak"
SETTINGS_FILE = APPDATA_DIR / "settings.json"
SETTINGS_BAK = APPDATA_DIR / "settings.json.bak"
LAST_USED_FILE = APPDATA_DIR / "last_used.json"
LAST_USED_BAK = APPDATA_DIR / "last_used.json.bak"
ICON_PATH = Path(sys._MEIPASS, "trayicon.ico") if hasattr(sys, "_MEIPASS") else Path("trayicon.ico")
TRACK_INTERVAL = 1
SAVE_INTERVAL = 5

def atomic_write_json(path: Path, data, bak_path: Path):
    tmp_path = path.with_suffix(path.suffix + ".tmp")
    with tmp_path.open("w", encoding="utf-8") as f:
        json.dump(data, f)
        f.flush()
        os.fsync(f.fileno())
    if path.exists():
        try:
            if bak_path:
                try:
                    if bak_path.exists():
                        bak_path.unlink()
                except:
                    pass
                try:
                    os.replace(path, bak_path)
                except:
                    pass
        except:
            pass
    os.replace(tmp_path, path)

def load_json_with_backup(path: Path, bak_path: Path):
    def _read(p: Path):
        try:
            if not p.exists():
                return None
            s = p.read_text(encoding="utf-8").strip()
            if not s:
                return None
            return json.loads(s)
        except:
            return None
    data = _read(path)
    if data is not None:
        return data
    data = _read(bak_path)
    if data is not None:
        try:
            atomic_write_json(path, data, bak_path)
        except:
            pass
        return data
    return {}

def _startup_shortcut_path():
    return Path(os.getenv("APPDATA")) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup" / "AllSeeingEye.lnk"

def _get_startup_command_parts():
    if getattr(sys, "frozen", False):
        return str(sys.executable), ""
    pythonw = Path(sys.executable).with_name("pythonw.exe")
    py = pythonw if pythonw.exists() else Path(sys.executable)
    script = str(Path(__file__).resolve())
    return str(py), f'"{script}"'

def is_startup_enabled(user=True):
    if user:
        return _startup_shortcut_path().exists()
    try:
        import winreg
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_READ) as key:
            try:
                _ = winreg.QueryValueEx(key, "AllSeeingEye")
                return True
            except FileNotFoundError:
                return False
    except Exception:
        return False

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def relaunch_as_admin_with_task(task):
    params = []
    for a in sys.argv[1:]:
        if not a.startswith("--apply-startup="):
            params.append(a)
    params.append(f"--apply-startup={task}")
    param_str = " ".join(f'"{p}"' if " " in p and not (p.startswith('"') and p.endswith('"')) else p for p in params)
    try:
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, param_str, None, 1)
        return True
    except:
        return False

def set_startup_hklm_no_elev(enable):
    try:
        import winreg
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_WRITE) as key:
            if enable:
                target, args = _get_startup_command_parts()
                cmd = f'"{target}" {args}'.strip()
                winreg.SetValueEx(key, "AllSeeingEye", 0, winreg.REG_SZ, cmd)
            else:
                try:
                    winreg.DeleteValue(key, "AllSeeingEye")
                except FileNotFoundError:
                    pass
    except Exception as e:
        try:
            messagebox.showerror("Startup Setting Error", f"Error setting startup: {e}")
        except:
            pass

def set_startup(enable, user=True):
    if user:
        lnk = _startup_shortcut_path()
        if enable:
            try:
                import win32com.client
                target, args = _get_startup_command_parts()
                shell = win32com.client.Dispatch("WScript.Shell")
                shortcut = shell.CreateShortcut(str(lnk))
                shortcut.TargetPath = target
                shortcut.Arguments = args
                shortcut.WorkingDirectory = str(Path(target).parent if getattr(sys, "frozen", False) else Path(__file__).parent)
                icon = str(ICON_PATH if ICON_PATH.exists() else target)
                shortcut.IconLocation = icon
                shortcut.Save()
            except Exception as e:
                try:
                    messagebox.showerror("Startup Setting Error", f"Error setting startup: {e}")
                except:
                    pass
        else:
            try:
                if lnk.exists():
                    lnk.unlink()
            except:
                pass
        return
    if not is_admin():
        task = "hklm_on" if enable else "hklm_off"
        if relaunch_as_admin_with_task(task):
            return
        else:
            try:
                messagebox.showerror("Admin Required", "You must run as administrator to set startup for all users.")
            except:
                pass
            return
    set_startup_hklm_no_elev(enable)

def ensure_single_instance():
    h_mutex = win32event.CreateMutex(None, False, "Global\\AllSeeingEyeMutex")
    if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
        try:
            messagebox.showinfo("AllSeeingEye", "Already running.")
        except:
            pass
        sys.exit(0)
    return h_mutex

class AppTracker:
    def __init__(self):
        self.running = True
        self.data = load_json_with_backup(DATA_FILE, DATA_BAK)
        self.settings = load_json_with_backup(SETTINGS_FILE, SETTINGS_BAK)
        self.recent_apps = []
        self.current_app = None
        self.start_time = time.time()
        self.lock = threading.Lock()
        self.last_used = load_json_with_backup(LAST_USED_FILE, LAST_USED_BAK)
        self.whitelist = set(self.settings.get("whitelist", []))
        self.idle_threshold = 300
        self._dirty_data = False
        self._dirty_settings = False
        self._dirty_last_used = False
        self._last_save = time.time()

    def start(self):
        threading.Thread(target=self.track_loop, daemon=True).start()

    def _maybe_save(self, force=False):
        now = time.time()
        if force or (now - self._last_save) >= SAVE_INTERVAL:
            if self._dirty_data:
                atomic_write_json(DATA_FILE, self.data, DATA_BAK)
                self._dirty_data = False
            if self._dirty_last_used:
                atomic_write_json(LAST_USED_FILE, self.last_used, LAST_USED_BAK)
                self._dirty_last_used = False
            if self._dirty_settings:
                atomic_write_json(SETTINGS_FILE, self.settings, SETTINGS_BAK)
                self._dirty_settings = False
            self._last_save = now

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
                            self._dirty_data = True
                            self._dirty_last_used = True
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
                                self._dirty_data = True
                                self._dirty_last_used = True
                            self.current_app = app
                            self.start_time = now
            self._maybe_save(False)
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
                self._dirty_data = True
                self._dirty_last_used = True
        self._maybe_save(True)

class TrackerGUI:
    def __init__(self, tracker):
        self.tracker = tracker
        self.root = tk.Tk()
        self.root.title("AllSeeingEye")
        self.root.geometry("400x400")
        if ICON_PATH.exists():
            try:
                self.root.iconbitmap(str(ICON_PATH))
            except Exception:
                pass
        self.recent_labels = [tk.Label(self.root, text="") for _ in range(3)]
        for lbl in self.recent_labels:
            lbl.pack(pady=2)
        self.export_btn = tk.Button(self.root, text="Export Data", command=self.export_data)
        self.export_btn.pack(pady=10)
        self.whitelist_frame = tk.Frame(self.root)
        self.whitelist_frame.pack(pady=10, fill="both", expand=True)
        tk.Label(self.whitelist_frame, text="Whitelisted Apps (empty = all):").pack()
        self.whitelist_listbox = tk.Listbox(self.whitelist_frame, height=6)
        self.whitelist_listbox.pack(side="left", fill="both", expand=True)
        scrollbar = tk.Scrollbar(self.whitelist_frame)
        scrollbar.pack(side="right", fill="y")
        self.whitelist_listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.whitelist_listbox.yview)
        self.btn_frame = tk.Frame(self.root)
        self.btn_frame.pack()
        self.add_btn = tk.Button(self.btn_frame, text="Add EXE to whitelist", command=self.add_whitelist)
        self.add_btn.grid(row=0, column=0, padx=5, pady=5)
        self.remove_btn = tk.Button(self.btn_frame, text="Remove Selected", command=self.remove_whitelist)
        self.remove_btn.grid(row=0, column=1, padx=5, pady=5)
        self.var_dark_mode = tk.IntVar(value=self.tracker.settings.get("dark_mode", 0))
        self.dark_mode_chk = tk.Checkbutton(self.root, text="Dark Mode", variable=self.var_dark_mode, command=self.toggle_dark_mode)
        self.dark_mode_chk.pack(pady=5)
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

    def build_tray_menu(self):
        return pystray.Menu(
            pystray.MenuItem("Restore", self.restore_from_tray),
            pystray.MenuItem("Exit", self.exit_app)
        )

    def update_ui(self):
        self.tracker.lock.acquire()
        recent = list(self.tracker.recent_apps)
        data = dict(self.tracker.data)
        last_used = dict(self.tracker.last_used)
        self.tracker.lock.release()
        for i in range(3):
            if i < len(recent):
                app = recent[i]
                seconds = int(data.get(app, 0))
                time_str = time.strftime("%H:%M:%S", time.gmtime(seconds))
                last_used_ts = last_used.get(app)
                last_used_str = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(last_used_ts)) if last_used_ts else "Never"
                self.recent_labels[i].config(text=f"{app} - {time_str} (Last used: {last_used_str})")
            else:
                self.recent_labels[i].config(text="")
        self.root.after(1000, self.update_ui)

    def export_data(self):
        def fmt_ago(now_ts, last_ts):
            if not last_ts:
                return "Never"
            diff = max(0, now_ts - last_ts)
            if diff < 60:
                return f"{int(diff)} seconds ago"
            if diff < 3600:
                return f"{int(diff // 60)} minutes ago"
            if diff < 86400:
                return f"{int(diff // 3600)} hours ago"
            return f"{int(diff // 86400)} days ago"

        sort_by_most_used = messagebox.askyesno(
            "Export Sort",
            "Sort by most used?\nYes = Most used\nNo = Last used"
        )

        filepath = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text Files", "*.txt")])
        if not filepath:
            return

        now = time.time()
        self.tracker.lock.acquire()
        seconds_map = dict(self.tracker.data)
        last_used_map = dict(self.tracker.last_used)
        self.tracker.lock.release()

        if sort_by_most_used:
            order = [k for k, _ in sorted(seconds_map.items(), key=lambda x: -x[1])]
        else:
            order = sorted(seconds_map.keys(), key=lambda a: last_used_map.get(a, 0), reverse=True)

        with open(filepath, "w", encoding="utf-8") as f:
            for app in order:
                seconds = int(seconds_map.get(app, 0))
                time_str = time.strftime("%H:%M:%S", time.gmtime(seconds))
                ts = last_used_map.get(app)
                if ts:
                    last_used_str = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(ts))
                    ago_str = fmt_ago(now, ts)
                    f.write(f"{app}: {time_str} (Last used: {last_used_str} | {ago_str})\n")
                else:
                    f.write(f"{app}: {time_str} (Last used: Never)\n")

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
        self.tracker._dirty_settings = False
        atomic_write_json(SETTINGS_FILE, self.tracker.settings, SETTINGS_BAK)

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
            if cls == "Frame":
                widget.configure(bg=bg)
                for child in widget.winfo_children():
                    try:
                        child.configure(bg=bg, fg=fg)
                        if isinstance(child, tk.Button):
                            child.configure(bg=btn_bg, fg=btn_fg, borderwidth=2, relief="raised", highlightthickness=1)
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
                        widget.configure(bg=btn_bg, fg=btn_fg, borderwidth=2, relief="raised", highlightthickness=1)
                    if isinstance(widget, tk.Checkbutton):
                        widget.configure(bg=bg, fg=fg, selectcolor=select_color)
                    if isinstance(widget, tk.Listbox):
                        widget.configure(bg=bg, fg=fg, borderwidth=2, highlightthickness=1)
                except:
                    pass

    def toggle_startup_user(self):
        enable = bool(self.var_startup_user.get())
        if enable and self.var_startup_all.get() == 1:
            self.var_startup_all.set(0)
            set_startup(False, user=False)
        set_startup(enable, user=True)

    def toggle_startup_all(self):
        enable = bool(self.var_startup_all.get())
        if enable and self.var_startup_user.get() == 1:
            self.var_startup_user.set(0)
            set_startup(False, user=True)
        set_startup(enable, user=False)

    def get_tray_title(self):
        current_app = self.tracker.current_app or "No active window"
        return f"Tracking: {current_app}"

    def minimize_to_tray(self):
        self.root.withdraw()
        if not self.is_tray_active:
            if not ICON_PATH.exists():
                return
            image = Image.open(str(ICON_PATH)).resize((64, 64), Image.LANCZOS)
            menu = self.build_tray_menu()
            self.tray_icon = pystray.Icon("AllSeeingEye", image, self.get_tray_title(), menu)
            self.tray_icon.run_detached()
            self.is_tray_active = True

    def restore_from_tray(self, icon=None, item=None):
        self.root.after(0, self._restore_main)

    def _restore_main(self):
        self.root.deiconify()
        self._stop_tray_async()

    def _stop_tray_icon(self, icon):
        try:
            icon.visible = False
            icon.stop()
        except:
            pass

    def _stop_tray_async(self):
        icon = self.tray_icon
        self.tray_icon = None
        self.is_tray_active = False
        if icon:
            threading.Thread(target=self._stop_tray_icon, args=(icon,), daemon=True).start()

    def exit_app(self, icon=None, item=None):
        self.root.after(0, self._exit_main)

    def _exit_main(self):
        self.tracker.stop()
        self._stop_tray_async()
        self.root.quit()
        self.root.destroy()

    def run(self):
        self.root.mainloop()

def get_active_app():
    try:
        hwnd = win32gui.GetForegroundWindow()
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        process = psutil.Process(pid)
        return process.name()
    except Exception:
        return None

def apply_startup_task_arg_if_present():
    for a in sys.argv[1:]:
        if a.startswith("--apply-startup="):
            v = a.split("=", 1)[1]
            if v == "hklm_on":
                set_startup_hklm_no_elev(True)
            elif v == "hklm_off":
                set_startup_hklm_no_elev(False)
            sys.exit(0)

if __name__ == "__main__":
    apply_startup_task_arg_if_present()
    _mutex = ensure_single_instance()
    tracker = AppTracker()
    tracker.start()
    gui = TrackerGUI(tracker)
    gui.run()
