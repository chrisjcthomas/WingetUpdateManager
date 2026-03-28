#!/usr/bin/env python3
"""
Winget Update Manager - A Modern GUI wrapper for winget
Full-featured: Dashboard, Installed Apps, Updates, History, Settings.
"""

import subprocess
import threading
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import sys
import os
import json
import csv
import re
import webbrowser
import ctypes
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from ctypes import wintypes
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlparse
from urllib.request import Request, urlopen

def get_resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return Path(sys._MEIPASS) / relative_path
    return Path(__file__).parent / relative_path

try:
    import winreg
except ImportError:
    winreg = None

try:
    from PIL import Image, ImageDraw, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False
    Image = ImageDraw = ImageTk = None

try:
    from winotify import Notification as WinNotification
    HAS_WINOTIFY = True
except ImportError:
    HAS_WINOTIFY = False

try:
    import pystray
    HAS_TRAY = True
except ImportError:
    HAS_TRAY = False

VERSION = "3.0.0"

APP_DIR = Path(os.environ.get("APPDATA", str(Path.home()))) / "WingetUpdateManager"
CONFIG_FILE = APP_DIR / "config.json"
HISTORY_FILE = APP_DIR / "history.json"
CACHE_FILE = APP_DIR / "scan_cache.json"

DARK_COLORS = {
    "bg": "#182235",
    "surface": "#202b3f",
    "surface_alt": "#223149",
    "surface_hover": "#2a3954",
    "surface_soft": "#1c2638",
    "border": "#31415e",
    "border_soft": "#293753",
    "primary": "#5ea2ff",
    "primary_dim": "#447ed2",
    "secondary": "#21c087",
    "accent": "#8b5cf6",
    "text_main": "#f3f7ff",
    "text_dim": "#8fa2c3",
    "text_soft": "#6f84a8",
    "danger": "#ff5d6c",
    "warning": "#f4b740",
    "console_bg": "#07101f",
    "console_header": "#0d1727",
    "success": "#2dde98",
    "card_bg": "#202b3f",
    "card_inner": "#1b2739",
    "nav_active": "#2a3b5a",
    "nav_badge_bg": "#194e5a",
    "nav_badge_fg": "#7df2ba",
}

LIGHT_COLORS = {
    "bg": "#eef4fb",
    "surface": "#ffffff",
    "surface_alt": "#f5f8fc",
    "surface_hover": "#eef4fb",
    "surface_soft": "#f7f9fd",
    "border": "#d7e1ef",
    "border_soft": "#e4ebf5",
    "primary": "#377dff",
    "primary_dim": "#2d64d1",
    "secondary": "#19b57d",
    "accent": "#8b5cf6",
    "text_main": "#142033",
    "text_dim": "#5e718e",
    "text_soft": "#7d8faa",
    "danger": "#e5485d",
    "warning": "#ca8a04",
    "console_bg": "#0f172a",
    "console_header": "#182436",
    "success": "#109e69",
    "card_bg": "#ffffff",
    "card_inner": "#f6f8fc",
    "nav_active": "#dfeafb",
    "nav_badge_bg": "#ddf7eb",
    "nav_badge_fg": "#0f8d61",
}

FONTS = {
    "title": ("Segoe UI", 21, "bold"),
    "header": ("Segoe UI", 19, "bold"),
    "sub_header": ("Segoe UI", 14, "bold"),
    "body": ("Segoe UI", 10),
    "body_bold": ("Segoe UI", 10, "bold"),
    "small": ("Segoe UI", 9),
    "small_bold": ("Segoe UI", 9, "bold"),
    "micro": ("Segoe UI", 8, "bold"),
    "nav": ("Segoe UI", 10, "bold"),
    "mono": ("Consolas", 10),
    "mono_small": ("Consolas", 9),
}

SPACING = {
    "page_x": 28,
    "page_y": 24,
    "card_gap": 18,
    "card_pad": 22,
    "control_gap": 10,
    "row_pad_y": 14,
}

DEFAULT_CONFIG = {
    "theme": "dark",
    "auto_check_on_launch": True,
    "auto_check_interval_hours": 24,
    "excluded_packages": [],
    "update_groups": {},
    "quiet_mode_packages": [],
    "quiet_mode_enabled": False,
    "silent_mode": True,
    "scheduled_scan": False,
    "cache_ttl": 3600,
    "parallel_workers": 1,
    "start_at_login": False,
    "window_geometry": "1100x750",
    "minimize_to_tray": False,
}


if os.name == "nt":
    class ICONINFO(ctypes.Structure):
        _fields_ = [
            ("fIcon", wintypes.BOOL),
            ("xHotspot", wintypes.DWORD),
            ("yHotspot", wintypes.DWORD),
            ("hbmMask", wintypes.HBITMAP),
            ("hbmColor", wintypes.HBITMAP),
        ]


    class BITMAP(ctypes.Structure):
        _fields_ = [
            ("bmType", wintypes.LONG),
            ("bmWidth", wintypes.LONG),
            ("bmHeight", wintypes.LONG),
            ("bmWidthBytes", wintypes.LONG),
            ("bmPlanes", wintypes.WORD),
            ("bmBitsPixel", wintypes.WORD),
            ("bmBits", wintypes.LPVOID),
        ]


    class BITMAPINFOHEADER(ctypes.Structure):
        _fields_ = [
            ("biSize", wintypes.DWORD),
            ("biWidth", wintypes.LONG),
            ("biHeight", wintypes.LONG),
            ("biPlanes", wintypes.WORD),
            ("biBitCount", wintypes.WORD),
            ("biCompression", wintypes.DWORD),
            ("biSizeImage", wintypes.DWORD),
            ("biXPelsPerMeter", wintypes.LONG),
            ("biYPelsPerMeter", wintypes.LONG),
            ("biClrUsed", wintypes.DWORD),
            ("biClrImportant", wintypes.DWORD),
        ]


    class RGBQUAD(ctypes.Structure):
        _fields_ = [
            ("rgbBlue", ctypes.c_byte),
            ("rgbGreen", ctypes.c_byte),
            ("rgbRed", ctypes.c_byte),
            ("rgbReserved", ctypes.c_byte),
        ]


    class BITMAPINFO(ctypes.Structure):
        _fields_ = [
            ("bmiHeader", BITMAPINFOHEADER),
            ("bmiColors", RGBQUAD * 1),
        ]


class Config:
    def __init__(self):
        APP_DIR.mkdir(parents=True, exist_ok=True)
        self.data = DEFAULT_CONFIG.copy()
        self.load()

    def load(self):
        try:
            if CONFIG_FILE.exists():
                with open(CONFIG_FILE, "r") as f:
                    saved = json.load(f)
                for k, v in saved.items():
                    if k in self.data:
                        self.data[k] = v
        except Exception:
            pass

    def save(self):
        try:
            with open(CONFIG_FILE, "w") as f:
                json.dump(self.data, f, indent=2)
        except Exception:
            pass

    def get(self, key, default=None):
        return self.data.get(key, default)

    def set(self, key, value):
        self.data[key] = value
        self.save()


class HistoryManager:
    def __init__(self):
        APP_DIR.mkdir(parents=True, exist_ok=True)
        self.entries = []
        self.rollback_ledger = {}
        self.load()

    def load(self):
        try:
            if HISTORY_FILE.exists():
                with open(HISTORY_FILE, "r") as f:
                    raw = json.load(f)
                if isinstance(raw, list):
                    self.entries = raw
                elif isinstance(raw, dict):
                    self.entries = raw.get("entries", [])
                    self.rollback_ledger = self._normalize_ledger(raw.get("rollback_ledger", {}))
        except Exception:
            pass
        if not self.rollback_ledger:
            self._rebuild_ledger_from_entries()

    def save(self):
        try:
            with open(HISTORY_FILE, "w") as f:
                json.dump({"entries": self.entries, "rollback_ledger": self.rollback_ledger}, f, indent=2)
        except Exception:
            pass

    def _normalize_ledger(self, ledger):
        if not isinstance(ledger, dict): return {}
        normalized = {}
        for pkg_id, items in ledger.items():
            clean = [{"version": str(i.get("version", "")).strip(), 
                      "manager": i.get("manager", "winget"), 
                      "timestamp": i.get("timestamp", datetime.now().isoformat())} 
                     for i in items if isinstance(i, dict) and str(i.get("version", "")).strip()]
            if clean: normalized[str(pkg_id)] = clean[-100:]
        return normalized

    def _append_ledger_version(self, package_id, version, manager="winget", timestamp=None):
        pkg_id, ver = str(package_id or "").strip(), str(version or "").strip()
        if not pkg_id or not ver: return
        items = self.rollback_ledger.setdefault(pkg_id, [])
        entry = {"version": ver, "manager": manager, "timestamp": timestamp or datetime.now().isoformat()}
        if not items or items[-1]["version"] != ver:
            items.append(entry)
        else:
            items[-1] = entry
        self.rollback_ledger[pkg_id] = items[-100:]

    def _rebuild_ledger_from_entries(self):
        for entry in reversed(self.entries):
            if entry.get("status") == "success" and entry.get("manager", "winget") == "winget":
                pkg_id = entry.get("package_id")
                for ver_key in ["old_version", "new_version"]:
                    if entry.get(ver_key):
                        self._append_ledger_version(pkg_id, entry[ver_key], "winget", entry.get("timestamp"))

    def add(self, package_id, package_name, old_ver, new_ver, status, manager="winget"):
        self.entries.insert(0, {
            "timestamp": datetime.now().isoformat(),
            "package_id": package_id,
            "package_name": package_name,
            "old_version": old_ver,
            "new_version": new_ver,
            "status": status,
            "manager": manager,
        })
        self.entries = self.entries[:500]
        self.save()

    def record_version(self, package_id, version, manager="winget", timestamp=None):
        if manager == "winget":
            self._append_ledger_version(package_id, version, manager, timestamp)
            self.save()

    def previous_version(self, package_id, current_version, manager="winget"):
        if manager != "winget": return None
        cur = str(current_version or "").strip()
        for item in reversed(self.rollback_ledger.get(package_id, [])):
            if item["version"] != cur: return item["version"]
        return None

    def get_entries(self, limit=100):
        return self.entries[:limit]

    def clear(self):
        self.entries, self.rollback_ledger = [], {}
        self.save()


class ScanCache:
    def __init__(self):
        APP_DIR.mkdir(parents=True, exist_ok=True)
        self.data = {"updates": [], "installed": [], "timestamp": None}
        self.load()

    def load(self):
        try:
            if CACHE_FILE.exists():
                with open(CACHE_FILE, "r") as f:
                    self.data = json.load(f)
        except Exception:
            self.data = {"updates": [], "installed": [], "timestamp": None}

    def save(self, updates=None, installed=None):
        if updates is not None:
            self.data["updates"] = updates
        if installed is not None:
            self.data["installed"] = installed
        self.data["timestamp"] = datetime.now().isoformat()
        try:
            with open(CACHE_FILE, "w") as f:
                json.dump(self.data, f, indent=2)
        except Exception:
            pass

    def get_updates(self):
        return self.data.get("updates", [])

    def get_installed(self):
        return self.data.get("installed", [])

    def age_seconds(self):
        ts = self.data.get("timestamp")
        if not ts:
            return float("inf")
        try:
            return (datetime.now() - datetime.fromisoformat(ts)).total_seconds()
        except Exception:
            return float("inf")

    def is_stale(self, ttl=3600):
        return self.age_seconds() > ttl

    def last_scan_time(self):
        ts = self.data.get("timestamp")
        if not ts:
            return "Never"
        try:
            return datetime.fromisoformat(ts).strftime("%H:%M")
        except Exception:
            return "Unknown"


class WingetParser:
    @staticmethod
    def _parse_tabular_output(output, required_cols):
        items = []
        lines = output.strip().split("\n")
        header_idx = -1
        for i, line in enumerate(lines):
            if all(col in line for col in required_cols):
                header_idx = i
                break
        if header_idx == -1:
            return []

        header = lines[header_idx]
        cols = {col.lower(): header.find(col) for col in ["Name", "Id", "Version", "Available", "Source"] if header.find(col) >= 0}
        sorted_cols = sorted(cols.items(), key=lambda x: x[1])

        for line in lines[header_idx + 1:]:
            if not line.strip() or line.strip().startswith("-") or "upgrades available" in line.lower():
                continue
            if len(line) < max(cols.values()):
                continue
            try:
                parts = {}
                for j, (name, start) in enumerate(sorted_cols):
                    end = sorted_cols[j + 1][1] if j + 1 < len(sorted_cols) else len(line)
                    parts[name] = line[start:end].strip()
                if all(parts.get(col.lower()) for col in required_cols):
                    items.append(parts)
            except Exception:
                continue
        return items

    @staticmethod
    def parse_upgrade_output(output):
        raw_items = WingetParser._parse_tabular_output(output, ["Name", "Id", "Version", "Available"])
        return [
            {
                "name": item["name"],
                "id": item["id"],
                "version": item["version"],
                "available": item["available"],
                "manager": "winget",
            }
            for item in raw_items if item["available"] != item["version"]
        ]

    @staticmethod
    def parse_list_output(output):
        raw_items = WingetParser._parse_tabular_output(output, ["Name", "Id"])
        return [
            {
                "name": item["name"],
                "id": item["id"],
                "version": item.get("version", ""),
                "source": item.get("source", ""),
                "manager": "winget",
            }
            for item in raw_items
        ]


class ScrollableFrame(tk.Frame):
    def __init__(self, parent, bg_color, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0, bg=bg_color)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg=bg_color)
        self.bg_color = bg_color

        self.scrollable_frame.bind("<Configure>", self._on_frame_configure)
        self.window_id = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.canvas.bind("<Enter>", self._bind_mousewheel)
        self.canvas.bind("<Leave>", self._unbind_mousewheel)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

    def _on_frame_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.window_id, width=event.width)

    def _bind_mousewheel(self, event):
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _unbind_mousewheel(self, event):
        self.canvas.unbind_all("<MouseWheel>")

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


class Toast:
    _notifications = []
    _change_callback = None

    def __init__(self, root, message, type_="info", duration=3000, record=True):
        colors = {
            "info": "#3b82f6", "success": "#10b981",
            "error": "#ef4444", "warning": "#f59e0b",
        }
        if record:
            Toast._notifications.insert(0, {
                "message": message, "type": type_,
                "timestamp": datetime.now().strftime("%H:%M:%S"),
            })
            Toast._notifications = Toast._notifications[:50]
            if callable(Toast._change_callback):
                try:
                    root.after(0, Toast._change_callback)
                except Exception:
                    pass
        bg = colors.get(type_, "#3b82f6")
        self.top = tk.Toplevel(root)
        self.top.overrideredirect(True)
        self.top.attributes('-topmost', True)
        frame = tk.Frame(self.top, bg=bg, padx=20, pady=12)
        frame.pack(fill=tk.BOTH, expand=True)
        tk.Label(frame, text=message, bg=bg, fg="white",
                 font=("Segoe UI", 10, "bold")).pack()
        root.update_idletasks()
        rx = root.winfo_rootx() + root.winfo_width() - 320
        ry = root.winfo_rooty() + 10
        self.top.geometry(f"300x44+{rx}+{ry}")
        self.top.after(duration, self._dismiss)

    def _dismiss(self):
        try:
            self.top.destroy()
        except Exception:
            pass


class WingetUpdateManager:
    def __init__(self, root):
        self.root = root
        self.config = Config()
        self.history = HistoryManager()
        self.scan_cache = ScanCache()
        self.parser = WingetParser()
        self.scan_only_mode = "--scan-only" in sys.argv
        self.minimized_mode = "--minimized" in sys.argv

        self.root.title("Winget Update Manager")
        icon_path = get_resource_path("assets/icon.png")
        if HAS_PIL and icon_path.exists():
            try:
                img = Image.open(str(icon_path)).convert("RGBA")
                self.app_icon = ImageTk.PhotoImage(img)
                self.root.iconphoto(True, self.app_icon)
            except Exception:
                pass
        geo = self.config.get("window_geometry", "1100x750")
        self.root.geometry(geo)
        self.root.minsize(1120, 720)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        self.current_page = "updates"
        self.updates_list = []
        self.installed_list = []
        self.checkboxes = {}
        self.is_checking = False
        self.is_updating = False
        self.all_selected_var = False
        self.sort_column = None
        self.sort_ascending = True
        self.search_var = tk.StringVar()
        self.installed_search_var = tk.StringVar()
        self.group_filter_var = tk.StringVar(value="All Packages")
        self.active_process = None
        self.active_processes = set()
        self.active_process_lock = threading.Lock()
        self.cancel_requested = threading.Event()
        self.update_count = 0
        self.tray_icon = None
        self.console_frame = None
        self.console_idle_height = 168
        self.console_active_height = 238
        self.console_anim_job = None
        self.pages = {}
        self.nav_items = {}
        self.nav_accents = {}
        self.header_dividers = {}
        self.registry_app_index = None
        self.registry_apps_by_name = {}
        self.package_icon_refs = {}
        self.tk_icon_cache = {}
        self.brand_photo = None
        self.focused_row_index = -1
        self.row_widgets = []
        self.notifications = Toast._notifications
        self.notification_badges = []
        self.discover_busy = False
        self.discover_action_var = tk.StringVar(value="Idle")
        self.selected_group_name = None

        self.colors = DARK_COLORS if self.config.get("theme") == "dark" else LIGHT_COLORS
        self.root.configure(bg=self.colors["bg"])

        Toast._change_callback = self._refresh_notification_indicators
        self.setup_styles()
        self.setup_ui()
        self.bind_shortcuts()

        if self.config.get("auto_check_on_launch") and not self.scan_only_mode:
            self.root.after(500, self.check_updates)

        self.root.after(100, self.verify_winget)
        self._load_from_cache()

    def _load_from_cache(self):
        cached_updates = self.scan_cache.get_updates()
        cached_installed = self.scan_cache.get_installed()
        if cached_updates:
            self.updates_list = cached_updates
            for pkg in self.updates_list:
                pkg["icon_ref"] = self._get_package_icon_ref(pkg) if hasattr(self, '_get_package_icon_ref') else None
            self.root.after(200, self._render_updates)
        if cached_installed:
            self.installed_list = cached_installed
            self.root.after(300, self._render_installed)

    def verify_winget(self):
        def check():
            try:
                result = subprocess.run(
                    ["winget", "--version"],
                    capture_output=True, text=True, timeout=10,
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
                if result.returncode != 0:
                    raise Exception("winget returned non-zero")
            except FileNotFoundError:
                self.root.after(0, self._show_winget_missing)
            except Exception:
                pass
        threading.Thread(target=check, daemon=True).start()

    def _show_winget_missing(self):
        messagebox.showwarning(
            "Winget Not Found",
            "winget is not installed or not in PATH.\n\n"
            "Install it from the Microsoft Store (App Installer)\n"
            "or visit: https://github.com/microsoft/winget-cli"
        )

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Vertical.TScrollbar",
            gripcount=0,
            background=self.colors["surface_hover"],
            darkcolor=self.colors["surface_soft"],
            lightcolor=self.colors["surface_hover"],
            troughcolor=self.colors["card_bg"],
            bordercolor=self.colors["card_bg"],
            arrowcolor=self.colors["text_dim"]
        )

    def setup_ui(self):
        self.sidebar = tk.Frame(self.root, bg=self.colors["surface_soft"], width=260)
        self.sidebar.pack(side=tk.LEFT, fill=tk.Y)
        self.sidebar.pack_propagate(False)
        self._build_sidebar()

        self.info_panel = tk.Frame(self.root, bg=self.colors["surface"], width=0,
                                   highlightthickness=1,
                                   highlightbackground=self.colors["border"])
        self.info_panel.pack_propagate(False)
        self._build_info_panel()

        self.main_container = tk.Frame(self.root, bg=self.colors["bg"])
        self.main_container.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._build_dashboard_page()
        self._build_installed_page()
        self._build_updates_page()
        self._build_discover_page()
        self._build_history_page()
        self._build_settings_page()

        self.switch_page("updates")

    # ------------------------------------------------------------------ sidebar
    def _build_sidebar(self):
        brand_bg = self.colors["surface_soft"]
        brand = tk.Frame(self.sidebar, bg=brand_bg, padx=18, pady=20)
        brand.pack(fill=tk.X)
        
        icon_wrap = tk.Frame(brand, bg=brand_bg, width=48, height=48)
        icon_wrap.pack(side=tk.LEFT)
        icon_wrap.pack_propagate(False)

        img_path = get_resource_path("assets/icon.png")
        if img_path.exists() and HAS_PIL:
            try:
                img = Image.open(str(img_path)).convert("RGBA")
                img = img.resize((48, 48), Image.LANCZOS)
                self.brand_photo = ImageTk.PhotoImage(img, master=self.root)
                tk.Label(icon_wrap, image=self.brand_photo, borderwidth=0, bg=brand_bg).pack(expand=True)
            except Exception:
                tk.Label(icon_wrap, text="\u25B7_", bg=self.colors["primary"],
                         fg="white", font=("Segoe UI", 14, "bold")).pack(expand=True)
        else:
            tk.Label(icon_wrap, text="\u25B7_", bg=self.colors["primary"],
                     fg="white", font=("Segoe UI", 14, "bold")).pack(expand=True)

        brand_text = tk.Frame(brand, bg=brand_bg)
        brand_text.pack(side=tk.LEFT, padx=(12, 0))
        tk.Label(brand_text, text="Winget UM", bg=brand_bg,
                 fg=self.colors["text_main"], font=("Segoe UI", 18, "bold")).pack(anchor="w")
        tk.Label(brand_text, text="Update workspace", bg=brand_bg,
                 fg=self.colors["primary"], font=FONTS["small"]).pack(anchor="w")

        nav = tk.Frame(self.sidebar, bg=self.colors["surface_soft"], padx=14, pady=8)
        nav.pack(fill=tk.BOTH, expand=True)

        for pid, label, icon in [
            ("dashboard", "Dashboard", "\u25A6"),
            ("installed", "Installed Apps", "\u25A4"),
            ("updates", "Updates", "\u27F3"),
            ("discover", "Discover", "\U0001F50D"),
            ("history", "History", "\u29D7"),
        ]:
            self._build_nav_item(nav, pid, label, icon)

        bottom = tk.Frame(self.sidebar, bg=self.colors["surface_soft"], padx=14, pady=12)
        bottom.pack(side=tk.BOTTOM, fill=tk.X)
        self._build_nav_item(bottom, "settings", "Settings", "\u2699")

        ver = tk.Frame(bottom, bg=self.colors["surface"], padx=14, pady=8,
                       highlightthickness=1,
                       highlightbackground=self.colors["border_soft"])
        ver.pack(fill=tk.X, pady=(12, 2))
        tk.Label(ver, text=f"\u25CF v{VERSION} stable", bg=self.colors["surface"],
                 fg=self.colors["text_soft"], font=("Consolas", 8)).pack(anchor="w")

    def _build_info_panel(self):
        header = tk.Frame(self.info_panel, bg=self.colors["surface_alt"], pady=18, padx=18)
        header.pack(fill=tk.X)
        self.info_title = tk.Label(header, text="Package Details", bg=self.colors["surface_alt"],
                                   fg=self.colors["text_main"], font=FONTS["sub_header"])
        self.info_title.pack(side=tk.LEFT)
        close_btn = tk.Label(header, text="\u2715", bg=self.colors["surface_alt"],
                             fg=self.colors["text_dim"], font=("Segoe UI", 12), cursor="hand2")
        close_btn.pack(side=tk.RIGHT)
        close_btn.bind("<Button-1>", lambda e: self.hide_info_panel())

        sc = ScrollableFrame(self.info_panel, bg_color=self.colors["surface"])
        sc.pack(fill=tk.BOTH, expand=True)
        self.info_content = sc.scrollable_frame

        self.info_msg = tk.Label(self.info_content, text="", bg=self.colors["surface"],
                                 fg=self.colors["text_dim"], font=FONTS["body"])
        self.info_msg.pack(pady=20)

        self.info_labels = {}
        for key in ["Description", "Homepage", "License", "Installer Type", "Publisher",
                    "Install Date", "Size", "Tags"]:
            f = tk.Frame(self.info_content, bg=self.colors["surface"], pady=5)
            f.pack(fill=tk.X, padx=18)
            tk.Label(f, text=key.upper(), bg=self.colors["surface"], fg=self.colors["text_soft"],
                     font=FONTS["micro"]).pack(anchor="w")
            lbl = tk.Label(f, text="---", bg=self.colors["surface"], fg=self.colors["text_main"],
                           font=FONTS["body"], wraplength=250, justify="left")
            lbl.pack(anchor="w")
            self.info_labels[key] = lbl

        btn_frame = tk.Frame(self.info_content, bg=self.colors["surface"], pady=12)
        btn_frame.pack(fill=tk.X, padx=18)
        self.info_homepage_btn = self._btn(btn_frame, "\U0001F310 Visit Homepage",
                                           self.colors["primary"], "white",
                                           lambda: self._open_info_homepage())
        self.info_homepage_btn.pack(fill=tk.X, pady=(0, 6))
        self.info_changelog_btn = self._btn(
            btn_frame, "\U0001F4DD View Changelog", self.colors["surface_alt"],
            self.colors["text_main"], lambda: self._open_info_changelog(), outline=True
        )
        self.info_changelog_btn.pack(fill=tk.X, pady=(0, 6))
        self.info_uninstall_btn = self._btn(btn_frame, "\U0001F5D1 Uninstall",
                                            self.colors["danger"], "white",
                                            lambda: self._uninstall_info_package())
        self.info_uninstall_btn.pack(fill=tk.X)
        self._info_current_pkg = None
        self._info_homepage_url = None
        self._info_release_notes_url = None
             
    def show_info_panel(self, pkg_id, pkg_name, manager="winget"):
        if not self.info_panel.winfo_ismapped():
            self.info_panel.pack(side=tk.RIGHT, fill=tk.Y, before=self.main_container)
        self._animate_panel_width(self.info_panel, 330)
        self.info_title.config(text=pkg_name)
        self.info_msg.config(text="Fetching details...")
        for lbl in self.info_labels.values():
            lbl.config(text="---")
        current_pkg = self._find_installed_package(pkg_id, manager)
        if not current_pkg:
            current_pkg = next(
                (pkg for pkg in self.updates_list if pkg.get("id") == pkg_id and pkg.get("manager", "winget") == manager),
                None,
            )
        self._info_current_pkg = {
            "id": pkg_id,
            "name": pkg_name,
            "manager": manager,
            "version": (current_pkg or {}).get("version", ""),
        }
        self._info_homepage_url = None
        self._info_release_notes_url = None
        self.info_homepage_btn.config(state=tk.DISABLED)
        self.info_changelog_btn.config(state=tk.DISABLED, bg=self.colors["surface_alt"])
             
        def fetch():
            try:
                if manager == "npm":
                    proc = subprocess.run(
                        [self._npm_command(), "view", pkg_id, "--json"],
                        capture_output=True, text=True, encoding='utf-8',
                        errors='ignore', creationflags=subprocess.CREATE_NO_WINDOW
                    )
                    if proc.returncode == 0:
                        raw = json.loads(proc.stdout or "{}")
                        publisher = raw.get("author")
                        if isinstance(publisher, dict):
                            publisher = publisher.get("name") or publisher.get("email")
                        if not publisher:
                            maintainers = raw.get("maintainers")
                            if isinstance(maintainers, list) and maintainers:
                                publisher = maintainers[0]
                        if isinstance(publisher, str):
                            publisher = publisher.split("<", 1)[0].strip()

                        homepage = raw.get("homepage")
                        if not homepage:
                            repo = raw.get("repository")
                            if isinstance(repo, dict):
                                homepage = repo.get("url")
                            elif isinstance(repo, str):
                                homepage = repo

                        data = {
                            "Description": raw.get("description") or "---",
                            "Homepage": homepage or "---",
                            "License": raw.get("license") or "---",
                            "Installer Type": "npm global package",
                            "Publisher": publisher or "npm registry",
                            "_release_notes_url": None,
                        }
                        self.root.after(0, lambda: self._update_info_panel(data))
                    else:
                        self.root.after(0, lambda: self.info_msg.config(text="Failed to fetch npm details."))
                else:
                    proc = subprocess.run(
                        ["winget", "show", "--id", pkg_id, "--exact", "--accept-source-agreements"],
                        capture_output=True, text=True, encoding='utf-8',
                        errors='ignore', creationflags=subprocess.CREATE_NO_WINDOW
                    )
                    if proc.returncode == 0:
                        data = self._parse_labelled_output(proc.stdout)
                        registry = self._match_registry_record(self._info_current_pkg)
                        if registry:
                            data.setdefault("Publisher", registry.get("publisher") or "---")
                            if registry.get("install_date") and not data.get("Install Date"):
                                data["Install Date"] = registry["install_date"]
                            if registry.get("estimated_size") and not data.get("Size"):
                                data["Size"] = registry["estimated_size"]
                        data["_release_notes_url"] = (
                            data.get("Release Notes Url")
                            or data.get("Release Notes URL")
                            or data.get("Release Notes")
                        )
                        self.root.after(0, lambda: self._update_info_panel(data))
                    else:
                        self.root.after(0, lambda: self.info_msg.config(text="Failed to fetch details."))
            except Exception:
                self.root.after(0, lambda: self.info_msg.config(text="Error fetching details."))
                
        threading.Thread(target=fetch, daemon=True).start()
        
    def _update_info_panel(self, data):
        self.info_msg.config(text="")
        mapping = {
            "Description": ["Description", "Short Description"],
            "Homepage": ["Homepage", "Publisher Url", "Author Url"],
            "License": ["License"],
            "Installer Type": ["Installer Type", "Installer"],
            "Publisher": ["Publisher"],
            "Install Date": ["Install Date", "InstallDate"],
            "Size": ["Size", "Installer Size", "Download Size"],
            "Tags": ["Tags", "Moniker"],
        }
        for ui_key, source_keys in mapping.items():
            for sk in source_keys:
                if sk in data and data[sk] and data[sk] != "---":
                    self.info_labels[ui_key].config(text=str(data[sk]))
                    break
        hp = data.get("Homepage") or data.get("Publisher Url") or data.get("Author Url")
        self._info_homepage_url = hp if hp and hp != "---" else None
        self._info_release_notes_url = data.get("_release_notes_url") or None
        self.info_homepage_btn.config(state=tk.NORMAL if self._info_homepage_url else tk.DISABLED)
        changelog_supported = (
            self._info_current_pkg
            and self._info_current_pkg.get("manager", "winget") == "winget"
            and bool(self._info_release_notes_url)
        )
        self.info_changelog_btn.config(
            state=tk.NORMAL if changelog_supported else tk.DISABLED,
            bg=self.colors["primary"] if changelog_supported else self.colors["surface_alt"],
        )

    def _open_info_homepage(self):
        if self._info_homepage_url:
            webbrowser.open(self._info_homepage_url)

    def _open_info_changelog(self):
        if self._info_current_pkg:
            self._show_changelog_for_package(self._info_current_pkg, self._info_release_notes_url)

    def _uninstall_info_package(self):
        if not self._info_current_pkg:
            return
        pkg = self._info_current_pkg
        if not messagebox.askyesno("Uninstall", f"Uninstall {pkg.get('name', pkg.get('id'))}?"):
            return
        manager = pkg.get("manager", "winget")
        def do_uninstall():
            try:
                if manager == "npm":
                    cmd = [self._npm_command(), "uninstall", "-g", pkg["id"]]
                else:
                    cmd = ["winget", "uninstall", "--id", pkg["id"], "--exact",
                           "--accept-source-agreements", "--disable-interactivity"]
                proc = subprocess.run(
                    cmd, capture_output=True, text=True, timeout=120,
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
                if proc.returncode != 0:
                    raise RuntimeError(self._tail_message(
                        proc.stderr or proc.stdout,
                        f"Exit code {proc.returncode}",
                    ))
                self.root.after(0, lambda: Toast(self.root, f"Uninstalled {pkg['id']}", "success"))
                self.root.after(200, self.load_installed)
                self.root.after(200, self.check_updates)
            except Exception as e:
                self.root.after(0, lambda: Toast(self.root, f"Uninstall failed: {e}", "error"))
        threading.Thread(target=do_uninstall, daemon=True).start()

    def _show_changelog_for_package(self, pkg, release_notes_url=None):
        if not pkg or pkg.get("manager", "winget") != "winget":
            Toast(self.root, "Changelog is only supported for winget packages", "info")
            return

        url = release_notes_url or ""
        if not url:
            try:
                proc = self._run_winget_capture(
                    ["show", "--id", pkg.get("id", ""), "--exact", "--accept-source-agreements"],
                    timeout=60,
                )
                if proc.returncode == 0:
                    data = self._parse_labelled_output(proc.stdout)
                    url = (
                        data.get("Release Notes Url")
                        or data.get("Release Notes URL")
                        or data.get("Release Notes")
                        or ""
                    )
            except Exception:
                url = ""
            if not url:
                Toast(self.root, "No changelog URL is available for this package", "warning")
                return

        repo = self._github_repo_from_url(url)
        if not repo:
            webbrowser.open(url)
            return

        owner, name = repo
        popup = tk.Toplevel(self.root)
        popup.title(f"Changelog - {pkg.get('name', pkg.get('id', 'Package'))}")
        popup.geometry("760x560")
        popup.configure(bg=self.colors["bg"])

        header = tk.Frame(popup, bg=self.colors["surface"], padx=18, pady=12)
        header.pack(fill=tk.X)
        tk.Label(
            header, text=f"{pkg.get('name', pkg.get('id', 'Package'))} changelog",
            bg=self.colors["surface"], fg=self.colors["text_main"], font=FONTS["sub_header"]
        ).pack(side=tk.LEFT)
        self._btn(
            header, "Open in Browser", self.colors["surface_alt"], self.colors["text_main"],
            lambda: webbrowser.open(url), outline=True
        ).pack(side=tk.RIGHT)

        body = scrolledtext.ScrolledText(
            popup, bg=self.colors["card_bg"], fg=self.colors["text_main"],
            font=FONTS["body"], relief=tk.FLAT, wrap=tk.WORD, padx=16, pady=16
        )
        body.pack(fill=tk.BOTH, expand=True, padx=18, pady=(0, 18))
        body.insert(tk.END, "Loading latest GitHub release notes...")
        body.config(state=tk.DISABLED)

        def load_notes():
            try:
                notes = self._load_github_release_notes(owner, name)
            except Exception:
                notes = None

            def finish():
                if not notes:
                    popup.destroy()
                    webbrowser.open(url)
                    Toast(self.root, "Opened changelog in browser", "info")
                    return
                body.config(state=tk.NORMAL)
                body.delete("1.0", tk.END)
                title = notes.get("title") or f"{owner}/{name} latest release"
                browser_url = notes.get("url") or url
                body.insert(tk.END, f"{title}\n{browser_url}\n\n")
                body.insert(tk.END, notes.get("body") or "No release notes body was provided.")
                body.config(state=tk.DISABLED)

            self.root.after(0, finish)

        threading.Thread(target=load_notes, daemon=True).start()

    def _github_repo_from_url(self, url):
        try:
            parsed = urlparse(url)
        except Exception:
            return None
        if "github.com" not in parsed.netloc.lower():
            return None
        parts = [part for part in parsed.path.strip("/").split("/") if part]
        if len(parts) < 2:
            return None
        return parts[0], parts[1].removesuffix(".git")

    def _load_github_release_notes(self, owner, repo):
        request = Request(
            f"https://api.github.com/repos/{owner}/{repo}/releases/latest",
            headers={
                "Accept": "application/vnd.github+json",
                "User-Agent": "WingetUpdateManager/3.0.0",
            },
        )
        with urlopen(request, timeout=15) as response:
            payload = json.loads(response.read().decode("utf-8"))
        return {
            "title": payload.get("name") or payload.get("tag_name") or f"{owner}/{repo}",
            "body": payload.get("body") or "",
            "url": payload.get("html_url") or f"https://github.com/{owner}/{repo}/releases/latest",
        }

    def hide_info_panel(self):
        self._animate_panel_width(self.info_panel, 0)
        
    def _animate_panel_width(self, panel, target, step=40):
        current = panel.winfo_width()
        if current == target:
            if target == 0:
                panel.pack_forget()
            return
            
        if current < target:
            new_w = min(current + step, target)
        else:
            new_w = max(current - step, target)
            
        panel.config(width=new_w)
        
        if new_w != target:
            self.root.after(15, lambda: self._animate_panel_width(panel, target, step))
        elif target == 0:
            panel.pack_forget()

    def _build_nav_item(self, parent, page_id, text, icon):
        outer = tk.Frame(parent, bg=self.colors["surface_soft"])
        outer.pack(fill=tk.X, pady=4)
        accent = tk.Frame(outer, bg=self.colors["surface_soft"], width=4, height=44)
        accent.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        accent.pack_propagate(False)

        frame = tk.Frame(outer, bg=self.colors["surface_soft"], pady=12, padx=14, cursor="hand2")
        frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        icon_lbl = tk.Label(frame, text=icon, font=("Segoe UI Symbol", 13),
                            bg=self.colors["surface_soft"], fg=self.colors["text_soft"])
        icon_lbl.pack(side=tk.LEFT, padx=(0, 12))
        text_lbl = tk.Label(frame, text=text, font=FONTS["nav"],
                            bg=self.colors["surface_soft"], fg=self.colors["text_dim"])
        text_lbl.pack(side=tk.LEFT)

        if page_id == "updates":
            self.update_badge_label = tk.Label(
                frame, text="", bg=self.colors["surface_soft"],
                fg=self.colors["nav_badge_fg"], font=FONTS["small_bold"],
                padx=8, pady=2)
            self.update_badge_label.pack(side=tk.RIGHT)

        for w in [outer, frame, accent, icon_lbl, text_lbl]:
            w.bind("<Button-1>", lambda e, p=page_id: self.switch_page(p))

        outer.body = frame
        outer.icon_label = icon_lbl
        outer.text_label = text_lbl
        outer.accent = accent
        self.nav_items[page_id] = outer
        self.nav_accents[page_id] = accent

    def switch_page(self, page_id):
        for frame in self.pages.values():
            frame.pack_forget()
        for pid, nav in self.nav_items.items():
            active = pid == page_id
            nav_bg = self.colors["surface_soft"]
            body_bg = self.colors["nav_active"] if active else self.colors["surface_soft"]
            fg = self.colors["primary"] if active else self.colors["text_dim"]
            icon_fg = self.colors["primary"] if active else self.colors["text_soft"]
            nav.configure(bg=nav_bg)
            nav.body.configure(bg=body_bg)
            nav.icon_label.configure(bg=body_bg, fg=icon_fg)
            nav.text_label.configure(bg=body_bg, fg=fg)
            nav.accent.configure(bg=self.colors["primary"] if active else nav_bg)

        if hasattr(self, "update_badge_label"):
            badge_text = self.update_badge_label.cget("text").strip()
            badge_parent = self.update_badge_label.master
            badge_parent_bg = badge_parent.cget("bg")
            self.update_badge_label.configure(
                bg=self.colors["nav_badge_bg"] if badge_text else badge_parent_bg,
                fg=self.colors["nav_badge_fg"],
            )
        if page_id in self.pages:
            self.pages[page_id].pack(fill=tk.BOTH, expand=True)
        self.current_page = page_id
        if page_id == "history":
            self._refresh_history()
        elif page_id == "dashboard":
            self._refresh_dashboard()
        elif page_id == "updates":
            self._animate_console_height(self.console_idle_height)
            self._refresh_admin_indicator()
            self._refresh_group_ui()

    def _build_page_header(self, page, title, subtitle):
        header = tk.Frame(page, bg=self.colors["surface"], padx=32, pady=22)
        header.pack(fill=tk.X)

        left = tk.Frame(header, bg=self.colors["surface"])
        left.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Label(left, text=title, bg=self.colors["surface"],
                 fg=self.colors["text_main"], font=FONTS["title"]).pack(anchor="w")
        subtitle_lbl = tk.Label(left, text=subtitle, bg=self.colors["surface"],
                                fg=self.colors["text_dim"], font=FONTS["small"])
        subtitle_lbl.pack(anchor="w", pady=(4, 0))

        actions = tk.Frame(header, bg=self.colors["surface"])
        actions.pack(side=tk.RIGHT, anchor="n")

        bell_wrap = tk.Frame(actions, bg=self.colors["surface"], width=28, height=24)
        bell_wrap.pack(side=tk.RIGHT, padx=(12, 0))
        bell_wrap.pack_propagate(False)
        bell = tk.Label(bell_wrap, text="\U0001F514", bg=self.colors["surface"],
                        fg=self.colors["text_dim"], font=("Segoe UI", 14), cursor="hand2")
        bell.pack(anchor="center")
        bell.bind("<Button-1>", lambda e: self._show_notification_center(e))
        badge = tk.Label(
            bell_wrap, text="", bg=self.colors["danger"], fg="white",
            font=("Segoe UI", 7, "bold"), padx=4, pady=0
        )
        badge.place_forget()
        bell_wrap.badge_label = badge
        bell_wrap.bind("<Button-1>", lambda e: self._show_notification_center(e))
        self.notification_badges.append(bell_wrap)
        self.root.after(0, self._refresh_notification_indicators)

        return subtitle_lbl, actions

    def _show_notification_center(self, event):
        if hasattr(self, '_notif_popup') and self._notif_popup and self._notif_popup.winfo_exists():
            self._notif_popup.destroy()
            self._notif_popup = None
            return
        w = event.widget
        x = w.winfo_rootx() - 260
        y = w.winfo_rooty() + 28
        popup = tk.Toplevel(self.root)
        popup.overrideredirect(True)
        popup.attributes('-topmost', True)
        popup.geometry(f"320x380+{x}+{y}")
        self._notif_popup = popup

        outer = tk.Frame(popup, bg=self.colors["border"], padx=1, pady=1)
        outer.pack(fill=tk.BOTH, expand=True)
        inner = tk.Frame(outer, bg=self.colors["card_bg"])
        inner.pack(fill=tk.BOTH, expand=True)

        hdr = tk.Frame(inner, bg=self.colors["surface_alt"], padx=14, pady=10)
        hdr.pack(fill=tk.X)
        tk.Label(hdr, text="\U0001F514 Notifications", bg=self.colors["surface_alt"],
                 fg=self.colors["text_main"], font=FONTS["body_bold"]).pack(side=tk.LEFT)
        clr = tk.Label(hdr, text="Clear All", bg=self.colors["surface_alt"],
                       fg=self.colors["primary"], font=FONTS["small"], cursor="hand2")
        clr.pack(side=tk.RIGHT)
        clr.bind("<Button-1>", lambda e: self._clear_notifications(popup))

        body = ScrollableFrame(inner, bg_color=self.colors["card_bg"])
        body.pack(fill=tk.BOTH, expand=True)
        content = body.scrollable_frame

        notifs = Toast._notifications
        if not notifs:
            tk.Label(content, text="No notifications yet", bg=self.colors["card_bg"],
                     fg=self.colors["text_dim"], font=FONTS["body"], pady=40).pack(fill=tk.X)
        else:
            type_colors = {"info": "#3b82f6", "success": "#10b981",
                           "error": "#ef4444", "warning": "#f59e0b"}
            for n in notifs:
                row = tk.Frame(content, bg=self.colors["card_bg"], padx=12, pady=8)
                row.pack(fill=tk.X, pady=1)
                tk.Label(row, text="\u25CF", bg=self.colors["card_bg"],
                         fg=type_colors.get(n["type"], "#3b82f6"),
                         font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=(0, 8))
                tk.Label(row, text=n["message"], bg=self.colors["card_bg"],
                         fg=self.colors["text_main"], font=FONTS["small"],
                         wraplength=220, justify="left", anchor="w").pack(side=tk.LEFT, fill=tk.X, expand=True)
                tk.Label(row, text=n["timestamp"], bg=self.colors["card_bg"],
                         fg=self.colors["text_soft"], font=FONTS["micro"]).pack(side=tk.RIGHT)

        popup.bind("<FocusOut>", lambda e: self._dismiss_notif_popup(popup))
        popup.focus_set()

    def _dismiss_notif_popup(self, popup):
        try:
            if popup.winfo_exists():
                self.root.after(200, lambda: self._safe_destroy_notif(popup))
        except Exception:
            pass

    def _safe_destroy_notif(self, popup):
        try:
            if popup.winfo_exists() and not popup.focus_get():
                popup.destroy()
                self._notif_popup = None
        except Exception:
            pass

    def _clear_notifications(self, popup):
        Toast._notifications.clear()
        self._refresh_notification_indicators()
        try:
            popup.destroy()
            self._notif_popup = None
        except Exception:
            pass
        Toast(self.root, "Notifications cleared", "info", record=False)

    def _refresh_notification_indicators(self):
        count = len(Toast._notifications)
        for wrap in list(self.notification_badges):
            try:
                badge = getattr(wrap, "badge_label", None)
                if badge is None:
                    continue
                if count:
                    badge.configure(text=str(min(count, 99)))
                    badge.place(relx=1.0, rely=0.0, x=-2, y=0, anchor="ne")
                else:
                    badge.place_forget()
            except Exception:
                continue

    def _record_notification(self, message, kind="info"):
        Toast._notifications.insert(0, {
            "message": message,
            "type": kind,
            "timestamp": datetime.now().strftime("%H:%M:%S"),
        })
        Toast._notifications = Toast._notifications[:50]
        self._refresh_notification_indicators()

    def _build_card(self, parent, pack_options=None):
        card = tk.Frame(parent, bg=self.colors["card_bg"], highlightthickness=1,
                        highlightbackground=self.colors["border"], bd=0)
        if pack_options is None:
            pack_options = {"fill": tk.BOTH, "expand": True}
        card.pack(**pack_options)
        return card

    def _build_empty_state(self, parent, title, detail, bg=None, pady=54):
        bg_color = bg or self.colors["card_bg"]
        wrap = tk.Frame(parent, bg=bg_color, pady=pady)
        wrap.pack(fill=tk.BOTH, expand=True)
        tk.Label(wrap, text="\u25CE", bg=bg_color, fg=self.colors["text_soft"],
                 font=("Segoe UI Symbol", 26)).pack()
        tk.Label(wrap, text=title, bg=bg_color, fg=self.colors["text_main"],
                 font=FONTS["sub_header"]).pack(pady=(10, 4))
        tk.Label(wrap, text=detail, bg=bg_color, fg=self.colors["text_dim"],
                 font=FONTS["body"]).pack()
        return wrap

    def _make_search_entry(self, parent, variable, placeholder):
        wrap = tk.Frame(parent, bg=self.colors["surface_alt"], padx=14, pady=10,
                        highlightthickness=1,
                        highlightbackground=self.colors["border_soft"])
        wrap.pack(fill=tk.X)
        entry = tk.Entry(
            wrap, textvariable=variable, bg=self.colors["surface_alt"],
            fg=self.colors["text_dim"], insertbackground=self.colors["text_main"],
            font=FONTS["body"], relief=tk.FLAT, highlightthickness=0, borderwidth=0
        )
        entry.pack(fill=tk.X, ipady=4)
        entry.insert(0, placeholder)
        entry.bind("<FocusIn>", lambda e: self._ph_in(entry, placeholder))
        entry.bind("<FocusOut>", lambda e: self._ph_out(entry, placeholder))
        return entry

    def _sanitize_update_groups(self, groups):
        cleaned = {}
        if not isinstance(groups, dict):
            return cleaned
        for group_name, items in groups.items():
            name = str(group_name or "").strip()
            if not name:
                continue
            if not isinstance(items, list):
                items = [items]
            package_ids = []
            seen = set()
            for item in items:
                pkg_id = str(item or "").strip()
                lowered = pkg_id.lower()
                if not pkg_id or lowered in seen:
                    continue
                seen.add(lowered)
                package_ids.append(pkg_id)
            cleaned[name] = package_ids
        return dict(sorted(cleaned.items(), key=lambda item: item[0].lower()))

    def _get_update_groups(self):
        groups = self.config.get("update_groups", {})
        cleaned = self._sanitize_update_groups(groups)
        if cleaned != groups:
            self.config.set("update_groups", cleaned)
        return cleaned

    def _set_update_groups(self, groups):
        cleaned = self._sanitize_update_groups(groups)
        self.config.set("update_groups", cleaned)
        self._refresh_group_ui()
        self._filter_updates()

    def _all_known_package_ids(self):
        package_ids = set()
        for source in [self.installed_list, self.updates_list]:
            for pkg in source:
                pkg_id = str(pkg.get("id") or "").strip()
                if pkg_id:
                    package_ids.add(pkg_id)
        for items in self._get_update_groups().values():
            package_ids.update(items)
        return sorted(package_ids, key=str.lower)

    def _group_filter_name(self):
        value = (self.group_filter_var.get() or "").strip()
        return value if value and value != "All Packages" else None

    def _group_package_ids(self, group_name=None):
        selected = group_name if group_name is not None else self._group_filter_name()
        if not selected:
            return set()
        return {pkg_id.lower() for pkg_id in self._get_update_groups().get(selected, [])}

    def _find_installed_package(self, package_id, manager=None):
        for pkg in self.installed_list:
            if pkg.get("id") != package_id:
                continue
            if manager and pkg.get("manager", "winget") != manager:
                continue
            return pkg
        return None

    def _refresh_group_ui(self):
        groups = self._get_update_groups()
        names = ["All Packages", *groups.keys()]
        current = self.group_filter_var.get() or "All Packages"
        if current not in names:
            current = "All Packages"
            self.group_filter_var.set(current)
        if hasattr(self, "group_filter_combo"):
            self.group_filter_combo.configure(values=names)
            self.group_filter_var.set(current)
        if hasattr(self, "group_package_combo"):
            self.group_package_combo.configure(values=self._all_known_package_ids())
        if hasattr(self, "group_listbox"):
            self.group_listbox.delete(0, tk.END)
            for name in groups:
                self.group_listbox.insert(tk.END, name)
            if self.selected_group_name not in groups:
                self.selected_group_name = next(iter(groups), None)
            if self.selected_group_name in groups:
                idx = list(groups.keys()).index(self.selected_group_name)
                self.group_listbox.selection_clear(0, tk.END)
                self.group_listbox.selection_set(idx)
                self.group_listbox.see(idx)
            self._populate_group_editor()
        if hasattr(self, "btn_update_group"):
            targets = self._group_targets(self._group_filter_name())
            enabled = bool(self._group_filter_name() and targets and not self.is_updating and not self.is_checking)
            self.btn_update_group.config(
                state=tk.NORMAL if enabled else tk.DISABLED,
                bg=self.colors["secondary"] if enabled else self.colors["surface_alt"],
            )

    def _group_targets(self, group_name):
        if not group_name:
            return []
        package_ids = self._group_package_ids(group_name)
        excluded = {pkg.lower() for pkg in self.config.get("excluded_packages", [])}
        targets = []
        for pkg in self.updates_list:
            pkg_id = str(pkg.get("id") or "").lower()
            if pkg_id in package_ids and pkg_id not in excluded:
                targets.append(pkg)
        return targets

    def _parse_labelled_output(self, output):
        data = {}
        current_key = None
        current_value = []
        for raw_line in output.splitlines():
            line = raw_line.rstrip()
            stripped = line.strip()
            if not stripped:
                continue
            if ":" in line and not line.startswith((" ", "\t")):
                key, value = line.split(":", 1)
                key = key.strip()
                if key:
                    if current_key:
                        data[current_key] = " ".join(current_value).strip()
                    current_key = key
                    current_value = [value.strip()] if value.strip() else []
                    continue
            if current_key:
                current_value.append(stripped)
        if current_key:
            data[current_key] = " ".join(current_value).strip()
        return data

    def _tail_message(self, output, fallback):
        for line in reversed((output or "").splitlines()):
            stripped = line.strip()
            if stripped:
                return stripped
        return fallback

    def _normalize_lookup(self, text):
        if not text:
            return ""
        return re.sub(r"[^a-z0-9]+", "", text.lower())

    def _read_registry_value(self, key, name):
        try:
            return winreg.QueryValueEx(key, name)[0]
        except OSError:
            return None

    def _format_registry_install_date(self, value):
        raw = str(value or "").strip()
        if len(raw) == 8 and raw.isdigit():
            return f"{raw[:4]}-{raw[4:6]}-{raw[6:]}"
        return raw or None

    def _format_registry_estimated_size(self, value):
        try:
            kb = int(str(value).strip())
        except Exception:
            return None
        if kb <= 0:
            return None
        mb = kb / 1024
        if mb >= 1024:
            return f"{mb / 1024:.2f} GB"
        return f"{mb:.1f} MB"

    def _load_registry_app_index(self):
        if self.registry_app_index is not None or winreg is None:
            return

        records = []
        seen = set()
        roots = [
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", winreg.KEY_WOW64_64KEY),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", winreg.KEY_WOW64_32KEY),
            (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", 0),
        ]

        for hive, path, view_flag in roots:
            flags = winreg.KEY_READ | view_flag
            try:
                with winreg.OpenKey(hive, path, 0, flags) as root:
                    count = winreg.QueryInfoKey(root)[0]
                    for idx in range(count):
                        try:
                            sub_name = winreg.EnumKey(root, idx)
                            with winreg.OpenKey(root, sub_name) as subkey:
                                display_name = self._read_registry_value(subkey, "DisplayName")
                                if not display_name:
                                    continue
                                display_icon = self._read_registry_value(subkey, "DisplayIcon") or ""
                                version = self._read_registry_value(subkey, "DisplayVersion") or ""
                                publisher = self._read_registry_value(subkey, "Publisher") or ""
                                dedupe_key = (display_name.lower(), version.lower(), display_icon.lower())
                                if dedupe_key in seen:
                                    continue
                                seen.add(dedupe_key)
                                record = {
                                    "display_name": display_name,
                                    "norm_name": self._normalize_lookup(display_name),
                                    "display_icon": display_icon,
                                    "version": version,
                                    "publisher": publisher,
                                    "install_date": self._format_registry_install_date(
                                        self._read_registry_value(subkey, "InstallDate")
                                    ),
                                    "estimated_size": self._format_registry_estimated_size(
                                        self._read_registry_value(subkey, "EstimatedSize")
                                    ),
                                }
                                records.append(record)
                                self.registry_apps_by_name.setdefault(record["norm_name"], []).append(record)
                        except OSError:
                            continue
            except OSError:
                continue

        self.registry_app_index = records

    def _parse_icon_location(self, value):
        if not value:
            return None

        raw = os.path.expandvars(str(value).strip())
        if not raw:
            return None

        path = None
        index = 0

        quoted = re.match(r'^\s*"([^"]+)"(?:\s*,\s*(-?\d+))?', raw)
        if quoted:
            path = quoted.group(1)
            index = int(quoted.group(2) or 0)
        else:
            match = re.match(
                r'^\s*([^,]+?\.(?:ico|exe|dll|png|jpg|jpeg|bmp|gif))(?:\s*,\s*(-?\d+))?',
                raw,
                re.IGNORECASE,
            )
            if match:
                path = match.group(1).strip().strip('"')
                index = int(match.group(2) or 0)
            else:
                path = raw.split(",")[0].strip().strip('"')

        if not path:
            return None

        path = os.path.expandvars(path)
        if not os.path.exists(path):
            return None

        return {"kind": "file", "path": path, "index": index}

    def _match_registry_record(self, pkg):
        self._load_registry_app_index()
        if not self.registry_app_index:
            return None

        norm_name = self._normalize_lookup(pkg.get("name"))
        if norm_name and norm_name in self.registry_apps_by_name:
            return self.registry_apps_by_name[norm_name][0]

        norm_id = self._normalize_lookup(pkg.get("id", "").split(".")[-1])
        candidates = []
        for record in self.registry_app_index:
            if not record.get("display_icon"):
                continue
            if norm_name and (norm_name in record["norm_name"] or record["norm_name"] in norm_name):
                candidates.append(record)
                continue
            if norm_id and norm_id and (norm_id in record["norm_name"] or record["norm_name"] in norm_id):
                candidates.append(record)

        if not candidates:
            return None

        candidates.sort(key=lambda rec: abs(len(rec["norm_name"]) - len(norm_name or norm_id)))
        return candidates[0]

    def _get_package_icon_ref(self, pkg):
        cache_key = f"{pkg.get('manager', 'winget')}::{pkg.get('id', '')}::{pkg.get('name', '')}"
        if cache_key in self.package_icon_refs:
            return self.package_icon_refs[cache_key]

        ref = None
        manager = pkg.get("manager", "winget")
        if manager == "npm":
            ref = {"kind": "npm"}
        else:
            record = self._match_registry_record(pkg)
            if record:
                ref = self._parse_icon_location(record.get("display_icon"))

        self.package_icon_refs[cache_key] = ref
        pkg["icon_ref"] = ref
        return ref

    def _hicon_to_pil(self, hicon):
        if not HAS_PIL or os.name != "nt":
            return None

        icon_info = ICONINFO()
        if not ctypes.windll.user32.GetIconInfo(hicon, ctypes.byref(icon_info)):
            return None

        color_bitmap = icon_info.hbmColor or icon_info.hbmMask
        if not color_bitmap:
            return None

        bmp = BITMAP()
        ctypes.windll.gdi32.GetObjectW(color_bitmap, ctypes.sizeof(BITMAP), ctypes.byref(bmp))
        width = bmp.bmWidth
        height = bmp.bmHeight if icon_info.hbmColor else bmp.bmHeight // 2
        if width <= 0 or height <= 0:
            return None

        bmi = BITMAPINFO()
        bmi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
        bmi.bmiHeader.biWidth = width
        bmi.bmiHeader.biHeight = -height
        bmi.bmiHeader.biPlanes = 1
        bmi.bmiHeader.biBitCount = 32
        bmi.bmiHeader.biCompression = 0

        buffer = ctypes.create_string_buffer(width * height * 4)
        hdc = ctypes.windll.user32.GetDC(None)
        try:
            bits = ctypes.windll.gdi32.GetDIBits(
                hdc, color_bitmap, 0, height, buffer, ctypes.byref(bmi), 0
            )
        finally:
            ctypes.windll.user32.ReleaseDC(None, hdc)
            if icon_info.hbmColor:
                ctypes.windll.gdi32.DeleteObject(icon_info.hbmColor)
            if icon_info.hbmMask:
                ctypes.windll.gdi32.DeleteObject(icon_info.hbmMask)

        if not bits:
            return None

        return Image.frombuffer(
            "RGBA", (width, height), buffer, "raw", "BGRA", 0, 1
        ).copy()

    def _load_icon_image(self, ref, size):
        if not HAS_PIL or not ref or ref.get("kind") != "file":
            return None

        path = ref["path"]
        suffix = Path(path).suffix.lower()
        try:
            if suffix in {".ico", ".png", ".jpg", ".jpeg", ".bmp", ".gif"}:
                img = Image.open(path).convert("RGBA")
            elif suffix in {".exe", ".dll"} and os.name == "nt":
                large = wintypes.HICON()
                small = wintypes.HICON()
                count = ctypes.windll.shell32.ExtractIconExW(
                    path, ref.get("index", 0), ctypes.byref(large), ctypes.byref(small), 1
                )
                if count <= 0:
                    return None
                handles = [h for h in [small.value, large.value] if h]
                if not handles:
                    return None
                try:
                    img = self._hicon_to_pil(handles[0])
                finally:
                    for handle in set(handles):
                        ctypes.windll.user32.DestroyIcon(handle)
            else:
                return None
        except Exception:
            return None

        if img is None:
            return None
        return img.resize((size, size), Image.LANCZOS)

    def _get_package_photo(self, pkg, size=28):
        if not HAS_PIL:
            return None

        ref = pkg.get("icon_ref") if isinstance(pkg, dict) else None
        if ref is None and isinstance(pkg, dict):
            ref = self._get_package_icon_ref(pkg)

        if not ref:
            return None

        ref_key = json.dumps(ref, sort_keys=True)
        cache_key = f"{ref_key}:{size}"
        if cache_key in self.tk_icon_cache:
            return self.tk_icon_cache[cache_key]

        img = self._load_icon_image(ref, size)
        if img is None:
            return None

        photo = ImageTk.PhotoImage(img)
        self.tk_icon_cache[cache_key] = photo
        return photo

    def _build_package_icon_widget(self, parent, pkg, bg, size=28):
        photo = self._get_package_photo(pkg, size=size)
        if photo is not None:
            lbl = tk.Label(parent, image=photo, bg=bg)
            lbl.image = photo
            return lbl

        if pkg.get("manager") == "npm":
            text = "npm"
            tile_bg = "#cb3837"
            tile_fg = "white"
        else:
            initials = (pkg.get("name") or "??")[:2].upper()
            text = initials
            palette = [
                ("#fca5a5", "#991b1b"), ("#fdba74", "#9a3412"),
                ("#fcd34d", "#854d0e"), ("#86efac", "#166534"),
                ("#93c5fd", "#1e40af"), ("#c4b5fd", "#4c1d95"),
                ("#f0abfc", "#86198f"),
            ]
            idx = sum(ord(c) for c in text) % len(palette)
            tile_bg, tile_fg = palette[idx]

        lbl = tk.Label(parent, text=text, bg=tile_bg, fg=tile_fg,
                       font=("Segoe UI", 8, "bold"), width=4, height=2)
        lbl.preserve_bg = True
        return lbl

    def _npm_command(self):
        return "npm.cmd" if os.name == "nt" else "npm"

    def _load_npm_globals(self):
        try:
            proc = subprocess.run(
                [self._npm_command(), "list", "-g", "--depth=0", "--json"],
                capture_output=True, text=True, encoding="utf-8",
                errors="ignore", timeout=30,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            if proc.returncode not in (0, 1):
                return []
            data = json.loads(proc.stdout or "{}")
            deps = data.get("dependencies", {}) or {}
            packages = []
            for pkg_name, meta in deps.items():
                packages.append({
                    "name": pkg_name,
                    "id": pkg_name,
                    "version": meta.get("version", ""),
                    "source": "npm-global",
                    "manager": "npm",
                })
            packages.sort(key=lambda item: item["name"].lower())
            return packages
        except Exception:
            return []

    def _load_npm_updates(self):
        try:
            proc = subprocess.run(
                [self._npm_command(), "outdated", "-g", "--depth=0", "--json"],
                capture_output=True, text=True, encoding="utf-8",
                errors="ignore", timeout=60,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            if proc.returncode not in (0, 1):
                return []
            data = json.loads(proc.stdout or "{}")
            if not isinstance(data, dict):
                return []
            packages = []
            for pkg_name, meta in data.items():
                if not isinstance(meta, dict):
                    continue
                current = meta.get("current", "")
                available = meta.get("latest") or meta.get("wanted") or ""
                if not available or available == current:
                    continue
                packages.append({
                    "name": pkg_name,
                    "id": pkg_name,
                    "version": current,
                    "available": available,
                    "source": "npm-global",
                    "manager": "npm",
                })
            packages.sort(key=lambda item: item["name"].lower())
            return packages
        except Exception:
            return []

    def _build_dashboard_page(self):
        page = tk.Frame(self.main_container, bg=self.colors["bg"])
        self.pages["dashboard"] = page

        self._build_page_header(page, "Dashboard", "\u25CF Overview of packages, activity, and scan health")

        cards = tk.Frame(page, bg=self.colors["bg"], padx=SPACING["page_x"], pady=SPACING["page_y"])
        cards.pack(fill=tk.X)

        self.dash_cards = {}
        for key, label, value, color in [
            ("installed", "Installed Packages", "---", self.colors["primary"]),
            ("updates", "Pending Updates", "---", self.colors["warning"]),
            ("last_scan", "Last Scan", "Never", self.colors["text_dim"]),
            ("total_updated", "Total Updated",
             str(len([e for e in self.history.entries if e["status"] == "success"])),
             self.colors["secondary"]),
        ]:
            card = tk.Frame(cards, bg=self.colors["card_bg"], padx=20, pady=20,
                            highlightthickness=1, highlightbackground=self.colors["border_soft"])
            card.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
            tk.Label(card, text=label, bg=self.colors["card_bg"],
                     fg=self.colors["text_soft"], font=FONTS["small"]).pack(anchor="w")
            lbl = tk.Label(card, text=value, bg=self.colors["card_bg"],
                           fg=color, font=("Segoe UI", 24, "bold"))
            lbl.pack(anchor="w", pady=(5, 0))
            self.dash_cards[key] = lbl

        self.dashboard_cache_label = tk.Label(
            page, text="", bg=self.colors["bg"], fg=self.colors["text_dim"],
            font=FONTS["small"]
        )
        self.dashboard_cache_label.pack(anchor="w", padx=SPACING["page_x"], pady=(0, 10))

        # -- Health Score card --
        health_row = tk.Frame(page, bg=self.colors["bg"], padx=SPACING["page_x"])
        health_row.pack(fill=tk.X, pady=(0, 5))
        health_card = tk.Frame(health_row, bg=self.colors["card_bg"], padx=24, pady=16,
                               highlightthickness=1, highlightbackground=self.colors["border_soft"])
        health_card.pack(fill=tk.X)
        tk.Label(health_card, text="System Health Score", bg=self.colors["card_bg"],
                 fg=self.colors["text_soft"], font=FONTS["small"]).pack(anchor="w")
        self.health_canvas = tk.Canvas(health_card, width=320, height=28,
                                       bg=self.colors["card_bg"], highlightthickness=0)
        self.health_canvas.pack(anchor="w", pady=(6, 2))
        self.health_label = tk.Label(health_card, text="---", bg=self.colors["card_bg"],
                                     fg=self.colors["primary"], font=FONTS["body_bold"])
        self.health_label.pack(anchor="w")

        # -- Charts row --
        charts_row = tk.Frame(page, bg=self.colors["bg"], padx=SPACING["page_x"])
        charts_row.pack(fill=tk.X, pady=(5, 5))

        # Bar chart card
        bar_card = tk.Frame(charts_row, bg=self.colors["card_bg"], padx=18, pady=14,
                            highlightthickness=1, highlightbackground=self.colors["border_soft"])
        bar_card.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        tk.Label(bar_card, text="Updates (Last 7 Days)", bg=self.colors["card_bg"],
                 fg=self.colors["text_soft"], font=FONTS["small"]).pack(anchor="w")
        self.bar_chart_canvas = tk.Canvas(bar_card, width=300, height=120,
                                          bg=self.colors["card_bg"], highlightthickness=0)
        self.bar_chart_canvas.pack(fill=tk.X, pady=(8, 0))

        # Pie chart card
        pie_card = tk.Frame(charts_row, bg=self.colors["card_bg"], padx=18, pady=14,
                            highlightthickness=1, highlightbackground=self.colors["border_soft"])
        pie_card.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0))
        tk.Label(pie_card, text="Success / Failure Ratio", bg=self.colors["card_bg"],
                 fg=self.colors["text_soft"], font=FONTS["small"]).pack(anchor="w")
        self.pie_chart_canvas = tk.Canvas(pie_card, width=140, height=120,
                                          bg=self.colors["card_bg"], highlightthickness=0)
        self.pie_chart_canvas.pack(pady=(8, 0))

        # -- Recent Activity --
        tk.Label(page, text="Recent Activity", bg=self.colors["bg"],
                 fg=self.colors["text_main"], font=FONTS["sub_header"]).pack(
                     anchor="w", padx=SPACING["page_x"], pady=(10, 5))

        act_container = tk.Frame(page, bg=self.colors["bg"], padx=SPACING["page_x"])
        act_container.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        self.dash_activity_scroll = ScrollableFrame(act_container, bg_color=self.colors["card_bg"])
        self.dash_activity_scroll.pack(fill=tk.BOTH, expand=True)
        self.dash_activity_content = self.dash_activity_scroll.scrollable_frame

    def _refresh_dashboard(self):
        self.dash_cards["updates"].config(text=str(self.update_count))
        self.dash_cards["total_updated"].config(
            text=str(len([e for e in self.history.entries if e["status"] == "success"])))
        if self.installed_list:
            self.dash_cards["installed"].config(text=str(len(self.installed_list)))
        self.dash_cards["last_scan"].config(text=self.scan_cache.last_scan_time())
        age_seconds = int(self.scan_cache.age_seconds()) if self.scan_cache.age_seconds() != float("inf") else None
        is_stale = self.scan_cache.is_stale(self.config.get("cache_ttl", 3600))
        if age_seconds is None:
            cache_text = "Scan cache status: no cached scan available"
            cache_color = self.colors["text_dim"]
        else:
            age_minutes = max(1, age_seconds // 60) if age_seconds >= 60 else age_seconds
            unit = "m" if age_seconds >= 60 else "s"
            cache_text = f"Scan cache is {'stale' if is_stale else 'fresh'} • age {age_minutes}{unit}"
            cache_color = self.colors["warning"] if is_stale else self.colors["secondary"]
        self.dashboard_cache_label.config(text=cache_text, fg=cache_color)

        # Draw health score
        self._draw_health_score()
        # Draw charts
        self._draw_bar_chart()
        self._draw_pie_chart()

        for w in self.dash_activity_content.winfo_children():
            w.destroy()
        entries = self.history.get_entries(20)
        if not entries:
            tk.Label(self.dash_activity_content, text="No activity yet. Run an update scan!",
                     bg=self.colors["card_bg"], fg=self.colors["text_dim"],
                     font=FONTS["body"], pady=30).pack(fill=tk.X)
            return
        for entry in entries:
            row = tk.Frame(self.dash_activity_content, bg=self.colors["card_bg"], pady=8, padx=15)
            row.pack(fill=tk.X, pady=1)
            sc = {"success": self.colors["secondary"],
                  "failed": self.colors["danger"],
                  "skipped": self.colors["warning"]}
            tk.Label(row, text="\u25CF", bg=self.colors["card_bg"],
                     fg=sc.get(entry["status"], self.colors["text_dim"]),
                     font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=(0, 8))
            tk.Label(row, text=entry["package_name"], bg=self.colors["card_bg"],
                     fg=self.colors["text_main"], font=FONTS["body_bold"]).pack(side=tk.LEFT)
            tk.Label(row, text=f'{entry["old_version"]} \u2192 {entry["new_version"]}',
                     bg=self.colors["card_bg"], fg=self.colors["text_dim"],
                     font=FONTS["mono_small"]).pack(side=tk.LEFT, padx=(10, 0))
            try:
                dt = datetime.fromisoformat(entry["timestamp"])
                ts = dt.strftime("%b %d, %H:%M")
            except Exception:
                ts = entry.get("timestamp", "")[:16]
            tk.Label(row, text=ts, bg=self.colors["card_bg"],
                     fg=self.colors["text_dim"], font=FONTS["small"]).pack(side=tk.RIGHT)

    def _draw_health_score(self):
        """Draw a health score gauge bar on the dashboard."""
        c = self.health_canvas
        c.delete("all")
        w, h = 300, 22
        total_installed = len(self.installed_list)
        pending = self.update_count if hasattr(self, 'update_count') else 0
        if total_installed <= 0:
            score = 100
        else:
            score = int(max(0, min(1, 1 - (pending / total_installed))) * 100)
        # Background track
        c.create_rectangle(0, 4, w, h, fill=self.colors["surface_alt"], outline="")
        # Filled bar
        bar_color = self.colors["secondary"] if score >= 90 else (
            self.colors["warning"] if score >= 70 else self.colors["danger"])
        bar_w = int((score / 100) * w)
        c.create_rectangle(0, 4, bar_w, h, fill=bar_color, outline="")
        # Score text
        c.create_text(bar_w + 8, 13, text=f"{score}%", fill=self.colors["text_main"],
                      font=("Segoe UI", 9, "bold"), anchor="w")
        grade = "A+" if score >= 97 else "A" if score >= 93 else "B" if score >= 85 else (
            "C" if score >= 75 else "D" if score >= 60 else "F")
        self.health_label.config(text=f"Grade: {grade}  •  {score}% up to date",
                                 fg=bar_color)

    def _draw_bar_chart(self):
        """Draw a 7-day update frequency bar chart."""
        c = self.bar_chart_canvas
        c.delete("all")
        c.update_idletasks()
        cw = max(c.winfo_width(), 280)
        ch = 110
        from collections import Counter
        today = datetime.now().date()
        day_counts = Counter()
        for e in self.history.entries:
            try:
                d = datetime.fromisoformat(e["timestamp"]).date()
                delta = (today - d).days
                if delta < 7:
                    day_counts[delta] += 1
            except Exception:
                pass
        max_val = max(day_counts.values(), default=1) or 1
        bar_w = max(int((cw - 40) / 7) - 6, 12)
        for i in range(7):
            x = 20 + i * (bar_w + 6)
            count = day_counts.get(6 - i, 0)
            bar_h = int((count / max_val) * 80) if max_val else 0
            # Bar
            c.create_rectangle(x, ch - bar_h - 18, x + bar_w, ch - 18,
                               fill=self.colors["primary"], outline="")
            # Count label
            if count > 0:
                c.create_text(x + bar_w // 2, ch - bar_h - 22,
                              text=str(count), fill=self.colors["text_soft"],
                              font=("Segoe UI", 7), anchor="s")
            # Day label
            from datetime import timedelta
            day = today - timedelta(days=6 - i)
            c.create_text(x + bar_w // 2, ch - 4,
                          text=day.strftime("%a"), fill=self.colors["text_dim"],
                          font=("Segoe UI", 7))

    def _draw_pie_chart(self):
        """Draw a success/fail/skip pie chart."""
        c = self.pie_chart_canvas
        c.delete("all")
        total = len(self.history.entries)
        if total == 0:
            c.create_oval(20, 10, 120, 110, fill=self.colors["surface_alt"], outline="")
            c.create_text(70, 60, text="No data", fill=self.colors["text_dim"],
                          font=("Segoe UI", 8))
            return
        counts = {
            "success": len([e for e in self.history.entries if e["status"] == "success"]),
            "failed": len([e for e in self.history.entries if e["status"] == "failed"]),
            "skipped": len([e for e in self.history.entries if e["status"] == "skipped"]),
        }
        colors = {
            "success": self.colors["secondary"],
            "failed": self.colors["danger"],
            "skipped": self.colors["warning"],
        }
        start_angle = 90
        for key in ["success", "failed", "skipped"]:
            count = counts[key]
            if count <= 0:
                continue
            extent = -(count / total) * 360
            c.create_arc(20, 10, 120, 110, start=start_angle, extent=extent,
                         fill=colors[key], outline="")
            start_angle += extent
        legend_items = [("success", "OK"), ("failed", "Fail"), ("skipped", "Skip")]
        for idx, (key, label) in enumerate(legend_items):
            y = 28 + (idx * 22)
            c.create_rectangle(130, y, 142, y + 12, fill=colors[key], outline="")
            c.create_text(148, y + 6, text=f"{label} ({counts[key]})",
                          fill=self.colors["text_main"], font=("Segoe UI", 8), anchor="w")

    # --------------------------------------------------------- installed apps
    def _build_installed_page(self):
        page = tk.Frame(self.main_container, bg=self.colors["bg"])
        self.pages["installed"] = page

        self.installed_status, actions = self._build_page_header(
            page, "Installed Applications", "\u25CF Load your installed package list"
        )

        self.btn_import = self._btn(
            actions, "Import", self.colors["surface_alt"], self.colors["text_main"],
            self.import_installed, outline=True
        )
        self.btn_import.pack(side=tk.RIGHT)

        self.btn_export = self._btn(
            actions, "Export", self.colors["surface_alt"], self.colors["text_main"],
            self.export_installed, outline=True
        )
        self.btn_export.pack(side=tk.RIGHT, padx=(0, 10))

        self.btn_load_installed = self._btn(
            actions, "Load Installed", self.colors["primary"], "white",
            self.load_installed
        )
        self.btn_load_installed.pack(side=tk.RIGHT, padx=(0, 10))

        self.btn_export_script = self._btn(
            actions, "\U0001F4C4 Export Script", self.colors["surface_alt"],
            self.colors["text_main"], self._export_as_script, outline=True
        )
        self.btn_export_script.pack(side=tk.RIGHT, padx=(0, 10))

        body = tk.Frame(page, bg=self.colors["bg"], padx=SPACING["page_x"],
                        pady=SPACING["page_y"])
        body.pack(fill=tk.BOTH, expand=True)

        list_card = self._build_card(body)
        top = tk.Frame(list_card, bg=self.colors["card_bg"], padx=SPACING["card_pad"],
                       pady=18)
        top.pack(fill=tk.X)
        self._make_search_entry(top, self.installed_search_var, "Search installed apps...")
        self.installed_search_var.trace_add("write", lambda *a: self._filter_installed())

        col_h = tk.Frame(list_card, bg=self.colors["surface_alt"], padx=SPACING["card_pad"],
                         pady=14)
        col_h.pack(fill=tk.X)
        col_h.columnconfigure(0, weight=0, minsize=56)
        col_h.columnconfigure(1, weight=3, minsize=220)
        col_h.columnconfigure(2, weight=3, minsize=220)
        col_h.columnconfigure(3, weight=1, minsize=110)
        col_h.columnconfigure(4, weight=1, minsize=100)
        for i, t in enumerate(["", "APPLICATION NAME", "PACKAGE ID", "VERSION", "SOURCE"]):
            tk.Label(col_h, text=t, bg=self.colors["surface_alt"], fg=self.colors["text_dim"],
                     font=FONTS["small_bold"], anchor="w").grid(
                         row=0, column=i, sticky="w", padx=4)

        lc = tk.Frame(list_card, bg=self.colors["card_bg"], padx=18, pady=10)
        lc.pack(fill=tk.BOTH, expand=True)
        self.installed_scroll = ScrollableFrame(lc, bg_color=self.colors["card_bg"])
        self.installed_scroll.pack(fill=tk.BOTH, expand=True)
        self.installed_content = self.installed_scroll.scrollable_frame
        self._build_empty_state(
            self.installed_content, "No installed packages loaded",
            "Use Load Installed to build your current app inventory.",
            bg=self.colors["card_bg"], pady=62
        )

    def load_installed(self):
        self.btn_load_installed.config(state=tk.DISABLED)
        self.installed_status.config(text="\u25CF Scanning installed package inventory...", fg=self.colors["warning"])
        for w in self.installed_content.winfo_children():
            w.destroy()
        self._build_empty_state(
            self.installed_content, "Loading installed packages...",
            "This can take a moment on systems with a large app inventory.",
            bg=self.colors["card_bg"], pady=56
        )

        def scan():
            try:
                proc = subprocess.run(
                    ["winget", "list", "--disable-interactivity",
                     "--accept-source-agreements"],
                    capture_output=True, text=True, encoding='utf-8',
                    errors='ignore', timeout=120,
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
                winget_packages = self.parser.parse_list_output(proc.stdout)
                npm_packages = self._load_npm_globals()
                combined = winget_packages + npm_packages
                deduped = []
                seen = set()
                for pkg in combined:
                    dedupe_key = (pkg.get("manager", "winget"), pkg.get("id", "").lower())
                    if dedupe_key in seen:
                        continue
                    seen.add(dedupe_key)
                    pkg["icon_ref"] = self._get_package_icon_ref(pkg)
                    deduped.append(pkg)
                self.installed_list = sorted(
                    deduped, key=lambda item: (item.get("name") or "").lower()
                )
                self.scan_cache.save(installed=[{k: v for k, v in p.items() if k != 'icon_ref'} for p in self.installed_list])
                self.root.after(0, self._render_installed)
            except Exception as e:
                self.root.after(0, lambda: self._installed_error(str(e)))
        threading.Thread(target=scan, daemon=True).start()

    def _render_installed(self):
        self.installed_status.config(
            text=f"\u25CF {len(self.installed_list)} packages ready to browse",
            fg=self.colors["success"])
        self.btn_load_installed.config(state=tk.NORMAL)
        self._refresh_group_ui()
        self._filter_installed()

    def _filter_installed(self):
        q = self.installed_search_var.get().lower()
        if q in ("", "search installed apps..."):
            q = ""
        for w in self.installed_content.winfo_children():
            w.destroy()
        filtered = [p for p in self.installed_list
                    if not q or q in p['name'].lower() or q in p['id'].lower()]
        if not filtered and self.installed_list:
            self._build_empty_state(
                self.installed_content, "No matching packages",
                "Try a different search term to filter the installed app list.",
                bg=self.colors["card_bg"], pady=48
            )
            return
        if not filtered:
            self._build_empty_state(
                self.installed_content, "No installed packages loaded",
                "Use Load Installed to build your current app inventory.",
                bg=self.colors["card_bg"], pady=62
            )
            return
        self._install_render_id = getattr(self, '_install_render_id', 0) + 1
        current_id = self._install_render_id

        def _build_installed_row(pkg):
            row = tk.Frame(
                self.installed_content, bg=self.colors["card_inner"], pady=SPACING["row_pad_y"],
                padx=18, cursor="hand2", highlightthickness=1,
                highlightbackground=self.colors["border_soft"]
            )
            row.pack(fill=tk.X, pady=4)
            row.columnconfigure(0, weight=0, minsize=56)
            row.columnconfigure(1, weight=3, minsize=220)
            row.columnconfigure(2, weight=3, minsize=220)
            row.columnconfigure(3, weight=1, minsize=110)
            row.columnconfigure(4, weight=1, minsize=100)

            icon_lbl = self._build_package_icon_widget(row, pkg, self.colors["card_inner"], size=28)
            icon_lbl.grid(row=0, column=0, sticky="w", padx=4)
            name_wrap = tk.Frame(row, bg=self.colors["card_inner"])
            name_wrap.grid(row=0, column=1, sticky="w", padx=4)
            tk.Label(name_wrap, text=pkg['name'], bg=self.colors["card_inner"],
                     fg=self.colors["text_main"], font=FONTS["body_bold"],
                     anchor="w").pack(anchor="w")
            secondary = "Global npm package" if pkg.get("manager") == "npm" else (pkg.get('source') or "Installed package")
            tk.Label(name_wrap, text=secondary, bg=self.colors["card_inner"],
                     fg=self.colors["text_soft"], font=FONTS["small"],
                     anchor="w").pack(anchor="w", pady=(3, 0))
            tk.Label(row, text=pkg['id'], bg=self.colors["card_inner"],
                     fg=self.colors["text_dim"], font=FONTS["mono_small"],
                     anchor="w").grid(row=0, column=2, sticky="w", padx=4)
            tk.Label(row, text=pkg['version'], bg=self.colors["card_inner"],
                     fg=self.colors["text_dim"], font=FONTS["mono_small"],
                     anchor="w").grid(row=0, column=3, sticky="w", padx=4)
            tk.Label(row, text=pkg.get('source', '') or "Unknown", bg=self.colors["card_inner"],
                     fg=self.colors["text_dim"], font=FONTS["small"],
                     anchor="w").grid(row=0, column=4, sticky="w", padx=4)

            def _bind_click(widget, p_id, p_name, p_manager):
                widget.bind(
                    "<Button-1>",
                    lambda e, pid=p_id, pname=p_name, pm=p_manager: self.show_info_panel(pid, pname, pm),
                )
                for child in widget.winfo_children():
                    _bind_click(child, p_id, p_name, p_manager)

            def _bind_context(widget, package):
                widget.bind("<Button-3>", lambda e, p=package: self._post_installed_context_menu(e, p), add="+")
                for child in widget.winfo_children():
                    _bind_context(child, package)
                     
            _bind_click(row, pkg['id'], pkg['name'], pkg.get("manager", "winget"))
            _bind_context(row, pkg)

            def _set_bg(widget, bg):
                try:
                    if getattr(widget, "preserve_bg", False):
                        pass
                    else:
                        widget.config(bg=bg)
                except Exception:
                    pass
                for child in widget.winfo_children():
                    _set_bg(child, bg)

            row.bind("<Enter>", lambda e, r=row: _set_bg(r, self.colors["surface_hover"]))
            row.bind("<Leave>", lambda e, r=row: _set_bg(r, self.colors["card_inner"]))
             
        def _pack_delayed(items, idx=0):
            if getattr(self, '_install_render_id', 0) != current_id: return
            batch_size = max(1, len(items) // 20)
            end_idx = min(idx + batch_size, len(items))
            for i in range(idx, end_idx):
                _build_installed_row(items[i])
            if end_idx < len(items):
                self.root.after(10, lambda: _pack_delayed(items, end_idx))
                
        _pack_delayed(filtered)

    def _installed_error(self, msg):
        self.installed_status.config(text=f"\u25CF Error loading installed packages: {msg}",
                                     fg=self.colors["danger"])
        self.btn_load_installed.config(state=tk.NORMAL)

    def export_installed(self):
        if not self.installed_list:
            messagebox.showinfo("Export", "Please load installed packages first.")
            return
        
        filepath = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="Export Installed Packages"
        )
        if not filepath:
            return
        
        try:
            with open(filepath, 'w') as f:
                json.dump(self.installed_list, f, indent=2)
            Toast(self.root, f"Exported {len(self.installed_list)} packages", "success")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export: {str(e)}")

    def import_installed(self):
        filepath = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="Import Packages"
        )
        if not filepath:
            return
            
        try:
            with open(filepath, 'r') as f:
                packages = json.load(f)
            
            if not isinstance(packages, list) or not all(isinstance(p, dict) and 'id' in p for p in packages):
                raise ValueError("Invalid format. Expected a list of package objects with 'id'.")
            
            self.switch_page("updates")
            self.log(f"Starting import of {len(packages)} packages...", self.colors["primary"])
            
            def do_import():
                for i, pkg in enumerate(packages):
                    pkg_id = pkg['id']
                    manager = (pkg.get("manager") or "winget").lower()
                    version = (pkg.get("version") or "").strip()
                    install_target = pkg_id
                    self.root.after(
                        0,
                        lambda p=install_target, m=manager, idx=i + 1, tot=len(packages):
                            self.log(
                                f"[{idx}/{tot}] Installing {p} via {m}...",
                                self.colors["text_main"],
                            ),
                    )

                    if manager == "npm":
                        npm_target = f"{pkg_id}@{version}" if version else pkg_id
                        proc = subprocess.run(
                            [self._npm_command(), "install", "-g", npm_target],
                            capture_output=True, text=True, encoding='utf-8', errors='ignore',
                            creationflags=subprocess.CREATE_NO_WINDOW
                        )
                    else:
                        proc = subprocess.run(
                            ["winget", "install", "--id", pkg_id, "--exact",
                             "--accept-source-agreements", "--accept-package-agreements",
                             "--silent"],
                            capture_output=True, text=True, encoding='utf-8', errors='ignore',
                            creationflags=subprocess.CREATE_NO_WINDOW
                        )
                    
                    if proc.returncode == 0:
                        self.root.after(0, lambda: self.log("Success", self.colors["success"]))
                    else:
                        err_msg = proc.stderr.strip() or f"Exit code {proc.returncode}"
                        self.root.after(0, lambda e=err_msg: self.log(f"Failed: {e}", self.colors["danger"]))
                
                self.root.after(0, lambda: self.log("Import complete.", self.colors["success"]))
                self.root.after(0, lambda: Toast(self.root, "Import completed", "success"))
                
            threading.Thread(target=do_import, daemon=True).start()
            
        except Exception as e:
            messagebox.showerror("Import Error", f"Failed to import: {str(e)}")


    # -------------------------------------------------------------- updates
    def _build_updates_page(self):
        page = tk.Frame(self.main_container, bg=self.colors["bg"])
        self.pages["updates"] = page

        self.status_sub, h_right = self._build_page_header(
            page, "Updates Available", "\u25CF Ready to scan for pending upgrades"
        )

        self.btn_check = self._btn(
            h_right, "Check Updates", self.colors["surface_alt"],
            self.colors["text_main"], self.check_updates, outline=True
        )
        self.btn_check.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_cancel = self._btn(
            h_right, "Cancel", self.colors["danger"], "white",
            self.cancel_operation
        )

        self.btn_update_sel = self._btn(
            h_right, "Update Selected", self.colors["accent"], "white",
            self.update_selected
        )
        self.btn_update_sel.pack(side=tk.LEFT, padx=(0, 10))
        self.btn_update_sel.config(state=tk.DISABLED, bg=self.colors["surface_alt"])

        self.btn_update_group = self._btn(
            h_right, "Update Group", self.colors["surface_alt"], self.colors["text_main"],
            self.update_group
        )
        self.btn_update_group.pack(side=tk.LEFT, padx=(0, 10))
        self.btn_update_group.config(state=tk.DISABLED, bg=self.colors["surface_alt"])

        self.btn_update_all = self._btn(
            h_right, "Update All", self.colors["secondary"], "white",
            self.update_all
        )
        self.btn_update_all.pack(side=tk.LEFT)

        self.admin_badge = tk.Label(
            h_right, text="", bg=self.colors["surface"], fg=self.colors["warning"],
            font=FONTS["small_bold"]
        )
        self.admin_badge.pack(side=tk.LEFT, padx=(12, 0))
        self._refresh_admin_indicator()

        self.progress_var = tk.DoubleVar(value=0)
        style = ttk.Style()
        style.configure("Custom.Horizontal.TProgressbar",
                        troughcolor=self.colors["surface_alt"],
                        background=self.colors["primary"],
                        bordercolor=self.colors["border"],
                        lightcolor=self.colors["primary"],
                        darkcolor=self.colors["primary"])
        self.progress_bar = ttk.Progressbar(
            page, variable=self.progress_var, maximum=100,
            style="Custom.Horizontal.TProgressbar", mode="determinate")

        body = tk.Frame(page, bg=self.colors["bg"], padx=SPACING["page_x"],
                        pady=SPACING["page_y"])
        body.pack(fill=tk.BOTH, expand=True)

        list_card = self._build_card(
            body, {"side": tk.TOP, "fill": tk.BOTH, "expand": True, "pady": (0, SPACING["card_gap"])}
        )

        top = tk.Frame(list_card, bg=self.colors["card_bg"], padx=SPACING["card_pad"],
                       pady=18)
        top.pack(fill=tk.X)
        search_left = tk.Frame(top, bg=self.colors["card_bg"])
        search_left.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._make_search_entry(search_left, self.search_var, "Search updates...")
        self.search_var.trace_add("write", lambda *a: self._filter_updates())

        group_right = tk.Frame(top, bg=self.colors["card_bg"])
        group_right.pack(side=tk.RIGHT, padx=(12, 0))
        tk.Label(
            group_right, text="Group Filter", bg=self.colors["card_bg"],
            fg=self.colors["text_soft"], font=FONTS["small_bold"]
        ).pack(anchor="w")
        self.group_filter_combo = ttk.Combobox(
            group_right, textvariable=self.group_filter_var, state="readonly", width=20
        )
        self.group_filter_combo.pack(anchor="e", pady=(6, 0))
        self.group_filter_combo.bind("<<ComboboxSelected>>", lambda e: self._on_group_filter_changed())
        self._refresh_group_ui()

        col_h = tk.Frame(list_card, bg=self.colors["surface_alt"], padx=SPACING["card_pad"],
                         pady=14)
        col_h.pack(fill=tk.X)
        col_h.columnconfigure(0, weight=0, minsize=44)
        col_h.columnconfigure(1, weight=0, minsize=58)
        col_h.columnconfigure(2, weight=3, minsize=220)
        col_h.columnconfigure(3, weight=2, minsize=180)
        col_h.columnconfigure(4, weight=1, minsize=110)
        col_h.columnconfigure(5, weight=1, minsize=120)

        self.select_all_lbl = tk.Label(
            col_h, text="\u2610", bg=self.colors["surface_alt"],
            fg=self.colors["text_dim"], font=("Segoe UI Symbol", 13),
            cursor="hand2"
        )
        self.select_all_lbl.grid(row=0, column=0, sticky="w", padx=4)
        self.select_all_lbl.bind("<Button-1>", lambda e: self.toggle_all_selection())

        tk.Label(col_h, text="", bg=self.colors["surface_alt"]).grid(row=0, column=1)

        for i, (col_key, text) in enumerate([
            ("name", "Application Name"), ("id", "Package ID"),
            ("version", "Current"), ("available", "Latest"),
        ]):
            lbl = tk.Label(
                col_h, text=f"{text.upper()} \u2195", bg=self.colors["surface_alt"],
                fg=self.colors["text_dim"], font=FONTS["small_bold"],
                anchor="w", cursor="hand2"
            )
            lbl.grid(row=0, column=i + 2, sticky="w", padx=4)
            lbl.bind("<Button-1>", lambda e, k=col_key: self._sort_updates(k))

        lc = tk.Frame(list_card, bg=self.colors["card_bg"], padx=18, pady=10)
        lc.pack(fill=tk.BOTH, expand=True)
        self.scroll_frame = ScrollableFrame(lc, bg_color=self.colors["card_bg"])
        self.scroll_frame.pack(fill=tk.BOTH, expand=True)
        self.list_content = self.scroll_frame.scrollable_frame
        self._build_empty_state(
            self.list_content, "No updates found yet",
            "Run a scan to see which packages can be upgraded.",
            bg=self.colors["card_bg"], pady=62
        )

        self.console_frame = tk.Frame(
            body, bg=self.colors["console_bg"], height=self.console_idle_height,
            highlightthickness=1, highlightbackground=self.colors["border"])
        self.console_frame.pack(side=tk.BOTTOM, fill=tk.X)
        self.console_frame.pack_propagate(False)

        cbar = tk.Frame(self.console_frame, bg=self.colors["console_header"], padx=14, pady=10)
        cbar.pack(fill=tk.X)
        cbar_left = tk.Frame(cbar, bg=self.colors["console_header"])
        cbar_left.pack(side=tk.LEFT)
        tk.Label(cbar_left, text="\u25CF", bg=self.colors["console_header"],
                 fg=self.colors["secondary"], font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT)
        tk.Label(cbar_left, text="LIVE CONSOLE OUTPUT", bg=self.colors["console_header"],
                 fg=self.colors["text_dim"], font=FONTS["small_bold"]).pack(side=tk.LEFT, padx=(6, 0))
        clr = tk.Label(cbar, text="\U0001F5D1", bg=self.colors["console_header"],
                       fg=self.colors["text_soft"], font=("Segoe UI Symbol", 10), cursor="hand2")
        clr.pack(side=tk.RIGHT, padx=(8, 0))
        clr.bind("<Button-1>", lambda e: self.clear_console())
        cpy = tk.Label(cbar, text="\U0001F4CB", bg=self.colors["console_header"],
                       fg=self.colors["text_soft"], font=("Segoe UI Symbol", 10), cursor="hand2")
        cpy.pack(side=tk.RIGHT)
        cpy.bind("<Button-1>", lambda e: self.copy_console())

        self.console = scrolledtext.ScrolledText(
            self.console_frame, bg=self.colors["console_bg"], fg=self.colors["text_dim"],
            font=FONTS["mono"], borderwidth=0, padx=14, pady=12, state=tk.DISABLED,
            insertbackground=self.colors["text_main"]
        )
        self.console.pack(fill=tk.BOTH, expand=True)

    def _sort_updates(self, column):
        if self.sort_column == column:
            self.sort_ascending = not self.sort_ascending
        else:
            self.sort_column = column
            self.sort_ascending = True
        self.updates_list.sort(key=lambda x: x.get(column, '').lower(),
                               reverse=not self.sort_ascending)
        self._filter_updates()

    def _filter_updates(self):
        q = self.search_var.get().lower()
        if q in ("", "search updates..."):
            q = ""
        excluded = [p.lower() for p in self.config.get("excluded_packages", [])]
        group_ids = self._group_package_ids()
        for w in self.list_content.winfo_children():
            w.destroy()
        self.checkboxes = {}
        self.row_widgets = []
        self.focused_row_index = -1
        filtered = [u for u in self.updates_list
                    if (not q or q in u['name'].lower() or q in u['id'].lower())
                    and u['id'].lower() not in excluded
                    and (not group_ids or u['id'].lower() in group_ids)]
        if not filtered:
            if q:
                title = "No matching updates"
                detail = "Try a different search term to filter the available upgrades."
            elif group_ids:
                title = "No updates in selected group"
                detail = "This group has no pending upgrades after exclusions were applied."
            elif self.updates_list:
                title = "Everything is up to date"
                detail = "There are no visible upgrades left after your current filters."
            else:
                title = "No updates found"
                detail = "Run Check Updates to scan your winget sources."
            self._build_empty_state(
                self.list_content, title, detail, bg=self.colors["card_bg"], pady=62
            )
            return
        self._update_render_id = getattr(self, '_update_render_id', 0) + 1
        current_id = self._update_render_id

        def _pack_delayed(items, idx=0):
            if getattr(self, '_update_render_id', 0) != current_id: return
            batch_size = max(1, len(items) // 20)
            end_idx = min(idx + batch_size, len(items))
            for i in range(idx, end_idx):
                self._build_update_row(items[i])
            if end_idx < len(items):
                self.root.after(10, lambda: _pack_delayed(items, end_idx))
        self._refresh_group_ui()
        _pack_delayed(filtered)

    def _build_update_row(self, data):
        data["icon_ref"] = data.get("icon_ref") or self._get_package_icon_ref(data)
        row = tk.Frame(
            self.list_content, bg=self.colors["card_inner"], pady=SPACING["row_pad_y"],
            padx=18, highlightthickness=1,
            highlightbackground=self.colors["border_soft"]
        )
        row.pack(fill=tk.X, pady=4)
        self.row_widgets.append(row)

        bg_normal = self.colors["card_inner"]
        bg_hover = self.colors["surface_hover"]

        def _set_bg(widget, bg):
            try:
                if not getattr(widget, "preserve_bg", False):
                    widget.config(bg=bg)
            except tk.TclError:
                pass
            for child in widget.winfo_children():
                _set_bg(child, bg)

        row.bind("<Enter>", lambda e: _set_bg(row, bg_hover))
        row.bind("<Leave>", lambda e: _set_bg(row, bg_normal))

        var = tk.BooleanVar(value=False)

        def toggle():
            var.set(not var.get())
            self._refresh_checks()

        row.columnconfigure(0, weight=0, minsize=44)
        row.columnconfigure(1, weight=0, minsize=58)
        row.columnconfigure(2, weight=3, minsize=220)
        row.columnconfigure(3, weight=2, minsize=180)
        row.columnconfigure(4, weight=1, minsize=110)
        row.columnconfigure(5, weight=1, minsize=120)

        chk = tk.Label(row, text="\u2610", bg=bg_normal,
                       fg=self.colors["text_dim"], font=("Segoe UI Symbol", 14),
                       cursor="hand2")
        chk.grid(row=0, column=0, sticky="w", padx=4)
        chk.bind("<Button-1>", lambda e: toggle())
        self.checkboxes[data['id']] = (var, chk, data)

        avatar = self._build_package_icon_widget(row, data, bg_normal, size=28)
        avatar.grid(row=0, column=1, sticky="w", padx=4)

        nf = tk.Frame(row, bg=bg_normal, cursor="hand2")
        nf.grid(row=0, column=2, sticky="w", padx=4)
        lbl_name = tk.Label(nf, text=data['name'], bg=bg_normal,
                            fg=self.colors["text_main"], font=FONTS["body_bold"],
                            anchor="w", cursor="hand2")
        lbl_name.pack(anchor="w")
        if data.get("manager") == "npm":
            publisher = "npm global package"
        else:
            publisher = data['id'].split('.')[0] if '.' in data['id'] else "Winget package"
        lbl_id = tk.Label(nf, text=publisher, bg=bg_normal,
                          fg=self.colors["text_soft"], font=FONTS["small"],
                          anchor="w", cursor="hand2")
        lbl_id.pack(anchor="w")

        def on_info_click(e, p_id=data['id'], p_name=data['name'], p_manager=data.get("manager", "winget")):
            self.show_info_panel(p_id, p_name, p_manager)

        pkg_id_lbl = tk.Label(row, text=data['id'], bg=bg_normal,
                              fg=self.colors["text_dim"], font=FONTS["mono_small"],
                              anchor="w")
        pkg_id_lbl.grid(row=0, column=3, sticky="w", padx=4)

        current_lbl = tk.Label(row, text=data['version'], bg=bg_normal,
                               fg=self.colors["text_dim"], font=FONTS["mono"],
                               anchor="w")
        current_lbl.grid(row=0, column=4, sticky="w", padx=4)

        latest_wrap = tk.Frame(row, bg=bg_normal)
        latest_wrap.grid(row=0, column=5, sticky="w", padx=4)
        latest_lbl = tk.Label(latest_wrap, text=data['available'], bg=bg_normal,
                              fg=self.colors["success"], font=FONTS["mono"],
                              anchor="w")
        latest_lbl.pack(side=tk.LEFT)
        latest_arrow = tk.Label(
            latest_wrap, text="\u2191", bg=bg_normal, fg=self.colors["success"],
            font=FONTS["body_bold"]
        )
        latest_arrow.pack(side=tk.LEFT, padx=(8, 0))

        for widget in [row, avatar, nf, lbl_name, lbl_id,
                       pkg_id_lbl, current_lbl, latest_wrap, latest_lbl, latest_arrow]:
            widget.bind("<Button-1>", on_info_click, add="+")

        for w in [row, avatar, nf, lbl_name, lbl_id,
                  pkg_id_lbl, current_lbl, latest_wrap, latest_lbl, latest_arrow]:
            w.bind("<Button-3>", lambda e, pkg=data: self._post_update_context_menu(e, pkg))

    def _exclude_pkg(self, pkg_id):
        excluded = self.config.get("excluded_packages", [])
        if pkg_id not in excluded:
            excluded.append(pkg_id)
            self.config.set("excluded_packages", excluded)
            self._filter_updates()
            Toast(self.root, f"Excluded: {pkg_id}", "info")

    def _clip(self, text):
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        Toast(self.root, "Copied to clipboard", "success")

    def _context_menu(self):
        return tk.Menu(
            self.root, tearoff=0, bg=self.colors["surface"], fg=self.colors["text_main"],
            activebackground=self.colors["primary"], activeforeground="white", font=FONTS["body"]
        )

    def _post_update_context_menu(self, event, pkg):
        menu = self._context_menu()
        menu.add_command(label="Update this package", command=lambda: self._run_updates([pkg]))
        menu.add_command(label="Exclude from updates", command=lambda: self._exclude_pkg(pkg["id"]))
        quiet_label = ("Remove from Quiet Mode" if self._is_quiet_pkg(pkg["id"]) else "Add to Quiet Mode")
        menu.add_command(label=quiet_label, command=lambda: self._toggle_quiet_pkg(pkg["id"]))
        menu.add_separator()
        menu.add_command(label="Copy Package ID", command=lambda: self._clip(pkg["id"]))
        if pkg.get("manager") == "npm":
            menu.add_command(
                label="View on npmjs.com",
                command=lambda: webbrowser.open(f"https://www.npmjs.com/package/{quote(pkg['id'], safe='@/')}"),
            )
        else:
            menu.add_command(label="View Changelog", command=lambda: self._show_changelog_for_package(pkg))
            menu.add_command(
                label="View on winget.run",
                command=lambda: webbrowser.open(f"https://winget.run/pkg/{pkg['id'].replace('.', '/')}"),
            )
        menu.post(event.x_root, event.y_root)

    def _post_installed_context_menu(self, event, pkg):
        manager = pkg.get("manager", "winget")
        menu = self._context_menu()
        menu.add_command(label="Show Details", command=lambda: self.show_info_panel(
            pkg["id"], pkg["name"], manager
        ))
        if manager == "winget":
            previous = self.history.previous_version(pkg["id"], pkg.get("version"), manager)
            menu.add_command(
                label="View Changelog",
                command=lambda: self._show_changelog_for_package(pkg),
            )
            menu.add_command(
                label="Rollback to Previous Version",
                state=tk.NORMAL if previous else tk.DISABLED,
                command=lambda: self._rollback_package(pkg),
            )
        menu.add_command(label="Copy Package ID", command=lambda: self._clip(pkg["id"]))
        menu.post(event.x_root, event.y_root)

    def _history_entry_manager(self, entry):
        manager = entry.get("manager")
        if manager:
            return manager
        installed = self._find_installed_package(entry.get("package_id"))
        return installed.get("manager", "winget") if installed else "winget"

    def _post_history_context_menu(self, event, entry):
        manager = self._history_entry_manager(entry)
        package_id = entry.get("package_id", "")
        package_name = entry.get("package_name", package_id)
        installed_pkg = self._find_installed_package(package_id, manager)
        menu = self._context_menu()
        menu.add_command(label="Show Details", command=lambda: self.show_info_panel(
            package_id, package_name, manager
        ))
        if manager == "winget":
            can_rollback = bool(installed_pkg and self.history.previous_version(
                package_id, installed_pkg.get("version"), manager
            ))
            menu.add_command(
                label="View Changelog",
                command=lambda: self._show_changelog_for_package({
                    "id": package_id,
                    "name": package_name,
                    "manager": manager,
                    "version": (installed_pkg or {}).get("version", ""),
                }),
            )
            menu.add_command(
                label="Rollback to Previous Version",
                state=tk.NORMAL if can_rollback else tk.DISABLED,
                command=lambda pkg=installed_pkg: self._rollback_package(pkg),
            )
        menu.add_command(label="Copy Package ID", command=lambda: self._clip(package_id))
        menu.post(event.x_root, event.y_root)

    def _rollback_package(self, pkg):
        if not pkg or pkg.get("manager", "winget") != "winget":
            Toast(self.root, "Rollback is only supported for winget packages", "info")
            return
        if self.is_checking or self.is_updating:
            Toast(self.root, "Wait for the current operation to finish first", "warning")
            return
        pkg_id = pkg.get("id", "")
        current_version = pkg.get("version", "")
        target_version = self.history.previous_version(pkg_id, current_version, "winget")
        if not target_version:
            Toast(self.root, "No previous version is available for rollback", "warning")
            return
        if not messagebox.askyesno(
            "Rollback Package",
            f"Rollback {pkg.get('name', pkg_id)} from {current_version or '?'} to {target_version}?",
        ):
            return

        self.is_updating = True
        self.cancel_requested.clear()
        self._animate_console_height(self.console_active_height)
        self.btn_check.config(state=tk.DISABLED)
        self.btn_update_sel.config(state=tk.DISABLED)
        self.btn_update_all.config(state=tk.DISABLED)
        if hasattr(self, "btn_update_group"):
            self.btn_update_group.config(state=tk.DISABLED, bg=self.colors["surface_alt"])
        self.btn_cancel.pack(side=tk.LEFT, padx=5)
        self.progress_var.set(0)
        self.progress_bar.config(mode="indeterminate")
        self.progress_bar.pack(fill=tk.X, padx=32)
        self.progress_bar.start(16)

        def run():
            try:
                self.root.after(0, lambda: self.log(
                    f"[{pkg_id}:rollback] Rolling back to {target_version}...", self.colors["warning"]
                ))
                self._run_with_retry(
                    [
                        "winget", "install", "--id", pkg_id, "--exact", "--version", target_version,
                        "--force", "--accept-package-agreements", "--accept-source-agreements",
                        "--disable-interactivity",
                    ],
                    log_prefix=f"[{pkg_id}:rollback]",
                    timeout=300,
                    emit_logs=True,
                )
                self.history.add(
                    pkg_id, pkg.get("name", pkg_id), current_version, target_version, "success", manager="winget"
                )
                self.history.record_version(pkg_id, target_version, manager="winget")
                self.root.after(0, lambda: Toast(self.root, f"Rolled back {pkg_id} to {target_version}", "success"))
                self.root.after(0, self._refresh_history)
                self.root.after(150, self.load_installed)
                self.root.after(150, self.check_updates)
            except Exception as exc:
                self.root.after(0, lambda e=str(exc): Toast(self.root, f"Rollback failed: {e}", "error"))
            finally:
                self.is_updating = False
                self.root.after(0, lambda: self.btn_cancel.pack_forget())
                self.root.after(0, lambda: self._animate_console_height(self.console_idle_height))
                self.root.after(0, self._stop_progress)
                self.root.after(0, self._refresh_group_ui)

    def _refresh_checks(self):
        any_on = False
        for uid, (var, lbl, _) in self.checkboxes.items():
            if var.get():
                lbl.config(text="\u2611", fg=self.colors["primary"])
                any_on = True
            else:
                lbl.config(text="\u2610", fg=self.colors["text_dim"])
        st = tk.NORMAL if any_on else tk.DISABLED
        bg = self.colors["accent"] if any_on else self.colors["surface_alt"]
        self.btn_update_sel.config(state=st, bg=bg)
        self._refresh_group_ui()

    def toggle_all_selection(self):
        self.all_selected_var = not self.all_selected_var
        for uid, (var, lbl, _) in self.checkboxes.items():
            var.set(self.all_selected_var)
        self.select_all_lbl.config(
            text="\u2611" if self.all_selected_var else "\u2610",
            fg=self.colors["primary"] if self.all_selected_var else self.colors["text_dim"])
        self._refresh_checks()

    # ----------------------------------------------------------------- history
    # ----------------------------------------------------------- discover page (feature 6)
    def _build_discover_page(self):
        page = tk.Frame(self.main_container, bg=self.colors["bg"])
        self.pages["discover"] = page

        _, h_right = self._build_page_header(
            page, "Discover", "\u25CF Search winget repository and install new packages"
        )

        body = tk.Frame(page, bg=self.colors["bg"], padx=SPACING["page_x"],
                        pady=SPACING["page_y"])
        body.pack(fill=tk.BOTH, expand=True)

        search_card = self._build_card(
            body, {"side": tk.TOP, "fill": tk.X, "pady": (0, SPACING["card_gap"])}
        )
        sf = tk.Frame(search_card, bg=self.colors["card_bg"], padx=SPACING["card_pad"],
                      pady=14)
        sf.pack(fill=tk.X)
        self.discover_search_var = tk.StringVar()
        self._make_search_entry(sf, self.discover_search_var, "Search winget packages...")
        search_btn = self._btn(sf, "Search", self.colors["primary"], "white",
                               self._do_discover_search)
        search_btn.pack(side=tk.RIGHT, padx=(10, 0))
        self.discover_search_btn = search_btn
        self.root.bind("<Return>", lambda e: self._do_discover_search()
                       if self.current_page == "discover" else self._kb_enter())

        self.discover_status = tk.Label(
            search_card, textvariable=self.discover_action_var, bg=self.colors["card_bg"],
            fg=self.colors["text_dim"], font=FONTS["small"], padx=SPACING["card_pad"]
        )
        self.discover_status.pack(anchor="w", pady=(0, 14))
        self.discover_progress = ttk.Progressbar(
            search_card, mode="indeterminate", style="Custom.Horizontal.TProgressbar"
        )

        results_card = self._build_card(
            body, {"side": tk.TOP, "fill": tk.BOTH, "expand": True}
        )
        self.discover_scroll = ScrollableFrame(results_card, bg_color=self.colors["card_bg"])
        self.discover_scroll.pack(fill=tk.BOTH, expand=True)
        self.discover_content = self.discover_scroll.scrollable_frame
        self._build_empty_state(self.discover_content, "Search for packages",
                                "Enter a name above to find packages to install.",
                                bg=self.colors["card_bg"], pady=62)

    def _do_discover_search(self):
        q = self.discover_search_var.get().strip()
        if not q or self.discover_busy:
            return
        self._set_discover_busy(f'Searching for "{q}"...')
        for w in self.discover_content.winfo_children():
            w.destroy()
        self._build_empty_state(self.discover_content, "Searching...",
                                f"Looking for \"{q}\" in winget repository...",
                                bg=self.colors["card_bg"], pady=62)

        def search():
            try:
                proc = subprocess.run(
                    ["winget", "search", q, "--accept-source-agreements",
                     "--disable-interactivity"],
                    capture_output=True, text=True, encoding='utf-8',
                    errors='ignore', timeout=30,
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
                if proc.returncode != 0:
                    raise RuntimeError(self._tail_message(
                        proc.stderr or proc.stdout,
                        f"Exit code {proc.returncode}",
                    ))
                results = self.parser.parse_list_output(proc.stdout)
                self.root.after(0, lambda: self._render_discover_results(results))
                self.root.after(0, lambda count=len(results): self._set_discover_idle(
                    f"Search complete • {count} result(s)", "success"
                ))
            except Exception as e:
                self.root.after(0, lambda: self._discover_error(str(e)))
        threading.Thread(target=search, daemon=True).start()

    def _render_discover_results(self, results):
        for w in self.discover_content.winfo_children():
            w.destroy()
        if not results:
            self._build_empty_state(self.discover_content, "No packages found",
                                    "Try a different search term.",
                                    bg=self.colors["card_bg"], pady=62)
            return
        for pkg in results[:50]:
            row = tk.Frame(self.discover_content, bg=self.colors["card_inner"],
                           padx=18, pady=10, highlightthickness=1,
                           highlightbackground=self.colors["border_soft"])
            row.pack(fill=tk.X, pady=3)
            info = tk.Frame(row, bg=self.colors["card_inner"])
            info.pack(side=tk.LEFT, fill=tk.X, expand=True)
            tk.Label(info, text=pkg.get("name", pkg.get("id", "?")),
                     bg=self.colors["card_inner"], fg=self.colors["text_main"],
                     font=FONTS["body_bold"]).pack(anchor="w")
            tk.Label(info, text=pkg.get("id", ""),
                     bg=self.colors["card_inner"], fg=self.colors["text_dim"],
                     font=FONTS["small"]).pack(anchor="w")
            ver = pkg.get("version", "")
            if ver:
                tk.Label(info, text=f"v{ver}",
                         bg=self.colors["card_inner"], fg=self.colors["text_soft"],
                         font=FONTS["micro"]).pack(anchor="w")
            install_btn = self._btn(row, "Install", self.colors["success"], "white",
                                    lambda p=pkg: self._discover_install(p))
            install_btn.pack(side=tk.RIGHT, padx=4)
            details_btn = self._btn(row, "Details", self.colors["surface_alt"],
                                    self.colors["text_main"],
                                    lambda p=pkg: self.show_info_panel(
                                        p.get("id", ""), p.get("name", "")),
                                    outline=True)
            details_btn.pack(side=tk.RIGHT, padx=4)

    def _discover_install(self, pkg):
        pkg_id = pkg.get("id", "")
        if self.discover_busy:
            return
        self._set_discover_busy(f"Installing {pkg_id}...")
        Toast(self.root, f"Installing {pkg_id}...", "info")
        def install():
            try:
                proc = subprocess.run(
                    ["winget", "install", "--id", pkg_id, "--exact",
                     "--accept-package-agreements", "--accept-source-agreements",
                     "--disable-interactivity"],
                    capture_output=True, text=True, timeout=300,
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
                if proc.returncode != 0:
                    raise RuntimeError(self._tail_message(
                        proc.stderr or proc.stdout,
                        f"Exit code {proc.returncode}",
                    ))
                version = pkg.get("version") or pkg.get("available") or ""
                self.history.add(pkg_id, pkg.get("name", pkg_id), "", version, "success", manager="winget")
                self.history.record_version(pkg_id, version, manager="winget")
                self.root.after(0, lambda: Toast(self.root, f"Installed {pkg_id}", "success"))
                self.root.after(0, lambda: self._set_discover_idle(
                    f"Installed {pkg_id}", "success"
                ))
                self.root.after(200, self.load_installed)
                self.root.after(200, self.check_updates)
            except Exception as e:
                self.root.after(0, lambda: self._set_discover_idle(
                    f"Install failed for {pkg_id}", "error"
                ))
                self.root.after(0, lambda: Toast(self.root, f"Install failed: {e}", "error"))
        threading.Thread(target=install, daemon=True).start()

    def _discover_error(self, msg):
        for w in self.discover_content.winfo_children():
            w.destroy()
        self._build_empty_state(self.discover_content, "Search failed", msg,
                                bg=self.colors["card_bg"], pady=62)
        self._set_discover_idle("Search failed", "error")

    def _set_discover_busy(self, message):
        self.discover_busy = True
        self.discover_action_var.set(message)
        if hasattr(self, "discover_search_btn"):
            self.discover_search_btn.config(state=tk.DISABLED, bg=self.colors["surface_alt"])
        self.discover_progress.pack(fill=tk.X, padx=SPACING["card_pad"], pady=(0, 16))
        self.discover_progress.start(18)

    def _set_discover_idle(self, message, kind="normal"):
        self.discover_busy = False
        self.discover_action_var.set(message)
        if hasattr(self, "discover_search_btn"):
            self.discover_search_btn.config(state=tk.NORMAL, bg=self.colors["primary"])
        try:
            self.discover_progress.stop()
            self.discover_progress.pack_forget()
        except Exception:
            pass
        color = {
            "normal": self.colors["text_dim"],
            "success": self.colors["success"],
            "error": self.colors["danger"],
        }.get(kind, self.colors["text_dim"])
        self.discover_status.config(fg=color)

    # ----------------------------------------------------------- export script (feature 10)
    def _export_as_script(self):
        packages = self.installed_list
        if not packages:
            Toast(self.root, "No installed packages to export", "warning")
            return
        path = filedialog.asksaveasfilename(
            title="Export Reinstall Script",
            defaultextension=".ps1",
            filetypes=[("PowerShell", "*.ps1"), ("Batch", "*.bat"), ("All", "*.*")]
        )
        if not path:
            return
        is_bat = path.endswith(".bat")
        lines = []
        if is_bat:
            lines.append("@echo off")
            lines.append("echo Reinstalling packages...")
        else:
            lines.append("# Winget Update Manager — Reinstall Script")
            lines.append(f"# Generated {datetime.now().isoformat()}")
            lines.append("")
        for pkg in packages:
            mgr = pkg.get("manager", "winget")
            if mgr == "npm":
                cmd = f"npm install -g {pkg['id']}"
            else:
                cmd = f'winget install --id {pkg["id"]} --exact --accept-package-agreements --accept-source-agreements --disable-interactivity'
            if is_bat:
                lines.append(cmd)
            else:
                lines.append(cmd)
        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        Toast(self.root, f"Script exported to {Path(path).name}", "success")

    # ----------------------------------------------------------- retry logic (feature 11)
    def _run_with_retry(self, cmd, max_retries=3, backoff=2, log_prefix="", timeout=300,
                        emit_logs=True):
        last_error = None
        for attempt in range(max_retries):
            if self.cancel_requested.is_set():
                raise RuntimeError("Operation cancelled by user")
            try:
                proc = self._run_command_logged(
                    cmd, timeout=timeout, log_prefix=log_prefix, emit_logs=emit_logs
                )
                if proc.returncode == 0:
                    return proc
                last_error = RuntimeError(self._tail_message(
                    proc.stdout or proc.stderr,
                    f"{Path(cmd[0]).stem} exited with code {proc.returncode}",
                ))
            except subprocess.TimeoutExpired:
                last_error = RuntimeError(f"{Path(cmd[0]).stem} timed out after {timeout}s")
            except Exception as exc:
                last_error = exc

            if self.cancel_requested.is_set():
                raise RuntimeError("Operation cancelled by user")
            if attempt < max_retries - 1:
                wait = backoff ** (attempt + 1)
                retry_msg = f"{log_prefix} Retry {attempt + 1}/{max_retries - 1} in {wait}s..."
                self.root.after(0, lambda msg=retry_msg: self.log(msg, self.colors["warning"]))
                time.sleep(wait)

        raise last_error or RuntimeError(f"Command failed after {max_retries} attempts")

    # ----------------------------------------------------------- admin elevation (feature 15)
    @staticmethod
    def _is_admin():
        try:
            return ctypes.windll.shell32.IsUserAnAdmin() != 0
        except Exception:
            return False

    def _elevate(self):
        if self._is_admin():
            Toast(self.root, "Already running as Administrator", "info")
            return
        try:
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas", sys.executable, " ".join(sys.argv), None, 1)
            self.root.destroy()
        except Exception:
            Toast(self.root, "Failed to elevate privileges", "error")

    # ----------------------------------------------------------- startup integration (feature 14)
    def _toggle_startup(self, enable):
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        app_name = "WingetUpdateManager"
        try:
            if winreg is None:
                Toast(self.root, "Registry access not available", "error")
                return
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0,
                                winreg.KEY_SET_VALUE)
            if enable:
                winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ,
                                  subprocess.list2cmdline(self._background_command_parts("--minimized")))
            else:
                try:
                    winreg.DeleteValue(key, app_name)
                except FileNotFoundError:
                    pass
            winreg.CloseKey(key)
            Toast(self.root, f"Startup {'enabled' if enable else 'disabled'}", "success")
        except Exception as e:
            Toast(self.root, f"Startup toggle failed: {e}", "error")

    # ----------------------------------------------------------- windows notifications (feature 13)
    def _app_launch_target(self):
        target = Path(sys.executable if getattr(sys, "frozen", False) else __file__).resolve()
        try:
            return target.as_uri()
        except Exception:
            return None

    def _notify_native(self, title, message):
        if HAS_WINOTIFY:
            try:
                icon_path = get_resource_path("assets/icon.png")
                n = WinNotification(
                    app_id="Winget Update Manager",
                    title=title,
                    msg=message,
                    icon=str(icon_path) if icon_path.exists() else "",
                    duration="short"
                )
                n.set_audio(None, loop=False)
                launch_target = self._app_launch_target()
                if launch_target:
                    try:
                        n.launch = launch_target
                    except Exception:
                        pass
                    try:
                        n.add_actions(label="Open App", launch=launch_target)
                    except Exception:
                        pass
                n.show()
                return True
            except Exception:
                pass
        return False

    # ----------------------------------------------------------- history page
    def _build_history_page(self):
        page = tk.Frame(self.main_container, bg=self.colors["bg"])
        self.pages["history"] = page

        self.history_status, hr = self._build_page_header(
            page, "Update History", "\u25CF Recent upgrade activity and outcomes"
        )
        self._btn(hr, "Export CSV", self.colors["surface_alt"],
                  self.colors["text_main"], self._export_history,
                  outline=True).pack(side=tk.LEFT, padx=(0, 10))
        self._btn(hr, "Clear History", self.colors["danger"],
                  "white", self._clear_history).pack(side=tk.LEFT)

        body = tk.Frame(page, bg=self.colors["bg"], padx=SPACING["page_x"],
                        pady=SPACING["page_y"])
        body.pack(fill=tk.BOTH, expand=True)

        history_card = self._build_card(body)
        col_h = tk.Frame(history_card, bg=self.colors["surface_alt"], padx=SPACING["card_pad"],
                         pady=14)
        col_h.pack(fill=tk.X)
        for t, w in [("STATUS", 8), ("PACKAGE", 28), ("FROM \u2192 TO", 24), ("DATE", 20)]:
            tk.Label(col_h, text=t, bg=self.colors["surface_alt"], fg=self.colors["text_dim"],
                     font=FONTS["small_bold"], width=w, anchor="w").pack(side=tk.LEFT, padx=2)

        lc = tk.Frame(history_card, bg=self.colors["card_bg"], padx=18, pady=12)
        lc.pack(fill=tk.BOTH, expand=True)
        self.history_scroll = ScrollableFrame(lc, bg_color=self.colors["card_bg"])
        self.history_scroll.pack(fill=tk.BOTH, expand=True)
        self.history_content = self.history_scroll.scrollable_frame

    def _refresh_history(self):
        for w in self.history_content.winfo_children():
            w.destroy()
        entries = self.history.get_entries(200)
        self.history_status.config(
            text=f"\u25CF {len(entries)} recorded event{'s' if len(entries) != 1 else ''}",
            fg=self.colors["text_dim"]
        )
        if not entries:
            self._build_empty_state(
                self.history_content, "No update history yet",
                "Completed upgrades will show up here once you run them.",
                bg=self.colors["card_bg"], pady=72
            )
            return
        icons = {"success": ("\u2713", self.colors["secondary"]),
                 "failed": ("\u2717", self.colors["danger"]),
                 "skipped": ("\u2298", self.colors["warning"])}
        for entry in entries:
            row = tk.Frame(
                self.history_content, bg=self.colors["card_inner"], pady=SPACING["row_pad_y"],
                padx=18, highlightthickness=1,
                highlightbackground=self.colors["border_soft"]
            )
            row.pack(fill=tk.X, pady=4)
            def _bind_history_context(widget, item):
                widget.bind("<Button-3>", lambda e, entry=item: self._post_history_context_menu(e, entry), add="+")
                for child in widget.winfo_children():
                    _bind_history_context(child, item)
            ic, co = icons.get(entry["status"], ("?", self.colors["text_dim"]))
            tk.Label(row, text=ic, bg=self.colors["card_inner"], fg=co,
                     font=("Segoe UI", 12, "bold"), width=5).pack(side=tk.LEFT)
            nf = tk.Frame(row, bg=self.colors["card_inner"], width=220)
            nf.pack(side=tk.LEFT)
            nf.pack_propagate(False)
            tk.Label(nf, text=entry["package_name"], bg=self.colors["card_inner"],
                     fg=self.colors["text_main"], font=FONTS["body_bold"],
                     anchor="w").pack(fill=tk.X)
            tk.Label(nf, text=entry["package_id"], bg=self.colors["card_inner"],
                     fg=self.colors["text_soft"], font=FONTS["small"],
                     anchor="w").pack(fill=tk.X)
            tk.Label(row, text=f'{entry["old_version"]} \u2192 {entry["new_version"]}',
                     bg=self.colors["card_inner"], fg=self.colors["text_dim"],
                     font=FONTS["mono_small"], width=22, anchor="w").pack(
                         side=tk.LEFT, padx=10)
            try:
                dt = datetime.fromisoformat(entry["timestamp"])
                ts = dt.strftime("%Y-%m-%d %H:%M:%S")
            except Exception:
                ts = entry.get("timestamp", "")[:19]
            tk.Label(row, text=ts, bg=self.colors["card_inner"],
                     fg=self.colors["text_dim"], font=FONTS["small"]).pack(side=tk.LEFT)
            _bind_history_context(row, entry)

    def _export_history(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv"), ("JSON", "*.json")])
        if not path:
            return
        entries = self.history.get_entries(500)
        if path.endswith(".json"):
            with open(path, "w") as f:
                json.dump(entries, f, indent=2)
        else:
            with open(path, "w", newline="") as f:
                w = csv.DictWriter(f, fieldnames=[
                    "timestamp", "package_id", "package_name",
                    "old_version", "new_version", "status", "manager"])
                w.writeheader()
                w.writerows(entries)
        Toast(self.root, f"Exported {len(entries)} entries", "success")

    def _clear_history(self):
        if messagebox.askyesno("Clear History", "Clear all update history?"):
            self.history.clear()
            self._refresh_history()
            Toast(self.root, "History cleared", "info")

    # ---------------------------------------------------------------- settings
    def _build_settings_page(self):
        page = tk.Frame(self.main_container, bg=self.colors["bg"])
        self.pages["settings"] = page

        self._build_page_header(page, "Settings", "\u25CF Preferences, scheduling, and package controls")

        sc = tk.Frame(page, bg=self.colors["bg"], padx=SPACING["page_x"], pady=SPACING["page_y"])
        sc.pack(fill=tk.BOTH, expand=True)
        ss = ScrollableFrame(sc, bg_color=self.colors["bg"])
        ss.pack(fill=tk.BOTH, expand=True)
        content = ss.scrollable_frame

        # -- Appearance --
        self._section(content, "Appearance")
        r = self._setting_row(content, "Theme", "Switch between dark and light mode")
        self.theme_var = tk.StringVar(value=self.config.get("theme", "dark"))
        for val, txt in [("dark", "Dark"), ("light", "Light")]:
            tk.Radiobutton(r, text=txt, variable=self.theme_var, value=val,
                           bg=self.colors["card_bg"], fg=self.colors["text_main"],
                           selectcolor=self.colors["surface"],
                           activebackground=self.colors["card_bg"],
                           activeforeground=self.colors["text_main"],
                           font=FONTS["body"], command=self._apply_theme).pack(
                               side=tk.LEFT, padx=10)

        # -- Updates --
        self._section(content, "Updates")
        r = self._setting_row(content, "Auto-check on launch",
                              "Automatically scan for updates when the app starts")
        self.auto_check_var = tk.BooleanVar(
            value=self.config.get("auto_check_on_launch", True))
        tk.Checkbutton(r, variable=self.auto_check_var,
                       bg=self.colors["card_bg"],
                       activebackground=self.colors["card_bg"],
                       selectcolor=self.colors["surface"],
                       command=lambda: self.config.set(
                           "auto_check_on_launch",
                           self.auto_check_var.get())).pack(side=tk.LEFT)

        r = self._setting_row(content, "Cache freshness window",
                              "How long cached scan results stay fresh on the dashboard")
        self.cache_ttl_var = tk.IntVar(value=self.config.get("cache_ttl", 3600))
        for value, label in [(900, "15m"), (3600, "1h"), (14400, "4h")]:
            tk.Radiobutton(
                r, text=label, variable=self.cache_ttl_var, value=value,
                bg=self.colors["card_bg"], fg=self.colors["text_main"],
                selectcolor=self.colors["surface"],
                activebackground=self.colors["card_bg"],
                activeforeground=self.colors["text_main"],
                font=FONTS["body"],
                command=lambda: self.config.set("cache_ttl", self.cache_ttl_var.get()),
            ).pack(side=tk.LEFT, padx=10)

        r = self._setting_row(content, "Scheduled daily scans",
                              "Create a Windows Scheduled Task to scan daily in background")
        self.scheduled_scan_var = tk.BooleanVar(
            value=self.config.get("scheduled_scan", False))
        tk.Checkbutton(r, variable=self.scheduled_scan_var,
                       bg=self.colors["card_bg"],
                       activebackground=self.colors["card_bg"],
                       selectcolor=self.colors["surface"],
                       command=self._toggle_scheduled_scan).pack(side=tk.LEFT)

        r = self._setting_row(content, "Silent updates",
                              "Run updates with --silent flag")
        self.silent_var = tk.BooleanVar(value=self.config.get("silent_mode", True))
        tk.Checkbutton(r, variable=self.silent_var,
                       bg=self.colors["card_bg"],
                       activebackground=self.colors["card_bg"],
                       selectcolor=self.colors["surface"],
                       command=lambda: self.config.set(
                           "silent_mode", self.silent_var.get())).pack(side=tk.LEFT)

        if HAS_TRAY:
            r = self._setting_row(content, "Minimize to system tray",
                                  "Keep running in background when closed")
            self.tray_var = tk.BooleanVar(
                value=self.config.get("minimize_to_tray", False))
            tk.Checkbutton(r, variable=self.tray_var,
                           bg=self.colors["card_bg"],
                           activebackground=self.colors["card_bg"],
                           selectcolor=self.colors["surface"],
                           command=lambda: self.config.set(
                               "minimize_to_tray",
                               self.tray_var.get())).pack(side=tk.LEFT)

        r = self._setting_row(content, "Quiet Mode background updates",
                              "When --scan-only runs, silently update marked packages")
        self.quiet_mode_var = tk.BooleanVar(
            value=self.config.get("quiet_mode_enabled", False))
        tk.Checkbutton(r, variable=self.quiet_mode_var,
                       bg=self.colors["card_bg"],
                       activebackground=self.colors["card_bg"],
                       selectcolor=self.colors["surface"],
                       command=lambda: self.config.set(
                           "quiet_mode_enabled",
                           self.quiet_mode_var.get())).pack(side=tk.LEFT)

        r = self._setting_row(content, "Parallel updates",
                              "Run multiple updates concurrently (2-6 workers)")
        self.parallel_var = tk.IntVar(value=self.config.get("parallel_workers", 1))
        for val in [1, 2, 4]:
            tk.Radiobutton(r, text=str(val), variable=self.parallel_var, value=val,
                           bg=self.colors["card_bg"], fg=self.colors["text_main"],
                           selectcolor=self.colors["surface"],
                           activebackground=self.colors["card_bg"],
                           activeforeground=self.colors["text_main"],
                            font=FONTS["body"], command=lambda: self.config.set(
                                "parallel_workers", self.parallel_var.get())).pack(
                                    side=tk.LEFT, padx=10)

        self._section(content, "Update Groups")
        tk.Label(
            content, text="Create reusable package groups for targeted updates.",
            bg=self.colors["bg"], fg=self.colors["text_soft"], font=FONTS["small"]
        ).pack(anchor="w", padx=10, pady=(0, 6))

        group_wrap = tk.Frame(
            content, bg=self.colors["card_bg"], padx=20, pady=18,
            highlightthickness=1, highlightbackground=self.colors["border_soft"]
        )
        group_wrap.pack(fill=tk.X, padx=10, pady=3)
        group_wrap.columnconfigure(0, weight=1, minsize=180)
        group_wrap.columnconfigure(1, weight=2, minsize=360)

        group_left = tk.Frame(group_wrap, bg=self.colors["card_bg"])
        group_left.grid(row=0, column=0, sticky="nsew", padx=(0, 16))
        tk.Label(
            group_left, text="Groups", bg=self.colors["card_bg"],
            fg=self.colors["text_main"], font=FONTS["body_bold"]
        ).pack(anchor="w", pady=(0, 6))
        self.group_listbox = tk.Listbox(
            group_left, bg=self.colors["surface"], fg=self.colors["text_main"],
            font=FONTS["body"], selectbackground=self.colors["primary"],
            highlightthickness=0, relief=tk.FLAT, height=8
        )
        self.group_listbox.pack(fill=tk.BOTH, expand=True)
        self.group_listbox.bind("<<ListboxSelect>>", lambda e: self._on_group_listbox_select())

        group_right = tk.Frame(group_wrap, bg=self.colors["card_bg"])
        group_right.grid(row=0, column=1, sticky="nsew")
        group_right.columnconfigure(0, weight=1)

        tk.Label(
            group_right, text="Group name", bg=self.colors["card_bg"],
            fg=self.colors["text_main"], font=FONTS["body_bold"]
        ).grid(row=0, column=0, sticky="w")
        self.group_name_entry = tk.Entry(
            group_right, bg=self.colors["surface"], fg=self.colors["text_main"],
            insertbackground=self.colors["text_main"], font=FONTS["body"],
            relief=tk.FLAT, highlightthickness=1, highlightbackground=self.colors["border"]
        )
        self.group_name_entry.grid(row=1, column=0, sticky="ew", pady=(6, 8))

        group_actions = tk.Frame(group_right, bg=self.colors["card_bg"])
        group_actions.grid(row=2, column=0, sticky="w", pady=(0, 12))
        self._btn(group_actions, "Save Group", self.colors["primary"], "white",
                  self._save_group).pack(side=tk.LEFT)
        self._btn(group_actions, "Delete Group", self.colors["danger"], "white",
                  self._delete_group).pack(side=tk.LEFT, padx=(8, 0))

        tk.Label(
            group_right, text="Add package ID", bg=self.colors["card_bg"],
            fg=self.colors["text_main"], font=FONTS["body_bold"]
        ).grid(row=3, column=0, sticky="w")
        pkg_row = tk.Frame(group_right, bg=self.colors["card_bg"])
        pkg_row.grid(row=4, column=0, sticky="ew", pady=(6, 8))
        pkg_row.columnconfigure(0, weight=1)
        self.group_package_combo = ttk.Combobox(pkg_row, values=[], state="normal")
        self.group_package_combo.grid(row=0, column=0, sticky="ew")
        self._btn(pkg_row, "Add Package", self.colors["secondary"], "white",
                  self._add_package_to_group).grid(row=0, column=1, padx=(8, 0))

        tk.Label(
            group_right, text="Packages in group", bg=self.colors["card_bg"],
            fg=self.colors["text_main"], font=FONTS["body_bold"]
        ).grid(row=5, column=0, sticky="w")
        self.group_members_listbox = tk.Listbox(
            group_right, bg=self.colors["surface"], fg=self.colors["text_main"],
            font=FONTS["mono_small"], selectbackground=self.colors["primary"],
            highlightthickness=0, relief=tk.FLAT, height=8
        )
        self.group_members_listbox.grid(row=6, column=0, sticky="ew", pady=(6, 8))
        self._btn(group_right, "Remove Selected Package", self.colors["surface_alt"],
                  self.colors["text_main"], self._remove_package_from_group,
                  outline=True).grid(row=7, column=0, sticky="w")
        self._refresh_group_ui()

        # -- System --
        self._section(content, "System Integration")

        r = self._setting_row(content, "Start at login",
                              "Launch Winget Update Manager when Windows starts")
        self.startup_var = tk.BooleanVar(
            value=self.config.get("start_at_login", False))
        tk.Checkbutton(r, variable=self.startup_var,
                       bg=self.colors["card_bg"],
                       activebackground=self.colors["card_bg"],
                       selectcolor=self.colors["surface"],
                       command=lambda: (self.config.set(
                           "start_at_login", self.startup_var.get()),
                           self._toggle_startup(self.startup_var.get()))).pack(
                               side=tk.LEFT)

        r = self._setting_row(content, "Admin elevation",
                              "Restart the app with elevated privileges")
        admin_btn = self._btn(r, "\U0001F6E1 Run as Admin" if not self._is_admin() else "\u2705 Admin",
                              self.colors["warning"] if not self._is_admin() else self.colors["success"],
                              "white", self._elevate)
        admin_btn.pack(side=tk.LEFT, padx=10)

        # -- Quiet Mode --
        self._section(content, "Quiet Mode Packages")
        self._setting_row(content, "Marked packages",
                          "These packages will be silently updated by Quiet Mode")

        qf = tk.Frame(content, bg=self.colors["card_bg"], padx=20, pady=10,
                      highlightthickness=1, highlightbackground=self.colors["border"])
        qf.pack(fill=tk.X, padx=10, pady=5)

        self.quiet_listbox = tk.Listbox(
            qf, bg=self.colors["surface"], fg=self.colors["text_main"],
            font=FONTS["mono"], selectbackground=self.colors["primary"],
            highlightthickness=0, relief=tk.FLAT, height=6)
        self.quiet_listbox.pack(fill=tk.X, pady=(0, 5))
        self._refresh_quiet_listbox()

        qbf = tk.Frame(qf, bg=self.colors["card_bg"])
        qbf.pack(fill=tk.X)
        self.quiet_entry = tk.Entry(
            qbf, bg=self.colors["surface"], fg=self.colors["text_dim"],
            insertbackground=self.colors["text_main"], font=FONTS["mono"],
            relief=tk.FLAT, highlightthickness=1,
            highlightbackground=self.colors["border"])
        self.quiet_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=4)
        self.quiet_entry.insert(0, "Enter package ID...")
        self.quiet_entry.bind("<FocusIn>",
                              lambda e: self._ph_in(self.quiet_entry, "Enter package ID..."))
        self.quiet_entry.bind("<FocusOut>",
                              lambda e: self._ph_out(self.quiet_entry, "Enter package ID..."))
        self._btn(qbf, "Add", self.colors["primary"], "white",
                  self._add_quiet_package).pack(side=tk.LEFT, padx=(5, 0))
        self._btn(qbf, "Remove", self.colors["danger"], "white",
                  self._remove_quiet_package).pack(side=tk.LEFT, padx=(5, 0))

        # -- Exclusions --
        self._section(content, "Package Exclusions")
        self._setting_row(content, "Excluded packages",
                          "These packages will be hidden from the updates list")

        ef = tk.Frame(content, bg=self.colors["card_bg"], padx=20, pady=10,
                      highlightthickness=1, highlightbackground=self.colors["border"])
        ef.pack(fill=tk.X, padx=10, pady=5)

        self.excl_listbox = tk.Listbox(
            ef, bg=self.colors["surface"], fg=self.colors["text_main"],
            font=FONTS["mono"], selectbackground=self.colors["primary"],
            highlightthickness=0, relief=tk.FLAT, height=6)
        self.excl_listbox.pack(fill=tk.X, pady=(0, 5))
        for pkg in self.config.get("excluded_packages", []):
            self.excl_listbox.insert(tk.END, pkg)

        bf = tk.Frame(ef, bg=self.colors["card_bg"])
        bf.pack(fill=tk.X)
        self.excl_entry = tk.Entry(
            bf, bg=self.colors["surface"], fg=self.colors["text_dim"],
            insertbackground=self.colors["text_main"], font=FONTS["mono"],
            relief=tk.FLAT, highlightthickness=1,
            highlightbackground=self.colors["border"])
        self.excl_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=4)
        self.excl_entry.insert(0, "Enter package ID...")
        self.excl_entry.bind("<FocusIn>",
                             lambda e: self._ph_in(self.excl_entry, "Enter package ID..."))
        self.excl_entry.bind("<FocusOut>",
                             lambda e: self._ph_out(self.excl_entry, "Enter package ID..."))
        self._btn(bf, "Add", self.colors["primary"], "white",
                  self._add_exclusion).pack(side=tk.LEFT, padx=(5, 0))
        self._btn(bf, "Remove", self.colors["danger"], "white",
                  self._remove_exclusion).pack(side=tk.LEFT, padx=(5, 0))

    def _refresh_admin_indicator(self):
        if not hasattr(self, "admin_badge"):
            return
        if self._is_admin():
            self.admin_badge.config(text="\u2705 Elevated", fg=self.colors["success"])
        else:
            self.admin_badge.config(text="\U0001F6E1 Not Elevated", fg=self.colors["warning"])

    def _on_group_filter_changed(self):
        self._refresh_group_ui()
        self._filter_updates()

    def _on_group_listbox_select(self):
        if not hasattr(self, "group_listbox"):
            return
        sel = self.group_listbox.curselection()
        if not sel:
            self.selected_group_name = None
        else:
            self.selected_group_name = self.group_listbox.get(sel[0])
        self._populate_group_editor()

    def _populate_group_editor(self):
        if not hasattr(self, "group_name_entry"):
            return
        groups = self._get_update_groups()
        name = self.selected_group_name if self.selected_group_name in groups else ""
        self.group_name_entry.delete(0, tk.END)
        self.group_name_entry.insert(0, name)
        if hasattr(self, "group_members_listbox"):
            self.group_members_listbox.delete(0, tk.END)
            for pkg_id in groups.get(name, []):
                self.group_members_listbox.insert(tk.END, pkg_id)

    def _save_group(self):
        name = self.group_name_entry.get().strip() if hasattr(self, "group_name_entry") else ""
        if not name:
            Toast(self.root, "Enter a group name first", "warning")
            return
        groups = self._get_update_groups()
        members = groups.get(self.selected_group_name, []) if self.selected_group_name in groups else []
        if self.selected_group_name and self.selected_group_name in groups and self.selected_group_name != name:
            groups.pop(self.selected_group_name, None)
        groups[name] = members
        self.selected_group_name = name
        self._set_update_groups(groups)
        Toast(self.root, f"Saved group: {name}", "success")

    def _delete_group(self):
        if not self.selected_group_name:
            Toast(self.root, "Select a group to delete", "warning")
            return
        groups = self._get_update_groups()
        groups.pop(self.selected_group_name, None)
        deleted = self.selected_group_name
        self.selected_group_name = None
        self._set_update_groups(groups)
        Toast(self.root, f"Deleted group: {deleted}", "info")

    def _add_package_to_group(self):
        if not self.selected_group_name:
            Toast(self.root, "Save or select a group first", "warning")
            return
        pkg_id = self.group_package_combo.get().strip() if hasattr(self, "group_package_combo") else ""
        if not pkg_id:
            Toast(self.root, "Enter a package ID to add", "warning")
            return
        groups = self._get_update_groups()
        members = groups.setdefault(self.selected_group_name, [])
        if pkg_id.lower() not in {item.lower() for item in members}:
            members.append(pkg_id)
            groups[self.selected_group_name] = members
            self._set_update_groups(groups)
            Toast(self.root, f"Added {pkg_id} to {self.selected_group_name}", "success")
        else:
            Toast(self.root, f"{pkg_id} is already in {self.selected_group_name}", "info")

    def _remove_package_from_group(self):
        if not self.selected_group_name or not hasattr(self, "group_members_listbox"):
            return
        sel = self.group_members_listbox.curselection()
        if not sel:
            Toast(self.root, "Select a package to remove", "warning")
            return
        pkg_id = self.group_members_listbox.get(sel[0])
        groups = self._get_update_groups()
        members = [item for item in groups.get(self.selected_group_name, [])
                   if item.lower() != pkg_id.lower()]
        groups[self.selected_group_name] = members
        self._set_update_groups(groups)
        Toast(self.root, f"Removed {pkg_id} from {self.selected_group_name}", "info")

    def update_group(self):
        group_name = self._group_filter_name()
        targets = self._group_targets(group_name)
        if not group_name:
            Toast(self.root, "Select a group to update", "warning")
            return
        if not targets:
            Toast(self.root, f"No pending updates found for {group_name}", "info")
            return
        self._run_updates(targets)

    def _section(self, parent, title):
        tk.Label(parent, text=title, bg=self.colors["bg"],
                 fg=self.colors["text_main"],
                 font=FONTS["sub_header"]).pack(anchor="w", padx=10, pady=(24, 8))

    def _setting_row(self, parent, title, desc):
        frame = tk.Frame(parent, bg=self.colors["card_bg"], padx=20, pady=16,
                         highlightthickness=1,
                         highlightbackground=self.colors["border_soft"])
        frame.pack(fill=tk.X, padx=10, pady=3)
        left = tk.Frame(frame, bg=self.colors["card_bg"])
        left.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Label(left, text=title, bg=self.colors["card_bg"],
                 fg=self.colors["text_main"], font=FONTS["body_bold"]).pack(anchor="w")
        tk.Label(left, text=desc, bg=self.colors["card_bg"],
                 fg=self.colors["text_soft"], font=FONTS["small"]).pack(anchor="w", pady=(3, 0))
        right = tk.Frame(frame, bg=self.colors["card_bg"])
        right.pack(side=tk.RIGHT)
        return right

    def _add_exclusion(self):
        pkg = self.excl_entry.get().strip()
        if not pkg or pkg == "Enter package ID...":
            return
        excluded = self.config.get("excluded_packages", [])
        if pkg not in excluded:
            excluded.append(pkg)
            self.config.set("excluded_packages", excluded)
            self.excl_listbox.insert(tk.END, pkg)
            self.excl_entry.delete(0, tk.END)
            Toast(self.root, f"Excluded: {pkg}", "info")

    def _remove_exclusion(self):
        sel = self.excl_listbox.curselection()
        if not sel:
            return
        pkg = self.excl_listbox.get(sel[0])
        self.excl_listbox.delete(sel[0])
        excluded = self.config.get("excluded_packages", [])
        if pkg in excluded:
            excluded.remove(pkg)
            self.config.set("excluded_packages", excluded)
            Toast(self.root, f"Removed: {pkg}", "info")

    def _apply_theme(self):
        self.config.set("theme", self.theme_var.get())
        self.root.after(10, self._rebuild_ui)

    def _rebuild_ui(self):
        if getattr(self, "_is_rebuilding", False): return
        self._is_rebuilding = True
        
        self.colors = DARK_COLORS if self.config.get("theme") == "dark" else LIGHT_COLORS
        self.root.configure(bg=self.colors["bg"])
        self.setup_styles()
        
        old_page = self.current_page
        
        self.sidebar.destroy()
        self.info_panel.destroy()
        self.main_container.destroy()
        
        self.pages = {}
        self.nav_items = {}
        self.nav_accents = {}
        self.search_var = tk.StringVar()
        self.installed_search_var = tk.StringVar()
        self.notification_badges = []
        self.group_filter_var = tk.StringVar(value=self.group_filter_var.get() or "All Packages")
        
        self.setup_ui()
        self.switch_page(old_page)
        
        if hasattr(self, 'updates_list') and self.updates_list:
            self._render_updates()
        if hasattr(self, 'installed_list') and self.installed_list:
            self._render_installed()
            
        self._is_rebuilding = False

    # -------------------------------------------------------- winget operations
    def check_updates(self):
        if self.is_checking or self.is_updating:
            return
        self.is_checking = True
        self.cancel_requested.clear()
        self._animate_console_height(self.console_active_height)
        self.btn_check.config(state=tk.DISABLED, bg=self.colors["surface_alt"])
        if hasattr(self, "btn_update_group"):
            self.btn_update_group.config(state=tk.DISABLED, bg=self.colors["surface_alt"])
        self.btn_cancel.pack(side=tk.LEFT, padx=5)
        for w in self.list_content.winfo_children():
            w.destroy()
        self._build_empty_state(
            self.list_content, "Scanning for updates...",
            "Querying winget and npm sources for upgrade candidates.",
            bg=self.colors["card_bg"], pady=62
        )
        self._set_status("Scanning for updates...", "warning")
        self.log("Initializing package scan...")
        self.progress_var.set(0)
        self.progress_bar.config(mode="indeterminate")
        self.progress_bar.pack(fill=tk.X, padx=32)
        self.progress_bar.start(20)

        def run():
            try:
                self.root.after(0, lambda: self.log("Checking winget sources...", self.colors["text_dim"]))
                self.active_process = subprocess.Popen(
                    ["winget", "upgrade", "--accept-source-agreements",
                     "--disable-interactivity"],
                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    text=True, encoding='utf-8', errors='ignore',
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
                full = ""
                for line in self.active_process.stdout:
                    full += line
                    if line.strip():
                        self.root.after(0, lambda l=line.strip(): self.log(l))
                self.active_process.wait()
                winget_updates = self.parser.parse_upgrade_output(full)
                self.root.after(0, lambda: self.log("Checking npm global packages...", self.colors["text_dim"]))
                npm_updates = self._load_npm_updates()
                if npm_updates:
                    self.root.after(
                        0,
                        lambda count=len(npm_updates): self.log(
                            f"Found {count} npm package update(s).", self.colors["primary"]
                        ),
                    )
                self.updates_list = sorted(
                    winget_updates + npm_updates,
                    key=lambda item: (item.get("name") or "").lower(),
                )
                for pkg in self.updates_list:
                    pkg["icon_ref"] = self._get_package_icon_ref(pkg)
                self.scan_cache.save(updates=[{k: v for k, v in p.items() if k != 'icon_ref'} for p in self.updates_list])
                self.root.after(0, self._render_updates)
            except Exception as e:
                self.root.after(0, lambda: self._handle_error(str(e)))
            finally:
                self.is_checking = False
                self.active_process = None
                self.root.after(0, lambda: self.btn_cancel.pack_forget())
                self.root.after(0, lambda: self._animate_console_height(self.console_idle_height))
                self.root.after(0, self._stop_progress)
                self.root.after(0, self._refresh_group_ui)
        threading.Thread(target=run, daemon=True).start()

    def _render_updates(self):
        for w in self.list_content.winfo_children():
            w.destroy()
        self.checkboxes = {}
        count = len(self.updates_list)
        self.update_count = count
        if hasattr(self, 'update_badge_label'):
            self.update_badge_label.config(
                text=f"{count}" if count else "",
                bg=(self.colors["nav_badge_bg"] if count else self.update_badge_label.master.cget("bg"))
            )
        if hasattr(self, 'dash_cards'):
            self.dash_cards["last_scan"].config(
                text=self.scan_cache.last_scan_time())
        cache_state = "stale" if self.scan_cache.is_stale(self.config.get("cache_ttl", 3600)) else "fresh"
        if count == 0:
            self._build_empty_state(
                self.list_content, "Everything is up to date",
                "No packages currently need attention from winget or npm.",
                bg=self.colors["card_bg"], pady=62
            )
            self._set_status(f"All systems operational • cache {cache_state}", "warning" if cache_state == "stale" else "success")
        else:
            self._set_status(f"{count} updates pending • cache {cache_state}", "warning")
            self._filter_updates()
        self.btn_check.config(state=tk.NORMAL, bg=self.colors["surface_alt"])
        self._refresh_group_ui()

    def _stop_progress(self):
        try:
            self.progress_bar.stop()
            self.progress_bar.pack_forget()
            self.progress_var.set(0)
        except Exception:
            pass

    def update_selected(self):
        packages = [pkg for _, (var, _, pkg) in self.checkboxes.items() if var.get()]
        if packages:
            self._run_updates(packages)

    def update_all(self):
        self._run_updates(all_apps=True)

    def _resolve_update_targets(self, app_ids=None, all_apps=False):
        if all_apps:
            candidates = list(self.updates_list)
        else:
            candidates = []
            for item in app_ids or []:
                if isinstance(item, dict):
                    candidates.append(item)
                    continue
                pkg = next((u for u in self.updates_list if u["id"] == item), None)
                if pkg:
                    candidates.append(pkg)
        deduped = []
        seen = set()
        for pkg in candidates:
            key = (pkg.get("manager", "winget"), pkg.get("id"))
            if key in seen:
                continue
            seen.add(key)
            deduped.append(pkg)
        return deduped

    def cancel_operation(self):
        self.cancel_requested.set()
        terminated = False
        if self.active_process:
            try:
                self.active_process.terminate()
                terminated = True
            except Exception:
                pass
        with self.active_process_lock:
            active = list(self.active_processes)
        for proc in active:
            try:
                proc.terminate()
                terminated = True
            except Exception:
                pass
        if terminated or self.is_checking or self.is_updating:
            self.log("Operation cancelled by user.", self.colors["warning"])
            self._notify_status("Operation cancelled", "warning")

    def _run_updates(self, app_ids=None, all_apps=False):
        if self.is_updating:
            return
        targets = self._resolve_update_targets(app_ids=app_ids, all_apps=all_apps)
        if not targets:
            Toast(self.root, "No packages are queued for update", "info")
            return
        self.is_updating = True
        self.cancel_requested.clear()
        self._animate_console_height(self.console_active_height)
        self.btn_check.config(state=tk.DISABLED)
        self.btn_update_sel.config(state=tk.DISABLED)
        self.btn_update_all.config(state=tk.DISABLED)
        if hasattr(self, "btn_update_group"):
            self.btn_update_group.config(state=tk.DISABLED, bg=self.colors["surface_alt"])
        self.btn_cancel.pack(side=tk.LEFT, padx=5)
        self.progress_var.set(0)
        self.progress_bar.config(mode="determinate")
        self.progress_bar.pack(fill=tk.X, padx=32)

        def run():
            succeeded = 0
            failed = 0
            cancelled = 0
            fatal_error = None
            try:
                total = len(targets)
                workers = min(max(1, int(self.config.get("parallel_workers", 1) or 1)), max(total, 1))
                self.root.after(0, lambda w=workers, t=total: self.log(
                    f"Running {t} update(s) with {w} worker(s)...", self.colors["text_main"]
                ))

                def worker(pkg):
                    manager = pkg.get("manager", "winget")
                    prefix = f"[{pkg['id']}:{manager}]"
                    self.root.after(0, lambda p=prefix: self.log(f"{p} Starting update...", self.colors["text_main"]))
                    self._run_with_retry(
                        self._update_command_for_package(pkg),
                        log_prefix=prefix,
                        timeout=300,
                        emit_logs=True,
                    )
                    return pkg

                completed = 0
                with ThreadPoolExecutor(max_workers=workers) as executor:
                    future_map = {executor.submit(worker, pkg): pkg for pkg in targets}
                    for future in as_completed(future_map):
                        pkg = future_map[future]
                        completed += 1
                        pct = (completed / total) * 100 if total else 100
                        self.root.after(0, lambda value=pct: self.progress_var.set(value))
                        try:
                            future.result()
                            self.history.add(
                                pkg["id"], pkg["name"], pkg.get("version", ""),
                                pkg.get("available", ""), "success",
                                manager=pkg.get("manager", "winget"),
                            )
                            if pkg.get("manager", "winget") == "winget":
                                self.history.record_version(pkg["id"], pkg.get("available", ""), manager="winget")
                            succeeded += 1
                        except Exception as exc:
                            if "cancelled" in str(exc).lower():
                                cancelled += 1
                                continue
                            failed += 1
                            self.history.add(
                                pkg["id"], pkg["name"], pkg.get("version", ""),
                                pkg.get("available", ""), "failed",
                                manager=pkg.get("manager", "winget"),
                            )
                            self.root.after(
                                0,
                                lambda p=pkg["id"], e=str(exc): self.log(
                                    f"[{p}] Failed: {e}", self.colors["danger"]
                                ),
                            )
                self.root.after(0, lambda: self.progress_var.set(100))
            except Exception as e:
                fatal_error = e
                self.root.after(0, lambda: self._handle_error(str(e)))
            finally:
                if fatal_error is not None:
                    toast_msg = None
                elif self.cancel_requested.is_set() and not failed and not succeeded:
                    toast_msg = "Update operation cancelled."
                    toast_kind = "warning"
                elif cancelled:
                    toast_msg = f"Updated {succeeded} package(s), {failed} failed, {cancelled} cancelled."
                    toast_kind = "warning"
                elif failed:
                    toast_msg = f"Updated {succeeded} package(s), {failed} failed."
                    toast_kind = "warning"
                else:
                    toast_msg = f"Updated {succeeded} package(s)."
                    toast_kind = "success"
                if toast_msg:
                    self.root.after(0, lambda m=toast_msg, k=toast_kind: self._notify_status(m, k))
                self.is_updating = False
                self.active_process = None
                self.root.after(0, lambda: self.btn_cancel.pack_forget())
                self.root.after(0, lambda: self._animate_console_height(self.console_idle_height))
                self.root.after(0, self._stop_progress)
                self.root.after(0, self._refresh_history)
                self.root.after(0, self._refresh_group_ui)
                self.root.after(0, self.check_updates)
        threading.Thread(target=run, daemon=True).start()

    def _update_command_for_package(self, pkg):
        manager = pkg.get("manager", "winget")
        silent = ["--silent"] if self.config.get("silent_mode") else []
        if manager == "npm":
            target = f"{pkg['id']}@{pkg.get('available') or 'latest'}"
            return [self._npm_command(), "install", "-g", target] + silent
        return [
            "winget", "upgrade", "--id", pkg["id"], "--exact",
            "--accept-package-agreements", "--accept-source-agreements",
            "--disable-interactivity",
        ] + silent

    def _exec_command(self, command):
        self.active_process = subprocess.Popen(
            command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding='utf-8', errors='ignore',
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        for line in self.active_process.stdout:
            if line.strip():
                self.root.after(0, lambda l=line.strip(): self.log(l))
        self.active_process.wait()
        if self.active_process.returncode != 0:
            raise RuntimeError(f"{Path(command[0]).stem} exited with code {self.active_process.returncode}")

    def _exec_package_update(self, pkg):
        self._exec_command(self._update_command_for_package(pkg))

    def _exec_winget(self, args):
        self._exec_command(["winget"] + args)

    # ----------------------------------------------------------------- console
    def log(self, message, color=None):
        ts = datetime.now().strftime("%H:%M:%S")
        self.console.config(state=tk.NORMAL)
        
        self.console.insert(tk.END, f"[{ts}] > ", "text_dim")
        self.console.tag_config("text_dim", foreground=self.colors["text_dim"])
        
        if color:
            tag = f"c_{color}"
            self.console.tag_config(tag, foreground=color)
            self.console.insert(tk.END, f"{message}\n", tag)
        else:
            ANSI_PATTERN = re.compile(r'\x1b\[([0-9;]*)[mK]')
            parts = ANSI_PATTERN.split(message)
            current_tags = []
            
            for i in range(len(parts)):
                if i % 2 == 0:
                    text = parts[i]
                    if text:
                        if current_tags:
                            tag_name = "tag_" + "_".join(current_tags)
                            fg_color = None
                            for c in reversed(current_tags):
                                if c in ['30','31','32','33','34','35','36','37', '90','91','92','93','94','95','96','97']:
                                    fg_color = self._get_ansi_color(c)
                                    break
                            if fg_color:
                                self.console.tag_config(tag_name, foreground=fg_color)
                            self.console.insert(tk.END, text, tag_name)
                        else:
                            self.console.insert(tk.END, text)
                else:
                    codes = parts[i].split(';')
                    for code in codes:
                        if code == '0' or code == '':
                            current_tags = []
                        else:
                            if code not in current_tags:
                                current_tags.append(code)
            
            self.console.insert(tk.END, "\n")
            
        self.console.see(tk.END)
        self.console.config(state=tk.DISABLED)
        self.root.update_idletasks()

    def _get_ansi_color(self, code):
        colors = {
            '30': self.colors["bg"], '31': self.colors["danger"], '32': self.colors["success"],
            '33': self.colors["warning"], '34': self.colors["primary"], '35': self.colors["accent"],
            '36': '#06b6d4', '37': self.colors.get("text_main", "white"),
            '90': self.colors["text_dim"], '91': '#fca5a5', '92': '#86efac',
            '93': '#fde047', '94': '#93c5fd', '95': '#d8b4fe', '96': '#67e8f9', '97': 'white'
        }
        return colors.get(code, self.colors["text_main"])

    def clear_console(self):
        self.console.config(state=tk.NORMAL)
        self.console.delete(1.0, tk.END)
        self.console.config(state=tk.DISABLED)
        Toast(self.root, "Console cleared", "info")

    def copy_console(self):
        self.root.clipboard_clear()
        self.root.clipboard_append(self.console.get(1.0, tk.END).strip())
        Toast(self.root, "Console output copied", "success")

    def _set_status(self, text, type_="normal"):
        c = {"normal": self.colors["text_dim"], "success": self.colors["success"],
             "error": self.colors["danger"], "warning": self.colors["warning"]}
        self.status_sub.config(text=f"\u25CF  {text}",
                               fg=c.get(type_, self.colors["text_dim"]))

    def _handle_error(self, msg):
        self.log(f"ERROR: {msg}", self.colors["danger"])
        Toast(self.root, f"Error: {msg[:50]}", "error", 5000)

    # -------------------------------------------------------- keyboard shortcuts
    def bind_shortcuts(self):
        self.root.bind("<Control-r>", lambda e: self.check_updates())
        self.root.bind("<Control-a>", lambda e: self.toggle_all_selection()
                       if self.current_page == "updates" else None)
        self.root.bind("<Return>", lambda e: self._kb_enter())
        self.root.bind("<Escape>", lambda e: self.cancel_operation())
        self.root.bind("<Control-Key-1>", lambda e: self.switch_page("dashboard"))
        self.root.bind("<Control-Key-2>", lambda e: self.switch_page("installed"))
        self.root.bind("<Control-Key-3>", lambda e: self.switch_page("updates"))
        self.root.bind("<Control-Key-4>", lambda e: self.switch_page("history"))
        self.root.bind("<Control-Key-5>", lambda e: self.switch_page("settings"))
        self.root.bind("<Up>", lambda e: self._kb_navigate(-1))
        self.root.bind("<Down>", lambda e: self._kb_navigate(1))
        self.root.bind("<space>", lambda e: self._kb_toggle_checkbox())

    def _kb_navigate(self, direction):
        if self.current_page != "updates" or not self.row_widgets:
            return
        self.focused_row_index = max(0, min(
            len(self.row_widgets) - 1, self.focused_row_index + direction))
        self._highlight_focused_row()

    def _highlight_focused_row(self):
        for i, row in enumerate(self.row_widgets):
            try:
                bg = self.colors["surface_hover"] if i == self.focused_row_index else self.colors["card_inner"]
                row.configure(bg=bg)
                for child in row.winfo_children():
                    if not getattr(child, 'preserve_bg', False):
                        child.configure(bg=bg)
                        for sub in child.winfo_children():
                            if not getattr(sub, 'preserve_bg', False):
                                sub.configure(bg=bg)
            except Exception:
                pass

    def _kb_toggle_checkbox(self):
        if self.current_page != "updates" or self.focused_row_index < 0:
            return
        keys = list(self.checkboxes.keys())
        if 0 <= self.focused_row_index < len(keys):
            key = keys[self.focused_row_index]
            var, lbl, _ = self.checkboxes[key]
            var.set(not var.get())
            lbl.config(text="\u2611" if var.get() else "\u2610")
            self._refresh_checks()

    def _kb_enter(self):
        if self.current_page == "updates":
            if self.focused_row_index >= 0:
                keys = list(self.checkboxes.keys())
                if 0 <= self.focused_row_index < len(keys):
                    key = keys[self.focused_row_index]
                    _, _, pkg = self.checkboxes[key]
                    self.show_info_panel(pkg["id"], pkg["name"], pkg.get("manager", "winget"))
                    return
            self.update_selected()

    # ---------------------------------------------------------------- sys tray
    def _setup_tray(self):
        if not HAS_TRAY:
            return
        img_path = get_resource_path("assets/icon.png")
        if HAS_PIL and img_path.exists():
            try:
                img = Image.open(str(img_path)).convert("RGBA")
            except Exception:
                img = Image.new('RGB', (64, 64), '#3b82f6')
                d = ImageDraw.Draw(img)
                d.rectangle([16, 16, 48, 48], fill='white')
                d.text((24, 20), "W", fill='#3b82f6')
        else:
            img = Image.new('RGB', (64, 64), '#3b82f6')
            d = ImageDraw.Draw(img)
            d.rectangle([16, 16, 48, 48], fill='white')
            d.text((24, 20), "W", fill='#3b82f6')
        menu = pystray.Menu(
            pystray.MenuItem("Show", self._tray_show),
            pystray.MenuItem("Check Updates",
                             lambda: self.root.after(0, self.check_updates)),
            pystray.MenuItem("Run Quiet Mode Now",
                             lambda: self.root.after(0, self._run_quiet_mode_now)),
            pystray.MenuItem("Update All Silently", 
                             lambda: self.root.after(0, lambda: self._run_updates(all_apps=True))),
            pystray.MenuItem("Quit", self._tray_quit),
        )
        self.tray_icon = pystray.Icon("WingetUM", img,
                                       "Winget Update Manager", menu)
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def _tray_show(self):
        self.root.after(0, self.root.deiconify)

    def _tray_quit(self):
        if self.tray_icon:
            self.tray_icon.stop()
        self.root.after(0, self.root.destroy)

    # --------------------------------------------------------------- lifecycle
    def on_close(self):
        self.config.set("window_geometry", self.root.geometry())
        if HAS_TRAY and self.config.get("minimize_to_tray"):
            self.root.withdraw()
            if not self.tray_icon:
                self._setup_tray()
        else:
            if self.tray_icon:
                self.tray_icon.stop()
            self.root.destroy()

    # ----------------------------------------------------------------- helpers
    def _btn(self, parent, text, bg, fg, cmd, outline=False):
        button = tk.Button(
            parent, text=text, command=cmd, bg=bg, fg=fg,
            font=FONTS["body_bold"], relief=tk.FLAT, padx=18, pady=10,
            activebackground=(self.colors["surface_hover"] if outline else bg),
            activeforeground=fg, cursor="hand2", borderwidth=0,
            disabledforeground=self.colors["text_soft"], highlightthickness=1,
            highlightbackground=(self.colors["border"] if outline else bg),
            highlightcolor=(self.colors["border"] if outline else bg)
        )
        return button

    def _ph_in(self, widget, placeholder):
        if widget.get() == placeholder:
            widget.delete(0, tk.END)
            widget.config(fg=self.colors["text_main"])

    def _ph_out(self, widget, placeholder):
        if not widget.get():
            widget.insert(0, placeholder)
            widget.config(fg=self.colors["text_dim"])

    def _animate_console_height(self, target, step=18):
        if not self.console_frame:
            return
        if self.console_anim_job:
            try:
                self.root.after_cancel(self.console_anim_job)
            except Exception:
                pass
            self.console_anim_job = None

        def tick():
            current = self.console_frame.winfo_height()
            if current <= 1:
                current = int(self.console_frame.cget("height"))
            if current == target:
                self.console_anim_job = None
                return
            delta = step if current < target else -step
            new_height = current + delta
            if abs(target - current) <= step:
                new_height = target
            self.console_frame.config(height=new_height)
            if new_height != target:
                self.console_anim_job = self.root.after(15, tick)
            else:
                self.console_anim_job = None

        tick()

    def _background_command_parts(self, *extra_args):
        if getattr(sys, "frozen", False):
            return [sys.executable, *extra_args]

        python_exe = Path(sys.executable)
        pythonw = python_exe.with_name("pythonw.exe")
        interpreter = str(pythonw) if pythonw.exists() else sys.executable
        return [interpreter, os.path.abspath(__file__), *extra_args]

    def _build_task_action(self, *extra_args):
        return subprocess.list2cmdline(self._background_command_parts(*extra_args))

    def _run_command_capture(self, command, timeout=None):
        return subprocess.run(
            command,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
            timeout=timeout,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )

    def _track_process(self, proc):
        with self.active_process_lock:
            self.active_processes.add(proc)

    def _untrack_process(self, proc):
        with self.active_process_lock:
            self.active_processes.discard(proc)

    def _run_command_logged(self, command, timeout=300, log_prefix="", emit_logs=True):
        proc = subprocess.Popen(
            command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding="utf-8", errors="ignore",
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        output = []
        self._track_process(proc)

        def reader():
            try:
                for line in iter(proc.stdout.readline, ""):
                    if not line:
                        break
                    output.append(line)
                    stripped = line.rstrip()
                    if emit_logs and stripped:
                        msg = f"{log_prefix} {stripped}".strip()
                        self.root.after(0, lambda message=msg: self.log(message))
            except Exception:
                pass

        reader_thread = threading.Thread(target=reader, daemon=True)
        reader_thread.start()
        started = time.time()
        try:
            while proc.poll() is None:
                if self.cancel_requested.is_set():
                    try:
                        proc.terminate()
                    except Exception:
                        pass
                    raise RuntimeError("Operation cancelled by user")
                if timeout and (time.time() - started) > timeout:
                    proc.kill()
                    raise subprocess.TimeoutExpired(command, timeout)
                time.sleep(0.1)
            reader_thread.join(timeout=1.5)
            stdout = "".join(output)
            return subprocess.CompletedProcess(command, proc.returncode, stdout, "")
        finally:
            self._untrack_process(proc)

    def _run_winget_capture(self, args, timeout=None):
        return self._run_command_capture(["winget"] + list(args), timeout=timeout)

    def _run_npm_capture(self, args, timeout=None):
        return self._run_command_capture([self._npm_command()] + list(args), timeout=timeout)

    def _run_package_update_capture(self, pkg):
        return self._run_command_capture(self._update_command_for_package(pkg), timeout=300)

    def _scan_available_updates(self):
        proc = self._run_winget_capture(
            ["upgrade", "--accept-source-agreements", "--disable-interactivity"],
            timeout=120,
        )
        winget_updates = self.parser.parse_upgrade_output(proc.stdout)
        npm_updates = self._load_npm_updates()
        updates = sorted(
            winget_updates + npm_updates,
            key=lambda item: (item.get("name") or "").lower(),
        )
        for pkg in updates:
            pkg["icon_ref"] = self._get_package_icon_ref(pkg)
        return updates

    def _quiet_mode_packages(self):
        return [pkg.strip() for pkg in self.config.get("quiet_mode_packages", [])
                if isinstance(pkg, str) and pkg.strip()]

    def _is_quiet_pkg(self, pkg_id):
        return pkg_id.lower() in {pkg.lower() for pkg in self._quiet_mode_packages()}

    def _refresh_quiet_listbox(self):
        if not hasattr(self, "quiet_listbox"):
            return
        self.quiet_listbox.delete(0, tk.END)
        for pkg in self._quiet_mode_packages():
            self.quiet_listbox.insert(tk.END, pkg)

    def _set_quiet_packages(self, packages):
        cleaned = []
        seen = set()
        for pkg in packages:
            if not isinstance(pkg, str):
                continue
            item = pkg.strip()
            if not item or item.lower() in seen:
                continue
            seen.add(item.lower())
            cleaned.append(item)
        self.config.set("quiet_mode_packages", cleaned)
        self._refresh_quiet_listbox()

    def _toggle_quiet_pkg(self, pkg_id, announce=True):
        packages = self._quiet_mode_packages()
        lowered = {pkg.lower() for pkg in packages}
        if pkg_id.lower() in lowered:
            packages = [pkg for pkg in packages if pkg.lower() != pkg_id.lower()]
            self._set_quiet_packages(packages)
            if announce:
                Toast(self.root, f"Quiet Mode removed: {pkg_id}", "info")
            return False

        packages.append(pkg_id)
        self._set_quiet_packages(packages)
        if announce:
            Toast(self.root, f"Quiet Mode added: {pkg_id}", "success")
        return True

    def _add_quiet_package(self):
        pkg = self.quiet_entry.get().strip()
        if not pkg or pkg == "Enter package ID...":
            return
        packages = self._quiet_mode_packages()
        if pkg.lower() not in {item.lower() for item in packages}:
            packages.append(pkg)
            self._set_quiet_packages(packages)
            added = True
        else:
            added = False
        self.quiet_entry.delete(0, tk.END)
        Toast(self.root,
              f"{'Added to' if added else 'Already in'} Quiet Mode: {pkg}",
              "success" if added else "info")

    def _remove_quiet_package(self):
        sel = self.quiet_listbox.curselection()
        if not sel:
            return
        pkg = self.quiet_listbox.get(sel[0])
        packages = [item for item in self._quiet_mode_packages()
                    if item.lower() != pkg.lower()]
        self._set_quiet_packages(packages)
        Toast(self.root, f"Quiet Mode removed: {pkg}", "info")

    def _notify_status(self, message, kind="info"):
        self._record_notification(message, kind)
        if self._notify_native("Winget Update Manager", message):
            return
        if self.tray_icon:
            try:
                self.tray_icon.notify(message, "Winget UM")
                return
            except Exception:
                pass

        if self.root.winfo_exists() and self.root.state() != "withdrawn":
            Toast(self.root, message, kind, record=False)

    def _execute_quiet_mode(self, log_to_console=False):
        summary = {"configured": 0, "matched": 0, "updated": [], "failed": [], "error": None}
        quiet_lookup = {pkg.lower(): pkg for pkg in self._quiet_mode_packages()}
        summary["configured"] = len(quiet_lookup)
        if not quiet_lookup:
            return summary

        try:
            updates = self._scan_available_updates()
            targets = [pkg for pkg in updates if pkg["id"].lower() in quiet_lookup]
            summary["matched"] = len(targets)

            if log_to_console:
                self.root.after(0, lambda: self.log(
                    f"Quiet Mode: found {len(targets)} marked package updates.",
                    self.colors["primary"]))

            for pkg in targets:
                if log_to_console:
                    self.root.after(0, lambda p=pkg["id"], m=pkg.get("manager", "winget"): self.log(
                        f"Quiet Mode updating {m}: {p}...", self.colors["text_main"]))
                try:
                    proc = self._run_with_retry(
                        self._update_command_for_package(pkg),
                        log_prefix=f"[{pkg['id']}:{pkg.get('manager', 'winget')}]",
                        timeout=300,
                        emit_logs=log_to_console,
                    )
                except Exception as exc:
                    proc = None
                    error_message = str(exc)
                else:
                    error_message = ""

                if proc and proc.returncode == 0:
                    summary["updated"].append(pkg["id"])
                    self.history.add(pkg["id"], pkg["name"],
                                     pkg["version"], pkg["available"], "success",
                                     manager=pkg.get("manager", "winget"))
                    if pkg.get("manager", "winget") == "winget":
                        self.history.record_version(pkg["id"], pkg.get("available", ""), manager="winget")
                    if log_to_console:
                        self.root.after(0, lambda: self.log(
                            "Quiet Mode success", self.colors["success"]))
                else:
                    summary["failed"].append({
                        "id": pkg["id"],
                        "error": error_message or (proc.stderr.strip() if proc and proc.stderr else "") or (
                            proc.stdout.strip() if proc and proc.stdout else ""
                        ),
                    })
                    self.history.add(pkg["id"], pkg["name"],
                                     pkg["version"], pkg["available"], "failed",
                                     manager=pkg.get("manager", "winget"))
                    if log_to_console:
                        self.root.after(0, lambda p=pkg["id"]: self.log(
                            f"Quiet Mode failed: {p}", self.colors["danger"]))
        except Exception as e:
            summary["error"] = str(e)
        return summary

    def _run_quiet_mode_now(self):
        if self.is_checking or self.is_updating:
            return
        self.is_updating = True
        self.cancel_requested.clear()
        self._animate_console_height(self.console_active_height)
        self.btn_check.config(state=tk.DISABLED)
        self.btn_update_sel.config(state=tk.DISABLED)
        self.btn_update_all.config(state=tk.DISABLED)
        if hasattr(self, "btn_update_group"):
            self.btn_update_group.config(state=tk.DISABLED, bg=self.colors["surface_alt"])
        self.log("Quiet Mode: scanning marked packages...", self.colors["primary"])

        def run():
            summary = self._execute_quiet_mode(log_to_console=True)

            def finish():
                self.is_updating = False
                self._animate_console_height(self.console_idle_height)
                if summary["error"]:
                    self._handle_error(summary["error"])
                elif summary["updated"]:
                    self._notify_status(
                        f"Quiet Mode updated {len(summary['updated'])} package(s).",
                        "success")
                elif summary["failed"]:
                    self._notify_status(
                        f"Quiet Mode failed for {len(summary['failed'])} package(s).",
                        "error")
                elif summary["matched"]:
                    self._notify_status("Quiet Mode found updates but nothing needed to run.", "info")
                elif summary["configured"]:
                    self._notify_status("Quiet Mode found no pending updates for marked packages.", "info")
                else:
                    self._notify_status("Quiet Mode has no marked packages configured.", "info")
                self._refresh_group_ui()
                self.check_updates()

            self.root.after(0, finish)

        threading.Thread(target=run, daemon=True).start()

    def run_scan_only(self):
        def run():
            try:
                if self.config.get("quiet_mode_enabled") and self._quiet_mode_packages():
                    self._execute_quiet_mode(log_to_console=False)
                else:
                    self._scan_available_updates()
            finally:
                self.root.after(0, self.root.destroy)

        threading.Thread(target=run, daemon=True).start()

    def _toggle_scheduled_scan(self):
        val = self.scheduled_scan_var.get()
        self.config.set("scheduled_scan", val)
        task_name = "WingetUpdateManager_Daily"
        
        if val:
            try:
                subprocess.run(
                    ["schtasks", "/Create", "/SC", "DAILY", "/TN", task_name,
                     "/TR", self._build_task_action("--scan-only"),
                     "/ST", "12:00", "/F"],
                    check=True,
                    creationflags=subprocess.CREATE_NO_WINDOW,
                )
                Toast(self.root, "Scheduled task created", "success")
            except Exception as e:
                self.scheduled_scan_var.set(False)
                self.config.set("scheduled_scan", False)
                messagebox.showerror("Task Error", f"Failed to create scheduled task.\nMake sure you have permissions.\n{e}")
        else:
            try:
                subprocess.run(
                    ["schtasks", "/Delete", "/TN", task_name, "/F"],
                    check=True,
                    creationflags=subprocess.CREATE_NO_WINDOW,
                )
                Toast(self.root, "Scheduled task removed", "info")
            except Exception:
                pass

def main():
    root = tk.Tk()
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    app = WingetUpdateManager(root)
    
    if "--scan-only" in sys.argv:
        root.withdraw()
        app.run_scan_only()
    elif "--minimized" in sys.argv:
        if HAS_TRAY:
            root.withdraw()
            app._setup_tray()
        else:
            root.iconify()
             
    root.mainloop()

if __name__ == "__main__":
    main()
