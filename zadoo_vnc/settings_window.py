"""Small native settings window for installed Zadoo builds."""
from __future__ import annotations

import subprocess
import sys
import urllib.request
from pathlib import Path
from tkinter import BOTH, END, LEFT, RIGHT, StringVar, Tk, messagebox, ttk
import tkinter as tk

from .config import PROJECT_DIR
from .dependencies import HAS_WINPTY
from .settings import ACCESS_CODE_MAX_LENGTH, PERMISSION_KEYS, SettingsStore, get_settings_store

APP_PORT = 6173
BG = "#f4f7fb"
SURFACE = "#ffffff"
SURFACE_ALT = "#eaf0f7"
TEXT = "#17202a"
MUTED = "#657386"
BORDER = "#d7dee8"
ACCENT = "#18bfa6"
ACCENT_DARK = "#0d8f7d"
DANGER = "#c63f4c"

PERMISSION_LABELS = {
    "mouse": "Mouse",
    "keyboard": "Keyboard",
    "clipboard_pull": "Clipboard Pull",
    "clipboard_push": "Clipboard Push",
    "system_audio": "System Audio",
    "mic": "Mic",
    "camera": "Camera",
    "terminal": "Terminal",
    "snapshots": "Snapshots",
    "advanced_video": "Advanced Video",
    "tunnel_refresh": "Tunnel Refresh",
    "remote_alerts": "Remote Alerts",
}


def _local_url(path: str = "/") -> str:
    return f"http://127.0.0.1:{APP_PORT}{path}"


def _local_server_running() -> bool:
    try:
        with urllib.request.urlopen(_local_url("/api/settings/status"), timeout=0.6) as response:
            return int(getattr(response, "status", 0) or 0) < 500
    except Exception:
        return False


def _notify_settings_reload(access_code: str) -> None:
    request = urllib.request.Request(
        _local_url("/api/settings/reload"),
        headers={"X-Zadoo-Code": str(access_code or "")},
        method="GET",
    )
    try:
        with urllib.request.urlopen(request, timeout=1.2) as response:
            response.read()
    except Exception:
        pass


def _runtime_command() -> list[str]:
    if getattr(sys, "frozen", False):
        return [sys.executable, "--open"]
    return [sys.executable, "-m", "zadoo_vnc.app", "--open"]


def _find_icon() -> str:
    candidates = [
        Path(getattr(sys, "_MEIPASS", "")) / "app_icon.ico",
        Path(sys.executable).with_name("app_icon.ico"),
        Path(r"C:\Users\divya\Real\app_icon.ico"),
        PROJECT_DIR / "app_icon.ico",
    ]
    for path in candidates:
        try:
            if path.exists():
                return str(path)
        except Exception:
            pass
    return ""


class ZadooSettingsWindow:
    def __init__(self, store: SettingsStore | None = None):
        self.store = store or get_settings_store()
        self.root = Tk()
        self.root.title("Zadoo Settings")
        self.root.geometry("700x500")
        self.root.minsize(620, 460)
        self.root.configure(bg=BG)
        self.root.protocol("WM_DELETE_WINDOW", self.hide)
        self.root.bind("<Unmap>", self._on_unmap)
        self._hidden = False
        icon = _find_icon()
        if icon:
            try:
                self.root.iconbitmap(icon)
            except Exception:
                pass

        self.data = {}
        self.saved_access_code = ""
        self.alert_vars: dict[str, dict[str, tk.Variable]] = {}
        self.permission_vars: dict[str, tk.BooleanVar] = {}

        self._build_style()
        self._build_ui()
        self.reload()

    def _build_style(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure(".", font=("Segoe UI", 9), background=BG, foreground=TEXT)
        style.configure("Root.TFrame", background=BG)
        style.configure("Surface.TFrame", background=SURFACE)
        style.configure("TLabel", background=BG, foreground=TEXT)
        style.configure("Surface.TLabel", background=SURFACE, foreground=TEXT)
        style.configure("Muted.TLabel", background=BG, foreground=MUTED)
        style.configure("Status.TLabel", background=BG, foreground=MUTED)
        style.configure("Title.TLabel", background=BG, foreground=TEXT, font=("Segoe UI", 16, "bold"))
        style.configure(
            "TButton",
            background=SURFACE,
            foreground=TEXT,
            bordercolor=BORDER,
            focusthickness=1,
            focuscolor=BORDER,
            padding=(14, 8),
        )
        style.map("TButton", background=[("active", SURFACE_ALT)], foreground=[("disabled", "#9aa6b5")])
        style.configure(
            "Primary.TButton",
            background=ACCENT,
            foreground="#ffffff",
            bordercolor=ACCENT,
            font=("Segoe UI", 9, "bold"),
            padding=(16, 8),
        )
        style.map("Primary.TButton", background=[("active", ACCENT_DARK), ("pressed", ACCENT_DARK)])
        style.configure(
            "TEntry",
            fieldbackground=SURFACE,
            foreground=TEXT,
            bordercolor=BORDER,
            lightcolor=BORDER,
            darkcolor=BORDER,
            insertcolor=TEXT,
            padding=6,
        )
        style.map("TEntry", fieldbackground=[("disabled", SURFACE_ALT)])
        style.configure("TCheckbutton", background=SURFACE, foreground=TEXT, padding=(4, 4))
        style.map("TCheckbutton", background=[("active", SURFACE)])
        style.configure("TNotebook", background=BG, borderwidth=0, tabmargins=(0, 8, 0, 0))
        style.configure("TNotebook.Tab", background=SURFACE_ALT, foreground=MUTED, padding=(18, 9), borderwidth=0)
        style.map(
            "TNotebook.Tab",
            background=[("selected", SURFACE), ("active", "#f9fbfd")],
            foreground=[("selected", TEXT), ("active", TEXT)],
        )
        style.configure(
            "Card.TLabelframe",
            background=SURFACE,
            bordercolor=BORDER,
            relief="solid",
            padding=12,
        )
        style.configure(
            "Card.TLabelframe.Label",
            background=SURFACE,
            foreground=TEXT,
            font=("Segoe UI", 9, "bold"),
        )

    def _build_ui(self) -> None:
        outer = ttk.Frame(self.root, padding=16, style="Root.TFrame")
        outer.pack(fill=BOTH, expand=True)

        header = ttk.Frame(outer, style="Root.TFrame")
        header.pack(fill="x")
        ttk.Label(header, text="Zadoo Settings", style="Title.TLabel").pack(side=LEFT)
        self.state_label = ttk.Label(header, text="", style="Muted.TLabel")
        self.state_label.pack(side=RIGHT)

        self.tabs = ttk.Notebook(outer)
        self.tabs.pack(fill=BOTH, expand=True, pady=(10, 8))
        self._build_access_tab()
        self._build_permissions_tab()
        self._build_alert_tab()

        footer = ttk.Frame(outer, style="Root.TFrame")
        footer.pack(fill="x")
        self.status_label = ttk.Label(footer, text="", style="Status.TLabel")
        self.status_label.pack(side=LEFT, fill="x", expand=True)
        ttk.Button(footer, text="Hide", command=self.hide).pack(side=RIGHT, padx=(6, 0))
        ttk.Button(footer, text="Start Zadoo", command=self.start_zadoo, style="Primary.TButton").pack(side=RIGHT, padx=(6, 0))
        ttk.Button(footer, text="Save", command=self.save, style="Primary.TButton").pack(side=RIGHT)

    def _access_code_validator(self, value: str) -> bool:
        return len(value or "") <= ACCESS_CODE_MAX_LENGTH

    def _build_access_tab(self) -> None:
        tab = ttk.Frame(self.tabs, padding=16, style="Surface.TFrame")
        self.tabs.add(tab, text="Access")
        validator = (self.root.register(self._access_code_validator), "%P")

        form = ttk.Frame(tab, style="Surface.TFrame")
        form.pack(fill="x")
        ttk.Label(form, text=f"Access code ({ACCESS_CODE_MAX_LENGTH} chars max)", style="Surface.TLabel").grid(row=0, column=0, sticky="w")
        self.access_code = ttk.Entry(form, validate="key", validatecommand=validator)
        self.access_code.grid(row=1, column=0, sticky="ew", pady=(2, 10))

        ttk.Label(form, text="Email To", style="Surface.TLabel").grid(row=2, column=0, sticky="w")
        self.email_to = ttk.Entry(form)
        self.email_to.grid(row=3, column=0, sticky="ew", pady=(2, 10))

        ttk.Label(form, text="Resend API Key", style="Surface.TLabel").grid(row=4, column=0, sticky="w")
        key_row = ttk.Frame(form, style="Surface.TFrame")
        key_row.grid(row=5, column=0, sticky="ew", pady=(2, 8))
        key_row.columnconfigure(0, weight=1)
        self.resend_key = ttk.Entry(key_row, show="*")
        self.resend_key.grid(row=0, column=0, sticky="ew")
        ttk.Button(key_row, text="Reveal", command=self.reveal_resend_key).grid(row=0, column=1, padx=(6, 0))
        ttk.Button(key_row, text="Copy", command=self.copy_resend_key).grid(row=0, column=2, padx=(6, 0))

        self.email_state = ttk.Label(form, text="Email not Set", style="Surface.TLabel")
        self.email_state.grid(row=6, column=0, sticky="w")
        form.columnconfigure(0, weight=1)

    def _build_permissions_tab(self) -> None:
        tab = ttk.Frame(self.tabs, padding=16, style="Surface.TFrame")
        self.tabs.add(tab, text="Permissions")

        perms = ttk.LabelFrame(tab, text="Allowed controls for this password", style="Card.TLabelframe")
        perms.pack(fill=BOTH, expand=True)
        for index, key in enumerate(PERMISSION_KEYS):
            var = tk.BooleanVar(value=False)
            self.permission_vars[key] = var
            button = ttk.Checkbutton(perms, text=PERMISSION_LABELS.get(key, key), variable=var)
            if key == "terminal" and not HAS_WINPTY:
                button.configure(state="disabled")
            button.grid(
                row=index // 3,
                column=index % 3,
                sticky="w",
                padx=10,
                pady=6,
            )

        actions = ttk.Frame(tab, style="Surface.TFrame")
        actions.pack(fill="x", pady=(10, 0))
        ttk.Button(actions, text="Allow All", command=self.allow_all_permissions).pack(side=LEFT)
        ttk.Button(actions, text="Clear All", command=self.clear_permissions).pack(side=LEFT, padx=(6, 0))

    def _build_alert_tab(self) -> None:
        tab = ttk.Frame(self.tabs, padding=12, style="Surface.TFrame")
        self.tabs.add(tab, text="Alerts")
        alert_tabs = ttk.Notebook(tab)
        alert_tabs.pack(fill=BOTH, expand=True)
        for code in ("A", "B", "C", "D"):
            frame = ttk.Frame(alert_tabs, padding=16, style="Surface.TFrame")
            alert_tabs.add(frame, text=f"Alert {code}")
            enabled = tk.BooleanVar(value=False)
            title = StringVar(value="")
            message = StringVar(value="")
            self.alert_vars[code] = {"enabled": enabled, "title": title, "message": message}
            ttk.Checkbutton(frame, text="Enabled", variable=enabled).pack(anchor="w", pady=(0, 10))
            ttk.Label(frame, text="Title", style="Surface.TLabel").pack(anchor="w")
            ttk.Entry(frame, textvariable=title).pack(fill="x", pady=(2, 8))
            ttk.Label(frame, text="Message", style="Surface.TLabel").pack(anchor="w")
            ttk.Entry(frame, textvariable=message).pack(fill="x", pady=(2, 8))

    def _set_status(self, text: str) -> None:
        self.status_label.configure(text=text or "")

    def _setup_complete(self) -> bool:
        return bool(self.data.get("setup_complete"))

    def reload(self) -> None:
        self.data = self.store.load(reload=True)
        self.saved_access_code = str(self.data.get("access_code_plain") or "")
        self._load_access_values()
        self._load_permission_values()
        self._load_alert_values()
        self.state_label.configure(text="Configured" if self._setup_complete() else "First launch setup required")

    def _load_access_values(self) -> None:
        self.access_code.delete(0, END)
        self.access_code.insert(0, self.saved_access_code)
        self.email_to.delete(0, END)
        self.email_to.insert(0, str(self.data.get("email_to") or ""))
        self.resend_key.delete(0, END)
        has_email = bool(str(self.data.get("email_to") or "").strip())
        has_key = bool(self.store.get_resend_api_key())
        self.email_state.configure(text="Email configured" if has_email and has_key else "Email not Set")

    def _load_permission_values(self) -> None:
        permissions = self.data.get("permissions") or {}
        for key, var in self.permission_vars.items():
            var.set(bool(permissions.get(key)))
        if not HAS_WINPTY and "terminal" in self.permission_vars:
            self.permission_vars["terminal"].set(False)

    def _load_alert_values(self) -> None:
        alerts = self.data.get("alerts") or {}
        for code, vars_for_code in self.alert_vars.items():
            item = alerts.get(code) or {}
            vars_for_code["enabled"].set(bool(item.get("enabled")))
            vars_for_code["title"].set(str(item.get("title") or ""))
            vars_for_code["message"].set(str(item.get("message") or ""))

    def allow_all_permissions(self) -> None:
        for var in self.permission_vars.values():
            var.set(True)

    def clear_permissions(self) -> None:
        for var in self.permission_vars.values():
            var.set(False)

    def _collect_payload(self) -> dict:
        alerts = {}
        for code, vars_for_code in self.alert_vars.items():
            title = str(vars_for_code["title"].get() or "").strip()
            message = str(vars_for_code["message"].get() or "").strip()
            alerts[code] = {
                "enabled": bool(vars_for_code["enabled"].get()) and bool(title or message),
                "title": title,
                "message": message,
            }
        payload = {
            "admin_code": self.saved_access_code or self.access_code.get().strip(),
            "access_code": self.access_code.get().strip(),
            "email_to": self.email_to.get().strip(),
            "permissions": {key: bool(var.get()) for key, var in self.permission_vars.items()},
            "alerts": alerts,
        }
        if not HAS_WINPTY:
            payload["permissions"]["terminal"] = False
        resend_key = self.resend_key.get().strip()
        if resend_key:
            payload["resend_api_key"] = resend_key
        return payload

    def save(self) -> bool:
        try:
            payload = self._collect_payload()
            if not payload["access_code"]:
                raise ValueError("access code is required")
            if len(payload["access_code"]) > ACCESS_CODE_MAX_LENGTH:
                raise ValueError(f"access code must be {ACCESS_CODE_MAX_LENGTH} characters or fewer")
            self.store.apply_setup(payload, require_code=payload.get("admin_code"))
            if _local_server_running():
                _notify_settings_reload(payload.get("access_code", ""))
            self.reload()
            self._set_status("Saved")
            return True
        except PermissionError:
            self._set_status("Incorrect access code")
            return False
        except Exception as exc:
            self._set_status(str(exc) or "Save failed")
            return False

    def _admin_ok(self) -> bool:
        if not self._setup_complete():
            return True
        code = self.saved_access_code or self.access_code.get().strip()
        if not code:
            self._set_status("Save access code first")
            return False
        if not self.store.verify_access_code(code):
            self._set_status("Incorrect access code")
            return False
        return True

    def reveal_resend_key(self) -> None:
        if not self._admin_ok():
            return
        key = self.store.get_resend_api_key()
        self.resend_key.delete(0, END)
        self.resend_key.insert(0, key)
        self.resend_key.configure(show="")
        self._set_status("Resend key revealed" if key else "No Resend key saved")

    def copy_resend_key(self) -> None:
        if not self._admin_ok():
            return
        key = self.resend_key.get().strip() or self.store.get_resend_api_key()
        if not key:
            self._set_status("No Resend key saved")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(key)
        self._set_status("Resend key copied")

    def start_zadoo(self) -> None:
        if not self._setup_complete():
            if not self.save():
                return
        try:
            subprocess.Popen(
                _runtime_command(),
                cwd=str(PROJECT_DIR),
                close_fds=True,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            self.hide()
        except Exception as exc:
            messagebox.showerror("Zadoo", f"Could not start Zadoo: {exc}")

    def _on_unmap(self, event) -> None:
        if event.widget is self.root and self.root.state() == "iconic":
            self.root.after(80, self.hide)

    def hide(self) -> None:
        if self._hidden:
            return
        self._hidden = True
        try:
            self.root.withdraw()
        except Exception:
            pass
        self.root.after(50, self.root.quit)

    def run(self) -> None:
        self.root.mainloop()
        try:
            self.root.destroy()
        except Exception:
            pass


def run_settings_window() -> None:
    ZadooSettingsWindow().run()
