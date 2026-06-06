"""Small native settings window for installed Zadoo builds."""
from __future__ import annotations

import subprocess
import sys
import json
import time
import urllib.request
import webbrowser
from pathlib import Path
from tkinter import BOTH, END, LEFT, RIGHT, StringVar, Tk, messagebox, ttk
import tkinter as tk

from .config import PROJECT_DIR
from .dependencies import HAS_WINPTY
from .saas import ZadooCloudClient
from .settings import ACCESS_CODE_MAX_LENGTH, DEFAULT_ACCESS_CODE, DEFAULT_CLOUD_API_BASE, PERMISSION_KEYS, SettingsStore, get_settings_store
from .windows_startup import set_startup_task, startup_task_exists

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


def _post_local_json(path: str, payload: dict, timeout: float = 1.5) -> dict:
    request = urllib.request.Request(
        _local_url(path),
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json", "X-Zadoo-Code": str(payload.get("admin_code") or "")},
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8", "replace"))


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
        self.cloud = ZadooCloudClient(self.store)
        self.root = Tk()
        self.root.title("Zadoo Settings")
        self.root.geometry("520x380")
        self.root.minsize(520, 380)
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
        self.autostart_var = tk.BooleanVar(value=True)
        self.show_taskbar_var = tk.BooleanVar(value=False)
        self._activation_poll_after: str | None = None
        self._activation_poll_deadline = 0.0
        self._has_credits = False
        # Account tab avatar image reference (prevent GC)
        self._avatar_photo: tk.PhotoImage | None = None

        self._build_style()
        self._build_ui()
        self.reload()
        if self.store.get_device_token():
            self._fetch_latest_account_details_async(show_status=False)

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

        # Build tabs (but do not pack immediately, will be packed dynamically in reload())
        self.tabs = ttk.Notebook(outer)
        self._build_access_tab()
        self._build_account_tab()
        self._build_permissions_tab()
        self._build_alert_tab()
        self._build_runtime_tab()

        # Build welcome frame (but do not pack immediately)
        self.signin_welcome_frame = ttk.Frame(outer, style="Root.TFrame")
        card = ttk.Frame(self.signin_welcome_frame, style="Surface.TFrame", padding=30)
        card.place(relx=0.5, rely=0.5, anchor="center")

        ttk.Label(card, text="Welcome to Zadoo", style="Title.TLabel", font=("Segoe UI", 18, "bold")).pack(pady=(0, 10))
        ttk.Label(card, text="Please sign in to activate and link this device.", style="Surface.TLabel", font=("Segoe UI", 10)).pack(pady=(0, 20))
        ttk.Button(card, text="Sign in to Zadoo", command=self.start_activation, style="Primary.TButton").pack(pady=10)
        
        self.welcome_status_label = ttk.Label(card, text="", style="Surface.TLabel", font=("Segoe UI", 9), foreground=ACCENT)
        self.welcome_status_label.pack(pady=(10, 0))

        # Build footer (but do not pack immediately)
        self.footer = ttk.Frame(outer, style="Root.TFrame")
        self.status_label = ttk.Label(self.footer, text="", style="Status.TLabel")
        self.status_label.pack(side=LEFT, fill="x", expand=True)
        ttk.Button(self.footer, text="Hide", command=self.hide).pack(side=RIGHT, padx=(6, 0))
        ttk.Button(self.footer, text="Stop", command=self.stop_zadoo).pack(side=RIGHT, padx=(6, 0))
        ttk.Button(self.footer, text="Start Zadoo", command=self.start_zadoo, style="Primary.TButton").pack(side=RIGHT, padx=(6, 0))
        ttk.Button(self.footer, text="Save", command=self.save, style="Primary.TButton").pack(side=RIGHT)

    def _access_code_validator(self, value: str) -> bool:
        return len(value or "") <= ACCESS_CODE_MAX_LENGTH

    def _build_access_tab(self) -> None:
        self.tab_access = ttk.Frame(self.tabs, padding=16, style="Surface.TFrame")
        self.tabs.add(self.tab_access, text="Access")
        validator = (self.root.register(self._access_code_validator), "%P")

        form = ttk.Frame(self.tab_access, style="Surface.TFrame")
        form.pack(fill="x")
        ttk.Label(form, text=f"Access code ({ACCESS_CODE_MAX_LENGTH} chars max)", style="Surface.TLabel").grid(row=0, column=0, sticky="w")
        self.access_code = ttk.Entry(form, validate="key", validatecommand=validator)
        self.access_code.grid(row=1, column=0, sticky="ew", pady=(2, 10))

        ttk.Label(form, text="Email To", style="Surface.TLabel").grid(row=2, column=0, sticky="w")
        self.email_to = ttk.Entry(form)
        self.email_to.grid(row=3, column=0, sticky="ew", pady=(2, 10))

        # Resend API key is configured through the backend, not exposed in the UI.
        self.email_state = ttk.Label(form, text="Email not Set", style="Surface.TLabel")
        self.email_state.grid(row=4, column=0, sticky="w")
        form.columnconfigure(0, weight=1)

        # ── Remote access (public link + start) ───────────────────────
        self.public_url_frame = ttk.Frame(self.tab_access, style="Surface.TFrame")
        self.public_url_frame.pack(fill="x", pady=(18, 0))
        ttk.Label(self.public_url_frame, text="Public link", style="Surface.TLabel").pack(anchor="w")
        url_row = ttk.Frame(self.public_url_frame, style="Surface.TFrame")
        url_row.pack(fill="x", pady=(2, 6))
        self.public_url_label = ttk.Label(url_row, text="Zadoo not running", style="Muted.TLabel",
                                          font=("Segoe UI", 9), cursor="hand2")
        self.public_url_label.pack(side=LEFT, fill="x", expand=True)
        self.public_url_label.bind("<Button-1>", lambda _e: self._open_public_url())
        self.start_btn = ttk.Button(url_row, text="Start", command=self.start_zadoo, style="Primary.TButton")
        self.start_btn.pack(side=LEFT, padx=(6, 0))
        ttk.Button(url_row, text="Stop", command=self.stop_zadoo).pack(side=LEFT, padx=(6, 0))
        ttk.Button(url_row, text="Refresh", command=self._refresh_public_url).pack(side=LEFT, padx=(6, 0))
        ttk.Label(self.public_url_frame,
                  text="Click Start to launch Zadoo — the public link appears here, then click it to open.",
                  style="Muted.TLabel", font=("Segoe UI", 8)).pack(anchor="w")

    def _build_account_tab(self) -> None:
        self.tab_account = ttk.Frame(self.tabs, padding=16, style="Surface.TFrame")
        self.tabs.add(self.tab_account, text="Account")

        # ── Profile card ──────────────────────────────────────────────
        self.profile_card = ttk.LabelFrame(self.tab_account, text="Signed-in account", style="Card.TLabelframe")

        avatar_frame = ttk.Frame(self.profile_card, style="Surface.TFrame")
        avatar_frame.pack(side=LEFT, padx=(0, 12))
        self.avatar_label = tk.Label(
            avatar_frame, bg=SURFACE, width=5, height=2,
            relief="flat", font=("Segoe UI", 12, "bold"), fg=ACCENT
        )
        self.avatar_label.pack()

        info_frame = ttk.Frame(self.profile_card, style="Surface.TFrame")
        info_frame.pack(side=LEFT, fill="x", expand=True)
        self.profile_name_label = ttk.Label(info_frame, text="Not signed in", style="Surface.TLabel",
                                             font=("Segoe UI", 10, "bold"))
        self.profile_name_label.pack(anchor="w")
        self.profile_email_label = ttk.Label(info_frame, text="", style="Muted.TLabel",
                                              font=("Segoe UI", 9))
        self.profile_email_label.pack(anchor="w")
        self.account_state = ttk.Label(info_frame, text="", style="Muted.TLabel",
                                       font=("Segoe UI", 8))
        self.account_state.pack(anchor="w")

        ttk.Button(self.profile_card, text="Sign Out", command=self.sign_out).pack(side=RIGHT, padx=(0, 4))

        # ── Device name container ─────────────────────────────────────
        self.device_name_frame = ttk.Frame(self.tab_account, style="Surface.TFrame")
        ttk.Label(self.device_name_frame, text="Device name", style="Surface.TLabel").pack(anchor="w")
        self.device_name = ttk.Entry(self.device_name_frame)
        self.device_name.pack(fill="x", pady=(2, 10))

        # ── Sign-in Container ─────────────────────────────────────────
        self.signin_frame = ttk.Frame(self.tab_account, style="Surface.TFrame")
        ttk.Label(self.signin_frame, text="Sign in to your Zadoo account to start using this device.",
                  style="Surface.TLabel", font=("Segoe UI", 10, "bold"), foreground=ACCENT).pack(anchor="w", pady=(5, 10))

        code_row = ttk.Frame(self.signin_frame, style="Surface.TFrame")
        code_row.pack(fill="x")
        ttk.Label(code_row, text="Browser sign-in code", style="Surface.TLabel").pack(anchor="w")
        
        btn_row = ttk.Frame(code_row, style="Surface.TFrame")
        btn_row.pack(fill="x", pady=(2, 10))
        self.activation_code = ttk.Entry(btn_row, state="readonly")
        self.activation_code.pack(side=LEFT, fill="x", expand=True)
        ttk.Button(btn_row, text="Sign in to Zadoo", command=self.start_activation, style="Primary.TButton").pack(side=LEFT, padx=(8, 0))
        ttk.Button(btn_row, text="Check sign-in", command=self.poll_activation).pack(side=LEFT, padx=(8, 0))

        # ── Credits ───────────────────────────────────────────────────
        self.credits_card = ttk.LabelFrame(self.tab_account, text="Credits", style="Card.TLabelframe")
        self.credits_included_label = ttk.Label(self.credits_card, text="", style="Surface.TLabel")
        self.credits_included_label.pack(anchor="w")
        self.credits_wallet_label = ttk.Label(self.credits_card, text="", style="Surface.TLabel")
        self.credits_wallet_label.pack(anchor="w")
        self.credits_total_label = ttk.Label(self.credits_card, text="", style="Surface.TLabel",
                                             font=("Segoe UI", 9, "bold"))
        self.credits_total_label.pack(anchor="w", pady=(2, 6))
        self.billing_state = ttk.Label(self.credits_card, text="", style="Muted.TLabel",
                                       font=("Segoe UI", 8))
        self.billing_state.pack(anchor="w")

        credits_btns = ttk.Frame(self.credits_card, style="Surface.TFrame")
        credits_btns.pack(anchor="w", pady=(8, 0))
        ttk.Button(credits_btns, text="Add Balance →", command=self._add_balance,
                   style="Primary.TButton").pack(side=LEFT)
        ttk.Button(credits_btns, text="Add Credits →", command=self._open_pricing).pack(side=LEFT, padx=(8, 0))
        ttk.Button(credits_btns, text="Refresh Credits", command=self._refresh_credits).pack(side=LEFT, padx=(8, 0))
        ttk.Button(credits_btns, text="Copy Sign-in Code", command=self.copy_activation_code).pack(side=LEFT, padx=(8, 0))

    def _build_runtime_tab(self) -> None:
        self.tab_runtime = ttk.Frame(self.tabs, padding=16, style="Surface.TFrame")
        self.tabs.add(self.tab_runtime, text="Runtime")
        card = ttk.LabelFrame(self.tab_runtime, text="Windows behavior", style="Card.TLabelframe")
        card.pack(fill="x")
        ttk.Checkbutton(card, text="Start Zadoo when Windows starts", variable=self.autostart_var).grid(row=0, column=0, sticky="w", pady=4)
        ttk.Checkbutton(card, text="Keep Settings visible on taskbar when minimized", variable=self.show_taskbar_var).grid(row=1, column=0, sticky="w", pady=4)
        ttk.Button(card, text="Apply Startup", command=self.apply_startup).grid(row=2, column=0, sticky="w", pady=(12, 0))
        self.runtime_state = ttk.Label(self.tab_runtime, text="", style="Surface.TLabel")
        self.runtime_state.pack(anchor="w", pady=(14, 0))

    def _ensure_check_images(self) -> None:
        """Build a green ✓ 'checked' indicator and an empty 'unchecked' box (cached)."""
        if getattr(self, "_chk_on", None) is not None or getattr(self, "_chk_images_failed", False):
            return
        try:
            from PIL import Image, ImageDraw, ImageTk
            size = 18
            off = Image.new("RGBA", (size, size), (0, 0, 0, 0))
            on = Image.new("RGBA", (size, size), (0, 0, 0, 0))
            doff, don = ImageDraw.Draw(off), ImageDraw.Draw(on)
            try:
                doff.rounded_rectangle([1, 1, size - 2, size - 2], radius=4, outline="#9aa6b5", width=2)
                don.rounded_rectangle([1, 1, size - 2, size - 2], radius=4, fill=ACCENT, outline=ACCENT_DARK, width=1)
            except Exception:
                doff.rectangle([1, 1, size - 2, size - 2], outline="#9aa6b5", width=2)
                don.rectangle([1, 1, size - 2, size - 2], fill=ACCENT, outline=ACCENT_DARK, width=1)
            # white checkmark
            don.line([(4, 9), (8, 13)], fill="#ffffff", width=2)
            don.line([(8, 13), (14, 5)], fill="#ffffff", width=2)
            self._chk_off = ImageTk.PhotoImage(off)
            self._chk_on = ImageTk.PhotoImage(on)
        except Exception:
            self._chk_off = None
            self._chk_on = None
            self._chk_images_failed = True

    def _build_permissions_tab(self) -> None:
        self.tab_permissions = ttk.Frame(self.tabs, padding=16, style="Surface.TFrame")
        self.tabs.add(self.tab_permissions, text="Permissions")

        perms = ttk.LabelFrame(self.tab_permissions, text="Allowed controls for this password", style="Card.TLabelframe")
        perms.pack(fill=BOTH, expand=True)
        self._ensure_check_images()
        for index, key in enumerate(PERMISSION_KEYS):
            var = tk.BooleanVar(value=False)
            self.permission_vars[key] = var
            if self._chk_on is not None:
                # Custom green ✓ indicator (the themed glyph rendered like an ✗).
                button = tk.Checkbutton(
                    perms, text="  " + PERMISSION_LABELS.get(key, key), variable=var,
                    image=self._chk_off, selectimage=self._chk_on, indicatoron=False,
                    compound="left", bg=SURFACE, activebackground=SURFACE, selectcolor=SURFACE,
                    fg=TEXT, activeforeground=TEXT, font=("Segoe UI", 9),
                    borderwidth=0, highlightthickness=0, relief="flat",
                    offrelief="flat", overrelief="flat", anchor="w", cursor="hand2",
                )
            else:
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

        actions = ttk.Frame(self.tab_permissions, style="Surface.TFrame")
        actions.pack(fill="x", pady=(10, 0))
        ttk.Button(actions, text="Allow All", command=self.allow_all_permissions).pack(side=LEFT)
        ttk.Button(actions, text="Clear All", command=self.clear_permissions).pack(side=LEFT, padx=(6, 0))

    def _build_alert_tab(self) -> None:
        self.tab_alerts = ttk.Frame(self.tabs, padding=12, style="Surface.TFrame")
        self.tabs.add(self.tab_alerts, text="Alerts")
        alert_tabs = ttk.Notebook(self.tab_alerts)
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
        is_signed_in = bool(self.store.get_device_token())
        self._load_account_values()
        if is_signed_in:
            self._load_access_values()
            self._load_permission_values()
            self._load_alert_values()
            self._load_runtime_values()
        if is_signed_in:
            self.state_label.configure(text="Configured" if self._setup_complete() else "Ready")
        else:
            self.state_label.configure(text="")

    def _load_access_values(self) -> None:
        self.access_code.delete(0, END)
        self.access_code.insert(0, self.saved_access_code or DEFAULT_ACCESS_CODE)
        # Default "Email To" to the account used for sign-in when not explicitly set.
        email_to = str(self.data.get("email_to") or "").strip() or str(self.data.get("user_email") or "").strip()
        self.email_to.delete(0, END)
        self.email_to.insert(0, email_to)
        self.email_state.configure(text="Email configured" if email_to else "Email not Set")

    def _load_account_values(self) -> None:
        # Check signed in state
        is_signed_in = bool(self.store.get_device_token())

        # Adjust window controls, tabs, and footer visibility dynamically based on sign in status
        if is_signed_in:
            self.signin_welcome_frame.pack_forget()
            self.tabs.pack(fill=BOTH, expand=True, pady=(10, 8))
            self.footer.pack(fill="x")

            # Insert tabs only if not already present (guard against duplicate insert errors)
            existing_tabs = list(self.tabs.tabs())
            ordered = [
                (self.tab_access, "Access"),
                (self.tab_account, "Account"),
                (self.tab_permissions, "Permissions"),
                (self.tab_alerts, "Alerts"),
                (self.tab_runtime, "Runtime"),
            ]
            for idx, (tab_widget, label) in enumerate(ordered):
                tab_id = str(tab_widget)
                if tab_id not in existing_tabs:
                    self.tabs.insert(idx, tab_widget, text=label)

            self.root.minsize(680, 500)
            if self.root.winfo_width() < 680:
                self.root.geometry("780x560")
        else:
            self.tabs.pack_forget()
            self.footer.pack_forget()
            self.signin_welcome_frame.pack(fill=BOTH, expand=True, pady=(10, 8))

            self.root.minsize(520, 380)
            if self.root.winfo_width() > 540:
                self.root.geometry("520x380")

            # Check if there is an active activation process running
            activation = self.data.get("activation") or {}
            code = str(activation.get("code") or "")
            if code:
                self.welcome_status_label.configure(text=f"Waiting for browser sign-in approval...")
                if not self._activation_poll_after and self._activation_poll_deadline == 0.0:
                    self._activation_poll_deadline = time.time() + 900
                    self._schedule_activation_poll()
            else:
                self.welcome_status_label.configure(text="")

        # Show/Hide account tab components dynamically
        self.profile_card.pack_forget()
        self.device_name_frame.pack_forget()
        self.signin_frame.pack_forget()
        self.credits_card.pack_forget()

        if is_signed_in:
            self.profile_card.pack(fill="x", pady=(0, 10))
            self.device_name_frame.pack(fill="x", pady=(0, 10))
            self.credits_card.pack(fill="x", pady=(4, 0))
        else:
            self.device_name_frame.pack(fill="x", pady=(0, 10))
            self.signin_frame.pack(fill="x", pady=(0, 10))

        # Profile card details
        name = str(self.data.get("user_name") or "").strip()
        email = str(self.data.get("user_email") or "").strip()
        workspace = str(self.data.get("workspace_id") or "")
        device = str(self.data.get("device_id") or "")

        if is_signed_in:
            display_name = name or email or "Active Account"
            self.profile_name_label.configure(text=display_name)
            self.profile_email_label.configure(text=email if name else "")
            self._load_avatar_async(str(self.data.get("user_image_url") or ""), display_name)
        else:
            self.profile_name_label.configure(text="Not signed in")
            self.profile_email_label.configure(text="")
            self.avatar_label.configure(text="?", image="")
            self._avatar_photo = None

        self.account_state.configure(
            text=(f"Workspace {workspace[:8]}  ·  Device {device[:8]}" if workspace and device else "")
        )

        # Device name field
        cloud = (self.data.get("cloud") or {}) if "cloud" in self.data else self.data
        self.device_name.delete(0, END)
        self.device_name.insert(0, str(cloud.get("device_name") or self.data.get("device_name") or ""))

        # Activation code
        activation = self.data.get("activation") or {}
        code = str(activation.get("code") or "")
        self.activation_code.configure(state="normal")
        self.activation_code.delete(0, END)
        self.activation_code.insert(0, code)
        self.activation_code.configure(state="readonly")

        # Public URL label (Open button state is driven by credits below, not URL presence)
        pub_url = str(self.data.get("public_url") or "").strip()
        if pub_url:
            self.public_url_label.configure(text=pub_url, foreground=ACCENT)
        else:
            self.public_url_label.configure(text="Zadoo not running", foreground=MUTED)

        # Credits — heartbeat updates entitlement_cache (not credits_cache), so fall
        # back to entitlement_cache values to avoid a stale display.
        credits = self.data.get("credits_cache") or {}
        entitlement = self.data.get("entitlement_cache") or {}
        included = int(credits.get("includedMinutesRemaining")
                       or entitlement.get("includedMinutesRemaining") or 0)
        wallet = int(credits.get("walletMinutes")
                     or entitlement.get("walletMinutesRemaining") or 0)
        total = int(credits.get("totalMinutesRemaining") or (included + wallet))
        plan = str(credits.get("planCode") or entitlement.get("planCode") or "")

        allowed_flag = entitlement.get("allowed", credits.get("allowed"))
        self._has_credits = bool(total > 0 or allowed_flag is True)

        if credits or entitlement:
            self.credits_included_label.configure(text=f"Included: {included} min")
            self.credits_wallet_label.configure(text=f"Wallet: {wallet} min")
            self.credits_total_label.configure(text=f"Total remaining: {total} min{(' · ' + plan) if plan else ''}")
        else:
            self.credits_included_label.configure(text="Credits not loaded")
            self.credits_wallet_label.configure(text="")
            self.credits_total_label.configure(text="")

        allowed = entitlement.get("allowed", credits.get("allowed"))
        reason = entitlement.get("reason") or credits.get("reason") or ""
        if allowed is not None:
            self.billing_state.configure(
                text=f"{'✓ Active' if allowed else '✗ Blocked'}{': ' + reason if reason else ''}",
                foreground=(ACCENT if allowed else DANGER)
            )
        else:
            self.billing_state.configure(text="", foreground=MUTED)

    def _load_avatar_async(self, image_url: str, name: str) -> None:
        """Download avatar in background thread, fall back to initials canvas."""
        initials = "".join(p[0].upper() for p in name.split() if p)[:2] or "?"
        self.avatar_label.configure(text=initials, image="")
        self._avatar_photo = None
        if not image_url:
            return

        def _do_load():
            try:
                import io
                import tempfile
                from PIL import Image, ImageTk
                import urllib.request as _ur
                req = _ur.Request(image_url, headers={"User-Agent": "ZadooDesktop/1.0"})
                with _ur.urlopen(req, timeout=5) as resp:
                    data = resp.read()
                img = Image.open(io.BytesIO(data)).resize((48, 48), Image.LANCZOS)
                # Circular crop
                mask = Image.new("L", (48, 48), 0)
                from PIL import ImageDraw
                ImageDraw.Draw(mask).ellipse((0, 0, 47, 47), fill=255)
                img.putalpha(mask)
                photo = ImageTk.PhotoImage(img)
                def _set():
                    try:
                        self._avatar_photo = photo
                        self.avatar_label.configure(image=photo, text="")
                    except Exception:
                        pass
                self.root.after(0, _set)
            except Exception:
                pass  # Keep initials fallback

        import threading
        threading.Thread(target=_do_load, daemon=True).start()

    def _admin_code(self) -> str:
        return self.saved_access_code or self.access_code.get().strip()

    def _store_public_url(self, pub_url: str) -> None:
        try:
            data = self.store.load(reload=True)
            data["public_url"] = pub_url
            self.store.save(data)
            self.data = data
        except Exception:
            pass

    def _open_public_url(self) -> None:
        # Clicking the public link opens it in the browser; if there's none yet, hint to Start.
        pub_url = str(self.data.get("public_url") or "").strip()
        if pub_url:
            webbrowser.open(pub_url)
        elif not _local_server_running():
            self._set_status("Zadoo is not running — click Start first.")
        else:
            self._set_status("Tunnel not ready yet — click Refresh in a moment.")

    def _begin_public_url_autopoll(self, attempts: int = 8) -> None:
        """After Start, poll the runtime until the public link appears — or until the
        runtime tells us the tunnel is disabled (then show WHY instead of hanging)."""
        if attempts <= 0:
            return
        import threading
        def _poll():
            url = ""
            block_reason = ""
            tunnel_enabled = True
            running = False
            try:
                with urllib.request.urlopen(_local_url("/api/runtime/status"), timeout=3.0) as resp:
                    status = json.loads(resp.read().decode("utf-8", "replace"))
                    running = True
                    url = str(status.get("public_url") or "").strip()
                    tunnel_enabled = bool(status.get("tunnel_enabled", True))
                    block_reason = str(status.get("tunnel_block_reason") or "").strip()
            except Exception:
                pass
            def _update():
                if url:
                    self.public_url_label.configure(text=url, foreground=ACCENT)
                    self._store_public_url(url)
                    self._set_status("Public link ready — click it to open.")
                elif running and not tunnel_enabled:
                    # Runtime is up but the tunnel is off (billing/sign-in/etc) — stop hanging.
                    msg = block_reason or "Tunnel is disabled."
                    self.public_url_label.configure(text=msg, foreground=DANGER)
                    self._set_status(msg + "  Fix it, then click Start again.")
                elif attempts <= 1:
                    self.public_url_label.configure(text="Tunnel not ready — click Refresh", foreground=MUTED)
                    self._set_status("Tunnel is taking longer than expected. Click Refresh, or Stop and Start again.")
                else:
                    self.root.after(2500, lambda: self._begin_public_url_autopoll(attempts - 1))
            self.root.after(0, _update)
        threading.Thread(target=_poll, daemon=True).start()

    def _refresh_public_url(self) -> None:
        # "Refresh" always rotates the tunnel to produce a brand-new public URL.
        if not _local_server_running():
            self.public_url_label.configure(text="Zadoo not running", foreground=MUTED)
            self._set_status("Zadoo is not running — click Start Zadoo first.")
            return
        self._set_status("Generating new public URL...")
        import threading
        def _do_refresh():
            pub_url = ""
            error = ""
            try:
                result = _post_local_json(
                    "/api/runtime/refresh-tunnel",
                    {"admin_code": self._admin_code()},
                    timeout=30.0,
                )
                if result.get("success"):
                    pub_url = str(result.get("public_url") or result.get("url") or "").strip()
                else:
                    error = str(result.get("error") or "Could not refresh tunnel")
            except Exception as exc:
                error = str(exc) or "Could not refresh tunnel"
            def _update():
                if pub_url:
                    self.public_url_label.configure(text=pub_url, foreground=ACCENT)
                    self._store_public_url(pub_url)
                    self._set_status("New public URL ready.")
                else:
                    self.public_url_label.configure(text="Waiting for tunnel...", foreground=MUTED)
                    self._set_status(error or "Tunnel not ready yet — try again in a moment.")
            self.root.after(0, _update)
        threading.Thread(target=_do_refresh, daemon=True).start()

    def _refresh_credits(self) -> None:
        self._fetch_latest_account_details_async(show_status=True)

    def _fetch_latest_account_details_async(self, show_status: bool = False) -> None:
        if show_status:
            self._set_status("Refreshing account info...")
        import threading
        def _do_fetch():
            fetched_profile = False
            fetched_credits = False
            try:
                r = self.cloud.fetch_profile()
                if r.get("success"):
                    fetched_profile = True
            except Exception:
                pass
            try:
                r = self.cloud.fetch_credits()
                if r.get("success"):
                    fetched_credits = True
            except Exception:
                pass
            def _done():
                try:
                    self.reload()
                    if show_status:
                        if fetched_profile and fetched_credits:
                            self._set_status("Account info updated.")
                        else:
                            self._set_status("Partial refresh — check your internet connection.")
                except Exception:
                    pass
            self.root.after(0, _done)
        threading.Thread(target=_do_fetch, daemon=True).start()

    def _open_pricing(self) -> None:
        base = str(self.data.get("cloud_api_base") or DEFAULT_CLOUD_API_BASE).rstrip("/")
        webbrowser.open(f"{base}/pricing")

    def _add_balance(self) -> None:
        base = str(self.data.get("cloud_api_base") or DEFAULT_CLOUD_API_BASE).rstrip("/")
        webbrowser.open(f"{base}/dashboard/billing")



    def _load_runtime_values(self) -> None:
        self.autostart_var.set(bool(self.data.get("autostart_enabled", True)))
        self.show_taskbar_var.set(bool(self.data.get("show_settings_in_taskbar", False)))
        self.runtime_state.configure(text="Startup task installed" if startup_task_exists() else "Startup task not installed")

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
            "device_name": self.device_name.get().strip(),
            "autostart_enabled": bool(self.autostart_var.get()),
            "show_settings_in_taskbar": bool(self.show_taskbar_var.get()),
            "permissions": {key: bool(var.get()) for key, var in self.permission_vars.items()},
            "alerts": alerts,
        }
        if not HAS_WINPTY:
            payload["permissions"]["terminal"] = False
        return payload

    def save(self) -> bool:
        try:
            payload = self._collect_payload()
            if not payload["access_code"]:
                raise ValueError("access code is required")
            if len(payload["access_code"]) > ACCESS_CODE_MAX_LENGTH:
                raise ValueError(f"access code must be {ACCESS_CODE_MAX_LENGTH} characters or fewer")
            self.store.apply_setup(payload, require_code=payload.get("admin_code"))
            self.apply_startup(show_status=False)
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

    def _save_cloud_fields_only(self) -> None:
        data = self.store.load(reload=True)
        data["device_name"] = self.device_name.get().strip() or data.get("device_name") or "Windows PC"
        self.store.save(data)
        self.data = data

    def start_activation(self) -> None:
        try:
            # Auto-populate device name from hostname if not set
            import platform as _platform
            data = self.store.load(reload=True)
            if not data.get("device_name"):
                data["device_name"] = _platform.node() or "Windows PC"
                self.store.save(data)
                self.data = data
            self._save_cloud_fields_only()
            self.welcome_status_label.configure(text="Opening browser...")
            result = self.cloud.start_activation()
            self.reload()
            if result.get("success"):
                connect_url = str(result.get("connectUrl") or (self.data.get("activation") or {}).get("connect_url") or "")
                if connect_url:
                    webbrowser.open(connect_url)
                self._activation_poll_deadline = time.time() + 900
                self._schedule_activation_poll()
                self.welcome_status_label.configure(text="Browser opened — please sign in then come back here.")
            else:
                err = str(result.get("error") or "Activation failed")
                self.welcome_status_label.configure(text=f"Error: {err}")
        except Exception as exc:
            self.welcome_status_label.configure(text=str(exc) or "Activation failed")

    def _schedule_activation_poll(self, delay_ms: int = 2500) -> None:
        try:
            if self._activation_poll_after:
                self.root.after_cancel(self._activation_poll_after)
        except Exception:
            pass
        self._activation_poll_after = self.root.after(delay_ms, self._poll_activation_auto)

    def _poll_activation_auto(self) -> None:
        self._activation_poll_after = None
        self.poll_activation(auto=True)

    def poll_activation(self, auto: bool = False) -> None:
        try:
            result = self.cloud.poll_activation()
            if result.get("status") == "claimed":
                self._activation_poll_deadline = 0.0
                # Fetch profile + credits then do one final reload
                import threading
                def _post_signin():
                    try:
                        self.cloud.fetch_profile()
                        self.cloud.fetch_credits()
                    except Exception:
                        pass
                    self.root.after(0, lambda: (self.reload(), self._set_status("Signed in ✓")))
                threading.Thread(target=_post_signin, daemon=True).start()
            elif result.get("status") == "pending":
                if not auto:
                    self.welcome_status_label.configure(text="Still waiting — please complete sign-in in the browser.")
                if auto and time.time() < self._activation_poll_deadline:
                    self._schedule_activation_poll()
            else:
                err = str(result.get("error") or "Activation check failed")
                if not auto:
                    self.welcome_status_label.configure(text=f"Error: {err}")
        except Exception as exc:
            if not auto:
                self.welcome_status_label.configure(text=str(exc) or "Check failed")


    def refresh_entitlement(self) -> None:
        try:
            result = self.cloud.entitlement()
            self.reload()
            if result.get("success"):
                entitlement = result.get("entitlement") or {}
                self._set_status("Entitlement refreshed" if entitlement.get("allowed") else str(entitlement.get("reason") or "Billing blocked"))
            else:
                self._set_status(str(result.get("error") or "Entitlement refresh failed"))
        except Exception as exc:
            self._set_status(str(exc) or "Entitlement refresh failed")

    def sign_out(self) -> None:
        """Clear device token and all cached user data, return to welcome screen."""
        if not messagebox.askyesno(
            "Sign Out",
            "Are you sure you want to sign out?\n\nThis will unlink this device from your Zadoo account.",
            icon="warning",
        ):
            return
        try:
            # Clear token + profile cache
            self.store.clear_device_token()
            data = self.store.load(reload=True)
            data["user_name"] = ""
            data["user_email"] = ""
            data["user_image_url"] = ""
            data["credits_cache"] = {}
            data["entitlement_cache"] = {}
            data["activation"] = {}
            self.store.save(data)
        except Exception as exc:
            self._set_status(f"Sign out error: {exc}")
            return
        self.reload()

    def copy_activation_code(self) -> None:
        code = self.activation_code.get().strip()
        if not code:
            self._set_status("No activation code")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(code)
        self._set_status("Sign-in code copied")

    def apply_startup(self, *, show_status: bool = True) -> None:
        data = self.store.load(reload=True)
        data["autostart_enabled"] = bool(self.autostart_var.get())
        data["show_settings_in_taskbar"] = bool(self.show_taskbar_var.get())
        self.store.save(data)
        ok, message = set_startup_task(bool(self.autostart_var.get()))
        if show_status:
            self._set_status(message if ok else f"Startup not changed: {message}")
        try:
            self.runtime_state.configure(text="Startup task installed" if startup_task_exists() else "Startup task not installed")
        except Exception:
            pass

    def stop_zadoo(self) -> None:
        if not _local_server_running():
            self._set_status("Zadoo runtime is not running")
            return
        try:
            payload = {"admin_code": self.saved_access_code or self.access_code.get().strip()}
            result = _post_local_json("/api/runtime/stop", payload)
            self._set_status(str(result.get("message") or "Zadoo runtime stopping"))
        except Exception as exc:
            self._set_status(str(exc) or "Could not stop Zadoo")

    def start_zadoo(self) -> None:
        if not self.store.get_device_token():
            self._set_status("Sign in to Zadoo first")
            return
        # Try to check entitlement but don't block start if the check itself fails
        try:
            result = self.cloud.entitlement()
            entitlement = (result.get("entitlement") if isinstance(result, dict) else None) or {}
            self.reload()  # refresh UI with latest entitlement
            if entitlement.get("revoked"):
                self._set_status(str(entitlement.get("reason") or "Device is revoked"))
                return
            # Hard-block only when the cloud RESPONDED and says we're not allowed
            # (a successful response with allowed:false = genuinely blocked).
            if entitlement.get("allowed") is False and result.get("success"):
                self._set_status(str(entitlement.get("reason") or result.get("error") or "Billing blocked"))
                return
        except Exception:
            pass  # Don't block start on network error — let the runtime handle it
        # If a (possibly stale) runtime is already running, restart it GRACEFULLY so a
        # fresh tunnel/public link is created. We use the graceful stop endpoint (not
        # taskkill /T), so this Settings window is never tree-killed.
        if _local_server_running():
            self.public_url_label.configure(text="Restarting Zadoo…", foreground=MUTED)
            self._set_status("Restarting Zadoo to refresh the public link…")
            try:
                _post_local_json("/api/runtime/stop", {"admin_code": self._admin_code()})
            except Exception:
                pass
            self.root.after(3000, self._launch_runtime)
            return
        self._launch_runtime()

    def _launch_runtime(self) -> None:
        try:
            subprocess.Popen(
                _runtime_command(),
                cwd=str(PROJECT_DIR),
                close_fds=True,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            # Keep Settings open (do not hide). Show the public link as the tunnel comes up.
            self.public_url_label.configure(text="Starting Zadoo…", foreground=MUTED)
            self._set_status("Starting Zadoo… the public link will appear here shortly.")
            self.root.after(3500, lambda: self._begin_public_url_autopoll(20))
        except Exception as exc:
            messagebox.showerror("Zadoo", f"Could not start Zadoo: {exc}")

    def _on_unmap(self, event) -> None:
        if event.widget is self.root and self.root.state() == "iconic" and not bool(self.show_taskbar_var.get()):
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
