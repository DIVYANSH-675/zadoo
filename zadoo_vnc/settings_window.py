"""Small native settings window for installed Zadoo builds."""
from __future__ import annotations

import json
import logging
import os
import subprocess
import sys
import threading
import time
import tkinter as tk
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
from contextlib import suppress
from tkinter import BOTH, END, LEFT, RIGHT, StringVar, Tk, messagebox, ttk

from .config import APP_PORT, PROJECT_DIR, resource_path
from .saas import ZadooCloudClient
from .settings import (
    ACCESS_CODE_MAX_LENGTH,
    PERMISSION_KEYS,
    SettingsStore,
    get_settings_store,
)
from .windows_startup import set_startup_task, startup_task_exists

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
        with urllib.request.urlopen(_local_url("/api/runtime/status"), timeout=0.6) as response:
            return response.status == 200
    except OSError:
        return False


def _get_local_json(path: str, payload: dict, timeout: float = 1.5) -> dict:
    code = str(payload.get("admin_code", ""))
    extras = {k: v for k, v in payload.items() if k != "admin_code" and v is not None}
    url = _local_url(path)
    if extras:
        url += ("&" if "?" in url else "?") + urllib.parse.urlencode(extras)
    request = urllib.request.Request(
        url,
        headers={"X-Zadoo-Code": code},
        method="GET",
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            status = response.status
            raw = response.read()
    except urllib.error.HTTPError as exc:
        status = exc.code
        raw = exc.read()
    try:
        data = json.loads(raw.decode("utf-8"))
    except (json.JSONDecodeError, UnicodeDecodeError) as exc:
        raise RuntimeError(f"Local API {path} returned invalid JSON (HTTP {status}): {exc}") from exc
    if not isinstance(data, dict):
        raise RuntimeError(f"Local API {path} returned a non-object response (HTTP {status})")
    if type(data.get("success")) is not bool:
        raise RuntimeError(f"Local API {path} response is missing boolean success (HTTP {status})")
    if status >= 400 and data["success"]:
        raise RuntimeError(f"Local API {path} returned success=true with HTTP {status}")
    if not data["success"] and (not isinstance(data.get("error"), str) or not data["error"].strip()):
        raise RuntimeError(f"Local API {path} failure is missing error (HTTP {status})")
    return data


def _notify_settings_reload(access_code: str) -> None:
    result = _get_local_json("/api/settings/reload", {"admin_code": access_code}, timeout=1.2)
    if result["success"] is not True:
        raise RuntimeError(result["error"])


def _find_icon() -> str:
    path = resource_path("app_icon.ico")
    if not path.is_file():
        raise FileNotFoundError(f"Zadoo icon not found: {path}")
    return str(path)


class ZadooSettingsWindow:
    def __init__(self, store: SettingsStore | None = None):
        self.store = store or get_settings_store()
        self.cloud = ZadooCloudClient(self.store)
        # Make the process DPI-aware BEFORE creating Tk so the window renders crisply and at the
        # right size on high-DPI / scaled displays instead of being bitmap-stretched.
        from .dpi import ensure_process_dpi_aware_once
        ensure_process_dpi_aware_once()
        self.root = Tk()
        self.root.title("Zadoo Settings")
        # Match Tk scaling to the monitor DPI (tk scaling = DPI/72) so fonts and widgets are
        # sized correctly for every device, and remember the ratio for window sizing below.
        dpi = float(self.root.winfo_fpixels("1i"))
        if dpi <= 0:
            raise RuntimeError(f"Tk reported an invalid display DPI: {dpi}")
        self._ui_scale = max(1.0, dpi / 96.0)
        self.root.tk.call("tk", "scaling", dpi / 72.0)
        self.root.minsize(int(420 * self._ui_scale), int(320 * self._ui_scale))
        self._fit_geometry(520, 380)
        self.root.configure(bg=BG)
        self.root.protocol("WM_DELETE_WINDOW", self.hide)
        self.root.bind("<Unmap>", self._on_unmap)
        self._hidden = False
        self.root.iconbitmap(_find_icon())

        self.data = {}
        self.saved_access_code = ""
        self.alert_vars: dict[str, dict[str, tk.Variable]] = {}
        self.permission_vars: dict[str, tk.BooleanVar] = {}
        self.autostart_var = tk.BooleanVar(value=True)
        self.show_taskbar_var = tk.BooleanVar(value=False)
        self._activation_poll_after: str | None = None
        self._activation_poll_deadline = 0.0
        self._activation_request_active = False
        self._start_request_active = False
        self._signout_active = False
        # Live runtime state (kept fresh by a background poll so the public-link label
        # updates automatically when Zadoo starts/stops — without blocking the UI).
        self._runtime_running = False
        self._runtime_url = ""
        self._last_focus_refresh = 0.0
        self._startup_status_loaded = False
        self._autosave_after: str | None = None
        self._loading = False
        self._build_style()
        self._build_ui()
        self.reload()
        self._wire_autosave()  # settings persist automatically (no Save button)
        self._poll_runtime_status()  # start the live public-link / running-state poll
        self.root.after(60_000, self._poll_account_balance)
        # Refresh balance the moment the window regains focus (e.g. back from paying).
        self.root.bind("<FocusIn>", lambda _e: self._refresh_balance_now())
        if self.store.get_device_token():
            self._last_focus_refresh = time.time()
            self._fetch_latest_account_details_async(show_status=False)

    def _build_style(self) -> None:
        style = ttk.Style(self.root)
        style.theme_use("clam")
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
        # Settings autosave automatically — no Save button. Keep only Start/Stop here.
        ttk.Button(self.footer, text="Stop", command=self.stop_zadoo).pack(side=RIGHT, padx=(6, 0))
        ttk.Button(self.footer, text="Start Zadoo", command=self.start_zadoo, style="Primary.TButton").pack(side=RIGHT, padx=(6, 0))

    def _fit_geometry(self, base_w: int, base_h: int) -> None:
        """Scale the requested size to the monitor DPI, then clamp to the screen and centre so
        the window fits and looks right on every device — no scrollbars needed."""
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        if min(sw, sh) <= 0:
            raise RuntimeError(f"Tk reported an invalid screen size: {sw}x{sh}")
        w = max(int(360 * self._ui_scale), min(int(base_w * self._ui_scale), sw - 40))
        h = max(int(300 * self._ui_scale), min(int(base_h * self._ui_scale), sh - 96))
        x = max(0, (sw - w) // 2)
        y = max(0, (sh - h) // 3)
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def _tab_body(self, parent):
        """Return a padded content frame filling `parent`. No scrollbars — the window itself
        auto-fits the screen and scales to the monitor DPI, so content stays right-sized and
        visible on every device."""
        inner = ttk.Frame(parent, padding=16, style="Surface.TFrame")
        inner.pack(fill=BOTH, expand=True)
        return inner

    def _access_code_validator(self, value: str) -> bool:
        return len(value or "") <= ACCESS_CODE_MAX_LENGTH

    def _build_access_tab(self) -> None:
        self.tab_access = ttk.Frame(self.tabs, style="Surface.TFrame")
        self.tabs.add(self.tab_access, text="Access")
        body = self._tab_body(self.tab_access)
        validator = (self.root.register(self._access_code_validator), "%P")

        form = ttk.Frame(body, style="Surface.TFrame")
        form.pack(fill="x")
        ttk.Label(form, text=f"Access code ({ACCESS_CODE_MAX_LENGTH} chars max)", style="Surface.TLabel").grid(row=0, column=0, sticky="w")
        self.access_code = ttk.Entry(form, validate="key", validatecommand=validator)
        self.access_code.grid(row=1, column=0, sticky="ew", pady=(2, 10))

        ttk.Label(form, text="Email To", style="Surface.TLabel").grid(row=2, column=0, sticky="w")
        self.email_to = ttk.Entry(form)
        self.email_to.grid(row=3, column=0, sticky="ew", pady=(2, 10))

        self.email_state = ttk.Label(form, text="", style="Surface.TLabel")
        self.email_state.grid(row=4, column=0, sticky="w")
        form.columnconfigure(0, weight=1)

        # ── Remote access (public link + start) ───────────────────────
        self.public_url_frame = ttk.Frame(body, style="Surface.TFrame")
        self.public_url_frame.pack(fill="x", pady=(18, 0))
        ttk.Label(self.public_url_frame, text="Public link", style="Surface.TLabel").pack(anchor="w")
        url_row = ttk.Frame(self.public_url_frame, style="Surface.TFrame")
        url_row.pack(fill="x", pady=(2, 6))
        self.public_url_label = ttk.Label(url_row, text="Zadoo not running", style="Muted.TLabel",
                                          font=("Segoe UI", 9), cursor="hand2")
        self.public_url_label.pack(side=LEFT, fill="x", expand=True)
        self.public_url_label.bind("<Button-1>", lambda _e: self._open_public_url())
        # Start/Stop live in the always-visible footer; keep only Refresh inline with the link.
        ttk.Button(url_row, text="Refresh", command=self._refresh_public_url).pack(side=LEFT, padx=(6, 0))
        ttk.Label(self.public_url_frame,
                  text="Click Start to launch Zadoo — the public link appears here, then click it to open.",
                  style="Muted.TLabel", font=("Segoe UI", 8)).pack(anchor="w")

    def _build_account_tab(self) -> None:
        self.tab_account = ttk.Frame(self.tabs, style="Surface.TFrame")
        self.tabs.add(self.tab_account, text="Account")
        body = self._tab_body(self.tab_account)

        # ── Profile card ──────────────────────────────────────────────
        self.profile_card = ttk.LabelFrame(body, text="Signed-in account", style="Card.TLabelframe")

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
        self.device_name_frame = ttk.Frame(body, style="Surface.TFrame")
        ttk.Label(self.device_name_frame, text="Device name", style="Surface.TLabel").pack(anchor="w")
        self.device_name = ttk.Entry(self.device_name_frame)
        self.device_name.pack(fill="x", pady=(2, 10))

        # ── Sign-in Container ─────────────────────────────────────────
        self.signin_frame = ttk.Frame(body, style="Surface.TFrame")
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
        self.credits_card = ttk.LabelFrame(body, text="Credits", style="Card.TLabelframe")
        # Icon-only refresh (loads latest balance) floated at the top-right so it does NOT
        # push the credit lines down / disturb the layout.
        ttk.Button(self.credits_card, text="⟳", width=3, command=self._refresh_credits).place(relx=1.0, x=-4, y=2, anchor="ne")
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
        ttk.Button(credits_btns, text="+ Add Credit", command=self._add_balance,
                   style="Primary.TButton").pack(side=LEFT)
        ttk.Button(credits_btns, text="View Plans →", command=self._open_pricing).pack(side=LEFT, padx=(8, 0))
        ttk.Button(credits_btns, text="Copy Sign-in Code", command=self.copy_activation_code).pack(side=LEFT, padx=(8, 0))

    def _build_runtime_tab(self) -> None:
        self.tab_runtime = ttk.Frame(self.tabs, style="Surface.TFrame")
        self.tabs.add(self.tab_runtime, text="Runtime")
        body = self._tab_body(self.tab_runtime)
        card = ttk.LabelFrame(body, text="Windows behavior", style="Card.TLabelframe")
        card.pack(fill="x")
        self._check_button(card, "Start Zadoo when Windows starts", self.autostart_var).grid(row=0, column=0, sticky="w", pady=4)
        self._check_button(card, "Keep Settings visible on taskbar when minimized", self.show_taskbar_var).grid(row=1, column=0, sticky="w", pady=4)
        ttk.Button(card, text="Apply Startup", command=self.apply_startup).grid(row=2, column=0, sticky="w", pady=(12, 0))
        self.runtime_state = ttk.Label(body, text="", style="Surface.TLabel")
        self.runtime_state.pack(anchor="w", pady=(14, 0))

    def _check_button(self, parent, text, var):
        return ttk.Checkbutton(parent, text=text, variable=var)

    def _build_permissions_tab(self) -> None:
        self.tab_permissions = ttk.Frame(self.tabs, style="Surface.TFrame")
        self.tabs.add(self.tab_permissions, text="Permissions")
        body = self._tab_body(self.tab_permissions)

        perms = ttk.LabelFrame(body, text="Allowed controls for this password", style="Card.TLabelframe")
        perms.pack(fill=BOTH, expand=True)
        for index, key in enumerate(PERMISSION_KEYS):
            var = tk.BooleanVar(value=False)
            self.permission_vars[key] = var
            button = self._check_button(perms, PERMISSION_LABELS.get(key, key), var)
            button.grid(
                row=index // 3,
                column=index % 3,
                sticky="w",
                padx=10,
                pady=6,
            )

        actions = ttk.Frame(body, style="Surface.TFrame")
        actions.pack(fill="x", pady=(10, 0))
        ttk.Button(actions, text="Allow All", command=self.allow_all_permissions).pack(side=LEFT)
        ttk.Button(actions, text="Clear All", command=self.clear_permissions).pack(side=LEFT, padx=(6, 0))

    def _build_alert_tab(self) -> None:
        self.tab_alerts = ttk.Frame(self.tabs, style="Surface.TFrame")
        self.tabs.add(self.tab_alerts, text="Alerts")
        body = self._tab_body(self.tab_alerts)
        alert_tabs = ttk.Notebook(body)
        alert_tabs.pack(fill=BOTH, expand=True)
        for code in ("A", "B", "C", "D"):
            frame = ttk.Frame(alert_tabs, padding=16, style="Surface.TFrame")
            alert_tabs.add(frame, text=f"Alert {code}")
            enabled = tk.BooleanVar(value=False)
            title = StringVar(value="")
            message = StringVar(value="")
            self.alert_vars[code] = {"enabled": enabled, "title": title, "message": message}
            self._check_button(frame, "Enabled", enabled).pack(anchor="w", pady=(0, 10))
            ttk.Label(frame, text="Title", style="Surface.TLabel").pack(anchor="w")
            ttk.Entry(frame, textvariable=title).pack(fill="x", pady=(2, 8))
            ttk.Label(frame, text="Message", style="Surface.TLabel").pack(anchor="w")
            ttk.Entry(frame, textvariable=message).pack(fill="x", pady=(2, 8))

    def _set_status(self, text: str) -> None:
        self.status_label.configure(text=text)

    def _setup_complete(self) -> bool:
        return self.data["setup_complete"]

    def reload(self) -> None:
        self._loading = True  # suppress autosave while we populate fields programmatically
        try:
            self.data = self.store.load(reload=True)
            self.saved_access_code = self.store.get_access_code()
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
        finally:
            self._loading = False

    def _load_access_values(self) -> None:
        self.access_code.delete(0, END)
        self.access_code.insert(0, self.saved_access_code)
        email_to = self.data["email_to"].strip()
        self.email_to.delete(0, END)
        self.email_to.insert(0, email_to)
        missing = []
        if not os.getenv("RESEND_API_KEY", "").strip():
            missing.append("RESEND_API_KEY")
        if not os.getenv("RESEND_FROM", "").strip():
            missing.append("RESEND_FROM")
        if not email_to:
            missing.append("Email To")
        self.email_state.configure(text=f"Missing: {', '.join(missing)}" if missing else "Email configured")

    def _load_account_values(self) -> None:
        # Check signed in state
        is_signed_in = bool(self.store.get_device_token())

        # Adjust window controls, tabs, and footer visibility dynamically based on sign in status
        if is_signed_in:
            self.signin_welcome_frame.pack_forget()
            # Pack the footer FIRST at the bottom so Start/Stop stay visible on any screen
            # even when the tab area is squeezed; the tabs then fill the space above it.
            self.footer.pack(side="bottom", fill="x")
            self.tabs.pack(fill=BOTH, expand=True, pady=(10, 8))

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

            _s = self._ui_scale
            self.root.minsize(int(480 * _s), int(360 * _s))
            if self.root.winfo_width() < int(660 * _s):
                self._fit_geometry(780, 560)
        else:
            self.tabs.pack_forget()
            self.footer.pack_forget()
            self.signin_welcome_frame.pack(fill=BOTH, expand=True, pady=(10, 8))

            _s = self._ui_scale
            self.root.minsize(int(420 * _s), int(320 * _s))
            if self.root.winfo_width() > int(560 * _s):
                self._fit_geometry(520, 380)

            # Check if there is an active activation process running
            activation = self.data["activation"]
            code = str(activation.get("code") or "")
            if code:
                self.welcome_status_label.configure(text="Waiting for browser sign-in approval...")
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
        name = self.data["user_name"].strip()
        email = self.data["user_email"].strip()
        workspace = self.data["workspace_id"]
        device = self.data["device_id"]

        if is_signed_in:
            display_name = name or email or "Active Account"
            self.profile_name_label.configure(text=display_name)
            self.profile_email_label.configure(text=email if name else "")
            initials = "".join(part[0].upper() for part in display_name.split() if part)[:2] or "?"
            self.avatar_label.configure(text=initials, image="")
        else:
            self.profile_name_label.configure(text="Not signed in")
            self.profile_email_label.configure(text="")
            self.avatar_label.configure(text="?", image="")

        self.account_state.configure(
            text=(f"Workspace {workspace[:8]}  ·  Device {device[:8]}" if workspace and device else "")
        )

        # Device name field
        self.device_name.delete(0, END)
        self.device_name.insert(0, self.data["device_name"])

        # Activation code
        activation = self.data["activation"]
        code = str(activation.get("code") or "")
        self.activation_code.configure(state="normal")
        self.activation_code.delete(0, END)
        self.activation_code.insert(0, code)
        self.activation_code.configure(state="readonly")

        # Public URL label — show the live link ONLY while Zadoo is actually running; a
        # stored URL from a previous run is stale. The background poll keeps this fresh.
        if self._runtime_running and self._runtime_url:
            self.public_url_label.configure(text=self._runtime_url, foreground=ACCENT)
        elif self._runtime_running:
            self.public_url_label.configure(text="Starting Zadoo…", foreground=MUTED)
        else:
            self.public_url_label.configure(text="Zadoo not running", foreground=MUTED)

        credits = self.data["credits_cache"]
        entitlement = self.data["entitlement_cache"]
        included = credits["includedMinutesRemaining"] if credits else 0
        wallet = credits["walletMinutes"] if credits else 0
        total = credits["totalMinutesRemaining"] if credits else 0
        plan = credits["planCode"] if credits else None

        if credits:
            self.credits_included_label.configure(text=f"Included: {included} min")
            self.credits_wallet_label.configure(text=f"Wallet: {wallet} min")
            self.credits_total_label.configure(text=f"Total remaining: {total} min{(' · ' + plan) if plan else ''}")
        else:
            self.credits_included_label.configure(text="Credits not loaded")
            self.credits_wallet_label.configure(text="")
            self.credits_total_label.configure(text="")

        if entitlement:
            allowed = entitlement["allowed"]
            reason = entitlement["reason"]
            self.billing_state.configure(
                text=f"{'✓ Active' if allowed else '✗ Blocked'}{': ' + reason if reason else ''}",
                foreground=(ACCENT if allowed else DANGER)
            )
        else:
            self.billing_state.configure(text="", foreground=MUTED)

    def _admin_code(self) -> str:
        return self.saved_access_code or self.access_code.get().strip()

    def _open_public_url(self) -> None:
        # Clicking the public link opens it in the browser; if there's none yet, hint to Start.
        pub_url = self._runtime_url
        if pub_url:
            webbrowser.open(pub_url)
        elif not _local_server_running():
            self._set_status("Zadoo is not running — click Start first.")
        else:
            self._set_status("Tunnel not ready yet — click Refresh in a moment.")

    def _poll_runtime_status(self) -> None:
        """Background-poll the local runtime so the public-link label reflects whether
        Zadoo is actually running — auto-clearing a stale link the moment it's stopped."""
        def _check():
            running = False
            url = ""
            error = ""
            try:
                data = _get_local_json("/api/runtime/status", {}, timeout=1.2)
                if not data["success"]:
                    raise RuntimeError(data["error"])
                if not isinstance(data.get("running"), bool):
                    raise RuntimeError("Runtime status is missing running")
                if data.get("public_url") is not None and not isinstance(data["public_url"], str):
                    raise RuntimeError("Runtime status public_url must be a string or null")
                running = data["running"]
                url = data["public_url"] or ""
            except OSError:
                pass
            except Exception as exc:
                error = str(exc)
                logging.error("Runtime status failed: %s", error)
            with suppress(tk.TclError):
                self.root.after(0, lambda: self._apply_runtime_status(running, url, error))
        threading.Thread(target=_check, daemon=True).start()

    def _apply_runtime_status(self, running: bool, url: str, error: str) -> None:
        self._runtime_running = running
        self._runtime_url = url
        with suppress(tk.TclError):
            if error:
                self.public_url_label.configure(text=error, foreground=DANGER)
            elif running and url:
                self.public_url_label.configure(text=url, foreground=ACCENT)
            elif running:
                self.public_url_label.configure(text="Starting Zadoo…", foreground=MUTED)
            else:
                self.public_url_label.configure(text="Zadoo not running", foreground=MUTED)
        with suppress(tk.TclError):
            self.root.after(4000, self._poll_runtime_status)

    def _poll_account_balance(self) -> None:
        """Background-poll the cloud so the Credits card updates on its own (e.g. after a
        top-up) without the user having to click Refresh."""
        if self.store.get_device_token():
            self._refresh_account_in_background()
        with suppress(tk.TclError):
            self.root.after(60_000, self._poll_account_balance)

    def _refresh_account_in_background(self) -> None:
        def _check():
            for error in self._cloud_fetch_errors(
                (("credits", self.cloud.fetch_credits), ("entitlement", self.cloud.entitlement))
            ):
                logging.error("Cloud refresh failed: %s", error)
            with suppress(tk.TclError):
                self.root.after(0, self.reload)

        threading.Thread(target=_check, daemon=True).start()

    def _refresh_balance_now(self) -> None:
        """Immediate one-shot balance refresh (used when the window regains focus, e.g.
        returning from the billing page after paying)."""
        if not self.store.get_device_token():
            return
        now = time.time()
        if now - self._last_focus_refresh < 3.0:
            return  # debounce: FocusIn can fire repeatedly
        self._last_focus_refresh = now
        self._refresh_account_in_background()

    def _begin_public_url_autopoll(self, attempts: int = 8) -> None:
        """After Start, poll the runtime until the public link appears — or until the
        runtime tells us the tunnel is disabled (then show WHY instead of hanging)."""
        if attempts <= 0:
            return
        def _poll():
            url = ""
            block_reason = ""
            tunnel_error = ""
            tunnel_enabled = True
            running = False
            try:
                status = _get_local_json("/api/runtime/status", {}, timeout=3.0)
                if not status["success"]:
                    raise RuntimeError(status["error"])
                if not isinstance(status.get("running"), bool):
                    raise RuntimeError("Runtime status is missing running")
                if not isinstance(status.get("tunnel_enabled"), bool):
                    raise RuntimeError("Runtime status is missing tunnel_enabled")
                for field in ("public_url", "tunnel_error"):
                    if status.get(field) is not None and not isinstance(status[field], str):
                        raise RuntimeError(f"Runtime status {field} must be a string or null")
                if not isinstance(status.get("tunnel_block_reason"), str):
                    raise RuntimeError("Runtime status tunnel_block_reason must be a string")
                running = status["running"]
                url = (status["public_url"] or "").strip()
                tunnel_enabled = status["tunnel_enabled"]
                block_reason = status["tunnel_block_reason"].strip()
                tunnel_error = (status["tunnel_error"] or "").strip()
                if not tunnel_enabled and not block_reason:
                    raise RuntimeError("Runtime status is missing tunnel_block_reason")
            except (
                OSError,
                TimeoutError,
                urllib.error.URLError,
                json.JSONDecodeError,
                UnicodeDecodeError,
                RuntimeError,
            ) as exc:
                tunnel_error = str(exc)
            def _update():
                if url:
                    self._runtime_running = True
                    self._runtime_url = url
                    self.public_url_label.configure(text=url, foreground=ACCENT)
                    self._set_status("Public link ready — click it to open.")
                elif running and tunnel_error:
                    self.public_url_label.configure(text=tunnel_error, foreground=DANGER)
                    self._set_status(tunnel_error)
                elif running and not tunnel_enabled:
                    # Runtime is up but the tunnel is off (billing/sign-in/etc) — stop hanging.
                    self.public_url_label.configure(text=block_reason, foreground=DANGER)
                    self._set_status(block_reason + "  Fix it, then click Start again.")
                elif attempts <= 1 and tunnel_error:
                    self.public_url_label.configure(text=tunnel_error, foreground=DANGER)
                    self._set_status(tunnel_error)
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
        def _do_refresh():
            pub_url = ""
            error = ""
            try:
                result = _get_local_json(
                    "/api/runtime/refresh-tunnel",
                    {"admin_code": self._admin_code()},
                    timeout=30.0,
                )
                if result["success"] is True:
                    pub_url = result.get("url")
                    if not isinstance(pub_url, str) or not pub_url.strip():
                        raise RuntimeError("Tunnel refresh response is missing url")
                    pub_url = pub_url.strip()
                else:
                    error = result["error"]
            except Exception as exc:
                error = str(exc)
            def _update():
                if pub_url:
                    self._runtime_running = True
                    self._runtime_url = pub_url
                    self.public_url_label.configure(text=pub_url, foreground=ACCENT)
                    self._set_status("New public URL ready.")
                else:
                    self.public_url_label.configure(text=error, foreground=DANGER)
                    self._set_status(error)
            self.root.after(0, _update)
        threading.Thread(target=_do_refresh, daemon=True).start()

    def _refresh_credits(self) -> None:
        self._fetch_latest_account_details_async(show_status=True)

    def _fetch_latest_account_details_async(self, show_status: bool = False) -> None:
        if show_status:
            self._set_status("Refreshing account info...")
        def _do_fetch():
            errors = self._cloud_fetch_errors(
                (("profile", self.cloud.fetch_profile), ("credits", self.cloud.fetch_credits))
            )
            def _done():
                try:
                    self.reload()
                    if show_status:
                        self._set_status("; ".join(errors) if errors else "Account info updated.")
                except Exception as exc:
                    self._set_status(f"Account reload failed: {exc}")
            self.root.after(0, _done)
        threading.Thread(target=_do_fetch, daemon=True).start()

    @staticmethod
    def _cloud_fetch_errors(fetches) -> list[str]:
        errors = []
        for label, fetch in fetches:
            try:
                result = fetch()
                if result["success"] is not True:
                    errors.append(f"{label}: {result['error']}")
            except Exception as exc:
                errors.append(f"{label}: {exc}")
        return errors

    def _open_pricing(self) -> None:
        base = self.data["cloud_api_base"].rstrip("/")
        webbrowser.open(f"{base}/pricing")

    def _add_balance(self) -> None:
        base = self.data["cloud_api_base"].rstrip("/")
        webbrowser.open(f"{base}/dashboard/billing")



    def _load_runtime_values(self) -> None:
        self.autostart_var.set(self.data["autostart_enabled"])
        self.show_taskbar_var.set(self.data["show_settings_in_taskbar"])
        if not self._startup_status_loaded:
            self.runtime_state.configure(
                text="Startup task installed" if startup_task_exists() else "Startup task not installed"
            )
            self._startup_status_loaded = True

    def _load_permission_values(self) -> None:
        permissions = self.data["permissions"]
        for key, var in self.permission_vars.items():
            var.set(permissions[key])

    def _load_alert_values(self) -> None:
        alerts = self.data["alerts"]
        for code, vars_for_code in self.alert_vars.items():
            item = alerts[code]
            vars_for_code["enabled"].set(item["enabled"])
            vars_for_code["title"].set(item["title"])
            vars_for_code["message"].set(item["message"])

    def allow_all_permissions(self) -> None:
        for var in self.permission_vars.values():
            var.set(True)

    def clear_permissions(self) -> None:
        for var in self.permission_vars.values():
            var.set(False)

    def _collect_payload(self) -> dict:
        alerts = {}
        for code, vars_for_code in self.alert_vars.items():
            title = str(vars_for_code["title"].get()).strip()
            message = str(vars_for_code["message"].get()).strip()
            alerts[code] = {
                "enabled": bool(vars_for_code["enabled"].get()) and bool(title or message),
                "title": title,
                "message": message,
            }
        return {
            "admin_code": self.saved_access_code or self.access_code.get().strip(),
            "access_code": self.access_code.get().strip(),
            "email_to": self.email_to.get().strip(),
            "device_name": self.device_name.get().strip(),
            "autostart_enabled": bool(self.autostart_var.get()),
            "show_settings_in_taskbar": bool(self.show_taskbar_var.get()),
            "permissions": {key: bool(var.get()) for key, var in self.permission_vars.items()},
            "alerts": alerts,
        }
    def _wire_autosave(self) -> None:
        """Persist settings automatically whenever a field changes (replaces the Save button)."""
        for entry in (self.access_code, self.email_to, self.device_name):
            entry.bind("<KeyRelease>", self._schedule_autosave, add="+")
            entry.bind("<FocusOut>", self._schedule_autosave, add="+")
        toggles = [self.autostart_var, self.show_taskbar_var]
        toggles.extend(self.permission_vars.values())
        for vars_for_code in self.alert_vars.values():
            toggles.extend(vars_for_code.values())
        for var in toggles:
            var.trace_add("write", self._schedule_autosave)

    def _schedule_autosave(self, *_args) -> None:
        if self._loading:
            return
        with suppress(tk.TclError):
            if self._autosave_after:
                self.root.after_cancel(self._autosave_after)
        self._autosave_after = self.root.after(700, self._autosave)

    def _autosave(self) -> None:
        self._autosave_after = None
        # Nothing to persist until the device is signed in (the welcome screen has no form).
        if not self.store.get_device_token():
            return
        try:
            payload = self._collect_payload()
        except Exception as exc:
            self._set_status(f"Save failed: {exc}")
            return
        code = str(payload.get("access_code") or "")
        # Wait for a complete, valid access code before writing — don't nag mid-typing.
        if not code or len(code) > ACCESS_CODE_MAX_LENGTH:
            return
        try:
            previous_autostart = self.data["autostart_enabled"]
            setup = {key: value for key, value in payload.items() if key != "admin_code"}
            self.data = self.store.apply_setup(setup, require_code=payload["admin_code"])
            desired_autostart = payload["autostart_enabled"]
            if desired_autostart != previous_autostart:
                ok, message = set_startup_task(desired_autostart)
                if not ok:
                    self.data = self.store.atomic_update(
                        lambda data: data.__setitem__("autostart_enabled", previous_autostart)
                    )
                    self._loading = True
                    try:
                        self.autostart_var.set(previous_autostart)
                    finally:
                        self._loading = False
                    raise RuntimeError(message)
                self.runtime_state.configure(
                    text="Startup task installed" if startup_task_exists() else "Startup task not installed"
                )
                self._startup_status_loaded = True
            self.saved_access_code = code
            reload_error = ""
            if _local_server_running():
                try:
                    _notify_settings_reload(code)
                except Exception as exc:
                    reload_error = str(exc)
                    logging.error("Runtime settings reload failed: %s", exc)
            self._set_status(
                f"Saved locally; runtime reload failed: {reload_error}"
                if reload_error else "Saved ✓"
            )
        except Exception as exc:
            self._set_status(f"Save failed: {exc}")

    def _save_cloud_fields_only(self) -> None:
        device_name = self.device_name.get().strip()
        if not device_name:
            raise ValueError("device name is required")
        self.data = self.store.atomic_update(lambda data: data.__setitem__("device_name", device_name))

    def start_activation(self) -> None:
        if self._activation_request_active:
            return
        try:
            self._save_cloud_fields_only()
            self.welcome_status_label.configure(text="Opening browser...")
        except Exception as exc:
            self.welcome_status_label.configure(text=str(exc))
            return

        self._activation_request_active = True

        def _start():
            try:
                result = self.cloud.start_activation()
            except Exception as exc:
                result = {"success": False, "error": str(exc)}

            def _done():
                self._activation_request_active = False
                self.reload()
                if result["success"]:
                    webbrowser.open(result["connectUrl"])
                    self._activation_poll_deadline = time.time() + 900
                    self._schedule_activation_poll()
                    self.welcome_status_label.configure(text="Browser opened — please sign in then come back here.")
                else:
                    self.welcome_status_label.configure(text=f"Error: {result['error']}")

            self.root.after(0, _done)

        threading.Thread(target=_start, daemon=True).start()

    def _schedule_activation_poll(self, delay_ms: int = 2500) -> None:
        with suppress(tk.TclError):
            if self._activation_poll_after:
                self.root.after_cancel(self._activation_poll_after)
        self._activation_poll_after = self.root.after(delay_ms, self._poll_activation_auto)

    def _poll_activation_auto(self) -> None:
        self._activation_poll_after = None
        self.poll_activation(auto=True)

    def poll_activation(self, auto: bool = False) -> None:
        if self._activation_request_active:
            return
        self._activation_request_active = True

        def _poll():
            errors = []
            try:
                result = self.cloud.poll_activation()
            except Exception as exc:
                result = {"success": False, "error": str(exc)}
            if result["success"] and result["status"] == "claimed":
                try:
                    def _wipe(data):
                        data["entitlement_cache"] = {}
                        data["credits_cache"] = {}
                    self.store.atomic_update(_wipe)
                except Exception as exc:
                    errors.append(f"local account reset: {exc}")
                for label, fetch in (
                    ("profile", self.cloud.fetch_profile),
                    ("credits", self.cloud.fetch_credits),
                    ("entitlement", self.cloud.entitlement),
                ):
                    try:
                        fetched = fetch()
                        if not fetched["success"]:
                            errors.append(f"{label}: {fetched['error']}")
                    except Exception as exc:
                        errors.append(f"{label}: {exc}")

            def _done():
                self._activation_request_active = False
                if result["success"] and result["status"] == "claimed":
                    self._activation_poll_deadline = 0.0
                    self.reload()
                    self._set_status("; ".join(errors) if errors else "Signed in ✓")
                elif result["success"] and result["status"] == "pending":
                    if not auto:
                        self.welcome_status_label.configure(text="Still waiting — please complete sign-in in the browser.")
                    if auto and time.time() < self._activation_poll_deadline:
                        self._schedule_activation_poll()
                else:
                    self.welcome_status_label.configure(text=f"Error: {result['error']}")

            self.root.after(0, _done)

        threading.Thread(target=_poll, daemon=True).start()

    def sign_out(self) -> None:
        """Stop Zadoo immediately, mark the device offline, and clear all account data."""
        if self._signout_active:
            return
        if not messagebox.askyesno(
            "Sign Out",
            "Are you sure you want to sign out?\n\nThis will stop Zadoo and unlink this device from your Zadoo account.",
            icon="warning",
        ):
            return
        self._signout_active = True
        self._set_status("Signing out...")
        admin_code = self._admin_code()

        def _sign_out():
            errors = []
            try:
                result = self.cloud.go_offline()
                if not result["success"]:
                    errors.append(f"cloud offline: {result['error']}")
            except Exception as exc:
                errors.append(f"cloud offline: {exc}")
            runtime_running = _local_server_running()
            if runtime_running:
                try:
                    result = _get_local_json("/api/runtime/stop", {"admin_code": admin_code})
                    if not result["success"]:
                        errors.append(f"runtime stop: {result['error']}")
                except Exception as exc:
                    errors.append(f"runtime stop: {exc}")
                deadline = time.monotonic() + 5
                while _local_server_running() and time.monotonic() < deadline:
                    time.sleep(0.25)
                if _local_server_running():
                    errors.append("runtime stop: port 6173 remained active after 5 seconds")
            try:
                self.store.clear_account()
            except Exception as exc:
                errors.append(f"local account clear: {exc}")

            def _done():
                self._signout_active = False
                self.reload()
                self._set_status("; ".join(errors) if errors else "Signed out — Zadoo stopped")

            self.root.after(0, _done)

        threading.Thread(target=_sign_out, daemon=True).start()

    def copy_activation_code(self) -> None:
        code = self.activation_code.get().strip()
        if not code:
            self._set_status("No activation code")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(code)
        self._set_status("Sign-in code copied")

    def apply_startup(self, *, show_status: bool = True) -> None:
        desired = bool(self.autostart_var.get())
        ok, message = set_startup_task(desired)
        if not ok:
            if show_status:
                self._set_status(f"Startup not changed: {message}")
            return
        show_taskbar = bool(self.show_taskbar_var.get())
        self.data = self.store.atomic_update(
            lambda data: data.update(
                autostart_enabled=desired,
                show_settings_in_taskbar=show_taskbar,
            )
        )
        if show_status:
            self._set_status(message)
        self.runtime_state.configure(text="Startup task installed" if startup_task_exists() else "Startup task not installed")
        self._startup_status_loaded = True

    def stop_zadoo(self) -> None:
        if not _local_server_running():
            self._set_status("Zadoo runtime is not running")
            return
        try:
            payload = {"admin_code": self.saved_access_code or self.access_code.get().strip()}
            result = _get_local_json("/api/runtime/stop", payload)
            if not result["success"]:
                raise RuntimeError(result["error"])
            if not isinstance(result.get("message"), str) or not result["message"].strip():
                raise RuntimeError("Runtime stop response is missing message")
            self._set_status(result["message"])
        except Exception as exc:
            self._set_status(str(exc))

    def start_zadoo(self) -> None:
        if self._start_request_active:
            return
        if not self.store.get_device_token():
            self._set_status("Sign in to Zadoo first")
            return
        code = self.access_code.get().strip()
        if not code:
            self._set_status("Access code is required")
            return
        try:
            payload = self._collect_payload()
            setup = {key: value for key, value in payload.items() if key != "admin_code"}
            self.store.apply_setup(setup, require_code=payload["admin_code"])
            self.saved_access_code = code
        except Exception as exc:
            self._set_status(str(exc))
            return
        admin_code = self._admin_code()
        self._start_request_active = True
        self._set_status("Checking cloud entitlement...")

        def _check_and_stop():
            error = ""
            restarted = False
            try:
                result = self.cloud.entitlement()
                if not result["success"]:
                    raise RuntimeError(f"Cloud entitlement check failed: {result['error']}")
                entitlement = result["entitlement"]
                if entitlement["revoked"]:
                    raise RuntimeError(entitlement["reason"])
                if _local_server_running():
                    stopped = _get_local_json("/api/runtime/stop", {"admin_code": admin_code})
                    if not stopped["success"]:
                        raise RuntimeError(stopped["error"])
                    deadline = time.monotonic() + 5
                    while _local_server_running() and time.monotonic() < deadline:
                        time.sleep(0.1)
                    if _local_server_running():
                        raise RuntimeError("Zadoo runtime did not stop within 5 seconds")
                    restarted = True
            except Exception as exc:
                error = str(exc)

            def _done():
                self._start_request_active = False
                self.reload()
                if error:
                    self._set_status(error)
                elif restarted:
                    self.public_url_label.configure(text="Restarting Zadoo…", foreground=MUTED)
                    self._set_status("Restarting Zadoo to refresh the public link…")
                    self._launch_runtime()
                else:
                    self._launch_runtime()

            self.root.after(0, _done)

        threading.Thread(target=_check_and_stop, daemon=True).start()

    def _launch_runtime(self) -> None:
        try:
            subprocess.Popen(
                [sys.executable] if getattr(sys, "frozen", False) else [sys.executable, "-m", "zadoo_vnc"],
                cwd=str(PROJECT_DIR),
                close_fds=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            # Keep Settings open (do not hide). Show the public link as the tunnel comes up.
            self.public_url_label.configure(text="Starting Zadoo…", foreground=MUTED)
            self._set_status("Starting Zadoo… the public link will appear here shortly.")
            self.root.after(1500, lambda: self._begin_public_url_autopoll(20))
        except Exception as exc:
            messagebox.showerror("Zadoo", f"Could not start Zadoo: {exc}")

    def _on_unmap(self, event) -> None:
        if event.widget is self.root and self.root.state() == "iconic" and not bool(self.show_taskbar_var.get()):
            self.root.after(80, self.hide)

    def hide(self) -> None:
        if self._hidden:
            return
        self._hidden = True
        with suppress(tk.TclError):
            self.root.withdraw()
        self.root.after(50, self.root.quit)

    def run(self) -> None:
        self.root.mainloop()
        with suppress(tk.TclError):
            self.root.destroy()


def run_settings_window() -> None:
    ZadooSettingsWindow().run()
