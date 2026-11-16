"""Small native settings window for installed Zadoo builds."""
from __future__ import annotations

import ctypes
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
import uuid
import webbrowser
from contextlib import suppress
from pathlib import Path
from tkinter import BOTH, END, LEFT, RIGHT, StringVar, Tk, filedialog, messagebox, ttk

from websockets.sync.client import connect as websocket_connect

from .config import APP_PORT, PROJECT_DIR, resource_path
from .saas import ZadooCloudClient
from .settings import (
    ACCESS_CODE_MAX_LENGTH,
    PERMISSION_KEYS,
    SettingsStore,
    get_settings_store,
)
from .windows_startup import set_startup_task, startup_task_exists

BG = "#080e14"
SURFACE = "#101e29"
SURFACE_ALT = "#19323f"
TEXT = "#edf8ff"
MUTED = "#a2bac9"
BORDER = "#304d5e"
ACCENT = "#60dbea"
ACCENT_DARK = "#89e9f3"
DANGER = "#ffaaaa"
SUPPORT_UPI = "divyanshgupta675@oksbi"

PERMISSION_LABELS = {
    "mouse": "Mouse",
    "keyboard": "Keyboard",
    "clipboard_pull": "Read clipboard",
    "clipboard_push": "Write clipboard",
    "system_audio": "Computer audio",
    "mic": "Microphone",
    "camera": "Camera",
    "terminal": "Terminal",
    "snapshots": "Snapshots",
    "advanced_video": "Video controls",
    "tunnel_refresh": "Refresh public link",
    "remote_alerts": "Remote alerts",
}
PERMISSION_HELP = {
    "mouse": "Move the pointer and click on this PC.",
    "keyboard": "Type and use keyboard shortcuts on this PC.",
    "clipboard_pull": "Copy text from this PC to the viewer.",
    "clipboard_push": "Send text or images to this PC's clipboard.",
    "system_audio": "Listen to sound playing on this PC.",
    "mic": "Listen through this PC's microphone.",
    "camera": "View this PC's connected camera.",
    "terminal": "Run commands on this PC. Only for trusted viewers.",
    "snapshots": "Save an image of this PC's screen.",
    "advanced_video": "Choose a screen region, scale or grayscale.",
    "tunnel_refresh": "Replace the public link used to connect.",
    "remote_alerts": "Send the preset messages from the Alerts tab.",
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


def _local_rpc(action: str, payload: dict, timeout: float = 1.5) -> dict:
    if not isinstance(action, str) or not action:
        raise ValueError("Local RPC action is required")
    if not isinstance(payload, dict):
        raise TypeError("Local RPC payload must be an object")
    request_id = uuid.uuid4().hex
    request = json.dumps({"id": request_id, "action": action, "params": payload})
    uri = f"ws://127.0.0.1:{APP_PORT}/rpc"
    with websocket_connect(
        uri,
        origin=f"http://127.0.0.1:{APP_PORT}",
        open_timeout=min(timeout, 5.0),
        close_timeout=1.0,
        max_size=1_000_000,
        proxy=None,
    ) as websocket:
        websocket.send(request)
        raw = websocket.recv(timeout=timeout)
    if not isinstance(raw, str):
        raise RuntimeError(f"Local RPC {action} returned a non-text response")
    try:
        data = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"Local RPC {action} returned invalid JSON: {exc}") from exc
    if not isinstance(data, dict):
        raise RuntimeError(f"Local RPC {action} returned a non-object response")
    if data.get("id") != request_id:
        raise RuntimeError(f"Local RPC {action} returned a mismatched id")
    if type(data.get("success")) is not bool:
        raise RuntimeError(f"Local RPC {action} response is missing boolean success")
    if not data["success"]:
        error = data.get("error")
        if not isinstance(error, str) or not error.strip():
            raise RuntimeError(f"Local RPC {action} failure is missing error")
        return {"success": False, "error": error}
    result = data.get("result")
    if not isinstance(result, dict):
        raise RuntimeError(f"Local RPC {action} result must be an object")
    return {"success": True, **result}


def _notify_settings_reload(access_code: str) -> None:
    result = _local_rpc("settings.reload", {"admin_code": access_code}, timeout=1.2)
    if result["success"] is not True:
        raise RuntimeError(result["error"])


def _find_icon() -> str:
    path = resource_path("app_icon.ico")
    if not path.is_file():
        raise FileNotFoundError(f"Zadoo icon not found: {path}")
    return str(path)


def _windows_animations_enabled() -> bool:
    """Fail closed to a still illustration if the OS preference cannot be read."""
    enabled = ctypes.c_int(0)
    try:
        return bool(ctypes.windll.user32.SystemParametersInfoW(0x1042, 0, ctypes.byref(enabled), 0) and enabled.value)
    except (AttributeError, OSError):
        return False


class ZadooSettingsWindow:
    def __init__(self, store: SettingsStore | None = None):
        self.store = store or get_settings_store()
        self.cloud = ZadooCloudClient(self.store)
        # Make the process DPI-aware BEFORE creating Tk so the window renders crisply and at the
        # right size on high-DPI / scaled displays instead of being bitmap-stretched.
        from .dpi import ensure_process_dpi_aware_once
        ensure_process_dpi_aware_once()
        self.root = Tk()
        self.root.title("Zadoo · Configure")
        # Match Tk scaling to the monitor DPI (tk scaling = DPI/72) so fonts and widgets are
        # sized correctly for every device, and remember the ratio for window sizing below.
        dpi = float(self.root.winfo_fpixels("1i"))
        if dpi <= 0:
            raise RuntimeError(f"Tk reported an invalid display DPI: {dpi}")
        self._ui_scale = max(1.0, dpi / 96.0)
        self.root.tk.call("tk", "scaling", dpi / 72.0)
        self.root.minsize(min(int(480 * self._ui_scale), self.root.winfo_screenwidth() - 40),
                          min(int(360 * self._ui_scale), self.root.winfo_screenheight() - 96))
        self._fit_geometry(880, 700)
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
        self._runtime_status_result = None
        self._startup_status_loaded = False
        self._autosave_after: str | None = None
        self._loading = False
        self._scroll_canvases: list[tk.Canvas] = []
        self._support_window: tk.Toplevel | None = None
        self._motion_after: str | None = None
        self._motion_tick = 0
        self.motion_var = tk.BooleanVar(value=_windows_animations_enabled())
        self._build_style()
        self._build_ui()
        self.reload()
        self._wire_autosave()  # settings persist automatically (no Save button)
        self._poll_runtime_status()  # start the live public-link / running-state poll
        if self._dashboard_linked():
            self._fetch_latest_account_details_async(show_status=False)

    def _build_style(self) -> None:
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure(".", font=("Segoe UI", 10), background=BG, foreground=TEXT)
        style.configure("Root.TFrame", background=BG)
        style.configure("Surface.TFrame", background=SURFACE)
        style.configure("TLabel", background=BG, foreground=TEXT)
        style.configure("Surface.TLabel", background=SURFACE, foreground=TEXT)
        style.configure("Muted.TLabel", background=BG, foreground=MUTED)
        style.configure("SurfaceMuted.TLabel", background=SURFACE, foreground=MUTED)
        style.configure("Eyebrow.TLabel", background=SURFACE, foreground=ACCENT, font=("Segoe UI", 9, "bold"))
        style.configure("Section.TLabel", background=SURFACE, foreground=TEXT, font=("Segoe UI", 18, "bold"))
        style.configure("Status.TLabel", background=BG, foreground=MUTED)
        style.configure("Title.TLabel", background=BG, foreground=TEXT, font=("Segoe UI", 25, "bold"))
        style.configure(
            "TButton",
            background=SURFACE,
            foreground=TEXT,
            bordercolor=BORDER,
            lightcolor=BORDER,
            darkcolor=BORDER,
            borderwidth=1,
            focusthickness=1,
            focuscolor=BORDER,
            padding=(14, 8),
        )
        style.map("TButton", background=[("active", SURFACE_ALT)], foreground=[("disabled", MUTED)])
        style.configure(
            "Primary.TButton",
            background=ACCENT,
            foreground="#06212c",
            bordercolor=ACCENT,
            font=("Segoe UI", 10, "bold"),
            padding=(16, 8),
        )
        style.map("Primary.TButton", background=[("active", ACCENT_DARK), ("pressed", ACCENT_DARK)])
        style.configure("Nav.TButton", padding=(10, 10), background=BG, foreground=MUTED, borderwidth=0)
        style.map("Nav.TButton", background=[("active", SURFACE_ALT)])
        style.configure("SelectedNav.TButton", padding=(10, 10), background=SURFACE_ALT, foreground=ACCENT, bordercolor=BORDER)
        style.map("SelectedNav.TButton", background=[("active", SURFACE_ALT)])
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
        style.map("TEntry", fieldbackground=[("disabled", SURFACE_ALT), ("readonly", SURFACE_ALT)],
                  bordercolor=[("focus", ACCENT)], foreground=[("readonly", TEXT)])
        style.configure("TCheckbutton", background=SURFACE, foreground=TEXT, padding=(4, 4))
        style.map("TCheckbutton", background=[("active", SURFACE)], indicatorbackground=[("selected", ACCENT)])
        style.configure("Vertical.TScrollbar", background=SURFACE_ALT, troughcolor=SURFACE, borderwidth=0, arrowsize=13)
        style.configure("TNotebook", background=BG, bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER, borderwidth=0, tabmargins=(0, 8, 0, 0))
        style.configure("TNotebook.Tab", background=SURFACE_ALT, foreground=MUTED, padding=(18, 9), borderwidth=0)
        style.map(
            "TNotebook.Tab",
            background=[("selected", SURFACE), ("active", SURFACE_ALT)],
            foreground=[("selected", TEXT), ("active", TEXT)],
        )
        style.layout("Main.TNotebook.Tab", [])
        style.configure("Main.TNotebook", background=SURFACE, borderwidth=0, tabmargins=0)
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
            font=("Segoe UI", 10, "bold"),
        )

    def _build_ui(self) -> None:
        outer = ttk.Frame(self.root, padding=16, style="Root.TFrame")
        outer.pack(fill=BOTH, expand=True)

        header = ttk.Frame(outer, style="Root.TFrame")
        header.pack(fill="x")
        branding = ttk.Frame(header, style="Root.TFrame")
        branding.pack(side=LEFT, fill="x", expand=True)
        ttk.Label(branding, text="zadoo", style="Title.TLabel").pack(anchor="w")
        self.state_label = ttk.Label(branding, text="", style="Muted.TLabel")
        self.state_label.pack(anchor="w", pady=(0, 4))
        ttk.Button(header, text="♡  Support", command=self.show_support).pack(side=RIGHT, padx=(8, 0))

        self.nav = ttk.Frame(outer, style="Root.TFrame")
        self.nav.pack(fill="x", pady=(16, 0))

        # Build tabs (but do not pack immediately, will be packed dynamically in reload())
        self.tabs = ttk.Notebook(outer, style="Main.TNotebook", takefocus=False)
        self._build_access_tab()
        self._build_account_tab()
        self._build_permissions_tab()
        self._build_alert_tab()
        self._build_runtime_tab()

        # Device linking is optional and never hides the local settings.
        self.welcome_status_label = self._hint(self.account_body, "")
        self._nav_columns = 0
        self.nav_buttons = []
        for index, name in enumerate(("Connect", "My PC", "Permissions", "Alerts", "Preferences")):
            button = ttk.Button(self.nav, text=name, style="Nav.TButton",
                                command=lambda i=index: self.tabs.select(i))
            self.nav_buttons.append(button)
        self.nav.bind("<Configure>", self._layout_nav)
        self.tabs.bind("<<NotebookTabChanged>>", self._sync_nav)
        self._sync_nav()
        self.root.bind("<MouseWheel>", self._scroll_settings, add="+")
        self.root.bind("<FocusIn>", self._reveal_focus, add="+")

        # Build footer (but do not pack immediately)
        self.footer = ttk.Frame(outer, style="Root.TFrame")
        self.status_label = ttk.Label(self.footer, text="Changes save automatically", style="Status.TLabel")
        self.status_label.pack(fill="x", pady=(0, 8))
        self.footer.bind("<Configure>", lambda e: self.status_label.configure(wraplength=max(160, e.width - 8)))
        actions = ttk.Frame(self.footer, style="Root.TFrame")
        actions.pack(fill="x")
        ttk.Label(actions, text="Free. No limits to unlock.", style="Muted.TLabel").pack(side=LEFT)
        # Settings autosave automatically — no Save button. Keep only Start/Stop here.
        ttk.Button(actions, text="Stop", command=self.stop_zadoo).pack(side=RIGHT, padx=(6, 0))
        ttk.Button(actions, text="Start Zadoo", command=self.start_zadoo, style="Primary.TButton").pack(side=RIGHT, padx=(6, 0))

    def _layout_nav(self, event) -> None:
        columns = 5 if event.width >= 670 * self._ui_scale else 3
        if columns == self._nav_columns:
            return
        self._nav_columns = columns
        for column in range(5):
            self.nav.columnconfigure(column, weight=1 if column < columns else 0, uniform="nav" if column < columns else "")
        for index, button in enumerate(self.nav_buttons):
            button.grid(row=index // columns, column=index % columns, sticky="ew", padx=2, pady=2)

    def _sync_nav(self, _event=None) -> None:
        selected = self.tabs.index(self.tabs.select())
        for index, button in enumerate(getattr(self, "nav_buttons", [])):
            button.configure(style="SelectedNav.TButton" if index == selected else "Nav.TButton")
        self._stop_setup_animation()
        if selected == 0:
            self._motion_tick = 0
            self._motion_after = self.root.after(40, self._animate_setup)

    def _draw_setup_story(self, event) -> None:
        """Small vector illustration: no image decoding, GPU effects or background loop."""
        canvas = self.setup_story
        canvas.delete("all")
        width = event.width
        scale = self._ui_scale
        y = 18 * scale
        size = min(64 * scale, width / 4)
        centers = (width * 0.22, width * 0.78)
        for center, label in zip(centers, ("THIS PC", "YOUR BROWSER"), strict=True):
            x = center - size / 2
            canvas.create_rectangle(x, y, x + size, y + size * .62, fill=SURFACE_ALT, outline=ACCENT, width=2)
            canvas.create_line(center, y + size * .62, center, y + size * .76, fill=ACCENT, width=2)
            canvas.create_line(center - size * .2, y + size * .76, center + size * .2, y + size * .76, fill=ACCENT, width=2)
            for eye in (-.13, .13):
                ex = center + size * eye
                canvas.create_oval(ex - 2, y + size * .23, ex + 2, y + size * .23 + 4, fill=TEXT, outline="")
            canvas.create_arc(center - size * .13, y + size * .26, center + size * .13, y + size * .47,
                              start=180, extent=180, style="arc", outline=TEXT, width=2)
            canvas.create_text(center, y + size * .76 + 18 * scale, text=label, fill=MUTED,
                               font=("Segoe UI", 8, "bold"))
        self._story_track = (centers[0] + size * .7, centers[1] - size * .7, y + size * .32)
        start, end, middle = self._story_track
        canvas.create_line(start, middle, end, middle, fill=BORDER, dash=(3, 5), width=2)
        canvas.create_oval(start - 4, middle - 4, start + 4, middle + 4, fill=ACCENT, outline="", tags="packet")

    def _stop_setup_animation(self) -> None:
        if self._motion_after is not None:
            self.root.after_cancel(self._motion_after)
            self._motion_after = None

    def _animate_setup(self) -> None:
        self._motion_after = None
        if (self._hidden or not self.motion_var.get() or not self.root.winfo_viewable()
                or self.tabs.index(self.tabs.select()) != 0 or not hasattr(self, "_story_track")):
            return
        self._motion_tick += 1
        start, end, y = self._story_track
        progress = min(1.0, self._motion_tick / 24)
        eased = 1 - (1 - progress) ** 3
        x = start + (end - start) * eased
        self.setup_story.coords("packet", x - 4, y - 4, x + 4, y + 4)
        if self._motion_tick < 24:
            self._motion_after = self.root.after(40, self._animate_setup)

    def _fit_geometry(self, base_w: int, base_h: int) -> None:
        """Scale the requested size to the monitor DPI, then clamp to the screen and centre so
        the window fits on the current screen. Long pages scroll independently."""
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        if min(sw, sh) <= 0:
            raise RuntimeError(f"Tk reported an invalid screen size: {sw}x{sh}")
        w = max(1, min(int(base_w * self._ui_scale), sw - 40))
        h = max(1, min(int(base_h * self._ui_scale), sh - 96))
        x = max(0, (sw - w) // 2)
        y = max(0, (sh - h) // 3)
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def _tab_body(self, parent):
        """Width-tracking page; scrolling keeps all controls accessible on small displays."""
        canvas = tk.Canvas(parent, bg=SURFACE, highlightthickness=0, bd=0, takefocus=False)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollbar.pack(side=RIGHT, fill="y")
        canvas.pack(side=LEFT, fill=BOTH, expand=True)
        canvas.configure(yscrollcommand=scrollbar.set)
        inner = ttk.Frame(canvas, padding=20, style="Surface.TFrame")
        item = canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.bind("<Configure>", lambda event: canvas.itemconfigure(item, width=event.width))
        inner.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))
        self._scroll_canvases.append(canvas)
        return inner

    def _scroll_settings(self, event):
        widget = event.widget
        while widget is not None:
            if widget in self._scroll_canvases:
                if widget.yview() != (0.0, 1.0):
                    widget.yview_scroll(-1 if event.delta > 0 else 1, "units")
                return "break"
            widget = getattr(widget, "master", None)
        return None

    def _reveal_focus(self, event):
        widget = event.widget
        parent = getattr(widget, "master", None)
        while parent is not None:
            if parent in self._scroll_canvases:
                bounds = parent.bbox("all")
                if bounds and bounds[3] > parent.winfo_height():
                    top = widget.winfo_rooty() - parent.winfo_rooty()
                    bottom = top + widget.winfo_height()
                    if top < 0 or bottom > parent.winfo_height():
                        offset = top - 12 if top < 0 else bottom - parent.winfo_height() + 12
                        parent.yview_moveto((parent.canvasy(0) + offset) / bounds[3])
                break
            parent = getattr(parent, "master", None)

    def _hint(self, parent, text: str):
        label = ttk.Label(parent, text=text, style="SurfaceMuted.TLabel", justify=LEFT, wraplength=280)
        label.pack(fill="x", pady=(4, 12))
        label.bind("<Configure>", lambda event: label.configure(wraplength=max(120, event.width)))
        return label

    def _page_heading(self, parent, eyebrow: str, title: str, detail: str) -> None:
        ttk.Label(parent, text=eyebrow.upper(), style="Eyebrow.TLabel").pack(anchor="w", pady=(0, 6))
        heading = ttk.Label(parent, text=title, style="Section.TLabel", wraplength=300)
        heading.pack(fill="x")
        heading.bind("<Configure>", lambda event: heading.configure(wraplength=max(120, event.width)))
        self._hint(parent, detail)

    def show_support(self) -> None:
        """Display the supplied QR locally, only when explicitly opened."""
        if self._support_window and self._support_window.winfo_exists():
            self._support_window.lift()
            return
        from PIL import Image, ImageTk
        qr_path = resource_path("zadoo_vnc/static/ui/developer-qr.png")
        dialog = tk.Toplevel(self.root)
        self._support_window = dialog
        dialog.title("Support the developer · Zadoo")
        dialog.configure(bg=SURFACE)
        dialog.transient(self.root)
        dialog.resizable(False, False)
        dialog.iconbitmap(_find_icon())
        body = ttk.Frame(dialog, padding=20, style="Surface.TFrame")
        body.pack(fill=BOTH, expand=True)
        ttk.Label(body, text="A little support. A lot of thanks.", style="Surface.TLabel",
                  font=("Segoe UI", 13, "bold")).pack(anchor="w")
        self._hint(body, "Completely optional. Every Zadoo feature stays free, whether you contribute or not.")
        # Resize for display only; the packaged/downloadable source image stays untouched.
        with Image.open(qr_path) as original:
            limit = max(120, min(int(300 * self._ui_scale), self.root.winfo_screenheight() - int(320 * self._ui_scale)))
            original.thumbnail((int(250 * self._ui_scale), limit), Image.Resampling.LANCZOS)
            dialog.qr_image = ImageTk.PhotoImage(original)
        ttk.Label(body, image=dialog.qr_image, style="Surface.TLabel").pack(pady=4)
        ttk.Label(body, text=SUPPORT_UPI, style="Surface.TLabel").pack(pady=(10, 0))
        self._hint(body, "Scan in your UPI app. Confirm Divyansh Gupta as the recipient before sending.")
        copied = ttk.Label(body, text="", style="SurfaceMuted.TLabel")
        copied.pack(fill="x")
        def copy_upi():
            self.root.clipboard_clear()
            self.root.clipboard_append(SUPPORT_UPI)
            copied.configure(text="UPI ID copied")
        actions = ttk.Frame(body, style="Surface.TFrame")
        actions.pack(fill="x", pady=(10, 0))
        ttk.Button(actions, text="Copy UPI ID", command=copy_upi).pack(side=LEFT)
        close = ttk.Button(actions, text="Done", style="Primary.TButton", command=dialog.destroy)
        close.pack(side=RIGHT)
        dialog.bind("<Escape>", lambda _event: dialog.destroy())
        dialog.update_idletasks()
        x = max(0, self.root.winfo_rootx() + (self.root.winfo_width() - dialog.winfo_width()) // 2)
        y = max(0, min(self.root.winfo_rooty() + 30, self.root.winfo_screenheight() - dialog.winfo_height() - 60))
        dialog.geometry(f"+{x}+{y}")
        close.focus_set()

    def _access_code_validator(self, value: str) -> bool:
        return len(value or "") <= ACCESS_CODE_MAX_LENGTH

    def _build_access_tab(self) -> None:
        self.tab_access = ttk.Frame(self.tabs, style="Surface.TFrame")
        self.tabs.add(self.tab_access, text="Access")
        body = self._tab_body(self.tab_access)
        validator = (self.root.register(self._access_code_validator), "%P")
        self._setup_hero = ttk.Frame(body, style="Surface.TFrame")
        self._setup_hero.pack(fill="x", pady=(0, 12))
        self._setup_intro = ttk.Frame(self._setup_hero, style="Surface.TFrame")
        self._page_heading(self._setup_intro, "Configure / Quick start", "Connect in three easy steps.",
                           "Set a code. Choose what viewers can do. Start sharing.")
        self.setup_story = tk.Canvas(self._setup_hero, width=210, height=int(104 * self._ui_scale), bg=SURFACE,
                                     highlightthickness=0, takefocus=False)
        self.setup_story.bind("<Configure>", self._draw_setup_story)

        self._setup_grid = ttk.Frame(body, style="Surface.TFrame")
        self._setup_grid.pack(fill="x")
        form = ttk.LabelFrame(self._setup_grid, text="1  ·  Choose your access code", style="Card.TLabelframe")
        ttk.Label(form, text=f"Access code · up to {ACCESS_CODE_MAX_LENGTH} characters", style="Surface.TLabel").pack(anchor="w")
        code_row = ttk.Frame(form, style="Surface.TFrame")
        code_row.pack(fill="x", pady=(6, 0))
        self.access_code = ttk.Entry(code_row, validate="key", validatecommand=validator, show="•",
                                     font=("Consolas", 14))
        self.access_code.pack(side=LEFT, fill="x", expand=True)
        self.show_code_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(code_row, text="Show", variable=self.show_code_var,
                        command=lambda: self.access_code.configure(show="" if self.show_code_var.get() else "•")).pack(side=RIGHT, padx=(8, 0))
        self._hint(form, "This is the code viewers enter on the connection screen. Share it only with people you trust.")

        permissions = ttk.LabelFrame(self._setup_grid, text="2  ·  Choose viewer permissions", style="Card.TLabelframe")
        self.permission_summary = self._hint(permissions, "Review access before you share this PC.")
        ttk.Button(permissions, text="Review permissions  →", command=lambda: self.tabs.select(2)).pack(anchor="w")

        # ── Remote access (public link + start) ───────────────────────
        self.public_url_frame = ttk.LabelFrame(self._setup_grid, text="3  ·  Start and share your link", style="Card.TLabelframe")
        url_row = ttk.Frame(self.public_url_frame, style="Surface.TFrame")
        url_row.pack(fill="x", pady=(2, 6))
        self.public_url_label = ttk.Label(url_row, text="Zadoo not running", style="SurfaceMuted.TLabel",
                                          font=("Segoe UI", 9), cursor="hand2")
        self.public_url_label.pack(fill="x", pady=(6, 10))
        self.public_url_label.bind("<Configure>", lambda event: self.public_url_label.configure(wraplength=max(100, event.width)))
        self.public_url_label.bind("<Button-1>", lambda _e: self._open_public_url())
        buttons = ttk.Frame(url_row, style="Surface.TFrame")
        buttons.pack(fill="x")
        self.copy_link_button = ttk.Button(buttons, text="Copy link", command=self.copy_public_url, state="disabled")
        self.copy_link_button.pack(side=LEFT)
        ttk.Button(buttons, text="New link", command=self._refresh_public_url).pack(side=LEFT, padx=(6, 0))
        self._hint(self.public_url_frame, "Select Start Zadoo below. When the link appears, send it with your code. New link replaces the current one; Stop ends sharing.")
        self._setup_cards = (form, permissions, self.public_url_frame)
        self._setup_columns = 0
        body.bind("<Configure>", self._layout_setup, add="+")

    def _layout_setup(self, event) -> None:
        columns = 2 if event.width >= 720 * self._ui_scale else 1
        if columns == self._setup_columns:
            return
        self._setup_columns = columns
        wide = columns == 2
        self._setup_hero.columnconfigure(0, weight=3 if wide else 1)
        self._setup_hero.columnconfigure(1, weight=2 if wide else 0)
        self._setup_intro.grid(row=0, column=0, sticky="ew")
        self.setup_story.grid(row=0 if wide else 1, column=1 if wide else 0, sticky="ew")
        for column in range(2):
            self._setup_grid.columnconfigure(column, weight=1 if column < columns else 0,
                                              uniform="setup" if column < columns else "")
        for index, card in enumerate(self._setup_cards[:2]):
            card.grid(row=0 if wide else index, column=index if wide else 0, sticky="nsew",
                      padx=(0, 12 if wide and index == 0 else 0), pady=(0, 12))
        self.public_url_frame.grid(row=1 if wide else 2, column=0, columnspan=columns, sticky="ew")

    def copy_public_url(self) -> None:
        if not self._runtime_running or not self._runtime_url:
            self._set_status("No connection link yet. Select Start Zadoo and wait for the link.")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(self._runtime_url)
        self._set_status("Connection link copied. Share it and your access code only with someone you trust.")

    def _update_permission_summary(self, *_args) -> None:
        count = sum(bool(var.get()) for var in self.permission_vars.values())
        text = (f"{count} of {len(PERMISSION_KEYS)} controls allowed. You choose the access on this PC."
                if count else "View only. Viewers can see the screen; all remote controls are off.")
        self.permission_summary.configure(text=text)

    def _build_account_tab(self) -> None:
        self.tab_account = ttk.Frame(self.tabs, style="Surface.TFrame")
        self.tabs.add(self.tab_account, text="Device")
        body = self._tab_body(self.tab_account)
        self.account_body = body
        self._page_heading(body, "02 / Device", "Make yourself at home.",
                           "Name this PC so it is easy to recognize. Dashboard sign-in is optional; remote access works without it.")

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
        self.profile_email_label = ttk.Label(info_frame, text="", style="SurfaceMuted.TLabel",
                                              font=("Segoe UI", 9))
        self.profile_email_label.pack(anchor="w")
        self.account_state = ttk.Label(info_frame, text="", style="SurfaceMuted.TLabel",
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
        ttk.Label(self.signin_frame, text="Dashboard · optional", style="Eyebrow.TLabel").pack(anchor="w", pady=(5, 10))
        self._hint(self.signin_frame, "Select Sign in, complete the browser steps, then return here. Use Check sign-in if the status has not updated.")

        code_row = ttk.Frame(self.signin_frame, style="Surface.TFrame")
        code_row.pack(fill="x")
        ttk.Label(code_row, text="Browser sign-in code", style="Surface.TLabel").pack(anchor="w")
        
        btn_row = ttk.Frame(code_row, style="Surface.TFrame")
        btn_row.pack(fill="x", pady=(2, 10))
        self.activation_code = ttk.Entry(btn_row, state="readonly")
        self.activation_code.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        ttk.Button(btn_row, text="Copy code", command=self.copy_activation_code).grid(row=0, column=2, sticky="ew", padx=(6, 0), pady=(0, 8))
        ttk.Button(btn_row, text="Sign in", command=self.start_activation, style="Primary.TButton").grid(row=1, column=0, sticky="ew")
        ttk.Button(btn_row, text="Check sign-in", command=self.poll_activation).grid(row=1, column=1, columnspan=2, sticky="ew", padx=(6, 0))
        for column in range(3):
            btn_row.columnconfigure(column, weight=1)

    def _build_runtime_tab(self) -> None:
        self.tab_runtime = ttk.Frame(self.tabs, style="Surface.TFrame")
        self.tabs.add(self.tab_runtime, text="Preferences")
        body = self._tab_body(self.tab_runtime)
        self._page_heading(body, "05 / Preferences", "Fits your routine.",
                           "Choose how Zadoo behaves on this PC. Changes save automatically.")
        card = ttk.LabelFrame(body, text="Windows behavior", style="Card.TLabelframe")
        card.pack(fill="x")
        self._check_button(card, "Start when I sign in to Windows", self.autostart_var).pack(anchor="w", pady=4)
        self._hint(card, "Keeps this PC available after you sign in. Windows may require administrator access to update the startup task.")
        self._check_button(card, "Keep minimized Settings on the taskbar", self.show_taskbar_var).pack(anchor="w", pady=4)
        self._hint(card, "Off: minimizing closes the Settings window, while sharing continues. Reopen Settings from the Zadoo shortcut.")
        ttk.Checkbutton(card, text="Animate the setup illustration", variable=self.motion_var,
                        command=self._stop_setup_animation).pack(anchor="w", pady=4)
        self._hint(card, "Follows Windows on launch. A short animation plays only when you open Connect; nothing loops while sharing.")
        ttk.Button(card, text="Reapply startup preference", command=self.apply_startup).pack(anchor="w", pady=(4, 0))
        support = ttk.LabelFrame(body, text="Local support", style="Card.TLabelframe")
        support.pack(fill="x", pady=(14, 0))
        self._hint(support, "Save recent logs and system details to a file when troubleshooting. Known access codes, tokens, IDs, email and public links are redacted. Review the file before sharing it.")
        ttk.Button(support, text="Export Diagnostics", command=self.export_diagnostics).pack(anchor="w", pady=(10, 0))
        self.runtime_state = self._hint(body, "")

    def _check_button(self, parent, text, var):
        return ttk.Checkbutton(parent, text=text, variable=var)

    def _build_permissions_tab(self) -> None:
        self.tab_permissions = ttk.Frame(self.tabs, style="Surface.TFrame")
        self.tabs.add(self.tab_permissions, text="Permissions")
        body = self._tab_body(self.tab_permissions)
        self._page_heading(body, "03 / Permissions", "You decide what's allowed.",
                           "These switches apply to viewers using your access code. Enable only what they need; your screen can still be viewed when controls are off.")
        perms = ttk.Frame(body, style="Surface.TFrame")
        perms.pack(fill="x")
        cards = []
        for key in PERMISSION_KEYS:
            var = tk.BooleanVar(value=False)
            self.permission_vars[key] = var
            var.trace_add("write", self._update_permission_summary)
            card = ttk.LabelFrame(perms, style="Card.TLabelframe", padding=10)
            self._check_button(card, PERMISSION_LABELS.get(key, key), var).pack(anchor="w")
            self._hint(card, PERMISSION_HELP[key])
            cards.append(card)
        state = {"columns": 0}
        def layout_permissions(event):
            columns = 2 if event.width >= 580 * self._ui_scale else 1
            if state["columns"] == columns:
                return
            state["columns"] = columns
            perms.columnconfigure(0, weight=1, uniform="permissions")
            perms.columnconfigure(1, weight=1 if columns == 2 else 0, uniform="permissions" if columns == 2 else "")
            for index, card in enumerate(cards):
                card.grid(row=index // columns, column=index % columns, sticky="nsew", padx=(0, 8 if columns == 2 and index % 2 == 0 else 0), pady=(0, 8))
        perms.bind("<Configure>", layout_permissions)

        actions = ttk.Frame(body, style="Surface.TFrame")
        actions.pack(fill="x", pady=(10, 0))
        ttk.Button(actions, text="Allow all controls", command=self.allow_all_permissions).pack(side=LEFT)
        ttk.Button(actions, text="View only", command=self.clear_permissions).pack(side=LEFT, padx=(6, 0))

    def _build_alert_tab(self) -> None:
        self.tab_alerts = ttk.Frame(self.tabs, style="Surface.TFrame")
        self.tabs.add(self.tab_alerts, text="Alerts")
        body = self._tab_body(self.tab_alerts)
        self._page_heading(body, "04 / Alerts", "A message, one tap away.",
                           "Prepare up to four messages for connected viewers. Enable a preset and fill in its title or message, then send it from Host Controls.")
        email_card = ttk.LabelFrame(body, text="Email notifications · optional", style="Card.TLabelframe")
        email_card.pack(fill="x", pady=(0, 14))
        ttk.Label(email_card, text="Send notifications to", style="Surface.TLabel").pack(anchor="w")
        self.email_to = ttk.Entry(email_card)
        self.email_to.pack(fill="x", pady=(6, 0))
        self.email_state = self._hint(email_card, "")
        alert_tabs = ttk.Notebook(body)
        alert_tabs.pack(fill=BOTH, expand=True)
        for code in ("A", "B", "C", "D"):
            frame = ttk.Frame(alert_tabs, padding=16, style="Surface.TFrame")
            alert_tabs.add(frame, text=f"Alert {code}")
            enabled = tk.BooleanVar(value=False)
            title = StringVar(value="")
            message = StringVar(value="")
            self.alert_vars[code] = {"enabled": enabled, "title": title, "message": message}
            self._check_button(frame, "Enable this preset", enabled).pack(anchor="w", pady=(0, 10))
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
            self._load_account_values()
            self._load_access_values()
            self._load_permission_values()
            self._load_alert_values()
            self._load_runtime_values()
            self.state_label.configure(text="Free · Configured" if self._setup_complete() else "Free · Set an access code")
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
        self.email_state.configure(text=(f"Notifications are not configured. Missing: {', '.join(missing)}. Remote access still works."
                                         if missing else "Email notifications are configured."))

    def _dashboard_linked(self) -> bool:
        try:
            return bool(self.store.get_device_token())
        except Exception as exc:
            logging.warning("Optional dashboard credentials unavailable: %s", exc)
            return False

    def _load_account_values(self) -> None:
        is_signed_in = self._dashboard_linked()

        self.footer.pack(side="bottom", fill="x")
        self.tabs.pack(fill=BOTH, expand=True, pady=(10, 8))
        self.profile_card.pack_forget()
        self.device_name_frame.pack_forget()
        self.signin_frame.pack_forget()
        self.device_name_frame.pack(fill="x", pady=(0, 10))
        if is_signed_in:
            self.profile_card.pack(fill="x", pady=(0, 10))
        else:
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
            # Do not call Tk from the worker: the settings main loop may already
            # have stopped by the time this network request completes.
            self._runtime_status_result = (running, url, error)
        threading.Thread(target=_check, daemon=True).start()
        self.root.after(50, self._finish_runtime_status)

    def _finish_runtime_status(self) -> None:
        result = self._runtime_status_result
        if result is None:
            self.root.after(50, self._finish_runtime_status)
            return
        self._runtime_status_result = None
        self._apply_runtime_status(*result)

    def _apply_runtime_status(self, running: bool, url: str, error: str) -> None:
        self._runtime_running = running
        self._runtime_url = url
        with suppress(tk.TclError):
            self.copy_link_button.configure(state="normal" if running and url and not error else "disabled")
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
                    # Runtime is up but the tunnel is disabled — stop hanging.
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
                result = _local_rpc(
                    "runtime.refresh_tunnel",
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

    def _fetch_latest_account_details_async(self, show_status: bool = False) -> None:
        if show_status:
            self._set_status("Refreshing account info...")
        def _do_fetch():
            errors = self._cloud_fetch_errors(
                (("profile", self.cloud.fetch_profile),)
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
                for label, fetch in (
                    ("profile", self.cloud.fetch_profile),
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
        """Unlink the optional dashboard account without interrupting remote access."""
        if self._signout_active:
            return
        self._signout_active = True
        self._set_status("Unlinking dashboard account...")

        def _unlink():
            error = ""
            try:
                result = self.cloud.go_offline()
                if not result["success"]:
                    logging.warning("Optional offline sync failed: %s", result["error"])
            except Exception as exc:
                logging.warning("Optional offline sync failed: %s", exc)
            try:
                self.store.clear_account()
            except Exception as exc:
                error = str(exc)

            def _done():
                self._signout_active = False
                self.reload()
                self._set_status(error or "Dashboard account unlinked. Local access is still available.")

            self.root.after(0, _done)

        threading.Thread(target=_unlink, daemon=True).start()

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

    def export_diagnostics(self) -> None:
        default_name = f"Zadoo-diagnostics-{time.strftime('%Y%m%d-%H%M%S')}.zip"
        initial_dir = Path.home() / "Downloads"
        target = filedialog.asksaveasfilename(
            title="Export redacted Zadoo diagnostics",
            initialdir=str(initial_dir if initial_dir.is_dir() else Path.home()),
            initialfile=default_name,
            defaultextension=".zip",
            filetypes=(("ZIP archive", "*.zip"),),
        )
        if not target:
            return
        self._set_status("Exporting redacted diagnostics...")

        def _export():
            try:
                from .diagnostics import build_diagnostic_bundle

                runtime_status = None
                if _local_server_running():
                    runtime_status = _get_local_json("/api/runtime/status", {}, timeout=1.2)
                path = build_diagnostic_bundle(
                    Path(target), store=self.store, runtime_status=runtime_status
                )
                message = f"Diagnostics exported: {path}"
            except Exception as exc:
                message = f"Diagnostics export failed: {exc}"
            with suppress(tk.TclError):
                self.root.after(0, lambda: self._set_status(message))

        threading.Thread(target=_export, daemon=True).start()

    def stop_zadoo(self) -> None:
        if not _local_server_running():
            self._set_status("Zadoo runtime is not running")
            return
        try:
            payload = {"admin_code": self.saved_access_code or self.access_code.get().strip()}
            result = _local_rpc("runtime.stop", payload)
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
        self._set_status("Starting Zadoo...")

        def _check_and_stop():
            error = ""
            restarted = False
            try:
                if _local_server_running():
                    stopped = _local_rpc("runtime.stop", {"admin_code": admin_code})
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
        if event.widget is self.root:
            self._stop_setup_animation()
        if event.widget is self.root and self.root.state() == "iconic" and not bool(self.show_taskbar_var.get()):
            self.root.after(80, self.hide)

    def hide(self) -> None:
        if self._hidden:
            return
        self._hidden = True
        self._stop_setup_animation()
        with suppress(tk.TclError):
            self.root.withdraw()
        self.root.after(50, self.root.quit)

    def run(self) -> None:
        self.root.mainloop()
        with suppress(tk.TclError):
            self.root.destroy()


def run_settings_window() -> None:
    ZadooSettingsWindow().run()
