"""Isolated Tk layout checks; never starts sharing or changes Windows preferences."""
from __future__ import annotations

import argparse
import json
import sys
import tempfile
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from zadoo_vnc.settings import PERMISSION_KEYS, SettingsStore  # noqa: E402
from zadoo_vnc.settings_window import ZadooSettingsWindow  # noqa: E402


def descendants(widget):
    for child in widget.winfo_children():
        yield child
        yield from descendants(child)


def check(condition, message):
    if not condition:
        raise AssertionError(message)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--preview", action="store_true")
    args = parser.parse_args()
    with tempfile.TemporaryDirectory(prefix="zadoo-ui-test-") as folder, \
            patch("zadoo_vnc.settings_window.startup_task_exists", return_value=False), \
            patch.object(ZadooSettingsWindow, "_poll_runtime_status"), \
            patch.object(ZadooSettingsWindow, "_dashboard_linked", return_value=False), \
            patch.object(ZadooSettingsWindow, "_wire_autosave"), \
            patch.object(ZadooSettingsWindow, "start_zadoo"), \
            patch.object(ZadooSettingsWindow, "stop_zadoo"), \
            patch.object(ZadooSettingsWindow, "apply_startup"):
        store = SettingsStore(Path(folder) / "config.json")
        store.apply_setup({"access_code": "UI-PREVIEW", "autostart_enabled": False})
        app = ZadooSettingsWindow(store)
        if args.preview:
            app.root.title("Zadoo · UI preview (isolated test)")
            app.run()
            return
        errors = []
        app.root.report_callback_exception = lambda _kind, value, _trace: errors.append(str(value))
        try:
            for width, height in ((480, 360), (640, 480), (880, 700)):
                app.root.geometry(f"{int(width * app._ui_scale)}x{int(height * app._ui_scale)}")
                for tab in range(5):
                    app.tabs.select(tab)
                    for _ in range(5):
                        app.root.update()
                    canvas = app._scroll_canvases[tab]
                    left = canvas.winfo_rootx()
                    right = left + canvas.winfo_width()
                    for widget in descendants(canvas):
                        if not widget.winfo_ismapped() or widget.winfo_class() not in ("TButton", "TEntry", "TCheckbutton", "TLabel"):
                            continue
                        x = widget.winfo_rootx()
                        if x < left - 2 or x + widget.winfo_width() > right + 2:
                            errors.append(f"{width}x{height} tab {tab}: {widget.winfo_class()} exceeds page width: {widget}")
                    if app.status_label.winfo_height() == 1:
                        errors.append(f"Footer clipped at {width}x{height}")
                expected_columns = 3 if width < 700 else 5
                if app._nav_columns != expected_columns:
                    errors.append(f"Navigation did not reflow at {width}px")
            payload = app._collect_payload()
            check(set(payload["permissions"]) == set(PERMISSION_KEYS), "Permission mapping changed")
            check(payload["access_code"] == "UI-PREVIEW", "Access code changed")
            check(app.access_code.cget("show") == "•", "Access code is not masked")
            app.clear_permissions()
            check(not any(app._collect_payload()["permissions"].values()), "View-only did not clear controls")
            app.allow_all_permissions()
            check(all(app._collect_payload()["permissions"].values()), "Allow-all missed controls")
            app.show_support()
            app.root.update()
            check(app._support_window.winfo_exists(), "Support window did not open")
            app._support_window.destroy()
            check(not errors, "\n".join(errors))
        finally:
            app.root.destroy()
        print(json.dumps({"native_layouts": ["480x360", "640x480", "880x700"], "pages": 5,
                          "permission_mapping": "pass", "support_dialog": "pass", "callback_errors": errors}, indent=2))


if __name__ == "__main__":
    main()
