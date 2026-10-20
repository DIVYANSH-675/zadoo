"""Integration checks for the account-optional Windows settings and status monitor."""
from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


def main():
    with tempfile.TemporaryDirectory(prefix="zadoo-free-settings-") as folder:
        os.environ["ZADOO_SETTINGS_DIR"] = folder
        os.environ["ZADOO_DISABLE_TUNNEL"] = "1"
        from zadoo_vnc.settings import SettingsStore
        from zadoo_vnc.settings_window import ZadooSettingsWindow

        store = SettingsStore(Path(folder) / "config.json")
        store.apply_setup({"access_code": "FREE-TEST", "autostart_enabled": False})
        window = ZadooSettingsWindow(store)
        launched = []
        try:
            if window.store.get_device_token():
                raise AssertionError("Test must run without an account")

            def launch():
                launched.append(True)
                window.root.after(0, window.root.quit)

            window._launch_runtime = launch
            with (
                patch("zadoo_vnc.settings_window._local_server_running", return_value=False),
                patch.object(window.cloud, "_request", side_effect=AssertionError("Start contacted cloud")),
            ):
                window.root.after(50, window.start_zadoo)
                window.root.after(3000, window.root.quit)
                window.root.mainloop()
            if not window.tabs.winfo_ismapped() or not window.footer.winfo_ismapped():
                raise AssertionError("Local controls are hidden without a dashboard account")
            if not launched:
                raise AssertionError("Free host did not start without an account")
            print("FREE_SETTINGS: PASS visible local controls and account-free start")
            with patch.object(store, "get_device_token", side_effect=RuntimeError("Simulated invalid dashboard credential")):
                window._load_account_values()
            print("OPTIONAL_DASHBOARD_CREDENTIAL: PASS local settings remain available")
            cleared = []

            def clear_account():
                cleared.append(True)

            def unlink_status(message):
                if "unlinked" in message:
                    window.root.after(0, window.root.quit)

            with (
                patch.object(window.cloud, "go_offline", side_effect=OSError("Simulated dashboard outage")),
                patch.object(store, "clear_account", clear_account),
                patch.object(window, "_set_status", unlink_status),
            ):
                window.root.after(50, window.sign_out)
                window.root.after(3000, window.root.quit)
                window.root.mainloop()
            if not cleared:
                raise AssertionError("An offline dashboard prevented local account unlinking")
            print("OFFLINE_UNLINK: PASS optional account can be unlinked during an outage")
        finally:
            window.root.destroy()

        from zadoo_vnc.saas import ZadooCloudClient
        from zadoo_vnc.server import VNCServer

        server = VNCServer(6173)
        server.enable_tunnel = False
        server.settings_store = store
        with patch.object(store, "get_device_token", return_value="test-only-device-token"):
            def unavailable(_client, _url):
                server._status_monitor_stop.set()
                raise OSError("Simulated dashboard outage")

            with patch.object(ZadooCloudClient, "heartbeat", unavailable):
                server._start_status_monitor()
                server._status_monitor_thread.join(timeout=3)
        if server._status_monitor_thread.is_alive() or server._fatal_error is not None:
            raise AssertionError("Optional dashboard outage blocked free remote access")
        print("OPTIONAL_DASHBOARD_OUTAGE: PASS host remains available")


if __name__ == "__main__":
    main()
