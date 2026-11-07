"""Integration checks for the account-optional Windows settings and status monitor."""
from __future__ import annotations

import os
import sys
import tempfile
import threading
import time
from pathlib import Path
from types import SimpleNamespace
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

        # A delayed status response must never call Tk after its loop has stopped.
        main_thread = threading.get_ident()
        callbacks = []
        def schedule(_delay, callback):
            if threading.get_ident() != main_thread:
                raise AssertionError("Runtime status worker called Tk outside the main thread")
            callbacks.append(callback)

        delayed = ZadooSettingsWindow.__new__(ZadooSettingsWindow)
        delayed.root = SimpleNamespace(after=schedule)
        delayed._runtime_status_result = None
        release = threading.Event()
        def delayed_status(*_args, **_kwargs):
            if not release.wait(2):
                raise TimeoutError("Status-poll test did not release its worker")
            return {"success":True,"running":False,"public_url":None}

        with patch("zadoo_vnc.settings_window._get_local_json", delayed_status):
            delayed._poll_runtime_status()
            release.set()
            deadline = time.monotonic() + 2
            while delayed._runtime_status_result is None and time.monotonic() < deadline:
                time.sleep(.01)
        if delayed._runtime_status_result != (False, "", "") or len(callbacks) != 1:
            raise AssertionError("Delayed runtime status was not handed off without a worker Tk call")
        applied = []
        delayed._apply_runtime_status = lambda *value: applied.append(value)
        callbacks[0]()
        if applied != [(False, "", "")]:
            raise AssertionError("Main-thread status completion did not apply its result")
        print("SETTINGS_STATUS_SHUTDOWN: PASS delayed worker never calls Tk")

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

        from zadoo_vnc.tunnel import CloudflareTunnelManager

        binary = Path(folder) / "test-cloudflared.exe"
        binary.touch()
        manager = CloudflareTunnelManager.__new__(CloudflareTunnelManager)
        manager.cloudflared_path = str(binary)
        manager._signature_cache = None
        valid = SimpleNamespace(returncode=0, stdout="Valid", stderr="")
        with patch("zadoo_vnc.tunnel.subprocess.run", return_value=valid) as signature:
            manager._verify_cloudflared_signature()
            manager._verify_cloudflared_signature()
            if signature.call_count != 1:
                raise AssertionError("Signature verification was not cached within the session")
            command = signature.call_args.args[0]
            if "-NonInteractive" not in command or "Join-Path $PSHOME" not in command[-1]:
                raise AssertionError("Signature check may autoload an incompatible PowerShell module")
        changed = binary.stat().st_mtime_ns + 1_000_000_000
        os.utime(binary, ns=(changed, changed))
        unsigned = SimpleNamespace(returncode=0, stdout="NotSigned", stderr="")
        with patch("zadoo_vnc.tunnel.subprocess.run", return_value=unsigned):
            try:
                manager._verify_cloudflared_signature()
            except RuntimeError as exc:
                if str(exc) != "cloudflared signature check failed: NotSigned":
                    raise
            else:
                raise AssertionError("A changed or unsigned tunnel binary was accepted")
        print("TUNNEL_SIGNATURE: PASS native module, cache invalidation and unsigned rejection")


if __name__ == "__main__":
    main()
