"""Application entrypoint and startup orchestration."""
from __future__ import annotations

import atexit
import asyncio
import ctypes
import logging
import os
import platform
import signal
import subprocess
import sys
import time
import urllib.request
import webbrowser
from pathlib import Path

from .config import _load_dotenv
from .logging_utils import _log_fallback, _setup_logging_to_file
from .network import get_local_ip
from .settings import get_settings_store


def configure_event_loop_policy():
    try:
        if sys.platform.startswith("win"):
            asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
    except Exception:
        pass


def configure_stdout_encoding():
    try:
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


def maybe_hide_console():
    try:
        if getattr(sys, "frozen", False) or os.environ.get("HIDE_CONSOLE") == "1":
            hwnd = ctypes.windll.kernel32.GetConsoleWindow()
            if hwnd:
                ctypes.windll.user32.ShowWindow(hwnd, 0)
    except Exception:
        pass


def configure_logging():
    _setup_logging_to_file()
    log_level_name = os.environ.get("ZADOO_LOG_LEVEL", "INFO").upper()
    log_level = getattr(logging, log_level_name, logging.INFO)
    try:
        logging.basicConfig(
            level=log_level,
            format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
            datefmt="%H:%M:%S",
        )
        logging.getLogger("asyncio").setLevel(logging.WARNING)
        logging.getLogger("websockets.server").setLevel(logging.WARNING)
        logging.getLogger("websockets").setLevel(logging.WARNING)
        try:
            logging.getLogger("comtypes").setLevel(logging.WARNING)
            logging.getLogger("comtypes.client").setLevel(logging.WARNING)
            logging.getLogger("comtypes.client._code_cache").setLevel(logging.WARNING)
            logging.getLogger("comtypes._post_coinit").setLevel(logging.WARNING)
        except Exception:
            pass
    except Exception:
        pass


def install_startup_task():
    try:
        from .windows_startup import set_startup_task

        ok, message = set_startup_task(True)
        if not ok:
            logging.warning("Startup task installation failed: %s", message)
            print(f"Startup task installation failed: {message}")
    except Exception as exc:
        logging.warning("Startup task installation failed", exc_info=True)
        try:
            print(f"Startup task installation failed: {exc}")
        except Exception:
            pass


def _local_url(path="/"):
    return f"http://127.0.0.1:6173{path}"


def _local_server_running():
    try:
        with urllib.request.urlopen(_local_url("/api/settings/status"), timeout=0.6) as response:
            return int(getattr(response, "status", 0) or 0) < 500
    except Exception:
        return False


def _open_local_page(path="/"):
    try:
        webbrowser.open(_local_url(path))
    except Exception:
        pass


def _zadoo_pids_on_port(port):
    if os.name != "nt":
        return []
    command = (
        f"$pids=(Get-NetTCPConnection -LocalPort {int(port)} -State Listen -ErrorAction SilentlyContinue | "
        "Select-Object -ExpandProperty OwningProcess -Unique); "
        "foreach($p in $pids){ "
        "$proc=Get-CimInstance Win32_Process -Filter \"ProcessId=$p\" -ErrorAction SilentlyContinue; "
        "if($proc){ [pscustomobject]@{ProcessId=$proc.ProcessId;CommandLine=$proc.CommandLine;ExecutablePath=$proc.ExecutablePath} } "
        "} | ConvertTo-Json -Compress"
    )
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", command],
            capture_output=True,
            text=True,
            timeout=5,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        if result.returncode != 0 or not (result.stdout or "").strip():
            return []
        import json

        data = json.loads(result.stdout)
        rows = data if isinstance(data, list) else [data]
        pids = []
        for row in rows:
            cmd = str(row.get("CommandLine") or row.get("ExecutablePath") or "").lower()
            pid = int(row.get("ProcessId") or 0)
            if pid > 0 and ("zadoo" in cmd or "zadoo_vnc" in cmd):
                pids.append(pid)
        return pids
    except Exception:
        return []


def _stop_existing_zadoo_on_port(port):
    stopped = False
    for pid in _zadoo_pids_on_port(port):
        if pid == os.getpid():
            continue
        try:
            subprocess.run(
                ["taskkill", "/PID", str(pid), "/T", "/F"],
                capture_output=True,
                timeout=5,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            stopped = True
        except Exception:
            pass
    return stopped


class ProcessProtector:
    def __init__(self):
        self.protected = True
        self.start_protection()

    def start_protection(self):
        atexit.register(self.cleanup)
        if sys.platform != "win32" and hasattr(signal, "SIGTERM"):
            signal.signal(signal.SIGTERM, self.signal_handler)
        signal.signal(signal.SIGINT, self.signal_handler)

    def restart_protection(self):
        try:
            if getattr(sys, "frozen", False):
                cmd = [sys.executable]
            else:
                script = Path(__file__).resolve().parent.parent / "zadoo_vnc_single.py"
                if script.exists():
                    cmd = [sys.executable, str(script)]
                else:
                    _log_fallback("process_protector.restart", "python_module_entrypoint", f"missing={script}")
                    cmd = [sys.executable, "-m", "zadoo_vnc.app"]
            subprocess.Popen(cmd, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        except Exception:
            pass

    def signal_handler(self, signum, frame):
        if signum == signal.SIGINT:
            self.cleanup()
            raise KeyboardInterrupt
        if self.protected:
            self.restart_protection()

    def cleanup(self):
        self.protected = False


def main():
    configure_event_loop_policy()
    configure_stdout_encoding()
    maybe_hide_console()
    _load_dotenv()
    configure_logging()
    args = {arg.lower() for arg in sys.argv[1:]}
    if "--settings" in args:
        from .settings_window import run_settings_window

        run_settings_window()
        return
    open_settings_store = get_settings_store()
    # Only redirect to settings window if the device has not been signed in.
    # Do NOT block on configured()/setup_complete — that field is only set when
    # an access code is saved, which is optional. The device token is the real
    # gate for whether the server can start.
    if "--open" in args and not open_settings_store.get_device_token():
        from .settings_window import run_settings_window

        run_settings_window()
        return

    if "--open" in args and _local_server_running():
        print("Existing Zadoo instance detected; restarting it...")
        _stop_existing_zadoo_on_port(6173)
        time.sleep(0.8)

    disable_startup_task = os.environ.get("ZADOO_DISABLE_STARTUP_TASK", "").strip().lower() in {
        "1",
        "true",
        "yes",
        "on",
    }
    autostart_enabled = True
    try:
        autostart_enabled = bool(get_settings_store().load(reload=True).get("autostart_enabled", True))
    except Exception:
        autostart_enabled = True
    if getattr(sys, "frozen", False) and not disable_startup_task and autostart_enabled:
        install_startup_task()

    use_process_protector = os.environ.get("ZADOO_DISABLE_PROCESS_PROTECTOR", "").strip().lower() not in {
        "1",
        "true",
        "yes",
        "on",
    }
    _protector = ProcessProtector() if use_process_protector else None

    from .dependencies import HAS_PIL, HAS_PYAUTOGUI
    from .screen_capture import ScreenCapturer
    from .server import VNCServer
    from .tunnel import CloudflareTunnelManager

    if not (HAS_PYAUTOGUI and HAS_PIL):
        print("Missing critical dependencies: pyautogui and Pillow are required")
        sys.exit(1)

    local_ip = get_local_ip()
    web_port = 6173
    if _local_server_running():
        print("Existing Zadoo instance detected; restarting it...")
        _stop_existing_zadoo_on_port(web_port)
    elif _zadoo_pids_on_port(web_port):
        _stop_existing_zadoo_on_port(web_port)

    settings_store = get_settings_store()
    setup_complete = settings_store.configured()
    signed_in = bool(settings_store.get_device_token())
    tunnel_disabled_by_env = os.environ.get("ZADOO_DISABLE_TUNNEL", "").strip().lower() in {"1", "true", "yes", "on"}
    cloud_ready = False
    revoked = False
    cloud_block_reason = ""
    # The public link / tunnel generates whenever the device is SIGNED IN, regardless of
    # remaining credits — billing is enforced inside the session (grace/lock + pay), not
    # at startup. Only a missing sign-in, an explicit revoke, or the env switch disable it.
    if not signed_in:
        cloud_block_reason = "Tunnel disabled until Zadoo is signed in"
    else:
        try:
            from .saas import ZadooCloudClient

            cloud_result = ZadooCloudClient(settings_store).entitlement()
            entitlement = (cloud_result.get("entitlement") if isinstance(cloud_result, dict) else None) or {}
            if not entitlement:
                entitlement = settings_store.load(reload=True).get("entitlement_cache") or {}
            if entitlement.get("revoked"):
                revoked = True
                cloud_block_reason = "Tunnel disabled because this device was revoked"
            elif entitlement.get("allowed"):
                cloud_ready = True
        except Exception:
            entitlement = settings_store.load(reload=True).get("entitlement_cache") or {}
            if isinstance(entitlement, dict) and entitlement.get("revoked"):
                revoked = True
                cloud_block_reason = "Tunnel disabled because this device was revoked"
    use_tunnel = signed_in and not revoked and not tunnel_disabled_by_env
    tunnel_manager = None
    if use_tunnel:
        tunnel_manager = CloudflareTunnelManager(primary_port=web_port)

    print(f"\n{'=' * 60}")
    print("COMPLETE VNC WITH TUNNEL")
    print(f"System: {platform.system()} {platform.release()}")
    print(f"Local: http://localhost:{web_port}")
    print(f"Network: http://{local_ip}:{web_port}")
    if use_tunnel and tunnel_manager:
        print(f"Tunnel port: {web_port}")
    elif not setup_complete:
        print("Tunnel disabled until first-launch setup is completed")
    elif tunnel_disabled_by_env:
        print("Tunnel disabled by ZADOO_DISABLE_TUNNEL")
    elif cloud_block_reason:
        print(cloud_block_reason)
    print("To stop: Press Ctrl+C")
    print("=" * 60)

    vnc_server = VNCServer(web_port)
    vnc_server.enable_tunnel = use_tunnel
    # Surface WHY the tunnel is off so the Settings window can show it instead of
    # hanging forever on "Starting Zadoo…".
    vnc_server.tunnel_block_reason = "" if use_tunnel else (
        cloud_block_reason or ("Tunnel disabled by ZADOO_DISABLE_TUNNEL" if tunnel_disabled_by_env else "Tunnel disabled")
    )
    if tunnel_manager:
        tunnel_manager.email_port = web_port
        print(f"Email will include port: {tunnel_manager.email_port}")
        vnc_server.set_tunnel_manager(tunnel_manager)
        print("Tunnel will launch after the local server binds successfully...")

    capturer = ScreenCapturer(fps=vnc_server.current_fps, quality=vnc_server.current_quality)
    vnc_server.screen_capturer = capturer
    try:
        vnc_server._apply_stream_profile("startup")
    except Exception:
        vnc_server.screen_capturer.quality = vnc_server.current_quality
    capturer.start()

    try:
        # Do NOT auto-open the local viewer in a browser. The machine is reachable only
        # via its public link (shown in Settings); localhost is not a usable entry point.
        asyncio.run(vnc_server.start_server())
    except OSError as e:
        err_no = getattr(e, "errno", None)
        if err_no == 10048:
            print(f"Bind failed: fixed port {vnc_server.port} is already in use.")
        else:
            raise
    except (KeyboardInterrupt, SystemExit):
        print("\nShutting down...")
    finally:
        capturer.stop()
        if vnc_server.tunnel_manager:
            vnc_server.tunnel_manager.cleanup()
    print("Goodbye!")


if __name__ == "__main__":
    main()
