"""Application entrypoint and startup orchestration."""
from __future__ import annotations

import asyncio
import ctypes
import logging
import os
import sys

from .config import APP_PORT, _load_dotenv, env_bool
from .logging_utils import _setup_logging_to_file


def require_windows_x64():
    if sys.platform != "win32":
        raise RuntimeError(f"Zadoo requires Windows; current platform is {sys.platform}")
    if ctypes.sizeof(ctypes.c_void_p) != 8:
        raise RuntimeError("Zadoo requires a 64-bit Python runtime")


def main():
    require_windows_x64()
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
    if sys.stdout and hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    _load_dotenv()
    if getattr(sys, "frozen", False) or env_bool("HIDE_CONSOLE"):
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd:
            ctypes.windll.user32.ShowWindow(hwnd, 0)
    _setup_logging_to_file()
    log_level_name = os.environ.get("ZADOO_LOG_LEVEL", "INFO").upper()
    try:
        log_level = logging.getLevelNamesMapping()[log_level_name]
    except KeyError as exc:
        raise ValueError(f"Invalid ZADOO_LOG_LEVEL: {log_level_name}") from exc
    logging.basicConfig(
        level=log_level,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )
    for logger_name in (
        "asyncio",
        "websockets.server",
        "websockets",
        "comtypes",
        "comtypes.client",
        "comtypes.client._code_cache",
        "comtypes._post_coinit",
    ):
        logging.getLogger(logger_name).setLevel(logging.WARNING)
    from .settings import get_settings_store

    args = {arg.lower() for arg in sys.argv[1:]}
    unknown_args = sorted(args - {"--open", "--settings"})
    if unknown_args:
        raise ValueError(f"Unsupported Zadoo arguments: {', '.join(unknown_args)}")
    if args:
        from .settings_window import run_settings_window

        run_settings_window()
        return

    from .screen_capture import ScreenCapturer
    from .server import VNCServer
    from .tunnel import CloudflareTunnelManager

    settings_store = get_settings_store()
    signed_in = bool(settings_store.get_device_token())
    if signed_in and not settings_store.configured():
        raise RuntimeError("Access code is not configured; set it in Zadoo Settings")
    tunnel_disabled_by_env = env_bool("ZADOO_DISABLE_TUNNEL")
    if not signed_in and not tunnel_disabled_by_env:
        raise RuntimeError("Zadoo is not signed in; sign in from Zadoo Settings")
    if signed_in and not tunnel_disabled_by_env:
        from .saas import ZadooCloudClient

        cloud_result = ZadooCloudClient(settings_store).entitlement()
        if not cloud_result["success"]:
            raise RuntimeError(f"Cloud entitlement check failed: {cloud_result['error']}")
        entitlement = cloud_result["entitlement"]
        if entitlement["revoked"]:
            raise RuntimeError(f"Cloud entitlement check failed: {entitlement['reason']}")
    use_tunnel = signed_in and not tunnel_disabled_by_env
    tunnel_manager = None
    if use_tunnel:
        tunnel_manager = CloudflareTunnelManager(primary_port=APP_PORT)

    vnc_server = VNCServer(APP_PORT)
    vnc_server.enable_tunnel = use_tunnel
    vnc_server.tunnel_block_reason = "" if use_tunnel else "Tunnel disabled by ZADOO_DISABLE_TUNNEL"
    if tunnel_manager:
        vnc_server.tunnel_manager = tunnel_manager
        print("Tunnel will launch after the local server binds successfully...")

    capturer = ScreenCapturer(fps=vnc_server.current_fps, quality=vnc_server.current_quality)
    vnc_server.screen_capturer = capturer
    vnc_server._apply_stream_profile("startup")
    capturer.start()

    try:
        # Settings displays the public link; starting the host does not open a browser.
        asyncio.run(vnc_server.start_server())
    except OSError as e:
        if e.errno == 10048:
            raise RuntimeError(f"Fixed port {vnc_server.port} is already in use") from e
        raise
    except (KeyboardInterrupt, SystemExit):
        print("\nShutting down...")
    finally:
        capturer.stop()
        capturer.join(timeout=2)
        if signed_in and vnc_server.ever_bound:
            try:
                from .saas import ZadooCloudClient

                offline = ZadooCloudClient(settings_store).go_offline()
                if not offline["success"]:
                    logging.error("Cloud offline notification failed: %s", offline["error"])
            except Exception as exc:
                logging.error("Cloud offline notification failed: %s", exc, exc_info=True)
        if vnc_server.tunnel_manager:
            vnc_server.tunnel_manager.cleanup()
        if capturer.is_alive():
            raise RuntimeError("Screen capture thread did not stop within 2 seconds")


if __name__ == "__main__":
    main()
