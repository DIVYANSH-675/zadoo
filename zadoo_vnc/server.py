"""HTTP/WebSocket server lifecycle for the Zadoo VNC runtime."""
from __future__ import annotations

import asyncio
import logging
import os
import queue
import threading
import time
import urllib.parse
from contextlib import suppress

import websockets

from .config import env_bool, env_int
from .input_control import InputControlMixin
from .media import MediaMixin
from .routes import RoutesMixin
from .settings import _clean_http_origin, get_settings_store
from .streaming import AdaptiveStreamController
from .tunnel import CloudflareTunnelManager


class VNCServer(RoutesMixin, MediaMixin, InputControlMixin):
    VIEWER_ROUTES = frozenset({"/video", "/input", "/audio", "/mic", "/terminal", "/webcam"})

    def __init__(self, port):
        self.port = port
        self.settings_store = get_settings_store()
        self.tunnel_manager = None
        self.enable_tunnel = True
        self.tunnel_block_reason = ""
        self.screen_capturer = None
        self.current_quality = 85
        # Do NOT lock quality/full-resolution by default. The lock pins quality=85 -> scale_div=1
        # and disables the adaptive RESOLUTION ladder, so a slow (e.g. 2 Mbps) client can never be
        # downshifted in resolution and stays flooded. The lock is enabled only when the user
        # explicitly changes quality (see _apply_quality).
        self._quality_locked_by_user = False
        self._manual_performance = None
        self.adaptive_stream = AdaptiveStreamController()
        self.current_fps = self.adaptive_stream.profile.target_fps
        self.video_clients: set[websockets.WebSocketServerProtocol] = set()
        self._media_clients_lock = threading.Lock()
        self.audio_clients: dict[websockets.WebSocketServerProtocol, queue.Queue] = {}
        self.mic_clients: dict[websockets.WebSocketServerProtocol, queue.Queue] = {}
        self.input_clients: set[websockets.WebSocketServerProtocol] = set()
        self.cursor_subscribers: set[websockets.WebSocketServerProtocol] = set()
        self.cursor_broadcast_enabled = False
        self._audio_running = False
        self._audio_thread = None
        self._audio_stream = None
        self._audio_stop_evt = None
        self._audio_sr = 48000
        self._audio_error = None
        self.stop_event = None
        self.loop = None
        self.frame_ready_event = None
        self._alert_hotkeys_active = False
        self._hk_ids = []
        self._custom_capture_hotkeys = []
        self._controller_alert_seq = 0
        self._alert_preset_cooldowns = {}
        # mic streaming state
        self.mic_running = False
        self.mic_stream = None
        self.mic_samplerate = 48000
        self.mic_blocksize = 960
        self.mic_device_id = "default"
        self.mic_device_name = None
        self._mic_error = None
        # Alert presets are user-defined. Blank slots are disabled.
        self.alert_presets = {}
        self._load_alert_presets_from_settings()
        # Custom type-to-alert capture state
        self.custom_alert_active = False
        self.custom_alert_buf = []
        self.auth_sessions = {}
        self._auth_failures = {}
        self._allow_direct_access = env_bool("ZADOO_ALLOW_DIRECT_ACCESS")
        try:
            self._allowed_origins = frozenset(
                "*" if item == "*" else _clean_http_origin(item, "ZADOO_ALLOWED_ORIGINS entry").lower()
                for raw in os.getenv("ZADOO_ALLOWED_ORIGINS", "").split(",")
                if (item := raw.strip())
            )
        except ValueError as exc:
            raise ValueError(f"Invalid ZADOO_ALLOWED_ORIGINS: {exc}") from exc
        self._auth_window_seconds = env_int("ZADOO_AUTH_WINDOW_SECONDS", 60, 1, 3600)
        self._auth_max_failures = env_int("ZADOO_AUTH_MAX_FAILURES", 5, 1, 100)
        self._auth_lockout_seconds = env_int("ZADOO_AUTH_LOCKOUT_SECONDS", 300, 1, 86400)
        self._cloud_heartbeat_grace_seconds = env_int(
            "ZADOO_CLOUD_HEARTBEAT_GRACE_SECONDS", 120, 60, 120
        )
        self._max_viewers = env_int("ZADOO_MAX_VIEWERS", 5, 1, 5)
        self._clipboard_image_max_bytes = env_int(
            "ZADOO_CLIPBOARD_IMAGE_MAX_BYTES", 5_000_000, 1, 100_000_000
        )
        self._clipboard_text_max_bytes = env_int(
            "ZADOO_CLIPBOARD_TEXT_MAX_BYTES", 1_000_000, 1, 50_000_000
        )
        self._websocket_max_size = max(
            1_048_576,
            ((self._clipboard_image_max_bytes + 2) // 3) * 4 + 4096,
            self._clipboard_text_max_bytes * 6 + 4096,
        )
        self._video_send_tasks = {}
        self._background_send_tasks = set()
        self._video_skipped_sends = 0
        self._server = None
        self.ever_bound = False
        self._cloud_heartbeat_stop = threading.Event()
        self._cloud_heartbeat_thread = None
        self._fatal_error = None
        self._grace_locks_by_session = {}
        self._viewer_connections_by_session = {}
        self._hotkey_lock = threading.RLock()
        self._keyboard_cleanup_registered = False
        self.live_typing_text_by_client = {}

    def _load_alert_presets_from_settings(self):
        settings = self.settings_store.load(reload=True)
        presets = {}
        for code, item in settings["alerts"].items():
            if not item["enabled"]:
                continue
            presets[code] = (item["title"], item["message"])
        self.alert_presets = presets
        return presets

    async def start_server(self):
        """Start WebSocket servers."""
        self.loop = asyncio.get_running_loop()
        self.stop_event = asyncio.Event()
        self.frame_ready_event = asyncio.Event()
        self._fatal_error = None
        if not self.screen_capturer:
            raise RuntimeError("Screen capturer must be configured before the server starts")
        self.screen_capturer.set_frame_event(self.loop, self.frame_ready_event)
        self._announce_auth_codes()

        # Bind before starting tunnels or background stream tasks, so failed
        # starts cannot leave fresh cloudflared processes behind.
        primary_server = await websockets.serve(
            self.main_handler,
            "0.0.0.0",
            self.port,
            process_request=self.process_request,
            compression=None,
            max_size=self._websocket_max_size,
            max_queue=2,
        )

        self._server = primary_server
        self.ever_bound = True

        async def cleanup(tasks=()):
            self._cloud_heartbeat_stop.set()
            self.stop_host_hotkeys()
            primary_server.close()
            for task in tasks:
                if not task.done():
                    task.cancel()
            await primary_server.wait_closed()
            if tasks:
                await asyncio.gather(*tasks, return_exceptions=True)
            background_sends = tuple(self._background_send_tasks)
            for task in background_sends:
                task.cancel()
            if background_sends:
                await asyncio.gather(*background_sends, return_exceptions=True)
            self._background_send_tasks.clear()
            heartbeat_thread = self._cloud_heartbeat_thread
            if heartbeat_thread is not None and heartbeat_thread.is_alive():
                await asyncio.to_thread(heartbeat_thread.join, 9)
                if heartbeat_thread.is_alive():
                    raise RuntimeError("Cloud heartbeat thread did not stop within 9 seconds")
            self._cloud_heartbeat_thread = None
            self._server = None

        try:
            registered_hotkeys = self.start_host_hotkeys()
        except Exception as exc:
            await cleanup()
            raise RuntimeError(f"Host alert hotkey registration failed: {exc}") from exc
        print(f" Host alert hotkeys armed: {registered_hotkeys}")
        
        if not self.screen_capturer.latest_frame_jpeg:
            with suppress(asyncio.TimeoutError):
                await asyncio.wait_for(self.frame_ready_event.wait(), timeout=1.0)
        if not self.screen_capturer.latest_frame_jpeg:
            error = self.screen_capturer.last_error or "Screen capture produced no frame within one second"
            await cleanup()
            raise RuntimeError(error)
        self.screen_capturer.set_streaming_active(False)
        
        # Start a tunnel only after the server has bound successfully.
        try:
            if self.enable_tunnel:
                if not self.tunnel_manager:
                    self.tunnel_manager = CloudflareTunnelManager(self.port)
                await asyncio.to_thread(self.tunnel_manager.start_primary_tunnel)
            self._start_cloud_heartbeat()
        except Exception:
            await cleanup()
            raise
                
        # Start broadcast task for video streaming
        broadcast_task = asyncio.create_task(self.broadcast_frames())
        
        # Start cursor broadcasting task
        cursor_broadcast_task = asyncio.create_task(self.broadcast_cursor_position())
        
        print(f" VNC server running on port: {self.port}")
        print(" Video streaming started")
        
        stop_wait_task = asyncio.create_task(self.stop_event.wait())
        try:
            done, _ = await asyncio.wait(
                {stop_wait_task, broadcast_task, cursor_broadcast_task},
                return_when=asyncio.FIRST_COMPLETED,
            )
            self._raise_for_completed_required_tasks(
                done, stop_wait_task, self._fatal_error
            )
        finally:
            await cleanup((stop_wait_task, broadcast_task, cursor_broadcast_task))

    @staticmethod
    def _raise_for_completed_required_tasks(done, stop_wait_task, fatal_error):
        # Graceful stop can wake the waiter at the same moment that media loops
        # observe the stop event and finish. Stop must win that benign race.
        if stop_wait_task in done:
            if fatal_error is not None:
                raise fatal_error
            return
        for task in done:
            task.result()
        raise RuntimeError("A required media broadcast task stopped unexpectedly")

    def _start_cloud_heartbeat(self):
        if not self.settings_store.get_device_token():
            return
        self._cloud_heartbeat_stop.clear()
        stop_event = self._cloud_heartbeat_stop

        def _loop():
            try:
                from .saas import ZadooCloudClient

                client = ZadooCloudClient(self.settings_store)
                seen_url = not self.enable_tunnel
                failure_started = None
                consecutive_failures = 0
                while not stop_event.is_set():
                    try:
                        process = self.tunnel_manager.primary_tunnel_process if self.tunnel_manager else None
                        if process is not None and process.poll() is not None:
                            raise RuntimeError(f"cloudflared exited with code {process.returncode}")
                        public_url = self.tunnel_manager.get_current_url() if self.tunnel_manager else None
                        result = client.heartbeat(public_url)
                        if not result["success"]:
                            raise RuntimeError(f"Cloud heartbeat failed: {result['error']}")
                    except Exception as exc:
                        now = time.monotonic()
                        try:
                            failure_started, consecutive_failures, retry_delay, elapsed = (
                                self._heartbeat_retry_policy(
                                    failure_started,
                                    consecutive_failures,
                                    now,
                                    exc,
                                )
                            )
                        except RuntimeError as final_error:
                            raise final_error from exc
                        logging.warning(
                            "Cloud heartbeat failed; retrying in %.1f seconds (%.1f/%ss): %s",
                            retry_delay,
                            elapsed,
                            self._cloud_heartbeat_grace_seconds,
                            exc,
                        )
                        if stop_event.wait(retry_delay):
                            break
                        continue
                    failure_started = None
                    consecutive_failures = 0
                    if public_url:
                        seen_url = True
                    # Heartbeat quickly until the tunnel URL first appears (so the website
                    # link syncs within seconds of startup), then settle to once a minute.
                    if stop_event.wait(60 if seen_url else 5):
                        break
            except Exception as exc:
                logging.exception("Required cloud heartbeat stopped")
                self._fatal_error = RuntimeError(str(exc))
                if self.loop and self.stop_event:
                    self.loop.call_soon_threadsafe(self.stop_event.set)

        self._cloud_heartbeat_thread = threading.Thread(target=_loop, daemon=True)
        self._cloud_heartbeat_thread.start()

    def _heartbeat_retry_policy(self, failure_started, consecutive_failures, now, error):
        if failure_started is None:
            failure_started = now
        consecutive_failures += 1
        elapsed = now - failure_started
        if elapsed >= self._cloud_heartbeat_grace_seconds:
            raise RuntimeError(
                "Cloud heartbeat failed for "
                f"{self._cloud_heartbeat_grace_seconds} seconds: {error}"
            )
        retry_delay = min(
            30,
            5 * (2 ** min(consecutive_failures - 1, 3)),
            self._cloud_heartbeat_grace_seconds - elapsed,
        )
        return failure_started, consecutive_failures, retry_delay, elapsed

    def _register_viewer_connection(self, websocket, route_path):
        if route_path not in self.VIEWER_ROUTES:
            return None
        token = self._cookie_value(self._headers_for_websocket(websocket), self.AUTH_COOKIE_NAME)
        if not token:
            raise PermissionError("Authenticated viewer session is missing")
        sessions = self._viewer_connections_by_session
        connections = sessions.get(token)
        if connections is None:
            if len(sessions) >= self._max_viewers:
                raise RuntimeError(f"Viewer limit reached ({self._max_viewers})")
            connections = set()
            sessions[token] = connections
        connections.add(websocket)
        return token

    def _unregister_viewer_connection(self, token, websocket):
        if token is None:
            return
        connections = self._viewer_connections_by_session.get(token)
        if connections is None:
            return
        connections.discard(websocket)
        if not connections:
            del self._viewer_connections_by_session[token]

    async def main_handler(self, websocket):
        """Handles incoming connections and routes them (websockets v15 ServerConnection)."""
        path = websocket.request.path
        route_path = urllib.parse.urlparse(path).path
        request_headers = websocket.request.headers
        try:
            if not self._request_origin_allowed(request_headers):
                await websocket.close(code=1008, reason="Forbidden")
                return
            if route_path == "/rpc":
                if not self._rpc_connection_allowed(request_headers, websocket):
                    await websocket.close(code=1008, reason="Forbidden")
                    return
            elif not self._is_ws_authorized(route_path, request_headers):
                await websocket.close(code=1008, reason="Forbidden")
                return
        except Exception:
            logging.exception("WebSocket authorization failed")
            await websocket.close(code=1008, reason="Forbidden")
            return
        
        viewer_token = None
        try:
            try:
                viewer_token = self._register_viewer_connection(websocket, route_path)
            except RuntimeError as exc:
                await websocket.close(code=1013, reason=str(exc))
                return
            if route_path == "/rpc":
                await self.rpc_handler(websocket)
            elif route_path == "/video":
                await self.video_stream_handler(websocket)
            elif route_path == "/input":
                await self.input_event_handler(websocket)
            elif route_path == "/audio":
                await self.audio_stream_handler(websocket)
            elif route_path == "/mic":
                await self.mic_stream_handler(websocket, path=path)
            elif route_path == "/terminal":
                await self.local_shell_ws_handler(websocket)
            elif route_path == "/webcam":
                await self.webcam_stream_handler(websocket)
            else:
                await websocket.close(code=1008, reason=f"Unsupported WebSocket route: {route_path}")
        finally:
            self._unregister_viewer_connection(viewer_token, websocket)

    def stop(self):
        if self.loop and self.stop_event is not None:
            def _stop():
                if self.stop_event is not None:
                    self.stop_event.set()
                if self._server is not None:
                    self._server.close()
            self.loop.call_soon_threadsafe(_stop)
