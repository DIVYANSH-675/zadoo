"""HTTP/WebSocket server lifecycle for the Zadoo VNC runtime."""
from __future__ import annotations

import asyncio
import queue
import threading
import urllib.parse
from typing import Set

import websockets

from .tunnel import CloudflareTunnelManager

from .input_control import InputControlMixin
from .logging_utils import _log_fallback
from .media import MediaMixin
from .routes import RoutesMixin
from .settings import DEFAULT_PERMISSIONS, get_settings_store
from .streaming import AdaptiveStreamController, detect_encoder_capabilities

class VNCServer(RoutesMixin, MediaMixin, InputControlMixin):

    def __init__(self, port):
        self.port = port
        self.settings_store = get_settings_store()
        self.runtime_settings = self.settings_store.load()
        self.tunnel_manager = None
        self.enable_tunnel = True
        self.screen_capturer = None
        self.current_quality = 85
        self._quality_locked_by_user = True
        self.current_fps = 0
        self.encoder_capabilities = detect_encoder_capabilities()
        self.adaptive_stream = AdaptiveStreamController(self.encoder_capabilities)
        if self.adaptive_stream.enabled:
            startup_profile = self.adaptive_stream.profile
            self.current_fps = startup_profile.target_fps
        self.video_clients: Set[websockets.WebSocketServerProtocol] = set()
        self.audio_clients: Set[websockets.WebSocketServerProtocol] = set()
        self.mic_clients: Set[websockets.WebSocketServerProtocol] = set()
        self.input_clients: Set[websockets.WebSocketServerProtocol] = set()
        self.cursor_subscribers: Set[websockets.WebSocketServerProtocol] = set()
        self.cursor_broadcast_enabled = False
        self.audio_queue = queue.Queue(maxsize=10)
        self.stop_event = None
        self.loop = None
        self.frame_ready_event = None
        self.selected_camera = None
        self.selected_camera_id = None
        self.keyboard_hook_active = False
        self.keystroke_capture_enabled = False
        self.keyboard_hook = None
        # Global keyboard hook suppression flag (true only when explicitly enabled)
        self._global_hook_suppress = False           # default: don't swallow anything
        self._synth_injecting = False                # guard to bypass suppression when we inject
        # mic streaming state
        self.mic_running = False
        self.mic_stream = None
        self.mic_queue = queue.Queue(maxsize=64)
        self.mic_samplerate = 48000
        self.mic_blocksize = 960
        self.mic_channels = 1
        self.mic_device_id = "default"
        self.mic_device_name = "System Default"
        # Typematic repeat state for non-modifier keys
        self._repeat_keys = {}
        # Alert presets are user-defined. Blank slots are disabled.
        self.alert_presets = {}
        self._load_alert_presets_from_settings()
        # Custom type-to-alert capture state
        self.custom_alert_active = False
        self.custom_alert_buf = []
        self._blocked_keys_in_capture = set()
        self._custom_alert_cooldown_until = 0.0
        self.auth_sessions = {}
        self._auth_failures = {}
        self._runtime_auth_code = ""
        self._runtime_auth_generated = False
        self._video_send_tasks = {}
        self._video_skipped_sends = 0
        self._servers = []
        self._hotkey_lock = threading.RLock()
        self.live_typing_text_by_client = {}

    def _load_alert_presets_from_settings(self):
        try:
            settings = self.settings_store.load(reload=True)
            self.runtime_settings = settings
            presets = {}
            for code, item in (settings.get("alerts") or {}).items():
                if not isinstance(item, dict) or not item.get("enabled"):
                    continue
                title = str(item.get("title") or "").strip()
                message = str(item.get("message") or "").strip()
                if title or message:
                    presets[str(code).upper()] = (title or "Alert", message)
            self.alert_presets = presets
            return presets
        except Exception:
            self.alert_presets = {}
            return {}

    def _settings_configured(self):
        try:
            return bool(self.settings_store.configured())
        except Exception:
            return False

    def _profile_permissions(self, profile_id):
        try:
            settings = self.settings_store.load(reload=True)
            profile = (settings.get("profiles") or {}).get(str(profile_id or ""))
            if isinstance(profile, dict):
                perms = dict(DEFAULT_PERMISSIONS)
                perms.update({k: bool(v) for k, v in (profile.get("permissions") or {}).items() if k in perms})
                return perms
        except Exception:
            pass
        return dict(DEFAULT_PERMISSIONS)

    async def start_server(self):
        """Start WebSocket servers."""
        # Capture the running event loop for cross-thread broadcasts
        try:
            self.loop = asyncio.get_running_loop()
        except Exception:
            self.loop = None
        self.stop_event = asyncio.Event()
        self.frame_ready_event = asyncio.Event()
        if self.screen_capturer and hasattr(self.screen_capturer, "set_frame_event"):
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
        )

        servers = [primary_server]
        self._servers = servers

        # Start host hotkey capture on server start (A/B/C/D)
        try:
            hook_started = bool(self.start_global_keyboard_hook())
        except Exception as e:
            hook_started = False
            _log_fallback("host_hotkeys", "poller", str(e), e)
            print(f" Keyboard hook failed to initialize: {e}")
        # If hook isn't active, start fallback poller
        try:
            if not hook_started or not getattr(self, 'keyboard_hook_active', False) or not getattr(self, '_alert_hotkeys_active', False):
                _log_fallback("host_hotkeys", "poller", "keyboard_hook_inactive")
                self.start_host_hotkey_poller()
                print(" Fallback hotkey poller started (A/B/C/D)")
            else:
                print(" Global keyboard hook armed for host alerts (A/B/C/D)")
        except Exception:
            pass
        
        # Check if screen capturer is working
        if self.screen_capturer:
            print(" Screen capturer initialized")
            # Wait a moment for it to capture first frame
            await asyncio.sleep(1)
            if self.screen_capturer.latest_frame_jpeg:
                print(" Screen capture working - first frame captured")
            else:
                print("  Screen capture not producing frames yet")
        else:
            print(" Screen capturer not initialized")
        
        # Start a tunnel only after the server has bound successfully.
        if self.enable_tunnel:
            if not self.tunnel_manager:
                self.tunnel_manager = CloudflareTunnelManager(self.port)
            threading.Thread(target=self.tunnel_manager.start_primary_tunnel, daemon=True).start()
                
        # Start broadcast task for video streaming
        broadcast_task = asyncio.create_task(self.broadcast_frames())
        
        # Start cursor broadcasting task
        cursor_broadcast_task = asyncio.create_task(self.broadcast_cursor_position())
        
        print(f" VNC server running on port: {self.port}")
        print(" Video streaming started")
        
        # Keep servers running until stop() or task cancellation requests shutdown.
        try:
            await self.stop_event.wait()
        finally:
            try:
                self._host_hotkey_poller_active = False
                self.stop_global_keyboard_hook()
            except Exception:
                pass
            for server in servers:
                try:
                    server.close()
                except Exception:
                    pass
            for task in (broadcast_task, cursor_broadcast_task):
                if not task.done():
                    task.cancel()
            await asyncio.gather(*(server.wait_closed() for server in servers), return_exceptions=True)
            await asyncio.gather(broadcast_task, cursor_broadcast_task, return_exceptions=True)
            self._servers = []

    def set_tunnel_manager(self, tunnel_manager):
        """Set the tunnel manager for API access"""
        self.tunnel_manager = tunnel_manager

    async def main_handler(self, websocket):
        """Handles incoming connections and routes them (websockets v15 ServerConnection)."""
        # Support both legacy and v15 APIs
        path = getattr(websocket, 'path', None)
        if not isinstance(path, str):
            try:
                path = websocket.request.path
            except Exception:
                path = '/video'
        route_path = urllib.parse.urlparse(path).path
        try:
            request_headers = getattr(getattr(websocket, "request", None), "headers", None)
            if request_headers is None:
                request_headers = getattr(websocket, "request_headers", None)
        except Exception:
            request_headers = None
        try:
            if not self._request_origin_allowed(request_headers):
                await websocket.close(code=1008, reason="Forbidden")
                return
            if not self._is_ws_authorized(route_path, request_headers):
                await websocket.close(code=1008, reason="Forbidden")
                return
        except Exception:
            await websocket.close(code=1008, reason="Forbidden")
            return
        
        if route_path == "/video":
            await self.video_stream_handler(websocket)
        elif route_path == "/input":
            await self.input_event_handler(websocket)
        elif route_path == "/audio":
            await self.audio_stream_handler(websocket)
        elif route_path == "/mic":
            await self.mic_stream_handler(websocket, path=path)
        elif route_path == "/ssh":
            await self.local_shell_ws_handler(websocket)
        elif route_path == "/webcam":
            await self.webcam_stream_handler(websocket)
        else:
            # Default to video stream handler for compatibility
            await self.video_stream_handler(websocket)

    def stop(self):
        if self.loop and self.stop_event is not None:
            def _stop():
                if self.stop_event is not None:
                    self.stop_event.set()
                for server in list(getattr(self, "_servers", []) or []):
                    try:
                        server.close()
                    except Exception:
                        pass
            self.loop.call_soon_threadsafe(_stop)
