"""HTTP/WebSocket server lifecycle for the Zadoo VNC runtime."""
from __future__ import annotations

import asyncio
import queue
import threading
import urllib.parse
from typing import Set

import websockets

from .network import get_local_ip
from .tunnel import CloudflareTunnelManager

from .input_control import InputControlMixin
from .media import MediaMixin
from .routes import RoutesMixin

class VNCServer(RoutesMixin, MediaMixin, InputControlMixin):

    def __init__(self, port, secondary_port):
        self.port = port
        self.secondary_port = secondary_port
        self.tunnel_manager = None
        self.enable_tunnel = True
        self.screen_capturer = None
        self.current_quality = 75
        self.current_fps = 60
        self.video_clients: Set[websockets.WebSocketServerProtocol] = set()
        self.audio_clients: Set[websockets.WebSocketServerProtocol] = set()
        self.input_clients: Set[websockets.WebSocketServerProtocol] = set()
        self.cursor_subscribers: Set[websockets.WebSocketServerProtocol] = set()
        self.cursor_broadcast_enabled = False
        self.audio_thread = None
        self.audio_queue = queue.Queue(maxsize=10)
        self.audio_running = False
        self.stop_event = asyncio.Event()
        self.loop = None
        # Audio quality targets
        self.audio_samplerate_target = 24000
        self.audio_channels_target = 1
        self.audio_blocksize_target = 960
        # Audio format state
        self.audio_samplerate = 24000
        self.audio_channels = 1
        self.audio_blocksize = 960
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
        # Typematic repeat state for non-modifier keys
        self._repeat_keys = {}
        # Alert presets for host controls
        self.alert_presets = {
            'A': ("Heads up", "Please check this now."),
            'B': ("Break", "Take a short break."),
            'C': ("Call", "Join the call."),
            'D': ("Stop", "Stop and review immediately."),
        }
        # Allow overriding alert presets from environment / .env
        try:
            self._load_alert_presets_from_env()
        except Exception:
            pass
        # Custom type-to-alert capture state
        self.custom_alert_active = False
        self.custom_alert_buf = []
        self._blocked_keys_in_capture = set()
        self._custom_alert_cooldown_until = 0.0
        self.auth_sessions = {}

    async def start_server(self):
        """Start WebSocket servers on both ports"""
        # Capture the running event loop for cross-thread broadcasts
        try:
            self.loop = asyncio.get_running_loop()
        except Exception:
            self.loop = None
        print("=" * 60)
        print("🎮 VNC SERVER STARTING...")
        print(f"🌐 Primary: http://localhost:{self.port}")
        print(f"🌐 Secondary: http://localhost:{self.secondary_port}")
        print(f"🏠 Network: http://{get_local_ip()}:{self.secondary_port}")
        print("🌍 Internet: Check above for public URL")
        print("=" * 60)
        # Start host hotkey capture on server start (A/B/C/D)
        try:
            self.start_global_keyboard_hook()
        except Exception as e:
            print(f"⚠️ Keyboard hook failed to initialize: {e}")
        # If hook isn't active, start fallback poller
        try:
            if not getattr(self, 'keyboard_hook_active', False):
                self.start_host_hotkey_poller()
                print("✅ Fallback hotkey poller started (A/B/C/D)")
            else:
                print("✅ Global keyboard hook armed for host alerts (A/B/C/D)")
        except Exception:
            pass
        
        # Check if screen capturer is working
        if self.screen_capturer:
            print("✅ Screen capturer initialized")
            # Wait a moment for it to capture first frame
            await asyncio.sleep(1)
            if self.screen_capturer.latest_frame_jpeg:
                print("✅ Screen capture working - first frame captured")
            else:
                print("⚠️  Screen capture not producing frames yet")
        else:
            print("❌ Screen capturer not initialized")
        
        # Start a tunnel only when enabled and one wasn't injected by app.main().
        if self.enable_tunnel and not self.tunnel_manager:
            self.tunnel_manager = CloudflareTunnelManager(self.port)
            threading.Thread(target=self.tunnel_manager.start_primary_tunnel, daemon=True).start()
                
        # Start broadcast task for video streaming
        broadcast_task = asyncio.create_task(self.broadcast_frames())
        
        # Start cursor broadcasting task
        cursor_broadcast_task = asyncio.create_task(self.broadcast_cursor_position())
        
        # Start both servers
        primary_server = await websockets.serve(
            self.main_handler,
            "0.0.0.0",
            self.port,
            process_request=self.process_request,
        )
        
        secondary_server = await websockets.serve(
            self.main_handler,
            "0.0.0.0",
            self.secondary_port,
            process_request=self.process_request,
        )
        
        print(f"✅ VNC servers running on ports {self.port} and {self.secondary_port}")
        print("🎬 Video streaming started")
        
        # Keep servers running
        await asyncio.gather(
            primary_server.wait_closed(),
            secondary_server.wait_closed(),
            broadcast_task,
            cursor_broadcast_task
        )

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
            await self.mic_stream_handler(websocket)
        elif route_path == "/ssh":
            await self.ssh_ws_handler(websocket)
        elif route_path == "/webcam":
            await self.webcam_stream_handler(websocket)
        else:
            # Default to video stream handler for compatibility
            await self.video_stream_handler(websocket)

    def stop(self):
        if self.loop:
            self.loop.call_soon_threadsafe(self.stop_event.set)
