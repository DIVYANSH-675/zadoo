"""Remote input, clipboard, alerts, and hotkey controls."""
from __future__ import annotations

import asyncio
import atexit
import ctypes
import json
import logging
import os
import re
import subprocess
import sys
import threading
import time
from contextlib import contextmanager

import websockets

from .config import env_int
from .dependencies import HAS_KEYBOARD, HAS_PYAUTOGUI, HAS_PYPERCLIP, keyboard, pyautogui, pyperclip
from .logging_utils import _log_except, _log_fallback, _log_try_ok
from .win32_input import (
    INPUT,
    MOUSEEVENTF_LEFTDOWN,
    MOUSEEVENTF_LEFTUP,
    MOUSEEVENTF_MIDDLEDOWN,
    MOUSEEVENTF_MIDDLEUP,
    MOUSEEVENTF_RIGHTDOWN,
    MOUSEEVENTF_RIGHTUP,
    MOUSEEVENTF_WHEEL,
    MOUSEINPUT,
    SM_CXVIRTUALSCREEN,
    SM_CYVIRTUALSCREEN,
    SM_XVIRTUALSCREEN,
    SM_YVIRTUALSCREEN,
    _GetAsyncKeyState,
    _sendinput_mouse_button,
    _sendinput_mouse_move_abs,
    user32,
)

class InputControlMixin:
    ALERT_PRESET_HOTKEYS = {
        'A': ('shift+numpad 5', 'shift+clear'),
        'B': ('shift+numpad 0', 'shift+insert'),
        'C': ('shift+numpad 8', 'shift+up'),
        'D': ('shift+numpad 2', 'shift+down'),
    }
    ALERT_PRESET_KEY_NAMES = {
        'A': ('num 5', 'numpad 5', 'kp_5', 'num5', 'clear'),
        'B': ('num 0', 'numpad 0', 'kp_0', 'num0', 'insert'),
        'C': ('num 8', 'numpad 8', 'kp_8', 'num8', 'up'),
        'D': ('num 2', 'numpad 2', 'kp_2', 'num2', 'down'),
    }
    ALERT_PRESET_VK_KEYS = {
        'A': (0x65, 0x0C),  # NumPad5 / Clear
        'B': (0x60, 0x2D),  # NumPad0 / Insert
        'C': (0x68, 0x26),  # NumPad8 / Up
        'D': (0x62, 0x28),  # NumPad2 / Down
    }

    def _get_hotkey_lock(self):
        lock = getattr(self, "_hotkey_lock", None)
        if lock is None:
            lock = threading.RLock()
            self._hotkey_lock = lock
        return lock

    def _register_keyboard_cleanup(self):
        if getattr(self, "_keyboard_cleanup_registered", False):
            return
        try:
            atexit.register(self.stop_global_keyboard_hook)
            self._keyboard_cleanup_registered = True
        except Exception:
            pass

    def _get_send_loop(self, log_name=None):
        loop = getattr(self, "loop", None)
        if loop:
            return loop
        try:
            loop = asyncio.get_running_loop()
            self.loop = loop
            if log_name:
                _log_try_ok(log_name)
            return loop
        except Exception as exc:
            if log_name:
                _log_except(log_name, exc)
            return None

    def _queue_ws_send(self, websocket, payload, log_name):
        if websocket is None:
            return False
        try:
            loop = asyncio.get_running_loop()
            loop.create_task(websocket.send(payload))
            return True
        except RuntimeError:
            loop = self._get_send_loop(f"{log_name}.get_loop")
            if not loop:
                return False
            try:
                asyncio.run_coroutine_threadsafe(websocket.send(payload), loop)
                return True
            except Exception as exc:
                _log_except(f"{log_name}.queue", exc)
                return False
        except Exception as exc:
            _log_except(f"{log_name}.queue", exc)
            return False

    def _load_alert_presets_from_env(self):
        """Override alert presets (A/B/C/D) from environment variables.
        Supported keys per preset (e.g. for A):
          - ALERT_A_TITLE, ALERT_A_MESSAGE
          - ALERT_A="Title|Message" (or "Title::Message")
        """
        def _pair_for(key: str):
            base = f"ALERT_{key}"
            title = os.getenv(f"{base}_TITLE")
            message = os.getenv(f"{base}_MESSAGE")
            combined = os.getenv(base)
            if (not title and not message) and combined:
                if '|' in combined:
                    parts = combined.split('|', 1)
                elif '::' in combined:
                    parts = combined.split('::', 1)
                else:
                    parts = [combined, '']
                title = (parts[0] or '').strip()
                message = (parts[1] or '').strip() if len(parts) > 1 else ''
            return (title, message)
        changed = []
        for k in ('A','B','C','D'):
            t, m = _pair_for(k)
            if (t and t.strip()) or (m and m.strip()):
                cur = self.alert_presets.get(k, ("Alert", k))
                new_title = (t or '').strip() or cur[0]
                new_message = (m or '').strip() or cur[1]
                self.alert_presets[k] = (new_title, new_message)
                changed.append(k)
        try:
            if changed:
                print(f"[alerts] Presets overridden from env for: {', '.join(changed)}")
            else:
                print("[alerts] Using default alert presets (no env overrides)")
        except Exception:
            pass

    def _capture_char(self, ch: str):
        try:
            if getattr(self, "custom_alert_active", False):
                self.custom_alert_buf.append(ch)
                _log_try_ok("_capture_char.append", ch)
        except Exception as e:
            _log_except("_capture_char", e)

    def _capture_backspace(self):
        try:
            if getattr(self, "custom_alert_active", False) and self.custom_alert_buf:
                self.custom_alert_buf.pop()
                _log_try_ok("_capture_backspace.pop")
        except Exception as e:
            _log_except("_capture_backspace", e)

    async def input_event_handler(self, websocket):
        print(f" New input client connected from {websocket.remote_address}")
        self.input_clients.add(websocket)
        _log_try_ok("input_event_handler.connect", str(getattr(websocket, 'remote_address', '')))
        self._get_send_loop("input_event_handler.grab_loop")
        
        try:
            # On connection, immediately get the URL for the client UI
            await self.handle_get_url_via_websocket(websocket)

            async for message in websocket:
                try:
                    event = json.loads(message)
                    action = event.get('action')
                    _log_try_ok("input_event_handler.message", action or '')
                    if not self._is_ws_action_authorized(websocket, action):
                        await self._send_ws_forbidden(websocket, action)
                        continue

                    # Control toggle: block/unblock host input and cursor broadcast
                    if action == 'control':
                        try:
                            self.block_host_input = bool(event.get('block_host_input', False))
                            if 'cursor_broadcast' in event:
                                self.cursor_broadcast_enabled = bool(event.get('cursor_broadcast'))
                            _log_try_ok('input_event_handler.control', f"block={self.block_host_input} cursor={self.cursor_broadcast_enabled}")
                        except Exception as e:
                            _log_except('input_event_handler.control', e)
                        continue

                    # This handler should ONLY process input actions or cursor subscriptions
                    if action in ['click', 'move', 'drag', 'key', 'key_combo', 'scroll', 'type_text', 'get_clipboard', 'set_clipboard', 'set_clipboard_image']:
                        # Drop input while blocked
                        if getattr(self, 'block_host_input', False):
                            _log_try_ok('input_event_handler.blocked', action)
                            continue
                        self.process_event(event, websocket)
                    elif action == 'snap_event':
                        # Client-side instrumentation events to measure perceived latency
                        # Expect event like {action:'snap_event', phase:'right_click'|'selection_start'|'selection_end'|'request_sent'|'response_received'}
                        try:
                            phase = str(event.get('phase', ''))
                            now = time.time()
                            if phase:
                                # Store timestamps on server for delta calculations
                                if not hasattr(self, '_snap_ts'):
                                    self._snap_ts = {}
                                self._snap_ts[phase] = now
                                # Compute useful deltas when possible
                                if phase == 'selection_end' and 'selection_start' in self._snap_ts:
                                    dt = (self._snap_ts['selection_end'] - self._snap_ts['selection_start']) * 1000.0
                                    logging.info("[snap.client] selection_ms=%.1f", dt)
                                if phase == 'response_received' and 'request_sent' in self._snap_ts:
                                    dt = (self._snap_ts['response_received'] - self._snap_ts['request_sent']) * 1000.0
                                    logging.info("[snap.client] request_to_response_ms=%.1f", dt)
                                if phase in ('right_click','left_click'):
                                    logging.info("[snap.client] click=%s", phase)
                        except Exception:
                            _log_except('input_event_handler.snap_event', sys.exc_info()[1])
                    elif action == 'mic_event':
                        try:
                            phase = str(event.get('phase', 'unknown'))[:64]
                            detail = str(event.get('detail', ''))[:240]
                            socket_state = str(event.get('socketState', 'unknown'))[:32]
                            audio_state = str(event.get('audioContextState', 'unknown'))[:32]
                            page_ms = event.get('pageMs')
                            logging.getLogger("mic").info(
                                "Mic client button event: phase=%s detail=%s playing=%s socket_state=%s audio_context=%s audio_state=%s page_ms=%s remote=%s",
                                phase,
                                detail,
                                bool(event.get('playing', False)),
                                socket_state,
                                bool(event.get('hasAudioContext', False)),
                                audio_state,
                                page_ms,
                                getattr(websocket, 'remote_address', None),
                            )
                        except Exception:
                            _log_except('input_event_handler.mic_event', sys.exc_info()[1])
                    elif action == 'refresh_tunnel':
                        await self.handle_refresh_via_websocket(websocket)
                    elif action == 'get_public_url':
                        await self.handle_get_url_via_websocket(websocket)
                    elif action == 'toggle_keystroke_capture':
                        enabled = bool(event.get('enabled', False))
                        if enabled:
                            self.enable_keystroke_capture()
                        else:
                            self.disable_keystroke_capture()
                        try:
                            await websocket.send(json.dumps({
                                'type': 'keystroke_capture_status',
                                'enabled': self.keystroke_capture_enabled
                            }))
                        except Exception:
                            pass
                    elif action == 'cursor_broadcast':
                        enabled = bool(event.get('enabled', False))
                        if enabled:
                            self.cursor_subscribers.add(websocket)
                            self.cursor_broadcast_enabled = True
                            print(f" Cursor broadcast enabled for {websocket.remote_address}")
                        else:
                            self.cursor_subscribers.discard(websocket)
                            if not self.cursor_subscribers:
                                self.cursor_broadcast_enabled = False
                            print(f" Cursor broadcast disabled for {websocket.remote_address}")

                except json.JSONDecodeError:
                    logging.warning(f"Received non-JSON input message: {message}")
                    _log_except("input_event_handler.json", sys.exc_info()[1])
                except Exception as e:
                    logging.error(f"Error processing input message: {e}", exc_info=True)
                    _log_except("input_event_handler.message", e)
        
        except websockets.exceptions.ConnectionClosed:
            print(f" Input client {websocket.remote_address} disconnected")
            _log_try_ok("input_event_handler.disconnect", str(getattr(websocket, 'remote_address', '')))
        finally:
            # Clean up on disconnect and ALWAYS re-enable host input
            self.block_host_input = False
            self.input_clients.discard(websocket)
            self.cursor_subscribers.discard(websocket)
            try:
                if hasattr(self, 'live_typing_text_by_client'):
                    self.live_typing_text_by_client.pop(websocket, None)
            except Exception:
                pass
            _log_try_ok("input_event_handler.cleanup")

    async def handle_refresh_via_websocket(self, websocket):
        """Handle tunnel refresh via WebSocket"""
        try:
            print(" WebSocket refresh request received")
            await websocket.send(json.dumps({
                'type': 'refresh_status',
                'message': f'Refreshing tunnel on current port...'
            }))
            payload = await self._refresh_tunnel_payload()
            payload["type"] = "refresh_complete"
            await websocket.send(json.dumps(payload))
                
        except Exception as e:
            print(f" Error in handle_refresh_via_websocket: {e}")
            await websocket.send(json.dumps({
                'type': 'refresh_complete',
                'success': False,
                'error': f'Error refreshing tunnel: {str(e)}'
            }))

    async def handle_get_url_via_websocket(self, websocket):
        """Handle public URL request via WebSocket"""
        try:
            payload = self._public_url_payload()
            payload["type"] = "public_url_response"
            await websocket.send(json.dumps(payload))
        except Exception as e:
            await websocket.send(json.dumps({
                'type': 'public_url_response',
                'success': False,
                'error': f'Error getting URL: {str(e)}'
            }))

    @contextmanager
    def _injection_guard(self):
        setattr(self, "_synth_injecting", True)
        try:
            yield
        finally:
            setattr(self, "_synth_injecting", False)

    def process_event(self, event, websocket=None):
        action = event.get('action')
        
        try:
            if action in ['click', 'move', 'drag']:
                with self._injection_guard():
                    self._handle_mouse_event(event)
            elif action == 'key':
                with self._injection_guard():
                    self._handle_key_event(event)
            elif action == 'key_combo':
                with self._injection_guard():
                    self._handle_key_combo(event)
            elif action == 'scroll':
                with self._injection_guard():
                    self._handle_scroll_event(event)
            elif action == 'type_text':
                with self._injection_guard():
                    self._handle_type_text(event, websocket)
            elif action == 'get_clipboard':
                if websocket is not None:
                    self._handle_get_clipboard(websocket)
            elif action == 'set_clipboard':
                self._handle_set_clipboard(event, websocket)
            elif action == 'set_clipboard_image':
                self._handle_set_clipboard_image(event, websocket)
        except Exception as e:
            print(f"Error processing event: {e}")

    def _send_clipboard_image_result(self, websocket, payload):
        if websocket is None:
            return
        try:
            self._queue_ws_send(websocket, json.dumps(payload), "_send_clipboard_image_result")
        except Exception as e:
            logging.warning("Failed to send clipboard image result: %s", e)

    def _handle_set_clipboard_image(self, event, websocket=None):
        request_id = event.get("request_id")
        try:
            import base64
            import tempfile

            mime = str(event.get("mime") or "image/png").lower()
            if mime not in {"image/png", "image/jpeg", "image/jpg"}:
                raise ValueError("unsupported image MIME type")
            data_b64 = str(event.get("data_base64") or "")
            if "," in data_b64:
                data_b64 = data_b64.split(",", 1)[1]
            if not data_b64:
                raise ValueError("missing image data")

            max_bytes = env_int("ZADOO_CLIPBOARD_IMAGE_MAX_BYTES", 5_000_000, 1, 100_000_000)
            if len(data_b64.encode("ascii", errors="ignore")) > ((max_bytes + 2) // 3) * 4 + 4096:
                raise ValueError("image clipboard payload exceeds size limit")
            image_data = base64.b64decode(data_b64, validate=True)
            if not image_data:
                raise ValueError("empty image data")
            if len(image_data) > max_bytes:
                raise ValueError("image clipboard payload exceeds size limit")

            suffix = ".jpg" if "jpeg" in mime or "jpg" in mime else ".png"
            temp_file_path = None
            try:
                with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as temp_file:
                    temp_file.write(image_data)
                    temp_file_path = temp_file.name

                if os.name != "nt":
                    raise RuntimeError("no image clipboard backend available")
                ps_cmd = (
                    "Add-Type -AssemblyName System.Windows.Forms; "
                    "Add-Type -AssemblyName System.Drawing; "
                    "$img=[System.Drawing.Image]::FromFile($env:ZADOO_CLIPBOARD_IMAGE_PATH); "
                    "[System.Windows.Forms.Clipboard]::SetImage($img); "
                    "$img.Dispose()"
                )
                child_env = os.environ.copy()
                child_env["ZADOO_CLIPBOARD_IMAGE_PATH"] = temp_file_path
                result = subprocess.run(
                    ["powershell", "-NoProfile", "-Command", ps_cmd],
                    capture_output=True,
                    text=True,
                    env=child_env,
                    timeout=10,
                )
                if result.returncode != 0:
                    raise RuntimeError(result.stderr.strip() or "PowerShell clipboard operation failed")
            finally:
                if temp_file_path:
                    try:
                        os.unlink(temp_file_path)
                    except Exception:
                        pass

            self._send_clipboard_image_result(websocket, {
                "type": "clipboard_image_result",
                "success": True,
                "request_id": request_id,
                "bytes": len(image_data),
            })
            print(f" Image copied to clipboard successfully ({len(image_data)} bytes)")
        except Exception as e:
            self._send_clipboard_image_result(websocket, {
                "type": "clipboard_image_result",
                "success": False,
                "request_id": request_id,
                "error": str(e),
            })
            print(f" Failed to copy image to clipboard: {e}")

    def _handle_get_clipboard(self, websocket):
        try:
            content = None
            if HAS_PYPERCLIP:
                try:
                    content = pyperclip.paste()
                    print(f" Clipboard content retrieved via pyperclip: {len(content or '')} chars")
                except Exception as e:
                    print(f" pyperclip failed: {e}")
                    _log_fallback("clipboard.get_text", "powershell_get_clipboard", str(e), e)
                    content = None
            
            if content is None:
                # Fallback via PowerShell (Windows)
                _log_fallback("clipboard.get_text", "powershell_get_clipboard", "pyperclip_unavailable_or_failed")
                try:
                    ps = subprocess.run(['powershell', '-NoProfile', '-Command', 'Get-Clipboard -Raw'], 
                                     capture_output=True, text=True, timeout=5)
                    if ps.returncode == 0 and ps.stdout:
                        content = ps.stdout
                        print(f" Clipboard content retrieved via PowerShell: {len(content)} chars")
                    else:
                        print(f" PowerShell clipboard failed: returncode={ps.returncode}, stderr={ps.stderr}")
                except subprocess.TimeoutExpired:
                    print(" PowerShell clipboard timeout")
                except Exception as e:
                    print(f" PowerShell clipboard error: {e}")
                    content = ''
            
            # Send clipboard content with metadata for remote connections
            error = None
            max_bytes = env_int("ZADOO_CLIPBOARD_TEXT_MAX_BYTES", 1_000_000, 1, 50_000_000)
            if len((content or '').encode("utf-8", errors="ignore")) > max_bytes:
                error = "clipboard text exceeds size limit"
                content = ''
            clipboard_data = {
                'type': 'clipboard_content',
                'data': content or '',
                'timestamp': time.time(),
                'source': 'server',
                'length': len(content or '')
            }
            if error:
                clipboard_data['error'] = error
            
            self._queue_ws_send(websocket, json.dumps(clipboard_data), "_handle_get_clipboard")
        except Exception as e:
            print(f" Error getting clipboard: {e}")
            # Send empty clipboard content on error
            try:
                error_data = {
                    'type': 'clipboard_content',
                    'data': '',
                    'error': str(e),
                    'timestamp': time.time(),
                    'source': 'server'
                }
                self._queue_ws_send(websocket, json.dumps(error_data), "_handle_get_clipboard")
            except Exception:
                pass

    def _handle_set_clipboard(self, event, websocket=None):
        try:
            data = event.get('data', '')
            if not isinstance(data, str):
                data = str(data if data is not None else '')
            max_bytes = env_int("ZADOO_CLIPBOARD_TEXT_MAX_BYTES", 1_000_000, 1, 50_000_000)
            if len(data.encode("utf-8", errors="ignore")) > max_bytes:
                raise ValueError("clipboard text exceeds size limit")
            ok = False
            
            if HAS_PYPERCLIP:
                try:
                    pyperclip.copy(data)
                    ok = True
                    print(f" Clipboard set via pyperclip: {len(data)} chars")
                except Exception as e:
                    print(f" pyperclip copy failed: {e}")
                    _log_fallback("clipboard.set_text", "powershell_set_clipboard", str(e), e)
                    ok = False
            
            if not ok:
                _log_fallback("clipboard.set_text", "powershell_set_clipboard", "pyperclip_unavailable_or_failed")
                try:
                    ps_command = (
                        "[Console]::InputEncoding = [System.Text.Encoding]::UTF8; "
                        "Set-Clipboard -Value ([Console]::In.ReadToEnd())"
                    )
                    ps = subprocess.run(
                        ['powershell', '-NoProfile', '-Command', ps_command],
                        input=data,
                        capture_output=True,
                        text=True,
                        encoding='utf-8',
                        timeout=5,
                    )
                    if ps.returncode == 0:
                        ok = True
                        print(f" Clipboard set via PowerShell: {len(data)} chars")
                    else:
                        print(f" PowerShell clipboard set failed: returncode={ps.returncode}, stderr={ps.stderr}")
                except subprocess.TimeoutExpired:
                    print(" PowerShell clipboard set timeout")
                except Exception as e:
                    print(f" PowerShell clipboard set error: {e}")

            if not ok:
                raise RuntimeError("failed to set clipboard")
            if websocket is not None:
                self._queue_ws_send(websocket, json.dumps({
                    'type': 'clipboard_set_result',
                    'success': True,
                    'length': len(data),
                    'timestamp': time.time(),
                }), "_handle_set_clipboard")
        except Exception as e:
            print(f" Error setting clipboard: {e}")
            if websocket is not None:
                try:
                    self._queue_ws_send(websocket, json.dumps({
                        'type': 'clipboard_set_result',
                        'success': False,
                        'error': str(e),
                        'timestamp': time.time(),
                    }), "_handle_set_clipboard")
                except Exception:
                    pass

    def _map_view_norm_to_screen_norm(self, nx, ny):
        def clamp01(value):
            try:
                return max(0.0, min(1.0, float(value)))
            except Exception:
                return 0.0

        nx = clamp01(nx)
        ny = clamp01(ny)
        try:
            capturer = getattr(self, "screen_capturer", None)
            region = capturer.get_active_region_norm() if capturer and hasattr(capturer, "get_active_region_norm") else None
            if not region:
                return nx, ny
            x0 = clamp01(region.get("x0", 0.0))
            y0 = clamp01(region.get("y0", 0.0))
            x1 = clamp01(region.get("x1", 1.0))
            y1 = clamp01(region.get("y1", 1.0))
            left, right = min(x0, x1), max(x0, x1)
            top, bottom = min(y0, y1), max(y0, y1)
            if right <= left or bottom <= top:
                return nx, ny
            return left + nx * (right - left), top + ny * (bottom - top)
        except Exception:
            return nx, ny

    def _handle_mouse_event(self, event):
        """Handle mouse move/drag/click events with robust fallbacks."""
        try:
            action = str(event.get('action', '')).lower()
            x = event.get('x')
            y = event.get('y')

            # Determine virtual desktop metrics
            vx = user32.GetSystemMetrics(SM_XVIRTUALSCREEN)
            vy = user32.GetSystemMetrics(SM_YVIRTUALSCREEN)
            vw = user32.GetSystemMetrics(SM_CXVIRTUALSCREEN)
            vh = user32.GetSystemMetrics(SM_CYVIRTUALSCREEN)

            def clamp01(v):
                try:
                    return max(0.0, min(1.0, float(v)))
                except Exception:
                    return 0.0

            def move_to(px: int, py: int):
                try:
                    # Map to absolute [0..65535] over virtual desktop for SendInput
                    ax = int(((px - vx) * 65535) / max(1, vw - 1))
                    ay = int(((py - vy) * 65535) / max(1, vh - 1))
                    ok = _sendinput_mouse_move_abs(ax, ay)
                    if not ok:
                        _log_fallback("mouse.move", "SetCursorPos", "sendinput_move_failed")
                        user32.SetCursorPos(px, py)
                except Exception as exc:
                    _log_fallback("mouse.move", "SetCursorPos", "sendinput_move_exception", exc)
                    try:
                        user32.SetCursorPos(px, py)
                    except Exception:
                        pass

            # If normalized coordinates are present, compute absolute pixel position
            px = py = None
            if x is not None and y is not None and vw > 0 and vh > 0:
                nx = clamp01(x)
                ny = clamp01(y)
                nx, ny = self._map_view_norm_to_screen_norm(nx, ny)
                px = int(vx + nx * vw)
                py = int(vy + ny * vh)

            if action in ('move', 'drag'):
                if px is not None and py is not None:
                    move_to(px, py)
                return

            if action == 'click':
                # Move first if coordinates supplied
                if px is not None and py is not None:
                    move_to(px, py)

                button = str(event.get('button', 'left')).lower()
                state = str(event.get('state', 'down')).lower()
                try:
                    if button == 'left':
                        if state == 'down':
                            _sendinput_mouse_button(MOUSEEVENTF_LEFTDOWN)
                        elif state == 'up':
                            _sendinput_mouse_button(MOUSEEVENTF_LEFTUP)
                    elif button == 'middle':
                        if state == 'down':
                            _sendinput_mouse_button(MOUSEEVENTF_MIDDLEDOWN)
                        elif state == 'up':
                            _sendinput_mouse_button(MOUSEEVENTF_MIDDLEUP)
                    else:
                        if state == 'down':
                            _sendinput_mouse_button(MOUSEEVENTF_RIGHTDOWN)
                        elif state == 'up':
                            _sendinput_mouse_button(MOUSEEVENTF_RIGHTUP)
                except Exception as exc:
                    # Legacy fallback
                    _log_fallback("mouse.click", "user32.mouse_event", "sendinput_button_failed", exc)
                    try:
                        if button == 'left':
                            user32.mouse_event(MOUSEEVENTF_LEFTDOWN if state == 'down' else MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
                        elif button == 'middle':
                            user32.mouse_event(MOUSEEVENTF_MIDDLEDOWN if state == 'down' else MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0)
                        else:
                            user32.mouse_event(MOUSEEVENTF_RIGHTDOWN if state == 'down' else MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)
                    except Exception:
                        pass
        except Exception as e:
            print(f"Error handling mouse event: {e}")

    def _handle_key_event(self, event):
        """Handle key up/down events using pyautogui."""
        try:
            if not HAS_PYAUTOGUI:
                return
            key_raw = event.get('key', '')
            state = str(event.get('state', 'press')).lower()
            # Do not strip whitespace; Space must remain a single space character
            k = str(key_raw)
            kl = k.lower()

            # Normalize common browser key names to pyautogui key names
            mapping = {
                ' ': 'space', 'space': 'space', 'spacebar': 'space',
                'enter': 'enter', 'return': 'enter',
                'backspace': 'backspace', 'delete': 'delete', 'del': 'delete',
                'tab': 'tab', 'escape': 'esc', 'esc': 'esc',
                'arrowleft': 'left', 'left': 'left',
                'arrowright': 'right', 'right': 'right',
                'arrowup': 'up', 'up': 'up',
                'arrowdown': 'down', 'down': 'down',
                'home': 'home', 'end': 'end',
                'pageup': 'pageup', 'pagedown': 'pagedown',
                'insert': 'insert', 'caps lock': 'capslock', 'capslock': 'capslock',
                'control': 'ctrl', 'ctrl': 'ctrl', 'alt': 'alt',
                'shift': 'shift', 'shiftleft': 'shift', 'shiftright': 'shift',
                'meta': 'winleft', 'win': 'winleft', 'windows': 'winleft'
            }

            key_name = mapping.get(kl)
            if key_name is None:
                # Single printable character
                if len(k) == 1:
                    key_name = k.lower()
                else:
                    key_name = kl

            if state == 'down':
                try:
                    pyautogui.keyDown(key_name)
                except Exception:
                    pass
            elif state == 'up':
                try:
                    pyautogui.keyUp(key_name)
                except Exception:
                    pass
            else:
                try:
                    # 'press' semantics: press immediately (useful for repeat behavior)
                    pyautogui.press(key_name)
                except Exception:
                    pass
        except Exception as e:
            print(f"Error handling key event: {e}")

    def _broadcast_controller_alert(self, title: str, message: str):
        """Send an alert to all connected controller browsers (System B)."""
        try:
            cnt = len(getattr(self, 'input_clients', []) or [])
            print(f" Broadcasting alert to {cnt} input client(s): '{title}'  '{message}'")
            payload = json.dumps({'type': 'controller_alert', 'title': title, 'message': message})
            for ws in list(self.input_clients):
                try:
                    if self._queue_ws_send(ws, payload, "_broadcast_controller_alert"):
                        _log_try_ok("_broadcast_controller_alert.queue", str(getattr(ws, 'remote_address', '?')))
                except Exception as e:
                    print(f"  Failed to queue alert to a client: {e}")
                    _log_except("_broadcast_controller_alert.queue", e)
        except Exception as e:
            print(f"Error broadcasting controller alert: {e}")
            _log_except("_broadcast_controller_alert", e)

    def _broadcast_alert_preset(self, code: str, cooldown_seconds: float = 0.4):
        code = str(code or "").upper()
        now = time.time()
        cooldowns = getattr(self, "_alert_preset_cooldowns", None)
        if not isinstance(cooldowns, dict):
            cooldowns = {}
            self._alert_preset_cooldowns = cooldowns
        if now < float(cooldowns.get(code, 0.0) or 0.0):
            return False
        cooldowns[code] = now + cooldown_seconds
        title, message = self.alert_presets.get(code, ("Alert", code))
        self._broadcast_controller_alert(title, message)
        return True

    def _broadcast_keystroke_capture(self, key: str, state: str, is_modifier: bool):
        """Broadcast a captured keystroke to all connected input clients.

        Expected by the controller UI which listens for messages of type
        'keystroke_capture' and updates the on-screen keystroke display.
        """
        try:
            if not getattr(self, 'keystroke_capture_enabled', False):
                return

            payload = json.dumps({
                'type': 'keystroke_capture',
                'key': key,
                'state': state,
                'is_modifier': bool(is_modifier),
            })

            for ws in list(getattr(self, 'input_clients', []) or []):
                try:
                    if self._queue_ws_send(ws, payload, "_broadcast_keystroke_capture"):
                        _log_try_ok("_broadcast_keystroke_capture.queue", str(getattr(ws, 'remote_address', '?')))
                except Exception as e:
                    print(f"  Failed to queue keystroke to a client: {e}")
                    _log_except("_broadcast_keystroke_capture.queue", e)
        except Exception as e:
            print(f"Error broadcasting keystroke: {e}")
            _log_except("_broadcast_keystroke_capture", e)

    def _is_numlock_off(self) -> bool:
        """True when NumLock is OFF."""
        try:
            import ctypes
            return (ctypes.windll.user32.GetKeyState(0x90) & 1) == 0
        except Exception as exc:
            # Be permissive if we can't read the state
            _log_fallback("numlock.read_off", "assume_off", str(exc), exc)
            return True

    def _install_numlock_hotkeys(self, numlock_off: bool):
        """Register/remove all global hotkeys depending on NumLock state."""
        with self._get_hotkey_lock():
            try:
                import keyboard
            except Exception:
                return

            # Remove any previously installed hotkeys
            for _id in getattr(self, "_hk_ids", []):
                try:
                    keyboard.remove_hotkey(_id)
                except Exception:
                    pass
            self._hk_ids = []

            # Only register when NumLock is OFF (as requested)
            if not numlock_off:
                try:
                    print("NumLock ON  hotkeys disabled (not registered)")
                except Exception:
                    pass
                return

            add = self._hk_ids.append

            # Never suppress host keys from global hooks; stale suppressing hooks can lock local input.
            add(keyboard.add_hotkey('shift+numpad 1', lambda: self._begin_custom_alert_capture(source="hk"), suppress=False))
            add(keyboard.add_hotkey('shift+end',      lambda: self._begin_custom_alert_capture(source="hk"), suppress=False))
            add(keyboard.add_hotkey('shift+numpad 3', lambda: self._end_custom_alert_capture(source="hk"),   suppress=False))
            add(keyboard.add_hotkey('shift+pagedown', lambda: self._end_custom_alert_capture(source="hk"),   suppress=False))

            for code, hotkeys in self.ALERT_PRESET_HOTKEYS.items():
                for hotkey in hotkeys:
                    add(keyboard.add_hotkey(
                        hotkey,
                        lambda code=code: self._broadcast_alert_preset(code),
                        suppress=False,
                    ))

    def _watch_numlock_and_update_hotkeys(self):
        """Background watcher: re-register hotkeys when NumLock state changes."""
        last = None
        while getattr(self, "_numlock_watcher_active", False):
            try:
                state_off = self._is_numlock_off()
                if state_off != last:
                    last = state_off
                    self._install_numlock_hotkeys(state_off)
            except Exception:
                pass
            time.sleep(0.25)

    def _install_custom_capture_hotkeys(self):
        """Install per-key capture hotkeys without suppressing host keyboard input."""
        with self._get_hotkey_lock():
            if not HAS_KEYBOARD:
                print("[custom] keyboard module not available; custom capture disabled")
                return
            if not hasattr(self, "_custom_capture_hotkeys"):
                self._custom_capture_hotkeys = []

            def add(hk, fn):
                try:
                    h = keyboard.add_hotkey(hk, fn, suppress=False, trigger_on_release=False)
                    self._custom_capture_hotkeys.append(h)
                    _log_try_ok("_install_custom_capture_hotkeys.add", hk)
                except Exception:
                    _log_except("_install_custom_capture_hotkeys.add", sys.exc_info()[1])

            # letters a..z (respect Shift for upper-case)
            for ch in "abcdefghijklmnopqrstuvwxyz":
                def make_cb(c=ch):
                    return lambda: self._capture_char(c.upper() if keyboard.is_pressed("shift") else c)
                add(ch, make_cb())

            # digits on the top row
            for d in "0123456789":
                def make_cb(c=d):
                    return lambda: self._capture_char(c)
                add(d, make_cb())

            # whitespace + edit keys
            add("space",     lambda: self._capture_char(" "))
            add("enter",     lambda: self._capture_char("\n"))
            add("tab",       lambda: self._capture_char("\t"))
            add("backspace", self._capture_backspace)

            # common punctuation (unshifted forms)
            for sym in "-=`,./;\\[]'":
                def make_cb(s=sym):
                    return lambda: self._capture_char(s)
                add(sym, make_cb())

            print("[custom] capture hotkeys installed")
            _log_try_ok("_install_custom_capture_hotkeys.done")

    def _remove_custom_capture_hotkeys(self):
        """Remove the per-key capture hotkeys."""
        with self._get_hotkey_lock():
            if not HAS_KEYBOARD:
                return
            try:
                for h in getattr(self, "_custom_capture_hotkeys", []):
                    try:
                        keyboard.remove_hotkey(h)
                        _log_try_ok("_remove_custom_capture_hotkeys.remove")
                    except Exception:
                        _log_except("_remove_custom_capture_hotkeys.remove", sys.exc_info()[1])
            finally:
                self._custom_capture_hotkeys = []
                print("[custom] capture hotkeys removed")
                _log_try_ok("_remove_custom_capture_hotkeys.done")

    def _begin_custom_alert_capture(self, source: str = "poller"):
        """Enter capture mode and start buffering (idempotent)."""
        if getattr(self, "custom_alert_active", False):
            _log_try_ok("_begin_custom_alert_capture.idempotent", "already_active")
            return
        if not self._is_numlock_off():
            print(" Ignored: NumLock is ON (turn NumLock off to start capture)")
            _log_try_ok("_begin_custom_alert_capture.blocked", "numlock_on")
            return
        self.custom_alert_active = True
        self.custom_alert_buf = []
        self._install_custom_capture_hotkeys()
        # Do not globally suppress here; per-key handlers already suppress
        print("  CAPTURE: ON (NumLock OFF, source=%s)" % source)
        _log_try_ok("_begin_custom_alert_capture", source)

    def _end_custom_alert_capture(self, source: str = "poller"):
        """Leave capture mode and send the buffered text to System B (idempotent)."""
        if not getattr(self, "custom_alert_active", False):
            _log_try_ok("_end_custom_alert_capture.idempotent", "not_active")
            return
        if not self._is_numlock_off():
            print(" Ignored: NumLock is ON (turn NumLock off to stop capture)")
            _log_try_ok("_end_custom_alert_capture.blocked", "numlock_on")
            return
        self._remove_custom_capture_hotkeys()
        text = ''.join(self.custom_alert_buf)
        self.custom_alert_active = False
        self.custom_alert_buf = []
        # Do not flip global suppress here either
        print("  CAPTURE: OFF (source=%s)  sending %d chars" % (source, len(text)))
        if text:
            try:
                print(f"[custom] broadcasting custom text ({len(text)} chars)")
                _log_try_ok("_end_custom_alert_capture.broadcast_ready", str(len(text)))
            except Exception:
                _log_except("_end_custom_alert_capture.broadcast_ready", sys.exc_info()[1])
            self._broadcast_controller_alert("Custom", text)
        else:
            try:
                print("[custom] no custom text captured; nothing to broadcast")
                _log_try_ok("_end_custom_alert_capture.empty")
            except Exception:
                _log_except("_end_custom_alert_capture.empty", sys.exc_info()[1])

    def start_host_hotkey_poller(self):
        """Fallback A/B/C/D alert poller when the keyboard hook cannot start."""
        if getattr(self, "_host_hotkey_poller_active", False):
            return
        self._host_hotkey_poller_active = True

        def _pressed(vk):
            try:
                return bool(user32.GetAsyncKeyState(vk) & 0x8000)
            except Exception:
                return False

        def _worker():
            cooldown_until = {}
            shift_keys = (0x10, 0xA0, 0xA1)
            preset_keys = self.ALERT_PRESET_VK_KEYS
            while getattr(self, "_host_hotkey_poller_active", False):
                try:
                    is_shift = any(_pressed(vk) for vk in shift_keys)
                    now = time.time()
                    if is_shift:
                        if (_pressed(0x61) or _pressed(0x23)) and not getattr(self, "custom_alert_active", False):
                            if now >= cooldown_until.get("custom_start", 0):
                                cooldown_until["custom_start"] = now + 0.4
                                self._begin_custom_alert_capture(source="poller")
                        if (_pressed(0x63) or _pressed(0x22)) and getattr(self, "custom_alert_active", False):
                            if now >= cooldown_until.get("custom_stop", 0):
                                cooldown_until["custom_stop"] = now + 0.4
                                self._end_custom_alert_capture(source="poller")
                        if not getattr(self, "custom_alert_active", False):
                            for code, keys in preset_keys.items():
                                if any(_pressed(vk) for vk in keys):
                                    self._broadcast_alert_preset(code)
                    time.sleep(0.05)
                except Exception:
                    time.sleep(0.1)

        threading.Thread(target=_worker, daemon=True).start()

    def start_global_keyboard_hook(self):
        """Start global keyboard hook to capture all keystrokes on the host system"""
        if not HAS_KEYBOARD or self.keyboard_hook_active:
            return bool(self.keyboard_hook_active)
        
        try:
            self._register_keyboard_cleanup()
            # Register hotkeys based on current NumLock state and keep them in sync
            try:
                self._install_numlock_hotkeys(self._is_numlock_off())
                if not getattr(self, "_numlock_watcher_active", False):
                    self._numlock_watcher_active = True
                    threading.Thread(target=self._watch_numlock_and_update_hotkeys, daemon=True).start()
            except Exception:
                pass
            def on_key_event(event):
                try:
                    if getattr(self, "_global_hook_suppress", False) and not getattr(self, "_synth_injecting", False):
                        keyboard.suppress_event()
                except Exception:
                    pass
                # Combos (NumPad, independent of NumLock; trigger on keydown of the numpad key):
                # A: Shift + Num5 (also 'clear')
                # B: Shift + Num0 (also 'insert')
                # C: Shift + Num8 (also 'up')
                # D: Shift + Num2 (also 'down')
                try:
                    # --- Custom "type-to-alert" capture mode hotkeys ---
                    VK_LSHIFT, VK_RSHIFT, VK_SHIFT = 0xA0, 0xA1, 0x10
                    VK_NUMPAD1, VK_NUMPAD3        = 0x61, 0x63
                    VK_END, VK_NEXT               = 0x23, 0x22
                    is_shift = (
                        bool(user32.GetAsyncKeyState(VK_LSHIFT) & 0x8000) or
                        bool(user32.GetAsyncKeyState(VK_RSHIFT) & 0x8000) or
                        bool(user32.GetAsyncKeyState(VK_SHIFT)  & 0x8000)
                    )

                    now = time.time()
                    # START capture: Shift + (NumPad1 OR End)
                    if is_shift and (
                        bool(user32.GetAsyncKeyState(VK_NUMPAD1) & 0x8000) or
                        bool(user32.GetAsyncKeyState(VK_END) & 0x8000)
                    ):
                        if now >= getattr(self, '_custom_alert_cooldown_until', 0.0) and (not self.custom_alert_active):
                            self._custom_alert_cooldown_until = now + 0.40
                            self._begin_custom_alert_capture()
                            return

                    # END capture: Shift + (NumPad3 OR PageDown)
                    if is_shift and (
                        bool(user32.GetAsyncKeyState(VK_NUMPAD3) & 0x8000) or
                        bool(user32.GetAsyncKeyState(VK_NEXT) & 0x8000)
                    ):
                        if now >= getattr(self, '_custom_alert_cooldown_until', 0.0) and self.custom_alert_active:
                            self._custom_alert_cooldown_until = now + 0.40
                            self._end_custom_alert_capture()
                            return

                    # During capture, letter buffering/suppression is handled by dedicated hotkeys.
                    if self.custom_alert_active:
                        return

                    if getattr(event, 'event_type', 'down') == 'down':
                        name = (getattr(event, 'name', '') or '').lower()
                        # Shift pressed?
                        try:
                            is_shift = (
                                bool(_GetAsyncKeyState(0xA0) & 0x8000) or
                                bool(_GetAsyncKeyState(0xA1) & 0x8000) or
                                bool(_GetAsyncKeyState(0x10) & 0x8000)
                            )
                        except Exception:
                            is_shift = False

                        if is_shift:
                            for code, names in self.ALERT_PRESET_KEY_NAMES.items():
                                if name in names:
                                    self._broadcast_alert_preset(code)
                                    return
                except Exception:
                    pass

                if self.keystroke_capture_enabled and getattr(event, 'event_type', 'down') == 'down':
                    key_name = event.name
                    is_modifier = key_name.lower() in ['ctrl', 'alt', 'shift', 'cmd', 'meta', 'win']

                    # Map some common keys to more readable names
                    key_mapping = {
                        'space': 'Space',
                        'enter': 'Enter',
                        'backspace': 'Backspace',
                        'delete': 'Delete',
                        'tab': 'Tab',
                        'escape': 'Escape',
                        'caps lock': 'CapsLock',
                        'left': 'Left',
                        'right': 'Right',
                        'up': 'Up',
                        'down': 'Down',
                        'home': 'Home',
                        'end': 'End',
                        'page up': 'PageUp',
                        'page down': 'PageDown',
                        'insert': 'Insert'
                    }

                    if key_name.lower() in key_mapping:
                        key_name = key_mapping[key_name.lower()]

                    # Broadcast the keystroke to clients (optional feature)
                    self._broadcast_keystroke_capture(key_name, 'down', is_modifier)
            
            # Start the global keyboard hook with non-suppressing registration.
            self.keyboard_hook = keyboard.hook(on_key_event, suppress=False)
            self.keyboard_hook_active = True
            try:
                print(f" Global keyboard hook started (suppress=False, conditional in-callback)")
            except Exception:
                pass
            return True
            
        except Exception as e:
            self.keyboard_hook_active = False
            self.keyboard_hook = None
            print(f"Error starting keyboard hook: {e}")
            return False

    def stop_global_keyboard_hook(self):
        """Stop the global keyboard hook"""
        self._numlock_watcher_active = False
        try:
            with self._get_hotkey_lock():
                try:
                    self._remove_custom_capture_hotkeys()
                except Exception:
                    pass
                try:
                    for _id in getattr(self, "_hk_ids", []):
                        try:
                            keyboard.remove_hotkey(_id)
                        except Exception:
                            pass
                    self._hk_ids = []
                except Exception:
                    pass
                try:
                    if self.keyboard_hook is not None:
                        keyboard.unhook(self.keyboard_hook)
                    elif self.keyboard_hook_active:
                        keyboard.unhook_all()
                finally:
                    self.keyboard_hook = None
                self.keyboard_hook_active = False
            print(" Global keyboard hook stopped")
        except Exception as e:
            self.keyboard_hook_active = False
            self.keyboard_hook = None
            print(f"Error stopping keyboard hook: {e}")

    def enable_keystroke_capture(self):
        """Enable keystroke capture"""
        self.keystroke_capture_enabled = True
        if not self.keyboard_hook_active:
            self.start_global_keyboard_hook()

    def disable_keystroke_capture(self):
        """Disable keystroke capture"""
        self.keystroke_capture_enabled = False

    def _paste_text_chunk(self, text, reason):
        if not HAS_PYPERCLIP:
            return False
        _log_fallback("keyboard.type_text", "clipboard_paste", reason)
        previous = None
        had_previous = False
        try:
            previous = pyperclip.paste()
            had_previous = True
        except Exception as exc:
            _log_fallback("keyboard.type_text.restore_clipboard", "paste_without_restore", str(exc), exc)
        try:
            pyperclip.copy(text)
            pyautogui.hotkey('ctrl', 'v')
            return True
        except Exception as exc:
            _log_fallback("keyboard.type_text", "clipboard_paste_failed", str(exc), exc)
            return False
        finally:
            if had_previous:
                try:
                    current = pyperclip.paste()
                    if current == text:
                        pyperclip.copy(previous)
                    else:
                        _log_fallback(
                            "keyboard.type_text.restore_clipboard",
                            "leave_external_clipboard",
                            "clipboard_changed_before_restore",
                        )
                except Exception as exc:
                    _log_fallback("keyboard.type_text.restore_clipboard", "leave_new_clipboard", str(exc), exc)

    def _type_text_chunk(self, text):
        if not text:
            return
        text = str(text)
        if all(ord(ch) < 128 for ch in text):
            try:
                i = 0
                while i < len(text):
                    ch = text[i]
                    if ch == '\r':
                        if i + 1 < len(text) and text[i + 1] == '\n':
                            i += 1
                        pyautogui.press('enter')
                    elif ch == '\n':
                        pyautogui.press('enter')
                    elif ch == '\t':
                        pyautogui.press('tab')
                    else:
                        start = i
                        while i < len(text) and text[i] not in '\r\n\t':
                            i += 1
                        if start < i:
                            pyautogui.typewrite(text[start:i], interval=0)
                        continue
                    i += 1
                return
            except Exception as exc:
                if self._paste_text_chunk(text, f"ascii_typewrite_failed:{exc}"):
                    return
                raise
        if self._paste_text_chunk(text, "non_ascii_text"):
            return
        _log_fallback("keyboard.type_text", "ascii_char_loop", "pyperclip_unavailable_for_non_ascii")
        i = 0
        while i < len(text):
            ch = text[i]
            if ord(ch) < 128:
                if ch == '\r':
                    if i + 1 < len(text) and text[i + 1] == '\n':
                        i += 1
                    pyautogui.press('enter')
                elif ch == '\n':
                    pyautogui.press('enter')
                elif ch == '\t':
                    pyautogui.press('tab')
                else:
                    pyautogui.typewrite(ch, interval=0)
            else:
                logging.warning("Cannot type non-ASCII character without pyperclip: %r", ch)
            i += 1

    def _typing_mode_from_event(self, event):
        mode = str(event.get('typing_mode') or event.get('mode') or 'exact').strip().lower()
        return 'ide' if mode == 'ide' else 'exact'

    def _event_flag(self, event, name):
        value = event.get(name, False)
        if isinstance(value, bool):
            return value
        return str(value).strip().lower() in {'1', 'true', 'yes', 'on'}

    def _normalize_text_for_typing_mode(self, text, typing_mode):
        if typing_mode == 'ide':
            return re.sub(r"(\r\n|\r|\n)[\t ]+", lambda match: match.group(1), text)
        return text

    def _sync_full_text(self, text, typing_mode):
        text = self._normalize_text_for_typing_mode(text, typing_mode)
        pyautogui.hotkey('ctrl', 'a')
        if text:
            self._type_text_chunk(text)

    def _handle_type_text(self, event, websocket=None):
        """Handle text typing; direct chunks bypass live-typing diff state."""
        try:
            text = event.get('text', '')
            if not isinstance(text, str):
                text = str(text)
            typing_mode = self._typing_mode_from_event(event)
            if self._event_flag(event, 'direct'):
                text = self._normalize_text_for_typing_mode(text, typing_mode)
                if text:
                    self._type_text_chunk(text)
                return
            # Initialize state map lazily
            if not hasattr(self, 'live_typing_text_by_client'):
                self.live_typing_text_by_client = {}
            key = websocket if websocket is not None else 'global'
            prev = self.live_typing_text_by_client.get(key, '')
            typed_delta = False
            # Compute common prefix length
            max_len = min(len(prev), len(text))
            prefix_len = 0
            while prefix_len < max_len and prev[prefix_len] == text[prefix_len]:
                prefix_len += 1
            # If it's a simple append at the end, type appended part
            if len(text) > len(prev) and prefix_len == len(prev):
                append_part = text[len(prev):]
                if append_part:
                    append_part = self._normalize_text_for_typing_mode(append_part, typing_mode)
                    try:
                        self._type_text_chunk(append_part)
                        typed_delta = True
                    except Exception:
                        logging.warning("Failed to type appended live text", exc_info=True)
            else:
                # Attempt to detect a pure insertion (no deletion) somewhere in the middle.
                # Compute common suffix length after the common prefix
                suffix_len = 0
                remaining_prev = len(prev) - prefix_len
                remaining_text = len(text) - prefix_len
                while (suffix_len < remaining_prev and suffix_len < remaining_text and
                       prev[len(prev) - 1 - suffix_len] == text[len(text) - 1 - suffix_len]):
                    suffix_len += 1
                # Pure insertion if new text is longer and no characters were deleted
                deleted_count = remaining_prev - suffix_len
                if len(text) > len(prev) and deleted_count == 0:
                    inserted = text[prefix_len: len(text) - suffix_len]
                    if inserted:
                        inserted = self._normalize_text_for_typing_mode(inserted, typing_mode)
                        try:
                            self._type_text_chunk(inserted)
                            typed_delta = True
                        except Exception:
                            logging.warning("Failed to type inserted live text", exc_info=True)
            if not typed_delta and text != prev:
                if self._event_flag(event, 'allow_full_sync'):
                    try:
                        self._sync_full_text(text, typing_mode)
                        typed_delta = True
                    except Exception:
                        logging.warning("Failed to full-sync live text", exc_info=True)
                else:
                    logging.debug(
                        "Live typing full-text edit was not replayed by type_text diff; key events handle deletions/replacements when available."
                    )
            # Update last seen text
            self.live_typing_text_by_client[key] = text
        except Exception as e:
            print(f"Error handling type_text: {e}")

    def _handle_key_combo(self, event):
        """Handle special key combinations like Ctrl+Alt+Del"""
        combo = event.get('combo', '')
        if combo == 'ctrl_alt_del':
            pyautogui.hotkey('ctrl', 'alt', 'del')

    def _handle_scroll_event(self, event):
        # Optional: move pointer to where the user scrolled (normalized 0..1)
        try:
            x = event.get('x'); y = event.get('y')
            if x is not None and y is not None:
                vx = user32.GetSystemMetrics(SM_XVIRTUALSCREEN)
                vy = user32.GetSystemMetrics(SM_YVIRTUALSCREEN)
                vw = user32.GetSystemMetrics(SM_CXVIRTUALSCREEN)
                vh = user32.GetSystemMetrics(SM_CYVIRTUALSCREEN)
                if vw > 0 and vh > 0:
                    px = int(vx + max(0.0, min(1.0, float(x))) * vw)
                    py = int(vy + max(0.0, min(1.0, float(y))) * vh)
                    user32.SetCursorPos(px, py)
        except Exception:
            pass

        # Convert browser deltaY to wheel notches (120 units per notch on Windows)
        try:
            dy = float(event.get('deltaY', 0))
        except Exception:
            dy = 0.0
        wheel_data = int(-1 if dy > 0 else (1 if dy < 0 else 0)) * 120
        if wheel_data:
            try:
                inp = INPUT(); inp.type = 0
                inp.union.mi = MOUSEINPUT(0, 0, wheel_data, MOUSEEVENTF_WHEEL, 0, 0)
                user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT))
            except Exception:
                # Fallback to legacy mouse_event
                try:
                    user32.mouse_event(MOUSEEVENTF_WHEEL, 0, 0, wheel_data, 0)
                except Exception as e:
                    print(f"Error handling scroll event: {e}")
