"""Remote input, clipboard, alerts, and hotkey controls."""
from __future__ import annotations

import asyncio
import atexit
import base64
import ctypes
import io
import json
import logging
import math
import re
import time

import keyboard
import websockets
import win32clipboard
import win32con
from PIL import Image

from .logging_utils import _log_except
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
    _sendinput_key,
    _sendinput_mouse_button,
    _sendinput_mouse_move_abs,
    _sendinput_unicode,
    user32,
)


def _unit_float(value, name):
    if isinstance(value, bool):
        raise ValueError(f"{name} must be a number from 0 to 1")
    try:
        result = float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{name} must be a number from 0 to 1") from exc
    if not math.isfinite(result) or not 0.0 <= result <= 1.0:
        raise ValueError(f"{name} must be a number from 0 to 1")
    return result


class InputControlMixin:
    INPUT_ACTIONS = {
        'click', 'move', 'drag', 'key', 'scroll', 'type_text',
        'get_clipboard', 'set_clipboard', 'set_clipboard_image',
    }
    ALERT_PRESET_HOTKEYS = {
        'A': ('shift+num 5', 'shift+clear'),
        'B': ('shift+num 0', 'shift+insert'),
        'C': ('shift+num 8', 'shift+up'),
        'D': ('shift+num 2', 'shift+down'),
    }
    KEY_NAMES = {
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
        'meta': 'winleft', 'win': 'winleft', 'windows': 'winleft',
    }
    MOUSE_BUTTON_FLAGS = {
        ('left', 'down'): MOUSEEVENTF_LEFTDOWN,
        ('left', 'up'): MOUSEEVENTF_LEFTUP,
        ('middle', 'down'): MOUSEEVENTF_MIDDLEDOWN,
        ('middle', 'up'): MOUSEEVENTF_MIDDLEUP,
        ('right', 'down'): MOUSEEVENTF_RIGHTDOWN,
        ('right', 'up'): MOUSEEVENTF_RIGHTUP,
    }

    def _register_keyboard_cleanup(self):
        if self._keyboard_cleanup_registered:
            return
        atexit.register(self.stop_host_hotkeys)
        self._keyboard_cleanup_registered = True

    def _queue_ws_send(self, websocket, payload, log_name):
        def _done(completed):
            self._background_send_tasks.discard(completed)
            if completed.cancelled():
                return
            try:
                completed.result()
            except websockets.exceptions.ConnectionClosed:
                return
            except Exception as exc:
                _log_except(f"{log_name}.send", exc)

        try:
            loop = asyncio.get_running_loop()
            task = loop.create_task(websocket.send(payload))
            self._background_send_tasks.add(task)
            task.add_done_callback(_done)
            return True
        except RuntimeError:
            if not self.loop:
                return False
            try:
                future = asyncio.run_coroutine_threadsafe(websocket.send(payload), self.loop)
                future.add_done_callback(_done)
                return True
            except Exception as exc:
                _log_except(f"{log_name}.queue", exc)
                return False
        except Exception as exc:
            _log_except(f"{log_name}.queue", exc)
            return False

    def _capture_char(self, ch: str):
        if self.custom_alert_active:
            self.custom_alert_buf.append(ch)

    def _capture_backspace(self):
        if self.custom_alert_active and self.custom_alert_buf:
            self.custom_alert_buf.pop()

    async def input_event_handler(self, websocket):
        print(f" New input client connected from {websocket.remote_address}")
        self.input_clients.add(websocket)
        
        try:
            async for message in websocket:
                action = None
                try:
                    event = json.loads(message)
                    if not isinstance(event, dict):
                        raise ValueError("Input event must be a JSON object")
                    action = event.get('action')
                    if not self._is_ws_action_authorized(websocket, action):
                        await self._send_ws_forbidden(websocket, action)
                        continue

                    if action in self.INPUT_ACTIONS:
                        if action in {"set_clipboard_image", "type_text"}:
                            await asyncio.to_thread(self.process_event, event, websocket)
                        else:
                            self.process_event(event, websocket)
                    elif action == 'refresh_tunnel':
                        await self.handle_refresh_via_websocket(websocket)
                    elif action == 'get_public_url':
                        await self.handle_get_url_via_websocket(websocket)
                    elif action == 'cursor_broadcast':
                        enabled = event['enabled']
                        if not isinstance(enabled, bool):
                            raise ValueError("enabled must be a boolean")
                        if enabled:
                            self.cursor_subscribers.add(websocket)
                            self.cursor_broadcast_enabled = True
                            print(f" Cursor broadcast enabled for {websocket.remote_address}")
                        else:
                            self.cursor_subscribers.discard(websocket)
                            if not self.cursor_subscribers:
                                self.cursor_broadcast_enabled = False
                            print(f" Cursor broadcast disabled for {websocket.remote_address}")

                except Exception as e:
                    await websocket.send(json.dumps({
                        "type": "input_error",
                        "action": action,
                        "error": str(e),
                    }))
        
        except websockets.exceptions.ConnectionClosed:
            print(f" Input client {websocket.remote_address} disconnected")
        finally:
            self.input_clients.discard(websocket)
            self.cursor_subscribers.discard(websocket)
            self.live_typing_text_by_client.pop(websocket, None)

    async def handle_refresh_via_websocket(self, websocket):
        """Handle tunnel refresh via WebSocket"""
        try:
            print(" WebSocket refresh request received")
            await websocket.send(json.dumps({
                'type': 'refresh_status',
                'message': 'Refreshing tunnel on current port...'
            }))
            payload = await self._refresh_tunnel_payload()
            payload["type"] = "refresh_complete"
            await websocket.send(json.dumps(payload))
                
        except Exception as e:
            print(f" Error in handle_refresh_via_websocket: {e}")
            await websocket.send(json.dumps({
                'type': 'refresh_complete',
                'success': False,
                'error': str(e),
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
                'error': str(e),
            }))

    def process_event(self, event, websocket):
        action = event.get('action')
        if action in {'click', 'move', 'drag'}:
            self._handle_mouse_event(event)
        elif action == 'key':
            self._handle_key_event(event)
        elif action == 'scroll':
            self._handle_scroll_event(event)
        elif action == 'type_text':
            self._handle_type_text(event, websocket)
        elif action == 'get_clipboard':
            self._handle_get_clipboard(websocket)
        elif action == 'set_clipboard':
            self._handle_set_clipboard(event, websocket)
        elif action == 'set_clipboard_image':
            self._handle_set_clipboard_image(event, websocket)
        else:
            raise ValueError(f"Unsupported input action: {action}")

    def _send_clipboard_image_result(self, websocket, payload):
        self._queue_ws_send(websocket, json.dumps(payload), "_send_clipboard_image_result")

    @staticmethod
    def _validate_clipboard_image_signature(mime, image_data):
        if mime == "image/png":
            valid = image_data.startswith(b"\x89PNG\r\n\x1a\n")
        else:
            valid = image_data.startswith(b"\xff\xd8\xff")
        if not valid:
            raise ValueError(f"image data does not match declared MIME type {mime}")

    def _handle_set_clipboard_image(self, event, websocket):
        request_id = event.get("request_id")
        try:
            mime = str(event["mime"]).lower()
            if mime not in {"image/png", "image/jpeg", "image/jpg"}:
                raise ValueError("unsupported image MIME type")
            data_b64 = str(event["data_base64"])
            if not data_b64:
                raise ValueError("missing image data")

            max_bytes = self._clipboard_image_max_bytes
            if len(data_b64) > ((max_bytes + 2) // 3) * 4 + 4:
                raise ValueError("image clipboard payload exceeds size limit")
            image_data = base64.b64decode(data_b64, validate=True)
            if not image_data:
                raise ValueError("empty image data")
            if len(image_data) > max_bytes:
                raise ValueError("image clipboard payload exceeds size limit")
            self._validate_clipboard_image_signature(mime, image_data)

            with Image.open(io.BytesIO(image_data)) as source:
                if source.format not in {"PNG", "JPEG"}:
                    raise ValueError(f"unsupported decoded image format: {source.format}")
                if source.width * source.height > 25_000_000:
                    raise ValueError("image clipboard dimensions exceed 25 million pixels")
                dib = io.BytesIO()
                source.convert("RGB").save(dib, "BMP")
            win32clipboard.OpenClipboard()
            try:
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardData(win32con.CF_DIB, dib.getvalue()[14:])
            finally:
                win32clipboard.CloseClipboard()

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
            win32clipboard.OpenClipboard()
            try:
                if not win32clipboard.IsClipboardFormatAvailable(win32con.CF_UNICODETEXT):
                    raise RuntimeError("clipboard does not contain Unicode text")
                content = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
            finally:
                win32clipboard.CloseClipboard()
            max_bytes = self._clipboard_text_max_bytes
            if len(content.encode("utf-8")) > max_bytes:
                raise ValueError("clipboard text exceeds size limit")
            self._queue_ws_send(websocket, json.dumps({
                'type': 'clipboard_content',
                'data': content,
            }), "_handle_get_clipboard")
        except Exception as e:
            self._queue_ws_send(websocket, json.dumps({
                'type': 'clipboard_content',
                'error': str(e),
            }), "_handle_get_clipboard")

    def _handle_set_clipboard(self, event, websocket):
        try:
            data = event['data']
            if not isinstance(data, str):
                raise ValueError("clipboard data must be text")
            max_bytes = self._clipboard_text_max_bytes
            if len(data.encode("utf-8")) > max_bytes:
                raise ValueError("clipboard text exceeds size limit")
            win32clipboard.OpenClipboard()
            try:
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, data)
            finally:
                win32clipboard.CloseClipboard()
            self._queue_ws_send(websocket, json.dumps({
                'type': 'clipboard_set_result',
                'success': True,
                'length': len(data),
            }), "_handle_set_clipboard")
        except Exception as e:
            print(f" Error setting clipboard: {e}")
            self._queue_ws_send(websocket, json.dumps({
                'type': 'clipboard_set_result',
                'success': False,
                'error': str(e),
            }), "_handle_set_clipboard")

    def _map_view_norm_to_screen_norm(self, nx, ny):
        nx = _unit_float(nx, "x")
        ny = _unit_float(ny, "y")
        region = self.screen_capturer.get_active_region_norm() if self.screen_capturer else None
        if not region:
            return nx, ny
        left, top, right, bottom = region
        return left + nx * (right - left), top + ny * (bottom - top)

    def _handle_mouse_event(self, event):
        """Handle mouse move, drag, and click events through SendInput."""
        action = str(event['action']).lower()
        if action not in {'click', 'move', 'drag'}:
            raise ValueError(f"Unsupported mouse action: {action}")
        vx = user32.GetSystemMetrics(SM_XVIRTUALSCREEN)
        vy = user32.GetSystemMetrics(SM_YVIRTUALSCREEN)
        vw = user32.GetSystemMetrics(SM_CXVIRTUALSCREEN)
        vh = user32.GetSystemMetrics(SM_CYVIRTUALSCREEN)
        if vw < 2 or vh < 2:
            raise RuntimeError(f"Invalid virtual desktop size: {vw}x{vh}")

        if 'x' not in event or 'y' not in event:
            raise ValueError(f"{action} requires x and y")
        nx, ny = self._map_view_norm_to_screen_norm(event['x'], event['y'])
        px = vx + round(nx * (vw - 1))
        py = vy + round(ny * (vh - 1))
        ax = int(((px - vx) * 65535) / (vw - 1))
        ay = int(((py - vy) * 65535) / (vh - 1))
        _sendinput_mouse_move_abs(ax, ay)

        if action != 'click':
            return
        button = event.get('button')
        state = event.get('state')
        if not isinstance(button, str) or not isinstance(state, str):
            raise ValueError("click requires string button and state")
        button = button.lower()
        state = state.lower()
        try:
            flag = self.MOUSE_BUTTON_FLAGS[(button, state)]
        except KeyError as exc:
            raise ValueError(f"Unsupported mouse button/state: {button}/{state}") from exc
        _sendinput_mouse_button(flag)

    def _handle_key_event(self, event):
        """Handle a key event with Win32 SendInput."""
        key = event.get('key')
        state = event.get('state')
        if not isinstance(key, str) or not key:
            raise ValueError("key must be a non-empty string")
        if state not in {"down", "up"}:
            raise ValueError("key state must be down or up")
        key = key.lower()
        key_name = self.KEY_NAMES.get(key, key)
        _sendinput_key(key_name, state)

    def _broadcast_controller_alert(self, title: str, message: str):
        """Send an alert to all connected controller browsers (System B)."""
        clients = self.input_clients | self.video_clients

        self._controller_alert_seq += 1
        alert_id = f"{int(time.time() * 1000)}-{self._controller_alert_seq}"
        print(f" Broadcasting alert to {len(clients)} client socket(s)")
        payload = json.dumps({
            'type': 'controller_alert',
            'id': alert_id,
            'title': title,
            'message': message,
        })
        for ws in clients:
            self._queue_ws_send(ws, payload, "_broadcast_controller_alert")

    def _broadcast_alert_preset(self, code: str, cooldown_seconds: float = 0.4):
        if not isinstance(code, str):
            raise ValueError("Alert preset code must be a string")
        code = code.upper()
        self._load_alert_presets_from_settings()
        if code not in self.alert_presets:
            return False
        now = time.time()
        if now < self._alert_preset_cooldowns.get(code, 0.0):
            return False
        self._alert_preset_cooldowns[code] = now + cooldown_seconds
        title, message = self.alert_presets[code]
        self._broadcast_controller_alert(title, message)
        return True

    def _install_alert_hotkeys(self):
        """Register alert hotkeys for both NumLock states."""
        with self._hotkey_lock:
            # Remove any previously installed hotkeys
            for _id in self._hk_ids:
                try:
                    keyboard.remove_hotkey(_id)
                except Exception as exc:
                    _log_except("host_hotkeys.remove", exc)
            self._hk_ids = []

            add = self._hk_ids.append
            registered = 0

            def add_alert_hotkey(hotkey, callback):
                nonlocal registered
                add(keyboard.add_hotkey(hotkey, callback, suppress=False))
                registered += 1

            # Register both physical numpad names and navigation aliases so alerts
            # work whether NumLock is on or off.
            add_alert_hotkey('shift+num 1', lambda: self._begin_custom_alert_capture(source="hk"))
            add_alert_hotkey('shift+end', lambda: self._begin_custom_alert_capture(source="hk"))
            add_alert_hotkey('shift+num 3', lambda: self._end_custom_alert_capture(source="hk"))
            add_alert_hotkey('shift+pagedown', lambda: self._end_custom_alert_capture(source="hk"))

            for code, hotkeys in self.ALERT_PRESET_HOTKEYS.items():
                for hotkey in hotkeys:
                    add_alert_hotkey(hotkey, lambda code=code: self._broadcast_alert_preset(code))
            self._alert_hotkeys_active = registered > 0
            print(f" Alert hotkeys registered: {registered} (NumLock on/off supported)")
            return registered

    def _install_custom_capture_hotkeys(self):
        """Install per-key capture hotkeys and keep typed text out of the host app."""
        with self._hotkey_lock:
            def add(hk, fn):
                h = keyboard.add_hotkey(hk, fn, suppress=True, trigger_on_release=False)
                self._custom_capture_hotkeys.append(h)

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

    def _remove_custom_capture_hotkeys(self):
        """Remove the per-key capture hotkeys."""
        with self._hotkey_lock:
            try:
                for h in self._custom_capture_hotkeys:
                    keyboard.remove_hotkey(h)
            finally:
                self._custom_capture_hotkeys = []
                print("[custom] capture hotkeys removed")

    def _begin_custom_alert_capture(self, source: str = "hotkey"):
        """Enter capture mode and start buffering (idempotent)."""
        if self.custom_alert_active:
            return
        self.custom_alert_active = True
        self.custom_alert_buf = []
        self._install_custom_capture_hotkeys()
        print(f"  CAPTURE: ON (source={source})")

    def _end_custom_alert_capture(self, source: str = "hotkey"):
        """Leave capture mode and send the buffered text to System B (idempotent)."""
        if not self.custom_alert_active:
            return
        self._remove_custom_capture_hotkeys()
        text = ''.join(self.custom_alert_buf)
        self.custom_alert_active = False
        self.custom_alert_buf = []
        print(f"  CAPTURE: OFF (source={source})  sending {len(text)} chars")
        if text:
            print(f"[custom] broadcasting custom text ({len(text)} chars)")
            self._broadcast_controller_alert("Custom", text)
        else:
            print("[custom] no custom text captured; nothing to broadcast")

    def start_host_hotkeys(self):
        if self._alert_hotkeys_active:
            return len(self._hk_ids)
        self._register_keyboard_cleanup()
        registered = self._install_alert_hotkeys()
        if registered <= 0:
            raise RuntimeError("No host alert hotkeys were registered")
        return registered

    def stop_host_hotkeys(self):
        with self._hotkey_lock:
            try:
                self._remove_custom_capture_hotkeys()
            except Exception as exc:
                _log_except("host_hotkeys.custom_remove", exc)
            for _id in self._hk_ids:
                try:
                    keyboard.remove_hotkey(_id)
                except Exception as exc:
                    _log_except("host_hotkeys.remove", exc)
            self._hk_ids = []
            self._alert_hotkeys_active = False

    def _type_text_chunk(self, text):
        if not text:
            return
        index = 0
        while index < len(text):
            char = text[index]
            if char == '\r':
                if index + 1 < len(text) and text[index + 1] == '\n':
                    index += 1
                _sendinput_key('enter', 'press')
            elif char == '\n':
                _sendinput_key('enter', 'press')
            elif char == '\t':
                _sendinput_key('tab', 'press')
            else:
                end = index
                while end < len(text) and text[end] not in '\r\n\t':
                    end += 1
                chunk = text[index:end]
                _sendinput_unicode(chunk)
                index = end
                continue
            index += 1

    def _typing_mode_from_event(self, event):
        mode = event.get('typing_mode')
        if not isinstance(mode, str):
            raise ValueError("typing_mode must be exact or ide")
        mode = mode.strip().lower()
        if mode not in {'exact', 'ide'}:
            raise ValueError("typing_mode must be exact or ide")
        return mode

    def _event_flag(self, event, name):
        value = event.get(name, False)
        if not isinstance(value, bool):
            raise ValueError(f"{name} must be a boolean")
        return value

    def _normalize_text_for_typing_mode(self, text, typing_mode):
        if typing_mode == 'ide':
            return re.sub(r"(\r\n|\r|\n)[\t ]+", lambda match: match.group(1), text)
        return text

    def _handle_type_text(self, event, websocket):
        """Handle text typing; direct chunks bypass live-typing diff state."""
        text = event.get('text')
        if not isinstance(text, str):
            raise ValueError("text must be a string")
        if len(text) > 100_000:
            raise ValueError("text exceeds 100000 character limit")
        typing_mode = self._typing_mode_from_event(event)
        if self._event_flag(event, 'direct'):
            text = self._normalize_text_for_typing_mode(text, typing_mode)
            if text:
                self._type_text_chunk(text)
            return
        previous = self.live_typing_text_by_client.get(websocket, '')
        max_len = min(len(previous), len(text))
        prefix_len = 0
        while prefix_len < max_len and previous[prefix_len] == text[prefix_len]:
            prefix_len += 1

        typed_delta = False
        if len(text) > len(previous) and prefix_len == len(previous):
            appended = self._normalize_text_for_typing_mode(text[len(previous):], typing_mode)
            if appended:
                self._type_text_chunk(appended)
                typed_delta = True
        else:
            suffix_len = 0
            remaining_previous = len(previous) - prefix_len
            remaining_text = len(text) - prefix_len
            while (
                suffix_len < remaining_previous
                and suffix_len < remaining_text
                and previous[len(previous) - 1 - suffix_len] == text[len(text) - 1 - suffix_len]
            ):
                suffix_len += 1
            if len(text) > len(previous) and remaining_previous == suffix_len:
                inserted = self._normalize_text_for_typing_mode(
                    text[prefix_len:len(text) - suffix_len],
                    typing_mode,
                )
                if inserted:
                    self._type_text_chunk(inserted)
                    typed_delta = True
        if not typed_delta and text != previous:
            logging.debug("Live typing edit is handled by key events")
        self.live_typing_text_by_client[websocket] = text

    def _handle_scroll_event(self, event):
        x = event.get('x')
        y = event.get('y')
        if (x is None) != (y is None):
            raise ValueError("scroll x and y must be provided together")
        if x is not None and y is not None:
            self._handle_mouse_event({'action': 'move', 'x': x, 'y': y})
        if 'deltaY' not in event or isinstance(event['deltaY'], bool):
            raise ValueError("deltaY must be a non-zero number")
        try:
            dy = float(event['deltaY'])
        except (TypeError, ValueError) as exc:
            raise ValueError("deltaY must be a non-zero number") from exc
        if not math.isfinite(dy) or dy == 0:
            raise ValueError("deltaY must be a non-zero number")
        wheel_data = (-1 if dy > 0 else 1) * 120
        inp = INPUT()
        inp.type = 0
        inp.union.mi = MOUSEINPUT(0, 0, wheel_data, MOUSEEVENTF_WHEEL, 0, 0)
        if user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT)) != 1:
            raise OSError(
                f"SendInput mouse wheel failed (GetLastError={ctypes.get_last_error()})"
            )
