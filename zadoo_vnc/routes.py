"""HTTP route and snapshot handlers."""
from __future__ import annotations

import asyncio
import http
import io
import json
import logging
import os
import secrets
import sys
import time
import urllib.parse
from http.cookies import SimpleCookie

from websockets.datastructures import Headers
from websockets.http11 import Response as WSResponse

from .assets import load_benchmark_html, load_host_controls_html, load_index_html, load_terminal_html
from .camera_discovery import enumerate_camera_devices
from .config import BRAND_HEADER_IMAGE_PATH, SPLASH_IMAGE_PATH, TRIGGER_ICON_IMAGE_PATH, env_int
from .dependencies import HAS_FAST_CTYPES, HAS_IMAGECODECS, HAS_MSS, HAS_SOUNDDEVICE, Image, fast_ctypes_screenshots, imagecodecs, mss, np, sd
from .dpi import get_primary_screen_size
from .logging_utils import _log_except, _log_fallback, _log_try_ok

class RoutesMixin:
    AUTH_COOKIE_NAME = "zadoo_auth"
    CSRF_COOKIE_NAME = "zadoo_csrf"
    AUTH_TTL_SECONDS = 3600
    FALSE_VALUES = {"0", "false", "no", "off", "disabled"}
    VIEW_ACTIONS = {
        "client_stream_stats",
        "cursor_broadcast",
        "get_available_capture_methods",
        "get_capture_stats",
        "get_public_url",
        "get_stream_status",
        "mic_event",
        "snap_event",
        "stream_ping",
        "verify_capture_method",
    }
    CONTROL_ACTIONS = {
        "click",
        "control",
        "drag",
        "get_clipboard",
        "key",
        "key_combo",
        "move",
        "scroll",
        "set_clipboard",
        "set_clipboard_image",
        "set_fps",
        "set_quality",
        "type_text",
    }
    ADVANCED_ACTIONS = {"set_capture_method", "set_performance"}
    HOST_ACTIONS = {"refresh_tunnel", "toggle_keystroke_capture"}
    ROUTE_FEATURES = {
        "/": "public",
        "/api/auth": "public",
        "/brand-header.png": "public",
        "/trigger-icon.png": "public",
        "/splash.png": "public",
        "/video": "view",
        "/audio": "view",
        "/input": "control",
        "/api/public-url": "view",
        "/api/stream-stats": "view",
        "/benchmark.html": "view",
        "/snapshot": "view",
        "/api/set-quality": "control",
        "/api/set-fps": "control",
        "/api/refresh-tunnel": "host",
        "/terminal.html": "terminal",
        "/ssh": "terminal",
        "/webcam": "webcam",
        "/api/list-cameras": "webcam",
        "/mic": "mic",
        "/api/list-mics": "mic",
        "/host-controls": "host",
        "/api/alert": "host",
    }
    CSRF_HTTP_FEATURES = {"control", "host"}

    def _json_response(self, payload, status=http.HTTPStatus.OK, extra_headers=None):
        headers = Headers()
        headers["Content-Type"] = "application/json; charset=utf-8"
        headers["Cache-Control"] = "no-store"
        for key, value in (extra_headers or {}).items():
            if isinstance(value, (list, tuple)):
                for item in value:
                    headers[key] = str(item)
            else:
                headers[key] = str(value)
        body = json.dumps(payload).encode("utf-8")
        return WSResponse(
            status_code=int(status),
            reason_phrase=status.phrase,
            headers=headers,
            body=body,
        )

    def _plain_response(self, body, status=http.HTTPStatus.FORBIDDEN):
        headers = Headers()
        headers["Content-Type"] = "text/plain; charset=utf-8"
        headers["Cache-Control"] = "no-store"
        return WSResponse(
            status_code=int(status),
            reason_phrase=status.phrase,
            headers=headers,
            body=str(body).encode("utf-8"),
        )

    def _png_file_response(self, image_path, missing_message):
        try:
            with open(image_path, "rb") as file:
                body = file.read()
            headers = Headers()
            headers["Content-Type"] = "image/png"
            headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
            return WSResponse(
                status_code=int(http.HTTPStatus.OK),
                reason_phrase=http.HTTPStatus.OK.phrase,
                headers=headers,
                body=body,
            )
        except Exception:
            return self._plain_response(missing_message, http.HTTPStatus.NOT_FOUND)

    def _primary_screen_size_or_none(self, log_context):
        try:
            return get_primary_screen_size()
        except Exception as exc:
            _log_fallback(log_context, "pyautogui.size", "primary_screen_size_failed", exc)
            try:
                import pyautogui as _pg
                return _pg.size()
            except Exception:
                return None

    def _header_get(self, request_headers, name, default=None):
        try:
            return request_headers.get(name, default)
        except Exception:
            return default

    def _env_enabled(self, name, default="0"):
        return str(os.getenv(name, default)).strip().lower() not in self.FALSE_VALUES

    def _request_identity(self, request_headers):
        for name in ("CF-Connecting-IP", "X-Real-IP", "X-Forwarded-For"):
            raw = self._header_get(request_headers, name, "")
            if raw:
                return str(raw).split(",", 1)[0].strip() or "default"
        return self._header_get(request_headers, "Host", "default") or "default"

    def _auth_lockout_remaining(self, request_headers):
        key = self._request_identity(request_headers)
        state = getattr(self, "_auth_failures", {}).get(key)
        if not isinstance(state, dict):
            return 0
        remaining = float(state.get("locked_until", 0.0) or 0.0) - time.time()
        return max(0, int(remaining))

    def _record_auth_failure(self, request_headers):
        key = self._request_identity(request_headers)
        failures = getattr(self, "_auth_failures", None)
        if not isinstance(failures, dict):
            self._auth_failures = {}
            failures = self._auth_failures
        now = time.time()
        window = env_int("ZADOO_AUTH_WINDOW_SECONDS", 60, 1, 3600)
        max_failures = env_int("ZADOO_AUTH_MAX_FAILURES", 5, 1, 100)
        lockout = env_int("ZADOO_AUTH_LOCKOUT_SECONDS", 300, 1, 86400)
        state = failures.setdefault(key, {"times": [], "locked_until": 0.0})
        times = [float(ts) for ts in state.get("times", []) if now - float(ts) <= window]
        times.append(now)
        state["times"] = times
        if len(times) >= max_failures:
            state["locked_until"] = now + lockout
            return lockout
        return 0

    def _record_auth_success(self, request_headers):
        try:
            getattr(self, "_auth_failures", {}).pop(self._request_identity(request_headers), None)
        except Exception:
            pass

    def _request_origin_allowed(self, request_headers):
        origin = self._header_get(request_headers, "Origin", "")
        if not origin:
            return True
        allowed = [
            item.strip().rstrip("/")
            for item in str(os.getenv("ZADOO_ALLOWED_ORIGINS", "")).split(",")
            if item.strip()
        ]
        if "*" in allowed:
            return True
        try:
            parsed = urllib.parse.urlparse(origin)
            if parsed.scheme not in {"http", "https"} or not parsed.netloc:
                return False
            normalized_origin = f"{parsed.scheme.lower()}://{parsed.netloc.lower()}"
            if normalized_origin in {item.lower() for item in allowed}:
                return True
            host = str(self._header_get(request_headers, "Host", "") or "").lower()
            return bool(host and parsed.netloc.lower() == host)
        except Exception:
            return False

    def _state_changing_http_allowed(self, request_headers):
        header_token = str(self._header_get(request_headers, "X-Zadoo-CSRF", "") or "")
        if not header_token:
            return False
        session = self._session_for_headers(request_headers)
        if not isinstance(session, dict):
            return False
        expected = str(session.get("csrf") or "")
        if not expected:
            return False
        return secrets.compare_digest(header_token, expected)

    def _headers_for_websocket(self, websocket):
        try:
            request_headers = getattr(getattr(websocket, "request", None), "headers", None)
            if request_headers is None:
                request_headers = getattr(websocket, "request_headers", None)
            return request_headers
        except Exception:
            return None

    def _role_for_headers(self, request_headers):
        session = self._session_for_headers(request_headers)
        if not isinstance(session, dict):
            return None
        return session.get("role")

    def _feature_for_action(self, action):
        action = str(action or "")
        if action in self.VIEW_ACTIONS:
            return "view"
        if action in self.CONTROL_ACTIONS:
            return "control"
        if action in self.ADVANCED_ACTIONS:
            return "advanced"
        if action in self.HOST_ACTIONS:
            return "host"
        return None

    def _is_ws_action_authorized_for_headers(self, request_headers, action):
        feature = self._feature_for_action(action)
        if feature is None:
            return False
        return self._role_allows(self._role_for_headers(request_headers), feature)

    def _is_ws_action_authorized(self, websocket, action):
        return self._is_ws_action_authorized_for_headers(self._headers_for_websocket(websocket), action)

    async def _send_ws_forbidden(self, websocket, action):
        try:
            await websocket.send(json.dumps({
                "type": "error",
                "error": "Forbidden",
                "action": action,
            }))
        except Exception:
            pass

    def _cookie_value(self, request_headers, name):
        cookie_header = self._header_get(request_headers, "Cookie", "")
        if not isinstance(cookie_header, str) or not cookie_header:
            return None
        try:
            parsed = SimpleCookie()
            parsed.load(cookie_header)
            morsel = parsed.get(name)
            return morsel.value if morsel else None
        except Exception:
            return None

    def _auth_codes(self):
        def clean(value):
            return str(value if value is not None else "").strip()

        codes = {
            "full": clean(os.getenv("CODE_FULL")),
            "limited": clean(os.getenv("CODE_LIMITED")),
            "partial": clean(os.getenv("CODE_PARTIAL")),
            "lockdown": clean(os.getenv("CODE_LOCKDOWN")),
        }
        custom = clean(os.getenv("CUSTOM_PASSWORD"))
        if custom:
            codes["custom"] = custom
        configured = {role: value.upper() for role, value in codes.items() if value}
        if configured:
            return configured
        runtime_codes = getattr(self, "_runtime_auth_codes", None)
        if not isinstance(runtime_codes, dict) or not runtime_codes:
            code = secrets.token_urlsafe(6).replace("-", "").replace("_", "").upper()[:6]
            runtime_codes = {"full": code}
            self._runtime_auth_codes = runtime_codes
            self._runtime_auth_generated = True
        return runtime_codes

    def _announce_auth_codes(self):
        codes = self._auth_codes()
        if getattr(self, "_runtime_auth_generated", False):
            code = codes.get("full")
            if code:
                print("=" * 60)
                print(f" Temporary full access code: {code}")
                print(" Set CODE_FULL in .env to use your own persistent code.")
                print("=" * 60)
        else:
            print(" Access code source: environment")

    def _match_auth_code(self, code):
        submitted = str(code or "").strip().upper()
        if not submitted:
            return None
        for role, expected in self._auth_codes().items():
            expected_clean = str(expected or "").strip().upper()
            if secrets.compare_digest(submitted, expected_clean):
                return "full" if role == "custom" else role
        return None

    def _cleanup_auth_sessions(self):
        sessions = getattr(self, "auth_sessions", None)
        if not isinstance(sessions, dict):
            self.auth_sessions = {}
            return
        now = time.time()
        expired = [
            token
            for token, session in sessions.items()
            if not isinstance(session, dict) or float(session.get("expires_at", 0)) <= now
        ]
        for token in expired:
            sessions.pop(token, None)

    def _session_for_headers(self, request_headers):
        token = self._cookie_value(request_headers, self.AUTH_COOKIE_NAME)
        if not token:
            return None
        session = getattr(self, "auth_sessions", {}).get(token)
        if not isinstance(session, dict):
            return None
        if float(session.get("expires_at", 0)) <= time.time():
            getattr(self, "auth_sessions", {}).pop(token, None)
            return None
        return session

    def _role_allows(self, role, feature):
        role = str(role or "")
        if feature == "public":
            return True
        if feature == "view":
            return role in {"full", "limited", "partial", "lockdown"}
        if feature == "control":
            return role in {"full", "limited", "partial"}
        if feature == "advanced":
            return role in {"full", "partial"}
        if feature in {"terminal", "webcam", "mic", "host"}:
            return role == "full"
        return False

    def _is_authorized(self, request_headers, feature="view"):
        session = self._session_for_headers(request_headers)
        return bool(session and self._role_allows(session.get("role"), feature))

    def _feature_for_route(self, route_path):
        if route_path.startswith("/api/client-log"):
            return "view"
        return self.ROUTE_FEATURES.get(route_path)

    def _http_route_requires_csrf(self, route_path, feature=None):
        if feature is None:
            feature = self._feature_for_route(route_path)
        return route_path.startswith("/api/") and feature in self.CSRF_HTTP_FEATURES

    def _is_ws_authorized(self, route_path, request_headers):
        feature = self._feature_for_route(route_path)
        if feature is None:
            return False
        return self._is_authorized(request_headers, feature)

    def _auth_code_from_request(self, path, request_headers):
        header_code = self._header_get(request_headers, "X-Zadoo-Code", "")
        if header_code:
            return str(header_code), "header"
        parsed = urllib.parse.urlparse(str(path or ""))
        query = urllib.parse.parse_qs(parsed.query or "")
        query_code = (query.get("code") or [""])[0]
        if query_code:
            if not self._env_enabled("ZADOO_ALLOW_QUERY_AUTH", "0"):
                return None, "query_disabled"
            return str(query_code), "query"
        return "", "missing"

    async def handle_auth(self, path, request_headers=None):
        lockout_remaining = self._auth_lockout_remaining(request_headers)
        if lockout_remaining > 0:
            return self._json_response(
                {"success": False, "error": "Too many invalid attempts", "retry_after": lockout_remaining},
                http.HTTPStatus.TOO_MANY_REQUESTS,
            )
        parsed = urllib.parse.urlparse(str(path or ""))
        code, source = self._auth_code_from_request(parsed.geturl(), request_headers)
        if source == "query_disabled":
            return self._json_response(
                {"success": False, "error": "Query string auth is disabled"},
                http.HTTPStatus.BAD_REQUEST,
            )
        role = self._match_auth_code(code)
        if not role:
            if code:
                retry_after = self._record_auth_failure(request_headers)
                if retry_after:
                    return self._json_response(
                        {"success": False, "error": "Too many invalid attempts", "retry_after": retry_after},
                        http.HTTPStatus.TOO_MANY_REQUESTS,
                    )
            return self._json_response({"success": False, "error": "Invalid code"}, http.HTTPStatus.UNAUTHORIZED)

        self._record_auth_success(request_headers)
        token = secrets.token_urlsafe(32)
        expires_at = time.time() + self.AUTH_TTL_SECONDS
        self._cleanup_auth_sessions()
        sessions = getattr(self, "auth_sessions", None)
        if not isinstance(sessions, dict):
            self.auth_sessions = {}
            sessions = self.auth_sessions
        csrf_token = secrets.token_urlsafe(32)
        sessions[token] = {"role": role, "expires_at": expires_at, "csrf": csrf_token}
        auth_cookie = (
            f"{self.AUTH_COOKIE_NAME}={token}; Path=/; Max-Age={self.AUTH_TTL_SECONDS}; "
            "HttpOnly; SameSite=Lax"
        )
        csrf_cookie = (
            f"{self.CSRF_COOKIE_NAME}={csrf_token}; Path=/; Max-Age={self.AUTH_TTL_SECONDS}; "
            "SameSite=Lax"
        )
        return self._json_response(
            {"success": True, "mode": role, "expires_in": self.AUTH_TTL_SECONDS, "csrf_token": csrf_token},
            extra_headers={"Set-Cookie": [auth_cookie, csrf_cookie]},
        )

    def enumerate_cameras(self):
        return enumerate_camera_devices()

    def enumerate_microphones(self):
        if not HAS_SOUNDDEVICE or sd is None:
            return []
        try:
            devices = list(sd.query_devices())
        except Exception:
            return []
        try:
            hostapis = list(sd.query_hostapis())
        except Exception:
            hostapis = []
        try:
            default_input = sd.default.device[0]
        except Exception:
            default_input = None
        try:
            if default_input is not None and int(default_input) < 0:
                default_input = None
        except Exception:
            default_input = None

        def hostapi_name(index):
            try:
                api = hostapis[int(index)]
                return str(api.get("name") or "").strip()
            except Exception:
                return ""

        result = []
        default_name = "System Default"
        if default_input is not None:
            try:
                default_dev = devices[int(default_input)]
                default_name = str(default_dev.get("name") or default_name)
            except Exception:
                pass
        result.append({
            "id": "default",
            "device_index": None,
            "name": default_name,
            "label": f"System Default ({default_name})",
            "default": True,
            "channels": 1,
            "samplerate": None,
        })

        seen = set()
        for index, device in enumerate(devices):
            try:
                max_input = int(device.get("max_input_channels") or 0)
            except Exception:
                max_input = 0
            if max_input <= 0:
                continue
            name = str(device.get("name") or f"Microphone {index}").strip()
            api = hostapi_name(device.get("hostapi"))
            label = f"{name} ({api})" if api else name
            key = (name.casefold(), api.casefold(), max_input)
            if key in seen:
                label = f"{label} #{index}"
            seen.add(key)
            try:
                samplerate = int(float(device.get("default_samplerate") or 0)) or None
            except Exception:
                samplerate = None
            result.append({
                "id": str(index),
                "device_index": index,
                "name": name,
                "label": label,
                "default": default_input is not None and int(index) == int(default_input),
                "channels": max_input,
                "samplerate": samplerate,
            })
        return result

    def _apply_quality(self, raw_value, default=85):
        try:
            value = max(1, min(100, int(raw_value)))
        except Exception:
            value = max(1, min(100, int(default)))
        self.current_quality = value
        self._quality_locked_by_user = True
        if self.screen_capturer:
            self.screen_capturer.quality = value
        try:
            # Quality controls compression and visual resolution. Re-apply the
            # active adaptive profile so high quality immediately restores
            # full-resolution capture instead of waiting for the next profile change.
            previous_fps = getattr(self, "current_fps", None)
            self._apply_stream_profile("quality_changed")
            if previous_fps is not None:
                self.current_fps = previous_fps
                if self.screen_capturer:
                    self.screen_capturer.fps = previous_fps
        except Exception:
            pass
        return value

    def _apply_fps(self, raw_value=None, default=0):
        try:
            raw_text = str(raw_value if raw_value is not None else "").strip().lower()
            if raw_text in {"", "0", "max", "auto", "unlimited", "none"}:
                value = 0
            else:
                value = max(1, int(raw_value))
        except Exception:
            try:
                value = max(0, int(default))
            except Exception:
                value = 0
        self.current_fps = value
        if self.screen_capturer:
            self.screen_capturer.fps = value
        return value

    async def process_request(self, *args, **kwargs):
        """Process HTTP requests - compatible with websockets v10-v15.

        Accepts either (path, request_headers) or a single ServerConnection object.
        """
        path = None
        request_headers = None
        request_body = b""

        # Unpack arguments depending on websockets version
        try:
            # websockets classic style: (path, request_headers)
            if len(args) >= 2 and isinstance(args[0], str):
                path = args[0]
                request_headers = args[1]
            elif len(args) >= 2:
                # websockets v15 style: (ServerConnection, Request)
                connection, request = args[0], args[1]
                path = getattr(request, "path", None)
                request_headers = getattr(request, "headers", None)
                request_body = getattr(request, "body", b"") or b""
            elif len(args) == 1:
                # Fallback: older style may pass a single connection-like object
                connection = args[0]
                # Try to resolve request then path
                request = getattr(connection, "request", None)
                if request is not None:
                    path = getattr(request, "path", None)
                    request_headers = getattr(request, "headers", None)
                    request_body = getattr(request, "body", b"") or b""
                else:
                    path = getattr(connection, "path", None)
                    request_headers = getattr(connection, "request_headers", None)
        except Exception:
            pass

        if not isinstance(path, str):
            path = "/"
        route_path = urllib.parse.urlparse(path).path

        logging.debug("process_request: path=%s", path)

        if not self._request_origin_allowed(request_headers):
            return self._plain_response("Forbidden", http.HTTPStatus.FORBIDDEN)

        # WebSocket upgrades must be authenticated before the stream handlers run.
        try:
            hdrs = request_headers
            get = None
            try:
                get = hdrs.get  # websockets Headers
            except Exception:
                get = None
            upgrade_val = None
            connection_val = None
            if get:
                try:
                    upgrade_val = (get("Upgrade") or get("upgrade") or "").lower()
                    connection_val = (get("Connection") or get("connection") or "").lower()
                except Exception:
                    pass
            if (isinstance(upgrade_val, str) and "websocket" in upgrade_val) or (
                isinstance(connection_val, str) and "upgrade" in connection_val
            ):
                if self._is_ws_authorized(route_path, hdrs):
                    return None
                return self._plain_response("Forbidden", http.HTTPStatus.FORBIDDEN)
        except Exception:
            return self._plain_response("Forbidden", http.HTTPStatus.FORBIDDEN)

        feature = self._feature_for_route(route_path)
        if feature and feature != "public" and not self._is_authorized(request_headers, feature):
            return self._plain_response("Forbidden", http.HTTPStatus.FORBIDDEN)
        if self._http_route_requires_csrf(route_path, feature) and not self._state_changing_http_allowed(request_headers):
            return self._plain_response("Forbidden", http.HTTPStatus.FORBIDDEN)

        # Process routes
        if route_path == "/":
            headers = Headers()
            headers["Content-Type"] = "text/html; charset=utf-8"
            try:
                headers["Permissions-Policy"] = "clipboard-read=(self), clipboard-write=(self)"
            except Exception:
                pass
            try:
                headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
            except Exception:
                pass
            try:
                inj = (
                    "<script>"
                    "(function(){\n"
                    "  function safeUnblock(){ try{ if(typeof setSuppressHostInput==='function'){ setSuppressHostInput(false); } }catch(e){} }\n"
                    "  function enableCursor(){ try{ if(typeof sendEvent==='function'){ sendEvent({ action: 'cursor_broadcast', enabled: true }); } }catch(e){} }\n"
                    "  window.addEventListener('load', function(){ safeUnblock(); enableCursor(); });\n"
                    "  document.addEventListener('keydown', function(e){ if(e.key==='Escape' && window.__snapSelecting){ safeUnblock(); } });\n"
                    "  document.addEventListener('mouseup', function(){ if(window.__snapSelecting){ safeUnblock(); } });\n"
                    "  window.addEventListener('beforeunload', function(){ safeUnblock(); });\n"
                    "})();"
                    "</script>"
                )
                html = load_index_html().replace("</head>", inj + "\n</head>")
            except Exception:
                html = load_index_html()
            return WSResponse(
                status_code=int(http.HTTPStatus.OK),
                reason_phrase=http.HTTPStatus.OK.phrase,
                headers=headers,
                body=html.encode("utf-8"),
            )
        elif route_path == "/api/auth":
            return await self.handle_auth(path, request_headers)
        elif route_path == "/api/public-url":
            try:
                return self._json_response(json.loads(await self.handle_get_public_url()))
            except Exception as e:
                return self._json_response({"success": False, "error": str(e)}, http.HTTPStatus.INTERNAL_SERVER_ERROR)
        elif route_path == "/api/refresh-tunnel":
            try:
                return self._json_response(json.loads(await self.handle_refresh_tunnel()))
            except Exception as e:
                return self._json_response({"success": False, "error": str(e)}, http.HTTPStatus.INTERNAL_SERVER_ERROR)
        elif route_path == "/api/set-quality":
            return self._json_response(await self.handle_set_quality(path))
        elif route_path == "/api/set-fps":
            return self._json_response(await self.handle_set_fps(path))
        elif route_path == "/api/stream-stats":
            try:
                stats = self._capture_stats_payload()
                stream = self._stream_status_payload()
                return self._json_response({"success": True, "stats": stats, "stream": stream})
            except Exception as e:
                return self._json_response({"success": False, "error": str(e)}, http.HTTPStatus.INTERNAL_SERVER_ERROR)
        elif route_path.startswith("/api/client-log"):
            try:
                # Accept simple GET with ?msg=... or POST with text body
                if request_headers is None:
                    body_bytes = b""
                else:
                    body_bytes = request_body
                parsed = urllib.parse.urlparse(path)
                qs = urllib.parse.parse_qs(parsed.query or "")
                msg = (qs.get("msg") or [""])[0]
                if not msg and isinstance(body_bytes, (bytes, bytearray)) and body_bytes:
                    try:
                        msg = body_bytes.decode("utf-8", "ignore")
                    except Exception:
                        msg = str(body_bytes)
                msg = urllib.parse.unquote(msg)
                _log_try_ok("client.log", msg[:500])
                headers = Headers()
                headers["Content-Type"] = "application/json; charset=utf-8"
                return WSResponse(
                    status_code=int(http.HTTPStatus.OK),
                    reason_phrase=http.HTTPStatus.OK.phrase,
                    headers=headers,
                    body=b"{\"ok\":true}",
                )
            except Exception as e:
                _log_except("client.log", e)
                headers = Headers()
                headers["Content-Type"] = "application/json; charset=utf-8"
                return WSResponse(
                    status_code=int(http.HTTPStatus.INTERNAL_SERVER_ERROR),
                    reason_phrase=http.HTTPStatus.INTERNAL_SERVER_ERROR.phrase,
                    headers=headers,
                    body=b"{\"ok\":false}",
                )
        elif route_path == "/terminal.html":
            term_html = load_terminal_html().encode("utf-8")
            headers = Headers()
            headers["Content-Type"] = "text/html; charset=utf-8"
            return WSResponse(
                status_code=int(http.HTTPStatus.OK),
                reason_phrase=http.HTTPStatus.OK.phrase,
                headers=headers,
                body=term_html,
            )
        elif route_path == "/benchmark.html":
            bench_html = load_benchmark_html().encode("utf-8")
            headers = Headers()
            headers["Content-Type"] = "text/html; charset=utf-8"
            headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
            return WSResponse(
                status_code=int(http.HTTPStatus.OK),
                reason_phrase=http.HTTPStatus.OK.phrase,
                headers=headers,
                body=bench_html,
            )
        elif route_path == "/video":
            return None  # Let WebSocket handler take over
        elif route_path == "/input":
            return None  # Let WebSocket handler take over
        elif route_path == "/audio":
            return None  # Let WebSocket handler take over
        elif route_path == "/ssh":
            return None  # WebSocket bridge for SSH
        elif route_path == "/webcam":
            return None  # Let WebSocket handler take over
        elif route_path == "/api/list-cameras":
            try:
                loop = asyncio.get_running_loop()
                devices = await loop.run_in_executor(None, self.enumerate_cameras)
                payload = json.dumps({ 'success': True, 'devices': devices }).encode('utf-8')
            except Exception:
                payload = json.dumps({ 'success': False, 'devices': [] }).encode('utf-8')
            headers = Headers()
            headers["Content-Type"] = "application/json; charset=utf-8"
            return WSResponse(
                status_code=int(http.HTTPStatus.OK),
                reason_phrase=http.HTTPStatus.OK.phrase,
                headers=headers,
                body=payload,
            )
        elif route_path == "/api/list-mics":
            try:
                loop = asyncio.get_running_loop()
                devices = await loop.run_in_executor(None, self.enumerate_microphones)
                payload = json.dumps({"success": True, "devices": devices}).encode("utf-8")
            except Exception:
                payload = json.dumps({"success": False, "devices": []}).encode("utf-8")
            headers = Headers()
            headers["Content-Type"] = "application/json; charset=utf-8"
            headers["Cache-Control"] = "no-store"
            return WSResponse(
                status_code=int(http.HTTPStatus.OK),
                reason_phrase=http.HTTPStatus.OK.phrase,
                headers=headers,
                body=payload,
            )
        elif route_path == "/host-controls":
            page = load_host_controls_html().strip().encode("utf-8")
            headers = Headers()
            headers["Content-Type"] = "text/html; charset=utf-8"
            return WSResponse(
                status_code=int(http.HTTPStatus.OK),
                reason_phrase=http.HTTPStatus.OK.phrase,
                headers=headers,
                body=page,
            )
        elif route_path == "/api/alert":
            # Accept GET or POST (websockets.process_request exposes only path/headers)
            try:
                qs = urllib.parse.parse_qs(urllib.parse.urlparse(path).query or "")
                code = (qs.get("code") or [""])[0].upper().strip()
                _log_try_ok("api.alert.parse", code)
            except Exception:
                code = ""
                _log_except("api.alert.parse", sys.exc_info()[1])
            try:
                title, message = self.alert_presets.get(code, ("Alert", f"Code: {code}"))
                _log_try_ok("api.alert.lookup", f"{title}|{message}")
            except Exception:
                title, message = ("Alert", f"Code: {code}")
                _log_except("api.alert.lookup", sys.exc_info()[1])
            try:
                self._broadcast_controller_alert(title, message)
                resp = b'{"ok":true}'
                headers = Headers()
                headers["Content-Type"] = "application/json; charset=utf-8"
                _log_try_ok("api.alert.broadcast")
                return WSResponse(
                    status_code=int(http.HTTPStatus.OK),
                    reason_phrase=http.HTTPStatus.OK.phrase,
                    headers=headers,
                    body=resp,
                )
            except Exception as e:
                err = ("{\"ok\":false,\"error\":\"" + str(e) + "\"}").encode("utf-8", "ignore")
                headers = Headers()
                headers["Content-Type"] = "application/json; charset=utf-8"
                _log_except("api.alert.broadcast", e)
                return WSResponse(
                    status_code=int(http.HTTPStatus.INTERNAL_SERVER_ERROR),
                    reason_phrase=http.HTTPStatus.INTERNAL_SERVER_ERROR.phrase,
                    headers=headers,
                    body=err,
                )
        elif route_path == "/brand-header.png":
            return self._png_file_response(BRAND_HEADER_IMAGE_PATH, "Header image not found")
        elif route_path == "/trigger-icon.png":
            return self._png_file_response(TRIGGER_ICON_IMAGE_PATH, "Trigger icon not found")
        elif route_path == "/splash.png":
            return self._png_file_response(SPLASH_IMAGE_PATH, "Splash image not found")
        elif route_path == "/snapshot":
            try:
                t_req0 = time.perf_counter()
                parsed_url = urllib.parse.urlparse(path)
                query_params = urllib.parse.parse_qs(parsed_url.query)

                rect_norm = None
                if all(k in query_params for k in ['x0', 'y0', 'x1', 'y1']):
                    try:
                        rect_norm = {
                            'x0': float(query_params['x0'][0]),
                            'y0': float(query_params['y0'][0]),
                            'x1': float(query_params['x1'][0]),
                            'y1': float(query_params['y1'][0])
                        }
                    except (ValueError, IndexError):
                        pass
                fmt = query_params.get('fmt', ['png'])[0].lower().strip()
                if fmt not in ('jpeg', 'jpg', 'png'):
                    fmt = 'png'
                try:
                    quality = int(query_params.get('q', [85])[0])
                except Exception:
                    quality = 85
                try:
                    max_w = int(query_params.get('max_w', [0])[0]) or None
                except Exception:
                    max_w = None
                try:
                    max_h = int(query_params.get('max_h', [0])[0]) or None
                except Exception:
                    max_h = None
                try:
                    logging.info(
                        "[snapshot.request] fmt=%s q=%s max_w=%s max_h=%s rect=%s",
                        fmt,
                        quality,
                        str(max_w),
                        str(max_h),
                        'yes' if rect_norm else 'no'
                    )
                except Exception:
                    pass

                t_generate0 = time.perf_counter()
                img_bytes = self._generate_snapshot_png(rect_norm=rect_norm, fmt=fmt, quality=quality, max_w=max_w, max_h=max_h)
                t_generate1 = time.perf_counter()

                if img_bytes:
                    headers = Headers()
                    headers["Content-Type"] = ("image/png" if fmt == 'png' else "image/jpeg")
                    headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
                    try:
                        ext = ("png" if fmt == 'png' else "jpg")
                        fname = f"snap-{time.strftime('%Y%m%d-%H%M%S')}.{ext}"
                        headers["Content-Disposition"] = f"attachment; filename=\"{fname}\""
                    except Exception:
                        pass
                    try:
                        t_ready = time.perf_counter()
                        logging.info(
                            "[snapshot.server] total=%.2fms generate=%.2fms bytes=%.1fKB fmt=%s rect=%s",
                            (t_ready - t_req0) * 1000.0,
                            (t_generate1 - t_generate0) * 1000.0,
                            len(img_bytes) / 1024.0,
                            fmt,
                            'yes' if rect_norm else 'no'
                        )
                    except Exception:
                        pass
                    return WSResponse(
                        status_code=int(http.HTTPStatus.OK),
                        reason_phrase=http.HTTPStatus.OK.phrase,
                        headers=headers,
                        body=img_bytes,
                    )
                headers = Headers()
                headers["Content-Type"] = "text/plain; charset=utf-8"
                return WSResponse(
                    status_code=int(http.HTTPStatus.INTERNAL_SERVER_ERROR),
                    reason_phrase=http.HTTPStatus.INTERNAL_SERVER_ERROR.phrase,
                    headers=headers,
                    body=b"Failed to capture snapshot",
                )
            except Exception:
                logging.error("Error generating snapshot", exc_info=True)
                headers = Headers()
                headers["Content-Type"] = "text/plain; charset=utf-8"
                return WSResponse(
                    status_code=int(http.HTTPStatus.INTERNAL_SERVER_ERROR),
                    reason_phrase=http.HTTPStatus.INTERNAL_SERVER_ERROR.phrase,
                    headers=headers,
                    body=b"Error generating snapshot",
                )
        else:
            headers = Headers()
            headers["Content-Type"] = "text/plain; charset=utf-8"
            return WSResponse(
                status_code=int(http.HTTPStatus.NOT_FOUND),
                reason_phrase=http.HTTPStatus.NOT_FOUND.phrase,
                headers=headers,
                body=b"Not Found",
            )

    def _public_url_payload(self):
        try:
            current_url = self.tunnel_manager.get_current_url() if self.tunnel_manager else None
            if current_url:
                current_port = self.tunnel_manager.current_port
                return {
                    'success': True, 
                    'url': current_url,
                    'port': current_port,
                    'message': f'Current public URL for port {current_port}',
                    'email_status': (self.tunnel_manager.last_email_message if self.tunnel_manager else None),
                }
            return {
                'success': False,
                'error': 'No public URL available'
            }
        except Exception as e:
            return {
                'success': False, 
                'error': f'Error getting URL: {str(e)}'
            }

    async def _refresh_tunnel_payload(self):
        if not self.tunnel_manager:
            print(" No tunnel manager available")
            return {"success": False, "error": "No tunnel manager available"}

        print(" Refreshing tunnel...")
        loop = asyncio.get_running_loop()
        url = await loop.run_in_executor(None, self.tunnel_manager.refresh_tunnel)
        if not url:
            print(" Failed to restart tunnel")
            return {"success": False, "error": "Failed to refresh tunnel"}

        current_port = self.tunnel_manager.current_port
        print(f" New tunnel URL: {url}")
        return {
            "success": True,
            "url": url,
            "port": current_port,
            "message": f"Successfully refreshed tunnel on port {current_port}",
            "email_status": self.tunnel_manager.last_email_message,
        }

    async def handle_get_public_url(self):
        """API endpoint to get current public URL"""
        return json.dumps(self._public_url_payload())

    async def handle_refresh_tunnel(self):
        """API endpoint to refresh tunnel and get new URL."""
        try:
            return json.dumps(await self._refresh_tunnel_payload())
        except Exception as e:
            return json.dumps({"success": False, "error": str(e)})

    async def handle_set_quality(self, path=None):
        """API endpoint to set graphics quality via query string."""
        try:
            parsed = urllib.parse.urlparse(str(path or ""))
            qs = urllib.parse.parse_qs(parsed.query or "")
            raw = (qs.get("value") or qs.get("quality") or [self.current_quality])[0]
            value = self._apply_quality(raw, self.current_quality)
            return {"success": True, "quality": value}
        except Exception as e:
            return {"success": False, "error": str(e)}

    async def handle_set_fps(self, path=None):
        """API endpoint to set FPS via query string."""
        try:
            parsed = urllib.parse.urlparse(str(path or ""))
            qs = urllib.parse.parse_qs(parsed.query or "")
            raw = (qs.get("value") or qs.get("fps") or [self.current_fps])[0]
            value = self._apply_fps(raw, self.current_fps)
            return {"success": True, "fps": value}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def _capture_screen_image(self, rect_norm=None):
        """One-shot snapshot; reuse instances and capture at source (dxcam region  MSS region)."""
        t0 = time.perf_counter()
        # Track which backend we actually used for logging
        self._last_snapshot_backend = 'none'
        screen_size = self._primary_screen_size_or_none("snapshot.rect_screen_size") if rect_norm else None
        # Helper: convert normalized rect to absolute desktop rect
        def _norm_to_abs_rect():
            if not rect_norm or not screen_size:
                return None
            try:
                sw, sh = screen_size
                x0 = max(0, min(sw, int(float(rect_norm['x0']) * sw)))
                y0 = max(0, min(sh, int(float(rect_norm['y0']) * sh)))
                x1 = max(0, min(sw, int(float(rect_norm['x1']) * sw)))
                y1 = max(0, min(sh, int(float(rect_norm['y1']) * sh)))
                if x1 <= x0 or y1 <= y0:
                    return None
                return (x0, y0, x1, y1)
            except Exception:
                return None

        abs_rect = _norm_to_abs_rect()

        # Heuristic: for small ROIs, MSS region can be faster than dxcam
        prefer_mss_first = False
        if abs_rect:
            try:
                sw2, sh2 = screen_size or (0, 0)
                if sw2 > 0 and sh2 > 0:
                    l, t, r, b = abs_rect
                    area_ratio = ((r - l) * (b - t)) / float(sw2 * sh2)
                    prefer_mss_first = area_ratio <= 0.25
            except Exception:
                prefer_mss_first = False

        def _try_dxcam_then_none():
            try:
                cam = getattr(self, 'dxcam_camera', None)
                if cam is not None and getattr(self, "dxcam_started", False):
                    _log_fallback("snapshot.capture_backend", "mss_or_other", "dxcam_ring_buffer_active")
                    return None
                if cam is not None:
                    arr = cam.grab(region=abs_rect)
                    if arr is not None:
                        rgb = arr[:, :, :3][:, :, ::-1].copy(order='C')
                        try:
                            self._last_snapshot_backend = 'dxcam'
                        except Exception:
                            pass
                        # Pillow deprecation: mode parameter on fromarray will be removed in Pillow 13
                        try:
                            return Image.fromarray(rgb)
                        except Exception as exc:
                            _log_fallback("snapshot.dxcam.image", "Image.frombuffer", "Image.fromarray_failed", exc)
                            return Image.frombuffer('RGB', (rgb.shape[1], rgb.shape[0]), rgb.tobytes())
            except Exception:
                _log_fallback("snapshot.capture_backend", "mss_or_other", "dxcam_failed")
                logging.warning('Snapshot dxcam failed; falling back', exc_info=True)
            return None

        def _try_mss_then_none():
            if HAS_MSS:
                try:
                    if not hasattr(self, '_snapshot_sct') or self._snapshot_sct is None:
                        self._snapshot_sct = mss.mss()
                    sct = self._snapshot_sct
                    mon = sct.monitors[1] if len(sct.monitors) > 1 else sct.monitors[0]
                    rect = None
                    if abs_rect:
                        l, t, r, b = abs_rect
                        rect = {'left': l, 'top': t, 'width': max(1, r - l), 'height': max(1, b - t)}
                    sct_img = sct.grab(rect or mon)
                    h, w = sct_img.height, sct_img.width
                    arr = np.frombuffer(sct_img.bgra, dtype=np.uint8).reshape((h, w, 4))
                    rgb = np.ascontiguousarray(arr[..., :3][:, :, ::-1])
                    try:
                        self._last_snapshot_backend = 'mss'
                    except Exception:
                        pass
                    try:
                        return Image.fromarray(rgb)
                    except Exception as exc:
                        _log_fallback("snapshot.mss.image", "Image.frombuffer", "Image.fromarray_failed", exc)
                        return Image.frombuffer('RGB', (rgb.shape[1], rgb.shape[0]), rgb.tobytes())
                except Exception:
                    _log_fallback("snapshot.capture_backend", "next_backend", "mss_failed")
                    logging.warning('Snapshot MSS failed', exc_info=True)
            return None

        # Order backends based on ROI size
        if prefer_mss_first:
            img = _try_mss_then_none()
            if img is not None:
                return img
            img = _try_dxcam_then_none()
            if img is not None:
                return img
        else:
            img = _try_dxcam_then_none()
            if img is not None:
                return img
            img = _try_mss_then_none()
            if img is not None:
                return img

        # 2) fast_ctypes (full, then crop)
        if HAS_FAST_CTYPES:
            try:
                sct = getattr(self, 'fast_ctypes_capture', None)
                if sct is None:
                    self.fast_ctypes_capture = fast_ctypes_screenshots.ScreenshotOfAllMonitors()
                    sct = self.fast_ctypes_capture
                frame = sct.screenshot_monitors()
                if frame is not None:
                    img = Image.fromarray(frame)
                    if abs_rect:
                        img = img.crop(abs_rect)
                    return img
            except Exception:
                _log_fallback("snapshot.capture_backend", "none", "fast_ctypes_failed")
                logging.warning('Snapshot fast_ctypes failed; falling back', exc_info=True)

        return None

    def _generate_snapshot_png(self, rect_norm=None, fmt='png', quality=85, max_w=None, max_h=None):
        """
        Returns image bytes (PNG or JPEG) for the current (or region) snapshot.
        Supports format/quality/downscale and logs captureconvertencode timings.
        """
        try:
            t0 = time.perf_counter()
            img = self._capture_screen_image(rect_norm=rect_norm)
            t1 = time.perf_counter()
            if img is None:
                return None
            # Optional downscale to limit size
            try:
                if (max_w and img.width > max_w) or (max_h and img.height > max_h):
                    # Keep aspect ratio
                    tw = img.width; th = img.height
                    rw = (max_w / tw) if max_w else 1.0
                    rh = (max_h / th) if max_h else 1.0
                    r = min(rw, rh)
                    if r < 1.0:
                        new_size = (max(1, int(tw * r)), max(1, int(th * r)))
                        img = img.resize(new_size, Image.BILINEAR)
            except Exception:
                pass
            arr = np.array(img, copy=False)
            if arr.ndim == 3 and arr.shape[2] == 4:
                arr = arr[:, :, :3]
            t2 = time.perf_counter()
            out = None
            if fmt == 'png' and HAS_IMAGECODECS:
                try:
                    out = imagecodecs.png_encode(arr, level=0)
                except Exception as exc:
                    _log_fallback("snapshot.png_encoder", "pillow_png", "imagecodecs_png_failed", exc)
                    out = None
            if out is None:
                buf = io.BytesIO()
                try:
                    if fmt == 'png':
                        img.save(buf, format='PNG', optimize=False, compress_level=0)
                    else:
                        # JPEG fast path via imagecodecs not used here to keep code simple; Pillow is fine
                        img.save(buf, format='JPEG', quality=max(1, min(95, int(quality))))
                    out = buf.getvalue()
                except Exception:
                    out = None
            t3 = time.perf_counter()
            logging.info(
                "[snapshot.timing] total=%.2fms | capture=%.2fms convert=%.2fms encode=%.2fms | size=%dx%d bytes=%.1fKB backend=%s fmt=%s rect=%s",
                (t3 - t0) * 1000.0,
                (t1 - t0) * 1000.0,
                (t2 - t1) * 1000.0,
                (t3 - t2) * 1000.0,
                getattr(img, 'width', 0), getattr(img, 'height', 0),
                0.0 if not out else (len(out) / 1024.0),
                getattr(self, '_last_snapshot_backend', 'unknown'),
                fmt,
                'yes' if rect_norm else 'no'
            )
            return out
        except Exception:
            logging.error("Failed to generate snapshot PNG", exc_info=True)
            return None
