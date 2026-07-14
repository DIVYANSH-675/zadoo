"""HTTP route and snapshot handlers."""
from __future__ import annotations

import asyncio
import http
import ipaddress
import json
import logging
import math
import os
import secrets
import time
import urllib.error
import urllib.parse
import urllib.request

import mss
import numpy as np
import sounddevice as sd
import websockets
from imagecodecs._jpeg8 import jpeg8_encode
from imagecodecs._png import png_encode
from PIL import Image
from websockets.datastructures import Headers
from websockets.http11 import Response as WSResponse

from .assets import load_binary, load_static, load_template
from .camera_discovery import enumerate_camera_devices
from .dpi import get_primary_screen_size
from .logging_utils import _log_except
from .saas import ZadooCloudClient
from .settings import ACCESS_CODE_MAX_LENGTH, PERMISSION_KEYS, _clean_http_origin


class RoutesMixin:
    SECURITY_HEADERS = {
        "Content-Security-Policy": "base-uri 'none'; object-src 'none'; frame-ancestors 'self'",
        "Cross-Origin-Resource-Policy": "same-origin",
        "Referrer-Policy": "no-referrer",
        "X-Content-Type-Options": "nosniff",
        "X-Frame-Options": "SAMEORIGIN",
    }
    AUTH_COOKIE_NAME = "zadoo_auth"
    CSRF_COOKIE_NAME = "zadoo_csrf"
    AUTH_TTL_SECONDS = 3600
    AUTH_MAX_IDENTITIES = 4096
    AUTH_MAX_SESSIONS = 1024
    AUTH_CODE_ENV = "ZADOO_ACCESS_CODE"
    VIEW_ACTIONS = {
        "client_stream_stats",
        "cursor_broadcast",
        "get_capture_stats",
        "get_public_url",
        "get_stream_status",
        "stream_ping",
    }
    MOUSE_ACTIONS = {"click", "drag", "move", "scroll"}
    KEYBOARD_ACTIONS = {"key", "type_text"}
    ADVANCED_ACTIONS = {"set_performance", "set_quality"}
    ROUTE_FEATURES = {
        "/": "public",
        "/api/auth": "public",
        "/api/runtime/status": "settings",
        "/api/runtime/stop": "settings",
        "/api/runtime/refresh-tunnel": "settings",
        "/brand-header.png": "public",
        "/trigger-icon.png": "public",
        "/splash.png": "public",
        "/video": "view",
        "/audio": "system_audio",
        "/input": "view",
        "/benchmark.html": "view",
        "/snapshot": "snapshots",
        "/terminal.html": "terminal",
        "/terminal": "terminal",
        "/webcam": "camera",
        "/api/list-cameras": "camera",
        "/mic": "mic",
        "/api/list-mics": "mic",
        "/host-controls": "remote_alerts",
        "/api/alert": "remote_alerts",
        "/api/local/credits": "billing",
        "/api/local/payment-market": "billing",
        "/api/local/topup-order": "billing",
        "/api/local/topup-verify": "billing",
    }
    STATIC_ASSETS = {
        "/static/vendor/codemirror-5.65.21.min.css": (
            "vendor/codemirror-5.65.21.min.css",
            "text/css; charset=utf-8",
        ),
        "/static/vendor/codemirror-5.65.21.min.js": (
            "vendor/codemirror-5.65.21.min.js",
            "text/javascript; charset=utf-8",
        ),
        "/static/vendor/xterm-5.3.0.min.css": (
            "vendor/xterm-5.3.0.min.css",
            "text/css; charset=utf-8",
        ),
        "/static/vendor/xterm-5.3.0-fit-0.8.0.min.js": (
            "vendor/xterm-5.3.0-fit-0.8.0.min.js",
            "text/javascript; charset=utf-8",
        ),
    }
    PAYMENT_MARKETS = {
        "INDIA": {
            "market": "INDIA",
            "currency": "INR",
            "symbol": "₹",
            "minimum": 200,
            "maximum": 10_000,
            "slider_maximum": 5_000,
            "default": 500,
            "step": 50,
            "minor_per_minute": 100,
        },
        "GLOBAL": {
            "market": "GLOBAL",
            "currency": "USD",
            "symbol": "$",
            "minimum": 3,
            "maximum": 100,
            "slider_maximum": 50,
            "default": 5,
            "step": 1,
            "minor_per_minute": 3,
        },
    }
    CSRF_HTTP_FEATURES = {
        "advanced_video",
        "mouse",
        "keyboard",
        "clipboard_pull",
        "clipboard_push",
        "tunnel_refresh",
        "remote_alerts",
        "settings",
    }

    def _response(self, body, content_type, status=http.HTTPStatus.OK, extra_headers=None):
        headers = Headers()
        headers["Content-Type"] = content_type
        extra_headers = extra_headers or {}
        override_names = {key.lower() for key in extra_headers}
        for key, value in self.SECURITY_HEADERS.items():
            if key.lower() not in override_names:
                headers[key] = value
        for key, value in extra_headers.items():
            if isinstance(value, (list, tuple)):
                for item in value:
                    headers[key] = str(item)
            else:
                headers[key] = str(value)
        return WSResponse(
            status_code=int(status),
            reason_phrase=status.phrase,
            headers=headers,
            body=body,
        )

    def _json_response(self, payload, status=http.HTTPStatus.OK, extra_headers=None):
        headers = {"Cache-Control": "no-store"}
        headers.update(extra_headers or {})
        return self._response(
            json.dumps(payload).encode("utf-8"),
            "application/json; charset=utf-8",
            status,
            headers,
        )

    def _plain_response(self, body, status=http.HTTPStatus.FORBIDDEN):
        return self._response(
            str(body).encode("utf-8"),
            "text/plain; charset=utf-8",
            status,
            {"Cache-Control": "no-store"},
        )

    def _template_response(self, name, extra_headers=None, strip=False):
        template = load_template(name)
        if strip:
            template = template.strip()
        return self._response(template.encode("utf-8"), "text/html; charset=utf-8", extra_headers=extra_headers)

    def _static_response(self, name, content_type):
        return self._response(
            load_static(name),
            content_type,
            extra_headers={
                "Cache-Control": "public, max-age=31536000, immutable",
                "X-Content-Type-Options": "nosniff",
            },
        )

    def _png_file_response(self, name, missing_message):
        try:
            body = load_binary(name)
            return self._response(
                body,
                "image/png",
                extra_headers={"Cache-Control": "public, max-age=31536000, immutable"},
            )
        except FileNotFoundError:
            return self._plain_response(missing_message, http.HTTPStatus.NOT_FOUND)
        except OSError as exc:
            return self._plain_response(
                f"Asset read failed: {name}: {exc}",
                http.HTTPStatus.INTERNAL_SERVER_ERROR,
            )

    def _header_get(self, request_headers, name, default=None):
        return request_headers.get(name, default) if request_headers else default

    def _request_identity(self, request_headers):
        for name in ("CF-Connecting-IP", "X-Real-IP", "X-Forwarded-For"):
            raw = self._header_get(request_headers, name, "")
            if raw:
                return str(raw).split(",", 1)[0].strip() or "default"
        return self._header_get(request_headers, "Host", "default") or "default"

    def _auth_lockout_remaining(self, request_headers):
        key = self._request_identity(request_headers)
        state = self._auth_failures.get(key)
        if state is None:
            return 0
        remaining = state["locked_until"] - time.time()
        return max(0, math.ceil(remaining))

    def _record_auth_failure(self, request_headers):
        key = self._request_identity(request_headers)
        failures = self._auth_failures
        if key not in failures and len(failures) >= self.AUTH_MAX_IDENTITIES:
            failures.pop(next(iter(failures)))
        now = time.time()
        state = failures.setdefault(key, {"times": [], "locked_until": 0.0})
        times = [
            timestamp
            for timestamp in state["times"]
            if now - timestamp <= self._auth_window_seconds
        ]
        times.append(now)
        state["times"] = times
        if len(times) >= self._auth_max_failures:
            state["locked_until"] = now + self._auth_lockout_seconds
            return self._auth_lockout_seconds
        return 0

    def _record_auth_success(self, request_headers):
        self._auth_failures.pop(self._request_identity(request_headers), None)

    def _request_origin_allowed(self, request_headers):
        origin = self._header_get(request_headers, "Origin", "")
        if not origin:
            return True
        allowed = self._allowed_origins
        if "*" in allowed:
            return True
        try:
            normalized_origin = _clean_http_origin(origin, "Origin").lower()
        except ValueError:
            return False
        if normalized_origin in allowed:
            return True
        host = str(self._header_get(request_headers, "Host", "") or "").lower()
        return bool(host and urllib.parse.urlparse(normalized_origin).netloc == host)

    def _is_local_request(self, connection):
        remote = getattr(connection, "remote_address", None)
        host = remote[0] if isinstance(remote, tuple) and remote else ""
        try:
            address = ipaddress.ip_address(str(host).split("%", 1)[0])
        except ValueError:
            return False
        return address.is_loopback or bool(address.version == 6 and address.ipv4_mapped and address.ipv4_mapped.is_loopback)

    def _is_tunnel_request(self, request_headers):
        """True if the request arrived through the Cloudflare tunnel (public link).
        Cloudflare adds CF-Connecting-IP / CF-Ray; direct localhost/LAN access has neither."""
        if self._header_get(request_headers, "CF-Connecting-IP", "") or self._header_get(request_headers, "CF-Ray", ""):
            return True
        xff = str(self._header_get(request_headers, "X-Forwarded-For", "") or "").strip()
        return bool(xff and not xff.startswith(("127.", "::1", "localhost")))

    def _is_localhost_request(self, request_headers):
        """True if the request targets this machine itself (localhost). The host accessing its
        own server is safe, so it is allowed without the public link — this lets the owner open
        http://localhost:6173 to test directly. Remote LAN access stays blocked."""
        host = str(self._header_get(request_headers, "Host", "") or "").strip().lower()
        if not host:
            return False
        hostname = host.rsplit(":", 1)[0].strip("[]")
        return hostname in ("localhost", "127.0.0.1", "::1")

    def _public_only_response(self):
        body = (
            "<!doctype html><html><head><meta charset='utf-8'><title>Zadoo</title>"
            "<style>body{font-family:Segoe UI,system-ui,sans-serif;background:#0f1115;color:#e8eef5;"
            "display:flex;min-height:100vh;align-items:center;justify-content:center;margin:0;text-align:center}"
            "div{max-width:420px;padding:28px}h1{font-size:20px;margin:0 0 10px}p{color:#9aa6b5;line-height:1.6}</style>"
            "</head><body><div><h1>Open via your public link</h1>"
            "<p>This machine does not accept direct LAN access. "
            "Use the public Zadoo link shown in the Settings window.</p></div></body></html>"
        )
        headers = Headers()
        headers["Content-Type"] = "text/html; charset=utf-8"
        headers["Cache-Control"] = "no-store"
        return WSResponse(
            status_code=int(http.HTTPStatus.FORBIDDEN),
            reason_phrase=http.HTTPStatus.FORBIDDEN.phrase,
            headers=headers,
            body=body.encode("utf-8"),
        )

    def _forbidden_response(self, route_path, message="Forbidden", public_only=False):
        if str(route_path or "").startswith("/api/"):
            return self._json_response(
                {"success": False, "error": message},
                http.HTTPStatus.FORBIDDEN,
            )
        return self._public_only_response() if public_only else self._plain_response("Forbidden")

    def _state_changing_http_allowed(self, request_headers):
        header_token = str(self._header_get(request_headers, "X-Zadoo-CSRF", "") or "")
        if not header_token:
            return False
        session = self._session_for_headers(request_headers)
        if session is None:
            return False
        return secrets.compare_digest(header_token, session["csrf"])

    def _headers_for_websocket(self, websocket):
        request = getattr(websocket, "request", None)
        return request.headers if request else None

    def _feature_for_action(self, action):
        action = str(action or "")
        if action in self.VIEW_ACTIONS:
            return "view"
        if action in self.MOUSE_ACTIONS:
            return "mouse"
        if action in self.KEYBOARD_ACTIONS:
            return "keyboard"
        if action == "get_clipboard":
            return "clipboard_pull"
        if action in {"set_clipboard", "set_clipboard_image"}:
            return "clipboard_push"
        if action in self.ADVANCED_ACTIONS:
            return "advanced_video"
        if action == "refresh_tunnel":
            return "tunnel_refresh"
        return None

    def _is_ws_action_authorized_for_headers(self, request_headers, action):
        feature = self._feature_for_action(action)
        if feature is None:
            return False
        return self._feature_allowed_for_headers(request_headers, feature)

    def _is_ws_action_authorized(self, websocket, action):
        # During the post-zero grace lock (≈ -5 to -10 min), only screen-view actions are
        # allowed — mouse / keyboard / clipboard / everything else is disabled.
        request_headers = self._headers_for_websocket(websocket)
        session_token = self._cookie_value(request_headers, self.AUTH_COOKIE_NAME)
        if self._grace_locks_by_session.get(session_token) and str(action) not in self.VIEW_ACTIONS:
            return False
        return self._is_ws_action_authorized_for_headers(request_headers, action)

    def _set_grace_lock(self, websocket, locked):
        token = self._cookie_value(self._headers_for_websocket(websocket), self.AUTH_COOKIE_NAME)
        if not token:
            raise RuntimeError("Authenticated session cookie is missing")
        if locked:
            self._grace_locks_by_session.setdefault(token, set()).add(websocket)
            return
        locks = self._grace_locks_by_session.get(token)
        if locks:
            locks.discard(websocket)
            if not locks:
                del self._grace_locks_by_session[token]

    async def _send_ws_forbidden(self, websocket, action):
        try:
            await websocket.send(json.dumps({
                "type": "error",
                "error": "Forbidden",
                "action": action,
            }))
        except websockets.exceptions.ConnectionClosed:
            return

    def _feature_allowed_for_headers(self, request_headers, feature, connection=None):
        if feature == "public":
            return True
        if feature == "settings":
            return self._is_local_request(connection)
        session = self._session_for_headers(request_headers)
        if session is None:
            return False
        permissions = session["permissions"]
        if feature in {"billing", "view"}:
            return True
        return permissions.get(feature) is True

    def _cookie_value(self, request_headers, name):
        cookie_header = self._header_get(request_headers, "Cookie", "")
        if not isinstance(cookie_header, str) or not cookie_header:
            return None
        for item in cookie_header.split(";"):
            key, separator, value = item.strip().partition("=")
            if separator and key == name:
                return value
        return None

    def _auth_code(self):
        configured = str(os.getenv(self.AUTH_CODE_ENV) or "").strip()
        if not configured:
            raise RuntimeError(
                f"{self.AUTH_CODE_ENV} is required when installed Settings are not configured"
            )
        if len(configured) > ACCESS_CODE_MAX_LENGTH:
            raise RuntimeError(
                f"{self.AUTH_CODE_ENV} must be {ACCESS_CODE_MAX_LENGTH} characters or fewer"
            )
        return configured

    def _announce_auth_codes(self):
        if self.settings_store.configured():
            print(" Access code source: installed settings")
            return
        self._auth_code()
        print(" Access code source: environment")

    def _match_auth_code(self, code):
        submitted = str(code or "").strip()
        if not submitted:
            return False
        installed = self.settings_store.get_access_code()
        return secrets.compare_digest(submitted, installed or self._auth_code())

    def _cleanup_auth_sessions(self):
        sessions = self.auth_sessions
        now = time.time()
        expired = [
            token
            for token, session in sessions.items()
            if session["expires_at"] <= now
        ]
        for token in expired:
            del sessions[token]
            self._grace_locks_by_session.pop(token, None)

    def _session_for_headers(self, request_headers):
        token = self._cookie_value(request_headers, self.AUTH_COOKIE_NAME)
        if not token:
            return None
        session = self.auth_sessions.get(token)
        if session is None:
            return None
        if session["expires_at"] <= time.time():
            del self.auth_sessions[token]
            self._grace_locks_by_session.pop(token, None)
            return None
        return session

    def _feature_for_route(self, route_path):
        if route_path in self.STATIC_ASSETS:
            return "public"
        if route_path.startswith("/api/settings/"):
            return self.ROUTE_FEATURES.get(route_path, "settings")
        return self.ROUTE_FEATURES.get(route_path)

    def _http_route_requires_csrf(self, route_path, feature=None):
        if feature is None:
            feature = self._feature_for_route(route_path)
        if feature == "settings":
            return False
        return route_path.startswith("/api/") and feature in self.CSRF_HTTP_FEATURES

    def _is_ws_authorized(self, route_path, request_headers):
        feature = self._feature_for_route(route_path)
        if feature is None:
            return False
        return self._feature_allowed_for_headers(request_headers, feature)

    async def handle_auth(self, path, request_headers=None):
        lockout_remaining = self._auth_lockout_remaining(request_headers)
        if lockout_remaining > 0:
            return self._json_response(
                {"success": False, "error": "Too many invalid attempts", "retry_after": lockout_remaining},
                http.HTTPStatus.TOO_MANY_REQUESTS,
            )
        parsed = urllib.parse.urlparse(str(path or ""))
        if parsed.query or self._header_get(request_headers, "X-Zadoo-Access", ""):
            return self._json_response(
                {"success": False, "error": "Authentication access selectors are not supported"},
                http.HTTPStatus.BAD_REQUEST,
            )
        code = str(self._header_get(request_headers, "X-Zadoo-Code", "") or "")
        if not self._match_auth_code(code):
            if code:
                retry_after = self._record_auth_failure(request_headers)
                if retry_after:
                    return self._json_response(
                        {"success": False, "error": "Too many invalid attempts", "retry_after": retry_after},
                        http.HTTPStatus.TOO_MANY_REQUESTS,
                    )
            return self._json_response({"success": False, "error": "Invalid code"}, http.HTTPStatus.UNAUTHORIZED)
        settings = self.settings_store.load(reload=True)
        permissions = (
            dict(settings["permissions"])
            if settings["setup_complete"] or settings["device_token"]
            else dict.fromkeys(PERMISSION_KEYS, True)
        )

        self._record_auth_success(request_headers)
        token = secrets.token_urlsafe(32)
        expires_at = time.time() + self.AUTH_TTL_SECONDS
        self._cleanup_auth_sessions()
        sessions = self.auth_sessions
        csrf_token = secrets.token_urlsafe(32)
        session_payload = {
            "expires_at": expires_at,
            "csrf": csrf_token,
            "permissions": permissions,
        }
        sessions[token] = session_payload
        while len(sessions) > self.AUTH_MAX_SESSIONS:
            evicted = next(iter(sessions))
            del sessions[evicted]
            self._grace_locks_by_session.pop(evicted, None)
        # Mark cookies Secure when served over the HTTPS tunnel (but NOT for plain
        # http://localhost, where a Secure cookie would be rejected by the browser).
        xfproto = str(self._header_get(request_headers, "X-Forwarded-Proto", "") or "").lower()
        secure = "Secure; " if (not self._is_localhost_request(request_headers) or xfproto == "https") else ""
        auth_cookie = (
            f"{self.AUTH_COOKIE_NAME}={token}; Path=/; Max-Age={self.AUTH_TTL_SECONDS}; "
            f"{secure}HttpOnly; SameSite=Lax"
        )
        csrf_cookie = (
            f"{self.CSRF_COOKIE_NAME}={csrf_token}; Path=/; Max-Age={self.AUTH_TTL_SECONDS}; "
            f"{secure}SameSite=Lax"
        )
        payload = {
            "success": True,
            "expires_in": self.AUTH_TTL_SECONDS,
            "csrf_token": csrf_token,
            "permissions": permissions,
            "limits": {"clipboard_image_max_bytes": self._clipboard_image_max_bytes},
        }
        return self._json_response(
            payload,
            extra_headers={"Set-Cookie": [auth_cookie, csrf_cookie]},
        )

    def enumerate_cameras(self):
        return enumerate_camera_devices()

    def enumerate_microphones(self):
        try:
            devices = list(sd.query_devices())
            hostapis = list(sd.query_hostapis())
            default_index = int(sd.default.device[0])
        except Exception as exc:
            raise RuntimeError(f"Microphone enumeration failed: {exc}") from exc
        default_input = default_index if default_index >= 0 else None

        result = []
        seen = set()
        for index, device in enumerate(devices):
            try:
                max_input = int(device["max_input_channels"])
            except (KeyError, TypeError, ValueError, OverflowError) as exc:
                raise RuntimeError(f"Microphone {index} has invalid max_input_channels: {exc}") from exc
            if max_input < 0:
                raise RuntimeError(f"Microphone {index} max_input_channels cannot be negative")
            if max_input == 0:
                continue
            try:
                raw_name = device["name"]
            except KeyError as exc:
                raise RuntimeError(f"Microphone {index} is missing name") from exc
            if not isinstance(raw_name, str) or not raw_name.strip():
                raise RuntimeError(f"Microphone {index} has no name")
            name = raw_name.strip()
            try:
                hostapi_index = int(device["hostapi"])
                raw_api = hostapis[hostapi_index]["name"]
            except (IndexError, KeyError, TypeError, ValueError, OverflowError) as exc:
                raise RuntimeError(f"Microphone {index} has invalid host API metadata: {exc}") from exc
            if not isinstance(raw_api, str) or not raw_api.strip():
                raise RuntimeError(f"Microphone {index} host API has no name")
            api = raw_api.strip()
            label = f"{name} ({api})"
            key = (name.casefold(), api.casefold(), max_input)
            if key in seen:
                label = f"{label} #{index}"
            seen.add(key)
            try:
                samplerate = int(float(device["default_samplerate"]))
            except (KeyError, TypeError, ValueError, OverflowError) as exc:
                raise RuntimeError(f"Microphone {index} has invalid default sample rate: {exc}") from exc
            if samplerate <= 0:
                raise RuntimeError(f"Microphone {index} has no default sample rate")
            result.append({
                "id": str(index),
                "device_index": index,
                "name": name,
                "label": label,
                "default": default_input is not None and int(index) == int(default_input),
                "channels": max_input,
                "samplerate": samplerate,
            })
        if default_input is not None:
            default_device = next((item for item in result if item["device_index"] == default_input), None)
            if default_device is None:
                raise RuntimeError(f"Default microphone index {default_input} is not an input device")
            result.insert(0, {
                "id": "default",
                "device_index": None,
                "name": default_device["name"],
                "label": f"System Default ({default_device['name']})",
                "default": True,
                "channels": default_device["channels"],
                "samplerate": default_device["samplerate"],
            })
        return result

    async def _device_list_response(self, enumerate_devices):
        try:
            devices = await asyncio.get_running_loop().run_in_executor(None, enumerate_devices)
            return self._json_response({"success": True, "devices": devices})
        except Exception as exc:
            return self._json_response(
                {"success": False, "error": str(exc)},
                http.HTTPStatus.INTERNAL_SERVER_ERROR,
            )

    def _runtime_admin_code_valid(self, request_headers):
        code = str(self._header_get(request_headers, "X-Zadoo-Code", "") or "")
        return self._match_auth_code(code)

    def _apply_quality(self, raw_value):
        if isinstance(raw_value, bool) or not isinstance(raw_value, int):
            raise ValueError("JPEG quality must be an integer")
        value = raw_value
        if not 10 <= value <= 100:
            raise ValueError("JPEG quality must be between 10 and 100")
        self.current_quality = value
        self._quality_locked_by_user = True
        if self.screen_capturer:
            self.screen_capturer.quality = value
        self._apply_stream_profile("quality_changed")
        return value

    async def _proxy_cloud(self, method: str, cloud_path: str, request_headers=None, request_body=None):
        try:
            if method not in {"GET", "POST"}:
                raise ValueError(f"Unsupported cloud proxy method: {method}")
            if method == "GET" and request_body is not None:
                raise ValueError("Cloud GET body must be omitted")
            store = self.settings_store
            token = store.get_device_token()
            if not token:
                return self._json_response({"success": False, "error": "Device not signed in"}, http.HTTPStatus.UNAUTHORIZED)
            if method == "POST" and not isinstance(request_body, (bytes, bytearray)):
                raise TypeError("Cloud POST body must be bytes")
            url = store.load()["cloud_api_base"] + cloud_path
            headers_out = {
                "Authorization": f"Bearer {token}",
                "Accept": "application/json",
            }
            body = bytes(request_body) if request_body is not None else None
            if body is not None:
                headers_out["Content-Type"] = "application/json"
            for hdr in (
                "CF-Connecting-IP",
                "X-Forwarded-For",
                "CF-IPCountry",
                "X-Country-Code",
                "X-Vercel-IP-Country",
            ):
                val = self._header_get(request_headers, hdr, None) if request_headers else None
                if val:
                    headers_out[hdr] = str(val)

            def _do_request():
                request = urllib.request.Request(
                    url,
                    data=body,
                    headers=headers_out,
                    method=method,
                )
                try:
                    with urllib.request.urlopen(request, timeout=10 if body is not None else 8) as response:
                        status = int(response.getcode())
                        return status, ZadooCloudClient._read_response_body(response, cloud_path, status), False
                except urllib.error.HTTPError as exc:
                    status = int(exc.code)
                    return status, ZadooCloudClient._read_response_body(exc, cloud_path, status), True

            status, raw, http_error = await asyncio.to_thread(_do_request)
            data = ZadooCloudClient._response_object(raw, cloud_path, status)
            data = ZadooCloudClient._validated_result(data, cloud_path, status, http_error=http_error)
            http_status = http.HTTPStatus(status) if status in http.HTTPStatus._value2member_map_ else http.HTTPStatus.BAD_GATEWAY
            return self._json_response(data, http_status)
        except Exception as exc:
            return self._json_response({"success": False, "error": str(exc)}, http.HTTPStatus.BAD_GATEWAY)

    def _payment_market_payload(self, request_headers):
        country = next(
            (
                str(value).strip().upper()
                for name in ("CF-IPCountry", "X-Country-Code", "X-Vercel-IP-Country")
                if (value := self._header_get(request_headers, name, ""))
            ),
            "",
        )
        return self.PAYMENT_MARKETS["INDIA" if country == "IN" else "GLOBAL"]

    async def process_request(self, connection, request):
        """Process a websockets 15 HTTP request."""
        path = request.path
        request_headers = request.headers
        route_path = urllib.parse.urlparse(path).path

        logging.debug("process_request: path=%s", path)

        if not self._request_origin_allowed(request_headers):
            return self._forbidden_response(route_path, "Forbidden origin")

        # Reachable through the Cloudflare tunnel (public link) AND from localhost on the host
        # itself (so the owner can open localhost:6173 to test). Direct LAN access from another
        # machine (ip:6173) is still refused unless ZADOO_ALLOW_DIRECT_ACCESS=1. The local admin
        # API used by the Settings window is always allowed.
        is_admin_route = route_path.startswith(("/api/runtime/", "/api/settings/"))
        local_peer = self._is_local_request(connection)
        local_admin = (
            local_peer
            and self._is_localhost_request(request_headers)
            and not self._is_tunnel_request(request_headers)
        )
        if is_admin_route and not local_admin:
            return self._forbidden_response(route_path, "Local request required")
        if (not is_admin_route and not self._allow_direct_access
                and not (local_peer and (
                    self._is_tunnel_request(request_headers) or self._is_localhost_request(request_headers)
                ))):
            return self._forbidden_response(route_path, "Open Zadoo via your public link", public_only=True)

        # WebSocket upgrades must be authenticated before the stream handlers run.
        upgrade_val = (request_headers.get("Upgrade") or "").lower() if request_headers else ""
        connection_val = (request_headers.get("Connection") or "").lower() if request_headers else ""
        if "websocket" in upgrade_val or "upgrade" in connection_val:
            if self._is_ws_authorized(route_path, request_headers):
                return None
            return self._plain_response("Forbidden", http.HTTPStatus.FORBIDDEN)

        feature = self._feature_for_route(route_path)
        if feature and not self._feature_allowed_for_headers(request_headers, feature, connection):
            return self._forbidden_response(route_path)
        if self._http_route_requires_csrf(route_path, feature) and not self._state_changing_http_allowed(request_headers):
            return self._forbidden_response(route_path)

        # Process routes
        if static_asset := self.STATIC_ASSETS.get(route_path):
            return self._static_response(*static_asset)
        if route_path == "/":
            return self._template_response(
                "index.html",
                {
                    "Permissions-Policy": "clipboard-read=(self), clipboard-write=(self)",
                    "Cache-Control": "no-store, no-cache, must-revalidate, max-age=0",
                },
            )
        if route_path == "/api/auth":
            return await self.handle_auth(path, request_headers)
        if route_path == "/api/settings/reload":
            if not self._runtime_admin_code_valid(request_headers):
                return self._json_response({"success": False, "error": "Invalid access code"}, http.HTTPStatus.UNAUTHORIZED)
            self._load_alert_presets_from_settings()
            self.auth_sessions = {}
            return self._json_response({"success": True, "settings": self.settings_store.owner_view()})
        if route_path == "/api/runtime/status":
            return self._json_response({
                "success": True,
                "running": True,
                "port": self.port,
                "tunnel_enabled": self.enable_tunnel,
                "tunnel_block_reason": self.tunnel_block_reason,
                "tunnel_error": (self.tunnel_manager.last_error if self.tunnel_manager else None),
                "public_url": (self.tunnel_manager.get_current_url() if self.tunnel_manager else None),
            })
        if route_path == "/api/runtime/stop":
            if not self._runtime_admin_code_valid(request_headers):
                return self._json_response({"success": False, "error": "Invalid access code"}, http.HTTPStatus.UNAUTHORIZED)
            try:
                self.stop()  # graceful: signals stop_event, closes the listeners
            except Exception as exc:
                return self._json_response({"success": False, "error": str(exc)}, http.HTTPStatus.INTERNAL_SERVER_ERROR)
            return self._json_response({"success": True, "message": "Zadoo runtime stopping"})
        if route_path == "/api/runtime/refresh-tunnel":
            # Local-only tunnel rotation for the desktop Settings window (Open / Refresh).
            if not self._runtime_admin_code_valid(request_headers):
                return self._json_response({"success": False, "error": "Invalid access code"}, http.HTTPStatus.UNAUTHORIZED)
            try:
                return self._json_response(await self._refresh_tunnel_payload())
            except Exception as exc:
                return self._json_response({"success": False, "error": str(exc)}, http.HTTPStatus.INTERNAL_SERVER_ERROR)
        if route_path == "/terminal.html":
            return self._template_response("terminal.html")
        if route_path == "/benchmark.html":
            return self._template_response(
                "benchmark.html",
                {"Cache-Control": "no-store, no-cache, must-revalidate, max-age=0"},
            )
        if route_path in {"/video", "/input", "/audio", "/terminal", "/webcam"}:
            return None
        if route_path == "/api/list-cameras":
            return await self._device_list_response(self.enumerate_cameras)
        if route_path == "/api/list-mics":
            return await self._device_list_response(self.enumerate_microphones)
        if route_path == "/host-controls":
            return self._template_response("host_controls.html", strip=True)
        if route_path == "/api/alert":
            try:
                qs = urllib.parse.parse_qs(
                    urllib.parse.urlparse(path).query,
                    keep_blank_values=True,
                    strict_parsing=True,
                )
            except ValueError as exc:
                return self._json_response({"ok": False, "error": str(exc)}, http.HTTPStatus.BAD_REQUEST)
            if set(qs) != {"code"} or len(qs["code"]) != 1:
                return self._json_response(
                    {"ok": False, "error": "Alert request must contain exactly one code parameter"},
                    http.HTTPStatus.BAD_REQUEST,
                )
            code = qs["code"][0].upper().strip()
            if code not in {"A", "B", "C", "D"}:
                return self._json_response(
                    {"ok": False, "error": "Alert code must be A, B, C, or D"},
                    http.HTTPStatus.BAD_REQUEST,
                )
            try:
                self._load_alert_presets_from_settings()
                if code not in self.alert_presets:
                    return self._json_response({"ok": False, "error": "Alert slot is not set"}, http.HTTPStatus.NOT_FOUND)
                title, message = self.alert_presets[code]
                self._broadcast_controller_alert(title, message)
                return self._json_response({"ok": True})
            except Exception as exc:
                _log_except("api.alert", exc)
                return self._json_response(
                    {"ok": False, "error": str(exc)},
                    http.HTTPStatus.INTERNAL_SERVER_ERROR,
                )
        if route_path == "/api/local/credits":
            if not self.settings_store.get_device_token():
                return self._json_response({"success": False, "offline": True, "error": "Device not signed in"})
            return await self._proxy_cloud("GET", "/api/agent/credits", request_headers)
        if route_path == "/api/local/payment-market":
            return self._json_response({"success": True, **self._payment_market_payload(request_headers)})
        if route_path == "/api/local/topup-order":
            raw_amount = self._header_get(request_headers, "X-Zadoo-Amount-Minor", "")
            try:
                amount = int(raw_amount)
            except (TypeError, ValueError):
                return self._json_response(
                    {"success": False, "error": "X-Zadoo-Amount-Minor must be an integer"},
                    http.HTTPStatus.BAD_REQUEST,
                )
            if amount <= 0:
                return self._json_response(
                    {"success": False, "error": "X-Zadoo-Amount-Minor must be positive"},
                    http.HTTPStatus.BAD_REQUEST,
                )
            body = {"amountMinor": amount}
            return await self._proxy_cloud(
                "POST",
                "/api/agent/wallet/topup-order",
                request_headers,
                json.dumps(body).encode("utf-8"),
            )
        if route_path == "/api/local/topup-verify":
            header_fields = {
                "paymentId": "X-Zadoo-Payment-Id",
                "razorpay_payment_id": "X-Razorpay-Payment-Id",
                "razorpay_order_id": "X-Razorpay-Order-Id",
                "razorpay_signature": "X-Razorpay-Signature",
            }
            body = {
                key: str(self._header_get(request_headers, header, "") or "").strip()
                for key, header in header_fields.items()
            }
            missing = [header for key, header in header_fields.items() if not body[key]]
            if missing:
                return self._json_response(
                    {"success": False, "error": f"Missing payment headers: {', '.join(missing)}"},
                    http.HTTPStatus.BAD_REQUEST,
                )
            resp = await self._proxy_cloud(
                "POST",
                "/api/agent/wallet/topup-verify",
                request_headers,
                json.dumps(body).encode("utf-8"),
            )
            try:
                vd = json.loads(resp.body.decode("utf-8"))
                if vd["success"]:
                    refreshed = await asyncio.to_thread(ZadooCloudClient(self.settings_store).entitlement)
                    if not refreshed["success"]:
                        return self._json_response(
                            {
                                "success": False,
                                "error": f"Payment verified but entitlement refresh failed: {refreshed['error']}",
                            },
                            http.HTTPStatus.BAD_GATEWAY,
                        )
                    session_token = self._cookie_value(request_headers, self.AUTH_COOKIE_NAME)
                    self._grace_locks_by_session.pop(session_token, None)
            except Exception as exc:
                return self._json_response(
                    {"success": False, "error": f"Payment verification state update failed: {exc}"},
                    http.HTTPStatus.INTERNAL_SERVER_ERROR,
                )
            return resp
        if route_path == "/brand-header.png":
            return self._png_file_response("brand-header.png", "Header image not found")
        if route_path == "/trigger-icon.png":
            return self._png_file_response("trigger-icon.png", "Trigger icon not found")
        if route_path == "/splash.png":
            return self._png_file_response("splash.png", "Splash image not found")
        if route_path == "/snapshot":
            try:
                parsed_url = urllib.parse.urlparse(path)
                query_params = urllib.parse.parse_qs(
                    parsed_url.query,
                    keep_blank_values=True,
                    strict_parsing=True,
                )

                rect_norm = None
                rect_keys = {'x0', 'y0', 'x1', 'y1'}
                allowed_keys = rect_keys | {'fmt', 'q', 'max_w', 'max_h'}
                unknown_keys = sorted(set(query_params) - allowed_keys)
                if unknown_keys:
                    raise ValueError(f"Unsupported snapshot parameters: {', '.join(unknown_keys)}")
                duplicate_keys = sorted(key for key, values in query_params.items() if len(values) != 1)
                if duplicate_keys:
                    raise ValueError(f"Duplicate snapshot parameters: {', '.join(duplicate_keys)}")
                present_rect_keys = rect_keys.intersection(query_params)
                if present_rect_keys and present_rect_keys != rect_keys:
                    missing = ", ".join(sorted(rect_keys - present_rect_keys))
                    raise ValueError(f"Snapshot region is missing: {missing}")
                if present_rect_keys:
                    try:
                        rect_norm = {key: float(query_params[key][0]) for key in rect_keys}
                    except ValueError as exc:
                        raise ValueError("Snapshot region coordinates must be numbers") from exc
                    if any(not math.isfinite(value) or value < 0.0 or value > 1.0 for value in rect_norm.values()):
                        raise ValueError("Snapshot region coordinates must be finite values between 0 and 1")
                    if rect_norm['x1'] <= rect_norm['x0'] or rect_norm['y1'] <= rect_norm['y0']:
                        raise ValueError("Snapshot region must have positive width and height")
                fmt = query_params.get('fmt', ['png'])[0].lower().strip()
                if fmt not in ('jpeg', 'jpg', 'png'):
                    raise ValueError(f"Unsupported snapshot format: {fmt}")
                try:
                    quality = int(query_params.get('q', [85])[0])
                except ValueError as exc:
                    raise ValueError("Snapshot q must be an integer") from exc
                if not 1 <= quality <= 95:
                    raise ValueError("Snapshot JPEG quality must be between 1 and 95")

                def positive_limit(name):
                    if name not in query_params:
                        return None
                    try:
                        value = int(query_params[name][0])
                    except ValueError as exc:
                        raise ValueError(f"Snapshot {name} must be a positive integer") from exc
                    if value < 1:
                        raise ValueError(f"Snapshot {name} must be a positive integer")
                    return value

                max_w = positive_limit('max_w')
                max_h = positive_limit('max_h')
                img_bytes = await asyncio.to_thread(
                    self._generate_snapshot,
                    rect_norm=rect_norm,
                    fmt=fmt,
                    quality=quality,
                    max_w=max_w,
                    max_h=max_h,
                )
                ext = "png" if fmt == 'png' else "jpg"
                return self._response(
                    img_bytes,
                    "image/png" if fmt == 'png' else "image/jpeg",
                    extra_headers={
                        "Cache-Control": "no-store, no-cache, must-revalidate, max-age=0",
                        "Content-Disposition": f"attachment; filename=\"snap-{time.strftime('%Y%m%d-%H%M%S')}.{ext}\"",
                    },
                )
            except ValueError as exc:
                return self._plain_response(str(exc), http.HTTPStatus.BAD_REQUEST)
            except Exception as exc:
                return self._plain_response(str(exc), http.HTTPStatus.INTERNAL_SERVER_ERROR)
        if route_path.startswith("/api/"):
            return self._json_response({"success": False, "error": "Not Found"}, http.HTTPStatus.NOT_FOUND)
        return self._plain_response("Not Found", http.HTTPStatus.NOT_FOUND)

    def _public_url_payload(self):
        if not self.tunnel_manager:
            return {
                'success': False,
                'error': self.tunnel_block_reason or 'Tunnel manager is not initialized',
                'email_status': None,
            }
        current_url = self.tunnel_manager.get_current_url()
        if current_url:
            current_port = self.tunnel_manager.primary_port
            return {
                'success': True,
                'url': current_url,
                'port': current_port,
                'message': f'Current public URL for port {current_port}',
                'email_status': self.tunnel_manager.last_email_message,
            }
        return {
            'success': False,
            'error': self.tunnel_manager.last_error or 'Public URL is not available yet',
            'email_status': self.tunnel_manager.last_email_message,
        }

    async def _refresh_tunnel_payload(self):
        if not self.tunnel_manager:
            return self._public_url_payload()

        print(" Refreshing tunnel...")
        url = await asyncio.to_thread(self.tunnel_manager.refresh_tunnel)
        if not isinstance(url, str) or not url.strip():
            raise RuntimeError("Cloudflare tunnel refresh returned no public URL")
        url = url.strip()

        current_port = self.tunnel_manager.primary_port
        print(f" New tunnel URL: {url}")
        synced = await asyncio.to_thread(ZadooCloudClient(self.settings_store).heartbeat, url)
        if not synced["success"]:
            raise RuntimeError(f"Tunnel refreshed but cloud URL sync failed: {synced['error']}")
        return {
            "success": True,
            "url": url,
            "port": current_port,
            "message": f"Successfully refreshed tunnel on port {current_port}",
            "email_status": self.tunnel_manager.last_email_message,
        }

    def _capture_screen_image(self, rect_norm=None):
        """Capture a one-shot snapshot with MSS."""
        rect = None
        if rect_norm:
            width, height = get_primary_screen_size()
            left = max(0, min(width, int(float(rect_norm['x0']) * width)))
            top = max(0, min(height, int(float(rect_norm['y0']) * height)))
            right = max(0, min(width, int(float(rect_norm['x1']) * width)))
            bottom = max(0, min(height, int(float(rect_norm['y1']) * height)))
            if right <= left or bottom <= top:
                raise ValueError("Snapshot region must have positive width and height")
            rect = {'left': left, 'top': top, 'width': right - left, 'height': bottom - top}
        try:
            with mss.mss() as sct:
                shot = sct.grab(rect if rect is not None else sct.monitors[1])
            arr = np.frombuffer(shot.bgra, dtype=np.uint8).reshape((shot.height, shot.width, 4))
            return np.ascontiguousarray(arr[..., :3][:, :, ::-1])
        except Exception as exc:
            raise RuntimeError(f"MSS snapshot capture failed: {exc}") from exc

    def _generate_snapshot(self, rect_norm=None, fmt='png', quality=85, max_w=None, max_h=None):
        """Return PNG or JPEG bytes for a full or cropped one-shot snapshot."""
        try:
            arr = self._capture_screen_image(rect_norm=rect_norm)
            height, width = arr.shape[:2]
            if (max_w and width > max_w) or (max_h and height > max_h):
                ratio = min((max_w / width) if max_w else 1.0, (max_h / height) if max_h else 1.0)
                size = (max(1, int(width * ratio)), max(1, int(height * ratio)))
                arr = np.asarray(Image.fromarray(arr).resize(size, Image.Resampling.BILINEAR))
            return (
                png_encode(arr, level=0)
                if fmt == 'png'
                else jpeg8_encode(arr, level=quality)
            )
        except Exception as exc:
            raise RuntimeError(f"Snapshot {fmt} generation failed: {exc}") from exc
