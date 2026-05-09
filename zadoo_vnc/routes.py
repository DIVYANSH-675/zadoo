"""HTTP route and snapshot handlers."""
from __future__ import annotations

import asyncio
import http
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

from .assets import load_host_controls_html, load_index_html, load_terminal_html
from .camera_discovery import enumerate_camera_devices
from .config import BRAND_HEADER_IMAGE_PATH, SPLASH_IMAGE_PATH, TRIGGER_ICON_IMAGE_PATH
from .dependencies import *
from .logging_utils import _log_except, _log_try_ok

class RoutesMixin:
    AUTH_COOKIE_NAME = "zadoo_auth"
    AUTH_TTL_SECONDS = 3600

    def _json_response(self, payload, status=http.HTTPStatus.OK, extra_headers=None):
        headers = Headers()
        headers["Content-Type"] = "application/json; charset=utf-8"
        headers["Cache-Control"] = "no-store"
        for key, value in (extra_headers or {}).items():
            headers[key] = value
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

    def _coerce_response(self, response):
        if isinstance(response, WSResponse):
            return response
        if isinstance(response, tuple) and len(response) == 3:
            status, raw_headers, body = response
            try:
                status_obj = status if isinstance(status, http.HTTPStatus) else http.HTTPStatus(int(status))
            except Exception:
                status_obj = http.HTTPStatus.INTERNAL_SERVER_ERROR
            headers = Headers()
            try:
                iterator = raw_headers.items() if isinstance(raw_headers, dict) else raw_headers
                for key, value in iterator:
                    headers[str(key)] = str(value)
            except Exception:
                headers["Content-Type"] = "application/octet-stream"
            return WSResponse(
                status_code=int(status_obj),
                reason_phrase=status_obj.phrase,
                headers=headers,
                body=body if isinstance(body, (bytes, bytearray)) else str(body).encode("utf-8"),
            )
        return self._plain_response("Internal Server Error", http.HTTPStatus.INTERNAL_SERVER_ERROR)

    def _header_get(self, request_headers, name, default=None):
        try:
            return request_headers.get(name, default)
        except Exception:
            try:
                return request_headers.get(name.lower(), default)
            except Exception:
                return default

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
        def clean(value, default=""):
            return str(value if value is not None else default).strip()

        codes = {
            "full": clean(os.getenv("CODE_FULL"), "TERMINATOR"),
            "limited": clean(os.getenv("CODE_LIMITED"), "ADVENTURES"),
            "partial": clean(os.getenv("CODE_PARTIAL"), "INNOVATION"),
            "lockdown": clean(os.getenv("CODE_LOCKDOWN"), "CHALLENGER"),
        }
        custom = clean(os.getenv("CUSTOM_PASSWORD"))
        if custom:
            codes["custom"] = custom
        return {role: value.upper() for role, value in codes.items() if value}

    def _match_auth_code(self, code):
        submitted = str(code or "").strip().upper()
        if not submitted:
            return None
        for role, expected in self._auth_codes().items():
            if submitted == expected:
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
        self._cleanup_auth_sessions()
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
        if route_path in {"/video", "/audio"}:
            return "view"
        if route_path == "/input":
            return "control"
        if route_path in {"/ssh", "/terminal.html"}:
            return "terminal"
        if route_path == "/webcam":
            return "webcam"
        if route_path == "/mic":
            return "mic"
        return None

    def _http_feature_for_route(self, route_path):
        if route_path in {"/", "/api/auth", "/brand-header.png", "/trigger-icon.png", "/splash.png"}:
            return "public"
        if route_path in {"/api/public-url", "/snapshot"}:
            return "view"
        if route_path in {"/api/set-quality", "/api/set-fps", "/api/set-clipboard-image"}:
            return "control"
        if route_path == "/api/refresh-tunnel":
            return "host"
        if route_path in {"/terminal.html", "/ssh"}:
            return "terminal"
        if route_path in {"/webcam", "/api/list-cameras"}:
            return "webcam"
        if route_path == "/mic":
            return "mic"
        if route_path in {"/host-controls", "/api/alert"}:
            return "host"
        if route_path.startswith("/api/client-log"):
            return "view"
        return None

    def _is_ws_authorized(self, route_path, request_headers):
        feature = self._feature_for_route(route_path)
        if feature is None:
            return False
        return self._is_authorized(request_headers, feature)

    async def handle_auth(self, path):
        parsed = urllib.parse.urlparse(str(path or ""))
        query = urllib.parse.parse_qs(parsed.query or "")
        role = self._match_auth_code((query.get("code") or [""])[0])
        if not role:
            return self._json_response({"success": False, "error": "Invalid code"}, http.HTTPStatus.UNAUTHORIZED)

        token = secrets.token_urlsafe(32)
        expires_at = time.time() + self.AUTH_TTL_SECONDS
        sessions = getattr(self, "auth_sessions", None)
        if not isinstance(sessions, dict):
            self.auth_sessions = {}
            sessions = self.auth_sessions
        sessions[token] = {"role": role, "expires_at": expires_at}
        cookie = (
            f"{self.AUTH_COOKIE_NAME}={token}; Path=/; Max-Age={self.AUTH_TTL_SECONDS}; "
            "HttpOnly; SameSite=Lax"
        )
        return self._json_response(
            {"success": True, "mode": role, "expires_in": self.AUTH_TTL_SECONDS},
            extra_headers={"Set-Cookie": cookie},
        )

    def enumerate_cameras(self):
        return enumerate_camera_devices()

    def _apply_quality(self, raw_value, default=75):
        try:
            value = max(1, min(95, int(raw_value)))
        except Exception:
            value = max(1, min(95, int(default)))
        self.current_quality = value
        if self.screen_capturer:
            self.screen_capturer.quality = value
        return value

    def _apply_fps(self, raw_value, default=30):
        try:
            value = max(1, min(120, int(raw_value)))
        except Exception:
            value = max(1, min(120, int(default)))
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
            elif len(args) == 1:
                # Fallback: older style may pass a single connection-like object
                connection = args[0]
                # Try to resolve request then path
                request = getattr(connection, "request", None)
                if request is not None:
                    path = getattr(request, "path", None)
                    request_headers = getattr(request, "headers", None)
                else:
                    path = getattr(connection, "path", None)
                    request_headers = getattr(connection, "request_headers", None)
        except Exception:
            pass

        if not isinstance(path, str):
            path = "/"
        route_path = urllib.parse.urlparse(path).path

        logging.debug("process_request: path=%s", path)

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

        feature = self._http_feature_for_route(route_path)
        if feature and feature != "public" and not self._is_authorized(request_headers, feature):
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
            return await self.handle_auth(path)
        elif isinstance(path, str) and route_path == "/api/public-url":
            try:
                return self._json_response(json.loads(await self.handle_get_public_url()))
            except Exception as e:
                return self._json_response({"success": False, "error": str(e)}, http.HTTPStatus.INTERNAL_SERVER_ERROR)
        elif isinstance(path, str) and route_path == "/api/refresh-tunnel":
            try:
                return self._json_response(json.loads(await self.handle_refresh_tunnel()))
            except Exception as e:
                return self._json_response({"success": False, "error": str(e)}, http.HTTPStatus.INTERNAL_SERVER_ERROR)
        elif isinstance(path, str) and route_path == "/api/set-quality":
            return self._json_response(await self.handle_set_quality(path))
        elif isinstance(path, str) and route_path == "/api/set-fps":
            return self._json_response(await self.handle_set_fps(path))
        elif isinstance(path, str) and route_path == "/api/set-clipboard-image":
            status, headers_dict, body = await self.handle_set_clipboard_image(request_headers)
            headers = Headers()
            for key, value in headers_dict.items():
                headers[key] = value
            return WSResponse(
                status_code=int(status),
                reason_phrase=status.phrase,
                headers=headers,
                body=body,
            )
        elif isinstance(path, str) and path.startswith("/api/client-log"):
            try:
                # Accept simple GET with ?msg=... or POST with text body
                if request_headers is None:
                    body_bytes = b""
                else:
                    body_bytes = getattr(args[1], "body", b"") if len(args) >= 2 else b""
                from urllib.parse import urlparse, parse_qs, unquote
                parsed = urlparse(path)
                qs = parse_qs(parsed.query or "")
                msg = (qs.get("msg") or [""])[0]
                if not msg and isinstance(body_bytes, (bytes, bytearray)) and body_bytes:
                    try:
                        msg = body_bytes.decode("utf-8", "ignore")
                    except Exception:
                        msg = str(body_bytes)
                msg = unquote(msg)
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
                devices = self.enumerate_cameras()
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
                from urllib.parse import urlparse, parse_qs
                qs = parse_qs(urlparse(path).query or "")
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
            try:
                with open(BRAND_HEADER_IMAGE_PATH, 'rb') as f:
                    img_bytes = f.read()
                headers = Headers()
                headers["Content-Type"] = "image/png"
                headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
                return WSResponse(
                    status_code=int(http.HTTPStatus.OK),
                    reason_phrase=http.HTTPStatus.OK.phrase,
                    headers=headers,
                    body=img_bytes,
                )
            except Exception:
                headers = Headers()
                headers["Content-Type"] = "text/plain; charset=utf-8"
                return WSResponse(
                    status_code=int(http.HTTPStatus.NOT_FOUND),
                    reason_phrase=http.HTTPStatus.NOT_FOUND.phrase,
                    headers=headers,
                    body=b"Header image not found",
                )
        elif route_path == "/trigger-icon.png":
            try:
                with open(TRIGGER_ICON_IMAGE_PATH, 'rb') as f:
                    img_bytes = f.read()
                headers = Headers()
                headers["Content-Type"] = "image/png"
                headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
                return WSResponse(
                    status_code=int(http.HTTPStatus.OK),
                    reason_phrase=http.HTTPStatus.OK.phrase,
                    headers=headers,
                    body=img_bytes,
                )
            except Exception:
                headers = Headers()
                headers["Content-Type"] = "text/plain; charset=utf-8"
                return WSResponse(
                    status_code=int(http.HTTPStatus.NOT_FOUND),
                    reason_phrase=http.HTTPStatus.NOT_FOUND.phrase,
                    headers=headers,
                    body=b"Trigger icon not found",
                )
        elif route_path == "/splash.png":
            try:
                with open(SPLASH_IMAGE_PATH, 'rb') as f:
                    img_bytes = f.read()
                headers = Headers()
                headers["Content-Type"] = "image/png"
                headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
                return WSResponse(
                    status_code=int(http.HTTPStatus.OK),
                    reason_phrase=http.HTTPStatus.OK.phrase,
                    headers=headers,
                    body=img_bytes,
                )
            except Exception:
                headers = Headers()
                headers["Content-Type"] = "text/plain; charset=utf-8"
                return WSResponse(
                    status_code=int(http.HTTPStatus.NOT_FOUND),
                    reason_phrase=http.HTTPStatus.NOT_FOUND.phrase,
                    headers=headers,
                    body=b"Splash image not found",
            )
        elif path.startswith("/snapshot"):
                try:
                    import time as _t
                    t_req0 = _t.perf_counter()
                    # NEW: Parse URL for crop coordinates
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
                            pass # Ignore invalid coordinates
                    # Optional encode params (default to PNG to match clipboard expectations)
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
                    # Log incoming snapshot request
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

                    img_bytes = self._generate_snapshot_png(rect_norm=rect_norm, fmt=fmt, quality=quality, max_w=max_w, max_h=max_h)
                    t_req1 = _t.perf_counter()

                    if img_bytes:
                        headers = Headers()
                        headers["Content-Type"] = ("image/png" if fmt == 'png' else "image/jpeg")
                        headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
                        try:
                            import time as _t
                            ext = ("png" if fmt == 'png' else "jpg")
                            fname = f"snap-{_t.strftime('%Y%m%d-%H%M%S')}.{ext}"
                            headers["Content-Disposition"] = f"attachment; filename=\"{fname}\""
                        except Exception:
                            pass
                        # Log end-to-end server timing before returning
                        try:
                            logging.info(
                                "[snapshot.server] total=%.2fms generate=%.2fms bytes=%.1fKB fmt=%s rect=%s",
                                (t_req1 - t_req0) * 1000.0,
                                (t_req1 - t_req0) * 1000.0,
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
                    else:
                        headers = Headers()
                        headers["Content-Type"] = "text/plain; charset=utf-8"
                        return WSResponse(
                            status_code=int(http.HTTPStatus.INTERNAL_SERVER_ERROR),
                            reason_phrase=http.HTTPStatus.INTERNAL_SERVER_ERROR.phrase,
                            headers=headers,
                            body=b"Failed to capture snapshot",
                        )
                except Exception as e:
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

    async def handle_get_public_url(self):
        """API endpoint to get current public URL"""
        try:
            if self.tunnel_manager and self.tunnel_manager.get_current_url():
                current_url = self.tunnel_manager.get_current_url()
                current_port = self.tunnel_manager.current_port
                return json.dumps({
                    'success': True, 
                    'url': current_url,
                    'port': current_port,
                    'message': f'Current public URL for port {current_port}'
                })
            else:
                return json.dumps({
                    'success': False, 
                    'error': 'No public URL available'
                })
        except Exception as e:
            return json.dumps({
                'success': False, 
                'error': f'Error getting URL: {str(e)}'
            })

    async def handle_refresh_tunnel(self):
        """API endpoint to refresh tunnel and get new URL."""
        try:
            if not self.tunnel_manager:
                print(" No tunnel manager available")
                return json.dumps({"success": False, "error": "No tunnel manager available"})

            print(" Refreshing tunnel...")
            loop = asyncio.get_event_loop()
            url = await loop.run_in_executor(None, self.tunnel_manager.refresh_tunnel)
            if not url:
                print(" Failed to restart tunnel")
                return json.dumps({"success": False, "error": "Failed to refresh tunnel"})

            print(f" New tunnel URL: {url}")
            return json.dumps({
                "success": True,
                "url": url,
                "port": self.tunnel_manager.current_port,
                "email_status": self.tunnel_manager.last_email_message,
            })
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

    async def handle_set_clipboard_image(self, request_headers=None):
        """Compatibility response for the old HTTP image clipboard endpoint."""
        response_data = json.dumps({
            "success": False,
            "error": "HTTP image clipboard upload is not supported by this WebSocket server.",
            "use": "Send {action:'set_clipboard_image', mime, data_base64} to the /input WebSocket.",
        }).encode("utf-8")
        return http.HTTPStatus.BAD_REQUEST, {"Content-Type": "application/json; charset=utf-8"}, response_data

    def _capture_screen_image(self, rect_norm=None):
        """One-shot snapshot; reuse instances and capture at source (dxcam region  MSS region)."""
        import time as _t
        t0 = _t.perf_counter()
        # Track which backend we actually used for logging
        try:
            self._last_snapshot_backend = 'none'
        except Exception:
            pass
        # Helper: convert normalized rect to absolute desktop rect
        def _norm_to_abs_rect():
            if not rect_norm:
                return None
            try:
                try:
                    import ctypes as _ct
                    u32 = _ct.windll.user32
                    u32.SetProcessDPIAware()
                    sw = u32.GetSystemMetrics(0)
                    sh = u32.GetSystemMetrics(1)
                except Exception:
                    import pyautogui as _pg
                    sw, sh = _pg.size()
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
                # Get desktop size for ratio
                try:
                    import ctypes as _ct2
                    u322 = _ct2.windll.user32
                    u322.SetProcessDPIAware()
                    sw2 = u322.GetSystemMetrics(0)
                    sh2 = u322.GetSystemMetrics(1)
                except Exception:
                    try:
                        import pyautogui as _pg2
                        sw2, sh2 = _pg2.size()
                    except Exception:
                        sw2, sh2 = 0, 0
                if sw2 > 0 and sh2 > 0:
                    l, t, r, b = abs_rect
                    area_ratio = ((r - l) * (b - t)) / float(sw2 * sh2)
                    prefer_mss_first = area_ratio <= 0.25
            except Exception:
                prefer_mss_first = False

        def _try_dxcam_then_none():
            try:
                cam = getattr(self, 'dxcam_camera', None)
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
                        except Exception:
                            return Image.frombuffer('RGB', (rgb.shape[1], rgb.shape[0]), rgb.tobytes())
            except Exception:
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
                    except Exception:
                        return Image.frombuffer('RGB', (rgb.shape[1], rgb.shape[0]), rgb.tobytes())
                except Exception:
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
                logging.warning('Snapshot fast_ctypes failed; falling back', exc_info=True)

        return None

    def _generate_snapshot_png(self, rect_norm=None, fmt='png', quality=85, max_w=None, max_h=None):
        """
        Returns image bytes (PNG or JPEG) for the current (or region) snapshot.
        Supports format/quality/downscale and logs captureconvertencode timings.
        """
        try:
            import time as _t, io as _io
            t0 = _t.perf_counter()
            img = self._capture_screen_image(rect_norm=rect_norm)
            t1 = _t.perf_counter()
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
            t2 = _t.perf_counter()
            out = None
            if fmt == 'png' and HAS_IMAGECODECS:
                try:
                    out = imagecodecs.png_encode(arr, level=0)
                except Exception:
                    out = None
            if out is None:
                buf = _io.BytesIO()
                try:
                    if fmt == 'png':
                        img.save(buf, format='PNG', optimize=False, compress_level=0)
                    else:
                        # JPEG fast path via imagecodecs not used here to keep code simple; Pillow is fine
                        img.save(buf, format='JPEG', quality=max(1, min(95, int(quality))))
                    out = buf.getvalue()
                except Exception:
                    out = None
            t3 = _t.perf_counter()
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
