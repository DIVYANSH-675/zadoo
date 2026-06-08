"""Smoke checks for the Zadoo VNC runtime."""
from __future__ import annotations

import argparse
import asyncio
import base64
import json
import os
import re
import secrets
import sys
import tempfile
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SOURCE_GLOBS = ("zadoo_vnc/**/*.py", "zadoo_vnc/templates/*.html", "scripts/*.py", "*.py", "*.md", "*.toml", "*.txt", ".env.example")
MOJIBAKE_RE = re.compile(r"[\u00c2\u00c3\u00e2\u00f0][^\x00-\x7f]+")


def _legacy_auth_strings():
    return tuple(
        "".join(chr(part) for part in value)
        for value in (
            (84, 69, 82, 77, 73, 78, 65, 84, 79, 82),
            (65, 68, 86, 69, 78, 84, 85, 82, 69, 83),
            (73, 78, 78, 79, 86, 65, 84, 73, 79, 78),
            (67, 72, 65, 76, 76, 69, 78, 71, 69, 82),
        )
    )


FORBIDDEN_SOURCE_STRINGS = (
    "re_" + "GF7",
    "resend_" + "api_key_default",
    "email_" + "to_default",
    "iskssj07" + "@gmail.com",
    "Set-Clipboard" + " -Value @\"\"",
    "Popen(" + "'clip'",
) + _legacy_auth_strings()


def fail(message: str) -> None:
    print(f"FAIL: {message}")
    raise SystemExit(1)


def ok(message: str) -> None:
    print(f"OK: {message}")


def set_cookie_headers(headers) -> list[str]:
    try:
        values = headers.get_all("Set-Cookie")
        if values:
            return list(values)
    except Exception:
        pass
    try:
        return [value for key, value in headers.raw_items() if key.lower() == "set-cookie"]
    except Exception:
        pass
    try:
        value = headers.get("Set-Cookie")
        return [value] if value else []
    except Exception:
        return []


def cookie_header_from_set_cookie(headers) -> str:
    cookies = []
    for value in set_cookie_headers(headers):
        first = str(value).split(";", 1)[0].strip()
        if first:
            cookies.append(first)
    return "; ".join(cookies)


def assert_camera_payload(body: bytes, *, require_objects: bool = False) -> None:
    try:
        payload = json.loads(body.decode("utf-8"))
    except Exception as exc:
        fail(f"camera payload is not valid JSON: {exc}")
    if not isinstance(payload, dict) or payload.get("success") is not True:
        fail("camera payload missing success=true")
    devices = payload.get("devices")
    if not isinstance(devices, list):
        fail("camera payload devices is not a list")
    for index, device in enumerate(devices):
        if isinstance(device, str):
            if require_objects:
                fail(f"camera device {index} used legacy string shape")
            continue
        if not isinstance(device, dict):
            fail(f"camera device {index} is neither string nor object")
        for key in ("id", "name", "label", "open_name", "source"):
            if not isinstance(device.get(key), str) or not device.get(key).strip():
                fail(f"camera device {index} missing non-empty {key!r}")


def assert_mic_payload(body: bytes) -> None:
    try:
        payload = json.loads(body.decode("utf-8"))
    except Exception as exc:
        fail(f"mic payload is not valid JSON: {exc}")
    if not isinstance(payload, dict) or payload.get("success") is not True:
        fail("mic payload missing success=true")
    devices = payload.get("devices")
    if not isinstance(devices, list):
        fail("mic payload devices is not a list")
    for index, device in enumerate(devices):
        if not isinstance(device, dict):
            fail(f"mic device {index} is not an object")
        for key in ("id", "name", "label"):
            if not isinstance(device.get(key), str) or not device.get(key).strip():
                fail(f"mic device {index} missing non-empty {key!r}")


def source_files():
    seen = set()
    for glob in SOURCE_GLOBS:
        for path in ROOT.glob(glob):
            if path.is_file() and path not in seen:
                seen.add(path)
                yield path


def assert_imports() -> None:
    sys.path.insert(0, str(ROOT))
    test_auth_code = "SMOKE_" + secrets.token_hex(8).upper()
    temp_settings = tempfile.TemporaryDirectory(prefix="zadoo-smoke-")
    old_env = {
        key: os.environ.get(key)
        for key in (
            "ZADOO_ACCESS_CODE",
            "ZADOO_SETTINGS_PATH",
            "ZADOO_SETTINGS_DIR",
            "ZADOO_ALLOW_QUERY_AUTH",
            "ZADOO_ALLOWED_ORIGINS",
            "ZADOO_AUTH_MAX_FAILURES",
            "ZADOO_AUTH_WINDOW_SECONDS",
            "ZADOO_AUTH_LOCKOUT_SECONDS",
            "ZADOO_CLIPBOARD_IMAGE_MAX_BYTES",
            "ZADOO_CLIPBOARD_TEXT_MAX_BYTES",
            "ZADOO_ADAPTIVE_STREAM",
            "ZADOO_STREAM_START_PROFILE",
            "ALERT_A",
            "ALERT_B",
            "ALERT_C",
            "ALERT_D",
            "ALERT_A_TITLE",
            "ALERT_A_MESSAGE",
            "ALERT_B_TITLE",
            "ALERT_B_MESSAGE",
            "ALERT_C_TITLE",
            "ALERT_C_MESSAGE",
            "ALERT_D_TITLE",
            "ALERT_D_MESSAGE",
        )
    }
    os.environ["ZADOO_ACCESS_CODE"] = test_auth_code
    os.environ["ZADOO_ALLOW_DIRECT_ACCESS"] = "1"  # tests drive the viewer locally (no tunnel)
    os.environ["ZADOO_SETTINGS_PATH"] = str(Path(temp_settings.name) / "config.json")
    os.environ["ZADOO_AUTH_MAX_FAILURES"] = "3"
    os.environ["ZADOO_AUTH_WINDOW_SECONDS"] = "60"
    os.environ["ZADOO_AUTH_LOCKOUT_SECONDS"] = "120"
    os.environ["ZADOO_CLIPBOARD_IMAGE_MAX_BYTES"] = "8"
    os.environ["ZADOO_CLIPBOARD_TEXT_MAX_BYTES"] = "8"
    os.environ["ZADOO_ADAPTIVE_STREAM"] = "1"
    os.environ["ZADOO_STREAM_START_PROFILE"] = "720p120"
    for key in (
        "ZADOO_ALLOW_QUERY_AUTH",
        "ZADOO_ALLOWED_ORIGINS",
        "ZADOO_SETTINGS_DIR",
        "ALERT_A",
        "ALERT_B",
        "ALERT_C",
        "ALERT_D",
        "ALERT_A_TITLE",
        "ALERT_A_MESSAGE",
        "ALERT_B_TITLE",
        "ALERT_B_MESSAGE",
        "ALERT_C_TITLE",
        "ALERT_C_MESSAGE",
        "ALERT_D_TITLE",
        "ALERT_D_MESSAGE",
    ):
        os.environ.pop(key, None)
    import zadoo_vnc.config as config
    import zadoo_vnc.dependencies as deps
    import zadoo_vnc.screen_capture as screen_capture
    import zadoo_vnc.settings as settings_mod
    from zadoo_vnc.assets import load_benchmark_html, load_host_controls_html, load_index_html, load_terminal_html
    from zadoo_vnc.server import VNCServer
    from websockets.datastructures import Headers

    settings_mod._STORE = None

    if not (deps.HAS_PYAUTOGUI and deps.HAS_PIL):
        fail("core capture/input dependency flags are not usable")

    for name, loader in {
        "index.html": load_index_html,
        "terminal.html": load_terminal_html,
        "host_controls.html": load_host_controls_html,
        "benchmark.html": load_benchmark_html,
    }.items():
        if not loader().strip():
            fail(f"template did not load: {name}")

    for asset_path in (config.BRAND_HEADER_IMAGE_PATH, config.SPLASH_IMAGE_PATH, config.TRIGGER_ICON_IMAGE_PATH):
        path = Path(asset_path)
        if not path.exists() or path.stat().st_size <= 0:
            fail(f"asset missing or empty: {path}")

    isolated_store = settings_mod.get_settings_store()
    if isolated_store.configured():
        fail("smoke settings store should start unconfigured")
    if any(item.get("enabled") for item in isolated_store.public_view().get("alerts", {}).values()):
        fail("settings store should not create default alert text")
    access_record = settings_mod.hash_access_code("SMOKE-CODE")
    if not settings_mod.verify_access_code("SMOKE-CODE", access_record):
        fail("access-code verifier rejected the correct code")
    if settings_mod.verify_access_code("wrong", access_record):
        fail("access-code verifier accepted the wrong code")
    try:
        settings_mod.hash_access_code("TOO-LONG-CODE")
        fail("access-code hash accepted a code longer than 10 characters")
    except ValueError:
        pass
    isolated_store.set_access_code("SMOKE")
    external_store = settings_mod.SettingsStore(isolated_store.path)
    external_data = external_store.load(reload=True)
    external_data["setup_complete"] = True
    external_data["permissions"] = {key: True for key in settings_mod.PERMISSION_KEYS}
    time.sleep(0.02)
    external_store.save(external_data)
    if not all((isolated_store.load().get("permissions") or {}).values()):
        fail("settings store did not reload changes saved by another process")
    isolated_store.save(settings_mod.normalize_settings(None))

    try:
        server = VNCServer(6173)
        if server.port != 6173:
            fail("VNCServer did not instantiate with the fixed port")
        if not getattr(server, "adaptive_stream", None) or server.adaptive_stream.profile.name != "720p120":
            fail("adaptive stream did not select the default 720p120 startup profile")
        if server.current_fps != 120 or server.current_quality != 85:
            fail(f"startup profile did not apply fps/locked quality: fps={server.current_fps} quality={server.current_quality}")
        if not server._quality_locked_by_user:
            fail("startup quality should be locked at 85 until the user changes quality")
        methods = screen_capture.ScreenCapturer().get_available_methods()
        expected_methods = []
        if screen_capture.HAS_BETTERCAM:
            expected_methods.append("bettercam")
        if screen_capture.HAS_DXCAM:
            expected_methods.append("dxcam")
        if screen_capture.HAS_PYAUTOGUI and screen_capture.HAS_PIL:
            expected_methods.append("pyautogui")
        if methods != expected_methods:
            fail(f"available capture methods changed: {methods}, expected {expected_methods}")

        region_capturer = screen_capture.ScreenCapturer()
        region_capturer.set_performance_mode(True, "center_0.5", 1)
        if region_capturer.get_active_region_norm() != {"x0": 0.25, "y0": 0.25, "x1": 0.75, "y1": 0.75}:
            fail(f"center performance region changed: {region_capturer.get_active_region_norm()}")
        region_capturer.set_performance_mode(True, "custom", 1)
        region_capturer.set_custom_region({"x0": 0.2, "y0": 0.1, "x1": 0.6, "y1": 0.5})
        server.screen_capturer = region_capturer
        mapped = server._map_view_norm_to_screen_norm(0.5, 0.5)
        if tuple(round(v, 4) for v in mapped) != (0.4, 0.3):
            fail(f"custom ROI mouse mapping failed: {mapped}")
        mapped_corner = server._map_view_norm_to_screen_norm(1.0, 0.0)
        if tuple(round(v, 4) for v in mapped_corner) != (0.6, 0.1):
            fail(f"custom ROI corner mapping failed: {mapped_corner}")
        region_capturer.set_performance_mode(False, "full", 1)
        if server._map_view_norm_to_screen_norm(0.5, 0.5) != (0.5, 0.5):
            fail("disabled performance mode should not remap mouse coordinates")
        server.screen_capturer = None

        class FakeDXCam:
            width = 640
            height = 480
            region = (0, 0, 640, 480)

            def __init__(self):
                self.is_capturing = False
                self.start_calls = []
                self.latest_calls = 0
                self.grab_calls = 0
                self.stop_calls = 0
                self.release_calls = 0

            def start(self, region=None, target_fps=60, video_mode=False):
                self.start_calls.append((region, target_fps, video_mode))
                self.is_capturing = True

            def get_latest_frame(self, copy=True):
                self.latest_calls += 1
                return screen_capture.np.zeros((8, 8, 3), dtype=screen_capture.np.uint8)

            def grab(self, *args, **kwargs):
                self.grab_calls += 1
                raise AssertionError("DXCam streaming must use the ring buffer, not grab()")

            def stop(self):
                self.stop_calls += 1
                self.is_capturing = False

            def release(self):
                self.release_calls += 1

        fake_dxcam = FakeDXCam()
        capturer = screen_capture.ScreenCapturer(fps=240, quality=54)
        capturer.dxcam_camera = fake_dxcam
        frame = capturer._grab_screen_dxcam()
        if frame is None or fake_dxcam.start_calls != [(None, 240, True)]:
            fail(f"DXCam ring-buffer start was not used correctly: {fake_dxcam.start_calls}")
        if fake_dxcam.latest_calls != 1 or fake_dxcam.grab_calls:
            fail("DXCam capture did not use get_latest_frame() exclusively")
        capturer.fps = 120
        capturer._grab_screen_dxcam()
        if fake_dxcam.stop_calls != 1 or fake_dxcam.start_calls[-1] != (None, 120, True):
            fail("DXCam ring buffer did not restart when target FPS changed")
        capturer._release_dxcam()
        if fake_dxcam.release_calls != 1:
            fail("DXCam release path did not release the camera")

        async def route_checks():
            headers = Headers()
            headers["Host"] = "localhost:6173"
            unauth_response = await server.process_request("/api/public-url", headers)
            if unauth_response.status_code != 403:
                fail(f"unauthenticated public-url returned {unauth_response.status_code}, expected 403")
            snapshot_prefix_response = await server.process_request("/snapshot-extra?fmt=png", headers)
            if snapshot_prefix_response.status_code != 404:
                fail(f"snapshot prefix route returned {snapshot_prefix_response.status_code}, expected 404")
            query_auth_response = await server.process_request(f"/api/auth?code={test_auth_code}", headers)
            if query_auth_response.status_code != 400:
                fail(f"query auth returned {query_auth_response.status_code}, expected 400")
            try:
                os.environ["ZADOO_ALLOW_QUERY_AUTH"] = "1"
                compat_query_server = VNCServer(6176)
                compat_query_response = await compat_query_server.process_request(f"/api/auth?code={test_auth_code}", headers)
                if compat_query_response.status_code != 200:
                    fail(f"compat query auth returned {compat_query_response.status_code}, expected 200")
            finally:
                os.environ.pop("ZADOO_ALLOW_QUERY_AUTH", None)
            legacy_headers = Headers()
            legacy_headers["Host"] = "localhost:6173"
            legacy_headers["X-Zadoo-Code"] = _legacy_auth_strings()[0]
            legacy_response = await server.process_request("/api/auth", legacy_headers)
            if legacy_response.status_code != 401:
                fail(f"legacy hardcoded auth code returned {legacy_response.status_code}, expected 401")
            cross_origin_headers = Headers()
            cross_origin_headers["Host"] = "localhost:6173"
            cross_origin_headers["Origin"] = "https://evil.example"
            cross_origin_response = await server.process_request("/", cross_origin_headers)
            if cross_origin_response.status_code != 403:
                fail(f"cross-origin HTTP request returned {cross_origin_response.status_code}, expected 403")
            auth_request_headers = Headers()
            auth_request_headers["Host"] = "localhost:6173"
            auth_request_headers["Origin"] = "http://localhost:6173"
            auth_request_headers["X-Zadoo-Code"] = test_auth_code
            auth_response = await server.process_request("/api/auth", auth_request_headers)
            if auth_response.status_code != 200:
                fail(f"auth route returned {auth_response.status_code}, expected 200")
            auth_payload = json.loads(auth_response.body.decode("utf-8"))
            csrf_token = auth_payload.get("csrf_token")
            if auth_payload.get("mode") != "full":
                fail(f"auth route returned mode={auth_payload.get('mode')}, expected full")
            cookie = cookie_header_from_set_cookie(auth_response.headers)
            if not cookie or "zadoo_auth=" not in cookie or not csrf_token:
                fail("auth route did not set zadoo_auth cookie")
            headers["Cookie"] = cookie
            query_access_headers = Headers()
            query_access_headers["Host"] = "localhost:6173"
            query_access_headers["X-Zadoo-Code"] = test_auth_code
            query_access_response = await server.process_request("/api/auth?access=lockdown", query_access_headers)
            if query_access_response.status_code != 200:
                fail(f"query access auth returned {query_access_response.status_code}, expected 200")
            query_access_payload = json.loads(query_access_response.body.decode("utf-8"))
            if query_access_payload.get("mode") != "lockdown":
                fail(f"query access auth returned mode={query_access_payload.get('mode')}, expected lockdown")
            ws_bad_origin_headers = Headers()
            ws_bad_origin_headers["Host"] = "localhost:6173"
            ws_bad_origin_headers["Origin"] = "https://evil.example"
            ws_bad_origin_headers["Connection"] = "Upgrade"
            ws_bad_origin_headers["Upgrade"] = "websocket"
            ws_bad_origin_headers["Cookie"] = headers["Cookie"]
            ws_bad_origin = await server.process_request("/video", ws_bad_origin_headers)
            if ws_bad_origin.status_code != 403:
                fail(f"cross-origin websocket returned {ws_bad_origin.status_code}, expected 403")
            throttled_server = VNCServer(6174)
            bad_headers = Headers()
            bad_headers["Host"] = "localhost:6174"
            bad_headers["X-Forwarded-For"] = "203.0.113.10"
            bad_headers["X-Zadoo-Code"] = "BADCODE"
            for _ in range(3):
                await throttled_server.process_request("/api/auth", bad_headers)
            locked_headers = Headers()
            locked_headers["Host"] = "localhost:6174"
            locked_headers["X-Forwarded-For"] = "203.0.113.10"
            locked_headers["X-Zadoo-Code"] = test_auth_code
            locked_response = await throttled_server.process_request("/api/auth", locked_headers)
            if locked_response.status_code != 429:
                fail(f"auth lockout returned {locked_response.status_code}, expected 429")
            unsafe_quality_response = await server.process_request("/api/set-quality?value=80", headers)
            if unsafe_quality_response.status_code != 403:
                fail(f"state-changing route without CSRF returned {unsafe_quality_response.status_code}, expected 403")
            csrf_headers = Headers()
            csrf_headers["Host"] = "localhost:6173"
            csrf_headers["Cookie"] = headers["Cookie"]
            csrf_headers["X-Zadoo-CSRF"] = csrf_token
            checks = {
                "/api/public-url": 200,
                "/api/list-cameras": 200,
                "/api/list-mics": 200,
                "/api/set-quality?value=80": 200,
                "/api/set-fps?value=20": 200,
            }
            for path, expected_status in checks.items():
                request_headers = csrf_headers if path in {
                    "/api/set-quality?value=80",
                    "/api/set-fps?value=20",
                } else headers
                response = await server.process_request(path, request_headers)
                if response is None:
                    fail(f"route returned websocket pass-through unexpectedly: {path}")
                if response.status_code != expected_status:
                    fail(f"route {path} returned {response.status_code}, expected {expected_status}")
                if path == "/api/list-cameras":
                    assert_camera_payload(response.body, require_objects=True)
                elif path == "/api/list-mics":
                    assert_mic_payload(response.body)
                else:
                    payload = json.loads(response.body.decode("utf-8"))
                    if path.startswith("/api/set-quality") and payload.get("quality") != 80:
                        fail(f"set-quality returned {payload.get('quality')}, expected 80")
                    if path.startswith("/api/set-fps") and payload.get("fps") != 20:
                        fail(f"set-fps returned {payload.get('fps')}, expected 20")
            removed_clipboard_response = await server.process_request("/api/set-clipboard-image", csrf_headers)
            if removed_clipboard_response.status_code != 404:
                fail(
                    "removed /api/set-clipboard-image returned "
                    f"{removed_clipboard_response.status_code}, expected 404"
                )
            if server.current_quality != 80:
                fail(f"server current_quality is {server.current_quality}, expected 80")
            if server.current_fps != 20:
                fail(f"server current_fps is {server.current_fps}, expected 20")
            max_fps_response = await server.process_request("/api/set-fps?value=max", csrf_headers)
            if max_fps_response.status_code != 200:
                fail(f"set-fps max returned {max_fps_response.status_code}, expected 200")
            max_fps_payload = json.loads(max_fps_response.body.decode("utf-8"))
            if max_fps_payload.get("fps") != 0 or server.current_fps != 0:
                fail(f"set-fps max returned {max_fps_payload.get('fps')} and server current_fps={server.current_fps}, expected 0")
            quality_100_response = await server.process_request("/api/set-quality?value=100", csrf_headers)
            if quality_100_response.status_code != 200:
                fail(f"set-quality 100 returned {quality_100_response.status_code}, expected 200")
            quality_100_payload = json.loads(quality_100_response.body.decode("utf-8"))
            if quality_100_payload.get("quality") != 100 or server.current_quality != 100:
                fail(f"set-quality 100 returned {quality_100_payload.get('quality')} and server current_quality={server.current_quality}, expected 100")
            if server.current_fps != 0:
                fail(f"set-quality changed current_fps to {server.current_fps}, expected it to remain 0")
            class DummyCapturer:
                def __init__(self):
                    self.fps = 0
                    self.quality = 100
                    self.perf_enabled = None
                    self.perf_region = None
                    self.perf_scale_div = None
                    self.perf_grayscale = None
                def set_performance_mode(self, enabled, region, scale_div):
                    self.perf_enabled = enabled
                    self.perf_region = region
                    self.perf_scale_div = scale_div
                def set_grayscale(self, enabled):
                    self.perf_grayscale = enabled
            dummy = DummyCapturer()
            server.screen_capturer = dummy
            if getattr(server, "adaptive_stream", None):
                server.adaptive_stream.profile_index = 6
                server.adaptive_stream.measured_fps_cap = None
            server._apply_stream_profile("smoke_quality_lock")
            if server.current_quality != 100 or dummy.quality != 100:
                fail("adaptive stream profile changed user-selected quality")
            if dummy.perf_scale_div != 1 or dummy.perf_enabled:
                fail(f"quality 100 did not force full-resolution scale: enabled={dummy.perf_enabled} scale={dummy.perf_scale_div}")
            if getattr(server, "adaptive_stream", None):
                server.adaptive_stream.profile_index = 0
                server.adaptive_stream.measured_fps_cap = None
                server._apply_stream_profile("smoke_fps_cap_start")
            changed = server.adaptive_stream.observe_server(
                frame_bytes=6_000_000,
                max_write_buffer=0,
                skipped_total=0,
                inflight_sends=0,
                video_clients=1,
                frame_age_ms=0,
                capture_stats={
                    "quality": 100,
                    "last_capture_ms": 12.0,
                    "last_encode_ms": 12.0,
                    "current_fps": 40.0,
                },
            )
            server._apply_stream_profile("smoke_fps_cap")
            if not changed or not server.adaptive_stream.measured_fps_cap:
                fail("full-resolution overload did not create a measured FPS cap")
            if dummy.perf_scale_div != 1 or dummy.quality != 100:
                fail("measured FPS cap changed quality or full-resolution scale")
            if dummy.fps >= 120:
                fail(f"measured FPS cap did not reduce target fps: {dummy.fps}")
            server.screen_capturer = None
            stream_response = await server.process_request("/api/stream-stats", headers)
            if stream_response.status_code != 200:
                fail(f"stream-stats returned {stream_response.status_code}, expected 200")
            stream_payload = json.loads(stream_response.body.decode("utf-8"))
            stream = stream_payload.get("stream", {})
            if not stream_payload.get("success") or "stream" not in stream_payload:
                fail("stream-stats payload missing success=true or stream data")
            if not stream.get("measured_fps_cap") or stream.get("fps_cap_reason") != "host_capture_encode":
                fail("stream-stats missing measured FPS cap or host downshift reason")
            benchmark_response = await server.process_request("/benchmark.html", headers)
            if benchmark_response.status_code != 200 or b"Stream Benchmark" not in benchmark_response.body:
                fail("benchmark.html route did not return the benchmark page")

            class FakeRequest:
                def __init__(self, request_headers):
                    self.headers = request_headers

            class FakeWebSocket:
                remote_address = ("smoke", 1)
                def __init__(self, request_headers, messages):
                    self.request = FakeRequest(request_headers)
                    self.sent = []
                    self.messages = list(messages)
                async def send(self, data):
                    self.sent.append(data)
                def __aiter__(self):
                    return self
                async def __anext__(self):
                    if self.messages:
                        return self.messages.pop(0)
                    raise StopAsyncIteration

            view_auth_headers = Headers()
            view_auth_headers["Host"] = "localhost:6173"
            view_auth_headers["X-Zadoo-Code"] = test_auth_code
            view_auth_headers["X-Zadoo-Access"] = "lockdown"
            view_auth_response = await server.process_request("/api/auth", view_auth_headers)
            if view_auth_response.status_code != 200:
                fail(f"view auth returned {view_auth_response.status_code}, expected 200")
            view_auth_payload = json.loads(view_auth_response.body.decode("utf-8"))
            if view_auth_payload.get("mode") != "lockdown":
                fail(f"view auth returned mode={view_auth_payload.get('mode')}, expected lockdown")
            view_ws_headers = Headers()
            view_ws_headers["Host"] = "localhost:6173"
            view_ws_headers["Cookie"] = cookie_header_from_set_cookie(view_auth_response.headers)
            recorded_actions = []
            original_process_event = server.process_event
            try:
                server.process_event = lambda event, websocket=None: recorded_actions.append(event.get("action"))
                view_ws = FakeWebSocket(view_ws_headers, [json.dumps({"action": "click", "x": 0.5, "y": 0.5})])
                await server.video_stream_handler(view_ws)
                if recorded_actions:
                    fail(f"view-only /video processed control actions: {recorded_actions}")
                if not any("Forbidden" in str(item) for item in view_ws.sent):
                    fail("view-only /video did not report forbidden control action")
                full_ws = FakeWebSocket(headers, [json.dumps({"action": "click", "x": 0.5, "y": 0.5})])
                await server.video_stream_handler(full_ws)
                if recorded_actions != ["click"]:
                    fail(f"full /video did not process allowed control action: {recorded_actions}")
            finally:
                server.process_event = original_process_event

            class FakeSendWebSocket:
                remote_address = ("smoke-send", 1)
                def __init__(self):
                    self.sent = []
                async def send(self, data):
                    self.sent.append(data)

            alert_response = await server.process_request("/api/alert?code=A", csrf_headers)
            if alert_response.status_code != 404:
                fail(f"unset alert route returned {alert_response.status_code}, expected 404")

            class FakeHTTPBodyRequest:
                def __init__(self, path, headers, payload):
                    self.path = path
                    self.headers = headers
                    self.body = json.dumps(payload).encode("utf-8")

            async def post_json(path, payload):
                local_headers = Headers()
                local_headers["Host"] = "localhost:6173"
                local_headers["Content-Type"] = "application/json"
                return await server.process_request(object(), FakeHTTPBodyRequest(path, local_headers, payload))

            settings_code = "SMOKE" + secrets.token_hex(2).upper()
            no_permissions = {key: False for key in settings_mod.PERMISSION_KEYS}
            full_permissions = {key: True for key in settings_mod.PERMISSION_KEYS}
            setup_payload = {
                "access_code": settings_code,
                "email_to": "alerts@example.invalid",
                "resend_api_key": "re_smoke_test",
                "permissions": no_permissions,
                "alerts": {
                    "A": {"enabled": True, "title": "Smoke Alert", "message": "Settings-backed alert"},
                    "B": {"enabled": False, "title": "", "message": ""},
                    "C": {"enabled": False, "title": "", "message": ""},
                    "D": {"enabled": False, "title": "", "message": ""},
                },
            }
            setup_response = await post_json("/api/settings/save", setup_payload)
            if setup_response.status_code != 200:
                fail(f"settings save returned {setup_response.status_code}, expected 200: {setup_response.body!r}")
            if not server._settings_store().verify_access_code(settings_code):
                fail("settings access code did not verify after save")
            if server._settings_store().verify_access_code("wrong"):
                fail("settings access code accepted a wrong code")
            if server._settings_store().get_resend_api_key() != "re_smoke_test":
                fail("settings Resend key did not round-trip through storage")
            public_settings = server._settings_store().public_view()
            if public_settings.get("access_code") != settings_code:
                fail("settings access code was not visible in the local settings view")
            server._settings_store().update_cloud_status({
                "workspace_id": "smoke_workspace",
                "device_id": "smoke_device",
                "device_name": "Smoke Device",
                "cloud_api_base": "http://127.0.0.1:9",
                "device_token": "smoke_device_token",
                "entitlement_cache": {
                    "allowed": True,
                    "reason": "",
                    "revoked": False,
                },
            })

            bad_settings_response = await post_json("/api/settings/save", {
                **setup_payload,
                "admin_code": "wrong",
                "access_code": "",
            })
            if bad_settings_response.status_code != 401:
                fail(f"settings edit with wrong admin code returned {bad_settings_response.status_code}, expected 401")

            def auth_with_settings_code():
                auth_headers = Headers()
                auth_headers["Host"] = "localhost:6173"
                auth_headers["X-Zadoo-Code"] = settings_code
                return auth_headers

            denied_auth_response = await server.process_request("/api/auth", auth_with_settings_code())
            if denied_auth_response.status_code != 200:
                fail("settings code authentication failed")
            denied_payload = json.loads(denied_auth_response.body.decode("utf-8"))
            if denied_payload.get("mode") != "custom" or any(denied_payload.get("permissions", {}).values()):
                fail(f"settings auth payload did not use single permission matrix: {denied_payload}")
            denied_ws_headers = Headers()
            denied_ws_headers["Host"] = "localhost:6173"
            denied_ws_headers["Cookie"] = cookie_header_from_set_cookie(denied_auth_response.headers)
            recorded_actions = []
            original_process_event = server.process_event
            try:
                server.process_event = lambda event, websocket=None: recorded_actions.append(event.get("action"))
                denied_ws = FakeWebSocket(denied_ws_headers, [json.dumps({"action": "click", "x": 0.5, "y": 0.5})])
                await server.video_stream_handler(denied_ws)
                if recorded_actions:
                    fail(f"settings disabled permissions processed control actions: {recorded_actions}")
                if not any("Forbidden" in str(item) for item in denied_ws.sent):
                    fail("settings disabled permissions did not report forbidden control action")
            finally:
                server.process_event = original_process_event

            denied_terminal_response = await server.process_request("/terminal.html", denied_ws_headers)
            if denied_terminal_response.status_code != 403:
                fail(f"disabled terminal route returned {denied_terminal_response.status_code}, expected 403")
            if server._is_ws_authorized("/ssh", denied_ws_headers):
                fail("disabled terminal websocket was authorized")

            allow_response = await post_json("/api/settings/save", {
                **setup_payload,
                "admin_code": settings_code,
                "access_code": settings_code,
                "permissions": full_permissions,
            })
            if allow_response.status_code != 200:
                fail(f"settings permission update returned {allow_response.status_code}, expected 200")
            allowed_auth_response = await server.process_request("/api/auth", auth_with_settings_code())
            allowed_payload = json.loads(allowed_auth_response.body.decode("utf-8"))
            if allowed_auth_response.status_code != 200 or not all(allowed_payload.get("permissions", {}).values()):
                fail(f"settings allowed auth payload was wrong: {allowed_payload}")
            allowed_cookie = cookie_header_from_set_cookie(allowed_auth_response.headers)
            allowed_csrf = allowed_payload.get("csrf_token")
            allowed_csrf_headers = Headers()
            allowed_csrf_headers["Host"] = "localhost:6173"
            allowed_csrf_headers["Cookie"] = allowed_cookie
            allowed_csrf_headers["X-Zadoo-CSRF"] = allowed_csrf

            allowed_ws_headers = Headers()
            allowed_ws_headers["Host"] = "localhost:6173"
            allowed_ws_headers["Cookie"] = allowed_cookie
            allowed_terminal_response = await server.process_request("/terminal.html", allowed_ws_headers)
            if allowed_terminal_response.status_code != 200:
                fail(f"enabled terminal route returned {allowed_terminal_response.status_code}, expected 200")
            if not server._is_ws_authorized("/ssh", allowed_ws_headers):
                fail("enabled terminal websocket was not authorized")

            recorded_actions = []
            original_process_event = server.process_event
            try:
                server.process_event = lambda event, websocket=None: recorded_actions.append(event.get("action"))
                allowed_ws = FakeWebSocket(allowed_ws_headers, [json.dumps({"action": "click", "x": 0.5, "y": 0.5})])
                await server.video_stream_handler(allowed_ws)
                if recorded_actions != ["click"]:
                    fail(f"settings enabled permissions did not process allowed control action: {recorded_actions}")
            finally:
                server.process_event = original_process_event

            alert_video_ws = FakeSendWebSocket()
            alert_input_ws = FakeSendWebSocket()
            server.video_clients = {alert_video_ws}
            server.input_clients = {alert_input_ws}
            alert_response = await server.process_request("/api/alert?code=A", allowed_csrf_headers)
            if alert_response.status_code != 200:
                fail(f"settings alert route returned {alert_response.status_code}, expected 200")
            await asyncio.sleep(0.05)
            for name, ws in {"video": alert_video_ws, "input": alert_input_ws}.items():
                if not ws.sent:
                    fail(f"settings alert route did not send to {name} websocket")
                alert_payload = json.loads(ws.sent[-1])
                if alert_payload.get("type") != "controller_alert" or not alert_payload.get("id"):
                    fail(f"settings alert payload for {name} websocket was invalid: {alert_payload}")
            server.video_clients = set()
            server.input_clients = set()

            image_ws = FakeSendWebSocket()
            too_large_png = base64.b64encode(b"123456789").decode("ascii")
            server._handle_set_clipboard_image({"mime": "image/png", "data_base64": too_large_png}, image_ws)
            await asyncio.sleep(0.05)
            if not image_ws.sent:
                fail("clipboard image limit did not send a result")
            image_payload = json.loads(image_ws.sent[-1])
            if image_payload.get("success") is not False or "size limit" not in image_payload.get("error", ""):
                fail(f"clipboard image limit returned unexpected payload: {image_payload}")

            mic_a, mic_b = object(), object()
            server.mic_clients = {mic_a, mic_b}
            stop_calls = []
            original_stop_mic = server._stop_mic_capture
            try:
                server._stop_mic_capture = lambda: stop_calls.append("stop")
                if server._remove_mic_client(mic_a):
                    fail("mic capture stopped while another mic client remained")
                if stop_calls:
                    fail("mic stop called before last client disconnected")
                if not server._remove_mic_client(mic_b) or stop_calls != ["stop"]:
                    fail("mic capture did not stop after last client disconnected")
            finally:
                server._stop_mic_capture = original_stop_mic

        asyncio.run(route_checks())
    finally:
        for key, value in old_env.items():
            if value is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = value
        try:
            settings_mod._STORE = None
        except Exception:
            pass
        temp_settings.cleanup()

    ok("imports, dependency flags, templates, assets, server instantiation, and in-process routes")


def assert_source_clean() -> None:
    for path in source_files():
        text = path.read_text(encoding="utf-8", errors="ignore")
        for forbidden in FORBIDDEN_SOURCE_STRINGS:
            if forbidden in text:
                fail(f"forbidden source string {forbidden!r} found in {path.relative_to(ROOT)}")
        match = MOJIBAKE_RE.search(text)
        if match:
            fail(f"mojibake marker {match.group(0)!r} found in {path.relative_to(ROOT)}")
    ok("source has no forbidden secrets or mojibake markers")


def fetch(base_url: str, path: str, headers: dict[str, str] | None = None):
    url = base_url.rstrip("/") + path
    try:
        request = urllib.request.Request(url, headers=headers or {})
        with urllib.request.urlopen(request, timeout=8) as response:
            return response.status, response.headers.get("Content-Type", ""), response.read()
    except urllib.error.HTTPError as e:
        return e.code, e.headers.get("Content-Type", ""), e.read()


def assert_live(base_url: str) -> None:
    status, content_type, body = fetch(base_url, "/")
    if status == 403:
        fail(
            "live root returned 403 — direct (non-tunnel) access is blocked by default. "
            "Start the target server with ZADOO_ALLOW_DIRECT_ACCESS=1 to run live localhost "
            "smoke tests, e.g.  set ZADOO_ALLOW_DIRECT_ACCESS=1 && python zadoo_vnc_single.py"
        )
    if status != 200 or "text/html" not in content_type:
        fail(f"live root failed: status={status} content_type={content_type}")

    for path in ("/brand-header.png", "/splash.png", "/trigger-icon.png"):
        status, content_type, body = fetch(base_url, path)
        if status != 200 or "image/png" not in content_type or not body:
            fail(f"live asset failed: {path} status={status} content_type={content_type}")

    status, content_type, body = fetch(base_url, "/api/public-url")
    if status != 403:
        fail(f"live unauthenticated public-url returned status={status}, expected 403")

    live_auth_code = os.environ.get("ZADOO_SMOKE_AUTH_CODE") or os.environ.get("ZADOO_ACCESS_CODE")
    if not live_auth_code:
        fail("live checks require ZADOO_ACCESS_CODE or ZADOO_SMOKE_AUTH_CODE because no hardcoded auth default exists")

    legacy_url = base_url.rstrip("/") + "/api/auth"
    legacy_request = urllib.request.Request(legacy_url, headers={"X-Zadoo-Code": _legacy_auth_strings()[0]})
    try:
        with urllib.request.urlopen(legacy_request, timeout=8) as response:
            legacy_status = response.status
            response.read()
    except urllib.error.HTTPError as e:
        legacy_status = e.code
    if legacy_status != 401:
        fail(f"live legacy auth returned status={legacy_status}, expected 401")

    auth_url = base_url.rstrip("/") + "/api/auth"
    live_auth_headers = {"X-Zadoo-Code": live_auth_code}
    auth_request = urllib.request.Request(auth_url, headers=live_auth_headers)
    try:
        with urllib.request.urlopen(auth_request, timeout=8) as response:
            auth_status = response.status
            auth_cookies = response.headers.get_all("Set-Cookie") or []
            auth_cookie = "; ".join(
                cookie.split(";", 1)[0].strip()
                for cookie in auth_cookies
                if cookie.split(";", 1)[0].strip()
            )
            auth_body = response.read()
            try:
                live_csrf = json.loads(auth_body.decode("utf-8")).get("csrf_token")
            except Exception:
                live_csrf = None
    except urllib.error.HTTPError as e:
        auth_status = e.code
        auth_cookies = e.headers.get_all("Set-Cookie") or []
        auth_cookie = "; ".join(
            cookie.split(";", 1)[0].strip()
            for cookie in auth_cookies
            if cookie.split(";", 1)[0].strip()
        )
        live_csrf = None
    if auth_status != 200 or "zadoo_auth=" not in auth_cookie:
        fail(f"live auth failed: status={auth_status}")
    auth_headers = {"Cookie": auth_cookie.split(";", 1)[0]}

    status, content_type, body = fetch(base_url, "/api/list-cameras", auth_headers)
    if status != 200:
        fail(f"live list-cameras failed: status={status}")
    assert_camera_payload(body, require_objects=True)

    status, content_type, body = fetch(base_url, "/api/list-mics", auth_headers)
    if status != 200:
        fail(f"live list-mics failed: status={status}")
    assert_mic_payload(body)

    status, content_type, body = fetch(base_url, "/api/public-url", auth_headers)
    if status != 200:
        fail(f"live public-url failed: status={status}. Restart the app to load the updated routes.")
    json.loads(body.decode("utf-8"))

    status, content_type, body = fetch(base_url, "/api/stream-stats", auth_headers)
    if status != 200:
        fail(f"live stream-stats failed: status={status}. Restart the app to load the updated routes.")
    json.loads(body.decode("utf-8"))

    status, content_type, body = fetch(base_url, "/benchmark.html", auth_headers)
    if status != 200 or "text/html" not in content_type or b"Stream Benchmark" not in body:
        fail(f"live benchmark route failed: status={status} content_type={content_type}")

    ok(f"live HTTP checks passed for {base_url}")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--live", help="Base URL for live HTTP route checks, for example http://localhost:6173")
    args = parser.parse_args()

    assert_imports()
    assert_source_clean()
    if args.live:
        assert_live(args.live)


if __name__ == "__main__":
    main()
