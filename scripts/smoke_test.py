"""Smoke checks for the Zadoo VNC runtime."""
from __future__ import annotations

import argparse
import asyncio
import json
import re
import sys
import urllib.error
import urllib.request
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SOURCE_GLOBS = ("zadoo_vnc/**/*.py", "zadoo_vnc/templates/*.html", "*.py", "*.md", "*.toml", "*.txt")
MOJIBAKE_RE = re.compile(r"[\u00c2\u00c3\u00e2\u00f0][^\x00-\x7f]+")
FORBIDDEN_SOURCE_STRINGS = (
    "re_GF7",
    "resend_api_key_default",
    "email_to_default",
    "iskssj07@gmail.com",
)


def fail(message: str) -> None:
    print(f"FAIL: {message}")
    raise SystemExit(1)


def ok(message: str) -> None:
    print(f"OK: {message}")


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


def source_files():
    seen = set()
    for glob in SOURCE_GLOBS:
        for path in ROOT.glob(glob):
            if path.is_file() and path not in seen:
                seen.add(path)
                yield path


def assert_imports() -> None:
    sys.path.insert(0, str(ROOT))
    import zadoo_vnc.config as config
    import zadoo_vnc.dependencies as deps
    from zadoo_vnc.assets import load_benchmark_html, load_host_controls_html, load_index_html, load_terminal_html
    from zadoo_vnc.server import VNCServer
    from websockets.datastructures import Headers

    if not (deps.HAS_PYAUTOGUI and (deps.HAS_MSS or deps.WIN32_AVAILABLE or deps.HAS_PIL)):
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

    server = VNCServer(6173, 0)
    if server.port != 6173 or server.secondary_port != 0:
        fail("VNCServer did not instantiate with expected ports")

    async def route_checks():
        headers = Headers()
        unauth_response = await server.process_request("/api/public-url", headers)
        if unauth_response.status_code != 403:
            fail(f"unauthenticated public-url returned {unauth_response.status_code}, expected 403")
        auth_response = await server.process_request("/api/auth?code=TERMINATOR", headers)
        if auth_response.status_code != 200:
            fail(f"auth route returned {auth_response.status_code}, expected 200")
        cookie = auth_response.headers.get("Set-Cookie")
        if not cookie or "zadoo_auth=" not in cookie:
            fail("auth route did not set zadoo_auth cookie")
        headers["Cookie"] = cookie.split(";", 1)[0]
        checks = {
            "/api/public-url": 200,
            "/api/list-cameras": 200,
            "/api/set-quality?value=80": 200,
            "/api/set-fps?value=20": 200,
            "/api/set-clipboard-image": 400,
        }
        for path, expected_status in checks.items():
            response = await server.process_request(path, headers)
            if response is None:
                fail(f"route returned websocket pass-through unexpectedly: {path}")
            if response.status_code != expected_status:
                fail(f"route {path} returned {response.status_code}, expected {expected_status}")
            if path == "/api/list-cameras":
                assert_camera_payload(response.body, require_objects=True)
            else:
                payload = json.loads(response.body.decode("utf-8"))
                if path.startswith("/api/set-quality") and payload.get("quality") != 80:
                    fail(f"set-quality returned {payload.get('quality')}, expected 80")
                if path.startswith("/api/set-fps") and payload.get("fps") != 20:
                    fail(f"set-fps returned {payload.get('fps')}, expected 20")
        if server.current_quality != 80:
            fail(f"server current_quality is {server.current_quality}, expected 80")
        if server.current_fps != 20:
            fail(f"server current_fps is {server.current_fps}, expected 20")
        max_fps_response = await server.process_request("/api/set-fps?value=max", headers)
        if max_fps_response.status_code != 200:
            fail(f"set-fps max returned {max_fps_response.status_code}, expected 200")
        max_fps_payload = json.loads(max_fps_response.body.decode("utf-8"))
        if max_fps_payload.get("fps") != 0 or server.current_fps != 0:
            fail(f"set-fps max returned {max_fps_payload.get('fps')} and server current_fps={server.current_fps}, expected 0")
        quality_100_response = await server.process_request("/api/set-quality?value=100", headers)
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
        server._apply_stream_profile("smoke_quality_lock")
        if server.current_quality != 100 or dummy.quality != 100:
            fail("adaptive stream profile changed user-selected quality")
        if dummy.perf_scale_div != 1 or dummy.perf_enabled:
            fail(f"quality 100 did not force full-resolution scale: enabled={dummy.perf_enabled} scale={dummy.perf_scale_div}")
        server.screen_capturer = None
        stream_response = await server.process_request("/api/stream-stats", headers)
        if stream_response.status_code != 200:
            fail(f"stream-stats returned {stream_response.status_code}, expected 200")
        stream_payload = json.loads(stream_response.body.decode("utf-8"))
        if not stream_payload.get("success") or "stream" not in stream_payload:
            fail("stream-stats payload missing success=true or stream data")
        benchmark_response = await server.process_request("/benchmark.html", headers)
        if benchmark_response.status_code != 200 or b"Stream Benchmark" not in benchmark_response.body:
            fail("benchmark.html route did not return the benchmark page")

    asyncio.run(route_checks())

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
    if status != 200 or "text/html" not in content_type:
        fail(f"live root failed: status={status} content_type={content_type}")

    for path in ("/brand-header.png", "/splash.png", "/trigger-icon.png"):
        status, content_type, body = fetch(base_url, path)
        if status != 200 or "image/png" not in content_type or not body:
            fail(f"live asset failed: {path} status={status} content_type={content_type}")

    status, content_type, body = fetch(base_url, "/api/public-url")
    if status != 403:
        fail(f"live unauthenticated public-url returned status={status}, expected 403")

    auth_url = base_url.rstrip("/") + "/api/auth?code=TERMINATOR"
    auth_request = urllib.request.Request(auth_url)
    try:
        with urllib.request.urlopen(auth_request, timeout=8) as response:
            auth_status = response.status
            auth_cookie = response.headers.get("Set-Cookie", "")
            response.read()
    except urllib.error.HTTPError as e:
        auth_status = e.code
        auth_cookie = e.headers.get("Set-Cookie", "")
    if auth_status != 200 or "zadoo_auth=" not in auth_cookie:
        fail(f"live auth failed: status={auth_status}")
    auth_headers = {"Cookie": auth_cookie.split(";", 1)[0]}

    status, content_type, body = fetch(base_url, "/api/list-cameras", auth_headers)
    if status != 200:
        fail(f"live list-cameras failed: status={status}")
    assert_camera_payload(body, require_objects=True)

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
