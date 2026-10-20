"""Real HTTP, media, browser and shutdown checks against an isolated Windows host."""
from __future__ import annotations

import argparse
import json
import os
import socket
import subprocess
import sys
import tempfile
import time
import urllib.error
import urllib.request
from pathlib import Path

from websockets.sync.client import connect

ROOT = Path(__file__).resolve().parents[1]
BASE = "http://127.0.0.1:6173"
CODE = "FREE-E2E"


def request(path, headers=None):
    req = urllib.request.Request(BASE + path, headers=headers or {})
    return urllib.request.urlopen(req, timeout=15)


def wait_ready(process):
    deadline = time.monotonic() + 60
    while time.monotonic() < deadline:
        if process.poll() is not None:
            raise RuntimeError(f"Test host exited before startup: {process.returncode}")
        try:
            with request("/") as response:
                if response.status == 200:
                    return
        except OSError:
            time.sleep(0.1)
    raise TimeoutError("Test host did not listen within 60 seconds")


def media_check(path, headers):
    with connect(BASE.replace("http", "ws") + path, origin=BASE, additional_headers=headers, open_timeout=15) as stream:
        header = stream.recv(timeout=15)
        metadata = json.loads(header)
        if metadata.get("type") != "audio_format":
            raise AssertionError(f"{path} returned an error: {metadata}")
        chunk = stream.recv(timeout=15)
        if not isinstance(chunk, bytes) or not chunk:
            raise AssertionError(f"{path} did not produce PCM audio")
        return len(chunk)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--exe", type=Path, help="Packaged executable; omitted for source")
    parser.add_argument("--runtime-python", default=sys.executable)
    parser.add_argument("--hardware", action="store_true", help="Briefly test microphone, speaker loopback and camera without saving media")
    args = parser.parse_args()
    with socket.socket() as probe:
        if probe.connect_ex(("127.0.0.1", 6173)) == 0:
            raise RuntimeError("E2E requires port 6173 to be free; existing host was not touched")
    with tempfile.TemporaryDirectory(prefix="zadoo-free-e2e-") as folder:
        env = {**os.environ, "ZADOO_SETTINGS_DIR": folder, "ZADOO_DISABLE_TUNNEL": "1",
               "ZADOO_ALLOW_DIRECT_ACCESS": "1", "ZADOO_SMOKE_AUTH_CODE": CODE,
               "HIDE_CONSOLE": "1", "PYTHONPATH": str(ROOT)}
        os.environ["ZADOO_SETTINGS_DIR"] = folder
        sys.path.insert(0, str(ROOT))
        from zadoo_vnc.config import windows_system_executable
        from zadoo_vnc.settings import PERMISSION_KEYS, SettingsStore
        store = SettingsStore(Path(folder) / "config.json")
        store.apply_setup({"access_code": CODE, "permissions": dict.fromkeys(PERMISSION_KEYS, True), "autostart_enabled": False})
        if store.get_device_token():
            raise AssertionError("E2E must not have a linked account")
        command = [str(args.exe.resolve())] if args.exe else [args.runtime_python, "-m", "zadoo_vnc"]
        process = subprocess.Popen(command, cwd=ROOT, env=env, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        result = {"mode": "packaged" if args.exe else "source", "account": "not linked"}
        try:
            wait_ready(process)
            with request("/api/auth", {"X-Zadoo-Code": CODE}) as auth:
                if not json.loads(auth.read())["success"]:
                    raise AssertionError("Access-code authentication failed")
                cookie = "; ".join(value.split(";", 1)[0] for value in auth.headers.get_all("Set-Cookie"))
            headers = {"Cookie": cookie}
            with connect(BASE.replace("http", "ws") + "/video", origin=BASE, additional_headers=headers, open_timeout=15) as stream:
                deadline = time.monotonic() + 15
                while time.monotonic() < deadline:
                    frame = stream.recv(timeout=15)
                    if isinstance(frame, bytes):
                        if not frame.startswith(bytes.fromhex("ffd8")) or not frame.endswith(bytes.fromhex("ffd9")):
                            raise AssertionError("Screen stream returned an invalid JPEG")
                        result["screen_jpeg_bytes"] = len(frame)
                        break
                else:
                    raise TimeoutError("Screen stream did not produce a JPEG")
            with request("/snapshot?fmt=png", headers) as response:
                image = response.read()
                if not image.startswith(bytes.fromhex("89504e470d0a1a0a")) or not image.endswith(bytes.fromhex("49454e44ae426082")):
                    raise AssertionError("Snapshot returned an invalid PNG")
                result["snapshot_png_bytes"] = len(image)
            for route in ("/api/local/credits", "/api/local/topup-order", "/api/local/topup-verify"):
                try:
                    with request(route, headers):
                        raise AssertionError(f"Retired route still exists: {route}")
                except urllib.error.HTTPError as exc:
                    if exc.code != 404:
                        raise
            if args.hardware:
                result["system_audio_pcm_bytes"] = media_check("/audio", headers)
                result["microphone_pcm_bytes"] = media_check("/mic", headers)
                with request("/api/list-cameras", headers) as response:
                    cameras = json.loads(response.read())["devices"]
                if not cameras:
                    result["camera"] = "No camera attached"
                else:
                    with connect(BASE.replace("http", "ws") + "/webcam", origin=BASE, additional_headers=headers, open_timeout=15) as stream:
                        stream.send(json.dumps({"type": "select_camera", "camera_id": cameras[0]["id"]}))
                        frame = stream.recv(timeout=20)
                        if not isinstance(frame, bytes) or not frame.startswith(bytes.fromhex("ffd8")):
                            raise AssertionError(f"Camera did not return a JPEG: {frame!r}")
                        result["camera_jpeg_bytes"] = len(frame)
            subprocess.run([sys.executable, str(ROOT / "scripts/smoke_test.py"), "--live", BASE], cwd=ROOT, env=env, check=True)
            subprocess.run([sys.executable, str(ROOT / "scripts/terminal_e2e_test.py"), "--code", CODE], cwd=ROOT, env=env, check=True)
            with connect(BASE.replace("http", "ws") + "/rpc", origin=BASE, open_timeout=10) as rpc:
                rpc.send(json.dumps({"id": "free-e2e-stop", "action": "runtime.stop", "params": {"admin_code": CODE}}))
                if not json.loads(rpc.recv(timeout=10))["success"]:
                    raise AssertionError("Runtime stop failed")
            process.wait(timeout=20)
            if process.returncode != 0:
                raise RuntimeError(f"Runtime exited with code {process.returncode}")
            with socket.socket() as probe:
                if probe.connect_ex(("127.0.0.1", 6173)) == 0:
                    raise AssertionError("Runtime did not release its port")
            result["browser_live_smoke_clean_shutdown"] = "pass"
        finally:
            if process.poll() is None:
                subprocess.run([windows_system_executable("taskkill.exe"), "/PID", str(process.pid), "/T", "/F"], check=False, capture_output=True)
                process.wait(timeout=15)
        print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()

