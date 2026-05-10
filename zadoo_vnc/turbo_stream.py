"""Generic Windows low-latency WebRTC streaming engine.

This module keeps the current JPEG WebSocket stream as a compatibility
fallback while adding a reusable FFmpeg + MediaMTX path for ordinary Windows
laptops. It intentionally selects capture and encoder support by probing and
benchmarking the current PC instead of assuming a specific GPU vendor.
"""
from __future__ import annotations

import os
import platform
import re
import shutil
import socket
import subprocess
import threading
import time
import urllib.parse
from dataclasses import asdict, dataclass
from pathlib import Path


FALSE_VALUES = {"0", "false", "no", "off", "disabled"}
TRUE_VALUES = {"1", "true", "yes", "on", "enabled"}


@dataclass(frozen=True)
class TurboProfile:
    name: str
    width: int
    height: int
    fps: int
    bitrate_kbps: int
    maxrate_kbps: int
    status: str

    @property
    def bitrate(self) -> str:
        return f"{self.bitrate_kbps}k"

    @property
    def maxrate(self) -> str:
        return f"{self.maxrate_kbps}k"

    @property
    def bufsize(self) -> str:
        return f"{max(self.maxrate_kbps * 2, self.bitrate_kbps * 2)}k"


@dataclass(frozen=True)
class CaptureMethod:
    name: str
    label: str
    kind: str
    available: bool
    reason: str = ""


@dataclass(frozen=True)
class EncoderMethod:
    name: str
    label: str
    available: bool
    reason: str = ""


@dataclass(frozen=True)
class StreamSelection:
    capture: CaptureMethod
    encoder: EncoderMethod
    profile: TurboProfile
    benchmark_ms: float

    def as_dict(self) -> dict:
        return {
            "capture": asdict(self.capture),
            "encoder": asdict(self.encoder),
            "profile": asdict(self.profile),
            "benchmark_ms": round(float(self.benchmark_ms), 1),
        }


TURBO_PROFILES = (
    TurboProfile("540p30", 960, 540, 30, 4000, 6000, "Turbo 540p30"),
    TurboProfile("540p60", 960, 540, 60, 6000, 8000, "Turbo 540p60"),
    TurboProfile("720p30", 1280, 720, 30, 5000, 7000, "Turbo 720p30"),
    TurboProfile("360p30", 640, 360, 30, 1800, 2500, "Turbo 360p30"),
    TurboProfile("360p24", 640, 360, 24, 1200, 2000, "Turbo 360p24"),
)

CAPTURE_PRIORITY = ("ddagrab", "gfxcapture", "gdigrab")
ENCODER_PRIORITY = ("h264_mf", "h264_qsv", "h264_amf", "h264_nvenc", "libx264")


def _env_flag(name: str, default: bool = False) -> bool:
    value = os.environ.get(name)
    if value is None or str(value).strip() == "":
        return default
    normalized = str(value).strip().lower()
    if normalized in TRUE_VALUES:
        return True
    if normalized in FALSE_VALUES:
        return False
    return default


def _hidden_creationflags() -> int:
    if platform.system().lower() != "windows":
        return 0
    return getattr(subprocess, "CREATE_NO_WINDOW", 0)


def _repo_root() -> Path:
    return Path(__file__).resolve().parent.parent


def _split_env_paths(value: str | None) -> list[Path]:
    paths: list[Path] = []
    for chunk in str(value or "").split(os.pathsep):
        cleaned = chunk.strip().strip('"')
        if cleaned:
            paths.append(Path(cleaned))
    return paths


def _which_or_none(name: str) -> str | None:
    try:
        found = shutil.which(name)
        return str(Path(found).resolve()) if found else None
    except Exception:
        return None


def _candidate_paths(env_name: str, exe_name: str) -> list[Path]:
    root = _repo_root()
    candidates: list[Path] = []
    candidates.extend(_split_env_paths(os.getenv(env_name)))
    candidates.extend(
        [
            root / exe_name,
            root / "bin" / exe_name,
            root / "tools" / exe_name,
            root / "vendor" / exe_name,
            root / "vendor" / exe_name.replace(".exe", "") / exe_name,
        ]
    )
    path_hit = _which_or_none(exe_name)
    if path_hit:
        candidates.append(Path(path_hit))
    return candidates


def _find_executable(env_name: str, exe_name: str) -> str | None:
    for candidate in _candidate_paths(env_name, exe_name):
        try:
            if candidate.is_dir():
                candidate = candidate / exe_name
            if candidate.is_file():
                return str(candidate.resolve())
        except Exception:
            continue
    return None


def _run_tool(args: list[str], timeout: float = 5.0) -> tuple[int, str]:
    try:
        result = subprocess.run(
            args,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
            creationflags=_hidden_creationflags(),
        )
        return int(result.returncode), (result.stdout or "") + "\n" + (result.stderr or "")
    except Exception as exc:
        return 999, str(exc)


def _has_ffmpeg_name(text: str, name: str) -> bool:
    pattern = re.compile(rf"(^|\s){re.escape(name)}(\s|$)", re.IGNORECASE)
    return any(pattern.search(line) for line in str(text or "").splitlines())


def _tcp_open(host: str, port: int, timeout: float = 0.35) -> bool:
    try:
        with socket.create_connection((host, int(port)), timeout=timeout):
            return True
    except Exception:
        return False


class WindowsTurboStream:
    """Probe, benchmark, and run a generic Windows FFmpeg WebRTC stream."""

    def __init__(self) -> None:
        self.enabled = _env_flag("ZADOO_TURBO_STREAM", True)
        self.auto_start_enabled = _env_flag("ZADOO_TURBO_AUTO", True)
        self.benchmark_enabled = _env_flag("ZADOO_TURBO_BENCHMARK", True)
        self.allow_cpu_60 = _env_flag("ZADOO_TURBO_CPU_60", False)
        self.stream_name = os.getenv("ZADOO_TURBO_STREAM_NAME", "zadoo").strip() or "zadoo"
        self.display_index = self._int_env("ZADOO_TURBO_DISPLAY_INDEX", 0, 0, 16)
        self.webrtc_port = self._int_env("ZADOO_MEDIAMTX_WEBRTC_PORT", 8889, 1, 65535)
        self.benchmark_seconds = self._float_env("ZADOO_TURBO_BENCH_SECONDS", 1.8, 0.5, 8.0)
        self.force_profile_name = os.getenv("ZADOO_TURBO_PROFILE", "").strip().lower()
        self.force_capture_name = os.getenv("ZADOO_TURBO_CAPTURE", "").strip().lower()
        self.force_encoder_name = os.getenv("ZADOO_TURBO_ENCODER", "").strip().lower()
        self.public_base_url = os.getenv("ZADOO_TURBO_PUBLIC_URL", "").strip()

        self.ffmpeg_path: str | None = None
        self.mediamtx_path: str | None = None
        self.capabilities: dict = {}
        self.capture_methods: list[CaptureMethod] = []
        self.encoder_methods: list[EncoderMethod] = []
        self.has_whip = False
        self.has_opus = False
        self.reason = ""

        self._probe_lock = threading.RLock()
        self._start_lock = threading.RLock()
        self._probe_done = False
        self._last_probe_at = 0.0
        self._selection: StreamSelection | None = None
        self._benchmark_results: list[dict] = []
        self._last_error = ""
        self._ffmpeg_proc: subprocess.Popen | None = None
        self._mediamtx_proc: subprocess.Popen | None = None
        self._mediamtx_owned = False
        self._last_started_at = 0.0

    @staticmethod
    def _int_env(name: str, default: int, min_value: int, max_value: int) -> int:
        try:
            value = int(str(os.getenv(name, default)).strip())
            return max(min_value, min(max_value, value))
        except Exception:
            return default

    @staticmethod
    def _float_env(name: str, default: float, min_value: float, max_value: float) -> float:
        try:
            value = float(str(os.getenv(name, default)).strip())
            return max(min_value, min(max_value, value))
        except Exception:
            return default

    def probe(self, force: bool = False) -> dict:
        with self._probe_lock:
            if self._probe_done and not force:
                return dict(self.capabilities)

            self._last_probe_at = time.time()
            platform_name = platform.system() or "unknown"
            self.ffmpeg_path = _find_executable("ZADOO_FFMPEG_PATH", "ffmpeg.exe")
            self.mediamtx_path = _find_executable("ZADOO_MEDIAMTX_PATH", "mediamtx.exe")
            self.capture_methods = []
            self.encoder_methods = []
            self.has_whip = False
            self.has_opus = False
            filters_text = ""
            devices_text = ""
            encoders_text = ""
            muxers_text = ""

            if self.ffmpeg_path:
                _, filters_text = _run_tool([self.ffmpeg_path, "-hide_banner", "-filters"], timeout=5.0)
                _, devices_text = _run_tool([self.ffmpeg_path, "-hide_banner", "-devices"], timeout=5.0)
                _, encoders_text = _run_tool([self.ffmpeg_path, "-hide_banner", "-encoders"], timeout=5.0)
                _, muxers_text = _run_tool([self.ffmpeg_path, "-hide_banner", "-muxers"], timeout=5.0)

            has_ddagrab = _has_ffmpeg_name(filters_text, "ddagrab")
            has_gfxcapture = _has_ffmpeg_name(filters_text, "gfxcapture")
            has_gdigrab = _has_ffmpeg_name(devices_text, "gdigrab")
            self.has_whip = _has_ffmpeg_name(muxers_text, "whip")
            self.has_opus = _has_ffmpeg_name(encoders_text, "libopus") or _has_ffmpeg_name(encoders_text, "opus")

            self.capture_methods = [
                CaptureMethod("ddagrab", "Desktop Duplication (GPU capture)", "lavfi", has_ddagrab,
                              "" if has_ddagrab else "FFmpeg ddagrab filter is not available."),
                CaptureMethod("gfxcapture", "Windows Graphics Capture", "lavfi", has_gfxcapture,
                              "" if has_gfxcapture else "FFmpeg gfxcapture filter is not available."),
                CaptureMethod("gdigrab", "GDI desktop capture", "device", has_gdigrab,
                              "" if has_gdigrab else "FFmpeg gdigrab input device is not available."),
            ]
            self.capture_methods = self._ordered_capture_methods(self.capture_methods)

            encoder_labels = {
                "h264_mf": "Media Foundation H.264",
                "h264_qsv": "Intel Quick Sync H.264",
                "h264_amf": "AMD AMF H.264",
                "h264_nvenc": "NVIDIA NVENC H.264",
                "libx264": "CPU x264 H.264",
            }
            self.encoder_methods = [
                EncoderMethod(name, encoder_labels[name], _has_ffmpeg_name(encoders_text, name),
                              "" if _has_ffmpeg_name(encoders_text, name) else f"FFmpeg encoder {name} is not available.")
                for name in ENCODER_PRIORITY
            ]
            self.encoder_methods = self._ordered_encoder_methods(self.encoder_methods)

            available_captures = [item.name for item in self.capture_methods if item.available]
            available_encoders = [item.name for item in self.encoder_methods if item.available]
            reasons: list[str] = []
            if not self.enabled:
                reasons.append("Turbo WebRTC is disabled by ZADOO_TURBO_STREAM=0.")
            if platform_name.lower() != "windows":
                reasons.append("Turbo capture is Windows-only; JPEG fallback remains active.")
            if not self.ffmpeg_path:
                reasons.append("ffmpeg.exe was not found. Set ZADOO_FFMPEG_PATH or place ffmpeg.exe in the project, bin, or tools folder.")
            if not self.mediamtx_path:
                reasons.append("mediamtx.exe was not found. Set ZADOO_MEDIAMTX_PATH or place mediamtx.exe in the project, bin, or tools folder.")
            if self.ffmpeg_path and not self.has_whip:
                reasons.append("This FFmpeg build does not expose the WHIP muxer required for WebRTC publishing.")
            if self.ffmpeg_path and not available_captures:
                reasons.append("No FFmpeg Windows desktop capture source is available.")
            if self.ffmpeg_path and not available_encoders:
                reasons.append("No usable FFmpeg H.264 encoder is available.")
            self.reason = " ".join(reasons)

            available = (
                self.enabled
                and platform_name.lower() == "windows"
                and bool(self.ffmpeg_path)
                and bool(self.mediamtx_path)
                and self.has_whip
                and bool(available_captures)
                and bool(available_encoders)
            )
            self.capabilities = {
                "enabled": self.enabled,
                "available": available,
                "platform": platform_name,
                "ffmpeg_path": self.ffmpeg_path,
                "mediamtx_path": self.mediamtx_path,
                "has_whip": self.has_whip,
                "has_opus": self.has_opus,
                "capture_priority": list(CAPTURE_PRIORITY),
                "encoder_priority": list(ENCODER_PRIORITY),
                "capture_methods": [asdict(item) for item in self.capture_methods],
                "encoders": [asdict(item) for item in self.encoder_methods],
                "available_captures": available_captures,
                "available_encoders": available_encoders,
                "reason": self.reason,
                "last_probe_at": self._last_probe_at,
            }
            self._probe_done = True
            return dict(self.capabilities)

    def _ordered_capture_methods(self, methods: list[CaptureMethod]) -> list[CaptureMethod]:
        if not self.force_capture_name:
            return methods
        preferred = [item for item in methods if item.name == self.force_capture_name]
        rest = [item for item in methods if item.name != self.force_capture_name]
        return preferred + rest

    def _ordered_encoder_methods(self, methods: list[EncoderMethod]) -> list[EncoderMethod]:
        if not self.force_encoder_name:
            return methods
        preferred = [item for item in methods if item.name == self.force_encoder_name]
        rest = [item for item in methods if item.name != self.force_encoder_name]
        return preferred + rest

    def _profile_by_name(self, name: str) -> TurboProfile | None:
        needle = str(name or "").strip().lower()
        for profile in TURBO_PROFILES:
            if profile.name.lower() == needle:
                return profile
        return None

    def _base_profiles(self) -> list[TurboProfile]:
        forced = self._profile_by_name(self.force_profile_name)
        if forced:
            return [forced]
        return [
            self._profile_by_name("540p30"),
            self._profile_by_name("360p30"),
            self._profile_by_name("360p24"),
        ]

    def _upgrade_profiles(self, encoder_name: str) -> list[TurboProfile]:
        if self.force_profile_name:
            return []
        upgrades = [
            self._profile_by_name("540p60"),
            self._profile_by_name("720p30"),
        ]
        profiles = [item for item in upgrades if item is not None]
        if encoder_name == "libx264" and not self.allow_cpu_60:
            profiles = [item for item in profiles if item.name != "540p60"]
        return profiles

    def ensure_started(self, host_header: str | None = None) -> dict:
        with self._start_lock:
            self._reap_processes()
            capabilities = self.probe()
            if not capabilities.get("available"):
                return self.status(host_header=host_header)

            if not self._is_mediamtx_ready():
                self._start_mediamtx()
            if not self._is_mediamtx_ready():
                if not self._last_error:
                    self._last_error = "MediaMTX did not open its WebRTC port."
                return self.status(host_header=host_header)

            if self._selection is None or not self._is_ffmpeg_running():
                self._selection = self._select_stream_path()

            if self._selection and not self._is_ffmpeg_running():
                self._start_ffmpeg(self._selection)

            return self.status(host_header=host_header)

    def stop(self) -> None:
        with self._start_lock:
            self._terminate_process("_ffmpeg_proc")
            if self._mediamtx_owned:
                self._terminate_process("_mediamtx_proc")
            self._mediamtx_owned = False

    def status(self, host_header: str | None = None) -> dict:
        self._reap_processes()
        capabilities = self.probe()
        selected = self._selection.as_dict() if self._selection else None
        profile = selected.get("profile") if selected else None
        urls = self._urls_for_host(host_header)
        active = self._is_ffmpeg_running() and self._is_mediamtx_ready()
        reason = capabilities.get("reason") or self._last_error
        if capabilities.get("available") and not active and not reason:
            reason = "Turbo stream is ready but not started yet."
        return {
            "success": True,
            "enabled": self.enabled,
            "available": bool(capabilities.get("available")),
            "active": bool(active),
            "reason": reason,
            "last_error": self._last_error,
            "selected": selected,
            "profile": profile,
            "stream_name": self.stream_name,
            "playback_url": urls["playback_url"],
            "whep_url": urls["whep_url"],
            "whip_url": self._whip_url(),
            "jpeg_fallback": True,
            "auto_start": self.auto_start_enabled,
            "benchmark_enabled": self.benchmark_enabled,
            "benchmark_results": list(self._benchmark_results[-20:]),
            "capabilities": capabilities,
            "processes": {
                "ffmpeg": self._is_ffmpeg_running(),
                "mediamtx": self._is_mediamtx_ready(),
                "mediamtx_owned": self._mediamtx_owned,
            },
            "public_internet_note": (
                "For public internet WebRTC, expose MediaMTX WebRTC ports and configure STUN/TURN or "
                "webrtcAdditionalHosts. A normal HTTP-only tunnel can load the app but cannot reliably carry WebRTC media."
            ),
            "sources": {
                "ffmpeg_whip": "https://ffmpeg.org/ffmpeg-formats.html#whip",
                "mediamtx_webrtc": "https://mediamtx.org/docs/read/webrtc",
            },
        }

    def config_payload(self, host_header: str | None = None) -> dict:
        payload = self.status(host_header=host_header)
        payload["profiles"] = [asdict(item) for item in TURBO_PROFILES]
        payload["capture_priority"] = list(CAPTURE_PRIORITY)
        payload["encoder_priority"] = list(ENCODER_PRIORITY)
        payload["env"] = {
            "ZADOO_FFMPEG_PATH": "absolute path to ffmpeg.exe",
            "ZADOO_MEDIAMTX_PATH": "absolute path to mediamtx.exe",
            "ZADOO_TURBO_PROFILE": "optional fixed profile such as 540p30",
            "ZADOO_TURBO_CAPTURE": "optional fixed capture method",
            "ZADOO_TURBO_ENCODER": "optional fixed encoder",
            "ZADOO_TURBO_AUTO": "1 to auto-start, 0 to start on demand",
        }
        return payload

    def _select_stream_path(self) -> StreamSelection | None:
        self._benchmark_results = []
        captures = [item for item in self.capture_methods if item.available]
        encoders = [item for item in self.encoder_methods if item.available]
        base_profiles = [item for item in self._base_profiles() if item is not None]
        if not captures or not encoders or not base_profiles:
            self._last_error = "No capture, encoder, or profile candidates are available."
            return None

        selected: StreamSelection | None = None
        for capture in captures:
            for encoder in encoders:
                for profile in base_profiles:
                    selection = self._try_candidate(capture, encoder, profile)
                    if selection:
                        selected = selection
                        break
                if selected:
                    break
            if selected:
                break

        if selected is None:
            self._last_error = "Turbo benchmark failed for all capture and encoder candidates."
            return None

        for profile in self._upgrade_profiles(selected.encoder.name):
            upgraded = self._try_candidate(selected.capture, selected.encoder, profile)
            if upgraded:
                selected = upgraded

        self._last_error = ""
        return selected

    def _try_candidate(
        self,
        capture: CaptureMethod,
        encoder: EncoderMethod,
        profile: TurboProfile,
    ) -> StreamSelection | None:
        start = time.perf_counter()
        ok, output = self._benchmark_candidate(capture, encoder, profile)
        elapsed_ms = (time.perf_counter() - start) * 1000.0
        self._benchmark_results.append(
            {
                "capture": capture.name,
                "encoder": encoder.name,
                "profile": profile.name,
                "ok": bool(ok),
                "elapsed_ms": round(elapsed_ms, 1),
                "error": "" if ok else str(output or "")[:500],
            }
        )
        if ok:
            return StreamSelection(capture, encoder, profile, elapsed_ms)
        return None

    def _benchmark_candidate(
        self,
        capture: CaptureMethod,
        encoder: EncoderMethod,
        profile: TurboProfile,
    ) -> tuple[bool, str]:
        if not self.benchmark_enabled:
            return True, ""
        if not self.ffmpeg_path:
            return False, "ffmpeg.exe missing"
        cmd = self._build_ffmpeg_base_command(capture, profile)
        cmd.extend(["-t", f"{self.benchmark_seconds:.2f}", "-vf", self._video_filter(capture, profile), "-an"])
        cmd.extend(self._encoder_args(encoder, profile))
        cmd.extend(["-f", "null", os.devnull])
        code, output = _run_tool(cmd, timeout=self.benchmark_seconds + 8.0)
        return code == 0, output

    def _start_mediamtx(self) -> None:
        if not self.mediamtx_path:
            self._last_error = "mediamtx.exe missing"
            return
        if _tcp_open("127.0.0.1", self.webrtc_port):
            self._mediamtx_owned = False
            return
        try:
            self._mediamtx_proc = subprocess.Popen(
                [self.mediamtx_path],
                cwd=str(Path(self.mediamtx_path).resolve().parent),
                stdin=subprocess.DEVNULL,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=_hidden_creationflags(),
            )
            self._mediamtx_owned = True
            for _ in range(20):
                if self._is_mediamtx_ready():
                    return
                time.sleep(0.2)
            self._last_error = "MediaMTX started but its WebRTC port is not reachable."
        except Exception as exc:
            self._last_error = f"Failed to start MediaMTX: {exc}"

    def _start_ffmpeg(self, selection: StreamSelection) -> None:
        if not self.ffmpeg_path:
            self._last_error = "ffmpeg.exe missing"
            return
        cmd = self._build_publish_command(selection)
        try:
            self._ffmpeg_proc = subprocess.Popen(
                cmd,
                cwd=str(_repo_root()),
                stdin=subprocess.DEVNULL,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=_hidden_creationflags(),
            )
            self._last_started_at = time.time()
            time.sleep(0.5)
            if not self._is_ffmpeg_running():
                code = self._ffmpeg_proc.poll() if self._ffmpeg_proc else None
                self._last_error = f"FFmpeg exited during startup with code {code}."
                self._ffmpeg_proc = None
            else:
                self._last_error = ""
        except Exception as exc:
            self._last_error = f"Failed to start FFmpeg: {exc}"

    def _build_publish_command(self, selection: StreamSelection) -> list[str]:
        capture = selection.capture
        profile = selection.profile
        encoder = selection.encoder
        cmd = self._build_ffmpeg_base_command(capture, profile)
        cmd.extend(["-vf", self._video_filter(capture, profile)])
        if self.has_opus:
            cmd.extend(["-f", "lavfi", "-i", "anullsrc=channel_layout=stereo:sample_rate=48000"])
            cmd.extend(["-map", "0:v:0", "-map", "1:a:0"])
        else:
            cmd.extend(["-map", "0:v:0", "-an"])
        cmd.extend(self._encoder_args(encoder, profile))
        if self.has_opus:
            cmd.extend(["-c:a", "libopus", "-ar", "48000", "-ac", "2", "-b:a", "64k"])
        cmd.extend(["-strict", "experimental", "-f", "whip", self._whip_url()])
        return cmd

    def _build_ffmpeg_base_command(self, capture: CaptureMethod, profile: TurboProfile) -> list[str]:
        cmd = [self.ffmpeg_path or "ffmpeg", "-hide_banner", "-loglevel", "warning", "-nostdin", "-y"]
        if capture.name == "ddagrab":
            source = (
                f"ddagrab=output_idx={self.display_index}:framerate={profile.fps}:"
                "draw_mouse=1"
            )
            cmd.extend(["-f", "lavfi", "-i", source])
        elif capture.name == "gfxcapture":
            source = (
                f"gfxcapture=monitor_idx={self.display_index}:max_framerate={profile.fps}:"
                "capture_cursor=1"
            )
            cmd.extend(["-f", "lavfi", "-i", source])
        else:
            cmd.extend(["-f", "gdigrab", "-framerate", str(profile.fps), "-draw_mouse", "1", "-i", "desktop"])
        return cmd

    def _video_filter(self, capture: CaptureMethod, profile: TurboProfile) -> str:
        filters: list[str] = []
        if capture.name in {"ddagrab", "gfxcapture"}:
            filters.extend(["hwdownload", "format=bgra"])
        filters.extend(
            [
                f"fps={profile.fps}",
                f"scale=w={profile.width}:h={profile.height}:flags=fast_bilinear",
                "format=yuv420p",
            ]
        )
        return ",".join(filters)

    def _encoder_args(self, encoder: EncoderMethod, profile: TurboProfile) -> list[str]:
        fps = max(1, int(profile.fps))
        base = [
            "-c:v",
            encoder.name,
            "-b:v",
            profile.bitrate,
            "-maxrate",
            profile.maxrate,
            "-bufsize",
            profile.bufsize,
            "-g",
            str(fps),
            "-bf",
            "0",
            "-pix_fmt",
            "yuv420p",
        ]
        if encoder.name == "libx264":
            return [
                "-c:v",
                "libx264",
                "-preset",
                "ultrafast",
                "-tune",
                "zerolatency",
                "-profile:v",
                "baseline",
                "-x264-params",
                f"keyint={fps}:min-keyint={fps}:scenecut=0",
                "-b:v",
                profile.bitrate,
                "-maxrate",
                profile.maxrate,
                "-bufsize",
                profile.bufsize,
                "-g",
                str(fps),
                "-bf",
                "0",
                "-pix_fmt",
                "yuv420p",
            ]
        if encoder.name == "h264_nvenc":
            return base + ["-preset", "p1", "-tune", "ull", "-rc", "cbr"]
        if encoder.name == "h264_amf":
            return base + ["-usage", "lowlatency", "-rc", "cbr"]
        if encoder.name == "h264_qsv":
            return base + ["-preset", "veryfast"]
        if encoder.name == "h264_mf":
            return base + ["-profile:v", "baseline"]
        return base

    def _whip_url(self) -> str:
        quoted_name = urllib.parse.quote(self.stream_name.strip("/"))
        return f"http://127.0.0.1:{self.webrtc_port}/{quoted_name}/whip"

    def _urls_for_host(self, host_header: str | None) -> dict:
        if self.public_base_url:
            base = self.public_base_url.rstrip("/")
        else:
            host = str(host_header or "").split(",", 1)[0].strip()
            if not host:
                host = f"localhost:{self.webrtc_port}"
            hostname = host.rsplit("@", 1)[-1].split(":", 1)[0].strip("[]") or "localhost"
            if hostname in {"0.0.0.0", "::"}:
                hostname = "localhost"
            base = f"http://{hostname}:{self.webrtc_port}"
        quoted_name = urllib.parse.quote(self.stream_name.strip("/"))
        playback_url = f"{base}/{quoted_name}"
        return {
            "playback_url": playback_url,
            "whep_url": f"{playback_url}/whep",
        }

    def _is_mediamtx_ready(self) -> bool:
        proc_ready = self._mediamtx_proc is not None and self._mediamtx_proc.poll() is None
        return proc_ready or _tcp_open("127.0.0.1", self.webrtc_port)

    def _is_ffmpeg_running(self) -> bool:
        return self._ffmpeg_proc is not None and self._ffmpeg_proc.poll() is None

    def _reap_processes(self) -> None:
        if self._ffmpeg_proc is not None and self._ffmpeg_proc.poll() is not None:
            code = self._ffmpeg_proc.returncode
            if code not in (0, None):
                self._last_error = f"FFmpeg exited with code {code}."
            self._ffmpeg_proc = None
        if self._mediamtx_proc is not None and self._mediamtx_proc.poll() is not None:
            code = self._mediamtx_proc.returncode
            if code not in (0, None):
                self._last_error = f"MediaMTX exited with code {code}."
            self._mediamtx_proc = None
            self._mediamtx_owned = False

    def _terminate_process(self, attr_name: str) -> None:
        proc = getattr(self, attr_name, None)
        if proc is None:
            return
        try:
            if proc.poll() is None:
                proc.terminate()
                try:
                    proc.wait(timeout=3.0)
                except subprocess.TimeoutExpired:
                    proc.kill()
        except Exception:
            pass
        setattr(self, attr_name, None)
