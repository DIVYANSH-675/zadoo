"""Generic Windows low-latency WebRTC streaming engine.

The fast path is FFmpeg desktop capture -> H.264 -> MediaMTX -> browser
WebRTC. JPEG WebSocket stays alive as the final compatibility fallback.
Selection is capability and benchmark based so ordinary Windows laptops,
integrated GPUs, dedicated GPUs, and CPU-only systems can all choose a safe
stream path.
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
import json
import logging
from logging.handlers import RotatingFileHandler
from collections import deque
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Deque


FALSE_VALUES = {"0", "false", "no", "off", "disabled"}
TRUE_VALUES = {"1", "true", "yes", "on", "enabled"}
HW_ENCODERS = {"h264_mf", "h264_qsv", "h264_amf", "h264_nvenc"}


@dataclass(frozen=True)
class TurboProfile:
    name: str
    width: int
    height: int
    fps: int
    bitrate_kbps: int
    maxrate_kbps: int
    status: str
    native_resolution: bool = False

    @property
    def bitrate(self) -> str:
        return f"{self.bitrate_kbps}k"

    @property
    def maxrate(self) -> str:
        return f"{self.maxrate_kbps}k"

    @property
    def bufsize(self) -> str:
        return f"{max(self.maxrate_kbps * 2, self.bitrate_kbps * 2)}k"

    @property
    def uses_native_resolution(self) -> bool:
        return bool(self.native_resolution or self.width <= 0 or self.height <= 0)


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
    publish_transport: str
    zero_copy: bool
    benchmark_ms: float

    def as_dict(self) -> dict:
        profile = asdict(self.profile)
        profile["native_resolution"] = bool(self.profile.uses_native_resolution)
        return {
            "capture": asdict(self.capture),
            "encoder": asdict(self.encoder),
            "profile": profile,
            "publish_transport": self.publish_transport,
            "zero_copy": bool(self.zero_copy),
            "benchmark_ms": round(float(self.benchmark_ms), 1),
        }


TURBO_PROFILES = (
    TurboProfile("540p30", 960, 540, 30, 3500, 6000, "Turbo 540p30"),
    TurboProfile("540p45", 960, 540, 45, 5000, 7000, "Turbo 540p45"),
    TurboProfile("540p60", 960, 540, 60, 6000, 8000, "Turbo 540p60"),
    TurboProfile("540p100", 960, 540, 100, 6500, 8000, "Turbo 540p100"),
    TurboProfile("540p120", 960, 540, 120, 8000, 10000, "Turbo 540p120"),
    TurboProfile("720p30", 1280, 720, 30, 5000, 7000, "Turbo 720p30"),
    TurboProfile("720p60", 1280, 720, 60, 7000, 9000, "Turbo 720p60"),
    TurboProfile("360p30", 640, 360, 30, 1800, 2500, "Turbo 360p30"),
    TurboProfile("360p24", 640, 360, 24, 1200, 2000, "Turbo 360p24"),
    TurboProfile("native30", 0, 0, 30, 12000, 18000, "Turbo Native 30 (100% quality)", True),
    TurboProfile("native60", 0, 0, 60, 20000, 28000, "Turbo Native 60 (100% quality)", True),
    TurboProfile("native100", 0, 0, 100, 28000, 36000, "Turbo Native 100 (100% quality)", True),
)

CAPTURE_PRIORITY = ("ddagrab", "gfxcapture", "gdigrab")
ENCODER_PRIORITY = ("h264_mf", "h264_qsv", "h264_amf", "h264_nvenc", "libx264")
TRANSPORT_PRIORITY = ("rtsp", "whip")


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


def _turbo_logger() -> logging.Logger:
    logger = logging.getLogger("zadoo.turbo")
    logger.setLevel(logging.DEBUG)
    try:
        log_dir = _repo_root() / "logs"
        log_dir.mkdir(parents=True, exist_ok=True)
        log_path = log_dir / "turbo_stream.log"
        has_handler = any(
            isinstance(handler, RotatingFileHandler)
            and str(getattr(handler, "baseFilename", "")).lower() == str(log_path).lower()
            for handler in logger.handlers
        )
        if not has_handler:
            handler = RotatingFileHandler(
                log_path,
                maxBytes=5_000_000,
                backupCount=3,
                encoding="utf-8",
            )
            handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
            logger.addHandler(handler)
        logger.propagate = True
    except Exception:
        pass
    return logger


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
            root / "tools" / exe_name.replace(".exe", "") / exe_name,
            root / "tools" / "ffmpeg" / "bin" / exe_name,
            root / "tools" / "mediamtx" / exe_name,
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
        self.allow_cpu_720 = _env_flag("ZADOO_TURBO_CPU_720", False)
        self.stream_name = os.getenv("ZADOO_TURBO_STREAM_NAME", "zadoo").strip() or "zadoo"
        self.display_index = self._int_env("ZADOO_TURBO_DISPLAY_INDEX", 0, 0, 16)
        self.rtsp_port = self._int_env("ZADOO_MEDIAMTX_RTSP_PORT", 8554, 1, 65535)
        self.webrtc_port = self._int_env("ZADOO_MEDIAMTX_WEBRTC_PORT", 8889, 1, 65535)
        self.max_fps = self._int_env("ZADOO_TURBO_MAX_FPS", 100, 24, 240)
        self.quality_percent = self._int_env("ZADOO_TURBO_QUALITY", 65, 10, 100)
        self.full_quality_at = self._int_env("ZADOO_TURBO_FULL_QUALITY_AT", 100, 85, 100)
        self.benchmark_seconds = self._float_env("ZADOO_TURBO_BENCH_SECONDS", 1.8, 0.5, 8.0)
        self.force_profile_name = os.getenv("ZADOO_TURBO_PROFILE", "").strip().lower()
        self.force_capture_name = os.getenv("ZADOO_TURBO_CAPTURE", "").strip().lower()
        self.force_encoder_name = os.getenv("ZADOO_TURBO_ENCODER", "").strip().lower()
        self.transport_mode = os.getenv("ZADOO_TURBO_TRANSPORT", "rtsp").strip().lower() or "rtsp"
        self.public_base_url = os.getenv("ZADOO_TURBO_PUBLIC_URL", "").strip()

        self.ffmpeg_path: str | None = None
        self.mediamtx_path: str | None = None
        self.capabilities: dict = {}
        self.capture_methods: list[CaptureMethod] = []
        self.encoder_methods: list[EncoderMethod] = []
        self.has_whip = False
        self.has_rtsp = False
        self.has_opus = False
        self.has_scale_d3d11 = False
        self.reason = ""

        self._probe_lock = threading.RLock()
        self._start_lock = threading.RLock()
        self._probe_done = False
        self._last_probe_at = 0.0
        self._selection: StreamSelection | None = None
        self._benchmark_results: list[dict] = []
        self._last_error = ""
        self._last_limit_reason = ""
        self._ffmpeg_proc: subprocess.Popen | None = None
        self._mediamtx_proc: subprocess.Popen | None = None
        self._mediamtx_owned = False
        self._last_started_at = 0.0
        self._last_ffmpeg_command: list[str] = []
        self._ffmpeg_log_tail: Deque[str] = deque(maxlen=80)
        self._mediamtx_log_tail: Deque[str] = deque(maxlen=80)
        self._logger = _turbo_logger()
        self._log_event(
            "initialized",
            enabled=self.enabled,
            auto_start=self.auto_start_enabled,
            max_fps=self.max_fps,
            quality=self.quality_percent,
            full_quality_at=self.full_quality_at,
            transport=self.transport_mode,
        )

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

    def _log_event(self, message: str, **details) -> None:
        try:
            suffix = ""
            if details:
                suffix = " " + json.dumps(details, sort_keys=True, default=str)
            self._logger.info("%s%s", message, suffix)
        except Exception:
            pass

    def _full_quality_requested(self, quality: int | None = None) -> bool:
        value = self.quality_percent if quality is None else quality
        return int(value) >= int(self.full_quality_at)

    def set_quality(self, raw_value, restart_active: bool = True) -> dict:
        try:
            value = max(10, min(100, int(raw_value)))
        except Exception:
            value = self.quality_percent
        with self._start_lock:
            previous_quality = self.quality_percent
            previous_full_quality = self._full_quality_requested(previous_quality)
            next_full_quality = self._full_quality_requested(value)
            was_active = self._is_ffmpeg_running()
            mode_changed = previous_full_quality != next_full_quality
            self.quality_percent = value
            restart_required = bool(restart_active and mode_changed and (was_active or self._selection))
            if mode_changed:
                self._selection = None
                self._benchmark_results = []
                self._last_limit_reason = "quality_changed"
                if restart_required:
                    self._terminate_process("_ffmpeg_proc")
            if previous_quality != value or mode_changed:
                self._log_event(
                    "quality_changed",
                    previous_quality=previous_quality,
                    quality=value,
                    previous_full_quality=previous_full_quality,
                    full_quality=next_full_quality,
                    restart_required=restart_required,
                    was_active=was_active,
                )
            return {
                "quality": value,
                "full_quality": next_full_quality,
                "mode_changed": mode_changed,
                "restart_required": restart_required,
                "was_active": was_active,
            }

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
            self.has_rtsp = False
            self.has_opus = False
            self.has_scale_d3d11 = False
            filters_text = ""
            devices_text = ""
            encoders_text = ""
            muxers_text = ""
            protocols_text = ""

            if self.ffmpeg_path:
                _, filters_text = _run_tool([self.ffmpeg_path, "-hide_banner", "-filters"], timeout=5.0)
                _, devices_text = _run_tool([self.ffmpeg_path, "-hide_banner", "-devices"], timeout=5.0)
                _, encoders_text = _run_tool([self.ffmpeg_path, "-hide_banner", "-encoders"], timeout=5.0)
                _, muxers_text = _run_tool([self.ffmpeg_path, "-hide_banner", "-muxers"], timeout=5.0)
                _, protocols_text = _run_tool([self.ffmpeg_path, "-hide_banner", "-protocols"], timeout=5.0)

            has_ddagrab = _has_ffmpeg_name(filters_text, "ddagrab")
            has_gfxcapture = _has_ffmpeg_name(filters_text, "gfxcapture")
            has_gdigrab = _has_ffmpeg_name(devices_text, "gdigrab")
            self.has_scale_d3d11 = _has_ffmpeg_name(filters_text, "scale_d3d11")
            self.has_whip = _has_ffmpeg_name(muxers_text, "whip")
            self.has_rtsp = _has_ffmpeg_name(muxers_text, "rtsp") or _has_ffmpeg_name(protocols_text, "rtsp")
            self.has_opus = _has_ffmpeg_name(encoders_text, "libopus") or _has_ffmpeg_name(encoders_text, "opus")

            self.capture_methods = self._ordered_capture_methods(
                [
                    CaptureMethod("ddagrab", "Desktop Duplication (D3D11 capture)", "lavfi", has_ddagrab,
                                  "" if has_ddagrab else "FFmpeg ddagrab filter is not available."),
                    CaptureMethod("gfxcapture", "Windows Graphics Capture", "lavfi", has_gfxcapture,
                                  "" if has_gfxcapture else "FFmpeg gfxcapture filter is not available."),
                    CaptureMethod("gdigrab", "GDI desktop capture", "device", has_gdigrab,
                                  "" if has_gdigrab else "FFmpeg gdigrab input device is not available."),
                ]
            )

            encoder_labels = {
                "h264_mf": "Media Foundation H.264",
                "h264_qsv": "Intel Quick Sync H.264",
                "h264_amf": "AMD AMF H.264",
                "h264_nvenc": "NVIDIA NVENC H.264",
                "libx264": "CPU x264 H.264",
            }
            self.encoder_methods = self._ordered_encoder_methods(
                [
                    EncoderMethod(name, encoder_labels[name], _has_ffmpeg_name(encoders_text, name),
                                  "" if _has_ffmpeg_name(encoders_text, name) else f"FFmpeg encoder {name} is not available.")
                    for name in ENCODER_PRIORITY
                ]
            )

            available_captures = [item.name for item in self.capture_methods if item.available]
            available_encoders = [item.name for item in self.encoder_methods if item.available]
            transports = self._candidate_transports()
            reasons: list[str] = []
            if not self.enabled:
                reasons.append("Turbo WebRTC is disabled by ZADOO_TURBO_STREAM=0.")
            if platform_name.lower() != "windows":
                reasons.append("Turbo capture is Windows-only; JPEG fallback remains active.")
            if not self.ffmpeg_path:
                reasons.append("ffmpeg.exe was not found. Run scripts\\setup_turbo_stream.ps1 or set ZADOO_FFMPEG_PATH.")
            if not self.mediamtx_path:
                reasons.append("mediamtx.exe was not found. Run scripts\\setup_turbo_stream.ps1 or set ZADOO_MEDIAMTX_PATH.")
            if self.ffmpeg_path and not transports:
                reasons.append("No FFmpeg publish transport is available. RTSP is preferred; WHIP is optional.")
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
                and bool(transports)
                and bool(available_captures)
                and bool(available_encoders)
            )
            self.capabilities = {
                "enabled": self.enabled,
                "available": available,
                "platform": platform_name,
                "ffmpeg_path": self.ffmpeg_path,
                "mediamtx_path": self.mediamtx_path,
                "has_rtsp": self.has_rtsp,
                "has_whip": self.has_whip,
                "has_opus": self.has_opus,
                "has_scale_d3d11": self.has_scale_d3d11,
                "capture_priority": list(CAPTURE_PRIORITY),
                "encoder_priority": list(ENCODER_PRIORITY),
                "transport_priority": list(TRANSPORT_PRIORITY),
                "capture_methods": [asdict(item) for item in self.capture_methods],
                "encoders": [asdict(item) for item in self.encoder_methods],
                "available_captures": available_captures,
                "available_encoders": available_encoders,
                "available_transports": transports,
                "reason": self.reason,
                "last_probe_at": self._last_probe_at,
            }
            self._probe_done = True
            self._log_event(
                "probe_complete",
                available=available,
                captures=available_captures,
                encoders=available_encoders,
                transports=transports,
                reason=self.reason,
            )
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

    def _candidate_transports(self) -> list[str]:
        requested = self.transport_mode
        if requested == "auto":
            order = list(TRANSPORT_PRIORITY)
        elif requested in TRANSPORT_PRIORITY:
            order = [requested]
        else:
            order = ["rtsp"]
        transports: list[str] = []
        for transport in order:
            if transport == "rtsp" and self.has_rtsp:
                transports.append(transport)
            elif transport == "whip" and self.has_whip:
                transports.append(transport)
        return transports

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
        if self._full_quality_requested():
            return [
                self._profile_by_name("native100"),
                self._profile_by_name("native60"),
                self._profile_by_name("native30"),
            ]
        return [
            self._profile_by_name("540p30"),
            self._profile_by_name("360p30"),
            self._profile_by_name("360p24"),
        ]

    def _upgrade_profiles(self, encoder_name: str) -> list[TurboProfile]:
        if self.force_profile_name or self._full_quality_requested():
            return []
        upgrades = [
            self._profile_by_name("540p120"),
            self._profile_by_name("540p100"),
            self._profile_by_name("540p60"),
            self._profile_by_name("720p60"),
            self._profile_by_name("540p45"),
            self._profile_by_name("720p30"),
        ]
        profiles = [item for item in upgrades if item is not None]
        profiles = [item for item in profiles if int(item.fps) <= int(self.max_fps)]
        if encoder_name == "libx264" and not self.allow_cpu_60:
            profiles = [item for item in profiles if int(item.fps) <= 30]
        if encoder_name == "libx264" and not self.allow_cpu_720:
            profiles = [item for item in profiles if int(item.width or 0) <= 960]
        return profiles

    def ensure_started(self, host_header: str | None = None) -> dict:
        with self._start_lock:
            self._reap_processes()
            capabilities = self.probe()
            if not capabilities.get("available"):
                return self.status(host_header=host_header)

            if not self._selection:
                self._selection = self._select_stream_path()
            if not self._selection:
                return self.status(host_header=host_header)

            if not self._is_mediamtx_ready_for(self._selection.publish_transport):
                self._start_mediamtx(self._selection.publish_transport)
            if not self._is_mediamtx_ready_for(self._selection.publish_transport):
                if not self._last_error:
                    self._last_error = "MediaMTX did not open the required RTSP/WebRTC ports."
                return self.status(host_header=host_header)

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
        active = self._is_ffmpeg_running() and self._is_mediamtx_ready_for(
            (selected or {}).get("publish_transport", "rtsp")
        )
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
            "limit_reason": self._last_limit_reason,
            "selected": selected,
            "profile": profile,
            "stream_name": self.stream_name,
            "max_fps": self.max_fps,
            "quality": self.quality_percent,
            "full_quality": self._full_quality_requested(),
            "full_quality_at": self.full_quality_at,
            "playback_url": urls["playback_url"],
            "whep_url": urls["whep_url"],
            "publish_url": self._publish_url((selected or {}).get("publish_transport", "rtsp")),
            "jpeg_fallback": True,
            "auto_start": self.auto_start_enabled,
            "benchmark_enabled": self.benchmark_enabled,
            "benchmark_results": list(self._benchmark_results[-30:]),
            "capabilities": capabilities,
            "processes": {
                "ffmpeg": self._is_ffmpeg_running(),
                "mediamtx": self._is_mediamtx_ready_for((selected or {}).get("publish_transport", "rtsp")),
                "mediamtx_owned": self._mediamtx_owned,
                "rtsp_port_open": _tcp_open("127.0.0.1", self.rtsp_port),
                "webrtc_port_open": _tcp_open("127.0.0.1", self.webrtc_port),
            },
            "public_internet_note": (
                "For public internet WebRTC, expose MediaMTX WebRTC ports and configure STUN/TURN or "
                "webrtcAdditionalHosts. A normal HTTP-only tunnel can load the app but cannot reliably carry WebRTC media."
            ),
            "sources": {
                "ffmpeg_ddagrab": "https://ffmpeg.org/ffmpeg-all.html#ddagrab",
                "ffmpeg_gfxcapture": "https://ffmpeg.org/ffmpeg-all.html#gfxcapture",
                "mediamtx_webrtc": "https://mediamtx.org/docs/read/webrtc",
            },
        }

    def diagnostics(self, host_header: str | None = None) -> dict:
        payload = self.status(host_header=host_header)
        payload["diagnostics"] = {
            "ffmpeg_log_tail": list(self._ffmpeg_log_tail),
            "mediamtx_log_tail": list(self._mediamtx_log_tail),
            "last_ffmpeg_command": self._redact_command(self._last_ffmpeg_command),
            "turbo_log_path": str((_repo_root() / "logs" / "turbo_stream.log").resolve()),
            "setup_hint": r"Run scripts\setup_turbo_stream.ps1, then restart the app.",
            "fps_floor_explanation": self._fps_floor_explanation(payload),
        }
        return payload

    def config_payload(self, host_header: str | None = None) -> dict:
        payload = self.status(host_header=host_header)
        payload["profiles"] = [asdict(item) for item in TURBO_PROFILES]
        payload["capture_priority"] = list(CAPTURE_PRIORITY)
        payload["encoder_priority"] = list(ENCODER_PRIORITY)
        payload["transport_priority"] = list(TRANSPORT_PRIORITY)
        payload["env"] = {
            "ZADOO_FFMPEG_PATH": "absolute path to ffmpeg.exe",
            "ZADOO_MEDIAMTX_PATH": "absolute path to mediamtx.exe",
            "ZADOO_TURBO_PROFILE": "optional fixed profile such as 540p30",
            "ZADOO_TURBO_CAPTURE": "optional fixed capture method",
            "ZADOO_TURBO_ENCODER": "optional fixed encoder",
            "ZADOO_TURBO_TRANSPORT": "rtsp, whip, or auto; rtsp is preferred",
            "ZADOO_TURBO_AUTO": "1 to auto-start, 0 to start on demand",
            "ZADOO_TURBO_CPU_60": "1 to allow CPU x264 profiles above 30 FPS after benchmark success",
            "ZADOO_TURBO_CPU_720": "1 to allow CPU x264 720p profiles after benchmark success",
            "ZADOO_TURBO_QUALITY": "initial UI quality value from 10 to 100",
            "ZADOO_TURBO_FULL_QUALITY_AT": "quality threshold that switches Turbo to native-resolution profiles; default 100",
        }
        return payload

    def _select_stream_path(self) -> StreamSelection | None:
        self._benchmark_results = []
        captures = [item for item in self.capture_methods if item.available]
        encoders = [item for item in self.encoder_methods if item.available]
        transports = self._candidate_transports()
        base_profiles = [item for item in self._base_profiles() if item is not None]
        if not captures or not encoders or not transports or not base_profiles:
            self._last_error = "No capture, encoder, transport, or profile candidates are available."
            return None

        selected: StreamSelection | None = None
        disabled_zero_copy_captures: set[str] = set()
        for transport in transports:
            for capture in captures:
                for encoder in encoders:
                    for zero_copy in self._zero_copy_modes(capture, encoder):
                        if zero_copy and capture.name in disabled_zero_copy_captures:
                            self._log_event(
                                "benchmark_skip",
                                capture=capture.name,
                                encoder=encoder.name,
                                zero_copy=True,
                                reason="zero_copy_previously_failed_for_capture",
                            )
                            continue
                        for profile in base_profiles:
                            if zero_copy and profile.uses_native_resolution:
                                continue
                            selection = self._try_candidate(capture, encoder, profile, transport, zero_copy)
                            if selection:
                                selected = selection
                                break
                            last_error = ""
                            try:
                                last_error = str((self._benchmark_results[-1] or {}).get("error") or "")
                            except Exception:
                                last_error = ""
                            if zero_copy and self._is_zero_copy_path_failure(last_error):
                                disabled_zero_copy_captures.add(capture.name)
                                self._log_event(
                                    "benchmark_skip_remaining_zero_copy",
                                    capture=capture.name,
                                    encoder=encoder.name,
                                    profile=profile.name,
                                    reason="zero_copy_filter_failed",
                                )
                                break
                            if self._is_encoder_path_failure(last_error):
                                self._log_event(
                                    "benchmark_skip_remaining_profiles",
                                    capture=capture.name,
                                    encoder=encoder.name,
                                    profile=profile.name,
                                    zero_copy=zero_copy,
                                    reason="encoder_path_failed",
                                )
                                break
                        if selected:
                            break
                    if selected:
                        break
                if selected:
                    break
            if selected:
                break

        if selected is None:
            self._last_error = "Turbo benchmark failed for all capture and encoder candidates."
            self._last_limit_reason = "benchmark_failed"
            self._log_event("selection_failed", quality=self.quality_percent, full_quality=self._full_quality_requested())
            return None

        for profile in self._upgrade_profiles(selected.encoder.name):
            upgraded = self._try_candidate(
                selected.capture,
                selected.encoder,
                profile,
                selected.publish_transport,
                selected.zero_copy,
            )
            if upgraded:
                selected = upgraded
                break
            else:
                self._last_limit_reason = f"{profile.name}_benchmark_failed"

        self._last_error = ""
        self._log_event(
            "selection_ready",
            capture=selected.capture.name,
            encoder=selected.encoder.name,
            profile=selected.profile.name,
            transport=selected.publish_transport,
            zero_copy=selected.zero_copy,
            quality=self.quality_percent,
            full_quality=self._full_quality_requested(),
        )
        return selected

    def _zero_copy_modes(self, capture: CaptureMethod, encoder: EncoderMethod) -> list[bool]:
        if encoder.name not in HW_ENCODERS or capture.name not in {"ddagrab", "gfxcapture"}:
            return [False]
        if capture.name == "ddagrab" and not self.has_scale_d3d11:
            return [False]
        return [True, False]

    def _try_candidate(
        self,
        capture: CaptureMethod,
        encoder: EncoderMethod,
        profile: TurboProfile,
        transport: str,
        zero_copy: bool,
    ) -> StreamSelection | None:
        start = time.perf_counter()
        ok, output = self._benchmark_candidate(capture, encoder, profile, zero_copy)
        elapsed_ms = (time.perf_counter() - start) * 1000.0
        frame_budget_ms = 1000.0 / max(1, int(profile.fps))
        bench = {
            "capture": capture.name,
            "encoder": encoder.name,
            "profile": profile.name,
            "transport": transport,
            "zero_copy": bool(zero_copy),
            "ok": bool(ok),
            "elapsed_ms": round(elapsed_ms, 1),
            "frame_budget_ms": round(frame_budget_ms, 1),
            "error": "" if ok else str(output or "")[:500],
        }
        self._benchmark_results.append(bench)
        self._log_event("benchmark_candidate", **bench)
        if ok:
            return StreamSelection(capture, encoder, profile, transport, zero_copy, elapsed_ms)
        return None

    @staticmethod
    def _is_zero_copy_path_failure(error_text: str) -> bool:
        text = str(error_text or "").lower()
        zero_copy_markers = (
            "parsed_scale_d3d11",
            "failed to configure output pad",
            "could not create the texture",
            "impossible to convert between the formats supported by the filter",
            "link 'parsed_scale_d3d11",
        )
        return any(marker in text for marker in zero_copy_markers)

    @staticmethod
    def _is_encoder_path_failure(error_text: str) -> bool:
        text = str(error_text or "").lower()
        encoder_markers = (
            "dll amfrt64.dll failed to open",
            "error creating a mfx session",
            "the current mfx implementation is not supported",
            "failed to create  hardware device context",
            "format negotiation failed",
            "error while opening encoder",
            "no capable devices found",
            "cannot load nvcuda.dll",
            "encoder not found",
        )
        return any(marker in text for marker in encoder_markers)

    def _benchmark_candidate(
        self,
        capture: CaptureMethod,
        encoder: EncoderMethod,
        profile: TurboProfile,
        zero_copy: bool,
    ) -> tuple[bool, str]:
        if not self.benchmark_enabled:
            return True, ""
        if not self.ffmpeg_path:
            return False, "ffmpeg.exe missing"
        cmd = self._build_ffmpeg_base_command(capture, profile, zero_copy)
        filter_chain = self._video_filter(capture, encoder, profile, zero_copy)
        if filter_chain:
            cmd.extend(["-vf", filter_chain])
        cmd.extend(["-t", f"{self.benchmark_seconds:.2f}", "-an"])
        cmd.extend(self._encoder_args(encoder, profile, zero_copy))
        cmd.extend(["-f", "null", os.devnull])
        code, output = _run_tool(cmd, timeout=self.benchmark_seconds + 8.0)
        return code == 0, output

    def _start_mediamtx(self, transport: str) -> None:
        if not self.mediamtx_path:
            self._last_error = "mediamtx.exe missing"
            return
        if self._is_mediamtx_ready_for(transport):
            self._mediamtx_owned = False
            self._log_event("mediamtx_reused", transport=transport)
            return
        try:
            self._mediamtx_log_tail.clear()
            self._log_event("mediamtx_starting", transport=transport, path=self.mediamtx_path)
            self._mediamtx_proc = subprocess.Popen(
                [self.mediamtx_path],
                cwd=str(Path(self.mediamtx_path).resolve().parent),
                stdin=subprocess.DEVNULL,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                creationflags=_hidden_creationflags(),
            )
            self._spawn_log_reader(self._mediamtx_proc.stdout, self._mediamtx_log_tail)
            self._mediamtx_owned = True
            for _ in range(30):
                if self._is_mediamtx_ready_for(transport):
                    self._log_event("mediamtx_ready", transport=transport)
                    return
                time.sleep(0.2)
            self._last_error = "MediaMTX started but the required RTSP/WebRTC port is not reachable."
            self._log_event("mediamtx_not_ready", transport=transport, error=self._last_error)
        except Exception as exc:
            self._last_error = f"Failed to start MediaMTX: {exc}"
            self._log_event("mediamtx_start_failed", transport=transport, error=str(exc))

    def _start_ffmpeg(self, selection: StreamSelection) -> None:
        if not self.ffmpeg_path:
            self._last_error = "ffmpeg.exe missing"
            return
        cmd = self._build_publish_command(selection)
        self._last_ffmpeg_command = list(cmd)
        try:
            self._ffmpeg_log_tail.clear()
            self._log_event(
                "ffmpeg_starting",
                capture=selection.capture.name,
                encoder=selection.encoder.name,
                profile=selection.profile.name,
                transport=selection.publish_transport,
                zero_copy=selection.zero_copy,
                command=self._redact_command(cmd),
            )
            self._ffmpeg_proc = subprocess.Popen(
                cmd,
                cwd=str(_repo_root()),
                stdin=subprocess.DEVNULL,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                creationflags=_hidden_creationflags(),
            )
            self._spawn_log_reader(self._ffmpeg_proc.stdout, self._ffmpeg_log_tail)
            self._last_started_at = time.time()
            time.sleep(0.8)
            if not self._is_ffmpeg_running():
                code = self._ffmpeg_proc.poll() if self._ffmpeg_proc else None
                tail = " ".join(list(self._ffmpeg_log_tail)[-5:])
                self._last_error = f"FFmpeg exited during startup with code {code}. {tail}".strip()
                self._log_event("ffmpeg_start_failed", code=code, tail=tail)
                self._ffmpeg_proc = None
            else:
                self._last_error = ""
                self._log_event("ffmpeg_ready", profile=selection.profile.name)
        except Exception as exc:
            self._last_error = f"Failed to start FFmpeg: {exc}"
            self._log_event("ffmpeg_start_exception", error=str(exc))

    def _build_publish_command(self, selection: StreamSelection) -> list[str]:
        capture = selection.capture
        profile = selection.profile
        encoder = selection.encoder
        cmd = self._build_ffmpeg_base_command(capture, profile, selection.zero_copy)
        if self.has_opus:
            cmd.extend(["-f", "lavfi", "-i", "anullsrc=channel_layout=stereo:sample_rate=48000"])
            cmd.extend(["-map", "0:v:0", "-map", "1:a:0"])
        else:
            cmd.extend(["-map", "0:v:0", "-an"])
        filter_chain = self._video_filter(capture, encoder, profile, selection.zero_copy)
        if filter_chain:
            cmd.extend(["-vf", filter_chain])
        cmd.extend(self._encoder_args(encoder, profile, selection.zero_copy))
        if self.has_opus:
            cmd.extend(["-c:a", "libopus", "-ar", "48000", "-ac", "2", "-b:a", "64k"])
        if selection.publish_transport == "whip":
            cmd.extend(["-strict", "experimental", "-f", "whip", self._publish_url("whip")])
        else:
            cmd.extend(["-f", "rtsp", "-rtsp_transport", "tcp", self._publish_url("rtsp")])
        return cmd

    def _build_ffmpeg_base_command(self, capture: CaptureMethod, profile: TurboProfile, zero_copy: bool) -> list[str]:
        cmd = [self.ffmpeg_path or "ffmpeg", "-hide_banner", "-loglevel", "warning", "-nostdin", "-y"]
        if capture.name == "ddagrab":
            source = (
                f"ddagrab=output_idx={self.display_index}:framerate={profile.fps}:"
                "draw_mouse=1:dup_frames=1:output_fmt=8bit"
            )
            cmd.extend(["-f", "lavfi", "-i", source])
        elif capture.name == "gfxcapture":
            if profile.uses_native_resolution:
                source = (
                    f"gfxcapture=monitor_idx={self.display_index}:max_framerate={profile.fps}:"
                    "capture_cursor=1:output_fmt=8bit"
                )
            else:
                source = (
                    f"gfxcapture=monitor_idx={self.display_index}:max_framerate={profile.fps}:"
                    f"width={profile.width}:height={profile.height}:resize_mode=scale_aspect:"
                    "capture_cursor=1:output_fmt=8bit"
                )
            cmd.extend(["-f", "lavfi", "-i", source])
        else:
            cmd.extend(["-f", "gdigrab", "-framerate", str(profile.fps), "-draw_mouse", "1", "-i", "desktop"])
        return cmd

    def _video_filter(
        self,
        capture: CaptureMethod,
        encoder: EncoderMethod,
        profile: TurboProfile,
        zero_copy: bool,
    ) -> str:
        if zero_copy:
            filters: list[str] = []
            if capture.name == "ddagrab":
                if profile.uses_native_resolution:
                    filters.append("format=nv12")
                else:
                    filters.append(f"scale_d3d11=width={profile.width}:height={profile.height}:format=nv12")
            if capture.name == "gfxcapture":
                filters.append(f"fps={profile.fps}")
            return ",".join(filters)

        filters = []
        if capture.name in {"ddagrab", "gfxcapture"}:
            filters.extend(["hwdownload", "format=bgra"])
        filters.append(f"fps={profile.fps}")
        if not profile.uses_native_resolution:
            filters.append(f"scale=w={profile.width}:h={profile.height}:flags=fast_bilinear")
        filters.append("format=yuv420p" if encoder.name == "libx264" else "format=nv12")
        return ",".join(filters)

    def _encoder_args(self, encoder: EncoderMethod, profile: TurboProfile, zero_copy: bool) -> list[str]:
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
        ]
        if not zero_copy:
            base.extend(["-pix_fmt", "yuv420p"])
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
        if encoder.name == "h264_mf":
            return base + ["-hw_encoding", "1", "-rate_control", "cbr", "-scenario", "display_remoting"]
        if encoder.name == "h264_nvenc":
            return base + ["-preset", "p1", "-tune", "ull", "-rc", "cbr", "-zerolatency", "1", "-delay", "0"]
        if encoder.name == "h264_amf":
            return base + ["-usage", "lowlatency", "-quality", "speed", "-rc", "cbr"]
        if encoder.name == "h264_qsv":
            return base + ["-preset", "veryfast", "-low_delay_brc", "1", "-async_depth", "1"]
        return base

    def _publish_url(self, transport: str) -> str:
        quoted_name = urllib.parse.quote(self.stream_name.strip("/"))
        if transport == "whip":
            return f"http://127.0.0.1:{self.webrtc_port}/{quoted_name}/whip"
        return f"rtsp://127.0.0.1:{self.rtsp_port}/{quoted_name}"

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

    def _is_mediamtx_ready_for(self, transport: str) -> bool:
        webrtc_ready = _tcp_open("127.0.0.1", self.webrtc_port)
        if transport == "whip":
            return webrtc_ready
        return webrtc_ready and _tcp_open("127.0.0.1", self.rtsp_port)

    def _is_ffmpeg_running(self) -> bool:
        return self._ffmpeg_proc is not None and self._ffmpeg_proc.poll() is None

    def _reap_processes(self) -> None:
        if self._ffmpeg_proc is not None and self._ffmpeg_proc.poll() is not None:
            code = self._ffmpeg_proc.returncode
            if code not in (0, None):
                tail = " ".join(list(self._ffmpeg_log_tail)[-5:])
                self._last_error = f"FFmpeg exited with code {code}. {tail}".strip()
                self._log_event("ffmpeg_exited", code=code, tail=tail)
            self._ffmpeg_proc = None
        if self._mediamtx_proc is not None and self._mediamtx_proc.poll() is not None:
            code = self._mediamtx_proc.returncode
            if code not in (0, None):
                tail = " ".join(list(self._mediamtx_log_tail)[-5:])
                self._last_error = f"MediaMTX exited with code {code}. {tail}".strip()
                self._log_event("mediamtx_exited", code=code, tail=tail)
            self._mediamtx_proc = None
            self._mediamtx_owned = False

    def _terminate_process(self, attr_name: str) -> None:
        proc = getattr(self, attr_name, None)
        if proc is None:
            return
        try:
            if proc.poll() is None:
                self._log_event("process_terminating", process=attr_name)
                proc.terminate()
                try:
                    proc.wait(timeout=3.0)
                except subprocess.TimeoutExpired:
                    self._log_event("process_killing", process=attr_name)
                    proc.kill()
        except Exception:
            pass
        setattr(self, attr_name, None)

    def _spawn_log_reader(self, pipe, sink: Deque[str]) -> None:
        if pipe is None:
            return

        def read_lines() -> None:
            try:
                for line in pipe:
                    cleaned = str(line or "").strip()
                    if cleaned:
                        sink.append(cleaned)
                        self._log_event("process_output", line=cleaned)
            except Exception:
                pass

        threading.Thread(target=read_lines, daemon=True).start()

    def _redact_command(self, command: list[str]) -> list[str]:
        redacted: list[str] = []
        for value in command or []:
            text = str(value)
            if "://" in text and "@" in text:
                parsed = urllib.parse.urlparse(text)
                safe_netloc = parsed.hostname or ""
                if parsed.port:
                    safe_netloc = f"{safe_netloc}:{parsed.port}"
                text = urllib.parse.urlunparse(parsed._replace(netloc=safe_netloc))
            redacted.append(text)
        return redacted

    def _fps_floor_explanation(self, payload: dict) -> str:
        if not payload.get("available"):
            return str(payload.get("reason") or "Turbo dependencies are missing; JPEG fallback is active.")
        if not payload.get("active"):
            return str(payload.get("reason") or "Turbo stream has not started yet.")
        selected = payload.get("selected") or {}
        profile = selected.get("profile") or {}
        return (
            f"Turbo is active at {profile.get('name', 'unknown profile')} using "
            f"{(selected.get('capture') or {}).get('name', 'unknown capture')} and "
            f"{(selected.get('encoder') or {}).get('name', 'unknown encoder')}; "
            f"last limit reason: {self._last_limit_reason or 'benchmark-selected'}."
        )
