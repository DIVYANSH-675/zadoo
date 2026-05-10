"""Adaptive streaming policy and diagnostics for screen video."""
from __future__ import annotations

import os
import platform
import subprocess
import time
from dataclasses import asdict, dataclass


FALSE_VALUES = {"0", "false", "no", "off", "disabled"}


@dataclass(frozen=True)
class StreamProfile:
    name: str
    status: str
    target_fps: int
    quality: int
    scale_div: int
    grayscale: bool = False


STREAM_LADDER = (
    StreamProfile("1080p120", "Fast", 120, 65, 1),
    StreamProfile("900p120", "Fast", 120, 56, 1),
    StreamProfile("720p120", "Fast", 120, 54, 2),
    StreamProfile("720p60", "Balanced", 60, 60, 2),
    StreamProfile("540p60", "Balanced", 60, 52, 2),
    StreamProfile("360p30", "Saving Data", 30, 46, 3),
    StreamProfile("360p15", "Saving Data", 15, 40, 3),
)


def _env_enabled(name: str, default: str = "1") -> bool:
    return str(os.getenv(name, default)).strip().lower() not in FALSE_VALUES


def _hidden_creationflags() -> int:
    if platform.system().lower() != "windows":
        return 0
    return getattr(subprocess, "CREATE_NO_WINDOW", 0)


def _detect_gpu_names() -> list[str]:
    if platform.system().lower() != "windows":
        return []
    command = (
        "Get-CimInstance Win32_VideoController | "
        "ForEach-Object { $_.Name }"
    )
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", command],
            capture_output=True,
            text=True,
            timeout=2.0,
            creationflags=_hidden_creationflags(),
        )
    except Exception:
        return []
    names = []
    for line in (result.stdout or "").splitlines():
        cleaned = line.strip()
        if cleaned:
            names.append(cleaned)
    return names


def detect_encoder_capabilities() -> dict:
    """Detect likely hardware encoder options without requiring them to work."""
    gpu_names = _detect_gpu_names()
    joined = " ".join(gpu_names).lower()
    has_nvidia = "nvidia" in joined
    has_intel = "intel" in joined
    has_amd = any(token in joined for token in ("amd", "radeon", "advanced micro devices"))
    preferred = "software_jpeg"
    if has_nvidia:
        preferred = "nvenc_candidate"
    elif has_intel:
        preferred = "quick_sync_candidate"
    elif has_amd:
        preferred = "amf_candidate"
    return {
        "platform": platform.system() or "unknown",
        "gpu_names": gpu_names,
        "nvenc_candidate": has_nvidia,
        "quick_sync_candidate": has_intel,
        "amf_candidate": has_amd,
        "preferred_video_encoder": preferred,
        "active_encoder_backend": "software_jpeg",
    }


class AdaptiveStreamController:
    """Keeps public-link streaming responsive by adapting bitrate before delay builds."""

    def __init__(self, encoder_capabilities: dict | None = None):
        self.enabled = _env_enabled("ZADOO_ADAPTIVE_STREAM", "1")
        self.encoder_capabilities = dict(encoder_capabilities or detect_encoder_capabilities())
        requested = str(os.getenv("ZADOO_STREAM_MODE", "adaptive_jpeg_ws")).strip().lower()
        self.requested_transport = requested or "adaptive_jpeg_ws"
        self.webrtc_configured = bool(
            os.getenv("ZADOO_WEBRTC_TURN_URL")
            or os.getenv("ZADOO_WEBRTC_SFU_URL")
            or os.getenv("ZADOO_WEBRTC_SIGNALING_URL")
        )
        self.transport_mode = "jpeg_ws"
        self.fallback_reason = ""
        if self.requested_transport.startswith("webrtc") and not self.webrtc_configured:
            self.fallback_reason = "WebRTC transport requested but no TURN, SFU, or signaling URL is configured."
        elif self.requested_transport.startswith("webrtc"):
            self.fallback_reason = "WebRTC media mode is staged; adaptive JPEG WebSocket remains active until the media bridge is enabled."

        start_name = str(os.getenv("ZADOO_STREAM_START_PROFILE", "720p120")).strip().lower()
        self.profile_index = self._index_for_name(start_name, default=2)
        self.last_change_at = 0.0
        self.last_eval_at = 0.0
        self.good_intervals = 0
        self.last_reason = "startup"
        self.last_skipped_total = 0
        self.last_server_stats = {}
        self.last_client_stats = {}
        self.last_client_at = 0.0
        self.effective_scale_div = self.profile.scale_div
        self.quality_scale_lock = "adaptive"

    def _index_for_name(self, name: str, default: int = 0) -> int:
        for index, profile in enumerate(STREAM_LADDER):
            if profile.name.lower() == name:
                return index
        return max(0, min(len(STREAM_LADDER) - 1, default))

    @property
    def profile(self) -> StreamProfile:
        return STREAM_LADDER[self.profile_index]

    def apply_to(self, server) -> StreamProfile:
        profile = self.profile
        server.current_fps = int(profile.target_fps)
        capturer = getattr(server, "screen_capturer", None)
        quality = int(getattr(server, "current_quality", 65) or 65)
        scale_div = int(profile.scale_div)
        if quality >= 85:
            scale_div = 1
            self.quality_scale_lock = "full_resolution"
        elif quality >= 70:
            scale_div = min(scale_div, 2)
            self.quality_scale_lock = "balanced_resolution"
        else:
            self.quality_scale_lock = "adaptive"
        self.effective_scale_div = scale_div
        if capturer is not None:
            capturer.fps = int(profile.target_fps)
            try:
                capturer.set_performance_mode(scale_div > 1, "full", scale_div)
                capturer.set_grayscale(bool(profile.grayscale))
            except Exception:
                pass
        return profile

    def record_client_stats(self, payload: dict) -> None:
        now = time.time()
        stats = {}
        for key in (
            "display_fps",
            "decode_ms",
            "draw_ms",
            "recv_kbps",
            "dropped_blobs",
            "rtt_ms",
            "receive_delay_ms",
        ):
            try:
                value = float(payload.get(key, 0) or 0)
            except Exception:
                value = 0.0
            stats[key] = value
        self.last_client_stats = stats
        self.last_client_at = now

    def observe_server(
        self,
        *,
        frame_bytes: int,
        max_write_buffer: int,
        skipped_total: int,
        inflight_sends: int,
        video_clients: int,
        frame_age_ms: float,
        capture_stats: dict | None = None,
    ) -> bool:
        now = time.time()
        skipped_delta = max(0, int(skipped_total) - int(self.last_skipped_total))
        self.last_skipped_total = int(skipped_total)
        capture_stats = dict(capture_stats or {})
        self.last_server_stats = {
            "frame_bytes": int(frame_bytes or 0),
            "max_write_buffer": int(max_write_buffer or 0),
            "skipped_delta": skipped_delta,
            "skipped_total": int(skipped_total or 0),
            "inflight_sends": int(inflight_sends or 0),
            "video_clients": int(video_clients or 0),
            "frame_age_ms": float(frame_age_ms or 0.0),
            "capture_ms": float(capture_stats.get("last_capture_ms") or 0.0),
            "encode_ms": float(capture_stats.get("last_encode_ms") or 0.0),
            "capture_fps": float(capture_stats.get("current_fps") or 0.0),
        }
        if not self.enabled or now - self.last_eval_at < 1.0:
            return False
        self.last_eval_at = now

        profile = self.profile
        target_fps = max(1, int(profile.target_fps))
        budget_ms = 1000.0 / target_fps
        client = self.last_client_stats if now - self.last_client_at < 4.0 else {}
        decode_draw_ms = float(client.get("decode_ms", 0.0)) + float(client.get("draw_ms", 0.0))
        display_fps = float(client.get("display_fps", 0.0))
        dropped_blobs = float(client.get("dropped_blobs", 0.0))
        rtt_ms = float(client.get("rtt_ms", 0.0))
        capture_ms = self.last_server_stats["capture_ms"]
        encode_ms = self.last_server_stats["encode_ms"]
        capture_fps = self.last_server_stats["capture_fps"]

        network_backlog = max_write_buffer > 128_000 or skipped_delta > max(2, video_clients * 3)
        stale_frames = frame_age_ms > 220 or inflight_sends > video_clients
        browser_slow = bool(client) and (
            decode_draw_ms > max(12.0, budget_ms * 1.2)
            or dropped_blobs > 2
            or (display_fps > 0 and display_fps < target_fps * 0.55)
            or rtt_ms > 350
        )
        host_slow = (
            capture_ms + encode_ms > max(10.0, budget_ms * 1.35)
            or (capture_fps > 0 and capture_fps < target_fps * 0.55)
        )
        overloaded = network_backlog or stale_frames or browser_slow or host_slow

        if overloaded:
            self.good_intervals = 0
            if self.profile_index < len(STREAM_LADDER) - 1 and now - self.last_change_at >= 1.5:
                self.profile_index += 1
                self.last_change_at = now
                reasons = []
                if network_backlog:
                    reasons.append("network_backlog")
                if stale_frames:
                    reasons.append("stale_frames")
                if browser_slow:
                    reasons.append("browser_decode")
                if host_slow:
                    reasons.append("host_capture_encode")
                self.last_reason = ",".join(reasons) or "overloaded"
                return True
            self.last_reason = "overloaded"
            return False

        client_ok = not client or (
            display_fps <= 0
            or display_fps >= target_fps * 0.8
        )
        host_ok = capture_fps <= 0 or capture_fps >= target_fps * 0.8
        if client_ok and host_ok:
            self.good_intervals += 1
        else:
            self.good_intervals = 0

        if self.good_intervals >= 6 and self.profile_index > 0 and now - self.last_change_at >= 8.0:
            self.profile_index -= 1
            self.good_intervals = 0
            self.last_change_at = now
            self.last_reason = "stable_headroom"
            return True
        self.last_reason = "stable"
        return False

    def status(self) -> dict:
        profile = self.profile
        payload = {
            "enabled": self.enabled,
            "transport_mode": self.transport_mode,
            "requested_transport": self.requested_transport,
            "fallback_reason": self.fallback_reason,
            "webrtc_configured": self.webrtc_configured,
            "profile": asdict(profile),
            "profile_name": profile.name,
            "status": profile.status,
            "effective_scale_div": self.effective_scale_div,
            "quality_scale_lock": self.quality_scale_lock,
            "last_reason": self.last_reason,
            "ladder": [asdict(item) for item in STREAM_LADDER],
            "encoder_capabilities": self.encoder_capabilities,
            "server": dict(self.last_server_stats),
            "client": dict(self.last_client_stats),
            "client_age_ms": max(0, (time.time() - self.last_client_at) * 1000.0) if self.last_client_at else None,
        }
        return payload
