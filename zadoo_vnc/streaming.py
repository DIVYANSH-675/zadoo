"""Adaptive streaming policy and diagnostics for screen video."""
from __future__ import annotations

import logging
import math
import os
import time
from dataclasses import asdict, dataclass

from .config import env_int


@dataclass(frozen=True)
class StreamProfile:
    name: str
    status: str
    target_fps: int
    quality: int
    scale_div: int
    grayscale: bool = False


STREAM_LADDER = (
    StreamProfile("full-240-q52", "Max FPS", 240, 52, 1),
    StreamProfile("full-120-q65", "Fast", 120, 65, 1),
    StreamProfile("full-120-q56", "Fast", 120, 56, 1),
    StreamProfile("half-240-q54", "Fast", 240, 54, 2),
    StreamProfile("half-120-q54", "Fast", 120, 54, 2),
    StreamProfile("half-60-q60", "Balanced", 60, 60, 2),
    StreamProfile("half-60-q52", "Balanced", 60, 52, 2),
    StreamProfile("half-30-q54", "Balanced", 30, 54, 2),
    StreamProfile("half-15-q50", "Saving Data", 15, 50, 2),
    StreamProfile("third-30-q46", "Saving Data", 30, 46, 3),
    StreamProfile("third-15-q42", "Saving Data", 15, 42, 3),
    StreamProfile("third-10-q40", "Saving Data", 10, 40, 3),
)


class AdaptiveStreamController:
    """Keeps public-link streaming responsive by adapting bitrate before delay builds."""

    def __init__(self):
        self.enabled = True

        # Start conservative for an UNKNOWN link (safe at ~2 Mbps) and let the controller upshift
        # toward high FPS on fast/local links. Starting near the top of the ladder floods a slow
        # link for several seconds before it can converge down.
        start_name = str(os.getenv("ZADOO_STREAM_START_PROFILE", "half-60-q52")).strip().lower()
        self.profile_index = self._index_for_name(start_name)
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
        self.measured_fps_cap = None
        self.effective_target_fps = self.profile.target_fps
        self.fps_cap_reason = ""
        self.downshift_reason = "startup"
        self._last_applied_signature = None
        # Optional hard bandwidth target in kbps (e.g. 2000 for a 2 Mbps link). 0 = disabled, in
        # which case adaptation reacts to latency/backlog only. When set, the controller also
        # downshifts proactively whenever the estimated egress (frame_bytes x fps) exceeds it.
        self._target_kbps = env_int("ZADOO_TARGET_KBPS", 0, 0)

    def _index_for_name(self, name: str) -> int:
        for index, profile in enumerate(STREAM_LADDER):
            if profile.name.lower() == name:
                return index
        choices = ", ".join(profile.name for profile in STREAM_LADDER)
        raise ValueError(f"Invalid ZADOO_STREAM_START_PROFILE={name!r}; expected one of {choices}")

    @property
    def profile(self) -> StreamProfile:
        return STREAM_LADDER[self.profile_index]

    def apply_to(self, server) -> StreamProfile:
        profile = self.profile
        target_fps = int(profile.target_fps)
        if self.measured_fps_cap:
            target_fps = max(1, min(target_fps, int(self.measured_fps_cap)))
        self.effective_target_fps = target_fps
        server.current_fps = target_fps
        capturer = server.screen_capturer
        quality_locked = server._quality_locked_by_user
        if quality_locked:
            quality = server.current_quality
        else:
            quality = int(profile.quality)
            server.current_quality = quality
        scale_div = int(profile.scale_div)
        if quality >= 85:
            scale_div = 1
            self.quality_scale_lock = "full_resolution"
        elif quality >= 70:
            scale_div = min(scale_div, 2)
            self.quality_scale_lock = "balanced_resolution"
        else:
            self.quality_scale_lock = "adaptive"
        manual = server._manual_performance
        region = "full"
        rect_norm = None
        grayscale = bool(profile.grayscale)
        if manual:
            region, manual_scale, manual_grayscale, rect_norm = manual
            scale_div = max(scale_div, manual_scale)
            grayscale = grayscale or manual_grayscale
        self.effective_scale_div = scale_div
        if capturer is not None:
            capturer.fps = target_fps
            capturer.quality = quality
            capturer.configure_performance(bool(manual) or scale_div > 1, region, scale_div, grayscale, rect_norm)
        signature = (profile.name, target_fps, quality, scale_div, grayscale, region, quality_locked)
        if signature != self._last_applied_signature:
            logging.getLogger("streaming").info(
                "Stream profile applied profile=%s target_fps=%s quality=%s scale_div=%s grayscale=%s quality_locked=%s reason=%s",
                profile.name,
                target_fps,
                quality,
                scale_div,
                grayscale,
                quality_locked,
                self.last_reason,
            )
            self._last_applied_signature = signature
        return profile

    def record_client_stats(self, payload: dict) -> None:
        now = time.time()
        keys = (
            "display_fps",
            "decode_ms",
            "draw_ms",
            "recv_kbps",
            "dropped_blobs",
            "rtt_ms",
            "receive_delay_ms",
        )
        stats = {}
        for key in keys:
            try:
                raw = payload[key]
                if isinstance(raw, bool):
                    raise TypeError("boolean is not a number")
                value = float(raw)
            except (KeyError, TypeError, ValueError) as exc:
                raise ValueError(f"Invalid client stream statistic {key}: {exc}") from exc
            if not math.isfinite(value) or value < 0:
                raise ValueError(f"Client stream statistic {key} must be finite and non-negative")
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
        capture_stats: dict,
    ) -> bool:
        now = time.time()
        skipped_delta = max(0, skipped_total - self.last_skipped_total)
        self.last_skipped_total = skipped_total
        self.last_server_stats = {
            "frame_bytes": frame_bytes,
            "max_write_buffer": max_write_buffer,
            "skipped_delta": skipped_delta,
            "skipped_total": skipped_total,
            "inflight_sends": inflight_sends,
            "video_clients": video_clients,
            "frame_age_ms": frame_age_ms,
            "capture_ms": float(capture_stats["last_capture_ms"]),
            "encode_ms": float(capture_stats["last_encode_ms"]),
            "capture_fps": float(capture_stats["current_fps"]),
            "target_fps": int(capture_stats["target_fps"]),
            "quality": int(capture_stats["quality"]),
            "scale_div": int(capture_stats["perf_scale_div"]),
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
        quality = int(self.last_server_stats["quality"])
        full_resolution_locked = quality >= 85

        # Estimated egress for the current send rate (kbps). Used only when a target is configured.
        frame_bytes = int(self.last_server_stats["frame_bytes"])
        est_send_kbps = (frame_bytes * max(1, int(self.effective_target_fps)) * 8.0) / 1000.0
        bitrate_over = self._target_kbps > 0 and est_send_kbps > self._target_kbps * 1.1
        bitrate_far_over = self._target_kbps > 0 and est_send_kbps > self._target_kbps * 2.0

        network_backlog = max_write_buffer > 128_000 or skipped_delta > max(2, video_clients * 3)
        stale_frames = frame_age_ms > 220 or inflight_sends > video_clients
        browser_slow = bool(client) and (
            decode_draw_ms > max(12.0, budget_ms * 1.2)
            or dropped_blobs > 2
            or (display_fps > 0 and display_fps < target_fps * 0.55)
            or rtt_ms > 350
        )
        host_slow = capture_ms + encode_ms > max(10.0, budget_ms * 1.35)
        overloaded = network_backlog or stale_frames or browser_slow or host_slow or bitrate_over
        cap_reason = ""
        desired_cap = int(profile.target_fps)
        if host_slow:
            host_ms = max(1.0, capture_ms + encode_ms)
            desired_cap = min(desired_cap, max(15, int(1000.0 / (host_ms * 1.25))))
            if capture_fps > 0:
                desired_cap = min(desired_cap, max(15, int(capture_fps * 0.95)))
            cap_reason = "host_capture_encode"
        if browser_slow and display_fps > 0:
            desired_cap = min(desired_cap, max(15, int(display_fps * 0.95)))
            cap_reason = cap_reason or "browser_decode"
        if network_backlog:
            desired_cap = min(desired_cap, max(15, int(max(15, self.effective_target_fps) * 0.75)))
            cap_reason = cap_reason or "network_backlog"

        if overloaded:
            self.good_intervals = 0
            if full_resolution_locked:
                new_cap = max(15, min(int(profile.target_fps), int(desired_cap)))
                old_cap = int(self.measured_fps_cap or profile.target_fps)
                self.measured_fps_cap = min(old_cap, new_cap)
                self.fps_cap_reason = cap_reason or "overloaded"
                self.downshift_reason = self.fps_cap_reason
                self.last_reason = self.fps_cap_reason
                if self.measured_fps_cap != old_cap:
                    self.last_change_at = now
                    logging.getLogger("streaming").info(
                        "Stream FPS cap applied cap=%s reason=%s capture_ms=%.2f encode_ms=%.2f capture_fps=%.1f display_fps=%.1f max_buffer=%s skipped_delta=%s",
                        self.measured_fps_cap,
                        self.fps_cap_reason,
                        capture_ms,
                        encode_ms,
                        capture_fps,
                        display_fps,
                        max_write_buffer,
                        skipped_delta,
                    )
                return self.measured_fps_cap != old_cap
            # When the link is severely over budget, converge fast: shorten the inter-step guard
            # and allow a 2-rung jump so a slow (2 Mbps) link stops flooding within ~1s instead of
            # descending one rung every 1.5s.
            severe = (max_write_buffer > 512_000) or (skipped_delta > max(6, video_clients * 8)) or bitrate_far_over
            downshift_guard = 0.6 if severe else 1.5
            if self.profile_index < len(STREAM_LADDER) - 1 and now - self.last_change_at >= downshift_guard:
                step = 2 if (severe and self.profile_index < len(STREAM_LADDER) - 2) else 1
                self.profile_index = min(len(STREAM_LADDER) - 1, self.profile_index + step)
                self.last_change_at = now
                reasons = []
                if network_backlog:
                    reasons.append("network_backlog")
                if bitrate_over:
                    reasons.append("bitrate_over")
                if stale_frames:
                    reasons.append("stale_frames")
                if browser_slow:
                    reasons.append("browser_decode")
                if host_slow:
                    reasons.append("host_capture_encode")
                self.last_reason = ",".join(reasons) or "overloaded"
                self.downshift_reason = self.last_reason
                logging.getLogger("streaming").info(
                    "Stream downshift profile=%s reason=%s capture_ms=%.2f encode_ms=%.2f capture_fps=%.1f display_fps=%.1f max_buffer=%s skipped_delta=%s frame_age_ms=%.1f",
                    self.profile.name,
                    self.last_reason,
                    capture_ms,
                    encode_ms,
                    capture_fps,
                    display_fps,
                    max_write_buffer,
                    skipped_delta,
                    frame_age_ms,
                )
                return True
            self.last_reason = "overloaded"
            self.downshift_reason = cap_reason or self.last_reason
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

        if self.measured_fps_cap and self.good_intervals >= 6 and now - self.last_change_at >= 8.0:
            old_cap = int(self.measured_fps_cap)
            self.measured_fps_cap = min(int(profile.target_fps), old_cap + 15)
            if self.measured_fps_cap >= int(profile.target_fps):
                self.measured_fps_cap = None
                self.fps_cap_reason = ""
            self.good_intervals = 0
            self.last_change_at = now
            self.last_reason = "stable_headroom"
            self.downshift_reason = self.last_reason
            logging.getLogger("streaming").info("Stream FPS cap relaxed cap=%s", self.measured_fps_cap)
            return True

        bitrate_ok = self._target_kbps <= 0 or est_send_kbps < self._target_kbps * 0.8
        if self.good_intervals >= 6 and self.profile_index > 0 and now - self.last_change_at >= 8.0 and bitrate_ok:
            self.profile_index -= 1
            self.good_intervals = 0
            self.last_change_at = now
            self.last_reason = "stable_headroom"
            self.downshift_reason = self.last_reason
            logging.getLogger("streaming").info("Stream upshift profile=%s reason=stable_headroom", self.profile.name)
            return True
        self.last_reason = "stable"
        return False

    def status(self) -> dict:
        profile = self.profile
        return {
            "enabled": self.enabled,
            "profile": asdict(profile),
            "profile_name": profile.name,
            "status": profile.status,
            "effective_target_fps": self.effective_target_fps,
            "measured_fps_cap": self.measured_fps_cap,
            "fps_cap_reason": self.fps_cap_reason,
            "downshift_reason": self.downshift_reason,
            "effective_scale_div": self.effective_scale_div,
            "quality_scale_lock": self.quality_scale_lock,
            "last_reason": self.last_reason,
            "ladder": [asdict(item) for item in STREAM_LADDER],
            "server": dict(self.last_server_stats),
            "client": dict(self.last_client_stats),
            "client_age_ms": max(0, (time.time() - self.last_client_at) * 1000.0) if self.last_client_at else None,
        }
