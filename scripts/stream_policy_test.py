"""Deterministic adaptation checks: preserve detail without hiding real overload."""
from __future__ import annotations

import os
import sys
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from zadoo_vnc.streaming import AdaptiveStreamController


def check(condition, message):
    if not condition:
        raise AssertionError(message)


class Simulation:
    def __init__(self, target_kbps=0):
        self.now = 100.0
        with patch.dict(os.environ, {"ZADOO_STREAM_START_PROFILE": "full-60-q85", "ZADOO_TARGET_KBPS": str(target_kbps)}):
            self.controller = AdaptiveStreamController()

    def tick(self, *, produced=30, displayed=30, cost=5, buffer=0, skipped=0, frame_bytes=10000, dropped=0, viewers=1, produced_kbps=None):
        self.now += 1.1
        controller = self.controller
        profile = controller.profile
        with patch("zadoo_vnc.streaming.time.time", return_value=self.now):
            controller.record_client_stats({"display_fps": displayed, "decode_ms": 1, "draw_ms": .1,
                "recv_kbps": 100, "dropped_blobs": dropped, "rtt_ms": 10, "receive_delay_ms": 0})
            capture_stats = {"quality":profile.quality,"last_capture_ms":cost-1,"last_encode_ms":1,
                "current_fps":profile.target_fps,"produced_fps":produced,"target_fps":profile.target_fps,
                "perf_scale_div":profile.scale_div}
            if produced_kbps is not None:
                capture_stats["produced_kbps"] = produced_kbps
            return controller.observe_server(frame_bytes=frame_bytes, max_write_buffer=buffer,
                skipped_total=skipped, inflight_sends=0, video_clients=viewers, frame_age_ms=0,
                capture_stats=capture_stats)


def main():
    for produced in (1, 30):
        sim = Simulation()
        initial = sim.controller.profile_index
        for _ in range(12):
            sim.tick(produced=produced, displayed=produced)
        check(sim.controller.profile_index == initial, "Quiet/source-limited stream lost clarity")
    print("STREAM_POLICY: PASS quiet and source-limited desktops keep full-resolution clarity")

    sim = Simulation()
    initial = sim.controller.profile_index
    sim.tick(cost=500)
    for _ in range(8):
        sim.tick()
    check(sim.controller.profile_index == initial, "Cold initialization poisoned sustained cost estimates")
    print("STREAM_POLICY: PASS cold capture initialization does not poison the cost estimate")

    sim = Simulation()
    initial = sim.controller.profile_index
    for _ in range(5):
        sim.tick(cost=500, viewers=0)
    for _ in range(8):
        sim.tick()
    check(sim.controller.profile_index == initial, "No-viewer timing changed profile")
    sim.controller.begin_session(12)
    check(sim.controller.last_skipped_total == 12 and not sim.controller.last_client_stats,
        "New session retained the previous viewer's telemetry")
    check(sim.controller.profile_index == initial, "New session changed selected settings")
    print("STREAM_POLICY: PASS idle timing and previous-viewer telemetry are excluded")

    for name, produced in (("half-60-q60", 30), ("half-30-q54", 1)):
        sim = Simulation()
        sim.controller.profile_index = sim.controller._index_for_name(name)
        for _ in range(10):
            sim.tick(produced=produced, displayed=produced)
        check(sim.controller.profile.scale_div == 1, "Clarity recovery was blocked by an unreachable faster tier")
        check(sim.controller.profile.quality >= 75, "Clarity recovery did not restore readable quality")
        sim.controller.begin_session()
        check(sim.controller.profile.name == "full-60-q85", "Fresh viewer inherited a stale downshift")
    print("STREAM_POLICY: PASS clarity recovers without requiring an unreachable FPS tier")

    for baseline_cost in (5, 18):
        sim = Simulation()
        initial = sim.controller.profile_index
        for _ in range(5):
            sim.tick(cost=baseline_cost)
        sim.tick(cost=100)
        for _ in range(8):
            sim.tick(cost=baseline_cost)
        check(sim.controller.profile_index == initial, "One host spike changed profile")
    print("STREAM_POLICY: PASS isolated host spike does not cause a resolution/FPS collapse")

    sim = Simulation()
    for _ in range(4):
        sim.tick()
    initial = sim.controller.profile_index
    for _ in range(4):
        sim.tick(cost=100)
    check(sim.controller.profile_index > initial, "Sustained host overload was ignored")
    check(sim.controller.profile.target_fps <= 60, "Downshift increased capture work")
    print("STREAM_POLICY: PASS sustained overload downshifts without raising requested FPS")

    sim = Simulation()
    initial = sim.controller.profile_index
    sim.tick(buffer=600000, skipped=10)
    check(sim.controller.profile_index > initial, "Network backlog did not downshift immediately")
    check(sim.controller.profile.target_fps <= 60, "Network downshift increased capture work")
    print("STREAM_POLICY: PASS slow-link backlog still downshifts immediately")

    sim = Simulation(target_kbps=2000)
    initial = sim.controller.profile_index
    sim.tick(frame_bytes=100000)
    check(sim.controller.profile_index > initial, "Configured bandwidth ceiling was ignored")
    print("STREAM_POLICY: PASS configured bandwidth ceiling is still enforced")

    sim = Simulation(target_kbps=2000)
    initial = sim.controller.profile_index
    for _ in range(10):
        sim.tick(frame_bytes=150000, produced_kbps=1000)
    check(sim.controller.profile_index == initial, "An occasional keyframe was mistaken for continuous full-frame bitrate")
    sim.tick(frame_bytes=150000, produced_kbps=24000)
    check(sim.controller.profile_index > initial, "Actual measured bitrate ceiling was ignored")
    print("STREAM_POLICY: PASS incremental bitrate uses measured bytes, not keyframe size times nominal FPS")

    sim = Simulation()
    initial = sim.controller.profile_index
    sim.controller._quality_locked_by_user = True
    sim.tick(buffer=600000, skipped=10)
    check(sim.controller.profile_index == initial and sim.controller.measured_fps_cap < 60,
        "Explicit high-quality lock did not retain full resolution and cap FPS")
    print("STREAM_POLICY: PASS explicit quality lock remains distinct from the adaptive quality-85 default")

    sim = Simulation()
    initial = sim.controller.profile_index
    for _ in range(3):
        sim.tick()
    sim.tick(produced=60, displayed=10, dropped=10)
    sim.tick(produced=60, displayed=10, dropped=10)
    check(sim.controller.profile_index > initial, "Persistently slow browser was ignored")
    print("STREAM_POLICY: PASS real browser overload still downshifts")


if __name__ == "__main__":
    main()
