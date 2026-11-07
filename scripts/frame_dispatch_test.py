"""Regression checks for frame sends that complete while awaiting capture."""
from __future__ import annotations

import asyncio
import sys
import time
from pathlib import Path
from types import SimpleNamespace

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from zadoo_vnc.media import MediaMixin  # noqa: E402


class Capture:
    def __init__(self, delta=False):
        self.is_running = True
        self.last_error = None
        self.sequence = 0
        self.consumed = asyncio.Event()
        self.prefix = b"ZDP1" if delta else b""

    def get_latest_frame_packet(self):
        self.consumed.set()
        return self.sequence, self.prefix + str(self.sequence).encode()


class Client:
    def __init__(self, blocked=False):
        self.frames = []
        self.started = 0
        self.transport = SimpleNamespace(get_write_buffer_size=lambda: 0)
        self.release = asyncio.Event() if blocked else None

    async def send(self, frame):
        self.started += 1
        if self.release is not None:
            await self.release.wait()
        self.frames.append(frame)


async def check_dispatch(with_slow_client, delta=False):
    capture = Capture(delta)
    fast = Client()
    slow = Client(blocked=True) if with_slow_client else None
    clients = {fast, slow} if slow else {fast}
    server = SimpleNamespace(
        stop_event=asyncio.Event(),
        frame_ready_event=asyncio.Event(),
        _video_send_tasks={},
        _video_skipped_sends=0,
        video_clients=clients,
        _delta_video_clients=clients.copy() if delta else set(),
        screen_capturer=capture,
        adaptive_stream=SimpleNamespace(last_eval_at=time.time() + 100000),
    )
    task = asyncio.create_task(MediaMixin.broadcast_frames(server))
    try:
        for sequence in range(1, 21):
            capture.consumed.clear()
            capture.sequence = sequence
            server.frame_ready_event.set()
            await asyncio.wait_for(capture.consumed.wait(), 2)
            # Complete the fast socket send before capture produces another frame.
            await asyncio.sleep(0)
            await asyncio.sleep(0)
        if fast.frames != [capture.prefix + str(sequence).encode() for sequence in range(1, 21)]:
            raise AssertionError(f"Fast viewer lost frames: {len(fast.frames)}/20")
        if slow:
            if slow.started != 1 or slow.frames:
                raise AssertionError("Slow viewer accumulated queued frame sends")
            if len(server._video_send_tasks) > len(clients):
                raise AssertionError("In-flight sends exceeded the viewer bound")
            if server._video_skipped_sends != 19:
                raise AssertionError("Slow-client skips included a fast viewer")
        elif server._video_skipped_sends:
            raise AssertionError("A completed send was counted as network backpressure")
    finally:
        server.stop_event.set()
        server.frame_ready_event.set()
        if slow:
            slow.release.set()
        await asyncio.wait_for(task, 2)
    label = "slow viewer isolated, one pending send" if slow else "no false backpressure"
    print(f"FRAME_DISPATCH: PASS 20/20 fast frames; delta={delta}; {label}")


async def main():
    await check_dispatch(False)
    await check_dispatch(True)
    await check_dispatch(False, delta=True)
    await check_dispatch(True, delta=True)


if __name__ == "__main__":
    asyncio.run(main())

