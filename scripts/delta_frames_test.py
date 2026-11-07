"""Real-codec checks for bounded patches, fidelity, recovery and compatibility."""
from __future__ import annotations

import sys
import time
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import numpy as np
from imagecodecs._jpeg8 import jpeg8_decode

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from zadoo_vnc.delta_frames import FRAME_HEADER, DeltaFrameEncoder
from zadoo_vnc.screen_capture import ScreenCapturer, _OwnedFrameProcessor


def check(condition, message):
    if not condition:
        raise AssertionError(message)


def unpack(packet):
    return FRAME_HEADER.unpack_from(packet), jpeg8_decode(packet[FRAME_HEADER.size:])


def main():
    # Import capture first: BetterCam/comtypes and audio select different COM modes.
    from zadoo_vnc.media import MediaMixin, _video_format_from_path

    for path, expected in (("/video", "jpeg"), ("/video?format=jpeg", "jpeg"),
                           ("/video?format=jpeg-patch-v1", "jpeg-patch-v1")):
        check(_video_format_from_path(path) == expected, "Video handshake selected the wrong format")
    for path, expected in (("/video?format=unknown", "Video format must be jpeg or jpeg-patch-v1"),
                           ("/video?format=jpeg&format=jpeg", "Video query must contain at most one format parameter"),
                           ("/video?unknown=1", "Video query must contain at most one format parameter")):
        try:
            _video_format_from_path(path)
        except ValueError as exc:
            check(str(exc) == expected, "Invalid handshake returned the wrong exact error")
        else:
            raise AssertionError("Invalid video handshake was accepted")
    print("DELTA: PASS initial negotiation and exact invalid-query errors")

    capture = ScreenCapturer(quality=85)
    capture.set_delta_streaming(True)
    source = np.full((224, 320, 4), 255, dtype=np.uint8)
    source[24:64, 20:300:4] = (255, 0, 0, 255)
    source[80:120, 20:300:4] = (0, 0, 255, 255)
    with patch("zadoo_vnc.delta_frames.time.monotonic", return_value=100):
        header, display = unpack(capture._encode_stream_frame(source))
        check(header == (b"ZDF1", 1, 0, 320, 224, 0, 0), "Initial frame was not an independent keyframe")
        check(capture._encode_stream_frame(source) is None, "Unchanged pixels were encoded again")
        previous_sequence = header[1]
        sizes = []
        for x in (*range(30, 220, 10), *range(200, 20, -10)):
            moving = source.copy()
            moving[160:188, x:x+20] = (32, 160, 48, 255)
            packet = capture._encode_stream_frame(moving)
            header, pixels = unpack(packet)
            magic, sequence, base, width, height, left, top = header
            check(magic == b"ZDP1" and base == previous_sequence, "Patch did not reference its exact predecessor")
            check((width, height) == (320, 224), "Patch changed the canvas dimensions")
            ph, pw = pixels.shape[:2]
            check(pw*ph < width*height*.65 and left+pw <= width and top+ph <= height, "Patch was not bounded")
            display[top:top+ph, left:left+pw] = pixels
            independent = jpeg8_decode(capture._encode_frame(moving))
            check(float(np.abs(display.astype(float)-independent).mean()) < .2,
                "Moving patches accumulated seams or left a stale object behind")
            previous_sequence = sequence
            sizes.append(len(packet))
        check(max(sizes) < len(capture._encode_frame(source))*.6, "Small changes did not reduce encoded bytes")
        capture.request_keyframe()
        check(unpack(capture._encode_stream_frame(moving))[0][0] == b"ZDF1", "Explicit recovery did not produce a keyframe")
        capture.quality = 90
        check(unpack(capture._encode_stream_frame(moving))[0][0] == b"ZDF1", "Quality change retained an incompatible base")
        capture.configure_performance(True, "full", 2, True)
        header, gray = unpack(capture._encode_stream_frame(moving))
        check(header[0] == b"ZDF1" and header[3:5] == (160, 112) and gray.ndim == 2,
            "Scale/grayscale changes were not independently decodable")
    with patch("zadoo_vnc.delta_frames.time.monotonic", return_value=103):
        check(unpack(capture._encode_stream_frame(moving))[0][0] == b"ZDF1", "Periodic recovery keyframe was not emitted")
    print("DELTA: PASS real JPEG reconstruction, no moving-object ghosts, unchanged suppression and recovery")

    owner = np.full((8, 8, 4), 99, dtype=np.uint8)
    processor = _OwnedFrameProcessor(SimpleNamespace(process=lambda: owner))
    owned = processor.process()
    owner[:] = 0
    check(np.all(owned == 99), "Frame escaped the mapped surface without owning its pixels")
    print("DELTA: PASS mapped DXGI pixels are copied before unmap")

    motion = ScreenCapturer(quality=85)
    motion.set_delta_streaming(True)
    motion._encode_stream_frame(source.copy())
    inverted = 255-source
    inverted[:, :, 3] = 255
    check(unpack(motion._encode_stream_frame(inverted.copy()))[0][0] == b"ZDF1", "Large motion was not independently decodable")
    check(motion._last_jpeg_subsampling == '420' and motion._delta_encoder.detail_refresh_pending,
        "Large motion did not use the faster path or schedule detail recovery")
    motion.request_keyframe()
    motion._encode_stream_frame(inverted.copy())
    check(motion._last_jpeg_subsampling == '444' and not motion._delta_encoder.detail_refresh_pending,
        "Recovery keyframe did not restore full colour detail")
    motion.quality_locked_by_user = True
    motion._encode_stream_frame(source.copy())
    check(motion._last_jpeg_subsampling == '444', "Explicit high-quality selection lost full colour detail")
    print("DELTA: PASS content-adaptive colour, full-detail recovery and explicit quality lock")

    motion.quality_locked_by_user = False
    motion._encode_stream_frame(inverted.copy())
    check(motion._delta_encoder.detail_refresh_pending, "Idle-refresh test did not enter motion mode")
    motion.capture_stats['last_frame_ts'] = time.time()-1
    motion._initial_frame_pending = False
    def static_grab():
        if motion._initial_frame_pending:
            motion._initial_frame_pending = False
            return inverted.copy()
        return None
    motion._grab_screen_bettercam = static_grab
    motion.fps = 60
    motion.start()
    try:
        deadline = time.monotonic()+2
        while motion.capture_stats['sequence'] == 0 and time.monotonic() < deadline:
            time.sleep(.01)
        check(motion.capture_stats['sequence'] > 0 and motion._last_jpeg_subsampling == '444'
            and not motion._delta_encoder.detail_refresh_pending, "Static desktop did not automatically recover full detail")
    finally:
        motion.stop()
        motion.join(timeout=2)
    check(not motion.is_alive(), "Idle-refresh capture thread did not stop")
    print("DELTA: PASS capture loop refreshes full detail after motion stops, then stops cleanly")

    capture.configure_performance(False, "full", 1, False)
    delta, legacy = object(), object()
    server = SimpleNamespace(video_clients={delta}, _delta_video_clients={delta}, screen_capturer=capture)
    MediaMixin._sync_video_format(server)
    check(capture._delta_enabled, "Negotiated viewer did not enable updates")
    server.video_clients.add(legacy)
    MediaMixin._sync_video_format(server)
    check(not capture._delta_enabled, "Legacy viewer received a negotiated protocol")
    check(capture._encode_stream_frame(source).startswith(b"\xff\xd8"), "JPEG fallback is not an ordinary JPEG")
    server.video_clients.remove(legacy)
    MediaMixin._sync_video_format(server)
    check(unpack(capture._encode_stream_frame(source))[0][0] == b"ZDF1", "Leaving legacy mode did not reset the base")
    print("DELTA: PASS mixed-viewer JPEG fallback and safe return to incremental updates")

    encoder = DeltaFrameEncoder()
    with patch("zadoo_vnc.delta_frames.time.monotonic", return_value=100):
        try:
            encoder.encode(source, lambda _pixels: b"", (85,))
        except RuntimeError as exc:
            check(str(exc) == "JPEG encoder returned an empty or non-bytes frame", "Wrong encoder error")
        else:
            raise AssertionError("Empty JPEG was accepted")
    check(encoder.sequence == 0 and encoder.previous is None, "Failed encode advanced its base")
    print("DELTA: PASS exact encoder failure leaves the recovery base unchanged")


if __name__ == "__main__":
    main()
