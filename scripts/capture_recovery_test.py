"""Fault-injection tests; never locks Windows or captures the user's desktop."""
from __future__ import annotations

import sys
import threading
import unittest
from pathlib import Path
from types import MethodType, SimpleNamespace
from unittest.mock import Mock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import numpy as np
from comtypes import COMError

from zadoo_vnc import dxgi_capture
from zadoo_vnc.screen_capture import ScreenCapturer


def access_denied():
    return COMError(-2147024891, "Synthetic desktop access denied", None)


class CaptureRecoveryTests(unittest.TestCase):
    def test_output_loss_returns_control_instead_of_entering_library_retry_loop(self):
        camera = SimpleNamespace(_duplicator=SimpleNamespace(update_frame=lambda: False))
        camera._on_output_change = MethodType(dxgi_capture.RecoverableBetterCam._on_output_change, camera)
        with patch("bettercam.bettercam.time.sleep") as sleep, self.assertRaisesRegex(
            dxgi_capture.CaptureUnavailable, "DXGI_ERROR_ACCESS_LOST"
        ):
            dxgi_capture.RecoverableBetterCam._grab(camera, (0, 0, 16, 16))
        sleep.assert_not_called()

    def test_retry_delay_increases_and_total_recovery_is_bounded(self):
        capture = ScreenCapturer()
        capture._recovery_timeout = 1.0
        clock = [100.0]
        delays = []

        def wait(delay):
            delays.append(delay)
            clock[0] += delay
            return False

        with patch.object(capture, "_grab_screen_bettercam", side_effect=access_denied()) as grab, \
             patch.object(capture, "_release_bettercam"), \
             patch.object(capture._stop_capture_event, "wait", side_effect=wait), \
             patch("zadoo_vnc.screen_capture.time.monotonic", side_effect=lambda: clock[0]), \
             patch("zadoo_vnc.screen_capture.logging.warning"), patch("zadoo_vnc.screen_capture.logging.error"):
            capture.run()
        self.assertEqual(delays, [0.25, 0.5, 0.25])
        self.assertEqual(grab.call_count, 4)
        self.assertFalse(capture.is_running)
        self.assertIn("recovery failed after 1 seconds", capture.last_error)
        self.assertIn("E_ACCESSDENIED (0x80070005)", capture.last_error)
        self.assertIn("Synthetic desktop access denied", capture.last_error)
        self.assertEqual(capture.get_capture_stats()["capture_state"], "failed")

    def test_transient_failure_recovers_with_a_fresh_keyframe(self):
        capture = ScreenCapturer(fps=60, quality=85)
        capture.set_delta_streaming(True)
        pixels = np.full((16, 16, 4), 127, dtype=np.uint8)
        calls = 0

        def grab():
            nonlocal calls
            calls += 1
            if calls == 1:
                raise access_denied()
            if calls == 2:
                self.assertEqual(capture.get_capture_stats()["capture_state"], "recovering")
                self.assertFalse(capture.get_capture_stats()["is_working"])
                return pixels
            capture.stop()
            return None

        with patch.object(capture, "_grab_screen_bettercam", side_effect=grab), \
             patch.object(capture, "_release_bettercam"), \
             patch.object(capture._stop_capture_event, "wait", return_value=False), \
             patch("zadoo_vnc.screen_capture.time.sleep"), patch("zadoo_vnc.screen_capture.logging.warning"):
            capture.run()
        self.assertIsNone(capture.last_error)
        self.assertEqual(capture._recovery_attempts, 0)
        self.assertEqual(capture.capture_stats["sequence"], 1)
        self.assertTrue(capture.latest_frame_jpeg.startswith(b"ZDF1"))
        self.assertEqual((capture.fps, capture.quality, capture.perf_scale_div), (60, 85, 1))

    def test_stop_cancels_recovery_wait(self):
        capture = ScreenCapturer()
        entered = threading.Event()

        def fail_capture():
            entered.set()
            raise access_denied()

        with patch.object(capture, "_grab_screen_bettercam", side_effect=fail_capture), \
             patch.object(capture, "_release_bettercam"), patch("zadoo_vnc.screen_capture.logging.warning"):
            capture.start()
            self.assertTrue(entered.wait(1.0))
            capture.stop()
            capture.join(timeout=1.0)
        self.assertFalse(capture.is_alive())
        self.assertFalse(capture.is_running)

    def test_nonrecoverable_com_failure_is_not_hidden_or_retried(self):
        capture = ScreenCapturer()
        invalid = COMError(-2147024809, "Synthetic invalid argument", None)
        with patch.object(capture, "_grab_screen_bettercam", side_effect=invalid), \
             patch.object(capture, "_release_bettercam"), \
             patch.object(capture._stop_capture_event, "wait") as wait, \
             patch("zadoo_vnc.screen_capture.logging.exception"):
            capture.run()
        wait.assert_not_called()
        self.assertIn("COM_ERROR (0x80070057)", capture.last_error)
        self.assertIn("Synthetic invalid argument", capture.last_error)

    def test_snapshot_cannot_mask_a_still_broken_dxgi_device(self):
        capture = ScreenCapturer()
        capture._recovery_started = 10.0
        capture.bettercam_camera = SimpleNamespace(grab=Mock(side_effect=access_denied()))
        with patch.object(capture, "_capture_initial_frame") as snapshot, self.assertRaises(COMError):
            capture._grab_screen_bettercam()
        snapshot.assert_not_called()
        self.assertTrue(capture._initial_frame_pending)

    def test_static_desktop_is_not_treated_as_capture_loss(self):
        capture = ScreenCapturer()
        capture._initial_frame_pending = False
        capture.bettercam_camera = SimpleNamespace(width=16, height=16, grab=Mock(return_value=None))
        with patch.object(capture, "_recover_capture") as recover:
            for _ in range(20):
                self.assertIsNone(capture._grab_screen_bettercam())
        recover.assert_not_called()
        self.assertIsNone(capture.last_error)

    def test_recovery_enumerates_new_devices_instead_of_using_cached_factory(self):
        adapters = [object(), object()]
        device = SimpleNamespace(enum_outputs=lambda: [object()])
        output = SimpleNamespace(attached_to_desktop=True, resolution=(1920, 1080))
        with patch.object(dxgi_capture, "enum_dxgi_adapters", side_effect=[[adapters[0]], [adapters[1]]]), \
             patch.object(dxgi_capture, "Device", return_value=device) as make_device, \
             patch.object(dxgi_capture, "Output", return_value=output), \
             patch.object(dxgi_capture, "RecoverableBetterCam", side_effect=[object(), object()]) as make_camera:
            first, second = dxgi_capture.create_camera(), dxgi_capture.create_camera()
        self.assertIsNot(first, second)
        self.assertEqual([call.args[0] for call in make_device.call_args_list], adapters)
        self.assertEqual(make_camera.call_args.kwargs["output_color"], "BGRA")

    def test_no_active_display_has_a_specific_recoverable_error(self):
        with patch.object(dxgi_capture, "enum_dxgi_adapters", return_value=[]), self.assertRaisesRegex(
            dxgi_capture.CaptureUnavailable, "no active desktop display is available"
        ):
            dxgi_capture.create_camera()

    def test_partial_construction_and_repeated_cleanup_are_safe(self):
        camera = dxgi_capture.RecoverableBetterCam.__new__(dxgi_capture.RecoverableBetterCam)
        camera._stagesurf = SimpleNamespace(texture=object(), width=16, height=16)
        camera.release()
        camera.release()
        self.assertIsNone(camera._stagesurf.texture)
        self.assertEqual(camera._stagesurf.width, 0)


if __name__ == "__main__":
    unittest.main()
