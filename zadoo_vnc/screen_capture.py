"""Screen capture thread and encoding controls."""
from __future__ import annotations

import logging
import math
import threading
import time

import bettercam
import mss
import numpy as np
from imagecodecs import JPEG8
from imagecodecs._jpeg8 import jpeg8_encode


class ScreenCapturer(threading.Thread):
    def __init__(self, fps=0, quality=85):
        super().__init__(daemon=True)
        self.latest_frame_jpeg = None
        self.frame_lock = threading.Lock()
        self.is_running = False
        self.last_error = None
        self.quality = quality
        self.fps = fps
        self.bettercam_camera = None
        self._initial_frame_pending = True
        self.active_capture_method = "unknown"
        self.capture_stats = {
            'frame_count': 0,
            'last_fps_time': time.time(),
            'current_fps': 0,
            'sequence': 0,
            'last_frame_ts': 0.0,
            'last_capture_ms': 0.0,
            'last_encode_ms': 0.0,
        }
        # Performance mode settings
        self.perf_enabled = False
        self.perf_region = 'full'
        self.perf_scale_div = 1
        self.perf_grayscale = False
        self._custom_rect_norm = None
        self._frame_event_loop = None
        self._frame_ready_event = None
        self._capture_source_region_applied = False
        self._streaming_event = threading.Event()
        self._streaming_event.set()

    def set_frame_event(self, loop, event):
        self._frame_event_loop = loop
        self._frame_ready_event = event

    def set_streaming_active(self, active: bool):
        if not isinstance(active, bool):
            raise ValueError("Screen streaming state must be a boolean")
        if active:
            if not self._streaming_event.is_set():
                self._initial_frame_pending = True
            self._streaming_event.set()
        else:
            self._streaming_event.clear()
            with self.frame_lock:
                self.capture_stats['current_fps'] = 0.0
                self.capture_stats['frame_count'] = 0
                self.capture_stats['last_fps_time'] = time.time()

    def _notify_frame_ready(self):
        loop = self._frame_event_loop
        event = self._frame_ready_event
        if loop is None or event is None:
            return
        try:
            loop.call_soon_threadsafe(event.set)
        except RuntimeError as exc:
            if self.is_running:
                self.last_error = f"Screen frame notification failed: {exc}"
                logging.error(self.last_error)

    def get_latest_frame_packet(self):
        with self.frame_lock:
            return self.capture_stats['sequence'], self.latest_frame_jpeg

    def run(self):
        self.is_running = True

        next_deadline = time.perf_counter()
        while self.is_running:
            if not self._streaming_event.wait(timeout=0.25):
                next_deadline = time.perf_counter()
                continue
            try:
                capture_start = time.perf_counter()
                frame = self._grab_screen_bettercam()
                source_region_applied = self._capture_source_region_applied
                capture_ms = (time.perf_counter() - capture_start) * 1000.0
                current_time = time.time()
                with self.frame_lock:
                    self.capture_stats['last_capture_ms'] = capture_ms
                    self.capture_stats['frame_count'] += 1
                    time_diff = current_time - self.capture_stats['last_fps_time']
                    if time_diff >= 1.0:
                        self.capture_stats['current_fps'] = self.capture_stats['frame_count'] / time_diff
                        self.capture_stats['frame_count'] = 0
                        self.capture_stats['last_fps_time'] = current_time
                if isinstance(frame, np.ndarray) and not source_region_applied:
                    frame = self._apply_perf_region(frame)
                if frame is not None:
                    encode_start = time.perf_counter()
                    jpeg_bytes = self._encode_frame(frame)
                    encode_ms = (time.perf_counter() - encode_start) * 1000.0
                    if not isinstance(jpeg_bytes, bytes) or not jpeg_bytes:
                        raise RuntimeError("JPEG encoder returned an empty or non-bytes frame")
                    with self.frame_lock:
                        self.latest_frame_jpeg = jpeg_bytes
                        self.capture_stats['sequence'] += 1
                        self.capture_stats['last_frame_ts'] = current_time
                        self.capture_stats['last_encode_ms'] = encode_ms
                    self._notify_frame_ready()
                                
            except Exception as exc:
                self.last_error = f"Screen capture failed: {exc}"
                logging.exception(self.last_error)
                self.is_running = False
                break
            target_fps = int(self.fps)
            if target_fps > 0:
                frame_period = 1.0 / target_fps
                next_deadline += frame_period
                now = time.perf_counter()
                if next_deadline < now - frame_period:
                    next_deadline = now
                sleep_for = max(0.0, next_deadline - now)
                if sleep_for:
                    time.sleep(sleep_for)
            else:
                next_deadline = time.perf_counter()
        
        self._release_bettercam()

    def _grab_screen_bettercam(self):
        """Return a new desktop frame, or None when the desktop is unchanged."""
        if self.bettercam_camera is None:
            # Keep DirectX's native BGRA layout. libjpeg-turbo converts it directly,
            # avoiding BetterCam's OpenCV conversion and an extra full-frame copy.
            self.bettercam_camera = bettercam.create(device_idx=0, output_idx=0, output_color="BGRA")
        if self._initial_frame_pending:
            self._initial_frame_pending = False
            self._capture_source_region_applied = False
            frame = self._capture_initial_frame()
        else:
            region = self._roi_norm_to_pixels(
                int(self.bettercam_camera.width),
                int(self.bettercam_camera.height),
            )
            self._capture_source_region_applied = bool(region)
            frame = self.bettercam_camera.grab(region=region)
            if frame is None:
                return None
        if not isinstance(frame, np.ndarray):
            raise TypeError(f"BetterCam returned {type(frame).__name__}; expected ndarray")
        if frame.ndim != 3 or frame.shape[2] not in {3, 4}:
            raise ValueError(f"BetterCam returned invalid frame shape: {frame.shape}")
        self.active_capture_method = "bettercam"
        return frame

    @staticmethod
    def _capture_initial_frame():
        with mss.mss() as sct:
            shot = sct.grab(sct.monitors[1])
        return np.frombuffer(shot.bgra, dtype=np.uint8).reshape((shot.height, shot.width, 4))

    def _release_bettercam(self):
        cam = self.bettercam_camera
        self.bettercam_camera = None
        if cam is None:
            return
        cam.stop()
        # BetterCam 1.0.0 calls Release() before discarding comtypes pointers,
        # so comtypes releases the same COM references again in __del__.
        # Dropping each smart pointer lets its owner release the reference once.
        cam._duplicator.texture = None
        cam._duplicator.duplicator = None
        cam._stagesurf.texture = None
        cam._stagesurf.width = 0
        cam._stagesurf.height = 0

    def _encode_frame(self, frame):
        if not isinstance(frame, np.ndarray):
            raise TypeError(f"Unsupported frame type for JPEG encoding: {type(frame).__name__}")
        if frame.ndim != 3 or frame.shape[2] not in {3, 4}:
            raise ValueError(f"Unsupported frame shape for JPEG encoding: {frame.shape}")
        arr = frame
        if self.perf_enabled and self.perf_scale_div > 1:
            arr = arr[::self.perf_scale_div, ::self.perf_scale_div, ...]
        if self.perf_grayscale and arr.ndim == 3:
            if arr.shape[2] == 4:  # BGRA
                arr = (0.114 * arr[:, :, 0] + 0.587 * arr[:, :, 1] + 0.299 * arr[:, :, 2]).astype(np.uint8)
            else:  # RGB
                arr = (0.299 * arr[:, :, 0] + 0.587 * arr[:, :, 1] + 0.114 * arr[:, :, 2]).astype(np.uint8)
        if not arr.flags.c_contiguous:
            arr = np.ascontiguousarray(arr)
        if arr.ndim == 2:
            colorspace = outcolorspace = JPEG8.CS.GRAYSCALE
        else:
            colorspace = JPEG8.CS.EXT_BGRA if arr.shape[2] == 4 else JPEG8.CS.RGB
            outcolorspace = JPEG8.CS.YCbCr
        return jpeg8_encode(
            arr,
            level=self.quality,
            colorspace=colorspace,
            outcolorspace=outcolorspace,
        )

    def get_capture_stats(self):
        """Get capture performance statistics"""
        with self.frame_lock:
            stats = dict(self.capture_stats)
            active_method = self.active_capture_method
        is_working = (
            self.is_running
            and active_method != "unknown"
            and stats['sequence'] > 0
            and self.last_error is None
        )
        return {
            'current_fps': stats['current_fps'],
            'active_method': active_method,
            'sequence': stats['sequence'],
            'last_frame_ts': stats['last_frame_ts'],
            'last_capture_ms': stats['last_capture_ms'],
            'last_encode_ms': stats['last_encode_ms'],
            'target_fps': self.fps,
            'target_fps_mode': 'max' if not self.fps else 'fixed',
            'quality': self.quality,
            'perf_enabled': self.perf_enabled,
            'perf_region': self.perf_region,
            'perf_scale_div': self.perf_scale_div,
            'perf_grayscale': self.perf_grayscale,
            'error': self.last_error,
            'is_working': is_working
        }

    def configure_performance(
        self,
        enabled: bool,
        region: str,
        scale_div: int,
        grayscale: bool,
        rect_norm: dict | None = None,
    ):
        if not isinstance(enabled, bool):
            raise ValueError("Performance enabled must be a boolean")
        if not isinstance(region, str):
            raise ValueError("Performance region must be a string")
        if region not in {'full', 'center_0.75', 'center_0.5', 'custom'}:
            raise ValueError(f"Unsupported performance region: {region}")
        if not isinstance(scale_div, int) or isinstance(scale_div, bool):
            raise ValueError("Performance scale_div must be an integer")
        if scale_div not in {1, 2, 3}:
            raise ValueError("Performance scale_div must be 1, 2, or 3")
        if not isinstance(grayscale, bool):
            raise ValueError("Performance grayscale must be a boolean")
        custom_rect = self._custom_rect_norm if rect_norm is None else self._normalize_custom_region(rect_norm)
        if enabled and region == 'custom' and custom_rect is None:
            raise ValueError("Custom performance region coordinates are required")
        self.perf_enabled = enabled
        self.perf_region = region
        self.perf_scale_div = scale_div
        self.perf_grayscale = grayscale
        self._custom_rect_norm = custom_rect if region == 'custom' else None

    def _roi_norm_to_pixels(self, width: int, height: int):
        """Return (l, t, r, b) in pixels from current perf settings; None if not active."""
        if not self.perf_enabled:
            return None

        def center_box(scale: float):
            w = int(width * scale)
            h = int(height * scale)
            x = max(0, (width - w) // 2)
            y = max(0, (height - h) // 2)
            return (x, y, x + w, y + h)

        if self.perf_region == 'center_0.75':
            return center_box(0.75)
        if self.perf_region == 'center_0.5':
            return center_box(0.5)
        if self.perf_region == 'custom' and self._custom_rect_norm:
            x0 = self._custom_rect_norm['x0']
            y0 = self._custom_rect_norm['y0']
            x1 = self._custom_rect_norm['x1']
            y1 = self._custom_rect_norm['y1']
            left = max(0, min(width, int(x0 * width)))
            top = max(0, min(height, int(y0 * height)))
            right = max(0, min(width, int(x1 * width)))
            bottom = max(0, min(height, int(y1 * height)))
            if right > left and bottom > top:
                return left, top, right, bottom
            raise ValueError("Custom performance region is smaller than one capture pixel")
        if self.perf_region == 'full':
            return None
        raise RuntimeError(f"Unsupported active performance region: {self.perf_region}")

    def get_active_region_norm(self):
        """Return the visible capture region as normalized (left, top, right, bottom)."""
        if not self.perf_enabled:
            return None
        region = self.perf_region
        if region == 'center_0.75':
            return 0.125, 0.125, 0.875, 0.875
        if region == 'center_0.5':
            return 0.25, 0.25, 0.75, 0.75
        if region == 'custom' and self._custom_rect_norm:
            rect = self._custom_rect_norm
            return rect['x0'], rect['y0'], rect['x1'], rect['y1']
        if region == 'full':
            return None
        raise RuntimeError(f"Unsupported active performance region: {region}")

    def _apply_perf_region(self, frame: np.ndarray) -> np.ndarray:
        if not self.perf_enabled:
            return frame
        if not isinstance(frame, np.ndarray) or frame.ndim != 3 or frame.shape[2] not in {3, 4}:
            raise ValueError(f"Invalid performance frame shape: {getattr(frame, 'shape', None)}")
        h, w = frame.shape[:2]
        region = self._roi_norm_to_pixels(w, h)
        if not region:
            return frame
        left, top, right, bottom = region
        return frame[top:bottom, left:right, :]

    @staticmethod
    def _normalize_custom_region(rect_norm: dict):
        try:
            rect = {key: float(rect_norm[key]) for key in ('x0', 'y0', 'x1', 'y1')}
        except (KeyError, TypeError, ValueError) as exc:
            raise ValueError(f"Invalid custom performance region: {exc}") from exc
        if any(not math.isfinite(value) or value < 0.0 or value > 1.0 for value in rect.values()):
            raise ValueError("Custom performance region coordinates must be finite values between 0 and 1")
        if rect['x1'] <= rect['x0'] or rect['y1'] <= rect['y0']:
            raise ValueError("Custom performance region must have positive width and height")
        return rect

    def stop(self):
        self.is_running = False
        self._streaming_event.set()
