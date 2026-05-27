"""Screen capture thread and encoding controls."""
from __future__ import annotations

import io
import logging
import os
import threading
import time

from .dependencies import HAS_BETTERCAM, HAS_DXCAM, HAS_IMAGECODECS, HAS_PIL, Image, bettercam, dxcam, imagecodecs, np
from .logging_utils import _log_fallback


class ScreenCapturer(threading.Thread):
    def __init__(self, fps=0, quality=85):
        super().__init__(daemon=True)
        self.latest_frame_jpeg = None
        self.frame_lock = threading.Lock()
        self.capture_control_lock = threading.RLock()
        self.encoder_lock = threading.Lock()
        self.is_running = False
        self.quality = quality
        self.fps = fps
        self.dxcam_camera = None
        self.dxcam_started = False
        self._dxcam_active_region = None
        self._dxcam_active_target_fps = None
        self._dxcam_source_region_applied = False
        self.bettercam_camera = None
        self.bettercam_started = False
        # Lock BetterCam to the known working pair from diagnostics
        self.bettercam_output_idx = 0
        self.bettercam_device_idx = 0
        self.capture_method = self._initial_capture_method()
        self.active_capture_method = "unknown"  # Track which method is actually being used
        self.capture_stats = {
            'frame_count': 0,
            'last_fps_time': time.time(),
            'current_fps': 0,
            'method_switches': 0,
            'last_method_switch': time.time(),
            'sequence': 0,
            'last_frame_ts': 0.0,
            'last_frame_bytes': 0,
            'last_capture_ms': 0.0,
            'last_encode_ms': 0.0,
            'last_loop_ms': 0.0,
        }
        # Performance mode settings
        self.perf_enabled = False
        self.perf_region = 'full'
        self.perf_scale_div = 1
        self.perf_grayscale = False
        self._frame_event_loop = None
        self._frame_ready_event = None
        self.backend_perf = {}
        self._backend_disabled_until = {}
        self._backend_no_frame_counts = {}
        self._backend_no_frame_since = {}
        self._last_capture_source_region_applied = False
        self._last_jpeg_encoder = "unknown"
        self._imagecodecs_jpeg_available = bool(HAS_IMAGECODECS)

    def _initial_capture_method(self):
        configured_method = os.getenv("ZADOO_CAPTURE_METHOD")
        method = self._normalize_capture_method(configured_method)
        if method and self._capture_method_available(method):
            return method
        fallback = self._preferred_capture_method()
        if configured_method:
            logging.warning(
                "Ignoring unavailable ZADOO_CAPTURE_METHOD=%r; using %s",
                configured_method,
                fallback,
            )
        return fallback

    @staticmethod
    def _normalize_capture_method(method):
        method = str(method or "").strip().lower()
        aliases = {
            "better_cam": "bettercam",
            "better-cam": "bettercam",
            "dx": "dxcam",
            "dx_cam": "dxcam",
            "dx-cam": "dxcam",
        }
        return aliases.get(method, method)

    def _installed_capture_methods(self):
        methods = []
        if HAS_BETTERCAM:
            methods.append("bettercam")
        if HAS_DXCAM:
            methods.append("dxcam")
        return methods

    def _preferred_capture_method(self):
        methods = self._installed_capture_methods()
        if methods:
            return methods[0]
        return "none"

    def _capture_method_available(self, method):
        return self._normalize_capture_method(method) in self._installed_capture_methods()

    def set_frame_event(self, loop, event):
        self._frame_event_loop = loop
        self._frame_ready_event = event

    def _notify_frame_ready(self):
        loop = self._frame_event_loop
        event = self._frame_ready_event
        if loop is None or event is None:
            return
        try:
            loop.call_soon_threadsafe(event.set)
        except Exception:
            pass

    def get_latest_frame_packet(self):
        with self.frame_lock:
            return self.capture_stats.get('sequence', 0), self.latest_frame_jpeg

    def run(self):
        self.is_running = True

        next_deadline = time.perf_counter()
        while self.is_running:
            loop_start = time.perf_counter()
            try:
                capture_start = time.perf_counter()
                with self.capture_control_lock:
                    frame = self._grab_screen()
                    source_region_applied = self._last_capture_source_region_applied
                capture_ms = (time.perf_counter() - capture_start) * 1000.0
                self._record_backend_perf(self.active_capture_method, capture_ms, frame is not None)
                # Apply perf-region cropping before encoding (ndarray only)
                if isinstance(frame, np.ndarray) and not source_region_applied:
                    frame = self._apply_perf_region(frame)
                if frame is not None:
                    # Aggressive drop: if encoder is busy, skip this frame.
                    if not self.encoder_lock.acquire(blocking=False):
                        time.sleep(0)  # yield
                        continue
                    try:
                        encode_start = time.perf_counter()
                        jpeg_bytes = self._encode_frame(frame)
                        encode_ms = (time.perf_counter() - encode_start) * 1000.0
                    finally:
                        self.encoder_lock.release()
                    if jpeg_bytes is not None:
                        with self.frame_lock:
                            self.latest_frame_jpeg = jpeg_bytes
                            self.capture_stats['sequence'] += 1
                            self.capture_stats['last_frame_ts'] = time.time()
                            self.capture_stats['last_frame_bytes'] = len(jpeg_bytes)
                            self.capture_stats['last_capture_ms'] = capture_ms
                            self.capture_stats['last_encode_ms'] = encode_ms
                            self.capture_stats['last_loop_ms'] = (time.perf_counter() - loop_start) * 1000.0
                            self.capture_stats['frame_count'] += 1
                            current_time = time.time()
                            time_diff = current_time - self.capture_stats['last_fps_time']

                            # Calculate FPS every second
                            if time_diff >= 1.0:
                                self.capture_stats['current_fps'] = self.capture_stats['frame_count'] / time_diff
                                self.capture_stats['frame_count'] = 0
                                self.capture_stats['last_fps_time'] = current_time
                                log_fps = True
                            else:
                                log_fps = False
                        self._notify_frame_ready()

                        if log_fps and int(current_time) % 5 == 0:
                            # Log performance info every 5 seconds
                            with self.frame_lock:
                                current_fps = self.capture_stats["current_fps"]
                                active_method = self.active_capture_method
                                configured_method = self.capture_method
                            try:
                                logging.debug(
                                    "Capture: %s | FPS: %.1f | Method: %s",
                                    active_method,
                                    current_fps,
                                    configured_method,
                                )
                            except Exception:
                                pass
                                
            except Exception as e:
                logging.error("Exception in ScreenCapturer run loop", exc_info=True)
            try:
                target_fps = int(self.fps)
            except Exception:
                target_fps = 0
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
                    time.sleep(0)
            else:
                next_deadline = time.perf_counter()
                time.sleep(0)
        
        self._release_dxcam()
        self._release_bettercam()

    def _grab_screen(self):
        self._last_capture_source_region_applied = False
        method = self._normalize_capture_method(self.capture_method)
        if not self._capture_method_available(method):
            self.active_capture_method = "unknown"
            logging.error("Screen capture method '%s' is not available.", method)
            return None
        if self._backend_is_disabled(method):
            self.active_capture_method = "unknown"
            return None

        if method == "dxcam":
            self._prepare_backend("dxcam")
            if self._ensure_dxcam_camera() is None:
                return None
            try:
                frame = self._grab_screen_dxcam()
                if frame is not None:
                    self._reset_backend_no_frame("dxcam")
                    self.active_capture_method = "dxcam"
                else:
                    self._record_backend_no_frame("dxcam")
                return frame
            except Exception as e:
                _log_fallback("screen_capture.explicit_dxcam", "no_frame", str(e), e)
                logging.warning("DXCam capture failed", exc_info=True)
                self._record_backend_failure("dxcam")
                return None

        if method == "bettercam":
            self._prepare_backend("bettercam")
            try:
                logging.debug("Attempting BetterCam capture (explicit method)")
                frame = self._grab_screen_bettercam()
                if frame is not None:
                    self.active_capture_method = "bettercam"
                    logging.debug("BetterCam capture succeeded (explicit)")
                else:
                    self._record_backend_no_frame("bettercam")
                return frame
            except Exception as exc:
                _log_fallback("screen_capture.explicit_bettercam", "no_frame", str(exc), exc)
                logging.exception("BetterCam explicit capture threw exception")
                self._record_backend_failure("bettercam")
                return None
        
        logging.error("Screen capture method '%s' is unsupported.", method)
        return None

    def _record_backend_perf(self, method, capture_ms, success):
        if not success or method in (None, "", "unknown"):
            return
        stats = self.backend_perf.setdefault(str(method), {"samples": 0, "avg_ms": 0.0, "failures": 0})
        samples = int(stats.get("samples") or 0)
        avg_ms = float(stats.get("avg_ms") or 0.0)
        stats["samples"] = samples + 1
        stats["avg_ms"] = float(capture_ms) if samples == 0 else (avg_ms * 0.85 + float(capture_ms) * 0.15)
        stats["failures"] = max(0, int(stats.get("failures") or 0) - 1)
        self._reset_backend_no_frame(method)

    def _record_backend_failure(self, method):
        if method in (None, "", "unknown"):
            return
        stats = self.backend_perf.setdefault(str(method), {"samples": 0, "avg_ms": 999.0, "failures": 0})
        stats["failures"] = int(stats.get("failures") or 0) + 1
        if str(method) in {"bettercam", "dxcam"} and int(stats["failures"]) >= 2:
            self._disable_backend(method, seconds=45)

    def _disable_backend(self, method, seconds=30):
        method = str(method or "")
        if not method:
            return
        self._backend_disabled_until[method] = time.time() + max(1, int(seconds))
        if method == "bettercam":
            self._release_bettercam()
        elif method == "dxcam":
            self._release_dxcam()
        if method == self.capture_method:
            self.active_capture_method = "unknown"
        logging.warning("%s capture backend disabled for %ss after repeated failures", method, seconds)

    def _backend_is_disabled(self, method):
        until = float(self._backend_disabled_until.get(str(method), 0.0) or 0.0)
        if until <= 0:
            return False
        if time.time() >= until:
            self._backend_disabled_until.pop(str(method), None)
            return False
        return True

    def _record_backend_no_frame(self, method):
        method = str(method or "")
        if not method:
            return 0, 0.0
        now = time.time()
        stats = self.backend_perf.setdefault(method, {"samples": 0, "avg_ms": 999.0, "failures": 0})
        stats["no_frames"] = int(stats.get("no_frames") or 0) + 1
        count = int(self._backend_no_frame_counts.get(method, 0) or 0) + 1
        self._backend_no_frame_counts[method] = count
        first_seen = float(self._backend_no_frame_since.setdefault(method, now) or now)
        return count, max(0.0, now - first_seen)

    def _reset_backend_no_frame(self, method):
        method = str(method or "")
        if not method:
            return
        self._backend_no_frame_counts.pop(method, None)
        self._backend_no_frame_since.pop(method, None)

    def _prepare_backend(self, method):
        method = str(method or "")
        if method == "dxcam" and self.bettercam_camera is not None:
            logging.info("DXCam: releasing BetterCam before initialization")
            self._release_bettercam()
        elif method == "bettercam" and self.dxcam_camera is not None:
            logging.info("BetterCam: releasing DXCam before initialization")
            self._release_dxcam()

    def _ensure_dxcam_camera(self):
        if not HAS_DXCAM:
            return None
        if self.dxcam_camera is not None:
            return self.dxcam_camera
        try:
            self.dxcam_camera = dxcam.create()
        except Exception:
            self.dxcam_camera = None
        return self.dxcam_camera

    def _dxcam_target_fps(self):
        try:
            return max(0, int(self.fps))
        except Exception:
            return 0

    def _dxcam_full_size(self):
        cam = self.dxcam_camera
        if cam is None:
            return None
        try:
            width = int(getattr(cam, "width", 0) or 0)
            height = int(getattr(cam, "height", 0) or 0)
            if width > 0 and height > 0:
                return width, height
        except Exception:
            pass
        try:
            region = getattr(cam, "region", None)
            if region and len(region) == 4:
                width = int(region[2]) - int(region[0])
                height = int(region[3]) - int(region[1])
                if width > 0 and height > 0:
                    return width, height
        except Exception:
            pass
        return None

    def _dxcam_desired_region(self):
        size = self._dxcam_full_size()
        if not size:
            return None
        width, height = size
        region = self._roi_norm_to_pixels(width, height)
        if not region:
            return None
        return tuple(int(value) for value in region)

    def _ensure_dxcam_started(self):
        cam = self._ensure_dxcam_camera()
        if cam is None:
            return None
        target_fps = self._dxcam_target_fps()
        region = self._dxcam_desired_region()
        needs_restart = (
            not self.dxcam_started
            or region != self._dxcam_active_region
            or target_fps != self._dxcam_active_target_fps
            or not bool(getattr(cam, "is_capturing", False))
        )
        if not needs_restart:
            self._dxcam_source_region_applied = bool(region)
            return cam
        if self.dxcam_started and hasattr(cam, "stop"):
            try:
                cam.stop()
            except Exception:
                logging.debug("DXCam stop before restart failed", exc_info=True)
        logging.info(
            "DXCam: starting ring-buffer capture target_fps=%s region=%s video_mode=True",
            target_fps,
            region or "full",
        )
        try:
            cam.start(region=region, target_fps=target_fps, video_mode=True)
        except TypeError:
            cam.start(region=region, target_fps=target_fps)
        self.dxcam_started = True
        self._dxcam_active_region = region
        self._dxcam_active_target_fps = target_fps
        self._dxcam_source_region_applied = bool(region)
        return cam

    def _grab_screen_dxcam(self):
        """DXCam ring-buffer capture method - returns RGB ndarray."""
        try:
            cam = self._ensure_dxcam_started()
            if cam is None:
                return None
            frame = cam.get_latest_frame(copy=True)
            self._last_capture_source_region_applied = bool(self._dxcam_source_region_applied)
            return frame
        except Exception:
            logging.warning("DXCam ring-buffer capture failed", exc_info=True)
            return None

    def _grab_screen_bettercam(self):
        """Capture using BetterCam library"""
        try:
            # Ensure BetterCam is created and started safely
            if self.bettercam_camera is None:
                # Use the working indices from diagnostics only; do not probe
                try:
                    logging.info(f"BetterCam: creating (device_idx=0, output_idx=0)")
                    self.bettercam_camera = bettercam.create(
                        device_idx=0,
                        output_idx=0,
                        max_buffer_len=4,
                    )
                except Exception:
                    logging.exception("BetterCam create() failed for device_idx=0, output_idx=0")
                    self.bettercam_camera = None
                    return None
                # Start only if start exists and we haven't started yet
                if hasattr(self.bettercam_camera, 'start') and not self.bettercam_started:
                    try:
                        try:
                            tfps = int(self.fps)
                        except Exception:
                            tfps = 0
                        if tfps <= 0:
                            tfps = None
                        logging.info(f"BetterCam: starting capture target_fps={tfps}")
                        if tfps is not None:
                            self.bettercam_camera.start(target_fps=tfps)
                        else:
                            self.bettercam_camera.start()
                        self.bettercam_started = True
                    except Exception as exc:
                        _log_fallback("screen_capture.bettercam_start", "direct_grab", "start_failed", exc)
                        logging.exception("BetterCam start() failed (continuing)")
                        # Some versions auto-start; continue
                        pass
            # Try to obtain a frame with a few attempts (warm-up)
            frame = None
            for _ in range(5):
                # When capturing, prefer the latest captured frame
                if hasattr(self.bettercam_camera, 'get_latest_frame') and self.bettercam_started:
                    frame = self.bettercam_camera.get_latest_frame()
                # If no capture session or no frame yet, try a direct screenshot
                if frame is None and hasattr(self.bettercam_camera, 'grab'):
                    frame = self.bettercam_camera.grab()
                if frame is not None:
                    break
                time.sleep(0.01)
            if frame is None:
                logging.warning("BetterCam: no frame after attempts")
                return None
            # Ensure ndarray RGB
            try:
                if isinstance(frame, np.ndarray):
                    if frame.ndim == 3 and frame.shape[2] == 4:
                        frame = frame[:, :, :3]
                    return frame
                # If some versions return PIL Image, convert to ndarray
                if HAS_PIL and hasattr(frame, 'tobytes'):
                    arr = np.array(frame)
                    if arr.ndim == 3 and arr.shape[2] == 4:
                        arr = arr[:, :, :3]
                    return arr
            except Exception:
                logging.exception("BetterCam: failed to normalize frame to ndarray")
                self._disable_backend("bettercam", seconds=45)
                return None
        except Exception:
            logging.warning("BetterCam capture failed", exc_info=True)
            self._disable_backend("bettercam", seconds=45)
            return None

    def _release_dxcam(self):
        cam = self.dxcam_camera
        self.dxcam_camera = None
        self.dxcam_started = False
        self._dxcam_active_region = None
        self._dxcam_active_target_fps = None
        self._dxcam_source_region_applied = False
        if cam is None:
            return
        try:
            if hasattr(cam, "stop") and bool(getattr(cam, "is_capturing", False)):
                cam.stop()
        except Exception:
            logging.debug("DXCam stop failed", exc_info=True)
        try:
            cam.release()
        except Exception:
            logging.debug("DXCam release failed", exc_info=True)

    def _release_bettercam(self):
        cam = self.bettercam_camera
        self.bettercam_camera = None
        self.bettercam_started = False
        if cam is None:
            return
        try:
            if hasattr(cam, "release"):
                cam.release()
            elif hasattr(cam, "stop"):
                cam.stop()
        except Exception:
            logging.debug("BetterCam release failed", exc_info=True)

    def _encode_frame(self, frame):
        try:
            # Fast path: ndarray -> JPEG (supports RGB or grayscale)
            if isinstance(frame, np.ndarray):
                arr = frame
                if arr.ndim == 3 and arr.shape[2] in (3, 4):
                    if arr.shape[2] == 4:
                        arr = arr[:, :, :3]
                    # Optional grayscale conversion
                    if getattr(self, 'perf_grayscale', False):
                        try:
                            arr = (0.299*arr[:, :, 0] + 0.587*arr[:, :, 1] + 0.114*arr[:, :, 2]).astype(np.uint8)
                        except Exception as exc:
                            _log_fallback("screen_capture.grayscale", "color_frame", "grayscale_conversion_failed", exc)
                            pass
                # Optional integer downscale (decimation)
                if getattr(self, 'perf_enabled', False) and getattr(self, 'perf_scale_div', 1) and self.perf_scale_div > 1:
                    try:
                        if arr.ndim == 3:
                            arr = arr[::self.perf_scale_div, ::self.perf_scale_div, :]
                        else:
                            arr = arr[::self.perf_scale_div, ::self.perf_scale_div]
                    except Exception as exc:
                        _log_fallback("screen_capture.downscale", "undownscaled_frame", "downscale_failed", exc)
                        pass
                if self._imagecodecs_jpeg_available and HAS_IMAGECODECS:
                    try:
                        # Ensure contiguous memory for encoder (avoid implicit copy stalls)
                        if not arr.flags.c_contiguous:
                            arr = np.ascontiguousarray(arr)
                        self._last_jpeg_encoder = "imagecodecs"
                        return imagecodecs.jpeg_encode(arr, level=self.quality)
                    except Exception as exc:
                        _log_fallback("screen_capture.jpeg_encoder", "pillow_jpeg", "imagecodecs_jpeg_failed", exc)
                        logging.warning("imagecodecs jpeg_encode failed  falling back", exc_info=True)
                        self._imagecodecs_jpeg_available = False
                if HAS_PIL:
                    buffer = io.BytesIO()
                    img = Image.fromarray(arr)
                    if getattr(self, 'perf_grayscale', False) and img.mode != 'L':
                        try:
                            img = img.convert('L')
                        except Exception as exc:
                            _log_fallback("screen_capture.pillow_grayscale", "pillow_original_mode", "convert_l_failed", exc)
                            pass
                    img.save(buffer, format='JPEG', quality=self.quality)
                    self._last_jpeg_encoder = "pillow"
                    return buffer.getvalue()
                return None

            # PIL Image
            if HAS_PIL and hasattr(frame, 'save'):
                buffer = io.BytesIO()
                frame.save(buffer, format='JPEG', quality=self.quality)
                self._last_jpeg_encoder = "pillow"
                return buffer.getvalue()
            else:
                logging.error("No JPEG encoder available - both imagecodecs and Pillow failed")
                return None
        except Exception as e:
            logging.error("Failed to encode frame", exc_info=True)
            return None

    def set_capture_method(self, method):
        """Set the screen capture method"""
        method = self._normalize_capture_method(method)
        available_methods = self.get_available_methods()
        if method in available_methods:
            with self.capture_control_lock:
                old_method = self.capture_method
                if method == "dxcam" and self.bettercam_camera is not None:
                    self._release_bettercam()
                if method == "bettercam" and self.dxcam_camera is not None:
                    self._release_dxcam()
                self._backend_disabled_until.pop(method, None)
                self.capture_method = method
                self.capture_stats['method_switches'] += 1
                self.capture_stats['last_method_switch'] = time.time()
                
                # Reset active method to force re-detection
                self.active_capture_method = "unknown"
            
            logging.info(f" Screen capture method changed from '{old_method}' to '{method}'")
            return True
        else:
            logging.warning(f" Capture method '{method}' not available. Available: {available_methods}")
            return False

    def get_available_methods(self):
        """Get list of available capture methods"""
        methods = self._installed_capture_methods()
        
        logging.debug("Checking available capture methods")
        logging.debug("HAS_DXCAM=%s dxcam_camera=%s", HAS_DXCAM, self.dxcam_camera is not None)
        logging.debug("HAS_BETTERCAM=%s bettercam_camera=%s", HAS_BETTERCAM, self.bettercam_camera is not None)
        logging.debug("Available capture methods: %s", methods)
        return methods

    def get_current_method(self):
        """Get current capture method"""
        return self.capture_method

    def get_capture_stats(self):
        """Get capture performance statistics"""
        with self.frame_lock:
            stats = dict(self.capture_stats)
            active_method = self.active_capture_method
            capture_method = self.capture_method
            backend_perf = {k: dict(v) for k, v in self.backend_perf.items()}
            disabled_until = dict(self._backend_disabled_until)
        is_working = active_method != "unknown" and stats['current_fps'] > 0
        return {
            'current_fps': stats['current_fps'],
            'active_method': active_method,
            'set_method': capture_method,
            'method_switches': stats['method_switches'],
            'sequence': stats.get('sequence', 0),
            'last_frame_ts': stats.get('last_frame_ts', 0.0),
            'last_frame_bytes': stats.get('last_frame_bytes', 0),
            'last_capture_ms': stats.get('last_capture_ms', 0.0),
            'last_encode_ms': stats.get('last_encode_ms', 0.0),
            'last_loop_ms': stats.get('last_loop_ms', 0.0),
            'target_fps': self.fps,
            'target_fps_mode': 'max' if not self.fps else 'fixed',
            'quality': self.quality,
            'perf_enabled': self.perf_enabled,
            'perf_region': self.perf_region,
            'perf_scale_div': self.perf_scale_div,
            'perf_grayscale': self.perf_grayscale,
            'jpeg_encoder': self._last_jpeg_encoder,
            'dxcam_ring_buffer': bool(self.dxcam_started),
            'dxcam_target_fps': self._dxcam_active_target_fps,
            'dxcam_region': list(self._dxcam_active_region) if self._dxcam_active_region else None,
            'dxcam_source_region_applied': bool(self._dxcam_source_region_applied),
            'backend_perf': backend_perf,
            'backend_disabled_until': disabled_until,
            'is_working': is_working
        }

    def verify_capture_method(self):
        """Verify that the selected capture method is working"""
        stats = self.get_capture_stats()
        verification = {
            'method_requested': stats['set_method'],
            'method_active': stats['active_method'],
            'is_working': False,
            'fps': stats['current_fps'],
            'status': 'unknown'
        }
        
        if stats['active_method'] == "unknown":
            verification['status'] = 'not_working'
        elif stats['current_fps'] == 0:
            verification['status'] = 'no_frames'
        elif stats['active_method'] == stats['set_method']:
            verification['status'] = 'working_correctly'
            verification['is_working'] = True
        else:
            verification['status'] = 'fallback_working'
            verification['is_working'] = True
            
        return verification
    def set_performance_mode(self, enabled: bool, region: str, scale_div: int):
        self.perf_enabled = bool(enabled)
        self.perf_region = str(region or 'full')
        try:
            v = int(scale_div)
            self.perf_scale_div = v if v >= 1 else 1
        except Exception:
            self.perf_scale_div = 1
        if self.perf_region != 'custom':
            self._custom_rect_norm = None

    def _roi_norm_to_pixels(self, width: int, height: int):
        """Return (l, t, r, b) in pixels from current perf settings; None if not active."""
        if not getattr(self, 'perf_enabled', False):
            return None

        def center_box(scale: float):
            w = int(width * scale); h = int(height * scale)
            x = max(0, (width - w) // 2); y = max(0, (height - h) // 2)
            return (x, y, x + w, y + h)

        try:
            if self.perf_region == 'center_0.75':
                return center_box(0.75)
            if self.perf_region == 'center_0.5':
                return center_box(0.5)
            if self.perf_region == 'custom' and getattr(self, '_custom_rect_norm', None):
                x0 = float(self._custom_rect_norm.get('x0', 0.0))
                y0 = float(self._custom_rect_norm.get('y0', 0.0))
                x1 = float(self._custom_rect_norm.get('x1', 1.0))
                y1 = float(self._custom_rect_norm.get('y1', 1.0))
                l = max(0, min(width,  int(x0 * width)))
                t = max(0, min(height, int(y0 * height)))
                r = max(0, min(width,  int(x1 * width)))
                b = max(0, min(height, int(y1 * height)))
                if r > l and b > t:
                    return (l, t, r, b)
        except Exception:
            pass
        return None

    def get_active_region_norm(self):
        """Return the visible capture region in normalized full-screen coordinates."""
        if not getattr(self, 'perf_enabled', False):
            return None
        region = str(getattr(self, 'perf_region', 'full') or 'full')
        if region == 'center_0.75':
            return {'x0': 0.125, 'y0': 0.125, 'x1': 0.875, 'y1': 0.875}
        if region == 'center_0.5':
            return {'x0': 0.25, 'y0': 0.25, 'x1': 0.75, 'y1': 0.75}
        if region == 'custom' and getattr(self, '_custom_rect_norm', None):
            try:
                x0 = max(0.0, min(1.0, float(self._custom_rect_norm.get('x0', 0.0))))
                y0 = max(0.0, min(1.0, float(self._custom_rect_norm.get('y0', 0.0))))
                x1 = max(0.0, min(1.0, float(self._custom_rect_norm.get('x1', 1.0))))
                y1 = max(0.0, min(1.0, float(self._custom_rect_norm.get('y1', 1.0))))
                left, right = min(x0, x1), max(x0, x1)
                top, bottom = min(y0, y1), max(y0, y1)
                if right > left and bottom > top:
                    return {'x0': left, 'y0': top, 'x1': right, 'y1': bottom}
            except Exception:
                return None
        return None

    def _apply_perf_region(self, frame: np.ndarray) -> np.ndarray:
        if not (self.perf_enabled and isinstance(frame, np.ndarray) and frame.ndim == 3 and frame.shape[2] in (3,4)):
            return frame
        try:
            h, w = frame.shape[:2]
            region = self._roi_norm_to_pixels(w, h)
            if not region:
                return frame
            l, t, r, b = region
            return frame[t:b, l:r, :]
        except Exception:
            return frame

    def set_custom_region(self, rect_norm: dict):
        try:
            self._custom_rect_norm = {
                'x0': float(rect_norm.get('x0', 0)),
                'y0': float(rect_norm.get('y0', 0)),
                'x1': float(rect_norm.get('x1', 1)),
                'y1': float(rect_norm.get('y1', 1)),
            }
            self.perf_region = 'custom'
        except Exception:
            self._custom_rect_norm = None

    def set_grayscale(self, enabled: bool):
        self.perf_grayscale = bool(enabled)

    def stop(self):
        self.is_running = False
        self._release_dxcam()
        self._release_bettercam()
