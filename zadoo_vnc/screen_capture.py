"""Screen capture thread and encoding controls."""
from __future__ import annotations

import asyncio
import io
import logging
import threading
import time

from .dependencies import *

class ScreenCapturer(threading.Thread):
    # Auto-instrument all methods for detailed logging
    def __init__(self, fps=30, quality=85):
        super().__init__(daemon=True)
        self.latest_frame_jpeg = None
        self.frame_lock = threading.Lock()
        self.capture_control_lock = threading.RLock()
        self.is_running = False
        self.quality = quality # Added quality attribute
        self.fps = fps # Added fps attribute
        self.capture_method = "auto"  # auto, dxcam, fast_ctypes, mss, win32, pil
        self.dxcam_camera = None
        self.fast_ctypes_capture = None
        self.bettercam_camera = None
        self.bettercam_started = False
        # Lock BetterCam to the known working pair from diagnostics
        self.bettercam_output_idx = 0
        self.bettercam_device_idx = 0
        self.d3dshot_camera = None
        self.winrt_session = None
        self.active_capture_method = "unknown"  # Track which method is actually being used
        self.capture_stats = {
            'frame_count': 0,
            'last_fps_time': time.time(),
            'current_fps': 0,
            'method_switches': 0,
            'last_method_switch': time.time()
        }
        # Performance mode settings
        self.perf_enabled = False
        self.perf_region = 'full'
        self.perf_scale_div = 1
        self.encoder_busy = False
        self.perf_grayscale = False

    def run(self):
        self.is_running = True
        sct = None
        if HAS_MSS:
            try:
                sct = mss.mss()
            except Exception:
                sct = None
        
        # Initialize DXCam if available
        if HAS_DXCAM:
            try:
                self.dxcam_camera = dxcam.create()
            except Exception:
                self.dxcam_camera = None
        # BetterCam lazy init; created on first use
        # D3DShot disabled for Python 3.12 environment
        # Initialize WinRT GraphicsCapture if available (lazy start later)
        if HAS_WINRT:
            try:
                self.winrt_session = None
            except Exception:
                self.winrt_session = None
        
        # Initialize fast_ctypes_screenshots if available
        if HAS_FAST_CTYPES:
            try:
                self.fast_ctypes_capture = fast_ctypes_screenshots.ScreenshotOfAllMonitors()
                logging.info("fast_ctypes_screenshots capture backend initialized")
            except Exception as e:
                self.fast_ctypes_capture = None
                logging.warning("fast_ctypes_screenshots initialization failed; using fallback capture backends: %s", e)
        else:
            logging.debug("Optional fast_ctypes_screenshots backend is not installed; using fallback capture backends")
        
        while self.is_running:
            try:
                with self.capture_control_lock:
                    frame = self._grab_screen(sct)
                # Apply perf-region cropping before encoding (ndarray only)
                if isinstance(frame, np.ndarray):
                    frame = self._apply_perf_region(frame)
                if frame is not None:
                    # Aggressive drop: if encoder busy, skip this frame
                    if self.encoder_busy:
                        time.sleep(0)  # yield
                        continue
                    self.encoder_busy = True
                    jpeg_bytes = self._encode_frame(frame)
                    self.encoder_busy = False
                    if jpeg_bytes is not None:
                        with self.frame_lock:
                            self.latest_frame_jpeg = jpeg_bytes
                        
                        # Update capture statistics
                        self.capture_stats['frame_count'] += 1
                        current_time = time.time()
                        time_diff = current_time - self.capture_stats['last_fps_time']
                        
                        # Calculate FPS every second
                        if time_diff >= 1.0:
                            self.capture_stats['current_fps'] = self.capture_stats['frame_count'] / time_diff
                            self.capture_stats['frame_count'] = 0
                            self.capture_stats['last_fps_time'] = current_time
                            
                            # Log performance info every 5 seconds
                            if int(current_time) % 5 == 0:
                                logging.debug(
                                    "Capture: %s | FPS: %.1f | Method: %s",
                                    self.active_capture_method,
                                    self.capture_stats["current_fps"],
                                    self.capture_method,
                                )
                                
            except Exception as e:
                logging.error("Exception in ScreenCapturer run loop", exc_info=True)
            try:
                target_fps = max(1, int(self.fps))
            except Exception:
                target_fps = 30
            time.sleep(1 / target_fps)
        
        if sct:
            sct.close()
        if self.dxcam_camera:
            try:
                self.dxcam_camera.release()
            except:
                pass
        if self.bettercam_camera:
            try:
                # Stop BetterCam to avoid __del__ errors on some versions
                if hasattr(self.bettercam_camera, 'stop') and self.bettercam_started:
                    try:
                        self.bettercam_camera.stop()
                    except Exception:
                        pass
                self.bettercam_camera = None
            except:
                pass
        if self.d3dshot_camera:
            try:
                # d3dshot can be cleaned by deleting instance
                self.d3dshot_camera = None
            except:
                pass
        if self.fast_ctypes_capture:
            try:
                # fast_ctypes_screenshots uses context manager, no explicit close needed
                pass
            except:
                pass

    def _grab_screen(self, sct):
        # Use specific method if set, otherwise use auto-detection
        # Helper: ROI mode active?
        roi_mode = bool(getattr(self, "perf_enabled", False) and getattr(self, "perf_region", "full") != "full")
        # Ensure DXCam is available on demand (lazy init)
        if self.capture_method == "dxcam" and HAS_DXCAM:
            if self.dxcam_camera is None:
                try:
                    self.dxcam_camera = dxcam.create()
                except Exception:
                    self.dxcam_camera = None
                    return None
            try:
                frame = self._grab_screen_dxcam()
                if frame is not None:
                    self.active_capture_method = "dxcam"
                return frame
            except Exception as e:
                logging.warning("DXCam capture failed", exc_info=True)
                if self.capture_method != "auto":
                    return None
        
        if self.capture_method == "fast_ctypes" and HAS_FAST_CTYPES and self.fast_ctypes_capture and not roi_mode:
            try:
                frame = self._grab_screen_fast_ctypes()
                if frame is not None:
                    self.active_capture_method = "fast_ctypes"
                return frame
            except Exception as e:
                logging.warning("fast_ctypes capture failed", exc_info=True)
                if self.capture_method != "auto":
                    return None
        
        if self.capture_method == "mss" and sct:
            try:
                # Prefer primary monitor; sct.monitors[0] is "all monitors"
                mon = sct.monitors[1] if len(sct.monitors) > 1 else sct.monitors[0]
                full_w, full_h = int(mon['width']), int(mon['height'])

                # Map perf/custom region to pixels and then to absolute MSS rect
                roi = self._roi_norm_to_pixels(full_w, full_h)
                rect = self._roi_pixels_to_mss_rect(roi, mon) if roi else mon

                sct_img = sct.grab(rect)
                h, w = sct_img.height, sct_img.width
                arr = np.frombuffer(sct_img.bgra, dtype=np.uint8).reshape((h, w, 4))
                frame = np.ascontiguousarray(arr[..., :3][:, :, ::-1])
                self.active_capture_method = "mss"
                return frame
            except mss.exception.ScreenShotError:
                logging.warning("mss.grab failed", exc_info=True)
                if self.capture_method != "auto":
                    return None

        if self.capture_method == "bettercam" and HAS_BETTERCAM:
            try:
                logging.debug("Attempting BetterCam capture (explicit method)")
                frame = self._grab_screen_bettercam()
                if frame is not None:
                    self.active_capture_method = "bettercam"
                    logging.debug("BetterCam capture succeeded (explicit)")
                return frame
            except Exception:
                logging.exception("BetterCam explicit capture threw exception")
                if self.capture_method != "auto":
                    return None

        # D3DShot path disabled

        if self.capture_method == "winrt" and (HAS_WINRT or HAS_PIL):
            try:
                frame = self._grab_screen_winrt()
                if frame is not None:
                    self.active_capture_method = "winrt"
                return frame
            except Exception:
                if self.capture_method != "auto":
                    return None
        
        # Auto mode: try methods in order of performance
        if self.capture_method == "auto":
            # Try DXCam first (fastest)
            if HAS_DXCAM:
                if self.dxcam_camera is None:
                    try:
                        self.dxcam_camera = dxcam.create()
                    except Exception:
                        self.dxcam_camera = None
                try:
                    frame = self._grab_screen_dxcam()
                    if frame is not None:
                        self.active_capture_method = "dxcam"
                        return frame
                except Exception:
                    pass
            
            # Try fast_ctypes only when not in ROI mode
            if HAS_FAST_CTYPES and self.fast_ctypes_capture and not roi_mode:
                try:
                    frame = self._grab_screen_fast_ctypes()
                    if frame is not None:
                        self.active_capture_method = "fast_ctypes"
                        return frame
                except Exception:
                    pass
            
            # Try MSS (region-aware & primary-only by default)
            if sct:
                try:
                    # Prefer primary monitor; sct.monitors[0] is "all monitors"
                    mon = sct.monitors[1] if len(sct.monitors) > 1 else sct.monitors[0]
                    full_w, full_h = int(mon['width']), int(mon['height'])

                    # Map perf/custom region to pixels and then to absolute MSS rect
                    roi = self._roi_norm_to_pixels(full_w, full_h)
                    rect = self._roi_pixels_to_mss_rect(roi, mon) if roi else mon

                    sct_img = sct.grab(rect)
                    h, w = sct_img.height, sct_img.width
                    arr = np.frombuffer(sct_img.bgra, dtype=np.uint8).reshape((h, w, 4))
                    frame = np.ascontiguousarray(arr[..., :3][:, :, ::-1])
                    self.active_capture_method = "mss"
                    return frame
                except mss.exception.ScreenShotError:
                    pass

            # Try BetterCam
            if HAS_BETTERCAM:
                try:
                    logging.debug("Attempting BetterCam capture (auto mode)")
                    frame = self._grab_screen_bettercam()
                    if frame is not None:
                        self.active_capture_method = "bettercam"
                        logging.debug("BetterCam capture succeeded (auto)")
                        return frame
                except Exception:
                    logging.exception("BetterCam auto capture threw exception")
                    pass

            # D3DShot path disabled

            # Try WinRT last
            if HAS_WINRT or HAS_PIL:
                try:
                    frame = self._grab_screen_winrt()
                    if frame is not None:
                        self.active_capture_method = "winrt"
                        return frame
                except Exception:
                    pass
        
        logging.error("All screen capture methods failed.")
        return None

    def _grab_screen_dxcam(self):
        """DXCam capture method - returns RGB ndarray (region-aware)."""
        try:
            roi = None
            # Probe output size once to map normalized perf region → pixels
            if not hasattr(self, '_dx_out_size') or self._dx_out_size is None:
                probe = self.dxcam_camera.grab()
                if isinstance(probe, np.ndarray) and probe.ndim == 3:
                    self._dx_out_size = (probe.shape[1], probe.shape[0])  # (w, h)
                else:
                    self._dx_out_size = None
            if self._dx_out_size:
                w, h = self._dx_out_size
                r = self._roi_norm_to_pixels(w, h)
                if r:
                    roi = (int(r[0]), int(r[1]), int(r[2]), int(r[3]))

            frame = self.dxcam_camera.grab(region=roi) if roi else self.dxcam_camera.grab()
            return frame
        except Exception:
            logging.warning("DXCam grab failed", exc_info=True)
            return None

    def _grab_screen_fast_ctypes(self):
        """Fast capture using fast_ctypes_screenshots library - RGB ndarray"""
        try:
            frame = self.fast_ctypes_capture.screenshot_monitors()
            return frame
        except Exception:
            logging.warning("fast_ctypes capture failed", exc_info=True)
            return None

    def _grab_screen_bettercam(self):
        """Capture using BetterCam library"""
        try:
            # Ensure BetterCam is created and started safely
            if self.bettercam_camera is None:
                # Ensure DXCam is released before initializing BetterCam (avoid device/output conflicts)
                try:
                    if self.dxcam_camera is not None:
                        logging.info("BetterCam: releasing DXCam before initialization")
                        try:
                            self.dxcam_camera.release()
                        except Exception:
                            pass
                        self.dxcam_camera = None
                except Exception:
                    pass
                # Use the working indices from diagnostics only; do not probe
                try:
                    logging.info(f"BetterCam: creating (device_idx=0, output_idx=0)")
                    self.bettercam_camera = bettercam.create(
                        device_idx=0,
                        output_idx=0,
                        max_buffer_len=256,
                    )
                except Exception:
                    logging.exception("BetterCam create() failed for device_idx=0, output_idx=0")
                    self.bettercam_camera = None
                    return None
                # Start only if start exists and we haven't started yet
                if hasattr(self.bettercam_camera, 'start') and not self.bettercam_started:
                    try:
                        # Align capture rate to current FPS target when available
                        tfps = int(self.fps) if hasattr(self, 'fps') else None
                        logging.info(f"BetterCam: starting capture target_fps={tfps}")
                        if tfps is not None:
                            self.bettercam_camera.start(target_fps=tfps)
                        else:
                            self.bettercam_camera.start()
                        self.bettercam_started = True
                    except Exception:
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
                return None
        except Exception:
            logging.warning("BetterCam capture failed", exc_info=True)
            return None

    # D3DShot support removed for Python 3.12

    def _grab_screen_winrt(self):
        """Capture using WinRT GraphicsCapture (basic window capture)."""
        try:
            # Minimal fallback approach: if Pillow's ImageGrab is available on Windows, use it
            if HAS_PIL and hasattr(ImageGrab, 'grab'):
                # Note: ImageGrab uses GDI; this is a placeholder for true WinRT path
                frame = ImageGrab.grab()
                return frame.convert('RGB') if frame else None
            return None
        except Exception:
            logging.warning("WinRT capture failed (using ImageGrab fallback)", exc_info=True)
            return None

    def _release_dxcam(self):
        cam = self.dxcam_camera
        self.dxcam_camera = None
        self._dx_out_size = None
        if cam is None:
            return
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
        global HAS_IMAGECODECS
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
                        except Exception:
                            pass
                # Optional integer downscale (decimation)
                if getattr(self, 'perf_enabled', False) and getattr(self, 'perf_scale_div', 1) and self.perf_scale_div > 1:
                    try:
                        if arr.ndim == 3:
                            arr = arr[::self.perf_scale_div, ::self.perf_scale_div, :]
                        else:
                            arr = arr[::self.perf_scale_div, ::self.perf_scale_div]
                    except Exception:
                        pass
                if HAS_IMAGECODECS:
                    try:
                        # Ensure contiguous memory for encoder (avoid implicit copy stalls)
                        if not arr.flags.c_contiguous:
                            arr = np.ascontiguousarray(arr)
                        return imagecodecs.jpeg_encode(arr, level=self.quality)
                    except Exception:
                        logging.warning("imagecodecs jpeg_encode failed – falling back", exc_info=True)
                        HAS_IMAGECODECS = False
                if HAS_PIL:
                    buffer = io.BytesIO()
                    img = Image.fromarray(arr)
                    if getattr(self, 'perf_grayscale', False) and img.mode != 'L':
                        try:
                            img = img.convert('L')
                        except Exception:
                            pass
                    img.save(buffer, format='JPEG', quality=self.quality)
                    return buffer.getvalue()
                return None

            # PIL Image
            if HAS_PIL and hasattr(frame, 'save'):
                buffer = io.BytesIO()
                frame.save(buffer, format='JPEG', quality=self.quality)
                return buffer.getvalue()
            else:
                logging.error("No JPEG encoder available - both imagecodecs and Pillow failed")
                return None
        except Exception as e:
            logging.error("Failed to encode frame", exc_info=True)
            return None

    def get_frame(self):
        with self.frame_lock:
            return self.latest_frame_jpeg

    def set_capture_method(self, method):
        """Set the screen capture method"""
        available_methods = self.get_available_methods()
        if method in available_methods:
            with self.capture_control_lock:
                old_method = self.capture_method
                if method in ("auto", "dxcam") and self.bettercam_camera is not None:
                    self._release_bettercam()
                if method == "bettercam" and self.dxcam_camera is not None:
                    self._release_dxcam()
                self.capture_method = method
                self.capture_stats['method_switches'] += 1
                self.capture_stats['last_method_switch'] = time.time()
                
                # Reset active method to force re-detection
                self.active_capture_method = "unknown"
            
            logging.info(f"📹 Screen capture method changed from '{old_method}' to '{method}'")
            return True
        else:
            logging.warning(f"❌ Capture method '{method}' not available. Available: {available_methods}")
            return False

    def get_available_methods(self):
        """Get list of available capture methods"""
        methods = ["auto"]
        
        logging.debug("Checking available capture methods")
        logging.debug("HAS_DXCAM=%s dxcam_camera=%s", HAS_DXCAM, self.dxcam_camera is not None)
        logging.debug("HAS_FAST_CTYPES=%s fast_ctypes_capture=%s", HAS_FAST_CTYPES, self.fast_ctypes_capture is not None)
        logging.debug("HAS_MSS=%s", HAS_MSS)
        logging.debug("HAS_BETTERCAM=%s bettercam_camera=%s", HAS_BETTERCAM, self.bettercam_camera is not None)
        logging.debug("HAS_D3DSHOT=%s d3dshot_camera=%s", HAS_D3DSHOT, self.d3dshot_camera is not None)
        
        if HAS_DXCAM:
            methods.append("dxcam")
            logging.debug("Added dxcam capture method")
        if HAS_FAST_CTYPES:
            methods.append("fast_ctypes")
            logging.debug("Added fast_ctypes capture method")
        if HAS_MSS:
            methods.append("mss")
            logging.debug("Added mss capture method")
        if HAS_BETTERCAM:
            methods.append("bettercam")
            logging.debug("Added bettercam capture method")
        if HAS_D3DSHOT:
            methods.append("d3dshot")
            logging.debug("Added d3dshot capture method")
        # Expose WinRT option if either winrt is available or PIL ImageGrab fallback can be used
        if HAS_WINRT or HAS_PIL:
            if "winrt" not in methods:
                methods.append("winrt")
                logging.debug("Added winrt capture method")
        if HAS_D3DSHOT and self.d3dshot_camera:
            methods.append("d3dshot")
            logging.debug("Added d3dshot capture method")
            
        # de-dup and keep a stable order preference
        pref = ["auto", "dxcam", "fast_ctypes", "mss", "bettercam", "d3dshot", "winrt"]
        methods = [m for m in pref if m in dict.fromkeys(methods)]
        logging.debug("Available capture methods: %s", methods)
        return methods

    def get_current_method(self):
        """Get current capture method"""
        return self.capture_method

    def get_active_method(self):
        """Get the method that's actually being used (may differ from set method)"""
        return self.active_capture_method

    def get_capture_stats(self):
        """Get capture performance statistics"""
        is_working = (
            self.active_capture_method != "unknown"
            and self.capture_stats['current_fps'] > 0
            and (self.capture_method == "auto" or self.active_capture_method == self.capture_method)
        )
        return {
            'current_fps': self.capture_stats['current_fps'],
            'active_method': self.active_capture_method,
            'set_method': self.capture_method,
            'method_switches': self.capture_stats['method_switches'],
            'is_working': is_working
        }

    def verify_capture_method(self):
        """Verify that the selected capture method is working"""
        verification = {
            'method_requested': self.capture_method,
            'method_active': self.active_capture_method,
            'is_working': False,
            'fps': self.capture_stats['current_fps'],
            'status': 'unknown'
        }
        
        if self.active_capture_method == "unknown":
            verification['status'] = 'not_working'
        elif self.capture_stats['current_fps'] == 0:
            verification['status'] = 'no_frames'
        elif self.capture_method == "auto":
            verification['status'] = 'auto_selected'
            verification['is_working'] = True
        elif self.active_capture_method == self.capture_method:
            verification['status'] = 'working_correctly'
            verification['is_working'] = True
        else:
            verification['status'] = 'not_working'
            verification['is_working'] = False
            
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

    def _roi_pixels_to_mss_rect(self, roi, mon):
        """Convert (l,t,r,b) relative to a monitor to MSS dict with absolute desktop coords."""
        if not roi: return None
        l, t, r, b = roi
        return {
            'left': int(mon['left'] + l),
            'top':  int(mon['top']  + t),
            'width':  int(max(1, r - l)),
            'height': int(max(1, b - t)),
        }

    def _apply_perf_region(self, frame: np.ndarray) -> np.ndarray:
        if not (self.perf_enabled and isinstance(frame, np.ndarray) and frame.ndim == 3 and frame.shape[2] in (3,4)):
            return frame
        try:
            h, w = frame.shape[:2]
            if self.perf_region == 'center_0.75':
                rh, rw = int(h*0.75), int(w*0.75)
            elif self.perf_region == 'center_0.5':
                rh, rw = int(h*0.5), int(w*0.5)
            elif self.perf_region == 'custom' and getattr(self, '_custom_rect_norm', None):
                x0 = max(0.0, min(1.0, float(self._custom_rect_norm.get('x0', 0))))
                y0 = max(0.0, min(1.0, float(self._custom_rect_norm.get('y0', 0))))
                x1 = max(0.0, min(1.0, float(self._custom_rect_norm.get('x1', 1))))
                y1 = max(0.0, min(1.0, float(self._custom_rect_norm.get('y1', 1))))
                ix0, iy0 = int(x0 * w), int(y0 * h)
                ix1, iy1 = int(x1 * w), int(y1 * h)
                ix0, iy0 = max(0, ix0), max(0, iy0)
                ix1, iy1 = min(w, ix1), min(h, iy1)
                if ix1 > ix0 and iy1 > iy0:
                    return frame[iy0:iy1, ix0:ix1, :]
                return frame
            else:
                return frame
            y0 = max(0, (h - rh)//2)
            x0 = max(0, (w - rw)//2)
            return frame[y0:y0+rh, x0:x0+rw, :]
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
