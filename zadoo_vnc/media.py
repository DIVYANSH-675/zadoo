"""Video, audio, webcam, cursor, and shell stream handlers."""
from __future__ import annotations

import asyncio
import ctypes
import io
import json
import logging
import os
import queue
import subprocess
import sys
import threading
import time
from collections import deque

import websockets

from .camera_discovery import camera_open_candidates, normalize_camera_devices, resolve_camera_selection
from .dependencies import *
from .win32_input import *

class MediaMixin:

    def _put_realtime_frame(self, frame_queue, frame):
        """Offer a realtime media frame without ever blocking the capture thread."""
        try:
            frame_queue.put_nowait(frame)
            return True
        except queue.Full:
            try:
                frame_queue.get_nowait()
            except queue.Empty:
                pass
            try:
                frame_queue.put_nowait(frame)
                return True
            except queue.Full:
                return False

    def _clear_media_queue(self, frame_queue):
        try:
            with frame_queue.mutex:
                frame_queue.queue.clear()
                frame_queue.unfinished_tasks = 0
                frame_queue.all_tasks_done.notify_all()
        except Exception:
            pass

    def _wake_media_queue_readers(self, frame_queue, count=1):
        try:
            wake_count = max(1, int(count or 1))
        except Exception:
            wake_count = 1
        for _ in range(wake_count):
            if not self._put_realtime_frame(frame_queue, None):
                break

    def _apply_stream_profile(self, reason="adaptive"):
        controller = getattr(self, "adaptive_stream", None)
        if controller is None or not getattr(controller, "enabled", False):
            return None
        try:
            profile = controller.apply_to(self)
            controller.last_reason = str(reason or controller.last_reason)
            return profile
        except Exception:
            logging.debug("Adaptive stream profile apply failed", exc_info=True)
            return None

    def _stream_status_payload(self):
        controller = getattr(self, "adaptive_stream", None)
        if controller is not None:
            status = controller.status()
        else:
            status = {
                "enabled": False,
                "transport_mode": "jpeg_ws",
                "requested_transport": "jpeg_ws",
                "fallback_reason": "",
                "webrtc_configured": False,
                "profile_name": "manual",
                "status": "Balanced",
                "last_reason": "manual",
                "encoder_capabilities": getattr(self, "encoder_capabilities", {}),
                "server": {},
                "client": {},
            }
        status["video_clients"] = len(getattr(self, "video_clients", []))
        status["skipped_sends"] = getattr(self, "_video_skipped_sends", 0)
        status["inflight_sends"] = sum(
            1 for task in getattr(self, "_video_send_tasks", {}).values() if not task.done()
        )
        status["effective_quality"] = getattr(self, "current_quality", None)
        status["quality_locked"] = bool(getattr(self, "_quality_locked_by_user", True))
        return status

    async def _send_stream_status(self, websocket):
        try:
            await websocket.send(json.dumps({
                "type": "stream_adaptation",
                "stream": self._stream_status_payload(),
            }))
        except Exception:
            pass

    async def _broadcast_stream_status(self):
        if not getattr(self, "video_clients", None):
            return
        message = json.dumps({
            "type": "stream_adaptation",
            "stream": self._stream_status_payload(),
        })
        for ws in list(self.video_clients):
            try:
                await ws.send(message)
            except Exception:
                self.video_clients.discard(ws)

    def _capture_stats_payload(self):
        if self.screen_capturer:
            stats = self.screen_capturer.get_capture_stats()
        else:
            stats = {
                "is_working": False,
                "current_fps": 0,
                "target_fps": self.current_fps,
                "target_fps_mode": "max" if not self.current_fps else "fixed",
            }
        send_tasks = getattr(self, "_video_send_tasks", {})
        stream = self._stream_status_payload()
        encoder = stream.get("encoder_capabilities") or {}
        stats.update({
            "video_clients": len(self.video_clients),
            "skipped_sends": getattr(self, "_video_skipped_sends", 0),
            "inflight_sends": sum(1 for task in send_tasks.values() if not task.done()),
            "stream_profile": stream.get("profile_name"),
            "stream_status": stream.get("status"),
            "stream_transport": stream.get("transport_mode"),
            "requested_transport": stream.get("requested_transport"),
            "webrtc_configured": stream.get("webrtc_configured"),
            "stream_fallback_reason": stream.get("fallback_reason"),
            "encoder_backend": encoder.get("active_encoder_backend"),
            "preferred_video_encoder": encoder.get("preferred_video_encoder"),
            "client_stream_stats": stream.get("client", {}),
            "server_stream_stats": stream.get("server", {}),
        })
        return stats

    def _start_audio_capture(self, want_samplerate=48000, packet_frames=960, want_channels=2, mono_method="left"):
        """
        Robust SYSTEM-audio capture (loopback) for Windows 11 using python-soundcard.
        - No dependency on sounddevice.WasapiSettings(loopback=...)
        - Jitter buffer to emit exact 20ms frames (960 @ 48k)
        - Auto-rebind when default output changes (speakers/headphones)
        """
        if getattr(self, "_audio_running", False):
            return
        self._clear_media_queue(self.audio_queue)
        self._audio_running = True
        self.audio_running = True
        self._audio_backend = "soundcard_loopback"
        self._audio_thread = None
        self._audio_stream = None
        self._audio_dev_name = None
        self._audio_sr = want_samplerate

        log = logging.getLogger("sysaudio")
        ring = deque()
        ring_len = 0

        def _to_safe_mono(x: np.ndarray, method="left"):
            if x.ndim == 1 or x.shape[1] == 1:
                return x.reshape(-1, 1)
            if method == "left":
                return x[:, :1]
            if method == "right":
                return x[:, 1:2]
            return ((x[:, :1] + x[:, 1:2]) * 0.5).astype(np.float32)

        def _pick_loopback():
            spk = sc.default_speaker()
            loopback = None
            try:
                for mic in sc.all_microphones(include_loopback=True):
                    if getattr(mic, "isloopback", False) and (spk.name.split(" (")[0] in mic.name or getattr(spk, 'id', None) in getattr(mic, 'id', '')):
                        loopback = mic
                        break
                if loopback is None:
                    for mic in sc.all_microphones(include_loopback=True):
                        if getattr(mic, "isloopback", False):
                            loopback = mic
                            break
            except Exception:
                loopback = None
            return spk, loopback

        stop_evt = threading.Event()
        self._audio_stop_evt = stop_evt

        def _audio_worker():
            nonlocal ring_len
            current_loop = None
            last_pick = 0.0
            try:
                while self._audio_running and not stop_evt.is_set():
                    now = time.time()
                    if current_loop is None or (now - last_pick) > 1.5:
                        last_pick = now
                        spk, loop = _pick_loopback()
                        if loop is not None and loop != current_loop:
                            try:
                                if self._audio_stream is not None:
                                    self._audio_stream.__exit__(None, None, None)
                            except Exception:
                                pass
                            try:
                                self._audio_stream = loop.recorder(samplerate=want_samplerate, channels=2)
                                self._audio_stream.__enter__()
                                current_loop = loop
                                log.info(f" Loopback device: {loop.name} @ {want_samplerate} Hz")
                            except Exception as e:
                                log.warning(f"Loopback open failed ({e}); retrying...")
                                current_loop = None
                                time.sleep(0.2)
                                continue

                    if self._audio_stream is None:
                        time.sleep(0.01)
                        continue

                    try:
                        data = self._audio_stream.record(numframes=packet_frames)
                        # ensure 2-D float32
                        if data.ndim == 1:
                            data = data.reshape(-1, 1)
                        if want_channels == 1:
                            data = _to_safe_mono(data, mono_method)
                        else:
                            if data.shape[1] == 1 and want_channels == 2:
                                data = np.repeat(data, 2, axis=1)
                        ring.append(data)
                        ring_len += data.shape[0]

                        while ring_len >= packet_frames:
                            need = packet_frames
                            chunks = []
                            while need > 0:
                                block = ring[0]
                                if block.shape[0] <= need:
                                    chunks.append(block)
                                    ring.popleft()
                                    ring_len -= block.shape[0]
                                    need -= block.shape[0]
                                else:
                                    chunks.append(block[:need])
                                    ring[0] = block[need:]
                                    ring_len -= need
                                    need = 0
                            out = np.vstack(chunks).astype(np.float32, copy=False)
                            # Convert to s16le mono (if needed) and enqueue for /audio
                            try:
                                mono = out[:, 0] if out.ndim == 2 else out.reshape(-1)
                                pcm = (np.clip(mono, -1.0, 1.0) * 32767.0).astype(np.int16).tobytes()
                                self._put_realtime_frame(self.audio_queue, pcm)
                            except Exception as e:
                                log.debug(f"sender err: {e}")
                    except Exception as e:
                        try:
                            if self._audio_stream is not None:
                                self._audio_stream.__exit__(None, None, None)
                        except Exception:
                            pass
                        self._audio_stream = None
                        current_loop = None
                        time.sleep(0.05)
                        continue
            finally:
                try:
                    if self._audio_stream is not None:
                        self._audio_stream.__exit__(None, None, None)
                except Exception:
                    pass
                self._audio_stream = None
                self._audio_running = False
                self.audio_running = False
                if getattr(self, "_audio_stop_evt", None) is stop_evt:
                    self._audio_stop_evt = None
                log.info(" System-audio worker stopped")

        t = threading.Thread(target=_audio_worker, daemon=True)
        t.start()
        self._audio_thread = t

    def _stop_audio_capture(self):
        self._audio_running = False
        self.audio_running = False
        try:
            stop_evt = getattr(self, "_audio_stop_evt", None)
            if stop_evt is not None:
                stop_evt.set()
        except Exception:
            pass
        try:
            if getattr(self, '_audio_stream', None):
                if hasattr(self._audio_stream, 'stop'):
                    self._audio_stream.stop()
                if hasattr(self._audio_stream, 'close'):
                    self._audio_stream.close()
                if hasattr(self._audio_stream, '__exit__'):
                    try:
                        self._audio_stream.__exit__(None, None, None)
                    except Exception:
                        pass
        except Exception:
            pass
        try:
            if getattr(self, '_audio_thread', None) and self._audio_thread.is_alive():
                self._audio_thread.join(timeout=0.5)
        except Exception:
            pass
        self._audio_stream = None
        self._audio_thread = None
        self._clear_media_queue(self.audio_queue)
        self._wake_media_queue_readers(self.audio_queue, len(getattr(self, "audio_clients", [])) or 1)
        logging.info(" Audio capture stopped")

    def _start_mic_capture(self, samplerate=48000, blocksize=960, channels=1):
        """
        Microphone of System A using sounddevice InputStream (no WASAPI special args).
        Always sends mono s16le frames of 'blocksize' samples at 'samplerate'.
        """
        mic_log = logging.getLogger("mic")
        if self.mic_running:
            mic_log.info("Mic capture already running")
            return True
        if sd is None:
            mic_log.error("Failed to start microphone capture: sounddevice is not installed")
            return False
        mic_log.info("Starting mic capture request sr=%s block=%s channels=%s", samplerate, blocksize, channels)
        self._clear_media_queue(self.mic_queue)
        self.mic_running = True
        self.mic_samplerate = int(samplerate)
        self.mic_blocksize = int(blocksize)
        self.mic_channels = 1

        def mic_callback(indata, frames, time_info, status):
            try:
                if not self.mic_running or not self.mic_clients:
                    return
                if status:
                    logging.debug(f"Mic status: {status}")
                x = indata
                if x.ndim == 2:
                    if x.shape[1] > 1:
                        x = x[:, 0]
                    else:
                        x = x.reshape(-1)
                x = (np.clip(x, -1.0, 1.0) * 32767.0).astype(np.int16)
                self._put_realtime_frame(self.mic_queue, x.tobytes())
            except Exception:
                logging.error("Mic callback error", exc_info=True)

        default_samplerate = self.mic_samplerate
        max_input_channels = 1
        try:
            default_input = sd.query_devices(kind='input')
            max_input_channels = max(1, int(default_input.get('max_input_channels') or 1))
            if default_input.get('default_samplerate'):
                default_samplerate = int(default_input['default_samplerate'])
            mic_log.info(
                "Default mic device name=%s max_input_channels=%s default_samplerate=%s",
                default_input.get('name'),
                max_input_channels,
                default_samplerate,
            )
        except Exception as exc:
            mic_log.warning("Unable to query default microphone device; using requested mic format: %s", exc)

        requested_channels = max(1, int(channels or 1))
        channel_candidates = [min(requested_channels, max_input_channels)]
        if max_input_channels >= 2 and 2 not in channel_candidates:
            channel_candidates.append(2)
        samplerate_candidates = [self.mic_samplerate]
        if default_samplerate and default_samplerate not in samplerate_candidates:
            samplerate_candidates.append(default_samplerate)

        attempts = []
        for sr in samplerate_candidates:
            for ch in channel_candidates:
                pair = (int(sr), int(ch))
                if pair not in attempts:
                    attempts.append(pair)

        last_error = None
        try:
            for sr, ch in attempts:
                try:
                    mic_log.info("Opening mic stream sr=%s input_channels=%s block=%s", sr, ch, self.mic_blocksize)
                    self.mic_stream = sd.InputStream(
                        samplerate=sr,
                        channels=ch,
                        dtype='float32',
                        blocksize=self.mic_blocksize,
                        callback=mic_callback,
                        device=None,
                        latency='low'
                    )
                    self.mic_stream.start()
                    self.mic_samplerate = sr
                    message = f"Mic capture started @ {self.mic_samplerate} Hz input_channels={ch} block={self.mic_blocksize}"
                    mic_log.info(message)
                    print(message)
                    return True
                except Exception as exc:
                    last_error = exc
                    self.mic_stream = None
                    mic_log.warning("Mic open attempt failed sr=%s channels=%s: %s", sr, ch, exc)
        finally:
            if self.mic_stream is None:
                self.mic_running = False

        mic_log.error("Failed to start microphone capture after trying %s: %s", attempts, last_error)
        return False

    def _stop_mic_capture(self):
        mic_log = logging.getLogger("mic")
        try:
            self.mic_running = False
            if self.mic_stream:
                try:
                    self.mic_stream.stop()
                except Exception:
                    pass
                try:
                    self.mic_stream.close()
                except Exception:
                    pass
            self.mic_stream = None
            self._clear_media_queue(self.mic_queue)
            self._wake_media_queue_readers(self.mic_queue, len(getattr(self, "mic_clients", [])) or 1)
            mic_log.info("Mic capture stopped and queue cleared")
        except Exception:
            mic_log.error("Failed to stop mic capture", exc_info=True)

    async def audio_stream_handler(self, websocket: websockets.WebSocketServerProtocol):
        print(f" New audio client connected from {websocket.remote_address}")
        self.audio_clients.add(websocket)

        # Start capture on first client
        if not getattr(self, "_audio_running", False):
            self._start_audio_capture()

        try:
            # Simple header to let client know PCM format
            header = json.dumps({
                'type': 'audio_format',
                'codec': 'pcm_s16le',
                'samplerate': int(getattr(self, "_audio_sr", 48000) or 48000),
                'channels': 1,
                'blocksize': 960
            }).encode()
            await websocket.send(header)

            loop = asyncio.get_running_loop()
            while True:
                try:
                    chunk = await loop.run_in_executor(None, self.audio_queue.get, True, 0.25)
                except queue.Empty:
                    continue
                except Exception:
                    chunk = None
                if chunk is None:
                    if not getattr(self, "_audio_running", False):
                        break
                    await asyncio.sleep(0.005)
                    continue
                try:
                    await websocket.send(chunk)
                except websockets.exceptions.ConnectionClosed:
                    break
        except Exception as e:
            print(f"Error in audio_stream_handler: {e}")
        finally:
            self.audio_clients.discard(websocket)
            print(f" Removed audio client, {len(self.audio_clients)} clients remaining")
            if not self.audio_clients and getattr(self, "_audio_running", False):
                self._stop_audio_capture()
                print(" Audio capture stopped (no clients)")

    def _remove_mic_client(self, websocket):
        try:
            self.mic_clients.discard(websocket)
        except Exception:
            self.mic_clients = set()
        if not self.mic_clients:
            self._stop_mic_capture()
            return True
        return False

    async def mic_stream_handler(self, websocket: websockets.WebSocketServerProtocol):
        mic_log = logging.getLogger("mic")
        remote = getattr(websocket, 'remote_address', None)
        mic_log.info("New mic client connected from %s", remote)
        print(f"New mic client connected from {remote}")
        self.mic_clients.add(websocket)
        try:
            if not self.mic_running:
                ok = self._start_mic_capture(samplerate=48000, blocksize=960, channels=1)
                if not ok:
                    mic_log.error("Mic open failed for %s", remote)
                    print("Mic open failed")
                    try:
                        await websocket.close(code=1011, reason="Mic open failed")
                    except Exception:
                        pass
                    return

            # Send format header after opening so the samplerate reflects the actual stream.
            hdr = {
                "type": "audio_format",
                "samplerate": int(self.mic_samplerate),
                "channels": 1,
                "samplefmt": "s16le",
                "blocksize": int(self.mic_blocksize)
            }
            await websocket.send(json.dumps(hdr).encode('utf-8'))
            mic_log.info("Mic header sent to %s: sr=%s ch=%s fmt=%s block=%s", remote, hdr['samplerate'], hdr['channels'], hdr['samplefmt'], hdr['blocksize'])
            print(f"Mic header sent: sr={hdr['samplerate']} ch={hdr['channels']} fmt={hdr['samplefmt']} block={hdr['blocksize']}")

            sent_chunks = 0
            loop = asyncio.get_running_loop()
            while True:
                try:
                    chunk = await loop.run_in_executor(None, self.mic_queue.get, True, 0.25)
                except queue.Empty:
                    continue
                except Exception:
                    chunk = None
                if chunk is None:
                    if not self.mic_running:
                        break
                    await asyncio.sleep(0.005)
                    continue
                try:
                    await websocket.send(chunk)
                    sent_chunks += 1
                    if sent_chunks == 1 or sent_chunks % 250 == 0:
                        mic_log.info("Mic sent %s chunk(s) to %s; last_chunk_bytes=%s", sent_chunks, remote, len(chunk))
                except websockets.exceptions.ConnectionClosed:
                    break
        except Exception as e:
            mic_log.error("Error in mic_stream_handler for %s: %s", remote, e, exc_info=True)
            print(f"Error in mic_stream_handler: {e}")
        finally:
            stopped = self._remove_mic_client(websocket)
            mic_log.info("Mic client disconnected from %s; remaining=%s stopped=%s", remote, len(self.mic_clients), stopped)
            print(f"Mic client disconnected; remaining={len(self.mic_clients)} stopped={stopped}")

    async def video_stream_handler(self, websocket):
        print(f" New video client connected from {websocket.remote_address}")
        self.video_clients.add(websocket)
        
        try:
            self._apply_stream_profile("client_connected")
            await self._send_stream_status(websocket)
            # Keep connection alive and handle any incoming messages
            async for message in websocket:
                try:
                    # Handle any control messages for video stream
                    event = json.loads(message)
                    action = event.get('action')
                    if not self._is_ws_action_authorized(websocket, action):
                        await self._send_ws_forbidden(websocket, action)
                        continue
                     
                    if action == 'refresh_tunnel':
                        await self.handle_refresh_via_websocket(websocket)
                    elif action == 'get_public_url':
                        await self.handle_get_url_via_websocket(websocket)
                    elif action == 'set_quality':
                        value = self._apply_quality(event.get('value', 75))
                        print(f" Quality set to: {value}%")
                    elif action == 'set_fps':
                        value = self._apply_fps(event.get('value', 30))
                        print(f" FPS set to: {value}")
                    elif action == 'client_stream_stats':
                        controller = getattr(self, "adaptive_stream", None)
                        if controller is not None:
                            controller.record_client_stats(event)
                    elif action == 'stream_ping':
                        await websocket.send(json.dumps({
                            'type': 'stream_pong',
                            'client_ts': event.get('client_ts'),
                            'server_ts': time.time() * 1000.0,
                        }))
                    elif action == 'get_stream_status':
                        await self._send_stream_status(websocket)
                    elif action == 'set_capture_method':
                        method = event.get('method', 'auto')
                        if self.screen_capturer:
                            success = self.screen_capturer.set_capture_method(method)
                            if success:
                                print(f" Capture method changed to: {method}")
                            else:
                                print(f" Failed to set capture method to: {method}")
                        else:
                            print(" Screen capturer not available")
                    elif action == 'get_available_capture_methods':
                        if self.screen_capturer:
                            methods = self.screen_capturer.get_available_methods()
                            current = self.screen_capturer.get_current_method()
                            await websocket.send(json.dumps({
                                'type': 'available_capture_methods',
                                'methods': methods,
                                'current': current
                            }))
                        else:
                            await websocket.send(json.dumps({
                                'type': 'available_capture_methods',
                                'methods': ['auto'],
                                'current': 'auto'
                            }))
                    elif action == 'set_performance':
                        enabled = bool(event.get('enabled'))
                        region = str(event.get('region') or 'full')
                        scale_div = int(event.get('scale_div') or 1)
                        rect_norm = event.get('rect_norm')
                        grayscale = bool(event.get('grayscale'))
                        if self.screen_capturer:
                            self.screen_capturer.set_performance_mode(enabled, region, scale_div)
                            if rect_norm and isinstance(rect_norm, dict):
                                self.screen_capturer.set_custom_region(rect_norm)
                            self.screen_capturer.set_grayscale(grayscale)
                            msg = f"Performance mode: enabled={enabled} region={region} scale_div={scale_div} gray={grayscale} custom={bool(rect_norm)}"
                            logging.getLogger("performance").info(msg)
                            print(f" {msg}")
                    elif action == 'get_capture_stats':
                        await websocket.send(json.dumps({
                            'type': 'capture_stats',
                            'stats': self._capture_stats_payload()
                        }))
                    elif action == 'verify_capture_method':
                        if self.screen_capturer:
                            verification = self.screen_capturer.verify_capture_method()
                            await websocket.send(json.dumps({
                                'type': 'capture_verification',
                                'verification': verification
                            }))
                        else:
                            await websocket.send(json.dumps({
                                'type': 'capture_verification',
                                'verification': {'is_working': False, 'status': 'no_capturer'}
                            }))
                    elif action == 'toggle_keystroke_capture':
                        enabled = event.get('enabled')
                        if enabled is None:
                            enabled = not getattr(self, 'keystroke_capture_enabled', False)
                        if enabled:
                            self.enable_keystroke_capture()
                        else:
                            self.disable_keystroke_capture()
                        await websocket.send(json.dumps({
                            'type': 'keystroke_capture_status',
                            'enabled': self.keystroke_capture_enabled
                        }))
                    else:
                        # Handle other input events
                        self.process_event(event, websocket)
                except Exception as e:
                    print(f"Error handling video client message: {e}")

        except websockets.exceptions.ConnectionClosed:
            print(f" Video client {websocket.remote_address} disconnected")
        except Exception as e:
            print(f"Error in video_stream_handler: {e}")
        finally:
            self.video_clients.discard(websocket)
            task = getattr(self, "_video_send_tasks", {}).pop(websocket, None)
            if task and not task.done():
                task.cancel()
            print(f" Removed video client, {len(self.video_clients)} clients remaining")

    async def webcam_stream_handler(self, websocket):
        """Stream JPEG frames from the server's webcam (DirectShow on Windows) to the client."""
        print(f" Webcam client connected from {websocket.remote_address}")
        # Prefer PyAV (FFmpeg) on Windows via DirectShow
        if not HAS_AV:
            print(" PyAV not available; webcam streaming disabled")
            try:
                await websocket.send(b"")
            except Exception:
                pass
            return

        container = None
        player = None
        cv2_capture = None
        cv2_first_frame = None
        try:
            # First message may include selected device from client
            try:
                first = await asyncio.wait_for(websocket.recv(), timeout=0.2)
                try:
                    data = json.loads(first) if isinstance(first, str) else json.loads(first.decode('utf-8','ignore'))
                    if isinstance(data, dict) and data.get('type') == 'select_camera':
                        selected_device = data.get('device') or data.get('open_name')
                        selected_id = data.get('camera_id') or data.get('id')
                        self.selected_camera = str(selected_device).strip() if selected_device else None
                        self.selected_camera_id = str(selected_id).strip() if selected_id else None
                        if self.selected_camera_id or self.selected_camera:
                            print(f" Client selected camera id={self.selected_camera_id or ''} device={self.selected_camera or ''}")
                except Exception:
                    pass
            except asyncio.TimeoutError:
                pass
            open_ok = False
            candidates = []
            option_sets = [None]
            selected_cv2_only = False
            is_windows = sys.platform.startswith('win')
            loop = asyncio.get_running_loop()

            async def open_dshow_device(device, opts=None):
                def _open():
                    if opts is None:
                        return av.open(device, format='dshow')
                    return av.open(device, format='dshow', options=opts)
                return await loop.run_in_executor(None, _open)

            # Try common device names for Windows DirectShow
            if is_windows:
                # Try multiple option combinations for broader compatibility
                option_sets = [
                    None,
                    { 'video_size': '640x480' },
                    { 'framerate': '15', 'video_size': '640x480' },
                    { 'framerate': '30', 'video_size': '640x480' },
                    { 'framerate': '30', 'video_size': '1280x720' },
                    { 'framerate': '30', 'video_size': '1920x1080' },
                ]
                def add_candidate(value):
                    value = str(value or '').strip()
                    if value and value not in candidates:
                        candidates.append(value)

                cv2_indices = []

                def add_cv2_index(value):
                    try:
                        index = int(value)
                    except Exception:
                        return
                    if index >= 0 and index not in cv2_indices:
                        cv2_indices.append(index)

                def add_device_candidates(device):
                    if not device:
                        return
                    if device.get('device_path'):
                        for value in camera_open_candidates(device):
                            add_candidate(value)
                    add_cv2_index(device.get('index'))

                devices = []
                try:
                    devices = await loop.run_in_executor(None, lambda: normalize_camera_devices(self.enumerate_cameras()))
                except Exception as e:
                    print(f"  Camera enumeration failed while opening webcam: {e}")

                selected_device = resolve_camera_selection(
                    getattr(self, 'selected_camera_id', None),
                    getattr(self, 'selected_camera', None),
                    devices,
                )
                selected_cv2_only = bool(
                    selected_device
                    and not selected_device.get('device_path')
                    and selected_device.get('index') is not None
                )
                add_device_candidates(selected_device)
                if selected_device is None:
                    add_candidate(getattr(self, 'selected_camera', None))
                if not selected_cv2_only:
                    add_candidate(os.environ.get('CAMERA_NAME'))
                if not selected_cv2_only:
                    for device in devices:
                        add_device_candidates(device)
                if not candidates and not selected_cv2_only:
                    candidates.extend([
                        'Iriun Webcam',
                        'HP True Vision FHD Camera',
                        'Integrated Camera',
                        'USB Video Device',
                        'HD Webcam',
                        'HP Wide Vision FHD Camera',
                        'Logitech',
                        'OBS Virtual Camera'
                    ])
                # Try PyAV DirectShow device strings
                for name in candidates:
                    for opts in option_sets:
                        try:
                            device = f"video={name}"
                            if opts is None:
                                print(f" Trying dshow open: {device} opts=None")
                            else:
                                print(f" Trying dshow open: {device} opts={opts}")
                            container = await open_dshow_device(device, opts)
                            open_ok = True
                            print(f" Opened webcam via dshow device: {device} opts={opts}")
                            break
                        except Exception as e:
                            print(f"  Open failed: {device} opts={opts} err={e}")
                            container = None
                            continue
                    if open_ok:
                        break
                # Fallback attempts with generic device strings
                if not open_ok and not selected_cv2_only:
                    for generic in ("video=0", "0", "video=1", "1"):
                        for opts in option_sets:
                            try:
                                if opts is None:
                                    print(f" Trying dshow open: {generic} opts=None")
                                else:
                                    print(f" Trying dshow open: {generic} opts={opts}")
                                container = await open_dshow_device(generic, opts)
                                open_ok = True
                                print(f" Opened webcam via dshow generic: {generic} opts={opts}")
                                break
                            except Exception as e:
                                print(f"  Open failed: {generic} opts={opts} err={e}")
                                container = None
                                continue
                        if open_ok:
                            break
                # Non-Windows simple attempt (may not be used in this environment)
                

            if not open_ok and HAS_AIORTC and not selected_cv2_only:
                # Fallback to aiortc MediaPlayer with dshow
                try:
                    from aiortc.contrib.media import MediaPlayer
                    for name in candidates:
                        for opts in option_sets:
                            try:
                                if opts is None:
                                    print(f" Trying MediaPlayer open: video={name} opts=None")
                                    player = MediaPlayer(f"video={name}", format='dshow')
                                else:
                                    print(f" Trying MediaPlayer open: video={name} opts={opts}")
                                    player = MediaPlayer(f"video={name}", format='dshow', options=opts)
                                print(f" Opened webcam via aiortc MediaPlayer: video={name} opts={opts}")
                                break
                            except Exception as e:
                                print(f"  MediaPlayer open failed: video={name} opts={opts} err={e}")
                                player = None
                                continue
                        if player is not None:
                            break
                    if player is None:
                        for generic in ("video=0", "video=1"):
                            for opts in option_sets:
                                try:
                                    if opts is None:
                                        print(f" Trying MediaPlayer open: {generic} opts=None")
                                        player = MediaPlayer(generic, format='dshow')
                                    else:
                                        print(f" Trying MediaPlayer open: {generic} opts={opts}")
                                        player = MediaPlayer(generic, format='dshow', options=opts)
                                    print(f" Opened webcam via aiortc MediaPlayer: {generic} opts={opts}")
                                    break
                                except Exception as e:
                                    print(f"  MediaPlayer open failed: {generic} opts={opts} err={e}")
                                    player = None
                                    continue
                            if player is not None:
                                break
                except Exception as e:
                    player = None

            if container is None and player is None and is_windows:
                try:
                    import cv2
                except Exception as e:
                    cv2 = None
                    print(f"  OpenCV not available for webcam fallback: {e}")
                if cv2 is not None:
                    for index in cv2_indices:
                        try:
                            cap = cv2.VideoCapture(index, cv2.CAP_DSHOW)
                            if not cap or not cap.isOpened():
                                try:
                                    cap.release()
                                except Exception:
                                    pass
                                print(f"  OpenCV webcam open failed: index={index}")
                                continue
                            cap.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
                            cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)
                            cap.set(cv2.CAP_PROP_FPS, 30)
                            ok, frame = cap.read()
                            if not ok or frame is None:
                                cap.release()
                                print(f"  OpenCV webcam produced no frame: index={index}")
                                continue
                            cv2_capture = cap
                            cv2_first_frame = frame
                            print(f" Opened webcam via OpenCV DirectShow index={index}")
                            break
                        except Exception as e:
                            print(f"  OpenCV webcam open failed: index={index} err={e}")
                            try:
                                cap.release()
                            except Exception:
                                pass
                            cv2_capture = None

            if container is None and player is None and cv2_capture is None:
                print(" No webcam device could be opened")
                try:
                    await websocket.send(b"")
                except Exception:
                    pass
                return

            # Send frames
            if container is not None:
                stream = None
                try:
                    stream = next((s for s in container.streams if s.type == 'video'), None)
                except Exception:
                    stream = None
                if stream is not None:
                    try:
                        stream.thread_type = 'AUTO'
                    except Exception:
                        pass
                    # Decode using stream index; passing a stream object causes errors
                    try:
                        video_stream_index = 0 if stream is None else int(getattr(stream, 'index', 0))
                    except Exception:
                        video_stream_index = 0
                    for frame in container.decode(video=video_stream_index):
                        try:
                            img = frame.to_image()
                            buf = io.BytesIO()
                            img.save(buf, format='JPEG', quality=70)
                            await websocket.send(buf.getvalue())
                        except websockets.exceptions.ConnectionClosed:
                            break
                        except Exception:
                            pass
                        await asyncio.sleep(0.05)
            elif cv2_capture is not None:
                try:
                    import cv2
                    pending_frame = cv2_first_frame
                    while True:
                        try:
                            if pending_frame is None:
                                ok, frame = cv2_capture.read()
                            else:
                                ok, frame = True, pending_frame
                                pending_frame = None
                            if not ok or frame is None:
                                await asyncio.sleep(0.05)
                                continue
                            ok, encoded = cv2.imencode('.jpg', frame, [int(cv2.IMWRITE_JPEG_QUALITY), 70])
                            if ok:
                                await websocket.send(encoded.tobytes())
                        except websockets.exceptions.ConnectionClosed:
                            break
                        except Exception:
                            await asyncio.sleep(0.05)
                            continue
                        await asyncio.sleep(0.05)
                except Exception as e:
                    print(f"Error in OpenCV webcam stream: {e}")
            else:
                # aiortc MediaPlayer path
                video_track = getattr(player, 'video', None)
                if video_track is None:
                    print(" MediaPlayer has no video track")
                    return
                while True:
                    try:
                        frame = await video_track.recv()
                        img = frame.to_image()
                        buf = io.BytesIO()
                        img.save(buf, format='JPEG', quality=70)
                        await websocket.send(buf.getvalue())
                    except websockets.exceptions.ConnectionClosed:
                        break
                    except Exception:
                        await asyncio.sleep(0.05)
                        continue
        except Exception as e:
            print(f"Error in webcam_stream_handler: {e}")
        finally:
            try:
                if container is not None:
                    container.close()
            except Exception:
                pass
            try:
                if player is not None:
                    player.audio and player.audio.stop()
                    player.video and player.video.stop()
            except Exception:
                pass
            try:
                if cv2_capture is not None:
                    cv2_capture.release()
            except Exception:
                pass

    async def ssh_ws_handler(self, websocket):
        """Route immediately to a local interactive shell for responsiveness on Windows."""
        await self.local_shell_ws_handler(websocket)
        return

    async def local_shell_ws_handler(self, websocket):
        """Spawn a local shell (PowerShell/cmd) and bridge it over the websocket."""
        import shutil
        use_pty = False
        # Prefer Windows PowerShell for a full shell experience; fallback to cmd.exe
        ps_path = os.path.join(os.environ.get('SystemRoot', 'C\\Windows'), 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe')
        cmd_path = os.path.join(os.environ.get('SystemRoot', 'C\\Windows'), 'System32', 'cmd.exe')
        if not os.path.exists(cmd_path):
            cmd_path = os.environ.get('ComSpec', shutil.which('cmd')) or os.path.join(os.environ.get('SystemRoot', 'C\\Windows'), 'System32', 'cmd.exe')
        try:
            # Use a Windows PTY when available for correct interactive behavior (Backspace, arrow keys)
            if 'HAS_WINPTY' in globals() and HAS_WINPTY:
                try:
                    from winpty import PtyProcess
                    shell_cmd = ps_path if os.path.exists(ps_path) else cmd_path
                    # Start with a sane default size matching the client
                    proc = PtyProcess.spawn(shell_cmd, dimensions=(34, 120))
                    proc_writer = proc
                    proc_reader = proc
                    use_pty = True
                except Exception:
                    proc = None
                    use_pty = False

            if not use_pty:
                shell_cmd = [ps_path, '-NoLogo', '-NoExit'] if os.path.exists(ps_path) else [cmd_path, '/K', 'chcp 65001']
                proc = subprocess.Popen(
                        shell_cmd,
                    stdin=subprocess.PIPE,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    bufsize=0,
                    creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0)
                )
            proc_writer = proc
            proc_reader = proc
            # Proactively show something
            try:
                await websocket.send("Connected to local shell. Type commands and press Enter.\r\n")
            except Exception:
                pass
            # Trigger prompt output
            try:
                if use_pty:
                    proc_writer.write("\r\n")
                else:
                    proc_writer.stdin.write(b"\r\n")
                    proc_writer.stdin.flush()
            except Exception:
                pass
        except Exception as e:
            try:
                await websocket.send(json.dumps({'type': 'error', 'message': f'Local shell failed: {e}'}))
            except Exception:
                pass
            return

        stop_flag = False

        async def ws_to_proc():
            nonlocal stop_flag
            try:
                async for msg in websocket:
                    try:
                        if isinstance(msg, str):
                            data = msg
                        else:
                            data = msg.decode('utf-8', errors='ignore')
                        # Handle terminal resize from client
                        if data and data.startswith('{'):
                            try:
                                evt = json.loads(data)
                                if evt.get('type')=='resize':
                                    cols = int(evt.get('cols', 120))
                                    rows = int(evt.get('rows', 34))
                                    if use_pty:
                                        try:
                                            proc_writer.set_size(rows, cols)
                                        except Exception:
                                            pass
                                    continue
                            except Exception:
                                pass
                        
                        # Normalize input: map DEL->BS always; only expand CR to CRLF for non-PTY
                        if use_pty:
                            data = data.replace('\x7f', '\b')
                            proc_writer.write(data)
                        else:
                            data = data.replace('\r', '\r\n').replace('\x7f', '\b')
                            proc_writer.stdin.write(data.encode('utf-8', errors='ignore'))
                            proc_writer.stdin.flush()
                    except Exception:
                        break
            finally:
                stop_flag = True

        async def proc_to_ws():
            nonlocal stop_flag
            loop = asyncio.get_running_loop()
            import locale
            import re
            enc = 'utf-8'
            poll_alive = (lambda: (proc.isalive() if use_pty else proc.poll() is None))
            while not stop_flag and poll_alive():
                try:
                    # Read in small chunks for responsiveness
                    if use_pty:
                        chunk = await loop.run_in_executor(None, proc_reader.read, 256)
                    else:
                        chunk = await loop.run_in_executor(None, proc_reader.stdout.read, 256)
                    if not chunk:
                        await asyncio.sleep(0.02)
                        continue
                    if isinstance(chunk, bytes):
                        try:
                            text = chunk.decode(enc, errors='ignore')
                        except Exception:
                            text = chunk.decode(locale.getpreferredencoding(False) or 'utf-8', errors='ignore')
                    else:
                        # winpty returns str
                        text = chunk
                    # Normalize bare CR from PTY to CRLF to avoid cursor overlays
                    if use_pty and text:
                        text = re.sub(r"\r(?!\n)", "\r\n", text)
                    await websocket.send(text)
                except Exception:
                    break

        sender = asyncio.create_task(ws_to_proc())
        receiver = asyncio.create_task(proc_to_ws())
        try:
            await asyncio.wait({sender, receiver}, return_when=asyncio.FIRST_COMPLETED)
        finally:
            for t in (sender, receiver):
                if not t.done():
                    t.cancel()
        try:
            if proc:
                if use_pty:
                    try:
                        proc.close()
                    except Exception:
                        pass
                else:
                    if proc.poll() is None:
                        try:
                            proc.terminate()
                            proc.wait(timeout=2)
                        except Exception:
                            try:
                                proc.kill()
                            except Exception:
                                pass
        except Exception:
            pass

    async def broadcast_frames(self):
        """Broadcast new video frames to connected clients."""
        frame_count = 0
        last_debug = 0
        last_sequence = 0
        write_buffer_limit = 256_000
        send_tasks = getattr(self, "_video_send_tasks", None)
        if send_tasks is None:
            send_tasks = {}
            self._video_send_tasks = send_tasks
        if not hasattr(self, "_video_skipped_sends"):
            self._video_skipped_sends = 0

        while not self.stop_event.is_set():
            try:
                disconnected = set()
                for ws, task in list(send_tasks.items()):
                    if not task.done():
                        continue
                    send_tasks.pop(ws, None)
                    try:
                        task.result()
                        frame_count += 1
                    except websockets.exceptions.ConnectionClosed:
                        disconnected.add(ws)
                    except Exception as e:
                        print(f"Error sending frame to client: {e}")
                        disconnected.add(ws)

                for ws in disconnected:
                    self.video_clients.discard(ws)

                event = getattr(self, "frame_ready_event", None)
                if event is not None:
                    try:
                        await asyncio.wait_for(event.wait(), timeout=1.0)
                        event.clear()
                    except asyncio.TimeoutError:
                        pass
                else:
                    try:
                        target_fps = int(self.current_fps)
                    except Exception:
                        target_fps = 0
                    if target_fps > 0:
                        await asyncio.sleep(1 / target_fps)
                    else:
                        await asyncio.sleep(0)

                if not self.screen_capturer:
                    continue
                if hasattr(self.screen_capturer, "get_latest_frame_packet"):
                    sequence, frame = self.screen_capturer.get_latest_frame_packet()
                else:
                    sequence, frame = 0, self.screen_capturer.latest_frame_jpeg
                if not frame or sequence == last_sequence:
                    if self.video_clients and frame_count == 0:
                        print(f"{len(self.video_clients)} video clients waiting, but no frames available")
                        frame_count = 1
                    continue
                last_sequence = sequence

                if not self.video_clients:
                    continue

                max_write_buffer = 0
                for ws in self.video_clients.copy():
                    if ws in disconnected:
                        continue
                    try:
                        transport = getattr(ws, "transport", None)
                        buffer_size = transport.get_write_buffer_size() if transport else 0
                        max_write_buffer = max(max_write_buffer, int(buffer_size or 0))
                        if buffer_size > write_buffer_limit:
                            self._video_skipped_sends += 1
                            continue
                        existing = send_tasks.get(ws)
                        if existing is not None and not existing.done():
                            self._video_skipped_sends += 1
                            continue
                        send_tasks[ws] = asyncio.create_task(ws.send(frame))
                    except websockets.exceptions.ConnectionClosed:
                        disconnected.add(ws)
                    except Exception as e:
                        print(f"Error scheduling frame send: {e}")
                        disconnected.add(ws)

                for ws in disconnected:
                    self.video_clients.discard(ws)
                    task = send_tasks.pop(ws, None)
                    if task and not task.done():
                        task.cancel()

                try:
                    capture_stats = self.screen_capturer.get_capture_stats()
                    frame_ts = float(capture_stats.get("last_frame_ts") or 0.0)
                    frame_age_ms = max(0.0, (time.time() - frame_ts) * 1000.0) if frame_ts else 0.0
                    controller = getattr(self, "adaptive_stream", None)
                    if controller is not None and controller.observe_server(
                        frame_bytes=len(frame),
                        max_write_buffer=max_write_buffer,
                        skipped_total=getattr(self, "_video_skipped_sends", 0),
                        inflight_sends=sum(1 for task in send_tasks.values() if not task.done()),
                        video_clients=len(self.video_clients),
                        frame_age_ms=frame_age_ms,
                        capture_stats=capture_stats,
                    ):
                        self._apply_stream_profile("adaptive_" + str(controller.last_reason))
                        await self._broadcast_stream_status()
                except Exception:
                    logging.debug("Adaptive stream observe failed", exc_info=True)

                if frame_count - last_debug >= 60:
                    logging.debug(
                        "Sent %s frames to %s video client(s); seq=%s bytes=%s skipped=%s inflight=%s stream=%s",
                        frame_count,
                        len(self.video_clients),
                        sequence,
                        len(frame),
                        getattr(self, "_video_skipped_sends", 0),
                        sum(1 for task in send_tasks.values() if not task.done()),
                        getattr(getattr(self, "adaptive_stream", None), "profile", None).name if getattr(self, "adaptive_stream", None) else "manual",
                    )
                    last_debug = frame_count
            except Exception as e:
                print(f"Error in broadcast_frames: {e}")
                await asyncio.sleep(0.05)

        for task in send_tasks.values():
            if not task.done():
                task.cancel()
    async def broadcast_cursor_position(self):
        """Broadcast cursor position and button states to subscribed clients"""
        while not self.stop_event.is_set():
            if not (self.cursor_broadcast_enabled and self.cursor_subscribers):
                await asyncio.sleep(0.25)
                continue
            if self.cursor_broadcast_enabled and self.cursor_subscribers:
                try:
                    # Get cursor position using GetCursorInfo first (avoids LP_POINT issues)
                    ci = CURSORINFO()
                    ci.cbSize = ctypes.sizeof(CURSORINFO)
                    if user32.GetCursorInfo(ctypes.byref(ci)):
                        x, y = ci.ptScreenPos.x, ci.ptScreenPos.y
                    else:
                        # Fallback to GetCursorPos if needed
                        pt = POINT()
                        if _GetCursorPos(ctypes.byref(pt)):
                            x, y = pt.x, pt.y
                        else:
                            x, y = 0, 0
                    
                    # Get button states using Windows API
                    left_press = bool(_GetAsyncKeyState(VK_LBUTTON) & 0x8000)
                    right_press = bool(_GetAsyncKeyState(VK_RBUTTON) & 0x8000)
                    
                    # Get virtual screen bounds for normalization
                    vx = user32.GetSystemMetrics(SM_XVIRTUALSCREEN)
                    vy = user32.GetSystemMetrics(SM_YVIRTUALSCREEN)
                    vw = user32.GetSystemMetrics(SM_CXVIRTUALSCREEN)
                    vh = user32.GetSystemMetrics(SM_CYVIRTUALSCREEN)
                    
                    # Normalize coordinates to 0-1 range
                    if vw > 0 and vh > 0:
                        norm_x = (x - vx) / vw
                        norm_y = (y - vy) / vh
                    else:
                        norm_x = 0.5
                        norm_y = 0.5
                    
                    # Create cursor data
                    cursor_data = {
                        'type': 'cursor',
                        'x': norm_x,
                        'y': norm_y,
                        'left_pressed': left_press,
                        'right_pressed': right_press,
                        'cursor_css': _get_css_cursor_from_system()
                    }
                    
                    # Send to all subscribed clients
                    disconnected = []
                    for ws in self.cursor_subscribers.copy():
                        try:
                            await ws.send(json.dumps(cursor_data))
                        except websockets.exceptions.ConnectionClosed:
                            disconnected.append(ws)
                        except Exception as e:
                            print(f"Error sending cursor data to client: {e}")
                            disconnected.append(ws)
                    
                    # Remove disconnected clients
                    for ws in disconnected:
                        self.cursor_subscribers.discard(ws)
                        if not self.cursor_subscribers:
                            self.cursor_broadcast_enabled = False
                            
                except Exception as e:
                    print(f"Error in broadcast_cursor_position: {e}")
            
            # Update at 60 FPS for smooth cursor tracking while subscribed.
            await asyncio.sleep(1/60)
