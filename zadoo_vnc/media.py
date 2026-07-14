"""Video, audio, webcam, cursor, and shell stream handlers."""
from __future__ import annotations

import asyncio
import ctypes
import json
import logging
import queue
import sys
import threading
import time
import urllib.parse
from contextlib import suppress
from pathlib import Path

import av
import numpy as np
import soundcard as sc
import sounddevice as sd
import websockets
from imagecodecs._jpeg8 import jpeg8_encode
from winpty import Backend, PtyProcess

from .camera_discovery import camera_open_target
from .config import windows_system_executable
from .win32_input import (
    CURSOR_SHOWING,
    CURSORINFO,
    SM_CXVIRTUALSCREEN,
    SM_CYVIRTUALSCREEN,
    SM_XVIRTUALSCREEN,
    SM_YVIRTUALSCREEN,
    VK_LBUTTON,
    VK_RBUTTON,
    _get_css_cursor,
    _GetAsyncKeyState,
    user32,
)


async def _spawn_conpty(shell_cmd, shell_cwd, dimensions):
    """Start ConPTY off the event loop and recover its transient OpenConsole startup race."""
    for attempt in range(2):
        try:
            return await asyncio.to_thread(
                PtyProcess.spawn,
                shell_cmd,
                cwd=shell_cwd,
                dimensions=dimensions,
                backend=Backend.ConPTY,
            )
        except (KeyboardInterrupt, SystemExit):
            raise
        except BaseException as exc:
            detail = str(exc).strip() or exc.__class__.__name__
            if attempt == 0 and "HRESULT(0x800700BB)" in detail:
                logging.warning("ConPTY startup raced with OpenConsole; retrying once: %s", detail)
                await asyncio.sleep(0.1)
                continue
            raise RuntimeError(f"ConPTY startup failed: {detail}") from exc
    raise AssertionError("ConPTY retry loop returned without a result")


class MediaMixin:

    def _put_realtime_frame(self, frame_queue, frame):
        """Offer a realtime media frame without ever blocking the capture thread."""
        try:
            frame_queue.put_nowait(frame)
            return True
        except queue.Full:
            with suppress(queue.Empty):
                frame_queue.get_nowait()
            try:
                frame_queue.put_nowait(frame)
                return True
            except queue.Full:
                return False

    def _clear_media_queue(self, frame_queue):
        with frame_queue.mutex:
            frame_queue.queue.clear()

    def _media_queues(self, clients):
        with self._media_clients_lock:
            return tuple(clients.values())

    def _register_media_client(self, clients, websocket):
        frame_queue = queue.Queue(maxsize=5)
        with self._media_clients_lock:
            clients[websocket] = frame_queue
        return frame_queue

    def _unregister_media_client(self, clients, websocket):
        with self._media_clients_lock:
            clients.pop(websocket, None)
            return len(clients)

    def _broadcast_realtime_frame(self, clients, frame):
        for frame_queue in self._media_queues(clients):
            self._put_realtime_frame(frame_queue, frame)

    def _clear_media_queues(self, clients):
        for frame_queue in self._media_queues(clients):
            self._clear_media_queue(frame_queue)

    def _wake_media_queue_readers(self, clients):
        self._broadcast_realtime_frame(clients, None)

    def _close_audio_stream(self):
        stream = self._audio_stream
        if not stream:
            return
        self._audio_stream = None
        try:
            stream.__exit__(None, None, None)
        except Exception as exc:
            raise RuntimeError(f"System-audio stream cleanup failed: {exc}") from exc

    def _clarify_mono_audio(
        self,
        samples,
        state,
        *,
        target_rms=0.12,
        noise_floor=0.002,
        max_gain=4.0,
        limiter=0.96,
    ):
        """Small realtime clarity chain: DC removal, soft gate, smooth gain, limiter."""
        arr = np.asarray(samples, dtype=np.float32).reshape(-1)
        if arr.size == 0:
            return arr
        arr = np.nan_to_num(arr, copy=False)
        arr -= float(np.mean(arr))
        rms = float(np.sqrt(np.mean(arr * arr) + 1e-12))
        if rms < noise_floor:
            gate = max(0.0, min(1.0, (rms / max(noise_floor, 1e-9)) ** 2))
            arr *= gate
            desired_gain = 1.0
        else:
            desired_gain = max(0.5, min(float(max_gain), float(target_rms) / max(rms, 1e-6)))
        prev_gain = float(state["gain"])
        smoothing = 0.12 if desired_gain > prev_gain else 0.35
        gain = prev_gain + (desired_gain - prev_gain) * smoothing
        state["gain"] = gain
        arr *= gain
        peak = float(np.max(np.abs(arr)))
        if peak > limiter:
            arr *= float(limiter) / peak
        return np.clip(arr, -1.0, 1.0, out=arr)

    def _apply_stream_profile(self, reason="adaptive"):
        self.adaptive_stream.last_reason = str(reason or self.adaptive_stream.last_reason)
        return self.adaptive_stream.apply_to(self)

    def _set_manual_performance(self, event):
        capturer = self.screen_capturer
        if capturer is None:
            raise RuntimeError("Screen capturer is not configured")
        enabled = event['enabled']
        region = event['region']
        scale_div = event['scale_div']
        grayscale = event['grayscale']
        capturer.configure_performance(enabled, region, scale_div, grayscale, event.get('rect_norm'))
        self._manual_performance = (
            region,
            scale_div,
            grayscale,
            dict(capturer._custom_rect_norm) if region == 'custom' else None,
        ) if enabled else None
        self._apply_stream_profile("manual_performance")

    def _stream_status_payload(self):
        status = self.adaptive_stream.status()
        status["video_clients"] = len(self.video_clients)
        status["skipped_sends"] = self._video_skipped_sends
        status["inflight_sends"] = sum(
            1 for task in self._video_send_tasks.values() if not task.done()
        )
        status["effective_quality"] = self.current_quality
        status["quality_locked"] = self._quality_locked_by_user
        return status

    async def _send_stream_status(self, websocket):
        await websocket.send(json.dumps({
            "type": "stream_adaptation",
            "stream": self._stream_status_payload(),
        }))

    async def _broadcast_stream_status(self):
        if not self.video_clients:
            return
        message = json.dumps({
            "type": "stream_adaptation",
            "stream": self._stream_status_payload(),
        })
        for ws in list(self.video_clients):
            try:
                await ws.send(message)
            except Exception as exc:
                logging.error("Stream-status send failed: %s", exc)
                self.video_clients.discard(ws)

    def _capture_stats_payload(self):
        if self.screen_capturer is None:
            raise RuntimeError("Screen capturer is not configured")
        stats = self.screen_capturer.get_capture_stats()
        send_tasks = self._video_send_tasks
        stream = self._stream_status_payload()
        stats.update({
            "video_clients": len(self.video_clients),
            "skipped_sends": self._video_skipped_sends,
            "inflight_sends": sum(1 for task in send_tasks.values() if not task.done()),
            "stream_profile": stream["profile_name"],
            "stream_status": stream["status"],
            "encoder_backend": "imagecodecs_jpeg",
            "client_stream_stats": stream["client"],
            "server_stream_stats": stream["server"],
        })
        return stats

    def _start_audio_capture(self):
        """Capture 20 ms mono PCM packets from the current Windows output device."""
        if self._audio_running:
            return
        self._clear_media_queues(self.audio_clients)
        self._audio_error = None
        self._audio_clarity_state = {"gain": 1.0}

        log = logging.getLogger("sysaudio")

        def _pick_loopback():
            spk = sc.default_speaker()
            speaker_name = spk.name.split(" (")[0]
            for mic in sc.all_microphones(include_loopback=True):
                if mic.isloopback and (
                    speaker_name in mic.name or (spk.id and spk.id in mic.id)
                ):
                    return mic
            raise RuntimeError(f"No system-audio loopback matched the default speaker: {spk.name}")

        def _open_loopback(loopback):
            try:
                stream = loopback.recorder(samplerate=self._audio_sr, channels=2)
                stream.__enter__()
                return stream
            except Exception as exc:
                raise RuntimeError(f"Could not open system-audio loopback {loopback.name}: {exc}") from exc

        loopback = _pick_loopback()
        self._audio_stream = _open_loopback(loopback)
        stop_evt = threading.Event()
        self._audio_stop_evt = stop_evt
        self._audio_running = True

        def _audio_worker():
            try:
                while self._audio_running and not stop_evt.is_set():
                    data = self._audio_stream.record(numframes=960)
                    if data.ndim != 2 or data.shape[0] != 960 or data.shape[1] < 1:
                        raise RuntimeError(f"System-audio loopback returned invalid shape: {data.shape}")
                    mono = self._clarify_mono_audio(
                        data[:, 0],
                        self._audio_clarity_state,
                        target_rms=0.13,
                        noise_floor=0.0012,
                        max_gain=3.2,
                    )
                    pcm = (mono * 32767.0).astype(np.int16).tobytes()
                    self._broadcast_realtime_frame(self.audio_clients, pcm)
            except Exception as exc:
                self._audio_error = f"System audio capture failed: {exc}"
                log.error(self._audio_error)
            finally:
                try:
                    self._close_audio_stream()
                except Exception as exc:
                    self._audio_error = str(exc)
                    log.error(self._audio_error)
                self._audio_running = False
                if self._audio_stop_evt is stop_evt:
                    self._audio_stop_evt = None
                self._wake_media_queue_readers(self.audio_clients)
                log.info("System-audio worker stopped")

        t = threading.Thread(target=_audio_worker, daemon=True)
        t.start()
        self._audio_thread = t

    def _stop_audio_capture(self):
        self._audio_running = False
        if self._audio_stop_evt is not None:
            self._audio_stop_evt.set()
        if self._audio_thread and self._audio_thread.is_alive():
            self._audio_thread.join(timeout=0.5)
            if self._audio_thread.is_alive():
                raise RuntimeError("System-audio capture thread did not stop within 0.5 seconds")
        if self._audio_stream is not None:
            self._close_audio_stream()
        self._audio_stream = None
        self._audio_thread = None
        self._clear_media_queues(self.audio_clients)
        self._wake_media_queue_readers(self.audio_clients)
        logging.info(" Audio capture stopped")

    def _normalize_mic_device_id(self, device):
        if device is None:
            return "default"
        raw = str(device).strip()
        if not raw:
            raise ValueError("Microphone device id cannot be empty")
        if raw.lower() == "default":
            return "default"
        try:
            return str(int(raw))
        except ValueError as exc:
            raise ValueError(f"Invalid microphone device id: {raw}") from exc

    def _resolve_mic_device(self, device):
        device_id = self._normalize_mic_device_id(device)
        index = None if device_id == "default" else int(device_id)
        try:
            info = sd.query_devices(index, kind="input")
        except Exception as exc:
            label = "system default" if index is None else str(index)
            raise RuntimeError(f"Could not query microphone device {label}: {exc}") from exc
        label = "system default" if index is None else str(index)
        try:
            channels = int(info["max_input_channels"])
        except (KeyError, TypeError, ValueError, OverflowError) as exc:
            raise RuntimeError(f"Microphone device {label} has invalid max_input_channels: {exc}") from exc
        if channels <= 0:
            raise RuntimeError(f"Microphone device {label} has no input channels")
        return index, device_id, info

    def _start_mic_capture(self, device=None):
        """Start microphone capture using sounddevice InputStream. Sends mono s16le frames."""
        mic_log = logging.getLogger("mic")
        if self.mic_running:
            return
        self._clear_media_queues(self.mic_clients)
        self._mic_error = None
        self._mic_clarity_state = {"gain": 1.0}
        device_arg, device_id, dev = self._resolve_mic_device(device)
        try:
            use_sr = int(float(dev["default_samplerate"]))
        except (KeyError, TypeError, ValueError, OverflowError) as exc:
            raise RuntimeError(f"Microphone {device_id} has invalid default sample rate: {exc}") from exc
        if use_sr <= 0:
            raise RuntimeError(f"Microphone {device_id} has no default sample rate")
        try:
            raw_name = dev["name"]
        except KeyError as exc:
            raise RuntimeError(f"Microphone {device_id} is missing name") from exc
        if not isinstance(raw_name, str) or not raw_name.strip():
            raise RuntimeError(f"Microphone {device_id} has no name")
        device_name = raw_name.strip()
        self.mic_device_id = device_id
        self.mic_device_name = device_name
        mic_log.info("Mic device: id=%s name=%s sr=%s", self.mic_device_id, self.mic_device_name, use_sr)

        self.mic_samplerate = use_sr
        def mic_callback(indata, _frames, _time_info, status):
            try:
                if not self.mic_running or not self.mic_clients:
                    return
                if status:
                    logging.debug("Mic status: %s", status)
                x = indata[:, 0]
                x = self._clarify_mono_audio(x, self._mic_clarity_state,
                                             target_rms=0.16, noise_floor=0.0035, max_gain=5.0)
                self._broadcast_realtime_frame(
                    self.mic_clients,
                    (x * 32767.0).astype(np.int16).tobytes(),
                )
            except Exception as exc:
                self._mic_error = f"Microphone capture failed: {exc}"
                self.mic_running = False
                self._wake_media_queue_readers(self.mic_clients)
                mic_log.error(self._mic_error, exc_info=True)

        try:
            mic_log.info("Opening mic stream device=%s sr=%s block=%s", self.mic_device_id, use_sr, self.mic_blocksize)
            self.mic_stream = sd.InputStream(
                samplerate=use_sr, channels=1, dtype='float32',
                blocksize=self.mic_blocksize, callback=mic_callback,
                device=device_arg, latency='low'
            )
            self.mic_stream.start()
            self.mic_running = True
            mic_log.info("Mic capture started device=%s @ %s Hz block=%s", self.mic_device_name, use_sr, self.mic_blocksize)
        except Exception as exc:
            self.mic_stream = None
            self.mic_running = False
            raise RuntimeError(f"Could not open microphone {self.mic_device_name}: {exc}") from exc

    def _stop_mic_capture(self):
        mic_log = logging.getLogger("mic")
        self.mic_running = False
        stream = self.mic_stream
        self.mic_stream = None
        errors = []
        if stream:
            try:
                stream.stop()
            except Exception as exc:
                errors.append(f"stop failed: {exc}")
            try:
                stream.close()
            except Exception as exc:
                errors.append(f"close failed: {exc}")
        try:
            self._clear_media_queues(self.mic_clients)
            self._wake_media_queue_readers(self.mic_clients)
        except Exception as exc:
            errors.append(f"queue cleanup failed: {exc}")
        if errors:
            error = "Microphone cleanup failed: " + "; ".join(errors)
            mic_log.error(error)
            raise RuntimeError(error)
        mic_log.info("Mic capture stopped and queue cleared")

    async def audio_stream_handler(self, websocket: websockets.WebSocketServerProtocol):
        print(f" New audio client connected from {websocket.remote_address}")
        frame_queue = self._register_media_client(self.audio_clients, websocket)

        if not self._audio_running:
            try:
                self._start_audio_capture()
            except Exception as exc:
                error = f"System audio capture failed: {exc}"
                await websocket.send(json.dumps({"type": "error", "error": error}))
                await websocket.close(code=1011, reason="System audio capture failed")
                self._unregister_media_client(self.audio_clients, websocket)
                return

        try:
            # Simple header to let client know PCM format
            header = json.dumps({
                'type': 'audio_format',
                'codec': 'pcm_s16le',
                'samplerate': self._audio_sr,
                'channels': 1,
                'blocksize': 960
            }).encode()
            await websocket.send(header)

            loop = asyncio.get_running_loop()
            while True:
                try:
                    chunk = await loop.run_in_executor(None, frame_queue.get, True, 0.25)
                except queue.Empty:
                    continue
                if chunk is None:
                    if not self._audio_running:
                        if self._audio_error:
                            await websocket.send(json.dumps({"type": "error", "error": self._audio_error}))
                        break
                    await asyncio.sleep(0.005)
                    continue
                try:
                    # Keep audio low-latency: if the socket is already backed up, drop this frame
                    # instead of piling on more delay (a tiny gap beats seconds of lag).
                    if websocket.transport.get_write_buffer_size() > 16000:
                        continue
                    await websocket.send(chunk)
                except websockets.exceptions.ConnectionClosed:
                    break
        except websockets.exceptions.ConnectionClosed:
            pass
        except Exception as exc:
            error = f"System audio stream failed: {exc}"
            logging.exception(error)
            with suppress(websockets.exceptions.ConnectionClosed):
                await websocket.send(json.dumps({"type": "error", "error": error}))
                await websocket.close(code=1011, reason="System audio stream failed")
        finally:
            remaining = self._unregister_media_client(self.audio_clients, websocket)
            print(f" Removed audio client, {remaining} clients remaining")
            if not remaining and self._audio_running:
                self._stop_audio_capture()
                print(" Audio capture stopped (no clients)")

    def _remove_mic_client(self, websocket):
        if not self._unregister_media_client(self.mic_clients, websocket):
            self._stop_mic_capture()
            return True
        return False

    def _cloud_start_remote_session(self):
        if not self.settings_store.get_device_token():
            return None
        from .saas import ZadooCloudClient

        public_url = self.tunnel_manager.get_current_url() if self.tunnel_manager else None
        result = ZadooCloudClient(self.settings_store).start_session(public_url)
        if not result["success"] and result.get("status_code") == 402:
            entitlement = ZadooCloudClient._required_entitlement(result, "Session start")
            if (
                not entitlement["revoked"]
                and entitlement["includedMinutesRemaining"] == 0
                and entitlement["walletMinutesRemaining"] == 0
            ):
                return {"success": True, "no_credits": True}
        return result

    def _cloud_end_remote_session(self, session_id):
        if not session_id:
            return
        from .saas import ZadooCloudClient

        result = ZadooCloudClient(self.settings_store).end_session(session_id)
        if not result["success"]:
            raise RuntimeError(f"Cloud session end failed: {result['error']}")

    async def _cloud_remote_session_heartbeat(self, websocket, session_state, start_in_grace=False):
        # Grace policy once the balance reaches 0 (heartbeat reports allowed=False):
        #   minute 0..5  → full control kept; client shows "add credit" banner + payment panel
        #   minute 5..10 → controls locked (screen-share only); payment panel keeps prompting
        #   minute 10    → session ends
        # start_in_grace=True means we connected with zero credits and run the grace
        # policy from the start (no billing session until a top-up resumes us).
        GRACE_LOCK_AFTER = 5
        GRACE_END_AFTER = 10
        try:
            from .saas import ZadooCloudClient

            client = ZadooCloudClient(self.settings_store)
            grace_minutes = 0
            in_grace = bool(start_in_grace)
            if in_grace:
                self._set_grace_lock(websocket, False)
                await self._send_grace(websocket, 0, False)
            while True:
                await asyncio.sleep(60)
                if not in_grace:
                    result = await asyncio.to_thread(client.session_heartbeat, session_state["id"], 1)
                    if not result["success"]:
                        raise RuntimeError(f"Cloud session heartbeat failed: {result['error']}")
                    entitlement = result["entitlement"]
                    allowed = entitlement["allowed"]
                    if allowed is False:
                        # Balance just hit zero — begin the local grace window.
                        in_grace = True
                        grace_minutes = 1
                        self._set_grace_lock(websocket, False)
                        await self._send_grace(websocket, grace_minutes, False)
                    continue
                # In grace: the cloud has already ended the session record, so stop billing
                # and instead watch the workspace entitlement for a top-up that resumes us.
                ent = await asyncio.to_thread(client.entitlement)
                if not ent["success"]:
                    raise RuntimeError(f"Cloud entitlement refresh failed: {ent['error']}")
                e = ent["entitlement"]
                recharged = e["allowed"] and not e["revoked"]
                if recharged:
                    in_grace = False
                    grace_minutes = 0
                    self._set_grace_lock(websocket, False)
                    cs = await asyncio.to_thread(self._cloud_start_remote_session)
                    if cs is None:
                        raise RuntimeError("Cloud session restart requires an activated device")
                    if not cs["success"]:
                        raise RuntimeError(f"Cloud session restart failed: {cs['error']}")
                    if cs.get("no_credits") is True:
                        raise RuntimeError("Cloud session restart was rejected for insufficient credits")
                    session_state["id"] = cs["sessionId"]
                    await self._send_grace(websocket, 0, False, recharged=True)
                    continue
                grace_minutes += 1
                if grace_minutes >= GRACE_END_AFTER:
                    self._set_grace_lock(websocket, False)
                    with suppress(websockets.exceptions.ConnectionClosed):
                        await websocket.send(json.dumps({"type": "error", "error": "Session ended — out of credits"}))
                        await websocket.close(code=1008, reason="Out of credits")
                    return
                locked = grace_minutes >= GRACE_LOCK_AFTER
                self._set_grace_lock(websocket, locked)
                await self._send_grace(websocket, grace_minutes, locked)
        except asyncio.CancelledError:
            self._set_grace_lock(websocket, False)
            raise
        except Exception as exc:
            self._set_grace_lock(websocket, False)
            with suppress(websockets.exceptions.ConnectionClosed):
                await websocket.send(json.dumps({"type": "error", "error": f"Cloud session heartbeat failed: {exc}"}))
                await websocket.close(code=1011, reason="Cloud session heartbeat failed")

    async def _send_grace(self, websocket, grace_minutes, controls_locked, recharged=False):
        await websocket.send(json.dumps({
            "type": "grace",
            "graceMinutes": int(grace_minutes),
            "controlsLocked": bool(controls_locked),
            "recharged": bool(recharged),
        }))

    async def mic_stream_handler(self, websocket: websockets.WebSocketServerProtocol, path=None):
        mic_log = logging.getLogger("mic")
        remote = getattr(websocket, 'remote_address', None)
        mic_log.info("New mic client connected from %s", remote)
        print(f"New mic client connected from {remote}")
        frame_queue = self._register_media_client(self.mic_clients, websocket)
        try:
            params = urllib.parse.parse_qs(
                urllib.parse.urlparse(path or "").query,
                keep_blank_values=True,
                strict_parsing=True,
            )
            if set(params) - {"device"}:
                raise ValueError(f"Unsupported microphone parameters: {', '.join(sorted(set(params) - {'device'}))}")
            if "device" in params and len(params["device"]) != 1:
                raise ValueError("Duplicate microphone parameter: device")
            requested_device = self._normalize_mic_device_id(params["device"][0] if "device" in params else None)
            if self.mic_running and self.mic_device_id != requested_device:
                mic_log.info(
                    "Mic device changed from %s to %s; restarting capture",
                    self.mic_device_id,
                    requested_device,
                )
                self._stop_mic_capture()
            if not self.mic_running:
                try:
                    self._start_mic_capture(device=requested_device)
                except Exception as exc:
                    error = f"Microphone capture failed: {exc}"
                    mic_log.error("%s for %s", error, remote)
                    await websocket.send(json.dumps({"type": "error", "error": error}))
                    await websocket.close(code=1011, reason="Microphone capture failed")
                    return

            # Send format header after opening so the samplerate reflects the actual stream.
            hdr = {
                "type": "audio_format",
                "samplerate": int(self.mic_samplerate),
                "channels": 1,
                "samplefmt": "s16le",
                "blocksize": int(self.mic_blocksize),
                "device_id": self.mic_device_id,
                "device_name": self.mic_device_name,
            }
            await websocket.send(json.dumps(hdr).encode('utf-8'))
            mic_log.info("Mic header sent to %s: sr=%s ch=%s fmt=%s block=%s", remote, hdr['samplerate'], hdr['channels'], hdr['samplefmt'], hdr['blocksize'])
            print(f"Mic header sent: sr={hdr['samplerate']} ch={hdr['channels']} fmt={hdr['samplefmt']} block={hdr['blocksize']}")

            sent_chunks = 0
            loop = asyncio.get_running_loop()
            while True:
                try:
                    chunk = await loop.run_in_executor(None, frame_queue.get, True, 0.25)
                except queue.Empty:
                    continue
                if chunk is None:
                    if not self.mic_running:
                        if self._mic_error:
                            await websocket.send(json.dumps({"type": "error", "error": self._mic_error}))
                        break
                    await asyncio.sleep(0.005)
                    continue
                try:
                    # Keep voice low-latency: drop this frame if the socket is already backed up.
                    if websocket.transport.get_write_buffer_size() > 16000:
                        continue
                    await websocket.send(chunk)
                    sent_chunks += 1
                    if sent_chunks == 1 or sent_chunks % 250 == 0:
                        mic_log.info("Mic sent %s chunk(s) to %s; last_chunk_bytes=%s", sent_chunks, remote, len(chunk))
                except websockets.exceptions.ConnectionClosed:
                    break
        except websockets.exceptions.ConnectionClosed:
            pass
        except Exception as exc:
            error = f"Microphone stream failed: {exc}"
            mic_log.error("%s for %s", error, remote, exc_info=True)
            with suppress(websockets.exceptions.ConnectionClosed):
                await websocket.send(json.dumps({"type": "error", "error": error}))
                await websocket.close(code=1011, reason="Microphone stream failed")
        finally:
            try:
                stopped = self._remove_mic_client(websocket)
            except Exception as exc:
                stopped = False
                mic_log.error("Microphone cleanup failed for %s: %s", remote, exc, exc_info=True)
                with suppress(websockets.exceptions.ConnectionClosed):
                    await websocket.send(json.dumps({"type": "error", "error": str(exc)}))
            mic_log.info("Mic client disconnected from %s; remaining=%s stopped=%s", remote, len(self.mic_clients), stopped)
            print(f"Mic client disconnected; remaining={len(self.mic_clients)} stopped={stopped}")

    async def video_stream_handler(self, websocket):
        print(f" New video client connected from {websocket.remote_address}")
        cloud_session = {"id": None}
        cloud_heartbeat_task = None
        
        try:
            cloud_start = await asyncio.to_thread(self._cloud_start_remote_session)
            if cloud_start is not None and not cloud_start["success"]:
                with suppress(websockets.exceptions.ConnectionClosed):
                    await websocket.send(json.dumps({
                        "type": "error",
                        "error": cloud_start["error"],
                    }))
                    await websocket.close(code=1008, reason="Entitlement blocked")
                return
            if cloud_start is not None and cloud_start.get("no_credits") is True:
                cloud_heartbeat_task = asyncio.create_task(
                    self._cloud_remote_session_heartbeat(websocket, cloud_session, start_in_grace=True)
                )
            elif cloud_start is not None:
                cloud_session["id"] = cloud_start["sessionId"]
                self._set_grace_lock(websocket, False)
                cloud_heartbeat_task = asyncio.create_task(
                    self._cloud_remote_session_heartbeat(websocket, cloud_session)
                )
            self.video_clients.add(websocket)
            self.screen_capturer.set_streaming_active(True)
            self._apply_stream_profile("client_connected")
            await self._send_stream_status(websocket)
            # Keep connection alive and handle any incoming messages
            async for message in websocket:
                event = None
                try:
                    # Handle any control messages for video stream
                    event = json.loads(message)
                    if not isinstance(event, dict):
                        raise ValueError("Video control event must be a JSON object")
                    action = event.get('action')
                    if not self._is_ws_action_authorized(websocket, action):
                        await self._send_ws_forbidden(websocket, action)
                        continue
                     
                    if action == 'set_quality':
                        value = self._apply_quality(event['value'])
                        print(f" Quality set to: {value}%")
                    elif action == 'client_stream_stats':
                        self.adaptive_stream.record_client_stats(event)
                    elif action == 'stream_ping':
                        client_ts = event.get('client_ts')
                        if isinstance(client_ts, bool) or not isinstance(client_ts, (int, float)) or not np.isfinite(client_ts):
                            raise ValueError("stream_ping client_ts must be a finite number")
                        await websocket.send(json.dumps({
                            'type': 'stream_pong',
                            'client_ts': client_ts,
                            'server_ts': time.time() * 1000.0,
                        }))
                    elif action == 'get_stream_status':
                        await self._send_stream_status(websocket)
                    elif action == 'set_performance':
                        self._set_manual_performance(event)
                    elif action == 'get_capture_stats':
                        await websocket.send(json.dumps({
                            'type': 'capture_stats',
                            'stats': self._capture_stats_payload()
                        }))
                    else:
                        raise ValueError(f"Unsupported video action: {action}")
                except Exception as e:
                    await websocket.send(json.dumps({
                        "type": "control_error",
                        "action": event.get("action") if isinstance(event, dict) else None,
                        "error": str(e),
                    }))

        except websockets.exceptions.ConnectionClosed:
            print(f" Video client {websocket.remote_address} disconnected")
        except Exception as e:
            logging.exception("Video stream failed")
            with suppress(websockets.exceptions.ConnectionClosed):
                await websocket.send(json.dumps({"type": "error", "error": f"Video stream failed: {e}"}))
                await websocket.close(code=1011, reason="Video stream failed")
        finally:
            if cloud_heartbeat_task and not cloud_heartbeat_task.done():
                cloud_heartbeat_task.cancel()
                await asyncio.gather(cloud_heartbeat_task, return_exceptions=True)
            self.video_clients.discard(websocket)
            self._set_grace_lock(websocket, False)
            if not self.video_clients:
                self.screen_capturer.set_streaming_active(False)
            task = self._video_send_tasks.pop(websocket, None)
            if task and not task.done():
                task.cancel()
                await asyncio.gather(task, return_exceptions=True)
            print(f" Removed video client, {len(self.video_clients)} clients remaining")
            if cloud_session["id"]:
                await asyncio.to_thread(self._cloud_end_remote_session, cloud_session["id"])

    async def webcam_stream_handler(self, websocket):
        """Stream JPEG frames from the server's webcam (DirectShow on Windows) to the client."""
        print(f" Webcam client connected from {websocket.remote_address}")

        container = None
        try:
            try:
                first = await asyncio.wait_for(websocket.recv(), timeout=5.0)
            except TimeoutError as exc:
                raise TimeoutError("Camera selection was not received within 5 seconds") from exc
            if isinstance(first, bytes):
                first = first.decode("utf-8")
            if not isinstance(first, str):
                raise ValueError("Camera selection must be a JSON text message")
            try:
                data = json.loads(first)
            except json.JSONDecodeError as exc:
                raise ValueError(f"Invalid camera selection JSON: {exc.msg}") from exc
            if not isinstance(data, dict) or data.get("type") != "select_camera":
                raise ValueError("Camera selection type must be select_camera")
            selected_id = data.get("camera_id")
            if not isinstance(selected_id, str) or not selected_id.strip():
                raise ValueError("camera_id is required")
            selected_id = selected_id.strip()

            loop = asyncio.get_running_loop()
            devices = await loop.run_in_executor(None, self.enumerate_cameras)
            selected_device = next(
                (device for device in devices if device["id"] == selected_id),
                None,
            )
            if selected_device is None:
                raise ValueError(f"Selected camera was not found: {selected_id}")

            target = camera_open_target(selected_device)
            if not sys.platform.startswith('win'):
                raise RuntimeError("Webcam streaming requires Windows DirectShow")

            device = f"video={target}"
            options = {'framerate': '30', 'video_size': '640x480'}
            try:
                container = await asyncio.to_thread(av.open, device, format='dshow', options=options)
            except Exception as exc:
                raise RuntimeError(f"Selected camera could not be opened: {selected_id}; {exc}") from exc
            print(f" Opened webcam via DirectShow device={device} options={options}")

            stream = next((item for item in container.streams if item.type == 'video'), None)
            if stream is None:
                raise RuntimeError(f"Selected camera has no video stream: {selected_id}")
            stream.thread_type = 'AUTO'
            decoder = container.decode(video=int(stream.index))

            def next_pyav_jpeg():
                try:
                    frame = next(decoder)
                except StopIteration:
                    return None
                pixels = frame.to_ndarray(format="rgb24")
                return jpeg8_encode(pixels, level=70)

            while True:
                jpeg = await asyncio.to_thread(next_pyav_jpeg)
                if jpeg is None:
                    raise RuntimeError(f"Selected camera stopped producing frames: {selected_id}")
                await websocket.send(jpeg)
        except websockets.exceptions.ConnectionClosed:
            pass
        except Exception as exc:
            logging.exception("Webcam stream failed")
            with suppress(websockets.exceptions.ConnectionClosed):
                await websocket.send(json.dumps({"error": str(exc)}))
        finally:
            try:
                if container is not None:
                    container.close()
            except Exception:
                logging.exception("Webcam container cleanup failed")

    async def local_shell_ws_handler(self, websocket):
        """Spawn PowerShell and bridge it over the websocket."""
        initial_cols = 120
        initial_rows = 34
        pending_client_message = None
        ps_path = windows_system_executable("WindowsPowerShell", "v1.0", "powershell.exe")

        def _message_to_text(message):
            if isinstance(message, str):
                return message
            return message.decode('utf-8')

        def _resize_from_message(message):
            data = _message_to_text(message).strip()
            if not data.startswith('{'):
                return None
            try:
                event = json.loads(data)
            except json.JSONDecodeError:
                return None
            if not isinstance(event, dict) or event.get('type') != 'resize':
                return None
            if set(event) != {'type', 'cols', 'rows'}:
                raise ValueError("Terminal resize must contain exactly: type, cols, rows")
            cols = event['cols']
            rows = event['rows']
            if isinstance(cols, bool) or not isinstance(cols, int) or not 20 <= cols <= 500:
                raise ValueError("Terminal resize cols must be an integer from 20 to 500")
            if isinstance(rows, bool) or not isinstance(rows, int) or not 5 <= rows <= 200:
                raise ValueError("Terminal resize rows must be an integer from 5 to 200")
            return cols, rows

        try:
            first_message = await asyncio.wait_for(websocket.recv(), timeout=0.35)
            first_resize = _resize_from_message(first_message)
            if first_resize:
                initial_cols, initial_rows = first_resize
            else:
                pending_client_message = first_message
        except TimeoutError:
            pass
        except websockets.exceptions.ConnectionClosed:
            return
        except (UnicodeDecodeError, ValueError) as exc:
            with suppress(websockets.exceptions.ConnectionClosed):
                await websocket.send(f"\x1b[91mTerminal protocol failed: {exc}\x1b[0m\r\n")
                await websocket.close(code=1008, reason="Invalid terminal protocol")
            return

        try:
            shell_cmd = [ps_path, "-NoLogo", "-NoProfile", "-NoExit"]
            shell_cwd = str(Path.home())
            proc = await _spawn_conpty(shell_cmd, shell_cwd, (initial_rows, initial_cols))
            logging.debug("Terminal PTY started with ConPTY")
        except Exception as e:
            with suppress(websockets.exceptions.ConnectionClosed):
                await websocket.send(f"\x1b[91mLocal PTY failed: {e}\x1b[0m\r\n")
            return

        async def ws_to_proc():
            async def handle_client_message(msg):
                data = _message_to_text(msg)
                resize = _resize_from_message(msg)
                if resize:
                    cols, rows = resize
                    proc.set_size(rows, cols)
                    return

                await asyncio.to_thread(proc.write, data)

            if pending_client_message is not None:
                await handle_client_message(pending_client_message)
            async for msg in websocket:
                await handle_client_message(msg)

        async def proc_to_ws():
            while proc.isalive():
                try:
                    chunk = await asyncio.to_thread(proc.read, 4096)
                except EOFError:
                    return
                await websocket.send(chunk)

        sender = asyncio.create_task(ws_to_proc())
        receiver = asyncio.create_task(proc_to_ws())
        try:
            done, _ = await asyncio.wait({sender, receiver}, return_when=asyncio.FIRST_COMPLETED)
            for task in done:
                task.result()
        except websockets.exceptions.ConnectionClosed:
            pass
        except Exception as exc:
            with suppress(websockets.exceptions.ConnectionClosed):
                await websocket.send(f"\x1b[91mLocal PTY failed: {exc}\x1b[0m\r\n")
        finally:
            for t in (sender, receiver):
                if not t.done():
                    t.cancel()
            await asyncio.gather(sender, receiver, return_exceptions=True)
            try:
                proc.close()
            except Exception as exc:
                logging.error("Local PTY cleanup failed: %s", exc)

    async def broadcast_frames(self):
        """Broadcast new video frames to connected clients."""
        last_sequence = 0
        write_buffer_limit = 256_000
        send_tasks = self._video_send_tasks

        def cleanup_disconnected(disconnected_clients):
            for ws in disconnected_clients:
                self.video_clients.discard(ws)
                task = send_tasks.pop(ws, None)
                if task and not task.done():
                    task.cancel()

        while not self.stop_event.is_set():
            disconnected = set()
            for ws, task in list(send_tasks.items()):
                if not task.done():
                    continue
                send_tasks.pop(ws)
                try:
                    task.result()
                except websockets.exceptions.ConnectionClosed:
                    disconnected.add(ws)
                except Exception as exc:
                    logging.error("Frame send failed: %s", exc)
                    disconnected.add(ws)
            cleanup_disconnected(disconnected)

            try:
                await asyncio.wait_for(self.frame_ready_event.wait(), timeout=1.0)
                self.frame_ready_event.clear()
            except TimeoutError:
                pass

            sequence, frame = self.screen_capturer.get_latest_frame_packet()
            if frame is None or sequence == last_sequence:
                if not self.screen_capturer.is_running and self.screen_capturer.last_error:
                    raise RuntimeError(self.screen_capturer.last_error)
                continue
            if not frame:
                raise RuntimeError("Screen capturer produced an empty JPEG frame")
            last_sequence = sequence
            if not self.video_clients:
                continue

            max_write_buffer = 0
            for ws in self.video_clients:
                try:
                    buffer_size = ws.transport.get_write_buffer_size()
                    max_write_buffer = max(max_write_buffer, buffer_size)
                    if buffer_size > write_buffer_limit or ws in send_tasks:
                        self._video_skipped_sends += 1
                        continue
                    send_tasks[ws] = asyncio.create_task(ws.send(frame))
                except websockets.exceptions.ConnectionClosed:
                    disconnected.add(ws)
                except Exception as exc:
                    logging.error("Frame scheduling failed: %s", exc)
                    disconnected.add(ws)
            cleanup_disconnected(disconnected)

            now = time.time()
            controller = self.adaptive_stream
            if now - controller.last_eval_at >= 1.0:
                capture_stats = self.screen_capturer.get_capture_stats()
                frame_ts = float(capture_stats["last_frame_ts"])
                frame_age_ms = max(0.0, (now - frame_ts) * 1000.0) if frame_ts else 0.0
                if controller.observe_server(
                    frame_bytes=len(frame),
                    max_write_buffer=max_write_buffer,
                    skipped_total=self._video_skipped_sends,
                    inflight_sends=len(send_tasks),
                    video_clients=len(self.video_clients),
                    frame_age_ms=frame_age_ms,
                    capture_stats=capture_stats,
                ):
                    self._apply_stream_profile("adaptive_" + controller.last_reason)
                    await self._broadcast_stream_status()

        for task in send_tasks.values():
            if not task.done():
                task.cancel()
        await asyncio.gather(*send_tasks.values(), return_exceptions=True)
    async def broadcast_cursor_position(self):
        """Broadcast cursor position and button states to subscribed clients"""
        while not self.stop_event.is_set():
            if not (self.cursor_broadcast_enabled and self.cursor_subscribers):
                await asyncio.sleep(0.25)
                continue
            ci = CURSORINFO()
            ci.cbSize = ctypes.sizeof(CURSORINFO)
            if not user32.GetCursorInfo(ctypes.byref(ci)):
                raise ctypes.WinError()
            x, y = ci.ptScreenPos.x, ci.ptScreenPos.y

            vx = user32.GetSystemMetrics(SM_XVIRTUALSCREEN)
            vy = user32.GetSystemMetrics(SM_YVIRTUALSCREEN)
            vw = user32.GetSystemMetrics(SM_CXVIRTUALSCREEN)
            vh = user32.GetSystemMetrics(SM_CYVIRTUALSCREEN)
            if vw < 2 or vh < 2:
                raise RuntimeError(f"Invalid virtual screen dimensions: {vw}x{vh}")
            norm_x = (x - vx) / (vw - 1)
            norm_y = (y - vy) / (vh - 1)
            cursor_visible = bool(ci.flags & CURSOR_SHOWING)
            cursor_css = _get_css_cursor(ci.hCursor) if cursor_visible else "default"
            region = self.screen_capturer.get_active_region_norm()
            if region:
                left, top, right, bottom = region
                cursor_visible &= left <= norm_x <= right and top <= norm_y <= bottom
                norm_x = max(0.0, min(1.0, (norm_x - left) / (right - left)))
                norm_y = max(0.0, min(1.0, (norm_y - top) / (bottom - top)))

            message = json.dumps({
                'type': 'cursor',
                'x': norm_x,
                'y': norm_y,
                'cursor_visible': cursor_visible,
                'left_pressed': bool(_GetAsyncKeyState(VK_LBUTTON) & 0x8000),
                'right_pressed': bool(_GetAsyncKeyState(VK_RBUTTON) & 0x8000),
                'cursor_css': cursor_css,
            })
            disconnected = []
            for ws in self.cursor_subscribers.copy():
                try:
                    await ws.send(message)
                except websockets.exceptions.ConnectionClosed:
                    disconnected.append(ws)
                except Exception as exc:
                    logging.error("Cursor send failed: %s", exc)
                    disconnected.append(ws)

            self.cursor_subscribers.difference_update(disconnected)
            if not self.cursor_subscribers:
                self.cursor_broadcast_enabled = False

            # Update at 60 FPS for smooth cursor tracking while subscribed.
            await asyncio.sleep(1/60)
