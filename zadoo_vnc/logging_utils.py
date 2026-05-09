"""Logging helpers for the Zadoo VNC runtime."""
from __future__ import annotations

import atexit
import io
import logging
import os
import sys
import threading
from datetime import datetime

_LOG_FILE_HANDLE = None


def _setup_logging_to_file():
    global _LOG_FILE_HANDLE
    try:
        base_dir = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        log_dir = os.path.join(base_dir, "logs")
        os.makedirs(log_dir, exist_ok=True)
        log_path = os.path.join(log_dir, f"zadoo_{datetime.now():%Y%m%d}.log")

        class _Tee(io.TextIOBase):
            def __init__(self, stream, file):
                self._stream = stream
                self._file = file
                self._lock = threading.Lock()

            def write(self, s):
                with self._lock:
                    try:
                        self._stream.write(s)
                    except Exception:
                        pass
                    try:
                        self._file.write(s)
                        self._file.flush()
                    except Exception:
                        pass
                return len(s)

            def flush(self):
                with self._lock:
                    try:
                        self._stream.flush()
                    except Exception:
                        pass
                    try:
                        self._file.flush()
                    except Exception:
                        pass

        fh = open(log_path, "a", encoding="utf-8", buffering=1)
        sys.stdout = _Tee(sys.stdout, fh)
        sys.stderr = _Tee(sys.stderr, fh)
        _LOG_FILE_HANDLE = fh
        atexit.register(lambda: (_LOG_FILE_HANDLE and _LOG_FILE_HANDLE.close()))
        print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] Logging to {log_path}")
    except Exception as e:
        try:
            print(f"Logging setup failed: {e}")
        except Exception:
            pass


def _log_try_ok(label: str, extra: str = ""):
    try:
        logging.debug("TRY_OK %s%s", label, (" " + extra) if extra else "")
    except Exception:
        pass


def _log_except(label: str, e: Exception):
    try:
        logging.warning("EXCEPT %s: %s", label, e)
    except Exception:
        pass
