"""Logging helpers for the Zadoo VNC runtime."""
from __future__ import annotations

import atexit
import io
import logging
import sys
import threading
from datetime import datetime

from .config import PROJECT_DIR

_LOG_FILE_HANDLE = None
_ORIGINAL_STDOUT = None
_ORIGINAL_STDERR = None
_LOGGING_RESTORE_REGISTERED = False


def _restore_logging_streams():
    global _LOG_FILE_HANDLE
    if _ORIGINAL_STDOUT is not None:
        sys.stdout = _ORIGINAL_STDOUT
    if _ORIGINAL_STDERR is not None:
        sys.stderr = _ORIGINAL_STDERR
    if _LOG_FILE_HANDLE is not None and not _LOG_FILE_HANDLE.closed:
        _LOG_FILE_HANDLE.close()
    _LOG_FILE_HANDLE = None


def _setup_logging_to_file():
    global _LOG_FILE_HANDLE, _ORIGINAL_STDOUT, _ORIGINAL_STDERR, _LOGGING_RESTORE_REGISTERED
    if getattr(sys, "frozen", False):
        from .settings import settings_dir

        log_dir = settings_dir() / "logs"
    else:
        log_dir = PROJECT_DIR / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / f"zadoo_{datetime.now():%Y%m%d}.log"

    class _Tee(io.TextIOBase):
        def __init__(self, stream, file):
            self._stream = stream
            self._file = file
            self._lock = threading.Lock()

        @property
        def encoding(self):
            return self._stream.encoding if self._stream is not None else self._file.encoding

        @property
        def errors(self):
            return self._stream.errors if self._stream is not None else self._file.errors

        def write(self, s):
            with self._lock:
                self._file.write(s)
                if self._stream is not None:
                    self._stream.write(s)
            return len(s)

        def flush(self):
            with self._lock:
                self._file.flush()
                if self._stream is not None:
                    self._stream.flush()

        def fileno(self):
            return (self._stream or self._file).fileno()

        def isatty(self):
            return bool(self._stream and self._stream.isatty())

        def readable(self):
            return False

        def writable(self):
            return True

        def seekable(self):
            return False

        def __getattr__(self, name):
            return getattr(self._stream if self._stream is not None else self._file, name)

    fh = log_path.open("a", encoding="utf-8", buffering=1)
    if _ORIGINAL_STDOUT is None:
        _ORIGINAL_STDOUT = sys.stdout
    if _ORIGINAL_STDERR is None:
        _ORIGINAL_STDERR = sys.stderr
    if _LOG_FILE_HANDLE is not None and not _LOG_FILE_HANDLE.closed:
        _LOG_FILE_HANDLE.close()
    sys.stdout = _Tee(_ORIGINAL_STDOUT, fh)
    sys.stderr = _Tee(_ORIGINAL_STDERR, fh)
    _LOG_FILE_HANDLE = fh
    if not _LOGGING_RESTORE_REGISTERED:
        atexit.register(_restore_logging_streams)
        _LOGGING_RESTORE_REGISTERED = True
    print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] Logging to {log_path}")


def _log_except(label: str, e: Exception):
    logging.warning("%s failed: %s", label, e)
