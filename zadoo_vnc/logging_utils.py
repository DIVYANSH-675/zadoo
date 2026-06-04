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
_ORIGINAL_STDOUT = None
_ORIGINAL_STDERR = None
_LOGGING_RESTORE_REGISTERED = False


def _restore_logging_streams():
    global _LOG_FILE_HANDLE
    try:
        if _ORIGINAL_STDOUT is not None:
            sys.stdout = _ORIGINAL_STDOUT
        if _ORIGINAL_STDERR is not None:
            sys.stderr = _ORIGINAL_STDERR
    except Exception:
        pass
    try:
        if _LOG_FILE_HANDLE is not None and not _LOG_FILE_HANDLE.closed:
            _LOG_FILE_HANDLE.close()
    except Exception:
        pass
    _LOG_FILE_HANDLE = None


def _setup_logging_to_file():
    global _LOG_FILE_HANDLE, _ORIGINAL_STDOUT, _ORIGINAL_STDERR, _LOGGING_RESTORE_REGISTERED
    try:
        if getattr(sys, "frozen", False):
            try:
                from .settings import settings_dir

                log_dir = str(settings_dir() / "logs")
            except Exception:
                log_dir = os.path.join(os.path.dirname(sys.executable), "logs")
        else:
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            log_dir = os.path.join(base_dir, "logs")
        os.makedirs(log_dir, exist_ok=True)
        log_path = os.path.join(log_dir, f"zadoo_{datetime.now():%Y%m%d}.log")

        class _Tee(io.TextIOBase):
            def __init__(self, stream, file):
                self._stream = stream
                self._file = file
                self._lock = threading.Lock()

            @property
            def encoding(self):
                return getattr(self._stream, "encoding", "utf-8")

            @property
            def errors(self):
                return getattr(self._stream, "errors", "replace")

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

            def fileno(self):
                return self._stream.fileno()

            def isatty(self):
                try:
                    return bool(self._stream.isatty())
                except Exception:
                    return False

            def readable(self):
                return False

            def writable(self):
                return True

            def seekable(self):
                return False

            def __getattr__(self, name):
                return getattr(self._stream, name)

        fh = open(log_path, "a", encoding="utf-8", buffering=1)
        if _ORIGINAL_STDOUT is None:
            _ORIGINAL_STDOUT = sys.stdout
        if _ORIGINAL_STDERR is None:
            _ORIGINAL_STDERR = sys.stderr
        if _LOG_FILE_HANDLE is not None and not _LOG_FILE_HANDLE.closed:
            try:
                _LOG_FILE_HANDLE.close()
            except Exception:
                pass
        sys.stdout = _Tee(_ORIGINAL_STDOUT, fh)
        sys.stderr = _Tee(_ORIGINAL_STDERR, fh)
        _LOG_FILE_HANDLE = fh
        if not _LOGGING_RESTORE_REGISTERED:
            atexit.register(_restore_logging_streams)
            _LOGGING_RESTORE_REGISTERED = True
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


def _log_fallback(source: str, fallback: str, reason: str = "", exc: Exception | None = None):
    """Log every intentional runtime fallback with a stable searchable marker."""
    try:
        message = "FALLBACK_USED source=%s fallback=%s"
        args = [source, fallback]
        if reason:
            message += " reason=%s"
            args.append(reason)
        logging.getLogger("fallback").warning(message, *args, exc_info=exc is not None)
    except Exception:
        pass
