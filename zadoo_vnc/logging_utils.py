"""Logging helpers for the Zadoo VNC runtime."""
from __future__ import annotations

import atexit
import io
import logging
import sys
import threading
import time
from datetime import datetime
from pathlib import Path

from .config import PROJECT_DIR, env_int

_LOG_FILE_HANDLE = None
_ORIGINAL_STDOUT = None
_ORIGINAL_STDERR = None
_LOGGING_RESTORE_REGISTERED = False


class _BoundedLogFile:
    encoding = "utf-8"
    errors = "strict"

    def __init__(self, path: Path, max_bytes: int, backup_count: int):
        self.path = path
        self.max_bytes = max_bytes
        self.backup_count = backup_count
        self._file = self.path.open("a", encoding=self.encoding, buffering=1)

    @property
    def closed(self):
        return self._file.closed

    def _backup_path(self, number: int) -> Path:
        return self.path.with_name(f"{self.path.name}.{number}")

    def _rotate(self):
        self._file.close()
        self._backup_path(self.backup_count).unlink(missing_ok=True)
        for number in range(self.backup_count - 1, 0, -1):
            source = self._backup_path(number)
            if source.exists():
                source.replace(self._backup_path(number + 1))
        if self.path.exists():
            self.path.replace(self._backup_path(1))
        self._file = self.path.open("a", encoding=self.encoding, buffering=1)

    def write(self, text):
        if not text:
            return 0
        encoded_size = len(text.encode(self.encoding, errors="replace"))
        self._file.seek(0, 2)
        if self._file.tell() and self._file.tell() + encoded_size > self.max_bytes:
            self._rotate()
        return self._file.write(text)

    def flush(self):
        self._file.flush()

    def fileno(self):
        return self._file.fileno()

    def close(self):
        self._file.close()


def _prune_logs(log_dir: Path, retention_days: int):
    cutoff = time.time() - retention_days * 86400
    for path in log_dir.glob("zadoo_*.log*"):
        try:
            if path.is_file() and path.stat().st_mtime < cutoff:
                path.unlink()
        except OSError as exc:
            logging.warning("Failed to prune old log %s: %s", path, exc)


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
    retention_days = env_int("ZADOO_LOG_RETENTION_DAYS", 14, minimum=1, maximum=365)
    max_bytes = env_int("ZADOO_LOG_MAX_BYTES", 5 * 1024 * 1024, minimum=65536, maximum=1024 * 1024 * 1024)
    backup_count = env_int("ZADOO_LOG_BACKUP_COUNT", 2, minimum=1, maximum=20)
    _prune_logs(log_dir, retention_days)
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

    fh = _BoundedLogFile(log_path, max_bytes, backup_count)
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
