"""Logging and tracing helpers for the Zadoo VNC runtime."""
from __future__ import annotations

import atexit
import functools
import io
import logging
import os
import sys
import threading
import time
from datetime import datetime

_LOG_FILE_HANDLE = None

def _setup_logging_to_file():
    global _LOG_FILE_HANDLE
    try:
        base_dir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        log_dir = os.path.join(base_dir, 'logs')
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

        # Open file in line-buffered append mode
        fh = open(log_path, 'a', encoding='utf-8', buffering=1)
        sys.stdout = _Tee(sys.stdout, fh)
        sys.stderr = _Tee(sys.stderr, fh)
        _LOG_FILE_HANDLE = fh
        atexit.register(lambda: (_LOG_FILE_HANDLE and _LOG_FILE_HANDLE.close()))
        print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] 📜 Logging to {log_path}")
    except Exception as e:
        try:
            print(f"⚠️ Logging setup failed: {e}")
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


def log_calls(name: str = None, level: int = logging.INFO):
    """Decorator to log function entry/exit, duration, and exceptions (sync/async)."""
    def _decorator(func):
        label = name or getattr(func, '__qualname__', getattr(func, '__name__', 'func'))
        try:
            import asyncio as _asyncio
        except Exception:
            _asyncio = None

        def _safe(v, limit=200):
            try:
                s = str(v)
            except Exception:
                s = object.__repr__(v)
            return (s if len(s) <= limit else (s[:limit] + '…'))

        if _asyncio and _asyncio.iscoroutinefunction(func):
            async def _aw(*args, **kwargs):
                t0 = time.perf_counter()
                try:
                    logging.log(level, f"[CALL] {label} args={_safe(args[:1])} kwargs_keys={list(kwargs.keys())}")
                except Exception:
                    pass
                try:
                    result = await func(*args, **kwargs)
                    try:
                        dt = (time.perf_counter() - t0) * 1000.0
                        logging.log(level, f"[RETURN] {label} dt_ms={dt:.2f}")
                    except Exception:
                        pass
                    return result
                except Exception as e:
                    logging.error(f"[EXC] {label}", exc_info=True)
                    raise
            return functools.wraps(func)(_aw)
        else:
            def _w(*args, **kwargs):
                t0 = time.perf_counter()
                try:
                    logging.log(level, f"[CALL] {label} args={_safe(args[:1])} kwargs_keys={list(kwargs.keys())}")
                except Exception:
                    pass
                try:
                    result = func(*args, **kwargs)
                    try:
                        dt = (time.perf_counter() - t0) * 1000.0
                        logging.log(level, f"[RETURN] {label} dt_ms={dt:.2f}")
                    except Exception:
                        pass
                    return result
                except Exception as e:
                    logging.error(f"[EXC] {label}", exc_info=True)
                    raise
            return functools.wraps(func)(_w)
    return _decorator


def _instrument_class_methods(cls):
    """Wrap all methods of a class with log_calls, preserving static/class methods."""
    try:
        import inspect as _inspect
    except Exception:
        _inspect = None
    for _n, _attr in list(getattr(cls, '__dict__', {}).items()):
        if _n.startswith('__'):
            continue
        try:
            is_static = isinstance(_attr, staticmethod)
            is_class = isinstance(_attr, classmethod)
            func = _attr.__func__ if (is_static or is_class) else _attr
            if not callable(func):
                continue
            # Avoid wrapping properties or already-wrapped functions
            if _inspect and (isinstance(func, property) or getattr(func, '__wrapped__', None)):
                continue
            wrapped = log_calls(f"{cls.__name__}.{_n}")(func)
            if is_static:
                setattr(cls, _n, staticmethod(wrapped))
            elif is_class:
                setattr(cls, _n, classmethod(wrapped))
            else:
                setattr(cls, _n, wrapped)
        except Exception:
            # Best-effort; skip problematic attributes
            continue


def _make_line_tracer(include_files: tuple):
    """Return a sys.settrace function that logs each executed line for given files."""
    include_files = tuple(os.path.abspath(p) for p in include_files)

    def _tracer(frame, event, arg):
        try:
            fpath = os.path.abspath(frame.f_code.co_filename)
            if fpath not in include_files:
                return _tracer

            fname = frame.f_code.co_name
            lineno = frame.f_lineno

            if event == 'call':
                # Log function call with arguments
                try:
                    arg_names = frame.f_code.co_varnames[: frame.f_code.co_argcount]
                    local_map = frame.f_locals or {}
                    parts = []
                    for name in arg_names:
                        try:
                            val = local_map.get(name, '<unset>')
                            sval = repr(val)
                            if len(sval) > 200:
                                sval = sval[:200] + '…'
                            parts.append(f"{name}={sval}")
                        except Exception:
                            parts.append(f"{name}=<error>")
                    arg_str = ", ".join(parts)
                    logging.info("[trace.call] %s:%d in %s(%s)", fpath, lineno, fname, arg_str)
                    print(f"[trace.call] {fpath}:{lineno} in {fname}({arg_str})", flush=True)
                except Exception:
                    pass

            elif event == 'line':
                try:
                    # Attempt to fetch the current line text
                    with open(fpath, 'r', encoding='utf-8', errors='ignore') as f:
                        for i, line in enumerate(f, start=1):
                            if i == lineno:
                                code_line = line.rstrip('\n')
                                break
                        else:
                            code_line = ''
                except Exception:
                    code_line = ''
                try:
                    logging.info("[trace] %s:%d in %s | %s", fpath, lineno, fname, code_line)
                except Exception:
                    pass
                try:
                    # Tag control-flow lines explicitly
                    _sl = code_line.lstrip()
                    _tag = None
                    if _sl.startswith('try:'):
                        _tag = 'TRY'
                    elif _sl.startswith('except'):
                        _tag = 'EXCEPT'
                    elif _sl.startswith('finally:'):
                        _tag = 'FINALLY'
                    elif _sl.startswith('if '):
                        _tag = 'IF'
                    elif _sl.startswith('elif '):
                        _tag = 'ELIF'
                    elif _sl == 'else:':
                        _tag = 'ELSE'
                    if _tag:
                        print(f"{_tag} line {lineno} - {code_line}", flush=True)
                    # Minimal per-line progress format as requested
                    print(f"line {lineno} done - {code_line}", flush=True)
                except Exception:
                    pass

            elif event == 'return':
                try:
                    r = repr(arg)
                    if len(r) > 200:
                        r = r[:200] + '…'
                    logging.info("[trace.return] %s:%d in %s → %s", fpath, lineno, fname, r)
                    print(f"[trace.return] {fpath}:{lineno} in {fname} → {r}", flush=True)
                except Exception:
                    pass

            elif event == 'exception':
                try:
                    etype, evalue, _ = arg or (None, None, None)
                    ename = getattr(etype, '__name__', str(etype))
                    evals = repr(evalue)
                    if len(evals) > 200:
                        evals = evals[:200] + '…'
                    logging.info("[trace.exc] %s:%d in %s !! %s: %s", fpath, lineno, fname, ename, evals)
                    print(f"[trace.exc] {fpath}:{lineno} in {fname} !! {ename}: {evals}", flush=True)
                    # Plain exception tag line
                    print(f"EXCEPT raised at line {lineno} - {ename}: {evals}", flush=True)
                except Exception:
                    pass
        except Exception:
            pass
        return _tracer

    return _tracer
