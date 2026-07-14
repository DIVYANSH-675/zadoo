"""Network helpers."""
from __future__ import annotations

import socket


def get_local_ip() -> str:
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
            sock.connect(("8.8.8.8", 80))
            return sock.getsockname()[0]
    except OSError as exc:
        raise RuntimeError(f"Unable to determine the local IP address: {exc}") from exc
