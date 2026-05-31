#!/usr/bin/env python3
"""Compatibility entrypoint for the modular Zadoo VNC runtime."""
import logging

from zadoo_vnc.app import main


if __name__ == "__main__":
    try:
        main()
    except BaseException:
        logging.exception("Fatal Zadoo startup error")
        raise
