"""Bounded, independently recoverable JPEG desktop updates.

The optional wire format is a 28-byte network-endian header followed by JPEG:
magic (ZDF1 full / ZDP1 patch), sequence, base sequence, canvas width/height, x/y.
Legacy viewers continue receiving ordinary, independent JPEG images.
"""
from __future__ import annotations

import struct
import time

import numpy as np

FRAME_HEADER = struct.Struct("!4sIIIIII")
FRAME_MAGICS = (b"ZDF1", b"ZDP1")


class DeltaFrameEncoder:
    def __init__(self):
        self.previous = None
        self.signature = None
        self.sequence = 0
        self.last_keyframe_at = 0.0
        self.full_frames = 0
        self.patch_frames = 0
        self.unchanged_frames = 0
        self.last_update_pixels = 0
        self.detail_refresh_pending = False

    def reset(self):
        self.previous = None
        self.signature = None
        self.detail_refresh_pending = False

    def encode(self, pixels, encode_jpeg, signature, *, force_keyframe=False, take_ownership=False, encode_motion_jpeg=None):
        now = time.monotonic()
        height, width = pixels.shape[:2]
        full = (force_keyframe or self.previous is None or self.signature != signature
                or self.previous.shape != pixels.shape or now - self.last_keyframe_at >= 2.0)
        left = top = 0
        right, bottom = width, height
        large_change = False
        if not full:
            # Compare one packed BGRA value per pixel, rather than constructing
            # a four-channel boolean image or scanning Python-level tiles.
            if pixels.ndim == 3 and pixels.shape[2] == 4:
                changed = pixels.view(np.uint32).reshape(height, width) != self.previous.view(np.uint32).reshape(height, width)
            elif pixels.ndim == 2:
                changed = pixels != self.previous
            else:
                changed = np.any(pixels != self.previous, axis=2)
            rows = np.flatnonzero(np.any(changed, axis=1))
            if not rows.size:
                self.unchanged_frames += 1
                return None
            top, bottom = int(rows[0]), int(rows[-1]) + 1
            columns = np.flatnonzero(np.any(changed[top:bottom], axis=0))
            left, right = int(columns[0]), int(columns[-1]) + 1
            # Align to JPEG MCUs and include their neighbours. This avoids
            # accumulating edge/chroma seams when an object moves repeatedly.
            left, top = max(0, (left//16-1)*16), max(0, (top//16-1)*16)
            right, bottom = min(width, ((right+15)//16+1)*16), min(height, ((bottom+15)//16+1)*16)
            full = (right-left)*(bottom-top) >= width*height*.65
            large_change = full
            if full:
                left = top = 0
                right, bottom = width, height
        motion_encoding = large_change and encode_motion_jpeg is not None
        encode = encode_motion_jpeg if motion_encoding else encode_jpeg
        jpeg = encode(pixels[top:bottom, left:right])
        if not isinstance(jpeg, bytes) or not jpeg:
            raise RuntimeError("JPEG encoder returned an empty or non-bytes frame")
        base = 0 if full else self.sequence
        self.sequence = (self.sequence + 1) % 0x100000000
        if full:
            self.last_keyframe_at = now
            self.full_frames += 1
            self.detail_refresh_pending = motion_encoding
        else:
            self.patch_frames += 1
        self.last_update_pixels = (right-left)*(bottom-top)
        # Capture already gives the encoder a fresh, privately owned frame.
        # Retain it directly instead of copying the entire desktop a second time.
        self.previous = pixels if take_ownership else pixels.copy()
        self.signature = signature
        return FRAME_HEADER.pack(b"ZDF1" if full else b"ZDP1", self.sequence, base, width, height, left, top) + jpeg
