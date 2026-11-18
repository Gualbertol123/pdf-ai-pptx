"""Windows built-in OCR (Windows.Media.Ocr through the ``winrt`` packages): the
fallback engine when the ONNX models are not installed.

It returns words with boxes but no confidence (0.8 is assumed) and no character
positions: a word's characters share its box in proportion to their count.
"""

from __future__ import annotations

import asyncio
import functools
import sys
from typing import Any

import cv2
import numpy as np

from .. import log
from .ocr_engine import OcrChar, OcrLine

LOG = log.get(__name__)

ASSUMED_CONF = 0.8


@functools.cache
def _engine() -> Any | None:
    if sys.platform != "win32":
        return None
    try:
        from winrt.windows.media.ocr import OcrEngine  # type: ignore[import-not-found]

        return OcrEngine.try_create_from_user_profile_languages()
    except Exception as exc:
        LOG.debug("Windows OCR unavailable: %s", exc)
        return None


def available() -> bool:
    return _engine() is not None


async def _recognise(img: np.ndarray) -> Any:
    from winrt.windows.graphics.imaging import (  # type: ignore[import-not-found]
        BitmapAlphaMode,
        BitmapPixelFormat,
        SoftwareBitmap,
    )
    from winrt.windows.storage.streams import DataWriter  # type: ignore[import-not-found]

    bgra = cv2.cvtColor(img, cv2.COLOR_BGR2BGRA)
    dw = DataWriter()
    dw.write_bytes(bgra.tobytes())
    bmp = SoftwareBitmap.create_copy_from_buffer(dw.detach_buffer(), BitmapPixelFormat.BGRA8,
                                                 img.shape[1], img.shape[0],
                                                 BitmapAlphaMode.PREMULTIPLIED)
    return await _engine().recognize_async(bmp)


def recognise(img: np.ndarray) -> list[OcrLine]:
    eng = _engine()
    if eng is None:
        raise RuntimeError("Windows OCR unavailable")
    from winrt.windows.media.ocr import OcrEngine  # type: ignore[import-not-found]

    scale = 1.0
    mx = int(OcrEngine.max_image_dimension)
    if max(img.shape[:2]) > mx:
        scale = mx / max(img.shape[:2])
        img = cv2.resize(img, (int(img.shape[1] * scale), int(img.shape[0] * scale)),
                         interpolation=cv2.INTER_AREA)
    res = asyncio.run(_recognise(img))
    out: list[OcrLine] = []
    for line in res.lines:
        chars: list[OcrChar] = []
        boxes = []
        for wi, word in enumerate(line.words):
            r = word.bounding_rect
            x0, y0 = r.x / scale, r.y / scale
            x1, y1 = (r.x + r.width) / scale, (r.y + r.height) / scale
            boxes.append((x0, y0, x1, y1))
            if wi:
                gap0 = chars[-1].x1 if chars else x0
                chars.append(OcrChar(" ", gap0, x0, ASSUMED_CONF))
            text = str(word.text)
            step = (x1 - x0) / max(len(text), 1)
            chars.extend(OcrChar(c, x0 + i * step, x0 + (i + 1) * step, ASSUMED_CONF)
                         for i, c in enumerate(text))
        if not boxes:
            continue
        box = (min(b[0] for b in boxes), min(b[1] for b in boxes),
               max(b[2] for b in boxes), max(b[3] for b in boxes))
        pad = 0.15 * (box[3] - box[1])
        box = (box[0] - pad, box[1] - pad, box[2] + pad, box[3] + pad)
        out.append(OcrLine(box=box, text="".join(c.ch for c in chars), conf=ASSUMED_CONF,
                           chars=chars, engine="winocr"))
    return out
