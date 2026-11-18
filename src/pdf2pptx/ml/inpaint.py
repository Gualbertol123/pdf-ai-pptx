"""Clean backgrounds for scanned pages: mask the text, fill the holes.

Rules first (PLAN.md §8 rule 2): the mask is split into regions (nearby letters
together). A region whose surrounding ring is flat (paper, a solid band) is filled
with the ring's median colour plus noise of the ring's strength, which matches
paper texture and costs nothing. Only regions on a non-uniform background (photo,
gradient, pattern) go to LaMa (512 × 512 tiles; ~2 s each on the CPU, it does not
run on DirectML here). Without LaMa those fall back to OpenCV's Telea inpainting.
"""

from __future__ import annotations

import math
import time
from typing import TYPE_CHECKING, Any

import cv2
import numpy as np

from .. import log
from . import runtime

if TYPE_CHECKING:
    from .ocr import Frame, LineGeo

LOG = log.get(__name__)

LAMA_SIZE = 512
UNIFORM_STD = 16.0  # gray std of the ring around a region; above: not uniform
UNIFORM_SPREAD = 40.0  # 5th..95th percentile spread of the ring; above: not uniform
GROUP_PX = 0.006  # regions: mask dilated by this share of the page diagonal (~a word gap)
SEED = 1234


def text_mask(gray: np.ndarray, geos: list[LineGeo], fr: Frame,
              shape_hw: tuple[int, int]) -> np.ndarray:
    """Ink of the converted words (deskewed frame), dilated, mapped back to the
    original image. uint8, 255 = to fill."""
    from .ocr import MIN_WORD_CONF, _ink

    h, w = gray.shape
    m = np.zeros((h, w), np.uint8)
    for g in geos:
        pad = max(2, round(0.15 * g.x_height))
        y0 = max(0, math.floor(min(g.ink_top, g.line.box[1])) - pad)
        y1 = min(h, math.ceil(max(g.ink_bot, g.line.box[3])) + pad)
        for wd in g.words:
            if wd.conf < MIN_WORD_CONF or not wd.text:
                continue
            x0 = max(0, math.floor(wd.x0) - pad)
            x1 = min(w, math.ceil(wd.x1) + pad)
            if x1 - x0 < 2 or y1 - y0 < 2:
                continue
            ink, _ = _ink(gray[y0:y1, x0:x1])
            m[y0:y1, x0:x1][ink] = 255
        r = max(2, round(0.12 * g.x_height))
        x0 = max(0, math.floor(min([g.line.box[0], *(wd.x0 for wd in g.words)])) - r)
        x1 = min(w, math.ceil(g.line.box[2]) + r)
        k = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (2 * r + 1, 2 * r + 1))
        m[y0:y1, x0:x1] = cv2.dilate(m[y0:y1, x0:x1], k)
    if fr.skew:
        m = cv2.warpAffine(m, fr.minv, (shape_hw[1], shape_hw[0]), flags=cv2.INTER_NEAREST)
        m = cv2.dilate(m, np.ones((3, 3), np.uint8))
    return m


def _uniform(ring: np.ndarray) -> bool:
    if ring.size < 20:
        return True
    lo, hi = np.percentile(ring, [5, 95])
    return float(ring.std()) <= UNIFORM_STD and float(hi - lo) <= UNIFORM_SPREAD


def fill(img: np.ndarray, mask: np.ndarray) -> tuple[np.ndarray, dict[str, Any]]:
    """``img`` with the masked pixels filled; and what filled them."""
    out = img.copy()
    stats: dict[str, Any] = {"regions": 0, "uniform": 0, "lama": 0, "telea": 0,
                             "masked_share": round(float((mask > 0).mean()), 4)}
    if not mask.any():
        return out, stats
    h, w = mask.shape
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    grow = max(3, int(GROUP_PX * math.hypot(h, w)))
    groups = cv2.dilate(mask, cv2.getStructuringElement(cv2.MORPH_RECT, (grow, grow)))
    n, lab, st, _ = cv2.connectedComponentsWithStats((groups > 0).astype(np.uint8), 8)
    rng = np.random.default_rng(SEED)
    hard: list[tuple[int, int, int, int]] = []
    ring_r = max(4, grow // 2)
    for i in range(1, n):
        x, y, bw, bh, _ = (int(v) for v in st[i])
        x0, y0 = max(0, x - ring_r), max(0, y - ring_r)
        x1, y1 = min(w, x + bw + ring_r), min(h, y + bh + ring_r)
        reg = (lab[y0:y1, x0:x1] == i) & (mask[y0:y1, x0:x1] > 0)
        if not reg.any():
            continue
        stats["regions"] += 1
        near = cv2.dilate(reg.astype(np.uint8), np.ones((2 * ring_r + 1,) * 2, np.uint8)) > 0
        ring_m = near & (mask[y0:y1, x0:x1] == 0)
        ring = gray[y0:y1, x0:x1][ring_m]
        if _uniform(ring):
            px = img[y0:y1, x0:x1][ring_m].astype(np.float32)
            med = np.median(px, axis=0) if len(px) else np.full(3, 255.0)
            sd = float(np.std(ring)) * 0.8 if ring.size else 0.0
            k = int(reg.sum())
            noise = rng.normal(0.0, sd, size=(k, 1)) if sd > 0 else np.zeros((k, 1))
            vals = np.clip(med[None, :] + noise, 0, 255).astype(np.uint8)
            out[y0:y1, x0:x1][reg] = vals
            stats["uniform"] += 1
        else:
            LOG.debug("non-uniform ring around %s: std %.1f, p5-p95 %s", (x0, y0, x1, y1),
                      float(ring.std()), np.percentile(ring, [5, 95]).tolist())
            hard.append((x0, y0, x1, y1))
    if hard:
        t0 = time.perf_counter()
        sess = runtime.session("inpaint", _lama_warmup)
        for box in hard:
            if sess is not None:
                _lama(sess, out, mask, box)
                stats["lama"] += 1
            else:
                x0, y0, x1, y1 = box
                out[y0:y1, x0:x1] = cv2.inpaint(out[y0:y1, x0:x1], mask[y0:y1, x0:x1], 5,
                                                cv2.INPAINT_TELEA)
                stats["telea"] += 1
        stats["hard_seconds"] = round(time.perf_counter() - t0, 3)
        if sess is not None:
            stats["lama_provider"] = sess.provider
    return out, stats


def _lama_warmup(s: runtime.Session) -> None:
    """A real hole: an all-zero mask lets a provider skip work that fails later."""
    z = np.full((1, 3, LAMA_SIZE, LAMA_SIZE), 0.5, np.float32)
    m = np.zeros((1, 1, LAMA_SIZE, LAMA_SIZE), np.float32)
    m[..., 200:300, 100:400] = 1.0
    s.run({"image": z, "mask": m})


def _lama(sess: runtime.Session, out: np.ndarray, mask: np.ndarray,
          box: tuple[int, int, int, int]) -> None:
    """Inpaint one region with LaMa in place: a square window around it (native
    resolution when it fits in 512 px, else scaled), only masked pixels change."""
    h, w = mask.shape
    x0, y0, x1, y1 = box
    extent = max(x1 - x0, y1 - y0)
    side = LAMA_SIZE if extent <= 0.8 * LAMA_SIZE else int(extent * 1.25)
    cx, cy = (x0 + x1) // 2, (y0 + y1) // 2
    wx0 = int(min(max(cx - side // 2, 0), max(w - side, 0)))
    wy0 = int(min(max(cy - side // 2, 0), max(h - side, 0)))
    wx1, wy1 = min(w, wx0 + side), min(h, wy0 + side)
    win = out[wy0:wy1, wx0:wx1]
    m = mask[wy0:wy1, wx0:wx1]
    ph, pw = win.shape[:2]
    im = cv2.resize(win, (LAMA_SIZE, LAMA_SIZE), interpolation=cv2.INTER_AREA) if (
        ph, pw) != (LAMA_SIZE, LAMA_SIZE) else win
    mm = cv2.resize(m, (LAMA_SIZE, LAMA_SIZE), interpolation=cv2.INTER_NEAREST) if (
        ph, pw) != (LAMA_SIZE, LAMA_SIZE) else m
    x = cv2.cvtColor(im, cv2.COLOR_BGR2RGB).astype(np.float32).transpose(2, 0, 1)[None] / 255.0
    mk = (mm > 0).astype(np.float32)[None, None]
    res = sess.run({"image": x, "mask": mk})[0][0].transpose(1, 2, 0)
    if float(res.max()) <= 2.0:  # some exports return 0..1
        res = res * 255.0
    res = cv2.cvtColor(np.clip(res, 0, 255).astype(np.uint8), cv2.COLOR_RGB2BGR)
    if (ph, pw) != (LAMA_SIZE, LAMA_SIZE):
        res = cv2.resize(res, (pw, ph), interpolation=cv2.INTER_CUBIC)
    sel = m > 0
    win[sel] = res[sel]
