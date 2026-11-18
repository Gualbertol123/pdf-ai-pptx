"""OCR engines on a page image: ONNX PP-OCR (det + cls + rec) and Windows OCR.

Both return :class:`OcrLine` objects in image pixels of the (deskewed) image:
the line box, its text, a confidence and, per character, the x range it covers
(from the CTC time steps for PP-OCR, from word boxes for Windows OCR). Turning
lines into IR glyphs is ``ml/ocr.py``'s job.
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field

import cv2
import numpy as np

from .. import log
from . import models, runtime

LOG = log.get(__name__)

SKEW_SIDE = 1600  # px: longer side of the image the skew is measured on
SKEW_STRIPS = 4
DET_MAX_SIDE = 1600  # px: the det input's longer side (multiple of 32)
DET_THRESH = 0.3
DET_BOX_THRESH = 0.55
DET_UNCLIP = 1.6
REC_H = 48
REC_BATCH = 8  # measured: 1.05 s vs 1.16 s at 16 for a dense page on DirectML
REC_WIDTH_STEP = 64  # rec input widths are padded to a multiple (fewer DirectML shapes)
CLS_FLIP_CONF = 0.9
MULTI_RETRY_CONF = 0.6  # lines the Latin model reads below this go to the multilingual one


@dataclass
class OcrChar:
    ch: str
    x0: float  # image px, along the line
    x1: float
    conf: float


@dataclass
class OcrLine:
    box: tuple[float, float, float, float]  # x0, y0, x1, y1 (image px)
    text: str
    conf: float
    chars: list[OcrChar] = field(default_factory=list)
    engine: str = "ppocr"
    flipped: bool = False


# --------------------------------------------------------------------------- deskew


def estimate_skew(gray: np.ndarray, max_deg: float = 5.0) -> float:
    """Skew of the text lines (degrees, counter-clockwise on screen is positive).

    The ink is rotated by candidate angles and the row profile is scored by its
    variance, per vertical strip: columns whose lines do not line up (a two-column
    report) blur one whole-page profile and pulled the estimate 0.3° off on s02."""
    h, w = gray.shape[:2]
    scale = min(1.0, SKEW_SIDE / max(h, w))
    small = cv2.resize(gray, (max(1, int(w * scale)), max(1, int(h * scale))),
                       interpolation=cv2.INTER_AREA) if scale < 1.0 else gray
    _, ink = cv2.threshold(small, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)
    ink = cv2.medianBlur(ink, 3)
    ys, xs = np.nonzero(ink)
    if len(xs) < 50:
        return 0.0
    sh, sw = small.shape[:2]
    cx, cy = sw / 2.0, sh / 2.0
    xs_c, ys_c = (xs - cx).astype(np.float64), (ys - cy).astype(np.float64)
    strip = np.minimum((xs * SKEW_STRIPS) // max(sw, 1), SKEW_STRIPS - 1).astype(np.int64)
    span = math.ceil(math.hypot(sh, sw)) + 4

    kernel = cv2.getGaussianKernel(9, 1.5).ravel()

    def score(deg: float) -> float:
        # soft binning + smoothing: integer bins favour exactly 0° (each pixel row lands
        # in one bin there, in two at any other angle)
        r = math.radians(deg)
        f = ys_c * math.cos(r) + xs_c * math.sin(r) + span // 2
        i0 = np.floor(f).astype(np.int64)
        wt = f - i0
        idx = strip * span + i0
        n = SKEW_STRIPS * span + 1
        hist = (np.bincount(idx, weights=1.0 - wt, minlength=n)
                + np.bincount(idx + 1, weights=wt, minlength=n))
        hist = np.convolve(hist, kernel, mode="same")
        return float(np.sum(hist ** 2))

    best = max(np.arange(-max_deg, max_deg + 1e-6, 0.25), key=score)
    for step, reach in ((0.05, 0.3), (0.01, 0.06)):
        cand = np.arange(best - reach, best + reach + 1e-9, step)
        best = max(cand, key=score)
    return round(float(best), 2)


def rotate_image(img: np.ndarray, deg: float, border: int | tuple[int, ...]) -> np.ndarray:
    """Rotate about the centre by ``deg`` (counter-clockwise on screen), same size."""
    h, w = img.shape[:2]
    m = cv2.getRotationMatrix2D((w / 2.0, h / 2.0), deg, 1.0)
    return cv2.warpAffine(img, m, (w, h), flags=cv2.INTER_CUBIC,
                          borderMode=cv2.BORDER_CONSTANT, borderValue=border)


# --------------------------------------------------------------------------- det


def _det_warmup(s: runtime.Session) -> None:
    s.run({s.input_names[0]: np.zeros((1, 3, 64, 64), np.float32)})


def _rec_warmup(s: runtime.Session) -> None:
    s.run({s.input_names[0]: np.zeros((1, 3, REC_H, 128), np.float32)})


def _cls_warmup(s: runtime.Session) -> None:
    s.run({s.input_names[0]: np.zeros((1, 3, 48, 192), np.float32)})


def detect(img: np.ndarray) -> list[tuple[float, float, float, float, float]]:
    """Text line boxes (x0, y0, x1, y1, score) of a BGR image, axis aligned (the
    image is deskewed first)."""
    s = runtime.session("ocr_det", _det_warmup)
    if s is None:
        raise RuntimeError("ocr_det model unavailable")
    h, w = img.shape[:2]
    scale = min(1.0, DET_MAX_SIDE / max(h, w))
    nh = max(32, round(h * scale / 32) * 32)
    nw = max(32, round(w * scale / 32) * 32)
    x = cv2.resize(img, (nw, nh), interpolation=cv2.INTER_AREA).astype(np.float32) / 255.0
    x = (x - np.array([0.485, 0.456, 0.406], np.float32)) / np.array([0.229, 0.224, 0.225],
                                                                       np.float32)
    prob = s.run({s.input_names[0]: x.transpose(2, 0, 1)[None]})[0][0, 0]
    bitmap = (prob > DET_THRESH).astype(np.uint8)
    contours, _ = cv2.findContours(bitmap, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
    sx, sy = w / nw, h / nh
    out = []
    for c in contours:
        if len(c) < 4:
            continue
        x0, y0, bw, bh = cv2.boundingRect(c)
        if min(bw, bh) < 3:
            continue
        mask = np.zeros((bh, bw), np.uint8)
        cv2.fillPoly(mask, [c.reshape(-1, 2) - (x0, y0)], 1)
        score = float(cv2.mean(prob[y0:y0 + bh, x0:x0 + bw], mask)[0])
        if score < DET_BOX_THRESH:
            continue
        area, peri = float(cv2.contourArea(c)), float(cv2.arcLength(c, True))
        d = area * DET_UNCLIP / max(peri, 1.0)
        out.append(((x0 - d) * sx, (y0 - d) * sy, (x0 + bw + d) * sx, (y0 + bh + d) * sy,
                    score))
    out = [(max(0.0, a), max(0.0, b), min(float(w), c), min(float(h), d), sc)
           for a, b, c, d, sc in out]
    out.sort(key=lambda b: (round(b[1]), b[0]))
    return out


# --------------------------------------------------------------------------- rec


@dataclass
class _Rec:
    session: str
    chars: list[str]


def _latin() -> _Rec | None:
    p = models.path("ocr_rec_latin_dict")
    if p is None or models.path("ocr_rec_latin") is None:
        return None
    chars = p.read_text(encoding="utf-8").splitlines()
    return _Rec("ocr_rec_latin", ["", *chars, " "])


def _multi() -> _Rec | None:
    s = runtime.session("ocr_rec_multi", _rec_warmup)
    if s is None:
        return None
    chars = s.metadata().get("character", "").splitlines()
    return _Rec("ocr_rec_multi", ["", *chars, " "]) if chars else None


def _crop(img: np.ndarray, box: tuple[float, float, float, float]) -> np.ndarray:
    x0, y0, x1, y1 = (round(v) for v in box)
    return img[max(0, y0):max(y0 + 1, y1), max(0, x0):max(x0 + 1, x1)]


def _rec_input(crops: list[np.ndarray]) -> tuple[np.ndarray, list[int]]:
    widths = [max(8, math.ceil(REC_H * c.shape[1] / max(c.shape[0], 1))) for c in crops]
    wmax = max(widths)
    wmax = int(math.ceil(wmax / REC_WIDTH_STEP) * REC_WIDTH_STEP)
    batch = np.zeros((len(crops), 3, REC_H, wmax), np.float32)
    for i, (c, w) in enumerate(zip(crops, widths, strict=True)):
        r = cv2.resize(c, (w, REC_H), interpolation=cv2.INTER_LINEAR).astype(np.float32)
        batch[i, :, :, :w] = ((r / 255.0 - 0.5) / 0.5).transpose(2, 0, 1)
    return batch, widths


def _ctc(prob: np.ndarray, chars: list[str], width: int, crop_w: int,
         padded_w: int) -> tuple[str, float, list[OcrChar]]:
    """Greedy CTC decode with the x range of each character (crop pixels)."""
    t_n = prob.shape[0]
    idx = prob.argmax(axis=1)
    conf = prob.max(axis=1)
    step = padded_w / t_n  # resized px per time step
    to_crop = crop_w / max(width, 1)
    out: list[OcrChar] = []
    prev = 0
    for t in range(t_n):
        k = int(idx[t])
        if k != 0 and k != prev and k < len(chars):
            x = t * step * to_crop
            out.append(OcrChar(chars[k], x, x + step * to_crop, float(conf[t])))
        elif k != 0 and k == prev and out:  # the same character continues
            out[-1].x1 = (t + 1) * step * to_crop
            out[-1].conf = max(out[-1].conf, float(conf[t]))
        prev = k
    text = "".join(c.ch for c in out)
    c = float(np.mean([o.conf for o in out])) if out else 0.0
    return text, c, out


def recognise(img: np.ndarray, boxes: list[tuple[float, float, float, float]],
              rec: _Rec) -> list[tuple[str, float, list[OcrChar]]]:
    s = runtime.session(rec.session, _rec_warmup)
    if s is None:
        raise RuntimeError(f"{rec.session} model unavailable")
    crops = [_crop(img, b) for b in boxes]
    order = sorted(range(len(crops)), key=lambda i: crops[i].shape[1] / max(crops[i].shape[0],
                                                                            1))
    res: list[tuple[str, float, list[OcrChar]]] = [("", 0.0, [])] * len(crops)
    for k in range(0, len(order), REC_BATCH):
        ids = order[k:k + REC_BATCH]
        batch, widths = _rec_input([crops[i] for i in ids])
        probs = s.run({s.input_names[0]: batch}, items=len(ids))[0]
        for j, i in enumerate(ids):
            res[i] = _ctc(probs[j], rec.chars, widths[j], crops[i].shape[1], batch.shape[3])
    return res


def orientation(img: np.ndarray, boxes: list[tuple[float, float, float, float]]) -> list[bool]:
    """True for lines the cls model sees upside down (confidently)."""
    s = runtime.session("ocr_cls", _cls_warmup)
    if s is None or not boxes:
        return [False] * len(boxes)
    out: list[bool] = []
    for k in range(0, len(boxes), REC_BATCH):
        crops = [_crop(img, b) for b in boxes[k:k + REC_BATCH]]
        batch = np.zeros((len(crops), 3, 48, 192), np.float32)
        for i, c in enumerate(crops):
            w = min(192, max(8, math.ceil(48 * c.shape[1] / max(c.shape[0], 1))))
            r = cv2.resize(c, (w, 48), interpolation=cv2.INTER_LINEAR).astype(np.float32)
            batch[i, :, :, :w] = ((r / 255.0 - 0.5) / 0.5).transpose(2, 0, 1)
        p = s.run({s.input_names[0]: batch}, items=len(crops))[0]
        out.extend(bool(row[1] > CLS_FLIP_CONF and row[1] > row[0]) for row in p)
    return out


def ppocr_available() -> bool:
    return models.available("ocr_det", "ocr_rec_latin", "ocr_rec_latin_dict")


def ppocr(img: np.ndarray, multilingual: bool = True) -> list[OcrLine]:
    """Detect and read every text line of a (deskewed) BGR image."""
    boxes = detect(img)
    if not boxes:
        return []
    rects = [b[:4] for b in boxes]
    latin = _latin()
    if latin is None:
        raise RuntimeError("Latin recognition model unavailable")
    res = recognise(img, rects, latin)
    # the orientation model is not trusted alone (it calls some upright lines upside
    # down): a flagged line is read both ways and the more confident reading wins
    flagged = [i for i, f in enumerate(orientation(img, rects)) if f]
    flips = [False] * len(rects)
    if flagged:
        work = img.copy()
        for i in flagged:
            x0, y0, x1, y1 = (round(v) for v in rects[i])
            work[y0:y1, x0:x1] = cv2.rotate(img[y0:y1, x0:x1], cv2.ROTATE_180)
        again = recognise(work, [rects[i] for i in flagged], latin)
        for i, r in zip(flagged, again, strict=True):
            if r[1] > res[i][1] + 0.05:
                res[i], flips[i] = r, True
    if multilingual:
        weak = [i for i, r in enumerate(res) if r[1] < MULTI_RETRY_CONF and not flips[i]]
        multi = _multi() if weak else None
        if multi is not None:
            again = recognise(img, [rects[i] for i in weak], multi)
            for i, r in zip(weak, again, strict=True):
                if r[1] > res[i][1] + 0.1:
                    res[i] = r
    lines = []
    for (x0, y0, x1, y1, _), (text, conf, chars), flip in zip(boxes, res, flips, strict=True):
        if not text.strip():
            continue
        if flip:  # read on the rotated crop: x runs right to left (reading order kept)
            w = x1 - x0
            chars = [OcrChar(c.ch, w - c.x1, w - c.x0, c.conf) for c in chars]
        for c in chars:
            c.x0 += x0
            c.x1 += x0
        lines.append(OcrLine(box=(x0, y0, x1, y1), text=text, conf=conf, chars=chars,
                             flipped=flip))
    return lines
