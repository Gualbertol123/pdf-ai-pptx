"""Deterministic images used by the synthetic corpus."""

from __future__ import annotations

from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw


def photo(path: Path, w: int = 640, h: int = 420, seed: int = 7) -> Path:
    """A photo-like image: smooth gradients, soft blobs, a little noise."""
    rng = np.random.default_rng(seed)
    yy, xx = np.mgrid[0:h, 0:w].astype(np.float64)
    r = 90 + 120 * (xx / w)
    g = 120 + 80 * (yy / h)
    b = 200 - 100 * (xx / w) * (yy / h)
    img = np.stack([r, g, b], axis=-1)
    for _ in range(6):
        cx, cy = rng.uniform(0, w), rng.uniform(0, h)
        rad = rng.uniform(40, 140)
        col = rng.uniform(0, 255, size=3)
        mask = np.exp(-(((xx - cx) ** 2 + (yy - cy) ** 2) / (2 * rad**2)))[..., None]
        img = img * (1 - 0.6 * mask) + col * 0.6 * mask
    img += rng.normal(0, 4, size=img.shape)
    arr = np.clip(img, 0, 255).astype(np.uint8)
    path.parent.mkdir(parents=True, exist_ok=True)
    Image.fromarray(arr, "RGB").save(path, format="JPEG", quality=90, optimize=False)
    return path


def logo(path: Path, size: int = 256, color: tuple[int, int, int] = (31, 73, 125)) -> Path:
    """A simple flat logo with transparency (PNG)."""
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.ellipse((8, 8, size - 8, size - 8), fill=(*color, 255))
    d.polygon(
        [(size * 0.30, size * 0.70), (size * 0.50, size * 0.25), (size * 0.70, size * 0.70)],
        fill=(255, 255, 255, 255),
    )
    d.rectangle((size * 0.42, size * 0.55, size * 0.58, size * 0.70), fill=(*color, 255))
    path.parent.mkdir(parents=True, exist_ok=True)
    img.save(path, format="PNG", optimize=False)
    return path


def diagram(path: Path, w: int = 500, h: int = 300) -> Path:
    """A flat figure for report-style PDFs (bars + axis), PNG."""
    img = Image.new("RGB", (w, h), (255, 255, 255))
    d = ImageDraw.Draw(img)
    d.line((40, h - 30, w - 10, h - 30), fill=(0, 0, 0), width=2)
    d.line((40, 10, 40, h - 30), fill=(0, 0, 0), width=2)
    vals = [0.35, 0.62, 0.48, 0.85, 0.71, 0.55]
    bw = (w - 80) / len(vals)
    for i, v in enumerate(vals):
        x0 = 50 + i * bw
        d.rectangle((x0, h - 30 - v * (h - 60), x0 + bw * 0.7, h - 31), fill=(79, 129, 189))
    path.parent.mkdir(parents=True, exist_ok=True)
    img.save(path, format="PNG", optimize=False)
    return path
