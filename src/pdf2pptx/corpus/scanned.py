"""Fake "scanned" PDFs: rasterize, rotate slightly, add noise and JPEG artifacts."""

from __future__ import annotations

import io
from pathlib import Path

import numpy as np
import pypdfium2 as pdfium
from PIL import Image


def make_scanned(src_pdf: Path, out_pdf: Path, seed: int, dpi: int = 200) -> int:
    rng = np.random.default_rng(seed)
    pdf = pdfium.PdfDocument(str(src_pdf))
    pages: list[Image.Image] = []
    try:
        for i in range(len(pdf)):
            page = pdf[i]
            img = page.render(scale=dpi / 72).to_pil().convert("L")
            angle = float(rng.uniform(-1.0, 1.0))
            img = img.rotate(angle, resample=Image.Resampling.BICUBIC, expand=False, fillcolor=255)
            arr = np.asarray(img, dtype=np.float64)
            arr = arr * 0.92 + 12  # paper is never pure white, ink never pure black
            arr += rng.normal(0, 6.0, size=arr.shape)
            speckle = rng.random(arr.shape) < 0.0008
            arr[speckle] = rng.uniform(0, 120, size=int(speckle.sum()))
            img = Image.fromarray(np.clip(arr, 0, 255).astype(np.uint8), "L")
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=55)
            buf.seek(0)
            pages.append(Image.open(buf).convert("RGB"))
    finally:
        pdf.close()
    out_pdf.parent.mkdir(parents=True, exist_ok=True)
    pages[0].save(out_pdf, format="PDF", save_all=True, append_images=pages[1:], resolution=dpi)
    return len(pages)
