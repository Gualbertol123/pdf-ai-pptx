"""Page renders, created lazily and cached in ``work/<sha256>/renders/``.

Renders are written to disk and never kept in memory beyond the call that made
them, so a long document streams page by page.
"""

from __future__ import annotations

from pathlib import Path

import pypdfium2 as pdfium

DPIS = (150, 300)


def render_rel(index: int, dpi: int) -> str:
    """Path of a render relative to the document's work dir."""
    return f"renders/p{index:04d}_{dpi}.png"


def ensure_render(pdf: Path, work_dir: Path, index: int, dpi: int) -> Path:
    """Render page ``index`` at ``dpi`` unless the cached PNG already exists."""
    out = work_dir / render_rel(index, dpi)
    if out.is_file():
        return out
    out.parent.mkdir(parents=True, exist_ok=True)
    doc = pdfium.PdfDocument(str(pdf))
    try:
        page = doc[index]
        try:
            w, h = page.get_size()  # already swapped for /Rotate
            target = (max(1, round(w * dpi / 72)), max(1, round(h * dpi / 72)))
            img = page.render(scale=dpi / 72, draw_annots=True).to_pil().convert("RGB")
            if img.size != target:
                img = img.resize(target)
            tmp = out.with_suffix(".tmp.png")
            img.save(tmp, "PNG", optimize=False)
            tmp.replace(out)
        finally:
            page.close()
    finally:
        doc.close()
    return out
