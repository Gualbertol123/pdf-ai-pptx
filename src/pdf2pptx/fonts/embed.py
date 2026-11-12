"""Embedded PDF fonts: extraction helpers, embedding permission, PPTX embedding candidates.

Ingest (``ingest/parse.py``) writes every embedded font program to
``work/<sha>/fonts/<hash>.<ext>`` and records its OS/2 ``fsType``. This module decides
which fonts *should* travel with the PPTX: embedded in the PDF, not installed here,
and licensed for embedding.

**Gap (M4): the PPTX embedding itself is not written.** PowerPoint stores embedded
fonts as ``/ppt/fonts/fontN.fntdata`` parts in an EOT-style container, referenced from
``p:embeddedFontLst`` per family and style. PDF programs are almost always *subsets*
with re-mapped or missing ``cmap`` tables (and bare CFF / Type 1 programs have no
OS/2 table at all), so a faithful .fntdata needs a cmap rebuilt from the PDF's
ToUnicode, a full ``name``/``OS/2`` table, and the EOT wrapper. The candidates are
recorded in ``Document.metadata["pdf2pptx:embed_candidates"]`` and listed in
PROGRESS.md; the writer falls back to the closest installed font (``fonts/match.py``).
"""

from __future__ import annotations

import io
import json
from pathlib import Path

from ..ir import Document, FontRef

# OS/2 fsType bits (OpenType spec)
FS_RESTRICTED = 0x0002  # must not be embedded
FS_PREVIEW_PRINT = 0x0004
FS_EDITABLE = 0x0008
FS_NO_SUBSETTING = 0x0100
FS_BITMAP_ONLY = 0x0200


def program_extension(data: bytes) -> str:
    """File extension for a font program, by its magic number."""
    head = data[:4]
    if head in (b"\x00\x01\x00\x00", b"true", b"typ1"):
        return ".ttf"
    if head == b"OTTO":
        return ".otf"
    if head == b"ttcf":
        return ".ttc"
    if data[:2] == b"%!" or data[:1] == b"\x80":
        return ".pfb"
    if data[:1] == b"\x01":
        return ".cff"
    return ".bin"


def read_fs_type(path: Path) -> int | None:
    """OS/2 fsType of a TrueType/OpenType program, or None (bare CFF, Type 1, no OS/2)."""
    if path.suffix.lower() not in (".ttf", ".otf"):
        return None
    try:
        from fontTools.ttLib import TTFont

        f = TTFont(io.BytesIO(path.read_bytes()), lazy=True)
        try:
            if "OS/2" not in f:
                return None
            return int(f["OS/2"].fsType)
        finally:
            f.close()
    except Exception:
        return None


def embedding_allowed(fs_type: int | None) -> bool | None:
    """True if the fsType allows embedding in a document that can be edited/printed."""
    if fs_type is None:
        return None
    if fs_type & FS_BITMAP_ONLY:
        return False
    return not (fs_type & FS_RESTRICTED and not fs_type & (FS_PREVIEW_PRINT | FS_EDITABLE))


def embed_candidates(doc: Document) -> list[dict[str, object]]:
    """Fonts embedded in the PDF, not installed here, whose license allows embedding.

    Also sets ``FontRef.embeddable`` on every font with a known fsType.
    """
    out: list[dict[str, object]] = []
    for fid in sorted(doc.fonts):
        fr: FontRef = doc.fonts[fid]
        fr.embeddable = embedding_allowed(fr.fs_type) if fr.embedded else None
        installed = fr.match_source in ("name", "alias")
        if fr.embedded and not installed and fr.embeddable:
            out.append({"font": fid, "family": fr.family, "program": fr.program,
                        "fs_type": fr.fs_type, "subset": fr.subset,
                        "substitute": fr.matched_family})
    doc.metadata["pdf2pptx:embed_candidates"] = json.dumps(out, sort_keys=True)
    return out
