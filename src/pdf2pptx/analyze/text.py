"""ANALYZE (text): font matching, text grouping, roles and reading order for a document.

A page whose grouping fails keeps no TextBlocks; the writer then falls back to the
baseline one-box-per-line layout for it (graceful degradation, logged).
"""

from __future__ import annotations

from .. import log
from ..fonts.embed import embed_candidates
from ..fonts.match import match_document
from ..ir import Document, Image, TextBlock
from . import reading_order, roles, text_group

LOG = log.get(__name__)

BULLET_IMAGE_NOTE = "bullet_image"


def analyze_text(doc: Document) -> Document:
    matches = match_document(doc)
    embed_candidates(doc)
    blocks: dict[int, list[TextBlock]] = {}
    for page in doc.pages:
        if page.triage.fallback_image:
            continue
        try:
            blocks[page.index] = text_group.group_page(page, doc.fonts, matches)
        except Exception as exc:
            LOG.error("page %d: text grouping failed (%s: %s); using one box per line",
                      page.index, type(exc).__name__, exc, exc_info=LOG.isEnabledFor(10))
            page.triage.reason = (page.triage.reason + "; " if page.triage.reason else "") + (
                f"text grouping failed: {type(exc).__name__}")
    roles.assign_roles(doc, blocks)
    for page in doc.pages:
        bl = blocks.get(page.index)
        if not bl:
            continue
        boxes = [reading_order.rotated_aabb(tb.bbox, tb.rotation) for tb in bl]
        for rank, i in enumerate(reading_order.xy_cut(boxes)):
            bl[i].reading_order = rank
        bullet_ids = text_group.picture_bullet_ids(bl)
        for e in page.elements:
            if isinstance(e, Image) and e.id in bullet_ids and BULLET_IMAGE_NOTE not in e.notes:
                e.notes.append(BULLET_IMAGE_NOTE)
        page.elements.extend(bl)
    return doc
