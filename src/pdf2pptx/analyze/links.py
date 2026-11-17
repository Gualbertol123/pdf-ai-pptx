"""Document links and metadata (M7): hyperlinks on shapes, sections, speaker notes,
producer detection.

* **Shape links.** Text runs already carry the link annotation under them (M4,
  ``text_group``). A link whose rectangle matches a shape or picture frame (IoU >= 0.7)
  becomes that element's click action (``Shape.hyperlink`` / ``Image.hyperlink``), and
  the same link is removed from the runs inside it: a button's label is not a link of
  its own. Internal links are ``#slide=<n>`` (1-based page number).
* **Sections** from the PDF outline (``metadata["pdf2pptx:outline"]``, read at ingest):
  a nested outline gives one section per top-level entry that has children (this is
  what PowerPoint's PDF export writes for a deck with sections). A flat outline that
  bookmarks every page in order is a per-slide outline, not sections; another flat
  outline (chapters of a report) gives one section per entry. The first section always
  starts on the first page.
* **Speaker notes**: the contents of text, free-text and markup annotations (comments)
  on a page, in reading order.
* **Producer** (``metadata["pdf2pptx:producer"]``): ``powerpoint`` | ``keynote`` |
  ``google_slides`` | ``libreoffice_impress`` | ``other`` from Producer/Creator, and
  ``slides_like`` when the pages have a slide aspect ratio (16:9, 16:10, 4:3) in
  landscape. A slide producer enables stronger placeholder mapping (``masters.py``).
"""

from __future__ import annotations

import json
import re
from typing import Any

from ..ingest.parse import OUTLINE_KEY
from ..ir import Annotation, Document, Group, Image, Page, Section, Shape, TextBlock

PRODUCER_KEY = "pdf2pptx:producer"
SLIDE_ASPECTS = (16 / 9, 16 / 10, 4 / 3)
# comments: text (sticky note), free text, and the markup annotations (line, square, circle,
# polygon, polyline, underline, squiggly, strikeout, stamp, caret, ink = pdfium 4..15).
# Popups (16) repeat their parent's contents; links and form widgets are not comments.
NOTE_SUBTYPES = {"text", "freetext", "highlight", *(f"subtype{k}" for k in range(4, 16))}
DEFAULT_SECTION = "Default Section"
_SLIDE_BOOKMARK = re.compile(r"^(slide|diapositiva|folie|diapositive|diapositivo|dia)\s*\d+\b",
                             re.IGNORECASE)


# --------------------------------------------------------------------------- producer


def detect_producer(doc: Document) -> dict[str, Any]:
    meta = " ".join((doc.metadata.get("Producer", ""), doc.metadata.get("Creator", ""))).lower()
    if "powerpoint" in meta:
        app = "powerpoint"
    elif "keynote" in meta:
        app = "keynote"
    elif "google" in meta and ("slides" in meta or "presentations" in meta):
        app = "google_slides"
    elif "impress" in meta:
        app = "libreoffice_impress"
    else:
        app = "other"
    sizes = [(p.width, p.height) for p in doc.pages if p.width > 0 and p.height > 0]
    landscape = [w / h for w, h in sizes if w > h]
    slides_like = bool(sizes) and len(landscape) == len(sizes) and all(
        any(abs(r - a) / a <= 0.02 for a in SLIDE_ASPECTS) for r in landscape)
    info = {"app": app, "slides_like": slides_like,
            "slide_producer": app != "other" or slides_like}
    doc.metadata[PRODUCER_KEY] = json.dumps(info, sort_keys=True)
    return info


def producer_info(doc: Document) -> dict[str, Any]:
    raw = doc.metadata.get(PRODUCER_KEY)
    if raw:
        try:
            data = json.loads(raw)
            if isinstance(data, dict):
                return data
        except ValueError:
            pass
    return detect_producer(doc)


# --------------------------------------------------------------------------- shape links


def _iou(a: tuple[float, float, float, float], b: tuple[float, float, float, float]) -> float:
    ix = max(0.0, min(a[2], b[2]) - max(a[0], b[0]))
    iy = max(0.0, min(a[3], b[3]) - max(a[1], b[1]))
    inter = ix * iy
    union = (a[2] - a[0]) * (a[3] - a[1]) + (b[2] - b[0]) * (b[3] - b[1]) - inter
    return inter / union if union > 0 else 0.0


def _within(inner: tuple[float, float, float, float], outer: tuple[float, float, float, float],
            tol: float = 2.0) -> bool:
    return (inner[0] >= outer[0] - tol and inner[1] >= outer[1] - tol
            and inner[2] <= outer[2] + tol and inner[3] <= outer[3] + tol)


def link_target(a: Annotation) -> str | None:
    if a.uri:
        return a.uri
    if a.dest_page is not None:
        return f"#slide={a.dest_page + 1}"
    return None


def shape_links(page: Page) -> int:
    """Move link annotations that cover a whole shape or picture onto it."""
    targets: list[Shape | Image] = [e for e in page.elements if isinstance(e, (Shape, Image))
                                    and not e.is_background and e.bbox[2] > e.bbox[0]
                                    and e.bbox[3] > e.bbox[1]
                                    and not (isinstance(e, Shape) and e.connector)]
    in_group = {c for g in page.elements if isinstance(g, Group) for c in g.child_ids}
    n = 0
    for a in page.annotations:
        if a.subtype != "link":
            continue
        link = link_target(a)
        if link is None:
            continue
        scored = sorted(((_iou(a.bbox, e.bbox), -e.z, e.id, e) for e in targets),
                        key=lambda t: (t[0], t[1], t[2]), reverse=True)
        if not scored or scored[0][0] < 0.7:
            continue
        best = scored[0][3]
        # in a PowerPoint group (shape + its shadow picture) the shape is the target
        if isinstance(best, Image) and best.id in in_group:
            alt = [t[3] for t in scored if isinstance(t[3], Shape) and t[0] >= 0.7]
            best = alt[0] if alt else best
        best.hyperlink = link
        best.notes.append(f"link:annotation:{a.id}")
        n += 1
        for tb in page.elements:
            if isinstance(tb, TextBlock) and _within(tb.bbox, best.bbox):
                for p in tb.paragraphs:
                    for r in p.runs:
                        if r.hyperlink == link:
                            r.hyperlink = None
    return n


# --------------------------------------------------------------------------- sections


def sections_from_outline(doc: Document) -> list[Section]:
    raw = doc.metadata.get(OUTLINE_KEY)
    if not raw:
        return []
    try:
        toc = [(int(lv), str(t), int(pg)) for lv, t, pg in json.loads(raw)]
    except (ValueError, TypeError):
        return []
    n_pages = len(doc.pages)
    toc = [(lv, t.strip(), pg) for lv, t, pg in toc if 0 <= pg < n_pages]
    if not toc:
        return []
    found: list[tuple[str, int]] = []
    if any(lv > 0 for lv, _, _ in toc):
        for k, (lv, title, pg) in enumerate(toc):
            has_children = k + 1 < len(toc) and toc[k + 1][0] > lv
            if lv == 0 and has_children:
                found.append((title, pg))
    else:
        pages = [pg for _, _, pg in toc]
        per_page = pages == list(range(len(pages))) and len(pages) >= max(1, n_pages - 1)
        slide_titles = sum(1 for _, t, _ in toc if _SLIDE_BOOKMARK.match(t)) >= 0.5 * len(toc)
        if per_page or slide_titles:
            return []
        found = [(t, pg) for _, t, pg in toc]
    out: list[Section] = []
    for title, pg in sorted(found, key=lambda x: x[1]):
        if out and pg <= out[-1].first_page:
            continue  # one section per start page, in page order
        out.append(Section(name=title[:255] or f"Section {len(out) + 1}", first_page=pg))
    if out and out[0].first_page > 0:
        out.insert(0, Section(name=DEFAULT_SECTION, first_page=0))
    return out if len(out) >= 2 else []


# --------------------------------------------------------------------------- notes


def page_notes(page: Page) -> str:
    items = [a for a in page.annotations if a.subtype in NOTE_SUBTYPES and a.contents.strip()]
    items.sort(key=lambda a: (round(a.bbox[1] / 12), a.bbox[0], a.id))
    seen: list[str] = []
    for a in items:
        text = a.contents.replace("\r\n", "\n").replace("\r", "\n").strip()
        if text not in seen:
            seen.append(text)
    return "\n".join(seen)


# --------------------------------------------------------------------------- entry


def analyze_links(doc: Document) -> dict[str, int]:
    stats = {"shape_links": 0, "notes_pages": 0, "sections": 0}
    detect_producer(doc)
    for page in doc.pages:
        if page.triage.fallback_image:
            continue
        stats["shape_links"] += shape_links(page)
    for page in doc.pages:
        page.notes_text = page_notes(page)
        stats["notes_pages"] += bool(page.notes_text)
    doc.sections = sections_from_outline(doc)
    stats["sections"] = len(doc.sections)
    return stats
