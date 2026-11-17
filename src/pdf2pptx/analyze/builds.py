"""Build-up sequences (M7): page n+1 = page n + new elements.

Two consecutive pages form a build step when every top-level element of page n is on
page n+1 unchanged (same fingerprint, frame within 1 pt) or is a text block that grew
by trailing paragraphs (the first paragraphs identical, at the same place), and page
n+1 adds something (new elements or paragraphs). Identical pages are not a build. Page
n must have real content (a title, or two elements besides the background), so a
near-empty page followed by a full one is not mistaken for a build.

A chain p0 <- p1 <- ... <- pk is recorded as ``Page.build_of`` (p_j builds on
p_{j-1}). The last page holds every element; the writer (``TextOptions.builds``) keeps
only that page as one slide and animates each element / paragraph in at the step it
first appeared:

* ``build:<step>`` on an element appearing at click ``step`` (1..k);
* ``build_para:<paragraph>:<step>`` on a text block whose paragraphs appear one by one.

Collapsing is reversible: with builds off the writer writes every page as before.
Notes of absorbed pages are joined onto the final page.
"""

from __future__ import annotations

import itertools
import json
from collections.abc import Sequence
from typing import Any

from ..ir import Document, Group, Page, TextBlock
from .fingerprint import fingerprint, paragraph_key, same_box

BUILD_NOTE = "build:"
PARA_NOTE = "build_para:"
BUILDS_KEY = "pdf2pptx:builds"
TOL = 1.0


def top_items(page: Page) -> list[Any]:
    """Top-level elements: group members are compared through their group."""
    present = {e.id for e in page.elements}
    groups = [g for g in page.elements if isinstance(g, Group)
              and sum(1 for c in g.child_ids if c in present) >= 2]
    members = {c for g in groups for c in g.child_ids}
    return sorted((e for e in page.elements if e.id not in members and (
        not isinstance(e, Group) or e in groups)), key=lambda e: (e.z, e.id))


def _fp(page: Page, e: Any, cache: dict[str, str]) -> str:
    if e.id not in cache:
        kids = None
        if isinstance(e, Group):
            by_id = {x.id: x for x in page.elements}
            kids = [by_id[c] for c in e.child_ids if c in by_id]
        cache[e.id] = fingerprint(e, children=kids)
    return cache[e.id]


def _grown(a: TextBlock, b: TextBlock) -> bool:
    n = len(a.paragraphs)
    if not (0 < n < len(b.paragraphs)) or a.rotation != b.rotation or a.vertical != b.vertical:
        return False
    if abs(a.bbox[0] - b.bbox[0]) > TOL or abs(a.bbox[1] - b.bbox[1]) > TOL:
        return False
    return all(paragraph_key(pa) == paragraph_key(pb) and same_box(pa.bbox, pb.bbox, TOL)
               for pa, pb in zip(a.paragraphs, b.paragraphs[:n], strict=True))


def _has_content(page: Page, items: Sequence[Any]) -> bool:
    if any(isinstance(e, TextBlock) and e.role == "title" for e in items):
        return True
    return sum(1 for e in items if not getattr(e, "is_background", False)) >= 2


def match_pages(a: Page, b: Page) -> dict[str, tuple[str, int | None]] | None:
    """Map items of ``b`` to the item of ``a`` they continue: (id, None) when identical,
    (id, n paragraphs on ``a``) for a grown text block. None when ``b`` is not ``a`` plus
    something new."""
    if abs(a.width - b.width) > 0.5 or abs(a.height - b.height) > 0.5:
        return None
    if a.triage.fallback_image or b.triage.fallback_image:
        return None
    ia, ib = top_items(a), top_items(b)
    if not ia or not _has_content(a, ia):
        return None
    ca: dict[str, str] = {}
    cb: dict[str, str] = {}
    used: set[str] = set()
    out: dict[str, tuple[str, int | None]] = {}
    for ea in ia:
        fa = _fp(a, ea, ca)
        hit = next((eb for eb in ib if eb.id not in used and eb.kind == ea.kind
                    and _fp(b, eb, cb) == fa and same_box(ea.bbox, eb.bbox, TOL)), None)
        n: int | None = None
        if hit is None and isinstance(ea, TextBlock):
            hit = next((eb for eb in ib if eb.id not in used and isinstance(eb, TextBlock)
                        and _grown(ea, eb)), None)
            n = len(ea.paragraphs) if hit is not None else None
        if hit is None:
            return None
        used.add(hit.id)
        out[hit.id] = (ea.id, n)
    if len(used) == len(ib) and not any(n for _, n in out.values()):
        return None  # the same page again
    return out


def _chains(doc: Document) -> list[tuple[list[int], list[dict[str, tuple[str, int | None]]]]]:
    pages = doc.pages
    chains: list[tuple[list[int], list[dict[str, tuple[str, int | None]]]]] = []
    cur: list[int] = [0] if pages else []
    maps: list[dict[str, tuple[str, int | None]]] = []
    for k in range(1, len(pages)):
        m = match_pages(pages[k - 1], pages[k])
        if m is not None:
            cur.append(k)
            maps.append(m)
            continue
        if len(cur) > 1:
            chains.append((cur, maps))
        cur, maps = [k], []
    if len(cur) > 1:
        chains.append((cur, maps))
    return chains


def _annotate(doc: Document, chain: list[int],
              maps: list[dict[str, tuple[str, int | None]]]) -> int:
    """Notes on the final page's items; returns the number of animation effects."""
    pages = [doc.pages[i] for i in chain]
    final = pages[-1]
    effects = 0
    for e in top_items(final):
        cur, step = e.id, len(pages) - 1
        counts: dict[int, int] = {}
        if isinstance(e, TextBlock):
            counts[step] = len(e.paragraphs)
        for j in range(len(pages) - 1, 0, -1):
            hit = maps[j - 1].get(cur)
            if hit is None:
                break
            prev, n = hit
            step = j - 1
            if isinstance(e, TextBlock):
                counts[step] = n if n is not None else counts.get(j, len(e.paragraphs))
            cur = prev
        grows = isinstance(e, TextBlock) and len(set(counts.values())) > 1
        if grows:
            assert isinstance(e, TextBlock)
            for pi in range(len(e.paragraphs)):
                at = min(j for j, c in counts.items() if c > pi)
                if at > 0:
                    e.notes.append(f"{PARA_NOTE}{pi}:{at}")
                    effects += 1
        elif step > 0:
            e.notes.append(f"{BUILD_NOTE}{step}")
            effects += 1
    return effects


def analyze_builds(doc: Document) -> dict[str, int]:
    stats = {"build_sequences": 0, "build_pages_absorbed": 0, "build_effects": 0}
    records = []
    for chain, maps in _chains(doc):
        for prev, nxt in itertools.pairwise(chain):
            doc.pages[nxt].build_of = prev
        effects = _annotate(doc, chain, maps)
        final = doc.pages[chain[-1]]
        notes = [doc.pages[i].notes_text for i in chain if doc.pages[i].notes_text]
        final.notes_text = "\n".join(dict.fromkeys(notes))
        stats["build_sequences"] += 1
        stats["build_pages_absorbed"] += len(chain) - 1
        stats["build_effects"] += effects
        records.append({"pages": chain, "steps": len(chain) - 1, "effects": effects})
    if records:
        doc.metadata[BUILDS_KEY] = json.dumps(records, sort_keys=True)
    return stats


def element_step(e: Any) -> int:
    for n in getattr(e, "notes", ()):
        if n.startswith(BUILD_NOTE):
            return int(n[len(BUILD_NOTE):])
    return 0


def paragraph_steps(tb: TextBlock) -> dict[int, int]:
    out: dict[int, int] = {}
    for n in tb.notes:
        if n.startswith(PARA_NOTE):
            pi, step = n[len(PARA_NOTE):].split(":")
            out[int(pi)] = int(step)
    return out


def absorbed_pages(doc: Document) -> set[int]:
    """Pages folded into a later slide (some page builds on them)."""
    return {p.build_of for p in doc.pages if p.build_of is not None}


def page_to_slide(doc: Document, builds: bool) -> list[int]:
    """Slide index of every page (an absorbed page maps to its chain's final slide)."""
    gone = absorbed_pages(doc) if builds else set()
    slide_of: list[int] = []
    k = 0
    for p in doc.pages:
        slide_of.append(k)
        if p.index not in gone:
            k += 1
    # absorbed pages point at the next written slide: that is their chain's final page
    return slide_of
