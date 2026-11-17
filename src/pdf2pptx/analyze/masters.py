"""Slide master, layouts and placeholders (M7).

1. **Page furniture.** Top-level elements that repeat on at least ``MIN_PAGES`` pages
   with the same fingerprint (``fingerprint.py``) and frame (within ``TOL`` pt):
   pictures (logos), shapes and groups (background art, rules, bands), the slide
   background, and text blocks that are footers, headers or sit in the top / bottom
   12% of the page. A page keeps an occurrence on the slide when moving it would put it
   under a slide element it was drawn above (master and layout shapes render below
   every slide shape).
2. **Master vs layout.** Furniture on at least 60% of the pages goes to the master;
   the rest to the layouts of the pages that have it. A page missing some master
   furniture gets a layout that hides the master graphics (``show_master=False``) and
   carries that page's furniture itself.
3. **Page numbers.** A repeated page-number block (``roles.py``: role ``page_number``,
   run with ``field_type="slidenum"``) becomes the slide-number placeholder: the master
   and every layout get an ``sldNum`` placeholder at its place, and the slide's block
   is written into a ``sldNum`` placeholder with a slide-number field.
4. **Layouts.** Pages are clustered by layout signature: the title frame (centre or
   left edge, by alignment, plus top and bottom), up to two body frames and a subtitle,
   together with the page's layout furniture. At most ``max_layouts`` layouts are
   made (the most frequent signatures); other pages use a plain layout and keep their
   text in text boxes. Layouts are named like PowerPoint's (Title Slide, Title and
   Content, Two Content, Title Only, Blank).
5. **Placeholders.** A title / subtitle / body block becomes the matching placeholder
   of its layout (``TextBlock.placeholder``) when its frame fits the layout's slot.
   Placeholder mapping is stronger for slide producers (``links.producer_info``):
   titles from role confidence 0.55 and body placeholders; for other producers
   (reports) only confident titles (>= 0.8) and no body placeholders.

Pages keep every element: an element now drawn by the master or a layout carries the
note ``in_master:<copy id>`` / ``in_layout:<copy id>``. The flat writer (tuner,
``--no-structure``, the fidelity fallback) ignores these notes; the structured writer
skips such elements and writes the copies in ``Document.masters``. Copies carry
``from:<page>:<element id>`` so ``refresh_copies`` can take over the tuner's changes.
"""

from __future__ import annotations

import copy
import json
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from typing import Any

from ..ir import Document, Group, Image, Layout, Master, Page, Shape, Table, TextBlock
from .builds import absorbed_pages, top_items
from .fingerprint import fingerprint, same_box
from .links import producer_info

MASTER_NOTE = "in_master:"
LAYOUT_NOTE = "in_layout:"
FROM_NOTE = "from:"
MASTERS_KEY = "pdf2pptx:masters"
MIN_PAGES = 3
MASTER_SHARE = 0.6
TOL = 2.0
BAND = 0.12
FURNITURE_ROLES = ("footer", "header", "page_number")
PLAIN_LAYOUT = "Blank (no master graphics)"
LAYOUT_TYPES = {"Title Slide": "title", "Title and Content": "obj", "Two Content": "twoObj",
                "Title Only": "titleOnly", "Blank": "blank", PLAIN_LAYOUT: "blank"}
PH_ROLES = {"title": "title", "subtitle": "subTitle"}


@dataclass
class _Cluster:
    key: str
    kind: str
    rep: tuple[int, Any]  # (page index, element)
    members: dict[int, Any] = field(default_factory=dict)  # page index -> element
    page_number: bool = False


def _bullet_image(e: Any) -> bool:
    return isinstance(e, Image) and "bullet_image" in e.notes


def _eligible(e: Any, page: Page) -> bool:
    if isinstance(e, Table) or _bullet_image(e):
        return False
    if isinstance(e, Shape) and (e.connector is not None or e.fill.kind == "pattern"):
        return False
    if isinstance(e, TextBlock):
        if e.role in FURNITURE_ROLES:
            return True
        if e.role in ("title", "subtitle"):
            return False
        return e.bbox[3] <= BAND * page.height or e.bbox[1] >= (1 - BAND) * page.height
    return isinstance(e, (Image, Shape, Group))


def _is_page_number(e: Any) -> bool:
    return isinstance(e, TextBlock) and e.role == "page_number" and any(
        r.field_type == "slidenum" for p in e.paragraphs for r in p.runs)


def _number_box(e: TextBlock) -> tuple[float, ...]:
    """Page numbers grow with their digits: compare the aligned edge and the height."""
    align = e.paragraphs[0].alignment if e.paragraphs else "left"
    x = e.bbox[2] if align == "right" else ((e.bbox[0] + e.bbox[2]) / 2 if align == "center"
                                           else e.bbox[0])
    return (x, e.bbox[1], e.bbox[3])


def _children(page: Page, g: Group) -> list[Any]:
    by_id = {x.id: x for x in page.elements}
    return [by_id[c] for c in g.child_ids if c in by_id]


def _clusters(doc: Document, pages: list[Page]) -> list[_Cluster]:
    found: dict[str, list[_Cluster]] = defaultdict(list)
    for p in pages:
        for e in top_items(p):
            if not _eligible(e, p):
                continue
            pn = _is_page_number(e)
            kids = _children(p, e) if isinstance(e, Group) else None
            key = f"{e.kind}:{fingerprint(e, mask_digits=pn, children=kids)}"
            box = _number_box(e) if pn else e.bbox
            for cl in found[key]:
                rbox = _number_box(cl.rep[1]) if pn else cl.rep[1].bbox
                if p.index not in cl.members and same_box(box, rbox, TOL):
                    cl.members[p.index] = e
                    break
            else:
                found[key].append(_Cluster(key, e.kind, (p.index, e), {p.index: e},
                                           page_number=pn))
    out = [cl for key in sorted(found) for cl in found[key] if len(cl.members) >= MIN_PAGES]
    return sorted(out, key=lambda c: (c.rep[0], c.rep[1].z, c.rep[1].id))


def _overlap(a: tuple[float, float, float, float], b: tuple[float, float, float, float]) -> bool:
    return min(a[2], b[2]) - max(a[0], b[0]) > 0.5 and min(a[3], b[3]) - max(a[1], b[1]) > 0.5


def _z_check(pages: dict[int, Page], clusters: list[_Cluster]) -> None:
    """Drop occurrences that a slide element would cover wrongly (2 passes)."""
    for _ in range(2):
        moved: dict[int, set[str]] = defaultdict(set)
        for cl in clusters:
            if cl.page_number:
                continue
            for pi, e in cl.members.items():
                moved[pi].add(e.id)
        for cl in clusters:
            if cl.page_number:
                continue
            for pi in sorted(cl.members):
                e = cl.members[pi]
                stays = [o for o in top_items(pages[pi]) if o.id not in moved[pi]
                         and not getattr(o, "is_background", False)]
                if any(o.z < e.z and _overlap(o.bbox, e.bbox) for o in stays):
                    del cl.members[pi]
        clusters[:] = [c for c in clusters if len(c.members) >= MIN_PAGES]


# --------------------------------------------------------------------------- geometry


def _strong(doc: Document) -> bool:
    return bool(producer_info(doc).get("slide_producer"))


def _slots(page: Page, strong: bool, furniture: set[str]) -> dict[str, TextBlock]:
    """Text blocks that may become placeholders: title, subTitle, body, body2."""
    blocks = [e for e in page.elements if isinstance(e, TextBlock) and e.id not in furniture
              and not e.rotation and not e.vertical and any(
                  r.text.strip() for p in e.paragraphs for r in p.runs)]
    out: dict[str, TextBlock] = {}
    titles = [b for b in blocks if b.role == "title"
              and b.role_confidence >= (0.55 if strong else 0.8)]
    if titles:
        out["title"] = max(titles, key=lambda b: (b.role_confidence, -b.bbox[1]))
    if not strong:
        return out
    subs = [b for b in blocks if b.role == "subtitle" and "title" in out]
    if subs:
        out["subTitle"] = min(subs, key=lambda b: b.bbox[1])
    area = page.width * page.height
    bodies = [b for b in blocks if b.role == "body" and b.role_confidence >= 0.6
              and (len(b.paragraphs) >= 2 or any(p.bullet.kind != "none" for p in b.paragraphs)
                   or (b.bbox[2] - b.bbox[0]) * (b.bbox[3] - b.bbox[1]) >= 0.05 * area)]
    bodies = sorted(bodies, key=lambda b: -(b.bbox[2] - b.bbox[0]) * (b.bbox[3] - b.bbox[1]))[:2]
    for name, b in zip(("body", "body2"), sorted(bodies, key=lambda b: (b.bbox[0], b.bbox[1])),
                       strict=False):
        out[name] = b
    if "subTitle" in out and "body" in out:
        del out["subTitle"]  # a subtitle placeholder belongs to a title-slide layout
    return out


def _anchor_box(tb: TextBlock) -> tuple[float, float, float]:
    align = tb.paragraphs[0].alignment if tb.paragraphs else "left"
    x = (tb.bbox[0] + tb.bbox[2]) / 2 if align == "center" else (
        tb.bbox[2] if align == "right" else tb.bbox[0])
    return (x, tb.bbox[1], tb.bbox[3])


def _geom_match(a: dict[str, TextBlock], b: dict[str, TextBlock], page: Page) -> bool:
    if set(a) != set(b):
        return False
    tol = (0.04 * page.width, 0.04 * page.height, 0.08 * page.height)
    for k in a:
        pa, pb = _anchor_box(a[k]), _anchor_box(b[k])
        if abs(pa[0] - pb[0]) > tol[0] or abs(pa[1] - pb[1]) > tol[1]:
            return False
        if k.startswith("title") and abs(pa[2] - pb[2]) > tol[2]:
            return False
    return True


def _union(boxes: list[tuple[float, float, float, float]]) -> tuple[float, float, float, float]:
    return (round(min(b[0] for b in boxes), 2), round(min(b[1] for b in boxes), 2),
            round(max(b[2] for b in boxes), 2), round(max(b[3] for b in boxes), 2))


def _layout_name(slots: set[str]) -> str:
    if "title" in slots and "body2" in slots:
        return "Two Content"
    if "title" in slots and "body" in slots:
        return "Title and Content"
    if "title" in slots and "subTitle" in slots:
        return "Title Slide"
    if "title" in slots:
        return "Title Only"
    return "Blank"


def _fits(tb: TextBlock, slot: tuple[float, float, float, float], page: Page) -> bool:
    cx, cy = (tb.bbox[0] + tb.bbox[2]) / 2, (tb.bbox[1] + tb.bbox[3]) / 2
    gx, gy = 0.05 * page.width, 0.05 * page.height
    return slot[0] - gx <= cx <= slot[2] + gx and slot[1] - gy <= cy <= slot[3] + gy


# --------------------------------------------------------------------------- copies


def _copy(e: Any, page: Page, prefix: str, owner: list[Any]) -> str:
    """Deep copy ``e`` (and a group's members) into ``owner``; returns the copy's id."""
    c = copy.deepcopy(e)
    c.id = f"{prefix}_{e.id}"
    c.notes = [n for n in c.notes if not n.startswith((MASTER_NOTE, LAYOUT_NOTE))]
    c.notes.append(f"{FROM_NOTE}{page.index}:{e.id}")
    if isinstance(e, Group):
        kids = _children(page, e)
        c.child_ids = [_copy(k, page, prefix, owner) for k in kids]
    owner.append(c)
    return str(c.id)


def _mark(e: Any, page: Page, note: str, copy_id: str, prefix: str) -> None:
    e.notes.append(f"{note}{copy_id}")
    if isinstance(e, Group):
        for k in _children(page, e):
            k.notes.append(f"{note}{prefix}_{k.id}")


def refresh_copies(doc: Document) -> None:
    """Re-copy master / layout elements from their (possibly tuned) page originals."""
    pages = {p.index: p for p in doc.pages}

    def refresh(items: list[Any]) -> None:
        for k, c in enumerate(items):
            src = next((n[len(FROM_NOTE):] for n in c.notes if n.startswith(FROM_NOTE)), None)
            if src is None:
                continue
            pi, _, eid = src.partition(":")
            page = pages.get(int(pi))
            orig = next((e for e in page.elements if e.id == eid), None) if page else None
            if orig is None or orig.kind != c.kind:
                continue
            new = copy.deepcopy(orig)
            new.id, new.notes = c.id, list(c.notes)
            if isinstance(new, Group) and isinstance(c, Group):
                new.child_ids = list(c.child_ids)
            items[k] = new

    for m in doc.masters:
        refresh(m.elements)
        for lay in m.layouts:
            refresh(lay.elements)


# --------------------------------------------------------------------------- entry


def analyze_masters(doc: Document, max_layouts: int = 8) -> dict[str, int]:
    stats = {"master_elements": 0, "layout_elements": 0, "layouts": 0, "placeholders": 0,
             "slide_number_placeholders": 0}
    doc.masters = []
    live = [p for p in doc.pages if not p.triage.fallback_image]
    if not live:
        return stats
    sizes = Counter((p.width, p.height) for p in live)
    dom = min(sizes.items(), key=lambda kv: (-kv[1], kv[0]))[0]
    gone = absorbed_pages(doc)
    cand = [p for p in live if (p.width, p.height) == dom and p.index not in gone]
    pages = {p.index: p for p in doc.pages}
    strong = _strong(doc)

    clusters = _clusters(doc, cand) if len(cand) >= MIN_PAGES else []
    _z_check(pages, clusters)
    numbers = [c for c in clusters if c.page_number]
    furn = [c for c in clusters if not c.page_number]
    n_cand = len(cand)
    in_master = [c for c in furn if len(c.members) >= MASTER_SHARE * n_cand]
    master = Master(id="master1", name="pdf2pptx Master")

    # furniture element ids per page (not the page number: it stays a slide placeholder)
    furn_ids: dict[int, set[str]] = defaultdict(set)
    for c in clusters:
        for pi, e in c.members.items():
            furn_ids[pi].add(e.id)

    # layout signature per candidate page
    sig: dict[int, tuple[dict[str, TextBlock], frozenset[int], bool]] = {}
    for p in cand:
        has = {k for k, c in enumerate(furn) if p.index in c.members}
        master_set = {furn.index(c) for c in in_master}
        show_master = master_set <= has
        own = frozenset(has - master_set) if show_master else frozenset(has)
        sig[p.index] = (_slots(p, strong, furn_ids[p.index]), own, show_master)

    groups: list[list[int]] = []  # pages per layout signature, in page order
    for pi in sorted(sig):
        slots, own, show = sig[pi]
        for g in groups:
            s0, o0, sh0 = sig[g[0]]
            if o0 == own and sh0 == show and _geom_match(slots, s0, pages[pi]):
                g.append(pi)
                break
        else:
            groups.append([pi])
    keep = sorted(groups, key=lambda g: (-len(g), g[0]))[:max(max_layouts - 1, 1)]
    keep = sorted(keep, key=lambda g: g[0])

    master_copies: dict[int, str] = {}
    for c in in_master:
        rp, re_ = c.rep
        master_copies[id(c)] = _copy(re_, pages[rp], master.id, master.elements)
    stats["master_elements"] = len(in_master)

    page_layout: dict[int, Layout] = {}
    names: Counter[str] = Counter()
    for gi, g in enumerate(keep, start=1):
        slots0, own, show = sig[g[0]]
        lay_id = f"layout{gi}"
        name = _layout_name(set(slots0))
        names[name] += 1
        lay = Layout(id=lay_id, name=name if names[name] == 1 else f"{name} {names[name]}",
                     show_master=show)
        for k in sorted(slots0):
            boxes = [sig[pi][0][k].bbox for pi in g]
            lay.placeholders[k] = _union(boxes)
        for k in sorted(own):
            c = furn[k]
            rp = min(pi for pi in g if pi in c.members)
            _copy(c.members[rp], pages[rp], lay_id, lay.elements)
            stats["layout_elements"] += 1
        master.layouts.append(lay)
        for pi in g:
            page_layout[pi] = lay

    plain: Layout | None = None

    def plain_layout() -> Layout:
        nonlocal plain
        if plain is None:
            plain = Layout(id=f"layout{len(master.layouts) + 1}", name=PLAIN_LAYOUT,
                           show_master=False)
            master.layouts.append(plain)
        return plain

    for p in live:
        if p.index not in page_layout and p.index not in gone:
            page_layout[p.index] = plain_layout()
    if numbers:
        nb = _union([e.bbox for c in numbers for e in c.members.values()])
        for lay in master.layouts:
            lay.placeholders["sldNum"] = nb

    # absorbed build pages share the final page's layout (when they hold its furniture)
    for p in doc.pages:
        if p.index in gone:
            nxt = p
            while nxt.index in gone:
                nxt = next(q for q in doc.pages if q.build_of == nxt.index)
            if nxt.index in page_layout:
                page_layout[p.index] = page_layout[nxt.index]

    # mark page elements, map placeholders
    for p in live:
        lay = page_layout.get(p.index)
        if lay is None:
            continue
        p.layout_id = lay.id
        if lay.name == PLAIN_LAYOUT:
            _map_numbers(p, numbers, lay, stats)
            continue
        ok = _mark_page(p, lay, master, in_master, furn, master_copies, pages)
        if not ok:  # furniture missing (absorbed build page): plain layout, all on the slide
            _unmark(p)
            lay = plain_layout()
            p.layout_id = lay.id
            if numbers:
                lay.placeholders.setdefault("sldNum", _union(
                    [e.bbox for c in numbers for e in c.members.values()]))
            _map_numbers(p, numbers, lay, stats)
            continue
        _map_numbers(p, numbers, lay, stats)
        furniture = {e.id for e in p.elements if any(
            n.startswith((MASTER_NOTE, LAYOUT_NOTE)) for n in e.notes)}
        for k, tb in _slots(p, strong, furniture).items():
            slot = lay.placeholders.get(k)
            if slot is not None and tb.placeholder is None and _fits(tb, slot, p):
                tb.placeholder = k
                stats["placeholders"] += 1
    stats["layouts"] = len(master.layouts)
    if master.layouts:
        doc.masters = [master]
    summary = {"layouts": [{"id": lay.id, "name": lay.name, "show_master": lay.show_master,
                            "placeholders": sorted(lay.placeholders),
                            "elements": len(lay.elements),
                            "pages": sorted(i for i, lo in page_layout.items() if lo is lay)}
                           for lay in master.layouts],
               "master_elements": len(master.elements), "strong_placeholders": strong}
    doc.metadata[MASTERS_KEY] = json.dumps(summary, sort_keys=True)
    return stats


def _map_numbers(p: Page, numbers: list[_Cluster], lay: Layout, stats: dict[str, int]) -> None:
    if "sldNum" not in lay.placeholders:
        return
    for c in numbers:
        e = c.members.get(p.index)
        if e is None:  # an absorbed build page: same block as its final page
            e = next((x for x in p.elements if _is_page_number(x)
                      and same_box(_number_box(x), _number_box(c.rep[1]), TOL)), None)
        if e is not None and e.placeholder is None:
            e.placeholder = "sldNum"
            stats["slide_number_placeholders"] += 1


def _find_same(p: Page, orig: Any, orig_page: Page) -> Any:
    kids = _children(orig_page, orig) if isinstance(orig, Group) else None
    fp = fingerprint(orig, children=kids)
    for e in top_items(p):
        if e.kind != orig.kind or not same_box(e.bbox, orig.bbox, TOL):
            continue
        ek = _children(p, e) if isinstance(e, Group) else None
        if fingerprint(e, children=ek) == fp:
            return e
    return None


def _mark_page(p: Page, lay: Layout, master: Master, in_master: list[_Cluster],
               furn: list[_Cluster], master_copies: dict[int, str],
               pages: dict[int, Page]) -> bool:
    """Note the page elements drawn by the master / layout. False if one is missing."""
    todo: list[tuple[Any, str, str, str]] = []  # (element, note, copy id, prefix)
    if lay.show_master:
        for c in in_master:
            e = c.members.get(p.index)
            if e is None:
                rp, re_ = c.rep
                e = _find_same(p, re_, pages[rp])
            if e is None:
                return False
            todo.append((e, MASTER_NOTE, master_copies[id(c)], master.id))
    for cpy in lay.elements:
        src = next((n[len(FROM_NOTE):] for n in cpy.notes if n.startswith(FROM_NOTE)), "")
        pi, _, eid = src.partition(":")
        c = next((c for c in furn if int(pi) in c.members and c.members[int(pi)].id == eid),
                 None)
        if c is None:  # a group member copy: handled through its group
            continue
        e = c.members.get(p.index)
        if e is None:
            e = _find_same(p, c.members[int(pi)], pages[int(pi)])
        if e is None:
            return False
        todo.append((e, LAYOUT_NOTE, cpy.id, lay.id))
    for e, note, cid, prefix in todo:
        _mark(e, p, note, cid, prefix)
    return True


def _unmark(p: Page) -> None:
    for e in p.elements:
        e.notes = [n for n in e.notes if not n.startswith((MASTER_NOTE, LAYOUT_NOTE))]
        if isinstance(e, TextBlock) and e.placeholder != "sldNum":
            e.placeholder = None
