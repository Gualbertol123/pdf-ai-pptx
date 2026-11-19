"""Document -> deck (M10, ``--mode deck``): rewrite a long document as a presentation.

1. **Outline** (rules, no model): the analysed IR becomes sections (headings by size /
   weight / role), text blocks, tables, charts and pictures with their captions. Every
   item keeps its IR id and page.
2. **Plan** (text model, schema-constrained): slide titles, 3-5 bullets that cite block
   ids, the visual to reuse (an id from the outline) and speaker notes. A document that
   does not fit the context is first condensed chunk by chunk into grounded key points
   (each citing block ids); the deck is then planned from the points.
3. **Validate**: every bullet must cite a block that exists; every number in a bullet
   must appear in the cited blocks (else in another block, which is then cited too, else
   the bullet is dropped). Notes sentences with numbers absent from the document are
   dropped; titles with such numbers fall back to the section heading. Bullets whose
   words barely overlap their sources are flagged in ``deck_sources.json``.
4. **Build**: a clean master (title / content / two content / title only layouts) with
   the source theme, placeholders for titles and bullets, the reused tables / charts /
   pictures copied from the IR and scaled, notes with the source pages, then a
   PowerPoint fit loop (COM text layout): text that overflows its placeholder is set
   smaller until it fits.

Without a text model (Lite tier, or no model fits the free RAM) the plan is
extractive: the first sentence of each paragraph per section (``source="rule"``).
"""

from __future__ import annotations

import copy
import dataclasses
import json
import math
import re
import time
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from rapidfuzz import fuzz

from .. import log, progress
from ..config import Config
from ..ir import (
    Bullet,
    Chart,
    Color,
    Document,
    Fill,
    Image,
    Layout,
    Master,
    Page,
    Paragraph,
    Shape,
    Table,
    TextBlock,
    TextRun,
    Theme,
)

LOG = log.get(__name__)

DECK_VERSION = "10.0"
SLIDE_W, SLIDE_H = 960.0, 540.0
SKIP_ROLES = {"footer", "page_number", "header"}
CAPTION_RE = re.compile(r"^\s*(fig(\.|ure|ura)?|tab(\.|le|ella)?|chart|grafico|graph)\s*\d",
                        re.I)
NUM_RE = re.compile(r"(?<![\w])[-+]?\d+(?:[.,]\d+)*(?:%|‰)?")
SINGLE_PASS_CHARS = 22000  # source text that fits one planning call (16k ctx)
CHUNK_CHARS = 15000
MAX_BULLETS = 5
NOTE_OVERLAP = 0.5  # notes / subtitle sentences: share of content words found in the source
LANG_NAMES = {"it": "Italian", "en": "English"}
SOURCES_LABEL = {"it": "Fonti", "en": "Sources"}
CONTINUED = {"it": "(segue)", "en": "(cont.)"}
PAGE_LABEL = {"it": "p.", "en": "p."}
STOP = {
    "it": {"il", "lo", "la", "gli", "le", "di", "che", "è", "per", "una", "con", "non", "del",
           "della", "dei", "nel", "sono", "anche", "più", "alla", "delle", "questo", "ogni"},
    "en": {"the", "and", "of", "to", "is", "in", "that", "for", "with", "are", "was", "this",
           "were", "from", "have", "has", "which", "their", "after", "each", "all"},
}
SYSTEM = ("You turn documents into clear presentation decks. You use only information "
          "that is in the source document: never add facts, numbers, names or claims that "
          "are not there. Copy every number exactly as the source writes it. Every bullet "
          "lists the ids of the source blocks it comes from.")


# --------------------------------------------------------------------------- outline


@dataclass
class Src:
    id: str
    page: int
    kind: str  # text | table | chart | image
    text: str
    caption: str = ""


@dataclass
class Section:
    title: str
    page: int
    heading_id: str | None = None
    items: list[Src] = field(default_factory=list)

    def chars(self) -> int:
        return sum(len(s.text) + len(s.caption) + 12 for s in self.items) + len(self.title)


@dataclass
class Outline:
    title: str
    subtitle: str
    lang: str
    pages: int
    sections: list[Section]
    texts: dict[str, str]  # element id -> full text (numbers are checked against it)
    pages_of: dict[str, int]
    visuals: dict[str, Src]
    theme: Theme | None = None

    def all_text(self) -> str:
        return "\n".join(self.texts[k] for k in sorted(self.texts))

    def as_dict(self) -> dict[str, Any]:
        return {"title": self.title, "subtitle": self.subtitle, "lang": self.lang,
                "pages": self.pages,
                "sections": [{"title": s.title, "page": s.page + 1, "heading": s.heading_id,
                              "items": [{"id": i.id, "kind": i.kind, "page": i.page + 1,
                                         "chars": len(i.text), "caption": i.caption}
                                        for i in s.items]} for s in self.sections]}


def _para_text(p: Paragraph) -> str:
    return re.sub(r"\s+", " ", "".join(r.text for r in p.runs)).strip()


def _size(p: Paragraph) -> float:
    sizes = [r.size for r in p.runs if r.text.strip()]
    return max(sizes) if sizes else 0.0


def _bold(p: Paragraph) -> bool:
    runs = [r for r in p.runs if r.text.strip()]
    return bool(runs) and all(r.bold for r in runs)


def _in_master(e: Any) -> bool:
    return any(n.startswith(("in_master:", "in_layout:")) for n in e.notes)


def _body_size(doc: Document) -> float:
    c: Counter[float] = Counter()
    for p in doc.pages:
        for e in p.elements:
            if isinstance(e, TextBlock) and e.role == "body":
                for para in e.paragraphs:
                    for r in para.runs:
                        c[round(r.size, 1)] += len(r.text)
    return c.most_common(1)[0][0] if c else 10.0


def detect_lang(text: str) -> str:
    words = re.findall(r"[a-zàèéìòù]+", text.lower())
    hits = {k: sum(1 for w in words if w in v) for k, v in STOP.items()}
    return "it" if hits["it"] > hits["en"] else "en"


def _is_heading(p: Paragraph, tb: TextBlock, body: float) -> bool:
    t = _para_text(p)
    words = t.split()
    if not (1 <= len(words) <= 14) or len(t) < 3 or t.endswith((".", ";", ",")):
        return False
    if CAPTION_RE.match(t):
        return False
    if tb.role in ("title", "subtitle", "heading"):
        return True
    return _size(p) >= body * 1.15 or (_bold(p) and len(tb.paragraphs) > 1)


def _table_text(t: Table, max_rows: int = 14) -> str:
    rows: dict[int, list[tuple[int, str]]] = {}
    for c in t.cells:
        if c.merged_into is not None:
            continue
        txt = ""
        if c.text is not None:
            txt = " ".join(_para_text(p) for p in c.text.paragraphs).strip()
        rows.setdefault(c.row, []).append((c.col, txt))
    lines = [" | ".join(x for _, x in sorted(rows[r])) for r in sorted(rows)]
    more = len(lines) - max_rows
    return "\n".join(lines[:max_rows]) + (f"\n(+{more} rows)" if more > 0 else "")


def _chart_text(ch: Chart) -> str:
    out = [f"{ch.chart_type} chart" + (f" '{ch.title}'" if ch.title else "")]
    if ch.categories:
        out.append("categories: " + ", ".join(ch.categories))
    for s in ch.series:
        vals = ", ".join("" if v is None else f"{v:g}" for v in s.values)
        out.append(f"{s.name or 'series'}: {vals}")
    return "\n".join(out)


def _visual_order(items: list[tuple[float, float, Any]], v: Any) -> float:
    """Sort key of a visual among the page's text blocks: after the last block that
    starts above it."""
    before = [k for (y0, k, _) in items if y0 <= v.bbox[1]]
    return (max(before) if before else -1) + 0.5


def build_outline(doc: Document) -> Outline:
    body = _body_size(doc)
    sections: list[Section] = []
    texts: dict[str, str] = {}
    pages_of: dict[str, int] = {}
    visuals: dict[str, Src] = {}
    title = subtitle = ""
    cur = Section(title="", page=0)
    for page in doc.pages:
        if page.triage.fallback_image:
            continue
        blocks = [e for e in page.elements if isinstance(e, TextBlock) and e.role not in
                  SKIP_ROLES and not _in_master(e)]
        blocks.sort(key=lambda b: (b.reading_order if b.reading_order is not None else 1e9,
                                   b.bbox[1], b.bbox[0]))
        keyed: list[tuple[float, float, Any]] = [(b.bbox[1], float(k), b)
                                                 for k, b in enumerate(blocks)]
        area = max(page.width * page.height, 1.0)
        for e in page.elements:
            if _in_master(e):
                continue
            big = (e.bbox[2] - e.bbox[0]) * (e.bbox[3] - e.bbox[1]) >= 0.02 * area
            if isinstance(e, (Table, Chart)) or (isinstance(e, Image) and big
                                                 and not e.is_background):
                keyed.append((e.bbox[1], _visual_order(keyed, e), e))
        keyed.sort(key=lambda t: (t[1], t[0]))
        captions: list[tuple[TextBlock, str]] = []
        for _, _, e in keyed:
            if isinstance(e, TextBlock):
                paras = [p for p in e.paragraphs if _para_text(p)]
                full = "\n".join(_para_text(p) for p in paras)
                if not full:
                    continue
                texts[e.id], pages_of[e.id] = full, page.index
                if e.role == "caption" or (len(paras) == 1 and CAPTION_RE.match(full)):
                    captions.append((e, full))
                    continue
                if e.role == "title" and not title:
                    title = full.replace("\n", " ")
                    cur = Section(title=title, page=page.index, heading_id=e.id)
                    sections.append(cur)
                    continue
                rest: list[str] = []
                for p in paras:
                    t = _para_text(p)
                    if _is_heading(p, e, body) and not rest:
                        if cur.items or not cur.title or cur.heading_id == e.id:
                            cur = Section(title=t, page=page.index, heading_id=e.id)
                            sections.append(cur)
                        else:  # a heading right below a heading: one longer title
                            cur.title = f"{cur.title}: {t}" if cur.title != title else t
                            cur.heading_id = e.id
                        continue
                    rest.append(t)
                if rest:
                    if not sections:
                        sections.append(cur)
                    cur.items.append(Src(e.id, page.index, "text", "\n".join(rest)))
            else:
                kind = ("table" if isinstance(e, Table) else
                        "chart" if isinstance(e, Chart) else "image")
                txt = (_table_text(e) if isinstance(e, Table) else
                       _chart_text(e) if isinstance(e, Chart) else (e.alt_text or ""))
                s = Src(e.id, page.index, kind, txt)
                visuals[e.id] = s
                texts[e.id], pages_of[e.id] = txt, page.index
                if not sections:
                    sections.append(cur)
                cur.items.append(s)
        for tb, cap in captions:  # nearest visual on the page
            best, dist = None, 1e9
            for v in visuals.values():
                if v.page != page.index:
                    continue
                el = next(x for x in page.elements if x.id == v.id)
                d = min(abs(tb.bbox[1] - el.bbox[3]), abs(el.bbox[1] - tb.bbox[3]))
                if d < dist:
                    best, dist = v, d
            if best is not None and dist < 60 and not best.caption:
                best.caption = cap  # the caption counts as the visual's text (units...)
                texts[best.id] = (cap + "\n" + texts[best.id]).strip()
            else:
                if not sections:
                    sections.append(cur)
                cur.items.append(Src(tb.id, page.index, "text", cap))
    sections = [s for s in sections if s.items or s.heading_id]
    for s in sections:
        if not s.title:
            s.title = title or "Introduction"
    if not title:
        title = (doc.metadata.get("Title") or (sections[0].title if sections else "")
                 or Path(doc.source_path).stem)
    if len(sections) > 1 and sections[0].heading_id and not sections[0].items \
            and sections[0].title == title:
        sections = sections[1:]
    lang = detect_lang(" ".join(texts.values()))
    return Outline(title=title, subtitle=subtitle, lang=lang, pages=len(doc.pages),
                   sections=sections, texts=texts, pages_of=pages_of, visuals=visuals,
                   theme=doc.theme)


def render_sections(sections: list[Section], limit: int = 10**9) -> str:
    out: list[str] = []
    for s in sections:
        out.append(f"## {s.title} (p. {s.page + 1})")
        for i in s.items:
            if i.kind == "text":
                out.append(f"[{i.id}] {i.text}")
            else:
                cap = f' "{i.caption}"' if i.caption else ""
                body = f"\n{i.text}" if i.text else ""
                out.append(f"[{i.id}] {i.kind.upper()}{cap}{body}")
    text = "\n".join(out)
    return text[:limit]


def _split_section(s: Section, budget: int) -> list[Section]:
    parts: list[Section] = []
    cur = Section(s.title, s.page, s.heading_id)
    for it in s.items:
        if cur.items and cur.chars() + len(it.text) > budget:
            parts.append(cur)
            cur = Section(s.title, it.page, s.heading_id)
        if len(it.text) > budget:
            it = dataclasses.replace(it, text=it.text[:budget])
        cur.items.append(it)
    parts.append(cur)
    return parts


def chunks(outline: Outline, budget: int = CHUNK_CHARS) -> list[list[Section]]:
    out: list[list[Section]] = []
    cur: list[Section] = []
    size = 0
    for s in outline.sections:
        for part in (_split_section(s, budget) if s.chars() > budget else [s]):
            if cur and size + part.chars() > budget:
                out.append(cur)
                cur, size = [], 0
            cur.append(part)
            size += part.chars()
    if cur:
        out.append(cur)
    return out


# --------------------------------------------------------------------------- numbers


def _values(tok: str) -> set[float]:
    """Numeric readings of a number token (``1.250,5`` / ``1,250.5`` / ``42%``)."""
    t = tok.strip().rstrip("%‰").lstrip("+")
    out: set[float] = set()
    for thousands, dec in ((",", "."), (".", ",")):
        s = t
        parts = s.split(thousands)
        if len(parts) > 1 and all(len(p) == 3 for p in parts[1:]) and parts[0].lstrip("-"):
            s = "".join(parts)
        s = s.replace(dec, ".")
        if s.count(".") > 1:
            continue
        try:
            out.add(round(float(s), 6))
        except ValueError:
            continue
    return out


# typographic minus signs (U+2212, dashes) before a digit read as "-" (real decks set them)
_MINUS_RE = re.compile(r"[−‒–—﹣－](?=\d)")


def numbers(text: str) -> list[str]:
    text = _MINUS_RE.sub("-", text)
    return [m.group(0) for m in NUM_RE.finditer(text) if any(c.isdigit() for c in m.group(0))]


def number_index(text: str) -> set[float]:
    """Magnitudes of every number in ``text``: "down 23bps" is grounded by "-23bps"."""
    vals: set[float] = set()
    for tok in numbers(text):
        vals |= {abs(v) for v in _values(tok)}
    return vals


def missing_numbers(text: str, source: str | set[float]) -> list[str]:
    idx = number_index(source) if isinstance(source, str) else source
    return [tok for tok in numbers(text) if not ({abs(v) for v in _values(tok)} & idx)]


# --------------------------------------------------------------------------- plan model


@dataclass
class DeckBullet:
    text: str
    sources: list[str]
    pages: list[int] = field(default_factory=list)
    flags: list[str] = field(default_factory=list)


@dataclass
class DeckSlide:
    title: str
    bullets: list[DeckBullet] = field(default_factory=list)
    visual: str | None = None
    notes: str = ""
    kind: str = "content"  # title | content
    source: str = "rule"
    sources: list[str] = field(default_factory=list)  # every block cited on the slide

    def pages(self, outline: Outline) -> list[int]:
        ids = set(self.sources) | ({self.visual} if self.visual else set())
        return sorted({outline.pages_of[i] + 1 for i in ids if i in outline.pages_of})


@dataclass
class DeckPlan:
    title: str
    subtitle: str
    slides: list[DeckSlide]
    lang: str
    source: str
    dropped: list[dict[str, Any]] = field(default_factory=list)
    calls: list[dict[str, Any]] = field(default_factory=list)


def _str_list(ids: list[str], limit: int = 600) -> dict[str, Any]:
    if 0 < len(ids) <= limit:
        return {"type": "array", "items": {"type": "string", "enum": ids}, "maxItems": 4}
    return {"type": "array", "items": {"type": "string"}, "maxItems": 4}


def points_schema(ids: list[str], n: int) -> dict[str, Any]:
    return {"type": "object", "properties": {"points": {
        "type": "array", "maxItems": n, "items": {
            "type": "object",
            "properties": {"section": {"type": "string"}, "text": {"type": "string"},
                           "sources": _str_list(ids)},
            "required": ["section", "text", "sources"]}}},
        "required": ["points"]}


def plan_schema(ids: list[str], visuals: list[str], n: int) -> dict[str, Any]:
    vis: dict[str, Any] = {"type": "string", "enum": ["", *visuals]} if visuals else {
        "type": "string", "enum": [""]}
    return {"type": "object", "properties": {
        "title": {"type": "string"}, "subtitle": {"type": "string"},
        "slides": {"type": "array", "minItems": max(1, n - 1), "maxItems": n + 1, "items": {
            "type": "object", "properties": {
                "title": {"type": "string"},
                "bullets": {"type": "array", "maxItems": MAX_BULLETS, "items": {
                    "type": "object",
                    "properties": {"text": {"type": "string"}, "sources": _str_list(ids)},
                    "required": ["text", "sources"]}},
                "visual": vis,
                "notes": {"type": "string"}},
            "required": ["title", "bullets", "visual", "notes"]}}},
        "required": ["title", "subtitle", "slides"]}


def _brief(n: int, lang: str, audience: str) -> str:
    who = f" The audience is: {audience}." if audience else ""
    return (f"Write the deck in {LANG_NAMES.get(lang, lang)}.{who} Plan exactly {n} content "
            "slides (the title slide is added separately). Each slide: a short title "
            "(at most 8 words), 3 to 5 bullets of at most 15 words each, each bullet citing "
            "the ids of the source blocks it comes from; optionally one visual (the id of a "
            "TABLE, CHART or PICTURE from the source that supports the slide, else \"\"); "
            "speaker notes (2-4 sentences) with the detail behind the bullets. Follow the "
            "document's order and cover all of its sections. Do not use figure or table "
            "captions as bullets. Use only facts and numbers from the source and copy "
            "numbers exactly.")


def plan_prompt(outline: Outline, body: str, n: int, lang: str, audience: str,
                from_points: bool, where: str = "") -> str:
    what = ("Key points extracted from the document (id in brackets; visuals keep their "
            "source ids):" if from_points else "Source document (block id in brackets):")
    part = (f"This is {where}; the other parts get their own slides, so plan slides for "
            "this part only.\n") if where else ""
    return (f"Document title: {outline.title}\nPages: {outline.pages}\n{part}\n{what}\n"
            f"{body}\n\n" + _brief(n, lang, audience) + " Also give a deck title and a "
            "subtitle taken from the document.")


def points_prompt(body: str, n: int, lang: str) -> str:
    return (f"Source document part (block id in brackets):\n{body}\n\nExtract the {n} most "
            f"important key points of this part, in {LANG_NAMES.get(lang, lang)}, in the "
            "document's order. Each point: one sentence (at most 25 words) with the section "
            "it belongs to and the ids of the source blocks it comes from. Keep the key "
            "numbers, copied exactly as written.")


# --------------------------------------------------------------------------- validation


def _content_words(text: str, lang: str) -> list[str]:
    stop = STOP.get(lang, set())
    return [w for w in re.findall(r"\w{4,}", text.lower()) if w not in stop
            and not w.isdigit()]


def overlap(text: str, source: str, lang: str) -> float:
    words = _content_words(text, lang)
    if not words:
        return 1.0
    src = source.lower()
    hit = sum(1 for w in words if w in src or (len(w) > 5 and w[:-2] in src))
    return hit / len(words)


class Grounder:
    """Validates model text against the outline (citations, numbers, overlap)."""

    def __init__(self, outline: Outline, aliases: dict[str, list[str]] | None = None,
                 same_lang: bool = True) -> None:
        self.o = outline
        self.aliases = aliases or {}  # key point id -> block ids
        self.same_lang = same_lang
        self._all = outline.all_text()
        self.doc_numbers = number_index(self._all)
        self.dropped: list[dict[str, Any]] = []

    def expand(self, ids: list[str]) -> list[str]:
        out: list[str] = []
        for i in ids:
            for j in self.aliases.get(i, [i]):
                if j in self.o.texts and j not in out:
                    out.append(j)
        return out

    def _reground(self, text: str) -> list[str]:
        if not self.same_lang:
            return []
        best, score = None, 0.0
        for k, t in self.o.texts.items():
            s = fuzz.token_set_ratio(text.lower(), t.lower())
            if s > score:
                best, score = k, s
        return [best] if best is not None and score >= 70 else []

    def bullet(self, text: str, ids: list[str], where: str) -> DeckBullet | None:
        text = re.sub(r"\s+", " ", text).strip().lstrip("•-–* ").strip()
        if not text:
            return None
        flags: list[str] = []
        src = self.expand(ids)
        if not src:
            src = self._reground(text)
            if not src:
                self.dropped.append({"where": where, "text": text, "reason": "no valid source"})
                return None
            flags.append("regrounded")
        cited = "\n".join(self.o.texts[i] for i in src)
        miss = missing_numbers(text, cited)
        if miss:
            extra = [k for k, t in sorted(self.o.texts.items())
                     if k not in src and any(_values(m) & number_index(t) for m in miss)]
            still = missing_numbers(text, cited + "\n" + "\n".join(self.o.texts[k]
                                                                   for k in extra))
            if still:
                self.dropped.append({"where": where, "text": text,
                                     "reason": f"numbers not in the source: {still}"})
                return None
            src += extra[:3]
            flags.append("numbers cited from " + ", ".join(extra[:3]))
            cited = "\n".join(self.o.texts[i] for i in src)
        if self.same_lang and overlap(text, cited, self.o.lang) < 0.3:
            flags.append("low word overlap with the sources")
        pages = sorted({self.o.pages_of[i] + 1 for i in src})
        return DeckBullet(text=text, sources=src, pages=pages, flags=flags)

    def clean_text(self, text: str, where: str) -> str:
        """Sentences whose numbers are absent from the document are dropped, and (same
        language) sentences whose content words are mostly not in the document: the
        model's own interpretations."""
        keep = []
        for sent in re.split(r"(?<=[.!?])\s+", text.strip()):
            if not sent:
                continue
            miss = missing_numbers(sent, self.doc_numbers)
            if miss:
                self.dropped.append({"where": where, "text": sent,
                                     "reason": f"numbers not in the source: {miss}"})
                continue
            if self.same_lang and overlap(sent, self._all, self.o.lang) < NOTE_OVERLAP:
                self.dropped.append({"where": where, "text": sent,
                                     "reason": "words mostly not in the source"})
                continue
            keep.append(sent)
        return " ".join(keep)

    def title(self, text: str, fallback: str, where: str) -> str:
        text = re.sub(r"\s+", " ", text).strip()
        if not text or missing_numbers(text, self.doc_numbers):
            if text:
                self.dropped.append({"where": where, "text": text,
                                     "reason": "title with numbers not in the source"})
            return fallback
        return text


def _section_of(outline: Outline, ids: list[str]) -> str:
    for s in outline.sections:
        if any(i.id in ids for i in s.items):
            return s.title
    return outline.title


def validate_plan(out: Any, outline: Outline, g: Grounder, n: int, model: str
                  ) -> list[DeckSlide]:
    slides: list[DeckSlide] = []
    if not isinstance(out, dict) or not isinstance(out.get("slides"), list):
        return slides
    for k, s in enumerate(out["slides"][: n + 2]):
        if not isinstance(s, dict):
            continue
        bullets: list[DeckBullet] = []
        for j, b in enumerate(s.get("bullets") or []):
            if not isinstance(b, dict):
                continue
            ids = [str(x) for x in b.get("sources") or [] if isinstance(x, str)]
            db = g.bullet(str(b.get("text") or ""), ids, f"slide {k + 1} bullet {j + 1}")
            if db is not None and len(bullets) < MAX_BULLETS:
                bullets.append(db)
        vis = str(s.get("visual") or "") or None
        if vis is not None and vis not in outline.visuals:
            g.dropped.append({"where": f"slide {k + 1}", "text": vis,
                              "reason": "visual id not in the source"})
            vis = None
        if not bullets and vis is None:
            continue
        cited = [i for b in bullets for i in b.sources]
        title = g.title(str(s.get("title") or ""), _section_of(outline, [*cited, vis or ""]),
                        f"slide {k + 1} title")
        notes = g.clean_text(str(s.get("notes") or ""), f"slide {k + 1} notes")
        slides.append(DeckSlide(title=title, bullets=bullets, visual=vis, notes=notes,
                                source=f"model:{model}",
                                sources=list(dict.fromkeys(cited))))
    return slides


# --------------------------------------------------------------------------- extractive


def _first_sentence(text: str, max_words: int = 22) -> str:
    first = re.split(r"(?<=[.!?;])\s+", text.strip())[0]
    words = first.split()
    if len(words) > max_words:
        return " ".join(words[:max_words]).rstrip(",;:") + " …"
    return first


def _sentences(text: str) -> list[str]:
    return [s for s in re.split(r"(?<=[.!?;])\s+", text.replace("\n", " ").strip()) if s]


def _split_group(items: list[Src]) -> tuple[list[Src], list[Src]]:
    """Two contiguous halves, each with at most one visual where possible."""
    vis = [k for k, i in enumerate(items) if i.kind != "text"]
    cut = vis[1] if len(vis) >= 2 else max(1, len(items) // 2)
    return items[:cut], items[cut:]


def extractive(outline: Outline, n: int, skip: set[str] | None = None,
               only: set[str] | None = None) -> list[DeckSlide]:
    """Rule plan, grounded by construction: per section the first sentence(s) of each
    paragraph and its first visual. Sections are split (largest first) or the ones with
    the least text left out, to approach ``n`` slides."""
    skip = skip or set()
    groups: list[tuple[Section, list[Src]]] = []
    for s in outline.sections:
        items = [i for i in s.items if i.id not in skip and (only is None or i.id in only)
                 and (i.kind != "text" or len(i.text) > 20)]
        for k in range(0, len(items), MAX_BULLETS + 1):
            groups.append((s, items[k:k + MAX_BULLETS + 1]))
    while 0 < len(groups) < n:
        k = max(range(len(groups)), key=lambda j: (len(groups[j][1]), -j))
        if len(groups[k][1]) < 2:
            break
        a, b = _split_group(groups[k][1])
        groups[k:k + 1] = [(groups[k][0], a), (groups[k][0], b)]
    if len(groups) > n:  # keep the n groups with the most text, in document order
        rank = sorted(range(len(groups)), key=lambda k: -sum(len(i.text) for i in groups[k][1]))
        groups = [groups[k] for k in sorted(rank[:n])]
    slides = []
    seen: Counter[str] = Counter()
    for s, items in groups:
        texts = [i for i in items if i.kind == "text"]
        per = 1 if len(texts) >= 3 else (2 if len(texts) == 2 else 4)
        bullets = [DeckBullet(_first_sentence(x), [i.id], [i.page + 1])
                   for i in texts for x in _sentences(i.text)[:per]][:MAX_BULLETS]
        vis = next((i.id for i in items if i.kind != "text"), None)
        notes = " ".join(i.text.replace("\n", " ") for i in texts)
        seen[s.title] += 1
        # a repeated section title is marked as continued, not numbered (no numbers
        # that are not in the source)
        title = s.title if seen[s.title] == 1 else f"{s.title} {CONTINUED[outline.lang]}"
        slides.append(DeckSlide(title=title, bullets=bullets, visual=vis,
                                notes=notes[:1500], source="rule",
                                sources=[i.id for i in texts]))
    return slides


# --------------------------------------------------------------------------- planning


def target_slides(outline: Outline, requested: int | None) -> int:
    if requested:
        return max(2, int(requested))
    return max(4, min(20, 3 + outline.pages // 3))


def plan_deck(outline: Outline, cfg: Config, n_slides: int, *, lang: str = "auto",
              audience: str = "", use_model: bool = True) -> DeckPlan:
    """Plan ``n_slides`` slides in total (title slide included)."""
    lang = outline.lang if lang in ("", "auto") else lang
    n = max(1, n_slides - 1)
    model = _text_model(cfg) if use_model else None
    plan = DeckPlan(title=outline.title, subtitle=outline.subtitle, slides=[], lang=lang,
                    source="rule")
    g = Grounder(outline, same_lang=lang == outline.lang)
    if model is not None:
        try:
            _plan_with_model(outline, cfg, model, n, lang, audience, plan, g)
        except Exception as exc:  # the model layer is optional: extractive plan
            LOG.error("deck planning with %s failed (%s: %s); extractive deck", model,
                      type(exc).__name__, exc, exc_info=LOG.isEnabledFor(10))
            plan.calls.append({"error": f"{type(exc).__name__}: {exc}"})
            plan.slides = []
    if not plan.slides:
        plan.slides = extractive(outline, n)
        plan.source = "rule"
    elif len(plan.slides) < n - 2:  # pad from the sections the model left out
        used = {i for s in plan.slides for i in s.sources}
        extra = extractive(outline, n - len(plan.slides), skip=used)
        LOG.info("deck: %d model slides, %d extractive slides added", len(plan.slides),
                 len(extra))
        plan.slides += extra
    plan.slides = plan.slides[: n + 2]
    plan.dropped = g.dropped
    return plan


def _text_model(cfg: Config) -> str | None:
    try:
        from .max_tier import resolve_stack
        from .setup import llama_server

        if llama_server() is None:
            LOG.warning("deck: llama.cpp missing (pdf2pptx setup llm): extractive deck")
            return None
        name = resolve_stack(cfg, check_memory=True).text
        if name is None:
            LOG.warning("deck: no text model fits in the free RAM: extractive deck")
        return name
    except Exception as exc:
        LOG.warning("deck: model stack unavailable (%s): extractive deck", exc)
        return None


def _plan_with_model(outline: Outline, cfg: Config, model: str, n: int, lang: str,
                     audience: str, plan: DeckPlan, g: Grounder) -> None:
    from .scheduler import Job, Scheduler

    """Plan in contiguous parts that each fit one call; every part gets slides in
    proportion to its length (so a long document is covered to the end). A document
    with more parts than slides is first condensed chunk by chunk into grounded key
    points, and the points are planned the same way."""
    sched = Scheduler(cfg.work_dir / "llm_cache", log_dir=cfg.work_dir / "llm_logs")
    try:
        units = [Unit(render_sections([s]), {i.id for i in s.items}, {
            i.id for i in s.items if i.kind != "text"}, (s.page + 1, max(
                [i.page + 1 for i in s.items] or [s.page + 1])))
            for part in chunks(outline) for s in part]
        total = sum(len(u.text) for u in units)
        from_points = math.ceil(total / SINGLE_PASS_CHARS) > n
        if from_points:
            units = _condense(outline, sched, model, lang, n, plan, g)
        groups = group_units(units, SINGLE_PASS_CHARS, n)
        counts = allocate([sum(len(u.text) for u in grp) for grp in groups], n)
        jobs = []
        for k, (grp, n_k) in enumerate(zip(groups, counts, strict=True)):
            ids = sorted(set().union(*(u.ids for u in grp)) | {
                a for a in g.aliases if any(a in u.ids for u in grp)})
            vis = sorted(set().union(*(u.visuals for u in grp)))
            body = "\n".join(u.text for u in grp)[:SINGLE_PASS_CHARS]
            where = (f"part {k + 1} of {len(groups)} (pages {grp[0].pages[0]}-"
                     f"{grp[-1].pages[1]} of {outline.pages})" if len(groups) > 1 else "")
            jobs.append(sched.submit(Job(
                model=model, prompt=plan_prompt(outline, body, n_k, lang, audience,
                                                from_points, where),
                schema=plan_schema(ids, vis, n_k), system=SYSTEM,
                max_tokens=min(7000, 320 * n_k + 400), tag=f"deck:plan:{k}")))
        runs = sched.run()
        plan.calls += [dataclasses.asdict(r) for r in runs]
    finally:
        sched.close()
    plan.source = f"model:{model}"
    first: dict[str, Any] | None = None
    for k, (job, grp, n_k) in enumerate(zip(jobs, groups, counts, strict=True)):
        res = job.result or {}
        out = res.get("json")
        if out is None or not isinstance(out, dict):
            why = res.get("error") or f"invalid JSON (finish: {res.get('finish_reason')})"
            LOG.warning("deck part %d/%d: %s; extractive slides for it", k + 1, len(groups),
                        why)
            plan.calls.append({"part": k + 1, "error": why})
            only = set().union(*(set(g.expand(sorted(u.ids))) | u.ids for u in grp))
            plan.slides += extractive(outline, n_k, only=only)
            continue
        first = first or out
        plan.slides += validate_plan(out, outline, g, n_k, model)
    if first is not None:
        t = g.title(str(first.get("title") or ""), outline.title, "deck title")
        plan.title = t if overlap(t, outline.all_text(), outline.lang) >= 0.5 or \
            lang != outline.lang else outline.title
        plan.subtitle = g.clean_text(str(first.get("subtitle") or ""), "deck subtitle")[:160]
    else:
        plan.source = "rule"


@dataclass
class Unit:
    """A piece of the planning input: rendered text, the ids it shows, its visuals and
    its page range (1-based)."""

    text: str
    ids: set[str]
    visuals: set[str]
    pages: tuple[int, int]


def group_units(units: list[Unit], budget: int, n: int) -> list[list[Unit]]:
    """Contiguous groups of at most ``budget`` characters (a larger unit is a group of
    its own), then the smallest neighbours merged until there are at most ``n``."""
    groups: list[list[Unit]] = []
    size = 0
    for u in units:
        if groups and size + len(u.text) <= budget:
            groups[-1].append(u)
            size += len(u.text)
        else:
            groups.append([u])
            size = len(u.text)
    while len(groups) > max(n, 1):
        sizes = [sum(len(u.text) for u in g) for g in groups]
        k = min(range(len(groups) - 1), key=lambda i: sizes[i] + sizes[i + 1])
        groups[k:k + 2] = [groups[k] + groups[k + 1]]
    return groups


def allocate(sizes: list[int], n: int) -> list[int]:
    """``n`` slides over parts in proportion to their sizes (largest remainder), at
    least one each."""
    if not sizes:
        return []
    k = len(sizes)
    rest = max(n - k, 0)
    tot = sum(sizes) or 1
    raw = [rest * s / tot for s in sizes]
    out = [1 + math.floor(r) for r in raw]
    left = n - sum(out)
    for i in sorted(range(k), key=lambda i: -(raw[i] - math.floor(raw[i])))[:max(left, 0)]:
        out[i] += 1
    return out


def _condense(outline: Outline, sched: Any, model: str, lang: str, n: int, plan: DeckPlan,
              g: Grounder) -> list[Unit]:
    """Map step for very long documents: grounded key points per chunk (``k<n>`` ids that
    expand to block ids), one unit per chunk."""
    from .scheduler import Job

    parts = chunks(outline)
    per = max(3, min(8, math.ceil(n * 5 / len(parts))))
    jobs = []
    for k, part in enumerate(parts):
        pids = sorted({i.id for s in part for i in s.items})
        jobs.append(sched.submit(Job(
            model=model, prompt=points_prompt(render_sections(part), per, lang),
            schema=points_schema(pids, per), system=SYSTEM, max_tokens=120 * per + 200,
            tag=f"deck:points:{k}")))
    runs = sched.run()
    plan.calls += [dataclasses.asdict(r) for r in runs]
    units: list[Unit] = []
    for k, (job, part) in enumerate(zip(jobs, parts, strict=True)):
        out = (job.result or {}).get("json")
        lines: list[str] = []
        ids: set[str] = set()
        sec_title = ""
        for p in (out or {}).get("points") or []:
            if not isinstance(p, dict):
                continue
            b = g.bullet(str(p.get("text") or ""), [str(x) for x in p.get("sources") or []],
                         f"chunk {k + 1} point")
            if b is None:
                continue
            pid = f"k{len(g.aliases) + 1}"
            g.aliases[pid] = b.sources
            ids.add(pid)
            sec = str(p.get("section") or "")
            if sec and sec != sec_title:
                lines.append(f"## {sec} (p. {b.pages[0]})")
                sec_title = sec
            lines.append(f"[{pid}] {b.text}")
        vis = {i.id for s in part for i in s.items if i.kind != "text"}
        for s in part:  # visuals stay addressable by their own ids
            for i in s.items:
                if i.kind != "text":
                    cap = f' "{i.caption}"' if i.caption else ""
                    lines.append(f"[{i.id}] {i.kind.upper()}{cap} (p. {i.page + 1})")
        if not ids:  # an unusable chunk: the start of its text goes in as is
            lines.append(render_sections(part, 3000))
            ids = {i.id for s in part for i in s.items}
        pages = (part[0].page + 1, max(i.page + 1 for s in part for i in s.items)
                 if any(s.items for s in part) else part[-1].page + 1)
        units.append(Unit("\n".join(lines), ids | vis, vis, pages))
    return units


# --------------------------------------------------------------------------- build


@dataclass
class Box:
    x0: float
    y0: float
    x1: float
    y1: float

    @property
    def t(self) -> tuple[float, float, float, float]:
        return (self.x0, self.y0, self.x1, self.y1)


TITLE = Box(40, 26, 920, 92)
BODY = Box(40, 112, 920, 498)
LEFT = Box(40, 112, 452, 498)
RIGHT = Box(476, 112, 920, 498)
NUM = Box(840, 506, 920, 530)
COVER_TITLE = Box(80, 170, 880, 292)
COVER_SUB = Box(80, 304, 880, 400)
MARGINS = (7.2, 3.6, 7.2, 3.6)


def _theme(outline: Outline) -> Theme:
    t = copy.deepcopy(outline.theme) if outline.theme else Theme()
    defaults = {"dk1": "000000", "lt1": "FFFFFF", "dk2": "1F2A44", "lt2": "E7E6E6",
                "accent1": "2F5597", "accent2": "ED7D31", "accent3": "A5A5A5",
                "accent4": "FFC000", "accent5": "5B9BD5", "accent6": "70AD47",
                "hlink": "0563C1", "folHlink": "954F72"}
    t.colors = {**defaults, **(t.colors or {})}
    t.major_font = t.major_font or "Calibri"
    t.minor_font = t.minor_font or t.major_font
    t.name = t.name or "pdf2pptx deck"
    return t


def _color(theme: Theme, key: str) -> Color:
    return Color(rgb=theme.colors.get(key, "000000"), theme=key)


def _run(text: str, font: str, size: float, color: Color, bold: bool = False) -> TextRun:
    return TextRun(id="", bbox=(0, 0, 0, 0), text=text, font=font, size=size, color=color,
                   bold=bold)


def _block(bid: str, box: Box, paras: list[Paragraph], *, role: str, placeholder: str | None,
           anchor: str = "top") -> TextBlock:
    return TextBlock(id=bid, bbox=box.t, paragraphs=paras, role=role, autofit="shrink",
                     anchor=anchor, margins=MARGINS, placeholder=placeholder,
                     source="rule")


def _estimate_height(texts: list[str], width: float, size: float, gap: float,
                     indent: float = 0.0) -> float:
    """Rough wrapped height (avg glyph width 0.5 em) - the COM fit loop corrects it."""
    cpl = max(8.0, (width - MARGINS[0] - MARGINS[2] - indent) / (0.5 * size))
    lines = sum(max(1, math.ceil(len(t) / cpl)) for t in texts)
    return lines * size * 1.2 + gap * len(texts) + MARGINS[1] + MARGINS[3]


def fit_size(texts: list[str], box: Box, start: float, floor: float, gap_k: float = 0.5,
             indent: float = 0.0) -> float:
    size = start
    while size > floor and _estimate_height(texts, box.x1 - box.x0, size, size * gap_k,
                                            indent) > box.y1 - box.y0:
        size -= 1.0
    return size


def _bullets_block(bid: str, box: Box, bullets: list[DeckBullet], theme: Theme,
                   size: float) -> TextBlock:
    paras = []
    for b in bullets:
        paras.append(Paragraph(
            id="", bbox=(0, 0, 0, 0), runs=[_run(b.text, theme.minor_font or "Calibri", size,
                                                  _color(theme, "dk1"))],
            indent_left=0.9 * size, indent_first=-0.9 * size,
            bullet=Bullet(kind="char", char="•", color=Color(rgb=theme.colors["accent1"])),
            space_before=round(0.5 * size, 1)))
    return _block(bid, box, paras, role="body", placeholder="body")


def _title_block(bid: str, box: Box, text: str, theme: Theme, size: float,
                 placeholder: str = "title", anchor: str = "bottom",
                 align: str = "left", color: str = "dk2", bold: bool = True) -> TextBlock:
    para = Paragraph(id="", bbox=(0, 0, 0, 0), alignment=align,
                     runs=[_run(text, theme.major_font or "Calibri", size, _color(theme, color),
                                bold=bold)])
    return _block(bid, box, [para], role="title" if placeholder == "title" else "subtitle",
                  placeholder=placeholder, anchor=anchor)


def _slide_number(bid: str, k: int, theme: Theme) -> TextBlock:
    r = _run(str(k), theme.minor_font or "Calibri", 12, Color(rgb="7F7F7F"))
    r.field_type = "slidenum"
    para = Paragraph(id="", bbox=(0, 0, 0, 0), alignment="right", runs=[r])
    return _block(bid, NUM, [para], role="page_number", placeholder="sldNum")


# -- copying a source element into a slide box ------------------------------------------


_BOX_FIELDS = {"bbox", "plot_bbox", "legend_bbox"}
_SCALE_FIELDS = {"size", "indent_left", "indent_first", "space_before", "space_after",
                 "char_spacing", "width"}


def _xform(obj: Any, s: float, dx: float, dy: float, prefix: str) -> Any:
    """Scale by ``s`` and shift by (dx, dy) every geometric field of an IR element tree
    (boxes, paths, table rows / columns, font sizes, spacing, margins); ids get
    ``prefix`` so a visual reused on two slides stays unique."""
    if isinstance(obj, list):
        return [_xform(v, s, dx, dy, prefix) for v in obj]
    if not dataclasses.is_dataclass(obj) or isinstance(obj, type):
        return obj
    changes: dict[str, Any] = {}
    for f in dataclasses.fields(obj):
        if not f.init:
            continue
        v = getattr(obj, f.name)
        name = f.name
        if name in _BOX_FIELDS and isinstance(v, tuple) and len(v) == 4:
            changes[name] = (v[0] * s + dx, v[1] * s + dy, v[2] * s + dx, v[3] * s + dy)
        elif name == "pts" and isinstance(v, list):
            changes[name] = [(p[0] * s + dx, p[1] * s + dy) for p in v]
        elif name in ("row_heights", "col_widths") and isinstance(v, list):
            changes[name] = [x * s for x in v]
        elif name == "margins" and isinstance(v, tuple):
            changes[name] = tuple(x * s for x in v)
        elif isinstance(v, (int, float)) and not isinstance(v, bool) and (
                name in _SCALE_FIELDS or (name == "line_spacing" and getattr(
                    obj, "line_spacing_rule", "") == "exact")):
            changes[name] = v * s
        elif name == "style" and isinstance(v, dict):
            changes[name] = {k: (x * s if k.endswith("font_size") and isinstance(x, (int, float))
                                 else x) for k, x in v.items()}
        elif name in ("id",) and isinstance(v, str) and v:
            changes[name] = prefix + v
        elif name in ("child_ids", "fallback_ids") and isinstance(v, list):
            changes[name] = [prefix + x for x in v]
        elif isinstance(v, list) or dataclasses.is_dataclass(v):
            changes[name] = _xform(v, s, dx, dy, prefix)
    return dataclasses.replace(obj, **changes)


MAX_SCALE = {"image": 3.0, "table": 2.2, "chart": 2.5}


def place_visual(el: Any, box: Box, prefix: str) -> Any:
    x0, y0, x1, y1 = el.bbox
    w, h = max(x1 - x0, 1.0), max(y1 - y0, 1.0)
    kind = "table" if isinstance(el, Table) else "chart" if isinstance(el, Chart) else "image"
    s = min((box.x1 - box.x0) / w, (box.y1 - box.y0) / h, MAX_SCALE[kind])
    nx = box.x0 + ((box.x1 - box.x0) - w * s) / 2
    ny = box.y0 + 8.0  # top-aligned with the bullets / title
    out = _xform(copy.deepcopy(el), s, nx - x0 * s, ny - y0 * s, prefix)
    out.notes = [n for n in out.notes if not n.startswith(("in_master:", "in_layout:"))]
    out.notes.append(f"deck:from:{el.id}")
    return out


def _find(doc: Document, eid: str) -> Any:
    for p in doc.pages:
        for e in p.elements:
            if e.id == eid:
                return e
    return None


def _layouts() -> Master:
    lays = [
        Layout(id="L_title", name="Title Slide",
               placeholders={"title": COVER_TITLE.t, "subTitle": COVER_SUB.t}),
        Layout(id="L_content", name="Title and Content",
               placeholders={"title": TITLE.t, "body": BODY.t, "sldNum": NUM.t}),
        Layout(id="L_two", name="Two Content",
               placeholders={"title": TITLE.t, "body": LEFT.t, "body2": RIGHT.t,
                             "sldNum": NUM.t}),
        Layout(id="L_title_only", name="Title Only",
               placeholders={"title": TITLE.t, "sldNum": NUM.t}),
    ]
    return Master(id="M_deck", name="pdf2pptx deck", layouts=lays)


def _accent(theme: Theme, layout: str) -> Shape:
    return Shape(id=f"{layout}_accent", bbox=(TITLE.x0 + 7.2, 96, TITLE.x0 + 127.2, 99.5),
                 preset="rect", fill=Fill(kind="solid", color=_color(theme, "accent1")),
                 source="rule")


def build_deck_document(plan: DeckPlan, outline: Outline, src: Document,
                        sizes: dict[str, float] | None = None) -> Document:
    """The deck as an IR document (one page per slide) for the existing writer.
    ``sizes``: block id -> font size from the fit loop."""
    sizes = sizes or {}
    theme = _theme(outline)
    master = _layouts()
    for lay in master.layouts[1:]:
        lay.elements.append(_accent(theme, lay.id))
    doc = Document(source_path=src.source_path, metadata={
        "Title": plan.title, "Subject": "pdf2pptx deck mode",
        "pdf2pptx:deck": json.dumps({"version": DECK_VERSION, "source": plan.source})},
        fonts={}, theme=theme, masters=[master], stage="build")

    def page(k: int, layout: str, els: list[Any], notes: str = "") -> Page:
        p = Page(index=k, width=SLIDE_W, height=SLIDE_H, elements=els, layout_id=layout,
                 notes_text=notes)
        doc.pages.append(p)
        return p

    cover_t = _title_block("s0_title", COVER_TITLE, plan.title, theme,
                           sizes.get("s0_title") or fit_size([plan.title], COVER_TITLE, 40, 24,
                                                             0.0), anchor="bottom",
                           align="center")
    els: list[Any] = [cover_t]
    if plan.subtitle:
        els.append(_title_block("s0_sub", COVER_SUB, plan.subtitle, theme,
                                sizes.get("s0_sub") or 20, placeholder="subTitle",
                                anchor="top", align="center", color="dk1", bold=False))
    pages = sorted({outline.pages_of[i] + 1 for s in plan.slides for i in s.sources
                    if i in outline.pages_of})
    page(0, "L_title", els, _source_note(plan, outline, pages))
    for k, s in enumerate(plan.slides, start=1):
        tid, bid = f"s{k}_title", f"s{k}_body"
        title_size = sizes.get(tid) or fit_size([s.title], TITLE, 32, 22, 0.0)
        els = [_title_block(tid, TITLE, s.title, theme, title_size)]
        visual = _find(src, s.visual) if s.visual else None
        texts = [b.text for b in s.bullets]
        if visual is not None and s.bullets:
            layout, tbox, vbox = "L_two", LEFT, RIGHT
        elif visual is not None:
            layout, tbox, vbox = "L_title_only", None, Box(BODY.x0, BODY.y0, BODY.x1, BODY.y1)
        else:
            layout, tbox, vbox = "L_content", BODY, None
        if tbox is not None and s.bullets:
            size = sizes.get(bid) or fit_size(texts, tbox, 24 if tbox is BODY else 20, 12,
                                              indent=20)
            els.append(_bullets_block(bid, tbox, s.bullets, theme, size))
        if visual is not None and vbox is not None:
            v = place_visual(visual, vbox, f"s{k}_")
            v.z = 1
            if isinstance(v, Image):
                v.is_background = False
            els.append(v)
        els.append(_slide_number(f"s{k}_num", k + 1, theme))
        page(k, layout, els, _source_note(s, outline, s.pages(outline)))
    return doc


def _source_note(item: DeckSlide | DeckPlan, outline: Outline, pages: list[int]) -> str:
    lang = getattr(item, "lang", outline.lang) if isinstance(item, DeckPlan) else outline.lang
    label = SOURCES_LABEL.get(lang, "Sources")
    refs = ", ".join(str(p) for p in pages)
    notes = item.notes if isinstance(item, DeckSlide) else ""
    tail = f"{label}: {PAGE_LABEL.get(lang, 'p.')} {refs}" if refs else ""
    return "\n\n".join(x for x in (notes, tail) if x)


# --------------------------------------------------------------------------- fit loop


def check_fit(pptx: Path, doc: Document) -> dict[str, dict[str, float]]:
    """PowerPoint text layout of every slide: block id -> {"need", "have"} (pt) for the
    blocks whose text runs past the bottom of their box."""
    from ..build.writer import TEXT_SHAPE_PREFIX
    from ..verify import ppt_com

    boxes = {e.id: e for p in doc.pages for e in p.elements if isinstance(e, TextBlock)}
    over: dict[str, dict[str, float]] = {}
    ppt = ppt_com.shared()
    with ppt.opened(pptx) as pres:
        for i in range(pres.slides):
            for shape in ppt.text_layout(pres, i):
                if not shape.name.startswith(TEXT_SHAPE_PREFIX) or not shape.lines:
                    continue
                tb = boxes.get(shape.name[len(TEXT_SHAPE_PREFIX):])
                if tb is None:
                    continue
                top = min(ln.bound[1] for ln in shape.lines)
                bottom = max(ln.bound[1] + ln.bound[3] for ln in shape.lines)
                right = max(ln.bound[0] + ln.bound[2] for ln in shape.lines)
                have = tb.bbox[3] - tb.bbox[1] - tb.margins[1] - tb.margins[3]
                need = bottom - min(top, tb.bbox[1] + tb.margins[1])
                if need > have + 1.0 or right > tb.bbox[2] + 1.0:
                    over[tb.id] = {"need": round(need, 1), "have": round(have, 1),
                                   "size": max(r.size for p in tb.paragraphs for r in p.runs)}
    return over


def _write(doc: Document, out: Path, work_dir: Path, pdf: Path) -> Any:
    from ..build.writer import TextOptions, write_pptx

    return write_pptx(doc, out, work_dir, pdf, TextOptions.for_mode("deck"))


def fit_loop(plan: DeckPlan, outline: Outline, src: Document, out: Path, work_dir: Path,
             pdf: Path, rounds: int = 4, sizes: dict[str, float] | None = None
             ) -> tuple[Document, Any, dict[str, Any]]:
    """Write, let PowerPoint lay the text out, set overflowing boxes smaller, repeat.
    ``sizes``: starting font sizes by block id (default: the estimate of ``fit_size``)."""
    sizes = dict(sizes or {})
    rep_fit: dict[str, Any] = {"rounds": 0, "shrunk": {}, "overflow": {}}
    doc = build_deck_document(plan, outline, src, sizes)
    rep = _write(doc, out, work_dir, pdf)
    for r in range(rounds):
        try:
            over = check_fit(out, doc)
        except Exception as exc:  # no PowerPoint: the size estimate stands
            rep_fit["skipped"] = f"{type(exc).__name__}: {exc}"
            LOG.warning("deck: PowerPoint fit check unavailable (%s)", exc)
            return doc, rep, rep_fit
        rep_fit["rounds"] = r + 1
        rep_fit["overflow"] = over
        if not over:
            break
        for bid, o in over.items():
            floor = 18.0 if bid.endswith("_title") else 11.0
            new = max(floor, math.floor(o["size"] * min(0.95, (o["have"] / o["need"]) ** 0.5)))
            if new >= o["size"]:
                continue
            sizes[bid] = new
            rep_fit["shrunk"][bid] = new
        if not any(sizes.get(b) for b in over):
            break
        doc = build_deck_document(plan, outline, src, sizes)
        rep = _write(doc, out, work_dir, pdf)
    return doc, rep, rep_fit


# --------------------------------------------------------------------------- entry point


def sources_json(plan: DeckPlan, outline: Outline, fit: dict[str, Any], pdf: Path
                 ) -> dict[str, Any]:
    slides: list[dict[str, Any]] = [{
        "slide": 1, "kind": "title", "title": plan.title, "subtitle": plan.subtitle,
        "source": plan.source}]
    for k, s in enumerate(plan.slides, start=2):
        slides.append({
            "slide": k, "title": s.title, "source": s.source,
            "pages": s.pages(outline),
            "bullets": [{"text": b.text, "sources": b.sources, "pages": b.pages,
                         **({"flags": b.flags} if b.flags else {})} for b in s.bullets],
            "visual": ({"id": s.visual, "kind": outline.visuals[s.visual].kind,
                        "page": outline.visuals[s.visual].page + 1} if s.visual else None),
            "notes_chars": len(s.notes),
        })
    return {"version": DECK_VERSION, "source_pdf": pdf.name, "title": outline.title,
            "lang": plan.lang,
            "plan_source": plan.source, "slides": slides, "dropped": plan.dropped,
            "fit": fit, "outline": outline.as_dict()}


def audit(pdf: Path, pptx: Path, sources: dict[str, Any], src: Document,
          target: int, com: bool = True) -> dict[str, Any]:
    """Independent checks of a written deck (M10 CHECKS), read back from the PPTX:

    * every bullet paragraph of the deck is listed in ``deck_sources.json`` with at least
      one source id, and every id exists in the source IR;
    * every number on the slides and in the notes appears in the source (the PDF text
      layer plus the IR text, which includes OCR); slide numbers and the notes' source
      line are references, not content;
    * no text runs past its box (PowerPoint text layout; skipped without PowerPoint);
    * the slide count is within 2 of the target."""
    from pptx import Presentation

    from ..verify.pdf_text import page_texts

    ids = {e.id for p in src.pages for e in p.elements}
    for p in src.pages:
        for e in p.elements:
            if isinstance(e, Table):
                ids |= {c.text.id for c in e.cells if c.text is not None}
    src_text = "\n".join(page_texts(pdf))
    src_text += "\n" + "\n".join(
        "\n".join(_para_text(q) for q in e.paragraphs) for p in src.pages for e in p.elements
        if isinstance(e, TextBlock))
    for p in src.pages:
        for e in p.elements:
            if isinstance(e, Table):
                src_text += "\n" + _table_text(e, 10**6)
            elif isinstance(e, Chart):
                src_text += "\n" + _chart_text(e)
    idx = number_index(src_text)
    listed = {b["text"]: b for s in sources["slides"] for b in s.get("bullets", [])}
    bad_refs: list[str] = []
    for b in listed.values():
        if not b["sources"] or any(i not in ids for i in b["sources"]):
            bad_refs.append(b["text"])
    prs = Presentation(str(pptx))
    labels = tuple(SOURCES_LABEL.values())
    bullets = 0
    unlisted: list[str] = []
    absent: list[dict[str, Any]] = []
    for k, slide in enumerate(prs.slides, start=1):
        for shape in slide.shapes:
            ph = shape.placeholder_format.type if shape.is_placeholder else None
            if ph is not None and "SLIDE_NUMBER" in str(ph):
                continue
            frames = []
            if shape.has_text_frame:
                frames.append(shape.text_frame)
            if getattr(shape, "has_table", False) and shape.has_table:
                frames += [c.text_frame for r in shape.table.rows for c in r.cells]
            for tf in frames:
                for para in tf.paragraphs:
                    t = "".join(r.text for r in para.runs).strip()
                    if not t:
                        continue
                    if para._p.pPr is not None and para._p.pPr.find(
                            "{http://schemas.openxmlformats.org/drawingml/2006/main}buChar"
                    ) is not None:
                        bullets += 1
                        if t not in listed:
                            unlisted.append(t)
                    for tok in missing_numbers(t, idx):
                        absent.append({"slide": k, "where": "slide", "number": tok, "text": t})
        if slide.has_notes_slide:
            for line in slide.notes_slide.notes_text_frame.text.splitlines():
                if line.startswith(labels):
                    continue
                for tok in missing_numbers(line, idx):
                    absent.append({"slide": k, "where": "notes", "number": tok, "text": line})
    overflow: Any = None
    if com:
        try:
            boxes = Document(pages=[Page(index=0, width=SLIDE_W, height=SLIDE_H, elements=[
                e for p in _deck_ir(pptx) for e in p])])
            overflow = sorted(check_fit(pptx, boxes))
        except Exception as exc:
            overflow = f"unavailable: {type(exc).__name__}: {exc}"
    n = len(prs.slides)
    return {"bullets": bullets, "bullets_with_valid_source": bullets - len(unlisted) - sum(
                1 for t in bad_refs if t not in unlisted),
            "unlisted_bullets": unlisted, "invalid_refs": bad_refs,
            "numbers_absent": absent, "overflow": overflow, "slides": n, "target": target,
            "slide_count_ok": abs(n - target) <= 2,
            "ok": not unlisted and not bad_refs and not absent and overflow in ([], None)
            and abs(n - target) <= 2}


def _deck_ir(pptx: Path) -> list[list[TextBlock]]:
    """Text boxes of a deck as IR blocks (name + box + font size) for ``check_fit``."""
    from pptx import Presentation

    from ..build.writer import TEXT_SHAPE_PREFIX

    out: list[list[TextBlock]] = []
    prs = Presentation(str(pptx))
    emu = 12700.0
    for slide in prs.slides:
        page: list[TextBlock] = []
        for sh in slide.shapes:
            if not sh.has_text_frame or not sh.name.startswith(TEXT_SHAPE_PREFIX):
                continue
            tf = sh.text_frame
            sizes = [r.font.size.pt for p in tf.paragraphs for r in p.runs if r.font.size]
            m = tuple(float(x or 0) / emu for x in (tf.margin_left, tf.margin_top,
                                                     tf.margin_right, tf.margin_bottom))
            x0, y0 = sh.left / emu, sh.top / emu
            page.append(TextBlock(
                id=sh.name[len(TEXT_SHAPE_PREFIX):], bbox=(x0, y0, x0 + sh.width / emu,
                                                         y0 + sh.height / emu),
                margins=(m[0], m[1], m[2], m[3]),
                paragraphs=[Paragraph(id="", bbox=(0, 0, 0, 0), runs=[TextRun(
                    id="", bbox=(0, 0, 0, 0), text="", size=max(sizes, default=18.0))])]))
        out.append(page)
    return out


def convert_deck(pdf: Path, out: Path, cfg: Config, *, use_cache: bool = True,
                 use_model: bool | None = None) -> dict[str, Any]:
    """``--mode deck``: analyse (cached), outline, plan, validate, build, fit."""
    from ..pipeline import analyzed

    t0 = time.perf_counter()
    src, work_dir, cached = analyzed(pdf, cfg, use_cache=use_cache)
    progress.report("plan", 0, 1, "document outline")
    outline = build_outline(src)
    n = target_slides(outline, cfg.deck_slides)
    if use_model is None:
        use_model = cfg.resolved_tier() == "max"
        if not use_model:
            LOG.warning("deck mode needs the Max tier for the text model; writing an "
                        "extractive deck (first sentences per section)")
    plan = plan_deck(outline, cfg, n, lang=cfg.deck_lang, audience=cfg.deck_audience,
                     use_model=use_model)
    progress.report("write", 0, 1, "building and fitting the slides")
    _, rep, fit = fit_loop(plan, outline, src, out, work_dir, pdf)
    srcs = sources_json(plan, outline, fit, pdf)
    sources_path = out.with_name(out.stem + ".deck_sources.json")
    srcs["audit"] = audit(pdf, out, srcs, src, n, com=not fit.get("skipped"))
    sources_path.write_text(json.dumps(srcs, indent=1, ensure_ascii=False) + "\n",
                            encoding="utf-8", newline="\n")
    if not srcs["audit"]["ok"]:
        LOG.warning("%s: deck audit failed: %s", pdf.name,
                    {k: v for k, v in srcs["audit"].items() if v and k in (
                        "unlisted_bullets", "invalid_refs", "numbers_absent", "overflow")})
    return {"pdf": str(pdf), "output": str(out), "sources": str(sources_path),
            "pages": len(src.pages), "slides": rep.slides, "target_slides": n,
            "plan_source": plan.source, "lang": plan.lang, "dropped": len(plan.dropped),
            "flagged": sum(1 for s in plan.slides for b in s.bullets if b.flags),
            "overflow": sorted(fit.get("overflow", {})), "fit": fit, "audit": srcs["audit"],
            "cached_stages": cached, "llm": plan.calls, "errors": rep.errors,
            "seconds": round(time.perf_counter() - t0, 2)}
