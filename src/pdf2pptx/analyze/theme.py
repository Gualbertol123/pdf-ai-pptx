"""Theme colours and fonts (M7).

**Palette.** Every colour the page draws is weighted by how much of it there is:
shape fills by area (gradient stops share their shape's area), lines by length x
width, text runs by characters x size², table cell fills by area, bullets like text.
Near-identical colours (RGB distance <= 6) merge into the heavier one. Backgrounds
count only for ``lt1`` / ``dk1`` (dark slides make a dark ``lt1`` wrong: the light
end is taken from all colours).

* ``dk1`` = the heaviest dark (luma < 0.25) text colour, else the heaviest dark
  colour, else 000000; ``lt1`` = the heaviest light (luma > 0.9) background, else the
  heaviest light colour, else FFFFFF; ``dk2`` / ``lt2`` = the next dark / light
  colours.
* ``accent1..6`` = the heaviest chromatic colours (saturation >= 0.2 or a mid grey
  that a known palette uses), by weight. If at least two of the document's colours
  belong to one of PowerPoint's built-in palettes (Office 2007-2010, 2013-2022,
  2023+), that palette's accents are used in its own order and the remaining slots
  keep the document's colours: a deck made from a stock theme gets its theme back.
* ``hlink`` = the heaviest colour of hyperlinked runs (PowerPoint draws links in the
  theme's hyperlink colour), ``folHlink`` the built-in default.

**Fonts.** ``major`` = the font family with the most title characters, ``minor`` =
the one with the most body characters (falling back to all text).

**References.** Colours equal to a theme colour (exact RGB) get ``Color.theme``
(``rgb`` is kept), so the writer can emit ``schemeClr`` in editable / balanced mode.
The writer maps dk1/lt1/dk2/lt2 to tx1/bg1/tx2/bg2.
"""

from __future__ import annotations

import colorsys
import json
from collections import Counter
from collections.abc import Iterator
from typing import Any

from ..ir import Color, Document, Fill, Shape, Stroke, Table, TextBlock, Theme

THEME_KEY = "pdf2pptx:theme"
KNOWN_PALETTES: dict[str, dict[str, str]] = {
    "Office 2007-2010": {"dk2": "1F497D", "lt2": "EEECE1", "accent1": "4F81BD",
                         "accent2": "C0504D", "accent3": "9BBB59", "accent4": "8064A2",
                         "accent5": "4BACC6", "accent6": "F79646", "hlink": "0000FF",
                         "folHlink": "800080"},
    "Office 2013-2022": {"dk2": "44546A", "lt2": "E7E6E6", "accent1": "4472C4",
                         "accent2": "ED7D31", "accent3": "A5A5A5", "accent4": "FFC000",
                         "accent5": "5B9BD5", "accent6": "70AD47", "hlink": "0563C1",
                         "folHlink": "954F72"},
    "Office 2013": {"dk2": "44546A", "lt2": "E7E6E6", "accent1": "5B9BD5",
                    "accent2": "ED7D31", "accent3": "A5A5A5", "accent4": "FFC000",
                    "accent5": "4472C4", "accent6": "70AD47", "hlink": "0563C1",
                    "folHlink": "954F72"},
    "Office 2023": {"dk2": "0E2841", "lt2": "E8E8E8", "accent1": "156082",
                    "accent2": "E97132", "accent3": "196B24", "accent4": "0F9ED5",
                    "accent5": "A02B93", "accent6": "4EA72E", "hlink": "467886",
                    "folHlink": "96607D"},
}
DEFAULT_PALETTE = "Office 2013-2022"
ACCENTS = tuple(f"accent{i}" for i in range(1, 7))
SLOTS = ("dk1", "lt1", "dk2", "lt2", *ACCENTS, "hlink", "folHlink")
MERGE_DIST = 6.0


def _rgb(h: str) -> tuple[int, int, int]:
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def _dist(a: str, b: str) -> float:
    x, y = _rgb(a), _rgb(b)
    return float(sum((u - v) ** 2 for u, v in zip(x, y, strict=True)) ** 0.5)


def _luma(h: str) -> float:
    r, g, b = (v / 255 for v in _rgb(h))
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def _sat(h: str) -> float:
    r, g, b = (v / 255 for v in _rgb(h))
    return colorsys.rgb_to_hsv(r, g, b)[1]


def _ok(c: Color | None) -> str | None:
    if c is None or not c.rgb or len(c.rgb) != 6 or c.alpha < 0.05:
        return None
    return c.rgb.upper()


# --------------------------------------------------------------------------- collection


def _area(b: tuple[float, float, float, float]) -> float:
    return max(b[2] - b[0], 0.0) * max(b[3] - b[1], 0.0)


def _blocks(doc: Document) -> Iterator[TextBlock]:
    for p in doc.pages:
        for e in p.elements:
            if isinstance(e, TextBlock):
                yield e
            elif isinstance(e, Shape) and e.text is not None:
                yield e.text
            elif isinstance(e, Table):
                for c in e.cells:
                    if c.text is not None:
                        yield c.text


def collect(doc: Document) -> tuple[Counter[str], Counter[str], Counter[str], Counter[str]]:
    """(weights of drawn colours, background colours, hyperlink run colours, text
    colours)."""
    w: Counter[str] = Counter()
    bg: Counter[str] = Counter()
    links: Counter[str] = Counter()
    text: Counter[str] = Counter()

    def fill(f: Fill, area: float, into: Counter[str]) -> None:
        if f.kind == "solid" and (h := _ok(f.color)):
            into[h] += area
        elif f.kind == "gradient" and f.gradient and f.gradient.stops:
            for s in f.gradient.stops:
                if h := _ok(s.color):
                    into[h] += area / len(f.gradient.stops)

    def stroke(s: Stroke | None, length: float) -> None:
        if s is not None and (h := _ok(s.color)):
            w[h] += length * max(s.width, 0.25)

    for p in doc.pages:
        if p.triage.fallback_image:
            continue
        for e in p.elements:
            if isinstance(e, Shape):
                b = e.bbox
                fill(e.fill, _area(b), bg if e.is_background else w)
                stroke(e.stroke, 2 * ((b[2] - b[0]) + (b[3] - b[1])) if e.preset != "line"
                       else ((b[2] - b[0]) ** 2 + (b[3] - b[1]) ** 2) ** 0.5)
            elif isinstance(e, Table):
                for c in e.cells:
                    if c.merged_into is None:
                        fill(c.fill, _area(e.bbox) / max(len(e.cells), 1), w)
                    for s in c.borders.values():
                        stroke(s, 20.0)
    for tb in _blocks(doc):
        for para in tb.paragraphs:
            for r in para.runs:
                h = _ok(r.color)
                n = len(r.text.strip())
                if h and n:
                    w[h] += n * r.size * r.size
                    text[h] += n * r.size * r.size
                    if r.hyperlink and not r.hyperlink.startswith("#"):
                        links[h] += n
            if para.bullet.color is not None and (h := _ok(para.bullet.color)):
                w[h] += 50.0
    return w, bg, links, text


def _merge(w: Counter[str]) -> list[tuple[str, float]]:
    out: list[list[Any]] = []
    for h, v in sorted(w.items(), key=lambda kv: (-kv[1], kv[0])):
        for o in out:
            if _dist(o[0], h) <= MERGE_DIST:
                o[1] += v
                break
        else:
            out.append([h, v])
    return [(str(h), float(v)) for h, v in sorted(out, key=lambda o: (-o[1], o[0]))]


# --------------------------------------------------------------------------- theme


def _known_palette(colors: list[str]) -> str | None:
    best, best_n = None, 1
    for name, pal in KNOWN_PALETTES.items():
        accents = {pal[k] for k in ACCENTS}
        n = sum(1 for c in colors if c in accents)
        if n > best_n:
            best, best_n = name, n
    return best


def build_theme(doc: Document) -> tuple[Theme, dict[str, Any]]:
    w, bg, links, text = collect(doc)
    merged = _merge(w)
    every = _merge(w + bg)
    colors = [h for h, _ in merged]
    dark_text = [h for h, _ in _merge(text) if _luma(h) < 0.25]
    light_bg = [h for h, _ in _merge(bg) if _luma(h) > 0.9]
    darks = dark_text + [h for h, _ in every if _luma(h) < 0.25 and h not in dark_text]
    lights = light_bg + [h for h, _ in every if _luma(h) > 0.9 and h not in light_bg]
    pal_name = _known_palette(colors)
    pal = KNOWN_PALETTES[pal_name or DEFAULT_PALETTE]
    pal_accents = [pal[k] for k in ACCENTS]
    out: dict[str, str] = {}
    out["dk1"] = darks[0] if darks else "000000"
    out["lt1"] = lights[0] if lights else "FFFFFF"
    out["dk2"] = next((h for h in darks[1:] if _dist(h, out["dk1"]) > 30), pal["dk2"])
    out["lt2"] = next((h for h in lights[1:] if _dist(h, out["lt1"]) > 10), pal["lt2"])
    chroma = [h for h in colors if (_sat(h) >= 0.2 and 0.08 < _luma(h) < 0.95)
              or h in pal_accents]
    if pal_name:
        accents = list(pal_accents)
        extra = [h for h in chroma if h not in accents]
        # palette slots the document never uses may hold its other colours
        unused = [k for k, h in enumerate(accents) if h not in colors]
        for k, h in zip(unused[::-1], extra, strict=False):
            accents[k] = h
    else:
        accents = chroma[:6] + [h for h in pal_accents if h not in chroma][:max(0, 6 - len(
            chroma))]
    for k, h in zip(ACCENTS, accents, strict=True):
        out[k] = h
    out["hlink"] = links.most_common(1)[0][0] if links else pal["hlink"]
    out["folHlink"] = pal["folHlink"]
    major, minor = _fonts(doc)
    theme = Theme(name="pdf2pptx", colors={k: out[k] for k in SLOTS}, major_font=major,
                  minor_font=minor)
    info = {"palette": pal_name, "weights": [[h, round(v, 1)] for h, v in merged[:12]]}
    return theme, info


def _fonts(doc: Document) -> tuple[str | None, str | None]:
    title: Counter[str] = Counter()
    body: Counter[str] = Counter()
    for tb in _blocks(doc):
        into = title if tb.role in ("title", "subtitle") else body
        for para in tb.paragraphs:
            for r in para.runs:
                if r.font and r.text.strip():
                    into[r.font] += len(r.text.strip())
    allf = title + body

    def top(c: Counter[str]) -> str | None:
        return min(c.items(), key=lambda kv: (-kv[1], kv[0]))[0] if c else None

    return top(title) or top(allf), top(body) or top(allf)


def _refs(doc: Document, theme: Theme) -> int:
    by_rgb: dict[str, str] = {}
    for k in SLOTS:  # the first slot wins (dk1 before accents)
        if k in ("hlink", "folHlink"):
            continue
        by_rgb.setdefault(theme.colors[k], k)
    n = 0

    def ref(c: Color | None) -> None:
        nonlocal n
        if c is not None and c.rgb and (k := by_rgb.get(c.rgb.upper())):
            c.theme = k
            n += 1

    def fill(f: Fill) -> None:
        ref(f.color)
        if f.gradient:
            for s in f.gradient.stops:
                ref(s.color)

    for p in doc.pages:
        for e in p.elements:
            if isinstance(e, Shape):
                fill(e.fill)
                ref(e.stroke.color if e.stroke else None)
            elif isinstance(e, Table):
                for c in e.cells:
                    fill(c.fill)
                    for s in c.borders.values():
                        ref(s.color)
    for tb in _blocks(doc):
        for para in tb.paragraphs:
            ref(para.bullet.color)
            for r in para.runs:
                ref(r.color)
    return n


def analyze_theme(doc: Document) -> dict[str, int]:
    theme, info = build_theme(doc)
    doc.theme = theme
    n = _refs(doc, theme)
    info["references"] = n
    doc.metadata[THEME_KEY] = json.dumps(info, sort_keys=True)
    return {"theme_references": n}
