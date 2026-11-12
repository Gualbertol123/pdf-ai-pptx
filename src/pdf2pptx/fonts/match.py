"""Map a PDF font to the closest installed font.

Order of evidence (rules first):

1. **Name.** Strip the subset prefix (``ABCDEF+``) and PostScript suffixes (``MT``,
   ``PSMT``), split camel case, and try the family with and without its style words
   ("Calibri-Light" -> "Calibri Light", "SegoeUI-Semibold" -> "Segoe UI Semibold",
   "TimesNewRomanPS-BoldMT" -> "Times New Roman" bold). Base-14 names go through
   metric-compatible aliases (Helvetica -> Arial, Times -> Times New Roman,
   Courier -> Courier New). The width profile is still compared, and a large
   disagreement lowers the confidence.
2. **Width profile.** Otherwise every installed face is scored by the distance
   between the PDF's advance widths of the characters actually used and the face's
   advances (sum |dw| / sum w), plus style penalties (italic, weight, serif and
   fixed-pitch flags). The closest face wins; confidence falls with the distance.
3. **Fallback.** No usable widths: Times New Roman / Courier New / Arial by the PDF
   serif and fixed-pitch flags.

Bold/italic come from the style words in the name, the font flags (italic bit,
ForceBold) and the stem weight pdfium reports; a family that is itself a weight
("Segoe UI Semibold", "Arial Black") is not bolded again.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from ..ir import Document, FontRef
from . import index as fidx

SUBSET_RE = re.compile(r"^[A-Z]{6}\+")
PS_SUFFIX_RE = re.compile(r"(PSMT|PS|MT)$")
STYLE_WORDS = {"bold", "italic", "oblique", "regular", "roman", "book", "normal", "it",
               "bolditalic", "boldoblique", "boldit", "boldmt", "italicmt", "bolditalicmt"}
WEIGHT_FAMILY_WORDS = ("light", "semibold", "semilight", "black", "narrow", "condensed",
                       "medium", "thin", "extralight", "extrabold", "demibold", "heavy")

FLAG_FIXED = 1
FLAG_SERIF = 2
FLAG_ITALIC = 64
FLAG_FORCE_BOLD = 1 << 18

ALIASES = {
    "helvetica": "Arial",
    "helveticaneue": "Arial",
    "arial": "Arial",
    "times": "Times New Roman",
    "timesroman": "Times New Roman",
    "timesnewroman": "Times New Roman",
    "courier": "Courier New",
    "couriernew": "Courier New",
    "symbol": "Symbol",
    "zapfdingbats": "Wingdings",
}

MIN_COMMON_CHARS = 4


@dataclass(frozen=True)
class FontMatch:
    family: str
    bold: bool
    italic: bool
    confidence: float
    source: str  # name | alias | widths | fallback
    distance: float | None = None  # width-profile distance to the chosen face


def strip_subset(name: str) -> str:
    return SUBSET_RE.sub("", name)


def _split_camel(s: str) -> str:
    s = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", s)
    return re.sub(r"(?<=[A-Z])(?=[A-Z][a-z])", " ", s)


def _words(style: str) -> list[str]:
    return [w for w in re.split(r"[\s_,-]+", _split_camel(style)) if w]


def name_parts(pdf_name: str) -> tuple[str, list[str]]:
    """'ABCDEF+TimesNewRomanPS-BoldMT' -> ('TimesNewRoman', ['Bold'])."""
    name = strip_subset(pdf_name)
    m = re.match(r"^(.*?)[-,](.*)$", name)
    base, tail = (m.group(1).strip(), m.group(2).strip()) if m else (name.strip(), "")
    prev = None
    while prev != base:
        prev = base
        stripped = PS_SUFFIX_RE.sub("", base)
        base = stripped if stripped else base
    return base, _words(PS_SUFFIX_RE.sub("", tail))


def _family_candidates(fr: FontRef) -> list[tuple[str, str]]:
    """(candidate family, source) in priority order."""
    base, style = name_parts(fr.pdf_name)
    spaced = _split_camel(base)
    weight_words = [w for w in style if w.lower() in WEIGHT_FAMILY_WORDS]
    out: list[tuple[str, str]] = []
    for b in (spaced, base):
        if weight_words:
            out.append((f"{b} {' '.join(weight_words)}", "name"))
        out.append((b, "name"))
    # a family whose name carries a trailing weight word in the base ("ArialNarrow")
    out.append((_split_camel(fr.family), "name"))
    out.append((fr.family, "name"))
    key = re.sub(r"[\s_-]", "", base).lower()
    if key in ALIASES:
        out.append((ALIASES[key], "alias"))
    seen: set[str] = set()
    uniq = []
    for c, src in out:
        if c and c.lower() not in seen:
            seen.add(c.lower())
            uniq.append((c, src))
    return uniq


def _installed(name: str) -> str | None:
    fam = fidx.installed_family(name)
    if fam:
        return fam
    key = name.replace(" ", "").lower()
    for low, faces in fidx.faces_by_family().items():
        if low.replace(" ", "") == key:
            return faces[0].family if faces[0].family.lower() == low else faces[0].typo_family
    return None


def width_distance(pdf_widths: dict[str, float], face: fidx.Face) -> tuple[float, int]:
    """(sum |w_pdf - w_face| / sum w_pdf over shared chars, number of shared chars)."""
    num = den = 0.0
    n = 0
    for ch, w in pdf_widths.items():
        fw = face.widths.get(ch)
        if fw is None or w <= 0 or ch == " ":
            continue
        num += abs(w - fw)
        den += w
        n += 1
    if n == 0 or den <= 0:
        return 1.0, 0
    return num / den, n


def _style(fr: FontRef) -> tuple[bool, bool, bool]:
    """(bold, italic, name says a weight word) from name, flags and stem weight."""
    _, style = name_parts(fr.pdf_name)
    low = {w.lower() for w in style}
    joined = "".join(low)
    name_bold = bool(low & {"bold", "black", "heavy", "boldmt", "bolditalic", "boldoblique",
                            "boldit", "extrabold", "demibold"}) or "bold" in joined
    name_italic = bool(low & {"italic", "oblique", "it", "italicmt"}) or "italic" in joined
    weight_word = bool(low & set(WEIGHT_FAMILY_WORDS))
    bold = name_bold or (not weight_word and (fr.bold or bool(fr.flags & FLAG_FORCE_BOLD)))
    italic = name_italic or fr.italic or bool(fr.flags & FLAG_ITALIC)
    return bold, italic, weight_word


def _bold_for_family(family: str, bold: bool) -> bool:
    """Do not ask PowerPoint to embolden a family that is itself heavy."""
    if not bold:
        return False
    reg = fidx.best_face(family, False, False)
    return not (reg is not None and reg.weight >= 600 and not reg.bold)


def match_font(fr: FontRef) -> FontMatch:
    bold, italic, _ = _style(fr)
    widths = {k: v for k, v in fr.widths.items() if v > 0}
    for cand, src in _family_candidates(fr):
        fam = _installed(cand)
        if not fam:
            continue
        b = _bold_for_family(fam, bold)
        face = fidx.best_face(fam, b, italic)
        conf = 1.0 if src == "name" else 0.95
        dist = None
        if face is not None and len(widths) >= MIN_COMMON_CHARS:
            dist, n = width_distance(widths, face)
            if n >= MIN_COMMON_CHARS and dist > 0.08:
                conf = min(conf, max(0.5, 1.0 - dist * 4))
        return FontMatch(fam, b, italic, round(conf, 3), src,
                         round(dist, 4) if dist is not None else None)
    return _match_by_widths(fr, widths, bold, italic)


def _match_by_widths(fr: FontRef, widths: dict[str, float], bold: bool,
                     italic: bool) -> FontMatch:
    serif_flag = bool(fr.flags & FLAG_SERIF)
    fixed = bool(fr.flags & FLAG_FIXED)
    best: tuple[float, fidx.Face] | None = None
    if len(widths) >= MIN_COMMON_CHARS:
        want_w = 700 if bold else 400
        for face in fidx.face_index():
            dist, n = width_distance(widths, face)
            if n < min(MIN_COMMON_CHARS, len(widths)):
                continue
            pen = 0.0
            pen += 0.05 if face.italic != italic else 0.0
            pen += min(abs(face.weight - want_w), 400) / 400 * 0.03
            pen += 0.03 if face.serif is not None and face.serif != serif_flag else 0.0
            pen += 0.1 if face.monospace != fixed else 0.0
            pen += 0.01 if face.width != 5 else 0.0
            key = dist + pen
            if best is None or (key, face.family, face.path, face.number) < (
                    best[0], best[1].family, best[1].path, best[1].number):
                best = (key, face)
    if best is not None:
        score, face = best
        b = _bold_for_family(face.family, bold or face.bold)
        conf = max(0.2, min(0.95, 1.0 - score * 6))
        return FontMatch(face.family, b, italic or face.italic, round(conf, 3), "widths",
                         round(score, 4))
    fam = "Courier New" if fixed else ("Times New Roman" if serif_flag else "Arial")
    return FontMatch(fam, bold, italic, 0.3, "fallback", None)


def match_document(doc: Document) -> dict[str, FontMatch]:
    """Match every font of ``doc``; records the result on each FontRef."""
    out: dict[str, FontMatch] = {}
    for fid in sorted(doc.fonts):
        fr = doc.fonts[fid]
        try:
            m = match_font(fr)
        except Exception:  # an unreadable font index never stops a conversion
            m = FontMatch(fidx.resolve_family(fr.family), fr.bold, fr.italic, 0.3, "fallback")
        fr.matched_family = m.family
        fr.match_score = m.confidence
        fr.match_source = m.source
        out[fid] = m
    return out
