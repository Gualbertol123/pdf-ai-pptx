"""Installed-font index (Windows registry) and the metrics the writer needs.

Measured in M2 (PowerPoint 16, zero-inset text box, single spacing): the first
baseline sits ``1.2 * winAscent / (winAscent + winDescent) * size`` below the top
of the text box, for every font tried (Calibri, Arial, Times New Roman, Georgia,
Consolas, Verdana, Segoe UI, Courier New) to within ~0.3% of the size.

M4 (fonts/match.py) builds the real matcher on top of this.
"""

from __future__ import annotations

import functools
import os
import re
import sys
from pathlib import Path

DEFAULT_ASCENT_RATIO = 0.8  # typical winAscent / (winAscent + winDescent)
LINE_FACTOR = 1.2  # PowerPoint single line pitch relative to (ascent + descent)

# PDF base-14 / PostScript names -> Windows families.
ALIASES = {
    "helvetica": "Arial",
    "arialmt": "Arial",
    "times": "Times New Roman",
    "timesroman": "Times New Roman",
    "timesnewromanpsmt": "Times New Roman",
    "timesnewroman": "Times New Roman",
    "courier": "Courier New",
    "couriernewpsmt": "Courier New",
    "couriernew": "Courier New",
    "symbol": "Symbol",
    "zapfdingbats": "Wingdings",
}

_STYLE_WORDS = re.compile(r"\s+(Bold|Italic|Oblique|Regular|Light|Semibold|SemiBold|Black|"
                          r"Medium|Condensed|Semilight|SemiLight)\b.*$")


def _font_dirs() -> list[Path]:
    dirs = [Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"]
    local = os.environ.get("LOCALAPPDATA")
    if local:
        dirs.append(Path(local) / "Microsoft" / "Windows" / "Fonts")
    return dirs


@functools.cache
def installed_fonts() -> dict[str, Path]:
    """Lower-case registry name ("arial", "arial bold", ...) -> font file."""
    out: dict[str, Path] = {}
    if sys.platform != "win32":
        return out
    import winreg

    keys = [(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"),
            (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts")]
    for hive, sub in keys:
        try:
            k = winreg.OpenKey(hive, sub)
        except OSError:
            continue
        i = 0
        while True:
            try:
                name, value, _ = winreg.EnumValue(k, i)
            except OSError:
                break
            i += 1
            path = Path(str(value))
            if not path.is_absolute():
                for d in _font_dirs():
                    if (d / path).is_file():
                        path = d / path
                        break
            label = re.sub(r"\s*\((TrueType|OpenType)\)\s*$", "", str(name))
            for part in label.split(" & "):
                out.setdefault(part.strip().lower(), path)
    return dict(sorted(out.items()))


@functools.cache
def families() -> dict[str, str]:
    """Lower-case family -> display family, for regular faces in the registry."""
    fams: dict[str, str] = {}
    if sys.platform != "win32":
        return fams
    import winreg

    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            k = winreg.OpenKey(hive, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts")
        except OSError:
            continue
        i = 0
        while True:
            try:
                name, _, _ = winreg.EnumValue(k, i)
            except OSError:
                break
            i += 1
            label = re.sub(r"\s*\((TrueType|OpenType)\)\s*$", "", str(name))
            for part in label.split(" & "):
                fam = _STYLE_WORDS.sub("", part.strip())
                fams.setdefault(fam.lower(), fam)
    return dict(sorted(fams.items()))


def _split_camel(name: str) -> str:
    s = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", name)
    return re.sub(r"(?<=[A-Z])(?=[A-Z][a-z])", " ", s)


@functools.cache
def resolve_family(name: str) -> str:
    """Best installed family name for a cleaned PDF family ('TimesNewRoman' -> ...)."""
    key = re.sub(r"[\s_-]", "", name).lower()
    if key in ALIASES:
        return ALIASES[key]
    fams = families()
    for cand in (name, _split_camel(name), name.replace("-", " ")):
        if cand.lower() in fams:
            return fams[cand.lower()]
    for low, disp in fams.items():
        if low.replace(" ", "") == key:
            return disp
    return _split_camel(name)


@functools.cache
def ascent_ratio(family: str) -> float:
    """winAscent / (winAscent + winDescent) of the installed regular face."""
    path = installed_fonts().get(family.lower())
    if path is None or not path.is_file():
        return DEFAULT_ASCENT_RATIO
    try:
        from fontTools.ttLib import TTFont

        kw = {"fontNumber": 0} if path.suffix.lower() == ".ttc" else {}
        f = TTFont(str(path), lazy=True, **kw)  # type: ignore[arg-type]
        os2 = f["OS/2"]
        asc, desc = float(os2.usWinAscent), float(os2.usWinDescent)
        f.close()
        return asc / (asc + desc) if asc + desc > 0 else DEFAULT_ASCENT_RATIO
    except Exception:
        return DEFAULT_ASCENT_RATIO


@functools.cache
def _face_metrics(family: str, bold: bool, italic: bool) -> tuple[dict[int, int], float] | None:
    """(codepoint -> advance in font units, unitsPerEm) of the installed face."""
    style = " ".join(s for s, on in (("bold", bold), ("italic", italic)) if on)
    fonts = installed_fonts()
    path = None
    for name in ([f"{family} {style}".lower()] if style else []) + [family.lower()]:
        if name in fonts:
            path = fonts[name]
            break
    if path is None or not path.is_file():
        return None
    try:
        from fontTools.ttLib import TTFont

        kw = {"fontNumber": 0} if path.suffix.lower() == ".ttc" else {}
        f = TTFont(str(path), lazy=True, **kw)  # type: ignore[arg-type]
        cmap = f.getBestCmap() or {}
        hmtx = f["hmtx"].metrics
        adv = {cp: int(hmtx[g][0]) for cp, g in cmap.items() if g in hmtx}
        upm = float(f["head"].unitsPerEm)
        f.close()
        return adv, upm
    except Exception:
        return None


def advance_em(family: str, bold: bool, italic: bool, ch: str) -> float | None:
    """Advance width of ``ch`` in em for the installed face, or None if unknown."""
    m = _face_metrics(family, bold, italic)
    if m is None or len(ch) != 1:
        return None
    adv, upm = m
    a = adv.get(ord(ch))
    return None if a is None else a / upm


def first_baseline_offset(family: str, size: float) -> float:
    """Distance (pt) from a zero-inset text box top to PowerPoint's first baseline."""
    return LINE_FACTOR * ascent_ratio(family) * size
