"""Installed-font index (Windows registry) and the metrics the writer needs.

Measured in M2 (PowerPoint 16, zero-inset text box, single spacing): the first
baseline sits ``1.2 * winAscent / (winAscent + winDescent) * size`` below the top
of the text box, for every font tried (Calibri, Arial, Times New Roman, Georgia,
Consolas, Verdana, Segoe UI, Courier New) to within ~0.3% of the size.

M4 adds the face index (``face_index()``: every installed face with family, style,
weight, width, fsType and a width profile; cached on disk) used by ``fonts/match.py``.
"""

from __future__ import annotations

import functools
import json
import os
import re
import sys
from dataclasses import dataclass
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


# --------------------------------------------------------------------------- face index (M4)

INDEX_VERSION = 1
# characters whose advances form a face's "width profile" (fonts/match.py)
PROFILE_CHARS = (
    "".join(chr(c) for c in range(0x20, 0x7F))
    + "àèéìòùÀÈÉÌÒÙáíóúçñäöüßâêîôû«»–—‘’“”•…€"
)


@dataclass(frozen=True)
class Face:
    """One installed font face (a TTC member counts as its own face)."""

    path: str
    number: int  # index inside a .ttc
    family: str  # nameID 1 (the name PowerPoint shows, e.g. "Calibri Light")
    typo_family: str  # nameID 16, or the family
    subfamily: str  # nameID 2 / 17
    full_name: str
    ps_name: str
    weight: int
    width: int  # usWidthClass 1..9
    bold: bool
    italic: bool
    monospace: bool
    serif: bool | None  # from PANOSE; None = unknown
    fs_type: int
    win_ascent: int
    win_descent: int
    upm: int
    widths: dict[str, float]  # PROFILE_CHARS present in the face -> advance in em


def index_path() -> Path:
    env = os.environ.get("PDF2PPTX_FONT_INDEX")
    if env:
        return Path(env)
    base = os.environ.get("LOCALAPPDATA") or str(Path.home())
    return Path(base) / "pdf2pptx" / "cache" / "font_index.json"


def _font_files() -> list[Path]:
    out: list[Path] = []
    for d in _font_dirs():
        if d.is_dir():
            out.extend(p for p in d.iterdir()
                       if p.suffix.lower() in (".ttf", ".otf", ".ttc") and p.is_file())
    return sorted(out, key=lambda p: str(p).lower())


def _name(f: object, *ids: int) -> str:
    table = f["name"]  # type: ignore[index]
    for i in ids:
        rec = table.getName(i, 3, 1, 0x409) or table.getName(i, 3, 1) or table.getName(i, 1, 0, 0)
        if rec is None:
            rec = next((r for r in table.names if r.nameID == i), None)
        if rec is not None:
            try:
                s = rec.toUnicode().strip()
            except Exception:
                continue
            if s:
                return s
    return ""


def _read_faces(path: Path) -> list[dict[str, object]]:
    import logging

    from fontTools.ttLib import TTCollection, TTFont

    logging.getLogger("fontTools").setLevel(logging.ERROR)  # noisy timestamp warnings
    fonts = (TTCollection(str(path), lazy=True).fonts if path.suffix.lower() == ".ttc"
             else [TTFont(str(path), lazy=True)])
    out: list[dict[str, object]] = []
    for n, f in enumerate(fonts):
        try:
            if "OS/2" not in f or "hmtx" not in f or "name" not in f:
                continue
            os2, head = f["OS/2"], f["head"]
            upm = int(head.unitsPerEm) or 1000
            cmap = f.getBestCmap() or {}
            hmtx = f["hmtx"].metrics
            widths = {}
            for ch in PROFILE_CHARS:
                g = cmap.get(ord(ch))
                if g is not None and g in hmtx:
                    widths[ch] = round(hmtx[g][0] / upm, 4)
            fam = _name(f, 1)
            sub = _name(f, 2)
            panose = getattr(os2, "panose", None)
            serif = None
            if panose is not None and getattr(panose, "bFamilyType", 0) == 2:
                ss = int(getattr(panose, "bSerifStyle", 0))
                serif = True if 2 <= ss <= 10 else (False if ss >= 11 else None)
            weight = int(os2.usWeightClass)
            sel = int(os2.fsSelection)
            out.append({
                "path": str(path), "number": n, "family": fam,
                "typo_family": _name(f, 16) or fam, "subfamily": _name(f, 17, 2) or sub,
                "full_name": _name(f, 4), "ps_name": _name(f, 6),
                "weight": weight, "width": int(getattr(os2, "usWidthClass", 5)),
                "bold": bool(sel & 0x20) or (sub.lower() in ("bold", "bold italic")),
                "italic": bool(sel & 0x01) or bool(int(head.macStyle) & 0x02),
                "monospace": bool(f["post"].isFixedPitch) if "post" in f else False,
                "serif": serif, "fs_type": int(os2.fsType),
                "win_ascent": int(os2.usWinAscent), "win_descent": int(os2.usWinDescent),
                "upm": upm, "widths": widths,
            })
        except Exception:
            continue
        finally:
            f.close()
    return out


def build_index(path: Path | None = None) -> list[Face]:
    """(Re)index installed fonts, reusing cached entries of unchanged files."""
    path = path or index_path()
    old: dict[str, dict[str, object]] = {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        if data.get("version") == INDEX_VERSION:
            old = {str(k): v for k, v in data.get("files", {}).items()}
    except (OSError, ValueError, AttributeError):
        pass
    files: dict[str, dict[str, object]] = {}
    for fp in _font_files():
        try:
            st = fp.stat()
        except OSError:
            continue
        key = str(fp)
        stamp = f"{st.st_size}:{int(st.st_mtime)}"
        prev = old.get(key)
        if prev and prev.get("stamp") == stamp:
            files[key] = prev
            continue
        try:
            faces = _read_faces(fp)
        except Exception:
            faces = []
        files[key] = {"stamp": stamp, "faces": faces}
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_name(f"{path.name}.{os.getpid()}.tmp")
        tmp.write_text(json.dumps({"version": INDEX_VERSION, "files": files}, sort_keys=True,
                                  ensure_ascii=False), encoding="utf-8", newline="\n")
        tmp.replace(path)
    except OSError:
        pass  # a read-only cache dir only costs a re-index next time
    faces = [Face(**fd) for k in sorted(files) for fd in files[k]["faces"]]  # type: ignore[arg-type,attr-defined]
    return faces


@functools.cache
def face_index() -> tuple[Face, ...]:
    """All installed faces (cached on disk, see ``index_path()``)."""
    return tuple(build_index())


@functools.cache
def faces_by_family() -> dict[str, tuple[Face, ...]]:
    """Lower-case family (nameID 1 and nameID 16) -> faces."""
    out: dict[str, list[Face]] = {}
    for f in face_index():
        for fam in {f.family.lower(), f.typo_family.lower()}:
            if fam:
                out.setdefault(fam, []).append(f)
    return {k: tuple(v) for k, v in sorted(out.items())}


def installed_family(name: str) -> str | None:
    """The installed family spelled as PowerPoint knows it (nameID 1), or None."""
    faces = faces_by_family().get(name.strip().lower())
    if not faces:
        return None
    exact = [f for f in faces if f.family.lower() == name.strip().lower()]
    return exact[0].family if exact else _regular_of(faces).family


def _regular_of(faces: tuple[Face, ...] | list[Face]) -> Face:
    return min(faces, key=lambda f: (f.italic, abs(f.weight - 400), f.width != 5, f.path,
                                     f.number))


def best_face(family: str, bold: bool, italic: bool) -> Face | None:
    """The face of ``family`` closest to the requested style."""
    faces = [f for f in faces_by_family().get(family.lower(), ()) if f.family.lower()
             == family.lower()] or list(faces_by_family().get(family.lower(), ()))
    if not faces:
        return None
    want_w = 700 if bold else 400
    return min(faces, key=lambda f: (f.italic != italic, abs(f.weight - want_w), f.width != 5,
                                     f.path, f.number))
