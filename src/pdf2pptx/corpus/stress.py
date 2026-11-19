"""Stress / hardening inputs (M10), generated on demand into ``work/stress`` (not part
of the scored corpus, not in git):

* ``long_report_40p.pdf``: a 40-page report (sections, headings, numbers, tables,
  figures): the deck-mode showcase.
* ``long_report_300p.pdf``: the same generator, ~300 pages (time and peak RAM).
* ``images_50mb.pdf``: 13 pages of large photographic JPEGs, ~50 MB.
* ``encrypted.pdf``: password protected (user password ``secret``).
* ``corrupted.pdf``: a PDF header followed by damaged bytes.
* ``damaged_page.pdf``: a valid PDF with one page's content stream destroyed.

Deterministic (fixed seeds, reportlab ``invariant=1``).
"""

from __future__ import annotations

import random
from pathlib import Path
from typing import Any

import cv2
import numpy as np
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas
from reportlab.platypus import Image, PageBreak, Spacer

from . import assets
from .reports import _doc, _p, _styles, _table

TOPICS = [
    ("Program overview", "program", ["sites", "coordinators", "volunteers", "sessions"]),
    ("Participation", "participation", ["children", "families", "sessions", "visits"]),
    ("Reading outcomes", "outcomes", ["pupils tested", "reading level gains", "book loans"]),
    ("Volunteers and staff", "staff", ["volunteers", "training hours", "staff members"]),
    ("Partnerships", "partners", ["schools", "libraries", "local businesses"]),
    ("Budget and spending", "budget", ["euro for materials", "euro for staff", "euro for rent"]),
    ("Fundraising", "funds", ["donors", "grants", "events"]),
    ("Communication", "comms", ["newsletter subscribers", "social media posts", "press articles"]),
    ("Digital learning", "digital", ["tablets", "online sessions", "app users"]),
    ("Risks and lessons learned", "risks", ["cancelled sessions", "late reports", "complaints"]),
    ("Evaluation method", "method", ["interviews", "survey responses", "observations"]),
    ("Plans for next year", "plans", ["new sites", "new volunteers", "planned sessions"]),
]
OPENERS = [
    "In {year} the {topic} work reached {n1} {u1}, up from {n2} the year before.",
    "The team recorded {n1} {u1} and {n2} {u2} across all sites.",
    "Compared with the target of {n2}, the program reached {n1} {u1}.",
    "Most of the growth came from the two newest sites, which added {n1} {u1}.",
    "Spending on {topic} stayed within the plan, at {n1} {u1} against {n2} budgeted.",
]
MIDDLES = [
    "Coordinators reported that evening sessions were easier for working parents to join.",
    "The local library continued to host sessions and promoted them in its newsletter.",
    "Feedback from families was positive, especially about the small group format.",
    "Two sites changed their schedule in spring to avoid the school holidays.",
    "Volunteers received a short training course before starting their first session.",
    "The steering group reviewed the results every quarter and adjusted the plan.",
    "Some sessions were cancelled because of bad weather and staff illness.",
    "The partnership with schools made it easier to reach families who had not joined before.",
]
CLOSERS = [
    "About {pct}% of the families who joined in autumn were still attending in June.",
    "The share of sessions held in the evening rose to {pct}%.",
    "Overall, {pct}% of participants said they would recommend the program.",
    "The cost per participant fell to {n3} euro.",
    "We expect to keep this level of activity next year.",
]


def _para(rng: random.Random, topic: str, units: list[str], year: int) -> str:
    n2 = rng.randint(20, 400)
    vals = {"year": year, "topic": topic, "n1": n2 + rng.randint(5, 120), "n2": n2,
            "u1": rng.choice(units), "u2": rng.choice(units), "pct": rng.randint(35, 95),
            "n3": rng.randint(12, 90)}
    parts = [rng.choice(OPENERS), *rng.sample(MIDDLES, 2), rng.choice(CLOSERS)]
    return " ".join(p.format(**vals) for p in parts)


def long_report(path: Path, asset_dir: Path, sections: int = 12, paras: int = 43,
                seed: int = 10) -> Path:
    """About ``sections * paras / 12.7`` pages."""
    rng = random.Random(seed)
    st = _styles()
    doc = _doc(path, columns=1)
    story: list[Any] = [_p("Community Reading Program: Annual Report 2026", st["h1"], "title"),
                        _p("This report describes the activities, results and finances of "
                           "the program in 2026 and the plans for 2027.", st["body"])]
    diagram = assets.diagram(asset_dir / "diagram.png")
    for k in range(sections):
        title, topic, units = TOPICS[k % len(TOPICS)]
        name = title if k < len(TOPICS) else f"{title} ({k // len(TOPICS) + 1})"
        story.append(_p(f"{k + 1}. {name}", st["h2"], "heading"))
        for j in range(paras):
            story.append(_p(_para(rng, topic, units, 2026), st["body"]))
            if j == 5 and k % 2 == 0:
                rows = [["Site", "2024", "2025", "2026"]]
                for site in ("Centre", "North", "South", "East"):
                    a = rng.randint(40, 200)
                    rows.append([site, str(a), str(a + rng.randint(0, 40)),
                                 str(a + rng.randint(20, 80))])
                story.append(_p(f"Table {k + 1}. {title}: {units[0]} per site",
                                st["caption"], "caption"))
                story.append(_table(rows, ruled=True))
                story.append(Spacer(1, 10))
            if j == 10 and k % 3 == 1:
                story.append(Image(str(diagram), width=11 * cm, height=6.6 * cm))
                story.append(_p(f"Figure {k + 1}. {title}: sessions per month.",
                                st["caption"], "caption"))
        if k % 4 == 3:
            story.append(PageBreak())
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.build(story)
    return path


def _photo_jpeg(path: Path, seed: int, w: int = 2600, h: int = 1950) -> Path:
    rng = np.random.default_rng(seed)
    base = cv2.resize(rng.integers(0, 255, (h // 40, w // 40, 3), dtype=np.uint8), (w, h),
                      interpolation=cv2.INTER_CUBIC)
    grain = rng.normal(0, 28, (h, w, 3))
    img = np.clip(base.astype(np.float64) + grain, 0, 255).astype(np.uint8)
    cv2.imwrite(str(path), img, [cv2.IMWRITE_JPEG_QUALITY, 92])
    return path


def images_pdf(path: Path, tmp: Path, pages: int = 13) -> Path:
    tmp.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(path), pagesize=A4, invariant=1)
    for i in range(pages):
        jpg = _photo_jpeg(tmp / f"photo_{i:02d}.jpg", seed=100 + i)
        c.setFont("Helvetica-Bold", 16)
        c.drawString(2 * cm, A4[1] - 2 * cm, f"Photo {i + 1}")
        c.drawImage(str(jpg), 2 * cm, 4 * cm, width=A4[0] - 4 * cm, height=(A4[0] - 4 * cm) * 0.75)
        c.showPage()
    c.save()
    return path


def encrypted_pdf(path: Path) -> Path:
    from reportlab.lib.pdfencrypt import StandardEncryption

    enc = StandardEncryption("secret", ownerPassword="owner", canPrint=1)
    c = canvas.Canvas(str(path), pagesize=A4, encrypt=enc, invariant=1)
    c.setFont("Helvetica", 14)
    c.drawString(2 * cm, A4[1] - 3 * cm, "Confidential: this page is password protected.")
    c.showPage()
    c.save()
    return path


def corrupted_pdf(path: Path, source: Path) -> Path:
    data = bytearray(source.read_bytes())
    rng = random.Random(3)
    keep = 400  # the header and the first object survive
    for i in range(keep, len(data)):
        data[i] = rng.randrange(256)
    path.write_bytes(bytes(data[: len(data) // 2]))
    return path


def damaged_page_pdf(path: Path, source: Path) -> Path:
    """A valid file whose second page's content stream is garbage (same length)."""
    import pypdfium2 as pdfium

    src = pdfium.PdfDocument(str(source))
    try:
        src.save(str(path))
    finally:
        src.close()
    data = bytearray(path.read_bytes())
    marks = [i for i in range(len(data)) if data.startswith(b"stream", i)]
    if len(marks) >= 2:
        start = marks[1] + 8
        rng = random.Random(5)
        for i in range(start, min(start + 400, len(data))):
            if data[i:i + 9] == b"endstream":
                break
            data[i] = rng.randrange(256)
    path.write_bytes(bytes(data))
    return path


def build(out_dir: Path, corpus_dir: Path, only: list[str] | None = None) -> dict[str, Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    assets_dir = out_dir / "_assets"
    assets_dir.mkdir(parents=True, exist_ok=True)
    jobs = {
        "long_report_40p": lambda p: long_report(p, assets_dir),
        "long_report_300p": lambda p: long_report(p, assets_dir, sections=96, paras=41),
        "images_50mb": lambda p: images_pdf(p, assets_dir / "photos"),
        "encrypted": encrypted_pdf,
        "corrupted": lambda p: corrupted_pdf(
            p, corpus_dir / "synthetic" / "r01_report_two_column.pdf"),
        "damaged_page": lambda p: damaged_page_pdf(
            p, corpus_dir / "synthetic" / "d01_title_bullets.pdf"),
    }
    out: dict[str, Path] = {}
    for name, fn in jobs.items():
        if only and name not in only:
            continue
        p = out_dir / f"{name}.pdf"
        out[name] = fn(p)
    return out
