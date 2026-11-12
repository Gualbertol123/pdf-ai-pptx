"""Command line: ``pdf2pptx <command>``."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

from . import __version__, log
from .config import MODES, TIERS, Config, load_config

NOT_YET = {
    "setup": "Milestone 8 (models) / Milestone 9 (llama.cpp)",
}


def _config(args: argparse.Namespace) -> Config:
    overrides: dict[str, Any] = {}
    for key in ("mode", "tier", "workers", "work_dir", "out_dir", "render_dpi", "text_layout",
                "tune", "tune_rounds"):
        v = getattr(args, key, None)
        if v is not None:
            overrides[key] = v
    return load_config(overrides, Path(args.config) if args.config else None)


def cmd_doctor(args: argparse.Namespace) -> int:
    from . import doctor

    cfg = _config(args)
    r = doctor.run(cfg)
    print(json.dumps(r, indent=1) if args.json else doctor.format_report(r))
    return 0


def cmd_corpus(args: argparse.Namespace) -> int:
    from .corpus.build import build

    cfg = _config(args)
    if args.action == "build":
        m = build(cfg.corpus_dir, only=args.only or None)
        cats: dict[str, int] = {}
        for e in m["files"]:
            cats[e["category"]] = cats.get(e["category"], 0) + 1
        print(f"{len(m['files'])} PDFs: " + ", ".join(f"{k}={v}" for k, v in sorted(cats.items())))
        print(f"manifest: {cfg.corpus_dir / 'manifest.json'}")
        return 0
    if args.action == "list":
        from .corpus.build import load_manifest

        for e in load_manifest(cfg.corpus_dir):
            print(f"{e['name']:30s} {e['category']:15s} {e['pages']:3d} pages")
        return 0
    return 2


def cmd_score(args: argparse.Namespace) -> int:
    from .verify import score

    cfg = _config(args)
    if args.corpus:
        s = score.score_corpus(cfg, against=args.against, label=args.label,
                               only=args.only or None)
        print(score.summary_markdown(s))
        print(f"written: {cfg.scores_dir / (args.label or args.against)}")
        if args.as_baseline:
            base = cfg.scores_dir / "baseline.json"
            base.write_text(json.dumps(s, indent=1), encoding="utf-8", newline="\n")
            print(f"baseline: {base}")
        return 0
    if not args.pdf or not args.pptx:
        print("score: give <pdf> <pptx>, or --corpus", file=sys.stderr)
        return 2
    pdf, pptx = Path(args.pdf), Path(args.pptx)
    res = score.score_pair(pdf, pptx, cfg)
    out_dir = cfg.scores_dir / "single"
    out_dir.mkdir(parents=True, exist_ok=True)
    out = out_dir / f"{pdf.stem}.json"
    out.write_text(json.dumps(res, indent=1), encoding="utf-8")
    summary = score.corpus_summary([{**res, "category": "single"}], "single", str(pptx), [])
    md = score.summary_markdown(summary)
    (out_dir / f"{pdf.stem}.md").write_text(md, encoding="utf-8")
    print(md)
    print(f"written: {out}")
    return 0


def cmd_structure(args: argparse.Namespace) -> int:
    from .verify import structure

    cfg = _config(args)
    res = structure.structure_corpus(cfg, only=args.only or None, use_cache=not args.no_cache)
    print(structure.summary_text(res))
    print(f"written: {cfg.scores_dir / 'structure.json'}")
    return 0


def cmd_convert(args: argparse.Namespace) -> int:
    from . import pipeline

    cfg = _config(args)
    use_cache = not args.no_cache
    if args.corpus:
        s = pipeline.convert_corpus(cfg, only=args.only or None, use_cache=use_cache,
                                    check=not args.no_check)
        print(json.dumps(s, indent=1, sort_keys=True))
        print(f"written: {cfg.out_dir / 'corpus'}")
        return 0 if not s["failures"] and s["converted"] == s["files"] else 1
    if not args.pdf:
        print("convert: give <pdf>, or --corpus", file=sys.stderr)
        return 2
    pdf = Path(args.pdf)
    out = Path(args.output) if args.output else pdf.with_suffix(".pptx")
    res = pipeline.convert(pdf, out, cfg, use_cache=use_cache)
    print(json.dumps(res.as_dict(), indent=1))
    return 0


def cmd_report(args: argparse.Namespace) -> int:
    from .report import debug_html

    cfg = _config(args)
    pdfs = [Path(p) for p in args.pdf]
    if args.corpus:
        from .corpus.build import load_manifest

        pdfs += [cfg.corpus_dir / e["pdf"] for e in load_manifest(cfg.corpus_dir)
                 if not args.only or e["name"] in args.only]
    if not pdfs:
        print("report: give <pdf> [<pdf> ...], or --corpus", file=sys.stderr)
        return 2
    if args.pptx and len(pdfs) != 1:
        print("report: --pptx needs exactly one PDF", file=sys.stderr)
        return 2
    out = Path(args.output) if args.output else None
    single = len(pdfs) == 1
    out_dir = (out.parent if single and out else out) or cfg.out_dir / "report"
    done = []
    for pdf in pdfs:
        target = out if single and out else out_dir / f"{pdf.stem}.html"
        path, rep = debug_html.write_report(pdf, cfg, target,
                                            pptx=Path(args.pptx) if args.pptx else None,
                                            use_cache=not args.no_cache)
        err = rep["score"].get("powerpoint_error")
        print(f"{path}  ({len(rep['pages'])} pages, {rep['seconds']} s"
              + (f", PowerPoint unavailable: {err}" if err else "") + ")")
        done.append((path, rep))
    if len(done) > 1:
        index = out_dir / "index.html"
        index.write_text(debug_html.render_index(done, index), encoding="utf-8", newline="\n")
        print(f"index: {index}")
    return 0


def cmd_stub(args: argparse.Namespace) -> int:
    print(f"`pdf2pptx {args.command}` is not implemented yet: arrives in {NOT_YET[args.command]}.",
          file=sys.stderr)
    return 3


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="pdf2pptx", description="Local PDF -> editable PPTX")
    p.add_argument("--version", action="version", version=f"pdf2pptx {__version__}")
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("--config", help="path to pdf2pptx.toml")
    common.add_argument("--work-dir", dest="work_dir")
    common.add_argument("--out-dir", dest="out_dir",
                        help="output root (default out/); corpus files go to <out-dir>/corpus")
    common.add_argument("-v", "--verbose", action="store_true")
    sub = p.add_subparsers(dest="command", required=True)

    c = sub.add_parser("convert", parents=[common], help="convert a PDF to PPTX")
    c.add_argument("pdf", nargs="?")
    c.add_argument("-o", "--output")
    c.add_argument("--mode", choices=MODES)
    c.add_argument("--tier", choices=TIERS)
    c.add_argument("--corpus", action="store_true")
    c.add_argument("--workers", type=int, help="process pool size (default 6)")
    c.add_argument("--only", nargs="*", help="corpus mode: names to convert")
    c.add_argument("--no-cache", dest="no_cache", action="store_true",
                   help="ignore cached IR in work/<sha256>/")
    c.add_argument("--no-check", dest="no_check", action="store_true",
                   help="corpus mode: skip opening each output in PowerPoint")
    c.add_argument("--text-layout", dest="text_layout", choices=("blocks", "lines"),
                   help="blocks (default): one box per text block; lines: the M2 "
                        "baseline, one box per line (for comparison)")
    c.add_argument("--tune", action=argparse.BooleanOptionalAction, default=None,
                   help="fidelity tuner: render in PowerPoint, compare, adjust, re-render "
                        "(default: on for fidelity and balanced, off for editable)")
    c.add_argument("--tune-rounds", dest="tune_rounds", type=int,
                   help="max re-render rounds of the tuner (default 4)")
    c.set_defaults(func=cmd_convert)

    s = sub.add_parser("score", parents=[common], help="score a PPTX against its PDF")
    s.add_argument("pdf", nargs="?")
    s.add_argument("pptx", nargs="?")
    s.add_argument("--corpus", action="store_true", help="score every corpus file")
    s.add_argument("--against", choices=("out", "truth", "blank"), default="out",
                   help="corpus mode: converter output (out/corpus), ground-truth PPTX, "
                        "or blank slides (sanity)")
    s.add_argument("--label", help="name of the result folder under work/scores/")
    s.add_argument("--only", nargs="*", help="corpus names to include")
    s.add_argument("--render-dpi", dest="render_dpi", type=int)
    s.add_argument("--as-baseline", dest="as_baseline", action="store_true",
                   help="also save the summary as work/scores/baseline.json")
    s.set_defaults(func=cmd_score)

    k = sub.add_parser("corpus", parents=[common], help="build or list the synthetic corpus")
    t = sub.add_parser("structure", parents=[common],
                       help="compare paragraphs, levels and fonts with the synthetic truth")
    t.add_argument("--only", nargs="*", help="corpus names to include")
    t.add_argument("--no-cache", dest="no_cache", action="store_true")
    t.set_defaults(func=cmd_structure)

    k.add_argument("action", choices=("build", "list"))
    k.add_argument("--only", nargs="*")
    k.set_defaults(func=cmd_corpus)

    r = sub.add_parser("report", parents=[common], help="HTML debug report")
    r.add_argument("pdf", nargs="*", help="PDF(s) to convert, score and report")
    r.add_argument("-o", "--output", help="HTML file (one PDF) or folder (several); "
                                          "default out/report/")
    r.add_argument("--pptx", help="report on this PPTX instead of converting (one PDF only)")
    r.add_argument("--corpus", action="store_true", help="every corpus file")
    r.add_argument("--only", nargs="*", help="corpus names to include")
    r.add_argument("--no-cache", dest="no_cache", action="store_true")
    r.add_argument("--render-dpi", dest="render_dpi", type=int)
    r.set_defaults(func=cmd_report)

    d = sub.add_parser("doctor", parents=[common], help="check this machine")
    d.add_argument("--json", action="store_true")
    d.set_defaults(func=cmd_doctor)

    u = sub.add_parser("setup", parents=[common], help="download models / tools")
    u.add_argument("what", nargs="?", choices=("models", "llm"))
    u.add_argument("--lite", action="store_true")
    u.set_defaults(func=cmd_stub)
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    log.setup(getattr(args, "verbose", False))
    rc: int = args.func(args)
    return rc


if __name__ == "__main__":
    sys.exit(main())
