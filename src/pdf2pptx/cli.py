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
    "convert": "Milestone 2 (ingest + baseline converter)",
    "report": "Milestone 3 (debug report)",
    "setup": "Milestone 8 (models) / Milestone 9 (llama.cpp)",
}


def _config(args: argparse.Namespace) -> Config:
    overrides: dict[str, Any] = {}
    for key in ("mode", "tier", "workers", "work_dir", "render_dpi"):
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
    common.add_argument("-v", "--verbose", action="store_true")
    sub = p.add_subparsers(dest="command", required=True)

    c = sub.add_parser("convert", parents=[common], help="convert a PDF to PPTX")
    c.add_argument("pdf", nargs="?")
    c.add_argument("-o", "--output")
    c.add_argument("--mode", choices=MODES)
    c.add_argument("--tier", choices=TIERS)
    c.add_argument("--corpus", action="store_true")
    c.add_argument("--workers", type=int)
    c.set_defaults(func=cmd_stub)

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
    s.set_defaults(func=cmd_score)

    k = sub.add_parser("corpus", parents=[common], help="build or list the synthetic corpus")
    k.add_argument("action", choices=("build", "list"))
    k.add_argument("--only", nargs="*")
    k.set_defaults(func=cmd_corpus)

    r = sub.add_parser("report", parents=[common], help="HTML debug report")
    r.add_argument("pdf", nargs="?")
    r.set_defaults(func=cmd_stub)

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
