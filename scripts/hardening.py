"""M10 hardening checks: big, broken and hostile inputs, and a PC without PowerPoint.

    .venv\\Scripts\\pdf2pptx corpus stress          # generates work/stress/*.pdf
    .venv\\Scripts\\python scripts\\hardening.py [--only NAME ...] [--exe PATH_TO_EXE]

Each case runs the CLI in a child process; the process tree's memory (RSS, summed over
the converter and its workers; POWERPNT.EXE separately) is sampled every 0.25 s. Results
go to bench/hardening.json and bench/hardening.md.
"""

from __future__ import annotations

import argparse
import contextlib
import json
import os
import subprocess
import sys
import threading
import time
from pathlib import Path
from typing import Any

import psutil

ROOT = Path(__file__).resolve().parents[1]
STRESS = ROOT / "work" / "stress"
OUT = ROOT / "out" / "hardening"


def _tree_rss(p: psutil.Process) -> int:
    total = 0
    try:
        procs = [p, *p.children(recursive=True)]
    except psutil.Error:
        return 0
    for q in procs:
        try:
            if (q.name() or "").lower() != "powerpnt.exe":
                total += q.memory_info().rss
        except psutil.Error:
            continue
    return total


def _ppt_rss() -> int:
    tot = 0
    for q in psutil.process_iter(["name"]):
        try:
            if (q.info["name"] or "").lower() == "powerpnt.exe":
                tot += q.memory_info().rss
        except psutil.Error:
            continue
    return tot


def run_case(name: str, args: list[str], *, env: dict[str, str] | None = None,
             cli: list[str]) -> dict[str, Any]:
    t0 = time.perf_counter()
    proc = subprocess.Popen(cli + args, cwd=ROOT, stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE, text=True, encoding="utf-8",
                            errors="replace", env=dict(os.environ, **(env or {})))
    peak = {"tree": 0, "ppt": 0, "avail_min": psutil.virtual_memory().available}
    stop = threading.Event()

    def sample() -> None:
        ps = psutil.Process(proc.pid)
        while not stop.is_set():
            peak["tree"] = max(peak["tree"], _tree_rss(ps))
            peak["ppt"] = max(peak["ppt"], _ppt_rss())
            peak["avail_min"] = min(peak["avail_min"], psutil.virtual_memory().available)
            stop.wait(0.25)

    th = threading.Thread(target=sample, daemon=True)
    th.start()
    out, err = proc.communicate()
    stop.set()
    th.join()
    secs = time.perf_counter() - t0
    res: dict[str, Any] = {"case": name, "rc": proc.returncode, "seconds": round(secs, 1),
                           "peak_rss_gb": round(peak["tree"] / 2**30, 2),
                           "peak_powerpoint_gb": round(peak["ppt"] / 2**30, 2),
                           "lowest_available_gb": round(peak["avail_min"] / 2**30, 2)}
    i = out.find("{")
    if i >= 0:
        with contextlib.suppress(ValueError):
            res["result"] = json.loads(out[i:])
    lines = [ln for ln in err.splitlines() if " INFO " not in ln and ln.strip()]
    res["stderr_tail"] = lines[-6:]
    traceback = any("Traceback" in ln for ln in err.splitlines())
    res["traceback"] = traceback
    return res


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--only", nargs="*")
    ap.add_argument("--exe", help="run the packaged exe instead of the dev build")
    a = ap.parse_args()
    cli = [a.exe] if a.exe else [sys.executable, "-m", "pdf2pptx.cli"]
    OUT.mkdir(parents=True, exist_ok=True)
    junk = OUT / "not_a_pdf.pdf"
    junk.write_text("this is a text file renamed to .pdf\n", encoding="utf-8")
    lite = ["--tier", "lite"]
    cases: list[tuple[str, list[str], dict[str, str] | None, str]] = [
        ("long_report_300p", ["convert", str(STRESS / "long_report_300p.pdf"), "-o",
                              str(OUT / "long_report_300p.pptx"), *lite, "--no-cache"], None,
         "converts; time and peak RAM logged"),
        ("images_50mb", ["convert", str(STRESS / "images_50mb.pdf"), "-o",
                         str(OUT / "images_50mb.pptx"), *lite, "--no-cache"], None,
         "converts; pictures kept"),
        ("encrypted", ["convert", str(STRESS / "encrypted.pdf"), "-o",
                       str(OUT / "encrypted.pptx"), *lite], None,
         "exit code 3, message names the password"),
        ("corrupted", ["convert", str(STRESS / "corrupted.pdf"), "-o",
                       str(OUT / "corrupted.pptx"), *lite], None,
         "exit code 3, message says damaged"),
        ("not_a_pdf", ["convert", str(junk), "-o", str(OUT / "junk.pptx"), *lite], None,
         "exit code 3, message says not a PDF"),
        ("damaged_page", ["convert", str(STRESS / "damaged_page.pdf"), "-o",
                          str(OUT / "damaged_page.pptx"), *lite, "--no-cache"], None,
         "converts; the damaged content stream is skipped, no crash"),
        ("no_powerpoint_convert", ["convert", str(ROOT / "corpus" / "synthetic" /
                                                  "d26_kitchen_sink.pdf"),
                                   "-o", str(OUT / "no_ppt_d26.pptx"), *lite, "--no-cache"],
         {"PDF2PPTX_NO_POWERPOINT": "1"}, "converts without the tuner / checks"),
        ("no_powerpoint_deck", ["convert", str(ROOT / "corpus" / "synthetic" /
                                               "r02_report_tables_it.pdf"),
                                "-o", str(OUT / "no_ppt_deck.pptx"), "--mode", "deck",
                                *lite, "--slides", "4"],
         {"PDF2PPTX_NO_POWERPOINT": "1"}, "extractive deck, fit check skipped"),
        ("no_powerpoint_score", ["score", str(ROOT / "corpus" / "synthetic" /
                                              "d26_kitchen_sink.pdf"),
                                 str(OUT / "no_ppt_d26.pptx")],
         {"PDF2PPTX_NO_POWERPOINT": "1"}, "scores without the PowerPoint metrics"),
    ]
    results = []
    for name, args, env, expect in cases:
        if a.only and name not in a.only:
            continue
        print(f"== {name}: {expect}", flush=True)
        r = run_case(name, args, env=env, cli=cli)
        r["expect"] = expect
        res = r.get("result") or {}
        brief = {k: res.get(k) for k in ("pages", "slides", "seconds", "seconds_per_page",
                                         "fallback_pages", "errors") if k in res}
        if isinstance(res.get("tune"), dict):
            brief["tune"] = res["tune"].get("skipped") or "ran"
        r["brief"] = brief
        print(f"   rc={r['rc']} {r['seconds']} s, peak {r['peak_rss_gb']} GB "
              f"(+PowerPoint {r['peak_powerpoint_gb']} GB), traceback={r['traceback']} "
              f"{brief} {r['stderr_tail'][-2:]}", flush=True)
        results.append(r)
    bench = ROOT / "bench"
    old: list[dict[str, Any]] = []
    jp = bench / "hardening.json"
    if a.only and jp.is_file():
        old = [r for r in json.loads(jp.read_text(encoding="utf-8")) if r["case"] not in a.only]
    allr = sorted(old + [{k: v for k, v in r.items() if k != "result"} for r in results],
                  key=lambda r: r["case"])
    jp.write_text(json.dumps(allr, indent=1, ensure_ascii=False) + "\n", encoding="utf-8",
                  newline="\n")
    md = ["# Hardening (M10)", "", "| case | expected | exit | seconds | peak RAM GB "
          "(converter / all POWERPNT.EXE) | traceback | outcome |", "|---|---|---|---|---|---|---|"]
    for r in allr:
        msg = r["stderr_tail"][-1] if r["stderr_tail"] else ""
        out = r["brief"] or msg
        md.append(f"| {r['case']} | {r['expect']} | {r['rc']} | {r['seconds']} | "
                  f"{r['peak_rss_gb']} / {r['peak_powerpoint_gb']} | {r['traceback']} | "
                  f"{str(out).replace('|', '/')[:220]} |")
    (bench / "hardening.md").write_text("\n".join(md) + "\n", encoding="utf-8", newline="\n")
    print(f"written: {jp}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
