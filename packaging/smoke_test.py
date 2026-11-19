"""Smoke test of the packaged exe (M10): convert 3 corpus files with the built exe and
with the dev build, score both against the PDF with the dev scorer, and require the same
scores (within 0.5 %).

    .venv\\Scripts\\python packaging\\smoke_test.py [--exe dist\\pdf2pptx\\pdf2pptx.exe]
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
FILES = ["d01_title_bullets", "d26_kitchen_sink", "r02_report_tables_it"]
KEYS = ["ssim", "text_sim", "real_text_ratio", "linebreak_match", "native_shape_ratio",
        "composite"]
TOL = 0.005


def run(cmd: list[str], **kw: object) -> tuple[int, float, str]:
    t0 = time.perf_counter()
    p = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                       errors="replace", **kw)  # type: ignore[call-overload]
    return p.returncode, time.perf_counter() - t0, (p.stdout + p.stderr)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--exe", default=str(ROOT / "dist" / "pdf2pptx" / "pdf2pptx.exe"))
    ap.add_argument("--files", nargs="*", default=FILES)
    args = ap.parse_args()
    exe = Path(args.exe)
    if not exe.is_file():
        print(f"no exe at {exe}: build it first (packaging/build.py)")
        return 2
    from pdf2pptx.config import load_config
    from pdf2pptx.verify import score

    cfg = load_config({"tier": "lite"})
    out = ROOT / "out" / "package"
    out.mkdir(parents=True, exist_ok=True)
    rc, secs, text = run([str(exe), "--version"])
    print(f"--version: rc={rc} {text.strip()} ({secs:.1f} s)")
    ok = rc == 0
    rc, secs, text = run([str(exe), "doctor", "--no-models"])
    print(f"doctor: rc={rc} ({secs:.1f} s)")
    ok &= rc == 0
    rows = []
    for name in args.files:
        pdf = ROOT / "corpus" / "synthetic" / f"{name}.pdf"
        res = {}
        for who, cmd in (("exe", [str(exe)]),
                         ("dev", [sys.executable, "-m", "pdf2pptx.cli"])):
            pptx = out / f"{name}.{who}.pptx"
            rc, secs, text = run([*cmd, "convert", str(pdf), "-o", str(pptx), "--tier", "lite",
                                  "--no-cache", "--work-dir",
                                  str(ROOT / "build" / f"smoke_work_{who}")], cwd=ROOT)
            if rc != 0:
                print(f"{name} {who}: convert failed rc={rc}\n{text[-2000:]}")
                ok = False
                continue
            s = score.score_pair(pdf, pptx, cfg, name=f"smoke_{who}_{name}")["summary"]
            res[who] = {k: (s.get(k) or {}).get("mean") for k in KEYS} | {"seconds": round(
                secs, 1)}
        if len(res) == 2:
            diff = {k: abs((res["exe"][k] or 0) - (res["dev"][k] or 0)) for k in KEYS}
            same = all(d <= TOL * max(abs(res["dev"][k] or 0), 1e-9) for k, d in diff.items())
            ok &= same
            rows.append({"file": name, "exe": res["exe"], "dev": res["dev"], "same": same})
            print(f"{name}: composite exe {res['exe']['composite']:.4f} dev "
                  f"{res['dev']['composite']:.4f} | ssim {res['exe']['ssim']:.4f} / "
                  f"{res['dev']['ssim']:.4f} | {res['exe']['seconds']} s / "
                  f"{res['dev']['seconds']} s -> {'same' if same else 'DIFFERENT'}")
    (out / "smoke_test.json").write_text(json.dumps({"ok": ok, "rows": rows}, indent=1),
                                         encoding="utf-8", newline="\n")
    print("OK" if ok else "FAILED")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
