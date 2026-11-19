"""Build the Windows package with PyInstaller (M10).

    .venv\\Scripts\\python packaging\\build.py            # one-folder: dist\\pdf2pptx\\pdf2pptx.exe
    .venv\\Scripts\\python packaging\\build.py --onefile  # single file: dist\\pdf2pptx.exe

PyInstaller is a build tool only (``pip install pyinstaller``); it is not a runtime
dependency. Models and llama.cpp stay out of the bundle (``pdf2pptx setup ...``).
"""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--onefile", action="store_true")
    args = ap.parse_args()
    env = dict(os.environ, PDF2PPTX_ONEFILE="1" if args.onefile else "0")
    work = ROOT / "build" / ("onefile" if args.onefile else "onedir")
    dist = ROOT / "dist" / ("onefile" if args.onefile else "")
    t0 = time.perf_counter()
    cmd = [sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean",
           "--workpath", str(work), "--distpath", str(dist),
           str(ROOT / "packaging" / "pdf2pptx.spec")]
    rc = subprocess.call(cmd, cwd=ROOT, env=env)
    exe = dist / ("pdf2pptx.exe" if args.onefile else "pdf2pptx/pdf2pptx.exe")
    if rc == 0 and exe.is_file():
        size = (exe.stat().st_size if args.onefile else
                sum(p.stat().st_size for p in exe.parent.rglob("*") if p.is_file()))
        print(f"built {exe} ({size / 2**20:.0f} MB) in {time.perf_counter() - t0:.0f} s")
    return rc


if __name__ == "__main__":
    sys.exit(main())
