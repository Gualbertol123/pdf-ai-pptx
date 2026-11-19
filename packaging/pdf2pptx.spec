# PyInstaller spec for pdf2pptx (M10). Build: `python packaging/build.py` (one-folder),
# `python packaging/build.py --onefile` (single exe). Models and llama.cpp are NOT bundled:
# `pdf2pptx setup models` / `pdf2pptx setup llm` put them under %LOCALAPPDATA%\pdf2pptx.
# -*- mode: python -*-
import os
import sys
from pathlib import Path

from PyInstaller.utils.hooks import collect_all, collect_data_files, collect_submodules

ROOT = Path(SPECPATH).parent  # noqa: F821 - defined by PyInstaller
sys.path.insert(0, str(ROOT / "src"))
from pdf2pptx import __version__  # noqa: E402

ONEFILE = os.environ.get("PDF2PPTX_ONEFILE") == "1"

datas, binaries, hidden = [], [], []
datas += collect_data_files("pdf2pptx")          # manifests, ui/index.html
hidden += collect_submodules("pdf2pptx")         # stages are imported lazily
for pkg in ("pypdfium2", "pypdfium2_raw", "pptx", "pdfminer", "pdfplumber", "reportlab",
            "skimage", "onnxruntime", "winrt", "xlsxwriter", "fontTools"):
    d, b, h = collect_all(pkg)
    datas += d
    binaries += b
    hidden += h
hidden += ["win32com", "win32com.client", "pythoncom", "pywintypes", "psutil"]

v = tuple(int(x) for x in __version__.split(".")) + (0,)
version_file = ROOT / "build" / "version_info.txt"
version_file.parent.mkdir(parents=True, exist_ok=True)
version_file.write_text(f"""VSVersionInfo(
  ffi=FixedFileInfo(filevers={v}, prodvers={v}, mask=0x3f, flags=0x0, OS=0x40004,
                    fileType=0x1, subtype=0x0, date=(0, 0)),
  kids=[StringFileInfo([StringTable('040904B0', [
    StringStruct('CompanyName', 'pdf2pptx'),
    StringStruct('FileDescription', 'pdf2pptx: local PDF to editable PowerPoint converter'),
    StringStruct('FileVersion', '{__version__}'),
    StringStruct('InternalName', 'pdf2pptx'),
    StringStruct('OriginalFilename', 'pdf2pptx.exe'),
    StringStruct('ProductName', 'pdf2pptx'),
    StringStruct('ProductVersion', '{__version__}')])]),
    VarFileInfo([VarStruct('Translation', [1033, 1200])])])
""", encoding="utf-8")

a = Analysis(  # noqa: F821
    [str(ROOT / "packaging" / "entry.py")],
    pathex=[str(ROOT / "src")],
    binaries=binaries,
    datas=datas,
    hiddenimports=hidden,
    excludes=["tkinter", "pytest", "mypy", "ruff", "IPython", "matplotlib", "PyInstaller"],
    noarchive=False,
)
pyz = PYZ(a.pure)  # noqa: F821

if ONEFILE:
    exe = EXE(  # noqa: F821
        pyz, a.scripts, a.binaries, a.datas, [],
        name="pdf2pptx", console=True, upx=False, version=str(version_file),
        runtime_tmpdir=None,
    )
else:
    exe = EXE(  # noqa: F821
        pyz, a.scripts, [], exclude_binaries=True,
        name="pdf2pptx", console=True, upx=False, version=str(version_file),
    )
    coll = COLLECT(exe, a.binaries, a.datas, name="pdf2pptx", upx=False)  # noqa: F821
