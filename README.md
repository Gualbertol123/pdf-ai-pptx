# pdf2pptx

A local, offline PDF → PowerPoint converter. It rebuilds **native, editable objects**:
real text with paragraphs and bullets, built-in shapes, tables, charts, slide masters,
theme, animations, notes and sections. It checks its own output by rendering it in
PowerPoint and correcting mismatches. It can also **rewrite a long document as a
presentation** with a local language model.

Nothing leaves this PC: no cloud service, no upload, and no network at conversion time.

Design: [PLAN.md](PLAN.md). Architecture: [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md).
Results and measurements: [PROGRESS.md](PROGRESS.md).

## Install

Requirements: Windows 10/11 and Microsoft PowerPoint (desktop). The converter works
without PowerPoint, but then it cannot check or tune its output.

### Option A: the packaged build (no Python needed)

1. Copy the `pdf2pptx` folder (built with `packaging/build.py`, see below) anywhere you
   can write, for example `%LOCALAPPDATA%\Programs\pdf2pptx`.
2. Run `pdf2pptx.exe doctor` from that folder.
3. Use `pdf2pptx.exe` in place of `pdf2pptx` in the commands below.

The exe keeps its cache, tools and settings in `%LOCALAPPDATA%\pdf2pptx` (set
`PDF2PPTX_HOME` to use another folder).

### Option B: from source (Python 3.12)

```powershell
cd pdf_pptx
py -3.12 -m venv .venv
.venv\Scripts\python -m pip install -e ".[dev,ml]"
.venv\Scripts\pdf2pptx doctor
```

### Optional downloads (one-time, the only commands that use the network)

```powershell
pdf2pptx setup models --lite   # ~346 MB ONNX models: OCR, layout, background cleanup (Lite tier)
pdf2pptx setup llm             # llama.cpp (Vulkan) + local LLMs for the Max tier and deck mode
pdf2pptx bench                 # measures the installed LLMs and writes the best stack to pdf2pptx.toml
pdf2pptx setup llm --remove-except Qwen3.5-9B Qwen3.5-4B PaddleOCR-VL-1.6   # keep what the bench picked
```

Every file is pinned by sha256. Models go to `%LOCALAPPDATA%\pdf2pptx\models` and
llama.cpp to `tools\llama.cpp` (dev build) or `%LOCALAPPDATA%\pdf2pptx\tools\llama.cpp`
(packaged build). The two builds share the models folder. The packaged build needs
`pdf2pptx setup llm` once to fetch its own llama.cpp.

## Use

```powershell
pdf2pptx ui                                        # local web app: http://127.0.0.1:8765
pdf2pptx convert report.pdf -o report.pptx         # page -> slide, balanced mode
pdf2pptx convert report.pdf --mode deck --slides 12 --audience "board members" --lang it
pdf2pptx report report.pdf                         # HTML: PDF | PowerPoint | differences
```

### The local UI

`pdf2pptx ui` opens a page in your browser. If port 8765 is taken, the UI picks a free
port and prints it. On the page you can:

- drop a PDF;
- choose the mode and tier;
- follow progress stage by stage;
- see the PDF page, the PowerPoint slide and the difference heatmap side by side (for
  deck mode: the slides with each bullet's sources);
- download the PPTX;
- convert again any document in the cache.

The server listens on 127.0.0.1 only, refuses requests from other sites and loads
nothing from the internet. Jobs and uploads are kept in `work\ui\`.

### Modes

| Mode | What you get |
|---|---|
| `fidelity` | As close to the PDF as possible: forced line breaks, everything in fixed positions. Master and layouts are kept only if they render identically. |
| `balanced` (default) | Close to the PDF, with placeholders, theme colours and fonts, slide master, build animations. |
| `editable` | Clean placeholders and reflowable text (shrink on overflow). Small visual differences are allowed. No tuner. |
| `deck` | Rewrites the document as a presentation. Every bullet cites its source blocks, every number is checked against the source, the speaker notes hold the detail, and tables, charts and figures are reused. Also writes `<name>.deck_sources.json` (slide → pages and blocks, anything left out, and the audit). |

### Tiers

| Tier | Adds | Speed on this PC (median) |
|---|---|---|
| `lite` | Rules, OCR and layout models, background cleanup, the PowerPoint check-and-correct loop | ~1 s/page (editable), ~2 s/page (balanced, with the tuner) |
| `max` | A local vision model for judgment calls and chart reading. In deck mode, the text model writes the deck. | ~22 s/page for judgment. A deck from a 1-page report takes ~1 min; from a 39-page report, minutes (see PROGRESS.md). |
| `auto` (default) | `max` when llama.cpp, a vision model and 24 GB+ of RAM are present, else `lite` | |

Deck mode without the Max tier (or when no model fits in the free RAM) writes an
*extractive* deck: the first sentences of each section, still with sources.

### Other commands

| Command | What it does |
|---|---|
| `pdf2pptx doctor [--no-models]` | What this PC can do: PowerPoint COM, GPU / DirectML / Vulkan, models and whether each LLM fits in RAM right now |
| `pdf2pptx score <pdf> <pptx>` | Scores a PPTX against its PDF (SSIM, text, line breaks, editability) |
| `pdf2pptx convert --corpus` / `score --corpus --label L` | Converts and scores the test corpus |
| `pdf2pptx corpus build` / `corpus stress` | Generates the synthetic corpus, or the hardening inputs (long, image-heavy, encrypted, corrupted PDFs) |
| `pdf2pptx structure`, `pdf2pptx deck`, `pdf2pptx bench` | Developer checks: text structure, deck structure, LLM benchmark |

Settings: defaults, then `pdf2pptx.toml` (project root, or `%LOCALAPPDATA%\pdf2pptx` for
the packaged build), then command-line flags. The keys are the fields of `Config` in
`src/pdf2pptx/config.py`.

## What needs IT (corporate PC)

Nothing needs admin rights: Python packages, the packaged build, the models and
llama.cpp all install per user. The following help, and may need IT:

- **RAM upgrade to 64 GB** (2 SODIMM slots). This is the best single step. With about
  19 GB free today, only the 9B / 4B models fit. 64 GB would let the stronger
  Qwen3.6-35B-A3B be the judge and deck writer, and keep a model resident between
  documents.
- **GPU driver update** (Radeon 780M; the installed driver is from March 2024) for
  better Vulkan and DirectML speed.
- **Power mode**: *Best performance*, plugged in, for long Max runs. Some policies lock
  this setting.
- **Security software** may scan `pdf2pptx.exe` and `llama-server.exe` on first start
  (slow first run), or block an unsigned exe. Ask IT to allow the folder if needed.
- **PowerPoint** must be the desktop app, signed in and opened once interactively.

## Troubleshooting

| Symptom | What to do |
|---|---|
| `the PDF is encrypted (password protected)` | Open it in a PDF viewer, remove the password (or print to PDF) and convert the copy. |
| `the PDF is damaged and cannot be read` / `not a PDF file` | Re-save or print it to PDF from a viewer. If only some pages are damaged, the file still converts, and the damaged pages come out as pictures. |
| `PowerPoint COM not available` in `doctor` | Open PowerPoint once by hand (first-run dialogs and sign-in block automation). Without PowerPoint, conversion still works but the tuner, the checks and the deck fit loop are skipped (the log says so). |
| A POWERPNT.EXE stays open | pdf2pptx only closes PowerPoint instances it started. A PowerPoint you opened is left alone, and only pdf2pptx's own presentations are closed. Instances left behind by a crash are reaped on the next run (`%TEMP%\pdf2pptx\ppt_owned`). |
| `does NOT fit now: x GB short of RAM` | Close programs, or let it fall back to the smaller model (automatic). The 8 GB reserve can be changed with `PDF2PPTX_LLM_RESERVE_GB` (lower values risk swapping). |
| The UI does not open | The URL is printed in the console. Port 8765 may be taken by another program: the UI then picks a free port. You can also choose one with `pdf2pptx ui --port 9000`. |
| Deck mode wrote an extractive deck | The Max tier is not available (`doctor`): install llama.cpp and a model, or free some RAM. |
| Very slow first conversion | First runs build the font index and warm up the models. Later runs reuse `work\<sha256>\` (delete it, or use `--no-cache`, to start over). |
| Wrong colours / missing graphics on a page | `pdf2pptx report file.pdf` shows the PDF page, the slide and a difference heatmap. Pages the converter could not rebuild come out as pictures (`fallback_pages` in the JSON output). |

## Development

```powershell
.venv\Scripts\pytest -q                  # ~2 min with PowerPoint; PowerPoint / corpus / model tests skip when missing
.venv\Scripts\ruff check .
.venv\Scripts\mypy                       # ir.py and build/units.py, strict
.venv\Scripts\python -m pip install pyinstaller      # build tool only (GPL with the bootloader exception)
.venv\Scripts\python packaging\build.py              # dist\pdf2pptx\pdf2pptx.exe (one folder)
.venv\Scripts\python packaging\build.py --onefile    # dist\onefile\pdf2pptx.exe
.venv\Scripts\python packaging\smoke_test.py         # exe vs dev build on 3 corpus files
.venv\Scripts\pdf2pptx corpus stress; .venv\Scripts\python scripts\hardening.py
```

Put your own PDFs in `corpus/real/`. That folder is git-ignored and is scored together
with the synthetic corpus. Only permissive licenses are used (MIT, BSD, Apache-2.0).
There are no AGPL dependencies.
