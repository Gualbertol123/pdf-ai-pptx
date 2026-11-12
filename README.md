# pdf2pptx

A local, offline PDF → PPTX converter that rebuilds native, editable PowerPoint
objects and checks its own output by rendering it in PowerPoint. No document ever
leaves this PC. Design: [PLAN.md](PLAN.md). Build prompts: [PROMPTS.md](PROMPTS.md).
Status: [PROGRESS.md](PROGRESS.md).

## Setup (per user, no admin rights needed)

Requirements: Windows 11, Python 3.12 (`py -3.12`), Microsoft PowerPoint (desktop).

```powershell
cd pdf_pptx
py -3.12 -m venv .venv
.venv\Scripts\python -m pip install --upgrade pip
.venv\Scripts\python -m pip install -e ".[dev]"
.venv\Scripts\pdf2pptx doctor
```

`doctor` reports Python, PowerPoint COM, onnxruntime providers, the Vulkan DLL,
RAM and GPU. PowerPoint COM must say `available`. Without it the converter still
runs, but every score that needs PowerPoint comes out as `null`.

## Commands

| Command | What it does | Status |
|---|---|---|
| `pdf2pptx doctor [--json]` | Machine check | M1 |
| `pdf2pptx corpus build [--only NAME ...]` | Generates the synthetic corpus in `corpus/` (decks → PDF via PowerPoint, reportlab reports, fake scans) with ground truth | M1 |
| `pdf2pptx corpus list` | Lists the corpus | M1 |
| `pdf2pptx score <pdf> <pptx>` | Scores one PPTX against its PDF → `work/scores/single/` | M1 |
| `pdf2pptx score --corpus [--against out\|truth\|blank] [--label L]` | Scores the whole corpus → `work/scores/<label>/` (per-file JSON, `summary.json`, `summary.md`) | M1 |
| `pdf2pptx convert in.pdf -o out.pptx [--mode ...] [--text-layout blocks\|lines]` | Converts a PDF. Text comes out as one wrapping box per text block with paragraphs, real bullets and matched fonts (M4); `--text-layout lines` writes the M2 baseline (one box per line) for comparison. `--tune/--no-tune`, `--tune-rounds N`: the fidelity tuner renders in PowerPoint and adjusts width, spacing, size, line spacing and position until line breaks and baselines match (on by default in fidelity and balanced) | M2, M4, M5 |
| `pdf2pptx structure [--only NAME ...] [--no-cache]` | Compares the analyzed paragraphs, levels, bullets and font families with the synthetic decks' truth → `work/scores/structure.json` | M4 |
| `pdf2pptx report <pdf> [<pdf> ...] [-o file/dir] [--pptx P]` | Converts, scores and writes one self-contained HTML per PDF to `out/report/` (PDF \| PPTX \| diff heatmap, IR overlays, sortable scores); several PDFs (or `--corpus [--only ...]`) also get an `index.html` | M3 |
| `pdf2pptx setup models\|llm` | Downloads models and tools | stub, arrives in M8/M9 |

`--against truth` scores the answer-key decks against their own PDFs, which checks
that the metrics behave. `--against blank` scores empty slides, which must score
low. `--against out` (the default) scores the converter output in `out/corpus/`.

Configuration: defaults, then `pdf2pptx.toml` in the project root (keys of
`Config` in `src/pdf2pptx/config.py`, optionally under a `[pdf2pptx]` table),
then CLI flags. Models live in `%LOCALAPPDATA%\pdf2pptx\models`.

## Development

```powershell
.venv\Scripts\pytest -q          # tests marked powerpoint/corpus skip when unavailable
.venv\Scripts\ruff check .
.venv\Scripts\mypy               # ir.py and build/units.py, strict
```

Put your own PDFs in `corpus/real/`. That folder is git-ignored, and later
milestones score it too.

## What needs IT (corporate PC)

- **GPU driver update.** The Radeon 780M driver is 31.0.22056.6001 (March 2024).
  A newer AMD Adrenalin driver improves Vulkan and DirectML. It matters from
  Milestone 8 on and isn't needed for Milestones 1–7.
- **Power mode.** Set Windows to *Best performance* and stay plugged in for long runs.
  Some corporate policies lock this setting.
- **RAM upgrade to 64 GB** (2 SODIMM slots) is the best single hardware step for the
  Max tier (Milestone 9).
- Nothing else needs admin rights: Python packages, models and llama.cpp all install
  per user.

## Troubleshooting

- *PowerPoint COM not available*: open PowerPoint once interactively (first-run
  dialogs or sign-in can block automation), then re-run `pdf2pptx doctor`.
- *A POWERPNT.EXE stays open*: our code only closes PowerPoint instances it started
  itself. If you had PowerPoint open, it is left alone and only our presentations
  are closed.
- PowerPoint runs behind a worker process (`verify/ppt_com.py`). Every POWERPNT.EXE it
  starts is recorded in `%TEMP%\pdf2pptx\ppt_owned\`; if pdf2pptx is killed, the next
  run reaps those instances (and only those).
