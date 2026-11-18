"""LLM layer (M9): GGUF header reader, memory plan, cache + scheduler (fake server),
schema validation, judgment merge, bench parsers. The live test needs llama.cpp and a
vision model (``pdf2pptx setup llm``) and is skipped without them."""

from __future__ import annotations

import struct
from pathlib import Path
from typing import Any

import pytest

from pdf2pptx.ir import Document, Page, Paragraph, TextBlock, TextRun
from pdf2pptx.llm import bench, gguf, judge, schemas
from pdf2pptx.llm import scheduler as sch


def _s(text: str) -> bytes:
    b = text.encode()
    return struct.pack("<Q", len(b)) + b


def write_gguf(path: Path, layers: int = 4, expert_bytes: int = 1000,
               other_bytes: int = 500) -> Path:
    kv = [("general.architecture", 8, _s("toymoe")),
          ("toymoe.block_count", 4, struct.pack("<I", layers)),
          ("toymoe.attention.head_count", 4, struct.pack("<I", 8)),
          ("toymoe.attention.head_count_kv", 4, struct.pack("<I", 2)),
          ("toymoe.embedding_length", 4, struct.pack("<I", 64))]
    tensors = [(f"blk.{i}.ffn_up_exps.weight", expert_bytes) for i in range(layers)]
    tensors += [("token_embd.weight", other_bytes)]
    head = b"GGUF" + struct.pack("<IQQ", 3, len(tensors), len(kv))
    for k, t, v in kv:
        head += _s(k) + struct.pack("<I", t) + v
    off = 0
    for name, size in tensors:
        head += _s(name) + struct.pack("<I", 1) + struct.pack("<Q", size) + struct.pack(
            "<I", 0) + struct.pack("<Q", off)
        off += size
    pad = (-len(head)) % 32
    path.write_bytes(head + b"\0" * pad + b"\1" * off)
    return path


def test_gguf_reader_sizes(tmp_path: Path) -> None:
    info = gguf.read(write_gguf(tmp_path / "toy.gguf"))
    assert info.arch == "toymoe" and info.n_layers == 4
    assert info.expert_bytes_per_layer() == {0: 1000, 1: 1000, 2: 1000, 3: 1000}
    assert info.tensors["token_embd.weight"] == 500
    # 4 layers x 2 kv heads x (8 + 8) dims x 100 tokens x 1 byte
    assert info.kv_bytes(100, 1.0) == 4 * 2 * 16 * 100


def test_memory_plan(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    from pdf2pptx.llm import server

    p = write_gguf(tmp_path / "toy.gguf", layers=4, expert_bytes=1 << 30, other_bytes=1 << 20)
    monkeypatch.setattr(server, "candidate", lambda n: {"name": n, "model": "toy",
                                                       "mmproj": None})
    monkeypatch.setattr(server.models, "path", lambda n: p)
    monkeypatch.setattr(server, "COMPUTE_BUFFER", 0)
    monkeypatch.delenv("PDF2PPTX_LLM_RESERVE_GB", raising=False)
    gb = 1 << 30
    # 4 GB model: RAM need is the whole model whatever --n-cpu-moe is
    plan = server.plan_memory("toy", 16, available=13 * gb, gpu_free=10 * gb)
    assert plan.fits and plan.n_cpu_moe == 0 and plan.need >= 4 * gb
    # a 2.8 GB device budget moves two layers of experts to the CPU
    plan = server.plan_memory("toy", 16, available=13 * gb, gpu_free=int(2.8 * gb / 0.9))
    assert plan.fits and plan.n_cpu_moe == 2 and plan.suggestion == "--n-cpu-moe 2"
    # RAM short: refused, with how much is missing
    plan = server.plan_memory("toy", 16, available=10 * gb, gpu_free=10 * gb)
    assert not plan.fits and "short of RAM" in plan.suggestion
    monkeypatch.setenv("PDF2PPTX_LLM_RESERVE_GB", "4")
    assert server.plan_memory("toy", 16, available=10 * gb, gpu_free=10 * gb).fits


class FakeServer:
    started = 0

    def __init__(self, name: str, **kw: Any) -> None:
        self.name = name
        self.plan = None
        self.mmproj_cpu_fallback = False
        from pdf2pptx.llm.server import ServerStats

        self.stats = ServerStats()

    def start(self, **kw: Any) -> FakeServer:
        FakeServer.started += 1
        return self

    def stop(self) -> None:
        pass

    def chat(self, prompt: str, **kw: Any) -> dict[str, Any]:
        return {"content": '{"label": "process", "confidence": 0.9}',
                "json": {"label": "process", "confidence": 0.9}, "timings": {},
                "finish_reason": "stop"}


def test_scheduler_groups_by_model_and_caches(tmp_path: Path,
                                              monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(sch, "LlamaServer", FakeServer)
    monkeypatch.setattr(sch, "model_id", lambda n: f"{n}@abc")
    FakeServer.started = 0
    s = sch.Scheduler(tmp_path / "cache", resident=False)
    got: list[str] = []
    calls = [("A", "p1"), ("B", "p2"), ("A", "p3"), ("B", "p4")]
    for m, prompt in calls:
        s.submit(sch.Job(model=m, prompt=prompt, images=[b"x"],
                         schema=schemas.REGION, on_done=lambda r, m=m: got.append(m)))
    runs = s.run()
    assert [r.model for r in runs] == ["A", "B"] and FakeServer.started == 2
    assert sorted(got) == ["A", "A", "B", "B"]
    # same calls again: all from the cache, no server
    for m, prompt in calls:
        s.submit(sch.Job(model=m, prompt=prompt, images=[b"x"], schema=schemas.REGION))
    runs = s.run()
    assert FakeServer.started == 2 and sum(r.cached for r in runs) == 4


def test_schema_validation() -> None:
    assert schemas.valid_region({"label": "process", "confidence": 0.8})
    assert schemas.valid_region({"label": "nonsense", "confidence": 0.8}) is None
    assert schemas.valid_order({"order": [2, 0, 1], "confidence": 0.9}, 3) == ([2, 0, 1], 0.9)
    assert schemas.valid_order({"order": [0, 0, 1], "confidence": 0.9}, 3) is None
    t = schemas.valid_table({"rows": 2, "cols": 2, "header_rows": 1, "confidence": 0.9,
                             "cells": [["a", "b"], ["c"]]})
    assert t is not None and t["cells"][1] == ["c", ""]
    ok = {"chart_type": "pie", "categories": ["a", "b"], "confidence": 0.9,
          "series": [{"name": "s", "values": [1, 2]}]}
    assert schemas.valid_chart(ok)
    assert schemas.valid_chart(ok | {"chart_type": "column_clustered",
                                     "series": [{"name": "s", "values": [1]}]}) is None
    per_slice = schemas.valid_chart(ok | {"series": [{"name": "a", "values": [1]},
                                                     {"name": "b", "values": [2]}]})
    assert per_slice is not None and per_slice["series"][0]["values"] == [1, 2]
    t2 = schemas.valid_table({"rows": 1, "cols": 2, "header_rows": 1, "confidence": 0.9,
                              "cells": [["a", "b"], ["c", "d"]]})
    assert t2 is not None and t2["rows"] == 2
    assert schemas.valid_roles({"blocks": [{"id": "x", "role": "title", "confidence": 1},
                                           {"id": "zz", "role": "title", "confidence": 1}]},
                               {"x"}) == {"x": ("title", 1)}


def _block(bid: str, text: str, y: float, role: str = "body", conf: float = 0.5) -> TextBlock:
    r = TextRun(id=bid + "r", bbox=(10, y, 200, y + 12), text=text)
    return TextBlock(id=bid, bbox=(10, y, 200, y + 12), role=role, role_confidence=conf,
                     paragraphs=[Paragraph(id=bid + "p", bbox=(10, y, 200, y + 12), runs=[r])])


def test_judge_apply_roles_threshold_and_invalid() -> None:
    page = Page(index=0, width=300, height=300, elements=[
        _block("b1", "Results", 10), _block("b2", "Figure 1. x", 50), _block("b3", "text", 90)])
    doc = Document(pages=[page])
    rep = judge.JudgeReport(proposals=[
        judge.Proposal("roles", 0, "b1,b2,b3", {"model": "M@1", "json": {"blocks": [
            {"id": "b1", "role": "title", "confidence": 0.9},
            {"id": "b2", "role": "caption", "confidence": 0.6},  # below MIN_ADOPT
            {"id": "b3", "role": "other", "confidence": 0.99}]}}),
        judge.Proposal("roles", 0, "b1", {"model": "M@1", "json": {"bad": 1}}),
    ])
    judge.apply(doc, rep)
    b1, b2, b3 = page.elements
    assert b1.role == "title" and b1.source == "model:M" and b1.role_confidence == 0.9
    assert b2.role == "body" and b3.role == "body"
    assert rep.invalid == 1 and judge.affected_pages(rep) == [0]


def test_keep_best_without_powerpoint_drops_changes(tmp_path: Path,
                                                    monkeypatch: pytest.MonkeyPatch) -> None:
    from pdf2pptx.config import Config
    from pdf2pptx.llm import max_tier

    lite = Document(pages=[Page(index=0, width=300, height=300,
                                elements=[_block("b1", "x", 10)])])
    doc = Document(pages=[Page(index=0, width=300, height=300,
                               elements=[_block("b1", "x", 10, role="title")])])

    def boom(*a: Any, **k: Any) -> list[float]:
        raise RuntimeError("PowerPoint unavailable")

    monkeypatch.setattr(max_tier, "_score", boom)
    out = max_tier.keep_best(doc, lite, [0], tmp_path / "x.pdf", tmp_path, Config())
    assert out["reverted"] == [0] and doc.pages[0].elements[0].role == "body"


def test_bench_parsers() -> None:
    otsl = "<otsl><fcel>Area<fcel>2025<nl><fcel>North<fcel>41<nl></otsl>"
    assert bench.parse_grid(otsl) == [["Area", "2025"], ["North", "41"]]
    md = "| | 2024 | 2025 |\n|---|---|---|\n| North | 41 | 56 |\n| South | 58 | 58 |"
    grid = bench.parse_grid(md)
    assert grid[1] == ["North", "41", "56"]
    cats, ser = bench.grid_to_chart(grid)
    assert cats == ["North", "South"] and ser == [[41, 58], [56, 58]]
    html = "<table><tr><td>a</td><td>b</td></tr><tr><td>1</td><td>2</td></tr></table>"
    assert bench.parse_grid(html) == [["a", "b"], ["1", "2"]]


def test_write_stack_keeps_other_settings(tmp_path: Path) -> None:
    p = tmp_path / "pdf2pptx.toml"
    p.write_text('[pdf2pptx]\nworkers = 3\nllm_judge = "old"\n', encoding="utf-8")
    bench.write_stack({"llm_judge": "Qwen3.5-9B", "llm_chart": None}, p)
    from pdf2pptx.config import load_config

    cfg = load_config(config_path=p)
    assert cfg.workers == 3 and cfg.llm_judge == "Qwen3.5-9B" and cfg.llm_chart is None


def _live_model() -> str | None:
    from pdf2pptx.llm.setup import candidate, installed, llama_server

    if llama_server() is None:
        return None
    for n in ("Qwen3.5-4B", "Qwen3.5-9B", "Qwen3.6-35B-A3B"):
        try:
            if installed(candidate(n)):
                return n
        except KeyError:
            continue
    return None


@pytest.mark.corpus
def test_live_vlm_reads_a_chart(tmp_path: Path) -> None:
    name = _live_model()
    if name is None:
        pytest.skip("no llama.cpp / vision model (pdf2pptx setup llm)")
    import pypdfium2 as pdfium

    from pdf2pptx.config import load_config
    from pdf2pptx.llm.server import LlamaServer, LLMUnavailable

    cfg = load_config()
    pdf = pdfium.PdfDocument(str(cfg.corpus_dir / "synthetic" / "d16_chart_pie.pdf"))
    img = pdf[0].render(scale=0.75).to_pil()
    png = tmp_path / "p.png"
    img.save(png)
    try:
        with LlamaServer(name) as srv:
            r = srv.chat(schemas.chart_prompt(), images=[png.read_bytes()],
                         schema=schemas.CHART, system=schemas.SYSTEM, max_tokens=600)
    except LLMUnavailable as exc:
        pytest.skip(str(exc))
    v = schemas.valid_chart(r["json"])
    assert v is not None and v["chart_type"] in ("pie", "doughnut")
    assert [c.lower() for c in v["categories"]] == ["individuals", "foundations",
                                                    "government", "events"]
