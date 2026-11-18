"""Minimal GGUF header reader: metadata and tensor sizes (for memory planning).

Tensor sizes come from the data offsets (next offset - this offset), so no table of
quantisation block sizes is needed. Only the header is read, never the weights.
"""

from __future__ import annotations

import re
import struct
from dataclasses import dataclass, field
from functools import cache
from pathlib import Path
from typing import Any, BinaryIO

_SCALARS = {0: "B", 1: "b", 2: "H", 3: "h", 4: "I", 5: "i", 6: "f", 7: "?", 10: "Q", 11: "q",
            12: "d"}
EXPERT_RE = re.compile(r"^blk\.(\d+)\.ffn_(?:up|down|gate|gate_up)_exps\.")


@dataclass
class GGUFInfo:
    path: Path
    metadata: dict[str, Any] = field(default_factory=dict)
    tensors: dict[str, int] = field(default_factory=dict)  # name -> bytes
    file_size: int = 0

    @property
    def arch(self) -> str:
        return str(self.metadata.get("general.architecture", ""))

    def meta(self, key: str, default: Any = None) -> Any:
        return self.metadata.get(f"{self.arch}.{key}", default)

    @property
    def n_layers(self) -> int:
        return int(self.meta("block_count", 0) or 0)

    def expert_bytes_per_layer(self) -> dict[int, int]:
        out: dict[int, int] = {}
        for name, size in self.tensors.items():
            m = EXPERT_RE.match(name)
            if m:
                out[int(m.group(1))] = out.get(int(m.group(1)), 0) + size
        return out

    def kv_bytes(self, ctx: int, bytes_per_elem: float = 1.0625) -> int:
        """K + V cache for ``ctx`` tokens (q8_0 by default). Hybrid models (linear
        attention layers) need less: this is an upper bound."""
        n_kv = self.meta("attention.head_count_kv", self.meta("attention.head_count", 0))
        if isinstance(n_kv, list):
            n_kv = max(n_kv) if n_kv else 0
        k_len = self.meta("attention.key_length")
        v_len = self.meta("attention.value_length")
        if not k_len:
            emb = int(self.meta("embedding_length", 0) or 0)
            heads = int(self.meta("attention.head_count", 1) or 1)
            k_len = v_len = emb // max(heads, 1)
        interval = int(self.meta("full_attention_interval", 1) or 1)
        layers = self.n_layers // max(interval, 1) if interval > 1 else self.n_layers
        per_tok = layers * int(n_kv or 0) * (int(k_len or 0) + int(v_len or k_len or 0))
        return int(per_tok * ctx * bytes_per_elem)


def _read(fh: BinaryIO, fmt: str) -> Any:
    size = struct.calcsize("<" + fmt)
    return struct.unpack("<" + fmt, fh.read(size))[0]


def _str(fh: BinaryIO) -> str:
    n = _read(fh, "Q")
    return fh.read(n).decode("utf-8", "replace")


def _value(fh: BinaryIO, vtype: int, keep_arrays: bool) -> Any:
    if vtype in _SCALARS:
        return _read(fh, _SCALARS[vtype])
    if vtype == 8:
        return _str(fh)
    if vtype == 9:
        etype, n = _read(fh, "I"), _read(fh, "Q")
        if etype in _SCALARS and not keep_arrays:
            fh.seek(n * struct.calcsize("<" + _SCALARS[etype]), 1)
            return None
        if etype in _SCALARS:
            return [_read(fh, _SCALARS[etype]) for _ in range(n)]
        vals = [_value(fh, etype, False) for _ in range(n)]
        return vals if keep_arrays or n <= 64 else None
    raise ValueError(f"unknown GGUF value type {vtype}")


@cache
def read(path: Path) -> GGUFInfo:
    path = Path(path)
    info = GGUFInfo(path=path, file_size=path.stat().st_size)
    with path.open("rb") as fh:
        if fh.read(4) != b"GGUF":
            raise ValueError(f"{path}: not a GGUF file")
        version = _read(fh, "I")
        n_tensors, n_kv = _read(fh, "Q"), _read(fh, "Q")
        for _ in range(n_kv):
            key = _str(fh)
            vtype = _read(fh, "I")
            keep = key.endswith(("head_count_kv", "head_count")) and vtype == 9
            info.metadata[key] = _value(fh, vtype, keep)
        offsets: list[tuple[int, str]] = []
        for _ in range(n_tensors):
            name = _str(fh)
            nd = _read(fh, "I")
            fh.seek(8 * nd + 4, 1)  # dims, type
            offsets.append((_read(fh, "Q"), name))
        align = int(info.metadata.get("general.alignment", 32) or 32)
        pos = fh.tell()
        data_start = (pos + align - 1) // align * align
    offsets.sort()
    data_len = info.file_size - data_start
    for k, (off, name) in enumerate(offsets):
        nxt = offsets[k + 1][0] if k + 1 < len(offsets) else data_len
        info.tensors[name] = nxt - off
    info.metadata["gguf.version"] = version
    return info
