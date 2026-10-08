#!/usr/bin/env python3
"""Generates the code-page tables behind CHAR, CODE and the byte-counting text functions.

Outputs:
  src/eval/dbcs_jis0208_table.cpp  JIS X 0208 (ja-JP), Python iso2022_jp codec
  src/eval/dbcs_gb2312_table.cpp   GB 2312 (zh-CN), Python gb2312 codec
  src/eval/dbcs_ksx1001_table.cpp  KS X 1001 (ko-KR), Python euc_kr codec
  src/sbcs_codepage.cpp            the Mac Thai block, from the th-TH golden
                                   (locale_tokens.char_high_bytes_unicode)

Each double-byte table stores only row/cell -> Unicode for the 94x94 grid,
delta-encoded as an adaptive Exp-Golomb bit stream (see `encode_cells`;
the decoder is `DbcsCells` in src/eval/dbcs_table.cpp). The codec choice
reproduces every measured CODE value in code_char_jp_probes on the Mac
ja-JP, zh-CN and ko-KR targets.

Usage:
    python3 tools/dbcs/generate_tables.py          # rewrite the outputs
    python3 tools/dbcs/generate_tables.py --check  # exit 1 on drift
"""

from __future__ import annotations

import argparse
import json
import sys
import textwrap
from dataclasses import dataclass
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
SBCS_PATH = REPO_ROOT / "src" / "sbcs_codepage.cpp"
THAI_GOLDEN = REPO_ROOT / "tests" / "oracle" / "targets" / "mac-365-th_TH" / "golden" / "locale_tokens.golden.json"
REGENERATION_GUIDANCE = "Regenerate with: python3 tools/dbcs/generate_tables.py"

GRID = 94
SLOTS = GRID * GRID
# Adaptive parameter: k = bit_length(acc >> ACC_SHIFT) - K_BIAS; must match dbcs_table.cpp.
ACC_SHIFT = 1
K_BIAS = 2
ACC_INIT = 32

THAI_BEGIN = "// BEGIN GENERATED (tools/dbcs/generate_tables.py)\n"
THAI_END = "// END GENERATED\n"


@dataclass(frozen=True)
class DbcsTable:
    name: str
    title: str
    codec: str
    accessor: str
    budget_kb: int

    @property
    def path(self) -> Path:
        return REPO_ROOT / "src" / "eval" / f"dbcs_{self.name}_table.cpp"


TABLES = (
    DbcsTable("jis0208", "JIS X 0208", "iso2022_jp", "jis0208_cells", 10),
    DbcsTable("gb2312", "GB 2312", "gb2312", "gb2312_cells", 11),
    DbcsTable("ksx1001", "KS X 1001", "euc_kr", "ksx1001_cells", 11),
)


def decode_cell(codec: str, row: int, cell: int) -> int:
    """Return the Unicode code point at 1-based (row, cell), or 0 when unassigned."""
    if codec == "iso2022_jp":
        payload = b"\x1b\x24\x42" + bytes([row + 0x20, cell + 0x20])
    else:
        payload = bytes([row + 0xA0, cell + 0xA0])
    try:
        text = payload.decode(codec)
    except (UnicodeDecodeError, ValueError):
        return 0
    return ord(text) if len(text) == 1 and ord(text) <= 0xFFFF else 0


def cells_for(codec: str) -> list[int]:
    return [decode_cell(codec, r, c) for r in range(1, GRID + 1) for c in range(1, GRID + 1)]


class BitWriter:
    def __init__(self) -> None:
        self.out = bytearray()
        self.acc = 0
        self.count = 0

    def bits(self, value: int, width: int) -> None:
        for shift in range(width - 1, -1, -1):
            self.acc = (self.acc << 1) | ((value >> shift) & 1)
            self.count += 1
            if self.count == 8:
                self.out.append(self.acc)
                self.acc = 0
                self.count = 0

    def exp_golomb(self, value: int, k: int) -> None:
        w = (value >> k) + 1
        length = w.bit_length() - 1
        self.bits(0, length)
        self.bits(w, length + 1)
        self.bits(value, k)

    def finish(self) -> bytes:
        if self.count:
            self.out.append(self.acc << (8 - self.count))
        return bytes(self.out)


def encode_cells(cells: list[int]) -> bytes:
    """Encode slots in row-major order.

    A mapped slot is the zigzag delta from the previous mapped code point
    (never 0, as no table repeats a code point); 0 is followed by EG(0) of
    the length-1 of a run of unassigned slots. The order k tracks a running
    average of recent deltas.
    """
    writer = BitWriter()
    prev = 0
    acc = ACC_INIT
    i = 0
    while i < len(cells):
        k = max(0, (acc >> ACC_SHIFT).bit_length() - K_BIAS)
        if cells[i] == 0:
            j = i
            while j < len(cells) and cells[j] == 0:
                j += 1
            writer.exp_golomb(0, k)
            writer.exp_golomb(j - i - 1, 0)
            i = j
            continue
        delta = cells[i] - prev
        if delta == 0:
            raise ValueError("duplicate code point breaks the zero escape")
        value = 2 * delta if delta > 0 else -2 * delta - 1
        writer.exp_golomb(value, k)
        acc += value - (acc >> ACC_SHIFT)
        prev = cells[i]
        i += 1
    return writer.finish()


def render_dbcs(table: DbcsTable) -> str:
    cells = cells_for(table.codec)
    encoded = encode_cells(cells)
    mapped = sum(1 for cp in cells if cp)
    body = []
    for start in range(0, len(encoded), 19):
        chunk = encoded[start : start + 19]
        body.append("    " + ", ".join(f"0x{b:02X}" for b in chunk) + ",\n")
    head = textwrap.dedent(
        f"""\
        // GENERATED by tools/dbcs/generate_tables.py from Python's {table.codec} codec; do not edit.
        // {table.title}: {mapped} of {SLOTS} cells mapped, {len(encoded)} encoded bytes.
        // @size-budget: {table.budget_kb} KB

        #include <cstdint>

        #include "eval/dbcs_table.h"

        namespace formulon {{
        namespace eval {{
        namespace dbcs_detail {{
        namespace {{

        constexpr std::uint8_t kEncoded[] = {{
        """
    )
    tail = textwrap.dedent(
        f"""\
        }};

        }}  // namespace

        const std::uint16_t* {table.accessor}() noexcept {{
          static const DbcsCells cells(kEncoded, sizeof(kEncoded));
          return cells.unicode;
        }}

        }}  // namespace dbcs_detail
        }}  // namespace eval
        }}  // namespace formulon
        """
    )
    return head + "".join(body) + tail


def thai_high_bytes(golden: Path = THAI_GOLDEN) -> list[int]:
    data = json.loads(golden.read_text(encoding="utf-8"))
    for case in data["cases"]:
        if case["id"] == "char_high_bytes_unicode":
            values = [int(v) for v in case["expect"]["value"].split(",")]
            if len(values) != 128 or any(v < 0 for v in values):
                raise ValueError("char_high_bytes_unicode must list 128 decodable bytes")
            return values
    raise ValueError(f"char_high_bytes_unicode missing from {golden}")


def render_thai_block(values: list[int]) -> str:
    lines = [THAI_BEGIN, "// Mac Thai, measured: locale_tokens.char_high_bytes_unicode (mac-365-th_TH).\n"]
    lines.append("constexpr std::array<std::uint32_t, 128> kMacThaiHigh = {{\n")
    for start in range(0, len(values), 13):
        chunk = values[start : start + 13]
        lines.append("    " + ", ".join(f"0x{v:04X}u" for v in chunk) + ",\n")
    lines.append("}};\n")
    lines.append(THAI_END)
    return "".join(lines)


def splice_thai(source: str, block: str) -> str:
    begin = source.find(THAI_BEGIN)
    end = source.find(THAI_END)
    if begin < 0 or end < begin:
        raise ValueError("generated Mac Thai markers missing from sbcs_codepage.cpp")
    return source[:begin] + block + source[end + len(THAI_END) :]


def outputs(sbcs_source: str) -> dict[Path, str]:
    rendered = {table.path: render_dbcs(table) for table in TABLES}
    rendered[SBCS_PATH] = splice_thai(sbcs_source, render_thai_block(thai_high_bytes()))
    return rendered


def check(paths: dict[Path, Path] | None = None) -> int:
    """Compare tracked files with the generator output.

    `paths` maps a canonical output path to the file actually compared, so a
    test can point the check at a temporary fixture. Returns 1 on drift and
    2 when a file cannot be read.
    """
    paths = paths or {}
    try:
        sbcs_actual = paths.get(SBCS_PATH, SBCS_PATH).read_text(encoding="utf-8")
        expected = outputs(sbcs_actual)
        actual = {canonical: paths.get(canonical, canonical).read_text(encoding="utf-8") for canonical in expected}
    except (OSError, ValueError) as exc:
        print(f"cannot check generated tables: {exc}", file=sys.stderr)
        return 2
    drifted = [path for path, text in expected.items() if actual[path] != text]
    for path in drifted:
        print(f"generated table drift detected: {path}", file=sys.stderr)
    if drifted:
        print(REGENERATION_GUIDANCE, file=sys.stderr)
        return 1
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n", 1)[0])
    parser.add_argument("--check", action="store_true", help="compare the outputs with the tracked files")
    args = parser.parse_args(argv)
    if args.check:
        return check()
    for path, text in outputs(SBCS_PATH.read_text(encoding="utf-8")).items():
        path.write_text(text, encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
