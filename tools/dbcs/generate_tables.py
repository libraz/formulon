#!/usr/bin/env python3
"""Generates the code-page tables behind CHAR, CODE and the byte-counting text functions.

Outputs:
  src/eval/dbcs_jis0208_table.cpp  JIS X 0208 (ja-JP), Python iso2022_jp codec
  src/eval/dbcs_gb2312_table.cpp   GB 2312 (zh-CN), Python gb2312 codec
  src/eval/dbcs_ksx1001_table.cpp  KS X 1001 (ko-KR), Python euc_kr codec
  src/eval/dbcs_big5_table.cpp     Big5 (zh-TW), Python big5 codec without the ETEN rows
  src/sbcs_codepage.cpp            the Mac Thai and Mac Cyrillic blocks, from the
                                   th-TH and ru-RU goldens
                                   (locale_tokens.char_high_bytes_unicode)

Each double-byte table stores only lead/trail -> Unicode over its code
space (`DbcsShape`: a lead-byte range and up to two trail-byte ranges; the
94x94 row/cell grid for JIS, GB and KS, 89 leads x 157 trails for Big5),
delta-encoded as an adaptive Exp-Golomb bit stream (see `encode_cells`;
the decoder is `DbcsCells` in src/eval/dbcs_table.cpp). The codec choice
reproduces every measured CODE value in code_char_jp_probes on the Mac
ja-JP, zh-CN, ko-KR and zh-TW targets.

Usage:
    python3 tools/dbcs/generate_tables.py          # rewrite the outputs
    python3 tools/dbcs/generate_tables.py --check  # exit 1 on drift
"""

from __future__ import annotations

import argparse
import json
import sys
import textwrap
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
SBCS_PATH = REPO_ROOT / "src" / "sbcs_codepage.cpp"
TARGETS_DIR = REPO_ROOT / "tests" / "oracle" / "targets"
REGENERATION_GUIDANCE = "Regenerate with: python3 tools/dbcs/generate_tables.py"

# Adaptive parameter: k = bit_length(acc >> ACC_SHIFT) - K_BIAS; must match dbcs_table.cpp.
ACC_SHIFT = 1
K_BIAS = 2
ACC_INIT = 32

SBCS_BEGIN = "// BEGIN GENERATED (tools/dbcs/generate_tables.py)\n"
SBCS_END = "// END GENERATED\n"


@dataclass(frozen=True)
class MeasuredSbcs:
    title: str
    symbol: str
    target: str

    @property
    def golden(self) -> Path:
        return TARGETS_DIR / self.target / "golden" / "locale_tokens.golden.json"


MEASURED_SBCS = (
    MeasuredSbcs("Mac Thai", "kMacThaiHigh", "mac-365-th_TH"),
    MeasuredSbcs("Mac Cyrillic", "kMacCyrillicHigh", "mac-365-ru_RU"),
)


@dataclass(frozen=True)
class DbcsShape:
    """Code space of one table: lead bytes and up to two trail-byte ranges.

    Bytes are in the table's own space (the code minus `dbcs_code_bias` in
    src/eval/dbcs_table.cpp), so the 94x94 tables use 1..94 for both.
    """

    lead_first: int
    lead_count: int
    trails: tuple[tuple[int, int], ...]  # (first, count), at most two

    @property
    def trail_total(self) -> int:
        return sum(count for _, count in self.trails)

    @property
    def slots(self) -> int:
        return self.lead_count * self.trail_total

    def positions(self):
        """Yield (lead, trail) for every slot in row-major order."""
        for lead in range(self.lead_first, self.lead_first + self.lead_count):
            for first, count in self.trails:
                for trail in range(first, first + count):
                    yield lead, trail


GRID_94 = DbcsShape(1, 94, ((1, 94),))
BIG5_SHAPE = DbcsShape(0xA1, 0xF9 - 0xA1 + 1, ((0x40, 0x3F), (0xA1, 0x5E)))


def big5_assigned(lead: int, trail: int) -> bool:
    """Plain Big5; the ETEN rows (C6A1-C8FE, F9D6-F9FE) are absent on Mac Excel
    (kana at C6A5.. read as unmapped in the zh-TW code_char_jp_probes)."""
    if lead in (0xC7, 0xC8) or (lead == 0xC6 and trail >= 0xA1):
        return False
    return not (lead == 0xF9 and trail >= 0xD6)


@dataclass(frozen=True)
class DbcsTable:
    name: str
    title: str
    codec: str
    accessor: str
    budget_kb: int
    shape: DbcsShape = GRID_94
    # Drops a slot whose character encodes to a different code, so a character
    # duplicated in the code page maps back from exactly one code.
    canonical_only: bool = False
    assigned: Callable[[int, int], bool] = lambda lead, trail: True

    @property
    def path(self) -> Path:
        return REPO_ROOT / "src" / "eval" / f"dbcs_{self.name}_table.cpp"


TABLES = (
    DbcsTable("jis0208", "JIS X 0208", "iso2022_jp", "jis0208_grid", 10),
    DbcsTable("gb2312", "GB 2312", "gb2312", "gb2312_grid", 11),
    DbcsTable("ksx1001", "KS X 1001", "euc_kr", "ksx1001_grid", 11),
    DbcsTable("big5", "Big5", "big5", "big5_grid", 15, BIG5_SHAPE, True, big5_assigned),
)


def decode_cell(table: DbcsTable, lead: int, trail: int) -> int:
    """Return the Unicode code point at (lead, trail), or 0 when unassigned."""
    if not table.assigned(lead, trail):
        return 0
    if table.codec == "iso2022_jp":
        payload = b"\x1b\x24\x42" + bytes([lead + 0x20, trail + 0x20])
    elif table.shape is GRID_94:
        payload = bytes([lead + 0xA0, trail + 0xA0])
    else:
        payload = bytes([lead, trail])
    try:
        text = payload.decode(table.codec)
    except (UnicodeDecodeError, ValueError):
        return 0
    if len(text) != 1 or ord(text) > 0xFFFF:
        return 0
    if table.canonical_only and text.encode(table.codec, errors="replace") != payload:
        return 0
    return ord(text)


def cells_for(table: DbcsTable) -> list[int]:
    return [decode_cell(table, lead, trail) for lead, trail in table.shape.positions()]


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
    shape = table.shape
    pairs = list(shape.trails) + [(0, 0)] * (2 - len(shape.trails))
    trails = [f"{{{first}, {count}}}" for first, count in pairs]
    cells = cells_for(table)
    encoded = encode_cells(cells)
    mapped = sum(1 for cp in cells if cp)
    body = []
    for start in range(0, len(encoded), 19):
        chunk = encoded[start : start + 19]
        body.append("    " + ", ".join(f"0x{b:02X}" for b in chunk) + ",\n")
    head = textwrap.dedent(
        f"""\
        // GENERATED by tools/dbcs/generate_tables.py from Python's {table.codec} codec; do not edit.
        // {table.title}: {mapped} of {table.shape.slots} cells mapped, {len(encoded)} encoded bytes.
        // @size-budget: {table.budget_kb} KB

        #include <cstddef>
        #include <cstdint>

        #include "eval/dbcs_table.h"

        namespace formulon {{
        namespace eval {{
        namespace dbcs_detail {{
        namespace {{

        constexpr DbcsShape kShape = {{{shape.lead_first}, {shape.lead_count}, {{{trails[0]}, {trails[1]}}}}};
        constexpr std::size_t kSlots = {shape.slots};

        constexpr std::uint8_t kEncoded[] = {{
        """
    )
    tail = textwrap.dedent(
        f"""\
        }};

        }}  // namespace

        const DbcsGrid& {table.accessor}() noexcept {{
          static const DbcsCells<kSlots> cells(kEncoded, sizeof(kEncoded));
          static const DbcsGrid grid = {{kShape, cells.unicode}};
          return grid;
        }}

        }}  // namespace dbcs_detail
        }}  // namespace eval
        }}  // namespace formulon
        """
    )
    return head + "".join(body) + tail


def measured_high_bytes(golden: Path) -> list[int]:
    data = json.loads(golden.read_text(encoding="utf-8"))
    for case in data["cases"]:
        if case["id"] == "char_high_bytes_unicode":
            values = [int(v) for v in case["expect"]["value"].split(",")]
            if len(values) != 128 or any(v < 0 for v in values):
                raise ValueError("char_high_bytes_unicode must list 128 decodable bytes")
            return values
    raise ValueError(f"char_high_bytes_unicode missing from {golden}")


def render_sbcs_block(tables: list[tuple[MeasuredSbcs, list[int]]]) -> str:
    lines = [SBCS_BEGIN]
    for table, values in tables:
        lines.append(f"// {table.title}, measured: locale_tokens.char_high_bytes_unicode ({table.target}).\n")
        lines.append(f"constexpr std::array<std::uint32_t, 128> {table.symbol} = {{{{\n")
        for start in range(0, len(values), 13):
            chunk = values[start : start + 13]
            lines.append("    " + ", ".join(f"0x{v:04X}u" for v in chunk) + ",\n")
        lines.append("}};\n")
    lines.append(SBCS_END)
    return "".join(lines)


def splice_sbcs(source: str, block: str) -> str:
    begin = source.find(SBCS_BEGIN)
    end = source.find(SBCS_END)
    if begin < 0 or end < begin:
        raise ValueError("generated single-byte markers missing from sbcs_codepage.cpp")
    return source[:begin] + block + source[end + len(SBCS_END) :]


def outputs(sbcs_source: str) -> dict[Path, str]:
    rendered = {table.path: render_dbcs(table) for table in TABLES}
    measured = [(table, measured_high_bytes(table.golden)) for table in MEASURED_SBCS]
    rendered[SBCS_PATH] = splice_sbcs(sbcs_source, render_sbcs_block(measured))
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
