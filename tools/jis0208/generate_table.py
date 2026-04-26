#!/usr/bin/env python3
"""Generates the JIS X 0208 -> Unicode reverse table for Mac Excel ja-JP CODE/CHAR parity.

Mac Excel 365 ja-JP CODE returns the JIS X 0208 row-cell encoding for DBCS
characters (e.g. CODE("あ") = 9250 = 0x2422 = row 4, cell 2). CHAR is the
inverse and decodes the same encoding back to Unicode.

This script regenerates the dense reverse lookup table baked into
src/eval/jis0208_table.cpp. The forward direction (Unicode -> JIS) is
performed by linear scan over the same table at runtime to keep the WASM
binary small (per wasm-size-guardian decision 2026-04-27).

The hand-written lookup helpers (lookup_unicode_to_jis0208 and
lookup_jis0208_to_unicode) live in a separate translation unit
(src/eval/jis0208_table_lookup.cpp) so this generator can clobber the
data file freely without touching code.

Usage:
    python3 tools/jis0208/generate_table.py > src/eval/jis0208_table.cpp

The output is deterministic and depends only on Python's iso2022_jp codec,
which implements JIS X 0208-1983 identically to Mac Excel's CODE/CHAR
(verified against tests/oracle/golden/code_char_jp_probes.golden.json).
"""

from __future__ import annotations

import sys
import textwrap


def jis_to_unicode(row: int, cell: int) -> str | None:
    if not (1 <= row <= 94 and 1 <= cell <= 94):
        return None
    payload = b"\x1b\x24\x42" + bytes([row + 0x20, cell + 0x20])
    try:
        return payload.decode("iso2022_jp")
    except (UnicodeDecodeError, ValueError):
        return None


def main() -> None:
    table: list[int] = []
    for r in range(1, 95):
        for c in range(1, 95):
            u = jis_to_unicode(r, c)
            cp = ord(u) if (u and len(u) == 1) else 0
            table.append(cp)

    count_mapped = sum(1 for x in table if x != 0)

    out = sys.stdout.write
    out(
        textwrap.dedent(
            f"""\
            // Copyright 2026 libraz. Licensed under the MIT License.
            //
            // GENERATED FILE — do not edit by hand. Regenerate via:
            //   python3 tools/jis0208/generate_table.py > src/eval/jis0208_table.cpp
            //
            // JIS X 0208 -> Unicode dense reverse lookup table for Mac Excel
            // ja-JP CODE/CHAR parity. See src/eval/jis0208_table.h for usage.
            //
            // Layout: one uint16_t per (row, cell) slot, row-major. Index =
            // (row - 1) * 94 + (cell - 1) where row, cell are 1-based JIS X 0208
            // indices. Sentinel value 0 marks unassigned slots; U+0000 itself
            // never appears in JIS X 0208 so the sentinel is unambiguous.
            //
            // Source-of-truth: Python iso2022_jp codec, which implements
            // JIS X 0208-1983 byte-for-byte identical to Mac Excel CODE/CHAR
            // (verified empirically against the 38-case probe suite at
            // tests/oracle/cases/code_char_jp_probes.yaml on Mac Excel 16.108.1
            // ja-JP, 2026-04-27).
            //
            // Mapped entries: {count_mapped} / {len(table)} slots ({len(table)*2} bytes raw).
            //
            // @size-budget: 18 KB

            #include "eval/jis0208_table.h"

            #include <cstdint>

            namespace formulon {{
            namespace eval {{

            const std::uint16_t kJis0208Reverse[kJis0208RowCount * kJis0208CellCount] = {{
            """
        )
    )
    for i in range(0, len(table), 16):
        chunk = table[i : i + 16]
        out("    " + ", ".join(f"0x{x:04X}" for x in chunk) + ",\n")
    out(
        textwrap.dedent(
            """\
            };

            }  // namespace eval
            }  // namespace formulon
            """
        )
    )


if __name__ == "__main__":
    main()
