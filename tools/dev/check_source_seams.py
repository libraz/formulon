#!/usr/bin/env python3
"""Source-level seam guards.

Two invariants in this codebase are enforced by "there is exactly one place
that does this", and neither is expressible in the type system. A second
implementation is not a compile error -- it is a silent loss of whatever the
single implementation was there to guarantee. This script fails when one
appears.

Checks:
  array-alloc      Every `ArrayValue` header is allocated by
                   `eval::allocate_array_value`.

                   An `ArrayValue` is a `(rows, cols)` header over a flat
                   `Value[]` of length `rows * cols`. On wasm32 `size_t` is
                   32 bits, so a bare product of two `uint32_t` axes wraps:
                   the buffer comes back short while the caller's fill loop
                   still walks the full logical extent, writing past the
                   arena block. The seam is the one place that does the
                   bounds-checked multiplication, the grid-extent check, the
                   cell ceiling and the arena-exhaustion check together.

                   The check targets the header rather than the `Value[]`
                   buffer: a header cannot exist without one of these calls,
                   whereas plain `Value[]` scratch buffers (argument vectors,
                   for instance) are unrelated to array shape and legitimately
                   allocate on their own.

  xml-passthrough  Raw-XML retention goes through `io::append_raw_xml` /
                   `io::raw_xml` / `io::capture_unknown_children`.

                   A part the reader consumes drops out of the unknown-part
                   passthrough sweep, so anything in it the model does not
                   represent is lost on the next save unless it is retained
                   verbatim. Every reader that consumes a part needs the same
                   retention format; before it was shared, seven of them had
                   grown their own copy, and each new copy is a chance to get
                   the format -- or the "which children are unmodelled"
                   question -- subtly different.

  all              Run every check above (default).

Stdlib only; no build artifacts or network access required.
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
SRC_ROOT = REPO_ROOT / "src"

SOURCE_SUFFIXES = (".cpp", ".h", ".cc")


def _scan(pattern: re.Pattern[str], allowed: set[Path]) -> list[str]:
    """Returns `path:line: text` for every match outside `allowed`."""
    hits: list[str] = []
    for path in sorted(SRC_ROOT.rglob("*")):
        if path.suffix not in SOURCE_SUFFIXES or path in allowed:
            continue
        for lineno, line in enumerate(path.read_text(encoding="utf-8").splitlines(), start=1):
            if pattern.search(line):
                hits.append(f"{path.relative_to(REPO_ROOT)}:{lineno}: {line.strip()}")
    return hits


def _report(name: str, hits: list[str], remedy: str) -> bool:
    if not hits:
        print(f"{name}: ok")
        return True
    print(f"{name}: FAILED", file=sys.stderr)
    for hit in hits:
        print(f"  {hit}", file=sys.stderr)
    print(f"\n{remedy}", file=sys.stderr)
    return False


def check_array_alloc() -> bool:
    seam = SRC_ROOT / "eval" / "array_alloc.cpp"
    hits = _scan(re.compile(r"create\s*<\s*ArrayValue\s*>"), {seam})
    return _report(
        "array-alloc",
        hits,
        "Call `eval::allocate_array_value(rows, cols, arena, out_buffer, max_cells)`\n"
        f"instead; only {seam.relative_to(REPO_ROOT)} may allocate the header directly.",
    )


def check_xml_passthrough() -> bool:
    seam = SRC_ROOT / "io" / "xml_utils.cpp"
    hits = _scan(re.compile(r"\bformat_raw\b"), {seam})
    return _report(
        "xml-passthrough",
        hits,
        "Call `io::append_raw_xml` / `io::raw_xml` for one element, or\n"
        "`io::capture_unknown_children` for the unmodelled children of a parent;\n"
        f"only {seam.relative_to(REPO_ROOT)} may drive a pugixml writer directly.",
    )


CHECKS = {
    "array-alloc": check_array_alloc,
    "xml-passthrough": check_xml_passthrough,
}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("check", nargs="?", default="all", choices=[*CHECKS, "all"])
    args = parser.parse_args()

    selected = CHECKS.values() if args.check == "all" else [CHECKS[args.check]]
    return 0 if all([fn() for fn in selected]) else 1


if __name__ == "__main__":
    sys.exit(main())
