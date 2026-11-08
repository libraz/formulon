#!/usr/bin/env python3
"""Generate the error display-name tables from ``kErrorTable`` in src/value.h.

Each output holds an ``<error-display-names:begin>`` / ``<error-display-names:end>``
marker pair, placed by the file's owner; this tool replaces what lies between
the two marker lines and never creates them. A missing marker is an error.

* ``packages/npm/common.mjs``: ``ErrorDisplayName``, a frozen table keyed by
  ``ErrorCode`` members, so the region follows the ``ErrorCode`` table.
* ``src/wasm/formulon.d.ts``: its ``Readonly<Record<ErrorCode, string>>``
  declaration.
* ``packages/python/formulon/_workbook.py``: ``_ERROR_DISPLAY_NAMES``, a
  module-level dict keyed by ``ErrorCode`` member name.

Modes:
  (default)  rewrite the outputs
  --check    exit 1 when an output differs from what would be generated
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

REPO_ROOT = Path(__file__).resolve().parents[2]
VALUE_H = REPO_ROOT / "src/value.h"
NPM_COMMON = REPO_ROOT / "packages/npm/common.mjs"
WASM_DTS = REPO_ROOT / "src/wasm/formulon.d.ts"
PY_WORKBOOK = REPO_ROOT / "packages/python/formulon/_workbook.py"
OUTPUTS = {NPM_COMMON: "js", WASM_DTS: "dts", PY_WORKBOOK: "py"}

BEGIN = "<error-display-names:begin>"
END = "<error-display-names:end>"

Row = Tuple[str, str]  # (ErrorCode enumerator, display name)


class GenError(ValueError):
    """An input or output file does not have the expected shape."""


def _enumerators(text: str) -> List[str]:
    match = re.search(r"enum\s+class\s+ErrorCode\b[^{]*\{(.*?)\};", text, re.S)
    if match is None:
        raise GenError("src/value.h: `enum class ErrorCode` not found")
    body = re.sub(r"//[^\n]*|/\*.*?\*/", "", match.group(1), flags=re.S)
    return [m.group(1) for part in body.split(",") if (m := re.match(r"\s*([A-Za-z_]\w*)", part))]


def parse_error_table(text: str) -> List[Row]:
    """`(enumerator, display name)` per `kErrorTable` row, in `ErrorCode` order."""
    names = _enumerators(text)
    match = re.search(r"kErrorTable\[\]\s*=\s*\{(.*?)\n\};", text, re.S)
    if match is None:
        raise GenError("src/value.h: `kErrorTable` not found")
    rows = re.findall(r'\{\s*"([^"]*)"\s*,\s*-?\d+\s*\}\s*,?\s*(?://\s*ErrorCode::(\w+))?', match.group(1))
    if len(rows) != len(names):
        raise GenError(f"src/value.h: kErrorTable has {len(rows)} rows for {len(names)} ErrorCode enumerators")
    for (_, labelled), name in zip(rows, names):
        if labelled and labelled != name:
            raise GenError(f"src/value.h: the kErrorTable row of ErrorCode::{name} is labelled ErrorCode::{labelled}")
    return [(name, display) for (display, _), name in zip(rows, names)]


def upper_snake(camel: str) -> str:
    return re.sub(r"(?<=[a-z])(?=[A-Z])", "_", camel).upper()


def render(kind: str, rows: Sequence[Row]) -> str:
    """The text between the marker lines for an output of `kind` (`js`, `dts` or `py`)."""
    if kind == "js":
        body = "".join(f"  [ErrorCode.{name}]: '{display}',\n" for name, display in rows)
        return f"export const ErrorDisplayName = Object.freeze({{\n{body}}});\n"
    if kind == "dts":
        listed = "".join(f" *  - `{name}`: `{display}`\n" for name, display in rows)
        return (
            "/** The display text of each {@link ErrorCode}, as Excel spells the error in a cell:\n"
            f"{listed} *  An unknown code reads `undefined`. */\n"
            "export const ErrorDisplayName: Readonly<Record<ErrorCode, string>>;\n"
        )
    if kind == "py":
        body = "".join(f'    "{upper_snake(name)}": "{display}",\n' for name, display in rows)
        return f"_ERROR_DISPLAY_NAMES = {{\n{body}}}\n"
    raise ValueError(kind)


def fill(text: str, kind: str, rows: Sequence[Row], label: str) -> str:
    """`text` with the lines between its marker lines replaced by the generated table."""
    begin = text.find(BEGIN)
    end = text.find(END, begin + 1) if begin >= 0 else -1
    if begin < 0 or end < 0:
        raise GenError(f"{label}: no {BEGIN} / {END} region; its owner places the empty region")
    start = text.index("\n", begin) + 1
    stop = text.rindex("\n", 0, end) + 1
    if stop < start:
        raise GenError(f"{label}: {BEGIN} and {END} share a line")
    return text[:start] + render(kind, rows) + text[stop:]


def outputs() -> Dict[Path, Tuple[str, str]]:
    """`{path: (current text, generated text)}` for every output."""
    rows = parse_error_table(VALUE_H.read_text(encoding="utf-8"))
    result: Dict[Path, Tuple[str, str]] = {}
    for path, kind in OUTPUTS.items():
        current = path.read_text(encoding="utf-8") if path.is_file() else ""
        result[path] = (current, fill(current, kind, rows, str(path.relative_to(REPO_ROOT))))
    return result


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--check", action="store_true", help="exit 1 when an output is not what would be generated")
    args = parser.parse_args(argv)
    try:
        generated = outputs()
    except GenError as exc:
        print(f"gen_error_display_names: {exc}", file=sys.stderr)
        return 1
    stale = [path for path, (current, text) in generated.items() if current != text]
    if args.check:
        for path in stale:
            print(
                f"gen_error_display_names: {path.relative_to(REPO_ROOT)} is out of date; "
                "run tools/codegen/gen_error_display_names.py",
                file=sys.stderr,
            )
        return 1 if stale else 0
    for path in stale:
        path.write_text(generated[path][1], encoding="utf-8")
        print(f"gen_error_display_names: wrote {path.relative_to(REPO_ROOT)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
