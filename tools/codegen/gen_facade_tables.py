#!/usr/bin/env python3
"""Generate the `Workbook` / `Sheet` facade tables from ``tools/api/names.yaml``.

Outputs (read through ``facade_scope.py``, never edited by hand):

* ``packages/npm/common.mjs``: the ``// <facade-tables>`` region, which both npm
  packages share: ``WORKBOOK_METHODS`` and ``SHEET_METHODS`` (the public names
  the facade forwards or implements), ``SHEET_STRUCTURE_OPS`` (the ops after
  which the facade re-reads the sheet list), ``LIST_FAMILIES``
  (``{listName: [countName, getName]}``, the lists the facade assembles from a
  raw count and element getter), ``HANDWRITTEN`` (ops the ledger marks
  ``facade: handwritten``), ``CELL_ARGS`` (``{method: [[row, col], ...]}``, the
  argument positions of each one-cell ``(row, col)`` pair, the receiver
  excluded) and ``RECT_ARGS`` (``{method: [position, ...]}``, the ``CellRect``
  arguments), both from the ledger's ``cell_args`` / ``rect_args``, and
  ``REF_ARGS`` (``{method: [position, kind]}``, the ``*Ref`` argument) from its
  ``ref_args``.
* ``packages/python/formulon/sheet.py``: the same tables but ``REF_ARGS`` in its
  ``# <facade-tables>`` region, with the Python names. ``--check`` skips the
  file while it has no region or an empty one; ``--write`` fills an empty one.
* ``src/node_addon/generated/facade_scope.inc``: the sorted raw names of the
  sheet-scope methods, as string literals for a C++ array initialiser; the
  addon words its argument guards for these as ``Sheet.<name>``.
* ``tools/api/address_args.json``: ``CELL_ARGS`` and ``RECT_ARGS`` per language
  (``{"js": {"cell_args": ..., "rect_args": ...}, "py": ...}``), sorted, for
  the binding tests to read.

Modes:
  --write   rewrite the outputs
  --check   exit 1 when an output differs from what would be generated
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

REPO_ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(Path(__file__).resolve().parent))

import facade_scope  # noqa: E402

NPM_COMMON = REPO_ROOT / "packages/npm/common.mjs"
PYTHON_SHEET = REPO_ROOT / "packages/python/formulon/sheet.py"
NODE_SCOPE_INC = REPO_ROOT / "src/node_addon/generated/facade_scope.inc"
ADDRESS_ARGS_JSON = REPO_ROOT / "tools/api/address_args.json"

_REGION = {
    "js": re.compile(r"(// <facade-tables>[^\n]*\n)(.*?)(// </facade-tables>)", re.S),
    "py": re.compile(r"(# <facade-tables>[^\n]*\n)(.*?)(# </facade-tables>)", re.S),
}
_INDENT = "  "


class GenError(ValueError):
    """An output file does not have the expected shape."""


def _tables(language: str) -> Tuple[Dict[str, List[str]], Dict[str, Tuple[str, str]]]:
    ledger = facade_scope.load()
    names = {
        "WORKBOOK_METHODS": facade_scope.methods(ledger, "workbook", language),
        "SHEET_METHODS": facade_scope.methods(ledger, "sheet", language),
        "SHEET_STRUCTURE_OPS": facade_scope.structure_ops(ledger, language),
    }
    families = facade_scope.list_families(ledger, language)
    names["HANDWRITTEN"] = facade_scope.handwritten(ledger, language)
    return names, families


def _js_list(name: str, values: Sequence[str]) -> str:
    body = "".join(f"{_INDENT}'{v}',\n" for v in values)
    return f"const {name} = Object.freeze([\n{body}]);\n"


def _py_positions(value: list) -> str:
    items = [f"({p[0]}, {p[1]})" if isinstance(p, list) else str(p) for p in value]
    return "(" + ", ".join(items) + ("," if len(items) == 1 else "") + ")"


def _js_positions(value: object) -> str:
    """A position entry as the formatter lays it out: nested pairs one per line when there are several."""
    if isinstance(value, list) and len(value) > 1 and all(isinstance(v, list) for v in value):
        rows = "".join(f"{_INDENT * 2}{json.dumps(v)},\n" for v in value)
        return f"[\n{rows}{_INDENT}]"
    return json.dumps(value)


def address_tables(language: str, ledger: Optional[facade_scope.Ledger] = None) -> str:
    """The `CELL_ARGS` and `RECT_ARGS` tables of `language`'s facade region, and in JS `REF_ARGS`, which the
    generic wrapper reads (Python methods resolve their reference themselves)."""
    ledger = ledger or facade_scope.load()
    tables = {
        "CELL_ARGS": facade_scope.cell_args(ledger, language),
        "RECT_ARGS": facade_scope.rect_args(ledger, language),
    }
    out = []
    for table, entries in tables.items():
        if language == "js":
            body = "".join(f"{_INDENT}{name}: {_js_positions(value)},\n" for name, value in entries.items())
            out.append(f"const {table} = Object.freeze({{\n{body}}});\n")
        else:
            body = "".join(f'    "{name}": {_py_positions(value)},\n' for name, value in entries.items())
            out.append(f"{table} = {{\n{body}}}\n")
    if language == "js":
        refs = facade_scope.ref_args(ledger, language)
        body = "".join(f"{_INDENT}{name}: [{pos}, '{kind}'],\n" for name, (pos, kind) in refs.items())
        out.append(f"const REF_ARGS = Object.freeze({{\n{body}}});\n")
    return "".join(out)


def address_args_json(ledger: Optional[facade_scope.Ledger] = None) -> str:
    """`tools/api/address_args.json`: both tables per language, keys sorted, one method per line."""
    ledger = ledger or facade_scope.load()
    languages = []
    for language in sorted(facade_scope.LANGUAGES):
        tables = []
        for table, entries in (
            ("cell_args", facade_scope.cell_args(ledger, language)),
            ("rect_args", facade_scope.rect_args(ledger, language)),
        ):
            rows = ",\n".join(f"      {json.dumps(name)}: {json.dumps(value)}" for name, value in entries.items())
            tables.append(f'    "{table}": {{\n{rows}\n    }}' if rows else f'    "{table}": {{}}')
        languages.append(f'  "{language}": {{\n' + ",\n".join(tables) + "\n  }")
    return "{\n" + ",\n".join(languages) + "\n}\n"


def js_region() -> str:
    names, families = _tables("js")
    out = [_js_list("WORKBOOK_METHODS", names["WORKBOOK_METHODS"])]
    out.append(_js_list("SHEET_METHODS", names["SHEET_METHODS"]))
    out.append(_js_list("SHEET_STRUCTURE_OPS", names["SHEET_STRUCTURE_OPS"]))
    body = "".join(f"{_INDENT}{k}: ['{c}', '{g}'],\n" for k, (c, g) in families.items())
    out.append(f"const LIST_FAMILIES = Object.freeze({{\n{body}}});\n")
    out.append(_js_list("HANDWRITTEN", names["HANDWRITTEN"]))
    out.append(address_tables("js"))
    return "".join(out)


def _py_tuple(name: str, values: Sequence[str]) -> str:
    body = "".join(f'    "{v}",\n' for v in values)
    return f"{name} = (\n{body})\n"


def py_region() -> str:
    names, families = _tables("py")
    out = [_py_tuple("WORKBOOK_METHODS", names["WORKBOOK_METHODS"])]
    out.append(_py_tuple("SHEET_METHODS", names["SHEET_METHODS"]))
    out.append(_py_tuple("SHEET_STRUCTURE_OPS", names["SHEET_STRUCTURE_OPS"]))
    body = "".join(f'    "{k}": ("{c}", "{g}"),\n' for k, (c, g) in families.items())
    out.append(f"LIST_FAMILIES = {{\n{body}}}\n")
    out.append(_py_tuple("HANDWRITTEN", names["HANDWRITTEN"]))
    out.append(address_tables("py"))
    return "".join(out)


def node_scope_inc() -> str:
    names = facade_scope.raw_sheet_methods(facade_scope.load(), "js")
    head = (
        "// Generated by tools/codegen/gen_facade_tables.py from tools/api/names.yaml; do not edit.\n"
        "// Raw method names whose first argument is the sheet index the `Sheet` facade supplies, sorted.\n"
    )
    return head + "".join(f'"{n}",\n' for n in names)


def _splice(text: str, language: str, region: str, path: Path) -> str:
    match = _REGION[language].search(text)
    if match is None:
        raise GenError(f"{path}: no facade-tables region")
    return text[: match.start(2)] + region + text[match.end(2) :]


def _empty_region(text: Optional[str], language: str) -> bool:
    match = _REGION[language].search(text or "")
    return match is not None and not match.group(2).strip()


def outputs() -> List[Tuple[Path, Optional[str]]]:
    """`(path, generated text)` per output; the text is None for a skipped output."""
    result: List[Tuple[Path, Optional[str]]] = []
    common = NPM_COMMON.read_text(encoding="utf-8")
    result.append((NPM_COMMON, _splice(common, "js", js_region(), NPM_COMMON)))
    sheet = PYTHON_SHEET.read_text(encoding="utf-8") if PYTHON_SHEET.is_file() else None
    if sheet is not None and _REGION["py"].search(sheet):
        result.append((PYTHON_SHEET, _splice(sheet, "py", py_region(), PYTHON_SHEET)))
    else:
        result.append((PYTHON_SHEET, None))
    result.append((NODE_SCOPE_INC, node_scope_inc()))
    result.append((ADDRESS_ARGS_JSON, address_args_json()))
    return result


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--write", action="store_true", help="rewrite the outputs")
    mode.add_argument("--check", action="store_true", help="exit 1 when an output is not what would be generated")
    args = parser.parse_args(argv)
    try:
        generated = outputs()
    except GenError as exc:
        print(f"gen_facade_tables: {exc}", file=sys.stderr)
        return 1
    stale = []
    for path, text in generated:
        rel = path.relative_to(REPO_ROOT)
        if text is None:
            print(f"gen_facade_tables: skipped {rel} (no facade-tables region)", file=sys.stderr)
            continue
        current = path.read_text(encoding="utf-8") if path.is_file() else None
        if current == text:
            continue
        if args.check and path == PYTHON_SHEET and _empty_region(current, "py"):
            print(f"gen_facade_tables: skipped {rel} (empty facade-tables region)", file=sys.stderr)
            continue
        if args.check:
            stale.append(str(rel))
        else:
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(text, encoding="utf-8")
            print(f"gen_facade_tables: wrote {rel}")
    if stale:
        print(
            "gen_facade_tables: out of date, run tools/codegen/gen_facade_tables.py --write: " + ", ".join(stale),
            file=sys.stderr,
        )
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
