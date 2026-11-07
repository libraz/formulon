#!/usr/bin/env python3
"""Generate the facade tables from ``tools/api/names.yaml``.

Every table is keyed ``Class.method`` (``Sheet.setColumnWidth``,
``Pivot.addField``) unless it is a per-class name list. Outputs (read through
``facade_scope.py``, never edited by hand):

* ``packages/npm/common.mjs``: the ``// <facade-tables>`` region, which both npm
  packages share: ``WORKBOOK_METHODS`` and ``SHEET_METHODS`` (the public names
  the facade forwards or implements), ``SHEET_STRUCTURE_OPS`` (the methods
  after which the facade re-reads the sheet list), ``LIST_FAMILIES``
  (``{method: [countName, getName]}``, the lists the facade assembles from a
  raw count and element getter), ``HANDWRITTEN`` (ops the ledger marks
  ``facade: handwritten``), ``HANDLE_FORWARDERS`` (``{method: rawName}``, null
  for a handwritten one) and ``HANDLE_PROPERTIES`` (``{property: rawGetter}``)
  of the handle classes, ``CELL_ARGS`` (``{method: [[row, col], ...]}``),
  ``RECT_ARGS`` (``{method: [position, ...]}``), ``SPAN_ARGS``
  (``{method: [[position, axis, form], ...]}``) and ``REF_ARGS``
  (``{method: [position, kind]}``) from the ledger's ``cell_args`` /
  ``rect_args`` / ``span_args`` / ``ref_args``, positions as the class method
  sees them; ``PATCH_KEYS`` (``{method: [position, [key, ...]]}``) of the
  ``patch: true`` ops and ``PATCH_NESTED_KEYS`` (``{method: {path: [key,
  ...]}}``) of the ``patch_nested`` ones: the keys of the patch type and of
  the getter's record in src/wasm/formulon.d.ts.
* ``packages/python/formulon/sheet.py``: the ``Workbook`` / ``Sheet`` tables but
  ``REF_ARGS`` and the patch tables in its ``# <facade-tables>`` region, with
  the Python names. ``--check`` skips the file while it has no region or an
  empty one; ``--write`` fills an empty one.
* ``packages/python/formulon/_handle_tables.py``: the same tables for the handle
  classes, the whole file generated.
* ``src/node_addon/generated/facade_scope.inc``: ``kSheetMethods``, the sorted
  raw names of the sheet-scope methods (the addon words their argument guards
  as ``Sheet.<name>``), and ``kHandleMethods``, a ``HandleMethod`` per raw
  method a handle class calls: its ``Class.method`` label and the count of
  leading raw arguments the handle supplies. It reads the ledger alone.
* ``tools/api/address_args.json``: ``cell_args``, ``rect_args`` and
  ``span_args`` per language, sorted, for the binding tests to read.

Modes:
  --write   rewrite the outputs
  --check   exit 1 when an output differs from what would be generated
  --only node-scope
            limit either mode to ``facade_scope.inc``, which needs neither the
            d.ts nor the facade sources
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
PYTHON_HANDLE_TABLES = REPO_ROOT / "packages/python/formulon/_handle_tables.py"
NODE_SCOPE_INC = REPO_ROOT / "src/node_addon/generated/facade_scope.inc"
ADDRESS_ARGS_JSON = REPO_ROOT / "tools/api/address_args.json"

_REGION = {
    "js": re.compile(r"(// <facade-tables>[^\n]*\n)(.*?)(// </facade-tables>)", re.S),
    "py": re.compile(r"(# <facade-tables>[^\n]*\n)(.*?)(# </facade-tables>)", re.S),
}
_INDENT = "  "
_FACADE_CLASSES = ("Workbook.", "Sheet.")
ONLY_TARGETS = ("node-scope",)


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


def _py_value(value: object) -> str:
    """A table value as a Python literal: lists become tuples, strings take double quotes."""
    if isinstance(value, list):
        items = [_py_value(v) for v in value]
        return "(" + ", ".join(items) + ("," if len(items) == 1 else "") + ")"
    if isinstance(value, str):
        return json.dumps(value)
    if value is None:
        return "None"
    return str(value)


def _js_value(value: object) -> str:
    return json.dumps(value).replace('"', "'")


def _js_positions(value: object) -> str:
    """A position entry as the formatter lays it out: nested pairs one per line when there are several."""
    if isinstance(value, list) and len(value) > 1 and all(isinstance(v, list) for v in value):
        rows = "".join(f"{_INDENT * 2}{_js_value(v)},\n" for v in value)
        return f"[\n{rows}{_INDENT}]"
    return _js_value(value)


def _js_key(name: str) -> str:
    return f"'{name}'"


def _py_dict(name: str, body: str) -> str:
    """`name = {...}` as the formatter lays it out: an empty table stays on one line."""
    return f"{name} = {{\n{body}}}\n" if body else f"{name} = {{}}\n"


def _js_frozen(name: str, body: str) -> str:
    """`const name = Object.freeze({...});` as the formatter lays it out: an empty table stays on one line."""
    return f"const {name} = Object.freeze({{\n{body}}});\n" if body else f"const {name} = Object.freeze({{}});\n"


_JS_IDENTIFIER = re.compile(r"^[A-Za-z_$][A-Za-z0-9_$]*$")


def _js_bare_key(name: str) -> str:
    """An object key as the formatter writes it: unquoted when it is an identifier."""
    return name if _JS_IDENTIFIER.match(name) else f"'{name}'"


def _select(entries: Dict[str, object], facade: Optional[bool]) -> Dict[str, object]:
    """The rows of `Workbook` / `Sheet` (facade True), of the handle classes (False), or all (None)."""
    if facade is None:
        return entries
    return {k: v for k, v in entries.items() if k.startswith(_FACADE_CLASSES) == facade}


def address_tables(language: str, ledger: Optional[facade_scope.Ledger] = None, facade: Optional[bool] = None) -> str:
    """The `CELL_ARGS`, `RECT_ARGS` and `SPAN_ARGS` tables of `language`, and in JS `REF_ARGS`, which the
    generic wrapper reads (Python methods resolve their reference themselves)."""
    ledger = ledger or facade_scope.load()
    tables = {
        "CELL_ARGS": facade_scope.cell_args(ledger, language),
        "RECT_ARGS": facade_scope.rect_args(ledger, language),
        "SPAN_ARGS": facade_scope.span_args(ledger, language),
    }
    if language == "js":
        tables["REF_ARGS"] = facade_scope.ref_args(ledger, language)
    out = []
    for table, entries in tables.items():
        entries = _select(entries, facade)
        if language == "js":
            body = "".join(f"{_INDENT}{_js_key(name)}: {_js_positions(value)},\n" for name, value in entries.items())
            out.append(_js_frozen(table, body))
        else:
            body = "".join(f'    "{name}": {_py_value(value)},\n' for name, value in entries.items())
            out.append(_py_dict(table, body))
    return "".join(out)


_LINE_WIDTH = 120


def _js_patch_entry(name: str, position: int, keys: Sequence[str]) -> str:
    """One `PATCH_KEYS` entry as the formatter lays it out at its line width."""
    quoted = [f"'{k}'" for k in keys]
    inline = f"[{', '.join(quoted)}]"
    name = _js_key(name)
    one_line = f"{_INDENT}{name}: [{position}, {inline}],"
    if len(one_line) <= _LINE_WIDTH:
        return one_line + "\n"
    inner = f"{_INDENT * 2}{inline},"
    if len(inner) > _LINE_WIDTH:
        rows = "".join(f"{_INDENT * 3}{q},\n" for q in quoted)
        inner = f"{_INDENT * 2}[\n{rows}{_INDENT * 2}],"
    return f"{_INDENT}{name}: [\n{_INDENT * 2}{position},\n{inner}\n{_INDENT}],\n"


def patch_table(ledger: Optional[facade_scope.Ledger] = None, dts: Optional[str] = None) -> str:
    """The JS `PATCH_KEYS` table: each patch op's patch position and the keys it accepts."""
    entries = facade_scope.patch_keys(ledger or facade_scope.load(), dts)
    body = "".join(_js_patch_entry(name, position, keys) for name, (position, keys) in entries.items())
    return f"const PATCH_KEYS = Object.freeze({{\n{body}}});\n"


def _js_key_list(indent: str, name: str, keys: Sequence[str]) -> str:
    """`name: [...]` at `indent`, one key per line once the inline form passes the line width."""
    quoted = [f"'{k}'" for k in keys]
    one_line = f"{indent}{name}: [{', '.join(quoted)}],"
    if len(one_line) <= _LINE_WIDTH:
        return one_line + "\n"
    rows = "".join(f"{indent}{_INDENT}{q},\n" for q in quoted)
    return f"{indent}{name}: [\n{rows}{indent}],\n"


def patch_nested_table(ledger: Optional[facade_scope.Ledger] = None, dts: Optional[str] = None) -> str:
    """The JS `PATCH_NESTED_KEYS` table: for each nested patch op, the keys each path into its patch accepts."""
    entries = facade_scope.patch_nested_keys(ledger or facade_scope.load(), dts)
    body = []
    for name, paths in entries.items():
        rows = "".join(_js_key_list(_INDENT * 2, _js_bare_key(path), keys) for path, keys in paths.items())
        body.append(f"{_INDENT}{_js_key(name)}: {{\n{rows}{_INDENT}}},\n")
    return f"const PATCH_NESTED_KEYS = Object.freeze({{\n{''.join(body)}}});\n"


def handle_tables(language: str, ledger: Optional[facade_scope.Ledger] = None) -> str:
    """`HANDLE_FORWARDERS` and `HANDLE_PROPERTIES` of `language`."""
    ledger = ledger or facade_scope.load()
    forwarders = facade_scope.handle_forwarders(ledger, language)
    properties = facade_scope.handle_properties(ledger, language)
    if language == "js":
        body = "".join(f"{_INDENT}{_js_key(k)}: {_js_key(v) if v else 'null'},\n" for k, v in forwarders.items())
        out = f"const HANDLE_FORWARDERS = Object.freeze({{\n{body}}});\n"
        body = "".join(f"{_INDENT}{_js_key(k)}: {_js_key(v)},\n" for k, v in properties.items())
        return out + f"const HANDLE_PROPERTIES = Object.freeze({{\n{body}}});\n"
    body = "".join(f'    "{k}": {_py_value(v)},\n' for k, v in forwarders.items())
    out = _py_dict("HANDLE_FORWARDERS", body)
    body = "".join(f'    "{k}": {_py_value(v)},\n' for k, v in properties.items())
    return out + _py_dict("HANDLE_PROPERTIES", body)


def address_args_json(ledger: Optional[facade_scope.Ledger] = None) -> str:
    """`tools/api/address_args.json`: the position tables per language, keys sorted, one method per line."""
    ledger = ledger or facade_scope.load()
    languages = []
    for language in sorted(facade_scope.LANGUAGES):
        tables = []
        for table, entries in (
            ("cell_args", facade_scope.cell_args(ledger, language)),
            ("rect_args", facade_scope.rect_args(ledger, language)),
            ("span_args", facade_scope.span_args(ledger, language)),
        ):
            rows = ",\n".join(f"      {json.dumps(name)}: {json.dumps(value)}" for name, value in entries.items())
            tables.append(f'    "{table}": {{\n{rows}\n    }}' if rows else f'    "{table}": {{}}')
        languages.append(f'  "{language}": {{\n' + ",\n".join(tables) + "\n  }")
    return "{\n" + ",\n".join(languages) + "\n}\n"


def _list_families(language: str, families: Dict[str, Tuple[str, str]]) -> str:
    if language == "js":
        body = "".join(f"{_INDENT}{_js_key(k)}: ['{c}', '{g}'],\n" for k, (c, g) in families.items())
        return f"const LIST_FAMILIES = Object.freeze({{\n{body}}});\n"
    body = "".join(f'    "{k}": ("{c}", "{g}"),\n' for k, (c, g) in families.items())
    return _py_dict("LIST_FAMILIES", body)


def js_region() -> str:
    ledger = facade_scope.load()
    names, families = _tables("js")
    out = [_js_list("WORKBOOK_METHODS", names["WORKBOOK_METHODS"])]
    out.append(_js_list("SHEET_METHODS", names["SHEET_METHODS"]))
    out.append(_js_list("SHEET_STRUCTURE_OPS", names["SHEET_STRUCTURE_OPS"]))
    out.append(_list_families("js", families))
    out.append(_js_list("HANDWRITTEN", names["HANDWRITTEN"]))
    out.append(handle_tables("js", ledger))
    out.append(address_tables("js", ledger))
    out.append(patch_table(ledger))
    out.append(patch_nested_table(ledger))
    return "".join(out)


def _py_tuple(name: str, values: Sequence[str]) -> str:
    body = "".join(f'    "{v}",\n' for v in values)
    return f"{name} = (\n{body})\n"


def py_region() -> str:
    """The `Workbook` / `Sheet` tables of sheet.py's region."""
    ledger = facade_scope.load()
    names, families = _tables("py")
    out = [_py_tuple("WORKBOOK_METHODS", names["WORKBOOK_METHODS"])]
    out.append(_py_tuple("SHEET_METHODS", names["SHEET_METHODS"]))
    out.append(_py_tuple("SHEET_STRUCTURE_OPS", names["SHEET_STRUCTURE_OPS"]))
    out.append(_list_families("py", _select(families, True)))
    out.append(_py_tuple("HANDWRITTEN", names["HANDWRITTEN"]))
    out.append(address_tables("py", ledger, facade=True))
    return "".join(out)


def py_handle_tables() -> str:
    """The whole of `_handle_tables.py`: the handle-class tables with the Python names."""
    ledger = facade_scope.load()
    _names, families = _tables("py")
    head = (
        '"""Facade tables of the handle classes, keyed `Class.method`.\n\n'
        "Generated by tools/codegen/gen_facade_tables.py from tools/api/names.yaml; do not edit.\n"
        '"""\n\n'
    )
    out = [head, handle_tables("py", ledger)]
    out.append(_list_families("py", _select(families, False)))
    out.append(address_tables("py", ledger, facade=False))
    return "".join(out)


def node_scope_inc() -> str:
    """`facade_scope.inc`: the sheet-scope raw methods and the raw methods of the handle classes."""
    ledger = facade_scope.load()
    names = facade_scope.raw_sheet_methods(ledger, "js")
    handles = facade_scope.handle_methods(ledger)
    out = [
        "// Generated by tools/codegen/gen_facade_tables.py from tools/api/names.yaml; do not edit.\n",
        "\n",
        "// Raw method names whose first argument is the sheet index the `Sheet` facade supplies, sorted.\n",
        "constexpr const char* kSheetMethods[] = {\n",
    ]
    out += [f'    "{n}",\n' for n in names]
    out += [
        "};\n",
        "\n",
        "/// A raw method a handle class calls: the `Class.method` label its argument errors carry, and the count\n",
        "/// of leading raw arguments the handle supplies.\n",
        "struct HandleMethod {\n",
        "  const char* raw;\n",
        "  const char* label;\n",
        "  int supplied;\n",
        "};\n",
        "\n",
        "// Sorted by raw name.\n",
        "constexpr HandleMethod kHandleMethods[] = {\n",
    ]
    out += [f'    {{"{raw}", "{label}", {supplied}}},\n' for raw, label, supplied in handles]
    out.append("};\n")
    return "".join(out)


def _splice(text: str, language: str, region: str, path: Path) -> str:
    match = _REGION[language].search(text)
    if match is None:
        raise GenError(f"{path}: no facade-tables region")
    return text[: match.start(2)] + region + text[match.end(2) :]


def _empty_region(text: Optional[str], language: str) -> bool:
    match = _REGION[language].search(text or "")
    return match is not None and not match.group(2).strip()


def outputs(only: Optional[str] = None) -> List[Tuple[Path, Optional[str]]]:
    """`(path, generated text)` per output; the text is None for a skipped output."""
    if only == "node-scope":
        return [(NODE_SCOPE_INC, node_scope_inc())]
    result: List[Tuple[Path, Optional[str]]] = []
    common = NPM_COMMON.read_text(encoding="utf-8")
    result.append((NPM_COMMON, _splice(common, "js", js_region(), NPM_COMMON)))
    sheet = PYTHON_SHEET.read_text(encoding="utf-8") if PYTHON_SHEET.is_file() else None
    if sheet is not None and _REGION["py"].search(sheet):
        result.append((PYTHON_SHEET, _splice(sheet, "py", py_region(), PYTHON_SHEET)))
    else:
        result.append((PYTHON_SHEET, None))
    result.append((PYTHON_HANDLE_TABLES, py_handle_tables()))
    result.append((NODE_SCOPE_INC, node_scope_inc()))
    result.append((ADDRESS_ARGS_JSON, address_args_json()))
    return result


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--write", action="store_true", help="rewrite the outputs")
    mode.add_argument("--check", action="store_true", help="exit 1 when an output is not what would be generated")
    parser.add_argument("--only", choices=ONLY_TARGETS, help="limit the run to one output")
    args = parser.parse_args(argv)
    try:
        generated = outputs(args.only)
    except (GenError, ValueError) as exc:
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
