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
  the getter's record in src/wasm/formulon.d.ts. Once an op carries
  ``namespace: ooxml``, ``OOXML_WORKBOOK_METHODS`` and ``OOXML_SHEET_METHODS``
  list the names published under ``.ooxml``, which leave the two class lists.
* ``packages/python/formulon/sheet.py``: the ``Workbook`` / ``Sheet`` tables but
  ``REF_ARGS`` and the patch tables in its ``# <facade-tables>`` region, with
  the Python names. ``--check`` skips the file while it has no region or an
  empty one; ``--write`` fills an empty one.
* ``RECORD_RANGE_FIELDS`` (``{Record: {field: type}}``) in the
  ``<record-range-fields>`` region of ``common.mjs`` and of
  ``packages/python/formulon/_inputs.py``: the fields of the input records in
  src/wasm/formulon.d.ts typed ``CellRect``, ``SheetRect`` or ``CellAddress``
  (``[]`` for an array), or typed by another input record that has such a
  field, less the ledger's ``a1_exempt`` fields. A discriminated input (a
  union type alias of interfaces) has the fields of every alternative. Python keys and nested record
  names are the ledger's Python names, fields snake_case. Both files follow the
  ``sheet.py`` rule for an absent or empty region.
* ``METHOD_ARITY`` (``{Class.method: count}``), ``TYPED_ARGS``
  (``{Class.method: [[position, type], ...]}``) in the ``<method-signatures>``
  region of ``common.mjs``, read from the facade interfaces of
  src/wasm/formulon.d.ts: the most positional parameters any overload
  declares, and the parameters typed by an address record or a range type
  other than a plain ``CellRect``; ``DEFAULT_ARGS`` (``{Class.method:
  [[position, default], ...]}``), the ledger's ``params`` defaults at their
  positions in the longest overload. Same rule for an absent or empty region.
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

from tools.dev.binding_drift.api_conventions_arguments import _js_fields  # noqa: E402
from tools.dev.binding_drift.api_conventions_bindings import (  # noqa: E402
    OOXML_CLASSES,
    SHEET_CLASS,
    WORKBOOK_CLASS,
    JsSurface,
    _top_union,
)
from tools.dev.binding_drift.api_conventions_shapes import _strip_null  # noqa: E402

NPM_COMMON = REPO_ROOT / "packages/npm/common.mjs"
PYTHON_SHEET = REPO_ROOT / "packages/python/formulon/sheet.py"
PYTHON_HANDLE_TABLES = REPO_ROOT / "packages/python/formulon/_handle_tables.py"
NODE_SCOPE_INC = REPO_ROOT / "src/node_addon/generated/facade_scope.inc"
ADDRESS_ARGS_JSON = REPO_ROOT / "tools/api/address_args.json"
PYTHON_INPUTS = REPO_ROOT / "packages/python/formulon/_inputs.py"
WASM_DTS = REPO_ROOT / "src/wasm/formulon.d.ts"


def _region_re(comment: str, tag: str) -> "re.Pattern[str]":
    return re.compile(rf"({comment} <{tag}>[^\n]*\n)(.*?)({comment} </{tag}>)", re.S)


_REGION = {"js": _region_re("//", "facade-tables"), "py": _region_re("#", "facade-tables")}
_RANGE_REGION = {"js": _region_re("//", "record-range-fields"), "py": _region_re("#", "record-range-fields")}
_SIGNATURE_REGION = _region_re("//", "method-signatures")
OOXML_NAMESPACE = "ooxml"
# The d.ts types a record field holds an address in; the facade also takes an A1 string for each.
RANGE_TYPES = ("CellRect", "SheetRect", "CellAddress")
INPUT_ROLES = frozenset({"input", "patch", "both"})
_INDENT = "  "
_FACADE_CLASSES = ("Workbook.", "Sheet.")
ONLY_TARGETS = ("node-scope",)


class GenError(ValueError):
    """An output file does not have the expected shape."""


def ooxml_methods(ledger: facade_scope.Ledger, scope: str, language: str) -> List[str]:
    """Sorted public names of the `scope` ops and lists the ledger publishes under `.ooxml`."""
    out = set()
    for entry in ledger.entries.values():
        if isinstance(entry, dict) and entry.get("namespace") == OOXML_NAMESPACE and entry.get("scope") == scope:
            name = facade_scope.public_name(entry, language)
            if name is not None:
                out.add(name)
    return sorted(out)


def _tables(
    language: str, ledger: Optional[facade_scope.Ledger] = None
) -> Tuple[Dict[str, List[str]], Dict[str, Tuple[str, str]]]:
    ledger = ledger or facade_scope.load()
    names = {}
    for table, scope in (("WORKBOOK_METHODS", "workbook"), ("SHEET_METHODS", "sheet")):
        ooxml = ooxml_methods(ledger, scope, language)
        names[table] = [n for n in facade_scope.methods(ledger, scope, language) if n not in ooxml]
        names[f"OOXML_{table}"] = ooxml
    names["SHEET_STRUCTURE_OPS"] = facade_scope.structure_ops(ledger, language)
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


def _has_ooxml(names: Dict[str, List[str]]) -> bool:
    return bool(names["OOXML_WORKBOOK_METHODS"] or names["OOXML_SHEET_METHODS"])


def js_region(ledger: Optional[facade_scope.Ledger] = None) -> str:
    ledger = ledger or facade_scope.load()
    names, families = _tables("js", ledger)
    out = [_js_list("WORKBOOK_METHODS", names["WORKBOOK_METHODS"])]
    out.append(_js_list("SHEET_METHODS", names["SHEET_METHODS"]))
    if _has_ooxml(names):
        out.append(_js_list("OOXML_WORKBOOK_METHODS", names["OOXML_WORKBOOK_METHODS"]))
        out.append(_js_list("OOXML_SHEET_METHODS", names["OOXML_SHEET_METHODS"]))
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


def py_region(ledger: Optional[facade_scope.Ledger] = None) -> str:
    """The `Workbook` / `Sheet` tables of sheet.py's region."""
    ledger = ledger or facade_scope.load()
    names, families = _tables("py", ledger)
    out = [_py_tuple("WORKBOOK_METHODS", names["WORKBOOK_METHODS"])]
    out.append(_py_tuple("SHEET_METHODS", names["SHEET_METHODS"]))
    if _has_ooxml(names):
        out.append(_py_tuple("OOXML_WORKBOOK_METHODS", names["OOXML_WORKBOOK_METHODS"]))
        out.append(_py_tuple("OOXML_SHEET_METHODS", names["OOXML_SHEET_METHODS"]))
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


_ARRAY_RE = re.compile(r"(?:readonly\s+)?(?:\((.*)\)|([A-Za-z_$][\w$]*))\[\]|ReadonlyArray<(.*)>", re.S)


def _range_type(text: str, known: Sequence[str]) -> Optional[str]:
    """`T` or `T[]` when the field type `text` holds one `T` of `known` or an array of them, `string` and absent
    alternatives aside; None otherwise. An array that also takes one `T` stays `T[]`."""
    bases = set()
    array = False
    for alt in _strip_null(text, True):
        match = _ARRAY_RE.fullmatch(alt.strip())
        if match:
            inner = [a for a in _strip_null(next(g for g in match.groups() if g), True) if a != "string"]
            if len(inner) != 1:
                return None
            bases.add(inner[0].strip())
            array = True
        elif alt != "string":
            bases.add(alt.strip())
    if len(bases) != 1:
        return None
    base = bases.pop()
    return (base + "[]" if array else base) if base in known else None


def _record_fields(surface: JsSurface, name: str) -> Dict[str, str]:
    """Field name -> type of the input record `name`: an interface, or a union type alias of interfaces (a
    discriminated input, alternatives module-private or exported), whose fields are every alternative's."""
    pool = {**surface.private, **surface.decls}
    decl = pool.get(name)
    if decl is None or decl.kind != "type":
        return _js_fields(pool, surface.raw, name) or {}
    alternatives = [_js_fields(pool, surface.raw, alt) for alt in _top_union(decl.text)]
    if len(alternatives) < 2 or any(alt is None for alt in alternatives):
        return {}
    types: Dict[str, List[str]] = {}
    for alt in alternatives:
        for field, ftype in alt.items():
            if ftype not in types.setdefault(field, []):
                types[field].append(ftype)
    return {field: " | ".join(kinds) for field, kinds in types.items()}


def _snake(name: str) -> str:
    return re.sub(r"(?<!^)(?=[A-Z])", "_", name).lower()


def record_range_fields(
    ledger: Optional[facade_scope.Ledger] = None, dts: Optional[str] = None
) -> Dict[str, Dict[str, str]]:
    """`{InputRecord: {field: type}}` of the d.ts input records, JS names, records sorted, fields in order.

    A field whose type is another input record counts once that record has an address field of its own.
    """
    ledger = ledger or facade_scope.load()
    surface = JsSurface(WASM_DTS.read_text(encoding="utf-8") if dts is None else dts)
    exempt: Dict[str, set] = {}
    for entry in ledger.entries.values():
        if not isinstance(entry, dict) or not isinstance(entry.get("js"), str):
            continue
        if entry.get("kind") == "record" and entry.get("role") in INPUT_ROLES or "a1_exempt" in entry:
            exempt[entry["js"]] = set((entry.get("a1_exempt") or {}).get("fields") or [])
    fields = {name: _record_fields(surface, name) for name in exempt}
    out: Dict[str, Dict[str, str]] = {}
    while True:
        known = (*RANGE_TYPES, *out)
        found = {}
        for name, members in fields.items():
            row = {
                field: kind
                for field, ftype in members.items()
                if field not in exempt[name] and (kind := _range_type(ftype, known)) is not None
            }
            if row:
                found[name] = row
        if found == out:
            return dict(sorted(out.items()))
        out = found


def _py_record_names(ledger: facade_scope.Ledger) -> Dict[str, str]:
    return {
        e["js"]: e["py"]
        for e in ledger.entries.values()
        if isinstance(e, dict) and isinstance(e.get("js"), str) and isinstance(e.get("py"), str)
    }


def range_fields_region(language: str, ledger: Optional[facade_scope.Ledger] = None, dts: Optional[str] = None) -> str:
    """`RECORD_RANGE_FIELDS` of `language`; Python drops a record or a nested record the package does not name."""
    ledger = ledger or facade_scope.load()
    table = record_range_fields(ledger, dts)
    if language == "js":
        body = "".join(_js_field_row(_INDENT, name, row) for name, row in table.items())
        return _js_frozen("RECORD_RANGE_FIELDS", body)
    py_names = _py_record_names(ledger)
    rows = []
    for name, row in table.items():
        if name not in py_names:
            continue
        cells = []
        for field, kind in row.items():
            base, array = (kind[:-2], "[]") if kind.endswith("[]") else (kind, "")
            target = base if base in RANGE_TYPES else py_names.get(base)
            if target is not None:
                cells.append(f'"{_snake(field)}": "{target}{array}"')
        if cells:
            rows.append(_py_mapping_row(py_names[name], cells))
    return _py_dict("RECORD_RANGE_FIELDS", "".join(rows))


def _js_field_row(indent: str, name: str, row: Dict[str, str]) -> str:
    """`Name: { field: 'Type', ... },` at `indent`, one field per line once the inline form passes the width."""
    pairs = [f"{_js_bare_key(f)}: '{t}'" for f, t in row.items()]
    one_line = f"{indent}{_js_bare_key(name)}: {{ {', '.join(pairs)} }},"
    if len(one_line) <= _LINE_WIDTH:
        return one_line + "\n"
    body = "".join(f"{indent}{_INDENT}{p},\n" for p in pairs)
    return f"{indent}{_js_bare_key(name)}: {{\n{body}{indent}}},\n"


def _py_mapping_row(name: str, cells: Sequence[str]) -> str:
    """`"Name": {...},` as the formatter lays it out at its line width."""
    one_line = f'    "{name}": {{{", ".join(cells)}}},'
    if len(one_line) <= _LINE_WIDTH:
        return one_line + "\n"
    body = "".join(f"        {c},\n" for c in cells)
    return f'    "{name}": {{\n{body}    }},\n'


def _facade_interfaces(ledger: facade_scope.Ledger) -> Dict[str, str]:
    """d.ts interface -> the `Class` its methods are keyed under: `Workbook` / `Sheet` and their `ooxml` layers,
    and every handle class the ledger forwards to."""
    out = {WORKBOOK_CLASS: WORKBOOK_CLASS, SHEET_CLASS: SHEET_CLASS}
    out.update({layer: owner for owner, layer in OOXML_CLASSES.items()})
    handles = {*facade_scope.handle_forwarders(ledger, "js"), *facade_scope.handle_properties(ledger, "js")}
    for key in handles:
        owner = key.split(".", 1)[0]
        out.setdefault(owner, owner)
    return out


def method_signatures(
    ledger: Optional[facade_scope.Ledger] = None, dts: Optional[str] = None
) -> Tuple[Dict[str, int], Dict[str, List[List[object]]]]:
    """`(METHOD_ARITY, TYPED_ARGS)` of the facade classes, keyed `Class.method`, sorted.

    The arity is the most positional parameters any overload declares (the coordinate form). A method with
    `CELL_ARGS` or `SPAN_ARGS` must also declare exactly the string form, one parameter fewer per pair or span,
    which the facade derives at run time. `TYPED_ARGS` lists `[position, type]` per parameter typed by an input
    record of `RECORD_RANGE_FIELDS` or by another range type; a plain `CellRect` parameter is `RECT_ARGS`'.
    """
    ledger = ledger or facade_scope.load()
    text = WASM_DTS.read_text(encoding="utf-8") if dts is None else dts
    surface = JsSurface(text)
    records = record_range_fields(ledger, text)
    known = (*RANGE_TYPES, *records)
    arities: Dict[str, set] = {}
    typed: Dict[str, Dict[int, str]] = {}
    for interface, owner in _facade_interfaces(ledger).items():
        decl = surface.decls.get(interface)
        for member in decl.members if decl is not None and decl.kind == "interface" else ():
            if not member.is_method:
                continue
            key = f"{owner}.{member.name}"
            arities.setdefault(key, set()).add(len(member.params))
            for position, (_name, ptype) in enumerate(member.params):
                kind = _range_type(ptype, known)
                if kind is None or kind == "CellRect":
                    continue
                slots = typed.setdefault(key, {})
                if slots.setdefault(position, kind) != kind:
                    raise GenError(f"{key} argument {position + 1} is typed {slots[position]} and {kind}")
    collapsed = facade_scope.cell_args(ledger, "js")
    spans = facade_scope.span_args(ledger, "js")
    for key in sorted({*collapsed, *spans}):
        if key not in arities:
            continue
        most = max(arities[key])
        want = {most, most - len(collapsed.get(key, ())) - len(spans.get(key, ()))}
        if arities[key] != want:
            raise GenError(f"{key}: overloads take {sorted(arities[key])} arguments, expected {sorted(want)}")
    arity = {key: max(counts) for key, counts in sorted(arities.items())}
    args = {key: [[p, k] for p, k in sorted(slots.items())] for key, slots in sorted(typed.items())}
    return arity, args


def _camel(snake: str) -> str:
    head, *rest = snake.split("_")
    return head + "".join(w[:1].upper() + w[1:] for w in rest)


def default_args(
    ledger: Optional[facade_scope.Ledger] = None, dts: Optional[str] = None
) -> Dict[str, List[List[object]]]:
    """`DEFAULT_ARGS`: `{Class.method: [[position, default], ...]}` of the facade methods whose ledger op gives a
    parameter a `default` (`params: {<c name>: {default: <value>}}`), sorted. The position is the parameter's in
    the longest overload (the coordinate form); `'today'` stays the ledger's clock token."""
    ledger = ledger or facade_scope.load()
    surface = JsSurface(WASM_DTS.read_text(encoding="utf-8") if dts is None else dts)
    longest: Dict[str, List[str]] = {}
    for interface, owner in _facade_interfaces(ledger).items():
        decl = surface.decls.get(interface)
        for member in decl.members if decl is not None and decl.kind == "interface" else ():
            key = f"{owner}.{member.name}"
            if member.is_method and len(member.params) >= len(longest.get(key, ())):
                longest[key] = [name for name, _type in member.params]
    out: Dict[str, List[List[object]]] = {}
    for entry in ledger.entries.values():
        params = entry.get("params") if isinstance(entry, dict) and entry.get("kind") == "op" else None
        if not isinstance(params, dict):
            continue
        defaults = {n: s["default"] for n, s in params.items() if isinstance(s, dict) and "default" in s}
        for key in facade_scope.qualified(entry, "js", ledger.collections) if defaults else ():
            if key not in longest:
                continue
            slots = []
            for name, value in defaults.items():
                if _camel(name) not in longest[key]:
                    raise GenError(f"{key} declares no parameter `{_camel(name)}` for the ledger default of `{name}`")
                slots.append([longest[key].index(_camel(name)), value])
            out[key] = sorted(slots, key=lambda slot: slot[0])
    return dict(sorted(out.items()))


def signatures_region(ledger: Optional[facade_scope.Ledger] = None, dts: Optional[str] = None) -> str:
    """The JS `METHOD_ARITY`, `TYPED_ARGS` and `DEFAULT_ARGS` tables."""
    arity, typed = method_signatures(ledger, dts)
    out = _js_frozen("METHOD_ARITY", "".join(f"{_INDENT}{_js_key(k)}: {n},\n" for k, n in arity.items()))
    body = "".join(f"{_INDENT}{_js_key(k)}: {_js_positions(v)},\n" for k, v in typed.items())
    out += _js_frozen("TYPED_ARGS", body)
    body = "".join(f"{_INDENT}{_js_key(k)}: {_js_positions(v)},\n" for k, v in default_args(ledger, dts).items())
    return out + _js_frozen("DEFAULT_ARGS", body)


def _splice(text: str, pattern: "re.Pattern[str]", region: str, path: Path) -> str:
    match = pattern.search(text)
    if match is None:
        raise GenError(f"{path}: no generated region")
    return text[: match.start(2)] + region + text[match.end(2) :]


def _empty_region(text: Optional[str], pattern: "re.Pattern[str]") -> bool:
    match = pattern.search(text or "")
    return match is not None and not match.group(2).strip()


def _region_output(path: Path, pattern: "re.Pattern[str]", region) -> Tuple[Path, Optional[str]]:
    """`(path, text)` with `region()` spliced in, text None while the file has no such region."""
    text = path.read_text(encoding="utf-8") if path.is_file() else None
    if text is None or pattern.search(text) is None:
        return path, None
    return path, _splice(text, pattern, region(), path)


def outputs(only: Optional[str] = None) -> List[Tuple[Path, Optional[str]]]:
    """`(path, generated text)` per output; the text is None for a skipped output."""
    if only == "node-scope":
        return [(NODE_SCOPE_INC, node_scope_inc())]
    result: List[Tuple[Path, Optional[str]]] = []
    common = NPM_COMMON.read_text(encoding="utf-8")
    common = _splice(common, _REGION["js"], js_region(), NPM_COMMON)
    if _RANGE_REGION["js"].search(common):
        common = _splice(common, _RANGE_REGION["js"], range_fields_region("js"), NPM_COMMON)
    if _SIGNATURE_REGION.search(common):
        common = _splice(common, _SIGNATURE_REGION, signatures_region(), NPM_COMMON)
    result.append((NPM_COMMON, common))
    result.append(_region_output(PYTHON_SHEET, _REGION["py"], py_region))
    result.append(_region_output(PYTHON_INPUTS, _RANGE_REGION["py"], lambda: range_fields_region("py")))
    result.append((PYTHON_HANDLE_TABLES, py_handle_tables()))
    result.append((NODE_SCOPE_INC, node_scope_inc()))
    result.append((ADDRESS_ARGS_JSON, address_args_json()))
    return result


# Outputs whose region `--check` skips while it is empty; `--write` fills it.
_FILLABLE = {
    PYTHON_SHEET: (_REGION["py"],),
    PYTHON_INPUTS: (_RANGE_REGION["py"],),
    NPM_COMMON: (_RANGE_REGION["js"], _SIGNATURE_REGION),
}


def _skippable(path: Path, current: Optional[str], text: str) -> bool:
    """`--check` passes an output whose only difference is a fillable region that is still empty."""
    for pattern in _FILLABLE.get(path, ()):
        if _empty_region(current, pattern):
            match = pattern.search(text)
            if match is not None and current == text[: match.start(2)] + text[match.end(2) :]:
                return True
    return False


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
            print(f"gen_facade_tables: skipped {rel} (no generated region)", file=sys.stderr)
            continue
        current = path.read_text(encoding="utf-8") if path.is_file() else None
        if current == text:
            continue
        if args.check and _skippable(path, current, text):
            print(f"gen_facade_tables: skipped {rel} (empty generated region)", file=sys.stderr)
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
