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
  ...]}}``) of the ``patch_nested`` ones: the keys of the patch type in
  src/wasm/formulon.d.ts and the ledger ``read_only_keys`` of its record, a
  dotted one at its path. ``SPAN_ARGS`` forms are ``span`` (``first, last``),
  ``start_count`` and ``single`` (one row or column, or an A1 string naming
  one). Once an op carries
  ``namespace: ooxml``, ``OOXML_WORKBOOK_METHODS`` and ``OOXML_SHEET_METHODS``
  list the names published under ``.ooxml``, which leave the two class lists.
* ``packages/python/formulon/_sheet.py``: the ``Workbook`` / ``Sheet`` tables but
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
  ``_sheet.py`` rule for an absent or empty region.
* ``METHOD_ARITY`` (``{Class.method: count}``), ``TYPED_ARGS``
  (``{Class.method: [[position, type], ...]}``) in the ``<method-signatures>``
  region of ``common.mjs``, read from the facade interfaces of
  src/wasm/formulon.d.ts: the most positional parameters any overload
  declares (``Workbook.<static>`` and ``Formulon.<fn>`` included), and the
  parameters typed by an address record or a range type other than a plain
  ``CellRect``; ``DEFAULT_ARGS`` (``{Class.method:
  [[position, default], ...]}``), the ledger's ``params`` defaults at their
  positions in the longest overload; ``ARG_ROLES`` (``{Class.method:
  [[position, role], ...]}``), the role of each integer parameter of the
  backing C function that the longest overload of a generated method
  declares as a ``number`` (an ``_index`` suffix optional in the match):
  ``sheet`` (``sheet_index``), ``position`` (``index`` / ``*_index``),
  ``coordinate`` (``row`` / ``col`` / ``first`` / ``last`` / ``*_row`` /
  ``*_col``), else ``quantity``. Same rule for an absent or empty region.
  Python methods are all handwritten and pass the role of each argument
  themselves, by the same rule.
* ``INPUT_KEYS`` (``{Class.method: [position, [key, ...]]}``) and
  ``INPUT_NESTED_KEYS`` (``{Class.method: {path: [key, ...]}}``) in the
  ``<input-keys>`` region of ``common.mjs``: for each non-patch op whose
  argument is a d.ts input record, the position of that argument and, sorted,
  the keys it accepts (the record's, plus its ledger ``read_only_keys``, the
  read-record fields a passed-back read carries and the binding ignores).
  Nested paths are
  dotted field names; a segment ending in ``[]`` descends into each element of
  an array field. A discriminated input has the keys of every alternative.
  ``READ_ONLY_KEYS`` (``{Class.method: [key, ...]}``) in the same region lists
  the ``read_only_keys`` of each op's input or patch record, camelCase, a
  nested one as a dotted path.
* ``PARAM_NAMES`` (``{Call: [name, ...]}``), ``BOOLEAN_ARGS`` (``{Call:
  [position, ...]}``) and ``C_FIELD_NAMES`` (``{c_field: {js, py}}``) in the
  ``<param-names>``, ``<boolean-args>`` and ``<c-field-names>`` regions of
  ``common.mjs``: the parameter names and the plain ``boolean`` parameters of
  each facade method's first signature, keyed by the class a consumer calls
  (``SheetOoxml.getColumnXfIndex``, ``Picture.update``, ``Workbook.load``,
  ``Formulon.evaluateFormula``), and the ledger fields whose JS name is not
  the camelCase of the C name. Same rule for an absent or empty region.
* ``PATCH_NULLABLE`` (``{Class.method: {key: type}}``) in the
  ``<patch-nullable>`` region of ``common.mjs``: the ledger's ``clearable:``
  fields (snake_case, emitted camelCase) of each patch op's patch record, with
  the ``typeof`` name checkPatch tests (``number``, ``string``, ``boolean``,
  ``array``, else ``object``). An op without a clearable field has no entry.
  Same rule for an absent or empty region.
* ``packages/python/formulon/_handle_tables.py``: the same tables for the handle
  classes, ``READ_ONLY_KEYS`` of every class (snake_case) and ``C_FIELD_NAMES``
  (``{c_field: py_name}`` where the two differ), the whole file generated.
* ``src/c_api/generated/binding_record_params.inc``: ``kRecordParams``, the
  name of the record parameter of each call taking an input record or a patch
  (``{"Sheet.updateView", "patch"}``), which the WASM and native readers root
  a field path at. A parameter typed by a union of object types, one a d.ts
  record (``date: DateInput``), roots the record it carries at its own name.
* ``src/c_api/generated/binding_param_names.inc``: ``kParamNames``, the
  ``PARAM_NAMES`` rows plus the label a binding words a handle's raw method
  with (``Drawing.setZOrder``), each call's names space-separated, which the
  WASM and native positional guards name a refused argument with. A call
  whose raw registration (the Node masks in ``workbook_class.cc``, after the
  arguments the facade supplies) has more checked positions than the d.ts
  signature, or a checked position the d.ts parameter's type cannot hold, is
  left out, so a refused raw argument is never given another one's name.
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
  --only record-params
            limit either mode to ``binding_record_params.inc``
  --only param-names
            limit either mode to ``binding_param_names.inc``
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
import gen_wasm_positional_guards as positional_guards  # noqa: E402

from tools.dev.binding_drift.api_conventions import parse_header  # noqa: E402
from tools.dev.binding_drift.api_conventions_arguments import _js_fields, _patch_keys  # noqa: E402
from tools.dev.binding_drift.api_conventions_bindings import (  # noqa: E402
    OOXML_CLASSES,
    SHEET_CLASS,
    WORKBOOK_CLASS,
    JsSurface,
    _top_union,
)
from tools.dev.binding_drift.api_conventions_shapes import _strip_null  # noqa: E402
from tools.dev.binding_drift.surface_files import PYTHON_SHEET  # noqa: E402

NPM_COMMON = REPO_ROOT / "packages/npm/common.mjs"
PYTHON_HANDLE_TABLES = REPO_ROOT / "packages/python/formulon/_handle_tables.py"
NODE_SCOPE_INC = REPO_ROOT / "src/node_addon/generated/facade_scope.inc"
RECORD_PARAMS_INC = REPO_ROOT / "src/c_api/generated/binding_record_params.inc"
PARAM_NAMES_INC = REPO_ROOT / "src/c_api/generated/binding_param_names.inc"
ADDRESS_ARGS_JSON = REPO_ROOT / "tools/api/address_args.json"
PYTHON_INPUTS = REPO_ROOT / "packages/python/formulon/_inputs.py"
WASM_DTS = REPO_ROOT / "src/wasm/formulon.d.ts"
C_HEADER = REPO_ROOT / "src/c_api/formulon_c.h"


def _region_re(comment: str, tag: str) -> "re.Pattern[str]":
    return re.compile(rf"({comment} <{tag}>[^\n]*\n)(.*?)({comment} </{tag}>)", re.S)


_REGION = {"js": _region_re("//", "facade-tables"), "py": _region_re("#", "facade-tables")}
_RANGE_REGION = {"js": _region_re("//", "record-range-fields"), "py": _region_re("#", "record-range-fields")}
_SIGNATURE_REGION = _region_re("//", "method-signatures")
_INPUT_REGION = _region_re("//", "input-keys")
_NULLABLE_REGION = _region_re("//", "patch-nullable")
_PARAM_NAMES_REGION = _region_re("//", "param-names")
_BOOLEAN_ARGS_REGION = _region_re("//", "boolean-args")
_C_FIELD_NAMES_REGION = _region_re("//", "c-field-names")
OOXML_NAMESPACE = "ooxml"
# The d.ts types a record field holds an address in; the facade also takes an A1 string for each.
RANGE_TYPES = ("CellRect", "SheetRect", "CellAddress")
# Address records and value types an op may take that are not input records (the set R39 uses).
_NOT_INPUT_RECORDS = frozenset({"CellRect", "SheetRect", "CellAddress", "CellRef", "CivilTime", "ErrorValue"})
INPUT_ROLES = frozenset({"input", "patch", "both"})
_INDENT = "  "
_FACADE_CLASSES = ("Workbook.", "Sheet.")
ONLY_TARGETS = ("node-scope", "record-params", "param-names")


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


def patch_nested_keys(
    ledger: Optional[facade_scope.Ledger] = None, dts: Optional[str] = None
) -> Dict[str, Dict[str, list]]:
    """`{Class.method: {path: [key, ...]}}`: the ledger's `patch_nested` paths, joined by the paths into a patch
    op's patch at and below each field typed by a `Patch<...>` alias (see `_nested_paths`)."""
    ledger = ledger or facade_scope.load()
    surface = JsSurface(WASM_DTS.read_text(encoding="utf-8") if dts is None else dts)
    out = {name: dict(paths) for name, paths in facade_scope.patch_nested_keys(ledger, dts).items()}
    for _key, entry in facade_scope._ops(ledger, ("op",)):
        if entry.get("patch") is not True:
            continue
        for qualified_name in facade_scope._qualified_js(ledger, entry):
            cls, name = qualified_name.split(".", 1)
            member = facade_scope._patch_member(surface, entry, name, cls)
            root = facade_scope._patch_type(member.params) if member is not None else None
            if root is None:
                continue
            fields = {f: [t] for f, t in (_patch_keys(surface.decls, surface.raw, root) or {}).items()}
            derived = _nested_paths(surface, ledger, fields, (root,), patch_aliases_only=True)
            derived = _with_read_only_paths(surface, ledger, root, fields, derived)
            if not derived:
                continue
            paths = out.setdefault(qualified_name, {})
            for path, keys in derived.items():
                paths[path] = sorted(set(paths.get(path, ())) | set(keys))
    return {name: dict(sorted(paths.items())) for name, paths in sorted(out.items())}


def _with_read_only_paths(
    surface: JsSurface,
    ledger: facade_scope.Ledger,
    record: str,
    fields: Dict[str, List[str]],
    paths: Dict[str, List[str]],
) -> Dict[str, List[str]]:
    """`paths` with the dotted `read_only_keys` of `record` added at their paths. A path no table row covers yet
    (a nested record a string alternative keeps out of `_nested_paths`) takes its record's fields as well."""
    out = {path: list(keys) for path, keys in paths.items()}
    for dotted, keys in facade_scope.read_only_keys(ledger, record)[1].items():
        found = next((p for p in out if p.replace("[]", "") == dotted), None)
        if found is None:
            found, nested = _resolve_path(surface, fields, dotted.split("."))
            if found is None:
                raise GenError(f"{record}: read_only_keys path {dotted} names no nested record")
            out[found] = sorted(_record_fields(surface, nested))
        out[found] = sorted(set(out[found]) | set(keys))
    return dict(sorted(out.items()))


def _resolve_path(
    surface: JsSurface, fields: Dict[str, List[str]], segments: Sequence[str]
) -> Tuple[Optional[str], Optional[str]]:
    """`(table path, record)` of the nested record `segments` lead to from `fields`, `string` alternatives aside;
    `(None, None)` when a segment is no field typed by one record."""
    path, record = "", None
    for segment in segments:
        types = fields.get(segment)
        if not types:
            return None, None
        element = _array_element(types[0])
        kinds = [k.strip() for k in _strip_null(element or types[0], True) if k not in ("string", "undefined")]
        if len(kinds) != 1 or not _record_fields(surface, kinds[0]):
            return None, None
        record = kinds[0]
        path = f"{path}{'.' if path else ''}{segment}{'[]' if element is not None else ''}"
        fields = {f: [t] for f, t in _record_fields(surface, record).items()}
    return path, record


def patch_nested_table(ledger: Optional[facade_scope.Ledger] = None, dts: Optional[str] = None) -> str:
    """The JS `PATCH_NESTED_KEYS` table: for each nested patch op, the keys each path into its patch accepts."""
    entries = patch_nested_keys(ledger, dts)
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


def _py_tuple_row(name: str, values: Sequence[str]) -> str:
    """`"name": (...),` as the formatter lays it out: one value per line once the inline form passes the width."""
    one_line = f'    "{name}": {_py_value(list(values))},'
    if len(one_line) <= _LINE_WIDTH:
        return one_line + "\n"
    rows = "".join(f"        {json.dumps(v)},\n" for v in values)
    return f'    "{name}": (\n{rows}    ),\n'


def _py_tuple(name: str, values: Sequence[str]) -> str:
    body = "".join(f'    "{v}",\n' for v in values)
    return f"{name} = (\n{body})\n"


def py_region(ledger: Optional[facade_scope.Ledger] = None) -> str:
    """The `Workbook` / `Sheet` tables of __sheet.py's region."""
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
    body = "".join(_py_tuple_row(k, v) for k, v in read_only_table("py", ledger).items())
    out.append(_py_dict("READ_ONLY_KEYS", body))
    renames = {c: n["py"] for c, n in c_field_names(ledger).items() if n["py"] is not None and n["py"] != c}
    out.append(_py_dict("C_FIELD_NAMES", "".join(f'    "{c}": "{py}",\n' for c, py in renames.items())))
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


_INTERFACE_EXTENDS_RE = re.compile(
    r"^(?:export\s+)?(?:declare\s+)?interface\s+([A-Za-z_$][\w$]*)\s+extends\s+([^{]+)\{", re.M
)


def _interface_members(surface: JsSurface, text: str, interface: str) -> list:
    """The members of d.ts `interface`, then the methods it inherits through `extends` and does not redeclare,
    so a facade subclass overriding an inherited method has a row of its own."""
    heritage = {m.group(1): [b.strip() for b in m.group(2).split(",")] for m in _INTERFACE_EXTENDS_RE.finditer(text)}
    own = surface.decls.get(interface)
    if own is None or own.kind != "interface":
        return []
    members = list(own.members)
    seen = {member.name for member in members}
    pending = list(heritage.get(interface, ()))
    while pending:
        decl = surface.decls.get(pending.pop(0))
        if decl is None or decl.kind != "interface":
            continue
        members.extend(m for m in decl.members if m.is_method and m.name not in seen)
        seen.update(m.name for m in decl.members)
        pending.extend(heritage.get(decl.name, ()))
    return members


def method_signatures(
    ledger: Optional[facade_scope.Ledger] = None, dts: Optional[str] = None
) -> Tuple[Dict[str, int], Dict[str, List[List[object]]]]:
    """`(METHOD_ARITY, TYPED_ARGS)` of the facade classes, keyed `Class.method`, sorted.

    The arity is the most positional parameters any overload declares (the coordinate form). A method with
    `CELL_ARGS` or `SPAN_ARGS` must also declare exactly the string form, one parameter fewer per pair or span
    (as many for a `single` span), which the facade derives at run time. `TYPED_ARGS` lists `[position, type]`
    per parameter typed by an input record of `RECORD_RANGE_FIELDS` or by another range type; a plain `CellRect`
    parameter is `RECT_ARGS`'.
    """
    ledger = ledger or facade_scope.load()
    text = WASM_DTS.read_text(encoding="utf-8") if dts is None else dts
    surface = JsSurface(text)
    records = record_range_fields(ledger, text)
    known = (*RANGE_TYPES, *records)
    arities: Dict[str, set] = {}
    typed: Dict[str, Dict[int, str]] = {}
    for interface, owner in _facade_interfaces(ledger).items():
        for member in _interface_members(surface, text, interface):
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
        # A `single` span keeps its position: the A1 form replaces the number one for one.
        folded = [span for span in spans.get(key, ()) if span[2] != "single"]
        want = {most, most - len(collapsed.get(key, ())) - len(folded)}
        if arities[key] != want:
            raise GenError(f"{key}: overloads take {sorted(arities[key])} arguments, expected {sorted(want)}")
    for key, counts in _module_signatures(surface).items():
        if key in arities:
            raise GenError(f"{key} is both a static or module function and an instance method")
        arities[key] = {len(params) for params in counts}
    arity = {key: max(counts) for key, counts in sorted(arities.items())}
    args = {key: [[p, k] for p, k in sorted(slots.items())] for key, slots in sorted(typed.items())}
    return arity, args


# Interfaces whose methods are not instance methods, and the `Class` their calls are labelled with.
_MODULE_INTERFACES = {"WorkbookCtor": "Workbook", "Formulon": "Formulon"}


def _module_signatures(surface: JsSurface) -> Dict[str, List[List[Tuple[str, str]]]]:
    """`{Class.fn: [params of each overload]}` of the `Workbook` statics and the `Formulon` module functions."""
    out: Dict[str, List[List[Tuple[str, str]]]] = {}
    for interface, owner in _MODULE_INTERFACES.items():
        decl = surface.decls.get(interface)
        for member in decl.members if decl is not None and decl.kind == "interface" else ():
            if member.is_method:
                out.setdefault(f"{owner}.{member.name}", []).append(list(member.params))
    return out


def _first_signatures(ledger: facade_scope.Ledger, surface: JsSurface, text: str) -> Dict[str, list]:
    """`{Call: params}` of the first signature of every facade method, keyed by the class a consumer calls it on:
    `Workbook`, `Sheet`, their `WorkbookOoxml` / `SheetOoxml` layers, each handle class (inherited methods
    included), `Workbook.<static>` and `Formulon.<fn>`."""
    interfaces = {interface: interface for interface in _facade_interfaces(ledger)}
    interfaces.update({layer: layer for layer in OOXML_CLASSES.values()})
    out: Dict[str, list] = {}
    for interface, owner in sorted(interfaces.items()):
        for member in _interface_members(surface, text, interface):
            if member.is_method:
                out.setdefault(f"{owner}.{member.name}", list(member.params))
    for key, overloads in _module_signatures(surface).items():
        out.setdefault(key, overloads[0])
    return dict(sorted(out.items()))


def param_names(ledger: Optional[facade_scope.Ledger] = None, dts: Optional[str] = None) -> Dict[str, List[str]]:
    """`PARAM_NAMES`: `{Call: [name, ...]}`, the parameter names of each facade method's first signature (see
    `_first_signatures`), for the methods that take any."""
    ledger = ledger or facade_scope.load()
    text = WASM_DTS.read_text(encoding="utf-8") if dts is None else dts
    signatures = _first_signatures(ledger, JsSurface(text), text)
    return {key: [name for name, _type in params] for key, params in signatures.items() if params}


_BOOLEAN_TYPES = frozenset({"boolean", "true", "false"})


def boolean_args(ledger: Optional[facade_scope.Ledger] = None, dts: Optional[str] = None) -> Dict[str, List[int]]:
    """`BOOLEAN_ARGS`: `{Call: [position, ...]}` of the parameters a facade method's first signature types
    `boolean` (an optional one included; a `boolean | null` takes `null` and is left out)."""
    ledger = ledger or facade_scope.load()
    text = WASM_DTS.read_text(encoding="utf-8") if dts is None else dts
    out: Dict[str, List[int]] = {}
    for key, params in _first_signatures(ledger, JsSurface(text), text).items():
        positions = [
            i for i, (_name, ptype) in enumerate(params) if set(_top_union(ptype)) - {"undefined"} <= _BOOLEAN_TYPES
        ]
        if positions:
            out[key] = positions
    return out


def c_field_names(ledger: Optional[facade_scope.Ledger] = None) -> Dict[str, Dict[str, Optional[str]]]:
    """`C_FIELD_NAMES`: `{c field: {js, py}}` of the ledger fields whose JS name is not the camelCase of the C
    name, sorted; `py` is None for a field Python does not publish. A C name two records publish under different
    names is an error, since an engine message names the field by its C name alone."""
    ledger = ledger or facade_scope.load()
    out: Dict[str, Dict[str, Optional[str]]] = {}
    for key, entry in sorted(ledger.entries.items()):
        if not isinstance(entry, dict) or entry.get("kind") != "field":
            continue
        c_name, js_name = entry.get("c"), entry.get("js")
        if not isinstance(c_name, str) or not isinstance(js_name, str) or js_name == _camel(c_name):
            continue
        py_name = entry.get("py") if isinstance(entry.get("py"), str) else None
        row = {"js": js_name, "py": py_name}
        if out.setdefault(c_name, row) != row:
            raise GenError(f"{key}: C field {c_name} is published as {row} and as {out[c_name]}")
    return dict(sorted(out.items()))


def _record_args(ledger: facade_scope.Ledger, surface: JsSurface) -> List[Tuple[dict, str, int, str, str]]:
    """`(entry, Class.method, position, parameter name, type)` of every op argument that is an input record or a
    patch, in the first signature of its JS member."""
    out = []
    for _key, entry in facade_scope._ops(ledger, ("op",)):
        for qualified_name in facade_scope._qualified_js(ledger, entry):
            cls, name = qualified_name.split(".", 1)
            member = facade_scope._patch_member(surface, entry, name, cls)
            if member is None and entry.get("namespace") == OOXML_NAMESPACE:
                member = facade_scope._patch_member(surface, entry, name, f"{cls}Ooxml")
            if member is None or not member.params:
                continue
            if entry.get("patch") is True:
                position = len(member.params) - 1
                ptype = facade_scope._patch_type(member.params)
            else:
                found = _input_arg(surface, ledger, entry, member)
                position, ptype = found if found is not None else (-1, None)
            if ptype is not None:
                out.append((entry, qualified_name, position, member.params[position][0], ptype))
    return out


def read_only_table(
    language: str, ledger: Optional[facade_scope.Ledger] = None, dts: Optional[str] = None
) -> Dict[str, List[str]]:
    """`READ_ONLY_KEYS`: `{Class.method: [key, ...]}` of each op whose input or patch record lists
    `read_only_keys`, sorted; JS keys camelCase, Python keys snake_case, a nested one as a dotted path."""
    ledger = ledger or facade_scope.load()
    surface = JsSurface(WASM_DTS.read_text(encoding="utf-8") if dts is None else dts)
    out: Dict[str, List[str]] = {}
    for entry, qualified_name, _position, _param, ptype in _record_args(ledger, surface):
        records = facade_scope.type_records(ledger, ptype)
        keys = sorted({key for record in records for key in record.get("read_only_keys") or ()})
        if not keys:
            continue
        if language == "js":
            out[qualified_name] = sorted(".".join(_camel(s) for s in k.split(".")) for k in keys)
            continue
        cls = qualified_name.split(".", 1)[0]
        for py_name in facade_scope.qualified(entry, "py", ledger.collections):
            if py_name.split(".", 1)[0] == cls:
                out[py_name] = keys
    return dict(sorted(out.items()))


def _alternatives(surface: JsSurface, text: str, depth: int = 0) -> List[str]:
    """The top-level alternatives of the d.ts type `text`, type aliases expanded, `null` / `undefined` left out."""
    pool = {**surface.private, **surface.decls}
    out: List[str] = []
    for alternative in (a.strip() for a in _top_union(text)):
        if alternative in ("null", "undefined"):
            continue
        decl = pool.get(alternative)
        alias = decl.text if decl is not None and decl.kind == "type" else None
        if decl is not None and decl.kind == "const":
            # An enum is a `const` object and a type alias of the same name; the surface keeps the `const`.
            found = re.search(r"\bexport\s+type\s+%s\s*=\s*([^;]*);" % re.escape(alternative), surface.raw)
            alias = found.group(1) if found is not None else None
        if alias is not None and depth < 8:
            out += _alternatives(surface, alias, depth + 1)
        else:
            out.append(alternative)
    return out


_NUMBER_LITERAL = re.compile(r"-?\d+(?:\.\d+)?")


def _is_object_type(alternative: str) -> bool:
    return (
        alternative not in _PRIMITIVES
        and alternative not in ("bigint", "symbol", "unknown", "any", "void", "never")
        and alternative[:1] not in "'\"`"
        and not _NUMBER_LITERAL.fullmatch(alternative)
    )


def _object_union_position(surface: JsSurface, params: Sequence[Tuple[str, str]]) -> Optional[int]:
    """The first parameter whose type is a union of two or more object types, one of them a record the d.ts
    declares (`DateInput`, `CivilTimeInput | Date`): the record it carries is rooted at that parameter's own name.
    A union of built-in objects alone (`BytesInput`) carries no record."""
    pool = {**surface.private, **surface.decls}
    for position, (_name, ptype) in enumerate(params):
        alternatives = _alternatives(surface, ptype)
        if (
            len(alternatives) > 1
            and all(_is_object_type(a) for a in alternatives)
            and any(a in pool and pool[a].kind == "interface" for a in alternatives)
        ):
            return position
    return None


def record_params(ledger: Optional[facade_scope.Ledger] = None, dts: Optional[str] = None) -> Dict[str, str]:
    """`{call: parameter}`: the name of the record parameter of each call that takes an input record or a patch.
    A call is the facade's `Class.method`, an `.ooxml` op's `<Class>Ooxml.method` too, `Workbook.<static>`,
    `Formulon.<fn>`, and the label a binding words a handle's raw method with (`Drawing.update`, the class every
    publishing class shares)."""
    ledger = ledger or facade_scope.load()
    surface = JsSurface(WASM_DTS.read_text(encoding="utf-8") if dts is None else dts)
    out: Dict[str, str] = {}
    by_raw: Dict[str, set] = {}
    text = surface.raw
    signatures = _first_signatures(ledger, surface, text)
    for entry, qualified_name, position, param, _ptype in _record_args(ledger, surface):
        cls, name = qualified_name.split(".", 1)
        union = _object_union_position(surface, signatures.get(qualified_name, []))
        if union is not None and union < position:
            param = signatures[qualified_name][union][0]
        out[qualified_name] = param
        if entry.get("namespace") == OOXML_NAMESPACE:
            out[f"{cls}Ooxml.{name}"] = param
        raw = facade_scope.raw_name(entry, "js")
        if raw is not None:
            by_raw.setdefault(raw, set()).add(param)
    for interface, owner in _MODULE_INTERFACES.items():
        decl = surface.decls.get(interface)
        for member in decl.members if decl is not None and decl.kind == "interface" else ():
            if not member.is_method:
                continue
            found = _input_arg(surface, ledger, {}, member)
            position = found[0] if found is not None else len(member.params)
            union = _object_union_position(surface, list(member.params))
            if union is not None and union < position:
                position = union
            if position < len(member.params):
                out.setdefault(f"{owner}.{member.name}", member.params[position][0])
    for raw, label, _supplied in facade_scope.handle_methods(ledger):
        params = by_raw.get(raw, set())
        if label in out or not params:
            continue
        if len(params) > 1:
            raise GenError(f"{label}: the classes that publish it name its record {sorted(params)}")
        out[label] = params.pop()
    return dict(sorted(out.items()))


def record_params_inc(ledger: Optional[facade_scope.Ledger] = None, dts: Optional[str] = None) -> str:
    """`binding_record_params.inc`: `kRecordParams`, the record parameter of each call, sorted by call."""
    rows = "".join(f'    {{"{call}", "{param}"}},\n' for call, param in record_params(ledger, dts).items())
    return (
        "// Generated by tools/codegen/gen_facade_tables.py from tools/api/names.yaml and src/wasm/formulon.d.ts;\n"
        "// do not edit.\n"
        "\n"
        "/// A call that takes an input record or a patch, and the name of that parameter.\n"
        "struct RecordParam {\n"
        "  const char* call;\n"
        "  const char* param;\n"
        "};\n"
        "\n"
        "// Sorted by call.\n"
        "constexpr RecordParam kRecordParams[] = {\n"
        f"{rows}"
        "};\n"
    )


# The d.ts type a checked raw position must be able to hold, per Node mask.
_MASK_NEEDS = {
    "u32": "number",
    "i32": "number",
    "u64": "number",
    "double": "number",
    "optionalU32": "number",
    "optionalU64": "number",
    "string": "string",
    "object": "object",
}


def _holds(surface: JsSurface, ptype: str, need: str) -> bool:
    alternatives = _alternatives(surface, ptype)
    if need == "number":
        return any(a == "number" or _NUMBER_LITERAL.fullmatch(a) for a in alternatives)
    if need == "string":
        return any(a == "string" or a[:1] in "'\"`" for a in alternatives)
    return any(_is_object_type(a) for a in alternatives)


def _layout_fits(surface: JsSurface, params: Sequence[Tuple[str, str]], supplied: int, spec: Dict[str, int]) -> bool:
    """Whether every checked raw position of a registration is the d.ts parameter at the same place after the
    `supplied` leading arguments: no more checked positions than parameters, each of a type the parameter holds."""
    width = max((mask.bit_length() for kind, mask in spec.items() if kind != "optionalNull"), default=0)
    if width - supplied > len(params):
        return False
    for kind, mask in spec.items():
        need = _MASK_NEEDS.get(kind)
        for index in range(supplied, mask.bit_length()) if need is not None else ():
            if mask >> index & 1 and not _holds(surface, params[index - supplied][1], need):
                return False
    return True


def raw_layouts(ledger: Optional[facade_scope.Ledger] = None) -> Dict[str, List[Tuple[int, Dict[str, int]]]]:
    """`{label: [(supplied, spec), ...]}`: the positional masks of every registration a binding words its refusals
    with `label` (`Sheet.<raw>`, a handle label, `Workbook.<raw>`, `Workbook.<static>`, `Formulon.<fn>`), and the
    count of leading raw arguments the facade supplies."""
    ledger = ledger or facade_scope.load()
    sheet = set(facade_scope.raw_sheet_methods(ledger, "js"))
    handles = {raw: (label, supplied) for raw, label, supplied in facade_scope.handle_methods(ledger)}
    out: Dict[str, List[Tuple[int, Dict[str, int]]]] = {}
    for registration in positional_guards._parse_node_registrations(positional_guards.DEFAULT_NODE_SOURCE):
        raw = registration.js_name
        if raw in handles:
            label, supplied = handles[raw]
        elif raw in sheet:
            label, supplied = f"Sheet.{raw}", 1
        else:
            label, supplied = f"Workbook.{raw}", 0
        out.setdefault(label, []).append((supplied, registration.as_js_spec()))
    for name, spec in positional_guards.STATIC_FUNCTION_GUARDS.items():
        out.setdefault(f"Workbook.{name}", []).append((0, dict(spec)))
    for name, spec in positional_guards.FREE_FUNCTION_GUARDS.items():
        out.setdefault(f"Formulon.{name}", []).append((0, dict(spec)))
    return out


def binding_param_names(
    ledger: Optional[facade_scope.Ledger] = None,
    dts: Optional[str] = None,
    layouts: Optional[Dict[str, List[Tuple[int, Dict[str, int]]]]] = None,
    aligned_only: bool = True,
) -> Dict[str, List[str]]:
    """`{call: [name, ...]}`: the first-signature parameter names of each facade call, plus each handle label a
    binding words a raw method with (`Drawing.setZOrder`) that no facade class publishes under that name, with the
    names every publishing class gives the raw method's parameters. A call whose raw layout does not fit its d.ts
    signature (`_layout_fits`) is left out unless `aligned_only` is false; `layouts` defaults to `raw_layouts()`."""
    ledger = ledger or facade_scope.load()
    text = WASM_DTS.read_text(encoding="utf-8") if dts is None else dts
    surface = JsSurface(text)
    signatures = {key: params for key, params in _first_signatures(ledger, surface, text).items() if params}
    by_raw: Dict[str, set] = {}
    for _key, entry in facade_scope._ops(ledger, ("op",)):
        raw = facade_scope.raw_name(entry, "js")
        for qualified_name in facade_scope._qualified_js(ledger, entry) if raw is not None else ():
            if qualified_name in signatures:
                by_raw.setdefault(raw, set()).add(tuple(tuple(p) for p in signatures[qualified_name]))
    out = dict(signatures)
    for raw, label, _supplied in facade_scope.handle_methods(ledger):
        params = by_raw.get(raw, set())
        if label in out or not params:
            continue
        if len({tuple(name for name, _type in p) for p in params}) > 1:
            raise GenError(f"{label}: the classes that publish it name its parameters {sorted(params)}")
        out[label] = list(params.pop())
    layouts = raw_layouts(ledger) if layouts is None else layouts
    return {
        call: [name for name, _type in params]
        for call, params in sorted(out.items())
        if not aligned_only
        or all(_layout_fits(surface, params, supplied, spec) for supplied, spec in layouts.get(call, ()))
    }


def param_names_inc(
    ledger: Optional[facade_scope.Ledger] = None,
    dts: Optional[str] = None,
    layouts: Optional[Dict[str, List[Tuple[int, Dict[str, int]]]]] = None,
) -> str:
    """`binding_param_names.inc`: `kParamNames`, the positional parameter names of each call, sorted by call."""
    rows = "".join(
        f'    {{"{call}", "{" ".join(names)}"}},\n' for call, names in binding_param_names(ledger, dts, layouts).items()
    )
    declared = "".join(
        f'    {{"{call}", "{" ".join(names)}"}},\n'
        for call, names in binding_param_names(ledger, dts, layouts, aligned_only=False).items()
    )
    return (
        "// Generated by tools/codegen/gen_facade_tables.py from tools/api/names.yaml and src/wasm/formulon.d.ts;\n"
        "// do not edit.\n"
        "\n"
        "/// A call and the d.ts names of its positional parameters, in order, separated by single spaces.\n"
        "struct ParamNames {\n"
        "  const char* call;\n"
        "  const char* names;\n"
        "};\n"
        "\n"
        "// Sorted by call; only calls whose raw positions follow the d.ts signature.\n"
        "constexpr ParamNames kParamNames[] = {\n"
        f"{rows}"
        "};\n"
        "\n"
        "// Every call's d.ts parameter names, sorted by call: the roots a record field path may start with.\n"
        "constexpr ParamNames kDeclaredParamNames[] = {\n"
        f"{declared}"
        "};\n"
    )


def _js_name_lists(name: str, table: Dict[str, List[str]]) -> str:
    return _js_frozen(name, "".join(_js_key_list(_INDENT, _js_key(k), v) for k, v in table.items()))


def param_names_region(ledger: Optional[facade_scope.Ledger] = None, dts: Optional[str] = None) -> str:
    """The JS `PARAM_NAMES` table."""
    return _js_name_lists("PARAM_NAMES", param_names(ledger, dts))


def boolean_args_region(ledger: Optional[facade_scope.Ledger] = None, dts: Optional[str] = None) -> str:
    """The JS `BOOLEAN_ARGS` table."""
    body = "".join(f"{_INDENT}{_js_key(k)}: {_js_value(v)},\n" for k, v in boolean_args(ledger, dts).items())
    return _js_frozen("BOOLEAN_ARGS", body)


def c_field_names_region(ledger: Optional[facade_scope.Ledger] = None) -> str:
    """The JS `C_FIELD_NAMES` table."""
    rows = []
    for c_name, names in c_field_names(ledger).items():
        py = f"'{names['py']}'" if names["py"] is not None else "null"
        rows.append(f"{_INDENT}{_js_bare_key(c_name)}: {{ js: '{names['js']}', py: {py} }},\n")
    return _js_frozen("C_FIELD_NAMES", "".join(rows))


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
            if member.is_method and (key not in longest or len(member.params) > len(longest[key])):
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


_C_INTEGER_TYPES = frozenset(
    {"int8_t", "uint8_t", "int16_t", "uint16_t", "int32_t", "uint32_t", "int64_t", "uint64_t", "size_t"}
)
_COORDINATE_NAMES = frozenset({"row", "col", "first", "last"})


def arg_role(c_name: str) -> str:
    """The role of the integer C parameter `c_name`: which error class a refused value takes."""
    if c_name == "sheet_index":
        return "sheet"
    if c_name == "index" or c_name.endswith("_index"):
        return "position"
    if c_name in _COORDINATE_NAMES or c_name.endswith(("_row", "_col")):
        return "coordinate"
    return "quantity"


def _c_integer_params(header: Optional[str] = None) -> Dict[str, List[str]]:
    """`{c function: [integer by-value parameter, ...]}` of `formulon_c.h`, in declaration order."""
    text = C_HEADER.read_text(encoding="utf-8") if header is None else header
    functions, _typedefs = parse_header(text)
    return {
        fn.name: [p.name for p in fn.params if p.pointer_depth == 0 and p.base_type in _C_INTEGER_TYPES]
        for fn in functions
    }


def arg_roles(
    ledger: Optional[facade_scope.Ledger] = None, dts: Optional[str] = None, header: Optional[str] = None
) -> Dict[str, List[List[object]]]:
    """`ARG_ROLES`: `{Class.method: [[position, role], ...]}` at the positions of the longest overload. A
    parameter counts when the backing C function takes it as an integer by value and the JS overload declares it
    (camelCase, an `_index` suffix optional) as a `number`; the facade supplies the others (the handle, the
    sheet index of `Sheet`). A handwritten method passes its roles itself."""
    ledger = ledger or facade_scope.load()
    surface = JsSurface(WASM_DTS.read_text(encoding="utf-8") if dts is None else dts)
    integers = _c_integer_params(header)
    longest: Dict[str, List[Tuple[str, str]]] = {}
    for interface, owner in _facade_interfaces(ledger).items():
        decl = surface.decls.get(interface)
        for member in decl.members if decl is not None and decl.kind == "interface" else ():
            key = f"{owner}.{member.name}"
            if member.is_method and (key not in longest or len(member.params) > len(longest[key])):
                longest[key] = list(member.params)
    js: Dict[str, List[List[object]]] = {}
    for entry in ledger.entries.values():
        if not isinstance(entry, dict) or entry.get("kind") != "op" or not isinstance(entry.get("c"), str):
            continue
        if entry.get("facade") == facade_scope.HANDWRITTEN_MODE:
            continue
        params = integers.get(entry["c"], [])
        if not params:
            continue
        for key in facade_scope.qualified(entry, "js", ledger.collections):
            declared = longest.get(key)
            if declared is None:
                continue
            names = [name for name, _type in declared]
            slots = []
            for c_name in params:
                spellings = [_camel(c_name)]
                if c_name.endswith("_index"):
                    spellings.append(_camel(c_name[: -len("_index")]))
                camel = next((n for n in spellings if n in names), None)
                if camel is None:
                    continue
                ptype = declared[names.index(camel)][1]
                if "number" not in _strip_null(ptype, True):
                    continue
                slots.append([names.index(camel), arg_role(c_name)])
            if slots:
                js[key] = sorted(slots, key=lambda slot: slot[0])
    return dict(sorted(js.items()))


def signatures_region(ledger: Optional[facade_scope.Ledger] = None, dts: Optional[str] = None) -> str:
    """The JS `METHOD_ARITY`, `TYPED_ARGS`, `DEFAULT_ARGS` and `ARG_ROLES` tables."""
    arity, typed = method_signatures(ledger, dts)
    out = _js_frozen("METHOD_ARITY", "".join(f"{_INDENT}{_js_key(k)}: {n},\n" for k, n in arity.items()))
    body = "".join(f"{_INDENT}{_js_key(k)}: {_js_positions(v)},\n" for k, v in typed.items())
    out += _js_frozen("TYPED_ARGS", body)
    body = "".join(f"{_INDENT}{_js_key(k)}: {_js_positions(v)},\n" for k, v in default_args(ledger, dts).items())
    out += _js_frozen("DEFAULT_ARGS", body)
    body = "".join(f"{_INDENT}{_js_key(k)}: {_js_positions(v)},\n" for k, v in arg_roles(ledger, dts).items())
    return out + _js_frozen("ARG_ROLES", body)


def _array_element(text: str) -> Optional[str]:
    """The element type of the array field type `text` (`T[]`, `readonly T[]`, `ReadonlyArray<T>`), else None."""
    kinds = [k for k in _strip_null(text, True) if k != "undefined"]
    match = _ARRAY_RE.fullmatch(kinds[0].strip()) if len(kinds) == 1 else None
    if match is None:
        return None
    inner = [
        a.strip() for a in _strip_null(next(g for g in match.groups() if g), True) if a not in ("string", "undefined")
    ]
    return inner[0] if len(inner) == 1 else None


def _record_alternative(surface: JsSurface, text: str) -> Optional[str]:
    """The one input record or `Patch<...>` alias `text` names, `| null` / `undefined` aside, or None."""
    kinds = [k.strip() for k in _strip_null(text, True) if k != "undefined"]
    if len(kinds) != 1:
        return None
    name = kinds[0]
    return name if _record_fields(surface, name) or _patch_alias(surface, name) else None


def _patch_alias(surface: JsSurface, name: str) -> bool:
    """Whether `name` is a `type XPatch = Patch<...>` alias."""
    decl = surface.decls.get(name)
    return decl is not None and decl.kind == "type" and name.endswith("Patch") and "Patch<" in decl.text


def _nested_fields(surface: JsSurface, ledger: facade_scope.Ledger, name: str) -> Dict[str, List[str]]:
    """Field -> declared types of the nested record `name`; a `Patch<...>` alias has its keys. Each also takes
    the `read_only_keys` of its ledger record, typed `unknown`, which a record a getter returned carries back."""
    if _patch_alias(surface, name):
        out = {f: [t] for f, t in (_patch_keys(surface.decls, surface.raw, name) or {}).items()}
    else:
        out = {f: [t] for f, t in _record_fields(surface, name).items()}
    for field in facade_scope.read_only_keys(ledger, name)[0]:
        out.setdefault(field, ["unknown"])
    return out


def _nested_paths(
    surface: JsSurface,
    ledger: facade_scope.Ledger,
    fields: Dict[str, List[str]],
    seen: Tuple[str, ...],
    patch_aliases_only: bool = False,
) -> Dict[str, List[str]]:
    """`{path: [key, ...]}` of every record nested under `fields` (field -> declared types), sorted by path. With
    `patch_aliases_only` a top-level field starts a path only when typed by a `Patch<...>` alias."""
    out: Dict[str, List[str]] = {}

    def walk(current: Dict[str, List[str]], prefix: str, seen: Tuple[str, ...]) -> None:
        for field, ftypes in current.items():
            for ftype in ftypes:
                element = _array_element(ftype)
                nested = _record_alternative(surface, element if element is not None else ftype)
                if nested is None or nested in seen:
                    continue
                if patch_aliases_only and not prefix and not _patch_alias(surface, nested):
                    continue
                path = f"{prefix}{field}{'[]' if element is not None else ''}"
                inner = _nested_fields(surface, ledger, nested)
                out[path] = sorted(set(out.get(path, ())) | set(inner))
                walk(inner, path + ".", (*seen, nested))

    walk(fields, "", seen)
    return dict(sorted(out.items()))


def _input_arg(surface: JsSurface, ledger: facade_scope.Ledger, entry: dict, member) -> Optional[Tuple[int, str]]:
    """`(position, record)` of the argument of `member` that is an input record, None when it takes none. A value
    union with several object alternatives (`Date | ErrorValue | PivotCacheValue`) is not a record argument."""
    addressed = {p for field in ("cell_args", "rect_args", "span_args", "ref_args") for p in _flat(entry.get(field))}
    for position, (_name, ptype) in enumerate(member.params):
        if position in addressed:
            continue
        alternatives = [a for a in _strip_null(ptype, True) if a != "undefined"]
        if sum(a not in _PRIMITIVES and a[:1] not in "'\"" for a in alternatives) > 1:
            continue
        for alternative in alternatives:
            decl = surface.decls.get(alternative)
            if (
                decl is not None
                and (decl.kind == "interface" or (decl.kind == "type" and _record_fields(surface, alternative)))
                and alternative not in _NOT_INPUT_RECORDS
                and not alternative.endswith("Patch")
                and not any(m.is_method for m in decl.members)
            ):
                return position, alternative
    return None


def _flat(value: object) -> List[int]:
    if isinstance(value, int) and not isinstance(value, bool):
        return [value]
    if isinstance(value, (list, tuple)):
        return [p for item in value for p in _flat(item)]
    if isinstance(value, dict):
        return [p for item in value.values() for p in _flat(item)]
    return []


def _input_keys_of(surface: JsSurface, ledger: facade_scope.Ledger, record: str) -> List[str]:
    """The keys `record` declares and its `read_only_keys`, sorted."""
    return sorted(set(_record_fields(surface, record)) | set(facade_scope.read_only_keys(ledger, record)[0]))


def _input_nested(surface: JsSurface, ledger: facade_scope.Ledger, record: str) -> Dict[str, List[str]]:
    """`{path: [key, ...]}` of every record nested in `record`, sorted by path."""
    fields = {f: [t] for f, t in _record_fields(surface, record).items()}
    paths = _nested_paths(surface, ledger, fields, (record,))
    return _with_read_only_paths(surface, ledger, record, fields, paths)


def input_keys(
    ledger: Optional[facade_scope.Ledger] = None, dts: Optional[str] = None
) -> Tuple[Dict[str, list], Dict[str, Dict[str, list]]]:
    """`(INPUT_KEYS, INPUT_NESTED_KEYS)` rows of every non-patch op taking a d.ts input record."""
    ledger = ledger or facade_scope.load()
    surface = JsSurface(WASM_DTS.read_text(encoding="utf-8") if dts is None else dts)
    top: Dict[str, list] = {}
    nested: Dict[str, Dict[str, list]] = {}
    for _key, entry in facade_scope._ops(ledger, ("op",)):
        if entry.get("patch") is True:
            continue
        for qualified_name in facade_scope._qualified_js(ledger, entry):
            cls, name = qualified_name.split(".", 1)
            member = facade_scope._patch_member(surface, entry, name, cls)
            if member is None and entry.get("namespace") == "ooxml":
                # The facade installs the `Workbook` / `Sheet` rows on the OOXML namespace objects too.
                member = facade_scope._patch_member(surface, entry, name, f"{cls}Ooxml")
            found = _input_arg(surface, ledger, entry, member) if member is not None else None
            if found is None:
                continue
            position, record = found
            top[qualified_name] = [position, _input_keys_of(surface, ledger, record)]
            paths = _input_nested(surface, ledger, record)
            if paths:
                nested[qualified_name] = paths
    return dict(sorted(top.items())), dict(sorted(nested.items()))


def input_keys_region(ledger: Optional[facade_scope.Ledger] = None, dts: Optional[str] = None) -> str:
    """The JS `INPUT_KEYS`, `INPUT_NESTED_KEYS` and `READ_ONLY_KEYS` tables."""
    top, nested = input_keys(ledger, dts)
    body = "".join(_js_patch_entry(name, position, keys) for name, (position, keys) in top.items())
    out = f"const INPUT_KEYS = Object.freeze({{\n{body}}});\n"
    rows = []
    for name, paths in nested.items():
        lines = "".join(_js_key_list(_INDENT * 2, _js_bare_key(path), keys) for path, keys in paths.items())
        rows.append(f"{_INDENT}{_js_key(name)}: {{\n{lines}{_INDENT}}},\n")
    out += f"const INPUT_NESTED_KEYS = Object.freeze({{\n{''.join(rows)}}});\n"
    return out + _js_name_lists("READ_ONLY_KEYS", read_only_table("js", ledger, dts))


_PRIMITIVES = {"number": "number", "string": "string", "boolean": "boolean", "true": "boolean", "false": "boolean"}


def _typeof_name(surface: JsSurface, text: str) -> str:
    """The `typeof` word checkPatch tests a field of declared type `text` against."""
    names = set()
    for kind in _strip_null(text, True):
        kind = kind.strip()
        if kind == "undefined":
            continue
        decl = surface.decls.get(kind)
        if kind in _PRIMITIVES:
            names.add(_PRIMITIVES[kind])
        elif kind[:1] in "'\"":
            names.add("string")
        elif _array_element(kind) is not None or kind.endswith("[]"):
            names.add("array")
        elif decl is not None and decl.kind == "const" and re.search(r":\s*-?\d", decl.text):
            names.add("number")
        elif decl is not None and decl.kind == "const" and re.search(r":\s*['\"]", decl.text):
            names.add("string")
        else:
            names.add("object")
    return names.pop() if len(names) == 1 else "object"


def patch_nullable(
    ledger: Optional[facade_scope.Ledger] = None, dts: Optional[str] = None
) -> Dict[str, Dict[str, str]]:
    """`{Class.method: {key: typeof}}` of the patch ops whose patch record lists `clearable` fields."""
    ledger = ledger or facade_scope.load()
    surface = JsSurface(WASM_DTS.read_text(encoding="utf-8") if dts is None else dts)
    clearable: Dict[str, List[str]] = {}
    for entry in ledger.entries.values():
        if isinstance(entry, dict) and isinstance(entry.get("js"), str) and entry.get("clearable"):
            clearable[entry["js"]] = [_camel(f) for f in entry["clearable"]]
    out: Dict[str, Dict[str, str]] = {}
    for key, entry in facade_scope._ops(ledger, ("op",)):
        if entry.get("patch") is not True:
            continue
        for qualified_name in facade_scope._qualified_js(ledger, entry):
            cls, name = qualified_name.split(".", 1)
            member = facade_scope._patch_member(surface, entry, name, cls)
            patch_type = facade_scope._patch_type(member.params) if member is not None else None
            # A handle's patch may be an intersection with the common drawing keys: the kind's own record lists them.
            parts = [re.match(r"[\w$]*", p.strip()).group(0) for p in (patch_type or "").split("&")]
            fields = next((clearable[p] for p in parts if clearable.get(p)), None)
            if not fields:
                continue
            declared = facade_scope._patch_keys(surface.decls, surface.raw, patch_type) or {}
            unknown = sorted(set(fields) - set(declared))
            if unknown:
                raise ValueError(f"{key}: clearable fields {patch_type} lacks: {', '.join(unknown)}")
            out[qualified_name] = {f: _typeof_name(surface, declared[f]) for f in sorted(fields)}
    return dict(sorted(out.items()))


def patch_nullable_region(ledger: Optional[facade_scope.Ledger] = None, dts: Optional[str] = None) -> str:
    """The JS `PATCH_NULLABLE` table."""
    rows = ""
    for name, fields in patch_nullable(ledger, dts).items():
        body = ", ".join(f"{_js_bare_key(f)}: '{t}'" for f, t in fields.items())
        line = f"{_INDENT}{_js_key(name)}: {{ {body} }},"
        if len(line) > _LINE_WIDTH:
            inner = "".join(f"{_INDENT * 2}{_js_bare_key(f)}: '{t}',\n" for f, t in fields.items())
            line = f"{_INDENT}{_js_key(name)}: {{\n{inner}{_INDENT}}},"
        rows += line + "\n"
    return f"const PATCH_NULLABLE = Object.freeze({{\n{rows}}});\n"


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
    if only == "record-params":
        return [(RECORD_PARAMS_INC, record_params_inc())]
    if only == "param-names":
        return [(PARAM_NAMES_INC, param_names_inc())]
    result: List[Tuple[Path, Optional[str]]] = []
    common = NPM_COMMON.read_text(encoding="utf-8")
    common = _splice(common, _REGION["js"], js_region(), NPM_COMMON)
    if _RANGE_REGION["js"].search(common):
        common = _splice(common, _RANGE_REGION["js"], range_fields_region("js"), NPM_COMMON)
    if _SIGNATURE_REGION.search(common):
        common = _splice(common, _SIGNATURE_REGION, signatures_region(), NPM_COMMON)
    if _INPUT_REGION.search(common):
        common = _splice(common, _INPUT_REGION, input_keys_region(), NPM_COMMON)
    if _NULLABLE_REGION.search(common):
        common = _splice(common, _NULLABLE_REGION, patch_nullable_region(), NPM_COMMON)
    for pattern, region in (
        (_PARAM_NAMES_REGION, param_names_region),
        (_BOOLEAN_ARGS_REGION, boolean_args_region),
        (_C_FIELD_NAMES_REGION, c_field_names_region),
    ):
        if pattern.search(common):
            common = _splice(common, pattern, region(), NPM_COMMON)
    result.append((NPM_COMMON, common))
    result.append(_region_output(PYTHON_SHEET, _REGION["py"], py_region))
    result.append(_region_output(PYTHON_INPUTS, _RANGE_REGION["py"], lambda: range_fields_region("py")))
    result.append((PYTHON_HANDLE_TABLES, py_handle_tables()))
    result.append((NODE_SCOPE_INC, node_scope_inc()))
    result.append((RECORD_PARAMS_INC, record_params_inc()))
    result.append((PARAM_NAMES_INC, param_names_inc()))
    result.append((ADDRESS_ARGS_JSON, address_args_json()))
    return result


# Outputs whose region `--check` skips while it is empty; `--write` fills it.
_FILLABLE = {
    PYTHON_SHEET: (_REGION["py"],),
    PYTHON_INPUTS: (_RANGE_REGION["py"],),
    NPM_COMMON: (
        _RANGE_REGION["js"],
        _SIGNATURE_REGION,
        _INPUT_REGION,
        _NULLABLE_REGION,
        _PARAM_NAMES_REGION,
        _BOOLEAN_ARGS_REGION,
        _C_FIELD_NAMES_REGION,
    ),
}


def _skippable(path: Path, current: Optional[str], text: str) -> bool:
    """`--check` passes an output whose only differences are fillable regions that are still empty."""
    emptied = text
    for pattern in _FILLABLE.get(path, ()):
        if _empty_region(current, pattern):
            match = pattern.search(emptied)
            if match is not None:
                emptied = emptied[: match.start(2)] + emptied[match.end(2) :]
    return emptied != text and current == emptied


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
