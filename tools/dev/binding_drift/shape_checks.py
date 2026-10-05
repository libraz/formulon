"""Drift checks on the shape of shared declarations: return types and record field sets."""

from __future__ import annotations

import re
from typing import List, Optional, Set

from .c_header import _c_struct_fields
from .surface_files import (
    CAPI_HEADER,
    NODE_DTS,
    PYTHON_WORKBOOK,
    REPO_ROOT,
    WASM_DTS,
    _format_diff,
    _read,
)
from .ts_decl import (
    _find_interface_body,
    _ts_field_set,
    _ts_interface_declaration,
    _ts_return_types,
)

# ---------------------------------------------------------------------------
# Check 2c: shared `.d.ts` shapes -- return types and record field sets.
#
# `dts-wasm` and `dts-node` compare method *names* only. That is enough to
# notice a method that exists on one surface and not the other, and nothing
# else: a shared method whose two declarations disagree about what comes back,
# or a record type that grew a field on one side, passes both of them. The
# README's parity claim is about the shapes, not the names, so this check
# compares the shapes.
#
# Types declared in only one of the two published declaration files. Each has
# to be reachable only from that surface's own method allowlist -- a type that
# becomes one-sided for any other reason is a missing declaration, and listing
# it here is the deliberate act that says otherwise.
# ---------------------------------------------------------------------------

_DTS_SURFACE_ONLY_TYPES = {
    # The embind module factory and its options: there is no equivalent on a
    # native addon, which is `require`d rather than instantiated.
    ("wasm", "FormulonModule"),
    ("wasm", "FormulonModuleOptions"),
    # Reachable only from `getSheetAutoFilterXml` / `createTable` /
    # `updateTable`, all of which are in WASM_ONLY_METHODS.
    ("wasm", "SheetAutoFilterXmlResult"),
    ("wasm", "TableInput"),
    ("wasm", "TableUpdateInput"),
}

# Shared methods whose two declarations are allowed to differ, with the
# reason. Empty by design: a difference that is not worth an entry here is a
# difference that should not exist.
_DTS_RETURN_TYPE_EXEMPT_METHODS: dict[str, str] = {}


def check_dts_shared_shapes() -> List[str]:
    problems: List[str] = []
    wasm_dts = _read(WASM_DTS)
    node_dts = _read(NODE_DTS)
    wasm_label = str(WASM_DTS.relative_to(REPO_ROOT))
    node_label = str(NODE_DTS.relative_to(REPO_ROOT))

    for interface in ("Workbook", "WorkbookCtor"):
        wasm_returns = _ts_return_types(_find_interface_body(wasm_dts, interface, WASM_DTS))
        node_returns = _ts_return_types(_find_interface_body(node_dts, interface, NODE_DTS))
        for name in sorted(set(wasm_returns) & set(node_returns)):
            if wasm_returns[name] == node_returns[name]:
                if name in _DTS_RETURN_TYPE_EXEMPT_METHODS:
                    problems.append(
                        f"dts-shared-shapes: {name} is listed as a deliberate return-type "
                        "difference but both surfaces now declare the same type; drop the entry"
                    )
                continue
            if name in _DTS_RETURN_TYPE_EXEMPT_METHODS:
                continue
            problems.append(
                f"dts-shared-shapes: {interface}.{name} returns a different type on each surface:\n"
                f"  {wasm_label}: {wasm_returns[name]}\n"
                f"  {node_label}: {node_returns[name]}"
            )

    wasm_types = set(re.findall(r"^export interface ([A-Za-z0-9_]+)", wasm_dts, re.MULTILINE))
    node_types = set(re.findall(r"^export interface ([A-Za-z0-9_]+)", node_dts, re.MULTILINE))
    surface_only = {("wasm", name) for name in wasm_types - node_types}
    surface_only |= {("node", name) for name in node_types - wasm_types}
    undeclared = surface_only - _DTS_SURFACE_ONLY_TYPES
    if undeclared:
        problems.append(
            "dts-shared-shapes: type declared on one surface only, without an entry in "
            f"_DTS_SURFACE_ONLY_TYPES: {sorted(undeclared)}"
        )
    stale = _DTS_SURFACE_ONLY_TYPES - surface_only
    if stale:
        problems.append(f"dts-shared-shapes: stale _DTS_SURFACE_ONLY_TYPES entries: {sorted(stale)}")

    for name in sorted(wasm_types & node_types):
        if name in ("Workbook", "WorkbookCtor"):
            continue
        wasm_extends, wasm_body = _ts_interface_declaration(wasm_dts, name, WASM_DTS)
        node_extends, node_body = _ts_interface_declaration(node_dts, name, NODE_DTS)
        if wasm_extends != node_extends:
            problems.append(
                f"dts-shared-shapes: {name} inherits differently on each surface:\n"
                f"  {wasm_label}: {wasm_extends or '(nothing)'}\n"
                f"  {node_label}: {node_extends or '(nothing)'}"
            )
        diff = _format_diff(
            f"{name} in {wasm_label}",
            _ts_field_set(wasm_body) - _ts_field_set(node_body),
            f"{name} in {node_label}",
            _ts_field_set(node_body) - _ts_field_set(wasm_body),
        )
        if diff:
            problems.append(f"dts-shared-shapes: {name} field set mismatch:\n" + "\n".join(diff))

    return problems


# Public type name -> the C struct it projects.
_STYLE_RECORD_STRUCTS = {
    "CellXf": "fm_cell_xf",
    "ColorSpec": "fm_color_spec",
    "FontRecord": "fm_font_record",
    "FillRecord": "fm_fill_record",
    "BorderSide": "fm_border_side",
}

# Deliberate omissions, keyed by (surface, type name). Python models a border
# side as a plain dict rather than a dataclass, so it has no `BorderSide` type
# to compare; the dict shape is covered by the wasm32 struct-layout check.
_STYLE_RECORD_EXEMPT_TYPES = {
    ("python", "BorderSide"),
}

# C presence flags that neither host language mirrors as a field, because both
# already have a way to say "absent": the TS surface omits the optional
# property, Python leaves it `None`. Naming the value field alone would let a
# host drop the distinction entirely, so each entry lists the value field the
# flag governs and both are required to exist together.
_STYLE_RECORD_PRESENCE_FLAGS = {
    "fm_cell_xf": {
        "has_text_rotation": "text_rotation",
        "has_indent": "indent",
        "has_relative_indent": "relative_indent",
        "has_shrink_to_fit": "shrink_to_fit",
        "has_reading_order": "reading_order",
    },
}

_TS_FIELD_RE = re.compile(r"^\s*([A-Za-z_$][A-Za-z0-9_$]*)\??\s*:", re.MULTILINE)
_PY_FIELD_RE = re.compile(r"^    ([a-z_][A-Za-z0-9_]*)\s*:", re.MULTILINE)


def _snake_to_camel(name: str) -> str:
    head, *rest = name.split("_")
    return head + "".join(part.capitalize() for part in rest)


def _python_class_fields(text: str, class_name: str) -> Optional[Set[str]]:
    match = re.search(r"^class %s:\s*$" % re.escape(class_name), text, re.MULTILINE)
    if not match:
        return None
    rest = text[match.end() :]
    end = re.search(r"^(?:@|class |def )", rest, re.MULTILINE)
    body = rest[: end.start()] if end else rest
    return set(_PY_FIELD_RE.findall(body))


def check_style_record_fields() -> List[str]:
    problems: List[str] = []
    header = re.sub(r"/\*.*?\*/", "", _read(CAPI_HEADER), flags=re.S)
    wasm_dts = _read(WASM_DTS)
    node_dts = _read(NODE_DTS)
    python_text = _read(PYTHON_WORKBOOK)

    for type_name, struct_name in sorted(_STYLE_RECORD_STRUCTS.items()):
        c_fields = _c_struct_fields(header, struct_name, CAPI_HEADER)
        # A presence flag is dropped from the comparison, but only once its
        # governed value field is confirmed present on the C side -- otherwise
        # renaming the value field would silently retire the flag with it.
        presence = _STYLE_RECORD_PRESENCE_FLAGS.get(struct_name, {})
        for flag, value_field in sorted(presence.items()):
            if flag in c_fields and value_field not in c_fields:
                problems.append(
                    f"style-record-fields: {struct_name}.{flag} is exempted as a presence flag for "
                    f"{value_field!r}, which no longer exists in {CAPI_HEADER.relative_to(REPO_ROOT)}"
                )
        stale_flags = sorted(flag for flag in presence if flag not in c_fields)
        if stale_flags:
            problems.append(f"style-record-fields: stale presence-flag exemptions for {struct_name}: {stale_flags}")
        c_fields = [name for name in c_fields if name not in presence]
        camel_fields = {_snake_to_camel(name) for name in c_fields}
        for surface, source, dts_text in (
            ("wasm", WASM_DTS, wasm_dts),
            ("node", NODE_DTS, node_dts),
        ):
            if (surface, type_name) in _STYLE_RECORD_EXEMPT_TYPES:
                continue
            ts_fields = set(_TS_FIELD_RE.findall(_find_interface_body(dts_text, type_name, source)))
            diff = _format_diff(
                f"{struct_name} in {CAPI_HEADER.relative_to(REPO_ROOT)}",
                camel_fields - ts_fields,
                f"{type_name} in {source.relative_to(REPO_ROOT)}",
                ts_fields - camel_fields,
            )
            if diff:
                problems.append(f"style-record-fields: {surface} {type_name} mismatch:\n" + "\n".join(diff))

        if ("python", type_name) in _STYLE_RECORD_EXEMPT_TYPES:
            continue
        py_fields = _python_class_fields(python_text, type_name)
        if py_fields is None:
            problems.append(
                f"style-record-fields: dataclass {type_name!r} not found in {PYTHON_WORKBOOK.relative_to(REPO_ROOT)}"
            )
            continue
        diff = _format_diff(
            f"{struct_name} in {CAPI_HEADER.relative_to(REPO_ROOT)}",
            set(c_fields) - py_fields,
            f"{type_name} in {PYTHON_WORKBOOK.relative_to(REPO_ROOT)}",
            py_fields - set(c_fields),
        )
        if diff:
            problems.append(f"style-record-fields: python {type_name} mismatch:\n" + "\n".join(diff))

    stale = {
        (surface, type_name)
        for surface, type_name in _STYLE_RECORD_EXEMPT_TYPES
        if type_name not in _STYLE_RECORD_STRUCTS
    }
    if stale:
        problems.append(f"style-record-fields: stale exemption entries: {sorted(stale)}")
    return problems
