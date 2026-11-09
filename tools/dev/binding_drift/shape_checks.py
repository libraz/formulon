"""Drift checks on the shape of shared declarations: return types and record field sets."""

from __future__ import annotations

import re
from typing import List, Optional, Set, Tuple

from .api_conventions import LEDGER_PATH, handle_class_names, load_ledger
from .c_header import _c_struct_fields
from .surface_files import (
    CAPI_HEADER,
    NODE_DTS,
    PYTHON_SHEET,
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

_DTS_SURFACE_ONLY_TYPES: set = set()

# Shared methods whose two declarations are allowed to differ, with the
# reason. Empty by design: a difference that is not worth an entry here is a
# difference that should not exist.
_DTS_RETURN_TYPE_EXEMPT_METHODS: dict[str, str] = {}

_TS_COMMENT_RE = re.compile(r"/\*.*?\*/|//[^\n]*", re.DOTALL)
_TS_MEMBER_KEY_RE = re.compile(r"(?:readonly\s+)?([A-Za-z_$][A-Za-z0-9_$]*)")


def _ts_member_declarations(body: str) -> dict[str, List[str]]:
    """Maps each top-level member of an interface body to its declarations.

    A declaration is the member's whole text -- parameter names, optional
    markers and types, or a field's type -- with comments dropped and
    whitespace collapsed, so the comparison sees the signature a caller
    compiles against rather than its formatting. Overloads keep one entry
    each, sorted.
    """
    text = _TS_COMMENT_RE.sub("", body)
    out: dict[str, List[str]] = {}
    depth = 0
    current: List[str] = []
    for char in text:
        if char in "({[<":
            depth += 1
        elif char in ")}]>" and not (char == ">" and current and current[-1] == "="):
            depth -= 1
        if char == ";" and depth == 0:
            declaration = " ".join("".join(current).split())
            declaration = re.sub(r",\s*\)", ")", declaration).replace("( ", "(")
            match = _TS_MEMBER_KEY_RE.match(declaration)
            if match:
                out.setdefault(match.group(1), []).append(declaration)
            current = []
        else:
            current.append(char)
    return {name: sorted(declarations) for name, declarations in out.items()}


def _declaration_mismatches(kind: str, owner: str, wasm_body: str, node_body: str) -> List[str]:
    wasm_members = _ts_member_declarations(wasm_body)
    node_members = _ts_member_declarations(node_body)
    wasm_label = str(WASM_DTS.relative_to(REPO_ROOT))
    node_label = str(NODE_DTS.relative_to(REPO_ROOT))
    problems: List[str] = []
    for name in sorted(set(wasm_members) & set(node_members)):
        if wasm_members[name] == node_members[name]:
            continue
        problems.append(
            f"dts-shared-shapes: {owner}.{name} {kind} differently on each surface:\n"
            f"  {wasm_label}: {' | '.join(wasm_members[name])}\n"
            f"  {node_label}: {' | '.join(node_members[name])}"
        )
    return problems


def _handle_classes() -> List[str]:
    """The handle classes the ledger names, `Sheet` included; none when the ledger does not load."""
    ledger = load_ledger(LEDGER_PATH)
    return [] if ledger.problems else handle_class_names(ledger.collections)


def check_dts_shared_shapes() -> List[str]:
    problems: List[str] = []
    wasm_dts = _read(WASM_DTS)
    node_dts = _read(NODE_DTS)
    wasm_label = str(WASM_DTS.relative_to(REPO_ROOT))
    node_label = str(NODE_DTS.relative_to(REPO_ROOT))

    # `Sheet` and each handle class of the ledger join the method interfaces once both files declare them.
    method_interfaces = ["Workbook", "WorkbookCtor"]
    for interface in ["Sheet", *_handle_classes()]:
        pattern = r"^export interface %s\b" % re.escape(interface)
        if interface not in method_interfaces and all(
            re.search(pattern, text, re.MULTILINE) for text in (wasm_dts, node_dts)
        ):
            method_interfaces.append(interface)
    for interface in method_interfaces:
        wasm_body = _find_interface_body(wasm_dts, interface, WASM_DTS)
        node_body = _find_interface_body(node_dts, interface, NODE_DTS)
        # Parameter lists, optional markers and enum-typed parameters, which
        # the return-type comparison below does not see.
        problems.extend(_declaration_mismatches("is declared", interface, wasm_body, node_body))
        wasm_returns = _ts_return_types(wasm_body)
        node_returns = _ts_return_types(node_body)
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
        if name in method_interfaces:
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
            continue
        problems.extend(_declaration_mismatches("is typed", name, wasm_body, node_body))

    return problems


# Public type name -> the C struct it projects.
_STYLE_RECORD_STRUCTS = {
    "CellXf": "fm_cell_xf_t",
    "Color": "fm_color_t",
    "Font": "fm_font_t",
    "Fill": "fm_fill_t",
    "BorderSide": "fm_border_side_t",
}

# Deliberate omissions, keyed by (surface, type name); every surface currently
# models every style record.
_STYLE_RECORD_EXEMPT_TYPES: Set[Tuple[str, str]] = set()

# C presence flags that neither host language mirrors as a field, because both
# say "absent" with a nullable value field: `null` in TS, `None` in Python.
# Naming the value field alone would let a host drop the distinction entirely,
# so each entry lists the value field (or the group of fields) the flag
# governs, and all of them are required to exist together with the flag.
_STYLE_RECORD_PRESENCE_FLAGS = {
    "fm_cell_xf_t": {
        "has_alignment": (
            "horizontal",
            "vertical",
            "wrap_text",
            "justify_last_line",
            "text_rotation",
            "indent",
        ),
        "has_horizontal": "horizontal",
        "has_vertical": "vertical",
        "has_wrap_text": "wrap_text",
        "has_justify_last_line": "justify_last_line",
        "has_protection": ("locked", "hidden"),
        "has_text_rotation": "text_rotation",
        "has_indent": "indent",
        "has_relative_indent": "relative_indent",
        "has_shrink_to_fit": "shrink_to_fit",
        "has_reading_order": "reading_order",
    },
    "fm_font_t": {
        "has_bold": "bold",
        "has_italic": "italic",
        "has_strike": "strike",
        "has_family": "family",
        "has_charset": "charset",
        "has_size_pt": "size_pt",
        "has_outline": "outline",
        "has_shadow": "shadow",
        "has_condense": "condense",
        "has_extend": "extend",
    },
}

# C fields the host languages carry as one nested value instead of flat
# members: a gradient fill's parameters and its stop array travel as a single
# `gradient` object. Each entry names the binding field and the C fields it
# replaces, all of which must still exist in the struct.
_STYLE_RECORD_COMPOSITES = {
    "fm_fill_t": {
        "gradient": (
            "has_gradient",
            "gradient_type",
            "gradient_angle_deg",
            "gradient_left",
            "gradient_right",
            "gradient_top",
            "gradient_bottom",
            "gradient_stops",
            "gradient_stop_count",
        ),
    },
}

_TS_FIELD_RE = re.compile(r"^\s*(?:readonly\s+)?([A-Za-z_$][A-Za-z0-9_$]*)\??\s*:", re.MULTILINE)
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
    # Records stay in _workbook.py or move next to `Sheet` in _sheet.py; either module counts.
    python_text = "\n".join(_read(path) for path in (PYTHON_WORKBOOK, PYTHON_SHEET) if path.is_file())

    for type_name, struct_name in sorted(_STYLE_RECORD_STRUCTS.items()):
        c_fields = _c_struct_fields(header, struct_name, CAPI_HEADER)
        # A presence flag is dropped from the comparison, but only once its
        # governed value field is confirmed present on the C side -- otherwise
        # renaming the value field would silently retire the flag with it.
        presence = _STYLE_RECORD_PRESENCE_FLAGS.get(struct_name, {})
        for flag, governed in sorted(presence.items()):
            for value_field in (governed,) if isinstance(governed, str) else governed:
                if flag in c_fields and value_field not in c_fields:
                    problems.append(
                        f"style-record-fields: {struct_name}.{flag} is exempted as a presence flag for "
                        f"{value_field!r}, which no longer exists in {CAPI_HEADER.relative_to(REPO_ROOT)}"
                    )
        stale_flags = sorted(flag for flag in presence if flag not in c_fields)
        if stale_flags:
            problems.append(f"style-record-fields: stale presence-flag exemptions for {struct_name}: {stale_flags}")
        c_fields = [name for name in c_fields if name not in presence]
        composite_fields: Set[str] = set()
        for binding_field, members in sorted(_STYLE_RECORD_COMPOSITES.get(struct_name, {}).items()):
            missing = [member for member in members if member not in c_fields]
            if missing:
                problems.append(
                    f"style-record-fields: {struct_name} composite {binding_field!r} names fields "
                    f"missing from {CAPI_HEADER.relative_to(REPO_ROOT)}: {missing}"
                )
            c_fields = [name for name in c_fields if name not in members]
            composite_fields.add(binding_field)
        camel_fields = {_snake_to_camel(name) for name in c_fields} | composite_fields
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
            (set(c_fields) | composite_fields) - py_fields,
            f"{type_name} in {PYTHON_WORKBOOK.relative_to(REPO_ROOT)}",
            py_fields - (set(c_fields) | composite_fields),
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
