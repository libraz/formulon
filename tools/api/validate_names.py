"""Validates the public-name ledger `tools/api/names.yaml`.

Exit 0 iff the ledger is a well-formed `schema: 1` document; no name is
renamed into a name another entry is renamed away from (no rename chain)
and no name is claimed twice; every new C function name passes R1 / R2;
every binding name of a C-backed op equals its R5 derivation; every record
and enum name equals `derive_type` of its C typedef; and every op whose C
signature carries an enum or a predicate verb states it in `returns` /
`params`. An `exception:` exempts an entry from the derivation checks, and so
does `review: true`, which marks a proposal still to be decided. Every op, list,
factory and free function on a binding surface carries a `scope`; index-keyed
collections carry `access` (and `plural`, or `exempt` with a reason);
`enum_families` and `enum_distinct` name existing enum entries. An op's
`cell_args` / `rect_args` equal the `(row, col)` pairs and `CellRect` arguments
of its first signature in src/wasm/formulon.d.ts, and an op's or list's
`ref_args` its `*Ref` argument (`ref_exempt: {params: [name, ...], reason}` keeps
the named arguments off `*Ref`); an op whose first signature takes a row or
column span (`(first, last)` or `(row | col, count)`) carries `span_args`. The
positions of an op with `handle:` are those of the handle method, checked
against `interface <handle.class>`. A lone `row` / `col` argument that a
sibling overload takes as `a1: string` is the span form `single` (one row or
column). `patch: true` names its getter's record in
`patch_record`, never `none`. A public name is unique per scope (or handle
class), a raw name per scope.

`reads_values: true` marks a C-backed op that computes pending cells before it
reads; `resolved: true` a read whose XF-family indices are resolved output;
`color: resolved` a record, field or op whose colours are rendered results;
`namespace: ooxml` a `Workbook` / `Sheet` op published under `.ooxml`;
`a1_exempt` the `*Row` / `*Col` fields of an input record that stay numeric. A
`params` slot is a type, or `{type, default}` for a trailing argument the
binding may omit; a default given there names a parameter of the C signature.

`add_returns: handle | key | index` on an `add` op states what it returns: the
element's handle, its stable key (id, GUID, name) or its position (checked
against the collection by the R37 rule). `clearable: [field, ...]` on a record
or typedef lists the snake_case fields a patch may clear (`null` / `None` /
`*_op = CLEAR`); each names a field of the record's JS interface, and
`tools/api/gen_patch_clear_cases.py` derives the shared clear-case table from
it. `read_only_keys: [field, ...]` on an input, patch or both record lists the
snake_case fields (a dotted path for a nested one) of the read record that
passes back into it which the record itself does not declare; each is a field
of a read twin (the `patch_record` of a patch op taking it, `X` of `XInput` /
`XPatch`, `XProperties`) and none of the record. The top-level
`unitless: [name, ...]` lists the field and parameter names (any casing;
compared in snake_case) that carry a dimension without a unit
suffix (R42). A `kind: constant` entry with `enum:` and a `js` / `py` name is
an enum member (`Enum.Member`), renamed through `was: {js, py}`; its names are
unique within its enum. After every rename, no two entries take the same new
name in one namespace of one surface.

`--ledger-only` runs the schema, chain, index and rule checks without reading
src/wasm/formulon.d.ts or the C header.

Reading the tree, every current record, enum or typedef name must be exported on the
surfaces its entry lists unless a `drop:` removes it, and no other entry may drop it
(R0.presence).

`--coverage` additionally requires every symbol the current tree exports
to be named by some entry's current or `was` name. Entries carrying `flag:` or `review: true` are listed in
either mode; they do not fail validation.
"""

from __future__ import annotations

import argparse
import re
import sys
from collections import defaultdict
from pathlib import Path
from typing import Dict, List, Optional, Sequence

sys.path.insert(0, str(Path(__file__).resolve().parent))

import extract_names  # noqa: E402
from binding_drift.api_conventions import (  # noqa: E402
    HANDLE_IDENTITY_PROPERTIES,
    LEDGER_PATH,
    REPO_ROOT,
    SCOPES,
    SURFACES,
    Ledger,
    check_c_name,
    derive_js,
    derive_py,
    derive_record_type,
    derive_type,
    entry_handle,
    handle_class_names,
    handle_classes,
    handle_collections,
    handle_name,
    split_c_name,
)
from binding_drift.api_conventions_arguments import _patch_keys  # noqa: E402
from binding_drift.api_conventions_bindings import JsSurface, _js_record_keys  # noqa: E402
from binding_drift.api_conventions_shapes import _strip_null  # noqa: E402

KINDS = frozenset({"op", "list", "free", "factory", "record", "field", "typedef", "enum", "constant"})
FIELDS = frozenset(
    {
        "kind", "scope", "noun", "qualifies", "record", "role", "from", "c", "js", "py", "enum", "returns",
        "params", "surfaces", "exception", "replaced_by", "closes", "was", "drop", "flag", "sheet_arg", "total",
        "missing", "facade", "review", "note", "cell_args", "rect_args", "ref_args", "ref_exempt", "patch",
        "patch_record", "patch_reason", "patch_nested", "span_args", "internal", "handle", "reads_values", "resolved",
        "color", "namespace", "a1_exempt", "add_returns", "clearable", "read_only_keys", "was_semantics",
    }
)  # fmt: skip
HANDLE_FIELDS = frozenset({"class", "js", "py", "supplies", "property"})
COLLECTION_FIELDS = frozenset(
    {"key", "by", "access", "plural", "ops", "no_getter", "exempt", "flag", "handle_class", "handle_base"}
)
SPAN_AXES = frozenset({"row", "col"})
SPAN_FORMS = frozenset({"span", "start_count", "single"})
# Reference types of the JS surface, by the `ref_args` kind that names them.
JS_REF_KINDS = {
    "TableRef": "table",
    "PivotRef": "pivot",
    "CfRuleRef": "cfRule",
    "ProtectedRangeRef": "protectedRange",
    "ValidationRef": "validation",
    "SparklineGroupRef": "sparklineGroup",
    "PivotCacheRef": "pivotCache",
}
ROLES = frozenset({"result", "input", "patch", "both"})
SCOPED_KINDS = frozenset({"op", "list", "factory", "free"})
BINDING_SURFACES = frozenset({"wasm", "native", "py"})
ACCESS_MODES = frozenset({"handles", "list", "indexed", "property", "exempt"})
MISSING_MODES = (None, "throw")  # YAML reads `missing: null` as None
FACADE_MODES = frozenset({"handwritten"})
NAME_FIELDS = ("c", "js", "py")
FUNCTION_KINDS = frozenset({"op", "factory"})
TYPE_KINDS = frozenset({"record", "enum", "typedef"})
_RULE_ID_RE = re.compile(r"R\d+(\.[a-z0-9_]+)*")
_SLOT_RE = re.compile(r"bool|string|number|int|double|enum:[a-z0-9_]+|union:[A-Za-z0-9_]+")
NAMESPACES = frozenset({"ooxml"})
COLOR_MODES = frozenset({"resolved"})
# Defaults the binding computes rather than spells: `today` is the pinned clock, else the current date.
DYNAMIC_DEFAULTS = frozenset({"today"})
A1_EXEMPT_ROLES = frozenset({"input", "patch", "both"})
INPUT_ROLES = A1_EXEMPT_ROLES
_SNAKE_PATH_RE = re.compile(r"[a-z][a-z0-9]*(?:_[a-z0-9]+)*(?:\.[a-z][a-z0-9]*(?:_[a-z0-9]+)*)*")
ADD_RETURNS = frozenset({"handle", "key", "index"})
WASM_DTS = REPO_ROOT / "src" / "wasm" / "formulon.d.ts"
PATCH_RECORD_NONE = "none"
_JS_CELL_PAIR_RE = re.compile(r"(\w*?)(?:row|Row)")


_LEDGER_CACHE: dict = {}


def _load(path: Path):
    """Parses the ledger once per (path, mtime, size); callers get a private copy."""
    import copy

    import yaml  # type: ignore

    st = path.stat()
    key = (str(path.resolve()), st.st_mtime_ns, st.st_size)
    if key not in _LEDGER_CACHE:
        _LEDGER_CACHE[key] = yaml.safe_load(path.read_text(encoding="utf-8"))
    return copy.deepcopy(_LEDGER_CACHE[key])


def _surface_of(name_field: str) -> Sequence[str]:
    return {"c": ("c",), "js": ("wasm", "native"), "py": ("py",)}[name_field]


def check_schema(data) -> List[str]:
    if not isinstance(data, dict) or data.get("schema") != 1:
        return ["ledger is not a `schema: 1` mapping"]
    problems: List[str] = []
    entries = data.get("entries")
    collections = data.get("collections")
    if not isinstance(entries, dict) or not isinstance(collections, dict):
        return ["ledger needs `collections:` and `entries:` mappings"]
    for key, entry in entries.items():
        where = f"entries.{key}"
        if not isinstance(entry, dict):
            problems.append(f"{where}: not a mapping")
            continue
        unknown = set(entry) - FIELDS
        if unknown:
            problems.append(f"{where}: unknown fields {sorted(unknown)}")
        kind = entry.get("kind")
        if kind not in KINDS:
            problems.append(f"{where}: kind {kind!r} is not one of {sorted(KINDS)}")
        for name_field in NAME_FIELDS:
            if name_field in entry and not isinstance(entry[name_field], str):
                problems.append(f"{where}: {name_field} must be a string")
        surfaces = entry.get("surfaces")
        if not isinstance(surfaces, list) or not surfaces or not set(surfaces) <= set(SURFACES):
            problems.append(f"{where}: surfaces must be a non-empty subset of {list(SURFACES)}")
            surfaces = []
        for name_field in NAME_FIELDS:
            if name_field in entry and not set(_surface_of(name_field)) & set(surfaces):
                problems.append(f"{where}: has a {name_field} name but no {name_field} surface")
        for mapping in ("was", "drop"):
            value = entry.get(mapping)
            if value is None:
                continue
            allowed = set(NAME_FIELDS) | ({"closes", "note"} if mapping == "drop" else set())
            if (
                not isinstance(value, dict)
                or not set(value) <= allowed
                or not all(isinstance(v, str) for v in value.values())
            ):
                problems.append(f"{where}: {mapping} must map {sorted(allowed)} to strings")
        if "drop" in entry and not _RULE_ID_RE.fullmatch(str(entry["drop"].get("closes", ""))):
            problems.append(f"{where}: drop needs a rule-id in `closes`")
        if "closes" in entry and not _RULE_ID_RE.fullmatch(str(entry["closes"])):
            problems.append(f"{where}: closes {entry['closes']!r} is not a rule-id")
        if "replaced_by" in entry:
            if entry["replaced_by"] not in entries:
                problems.append(f"{where}: replaced_by {entry['replaced_by']!r} is not an entry")
            if "closes" not in entry and "closes" not in (entry.get("drop") or {}):
                problems.append(f"{where}: replaced_by needs `closes`")
        for ref_field in ("record", "enum"):
            if ref_field in entry and entry[ref_field] not in entries:
                problems.append(f"{where}: {ref_field} {entry[ref_field]!r} is not an entry")
        if "from" in entry and (not isinstance(entry["from"], list) or not all(k in entries for k in entry["from"])):
            problems.append(f"{where}: from must list entry keys")
        if "role" in entry and entry["role"] not in ROLES:
            problems.append(f"{where}: role {entry['role']!r} is not one of {sorted(ROLES)}")
        if "scope" in entry and entry["scope"] not in (*SCOPES, "process"):
            problems.append(f"{where}: scope {entry['scope']!r} is unknown")
        if kind in SCOPED_KINDS and set(surfaces) & BINDING_SURFACES and "scope" not in entry:
            problems.append(f"{where}: a {kind} on a binding surface needs `scope`")
        if "missing" in entry and entry["missing"] not in MISSING_MODES:
            problems.append(f"{where}: missing must be `null` or `throw`")
        if "facade" in entry and entry["facade"] not in FACADE_MODES:
            problems.append(f"{where}: facade must be one of {sorted(FACADE_MODES)}")
        if "review" in entry and entry["review"] is not True:
            problems.append(f"{where}: review is `true` or absent")
        if "note" in entry and (not isinstance(entry["note"], str) or not entry["note"].strip()):
            problems.append(f"{where}: note must be a non-empty string")
        if "was_semantics" in entry and (entry["was_semantics"] != "inverted" or not entry.get("was")):
            problems.append(f"{where}: was_semantics is `inverted`, on an entry with `was`")
        slots = {}
        for param, spec in (entry.get("params") or {}).items():
            if isinstance(spec, dict):
                problems += _param_default_schema(f"{where}: param {param}", spec)
                if "type" not in spec:
                    continue
            slots[param] = slot_type(spec)
        if "returns" in entry:
            slots["<return>"] = entry["returns"]
        for slot, value in slots.items():
            if not isinstance(value, str) or not _SLOT_RE.fullmatch(value):
                problems.append(
                    f"{where}: {slot} type {value!r} is not bool, string, number, int, double, enum:<key> or union:<Type>"
                )
            elif value.startswith("enum:") and f"enum.{value[5:]}" not in entries:
                problems.append(f"{where}: {slot} names enum {value[5:]!r}, which has no `enum.{value[5:]}` entry")
        for text_field in ("flag", "exception", "sheet_arg", "total"):
            if text_field in entry and (not isinstance(entry[text_field], str) or not entry[text_field].strip()):
                problems.append(f"{where}: {text_field} must be a non-empty string")
        if "sheet_arg" in entry and entry.get("scope") != "workbook":
            problems.append(f"{where}: `sheet_arg` excuses a sheet index on a workbook-scope entry only")
        problems += _address_and_patch_schema(where, entry)
        problems += _handle_schema(where, entry, data)
        problems += _adoption_schema(where, entry)
        if not any(f in entry for f in (*NAME_FIELDS, "was", "replaced_by", "handle")):
            problems.append(f"{where}: names nothing")
    for name, collection in collections.items():
        where = f"collections.{name}"
        if not isinstance(collection, dict) or collection.get("key") not in ("index", "natural"):
            problems.append(f"{where}: key must be index or natural")
            continue
        unknown = set(collection) - COLLECTION_FIELDS
        if unknown:
            problems.append(f"{where}: unknown fields {sorted(unknown)}")
        if "handle_class" in collection and (
            collection.get("access") != "handles" or not isinstance(collection["handle_class"], str)
        ):
            problems.append(f"{where}: `handle_class` names the class of an `access: handles` collection")
        if "handle_base" in collection and (
            "handle_class" not in collection or collection["handle_base"] not in handle_class_names(collections)
        ):
            problems.append(f"{where}: `handle_base` names another handle class of a handle collection")
        if not all(k in entries for k in collection.get("ops", [])):
            problems.append(f"{where}: ops must list entry keys")
        if "no_getter" in collection and (
            collection["key"] != "natural"
            or not isinstance(collection["no_getter"], str)
            or not collection["no_getter"]
        ):
            problems.append(f"{where}: `no_getter` is a reason string on a natural-key collection")
        access = collection.get("access")
        if collection["key"] == "index" and access not in ACCESS_MODES:
            problems.append(f"{where}: an index-keyed collection needs `access` in {sorted(ACCESS_MODES)}")
        elif collection["key"] == "natural" and access not in (None, "handles"):
            problems.append(f"{where}: a natural-key collection takes `access: handles` or none")
        if access is not None and not isinstance(collection.get("plural"), str):
            problems.append(f"{where}: `access` needs `plural`")
        if (access == "exempt") != isinstance(collection.get("exempt"), str):
            problems.append(f"{where}: `exempt` is the reason string of an `access: exempt` collection")
    problems += check_enum_groups(data, entries)
    problems += check_unitless(data)
    return problems


def _names_list(value) -> bool:
    """A non-empty list of distinct non-empty strings."""
    return (
        isinstance(value, list)
        and bool(value)
        and all(isinstance(v, str) and v for v in value)
        and len(set(value)) == len(value)
    )


def check_unitless(data: dict) -> List[str]:
    """The top-level `unitless:` lists distinct names, distinct in snake_case too."""
    if "unitless" not in data:
        return []
    names = data["unitless"]
    if not _names_list(names):
        return ["unitless: lists distinct non-empty names"]
    snake = [_snake_case(n) for n in names]
    doubled = sorted({n for n in snake if snake.count(n) > 1})
    return [f"unitless: {n} is listed in two casings" for n in doubled]


def _snake_case(name: str) -> str:
    return re.sub(r"(?<=[a-z0-9])(?=[A-Z])", "_", name).lower()


def _index(value) -> bool:
    return isinstance(value, int) and not isinstance(value, bool) and value >= 0


def _address_and_patch_schema(where: str, entry: dict) -> List[str]:
    """`cell_args` lists `[i, j]` pairs, `rect_args` positions and `ref_args` one `[i, kind]`; `patch: true`
    carries `patch_record`, and `patch_record: none` a `patch_reason`."""
    problems: List[str] = []
    if "ref_args" in entry:
        ref = entry["ref_args"]
        if entry.get("kind") not in ("op", "list"):
            problems.append(f"{where}: ref_args belong to an op or a list")
        if not (isinstance(ref, list) and len(ref) == 2 and _index(ref[0]) and ref[1] in JS_REF_KINDS.values()):
            problems.append(f"{where}: ref_args is `[position, kind]`, kind one of {sorted(JS_REF_KINDS.values())}")
    if "ref_exempt" in entry:
        exempt = entry["ref_exempt"]
        if entry.get("kind") != "op":
            problems.append(f"{where}: ref_exempt belongs to an op")
        if (
            not isinstance(exempt, dict)
            or set(exempt) != {"params", "reason"}
            or not _names_list(exempt["params"])
            or not isinstance(exempt["reason"], str)
            or not exempt["reason"].strip()
        ):
            problems.append(f"{where}: ref_exempt is `{{params: [name, ...], reason}}`")
    fields = [f for f in ("cell_args", "rect_args", "patch", "patch_record", "patch_reason") if f in entry]
    if fields and entry.get("kind") != "op":
        problems.append(f"{where}: {', '.join(fields)} belong to an op")
    if "cell_args" in entry:
        pairs = entry["cell_args"]
        if (
            not isinstance(pairs, list)
            or not pairs
            or not all(isinstance(p, list) and len(p) == 2 and all(_index(i) for i in p) for p in pairs)
        ):
            problems.append(f"{where}: cell_args lists `[row_position, col_position]` pairs")
    if "rect_args" in entry:
        positions = entry["rect_args"]
        if not isinstance(positions, list) or not positions or not all(_index(i) for i in positions):
            problems.append(f"{where}: rect_args lists argument positions")
    if "patch" in entry and entry["patch"] is not True:
        problems.append(f"{where}: patch is `true` or absent")
    record = entry.get("patch_record")
    if entry.get("patch") is True and (not isinstance(record, str) or not record):
        problems.append(f"{where}: `patch: true` names its getter's record in `patch_record`")
    if ("patch_record" in entry or "patch_reason" in entry) and entry.get("patch") is not True:
        problems.append(f"{where}: patch_record / patch_reason belong to a `patch: true` op")
    reason = entry.get("patch_reason")
    if record == PATCH_RECORD_NONE and (not isinstance(reason, str) or not reason.strip()):
        problems.append(f"{where}: `patch_record: none` needs a `patch_reason`")
    if "patch_reason" in entry and record != PATCH_RECORD_NONE:
        problems.append(f"{where}: patch_reason explains a `patch_record: none` only")
    nested = entry.get("patch_nested")
    if nested is not None and (
        entry.get("patch") is not True
        or not isinstance(nested, dict)
        or not nested
        or not all(isinstance(k, str) and isinstance(v, str) and v for k, v in nested.items())
    ):
        problems.append(f"{where}: patch_nested maps a path in the patch to its record, on a `patch: true` op")
    elif isinstance(nested, dict):
        for path in nested:
            parent = path.rpartition(".")[0]
            if parent and parent not in nested:
                problems.append(f"{where}: patch_nested path {path} has no entry for {parent}")
    if "span_args" in entry:
        spans = entry["span_args"]
        if entry.get("kind") != "op":
            problems.append(f"{where}: span_args belong to an op")
        if not (
            isinstance(spans, list)
            and spans
            and all(
                isinstance(s, list) and len(s) == 3 and _index(s[0]) and s[1] in SPAN_AXES and s[2] in SPAN_FORMS
                for s in spans
            )
        ):
            problems.append(f"{where}: span_args lists `[position, row|col, span|start_count|single]` triples")
    return problems


def slot_type(spec) -> Optional[str]:
    """The type of a `params` slot: the slot itself, or the `type` of a `{type, default}` mapping."""
    return spec.get("type") if isinstance(spec, dict) else spec


def _param_default_schema(where: str, spec: dict) -> List[str]:
    """A mapping slot is `{type?, default}`; the default is a literal or one of `DYNAMIC_DEFAULTS`."""
    problems: List[str] = []
    if not set(spec) <= {"type", "default"} or "default" not in spec:
        problems.append(f"{where}: a mapping slot is `{{type, default}}` with `default` given")
    default = spec.get("default")
    if isinstance(default, str) and default not in DYNAMIC_DEFAULTS:
        problems.append(f"{where}: default {default!r} is a literal or one of {sorted(DYNAMIC_DEFAULTS)}")
    elif default is not None and not isinstance(default, (bool, int, float, str)):
        problems.append(f"{where}: default must be a scalar")
    return problems


def _adoption_schema(where: str, entry: dict) -> List[str]:
    """`reads_values`, `resolved`, `color`, `namespace`, `a1_exempt`, `add_returns` and `clearable` sit on the
    kinds they describe."""
    problems: List[str] = []
    kind = entry.get("kind")
    if "add_returns" in entry and (entry["add_returns"] not in ADD_RETURNS or kind != "op"):
        problems.append(f"{where}: add_returns is one of {sorted(ADD_RETURNS)} on an op")
    if "clearable" in entry:
        if kind not in ("record", "typedef"):
            problems.append(f"{where}: clearable belongs to a record or a typedef")
        if entry["clearable"] != [] and not _names_list(entry["clearable"]):
            problems.append(f"{where}: clearable lists distinct field names, or is `[]` for a patch with none")
        elif any(_snake_case(f) != f for f in entry["clearable"]):
            problems.append(f"{where}: clearable names fields in snake_case")
    if "read_only_keys" in entry:
        keys = entry["read_only_keys"]
        if kind != "record" or entry.get("role") not in INPUT_ROLES:
            problems.append(f"{where}: read_only_keys belongs to an input, patch or both record")
        if not _names_list(keys):
            problems.append(f"{where}: read_only_keys lists distinct field names")
        elif any(not _SNAKE_PATH_RE.fullmatch(k) for k in keys):
            problems.append(f"{where}: read_only_keys names fields in snake_case, nested ones as a dotted path")
    if "reads_values" in entry and (entry["reads_values"] is not True or kind != "op" or "c" not in entry):
        problems.append(f"{where}: reads_values is `true` on an op with a C name")
    if "resolved" in entry and (entry["resolved"] is not True or kind != "op"):
        problems.append(f"{where}: resolved is `true` on an op")
    if "color" in entry and (entry["color"] not in COLOR_MODES or kind not in ("record", "field", "op")):
        problems.append(f"{where}: color is one of {sorted(COLOR_MODES)} on a record, field or op")
    if "namespace" in entry:
        if entry["namespace"] not in NAMESPACES:
            problems.append(f"{where}: namespace is one of {sorted(NAMESPACES)}")
        if kind not in ("op", "list") or entry.get("scope") not in SCOPES or "handle" in entry:
            problems.append(f"{where}: a namespace holds `Workbook` / `Sheet` ops and lists, not handle methods")
    if "a1_exempt" in entry:
        exempt = entry["a1_exempt"]
        if kind not in ("record", "typedef") or (kind == "record" and entry.get("role") not in A1_EXEMPT_ROLES):
            problems.append(f"{where}: a1_exempt belongs to an input record or an options type")
        if (
            not isinstance(exempt, dict)
            or set(exempt) != {"fields", "reason"}
            or not isinstance(exempt["fields"], list)
            or not exempt["fields"]
            or not all(isinstance(f, str) and f for f in exempt["fields"])
            or not isinstance(exempt["reason"], str)
            or not exempt["reason"].strip()
        ):
            problems.append(f"{where}: a1_exempt is `{{fields: [name, ...], reason}}`")
    return problems


def _handle_schema(where: str, entry: dict, data: dict) -> List[str]:
    """`internal: true` and `handle:` on an op or list; a handle names known classes and its public names."""
    problems: List[str] = []
    kind = entry.get("kind")
    if "internal" in entry:
        if entry["internal"] is not True:
            problems.append(f"{where}: internal is `true` or absent")
        if kind not in ("op", "list"):
            problems.append(f"{where}: internal belongs to an op or a list")
        if "handle" in entry:
            problems.append(f"{where}: an internal op publishes no handle method")
    if "replaced_by" in entry and kind in ("op", "list"):
        target = (data.get("entries") or {}).get(entry["replaced_by"])
        if isinstance(target, dict) and target.get("kind") not in ("op", "list"):
            problems.append(f"{where}: an op or list is replaced by an op or list")
    if "handle" not in entry:
        return problems
    handle = entry_handle(entry)
    if kind not in ("op", "list") or handle is None:
        return problems + [f"{where}: handle is a mapping on an op or a list"]
    unknown = set(handle) - HANDLE_FIELDS
    if unknown:
        problems.append(f"{where}: handle has unknown fields {sorted(unknown)}")
    known = handle_class_names(data.get("collections") or {})
    classes = handle_classes(entry)
    raw_class = handle.get("class")
    if not classes or (isinstance(raw_class, list) and len(classes) != len(raw_class)):
        problems.append(f"{where}: handle.class is a class name or a list of them")
    for cls in classes:
        if cls not in known:
            problems.append(f"{where}: handle.class {cls} is not a handle class of the ledger ({', '.join(known)})")
    surfaces = set(entry.get("surfaces") or [])
    collections = data.get("collections") or {}
    for name_field in ("js", "py"):
        if name_field in handle and not isinstance(handle[name_field], str):
            problems.append(f"{where}: handle.{name_field} must be a string")
        elif set(_surface_of(name_field)) & surfaces and not handle_name(entry, name_field, collections):
            problems.append(f"{where}: handle derives no {name_field} name; give `handle.{name_field}`")
    supplies = handle.get("supplies", [])
    if not isinstance(supplies, list) or not all(isinstance(s, str) and s for s in supplies):
        problems.append(f"{where}: handle.supplies lists the raw arguments the handle supplies")
    if "property" in handle and handle["property"] is not True:
        problems.append(f"{where}: handle.property is `true` or absent")
    return problems


def check_enum_groups(data: dict, entries: Dict[str, dict]) -> List[str]:
    """`enum_families` lists enum entries; each `enum_distinct` item is a pair with a reason or a review note."""
    problems: List[str] = []

    def enum_keys(where: str, keys) -> None:
        if not isinstance(keys, list) or len(keys) < 2:
            problems.append(f"{where}: lists two or more enum entry keys")
            return
        for key in keys:
            if (entries.get(key) or {}).get("kind") != "enum":
                problems.append(f"{where}: {key!r} is not an enum entry")

    families = data.get("enum_families") or {}
    if not isinstance(families, dict):
        return ["enum_families: must map a family name to enum entry keys"]
    for name, keys in families.items():
        enum_keys(f"enum_families.{name}", keys)
    distinct = data.get("enum_distinct") or []
    if not isinstance(distinct, list):
        return problems + ["enum_distinct: must be a list"]
    for i, item in enumerate(distinct):
        where = f"enum_distinct[{i}]"
        if not isinstance(item, dict) or not set(item) <= {"enums", "reason", "review", "note"}:
            problems.append(f"{where}: takes `enums`, `reason`, `review`, `note`")
            continue
        enum_keys(where, item.get("enums"))
        if isinstance(item.get("enums"), list) and len(item["enums"]) != 2:
            problems.append(f"{where}: names exactly two enums")
        if item.get("review") is True:
            if not isinstance(item.get("note"), str) or not item["note"].strip():
                problems.append(f"{where}: a review item carries a `note`")
        elif not isinstance(item.get("reason"), str) or not item["reason"].strip():
            problems.append(f"{where}: needs a `reason`")
    return problems


_SCOPED_NAME_KINDS = frozenset({"op", "list", "free", "factory"})


def _namespace(entry: dict, name_field: str) -> str:
    """The namespace of the entry's raw (or former) `name_field` name.

    C names share one namespace; a binding name of an op, list, factory or free function is unique within
    its scope (`process` included), and every other binding name within its language.
    """
    if entry.get("kind") == "field":
        return f"field:{entry.get('record')}"
    if entry.get("kind") == "constant" and "enum" in entry and name_field != "c":
        return f"member:{name_field}:{entry.get('enum')}"
    if name_field != "c" and entry.get("kind") in _SCOPED_NAME_KINDS:
        return f"{name_field}@{entry.get('scope')}"
    return name_field


def _public_namespaces(entry: dict, name_field: str, collections: Dict[str, dict]) -> List[tuple]:
    """`(namespace, name)` of each public name a `handle:` gives the entry: one per class it lists."""
    name = handle_name(entry, name_field, collections) if name_field != "c" else None
    if not isinstance(name, str):
        return []
    out = []
    for cls in handle_classes(entry):
        space = f"{name_field}@sheet" if cls == "Sheet" else f"{name_field}#{cls}"
        out.append((space, name))
    return out


def _is_c_swap(entries: Dict[str, dict], surfaces: Optional[extract_names.Surfaces], a: str, b: str) -> bool:
    """Entry `b` takes the C name entry `a` renames away, and the two functions differ in parameter count.

    A stale call then fails to compile instead of binding to the other function.
    """
    was_a, was_b = (entries[a].get("was") or {}).get("c"), (entries[b].get("was") or {}).get("c")
    if not (was_a and was_b and was_a == entries[b].get("c")):
        return False
    if surfaces is None:
        # Ledger-only: the parameter counts are proved by the run that reads the header.
        return True
    fa, fb = surfaces.functions.get(was_a), surfaces.functions.get(was_b)
    if fb is None:
        # After the rename the tree declares the two current names instead.
        fa, fb = surfaces.functions.get(entries[a].get("c") or ""), surfaces.functions.get(was_a)
    return fa is not None and fb is not None and len(fa.params) != len(fb.params)


def check_chains(
    entries: Dict[str, dict],
    surfaces: Optional[extract_names.Surfaces] = None,
    collections: Optional[Dict[str, dict]] = None,
) -> List[str]:
    """New names are unique per namespace, old names are claimed once, and no new name is an old name renamed
    away.

    A public name is unique per `(scope or handle.class, name)`, a raw name per `(scope, name)`. The exceptions
    are a C name handed from one function to another with a different parameter count, and a record name a
    handle class takes over.
    """
    problems: List[str] = []
    new: Dict[tuple, List[str]] = defaultdict(list)
    old: Dict[tuple, List[str]] = defaultdict(list)
    renamed_from: Dict[tuple, str] = {}
    handle_names = set(handle_class_names(collections or {}))
    for key, entry in entries.items():
        for name_field in NAME_FIELDS:
            space = _namespace(entry, name_field)
            target = entry.get(name_field)
            was = (entry.get("was") or {}).get(name_field)
            if isinstance(target, str):
                new[(space, target)].append(key)
            for public in _public_namespaces(entry, name_field, collections or {}):
                if public != (space, target):
                    new[public].append(key)
            if isinstance(was, str):
                old[(space, was)].append(key)
                if was != target:
                    renamed_from[(space, was)] = key
    for (space, name), keys in sorted(new.items()):
        if len(set(keys)) > 1:
            problems.append(f"{space} name {name} is the new name of several entries: {', '.join(keys)}")
        owner = renamed_from.get((space, name))
        if owner is not None and owner not in keys:
            if space == "c" and _is_c_swap(entries, surfaces, owner, keys[0]):
                continue
            # A record renamed to `<Class>Properties` hands its former name to the handle class.
            if entries[keys[0]].get("kind") == "typedef" and name in handle_names:
                continue
            problems.append(f"rename chain: {space} name {name} is renamed away by {owner} and into by {keys[0]}")
    for (space, name), keys in sorted(old.items()):
        if len(set(keys)) > 1:
            problems.append(f"{space} name {name} is claimed by several `was`: {', '.join(keys)}")
    return problems


def check_ledger_index(data: dict, surfaces: Optional[extract_names.Surfaces] = None) -> List[str]:
    """The checker's own name index must resolve every name without a double claim."""
    ledger = Ledger(data)
    if surfaces is not None:
        ledger.c_param_counts = {name: len(fn.params) for name, fn in surfaces.functions.items()}
    else:
        ledger.trust_c_handover = True
    for name_field in NAME_FIELDS:
        ledger.key_for(name_field, "")
    return list(dict.fromkeys(ledger.problems))


def check_clearable(entries: Dict[str, dict], dts_text: str) -> List[str]:
    """Each `clearable` field is a field of the record's type (an interface or a `Patch<...>` alias) in the d.ts."""
    problems: List[str] = []
    surface = JsSurface(dts_text)
    for key, entry in sorted(entries.items()):
        if not isinstance(entry, dict) or not (entry.get("clearable") == [] or _names_list(entry.get("clearable"))):
            continue
        name = entry.get("js")
        keys = _patch_keys(surface.decls, surface.raw, name) if isinstance(name, str) else None
        if keys is None:
            problems.append(f"entries.{key}: clearable needs the record's `js` type in the d.ts")
            continue
        fields = {_snake_case(k) for k in keys}
        for field in entry["clearable"]:
            if field not in fields:
                problems.append(f"entries.{key}: clearable field {field} is not a field of {entry['js']}")
    return problems


def _camel_path(path: str) -> List[str]:
    return [_list_js(segment) for segment in path.split(".")]


def _path_resolves(surface: JsSurface, type_name: str, segments: Sequence[str]) -> bool:
    """Whether `segments` name a field of `type_name`, each earlier one typed by one nested record (an array of
    one too); `string` and absent alternatives aside."""
    keys = _js_record_keys(surface.decls, surface.raw, type_name)
    if keys is None or segments[0] not in keys:
        return False
    if len(segments) == 1:
        return True
    kinds = [k for k in _strip_null(keys[segments[0]], True) if k not in ("string", "undefined")]
    if len(kinds) != 1:
        return False
    nested = re.fullmatch(r"(?:readonly\s+)?\(?([\w$]+)\)?\[\]|ReadonlyArray<([\w$]+)>|([\w$]+)", kinds[0].strip())
    return nested is not None and _path_resolves(surface, next(g for g in nested.groups() if g), segments[1:])


def _patch_records_by_type(entries: Dict[str, dict], surface: JsSurface, collections: Dict[str, dict]) -> dict:
    """`{type name: {patch_record, ...}}`: the records each type named in a patch op's patch argument is the
    patch of."""
    out: Dict[str, set] = defaultdict(set)
    scope_classes = {"workbook": ["Workbook", "WorkbookOoxml"], "sheet": ["Sheet", "SheetOoxml"]}
    for entry in entries.values():
        record = entry.get("patch_record") if isinstance(entry, dict) and entry.get("patch") is True else None
        if not isinstance(record, str) or record == PATCH_RECORD_NONE:
            continue
        if entry_handle(entry) is not None:
            name = handle_name(entry, "js", collections)
            classes = handle_classes(entry)
        else:
            name = entry.get("js")
            classes = scope_classes.get(entry.get("scope"), [])
        for cls in classes:
            member = surface.interface_members(cls).get(name) if isinstance(name, str) else None
            if member is not None and member.is_method and member.params:
                for token in re.findall(r"[A-Za-z_$][\w$]*", member.params[-1][1]):
                    out[token].add(record)
    return out


def check_read_only_keys(
    entries: Dict[str, dict], dts_text: str, collections: Optional[Dict[str, dict]] = None
) -> List[str]:
    """Each `read_only_keys` path is no field of the record and a field of one of its read twins: the
    `patch_record` of a patch op taking the record, `X` of `XInput` / `XPatch` and `XProperties`. A record with
    no read twin in the d.ts (a read of another name, such as `RowBreak` for `PageBreakInput`) is held to the
    first half only."""
    surface = JsSurface(dts_text)
    patched = _patch_records_by_type(entries, surface, collections or {})
    problems: List[str] = []
    for key, entry in sorted(entries.items()):
        if not isinstance(entry, dict) or not _names_list(entry.get("read_only_keys")):
            continue
        name = entry.get("js")
        if not isinstance(name, str) or _js_record_keys(surface.decls, surface.raw, name) is None:
            problems.append(f"entries.{key}: read_only_keys needs the record's `js` type in the d.ts")
            continue
        base = re.sub(r"(Input|Patch)$", "", name)
        twins = sorted(
            t
            for t in {*patched.get(name, ()), base, f"{base}Properties"} - {name}
            if _js_record_keys(surface.decls, surface.raw, t) is not None
        )
        for path in entry["read_only_keys"]:
            segments = _camel_path(path)
            if _path_resolves(surface, name, segments):
                problems.append(f"entries.{key}: read_only_keys {path} is a field of {name}; it is not read-only")
            elif twins and not any(_path_resolves(surface, t, segments) for t in twins):
                problems.append(f"entries.{key}: read_only_keys {path} is a field of no read twin ({', '.join(twins)})")
    return problems


def _list_js(py_name: str) -> str:
    return py_name.split("_")[0] + "".join(w[:1].upper() + w[1:] for w in py_name.split("_")[1:])


def check_rules(entries: Dict[str, dict], collections: Optional[Dict[str, dict]] = None) -> List[str]:
    """R1 / R2 on C function names, R5 on binding names, `derive_type` on type names.

    The result record of a collection published as handles derives to `<handle_class>Properties`.
    """
    problems: List[str] = []
    handles = handle_collections(collections or {})
    for key, entry in entries.items():
        kind = entry.get("kind")
        c_name = entry.get("c")
        exempt = "exception" in entry or entry.get("review") is True
        if kind in FUNCTION_KINDS and isinstance(c_name, str):
            for rule_id, message in check_c_name(c_name):
                problems.append(f"{key}: {rule_id}: {message}")
            if not exempt:
                prefix = "Workbook." if kind == "factory" else ""
                for name_field, derive in (("js", derive_js), ("py", derive_py)):
                    want = prefix + derive(c_name)
                    if name_field in entry and entry[name_field] != want:
                        problems.append(f"{key}: R5.derive: {name_field} {entry[name_field]} should be {want}")
            scope, _ = split_c_name(c_name)
            if kind == "op" and entry.get("scope") not in (scope, "process" if scope is None else scope):
                problems.append(f"{key}: scope {entry.get('scope')} disagrees with the C name {c_name}")
        elif kind in TYPE_KINDS and isinstance(c_name, str):
            if not c_name.endswith("_t"):
                problems.append(f"{key}: R7.c.typedef: {c_name} is not named fm_<name>_t")
            if not exempt:
                wants = {derive_type(c_name, entry.get("role"))}
                if kind == "record":
                    wants.add(derive_record_type(c_name, entry.get("role"), handles))
                for name_field in ("js", "py"):
                    if name_field in entry and entry[name_field] not in wants:
                        want = " or ".join(sorted(wants))
                        problems.append(f"{key}: R5.derive: {name_field} {entry[name_field]} should be {want}")
        elif kind == "list" and "js" in entry and "py" in entry and entry["js"] != _list_js(entry["py"]):
            problems.append(f"{key}: R5.derive: list names {entry['js']} / {entry['py']} disagree")
        elif kind in ("free", "op") and c_name is None and not exempt and "flag" not in entry and "drop" not in entry:
            if any(f in entry for f in ("js", "py")):
                problems.append(f"{key}: a binding-only {kind} needs `exception:`")
    return problems


def check_typed_slots(entries: Dict[str, dict], surfaces: Optional[extract_names.Surfaces]) -> List[str]:
    """Every op whose C signature carries an enum, or whose verb is a predicate, says so."""
    if surfaces is None:
        return []
    enums = {name for name, t in surfaces.typedefs.items() if t.kind == "enum"}
    problems: List[str] = []
    for key, entry in entries.items():
        if entry.get("kind") not in FUNCTION_KINDS:
            continue
        fn = surfaces.functions.get((entry.get("was") or {}).get("c", "")) or surfaces.functions.get(entry.get("c", ""))
        if fn is None:
            continue
        params = entry.get("params") or {}
        for p in fn.params:
            if p.base_type not in enums:
                continue
            if p.name.startswith("out") and p.pointer_depth == 1:
                if not str(entry.get("returns", "")).startswith("enum:"):
                    problems.append(f"{key}: returns enum {p.base_type} but carries no `returns: enum:<key>`")
            elif p.pointer_depth == 0 and not str(slot_type(params.get(p.name)) or "").startswith("enum:"):
                problems.append(f"{key}: parameter {p.name} is enum {p.base_type} but `params` omits it")
        names = {p.name for p in fn.params}
        for param, spec in params.items():
            if isinstance(spec, dict) and param not in names:
                problems.append(f"{key}: params gives {param} a default, but {fn.name} has no such parameter")
        _, words = split_c_name(entry.get("c", ""))
        if words[:1] in (["is"], ["has"]) and entry.get("returns") != "bool":
            problems.append(f"{key}: predicate {entry.get('c')} needs `returns: bool`")
    return problems


def _js_cell_pairs(params) -> List[List[int]]:
    """Positions of adjacent `row` / `col` (or `<x>Row` / `<x>Col`) arguments."""
    out = []
    for i in range(len(params) - 1):
        match = _JS_CELL_PAIR_RE.fullmatch(params[i][0])
        if match and params[i + 1][0] in (f"{match.group(1)}col", f"{match.group(1)}Col"):
            out.append([i, i + 1])
    return out


def _js_rect_positions(params) -> List[int]:
    """Positions of arguments typed `CellRect`, a union holding it, or an array of either."""
    return [i for i, (_, ptype) in enumerate(params) if re.search(r"(?<![\w$])CellRect(?![\w$])", ptype)]


def _js_ref_arg(params) -> List:
    """`[position, kind]` of the argument typed by a `*Ref` of `JS_REF_KINDS`, or `[]`."""
    for i, (_, ptype) in enumerate(params):
        for ref, kind in JS_REF_KINDS.items():
            if re.search(rf"(?<![\w$]){ref}(?![\w$])", ptype):
                return [i, kind]
    return []


_SPAN_AXIS_WORDS = (("Column", "col"), ("Row", "row"))


def _js_span_args(name: str, params, overloads: Sequence = ()) -> List[List]:
    """`[position, axis, form]` of each `(first, last)` or `(row | col, count)` pair of arguments, and of each
    lone `row` / `col` argument a sibling overload takes as `a1: string` at the same position (`single`).

    A `(first, last)` pair takes its axis from the method name (`setColumnWidth` -> `col`), a `(row | col,
    count)` pair and a `single` argument from the argument's name.
    """
    out: List[List] = []
    paired = set()
    for i in range(len(params) - 1):
        a, b = params[i][0], params[i + 1][0]
        if (a, b) == ("first", "last"):
            axis = next((ax for word, ax in _SPAN_AXIS_WORDS if word in name), None)
            out.append([i, axis or "?", "span"])
        elif a in ("row", "col") and b == "count":
            out.append([i, a, "start_count"])
        elif a == "row" and b == "col":
            paired.add(i)
        else:
            continue
        paired.update((i, i + 1))
    for i, (pname, _ptype) in enumerate(params):
        if pname not in ("row", "col") or i in paired:
            continue
        if any(len(o.params) > i and o.params[i] == ("a1", "string") for o in overloads):
            out.append([i, pname, "single"])
    return sorted(out, key=lambda span: span[0])


def _overloads(surface: JsSurface, member) -> List:
    """Every signature of `member`'s method in the interface that declares `member`."""
    for decl in surface.decls.values():
        if decl.kind == "interface" and member in decl.members:
            return [m for m in decl.members if m.name == member.name and m.is_method]
    return [member]


def _handle_members(dts_text: str, entry: dict, collections: Dict[str, dict]) -> List[tuple]:
    """`(owner, member or None)` of each class the entry's handle method is declared on."""
    surface = JsSurface(dts_text)
    name = handle_name(entry, "js", collections)
    out = []
    for cls in handle_classes(entry):
        members = {n: m for n, m in surface.interface_members(cls).items() if m.is_method}
        out.append((f"interface {cls}", members.get(name) if isinstance(name, str) else None))
    return out


def check_address_args(
    entries: Dict[str, dict], dts_text: str, collections: Optional[Dict[str, dict]] = None
) -> List[str]:
    """Each op's `cell_args` / `rect_args` / `span_args`, and each op's or list's `ref_args`, equal what the
    first signature of its JS member declares; an op with `handle:` is read from `interface <handle.class>`,
    whose methods omit the arguments the handle supplies, any other op from the class of its scope first. An op
    whose first signature takes a row or column span without `span_args` is reported."""
    surface = JsSurface(dts_text)
    pool = surface.callables()
    scope_classes = {"workbook": "Workbook", "sheet": "Sheet"}
    problems: List[str] = []
    for key, entry in sorted(entries.items()):
        kind = entry.get("kind") if isinstance(entry, dict) else None
        if kind not in ("op", "list") or entry.get("internal") is True:
            continue
        fields = ("cell_args", "rect_args", "ref_args", "span_args") if kind == "op" else ("ref_args",)
        if entry_handle(entry) is not None:
            members = _handle_members(dts_text, entry, collections or {})
        else:
            js = entry.get("js")
            cls = scope_classes.get(entry.get("scope"))
            member = surface.interface_members(cls).get(js) if cls and isinstance(js, str) else None
            members = [(WASM_DTS.name, member or (pool.get(js) if isinstance(js, str) else None))]
        for owner, member in members:
            if member is None:
                for field in fields:
                    if field in entry:
                        problems.append(f"{key}: {field} names no member of {owner}")
                continue
            derived = {
                "cell_args": lambda: _js_cell_pairs(member.params),
                "rect_args": lambda: _js_rect_positions(member.params),
                "ref_args": lambda: _js_ref_arg(member.params),
                "span_args": lambda: _js_span_args(member.name, member.params, _overloads(surface, member)),
            }
            for field, want in ((f, derived[f]()) for f in fields):
                have = entry.get(field) or []
                if have != want:
                    problems.append(
                        f"{key}: {field} is {have or 'absent'}, but `{member.name}` in {owner} declares {want or 'none'}"
                    )
    return problems


# Entry kinds whose names are top-level exports, so a current name must exist where the entry says it does.
_PRESENT_KINDS = frozenset({"record", "enum", "typedef"})
_SURFACES_OF_FIELD = {"c": ("c",), "js": ("wasm", "native"), "py": ("py",)}
_DTS_EXPORT_RE = re.compile(
    r"^export\s+(?:declare\s+)?(?:class|interface|type|const|function|enum)\s+([A-Za-z_$][\w$]*)", re.M
)


def check_presence(entries: Dict[str, dict], surfaces: extract_names.Surfaces, dts_text: str) -> List[str]:
    """Each current record, enum or typedef name exists on every surface its entry lists, and no entry drops it.

    A name that leaves a surface carries `drop:` (or `replaced_by:`); a current name the tree no longer exports
    would otherwise reach the migration table as a replacement nobody can import.
    """
    exported: Dict[str, set] = defaultdict(set)
    for surface, symbol in surfaces.exported():
        exported[surface].add(symbol)
    for name in _DTS_EXPORT_RE.findall(dts_text):
        exported["wasm"].add(name)
        exported["native"].add(name)
    dropped: Dict[tuple, List[str]] = defaultdict(list)
    for key, entry in entries.items():
        for field, name in (entry.get("drop") or {}).items():
            if field in _SURFACES_OF_FIELD:
                dropped[(field, name)].append(key)
    problems = []
    for key, entry in entries.items():
        if entry.get("kind") not in _PRESENT_KINDS or "replaced_by" in entry:
            continue
        for field, field_surfaces in _SURFACES_OF_FIELD.items():
            name = entry.get(field)
            if not isinstance(name, str) or (entry.get("drop") or {}).get(field) == name:
                continue
            missing = [s for s in field_surfaces if s in entry.get("surfaces", []) and name not in exported[s]]
            if missing:
                problems.append(f"R0.presence: {key}: {field} {name} is not exported on {', '.join(missing)}")
            others = [k for k in dropped[(field, name)] if k != key]
            if others:
                problems.append(
                    f"R0.presence: {key}: {field} {name} is current here but dropped by {', '.join(others)}"
                )
    return problems


def check_coverage(
    entries: Dict[str, dict], surfaces: extract_names.Surfaces, collections: Optional[Dict[str, dict]] = None
) -> List[str]:
    """Each exported symbol is some entry's current or `was` name, on a surface that entry lists.

    A current name wins, so the tree is covered both before and after the renames land. A dropped
    name or a `replaced_by` entry needs no surface: it leaves the tree. A handle method is the
    `Class.method` of an entry's `handle:`; an identity property of a handle (`index`, `objectId`, ...)
    that no entry claims is built by the facade and needs none.
    """
    owner: Dict[tuple, str] = {}
    for key, entry in entries.items():
        if entry.get("kind") == "field":
            continue
        for name_field in NAME_FIELDS:
            name = entry.get(name_field)
            if isinstance(name, str):
                owner.setdefault((name_field, name), key)
            if name_field != "c":
                for cls in handle_classes(entry):
                    method = handle_name(entry, name_field, collections or {})
                    if method:
                        owner.setdefault((name_field, f"{cls}.{method}"), key)
    for key, entry in entries.items():
        for name_field, name in (entry.get("was") or {}).items():
            owner.setdefault((name_field, name), key)
    problems = []
    for surface, symbol in surfaces.exported():
        name_field = "c" if surface == "c" else "py" if surface == "py" else "js"
        key = owner.get((name_field, symbol))
        cls, _, member = symbol.partition(".")
        if key is None and member in HANDLE_IDENTITY_PROPERTIES and cls in handle_class_names(collections or {}):
            continue
        if key is None:
            problems.append(
                f"R0.coverage: [{surface}] {symbol} is exported but no entry's `{name_field}` or `was.{name_field}` names it"
            )
            continue
        entry = entries[key]
        leaving = "replaced_by" in entry or symbol in (entry.get("drop") or {}).values()
        if not leaving and surface not in entry.get("surfaces", []):
            problems.append(f"R0.coverage: [{surface}] {symbol} is covered by {key}, whose surfaces omit {surface}")
    return problems


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--ledger", type=Path, default=LEDGER_PATH)
    parser.add_argument("--coverage", action="store_true", help="also require every exported symbol to be covered")
    parser.add_argument(
        "--ledger-only", action="store_true", help="check the ledger alone; read neither the d.ts nor the C header"
    )
    args = parser.parse_args(argv)
    if args.coverage and args.ledger_only:
        parser.error("--coverage reads the tree; it does not combine with --ledger-only")
    if not args.ledger.is_file():
        print(f"validate_names: {args.ledger} does not exist")
        return 1
    try:
        data = _load(args.ledger)
    except ImportError:
        print("validate_names: PyYAML is not importable; run under tools/oracle/.venv")
        return 2
    problems = check_schema(data)
    if not problems:
        entries = data["entries"]
        surfaces = None if args.ledger_only else extract_names.scan()
        problems += check_chains(entries, surfaces, data["collections"])
        problems += check_ledger_index(data, surfaces)
        problems += check_rules(entries, data["collections"])
        if surfaces is not None:
            problems += check_typed_slots(entries, surfaces)
            dts_text = WASM_DTS.read_text(encoding="utf-8")
            problems += check_address_args(entries, dts_text, data["collections"])
            problems += check_clearable(entries, dts_text)
            problems += check_read_only_keys(entries, dts_text, data["collections"])
            problems += check_presence(entries, surfaces, dts_text)
        if args.coverage:
            problems += check_coverage(entries, surfaces, data["collections"])
        flagged = [(f"entries.{k}", e["flag"]) for k, e in entries.items() if isinstance(e, dict) and "flag" in e]
        flagged += [
            (f"entries.{k}", f"review: {e.get('note', '')}")
            for k, e in entries.items()
            if isinstance(e, dict) and e.get("review") is True
        ]
        flagged += [
            (f"enum_distinct[{i}]", f"review: {item.get('note', '')}")
            for i, item in enumerate(data.get("enum_distinct") or [])
            if isinstance(item, dict) and item.get("review") is True
        ]
        flagged += [(f"collections.{k}", c["flag"]) for k, c in data["collections"].items() if "flag" in c]
        for where, reason in flagged:
            print(f"flag: {where}: {reason}")
        print(f"validate_names: {len(entries)} entries, {len(flagged)} flagged")
    for problem in problems:
        print(f"validate_names: {problem}")
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
