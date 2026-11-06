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
of its first signature in src/wasm/formulon.d.ts; `patch: true` names its
getter's record in `patch_record`, or `none` with a `patch_reason`.

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
    split_c_name,
)
from binding_drift.api_conventions_bindings import JsSurface  # noqa: E402

KINDS = frozenset({"op", "list", "free", "factory", "record", "field", "typedef", "enum", "constant"})
FIELDS = frozenset(
    {
        "kind", "scope", "noun", "qualifies", "record", "role", "from", "c", "js", "py", "enum", "returns",
        "params", "surfaces", "exception", "replaced_by", "closes", "was", "drop", "flag", "sheet_arg", "total",
        "missing", "facade", "review", "note", "cell_args", "rect_args", "patch", "patch_record", "patch_reason",
    }
)  # fmt: skip
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
WASM_DTS = REPO_ROOT / "src" / "wasm" / "formulon.d.ts"
PATCH_RECORD_NONE = "none"
_JS_CELL_PAIR_RE = re.compile(r"(\w*?)(?:row|Row)")


def _load(path: Path):
    import yaml  # type: ignore

    return yaml.safe_load(path.read_text(encoding="utf-8"))


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
            allowed = set(NAME_FIELDS) | ({"closes"} if mapping == "drop" else set())
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
            if "closes" not in entry:
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
        slots = dict(entry.get("params") or {})
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
        if not any(f in entry for f in (*NAME_FIELDS, "was", "replaced_by")):
            problems.append(f"{where}: names nothing")
    for name, collection in collections.items():
        where = f"collections.{name}"
        if not isinstance(collection, dict) or collection.get("key") not in ("index", "natural"):
            problems.append(f"{where}: key must be index or natural")
            continue
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
        elif collection["key"] == "natural" and "access" in collection:
            problems.append(f"{where}: `access` belongs to index-keyed collections")
        if access is not None and not isinstance(collection.get("plural"), str):
            problems.append(f"{where}: `access` needs `plural`")
        if (access == "exempt") != isinstance(collection.get("exempt"), str):
            problems.append(f"{where}: `exempt` is the reason string of an `access: exempt` collection")
    problems += check_enum_groups(data, entries)
    return problems


def _index(value) -> bool:
    return isinstance(value, int) and not isinstance(value, bool) and value >= 0


def _address_and_patch_schema(where: str, entry: dict) -> List[str]:
    """`cell_args` lists `[i, j]` pairs and `rect_args` positions; `patch: true` carries `patch_record`, and
    `patch_record: none` a `patch_reason`."""
    problems: List[str] = []
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


def _namespace(entry: dict, name_field: str) -> str:
    return f"field:{entry.get('record')}" if entry.get("kind") == "field" else name_field


def _is_c_swap(entries: Dict[str, dict], surfaces: Optional[extract_names.Surfaces], a: str, b: str) -> bool:
    """Entry `b` takes the C name entry `a` renames away, and the two functions differ in parameter count.

    A stale call then fails to compile instead of binding to the other function.
    """
    if surfaces is None:
        return False
    was_a, was_b = (entries[a].get("was") or {}).get("c"), (entries[b].get("was") or {}).get("c")
    if not (was_a and was_b and was_a == entries[b].get("c")):
        return False
    fa, fb = surfaces.functions.get(was_a), surfaces.functions.get(was_b)
    if fb is None:
        # After the rename the tree declares the two current names instead.
        fa, fb = surfaces.functions.get(entries[a].get("c") or ""), surfaces.functions.get(was_a)
    return fa is not None and fb is not None and len(fa.params) != len(fb.params)


def check_chains(entries: Dict[str, dict], surfaces: Optional[extract_names.Surfaces] = None) -> List[str]:
    """New names are unique, old names are claimed once, and no new name is an old name renamed away.

    The one exception is a C name handed from one function to another with a different parameter count.
    """
    problems: List[str] = []
    new: Dict[tuple, List[str]] = defaultdict(list)
    old: Dict[tuple, List[str]] = defaultdict(list)
    renamed_from: Dict[tuple, str] = {}
    for key, entry in entries.items():
        for name_field in NAME_FIELDS:
            space = _namespace(entry, name_field)
            target = entry.get(name_field)
            was = (entry.get("was") or {}).get(name_field)
            if isinstance(target, str):
                new[(space, target)].append(key)
            if isinstance(was, str):
                old[(space, was)].append(key)
                if was != target:
                    renamed_from[(space, was)] = key
    for (space, name), keys in sorted(new.items()):
        if len(keys) > 1:
            problems.append(f"{space} name {name} is the new name of several entries: {', '.join(keys)}")
        owner = renamed_from.get((space, name))
        if owner is not None and owner not in keys:
            if space == "c" and _is_c_swap(entries, surfaces, owner, keys[0]):
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
    for name_field in NAME_FIELDS:
        ledger.key_for(name_field, "")
    return list(dict.fromkeys(ledger.problems))


def _list_js(py_name: str) -> str:
    return py_name.split("_")[0] + "".join(w[:1].upper() + w[1:] for w in py_name.split("_")[1:])


def check_rules(entries: Dict[str, dict]) -> List[str]:
    """R1 / R2 on C function names, R5 on binding names, `derive_type` on type names."""
    problems: List[str] = []
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
                    wants.add(derive_record_type(c_name, entry.get("role")))
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
            elif p.pointer_depth == 0 and not str(params.get(p.name, "")).startswith("enum:"):
                problems.append(f"{key}: parameter {p.name} is enum {p.base_type} but `params` omits it")
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


def check_address_args(entries: Dict[str, dict], dts_text: str) -> List[str]:
    """Each op's `cell_args` / `rect_args` equal what the first signature of its JS member declares."""
    pool = JsSurface(dts_text).callables()
    problems: List[str] = []
    for key, entry in sorted(entries.items()):
        if not isinstance(entry, dict) or entry.get("kind") != "op":
            continue
        member = pool.get(entry.get("js")) if isinstance(entry.get("js"), str) else None
        if member is None:
            for field in ("cell_args", "rect_args"):
                if field in entry:
                    problems.append(f"{key}: {field} names no member of {WASM_DTS.name}")
            continue
        for field, want in (
            ("cell_args", _js_cell_pairs(member.params)),
            ("rect_args", _js_rect_positions(member.params)),
        ):
            have = entry.get(field) or []
            if have != want:
                problems.append(
                    f"{key}: {field} is {have or 'absent'}, but `{member.name}` in {WASM_DTS.name} declares {want or 'none'}"
                )
    return problems


def check_coverage(entries: Dict[str, dict], surfaces: extract_names.Surfaces) -> List[str]:
    """Each exported symbol is some entry's current or `was` name, on a surface that entry lists.

    A current name wins, so the tree is covered both before and after the renames land. A dropped
    name or a `replaced_by` entry needs no surface: it leaves the tree.
    """
    owner: Dict[tuple, str] = {}
    for key, entry in entries.items():
        if entry.get("kind") == "field":
            continue
        for name_field in NAME_FIELDS:
            name = entry.get(name_field)
            if isinstance(name, str):
                owner.setdefault((name_field, name), key)
    for key, entry in entries.items():
        for name_field, name in (entry.get("was") or {}).items():
            owner.setdefault((name_field, name), key)
    problems = []
    for surface, symbol in surfaces.exported():
        name_field = "c" if surface == "c" else "py" if surface == "py" else "js"
        key = owner.get((name_field, symbol))
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
    args = parser.parse_args(argv)
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
        surfaces = extract_names.scan()
        problems += check_chains(entries, surfaces)
        problems += check_ledger_index(data, surfaces)
        problems += check_rules(entries)
        problems += check_typed_slots(entries, surfaces)
        problems += check_address_args(entries, WASM_DTS.read_text(encoding="utf-8"))
        if args.coverage:
            problems += check_coverage(entries, surfaces)
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
