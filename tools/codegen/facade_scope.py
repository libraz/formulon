"""Which binding class owns each ledger op, read from ``tools/api/names.yaml``.

The JS and Python bindings publish two classes: ``Workbook`` holds the
``scope: workbook`` ops and ``Sheet`` the ``scope: sheet`` ones, without their
leading sheet argument. An op with ``handle:`` is published on its handle
classes instead (``Pivot``, ``Chart``, ...), which supply the leading raw
arguments ``handle.supplies`` names. The raw engine classes (embind, N-API, the
Python FFI layer) stay flat and take the sheet index first. This module is the
one reader of that split; ``gen_facade_tables.py`` writes the facade tables
from it and ``gen_wasm_positional_guards.py`` reads the generated copy to word
the WASM argument guards.

A name is *public* on a language when its op lists one of the language's
binding surfaces, does not drop that name, is not ``internal`` and has no
``handle:``. A *raw* name also counts a dropped, internal or handle one: the raw
class still registers it, and the facade calls it. Generated tables are keyed
``Class.method``.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path
from typing import Dict, Iterator, List, Optional, Tuple

REPO_ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO_ROOT))

from tools.dev.binding_drift.api_conventions import (  # noqa: E402
    Ledger,
    c_verb,
    common_base,
    entry_handle,
    handle_classes,
    handle_name,
    load_ledger,
    parse_header,
)
from tools.dev.binding_drift.api_conventions_arguments import _js_fields, _patch_keys  # noqa: E402
from tools.dev.binding_drift.api_conventions_bindings import JsSurface  # noqa: E402
from tools.dev.binding_drift.api_conventions_shapes import _strip_null  # noqa: E402
from tools.dev.binding_drift.surface_files import CAPI_HEADER, WASM_DTS  # noqa: E402

LANGUAGES = ("js", "py")
SURFACES = {"js": ("wasm", "native"), "py": ("py",)}
SCOPES = ("workbook", "sheet")
SCOPE_CLASS = {"workbook": "Workbook", "sheet": "Sheet"}
HANDWRITTEN_MODE = "handwritten"
NPM_COMMON = REPO_ROOT / "packages" / "npm" / "common.mjs"


def _ops(ledger: Ledger, kinds: Tuple[str, ...] = ("op", "list")) -> Iterator[Tuple[str, dict]]:
    for key, entry in sorted(ledger.entries.items()):
        if isinstance(entry, dict) and entry.get("kind") in kinds:
            yield key, entry


def _on_language(entry: dict, language: str) -> bool:
    listed = entry.get("surfaces")
    return listed is None or any(s in listed for s in SURFACES[language])


def _dropped(entry: dict, language: str) -> Optional[str]:
    drop = entry.get("drop")
    name = drop.get(language) if isinstance(drop, dict) else None
    return name if isinstance(name, str) else None


def public_name(entry: dict, language: str) -> Optional[str]:
    """The entry's name on `language`'s `Workbook` / `Sheet` (or module), or None when it has none there.

    An `internal` op and an op published on a handle class have none.
    """
    name = entry.get(language)
    if not isinstance(name, str) or not _on_language(entry, language) or _dropped(entry, language) == name:
        return None
    if entry.get("internal") is True or entry_handle(entry) is not None:
        return None
    return name


def qualified(entry: dict, language: str, collections: Dict[str, dict]) -> List[str]:
    """`Class.method` of every public name of the entry: its handle methods, else its `Workbook` / `Sheet` one."""
    if not _on_language(entry, language):
        return []
    if entry_handle(entry) is not None:
        name = handle_name(entry, language, collections)
        return [f"{cls}.{name}" for cls in handle_classes(entry)] if name else []
    name = public_name(entry, language)
    owner = SCOPE_CLASS.get(entry.get("scope") or "")
    return [f"{owner}.{name}"] if name and owner else []


def raw_name(entry: dict, language: str) -> Optional[str]:
    """The name the raw class registers for the entry: the public name, else the dropped one."""
    if not _on_language(entry, language):
        return None
    name = entry.get(language)
    if isinstance(name, str):
        return name
    return _dropped(entry, language)


def module_members() -> frozenset:
    """The members of the object `createFormulon()` resolves to: `FORMULON_MEMBERS` of `common.mjs`."""
    text = NPM_COMMON.read_text(encoding="utf-8")
    found = re.search(r"\bFORMULON_MEMBERS\s*=\s*Object\.freeze\(\[(.*?)\]\)", text, re.S)
    if found is None:
        raise SystemExit(f"facade_scope: {NPM_COMMON} does not declare FORMULON_MEMBERS")
    return frozenset(re.findall(r"'([^']+)'", found.group(1)))


def methods(ledger: Ledger, scope: str, language: str) -> List[str]:
    """Sorted public names of the `scope` ops and lists on `language`.

    The JS `Workbook` also carries the `scope: process` ops that are not members of the `Formulon` object
    (`getLocaleFacts`, `compareText`, ...): they need no workbook, but the npm packages publish them as
    `Workbook` methods.
    """
    module = module_members() if scope == "workbook" and language == "js" else frozenset()
    out = set()
    for _key, entry in _ops(ledger):
        owned = entry.get("scope") == scope
        if module and entry.get("scope") == "process":
            owned = entry.get(language) not in module
        if owned:
            name = public_name(entry, language)
            if name is not None:
                out.add(name)
    return sorted(out)


def raw_sheet_methods(ledger: Ledger, language: str) -> List[str]:
    """Sorted raw names of the sheet-scope ops on `language`: every raw method whose first argument is the
    sheet index the `Sheet` facade supplies, the ones the published classes drop included."""
    out = set()
    for _key, entry in _ops(ledger, ("op",)):
        if entry.get("scope") == "sheet":
            name = raw_name(entry, language)
            if name is not None:
                out.add(name)
    return sorted(out)


def handwritten(ledger: Ledger, language: str) -> List[str]:
    """Sorted public names of the ops the facade implements by hand (`facade: handwritten`)."""
    out = set()
    for _key, entry in _ops(ledger):
        if entry.get("facade") == HANDWRITTEN_MODE:
            name = public_name(entry, language)
            if name is not None:
                out.add(name)
    return sorted(out)


def list_families(ledger: Ledger, language: str) -> Dict[str, Tuple[str, str]]:
    """`{Class.listName: (countName, getName)}` for every `kind: list` entry the facade assembles from a raw
    count and element getter, that is one whose `from` ops are raw methods on `language`."""
    out: Dict[str, Tuple[str, str]] = {}
    for _key, entry in _ops(ledger, ("list",)):
        sources = entry.get("from") or []
        if len(sources) != 2:
            continue
        raw = [raw_name(ledger.entry(op), language) for op in sources]
        if all(isinstance(n, str) for n in raw):
            for name in qualified(entry, language, ledger.collections):
                out[name] = (raw[0], raw[1])
    return dict(sorted(out.items()))


def _positions(ledger: Ledger, language: str, field: str, kinds: Tuple[str, ...] = ("op",)) -> Dict[str, list]:
    out: Dict[str, list] = {}
    for _key, entry in _ops(ledger, kinds):
        if not entry.get(field):
            continue
        for name in qualified(entry, language, ledger.collections):
            out[name] = [list(v) if isinstance(v, list) else v for v in entry[field]]
    return dict(sorted(out.items()))


def cell_args(ledger: Ledger, language: str) -> Dict[str, list]:
    """`{Class.method: [[row_position, col_position], ...]}` of the ops whose ledger entry carries `cell_args`."""
    return _positions(ledger, language, "cell_args")


def rect_args(ledger: Ledger, language: str) -> Dict[str, list]:
    """`{Class.method: [position, ...]}` of the ops whose ledger entry carries `rect_args`."""
    return _positions(ledger, language, "rect_args")


def span_args(ledger: Ledger, language: str) -> Dict[str, list]:
    """`{Class.method: [[position, axis, form], ...]}` of the ops whose ledger entry carries `span_args`."""
    return _positions(ledger, language, "span_args")


def ref_args(ledger: Ledger, language: str) -> Dict[str, list]:
    """`{Class.method: [position, kind]}` of the ops and lists whose ledger entry carries `ref_args`."""
    out: Dict[str, list] = {}
    for _key, entry in _ops(ledger, ("op", "list")):
        if entry.get("ref_args"):
            for name in qualified(entry, language, ledger.collections):
                out[name] = list(entry["ref_args"])
    return dict(sorted(out.items()))


def handle_forwarders(ledger: Ledger, language: str) -> Dict[str, Optional[str]]:
    """`{Class.method: rawName}` of every handle method that is not a property; None for one the facade builds
    (a handwritten op, or a list `LIST_FAMILIES` assembles)."""
    out: Dict[str, Optional[str]] = {}
    for _key, entry in _ops(ledger):
        handle = entry_handle(entry)
        if handle is None or handle.get("property") is True:
            continue
        built = entry.get("facade") == HANDWRITTEN_MODE or entry.get("kind") == "list"
        raw = None if built else raw_name(entry, language)
        for name in qualified(entry, language, ledger.collections):
            out[name] = raw
    return dict(sorted(out.items()))


def handle_properties(ledger: Ledger, language: str) -> Dict[str, str]:
    """`{Class.property: rawGetter}` of the identity properties a handle reads through a raw getter."""
    out: Dict[str, str] = {}
    for _key, entry in _ops(ledger):
        handle = entry_handle(entry)
        raw = raw_name(entry, language)
        if handle is None or handle.get("property") is not True or raw is None:
            continue
        for name in qualified(entry, language, ledger.collections):
            out[name] = raw
    return dict(sorted(out.items()))


def handle_methods(ledger: Ledger) -> List[Tuple[str, str, int]]:
    """`(raw JS name, label, supplied)` of every C-backed raw method a handle class calls, sorted by raw name.

    The label is `Class.method` on the nearest class every publishing class shares (`Drawing.rename` for a raw
    `Picture` and `Chart` both call); `supplied` counts the leading raw arguments the handle fills in.
    """
    bases = ledger.handle_bases()
    out: Dict[str, Tuple[str, str, int]] = {}
    for _key, entry in _ops(ledger):
        handle = entry_handle(entry)
        raw = raw_name(entry, "js")
        name = handle_name(entry, "js", ledger.collections) if handle is not None else None
        # Only an op a C function backs has a raw registration; a handwritten one still calls it.
        if handle is None or raw is None or name is None or entry.get("kind") != "op":
            continue
        if not isinstance(entry.get("c"), str):
            continue
        owner = common_base(handle_classes(entry), bases)
        out[raw] = (raw, f"{owner}.{name}", len(handle.get("supplies") or []))
    return [out[k] for k in sorted(out)]


def _patch_member(surface: JsSurface, entry: dict, language_name: str, cls: str):
    """The first signature of the op's JS member on `cls`: a handle interface, `Workbook` or `Sheet`."""
    members = surface.interface_members(cls)
    member = members.get(language_name)
    return member if member is not None and member.is_method else None


def _patch_type(params) -> Optional[str]:
    kinds = _strip_null(params[-1][1], True) if params else []
    return kinds[0] if len(kinds) == 1 else None


def patch_keys(ledger: Ledger, dts: Optional[str] = None) -> Dict[str, list]:
    """`{Class.method: [position, [key, ...]]}` of the `patch: true` ops: the position of the patch argument
    (the last one of the member's first signature in src/wasm/formulon.d.ts) and, sorted, the keys the patch
    accepts: those its type declares (`| null` aside) and the fields of its `patch_record`, so a record a
    getter returned passes back whole. An op whose patch type the declaration reader cannot resolve is an
    error."""
    surface = JsSurface(WASM_DTS.read_text(encoding="utf-8") if dts is None else dts)
    out: Dict[str, list] = {}
    for key, entry in _ops(ledger, ("op",)):
        if entry.get("patch") is not True:
            continue
        for qualified_name in _qualified_js(ledger, entry):
            cls, name = qualified_name.split(".", 1)
            member = _patch_member(surface, entry, name, cls)
            if member is None:
                continue
            params = member.params
            patch_type = _patch_type(params)
            keys = _patch_keys(surface.decls, surface.raw, patch_type) if patch_type else None
            if not keys:
                raise ValueError(f"{key}: the patch type of `{qualified_name}` declares no keys the reader can resolve")
            record = entry.get("patch_record")
            fields = _js_fields(surface.decls, surface.raw, record) if record and record != "none" else None
            out[qualified_name] = [len(params) - 1, sorted(set(keys) | set(fields or {}))]
    return dict(sorted(out.items()))


def _qualified_js(ledger: Ledger, entry: dict) -> List[str]:
    return qualified(entry, "js", ledger.collections)


def _field_type(surface: JsSurface, type_name: str, field: str) -> Optional[str]:
    """The single declared type of `field` in the patch type `type_name`, `| null` / `undefined` aside."""
    keys = _patch_keys(surface.decls, surface.raw, type_name) or {}
    kinds = [k for k in _strip_null(keys.get(field, ""), True) if k != "undefined"]
    return kinds[0] if len(kinds) == 1 else None


def patch_nested_keys(ledger: Ledger, dts: Optional[str] = None) -> Dict[str, Dict[str, list]]:
    """`{Class.method: {path: [key, ...]}}` of the `patch_nested` ops: for each path into the patch, sorted, the
    keys of the nested patch type there and the fields of the record the ledger names for it."""
    surface = JsSurface(WASM_DTS.read_text(encoding="utf-8") if dts is None else dts)
    out: Dict[str, Dict[str, list]] = {}
    for key, entry in _ops(ledger, ("op",)):
        nested = entry.get("patch_nested")
        if entry.get("patch") is not True or not isinstance(nested, dict):
            continue
        for qualified_name in _qualified_js(ledger, entry):
            cls, name = qualified_name.split(".", 1)
            member = _patch_member(surface, entry, name, cls)
            root = _patch_type(member.params) if member is not None else None
            if root is None:
                continue
            paths: Dict[str, list] = {}
            for path, record in sorted(nested.items()):
                type_name: Optional[str] = root
                for segment in path.split("."):
                    type_name = _field_type(surface, type_name, segment) if type_name else None
                keys = _patch_keys(surface.decls, surface.raw, type_name) if type_name else None
                if not keys:
                    raise ValueError(f"{key}: `{path}` of `{qualified_name}` resolves to no patch type with keys")
                fields = _js_fields(surface.decls, surface.raw, record) or {}
                paths[path] = sorted(set(keys) | set(fields))
            out[qualified_name] = paths
    return dict(sorted(out.items()))


def structure_ops(ledger: Ledger, language: str, header: Optional[str] = None) -> List[str]:
    """Sorted `Class.method` of the ops that add, remove or reorder sheets: every non-`get` op of the sheet
    collection, and every op whose C function writes `out_sheet_index` without a `get` verb."""
    if header is None:
        header = CAPI_HEADER.read_text(encoding="utf-8")
    functions, _ = parse_header(header)
    by_c = {fn.name: fn for fn in functions}
    sheet_ops = set((ledger.collections.get("sheet") or {}).get("ops") or [])
    out = set()
    for key, entry in _ops(ledger, ("op",)):
        target = entry.get("c")
        fn = by_c.get(target) if isinstance(target, str) else None
        if fn is None or c_verb(target) == "get":
            continue
        writes_sheet = any(p.name == "out_sheet_index" for p in fn.params)
        if writes_sheet or key in sheet_ops:
            out.update(qualified(entry, language, ledger.collections))
    return sorted(out)


def load(path: Optional[Path] = None) -> Ledger:
    """The ledger, failing loudly when it does not parse."""
    ledger = load_ledger(path) if path is not None else load_ledger()
    if ledger.problems:
        raise SystemExit("facade_scope: " + "; ".join(ledger.problems))
    return ledger
