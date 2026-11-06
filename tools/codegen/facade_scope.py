"""Which binding class owns each ledger op, read from ``tools/api/names.yaml``.

The JS and Python bindings publish two classes: ``Workbook`` holds the
``scope: workbook`` ops and ``Sheet`` the ``scope: sheet`` ones, without their
leading sheet argument. The raw engine classes (embind, N-API, the Python FFI
layer) stay flat and take the sheet index first. This module is the one reader
of that split; ``gen_facade_tables.py`` writes the facade tables from it and
``gen_wasm_positional_guards.py`` reads it to word the WASM argument guards.

A name is *public* on a language when its op lists one of the language's
binding surfaces and does not drop that name. A *raw* name also counts a
dropped one: the raw class still registers it, and the facade calls it.
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
    """The entry's name on `language`'s public classes, or None when it has none there."""
    name = entry.get(language)
    if not isinstance(name, str) or not _on_language(entry, language) or _dropped(entry, language) == name:
        return None
    return name


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
    """`{listName: (countName, getName)}` for every `kind: list` entry the facade assembles from a raw count
    and element getter, that is one whose `from` ops are raw methods on `language`."""
    out: Dict[str, Tuple[str, str]] = {}
    for _key, entry in _ops(ledger, ("list",)):
        name = public_name(entry, language)
        sources = entry.get("from") or []
        if name is None or len(sources) != 2:
            continue
        raw = [raw_name(ledger.entry(op), language) for op in sources]
        if all(isinstance(n, str) for n in raw):
            out[name] = (raw[0], raw[1])
    return dict(sorted(out.items()))


def _positions(ledger: Ledger, language: str, field: str) -> Dict[str, list]:
    out: Dict[str, list] = {}
    for _key, entry in _ops(ledger, ("op",)):
        name = public_name(entry, language)
        if name is not None and entry.get(field):
            out[name] = list(entry[field])
    return dict(sorted(out.items()))


def cell_args(ledger: Ledger, language: str) -> Dict[str, list]:
    """`{publicName: [[row_position, col_position], ...]}` of the ops whose ledger entry carries `cell_args`."""
    return _positions(ledger, language, "cell_args")


def rect_args(ledger: Ledger, language: str) -> Dict[str, list]:
    """`{publicName: [position, ...]}` of the ops whose ledger entry carries `rect_args`."""
    return _positions(ledger, language, "rect_args")


def ref_args(ledger: Ledger, language: str) -> Dict[str, list]:
    """`{publicName: [position, kind]}` of the ops and lists whose ledger entry carries `ref_args`."""
    out: Dict[str, list] = {}
    for _key, entry in _ops(ledger, ("op", "list")):
        name = public_name(entry, language)
        if name is not None and entry.get("ref_args"):
            out[name] = list(entry["ref_args"])
    return dict(sorted(out.items()))


def patch_keys(ledger: Ledger, dts: Optional[str] = None) -> Dict[str, list]:
    """`{jsName: [position, [key, ...]]}` of the `patch: true` ops: the position of the patch argument (the
    last one of the member's first signature in src/wasm/formulon.d.ts) and, sorted, the keys the patch
    accepts: those its type declares (`| null` aside) and the fields of its `patch_record`, so a record a
    getter returned passes back whole. An op whose patch type the declaration reader cannot resolve is an
    error."""
    surface = JsSurface(WASM_DTS.read_text(encoding="utf-8") if dts is None else dts)
    pool = surface.callables()
    out: Dict[str, list] = {}
    for key, entry in _ops(ledger, ("op",)):
        name = public_name(entry, "js")
        if entry.get("patch") is not True or name is None or name not in pool:
            continue
        params = pool[name].params
        kinds = _strip_null(params[-1][1], True) if params else []
        keys = _patch_keys(surface.decls, surface.raw, kinds[0]) if len(kinds) == 1 else None
        if not keys:
            raise ValueError(f"{key}: the patch type of `{name}` declares no keys the reader can resolve")
        record = entry.get("patch_record")
        fields = _js_fields(surface.decls, surface.raw, record) if record and record != "none" else None
        out[name] = [len(params) - 1, sorted(set(keys) | set(fields or {}))]
    return dict(sorted(out.items()))


def structure_ops(ledger: Ledger, language: str, header: Optional[str] = None) -> List[str]:
    """Sorted public names of the ops that add, remove or reorder sheets: every non-`get` op of the sheet
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
            name = public_name(entry, language)
            if name is not None:
                out.add(name)
    return sorted(out)


def load(path: Optional[Path] = None) -> Ledger:
    """The ledger, failing loudly when it does not parse."""
    ledger = load_ledger(path) if path is not None else load_ledger()
    if ledger.problems:
        raise SystemExit("facade_scope: " + "; ".join(ledger.problems))
    return ledger
