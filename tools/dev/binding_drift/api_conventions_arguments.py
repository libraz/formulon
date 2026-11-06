"""Ledger-bound rules on how the binding methods take their arguments.

A cell is a `(row, col)` pair or one A1 string, a rectangle a `CellRect` or an
A1 range string (R25.a1 / R25.rect, driven by the ledger's `cell_args` and
`rect_args`). A collection is enumerated by `list<Plural>()`, not counted
(R25.count). A partial update takes a declared patch whose keys are fields of
the getter's record (R18.all, driven by `patch` / `patch_record`). Tables,
pivots, CF rules and protected ranges are addressed by index or name
(R26.ref), sheets by `SheetRef` (R27.sheetparam).

A member's signature is its first overload (the coordinate form); R25.a1
alone reads every overload.
"""

from __future__ import annotations

import re
from typing import Dict, Iterable, Iterator, List, Optional, Sequence, Set, Tuple

from .api_conventions import Context, Violation
from .api_conventions_bindings import (
    BINDING_SURFACES,
    TsDecl,
    _callables,
    _entry_surfaces,
    _field,
    _in_tree,
    _key,
    _lookup,
    _norm,
    _owned,
    _parse_members,
    _split_top,
    _violation,
    surfaces,
)
from .api_conventions_shapes import _UNCHANGED_DEFAULTS, _public_pool, _strip_null

PATCH_RECORD_NONE = "none"
_A1_NAMES = ("a1", "anchor")
_PY_RECT = frozenset({"CellRect", "str"})
_JS_RECT = frozenset({"CellRect", "string"})

# Collection index arguments R26.ref replaces, by the reference argument and type that replace them.
_JS_REFS = {
    "tableIndex": ("table", "TableRef"),
    "pivotIndex": ("pivot", "PivotRef"),
    "cfRuleIndex": ("cfRule", "CfRuleRef"),
    "protectedRangeIndex": ("protectedRange", "ProtectedRangeRef"),
}
_PY_REFS = {
    "table_index": "table",
    "pivot_index": "pivot",
    "cf_rule_index": "cf_rule",
    "protected_range_index": "protected_range",
}
_JS_SHEET_PARAM_RE = re.compile(r"sheet|sheetIndex|localSheetId|scope|[a-z]\w*Sheet")
_PY_SHEET_PARAM_RE = re.compile(r"sheet|sheet_index|local_sheet_id|scope|\w+_sheet")
_JS_COUNT_AT_RE = re.compile(r"get[A-Z]\w*(?:Count|At)")
_PY_COUNT_AT_RE = re.compile(r"get_\w+_(?:count|at)")


def _ledger_ops(ctx: Context, field_name: str) -> Iterator[Tuple[str, dict]]:
    for key, entry in sorted(ctx.ledger.entries.items()):
        if isinstance(entry, dict) and entry.get("kind") == "op" and field_name in entry:
            yield key, entry


def _on(entry: dict, surface: str) -> bool:
    listed = _entry_surfaces(entry)
    return listed is None or surface in listed


def _pairs(entry: dict) -> List[Tuple[int, int]]:
    return [(int(i), int(j)) for i, j in entry.get("cell_args") or []]


def _collapse(params: Sequence[Tuple[str, str]], pairs: Sequence[Tuple[int, int]]) -> Optional[List[int]]:
    """Indices of `params` kept by the A1 form, the first of each pair standing for the pair; None when out of range."""
    if any(j >= len(params) for _, j in pairs):
        return None
    dropped = {j for _, j in pairs}
    return [i for i in range(len(params)) if i not in dropped]


# ---------------------------------------------------------------------------
# R25.a1 / R25.rect: cells and rectangles also take A1 text.
# ---------------------------------------------------------------------------


def _js_a1_problems(overloads: Sequence, pairs: Sequence[Tuple[int, int]]) -> Iterator[Tuple[str, str]]:
    first = overloads[0]
    kept = _collapse(first.params, pairs)
    if kept is None:
        yield "a1", f"`{first.name}` has fewer arguments than its `cell_args` name"
        return
    firsts = {i for i, _ in pairs}
    want = ["string" if i in firsts else _norm(first.params[i][1]) for i in kept]
    a1 = [m for m in overloads[1:] if [_norm(t) for _, t in m.params] == want]
    if not a1:
        yield "a1", f"`{first.name}` lacks the overload taking each cell as one A1 string ({', '.join(want)})"
    if len(overloads) - 1 > len(a1[:1]):
        yield "mixed", f"`{first.name}` declares an overload other than the coordinate and the A1 form"


def _py_a1_problems(name: str, stubs: Sequence, pairs: Sequence[Tuple[int, int]]) -> Iterator[Tuple[str, str]]:
    if not stubs:
        yield "stubs", f"`{name}` has no `@overload` stubs for the coordinate and the A1 form"
        return
    coordinate = [(p, t) for p, t in stubs[0].params if p not in stubs[0].kwonly]
    kept = _collapse(coordinate, pairs)
    if kept is None or any(coordinate[k][1] != "int" for pair in pairs for k in pair):
        yield "coordinate", f"`{name}`: the first stub is the coordinate form, with `int` at the `cell_args` positions"
        return
    firsts = [i for i, _ in pairs]
    want = [(_A1_NAMES[firsts.index(i)], "str") if i in firsts else (coordinate[i][0], None) for i in kept]
    second = [(p, t) for p, t in stubs[1].params if p not in stubs[1].kwonly] if len(stubs) > 1 else []
    matches = len(second) == len(want) and all(
        p == wp and (wt is None or t == wt) for (p, t), (wp, wt) in zip(second, want)
    )
    if not matches:
        shown = ", ".join(p for p, _ in want)
        yield "a1", f"`{name}`: the second stub is the A1 form `({shown})`"
    if len(stubs) > 2:
        yield "mixed", f"`{name}` has a stub other than the coordinate and the A1 form"


@_owned("R25.a1")
def _rule_cell_a1(ctx: Context) -> Iterable[Violation]:
    """An op with `cell_args` takes each cell as a `(row, col)` pair or as one A1 string, never a mix: the JS
    declaration has the coordinate signature and the signature with every pair replaced by a `string`; the Python
    method has the same two `@overload` stubs, the A1 one naming the cells `a1` (then `anchor`).
    """
    found = surfaces(ctx)
    for key, entry in _ledger_ops(ctx, "cell_args"):
        pairs = _pairs(entry)
        for surface in BINDING_SURFACES:
            if not _on(entry, surface):
                continue
            name = _in_tree(_callables(ctx, surface), entry, _field(surface))
            if name is None:
                continue
            if surface == "py":
                problems = _py_a1_problems(name, found.py.overloads(name), pairs)
            else:
                problems = _js_a1_problems(
                    found.js[surface].overloads(name) or (_callables(ctx, surface)[name],), pairs
                )
            for disc, message in problems:
                yield _violation(ctx, "R25.a1", surface, key, disc, message)


def _element(alternatives: List[str], js: bool) -> List[str]:
    """The element type of a single array alternative, else the alternatives unchanged."""
    if len(alternatives) != 1:
        return alternatives
    text = alternatives[0]
    if js:
        array = re.fullmatch(r"(?:readonly\s+)?\(?(.*?)\)?\[\]|ReadonlyArray<(.*)>", text)
        return _strip_null(next(g for g in array.groups() if g is not None), True) if array else alternatives
    array = re.fullmatch(r"(?:typing\.)?(?:Sequence|List|Iterable|Tuple)\[(.*?)(?:,\s*\.\.\.)?\]", text)
    return _py_union(array.group(1)) if array else _py_union(text)


def _py_union(text: str) -> List[str]:
    text = text.strip().strip("'\"")
    union = re.fullmatch(r"(?:typing\.)?Union\[(.*)\]", text)
    if union:
        return [a.strip("'\"") for a in _split_top(union.group(1), ",")]
    return [a.strip("'\"") for a in _split_top(text, "|")]


@_owned("R25.rect")
def _rule_rect_a1(ctx: Context) -> Iterable[Violation]:
    """An op's `rect_args` argument takes a rectangle or an A1 range: `CellRect | string` in JS,
    `Union[CellRect, str]` in Python, as the element type when the argument is an array.
    """
    for key, entry in _ledger_ops(ctx, "rect_args"):
        for surface in BINDING_SURFACES:
            if not _on(entry, surface):
                continue
            pool = _callables(ctx, surface)
            name = _in_tree(pool, entry, _field(surface))
            if name is None:
                continue
            js = surface != "py"
            params = [(p, t) for p, t in pool[name].params if js or p not in pool[name].kwonly]
            for position in entry.get("rect_args") or []:
                ptype = params[position][1] if position < len(params) else ""
                kinds = set(_element(_strip_null(ptype, js), js))
                if kinds != (_JS_RECT if js else _PY_RECT):
                    want = "CellRect | string" if js else "Union[CellRect, str]"
                    yield _violation(
                        ctx,
                        "R25.rect",
                        surface,
                        key,
                        f"arg{position}",
                        f"`{name}` argument {position} is `{ptype}`, not `{want}`",
                    )


# ---------------------------------------------------------------------------
# R25.count: no count or position API outside an exempt collection.
# ---------------------------------------------------------------------------


def _exempt_ops(ctx: Context) -> Set[str]:
    out: Set[str] = set()
    for coll in ctx.ledger.collections.values():
        if isinstance(coll, dict) and coll.get("access") == "exempt":
            out.update(op for op in coll.get("ops") or [] if isinstance(op, str))
    return out


@_owned("R25.count")
def _rule_count_at(ctx: Context) -> Iterable[Violation]:
    """No public method is named `get<X>Count` / `get<X>At` (`get_<x>_count` / `get_<x>_at`) unless its op belongs
    to an `access: exempt` collection; a collection is read through `list<Plural>()`.
    """
    exempt = _exempt_ops(ctx)
    for surface in BINDING_SURFACES:
        pattern = _PY_COUNT_AT_RE if surface == "py" else _JS_COUNT_AT_RE
        for name in sorted(_public_pool(ctx, surface)):
            if not pattern.fullmatch(name.rpartition(".")[2]):
                continue
            key, _ = _lookup(ctx, surface, name)
            if key in exempt:
                continue
            yield _violation(
                ctx,
                "R25.count",
                surface,
                key or name,
                name,
                f"`{name}` counts or positions; read the collection as a list",
            )


# ---------------------------------------------------------------------------
# R18.all: a patch's keys are fields of the getter's record.
# ---------------------------------------------------------------------------


_TS_EXTENDS_RE = re.compile(r"^export\s+interface\s+([A-Za-z_$][\w$]*)\s+extends\s+([^{]+)\{", re.M)


def _js_fields(decls: Dict[str, TsDecl], raw: str, name: str, seen: Tuple[str, ...] = ()) -> Optional[Dict[str, str]]:
    """Field name -> type of interface `name`, inherited fields included; None when it is not an interface."""
    decl = decls.get(name)
    if decl is None or decl.kind != "interface":
        return None
    out: Dict[str, str] = {}
    for match in _TS_EXTENDS_RE.finditer(raw):
        if match.group(1) == name:
            for parent in (p.strip() for p in match.group(2).split(",")):
                if parent not in seen:
                    out.update(_js_fields(decls, raw, parent, seen + (name,)) or {})
    out.update({m.name: m.type for m in decl.members if not m.is_method})
    return out


def _patch_keys(decls: Dict[str, TsDecl], raw: str, text: str) -> Optional[Dict[str, str]]:
    """Key -> type of a patch type expression: an interface, `Patch<T>`, `Pick<T, 'a' | 'b'>` or `{ ... }`."""
    text = text.strip()
    inner = re.fullmatch(r"(?:Patch|Partial)<(.*)>", text, re.S)
    if inner:
        return _patch_keys(decls, raw, inner.group(1))
    pick = re.fullmatch(r"Pick<\s*([A-Za-z_$][\w$]*)\s*,(.*)>", text, re.S)
    if pick:
        fields = _patch_keys(decls, raw, pick.group(1))
        keys = re.findall(r"'([^']*)'|\"([^\"]*)\"", pick.group(2))
        return None if fields is None else {k: fields[k] for k in (a or b for a, b in keys) if k in fields}
    if text.startswith("{"):
        return {m.name: m.type for m in _parse_members(text.strip("{} \n")) if not m.is_method}
    decl = decls.get(text)
    if decl is not None and decl.kind == "type":
        return _patch_keys(decls, raw, decl.text)
    return _js_fields(decls, raw, text)


def _same_type(patch_type: str, record_type: str) -> bool:
    """The patch key's type is the record field's, `null` aside, or the patch of a nested record."""
    patch = sorted(_strip_null(patch_type, True))
    record = sorted(_strip_null(record_type, True))
    return patch == record or (len(record) == 1 and patch == [f"{record[0]}Patch"])


def _js_patch_problems(decls: Dict[str, TsDecl], raw: str, member, record: Optional[str]) -> Iterator[Tuple[str, str]]:
    last = member.params[-1] if member.params else ("", "")
    kinds = _strip_null(last[1], True)
    if len(kinds) != 1 or not kinds[0].endswith("Patch") or kinds[0] not in decls:
        yield "patch", f"`{member.name}` takes `{last[1] or 'nothing'}` last, not a declared `<X>Patch`"
        return
    keys = _patch_keys(decls, raw, kinds[0])
    if keys is None:
        yield "patch", f"`{kinds[0]}` is not a record or `Patch<...>` the checker can read"
        return
    if record is None:
        return
    fields = _js_fields(decls, raw, record)
    if fields is None:
        yield "record", f"`patch_record` `{record}` is not a declared interface"
        return
    for key, ptype in sorted(keys.items()):
        if key not in fields:
            yield f"key:{key}", f"`{kinds[0]}.{key}` is not a field of `{record}`"
        elif not _same_type(ptype, fields[key]):
            yield f"type:{key}", f"`{kinds[0]}.{key}: {ptype}` differs from `{record}.{key}: {fields[key]}`"


def _py_patch_problems(fn, name: str, fields: Optional[Set[str]]) -> Iterator[Tuple[str, str]]:
    if not fn.kwonly:
        yield "patch", f"`{name}` takes no keyword-only patch arguments"
        return
    defaults = dict(fn.defaults)
    for param in fn.kwonly:
        if defaults.get(param) not in _UNCHANGED_DEFAULTS:
            yield f"default:{param}", f"`{name}({param})` defaults to `{defaults.get(param)}`, not `UNCHANGED`"
        if fields is not None and param not in fields:
            yield f"key:{param}", f"`{name}({param})` is not a field of the record"
    for param, _ in fn.params:
        if param not in fn.kwonly and fields is not None and param in fields:
            yield f"positional:{param}", f"`{name}({param})` is a patch key; make it keyword-only"


@_owned("R18.all")
def _rule_patch_shape(ctx: Context) -> Iterable[Violation]:
    """An op with `patch: true` takes a declared `<X>Patch` last in JS, whose keys are fields of the
    `patch_record` interface with the same types (`null` added, a nested record as its patch); in Python it takes
    keyword-only arguments defaulting to `UNCHANGED`, named after fields of the `patch_record` dataclass. A
    `patch_record: none` op has no getter record, so its keys are not compared.
    """
    found = surfaces(ctx)
    for key, entry in _ledger_ops(ctx, "patch"):
        if entry.get("patch") is not True:
            continue
        record = entry.get("patch_record")
        record = None if record == PATCH_RECORD_NONE or not isinstance(record, str) else record
        for surface in BINDING_SURFACES:
            if not _on(entry, surface):
                continue
            pool = _callables(ctx, surface)
            name = _in_tree(pool, entry, _field(surface))
            if name is None:
                continue
            if surface == "py":
                cls = found.py.classes.get(record) if record else None
                if record and cls is None:
                    yield _violation(ctx, "R18.all", "py", key, "record", f"`patch_record` `{record}` is not a class")
                    continue
                problems = _py_patch_problems(pool[name], name, set(cls.fields) if cls else None)
            else:
                js = found.js[surface]
                problems = _js_patch_problems(js.decls, js.raw, pool[name], record)
            for disc, message in problems:
                yield _violation(ctx, "R18.all", surface, key, disc, message)


# ---------------------------------------------------------------------------
# R26.ref / R27.sheetparam: references by index or name.
# ---------------------------------------------------------------------------


def _alternatives(ptype: str, js: bool) -> List[str]:
    return _strip_null(ptype, True) if js else _py_union(" | ".join(_strip_null(ptype, False)))


@_owned("R26.ref")
def _rule_refs(ctx: Context) -> Iterable[Violation]:
    """A table, pivot, CF rule or protected range argument is a reference by index or name: `table: TableRef`,
    `pivot: PivotRef`, `cfRule: CfRuleRef`, `protectedRange: ProtectedRangeRef` in JS, `Union[int, str]` under
    the snake-case name in Python; never `<thing>Index`.
    """
    for surface in BINDING_SURFACES:
        js = surface != "py"
        for name, member in sorted(_public_pool(ctx, surface).items()):
            for param, ptype in member.params:
                if js and param in _JS_REFS:
                    want, ref = _JS_REFS[param]
                    message = f"`{name}({param})` addresses by position; take `{want}: {ref}`"
                elif not js and param in _PY_REFS:
                    message = f"`{name}({param})` addresses by position; take `{_PY_REFS[param]}: Union[int, str]`"
                elif js and param in {v[0] for v in _JS_REFS.values()}:
                    ref = next(r for p, r in _JS_REFS.values() if p == param)
                    if _alternatives(ptype, True) == [ref]:
                        continue
                    message = f"`{name}({param}: {ptype})` is not `{ref}`"
                elif not js and param in _PY_REFS.values():
                    if set(_alternatives(ptype, False)) == {"int", "str"}:
                        continue
                    message = f"`{name}({param}: {ptype})` is not `Union[int, str]`"
                else:
                    continue
                yield _violation(ctx, "R26.ref", surface, _key(ctx, surface, name), param, message)


@_owned("R27.sheetparam")
def _rule_sheet_params(ctx: Context) -> Iterable[Violation]:
    """A method argument that names a sheet (`sheet`, `sheetIndex`, `localSheetId`, `scope`, `<x>Sheet`, and the
    snake-case forms) is a `SheetRef`. Record fields are not judged.
    """
    for surface in BINDING_SURFACES:
        js = surface != "py"
        pattern = _JS_SHEET_PARAM_RE if js else _PY_SHEET_PARAM_RE
        for name, member in sorted(_public_pool(ctx, surface).items()):
            for param, ptype in member.params:
                if pattern.fullmatch(param) and _alternatives(ptype, js) != ["SheetRef"]:
                    yield _violation(
                        ctx,
                        "R27.sheetparam",
                        surface,
                        _key(ctx, surface, name),
                        param,
                        f"`{name}({param}: {ptype})` names a sheet; type it `SheetRef`",
                    )
