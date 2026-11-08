"""Ledger-bound rules on how the binding methods take their arguments.

A cell is a `(row, col)` pair or one A1 string, a rectangle a `CellRect` or an
A1 range string (R25.a1 / R25.rect, driven by the ledger's `cell_args` and
`rect_args`), a row or column span a number pair or one span string (R25.span,
driven by `span_args`). A collection is enumerated by `list<Plural>()`, not
counted (R25.count). A partial update takes a declared patch whose keys are
fields of the getter's record, and every patch has one (R18.all, driven by
`patch` / `patch_record`). Tables, pivots, CF rules and protected ranges are
addressed by index or name (R26.ref), sheets by `SheetRef` (R27.sheetparam).
The element of a collection published as handles is edited through its
handle; its parent keeps add, remove, the lookup and the list (R28.handles).

A member's signature is its first overload (the coordinate form); R25.a1
alone reads every overload.
"""

from __future__ import annotations

import re
from typing import Dict, Iterable, Iterator, List, NamedTuple, Optional, Sequence, Set, Tuple

from .api_conventions import (
    SHEET_HANDLE_CLASS,
    Context,
    Violation,
    c_verb,
    entry_handle,
    handle_classes,
    handle_collections,
    handle_name,
)
from .api_conventions_bindings import (
    BINDING_SURFACES,
    SHEET_CLASS,
    WORKBOOK_CLASS,
    TsDecl,
    _callables,
    _dropped,
    _entry_surfaces,
    _field,
    _in_tree,
    _key,
    _lookup,
    _norm,
    _owned,
    _parse_members,
    _split_top,
    _top_union,
    _tree_c,
    _violation,
    surfaces,
)
from .api_conventions_shapes import _UNCHANGED_DEFAULTS, _public_pool, _strip_null

PATCH_RECORD_NONE = "none"
_A1_NAMES = ("a1", "anchor")
_PY_RECT = frozenset({"CellRect", "str"})
_JS_RECT = frozenset({"CellRect", "string"})
_RECT_RECORDS = frozenset({"CellRect", "SheetRect"})
_JS_UPDATE_RE = re.compile(r"update(?:[A-Z]\w*)?")
_PY_UPDATE_RE = re.compile(r"update(?:_\w+)?")

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
# Python reference arguments that also take the handle of the object, as the JS `TableRef` / `PivotRef` do.
_PY_REF_HANDLES = {"table": ("Table",), "pivot": ("Pivot",)}


def _py_ref_union(param: str) -> str:
    return f"Union[{', '.join(('int', 'str', *_PY_REF_HANDLES.get(param, ())))}]"


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


class _Target(NamedTuple):
    """Where an op is declared on one surface: its label, its first signature, and every JS signature or the
    Python `@overload` stubs."""

    name: str
    member: object
    overloads: Tuple[object, ...]


def _targets(ctx: Context, entry: dict, surface: str) -> List[_Target]:
    """The declarations of the op on `surface`: each handle class's method for an op with `handle:`, else its
    `Workbook` / `Sheet` method or free function. An absent declaration yields nothing."""
    found = surfaces(ctx)
    if entry_handle(entry) is None:
        pool = _callables(ctx, surface)
        name = _in_tree(pool, entry, _field(surface))
        if name is None:
            return []
        if surface == "py":
            return [_Target(name, pool[name], tuple(found.py.overloads(name)))]
        return [_Target(name, pool[name], tuple(found.js[surface].overloads(name)) or (pool[name],))]
    name = handle_name(entry, _field(surface), ctx.ledger.collections)
    out: List[_Target] = []
    for cls in handle_classes(entry):
        if surface == "py":
            py_class = found.py.classes.get(cls)
            member = py_class.methods.get(name) if py_class and name else None
            overloads = tuple(py_class.overloads.get(name, ())) if py_class and name else ()
        else:
            decl = found.js[surface].decls.get(cls)
            overloads = tuple(m for m in decl.members if m.name == name and m.is_method) if decl else ()
            member = overloads[0] if overloads else None
        if member is not None:
            out.append(_Target(f"{cls}.{name}", member, overloads))
    return out


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
    for key, entry in _ledger_ops(ctx, "cell_args"):
        pairs = _pairs(entry)
        for surface in BINDING_SURFACES:
            if not _on(entry, surface):
                continue
            for target in _targets(ctx, entry, surface):
                if surface == "py":
                    problems = _py_a1_problems(target.name, target.overloads, pairs)
                else:
                    problems = _js_a1_problems(target.overloads, pairs)
                for disc, message in problems:
                    yield _violation(ctx, "R25.a1", surface, key, disc, message)


def _element(alternatives: List[str], js: bool) -> List[str]:
    """The element type of a single array alternative, or of an array that also takes one A1 string in its place;
    else the alternatives unchanged."""
    a1 = "string" if js else "str"
    if len(alternatives) == 2 and a1 in alternatives:
        array = [a for a in alternatives if a != a1]
        element = _element(array, js)
        if element != array:
            return element
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
    `Union[CellRect, str]` in Python, as the element type when the argument is an array (which may also take
    one A1 string in place of the array).
    """
    for key, entry in _ledger_ops(ctx, "rect_args"):
        for surface in BINDING_SURFACES:
            if not _on(entry, surface):
                continue
            js = surface != "py"
            for name, member, _overloads in _targets(ctx, entry, surface):
                params = [(p, t) for p, t in member.params if js or p not in member.kwonly]
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


_TS_EXTENDS_RE = re.compile(r"^(?:export\s+)?(?:declare\s+)?interface\s+([A-Za-z_$][\w$]*)\s+extends\s+([^{]+)\{", re.M)


def _js_fields(decls: Dict[str, TsDecl], raw: str, name: str, seen: Tuple[str, ...] = ()) -> Optional[Dict[str, str]]:
    """Field name -> type of interface `name`, inherited fields included (from any interface in `decls`, which
    may hold module-private ones); None when it is not an interface."""
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


def _same_type(patch_type: str, record_type: str, decls: Optional[Dict[str, TsDecl]] = None) -> bool:
    """The patch key's type is the record field's, `null` aside, the patch of a nested record, a rectangle
    field that also takes its A1 string, or the record type's declared input alias (`ColorInput = string | Color`
    for a `Color` field)."""
    patch = sorted(_strip_null(patch_type, True))
    record = sorted(_strip_null(record_type, True))
    if len(record) == 1 and record[0] in _RECT_RECORDS and patch == sorted([record[0], "string"]):
        return True
    if len(record) == 1 and patch == [f"{record[0]}Input"]:
        alias = (decls or {}).get(patch[0])
        if alias is not None and alias.kind == "type" and record[0] in _top_union(alias.text):
            return True
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
        elif not _same_type(ptype, fields[key], decls):
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
    """An op with `patch: true` is named `update` / `update<X>` on every binding surface (a `set` replaces a whole
    value and never merges) and takes a declared `<X>Patch` last in JS, whose keys are fields of the
    `patch_record` interface with the same types (`null` added, a nested record as its patch); in Python it takes
    keyword-only arguments defaulting to `UNCHANGED`, named after fields of the `patch_record` dataclass. Every
    patch reads back through a getter: `patch_record: none` is a violation. An op named `update` / `update<X>` is
    always a patch. A rectangle field may take its A1 string as well in the patch.
    """
    for key, entry in sorted(ctx.ledger.entries.items()):
        if not isinstance(entry, dict) or entry.get("kind") != "op" or entry.get("internal") is True:
            continue
        if entry.get("patch") is True:
            continue
        js, py = entry.get("js"), entry.get("py")
        if (isinstance(js, str) and _JS_UPDATE_RE.fullmatch(js)) or (
            isinstance(py, str) and _PY_UPDATE_RE.fullmatch(py)
        ):
            yield _violation(
                ctx, "R18.all", "c", key, "update", f"`{js or py}` is an update; it is a partial update, `patch: true`"
            )
    found = surfaces(ctx)
    for key, entry in _ledger_ops(ctx, "patch"):
        if entry.get("patch") is not True:
            continue
        if entry.get("patch_record") == PATCH_RECORD_NONE:
            yield _violation(
                ctx, "R18.all", "c", key, "patch_record", "a patch reads back through a getter; `patch_record` is none"
            )
        record = entry.get("patch_record")
        record = None if record == PATCH_RECORD_NONE or not isinstance(record, str) else record
        for surface in BINDING_SURFACES:
            if not _on(entry, surface):
                continue
            for name, member, _overloads in _targets(ctx, entry, surface):
                method = name.rpartition(".")[2]
                if not (_PY_UPDATE_RE if surface == "py" else _JS_UPDATE_RE).fullmatch(method):
                    yield _violation(
                        ctx, "R18.all", surface, key, "verb", f"`{name}` merges a patch; name it with `update`"
                    )
                if surface == "py":
                    cls = found.py.classes.get(record) if record else None
                    if record and cls is None:
                        yield _violation(
                            ctx, "R18.all", "py", key, "record", f"`patch_record` `{record}` is not a class"
                        )
                        continue
                    problems = _py_patch_problems(member, name, set(cls.fields) if cls else None)
                else:
                    js = found.js[surface]
                    problems = _js_patch_problems(js.decls, js.raw, member, record)
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
    the snake-case name in Python (plus the handle class for a table or pivot, as `TableRef` / `PivotRef` carry
    it); never `<thing>Index`.
    """
    for surface in BINDING_SURFACES:
        js = surface != "py"
        for name, member in sorted(_public_pool(ctx, surface).items()):
            for param, ptype in member.params:
                if js and param in _JS_REFS:
                    want, ref = _JS_REFS[param]
                    message = f"`{name}({param})` addresses by position; take `{want}: {ref}`"
                elif not js and param in _PY_REFS:
                    want = _py_ref_union(_PY_REFS[param])
                    message = f"`{name}({param})` addresses by position; take `{_PY_REFS[param]}: {want}`"
                elif js and param in {v[0] for v in _JS_REFS.values()}:
                    ref = next(r for p, r in _JS_REFS.values() if p == param)
                    if _alternatives(ptype, True) == [ref]:
                        continue
                    message = f"`{name}({param}: {ptype})` is not `{ref}`"
                elif not js and param in _PY_REFS.values():
                    want = _py_ref_union(param)
                    if set(_alternatives(ptype, False)) == {"int", "str", *_PY_REF_HANDLES.get(param, ())}:
                        continue
                    message = f"`{name}({param}: {ptype})` is not `{want}`"
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


# ---------------------------------------------------------------------------
# R25.span: a row or column span is a number pair or one span string.
# ---------------------------------------------------------------------------


def _js_span_problems(overloads: Sequence, spans: Sequence[Sequence]) -> Iterator[Tuple[str, str]]:
    first = overloads[0]
    starts = [int(span[0]) for span in spans]
    if any(i + 1 >= len(first.params) for i in starts):
        yield "span", f"`{first.name}` has fewer arguments than its `span_args` name"
        return
    want = ["string" if i in starts else _norm(t) for i, (_, t) in enumerate(first.params) if i - 1 not in starts]
    if not any([_norm(t) for _, t in m.params] == want for m in overloads[1:]):
        yield "span", f"`{first.name}` lacks the overload taking the span as one string ({', '.join(want)})"


def _py_span_problems(name: str, stubs: Sequence, spans: Sequence[Sequence]) -> Iterator[Tuple[str, str]]:
    if len(stubs) < 2:
        yield "stubs", f"`{name}` has no `@overload` stubs for the number pair and the span string"
        return
    starts = [int(span[0]) for span in spans]
    numbers = [(p, t) for p, t in stubs[0].params if p not in stubs[0].kwonly]
    if any(i + 1 >= len(numbers) or numbers[i][1] != "int" or numbers[i + 1][1] != "int" for i in starts):
        yield "coordinate", f"`{name}`: the first stub takes `int` at the `span_args` positions"
        return
    want = [("span", "str") if i in starts else (p, None) for i, (p, _) in enumerate(numbers) if i - 1 not in starts]
    second = [(p, t) for p, t in stubs[1].params if p not in stubs[1].kwonly]
    if not (
        len(second) == len(want) and all(p == wp and (wt is None or t == wt) for (p, t), (wp, wt) in zip(second, want))
    ):
        shown = ", ".join(p for p, _ in want)
        yield "span", f"`{name}`: the second stub is the span form `({shown})`"


@_owned("R25.span")
def _rule_span(ctx: Context) -> Iterable[Violation]:
    """An op with `span_args` takes each span as its number pair or as one span string (`"B:D"`, `"3:5"`): the JS
    declaration has the number signature first and the signature with the pair replaced by a `string`; the Python
    method has the same two `@overload` stubs, the span one naming the argument `span`.
    """
    for key, entry in _ledger_ops(ctx, "span_args"):
        spans = entry.get("span_args") or []
        for surface in BINDING_SURFACES:
            if not _on(entry, surface):
                continue
            for target in _targets(ctx, entry, surface):
                if surface == "py":
                    problems = _py_span_problems(target.name, target.overloads, spans)
                else:
                    problems = _js_span_problems(target.overloads, spans)
                for disc, message in problems:
                    yield _violation(ctx, "R25.span", surface, key, disc, message)


# ---------------------------------------------------------------------------
# R28.handles: the element of a handle collection is edited through its handle.
# ---------------------------------------------------------------------------

# Verbs of the C ops a parent keeps for a handle collection: it adds and removes elements.
_PARENT_VERBS = frozenset({"add", "create", "copy", "remove"})
_FIRST_WORD_RE = re.compile(r"[a-z]+")
_FIRST_ARG_KEYS = ("object_id", "cache_id")


def _exposed(entry: dict) -> bool:
    """The op has a binding name a published class could carry: not internal, not dropped on every language."""
    if entry.get("internal") is True or entry.get("kind") not in ("op", "list"):
        return False
    names = [(field, entry.get(field)) for field in ("js", "py") if isinstance(entry.get(field), str)]
    return any(_dropped(entry, field) != name for field, name in names) or entry_handle(entry) is not None


def handle_collection_of(ctx: Context, key: str, entry: dict) -> Optional[str]:
    """The handle collection an op belongs to: the one whose `ops` lists it, else the one its first argument
    keys (`ref_args[0]`'s kind, or a C argument `object_id` / `cache_id` named by the collection's `by:`)."""
    handles = handle_collections(ctx.ledger.collections)
    for name in handles:
        if key in ((ctx.ledger.collections.get(name) or {}).get("ops") or []):
            return name
    for name, coll in ctx.ledger.collections.items():
        if key in (coll.get("ops") or []):
            return None
    ref = entry.get("ref_args")
    if isinstance(ref, list) and len(ref) == 2 and ref[0] == 0 and ref[1] in handles:
        return ref[1]
    fn = _tree_c(ctx, entry)
    rest = [p.name for p in (fn.params if fn else ()) if p.name not in ("wb", "sheet_index")]
    if not rest or rest[0] not in _FIRST_ARG_KEYS:
        return None
    keyed = sorted(n for n in handles if (ctx.ledger.collections[n] or {}).get("by") == rest[0])
    noun = key.split(".")[0]
    prefixed = [n for n in keyed if noun == n or noun.startswith(n + "_")]
    return (prefixed or keyed or [None])[0]


def _parent_op(ctx: Context, key: str, entry: dict, collection: str) -> bool:
    """An op the parent keeps: in the collection's `ops`, it adds or removes, lists, or looks the handle up."""
    if key not in ((ctx.ledger.collections.get(collection) or {}).get("ops") or []):
        return False
    if entry.get("kind") == "list":
        return True
    target = entry.get("c")
    if isinstance(target, str):
        return c_verb(target) in _PARENT_VERBS
    name = entry.get("js") if isinstance(entry.get("js"), str) else entry.get("py")
    verb = _FIRST_WORD_RE.match(name or "")
    lookup = name in (_lower_first(handle_collections(ctx.ledger.collections)[collection]), collection)
    return lookup or (verb is not None and verb.group(0) in _PARENT_VERBS)


def _lower_first(name: str) -> str:
    return name[:1].lower() + name[1:]


@_owned("R28.handles")
def _rule_handles(ctx: Context) -> Iterable[Violation]:
    """An op of a collection with `access: handles` and a `handle_class` -- listed in its `ops`, else keyed by its
    first argument -- has `handle:` unless the parent keeps it (add / remove / the handle lookup / the list); an
    op with `handle:` is declared on no `Workbook` or `Sheet`.
    """
    for key, entry in sorted(ctx.ledger.entries.items()):
        if not isinstance(entry, dict) or not _exposed(entry):
            continue
        collection = handle_collection_of(ctx, key, entry)
        if collection is None or entry_handle(entry) is not None or _parent_op(ctx, key, entry, collection):
            continue
        handle_class = handle_collections(ctx.ledger.collections)[collection]
        yield _violation(
            ctx,
            "R28.handles",
            "c",
            key,
            "handle",
            f"`{key}` edits a `{collection}` element; publish it on `{handle_class}` with `handle:`",
        )
    found = surfaces(ctx)
    for key, entry in sorted(ctx.ledger.entries.items()):
        if not isinstance(entry, dict) or entry_handle(entry) is None:
            continue
        for surface in BINDING_SURFACES:
            if not _on(entry, surface):
                continue
            field = _field(surface)
            classes = found.py.class_methods() if surface == "py" else found.js[surface].classes()
            for owner in (WORKBOOK_CLASS, SHEET_CLASS):
                name = _in_tree(classes.get(owner) or {}, entry, field)
                if name is None or (owner == SHEET_CLASS and SHEET_HANDLE_CLASS in _handle_class_list(entry)):
                    continue
                yield _violation(
                    ctx,
                    "R28.handles",
                    surface,
                    key,
                    owner,
                    f"`{owner}.{name}` is a handle method; only the handle class declares it",
                )


def _handle_class_list(entry: dict) -> List[str]:
    cls = (entry_handle(entry) or {}).get("class")
    return [cls] if isinstance(cls, str) else list(cls or [])
