"""Ledger-bound rules on the shape of the binding surfaces.

These rules judge what a caller of the JS and Python bindings sees beyond the
names: which class (`Workbook` or `Sheet`) owns an op, how an index-keyed
collection is exposed, when arguments fold into an options object, how a
setter mirrors its getter, how "absent" is spelled, which values are closed
sets, how same-concept enums agree, and how records and index arguments are
named. They read the same parsed surfaces as `api_conventions_bindings`, plus
the generated facade tables of `packages/npm/common.mjs` and
`packages/python/formulon/sheet.py`.
"""

from __future__ import annotations

import itertools
import re
from typing import Dict, Iterable, Iterator, List, Optional, Sequence, Set, Tuple

from .api_conventions import (
    _COMMENT_RE,
    _HEADER_DEFINE_RE,
    MEANINGLESS_SUFFIXES,
    Context,
    Violation,
    _enum_family,
    _struct_field_comments,
    c_verb,
    strip_record_suffix,
)
from .api_conventions_bindings import (
    BINDING_SURFACES,
    JS_SURFACES,
    SHEET_CLASS,
    WORKBOOK_CLASS,
    PyFunc,
    _callables,
    _camel,
    _dropped,
    _entry_surfaces,
    _field,
    _in_tree,
    _key,
    _optional_lookups,
    _owned,
    _pascal,
    _split_top,
    _top_union,
    _tree_c,
    _violation,
    surfaces,
)
from .surface_files import NPM_COMMON_MJS, PYTHON_SHEET, REPO_ROOT

# Arguments that address something rather than carry a value; they stay positional (R7.opts) and are not the
# value argument of a setter (R18).
_JS_COORDINATES = frozenset(
    {"sheet", "row", "col", "first", "last", "firstRow", "lastRow", "firstCol", "lastCol", "rect"}
)
_PY_COORDINATES = frozenset(
    {"sheet", "row", "col", "first", "last", "first_row", "last_row", "first_col", "last_col", "rect"}
)


def _is_coordinate(name: str, js: bool) -> bool:
    if name in (_JS_COORDINATES if js else _PY_COORDINATES):
        return True
    return name.endswith(("Index", "Id")) if js else name.endswith(("_index", "_id"))


def _value_params(params: Sequence[Tuple[str, str]], js: bool) -> List[Tuple[str, str]]:
    return [(name, ptype) for name, ptype in params if not _is_coordinate(name, js)]


def _strip_null(text: str, js: bool) -> List[str]:
    """Top-level union alternatives without the absent alternative (`null` / `undefined` / `None`)."""
    text = text.strip().strip("'\"")
    if not js:
        inner = re.fullmatch(r"(?:typing\.)?Optional\[(.*)\]", text)
        if inner:
            text = inner.group(1)
        union = re.fullmatch(r"(?:typing\.)?Union\[(.*)\]", text)
        if union:
            return [a for a in _split_commas(union.group(1)) if a not in ("None", "NoneType")]
    return [a.strip("'\"") for a in _top_union(text) if a not in ("null", "undefined", "None")]


def _split_commas(text: str) -> List[str]:
    return _split_top(text, ",")


_PY_BOOL = frozenset({"bool", "Optional[bool]", "bool | None", "None | bool", "typing.Optional[bool]"})


def _is_bool(annotation: str, js: bool) -> bool:
    return annotation == "boolean" if js else annotation.strip("'\"") in _PY_BOOL


# ---------------------------------------------------------------------------
# Class membership: R23.
# ---------------------------------------------------------------------------


def _scoped_entries(ctx: Context, field: str, surface: str) -> Iterator[Tuple[str, dict, str]]:
    """`(key, entry, scope)` of every op / list entry with a `field` name on `surface` and a class scope."""
    for key, entry in sorted(ctx.ledger.entries.items()):
        if not isinstance(entry, dict) or entry.get("kind") not in ("op", "list"):
            continue
        scope = entry.get("scope")
        listed = _entry_surfaces(entry)
        if scope not in ("workbook", "sheet") or (listed is not None and surface not in listed):
            continue
        if isinstance(entry.get(field), str):
            yield key, entry, scope


def _class_problems(
    ctx: Context, surface: str, classes: Dict[str, Dict[str, object]], params_of
) -> Iterator[Tuple[str, str, str]]:
    field = _field(surface)
    union = {name: True for members in classes.values() for name in members}
    for key, entry, scope in _scoped_entries(ctx, field, surface):
        name = _in_tree(union, entry, field)
        if name is None or name == _dropped(entry, field):
            continue
        want = SHEET_CLASS if scope == "sheet" else WORKBOOK_CLASS
        other = WORKBOOK_CLASS if want == SHEET_CLASS else SHEET_CLASS
        if name in classes[other]:
            yield key, other, f"`{name}` is a {scope}-scope op; it belongs on `{want}`, not `{other}`"
        elif name not in classes[want]:
            continue
        if want == SHEET_CLASS and name in classes[SHEET_CLASS]:
            params = params_of(classes[SHEET_CLASS][name])
            if params and params[0] == "sheet":
                yield key, "sheet-arg", f"`{SHEET_CLASS}.{name}` takes the sheet its handle already names"


_MAKE_FACADE_RE = re.compile(r"\bfunction\s+makeFacade\s*\(")


@_owned("R23.js")
def _rule_js_classes(ctx: Context) -> Iterable[Violation]:
    """The JS `Workbook` declares no `scope: sheet` op and `Sheet` no `scope: workbook` op; `Sheet` declares its
    ops without the leading sheet argument. `packages/npm/common.mjs` builds the classes with `makeFacade`.
    """
    for surface in JS_SURFACES:
        classes = surfaces(ctx).js[surface].classes()
        for key, disc, message in _class_problems(ctx, surface, classes, lambda m: [p for p, _ in m.params]):
            yield _violation(ctx, "R23.js", surface, key, disc, message)
    common = ctx.repo_root / NPM_COMMON_MJS.relative_to(REPO_ROOT)
    text = common.read_text(encoding="utf-8") if common.is_file() else ""
    if not _MAKE_FACADE_RE.search(text):
        yield _violation(
            ctx,
            "R23.js",
            "wasm",
            WORKBOOK_CLASS,
            "makeFacade",
            "common.mjs does not build the facade with `makeFacade`",
        )


@_owned("R23.py")
def _rule_py_classes(ctx: Context) -> Iterable[Violation]:
    """The Python `Workbook` has no `scope: sheet` op and `Sheet` no `scope: workbook` op; `Sheet` takes no
    leading sheet argument.
    """
    classes = surfaces(ctx).py.class_methods()
    for key, disc, message in _class_problems(ctx, "py", classes, lambda fn: [p for p, _ in fn.params]):
        yield _violation(ctx, "R23.py", "py", key, disc, message)


_FACADE_REGION = {
    "js": re.compile(r"//\s*<facade-tables>(.*?)//\s*</facade-tables>", re.S),
    "py": re.compile(r"#\s*<facade-tables>(.*?)#\s*</facade-tables>", re.S),
}


def facade_table(text: str, language: str, table: str) -> Optional[Set[str]]:
    """The quoted names of `table` inside the generated facade region of `text`, or None when absent."""
    region = _FACADE_REGION[language].search(text)
    if region is None:
        return None
    found = re.search(r"\b%s\b[^\[({]*[\[({]" % re.escape(table), region.group(1))
    if found is None:
        return None
    body = region.group(1)[found.end() :]
    depth, end = 1, 0
    for end, char in enumerate(body):
        depth += char in "[({"
        depth -= char in "])}"
        if depth == 0:
            break
    return {a or b for a, b in re.findall(r"'([^']*)'|\"([^\"]*)\"", body[:end])}


def _structure_ops(ctx: Context) -> List[Tuple[str, dict]]:
    """Ops that change the number or order of sheets: a non-`get` op returning `out_sheet_index`, and every
    non-`get` op of the sheet collection."""
    sheet_ops = set((ctx.ledger.collections.get("sheet") or {}).get("ops") or [])
    out = []
    for key, entry in sorted(ctx.ledger.entries.items()):
        if not isinstance(entry, dict) or entry.get("kind") != "op":
            continue
        fn = _tree_c(ctx, entry)
        if fn is None or c_verb(ctx.target_c(fn.name)) == "get":
            continue
        returns_sheet = any(p.name == "out_sheet_index" and p.pointer_depth == 1 for p in fn.params)
        if returns_sheet or key in sheet_ops:
            out.append((key, entry))
    return out


@_owned("R23.track")
def _rule_structure_tracking(ctx: Context) -> Iterable[Violation]:
    """Every op that adds, removes or reorders sheets is in `SHEET_STRUCTURE_OPS` of the generated facade
    tables, in `packages/npm/common.mjs` (both npm packages) and `packages/python/formulon/sheet.py`.
    """
    sources = {
        "wasm": ("js", ctx.repo_root / NPM_COMMON_MJS.relative_to(REPO_ROOT)),
        "py": ("py", ctx.repo_root / PYTHON_SHEET.relative_to(REPO_ROOT)),
    }
    for surface, (language, path) in sources.items():
        text = path.read_text(encoding="utf-8") if path.is_file() else ""
        listed = facade_table(text, language, "SHEET_STRUCTURE_OPS") or set()
        field = _field(surface)
        for key, entry in _structure_ops(ctx):
            name = entry.get(field)
            listed_surfaces = _entry_surfaces(entry)
            if not isinstance(name, str) or (listed_surfaces is not None and surface not in listed_surfaces):
                continue
            if name not in listed:
                yield _violation(
                    ctx,
                    "R23.track",
                    surface,
                    key,
                    "SHEET_STRUCTURE_OPS",
                    f"`{name}` changes the sheet list but `{path.name}` does not track it in `SHEET_STRUCTURE_OPS`",
                )


# ---------------------------------------------------------------------------
# R23.list: the binding shape of an index-keyed collection.
# ---------------------------------------------------------------------------

ACCESS_MODES = ("handles", "list", "indexed", "property", "exempt")


def _collection_names(ctx: Context, ops: Sequence[str], field: str) -> Tuple[List[str], List[str]]:
    """Ledger keys of the count op and of the element getters among `ops` (C verb `get`, no `_at`)."""
    counts, getters = [], []
    for op in ops:
        entry = ctx.ledger.entry(op)
        target = entry.get("c")
        if not isinstance(target, str):
            continue
        if c_verb(target) != "get" or target.endswith("_at"):
            continue
        (counts if target.endswith("_count") else getters).append(op)
    return counts, getters


def _tree_names(ctx: Context, keys: Sequence[str], field: str, pool: Dict[str, object]) -> List[Tuple[str, str]]:
    out = []
    for key in keys:
        entry = ctx.ledger.entry(key)
        name = _in_tree(pool, entry, field)
        if name is not None and name != _dropped(entry, field):
            out.append((key, name))
    return out


@_owned("R23.list")
def _rule_collection_access(ctx: Context) -> Iterable[Violation]:
    """An index-keyed collection's binding shape follows its ledger `access:`. `handles` and `list`: only
    `list<Plural>()`, no count and no element getter; `indexed`: the element getter and `list<Plural>()`, no
    count; `property`: `get<Plural>()` and `set<Plural>()`, no count and no element getter; `exempt`: unchecked.
    A surface none of the collection's ops lists is not judged.
    """
    for surface in BINDING_SURFACES:
        field = _field(surface)
        pool = _callables(ctx, surface)
        for name, coll in sorted(ctx.ledger.collections.items()):
            if not isinstance(coll, dict) or coll.get("access") in (None, "exempt"):
                continue
            access, plural = coll["access"], coll.get("plural")
            if not isinstance(plural, str):
                continue
            ops = coll.get("ops") or []
            if not any(surface in (_entry_surfaces(ctx.ledger.entry(op)) or ()) for op in ops):
                continue
            counts, getters = _collection_names(ctx, ops, field)
            present = _tree_names(ctx, counts + getters, field, pool)
            pascal = _pascal(plural)
            if access in ("handles", "list"):
                want = [f"list_{plural}" if field == "py" else f"list{pascal}"]
                forbidden = present
            elif access == "property":
                want = [f"get_{plural}", f"set_{plural}"] if field == "py" else [f"get{pascal}", f"set{pascal}"]
                forbidden = present
            else:
                want = [f"list_{plural}" if field == "py" else f"list{pascal}"]
                forbidden = _tree_names(ctx, counts, field, pool)
                if not _tree_names(ctx, getters, field, pool):
                    yield _violation(
                        ctx,
                        "R23.list",
                        surface,
                        f"collection.{name}",
                        "getter",
                        f"`{name}` is indexed; it exposes the element getter by index",
                    )
            for wanted in want:
                if wanted not in pool:
                    yield _violation(
                        ctx,
                        "R23.list",
                        surface,
                        f"collection.{name}",
                        wanted,
                        f"`{name}` is `access: {access}`; `{wanted}` is missing",
                    )
            for key, tree_name in forbidden:
                yield _violation(
                    ctx,
                    "R23.list",
                    surface,
                    key,
                    tree_name,
                    f"`{name}` is `access: {access}`; `{tree_name}` leaves the {surface} surface",
                )


# ---------------------------------------------------------------------------
# R7.opts: options objects and keyword-only flags.
# ---------------------------------------------------------------------------


def _longest_run(types: Sequence[str]) -> Tuple[int, str]:
    """Length and type of the longest run of adjacent arguments of one type."""
    best, kind, run = 0, "", 0
    previous = None
    for ptype in types:
        run = run + 1 if ptype == previous else 1
        if run > best:
            best, kind = run, ptype
        previous = ptype
    return best, kind


def single_boolean_setter(name: str, params: Sequence[Tuple[str, str]], js: bool) -> bool:
    """A `set<X>` whose only value argument is the one boolean its name points at."""
    values = _value_params(params, js)
    return name.startswith("set") and len(values) == 1 and _is_bool(values[0][1], js)


def _public_pool(ctx: Context, surface: str) -> Dict[str, object]:
    return {n: m for n, m in _callables(ctx, surface).items() if not n.rpartition(".")[2].startswith("_")}


@_owned("R7.opts.js")
def _rule_js_options(ctx: Context) -> Iterable[Violation]:
    """A JS method whose value arguments (coordinates and identifiers excluded) hold two or more booleans, or
    three or more adjacent arguments of one type, takes an options object instead.
    """
    for surface in JS_SURFACES:
        for name, member in sorted(_public_pool(ctx, surface).items()):
            values = _value_params(member.params, True)
            if len(values) < 2:
                continue
            types = [t.rstrip("?") for _, t in values]
            if sum(_is_bool(t, True) for t in types) >= 2:
                yield _violation(
                    ctx, "R7.opts.js", surface, _key(ctx, surface, name), "bool", f"`{name}` takes two booleans"
                )
            run, ptype = _longest_run(types)
            if run >= 3:
                yield _violation(
                    ctx,
                    "R7.opts.js",
                    surface,
                    _key(ctx, surface, name),
                    "run",
                    f"`{name}` takes {run} adjacent `{ptype}` arguments; take an options object",
                )


@_owned("R7.opts.py")
def _rule_py_options(ctx: Context) -> Iterable[Violation]:
    """A Python boolean value argument is keyword-only, except the one argument of a single-boolean setter;
    three or more adjacent positional value arguments of one type are keyword-only.
    """
    for name, fn in sorted(_public_pool(ctx, "py").items()):
        short = name.rpartition(".")[2]
        values = _value_params(fn.params, False)
        positional = [(p, t) for p, t in values if p not in fn.kwonly]
        if not single_boolean_setter(short, fn.params, False):
            for param, annotation in positional:
                if _is_bool(annotation, False):
                    yield _violation(
                        ctx,
                        "R7.opts.py",
                        "py",
                        _key(ctx, "py", name),
                        f"bool:{param}",
                        f"`{name}({param})` is a boolean; make it keyword-only",
                    )
        if _longest_run([t for _, t in positional])[0] >= 3:
            yield _violation(
                ctx,
                "R7.opts.py",
                "py",
                _key(ctx, "py", name),
                "run",
                f"`{name}` takes three or more adjacent positional arguments of one type; make them keyword-only",
            )


# ---------------------------------------------------------------------------
# R18: a setter mirrors its getter's record.
# ---------------------------------------------------------------------------


def _record_of(alternatives: Sequence[str], records: Dict[str, Set[str]]) -> Optional[str]:
    return alternatives[0] if len(alternatives) == 1 and alternatives[0] in records else None


_TS_EXTENDS_RE = re.compile(r"^export\s+interface\s+([A-Za-z_$][\w$]*)\s+extends\s+([^{]+)\{", re.M)


def _js_records(ctx: Context, surface: str) -> Dict[str, Set[str]]:
    """Interface name -> its field names, inherited ones included."""
    js = surfaces(ctx).js[surface]
    own = {n: {m.name for m in d.members if not m.is_method} for n, d in js.decls.items() if d.kind == "interface"}
    parents = {m.group(1): [p.strip() for p in m.group(2).split(",")] for m in _TS_EXTENDS_RE.finditer(js.raw)}

    def fields(name: str, seen: Tuple[str, ...] = ()) -> Set[str]:
        out = set(own.get(name, ()))
        for parent in parents.get(name, ()):
            if parent not in seen:
                out |= fields(parent, seen + (name,))
        return out

    return {name: fields(name) for name in own}


def _alias_alternatives(decls: Dict[str, object], text: str, seen: Tuple[str, ...] = ()) -> List[str]:
    """Top-level union alternatives of `text` without `null`, each union type alias also expanded in place."""
    out: List[str] = []
    for alternative in _strip_null(text, True):
        out.append(alternative)
        decl = decls.get(alternative)
        if decl is not None and decl.kind == "type" and alternative not in seen and len(_top_union(decl.text)) > 1:
            out.extend(_alias_alternatives(decls, decl.text, seen + (alternative,)))
    return out


def _accessor_pairs(pool: Dict[str, object], js: bool) -> Iterator[Tuple[str, str]]:
    for name in sorted(pool):
        if js and name.startswith("get") and "set" + name[3:] in pool:
            yield name, "set" + name[3:]
        elif not js and name.startswith("get_") and "set_" + name[4:] in pool:
            yield name, "set_" + name[4:]


@_owned("R18.js")
def _rule_js_symmetry(ctx: Context) -> Iterable[Violation]:
    """When `get<X>` returns a record `R` and `set<X>` exists, `set<X>` takes one value argument typed `R`,
    `<R>Patch` or `<R>Input` (an input's fields are `R`'s). No public argument is `Partial<...>`.
    """
    for surface in JS_SURFACES:
        pool = _public_pool(ctx, surface)
        decls = surfaces(ctx).js[surface].decls
        records = _js_records(ctx, surface)
        for getter, setter in _accessor_pairs(pool, True):
            record = _record_of(_strip_null(pool[getter].type, True), records)
            if record is None:
                continue
            shared = {p for p, _ in pool[getter].params}
            # A `sheet` argument typed by the record (through an alias such as `SheetRef`) is the value, not an
            # address.
            values = [
                (p, t)
                for p, t in pool[setter].params
                if p not in shared and (not _is_coordinate(p, True) or record in _alias_alternatives(decls, t))
            ]
            accepted = {record, f"{record}Patch", f"{record}Input", f"Patch<{record}>"}
            kinds = _strip_null(values[0][1], True) if len(values) == 1 else []
            if len(values) == 1 and record in _alias_alternatives(decls, values[0][1]):
                continue
            if len(kinds) != 1 or kinds[0] not in accepted:
                shown = ", ".join(f"{p}: {t}" for p, t in values) or "nothing"
                yield _violation(
                    ctx,
                    "R18.js",
                    surface,
                    _key(ctx, surface, setter),
                    "shape",
                    f"`{getter}` returns `{record}` but `{setter}` takes ({shown})",
                )
            elif kinds[0] == f"{record}Input" and kinds[0] in records and not records[kinds[0]] <= records[record]:
                extra = sorted(records[kinds[0]] - records[record])
                yield _violation(
                    ctx,
                    "R18.js",
                    surface,
                    _key(ctx, surface, setter),
                    "fields",
                    f"`{kinds[0]}` has fields `{record}` lacks: {extra}",
                )
        for name, member in sorted(pool.items()):
            for param, ptype in member.params:
                if "Partial<" in ptype:
                    yield _violation(
                        ctx,
                        "R18.js",
                        surface,
                        _key(ctx, surface, name),
                        f"Partial:{param}",
                        f"`{name}({param})` is `Partial<...>`; name a `<Record>Patch`",
                    )
        for decl in decls.values():
            if decl.kind == "type" and decl.text.startswith("Partial<"):
                yield _violation(
                    ctx, "R18.js", surface, _key(ctx, surface, decl.name), "Partial", f"`{decl.name}` is `Partial<...>`"
                )


_UNCHANGED_DEFAULTS = frozenset({"UNCHANGED", "formulon.UNCHANGED", "_UNCHANGED"})


def _py_dataclasses(ctx: Context) -> Dict[str, Set[str]]:
    return {
        n: set(c.fields)
        for n, c in surfaces(ctx).py.classes.items()
        if any(d.startswith(("dataclass", "dataclasses.dataclass")) for d in c.decorators)
    }


@_owned("R18.py")
def _rule_py_symmetry(ctx: Context) -> Iterable[Violation]:
    """When `get_<x>` returns a dataclass `R` and `set_<x>` exists, `set_<x>` takes one value argument typed
    `R` / `<R>Input` / `<R>Patch`, or keyword-only arguments named after `R`'s fields that default to
    `UNCHANGED`.
    """
    pool = _public_pool(ctx, "py")
    records = _py_dataclasses(ctx)
    for getter, setter in _accessor_pairs(pool, False):
        record = _record_of(_strip_null(pool[getter].returns, False), records)
        if record is None:
            continue
        fn: PyFunc = pool[setter]
        shared = {p for p, _ in pool[getter].params}
        values = [(p, t) for p, t in _value_params(fn.params, False) if p not in shared]
        defaults = dict(fn.defaults)
        kinds = _strip_null(values[0][1], False) if len(values) == 1 else []
        whole = record in kinds or (len(kinds) == 1 and kinds[0] in {f"{record}Input", f"{record}Patch"})
        partial = bool(values) and all(
            p in fn.kwonly and p in records[record] and defaults.get(p) in _UNCHANGED_DEFAULTS for p, _ in values
        )
        if not (whole or partial):
            shown = ", ".join(p for p, _ in values) or "nothing"
            yield _violation(
                ctx,
                "R18.py",
                "py",
                _key(ctx, "py", setter),
                "shape",
                f"`{getter}` returns `{record}` but `{setter}` takes ({shown})",
            )


# ---------------------------------------------------------------------------
# R7.absent: "absent" is null / None; clearing a property takes null / None.
# ---------------------------------------------------------------------------

# Doc wording for an absent value spelled `""` / `-1`; a code span is `x` in TSDoc and ``x`` in a docstring.
_EMPTY = r"`{1,2}\"\"`{1,2}"
_READ_ABSENT_RE = re.compile(
    rf"empty (?:string )?when|zero or empty when|{_EMPTY} (?:when|if|for none)|`{{1,2}}-1`{{1,2}}|"
    r"(?<![\w`])-1 when|\bempty if\b",
    re.I,
)
_CLEAR_EMPTY_RE = re.compile(
    rf"{_EMPTY} (?:clears|unsets|removes|deletes)|pass {_EMPTY}|{_EMPTY} to (?:clear|unset|remove)|"
    rf"empty `{{1,2}}\w+`{{1,2}} (?:to )?remov|(?:clears|unsets|removes) [^.]*{_EMPTY}|`{{1,2}}\[\]`{{1,2}} (?:clears|removes)",
    re.I,
)
_TS_DOC_MEMBER_RE = re.compile(r"/\*\*((?:(?!\*/).)*)\*/\s*(?:readonly\s+)?([A-Za-z_$][\w$]*)\??\s*[(:<]", re.S)
_TS_DOC_DECL_RE = re.compile(
    r"/\*\*((?:(?!\*/).)*)\*/\s*export\s+(?:declare\s+)?(?:interface|type)\s+([A-Za-z_$][\w$]*)", re.S
)
_TS_INTERFACE_RE = re.compile(r"^export\s+interface\s+([A-Za-z_$][\w$]*)[^{]*\{", re.M)


def _js_docs(raw: str) -> Dict[Tuple[str, str], str]:
    """`(owner, member)` -> the doc comment right before the member; a declaration's own doc is `(name, "")`."""
    from .ts_decl import _extract_braced_block

    out: Dict[Tuple[str, str], str] = {}
    for match in _TS_DOC_DECL_RE.finditer(raw):
        out[(match.group(2), "")] = " ".join(match.group(1).split())
    for match in _TS_INTERFACE_RE.finditer(raw):
        body = _extract_braced_block(raw, match.end())
        for doc in _TS_DOC_MEMBER_RE.finditer(body):
            out.setdefault((match.group(1), doc.group(2)), " ".join(doc.group(1).replace("*", " ").split()))
    return out


def _has_minus_one(text: str) -> bool:
    return "-1" in _top_union(text)


def _ledger_records(ctx: Context, field: str) -> Dict[str, dict]:
    """Binding record name -> ledger entry, for records that are not inputs or patches."""
    out = {}
    for entry in ctx.ledger.entries.values():
        if isinstance(entry, dict) and entry.get("kind") == "record" and entry.get("role") not in ("input", "patch"):
            for name in (entry.get(field), (entry.get("was") or {}).get(field)):
                if isinstance(name, str):
                    out[name] = entry
    return out


def _total_names(ctx: Context, field: str) -> Set[str]:
    out = set()
    for entry in ctx.ledger.entries.values():
        if isinstance(entry, dict) and "total" in entry:
            out.update(n for n in (entry.get(field), (entry.get("was") or {}).get(field)) if isinstance(n, str))
    return out


def _presence_flag(name: str, js: bool) -> bool:
    return re.fullmatch(r"has[A-Z]\w*", name) is not None if js else name.startswith("has_")


@_owned("R7.absent.js")
def _rule_js_absent(ctx: Context) -> Iterable[Violation]:
    """JS reads "absent" as `null`: no doc says a member is empty / `""` / `-1` when absent, no type carries a
    `-1` alternative, and no result record has a `has<X>` presence flag. A setter whose doc clears with `""`
    or `[]` also accepts `null`. `total:` getters are exempt from the wording.
    """
    totals = _total_names(ctx, "js")
    for surface in JS_SURFACES:
        js = surfaces(ctx).js[surface]
        members = {(owner, m.name): m for owner, decl in js.decls.items() for m in decl.members}
        for (owner, member), doc in sorted(_js_docs(js.raw).items()):
            key = _key(ctx, surface, member if owner in (WORKBOOK_CLASS, SHEET_CLASS) and member else owner)
            decl_member = members.get((owner, member))
            where = f"{owner}.{member}" if member else owner
            if member not in totals and _READ_ABSENT_RE.search(doc):
                yield _violation(
                    ctx,
                    "R7.absent.js",
                    surface,
                    key,
                    f"doc:{where}",
                    f'`{where}` documents an absent value as `""` / `-1`',
                )
            if decl_member is not None and decl_member.is_method and _CLEAR_EMPTY_RE.search(doc):
                if not any("null" in _top_union(t) for _, t in decl_member.params):
                    yield _violation(
                        ctx,
                        "R7.absent.js",
                        surface,
                        key,
                        f"clear:{where}",
                        f"`{where}` clears with an empty value; accept `null`",
                    )
        for owner, decl in sorted(js.decls.items()):
            for m in decl.members:
                slots = [(m.name, m.type)] if not m.is_method else [(f"{m.name}({p})", t) for p, t in m.params]
                for slot, text in slots:
                    if _has_minus_one(text):
                        key = _key(ctx, surface, m.name if m.is_method else owner)
                        yield _violation(
                            ctx,
                            "R7.absent.js",
                            surface,
                            key,
                            f"-1:{owner}.{slot}",
                            f"`{owner}.{slot}` has a `-1` alternative",
                        )
        for name, entry in sorted(_ledger_records(ctx, "js").items()):
            decl = js.decls.get(name)
            for m in decl.members if decl else ():
                if not m.is_method and _presence_flag(m.name, True):
                    yield _violation(
                        ctx,
                        "R7.absent.js",
                        surface,
                        _key(ctx, surface, name),
                        f"has:{m.name}",
                        f"`{name}.{m.name}` is a presence flag; make the value field nullable",
                    )


@_owned("R7.absent.py")
def _rule_py_absent(ctx: Context) -> Iterable[Violation]:
    """Python reads "absent" as `None`: no docstring says a result is empty / `""` / `-1` when absent, no
    argument defaults to `-1`, and no result record has a `has_<x>` presence flag. A setter whose docstring
    clears with `""` or `[]` also accepts `None`. `total:` getters are exempt from the wording.
    """
    py = surfaces(ctx).py
    totals = _total_names(ctx, "py")
    for name, fn in sorted(_public_pool(ctx, "py").items()):
        key = _key(ctx, "py", name)
        if name not in totals and _READ_ABSENT_RE.search(fn.doc):
            yield _violation(
                ctx, "R7.absent.py", "py", key, "doc", f'`{name}` documents an absent value as `""` / `-1`'
            )
        if _CLEAR_EMPTY_RE.search(fn.doc) and not any(
            re.search(r"Optional\[|None", t) for _, t in _value_params(fn.params, False)
        ):
            yield _violation(
                ctx, "R7.absent.py", "py", key, "clear", f"`{name}` clears with an empty value; accept `None`"
            )
        for param, default in fn.defaults:
            if default == "-1":
                yield _violation(
                    ctx, "R7.absent.py", "py", key, f"-1:{param}", f"`{name}({param})` defaults to `-1`; use `None`"
                )
    for name, entry in sorted(_ledger_records(ctx, "py").items()):
        cls = py.classes.get(name)
        for field_name in sorted(cls.fields) if cls else ():
            if _presence_flag(field_name, False):
                yield _violation(
                    ctx,
                    "R7.absent.py",
                    "py",
                    _key(ctx, "py", name),
                    f"has:{field_name}",
                    f"`{name}.{field_name}` is a presence flag; make the value field `Optional`",
                )


# ---------------------------------------------------------------------------
# R6.closed: closed value sets are typed by their binding table.
# ---------------------------------------------------------------------------

# String parameters and fields whose values the C side closes without an enum: the binding names the alias,
# or a string-literal union where the value is `None`.
CLOSED_STRINGS = {"profile_id": "ExcelProfileId", "tick_label_position": None}


def _c_enum_fields(ctx: Context) -> Dict[str, Dict[str, str]]:
    """C struct -> field -> the value set it carries (enum typedef or macro family)."""
    macros = set(_HEADER_DEFINE_RE.findall(ctx.header))
    out: Dict[str, Dict[str, str]] = {}
    for struct, fields in _struct_field_comments(ctx.header).items():
        for ctype, name, comment in fields:
            family = ctype if ctype in ctx.enum_names else _enum_family(comment, ctx.enum_names, macros)
            if family is not None:
                out.setdefault(struct, {})[name] = family
    return out


def _js_tables(ctx: Context, surface: str) -> Set[str]:
    decls = surfaces(ctx).js[surface].decls
    return {n for n, d in decls.items() if d.kind == "const" and d.text.startswith("Readonly<{")}


def _py_enums(ctx: Context) -> Set[str]:
    return {n for n, c in surfaces(ctx).py.classes.items() if any(b.endswith(("IntEnum", "IntFlag")) for b in c.bases)}


def _open_alternatives(text: str, js: bool, tables: Set[str]) -> List[str]:
    """Alternatives of a closed-set slot that are not one of the binding's tables.

    A numeric literal (a `-1` sentinel) is `R7.absent`'s to judge.
    """
    return [a for a in _strip_null(text, js) if a not in tables and not re.fullmatch(r"-?\d+", a)]


def _field_exceptions(ctx: Context, field: str) -> Set[Tuple[str, str]]:
    """`(record name, field name)` pairs a `kind: field` entry with `exception:` exempts."""
    out = set()
    for entry in ctx.ledger.entries.values():
        if isinstance(entry, dict) and entry.get("kind") == "field" and "exception" in entry:
            record = ctx.ledger.entry(entry.get("record"))
            if isinstance(record.get(field), str) and isinstance(entry.get(field), str):
                out.add((record[field], entry[field]))
    return out


def _closed_problems(ctx: Context, surface: str) -> Iterator[Tuple[str, str, str]]:
    field = _field(surface)
    js = field == "js"
    found = surfaces(ctx)
    tables = _js_tables(ctx, surface) if js else _py_enums(ctx)
    exempt = _field_exceptions(ctx, field)
    enum_fields = _c_enum_fields(ctx)
    for key, entry in sorted(ctx.ledger.entries.items()):
        if not isinstance(entry, dict) or entry.get("kind") != "record" or entry.get("c") not in enum_fields:
            continue
        name = _in_tree(found.js[surface].decls if js else found.py.classes, entry, field)
        if name is None:
            continue
        if js:
            types = {m.name: m.type for m in found.js[surface].decls[name].members if not m.is_method}
        else:
            types = dict(found.py.classes[name].fields)
        for c_field, family in sorted(enum_fields[entry["c"]].items()):
            binding = _camel(c_field) if js else c_field
            if binding not in types or (name, binding) in exempt:
                continue
            loose = _open_alternatives(types[binding], js, tables)
            if loose:
                yield (
                    key,
                    f"{name}.{binding}",
                    f"`{name}.{binding}: {types[binding]}` carries `{family}`; type it by its table",
                )
    for key, entry in sorted(ctx.ledger.entries.items()):
        params = entry.get("params") if isinstance(entry, dict) else None
        if not isinstance(params, dict):
            continue
        pool = _callables(ctx, surface)
        name = _in_tree(pool, entry, field)
        if name is None:
            continue
        actual = dict(pool[name].params)
        for param, spec in sorted(params.items()):
            binding = _camel(param) if js else param
            if isinstance(spec, str) and spec.startswith("enum:") and binding in actual:
                loose = _open_alternatives(actual[binding], js, tables)
                if loose:
                    yield key, f"{name}({binding})", f"`{name}({binding}: {actual[binding]})` admits {loose}"
    for name, member in sorted(_callables(ctx, surface).items()):
        params = member.params
        for param, ptype in params:
            snake = re.sub(r"(?<!^)(?=[A-Z])", "_", param).lower()
            if snake in CLOSED_STRINGS and _bare_string(ptype, js):
                want = CLOSED_STRINGS[snake] or "a string-literal union"
                yield (
                    _key(ctx, surface, name),
                    f"{name}({param})",
                    f"`{name}({param}: {ptype})` is closed; type it `{want}`",
                )
    decls = found.js[surface].decls if js else {}
    for owner, decl in sorted(decls.items()):
        for m in decl.members:
            snake = re.sub(r"(?<!^)(?=[A-Z])", "_", m.name).lower()
            if not m.is_method and snake in CLOSED_STRINGS and _bare_string(m.type, js):
                want = CLOSED_STRINGS[snake] or "a string-literal union"
                yield (
                    _key(ctx, surface, owner),
                    f"{owner}.{m.name}",
                    f"`{owner}.{m.name}: {m.type}` is closed; type it `{want}`",
                )


def _bare_string(text: str, js: bool) -> bool:
    return ("string" if js else "str") in _strip_null(text, js)


@_owned("R6.closed.js")
def _rule_js_closed(ctx: Context) -> Iterable[Violation]:
    """A JS field whose C field carries an enum (enum typedef, or a comment naming an enum typedef or a
    `FM_<FAMILY>_*` set) and an argument the ledger types `enum:<key>` admit only the binding's `as const`
    tables; a closed string set (`CLOSED_STRINGS`) is its alias or a literal union, never bare `string`.
    """
    for surface in JS_SURFACES:
        for key, disc, message in _closed_problems(ctx, surface):
            yield _violation(ctx, "R6.closed.js", surface, key, disc, message)


@_owned("R6.closed.py")
def _rule_py_closed(ctx: Context) -> Iterable[Violation]:
    """The Python counterpart of `R6.closed.js`: an enum slot is annotated with its `IntEnum` alone (no bare
    or added `int`), a closed string set with its `Literal` alias.
    """
    for key, disc, message in _closed_problems(ctx, "py"):
        yield _violation(ctx, "R6.closed.py", "py", key, disc, message)


# ---------------------------------------------------------------------------
# R19: enum families.
# ---------------------------------------------------------------------------

# Enum members that stand for "the attribute is absent": the binding table drops them, and the C
# value of a field-level absence is `-1`. Keyed by the binding enum name; the flag says whether C keeps the
# member as `-1` (a field that reads `null`) or as a patch instruction (a record that reads `null`).
ABSENT_MEMBERS = {
    "ChartBlanksAs": ("Absent", True),
    "ChartAxisCrosses": ("Absent", True),
    "ShapePaint": ("Unset", True),
    "ChartBarDirection": ("Absent", True),
    "PivotFilterValueKind": ("None", True),
    "ChartTitleKind": ("None", False),
    "ChartLegendPosition": ("None", False),
}


def _norm_member(name: str) -> str:
    return name.replace("_", "").lower()


def c_enum_members(ctx: Context) -> Dict[str, Dict[str, int]]:
    """C enum typedef -> normalized member name (common prefix removed) -> value."""
    out: Dict[str, Dict[str, int]] = {}
    header = _COMMENT_RE.sub(" ", ctx.header)
    for match in re.finditer(r"typedef\s+enum\s*\{([^}]*)\}\s*(\w+)\s*;", header):
        names: List[str] = []
        values: List[Optional[int]] = []
        current: Optional[int] = 0
        for part in match.group(1).split(","):
            found = re.match(r"\s*(\w+)\s*(?:=\s*(-?\w+))?", part)
            if not found or not found.group(1):
                continue
            if found.group(2) is not None:
                try:
                    current = int(found.group(2), 0)
                except ValueError:
                    current = None
            names.append(found.group(1))
            values.append(current)
            current = current + 1 if current is not None else None
        words = [n.split("_") for n in names]
        depth = 0
        if len(names) > 1:
            while all(len(w) > depth + 1 for w in words) and len({tuple(w[: depth + 1]) for w in words}) == 1:
                depth += 1
        else:
            stem = match.group(2)[3:-2].upper().split("_")
            depth = len(stem) + 1 if names and names[0].startswith("FM_" + "_".join(stem) + "_") else 1
        out[match.group(2)] = {"".join(w[depth:]).lower(): v for w, v in zip(words, values) if v is not None}
    return out


def _enum_key_by_c(ctx: Context) -> Dict[str, str]:
    return {
        e["c"]: k
        for k, e in ctx.ledger.entries.items()
        if isinstance(e, dict) and e.get("kind") in ("enum", "typedef") and isinstance(e.get("c"), str)
    }


def _families(ctx: Context) -> Dict[str, List[str]]:
    data = ctx.ledger.data.get("enum_families") or {}
    return {
        f: [k for k in (m.get("enums") if isinstance(m, dict) else m) or [] if isinstance(k, str)]
        for f, m in data.items()
    }


def _distinct_pairs(ctx: Context) -> Set[frozenset]:
    out = set()
    for item in ctx.ledger.data.get("enum_distinct") or []:
        enums = item.get("enums") if isinstance(item, dict) else None
        if isinstance(enums, list) and len(enums) == 2:
            out.add(frozenset(enums))
    return out


def candidate_pairs(ctx: Context) -> List[Tuple[str, str, List[str]]]:
    """Ledger enum pairs sharing two or more member names, measured on the C enums."""
    members = c_enum_members(ctx)
    keys = _enum_key_by_c(ctx)
    out = []
    for a, b in itertools.combinations(sorted(members), 2):
        shared = sorted(set(members[a]) & set(members[b]))
        if len(shared) >= 2 and a in keys and b in keys:
            out.append((keys[a], keys[b], shared))
    return out


def _binding_values(ctx: Context, surface: str, entry: dict) -> Optional[Dict[str, int]]:
    """Normalized member name -> value of an enum's binding table, or None when it has none."""
    field = _field(surface)
    name = entry.get(field)
    if not isinstance(name, str):
        return None
    if field == "py":
        cls = surfaces(ctx).py.classes.get(name)
        return {_norm_member(k): v for k, v in cls.constants.items()} if cls else None
    from .enum_decl import _parse_ts_named_enum

    table = _parse_ts_named_enum(surfaces(ctx).js[surface].raw, name)
    return {_norm_member(k): v for k, v in table.items()} if table else None


def _family_mismatches(values: Dict[str, Dict[str, int]]) -> Iterator[Tuple[str, str, str]]:
    for a, b in itertools.combinations(sorted(values), 2):
        for member in sorted(set(values[a]) & set(values[b])):
            if values[a][member] != values[b][member]:
                yield b, member, f"`{member}` is {values[a][member]} in `{a}` but {values[b][member]} in `{b}`"


@_owned("R19.c")
def _rule_c_enum_families(ctx: Context) -> Iterable[Violation]:
    """Enums of one ledger `enum_families` family give a same-named member the same C value; every enum pair
    sharing two or more member names sits in one family or in `enum_distinct`; a field-level absence member
    (`ABSENT_MEMBERS`) is `-1`.
    """
    members = c_enum_members(ctx)
    by_key = {v: k for k, v in _enum_key_by_c(ctx).items()}
    families = _families(ctx)
    for family, keys in sorted(families.items()):
        values = {k: members[by_key[k]] for k in keys if k in by_key and by_key[k] in members}
        for key, member, message in _family_mismatches(values):
            yield _violation(ctx, "R19.c", "c", key, member, f"family `{family}`: {message}")
    grouped = {frozenset(p) for keys in families.values() for p in itertools.combinations(keys, 2)}
    distinct = _distinct_pairs(ctx)
    for a, b, shared in candidate_pairs(ctx):
        pair = frozenset((a, b))
        if pair not in grouped and pair not in distinct:
            yield _violation(
                ctx, "R19.c", "c", b, f"pair:{a}", f"`{a}` and `{b}` share {shared}; list them as a family or distinct"
            )
    for key, entry in sorted(ctx.ledger.entries.items()):
        if not isinstance(entry, dict) or entry.get("js") not in ABSENT_MEMBERS:
            continue
        member, field_level = ABSENT_MEMBERS[entry["js"]]
        value = members.get(entry.get("c"), {}).get(_norm_member(member))
        if field_level and value is not None and value != -1:
            yield _violation(
                ctx, "R19.c", "c", key, f"absent:{member}", f"`{member}` means absent; its C value is -1, not {value}"
            )


def _binding_family_rule(ctx: Context, rule_id: str, surface_list: Sequence[str]) -> Iterator[Violation]:
    for surface in surface_list:
        for family, keys in sorted(_families(ctx).items()):
            values = {}
            for key in keys:
                table = _binding_values(ctx, surface, ctx.ledger.entry(key))
                if table is not None:
                    values[key] = table
            for key, member, message in _family_mismatches(values):
                yield _violation(ctx, rule_id, surface, key, member, f"family `{family}`: {message}")
        for key, entry in sorted(ctx.ledger.entries.items()):
            if not isinstance(entry, dict) or entry.get("js") not in ABSENT_MEMBERS:
                continue
            member = ABSENT_MEMBERS[entry["js"]][0]
            table = _binding_values(ctx, surface, entry) or {}
            if _norm_member(member) in table:
                yield _violation(
                    ctx,
                    rule_id,
                    surface,
                    key,
                    f"absent:{member}",
                    f"`{entry.get(_field(surface))}` lists `{member}`; absent reads as null / None",
                )


@_owned("R19.js")
def _rule_js_enum_families(ctx: Context) -> Iterable[Violation]:
    """The JS tables of one enum family agree on every same-named member; no table lists an absence member."""
    yield from _binding_family_rule(ctx, "R19.js", JS_SURFACES)


@_owned("R19.py")
def _rule_py_enum_families(ctx: Context) -> Iterable[Violation]:
    """The Python `IntEnum`s of one enum family agree on every same-named member; none lists an absence member."""
    yield from _binding_family_rule(ctx, "R19.py", ("py",))


# ---------------------------------------------------------------------------
# R20: results are read-only.
# ---------------------------------------------------------------------------


def _mutable_array(text: str) -> bool:
    return any(a.endswith("[]") and not a.startswith("readonly ") for a in _strip_null(text, True))


@_owned("R20.js")
def _rule_js_readonly(ctx: Context) -> Iterable[Violation]:
    """A JS method returning an array returns `readonly T[]`; an array field of a result record is
    `readonly T[]`; `ReadonlyArray<T>` is spelled `readonly T[]`.
    """
    records = _ledger_records(ctx, "js")
    for surface in JS_SURFACES:
        js = surfaces(ctx).js[surface]
        for name, member in sorted(_public_pool(ctx, surface).items()):
            if _mutable_array(member.type):
                yield _violation(
                    ctx, "R20.js", surface, _key(ctx, surface, name), "return", f"`{name}` returns `{member.type}`"
                )
        for owner, decl in sorted(js.decls.items()):
            for m in decl.members:
                texts = [m.type] + [t for _, t in m.params]
                if any("ReadonlyArray<" in t for t in texts):
                    yield _violation(
                        ctx,
                        "R20.js",
                        surface,
                        _key(ctx, surface, m.name if m.is_method else owner),
                        f"ReadonlyArray:{owner}.{m.name}",
                        f"`{owner}.{m.name}` spells `ReadonlyArray<T>`; write `readonly T[]`",
                    )
                elif owner in records and not m.is_method and _mutable_array(m.type):
                    yield _violation(
                        ctx,
                        "R20.js",
                        surface,
                        _key(ctx, surface, owner),
                        f"field:{m.name}",
                        f"`{owner}.{m.name}: {m.type}` is a result array; make it `readonly T[]`",
                    )


@_owned("R20.py")
def _rule_py_readonly(ctx: Context) -> Iterable[Violation]:
    """An array field of a Python result record is a `List[T]`, not a tuple or a `Sequence`."""
    py = surfaces(ctx).py
    for name, _entry in sorted(_ledger_records(ctx, "py").items()):
        cls = py.classes.get(name)
        for field_name, annotation in sorted(cls.fields.items()) if cls else ():
            text = annotation.strip("'\"")
            homogeneous = re.match(r"(?:typing\.)?(?:Tuple|tuple)\[.*,\s*\.\.\.\]$", text)
            if homogeneous or re.match(r"(?:typing\.)?(?:Sequence|FrozenSet|frozenset)\[", text):
                yield _violation(
                    ctx,
                    "R20.py",
                    "py",
                    _key(ctx, "py", name),
                    f"field:{field_name}",
                    f"`{name}.{field_name}: {annotation}` is a result array; make it `List[...]`",
                )


# ---------------------------------------------------------------------------
# R21: record roles and names.
# ---------------------------------------------------------------------------

RECORD_ROLES = ("result", "input", "patch", "both")


def _records_in_ledger(ctx: Context) -> Iterator[Tuple[str, dict]]:
    for key, entry in sorted(ctx.ledger.entries.items()):
        if isinstance(entry, dict) and entry.get("kind") == "record" and "replaced_by" not in entry:
            yield key, entry


@_owned("R21.role")
def _rule_record_roles(ctx: Context) -> Iterable[Violation]:
    """Every ledger record carries `role:` (result / input / patch / both)."""
    for key, entry in _records_in_ledger(ctx):
        if entry.get("role") not in RECORD_ROLES:
            listed = _entry_surfaces(entry) or ["c"]
            surface = next((s for s in BINDING_SURFACES if s in listed), "c")
            yield _violation(ctx, "R21.role", surface, key, "role", f"`{key}` has no `role:`")


@_owned("R21.name")
def _rule_record_names(ctx: Context) -> Iterable[Violation]:
    """A binding record name carries no meaningless suffix (`Entry`, `Record`, `Info`, `Spec`), and the
    Python record is named like the JS one; an `exception:` keeps a name.
    """
    found = surfaces(ctx)
    for key, entry in _records_in_ledger(ctx):
        if "exception" in entry:
            continue
        for surface in BINDING_SURFACES:
            field = _field(surface)
            pool = found.py.classes if field == "py" else found.js[surface].decls
            name = _in_tree(pool, entry, field)
            if name is not None and strip_record_suffix(name) != name:
                yield _violation(
                    ctx,
                    "R21.name",
                    surface,
                    key,
                    "suffix",
                    f"`{name}` ends with one of {list(MEANINGLESS_SUFFIXES)}; name it `{strip_record_suffix(name)}`",
                )
        js, py = entry.get("js"), entry.get("py")
        if isinstance(js, str) and isinstance(py, str) and js != py:
            yield _violation(ctx, "R21.name", "py", key, "py-name", f"the Python record `{py}` is `{js}` in JS")


# ---------------------------------------------------------------------------
# R22: index argument names.
# ---------------------------------------------------------------------------

_BARE_INDEX = frozenset({"index", "idx"})


def _type_derived(name: str) -> bool:
    return name == "utf8" or name.startswith("utf8_") or re.fullmatch(r"utf8[A-Z]\w*", name) is not None


@_owned("R22.c")
def _rule_c_index_names(ctx: Context) -> Iterable[Violation]:
    """A C argument addressing a collection element is `<thing>_index` (an `_at` enumeration position keeps
    `index`); no argument is named after its C type (`utf8`).
    """
    for fn in ctx.functions:
        at = ctx.target_c(fn.name).endswith("_at")
        scope = ctx.ledger.entry(ctx.ledger.key_for("c", fn.name)).get("scope")
        if scope not in (None, "workbook", "sheet", "process"):
            continue
        for param in fn.params:
            if param.role_name in _BARE_INDEX and param.pointer_depth <= 1 and not (at and param.name == "index"):
                yield ctx.c_violation(
                    "R22.c",
                    fn.name,
                    param.name,
                    f"{fn.name}: `{param.name}` names no collection; use `<thing>_index`",
                    fn.line,
                )
            elif _type_derived(param.role_name):
                yield ctx.c_violation(
                    "R22.c", fn.name, param.name, f"{fn.name}: `{param.name}` is named after its C type", fn.line
                )


def _binding_index_names(ctx: Context, surface: str, rule_id: str) -> Iterator[Violation]:
    js = surface != "py"
    for name, member in sorted(_public_pool(ctx, surface).items()):
        short = name.rpartition(".")[2]
        at = short.endswith("At") if js else short.endswith("_at")
        for param, _ in member.params:
            if param in _BARE_INDEX and not (at and param == "index"):
                want = "<thing>Index" if js else "<thing>_index"
                yield _violation(
                    ctx, rule_id, surface, _key(ctx, surface, name), param, f"`{name}({param})`: use `{want}`"
                )
            elif _type_derived(param):
                yield _violation(
                    ctx, rule_id, surface, _key(ctx, surface, name), param, f"`{name}({param})` is named after a C type"
                )


@_owned("R22.js")
def _rule_js_index_names(ctx: Context) -> Iterable[Violation]:
    """A JS argument addressing a collection element is `<thing>Index`; none is named after a C type."""
    for surface in JS_SURFACES:
        yield from _binding_index_names(ctx, surface, "R22.js")


@_owned("R22.py")
def _rule_py_index_names(ctx: Context) -> Iterable[Violation]:
    """A Python argument addressing a collection element is `<thing>_index`; none is named after a C type."""
    yield from _binding_index_names(ctx, "py", "R22.py")


# ---------------------------------------------------------------------------
# R24.missing: what a lookup does when nothing is there.
# ---------------------------------------------------------------------------

MISSING_MODES = ("null", "throw")


def missing_mode(entry: object) -> Optional[str]:
    """An entry's `missing:` mode; YAML reads `missing: null` as None, so the key's presence decides."""
    if not isinstance(entry, dict) or "missing" not in entry:
        return None
    return "null" if entry["missing"] is None else entry["missing"]


@_owned("R24.missing.js")
def _rule_js_missing(ctx: Context) -> Iterable[Violation]:
    """An op with `missing: null` returns `... | null` in JS; one with `missing: throw` does not. Natural-key
    getters are judged by `R8.js`.
    """
    lookups = _optional_lookups(ctx)
    for surface in JS_SURFACES:
        pool = _callables(ctx, surface)
        for key, entry in sorted(ctx.ledger.entries.items()):
            mode = missing_mode(entry)
            name = _in_tree(pool, entry, "js") if mode in MISSING_MODES and key not in lookups else None
            if name is None:
                continue
            nullable = "null" in _top_union(pool[name].type)
            if nullable != (mode == "null"):
                want = "returns `... | null`" if mode == "null" else "throws `NotFound` instead of returning null"
                yield _violation(
                    ctx, "R24.missing.js", surface, key, mode, f"`{name}` returns `{pool[name].type}`; it {want}"
                )


@_owned("R24.missing.py")
def _rule_py_missing(ctx: Context) -> Iterable[Violation]:
    """An op with `missing: null` returns `Optional[...]` in Python; one with `missing: throw` documents
    `Raises:` with `NotFoundError`.
    """
    lookups = _optional_lookups(ctx)
    pool = _callables(ctx, "py")
    for key, entry in sorted(ctx.ledger.entries.items()):
        mode = missing_mode(entry)
        name = _in_tree(pool, entry, "py") if mode in MISSING_MODES and key not in lookups else None
        if name is None:
            continue
        fn = pool[name]
        if mode == "null" and fn.returns and not re.search(r"Optional\[|None", fn.returns):
            yield _violation(ctx, "R24.missing.py", "py", key, mode, f"`{name}` returns `{fn.returns}`, not `Optional`")
        elif mode == "throw" and not ("Raises" in fn.doc and "NotFoundError" in fn.doc):
            yield _violation(
                ctx, "R24.missing.py", "py", key, mode, f"`{name}` does not document `Raises: NotFoundError`"
            )
