"""API consistency rules: verbs of patches and singletons, add results, closed sets, Python argument shapes, C
patch fields, boolean polarity, units and C status documentation.

R36.verb.update    A patch op (`patch: true`, or a `set` op taking a C `fm_*_patch_t` or a JS `*Patch`)
                   is named with the verb `update` on every surface the ledger names; `set` never merges.
R37.add.returns    An `add` op states `add_returns: handle | key | index`, and it is the kind its collection's
                   elements are addressed by: `handle` for a handle collection, `key` for a natural key,
                   `index` for a position.
R38.enum.numeric   No field of a d.ts record or input is a union of string literals (directly or through a
                   type alias); a closed set is a numeric table. `ProfileId` is the one string-closed set.
R39.py.args        A Python `add_*` / `set_*` whose JS twin takes a record or input takes it as one positional
                   dataclass argument; an `update_*` takes its keys keyword-only, each defaulting to
                   `UNCHANGED`.
R40.c.patch_op     A C patch struct (`fm_*_patch_t`, or the struct an `update` function takes) gives every value
                   field a `<field>_op` typed `fm_patch_op_t`; no `_mask`, `has_` or `_engaged` field and no
                   negative value keeps a setting.
R41.bool.polarity  A boolean field of a protection record starts with `allow` or `lock`, and its input and patch
                   records use the read record's field names.
R42.unit.suffix    A numeric name that carries a dimension (width, height, margin, angle, ...) ends with a unit
                   (`Pt`, `Inches`, `Deg`, ...), unless the ledger's top-level `unitless:` list names it.
R43.c.status_doc   Every `FM_API` declaration returning a status documents its statuses with `@return`, or refers to the common
                   failures stated at the head of its section.
R44.singleton.null A singleton (a noun with no `add`, outside every collection) has no `remove` / `clear` op; it
                   is removed with `set<X>(null)`. A `reset` op restores the default of a noun that has a `set`.

R36, R37 and R44 report on the `ledger` surface: the ledger names the op and changes first.
"""

from __future__ import annotations

import re
from typing import Dict, Iterable, Iterator, List, Optional, Sequence, Set, Tuple

from .api_conventions import (
    _C_STRUCT_RE,
    LEDGER_PATH,
    LEDGER_SURFACE,
    NO_DISCRIMINATOR,
    CFunction,
    Context,
    CParam,
    Violation,
    _doc_before,
    _struct_field_comments,
    c_verb,
)
from .api_conventions_arguments import _targets
from .api_conventions_bindings import (
    BINDING_SURFACES,
    JS_SURFACES,
    PyClass,
    TsDecl,
    TsMember,
    _key,
    _owned,
    _parse_members,
    _top_union,
    _tree_c,
    _violation,
    _words,
    surfaces,
)
from .api_conventions_shapes import _UNCHANGED_DEFAULTS, _address_positions, _is_coordinate, _strip_null
from .surface_files import REPO_ROOT
from .ts_decl import _extract_braced_block

_LEDGER_SITE = str(LEDGER_PATH.relative_to(REPO_ROOT))


def _ledger_violation(rule_id: str, key: str, disc: str, message: str) -> Violation:
    return Violation(rule_id, LEDGER_SURFACE, key, disc, message, _LEDGER_SITE)


def _ops(ctx: Context) -> Iterator[Tuple[str, dict]]:
    """Every published op of the ledger, by key."""
    for key, entry in sorted(ctx.ledger.entries.items()):
        if isinstance(entry, dict) and entry.get("kind") == "op" and entry.get("internal") is not True:
            if "replaced_by" not in entry:
                yield key, entry


def _js_verb(name: object) -> Optional[str]:
    """The leading lower-case word of a JS name (`setChartAxis` -> `set`), the class prefix dropped."""
    if not isinstance(name, str):
        return None
    found = re.match(r"[a-z]+", name.rpartition(".")[2])
    return found.group(0) if found else None


def _py_verb(name: object) -> Optional[str]:
    return name.rpartition(".")[2].split("_", 1)[0] if isinstance(name, str) else None


# ---------------------------------------------------------------------------
# R36.verb.update
# ---------------------------------------------------------------------------


def _is_patch_type(text: str, js: bool) -> bool:
    return any(a.endswith("Patch") or a.startswith("Patch<") for a in _strip_null(text, js))


def _patch_evidence(ctx: Context, key: str, entry: dict) -> Optional[str]:
    """Why an op merges a patch, or None: `patch: true`, or a `set` op whose C or JS signature takes a patch."""
    if entry.get("patch") is True:
        return "`patch: true`"
    if key.rpartition(".")[2] != "set":
        return None
    fn = _tree_c(ctx, entry)
    for param in fn.params if fn else ():
        if param.base_type.endswith("_patch_t") and param.pointer_depth == 1:
            return f"its C function takes `{param.base_type}`"
    for surface in JS_SURFACES:
        for target in _targets(ctx, entry, surface):
            for name, ptype in target.member.params:  # type: ignore[attr-defined]
                if _is_patch_type(ptype, True):
                    return f"`{target.name}({name})` is a patch"
    return None


@_owned("R36.verb.update")
def _rule_update_verb(ctx: Context) -> Iterable[Violation]:
    """A patch op (`patch: true`, or a `set` op taking a C `fm_*_patch_t` or a JS `*Patch`) carries the verb
    `update` in each of its C, JS and Python names; `set` replaces a whole value and never merges."""
    for key, entry in _ops(ctx):
        why = _patch_evidence(ctx, key, entry)
        if why is None:
            continue
        verbs = {"c": c_verb(entry["c"]) if isinstance(entry.get("c"), str) else None}
        verbs["js"] = _js_verb(entry.get("js"))
        verbs["py"] = _py_verb(entry.get("py"))
        wrong = [f"{field} `{entry[field]}`" for field, verb in verbs.items() if verb not in (None, "update")]
        if wrong:
            yield _ledger_violation(
                "R36.verb.update", key, "verb", f"{why}; a merge takes the verb `update`: {', '.join(wrong)}"
            )


# ---------------------------------------------------------------------------
# R37.add.returns
# ---------------------------------------------------------------------------

ADD_RETURNS = ("handle", "key", "index")


def _collection_of(ctx: Context) -> Dict[str, dict]:
    out: Dict[str, dict] = {}
    for collection in ctx.ledger.collections.values():
        if isinstance(collection, dict):
            for op in collection.get("ops") or []:
                out.setdefault(op, collection)
    return out


def addressed_by(collection: dict) -> str:
    """How a collection's elements are addressed: by `handle`, natural `key` or `index`."""
    if collection.get("access") == "handles":
        return "handle"
    return "key" if collection.get("key") == "natural" else "index"


def _is_add(key: str, entry: dict) -> bool:
    if key.rpartition(".")[2] == "add":
        return True
    c_name = entry.get("c")
    return (isinstance(c_name, str) and c_verb(c_name) == "add") or _js_verb(entry.get("js")) == "add"


@_owned("R37.add.returns")
def _rule_add_returns(ctx: Context) -> Iterable[Violation]:
    """An `add` op states `add_returns: handle | key | index`, the kind its collection is addressed by (the
    argument kind of the collection's `remove`)."""
    owners = _collection_of(ctx)
    for key, entry in _ops(ctx):
        if not _is_add(key, entry):
            continue
        stated = entry.get("add_returns")
        collection = owners.get(key)
        if stated is None:
            want = f"; its collection is addressed by {addressed_by(collection)}" if collection else ""
            yield _ledger_violation("R37.add.returns", key, "add_returns", f"`add_returns` is missing{want}")
        elif collection is not None and stated != addressed_by(collection):
            yield _ledger_violation(
                "R37.add.returns",
                key,
                "remove",
                f"`add_returns: {stated}` but the collection is addressed by {addressed_by(collection)}",
            )


# ---------------------------------------------------------------------------
# R38.enum.numeric
# ---------------------------------------------------------------------------

# The one closed set the C side also spells as strings.
STRING_CLOSED_TYPES = frozenset({"ProfileId"})
_STRING_LITERAL_RE = re.compile(r"'[^']*'|\"[^\"]*\"")


def _alternatives(text: str) -> List[str]:
    """Top-level union alternatives without `null` / `undefined`, quotes kept."""
    return [a.strip() for a in _top_union(text) if a.strip() not in ("null", "undefined")]


def _string_union(decls: Dict[str, object], text: str) -> Optional[str]:
    """The name of the string-literal union `text` is (itself when spelled inline), or None."""
    alternatives = _alternatives(text)
    if len(alternatives) == 1 and alternatives[0] in decls:
        decl = decls[alternatives[0]]
        if getattr(decl, "kind", "") != "type" or alternatives[0] in STRING_CLOSED_TYPES:
            return None
        name, alternatives = alternatives[0], _alternatives(decl.text)  # type: ignore[attr-defined]
    else:
        name = text.strip()
    literals = [a for a in alternatives if _STRING_LITERAL_RE.fullmatch(a.strip())]
    return name if len(literals) >= 2 and len(literals) == len(alternatives) else None


def _record_members(decl: TsDecl) -> Tuple[TsMember, ...]:
    """Fields of a record interface, or of the object literal a `Patch<{ ... }>` alias wraps."""
    if decl.kind == "interface":
        return () if any(m.is_method for m in decl.members) else decl.members
    if decl.kind == "type" and decl.text.startswith("Patch<{"):
        open_at = decl.text.index("{")
        return _parse_members(_extract_braced_block(decl.text, open_at + 1))
    return ()


@_owned("R38.enum.numeric")
def _rule_numeric_enums(ctx: Context) -> Iterable[Violation]:
    """No field of a d.ts record or input is a union of string literals; a closed set is a numeric table
    (`ProfileId` excepted)."""
    for surface in JS_SURFACES:
        decls = surfaces(ctx).js[surface].decls
        for name, decl in sorted(decls.items()):
            for member in _record_members(decl):
                union = _string_union(decls, member.type)
                if union:
                    yield _violation(
                        ctx,
                        "R38.enum.numeric",
                        surface,
                        _key(ctx, surface, name),
                        member.name,
                        f"`{name}.{member.name}` is the string union `{union}`; make it a numeric table",
                    )


# ---------------------------------------------------------------------------
# R39.py.args
# ---------------------------------------------------------------------------

# Records that address a cell or range, and values the bindings pass natively (a date, an error value); an
# argument of these types is not an input record.
_NOT_INPUT_RECORDS = frozenset({"CellRect", "SheetRect", "CellAddress", "CellRef", "CivilTime", "ErrorValue"})


def _takes_record(ctx: Context, entry: dict) -> Optional[str]:
    """The record or input type the JS twin of an op takes as its value, or None."""
    decls = surfaces(ctx).js["wasm"].decls
    addressed = _address_positions(entry)
    for target in _targets(ctx, entry, "wasm"):
        for position, (name, ptype) in enumerate(target.member.params):  # type: ignore[attr-defined]
            if position in addressed or _is_coordinate(name, True):
                continue
            for alternative in _strip_null(ptype, True):
                decl = decls.get(alternative)
                if (
                    decl is not None
                    and decl.kind == "interface"
                    and alternative not in _NOT_INPUT_RECORDS
                    and not alternative.endswith("Patch")
                    and not any(m.is_method for m in decl.members)
                ):
                    return alternative
    return None


def _dataclass_names(classes: Dict[str, PyClass]) -> Set[str]:
    return {
        n for n, c in classes.items() if any(d.startswith(("dataclass", "dataclasses.dataclass")) for d in c.decorators)
    }


def _mentions(text: str, names: Iterable[str]) -> bool:
    return any(re.search(rf"(?<![\w.]){re.escape(n)}(?!\w)", text) for n in names)


@_owned("R39.py.args")
def _rule_py_args(ctx: Context) -> Iterable[Violation]:
    """A Python `add_*` / `set_*` whose JS twin takes a record or input takes one positional dataclass argument;
    an `update_*` takes its keys keyword-only, each defaulting to `UNCHANGED`."""
    dataclasses = _dataclass_names(surfaces(ctx).py.classes)
    for key, entry in _ops(ctx):
        if not isinstance(entry.get("py"), str):
            continue
        for target in _targets(ctx, entry, "py"):
            fn = target.member
            verb = target.name.rpartition(".")[2].split("_", 1)[0]
            positional = [(p, t) for p, t in fn.params if p not in fn.kwonly]  # type: ignore[attr-defined]
            if verb in ("add", "set"):
                record = _takes_record(ctx, entry)
                if record is None:
                    continue
                if not any(_mentions(t, dataclasses) for _, t in positional):
                    yield _violation(
                        ctx,
                        "R39.py.args",
                        "py",
                        key,
                        "input",
                        f"`{target.name}` takes `{record}` in JS; take one positional dataclass argument",
                    )
            elif verb == "update":
                defaults = dict(fn.defaults)  # type: ignore[attr-defined]
                if not fn.kwonly:  # type: ignore[attr-defined]
                    yield _violation(
                        ctx, "R39.py.args", "py", key, "keywords", f"`{target.name}` takes no keyword-only keys"
                    )
                for param in fn.kwonly:  # type: ignore[attr-defined]
                    if defaults.get(param) not in _UNCHANGED_DEFAULTS:
                        yield _violation(
                            ctx,
                            "R39.py.args",
                            "py",
                            key,
                            f"default:{param}",
                            f"`{target.name}({param})` defaults to `{defaults.get(param)}`, not `UNCHANGED`",
                        )
                ref = entry.get("ref_args") or [None]
                for position, (param, _) in enumerate(positional):
                    if position != ref[0] and not _is_coordinate(param, False):
                        yield _violation(
                            ctx,
                            "R39.py.args",
                            "py",
                            key,
                            f"positional:{param}",
                            f"`{target.name}({param})` is a patch key; make it keyword-only",
                        )


# ---------------------------------------------------------------------------
# R40.c.patch_op
# ---------------------------------------------------------------------------

PATCH_OP_TYPE = "fm_patch_op_t"
# Address structs an `update` function may take after its patch.
_NOT_PATCH_STRUCTS = frozenset({"fm_cell_rect_t", "fm_sheet_rect_t"})
_NEGATIVE_KEEP_RE = re.compile(r"negative\b[^.]*\b(?:keeps?|leaves?|unchanged)|`?-1`? (?:keeps|leaves)", re.I)


def _patch_structs(ctx: Context) -> Set[str]:
    """`fm_*_patch_t` structs and the first struct an `update` function takes by `const T*` (its patch)."""
    out = {name for name in ctx.struct_names if name.endswith("_patch_t")}
    for fn in ctx.functions:
        if c_verb(fn.name) != "update":
            continue
        out.update(
            [
                p.base_type
                for p in fn.params
                if p.base_type in ctx.struct_names and "const" in p.ctype and p.base_type not in _NOT_PATCH_STRUCTS
            ][:1]
        )
    return out


def _struct_docs(header: str) -> Dict[str, str]:
    """Struct name -> its doc comment and its field comments, as one text."""
    out: Dict[str, str] = {}
    comments = _struct_field_comments(header)
    for match in _C_STRUCT_RE.finditer(header):
        name = match.group(2)
        fields = " ".join(comment for _, _, comment in comments.get(name, ()))
        out[name] = f"{_doc_before(header, match.start())} {fields}"
    return out


def _patch_field_problems(fields: Sequence[CParam]) -> Iterator[Tuple[str, str]]:
    names = {f.name for f in fields}
    types = {f.name: f.ctype for f in fields}
    for f in fields:
        name = f.name
        if name == "struct_size" or f.base_type.endswith("_patch_t"):
            continue
        if name.endswith("_op"):
            if types[name] != PATCH_OP_TYPE:
                yield f"op-type:{name}", f"`{name}` is `{types[name]}`, not `{PATCH_OP_TYPE}`"
        elif name.endswith("_mask"):
            yield f"mask:{name}", f"`{name}` is a mask; give each field a `<field>_op`"
        elif name.startswith("has_"):
            yield f"has:{name}", f"`{name}` keeps a field; give it a `<field>_op`"
        elif name.endswith("_engaged"):
            yield f"engaged:{name}", f"`{name}` keeps a field; give it a `<field>_op`"
        elif name.endswith("_count") and name[: -len("_count")] in names:
            continue
        elif f"{name}_op" not in names:
            yield f"op:{name}", f"`{name}` has no `{name}_op`"


@_owned("R40.c.patch_op")
def _rule_patch_ops(ctx: Context) -> Iterable[Violation]:
    """A C patch struct gives every value field a `<field>_op` typed `fm_patch_op_t` (an array its `_count` and
    one op, an embedded sub-patch its own ops); no `_mask`, `has_` or `_engaged` field, no negative keep."""
    structs = {t.name: t for t in ctx.typedefs if t.kind == "struct"}
    docs = _struct_docs(ctx.header)
    for name in sorted(_patch_structs(ctx)):
        typedef = structs[name]
        for disc, message in _patch_field_problems(typedef.fields):
            yield ctx.c_violation("R40.c.patch_op", name, disc, message, typedef.line)
        if _NEGATIVE_KEEP_RE.search(docs.get(name, "")):
            yield ctx.c_violation(
                "R40.c.patch_op", name, "negative-keep", "a negative value keeps a setting", typedef.line
            )


# ---------------------------------------------------------------------------
# R41.bool.polarity
# ---------------------------------------------------------------------------

# A cell format part: its flags are the OOXML `<protection locked hidden>` pair, not permissions.
PROTECTION_EXEMPT = frozenset({"CellProtection", "CellProtectionPatch", "CellProtectionInput"})
_PROTECTION_RE = re.compile(r"[A-Za-z]*Protection(Input|Patch)?")
_JS_POLARITY = ("allow", "lock")
_PY_POLARITY = ("allow_", "lock_")


def _protection_records(names: Iterable[str]) -> List[Tuple[str, str]]:
    """`(name, read record)` of every protection record, input and patch."""
    out = []
    for name in sorted(names):
        match = _PROTECTION_RE.fullmatch(name)
        if match and name not in PROTECTION_EXEMPT:
            out.append((name, name[: len(name) - len(match.group(1) or "")]))
    return out


def _js_bool(ptype: str) -> bool:
    return _strip_null(ptype, True) == ["boolean"]


def _py_bool(annotation: str) -> bool:
    return [a for a in _strip_null(annotation, False) if a.strip() not in ("Unchanged", "_Unchanged")] == ["bool"]


@_owned("R41.bool.polarity")
def _rule_bool_polarity(ctx: Context) -> Iterable[Violation]:
    """A boolean field of a protection record starts with `allow` / `lock`, and the input and patch records of a
    protection use the read record's field names."""
    found = surfaces(ctx)
    for surface in BINDING_SURFACES:
        if surface == "py":
            fields = {n: dict(c.fields) for n, c in found.py.classes.items()}
            is_bool, prefixes = _py_bool, _PY_POLARITY
        else:
            decls = found.js[surface].decls
            fields = {
                n: {m.name: m.type for m in d.members if not m.is_method}
                for n, d in decls.items()
                if d.kind == "interface"
            }
            is_bool, prefixes = _js_bool, _JS_POLARITY
        for name, record in _protection_records(fields):
            key = _key(ctx, surface, name)
            for field, ftype in sorted(fields[name].items()):
                if is_bool(ftype) and not field.startswith(prefixes):
                    yield _violation(
                        ctx,
                        "R41.bool.polarity",
                        surface,
                        key,
                        f"polarity:{field}",
                        f"`{name}.{field}` is a boolean; name what `true` allows (`allow...` / `lock...`)",
                    )
            if name != record and record in fields:
                for field in sorted(set(fields[name]) - set(fields[record])):
                    yield _violation(
                        ctx,
                        "R41.bool.polarity",
                        surface,
                        key,
                        f"name:{field}",
                        f"`{name}.{field}` is not a field of `{record}`",
                    )


# ---------------------------------------------------------------------------
# R42.unit.suffix
# ---------------------------------------------------------------------------

DIMENSION_WORDS = frozenset(
    {
        "width", "height", "margin", "margins", "indent", "rotation", "angle", "offset", "spacing", "gap",
        "thickness", "transparency", "zoom", "scale",
    }
)  # fmt: skip
# Word pairs that name a dimension together.
DIMENSION_PAIRS = frozenset({("font", "size"), ("marker", "size")})
UNIT_WORDS = frozenset(
    {
        "pt", "pts", "points", "px", "pixels", "emu", "emus", "in", "inch", "inches", "deg", "degrees",
        "chars", "characters", "twips", "pct", "percent", "cm", "mm",
    }
)  # fmt: skip
# Marker words a C patch or presence field appends to the name it describes.
_MARKER_WORDS = frozenset({"op", "engaged", "mask"})
_C_NUMERIC = frozenset({"int8_t", "int16_t", "int32_t", "int64_t", "uint8_t", "uint16_t", "uint32_t", "uint64_t",
                        "double", "float", "size_t"})  # fmt: skip


# Last words that make a name a count, an index or a selector rather than a measure.
_NOT_MEASURE_WORDS = frozenset({"count", "index", "id", "kind", "type", "mode"})


def _name_words(name: str) -> List[str]:
    return [w.lower() for w in _words(name)]


def _snake(name: str) -> str:
    return "_".join(w.lower() for w in _words(name))


def needs_unit(name: str, owner: str = "") -> bool:
    """`name` carries a dimension (or is a field of a margins record) and does not end with a unit word."""
    words = _name_words(name)
    # A presence flag or a patch op follows the field it describes, which is checked itself.
    if not words or words[0] == "has" or words[-1] in UNIT_WORDS | _MARKER_WORDS | _NOT_MEASURE_WORDS:
        return False
    if DIMENSION_WORDS & set(words) or any(pair in DIMENSION_PAIRS for pair in zip(words, words[1:])):
        return True
    return any(w in ("margin", "margins") for w in _name_words(owner))


def _unitless(ctx: Context) -> Set[str]:
    listed = ctx.ledger.data.get("unitless") or []
    return {_snake(n) for n in listed if isinstance(n, str)}


def _js_numeric(ptype: str) -> bool:
    return "number" in _strip_null(ptype, True)


def _py_numeric(annotation: str) -> bool:
    return bool({"int", "float"} & {a.strip() for a in _top_union(" | ".join(_strip_null(annotation, False)))})


@_owned("R42.unit.suffix")
def _rule_unit_suffix(ctx: Context) -> Iterable[Violation]:
    """A numeric field or parameter whose name carries a dimension ends with a unit word, unless the ledger's
    top-level `unitless:` list names it (in any casing)."""
    exempt = _unitless(ctx)

    def flagged(name: str, owner: str = "") -> bool:
        return needs_unit(name, owner) and _snake(name) not in exempt

    message = "`{}` carries a dimension; end it with its unit or list it under `unitless:`"
    for typedef in ctx.typedefs:
        if typedef.kind != "struct":
            continue
        for f in typedef.fields:
            if f.pointer_depth == 0 and f.base_type in _C_NUMERIC and flagged(f.name, typedef.name):
                yield ctx.c_violation("R42.unit.suffix", typedef.name, f.name, message.format(f.name), typedef.line)
    for fn in ctx.functions:
        for p in fn.params:
            if p.base_type in _C_NUMERIC and flagged(p.role_name):
                yield ctx.c_violation("R42.unit.suffix", fn.name, p.name, message.format(p.name), fn.line)
    found = surfaces(ctx)
    for surface in JS_SURFACES:
        for name, decl in sorted(found.js[surface].decls.items()):
            if decl.kind != "interface":
                continue
            key = _key(ctx, surface, name)
            for member in decl.members:
                if member.is_method:
                    for param, ptype in member.params:
                        if _js_numeric(ptype) and flagged(param):
                            disc = f"{member.name}.{param}"
                            yield _violation(ctx, "R42.unit.suffix", surface, key, disc, message.format(disc))
                elif _js_numeric(member.type) and flagged(member.name, name):
                    yield _violation(ctx, "R42.unit.suffix", surface, key, member.name, message.format(member.name))
    for name, cls in sorted(found.py.classes.items()):
        if name.startswith("_"):
            continue
        key = _key(ctx, "py", name)
        for field, annotation in sorted(cls.fields.items()):
            if not field.startswith("_") and _py_numeric(annotation) and flagged(field, name):
                yield _violation(ctx, "R42.unit.suffix", "py", key, field, message.format(field))
        for method, fn in sorted(cls.methods.items()):
            if method.startswith("_"):
                continue
            for param, annotation in fn.params:
                if _py_numeric(annotation) and flagged(param):
                    disc = f"{method}.{param}"
                    yield _violation(ctx, "R42.unit.suffix", "py", key, disc, message.format(disc))


# ---------------------------------------------------------------------------
# R43.c.status_doc
# ---------------------------------------------------------------------------

_STATUS_DOC_RE = re.compile(r"@return|\bcommon (?:[\w-]+ )?failures?\b", re.I)


def documents_status(fn: CFunction) -> bool:
    return _STATUS_DOC_RE.search(fn.doc) is not None


@_owned("R43.c.status_doc")
def _rule_status_doc(ctx: Context) -> Iterable[Violation]:
    """Every `FM_API` declaration returning a status documents its statuses with `@return`, or refers to the common failures stated
    at the head of its section."""
    for fn in ctx.functions:
        if fn.ret != "void" and not documents_status(fn):
            yield ctx.c_violation(
                "R43.c.status_doc", fn.name, NO_DISCRIMINATOR, "no `@return` and no common-failure reference", fn.line
            )


# ---------------------------------------------------------------------------
# R44.singleton.null
# ---------------------------------------------------------------------------


@_owned("R44.singleton.null")
def _rule_singleton_null(ctx: Context) -> Iterable[Violation]:
    """A singleton (a noun with no `add` / `create`, outside every collection, not a plural) has no `remove` or
    `clear` op: `set<X>(null)` removes it. A `reset` op restores the default of a noun that has a `set`."""
    in_collection = set(_collection_of(ctx))
    plurals = {c.get("plural") for c in ctx.ledger.collections.values() if isinstance(c, dict)}
    nouns: Dict[str, Dict[str, str]] = {}
    for key, _entry in _ops(ctx):
        stem, _, verb = key.rpartition(".")
        nouns.setdefault(stem, {})[verb] = key
    for stem, verbs in sorted(nouns.items()):
        if "add" in verbs or "create" in verbs or stem in plurals or stem.endswith("s"):
            continue
        for verb in ("remove", "clear"):
            key = verbs.get(verb)
            if key is not None and key not in in_collection:
                yield _ledger_violation(
                    "R44.singleton.null",
                    key,
                    verb,
                    f"`{stem}` is a singleton; remove it with `set` and `null`, not `{verb}`",
                )
        if "reset" in verbs and "set" not in verbs:
            yield _ledger_violation(
                "R44.singleton.null", verbs["reset"], "reset", f"`{stem}` has no `set` whose default `reset` restores"
            )
