"""API adoption rules: A1 in records, the ooxml layer, colour specs, record versioning, pointer lifetimes,
calculating reads and trailing defaults.

R29.record_a1      An input / patch record field typed `CellRect`, `SheetRect` or `CellAddress` also takes an
                   A1 string (an array field also takes one string), and no input record spells a cell as a
                   `<p>Row` / `<p>Col` pair, except the fields its `a1_exempt` lists.
R30.ooxml_namespace
                   An op the ledger marks `namespace: ooxml` is not a direct `Workbook` / `Sheet` member, and
                   no direct member exposes an XF index or an XML fragment unless marked `resolved: true`.
R31.color          A colour is a `fm_color_t` / `Color` (an input a `ColorInput`); a record or op the
                   ledger marks `color: resolved` is exempt.
R32.struct_size    A C struct passed by pointer that is not a frozen value type starts with
                   `uint32_t struct_size`.
R33.lifetime       A C function that outputs a pointer states `Lifetime: static`, `Lifetime: transient` or
                   `Lifetime: owned (until <declared fm_* function>)` in its doc.
R34.reads_values   The C definition of an op the ledger marks `reads_values: true` calls `ensure_calculated`,
                   and no other definition does.
R35.trailing_default
                   A parameter the ledger gives a `default` (`params: {<c name>: {default: <value>}}`) may be
                   omitted on every binding surface, as may every parameter after it.
"""

from __future__ import annotations

import ast
import re
from pathlib import Path
from typing import Dict, Iterable, Iterator, List, Optional, Sequence, Set, Tuple

from .api_conventions import NO_DISCRIMINATOR, CFunction, Context, CParam, Violation, rule
from .api_conventions_bindings import (
    _TS_HEAD_RE,
    BINDING_SURFACES,
    SHEET_CLASS,
    WORKBOOK_CLASS,
    _active,
    _cached,
    _callables,
    _entry_surfaces,
    _field,
    _in_tree,
    _lookup,
    _norm,
    _scan,
    _split_top,
    _top_union,
    _tree_c,
    _violation,
    _words,
    surfaces,
)

C_API_DIR = Path("src") / "c_api"


def _mentions(text: str, name: str) -> bool:
    return re.search(rf"(?<![\w$]){re.escape(name)}(?![\w$])", text) is not None


def _parse_annotation(text: str) -> Optional[ast.AST]:
    try:
        return ast.parse(text, mode="eval").body
    except SyntaxError:
        return None


def _py_alternatives(node: ast.AST) -> List[ast.AST]:
    """Union alternatives of a Python annotation node; a quoted annotation is parsed first."""
    if isinstance(node, ast.Constant) and isinstance(node.value, str):
        parsed = _parse_annotation(node.value)
        if parsed is None:
            return [node]
        node = parsed
    if isinstance(node, ast.BinOp) and isinstance(node.op, ast.BitOr):
        return _py_alternatives(node.left) + _py_alternatives(node.right)
    if isinstance(node, ast.Subscript) and ast.unparse(node.value).rpartition(".")[2] in ("Union", "Optional"):
        inner = node.slice.elts if isinstance(node.slice, ast.Tuple) else [node.slice]
        return [alt for item in inner for alt in _py_alternatives(item)]
    return [node]


def _py_names(nodes: Sequence[ast.AST]) -> List[str]:
    """Unqualified names of annotation alternatives, `None` left out."""
    return [ast.unparse(n).rpartition(".")[2] for n in nodes if not (isinstance(n, ast.Constant) and n.value is None)]


def _py_type_alternatives(declared: str) -> List[str]:
    node = _parse_annotation(declared) if declared else None
    return _py_names(_py_alternatives(node)) if node is not None else ([declared] if declared else [])


def _js_alternatives(declared: str) -> List[str]:
    return [a.strip() for a in _top_union(_norm(declared)) if a not in ("null", "undefined")]


# ---------------------------------------------------------------------------
# R29.record_a1
# ---------------------------------------------------------------------------

_RANGE_TYPES = ("CellRect", "SheetRect", "CellAddress")
_INPUT_ROLES = ("input", "patch")
_ARRAY_GENERICS = ("Array", "ReadonlyArray")
_PY_SEQUENCES = frozenset({"list", "List", "Sequence", "tuple", "Tuple", "Iterable", "Collection"})


def _js_element(alternative: str) -> Optional[str]:
    """The element type of a JS array alternative, or None."""
    text = alternative.strip()
    if text.startswith("readonly "):
        text = text[len("readonly ") :].strip()
    if text.endswith("[]"):
        inner = text[:-2].strip()
        return inner[1:-1] if inner.startswith("(") and inner.endswith(")") else inner
    for generic in _ARRAY_GENERICS:
        if text.startswith(generic + "<") and text.endswith(">"):
            return text[len(generic) + 1 : -1]
    return None


def _range_problem(alternatives: List[str], elements: Optional[List[str]], text_type: str) -> Optional[str]:
    """What an A1-capable range field lacks, given its top-level and (for an array) element alternatives."""
    if elements is None:
        return None if text_type in alternatives else f"also takes `{text_type}`"
    if text_type not in elements or text_type not in alternatives:
        return f"takes `{text_type}` elements and one `{text_type}`"
    return None


def _js_range_problem(declared: str) -> Optional[Tuple[str, str]]:
    """`(range type, problem)` of a JS field type naming a range type without the A1 string form."""
    found = next((t for t in _RANGE_TYPES if _mentions(declared, t)), None)
    if found is None:
        return None
    alternatives = _js_alternatives(declared)
    elements: Optional[List[str]] = None
    for alternative in alternatives:
        element = _js_element(alternative)
        if element is not None:
            elements = (elements or []) + [e.strip() for e in _top_union(element)]
    problem = _range_problem(alternatives, elements, "string")
    return (found, problem) if problem else None


def _py_range_problem(declared: str) -> Optional[Tuple[str, str]]:
    """`(range type, problem)` of a Python field annotation naming a range type without `str`."""
    found = next((t for t in _RANGE_TYPES if _mentions(declared, t)), None)
    root = _parse_annotation(declared) if found else None
    if root is None:
        return None
    nodes = _py_alternatives(root)
    elements: Optional[List[str]] = None
    for node in nodes:
        if isinstance(node, ast.Subscript) and ast.unparse(node.value).rpartition(".")[2] in _PY_SEQUENCES:
            item = node.slice.elts[0] if isinstance(node.slice, ast.Tuple) else node.slice
            elements = (elements or []) + _py_names(_py_alternatives(item))
    problem = _range_problem(_py_names(nodes), elements, "str")
    return (found, problem) if problem else None


def _record_members(ctx: Context, surface: str, entry: dict) -> Optional[Tuple[str, Dict[str, str]]]:
    """`(name, {field: declared type})` of a record interface or class on `surface`; None for a type alias such
    as `Patch<T>`, whose fields are the result record's."""
    found = surfaces(ctx)
    if surface == "py":
        name = _in_tree(found.py.classes, entry, "py")
        return (name, dict(found.py.classes[name].fields)) if name else None
    decls = found.js[surface].decls
    name = _in_tree(decls, entry, "js")
    if name is None or decls[name].kind != "interface":
        return None
    return name, {m.name: m.type for m in decls[name].members if not m.is_method}


def a1_exempt(entry: dict) -> Set[str]:
    """JS field names the entry's `a1_exempt: {fields: [...], reason: ...}` keeps numeric."""
    exempt = entry.get("a1_exempt")
    fields = exempt.get("fields") if isinstance(exempt, dict) else None
    return {f for f in fields if isinstance(f, str)} if isinstance(fields, list) else set()


def _cell_pairs(fields: Iterable[str], js: bool) -> List[str]:
    """Prefixes `p` of the `<p>Row` / `<p>Col` (`<p>_row` / `<p>_col`) field pairs, `p` non-empty."""
    names = set(fields)
    row, col = ("Row", "Col") if js else ("_row", "_col")
    return sorted(n[: -len(row)] for n in names if n.endswith(row) and n != row and n[: -len(row)] + col in names)


@rule("R29.record_a1")
def _rule_record_a1(ctx: Context) -> Iterable[Violation]:
    """An input or patch record field typed `CellRect`, `SheetRect` or `CellAddress` also takes an A1 string
    (an array field takes string elements and one string); an input record names a cell by a `CellAddress`
    field, not a `<p>Row` / `<p>Col` pair. The fields the entry's `a1_exempt: {fields: [...]}` lists are exempt.
    """
    if not _active(ctx):
        return
    for key, entry in sorted(ctx.ledger.entries.items()):
        if not isinstance(entry, dict) or entry.get("kind") != "record" or "replaced_by" in entry:
            continue
        if entry.get("role") not in _INPUT_ROLES:
            continue
        exempt = a1_exempt(entry)
        listed = _entry_surfaces(entry)
        for surface in BINDING_SURFACES:
            shape = _record_members(ctx, surface, entry) if listed is None or surface in listed else None
            if shape is None:
                continue
            name, fields = shape
            js = surface != "py"
            for field, declared in sorted(fields.items()):
                if _camel(field) in exempt:
                    continue
                found = _js_range_problem(declared) if js else _py_range_problem(declared)
                if found is not None:
                    yield _violation(
                        ctx, "R29.record_a1", surface, key, field, f"`{name}.{field}` is a `{found[0]}`; it {found[1]}"
                    )
            if entry.get("role") != "input":
                continue
            for prefix in _cell_pairs(fields, js):
                if not {_camel(prefix + "_row"), _camel(prefix + "_col")} <= exempt:
                    yield _violation(
                        ctx,
                        "R29.record_a1",
                        surface,
                        key,
                        prefix,
                        f"`{name}` spells the cell `{prefix}` as a row / column pair; take one `CellAddress` field",
                    )


# ---------------------------------------------------------------------------
# R30.ooxml_namespace
# ---------------------------------------------------------------------------

OOXML_NAMESPACE = "ooxml"
_XF_FIELDS = {
    "js": frozenset({"xfIndex", "xfId", "cellStyleXfIndex"}),
    "py": frozenset({"xf_index", "xf_id", "cell_style_xf_index"}),
}
_IDENT_RE = re.compile(r"[A-Za-z_$][\w$]*")


def _direct_members(ctx: Context, surface: str) -> Iterator[Tuple[str, str, object]]:
    """`(class, name, member)` of every public `Workbook` / `Sheet` method on `surface`."""
    found = surfaces(ctx)
    classes = found.py.class_methods() if surface == "py" else found.js[surface].classes()
    for cls in (WORKBOOK_CLASS, SHEET_CLASS):
        for name, member in sorted(classes.get(cls, {}).items()):
            if not name.rpartition(".")[2].startswith("_"):
                yield cls, name, member


def _type_fields(ctx: Context, surface: str) -> Dict[str, Dict[str, str]]:
    """`{type name: {field: declared type}}` of every interface (JS) or class (Python) on `surface`."""

    def build() -> Dict[str, Dict[str, str]]:
        found = surfaces(ctx)
        if surface == "py":
            return {name: dict(cls.fields) for name, cls in found.py.classes.items()}
        return {
            name: {m.name: m.type for m in decl.members if not m.is_method}
            for name, decl in found.js[surface].decls.items()
            if decl.kind == "interface"
        }

    return _cached(ctx, f"adoption-type-fields:{surface}", build)


def _xf_exposures(ctx: Context, surface: str, name: str, member) -> Iterator[Tuple[str, str]]:
    """`(discriminator, what)` of every way the direct member `name` exposes an XF index or an XML fragment."""
    forbidden = _XF_FIELDS[_field(surface)]
    words = _words(name.rpartition(".")[2])
    if "xf" in words or "xfs" in words:
        yield "name", "its name addresses an XF record"
    if words and words[-1] == "xml":
        yield "xml", "it reads or writes an XML fragment"
    for param, _ in member.params:
        if param in forbidden:
            yield param, f"it takes `{param}`"
    types = _type_fields(ctx, surface)
    returns = member.returns if surface == "py" else member.type
    for slot, declared in [*member.params, ("returns", returns)]:
        for ident in sorted(set(_IDENT_RE.findall(declared or ""))):
            for field in sorted(set(types.get(ident, {})) & forbidden):
                yield f"{slot}.{field}", f"`{ident}.{field}` reaches it through `{slot}`"


@rule("R30.ooxml_namespace")
def _rule_ooxml_namespace(ctx: Context) -> Iterable[Violation]:
    """An op the ledger marks `namespace: ooxml` is reached through `Workbook.ooxml` / `Sheet.ooxml`, never as
    a direct member; a direct member exposes no XF index (`xf` in its name, or an `xfIndex` / `xfId` /
    `cellStyleXfIndex` argument, returned field or argument field) and no XML fragment (`*Xml`), unless the
    ledger marks it `resolved: true`.
    """
    if not _active(ctx):
        return
    for surface in BINDING_SURFACES:
        for cls, name, member in _direct_members(ctx, surface):
            key, entry = _lookup(ctx, surface, name)
            key = key or name
            if entry.get("namespace") == OOXML_NAMESPACE:
                yield _violation(
                    ctx,
                    "R30.ooxml_namespace",
                    surface,
                    key,
                    "namespace",
                    f"`{cls}.{name}` belongs under `{cls}.{OOXML_NAMESPACE}`",
                )
                continue
            if entry.get("resolved") is True:
                continue
            for disc, what in _xf_exposures(ctx, surface, name, member):
                yield _violation(
                    ctx,
                    "R30.ooxml_namespace",
                    surface,
                    key,
                    disc,
                    f"`{cls}.{name}` is a direct member, but {what}; move it under `{cls}.{OOXML_NAMESPACE}`",
                )


# ---------------------------------------------------------------------------
# R31.color
# ---------------------------------------------------------------------------

COLOR_SPEC = {"c": "fm_color_t", "js": "Color", "py": "Color"}
COLOR_INPUT = "ColorInput"
RETIRED_COLOR = {"c": "fm_cf_color_t", "js": "CfColor", "py": "CfColor"}
RESOLVED = "resolved"
# Per naming style: a raw ARGB / RGB number, a colour-named slot, a presence flag.
_SNAKE_COLOR_RES = (
    re.compile(r"(?:^|_)a?rgb$"),
    re.compile(r"^colou?r(?:_[a-z]+)?$|_colou?r$"),
    re.compile(r"^has_(?:\w+_)?colou?r$"),
)
_CAMEL_COLOR_RES = (
    re.compile(r"(?:^a?rgb|Argb|Rgb)$"),
    re.compile(r"^colou?r(?:[A-Z][a-z]*)?$|Colou?r$"),
    re.compile(r"^has\w*Colou?r$"),
)
_NUMBER_TYPES = {"c": frozenset({"uint32_t"}), "js": frozenset({"number"}), "py": frozenset({"int", "float"})}


def _color_problem(name: str, alternatives: Sequence[str], field: str, takes_input: bool) -> Optional[str]:
    """What is wrong with one colour-carrying field or argument, or None."""
    raw_re, name_re, has_re = _CAMEL_COLOR_RES if field == "js" else _SNAKE_COLOR_RES
    spec = COLOR_SPEC[field]
    if any(_mentions(a, RETIRED_COLOR[field]) for a in alternatives):
        return f"`{name}` is a `{RETIRED_COLOR[field]}`; use `{spec}`"
    if raw_re.search(name):
        return f"`{name}` holds a colour as a number; carry one `{spec}` instead"
    if has_re.match(name):
        return f"`{name}` flags a colour's presence; an absent `{spec}` is kind None"
    if not name_re.search(name):
        return None
    if any(a in _NUMBER_TYPES[field] for a in alternatives):
        return f"`{name}` holds a colour as a number; make it a `{spec}`"
    if takes_input and spec in alternatives and COLOR_INPUT not in alternatives:
        return f"`{name}` is an input colour; take a `{COLOR_INPUT}`"
    return None


def _ledger_records(ctx: Context, field: str) -> Dict[str, Tuple[str, dict]]:
    """`{tree name: (ledger key, entry)}` of every ledger record by its current and former `field` name."""

    def build() -> Dict[str, Tuple[str, dict]]:
        out: Dict[str, Tuple[str, dict]] = {}
        for key, entry in sorted(ctx.ledger.entries.items()):
            if isinstance(entry, dict) and entry.get("kind") == "record":
                was = entry.get("was") if isinstance(entry.get("was"), dict) else {}
                for name in (entry.get(field), was.get(field)):
                    if isinstance(name, str):
                        out.setdefault(name, (key, entry))
        return out

    return _cached(ctx, f"adoption-records:{field}", build)


def _resolved(ctx: Context, field: str, name: str) -> bool:
    return _ledger_records(ctx, field).get(name, ("", {}))[1].get("color") == RESOLVED


def _c_color_violations(ctx: Context) -> Iterator[Violation]:
    for typedef in ctx.typedefs:
        if typedef.name == RETIRED_COLOR["c"]:
            message = f"`{typedef.name}` is retired; use `{COLOR_SPEC['c']}`"
            yield ctx.c_violation("R31.color", typedef.name, NO_DISCRIMINATOR, message, typedef.line)
            continue
        if typedef.kind != "struct" or typedef.name == COLOR_SPEC["c"] or _resolved(ctx, "c", typedef.name):
            continue
        for member in typedef.fields:
            problem = _color_problem(member.name, [member.base_type], "c", False)
            if problem:
                yield ctx.c_violation(
                    "R31.color", typedef.name, member.name, f"{typedef.name}: {problem}", typedef.line
                )
    for fn in ctx.functions:
        if ctx.ledger.entry(ctx.ledger.key_for("c", fn.name)).get("color") == RESOLVED:
            continue
        for param in fn.params:
            problem = _color_problem(param.role_name, [param.base_type], "c", False)
            if problem:
                yield ctx.c_violation("R31.color", fn.name, param.name, f"{fn.name}: {problem}", fn.line)


def _binding_color_violations(ctx: Context, surface: str) -> Iterator[Violation]:
    field = _field(surface)
    alternatives_of = _py_type_alternatives if field == "py" else _js_alternatives
    records = _ledger_records(ctx, field)
    retired = RETIRED_COLOR[field]
    types = _type_fields(ctx, surface)
    for name, fields in sorted(types.items()):
        key, entry = records.get(name, (name, {}))
        if name == retired:
            yield _violation(ctx, "R31.color", surface, key, NO_DISCRIMINATOR, f"`{name}` is retired; use `Color`")
            continue
        if name == COLOR_SPEC[field] or entry.get("color") == RESOLVED:
            continue
        for member, declared in sorted(fields.items()):
            problem = _color_problem(member, alternatives_of(declared), field, entry.get("role") in _INPUT_ROLES)
            if problem:
                yield _violation(ctx, "R31.color", surface, key, f"{name}.{member}", f"{name}: {problem}")
    for cls, name, member in _direct_members(ctx, surface):
        key, op = _lookup(ctx, surface, name)
        if op.get("color") == RESOLVED:
            continue
        for param, declared in member.params:
            problem = _color_problem(param, alternatives_of(declared), field, True)
            if problem:
                yield _violation(ctx, "R31.color", surface, key or name, param, f"{cls}.{name}: {problem}")


@rule("R31.color")
def _rule_color(ctx: Context) -> Iterable[Violation]:
    """A colour crosses every surface as a colour spec: `fm_color_t` in C, `Color` in the bindings, and a
    `ColorInput` where an input record field or a method argument takes one. No `fm_cf_color_t` / `CfColor`,
    no `*_argb` / `*_rgb` / `*Argb` / `*Rgb` sibling, no `has_*color` flag, no colour held as a bare number.
    A record or op the ledger marks `color: resolved` (a rendered result) is exempt.
    """
    yield from _c_color_violations(ctx)
    if not _active(ctx):
        return
    for surface in BINDING_SURFACES:
        yield from _binding_color_violations(ctx, surface)


# ---------------------------------------------------------------------------
# R32.struct_size
# ---------------------------------------------------------------------------

STRUCT_SIZE = CParam("uint32_t", "struct_size")
# Frozen value types: array elements passed as `const T*` with a count, the coordinate / value / calendar /
# colour primitives, and structs embedded by value. They carry no `struct_size`.
VALUE_STRUCTS = frozenset(
    {
        "fm_cell_rect_t", "fm_sheet_rect_t", "fm_value_t", "fm_civil_time_t",
        "fm_color_t", "fm_phonetic_run_t", "fm_gradient_stop_t", "fm_mention_t", "fm_date_group_item_t",
        "fm_filter_column_t", "fm_sort_condition_t", "fm_sparkline_t", "fm_cfvo_t", "fm_cf_icon_t",
        "fm_drawing_anchor_t",
        "fm_font_t", "fm_fill_t", "fm_border_t", "fm_border_side_t", "fm_cell_xf_t",
        "fm_dxf_t", "fm_cell_style_t", "fm_table_style_element_t",
        "fm_chart_data_source_t", "fm_chart_title_t", "fm_chart_line_t", "fm_chart_line_patch_t", "fm_chart_fill_t",
        "fm_chart_fill_patch_t", "fm_chart_marker_t", "fm_chart_marker_patch_t", "fm_chart_labels_t",
        "fm_chart_labels_patch_t", "fm_cell_address_t", "fm_xf_remap_t", "fm_cell_xf_index_t",
    }
)  # fmt: skip


@rule("R32.struct_size")
def _rule_struct_size(ctx: Context) -> Iterable[Violation]:
    """A record (a struct some function takes or fills through a pointer, other than the frozen value types)
    starts with `uint32_t struct_size`.
    """
    structs = {t.name: t for t in ctx.typedefs if t.kind == "struct"}
    passed = {p.base_type for fn in ctx.functions for p in fn.params if p.pointer_depth >= 1 and p.base_type in structs}
    for name in sorted(passed - VALUE_STRUCTS):
        typedef = structs[name]
        if not typedef.fields or typedef.fields[0] != STRUCT_SIZE:
            yield ctx.c_violation(
                "R32.struct_size",
                name,
                NO_DISCRIMINATOR,
                f"{name} is a record passed by pointer; its first field is `uint32_t struct_size`",
                typedef.line,
            )


# ---------------------------------------------------------------------------
# R33.lifetime
# ---------------------------------------------------------------------------

LIFETIME_WORDS = ("static", "transient", "owned (until <fm_* function>)")
_LIFETIME_TAG_RE = re.compile(r"\bLifetime:\s*")
_LIFETIME_RE = re.compile(r"(?:static|transient|owned\s*\(\s*until\s+(fm_\w+)\s*\))(?![\w-])")


def _carries_pointer(ctx: Context, struct: str, seen: Tuple[str, ...] = ()) -> bool:
    """True when `struct` holds a pointer, directly or in a struct it embeds by value."""
    structs = _cached(ctx, "adoption-structs", lambda: {t.name: t for t in ctx.typedefs if t.kind == "struct"})
    typedef = structs.get(struct)
    if typedef is None or struct in seen:
        return False
    return any(m.pointer_depth > 0 or _carries_pointer(ctx, m.base_type, (*seen, struct)) for m in typedef.fields)


def outputs_pointer(ctx: Context, fn: CFunction) -> bool:
    """True when `fn` hands a pointer back: a pointer return, a `T**` output, or an output record holding one."""
    if "*" in fn.ret:
        return True
    for param in fn.params:
        compact = param.ctype.replace(" ", "")
        if param.pointer_depth >= 2 and "*const*" not in compact:
            return True
        if param.pointer_depth == 1 and not compact.startswith("const") and _carries_pointer(ctx, param.base_type):
            return True
    return False


def _lifetime_problem(doc: str, declared: Set[str]) -> Optional[Tuple[str, str]]:
    tags = list(_LIFETIME_TAG_RE.finditer(doc))
    if not tags:
        return NO_DISCRIMINATOR, "its doc states no `Lifetime:`"
    for tag in tags:
        match = _LIFETIME_RE.match(doc, tag.end())
        if match is None:
            shown = doc[tag.start() : tag.end() + 24].strip()
            return "malformed", f"`{shown}...` is not one of: {', '.join(LIFETIME_WORDS)}"
        if match.group(1) and match.group(1) not in declared:
            return "malformed", f"`Lifetime: owned` names `{match.group(1)}`, which is not declared"
    return None


@rule("R33.lifetime")
def _rule_lifetime(ctx: Context) -> Iterable[Violation]:
    """A C function that hands a pointer back (a pointer return, a `T**` output, or an output record holding a
    pointer) states its lifetime in its doc as `Lifetime: static`, `Lifetime: transient` or
    `Lifetime: owned (until <fm_* function>)` naming a declared function.
    """
    declared = {fn.name for fn in ctx.functions}
    for fn in ctx.functions:
        problem = _lifetime_problem(fn.doc, declared) if outputs_pointer(ctx, fn) else None
        if problem:
            message = f"{fn.name} outputs a pointer, but {problem[1]}"
            yield ctx.c_violation("R33.lifetime", fn.name, problem[0], message, fn.line)


# ---------------------------------------------------------------------------
# R34.reads_values
# ---------------------------------------------------------------------------

GATE_CALL = "ensure_calculated"
_GATE_RE = re.compile(rf"\b{GATE_CALL}\s*\(")
_DEFINITION_RE = re.compile(r'extern\s+"C"\s+[^;{}()"]*?\b(fm_\w+)\s*\(')
_BLANKED_RE = re.compile(r"//[^\n]*|/\*.*?\*/|\"(?:\\.|[^\"\\\n])*\"|'(?:\\.|[^'\\\n])*'", re.S)


def _blank(text: str) -> str:
    """`text` with comments and literals replaced by spaces, offsets and newlines kept."""
    return _BLANKED_RE.sub(lambda m: re.sub(r"[^\n]", " ", m.group(0)), text)


def _c_scan(clean: str, start: int, stop: str) -> int:
    """Index of the first `stop` at bracket depth 0 from `start` in blanked C++, or `len(clean)`.

    Unlike `_scan`, angle brackets are not brackets here: `->`, `<` and `>` are operators in a body.
    """
    depth = 0
    for i in range(start, len(clean)):
        char = clean[i]
        if depth == 0 and char == stop:
            return i
        if char in "([{":
            depth += 1
        elif char in ")]}":
            depth -= 1
    return len(clean)


def c_definitions(root: Path) -> Dict[str, Tuple[str, int, str]]:
    """`{fm_* name: (path, line, body)}` of the `extern "C"` definitions under src/c_api of `root`."""
    out: Dict[str, Tuple[str, int, str]] = {}
    base = root / C_API_DIR
    for path in sorted(base.rglob("*.cpp")) if base.is_dir() else ():
        raw = path.read_text(encoding="utf-8")
        clean = _blank(raw)
        for match in _DEFINITION_RE.finditer(raw):
            brace = _c_scan(clean, match.end(), ")") + 1
            while brace < len(clean) and clean[brace].isspace():
                brace += 1
            if brace >= len(clean) or clean[brace] != "{":
                continue
            end = _c_scan(clean, brace + 1, "}")
            line = raw.count("\n", 0, match.start()) + 1
            out.setdefault(match.group(1), (str(path.relative_to(root)), line, clean[brace : end + 1]))
    return out


@rule("R34.reads_values")
def _rule_reads_values(ctx: Context) -> Iterable[Violation]:
    """The C definition of an op the ledger marks `reads_values: true` calls `ensure_calculated`, and a
    definition that calls it belongs to such an op.
    """
    if not _active(ctx):
        return
    definitions = _cached(ctx, "adoption-c-definitions", lambda: c_definitions(ctx.repo_root))
    gated: Set[str] = set()
    for key, entry in sorted(ctx.ledger.entries.items()):
        if not isinstance(entry, dict) or entry.get("reads_values") is not True:
            continue
        fn = _tree_c(ctx, entry)
        name = fn.name if fn else entry.get("c")
        if not isinstance(name, str):
            continue
        gated.add(name)
        found = definitions.get(name)
        if found is None:
            message = f"{name} reads values but has no definition under {C_API_DIR}"
            yield Violation("R34.reads_values", "c", key, "definition", message, ctx.site(fn.line) if fn else "")
        elif not _GATE_RE.search(found[2]):
            message = f"{name} reads values but does not call `{GATE_CALL}`"
            yield Violation("R34.reads_values", "c", key, "gate", message, f"{found[0]}:{found[1]}")
    for name, (path, line, body) in sorted(definitions.items()):
        if name not in gated and _GATE_RE.search(body):
            message = f"{name} calls `{GATE_CALL}` but its ledger op lacks `reads_values: true`"
            yield Violation("R34.reads_values", "c", ctx.op_key(name), "undeclared", message, f"{path}:{line}")


# ---------------------------------------------------------------------------
# R35.trailing_default
# ---------------------------------------------------------------------------


def _camel(snake: str) -> str:
    words = [w for w in snake.split("_") if w]
    return words[0] + "".join(w[:1].upper() + w[1:] for w in words[1:]) if words else ""


def ledger_defaults(entry: dict) -> Dict[str, object]:
    """`{C parameter name: default}` of an op's `params` slots that carry a `default`."""
    params = entry.get("params")
    if not isinstance(params, dict):
        return {}
    return {name: slot["default"] for name, slot in params.items() if isinstance(slot, dict) and "default" in slot}


def _js_signatures(body: str, method: str) -> List[List[Tuple[str, bool]]]:
    """`[(param, optional), ...]` of every signature of `method` in an interface body."""
    out: List[List[Tuple[str, bool]]] = []
    for chunk in _split_top(body, ";,"):
        head = _TS_HEAD_RE.match(chunk)
        if head is None or head.group(2) != method:
            continue
        rest = chunk[head.end() :]
        if rest.startswith("<"):
            rest = rest[_scan(rest, 1, ">") + 1 :].lstrip()
        if not rest.startswith("("):
            continue
        params = []
        for raw in _split_top(rest[1 : _scan(rest, 1, ")")], ","):
            match = re.match(r"(\.\.\.)?([A-Za-z_$][\w$]*)\s*(\?)?", raw)
            if match:
                params.append((match.group(2), bool(match.group(1) or match.group(3))))
        out.append(params)
    return out


def _js_omittable(signatures: List[List[Tuple[str, bool]]], param: str) -> Optional[bool]:
    """Whether some signature lets `param` and every later parameter be omitted (an optional tail, or another
    signature that stops right before it); None when no signature has `param`."""
    lists = [[p for p, _ in sig] for sig in signatures]
    holding = [sig for sig in signatures if any(p == param for p, _ in sig)]
    if not holding:
        return None
    for sig in holding:
        at = [p for p, _ in sig].index(param)
        if all(optional for _, optional in sig[at:]) or [p for p, _ in sig[:at]] in lists:
            return True
    return False


def _py_number(text: str) -> Optional[float]:
    try:
        return float(text)
    except ValueError:
        return None


def _py_omittable_problem(fn, param: str, default: object) -> Optional[str]:
    """Why the Python `param` of `fn` cannot be omitted, or contradicts the ledger's numeric default; None if fine."""
    names = [p for p, _ in fn.params]
    if param not in names:
        return None
    defaults = dict(fn.defaults)
    if any(p not in defaults and p not in fn.kwonly for p in names[names.index(param) :]):
        return f"`{param}` and the parameters after it need defaults"
    given = defaults.get(param, "None")
    if isinstance(default, (int, float)) and not isinstance(default, bool) and given != "None":
        if _py_number(given) != float(default):
            return f"`{param}` defaults to {given}"
    return None


@rule("R35.trailing_default")
def _rule_trailing_default(ctx: Context) -> Iterable[Violation]:
    """A parameter the ledger gives a `default` (`params: {<c name>: {default: <value>}}`) may be omitted on
    every binding surface, together with every parameter after it; a Python numeric default equals the
    ledger's (or is `None`).
    """
    if not _active(ctx):
        return
    found = surfaces(ctx)
    for key, entry in sorted(ctx.ledger.entries.items()):
        defaults = ledger_defaults(entry) if isinstance(entry, dict) else {}
        listed = _entry_surfaces(entry) if defaults else None
        for surface in BINDING_SURFACES if defaults else ():
            pool = _callables(ctx, surface)
            name = _in_tree(pool, entry, _field(surface)) if listed is None or surface in listed else None
            if name is None:
                continue
            for c_param, default in sorted(defaults.items()):
                if surface == "py":
                    problem = _py_omittable_problem(pool[name], c_param, default)
                else:
                    js = found.js[surface]
                    bodies = [js.decls[c].text for c in (SHEET_CLASS, WORKBOOK_CLASS) if c in js.decls]
                    signatures = [s for body in bodies for s in _js_signatures(body, name)]
                    omittable = _js_omittable(signatures, _camel(c_param))
                    problem = (
                        f"`{_camel(c_param)}` and the parameters after it must be optional"
                        if omittable is False
                        else None
                    )
                if problem:
                    message = f"{name}: {problem} (ledger default {default!r})"
                    yield _violation(ctx, "R35.trailing_default", surface, key, c_param, message)
