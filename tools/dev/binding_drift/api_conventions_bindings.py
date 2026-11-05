"""Ledger-bound API convention rules: the tree's surfaces judged against `tools/api/names.yaml`.

The rules here read the ledger's entries and collections alongside the C
header, both TypeScript declaration files and the Python package (parsed with
`ast`, never imported). Without a ledger they report nothing and leave a skip
note, so a header-only context sees only the ledger-free rules.

Each violation lands on the surface whose file has to change: `c` for the
header, `wasm` / `native` for the two declaration files, `py` for the package.
A binding name the ledger marks `drop:` is reported under the rule-id its
`closes:` names (default `R5.list`) and by no other rule.
"""

from __future__ import annotations

import ast
import re
from pathlib import Path
from typing import Dict, Iterable, Iterator, List, NamedTuple, Optional, Sequence, Set, Tuple

from .api_conventions import (
    ACRONYMS,
    CELL_RECT,
    CELL_RECT_FIELDS,
    NO_DISCRIMINATOR,
    SCOPES,
    CFunction,
    Context,
    CParam,
    Violation,
    c_verb,
    derive_js,
    derive_py,
    derive_type,
    rule,
)
from .surface_files import _SKIPPED, NODE_DTS, PYTHON_PKG_DIR, REPO_ROOT, WASM_DTS
from .ts_decl import _extract_braced_block

JS_SURFACES = ("wasm", "native")
BINDING_SURFACES = ("wasm", "native", "py")
SHEET_RECT = "fm_sheet_rect_t"
WORKBOOK_HANDLE = "fm_workbook_t"
DEFAULT_DROP_RULE = "R5.list"

# Synonyms whose left side must not appear in a name (R4); a pair is two adjacent words.
SYNONYMS = {("horizontal", "break"): "row_break", ("vertical", "break"): "col_break", ("idx",): "index"}

# Names a multi-range argument or field carries while it is still A1 text or a non-canonical list.
_A1_MULTI_NAMES = frozenset({"sqref", "ranges_a1", "ranges"})
_JS_A1_MULTI_NAMES = frozenset({"sqref", "rangesA1", "ranges"})
_TEXT_VERBS = frozenset({"parse", "format", "resolve", "convert", "localize", "canonicalize", "evaluate"})
_SHEET_FIELDS = frozenset({"sheet", "sheet_index"})
# Scopes whose `_at` / `_count` address a ledger collection; any other scope is a result handle.
_COLLECTION_SCOPES = ("workbook", "sheet", "process")
_EDGE_QUADS = (CELL_RECT_FIELDS, ("top", "left", "bottom", "right"), ("row", "col", "last_row", "last_col"))
_JS_EDGE_QUADS = (("firstRow", "firstCol", "lastRow", "lastCol"), ("top", "left", "bottom", "right"))

_SCALAR_RESULTS = {"string": "StringResult", "number": "NumberResult", "int": "NumberResult", "double": "NumberResult"}
_PY_SCALARS = {"string": ("str",), "number": ("int", "float"), "int": ("int",), "double": ("float",)}
_VALUE_ENVELOPES = ("StringResult", "NumberResult", "BoolResult")
_PATCH_ALIAS = "{ [K in keyof T]?: T[K] | null }"


# ---------------------------------------------------------------------------
# TypeScript declarations.
# ---------------------------------------------------------------------------


class TsMember(NamedTuple):
    name: str
    is_method: bool
    params: Tuple[Tuple[str, str], ...]
    type: str  # return type of a method, declared type of a field
    readonly: bool


class TsDecl(NamedTuple):
    name: str
    kind: str  # interface | type | enum | const | function | class
    text: str  # interface / enum body, type alias right-hand side, const annotation
    members: Tuple[TsMember, ...]


_TS_COMMENT_RE = re.compile(r"/\*.*?\*/|//[^\n]*", re.S)
_TS_DECL_RE = re.compile(
    r"^export\s+(?:declare\s+)?(?:default\s+)?(interface|type|enum|const|function|class)\s+([A-Za-z_$][\w$]*)",
    re.M,
)
_TS_DEFAULT_OBJECT_RE = re.compile(r"^declare\s+const\s+_default\s*:\s*\{", re.M)
_TS_HEAD_RE = re.compile(r"(readonly\s+)?([A-Za-z_$][\w$]*|\[[^\]]+\])\??\s*")
_TS_PARAM_RE = re.compile(r"(?:\.\.\.)?([A-Za-z_$][\w$]*|\{[^}]*\}|\[[^\]]*\])\??\s*(?::\s*(.*))?$", re.S)
_OPEN, _CLOSE = "([{<", ")]}>"


def _scan(text: str, start: int, stops: str) -> int:
    """Index of the first `stops` character at bracket depth 0 from `start`, or `len(text)`."""
    depth = 0
    i = start
    quote = ""
    while i < len(text):
        char = text[i]
        if quote:
            if char == "\\":
                i += 1
            elif char == quote:
                quote = ""
        elif char in "'\"`":
            quote = char
        elif char == "=" and text.startswith("=>", i):
            i += 1
        elif depth == 0 and char in stops:
            return i
        elif char in _OPEN:
            depth += 1
        elif char in _CLOSE:
            depth -= 1
        i += 1
    return len(text)


def _split_top(text: str, sep: str) -> List[str]:
    parts: List[str] = []
    start = 0
    while start <= len(text):
        end = _scan(text, start, sep)
        parts.append(text[start:end])
        start = end + 1
    return [part.strip() for part in parts if part.strip()]


def _norm(text: str) -> str:
    return " ".join(text.split())


def _parse_member(chunk: str) -> Optional[TsMember]:
    head = _TS_HEAD_RE.match(chunk)
    if not head:
        return None
    readonly, name = bool(head.group(1)), head.group(2)
    rest = chunk[head.end() :]
    if rest.startswith("<"):
        rest = rest[_scan(rest, 1, ">") + 1 :].lstrip()
    if rest.startswith("("):
        close = _scan(rest, 1, ")")
        params = []
        for raw in _split_top(rest[1:close], ","):
            match = _TS_PARAM_RE.match(raw)
            if match:
                params.append((match.group(1), _norm(match.group(2) or "")))
        returns = rest[close + 1 :].strip()
        return TsMember(name, True, tuple(params), _norm(returns[1:] if returns.startswith(":") else returns), readonly)
    if rest.startswith(":"):
        return TsMember(name, False, (), _norm(rest[1:]), readonly)
    return None


def _parse_members(body: str) -> Tuple[TsMember, ...]:
    members = (_parse_member(chunk) for chunk in _split_top(body, ";,"))
    return tuple(m for m in members if m is not None)


class JsSurface:
    """The exported declarations of one TypeScript declaration file."""

    def __init__(self, text: str) -> None:
        text = _TS_COMMENT_RE.sub(" ", text)
        self.decls: Dict[str, TsDecl] = {}
        for match in _TS_DECL_RE.finditer(text):
            kind, name = match.groups()
            after = match.end()
            if kind in ("interface", "enum", "class"):
                brace = text.find("{", after)
                body = _extract_braced_block(text, brace + 1) if brace >= 0 else ""
                members = _parse_members(body) if kind != "enum" else ()
                decl = TsDecl(name, kind, body, members)
            elif kind == "function":
                end = _scan(text, after, ";\n")
                member = _parse_member(text[match.start(2) : end])
                decl = TsDecl(name, kind, "", (member,) if member else ())
            else:
                end = _scan(text, after, ";")
                rhs = text[after:end].strip()
                if kind == "type":
                    rhs = rhs[_scan(rhs, 0, "=") + 1 :]
                elif rhs.startswith(":"):
                    rhs = rhs[1:]
                decl = TsDecl(name, kind, _norm(rhs), ())
            self.decls.setdefault(name, decl)
        default = _TS_DEFAULT_OBJECT_RE.search(text)
        self.default_object: Tuple[TsMember, ...] = (
            _parse_members(_extract_braced_block(text, default.end())) if default else ()
        )

    def interface_members(self, name: str) -> Dict[str, TsMember]:
        decl = self.decls.get(name)
        return {m.name: m for m in decl.members} if decl and decl.kind == "interface" else {}

    @property
    def workbook(self) -> Dict[str, TsMember]:
        return {n: m for n, m in self.interface_members("Workbook").items() if m.is_method}

    @property
    def statics(self) -> Dict[str, TsMember]:
        return {n: m for n, m in self.interface_members("WorkbookCtor").items() if m.is_method}

    @property
    def module_object(self) -> Dict[str, TsMember]:
        """Members of the module object: `FormulonModule`, else the default export's object type."""
        members = self.interface_members("FormulonModule")
        return members or {m.name: m for m in self.default_object}

    @property
    def functions(self) -> Dict[str, TsMember]:
        """Free functions: methods of `FormulonModule`, else top-level `export function`s."""
        module = {n: m for n, m in self.interface_members("FormulonModule").items() if m.is_method}
        exported = {d.name: d.members[0] for d in self.decls.values() if d.kind == "function" and d.members}
        return {**exported, **module}

    def callables(self) -> Dict[str, TsMember]:
        """Instance methods, free functions and `Workbook.<static>` factories by tree name."""
        out = dict(self.functions)
        out.update(self.workbook)
        out.update({f"Workbook.{n}": m for n, m in self.statics.items()})
        return out

    def symbols(self) -> Dict[str, str]:
        """Every exported symbol of the file, by the name the ledger's `js` field uses."""
        out = {name: "type" for name in self.decls}
        out.update({name: "function" for name in self.functions})
        out.update({name: "method" for name in self.workbook})
        out.update({f"Workbook.{name}": "static" for name in self.statics})
        return out


# ---------------------------------------------------------------------------
# Python package.
# ---------------------------------------------------------------------------


class PyFunc(NamedTuple):
    name: str
    params: Tuple[Tuple[str, str], ...]
    returns: str
    decorators: Tuple[str, ...]


class PyClass(NamedTuple):
    name: str
    bases: Tuple[str, ...]
    decorators: Tuple[str, ...]
    keywords: Dict[str, str]
    fields: Dict[str, str]
    methods: Dict[str, PyFunc]


def _py_func(node: ast.AST, method: bool) -> PyFunc:
    args = node.args  # type: ignore[attr-defined]
    positional = [*args.posonlyargs, *args.args]
    if method and positional and positional[0].arg in ("self", "cls"):
        positional = positional[1:]
    params = tuple(
        (a.arg, ast.unparse(a.annotation) if a.annotation is not None else "") for a in (*positional, *args.kwonlyargs)
    )
    returns = ast.unparse(node.returns) if node.returns is not None else ""  # type: ignore[attr-defined]
    decorators = tuple(ast.unparse(d) for d in node.decorator_list)  # type: ignore[attr-defined]
    return PyFunc(node.name, params, returns, decorators)  # type: ignore[attr-defined]


def _py_class(node: ast.ClassDef) -> PyClass:
    fields: Dict[str, str] = {}
    methods: Dict[str, PyFunc] = {}
    for item in node.body:
        if isinstance(item, ast.AnnAssign) and isinstance(item.target, ast.Name):
            fields[item.target.id] = ast.unparse(item.annotation)
        elif isinstance(item, (ast.FunctionDef, ast.AsyncFunctionDef)):
            methods[item.name] = _py_func(item, method=True)
            if item.name == "__init__":
                for sub in ast.walk(item):
                    targets = sub.targets if isinstance(sub, ast.Assign) else [getattr(sub, "target", None)]
                    for target in targets:
                        if (
                            isinstance(target, ast.Attribute)
                            and isinstance(target.value, ast.Name)
                            and target.value.id == "self"
                        ):
                            fields.setdefault(target.attr, "")
    return PyClass(
        node.name,
        tuple(ast.unparse(b) for b in node.bases),
        tuple(ast.unparse(d) for d in node.decorator_list),
        {kw.arg: ast.unparse(kw.value) for kw in node.keywords if kw.arg},
        fields,
        methods,
    )


def _all_names(tree: ast.AST) -> List[str]:
    names: List[str] = []
    for node in getattr(tree, "body", []):
        target = None
        if isinstance(node, ast.Assign) and len(node.targets) == 1:
            target = node.targets[0]
        elif isinstance(node, (ast.AugAssign, ast.AnnAssign)):
            target = node.target
        if isinstance(target, ast.Name) and target.id == "__all__" and node.value is not None:
            if isinstance(node.value, (ast.List, ast.Tuple)):
                names.extend(
                    e.value for e in node.value.elts if isinstance(e, ast.Constant) and isinstance(e.value, str)
                )
    return names


class PySurface:
    """The Python package's exports, `Workbook`'s methods and every top-level class and function."""

    def __init__(self, sources: Dict[str, str]) -> None:
        self.classes: Dict[str, PyClass] = {}
        self.functions: Dict[str, PyFunc] = {}
        self.exports: List[str] = []
        self.problems: List[str] = []
        for file_name, text in sorted(sources.items()):
            try:
                tree = ast.parse(text, filename=file_name)
            except SyntaxError as exc:
                self.problems.append(f"api-conventions: cannot parse {file_name}: {exc}")
                continue
            if file_name == "__init__.py":
                self.exports = _all_names(tree)
            for node in tree.body:
                if isinstance(node, ast.ClassDef):
                    self.classes.setdefault(node.name, _py_class(node))
                elif isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                    self.functions.setdefault(node.name, _py_func(node, method=False))

    @property
    def workbook(self) -> Dict[str, PyFunc]:
        """Public `Workbook` methods; class and static methods as `Workbook.<name>`."""
        cls = self.classes.get("Workbook")
        out: Dict[str, PyFunc] = {}
        for name, fn in cls.methods.items() if cls else ():
            if name.startswith("_"):
                continue
            static = any(d in ("classmethod", "staticmethod") for d in fn.decorators)
            out[f"Workbook.{name}" if static else name] = fn
        return out

    def callables(self) -> Dict[str, PyFunc]:
        out = {name: fn for name, fn in self.functions.items() if name in self.exports}
        out.update(self.workbook)
        return out

    def symbols(self) -> Dict[str, str]:
        out = {name: "export" for name in self.exports}
        out.update({name: "method" for name in self.workbook})
        return out


# ---------------------------------------------------------------------------
# Per-context model.
# ---------------------------------------------------------------------------


class Surfaces(NamedTuple):
    js: Dict[str, JsSurface]
    py: PySurface
    sites: Dict[str, str]


_JS_FILES = {"wasm": WASM_DTS, "native": NODE_DTS}


def _read_optional(path: Path, root: Path) -> str:
    if path.is_file():
        return path.read_text(encoding="utf-8")
    if root == REPO_ROOT:
        _note(f"api-conventions: {path.relative_to(root)} is absent; its surface was not checked")
    return ""


def _note(text: str) -> None:
    if text not in _SKIPPED:
        _SKIPPED.append(text)


def surfaces(ctx: Context) -> Surfaces:
    """The parsed binding surfaces of `ctx.repo_root`, built once per context."""
    cached = ctx.__dict__.get("_binding_surfaces")
    if cached is not None:
        return cached
    root = ctx.repo_root
    js = {}
    sites = {"c": ctx.header_name}
    for surface, path in _JS_FILES.items():
        local = root / path.relative_to(REPO_ROOT)
        js[surface] = JsSurface(_read_optional(local, root))
        sites[surface] = str(path.relative_to(REPO_ROOT))
    package = root / PYTHON_PKG_DIR.relative_to(REPO_ROOT)
    if not package.is_dir() and root == REPO_ROOT:
        _note(f"api-conventions: {PYTHON_PKG_DIR.relative_to(REPO_ROOT)} is absent; its surface was not checked")
    sources = {p.name: p.read_text(encoding="utf-8") for p in sorted(package.glob("*.py"))} if package.is_dir() else {}
    sites["py"] = str(PYTHON_PKG_DIR.relative_to(REPO_ROOT))
    built = Surfaces(js, PySurface(sources), sites)
    ctx.__dict__["_binding_surfaces"] = built
    return built


def _active(ctx: Context) -> bool:
    if not ctx.ledger.present:
        _note("api-conventions: no name ledger; the ledger-bound rules were not checked")
        return False
    return True


def _cached(ctx: Context, name: str, build):
    cache = ctx.__dict__.setdefault("_binding_cache", {})
    if name not in cache:
        cache[name] = build()
    return cache[name]


def _symbols(ctx: Context, surface: str) -> Dict[str, str]:
    def build() -> Dict[str, str]:
        if surface == "c":
            return _c_symbols(ctx)
        found = surfaces(ctx)
        return found.py.symbols() if surface == "py" else found.js[surface].symbols()

    return _cached(ctx, f"symbols:{surface}", build)


def _callables(ctx: Context, surface: str) -> Dict[str, object]:
    found = surfaces(ctx)
    return _cached(
        ctx,
        f"callables:{surface}",
        lambda: found.py.callables() if surface == "py" else found.js[surface].callables(),
    )


_C_ENUM_BODY_RE = re.compile(r"typedef\s+enum\s*\{([^}]*)\}")
_C_DEFINE_RE = re.compile(r"^#define\s+(FM_[A-Z0-9_]+)\s+\(?-?(?:0[xX][0-9A-Fa-f]+|\d+)[uUlL]*\)?\s*$", re.M)


def _c_symbols(ctx: Context) -> Dict[str, str]:
    out = {fn.name: "function" for fn in ctx.functions}
    out.update({t.name: "typedef" for t in ctx.typedefs})
    header = re.sub(r"/\*.*?\*/|//[^\n]*", " ", ctx.header, flags=re.S)
    for body in _C_ENUM_BODY_RE.findall(header):
        for member in body.split(","):
            name = re.match(r"\s*([A-Za-z_]\w*)", member)
            if name:
                out[name.group(1)] = "constant"
    out.update({name: "constant" for name in _C_DEFINE_RE.findall(header)})
    return out


def _field(surface: str) -> str:
    return "c" if surface == "c" else "py" if surface == "py" else "js"


def _lookup(ctx: Context, surface: str, symbol: str) -> Tuple[Optional[str], dict]:
    key = ctx.ledger.key_for(_field(surface), symbol)
    return key, ctx.ledger.entry(key)


def _key(ctx: Context, surface: str, symbol: str) -> str:
    return ctx.ledger.key_for(_field(surface), symbol) or symbol


def _names(entry: dict, field: str) -> List[str]:
    """The entry's current and former name for `field`."""
    was = entry.get("was") if isinstance(entry.get("was"), dict) else {}
    return [n for n in (entry.get(field), was.get(field)) if isinstance(n, str)]


def _dropped(entry: dict, field: str) -> Optional[str]:
    drop = entry.get("drop")
    name = drop.get(field) if isinstance(drop, dict) else None
    return name if isinstance(name, str) else None


def _entry_surfaces(entry: dict) -> Optional[List[str]]:
    listed = entry.get("surfaces")
    return [s for s in listed if isinstance(s, str)] if isinstance(listed, list) else None


def _violation(ctx: Context, rule_id: str, surface: str, key: str, disc: str, message: str) -> Violation:
    return Violation(rule_id, surface, key, disc, message, surfaces(ctx).sites.get(surface, ""))


def _tree_c(ctx: Context, entry: dict) -> Optional[CFunction]:
    """The header function carrying the entry's current or former C name."""
    by_name = _cached(ctx, "functions", lambda: {fn.name: fn for fn in ctx.functions})
    for name in _names(entry, "c"):
        if name in by_name:
            return by_name[name]
    return None


def _functions_by_key(ctx: Context) -> Dict[str, CFunction]:
    return _cached(ctx, "functions-by-key", lambda: {ctx.op_key(fn.name): fn for fn in ctx.functions})


def _in_tree(found: Dict[str, object], entry: dict, field: str) -> Optional[str]:
    """The tree name under which `entry` appears in `found`, excluding a dropped name."""
    dropped = _dropped(entry, field)
    for name in _names(entry, field):
        if name in found and name != dropped:
            return name
    return None


def _owned(rule_id: str):
    """Registers a rule that also reports the dropped binding names its rule-id closes."""

    def wrap(fn):
        def run(ctx: Context) -> Iterator[Violation]:
            if not _active(ctx):
                return
            yield from fn(ctx)
            yield from _drop_violations(ctx, rule_id)

        run.__doc__ = fn.__doc__
        run.__name__ = fn.__name__
        return rule(rule_id)(run)

    return wrap


def _drop_violations(ctx: Context, rule_id: str) -> Iterator[Violation]:
    for key, entry in ctx.ledger.entries.items():
        drop = entry.get("drop") if isinstance(entry, dict) else None
        if not isinstance(drop, dict) or drop.get("closes", DEFAULT_DROP_RULE) != rule_id:
            continue
        for surface in ("c", *BINDING_SURFACES):
            name = drop.get(_field(surface))
            if isinstance(name, str) and name in _symbols(ctx, surface):
                yield _violation(ctx, rule_id, surface, key, name, f"`{name}` is dropped from the {surface} surface")


def _words(name: str) -> List[str]:
    """Lower-case words of a snake_case, camelCase or PascalCase name."""
    out: List[str] = []
    for part in re.split(r"[_.]", name):
        out.extend(w.lower() for w in re.findall(r"[A-Z]+(?=[A-Z][a-z]|\d|$)|[A-Z]?[a-z]+\d*|[A-Z]+\d*|\d+", part))
    return out


def _camel(snake: str) -> str:
    words = [w for w in snake.split("_") if w]
    return words[0] + "".join(ACRONYMS.get(w, w[:1].upper() + w[1:]) for w in words[1:]) if words else ""


def _pascal(snake: str) -> str:
    return "".join(ACRONYMS.get(w, w[:1].upper() + w[1:]) for w in snake.split("_") if w)


def _target_verb(ctx: Context, fn: CFunction) -> Optional[str]:
    return c_verb(ctx.target_c(fn.name))


def _suffixed(name: str) -> bool:
    return name.endswith("_at") or name.endswith("_count")


def _positional_params(fn: CFunction) -> List[CParam]:
    return [
        p
        for p in fn.params
        if p.pointer_depth == 0
        and (p.role_name in ("index", "idx") or (p.role_name.endswith("_index") and p.role_name != "sheet_index"))
    ]


# ---------------------------------------------------------------------------
# R0: every tree symbol is an entry's current name.
# ---------------------------------------------------------------------------


@_owned("R0.coverage")
def _rule_coverage(ctx: Context) -> Iterable[Violation]:
    """Every exported symbol on every surface is the current name of a ledger entry that lists the surface.

    A dropped binding name is reported by the rule its `drop.closes` names; an entry with `replaced_by` is
    exempt until the type that replaces it lands.
    """
    for surface in ("c", *BINDING_SURFACES):
        field = _field(surface)
        for symbol in sorted(_symbols(ctx, surface)):
            key, entry = _lookup(ctx, surface, symbol)
            if key is None:
                yield _violation(ctx, "R0.coverage", surface, symbol, symbol, f"`{symbol}` has no ledger entry")
                continue
            if "replaced_by" in entry or _dropped(entry, field) == symbol:
                continue
            current = entry.get(field)
            if current != symbol:
                target = f"`{current}`" if isinstance(current, str) else "no current name"
                yield _violation(
                    ctx, "R0.coverage", surface, key, symbol, f"`{symbol}` is a former name; the ledger has {target}"
                )
                continue
            listed = _entry_surfaces(entry)
            if listed is not None and surface not in listed:
                yield _violation(
                    ctx, "R0.coverage", surface, key, symbol, f"`{symbol}` is on {surface}, which `{key}` does not list"
                )


# ---------------------------------------------------------------------------
# R1.scope: the ledger scope matches the argument shape.
# ---------------------------------------------------------------------------


@_owned("R1.scope")
def _rule_scope(ctx: Context) -> Iterable[Violation]:
    """A `sheet`-scope function takes `(fm_workbook_t*, size_t sheet_index, ...)`; a `workbook`-scope function
    takes the workbook handle and no sheet index unless its noun is `sheet`; a result-handle scope takes
    `fm_<scope>_t*` first.
    """
    for fn in ctx.functions:
        key = ctx.op_key(fn.name)
        entry = ctx.ledger.entry(ctx.ledger.key_for("c", fn.name))
        scope = entry.get("scope")
        if scope not in SCOPES:
            continue
        handle = fn.params[0] if fn.params else None
        want = f"fm_{'workbook' if scope in ('workbook', 'sheet') else scope}_t"
        if handle is None or ctx.target_c(handle.base_type) != want:
            yield ctx.c_violation(
                "R1.scope", fn.name, "handle", f"{fn.name}: `{scope}` scope takes `{want}*` first", fn.line
            )
        if scope == "sheet":
            second = fn.params[1] if len(fn.params) > 1 else None
            if second is None or (second.ctype, second.name) != ("size_t", "sheet_index"):
                yield ctx.c_violation(
                    "R1.scope",
                    fn.name,
                    "sheet_index",
                    f"{fn.name}: `sheet` scope takes `size_t sheet_index` second",
                    fn.line,
                )
        elif scope == "workbook" and entry.get("noun") != "sheet":
            if any(p.role_name in _SHEET_FIELDS and p.pointer_depth == 0 for p in fn.params[1:]):
                yield ctx.c_violation(
                    "R1.scope",
                    fn.name,
                    "scope",
                    f"{fn.name}: takes a sheet index, so `{key}` belongs to `sheet` scope",
                    fn.line,
                )


# ---------------------------------------------------------------------------
# R3: collection keys and suffixes.
# ---------------------------------------------------------------------------


def _collections(ctx: Context) -> Iterator[Tuple[str, str, List[str]]]:
    for name, coll in sorted(ctx.ledger.collections.items()):
        if isinstance(coll, dict):
            ops = [op for op in coll.get("ops") or [] if isinstance(op, str)]
            yield name, str(coll.get("key", "")), ops


@_owned("R3.suffix")
def _rule_suffix(ctx: Context) -> Iterable[Violation]:
    """`_at` and `_count` belong to collections: an index-keyed collection addresses by `<thing>_index` with no
    `_at`; an `_at` elsewhere is an enumeration position taking `size_t index`; every `_at` / `_count` function
    is an op of a ledger collection, or enumerates the elements of a result handle.
    """
    by_key = _functions_by_key(ctx)
    member_of: Dict[str, str] = {}
    for name, key_kind, ops in _collections(ctx):
        for op in ops:
            member_of.setdefault(op, name)
            fn = by_key.get(op)
            if fn is None or not fn.name.endswith("_at"):
                continue
            if key_kind == "index":
                yield ctx.c_violation(
                    "R3.suffix",
                    fn.name,
                    "at",
                    f"{fn.name}: `{name}` is index-keyed; address it by `<thing>_index` without `_at`",
                    fn.line,
                )
            elif not any(p.role_name in ("index", "idx") and p.pointer_depth == 0 for p in fn.params):
                yield ctx.c_violation(
                    "R3.suffix",
                    fn.name,
                    "at",
                    f"{fn.name}: `_at` is an enumeration position taking `size_t index`",
                    fn.line,
                )
    for key, fn in sorted(by_key.items()):
        handle_element = ctx.ledger.entry(key).get("scope") not in (None, *_COLLECTION_SCOPES)
        if _suffixed(fn.name) and c_verb(fn.name) != "set" and key not in member_of and not handle_element:
            yield ctx.c_violation(
                "R3.suffix",
                fn.name,
                "collection",
                f"{fn.name}: `_at` / `_count` outside every ledger collection",
                fn.line,
            )


def _id_key(fns: Sequence[CFunction]) -> Optional[str]:
    for fn in fns:
        for param in fn.params:
            if param.pointer_depth == 0 and param.name in ("cache_id", "num_fmt_id", "object_id"):
                return param.name
    return None


@_owned("R3.keys")
def _rule_keys(ctx: Context) -> Iterable[Violation]:
    """A natural-key collection has an unsuffixed getter, and its unsuffixed get / update / remove take the
    natural key rather than a position; an `add` that allocates the key (index, `object_id`, `cache_id`,
    `num_fmt_id`) returns it through an out parameter.
    """
    by_key = _functions_by_key(ctx)
    for name, key_kind, ops in _collections(ctx):
        fns = [by_key[op] for op in ops if op in by_key]
        if not fns:
            continue
        unsuffixed = [fn for fn in fns if not _suffixed(ctx.target_c(fn.name))]
        id_key = _id_key(fns) if key_kind == "natural" else None
        if key_kind == "natural":
            exempt = isinstance(ctx.ledger.collections.get(name), dict) and "no_getter" in ctx.ledger.collections[name]
            if not exempt and not any(_target_verb(ctx, fn) == "get" for fn in unsuffixed):
                yield _violation(
                    ctx, "R3.keys", "c", name, "getter", f"natural-key collection `{name}` has no getter by key"
                )
            for fn in unsuffixed:
                if _target_verb(ctx, fn) not in ("get", "update", "remove"):
                    continue
                for param in _positional_params(fn):
                    yield ctx.c_violation(
                        "R3.keys",
                        fn.name,
                        param.name,
                        f"{fn.name}: `{name}` is natural-keyed; address it by key, not `{param.name}`",
                        fn.line,
                    )
        for fn in unsuffixed:
            if _target_verb(ctx, fn) != "add":
                continue
            outs = {p.name for p in fn.params if p.pointer_depth == 1 and p.name.startswith("out_")}
            if key_kind == "index":
                if not any(o == "out_index" or o.endswith("_index") for o in outs):
                    yield ctx.c_violation(
                        "R3.keys",
                        fn.name,
                        "out-key",
                        f"{fn.name}: returns the allocated index through `out_index`",
                        fn.line,
                    )
            elif id_key and id_key not in {p.name for p in fn.params} and f"out_{id_key}" not in outs:
                yield ctx.c_violation(
                    "R3.keys",
                    fn.name,
                    "out-key",
                    f"{fn.name}: returns the allocated key through `out_{id_key}`",
                    fn.line,
                )


# ---------------------------------------------------------------------------
# R4: nouns.
# ---------------------------------------------------------------------------


def _synonym(name: str) -> Optional[str]:
    words = [w[:-1] if w.endswith("breaks") else w for w in _words(name)]
    for i in range(len(words)):
        for left in SYNONYMS:
            if tuple(words[i : i + len(left)]) == left:
                return "_".join(left)
    return None


@_owned("R4.unique")
def _rule_nouns(ctx: Context) -> Iterable[Violation]:
    """No name uses the left side of the synonym table (`horizontal_break`, `vertical_break`, `idx`); one
    `noun` is not shared by two scopes unless an entry of the pair carries `qualifies:`.
    """
    for surface in ("c", *BINDING_SURFACES):
        for symbol in sorted(_symbols(ctx, surface)):
            word = _synonym(symbol)
            if word is not None:
                replacement = SYNONYMS[tuple(word.split("_"))]
                yield _violation(
                    ctx,
                    "R4.unique",
                    surface,
                    _key(ctx, surface, symbol),
                    word,
                    f"`{symbol}` says `{word}`; use `{replacement}`",
                )
    groups: Dict[str, List[Tuple[str, dict]]] = {}
    for key, entry in ctx.ledger.entries.items():
        if isinstance(entry, dict) and isinstance(entry.get("noun"), str) and isinstance(entry.get("scope"), str):
            groups.setdefault(entry["noun"], []).append((key, entry))
    for noun, members in sorted(groups.items()):
        scopes = {entry["scope"] for _, entry in members}
        if len(scopes) < 2 or any("qualifies" in entry for _, entry in members):
            continue
        for key, entry in sorted(members, key=lambda m: m[0]):
            listed = _entry_surfaces(entry) or ["c"]
            surface = "c" if "c" in listed or "c" in entry else listed[0]
            yield _violation(
                ctx, "R4.unique", surface, key, "noun", f"noun `{noun}` is used by scopes {sorted(scopes)}; qualify one"
            )


# ---------------------------------------------------------------------------
# R5: binding names.
# ---------------------------------------------------------------------------


@_owned("R5.derive")
def _rule_derive(ctx: Context) -> Iterable[Violation]:
    """A binding name in the tree derives from the tree's C name of the same entry: functions by
    `derive_js` / `derive_py`, factories as `Workbook.<derived>`, records and enums by `derive_type`.
    Entries with `exception:` are exempt.
    """
    found = surfaces(ctx)
    typedefs = {t.name for t in ctx.typedefs}
    for key, entry in sorted(ctx.ledger.entries.items()):
        if not isinstance(entry, dict) or "exception" in entry:
            continue
        kind = entry.get("kind")
        if kind in ("op", "factory"):
            fn = _tree_c(ctx, entry)
            if fn is None:
                continue
            prefix = "Workbook." if kind == "factory" else ""
            wants = {"js": prefix + derive_js(fn.name), "py": prefix + derive_py(fn.name)}
            tree_c = fn.name
        elif kind in ("record", "enum"):
            tree_c = next((n for n in _names(entry, "c") if n in typedefs), None)
            if tree_c is None or "replaced_by" in entry:
                continue
            typed = derive_type(tree_c, entry.get("role"))
            wants = {"js": typed, "py": typed}
        else:
            continue
        for surface in BINDING_SURFACES:
            field = _field(surface)
            pool = (
                _callables(ctx, surface)
                if kind in ("op", "factory")
                else (found.py.classes if surface == "py" else found.js[surface].decls)
            )
            name = _in_tree(pool, entry, field)
            if name is not None and name != wants[field]:
                yield _violation(
                    ctx,
                    "R5.derive",
                    surface,
                    key,
                    NO_DISCRIMINATOR,
                    f"`{name}` does not derive from `{tree_c}` (`{wants[field]}`)",
                )


def _plural_name(ctx: Context, entry: dict) -> Optional[str]:
    sources = set(entry.get("from") or [])
    for _, coll in ctx.ledger.collections.items():
        if isinstance(coll, dict) and isinstance(coll.get("plural"), str) and sources & set(coll.get("ops") or []):
            return coll["plural"]
    return None


@_owned("R5.list")
def _rule_list(ctx: Context) -> Iterable[Violation]:
    """A `kind: list` entry is `list<Plural>` / `list_<plural>` of its collection and exists on each binding
    surface it lists; dropped count / at names are gone from the binding surfaces.
    """
    found = surfaces(ctx)
    for key, entry in sorted(ctx.ledger.entries.items()):
        if not isinstance(entry, dict) or entry.get("kind") != "list":
            continue
        plural = _plural_name(ctx, entry)
        listed = _entry_surfaces(entry) or list(BINDING_SURFACES)
        for surface in BINDING_SURFACES:
            if surface not in listed:
                continue
            field = _field(surface)
            name = entry.get(field)
            if not isinstance(name, str):
                continue
            if plural is not None:
                want = f"list_{plural}" if field == "py" else "list" + _pascal(plural)
                if name != want:
                    yield _violation(ctx, "R5.list", surface, key, "name", f"`{name}` should be `{want}`")
            pool = found.py.workbook if surface == "py" else found.js[surface].workbook
            if name not in pool:
                yield _violation(ctx, "R5.list", surface, key, "missing", f"`{name}` is not on the {surface} Workbook")


# ---------------------------------------------------------------------------
# R6: binding scalar types.
# ---------------------------------------------------------------------------


def _type_name(ctx: Context, spec: str, field: str) -> Optional[str]:
    """Binding type name a `returns` / `params` spec requires, or None for an unconstrained spec."""
    prefix, _, ref = spec.partition(":")
    if not ref:
        return None
    if prefix == "enum":
        for key in (ref, f"enum.{ref}"):
            name = ctx.ledger.entry(key).get(field)
            if isinstance(name, str):
                return name
        return _pascal(ref)
    if prefix in ("union", "literal"):
        return ref
    return None


def _member_for(pool: Dict[str, object], entry: dict, field: str):
    name = _in_tree(pool, entry, field)
    return (name, pool[name]) if name is not None else (None, None)


def _typed_entries(ctx: Context) -> Iterator[Tuple[str, dict]]:
    for key, entry in sorted(ctx.ledger.entries.items()):
        if isinstance(entry, dict) and (isinstance(entry.get("returns"), str) or isinstance(entry.get("params"), dict)):
            yield key, entry


def _mentions(text: str, name: str) -> bool:
    return re.search(rf"(?<![\w$]){re.escape(name)}(?![\w$])", text) is not None


def _scalar_type_problems(ctx: Context, entry: dict, field: str, returns: str, params: Dict[str, str]):
    """(discriminator, message) pairs for the ledger's `returns` / `params` on one binding member."""
    bool_return = ("BoolResult", "boolean") if field == "js" else ("bool",)
    spec = entry.get("returns")
    if isinstance(spec, str):
        if spec == "bool" and not any(returns == t or returns.startswith(t + "<") for t in bool_return):
            yield "return", f"returns `{returns}`; a boolean result is `{bool_return[0]}`"
        name = _type_name(ctx, spec, field)
        if name and not _mentions(returns, name):
            yield "return", f"returns `{returns}`; the ledger types it `{name}`"
    for param, spec in sorted((entry.get("params") or {}).items()):
        if not isinstance(spec, str):
            continue
        actual = params.get(_camel(param) if field == "js" else param)
        if actual is None:
            continue
        if spec == "bool" and actual != ("boolean" if field == "js" else "bool"):
            yield param, f"parameter `{param}: {actual}` is a boolean"
        name = _type_name(ctx, spec, field)
        if name and not _mentions(actual, name):
            yield param, f"parameter `{param}: {actual}`; the ledger types it `{name}`"


@_owned("R6.js")
def _rule_js_types(ctx: Context) -> Iterable[Violation]:
    """Ledger `returns` / `params` on the JS surfaces: a boolean is `boolean` (`BoolResult` as a result), an
    `enum:<key>` names the enum's type, a `union:<Name>` names the string-literal union.
    """
    for key, entry in _typed_entries(ctx):
        for surface in JS_SURFACES:
            _, member = _member_for(_callables(ctx, surface), entry, "js")
            if member is None:
                continue
            for disc, message in _scalar_type_problems(ctx, entry, "js", member.type, dict(member.params)):
                yield _violation(ctx, "R6.js", surface, key, disc, f"`{member.name}`: {message}")


@_owned("R6.py")
def _rule_py_types(ctx: Context) -> Iterable[Violation]:
    """Ledger `returns` / `params` on the Python surface: a boolean is `bool`, an `enum:<key>` names the
    `IntEnum`, a `union:<Name>` names the `Literal` alias.
    """
    for key, entry in _typed_entries(ctx):
        name, fn = _member_for(_callables(ctx, "py"), entry, "py")
        if fn is None:
            continue
        for disc, message in _scalar_type_problems(ctx, entry, "py", fn.returns, dict(fn.params)):
            yield _violation(ctx, "R6.py", "py", key, disc, f"`{name}`: {message}")


# ---------------------------------------------------------------------------
# R7: record roles.
# ---------------------------------------------------------------------------


def _records(ctx: Context) -> Iterator[Tuple[str, dict, str]]:
    for key, entry in sorted(ctx.ledger.entries.items()):
        if isinstance(entry, dict) and entry.get("kind") == "record" and "replaced_by" not in entry:
            role = entry.get("role")
            if role in ("result", "input", "patch"):
                yield key, entry, role


def _role_name_problem(name: str, role: str) -> Optional[str]:
    if role == "input" and not name.endswith("Input"):
        return f"`{name}` is an input record; name it `<Name>Input`"
    if role == "patch" and not name.endswith("Patch"):
        return f"`{name}` is a patch record; name it `<Name>Patch`"
    if role == "result" and (name.endswith("Input") or name.endswith("Patch")):
        return f"`{name}` is a result record; drop the `Input` / `Patch` suffix"
    return None


@_owned("R7.js")
def _rule_js_records(ctx: Context) -> Iterable[Violation]:
    """JS record shape by role: a result is an interface of `readonly` fields, an input an interface named
    `<Name>Input`, a patch `type <Name>Patch = Patch<...>`; `Patch<T>` is declared once as
    `{ [K in keyof T]?: T[K] | null }`.
    """
    found = surfaces(ctx)
    records = list(_records(ctx))
    for surface in JS_SURFACES:
        decls = found.js[surface].decls
        for key, entry, role in records:
            name = _in_tree(decls, entry, "js")
            if name is None:
                continue
            decl = decls[name]
            problem = _role_name_problem(name, role)
            if problem:
                yield _violation(ctx, "R7.js", surface, key, "name", problem)
            if role == "patch":
                if decl.kind != "type" or not decl.text.startswith("Patch<"):
                    yield _violation(ctx, "R7.js", surface, key, "patch", f"`{name}` is not declared as `Patch<...>`")
            elif decl.kind != "interface":
                yield _violation(
                    ctx, "R7.js", surface, key, "interface", f"`{name}` is a {decl.kind}, not an interface"
                )
            elif role == "result":
                mutable = [m.name for m in decl.members if not m.readonly]
                if mutable:
                    yield _violation(
                        ctx,
                        "R7.js",
                        surface,
                        key,
                        "readonly",
                        f"`{name}` is a result; `{mutable[0]}` is not `readonly`",
                    )
        if any(role == "patch" for _, _, role in records):
            alias = decls.get("Patch")
            if alias is None or alias.kind != "type" or _norm(alias.text) != _PATCH_ALIAS:
                yield _violation(
                    ctx, "R7.js", surface, "Patch", "alias", f"declare `export type Patch<T> = {_PATCH_ALIAS};` once"
                )


def _is_dataclass(cls: PyClass, frozen: bool) -> bool:
    for decorator in cls.decorators:
        if decorator in ("dataclass", "dataclasses.dataclass") and not frozen:
            return True
        if re.match(r"(?:dataclasses\.)?dataclass\(", decorator):
            if not frozen or re.search(r"\bfrozen\s*=\s*True\b", decorator):
                return True
    return False


@_owned("R7.py")
def _rule_py_records(ctx: Context) -> Iterable[Violation]:
    """Python record shape by role: a result is `@dataclass(frozen=True)` (no `NamedTuple`), an input a
    `@dataclass` named `<Name>Input`, a patch a `TypedDict(total=False)` named `<Name>Patch` with `Optional`
    fields.
    """
    classes = surfaces(ctx).py.classes
    for key, entry, role in _records(ctx):
        name = _in_tree(classes, entry, "py")
        if name is None:
            continue
        cls = classes[name]
        problem = _role_name_problem(name, role)
        if problem:
            yield _violation(ctx, "R7.py", "py", key, "name", problem)
        if role == "result" and (not _is_dataclass(cls, frozen=True) or "NamedTuple" in cls.bases):
            yield _violation(ctx, "R7.py", "py", key, "dataclass", f"`{name}` is a result; make it a frozen dataclass")
        elif role == "input" and not _is_dataclass(cls, frozen=False):
            yield _violation(ctx, "R7.py", "py", key, "dataclass", f"`{name}` is an input; make it a dataclass")
        elif role == "patch":
            if "TypedDict" not in cls.bases or cls.keywords.get("total") != "False":
                yield _violation(
                    ctx, "R7.py", "py", key, "typed-dict", f"`{name}` is a patch; make it `TypedDict, total=False`"
                )
            for field_name, annotation in sorted(cls.fields.items()):
                if annotation and not re.match(r"(?:typing\.)?Optional\[|.*\|\s*None$", annotation):
                    yield _violation(
                        ctx,
                        "R7.py",
                        "py",
                        key,
                        field_name,
                        f"`{name}.{field_name}: {annotation}` is not `Optional[...]`",
                    )


# ---------------------------------------------------------------------------
# R8: result shapes.
# ---------------------------------------------------------------------------


def _top_union(text: str) -> List[str]:
    return _split_top(text, "|")


def _optional_lookups(ctx: Context) -> Set[str]:
    """Ledger keys of the unsuffixed getters of natural-key collections."""
    by_key = _functions_by_key(ctx)
    out = set()
    for _, key_kind, ops in _collections(ctx):
        if key_kind != "natural":
            continue
        for op in ops:
            target = ctx.ledger.entry(op).get("c")
            if not isinstance(target, str) and op in by_key:
                target = by_key[op].name
            if isinstance(target, str) and c_verb(target) == "get" and not _suffixed(target):
                out.add(op)
    return out


def _allocated_key_result(fn: CFunction) -> Optional[str]:
    names = {p.name for p in fn.params if p.pointer_depth == 1}
    if "out_object_id" in names:
        return "ObjectIdResult"
    if any(n == "out_index" or (n.startswith("out_") and n.endswith("_index")) for n in names):
        return "IndexResult"
    return None


def _record_fields(found: Surfaces, record: dict, surface: str) -> Optional[Tuple[str, Set[str]]]:
    field = _field(surface)
    if surface == "py":
        name = _in_tree(found.py.classes, record, field)
        return (name, set(found.py.classes[name].fields)) if name else None
    decls = found.js[surface].decls
    name = _in_tree(decls, record, field)
    return (name, {m.name for m in decls[name].members}) if name else None


def _field_renames(ctx: Context, surface: str, rule_id: str) -> Iterator[Violation]:
    found = surfaces(ctx)
    field = _field(surface)
    for key, entry in sorted(ctx.ledger.entries.items()):
        if not isinstance(entry, dict) or entry.get("kind") != "field" or not isinstance(entry.get(field), str):
            continue
        record = ctx.ledger.entry(entry.get("record"))
        shape = _record_fields(found, record, surface) if record else None
        if shape is None:
            continue
        name, fields = shape
        new = entry[field]
        old = (entry.get("was") or {}).get(field) if isinstance(entry.get("was"), dict) else None
        if new not in fields or (isinstance(old, str) and old != new and old in fields):
            yield _violation(
                ctx, rule_id, surface, key, new, f"`{name}` should carry `{new}`" + (f", not `{old}`" if old else "")
            )


@_owned("R8.js")
def _rule_js_results(ctx: Context) -> Iterable[Violation]:
    """JS results: a scalar result is `StringResult` / `NumberResult` (`.value`), an allocated key
    `IndexResult` / `ObjectIdResult`, a natural-key lookup or a factory `OptionalResult<T>`; no method returns
    `null` bare; `kind: field` renames are applied.
    """
    found = surfaces(ctx)
    lookups = _optional_lookups(ctx)
    for surface in JS_SURFACES:
        js = found.js[surface]
        pool = _callables(ctx, surface)
        for envelope in _VALUE_ENVELOPES:
            decl = js.decls.get(envelope)
            if decl is not None and "value" not in {m.name for m in decl.members}:
                yield _violation(
                    ctx, "R8.js", surface, _key(ctx, surface, envelope), "value", f"`{envelope}` has no `value`"
                )
        for name, member in sorted(pool.items()):
            if "null" in _top_union(member.type):
                yield _violation(
                    ctx,
                    "R8.js",
                    surface,
                    _key(ctx, surface, name),
                    "null",
                    f"`{name}` returns `{member.type}`; wrap it",
                )
        for key, entry in sorted(ctx.ledger.entries.items()):
            if not isinstance(entry, dict):
                continue
            name, member = _member_for(pool, entry, "js")
            if member is None:
                continue
            want = _SCALAR_RESULTS.get(entry.get("returns")) if isinstance(entry.get("returns"), str) else None
            fn = _tree_c(ctx, entry)
            if fn is not None and c_verb(ctx.target_c(fn.name)) == "add":
                want = _allocated_key_result(fn) or want
            if key in lookups or entry.get("kind") == "factory":
                want = "OptionalResult"
            if want and not (member.type == want or member.type.startswith(want + "<")):
                yield _violation(
                    ctx, "R8.js", surface, key, "result", f"`{name}` returns `{member.type}`, not `{want}`"
                )
        yield from _field_renames(ctx, surface, "R8.js")


def _is_tuple(annotation: str, classes: Dict[str, PyClass]) -> bool:
    if re.match(r"(?:typing\.)?(?:Tuple|tuple)\b", annotation):
        return True
    base = re.sub(r"^(?:typing\.)?Optional\[(.*)\]$", r"\1", annotation)
    cls = classes.get(base)
    return cls is not None and "NamedTuple" in cls.bases


def _is_optional(annotation: str) -> bool:
    return bool(re.match(r"(?:typing\.)?Optional\[", annotation)) or "None" in _top_union(annotation)


@_owned("R8.py")
def _rule_py_results(ctx: Context) -> Iterable[Violation]:
    """Python results: no method returns a bare tuple or a `NamedTuple`; a scalar result is the scalar; a
    natural-key lookup returns `Optional[T]`; `kind: field` renames are applied.
    """
    py = surfaces(ctx).py
    pool = _callables(ctx, "py")
    lookups = _optional_lookups(ctx)
    for name, fn in sorted(pool.items()):
        if fn.returns and _is_tuple(fn.returns, py.classes):
            yield _violation(ctx, "R8.py", "py", _key(ctx, "py", name), "tuple", f"`{name}` returns `{fn.returns}`")
    for key, entry in sorted(ctx.ledger.entries.items()):
        if not isinstance(entry, dict):
            continue
        name, fn = _member_for(pool, entry, "py")
        if fn is None or not fn.returns:
            continue
        spec = entry.get("returns")
        allowed = _PY_SCALARS.get(spec) if isinstance(spec, str) else None
        if allowed and fn.returns not in allowed:
            yield _violation(ctx, "R8.py", "py", key, "result", f"`{name}` returns `{fn.returns}`, not `{allowed[0]}`")
        if key in lookups and not _is_optional(fn.returns):
            yield _violation(
                ctx, "R8.py", "py", key, "optional", f"`{name}` returns `{fn.returns}`, not `Optional[...]`"
            )
    yield from _field_renames(ctx, "py", "R8.py")


# ---------------------------------------------------------------------------
# R11: multiple ranges and ranges on another sheet.
# ---------------------------------------------------------------------------


def _is_text(ctype: str) -> bool:
    return ctype.replace("const ", "").strip() in ("char*", "char**")


def _multi_problems(members: Sequence[CParam], owner: str) -> Iterator[Tuple[str, str]]:
    for i, member in enumerate(members):
        role = member.role_name
        if _is_text(member.ctype) and (role in _A1_MULTI_NAMES or (role == "ref" and "table" in owner)):
            yield (
                member.name,
                f"`{member.ctype} {member.name}` is A1 text; take `const {CELL_RECT}* ranges, size_t range_count`",
            )
            continue
        follower = members[i + 1] if i + 1 < len(members) else None
        if (
            member.pointer_depth == 1
            and follower is not None
            and follower.role_name.endswith("count")
            and (role in _A1_MULTI_NAMES or member.base_type == CELL_RECT)
        ):
            if (member.base_type, member.name, follower.ctype, follower.name) != (
                CELL_RECT,
                "ranges",
                "size_t",
                "range_count",
            ):
                yield (
                    member.name,
                    (
                        f"`{member.ctype} {member.name}, {follower.ctype} {follower.name}` should be "
                        f"`const {CELL_RECT}* ranges, size_t range_count`"
                    ),
                )


@_owned("R11.multi")
def _rule_multi_ranges(ctx: Context) -> Iterable[Violation]:
    """Several ranges cross as `const fm_cell_rect_t* ranges, size_t range_count` in C and `CellRect[]` in
    the bindings, never as A1 text (`sqref`, `ranges_a1`, a table's `ref`).
    """
    for typedef in ctx.typedefs:
        for disc, message in _multi_problems(typedef.fields, typedef.name):
            yield ctx.c_violation("R11.multi", typedef.name, disc, f"{typedef.name}: {message}", typedef.line)
    for fn in ctx.functions:
        target = ctx.target_c(fn.name)
        if c_verb(target) in _TEXT_VERBS:
            continue
        for disc, message in _multi_problems(fn.params, target):
            yield ctx.c_violation("R11.multi", fn.name, disc, f"{fn.name}: {message}", fn.line)
    found = surfaces(ctx)
    for surface in JS_SURFACES:
        js = found.js[surface]
        for decl in js.decls.values():
            for member in decl.members:
                if not member.is_method and _js_multi_problem(member.name, member.type, decl.name):
                    yield _violation(
                        ctx,
                        "R11.multi",
                        surface,
                        _key(ctx, surface, decl.name),
                        member.name,
                        f"`{decl.name}.{member.name}: {member.type}` should be `ranges: CellRect[]`",
                    )
        for name, member in sorted(_callables(ctx, surface).items()):
            for param, ptype in member.params:
                if _js_multi_problem(param, ptype, name):
                    yield _violation(
                        ctx,
                        "R11.multi",
                        surface,
                        _key(ctx, surface, name),
                        param,
                        f"`{name}({param}: {ptype})` should take `CellRect[]`",
                    )
    py = found.py
    for cls_name, cls in sorted(py.classes.items()):
        for field_name, annotation in sorted(cls.fields.items()):
            if _py_multi_problem(field_name, annotation, cls_name):
                yield _violation(
                    ctx,
                    "R11.multi",
                    "py",
                    _key(ctx, "py", cls_name),
                    field_name,
                    f"`{cls_name}.{field_name}: {annotation}` should be `ranges: List[CellRect]`",
                )
    for name, fn in sorted(_callables(ctx, "py").items()):
        for param, annotation in fn.params:
            if _py_multi_problem(param, annotation, name):
                yield _violation(
                    ctx,
                    "R11.multi",
                    "py",
                    _key(ctx, "py", name),
                    param,
                    f"`{name}({param}: {annotation})` should take `List[CellRect]`",
                )


def _js_multi_problem(name: str, ptype: str, owner: str) -> bool:
    if name == "sqref" or name == "rangesA1":
        return True
    is_text = ptype in ("string", "string | null") or ptype.startswith("string")
    return is_text and (name == "ranges" or (name == "ref" and "table" in owner.lower()))


def _py_multi_problem(name: str, annotation: str, owner: str) -> bool:
    if name in ("sqref", "ranges_a1"):
        return True
    is_text = bool(re.match(r"(?:Optional\[)?str\b", annotation))
    return is_text and (name == "ranges" or (name == "ref" and "table" in owner.lower()))


def _edge_quad(names: Set[str], quads: Sequence[Sequence[str]]) -> bool:
    return any(all(part in names for part in quad) for quad in quads)


@_owned("R11.sheet")
def _rule_sheet_ranges(ctx: Context) -> Iterable[Violation]:
    """A range on another sheet is `fm_sheet_rect_t { size_t sheet_index; fm_cell_rect_t rect; }` in C and
    `SheetRect { sheet, rect }` in the bindings; no other record pairs a sheet index with a rectangle.
    """
    for typedef in ctx.typedefs:
        if typedef.kind != "struct":
            continue
        fields = {m.name: m for m in typedef.fields}
        if typedef.name == SHEET_RECT:
            shape = tuple((m.ctype, m.name) for m in typedef.fields)
            if shape != (("size_t", "sheet_index"), (CELL_RECT, "rect")):
                yield ctx.c_violation(
                    "R11.sheet",
                    typedef.name,
                    NO_DISCRIMINATOR,
                    f"{SHEET_RECT} must be exactly `size_t sheet_index; {CELL_RECT} rect;`",
                    typedef.line,
                )
            continue
        has_sheet = any(n in _SHEET_FIELDS and fields[n].pointer_depth == 0 for n in fields)
        has_rect = _edge_quad(set(fields), _EDGE_QUADS) or any(m.ctype == CELL_RECT for m in typedef.fields)
        if has_sheet and has_rect:
            yield ctx.c_violation(
                "R11.sheet",
                typedef.name,
                NO_DISCRIMINATOR,
                f"{typedef.name} pairs a sheet with a rectangle; use `{SHEET_RECT}`",
                typedef.line,
            )
    found = surfaces(ctx)
    for surface in JS_SURFACES:
        for decl in found.js[surface].decls.values():
            if decl.kind != "interface":
                continue
            names = {m.name for m in decl.members if not m.is_method}
            if not ({"sheet", "sheetIndex"} & names):
                continue
            has_rect = _edge_quad(names, _JS_EDGE_QUADS) or "rect" in names
            if decl.name == "SheetRect":
                if names != {"sheet", "rect"}:
                    yield _violation(
                        ctx,
                        "R11.sheet",
                        surface,
                        _key(ctx, surface, decl.name),
                        "shape",
                        "`SheetRect` is `{ sheet, rect }`",
                    )
            elif has_rect:
                yield _violation(
                    ctx,
                    "R11.sheet",
                    surface,
                    _key(ctx, surface, decl.name),
                    NO_DISCRIMINATOR,
                    f"`{decl.name}` pairs a sheet with a rectangle; use `SheetRect`",
                )
    for name, cls in sorted(found.py.classes.items()):
        names = set(cls.fields)
        if not ({"sheet", "sheet_index"} & names):
            continue
        has_rect = _edge_quad(names, _EDGE_QUADS) or "rect" in names
        if name == "SheetRect":
            if names != {"sheet", "rect"}:
                yield _violation(
                    ctx, "R11.sheet", "py", _key(ctx, "py", name), "shape", "`SheetRect` is `(sheet, rect)`"
                )
        elif has_rect:
            yield _violation(
                ctx,
                "R11.sheet",
                "py",
                _key(ctx, "py", name),
                NO_DISCRIMINATOR,
                f"`{name}` pairs a sheet with a rectangle; use `SheetRect`",
            )


# ---------------------------------------------------------------------------
# R12: JS entry points.
# ---------------------------------------------------------------------------


def _constant_tables(js: JsSurface) -> List[str]:
    return sorted(
        d.name for d in js.decls.values() if d.kind == "enum" or (d.kind == "const" and d.text.startswith("Readonly<{"))
    )


def _one_sided(ctx: Context, names: Dict[str, Set[str]], disc: str) -> Iterator[Violation]:
    """Names present on one JS surface and missing on the other, unless the ledger confines them."""
    for surface, other in (("wasm", "native"), ("native", "wasm")):
        for name in sorted(names[other] - names[surface]):
            key, entry = _lookup(ctx, other, name)
            listed = _entry_surfaces(entry)
            if listed is not None and surface not in listed:
                continue
            yield _violation(ctx, "R12", surface, key or name, disc, f"`{name}` is on {other} but not on {surface}")


@_owned("R12")
def _rule_entry_points(ctx: Context) -> Iterable[Violation]:
    """Both npm packages put every constant table (enums, `StatusCode`) on the module object as well as the
    named exports, and declare the same `Workbook` methods, factories and free functions, except where the
    ledger's `surfaces:` confines a name to one package.
    """
    js = surfaces(ctx).js
    for surface in JS_SURFACES:
        module = js[surface].module_object
        for table in _constant_tables(js[surface]):
            if table not in module:
                yield _violation(
                    ctx, "R12", surface, _key(ctx, surface, table), "module", f"`{table}` is not on the module object"
                )
    yield from _one_sided(ctx, {s: set(js[s].workbook) for s in JS_SURFACES}, "method")
    yield from _one_sided(ctx, {s: {f"Workbook.{n}" for n in js[s].statics} for s in JS_SURFACES}, "static")
    yield from _one_sided(ctx, {s: set(js[s].functions) - {"createFormulon"} for s in JS_SURFACES}, "function")
