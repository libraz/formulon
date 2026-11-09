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
    derive_record_type,
    derive_type,
    entry_handle,
    handle_classes,
    handle_collections,
    handle_name,
    rule,
)
from .surface_files import _SKIPPED, NODE_DTS, NPM_COMMON_MJS, PYTHON_PKG_DIR, REPO_ROOT, WASM_DTS
from .ts_decl import _extract_braced_block

JS_SURFACES = ("wasm", "native")
BINDING_SURFACES = ("wasm", "native", "py")
SHEET_RECT = "fm_sheet_rect_t"
WORKBOOK_HANDLE = "fm_workbook_t"
DEFAULT_DROP_RULE = "R5.list"
ENTRY_INTERFACE = "Formulon"
WORKBOOK_CLASS = "Workbook"
SHEET_CLASS = "Sheet"
# The interface each class's read-only `ooxml` property returns; its methods are binding ops like the class's own.
OOXML_CLASSES = {WORKBOOK_CLASS: "WorkbookOoxml", SHEET_CLASS: "SheetOoxml"}

# Hand-written facade members that no ledger op stands behind (or that the facade adds around one); the
# `Workbook` statics are the ledger factories and are not repeated here.
FACADE_EXTRA = {
    "js": {
        WORKBOOK_CLASS: frozenset(
            {"sheet", "listSheets", "dispose", "disposed", "[Symbol.dispose]", "getActiveSheet", "setActiveSheet"}
        ),
        SHEET_CLASS: frozenset({"workbook", "index", "name"}),
    },
    "py": {
        WORKBOOK_CLASS: frozenset({"sheet", "list_sheets", "close", "closed", "get_active_sheet", "set_active_sheet"}),
        SHEET_CLASS: frozenset({"workbook", "index", "name"}),
    },
}

# Synonyms whose left side must not appear in a name (R4); a pair is two adjacent words.
SYNONYMS = {("horizontal", "break"): "row_break", ("vertical", "break"): "column_break", ("idx",): "index"}

# Names a multi-range argument or field carries while it is still A1 text or a non-canonical list.
_A1_MULTI_NAMES = frozenset({"sqref", "ranges_a1", "ranges"})
_JS_A1_MULTI_NAMES = frozenset({"sqref", "rangesA1", "ranges"})
_TEXT_VERBS = frozenset({"parse", "format", "resolve", "convert", "localize", "canonicalize", "evaluate"})
_SHEET_FIELDS = frozenset({"sheet", "sheet_index"})
# Scopes whose `_at` / `_count` address a ledger collection; any other scope is a result handle.
_COLLECTION_SCOPES = ("workbook", "sheet", "process")
_EDGE_QUADS = (CELL_RECT_FIELDS, ("top", "left", "bottom", "right"), ("row", "col", "last_row", "last_col"))
_JS_EDGE_QUADS = (("firstRow", "firstCol", "lastRow", "lastCol"), ("top", "left", "bottom", "right"))
# Loose corners a record keyed by a sheet name (a reference) may spell instead of `rect`.
_NAMED_EDGE_QUADS = (("row0", "col0", "row1", "col1"),)

_SCALAR_JS = {"string": "string", "number": "number", "int": "number", "double": "number"}
_PY_SCALARS = {"string": ("str",), "number": ("int", "float"), "int": ("int",), "double": ("float",)}
# Result-envelope types that no JS declaration may carry: a failure is thrown, a success is the payload.
_RETIRED_RESULT_TYPES = (
    "Status",
    "OptionalResult",
    "ListResult",
    "StringResult",
    "NumberResult",
    "BoolResult",
    "IndexResult",
    "ObjectIdResult",
)
_PATCH_ALIAS = "{ [K in keyof T]?: K extends Clearable ? T[K] | null : T[K] }"


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
# A module-private interface: an alternative of an exported discriminated union, never published by name.
_TS_PRIVATE_INTERFACE_RE = re.compile(r"^(?:declare\s+)?interface\s+([A-Za-z_$][\w$]*)", re.M)
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
        self.raw = text
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
        self.private: Dict[str, TsDecl] = {}
        for match in _TS_PRIVATE_INTERFACE_RE.finditer(text):
            brace = text.find("{", match.end())
            body = _extract_braced_block(text, brace + 1) if brace >= 0 else ""
            self.private.setdefault(match.group(1), TsDecl(match.group(1), "interface", body, _parse_members(body)))

    def interface_members(self, name: str) -> Dict[str, TsMember]:
        """Members of interface `name`; an overloaded method is its first signature (the coordinate form)."""
        decl = self.decls.get(name)
        out: Dict[str, TsMember] = {}
        for member in decl.members if decl and decl.kind == "interface" else ():
            out.setdefault(member.name, member)
        return out

    def overloads(self, name: str) -> Tuple[TsMember, ...]:
        """Every signature of the `Workbook` or `Sheet` method `name` (or of their `ooxml` layers), in declaration
        order."""
        for interface in (SHEET_CLASS, WORKBOOK_CLASS, *OOXML_CLASSES.values()):
            decl = self.decls.get(interface)
            found = tuple(m for m in decl.members if m.name == name and m.is_method) if decl else ()
            if found:
                return found
        return ()

    @property
    def workbook(self) -> Dict[str, TsMember]:
        return {n: m for n, m in self.interface_members(WORKBOOK_CLASS).items() if m.is_method}

    @property
    def sheet(self) -> Dict[str, TsMember]:
        """Methods of the `Sheet` facade; they omit the leading sheet argument of the raw registration."""
        return {n: m for n, m in self.interface_members(SHEET_CLASS).items() if m.is_method}

    @property
    def ooxml(self) -> Dict[str, TsMember]:
        """Methods of the `ooxml` layers of `Workbook` and `Sheet`; the sheet-level ones omit the sheet argument."""
        out: Dict[str, TsMember] = {}
        for interface in OOXML_CLASSES.values():
            out.update({n: m for n, m in self.interface_members(interface).items() if m.is_method})
        return out

    def classes(self) -> Dict[str, Dict[str, TsMember]]:
        """The direct methods of `Workbook` and `Sheet`; the `ooxml` layers are not among them."""
        return {WORKBOOK_CLASS: self.workbook, SHEET_CLASS: self.sheet}

    @property
    def statics(self) -> Dict[str, TsMember]:
        return {n: m for n, m in self.interface_members("WorkbookCtor").items() if m.is_method}

    @property
    def module_object(self) -> Dict[str, TsMember]:
        """Members of the object `createFormulon` resolves to: the `Formulon` interface."""
        return self.interface_members(ENTRY_INTERFACE)

    @property
    def functions(self) -> Dict[str, TsMember]:
        """Free functions: methods of `Formulon`, and the top-level `export function`s."""
        module = {n: m for n, m in self.module_object.items() if m.is_method}
        exported = {d.name: d.members[0] for d in self.decls.values() if d.kind == "function" and d.members}
        return {**exported, **module}

    def callables(self) -> Dict[str, TsMember]:
        """Instance methods (the `ooxml` layers' included), free functions and `Workbook.<static>` factories by
        tree name."""
        out = dict(self.functions)
        out.update(self.ooxml)
        out.update(self.sheet)
        out.update(self.workbook)
        out.update({f"Workbook.{n}": m for n, m in self.statics.items()})
        return out

    def symbols(self) -> Dict[str, str]:
        """Every exported symbol of the file, by the name the ledger's `js` field uses.

        The hand-written facade members of `Sheet` (`FACADE_EXTRA`) have no ledger entry and are left out.
        """
        out = {name: "type" for name in self.decls}
        out.update({name: "function" for name in self.functions})
        out.update({name: "method" for name in self.ooxml})
        out.update({name: "method" for name in self.sheet if name not in FACADE_EXTRA["js"][SHEET_CLASS]})
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
    defaults: Tuple[Tuple[str, str], ...] = ()
    kwonly: Tuple[str, ...] = ()
    doc: str = ""


class PyClass(NamedTuple):
    name: str
    bases: Tuple[str, ...]
    decorators: Tuple[str, ...]
    keywords: Dict[str, str]
    fields: Dict[str, str]
    methods: Dict[str, PyFunc]
    constants: Dict[str, int] = {}
    doc: str = ""
    overloads: Dict[str, Tuple[PyFunc, ...]] = {}


_OVERLOAD_DECORATORS = frozenset({"overload", "typing.overload"})


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
    defaults = [(a.arg, ast.unparse(d)) for a, d in zip(reversed(positional), reversed(args.defaults))]
    defaults += [(a.arg, ast.unparse(d)) for a, d in zip(args.kwonlyargs, args.kw_defaults) if d is not None]
    kwonly = tuple(a.arg for a in args.kwonlyargs)
    doc = ast.get_docstring(node) or ""  # type: ignore[arg-type]
    return PyFunc(node.name, params, returns, decorators, tuple(defaults), kwonly, doc)  # type: ignore[attr-defined]


def _py_class(node: ast.ClassDef) -> PyClass:
    fields: Dict[str, str] = {}
    methods: Dict[str, PyFunc] = {}
    constants: Dict[str, int] = {}
    stubs: Dict[str, List[PyFunc]] = {}
    for item in node.body:
        if isinstance(item, ast.AnnAssign) and isinstance(item.target, ast.Name):
            fields[item.target.id] = ast.unparse(item.annotation)
        elif isinstance(item, ast.Assign) and len(item.targets) == 1 and isinstance(item.targets[0], ast.Name):
            value = item.value
            if isinstance(value, ast.UnaryOp) and isinstance(value.op, ast.USub):
                value = ast.Constant(-value.operand.value) if isinstance(value.operand, ast.Constant) else value
            if isinstance(value, ast.Constant) and isinstance(value.value, int) and not isinstance(value.value, bool):
                constants[item.targets[0].id] = value.value
        elif isinstance(item, (ast.FunctionDef, ast.AsyncFunctionDef)):
            func = _py_func(item, method=True)
            if _OVERLOAD_DECORATORS & set(func.decorators):
                stubs.setdefault(item.name, []).append(func)
            # A method with `@overload` stubs reads as its first stub (the coordinate form), else as its last `def`.
            methods[item.name] = stubs[item.name][0] if item.name in stubs else func
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
        constants,
        ast.get_docstring(node) or "",
        {name: tuple(found) for name, found in stubs.items()},
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
        # Module-level `Name = <type expression>` assignments: alias name -> right-hand side source.
        self.aliases: Dict[str, str] = {}
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
                elif isinstance(node, ast.Assign) and len(node.targets) == 1 and isinstance(node.targets[0], ast.Name):
                    self.aliases.setdefault(node.targets[0].id, ast.unparse(node.value))

    @property
    def workbook(self) -> Dict[str, PyFunc]:
        """Public `Workbook` methods; class and static methods as `Workbook.<name>`."""
        cls = self.classes.get(WORKBOOK_CLASS)
        out: Dict[str, PyFunc] = {}
        for name, fn in cls.methods.items() if cls else ():
            if name.startswith("_"):
                continue
            static = any(d in ("classmethod", "staticmethod") for d in fn.decorators)
            out[f"Workbook.{name}" if static else name] = fn
        return out

    @property
    def sheet(self) -> Dict[str, PyFunc]:
        """Public `Sheet` methods other than properties; they omit the leading sheet argument."""
        cls = self.classes.get(SHEET_CLASS)
        return {
            name: fn
            for name, fn in (cls.methods.items() if cls else ())
            if not name.startswith("_")
            and not any(d.endswith("property") or d.endswith(".setter") for d in fn.decorators)
        }

    @property
    def ooxml(self) -> Dict[str, PyFunc]:
        """Public methods of the `ooxml` layers of `Workbook` and `Sheet` (`WorkbookOoxml`, `SheetOoxml`)."""
        out: Dict[str, PyFunc] = {}
        for class_name in OOXML_CLASSES.values():
            cls = self.classes.get(class_name)
            out.update({name: fn for name, fn in (cls.methods.items() if cls else ()) if not name.startswith("_")})
        return out

    def class_methods(self) -> Dict[str, Dict[str, PyFunc]]:
        """The direct methods of `Workbook` and `Sheet`; the `ooxml` layers are not among them."""
        return {WORKBOOK_CLASS: self.workbook, SHEET_CLASS: self.sheet}

    def overloads(self, name: str) -> Tuple[PyFunc, ...]:
        """The `@overload` stubs of the `Sheet` or `Workbook` method `name` (or of their `ooxml` layers), in
        definition order."""
        for class_name in (SHEET_CLASS, WORKBOOK_CLASS, *OOXML_CLASSES.values()):
            cls = self.classes.get(class_name)
            if cls is not None and name in cls.overloads:
                return cls.overloads[name]
        return ()

    def callables(self) -> Dict[str, PyFunc]:
        """Exported free functions and instance methods, the `ooxml` layers' included, by tree name."""
        out = {name: fn for name, fn in self.functions.items() if name in self.exports}
        out.update(self.ooxml)
        out.update(self.sheet)
        out.update(self.workbook)
        return out

    def symbols(self) -> Dict[str, str]:
        out = {name: "export" for name in self.exports}
        out.update({name: "method" for name in self.ooxml})
        out.update({name: "method" for name in self.sheet if name not in FACADE_EXTRA["py"][SHEET_CLASS]})
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


def handle_members(ctx: Context, surface: str, cls: str) -> Dict[str, object]:
    """Members of the handle class `cls` on `surface`: the interface's methods and fields, or the Python class's
    methods and properties."""
    found = surfaces(ctx)
    if surface == "py":
        py_class = found.py.classes.get(cls)
        return dict(py_class.methods) if py_class else {}
    decl = found.js[surface].decls.get(cls)
    out: Dict[str, object] = {}
    for member in decl.members if decl else ():
        out.setdefault(member.name, member)
    return out


def handle_in_tree(ctx: Context, surface: str, entry: dict) -> Optional[str]:
    """`Class.method` under which a handle op is declared on `surface`, or None (also for a non-handle op)."""
    if entry_handle(entry) is None:
        return None
    name = handle_name(entry, _field(surface), ctx.ledger.collections)
    for cls in handle_classes(entry):
        if name and name in handle_members(ctx, surface, cls):
            return f"{cls}.{name}"
    return None


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
    """Dropped names still in the tree; a dropped binding field is looked up among its record's fields."""
    for key, entry in ctx.ledger.entries.items():
        drop = entry.get("drop") if isinstance(entry, dict) else None
        if not isinstance(drop, dict) or drop.get("closes", DEFAULT_DROP_RULE) != rule_id:
            continue
        record = ctx.ledger.entry(entry.get("record")) if entry.get("kind") == "field" else None
        for surface in ("c", *BINDING_SURFACES):
            name = drop.get(_field(surface))
            if not isinstance(name, str):
                continue
            if record is not None and surface != "c":
                shape = _record_fields(surfaces(ctx), record, surface) if record else None
                present = shape is not None and name in shape[1]
            else:
                present = name in _symbols(ctx, surface)
            if present:
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


# Symbol-keyed members that alias a named method and are covered by that method's entry.
_SYMBOL_ALIASES = {"[Symbol.dispose]": "dispose"}


def _sheet_handle_names(ctx: Context, field: str) -> Set[str]:
    """Names the ledger publishes on `Sheet` through `handle: {class: Sheet}` (`rename`, `move`)."""
    out: Set[str] = set()
    for entry in ctx.ledger.entries.values():
        if isinstance(entry, dict) and SHEET_CLASS in handle_classes(entry):
            name = handle_name(entry, field, ctx.ledger.collections)
            if name:
                out.add(name)
    return out


@_owned("R0.coverage")
def _rule_coverage(ctx: Context) -> Iterable[Violation]:
    """Every exported symbol on every surface is the current name of a ledger entry that lists the surface.

    A dropped binding name is reported by the rule its `drop.closes` names; an entry with `replaced_by` is
    exempt until the type that replaces it lands.
    """
    for surface in ("c", *BINDING_SURFACES):
        field = _field(surface)
        for symbol in sorted(_symbols(ctx, surface)):
            if symbol in _SYMBOL_ALIASES and surface != "c":
                key, entry = _lookup(ctx, surface, _SYMBOL_ALIASES[symbol])
                if key is not None and entry.get(field) == _SYMBOL_ALIASES[symbol]:
                    continue
            key, entry = _lookup(ctx, surface, symbol)
            if key is None and surface != "c" and symbol in _sheet_handle_names(ctx, field):
                continue
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


# Scopes whose handle is not an `fm_<scope>_t`: a buffer is the bytes pointer a save returns.
_SCOPE_HANDLE_TYPES = {"buffer": "uint8_t"}


@_owned("R1.scope")
def _rule_scope(ctx: Context) -> Iterable[Violation]:
    """A `sheet`-scope function takes `(fm_workbook_t*, size_t sheet_index, ...)`; a `workbook`-scope function
    takes the workbook handle and no sheet index unless its noun is `sheet` or the entry carries `sheet_arg:`;
    a result-handle scope takes `fm_<scope>_t*` (or its `_SCOPE_HANDLE_TYPES` type) first. A factory creates
    the handle, so it takes none.
    """
    for fn in ctx.functions:
        key = ctx.op_key(fn.name)
        entry = ctx.ledger.entry(ctx.ledger.key_for("c", fn.name))
        scope = entry.get("scope")
        if scope not in SCOPES or entry.get("kind") == "factory":
            continue
        handle = fn.params[0] if fn.params else None
        want = _SCOPE_HANDLE_TYPES.get(scope) or f"fm_{'workbook' if scope in ('workbook', 'sheet') else scope}_t"
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
        elif scope == "workbook" and entry.get("noun") != "sheet" and "sheet_arg" not in entry:
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


def _is_position(role_name: str) -> bool:
    return role_name in ("index", "idx") or role_name.endswith("_index")


@_owned("R3.suffix")
def _rule_suffix(ctx: Context) -> Iterable[Violation]:
    """`_at` and `_count` belong to collections: an `_at` is a position whose last by-value parameter is an
    `index` / `<thing>_index`;
    every `_at` / `_count` function is an op of a ledger collection, or enumerates the elements of a result handle.
    """
    by_key = _functions_by_key(ctx)
    member_of: Dict[str, str] = {}
    for name, _key_kind, ops in _collections(ctx):
        for op in ops:
            member_of.setdefault(op, name)
            fn = by_key.get(op)
            if fn is None or not fn.name.endswith("_at"):
                continue
            values = [p for p in fn.params if p.pointer_depth == 0]
            if not values or not _is_position(values[-1].role_name):
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
    `derive_js` / `derive_py`, factories as `Workbook.<derived>`, enums by `derive_type`, records by
    `derive_type` with or without a meaningless suffix (R21.name judges the suffix). Entries with
    `exception:` or `review: true` are exempt.
    """
    found = surfaces(ctx)
    typedefs = {t.name for t in ctx.typedefs}
    for key, entry in sorted(ctx.ledger.entries.items()):
        if not isinstance(entry, dict) or "exception" in entry or entry.get("review") is True:
            continue
        kind = entry.get("kind")
        if kind in ("op", "factory"):
            fn = _tree_c(ctx, entry)
            if fn is None:
                continue
            prefix = "Workbook." if kind == "factory" else ""
            wants = {"js": {prefix + derive_js(fn.name)}, "py": {prefix + derive_py(fn.name)}}
            tree_c = fn.name
        elif kind in ("record", "enum"):
            tree_c = next((n for n in _names(entry, "c") if n in typedefs), None)
            if tree_c is None or "replaced_by" in entry:
                continue
            typed = {derive_type(tree_c, entry.get("role"))}
            if kind == "record":
                typed.add(derive_record_type(tree_c, entry.get("role"), ctx.ledger.handle_collections()))
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
            if name is not None and name not in wants[field]:
                yield _violation(
                    ctx,
                    "R5.derive",
                    surface,
                    key,
                    NO_DISCRIMINATOR,
                    f"`{name}` does not derive from `{tree_c}` (`{' or '.join(sorted(wants[field]))}`)",
                )


def _plural_name(ctx: Context, entry: dict) -> Optional[str]:
    sources = set(entry.get("from") or [])
    for _, coll in ctx.ledger.collections.items():
        if isinstance(coll, dict) and isinstance(coll.get("plural"), str) and sources & set(coll.get("ops") or []):
            return coll["plural"]
    return None


@_owned("R5.list")
def _rule_list(ctx: Context) -> Iterable[Violation]:
    """A `kind: list` entry is `list<Plural>` / `list_<plural>` of its collection and is callable on each binding
    surface it lists (a `Workbook` method or an exported function, or a method of its handle class); dropped
    count / at names are gone from the binding surfaces. An `internal` list is raw-only.
    """
    for key, entry in sorted(ctx.ledger.entries.items()):
        if not isinstance(entry, dict) or entry.get("kind") != "list" or entry.get("internal") is True:
            continue
        if entry_handle(entry) is not None:
            for surface in BINDING_SURFACES:
                if surface in (_entry_surfaces(entry) or list(BINDING_SURFACES)) and not handle_in_tree(
                    ctx, surface, entry
                ):
                    shown = handle_name(entry, _field(surface), ctx.ledger.collections)
                    yield _violation(
                        ctx, "R5.list", surface, key, "missing", f"`{shown}` is not on its handle class on {surface}"
                    )
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
            if name not in _callables(ctx, surface):
                yield _violation(ctx, "R5.list", surface, key, "missing", f"`{name}` is not on the {surface} surface")


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
    if prefix == "union":
        record = ctx.ledger.entries.get(f"record.{'_'.join(_words(ref))}")
        if isinstance(record, dict):
            name = record.get(field)
            return name if isinstance(name, str) else None
        return ref
    if prefix == "literal":
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


def _without_unchanged(annotation: str) -> str:
    """`Union[T, Unchanged]` (a partial-update argument defaulting to `UNCHANGED`) as `T`."""
    union = re.fullmatch(r"(?:typing\.)?Union\[(.*)\]", annotation.strip())
    if union is None:
        return annotation
    parts = [a.strip() for a in _split_top(union.group(1), ",")]
    kept = [a for a in parts if a.rpartition(".")[2] != "Unchanged"]
    return kept[0] if len(kept) == 1 else annotation


_OPTIONAL_BOOL_RE = re.compile(r"(?:typing\.)?Optional\[bool\]|bool\s*\|\s*None|None\s*\|\s*bool")


def _scalar_type_problems(
    ctx: Context,
    entry: dict,
    field: str,
    returns: str,
    params: Dict[str, str],
    defaults: Dict[str, str],
    lookup: bool = False,
):
    """(discriminator, message) pairs for the ledger's `returns` / `params` on one binding member.

    On the Python surface a `bool` parameter may be `Optional[bool] = None` (absent means unchanged), and a
    boolean natural-key lookup returns `Optional[bool]`.
    """
    bool_return = ("boolean",) if field == "js" else ("bool",)
    spec = entry.get("returns")
    if isinstance(spec, str):
        if spec == "bool" and field == "py" and lookup:
            if not _OPTIONAL_BOOL_RE.fullmatch(returns):
                yield "return", f"returns `{returns}`; a boolean lookup result is `Optional[bool]`"
        elif spec == "bool" and field == "js" and lookup and returns == "boolean | null":
            pass
        elif spec == "bool" and not any(returns == t or returns.startswith(t + "<") for t in bool_return):
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
        if field == "py" and (defaults.get(param) or "").rpartition(".")[2] == "UNCHANGED":
            actual = _without_unchanged(actual)
        optional_default = (
            field == "py" and _OPTIONAL_BOOL_RE.fullmatch(actual) is not None and defaults.get(param) == "None"
        )
        if spec == "bool" and actual != ("boolean" if field == "js" else "bool") and not optional_default:
            yield param, f"parameter `{param}: {actual}` is a boolean"
        name = _type_name(ctx, spec, field)
        if name and not _mentions(actual, name):
            yield param, f"parameter `{param}: {actual}`; the ledger types it `{name}`"


@_owned("R6.js")
def _rule_js_types(ctx: Context) -> Iterable[Violation]:
    """Ledger `returns` / `params` on the JS surfaces: a boolean is `boolean` (`boolean | null` as a lookup), an
    `enum:<key>` names the enum's type, a `union:<Name>` names the string-literal union.
    """
    lookups = _optional_lookups(ctx)
    for key, entry in _typed_entries(ctx):
        for surface in JS_SURFACES:
            _, member = _member_for(_callables(ctx, surface), entry, "js")
            if member is None:
                continue
            problems = _scalar_type_problems(
                ctx, entry, "js", member.type, dict(member.params), {}, lookup=key in lookups
            )
            for disc, message in problems:
                yield _violation(ctx, "R6.js", surface, key, disc, f"`{member.name}`: {message}")


@_owned("R6.py")
def _rule_py_types(ctx: Context) -> Iterable[Violation]:
    """Ledger `returns` / `params` on the Python surface: a boolean is `bool`, an `enum:<key>` names the
    `IntEnum`, a `union:<Name>` names the `Literal` alias.
    """
    lookups = _optional_lookups(ctx)
    for key, entry in _typed_entries(ctx):
        name, fn = _member_for(_callables(ctx, "py"), entry, "py")
        if fn is None:
            continue
        problems = _scalar_type_problems(
            ctx, entry, "py", fn.returns, dict(fn.params), dict(fn.defaults), lookup=key in lookups
        )
        for disc, message in problems:
            yield _violation(ctx, "R6.py", "py", key, disc, f"`{name}`: {message}")


# ---------------------------------------------------------------------------
# R7: record roles.
# ---------------------------------------------------------------------------


def _records(ctx: Context) -> Iterator[Tuple[str, dict, str]]:
    """`(key, entry, role)` of the records whose role fixes their binding shape; an `exception:` keeps a record's
    name and form, as it does for R21.name."""
    for key, entry in sorted(ctx.ledger.entries.items()):
        if isinstance(entry, dict) and entry.get("kind") == "record" and "replaced_by" not in entry:
            if "exception" in entry:
                continue
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


def _is_interface_union(decls: Dict[str, TsDecl], decl: TsDecl, private: Optional[Dict[str, TsDecl]] = None) -> bool:
    """True for a type alias that is a union of two or more interfaces (a discriminated record): exported ones, or
    the module-private alternatives in `private`."""
    alternatives = _top_union(decl.text)
    pool = {**(private or {}), **decls}
    return len(alternatives) > 1 and all(a in pool and pool[a].kind == "interface" for a in alternatives)


@_owned("R7.js")
def _rule_js_records(ctx: Context) -> Iterable[Violation]:
    """JS record shape by role: a result is an interface of `readonly` fields (or a union of such interfaces,
    which R16.js judges), an input an interface named
    `<Name>Input`, a patch `type <Name>Patch = Patch<...>` or a dedicated interface whose every field is optional
    (a patch whose keys take no `null`, or only the nullable ones); `Patch<T>` is declared once as
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
                if not (decl.kind == "type" and decl.text.startswith("Patch<")) and not _optional_interface(decl):
                    yield _violation(
                        ctx,
                        "R7.js",
                        surface,
                        key,
                        "patch",
                        f"`{name}` is neither `Patch<...>` nor an interface of optional fields",
                    )
            elif decl.kind == "type" and _is_interface_union(decls, decl):
                continue
            elif (
                role == "input" and decl.kind == "type" and _is_interface_union(decls, decl, found.js[surface].private)
            ):
                # A discriminated input: every alternative is itself an input record.
                for alternative in _top_union(decl.text):
                    problem = _role_name_problem(alternative, role)
                    if problem:
                        yield _violation(ctx, "R7.js", surface, key, f"name:{alternative}", problem)
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
                    ctx,
                    "R7.js",
                    surface,
                    "Patch",
                    "alias",
                    f"declare `export type Patch<T, Clearable extends keyof T = never> = {_PATCH_ALIAS};` once",
                )


_TS_FIELD_RE = re.compile(r"(?:^|[;{\n])\s*(?:readonly\s+)?[A-Za-z_$][\w$]*(\??)\s*:")


def _optional_interface(decl: TsDecl) -> bool:
    """An interface with fields, every one of them optional (`name?: T`)."""
    if decl.kind != "interface" or any(m.is_method for m in decl.members):
        return False
    marks = _TS_FIELD_RE.findall(_extract_top_level(decl.text))
    return bool(marks) and all(mark == "?" for mark in marks)


def _extract_top_level(body: str) -> str:
    """`body` with every nested `{...}` / `<...>` / `(...)` / `[...]` removed, so only its own fields remain."""
    out, depth = [], 0
    for char in body:
        if char in "{<([":
            depth += 1
        elif char in "}>)]":
            depth -= 1
        elif depth == 0:
            out.append(char)
    return "".join(out)


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
    `@dataclass` named `<Name>Input`; an argument typed by an input also accepts the read record `<Name>` when the
    package declares it, so a getter's value feeds back unchanged. A patch is a `TypedDict(total=False)` named
    `<Name>Patch` that only types a nested field: a method takes the keys of its patch as keyword-only arguments
    defaulting to `UNCHANGED` (R18.all), never the patch itself.
    """
    found = surfaces(ctx).py
    classes = found.classes
    # A nested patch field arrives as a keyword-only argument; only a positional one passes a whole patch.
    callables = found.callables()
    every = [(name, p, t) for name, fn in sorted(callables.items()) for p, t in fn.params]
    params = [(name, p, t) for name, p, t in every if p not in callables[name].kwonly]
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
        read = name[: -len("Input")] if role == "input" and name.endswith("Input") else None
        if read in classes and _is_dataclass(classes[read], frozen=False):
            for method in sorted({m for m, _, a in every if _mentions(a, name) and not _mentions(a, read)}):
                yield _violation(
                    ctx, "R7.py", "py", key, f"record:{method}", f"`{method}` takes `{name}`; accept `{read}` too"
                )
        elif role == "patch":
            if "TypedDict" not in cls.bases or cls.keywords.get("total") != "False":
                yield _violation(
                    ctx, "R7.py", "py", key, "typed-dict", f"`{name}` is a patch; make it `TypedDict, total=False`"
                )
            for method in sorted({m for m, _, annotation in params if _mentions(annotation, name)}):
                yield _violation(
                    ctx,
                    "R7.py",
                    "py",
                    key,
                    method,
                    f"`{method}` takes `{name}`; take its keys as keyword-only arguments defaulting to `UNCHANGED`",
                )


# ---------------------------------------------------------------------------
# R8: result shapes.
# ---------------------------------------------------------------------------


def _top_union(text: str) -> List[str]:
    return _split_top(text, "|")


def _optional_lookups(ctx: Context) -> Set[str]:
    """Ledger keys of the unsuffixed getters of natural-key collections; a `total:` getter answers every key."""
    by_key = _functions_by_key(ctx)
    out = set()
    for _, key_kind, ops in _collections(ctx):
        if key_kind != "natural":
            continue
        for op in ops:
            if ctx.ledger.entry(op).get("total"):
                continue
            target = ctx.ledger.entry(op).get("c")
            if not isinstance(target, str) and op in by_key:
                target = by_key[op].name
            if isinstance(target, str) and c_verb(target) == "get" and not _suffixed(target):
                out.add(op)
    return out


def _allocates_key(fn: CFunction) -> bool:
    names = {p.name for p in fn.params if p.pointer_depth == 1}
    return "out_object_id" in names or any(
        n == "out_index" or (n.startswith("out_") and n.endswith("_index")) for n in names
    )


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
            for parent in _split_top_level(match.group(2).strip(), ","):
                if parent in seen:
                    continue
                if parent in decls:
                    out.update(_js_fields(decls, raw, parent, seen + (name,)) or {})
                else:
                    # A utility-type parent (`Omit<T, ...>`, `Pick<T, ...>`).
                    out.update(_patch_keys(decls, raw, parent) or {})
    out.update({m.name: m.type for m in decl.members if not m.is_method})
    return out


def _split_top_level(text: str, sep: str) -> List[str]:
    """`text` split at `sep` outside any bracket."""
    out, depth, start = [], 0, 0
    for i, ch in enumerate(text):
        if ch in "<{([":
            depth += 1
        elif ch in ">})]":
            depth -= 1
        elif ch == sep and depth == 0:
            out.append(text[start:i].strip())
            start = i + 1
    out.append(text[start:].strip())
    return out


def _first_type_arg(args: str) -> str:
    """The first top-level argument of a type argument list."""
    depth = 0
    for i, ch in enumerate(args):
        if ch in "<{([":
            depth += 1
        elif ch in ">})]":
            depth -= 1
        elif ch == "," and depth == 0:
            return args[:i]
    return args


def _patch_keys(decls: Dict[str, TsDecl], raw: str, text: str) -> Optional[Dict[str, str]]:
    """Key -> type of a patch type expression: an interface, `Patch<T[, C]>`, `Pick` / `Omit<T, 'a' | 'b'>`, `{ ... }`, or an `&` of these."""
    text = text.strip()
    parts = _split_top_level(text, "&")
    if len(parts) > 1:
        merged: Dict[str, str] = {}
        for part in parts:
            keys = _patch_keys(decls, raw, part)
            if keys is None:
                return None
            merged.update(keys)
        return merged
    omit = re.fullmatch(r"Omit<\s*([A-Za-z_$][\w$]*)\s*,(.*)>", text, re.S)
    if omit:
        fields = _patch_keys(decls, raw, omit.group(1))
        gone = {a or b for a, b in re.findall(r"'([^']*)'|\"([^\"]*)\"", omit.group(2))}
        return None if fields is None else {k: v for k, v in fields.items() if k not in gone}
    inner = re.fullmatch(r"(?:Patch|Partial)<(.*)>", text, re.S)
    if inner:
        # `Patch<T, 'a' | 'b'>`: the keys are T's; the second argument names the clearable ones.
        return _patch_keys(decls, raw, _first_type_arg(inner.group(1)))
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


def _js_record_keys(decls: Dict[str, TsDecl], raw: str, name: str) -> Optional[Dict[str, str]]:
    """Field name -> type of record type `name`: an interface (inherited fields included), a patch alias, or the
    fields of every alternative of a discriminated union; None when the checker cannot read it."""
    keys = _patch_keys(decls, raw, name)
    if keys is not None:
        return keys
    decl = decls.get(name)
    alternatives = _top_union(decl.text) if decl is not None and decl.kind == "type" else []
    if len(alternatives) < 2:
        return None
    out: Dict[str, str] = {}
    for alternative in alternatives:
        out.update(_js_record_keys(decls, raw, alternative) or {})
    return out


def _record_fields(found: Surfaces, record: dict, surface: str) -> Optional[Tuple[str, Set[str]]]:
    field = _field(surface)
    if surface == "py":
        name = _in_tree(found.py.classes, record, field)
        return (name, set(found.py.classes[name].fields)) if name else None
    js = found.js[surface]
    name = _in_tree(js.decls, record, field)
    if not name:
        return None
    return name, set(_js_record_keys({**js.private, **js.decls}, js.raw, name) or ())


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
    """JS results carry no envelope: no declaration is named like a retired result type, no `export interface`
    has a `status` field, and no callable returns a retired type. A scalar result is the scalar, an allocated
    key a `number` (or the type the ledger's `returns: union:<Type>` names, or the handle class of a collection
    published as handles), a natural-key lookup `T | null`;
    `kind: field` renames are applied.
    """
    found = surfaces(ctx)
    lookups = _optional_lookups(ctx)
    for surface in JS_SURFACES:
        js = found.js[surface]
        pool = _callables(ctx, surface)
        for name in _RETIRED_RESULT_TYPES:
            if name in js.decls:
                yield _violation(
                    ctx, "R8.js", surface, _key(ctx, surface, name), name, f"`{name}` is a retired result type"
                )
        for decl in sorted(js.decls.values()):
            if decl.kind == "interface" and any(m.name == "status" and not m.is_method for m in decl.members):
                yield _violation(
                    ctx,
                    "R8.js",
                    surface,
                    _key(ctx, surface, decl.name),
                    "status",
                    f"`{decl.name}` has a `status` field",
                )
        for name, member in sorted(pool.items()):
            for retired in _RETIRED_RESULT_TYPES:
                if _mentions(member.type, retired):
                    yield _violation(
                        ctx,
                        "R8.js",
                        surface,
                        _key(ctx, surface, name),
                        retired,
                        f"`{name}` returns `{member.type}`; a failure throws and a success is the payload",
                    )
        for key, entry in sorted(ctx.ledger.entries.items()):
            if not isinstance(entry, dict):
                continue
            name, member = _member_for(pool, entry, "js")
            if member is None:
                continue
            if key in lookups:
                if "null" not in _top_union(member.type):
                    yield _violation(
                        ctx, "R8.js", surface, key, "optional", f"`{name}` returns `{member.type}`, not `... | null`"
                    )
                continue
            spec = entry.get("returns")
            want = _SCALAR_JS.get(spec) if isinstance(spec, str) else None
            typed = _type_name(ctx, spec, "js") if isinstance(spec, str) else None
            fn = _tree_c(ctx, entry)
            # An allocated key is a `number` unless the ledger types the result (a handle the facade wraps).
            if fn is not None and c_verb(ctx.target_c(fn.name)) == "add" and _allocates_key(fn) and typed is None:
                # An add of a collection published as handles returns the new element's handle.
                want = _handle_result(ctx, key) or "number"
            if want and member.type != want:
                yield _violation(
                    ctx, "R8.js", surface, key, "result", f"`{name}` returns `{member.type}`, not `{want}`"
                )
        yield from _field_renames(ctx, surface, "R8.js")


def _handle_result(ctx: Context, key: str) -> Optional[str]:
    """The handle class of the collection whose `ops` list `key`, when it is published as handles."""
    for name, cls in handle_collections(ctx.ledger.collections).items():
        if key in ((ctx.ledger.collections.get(name) or {}).get("ops") or []):
            return cls
    return None


def _is_tuple(annotation: str, classes: Dict[str, PyClass]) -> bool:
    annotation = annotation.replace("'", "").replace('"', "")
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
        if not name.rpartition(".")[2].startswith("_") and fn.returns and _is_tuple(fn.returns, py.classes):
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


# A struct field counts in `uint32_t` (R6.c.field forbids `size_t` there); a function parameter in `size_t`.
_MULTI_FIELD_COUNT_TYPES = ("uint32_t", "size_t")
_MULTI_PARAM_COUNT_TYPES = ("size_t",)


def _multi_problems(members: Sequence[CParam], owner: str, count_types: Sequence[str]) -> Iterator[Tuple[str, str]]:
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
            and follower.pointer_depth == 0
            and follower.role_name.endswith("count")
            and (role in _A1_MULTI_NAMES or member.base_type == CELL_RECT)
        ):
            if (
                member.base_type != CELL_RECT
                or member.name != "ranges"
                or follower.ctype not in count_types
                or follower.name != "range_count"
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
        for disc, message in _multi_problems(typedef.fields, typedef.name, _MULTI_FIELD_COUNT_TYPES):
            yield ctx.c_violation("R11.multi", typedef.name, disc, f"{typedef.name}: {message}", typedef.line)
    for fn in ctx.functions:
        target = ctx.target_c(fn.name)
        if c_verb(target) in _TEXT_VERBS:
            continue
        for disc, message in _multi_problems(fn.params, target, _MULTI_PARAM_COUNT_TYPES):
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
    """A range on another sheet is `fm_sheet_rect_t { sheet_index; fm_cell_rect_t rect; }` in C and
    `SheetRect { sheet, rect }` in the bindings; no other record pairs a sheet index with a rectangle. A
    record keyed by a sheet name (`ParsedReference`) carries its rectangle as `rect`, not loose corners.
    """
    for typedef in ctx.typedefs:
        if typedef.kind != "struct":
            continue
        fields = {m.name: m for m in typedef.fields}
        if typedef.name == SHEET_RECT:
            shape = tuple((m.ctype, m.name) for m in typedef.fields)
            if shape[1:] != ((CELL_RECT, "rect"),) or shape[:1] not in (
                (("size_t", "sheet_index"),),
                (("uint32_t", "sheet_index"),),
            ):
                yield ctx.c_violation(
                    "R11.sheet",
                    typedef.name,
                    NO_DISCRIMINATOR,
                    f"{SHEET_RECT} must be exactly `<index> sheet_index; {CELL_RECT} rect;` (R6.c.field types the index)",
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
            if "sheetName" in names and _edge_quad(names, _JS_EDGE_QUADS + _NAMED_EDGE_QUADS):
                yield _violation(
                    ctx,
                    "R11.sheet",
                    surface,
                    _key(ctx, surface, decl.name),
                    "rect",
                    f"`{decl.name}` spells its rectangle as loose coordinates; use `rect: CellRect`",
                )
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
        if "sheet_name" in names and _edge_quad(names, _EDGE_QUADS + _NAMED_EDGE_QUADS):
            yield _violation(
                ctx,
                "R11.sheet",
                "py",
                _key(ctx, "py", name),
                "rect",
                f"`{name}` spells its rectangle as loose coordinates; use `rect: CellRect`",
            )
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
    yield from _one_sided(ctx, {s: set(js[s].sheet) for s in JS_SURFACES}, "sheet-method")
    yield from _one_sided(ctx, {s: {f"Workbook.{n}" for n in js[s].statics} for s in JS_SURFACES}, "static")
    yield from _one_sided(ctx, {s: set(js[s].functions) - {"createFormulon"} for s in JS_SURFACES}, "function")


_ENTRY_REQUIRED_EXPORTS = frozenset({"FormulonError", "createFormulon"})
# Host-side values and pure helpers both packages also export by name, besides carrying them as `Formulon` members.
_ENTRY_SHARED_EXPORTS = frozenset({"ExcelError", "enumName", "mergeFunctionMetadata", "toNative"})
_SCALAR_CONSTANT_RE = re.compile(r"=?\s*-?\d+")
_MEMBERS_LIST_RE = re.compile(r"\bFORMULON_MEMBERS\s*=\s*(?:Object\.freeze\(\s*)?\[(.*?)\]", re.S)
_MJS_COMMENT_RE = re.compile(r"/\*.*?\*/|//[^\n]*", re.S)


def _is_constant(decl: TsDecl) -> bool:
    if decl.kind == "enum":
        return True
    if decl.kind != "const":
        return False
    return (
        decl.text.startswith(("Readonly<{", "Readonly<Record<")) or _SCALAR_CONSTANT_RE.fullmatch(decl.text) is not None
    )


def _formulon_members(ctx: Context) -> Optional[Set[str]]:
    """The names `FORMULON_MEMBERS` lists in the npm package's `common.mjs`, or None when it is not declared."""
    path = ctx.repo_root / NPM_COMMON_MJS.relative_to(REPO_ROOT)
    if not path.is_file():
        return None
    match = _MEMBERS_LIST_RE.search(path.read_text(encoding="utf-8"))
    if match is None:
        return None
    body = _MJS_COMMENT_RE.sub(" ", match.group(1))
    return {a or b for a, b in re.findall(r"'([^']*)'|\"([^\"]*)\"", body)}


@_owned("R17.entry")
def _rule_entry_shape(ctx: Context) -> Iterable[Violation]:
    """Each declaration file's `Formulon` interface has exactly the members `FORMULON_MEMBERS` lists, and its
    runtime named exports are the constant tables, `FormulonError`, `createFormulon` and the host-side values
    `ExcelError` / `enumName` / `mergeFunctionMetadata` / `toNative` (which must also be `Formulon` members) only.
    """
    js = surfaces(ctx).js
    listed = _formulon_members(ctx)
    for surface in JS_SURFACES:
        key = _key(ctx, surface, ENTRY_INTERFACE)
        declared = set(js[surface].module_object)
        if listed is None:
            yield _violation(
                ctx, "R17.entry", surface, key, "FORMULON_MEMBERS", "`FORMULON_MEMBERS` is not declared in common.mjs"
            )
        else:
            for name in sorted(declared - listed):
                yield _violation(
                    ctx, "R17.entry", surface, key, name, f"`{ENTRY_INTERFACE}.{name}` is not in `FORMULON_MEMBERS`"
                )
            for name in sorted(listed - declared):
                yield _violation(
                    ctx,
                    "R17.entry",
                    surface,
                    key,
                    name,
                    f"`FORMULON_MEMBERS` lists `{name}`; `{ENTRY_INTERFACE}` lacks it",
                )
        values = {d.name: d for d in js[surface].decls.values() if d.kind in ("const", "function", "class", "enum")}
        for name in sorted(values):
            if name in _ENTRY_SHARED_EXPORTS and name in declared:
                continue
            if name not in _ENTRY_REQUIRED_EXPORTS and not _is_constant(values[name]):
                yield _violation(
                    ctx,
                    "R17.entry",
                    surface,
                    _key(ctx, surface, name),
                    name,
                    f"`{name}` is a named export; only constants, `FormulonError`, `createFormulon` and the "
                    f"`Formulon` members {' / '.join(f'`{n}`' for n in sorted(_ENTRY_SHARED_EXPORTS))} are",
                )
        for name in sorted(_ENTRY_REQUIRED_EXPORTS - set(values)):
            yield _violation(ctx, "R17.entry", surface, _key(ctx, surface, name), name, f"`{name}` is not exported")


# ---------------------------------------------------------------------------
# R16: the JS `Value` union.
# ---------------------------------------------------------------------------


@_owned("R16.js")
def _rule_value_union(ctx: Context) -> Iterable[Violation]:
    """`Value` is a union of eight interfaces, each with a `readonly kind: typeof ValueKind.X` of its own;
    the `Bool` member's `boolean` is `boolean`, the `Error` member's `errorCode` is `ErrorCode`.
    """
    for surface in JS_SURFACES:
        js = surfaces(ctx).js[surface]
        key = _key(ctx, surface, "Value")
        decl = js.decls.get("Value")
        if decl is None or decl.kind != "type":
            yield _violation(ctx, "R16.js", surface, key, "union", "`Value` is not an exported type alias")
            continue
        alternatives = _top_union(decl.text)
        if len(alternatives) != 8:
            yield _violation(
                ctx, "R16.js", surface, key, "union", f"`Value` has {len(alternatives)} members, not eight"
            )
        kinds: Dict[str, List[str]] = {}
        for alt in alternatives:
            member = js.decls.get(alt)
            if member is None or member.kind != "interface":
                yield _violation(ctx, "R16.js", surface, key, alt, f"`{alt}` is not an exported interface")
                continue
            fields = {m.name: m for m in member.members if not m.is_method}
            tag = fields.get("kind")
            if tag is None or not tag.readonly or re.fullmatch(r"typeof ValueKind\.\w+", tag.type) is None:
                yield _violation(
                    ctx,
                    "R16.js",
                    surface,
                    key,
                    f"{alt}.kind",
                    f"`{alt}.kind` is not `readonly kind: typeof ValueKind.X`",
                )
                continue
            kinds.setdefault(tag.type, []).append(alt)
            for kind, field, want in (("Bool", "boolean", "boolean"), ("Error", "errorCode", "ErrorCode")):
                if tag.type == f"typeof ValueKind.{kind}" and (field not in fields or fields[field].type != want):
                    yield _violation(ctx, "R16.js", surface, key, f"{alt}.{field}", f"`{alt}.{field}` is not `{want}`")
        for tag, owners in sorted(kinds.items()):
            if len(owners) > 1:
                yield _violation(
                    ctx, "R16.js", surface, key, tag, f"`{tag}` is the kind of {', '.join(f'`{o}`' for o in owners)}"
                )


# ---------------------------------------------------------------------------
# R13 / R14: enum-typed parameters and abbreviations, on both binding languages.
# ---------------------------------------------------------------------------


class _Member(NamedTuple):
    owner: str
    name: str
    is_method: bool
    params: Tuple[Tuple[str, str], ...]
    type: str
    js: bool

    @property
    def where(self) -> str:
        return self.name if self.owner == self.name else f"{self.owner}.{self.name}"


def _members_of(ctx: Context, surface: str) -> Iterator[_Member]:
    """Every public field and callable of a binding surface, with the type that declares it."""
    found = surfaces(ctx)
    if surface != "py":
        for decl in sorted(found.js[surface].decls.values()):
            for m in decl.members:
                yield _Member(decl.name, m.name, m.is_method, m.params, m.type, True)
        return
    py = found.py
    public = set(py.exports) | {WORKBOOK_CLASS, SHEET_CLASS}
    for cls in sorted(py.classes.values()):
        if cls.name not in public:
            continue
        for name, annotation in cls.fields.items():
            if not name.startswith("_"):
                yield _Member(cls.name, name, False, (), annotation, False)
        for name, fn in cls.methods.items():
            if not name.startswith("_"):
                yield _Member(cls.name, name, True, fn.params, fn.returns, False)
    for name, fn in sorted(py.functions.items()):
        if name in public:
            yield _Member(name, name, True, fn.params, fn.returns, False)


def _identifier(name: str) -> bool:
    return re.fullmatch(r"[A-Za-z_]\w*", name) is not None


# Parameter and field names that carry an enum, whatever enum it is.
_ENUM_NAMES = {
    True: frozenset({"format", "errorCode", "level", "status"}),
    False: frozenset({"format", "error_code", "level", "status"}),
}


# Functions whose `level` is an outline depth, a bounded integer rather than a log level.
_OUTLINE_DEPTH_FUNCTIONS = frozenset({"setRowOutline", "setColumnOutline", "set_row_outline", "set_column_outline"})


def _bare_number(annotation: str, js: bool) -> bool:
    text = annotation.strip()
    inner = re.fullmatch(r"(?:typing\.)?Optional\[(.*)\]", text)
    return ("number" if js else "int") in _top_union(inner.group(1) if inner else text)


@_owned("R13.enum")
def _rule_enum_types(ctx: Context) -> Iterable[Violation]:
    """A parameter or field named `format`, `errorCode` (`error_code`), `level` or `status` is typed by an enum;
    a bare `number` (`int`), alone or in a union, is not one. The outline-depth setters' `level` is exempt.
    """
    for surface in BINDING_SURFACES:
        for m in _members_of(ctx, surface):
            slots = [(m.name, m.type, f"{m.where}")] if not m.is_method else []
            if m.name not in _OUTLINE_DEPTH_FUNCTIONS:
                slots += [(pname, ptype, f"{m.where}({pname})") for pname, ptype in m.params]
            for name, annotation, path in slots:
                if name in _ENUM_NAMES[m.js] and annotation and _bare_number(annotation, m.js):
                    yield _violation(
                        ctx,
                        "R13.enum",
                        surface,
                        _key(ctx, surface, m.name if m.is_method else m.owner),
                        path,
                        f"`{path}` is `{annotation}`; type it with an enum",
                    )


def _field_accessor(name: str, js: bool, fields: Set[str]) -> bool:
    """True for `get<Field>` / `set<Field>` (`get_<field>` / `set_<field>`) over an existing field."""
    head = name[:3]
    if head not in ("get", "set"):
        return False
    rest = name[3:] if js else name[4:]
    return (rest[:1].lower() + rest[1:] if js else rest) in fields


@_owned("R14.abbrev")
def _rule_abbreviations(ctx: Context) -> Iterable[Violation]:
    """A method name does not use the word `col` / `cols` (`column` / `columns`), except a `get` / `set` over an
    existing field, which follows the field's spelling. A field or parameter name does not use `off`
    (`offset`); `col` is a coordinate word there and stays. No name uses `idx` (`index`).
    """
    for surface in BINDING_SURFACES:
        members = list(_members_of(ctx, surface))
        fields = {m.name for m in members if not m.is_method}
        for m in members:
            names = [(m.name, m.where, m.is_method)]
            names += [(pname, f"{m.where}({pname})", False) for pname, _ in m.params]
            for name, path, method in names:
                if not _identifier(name):
                    continue
                words = set(_words(name))
                problems = []
                if method and words & {"col", "cols"} and not _field_accessor(name, m.js, fields):
                    problems.append(("col", "`col` / `cols` is spelled `column` / `columns` in a method name"))
                if not method and "off" in words:
                    problems.append(("off", "`off` is spelled `offset`"))
                if "idx" in words:
                    problems.append(("idx", "`idx` is spelled `index`"))
                for word, text in problems:
                    yield _violation(
                        ctx,
                        "R14.abbrev",
                        surface,
                        _key(ctx, surface, m.name if m.is_method else m.owner),
                        f"{word}:{path}",
                        f"`{path}`: {text}",
                    )
