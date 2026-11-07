"""Scans the public API surfaces and proposes a name-ledger entry per exported symbol.

The surfaces are the C ABI headers (`formulon_c.h`, `formulon_status.h`),
the two JS packages (embind registrations + `src/wasm/formulon.d.ts`, the
N-API registrations + `packages/npm-native/index.d.ts`) and the Python
package (`__all__` plus the public `Workbook` methods, read with `ast`).

Binding methods are tied to the C functions they reach: the binding sources
are read as a call graph and every `fm_*` name a method references, directly
or through a helper, is attributed to it. New names follow the rules the
`api-conventions` checker holds (`tools/dev/binding_drift/api_conventions.py`);
an entry the rules do not settle carries `flag: <reason>` for a human to
decide. `tools/api/names.yaml` is this output after curation, so re-running
writes to stdout unless `-o` is given.
"""

from __future__ import annotations

import argparse
import ast
import re
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Set, Tuple

REPO_ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO_ROOT / "tools" / "dev"))

from binding_drift.api_conventions import (  # noqa: E402
    C_VERBS,
    CFunction,
    CTypedef,
    derive_js,
    derive_py,
    derive_type,
    parse_header,
    split_c_name,
)

CAPI_HEADER = Path("src/c_api/formulon_c.h")
STATUS_HEADER = Path("src/c_api/formulon_status.h")
WASM_REGISTER = Path("src/wasm/parts/bindings_register.cpp")
WASM_SOURCES = (Path("src/wasm/parts"), Path("src/wasm/generated"))
WASM_DTS = Path("src/wasm/formulon.d.ts")
NODE_CLASS = Path("src/node_addon/parts/workbook_class.cc")
NODE_ADDON = Path("src/node_addon/addon.cc")
NODE_SOURCES = (Path("src/node_addon/parts"), Path("src/node_addon/generated"), Path("src/node_addon"))
NODE_DTS = Path("packages/npm-native/index.d.ts")
PY_PKG = Path("packages/python/formulon")
# The modules of the Python handle classes; a module not written yet is skipped.
PY_HANDLE_MODULES = ("table.py", "pivot.py", "pivot_cache.py", "drawing.py")
# The facade source whose generated tables name the JS handle methods `Class.method`.
NPM_COMMON = Path("packages/npm/common.mjs")
_JS_HANDLE_TABLES = ("HANDLE_FORWARDERS", "HANDLE_PROPERTIES")

JS_SURFACES = ("wasm", "native")
LEDGER_SURFACES = ("c", "wasm", "native", "py")

# C functions every binding method may touch without being "about" them.
_PLUMBING = frozenset(
    {
        "fm_last_error_message",
        "fm_last_error_context",
        "fm_status_string",
        "fm_buffer_free",
    }
)


# ---------------------------------------------------------------------------
# Scanning.
# ---------------------------------------------------------------------------


@dataclass
class Member:
    """One binding-surface symbol and the C functions it reaches."""

    name: str
    kind: str  # method | static | free | type | enum | const
    calls: Set[str] = field(default_factory=set)
    detail: str = ""  # d.ts declaration text or Python annotation summary


@dataclass
class Surfaces:
    functions: Dict[str, CFunction]
    typedefs: Dict[str, CTypedef]
    constants: Dict[str, str]  # constant -> owning enum typedef, or "" for a #define
    status: List[str]
    js: Dict[str, Dict[str, Member]]  # surface -> js name -> member
    py: Dict[str, Member]  # py name -> member
    py_enums: Dict[str, str]  # Python IntEnum class -> C enum typedef it cites
    py_types: Dict[str, str]  # Python class -> C typedef its docstring cites

    def exported(self) -> List[Tuple[str, str]]:
        """Every exported symbol as `(surface, name)`; JS statics are `Workbook.<name>`."""
        out = [("c", n) for n in self.functions]
        out += [("c", n) for n in self.typedefs]
        out += [("c", n) for n in self.constants]
        out += [("c", n) for n in self.status]
        for surface in JS_SURFACES:
            out += [(surface, n) for n in self.js[surface]]
        out += [("py", n) for n in self.py]
        return out


_COMMENT_RE = re.compile(r"/\*.*?\*/|//[^\n]*", re.S)
_ENUM_BODY_RE = re.compile(r"typedef\s+enum\s*\{([^}]*)\}\s*([A-Za-z_]\w*)\s*;")
_DEFINE_RE = re.compile(r"^#define\s+(FM_[A-Z0-9_]+)\s+\S", re.M)
_NON_SYMBOL_DEFINES = frozenset({"FM_API"})


def _read(root: Path, rel: Path) -> str:
    return (root / rel).read_text(encoding="utf-8")


def _scan_c(root: Path):
    header = _read(root, CAPI_HEADER)
    functions, typedefs = parse_header(header)
    stripped = _COMMENT_RE.sub(" ", header)
    constants: Dict[str, str] = {}
    for body, name in _ENUM_BODY_RE.findall(stripped):
        for item in body.split(","):
            item = item.strip()
            if item:
                constants[item.split("=")[0].strip()] = name
    for name in _DEFINE_RE.findall(header):
        if name not in _NON_SYMBOL_DEFINES:
            constants.setdefault(name, "")
    status_path = root / STATUS_HEADER
    status = _DEFINE_RE.findall(status_path.read_text(encoding="utf-8")) if status_path.is_file() else []
    return {f.name: f for f in functions}, {t.name: t for t in typedefs}, constants, status


# -- C++ binding sources as a call graph -------------------------------------

_CPP_DEF_RE = re.compile(
    r"^[A-Za-z_][\w:<>,\s\*&]*?\b((?:[A-Za-z_]\w*::)*[A-Za-z_]\w*)\s*\(([^;{}]*?)\)\s*(?:const\s*)?(?:noexcept\s*)?\{",
    re.M,
)
_IDENT_RE = re.compile(r"\b([A-Za-z_]\w*)\b")


def _brace_body(text: str, open_at: int) -> str:
    depth = 0
    for i in range(open_at, len(text)):
        if text[i] == "{":
            depth += 1
        elif text[i] == "}":
            depth -= 1
            if depth == 0:
                return text[open_at + 1 : i]
    return text[open_at + 1 :]


def _cpp_graph(root: Path, dirs: Sequence[Path]) -> Dict[str, Set[str]]:
    """Unqualified function name -> identifiers its definitions reference."""
    graph: Dict[str, Set[str]] = defaultdict(set)
    for rel in dirs:
        for path in sorted((root / rel).glob("*.c*")):
            if path.suffix not in (".cc", ".cpp"):
                continue
            text = _COMMENT_RE.sub(" ", path.read_text(encoding="utf-8"))
            for match in _CPP_DEF_RE.finditer(text):
                name = match.group(1).split("::")[-1]
                if name in ("if", "for", "while", "switch", "return"):
                    continue
                body = _brace_body(text, match.end() - 1)
                graph[name].update(_IDENT_RE.findall(body))
    return graph


def _reach(graph: Dict[str, Set[str]], start: str, limit: int = 4) -> Set[str]:
    """`fm_*` names reachable from `start` through helper definitions."""
    seen = {start}
    frontier = [start]
    found: Set[str] = set()
    for _ in range(limit):
        nxt = []
        for name in frontier:
            for ident in graph.get(name, ()):
                if ident.startswith("fm_"):
                    found.add(ident)
                elif ident in graph and ident not in seen:
                    seen.add(ident)
                    nxt.append(ident)
        frontier = nxt
    return {f for f in found if not f.endswith("_t")}


# -- d.ts --------------------------------------------------------------------

_DTS_TOP_RE = re.compile(
    r"^export\s+(?:declare\s+)?(?:default\s+)?(interface|type|enum|const|function)\s+([A-Za-z_$][\w$]*)",
    re.M,
)
_DTS_MEMBER_RE = re.compile(r"^ {2}(?:readonly\s+)?([A-Za-z_$][\w$]*)\??\s*[(:<]", re.M)


def _dts_block(text: str, name: str) -> str:
    match = re.search(r"^export interface %s\b[^{]*\{" % re.escape(name), text, re.M)
    return _brace_body(text, match.end() - 1) if match else ""


def _dts_member_decls(body: str) -> Dict[str, str]:
    """Top-level members of an interface body -> declaration text up to `;`."""
    out: Dict[str, str] = {}
    for match in _DTS_MEMBER_RE.finditer(body):
        end = body.find(";\n", match.start())
        out.setdefault(match.group(1), " ".join(body[match.start() : end if end > 0 else None].split()))
    return out


def _scan_dts(root: Path, rel: Path):
    text = _read(root, rel)
    stripped = _COMMENT_RE.sub(" ", text)
    tops: Dict[str, str] = {}
    for kind, name in _DTS_TOP_RE.findall(stripped):
        tops.setdefault(name, kind)
    workbook = _dts_member_decls(_dts_block(stripped, "Workbook"))
    statics = _dts_member_decls(_dts_block(stripped, "WorkbookCtor"))
    module = _dts_member_decls(_dts_block(stripped, "Formulon"))
    # Both enum spellings: a TS `enum`, and a frozen `as const` table with its union type.
    enums = set(re.findall(r"^export enum (\w+)", stripped, re.M))
    enums |= set(re.findall(r"^export const (\w+): Readonly<\{", stripped, re.M))
    return tops, workbook, statics, module, enums


# Addon exports the package entry point calls and the public surface does not carry.
NODE_INTERNAL_EXPORTS = frozenset({"__setErrorClass"})


def _scan_js(root: Path) -> Dict[str, Dict[str, Member]]:
    js: Dict[str, Dict[str, Member]] = {}

    # WASM: embind registration -> JsWorkbook member or free function.
    register = _COMMENT_RE.sub(" ", _read(root, WASM_REGISTER))
    graph = _cpp_graph(root, WASM_SOURCES)
    targets: Dict[str, Tuple[str, str]] = {}
    for kind, name, target in re.findall(r'\.(function|class_function)\("([^"]+)",\s*&JsWorkbook::(\w+)', register):
        targets.setdefault(name, ("static" if kind == "class_function" else "method", target))
    for name, target in re.findall(r'(?<![.\w])function\("([^"]+)",\s*&(\w+)', register):
        targets.setdefault(name, ("free", target))
    js["wasm"] = _js_members(root, WASM_DTS, targets, graph)

    # Native: N-API registration table and addon exports.
    klass = _read(root, NODE_CLASS)
    addon = _read(root, NODE_ADDON)
    graph = _cpp_graph(root, NODE_SOURCES)
    targets = {}
    for target, name in re.findall(r'(?:Guarded)?InstanceMethod<&Workbook::(\w+)[^>]*>\("([^"]+)"\)', klass):
        targets.setdefault(name, ("method", target))
    for target, name in re.findall(r'StaticMethod<&Workbook::(\w+)>\("([^"]+)"\)', klass):
        targets.setdefault(name, ("static", target))
    for name, target in re.findall(r'exports\.Set\("([^"]+)",\s*Napi::Function::New\(env,\s*&[\w:]*?(\w+)\b', addon):
        if name not in NODE_INTERNAL_EXPORTS:
            targets.setdefault(name, ("free", target))
    js["native"] = _js_members(root, NODE_DTS, targets, graph)
    for name in _js_handle_members(root):
        for surface in JS_SURFACES:
            js[surface].setdefault(name, Member(name, "method"))
    return js


def _js_handle_members(root: Path) -> List[str]:
    """`Class.method` of every handle method the generated facade tables of common.mjs name."""
    path = root / NPM_COMMON
    if not path.is_file():
        return []
    region = re.search(r"// <facade-tables>(.*?)// </facade-tables>", path.read_text(encoding="utf-8"), re.S)
    out: List[str] = []
    for table in _JS_HANDLE_TABLES:
        found = re.search(r"\b%s\s*=\s*Object\.freeze\(\{(.*?)\}\);" % table, region.group(1), re.S) if region else None
        if found:
            out += re.findall(r"^\s*'([A-Z]\w*\.\w+)':", found.group(1), re.M)
    return out


def _js_members(root: Path, rel: Path, targets: Dict[str, Tuple[str, str]], graph) -> Dict[str, Member]:
    """d.ts exports plus registrations, each with the C functions its binding body reaches."""
    tops, workbook, statics, module, enums = _scan_dts(root, rel)
    out: Dict[str, Member] = {}
    for name, decl in workbook.items():
        kind, target = targets.get(name, ("method", ""))
        out[name] = Member(name, "method", _reach(graph, target) if target else set(), decl)
    for name, decl in statics.items():
        _, target = targets.get(name, ("static", ""))
        out[f"Workbook.{name}"] = Member(name, "static", _reach(graph, target) if target else set(), decl)
    for name, decl in module.items():
        if name == "Workbook" or not decl.split(name, 1)[1].lstrip("?").startswith("("):
            continue
        _, target = targets.get(name, ("free", ""))
        out[name] = Member(name, "free", _reach(graph, target) if target else set(), decl)
    for name, kind in tops.items():
        if name in out or name == "_default":
            continue
        if kind == "function":
            _, target = targets.get(name, ("free", ""))
            out[name] = Member(name, "free", _reach(graph, target) if target else set())
        elif kind == "enum" or name in enums:
            out[name] = Member(name, "enum")
        elif kind == "const":
            out[name] = Member(name, "const")
        else:
            out[name] = Member(name, "type")
    # A registration the d.ts omits is still exported.
    for name, (kind, target) in targets.items():
        key = f"Workbook.{name}" if kind == "static" else name
        if key not in out:
            out[key] = Member(name, kind, _reach(graph, target))
    return out


# -- Python ------------------------------------------------------------------

_PY_C_TYPE_RE = re.compile(r"``(fm_[a-z0-9_]+)``")


def _py_refs(node: ast.AST) -> Tuple[Set[str], Set[str]]:
    """(`fm_*` names, self/Workbook/module callees) referenced inside `node`."""
    fms: Set[str] = set()
    callees: Set[str] = set()
    for sub in ast.walk(node):
        if isinstance(sub, ast.Attribute):
            if sub.attr.startswith("fm_"):
                fms.add(sub.attr)
            elif isinstance(sub.value, ast.Name) and sub.value.id in ("self", "cls", "Workbook"):
                callees.add(sub.attr)
        elif isinstance(sub, ast.Constant) and isinstance(sub.value, str) and re.fullmatch(r"fm_[a-z0-9_]+", sub.value):
            fms.add(sub.value)
        elif isinstance(sub, ast.Name):
            callees.add(sub.id)
    return fms, callees


def _scan_py(root: Path):
    pkg = root / PY_PKG
    graph: Dict[str, Tuple[Set[str], Set[str]]] = {}
    methods: Dict[str, ast.FunctionDef] = {}
    module_defs: Dict[str, ast.AST] = {}
    py_enums: Dict[str, str] = {}
    py_types: Dict[str, str] = {}
    for path in sorted(pkg.glob("*.py")):
        tree = ast.parse(path.read_text(encoding="utf-8"))
        for node in tree.body:
            if isinstance(node, ast.FunctionDef):
                graph.setdefault(node.name, _py_refs(node))
                module_defs.setdefault(node.name, node)
            elif (
                isinstance(node, ast.Assign)
                and len(node.targets) == 1
                and isinstance(node.targets[0], ast.Name)
                and isinstance(node.value, ast.Name)
                and isinstance(module_defs.get(node.value.id), ast.FunctionDef)
            ):  # `alias = function`
                module_defs.setdefault(node.targets[0].id, module_defs[node.value.id])
                graph.setdefault(node.targets[0].id, graph[node.value.id])
            elif isinstance(node, ast.ClassDef):
                module_defs.setdefault(node.name, node)
                doc = ast.get_docstring(node) or ""
                cited = _PY_C_TYPE_RE.findall(doc)
                bases = {getattr(b, "id", getattr(b, "attr", "")) for b in node.bases}
                if cited:
                    (py_enums if bases & {"IntEnum", "IntFlag"} else py_types).setdefault(node.name, cited[0])
                if node.name == "Workbook" and path.name == "workbook.py":
                    for item in node.body:
                        if isinstance(item, ast.FunctionDef):
                            methods[item.name] = item
                            graph[f"Workbook.{item.name}"] = _py_refs(item)

    def reach(start: str) -> Set[str]:
        seen = {start}
        frontier = [start]
        found: Set[str] = set()
        for _ in range(5):
            nxt = []
            for name in frontier:
                fms, callees = graph.get(name, (set(), set()))
                found |= fms
                for callee in callees:
                    for candidate in (f"Workbook.{callee}", callee):
                        if candidate in graph and candidate not in seen:
                            seen.add(candidate)
                            nxt.append(candidate)
            frontier = nxt
        return {f for f in found if not f.endswith("_t")}

    out: Dict[str, Member] = {}
    for name, node in methods.items():
        if name.startswith("_"):
            continue
        static = any(getattr(d, "id", "") in ("staticmethod", "classmethod") for d in node.decorator_list)
        prop = any(getattr(d, "id", "") == "property" for d in node.decorator_list)
        key = f"Workbook.{name}" if static else name
        out[key] = Member(
            name,
            "static" if static else ("property" if prop else "method"),
            reach(f"Workbook.{name}"),
            _py_signature(node),
        )
    for module in PY_HANDLE_MODULES:
        path = pkg / module
        if not path.is_file():
            continue
        for node in ast.parse(path.read_text(encoding="utf-8")).body:
            if not isinstance(node, ast.ClassDef):
                continue
            for item in node.body:
                if not isinstance(item, ast.FunctionDef) or item.name.startswith("_"):
                    continue
                prop = any(getattr(d, "id", "") == "property" for d in item.decorator_list)
                name = f"{node.name}.{item.name}"
                out.setdefault(
                    name, Member(name, "property" if prop else "method", _py_refs(item)[0], _py_signature(item))
                )
    init = ast.parse((pkg / "__init__.py").read_text(encoding="utf-8"))
    for node in init.body:
        if isinstance(node, ast.Assign) and getattr(node.targets[0], "id", "") == "__all__":
            for elt in node.value.elts:
                name = elt.value
                definition = module_defs.get(name)
                if isinstance(definition, ast.FunctionDef):
                    out[name] = Member(name, "free", reach(name), _py_signature(definition))
                elif isinstance(definition, ast.ClassDef):
                    bases = {getattr(b, "id", getattr(b, "attr", "")) for b in definition.bases}
                    out[name] = Member(name, "enum" if bases & {"IntEnum", "IntFlag"} else "type")
                else:
                    out[name] = Member(name, "const")
    return out, py_enums, py_types


def _py_signature(node: ast.FunctionDef) -> str:
    args = []
    for arg in node.args.args + node.args.kwonlyargs:
        if arg.arg in ("self", "cls"):
            continue
        args.append(f"{arg.arg}: {ast.unparse(arg.annotation) if arg.annotation else '?'}")
    ret = ast.unparse(node.returns) if node.returns else "?"
    return f"({', '.join(args)}) -> {ret}"


def scan(root: Path = REPO_ROOT) -> Surfaces:
    functions, typedefs, constants, status = _scan_c(root)
    js = _scan_js(root)
    py, py_enums, py_types = _scan_py(root)
    return Surfaces(functions, typedefs, constants, status, js, py, py_enums, py_types)


# ---------------------------------------------------------------------------
# C name proposal.
# ---------------------------------------------------------------------------

# Process functions whose new names the rules list outright.
PROCESS_RENAMES = {
    "fm_version_string": "fm_get_version",
    "fm_status_string": "fm_get_status_name",
    "fm_last_error_message": "fm_get_last_error_message",
    "fm_last_error_context": "fm_get_last_error_context",
    "fm_error_display_name": "fm_get_error_display_name",
    "fm_function_count": "fm_get_function_count",
    "fm_function_name_at": "fm_get_function_name_at",
    "fm_function_metadata": "fm_get_function_metadata",
    "fm_function_localize": "fm_localize_function_name",
    "fm_function_canonicalize": "fm_canonicalize_function_name",
    "fm_formula_localize": "fm_localize_formula",
    "fm_formula_canonicalize": "fm_canonicalize_formula",
    "fm_locale_facts": "fm_get_locale_facts",
    "fm_locale_error_name_count": "fm_get_locale_error_name_count",
    "fm_locale_error_name": "fm_get_locale_error_name_at",
    "fm_text_compare": "fm_compare_text",
    "fm_collation_classes": "fm_get_collation_classes",
    "fm_buffer_free": "fm_buffer_destroy",
}

# Leading word sequences (after the scope) the verb rules rewrite outright.
PHRASE_RENAMES = {
    ("chart", "open"): ("get", "chart"),
    ("chart", "data", "open"): ("get", "chart", "data"),
    ("chart", "preview", "open"): ("get", "chart", "preview"),
    ("paginate",): ("get", "pagination"),
    ("precedents",): ("get", "precedents"),
    ("dependents",): ("get", "dependents"),
    ("probe", "image"): ("get", "image", "info"),
    ("list", "invalid", "cells"): ("get", "invalid", "cells"),
    ("edit", "threaded", "comment"): ("update", "threaded", "comment"),
    ("insert", "image"): ("add", "image"),
    ("insert", "shape"): ("add", "shape"),
    ("insert", "chart"): ("add", "chart"),
    ("column", "chars", "to", "pt"): ("convert", "column", "chars", "to", "pt"),
    ("column", "pt", "to", "chars"): ("convert", "column", "pt", "to", "chars"),
}
VERB_REWRITES = {"reset": "clear"}
# Noun synonyms; the left side may not appear in a name.
NOUN_SYNONYMS = {
    ("horizontal", "break"): ("row", "break"),
    ("vertical", "break"): ("col", "break"),
    ("idx",): ("index",),
}
# Words that read as verbs in a name although the vocabulary lacks them.
UNLISTED_VERBS = frozenset({"project", "show"})
# A vocabulary word after one of these is part of a noun phrase (`calc_on_save`).
_PREPOSITIONS = frozenset({"on", "to", "with", "in", "for", "of", "by", "as"})

# The R3 classification the rules settle. Keys are collection nouns.
INDEX_COLLECTIONS = (
    "sheet", "font", "fill", "border", "cell_xf", "cell_style_xf", "dxf", "table", "pivot",
    "pivot_field", "pivot_data_field", "pivot_filter", "pivot_field_item", "pivot_cache_field",
    "pivot_cache_record", "pivot_cache_field_shared_item", "pivot_cache_calculated_item",
    "cf_rule", "validation", "protected_range", "sparkline_group", "sparkline", "external_link",
    "cell_phonetic_run",
)  # fmt: skip
NATURAL_COLLECTIONS = {
    "comment": "cell", "threaded_comment": "id", "person": "id", "hyperlink": "cell",
    "merge": "rect", "defined_name": "name", "data_table": "cell", "lambda_text": "cell",
    "pivot_cache": "cache_id", "num_fmt": "num_fmt_id", "function": "name",
    "locale_error_name": "error code", "row_break": "row", "col_break": "col",
    "row_override": "row", "theme_script_font": "script", "shape": "object_id",
    "image": "object_id", "chart": "object_id", "slicer": "object_id",
    "drawing_object": "object_id", "passthrough": "part name", "cell": "cell",
}  # fmt: skip
# The element getter of a natural collection whose noun is not the collection's own.
NATURAL_ELEMENT_NOUNS = {"function_name": "function", "pivot_cache_id": "pivot_cache"}
# Element getters of index collections named by a property of the element.
ELEMENT_NOUNS = {**NATURAL_ELEMENT_NOUNS, "sheet_name": "sheet"}
# Result handles: elements are enumerated with count / at only.
HANDLE_SCOPES = ("cell_range", "cf_results", "cell_nodes", "pivot_cells", "pagination", "chart", "chart_data")
# Element nouns of handles whose count / at name none.
HANDLE_ELEMENT_NOUN = {"cell_range": "cell", "cell_nodes": "node", "pivot_cells": "cell"}
# Compound nouns whose last word is also a vocabulary verb.
_COMPOUND_NOUNS = frozenset({"sparkline_group", "date_group", "number_format", "num_fmt"})
# Index collections numbered across the workbook rather than per sheet.
_WORKBOOK_INDEXED = frozenset({"table"})
IRREGULAR_PLURALS = {"person": "people", "series": "series", "data": "data"}


def plural(noun: str) -> Tuple[str, bool]:
    """English plural of the last word of `noun`, and whether the plural is regular."""
    words = noun.split("_")
    last = words[-1]
    if last in IRREGULAR_PLURALS:
        return "_".join(words[:-1] + [IRREGULAR_PLURALS[last]]), False
    if re.search(r"(s|x|z|ch|sh)$", last):
        last += "es"
    elif re.search(r"[^aeiou]y$", last):
        last = last[:-1] + "ies"
    else:
        last += "s"
    return "_".join(words[:-1] + [last]), True


def _replace_seq(words: Sequence[str], table) -> List[str]:
    out = list(words)
    for old, new in table.items():
        i = 0
        while i <= len(out) - len(old):
            if tuple(out[i : i + len(old)]) == old:
                out[i : i + len(old)] = list(new)
                i += len(new)
            else:
                i += 1
    return out


def collection_kind(noun: str) -> Optional[str]:
    if noun in INDEX_COLLECTIONS:
        return "index"
    if noun in NATURAL_COLLECTIONS:
        return "natural"
    return None


@dataclass
class Proposal:
    c_old: str
    c_new: str
    scope: str  # an `api_conventions.SCOPES` member, or "process"
    verb: str
    noun: str  # words after the verb, without the `_at` / `_count` suffix
    suffix: str  # "", "at" or "count"
    factory: bool = False
    flags: List[str] = field(default_factory=list)

    @property
    def collection(self) -> Optional[str]:
        noun = NATURAL_ELEMENT_NOUNS.get(self.noun, self.noun)
        return noun if collection_kind(noun) else None


def _index_params(fn: CFunction) -> List[str]:
    return [
        p.name
        for p in fn.params
        if re.search(r"(^|_)(index|idx)$", p.name) and p.name != "sheet_index" and not p.name.startswith("out")
    ]


def _verb_position(words: Sequence[str]) -> Optional[int]:
    for k, word in enumerate(words):
        if word not in C_VERBS:
            continue
        if k and words[k - 1] in _PREPOSITIONS:
            continue
        if any("_".join(words[j : k + 1]) in _COMPOUND_NOUNS for j in range(k)):
            continue
        return k
    return None


def propose_c(fn: CFunction, typedefs: Dict[str, CTypedef]) -> Proposal:
    """New C name of one function under R1 (form, scope), R2 (verbs), R3 (suffixes) and R4 (synonyms)."""
    flags: List[str] = []
    name = fn.name
    first = fn.params[0] if fn.params else None
    if any(q.base_type == "fm_workbook_t" and q.pointer_depth == 2 for q in fn.params):
        words = name[len("fm_workbook_") :].split("_")
        return Proposal(name, name, "workbook", words[0], "_".join(words[1:]), "", True)
    if name in PROCESS_RENAMES or first is None or first.base_type not in typedefs:
        target = PROCESS_RENAMES.get(name, name)
        scope, words = split_c_name(target)
        if scope is None:
            suffix = words[-1] if words[-1] in ("at", "count") else ""
            core = words[1:-1] if suffix else words[1:]
            return Proposal(name, target, "process", words[0], "_".join(core), suffix)
        return Proposal(name, target, scope, words[0], "_".join(words[1:]), "")
    factory = False
    old_scope, words = split_c_name(name)
    if old_scope is None:
        words = name[3:].split("_")
    if first.base_type == "fm_workbook_t":
        if words[0] == "styles":
            words = words[1:]
        scope = "sheet" if len(fn.params) > 1 and fn.params[1].name in ("sheet", "sheet_index") else "workbook"
        if name == "fm_workbook_paginate":
            scope = "workbook"
    else:
        scope = first.base_type[3:-2]
    for old, new in PHRASE_RENAMES.items():
        if tuple(words[: len(old)]) == old:
            words = list(new) + words[len(old) :]
            break
    if words[0] in VERB_REWRITES:
        words = [VERB_REWRITES[words[0]]] + words[1:]
    if words[0] == "read":
        flags.append(f"`read_` prefix: get_{'_'.join(words[1:])} (read_* -> get_*) or get_{'_'.join(words)}")
        words = ["get"] + words[1:]
    words = _replace_seq(words, NOUN_SYNONYMS)
    position = _verb_position(words)
    if position is None:
        unlisted = [w for w in words if w in UNLISTED_VERBS]
        if unlisted:
            flags.append(f"`{unlisted[0]}` acts as a verb outside R2; proposed a get_ prefix, or add a verb")
        verb, noun_words = "get", list(words)
    else:
        verb, noun_words = words[position], words[:position] + words[position + 1 :]
    if verb == "match" and noun_words[-1:] == ["at"]:
        flags.append("`match_at` reads `match` as the verb: element getter get_match_at, or keep match_at")
    if verb == "create" and not factory:
        flags.append(f"`create` outside the workbook factories: add_{'_'.join(noun_words)} (R2 add) or keep create")
    if noun_words[:1] == ["cf"] and noun_words[1:2] != ["rule"] and noun_words[1:2] in ([], ["at"], ["count"]):
        noun_words = ["cf", "rule"] + noun_words[1:]
        flags.append("CF rule collection noun: cf_rule (as add_cf_rule, fm_cf_rule_t) or cf")
    suffix = ""
    if noun_words[-2:] == ["at", "index"]:
        noun_words = noun_words[:-2] + ["at"]
    if noun_words and noun_words[-1] in ("at", "count"):
        suffix = noun_words.pop()
    if scope in HANDLE_ELEMENT_NOUN and suffix and not noun_words:
        noun_words = [HANDLE_ELEMENT_NOUN[scope]]
        flags.append(f"{scope} element noun is absent: proposed `{noun_words[0]}`")
    noun = "_".join(noun_words)
    kind = collection_kind(NATURAL_ELEMENT_NOUNS.get(noun, noun))
    index_params = _index_params(fn)
    coords = {p.name for p in fn.params} >= {"row", "col"}
    if scope not in HANDLE_SCOPES and kind is not None:
        if suffix == "at" and verb in ("get", "remove") and kind == "index":
            suffix = ""  # an index-keyed getter takes `<thing>_index` without a suffix
        elif suffix == "at" and verb == "get" and kind == "natural" and coords and not index_params:
            suffix = ""  # `_at` named the natural key, not a position
        elif not suffix and verb == "get" and noun in NATURAL_COLLECTIONS and index_params and not coords:
            suffix = "at"  # positional read of a natural-key collection
    if scope == "sheet" and noun in _WORKBOOK_INDEXED:
        flags.append(f"`{noun}` is indexed workbook-wide but this takes sheet_index: workbook or sheet scope (D5)")
    if suffix == "at" and verb == "add":
        flags.append("`_at` is an insert position, not an enumeration (R3): keep add_..._at or rename")
    if verb == "clear" and kind is not None and not suffix:
        many, _ = plural(noun)
        flags.append(f"clear of the whole collection: clear_{many} (as clear_hyperlinks) or clear_{noun}")
    new_words = [verb] + noun_words + ([suffix] if suffix else [])
    prefix = "fm_" if scope == "process" else f"fm_{scope}_"
    return Proposal(name, prefix + "_".join(new_words), scope, verb, noun, suffix, factory, flags)


# ---------------------------------------------------------------------------
# Binding linkage.
# ---------------------------------------------------------------------------

_CAMEL_RE = re.compile(r"[A-Z]?[a-z0-9]+|[A-Z]+(?![a-z])")
_SCOPE_FILLER = frozenset({"workbook", "styles"})
# Abbreviations a binding name spells out.
_WORD_ALIASES = {"cf": ("conditional", "format")}


def words_of(name: str) -> List[str]:
    """Lower-case words of a camelCase, PascalCase or snake_case name."""
    name = name.split(".")[-1]
    if "_" in name or name.islower():
        return [w for w in name.lower().split("_") if w]
    return [w.lower() for w in _CAMEL_RE.findall(name)]


def _stem(word: str) -> str:
    return word[:-1] if len(word) > 3 and word.endswith("s") and not word.endswith("ss") else word


def _c_words(c_name: str) -> List[str]:
    out: List[str] = []
    for word in c_name[3:].split("_"):
        if word not in _SCOPE_FILLER:
            out.extend(_WORD_ALIASES.get(word, (word,)))
    return out


def similarity(binding: str, c_name: str) -> float:
    """How well a binding name spells a C function name, in [0, 1].

    The C name is read with and without its scope word, since bindings drop it.
    """
    a = [x for w in words_of(binding) for x in _WORD_ALIASES.get(w, (w,))]
    best = 0.0
    full = _c_words(c_name)
    scope, _ = split_c_name(c_name)
    variants = [full]
    if scope and full[:1] == [scope]:
        variants.append(full[1:])
    for b in variants:
        if "".join(a) == "".join(b):
            return 1.0
        if [_stem(w) for w in a] == [_stem(w) for w in b]:
            best = max(best, 0.95)
            continue
        sa, sb = {_stem(w) for w in a}, {_stem(w) for w in b}
        if sa | sb:
            best = max(best, min(0.9, len(sa & sb) / len(sa | sb)))
    return best


@dataclass
class Link:
    surface: str
    key: str  # the exported symbol (`Workbook.<name>` for statics)
    member: Member
    c: Optional[str] = None  # primary C function
    pair: Optional[Tuple[str, str]] = None  # (count, element getter) a list method walks
    score: float = 0.0


def collection_pairs(props: Dict[str, Proposal], s: Surfaces) -> Dict[str, Tuple[str, str]]:
    """C function -> the (count, element getter) pair of its collection, for both members."""
    counts: Dict[Tuple[str, str], str] = {}
    getters: Dict[Tuple[str, str], str] = {}
    for p in props.values():
        noun = NATURAL_ELEMENT_NOUNS.get(p.noun, p.noun)
        if p.suffix == "count":
            counts[(p.scope, noun)] = p.c_old
        elif p.verb == "get" and (p.suffix == "at" or _index_params(s.functions[p.c_old])):
            getters.setdefault((p.scope, noun), p.c_old)
    out: Dict[str, Tuple[str, str]] = {}
    for key, count in counts.items():
        if key in getters:
            out[count] = out[getters[key]] = (count, getters[key])
    return out


def _best(member: Member, cands: Iterable[str], props: Dict[str, Proposal]) -> Tuple[Optional[str], float]:
    scored = [
        (max(similarity(member.name, c), similarity(member.name, props[c].c_new)), -len(c), c) for c in sorted(cands)
    ]
    if not scored:
        return None, 0.0
    score, _, c = max(scored)
    return c, score


def link_members(s: Surfaces, props: Dict[str, Proposal]) -> List[Link]:
    """Ties each binding method, static and free function to the C function it wraps."""
    pair_of = collection_pairs(props, s)
    links: List[Link] = []
    surfaces = [(surface, s.js[surface]) for surface in JS_SURFACES] + [("py", s.py)]
    for surface, members in surfaces:
        for key, member in members.items():
            if member.kind not in ("method", "static", "free", "property"):
                continue
            link = Link(surface, key, member)
            cands = {c for c in member.calls if c in s.functions and c not in _PLUMBING}
            if not any(w in words_of(member.name) for w in ("dispose", "delete", "close", "destroy")):
                cands = {c for c in cands if not c.endswith("_destroy")} or cands
            # A workbook method opens a result handle; reads through the handle are not its op.
            owned = {c for c in cands if props[c].scope not in HANDLE_SCOPES}
            if member.kind != "free" and owned:
                cands = owned
            exact = [c for c in s.functions if similarity(member.name, c) == 1.0]
            best, score = _best(member, cands, props)
            if exact and score < 1.0:
                best, score = exact[0], 1.0
            if best is not None:
                link.c, link.score = best, score
                pair = pair_of.get(best)
                plural_name = _stem(words_of(member.name)[-1]) != words_of(member.name)[-1]
                if pair and pair[1] in cands and (pair[0] in cands or plural_name) and score < 1.0:
                    link.pair, link.c = pair, None
            links.append(link)
    # A JS method a helper table dispatches is read off the other JS package.
    by_name = {(lk.surface, lk.key): lk for lk in links}
    for lk in links:
        if lk.c is None and lk.pair is None and lk.surface in JS_SURFACES:
            other = by_name.get(("native" if lk.surface == "wasm" else "wasm", lk.key))
            if other is not None and (other.c or other.pair):
                lk.c, lk.pair, lk.score = other.c, other.pair, other.score
    return links


# ---------------------------------------------------------------------------
# Ledger assembly.
# ---------------------------------------------------------------------------

# Binding-only symbols whose ledger shape the rules settle.
BINDING_ONLY = {
    "evalFormula": ("formula.eval", "free", "evalFormula", "eval_formula",
                    "binding convenience: evaluates one formula in a default workbook"),
    "eval_formula": ("formula.eval", "free", "evalFormula", "eval_formula",
                     "binding convenience: evaluates one formula in a default workbook"),
    "createFormulon": ("module.create", "free", "createFormulon", None,
                       "entry point of both npm packages (default and named export)"),
}  # fmt: skip
# Second binding names of one C function the rules retire.
SPEC_DROPS = {
    "getCommentResult": "R8.js",  # the null-returning lookup and its envelope twin become one OptionalResult
    "version": "R0.coverage",  # alias of versionString; both become getVersion
    "version_string": "R0.coverage",  # alias of library_version; both become get_version
}
# Second binding forms of one C function the rules name.
SPEC_FORMS = {
    "addHyperlinkRange": ("hyperlink.add_range", "add_hyperlink_range", "binding form taking a CellRect (R7)"),
    "add_hyperlink_range": ("hyperlink.add_range", "add_hyperlink_range", "binding form taking a CellRect (R7)"),
}
# R12: native gains the C functions WASM already exposes.
NATIVE_ADDITIONS = frozenset(
    {"fm_styles_add_cell_style_xf", "fm_sheet_get_auto_filter_xml", "fm_sheet_set_auto_filter_xml"}
)
# Nouns both scopes carry; R4 qualifies each with its scope.
QUALIFIED_NOUNS = frozenset({"protection"})
# Proposals that would otherwise collide or chain: (C name, flag reason, entry key).
_LAYOUT_REASON = (
    "fm_workbook_pivot_layout (cells) and pivot_get/set_layout (report layout) both derive get_pivot_layout: "
    "proposed *_pivot_report_layout for the enum pair (as PivotReportLayout), or rename the cells reader"
)
_COMMENT_REASON = (
    "get_comment_at is the current name of the coordinate getter (now get_comment): the positional getter "
    "stays get_comment_at_index here; reaching get_comment_at needs a third name"
)
DISAMBIGUATE = {
    "fm_workbook_pivot_get_layout": ("fm_sheet_get_pivot_report_layout", _LAYOUT_REASON, "pivot_report_layout.get"),
    "fm_workbook_pivot_set_layout": ("fm_sheet_set_pivot_report_layout", _LAYOUT_REASON, "pivot_report_layout.set"),
    "fm_workbook_pivot_layout": ("fm_sheet_get_pivot_layout", _LAYOUT_REASON, "pivot_layout.get"),
    "fm_sheet_get_comment_at_index": ("fm_sheet_get_comment_at_index", _COMMENT_REASON, "comment.get_at"),
}
# Natural collections whose binding list the rules name although no binding walks them yet.
SPEC_LISTS = {"function_name": ("wasm", "native", "py"), "locale_error_name": ("wasm", "native", "py")}
LIFETIME_EXCEPTION = "R10 lifetime: JS dispose() and [Symbol.dispose](); Python close()"
NATIVE_ONLY_EXCEPTION = "process memory of the native addon; WASM linear memory has no equivalent"
# Records replaced by a single shape instead of renamed.
REPLACED_BY = {
    "fm_merge_range": ("record.cell_rect", "R11.rect"),
    "fm_print_range_t": ("record.cell_rect", "R11.rect"),
    "fm_cf_cell_range_t": ("record.cell_rect", "R11.rect"),
    "fm_chart_sheet_range_t": ("record.sheet_rect", "R11.sheet"),
}
# Shapes and ops the rules introduce.
NEW_ENTRIES = {
    "record.cell_rect": {"kind": "record", "role": "result", "c": "fm_cell_rect_t", "js": "CellRect", "py": "CellRect",
                         "surfaces": list(LEDGER_SURFACES)},
    "record.sheet_rect": {"kind": "record", "role": "result", "c": "fm_sheet_rect_t", "js": "SheetRect", "py": "SheetRect",
                          "surfaces": list(LEDGER_SURFACES)},
    "defined_name.get": {"kind": "op", "scope": "workbook", "noun": "defined_name", "c": "fm_workbook_get_defined_name",
                         "js": "getDefinedName", "py": "get_defined_name", "surfaces": list(LEDGER_SURFACES)},
    "defined_name.remove": {"kind": "op", "scope": "workbook", "noun": "defined_name",
                            "c": "fm_workbook_remove_defined_name", "js": "removeDefinedName",
                            "py": "remove_defined_name", "surfaces": list(LEDGER_SURFACES)},
}  # fmt: skip
_ENVELOPE_RE = re.compile(r"(Result|Results)$")


def _type_key(c_typedef: str) -> str:
    body = c_typedef[3:] if c_typedef.startswith("fm_") else c_typedef
    return body[:-2] if body.endswith("_t") else body


def _snake(name: str) -> str:
    return "_".join(words_of(name))


def _new_typedef(name: str) -> str:
    return name if name.endswith("_t") else name + "_t"


class Builder:
    def __init__(self, s: Surfaces) -> None:
        self.s = s
        self.props = {f.name: propose_c(f, s.typedefs) for f in s.functions.values()}
        self.links = link_members(s, self.props)
        self.entries: Dict[str, dict] = {}
        self.enum_key_of: Dict[Tuple[str, str], str] = {}  # (field, binding enum name) -> enum key
        self.claimed: Dict[Tuple[str, str], str] = {}  # (surface, symbol) -> entry key
        self.op_key_of: Dict[str, str] = {}  # current C function -> entry key

    # -- helpers ---------------------------------------------------------------

    def add(self, key: str, entry: dict) -> str:
        base, n = key, 2
        while key in self.entries:
            key = f"{base}.{n}"
            n += 1
        self.entries[key] = entry
        return key

    def flag(self, entry: dict, reason: str) -> None:
        if reason:
            entry["flag"] = f"{entry['flag']}; {reason}" if entry.get("flag") else reason

    def claim(self, key: str, surface: str, symbol: str) -> None:
        self.claimed[(surface, symbol)] = key

    # -- types -----------------------------------------------------------------

    def build_types(self) -> None:
        s = self.s
        js_types = {}
        for surface in JS_SURFACES:
            for name, m in s.js[surface].items():
                if m.kind in ("type", "enum"):
                    js_types.setdefault(name, set()).add(surface)
        py_types = {n for n, m in s.py.items() if m.kind in ("type", "enum")}
        roles = self._struct_roles()
        votes = self._type_votes()
        linked_js: Dict[str, str] = {}
        linked_py: Dict[str, str] = {}
        plan = []
        for name, typedef in s.typedefs.items():
            new_c = _new_typedef(name)
            derived = derive_type(new_c)
            current = derive_type(name if name.endswith("_t") else name + "_t")
            js = derived if derived in js_types else current if current in js_types else None
            py = derived if derived in py_types else current if current in py_types else None
            if typedef.kind == "scalar":  # a scalar code links to a binding enum, never to a record
                js = js if js and any(s.js[x].get(js, Member("", "")).kind == "enum" for x in JS_SURFACES) else None
                py = py if py and s.py[py].kind == "enum" else None
            if py is None and typedef.kind != "scalar":
                cited = [k for k, v in {**s.py_enums, **s.py_types}.items() if v == name]
                py = cited[0] if cited else None
            plan.append([name, typedef, new_c, derived, js, py])
            if js:
                linked_js[js] = name
            if py:
                linked_py[py] = name
        # A binding name Python already links carries over to the JS package.
        for row in plan:
            if row[4] is None and row[5] in js_types and row[5] not in linked_js:
                row[4] = row[5]
                linked_js[row[5]] = row[0]
        # Uncertain links: the binding type a C struct's ops use most, named alike; enums by name alone.
        for row in plan:
            if row[1].kind == "enum":
                for field_name, pool, linked, idx in (("js", js_types, linked_js, 4), ("py", py_types, linked_py, 5)):
                    if row[idx] is None:
                        enums = {n for n in pool if n not in linked and self._is_enum(field_name, n)}
                        guess = self._guess_type(row[0], row[3], Counter({n: 1 for n in enums}), pool, linked)
                        if guess:
                            row[idx] = guess
                            linked[guess] = row[0]
                continue
        for row in plan:
            if row[1].kind == "enum":
                continue
            name, typedef, new_c, derived, js, py = row
            for field_name, pool, linked, idx in (("js", js_types, linked_js, 4), ("py", py_types, linked_py, 5)):
                if row[idx] is not None:
                    continue
                guess = self._guess_type(name, derived, votes.get((field_name, name), Counter()), pool, linked)
                if guess:
                    row[idx] = guess
                    linked[guess] = name
                    row.append((field_name, guess))
        for row in plan:
            name, typedef, new_c, derived, js, py = row[:6]
            guesses = row[6:]
            kind = {"struct": "record", "enum": "enum"}.get(typedef.kind, "typedef")
            entry: dict = {"kind": kind}
            if kind == "record":
                entry["role"] = roles.get(name, "result")
            surfaces = ["c"]
            if js:
                surfaces += sorted(js_types[js], key=JS_SURFACES.index)
            if py:
                surfaces.append("py")
            was: dict = {"c": name}
            if js:
                was["js"] = js
            if py:
                was["py"] = py
            if name in REPLACED_BY:
                target, closes = REPLACED_BY[name]
                entry.update({"replaced_by": target, "closes": closes, "surfaces": surfaces, "was": was})
            else:
                entry["c"] = new_c
                binding_name = derived
                if kind == "record" and entry["role"] == "input" and not derived.endswith("Input"):
                    self.flag(entry, f"input-only record: {derived}Input (R7) or {derived} (R5 derive_type)")
                if js:
                    entry["js"] = binding_name
                if py:
                    entry["py"] = binding_name
                entry.update({"surfaces": surfaces, "was": was})
            for field_name, guess in guesses:
                other = py if field_name == "js" else js
                other_pool = py_types if field_name == "js" else js_types
                if guess == derived + "Entry" and other == derived and guess not in other_pool:
                    continue  # the `<Name>Entry` spelling of the record the other binding names `<Name>`
                self.flag(
                    entry, f"{field_name} {guess} linked by usage, not by name: rename to {derived} or keep apart"
                )
            if kind == "enum":
                if js and js != derived:
                    self.flag(entry, f"bindings spell {js}: rename to {derived} (R5) or exception")
                elif py and py != derived:
                    self.flag(entry, f"Python spells {py}: rename to {derived} (R5) or exception")
            key = self.add(f"{kind if kind != 'typedef' else 'type'}.{_type_key(new_c)}", entry)
            if kind == "enum":
                for field_name, binding in (("js", js), ("py", py)):
                    if binding:
                        self.enum_key_of[(field_name, binding)] = _type_key(new_c)
                self.enum_key_of[("c", name)] = _type_key(new_c)
            self.claim(key, "c", name)
            if js:
                for surface in js_types[js]:
                    self.claim(key, surface, js)
            if py:
                self.claim(key, "py", py)
        for key, entry in NEW_ENTRIES.items():
            if key.startswith("record."):
                self.entries[key] = dict(entry)
        # Binding-only types.
        for name, surfaces in sorted(js_types.items()):
            if name in linked_js:
                continue
            py = name if name in py_types and name not in linked_py else None
            kind = "enum" if any(s.js[x][name].kind == "enum" for x in surfaces) else "record"
            entry = {"kind": kind, "js": name}
            if py:
                entry["py"] = py
            entry["surfaces"] = sorted(surfaces, key=JS_SURFACES.index) + (["py"] if py else [])
            entry["was"] = {"js": name, **({"py": py} if py else {})}
            if kind == "record" and _ENVELOPE_RE.search(name):
                entry["role"] = "result"
            key = self.add(f"{kind}.{_snake(name)}", entry)
            if kind == "enum":
                self.enum_key_of[("js", name)] = _snake(name)
                if py:
                    self.enum_key_of[("py", py)] = _snake(name)
            for surface in surfaces:
                self.claim(key, surface, name)
            if py:
                linked_py[py] = key
                self.claim(key, "py", py)
        for name in sorted(py_types - set(linked_py)):
            kind = "enum" if s.py[name].kind == "enum" else "record"
            key = self.add(
                f"{kind}.{_snake(name)}", {"kind": kind, "py": name, "surfaces": ["py"], "was": {"py": name}}
            )
            if kind == "enum":
                self.enum_key_of[("py", name)] = _snake(name)
            self.claim(key, "py", name)

    def _is_enum(self, field_name: str, name: str) -> bool:
        if field_name == "py":
            return self.s.py[name].kind == "enum"
        return any(self.s.js[x].get(name, Member("", "")).kind == "enum" for x in JS_SURFACES)

    def _struct_roles(self) -> Dict[str, str]:
        s = self.s
        structs = {n for n, t in s.typedefs.items() if t.kind == "struct"}
        role: Dict[str, str] = {}
        for fn in s.functions.values():
            for p in fn.params:
                if p.base_type not in structs:
                    continue
                if p.pointer_depth == 0 or p.ctype.startswith("const"):
                    role.setdefault(p.base_type, "input")
                else:
                    role[p.base_type] = "result"
        for name in structs:
            if name.endswith(("_patch", "_patch_t")):
                role[name] = "patch"
        changed = True
        while changed:  # an embedded record takes its container's role
            changed = False
            for name in structs:
                for member in s.typedefs[name].fields:
                    inner = member.base_type
                    if inner in structs and role.get(name) == "result" and role.get(inner) != "result":
                        role[inner] = "result"
                        changed = True
                    elif inner in structs and inner not in role and name in role:
                        role[inner] = role[name] if role[name] != "patch" else "input"
                        changed = True
        return role

    def _type_votes(self) -> Dict[Tuple[str, str], Counter]:
        s = self.s
        structs = {n for n, t in s.typedefs.items() if t.kind == "struct"}
        votes: Dict[Tuple[str, str], Counter] = defaultdict(Counter)
        for link in self.links:
            fns = [link.c] if link.c else list(link.pair or ())
            cs = {p.base_type for fn in fns for p in s.functions[fn].params if p.base_type in structs}
            pool = s.py if link.surface == "py" else s.js[link.surface]
            names = set(re.findall(r"\b([A-Z][A-Za-z0-9]+)\b", link.member.detail))
            names = {n for n in names if n in pool and pool[n].kind == "type" and not _ENVELOPE_RE.search(n)}
            field_name = "py" if link.surface == "py" else "js"
            for c in cs:
                for n in names:
                    votes[(field_name, c)][n] += 1.0 / len(names)
        return votes

    @staticmethod
    def _guess_type(c_name, derived, votes: Counter, pool, linked) -> Optional[str]:
        for name, _ in votes.most_common():
            if name in linked or name not in pool:
                continue
            a = {_stem(w) for w in words_of(derived)}
            b = {_stem(w) for w in words_of(name)} - {"entry", "info", "record", "spec", "input"}
            if a and b and len(a & b) / len(a | b) >= 0.5:
                return name
        return None

    # -- ops -------------------------------------------------------------------

    def _op_key(self, p: Proposal) -> str:
        if p.factory:
            return f"workbook.{'_'.join(x for x in (p.verb, p.noun) if x)}"
        if p.scope in HANDLE_SCOPES or p.scope == "buffer":
            return f"{p.scope}.{'_'.join(x for x in (p.verb, p.noun, p.suffix) if x)}"
        noun = p.collection or p.noun or p.scope
        rest = p.noun[len(noun) :].strip("_") if p.collection and p.noun.startswith(noun) else ""
        verb = "_".join(x for x in (p.verb, rest, p.suffix) if x)
        return f"{noun}.{verb}"

    def _typing(self, fn: CFunction, links: Sequence[Link], p: Proposal) -> Tuple[Optional[str], Dict[str, str]]:
        """`returns` and `params` for boolean- and enum-typed slots."""
        params: Dict[str, str] = {}
        returns: Optional[str] = None
        c_names = {q.name for q in fn.params}
        for q in fn.params:
            key = self.enum_key_of.get(("c", q.base_type))
            if key is None:
                continue
            if q.name.startswith("out") and q.pointer_depth == 1:
                returns = f"enum:{key}"
            elif q.pointer_depth == 0:
                params[q.name] = f"enum:{key}"
        if p.verb in ("is", "has"):
            returns = "bool"
        for link in links:
            field_name = "py" if link.surface == "py" else "js"
            args, ret = _split_signature(link.member.detail, field_name)
            for arg, annotation in args:
                c_arg = arg if arg in c_names else _snake(arg) if _snake(arg) in c_names else None
                slot = c_arg or _snake(arg)
                kind = self._slot_kind(annotation, field_name)
                if kind and slot not in params:
                    params[slot] = kind
            kind = self._slot_kind(ret, field_name)
            if kind and returns is None:
                returns = kind
        return returns, params

    def _slot_kind(self, annotation: str, field_name: str) -> Optional[str]:
        annotation = re.sub(r"^Optional\[(.*)\]$", r"\1", annotation.strip())
        annotation = re.sub(r"\s*\|\s*(None|null|undefined)\b", "", annotation).strip()
        if annotation in ("bool", "boolean", "BoolResult"):
            return "bool"
        for name in re.findall(r"\b([A-Z][A-Za-z0-9]+)\b", annotation):
            key = self.enum_key_of.get((field_name, name))
            if key is not None and not re.search(r"\b(List|Sequence|Iterable|ReadonlyArray|ListResult)\b", annotation):
                return f"enum:{key}"
        return None

    def build_ops(self) -> None:
        s = self.s
        by_c: Dict[str, List[Link]] = defaultdict(list)
        for link in self.links:
            if link.c:
                by_c[link.c].append(link)
        for c_old, p in self.props.items():
            fn = s.functions[c_old]
            links = sorted(by_c.get(c_old, []), key=lambda lk: -lk.score)
            primary: Dict[str, Link] = {}
            extras: List[Link] = []
            for link in links:
                if link.surface in primary or link.member.name in SPEC_DROPS or link.member.name in SPEC_FORMS:
                    extras.append(link)
                elif link.member.name in BINDING_ONLY:
                    extras.append(link)
                else:
                    primary[link.surface] = link
            entry: dict = {"kind": "factory" if p.factory else "op"}
            if entry["kind"] == "op":
                entry["scope"] = p.scope
                entry["noun"] = p.noun or p.scope
            c_new = p.c_new
            if p.noun in QUALIFIED_NOUNS:
                c_new = c_new.replace(f"_{p.noun}", f"_{p.scope}_{p.noun}", 1)
                entry["noun"] = f"{p.scope}_{p.noun}"
            if c_old in DISAMBIGUATE:
                c_new, reason, _ = DISAMBIGUATE[c_old]
                self.flag(entry, reason)
            if c_old == "fm_workbook_create":
                c_new = "fm_workbook_create_default"
                self.flag(
                    entry,
                    "Workbook.createDefault (R5) needs fm_workbook_create_default; or keep fm_workbook_create with an exception",
                )
            entry["c"] = c_new
            for reason in p.flags:
                self.flag(entry, reason)
            surfaces = ["c"] + [x for x in ("wasm", "native", "py") if x in primary]
            if c_old in NATIVE_ADDITIONS and "native" not in surfaces:
                surfaces.insert(surfaces.index("wasm") + 1 if "wasm" in surfaces else 1, "native")
            has_js = any(x in surfaces for x in JS_SURFACES)
            js_name = derive_js(c_new)
            py_name = derive_py(c_new)
            if entry["kind"] == "factory":
                js_name, py_name = f"Workbook.{js_name}", f"Workbook.{py_name}"
            if c_old == "fm_workbook_destroy":
                js_name, py_name = "dispose", "close"
                entry["exception"] = LIFETIME_EXCEPTION
                if "wasm" not in surfaces:
                    surfaces.insert(1, "wasm")
            if "native" in surfaces and "wasm" not in surfaces and c_old not in NATIVE_ADDITIONS:
                entry["exception"] = NATIVE_ONLY_EXCEPTION
            if has_js:
                entry["js"] = js_name
            if "py" in surfaces:
                entry["py"] = py_name
            returns, params = self._typing(fn, list(primary.values()), p)
            if returns:
                entry["returns"] = returns
            if params:
                entry["params"] = params
            if returns == "bool" and p.verb == "get" and not p.flags:
                setter = c_new.replace("_get_", "_set_", 1)
                if not any(q.c_new == setter for q in self.props.values()):
                    self.flag(entry, f"boolean read with no setter: is_{p.noun} (R2 predicate) or get_{p.noun}")
            entry["surfaces"] = surfaces
            was = {"c": c_old}
            js_links = [primary[x] for x in JS_SURFACES if x in primary]
            if js_links:
                was["js"] = js_links[0].member.name if js_links[0].member.kind != "static" else js_links[0].key
                if len({lk.key for lk in js_links}) > 1:
                    extras.append(js_links[1])
                    del primary[js_links[1].surface]
            if "py" in primary:
                was["py"] = primary["py"].key
            entry["was"] = was
            if c_old in DISAMBIGUATE:
                key = DISAMBIGUATE[c_old][2]
            elif p.noun in QUALIFIED_NOUNS:
                key = f"{p.scope}_{p.noun}.{p.verb}"
            else:
                key = self._op_key(p)
            key = self.add(key, entry)
            self.claim(key, "c", c_old)
            for surface, link in primary.items():
                self.claim(key, surface, link.key)
            for link in extras:
                self._extra(link, key, entry)
            self.op_key_of[c_old] = key
        for key, entry in NEW_ENTRIES.items():
            if not key.startswith("record."):
                self.entries[key] = dict(entry)

    def _extra(self, link: Link, op_key: str, op: dict) -> None:
        """A second binding name reaching one C function."""
        name = link.member.name
        field_name = "py" if link.surface == "py" else "js"
        surface_key = (link.surface, link.key)
        if surface_key in self.claimed:
            return
        if name in SPEC_FORMS:
            key, py, reason = SPEC_FORMS[name]
            js = name if field_name == "js" else "addHyperlinkRange"
            entry = self.entries.get(key)
            if entry is None:
                entry = {
                    "kind": "op",
                    "scope": op.get("scope"),
                    "noun": op.get("noun"),
                    "js": js,
                    "py": py,
                    "surfaces": [],
                    "exception": reason,
                    "was": {},
                }
                self.entries[key] = entry
            entry["was"][field_name] = name
            entry["surfaces"] = sorted(set(entry["surfaces"]) | {link.surface}, key=LEDGER_SURFACES.index)
        elif name in BINDING_ONLY:
            key, kind, js, py, reason = BINDING_ONLY[name]
            entry = self.entries.setdefault(
                key, {"kind": kind, "js": js, "surfaces": [], "exception": reason, "was": {}}
            )
            if py:
                entry["py"] = py
            entry["was"][field_name] = name
            entry["surfaces"] = sorted(set(entry["surfaces"]) | {link.surface}, key=LEDGER_SURFACES.index)
        else:
            key = f"{op_key}.{_snake(name)}"
            entry = self.entries.get(key)
            if entry is None or entry.get("was", {}).get(field_name, link.key) != link.key:
                entry = {"kind": "op", "surfaces": [], "was": {}}
                closes = SPEC_DROPS.get(name)
                if closes:
                    entry["drop"] = {"closes": closes}
                else:
                    self.flag(
                        entry,
                        f"second binding name for {op.get('c')} (entry {op_key}): alias to drop, or a distinct op",
                    )
                key = self.add(key, entry)
            entry["was"][field_name] = link.key
            if "drop" in entry:
                entry["drop"] = {**{k: v for k, v in entry["drop"].items() if k != "closes"}, field_name: link.key,
                                 "closes": entry["drop"]["closes"]}  # fmt: skip
            entry["surfaces"] = sorted(set(entry["surfaces"]) | {link.surface}, key=LEDGER_SURFACES.index)
        self.claim(key, link.surface, link.key)

    # -- lists -----------------------------------------------------------------

    def build_lists(self) -> None:
        """`list<Plural>` for natural-key collections (R5); count / at leave the bindings."""
        groups: Dict[Tuple[str, str], List[Link]] = defaultdict(list)
        for link in self.links:
            if link.pair:
                groups[link.pair].append(link)
        pairs = collection_pairs(self.props, self.s)
        for pair in sorted(set(pairs.values())):
            count, at = pair
            noun = NATURAL_ELEMENT_NOUNS.get(self.props[at].noun, self.props[at].noun)
            kind = collection_kind(noun)
            links = groups.get(pair, [])
            if kind == "natural" and noun != "cell" and self.props[at].scope not in HANDLE_SCOPES:
                self._natural_list(pair, links)
            elif links:
                # One entry per walker: the first JS and Python names share it, any further name gets its own.
                firsts: Dict[str, str] = {}
                for lk in sorted(links, key=lambda lk: (lk.surface, lk.key)):
                    firsts.setdefault("py" if lk.surface == "py" else "js", lk.key)
                shared = [lk for lk in links if firsts["py" if lk.surface == "py" else "js"] == lk.key]
                self._flagged_list(pair, noun, kind, shared)
                for name in sorted({lk.key for lk in links if lk not in shared}):
                    self._flagged_list(pair, noun, kind, [lk for lk in links if lk.key == name])

    def _flagged_list(self, pair: Tuple[str, str], noun: str, kind: Optional[str], links: List[Link]) -> None:
        """A binding walker over a collection R5 gives no `list<Plural>`; decided by hand."""
        count, at = pair
        link = links[0]
        entry = {
            "kind": "list",
            "from": [self.op_key_of[count], self.op_key_of[at]],
            "surfaces": sorted({lk.surface for lk in links}, key=LEDGER_SURFACES.index),
            "was": {("py" if lk.surface == "py" else "js"): lk.key for lk in links},
        }
        if kind == "index":
            reason = (
                f"index-key collection `{noun}`: R5 exposes count + getter; drop this list or keep it as an exception"
            )
        elif noun == "cell":
            reason = (
                "cells keep getCellCount / getCellAt and the cursor (R5): drop this walker or keep it as an exception"
            )
        else:
            reason = f"collection `{noun}` is not in the R3 table: classify index or natural, then list or drop"
        self.flag(entry, reason)
        key = self.add(f"{noun}.list.{_snake(link.member.name)}", entry)
        for lk in links:
            self.claim(key, lk.surface, lk.key)

    def _natural_list(self, pair: Tuple[str, str], links: List[Link]) -> None:
        count, at = pair
        element = self.props[at].noun
        many, regular = plural(element)
        surfaces = set()
        for c in pair:
            entry = self.entries[self.op_key_of[c]]
            surfaces |= set(entry["surfaces"]) - {"c"}
        surfaces |= {lk.surface for lk in links}
        if element in SPEC_LISTS:
            surfaces |= set(SPEC_LISTS[element])
        if not surfaces:
            return  # a C-only collection has no binding list
        entry: dict = {"kind": "list", "from": [self.op_key_of[count], self.op_key_of[at]]}
        js = "list" + "".join(w[:1].upper() + w[1:] for w in many.split("_"))
        if surfaces & set(JS_SURFACES):
            entry["js"] = js
        if "py" in surfaces:
            entry["py"] = f"list_{many}"
        entry["surfaces"] = sorted(surfaces, key=LEDGER_SURFACES.index)
        if not regular:
            self.flag(entry, f"plural of `{element}`: {many} or another spelling")
        was: dict = {}
        extras: List[Link] = []
        for link in sorted(links, key=lambda lk: (lk.surface != "wasm", lk.surface)):
            field_name = "py" if link.surface == "py" else "js"
            if field_name not in was:
                was[field_name] = link.key
            elif was[field_name] != link.key:
                extras.append(link)
        if was:
            entry["was"] = was
        key = self.add(f"{many}.list", entry)
        for link in links:
            if link not in extras:
                self.claim(key, link.surface, link.key)
        for link in extras:
            alias = {"kind": "list", "surfaces": [link.surface], "was": {"js": link.key}}
            self.flag(alias, f"second list name for {key}: alias to drop, or a distinct op")
            self.claim(self.add(f"{key}.{_snake(link.member.name)}", alias), link.surface, link.key)
        for c in pair:
            op = self.entries[self.op_key_of[c]]
            for name_field in ("js", "py"):
                op.pop(name_field, None)
            was_names = {f: op["was"][f] for f in ("js", "py") if f in op["was"]}
            if was_names:
                op["drop"] = {**was_names, "closes": "R5.list"}
            op["surfaces"] = ["c"]

    # -- leftovers -------------------------------------------------------------

    def build_unlinked(self) -> None:
        """Binding symbols no C function or type claimed."""
        s = self.s
        for surface in LEDGER_SURFACES[1:]:
            pool = s.py if surface == "py" else s.js[surface]
            field_name = "py" if surface == "py" else "js"
            for symbol, member in pool.items():
                if (surface, symbol) in self.claimed:
                    continue
                existing = next(
                    (k for (sf, sym), k in self.claimed.items() if sym == symbol and sf != surface
                     and (sf == "py") == (surface == "py")), None)  # fmt: skip
                if existing is not None:
                    entry = self.entries[existing]
                    entry["surfaces"] = sorted(set(entry["surfaces"]) | {surface}, key=LEDGER_SURFACES.index)
                    self.claim(existing, surface, symbol)
                    continue
                name = member.name
                if name in BINDING_ONLY:
                    key, kind, js, py, reason = BINDING_ONLY[name]
                    entry = self.entries.setdefault(
                        key, {"kind": kind, "js": js, "surfaces": [], "exception": reason, "was": {}}
                    )
                    if py:
                        entry["py"] = py
                    entry["was"][field_name] = symbol
                    entry["surfaces"] = sorted(set(entry["surfaces"]) | {surface}, key=LEDGER_SURFACES.index)
                    self.claim(key, surface, symbol)
                    continue
                if member.kind == "const":
                    key = self._constant_for(field_name, symbol, surface)
                    self.claim(key, surface, symbol)
                    continue
                if name == "delete" and surface == "wasm":
                    entry = {"kind": "op", "surfaces": ["wasm"], "was": {"js": "delete"},
                             "drop": {"js": "delete", "closes": "R10.js"}}  # fmt: skip
                    self.claim(self.add("workbook.delete", entry), surface, symbol)
                    continue
                key = f"binding.{_snake(name)}"
                if key in self.entries and field_name not in self.entries[key]["was"]:
                    entry = self.entries[key]
                    entry[field_name] = name
                    entry["was"][field_name] = symbol
                    entry["surfaces"] = sorted(set(entry["surfaces"]) | {surface}, key=LEDGER_SURFACES.index)
                    self.claim(key, surface, symbol)
                    continue
                entry = {"kind": "free" if member.kind == "free" else "op", field_name: name,
                         "surfaces": [surface], "was": {field_name: symbol}}  # fmt: skip
                if name in ("isValid", "is_valid"):
                    reason = "binding-only handle liveness check: keep as an exception or drop"
                elif name in ("mergeFunctionMetadata", "merge_function_metadata"):
                    reason = "binding-only helper; verb `merge` is outside R2: keep as an exception or rename"
                else:
                    reason = "no C counterpart found: binding-only (exception) or link to a C function"
                self.flag(entry, reason)
                self.claim(self.add(f"binding.{_snake(name)}", entry), surface, symbol)

    def _constant_for(self, field_name: str, symbol: str, surface: str) -> str:
        c_name = f"FM_{symbol}"
        key = self.claimed.get(("c", c_name))
        if key is not None:
            entry = self.entries[key]
            entry[field_name] = symbol
            entry.setdefault("was", {})[field_name] = symbol
            entry["surfaces"] = sorted(set(entry["surfaces"]) | {surface}, key=LEDGER_SURFACES.index)
            return key
        for (sf, sym), k in self.claimed.items():
            if sym == symbol and ((sf == "py") == (field_name == "py")):
                entry = self.entries[k]
                entry["surfaces"] = sorted(set(entry["surfaces"]) | {surface}, key=LEDGER_SURFACES.index)
                return k
        return self.add(f"constant.{symbol.lower()}", {"kind": "constant", field_name: symbol, "surfaces": [surface],
                                                      "was": {field_name: symbol}})  # fmt: skip

    def build_constants(self) -> None:
        for name, enum in self.s.constants.items():
            entry: dict = {"kind": "constant", "c": name}
            if enum:
                entry["enum"] = f"enum.{_type_key(_new_typedef(enum))}"
            entry.update({"surfaces": ["c"], "was": {"c": name}})
            self.claim(self.add(f"constant.{name.lower()}", entry), "c", name)
        for name in self.s.status:
            entry = {"kind": "constant", "c": name, "enum": "enum.status_code", "surfaces": ["c"], "was": {"c": name}}
            self.claim(self.add(f"constant.{name.lower()}", entry), "c", name)

    # -- collections and cross-entry checks --------------------------------------

    def build_collections(self) -> Dict[str, dict]:
        collections: Dict[str, dict] = {}
        members: Dict[str, List[str]] = defaultdict(list)
        for key, entry in self.entries.items():
            if entry.get("kind") != "op" or "c" not in entry or entry.get("scope") in HANDLE_SCOPES:
                continue
            noun = ELEMENT_NOUNS.get(entry.get("noun"), entry.get("noun"))
            if collection_kind(noun):
                members[noun].append(key)
        for noun in INDEX_COLLECTIONS:
            collections[noun] = {"key": "index", "ops": sorted(members.get(noun, []))}
        for noun, natural_key in NATURAL_COLLECTIONS.items():
            entry = {"key": "natural", "by": natural_key}
            if noun != "cell":
                entry["plural"] = plural(noun)[0]
            entry["ops"] = sorted(members.get(noun, []))
            collections[noun] = entry
        # Collections the R3 table does not list, found by their count function.
        for p in self.props.values():
            if p.suffix != "count" or p.scope in HANDLE_SCOPES or collection_kind(p.noun) or p.scope == "process":
                continue
            ops = sorted(k for k, e in self.entries.items() if e.get("noun") == p.noun and e.get("kind") == "op")
            getter = next(
                (q for q in self.props.values() if q.noun == p.noun and q.verb == "get" and q.suffix != "count"), None
            )
            by_index = getter is not None and _index_params(self.s.functions[getter.c_old])
            collections[p.noun] = {
                "key": "index" if by_index else "natural",
                "ops": ops,
                "flag": f"`{p.noun}` is not in the R3 table: proposed {'index' if by_index else 'natural'} "
                "from its getter's parameters; index or natural",
            }
        return collections

    def check_names(self) -> None:
        """Flags new C / binding names that collide or chain into another entry's current name."""
        for name_field in ("c", "js", "py"):
            new: Dict[str, List[str]] = defaultdict(list)
            old: Dict[str, str] = {}
            for key, entry in self.entries.items():
                if entry.get("kind") == "field":
                    continue
                if isinstance(entry.get(name_field), str):
                    new[entry[name_field]].append(key)
                was = entry.get("was", {}).get(name_field)
                if isinstance(was, str) and was != entry.get(name_field):
                    old[was] = key
            for name, keys in new.items():
                if len(keys) > 1:
                    for key in keys:
                        others = ", ".join(k for k in keys if k != key)
                        self.flag(self.entries[key], f"{name_field} name {name} collides with {others}: rename one")
                if name in old and old[name] not in keys:
                    for key in keys:
                        self.flag(
                            self.entries[key],
                            f"{name_field} name {name} is the current name of {old[name]} (rename chain): "
                            "route through a third name or rename the other",
                        )

    def check_nouns(self) -> None:
        """Flags a noun two scopes share without a qualifier (R4), once per noun."""
        scopes: Dict[str, Dict[str, List[str]]] = defaultdict(lambda: defaultdict(list))
        for key, entry in self.entries.items():
            if entry.get("kind") == "op" and entry.get("scope") and "c" in entry:
                scopes[entry["noun"]][entry["scope"]].append(key)
        for noun, by_scope in sorted(scopes.items()):
            if len(by_scope) < 2:
                continue
            keys = [k for scope in sorted(by_scope) for k in sorted(by_scope[scope])]
            spread = ", ".join(f"{scope} ({len(ks)})" for scope, ks in sorted(by_scope.items()))
            self.flag(
                self.entries[keys[0]],
                f"noun `{noun}` spans scopes {spread} (R4): qualify per scope, mark `qualifies:`, or rename",
            )

    def build(self) -> dict:
        self.build_types()
        self.build_ops()
        self.build_lists()
        self.build_constants()
        self.build_unlinked()
        self.entries["field.status_code"] = {
            "kind": "field", "record": "record.status", "js": "code", "surfaces": ["wasm", "native"],
            "was": {"js": "status"}, "closes": "R8.js",
        }  # fmt: skip
        collections = self.build_collections()
        self.check_names()
        self.check_nouns()
        return {"schema": 1, "collections": collections, "entries": self.entries}


def _split_signature(detail: str, field_name: str) -> Tuple[List[Tuple[str, str]], str]:
    """`(name, annotation)` pairs and the return annotation of a d.ts or Python signature."""
    if not detail:
        return [], ""
    open_at = detail.find("(")
    if open_at < 0:
        return [], ""
    depth = 0
    close_at = -1
    for i in range(open_at, len(detail)):
        if detail[i] in "([{<":
            depth += 1
        elif detail[i] in ")]}>":
            depth -= 1
            if depth == 0:
                close_at = i
                break
    inner = detail[open_at + 1 : close_at]
    ret = detail[close_at + 1 :].lstrip(" :").replace("->", "", 1).strip()
    args: List[Tuple[str, str]] = []
    depth = 0
    current = ""
    for char in inner + ",":
        if char in "([{<":
            depth += 1
        elif char in ")]}>":
            depth -= 1
        if char == "," and depth == 0:
            if ":" in current:
                name, annotation = current.split(":", 1)
                args.append((name.strip().rstrip("?"), annotation.strip()))
            current = ""
        else:
            current += char
    return args, ret


# ---------------------------------------------------------------------------
# YAML emission.
# ---------------------------------------------------------------------------

_FIELD_ORDER = (
    "kind", "scope", "noun", "record", "role", "from", "c", "js", "py", "enum", "returns", "params",
    "surfaces", "exception", "replaced_by", "closes", "was", "drop", "flag",
)  # fmt: skip
_PLAIN_RE = re.compile(r"[A-Za-z_][A-Za-z0-9_.]*")
_YAML_WORDS = frozenset({"y", "n", "yes", "no", "on", "off", "true", "false", "null", "none", "~"})


def _scalar(value) -> str:
    if isinstance(value, bool):
        return "true" if value else "false"
    text = str(value)
    if _PLAIN_RE.fullmatch(text) and text.lower() not in _YAML_WORDS:
        return text
    return '"' + text.replace("\\", "\\\\").replace('"', '\\"') + '"'


def _flow(value) -> str:
    if isinstance(value, dict):
        return "{ " + ", ".join(f"{_scalar(k)}: {_flow(v)}" for k, v in value.items()) + " }"
    if isinstance(value, (list, tuple)):
        return "[" + ", ".join(_flow(v) for v in value) + "]"
    return _scalar(value)


def dump_ledger(ledger: dict) -> str:
    lines = [
        "# Public API name ledger: one entry per exported symbol, keyed by a stable op id.",
        "# `was` holds the current names, the top-level names the target ones. `scope: process`",
        "# marks a process-wide function (JS named export, Python module function). `drop` names",
        "# leave a surface when the `closes` rule closes; `replaced_by` folds a type into another.",
        "# `flag` marks an entry the rules do not settle; it is resolved by hand.",
        f"schema: {ledger['schema']}",
        "collections:",
    ]
    for name, entry in ledger["collections"].items():
        lines.append(f"  {_scalar(name)}: {_flow(entry)}")
    lines.append("entries:")
    for key, entry in ledger["entries"].items():
        lines.append(f"  {_scalar(key)}:")
        for field_name in sorted(entry, key=lambda f: _FIELD_ORDER.index(f) if f in _FIELD_ORDER else 99):
            lines.append(f"    {field_name}: {_flow(entry[field_name])}")
    return "\n".join(lines) + "\n"


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("-o", "--output", type=Path, help="write the proposed ledger here instead of stdout")
    args = parser.parse_args(argv)
    text = dump_ledger(Builder(scan()).build())
    if args.output:
        args.output.write_text(text, encoding="utf-8")
    else:
        sys.stdout.write(text)
    return 0


if __name__ == "__main__":
    sys.exit(main())
