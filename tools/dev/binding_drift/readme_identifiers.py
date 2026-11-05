"""Identifier and link checks for the public prose: READMEs, CONTRIBUTING and the doc comments of both bindings.

Every name a README code block or doc comment points at must exist in the
exported surface of the language it is written in, and no relative link in a
public document may lead to a git-ignored path. The surfaces are read the way
the sibling checks read them (header regex, `.d.ts` text, Python `ast`); no
package is imported.
"""

from __future__ import annotations

import ast
import re
import subprocess
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Set, Tuple

from .c_header import c_abi_declarations
from .surface_files import CAPI_HEADER, NODE_DTS, PYTHON_PKG_DIR, REPO_ROOT, WASM_DTS, _read
from .ts_decl import _extract_braced_block

PUBLIC_DOCS: Tuple[Path, ...] = (
    REPO_ROOT / "CONTRIBUTING.md",
    REPO_ROOT / "README.md",
    REPO_ROOT / "README_ja.md",
    *sorted((REPO_ROOT / "packages").glob("*/README.md")),
)

_FENCE_RE = re.compile(r"^```(\w*)[^\n]*\n(.*?)^```", re.S | re.M)
_C_NAME_RE = re.compile(r"(?<![\w])fm_[a-z0-9_]+")
_INLINE_CODE_RE = re.compile(r"`([^`\n]+)`")
_LINK_RE = re.compile(r"\]\(([^)\s]+)\)")
_JS_IMPORT_RE = re.compile(r"import\s+(?:(\w+)\s*,?\s*)?(?:\{([^}]*)\}\s*)?from\s*'(@libraz/formulon[^']*)'")
_TS_MEMBER_RE = re.compile(r"^ {2}(?:readonly\s+)?([A-Za-z_$][A-Za-z0-9_$]*)\??\s*[(:<]", re.M)
_TS_EXPORT_RE = re.compile(
    r"^export\s+(?:declare\s+)?(?:interface|type|enum|class|function|const|namespace)\s+([A-Za-z_$][A-Za-z0-9_$]*)",
    re.M,
)
_JSDOC_LINK_RE = re.compile(r"\{@link\s+([A-Za-z_$][\w$]*)(?:\.([A-Za-z_$][\w$]*))?[^}]*\}")
_PY_ROLE_RE = re.compile(r":(?:meth|func|class):`~?([^`]+)`")
_NATIVE_PACKAGE = "@libraz/formulon-native"


def _members(text: str, interface: str, source: Path) -> Set[str]:
    match = re.search(r"^export interface %s\b[^{]*\{" % re.escape(interface), text, re.M)
    if match is None:
        raise RuntimeError(f"interface {interface!r} not found in {source}")
    return set(_TS_MEMBER_RE.findall(_extract_braced_block(text, match.end())))


class _JsSurface:
    """What one JS package exports: top-level names, module members, `Workbook` members, factory members."""

    def __init__(self, dts: Path, native: bool) -> None:
        text = _read(dts)
        self.exports = set(_TS_EXPORT_RE.findall(text))
        self.workbook = _members(text, "Workbook", dts) | {"dispose"}
        self.ctor = _members(text, "WorkbookCtor", dts)
        self.module = self.exports - {"default"} if native else _members(text, "FormulonModule", dts)
        self.interfaces = {
            name: _members(text, name, dts)
            for name in self.exports
            if re.search(r"^export interface %s\b" % re.escape(name), text, re.M)
        }


def _js_surfaces() -> Dict[str, _JsSurface]:
    wasm = _JsSurface(WASM_DTS, native=False)
    return {"@libraz/formulon": wasm, _NATIVE_PACKAGE: _JsSurface(NODE_DTS, native=True)}


def _surface_for(package: str, surfaces: Dict[str, _JsSurface]) -> _JsSurface:
    return surfaces[_NATIVE_PACKAGE] if package == _NATIVE_PACKAGE else surfaces["@libraz/formulon"]


class _PySurface:
    """The Python package's `__all__`, class names and the public members of each class."""

    def __init__(self) -> None:
        self.all: Set[str] = set()
        self.classes: Dict[str, Set[str]] = {}
        self.functions: Set[str] = set()
        self.docstrings: List[Tuple[str, str]] = []
        for path in sorted(PYTHON_PKG_DIR.glob("*.py")):
            tree = ast.parse(_read(path), filename=str(path))
            public = not path.name.startswith("_") or path.name == "__init__.py"
            for node in tree.body:
                if isinstance(node, ast.Assign) and any(getattr(t, "id", "") == "__all__" for t in node.targets):
                    self.all |= {e.value for e in node.value.elts if isinstance(e, ast.Constant)}
                if isinstance(node, ast.ClassDef):
                    members = {
                        item.name
                        for item in node.body
                        if isinstance(item, (ast.FunctionDef, ast.AsyncFunctionDef)) and not item.name.startswith("_")
                    }
                    members |= {
                        t.id
                        for item in node.body
                        if isinstance(item, (ast.Assign, ast.AnnAssign))
                        for t in ([item.target] if isinstance(item, ast.AnnAssign) else item.targets)
                        if isinstance(t, ast.Name)
                    }
                    self.classes.setdefault(node.name, set()).update(members)
                elif isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                    self.functions.add(node.name)
            if public:
                for node in ast.walk(tree):
                    if isinstance(node, (ast.Module, ast.ClassDef, ast.FunctionDef, ast.AsyncFunctionDef)):
                        doc = ast.get_docstring(node, clean=False)
                        if doc:
                            self.docstrings.append((f"{path.relative_to(REPO_ROOT)}:{getattr(node, 'lineno', 1)}", doc))
        self.module = self.all | self.functions | set(self.classes)


def _code_blocks(text: str) -> Iterable[Tuple[str, str, int]]:
    for m in _FENCE_RE.finditer(text):
        yield m.group(1).lower(), m.group(2), text.count("\n", 0, m.start()) + 1


def _check_c_names(label: str, text: str, c_names: Set[str], problems: List[str]) -> None:
    spans = [body for _, body, _ in _code_blocks(text)]
    stripped = _FENCE_RE.sub("", text)
    spans += _INLINE_CODE_RE.findall(stripped)
    for span in spans:
        for name in _C_NAME_RE.findall(span):
            if name not in c_names:
                problems.append(f"readme-identifiers: {label}: `{name}` is not declared in formulon_c.h")


def _check_js_block(label: str, line: int, block: str, surfaces: Dict[str, _JsSurface], problems: List[str]) -> None:
    where = f"{label}:{line}"
    package: Optional[str] = None
    imported: Set[str] = set()
    factory: Optional[str] = None
    for m in _JS_IMPORT_RE.finditer(block):
        package = m.group(3).split("/threads")[0]
        surface = _surface_for(package, surfaces)
        if m.group(1):
            factory = m.group(1)
        for raw in (m.group(2) or "").split(","):
            name = raw.split(" as ")[0].strip()
            if not name:
                continue
            imported.add(name)
            if name not in surface.exports:
                problems.append(f"readme-identifiers: {where}: `{name}` is not exported by {package}")
    surface = _surface_for(package or "@libraz/formulon", surfaces)
    modules: Set[str] = set()
    if factory:
        modules |= {
            m.group(1) for m in re.finditer(r"(?:const|let)\s+(\w+)\s*=\s*await\s+%s\(" % re.escape(factory), block)
        }
    ctors = {f"{mod}.Workbook" for mod in modules} | ({"Workbook"} if "Workbook" in imported else set())
    ctor_re = "|".join(re.escape(c) for c in sorted(ctors, key=len, reverse=True))
    results: Set[str] = set()
    workbooks: Set[str] = set()
    if ctors:
        for m in re.finditer(r"(?:const|let)\s+(\w+)\s*=\s*(?:%s)\.\w+\([^)]*\)(\.value)?" % ctor_re, block):
            (workbooks if m.group(2) else results).add(m.group(1))
        for m in re.finditer(r"(?:const|let)\s+(\w+)\s*=\s*(\w+)\.value\b", block):
            if m.group(2) in results:
                workbooks.add(m.group(1))
    if ctors:
        for m in re.finditer(r"(?<![\w.])(?:%s)\.(\w+)\(" % ctor_re, block):
            if m.group(1) not in surface.ctor:
                problems.append(f"readme-identifiers: {where}: `{m.group(0)[:-1]}` is not a member of WorkbookCtor")
    for mod in modules:
        for m in re.finditer(r"(?<![\w.])%s\.(\w+)\(" % re.escape(mod), block):
            if m.group(1) not in surface.module:
                problems.append(f"readme-identifiers: {where}: `{mod}.{m.group(1)}` is not a member of the module")
    for wb in workbooks:
        for m in re.finditer(r"(?<![\w.])%s\.(\w+)\(" % re.escape(wb), block):
            if m.group(1) not in surface.workbook:
                problems.append(f"readme-identifiers: {where}: `{wb}.{m.group(1)}` is not a Workbook method")


def _check_py_block(label: str, line: int, block: str, py: _PySurface, problems: List[str]) -> None:
    where = f"{label}:{line}"
    imported: Set[str] = set()
    for m in re.finditer(r"^\s*from\s+formulon\s+import\s+([^\n(]+)", block, re.M):
        for raw in m.group(1).split(","):
            name = raw.split(" as ")[0].strip()
            if name:
                imported.add(name)
                if name not in py.module:
                    problems.append(f"readme-identifiers: {where}: `{name}` is not exported by formulon")
    has_module = bool(re.search(r"^\s*import\s+formulon\b", block, re.M))
    ctors: List[str] = (["formulon.Workbook"] if has_module else []) + (["Workbook"] if "Workbook" in imported else [])
    ctor_re = "|".join(re.escape(c) for c in ctors)
    if has_module:
        for m in re.finditer(r"(?<![\w.])formulon\.(\w+)", block):
            if m.group(1) not in py.module:
                problems.append(f"readme-identifiers: {where}: `formulon.{m.group(1)}` is not exported by formulon")
    if not ctors:
        return
    workbook = py.classes.get("Workbook", set())
    for m in re.finditer(r"(?<![\w.])(?:%s)\.(\w+)" % ctor_re, block):
        if m.group(1) not in workbook:
            problems.append(f"readme-identifiers: {where}: `Workbook.{m.group(1)}` is not a Workbook member")
    names = {m.group(1) for m in re.finditer(r"(?:with\s+(?:%s)\.\w+\([^)]*\)\s+as\s+(\w+))" % ctor_re, block)}
    names |= {m.group(1) for m in re.finditer(r"^\s*(\w+)\s*=\s*(?:%s)\.\w+\(" % ctor_re, block, re.M)}
    for wb in names:
        for m in re.finditer(r"(?<![\w.])%s\.(\w+)" % re.escape(wb), block):
            if m.group(1) not in workbook:
                problems.append(f"readme-identifiers: {where}: `{wb}.{m.group(1)}` is not a Workbook member")


def _check_links(path: Path, text: str, problems: List[str]) -> None:
    label = path.relative_to(REPO_ROOT)
    targets: List[Path] = []
    for raw in _LINK_RE.findall(_FENCE_RE.sub("", text)):
        if "://" in raw or raw.startswith(("#", "mailto:")):
            continue
        targets.append((path.parent / raw.split("#", 1)[0]).resolve())
    for target in targets:
        try:
            rel = target.relative_to(REPO_ROOT)
        except ValueError:
            continue
        ignored = subprocess.run(["git", "check-ignore", "-q", str(rel)], cwd=REPO_ROOT).returncode == 0
        if ignored:
            problems.append(f"readme-identifiers: {label}: links to git-ignored path `{rel}`")


def _check_js_doc_links(surfaces: Dict[str, _JsSurface], problems: List[str]) -> None:
    for path, surface in ((WASM_DTS, surfaces["@libraz/formulon"]), (NODE_DTS, surfaces[_NATIVE_PACKAGE])):
        label = path.relative_to(REPO_ROOT)
        for m in _JSDOC_LINK_RE.finditer(_read(path)):
            head, member = m.group(1), m.group(2)
            if head not in surface.exports:
                problems.append(f"readme-identifiers: {label}: {{@link {head}}} is not exported")
            elif member and member not in surface.interfaces.get(head, set()):
                problems.append(f"readme-identifiers: {label}: {{@link {head}.{member}}} has no such member")


def _check_py_doc_roles(py: _PySurface, problems: List[str]) -> None:
    members = set().union(*py.classes.values()) if py.classes else set()
    for where, doc in py.docstrings:
        for raw in _PY_ROLE_RE.findall(doc):
            parts = raw.split(".")
            if parts[0] != "formulon" and len(parts) > 1 and parts[0] not in py.classes:
                continue  # another package's symbol
            if parts[0] == "formulon":
                parts = parts[1:]
                if parts and parts[0] in ("workbook", "metadata"):
                    parts = parts[1:]
            if not parts:
                continue
            if len(parts) == 1:
                if parts[0] not in py.module and parts[0] not in members:
                    problems.append(f"readme-identifiers: {where}: `{raw}` does not resolve")
            elif parts[0] not in py.classes or parts[1] not in py.classes[parts[0]]:
                problems.append(f"readme-identifiers: {where}: `{raw}` does not resolve")


def check_readme_identifiers() -> List[str]:
    """Every identifier a public doc names exists on its surface, and no public link leads to an ignored path."""
    problems: List[str] = []
    c_names = set(c_abi_declarations(_read(CAPI_HEADER)))
    c_names |= set(re.findall(r"\}\s*(fm_[a-z0-9_]+)\s*;", _read(CAPI_HEADER)))
    c_names |= set(re.findall(r"\btypedef\b[^;]*?\b(fm_[a-z0-9_]+)\s*;", _read(CAPI_HEADER)))
    surfaces = _js_surfaces()
    py = _PySurface()
    for path in PUBLIC_DOCS:
        text = _read(path)
        label = str(path.relative_to(REPO_ROOT))
        _check_links(path, text, problems)
        if path.name == "CONTRIBUTING.md":
            continue
        _check_c_names(label, text, c_names, problems)
        for lang, block, line in _code_blocks(text):
            if lang in ("js", "javascript", "ts", "typescript", "mjs"):
                _check_js_block(label, line, block, surfaces, problems)
            elif lang in ("python", "py"):
                _check_py_block(label, line, block, py, problems)
    _check_js_doc_links(surfaces, problems)
    _check_py_doc_roles(py, problems)
    return problems
