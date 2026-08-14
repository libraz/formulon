#!/usr/bin/env python3
"""Binding drift guard.

Several files that describe the C-ABI / embind / N-API surface are
hand-synchronized rather than generated, and have gone out of sync more
than once: a newly added `fm_*` entry point lands in one binding but is
forgotten in another, or a `.d.ts` stops matching what is actually
registered. This script cross-checks those pairs and exits non-zero on
any mismatch, so the drift fails fast in CTest instead of shipping as a
binding-specific runtime regression.

Checks:
  python-exports  Every `LIB.fm_*` symbol called from
                  packages/python/formulon/*.py is present in
                  tools/wasm/capi_exports.txt (the staged WASM export
                  list the Python binding loads), and every symbol in
                  that list is declared in src/c_api/formulon_c.h. The
                  literal `_STATUS_RETURNING_EXPORT_NAMES` tuple in `_c.py`
                  must also equal exactly the intersection of the manifest
                  and header declarations returning `fm_status_t`.
  dts-wasm        src/wasm/formulon.d.ts (Workbook / WorkbookCtor /
                  FormulonModule method surface) matches what is
                  registered in src/wasm/parts/bindings_register.cpp.
  dts-node        packages/npm-native/index.d.ts (Workbook /
                  WorkbookCtor / free-function surface) matches what is
                  registered in src/node_addon/parts/workbook_class.cc
                  and src/node_addon/addon.cc. It also enforces the exact
                  intentional WASM-only / Node-only method allowlists.
  readme-counts   The instance-method count quoted in
                  packages/npm-native/README.md matches the actual count
                  registered in workbook_class.cc and its shared/WASM-only
                  projection is current.
  dts-enums       Every `export enum` in src/wasm/formulon.d.ts that
                  mirrors a C/C++ enum (embind passes these through as
                  plain numbers rather than registering a real
                  `enum_<T>`, so the `.d.ts` copy is the only place the
                  ordinal values live on the JS side) matches its
                  source enum's ordinal sequence. The frozen ordinal
                  tables the two published ESM entry points export
                  (packages/npm/index.mjs and
                  packages/npm-native/index.mjs) must then carry the same
                  names and the same values as each other, as that
                  canonical `.d.ts`, and as the declaration file their own
                  package ships -- swapping one JS package for the other
                  is only safe if a named constant means the same thing in
                  both.
  style-record-fields
                  Every public type that projects a style record
                  (`ColorSpec` / `FontRecord` / `FillRecord` /
                  `BorderSide`) carries the same field set as its C ABI
                  struct, across the WASM `.d.ts`, the Node `.d.ts` and
                  the Python dataclasses. Deliberate omissions live in
                  `_STYLE_RECORD_EXEMPT_TYPES`.
  staged-dist     packages/npm-native/dist/{index.d.ts,index.mjs} are
                  byte-identical to the package-root sources `stage.mjs`
                  copied them from, so the published package cannot ship a
                  declaration file or ordinal table that no longer matches
                  the source every other check here reads. Reports
                  "SKIPPED" when the package has not been staged, because
                  `dist/` is gitignored and absent in a fresh clone.
  all             Run every check above (default).

Stdlib only; no build artifacts or network access required.
"""

from __future__ import annotations

import argparse
import ast
import importlib.util
import re
import sys
from pathlib import Path
from typing import List, NamedTuple, Optional, Set

REPO_ROOT = Path(__file__).resolve().parents[2]

CAPI_HEADER = REPO_ROOT / "src" / "c_api" / "formulon_c.h"
CAPI_EXPORTS = REPO_ROOT / "tools" / "wasm" / "capi_exports.txt"
PYTHON_PKG_DIR = REPO_ROOT / "packages" / "python" / "formulon"
PYTHON_C_BINDING = PYTHON_PKG_DIR / "_c.py"

WASM_BINDINGS_CPP = REPO_ROOT / "src" / "wasm" / "parts" / "bindings_register.cpp"
WASM_DTS = REPO_ROOT / "src" / "wasm" / "formulon.d.ts"

NODE_WORKBOOK_CLASS_CC = REPO_ROOT / "src" / "node_addon" / "parts" / "workbook_class.cc"
NODE_ADDON_CC = REPO_ROOT / "src" / "node_addon" / "addon.cc"
NODE_DTS = REPO_ROOT / "packages" / "npm-native" / "index.d.ts"
NODE_README = REPO_ROOT / "packages" / "npm-native" / "README.md"
NODE_INDEX_MJS = REPO_ROOT / "packages" / "npm-native" / "index.mjs"
NODE_DIST_DIR = REPO_ROOT / "packages" / "npm-native" / "dist"
NPM_INDEX_MJS = REPO_ROOT / "packages" / "npm" / "index.mjs"

VALUE_H = REPO_ROOT / "src" / "value.h"
CF_MATCH_H = REPO_ROOT / "src" / "cf" / "cf_match.h"
CALC_MODE_H = REPO_ROOT / "src" / "io" / "calc_mode.h"
EXTERNAL_LINKS_H = REPO_ROOT / "src" / "io" / "external_links.h"
PYTHON_STRUCTS = PYTHON_PKG_DIR / "_structs.py"

# embind auto-adds a `delete()` finaliser to every `class_<T>` -- it has no
# corresponding `.function(...)` registration, so it must be excluded before
# comparing the WASM Workbook interface against bindings_register.cpp.
WASM_AUTO_METHODS = {"delete"}

# Intentional binding differences are kept exact. This makes a newly added
# method or a removed method fail drift checking instead of becoming a stale
# exception in one binding.
WASM_ONLY_METHODS = {
    "addCellStyleXf",
    "createTable",
    "getCellPhonetic",
    "getSheetAutoFilterXml",
    "removeTable",
    "setCellPhonetic",
    "setCellStyle",
    "setSheetAutoFilterXml",
    "updateTable",
}
NODE_ONLY_METHODS = {"dispose", "memoryUsage"}

# Pure-JS free functions declared in packages/npm-native/index.d.ts that are
# intentionally NOT backed by a native `exports.Set(...)` registration -- they
# are host-side helpers implemented in index.mjs (e.g. the function-metadata
# provider merge helper). Excluded before comparing the d.ts free-function
# surface against the addon's native exports.
NODE_PURE_JS_FREE_FUNCTIONS = {"mergeFunctionMetadata"}


# A sub-check appends here when an input it needs is legitimately absent, as
# opposed to wrong. `main` prints these alongside the verdict so "not checked"
# is visible in the output and can never be read as "checked and clean" -- a
# guard that quietly passes when its input is missing is not a guard.
_SKIPPED: List[str] = []


def _read(path: Path) -> str:
    if not path.is_file():
        print(f"check_binding_drift: missing file: {path}", file=sys.stderr)
        sys.exit(2)
    return path.read_text(encoding="utf-8")


def _read_bytes(path: Path) -> bytes:
    """`_read` for a comparison that must not normalise line endings."""
    if not path.is_file():
        print(f"check_binding_drift: missing file: {path}", file=sys.stderr)
        sys.exit(2)
    return path.read_bytes()


def _extract_braced_block(text: str, start: int) -> str:
    """Returns the substring between `start` (just after an opening `{`)
    and its matching closing `}`, using a simple depth counter.

    Safe for this codebase's `.d.ts` files: the only braces inside an
    interface body are either the one being matched or self-contained
    `{ ... }` pairs inside single-line JSDoc comments (e.g. `` `{ status,
    index }` ``), which do not change the net depth.
    """
    depth = 1
    i = start
    while i < len(text) and depth > 0:
        if text[i] == "{":
            depth += 1
        elif text[i] == "}":
            depth -= 1
        i += 1
    return text[start : i - 1]


def _find_interface_body(text: str, name: str, source: Path) -> str:
    match = re.search(r"export interface %s\s*\{" % re.escape(name), text)
    if not match:
        print(f"check_binding_drift: interface {name!r} not found in {source}", file=sys.stderr)
        sys.exit(2)
    return _extract_braced_block(text, match.end())


# Matches a method signature's leading `name(` at the start of a line
# (after indentation); this also catches multi-line signatures whose
# parameter list wraps, since only the opening token is needed. Plain
# data fields (`name: Type;`) do not match -- no `(` follows the name.
_TS_METHOD_RE = re.compile(r"^\s*([A-Za-z_$][A-Za-z0-9_$]*)\s*\(", re.MULTILINE)


def _extract_ts_methods(body: str) -> Set[str]:
    return set(_TS_METHOD_RE.findall(body))


def _format_diff(label_a: str, only_a: Set[str], label_b: str, only_b: Set[str]) -> List[str]:
    problems = []
    if only_a:
        problems.append(f"  in {label_a} but not {label_b}: {sorted(only_a)}")
    if only_b:
        problems.append(f"  in {label_b} but not {label_a}: {sorted(only_b)}")
    return problems


def _parse_literal_status_exports() -> tuple[Set[str], Optional[str]]:
    """Read `_STATUS_RETURNING_EXPORT_NAMES` without importing `_c.py`.

    Importing the binding would load wasmtime and, depending on the
    environment, initialize a real WASM instance. The drift check is a
    source-level consistency check, so an AST walk keeps it stdlib-only and
    side-effect-free.
    """
    try:
        tree = ast.parse(_read(PYTHON_C_BINDING), filename=str(PYTHON_C_BINDING))
    except SyntaxError as exc:
        return set(), f"python-exports: cannot parse {PYTHON_C_BINDING.relative_to(REPO_ROOT)}: {exc}"

    for statement in tree.body:
        if isinstance(statement, ast.Assign):
            targets = statement.targets
        elif isinstance(statement, ast.AnnAssign):
            targets = [statement.target]
        else:
            continue
        if not any(
            isinstance(target, ast.Name) and target.id == "_STATUS_RETURNING_EXPORT_NAMES" for target in targets
        ):
            continue
        value = statement.value
        if not isinstance(value, ast.Tuple):
            return set(), "python-exports: _STATUS_RETURNING_EXPORT_NAMES must be a literal tuple"
        names: Set[str] = set()
        for element in value.elts:
            if not isinstance(element, ast.Constant) or not isinstance(element.value, str):
                return set(), "python-exports: _STATUS_RETURNING_EXPORT_NAMES must contain only string literals"
            names.add(element.value)
        return names, None

    return (
        set(),
        "python-exports: _STATUS_RETURNING_EXPORT_NAMES literal tuple not found in packages/python/formulon/_c.py",
    )


# ---------------------------------------------------------------------------
# Check 1: Python `LIB.fm_*` calls <-> tools/wasm/capi_exports.txt <->
#          src/c_api/formulon_c.h declarations.
# ---------------------------------------------------------------------------


def check_python_exports() -> List[str]:
    problems: List[str] = []

    python_calls: Set[str] = set()
    for py_file in sorted(PYTHON_PKG_DIR.glob("*.py")):
        python_calls |= set(re.findall(r"\bLIB\.(fm_[A-Za-z0-9_]+)", _read(py_file)))

    exports_text = _read(CAPI_EXPORTS)
    exports: Set[str] = set()
    for line in exports_text.splitlines():
        stripped = line.strip()
        if stripped and not stripped.startswith("#"):
            exports.add(stripped)

    header_text = _read(CAPI_HEADER)
    header_fns = set(re.findall(r"\bfm_[A-Za-z0-9_]+(?=\s*\()", header_text))

    missing_from_exports = python_calls - exports
    if missing_from_exports:
        problems.append(
            "python-exports: Python calls a symbol not staged in "
            f"{CAPI_EXPORTS.relative_to(REPO_ROOT)}: {sorted(missing_from_exports)}"
        )

    # capi_exports.txt also lists a handful of non-`fm_` allocator symbols
    # (`malloc`/`free`) needed for host-side scratch-buffer plumbing; those
    # are libc, not part of the `fm_*` C ABI, and are never declared in
    # formulon_c.h. Only the `fm_*` subset is checked against the header.
    fm_exports = {symbol for symbol in exports if symbol.startswith("fm_")}
    exports_not_in_header = fm_exports - header_fns
    if exports_not_in_header:
        problems.append(
            "python-exports: symbol staged in "
            f"{CAPI_EXPORTS.relative_to(REPO_ROOT)} but not declared in "
            f"{CAPI_HEADER.relative_to(REPO_ROOT)}: {sorted(exports_not_in_header)}"
        )

    # The WASM result type is not enough to identify status calls: counts,
    # indices, and pointers are also i32. Keep the binding's capture list as
    # a literal and compare it to the authoritative header/manifest
    # intersection without importing `_c.py` or wasmtime.
    header_status_fns = set(re.findall(r"\bFM_API\s+fm_status_t\s+(fm_[A-Za-z0-9_]+)\s*\(", header_text))
    expected_status_exports = header_status_fns & fm_exports
    binding_status_exports, parse_problem = _parse_literal_status_exports()
    if parse_problem:
        problems.append(parse_problem)

    missing_status_exports = expected_status_exports - binding_status_exports
    if missing_status_exports:
        problems.append(
            "python-exports: _STATUS_RETURNING_EXPORT_NAMES is missing status-returning exports: "
            f"{sorted(missing_status_exports)}"
        )

    extra_status_exports = binding_status_exports - expected_status_exports
    if extra_status_exports:
        problems.append(
            f"python-exports: _STATUS_RETURNING_EXPORT_NAMES has extra exports: {sorted(extra_status_exports)}"
        )

    return problems


# ---------------------------------------------------------------------------
# Check 1b: Python WASM32 POD layouts <-> C ABI header.
# ---------------------------------------------------------------------------


def _wasm32_layout(fields: List[tuple[str, int, int]]) -> tuple[dict[str, int], int, int]:
    """Lays `fields` out in sequence, returning `(offsets, size, alignment)`."""
    offsets: dict[str, int] = {}
    offset = 0
    max_align = 1
    for name, size, align in fields:
        offset = (offset + align - 1) // align * align
        offsets[name] = offset
        offset += size
        max_align = max(max_align, align)
    return offsets, (offset + max_align - 1) // max_align * max_align, max_align


# wasm32 (ILP32) size and alignment for the leaf types the C ABI structs are
# built from. These are properties of the target ABI rather than anything the
# header states, which is why they are the only layout facts tabulated here:
# every type that has a declaration to read -- struct, union, array -- is
# measured from that declaration instead of quoted. A literal for an embedded
# struct would keep asserting its old size after the struct changed, which is
# the one thing this check exists to catch.
_WASM32_SCALARS = {
    "uint8_t": (1, 1),
    "uint16_t": (2, 2),
    "uint32_t": (4, 4),
    "int32_t": (4, 4),
    "double": (8, 8),
    "size_t": (4, 4),
}
_WASM32_POINTER = (4, 4)
# A C enum takes `int` representation. The header pins each enumerator's
# ordinal but says nothing about the storage width, so this is not derivable
# either; *which* names are enums is read off the header.
_WASM32_ENUM = (4, 4)

_C_STRUCT_RE = re.compile(r"typedef\s+struct\s*\{(.*?)\}\s*(fm_[A-Za-z0-9_]+)\s*;", re.S)
_C_ENUM_RE = re.compile(r"typedef\s+enum\s*\{[^}]*\}\s*([A-Za-z_][A-Za-z0-9_]*)\s*;")
# An anonymous `union { ... } name;` member, matched before the enclosing body
# is split on `;` so the union's own members are not mistaken for the parent's.
_C_UNION_MEMBER_RE = re.compile(r"union\s*\{(.*?)\}\s*([A-Za-z_][A-Za-z0-9_]*)\s*;", re.S)
_C_DECLARATION_RE = re.compile(r"(.+?)\s+([A-Za-z_][A-Za-z0-9_]*)(?:\[(\d+)\])?")


class _LayoutError(Exception):
    """A C declaration the wasm32 layout resolver cannot measure."""


class _Member(NamedTuple):
    """One measured struct member. `ctype` is empty for an anonymous union."""

    name: str
    ctype: str
    size: int
    align: int


class _Wasm32Layouts:
    """Measures the C ABI's structs against the wasm32 ABI, from the header.

    Struct members are resolved recursively, so an embedded POD is measured
    rather than quoted: changing a field of `fm_color_spec` moves every struct
    that embeds it, and the Python side is checked against the new offsets.
    """

    def __init__(self, header: str) -> None:
        self._structs = {name: body for body, name in _C_STRUCT_RE.findall(header)}
        self._enums = set(_C_ENUM_RE.findall(header))
        self._members: dict[str, List[_Member]] = {}
        self._pending: Set[str] = set()

    def is_struct(self, ctype: str) -> bool:
        return ctype in self._structs

    def members(self, struct_name: str) -> List[_Member]:
        """Members of `struct_name` in declaration order, each measured."""
        cached = self._members.get(struct_name)
        if cached is not None:
            return cached
        body = self._structs.get(struct_name)
        if body is None:
            raise _LayoutError(f"{struct_name} missing from C header")
        if struct_name in self._pending:
            raise _LayoutError(f"{struct_name} embeds itself")
        self._pending.add(struct_name)
        try:
            measured = self._parse_body(struct_name, body)
        finally:
            self._pending.discard(struct_name)
        self._members[struct_name] = measured
        return measured

    def extent(self, ctype: str) -> tuple[int, int]:
        """`(size, alignment)` of one member's declared type."""
        if "*" in ctype:
            return _WASM32_POINTER
        ctype = ctype.replace("const ", "").strip()
        if ctype in self._structs:
            _, size, align = _wasm32_layout([(m.name, m.size, m.align) for m in self.members(ctype)])
            return size, align
        if ctype in _WASM32_SCALARS:
            return _WASM32_SCALARS[ctype]
        if ctype in self._enums:
            return _WASM32_ENUM
        raise _LayoutError(f"unknown field type {ctype!r}")

    def _parse_body(self, owner: str, body: str) -> List[_Member]:
        measured: List[_Member] = []
        position = 0
        for union in _C_UNION_MEMBER_RE.finditer(body):
            measured.extend(self._parse_declarations(owner, body[position : union.start()]))
            size, align = self._union_extent(owner, union.group(1))
            measured.append(_Member(union.group(2), "", size, align))
            position = union.end()
        measured.extend(self._parse_declarations(owner, body[position:]))
        return measured

    def _parse_declarations(self, owner: str, text: str) -> List[_Member]:
        measured: List[_Member] = []
        for declaration in text.split(";"):
            declaration = " ".join(declaration.split())
            if not declaration:
                continue
            match = _C_DECLARATION_RE.fullmatch(declaration)
            if not match:
                raise _LayoutError(f"cannot parse {owner}: {declaration!r}")
            ctype, name, count_text = match.groups()
            size, align = self.extent(ctype)
            measured.append(_Member(name, ctype.replace("const ", "").strip(), size * int(count_text or "1"), align))
        return measured

    def _union_extent(self, owner: str, body: str) -> tuple[int, int]:
        """A union overlays its members: widest size, strictest alignment."""
        size = 0
        align = 1
        for member in self._parse_declarations(owner, body):
            size = max(size, member.size)
            align = max(align, member.align)
        return (size + align - 1) // align * align, align


def _flatten_for_python(
    layouts: _Wasm32Layouts, struct_name: str, py_field_names: Set[str]
) -> List[tuple[str, int, int]]:
    """C members of `struct_name` in the shape the Python layout models them.

    Python keeps one flat field map per struct, so an embedded POD appears in
    it one of two ways and both have to be accepted: as a single opaque blob
    under the C member's own name (`fm_pivot_cell_t.value`), or spread into
    the parent's map (`fm_cf_color_t color` contributing `color_r`). Which one
    applies is read off the Python layout, so a struct that switches
    representation needs no edit here -- only the nesting depth Python
    actually flattens is followed.
    """
    fields: List[tuple[str, int, int]] = []
    for member in layouts.members(struct_name):
        if member.name in py_field_names or not layouts.is_struct(member.ctype):
            fields.append((member.name, member.size, member.align))
            continue
        for inner in layouts.members(member.ctype):
            name = inner.name if inner.name in py_field_names else f"{member.name}_{inner.name}"
            fields.append((name, inner.size, inner.align))
    return fields


def check_python_struct_layouts() -> List[str]:
    """Verify Python's hand-written WASM32 structs against the C header.

    Every size and offset is measured from the authoritative C declarations
    rather than repeated in a Python table -- including the structs embedded
    inside other structs, which are laid out recursively. Only the wasm32
    scalar, pointer and enum widths are tabulated, because the header does not
    state those. It covers every ``Struct`` exported by ``_structs.py`` and
    reports field order, offsets, and final size.
    """
    layouts = _Wasm32Layouts(re.sub(r"/\*.*?\*/", "", _read(CAPI_HEADER), flags=re.S))
    spec = importlib.util.spec_from_file_location("formulon_struct_layouts", PYTHON_STRUCTS)
    if spec is None or spec.loader is None:
        return ["python-struct-layout: could not load packages/python/formulon/_structs.py"]
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)

    problems: List[str] = []
    for layout in (value for value in vars(module).values() if isinstance(value, module.Struct)):
        py_field_names = {name for name, _ in layout.fields}
        try:
            c_fields = _flatten_for_python(layouts, layout.name, py_field_names)
        except _LayoutError as exc:
            problems.append(f"python-struct-layout: {exc}")
            continue
        # Explicit C padding need not be represented in Struct: its alignment
        # effect is reproduced by the following semantic field. It is still
        # laid out here so the offsets after it are right.
        c_offsets, c_size, _ = _wasm32_layout(c_fields)
        c_sizes = {name: size for name, size, _ in c_fields}
        semantic_names = [name for name, _, _ in c_fields if not name.startswith("_pad")]
        py_names = [name for name, _ in layout.fields if not name.startswith("_pad")]
        py_sizes = {name: spec[1] for name, spec in layout.fields}
        if py_names != semantic_names:
            problems.append(f"python-struct-layout: {layout.name} fields differ: C={semantic_names}, Python={py_names}")
        for name in py_names:
            if name in c_offsets and layout.offsets[name][1] != c_offsets[name]:
                problems.append(
                    f"python-struct-layout: {layout.name}.{name} offset C={c_offsets[name]} Python={layout.offsets[name][1]}"
                )
            # Field width is checked separately from offset: narrowing a field
            # whose successor is more strictly aligned leaves every offset and
            # the total size intact, so the offset comparison alone reads the
            # wrong number of bytes without ever disagreeing.
            if name in c_sizes and py_sizes[name] != c_sizes[name]:
                problems.append(
                    f"python-struct-layout: {layout.name}.{name} size C={c_sizes[name]} Python={py_sizes[name]}"
                )
        if layout.size != c_size:
            problems.append(f"python-struct-layout: {layout.name} size C={c_size} Python={layout.size}")
    return problems


# ---------------------------------------------------------------------------
# Check 2a: WASM embind registration <-> src/wasm/formulon.d.ts.
# ---------------------------------------------------------------------------


def check_dts_wasm() -> List[str]:
    problems: List[str] = []
    cpp_text = _read(WASM_BINDINGS_CPP)

    instance_cpp = set(re.findall(r'\.function\("([^"]+)"', cpp_text))
    static_cpp = set(re.findall(r'\.class_function\("([^"]+)"', cpp_text))
    free_cpp = set(re.findall(r'(?<!\.)\bfunction\("([^"]+)"', cpp_text))

    dts_text = _read(WASM_DTS)
    instance_dts = _extract_ts_methods(_find_interface_body(dts_text, "Workbook", WASM_DTS)) - WASM_AUTO_METHODS
    static_dts = _extract_ts_methods(_find_interface_body(dts_text, "WorkbookCtor", WASM_DTS))
    free_dts = _extract_ts_methods(_find_interface_body(dts_text, "FormulonModule", WASM_DTS))

    for label, cpp_set, dts_set in (
        ("Workbook instance methods", instance_cpp, instance_dts),
        ("Workbook static factories", static_cpp, static_dts),
        ("free functions", free_cpp, free_dts),
    ):
        diff = _format_diff(
            f"{WASM_BINDINGS_CPP.relative_to(REPO_ROOT)}",
            cpp_set - dts_set,
            f"{WASM_DTS.relative_to(REPO_ROOT)}",
            dts_set - cpp_set,
        )
        if diff:
            problems.append(f"dts-wasm: {label} mismatch:\n" + "\n".join(diff))

    return problems


# ---------------------------------------------------------------------------
# Check 2b: Node N-API registration <-> packages/npm-native/index.d.ts.
# ---------------------------------------------------------------------------


def check_dts_node() -> List[str]:
    problems: List[str] = []
    cc_text = _read(NODE_WORKBOOK_CLASS_CC)

    instance_cc = set(re.findall(r'InstanceMethod<[^>]+>\("([^"]+)"\)', cc_text))
    static_cc = set(re.findall(r'StaticMethod<[^>]+>\("([^"]+)"\)', cc_text))

    addon_text = _read(NODE_ADDON_CC)
    free_cc = set(re.findall(r'exports\.Set\("([^"]+)"', addon_text)) - {"Workbook"}

    dts_text = _read(NODE_DTS)
    instance_dts = _extract_ts_methods(_find_interface_body(dts_text, "Workbook", NODE_DTS))
    static_dts = _extract_ts_methods(_find_interface_body(dts_text, "WorkbookCtor", NODE_DTS))
    free_dts = set(re.findall(r"^export function ([A-Za-z0-9_]+)\(", dts_text, re.MULTILINE))
    free_dts -= NODE_PURE_JS_FREE_FUNCTIONS

    wasm_cpp_text = _read(WASM_BINDINGS_CPP)
    wasm_instance = set(re.findall(r'\.function\("([^"]+)"', wasm_cpp_text))
    actual_wasm_only = wasm_instance - instance_cc
    actual_node_only = instance_cc - wasm_instance
    if actual_wasm_only != WASM_ONLY_METHODS:
        problems.append(
            "dts-node: WASM-only instance-method allowlist mismatch; "
            f"actual={sorted(actual_wasm_only)}, allowlist={sorted(WASM_ONLY_METHODS)}"
        )
    if actual_node_only != NODE_ONLY_METHODS:
        problems.append(
            "dts-node: Node-only instance-method allowlist mismatch; "
            f"actual={sorted(actual_node_only)}, allowlist={sorted(NODE_ONLY_METHODS)}"
        )
    stale_wasm_only = WASM_ONLY_METHODS - actual_wasm_only
    stale_node_only = NODE_ONLY_METHODS - actual_node_only
    if stale_wasm_only:
        problems.append(f"dts-node: stale WASM-only allowlist entries: {sorted(stale_wasm_only)}")
    if stale_node_only:
        problems.append(f"dts-node: stale Node-only allowlist entries: {sorted(stale_node_only)}")

    for label, cc_set, dts_set in (
        ("Workbook instance methods", instance_cc, instance_dts),
        ("Workbook static factories", static_cc, static_dts),
        ("free functions", free_cc, free_dts),
    ):
        diff = _format_diff(
            f"{NODE_WORKBOOK_CLASS_CC.relative_to(REPO_ROOT)} / {NODE_ADDON_CC.relative_to(REPO_ROOT)}",
            cc_set - dts_set,
            f"{NODE_DTS.relative_to(REPO_ROOT)}",
            dts_set - cc_set,
        )
        if diff:
            problems.append(f"dts-node: {label} mismatch:\n" + "\n".join(diff))

    return problems


# ---------------------------------------------------------------------------
# Check 3: README instance-method count <-> actual registration count.
# ---------------------------------------------------------------------------

_NUMBER_WORDS = {
    "zero": 0,
    "one": 1,
    "two": 2,
    "three": 3,
    "four": 4,
    "five": 5,
    "six": 6,
    "seven": 7,
    "eight": 8,
    "nine": 9,
    "ten": 10,
}

_README_COUNT_RE = re.compile(r"register (\d+) instance methods plus the\s+(\w+) static factories")
_README_SHARED_COUNT_RE = re.compile(r"instance methods, (\d+) are shared with WASM; (\w+) remain WASM-only")


def check_readme_counts() -> List[str]:
    problems: List[str] = []
    cc_text = _read(NODE_WORKBOOK_CLASS_CC)
    actual_instance = len(set(re.findall(r'InstanceMethod<[^>]+>\("([^"]+)"\)', cc_text)))
    actual_static = len(set(re.findall(r'StaticMethod<[^>]+>\("([^"]+)"\)', cc_text)))

    readme_text = _read(NODE_README)
    match = _README_COUNT_RE.search(readme_text)
    if not match:
        problems.append(
            f"readme-counts: could not find the instance/static method-count sentence in "
            f"{NODE_README.relative_to(REPO_ROOT)} (pattern: {_README_COUNT_RE.pattern!r})"
        )
        return problems

    quoted_instance = int(match.group(1))
    quoted_static_word = match.group(2).lower()
    quoted_static = _NUMBER_WORDS.get(quoted_static_word)

    if quoted_instance != actual_instance:
        problems.append(
            f"readme-counts: {NODE_README.relative_to(REPO_ROOT)} claims {quoted_instance} instance methods, "
            f"actual count in {NODE_WORKBOOK_CLASS_CC.relative_to(REPO_ROOT)} is {actual_instance}"
        )
    if quoted_static is None or quoted_static != actual_static:
        problems.append(
            f"readme-counts: {NODE_README.relative_to(REPO_ROOT)} claims {match.group(2)!r} static factories, "
            f"actual count in {NODE_WORKBOOK_CLASS_CC.relative_to(REPO_ROOT)} is {actual_static}"
        )

    wasm_cpp_text = _read(WASM_BINDINGS_CPP)
    wasm_instance = set(re.findall(r'\.function\("([^"]+)"', wasm_cpp_text))
    actual_node_methods = set(re.findall(r'InstanceMethod<[^>]+>\("([^"]+)"\)', cc_text))
    actual_shared = len(wasm_instance & actual_node_methods)
    actual_wasm_only = len(wasm_instance - actual_node_methods)
    shared_match = _README_SHARED_COUNT_RE.search(readme_text)
    if not shared_match:
        problems.append(
            f"readme-counts: could not find the shared/WASM-only method-count sentence in "
            f"{NODE_README.relative_to(REPO_ROOT)}"
        )
    else:
        quoted_shared = int(shared_match.group(1))
        quoted_wasm_only = _NUMBER_WORDS.get(shared_match.group(2).lower())
        if quoted_shared != actual_shared:
            problems.append(
                f"readme-counts: {NODE_README.relative_to(REPO_ROOT)} claims {quoted_shared} shared methods, "
                f"actual shared count is {actual_shared}"
            )
        if quoted_wasm_only is None or quoted_wasm_only != actual_wasm_only:
            problems.append(
                f"readme-counts: {NODE_README.relative_to(REPO_ROOT)} claims {shared_match.group(2)!r} WASM-only methods, "
                f"actual count is {actual_wasm_only}"
            )

    return problems


# ---------------------------------------------------------------------------
# Check 5: src/wasm/formulon.d.ts `export enum` ordinals <-> their C/C++
#          source enum. embind never registers these as a real `enum_<T>`
#          (that would cost WASM size for values only ever crossing the
#          boundary as plain numbers); the `.d.ts` copy is therefore the
#          sole place the JS-visible ordinals live, and nothing enforces
#          that it still matches the source enum after a reorder / insert.
# ---------------------------------------------------------------------------

_ENUM_BODY_COMMENT_RE = re.compile(r"//[^\n]*|/\*.*?\*/", re.DOTALL)
_ENUM_MEMBER_RE = re.compile(r"([A-Za-z_][A-Za-z0-9_]*)\s*(?:=\s*(-?(?:0[xX][0-9a-fA-F]+|\d+)))?\s*$")


def _parse_enum_body(body: str) -> Optional[List[int]]:
    """Parses a brace-delimited enumerator list into its ordinal values.

    Handles both explicit `Name = N` members and implicit sequential
    values (each unset member is one more than the previous value, C/C++/TS
    enum semantics). Returns `None` if a member's initializer is not a
    literal integer (e.g. it references another constant) -- the caller
    should skip comparison in that case rather than mis-flag a drift.
    """
    body = _ENUM_BODY_COMMENT_RE.sub("", body)
    values: List[int] = []
    next_value = 0
    for part in body.split(","):
        part = part.strip()
        if not part:
            continue
        match = _ENUM_MEMBER_RE.match(part)
        if not match:
            return None
        literal = match.group(2)
        if literal is not None:
            next_value = int(literal, 0)
        values.append(next_value)
        next_value += 1
    return values


def _extract_c_typedef_enum(text: str, type_name: str) -> Optional[List[int]]:
    # `[^}]*` (rather than a non-greedy `.*?`) deliberately cannot cross a
    # `}` boundary, so a search starting at an *earlier*, unrelated
    # `typedef enum { ... }` block cannot skip past its own close brace
    # and accidentally splice in a later block that happens to close with
    # this `type_name`.
    match = re.search(r"typedef enum\s*\{([^}]*)\}\s*" + re.escape(type_name) + r"\s*;", text, re.DOTALL)
    if not match:
        return None
    return _parse_enum_body(match.group(1))


def _extract_cpp_enum_class(text: str, enum_name: str) -> Optional[List[int]]:
    match = re.search(
        r"enum class\s+" + re.escape(enum_name) + r"\s*(?::\s*[\w:]+)?\s*\{([^}]*)\}\s*;", text, re.DOTALL
    )
    if not match:
        return None
    return _parse_enum_body(match.group(1))


def _extract_ts_enum(text: str, enum_name: str) -> Optional[List[int]]:
    match = re.search(r"export (?:const )?enum\s+" + re.escape(enum_name) + r"\s*\{([^}]*)\}", text, re.DOTALL)
    if not match:
        return None
    return _parse_enum_body(match.group(1))


# Maps each `.d.ts` enum name to the (source description, extractor) pair
# that finds its mirrored C/C++ enum's ordinal sequence. `fm_*_t` names
# resolve against formulon_c.h; the remaining few are plain C++ `enum
# class` types in their own small headers (comparing against the giant
# formulon_c.h text for those would just add false-positive risk from an
# unrelated same-named enum elsewhere).
_DTS_ENUM_SOURCES = {
    "ValueKind": (CAPI_HEADER, "fm_value_kind_t", _extract_c_typedef_enum),
    "WorkbookFormat": (CAPI_HEADER, "fm_workbook_format_t", _extract_c_typedef_enum),
    "PivotCellKind": (CAPI_HEADER, "fm_pivot_cell_kind_t", _extract_c_typedef_enum),
    "PivotAxis": (CAPI_HEADER, "fm_pivot_axis_t", _extract_c_typedef_enum),
    "PivotAggregation": (CAPI_HEADER, "fm_pivot_aggregation_t", _extract_c_typedef_enum),
    "PivotShowValuesAs": (CAPI_HEADER, "fm_pivot_show_as_t", _extract_c_typedef_enum),
    "PivotFilterType": (CAPI_HEADER, "fm_pivot_filter_type_t", _extract_c_typedef_enum),
    "PivotDateGrouping": (CAPI_HEADER, "fm_pivot_date_grouping_t", _extract_c_typedef_enum),
    "PivotCalendar": (CAPI_HEADER, "fm_pivot_calendar_t", _extract_c_typedef_enum),
    "PivotReportLayout": (CAPI_HEADER, "fm_pivot_layout_t", _extract_c_typedef_enum),
    "PivotFilterValueKind": (CAPI_HEADER, "fm_pivot_filter_value_kind_t", _extract_c_typedef_enum),
    "CfMatchKind": (CF_MATCH_H, "CFMatchKind", _extract_cpp_enum_class),
    "CalcMode": (CALC_MODE_H, "CalcMode", _extract_cpp_enum_class),
    "ErrorCode": (VALUE_H, "ErrorCode", _extract_cpp_enum_class),
    "ExternalLinkKind": (EXTERNAL_LINKS_H, "Kind", _extract_cpp_enum_class),
}


def check_dts_enums() -> List[str]:
    problems: List[str] = []
    dts_text = _read(WASM_DTS)

    for ts_name, (source_path, source_enum, extractor) in sorted(_DTS_ENUM_SOURCES.items()):
        ts_values = _extract_ts_enum(dts_text, ts_name)
        if ts_values is None:
            problems.append(
                f"dts-enums: could not find/parse `export enum {ts_name}` in {WASM_DTS.relative_to(REPO_ROOT)}"
            )
            continue

        source_text = _read(source_path)
        source_values = extractor(source_text, source_enum)
        if source_values is None:
            problems.append(
                f"dts-enums: could not find/parse `{source_enum}` in {source_path.relative_to(REPO_ROOT)} "
                f"(needed to check {WASM_DTS.relative_to(REPO_ROOT)}'s {ts_name})"
            )
            continue

        if ts_values != source_values:
            problems.append(
                f"dts-enums: {ts_name} in {WASM_DTS.relative_to(REPO_ROOT)} has ordinals {ts_values}, "
                f"but its source {source_enum} in {source_path.relative_to(REPO_ROOT)} has {source_values}"
            )

    problems.extend(_check_js_constant_tables(dts_text))
    return problems


# Frozen ordinal tables are the only way a JS consumer can name a value
# that crosses the boundary as a plain number, so both published ESM entry
# points have to carry the same ones. The WASM `.d.ts` above is already
# pinned to the C/C++ enums, which makes it the single source the two
# `index.mjs` files are measured against -- comparing them only to each
# other would let a shared mistake pass.
_JS_TABLE_RE = re.compile(r"^export const (\w+) = Object\.freeze\(\{(.*?)\}\);", re.DOTALL | re.MULTILINE)
_JS_SCALAR_RE = re.compile(r"^export const (\w+) = (-?\d+);", re.MULTILINE)
_JS_MEMBER_RE = re.compile(r"(\w+)\s*:\s*(-?\d+)")


def _parse_js_constants(path: Path) -> dict:
    """Reads an `index.mjs`'s exported ordinal tables and scalar constants."""
    text = _read(path)
    out: dict = {}
    for match in _JS_TABLE_RE.finditer(text):
        out[match.group(1)] = {name: int(value) for name, value in _JS_MEMBER_RE.findall(match.group(2))}
    for match in _JS_SCALAR_RE.finditer(text):
        out[match.group(1)] = int(match.group(2))
    return out


def _parse_ts_named_enum(text: str, name: str) -> Optional[dict]:
    """Reads a `.d.ts` ordinal table in either idiom used across the two
    declaration files: a TS `export enum` or an `export const` whose type
    is a `Readonly<{...}>` literal."""
    match = re.search(r"export enum\s+" + re.escape(name) + r"\s*\{([^}]*)\}", text, re.DOTALL)
    if match is None:
        match = re.search(r"export const\s+" + re.escape(name) + r"\s*:\s*Readonly<\{([^}]*)\}>", text, re.DOTALL)
    if match is None:
        return None
    body = _ENUM_BODY_COMMENT_RE.sub("", match.group(1))
    members: dict = {}
    next_value = 0
    for part in body.replace(";", ",").split(","):
        part = part.strip()
        if not part:
            continue
        member = re.match(r"([A-Za-z_]\w*)\s*(?:[:=]\s*(-?\d+))?$", part)
        if member is None:
            return None
        if member.group(2) is not None:
            next_value = int(member.group(2))
        members[member.group(1)] = next_value
        next_value += 1
    return members


def _check_js_constant_tables(wasm_dts_text: str) -> List[str]:
    problems: List[str] = []
    npm = _parse_js_constants(NPM_INDEX_MJS)
    native = _parse_js_constants(NODE_INDEX_MJS)
    npm_label = NPM_INDEX_MJS.relative_to(REPO_ROOT)
    native_label = NODE_INDEX_MJS.relative_to(REPO_ROOT)

    only_npm = sorted(set(npm) - set(native))
    only_native = sorted(set(native) - set(npm))
    if only_npm:
        problems.append(f"dts-enums: exported by {npm_label} but not {native_label}: {only_npm}")
    if only_native:
        problems.append(f"dts-enums: exported by {native_label} but not {npm_label}: {only_native}")

    for name in sorted(set(npm) & set(native)):
        if npm[name] != native[name]:
            problems.append(f"dts-enums: {name} is {npm[name]} in {npm_label} but {native[name]} in {native_label}")

    # Each runtime table must also match the declaration file its own
    # package ships, and the tables (not the bare scalars) must match the
    # canonical WASM `.d.ts`.
    node_dts_text = _read(NODE_DTS)
    for label, constants, dts_text, dts_path in (
        (npm_label, npm, wasm_dts_text, WASM_DTS),
        (native_label, native, node_dts_text, NODE_DTS),
    ):
        for name, value in sorted(constants.items()):
            if not isinstance(value, dict):
                if not re.search(r"export const\s+" + re.escape(name) + r"\s*=\s*" + str(value) + r"\b", dts_text):
                    problems.append(
                        f"dts-enums: {label} exports {name} = {value}, "
                        f"not declared with that value in {dts_path.relative_to(REPO_ROOT)}"
                    )
                continue
            declared = _parse_ts_named_enum(dts_text, name)
            if declared is None:
                problems.append(
                    f"dts-enums: {label} exports the table {name}, "
                    f"which {dts_path.relative_to(REPO_ROOT)} does not declare as a value"
                )
            elif declared != value:
                problems.append(
                    f"dts-enums: {name} is {value} in {label} but {declared} in {dts_path.relative_to(REPO_ROOT)}"
                )

    canonical_only = sorted(
        name for name in _DTS_ENUM_SOURCES if name not in npm or not isinstance(npm.get(name), dict)
    )
    if canonical_only:
        problems.append(
            f"dts-enums: {WASM_DTS.relative_to(REPO_ROOT)} declares these enums but {npm_label} "
            f"does not export them as runtime tables: {canonical_only}"
        )
    return problems


# ---------------------------------------------------------------------------
# Check 6: public style-record field sets <-> their C ABI struct.
#
# `FontRecord` and friends are declared independently in three public
# surfaces and each one claims to mirror the C struct. Method-level drift
# checking cannot see a field that one surface forgot, which is how a font's
# `vertAlign` reached the WASM `.d.ts` alone: reading a superscript font and
# writing it back through the other bindings silently demoted it.
# ---------------------------------------------------------------------------

PYTHON_WORKBOOK = PYTHON_PKG_DIR / "workbook.py"

# Public type name -> the C struct it projects.
_STYLE_RECORD_STRUCTS = {
    "CellXf": "fm_cell_xf",
    "ColorSpec": "fm_color_spec",
    "FontRecord": "fm_font_record",
    "FillRecord": "fm_fill_record",
    "BorderSide": "fm_border_side",
}

# Deliberate omissions, keyed by (surface, type name). Python models a border
# side as a plain dict rather than a dataclass, so it has no `BorderSide` type
# to compare; the dict shape is covered by the wasm32 struct-layout check.
_STYLE_RECORD_EXEMPT_TYPES = {
    ("python", "BorderSide"),
}

# C presence flags that neither host language mirrors as a field, because both
# already have a way to say "absent": the TS surface omits the optional
# property, Python leaves it `None`. Naming the value field alone would let a
# host drop the distinction entirely, so each entry lists the value field the
# flag governs and both are required to exist together.
_STYLE_RECORD_PRESENCE_FLAGS = {
    "fm_cell_xf": {
        "has_text_rotation": "text_rotation",
        "has_indent": "indent",
        "has_relative_indent": "relative_indent",
        "has_shrink_to_fit": "shrink_to_fit",
        "has_reading_order": "reading_order",
    },
}

_TS_FIELD_RE = re.compile(r"^\s*([A-Za-z_$][A-Za-z0-9_$]*)\??\s*:", re.MULTILINE)
_PY_FIELD_RE = re.compile(r"^    ([a-z_][A-Za-z0-9_]*)\s*:", re.MULTILINE)


def _snake_to_camel(name: str) -> str:
    head, *rest = name.split("_")
    return head + "".join(part.capitalize() for part in rest)


def _c_struct_fields(header: str, struct_name: str, source: Path) -> List[str]:
    blocks = {name: body for body, name in _C_STRUCT_RE.findall(header)}
    body = blocks.get(struct_name)
    if body is None:
        print(f"check_binding_drift: struct {struct_name!r} not found in {source}", file=sys.stderr)
        sys.exit(2)
    fields: List[str] = []
    for declaration in body.split(";"):
        declaration = " ".join(declaration.split())
        if not declaration:
            continue
        match = re.fullmatch(r"(.+?)\s+([A-Za-z_][A-Za-z0-9_]*)(?:\[(\d+)\])?", declaration)
        if match and not match.group(2).startswith("_pad"):
            fields.append(match.group(2))
    return fields


def _python_class_fields(text: str, class_name: str) -> Optional[Set[str]]:
    match = re.search(r"^class %s:\s*$" % re.escape(class_name), text, re.MULTILINE)
    if not match:
        return None
    rest = text[match.end() :]
    end = re.search(r"^(?:@|class |def )", rest, re.MULTILINE)
    body = rest[: end.start()] if end else rest
    return set(_PY_FIELD_RE.findall(body))


def check_style_record_fields() -> List[str]:
    problems: List[str] = []
    header = re.sub(r"/\*.*?\*/", "", _read(CAPI_HEADER), flags=re.S)
    wasm_dts = _read(WASM_DTS)
    node_dts = _read(NODE_DTS)
    python_text = _read(PYTHON_WORKBOOK)

    for type_name, struct_name in sorted(_STYLE_RECORD_STRUCTS.items()):
        c_fields = _c_struct_fields(header, struct_name, CAPI_HEADER)
        # A presence flag is dropped from the comparison, but only once its
        # governed value field is confirmed present on the C side -- otherwise
        # renaming the value field would silently retire the flag with it.
        presence = _STYLE_RECORD_PRESENCE_FLAGS.get(struct_name, {})
        for flag, value_field in sorted(presence.items()):
            if flag in c_fields and value_field not in c_fields:
                problems.append(
                    f"style-record-fields: {struct_name}.{flag} is exempted as a presence flag for "
                    f"{value_field!r}, which no longer exists in {CAPI_HEADER.relative_to(REPO_ROOT)}"
                )
        stale_flags = sorted(flag for flag in presence if flag not in c_fields)
        if stale_flags:
            problems.append(f"style-record-fields: stale presence-flag exemptions for {struct_name}: {stale_flags}")
        c_fields = [name for name in c_fields if name not in presence]
        camel_fields = {_snake_to_camel(name) for name in c_fields}
        for surface, source, dts_text in (
            ("wasm", WASM_DTS, wasm_dts),
            ("node", NODE_DTS, node_dts),
        ):
            if (surface, type_name) in _STYLE_RECORD_EXEMPT_TYPES:
                continue
            ts_fields = set(_TS_FIELD_RE.findall(_find_interface_body(dts_text, type_name, source)))
            diff = _format_diff(
                f"{struct_name} in {CAPI_HEADER.relative_to(REPO_ROOT)}",
                camel_fields - ts_fields,
                f"{type_name} in {source.relative_to(REPO_ROOT)}",
                ts_fields - camel_fields,
            )
            if diff:
                problems.append(f"style-record-fields: {surface} {type_name} mismatch:\n" + "\n".join(diff))

        if ("python", type_name) in _STYLE_RECORD_EXEMPT_TYPES:
            continue
        py_fields = _python_class_fields(python_text, type_name)
        if py_fields is None:
            problems.append(
                f"style-record-fields: dataclass {type_name!r} not found in {PYTHON_WORKBOOK.relative_to(REPO_ROOT)}"
            )
            continue
        diff = _format_diff(
            f"{struct_name} in {CAPI_HEADER.relative_to(REPO_ROOT)}",
            set(c_fields) - py_fields,
            f"{type_name} in {PYTHON_WORKBOOK.relative_to(REPO_ROOT)}",
            py_fields - set(c_fields),
        )
        if diff:
            problems.append(f"style-record-fields: python {type_name} mismatch:\n" + "\n".join(diff))

    stale = {
        (surface, type_name)
        for surface, type_name in _STYLE_RECORD_EXEMPT_TYPES
        if type_name not in _STYLE_RECORD_STRUCTS
    }
    if stale:
        problems.append(f"style-record-fields: stale exemption entries: {sorted(stale)}")
    return problems


# ---------------------------------------------------------------------------
# Check 7: the npm-native package's staged copies <-> their sources.
#
# `stage.mjs` publishes `packages/npm-native/dist/` by copying `index.d.ts`
# and `index.mjs` verbatim out of the package root. Every other check in this
# file reads the source side only, so an edit that lands in the source but is
# never re-staged publishes a declaration file that no longer describes the
# runtime, or an ordinal table that no longer matches the one `dts-enums`
# verified. The WASM package already guards its own staged `.d.ts`
# (packages/npm/scripts/check-dts.mjs); npm-native had no equivalent, which
# matters most for a mechanical rewrite that touches one side of the pair.
# ---------------------------------------------------------------------------

# (source, staged copy) pairs `stage.mjs` produces with a verbatim copyFile.
_NODE_STAGED_COPIES = (
    (NODE_DTS, NODE_DIST_DIR / "index.d.ts"),
    (NODE_INDEX_MJS, NODE_DIST_DIR / "index.mjs"),
)


def check_staged_dist() -> List[str]:
    problems: List[str] = []
    for source, staged in _NODE_STAGED_COPIES:
        if not staged.is_file():
            # `dist/` is gitignored, so a fresh clone has nothing to compare
            # until the package is staged. Record that as not-checked instead
            # of passing: this check exists because a silent pass is
            # indistinguishable from a real one.
            _SKIPPED.append(
                f"staged-dist: {staged.relative_to(REPO_ROOT)} is absent; the npm-native "
                "package has not been staged in this tree (`make node-package`)"
            )
            continue
        if _read_bytes(source) != _read_bytes(staged):
            problems.append(
                f"staged-dist: {staged.relative_to(REPO_ROOT)} differs from "
                f"{source.relative_to(REPO_ROOT)}; the published copy is stale -- "
                "re-stage the package with `make node-package`"
            )
    return problems


CHECKS = {
    "python-exports": check_python_exports,
    "python-struct-layouts": check_python_struct_layouts,
    "dts-wasm": check_dts_wasm,
    "dts-node": check_dts_node,
    "readme-counts": check_readme_counts,
    "dts-enums": check_dts_enums,
    "style-record-fields": check_style_record_fields,
    "staged-dist": check_staged_dist,
}


def main(argv: List[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument(
        "check",
        nargs="?",
        default="all",
        choices=[*CHECKS.keys(), "all"],
        help="which drift check to run (default: all)",
    )
    args = parser.parse_args(argv)

    names = list(CHECKS.keys()) if args.check == "all" else [args.check]
    _SKIPPED.clear()
    problems: List[str] = []
    for name in names:
        problems.extend(CHECKS[name]())

    # Printed either way: a skip is not a pass, and the reader has to be able
    # to tell which parts of the run actually looked at anything.
    for note in _SKIPPED:
        print(f"check_binding_drift ({args.check}): SKIPPED {note}")

    if problems:
        print(f"check_binding_drift ({args.check}): DRIFT DETECTED")
        for problem in problems:
            print(problem)
        return 1

    print(f"check_binding_drift ({args.check}): no drift detected")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
