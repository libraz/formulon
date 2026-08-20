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
  python-call-signatures
                  Every `LIB.fm_*(...)` call in packages/python/formulon
                  passes as many arguments as src/c_api/formulon_c.h declares
                  parameters, and every scratch block passed to a pointer
                  parameter is at least (for a struct, exactly) the wasm32
                  size of that parameter's pointee. Argument types at
                  non-pointer positions are not covered; see the function's
                  docstring for the exact boundary.
  python-inline-structs
                  The C structs the Python binding decodes with a bare
                  `struct.unpack` instead of a `_structs.Struct` entry --
                  `fm_value_t`, `fm_print_range_t` -- have their size
                  literals and their decoders' offsets and field widths
                  measured against the header.
  dts-wasm        src/wasm/formulon.d.ts (Workbook / WorkbookCtor /
                  FormulonModule method surface) matches what is
                  registered in src/wasm/parts/bindings_register.cpp.
  dts-node        packages/npm-native/index.d.ts (Workbook /
                  WorkbookCtor / free-function surface) matches what is
                  registered in src/node_addon/parts/workbook_class.cc
                  and src/node_addon/addon.cc. It also enforces the exact
                  intentional WASM-only / Node-only method allowlists.
  dts-shared-shapes
                  The two published declaration files agree on the *shape*
                  of everything they both declare, not merely on the names:
                  a method present in both `Workbook` (or `WorkbookCtor`)
                  interfaces declares the same return type, and a record
                  type present in both declares the same field set, with
                  optionality counted as part of the field. A type declared
                  on one surface only needs an entry in
                  `_DTS_SURFACE_ONLY_TYPES`; a deliberately divergent return
                  type needs one in `_DTS_RETURN_TYPE_EXEMPT_METHODS`.
  pure-js-helpers Helpers with no native entry point behind them
                  (`NODE_PURE_JS_FREE_FUNCTIONS`) are exported by both npm
                  packages' `index.mjs`, declared in both declaration files,
                  and implemented with identical source. The two packages
                  share no module, so this is the only thing holding the
                  copies together.
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
  abi-baseline    Every entry point in the released C ABI surface
                  (tools/dev/c_abi_baseline.txt) is still declared in
                  src/c_api/formulon_c.h with the same signature, unless
                  the divergence is recorded in tools/dev/c_abi_breaks.txt.
                  Five of the eight base/`_ex` families are reached only
                  through their `_ex` variant, so deleting or renaming the
                  base leaves all four surfaces green and the break ships
                  silently -- this is the only artifact that notices. A
                  stale ledger entry fails too, so the ledger cannot
                  outlive the break it excuses.
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

C_ABI_BASELINE = Path(__file__).resolve().parent / "c_abi_baseline.txt"
C_ABI_BREAKS = Path(__file__).resolve().parent / "c_abi_breaks.txt"
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
# Check 1c: Python `LIB.fm_*` call sites <-> the C header's declarations.
#
# `python-exports` matches call *names* only. Nothing compares what a call
# passes against what the entry point declares, so a parameter added to an
# existing `fm_*` function -- the shape a 1.0-frozen ABI is most likely to
# grow -- leaves the Python side calling the old signature and is caught only
# if a runtime binding test happens to exercise that path.
# ---------------------------------------------------------------------------

_C_FUNCTION_DECL_RE = re.compile(
    r"FM_API\s+([A-Za-z_][A-Za-z0-9_ *]*?)\b(fm_[A-Za-z0-9_]+)\s*\(([^;]*?)\)\s*;",
    re.S,
)

# Scratch readers whose name pins how many bytes come back out of the slot.
# `read_cstr` is deliberately absent: it walks to a NUL rather than reading a
# fixed width, so it is checked for the deref mistake only (below).
_SCRATCH_READER_WIDTHS = {"read_u32": 4, "read_i32": 4, "read_f64": 8}


def _split_c_params(text: str) -> List[str]:
    """Splits a parameter list on its top-level commas."""
    parts: List[str] = []
    depth = 0
    current = ""
    for char in text:
        if char == "(":
            depth += 1
        elif char == ")":
            depth -= 1
        if char == "," and depth == 0:
            parts.append(current.strip())
            current = ""
        else:
            current += char
    parts.append(current.strip())
    return parts


def _c_function_params(header: str) -> dict[str, List[str]]:
    """Maps each `FM_API` entry point to its declared parameter types.

    The parameter *name* is stripped, so the values are types in declaration
    order (`["fm_workbook_t*", "size_t", "fm_value_t*"]`). Every parameter in
    this header is named, so dropping the last whitespace-delimited token is
    unambiguous.
    """
    return {name: params for name, (_, params) in c_abi_declarations(header).items()}


def c_abi_declarations(header: str) -> dict[str, tuple[str, List[str]]]:
    """Maps each `FM_API` entry point to its `(return type, parameter types)`.

    Public because `gen_c_abi_baseline.py` writes the released-surface file
    with it: generating the baseline through the same parser the check reads
    it back with keeps a formatting difference from masquerading as a break.
    """
    declarations: dict[str, tuple[str, List[str]]] = {}
    for ret, name, params in _C_FUNCTION_DECL_RE.findall(header):
        ret = " ".join(ret.split())
        params = " ".join(params.split())
        if params in ("", "void"):
            declarations[name] = (ret, [])
            continue
        declarations[name] = (ret, [" ".join(param.split()[:-1]) for param in _split_c_params(params)])
    return declarations


def _module_int_constants(paths: List[Path]) -> dict[str, int]:
    """Module-level `NAME = <int literal>` bindings across the binding package.

    Collected across files because the constant a scratch allocation is sized
    with may be defined in one module and imported into another
    (`fm_value_t_size` lives in `_c.py` and is used from `workbook.py`).
    """
    constants: dict[str, int] = {}
    for path in paths:
        for statement in ast.parse(_read(path), filename=str(path)).body:
            if not isinstance(statement, ast.Assign) or len(statement.targets) != 1:
                continue
            target = statement.targets[0]
            if isinstance(target, ast.Name) and isinstance(statement.value, ast.Constant):
                if isinstance(statement.value.value, int) and not isinstance(statement.value.value, bool):
                    constants[target.id] = statement.value.value
    return constants


def _is_lib_call(node: ast.AST, attribute: Optional[str] = None) -> bool:
    """True for `LIB.<attribute>(...)` (any `LIB.*` call when unspecified)."""
    if not isinstance(node, ast.Call) or not isinstance(node.func, ast.Attribute):
        return False
    if not (isinstance(node.func.value, ast.Name) and node.func.value.id == "LIB"):
        return False
    return attribute is None or node.func.attr == attribute


def _scratch_slot_widths(
    function: ast.AST, constants: dict[str, int], struct_module: object
) -> dict[str, Optional[int]]:
    """Local name -> byte width of the WASM scratch block bound to it.

    A width of `None` marks a slot whose size this resolver cannot pin (an
    array allocation sized from a runtime length, or a name rebound to two
    different widths). Those are reported rather than skipped when the C side
    says the parameter is a struct pointer, because an unmeasurable buffer
    behind a by-pointer struct is exactly the case a size guard exists for.
    """
    widths: dict[str, Optional[int]] = {}

    def record(name: str, width: Optional[int]) -> None:
        if name in widths and widths[name] != width:
            widths[name] = None
        else:
            widths[name] = width

    for node in ast.walk(function):
        if not isinstance(node, ast.Assign) or len(node.targets) != 1:
            continue
        target = node.targets[0]
        if not isinstance(target, ast.Name) or not isinstance(node.value, ast.Call):
            continue
        call = node.value
        callee = call.func
        # `x = _alloc_out_ptr()` -- the package's 4-byte out-i32 / out-ptr slot.
        if isinstance(callee, ast.Name) and callee.id == "_alloc_out_ptr":
            record(target.id, 4)
        # `x = S.alloc_struct(LIB, S.LAYOUT)` -- sized from the layout table
        # `python-struct-layouts` already pins against the header.
        elif (isinstance(callee, ast.Attribute) and callee.attr == "alloc_struct") or (
            isinstance(callee, ast.Name) and callee.id == "alloc_struct"
        ):
            layout = call.args[1] if len(call.args) > 1 else None
            layout_name = layout.attr if isinstance(layout, ast.Attribute) else None
            struct = getattr(struct_module, layout_name, None) if layout_name else None
            record(target.id, struct.size if isinstance(struct, struct_module.Struct) else None)
        # `x = LIB.alloc(N)` / `LIB.alloc(CONSTANT)` -- a hand-sized block.
        elif _is_lib_call(call, "alloc") and call.args:
            size = call.args[0]
            if isinstance(size, ast.Constant) and isinstance(size.value, int):
                record(target.id, size.value)
            elif isinstance(size, ast.Name) and size.id in constants:
                record(target.id, constants[size.id])
            else:
                record(target.id, None)
    return widths


def _scratch_slot_readers(function: ast.AST) -> dict[str, Set[object]]:
    """Local name -> the reads taken against that scratch slot.

    An entry is either a `_SCRATCH_READER_WIDTHS` key, `("bytes", N)` for a
    constant-length `LIB.read_bytes`, or `"read_cstr"`.
    """
    readers: dict[str, Set[object]] = {}
    for node in ast.walk(function):
        if not isinstance(node, ast.Call) or not _is_lib_call(node) or not node.args:
            continue
        first = node.args[0]
        if not isinstance(first, ast.Name):
            continue
        attribute = node.func.attr if isinstance(node.func, ast.Attribute) else ""
        if attribute in _SCRATCH_READER_WIDTHS or attribute == "read_cstr":
            readers.setdefault(first.id, set()).add(attribute)
        elif attribute == "read_bytes" and len(node.args) > 1 and isinstance(node.args[1], ast.Constant):
            readers.setdefault(first.id, set()).add(("bytes", node.args[1].value))
    return readers


def _pointee(ctype: str) -> Optional[str]:
    """The type a parameter points at, or `None` if it is not a pointer."""
    ctype = ctype.strip()
    if not ctype.endswith("*"):
        return None
    return ctype[:-1].strip()


def _check_out_param_width(
    layouts: _Wasm32Layouts,
    site: str,
    fn_name: str,
    position: int,
    ctype: str,
    width: Optional[int],
    reads: Set[object],
) -> List[str]:
    """One scratch slot against the parameter it is passed to."""
    problems: List[str] = []
    pointee = _pointee(ctype)
    if pointee is None:
        # Emscripten lowers a by-value struct parameter to a pointer into
        # linear memory, so a scratch block is the correct thing to pass for
        # one. Any other non-pointer parameter takes a scalar, and handing it
        # a scratch address means the callee reads the pointer as the value.
        bare = ctype.replace("const ", "").strip()
        if not layouts.is_struct(bare):
            return [
                f"python-call-signatures: {site}: {fn_name} parameter {position} is {ctype!r}, "
                "which takes a scalar by value, but the binding passes a scratch-block pointer"
            ]
        pointee_size, _ = layouts.extent(bare)
        is_struct = True
    else:
        bare = pointee.replace("const ", "").strip()
        if bare.endswith("*"):
            pointee_size, is_struct = 4, False  # wasm32 pointer-to-pointer
        else:
            try:
                pointee_size, _ = layouts.extent(bare)
            except _LayoutError as exc:
                return [f"python-call-signatures: {site}: {fn_name} parameter {position} ({ctype}): {exc}"]
            is_struct = layouts.is_struct(bare)

    if width is None:
        if is_struct:
            problems.append(
                f"python-call-signatures: {site}: {fn_name} parameter {position} is {ctype!r}, "
                "but the size of the scratch block passed to it cannot be resolved; give the "
                "struct a `_structs.Struct` layout or size the allocation from a module constant"
            )
        return problems

    if width < pointee_size:
        problems.append(
            f"python-call-signatures: {site}: {fn_name} parameter {position} is {ctype!r} "
            f"(wasm32 pointee {pointee_size} bytes) but the binding allocates {width} bytes"
        )
    elif is_struct and width != pointee_size:
        problems.append(
            f"python-call-signatures: {site}: {fn_name} parameter {position} is {ctype!r} "
            f"(wasm32 sizeof {pointee_size}) but the binding allocates {width} bytes; a struct "
            "out-parameter block must match the C size exactly"
        )

    for read in sorted(reads, key=repr):
        read_width = _SCRATCH_READER_WIDTHS.get(read) if isinstance(read, str) else read[1]
        if read == "read_cstr":
            if bare.endswith("*"):
                problems.append(
                    f"python-call-signatures: {site}: {fn_name} parameter {position} is {ctype!r}, "
                    "so its slot holds a pointer; `read_cstr` on the slot itself decodes the "
                    "pointer's bytes as text instead of dereferencing it"
                )
            continue
        if read_width is None:
            continue
        if read_width > width:
            problems.append(
                f"python-call-signatures: {site}: {fn_name} parameter {position}'s slot is "
                f"{width} bytes but is read with {read!r} ({read_width} bytes)"
            )
        elif read_width > pointee_size:
            problems.append(
                f"python-call-signatures: {site}: {fn_name} parameter {position} is {ctype!r} "
                f"(wasm32 pointee {pointee_size} bytes) but its slot is read with {read!r} "
                f"({read_width} bytes)"
            )
    return problems


def check_python_call_signatures() -> List[str]:
    """Verify Python's `LIB.fm_*` calls against the C header's declarations.

    Covered, for every `LIB.fm_*(...)` call in `packages/python/formulon`:

    * **Arity.** The number of positional arguments equals the header's
      parameter count. A call that splats (`*pointers`) is checked as a lower
      bound only and named in the run's SKIPPED list.
    * **Scratch-block width.** When an argument is a local bound to a
      recognised scratch allocation (``_alloc_out_ptr()``,
      ``S.alloc_struct(LIB, S.LAYOUT)``, or ``LIB.alloc(...)`` with a constant
      or module-constant size), the block is compared against the wasm32 size
      of what the parameter addresses: never smaller, and exactly equal when
      that is a struct. This covers both pointer parameters and the by-value
      struct parameters Emscripten lowers to a pointer. A struct whose block
      size cannot be resolved is reported rather than skipped.
    * **Read width.** Reads taken against such a slot (``read_u32`` /
      ``read_i32`` / ``read_f64`` / constant-length ``read_bytes``) must not
      exceed either the block or the pointee, and ``read_cstr`` must not be
      applied to a slot that holds a pointer.

    NOT covered: argument *types* at non-pointer positions (nothing on the
    Python side records whether a value was meant to be a `uint32_t` or a
    `double`; the wasmtime layer passes plain Python ints and floats), return
    types beyond what `python-exports` already pins, and both non-Python
    bindings -- embind and N-API bind C++ entry points directly rather than
    the `fm_*` C ABI, so this header has no arity relationship to them.
    """
    header = re.sub(r"/\*.*?\*/", "", _read(CAPI_HEADER), flags=re.S)
    declarations = _c_function_params(header)
    layouts = _Wasm32Layouts(header)

    spec = importlib.util.spec_from_file_location("formulon_struct_layouts_sig", PYTHON_STRUCTS)
    if spec is None or spec.loader is None:
        return ["python-call-signatures: could not load packages/python/formulon/_structs.py"]
    struct_module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(struct_module)

    sources = sorted(PYTHON_PKG_DIR.glob("*.py"))
    constants = _module_int_constants(sources)

    problems: List[str] = []
    checked_calls = 0
    for path in sources:
        label = path.relative_to(REPO_ROOT)
        tree = ast.parse(_read(path), filename=str(path))
        # Analysis is per enclosing function: a scratch local is only
        # meaningfully tied to the calls in the scope that allocated it.
        scopes = [node for node in ast.walk(tree) if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))]
        for scope in scopes:
            widths = _scratch_slot_widths(scope, constants, struct_module)
            readers = _scratch_slot_readers(scope)
            for node in ast.walk(scope):
                if not isinstance(node, ast.Call) or not isinstance(node.func, ast.Attribute):
                    continue
                if not _is_lib_call(node):
                    continue
                fn_name = node.func.attr
                if not fn_name.startswith("fm_"):
                    continue
                site = f"{label}:{node.lineno}"
                params = declarations.get(fn_name)
                if params is None:
                    problems.append(f"python-call-signatures: {site}: {fn_name} is not declared in the C header")
                    continue
                checked_calls += 1
                if node.keywords:
                    problems.append(
                        f"python-call-signatures: {site}: {fn_name} is called with keyword arguments; "
                        "the wasmtime export takes positional arguments only"
                    )
                    continue
                starred = any(isinstance(arg, ast.Starred) for arg in node.args)
                if starred:
                    fixed = len(node.args) - 1
                    if fixed > len(params):
                        problems.append(
                            f"python-call-signatures: {site}: {fn_name} passes {fixed} arguments before its "
                            f"splat but the C header declares only {len(params)} parameters"
                        )
                    _SKIPPED.append(
                        f"python-call-signatures: {site}: {fn_name} splats its trailing arguments, so only "
                        f"the {fixed}-argument lower bound is checked against the header's {len(params)}"
                    )
                    continue
                if len(node.args) != len(params):
                    problems.append(
                        f"python-call-signatures: {site}: {fn_name} is called with {len(node.args)} "
                        f"arguments but the C header declares {len(params)}: {params}"
                    )
                    continue
                for position, argument in enumerate(node.args):
                    if not isinstance(argument, ast.Name) or argument.id not in widths:
                        continue
                    problems.extend(
                        _check_out_param_width(
                            layouts,
                            site,
                            fn_name,
                            position,
                            params[position],
                            widths[argument.id],
                            readers.get(argument.id, set()),
                        )
                    )

    if not checked_calls:
        problems.append(
            "python-call-signatures: no `LIB.fm_*` call sites found in "
            f"{PYTHON_PKG_DIR.relative_to(REPO_ROOT)}; the check has stopped looking at anything"
        )
    # A call inside a nested `def` is reached both as part of the enclosing
    # scope's walk and as its own scope, so the same finding can be produced
    # twice. Deduplicate in place rather than restricting the walk: an inner
    # closure that uses a scratch slot its parent allocated still has to be
    # checked against that slot's width.
    return list(dict.fromkeys(problems))


# ---------------------------------------------------------------------------
# Check 1d: C structs the Python binding decodes inline <-> the C header.
#
# `python-struct-layouts` only sees structs that have a `_structs.Struct`
# entry. A struct small enough to decode with a bare `struct.unpack` -- an
# `fm_value_t`, an `fm_print_range_t` -- has no such entry, so its field
# offsets and widths live as literals in the decoding function and nothing
# compares them to the header.
# ---------------------------------------------------------------------------

# (source file, decoder qualname, C struct). Each decoder is expected to
# consume the whole struct, so its `struct` calls are read as a description
# of the C layout and compared field for field.
_PYTHON_INLINE_DECODERS = (
    ("Value._from_wasm", "fm_value_t"),
    ("Workbook.paginate", "fm_print_range_t"),
)

# Python-side size literals for a struct with no `_structs.Struct` entry, as
# (module, binding name, C struct). A `Struct`-shaped tuple binding is
# compared on both size and alignment; a bare int on size alone.
_PYTHON_SIZE_CONSTANTS = (
    (PYTHON_C_BINDING, "fm_value_t_size", "fm_value_t"),
    (PYTHON_STRUCTS, "VALUE_BLOB", "fm_value_t"),
)

# C structs that cross the ABI but that no binding marshals, so there is no
# second side to compare against. Each is pinned by a native/wasm32
# `static_assert` tripwire in tests/c_api instead; listing them here keeps the
# absence deliberate, and `python-call-signatures` fails if a binding starts
# passing one an unmeasurable block.
_UNMODELLED_C_STRUCTS = {"fm_styles_batch"}

_STRUCT_FORMAT_WIDTHS = {"b": 1, "B": 1, "h": 2, "H": 2, "i": 4, "I": 4, "q": 8, "Q": 8, "f": 4, "d": 8}


def _struct_format_widths(fmt: str) -> Optional[List[int]]:
    """Per-item byte widths of a little-endian `struct` format string."""
    if not fmt.startswith("<"):
        return None
    widths: List[int] = []
    for char in fmt[1:]:
        width = _STRUCT_FORMAT_WIDTHS.get(char)
        if width is None:
            return None
        widths.append(width)
    return widths


def _find_qualified_function(tree: ast.AST, qualname: str) -> Optional[ast.AST]:
    class_name, _, function_name = qualname.rpartition(".")
    for node in ast.walk(tree):
        if not isinstance(node, ast.ClassDef) or node.name != class_name:
            continue
        for member in node.body:
            if isinstance(member, (ast.FunctionDef, ast.AsyncFunctionDef)) and member.name == function_name:
                return member
    return None


def check_python_inline_structs() -> List[str]:
    """Verify Python's inline `struct.unpack` decoders against the C header.

    Two things are compared, both measured from the C declarations rather
    than tabulated here:

    * A size literal the binding carries for a struct with no
      `_structs.Struct` entry (`fm_value_t_size`, `VALUE_BLOB`) equals the
      struct's wasm32 size, and its alignment where the binding records one.
    * Inside each decoder in `_PYTHON_INLINE_DECODERS`, every
      `struct.unpack_from(fmt, buf, offset)` lands on a C member offset and
      reads no wider than that member, every whole-struct `struct.unpack(fmt,
      ...)` describes the C members' widths in order, and every
      constant-length `LIB.read_bytes` spans exactly the struct.

    NOT covered: the *meaning* of a field (a decoder that reads the right
    width from the right offset into the wrong attribute still passes), and
    any struct in `_UNMODELLED_C_STRUCTS`, which no binding marshals and
    which therefore has only a `static_assert` tripwire.
    """
    header = re.sub(r"/\*.*?\*/", "", _read(CAPI_HEADER), flags=re.S)
    layouts = _Wasm32Layouts(header)
    constants = _module_int_constants(sorted(PYTHON_PKG_DIR.glob("*.py")))
    problems: List[str] = []

    def measure(struct_name: str) -> Optional[tuple[dict[str, int], int, int, dict[str, int]]]:
        try:
            members = layouts.members(struct_name)
        except _LayoutError as exc:
            problems.append(f"python-inline-structs: {exc}")
            return None
        offsets, size, align = _wasm32_layout([(m.name, m.size, m.align) for m in members])
        return offsets, size, align, {m.name: m.size for m in members}

    for struct_name in sorted(_UNMODELLED_C_STRUCTS):
        if not layouts.is_struct(struct_name):
            problems.append(
                f"python-inline-structs: {struct_name} is listed as unmodelled by every binding "
                f"but no longer exists in {CAPI_HEADER.relative_to(REPO_ROOT)}"
            )

    for module_path, binding_name, struct_name in _PYTHON_SIZE_CONSTANTS:
        measured = measure(struct_name)
        if measured is None:
            continue
        _, c_size, c_align, _ = measured
        label = module_path.relative_to(REPO_ROOT)
        found = False
        for statement in ast.parse(_read(module_path), filename=str(module_path)).body:
            if not isinstance(statement, ast.Assign) or len(statement.targets) != 1:
                continue
            target = statement.targets[0]
            if not isinstance(target, ast.Name) or target.id != binding_name:
                continue
            found = True
            value = statement.value
            if isinstance(value, ast.Constant) and isinstance(value.value, int):
                if value.value != c_size:
                    problems.append(
                        f"python-inline-structs: {label}: {binding_name} is {value.value}, "
                        f"but wasm32 sizeof({struct_name}) is {c_size}"
                    )
            elif isinstance(value, ast.Tuple) and len(value.elts) == 3:
                elements = [element.value if isinstance(element, ast.Constant) else None for element in value.elts]
                if elements[1] != c_size:
                    problems.append(
                        f"python-inline-structs: {label}: {binding_name} declares size {elements[1]}, "
                        f"but wasm32 sizeof({struct_name}) is {c_size}"
                    )
                if elements[2] != c_align:
                    problems.append(
                        f"python-inline-structs: {label}: {binding_name} declares alignment {elements[2]}, "
                        f"but wasm32 alignof({struct_name}) is {c_align}"
                    )
            else:
                problems.append(
                    f"python-inline-structs: {label}: {binding_name} is neither an int literal nor a "
                    "(kind, size, alignment) literal tuple, so its layout claim cannot be read"
                )
        if not found:
            problems.append(
                f"python-inline-structs: {label} no longer defines {binding_name}, "
                f"which is where the binding's {struct_name} size lives"
            )

    tree = ast.parse(_read(PYTHON_WORKBOOK), filename=str(PYTHON_WORKBOOK))
    label = PYTHON_WORKBOOK.relative_to(REPO_ROOT)
    for qualname, struct_name in _PYTHON_INLINE_DECODERS:
        function = _find_qualified_function(tree, qualname)
        if function is None:
            problems.append(
                f"python-inline-structs: {label} no longer defines {qualname}, "
                f"which is where the inline {struct_name} decoding lives"
            )
            continue
        measured = measure(struct_name)
        if measured is None:
            continue
        c_offsets, c_size, _, c_sizes = measured
        member_widths = [c_sizes[name] for name in c_offsets]

        for node in ast.walk(function):
            if isinstance(node, ast.Call) and _is_lib_call(node, "read_bytes") and len(node.args) > 1:
                length = node.args[1]
                span = None
                if isinstance(length, ast.Constant) and isinstance(length.value, int):
                    span = length.value
                elif isinstance(length, ast.Name):
                    span = constants.get(length.id)
                if span is not None and span != c_size:
                    problems.append(
                        f"python-inline-structs: {label}:{node.lineno}: {qualname} reads {span} bytes "
                        f"for a {struct_name}, whose wasm32 size is {c_size}"
                    )
            if not (
                isinstance(node, ast.Call)
                and isinstance(node.func, ast.Attribute)
                and isinstance(node.func.value, ast.Name)
                and node.func.value.id == "struct"
            ):
                continue
            if not node.args or not isinstance(node.args[0], ast.Constant):
                continue
            fmt = node.args[0].value
            if not isinstance(fmt, str):
                continue
            widths = _struct_format_widths(fmt)
            if widths is None:
                problems.append(
                    f"python-inline-structs: {label}:{node.lineno}: {qualname} uses the format {fmt!r}, "
                    "which is not a little-endian fixed-width layout this check can measure"
                )
                continue
            if node.func.attr == "unpack_from":
                if len(node.args) < 3 or not isinstance(node.args[2], ast.Constant):
                    continue
                offset = node.args[2].value
                owner = next((name for name, at in c_offsets.items() if at == offset), None)
                if owner is None:
                    problems.append(
                        f"python-inline-structs: {label}:{node.lineno}: {qualname} decodes at offset "
                        f"{offset}, which is not a member offset of {struct_name} "
                        f"({sorted(c_offsets.items(), key=lambda item: item[1])})"
                    )
                elif sum(widths) > c_sizes[owner]:
                    problems.append(
                        f"python-inline-structs: {label}:{node.lineno}: {qualname} reads {sum(widths)} bytes "
                        f"with {fmt!r} at offset {offset}, but {struct_name}.{owner} is {c_sizes[owner]} bytes"
                    )
            elif node.func.attr == "unpack" and widths != member_widths:
                problems.append(
                    f"python-inline-structs: {label}:{node.lineno}: {qualname} decodes a whole "
                    f"{struct_name} with {fmt!r} (field widths {widths}), but the C members are "
                    f"{list(zip(c_offsets, member_widths))}"
                )
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
# Check 2c: shared `.d.ts` shapes -- return types and record field sets.
#
# `dts-wasm` and `dts-node` compare method *names* only. That is enough to
# notice a method that exists on one surface and not the other, and nothing
# else: a shared method whose two declarations disagree about what comes back,
# or a record type that grew a field on one side, passes both of them. The
# README's parity claim is about the shapes, not the names, so this check
# compares the shapes.
#
# Types declared in only one of the two published declaration files. Each has
# to be reachable only from that surface's own method allowlist -- a type that
# becomes one-sided for any other reason is a missing declaration, and listing
# it here is the deliberate act that says otherwise.
# ---------------------------------------------------------------------------

_DTS_SURFACE_ONLY_TYPES = {
    # The embind module factory and its options: there is no equivalent on a
    # native addon, which is `require`d rather than instantiated.
    ("wasm", "FormulonModule"),
    ("wasm", "FormulonModuleOptions"),
    # Reachable only from `getSheetAutoFilterXml` / `createTable` /
    # `updateTable`, all of which are in WASM_ONLY_METHODS.
    ("wasm", "SheetAutoFilterXmlResult"),
    ("wasm", "TableInput"),
    ("wasm", "TableUpdateInput"),
}

# Shared methods whose two declarations are allowed to differ, with the
# reason. Empty by design: a difference that is not worth an entry here is a
# difference that should not exist.
_DTS_RETURN_TYPE_EXEMPT_METHODS: dict[str, str] = {}

# Top-level members of an exported interface are indented exactly two spaces;
# anything deeper belongs to a nested object type and is compared as part of
# its parent's field text, not on its own.
_TS_MEMBER_NAME_RE = re.compile(r"^ {2}([A-Za-z_$][A-Za-z0-9_$]*)\s*\(", re.MULTILINE)
_TS_MEMBER_FIELD_RE = re.compile(r"^ {2}(?:readonly\s+)?([A-Za-z_$][A-Za-z0-9_$]*)(\??)\s*:", re.MULTILINE)


def _ts_return_types(body: str) -> dict[str, str]:
    """Maps each method in an interface body to its declared return type.

    The parameter list is skipped by bracket balancing so a wrapped or
    generic signature is read the same as a single-line one; whitespace in
    the return type is collapsed so formatting differences do not register
    as drift.
    """
    out: dict[str, str] = {}
    for match in _TS_MEMBER_NAME_RE.finditer(body):
        open_paren = body.index("(", match.start())
        depth = 0
        close_paren = -1
        for i in range(open_paren, len(body)):
            char = body[i]
            if char in "(<[":
                depth += 1
            elif char in ")>]":
                depth -= 1
                if depth == 0 and char == ")":
                    close_paren = i
                    break
        if close_paren < 0:
            continue
        rest = body[close_paren + 1 :]
        if ";" not in rest:
            continue
        declared = rest[: rest.index(";")].strip()
        if declared.startswith(":"):
            declared = declared[1:]
        out[match.group(1)] = " ".join(declared.split())
    return out


def _ts_interface_declaration(text: str, name: str, source: Path) -> tuple[str, str]:
    """Returns an interface's `extends` clause (may be empty) and its body.

    Unlike `_find_interface_body` this tolerates `extends`, which matters
    here: a record type that inherits its fields declares none of them
    locally, so the base list is part of the shape being compared.
    """
    match = re.search(r"^export interface %s\b([^{]*)\{" % re.escape(name), text, re.MULTILINE)
    if not match:
        print(f"check_binding_drift: interface {name!r} not found in {source}", file=sys.stderr)
        sys.exit(2)
    return " ".join(match.group(1).split()), _extract_braced_block(text, match.end())


def _ts_field_set(body: str) -> Set[str]:
    """Field names of an interface body, each suffixed with `?` when optional.

    Optionality is part of the compared shape on purpose: "present only on
    success" is how both declaration files spell a payload key the runtime
    drops, so a required declaration on one side and an optional one on the
    other is a real disagreement about what a caller may dereference.
    """
    return {name + optional for name, optional in _TS_MEMBER_FIELD_RE.findall(body)}


def check_dts_shared_shapes() -> List[str]:
    problems: List[str] = []
    wasm_dts = _read(WASM_DTS)
    node_dts = _read(NODE_DTS)
    wasm_label = str(WASM_DTS.relative_to(REPO_ROOT))
    node_label = str(NODE_DTS.relative_to(REPO_ROOT))

    for interface in ("Workbook", "WorkbookCtor"):
        wasm_returns = _ts_return_types(_find_interface_body(wasm_dts, interface, WASM_DTS))
        node_returns = _ts_return_types(_find_interface_body(node_dts, interface, NODE_DTS))
        for name in sorted(set(wasm_returns) & set(node_returns)):
            if wasm_returns[name] == node_returns[name]:
                if name in _DTS_RETURN_TYPE_EXEMPT_METHODS:
                    problems.append(
                        f"dts-shared-shapes: {name} is listed as a deliberate return-type "
                        "difference but both surfaces now declare the same type; drop the entry"
                    )
                continue
            if name in _DTS_RETURN_TYPE_EXEMPT_METHODS:
                continue
            problems.append(
                f"dts-shared-shapes: {interface}.{name} returns a different type on each surface:\n"
                f"  {wasm_label}: {wasm_returns[name]}\n"
                f"  {node_label}: {node_returns[name]}"
            )

    wasm_types = set(re.findall(r"^export interface ([A-Za-z0-9_]+)", wasm_dts, re.MULTILINE))
    node_types = set(re.findall(r"^export interface ([A-Za-z0-9_]+)", node_dts, re.MULTILINE))
    surface_only = {("wasm", name) for name in wasm_types - node_types}
    surface_only |= {("node", name) for name in node_types - wasm_types}
    undeclared = surface_only - _DTS_SURFACE_ONLY_TYPES
    if undeclared:
        problems.append(
            "dts-shared-shapes: type declared on one surface only, without an entry in "
            f"_DTS_SURFACE_ONLY_TYPES: {sorted(undeclared)}"
        )
    stale = _DTS_SURFACE_ONLY_TYPES - surface_only
    if stale:
        problems.append(f"dts-shared-shapes: stale _DTS_SURFACE_ONLY_TYPES entries: {sorted(stale)}")

    for name in sorted(wasm_types & node_types):
        if name in ("Workbook", "WorkbookCtor"):
            continue
        wasm_extends, wasm_body = _ts_interface_declaration(wasm_dts, name, WASM_DTS)
        node_extends, node_body = _ts_interface_declaration(node_dts, name, NODE_DTS)
        if wasm_extends != node_extends:
            problems.append(
                f"dts-shared-shapes: {name} inherits differently on each surface:\n"
                f"  {wasm_label}: {wasm_extends or '(nothing)'}\n"
                f"  {node_label}: {node_extends or '(nothing)'}"
            )
        diff = _format_diff(
            f"{name} in {wasm_label}",
            _ts_field_set(wasm_body) - _ts_field_set(node_body),
            f"{name} in {node_label}",
            _ts_field_set(node_body) - _ts_field_set(wasm_body),
        )
        if diff:
            problems.append(f"dts-shared-shapes: {name} field set mismatch:\n" + "\n".join(diff))

    return problems


# ---------------------------------------------------------------------------
# Check 2d: pure-JS helpers shipped by both npm packages.
#
# A helper with no native entry point behind it has to be written out once per
# package, because the two packages share no module. Nothing then holds the
# copies together, and nothing notices when one package documents a helper it
# does not ship -- which is how the WASM declaration file came to send its
# readers to a package that is not published. Both halves are checked here:
# the helper is exported and declared by both packages, and the two
# implementations are the same text.
# ---------------------------------------------------------------------------


def _js_exported_function_source(text: str, name: str) -> Optional[str]:
    """Whitespace-normalised body of `export function <name>(...) { ... }`."""
    match = re.search(r"^export function %s\s*\(" % re.escape(name), text, re.MULTILINE)
    if not match:
        return None
    body_open = text.index("{", match.end())
    return " ".join(_extract_braced_block(text, body_open + 1).split())


def check_pure_js_helpers() -> List[str]:
    problems: List[str] = []
    sources = {
        "npm": (NPM_INDEX_MJS, WASM_DTS),
        "npm-native": (NODE_INDEX_MJS, NODE_DTS),
    }

    for name in sorted(NODE_PURE_JS_FREE_FUNCTIONS):
        bodies: dict[str, str] = {}
        for package, (mjs_path, dts_path) in sources.items():
            body = _js_exported_function_source(_read(mjs_path), name)
            if body is None:
                problems.append(
                    f"pure-js-helpers: {mjs_path.relative_to(REPO_ROOT)} does not export a "
                    f"`{name}` function. A helper only one package ships cannot be named in a "
                    "declaration file both packages publish."
                )
            else:
                bodies[package] = body
            if not re.search(r"^export function %s\s*\(" % re.escape(name), _read(dts_path), re.MULTILINE):
                problems.append(
                    f"pure-js-helpers: {dts_path.relative_to(REPO_ROOT)} does not declare `export function {name}`"
                )
        if len(bodies) == len(sources) and len(set(bodies.values())) != 1:
            problems.append(
                f"pure-js-helpers: the `{name}` implementations have diverged between "
                f"{NPM_INDEX_MJS.relative_to(REPO_ROOT)} and {NODE_INDEX_MJS.relative_to(REPO_ROOT)}. "
                "The two packages share no module, so the copies are held together here or not "
                "at all."
            )

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


def _parse_abi_manifest(path: Path) -> List[str]:
    """Non-comment, non-blank lines of a baseline / ledger file."""
    lines = []
    for raw in _read(path).splitlines():
        line = raw.strip()
        if line and not line.startswith("#"):
            lines.append(line)
    return lines


_ABI_BREAK_RE = re.compile(r"^(removed|signature|retyped)\s+(fm_[A-Za-z0-9_]+)\s*::\s*(\S.*)$")
_ABI_BASELINE_RE = re.compile(r"^(.+?)\s+(fm_[A-Za-z0-9_]+)\((.*)\)$")


def check_abi_baseline() -> List[str]:
    """The released C ABI surface is still declared, or its loss is recorded.

    A deletion or rename of a base entry point is invisible to every other
    check here: five of the eight base/`_ex` families have no in-repo caller
    at all, so the C API tests recompile against the new header and pass. The
    only party that notices is a third-party consumer compiled against the
    released header, which no test can stand in for -- hence a pinned copy of
    that header's surface plus an explicit ledger of what was broken on
    purpose.

    Signature comparison is by declared parameter *type*, so it catches an
    added, dropped or retyped parameter. It does NOT catch a by-value struct
    that keeps its name and changes size; that is a calling convention change
    the `static_assert(sizeof(...))` tripwires in the C ABI tests pin instead.
    """
    problems: List[str] = []

    baseline: dict[str, tuple[str, List[str]]] = {}
    for line in _parse_abi_manifest(C_ABI_BASELINE):
        match = _ABI_BASELINE_RE.match(line)
        if match is None:
            problems.append(f"abi-baseline: unparsable baseline line: {line!r}")
            continue
        ret, name, params = match.groups()
        parsed = [] if params.strip() in ("", "void") else [p.strip() for p in params.split(",")]
        baseline[name] = (ret.strip(), parsed)

    ledger: dict[str, str] = {}
    for line in _parse_abi_manifest(C_ABI_BREAKS):
        match = _ABI_BREAK_RE.match(line)
        if match is None:
            problems.append(f"abi-baseline: unparsable ledger line: {line!r}")
            continue
        kind, name, _reason = match.groups()
        if name in ledger:
            problems.append(f"abi-baseline: {name} is listed twice in {C_ABI_BREAKS.name}")
        ledger[name] = kind

    current = c_abi_declarations(_read(CAPI_HEADER))

    for name, declared in sorted(baseline.items()):
        recorded = ledger.get(name)
        if name not in current:
            if recorded != "removed":
                problems.append(
                    f"abi-baseline: {name} shipped in the released header but is gone from "
                    f"{CAPI_HEADER.relative_to(REPO_ROOT)}. A consumer built against the release "
                    f"loses it at link time. Record the break as `removed {name} :: <reason>` in "
                    f"{C_ABI_BREAKS.relative_to(REPO_ROOT)} and carry it into the CHANGELOG's "
                    "Removed section, or restore the entry point."
                )
            continue
        if current[name] != declared:
            if recorded not in ("signature", "retyped"):
                problems.append(
                    f"abi-baseline: {name} changed signature since the release.\n"
                    f"  released: {declared[0]} {name}({', '.join(declared[1]) or 'void'})\n"
                    f"  current:  {current[name][0]} {name}({', '.join(current[name][1]) or 'void'})\n"
                    f"  Record it in {C_ABI_BREAKS.relative_to(REPO_ROOT)} as `signature` (breaks a "
                    "stale caller) or `retyped` (invisible to a C caller and to the calling "
                    "convention), with the reason."
                )
            continue
        # Declared and unchanged: any ledger entry naming it is stale.
        if recorded is not None:
            problems.append(
                f"abi-baseline: {C_ABI_BREAKS.relative_to(REPO_ROOT)} records {name} as "
                f"`{recorded}`, but it is declared unchanged from the release. Drop the line -- a "
                "ledger that outlives its break stops meaning anything."
            )

    for name in sorted(set(ledger) - set(baseline)):
        problems.append(
            f"abi-baseline: {C_ABI_BREAKS.relative_to(REPO_ROOT)} records {name}, which was never "
            f"in the released surface. Only an entry point a consumer could have compiled against "
            "can be broken; drop the line."
        )

    return problems


CHECKS = {
    "python-exports": check_python_exports,
    "python-struct-layouts": check_python_struct_layouts,
    "python-call-signatures": check_python_call_signatures,
    "python-inline-structs": check_python_inline_structs,
    "dts-wasm": check_dts_wasm,
    "dts-node": check_dts_node,
    "dts-shared-shapes": check_dts_shared_shapes,
    "pure-js-helpers": check_pure_js_helpers,
    "readme-counts": check_readme_counts,
    "dts-enums": check_dts_enums,
    "style-record-fields": check_style_record_fields,
    "staged-dist": check_staged_dist,
    "abi-baseline": check_abi_baseline,
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
