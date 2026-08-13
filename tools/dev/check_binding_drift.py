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
                  that list is declared in src/c_api/formulon_c.h.
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
                  source enum's ordinal sequence.
  style-record-fields
                  Every public type that projects a style record
                  (`ColorSpec` / `FontRecord` / `FillRecord` /
                  `BorderSide`) carries the same field set as its C ABI
                  struct, across the WASM `.d.ts`, the Node `.d.ts` and
                  the Python dataclasses. Deliberate omissions live in
                  `_STYLE_RECORD_EXEMPT_TYPES`.
  all             Run every check above (default).

Stdlib only; no build artifacts or network access required.
"""

from __future__ import annotations

import argparse
import importlib.util
import re
import sys
from pathlib import Path
from typing import List, Optional, Set

REPO_ROOT = Path(__file__).resolve().parents[2]

CAPI_HEADER = REPO_ROOT / "src" / "c_api" / "formulon_c.h"
CAPI_EXPORTS = REPO_ROOT / "tools" / "wasm" / "capi_exports.txt"
PYTHON_PKG_DIR = REPO_ROOT / "packages" / "python" / "formulon"

WASM_BINDINGS_CPP = REPO_ROOT / "src" / "wasm" / "parts" / "bindings_register.cpp"
WASM_DTS = REPO_ROOT / "src" / "wasm" / "formulon.d.ts"

NODE_WORKBOOK_CLASS_CC = REPO_ROOT / "src" / "node_addon" / "parts" / "workbook_class.cc"
NODE_ADDON_CC = REPO_ROOT / "src" / "node_addon" / "addon.cc"
NODE_DTS = REPO_ROOT / "packages" / "npm-native" / "index.d.ts"
NODE_README = REPO_ROOT / "packages" / "npm-native" / "README.md"

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


def _read(path: Path) -> str:
    if not path.is_file():
        print(f"check_binding_drift: missing file: {path}", file=sys.stderr)
        sys.exit(2)
    return path.read_text(encoding="utf-8")


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

    header_fns = set(re.findall(r"\bfm_[A-Za-z0-9_]+(?=\s*\()", _read(CAPI_HEADER)))

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

    return problems


# ---------------------------------------------------------------------------
# Check 1b: Python WASM32 POD layouts <-> C ABI header.
# ---------------------------------------------------------------------------


def _wasm32_layout(fields: List[tuple[str, int, int]]) -> tuple[dict[str, int], int]:
    offsets: dict[str, int] = {}
    offset = 0
    max_align = 1
    for name, size, align in fields:
        offset = (offset + align - 1) // align * align
        offsets[name] = offset
        offset += size
        max_align = max(max_align, align)
    return offsets, (offset + max_align - 1) // max_align * max_align


def check_python_struct_layouts() -> List[str]:
    """Verify Python's hand-written WASM32 structs against the C header.

    This intentionally parses the authoritative C declarations rather than
    repeating sizes in a Python table. It covers every ``Struct`` exported by
    ``_structs.py`` and reports field order, offsets, and final size.
    """
    header = re.sub(r"/\*.*?\*/", "", _read(CAPI_HEADER), flags=re.S)
    blocks = {
        name: body for body, name in re.findall(r"typedef\s+struct\s*\{(.*?)\}\s*(fm_[A-Za-z0-9_]+)\s*;", header, re.S)
    }
    spec = importlib.util.spec_from_file_location("formulon_struct_layouts", PYTHON_STRUCTS)
    if spec is None or spec.loader is None:
        return ["python-struct-layout: could not load packages/python/formulon/_structs.py"]
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)

    primitive = {
        "uint8_t": (1, 1),
        "uint16_t": (2, 2),
        "uint32_t": (4, 4),
        "int32_t": (4, 4),
        "double": (8, 8),
        "size_t": (4, 4),
        "fm_value_kind_t": (4, 4),
        "fm_function_availability_t": (4, 4),
        "fm_cf_match_kind_t": (4, 4),
        "fm_pivot_cell_kind_t": (4, 4),
        "fm_pivot_axis_t": (4, 4),
        "fm_pivot_aggregation_t": (4, 4),
        "fm_pivot_show_as_t": (4, 4),
        "fm_pivot_filter_type_t": (4, 4),
        "fm_pivot_filter_value_kind_t": (4, 4),
        "fm_calc_mode_t": (4, 4),
        "fm_value_t": (16, 8),
        "fm_cf_color_t": (4, 1),
        "fm_cfvo_t": (12, 4),
        "fm_color_spec": (24, 8),
        "fm_border_side": (32, 8),
        "fm_font_record": (80, 8),
        "fm_fill_record": (64, 8),
        "fm_border_record": (168, 8),
    }
    # Versioned style records embed the stable ``fm_cell_xf`` prefix rather
    # than repeating its fields. Expand that nested POD here so the Python
    # marshaller can keep its flat field map while still checking the exact
    # C offsets and final size.
    embedded_cell_xf = [
        ("font_index", 4, 4),
        ("fill_index", 4, 4),
        ("border_index", 4, 4),
        ("num_fmt_id", 2, 2),
        ("horizontal_align", 1, 1),
        ("vertical_align", 1, 1),
        ("wrap_text", 4, 4),
    ]
    problems: List[str] = []
    for layout in (value for value in vars(module).values() if isinstance(value, module.Struct)):
        body = blocks.get(layout.name)
        if body is None:
            problems.append(f"python-struct-layout: {layout.name} missing from C header")
            continue
        c_fields: List[tuple[str, int, int]] = []
        py_field_names = {name for name, _ in layout.fields}
        for declaration in body.split(";"):
            declaration = " ".join(declaration.split())
            if not declaration:
                continue
            match = re.fullmatch(r"(.+?)\s+([A-Za-z_][A-Za-z0-9_]*)(?:\[(\d+)\])?", declaration)
            if not match:
                problems.append(f"python-struct-layout: cannot parse {layout.name}: {declaration!r}")
                continue
            ctype, name, count_text = match.groups()
            count = int(count_text or "1")
            if "*" in ctype:
                size, align = 4, 4
            else:
                ctype = ctype.replace("const ", "").strip()
                if ctype == "fm_cell_xf":
                    c_fields.extend(embedded_cell_xf)
                    continue
                if ctype == "fm_cf_color_t":
                    if name in py_field_names:
                        c_fields.append((name, 4, 1))
                    else:
                        c_fields.extend((f"{name}_{channel}", 1, 1) for channel in ("r", "g", "b", "a"))
                    continue
                if ctype not in primitive:
                    problems.append(f"python-struct-layout: unknown {layout.name} field type {ctype!r}")
                    continue
                size, align = primitive[ctype]
            # Explicit C padding need not be represented in Struct: its
            # alignment effect is reproduced by the following semantic field.
            if not name.startswith("_pad"):
                c_fields.append((name, size * count, align))
            else:
                c_fields.append((name, size * count, align))
        c_offsets, c_size = _wasm32_layout(c_fields)
        semantic_names = [name for name, _, _ in c_fields if not name.startswith("_pad")]
        py_names = [name for name, _ in layout.fields if not name.startswith("_pad")]
        if py_names != semantic_names:
            problems.append(f"python-struct-layout: {layout.name} fields differ: C={semantic_names}, Python={py_names}")
        for name in py_names:
            if name in c_offsets and layout.offsets[name][1] != c_offsets[name]:
                problems.append(
                    f"python-struct-layout: {layout.name}.{name} offset C={c_offsets[name]} Python={layout.offsets[name][1]}"
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

_TS_FIELD_RE = re.compile(r"^\s*([A-Za-z_$][A-Za-z0-9_$]*)\??\s*:", re.MULTILINE)
_PY_FIELD_RE = re.compile(r"^    ([a-z_][A-Za-z0-9_]*)\s*:", re.MULTILINE)


def _snake_to_camel(name: str) -> str:
    head, *rest = name.split("_")
    return head + "".join(part.capitalize() for part in rest)


def _c_struct_fields(header: str, struct_name: str, source: Path) -> List[str]:
    blocks = {
        name: body for body, name in re.findall(r"typedef\s+struct\s*\{(.*?)\}\s*(fm_[A-Za-z0-9_]+)\s*;", header, re.S)
    }
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


CHECKS = {
    "python-exports": check_python_exports,
    "python-struct-layouts": check_python_struct_layouts,
    "dts-wasm": check_dts_wasm,
    "dts-node": check_dts_node,
    "readme-counts": check_readme_counts,
    "dts-enums": check_dts_enums,
    "style-record-fields": check_style_record_fields,
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
    problems: List[str] = []
    for name in names:
        problems.extend(CHECKS[name]())

    if problems:
        print(f"check_binding_drift ({args.check}): DRIFT DETECTED")
        for problem in problems:
            print(problem)
        return 1

    print(f"check_binding_drift ({args.check}): no drift detected")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
