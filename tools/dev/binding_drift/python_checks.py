"""Drift checks between the C ABI header and the Python binding."""

from __future__ import annotations

import ast
import importlib.util
import re
from pathlib import Path
from typing import List, Optional, Set

from .c_header import _c_function_params, _LayoutError, _wasm32_layout, _Wasm32Layouts
from .surface_files import (
    _SKIPPED,
    CAPI_EXPORTS,
    CAPI_HEADER,
    PYTHON_C_BINDING,
    PYTHON_PKG_DIR,
    PYTHON_STRUCTS,
    PYTHON_WORKBOOK,
    REPO_ROOT,
    _read,
)


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


# Scratch readers whose name pins how many bytes come back out of the slot.
# `read_cstr` is deliberately absent: it walks to a NUL rather than reading a
# fixed width, so it is checked for the deref mistake only (below).
_SCRATCH_READER_WIDTHS = {"read_u32": 4, "read_i32": 4, "read_f64": 8}


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
        # `x = _alloc_out_u64()` -- the 8-byte out-uint64 slot.
        elif isinstance(callee, ast.Name) and callee.id == "_alloc_out_u64":
            record(target.id, 8)
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
