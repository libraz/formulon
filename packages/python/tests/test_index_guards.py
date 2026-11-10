"""Range guards on the integer arguments the binding hands to WebAssembly.

Two layers are covered:

  * A static sweep pairing every integer the package hands to the C ABI
    with the slot it lands in, asserting that each one is either a literal
    or is wrapped in ``_uint`` / ``_sint``. Three routes reach that ABI and
    all three are swept: a direct ``LIB.fm_*(...)`` call, a bare ``LIB.fm_*``
    reference handed to a helper that makes the call itself, and a field
    packed into a struct the ABI reads through a pointer. A sweep that
    followed only the first route passed while the other two carried
    unguarded coordinates, so the routes are enumerated in
    ``_INDIRECT_CALLERS`` and ``_struct_pack_slots`` rather than assumed.
  * Behavioural tests for the wrap that motivates the guard: ``wasmtime``
    marshals a Python int into an ``i32`` modulo 2**32, so ``2**32 + 5``
    would otherwise land on row 5 and overwrite a live cell. On the packed
    route the same value reaches ``struct.pack`` instead, which raises
    ``struct.error`` -- an exception a caller guarding with ``ValueError``
    never sees.
"""

from __future__ import annotations

import ast
import re
import unittest
from pathlib import Path

import formulon
from formulon import Workbook, _structs
from formulon._c import _sint, _uint

_PKG_ROOT = Path(__file__).resolve().parent.parent
_HEADER = _PKG_ROOT.parent.parent / "src" / "c_api" / "formulon_c.h"

# C parameter types passed by value in a 32-bit WASM ABI. Anything else in
# a signature is a pointer (an offset the binding computes itself) or a
# double, neither of which a caller-supplied index can corrupt.
_UNSIGNED_TYPES = {"uint32_t", "size_t", "uint16_t", "uint8_t"}
_SIGNED_TYPES = {"int32_t", "fm_status_t", "fm_error_code_t"}

# Helpers that take a bare ``LIB.fm_*`` reference and make the ABI call
# themselves. Each maps to `(index of the callable argument at the call
# site, index of the C parameter the argument after it lands on)`. A
# ``None`` second element means the helper supplies every remaining
# argument itself, so the call site passes no integer through -- the entry
# still exists so the reference is checked against the header.
#
# The handle argument tells the two shapes apart: ``_read_count`` is a
# module-level function whose caller passes the handle, so its forwarded
# arguments start at parameter 0, while the methods supply the handle
# themselves and their callers start at parameter 1.
_INDIRECT_CALLERS = {
    "_read_count": (0, 0),
    "_style_count": (0, None),
    "_add_xf_record": (1, None),
    "_pivot_set_field_order": (0, 1),
    "_trace": (0, 1),
}

# Struct field kinds that go through ``struct.pack`` as integers, and the
# width each one can hold. A value outside it raises ``struct.error``
# rather than the ``ValueError`` the binding documents, which is why the
# packed route needs the same guard as a scalar argument.
_PACKED_INT_KINDS = {"u32": 32, "i32": 32, "u16": 16, "u8": 8, "i64": 64}


def _parse_header(path: Path) -> dict:
    """Map each ``fm_*`` function to its ``[(type, name), ...]`` parameters."""
    text = path.read_text(encoding="utf-8")
    text = re.sub(r"/\*.*?\*/", " ", text, flags=re.S)
    text = re.sub(r"//[^\n]*", " ", text)
    functions = {}
    for decl in re.findall(r"FM_API\s+([^;]+);", text):
        decl = " ".join(decl.split())
        match = re.match(r"(.+?)\b(fm_\w+)\s*\((.*)\)$", decl)
        if match is None:
            continue
        _ret, name, params = match.groups()
        parsed = []
        if params.strip() not in ("void", ""):
            for param in params.split(","):
                bits = re.match(r"(.*?)(\w+)$", param.strip())
                parsed.append((bits.group(1).strip(), bits.group(2)))
        functions[name] = parsed
    return functions


def _is_literal_int(node: ast.AST) -> bool:
    if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.USub, ast.UAdd)):
        return _is_literal_int(node.operand)
    return isinstance(node, ast.Constant) and isinstance(node.value, int)


def _is_int_call(node: ast.AST) -> bool:
    """True for a bare ``int(...)``, which truncates a fraction and accepts a numeric string."""
    return isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id == "int"


# Methods whose result tuple is range-checked where it is built (verified by
# `ResolverGuardTests`), so the names a call site unpacks from one are guarded.
_GUARDED_RESOLVERS = {"_at", "_call"}

# Module-level encoders that only ever return literal integers, and the patch
# operation constants (`BoundedEncoderTests` verifies both).
_BOUNDED_ENCODERS = {"_tristate", "_tristate_default"}
_PATCH_OP_CONSTANTS = {"_PATCH_SET", "_PATCH_CLEAR"}


def _is_bool_check(node: ast.AST) -> bool:
    """True for a ``_require_bool(...)`` call, which returns its argument only when it is a ``bool``."""
    return isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id == "_require_bool"


def _is_guarded(node: ast.AST, known: "frozenset[str]" = frozenset()) -> bool:
    """True when ``node`` cannot reach WASM with an out-of-range value.

    ``known`` holds the local names every assignment of which is itself guarded.
    """
    if _is_literal_int(node):
        return True
    if isinstance(node, ast.Name):
        return node.id in known or node.id in _PATCH_OP_CONSTANTS
    if isinstance(node, ast.IfExp):
        return _is_guarded(node.body, known) and _is_guarded(node.orelse, known)
    if isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and f"{node.func.id}()" in known:
        return True
    if _is_int_call(node) and len(node.args) == 1 and _is_bool_check(node.args[0]):
        # `int(True)` is 1: the flag was verified to be a `bool` before it widened.
        return True
    return (
        isinstance(node, ast.Call)
        and isinstance(node.func, ast.Name)
        and node.func.id in ("_uint", "_sint", "_require_bool", *_BOUNDED_ENCODERS)
        and not (node.args and _is_int_call(node.args[0]))
    ) or (
        # `_ArgScope.cstr` hands back a pointer into the scope, not a caller's integer.
        isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute) and node.func.attr == "cstr"
    )


class _Facts:
    """What a function body establishes about its locals: guarded names and tuple-valued names."""

    def __init__(self, func: "ast.AST | None") -> None:
        self.known: "frozenset[str]" = frozenset()
        self.tuples: "dict[str, list]" = {}
        if func is None:
            return
        params = set()
        if isinstance(func, (ast.FunctionDef, ast.AsyncFunctionDef)):
            args = func.args
            params = {a.arg for a in (*args.posonlyargs, *args.args, *args.kwonlyargs)}
            params |= {a.arg for a in (args.vararg, args.kwarg) if a is not None}
        bad: set = set(params)
        assigned: "dict[str, list]" = {}
        resolver_names: set = set()
        # A nested helper whose every return is guarded makes its calls guarded.
        nested = set()
        for node in ast.walk(func):
            if isinstance(node, ast.FunctionDef) and node is not func:
                returns = [r.value for r in ast.walk(node) if isinstance(r, ast.Return) and r.value is not None]
                if returns and all(_is_guarded(v) for v in returns):
                    nested.add(f"{node.name}()")
        for node in ast.walk(func):
            if isinstance(node, ast.Assign):
                if len(node.targets) == 1 and isinstance(node.targets[0], ast.Name):
                    assigned.setdefault(node.targets[0].id, []).append(node.value)
                    continue
                target = node.targets[0] if len(node.targets) == 1 else None
                if (
                    isinstance(target, ast.Tuple)
                    and isinstance(node.value, ast.Tuple)
                    and len(target.elts) == len(node.value.elts)
                    and all(isinstance(e, ast.Name) for e in target.elts)
                ):
                    for elt, value in zip(target.elts, node.value.elts):
                        assigned.setdefault(elt.id, []).append(value)
                    continue
                target = node.targets[0] if len(node.targets) == 1 else None
                if (
                    isinstance(target, ast.Tuple)
                    and isinstance(node.value, ast.Call)
                    and _called_name(node.value) in _GUARDED_RESOLVERS
                ):
                    for elt in target.elts[1:]:
                        if isinstance(elt, ast.Name):
                            resolver_names.add(elt.id)
                    continue
                for tgt in node.targets:
                    bad.update(n.id for n in ast.walk(tgt) if isinstance(n, ast.Name))
            elif isinstance(node, (ast.AugAssign, ast.AnnAssign)):
                bad.update(n.id for n in ast.walk(node.target) if isinstance(n, ast.Name))
            elif isinstance(node, (ast.For, ast.AsyncFor, ast.comprehension)):
                bad.update(n.id for n in ast.walk(node.target) if isinstance(n, ast.Name))
            elif isinstance(node, ast.NamedExpr):
                bad.add(node.target.id)
            elif isinstance(node, ast.withitem) and node.optional_vars is not None:
                bad.update(n.id for n in ast.walk(node.optional_vars) if isinstance(n, ast.Name))
        known: set = {n for n in resolver_names if n not in bad and n not in assigned and n != "_"} | nested
        for _ in range(3):
            frozen = frozenset(known)
            known |= {
                n for n, values in assigned.items() if n not in bad and all(_is_guarded(v, frozen) for v in values)
            }
        self.known = frozenset(known)
        for name, values in assigned.items():
            if len(values) == 1 and isinstance(values[0], ast.Tuple) and name not in bad:
                self.tuples[name] = list(values[0].elts)


def _lib_fn_name(node: ast.AST) -> "str | None":
    """Return the ``fm_*`` name when ``node`` is a bare ``LIB.fm_*`` reference."""
    if (
        isinstance(node, ast.Attribute)
        and isinstance(node.value, ast.Name)
        and node.value.id == "LIB"
        and node.attr.startswith("fm_")
    ):
        return node.attr
    return None


def _called_name(node: ast.Call) -> "str | None":
    """Name of the function a call invokes, for a plain or ``self.`` call."""
    if isinstance(node.func, ast.Name):
        return node.func.id
    if isinstance(node.func, ast.Attribute):
        return node.func.attr
    return None


def _check_args(args, params, offset: int, where: str, findings: list, facts: "_Facts | None" = None) -> None:
    """Pair positional arguments against C parameters starting at ``offset``."""
    facts = facts or _Facts(None)
    expanded: list = []
    for arg in args:
        if isinstance(arg, ast.Starred) and isinstance(arg.value, ast.Name) and arg.value.id in facts.tuples:
            expanded.extend(facts.tuples[arg.value.id])
        else:
            expanded.append(arg)
    for position, arg in enumerate(expanded):
        index = offset + position
        if index >= len(params):
            findings.append(f"{where}: argument {position} has no parameter at index {index}")
            return
        ptype, pname = params[index]
        if ptype not in _UNSIGNED_TYPES and ptype not in _SIGNED_TYPES:
            continue
        if not _is_guarded(arg, facts.known):
            findings.append(f"{where}: {ptype} {pname} <- {ast.unparse(arg)}")


def _direct_call_slots(
    node: ast.Call, functions: dict, module_name: str, findings: list, facts: "_Facts | None" = None
) -> None:
    """Check a ``LIB.fm_*(...)`` call site against the header signature."""
    name = _lib_fn_name(node.func)
    if name is None:
        return
    params = functions.get(name)
    where = f"{module_name}:{node.lineno} {name}"
    if params is None:
        findings.append(f"{where}: not declared in formulon_c.h")
        return
    if len(node.args) > len(params):
        findings.append(f"{where}: passes {len(node.args)} args for {len(params)} parameters")
        return
    _check_args(node.args, params, 0, where, findings, facts)


def _indirect_call_slots(
    node: ast.Call, functions: dict, module_name: str, findings: list, facts: "_Facts | None" = None
) -> None:
    """Check a ``LIB.fm_*`` reference passed to a helper that calls it.

    The reference is not a call expression, so a sweep looking only for
    ``LIB.fm_*(...)`` walks straight past every argument handed to it.
    """
    helper = _called_name(node)
    route = _INDIRECT_CALLERS.get(helper)
    if route is None:
        return
    callable_index, first_param = route
    if callable_index >= len(node.args):
        return
    name = _lib_fn_name(node.args[callable_index])
    if name is None:
        # A helper can also be handed a local callable; only the ABI
        # references are this sweep's business.
        return
    params = functions.get(name)
    where = f"{module_name}:{node.lineno} {name} (via {helper})"
    if params is None:
        findings.append(f"{where}: not declared in formulon_c.h")
        return
    if first_param is None:
        return
    _check_args(node.args[callable_index + 1 :], params, first_param, where, findings, facts)


def _struct_layout(node: ast.AST):
    """Resolve ``S.<LAYOUT>`` to the ``Struct`` the binding packs with."""
    if not (isinstance(node, ast.Attribute) and isinstance(node.value, ast.Name) and node.value.id == "S"):
        return None
    return getattr(_structs, node.attr, None)


def _is_non_integer(node: ast.AST) -> bool:
    """True when ``node`` cannot be an integer, so no range guard applies.

    Used only where the destination field is unknown: a ``float`` lands in
    an ``f64`` slot, and ``struct.pack`` would reject it outright from an
    integer one.
    """
    if isinstance(node, ast.Constant) and isinstance(node.value, float):
        return True
    if (
        isinstance(node, ast.Call)
        and isinstance(node.func, ast.Name)
        and node.func.id in ("float", "_finite", "_require_number")
    ):
        # `_finite` and `_require_number` check and return a float.
        return True
    if isinstance(node, ast.IfExp):
        return _is_non_integer(node.body) and _is_non_integer(node.orelse)
    return False


def _returned_dict(func: ast.FunctionDef) -> "ast.Dict | None":
    """The dict literal a one-expression field-projection helper returns."""
    body = [
        stmt
        for stmt in func.body
        if not (isinstance(stmt, ast.Expr) and isinstance(stmt.value, ast.Constant))
        and not isinstance(stmt, ast.FunctionDef)
        and not (isinstance(stmt, ast.If) and all(isinstance(s, ast.Raise) for s in stmt.body) and not stmt.orelse)
    ]
    if len(body) == 1 and isinstance(body[0], ast.Return) and isinstance(body[0].value, ast.Dict):
        return body[0].value
    return None


def _returned_local_dict(func: ast.FunctionDef) -> "list | None":
    """The fields written into the one local dict a helper builds and returns, or ``None`` for any other shape."""
    returns = [n for n in ast.walk(func) if isinstance(n, ast.Return)]
    if len(returns) != 1 or not isinstance(returns[0].value, ast.Name):
        return None
    return _collected_payload(returns[0].value.id, func)


def _collected_payload(name: str, scope: ast.AST) -> "list | None":
    """Every ``(key_node_or_None, value)`` written into a local dict ``name``.

    Covers the build-then-pack shape: a ``values = {...}`` seed plus
    ``values[key] = expr`` writes, including the loops that derive a field
    name from a variable. A dynamic key yields ``None`` so the caller
    range-checks the value without knowing which field it lands in.
    """
    writes: list = []
    seen = False
    for node in ast.walk(scope):
        if isinstance(node, ast.Assign) and len(node.targets) == 1:
            target = node.targets[0]
            if isinstance(target, ast.Name) and target.id == name:
                seen = True
                if isinstance(node.value, ast.Dict):
                    writes.extend(zip(node.value.keys, node.value.values))
                continue
            if isinstance(target, ast.Subscript) and isinstance(target.value, ast.Name) and target.value.id == name:
                seen = True
                writes.append((target.slice, node.value))
        elif isinstance(node, ast.AnnAssign) and isinstance(node.target, ast.Name) and node.target.id == name:
            seen = True
            if isinstance(node.value, ast.Dict):
                writes.extend(zip(node.value.keys, node.value.values))
    return writes if seen else None


def _check_packed_field(layout, key, value, where: str, findings: list, known: "frozenset[str]" = frozenset()) -> None:
    """Range-check one field written into a packed struct."""
    if isinstance(key, ast.Constant) and isinstance(key.value, str):
        entry = layout.offsets.get(key.value)
        if entry is None:
            findings.append(f"{where}: no field named {key.value!r}")
            return
        kind = entry[0]
        if kind not in _PACKED_INT_KINDS:
            return
        if not _is_guarded(value, known):
            findings.append(f"{where}: {kind} {key.value} <- {ast.unparse(value)}")
        return
    # Field name computed at runtime: the value has to be safe for
    # whichever integer field it reaches.
    if not (_is_guarded(value, known) or _is_non_integer(value)):
        findings.append(f"{where}: {ast.unparse(key)} <- {ast.unparse(value)}")


def _struct_pack_slots(
    node: ast.Call, scope, helpers: dict, module_name: str, findings: list, facts: "_Facts | None" = None
) -> None:
    """Check the integer fields written by ``S.<LAYOUT>.pack(LIB, ptr, ...)``.

    The ABI reads these through a pointer, so the header signature says
    only ``fm_hyperlink_t*`` and the scalar sweep sees nothing. The value
    still crosses into WASM, and out of range it dies inside
    ``struct.pack`` with an exception the API never documents.
    """
    if not (isinstance(node.func, ast.Attribute) and node.func.attr == "pack"):
        return
    layout = _struct_layout(node.func.value)
    if layout is None or len(node.args) < 3:
        return
    where = f"{module_name}:{node.lineno} {layout.name}"
    payload = node.args[2]

    pairs = None
    known = (facts or _Facts(None)).known
    if isinstance(payload, ast.Dict):
        pairs = list(zip(payload.keys, payload.values))
    elif isinstance(payload, ast.Call) and isinstance(payload.func, ast.Name):
        helper = helpers.get(payload.func.id)
        returned = _returned_dict(helper) if helper is not None else None
        if returned is not None:
            pairs = list(zip(returned.keys, returned.values))
            known = _Facts(helper).known
        elif helper is not None and (built := _returned_local_dict(helper)) is not None:
            pairs = built
            known = _Facts(helper).known
    elif isinstance(payload, ast.Name) and scope is not None:
        pairs = _collected_payload(payload.id, scope)

    if pairs is None:
        # Report rather than skip: a payload the sweep cannot read is a
        # hole in it, not evidence that the site is safe.
        findings.append(f"{where}: packs a payload this sweep cannot inspect ({ast.unparse(payload)})")
        return
    for key, value in pairs:
        _check_packed_field(layout, key, value, where, findings, known)


# Encoders whose input a constructor has already validated, so the fields
# they pack are checked even though no `_uint` / `_sint` is in sight.
_VALIDATED_ENCODERS = {
    # `Value.__post_init__` coerces `kind` and type-checks the one payload it takes.
    "_write_value_input",
}


def _raw_struct_pack_slots(node: ast.Call, scope, module_name: str, findings: list) -> None:
    """Check the values handed to a bare ``struct.pack(fmt, ...)``.

    A pointer, a length or a local the binding computed is safe; a caller's
    field read straight off an object, or anything coerced with ``int()``,
    has not been range-checked and must go through ``_uint`` / ``_sint``.
    """
    if ast.unparse(node.func) != "struct.pack" or len(node.args) < 2:
        return
    if scope is not None and scope.name in _VALIDATED_ENCODERS:
        return
    for arg in node.args[1:]:
        value = arg.value if isinstance(arg, ast.Starred) else arg
        if _is_guarded(value):
            continue
        if isinstance(value, ast.Attribute) or any(_is_int_call(n) for n in ast.walk(value)):
            findings.append(f"{module_name}:{node.lineno} struct.pack <- {ast.unparse(value)}")


def _scopes(tree: ast.AST) -> dict:
    """Map each node to the innermost function that contains it."""
    enclosing = {}
    for func in ast.walk(tree):
        if not isinstance(func, (ast.FunctionDef, ast.AsyncFunctionDef)):
            continue
        for node in ast.walk(func):
            enclosing[node] = func
    return enclosing


def _package_helpers() -> dict:
    """The module-level functions of every module of the package, for payloads built in a sibling module."""
    out: dict = {}
    for module in sorted(_PKG_ROOT.joinpath("formulon").glob("*.py")):
        for node in ast.parse(module.read_text(encoding="utf-8")).body:
            if isinstance(node, ast.FunctionDef):
                out.setdefault(node.name, node)
    return out


def _sweep_source(source: str, functions: dict, module_name: str = "<probe>", shared: "dict | None" = None) -> list:
    """Run the three checkers over one module's source."""
    tree = ast.parse(source)
    helpers = {**(shared or {}), **{n.name: n for n in tree.body if isinstance(n, ast.FunctionDef)}}
    enclosing = _scopes(tree)
    facts_of: dict = {}
    findings: list = []
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        scope = enclosing.get(node)
        if id(scope) not in facts_of:
            facts_of[id(scope)] = _Facts(scope)
        facts = facts_of[id(scope)]
        _direct_call_slots(node, functions, module_name, findings, facts)
        _indirect_call_slots(node, functions, module_name, findings, facts)
        _struct_pack_slots(node, scope, helpers, module_name, findings, facts)
        _raw_struct_pack_slots(node, scope, module_name, findings)
    return findings


def _unguarded_slots() -> list:
    """Return every unguarded integer argument slot in the package."""
    functions = _parse_header(_HEADER)
    shared = _package_helpers()
    findings: list = []
    for module in sorted(_PKG_ROOT.joinpath("formulon").glob("*.py")):
        findings.extend(_sweep_source(module.read_text(encoding="utf-8"), functions, module.name, shared))
    return findings


class ArgumentSweepTests(unittest.TestCase):
    """Static pairing of the Python call sites with the C ABI signatures."""

    def setUp(self) -> None:
        if not _HEADER.is_file():
            self.skipTest(f"C ABI header not reachable at {_HEADER}")

    def test_every_integer_argument_is_range_checked(self) -> None:
        findings = _unguarded_slots()
        self.assertEqual(
            findings,
            [],
            "integer arguments reach WASM without a _uint/_sint range check:\n  " + "\n  ".join(findings),
        )

    def test_bounded_encoders_return_only_literal_integers(self) -> None:
        helpers = _package_helpers()
        for name in sorted(_BOUNDED_ENCODERS):
            returns = [r.value for r in ast.walk(helpers[name]) if isinstance(r, ast.Return) and r.value is not None]
            self.assertTrue(returns, name)
            for value in returns:
                arms = [value.body, value.orelse] if isinstance(value, ast.IfExp) else [value]
                self.assertTrue(all(_is_literal_int(v) for v in arms), f"{name} returns {ast.unparse(value)}")
        source = (_PKG_ROOT / "formulon" / "_workbook.py").read_text(encoding="utf-8")
        constants = {
            t.id: n.value.value
            for n in ast.parse(source).body
            if isinstance(n, ast.Assign) and isinstance(n.value, ast.Constant) and isinstance(n.value.value, int)
            for t in n.targets
            if isinstance(t, ast.Name)
        }
        for name in _PATCH_OP_CONSTANTS:
            self.assertIn(name, constants)

    def test_sweep_actually_inspects_call_sites(self) -> None:
        # Guards the guard: a header-parsing or AST-walking regression would
        # otherwise make the sweep above vacuously pass.
        functions = _parse_header(_HEADER)
        self.assertIn("fm_sheet_set_number", functions)
        self.assertEqual(
            [t for t, _n in functions["fm_sheet_set_number"]],
            ["fm_workbook_t*", "size_t", "uint32_t", "uint32_t", "double"],
        )

    def test_sweep_catches_a_direct_call(self) -> None:
        findings = _sweep_source("LIB.fm_sheet_set_number(h, int(sheet), 0, 0, 1.0)", _parse_header(_HEADER))
        self.assertEqual(len(findings), 1, findings)
        self.assertIn("sheet_index <- int(sheet)", findings[0])

    def test_sweep_catches_a_reference_handed_to_a_helper(self) -> None:
        # The route the previous sweep missed: the ABI function is named
        # but never called here, so there is no `LIB.fm_*(...)` expression
        # for a call-shaped matcher to find.
        findings = _sweep_source(
            "_read_count(LIB.fm_sheet_get_merge_count, h, int(sheet))",
            _parse_header(_HEADER),
        )
        self.assertEqual(len(findings), 1, findings)
        self.assertIn("via _read_count", findings[0])
        self.assertIn("sheet_index <- int(sheet)", findings[0])

    def test_sweep_catches_a_packed_struct_field(self) -> None:
        findings = _sweep_source('S.COMMENT.pack(LIB, ptr, {"row": int(row)})', _parse_header(_HEADER))
        self.assertEqual(len(findings), 1, findings)
        self.assertIn("u32 row <- int(row)", findings[0])

    def test_sweep_catches_a_field_packed_through_a_local_dict(self) -> None:
        findings = _sweep_source(
            "def f():\n"
            "    values = {}\n"
            '    values["spin_count"] = int(spin_count)\n'
            "    S.SHEET_PROTECTION.pack(LIB, ptr, values)\n",
            _parse_header(_HEADER),
        )
        self.assertEqual(len(findings), 1, findings)
        self.assertIn("u32 spin_count", findings[0])

    def test_sweep_catches_a_64_bit_field_coerced_with_int(self) -> None:
        findings = _sweep_source(
            'S.DRAWING_ANCHOR.pack(LIB, ptr, {"width_emu": int(width_emu)})', _parse_header(_HEADER)
        )
        self.assertEqual(len(findings), 1, findings)
        self.assertIn("i64 width_emu <- int(width_emu)", findings[0])

    def test_sweep_follows_a_helper_that_builds_and_returns_a_local_dict(self) -> None:
        template = (
            "def fields(scope, spec):\n"
            "    values = {}\n"
            '    values["width_emu"] = %s\n'
            "    return values\n"
            "def add(spec):\n"
            "    S.DRAWING_ANCHOR.pack(LIB, ptr, fields(scope, spec))\n"
        )
        findings = _sweep_source(template % "int(spec.width)", _parse_header(_HEADER))
        self.assertEqual(len(findings), 1, findings)
        self.assertIn("i64 width_emu <- int(spec.width)", findings[0])
        self.assertEqual(_sweep_source(template % '_sint(spec.width, "width_emu", 64)', _parse_header(_HEADER)), [])

    def test_sweep_catches_int_inside_a_guard(self) -> None:
        findings = _sweep_source('S.COMMENT.pack(LIB, ptr, {"row": _uint(int(row), "row")})', _parse_header(_HEADER))
        self.assertEqual(len(findings), 1, findings)

    def test_sweep_catches_a_raw_struct_pack_of_a_caller_value(self) -> None:
        functions = _parse_header(_HEADER)
        for source in (
            'struct.pack("<B", int(bar.direction))',
            'struct.pack("<B", bar.direction)',
            'b"".join(struct.pack("<I", int(i)) for i in indices)',
        ):
            with self.subTest(source=source):
                self.assertEqual(len(_sweep_source(source, functions)), 1)
        for source in (
            'struct.pack("<B", _uint(bar.direction, "direction", 8))',
            'struct.pack("<I", text_ptr)',
            'struct.pack("<I", len(values))',
        ):
            with self.subTest(source=source):
                self.assertEqual(_sweep_source(source, functions), [])

    def test_sweep_reports_a_payload_it_cannot_read(self) -> None:
        # An unreadable payload is a hole in the sweep, not evidence that
        # the site is safe, so it has to surface rather than be skipped.
        findings = _sweep_source("S.COMMENT.pack(LIB, ptr, build_it())", _parse_header(_HEADER))
        self.assertEqual(len(findings), 1, findings)
        self.assertIn("cannot inspect", findings[0])

    def test_sweep_accepts_the_guarded_forms(self) -> None:
        functions = _parse_header(_HEADER)
        for source in (
            'LIB.fm_sheet_set_number(h, _uint(sheet, "sheet_index"), 0, 0, 1.0)',
            '_read_count(LIB.fm_sheet_get_merge_count, h, _uint(sheet, "sheet"))',
            'S.COMMENT.pack(LIB, ptr, {"row": _uint(row, "row")})',
            "def f(name, value):\n    values = {}\n    values[name] = 0.0 if value is None else _finite(value, name)\n    S.COMMENT.pack(LIB, ptr, values)\n",
        ):
            with self.subTest(source=source):
                self.assertEqual(_sweep_source(source, functions), [])


class ResolverGuardTests(unittest.TestCase):
    """The helpers the sweep trusts to return range-checked values do."""

    def test_guarded_resolvers_return_range_checked_values(self) -> None:
        seen = set()
        for module in sorted(_PKG_ROOT.joinpath("formulon").glob("*.py")):
            for func in ast.walk(ast.parse(module.read_text(encoding="utf-8"))):
                if not (isinstance(func, ast.FunctionDef) and func.name in _GUARDED_RESOLVERS):
                    continue
                returns = [r.value for r in ast.walk(func) if isinstance(r, ast.Return)]
                self.assertTrue(returns, f"{module.name}:{func.name} returns nothing")
                for value in returns:
                    self.assertIsInstance(value, ast.Tuple, f"{module.name}:{func.name}")
                    for element in value.elts[1:]:
                        self.assertTrue(
                            _is_guarded(element), f"{module.name}:{func.name} returns {ast.unparse(element)}"
                        )
                seen.add(func.name)
        self.assertEqual(seen, _GUARDED_RESOLVERS)

    def test_sweep_trusts_only_the_resolvers(self) -> None:
        functions = _parse_header(_HEADER)
        trusted = "def f(ws):\n    h, sheet, oid = ws._at()\n    LIB.fm_sheet_set_drawing_z_order(h, sheet, oid, 0)\n"
        self.assertEqual(_sweep_source(trusted, functions), [])
        untrusted = trusted.replace("_at", "_other")
        self.assertEqual(len(_sweep_source(untrusted, functions)), 2)

    def test_sweep_checks_a_local_assigned_from_a_guard(self) -> None:
        functions = _parse_header(_HEADER)
        guarded = 'def f(sheet):\n    s = _uint(sheet, "s")\n    LIB.fm_sheet_set_number(h, s, 0, 0, 1.0)\n'
        self.assertEqual(_sweep_source(guarded, functions), [])
        reassigned = guarded + "    s = sheet\n    LIB.fm_sheet_set_number(h, s, 0, 0, 1.0)\n"
        self.assertEqual(len(_sweep_source(reassigned, functions)), 2)
        parameter = "def f(s):\n    LIB.fm_sheet_set_number(h, s, 0, 0, 1.0)\n"
        self.assertEqual(len(_sweep_source(parameter, functions)), 1)

    def test_sweep_expands_a_starred_tuple_of_guards(self) -> None:
        functions = _parse_header(_HEADER)
        lead = 'def f(a, b):\n    lead = (_uint(a, "a"), 0)\n    LIB.fm_sheet_set_number(h, *lead, 0, 1.0)\n'
        self.assertEqual(_sweep_source(lead, functions), [])
        raw = "def f(a, b):\n    lead = (a, 0)\n    LIB.fm_sheet_set_number(h, *lead, 0, 1.0)\n"
        self.assertEqual(len(_sweep_source(raw, functions)), 1)


class IndexRangeTests(unittest.TestCase):
    """Out-of-range indices are rejected before any WASM call happens."""

    def test_fractional_integer_arguments_are_rejected_without_truncation(self) -> None:
        for guard in (_uint, _sint):
            with self.subTest(guard=guard.__name__):
                for not_an_integer in (1.5, 1.0, float("nan"), True, None, "5"):
                    with self.assertRaises(TypeError):
                        guard(not_an_integer, "integer")

        with Workbook.create_default() as wb:
            with self.assertRaises(TypeError):
                wb.sheet(0).get_value(1.5, 0)
            with self.assertRaises(TypeError):
                wb.sheet(0).get_value(1.0, 0)
            with self.assertRaises(TypeError):
                wb.sheet(0).set_hyperlink(1.5, 0, formulon.HyperlinkInput(target="https://example.test"))
            wb.update_iterative(enabled=True, max_iterations=100, max_change=0.001)

    def test_wrapping_row_does_not_overwrite_the_aliased_cell(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_number(5, 0, 1.0)
            with self.assertRaises(ValueError):
                wb.sheet(0).set_number(2**32 + 5, 0, 99.0)
            self.assertEqual(wb.sheet(0).get_typed_value(5, 0).number, 1.0)

    def test_wrapping_sheet_index_is_rejected(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_number(0, 0, 7.0)
            with self.assertRaises(IndexError):
                wb.sheet(2**32).set_number(0, 0, 99.0)
            with self.assertRaises(IndexError):
                wb.sheet(2**32).get_value(0, 0)
            self.assertEqual(wb.sheet(0).get_typed_value(0, 0).number, 7.0)

    def test_negative_indices_are_rejected(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_number(0, 0, 3.0)
            for row, col in ((-1, 0), (0, -1)):
                with self.assertRaises(formulon.InvalidArgumentError) as ctx:
                    wb.sheet(0).set_number(row, col, 99.0)
                self.assertEqual(ctx.exception.code, formulon.StatusCode.INVALID_ARGUMENT)
            with self.assertRaises(formulon.IndexOutOfRangeError) as ctx:
                wb.sheet(-1)
            self.assertEqual(ctx.exception.code, formulon.StatusCode.SHEET_INDEX_OUT_OF_RANGE)
            self.assertEqual(wb.sheet(0).get_typed_value(0, 0).number, 3.0)

    def test_huge_indices_are_rejected(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(formulon.InvalidArgumentError):
                wb.sheet(0).set_text(2**64, 0, "wrapped")
            with self.assertRaises(formulon.InvalidArgumentError):
                wb.sheet(0).set_formula(0, 2**48, "=1")
            with self.assertRaises(IndexError):
                wb.sheet(2**32 + 0).name
            with self.assertRaises(ValueError):
                wb.sheet(2**64)

    def test_error_message_names_the_parameter(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(formulon.InvalidArgumentError) as ctx:
                wb.sheet(0).set_number(2**32 + 5, 0, 1.0)
            self.assertEqual(ctx.exception.code, formulon.StatusCode.INVALID_ARGUMENT)
            self.assertIn("row", ctx.exception.context)

    def test_a_wrapping_sheet_index_finds_no_handle(self) -> None:
        # A sheet index that would wrap modulo 2**32 never reaches the ABI:
        # the handle lookup rejects it before any per-sheet call.
        with Workbook.create_default() as wb:
            with self.assertRaises(IndexError):
                wb.sheet(2**32)

    def test_trace_walk_rejects_a_wrapping_index(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(ValueError):
                wb.sheet(0).get_precedents(2**32 + 5, 0, 1)
            with self.assertRaises(ValueError):
                wb.sheet(0).get_dependents(0, 0, 2**32)

    def test_packed_coordinates_raise_value_error_naming_the_parameter(self) -> None:
        # A coordinate packed into a struct never appears as a scalar
        # argument, so it used to die inside ``struct.pack`` with a
        # ``struct.error`` that no caller guarding on ``ValueError`` sees.
        with Workbook.create_default() as wb:
            with self.assertRaises(ValueError) as ctx:
                wb.sheet(0).set_hyperlink(2**32 + 5, 0, formulon.HyperlinkInput(target="http://example.invalid"))
            self.assertIn("row", ctx.exception.context)

            with self.assertRaises(formulon.InvalidArgumentError) as ctx:
                wb.sheet(0).add_merge(formulon.CellRect(first_row=-1, first_col=0, last_row=1, last_col=1))
            self.assertEqual(ctx.exception.context, "first_row=-1")

    def test_narrow_struct_fields_are_checked_at_their_own_width(self) -> None:
        # ``u8`` / ``u16`` fields overflow long before uint32 does, and
        # ``struct.pack`` is the only thing that used to notice.
        with Workbook.create_default() as wb:
            with self.assertRaises(ValueError) as ctx:
                wb.ooxml.add_font(formulon.Font(name="Calibri", size_pt=11.0, underline=256))
            self.assertIn("underline", str(ctx.exception))

    def test_in_range_extremes_still_reach_the_engine(self) -> None:
        # The guard rejects only what cannot survive marshalling; a value
        # inside uint32 must still be handled by the engine's own bounds.
        with Workbook.create_default() as wb:
            with self.assertRaises(formulon.FormulonError):
                wb.sheet(0).set_number(0xFFFFFFFF, 0, 1.0)


if __name__ == "__main__":
    unittest.main()
