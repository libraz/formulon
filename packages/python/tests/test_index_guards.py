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
_PACKED_INT_KINDS = {"u32": 32, "i32": 32, "u16": 16, "u8": 8}


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
    return isinstance(node, ast.Constant) and isinstance(node.value, int)


def _is_guarded(node: ast.AST) -> bool:
    """True when ``node`` cannot reach WASM with an out-of-range value."""
    if _is_literal_int(node):
        return True
    if isinstance(node, ast.IfExp):
        return _is_literal_int(node.body) and _is_literal_int(node.orelse)
    return isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id in ("_uint", "_sint")


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


def _check_args(args, params, offset: int, where: str, findings: list) -> None:
    """Pair positional arguments against C parameters starting at ``offset``."""
    for position, arg in enumerate(args):
        index = offset + position
        if index >= len(params):
            findings.append(f"{where}: argument {position} has no parameter at index {index}")
            return
        ptype, pname = params[index]
        if ptype not in _UNSIGNED_TYPES and ptype not in _SIGNED_TYPES:
            continue
        if not _is_guarded(arg):
            findings.append(f"{where}: {ptype} {pname} <- {ast.unparse(arg)}")


def _direct_call_slots(node: ast.Call, functions: dict, module_name: str, findings: list) -> None:
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
    _check_args(node.args, params, 0, where, findings)


def _indirect_call_slots(node: ast.Call, functions: dict, module_name: str, findings: list) -> None:
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
    _check_args(node.args[callable_index + 1 :], params, first_param, where, findings)


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
    if isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id == "float":
        return True
    if isinstance(node, ast.IfExp):
        return _is_non_integer(node.body) and _is_non_integer(node.orelse)
    return False


def _returned_dict(func: ast.FunctionDef) -> "ast.Dict | None":
    """The dict literal a one-expression field-projection helper returns."""
    body = [stmt for stmt in func.body if not (isinstance(stmt, ast.Expr) and isinstance(stmt.value, ast.Constant))]
    if len(body) == 1 and isinstance(body[0], ast.Return) and isinstance(body[0].value, ast.Dict):
        return body[0].value
    return None


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


def _check_packed_field(layout, key, value, where: str, findings: list) -> None:
    """Range-check one field written into a packed struct."""
    if isinstance(key, ast.Constant) and isinstance(key.value, str):
        entry = layout.offsets.get(key.value)
        if entry is None:
            findings.append(f"{where}: no field named {key.value!r}")
            return
        kind = entry[0]
        if kind not in _PACKED_INT_KINDS:
            return
        if not _is_guarded(value):
            findings.append(f"{where}: {kind} {key.value} <- {ast.unparse(value)}")
        return
    # Field name computed at runtime: the value has to be safe for
    # whichever integer field it reaches.
    if not (_is_guarded(value) or _is_non_integer(value)):
        findings.append(f"{where}: {ast.unparse(key)} <- {ast.unparse(value)}")


def _struct_pack_slots(node: ast.Call, scope, helpers: dict, module_name: str, findings: list) -> None:
    """Check the integer fields written by ``S.<LAYOUT>.pack(LIB, ptr, ...)``.

    The ABI reads these through a pointer, so the header signature says
    only ``fm_hyperlink*`` and the scalar sweep sees nothing. The value
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
    if isinstance(payload, ast.Dict):
        pairs = list(zip(payload.keys, payload.values))
    elif isinstance(payload, ast.Call) and isinstance(payload.func, ast.Name):
        helper = helpers.get(payload.func.id)
        returned = _returned_dict(helper) if helper is not None else None
        if returned is not None:
            pairs = list(zip(returned.keys, returned.values))
    elif isinstance(payload, ast.Name) and scope is not None:
        pairs = _collected_payload(payload.id, scope)

    if pairs is None:
        # Report rather than skip: a payload the sweep cannot read is a
        # hole in it, not evidence that the site is safe.
        findings.append(f"{where}: packs a payload this sweep cannot inspect ({ast.unparse(payload)})")
        return
    for key, value in pairs:
        _check_packed_field(layout, key, value, where, findings)


def _scopes(tree: ast.AST) -> dict:
    """Map each node to the innermost function that contains it."""
    enclosing = {}
    for func in ast.walk(tree):
        if not isinstance(func, (ast.FunctionDef, ast.AsyncFunctionDef)):
            continue
        for node in ast.walk(func):
            enclosing[node] = func
    return enclosing


def _sweep_source(source: str, functions: dict, module_name: str = "<probe>") -> list:
    """Run the three checkers over one module's source."""
    tree = ast.parse(source)
    helpers = {n.name: n for n in tree.body if isinstance(n, ast.FunctionDef)}
    enclosing = _scopes(tree)
    findings: list = []
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        _direct_call_slots(node, functions, module_name, findings)
        _indirect_call_slots(node, functions, module_name, findings)
        _struct_pack_slots(node, enclosing.get(node), helpers, module_name, findings)
    return findings


def _unguarded_slots() -> list:
    """Return every unguarded integer argument slot in the package."""
    functions = _parse_header(_HEADER)
    findings: list = []
    for module in sorted(_PKG_ROOT.joinpath("formulon").glob("*.py")):
        findings.extend(_sweep_source(module.read_text(encoding="utf-8"), functions, module.name))
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

    def test_sweep_actually_inspects_call_sites(self) -> None:
        # Guards the guard: a header-parsing or AST-walking regression would
        # otherwise make the sweep above vacuously pass.
        functions = _parse_header(_HEADER)
        self.assertIn("fm_workbook_set_number", functions)
        self.assertEqual(
            [t for t, _n in functions["fm_workbook_set_number"]],
            ["fm_workbook_t*", "size_t", "uint32_t", "uint32_t", "double"],
        )

    def test_sweep_catches_a_direct_call(self) -> None:
        findings = _sweep_source("LIB.fm_workbook_set_number(h, int(sheet), 0, 0, 1.0)", _parse_header(_HEADER))
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
        self.assertIn("sheet <- int(sheet)", findings[0])

    def test_sweep_catches_a_packed_struct_field(self) -> None:
        findings = _sweep_source('S.HYPERLINK.pack(LIB, ptr, {"row": int(row)})', _parse_header(_HEADER))
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

    def test_sweep_reports_a_payload_it_cannot_read(self) -> None:
        # An unreadable payload is a hole in the sweep, not evidence that
        # the site is safe, so it has to surface rather than be skipped.
        findings = _sweep_source("S.HYPERLINK.pack(LIB, ptr, build_it())", _parse_header(_HEADER))
        self.assertEqual(len(findings), 1, findings)
        self.assertIn("cannot inspect", findings[0])

    def test_sweep_accepts_the_guarded_forms(self) -> None:
        functions = _parse_header(_HEADER)
        for source in (
            'LIB.fm_workbook_set_number(h, _uint(sheet, "sheet_index"), 0, 0, 1.0)',
            '_read_count(LIB.fm_sheet_get_merge_count, h, _uint(sheet, "sheet"))',
            'S.HYPERLINK.pack(LIB, ptr, {"row": _uint(row, "row")})',
        ):
            with self.subTest(source=source):
                self.assertEqual(_sweep_source(source, functions), [])


class IndexRangeTests(unittest.TestCase):
    """Out-of-range indices are rejected before any WASM call happens."""

    def test_wrapping_row_does_not_overwrite_the_aliased_cell(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_number(0, 5, 0, 1.0)
            with self.assertRaises(ValueError):
                wb.set_number(0, 2**32 + 5, 0, 99.0)
            self.assertEqual(wb.get_value(0, 5, 0).number, 1.0)

    def test_wrapping_sheet_index_is_rejected(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_number(0, 0, 0, 7.0)
            with self.assertRaises(ValueError):
                wb.set_number(2**32, 0, 0, 99.0)
            with self.assertRaises(ValueError):
                wb.get_value(2**32, 0, 0)
            self.assertEqual(wb.get_value(0, 0, 0).number, 7.0)

    def test_negative_indices_are_rejected(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_number(0, 0, 0, 3.0)
            for sheet, row, col in ((0, -1, 0), (0, 0, -1), (-1, 0, 0)):
                with self.assertRaises(ValueError):
                    wb.set_number(sheet, row, col, 99.0)
            self.assertEqual(wb.get_value(0, 0, 0).number, 3.0)

    def test_huge_indices_are_rejected(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(ValueError):
                wb.set_text(0, 2**64, 0, "wrapped")
            with self.assertRaises(ValueError):
                wb.set_formula(0, 0, 2**48, "=1")
            with self.assertRaises(ValueError):
                wb.sheet_name(2**32 + 0)

    def test_error_message_names_the_parameter(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(ValueError) as ctx:
                wb.set_number(0, 2**32 + 5, 0, 1.0)
            self.assertIn("row", str(ctx.exception))

    def test_count_getters_reject_a_wrapping_sheet(self) -> None:
        # These reach the ABI through a helper that is handed the function
        # rather than calling it inline. Unguarded, the index wrapped to 0
        # and the getter answered about a different sheet -- a plausible
        # count for a question the caller never asked.
        with Workbook.create_default() as wb:
            for name in ("merge_count", "hyperlink_count", "comment_count", "validation_count"):
                with self.subTest(getter=name):
                    with self.assertRaises(ValueError):
                        getattr(wb, name)(2**32)

    def test_trace_walk_rejects_a_wrapping_index(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(ValueError):
                wb.precedents(0, 2**32 + 5, 0, 1)
            with self.assertRaises(ValueError):
                wb.dependents(0, 0, 0, 2**32)

    def test_packed_coordinates_raise_value_error_naming_the_parameter(self) -> None:
        # A coordinate packed into a struct never appears as a scalar
        # argument, so it used to die inside ``struct.pack`` with a
        # ``struct.error`` that no caller guarding on ``ValueError`` sees.
        with Workbook.create_default() as wb:
            with self.assertRaises(ValueError) as ctx:
                wb.add_hyperlink(0, 2**32 + 5, 0, "http://example.invalid")
            self.assertIn("row", str(ctx.exception))

            with self.assertRaises(ValueError) as ctx:
                wb.add_merge(0, formulon.MergeRange(first_row=-1, first_col=0, last_row=1, last_col=1))
            self.assertIn("first_row", str(ctx.exception))

    def test_narrow_struct_fields_are_checked_at_their_own_width(self) -> None:
        # ``u8`` / ``u16`` fields overflow long before uint32 does, and
        # ``struct.pack`` is the only thing that used to notice.
        with Workbook.create_default() as wb:
            with self.assertRaises(ValueError) as ctx:
                wb.add_font(formulon.FontRecord(name="Calibri", size=11.0, underline=256))
            self.assertIn("underline", str(ctx.exception))

    def test_in_range_extremes_still_reach_the_engine(self) -> None:
        # The guard rejects only what cannot survive marshalling; a value
        # inside uint32 must still be handled by the engine's own bounds.
        with Workbook.create_default() as wb:
            with self.assertRaises(formulon.FormulonError):
                wb.set_number(0, 0xFFFFFFFF, 0, 1.0)


if __name__ == "__main__":
    unittest.main()
