"""Range guards on the integer arguments the binding hands to WebAssembly.

Two layers are covered:

  * A static sweep pairing every ``LIB.fm_*`` call in the package with the
    parameter list the C ABI header declares, asserting that each by-value
    integer slot is either a literal or is wrapped in ``_uint`` / ``_sint``.
    It fails as soon as a new call site (or a new parameter on an existing
    one) reaches WASM unguarded.
  * Behavioural tests for the wrap that motivates the guard: ``wasmtime``
    marshals a Python int into an ``i32`` modulo 2**32, so ``2**32 + 5``
    would otherwise land on row 5 and overwrite a live cell.
"""

from __future__ import annotations

import ast
import re
import unittest
from pathlib import Path

import formulon
from formulon import Workbook

_PKG_ROOT = Path(__file__).resolve().parent.parent
_HEADER = _PKG_ROOT.parent.parent / "src" / "c_api" / "formulon_c.h"

# C parameter types passed by value in a 32-bit WASM ABI. Anything else in
# a signature is a pointer (an offset the binding computes itself) or a
# double, neither of which a caller-supplied index can corrupt.
_UNSIGNED_TYPES = {"uint32_t", "size_t", "uint16_t", "uint8_t"}
_SIGNED_TYPES = {"int32_t", "fm_status_t", "fm_error_code_t"}


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


def _unguarded_slots() -> list:
    """Return every unguarded integer argument slot in the package."""
    functions = _parse_header(_HEADER)
    findings = []
    for module in sorted(_PKG_ROOT.joinpath("formulon").glob("*.py")):
        tree = ast.parse(module.read_text(encoding="utf-8"))
        for node in ast.walk(tree):
            if not (
                isinstance(node, ast.Call)
                and isinstance(node.func, ast.Attribute)
                and isinstance(node.func.value, ast.Name)
                and node.func.value.id == "LIB"
                and node.func.attr.startswith("fm_")
            ):
                continue
            params = functions.get(node.func.attr)
            where = f"{module.name}:{node.lineno} {node.func.attr}"
            if params is None:
                findings.append(f"{where}: not declared in formulon_c.h")
                continue
            if len(node.args) > len(params):
                findings.append(f"{where}: passes {len(node.args)} args for {len(params)} parameters")
                continue
            for arg, (ptype, pname) in zip(node.args, params):
                if ptype not in _UNSIGNED_TYPES and ptype not in _SIGNED_TYPES:
                    continue
                if not _is_guarded(arg):
                    findings.append(f"{where}: {ptype} {pname} <- {ast.unparse(arg)}")
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

    def test_in_range_extremes_still_reach_the_engine(self) -> None:
        # The guard rejects only what cannot survive marshalling; a value
        # inside uint32 must still be handled by the engine's own bounds.
        with Workbook.create_default() as wb:
            with self.assertRaises(formulon.FormulonError):
                wb.set_number(0, 0xFFFFFFFF, 0, 1.0)


if __name__ == "__main__":
    unittest.main()
