"""Python binding parity surface tests."""

from __future__ import annotations

import inspect
import json
import re
import unittest
from pathlib import Path
from typing import Dict, Iterable, Set

import formulon
from formulon import (
    Sheet,
    Workbook,
)

REPO_ROOT = Path(__file__).resolve().parents[3]
DTS = REPO_ROOT / "src" / "wasm" / "formulon.d.ts"
PLATFORM_DIFF = REPO_ROOT / "tools" / "api" / "platform_diff.json"


def _snake(name: str) -> str:
    split = re.compile(r"(?<=[a-z])([A-Z])|(?<=[A-Z])([A-Z][a-z])")
    return split.sub(lambda m: "_" + (m.group(1) or m.group(2)), name).lower()


def _interface_members(name: str) -> Set[str]:
    """Member names of one exported interface in ``formulon.d.ts`` (methods, properties, ``[Symbol.dispose]``)."""
    text = DTS.read_text(encoding="utf-8")
    text = re.sub(r"/\*[\s\S]*?\*/", "", text)
    text = re.sub(r"//[^\n]*", "", text)
    match = re.search(rf"export interface {name}\s*\{{", text)
    if match is None:
        raise AssertionError(f"no interface {name} in {DTS}")
    start = match.end()
    depth = 1
    pos = start
    while depth:
        ch = text[pos]
        depth += ch == "{"
        depth -= ch == "}"
        pos += 1
    body = text[start : pos - 1]
    names: Set[str] = set()
    current = ""
    depth = 0
    for ch in body:
        if ch == ";" and depth == 0:
            found = re.match(r"^(?:readonly\s+)?(\[Symbol\.dispose\]|[A-Za-z_$][\w$]*)\??\s*[(<:]", current.strip())
            if found:
                names.add(found.group(1))
            current = ""
            continue
        depth += ch in "({"
        depth -= ch in ")}"
        current += ch
    return names


def _platform_diff() -> Dict[str, Dict[str, Set[str]]]:
    raw = json.loads(PLATFORM_DIFF.read_text(encoding="utf-8"))
    return {
        key: {scope: set(names) for scope, names in raw[key].items()} for key in ("js_only", "native_only", "py_only")
    }


def _public(owner: object) -> Set[str]:
    return {n for n in dir(owner) if not n.startswith("_")}


def _dunders(owner: object) -> Set[str]:
    return {n for n in dir(owner) if n.startswith("__") and n.endswith("__") and n in _DUNDER_SURFACE}


# The dunders the platform diff may name; every other dunder is Python plumbing.
_DUNDER_SURFACE = {"__getitem__", "__setitem__", "__len__", "__iter__", "__contains__", "__enter__", "__exit__"}


class SurfaceSetParityTests(unittest.TestCase):
    """Criterion 15: the JS and Python member sets match apart from the entries in ``platform_diff.json``."""

    @classmethod
    def setUpClass(cls) -> None:
        cls.diff = _platform_diff()

    def _compare(self, js: Iterable[str], py: Iterable[str], js_only: Set[str], py_only: Set[str]) -> None:
        js_names = {_snake(n) for n in js if not n.startswith("[")} - {_snake(n) for n in js_only}
        py_names = set(py) - py_only
        self.assertEqual(sorted(js_names - py_names), [], "on the JS surface, missing from Python")
        self.assertEqual(sorted(py_names - js_names), [], "on the Python surface, missing from JS")

    def _surface(self, scope: str, js: Set[str], py_owner: object) -> None:
        # A native-only member is on the native npm package, so it counts as JS surface.
        js = js | self.diff["native_only"].get(scope, set())
        js_only = self.diff["js_only"].get(scope, set())
        py_only = self.diff["py_only"].get(scope, set())
        py_names = _public(py_owner) | (_dunders(py_owner) & py_only)
        self._compare(js, py_names, js_only, py_only)

    def test_workbook(self) -> None:
        # The static factories are declared on `WorkbookCtor`; Python has them as class methods.
        js = _interface_members("Workbook") | _interface_members("WorkbookCtor")
        self._surface("Workbook", js, formulon.Workbook)

    def test_sheet(self) -> None:
        self._surface("Sheet", _interface_members("Sheet"), formulon.Sheet)

    def test_module_functions_match_formulon(self) -> None:
        js_members = _interface_members("Formulon")
        js_functions = {n for n in js_members if n[:1].islower()}
        js_only = {n for n in self.diff["js_only"].get("Formulon", set()) if n[:1].islower()}
        py_only = self.diff["py_only"].get("module", set())
        py_functions = {
            n
            for n in _public(formulon)
            if inspect.isfunction(getattr(formulon, n)) and getattr(formulon, n).__module__.startswith("formulon")
        }
        self._compare(js_functions, py_functions, js_only, py_only)

    def test_value_tables_exist_on_the_module(self) -> None:
        js_tables = {n for n in _interface_members("Formulon") if n[:1].isupper() and n != "Workbook"}
        js_only = self.diff["js_only"].get("Formulon", set())
        missing = sorted(n for n in js_tables - js_only if not hasattr(formulon, n))
        self.assertEqual(missing, [], "value tables without a module counterpart")

    def test_error_display_name_is_an_enum_property(self) -> None:
        self.assertIn("ErrorDisplayName", self.diff["js_only"]["Formulon"])
        self.assertIsInstance(formulon.ErrorCode.NA.display_name, str)

    def test_platform_diff_names_exist(self) -> None:
        for name in self.diff["py_only"]["Workbook"]:
            self.assertTrue(hasattr(formulon.Workbook, name), f"Workbook.{name}")
        for name in self.diff["py_only"]["Sheet"]:
            self.assertTrue(hasattr(formulon.Sheet, name), f"Sheet.{name}")
        for name in self.diff["py_only"]["module"]:
            self.assertTrue(hasattr(formulon, name), name)


class SurfaceParityTests(unittest.TestCase):
    """Every advertised method must exist on ``Workbook`` or ``Sheet``."""

    METHODS = [
        "Sheet.move",
        "copy_sheet",
        "remove_sheet",
        "Sheet.rename",
        "set_defined_name",
        "set_error",
        "insert_rows",
        "delete_rows",
        "insert_columns",
        "delete_columns",
        "get_calc_mode",
        "set_calc_mode",
        "add_data_table",
        "get_data_table",
        "remove_data_table",
        "recalc_data_tables",
        "get_pinned_now",
        "set_pinned_now",
        "get_profile_id",
        "set_profile_id",
        "recalc_viewport",
        "get_lambda_text",
        "evaluate_formula_array",
        "add_merge",
        "remove_merge",
        "clear_merges",
        "list_merges",
        "set_hyperlink",
        "insert_hyperlink_at",
        "remove_hyperlink",
        "remove_hyperlink_at",
        "clear_hyperlinks",
        "list_hyperlinks",
        "get_comment",
        "set_comment",
        "list_validations",
        "list_protected_ranges",
        "list_sparkline_groups",
        "get_cell_phonetic_runs",
        "list_show_phonetic_cells",
        "list_tables",
        "add_validation",
        "remove_validation",
        "clear_validations",
        "get_sheet_protection",
        "set_sheet_protection",
        "get_view",
        "update_view",
        "get_values",
        "set_values",
        "list_cells",
        "list_pivots",
        "Pivot.list_field_items",
        "Pivot.update_field",
        "Pivot.update_field_item",
        "has_dirty_cells",
        "list_fonts",
        "list_fills",
        "list_borders",
        "list_cell_xfs",
        "list_cell_style_xfs",
        "list_dxfs",
        "PivotCache.list_fields",
        "PivotCache.list_field_shared_items",
        "PivotCache.add_field_shared_item",
        "PivotCache.set_record_value",
        "list_column_overrides",
        "set_column_width",
        "set_column_hidden",
        "set_column_outline",
        "set_column_xf_index",
        "reset_column_xf_index",
        "list_row_overrides",
        "set_row_height",
        "set_row_auto_height",
        "set_row_hidden",
        "set_row_outline",
        "set_row_xf_index",
        "reset_row_xf_index",
        "get_range_xf_bases",
        "get_range_xf_bases_excluding",
        "list_cell_xf_indices",
        "remap_range_xfs",
        "snapshot_range_xfs",
        "restore_range_xfs",
        "evaluate_cf_range",
        "list_cf_rules",
        "add_cf_rule",
        "remove_cf_rule",
        "clear_cf_rules",
        "get_cell_xf_index",
        "set_cell_xf_index",
        "reset_cell_xf_index",
        "get_cell_xf",
        "get_font",
        "get_fill",
        "get_border",
        "get_number_format",
        "list_cell_styles",
        "add_font",
        "add_fill",
        "add_border",
        "add_number_format",
        "add_cell_xf",
        "get_cell_style",
        "get_cell_style_xf",
        "Pivot.get_layout",
        "list_pivot_caches",
        "add_pivot_cache",
        "remove_pivot_cache",
        "PivotCache.add_field",
        "PivotCache.clear_fields",
        "PivotCache.add_record",
        "add_pivot",
        "remove_pivot",
        "Pivot.rename",
        "Pivot.set_anchor",
        "Pivot.update_grand_totals",
        "Pivot.update_style_options",
        "Pivot.add_field_item",
        "Pivot.set_field_axis",
        "Pivot.add_data_field",
        "Pivot.set_data_field",
        "Pivot.add_filter",
        "Pivot.set_row_field_order",
        "Pivot.set_column_field_order",
        "get_precedents",
        "get_dependents",
        "get_spill",
        "parse_reference",
        "format_reference",
        "get_array_formula",
        "set_array_formula",
        "list_external_links",
        "save",
        "sheet",
        "save_with_diagnostics",
        "get_diagnostics",
    ]

    def test_all_methods_present(self) -> None:
        handles = {
            "Pivot": formulon.Pivot,
            "PivotCache": formulon.PivotCache,
            "Sheet": Sheet,
            "Table": formulon.Table,
            "Chart": formulon.Chart,
        }
        for name in self.METHODS:
            if "." in name:
                owner, member = name.split(".")
                self.assertTrue(hasattr(handles[owner], member), f"missing method: {name}")
            else:
                owners = (Workbook, Sheet, formulon.WorkbookOoxml, formulon.SheetOoxml)
                self.assertTrue(any(hasattr(owner, name) for owner in owners), f"missing method: {name}")

    def test_default_workbook_uses_english_profile(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(wb.get_profile_id(), "win-365-en_US")

    def test_public_types_exported(self) -> None:
        for name in (
            "CellRect",
            "Sheet",
            "ExcelError",
            "IndexOutOfRangeError",
            "WorkbookClosedError",
            "Comment",
            "Hyperlink",
            "DataValidation",
            "DataValidationInput",
            "SheetProtection",
            "SheetView",
            "ConditionalFormat",
            "ConditionalFormatInput",
            "CfMatch",
            "CfCell",
            "CellNode",
            "Spill",
            "ParsedReference",
            "ChartDisplayPatch",
            "FunctionMetadata",
            "LocaleFacts",
            "LocaleErrorName",
            "CollationStrength",
            "CollationClasses",
            "CellXf",
            "Color",
            "Font",
            "Fill",
            "CellStyle",
            "ExternalLink",
            "PivotCell",
            "PivotLayout",
            "PivotDataFieldInput",
            "PivotFilterInput",
            "CalcMode",
            "PivotAxis",
            "PivotAggregation",
            "MergedFunctionMetadata",
            "merge_function_metadata",
        ):
            self.assertTrue(hasattr(formulon, name), f"formulon missing: {name}")


if __name__ == "__main__":
    unittest.main()
