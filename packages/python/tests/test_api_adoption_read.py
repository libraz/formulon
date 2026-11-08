"""Sheet-level read contracts: native and typed reads, range keys, dates, the used range and the ooxml layers."""

from __future__ import annotations

import dataclasses
import datetime
import json
import math
import unittest
from pathlib import Path

from formulon import (
    CalcMode,
    CellRect,
    Color,
    ColorKind,
    ColumnLayout,
    ErrorCode,
    ExcelError,
    InvalidArgumentError,
    RangeCell,
    RowLayout,
    Value,
    ValueKind,
    Workbook,
    to_native,
)

REPO_ROOT = Path(__file__).resolve().parents[3]
API_DIR = REPO_ROOT / "tools" / "api"


def _case_table(name: str) -> dict:
    return json.loads((API_DIR / name).read_text(encoding="utf-8"))


def _expected_native(value: object) -> object:
    """A case-table native value with `{error: <ErrorCode member>}` mapped to an `ExcelError`."""
    if isinstance(value, dict):
        return ExcelError(ErrorCode[value["error"].upper()])
    if isinstance(value, list):
        return [_expected_native(item) for item in value]
    return float(value) if isinstance(value, int) and not isinstance(value, bool) else value


class NativeReadTests(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.ws = self.wb[0]

    def tearDown(self) -> None:
        self.wb.close()

    def test_facade_cases_of_the_sheet_reads(self) -> None:
        for case in _case_table("to_native_cases.json")["facade_cases"]:
            with self.subTest(case["id"]):
                call = case["call"]
                if call == "get_value":
                    got = self.ws.get_value(case["cell"])
                elif call == "get_values":
                    got = self.ws.get_values(case["range"])
                else:
                    got = getattr(self.ws, call)("H20", case["formula"])
                self.assertEqual(got, _expected_native(case["native"]))

    def test_get_value_is_native_and_the_typed_read_keeps_the_kind(self) -> None:
        self.ws.set_values("A1", [[42, "x", True, ExcelError(ErrorCode.NA)]])
        self.assertEqual(self.ws.get_value("A1"), 42.0)
        self.assertIsInstance(self.ws.get_value(0, 0), float)
        self.assertEqual(self.ws.get_value("B1"), "x")
        self.assertIs(self.ws.get_value("C1"), True)
        self.assertEqual(self.ws.get_value("D1"), ExcelError(ErrorCode.NA))
        self.assertIsNone(self.ws.get_value("E1"))
        self.assertEqual(self.ws.get_typed_value("A1"), Value(ValueKind.NUMBER, number=42.0))
        self.assertEqual(self.ws.get_typed_value(0, 4), Value(ValueKind.BLANK))

    def test_grids_native_and_typed(self) -> None:
        self.ws.set_values("A1", [[1, "a"], [None, False]])
        self.assertEqual(self.ws.get_values("A1:B2"), [[1.0, "a"], [None, False]])
        typed = self.ws.get_typed_values(CellRect(first_row=0, first_col=0, last_row=1, last_col=1))
        self.assertEqual(typed[0][1], Value(ValueKind.TEXT, text="a"))
        self.assertEqual(typed[1][0].kind, ValueKind.BLANK)

    def test_typed_grid_round_trips_through_set_values(self) -> None:
        self.ws.set_values("A1", [["=1+1", "x"]])
        self.ws.set_text("C1", "=not a formula")
        grid = self.ws.get_typed_values("A1:C1")
        self.ws.set_values("A2", grid)
        self.assertEqual(self.ws.get_values("A2:C2"), [[2.0, "x", "=not a formula"]])
        self.assertIsNone(self.ws.get_formula("A2"))
        self.assertIsNone(self.ws.get_formula("C2"))

    def test_list_cells_and_pages_are_range_cells(self) -> None:
        self.ws.set_values("A1", [[1, "=A1*2"]])
        cells = self.ws.list_cells()
        self.assertEqual(
            cells,
            [
                RangeCell(row=0, col=0, formula=None, value=1.0, kind=ValueKind.NUMBER),
                RangeCell(row=0, col=1, formula="=A1*2", value=2.0, kind=ValueKind.NUMBER),
            ],
        )
        page = self.ws.get_cells_in_range("A1:B1", None, 1)
        self.assertEqual(page.cells, cells[:1])
        rest = self.ws.get_cells_in_range("A1:B1", page.next_cursor)
        self.assertEqual(rest.cells, cells[1:])
        self.assertIsNone(rest.next_cursor)

    def test_evaluate_typed_forms(self) -> None:
        folded = self.ws.evaluate_typed_formula("H20", "=SEQUENCE(2,1)")
        self.assertEqual(folded.kind, ValueKind.ARRAY)
        self.assertEqual(to_native(folded), [[1.0], [2.0]])
        self.assertEqual(self.ws.evaluate_typed_formula(19, 7, "=1+2"), Value(ValueKind.NUMBER, number=3.0))
        grid = self.ws.evaluate_typed_formula_array("H20", "=1+2")
        self.assertEqual(grid, [[Value(ValueKind.NUMBER, number=3.0)]])
        with self.assertRaises(TypeError):
            self.ws.evaluate_formula("A1", 1)  # type: ignore[arg-type]
        with self.assertRaises(TypeError):
            self.ws.evaluate_formula_array("A1", None)  # type: ignore[arg-type]

    def test_a_read_calculates_pending_cells_under_auto_only(self) -> None:
        self.ws.set_value("A1", 2)
        self.ws.set_formula("B1", "=A1*10")
        self.ws.set_value("A1", 3)
        self.assertEqual(self.ws.get_value("B1"), 30.0)
        self.wb.set_calc_mode(CalcMode.MANUAL)
        self.ws.set_value("A1", 4)
        self.assertEqual(self.ws.get_value("B1"), 30.0)
        self.wb.recalc()
        self.assertEqual(self.ws.get_value("B1"), 40.0)


class RangeKeyTests(unittest.TestCase):
    def test_range_key_reads_and_writes_a_grid(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb[0]
            ws["A1:B2"] = [[1, "=A1+1"], ["x", None]]
            self.assertEqual(ws["A1:B2"], [[1.0, 2.0], ["x", None]])
            self.assertEqual(ws["B1"], 2.0)
            self.assertEqual(ws[0, 1], 2.0)
            self.assertEqual(ws["C3:C3"], [[None]])

    def test_range_key_rejects_a_grid_of_another_shape(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb[0]
            with self.assertRaises(InvalidArgumentError):
                ws["A1:B2"] = [[1, 2]]
            with self.assertRaises(TypeError):
                ws["A1:B2"] = 5
            self.assertIsNone(ws["A1"])

    def test_int_and_slice_keys_are_type_errors(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb[0]
            for key in (0, slice(0, 2)):
                with self.subTest(key=key):
                    with self.assertRaises(TypeError):
                        ws[key]  # type: ignore[index]
                    with self.assertRaises(TypeError):
                        ws[key] = 1  # type: ignore[index]


class DateReadTests(unittest.TestCase):
    def _fill(self, ws, kind: str) -> None:
        if kind == "text":
            ws.set_text("A1", "2024-01-01")
        elif kind == "bool":
            ws.set_bool("A1", True)
        elif kind == "error":
            ws.set_error("A1", ErrorCode.NA)

    def test_table_rows(self) -> None:
        ran = 0
        for case in _case_table("date_value_cases.json")["cases"]:
            expect = case["expect"].get("py")
            if case["direction"] != "read" or expect is None:
                continue
            spec = case["input"]
            with self.subTest(case["id"]), Workbook.create_default() as wb:
                wb.set_date1904(case["date1904"])
                ws = wb[0]
                if "serial" in spec:
                    serial = float(spec["serial"]) if isinstance(spec["serial"], str) else spec["serial"]
                    if not math.isfinite(serial):
                        # A cell cannot hold a non-finite number; the shared helper's rows cover these.
                        continue
                    ws.set_number("A1", serial)
                else:
                    self._fill(ws, spec["cell"])
                ran += 1
                if "error" in expect:
                    with self.assertRaises(InvalidArgumentError):
                        ws.get_date("A1")
                elif "value" in expect:
                    self.assertIsNone(ws.get_date(0, 0))
                else:
                    d = expect["date"]
                    self.assertEqual(
                        ws.get_date("A1"),
                        datetime.datetime(
                            d["year"], d["month"], d["day"], d["hour"], d["minute"], d["second"], d["microsecond"]
                        ),
                    )
        self.assertGreater(ran, 20)


class UsedRangeTests(unittest.TestCase):
    def test_empty_sheet_has_none(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb[0]
            ws.update_cell_format("D4", number_format="0.00")
            self.assertIsNone(ws.get_used_range())
            self.assertEqual(list(ws.iter_rows()), [])

    def test_bounding_rectangle_of_values_and_formulas(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb[0]
            ws.set_value("C2", 1)
            ws.set_formula("B4", "=C2")
            ws.set_text("E3", "x")
            ws.update_cell_format("H10", number_format="0.00")
            self.assertEqual(ws.get_used_range(), CellRect(first_row=1, first_col=1, last_row=3, last_col=4))
            self.assertEqual(
                list(ws.iter_rows()),
                [
                    [None, 1.0, None, None],
                    [None, None, None, "x"],
                    [1.0, None, None, None],
                ],
            )
            self.assertEqual(list(ws.iter_rows("C2:C3")), [[1.0], [None]])


class OoxmlLayerTests(unittest.TestCase):
    def test_layers_are_stable_and_hold_the_index_ops(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb[0]
            self.assertIs(wb.ooxml, wb.ooxml)
            self.assertIs(ws.ooxml, ws.ooxml)
            for name in ("get_font", "add_cell_xf", "list_passthroughs"):
                self.assertFalse(hasattr(wb, name), name)
                self.assertTrue(callable(getattr(wb.ooxml, name)), name)
            for name in ("list_cell_styles", "get_cell_style", "set_cell_style", "remove_cell_style"):
                self.assertTrue(callable(getattr(wb, name)), name)
                self.assertFalse(hasattr(wb.ooxml, name), name)
            for name in ("get_cell_xf_index", "set_column_xf_index", "get_sheet_pr_xml", "remap_range_xfs"):
                self.assertFalse(hasattr(ws, name), name)
                self.assertTrue(callable(getattr(ws.ooxml, name)), name)

    def test_column_and_row_style_reads(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb[0]
            ws.update_cell_format("Z1", font={"bold": True})
            xf = ws.ooxml.get_cell_xf_index("Z1")
            self.assertNotEqual(xf, 0)
            ws.ooxml.set_column_xf_index("B:C", xf)
            ws.ooxml.set_row_xf_index(4, 4, xf)
            self.assertEqual(ws.ooxml.get_column_xf_index(2), xf)
            self.assertIsNone(ws.ooxml.get_column_xf_index(3))
            self.assertEqual(ws.ooxml.get_row_xf_index(4), xf)
            self.assertIsNone(ws.ooxml.get_row_xf_index(5))
            for record in (ColumnLayout, RowLayout):
                self.assertFalse({"style_xf", "xf_index"} & {f.name for f in dataclasses.fields(record)})
            self.assertEqual([(c.first, c.last) for c in ws.list_column_overrides()], [(1, 2)])
            ws.ooxml.set_cell_xf_index("A1", xf)
            self.assertEqual(ws.ooxml.get_cell_xf_index(0, 0), xf)
            self.assertTrue(wb.ooxml.get_font(wb.ooxml.get_cell_xf(xf).font_index).bold)

    def test_set_cell_format_writes_through_the_style_tables(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb[0]
            ws.update_cell_format("A1", font={"color": "#FF0000"})
            fmt = ws.get_cell_format("A1")
            self.assertEqual(fmt.font.color, Color(kind=ColorKind.RGB, argb=0xFFFF0000))


class TabColorTests(unittest.TestCase):
    def test_absent_set_and_cleared(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb[0]
            self.assertIsNone(ws.get_tab_color())
            ws.set_tab_color("#00FF00")
            self.assertEqual(ws.get_tab_color(), Color(kind=ColorKind.RGB, argb=0xFF00FF00))
            theme = Color(kind=ColorKind.THEME, theme=4, tint=0.25)
            ws.set_tab_color(theme)
            self.assertEqual(ws.get_tab_color(), theme)
            ws.set_tab_color(None)
            self.assertIsNone(ws.get_tab_color())

    def test_tab_colour_keeps_the_rest_of_sheet_pr(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb[0]
            ws.update_page_setup(fit_to_page=True)
            ws.set_tab_color("#123456")
            self.assertTrue(ws.get_page_setup().fit_to_page)
            self.assertIn("tabColor", ws.ooxml.get_sheet_pr_xml() or "")

    def test_bad_colour_input(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(TypeError):
                wb[0].set_tab_color("red")


if __name__ == "__main__":
    unittest.main()
