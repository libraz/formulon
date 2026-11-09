"""Sheet-level input contracts: formula text, grids, dates, A1 in records, colours and typed errors."""

from __future__ import annotations

import dataclasses
import datetime
import decimal
import json
import unittest
from pathlib import Path

from formulon import (
    AnchorKind,
    AutoFilter,
    AutoFilterSortCondition,
    AutoFilterSortState,
    CellAddress,
    CellRect,
    CfExpressionRuleInput,
    ChartFill,
    ChartLine,
    Color,
    ColorKind,
    DataValidationInput,
    DrawingAnchorInput,
    FormulaSyntaxError,
    FormulonError,
    InvalidArgumentError,
    InvalidFileError,
    ParseError,
    ParseErrorCode,
    PivotInput,
    ShapeInput,
    SheetSelectionInput,
    SparklineGroupInput,
    SparklineInput,
    StatusCode,
    ValidationType,
    Value,
    ValueKind,
    Workbook,
)

REPO_ROOT = Path(__file__).resolve().parents[3]
API_DIR = REPO_ROOT / "tools" / "api"


def _case_table(name: str) -> dict:
    return json.loads((API_DIR / name).read_text(encoding="utf-8"))


class FormulaTextTests(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.ws = self.wb[0]

    def tearDown(self) -> None:
        self.wb.close()

    def test_a_native_string_starting_with_equals_is_a_formula(self) -> None:
        self.ws.set_value("A1", 2)
        self.ws.set_value("B1", "=A1*3")
        self.ws["C1"] = "=B1+1"
        self.assertEqual(self.ws.get_formula("B1"), "=A1*3")
        self.assertEqual(self.ws.get_values("B1:C1"), [[6.0, 7.0]])

    def test_other_leading_characters_and_tagged_text_stay_text(self) -> None:
        for i, text in enumerate(("+1", "-1", "@A1", "＝1", " =1")):
            with self.subTest(text=text):
                self.ws.set_value(i, 3, text)
                self.assertEqual(self.ws.get_value(i, 3), text)
                self.assertIsNone(self.ws.get_formula(i, 3))
        self.ws.set_value("E1", Value(ValueKind.TEXT, text="=1"))
        self.ws.set_text("E2", "=1")
        self.assertEqual(self.ws.get_values("E1:E2"), [["=1"], ["=1"]])

    def test_a_lone_equals_sign_is_a_syntax_error(self) -> None:
        with self.assertRaises(FormulaSyntaxError) as caught:
            self.ws.set_value("A1", "=")
        self.assertNotIsInstance(caught.exception, InvalidArgumentError)
        self.assertIsInstance(caught.exception, ValueError)
        self.assertIsNone(self.ws.get_value("A1"))

    def test_set_values_stops_at_the_failing_cell(self) -> None:
        with self.assertRaises(FormulaSyntaxError) as caught:
            self.ws.set_values("B2", [[1, 2], ["=(", 4]])
        error = caught.exception
        self.assertIn("row=2", error.context)
        self.assertIn("col=1", error.context)
        self.assertIsInstance(error.parse_error, ParseError)
        self.assertEqual(self.ws.get_values("B2:C3"), [[1.0, 2.0], [None, None]])

    def test_grid_checks_run_before_any_write(self) -> None:
        with self.assertRaises(TypeError):
            self.ws.set_values("A1", [[1, object()]])
        with self.assertRaises(InvalidArgumentError):
            self.ws.set_values("A1", [[1, 2], [3]])
        self.assertIsNone(self.ws.get_value("A1"))


class DateWriteTests(unittest.TestCase):
    @staticmethod
    def _input(spec: dict) -> object:
        kind = spec["kind"]
        if kind == "date":
            return datetime.date(spec["year"], spec["month"], spec["day"])
        if kind in ("datetime", "aware_datetime"):
            tz = (
                datetime.timezone(datetime.timedelta(minutes=spec["utc_offset_minutes"]))
                if "utc_offset_minutes" in spec
                else None
            )
            micro = spec.get("subsecond", {}).get("py_us", 0)
            return datetime.datetime(
                spec["year"], spec["month"], spec["day"], spec["hour"], spec["minute"], spec["second"], micro, tz
            )
        if kind == "time":
            return datetime.time(spec["hour"], spec["minute"], spec["second"], spec["subsecond"]["py_us"])
        if kind == "timedelta":
            return datetime.timedelta(days=spec["days"])
        if kind == "decimal":
            return decimal.Decimal(spec["value"])
        raise AssertionError(f"no Python input for {kind}")

    @staticmethod
    def _preformat(ws, effective: str) -> None:
        if effective == "custom":
            ws.update_cell_format("A1", number_format="0.00")
        elif effective == "row_style":
            ws.update_cell_format("1:1", number_format="0.00")
        elif effective == "column_style":
            ws.update_cell_format("A:A", number_format="0.00")

    def test_table_rows(self) -> None:
        errors = {"TypeError": TypeError, "ValueError": ValueError, "InvalidArgument": InvalidArgumentError}
        ran = 0
        for case in _case_table("date_value_cases.json")["cases"]:
            expect = case["expect"].get("py")
            if case["direction"] != "write" or expect is None:
                continue
            with self.subTest(case["id"]), Workbook.create_default() as wb:
                wb.set_date1904(case["date1904"])
                ws = wb[0]
                self._preformat(ws, case["effective_format"])
                before = ws.get_cell_format("A1").number_format
                value = self._input(case["input"])
                ran += 1
                if "error" in expect:
                    with self.assertRaises(errors[expect["error"]]):
                        ws.set_value("A1", value)
                    self.assertIsNone(ws.get_value("A1"))
                    continue
                ws["A1"] = value
                self.assertAlmostEqual(ws.get_value("A1"), expect["serial"], delta=abs(expect["serial"]) * 1e-12)
                fmt = ws.get_cell_format("A1").number_format
                if expect["format_id"] is None:
                    self.assertEqual(fmt, before)
                else:
                    # The en-US code lands on its built-in id, which renders per profile.
                    xf = wb.ooxml.get_cell_xf(ws.ooxml.get_cell_xf_index("A1"))
                    self.assertEqual(xf.num_fmt_id, expect["format_id"])
        self.assertGreater(ran, 25)

    def test_set_values_takes_dates(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb[0]
            ws.set_values("A1", [[datetime.date(2024, 3, 15), datetime.time(12)]])
            self.assertEqual(ws.get_values("A1:B1"), [[45366.0, 0.5]])
            self.assertEqual(ws.get_date("A1"), datetime.datetime(2024, 3, 15))
            self.assertEqual(wb.ooxml.get_cell_xf(ws.ooxml.get_cell_xf_index("B1")).num_fmt_id, 21)
            self.assertEqual(ws.get_display_text("B1").text, "12:00:00")


class RecordA1Tests(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.ws = self.wb[0]

    def tearDown(self) -> None:
        self.wb.close()

    def test_cf_rule_ranges(self) -> None:
        self.ws.add_cf_rule(CfExpressionRuleInput(ranges="A1:B2", formula1="TRUE"))
        self.ws.add_cf_rule(
            CfExpressionRuleInput(
                ranges=["C3", CellRect(first_row=4, first_col=0, last_row=4, last_col=0)], formula1="TRUE"
            )
        )
        rules = self.ws.list_cf_rules()
        self.assertEqual(rules[0].ranges, [CellRect(first_row=0, first_col=0, last_row=1, last_col=1)])
        self.assertEqual(
            rules[1].ranges,
            [
                CellRect(first_row=2, first_col=2, last_row=2, last_col=2),
                CellRect(first_row=4, first_col=0, last_row=4, last_col=0),
            ],
        )

    def test_validation_ranges(self) -> None:
        self.ws.add_validation(
            DataValidationInput(type=ValidationType.WHOLE, ranges="B2:B5", formula1="1", formula2="9")
        )
        self.assertEqual(
            self.ws.list_validations()[0].ranges, [CellRect(first_row=1, first_col=1, last_row=4, last_col=1)]
        )

    def test_a_qualified_range_in_a_record_is_refused(self) -> None:
        with self.assertRaises(InvalidArgumentError):
            self.ws.add_validation(DataValidationInput(type=ValidationType.WHOLE, ranges="Sheet1!B2", formula1="1"))
        self.assertEqual(self.ws.list_validations(), [])

    def test_pivot_anchor(self) -> None:
        cache = self.wb.add_pivot_cache()
        by_text = self.ws.add_pivot(PivotInput(name="ByText", cache_id=cache.cache_id, anchor="E3"))
        by_address = self.ws.add_pivot(
            PivotInput(name="ByAddress", cache_id=cache.cache_id, anchor=CellAddress(row=20, col=1))
        )
        self.assertEqual([p.name for p in self.ws.list_pivots()], [by_text.name, by_address.name])
        with self.assertRaises(InvalidArgumentError):
            self.ws.add_pivot(PivotInput(name="Wide", cache_id=cache.cache_id, anchor="E3:F4"))

    def test_auto_filter_sort_refs(self) -> None:
        self.ws.set_values("A1", [["k", "v"], [2, "b"], [1, "a"]])
        sort = AutoFilterSortState(range="A2:B3", conditions=[AutoFilterSortCondition(range="A2:A3")])
        self.ws.set_auto_filter(
            AutoFilter(range=CellRect(first_row=0, first_col=0, last_row=2, last_col=1), sort=sort)  # type: ignore[arg-type]
        )
        got = self.ws.get_auto_filter()
        assert got is not None and got.sort is not None
        self.assertEqual(got.sort.range, CellRect(first_row=1, first_col=0, last_row=2, last_col=1))
        self.assertEqual(got.sort.conditions[0].range, CellRect(first_row=1, first_col=0, last_row=2, last_col=0))

    def test_print_area_and_selection_take_one_string(self) -> None:
        self.ws.set_print_area("A1:C3")
        self.assertEqual(self.ws.get_print_area(), [CellRect(first_row=0, first_col=0, last_row=2, last_col=2)])
        self.ws.set_selection(SheetSelectionInput(CellAddress(1, 1), ranges="B2:C3"))
        self.assertEqual(self.ws.get_selection().ranges, [CellRect(first_row=1, first_col=1, last_row=2, last_col=2)])


class ColourInputTests(unittest.TestCase):
    def test_shape_colours(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb[0]
            shape = ws.add_shape(
                ShapeInput(
                    anchor=DrawingAnchorInput(kind=AnchorKind.TWO_CELL, to_row=3, to_col=3),
                    fill=ChartFill(color="#FF0000"),  # type: ignore[arg-type]
                    line=ChartLine(color=Color(kind=ColorKind.RGB, argb=0x800000FF)),
                ),
            )
            # A shape colour reads back opaque; the alpha of the input is ignored.
            props = shape.get_properties()
            self.assertEqual(props.fill.color, Color(kind=ColorKind.RGB, argb=0xFFFF0000))
            self.assertEqual(props.line.color, Color(kind=ColorKind.RGB, argb=0xFF0000FF))
            shape.update(fill={"color": "#00FF00"})
            self.assertEqual(shape.get_properties().fill.color, Color(kind=ColorKind.RGB, argb=0xFF00FF00))
            passed_back = {
                f.name: getattr(props, f.name) for f in dataclasses.fields(props) if f.name in ("fill", "line")
            }
            shape.update(**passed_back)
            self.assertEqual(shape.get_properties().fill.color, props.fill.color)
            shape.update(fill=None)
            self.assertIsNone(shape.get_properties().fill)
            self.assertEqual(shape.get_properties().line.color, props.line.color)

    def test_sparkline_colours_take_text(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb[0]
            ws.set_values("A1", [[1, 2, 3]])
            group_id = ws.add_sparkline_group(
                SparklineGroupInput(color_series="#112233", sparklines=[SparklineInput("Sheet1!A1:C1", "D1")])
            )
            self.assertEqual(ws.list_sparkline_groups()[0].id, group_id)
            self.assertEqual(ws.list_sparkline_groups()[0].color_series, Color(kind=ColorKind.RGB, argb=0xFF112233))
            ws.update_sparkline_group(group_id, color_high="#445566")
            self.assertEqual(ws.list_sparkline_groups()[0].color_high, Color(kind=ColorKind.RGB, argb=0xFF445566))

    def test_cell_format_colour_keys(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb[0]
            ws.update_cell_format("A1", fill={"foreground": "#336699"}, border={"left": {"color": "#000000"}})
            fmt = ws.get_cell_format("A1")
            self.assertEqual(fmt.fill.foreground, Color(kind=ColorKind.RGB, argb=0xFF336699))
            self.assertEqual(fmt.border.left.color, Color(kind=ColorKind.RGB, argb=0xFF000000))
            for patch in ({"font": {"color_argb": 0xFF000000}}, {"fill": {"fg_argb": 0}}):
                with self.subTest(patch=patch), self.assertRaises(TypeError):
                    ws.update_cell_format("A1", **patch)


class TypedErrorTests(unittest.TestCase):
    def test_parse_error_is_a_frozen_record(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(FormulaSyntaxError) as caught:
                wb[0].set_formula("A1", "=SUM(1,")
        record = caught.exception.parse_error
        self.assertIsInstance(record, ParseError)
        self.assertIsInstance(record.code, ParseErrorCode)
        self.assertFalse(hasattr(caught.exception, "parse_error_code"))
        with self.assertRaises(dataclasses.FrozenInstanceError):
            record.start = 0  # type: ignore[misc]

    def test_invalid_file_statuses_construct_the_subclass(self) -> None:
        for status in (StatusCode.IO_ZIP_CORRUPT, StatusCode.IO_XML_PARSE, StatusCode.IO_XLSB_CORRUPT):
            with self.subTest(status=status.name):
                error = FormulonError(status, "bad")
                self.assertIs(type(error), InvalidFileError)
        for status in (StatusCode.IO_WRITE_FAILED, StatusCode.NOT_FOUND):
            with self.subTest(status=status.name):
                self.assertNotIsInstance(FormulonError(status, "x"), InvalidFileError)
        with self.assertRaises(InvalidFileError):
            Workbook.load(b"not a zip archive")


if __name__ == "__main__":
    unittest.main()
