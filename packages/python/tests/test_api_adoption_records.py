"""Workbook-level adoption contracts: native values, syntax errors, records, colours, enums and packaging."""

from __future__ import annotations

import dataclasses
import datetime
import json
import re
import tempfile
import typing
import unittest
from pathlib import Path

import formulon
from formulon import Border as BorderRecord
from formulon import BorderSide as BorderSideRecord
from formulon import (
    CalcMode,
    CellAddress,
    CellRect,
    CfRuleType,
    CfValueObjectType,
    Cfvo,
    Color,
    ColorKind,
    ComparisonOperator,
    ConditionalFormatInput,
    DataValidationInput,
    DatePeriod,
    DefinedNameInput,
    ErrorCode,
    ExcelError,
    Fill,
    FillPattern,
    Font,
    FormulaSyntaxError,
    FormulonError,
    GradientFill,
    GradientStop,
    IconSet,
    IconSetName,
    InvalidArgumentError,
    InvalidFileError,
    ParseErrorCode,
    PivotInput,
    RangeCell,
    SheetRect,
    TableInput,
    ValidationType,
    Value,
    ValueKind,
    Workbook,
    WorkbookFormat,
    _inputs,
    enum_name,
    evaluate_formula,
    evaluate_typed_formula,
)

REPO_ROOT = Path(__file__).resolve().parents[3]
API_DIR = REPO_ROOT / "tools" / "api"
HEADER = REPO_ROOT / "src" / "c_api" / "formulon_c.h"
PYPROJECT = REPO_ROOT / "packages" / "python" / "pyproject.toml"


def _case_table(name: str) -> dict:
    return json.loads((API_DIR / name).read_text(encoding="utf-8"))


def _expected_native(value: object) -> object:
    """A case-table native value with `{error: <ErrorCode member>}` mapped to an `ExcelError`."""
    if isinstance(value, dict):
        return ExcelError(ErrorCode[value["error"].upper()])
    if isinstance(value, list):
        return [_expected_native(item) for item in value]
    return float(value) if isinstance(value, int) and not isinstance(value, bool) else value


class NativeEvaluateTests(unittest.TestCase):
    def test_facade_cases_of_the_module_evaluate(self) -> None:
        rows = [c for c in _case_table("to_native_cases.json")["facade_cases"] if c["call"] == "evaluate_formula"]
        self.assertTrue(rows)
        for case in rows:
            with self.subTest(case["id"]):
                self.assertEqual(evaluate_formula(case["formula"]), _expected_native(case["native"]))

    def test_scalar_rows_of_to_native(self) -> None:
        for case in _case_table("to_native_cases.json")["cases"]:
            if "formula" not in case:
                continue
            with self.subTest(case["formula"]):
                self.assertEqual(evaluate_formula(case["formula"]), _expected_native(case["native"]))

    def test_typed_evaluate_keeps_the_kind_and_the_array(self) -> None:
        self.assertEqual(evaluate_typed_formula("=1+1"), Value(kind=ValueKind.NUMBER, number=2.0))
        grid = evaluate_typed_formula('={1,"a";TRUE,#N/A}')
        self.assertIs(grid.kind, ValueKind.ARRAY)
        self.assertEqual(len(grid.array), 2)
        self.assertEqual(grid.array[0][1], Value(kind=ValueKind.TEXT, text="a"))
        self.assertEqual(formulon.to_native(grid), [[1.0, "a"], [True, ExcelError(ErrorCode.NA)]])
        self.assertEqual(repr(grid), "Value(ARRAY, 2x2)")

    def test_a_non_string_formula_is_a_type_error(self) -> None:
        with self.assertRaises(TypeError):
            evaluate_formula(1)  # type: ignore[arg-type]

    def test_an_array_value_must_be_a_rectangle_of_scalars(self) -> None:
        one = Value(kind=ValueKind.NUMBER, number=1.0)
        with self.assertRaises(ValueError):
            Value(kind=ValueKind.ARRAY)
        with self.assertRaises(ValueError):
            Value(kind=ValueKind.ARRAY, array=[[one], [one, one]])
        with self.assertRaises(TypeError):
            Value(kind=ValueKind.ARRAY, array=[[1.0]])
        self.assertEqual(Value(kind=ValueKind.ARRAY, array=[[one]]).array, ((one,),))

    def test_range_cell_carries_a_native_value_and_its_kind(self) -> None:
        cell = RangeCell(row=0, col=1, formula=None, value=2.0, kind=ValueKind.NUMBER)
        self.assertEqual((cell.value, cell.kind), (2.0, ValueKind.NUMBER))
        self.assertNotIn("Cell", formulon.__all__)


class SyntaxErrorTests(unittest.TestCase):
    def test_evaluate_rows_of_the_parse_error_table(self) -> None:
        rows = [c for c in _case_table("parse_error_cases.json")["cases"] if c["entry"] == "evaluate_formula"]
        self.assertTrue(rows)
        for case in rows:
            with self.subTest(case["id"]):
                with self.assertRaises(FormulaSyntaxError) as caught:
                    evaluate_formula(case["formula"])
                error = caught.exception
                self.assertEqual(int(error.code), case["status"])
                self.assertIn(
                    error.parse_error.code.name, {re.sub(r"(?<!^)(?=[A-Z])", "_", c).upper() for c in case["codes"]}
                )
                if case["start"] is not None:
                    self.assertEqual(
                        (
                            error.parse_error.start,
                            error.parse_error.end,
                            error.parse_error.line,
                            error.parse_error.column,
                        ),
                        (case["start"], case["end"], case["line"], case["column"]),
                    )

    def test_a_syntax_error_is_an_invalid_argument(self) -> None:
        with self.assertRaises(InvalidArgumentError):
            evaluate_formula("=(1+")
        self.assertIs(ParseErrorCode.UNEXPECTED_EOF, formulon.ParseErrorCode.UNEXPECTED_EOF)


class ExceptionTests(unittest.TestCase):
    def test_bytes_that_are_not_a_package_raise_invalid_file(self) -> None:
        with self.assertRaises(InvalidFileError) as caught:
            Workbook.load(b"not a workbook")
        self.assertEqual(int(caught.exception.code), 5003)
        self.assertIsInstance(caught.exception, FormulonError)

    def test_public_exceptions_report_the_package(self) -> None:
        for cls in (
            FormulonError,
            InvalidArgumentError,
            FormulaSyntaxError,
            InvalidFileError,
            formulon.NotFoundError,
            formulon.IndexOutOfRangeError,
            formulon.WorkbookClosedError,
        ):
            with self.subTest(cls.__name__):
                self.assertEqual(cls.__module__, "formulon")
                self.assertIs(getattr(formulon, cls.__name__), cls)


def _header_enum(c_type: str) -> dict:
    text = HEADER.read_text(encoding="utf-8")
    body = re.search(r"typedef enum \{([^}]*)\} " + re.escape(c_type) + ";", text)
    assert body is not None, c_type
    return {m[0]: int(m[1]) for m in re.findall(r"^\s*(FM_\w+) = (\d+)", body.group(1), re.M)}


class EnumTests(unittest.TestCase):
    TABLES = (
        (ValidationType, "fm_validation_type_t", "FM_VALIDATION_TYPE_"),
        (ComparisonOperator, "fm_comparison_operator_t", "FM_COMPARISON_OPERATOR_"),
        (CfRuleType, "fm_cf_rule_type_t", "FM_CF_RULE_TYPE_"),
        (DatePeriod, "fm_date_period_t", "FM_DATE_PERIOD_"),
        (CfValueObjectType, "fm_cf_value_object_type_t", "FM_CF_VALUE_OBJECT_TYPE_"),
        (IconSetName, "fm_icon_set_name_t", "FM_ICON_SET_NAME_"),
    )
    DIGITS = {"3": "THREE", "4": "FOUR", "5": "FIVE"}

    def test_values_match_the_header(self) -> None:
        for table, c_type, prefix in self.TABLES:
            with self.subTest(c_type):
                expected = {}
                for name, value in _header_enum(c_type).items():
                    member = name[len(prefix) :]
                    member = re.sub(r"^(\d)_", lambda m: self.DIGITS[m[1]] + "_", member)
                    expected[member] = value
                self.assertEqual({m.name: m.value for m in table}, expected)

    def test_enum_name(self) -> None:
        self.assertEqual(enum_name(CfRuleType, 3), "DATA_BAR")
        self.assertEqual(enum_name(IconSetName, IconSetName.FIVE_QUARTERS), "FIVE_QUARTERS")
        self.assertIsNone(enum_name(CfRuleType, 99))
        with self.assertRaises(TypeError):
            enum_name(int, 1)  # type: ignore[arg-type]

    def test_records_read_their_codes_as_members(self) -> None:
        self.assertIs(Cfvo(type=3).type, CfValueObjectType.MIN)
        self.assertIs(IconSet(name=16, thresholds=[]).name, IconSetName.FIVE_QUARTERS)


class StyleColourTests(unittest.TestCase):
    def test_font_colour_is_one_color_and_takes_text(self) -> None:
        self.assertNotIn("color_argb", {f.name for f in dataclasses.fields(Font)})
        with Workbook.create_default() as wb:
            index = wb.ooxml.add_font(Font(name="Arial", color="#FF0000"))
            self.assertEqual(wb.ooxml.get_font(index).color, Color(kind=ColorKind.RGB, argb=0xFFFF0000))
            theme = Color(kind=ColorKind.THEME, theme=4, tint=0.5)
            self.assertEqual(wb.ooxml.get_font(wb.ooxml.add_font(Font(name="Arial", color=theme))).color, theme)
            self.assertEqual(wb.ooxml.get_font(wb.ooxml.add_font(Font(name="Arial Black"))).color.kind, ColorKind.NONE)

    def test_fill_border_and_gradient_colours(self) -> None:
        self.assertFalse({"fg_argb", "bg_argb"} & {f.name for f in dataclasses.fields(Fill)})
        with Workbook.create_default() as wb:
            fill = wb.ooxml.get_fill(wb.ooxml.add_fill(Fill(pattern=FillPattern.SOLID, foreground="#00FF00")))
            self.assertEqual(fill.foreground, Color(kind=ColorKind.RGB, argb=0xFF00FF00))
            stops = [GradientStop(0.0, "#80112233"), GradientStop(1.0, Color(kind=ColorKind.THEME, theme=1))]
            gradient = wb.ooxml.get_fill(wb.ooxml.add_fill(Fill(gradient=GradientFill(stops=stops)))).gradient
            self.assertEqual([s.color for s in gradient.stops], [Color(ColorKind.RGB, argb=0x80112233), stops[1].color])
            side = BorderSideRecord(style=formulon.BorderStyle.THIN, color="#0000FF")
            border = wb.ooxml.get_border(wb.ooxml.add_border(BorderRecord(left=side)))
            self.assertEqual(border.left.color, Color(kind=ColorKind.RGB, argb=0xFF0000FF))

    def test_a_malformed_colour_is_a_type_error(self) -> None:
        with Workbook.create_default() as wb:
            for bad in ("red", "#12345", 0xFF0000):
                with self.subTest(bad):
                    with self.assertRaises(TypeError):
                        wb.ooxml.add_font(Font(name="Arial", color=bad))  # type: ignore[arg-type]

    def test_colour_input_helpers(self) -> None:
        self.assertEqual(_inputs.color_input("#abcdef", "c"), Color(kind=ColorKind.RGB, argb=0xFFABCDEF))
        self.assertEqual(_inputs.optional_color_input(None, "c"), Color())


class RecordA1Tests(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.wb.add_sheet("Data")

    def tearDown(self) -> None:
        self.wb.close()

    def test_rect_list_field_takes_strings_and_one_string(self) -> None:
        record = ConditionalFormatInput(ranges="B2:C3", type=CfRuleType.EXPRESSION, formula1="=TRUE")
        normalized = _inputs.normalize_record_ranges(record, self.wb)
        self.assertEqual(normalized.ranges, [CellRect(1, 1, 2, 2)])
        mixed = DataValidationInput(type=ValidationType.WHOLE, ranges=["A1", CellRect(4, 0, 5, 0)])
        self.assertEqual(
            _inputs.normalize_record_ranges(mixed, self.wb).ranges, [CellRect(0, 0, 0, 0), CellRect(4, 0, 5, 0)]
        )
        self.assertEqual(_inputs.normalize_record_ranges(dataclasses.replace(mixed, ranges=[]), self.wb).ranges, [])

    def test_grid_edges(self) -> None:
        record = DataValidationInput(type=ValidationType.WHOLE, ranges="XFD1048576")
        self.assertEqual(
            _inputs.normalize_record_ranges(record, self.wb).ranges, [CellRect(1048575, 16383, 1048575, 16383)]
        )
        for bad in ("XFE1", "", "A1:B2 C3", "Data!A1"):
            with self.subTest(bad):
                with self.assertRaises(InvalidArgumentError):
                    _inputs.normalize_record_ranges(dataclasses.replace(record, ranges=bad), self.wb)
        with self.assertRaises(TypeError):
            _inputs.normalize_record_ranges(dataclasses.replace(record, ranges=[3]), self.wb)

    def test_cell_address_field(self) -> None:
        self.assertNotIn("anchor_row", {f.name for f in dataclasses.fields(PivotInput)})
        field_types = {"PivotInput": {"anchor": "CellAddress"}}
        original = dict(_inputs.RECORD_RANGE_FIELDS)
        _inputs.RECORD_RANGE_FIELDS.update(field_types)
        try:
            pivot = _inputs.normalize_record_ranges(PivotInput(name="P", cache_id=1, anchor="C3"), self.wb)
            self.assertEqual(pivot.anchor, CellAddress(row=2, col=2))
            with self.assertRaises(InvalidArgumentError):
                _inputs.normalize_record_ranges(PivotInput(name="P", cache_id=1, anchor="C3:D4"), self.wb)
        finally:
            _inputs.RECORD_RANGE_FIELDS.clear()
            _inputs.RECORD_RANGE_FIELDS.update(original)

    def test_sheet_rect_takes_a_qualified_range(self) -> None:
        self.assertEqual(_inputs.sheet_rect(self.wb, "Data!A1:B2", None, "r"), SheetRect(1, CellRect(0, 0, 1, 1)))
        self.assertEqual(_inputs.sheet_rect(self.wb, "A1", 1, "r"), SheetRect(1, CellRect(0, 0, 0, 0)))
        with self.assertRaises(InvalidArgumentError):
            _inputs.sheet_rect(self.wb, "A1", None, "r")

    def test_unlisted_record_is_returned_as_it_is(self) -> None:
        record = DefinedNameInput(name="X", formula="=1")
        self.assertIs(_inputs.normalize_record_ranges(record, self.wb), record)


class WorkbookEntryTests(unittest.TestCase):
    def test_defined_name_scope_and_formula_without_equals(self) -> None:
        with Workbook.create_default() as wb:
            sheet = wb.add_sheet("Rates")
            wb.set_defined_name(DefinedNameInput(name="Rate", formula="0.5", scope=sheet))
            wb.set_defined_name(DefinedNameInput(name="Base", formula="=2", scope="Sheet1"))
            wb.set_defined_name(DefinedNameInput(name="Global", formula="3"))
            self.assertEqual(wb.get_defined_name("Rate", "Rates").formula, "=0.5")
            self.assertEqual(wb.get_defined_name("Base", 0).scope, 0)
            self.assertIsNone(wb.get_defined_name("Global").scope)
            self.assertNotIn("local_sheet_id", {f.name for f in dataclasses.fields(DefinedNameInput)})

    def test_profile_id_accessors(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(wb.get_profile_id(), "win-365-en_US")
            wb.set_profile_id("mac-365-ja_JP")
            self.assertEqual(wb.get_profile_id(), "mac-365-ja_JP")
            self.assertFalse(hasattr(wb, "get_excel_profile_id"))
        self.assertFalse(hasattr(formulon, "ExcelProfileId"))

    def test_copy_sheet_appends_by_default(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(wb.copy_sheet(0).name, "Sheet1 (2)")
            self.assertEqual(wb.copy_sheet("Sheet1", None, "Copy").index, 2)
            self.assertEqual([s.name for s in wb], ["Sheet1", "Sheet1 (2)", "Copy"])

    def test_add_table_derives_columns_and_takes_a1(self) -> None:
        with Workbook.create_default() as wb:
            wb[0].set_text("A1", "Item")
            table = wb[0].add_table(TableInput(range="A1:B3", name="Orders"))
            self.assertEqual([c.name for c in table.get_properties().columns], ["Item", "Column1"])
            with self.assertRaises(InvalidArgumentError):
                wb[0].add_table(TableInput(range="D1:E3", name="T2"))

    def test_save_needs_a_known_extension_or_a_format(self) -> None:
        with Workbook.create_default() as wb, tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / "book.xlsm"
            with self.assertRaises(ValueError) as caught:
                wb.save(target)
            self.assertEqual(
                str(caught.exception),
                f"formulon: cannot infer the workbook format from {str(target)!r}; use a .xlsx or .xlsb extension "
                "or pass format=WorkbookFormat.XLSX or WorkbookFormat.XLSB",
            )
            self.assertFalse(target.exists())
            wb.save(target, format=WorkbookFormat.XLSX)
            self.assertTrue(target.read_bytes().startswith(b"PK"))
            with self.assertRaises(ValueError):
                wb.save_with_diagnostics(Path(tmp) / "book")
            self.assertTrue(wb.save().startswith(b"PK"))
            wb.save(Path(tmp) / "book.XLSB")

    def test_calc_mode_reads_are_documented_as_calculating(self) -> None:
        self.assertIn("calculate the", CalcMode.__doc__)


class DateHelperTests(unittest.TestCase):
    """The serial arithmetic of the shared date table, through the helpers the setters and `get_date` use."""

    @staticmethod
    def _write_input(spec: dict) -> object:
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
        return None

    def test_table_rows(self) -> None:
        errors = {"TypeError": TypeError, "ValueError": ValueError, "InvalidArgument": InvalidArgumentError}
        ran = 0
        for case in _case_table("date_value_cases.json")["cases"]:
            expect = case["expect"].get("py")
            if expect is None:
                continue
            spec = case["input"]
            with self.subTest(case["id"]):
                if case["direction"] == "write":
                    value = self._write_input(spec)
                    if value is None:
                        continue
                    call = lambda: _inputs.date_serial(value, date1904=case["date1904"])  # noqa: E731
                else:
                    if "serial" not in spec:
                        continue
                    serial = float(spec["serial"]) if isinstance(spec["serial"], str) else spec["serial"]
                    call = lambda: _inputs.serial_datetime(serial, date1904=case["date1904"])  # noqa: E731
                ran += 1
                if "error" in expect:
                    with self.assertRaises(errors[expect["error"]]):
                        call()
                elif case["direction"] == "write":
                    serial_out, code = call()
                    self.assertAlmostEqual(serial_out, expect["serial"], delta=abs(expect["serial"]) * 1e-12)
                    if case["effective_format"] == "general":
                        self.assertEqual(code, expect["format"])
                else:
                    d = expect["date"]
                    self.assertEqual(
                        call(),
                        datetime.datetime(
                            d["year"], d["month"], d["day"], d["hour"], d["minute"], d["second"], d["microsecond"]
                        ),
                    )
        self.assertGreater(ran, 40)


class TypingAndPackagingTests(unittest.TestCase):
    def test_workbook_hints_resolve_at_runtime(self) -> None:
        for name, member in vars(Workbook).items():
            func = getattr(member, "__func__", member)
            func = getattr(func, "fget", func)
            if callable(func):
                with self.subTest(name):
                    typing.get_type_hints(func)
        typing.get_type_hints(evaluate_formula)
        typing.get_type_hints(_inputs.normalize_record_ranges)

    def test_no_helper_leaks_into_the_package(self) -> None:
        for name in ("Path", "PackageNotFoundError", "CfColor"):
            self.assertFalse(hasattr(formulon, name), name)
        for name in formulon.__all__:
            self.assertTrue(hasattr(formulon, name), name)

    def test_project_metadata(self) -> None:
        text = PYPROJECT.read_text(encoding="utf-8")
        self.assertIn('requires-python = ">=3.10"', text)
        self.assertIn('"Programming Language :: Python :: 3.14"', text)
        self.assertIn('"Typing :: Typed"', text)
        self.assertNotIn("3.9", text)


if __name__ == "__main__":
    unittest.main()
