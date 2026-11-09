"""Argument errors: the class and status code each kind of bad argument raises.

A non-integer where an integer is required is a ``TypeError`` (every ``float``
and ``bool`` included); a negative position is an out-of-range failure; a
negative coordinate or quantity is an invalid argument; a value past the integer
width is a ``ValueError``. Only classes and codes are asserted, never message text.
"""

from __future__ import annotations

import copy
import unittest

import _chart_input

import formulon
from formulon import (
    CellAddress,
    CellRect,
    ChartLegendPosition,
    FormulaSyntaxError,
    FormulonError,
    GeometryMode,
    IndexOutOfRangeError,
    InvalidArgumentError,
    PivotAggregation,
    PivotAxis,
    PivotDataFieldInput,
    PivotFieldInput,
    PivotInput,
    ShapeInput,
    StatusCode,
    TableInput,
    Workbook,
)

_PROFILE = "mac-365-de_DE"


class SheetReferenceTests(unittest.TestCase):
    def test_negative_sheet_index_is_a_sheet_index_failure(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(IndexOutOfRangeError) as caught:
                wb.sheet(-1)
            self.assertEqual(caught.exception.code, StatusCode.SHEET_INDEX_OUT_OF_RANGE)

    def test_past_the_end_sheet_index_is_a_sheet_index_failure(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(IndexOutOfRangeError) as caught:
                wb.sheet(5)
            self.assertEqual(caught.exception.code, StatusCode.SHEET_INDEX_OUT_OF_RANGE)

    def test_non_integers_are_type_errors(self) -> None:
        with Workbook.create_default() as wb:
            for bad in (1.5, 1.0, float("nan"), True, False, None, b"0"):
                with self.subTest(bad=bad):
                    with self.assertRaises(TypeError):
                        wb.sheet(bad)  # type: ignore[arg-type]

    def test_index_past_the_integer_width_is_a_value_error(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(ValueError) as caught:
                wb.sheet(2**64)
            self.assertNotIsInstance(caught.exception, FormulonError)

    def test_unknown_sheet_name_is_not_found(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(formulon.NotFoundError) as caught:
                wb.sheet("missing")
            self.assertEqual(caught.exception.code, StatusCode.NOT_FOUND)


class CoordinateTests(unittest.TestCase):
    def test_negative_coordinates_are_invalid_arguments(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            for row, col in ((-1, 0), (0, -1)):
                with self.subTest(row=row, col=col):
                    with self.assertRaises(InvalidArgumentError) as caught:
                        ws.get_value(row, col)
                    self.assertEqual(caught.exception.code, StatusCode.INVALID_ARGUMENT)

    def test_a_coordinate_the_grid_cannot_hold_is_an_invalid_argument(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            for row in (2**31, 2**32 + 5):
                with self.subTest(row=row):
                    with self.assertRaises(InvalidArgumentError):
                        ws.get_value(row, 0)

    def test_a_coordinate_past_the_integer_width_is_refused(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(ValueError):
                wb.sheet(0).get_value(0, 2**64)

    def test_non_integer_coordinates_are_type_errors(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            for bad in (1.5, 2.0, True, "1", None):
                with self.subTest(bad=bad):
                    with self.assertRaises(TypeError):
                        ws.get_value(bad, 0)  # type: ignore[arg-type]

    def test_an_integer_like_object_is_accepted(self) -> None:
        class Index:
            def __index__(self) -> int:
                return 0

        with Workbook.create_default() as wb:
            wb.sheet(0).set_number(0, 0, 4.0)
            self.assertEqual(wb.sheet(0).get_typed_value(Index(), 0).number, 4.0)


class PositionTests(unittest.TestCase):
    @staticmethod
    def _table(wb: Workbook):
        ws = wb.sheet(0)
        ws.set_text(0, 0, "Product")
        ws.set_text(0, 1, "Amount")
        return ws.add_table(TableInput(range=CellRect(0, 0, 1, 1), name="Tbl", columns=["Product", "Amount"]))

    def test_negative_and_past_the_end_table_positions_are_out_of_range(self) -> None:
        with Workbook.create_default() as wb:
            self._table(wb)
            for position in (-1, 1):
                with self.subTest(position=position):
                    with self.assertRaises(IndexOutOfRangeError) as caught:
                        wb.sheet(0).table(position)
                    self.assertEqual(caught.exception.code, StatusCode.INDEX_OUT_OF_RANGE)

    def test_table_column_position(self) -> None:
        with Workbook.create_default() as wb:
            table = self._table(wb)
            with self.assertRaises(IndexOutOfRangeError) as caught:
                table.rename_column(-1, "x")
            self.assertEqual(caught.exception.code, StatusCode.INDEX_OUT_OF_RANGE)
            with self.assertRaises(TypeError):
                table.rename_column(1.0, "x")  # type: ignore[arg-type]

    def test_negative_pivot_position_is_out_of_range(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(IndexOutOfRangeError) as caught:
                wb.sheet(0).pivot(-1)
            self.assertEqual(caught.exception.code, StatusCode.INDEX_OUT_OF_RANGE)


class GeometryModeTests(unittest.TestCase):
    def test_an_integer_outside_the_enum_is_an_invalid_argument(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            for bad in (5, -1, 99):
                with self.subTest(mode=bad):
                    with self.assertRaises(InvalidArgumentError) as caught:
                        ws.get_column_width_pt(0, bad)  # type: ignore[arg-type]
                    self.assertEqual(caught.exception.code, StatusCode.INVALID_ARGUMENT)

    def test_a_non_integer_mode_is_a_type_error(self) -> None:
        with Workbook.create_default() as wb:
            for bad in ("1", 1.0, None, True):
                with self.subTest(mode=bad):
                    with self.assertRaises(TypeError):
                        wb.sheet(0).get_column_width_pt(0, bad)  # type: ignore[arg-type]

    def test_every_member_is_accepted(self) -> None:
        with Workbook.create_default() as wb:
            for mode in GeometryMode:
                with self.subTest(mode=mode.name):
                    self.assertGreaterEqual(wb.sheet(0).get_column_width_pt(0, mode), 0.0)


class SetValuesShapeTests(unittest.TestCase):
    def test_ragged_rows_are_an_invalid_argument(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(InvalidArgumentError) as caught:
                wb.sheet(0).set_values("A1", [[1, 2], [3]])
            self.assertEqual(caught.exception.code, StatusCode.INVALID_ARGUMENT)
            self.assertIsNone(wb.sheet(0).get_value("A1"))

    def test_a_grid_past_the_sheet_edge_is_an_invalid_argument(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(InvalidArgumentError):
                wb.sheet(0).set_values(1048575, 0, [[1], [2]])

    def test_a_non_sequence_grid_or_row_is_a_type_error(self) -> None:
        with Workbook.create_default() as wb:
            for bad in (5, [1, 2]):
                with self.subTest(grid=bad):
                    with self.assertRaises(TypeError):
                        wb.sheet(0).set_values("A1", bad)  # type: ignore[arg-type]


class ParseReferenceBaseTests(unittest.TestCase):
    def test_one_base_coordinate_alone_is_a_type_error(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(TypeError):
                wb.parse_reference("RC", base_row=1)
            with self.assertRaises(TypeError):
                wb.parse_reference("RC", base_col=1)

    def test_both_base_coordinates_are_accepted(self) -> None:
        with Workbook.create_default() as wb:
            ref = wb.parse_reference("R[-1]C", r1c1=True, base_row=5, base_col=5)
            self.assertEqual(ref.range, CellRect(4, 5, 4, 5))


class FormulaLocaleValidationTests(unittest.TestCase):
    def test_canonicalize_reports_a_syntax_error_in_the_localized_text(self) -> None:
        text = "=SUMME(1;"
        with self.assertRaises(FormulaSyntaxError) as caught:
            formulon.canonicalize_formula(text, _PROFILE)
        self.assertGreaterEqual(int(caught.exception.code), 1000)
        self.assertLessEqual(int(caught.exception.code), 1014)
        record = caught.exception.parse_error
        self.assertIsNotNone(record)
        self.assertEqual(record.end, len(text))

    def test_canonicalize_locates_an_error_in_the_caller_text(self) -> None:
        text = '=WENN(A1>0;SUMME(B1:B3;"x);0)'
        with self.assertRaises(FormulaSyntaxError) as caught:
            formulon.canonicalize_formula(text, _PROFILE)
        record = caught.exception.parse_error
        self.assertIsNotNone(record)
        self.assertEqual(record.start, 23)
        self.assertLessEqual(record.start, record.end)
        self.assertLessEqual(record.end, len(text))

    def test_localize_reports_a_syntax_error_in_the_canonical_text(self) -> None:
        text = "=SUM(1,"
        with self.assertRaises(FormulaSyntaxError) as caught:
            formulon.localize_formula(text, _PROFILE)
        record = caught.exception.parse_error
        self.assertIsNotNone(record)
        self.assertEqual(record.end, len(text))

    def test_a_valid_formula_round_trips(self) -> None:
        local = formulon.localize_formula("=SUM(1.5,2)", _PROFILE)
        self.assertEqual(local, "=SUMME(1,5;2)")
        self.assertEqual(formulon.canonicalize_formula(local, _PROFILE), "=SUM(1.5,2)")


def _add_pivot(wb: Workbook):
    ws = wb.sheet(0)
    ws.set_text(0, 0, "Region")
    ws.set_text(0, 1, "Amount")
    ws.set_text(1, 0, "East")
    ws.set_number(1, 1, 10)
    ws.set_text(2, 0, "West")
    ws.set_number(2, 1, 30)
    cache = wb.add_pivot_cache()
    cache.set_worksheet_source(formulon.PivotWorksheetSourceInput(range="A1:B3", sheet="Sheet1"))
    cache.add_field("Region")
    cache.add_field("Amount")
    for region, amount in (("East", 10), ("West", 30)):
        record = cache.add_record()
        cache.set_record_value(record, 0, region)
        cache.set_record_value(record, 1, amount)
    pivot = ws.add_pivot(PivotInput(name="P1", cache_id=cache.cache_id, anchor=CellAddress(0, 4)))
    pivot.add_field(PivotFieldInput(field_index="Region", axis=PivotAxis.ROW))
    amount = pivot.add_field(PivotFieldInput(field_index="Amount", axis=PivotAxis.VALUE))
    pivot.add_data_field(PivotDataFieldInput(name="Sum", field_index=amount, aggregation=PivotAggregation.SUM))
    return cache, pivot


class HandleCopyTests(unittest.TestCase):
    def _handles(self, wb: Workbook) -> dict:
        ws = wb.sheet(0)
        cache, pivot = _add_pivot(wb)
        table_sheet = wb.add_sheet("T")
        table_sheet.set_text(0, 0, "Product")
        table_sheet.set_text(0, 1, "Amount")
        table = table_sheet.add_table(TableInput(range=CellRect(0, 0, 1, 1), name="Tbl", columns=["Product", "Amount"]))
        chart = _chart_input.add_chart(ws, _chart_input.COLUMN, "Sheet1!$A$1:$B$3", legend=ChartLegendPosition.RIGHT)
        shape = ws.add_shape(
            ShapeInput(
                preset_geometry="rect",
                anchor=formulon.DrawingAnchorInput(
                    kind=formulon.AnchorKind.ONE_CELL, from_row=2, width_emu=914400, height_emu=457200
                ),
            )
        )
        return {
            "Workbook": wb,
            "WorkbookOoxml": wb.ooxml,
            "Sheet": ws,
            "SheetOoxml": ws.ooxml,
            "Table": table,
            "Pivot": pivot,
            "PivotCache": cache,
            "Chart": chart,
            "Shape": shape,
        }

    def test_copy_deepcopy_and_serialization_are_type_errors(self) -> None:
        with Workbook.create_default() as wb:
            for name, handle in self._handles(wb).items():
                with self.subTest(handle=name):
                    with self.assertRaises(TypeError):
                        copy.copy(handle)
                    with self.assertRaises(TypeError):
                        copy.deepcopy(handle)
                    # The hook the stdlib serializer calls, at every protocol it supports.
                    for protocol in range(6):
                        with self.assertRaises(TypeError):
                            handle.__reduce_ex__(protocol)

    def test_the_original_stays_usable_after_a_refused_copy(self) -> None:
        with Workbook.create_default() as wb:
            handles = self._handles(wb)
            for handle in handles.values():
                with self.assertRaises(TypeError):
                    copy.copy(handle)
                with self.assertRaises(TypeError):
                    copy.deepcopy(handle)
            self.assertFalse(wb.closed)
            self.assertEqual(wb.sheet(0).get_value(0, 0), "Region")
            self.assertEqual(handles["Table"].name, "Tbl")

    def test_the_refusal_names_the_class_and_the_supported_route(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(TypeError) as caught:
                copy.copy(wb)
            self.assertIn("Workbook", str(caught.exception))
            self.assertIn("save", str(caught.exception))

    def test_save_and_load_is_the_supported_copy(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_number(0, 0, 9.0)
            with Workbook.load(wb.save()) as other:
                self.assertEqual(other.sheet(0).get_typed_value(0, 0).number, 9.0)
            self.assertFalse(wb.closed)


if __name__ == "__main__":
    unittest.main()
