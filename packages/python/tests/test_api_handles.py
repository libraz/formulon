"""Handle identity (twelve cases), handle references, and the edge rows of the lower-object API.

Covers the print getters, pivot number formats, threaded-comment ``done`` and sparkline replacement.
"""

from __future__ import annotations

import io
import re
import unittest
import zipfile
from typing import Any, Tuple

import formulon
from formulon import (
    CellAddress,
    CellRect,
    ChartInsertKind,
    ChartLegendPosition,
    DrawingObjectKind,
    FormulonError,
    IndexOutOfRangeError,
    InvalidArgumentError,
    NotFoundError,
    Person,
    Pivot,
    PivotAggregation,
    PivotAxis,
    PivotCache,
    PivotDataFieldInput,
    PivotFieldInput,
    PivotInput,
    ShapeInput,
    Sheet,
    SlicerInput,
    Sparkline,
    SparklineGroup,
    StatusCode,
    Table,
    TableInput,
    ThreadedCommentInput,
    Workbook,
)

# ---- Fixtures ---------------------------------------------------------------------


def fill_source(sheet: Sheet) -> None:
    """Fill ``sheet`` with the Region / Amount table (A1:B3) the caches read."""
    sheet.set_text(0, 0, "Region")
    sheet.set_text(0, 1, "Amount")
    sheet.set_text(1, 0, "East")
    sheet.set_number(1, 1, 10)
    sheet.set_text(2, 0, "West")
    sheet.set_number(2, 1, 30)


def add_cache(wb: Workbook, source_sheet: str) -> PivotCache:
    """A cache over ``source_sheet``'s A1:B3 with the two records filled in."""
    cache = wb.add_pivot_cache()
    cache.set_worksheet_source(ref="A1:B3", sheet=source_sheet)
    cache.add_field("Region")
    cache.add_field("Amount")
    for region, amount in (("East", 10), ("West", 30)):
        record = cache.add_record()
        cache.set_record_value(record, 0, region)
        cache.set_record_value(record, 1, amount)
    return cache


def add_pivot(sheet: Sheet, cache: PivotCache, name: str, anchor_row: int, anchor_col: int) -> Tuple[Pivot, int]:
    """A pivot of ``cache`` on ``sheet``: Region on rows, Sum of Amount as the data field."""
    pivot = sheet.add_pivot(PivotInput(name=name, cache_id=cache.cache_id, anchor=CellAddress(anchor_row, anchor_col)))
    pivot.add_field(PivotFieldInput(source_name="Region", axis=PivotAxis.ROW))
    amount = pivot.add_field(PivotFieldInput(source_name="Amount", axis=PivotAxis.VALUE))
    pivot.add_data_field(
        PivotDataFieldInput(name="Sum of Amount", field_index=amount, aggregation=PivotAggregation.SUM)
    )
    return pivot, amount


def add_table(wb: Workbook, sheet: Sheet, name: str, rect: CellRect = CellRect(0, 0, 1, 1)) -> Table:
    """A two-column table over ``rect`` on ``sheet``."""
    sheet.set_text(rect.first_row, rect.first_col, "Product")
    sheet.set_text(rect.first_row, rect.first_col + 1, "Amount")
    sheet.set_number(rect.first_row + 1, rect.first_col + 1, 5)
    return sheet.add_table(TableInput(range=rect, name=name, columns=["Product", "Amount"]))


def add_chart(sheet: Sheet) -> formulon.Chart:
    """A column chart over Sheet1!A1:B3."""
    return sheet.add_chart(ChartInsertKind.COLUMN, "Sheet1!$A$1:$B$3", legend=ChartLegendPosition.RIGHT)


def add_shape(sheet: Sheet) -> formulon.Shape:
    return sheet.add_shape(
        ShapeInput(
            preset_geometry="rect",
            anchor_kind=formulon.AnchorKind.ONE_CELL,
            from_row=2,
            width_emu=914400,
            height_emu=457200,
        )
    )


class HandleTestCase(unittest.TestCase):
    def assertCode(self, exception: type, call: Any, code: StatusCode) -> None:
        with self.assertRaises(exception) as caught:
            call()
        self.assertEqual(caught.exception.code, code)


# ---- Identity: twelve cases ---------------------------------------------------------


class IdentityTests(HandleTestCase):
    def test_1_every_way_to_reach_one_target_returns_the_same_handle(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            fill_source(ws)
            table_sheet = wb.add_sheet("T")
            table = add_table(wb, table_sheet, "Tbl")
            self.assertIs(wb.table("Tbl"), table)
            self.assertIs(table_sheet.table("Tbl"), table)
            self.assertIs(table_sheet.table(0), table)
            self.assertIs(table_sheet.table(table), table)
            self.assertIs(table_sheet.list_tables()[0], table)
            self.assertIs(table.workbook, wb)
            self.assertIs(table.sheet, table_sheet)

            cache = add_cache(wb, "Sheet1")
            pivot, _ = add_pivot(ws, cache, "P1", 0, 4)
            self.assertIs(ws.pivot("P1"), pivot)
            self.assertIs(ws.pivot(0), pivot)
            self.assertIs(ws.pivot(pivot), pivot)
            self.assertIs(ws.list_pivots()[0], pivot)
            self.assertIs(pivot.get_cache(), cache)
            self.assertIs(wb.pivot_cache(cache.cache_id), cache)
            self.assertIs(wb.pivot_cache(cache), cache)
            self.assertIs(wb.list_pivot_caches()[0], cache)

            chart = add_chart(ws)
            self.assertIs(ws.chart(chart.object_id), chart)
            self.assertIs(ws.drawing(chart.object_id), chart)
            self.assertIs(ws.drawing(chart), chart)
            self.assertIs(ws.list_charts()[0], chart)
            self.assertIs(ws.list_drawings()[0], chart)
            shape = add_shape(ws)
            self.assertIs(ws.shape(shape.object_id), shape)
            self.assertIs(ws.drawing(shape.object_id), shape)
            self.assertIs(ws.list_shapes()[0], shape)
            self.assertIs(ws.list_drawings()[-1], shape)
            self.assertEqual(chart.kind, DrawingObjectKind.CHART)
            self.assertEqual(shape.kind, DrawingObjectKind.SHAPE)

    def test_2_removing_an_earlier_table_or_pivot_leaves_the_handle_on_its_own_target(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            fill_source(wb.add_sheet("Data"))
            a = add_table(wb, ws, "TblA", CellRect(0, 0, 1, 1))
            b = add_table(wb, ws, "TblB", CellRect(5, 0, 6, 1))
            c = add_table(wb, ws, "TblC", CellRect(10, 0, 11, 1))
            self.assertEqual([a.index, b.index, c.index], [0, 1, 2])
            ws.remove_table(a)
            self.assertEqual(b.index, 0)
            self.assertEqual(c.index, 1)
            self.assertEqual(c.name, "TblC")
            self.assertIs(wb.table("TblC"), c)
            self.assertIs(ws.table(0), b)

            cache = add_cache(wb, "Data")
            out = wb.add_sheet("Out")
            p1, _ = add_pivot(out, cache, "P1", 0, 0)
            p2, _ = add_pivot(out, cache, "P2", 20, 0)
            p3, _ = add_pivot(out, cache, "P3", 40, 0)
            out.remove_pivot(p1)
            self.assertEqual(p2.name, "P2")
            self.assertEqual(p2.index, 0)
            self.assertEqual(p3.index, 1)
            self.assertIs(out.pivot(1), p3)
            self.assertEqual(p3.get_properties().name, "P3")

    def test_3_the_handle_of_a_removed_target_raises_not_found(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            fill_source(ws)
            table_sheet = wb.add_sheet("T")
            table = add_table(wb, table_sheet, "Tbl")
            table_sheet.remove_table(table)
            for read in (lambda: table.name, lambda: table.index, table.get_properties, lambda: wb.table("Tbl")):
                self.assertCode(NotFoundError, read, StatusCode.NOT_FOUND)

            cache = add_cache(wb, "Sheet1")
            pivot, _ = add_pivot(ws, cache, "P1", 0, 4)
            ws.remove_pivot(pivot)
            for read in (
                lambda: pivot.name,
                pivot.get_properties,
                lambda: pivot.add_field(PivotFieldInput(source_name="Region", axis=PivotAxis.ROW)),
                lambda: ws.pivot("P1"),
            ):
                self.assertCode(NotFoundError, read, StatusCode.NOT_FOUND)

            chart = add_chart(ws)
            shape = add_shape(ws)
            ws.remove_drawing(chart)
            ws.remove_drawing(shape.object_id)
            for read in (
                chart.get_drawing_properties,
                chart.get_properties,
                shape.get_properties,
                lambda: ws.drawing(chart.object_id),
            ):
                self.assertCode(NotFoundError, read, StatusCode.NOT_FOUND)

            spare = wb.add_pivot_cache()
            wb.remove_pivot_cache(spare)
            self.assertCode(NotFoundError, spare.get_options, StatusCode.NOT_FOUND)
            self.assertCode(NotFoundError, lambda: wb.pivot_cache(spare.cache_id), StatusCode.NOT_FOUND)

    def test_4_removing_a_sheet_orphans_its_pivot_and_drawing_handles(self) -> None:
        with Workbook.create_default() as wb:
            fill_source(wb.sheet(0))
            out = wb.add_sheet("Out")
            keep = wb.add_sheet("Keep")
            cache = add_cache(wb, "Sheet1")
            pivot, _ = add_pivot(out, cache, "P1", 0, 0)
            kept_pivot, _ = add_pivot(keep, cache, "P2", 0, 0)
            chart = add_chart(out)
            kept_chart = add_chart(keep)
            wb.remove_sheet(out)
            self.assertCode(NotFoundError, pivot.get_properties, StatusCode.NOT_FOUND)
            self.assertCode(NotFoundError, chart.get_drawing_properties, StatusCode.NOT_FOUND)
            self.assertEqual(kept_pivot.name, "P2")
            self.assertIs(kept_pivot.sheet, keep)
            self.assertEqual(kept_chart.get_drawing_properties().kind, DrawingObjectKind.CHART)

    def test_5_deleting_the_rows_of_a_pivot_report_orphans_it_and_no_other_pivot(self) -> None:
        with Workbook.create_default() as wb:
            fill_source(wb.add_sheet("Data"))
            ws = wb.sheet(0)
            cache = add_cache(wb, "Data")
            p1, _ = add_pivot(ws, cache, "P1", 0, 0)
            p2, _ = add_pivot(ws, cache, "P2", 20, 0)
            ws.delete_rows("1:10")
            self.assertCode(NotFoundError, p1.get_properties, StatusCode.NOT_FOUND)
            self.assertEqual(p2.get_properties().name, "P2")
            self.assertEqual(p2.index, 0)
            self.assertIs(ws.list_pivots()[0], p2)
            self.assertIs(ws.pivot("P2"), p2)

    def test_6_removing_a_source_sheet_orphans_the_pivots_and_the_cache_that_read_it(self) -> None:
        with Workbook.create_default() as wb:
            out = wb.sheet(0)
            data_a = wb.add_sheet("DataA")
            data_b = wb.add_sheet("DataB")
            fill_source(data_a)
            fill_source(data_b)
            cache_a = add_cache(wb, "DataA")
            cache_b = add_cache(wb, "DataB")
            pa, _ = add_pivot(out, cache_a, "PA", 0, 0)
            pb, _ = add_pivot(out, cache_b, "PB", 20, 0)
            wb.remove_sheet(data_a)
            self.assertCode(NotFoundError, pa.get_properties, StatusCode.NOT_FOUND)
            self.assertCode(NotFoundError, cache_a.get_options, StatusCode.NOT_FOUND)
            self.assertEqual(pb.get_properties().name, "PB")
            self.assertIs(pb.get_cache(), cache_b)
            self.assertEqual(cache_b.get_worksheet_source().sheet, "DataB")
            self.assertEqual(len(out.list_pivots()), 1)
            self.assertIs(out.list_pivots()[0], pb)

    def test_7_copy_sheet_keeps_the_existing_handles_and_a_sheet_with_a_slicer_is_not_copied(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            fill_source(ws)
            cache = add_cache(wb, "Sheet1")
            pivot, _ = add_pivot(ws, cache, "P1", 0, 4)
            chart = add_chart(ws)
            copy = wb.copy_sheet(ws, 1, "Copy")
            self.assertIsNot(copy, ws)
            self.assertIs(wb.sheet("Sheet1"), ws)
            # The original handles are undisturbed and the copy has its own.
            self.assertIs(ws.pivot("P1"), pivot)
            self.assertIs(ws.list_charts()[0], chart)
            self.assertEqual(pivot.get_properties().name, "P1")
            copied_pivot = copy.list_pivots()[0]
            self.assertIsNot(copied_pivot, pivot)
            self.assertIs(copied_pivot.sheet, copy)
            self.assertEqual(copied_pivot.name, "P1")
            self.assertIsNot(copy.list_charts()[0], chart)
            self.assertIs(copy.list_charts()[0].sheet, copy)

            # A slicer on the sheet makes the copy fail; the handles stay as they were.
            table_sheet = wb.add_sheet("Tables")
            table = add_table(wb, table_sheet, "Tbl")
            slicer = table_sheet.add_slicer(table, 0, SlicerInput(from_row=4, from_col=3, to_row=14, to_col=5))
            sheet_count = len(wb.list_sheets())
            self.assertCode(
                FormulonError, lambda: wb.copy_sheet(table_sheet, 0, "TablesCopy"), StatusCode.NOT_IMPLEMENTED
            )
            self.assertEqual(len(wb.list_sheets()), sheet_count)
            self.assertIs(table_sheet.list_slicers()[0], slicer)
            self.assertEqual(slicer.get_properties().table_name, "Tbl")
            self.assertIs(wb.table("Tbl"), table)

    def test_8_a_rename_leaves_the_handle_in_place(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            fill_source(ws)
            cache = add_cache(wb, "Sheet1")
            pivot, _ = add_pivot(ws, cache, "P1", 0, 4)
            pivot.rename("Q")
            self.assertEqual(pivot.name, "Q")
            self.assertIs(ws.pivot("Q"), pivot)
            self.assertCode(NotFoundError, lambda: ws.pivot("P1"), StatusCode.NOT_FOUND)

            table_sheet = wb.add_sheet("T")
            table = add_table(wb, table_sheet, "Tbl")
            table.rename("NewName")
            self.assertEqual(table.name, "NewName")
            self.assertIs(wb.table("NewName"), table)
            self.assertIs(table_sheet.table(0), table)

            chart = add_chart(ws)
            chart.rename("Revenue")
            self.assertIs(ws.chart(chart.object_id), chart)
            self.assertEqual(chart.get_drawing_properties().name, "Revenue")

            ws.rename("Renamed1")
            self.assertEqual(ws.name, "Renamed1")
            self.assertIs(wb.sheet("Renamed1"), ws)
            self.assertIs(wb.sheet(0), ws)
            self.assertIs(chart.sheet, ws)

    def test_9_restoring_a_drawing_under_its_old_id_revives_the_old_handle(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            fill_source(ws)
            chart = add_chart(ws)
            object_id = chart.object_id
            snapshot = chart.snapshot()
            ws.remove_drawing(chart)
            self.assertCode(NotFoundError, chart.get_properties, StatusCode.NOT_FOUND)
            restored = ws.restore_drawing(snapshot)
            self.assertIs(restored, chart)
            self.assertEqual(restored.object_id, object_id)
            self.assertEqual(chart.get_properties().object_id, object_id)
            self.assertIs(ws.chart(object_id), chart)

            # A cache id comes back the same way.
            cache = wb.add_pivot_cache(41)
            wb.remove_pivot_cache(cache)
            self.assertCode(NotFoundError, cache.get_options, StatusCode.NOT_FOUND)
            again = wb.add_pivot_cache(41)
            self.assertIs(again, cache)
            cache.get_options()

    def test_10_an_accessor_of_the_wrong_kind_is_invalid_argument(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            fill_source(ws)
            shape = add_shape(ws)
            chart = add_chart(ws)
            for access in (ws.chart, ws.picture, ws.slicer):
                self.assertCode(InvalidArgumentError, lambda: access(shape.object_id), StatusCode.INVALID_ARGUMENT)
            for access in (ws.shape, ws.picture):
                self.assertCode(InvalidArgumentError, lambda: access(chart.object_id), StatusCode.INVALID_ARGUMENT)
            # The right accessor still finds both.
            self.assertIs(ws.shape(shape.object_id), shape)
            self.assertIs(ws.chart(chart.object_id), chart)

    def test_11_a_handle_of_another_workbook_is_invalid_argument(self) -> None:
        with Workbook.create_default() as wb, Workbook.create_default() as other:
            ws = wb.sheet(0)
            os_ = other.sheet(0)
            fill_source(ws)
            fill_source(os_)
            foreign_table_sheet = other.add_sheet("T")
            foreign_table = add_table(other, foreign_table_sheet, "Tbl")
            foreign_cache = add_cache(other, "Sheet1")
            foreign_pivot, _ = add_pivot(os_, foreign_cache, "P1", 0, 4)
            foreign_chart = add_chart(os_)
            local_table_sheet = wb.add_sheet("T")
            local_table = add_table(wb, local_table_sheet, "Tbl")
            local_cache = add_cache(wb, "Sheet1")
            add_pivot(ws, local_cache, "P1", 0, 4)
            add_chart(ws)

            refused = (
                lambda: local_table_sheet.table(foreign_table),
                lambda: local_table_sheet.remove_table(foreign_table),
                lambda: ws.pivot(foreign_pivot),
                lambda: ws.remove_pivot(foreign_pivot),
                lambda: wb.pivot_cache(foreign_cache),
                lambda: wb.remove_pivot_cache(foreign_cache),
                lambda: ws.drawing(foreign_chart),
                lambda: ws.chart(foreign_chart),
                lambda: ws.remove_drawing(foreign_chart),
            )
            for call in refused:
                self.assertCode(InvalidArgumentError, call, StatusCode.INVALID_ARGUMENT)
            # A drawing of another sheet of the same workbook is refused too.
            second = wb.add_sheet("Second")
            self.assertCode(
                InvalidArgumentError, lambda: second.drawing(ws.list_charts()[0]), StatusCode.INVALID_ARGUMENT
            )
            # The local handles are unaffected.
            self.assertIs(local_table_sheet.table(0), local_table)

    def test_12_a_failed_pivot_edit_and_table_churn_leave_the_handles_on_their_targets(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            fill_source(wb.add_sheet("Data"))
            cache = add_cache(wb, "Data")
            # Each report is four rows tall: moving the first to row 5 reaches the second, which cannot be placed.
            first, _ = add_pivot(ws, cache, "P1", 0, 0)
            second, _ = add_pivot(ws, cache, "P2", 7, 0)
            self.assertCode(FormulonError, lambda: first.set_anchor("A5"), StatusCode.PRECONDITION_FAILED)
            self.assertEqual(first.name, "P1")
            self.assertIs(ws.pivot("P1"), first)
            self.assertIs(ws.list_pivots()[0], first)
            self.assertIs(ws.list_pivots()[1], second)
            self.assertEqual(first.get_layout().rect.first_row, 0)
            self.assertEqual(second.get_layout().rect.first_row, 7)

            # Tables survive growth of the list, row deletion and sheet removal elsewhere, and removal of an earlier one.
            home = wb.add_sheet("Home")
            target = add_table(wb, home, "Target", CellRect(0, 0, 1, 1))
            early_sheet = wb.add_sheet("Early")
            early = add_table(wb, early_sheet, "Early")
            for i in range(12):
                add_table(wb, wb.add_sheet(f"S{i}"), f"Extra{i}")
            self.assertIs(wb.table("Target"), target)
            self.assertEqual(target.name, "Target")
            elsewhere = wb.add_sheet("Elsewhere")
            elsewhere.set_number(0, 0, 1)
            elsewhere.delete_rows("1:1")
            self.assertEqual(target.name, "Target")
            wb.remove_sheet(elsewhere)
            self.assertEqual(target.name, "Target")
            early_sheet.remove_table(early)
            self.assertEqual(target.name, "Target")
            self.assertEqual(target.get_properties().range.rect.last_col, 1)
            self.assertIs(next(t for t in home.list_tables() if t.name == "Target"), target)
            self.assertIs(home.list_tables()[target.index], target)


# ---- Handle references ----------------------------------------------------------------


class HandleRefTests(HandleTestCase):
    def test_indices_names_ids_and_their_rejected_spellings(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            fill_source(ws)
            table_sheet = wb.add_sheet("T")
            table = add_table(wb, table_sheet, "Tbl")
            cache = add_cache(wb, "Sheet1")
            pivot, _ = add_pivot(ws, cache, "P1", 0, 4)
            not_found, out_of_range = StatusCode.NOT_FOUND, StatusCode.INDEX_OUT_OF_RANGE

            # 0 is the first element; the last index is valid and the count is not.
            self.assertIs(table_sheet.table(0), table)
            self.assertIs(table_sheet.table(-0), table)
            self.assertIs(ws.pivot(0), pivot)
            self.assertIs(ws.pivot(-0), pivot)
            self.assertCode(IndexOutOfRangeError, lambda: table_sheet.table(1), out_of_range)
            self.assertCode(IndexOutOfRangeError, lambda: ws.pivot(1), out_of_range)
            # 2^31 is out of range, 2^32 is outside the integer type.
            self.assertCode(IndexOutOfRangeError, lambda: table_sheet.table(2**31), out_of_range)
            self.assertCode(IndexOutOfRangeError, lambda: ws.pivot(2**31), out_of_range)
            with self.assertRaises(ValueError):
                table_sheet.table(2**32)
            with self.assertRaises(ValueError):
                ws.drawing(2**32)
            with self.assertRaises(ValueError):
                wb.pivot_cache(2**32)
            # A negative index is a ValueError, like every other index; a float is not an index.
            for bad in (-1,):
                with self.subTest(bad=bad):
                    for access in (table_sheet.table, ws.pivot, wb.pivot_cache, ws.drawing):
                        with self.assertRaises(ValueError):
                            access(bad)
            for bad in (float("nan"), float("inf"), -float("inf"), 1.5):
                with self.subTest(bad=bad):
                    for access in (table_sheet.table, ws.pivot, wb.pivot_cache, ws.drawing):
                        with self.assertRaises(TypeError):
                            access(bad)
            # None and values of another type are type errors.
            for bad in (None, True, {}, []):
                with self.subTest(bad=bad):
                    for access in (table_sheet.table, ws.pivot, wb.pivot_cache, ws.drawing):
                        with self.assertRaises(TypeError):
                            access(bad)
            with self.assertRaises(TypeError):
                table_sheet.table()  # type: ignore[call-arg]
            with self.assertRaises(TypeError):
                ws.chart("x")  # type: ignore[arg-type]
            # An empty name matches nothing; a name matches case-insensitively.
            self.assertCode(NotFoundError, lambda: table_sheet.table(""), not_found)
            self.assertCode(NotFoundError, lambda: ws.pivot(""), not_found)
            self.assertIs(wb.table("tbl"), table)
            self.assertIs(ws.pivot("p1"), pivot)
            # An id nobody holds is NotFound.
            self.assertCode(NotFoundError, lambda: wb.pivot_cache(cache.cache_id + 100), not_found)
            self.assertCode(NotFoundError, lambda: ws.drawing(0), not_found)
            self.assertCode(NotFoundError, lambda: ws.drawing(987654), not_found)
            self.assertCode(NotFoundError, lambda: ws.remove_drawing(987654), not_found)

    def test_a_removed_drawing_cannot_be_removed_again(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            fill_source(ws)
            shape = add_shape(ws)
            ws.remove_drawing(shape)
            self.assertEqual(ws.list_drawings(), [])
            self.assertCode(NotFoundError, lambda: ws.remove_drawing(shape), StatusCode.NOT_FOUND)

    def test_handles_of_a_closed_workbook_raise_workbook_closed(self) -> None:
        wb = Workbook.create_default()
        fill_source(wb.sheet(0))
        table = add_table(wb, wb.add_sheet("T"), "Tbl")
        cache = add_cache(wb, "Sheet1")
        wb.close()
        with self.assertRaises(formulon.WorkbookClosedError):
            table.get_properties()
        with self.assertRaises(formulon.WorkbookClosedError):
            cache.get_options()
        self.assertIs(table.workbook, wb)


# ---- Print getters ----------------------------------------------------------------------


class PrintGetterTests(unittest.TestCase):
    def test_fit_to_page_false_removes_the_attribute_and_the_getters_report_the_stated_flags(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            self.assertFalse(ws.get_page_setup().fit_to_page_stated)
            ws.update_page_setup(fit_to_page=True)
            self.assertTrue(ws.get_page_setup().fit_to_page)
            self.assertTrue(ws.get_page_setup().fit_to_page_stated)
            ws.update_page_setup(fit_to_page=False)
            self.assertFalse(ws.get_page_setup().fit_to_page)
            self.assertFalse(ws.get_page_setup().fit_to_page_stated)

            # The print options and header / footer getters return every field, stated or not.
            self.assertEqual(
                ws.get_print_options(),
                formulon.PrintOptions(
                    grid_lines=False,
                    headings=False,
                    horizontal_centered=False,
                    vertical_centered=False,
                    grid_lines_stated=False,
                    headings_stated=False,
                    horizontal_centered_stated=False,
                    vertical_centered_stated=False,
                ),
            )
            ws.update_print_options(grid_lines=True, horizontal_centered=True)
            options = ws.get_print_options()
            self.assertTrue(options.grid_lines)
            self.assertTrue(options.grid_lines_stated)
            self.assertTrue(options.horizontal_centered)
            self.assertFalse(options.headings)
            self.assertFalse(options.headings_stated)

            empty = ws.get_header_footer()
            self.assertIsNone(empty.odd_header)
            self.assertFalse(empty.different_first)
            self.assertFalse(empty.different_first_stated)
            ws.update_header_footer(odd_header="&CReport", odd_footer="&P/&N", different_first=True)
            footer = ws.get_header_footer()
            self.assertEqual(footer.odd_header, "&CReport")
            self.assertEqual(footer.odd_footer, "&P/&N")
            self.assertIsNone(footer.even_header)
            self.assertTrue(footer.different_first)
            self.assertTrue(footer.different_first_stated)
            ws.update_header_footer(odd_header=None)
            self.assertIsNone(ws.get_header_footer().odd_header)
            self.assertEqual(ws.get_header_footer().odd_footer, "&P/&N")


# ---- Pivot number formats -------------------------------------------------------------------


def _pivot_part(wb: Workbook) -> str:
    with zipfile.ZipFile(io.BytesIO(wb.save())) as package:
        return package.read("xl/pivotTables/pivotTable1.xml").decode("utf-8")


class PivotNumberFormatTests(unittest.TestCase):
    def test_number_format_is_a_format_code_and_the_builtin_id_27_survives_a_round_trip(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_profile_id("win-365-ja_JP")
            ws = wb.sheet(0)
            fill_source(ws)
            cache = add_cache(wb, "Sheet1")
            pivot, amount = add_pivot(ws, cache, "P1", 0, 4)
            self.assertIsNone(pivot.list_fields()[amount].number_format)

            # The profile's code for id 27 maps back to id 27, and its read-back writes id 27 again.
            code27 = wb.ooxml.get_number_format(27)
            self.assertTrue(code27, "the profile has a code for numFmtId 27")
            pivot.update_field(amount, number_format=code27)
            self.assertRegex(_pivot_part(wb), r'<pivotField [^>]*numFmtId="27"')
            read_back = pivot.list_fields()[amount].number_format
            self.assertEqual(read_back, code27)
            pivot.update_field(amount, number_format=read_back)
            self.assertRegex(_pivot_part(wb), r'<pivotField [^>]*numFmtId="27"')
            self.assertEqual(pivot.list_fields()[amount].number_format, code27)

            # A builtin code gets its builtin id; a code that is not builtin gets a custom id.
            pivot.update_field(amount, number_format="#,##0.00")
            self.assertRegex(_pivot_part(wb), r'<pivotField [^>]*numFmtId="4"')
            self.assertEqual(pivot.list_fields()[amount].number_format, "#,##0.00")
            pivot.update_field(amount, number_format='0.0000"x"')
            custom = re.search(r'<pivotField [^>]*numFmtId="(\d+)"', _pivot_part(wb))
            self.assertIsNotNone(custom)
            self.assertGreaterEqual(int(custom.group(1)), 164, "a custom code gets an id from 164")
            self.assertEqual(pivot.list_fields()[amount].number_format, '0.0000"x"')

            # "" and None mean no format.
            for none in ("", None):
                with self.subTest(number_format=none):
                    pivot.update_field(amount, number_format="#,##0")
                    pivot.update_field(amount, number_format=none)
                    self.assertIsNone(pivot.list_fields()[amount].number_format)
                    self.assertNotRegex(_pivot_part(wb), r"<pivotField [^>]*numFmtId=")

            # Data fields take the same codes.
            pivot.set_data_field(
                0,
                PivotDataFieldInput(
                    name="Sum of Amount", field_index=amount, aggregation=PivotAggregation.SUM, number_format=code27
                ),
            )
            self.assertEqual(pivot.list_data_fields()[0].number_format, code27)
            self.assertRegex(_pivot_part(wb), r'<dataField [^>]*numFmtId="27"')
            pivot.set_data_field(
                0,
                PivotDataFieldInput(
                    name="Sum of Amount", field_index=amount, aggregation=PivotAggregation.SUM, number_format=None
                ),
            )
            self.assertIsNone(pivot.list_data_fields()[0].number_format)


# ---- Threaded comments --------------------------------------------------------------------------

PERSON = "{0A1B2C3D-0000-4000-8000-000000000001}"
THREAD = "{6F1E0D2C-1111-4111-8111-000000000001}"
REPLY = "{6F1E0D2C-1111-4111-8111-000000000002}"


class ThreadedDoneTests(HandleTestCase):
    def test_done_resolves_a_thread_is_invalid_on_a_reply_and_none_is_a_type_error(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            wb.add_person(Person(id=PERSON, display_name="Alice", user_id="alice", provider_id="None"))
            ws.add_threaded_comment(
                ThreadedCommentInput(
                    id=THREAD, row=1, col=1, person_id=PERSON, created="2026-01-02T03:04:05.67", text="Check"
                )
            )
            ws.add_threaded_comment(
                ThreadedCommentInput(
                    id=REPLY,
                    row=1,
                    col=1,
                    person_id=PERSON,
                    created="2026-01-02T03:05:05.67",
                    text="Done?",
                    parent_id=THREAD,
                )
            )
            self.assertFalse(ws.get_threaded_comment(THREAD).done)

            ws.update_threaded_comment(THREAD, done=True)
            self.assertTrue(ws.get_threaded_comment(THREAD).done)
            self.assertEqual(ws.get_threaded_comment(THREAD).text, "Check")
            ws.update_threaded_comment(THREAD, text="Edited")
            self.assertEqual(ws.get_threaded_comment(THREAD).text, "Edited")
            self.assertTrue(ws.get_threaded_comment(THREAD).done, "a patch without done keeps it")
            ws.update_threaded_comment(THREAD, done=False)
            self.assertFalse(ws.get_threaded_comment(THREAD).done)

            self.assertCode(
                FormulonError,
                lambda: ws.update_threaded_comment(REPLY, done=True),
                StatusCode.THREADED_COMMENT_INVALID,
            )
            self.assertEqual(StatusCode.THREADED_COMMENT_INVALID, 9006)
            self.assertFalse(ws.get_threaded_comment(REPLY).done)
            with self.assertRaises(TypeError):
                ws.update_threaded_comment(THREAD, done=None)  # type: ignore[arg-type]
            with self.assertRaises(TypeError):
                ws.update_threaded_comment(THREAD, done="yes")  # type: ignore[arg-type]
            self.assertFalse(ws.get_threaded_comment(THREAD).done)
            self.assertFalse(hasattr(Sheet, "set_thread_resolved"))


# ---- Sparklines -------------------------------------------------------------------------------------


class SparklineUpdateTests(HandleTestCase):
    def test_update_replaces_keeps_or_refuses_to_empty_the_group(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.add_sparkline_group(
                SparklineGroup(sparklines=[Sparkline("Sheet1!A1:D1", "F1"), Sparkline("Sheet1!A2:D2", "F2")])
            )
            original = list(ws.list_sparkline_groups()[0].sparklines)
            self.assertEqual(len(original), 2)

            # A partial update keeps the sparklines, whatever else changes.
            ws.update_sparkline_group(0, markers=True)
            self.assertTrue(ws.list_sparkline_groups()[0].markers)
            self.assertEqual(ws.list_sparkline_groups()[0].sparklines, original)

            # A new list replaces them.
            replacement = [Sparkline("Sheet1!B1:B5", "H1")]
            ws.update_sparkline_group(0, sparklines=replacement)
            self.assertEqual(ws.list_sparkline_groups()[0].sparklines, replacement)
            self.assertTrue(ws.list_sparkline_groups()[0].markers)

            # An empty list is refused and the group is untouched.
            self.assertCode(
                InvalidArgumentError,
                lambda: ws.update_sparkline_group(0, sparklines=[]),
                StatusCode.INVALID_ARGUMENT,
            )
            self.assertEqual(ws.list_sparkline_groups()[0].sparklines, replacement)
            self.assertFalse(hasattr(Sheet, "set_sparkline_group_sparklines"))


if __name__ == "__main__":
    unittest.main()
