"""Read-side shapes: default ranges, record fields named alike, and records that pass back into their setters."""

from __future__ import annotations

import dataclasses
import datetime
import unittest
from typing import Any

import formulon
from formulon import (
    CellAddress,
    CellRect,
    CellXfInput,
    ComparisonOperator,
    FilterKind,
    FormulonError,
    SheetRect,
    Workbook,
    XfCell,
)

_PERSON = "{11111111-1111-4111-8111-111111111111}"
_STAMP = datetime.datetime(2026, 1, 2, 3, 4, 5)
_T1 = "{22222222-2222-4222-8222-222222222221}"
_R1 = "{22222222-2222-4222-8222-222222222222}"
_T3 = "{22222222-2222-4222-8222-222222222223}"


def _grid(ws: Any) -> None:
    for c, name in enumerate(("Name", "North", "South")):
        ws.set_text(0, c, name)
    for r in range(1, 4):
        ws.set_text(r, 0, f"row{r}")
        ws.set_number(r, 1, float(r))
        ws.set_number(r, 2, float(r * 2))


class BlockReadDefaultRangeTests(unittest.TestCase):
    def test_empty_sheet_reads_are_empty(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            self.assertEqual(ws.get_values(), [])
            self.assertEqual(ws.get_values(None), [])
            self.assertEqual(ws.get_typed_values(), [])
            self.assertEqual(ws.get_typed_values(None), [])
            self.assertEqual(list(ws.iter_rows()), [])
            self.assertEqual(ws.list_cells(), [])
            page = ws.get_cells_in_range()
            self.assertEqual(page.cells, [])
            self.assertIsNone(page.next_cursor)

    def test_default_range_is_the_used_range(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            _grid(ws)
            used = ws.get_used_range()
            self.assertEqual(ws.get_values(), ws.get_values(used))
            self.assertEqual(ws.get_values(None), ws.get_values(used))
            self.assertEqual(ws.get_typed_values(), ws.get_typed_values(used))
            self.assertEqual(list(ws.iter_rows()), list(ws.iter_rows(used)))
            self.assertEqual(ws.list_cells(), ws.list_cells(used))
            self.assertEqual(ws.get_cells_in_range().cells, ws.get_cells_in_range(used).cells)

    def test_a_far_cell_does_not_widen_the_list_beyond_the_used_range(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.set_number(2, 2, 1.0)
            self.assertEqual([(c.row, c.col) for c in ws.list_cells()], [(2, 2)])


class ThreadedCommentCellTests(unittest.TestCase):
    def _threads(self, wb: Workbook) -> Any:
        ws = wb.sheet(0)
        wb.add_person(formulon.Person(id=_PERSON, display_name="Ann"))
        ws.add_threaded_comment(
            formulon.ThreadedCommentInput(id=_T1, person_id=_PERSON, text="one", cell="B2", created=_STAMP)
        )
        ws.add_threaded_comment(
            formulon.ThreadedCommentInput(id=_R1, person_id=_PERSON, text="re", parent_id=_T1, created=_STAMP)
        )
        ws.add_threaded_comment(
            formulon.ThreadedCommentInput(
                id=_T3, person_id=_PERSON, text="three", cell=CellAddress(2, 2), created=_STAMP
            )
        )
        return ws

    def test_cell_is_a_cell_address_on_read(self) -> None:
        with Workbook.create_default() as wb:
            ws = self._threads(wb)
            comment = ws.get_threaded_comment(_T1)
            self.assertEqual(comment.cell, CellAddress(1, 1))
            self.assertFalse(hasattr(comment, "row"))
            self.assertFalse(hasattr(comment, "col"))
            self.assertEqual(ws.get_threaded_comment(_T3).cell, CellAddress(2, 2))

    def test_list_filters_by_cell(self) -> None:
        with Workbook.create_default() as wb:
            ws = self._threads(wb)
            self.assertEqual(len(ws.list_threaded_comments()), 3)
            by_a1 = ws.list_threaded_comments("B2")
            self.assertEqual({c.id for c in by_a1}, {_T1, _R1})
            self.assertEqual(by_a1, ws.list_threaded_comments(1, 1))
            self.assertEqual([c.id for c in ws.list_threaded_comments("C3")], [_T3])
            self.assertEqual(ws.list_threaded_comments("F9"), [])
            with self.assertRaises(TypeError):
                ws.list_threaded_comments(cell="B2")  # type: ignore[call-overload]

    def test_input_rejects_the_old_row_and_column_fields(self) -> None:
        with self.assertRaises(TypeError):
            formulon.ThreadedCommentInput(id="x", person_id="p", text="t", row=1, col=1)  # type: ignore[call-arg]


class ReferenceFormattingTests(unittest.TestCase):
    def test_format_reference_takes_a_parsed_reference(self) -> None:
        with Workbook.create_default() as wb:
            parsed = wb.parse_reference("'My Sheet'!B2:D4")
            text = wb.format_reference(parsed)
            self.assertEqual(wb.parse_reference(text), parsed)
            self.assertEqual(parsed.range, CellRect(1, 1, 3, 3))
            self.assertFalse(hasattr(parsed, "rect"))
            with self.assertRaises(TypeError):
                wb.format_reference(parsed, sheet_name="Other")

    def test_format_reference_takes_a_rect_or_an_unqualified_string(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(
                wb.format_reference("B2:D4", sheet_name="S"), wb.format_reference(CellRect(1, 1, 3, 3), sheet_name="S")
            )
            with self.assertRaises(FormulonError):
                wb.format_reference("Sheet1!B2:D4")

    def test_range_fields_are_named_range(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.set_hyperlink(
                "A1",
                formulon.Hyperlink(
                    range=CellRect(0, 0, 0, 0), target="https://example.com", location=None, display=None, tooltip=None
                ),
            )
            self.assertEqual(ws.get_hyperlink("A1").range, CellRect(0, 0, 0, 0))
        for cls in (
            formulon.Hyperlink,
            formulon.ArrayFormula,
            formulon.Spill,
            formulon.ParsedReference,
            formulon.DataTable,
        ):
            names = {f.name for f in dataclasses.fields(cls)}
            self.assertIn("range", names, cls.__name__)
            self.assertNotIn("rect", names, cls.__name__)
        for cls in (formulon.AutoFilterSortCondition, formulon.AutoFilterSortState):
            names = {f.name for f in dataclasses.fields(cls)}
            self.assertIn("range", names, cls.__name__)
            self.assertNotIn("ref", names, cls.__name__)
        for cls in (formulon.PivotLayout, formulon.PaginationPage):
            names = {f.name for f in dataclasses.fields(cls)}
            self.assertIn("range", names, cls.__name__)
            self.assertNotIn("rect", names, cls.__name__)
        self.assertIn("rect", {f.name for f in dataclasses.fields(SheetRect)})


class XfRemapTests(unittest.TestCase):
    def test_excluding_cells_accepts_addresses_pairs_and_a1(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            as_pairs = ws.ooxml.get_range_xf_bases_excluding("A1:C3", [(0, 0), (1, 1)])
            self.assertEqual(ws.ooxml.get_range_xf_bases_excluding("A1:C3", [CellAddress(0, 0), "B2"]), as_pairs)
            self.assertEqual(
                ws.ooxml.get_range_xf_bases_excluding("A1:C3", [CellAddress(0, 0), CellAddress(1, 1)]), as_pairs
            )

    def test_remap_takes_any_mapping(self) -> None:
        import types

        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.set_number(0, 0, 1.0)
            base = ws.ooxml.get_range_xf_bases("A1:B2")[0]
            ws.ooxml.remap_range_xfs("A1:B2", types.MappingProxyType({base: base}))
            self.assertEqual(ws.ooxml.get_range_xf_bases("A1:B2")[0], base)
            ws.ooxml.remap_range_xfs("A1:B2", {base: base})


class ListCellXfIndicesTests(unittest.TestCase):
    @staticmethod
    def _styled(wb: Workbook) -> Any:
        """Styled blanks, styled content, unstyled content and an explicit XF 0; returns the two XFs used."""
        ws = wb.sheet(0)
        xfs = [
            wb.ooxml.add_cell_xf(
                CellXfInput(font_index=0, fill_index=0, border_index=0, num_fmt_id=num_fmt_id, wrap_text=None)
            )
            for num_fmt_id in (2, 3)
        ]
        ws.ooxml.set_cell_xf_index("U1001", xfs[0])  # styled blank, far from the content
        ws.ooxml.set_cell_xf_index("D6", xfs[1])  # styled blank
        ws.set_number("H3", 4.0)
        ws.ooxml.set_cell_xf_index("H3", xfs[0])  # content with a style
        ws.set_number("B5", 9.0)  # content, unstyled
        ws.ooxml.set_cell_xf_index("G7", 0)  # explicit default
        return xfs

    def test_cells_are_xf_cells_and_the_get_names_are_gone(self) -> None:
        with Workbook.create_default() as wb:
            self._styled(wb)
            ws = wb.sheet(0)
            cells = ws.ooxml.list_cell_xf_indices()
            self.assertTrue(all(isinstance(c, XfCell) for c in cells))
            self.assertEqual([f.name for f in dataclasses.fields(XfCell)], ["row", "col", "xf_index"])
            self.assertFalse(hasattr(ws.ooxml, "get_cell_xf_indices"))
            self.assertFalse(hasattr(ws, "get_show_phonetic_cells"))
            self.assertFalse(hasattr(formulon, "CellXfIndex"))

    def test_lists_styled_blank_and_content_cells_row_major(self) -> None:
        with Workbook.create_default() as wb:
            a, b = self._styled(wb)
            self.assertEqual(
                wb.sheet(0).ooxml.list_cell_xf_indices(),
                [XfCell(2, 7, a), XfCell(5, 3, b), XfCell(1000, 20, a)],
            )

    def test_the_default_range_covers_format_only_cells_outside_the_used_range(self) -> None:
        with Workbook.create_default() as wb:
            self._styled(wb)
            ws = wb.sheet(0)
            self.assertLess(ws.get_used_range().last_row, 1000)
            self.assertEqual(len(ws.ooxml.list_cell_xf_indices(None)), 3)
            self.assertEqual(ws.ooxml.list_cell_xf_indices(), ws.ooxml.list_cell_xf_indices("A1:XFD1048576"))
            whole = CellRect(0, 0, 1048575, 16383)
            self.assertEqual(ws.ooxml.list_cell_xf_indices(), ws.ooxml.list_cell_xf_indices(whole))

    def test_clips_to_a_rect_or_an_a1_range(self) -> None:
        with Workbook.create_default() as wb:
            a, b = self._styled(wb)
            ooxml = wb.sheet(0).ooxml
            clipped = [XfCell(2, 7, a), XfCell(5, 3, b)]
            self.assertEqual(ooxml.list_cell_xf_indices(CellRect(0, 0, 5, 1000)), clipped)
            self.assertEqual(ooxml.list_cell_xf_indices("A1:Z6"), clipped)
            self.assertEqual(ooxml.list_cell_xf_indices("D6"), [XfCell(5, 3, b)])
            self.assertEqual(ooxml.list_cell_xf_indices("A8:C10"), [])

    def test_the_record_is_frozen_and_survives_save_and_load(self) -> None:
        with Workbook.create_default() as wb:
            self._styled(wb)
            expected = wb.sheet(0).ooxml.list_cell_xf_indices()
            with self.assertRaises(dataclasses.FrozenInstanceError):
                expected[0].row = 1  # type: ignore[misc]
            saved = wb.save()
        with Workbook.load(saved) as reloaded:
            self.assertEqual(reloaded.sheet(0).ooxml.list_cell_xf_indices(), expected)

    def test_an_out_of_grid_rect_is_refused(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(FormulonError):
                wb.sheet(0).ooxml.list_cell_xf_indices(CellRect(0, 0, 1048576, 0))


class ListShowPhoneticCellsTests(unittest.TestCase):
    @staticmethod
    def _flagged(wb: Workbook) -> None:
        """A flagged constant, a flag-only blank, a flagged text, a far flagged blank, and unflagged cells."""
        ws = wb.sheet(0)
        ws.set_blank("U1001")
        ws.set_cell_show_phonetic("U1001", True)  # flag-only blank, far from the content
        ws.set_number("B2", 5.0)
        ws.set_cell_show_phonetic("B2", True)  # flagged constant
        ws.set_cell_show_phonetic("A3", True)  # flag-only blank
        ws.set_text("H5", "x")
        ws.set_cell_show_phonetic("H5", True)  # flagged text
        ws.set_number("C2", 1.0)  # content, no flag
        ws.set_blank("D6")  # stored blank, no flag

    ALL = [CellAddress(1, 1), CellAddress(2, 0), CellAddress(4, 7), CellAddress(1000, 20)]

    def test_lists_flagged_cells_row_major_and_leaves_out_unflagged_ones(self) -> None:
        with Workbook.create_default() as wb:
            self._flagged(wb)
            ws = wb.sheet(0)
            self.assertEqual(ws.list_show_phonetic_cells(), self.ALL)
            for cell in self.ALL:
                self.assertTrue(ws.get_cell_show_phonetic(cell.row, cell.col))

    def test_the_default_range_covers_flagged_cells_outside_the_used_range(self) -> None:
        with Workbook.create_default() as wb:
            self._flagged(wb)
            ws = wb.sheet(0)
            self.assertLess(ws.get_used_range().last_row, 1000)
            self.assertEqual(len(ws.list_show_phonetic_cells(None)), 4)
            self.assertEqual(ws.list_show_phonetic_cells(), ws.list_show_phonetic_cells("A1:XFD1048576"))
            self.assertEqual(ws.list_show_phonetic_cells(), ws.list_show_phonetic_cells(CellRect(0, 0, 1048575, 16383)))

    def test_clips_to_a_rect_or_an_a1_range(self) -> None:
        with Workbook.create_default() as wb:
            self._flagged(wb)
            ws = wb.sheet(0)
            clipped = [CellAddress(1, 1), CellAddress(2, 0)]
            self.assertEqual(ws.list_show_phonetic_cells(CellRect(0, 0, 2, 1)), clipped)
            self.assertEqual(ws.list_show_phonetic_cells("A1:B3"), clipped)
            self.assertEqual(ws.list_show_phonetic_cells("H5"), [CellAddress(4, 7)])
            self.assertEqual(ws.list_show_phonetic_cells("A8:C10"), [])

    def test_the_flag_only_blank_survives_save_and_load(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(wb.sheet(0).list_show_phonetic_cells(), [])
            self._flagged(wb)
            saved = wb.save()
        with Workbook.load(saved) as reloaded:
            self.assertEqual(reloaded.sheet(0).list_show_phonetic_cells(), self.ALL)

    def test_an_out_of_grid_rect_is_refused(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(FormulonError):
                wb.sheet(0).list_show_phonetic_cells(CellRect(0, 0, 1048576, 0))


class AutoFilterConditionTests(unittest.TestCase):
    def _filter(self, conditions: Any) -> Any:
        return formulon.AutoFilterInput(
            "A1:B5",
            columns=[
                formulon.AutoFilterColumn(column_offset=0, kind=FilterKind.CUSTOM, conditions=conditions),
            ],
        )

    def test_custom_conditions_are_a_list(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            _grid(ws)
            condition = formulon.AutoFilterCondition(operator=ComparisonOperator.GREATER_THAN, value="3")
            ws.set_auto_filter(self._filter([condition]))
            column = ws.get_auto_filter().columns[0]
            self.assertEqual(column.conditions, [condition])
            ws.set_auto_filter(ws.get_auto_filter())
            self.assertEqual(ws.get_auto_filter().columns[0].conditions, [condition])

    def test_two_conditions_round_trip(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            _grid(ws)
            pair = [
                formulon.AutoFilterCondition(ComparisonOperator.GREATER_THAN, "1"),
                formulon.AutoFilterCondition(ComparisonOperator.LESS_THAN, "9"),
            ]
            ws.set_auto_filter(self._filter(pair))
            self.assertEqual(ws.get_auto_filter().columns[0].conditions, pair)

    def test_condition_count_outside_one_or_two_is_refused(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            _grid(ws)
            one = formulon.AutoFilterCondition(ComparisonOperator.EQUAL, "1")
            for bad in ([], [one, one, one]):
                with self.assertRaises(FormulonError):
                    ws.set_auto_filter(self._filter(bad))
            self.assertIsNone(ws.get_auto_filter())

    def test_between_is_refused(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            _grid(ws)
            with self.assertRaises(FormulonError):
                ws.set_auto_filter(self._filter([formulon.AutoFilterCondition(ComparisonOperator.BETWEEN, "1")]))

    def test_the_five_flat_fields_are_gone(self) -> None:
        names = {f.name for f in dataclasses.fields(formulon.AutoFilterColumn)}
        self.assertIn("conditions", names)
        for gone in ("custom_count", "operator1", "value1", "operator2", "value2"):
            self.assertNotIn(gone, names)
        with self.assertRaises(TypeError):
            formulon.AutoFilterColumn(column_offset=0, custom_count=1)  # type: ignore[call-arg]

    def test_condition_coerces_its_operator(self) -> None:
        condition = formulon.AutoFilterCondition(operator=int(ComparisonOperator.EQUAL), value="1")  # type: ignore[arg-type]
        self.assertIs(condition.operator, ComparisonOperator.EQUAL)
        with self.assertRaises(dataclasses.FrozenInstanceError):
            condition.value = "2"  # type: ignore[misc]


class RecordShapeTests(unittest.TestCase):
    def test_cell_xf_input_carries_no_presence_flags(self) -> None:
        names = {f.name for f in dataclasses.fields(formulon.CellXfInput)}
        for gone in (
            "has_alignment",
            "has_horizontal",
            "has_vertical",
            "has_wrap_text",
            "has_justify_last_line",
            "has_protection",
        ):
            self.assertNotIn(gone, names)

    def test_cell_xf_round_trips_through_the_input(self) -> None:
        with Workbook.create_default() as wb:
            ooxml = wb.ooxml
            minimal = formulon.CellXfInput(font_index=0, fill_index=0, border_index=0, num_fmt_id=0)
            index = ooxml.add_cell_xf(minimal)
            read = ooxml.get_cell_xf(index)
            self.assertIsNone(read.horizontal)
            self.assertEqual(ooxml.add_cell_xf(read), index)

    def test_pivot_field_reads_show_detail(self) -> None:
        names = {f.name for f in dataclasses.fields(formulon.PivotField)}
        self.assertIn("show_detail", names)

    def test_merged_function_metadata_is_always_the_merged_shape(self) -> None:
        base = formulon.get_function_metadata("SUM")
        merged = formulon.merge_function_metadata(base, None, "fr-FR")
        self.assertIsInstance(merged, formulon.MergedFunctionMetadata)
        self.assertEqual(merged.localized_name, base.name)
        self.assertEqual(merged.name, base.name)
        merged = formulon.merge_function_metadata(base, {}, "fr-FR")
        self.assertIsInstance(merged, formulon.MergedFunctionMetadata)

    def test_set_defined_name_accepts_the_read_record(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_defined_name(formulon.DefinedNameInput(name="Rate", formula="=0.5"))
            wb.set_defined_name(wb.get_defined_name("Rate"))
            self.assertEqual(wb.get_defined_name("Rate").formula, "=0.5")


class CfMatchKindTests(unittest.TestCase):
    def test_a_match_carries_only_the_members_of_its_kind(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            for r in range(3):
                ws.set_number(r, 0, float(r + 1))
            rect = formulon.CellRect(0, 0, 2, 0)
            red = formulon.Color(formulon.ColorKind.RGB, argb=0xFFFF0000)
            ws.add_cf_rule(
                formulon.CfDataBarRuleInput(
                    ranges=[rect],
                    data_bar=formulon.CfDataBarInput(
                        formulon.CfValueObject(formulon.CfValueObjectType.MIN),
                        formulon.CfValueObject(formulon.CfValueObjectType.MAX),
                        red,
                    ),
                )
            )
            ws.add_cf_rule(
                formulon.CfCellIsRuleInput(ranges=[rect], operator=ComparisonOperator.GREATER_THAN, formula1="0")
            )
            matches = ws.evaluate_cf_range(rect)[2].matches
            by_kind = {m.kind: m for m in matches}
            bar = by_kind[formulon.CfMatchKind.DATA_BAR]
            self.assertIsInstance(bar.bar, formulon.CfMatchBar)
            self.assertIsNone(bar.icon)
            self.assertIsNone(bar.color)
            fmt = by_kind[formulon.CfMatchKind.DIFFERENTIAL_FORMAT]
            self.assertIsNone(fmt.bar)
            self.assertIsNone(fmt.icon)
            self.assertIsNone(fmt.color)
