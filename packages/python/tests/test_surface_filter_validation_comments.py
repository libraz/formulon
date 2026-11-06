"""AutoFilter, data-validation check and threaded-comment surface tests."""

from __future__ import annotations

import unittest

from formulon import (
    AutoFilter,
    AutoFilterColumn,
    AutoFilterDateGroup,
    AutoFilterEval,
    AutoFilterSortCondition,
    AutoFilterSortState,
    CellRect,
    DataValidationInput,
    DateTimeGrouping,
    DynamicFilterType,
    FilterKind,
    FilterOperator,
    FormulonError,
    Mention,
    Person,
    SheetRect,
    SortBy,
    SortMethod,
    ThreadedCommentInput,
    ValidationErrorStyle,
    ValidationOutcome,
    Value,
    ValueKind,
    Workbook,
)

ALICE = "{11111111-1111-4111-8111-111111111111}"
BOB = "{22222222-2222-4222-8222-222222222222}"
THREAD = "{33333333-3333-4333-8333-333333333333}"
REPLY = "{44444444-4444-4444-8444-444444444444}"
STAMP = "2026-01-01T10:00:00.00"


def _text(s: str) -> Value:
    return Value(kind=ValueKind.TEXT, text=s)


def _seed_numbers(wb: Workbook, col: int = 0, first_row: int = 0) -> None:
    wb.sheet(0).set_text(first_row, col, "n")
    for i, v in enumerate((1, 2, 3, 4, 5)):
        wb.sheet(0).set_number(first_row + 1 + i, col, float(v))


class EnumOrdinalTests(unittest.TestCase):
    def test_ordinals_match_the_c_header(self) -> None:
        self.assertEqual(
            (FilterKind.ICON, FilterOperator.GREATER_THAN, DynamicFilterType.YEAR_TO_DATE, DynamicFilterType.M12),
            (6, 5, 18, 34),
        )
        self.assertEqual(
            (SortBy.ICON, SortMethod.STROKE, DateTimeGrouping.SECOND, ValidationErrorStyle.INFORMATION),
            (3, 2, 5, 2),
        )


class AutoFilterTests(unittest.TestCase):
    def test_absent_filter_reads_none(self) -> None:
        with Workbook.create_default() as wb:
            self.assertIsNone(wb.sheet(0).get_auto_filter())

    def test_set_get_roundtrip_every_criterion(self) -> None:
        with Workbook.create_default() as wb:
            _seed_numbers(wb)
            sort = AutoFilterSortState(
                ref=CellRect(1, 0, 5, 1),
                column_sort=False,
                case_sensitive=True,
                sort_method=SortMethod.PIN_YIN,
                conditions=[
                    AutoFilterSortCondition(ref=CellRect(1, 0, 5, 0), descending=True, custom_list="a,b"),
                ],
            )
            af = AutoFilter(
                range=CellRect(0, 0, 5, 1),
                columns=[
                    AutoFilterColumn(
                        col_id=0,
                        hidden_button=True,
                        kind=1,
                        filter_blank=True,
                        values=["1", "3"],
                        date_groups=[AutoFilterDateGroup(year=2024, month=5, grouping=DateTimeGrouping.MONTH)],
                    ),
                    AutoFilterColumn(
                        col_id=1,
                        kind=FilterKind.CUSTOM,
                        custom_and=True,
                        custom_count=2,
                        op1=FilterOperator.GREATER_THAN,
                        val1="1",
                        op2=FilterOperator.LESS_THAN,
                        val2="9",
                    ),
                ],
                sort=sort,
            )
            wb.sheet(0).set_auto_filter(af)
            got = wb.sheet(0).get_auto_filter()
            assert got is not None
            self.assertEqual(got.range, af.range)
            self.assertEqual(len(got.columns), 2)
            c0, c1 = got.columns
            self.assertTrue(c0.hidden_button)
            self.assertTrue(c0.show_button)
            self.assertTrue(c0.filter_blank)
            self.assertEqual(c0.values, ["1", "3"])
            self.assertEqual(c0.date_groups, [AutoFilterDateGroup(year=2024, month=5, grouping=DateTimeGrouping.MONTH)])
            self.assertEqual(
                (c1.kind, c1.custom_count, c1.op1, c1.val1, c1.op2, c1.val2),
                (FilterKind.CUSTOM, 2, FilterOperator.GREATER_THAN, "1", FilterOperator.LESS_THAN, "9"),
            )
            self.assertTrue(c1.custom_and)
            assert got.sort is not None
            self.assertEqual(got.sort.ref, sort.ref)
            self.assertTrue(got.sort.case_sensitive)
            self.assertEqual(got.sort.sort_method, SortMethod.PIN_YIN)
            self.assertEqual(got.sort.conditions[0].custom_list, "a,b")
            self.assertTrue(got.sort.conditions[0].descending)
            # A get followed by a set of the unchanged record is stable.
            wb.sheet(0).set_auto_filter(got)
            self.assertEqual(wb.sheet(0).get_auto_filter(), got)

    def test_top10_and_dynamic_fields_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            _seed_numbers(wb)
            col = AutoFilterColumn(
                col_id=0,
                kind=FilterKind.TOP10,
                top=False,
                percent=True,
                top_val=25.0,
                filter_val=2.0,
            )
            dyn = AutoFilterColumn(
                col_id=1,
                kind=FilterKind.DYNAMIC,
                dynamic_type=DynamicFilterType.ABOVE_AVERAGE,
                dyn_val=3.5,
                val_iso="x",
            )
            wb.sheet(0).set_auto_filter(AutoFilter(range=CellRect(0, 0, 5, 1), columns=[col, dyn]))
            got = wb.sheet(0).get_auto_filter()
            assert got is not None
            self.assertFalse(got.columns[0].top)
            self.assertTrue(got.columns[0].percent)
            self.assertEqual((got.columns[0].top_val, got.columns[0].filter_val), (25.0, 2.0))
            self.assertEqual(got.columns[1].dynamic_type, DynamicFilterType.ABOVE_AVERAGE)
            self.assertEqual(got.columns[1].dyn_val, 3.5)

    def test_evaluate_apply_clear_remove(self) -> None:
        with Workbook.create_default() as wb:
            _seed_numbers(wb)
            wb.sheet(0).set_auto_filter(
                AutoFilter(
                    range=CellRect(0, 0, 5, 0),
                    columns=[
                        AutoFilterColumn(
                            col_id=0, kind=FilterKind.CUSTOM, custom_count=1, op1=FilterOperator.GREATER_THAN, val1="2"
                        )
                    ],
                ),
            )
            evaluated = wb.sheet(0).evaluate_auto_filter()
            first_row, match = evaluated.first_row, evaluated.match
            self.assertEqual(first_row, 1)
            self.assertEqual(match, [False, False, True, True, True])
            wb.sheet(0).apply_auto_filter()
            hidden = [r.row for r in wb.sheet(0).list_row_overrides() if r.hidden]
            self.assertEqual(hidden, [1, 2])
            wb.sheet(0).clear_auto_filter()
            got = wb.sheet(0).get_auto_filter()
            assert got is not None
            self.assertEqual(got.columns, [])
            self.assertEqual(wb.sheet(0).evaluate_auto_filter(), AutoFilterEval(1, [True] * 5))
            wb.sheet(0).remove_auto_filter()
            self.assertIsNone(wb.sheet(0).get_auto_filter())
            wb.sheet(0).remove_auto_filter()
            with self.assertRaises(FormulonError):
                wb.sheet(0).apply_auto_filter()
            with self.assertRaises(FormulonError):
                wb.sheet(0).evaluate_auto_filter()

    def test_invalid_filter_is_rejected(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(FormulonError):
                wb.sheet(0).set_auto_filter(
                    AutoFilter(range=CellRect(0, 0, 3, 0), columns=[AutoFilterColumn(col_id=5)]),
                )
            with self.assertRaises(FormulonError):
                wb.sheet(99).get_auto_filter()

    def test_table_auto_filter(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_text(0, 0, "n")
            wb.sheet(0).set_text(0, 1, "m")
            for i in range(1, 4):
                wb.sheet(0).set_number(i, 0, float(i))
                wb.sheet(0).set_number(i, 1, 0.0)
            idx = wb.add_table(
                range=SheetRect(0, CellRect(0, 0, 3, 1)), name="Tbl", display_name="Tbl", columns=["n", "m"]
            )
            wb.remove_table_auto_filter(idx)
            self.assertIsNone(wb.get_table_auto_filter(idx))
            wb.set_table_auto_filter(
                idx,
                AutoFilter(
                    range=CellRect(0, 0, 3, 1),
                    columns=[
                        AutoFilterColumn(
                            col_id=0,
                            kind=FilterKind.CUSTOM,
                            custom_count=1,
                            op1=FilterOperator.GREATER_THAN_OR_EQUAL,
                            val1="2",
                        )
                    ],
                ),
            )
            got = wb.get_table_auto_filter(idx)
            assert got is not None
            self.assertEqual(got.columns[0].val1, "2")
            self.assertEqual(wb.evaluate_table_auto_filter(idx), AutoFilterEval(1, [False, True, True]))
            wb.apply_table_auto_filter(idx)
            wb.clear_table_auto_filter(idx)
            self.assertEqual(wb.evaluate_table_auto_filter(idx), AutoFilterEval(1, [True, True, True]))
            wb.remove_table_auto_filter(idx)
            self.assertIsNone(wb.get_table_auto_filter(idx))
            with self.assertRaises(FormulonError):
                wb.get_table_auto_filter(idx + 5)


class ValidationCheckTests(unittest.TestCase):
    def _wb_with_list_rule(self) -> Workbook:
        wb = Workbook.create_default()
        wb.sheet(0).add_validation(
            DataValidationInput(
                type=3, ranges=[CellRect(0, 3, 9, 3)], formula1='"Yes,No"', error_style=ValidationErrorStyle.WARNING
            ),
        )
        return wb

    def test_validate_value(self) -> None:
        with self._wb_with_list_rule() as wb:
            ok = wb.sheet(0).validate_value(0, 3, _text("Yes"))
            self.assertEqual(ok, ValidationOutcome(valid=True, rule_index=0, error_style=ValidationErrorStyle.WARNING))
            bad = wb.sheet(0).validate_value(0, 3, _text("Maybe"))
            self.assertEqual(bad.rule_index, 0)
            self.assertFalse(bad.valid)
            self.assertEqual(bad.error_style, ValidationErrorStyle.WARNING)
            none = wb.sheet(0).validate_value(0, 0, Value(kind=ValueKind.NUMBER, number=1.0))
            self.assertEqual(none, ValidationOutcome(valid=True, rule_index=None, error_style=None))
            with self.assertRaises(FormulonError):
                wb.sheet(99).validate_value(0, 3, _text("Yes"))

    def test_list_invalid_cells_pages(self) -> None:
        with self._wb_with_list_rule() as wb:
            for r, v in enumerate(["Yes", "bad1", "No", "bad2", "bad3"]):
                wb.sheet(0).set_text(r, 3, v)
            first_page = wb.sheet(0).get_invalid_cells()
            cells, cursor = first_page.cells, first_page.next_cursor
            self.assertEqual([c.row for c in cells], [1, 3, 4])
            self.assertIsNone(cursor)
            seen = []
            cursor = None
            pages = 0
            while True:
                got = wb.sheet(0).get_invalid_cells(cursor, 2)
                page, cursor = got.cells, got.next_cursor
                pages += 1
                seen += [(c.row, c.value.text) for c in page]
                if cursor is None:
                    break
            self.assertEqual(seen, [(1, "bad1"), (3, "bad2"), (4, "bad3")])
            self.assertGreaterEqual(pages, 2)
            with self.assertRaises(FormulonError):
                wb.sheet(99).get_invalid_cells()


class ThreadedCommentTests(unittest.TestCase):
    def _wb(self) -> Workbook:
        wb = Workbook.create_default()
        wb.add_person(Person(id=ALICE, display_name="Alice", user_id="alice@example.com", provider_id="AD"))
        wb.add_person(Person(id=BOB, display_name="Bob", user_id="bob@example.com", provider_id="AD"))
        return wb

    def test_persons(self) -> None:
        with self._wb() as wb:
            persons = wb.list_people()
            self.assertEqual([p.display_name for p in persons], ["Alice", "Bob"])
            self.assertEqual(persons[0], Person(ALICE, "Alice", "alice@example.com", "AD"))
            wb.remove_person(BOB)
            self.assertEqual([p.id for p in wb.list_people()], [ALICE])
            with self.assertRaises(FormulonError):
                wb.remove_person(BOB)
            with self.assertRaises(FormulonError):
                wb.add_person(Person(id=ALICE, display_name="Dup"))

    def test_thread_with_reply_and_mention(self) -> None:
        with self._wb() as wb:
            self.assertEqual(wb.sheet(0).list_threaded_comments(), [])
            opener = ThreadedCommentInput(id=THREAD, row=2, col=1, person_id=ALICE, created=STAMP, text="Check this")
            wb.sheet(0).add_threaded_comment(opener)
            reply = ThreadedCommentInput(
                id=REPLY,
                person_id=BOB,
                created=STAMP,
                text="@Alice done",
                parent_id=THREAD,
                mentions=[
                    Mention(person_id=ALICE, mention_id="{55555555-5555-4555-8555-555555555555}", start=0, length=6)
                ],
            )
            wb.sheet(0).add_threaded_comment(reply)
            got = wb.sheet(0).list_threaded_comments()
            self.assertEqual([c.id for c in got], [THREAD, REPLY])
            self.assertEqual((got[0].row, got[0].col, got[0].text), (2, 1, "Check this"))
            self.assertEqual((got[1].row, got[1].col, got[1].parent_id), (2, 1, THREAD))
            self.assertEqual(got[1].mentions, reply.mentions)

            with self.assertRaises(FormulonError):
                wb.remove_person(ALICE)

            wb.sheet(0).update_threaded_comment(REPLY, text="plain", mentions=[])
            self.assertEqual(wb.sheet(0).list_threaded_comments()[1].text, "plain")
            self.assertEqual(wb.sheet(0).list_threaded_comments()[1].mentions, [])

            wb.sheet(0).set_thread_resolved(THREAD, True)
            self.assertTrue(wb.sheet(0).list_threaded_comments()[0].done)
            wb.sheet(0).set_thread_resolved(THREAD, False)
            self.assertFalse(wb.sheet(0).list_threaded_comments()[0].done)

            wb.sheet(0).remove_threaded_comment(REPLY)
            self.assertEqual([c.id for c in wb.sheet(0).list_threaded_comments()], [THREAD])
            wb.sheet(0).remove_threaded_comment(THREAD)
            self.assertEqual(wb.sheet(0).list_threaded_comments(), [])
            wb.remove_person(ALICE)

    def test_invalid_comment_is_rejected(self) -> None:
        with self._wb() as wb:
            with self.assertRaises(FormulonError):
                wb.sheet(0).add_threaded_comment(
                    ThreadedCommentInput(id="not-a-guid", person_id=ALICE, created=STAMP, text="x")
                )
            with self.assertRaises(FormulonError):
                wb.sheet(0).update_threaded_comment(THREAD, text="x")
            with self.assertRaises(FormulonError):
                wb.sheet(0).set_thread_resolved(THREAD, True)
            with self.assertRaises(FormulonError):
                wb.sheet(0).remove_threaded_comment(THREAD)


if __name__ == "__main__":
    unittest.main()
