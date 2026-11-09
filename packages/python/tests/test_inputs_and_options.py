"""Input spellings and options: cell-value formatting, keywords, page breaks, paging, single-axis reads, timestamps."""

from __future__ import annotations

import datetime
import inspect
import unittest

import formulon
from formulon import (
    CellAddress,
    InvalidArgumentError,
    PageBreakInput,
    PersonInput,
    RowBreak,
    ThreadedCommentInput,
    Workbook,
)

ALICE = "{11111111-1111-4111-8111-111111111111}"
THREAD = "{33333333-3333-4333-8333-333333333333}"
OTHER = "{44444444-4444-4444-8444-444444444444}"


class FormatValueTests(unittest.TestCase):
    def test_native_number(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(wb.format_value(1234.5, "#,##0.00").text, "1,234.50")
            self.assertEqual(wb.format_value("abc", "General").text, "abc")
            self.assertEqual(wb.format_value(True, "General").text, "TRUE")

    def test_date_is_read_under_the_workbook_date_system(self) -> None:
        with Workbook.create_default() as wb:
            day = datetime.date(2024, 1, 2)
            self.assertEqual(wb.format_value(day, "yyyy-mm-dd").text, "2024-01-02")
            wb.set_date1904(True)
            self.assertEqual(wb.format_value(day, "yyyy-mm-dd").text, "2024-01-02")

    def test_list_is_type_error(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(TypeError):
                wb.format_value([1.0, 2.0], "0")  # type: ignore[arg-type]


class StatusNameTests(unittest.TestCase):
    def test_get_status_name_removed(self) -> None:
        self.assertFalse(hasattr(formulon, "get_status_name"))
        self.assertNotIn("get_status_name", formulon.__all__)
        with self.assertRaises(ImportError):
            from formulon import get_status_name  # noqa: F401


class KeywordNameTests(unittest.TestCase):
    def test_set_pinned_now_time_keyword(self) -> None:
        with Workbook.create_default() as wb:
            moment = datetime.datetime(2026, 4, 23, 15, 30, 45)
            wb.set_pinned_now(time=moment)
            self.assertEqual(wb.get_pinned_now(), moment)
            with self.assertRaises(TypeError):
                wb.set_pinned_now(t=moment)  # type: ignore[call-arg]
            wb.set_pinned_now(time=None)
            self.assertIsNone(wb.get_pinned_now())


class SpellingTests(unittest.TestCase):
    def test_format_defaults_column_keywords(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.update_format_defaults(default_column_width_chars=12.5, base_column_width_chars=9.0)
            defaults = ws.get_format_defaults()
            self.assertEqual(defaults.default_column_width_chars, 12.5)
            self.assertEqual(defaults.base_column_width_chars, 9.0)
            self.assertEqual(defaults.effective_default_column_width_chars, 12.5)
            ws.update_format_defaults(default_column_width_chars=None)
            self.assertIsNone(ws.get_format_defaults().default_column_width_chars)
            for old in ("default_col_width_chars", "base_col_width_chars"):
                with self.assertRaises(TypeError):
                    ws.update_format_defaults(**{old: 10.0})  # type: ignore[arg-type]
            names = {f for f in vars(type(defaults))["__dataclass_fields__"]}
            self.assertTrue(
                {"default_column_width_chars", "base_column_width_chars", "effective_default_column_width_chars"}
                <= names
            )
            self.assertFalse(names & {"default_col_width_chars", "base_col_width_chars"})
            self.assertFalse("effective_default_col_width_chars" in names)

    def test_view_split_fields_spell_twips(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.update_view(split_x_twips=2000, split_y_twips=1500)
            view = ws.get_view()
            self.assertEqual((view.split_x_twips, view.split_y_twips), (2000.0, 1500.0))
            self.assertFalse(hasattr(view, "split_x"))


class PageBreakTests(unittest.TestCase):
    def test_page_break_input_defaults_to_whole_sheet(self) -> None:
        brk = PageBreakInput()
        self.assertEqual((brk.first, brk.last, brk.manual), (0, 0, True))
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.set_row_break(3, PageBreakInput())
            ws.set_column_break(2, PageBreakInput(first=1, last=9, manual=False))
            self.assertEqual(ws.list_row_breaks(), [RowBreak(row=3, first=0, last=16383, manual=True)])
            (column,) = ws.list_column_breaks()
            self.assertEqual((column.col, column.first, column.last, column.manual), (2, 1, 9, False))
            ws.set_row_break(3, None)
            ws.set_row_break(8, None)
            self.assertEqual(ws.list_row_breaks(), [])
            ws.set_column_break(2, None)
            self.assertEqual(ws.list_column_breaks(), [])
            with self.assertRaises(TypeError):
                ws.set_row_break(1, PageBreakInput(manual=1))  # type: ignore[arg-type]

    def test_read_break_passes_back(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.set_row_break(3, PageBreakInput(first=2, last=6, manual=False))
            (read,) = ws.list_row_breaks()
            ws.set_row_break(7, read)
            moved = {b.row: b for b in ws.list_row_breaks()}
            self.assertEqual((moved[7].first, moved[7].last, moved[7].manual), (2, 6, False))
            ws.set_column_break(4, read)
            (column,) = ws.list_column_breaks()
            self.assertEqual((column.col, column.first, column.last, column.manual), (4, 2, 6, False))


class PagingTests(unittest.TestCase):
    def test_cursor_and_limit_are_keyword_only(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            for row in range(5):
                ws.set_number(row, 0, float(row))
            first = ws.get_cells_in_range("A1:A5", limit=2)
            self.assertEqual([c.row for c in first.cells], [0, 1])
            self.assertIsNotNone(first.next_cursor)
            second = ws.get_cells_in_range("A1:A5", cursor=first.next_cursor, limit=2)
            self.assertEqual([c.row for c in second.cells], [2, 3])
            third = ws.get_cells_in_range("A1:A5", cursor=second.next_cursor, limit=2)
            self.assertEqual([c.row for c in third.cells], [4])
            self.assertIsNone(third.next_cursor)
            with self.assertRaises(TypeError):
                ws.get_cells_in_range("A1:A5", None, 2)  # type: ignore[misc]
            with self.assertRaises(TypeError):
                ws.get_invalid_cells(None, 1)  # type: ignore[misc]
            self.assertEqual(ws.get_invalid_cells(cursor=None, limit=1).cells, [])
            parameters = inspect.signature(type(ws).get_cells_in_range).parameters
            self.assertEqual(parameters["cursor"].kind, inspect.Parameter.KEYWORD_ONLY)
            self.assertEqual(parameters["limit"].kind, inspect.Parameter.KEYWORD_ONLY)
            parameters = inspect.signature(type(ws).get_invalid_cells).parameters
            self.assertEqual(parameters["cursor"].kind, inspect.Parameter.KEYWORD_ONLY)
            self.assertEqual(parameters["limit"].kind, inspect.Parameter.KEYWORD_ONLY)


def _with_comments(wb: Workbook) -> None:
    wb.add_person(PersonInput(id=ALICE, display_name="Alice"))
    ws = wb.sheet(0)
    ws.add_threaded_comment(ThreadedCommentInput(id=THREAD, cell="B2", person_id=ALICE, text="one"))
    ws.add_threaded_comment(ThreadedCommentInput(id=OTHER, cell="C3", person_id=ALICE, text="two"))


class ThreadedCommentListTests(unittest.TestCase):
    def test_row_and_a1(self) -> None:
        with Workbook.create_default() as wb:
            _with_comments(wb)
            ws = wb.sheet(0)
            self.assertEqual({c.id for c in ws.list_threaded_comments()}, {THREAD, OTHER})
            self.assertEqual([c.id for c in ws.list_threaded_comments(1, 1)], [THREAD])
            self.assertEqual([c.id for c in ws.list_threaded_comments("C3")], [OTHER])
            self.assertEqual(ws.list_threaded_comments(5, 5), [])

    def test_tuple_is_type_error(self) -> None:
        with Workbook.create_default() as wb:
            _with_comments(wb)
            ws = wb.sheet(0)
            for cell in ((1, 1), CellAddress(1, 1)):
                with self.assertRaises(TypeError) as caught:
                    ws.list_threaded_comments(cell)  # type: ignore[call-overload]
                self.assertEqual(
                    str(caught.exception), "Sheet.list_threaded_comments cell must be (row, col) or one A1 string"
                )


class CellFormatCopyTests(unittest.TestCase):
    def test_vars_splat_copies(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.update_cell_format("A1", number_format="0.00", font={"bold": True})
            ws.update_cell_format("B1", **vars(ws.get_cell_format("A1")))
            copy = ws.get_cell_format("B1")
            self.assertEqual(copy.number_format, "0.00")
            self.assertTrue(copy.font.bold)

    def test_source_not_in_signature(self) -> None:
        parameters = inspect.signature(formulon.Sheet.update_cell_format).parameters
        self.assertNotIn("source", parameters)
        self.assertEqual(parameters["read_only"].kind, inspect.Parameter.VAR_KEYWORD)
        with Workbook.create_default() as wb:
            with self.assertRaises(TypeError):
                wb.sheet(0).update_cell_format("A1", bogus=1)  # type: ignore[call-arg]


class SingleAxisGetterTests(unittest.TestCase):
    def test_a1_column_and_row(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.set_column_width(1, 1, 30.0)
            ws.set_row_height(2, 2, 31.0)
            self.assertEqual(ws.get_column_width_pt("B"), ws.get_column_width_pt(1))
            self.assertNotEqual(ws.get_column_width_pt("B"), ws.get_column_width_pt("A"))
            self.assertEqual(ws.get_row_height_pt("3"), ws.get_row_height_pt(2))
            self.assertEqual(ws.get_row_height_pt("3"), 31.0)
            self.assertEqual(ws.ooxml.get_column_xf_index("B"), ws.ooxml.get_column_xf_index(1))
            self.assertEqual(ws.ooxml.get_row_xf_index("3"), ws.ooxml.get_row_xf_index(2))

    def test_span_is_invalid_argument(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            with self.assertRaises(InvalidArgumentError) as caught:
                ws.get_column_width_pt("B:D")
            self.assertEqual(caught.exception.message, "Sheet.get_column_width_pt col must name one column")
            with self.assertRaises(InvalidArgumentError):
                ws.get_row_height_pt("3:5")
            with self.assertRaises(InvalidArgumentError):
                ws.ooxml.get_column_xf_index("B:D")
            with self.assertRaises(InvalidArgumentError):
                ws.ooxml.get_row_xf_index("3:5")


class ThreadedCommentCreatedTests(unittest.TestCase):
    def _created(self, created: datetime.date) -> datetime.datetime:
        with Workbook.create_default() as wb:
            wb.add_person(PersonInput(id=ALICE, display_name="Alice"))
            ws = wb.sheet(0)
            ws.add_threaded_comment(
                ThreadedCommentInput(id=THREAD, cell="A1", person_id=ALICE, text="t", created=created)
            )
            return ws.get_threaded_comment(THREAD).created

    def test_aware_datetime_converts_to_utc(self) -> None:
        tokyo = datetime.timezone(datetime.timedelta(hours=9))
        aware = datetime.datetime(2026, 11, 9, 17, 4, 48, tzinfo=tokyo)
        self.assertEqual(self._created(aware), datetime.datetime(2026, 11, 9, 8, 4, 48))
        west = datetime.timezone(datetime.timedelta(hours=-5))
        self.assertEqual(
            self._created(datetime.datetime(2026, 11, 9, 22, 0, 0, tzinfo=west)),
            datetime.datetime(2026, 11, 10, 3, 0, 0),
        )
        # The result is a naive UTC reading.
        self.assertIsNone(self._created(aware).tzinfo)

    def test_naive_datetime_is_utc(self) -> None:
        self.assertEqual(
            self._created(datetime.datetime(2026, 11, 9, 8, 4, 48)), datetime.datetime(2026, 11, 9, 8, 4, 48)
        )
        self.assertEqual(self._created(datetime.date(2026, 11, 9)), datetime.datetime(2026, 11, 9, 0, 0, 0))


if __name__ == "__main__":
    unittest.main()
