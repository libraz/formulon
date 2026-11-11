"""Hyperlink model through the Python binding: ranges, containment, whole removal and cell formats."""

from __future__ import annotations

import unittest
from typing import List

import formulon
from formulon import (
    CellRect,
    FormulaSyntaxError,
    HyperlinkInput,
    IndexOutOfRangeError,
    InvalidArgumentError,
    NotFoundError,
    Sheet,
    Workbook,
)

URL = "https://example.com/"


def _link(range_: object, display: object = None, **fields: object) -> HyperlinkInput:
    return HyperlinkInput(range=range_, target=URL, display=display, **fields)  # type: ignore[arg-type]


def _displays(ws: Sheet) -> List[object]:
    return [h.display for h in ws.list_hyperlinks()]


def _xf(ws: Sheet, a1: str) -> int:
    return ws.ooxml.get_cell_xf_index(a1) or 0


class _Base(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.addCleanup(self.wb.close)
        self.ws = self.wb.sheet(0)


class SetTests(_Base):
    def test_new_range_appends_and_equal_range_replaces_in_place(self) -> None:
        self.ws.set_hyperlink(_link("A1", "one"))
        self.ws.set_hyperlink(_link("B1:C1", "two"))
        self.ws.set_hyperlink(_link("A1", "one again", tooltip="tip"))
        self.assertEqual(_displays(self.ws), ["one again", "two"])
        self.assertEqual(self.ws.list_hyperlinks()[0].tooltip, "tip")
        self.assertEqual([h.range for h in self.ws.list_hyperlinks()], [CellRect(0, 0, 0, 0), CellRect(0, 1, 0, 2)])

    def test_overlapping_ranges_coexist_in_insertion_order(self) -> None:
        self.ws.set_hyperlink(_link("A1:C3", "wide"))
        self.ws.set_hyperlink(_link("B2", "cell"))
        self.ws.set_hyperlink(_link("B2:B3", "sub"))
        self.ws.set_hyperlink(_link("C3:D4", "partial"))
        self.assertEqual(_displays(self.ws), ["wide", "cell", "sub", "partial"])
        self.assertEqual(self.ws.list_hyperlinks()[0].range, CellRect(0, 0, 2, 2))
        self.assertEqual(self.ws.list_hyperlinks()[3].range, CellRect(2, 2, 3, 3))

    def test_round_trip_of_a_read_link_leaves_the_list_unchanged(self) -> None:
        self.ws.set_value("A1", "seed")
        self.ws.set_value("E5", "seed")
        self.ws.set_hyperlink(_link("A1:B2", "d", tooltip="t"))
        self.ws.set_hyperlink(HyperlinkInput(range=CellRect(4, 4, 4, 4), target=None, location="Sheet1!A1"))
        before = self.ws.list_hyperlinks()
        link = self.ws.get_hyperlink("A1")
        assert link is not None
        self.ws.set_hyperlink(link)
        other = self.ws.get_hyperlink(4, 4)
        assert other is not None
        self.ws.set_hyperlink(other)
        self.ws.set_hyperlink(before[1])
        self.assertEqual(self.ws.list_hyperlinks(), before)

    def test_missing_range_and_wrong_types_are_type_errors(self) -> None:
        with self.assertRaisesRegex(TypeError, r"Sheet\.set_hyperlink link\.range is required"):
            self.ws.set_hyperlink(HyperlinkInput(range=None, target=URL))  # type: ignore[arg-type]
        with self.assertRaises(TypeError):
            self.ws.set_hyperlink(None)  # type: ignore[arg-type]
        with self.assertRaises(TypeError):
            self.ws.set_hyperlink("A1", _link("A1"))  # type: ignore[call-arg,arg-type]
        with self.assertRaises(TypeError):
            self.ws.set_hyperlink(formulon.CommentInput("x"))  # type: ignore[arg-type]
        self.assertEqual(self.ws.list_hyperlinks(), [])

    def test_bad_ranges_are_refused(self) -> None:
        for bad in (CellRect(2, 2, 0, 0), CellRect(1048576, 0, 1048576, 0)):
            with self.subTest(range=bad):
                with self.assertRaises(InvalidArgumentError):
                    self.ws.set_hyperlink(_link(bad))
        for bad in ("A1048577", "not a range"):
            with self.subTest(range=bad):
                with self.assertRaises(FormulaSyntaxError):
                    self.ws.set_hyperlink(_link(bad))
        self.assertEqual(self.ws.list_hyperlinks(), [])

    def test_lone_surrogate_in_each_text_field_becomes_u_fffd(self) -> None:
        for field in ("target", "location", "display", "tooltip"):
            fields = {"target": URL, "location": None, "display": None, "tooltip": None}
            fields[field] = "x\ud800"
            with self.subTest(field=field):
                self.ws.clear_hyperlinks()
                self.ws.set_hyperlink(HyperlinkInput(range="A1", **fields))  # type: ignore[arg-type]
                self.assertEqual(getattr(self.ws.get_hyperlink("A1"), field), "x\ufffd")

    def test_text_goes_into_an_empty_top_left_cell_only(self) -> None:
        self.ws.set_value("B1", "kept")
        self.ws.set_hyperlink(_link("A1", "shown"))
        self.ws.set_hyperlink(_link("B1", "ignored"))
        self.ws.set_hyperlink(_link("C1"))
        self.ws.set_hyperlink(HyperlinkInput(range="D1", target=None, location="Sheet1!A1"))
        self.assertEqual(self.ws.get_value("A1"), "shown")
        self.assertEqual(self.ws.get_value("B1"), "kept")
        self.assertEqual(self.ws.get_value("C1"), URL)
        self.assertEqual(self.ws.get_value("D1"), "Sheet1!A1")


class GetTests(_Base):
    def test_earliest_link_containing_the_cell_wins(self) -> None:
        self.ws.set_hyperlink(_link("A1:C3", "wide"))
        self.ws.set_hyperlink(_link("B2", "cell"))
        self.assertEqual(self.ws.get_hyperlink("B2").display, "wide")
        self.assertEqual(self.ws.get_hyperlink(1, 1).display, "wide")
        for a1 in ("A1", "C1", "A3", "C3"):
            self.assertEqual(self.ws.get_hyperlink(a1).display, "wide", a1)
        self.assertIsNone(self.ws.get_hyperlink("D4"))
        self.ws.clear_hyperlinks()
        self.ws.set_hyperlink(_link("B2", "cell"))
        self.ws.set_hyperlink(_link("A1:C3", "wide"))
        self.assertEqual(self.ws.get_hyperlink("B2").display, "cell")
        self.assertEqual(self.ws.get_hyperlink("A1").display, "wide")

    def test_cell_beyond_the_grid_is_invalid_argument_and_the_last_cell_reads(self) -> None:
        with self.assertRaises(InvalidArgumentError):
            self.ws.get_hyperlink(1048576, 0)
        self.assertIsNone(self.ws.get_hyperlink(1048575, 16383))


class RemoveTests(_Base):
    def test_remove_by_range_removes_every_intersecting_link_whole(self) -> None:
        self.ws.set_hyperlink(_link("A1:C3", "wide"))
        self.ws.set_hyperlink(_link("B2", "cell"))
        self.ws.set_hyperlink(_link("E5", "far"))
        self.ws.remove_hyperlink("B2")
        self.assertEqual(_displays(self.ws), ["far"])
        with self.assertRaises(NotFoundError):
            self.ws.remove_hyperlink("B2")
        self.ws.set_hyperlink(_link("A1:C3", "wide"))
        self.ws.remove_hyperlink(CellRect(2, 2, 6, 6))
        self.assertEqual(_displays(self.ws), [])

    def test_remove_refuses_a_bad_range(self) -> None:
        with self.assertRaises(InvalidArgumentError):
            self.ws.remove_hyperlink(CellRect(2, 2, 0, 0))
        with self.assertRaises(FormulaSyntaxError):
            self.ws.remove_hyperlink("A1048577")

    def test_remove_at_removes_only_that_entry(self) -> None:
        self.ws.set_hyperlink(_link("A1:C3", "wide"))
        self.ws.set_hyperlink(_link("B2", "cell"))
        self.ws.remove_hyperlink_at(0)
        self.assertEqual(_displays(self.ws), ["cell"])
        with self.assertRaises(IndexOutOfRangeError):
            self.ws.remove_hyperlink_at(1)
        with self.assertRaises(IndexOutOfRangeError):
            self.ws.remove_hyperlink_at(-1)
        self.assertEqual(_displays(self.ws), ["cell"])

    def test_clear_empties_the_list_and_is_a_no_op_when_empty(self) -> None:
        self.ws.clear_hyperlinks()
        self.ws.set_hyperlink(_link("A1"))
        self.ws.set_hyperlink(_link("B2"))
        self.ws.clear_hyperlinks()
        self.assertEqual(self.ws.list_hyperlinks(), [])
        self.ws.clear_hyperlinks()


class InsertTests(_Base):
    def test_insert_keeps_an_equal_range_appends_at_the_count_and_returns_none(self) -> None:
        self.assertIsNone(self.ws.insert_hyperlink_at(0, _link("A1", "b")))
        self.ws.insert_hyperlink_at(0, _link("A1", "a"))
        self.ws.insert_hyperlink_at(1, _link("A1", "mid"))
        self.ws.insert_hyperlink_at(3, _link("A1", "end"))
        self.assertEqual(_displays(self.ws), ["a", "mid", "b", "end"])

    def test_insert_refusals_change_nothing(self) -> None:
        self.ws.insert_hyperlink_at(0, _link("A1"))
        with self.assertRaises(IndexOutOfRangeError):
            self.ws.insert_hyperlink_at(2, _link("A1"))
        with self.assertRaises(FormulaSyntaxError):
            self.ws.insert_hyperlink_at(0, _link("A1048577"))
        with self.assertRaisesRegex(TypeError, r"Sheet\.insert_hyperlink_at link\.range is required"):
            self.ws.insert_hyperlink_at(0, HyperlinkInput(range=None, target=URL))  # type: ignore[arg-type]
        self.assertEqual(len(self.ws.list_hyperlinks()), 1)


class StyleTests(_Base):
    def test_set_formats_the_range_and_remove_resets_it_to_the_default_format(self) -> None:
        self.ws.set_hyperlink(_link("A1:C3"))
        for a1 in ("A1", "B2", "C3", "C1"):
            self.assertNotEqual(_xf(self.ws, a1), 0, a1)
        self.assertEqual(_xf(self.ws, "D4"), 0)
        self.ws.remove_hyperlink("B2")
        for a1 in ("A1", "B2", "C3"):
            self.assertEqual(_xf(self.ws, a1), 0, a1)
        self.assertEqual(self.ws.list_hyperlinks(), [])

    def test_removing_a_link_resets_cells_a_remaining_link_still_covers(self) -> None:
        self.ws.set_hyperlink(_link("A1:C3", "wide"))
        self.ws.set_hyperlink(_link("B2", "cell"))
        self.ws.remove_hyperlink("A1")
        self.assertEqual(_displays(self.ws), ["cell"])
        self.assertEqual(self.ws.get_hyperlink("B2").display, "cell")
        self.assertEqual(_xf(self.ws, "B2"), 0)
        self.assertEqual(_xf(self.ws, "A1"), 0)

    def test_insert_and_remove_at_leave_cell_formats_unchanged(self) -> None:
        self.ws.set_hyperlink(_link("A1:B2"))

        def formats() -> List[int]:
            return [_xf(self.ws, a1) for a1 in ("A1", "B2", "C3")]

        before = formats()
        self.ws.insert_hyperlink_at(1, _link("C3"))
        self.assertEqual(formats(), before)
        self.ws.remove_hyperlink_at(0)
        self.assertEqual(formats(), before)
        self.assertNotEqual(before[0], 0)


if __name__ == "__main__":
    unittest.main()
