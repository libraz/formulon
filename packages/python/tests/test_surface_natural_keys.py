"""Lookups by natural key, the defined-name remove, and the index an index-keyed add returns."""

from __future__ import annotations

import unittest

from formulon import (
    AnchorKind,
    CellRect,
    DataValidationInput,
    DefinedName,
    FormulonError,
    Person,
    ThemeScriptFont,
    ThreadedComment,
    Workbook,
)

ALICE = "{0A1B2C3D-0000-4000-8000-000000000001}"
THREAD = "{6F1E0D2C-1111-4111-8111-000000000001}"


class NaturalKeyLookupTests(unittest.TestCase):
    def test_defined_names_are_read_and_removed_by_name_and_scope(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_defined_name("Rate", "Sheet1!$A$1")
            wb.set_defined_name_scoped("Rate", "Sheet1!$B$1", 0)
            self.assertEqual(
                wb.get_defined_name("rate"),
                DefinedName(name="Rate", formula="Sheet1!$A$1", local_sheet_id=-1, comment="", hidden=False),
            )
            scoped = wb.get_defined_name("Rate", 0)
            self.assertIsNotNone(scoped)
            assert scoped is not None
            self.assertEqual(scoped.formula, "Sheet1!$B$1")
            self.assertIsNone(wb.get_defined_name("Missing"))
            wb.remove_defined_name("Rate")
            self.assertIsNone(wb.get_defined_name("Rate"))
            self.assertIsNotNone(wb.get_defined_name("Rate", 0))
            with self.assertRaises(FormulonError):
                wb.remove_defined_name("Rate")

    def test_hyperlinks_people_and_threaded_comments_are_read_by_key(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_hyperlink(0, 2, 1, "https://example.com/")
            link = wb.get_hyperlink(0, 2, 1)
            self.assertIsNotNone(link)
            assert link is not None
            self.assertEqual(link.target, "https://example.com/")
            self.assertIsNone(wb.get_hyperlink(0, 2, 2))

            wb.add_person(Person(id=ALICE, display_name="Alice", user_id="alice", provider_id="None"))
            self.assertEqual(wb.get_person(ALICE), Person(ALICE, "Alice", "alice", "None"))
            self.assertIsNone(wb.get_person(THREAD))
            wb.add_threaded_comment(
                0,
                ThreadedComment(
                    id=THREAD, row=4, col=2, person_id=ALICE, created="2026-01-02T03:04:05.67", text="Check this"
                ),
            )
            thread = wb.get_threaded_comment(0, THREAD)
            self.assertIsNotNone(thread)
            assert thread is not None
            self.assertEqual((thread.row, thread.col, thread.text), (4, 2, "Check this"))
            self.assertIsNone(wb.get_threaded_comment(0, ALICE))

    def test_shapes_and_drawing_objects_are_read_by_object_id(self) -> None:
        with Workbook.create_default() as wb:
            object_id = wb.add_shape(0, anchor_kind=AnchorKind.ONE_CELL, from_row=2, cx=914400, cy=457200, prst="rect")
            shape = wb.get_shape(0, object_id)
            self.assertIsNotNone(shape)
            assert shape is not None
            self.assertEqual(shape.prst, "rect")
            drawing = wb.get_drawing_object(0, object_id)
            self.assertIsNotNone(drawing)
            assert drawing is not None
            self.assertEqual(drawing.from_row, 2)
            self.assertIsNone(wb.get_shape(0, object_id + 1))
            self.assertIsNone(wb.get_drawing_object(0, object_id + 1))

    def test_theme_script_fonts_and_passthrough_parts_are_read_by_key(self) -> None:
        with Workbook.create_default() as wb:
            jpan = wb.get_theme_script_font(True, "Jpan")
            self.assertIsNotNone(jpan)
            assert jpan is not None
            self.assertTrue(jpan.typeface)
            wb.set_theme_script_font(False, "Zzzz", "Test Face")
            self.assertEqual(wb.get_theme_script_font(False, "Zzzz"), ThemeScriptFont("Zzzz", "Test Face"))
            self.assertIsNone(wb.get_theme_script_font(True, "Zzzz"))
            path = next(wb.list_passthroughs()).path
            self.assertEqual(wb.get_passthrough(path), path)
            self.assertIsNone(wb.get_passthrough("/" + path))


class IndexKeyedAddTests(unittest.TestCase):
    def test_adds_return_the_index_they_allocate(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(wb.add_sheet("Second"), 1)
            rule = DataValidationInput(ranges=[CellRect(0, 0, 3, 0)], type=1, formula1="1", formula2="9")
            self.assertEqual(wb.add_validation(1, rule), 0)
            self.assertEqual(wb.add_validation(1, rule), 1)


if __name__ == "__main__":
    unittest.main()
