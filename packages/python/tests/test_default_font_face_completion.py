"""A bare rename of the default font saves with the family and charset Excel completes for the face."""

from __future__ import annotations

import unittest
from typing import Callable, Optional

from formulon import Font, FontFamily, Workbook

# Mac Excel 16.115: (face, family, charset), None when the font states none.
MEASURED = [
    ("Calibri", 2, None),
    ("Aptos", None, None),
    ("游明朝", 1, 128),
    ("等线", 4, 134),
    ("맑은 고딕", 2, 129),
    ("新細明體", 1, 136),
    ("ＭＳ Ｐゴシック", 2, 128),
    ("Meiryo", 2, 128),
    ("Arial", 2, None),
    ("Cambria", 1, None),
    ("Courier New", 1, None),
]


def _saved_default_font(edit: Callable[[Workbook], None]) -> Font:
    with Workbook.create_for_profile("mac-365-ja_JP") as wb:
        edit(wb)
        data = wb.save()
    # A theme-linked font is named for the profile a book is opened under, so reopen under the same one.
    with Workbook.load(data, profile_id="mac-365-ja_JP") as back:
        return back.get_default_font()


def _family(font: Font) -> Optional[int]:
    return None if font.family is None else int(font.family)


class DefaultFontFaceCompletionTest(unittest.TestCase):
    def test_a_bare_rename_saves_the_completed_family_and_charset(self) -> None:
        for face, family, charset in MEASURED:
            with self.subTest(face=face):
                font = _saved_default_font(lambda wb, face=face: wb.set_default_font(Font(name=face)))
                self.assertEqual(font.name, face)
                self.assertEqual(_family(font), family)
                self.assertEqual(font.charset, charset)

    def test_a_rename_to_yu_gothic_from_another_face_is_family_three(self) -> None:
        def edit(wb: Workbook) -> None:
            wb.set_default_font(Font(name="Arial"))
            wb.set_default_font(Font(name="游ゴシック"))

        font = _saved_default_font(edit)
        self.assertEqual(_family(font), 3)
        self.assertEqual(font.charset, 128)

    def test_a_stated_charset_or_family_is_written_as_given(self) -> None:
        charset = _saved_default_font(lambda wb: wb.set_default_font(Font(name="Courier New", charset=128)))
        self.assertIsNone(charset.family)
        self.assertEqual(charset.charset, 128)
        family = _saved_default_font(lambda wb: wb.set_default_font(Font(name="Arial", family=FontFamily.ROMAN)))
        self.assertEqual(_family(family), 1)
        self.assertIsNone(family.charset)

    def test_the_stock_default_keeps_its_own_record(self) -> None:
        font = _saved_default_font(lambda wb: None)
        self.assertEqual(font.name, "游ゴシック")
        self.assertEqual(_family(font), 2)
        self.assertEqual(font.charset, 128)


if __name__ == "__main__":
    unittest.main()
