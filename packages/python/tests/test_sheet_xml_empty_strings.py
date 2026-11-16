"""Empty-value rules of the raw sheet-XML setters, the phonetic setter and the header/footer patch.

``""`` is refused where it is not one element, and means "no value" where Excel stores one state for both.
"""

from __future__ import annotations

import unittest
from typing import Any, Callable, Optional

from formulon import InvalidArgumentError, Sheet, Workbook

XML_SETTERS = (
    ("auto_filter", '<autoFilter ref="A1:C10"/>'),
    ("page_setup", '<pageSetup paperSize="9" orientation="landscape"/>'),
    ("page_margins", '<pageMargins left="1" right="1" top="1" bottom="1" header="0.5" footer="0.5"/>'),
    ("print_options", '<printOptions gridLines="1"/>'),
    ("header_footer", "<headerFooter><oddHeader>&amp;CTitle</oddHeader></headerFooter>"),
    ("sheet_pr", '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>'),
)


def _accessors(sheet: Sheet, name: str) -> "tuple[Callable[[], Optional[str]], Callable[[Any], None]]":
    ooxml = sheet.ooxml
    return getattr(ooxml, f"get_{name}_xml"), getattr(ooxml, f"set_{name}_xml")


class SheetXmlEmptyStringTests(unittest.TestCase):
    def test_empty_string_is_refused_and_changes_nothing(self) -> None:
        for name, fragment in XML_SETTERS:
            with self.subTest(setter=name), Workbook.create_default() as wb:
                get, put = _accessors(wb.sheet(0), name)
                put(fragment)
                before = get()
                self.assertIsNotNone(before)
                with self.assertRaises(InvalidArgumentError):
                    put("")
                self.assertEqual(get(), before)

    def test_empty_string_on_an_absent_element_is_refused(self) -> None:
        for name, _fragment in XML_SETTERS:
            with self.subTest(setter=name), Workbook.create_default() as wb:
                get, put = _accessors(wb.sheet(0), name)
                before = get()
                with self.assertRaises(InvalidArgumentError):
                    put("")
                self.assertEqual(get(), before)

    def test_none_removes_the_element(self) -> None:
        for name, fragment in XML_SETTERS:
            with self.subTest(setter=name), Workbook.create_default() as wb:
                get, put = _accessors(wb.sheet(0), name)
                put(fragment)
                self.assertIsNotNone(get())
                put(None)
                self.assertIsNone(get())
                # Removing what is not there is not an error.
                put(None)
                self.assertIsNone(get())

    def test_a_non_string_is_a_type_error(self) -> None:
        for name, _fragment in XML_SETTERS:
            with self.subTest(setter=name), Workbook.create_default() as wb:
                _get, put = _accessors(wb.sheet(0), name)
                with self.assertRaises(TypeError):
                    put(123)

    def test_none_resets_the_structured_page_setup_and_empty_string_leaves_it(self) -> None:
        with Workbook.create_default() as wb:
            ooxml = wb.sheet(0).ooxml
            ooxml.set_page_setup_xml('<pageSetup orientation="landscape" scale="50"/>')
            with self.assertRaises(InvalidArgumentError):
                ooxml.set_page_setup_xml("")
            self.assertEqual(wb.sheet(0).get_page_setup().scale_pct, 50)
            ooxml.set_page_setup_xml(None)
            reset = wb.sheet(0).get_page_setup()
            self.assertEqual(reset.scale_pct, 100)
            self.assertFalse(reset.scale_pct_stated)

    def test_cell_phonetic_empty_string_and_none_both_clear(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.set_text("A1", "kanji")
            for empty in ("", None):
                with self.subTest(empty=empty):
                    ws.set_cell_phonetic("A1", "kana")
                    self.assertEqual(ws.get_cell_phonetic("A1"), "kana")
                    ws.set_cell_phonetic("A1", empty)
                    self.assertIsNone(ws.get_cell_phonetic("A1"))

    def test_header_footer_section_empty_string_and_none_both_remove(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            for empty in ("", None):
                with self.subTest(empty=empty):
                    ws.update_header_footer(odd_header="&CHead", odd_footer="&CFoot")
                    self.assertEqual(ws.get_header_footer().odd_header, "&CHead")
                    ws.update_header_footer(odd_header=empty)
                    read = ws.get_header_footer()
                    self.assertIsNone(read.odd_header)
                    self.assertEqual(read.odd_footer, "&CFoot")
                    ws.update_header_footer(odd_footer=empty)
                    self.assertIsNone(ws.get_header_footer().odd_footer)


if __name__ == "__main__":
    unittest.main()
