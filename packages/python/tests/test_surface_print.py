"""Print settings surface tests."""

from __future__ import annotations

import unittest

from formulon import (
    CellRect,
    CellXf,
    FormulonError,
    NotFoundError,
    PageOrientation,
    PrintTitles,
    Span,
    Workbook,
)


class PrintSettingsTests(unittest.TestCase):
    def test_raw_xml_roundtrip_and_removal(self) -> None:
        with Workbook.create_default() as wb:
            # An absent element reads back as "" rather than raising.
            self.assertIsNone(wb.sheet(0).get_page_setup_xml())
            fragment = '<pageSetup paperSize="9" orientation="portrait" scale="85"/>'
            wb.sheet(0).set_page_setup_xml(fragment)
            self.assertEqual(wb.sheet(0).get_page_setup_xml(), fragment)

            setup = wb.sheet(0).get_page_setup()
            self.assertEqual(setup.paper_size, 9)
            self.assertEqual(setup.scale, 85)
            self.assertTrue(setup.scale_stated)

            # Empty removes the element and restores the defaults.
            wb.sheet(0).set_page_setup_xml("")
            self.assertIsNone(wb.sheet(0).get_page_setup_xml())
            cleared = wb.sheet(0).get_page_setup()
            self.assertEqual(cleared.scale, 100)
            self.assertFalse(cleared.scale_stated)

    def test_malformed_fragment_is_rejected_at_set_time(self) -> None:
        with Workbook.create_default() as wb:
            for bad in ("<pageSetup/><pageSetup/>", '<pageMargins left="1"/>', '<pageSetup orientation="p"'):
                with self.assertRaises(FormulonError, msg=bad):
                    wb.sheet(0).set_page_setup_xml(bad)
            self.assertIsNone(wb.sheet(0).get_page_setup_xml())

    def test_fit_to_page_keeps_the_rest_of_sheet_pr(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_sheet_pr_xml('<sheetPr codeName="Sheet1"><tabColor rgb="FFFF0000"/></sheetPr>')
            wb.sheet(0).set_fit_to_page(True)
            xml = wb.sheet(0).get_sheet_pr_xml()
            # A VBA binding and a tab colour are not print settings, and
            # toggling fit-to-page must not be how they get lost.
            self.assertIn('codeName="Sheet1"', xml)
            self.assertIn('<tabColor rgb="FFFF0000"/>', xml)
            self.assertIn('fitToPage="true"', xml)
            self.assertTrue(wb.sheet(0).get_page_setup().fit_to_page)

    def test_typed_patch_preserves_unmodelled_attributes(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_page_setup_xml(
                '<pageSetup paperSize="9" orientation="portrait" horizontalDpi="600" copies="3"/>'
            )
            wb.sheet(0).set_page_setup(orientation=PageOrientation.LANDSCAPE)
            xml = wb.sheet(0).get_page_setup_xml()
            self.assertIn('orientation="landscape"', xml)
            self.assertIn('horizontalDpi="600"', xml)
            self.assertIn('copies="3"', xml)
            self.assertIn('paperSize="9"', xml)

    def test_scale_outside_excels_range_is_rejected(self) -> None:
        with Workbook.create_default() as wb:
            # Rejected rather than clamped: a mis-stated print scale lands
            # on paper, unlike the on-screen zoom.
            with self.assertRaises(FormulonError):
                wb.sheet(0).set_page_setup(scale=9)
            with self.assertRaises(FormulonError):
                wb.sheet(0).set_page_setup(scale=401)
            wb.sheet(0).set_page_setup(scale=400)
            self.assertEqual(wb.sheet(0).get_page_setup().scale, 400)

    def test_margins_patch_reports_presence_separately(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_page_margins(left=0.25, top=1.5)
            margins = wb.sheet(0).get_page_margins()
            self.assertAlmostEqual(margins.left, 0.25)
            self.assertAlmostEqual(margins.top, 1.5)
            self.assertTrue(margins.left_stated)
            # An unstated margin still reports its effective default, with
            # the flag clear so the caller can tell the two apart.
            self.assertAlmostEqual(margins.right, 0.7)
            self.assertFalse(margins.right_stated)

    def test_negative_margin_is_rejected(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(FormulonError):
                wb.sheet(0).set_page_margins(bottom=-1.0)

    def test_header_footer_sections_take_decoded_text(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_header_footer(odd_header="&C\u5e33\u7968", odd_footer="&R&P / &N")
            # `&` introduces Excel's codes but lives in XML element text, so
            # it reaches the file escaped and decodes back to `&C...`.
            self.assertEqual(
                wb.sheet(0).get_header_footer_xml(),
                "<headerFooter><oddHeader>&amp;C\u5e33\u7968</oddHeader>"
                "<oddFooter>&amp;R&amp;P / &amp;N</oddFooter></headerFooter>",
            )
            # Omitting a section leaves it; "" clears it.
            wb.sheet(0).set_header_footer(odd_footer="")
            self.assertEqual(
                wb.sheet(0).get_header_footer_xml(),
                "<headerFooter><oddHeader>&amp;C\u5e33\u7968</oddHeader></headerFooter>",
            )

    def test_print_options_patch(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_print_options(grid_lines=True, horizontal_centered=True)
            self.assertEqual(
                wb.sheet(0).get_print_options_xml(),
                '<printOptions gridLines="true" horizontalCentered="true"/>',
            )

    def test_print_area_and_titles(self) -> None:
        with Workbook.create_default() as wb:
            area = [CellRect(0, 0, 7, 5)]
            wb.sheet(0).set_print_area(area)
            self.assertEqual(wb.sheet(0).get_print_area(), area)
            with self.assertRaises(FormulonError):
                wb.sheet(0).set_print_area([CellRect(5, 0, 4, 0)])
            with self.assertRaises(TypeError):
                wb.sheet(0).set_print_area("A1:F8")  # type: ignore[arg-type]
            self.assertEqual(wb.sheet(0).get_print_area(), area)
            wb.sheet(0).set_print_area([])
            self.assertIsNone(wb.sheet(0).get_print_area())

            titles = PrintTitles(rows=Span(0, 1), columns=Span(0, 0))
            wb.sheet(0).set_print_titles(titles)
            self.assertEqual(wb.sheet(0).get_print_titles(), titles)
            wb.sheet(0).set_print_titles(None)
            self.assertIsNone(wb.sheet(0).get_print_titles())

    def test_manual_breaks(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).add_row_break(40)
            wb.sheet(0).add_row_break(10)
            wb.sheet(0).add_column_break(5)
            rows = wb.sheet(0).list_row_breaks()
            # Kept sorted by index regardless of insertion order.
            self.assertEqual([b.id for b in rows], [10, 40])
            self.assertEqual(rows[0].max, 16383)
            self.assertTrue(rows[0].manual)
            self.assertEqual(wb.sheet(0).list_column_breaks()[0].max, 1048575)

            # Re-adding an existing index replaces rather than duplicates.
            wb.sheet(0).add_row_break(10, manual=False)
            rows = wb.sheet(0).list_row_breaks()
            self.assertEqual(len(rows), 2)
            self.assertFalse(rows[0].manual)

            # Removing an absent break reports it as missing.
            with self.assertRaises(NotFoundError):
                wb.sheet(0).remove_row_break(999)
            wb.sheet(0).clear_breaks()
            self.assertEqual(wb.sheet(0).list_row_breaks(), [])
            self.assertEqual(wb.sheet(0).list_column_breaks(), [])

    def test_settings_reach_paginate_without_a_save_cycle(self) -> None:
        with Workbook.create_default() as wb:
            for row in range(200):
                for col in range(20):
                    wb.sheet(0).set_number(row, col, 1.0)
            portrait = wb.sheet(0).get_pagination().page_count
            wb.sheet(0).set_page_setup(orientation=PageOrientation.LANDSCAPE)
            self.assertNotEqual(wb.sheet(0).get_pagination().page_count, portrait)

    def test_authored_report_survives_save_and_load(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_text(0, 0, "\u58f2\u4e0a")
            wb.sheet(0).set_page_setup(
                paper_size=9, orientation=PageOrientation.PORTRAIT, fit_to_page=True, fit_to_width=1, fit_to_height=0
            )
            wb.sheet(0).set_page_margins(left=0.5, right=0.5, top=0.8, bottom=0.8)
            wb.sheet(0).set_header_footer(odd_header="&C\u6708\u6b21\u5831\u544a")
            wb.sheet(0).set_print_area([CellRect(0, 0, 79, 5)])
            wb.sheet(0).set_print_titles(PrintTitles(rows=Span(0, 1), columns=None))
            wb.sheet(0).add_row_break(39)
            data = wb.save()

        with Workbook.load(data) as reloaded:
            setup = reloaded.sheet(0).get_page_setup()
            self.assertEqual(setup.paper_size, 9)
            self.assertEqual(setup.orientation, PageOrientation.PORTRAIT)
            self.assertTrue(setup.fit_to_page)
            self.assertEqual(setup.fit_to_height, 0)
            self.assertAlmostEqual(reloaded.sheet(0).get_page_margins().left, 0.5)
            self.assertEqual(reloaded.sheet(0).get_print_area(), [CellRect(0, 0, 79, 5)])
            self.assertEqual(reloaded.sheet(0).get_print_titles(), PrintTitles(rows=Span(0, 1), columns=None))
            self.assertEqual([b.id for b in reloaded.sheet(0).list_row_breaks()], [39])

    def test_set_range_xf_index_materialises_the_rectangle(self) -> None:
        with Workbook.create_default() as wb:
            # Something other than the seeded default xf, so the range
            # write is observable rather than deduping to index 0.
            xf = wb.add_cell_xf(
                CellXf(
                    font_index=0,
                    fill_index=0,
                    border_index=0,
                    num_fmt_id=0,
                    horizontal_align=0,
                    vertical_align=2,
                    wrap_text=True,
                )
            )
            wb.sheet(0).set_range_xf_index(CellRect(0, 0, 2, 2), xf)
            for row, col in ((0, 0), (1, 1), (2, 2)):
                self.assertEqual(wb.sheet(0).get_cell_xf_index(row, col), xf)
            self.assertEqual(wb.sheet(0).get_cell_xf_index(3, 3), 0)


if __name__ == "__main__":
    unittest.main()
