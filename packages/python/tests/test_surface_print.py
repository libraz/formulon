"""Print settings surface tests."""

from __future__ import annotations

import unittest

from formulon import (
    CellXf,
    FormulonError,
    Workbook,
)


class PrintSettingsTests(unittest.TestCase):
    def test_raw_xml_roundtrip_and_removal(self) -> None:
        with Workbook.create_default() as wb:
            # An absent element reads back as "" rather than raising.
            self.assertEqual(wb.get_page_setup_xml(0), "")
            fragment = '<pageSetup paperSize="9" orientation="portrait" scale="85"/>'
            wb.set_page_setup_xml(0, fragment)
            self.assertEqual(wb.get_page_setup_xml(0), fragment)

            setup = wb.get_page_setup(0)
            self.assertEqual(setup.paper_size, 9)
            self.assertEqual(setup.scale, 85)
            self.assertTrue(setup.scale_stated)

            # Empty removes the element and restores the defaults.
            wb.set_page_setup_xml(0, "")
            self.assertEqual(wb.get_page_setup_xml(0), "")
            cleared = wb.get_page_setup(0)
            self.assertEqual(cleared.scale, 100)
            self.assertFalse(cleared.scale_stated)

    def test_malformed_fragment_is_rejected_at_set_time(self) -> None:
        with Workbook.create_default() as wb:
            for bad in ("<pageSetup/><pageSetup/>", '<pageMargins left="1"/>', '<pageSetup orientation="p"'):
                with self.assertRaises(FormulonError, msg=bad):
                    wb.set_page_setup_xml(0, bad)
            self.assertEqual(wb.get_page_setup_xml(0), "")

    def test_fit_to_page_keeps_the_rest_of_sheet_pr(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_sheet_pr_xml(0, '<sheetPr codeName="Sheet1"><tabColor rgb="FFFF0000"/></sheetPr>')
            wb.set_fit_to_page(0, True)
            xml = wb.get_sheet_pr_xml(0)
            # A VBA binding and a tab colour are not print settings, and
            # toggling fit-to-page must not be how they get lost.
            self.assertIn('codeName="Sheet1"', xml)
            self.assertIn('<tabColor rgb="FFFF0000"/>', xml)
            self.assertIn('fitToPage="true"', xml)
            self.assertTrue(wb.get_page_setup(0).fit_to_page)

    def test_typed_patch_preserves_unmodelled_attributes(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_page_setup_xml(0, '<pageSetup paperSize="9" orientation="portrait" horizontalDpi="600" copies="3"/>')
            wb.set_page_setup(0, orientation=2)
            xml = wb.get_page_setup_xml(0)
            self.assertIn('orientation="landscape"', xml)
            self.assertIn('horizontalDpi="600"', xml)
            self.assertIn('copies="3"', xml)
            self.assertIn('paperSize="9"', xml)

    def test_scale_outside_excels_range_is_rejected(self) -> None:
        with Workbook.create_default() as wb:
            # Rejected rather than clamped: a mis-stated print scale lands
            # on paper, unlike the on-screen zoom.
            with self.assertRaises(FormulonError):
                wb.set_page_setup(0, scale=9)
            with self.assertRaises(FormulonError):
                wb.set_page_setup(0, scale=401)
            wb.set_page_setup(0, scale=400)
            self.assertEqual(wb.get_page_setup(0).scale, 400)

    def test_margins_patch_reports_presence_separately(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_page_margins(0, left=0.25, top=1.5)
            margins = wb.get_page_margins(0)
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
                wb.set_page_margins(0, bottom=-1.0)

    def test_header_footer_sections_take_decoded_text(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_header_footer(0, odd_header="&C\u5e33\u7968", odd_footer="&R&P / &N")
            # `&` introduces Excel's codes but lives in XML element text, so
            # it reaches the file escaped and decodes back to `&C...`.
            self.assertEqual(
                wb.get_header_footer_xml(0),
                "<headerFooter><oddHeader>&amp;C\u5e33\u7968</oddHeader>"
                "<oddFooter>&amp;R&amp;P / &amp;N</oddFooter></headerFooter>",
            )
            # Omitting a section leaves it; "" clears it.
            wb.set_header_footer(0, odd_footer="")
            self.assertEqual(
                wb.get_header_footer_xml(0),
                "<headerFooter><oddHeader>&amp;C\u5e33\u7968</oddHeader></headerFooter>",
            )

    def test_print_options_patch(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_print_options(0, grid_lines=True, horizontal_centered=True)
            self.assertEqual(
                wb.get_print_options_xml(0),
                '<printOptions gridLines="true" horizontalCentered="true"/>',
            )

    def test_print_area_and_titles(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_print_area(0, "A1:F8")
            self.assertEqual(wb.get_print_area(0), "A1:F8")
            with self.assertRaises(FormulonError):
                wb.set_print_area(0, "not-a-range")
            self.assertEqual(wb.get_print_area(0), "A1:F8")
            wb.set_print_area(0, "")
            self.assertEqual(wb.get_print_area(0), "")

            wb.set_print_titles(0, "1:2", "A:A")
            self.assertEqual(wb.get_print_titles(0), ("1:2", "A:A"))
            wb.set_print_titles(0)
            self.assertEqual(wb.get_print_titles(0), ("", ""))

    def test_manual_breaks(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_row_break(0, 40)
            wb.add_row_break(0, 10)
            wb.add_col_break(0, 5)
            rows = wb.get_row_breaks(0)
            # Kept sorted by index regardless of insertion order.
            self.assertEqual([b.id for b in rows], [10, 40])
            self.assertEqual(rows[0].max, 16383)
            self.assertTrue(rows[0].manual)
            self.assertEqual(wb.get_col_breaks(0)[0].max, 1048575)

            # Re-adding an existing index replaces rather than duplicates.
            wb.add_row_break(0, 10, manual=False)
            rows = wb.get_row_breaks(0)
            self.assertEqual(len(rows), 2)
            self.assertFalse(rows[0].manual)

            # Removing an absent break is not an error.
            wb.remove_row_break(0, 999)
            wb.clear_breaks(0)
            self.assertEqual(wb.get_row_breaks(0), [])
            self.assertEqual(wb.get_col_breaks(0), [])

    def test_settings_reach_paginate_without_a_save_cycle(self) -> None:
        with Workbook.create_default() as wb:
            for row in range(200):
                for col in range(20):
                    wb.set_number(0, row, col, 1.0)
            portrait = wb.paginate(0).page_count
            wb.set_page_setup(0, orientation=2)
            self.assertNotEqual(wb.paginate(0).page_count, portrait)

    def test_authored_report_survives_save_and_load(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_text(0, 0, 0, "\u58f2\u4e0a")
            wb.set_page_setup(0, paper_size=9, orientation=1, fit_to_page=True, fit_to_width=1, fit_to_height=0)
            wb.set_page_margins(0, left=0.5, right=0.5, top=0.8, bottom=0.8)
            wb.set_header_footer(0, odd_header="&C\u6708\u6b21\u5831\u544a")
            wb.set_print_area(0, "A1:F80")
            wb.set_print_titles(0, "1:2")
            wb.add_row_break(0, 39)
            data = wb.save()

        with Workbook.load(data) as reloaded:
            setup = reloaded.get_page_setup(0)
            self.assertEqual(setup.paper_size, 9)
            self.assertEqual(setup.orientation, 1)
            self.assertTrue(setup.fit_to_page)
            self.assertEqual(setup.fit_to_height, 0)
            self.assertAlmostEqual(reloaded.get_page_margins(0).left, 0.5)
            self.assertEqual(reloaded.get_print_area(0), "A1:F80")
            self.assertEqual(reloaded.get_print_titles(0), ("1:2", ""))
            self.assertEqual([b.id for b in reloaded.get_row_breaks(0)], [39])

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
                    has_wrap_text=True,
                )
            )
            wb.set_range_xf_index(0, 0, 0, 2, 2, xf)
            for row, col in ((0, 0), (1, 1), (2, 2)):
                self.assertEqual(wb.get_cell_xf_index(0, row, col), xf)
            self.assertEqual(wb.get_cell_xf_index(0, 3, 3), 0)


if __name__ == "__main__":
    unittest.main()
