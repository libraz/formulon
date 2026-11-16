"""A position past uint32_t but within uint64_t is out of range, not a ValueError."""

from __future__ import annotations

import unittest

from formulon import IndexOutOfRangeError, StatusCode, Workbook


class PositionBeyondU32Test(unittest.TestCase):
    def test_a_style_record_position_past_u32_is_out_of_range(self) -> None:
        with Workbook.create_default() as wb:
            for value in (2**32, 2**64 - 1):
                with self.assertRaises(IndexOutOfRangeError) as caught:
                    wb.ooxml.get_font(value)
                self.assertEqual(caught.exception.code, StatusCode.INDEX_OUT_OF_RANGE)

    def test_a_position_past_u64_stays_a_value_error(self) -> None:
        with Workbook.create_default() as wb, self.assertRaises(ValueError):
            wb.ooxml.get_font(2**64)


if __name__ == "__main__":
    unittest.main()
