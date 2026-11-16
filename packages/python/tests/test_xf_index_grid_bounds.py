"""Style XF reads of one row or column refuse a coordinate the grid cannot hold."""

from __future__ import annotations

import unittest

from formulon import (
    InvalidArgumentError,
    Workbook,
)


class GridCoordinateTests(unittest.TestCase):
    def test_xf_index_reads_refuse_a_coordinate_beyond_the_grid(self) -> None:
        with Workbook.create_default() as wb:
            ooxml = wb.sheet(0).ooxml
            self.assertIsNone(ooxml.get_row_xf_index(1048575))
            self.assertIsNone(ooxml.get_column_xf_index(16383))
            with self.assertRaises(InvalidArgumentError):
                ooxml.get_row_xf_index(1048576)
            with self.assertRaises(InvalidArgumentError):
                ooxml.get_column_xf_index(16384)


if __name__ == "__main__":
    unittest.main()
