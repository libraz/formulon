"""A sparkline group whose date-axis text is not one rectangle writes that text back when the group read is added."""

from __future__ import annotations

import io
import re
import unittest
import zipfile

from formulon import DefinedNameInput, SparklineGroupInput, SparklineInput, Workbook


def _saved_range_texts(wb: Workbook) -> "list[str]":
    sheet = zipfile.ZipFile(io.BytesIO(wb.save())).read("xl/worksheets/sheet1.xml").decode()
    return re.findall(r"<xm:f>([^<]*)</xm:f>", sheet)


class SparklineDateAxisRoundTripTests(unittest.TestCase):
    def test_a_group_read_back_and_added_again_keeps_its_date_axis_text(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_defined_name(DefinedNameInput(name="Dates", formula="Sheet1!$A$1:$D$1"))
            ws = wb.sheet(0)
            group_id = ws.add_sparkline_group(
                SparklineGroupInput(
                    date_axis=True, date_axis_range="Dates", sparklines=[SparklineInput("Sheet1!A2:D2", "F2")]
                )
            )
            read = ws.get_sparkline_group(group_id)
            self.assertIsNone(read.date_axis_range)
            ws.remove_sparkline_group(group_id)
            ws.add_sparkline_group(read)
            self.assertEqual(_saved_range_texts(wb), ["Dates", "Sheet1!A2:D2"])


if __name__ == "__main__":
    unittest.main()
