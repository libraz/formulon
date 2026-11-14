"""Features the xlsb writer cannot store fail the save with IoFeatureNotWritable; inputs outside Excel's limits are
refused with InvalidArgumentError."""

from __future__ import annotations

import unittest
from pathlib import Path

from formulon import (
    CellRect,
    InvalidArgumentError,
    InvalidFileError,
    IoError,
    ProtectedRangeInput,
    SparklineGroupInput,
    SparklineInput,
    StatusCode,
    Workbook,
    WorkbookFormat,
)

FIXTURES = Path(__file__).resolve().parents[3] / "tests" / "fixtures" / "excel" / "xlsb_parity"


def _cells(count: int) -> list:
    return [CellRect(i, 0, i, 0) for i in range(count)]


class XlsbNotWritableTests(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.addCleanup(self.wb.close)
        self.ws = self.wb.sheet(0)

    def test_xlsb_save_of_a_range_with_a_security_descriptor_fails_and_names_the_feature(self) -> None:
        self.ws.add_protected_range(
            ProtectedRangeInput(name="guarded", ranges="A1:B2", security_descriptor="O:WDG:WDD:(A;;CC;;;WD)")
        )
        with self.assertRaises(IoError) as cm:
            self.wb.save(format=WorkbookFormat.XLSB)
        self.assertNotIsInstance(cm.exception, InvalidFileError)
        self.assertEqual(cm.exception.code, StatusCode.IO_FEATURE_NOT_WRITABLE)
        self.assertIn("kind=protected_range_descriptor", cm.exception.context)
        self.assertIn("sheet_index=0", cm.exception.context)
        self.assertTrue(self.wb.save(format=WorkbookFormat.XLSX))

    def test_protected_range_name_of_255_units_is_accepted_and_256_is_refused(self) -> None:
        self.assertEqual(self.ws.add_protected_range(ProtectedRangeInput(name="a" * 255, ranges="A1:B2")), "a" * 255)
        with self.assertRaises(InvalidArgumentError) as cm:
            self.ws.add_protected_range(ProtectedRangeInput(name="b" * 256, ranges="A1:B2"))
        self.assertEqual(cm.exception.code, StatusCode.INVALID_ARGUMENT)
        # An astral character counts as two units: 128 of them are 256.
        with self.assertRaises(InvalidArgumentError):
            self.ws.add_protected_range(ProtectedRangeInput(name="\U0001f600" * 128, ranges="A1:B2"))

    def test_sparkline_union_source_is_refused_and_a_3d_source_is_accepted(self) -> None:
        self.wb.add_sheet("Sheet2")
        with self.assertRaises(InvalidArgumentError) as cm:
            self.ws.add_sparkline_group(
                SparklineGroupInput(sparklines=[SparklineInput("Sheet1!A1:A3,Sheet1!C1:C3", "E1")])
            )
        self.assertEqual(cm.exception.code, StatusCode.INVALID_ARGUMENT)
        self.ws.add_sparkline_group(SparklineGroupInput(sparklines=[SparklineInput("Sheet1:Sheet2!A1:A3", "E1")]))
        self.assertEqual(len(self.ws.list_sparkline_groups()), 1)

    def test_xlsb_save_of_a_sparkline_with_a_union_source_fails_and_names_the_feature(self) -> None:
        with Workbook.load(FIXTURES / "sparklines" / "union_source_src.xlsx") as wb:
            with self.assertRaises(IoError) as cm:
                wb.save(format=WorkbookFormat.XLSB)
            self.assertEqual(cm.exception.code, StatusCode.IO_FEATURE_NOT_WRITABLE)
            self.assertIn("kind=sparklines", cm.exception.context)
            self.assertIn("sheet_index=0", cm.exception.context)
            self.assertTrue(wb.save(format=WorkbookFormat.XLSX))

    def test_xlsb_save_of_a_protected_range_with_a_256_unit_name_fails_and_names_the_feature(self) -> None:
        with Workbook.load(FIXTURES / "protection" / "title_256_src.xlsx") as wb:
            self.assertEqual(len(wb.sheet(0).get_protected_range(0).name), 256)
            with self.assertRaises(IoError) as cm:
                wb.save(format=WorkbookFormat.XLSB)
            self.assertEqual(cm.exception.code, StatusCode.IO_FEATURE_NOT_WRITABLE)
            self.assertIn("kind=protected_range_title", cm.exception.context)
            self.assertTrue(wb.save(format=WorkbookFormat.XLSX))

    def test_sparkline_inputs_excel_cannot_store_are_refused(self) -> None:
        with self.assertRaises(InvalidArgumentError):
            self.ws.add_sparkline_group(SparklineGroupInput(sparklines=[SparklineInput("Sheet1!E1:E6", "A1:A2")]))
        with self.assertRaises(InvalidArgumentError) as cm:
            self.ws.add_sparkline_group(
                SparklineGroupInput(
                    date_axis_range="Sheet1!A1:A3,Sheet1!C1:C3", sparklines=[SparklineInput("Sheet1!E1:E6", "G1")]
                )
            )
        self.assertIn("date_axis_ref=", cm.exception.context)
        with self.assertRaises(InvalidArgumentError) as cm:
            self.ws.add_sparkline_group(SparklineGroupInput(sparklines=[SparklineInput("Sheet1!#REF!", "G1")]))
        self.assertIn("formula=", cm.exception.context)
        self.assertEqual(len(self.ws.list_sparkline_groups()), 0)

    def test_sparkline_update_with_a_union_source_is_refused_and_keeps_the_group(self) -> None:
        gid = self.ws.add_sparkline_group(SparklineGroupInput(sparklines=[SparklineInput("Sheet1!E1:E6", "G1")]))
        with self.assertRaises(InvalidArgumentError) as cm:
            self.ws.update_sparkline_group(gid, sparklines=[SparklineInput("Sheet1!A1:A3,Sheet1!C1:C3", "G1")])
        self.assertIn("formula=", cm.exception.context)
        self.assertIsNotNone(self.ws.get_sparkline_group(gid).sparklines[0].source)

    def test_protected_range_holds_8191_rectangles_and_not_8192_and_its_name_is_not_empty(self) -> None:
        self.assertEqual(self.ws.add_protected_range(ProtectedRangeInput(name="ok", ranges=_cells(8191))), "ok")
        self.assertTrue(self.wb.save(format=WorkbookFormat.XLSB))
        with self.assertRaises(InvalidArgumentError) as cm:
            self.ws.add_protected_range(ProtectedRangeInput(name="big", ranges=_cells(8192)))
        self.assertIn("part=ranges requested=8192 limit=8191", cm.exception.context)
        with self.assertRaises(InvalidArgumentError):
            self.ws.add_protected_range(ProtectedRangeInput(name="", ranges="A1:B2"))
        with self.assertRaises(InvalidArgumentError) as cm:
            self.ws.update_protected_range("ok", ranges=_cells(8192))
        self.assertIn("part=ranges requested=8192 limit=8191", cm.exception.context)
        self.assertEqual(len(self.ws.list_protected_ranges()), 1)

    def test_protected_range_name_past_255_units_is_refused_with_its_count_on_add_and_update(self) -> None:
        with self.assertRaises(InvalidArgumentError) as cm:
            self.ws.add_protected_range(ProtectedRangeInput(name="c" * 256, ranges="A1:B2"))
        self.assertIn("part=name requested=256 limit=255", cm.exception.context)
        self.ws.add_protected_range(ProtectedRangeInput(name="ok", ranges="A1:B2"))
        with self.assertRaises(InvalidArgumentError) as cm:
            self.ws.update_protected_range("ok", name="d" * 256)
        self.assertIn("part=name requested=256 limit=255", cm.exception.context)
        self.assertEqual(self.ws.get_protected_range("ok").name, "ok")

    def test_sparkline_with_a_ref_source_reads_back_without_one_and_the_read_back_value_is_accepted(self) -> None:
        gid = self.ws.add_sparkline_group(SparklineGroupInput(sparklines=[SparklineInput("#REF!", "G1")]))
        read = self.ws.get_sparkline_group(gid).sparklines[0]
        self.assertIsNone(read.source)
        self.ws.add_sparkline_group(SparklineGroupInput(sparklines=[read]))
        self.ws.update_sparkline_group(gid, sparklines=[read])
        self.assertEqual(len(self.ws.list_sparkline_groups()), 2)


if __name__ == "__main__":
    unittest.main()
