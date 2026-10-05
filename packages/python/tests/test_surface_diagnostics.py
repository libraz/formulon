"""Package diagnostics surface tests."""

from __future__ import annotations

import struct
import unittest
from unittest import mock

import formulon
from formulon import (
    ConditionalFormatInput,
    DataValidationInput,
    MergeRange,
    ReadDiagnostics,
    SaveDiagnostics,
    Workbook,
    WorkbookFormat,
)


def _append_empty_zip_entry(data: bytes, name: str) -> bytes:
    """Add one deterministic, empty stored entry without recompressing XLSB."""
    eocd = data.rfind(b"PK\x05\x06")
    if eocd < 0:
        raise AssertionError("missing ZIP end record")
    _disk_count, count, central_size, central_offset = struct.unpack_from("<HHII", data, eocd + 8)
    encoded_name = name.encode("utf-8")
    local = struct.pack("<IHHHHHIIIHH", 0x04034B50, 20, 0, 0, 0, 0, 0, 0, 0, len(encoded_name), 0) + encoded_name
    central = (
        struct.pack(
            "<IHHHHHHIIIHHHHHII",
            0x02014B50,
            20,
            20,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            len(encoded_name),
            0,
            0,
            0,
            0,
            0,
            central_offset,
        )
        + encoded_name
    )
    new_eocd = struct.pack(
        "<IHHHHIIH",
        0x06054B50,
        0,
        0,
        count + 1,
        count + 1,
        central_size + len(central),
        central_offset + len(local),
        0,
    )
    return data[:central_offset] + local + data[central_offset : central_offset + central_size] + central + new_eocd


class PackageDiagnosticsTests(unittest.TestCase):
    def test_unknown_save_format_raises_formulon_error(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(formulon.FormulonError) as ctx:
                wb.save_with_diagnostics(WorkbookFormat.UNKNOWN)
            self.assertIn("unsupported format", str(ctx.exception))

    def test_save_and_read_diagnostics_report_every_counter(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_formula(0, 0, 0, "=SUM(T[C])")
            wb.add_validation(
                0,
                DataValidationInput(
                    type=3,
                    ranges=[MergeRange(0, 0, 0, 0)],
                    formula1='"Yes,No"',
                ),
            )
            wb.add_conditional_format(
                0,
                ConditionalFormatInput(
                    sqref=[MergeRange(0, 0, 0, 0)],
                    type=1,
                    op_engaged=True,
                    op=5,
                    formula1="10",
                ),
            )

            xlsx = wb.save_with_diagnostics(WorkbookFormat.XLSX)
            self.assertIsInstance(xlsx, SaveDiagnostics)
            self.assertIsInstance(xlsx.bytes, bytes)
            self.assertEqual(xlsx.downgraded_formula_count, 0)
            self.assertEqual(xlsx.deferred_feature_count, 0)
            self.assertEqual(xlsx.dropped_part_count, 0)
            self.assertEqual(xlsx.dropped_relationship_count, 0)
            self.assertEqual(xlsx.renumbered_part_count, 0)

            xlsb = wb.save_with_diagnostics(WorkbookFormat.XLSB)
            self.assertEqual(xlsb.downgraded_formula_count, 1)
            # The validation and the CF rule (no dxf) are both written.
            self.assertEqual(xlsb.deferred_feature_count, 0)
            # The binary writer never reassigns a part id.
            self.assertEqual(xlsb.renumbered_part_count, 0)

            # `save_as` remains a plain bytes return value.
            self.assertIsInstance(wb.save_as(WorkbookFormat.XLSB), bytes)

            with Workbook.load(xlsb.bytes) as loaded:
                read = loaded.read_diagnostics()
                self.assertIsInstance(read, ReadDiagnostics)
                self.assertEqual(read.undecoded_formula_count, 0)
                self.assertEqual(read.undecoded_defined_name_count, 0)
                self.assertEqual(read.undecoded_part_count, 0)
                self.assertEqual(read.skipped_feature_count, 0)
                self.assertEqual(read.unknown_content_type_count, 0)

            with Workbook.load(_append_empty_zip_entry(xlsb.bytes, "xl/preserved.bin")) as preserved_loaded:
                preserved = preserved_loaded.read_diagnostics()
                self.assertEqual(preserved.undecoded_part_count, 0)

    def test_diagnostic_scratch_is_freed_when_a_later_allocation_fails(self) -> None:
        # `save_with_diagnostics` takes three blocks (two out-pointers plus the
        # counter struct). If the second allocation raises, the first must
        # still be released rather than leaked into the WASM heap.
        with Workbook.create_default() as wb:
            original_alloc = formulon.workbook.LIB.alloc
            calls = 0

            def alloc_then_fail(size: int) -> int:
                nonlocal calls
                calls += 1
                if calls == 2:
                    raise RuntimeError("synthetic allocation failure")
                return original_alloc(size)

            with (
                mock.patch.object(formulon.workbook.LIB, "alloc", side_effect=alloc_then_fail),
                mock.patch.object(formulon.workbook.LIB, "free", wraps=formulon.workbook.LIB.free) as free,
            ):
                with self.assertRaises(RuntimeError):
                    wb.save_with_diagnostics(WorkbookFormat.XLSB)
            self.assertEqual(free.call_count, 1)


if __name__ == "__main__":
    unittest.main()
