"""Table slicer surface tests."""

from __future__ import annotations

import io
import unittest
import zipfile

from formulon import (
    AnchorKind,
    FormulonError,
    Slicer,
    SlicerCrossFilter,
    SlicerSortOrder,
    Workbook,
)

ANCHOR = (4, 3, 14, 5)


def _parts(data: bytes) -> dict:
    archive = zipfile.ZipFile(io.BytesIO(data))
    return {name: archive.read(name).decode("utf-8") for name in archive.namelist()}


def _sales() -> Workbook:
    wb = Workbook.create_default()
    wb.set_text(0, 0, 0, "Product")
    wb.set_text(0, 0, 1, "Amount")
    wb.set_number(0, 1, 1, 10.0)
    wb.table_create(0, "A1:B2", "Sales", "Sales", ["Product", "Amount"])
    return wb


class SlicerTests(unittest.TestCase):
    def test_add_list_update_remove_and_round_trip(self) -> None:
        with _sales() as wb:
            self.assertEqual(wb.list_slicers(0), [])
            object_id = wb.add_slicer(
                0,
                0,
                0,
                anchor=ANCHOR,
                caption="Pick a product",
                style="SlicerStyleLight3",
                column_count=2,
                sort_order=SlicerSortOrder.DESCENDING,
                cross_filter=SlicerCrossFilter.NONE,
            )
            self.assertEqual(object_id, 2)
            parts = _parts(wb.save())
            self.assertIn('sortOrder="descending"', parts["xl/slicerCaches/slicerCache1.xml"])
            self.assertIn('caption="Pick a product"', parts["xl/slicers/slicer1.xml"])
            self.assertIn("application/vnd.ms-excel.slicerCache+xml", parts["[Content_Types].xml"])
            self.assertIn("<mc:Fallback>", parts["xl/drawings/drawing1.xml"])

            with Workbook.load(wb.save()) as loaded:
                slicers = loaded.list_slicers(0)
                self.assertEqual(len(slicers), 1)
                s = slicers[0]
                self.assertIsInstance(s, Slicer)
                self.assertEqual((s.object_id, s.name, s.caption), (object_id, "Product", "Pick a product"))
                self.assertEqual((s.cache_name, s.table_name, s.column_name), ("Slicer_Product", "Sales", "Product"))
                self.assertEqual((s.style, s.column_count, s.row_height), ("SlicerStyleLight3", 2, 241300))
                self.assertEqual((s.sort_order, s.cross_filter), (SlicerSortOrder.DESCENDING, SlicerCrossFilter.NONE))
                self.assertEqual(s.anchor_kind, AnchorKind.TWO_CELL)
                self.assertEqual((s.from_row, s.from_col, s.to_row, s.to_col), ANCHOR)

                loaded.update_slicer(0, 0, caption="Renamed", style="", column_count=1, anchor=(0, 3, 8, 5))
                loaded.insert_rows(0, 0, 3)
                moved = loaded.list_slicers(0)[0]
                self.assertEqual((moved.caption, moved.style, moved.column_count), ("Renamed", "", 1))
                self.assertEqual(moved.sort_order, SlicerSortOrder.DESCENDING)
                self.assertEqual((moved.from_row, moved.to_row), (3, 11))

                loaded.remove_slicer(0, 0)
                self.assertEqual(loaded.list_slicers(0), [])
                emptied = _parts(loaded.save())
                self.assertNotIn("xl/slicers/slicer1.xml", emptied)
                self.assertNotIn("xl/slicerCaches/slicerCache1.xml", emptied)
                self.assertNotIn("Slicer_Product", emptied["xl/workbook.xml"])

    def test_removing_a_table_removes_its_slicers(self) -> None:
        with _sales() as wb:
            wb.add_slicer(0, 0, 0, anchor=ANCHOR)
            wb.add_slicer(0, 0, 1, anchor=ANCHOR)
            self.assertEqual(len(wb.list_slicers(0)), 2)
            wb.table_remove(0)
            self.assertEqual(wb.list_slicers(0), [])
            parts = _parts(wb.save())
            for name in ("xl/slicers/slicer1.xml", "xl/slicerCaches/slicerCache1.xml"):
                self.assertNotIn(name, parts)
            self.assertNotIn("xl/slicerCaches/slicerCache2.xml", parts)
            self.assertNotIn("Slicer_", parts["xl/workbook.xml"])

    def test_arguments_are_validated(self) -> None:
        with _sales() as wb:
            with self.assertRaises(FormulonError):
                wb.add_slicer(0, 1, 0, anchor=ANCHOR)
            with self.assertRaises(FormulonError):
                wb.add_slicer(0, 0, 2, anchor=ANCHOR)
            with self.assertRaises(FormulonError):
                wb.add_slicer(0, 0, 0, anchor=(4, 3, 1, 5))
            with self.assertRaises(FormulonError):
                wb.add_slicer(0, 0, 0, anchor=ANCHOR, column_count=0)
            with self.assertRaises(TypeError):
                wb.add_slicer(0, 0, 0, anchor=(1, 2, 3))
            self.assertEqual(wb.list_slicers(0), [])
            wb.add_slicer(0, 0, 0, anchor=ANCHOR)
            with self.assertRaises(FormulonError):
                wb.update_slicer(0, 3, caption="x")
            with self.assertRaises(FormulonError):
                wb.remove_slicer(0, 3)


if __name__ == "__main__":
    unittest.main()
