"""Table slicer surface tests."""

from __future__ import annotations

import io
import unittest
import zipfile

from formulon import (
    AnchorKind,
    CellRect,
    FormulonError,
    NotFoundError,
    SheetRect,
    Slicer,
    SlicerCrossFilter,
    SlicerProperties,
    SlicerSortOrder,
    Workbook,
)

ANCHOR = {"from_row": 4, "from_col": 3, "to_row": 14, "to_col": 5}


def _parts(data: bytes) -> dict:
    archive = zipfile.ZipFile(io.BytesIO(data))
    return {name: archive.read(name).decode("utf-8") for name in archive.namelist()}


def _sales() -> Workbook:
    wb = Workbook.create_default()
    wb.sheet(0).set_text(0, 0, "Product")
    wb.sheet(0).set_text(0, 1, "Amount")
    wb.sheet(0).set_number(1, 1, 10.0)
    wb.add_table(
        range=SheetRect(0, CellRect(0, 0, 1, 1)), name="Sales", display_name="Sales", columns=["Product", "Amount"]
    )
    return wb


class SlicerTests(unittest.TestCase):
    def test_add_list_update_remove_and_round_trip(self) -> None:
        with _sales() as wb:
            self.assertEqual(wb.sheet(0).list_slicers(), [])
            slicer = wb.sheet(0).add_slicer(
                0,
                0,
                **ANCHOR,
                caption="Pick a product",
                style="SlicerStyleLight3",
                column_count=2,
                sort_order=SlicerSortOrder.DESCENDING,
                cross_filter=SlicerCrossFilter.NONE,
            )
            self.assertIsInstance(slicer, Slicer)
            object_id = slicer.object_id
            self.assertEqual(object_id, 2)
            parts = _parts(wb.save())
            self.assertIn('sortOrder="descending"', parts["xl/slicerCaches/slicerCache1.xml"])
            self.assertIn('caption="Pick a product"', parts["xl/slicers/slicer1.xml"])
            self.assertIn("application/vnd.ms-excel.slicerCache+xml", parts["[Content_Types].xml"])
            self.assertIn("<mc:Fallback>", parts["xl/drawings/drawing1.xml"])

            with Workbook.load(wb.save()) as loaded:
                slicers = loaded.sheet(0).list_slicers()
                self.assertEqual(len(slicers), 1)
                self.assertIs(slicers[0], loaded.sheet(0).slicer(object_id))
                s = slicers[0].get_properties()
                self.assertIsInstance(s, SlicerProperties)
                self.assertEqual((s.object_id, s.name, s.caption), (object_id, "Product", "Pick a product"))
                self.assertEqual((s.cache_name, s.table_name, s.column_name), ("Slicer_Product", "Sales", "Product"))
                self.assertEqual((s.style, s.column_count, s.row_height_emu), ("SlicerStyleLight3", 2, 241300))
                self.assertEqual((s.sort_order, s.cross_filter), (SlicerSortOrder.DESCENDING, SlicerCrossFilter.NONE))
                self.assertEqual(s.anchor_kind, AnchorKind.TWO_CELL)
                self.assertEqual((s.from_row, s.from_col, s.to_row, s.to_col), (4, 3, 14, 5))

                loaded.sheet(0).slicer(object_id).update(
                    caption="Renamed", style="", column_count=1, from_row=0, from_col=3, to_row=8, to_col=5
                )
                found = loaded.sheet(0).slicer(object_id).get_properties()
                self.assertIsNotNone(found)
                assert found is not None
                self.assertEqual(found.caption, "Renamed")
                loaded.sheet(0).insert_rows(0, 3)
                moved = loaded.sheet(0).list_slicers()[0].get_properties()
                self.assertEqual((moved.caption, moved.style, moved.column_count), ("Renamed", None, 1))
                self.assertEqual(moved.sort_order, SlicerSortOrder.DESCENDING)
                self.assertEqual((moved.from_row, moved.to_row), (3, 11))

                loaded.sheet(0).remove_drawing(object_id)
                self.assertEqual(loaded.sheet(0).list_slicers(), [])
                with self.assertRaises(NotFoundError):
                    loaded.sheet(0).slicer(object_id)
                emptied = _parts(loaded.save())
                self.assertNotIn("xl/slicers/slicer1.xml", emptied)
                self.assertNotIn("xl/slicerCaches/slicerCache1.xml", emptied)
                self.assertNotIn("Slicer_Product", emptied["xl/workbook.xml"])

    def test_removing_a_table_removes_its_slicers(self) -> None:
        with _sales() as wb:
            wb.sheet(0).add_slicer(0, 0, **ANCHOR)
            wb.sheet(0).add_slicer(0, 1, **ANCHOR)
            self.assertEqual(len(wb.sheet(0).list_slicers()), 2)
            wb.remove_table(0)
            self.assertEqual(wb.sheet(0).list_slicers(), [])
            parts = _parts(wb.save())
            for name in ("xl/slicers/slicer1.xml", "xl/slicerCaches/slicerCache1.xml"):
                self.assertNotIn(name, parts)
            self.assertNotIn("xl/slicerCaches/slicerCache2.xml", parts)
            self.assertNotIn("Slicer_", parts["xl/workbook.xml"])

    def test_arguments_are_validated(self) -> None:
        with _sales() as wb:
            with self.assertRaises(FormulonError):
                wb.sheet(0).add_slicer(1, 0, **ANCHOR)
            with self.assertRaises(FormulonError):
                wb.sheet(0).add_slicer(0, 2, **ANCHOR)
            with self.assertRaises(FormulonError):
                wb.sheet(0).add_slicer(0, 0, from_row=4, from_col=3, to_row=1, to_col=5)
            with self.assertRaises(FormulonError):
                wb.sheet(0).add_slicer(0, 0, **ANCHOR, column_count=0)
            with self.assertRaises(TypeError):
                wb.sheet(0).add_slicer(0, 0, from_row=1, from_col=2, to_row=3)  # type: ignore[call-arg]
            self.assertEqual(wb.sheet(0).list_slicers(), [])
            object_id = wb.sheet(0).add_slicer(0, 0, **ANCHOR).object_id
            with self.assertRaises(FormulonError):
                wb.sheet(0).slicer(object_id + 1).update(caption="x")
            with self.assertRaises(FormulonError):
                wb.sheet(0).remove_drawing(object_id + 1)


if __name__ == "__main__":
    unittest.main()
