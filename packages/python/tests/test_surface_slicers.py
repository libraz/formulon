"""Table slicer surface tests."""

from __future__ import annotations

import io
import unittest
import zipfile

from formulon import (
    AnchorEditAs,
    AnchorKind,
    CellRect,
    DrawingAnchorInput,
    FormulonError,
    NotFoundError,
    Slicer,
    SlicerCrossFilter,
    SlicerInput,
    SlicerProperties,
    SlicerSortOrder,
    TableInput,
    Workbook,
)

ANCHOR = {"anchor": DrawingAnchorInput(from_row=4, from_col=3, to_row=14, to_col=5)}
SALES = "Sales"


def _parts(data: bytes) -> dict:
    archive = zipfile.ZipFile(io.BytesIO(data))
    return {name: archive.read(name).decode("utf-8") for name in archive.namelist()}


def _sales() -> Workbook:
    wb = Workbook.create_default()
    wb.sheet(0).set_text(0, 0, "Product")
    wb.sheet(0).set_text(0, 1, "Amount")
    wb.sheet(0).set_number(1, 1, 10.0)
    wb.sheet(0).add_table(TableInput(range=CellRect(0, 0, 1, 1), name=SALES, columns=["Product", "Amount"]))
    return wb


class SlicerTests(unittest.TestCase):
    def test_add_list_update_remove_and_round_trip(self) -> None:
        with _sales() as wb:
            self.assertEqual(wb.sheet(0).list_slicers(), [])
            slicer = wb.sheet(0).add_slicer(
                SALES,
                0,
                SlicerInput(
                    **ANCHOR,
                    caption="Pick a product",
                    style_name="SlicerStyleLight3",
                    column_count=2,
                    sort_order=SlicerSortOrder.DESCENDING,
                    cross_filter=SlicerCrossFilter.NONE,
                ),
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
                self.assertEqual((s.object_id, s.caption), (object_id, "Pick a product"))
                found_common = slicers[0].get_drawing_properties()
                self.assertEqual(found_common.name, "Product")
                self.assertEqual((s.cache_name, s.table_name, s.column_name), ("Slicer_Product", "Sales", "Product"))
                self.assertEqual((s.style_name, s.column_count, s.row_height_pt), ("SlicerStyleLight3", 2, 19.0))
                self.assertEqual((s.sort_order, s.cross_filter), (SlicerSortOrder.DESCENDING, SlicerCrossFilter.NONE))
                anchor = found_common.anchor
                self.assertEqual(anchor.kind, AnchorKind.TWO_CELL)
                self.assertEqual((anchor.from_row, anchor.from_col, anchor.to_row, anchor.to_col), (4, 3, 14, 5))
                self.assertGreater(anchor.width_emu, 0)
                self.assertGreater(anchor.height_emu, 0)
                self.assertNotEqual(anchor.edit_as, AnchorEditAs.DEFAULT)

                loaded.sheet(0).slicer(object_id).update(
                    caption="Renamed",
                    style_name=None,
                    column_count=1,
                    anchor=DrawingAnchorInput(from_row=0, from_col=3, to_row=8, to_col=5),
                )
                found = loaded.sheet(0).slicer(object_id).get_properties()
                self.assertIsNotNone(found)
                assert found is not None
                self.assertEqual(found.caption, "Renamed")
                loaded.sheet(0).insert_rows(0, 2)
                moved_handle = loaded.sheet(0).list_slicers()[0]
                moved = moved_handle.get_properties()
                self.assertEqual((moved.caption, moved.style_name, moved.column_count), ("Renamed", None, 1))
                self.assertEqual(moved.sort_order, SlicerSortOrder.DESCENDING)
                moved_anchor = moved_handle.get_drawing_properties().anchor
                self.assertEqual((moved_anchor.from_row, moved_anchor.to_row), (3, 11))

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
            wb.sheet(0).add_slicer(SALES, 0, SlicerInput(**ANCHOR))
            wb.sheet(0).add_slicer(SALES, 1, SlicerInput(**ANCHOR))
            self.assertEqual(len(wb.sheet(0).list_slicers()), 2)
            wb.sheet(0).remove_table(0)
            self.assertEqual(wb.sheet(0).list_slicers(), [])
            parts = _parts(wb.save())
            for name in ("xl/slicers/slicer1.xml", "xl/slicerCaches/slicerCache1.xml"):
                self.assertNotIn(name, parts)
            self.assertNotIn("xl/slicerCaches/slicerCache2.xml", parts)
            self.assertNotIn("Slicer_", parts["xl/workbook.xml"])

    def test_column_by_name_equals_column_by_position(self) -> None:
        with _sales() as wb:
            by_position = wb.sheet(0).add_slicer(SALES, 1, SlicerInput(**ANCHOR))
            by_name = wb.sheet(0).add_slicer(SALES, "Amount", SlicerInput(**ANCHOR))
            by_folded_name = wb.sheet(0).add_slicer(SALES, "amount", SlicerInput(**ANCHOR))
            for handle in (by_position, by_name, by_folded_name):
                self.assertEqual(handle.get_properties().column_name, "Amount")
            by_handle = wb.sheet(0).add_slicer(wb.table(SALES), "Product", SlicerInput(**ANCHOR))
            self.assertEqual(by_handle.get_properties().column_name, "Product")
            with self.assertRaises(NotFoundError):
                wb.sheet(0).add_slicer(SALES, "nope", SlicerInput(**ANCHOR))
            with self.assertRaises(TypeError):
                wb.sheet(0).add_slicer(SALES, 1.5, SlicerInput(**ANCHOR))  # type: ignore[arg-type]

    def test_common_update_keys_apply_to_a_slicer_and_a_name_is_not_a_keyword(self) -> None:
        with _sales() as wb:
            slicer = wb.sheet(0).add_slicer(SALES, 0, SlicerInput(**ANCHOR))
            slicer.update(description="pick one", hidden=True, caption="Cap", anchor=DrawingAnchorInput(to_col=6))
            common = slicer.get_drawing_properties()
            self.assertEqual((common.description, common.hidden), ("pick one", True))
            self.assertEqual(
                (common.anchor.kind, common.anchor.from_row, common.anchor.to_col), (AnchorKind.TWO_CELL, 4, 6)
            )
            self.assertEqual(slicer.get_properties().caption, "Cap")
            self.assertEqual(slicer.name, "Product")
            with self.assertRaises(TypeError):
                slicer.update(name="other")  # type: ignore[call-arg]
            self.assertEqual(slicer.name, "Product")
            self.assertEqual(slicer.get_drawing_properties().description, "pick one")

    def test_a_slicer_is_removed_by_the_one_remover(self) -> None:
        with _sales() as wb:
            slicer = wb.sheet(0).add_slicer(SALES, 0, SlicerInput(**ANCHOR))
            wb.sheet(0).remove_drawing(slicer)
            self.assertEqual(wb.sheet(0).list_slicers(), [])
            self.assertEqual(wb.sheet(0).list_drawings(), [])
            with self.assertRaises(NotFoundError):
                wb.sheet(0).remove_drawing(slicer)

    def test_arguments_are_validated(self) -> None:
        with _sales() as wb:
            with self.assertRaises(FormulonError):
                wb.sheet(0).add_slicer("Missing", 0, SlicerInput(**ANCHOR))
            with self.assertRaises(FormulonError):
                wb.sheet(0).add_slicer(SALES, 2, SlicerInput(**ANCHOR))
            with self.assertRaises(FormulonError):
                wb.sheet(0).add_slicer(
                    SALES, 0, SlicerInput(anchor=DrawingAnchorInput(from_row=4, from_col=3, to_row=1, to_col=5))
                )
            with self.assertRaises(TypeError):
                SlicerInput(from_row=1, from_col=2, to_row=3)  # type: ignore[call-arg]
            with self.assertRaises(TypeError):
                DrawingAnchorInput(row=1)  # type: ignore[call-arg]
            self.assertEqual(wb.sheet(0).list_slicers(), [])
            slicer = wb.sheet(0).add_slicer(SALES, 0, SlicerInput(**ANCHOR))
            object_id = slicer.object_id
            with self.assertRaises(FormulonError):
                slicer.update(column_count=0)
            with self.assertRaises(FormulonError):
                wb.sheet(0).slicer(object_id + 1).update(caption="x")
            with self.assertRaises(FormulonError):
                wb.sheet(0).remove_drawing(object_id + 1)


if __name__ == "__main__":
    unittest.main()
