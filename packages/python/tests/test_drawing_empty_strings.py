"""Empty strings in the DRAWING / SHAPE / SLICER / CHART / SPARKLINE setters (api-conventions §5 rule 30).

Where the file holds one state for ``""`` and the absent value they mean the same; where ``""`` is not a legal value it
is refused and the absent value clears or derives.
"""

from __future__ import annotations

import unittest

from formulon import (
    CellPictureInput,
    CellRect,
    ChartInput,
    ChartPlotType,
    ChartSourceKind,
    ChartTitle,
    ChartTitleKind,
    DrawingAnchorInput,
    DrawingPlacementInput,
    FormulaSyntaxError,
    InvalidArgumentError,
    ShapeInput,
    SlicerInput,
    SparklineGroupInput,
    SparklineInput,
    TableInput,
    Workbook,
)

_BOX = DrawingAnchorInput(from_row=1, from_col=1, to_row=5, to_col=4)
_SIDE_BOX = DrawingAnchorInput(from_row=1, from_col=6, to_row=5, to_col=9)


def png(width: int, height: int) -> bytes:
    return (
        b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR"
        + width.to_bytes(4, "big")
        + height.to_bytes(4, "big")
        + b"\x08\x06\x00\x00\x00\x1f\x15\xc4\x89"
    )


class DrawingEmptyStringTests(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.addCleanup(self.wb.close)
        self.ws = self.wb.sheet(0)

    def _fill_grid(self) -> None:
        for c in range(1, 4):
            self.ws.set_text(0, c, "S")
            for r in range(1, 5):
                self.ws.set_number(r, c, float(r * (c + 1)))
        for r in range(1, 5):
            self.ws.set_text(r, 0, "a")

    def _sales_table(self):
        self.ws.set_text(8, 0, "Product")
        self.ws.set_text(9, 0, "Widget")
        return self.ws.add_table(TableInput(range=CellRect(8, 0, 9, 0), name="Sales", columns=["Product"]))

    def test_picture_description_empty_and_absent_are_one_state(self) -> None:
        empty = self.ws.add_picture(png(4, 4), DrawingPlacementInput(description=""))
        absent = self.ws.add_picture(png(4, 4))
        self.assertFalse(empty.get_drawing_properties().description)
        self.assertFalse(absent.get_drawing_properties().description)

    def test_placement_name_empty_is_refused_on_every_kind_and_absent_derives_one(self) -> None:
        self._fill_grid()
        table = self._sales_table()
        chart = {"type": ChartPlotType.LINE, "range": "Sheet1!$A$1:$D$5", "anchor": _BOX}
        with self.assertRaises(InvalidArgumentError):
            self.ws.add_picture(png(4, 4), DrawingPlacementInput(name=""))
        with self.assertRaises(InvalidArgumentError):
            self.ws.add_shape(ShapeInput(name="", anchor=_BOX))
        with self.assertRaises(InvalidArgumentError):
            self.ws.add_chart(ChartInput(name="", **chart))
        with self.assertRaises(InvalidArgumentError):
            self.ws.add_slicer(table, 0, SlicerInput(name="", anchor=_SIDE_BOX))
        self.assertEqual(self.ws.list_drawings(), [])

        made = [
            self.ws.add_picture(png(4, 4)),
            self.ws.add_shape(ShapeInput(anchor=_BOX)),
            self.ws.add_chart(ChartInput(**chart)),
            self.ws.add_slicer(table, 0, SlicerInput(anchor=_SIDE_BOX)),
        ]
        for drawing in made:
            self.assertNotEqual(drawing.get_drawing_properties().name, "")

    def test_drawing_patch_description_empty_removes_it_as_none_does(self) -> None:
        shape = self.ws.add_shape(ShapeInput(anchor=_BOX, description="alt"))
        self.assertEqual(shape.get_drawing_properties().description, "alt")
        shape.update(description="")
        self.assertFalse(shape.get_drawing_properties().description)
        shape.update(description="again")
        shape.update(description=None)
        self.assertFalse(shape.get_drawing_properties().description)

    def test_cell_picture_alt_text_empty_means_none(self) -> None:
        self.ws.set_cell_picture("A1", CellPictureInput(png(2, 2), "old"))
        self.ws.set_cell_picture("A1", CellPictureInput(png(2, 2), ""))
        picture = self.ws.get_cell_picture("A1")
        assert picture is not None
        self.assertIsNone(picture.alt_text)

    def test_shape_preset_geometry_empty_is_refused_and_absent_is_rect(self) -> None:
        with self.assertRaises(InvalidArgumentError):
            self.ws.add_shape(ShapeInput(anchor=_BOX, preset_geometry=""))
        self.assertEqual(self.ws.list_drawings(), [])
        self.assertEqual(self.ws.add_shape(ShapeInput(anchor=_BOX)).get_properties().preset_geometry, "rect")

    def test_shape_text_empty_and_absent_are_one_state(self) -> None:
        empty = self.ws.add_shape(ShapeInput(anchor=_BOX, text=""))
        absent = self.ws.add_shape(ShapeInput(anchor=_BOX))
        self.assertEqual(empty.get_properties().text, absent.get_properties().text)

    def test_slicer_caption_empty_and_absent_both_take_the_column_name(self) -> None:
        table = self._sales_table()
        empty = self.ws.add_slicer(table, 0, SlicerInput(anchor=_BOX, caption=""))
        absent = self.ws.add_slicer(table, 0, SlicerInput(anchor=_SIDE_BOX))
        self.assertEqual(empty.get_properties().caption, "Product")
        self.assertEqual(absent.get_properties().caption, "Product")

    def test_sparkline_source_empty_is_refused_and_none_is_a_deleted_source(self) -> None:
        with self.assertRaises(InvalidArgumentError):
            self.ws.add_sparkline_group(SparklineGroupInput(sparklines=[SparklineInput(source="", location="F1")]))
        self.assertEqual(self.ws.list_sparkline_groups(), [])
        group_id = self.ws.add_sparkline_group(
            SparklineGroupInput(sparklines=[SparklineInput(source=None, location="F1")])
        )
        self.assertIsNone(self.ws.list_sparkline_groups()[0].sparklines[0].source)
        with self.assertRaises(InvalidArgumentError):
            self.ws.update_sparkline_group(group_id, sparklines=[SparklineInput(source="", location="F2")])

    def test_sparkline_read_back_with_a_deleted_source_writes_back_as_none(self) -> None:
        group_id = self.ws.add_sparkline_group(
            SparklineGroupInput(sparklines=[SparklineInput(source=None, location="F1")])
        )
        group = self.ws.list_sparkline_groups()[0]
        self.ws.update_sparkline_group(group_id, sparklines=group.sparklines)
        self.assertIsNone(self.ws.list_sparkline_groups()[0].sparklines[0].source)

    def test_sparkline_date_axis_range_empty_is_refused_and_none_is_no_date_axis(self) -> None:
        sparklines = [SparklineInput(source="Sheet1!A1:D1", location="F1")]
        with self.assertRaises(InvalidArgumentError):
            self.ws.add_sparkline_group(
                SparklineGroupInput(sparklines=sparklines, date_axis_range="")  # type: ignore[arg-type]
            )
        self.assertEqual(self.ws.list_sparkline_groups(), [])
        group_id = self.ws.add_sparkline_group(SparklineGroupInput(sparklines=sparklines, date_axis_range=None))
        self.assertIsNone(self.ws.list_sparkline_groups()[0].date_axis_range)
        with self.assertRaises(InvalidArgumentError):
            self.ws.update_sparkline_group(group_id, date_axis_range="")  # type: ignore[arg-type]

    def test_sparkline_location_empty_is_a_parse_failure(self) -> None:
        with self.assertRaises(FormulaSyntaxError):
            self.ws.add_sparkline_group(
                SparklineGroupInput(sparklines=[SparklineInput(source="Sheet1!A1:D1", location="")])
            )
        self.assertEqual(self.ws.list_sparkline_groups(), [])

    def test_chart_series_empty_reference_is_refused_and_none_removes(self) -> None:
        self._fill_grid()
        chart = self.ws.add_chart(ChartInput(type=ChartPlotType.LINE, range="Sheet1!$A$1:$D$5", anchor=_BOX))
        for key in ("name", "categories", "x", "values"):
            with self.subTest(key=key), self.assertRaises(InvalidArgumentError):
                chart.update_series(0, 0, **{key: ""})
        chart.update_series(0, 0, name="Sheet1!$B$1")
        self.assertEqual(chart.get_properties().plots[0].series[0].name.kind, ChartSourceKind.REF)
        chart.update_series(0, 0, name=None)
        self.assertEqual(chart.get_properties().plots[0].series[0].name.kind, ChartSourceKind.NONE)

    def test_chart_title_text_empty_and_missing_are_one_state(self) -> None:
        self._fill_grid()
        for title in (ChartTitle(kind=ChartTitleKind.TEXT, text=""), ChartTitle(kind=ChartTitleKind.TEXT, text=None)):
            with self.assertRaises(InvalidArgumentError):
                self.ws.add_chart(
                    ChartInput(type=ChartPlotType.LINE, range="Sheet1!$A$1:$D$5", anchor=_BOX, title=title)
                )
        self.assertEqual(self.ws.list_drawings(), [])


if __name__ == "__main__":
    unittest.main()
