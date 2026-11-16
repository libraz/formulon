"""The plot argument of the chart methods is a position."""

from __future__ import annotations

import unittest

import _chart_input

from formulon import (
    AnchorKind,
    ChartPlotType,
    ChartSeriesInput,
    DrawingAnchorInput,
    IndexOutOfRangeError,
    Workbook,
)


def _chart_book() -> "tuple[Workbook, int]":
    wb = Workbook.create_default()
    ws = wb.sheet(0)
    for r in range(4):
        ws.set_number(r, 0, float(r))
        ws.set_number(r, 1, float(r * 2))
    chart = _chart_input.add_chart(
        ws, _chart_input.COLUMN, "A1:B4", anchor=DrawingAnchorInput(kind=AnchorKind.TWO_CELL, from_row=6)
    )
    return wb, chart.object_id


class ChartPlotPositionTests(unittest.TestCase):
    def test_negative_plot_is_index_out_of_range_on_every_plot_method(self) -> None:
        wb, oid = _chart_book()
        with wb:
            chart = wb.sheet(0).chart(oid)
            calls = {
                "set_type": lambda: chart.set_type(-1, ChartPlotType.LINE),
                "update_series": lambda: chart.update_series(-1, 0, name=None),
                "update_point": lambda: chart.update_point(-1, 0, 0, fill=None),
                "update_plot": lambda: chart.update_plot(-1, gap_width_pct=50),
                "add_series": lambda: chart.add_series(-1, ChartSeriesInput(values="Sheet1!$B$1:$B$4")),
                "remove_series": lambda: chart.remove_series(-1, 0),
            }
            for name, call in calls.items():
                with self.subTest(method=name), self.assertRaises(IndexOutOfRangeError):
                    call()


if __name__ == "__main__":
    unittest.main()
