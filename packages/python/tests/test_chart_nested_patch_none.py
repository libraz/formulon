"""None in a nested chart patch clears only the keys that patch declares clearable."""

from __future__ import annotations

import unittest

import _chart_input

from formulon import (
    AnchorKind,
    DrawingAnchorInput,
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


class NestedPatchNoneTests(unittest.TestCase):
    def test_none_on_a_key_the_nested_patch_cannot_clear_is_a_type_error(self) -> None:
        wb, oid = _chart_book()
        with wb:
            chart = wb.sheet(0).chart(oid)
            for patch in (
                {"labels": {"show_value": None}},
                {"marker": {"symbol": None}},
                {"fill": {"none": None}},
                {"line": {"none": None}},
            ):
                with self.subTest(patch=patch), self.assertRaises(TypeError):
                    chart.update_series(0, 0, **patch)

    def test_nested_patch_clears_a_clearable_key(self) -> None:
        wb, oid = _chart_book()
        with wb:
            chart = wb.sheet(0).chart(oid)
            chart.update_series(0, 0, labels={"separator": None})
            chart.update_series(0, 0, fill={"transparency_pct": None}, line={"width_pt": None})

    def test_paint_patch_accepts_its_read_only_other_key(self) -> None:
        wb, oid = _chart_book()
        with wb:
            chart = wb.sheet(0).chart(oid)
            chart.update_series(0, 0, fill={"color": "#FFFF0000", "other": False})
            chart.update_series(0, 0, line={"width_pt": 2.0, "other": False})


if __name__ == "__main__":
    unittest.main()
