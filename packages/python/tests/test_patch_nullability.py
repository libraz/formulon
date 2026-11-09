"""``None`` in an ``update_*`` keyword clears only a field its patch record declares clearable; the static
ledger / d.ts / annotation comparison is the ``patch-nullability`` binding-drift check."""

from __future__ import annotations

import unittest
from typing import Any, Tuple

import formulon
from formulon import ChartPlotType, ProtectedRangeInput, Workbook


def _grid(ws: Any) -> None:
    for c, name in enumerate(("Name", "North", "South")):
        ws.set_text(0, c, name)
    for r in range(1, 4):
        ws.set_text(r, 0, f"row{r}")
        ws.set_number(r, 1, float(r))
        ws.set_number(r, 2, float(r * 2))


def _chart(wb: Workbook) -> Any:
    _grid(wb.sheet(0))
    return wb.sheet(0).add_chart(formulon.ChartInput(type=ChartPlotType.BAR, range="Sheet1!A1:C4"))


class NonClearableChartFieldsTest(unittest.TestCase):
    def _reject(self, call: Any, read: Any, names: Tuple[str, ...]) -> None:
        for name in names:
            with self.subTest(field=name):
                before = read()
                with self.assertRaises(TypeError):
                    call(**{name: None})
                self.assertEqual(read(), before)

    def test_update_plot_rejects_none(self) -> None:
        with Workbook.create_default() as wb:
            chart = _chart(wb)
            self._reject(
                lambda **kw: chart.update_plot(0, **kw),
                lambda: chart.get_properties().plots,
                ("gap_width_pct", "overlap_pct", "first_slice_angle_deg", "vary_colors", "hole_size_pct"),
            )

    def test_update_axis_rejects_none(self) -> None:
        with Workbook.create_default() as wb:
            chart = _chart(wb)
            axis_id = chart.get_properties().axes[0].id
            self._reject(
                lambda **kw: chart.update_axis(axis_id, **kw),
                lambda: chart.get_properties().axes,
                ("visible", "major_gridlines", "reversed", "crosses", "crosses_at"),
            )

    def test_update_series_rejects_none(self) -> None:
        with Workbook.create_default() as wb:
            chart = _chart(wb)
            series = chart.get_properties().plots[0].series[0]
            self._reject(
                lambda **kw: chart.update_series(0, series.series_id, **kw),
                lambda: chart.get_properties().plots,
                ("values", "y", "size", "smooth"),
            )


class ProtectedRangeClearTest(unittest.TestCase):
    def test_clearable_fields_accept_none_and_clear(self) -> None:
        fields = ("legacy_password_hash", "algorithm_name", "hash_value", "salt_value", "security_descriptor")
        values = {
            "legacy_password_hash": "CBEB",
            "algorithm_name": "SHA-512",
            "hash_value": "aGFzaA==",
            "salt_value": "c2FsdA==",
            "security_descriptor": "O:SYG:SYD:",
        }
        for field in fields:
            with self.subTest(field=field), Workbook.create_default() as wb:
                ws = wb.sheet(0)
                ws.add_protected_range(ProtectedRangeInput(name="P", ranges="A1:B2", **values))
                self.assertEqual(getattr(ws.get_protected_range("P"), field), values[field])
                ws.update_protected_range("P", **{field: None})
                read = ws.get_protected_range("P")
                self.assertIsNone(getattr(read, field))
                for other in fields:
                    if other != field:
                        self.assertEqual(getattr(read, other), values[other])


if __name__ == "__main__":
    unittest.main()
