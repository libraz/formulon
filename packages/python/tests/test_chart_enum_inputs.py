"""Enum fields of a chart input take integers only."""

from __future__ import annotations

import unittest

from formulon import (
    ChartInput,
    InvalidArgumentError,
    Workbook,
)


class EnumInputTests(unittest.TestCase):
    def test_chart_plot_type_refuses_a_bool_and_a_float(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            for value in (True, 1.0):
                with self.subTest(value=value), self.assertRaises(TypeError):
                    ws.add_chart(ChartInput(type=value, range="A1:B3"))  # type: ignore[arg-type]
            self.assertEqual(len(ws.list_drawings()), 0)

    def test_chart_plot_type_unknown_member_is_a_status_error(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(InvalidArgumentError):
                wb.sheet(0).add_chart(ChartInput(type=999, range="A1:B3"))  # type: ignore[arg-type]


if __name__ == "__main__":
    unittest.main()
