"""Sheet.list_circular_references: the unresolved cycle members on a sheet, row-major."""

from __future__ import annotations

import unittest

from formulon import CellAddress, Workbook


class ListCircularReferencesTests(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.addCleanup(self.wb.close)
        self.ws = self.wb.sheet(0)

    def test_empty_without_a_cycle(self) -> None:
        self.assertEqual(self.ws.list_circular_references(), [])

    def test_names_cycle_members_and_empties_once_broken(self) -> None:
        self.ws.set_formula("A1", "=B1")
        self.ws.set_formula("B1", "=A1")
        self.wb.recalc()
        self.assertEqual(self.ws.list_circular_references(), [CellAddress(row=0, col=0), CellAddress(row=0, col=1)])
        self.ws.set_value("B1", 1)
        self.wb.recalc()
        self.assertEqual(self.ws.list_circular_references(), [])

    def test_orders_members_row_major_and_recalculates_first(self) -> None:
        self.ws.set_formula("C2", "=A1")
        self.ws.set_formula("A1", "=B3")
        self.ws.set_formula("B3", "=C2")
        self.ws.set_formula("D1", "=A1+1")
        self.assertEqual(
            self.ws.list_circular_references(),
            [CellAddress(row=0, col=0), CellAddress(row=1, col=2), CellAddress(row=2, col=1)],
        )

    def test_lists_nothing_under_iterative_calculation(self) -> None:
        self.ws.set_formula("A1", "=B1+1")
        self.ws.set_formula("B1", "=A1/2")
        self.wb.recalc()
        self.assertEqual(len(self.ws.list_circular_references()), 2)
        self.wb.update_iterative(enabled=True)
        self.wb.recalc()
        self.assertEqual(self.ws.list_circular_references(), [])


if __name__ == "__main__":
    unittest.main()
