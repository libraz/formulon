"""Sheet.is_cell_stale: whether a formula cell holds a value its inputs have outdated, without recalculating."""

from __future__ import annotations

import unittest

from formulon import CalcMode, Workbook


class IsCellStaleTests(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.addCleanup(self.wb.close)
        self.ws = self.wb.sheet(0)

    def test_flags_a_dependent_after_an_edit_under_manual_calculation_and_clears_on_recalc(self) -> None:
        self.wb.set_calc_mode(CalcMode.MANUAL)
        self.ws.set_value("A1", 1)
        self.ws.set_formula("B1", "=A1")
        self.wb.recalc()
        self.assertFalse(self.ws.is_cell_stale("B1"))
        self.ws.set_value("A1", 2)
        self.assertTrue(self.ws.is_cell_stale("B1"))
        self.assertFalse(self.ws.is_cell_stale("A1"))
        self.assertTrue(self.ws.is_cell_stale(0, 1))
        self.assertFalse(self.ws.is_cell_stale(0, 0))
        self.wb.recalc()
        self.assertFalse(self.ws.is_cell_stale("B1"))
        self.assertFalse(self.ws.is_cell_stale(0, 1))

    def test_flags_every_volatile_and_its_dependents_after_any_edit(self) -> None:
        self.wb.set_calc_mode(CalcMode.MANUAL)
        self.ws.set_formula("A1", "=RAND()")
        self.ws.set_formula("B1", "=A1+1")
        self.ws.set_formula("C1", "=1+1")
        self.wb.recalc()
        self.assertFalse(self.ws.is_cell_stale("A1"))
        self.ws.set_value("F9", 1)  # nothing reads it
        self.assertTrue(self.ws.is_cell_stale("A1"))
        self.assertTrue(self.ws.is_cell_stale("B1"))
        self.assertFalse(self.ws.is_cell_stale("C1"))
        self.wb.recalc()
        self.assertFalse(self.ws.is_cell_stale("A1"))
        self.assertFalse(self.ws.is_cell_stale("B1"))

    def test_false_under_automatic_calculation_after_an_edit(self) -> None:
        self.wb.set_calc_mode(CalcMode.AUTO)
        self.ws.set_value("A1", 1)
        self.ws.set_formula("B1", "=A1")
        self.wb.recalc()
        self.ws.set_value("A1", 2)
        self.assertFalse(self.ws.is_cell_stale("B1"))
        self.assertFalse(self.ws.is_cell_stale(0, 1))

    def test_false_for_an_empty_cell(self) -> None:
        self.wb.set_calc_mode(CalcMode.MANUAL)
        self.assertFalse(self.ws.is_cell_stale("D4"))
        self.assertFalse(self.ws.is_cell_stale(3, 3))
