"""Sheet.is_cell_stale: whether a formula cell holds a value its inputs have outdated, without recalculating."""

from __future__ import annotations

import unittest

from formulon import CalcMode, InvalidArgumentError, Workbook


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

    def test_a_manual_mode_save_computes_stale_formulas_first_unless_calc_on_save_is_off(self) -> None:
        self.wb.set_calc_mode(CalcMode.MANUAL)
        self.ws.set_value("A1", 1)
        self.ws.set_formula("B1", "=A1*2")
        self.wb.recalc()
        self.ws.set_value("A1", 5)
        self.assertTrue(self.wb.get_calc_on_save())
        self.wb.save()
        self.assertFalse(self.ws.is_cell_stale("B1"))
        self.assertEqual(self.ws.get_value("B1"), 10)

        self.ws.set_value("A1", 7)
        self.wb.set_calc_on_save(False)
        data = self.wb.save()
        self.assertTrue(self.ws.is_cell_stale("B1"))
        self.assertEqual(self.ws.get_value("B1"), 10)
        with Workbook.load(data) as loaded:
            self.assertTrue(loaded.sheet(0).is_cell_stale("B1"))
            self.assertEqual(loaded.sheet(0).get_value("B1"), 10)


class LoadProfileTests(unittest.TestCase):
    """A profile given at load leaves the stale set of the default load; set_profile_id afterwards dirties it all."""

    def setUp(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            wb.set_calc_mode(CalcMode.MANUAL)
            ws.set_value("A1", 1)
            ws.set_formula("B1", "=A1*2")
            ws.set_formula("C1", "=B1+1")
            ws.set_formula("D1", "=1+1")
            wb.recalc()
            ws.set_value("A1", 2)
            wb.set_calc_on_save(False)
            self.data = wb.save()

    @staticmethod
    def _stale(wb: Workbook) -> list:
        ws = wb.sheet(0)
        return [(ref, ws.is_cell_stale(ref), ws.get_value(ref)) for ref in ("A1", "B1", "C1", "D1")]

    def test_keeps_the_stale_set_and_the_cached_values_of_the_default_load(self) -> None:
        with Workbook.load(self.data) as plain, Workbook.load(self.data, profile_id="mac-365-ja_JP") as profiled:
            self.assertEqual(plain.get_profile_id(), "win-365-en_US")
            self.assertEqual(profiled.get_profile_id(), "mac-365-ja_JP")
            self.assertEqual(self._stale(profiled), self._stale(plain))
            self.assertTrue(profiled.sheet(0).is_cell_stale("B1"))
            self.assertFalse(profiled.sheet(0).is_cell_stale("D1"))

    def test_none_keeps_the_load_profile(self) -> None:
        with Workbook.load(self.data, profile_id=None) as wb:
            self.assertEqual(wb.get_profile_id(), "win-365-en_US")

    def test_refuses_an_unknown_or_empty_profile_id(self) -> None:
        with self.assertRaises(InvalidArgumentError):
            Workbook.load(self.data, profile_id="not-a-profile")
        with self.assertRaises(InvalidArgumentError):
            Workbook.load(self.data, profile_id="")

    def test_refuses_a_profile_id_that_is_not_a_string(self) -> None:
        for bad in (5, True, b"mac-365-ja_JP"):
            with self.assertRaises(TypeError) as caught:
                Workbook.load(self.data, profile_id=bad)  # type: ignore[arg-type]
            self.assertIn("Workbook.load", str(caught.exception))

    def test_the_profile_is_keyword_only(self) -> None:
        with self.assertRaises(TypeError):
            Workbook.load(self.data, "mac-365-ja_JP")  # type: ignore[misc]

    def test_set_profile_id_after_a_load_still_makes_every_formula_stale(self) -> None:
        with Workbook.load(self.data) as wb:
            self.assertFalse(wb.sheet(0).is_cell_stale("D1"))
            wb.set_profile_id("mac-365-ja_JP")
            for ref in ("B1", "C1", "D1"):
                self.assertTrue(wb.sheet(0).is_cell_stale(ref), ref)
