"""Formulas entered through the binding take Excel's spelling for names and references."""

from __future__ import annotations

import unittest

from formulon import Workbook


class EnteredFormulaSpellingTest(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.addCleanup(self.wb.close)
        self.ws = self.wb.sheet(0)

    def test_builtins_and_references_are_upper_cased(self) -> None:
        self.ws.set_formula("A1", "=sum(b1:c2)+pi()")
        self.assertEqual(self.ws.get_formula("A1"), "=SUM(B1:C2)+PI()")

    def test_unknown_function_takes_the_first_entered_spelling(self) -> None:
        self.ws.set_formula("A1", "=IFERROR(fOo(1),0)")
        self.ws.set_formula("A2", "=IFERROR(FOO(2),0)")
        self.assertEqual(self.ws.get_formula("A1"), "=IFERROR(fOo(1),0)")
        self.assertEqual(self.ws.get_formula("A2"), "=IFERROR(fOo(2),0)")


if __name__ == "__main__":
    unittest.main()
