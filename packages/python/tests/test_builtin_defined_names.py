"""A defined name spelled as the profile locale's built-in name is stored as that built-in."""

from __future__ import annotations

import unittest

from formulon import CellRect, DefinedName, DefinedNameInput, InvalidArgumentError, Workbook


class BuiltinDefinedNameTest(unittest.TestCase):
    def test_print_area_defines_the_print_area_of_the_sheet_its_formula_names(self) -> None:
        wb = Workbook.create_for_profile("mac-365-ja_JP")
        self.addCleanup(wb.close)
        wb.add_sheet("Sheet2")
        wb.set_defined_name(
            DefinedNameInput(name="print_area", formula="=Sheet2!$A$1:$B$2", comment="note", hidden=True)
        )
        self.assertEqual(
            wb.list_defined_names(),
            [DefinedName(name="_xlnm.Print_Area", formula="=Sheet2!$A$1:$B$2", scope=1, comment="note", hidden=True)],
        )
        self.assertEqual(wb.sheet(1).get_print_area(), [CellRect(0, 0, 1, 1)])
        self.assertIsNone(wb.get_defined_name("Print_Area"))
        found = wb.get_defined_name("Print_Area", 1)
        self.assertIsNotNone(found)
        assert found is not None
        self.assertEqual(found.name, "_xlnm.Print_Area")
        wb.sheet(1).set_print_area("A1:C3")
        self.assertEqual(len(wb.list_defined_names()), 1)
        wb.remove_defined_name("PRINT_AREA", 1)
        self.assertEqual(wb.list_defined_names(), [])

    def test_only_the_locale_spelling_is_a_built_in(self) -> None:
        wb = Workbook.create_for_profile("mac-365-de_DE")
        self.addCleanup(wb.close)
        sheet = wb.sheet(0).name
        wb.set_defined_name(DefinedNameInput(name="Print_Area", formula=f"={sheet}!$A$1:$B$2"))
        wb.set_defined_name(DefinedNameInput(name="Druckbereich", formula=f"={sheet}!$A$1:$B$2"))
        self.assertEqual(
            [(n.name, n.scope) for n in wb.list_defined_names()],
            [("Print_Area", None), ("_xlnm.Print_Area", 0)],
        )
        with self.assertRaises(InvalidArgumentError):
            wb.set_defined_name(DefinedNameInput(name="_xlnm.Print_Titles", formula=f"={sheet}!$1:$1"))


if __name__ == "__main__":
    unittest.main()
