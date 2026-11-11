"""New defined and table names that open with the profile locale's R1C1 letters are refused."""

from __future__ import annotations

import unittest

from formulon import DefinedNameInput, InvalidArgumentError, TableInput, Workbook

DE = "win-365-de_DE"
EN = "win-365-en_US"


class NameR1C1RefusalTest(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_for_profile(DE)
        self.addCleanup(self.wb.close)

    def test_new_defined_name_is_refused_under_de_and_accepted_under_en(self) -> None:
        for name in ("Z1S1x", "z1s1"):
            with self.subTest(name=name):
                with self.assertRaises(InvalidArgumentError):
                    self.wb.set_defined_name(DefinedNameInput(name=name, formula="Sheet1!$A$1"))
        self.assertEqual(self.wb.list_defined_names(), [])
        self.wb.set_profile_id(EN)
        self.wb.set_defined_name(DefinedNameInput(name="Z1S1x", formula="Sheet1!$A$1"))
        self.assertEqual(len(self.wb.list_defined_names()), 1)

    def test_names_that_only_resemble_the_letters_are_accepted(self) -> None:
        for name in ("Zeile", "Zone"):
            self.wb.set_defined_name(DefinedNameInput(name=name, formula="Sheet1!$A$1"))
        self.assertEqual(len(self.wb.list_defined_names()), 2)

    def test_redefining_by_case_after_a_profile_switch_replaces_the_name(self) -> None:
        self.wb.set_profile_id(EN)
        self.wb.set_defined_name(DefinedNameInput(name="Z1S1x", formula="Sheet1!$A$1"))
        self.wb.set_profile_id(DE)
        self.wb.set_defined_name(DefinedNameInput(name="z1s1X", formula="Sheet1!$A$2"))
        names = self.wb.list_defined_names()
        self.assertEqual(len(names), 1)
        self.assertEqual(names[0].name, "z1s1X")

    def test_table_names_are_refused_and_a_case_only_rename_is_allowed(self) -> None:
        ws = self.wb.sheet(0)
        ws.set_values("A1", [["Item"], ["East"]])
        with self.assertRaises(InvalidArgumentError):
            ws.add_table(TableInput(range="A1:A2", name="Z1S1T"))
        table = ws.add_table(TableInput(range="A1:A2", name="Sales"))
        with self.assertRaises(InvalidArgumentError):
            table.rename("Z1S1T")
        self.assertEqual(table.name, "Sales")

        other = Workbook.create_for_profile(EN)
        self.addCleanup(other.close)
        sheet = other.sheet(0)
        sheet.set_values("A1", [["Item"], ["East"]])
        kept = sheet.add_table(TableInput(range="A1:A2", name="Z1S1T"))
        other.set_profile_id(DE)
        kept.rename("z1s1t")
        self.assertEqual(kept.name, "z1s1t")


if __name__ == "__main__":
    unittest.main()
