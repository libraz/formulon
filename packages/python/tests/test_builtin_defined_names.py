"""A defined name spelled as the profile locale's built-in name is stored as that built-in."""

from __future__ import annotations

import io
import unittest
import zipfile
from dataclasses import replace
from pathlib import Path

from formulon import (
    AlreadyExistsError,
    CellRect,
    DefinedName,
    DefinedNameInput,
    DefinedNameKey,
    FormulonError,
    IndexOutOfRangeError,
    InvalidArgumentError,
    TableInput,
    Workbook,
    canonicalize_defined_name,
    list_profile_ids,
    localize_defined_name,
)

BUILT_IN_NAMES = (
    "Print_Area",
    "Print_Titles",
    "_FilterDatabase",
    "Criteria",
    "Extract",
    "Database",
    "Consolidate_Area",
    "Sheet_Title",
    "Auto_Open",
    "Recorder",
    "Auto_Close",
    "Data_Form",
    "Auto_Activate",
    "Auto_Deactivate",
)
SHEET_BUILT_INS = ("Print_Area", "Print_Titles")

_FIXTURES = Path(__file__).resolve().parents[3] / "tests/fixtures/excel"
# The Excel-saved workbook of the worksheet `Sheet1` then the chartsheet `Chart1`.
_CHARTSHEET_BOOK = "chartsheet_twins/tab_color_rgb_red.xlsx"
_CHARTSHEET_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet"


def _fixture(relative: str) -> bytes:
    return (_FIXTURES / relative).read_bytes()


def _second_sheet_retyped(rel_type: str) -> bytes:
    """The chartsheet book with its second sheet's relationship retyped to ``rel_type``."""
    source = zipfile.ZipFile(io.BytesIO(_fixture(_CHARTSHEET_BOOK)))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as target:
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename == "xl/_rels/workbook.xml.rels":
                text = data.decode()
                assert text.count(_CHARTSHEET_REL) == 1
                data = text.replace(_CHARTSHEET_REL, rel_type).encode()
            target.writestr(item, data)
    return out.getvalue()


def mac_profiles() -> list:
    return [p for p in list_profile_ids() if p.startswith("mac-365-")]


def error_type(fn):
    try:
        fn()
    except FormulonError as e:
        return type(e)
    return None


class BuiltinDefinedNameTest(unittest.TestCase):
    def test_new_print_area_in_workbook_scope_goes_to_the_active_worksheet(self) -> None:
        wb = Workbook.create_for_profile("mac-365-ja_JP")
        self.addCleanup(wb.close)
        wb.add_sheet("Sheet2")
        wb.add_sheet("Sheet3")
        wb.set_active_sheet(2)
        wb.set_defined_name(
            DefinedNameInput(name="print_area", formula="=Sheet2!$A$1:$B$2", comment="note", hidden=True)
        )
        self.assertEqual(
            wb.list_defined_names(),
            [DefinedName(name="_xlnm.Print_Area", formula="=Sheet2!$A$1:$B$2", scope=2, comment="note", hidden=True)],
        )
        self.assertIsNone(wb.get_defined_name("Print_Area"))
        wb.set_defined_name(DefinedNameInput(name="_xlnm.Print_Titles", formula="=Sheet1!$1:$1"))
        titles = wb.get_defined_name("Print_Titles", 2)
        assert titles is not None
        self.assertEqual(titles.name, "_xlnm.Print_Titles")
        # One asked for in a sheet's scope stays there.
        wb.set_defined_name(DefinedNameInput(name="Print_Area", formula="=Sheet2!$A$1:$B$2", scope=1))
        self.assertEqual(wb.sheet(1).get_print_area(), [CellRect(0, 0, 1, 1)])
        wb.sheet(1).set_print_area("A1:C3")
        self.assertEqual(len(wb.list_defined_names()), 3)
        wb.remove_defined_name("PRINT_AREA", 1)
        self.assertIsNone(wb.get_defined_name("Print_Area", 1))
        kept = wb.get_defined_name("Print_Area", 2)
        assert kept is not None
        self.assertEqual(kept.formula, "=Sheet2!$A$1:$B$2")

    def test_only_the_locale_spelling_is_a_built_in(self) -> None:
        wb = Workbook.create_for_profile("mac-365-de_DE")
        self.addCleanup(wb.close)
        sheet = wb.sheet(0).name
        wb.set_defined_name(DefinedNameInput(name="Print_Area", formula=f"={sheet}!$A$1:$B$2"))
        wb.set_defined_name(DefinedNameInput(name="Druckbereich", formula=f"={sheet}!$A$1:$B$2"))
        wb.set_defined_name(DefinedNameInput(name="_xlnm.Print_Titles", formula=f"={sheet}!$1:$1"))
        self.assertEqual(
            [(n.name, n.scope) for n in wb.list_defined_names()],
            [("Print_Area", None), ("_xlnm.Print_Area", 0), ("_xlnm.Print_Titles", 0)],
        )

    def test_new_print_area_in_workbook_scope_stays_there_with_a_chartsheet_active(self) -> None:
        wb = Workbook.load(_fixture(_CHARTSHEET_BOOK))
        self.addCleanup(wb.close)
        wb.set_active_sheet(1)
        before = len(wb.list_defined_names())
        wb.set_defined_name(DefinedNameInput(name="Print_Area", formula="=Sheet1!$A$1:$B$2"))
        found = wb.get_defined_name("Print_Area")
        assert found is not None
        self.assertEqual((found.name, found.scope), ("_xlnm.Print_Area", None))
        wb.set_defined_name(DefinedNameInput(name="_xlnm.Print_Area", formula="=Sheet1!$A$1:$C$3"))
        self.assertEqual(len(wb.list_defined_names()), before + 1)
        found = wb.get_defined_name("Print_Area")
        assert found is not None
        self.assertEqual(found.formula, "=Sheet1!$A$1:$C$3")
        wb.remove_defined_name("Print_Area")
        self.assertEqual(len(wb.list_defined_names()), before)

    def test_new_print_area_in_workbook_scope_is_refused_with_another_sheet_kind_active(self) -> None:
        for rel_type in (
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet",
            "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet",
        ):
            with self.subTest(rel_type=rel_type):
                wb = Workbook.load(_second_sheet_retyped(rel_type))
                self.addCleanup(wb.close)
                wb.set_active_sheet(1)
                for name in ("Print_Area", "_xlnm.Print_Titles"):
                    with self.assertRaises(InvalidArgumentError):
                        wb.set_defined_name(DefinedNameInput(name=name, formula="=Sheet1!$A$1:$B$2"))
                    with self.assertRaises(InvalidArgumentError):
                        wb.resolve_defined_name_key(DefinedNameInput(name=name, formula="=Sheet1!$A$1:$B$2"))
                wb.set_defined_name(DefinedNameInput(name="Print_Area", formula="=Sheet1!$A$1:$B$2", scope=0))
                self.assertIsNotNone(wb.get_defined_name("Print_Area", 0))

    def test_ordinary_name_spelled_as_the_built_in_is_addressed_by_its_spelling(self) -> None:
        # An ordinary workbook-scope Print_Area, read as such under de; under en it is the built-in's spelling.
        wb = Workbook.load(_fixture("xlsb_builtin_name/ordinary_source.xlsb"), profile_id="mac-365-de_DE")
        self.addCleanup(wb.close)
        wb.set_profile_id("mac-365-en_US")
        wb.set_defined_name(DefinedNameInput(name="_xlnm.Print_Area", formula="=Sheet1!$A$1:$A$3"))
        builtin = wb.get_defined_name("_xlnm.Print_Area", 0)
        assert builtin is not None
        self.assertEqual(builtin.formula, "=Sheet1!$A$1:$A$3")
        wb.set_defined_name(DefinedNameInput(name="Print_Area", formula="=Sheet1!$A$1:$A$4"))
        ordinary = wb.get_defined_name("Print_Area")
        assert ordinary is not None
        self.assertEqual((ordinary.name, ordinary.formula), ("Print_Area", "=Sheet1!$A$1:$A$4"))
        wb.remove_defined_name("Print_Area")
        self.assertIsNone(wb.get_defined_name("Print_Area"))
        builtin = wb.get_defined_name("Print_Area", 0)
        assert builtin is not None
        self.assertEqual(builtin.name, "_xlnm.Print_Area")

    def test_load_folds_a_name_spelled_as_the_load_locales_built_in(self) -> None:
        for profile, expected in (("mac-365-en_US", "_xlnm.Print_Area"), ("mac-365-de_DE", "Print_Area")):
            with self.subTest(profile=profile):
                wb = Workbook.load(_fixture("xlsb_builtin_name/ordinary_source.xlsb"), profile_id=profile)
                self.addCleanup(wb.close)
                self.assertEqual(
                    [(n.name, n.scope) for n in wb.list_defined_names()], [("Ctl_8", None), (expected, None)]
                )

    def test_resolved_key_is_where_set_stores_every_built_in_in_every_locale(self) -> None:
        profiles = mac_profiles()
        self.assertEqual(len(profiles), 14)
        for profile in profiles:
            for english in BUILT_IN_NAMES:
                typed = localize_defined_name(f"_xlnm.{english}", profile)
                with self.subTest(profile=profile, name=typed):
                    wb = Workbook.create_for_profile(profile)
                    self.addCleanup(wb.close)
                    wb.add_sheet("Sheet2")
                    item = DefinedNameInput(name=typed, formula="=Sheet2!$A$1:$B$2")
                    key = wb.resolve_defined_name_key(item)
                    self.assertEqual(wb.list_defined_names(), [])
                    self.assertEqual(key.name, canonicalize_defined_name(typed, profile))
                    self.assertEqual(key.scope, 0 if english in SHEET_BUILT_INS else None)
                    wb.set_defined_name(item)
                    self.assertEqual([(n.name, n.scope) for n in wb.list_defined_names()], [(key.name, key.scope)])
                    found = wb.get_defined_name(key.name, key.scope)
                    self.assertIsNotNone(found)
                    self.assertEqual(localize_defined_name(key.name, profile), typed)

    def test_resolver_refuses_what_set_refuses_with_the_same_error(self) -> None:
        for profile in mac_profiles():
            with self.subTest(profile=profile):
                wb = Workbook.create_for_profile(profile)
                self.addCleanup(wb.close)
                wb.add_sheet("Sheet2")
                sheet = wb.sheet(0)
                sheet.set_text("A1", "a")
                sheet.set_text("B1", "b")
                sheet.add_table(TableInput(range="A1:B2", name="Sales"))
                wb.set_defined_name(DefinedNameInput(name="Taken", formula="=Sheet1!$A$1"))
                cases = [
                    DefinedNameInput(name="_xlnm.Print_Area", formula="=$A$1:$B$2"),
                    DefinedNameInput(name="_xlnm.Print_Titles", formula="=$1:$1"),
                    DefinedNameInput(name="Sales", formula="=Sheet1!$A$1"),
                    DefinedNameInput(name="Fresh", formula="=Sheet1!$A$1", scope=7),
                    DefinedNameInput(name="Fresh", formula=""),
                    DefinedNameInput(name="Fresh", formula="=SUM("),
                    DefinedNameInput(name="", formula="=Sheet1!$A$1"),
                    DefinedNameInput(name="1Fresh", formula="=Sheet1!$A$1"),
                    DefinedNameInput(name="taken", formula="=Sheet1!$B$1"),
                    DefinedNameInput(name="Fresh", formula="=Sheet1!$A$1", scope=1),
                ]
                for item in cases:
                    before = len(wb.list_defined_names())
                    keys = []
                    resolved = error_type(lambda: keys.append(wb.resolve_defined_name_key(item)))
                    self.assertEqual(len(wb.list_defined_names()), before)
                    refused = error_type(lambda: wb.set_defined_name(item))
                    self.assertIs(resolved, refused, item)
                    if refused is None:
                        self.assertIsNotNone(wb.get_defined_name(keys[0].name, keys[0].scope))
        wb = Workbook.create_for_profile("mac-365-en_US")
        self.addCleanup(wb.close)
        with self.assertRaises(InvalidArgumentError):
            wb.resolve_defined_name_key(DefinedNameInput(name="1Fresh", formula="=Sheet1!$A$1"))
        with self.assertRaises(IndexOutOfRangeError):
            wb.resolve_defined_name_key(DefinedNameInput(name="Fresh", formula="=Sheet1!$A$1", scope=9))
        wb.sheet(0).add_table(TableInput(range="A1:B2", name="Sales"))
        with self.assertRaises(AlreadyExistsError):
            wb.resolve_defined_name_key(DefinedNameInput(name="Sales", formula="=Sheet1!$A$2"))

    def test_resolved_key_lets_a_caller_keep_the_attributes_a_set_resets(self) -> None:
        wb = Workbook.create_for_profile("mac-365-de_DE")
        self.addCleanup(wb.close)
        wb.add_sheet("Sheet2")
        wb.set_defined_name(
            DefinedNameInput(name="Druckbereich", formula="=Sheet2!$A$1:$B$2", comment="note", hidden=True)
        )
        item = DefinedNameInput(name="DRUCKBEREICH", formula="=Sheet2!$A$1:$C$3")
        key = wb.resolve_defined_name_key(item)
        self.assertEqual(key, DefinedNameKey(name="_xlnm.Print_Area", scope=0))
        prev = wb.get_defined_name(key.name, key.scope)
        assert prev is not None
        wb.set_defined_name(replace(item, comment=prev.comment, hidden=prev.hidden))
        self.assertEqual(
            wb.list_defined_names(),
            [DefinedName(name="_xlnm.Print_Area", formula="=Sheet2!$A$1:$C$3", scope=0, comment="note", hidden=True)],
        )
        wb.set_defined_name(item)
        found = wb.get_defined_name(key.name, key.scope)
        assert found is not None
        self.assertIsNone(found.comment)
        # A read record is a valid input.
        self.assertEqual(wb.resolve_defined_name_key(wb.list_defined_names()[0]), key)

    def test_localize_and_canonicalize_map_between_stored_and_displayed_spelling(self) -> None:
        self.assertEqual(localize_defined_name("_xlnm.Print_Area", "mac-365-de_DE"), "Druckbereich")
        self.assertEqual(localize_defined_name("_xlnm.PRINT_AREA", "mac-365-de_DE"), "Druckbereich")
        self.assertEqual(localize_defined_name("_xlnm.Print_Area", "mac-365-fr_FR"), "Zone_d_impression")
        self.assertEqual(localize_defined_name("_xlnm.Print_Area", "mac-365-ja_JP"), "Print_Area")
        self.assertEqual(localize_defined_name("_xlnm.print_area", "mac-365-en_US"), "print_area")
        self.assertEqual(localize_defined_name("_xlnm.Foo", "mac-365-en_US"), "_xlnm.Foo")
        self.assertEqual(localize_defined_name("Sales", "mac-365-de_DE"), "Sales")
        self.assertEqual(canonicalize_defined_name("druckbereich", "mac-365-de_DE"), "_xlnm.Print_Area")
        self.assertEqual(canonicalize_defined_name("Print_Area", "mac-365-de_DE"), "Print_Area")
        self.assertEqual(canonicalize_defined_name("print_area", "mac-365-en_US"), "_xlnm.print_area")
        self.assertEqual(canonicalize_defined_name("_xlnm.Foo", "mac-365-en_US"), "_xlnm.Foo")
        self.assertEqual(canonicalize_defined_name("Sales", "mac-365-en_US"), "Sales")
        for profile in mac_profiles():
            for english in BUILT_IN_NAMES:
                stored = f"_xlnm.{english}"
                with self.subTest(profile=profile, name=stored):
                    shown = localize_defined_name(stored, profile)
                    self.assertEqual(canonicalize_defined_name(shown, profile), stored)
        for call in (
            lambda: localize_defined_name("", "mac-365-en_US"),
            lambda: canonicalize_defined_name("", "mac-365-en_US"),
            lambda: localize_defined_name("Print_Area", "no-such-profile"),
            lambda: canonicalize_defined_name("Print_Area", "no-such-profile"),
        ):
            with self.assertRaises(InvalidArgumentError):
                call()
        with self.assertRaises(TypeError):
            localize_defined_name(None, "mac-365-en_US")  # type: ignore[arg-type]


if __name__ == "__main__":
    unittest.main()
