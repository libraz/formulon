"""The workbook location: argument types, refusals, the canonical read back, CELL and saved hyperlink targets."""

import unittest

from formulon import (
    FormulonError,
    HyperlinkInput,
    InvalidArgumentError,
    SecurityError,
    StatusCode,
    Workbook,
    WorkbookLocation,
    WorkbookLocationInput,
)

DATA = "/Users/libraz/Library/Containers/com.microsoft.Excel/Data"


class LocationArgumentTests(unittest.TestCase):
    def test_none_unsets_and_other_types_raise_type_error(self) -> None:
        with Workbook.create_default() as wb:
            self.assertIsNone(wb.get_location())
            wb.set_location(WorkbookLocationInput("b.xlsx", directory="/d"))
            wb.set_location(None)
            self.assertIsNone(wb.get_location())
            for bad in ("b.xlsx", {"file_name": "b.xlsx"}, 1):
                with self.assertRaises(TypeError):
                    wb.set_location(bad)  # type: ignore[arg-type]
            with self.assertRaises(TypeError):
                wb.set_location(WorkbookLocationInput(None))  # type: ignore[arg-type]
            with self.assertRaises(TypeError):
                wb.set_location(WorkbookLocationInput("b.xlsx", directory=3))  # type: ignore[arg-type]
            self.assertIsNone(wb.get_location())

    def test_refusals_are_invalid_argument_and_change_nothing(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_location(WorkbookLocationInput("Keep.xlsx", directory="/keep"))
            rows = [
                ("", "/d"),
                ("Book", "/d"),
                ("a:b.xlsx", "/d"),
                ("b.xlsx", ""),
                ("b.xlsx", "relative"),
                ("b.xlsx", "C:"),
                ("b.xlsx", "//server"),
                ("b.xlsx", "/a/../b"),
            ]
            for file_name, directory in rows:
                with self.assertRaises(InvalidArgumentError, msg=(file_name, directory)):
                    wb.set_location(WorkbookLocationInput(file_name, directory=directory))
                self.assertEqual(wb.get_location(), WorkbookLocation("/keep", "Keep.xlsx"))

    def test_a_path_past_1024_bytes_is_a_resource_limit_and_changes_nothing(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_location(WorkbookLocationInput("Keep.xlsx", directory="/keep"))
            with self.assertRaises(SecurityError) as caught:
                wb.set_location(WorkbookLocationInput("b.xlsx", directory="/" + "d" * 1017))
            self.assertEqual(caught.exception.code, StatusCode.SECURITY_RESOURCE_LIMIT)
            self.assertEqual(caught.exception.context, "part=location_path requested=1025 limit=1024")
            self.assertEqual(wb.get_location(), WorkbookLocation("/keep", "Keep.xlsx"))

    def test_read_back_is_canonical_and_a_fixed_point(self) -> None:
        with Workbook.create_default() as wb:
            for given, canonical in (
                ("/a/b/", "/a/b"),
                ("/", "/"),
                ("c:\\a\\", "C:/a"),
                ("\\\\server\\share\\a", "//server/share/a"),
            ):
                wb.set_location(WorkbookLocationInput("b.xlsx", directory=given))
                read = wb.get_location()
                self.assertEqual(read, WorkbookLocation(canonical, "b.xlsx"))
                wb.set_location(read)
                self.assertEqual(wb.get_location(), read)


class LocationCellTests(unittest.TestCase):
    def test_cell_filename_and_address_follow_the_location(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_profile_id("mac-365-ja_JP")
            wb.add_sheet("Sheet2")
            ws = wb.sheet(0)
            ws.set_formula(0, 0, '=CELL("filename",A1)')
            ws.set_formula(1, 0, '=CELL("address",Sheet2!A1)')
            ws.set_formula(2, 0, '=CELL("address",Sheet1!A1)')
            wb.recalc()
            self.assertEqual(ws.get_value(0, 0), "")
            self.assertEqual(ws.get_value(1, 0), "Sheet2!$A$1")
            wb.set_location(WorkbookLocationInput("Book 1.xlsx", directory="/U/x"))
            wb.recalc()
            self.assertEqual(ws.get_value(0, 0), "/U/x/[Book 1.xlsx]Sheet1")
            self.assertEqual(ws.get_value(1, 0), "'[Book 1.xlsx]Sheet2'!$A$1")
            self.assertEqual(ws.get_value(2, 0), "$A$1")


class LocationHyperlinkSaveTests(unittest.TestCase):
    def test_saved_targets_are_relative_to_the_location(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.set_hyperlink(HyperlinkInput("A1", "/Users/libraz/Documents/a b/f.xlsx"))
            ws.set_hyperlink(HyperlinkInput("A2", "/private/tmp/x.xlsx"))
            self.assertEqual(ws.list_hyperlinks()[0].target, "file:///Users/libraz/Documents/a b/f.xlsx")

            unset = Workbook.load(wb.save()).sheet(0).list_hyperlinks()
            self.assertEqual(unset[0].target, "file:///Users/libraz/Documents/a%20b/f.xlsx")
            self.assertIsNone(unset[0].display)

            wb.set_location(WorkbookLocationInput("Book.xlsx", directory=DATA + "/loc_probe3/p q"))
            saved = Workbook.load(wb.save()).sheet(0).list_hyperlinks()
            self.assertEqual(saved[0].target, "../../../../../../Documents/a%20b/f.xlsx")
            self.assertEqual(saved[0].display, "../../../../../../Documents/a b/f.xlsx")
            self.assertEqual(saved[1].target, "file:///private/tmp/x.xlsx")
            self.assertIsNone(saved[1].display)

    def test_load_and_save_leave_the_location_alone(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_location(WorkbookLocationInput("Book.xlsx", directory="/d"))
            data = wb.save()
            self.assertEqual(wb.get_location(), WorkbookLocation("/d", "Book.xlsx"))
            with Workbook.load(data) as reloaded:
                self.assertIsNone(reloaded.get_location())


class LocationErrorTypeTests(unittest.TestCase):
    def test_invalid_argument_is_a_formulon_error(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(FormulonError):
                wb.set_location(WorkbookLocationInput("Book1"))


if __name__ == "__main__":
    unittest.main()
