"""Workbook.get_compatibility_version / set_compatibility_version: the stored version, its effect on
text-position functions, and the refusals."""

from __future__ import annotations

import io
import re
import unittest
import zipfile
from typing import Callable

from formulon import CalcMode, CompatibilityVersion, InvalidArgumentError, Workbook


class CompatibilityVersionTests(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.addCleanup(self.wb.close)
        self.ws = self.wb.sheet(0)

    def test_enum_values(self) -> None:
        self.assertEqual(
            {m.name: int(m) for m in CompatibilityVersion},
            {"VERSION_1": 1, "VERSION_2": 2, "VERSION_3": 3},
        )

    def test_new_workbook_reads_version_2(self) -> None:
        self.assertIs(self.wb.get_compatibility_version(), CompatibilityVersion.VERSION_2)

    def test_len_counts_utf16_units_under_version_1_and_code_points_under_version_2(self) -> None:
        self.ws.set_formula("A1", '=LEN("a\U0001f600")')
        self.wb.set_compatibility_version(CompatibilityVersion.VERSION_1)
        self.assertIs(self.wb.get_compatibility_version(), CompatibilityVersion.VERSION_1)
        self.assertEqual(self.ws.get_value("A1"), 3)
        self.wb.set_compatibility_version(CompatibilityVersion.VERSION_2)
        self.assertIs(self.wb.get_compatibility_version(), CompatibilityVersion.VERSION_2)
        self.assertEqual(self.ws.get_value("A1"), 2)

    def test_version_3_is_stored_and_marks_formulas_only_across_version_1(self) -> None:
        self.wb.set_calc_mode(CalcMode.MANUAL)
        self.ws.set_formula("A1", '=LEN("a\U0001f600")')
        self.wb.recalc()
        self.wb.set_compatibility_version(CompatibilityVersion.VERSION_3)
        self.assertIs(self.wb.get_compatibility_version(), CompatibilityVersion.VERSION_3)
        self.assertFalse(self.ws.is_cell_stale("A1"))
        self.wb.set_compatibility_version(CompatibilityVersion.VERSION_1)
        self.assertTrue(self.ws.is_cell_stale("A1"))
        self.wb.recalc()
        self.assertEqual(self.ws.get_value("A1"), 3)
        self.wb.set_compatibility_version(CompatibilityVersion.VERSION_3)
        self.assertIs(self.wb.get_compatibility_version(), CompatibilityVersion.VERSION_3)
        self.assertTrue(self.ws.is_cell_stale("A1"))
        self.wb.recalc()
        self.assertEqual(self.ws.get_value("A1"), 2)

    def test_value_outside_the_enum_is_invalid_argument(self) -> None:
        for bad in (0, 4, -1):
            with self.subTest(bad=bad):
                with self.assertRaises(InvalidArgumentError):
                    self.wb.set_compatibility_version(bad)  # type: ignore[arg-type]
        self.assertIs(self.wb.get_compatibility_version(), CompatibilityVersion.VERSION_2)

    def test_non_integer_argument_is_a_type_error(self) -> None:
        for bad in ("2", None, 1.5, float("nan"), True):
            with self.subTest(bad=bad):
                with self.assertRaises(TypeError):
                    self.wb.set_compatibility_version(bad)  # type: ignore[arg-type]
        self.assertIs(self.wb.get_compatibility_version(), CompatibilityVersion.VERSION_2)


_STAMP_EXT = re.compile(r'<ext uri="\{D14903EA-33C4-47F7-8F05-3474C54BE107\}".*?</ext>')


def _reload_with_stamp(edit: Callable[[str], str]) -> Workbook:
    """A new workbook saved and reloaded with its workbook.xml rewritten by ``edit``; A1 holds =LEN("a😀b")."""
    wb = Workbook.create_default()
    try:
        wb.sheet(0).set_formula("A1", '=LEN("a\U0001f600b")')
        saved = wb.save()
    finally:
        wb.close()
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(saved)) as src, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == "xl/workbook.xml":
                text = data.decode("utf-8")
                edited = edit(text)
                assert edited != text
                data = edited.encode("utf-8")
            dst.writestr(info, data)
    return Workbook.load(out.getvalue())


class LoadedStampTests(unittest.TestCase):
    def test_a_loaded_workbook_without_a_stamp_reads_version_1(self) -> None:
        wb = _reload_with_stamp(lambda xml: _STAMP_EXT.sub("", xml, count=1))
        self.addCleanup(wb.close)
        self.assertIs(wb.get_compatibility_version(), CompatibilityVersion.VERSION_1)
        wb.recalc()
        self.assertEqual(wb.sheet(0).get_value("A1"), 4)

    def test_a_stamp_above_3_reads_as_stored_and_evaluates_as_version_2(self) -> None:
        wb = _reload_with_stamp(lambda xml: xml.replace('setVersion="2"', 'setVersion="4"', 1))
        self.addCleanup(wb.close)
        version = wb.get_compatibility_version()
        self.assertIsInstance(version, CompatibilityVersion)
        self.assertEqual(version, 4)
        self.assertNotIn(version, list(CompatibilityVersion))
        wb.recalc()
        self.assertEqual(wb.sheet(0).get_value("A1"), 3)
        with self.assertRaises(InvalidArgumentError):
            wb.set_compatibility_version(version)
        again = Workbook.load(wb.save())
        self.addCleanup(again.close)
        self.assertEqual(again.get_compatibility_version(), 4)


if __name__ == "__main__":
    unittest.main()
