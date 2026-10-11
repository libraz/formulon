"""Tests for the collation-aware text compare and value classification."""

from __future__ import annotations

import unittest

from formulon import CollationClasses, CollationStrength, FormulonError, Value, ValueKind, Workbook

ERROR_DIV0 = 1
ERROR_NA = 6


def _num(x: float) -> Value:
    return Value(kind=ValueKind.NUMBER, number=x)


def _text(s: str) -> Value:
    return Value(kind=ValueKind.TEXT, text=s)


def _bool(b: bool) -> Value:
    return Value(kind=ValueKind.BOOL, boolean=b)


def _blank() -> Value:
    return Value(kind=ValueKind.BLANK)


def _err(code: int) -> Value:
    return Value(kind=ValueKind.ERROR, error_code=code)


class CompareTextTests(unittest.TestCase):
    def test_sign_and_argument_order(self) -> None:
        profile = "mac-365-ja_JP"
        s = CollationStrength.COMPARE
        self.assertEqual(Workbook.compare_text("a", "b", s, profile), -1)
        self.assertEqual(Workbook.compare_text("b", "a", s, profile), 1)
        self.assertEqual(Workbook.compare_text("a", "a", s, profile), 0)

    def test_strength_changes_the_fold(self) -> None:
        profile = "mac-365-zh_CN"
        self.assertNotEqual(Workbook.compare_text("A", "Ａ", CollationStrength.COMPARE, profile), 0)
        self.assertEqual(Workbook.compare_text("A", "Ａ", CollationStrength.DYNAMIC, profile), 0)

    def test_rejections(self) -> None:
        with self.assertRaises(FormulonError):
            Workbook.compare_text("a", "b", CollationStrength.COMPARE, "mac-365-xx_XX")  # type: ignore[arg-type]
        with self.assertRaises(FormulonError):
            Workbook.compare_text("a", "b", 3, "mac-365-ja_JP")  # type: ignore[arg-type]


class CollationClassesTests(unittest.TestCase):
    PROFILE = "mac-365-ja_JP"

    def test_type_boundaries_and_order(self) -> None:
        values = [
            _num(1),
            _text("1"),
            _bool(True),
            _text("TRUE"),
            _blank(),
            _text(""),
            _num(0),
            _num(-0.0),
            _bool(False),
        ]
        for strength in CollationStrength:
            result = Workbook.collation_classes(values, strength, self.PROFILE)
            self.assertIsInstance(result, CollationClasses)
            ids = result.class_ids
            self.assertEqual(result.class_count, 8)
            self.assertEqual(len(ids), len(values))
            self.assertEqual(ids[6], ids[7])
            self.assertEqual(sorted(set(ids)), list(range(8)))
            # number < text < bool < blank
            self.assertLess(ids[6], ids[0])
            self.assertLess(ids[0], ids[5])
            self.assertLess(ids[5], ids[1])
            self.assertLess(ids[1], ids[3])
            self.assertLess(ids[3], ids[8])
            self.assertLess(ids[8], ids[2])
            self.assertLess(ids[2], ids[4])

    def test_errors_order_by_code(self) -> None:
        result = Workbook.collation_classes(
            [_err(ERROR_NA), _err(ERROR_DIV0), _err(ERROR_NA)], CollationStrength.DYNAMIC, self.PROFILE
        )
        self.assertEqual(result.class_ids, [1, 0, 1])
        self.assertEqual(result.class_count, 2)

    def test_empty(self) -> None:
        result = Workbook.collation_classes([], CollationStrength.COMPARE, self.PROFILE)
        self.assertEqual(result.class_ids, [])
        self.assertEqual(result.class_count, 0)

    def test_rejections(self) -> None:
        ok = [_num(1)]
        with self.assertRaises(FormulonError):
            Workbook.collation_classes(ok, CollationStrength.COMPARE, "mac-365-xx_XX")  # type: ignore[arg-type]
        with self.assertRaises(FormulonError):
            Workbook.collation_classes(ok, 3, self.PROFILE)  # type: ignore[arg-type]
        with self.assertRaises(FormulonError):
            Workbook.collation_classes([_num(float("inf"))], CollationStrength.COMPARE, self.PROFILE)
        with self.assertRaises(FormulonError):
            Workbook.collation_classes([_num(float("nan"))], CollationStrength.COMPARE, self.PROFILE)
        with self.assertRaises(FormulonError):
            Workbook.collation_classes([_err(99)], CollationStrength.COMPARE, self.PROFILE)


class CollationMeasuredTests(unittest.TestCase):
    def test_follows_the_host(self) -> None:
        self.assertFalse(Workbook.locale_facts("win-365-ja_JP").collation_measured)
        self.assertTrue(Workbook.locale_facts("mac-365-ja_JP").collation_measured)


if __name__ == "__main__":
    unittest.main()
