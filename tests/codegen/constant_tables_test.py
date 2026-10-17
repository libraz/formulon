#!/usr/bin/env python3
"""Tests for the embedded numeric table checker: generators, parser and mismatch report."""

from __future__ import annotations

import importlib.util
import sys
import tempfile
import unittest
from decimal import Decimal
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
CHECKER_PATH = ROOT / "tools" / "numeric_tables" / "check_constant_tables.py"


def _load_checker():
    spec = importlib.util.spec_from_file_location("formulon_constant_tables", CHECKER_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load {CHECKER_PATH}")
    module = importlib.util.module_from_spec(spec)
    # The checker's dataclass resolves its module through sys.modules.
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


c = _load_checker()


class GeneratorTest(unittest.TestCase):
    def test_bernoulli(self):
        b = c.bernoulli(12)
        self.assertEqual(str(b[2]), "1/6")
        self.assertEqual(str(b[12]), "-691/2730")

    def test_stirling_leading_terms(self):
        t = c.gamma_stirling()
        self.assertEqual(t[-1], 2.0)
        self.assertEqual(t[-2], 2 * (1 / 12))
        self.assertEqual(t[-3], 2 / 288)

    def test_temme_leading_terms(self):
        t = c.temme_seed()
        self.assertEqual(t[-1], -1 / 3)
        self.assertEqual(t[-2], 1 / 12)
        self.assertEqual(t[-3], -2 / 135)

    def test_pi_and_euler_gamma(self):
        self.assertEqual(c.rn(c.dec_pi()), 3.141592653589793)
        self.assertEqual(c.rn(c.dec_euler_gamma()), 0.5772156649015329)

    def test_truncation_is_toward_zero(self):
        self.assertEqual(c.trunc_sig(Decimal("-1.239"), 3), -1.23)
        self.assertEqual(c.trunc_sig(Decimal("0.001239"), 3), 0.00123)


class ParserTest(unittest.TestCase):
    def test_literals(self):
        self.assertEqual(c.parse_literal("-0x1.8p+1"), -3.0)
        self.assertEqual(c.parse_literal("0x0.0p+0"), 0.0)
        self.assertEqual(c.parse_literal("1e-16"), 1e-16)

    def test_array_and_scalar(self):
        src = c.strip_comments(
            "// kA = {9}\nconstexpr double kA[] = {1.0, /* x */ -0x1p-2,\n 3e0};\nconstexpr double kS = 0.5;"
        )
        self.assertEqual(c.find_array(src, "kA"), [1.0, -0.25, 3.0])
        self.assertEqual(c.find_scalar(src, "kS"), 0.5)
        self.assertIsNone(c.find_array(src, "kMissing"))


class TreeTest(unittest.TestCase):
    def test_current_tree_matches(self):
        for chk in c.build_checks():
            ok, msg = c.run_check(chk, ROOT)
            self.assertTrue(ok, msg)

    def test_mismatch_reports_ulp(self):
        with tempfile.TemporaryDirectory() as tmp:
            dst = Path(tmp) / "src" / "eval" / "stats"
            dst.mkdir(parents=True)
            text = (ROOT / "src/eval/stats/incomplete_gamma.cpp").read_text(encoding="utf-8")
            text = text.replace("0x1.9884533d43651p-2", "0x1.9884533d43652p-2")
            (dst / "incomplete_gamma.cpp").write_text(text, encoding="utf-8")
            chk = next(x for x in c.build_checks() if x.name == "kPhi0")
            ok, msg = c.run_check(chk, Path(tmp))
            self.assertFalse(ok)
            self.assertIn("ulp diff 1", msg)

    def test_missing_symbol_fails(self):
        with tempfile.TemporaryDirectory() as tmp:
            dst = Path(tmp) / "src" / "eval" / "stats"
            dst.mkdir(parents=True)
            (dst / "incomplete_gamma.cpp").write_text("", encoding="utf-8")
            chk = next(x for x in c.build_checks() if x.name == "kTwoPi")
            self.assertFalse(c.run_check(chk, Path(tmp))[0])


if __name__ == "__main__":
    unittest.main()
