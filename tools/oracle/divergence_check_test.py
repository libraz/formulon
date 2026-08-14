#!/usr/bin/env python3
"""Regression tests for `divergence_check.py`'s verified-stamp gate.

CONTRIBUTING.md forbids merging goldens captured on Office 2019, and
`Application.Version` reports the same bare `"16.0"` string for every
Office SKU from 2016 through 365 -- it is not build evidence. These
tests pin `is_pending_stamp` against that and the other non-evidence
shapes actually seen in `tests/divergence.yaml`.
"""

from __future__ import annotations

import contextlib
import io
import tempfile
import unittest
from pathlib import Path

from tools.oracle import divergence_check


class IsPendingStampTest(unittest.TestCase):
    def test_real_build_stamps_are_not_pending(self) -> None:
        for stamp in ("16.108.1", "16.108.2", "16.109", "16.111.2", "16.112"):
            self.assertFalse(divergence_check.is_pending_stamp(stamp), stamp)

    def test_bare_office_major_version_is_pending(self) -> None:
        # `Application.Version` on Office 2016 through 365 alike -- see
        # tests/oracle/variants/win-365-ja_JP/ENVIRONMENT.md.
        self.assertTrue(divergence_check.is_pending_stamp("16.0"))

    def test_doc_placeholder_shape_is_pending(self) -> None:
        self.assertTrue(divergence_check.is_pending_stamp("16.xx.x (Build 24021522)"))

    def test_prose_wrapped_build_is_pending(self) -> None:
        # The build number is real, but wrapping it in prose defeats the
        # allowlist on purpose: divergence.yaml entries must record the
        # bare stamp so a future numeric-threshold stale check works
        # without a format migration.
        self.assertTrue(divergence_check.is_pending_stamp("Excel 365 (Mac, ja-JP, 16.111.2)"))

    def test_date_instead_of_build_is_pending(self) -> None:
        self.assertTrue(divergence_check.is_pending_stamp("Excel 365 (Mac, ja-JP, 2026-07)"))

    def test_explicit_unverified_markers_are_pending(self) -> None:
        self.assertTrue(
            divergence_check.is_pending_stamp("unverified (Office 2019 reference; M365 print probe pending)")
        )
        self.assertTrue(divergence_check.is_pending_stamp("needs live Excel"))
        self.assertTrue(divergence_check.is_pending_stamp("unknown"))

    def test_missing_or_blank_stamp_is_pending(self) -> None:
        self.assertTrue(divergence_check.is_pending_stamp(None))
        self.assertTrue(divergence_check.is_pending_stamp(""))
        self.assertTrue(divergence_check.is_pending_stamp("   "))


class SkipCauseTest(unittest.TestCase):
    """`cause` is required on skip entries, and the tally counts cases.

    Nothing in the tree computes the release gate's pass rate, so this
    tally is the only place the skipped-case population is quantified.
    That makes both halves load-bearing: a missing cause has to fail, and
    the number has to be in cases rather than entries.
    """

    COMMON = 'reason: "r"\n    prefer: formulon\n    last_verified_excel_version: "16.112"\n'

    def _run(self, body: str) -> tuple[int, str]:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "divergence.yaml"
            path.write_text(body, encoding="utf-8")
            out = io.StringIO()
            err = io.StringIO()
            with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
                status = divergence_check.validate(path, strict=False)
            return status, out.getvalue() + err.getvalue()

    def test_skip_without_cause_fails(self) -> None:
        status, output = self._run(f"entries:\n  - id: volatile_now_clock\n    mode: skip-oracle\n    {self.COMMON}")
        self.assertEqual(status, 1)
        self.assertIn("needs `cause`", output)

    def test_skip_with_unknown_cause_fails(self) -> None:
        status, output = self._run(
            f"entries:\n  - id: volatile_now_clock\n    mode: skip-oracle\n    cause: because\n    {self.COMMON}"
        )
        self.assertEqual(status, 1)
        self.assertIn("needs `cause`", output)

    def test_tolerance_entry_needs_no_cause(self) -> None:
        status, _ = self._run(
            f"entries:\n  - id: volatile_now_clock\n    tolerance: {{ abs: 1.0, rel: 0 }}\n    {self.COMMON}"
        )
        self.assertEqual(status, 0)

    def test_tally_counts_cases_not_entries(self) -> None:
        # One `suite:` entry removes every case in that suite; one `ids:`
        # entry removes as many cases as it lists.
        status, output = self._run(
            "entries:\n"
            f"  - suite: filterxml\n    mode: skip-oracle\n    cause: harness-cannot-capture\n    {self.COMMON}"
            "  - ids:\n      - volatile_now_clock\n      - volatile_today_clock\n"
            f"    mode: skip-oracle\n    cause: excel-no-value\n    {self.COMMON}"
        )
        self.assertEqual(status, 0)
        filterxml_cases = divergence_check.load_case_catalog()[1]["filterxml"]
        self.assertIn(f"harness-cannot-capture: {filterxml_cases}", output)
        self.assertIn("excel-no-value: 2", output)
        self.assertIn(f"skipped cases by cause: {filterxml_cases + 2} total", output)

    def test_cross_registry_prefer_conflict_is_detected(self) -> None:
        # The same case recorded with opposite verdicts in two registries:
        # found by hand once, checked mechanically now.
        common = 'reason: "r"\n    mode: skip-oracle\n    cause: accepted-divergence\n'
        with tempfile.TemporaryDirectory() as tmp:
            primary = Path(tmp) / "divergence.yaml"
            variant = Path(tmp) / "variant.yaml"
            primary.write_text(
                f"entries:\n  - id: volatile_now_clock\n    {common}    prefer: mac-excel-365\n", encoding="utf-8"
            )
            variant.write_text(
                f"entries:\n  - id: volatile_now_clock\n    {common}    prefer: formulon\n", encoding="utf-8"
            )
            conflicts = divergence_check.cross_file_prefer_conflicts([primary, variant])
            self.assertEqual(len(conflicts), 1, conflicts)
            self.assertIn("volatile_now_clock", conflicts[0])
            self.assertIn("mac-excel-365", conflicts[0])
            self.assertIn("formulon", conflicts[0])

            # Agreeing registries are not a conflict.
            variant.write_text(
                f"entries:\n  - id: volatile_now_clock\n    {common}    prefer: mac-excel-365\n", encoding="utf-8"
            )
            self.assertEqual(divergence_check.cross_file_prefer_conflicts([primary, variant]), [])

    def test_committed_registries_have_no_prefer_conflicts(self) -> None:
        self.assertEqual(divergence_check.cross_file_prefer_conflicts(), [])

    def test_non_oracle_scope_entry_does_not_inflate_the_tally(self) -> None:
        status, output = self._run(
            "entries:\n"
            "  - id: not_an_oracle_case\n"
            "    mode: skip-oracle\n"
            "    cause: accepted-divergence\n"
            "    scope: unit-test\n"
            "    evidence: [tools/oracle/divergence_check.py]\n"
            f"    {self.COMMON}"
        )
        self.assertEqual(status, 0)
        self.assertIn("skipped cases by cause: 0 total", output)


if __name__ == "__main__":
    unittest.main()
