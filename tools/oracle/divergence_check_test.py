#!/usr/bin/env python3
"""Regression tests for `divergence_check.py`'s verified-stamp gate.

CONTRIBUTING.md forbids merging goldens captured on Office 2019, and
`Application.Version` reports the same bare `"16.0"` string for every
Office SKU from 2016 through 365 -- it is not build evidence. These
tests pin `is_pending_stamp` against that and the other non-evidence
shapes actually seen in `tests/divergence.yaml`.
"""

from __future__ import annotations

import unittest

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


if __name__ == "__main__":
    unittest.main()
