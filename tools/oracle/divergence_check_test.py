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
import json
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
        # bare stamp, because that is what `is_stale_stamp` compares
        # numerically against the accepted build floor.
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


class IsStaleStampTest(unittest.TestCase):
    """A real but old build is evidence of what Excel did, not of what it does.

    `is_pending_stamp` cannot carry this: it also decides whether a fresh
    capture may be promoted, and a capture taken on an older build is
    still a genuine capture. So staleness is a second predicate, and only
    the divergence registry applies it.
    """

    def test_build_below_its_floor_is_stale(self) -> None:
        for stamp in ("16.108.1", "16.108.2", "16.109"):
            self.assertTrue(divergence_check.is_stale_stamp(stamp), stamp)

    def test_build_at_or_above_its_floor_is_not_stale(self) -> None:
        for stamp in ("16.110", "16.111.2", "16.112"):
            self.assertFalse(divergence_check.is_stale_stamp(stamp), stamp)

    def test_windows_build_is_compared_against_the_windows_floor(self) -> None:
        # `Application.Version` is pinned at the Office major on Windows,
        # so its builds read as 16.0.<n> and are not on the macOS number
        # line -- comparing them against the macOS floor would call every
        # Windows capture stale.
        self.assertFalse(divergence_check.is_stale_stamp("16.0.20228"))
        self.assertTrue(divergence_check.is_stale_stamp("16.0.18025"))

    def test_pending_stamps_are_not_reported_as_stale(self) -> None:
        for stamp in (None, "", "16.0", "unknown", "16.xx.x (Build 24021522)"):
            self.assertFalse(divergence_check.is_stale_stamp(stamp), stamp)


class StaleStampGateTest(unittest.TestCase):
    """`--strict` fails on a stale stamp, and the tally leaves it out."""

    def _run(self, body: str, *, strict: bool) -> tuple[int, str]:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "divergence.yaml"
            path.write_text(body, encoding="utf-8")
            out = io.StringIO()
            err = io.StringIO()
            with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
                status = divergence_check.validate(path, strict=strict)
            return status, out.getvalue() + err.getvalue()

    BODY = (
        "entries:\n"
        "  - id: volatile_now_clock\n"
        "    mode: skip-oracle\n"
        "    cause: excel-no-value\n"
        '    reason: "r"\n'
        "    prefer: formulon\n"
        '    last_verified_excel_version: "16.108.1"\n'
    )

    def test_strict_fails_on_a_stale_stamp(self) -> None:
        status, output = self._run(self.BODY, strict=True)
        self.assertEqual(status, 1)
        self.assertIn("STALE volatile_now_clock", output)

    def test_non_strict_reports_but_does_not_fail(self) -> None:
        status, output = self._run(self.BODY, strict=False)
        self.assertEqual(status, 0)
        self.assertIn("STALE volatile_now_clock", output)

    def test_stale_entry_is_not_tallied_under_its_cause(self) -> None:
        _, output = self._run(self.BODY, strict=False)
        self.assertIn("excel-no-value: 0", output)
        self.assertIn("not tallied: 1", output)

    def test_current_stamp_is_tallied(self) -> None:
        status, output = self._run(self.BODY.replace("16.108.1", "16.112"), strict=True)
        self.assertEqual(status, 0)
        self.assertIn("excel-no-value: 1", output)
        self.assertIn("not tallied: 0", output)


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


class IronCalcRegistryTest(unittest.TestCase):
    """The IronCalc skip registry is held to the same three properties.

    Every entry resolves to a case that still exists, carries a
    machine-checkable reason for not running, and is tallied in cases
    against the corpus it is removed from. Without the first, a renamed
    suite turns a skip into a permanent no-op that hides whatever
    regression later lands on the id.
    """

    GOLDEN = {"suite": "ironcalc_demo_Sheet1", "cases": [{"id": "A1"}, {"id": "A2"}]}
    PROBE = {"suite": "demo_probes", "cases": [{"id": "demo_case"}]}

    def _run(self, body: str) -> tuple[int, str]:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            golden_dir = root / "ironcalc"
            probe_dir = root / "probes"
            golden_dir.mkdir()
            probe_dir.mkdir()
            (golden_dir / "demo__Sheet1.golden.json").write_text(json.dumps(self.GOLDEN), encoding="utf-8")
            (probe_dir / "demo_probes.golden.json").write_text(json.dumps(self.PROBE), encoding="utf-8")
            path = root / "ironcalc_divergence.yaml"
            path.write_text(body, encoding="utf-8")
            out = io.StringIO()
            err = io.StringIO()
            with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
                status = divergence_check.validate_ironcalc(path, golden_dir=golden_dir, probe_dir=probe_dir)
            return status, out.getvalue() + err.getvalue()

    def _entry(self, **overrides: str) -> str:
        fields = {
            "id": "ironcalc_demo_Sheet1.A1",
            "reason": '"IronCalc caches a stale value. Probe: demo_probes."',
            "prefer": "mac",
            "first_noted": "2026-05-02",
        }
        fields.update(overrides)
        body = "".join(f"    {key}: {value}\n" for key, value in fields.items() if key != "id")
        return f"entries:\n  - id: {fields['id']}\n{body}"

    def test_probe_backed_entry_is_accepted_and_tallied(self) -> None:
        status, output = self._run(self._entry())
        self.assertEqual(status, 0, output)
        self.assertIn("mac-probe: 1", output)
        self.assertIn("2 imported cases in 1 goldens", output)

    def test_orphan_id_fails(self) -> None:
        status, output = self._run(self._entry(id="ironcalc_demo_Sheet1.Z99"))
        self.assertEqual(status, 1)
        self.assertIn("orphan id", output)

    def test_probe_that_names_no_golden_fails(self) -> None:
        status, output = self._run(self._entry(reason='"Probe: gone_probes."'))
        self.assertEqual(status, 1)
        self.assertIn("names no golden", output)

    def test_entry_without_a_probe_needs_a_cause(self) -> None:
        status, output = self._run(self._entry(reason='"Formulon and IronCalc simply differ."'))
        self.assertEqual(status, 1)
        self.assertIn("needs a `cause`", output)

    def test_declared_cause_stands_in_for_a_probe(self) -> None:
        status, output = self._run(
            self._entry(reason='"Single-sheet flatten drops the tab index."', cause="importer-flatten")
        )
        self.assertEqual(status, 0, output)
        self.assertIn("importer-flatten: 1", output)

    def test_unknown_cause_fails(self) -> None:
        status, output = self._run(self._entry(cause="because"))
        self.assertEqual(status, 1)
        self.assertIn("cause must be one of", output)

    def test_explicit_probe_key_is_honoured(self) -> None:
        status, output = self._run(self._entry(reason='"No prose citation here."', probe="demo_probes"))
        self.assertEqual(status, 0, output)
        self.assertIn("mac-probe: 1", output)

    def test_committed_registry_is_clean(self) -> None:
        out = io.StringIO()
        err = io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            status = divergence_check.validate_ironcalc(divergence_check.IRONCALC_DIVERGENCE)
        self.assertEqual(status, 0, out.getvalue() + err.getvalue())


if __name__ == "__main__":
    unittest.main()
