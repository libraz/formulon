"""Tests for the reprobe path: a skipped case that still owes an observation.

`mode: skip-oracle` says "do not verify this case". It used to also mean
"do not capture it", which quietly made the pending-reprobe queue
unsatisfiable: every entry asking to be re-checked against a live Excel was
excluded from the very run that would have checked it, so a capture on any
Excel reproduced the same "still pending" state. These tests pin the two
halves of the fix -- that a pending stamp selects the case for capture, and
that what Excel says lands beside the skip rather than in place of it.

Run: python3 -m tools.oracle.reprobe_test
"""

from __future__ import annotations

import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from typing import Any, Dict

try:  # pragma: no cover - trivial fallback
    from tools.oracle import divergence_check, oracle_gen, workbook_case_schema
except ImportError:  # pragma: no cover
    sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
    from tools.oracle import divergence_check, oracle_gen, workbook_case_schema

REPO_ROOT = Path(__file__).resolve().parents[2]

_DIVERGENCE = """entries:
  - id: settled_case
    cause: accepted-divergence
    mode: skip-oracle
    reason: "settled"
    prefer: formulon
    last_verified_excel_version: "16.111.2"
  - id: pending_case
    cause: engine-gap
    mode: skip-oracle
    reason: "needs a live look"
    prefer: formulon
    last_verified_excel_version: "unverified (external probe pending)"
  - id: unstamped_case
    cause: engine-gap
    mode: skip-oracle
    reason: "no stamp at all"
    prefer: formulon
  - id: scoped_out_case
    cause: accepted-divergence
    mode: skip-oracle
    applies_to: [win-365-ja_JP]
    reason: "other target"
    prefer: formulon
    last_verified_excel_version: "unverified (pending)"
  - id: tolerance_case
    cause: accepted-divergence
    mode: tolerance
    reason: "not a skip at all"
    prefer: formulon
    last_verified_excel_version: "unverified (pending)"
"""


class ReprobeSelectionTests(unittest.TestCase):
    def setUp(self) -> None:
        tmp = Path(tempfile.mkdtemp())
        self.addCleanup(lambda: shutil.rmtree(tmp, ignore_errors=True))
        self.path = tmp / "divergence.yaml"
        self.path.write_text(_DIVERGENCE, encoding="utf-8")

    def _reprobes(self, target: str = "mac-365-ja_JP") -> Dict[str, str]:
        return oracle_gen._load_divergence_reprobes(self.path, target)

    def test_a_pending_stamp_selects_the_case(self) -> None:
        self.assertIn("pending_case", self._reprobes())

    def test_a_missing_stamp_selects_the_case(self) -> None:
        # An entry that never recorded a version has, if anything, weaker
        # evidence than one that recorded "unverified".
        self.assertIn("unstamped_case", self._reprobes())

    def test_a_verified_stamp_is_left_alone(self) -> None:
        # Reprobing settled entries would turn every capture into a full
        # re-run of the skip list for no new information.
        self.assertNotIn("settled_case", self._reprobes())

    def test_only_skip_oracle_entries_are_reprobed(self) -> None:
        # A `tolerance` entry is verified normally; it is never skipped, so
        # there is nothing to reprobe.
        self.assertNotIn("tolerance_case", self._reprobes())

    def test_applies_to_scoping_is_honoured(self) -> None:
        self.assertNotIn("scoped_out_case", self._reprobes("mac-365-ja_JP"))
        self.assertIn("scoped_out_case", self._reprobes("win-365-ja_JP"))

    def test_reprobes_are_a_subset_of_skips(self) -> None:
        # The reprobe set only ever narrows the skip set -- a case that is
        # verified normally must never be diverted into `observed`.
        skips = set(oracle_gen._load_divergence_skips(self.path, "win-365-ja_JP"))
        self.assertTrue(set(self._reprobes("win-365-ja_JP")).issubset(skips))

    def test_repository_pending_queue_is_selected(self) -> None:
        # Guards the actual deadlock: the eight entries the Windows track
        # owes an observation must be the ones a Windows capture drives.
        reprobes = oracle_gen._load_divergence_reprobes(REPO_ROOT / "tests/divergence.yaml", "win-365-ja_JP")
        self.assertIn("page_axis_field", reprobes)
        self.assertIn("scale_50_shrinks_breaks", reprobes)


def _golden(case: Dict[str, Any]) -> Dict[str, Any]:
    return {"suite": "s", "kind": "workbook", "cases": [case]}


class ObservedSchemaTests(unittest.TestCase):
    """`observed` is evidence, never an assertion; the schema says so."""

    SPEC = {"id": "c", "print": {"sheet": "Sheet1"}}

    def test_observed_beside_a_skip_is_accepted(self) -> None:
        doc = _golden({"id": "c", "spec": self.SPEC, "skipped": "why", "observed": {"print": {"page_count": 2}}})
        self.assertEqual(workbook_case_schema.validate_golden_json(doc), ["c"])

    def test_observed_error_beside_a_skip_is_accepted(self) -> None:
        doc = _golden({"id": "c", "spec": self.SPEC, "skipped": "why", "observed_error": "ComError: boom"})
        self.assertEqual(workbook_case_schema.validate_golden_json(doc), ["c"])

    def test_observed_without_a_skip_is_rejected(self) -> None:
        doc = _golden({"id": "c", "spec": self.SPEC, "expect": {}, "observed": {"print": {}}})
        with self.assertRaisesRegex(workbook_case_schema.ValidationError, "only valid on a skipped case"):
            workbook_case_schema.validate_golden_json(doc)

    def test_observed_alongside_expect_is_rejected(self) -> None:
        # If a case is being asserted, it is not pending adjudication --
        # allowing both would leave two competing records of Excel.
        doc = _golden({"id": "c", "spec": self.SPEC, "expect": {}, "skipped": "why", "observed": {"print": {}}})
        with self.assertRaisesRegex(workbook_case_schema.ValidationError, "not pending adjudication"):
            workbook_case_schema.validate_golden_json(doc)

    def test_observed_and_observed_error_are_mutually_exclusive(self) -> None:
        doc = _golden(
            {"id": "c", "spec": self.SPEC, "skipped": "why", "observed": {"print": {}}, "observed_error": "boom"}
        )
        with self.assertRaisesRegex(workbook_case_schema.ValidationError, "mutually exclusive"):
            workbook_case_schema.validate_golden_json(doc)

    def test_observed_must_be_a_mapping(self) -> None:
        doc = _golden({"id": "c", "spec": self.SPEC, "skipped": "why", "observed": "2 pages"})
        with self.assertRaisesRegex(workbook_case_schema.ValidationError, "expected mapping"):
            workbook_case_schema.validate_golden_json(doc)


class ObservationReportingTests(unittest.TestCase):
    """An observation nobody is told about is the same as no observation."""

    def _golden_dir(self, case: Dict[str, Any]) -> Path:
        tmp = Path(tempfile.mkdtemp())
        self.addCleanup(lambda: shutil.rmtree(tmp, ignore_errors=True))
        (tmp / "s.golden.json").write_text(json.dumps(_golden(case), indent=2) + "\n", encoding="utf-8")
        return tmp

    def test_an_observation_is_reported_against_its_case(self) -> None:
        found = divergence_check.load_observations(
            [self._golden_dir({"id": "pending_case", "spec": {}, "skipped": "w", "observed": {"print": {}}})]
        )
        self.assertIn("pending_case", found)

    def test_a_failed_probe_is_reported_rather_than_swallowed(self) -> None:
        found = divergence_check.load_observations(
            [self._golden_dir({"id": "pending_case", "spec": {}, "skipped": "w", "observed_error": "ComError: boom"})]
        )
        self.assertIn("ComError: boom", found["pending_case"])

    def test_a_plain_skip_reports_nothing(self) -> None:
        found = divergence_check.load_observations([self._golden_dir({"id": "c", "spec": {}, "skipped": "w"})])
        self.assertEqual(found, {})


if __name__ == "__main__":
    unittest.main()
