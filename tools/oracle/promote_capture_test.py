"""Tests for `promote_capture`, the staged-capture landing step.

The point of these is the refusal side. Promotion is what decides whether
an externally captured golden earns repository coverage, and the whole
reason the step exists is that a previous Windows capture entered the tree
without anyone checking which Excel produced it. Each test below feeds the
promoter a capture that is wrong in exactly one way and asserts it is
turned away.

Run: python3 -m tools.oracle.promote_capture_test
"""

from __future__ import annotations

import hashlib
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock
from typing import Any, Dict

try:  # pragma: no cover - trivial fallback
    from tools.oracle import promote_capture
except ImportError:  # pragma: no cover
    sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
    from tools.oracle import promote_capture

_SENTINEL = "ARRAYTOTEXT(1) == text '1'"
# A Windows capture stamps `Application.Version` + `Application.Build`; the
# build is the only part that distinguishes an Office SKU.
_PRODUCT = "16.0.18025"
_LOCALE = "ja-JP"
_CAPTURE = "capture-0001"
_TARGET = "win-365-ja_JP"
_RECORD = {"locale": _LOCALE}


def _golden_doc(suite: str) -> Dict[str, Any]:
    return {
        "suite": suite,
        "kind": "workbook",
        "environment": {
            "excel_version": _PRODUCT,
            "excel_locale": _LOCALE,
            "date1904": False,
            "iterative": False,
            "generated_at": "2026-08-16T00:00:00Z",
            "capture_id": _CAPTURE,
        },
        "cases": [{"id": f"{suite}_case", "spec": {}, "expect": {}}],
    }


def _write_capture(root: Path, suites=("pivot_basic", "print_basic"), **candidate_overrides: Any) -> Path:
    """Materialises a complete, valid staged capture, then applies overrides."""

    root.mkdir(parents=True, exist_ok=True)
    inventory = []
    for suite in suites:
        path = root / f"{suite}.golden.json"
        path.write_text(json.dumps(_golden_doc(suite), indent=2) + "\n", encoding="utf-8")
        inventory.append(
            {
                "suite": suite,
                "case_count": 1,
                "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
                "capture_id": _CAPTURE,
            }
        )
    candidate: Dict[str, Any] = {
        "target": _TARGET,
        "status": "wanted",
        "classification": "candidate",
        "verified": True,
        "active_ctest": False,
        "capture_id": _CAPTURE,
        "all_suites_same_capture": True,
        "product": _PRODUCT,
        "locale": _LOCALE,
        "m365_sentinel": _SENTINEL,
        "required_suites": sorted(suites),
        "captured_suites": sorted(suites),
        "suite_inventory": inventory,
    }
    candidate.update(candidate_overrides)
    (root / "PROVENANCE.candidate.json").write_text(json.dumps(candidate, indent=2) + "\n", encoding="utf-8")
    return root


class CheckCaptureTests(unittest.TestCase):
    def _staged(self, **overrides: Any) -> Path:
        tmp = Path(tempfile.mkdtemp())
        self.addCleanup(lambda: __import__("shutil").rmtree(tmp, ignore_errors=True))
        return _write_capture(tmp / "capture", **overrides)

    def test_complete_capture_is_accepted(self) -> None:
        candidate = promote_capture.check_capture(self._staged(), _TARGET, _RECORD)
        self.assertEqual(candidate["product"], _PRODUCT)
        self.assertEqual(sorted(candidate["captured_suites"]), ["pivot_basic", "print_basic"])

    def test_missing_directory_is_reported_not_crashed(self) -> None:
        with self.assertRaisesRegex(RuntimeError, "no staged capture"):
            promote_capture.check_capture(Path("/nonexistent/formulon-capture"), _TARGET, _RECORD)

    def test_capture_without_a_candidate_marker_is_refused(self) -> None:
        # The generator emits the marker only after a complete run on an
        # M365 host, so its absence is the signal that one of those did not
        # hold -- a partial capture must never be promotable.
        staged = self._staged()
        (staged / "PROVENANCE.candidate.json").unlink()
        with self.assertRaisesRegex(RuntimeError, "PROVENANCE.candidate.json"):
            promote_capture.check_capture(staged, _TARGET, _RECORD)

    def test_unknown_product_is_refused(self) -> None:
        # This is the exact shape of the historical Office 2019 mix-up.
        staged = self._staged(product="unknown Excel 16.0 capture")
        with self.assertRaisesRegex(RuntimeError, "Office 2019"):
            promote_capture.check_capture(staged, _TARGET, _RECORD)

    def test_bare_office_major_product_is_refused(self) -> None:
        # `Application.Version` reports "16.0" for every SKU from Office
        # 2016 through Microsoft 365, so it identifies nothing. A capture
        # stamped with it is indistinguishable after the fact from the
        # Office 2019 one, which is the whole failure this step exists for.
        staged = self._staged(product="16.0")
        for path in staged.glob("*.golden.json"):
            document = json.loads(path.read_text())
            document["environment"]["excel_version"] = "16.0"
            path.write_text(json.dumps(document, indent=2) + "\n", encoding="utf-8")
        with self.assertRaisesRegex(RuntimeError, "names no particular Excel build"):
            promote_capture.check_capture(staged, _TARGET, _RECORD)

    def test_prose_wrapped_product_is_refused(self) -> None:
        staged = self._staged(product="Excel 365 (Windows, ja-JP)")
        with self.assertRaisesRegex(RuntimeError, "names no particular Excel build"):
            promote_capture.check_capture(staged, _TARGET, _RECORD)

    def test_missing_m365_sentinel_is_refused(self) -> None:
        staged = self._staged(m365_sentinel="")
        with self.assertRaisesRegex(RuntimeError, "sentinel"):
            promote_capture.check_capture(staged, _TARGET, _RECORD)

    def test_locale_disagreeing_with_the_target_is_refused(self) -> None:
        staged = self._staged(locale="en-US")
        with self.assertRaisesRegex(RuntimeError, "locale"):
            promote_capture.check_capture(staged, _TARGET, _RECORD)

    def test_capture_for_another_target_is_refused(self) -> None:
        staged = self._staged(target="win-365-de_DE")
        with self.assertRaisesRegex(RuntimeError, "target"):
            promote_capture.check_capture(staged, _TARGET, _RECORD)

    def test_partial_suite_capture_is_refused(self) -> None:
        staged = self._staged()
        candidate = json.loads((staged / "PROVENANCE.candidate.json").read_text())
        candidate["captured_suites"] = ["pivot_basic"]
        (staged / "PROVENANCE.candidate.json").write_text(json.dumps(candidate), encoding="utf-8")
        with self.assertRaisesRegex(RuntimeError, "missing suite"):
            promote_capture.check_capture(staged, _TARGET, _RECORD)

    def test_golden_edited_after_capture_is_refused(self) -> None:
        # The inventory is the capture's own account of itself; re-deriving
        # the digest is what stops a hand-edited value from being promoted
        # under a genuine capture's provenance.
        staged = self._staged()
        path = staged / "pivot_basic.golden.json"
        document = json.loads(path.read_text())
        document["cases"][0]["expect"] = {"tampered": True}
        path.write_text(json.dumps(document, indent=2) + "\n", encoding="utf-8")
        with self.assertRaisesRegex(RuntimeError, "SHA-256"):
            promote_capture.check_capture(staged, _TARGET, _RECORD)

    def test_golden_from_a_different_capture_is_refused(self) -> None:
        # Two Excel sessions must not be spliced into one provenance record.
        staged = self._staged()
        path = staged / "print_basic.golden.json"
        document = json.loads(path.read_text())
        document["environment"]["capture_id"] = "capture-0002"
        path.write_text(json.dumps(document, indent=2) + "\n", encoding="utf-8")
        candidate = json.loads((staged / "PROVENANCE.candidate.json").read_text())
        for item in candidate["suite_inventory"]:
            if item["suite"] == "print_basic":
                item["sha256"] = hashlib.sha256(path.read_bytes()).hexdigest()
        (staged / "PROVENANCE.candidate.json").write_text(json.dumps(candidate), encoding="utf-8")
        with self.assertRaisesRegex(RuntimeError, "different capture"):
            promote_capture.check_capture(staged, _TARGET, _RECORD)

    def test_inventory_naming_an_absent_golden_is_refused(self) -> None:
        staged = self._staged()
        (staged / "print_basic.golden.json").unlink()
        with self.assertRaisesRegex(RuntimeError, "absent"):
            promote_capture.check_capture(staged, _TARGET, _RECORD)


class PromotedProvenanceTests(unittest.TestCase):
    def test_promotion_marks_the_capture_active(self) -> None:
        candidate = {
            "classification": "candidate",
            "verified": True,
            "active_ctest": False,
            "status": "wanted",
            "reason": "staged",
            "product": _PRODUCT,
        }
        promoted = promote_capture._promoted_provenance(candidate)
        self.assertEqual(promoted["classification"], "active")
        self.assertIs(promoted["verified"], True)
        self.assertIs(promoted["active_ctest"], True)
        self.assertEqual(promoted["status"], "scaffolded")
        # A staging note must not survive into the repository marker.
        self.assertNotIn("reason", promoted)
        # The candidate itself is not mutated.
        self.assertEqual(candidate["classification"], "candidate")


class TargetStatusRewriteTests(unittest.TestCase):
    MANIFEST = """primary: mac-365-ja_JP

targets:
  mac-365-ja_JP:
    status: primary
    driver: macos_excel

  win-365-ja_JP:
    # A comment that must survive the rewrite.
    status: wanted
    driver: windows_excel
    locale: ja-JP

  win-365-en_US:
    status: wanted
    driver: windows_excel
"""

    def _manifest(self) -> Path:
        tmp = Path(tempfile.mkdtemp())
        self.addCleanup(lambda: __import__("shutil").rmtree(tmp, ignore_errors=True))
        path = tmp / "targets.yaml"
        path.write_text(self.MANIFEST, encoding="utf-8")
        return path

    def test_rewrites_only_the_named_target(self) -> None:
        path = self._manifest()
        self.assertTrue(promote_capture._set_target_status(path, "win-365-ja_JP", "scaffolded"))
        text = path.read_text(encoding="utf-8")
        self.assertIn(
            "  win-365-ja_JP:\n    # A comment that must survive the rewrite.\n    status: scaffolded\n", text
        )
        # The sibling target keeps its own status.
        self.assertIn("  win-365-en_US:\n    status: wanted\n", text)
        self.assertIn("    status: primary\n", text)

    def test_rewrite_is_idempotent(self) -> None:
        path = self._manifest()
        self.assertTrue(promote_capture._set_target_status(path, "win-365-ja_JP", "scaffolded"))
        before = path.read_text(encoding="utf-8")
        self.assertFalse(promote_capture._set_target_status(path, "win-365-ja_JP", "scaffolded"))
        self.assertEqual(path.read_text(encoding="utf-8"), before)

    def test_unknown_target_raises(self) -> None:
        with self.assertRaisesRegex(RuntimeError, "no `status:` line"):
            promote_capture._set_target_status(self._manifest(), "win-365-zz_ZZ", "scaffolded")


class ScaffoldedInPlacePromotionTests(unittest.TestCase):
    """Promotion of a target that no longer stages outside the tree.

    Only a `wanted` target stages into the cache; once it is scaffolded the
    generator writes its goldens and its candidate straight into the golden
    directory, so promotion is the review step alone. Defaulting to the
    staging path there would land whichever capture the cache still holds
    from the run that scaffolded the target, and copying the goldens onto
    themselves raises SameFileError.
    """

    MANIFEST = """tracks:
  workbook:
    primary: win-365-ja_JP

targets:
  win-365-ja_JP:
    status: scaffolded
    driver: windows_excel
    locale: ja-JP
"""

    def _fixture(self) -> "tuple[Path, Path]":
        tmp = Path(tempfile.mkdtemp())
        self.addCleanup(lambda: __import__("shutil").rmtree(tmp, ignore_errors=True))
        golden_dir = _write_capture(tmp / "golden_wb", status="scaffolded")
        targets = tmp / "targets.yaml"
        targets.write_text(self.MANIFEST, encoding="utf-8")
        return golden_dir, targets

    def _promote(self, golden_dir: Path, targets: Path, extra: "list[str]") -> int:
        with (
            mock.patch.dict(promote_capture.TRACKS["workbook"], {"default_golden_dir": golden_dir}),
            mock.patch.object(promote_capture, "REPO_ROOT", golden_dir.parent),
        ):
            return promote_capture.main(["--targets-file", str(targets), *extra])

    def test_capture_already_in_the_golden_dir_is_promoted_in_place(self) -> None:
        golden_dir, targets = self._fixture()
        goldens_before = {p.name: p.read_bytes() for p in golden_dir.glob("*.golden.json")}

        self.assertEqual(self._promote(golden_dir, targets, []), 0)

        promoted = json.loads((golden_dir / "PROVENANCE.json").read_text(encoding="utf-8"))
        self.assertEqual(promoted["capture_id"], _CAPTURE)
        self.assertEqual(promoted["product"], _PRODUCT)
        # The goldens are the capture's own; promotion must not touch them.
        self.assertEqual({p.name: p.read_bytes() for p in golden_dir.glob("*.golden.json")}, goldens_before)

    def test_explicit_from_pointing_at_the_golden_dir_is_accepted(self) -> None:
        golden_dir, targets = self._fixture()

        self.assertEqual(self._promote(golden_dir, targets, ["--from", str(golden_dir)]), 0)

        self.assertTrue((golden_dir / "PROVENANCE.json").is_file())

    def test_a_stale_staged_capture_is_not_promoted_for_a_scaffolded_target(self) -> None:
        golden_dir, targets = self._fixture()
        stale = _write_capture(golden_dir.parent / "stale", capture_id="stale-0000")
        with mock.patch.object(promote_capture, "staging_dir_for", return_value=stale) as staging:
            self.assertEqual(self._promote(golden_dir, targets, []), 0)

        staging.assert_not_called()
        promoted = json.loads((golden_dir / "PROVENANCE.json").read_text(encoding="utf-8"))
        self.assertEqual(promoted["capture_id"], _CAPTURE)


if __name__ == "__main__":
    unittest.main()
