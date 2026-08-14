#!/usr/bin/env python3
"""Regression tests for the tree-level `ENVIRONMENT.md` stamp.

`ENVIRONMENT.md` documents itself as covering every suite under
`tests/oracle/golden/`, not just whichever suite(s) a given
`oracle-gen --suite ...` invocation happened to touch. These tests pin
the helper that derives the tree-wide stamp from the committed goldens
themselves, so a partial regeneration can never advance the file past
the oldest suite still on disk.
"""

from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from tools.oracle import oracle_gen


class VersionTupleTest(unittest.TestCase):
    def test_parses_plain_dotted_versions(self) -> None:
        self.assertEqual(oracle_gen._version_tuple("16.111.2"), (16, 111, 2))
        self.assertEqual(oracle_gen._version_tuple("16.112"), (16, 112))

    def test_parses_build_suffixed_versions(self) -> None:
        self.assertEqual(oracle_gen._version_tuple("16.84 (Build 24021522)"), (16, 84))

    def test_non_numeric_stamp_sorts_as_zero(self) -> None:
        # A hand-seeded or otherwise non-Microsoft-365 stamp must not be
        # mistaken for a real (and therefore "oldest wins") build.
        self.assertEqual(oracle_gen._version_tuple("hand-seeded"), (0,))
        self.assertEqual(oracle_gen._version_tuple(""), (0,))


class TreeWideExcelVersionTest(unittest.TestCase):
    def _write_golden(self, golden_dir: Path, name: str, excel_version: str) -> None:
        (golden_dir / f"{name}.golden.json").write_text(
            json.dumps({"environment": {"excel_version": excel_version}}),
            encoding="utf-8",
        )

    def test_picks_the_oldest_committed_suite_version(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            golden_dir = Path(raw)
            self._write_golden(golden_dir, "aggregate", "16.112")
            self._write_golden(golden_dir, "logical", "16.108.1")
            self._write_golden(golden_dir, "math", "16.111.2")

            # This is exactly the M-35 scenario: a `--suite aggregate`
            # regeneration lands 16.112 for one suite while `logical`
            # is still committed at 16.108.1. The tree-wide stamp must
            # follow the older suite, not the suite that was just run.
            version = oracle_gen._tree_wide_excel_version(golden_dir, fallback="99.99")
            self.assertEqual(version, "16.108.1")

    def test_ignores_unparseable_stamps_like_hand_seeded(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            golden_dir = Path(raw)
            self._write_golden(golden_dir, "aggregate", "16.112")
            self._write_golden(golden_dir, "bootstrap", "hand-seeded")

            version = oracle_gen._tree_wide_excel_version(golden_dir, fallback="99.99")
            self.assertEqual(version, "16.112")

    def test_falls_back_when_no_golden_has_a_parseable_version(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            golden_dir = Path(raw)
            self._write_golden(golden_dir, "bootstrap", "hand-seeded")

            version = oracle_gen._tree_wide_excel_version(golden_dir, fallback="16.112")
            self.assertEqual(version, "16.112")

    def test_empty_directory_falls_back(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            golden_dir = Path(raw)
            version = oracle_gen._tree_wide_excel_version(golden_dir, fallback="16.112")
            self.assertEqual(version, "16.112")


class WriteEnvironmentMdTest(unittest.TestCase):
    def test_records_the_supplied_excel_version_not_the_probed_one(self) -> None:
        # `_write_environment_md` must record whatever `excel_version`
        # it is handed -- callers are expected to have already resolved
        # that to the tree-wide minimum, which may be older than the
        # Excel build that ran this particular invocation.
        from tools.oracle.drivers.base import EnvironmentInfo

        env = EnvironmentInfo(
            excel_version="16.112",
            excel_locale="ja-JP",
            date1904=False,
            iterative=False,
        )
        with tempfile.TemporaryDirectory() as raw:
            path = Path(raw) / "ENVIRONMENT.md"
            oracle_gen._write_environment_md(path, env, "2026-08-14T00:00:00Z", excel_version="16.108.1")
            body = path.read_text(encoding="utf-8")
            self.assertIn("`16.108.1`", body)
            self.assertNotIn("`16.112`", body)


if __name__ == "__main__":
    unittest.main()
