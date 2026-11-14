"""Unit tests for check_upgrade_guide_abi.py, plus the check over the repository."""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import check_upgrade_guide_abi as check  # noqa: E402

LEDGER = """# header
#   removed    gone from the header.
removed fm_old_one :: renamed fm_new_one
signature fm_changed :: takes a record
removed fm_never_published :: existed only before the release was cut
"""


class LedgerTest(unittest.TestCase):
    def test_reads_field_two_and_skips_comments(self) -> None:
        self.assertEqual(check.ledger_names(LEDGER), ["fm_old_one", "fm_changed", "fm_never_published"])


class GuidePathTest(unittest.TestCase):
    def test_guide_follows_the_project_version(self) -> None:
        guide = check.project_guide("project(formulon VERSION 2.3.4 LANGUAGES CXX)")
        self.assertEqual(guide.name, "upgrading-2.3.md")

    def test_no_version_is_an_error(self) -> None:
        with self.assertRaises(ValueError):
            check.project_guide("project(formulon LANGUAGES CXX)")


class MissingNamesTest(unittest.TestCase):
    released = {"fm_old_one", "fm_changed"}

    def test_a_published_name_must_appear_as_a_whole_word(self) -> None:
        guide = "`fm_old_one` is `fm_new_one`; fm_changed_ex is unrelated."
        self.assertEqual(check.missing_names(check.ledger_names(LEDGER), self.released, guide), ["fm_changed"])

    def test_an_unpublished_name_is_not_owed(self) -> None:
        guide = "`fm_old_one` and `fm_changed`."
        self.assertEqual(check.missing_names(check.ledger_names(LEDGER), self.released, guide), [])


class RepositoryTest(unittest.TestCase):
    def test_the_guide_names_every_published_break(self) -> None:
        self.assertEqual(check.main([]), 0)


if __name__ == "__main__":
    unittest.main()
