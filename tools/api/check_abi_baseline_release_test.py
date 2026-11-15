"""Unit tests for check_abi_baseline_release.py, plus the check over the repository."""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import check_abi_baseline_release as check  # noqa: E402


class BaselineVersionTest(unittest.TestCase):
    def test_reads_the_tag_of_the_first_line(self) -> None:
        self.assertEqual(check.baseline_version("# Released C ABI surface, as shipped in v0.18.0.\n#\n"), "0.18.0")

    def test_a_first_line_without_a_tag_is_an_error(self) -> None:
        with self.assertRaises(ValueError):
            check.baseline_version("# Released C ABI surface.\n")


class NewestReleaseTest(unittest.TestCase):
    def test_compares_versions_numerically(self) -> None:
        self.assertEqual(check.newest_release(["0.9.0", "0.18.0", "0.10.2"]), "0.18.0")


class RepositoryTest(unittest.TestCase):
    def test_the_baseline_is_the_newest_release(self) -> None:
        self.assertEqual(check.main([]), 0)


if __name__ == "__main__":
    unittest.main()
