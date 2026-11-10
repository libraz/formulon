"""Unit tests for check_changelog_identifiers.py."""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import check_changelog_identifiers as check  # noqa: E402

CHANGELOG = """# Changelog

## [Unreleased]

- `newName` is new.

## [2.0.0] - 2026-01-01

- `currentName()` replaces `oldName()`, and `Rec.gone` is removed.
- Formula text `=SUM(A1)`, XML `<x:y>`, a path `xl/a.xml`, a glob `fm_*` and `_suffix` are not checked.

| Before | After |
|---|---|
| `oldName` | `currentName` |
| `neverShipped` | `currentName` |

## [1.0.0] - 2025-01-01

- `ignoredBecauseOlder` is not in the current section.
"""


class SpanIdentifierTest(unittest.TestCase):
    def test_only_api_shaped_tokens_are_checked(self) -> None:
        self.assertEqual(list(check.span_identifiers("Workbook.save()")), ["Workbook", "save"])
        self.assertEqual(list(check.span_identifiers("fm_sheet_t")), ["fm_sheet_t"])
        self.assertEqual(list(check.span_identifiers("camelCase")), ["camelCase"])
        self.assertEqual(list(check.span_identifiers("plain")), [])
        self.assertEqual(list(check.span_identifiers("SUM")), [])
        for skipped in ("=SUM(A1)", "<x:y>", "xl/a.xml", "fm_*", "--json", "'#RRGGBB'"):
            self.assertEqual(list(check.span_identifiers(skipped)), [], skipped)


class MissesTest(unittest.TestCase):
    def test_prose_resolves_against_current_and_released_and_before_cells_against_released(self) -> None:
        current = {"newName", "currentName"}
        released = {"oldName", "Rec", "gone"}
        found = [(token, former) for _, token, former in check.misses(CHANGELOG, current, released)]
        self.assertEqual(found, [("neverShipped", True)])

    def test_a_current_name_in_a_before_cell_fails(self) -> None:
        text = CHANGELOG.replace("| `neverShipped` |", "| `currentName` |")
        found = [token for _, token, _ in check.misses(text, {"newName", "currentName"}, {"oldName", "Rec", "gone"})]
        self.assertEqual(found, ["currentName"])

    def test_previous_release_is_the_second_versioned_section(self) -> None:
        self.assertEqual(check.previous_release(CHANGELOG), "1.0.0")


class RepositoryTest(unittest.TestCase):
    def test_the_changelog_names_only_shipped_identifiers(self) -> None:
        self.assertEqual(check.main([]), 0)


if __name__ == "__main__":
    unittest.main()
