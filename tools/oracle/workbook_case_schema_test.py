#!/usr/bin/env python3
"""Focused negative tests for the workbook case/golden contract."""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import workbook_case_schema as schema  # noqa: E402


def _case(pivot):
    return {
        "suite": "negative",
        "kind": "workbook",
        "cases": [
            {
                "id": "case",
                "sheets": {"Data": {"A1": {"kind": "text", "value": "x"}}},
                "pivot": pivot,
            }
        ],
    }


class WorkbookCaseSchemaNegativeTest(unittest.TestCase):
    def test_axis_overlap_is_rejected(self):
        with self.assertRaisesRegex(schema.ValidationError, "multiple axes"):
            schema.validate_case_json(
                _case(
                    {
                        "source": "Data!A1:A1",
                        "anchor": "Report!A1",
                        "row_fields": ["Region"],
                        "page_fields": ["Region"],
                    }
                )
            )

    def test_formula_probe_requires_sheet_qualified_a1(self):
        with self.assertRaisesRegex(schema.ValidationError, "sheet-qualified A1"):
            schema.validate_case_json(
                _case(
                    {
                        "source": "Data!A1:A1",
                        "anchor": "Report!A1",
                        "row_fields": ["Region"],
                        "formula_probes": [{"id": "p", "cell": "Z1", "formula": "=1"}],
                    }
                )
            )

    def test_formula_probe_requires_excel_formula(self):
        with self.assertRaisesRegex(schema.ValidationError, "beginning with"):
            schema.validate_case_json(
                _case(
                    {
                        "source": "Data!A1:A1",
                        "anchor": "Report!A1",
                        "row_fields": ["Region"],
                        "formula_probes": [{"id": "p", "cell": "Report!Z1", "formula": "1"}],
                    }
                )
            )

    def test_golden_rejects_unknown_formula_result_kind(self):
        with self.assertRaisesRegex(schema.ValidationError, "unsupported result kind"):
            schema.validate_golden_json(
                {
                    "suite": "negative",
                    "kind": "workbook",
                    "cases": [
                        {
                            "id": "case",
                            "spec": {"pivot": {"formula_probes": [{"id": "p", "cell": "Report!Z1", "formula": "=1"}]}},
                            "expect": {"formula_probes": [{"id": "p", "result": {"kind": "object"}}]},
                        }
                    ],
                }
            )

    def test_full_source_normalization_exposes_case_data_drift(self):
        yaml_like = {
            "suite": "negative",
            "kind": "workbook",
            "description": "Same description",
            "cases": [
                {
                    "id": "case",
                    "description": "same",
                    "sheets": {"Data": {"A1": "wrong"}},
                }
            ],
        }
        json_like = {
            **yaml_like,
            "cases": [
                {
                    **yaml_like["cases"][0],
                    "sheets": {"Data": {"A1": {"kind": "text", "value": "right"}}},
                }
            ],
        }
        self.assertNotEqual(schema.normalise_case_source(yaml_like), schema.normalise_case_source(json_like))


def _layout_case(**extra):
    return {
        "suite": "negative",
        "kind": "workbook",
        "cases": [
            {
                "id": "case",
                "sheets": {"Sheet1": {"A1": {"kind": "text", "value": "x"}}},
                "print": {"sheet": "Sheet1"},
                **extra,
            }
        ],
    }


class HiddenLineSchemaTest(unittest.TestCase):
    """Hiding is its own field, not a zero entry in the size map."""

    def test_hidden_columns_are_accepted(self):
        self.assertEqual(schema.validate_case_json(_layout_case(hidden_columns=["C", "E"])), ["case"])

    def test_hidden_rows_accept_numbers_or_strings(self):
        self.assertEqual(schema.validate_case_json(_layout_case(hidden_rows=[3, "4"])), ["case"])

    def test_empty_hidden_list_is_rejected(self):
        # An empty list reads as "this case is about hiding" while hiding
        # nothing -- almost certainly a half-written case.
        with self.assertRaisesRegex(schema.ValidationError, "non-empty list"):
            schema.validate_case_json(_layout_case(hidden_columns=[]))

    def test_duplicate_hidden_key_is_rejected(self):
        with self.assertRaisesRegex(schema.ValidationError, "duplicate"):
            schema.validate_case_json(_layout_case(hidden_columns=["C", "C"]))

    def test_zero_width_is_still_rejected_by_the_size_map(self):
        # The reason hiding needed its own field: a hidden column is not a
        # zero-width one, and `column_widths` refuses to express either.
        with self.assertRaisesRegex(schema.ValidationError, "must be positive"):
            schema.validate_case_json(_layout_case(column_widths={"C": 0}))


if __name__ == "__main__":
    unittest.main()
