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


def _roundtrip_case(block):
    return {
        "suite": "negative",
        "kind": "workbook",
        "cases": [
            {
                "id": "case",
                "sheets": {"Sheet1": {"A1": {"kind": "text", "value": "x"}}},
                "roundtrip": block,
            }
        ],
    }


def _roundtrip_observed(**overrides):
    observed = {
        "page_setup": {
            "paper_size": 9,
            "orientation": 1,
            "zoom": 100,
            "fit_to_pages_wide": False,
            "fit_to_pages_tall": False,
        },
        "page_margins": {k: 0.5 for k in ("left", "right", "top", "bottom", "header", "footer")},
        "print_options": {
            "grid_lines": False,
            "headings": False,
            "horizontal_centered": False,
            "vertical_centered": False,
        },
        "header_footer": {},
        "print_area": "A1:D20",
        "print_title_rows": "",
        "print_title_cols": "",
        "manual_row_breaks": [],
        "manual_col_breaks": [],
        "xlsx_sha256": "0" * 64,
    }
    observed.update(overrides)
    return observed


def _roundtrip_golden(observed):
    return {
        "suite": "negative",
        "kind": "workbook",
        "cases": [
            {
                "id": "case",
                "spec": {"roundtrip": {"sheet": "Sheet1"}},
                "expect": {"roundtrip": observed},
            }
        ],
    }


class RoundtripSchemaTest(unittest.TestCase):
    """The roundtrip block is the only case shape whose fixture is authored
    by us rather than by Excel, so a field that quietly does nothing would
    produce a golden of Excel resolving defaults -- a clean-looking pass
    that tested nothing."""

    def test_minimal_block_is_accepted(self):
        schema.validate_case_json(_roundtrip_case({"sheet": "Sheet1"}))

    def test_sheet_is_required(self):
        with self.assertRaisesRegex(schema.ValidationError, "missing required 'sheet'"):
            schema.validate_case_json(_roundtrip_case({"page_setup": {"scale": 75}}))

    def test_misspelled_sub_block_is_rejected(self):
        with self.assertRaisesRegex(schema.ValidationError, "unknown roundtrip field"):
            schema.validate_case_json(_roundtrip_case({"sheet": "Sheet1", "page_setups": {}}))

    def test_misspelled_field_inside_a_sub_block_is_rejected(self):
        with self.assertRaisesRegex(schema.ValidationError, "unknown field"):
            schema.validate_case_json(_roundtrip_case({"sheet": "Sheet1", "page_setup": {"scal": 75}}))

    def test_fit_to_page_must_be_a_boolean(self):
        with self.assertRaisesRegex(schema.ValidationError, "expected a boolean"):
            schema.validate_case_json(_roundtrip_case({"sheet": "Sheet1", "page_setup": {"fit_to_page": 1}}))

    def test_row_breaks_are_one_based(self):
        with self.assertRaisesRegex(schema.ValidationError, "1-based row number"):
            schema.validate_case_json(_roundtrip_case({"sheet": "Sheet1", "row_breaks": [0]}))

    def test_col_breaks_are_column_letters(self):
        with self.assertRaisesRegex(schema.ValidationError, "column letters"):
            schema.validate_case_json(_roundtrip_case({"sheet": "Sheet1", "col_breaks": [4]}))

    def test_empty_header_footer_section_is_allowed(self):
        # An empty string clears the section, which is a thing a case may
        # legitimately want to pin.
        schema.validate_case_json(_roundtrip_case({"sheet": "Sheet1", "header_footer": {"odd_header": ""}}))

    def test_golden_must_report_a_roundtrip_block(self):
        golden = _roundtrip_golden(_roundtrip_observed())
        golden["cases"][0]["expect"] = {}
        with self.assertRaisesRegex(schema.ValidationError, "must report a roundtrip block"):
            schema.validate_golden_json(golden)

    def test_golden_rejects_a_missing_sub_block_field(self):
        observed = _roundtrip_observed()
        del observed["page_setup"]["zoom"]
        with self.assertRaisesRegex(schema.ValidationError, "missing field"):
            schema.validate_golden_json(_roundtrip_golden(observed))

    def test_golden_requires_the_fixture_digest(self):
        # Without it a stale golden cannot be told from a current one.
        observed = _roundtrip_observed()
        del observed["xlsx_sha256"]
        with self.assertRaisesRegex(schema.ValidationError, "sha256"):
            schema.validate_golden_json(_roundtrip_golden(observed))

    def test_golden_break_indices_are_zero_based(self):
        with self.assertRaisesRegex(schema.ValidationError, "zero-based index"):
            schema.validate_golden_json(_roundtrip_golden(_roundtrip_observed(manual_row_breaks=[-1])))

    def test_complete_golden_is_accepted(self):
        self.assertEqual(schema.validate_golden_json(_roundtrip_golden(_roundtrip_observed())), ["case"])


if __name__ == "__main__":
    unittest.main()
