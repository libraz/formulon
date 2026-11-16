from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from tools.oracle import coverage_gap

SUITE = """\
suite: synthetic
cases:
  - id: scalar
    formula: '=FOO(1,2)'
  - id: array_constant
    formula: '=INDEX(FOO({1,2},2),2)'
  - id: range
    setup: { A1: 1, A2: 2 }
    formula: '=FOO(A1:A2,2)'
  - id: error_cell
    setup: { B1: '=NA()' }
    formula: '=FOO(B1,2)'
  - id: blank_cell
    formula: '=FOO(B2,2)'
  - id: omitted
    formula: '=FOO(1,)'
  - id: nested_call_is_not_a_cell
    formula: '=BAR(LOG10(100))'
  - id: error_literal
    formula: '=BAZ(#N/A)'
  - id: text_is_not_a_cell
    formula: '=BAZ("A1")'
  - id: set_cell_is_not_blank
    setup: { C3: 5 }
    formula: '=QUX(C3)'
"""


class ShapeClassifierTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls._tmp = tempfile.TemporaryDirectory()
        Path(cls._tmp.name, "synthetic.yaml").write_text(SUITE, encoding="utf-8")
        cls.shapes = coverage_gap.classify_shapes(Path(cls._tmp.name), ["FOO", "BAR", "LOG10", "BAZ", "QUX", "NONE"])

    @classmethod
    def tearDownClass(cls) -> None:
        cls._tmp.cleanup()

    def test_each_shape_is_tagged(self) -> None:
        self.assertEqual(self.shapes["FOO"], {"calls": 6, "shapes": ["array", "blank", "error", "omitted"]})

    def test_a_call_inside_an_argument_is_not_a_blank_cell(self) -> None:
        self.assertEqual(self.shapes["BAR"]["shapes"], [])
        self.assertEqual(self.shapes["LOG10"], {"calls": 1, "shapes": []})

    def test_error_literal_counts_and_a_string_is_not_a_reference(self) -> None:
        self.assertEqual(self.shapes["BAZ"], {"calls": 2, "shapes": ["error"]})

    def test_a_cell_the_setup_fills_is_not_blank(self) -> None:
        self.assertEqual(self.shapes["QUX"]["shapes"], [])

    def test_gaps_rank_most_missing_first_and_skip_uncalled_and_exempt(self) -> None:
        classified = dict(self.shapes)
        classified["RAND"] = {"calls": 3, "shapes": []}
        gaps = coverage_gap.shape_gaps(classified)
        self.assertEqual(
            [(g["name"], g["missing"]) for g in gaps],
            [
                ("BAR", ["array", "error", "blank"]),
                ("LOG10", ["array", "error", "blank"]),
                ("QUX", ["array", "error", "blank"]),
                ("BAZ", ["array", "blank"]),
            ],
        )


if __name__ == "__main__":
    unittest.main()
