"""Empty-value rules of the conditional-format rule input (api-conventions §5 rule 30).

``""`` is refused for the id, the two formulas and a threshold value, and is a value for the text of a text rule, as
Excel stores it.
"""

from __future__ import annotations

import io
import re
import unittest
import zipfile

from formulon import (
    CellRect,
    CfCellIsRuleInput,
    CfColorScaleInput,
    CfColorScaleRuleInput,
    CfDataBarInput,
    CfDataBarRuleInput,
    CfExpressionRuleInput,
    CfRuleType,
    CfTextRuleInput,
    CfValueObject,
    CfValueObjectType,
    Color,
    ColorKind,
    ComparisonOperator,
    InvalidArgumentError,
    Workbook,
    WorkbookFormat,
)

GUID = re.compile(r"^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$")
RANGE = [CellRect(0, 0, 3, 0)]


def _rgb(red: int, green: int, blue: int) -> Color:
    return Color(ColorKind.RGB, argb=0xFF000000 | (red << 16) | (green << 8) | blue)


class _Base(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.addCleanup(self.wb.close)
        self.ws = self.wb.sheet(0)


class IdAndFormulaTests(_Base):
    def test_empty_id_is_refused_and_none_generates(self) -> None:
        with self.assertRaises(InvalidArgumentError):
            self.ws.add_cf_rule(CfExpressionRuleInput(ranges=RANGE, formula1="TRUE", id=""))
        self.assertEqual(len(self.ws.list_cf_rules()), 0)
        rule_id = self.ws.add_cf_rule(CfExpressionRuleInput(ranges=RANGE, formula1="TRUE"))
        self.assertRegex(rule_id, GUID)
        self.assertEqual(len(self.ws.list_cf_rules()), 1)

    def test_empty_formula1_is_refused_and_none_fills_a_text_rule(self) -> None:
        with self.assertRaises(InvalidArgumentError):
            self.ws.add_cf_rule(CfExpressionRuleInput(ranges=RANGE, formula1=""))
        with self.assertRaises(InvalidArgumentError):
            self.ws.add_cf_rule(CfCellIsRuleInput(ranges=RANGE, operator=ComparisonOperator.GREATER_THAN, formula1=""))
        with self.assertRaises(InvalidArgumentError):
            self.ws.add_cf_rule(CfTextRuleInput(ranges=RANGE, type=CfRuleType.CONTAINS_TEXT, text="ab", formula1=""))
        self.assertEqual(len(self.ws.list_cf_rules()), 0)
        rule_id = self.ws.add_cf_rule(CfTextRuleInput(ranges=RANGE, type=CfRuleType.CONTAINS_TEXT, text="ab"))
        self.assertEqual(self.ws.get_cf_rule(rule_id).formula1, '=NOT(ISERROR(SEARCH("ab",A1)))')

    def test_empty_formula2_is_refused_and_none_is_absent(self) -> None:
        with self.assertRaises(InvalidArgumentError):
            self.ws.add_cf_rule(
                CfCellIsRuleInput(ranges=RANGE, operator=ComparisonOperator.BETWEEN, formula1="1", formula2="")
            )
        self.assertEqual(len(self.ws.list_cf_rules()), 0)
        rule_id = self.ws.add_cf_rule(
            CfCellIsRuleInput(ranges=RANGE, operator=ComparisonOperator.GREATER_THAN, formula1="1", formula2=None)
        )
        self.assertIsNone(self.ws.get_cf_rule(rule_id).formula2)

    def test_update_refuses_an_empty_formula_and_changes_nothing(self) -> None:
        rule_id = self.ws.add_cf_rule(
            CfCellIsRuleInput(ranges=RANGE, operator=ComparisonOperator.BETWEEN, formula1="1", formula2="9")
        )
        with self.assertRaises(InvalidArgumentError):
            self.ws.update_cf_rule(rule_id, formula1="")
        with self.assertRaises(InvalidArgumentError):
            self.ws.update_cf_rule(rule_id, formula2="")
        rule = self.ws.get_cf_rule(rule_id)
        self.assertEqual(rule.formula1, "=1")
        self.assertEqual(rule.formula2, "=9")


class ThresholdValueTests(_Base):
    def test_empty_threshold_value_is_refused_and_none_reads_none(self) -> None:
        def color_scale(value: object) -> CfColorScaleRuleInput:
            return CfColorScaleRuleInput(
                ranges=RANGE,
                color_scale=CfColorScaleInput(
                    [CfValueObject(CfValueObjectType.MIN), CfValueObject(CfValueObjectType.NUMBER, value)],
                    [_rgb(255, 0, 0), _rgb(0, 255, 0)],
                ),
            )

        with self.assertRaises(InvalidArgumentError):
            self.ws.add_cf_rule(color_scale(""))
        with self.assertRaises(InvalidArgumentError):
            self.ws.add_cf_rule(
                CfDataBarRuleInput(
                    ranges=RANGE,
                    data_bar=CfDataBarInput(
                        CfValueObject(CfValueObjectType.MIN),
                        CfValueObject(CfValueObjectType.NUMBER, ""),
                        _rgb(0, 0, 255),
                    ),
                )
            )
        self.assertEqual(len(self.ws.list_cf_rules()), 0)
        rule_id = self.ws.add_cf_rule(color_scale("10"))
        self.assertIsNone(self.ws.get_cf_rule(rule_id).color_scale.thresholds[0].value)

    def test_color_scale_auto_threshold_is_refused_and_changes_nothing(self) -> None:
        def color_scale(first: CfValueObjectType, second: CfValueObjectType) -> CfColorScaleInput:
            return CfColorScaleInput([CfValueObject(first), CfValueObject(second)], [_rgb(255, 0, 0), _rgb(0, 255, 0)])

        with self.assertRaises(InvalidArgumentError):
            self.ws.add_cf_rule(
                CfColorScaleRuleInput(
                    ranges=RANGE, color_scale=color_scale(CfValueObjectType.AUTO_MIN, CfValueObjectType.MAX)
                )
            )
        self.assertEqual(len(self.ws.list_cf_rules()), 0)
        rule_id = self.ws.add_cf_rule(
            CfColorScaleRuleInput(ranges=RANGE, color_scale=color_scale(CfValueObjectType.MIN, CfValueObjectType.MAX))
        )
        with self.assertRaises(InvalidArgumentError):
            self.ws.update_cf_rule(rule_id, color_scale=color_scale(CfValueObjectType.MIN, CfValueObjectType.AUTO_MAX))
        thresholds = self.ws.get_cf_rule(rule_id).color_scale.thresholds
        self.assertEqual([t.type for t in thresholds], [CfValueObjectType.MIN, CfValueObjectType.MAX])


class TextTests(_Base):
    def _text_rule(self, **kwargs: object) -> CfTextRuleInput:
        return CfTextRuleInput(ranges=RANGE, type=CfRuleType.CONTAINS_TEXT, **kwargs)  # type: ignore[arg-type]

    def test_empty_text_is_a_value_and_none_is_absent(self) -> None:
        empty = self.ws.add_cf_rule(self._text_rule(text=""))
        self.assertEqual(self.ws.get_cf_rule(empty).text, "")
        self.assertEqual(self.ws.get_cf_rule(empty).formula1, '=NOT(ISERROR(SEARCH("",A1)))')
        absent = self.ws.add_cf_rule(self._text_rule(text=None))
        self.assertIsNone(self.ws.get_cf_rule(absent).text)

    def test_update_to_empty_text_is_kept(self) -> None:
        rule_id = self.ws.add_cf_rule(CfTextRuleInput(ranges=RANGE, type=CfRuleType.BEGINS_WITH, text="ab"))
        self.ws.update_cf_rule(rule_id, text="")
        self.assertEqual(self.ws.get_cf_rule(rule_id).text, "")

    def test_empty_text_is_saved_as_an_attribute_and_round_trips(self) -> None:
        self.ws.add_cf_rule(self._text_rule(text=""))
        xlsx = self.wb.save()
        with zipfile.ZipFile(io.BytesIO(xlsx)) as package:
            xml = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
        self.assertIn('type="containsText"', xml)
        self.assertIn('operator="containsText"', xml)
        self.assertIn('text=""', xml)
        for data in (xlsx, self.wb.save(format=WorkbookFormat.XLSB)):
            with Workbook.load(data) as reloaded:
                rules = reloaded.sheet(0).list_cf_rules()
                self.assertEqual(len(rules), 1)
                self.assertEqual(rules[0].text, "")

    def test_contains_empty_text_matches_every_cell_that_is_not_an_error(self) -> None:
        self.ws.set_text("A1", "abc")
        self.ws.set_number("A2", 5)
        self.ws.set_formula("A4", "=1/0")
        self.wb.recalc()
        self.ws.add_cf_rule(self._text_rule(text=""))
        self.assertEqual([cell.row for cell in self.ws.evaluate_cf_range("A1:A4")], [0, 1, 2])


if __name__ == "__main__":
    unittest.main()
