"""Conditional-format surface tests."""

from __future__ import annotations

import unittest

from formulon import (
    CellRect,
    CfAboveAverageRuleInput,
    CfCellIsRuleInput,
    CfColorScaleInput,
    CfColorScaleRuleInput,
    CfDataBarInput,
    CfDataBarRuleInput,
    CfIconSetInput,
    CfIconSetRuleInput,
    CfTimePeriodRuleInput,
    CfTop10RuleInput,
    CfValueObject,
    Color,
    ColorKind,
    ComparisonOperator,
    ConditionalFormatInput,
    DxfInput,
    Workbook,
)


def _rgb(red: int, green: int, blue: int) -> Color:
    return Color(ColorKind.RGB, argb=0xFF000000 | (red << 16) | (green << 8) | blue)


class ConditionalFormatTests(unittest.TestCase):
    def test_cf_add_get_evaluate_clear(self) -> None:
        with Workbook.create_default() as wb:
            for r in range(3):
                wb.sheet(0).set_number(r, 0, float(r * 5))  # A1=0, A2=5, A3=10
            # A rule's dxf_index must resolve against a registered dxf.
            dxf_index = wb.ooxml.add_dxf(DxfInput())
            wb.sheet(0).add_cf_rule(
                CfCellIsRuleInput(
                    ranges=[CellRect(0, 0, 2, 0)],  # cellIs
                    operator=ComparisonOperator.GREATER_THAN,
                    formula1="4",
                    dxf_index=dxf_index,
                ),
            )
            rules = wb.sheet(0).list_cf_rules()
            self.assertEqual(len(rules), 1)
            self.assertEqual(rules[0].type, 1)
            self.assertEqual(rules[0].formula1, "=4")

            wb.recalc()
            cells = wb.sheet(0).evaluate_cf_range(CellRect(0, 0, 2, 0))
            matched = {(c.row, c.col) for c in cells if c.matches}
            # A2 (=5) and A3 (=10) exceed 4; A1 (=0) does not.
            self.assertIn((1, 0), matched)
            self.assertIn((2, 0), matched)
            self.assertNotIn((0, 0), matched)

            wb.sheet(0).clear_cf_rules()
            self.assertEqual(wb.sheet(0).list_cf_rules(), [])

    def test_cell_is_op_alone_is_engaged(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).add_cf_rule(
                CfCellIsRuleInput(ranges=[CellRect(0, 0, 2, 0)], operator=3, formula1="5"),
            )
            rule = wb.sheet(0).list_cf_rules()[0]
            self.assertEqual(rule.operator, 3)
            self.assertIsNone(rule.rank)
            self.assertIsNone(rule.std_dev)
            self.assertIsNone(rule.time_period)
            self.assertIsNone(rule.dxf_index)

    def test_unset_rank_reads_back_none(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).add_cf_rule(CfTop10RuleInput(ranges=[CellRect(0, 0, 2, 0)]))
            rule = wb.sheet(0).list_cf_rules()[0]
            self.assertIsNone(rule.rank)
            self.assertIsNone(rule.operator)

    def test_get_then_add_reproduces_engaged_fields(self) -> None:
        with Workbook.create_default() as wb:
            dxf_index = wb.ooxml.add_dxf(DxfInput())
            sqref = [CellRect(0, 0, 2, 0)]
            wb.sheet(0).add_cf_rule(
                CfCellIsRuleInput(
                    ranges=sqref, operator=ComparisonOperator.LESS_THAN, formula1="1", dxf_index=dxf_index
                )
            )
            wb.sheet(0).add_cf_rule(CfTop10RuleInput(ranges=sqref, rank=0, dxf_index=dxf_index))
            wb.sheet(0).add_cf_rule(CfAboveAverageRuleInput(ranges=sqref, std_dev=0.0))
            wb.sheet(0).add_cf_rule(CfTimePeriodRuleInput(ranges=sqref, time_period=0))
            first = wb.sheet(0).list_cf_rules()
            wb.sheet(0).clear_cf_rules()
            for r in first:
                wb.sheet(0).add_cf_rule(r)
            again = wb.sheet(0).list_cf_rules()
            for x, y in zip(first, again):
                for f in ("operator", "rank", "std_dev", "time_period", "dxf_index"):
                    self.assertEqual(getattr(x, f), getattr(y, f), f)

    def test_visual_rules_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            rules = [
                CfColorScaleRuleInput(
                    ranges=[CellRect(0, 0, 2, 0)],
                    color_scale=CfColorScaleInput(
                        [CfValueObject(3), CfValueObject(4)],
                        [_rgb(255, 0, 0), _rgb(0, 255, 0)],
                    ),
                ),
                CfDataBarRuleInput(
                    ranges=[CellRect(0, 1, 2, 1)],
                    data_bar=CfDataBarInput(CfValueObject(3), CfValueObject(4), _rgb(0, 0, 255)),
                ),
                CfIconSetRuleInput(
                    ranges=[CellRect(0, 2, 2, 2)],
                    icon_set=CfIconSetInput(0, [CfValueObject(1, "33"), CfValueObject(1, "67")]),
                ),
            ]
            for rule in rules:
                wb.sheet(0).add_cf_rule(rule)
            got = wb.sheet(0).list_cf_rules()
            self.assertEqual(got[0].color_scale.colors[1], _rgb(0, 255, 0))
            self.assertEqual(got[1].data_bar.fill, _rgb(0, 0, 255))
            self.assertEqual(got[2].icon_set.thresholds[1].value, "67")

    # `CfValueObject.type`: 0 num, 1 percent, 2 percentile, 3 min, 4 max,
    # 5 formula, 6 autoMin, 7 autoMax. Every threshold below states a
    # `value`, which is what puts a borrowed string behind each CFVO.
    EXPECTED_CFVO_VALUES = ["0", "50", "100", "5", "95", "0", "33", "67"]

    @staticmethod
    def _cfvo_rules() -> list[ConditionalFormatInput]:
        """Build one color-scale, one data-bar and one icon-set rule."""
        return [
            CfColorScaleRuleInput(
                ranges=[CellRect(0, 0, 4, 0)],
                color_scale=CfColorScaleInput(
                    [CfValueObject(0, "0"), CfValueObject(1, "50"), CfValueObject(0, "100")],
                    [_rgb(248, 105, 107), _rgb(255, 235, 132), _rgb(99, 190, 123)],
                ),
            ),
            CfDataBarRuleInput(
                ranges=[CellRect(0, 1, 4, 1)],
                data_bar=CfDataBarInput(CfValueObject(0, "5"), CfValueObject(0, "95"), _rgb(0, 112, 192)),
            ),
            CfIconSetRuleInput(
                ranges=[CellRect(0, 2, 4, 2)],
                # 4Arrows: a 4-icon set carries three thresholds.
                icon_set=CfIconSetInput(8, [CfValueObject(1, "0"), CfValueObject(1, "33"), CfValueObject(1, "67")]),
            ),
        ]

    @staticmethod
    def _cfvo_values(rules: list) -> list[str]:
        """Flatten the threshold value strings of the three visual rules."""
        by_type = {rule.type: rule for rule in rules}
        color_scale = by_type[2].color_scale
        data_bar = by_type[3].data_bar
        icon_set = by_type[4].icon_set
        return (
            [t.value for t in color_scale.thresholds]
            + [data_bar.min.value, data_bar.max.value]
            + [t.value for t in icon_set.thresholds]
        )

    def test_cfvo_value_strings_survive_read_back_and_a_save_load_cycle(self) -> None:
        # Each CFVO `value` crosses the C ABI as a borrowed `const char*`.
        # A store that relocates while the later thresholds are pulled
        # publishes a pointer into freed bytes, so the strings come back
        # wrong or empty even though the rule count and colors look right.
        with Workbook.create_default() as wb:
            for rule in self._cfvo_rules():
                wb.sheet(0).add_cf_rule(rule)

            in_session = wb.sheet(0).list_cf_rules()
            self.assertEqual(len(in_session), 3)
            self.assertEqual(self._cfvo_values(in_session), self.EXPECTED_CFVO_VALUES)
            by_type = {rule.type: rule for rule in in_session}
            # The type travels with the value: a percent threshold read
            # back as type 0 would mean an absolute number instead.
            self.assertEqual([t.type for t in by_type[2].color_scale.thresholds], [0, 1, 0])
            self.assertEqual([t.type for t in by_type[4].icon_set.thresholds], [1, 1, 1])
            self.assertEqual(by_type[2].color_scale.colors[2], _rgb(99, 190, 123))
            self.assertEqual(by_type[3].data_bar.fill, _rgb(0, 112, 192))
            saved = wb.save()

        with Workbook.load(saved) as reloaded:
            rules = reloaded.sheet(0).list_cf_rules()
            self.assertEqual(len(rules), 3)
            self.assertEqual(self._cfvo_values(rules), self.EXPECTED_CFVO_VALUES)

    def test_data_bar_x14_fields_survive_save_and_load(self) -> None:
        # These six live in the `x14` extension, not the legacy `<dataBar>`
        # element. An in-session round-trip alone would not catch a writer
        # that never emits the extension, which is how they were lost
        # before: the values came back from the model and disappeared on
        # the way through the file.
        bar = CfDataBarInput(
            CfValueObject(3),
            CfValueObject(4),
            _rgb(0, 0, 255),
            gradient=False,
            axis_position=1,
            negative_fill=_rgb(255, 0, 0),
            border=_rgb(9, 9, 9),
            negative_border=_rgb(8, 8, 8),
            axis_color=_rgb(1, 2, 3),
            direction=2,
        )
        with Workbook.create_default() as wb:
            wb.sheet(0).add_cf_rule(CfDataBarRuleInput(ranges=[CellRect(0, 0, 2, 0)], data_bar=bar))
            saved = wb.save()

        with Workbook.load(saved) as reloaded:
            got = reloaded.sheet(0).list_cf_rules()[0].data_bar
            self.assertIs(got.gradient, False)
            self.assertEqual(got.axis_position, 1)
            self.assertEqual(got.negative_fill, _rgb(255, 0, 0))
            self.assertEqual(got.border, _rgb(9, 9, 9))
            self.assertEqual(got.negative_border, _rgb(8, 8, 8))
            self.assertEqual(got.axis_color, _rgb(1, 2, 3))
            self.assertEqual(got.direction, 2)
            # Feeding the decoded bar straight back must reproduce it.
            reloaded.sheet(0).add_cf_rule(CfDataBarRuleInput(ranges=[CellRect(4, 0, 6, 0)], data_bar=got))
            self.assertEqual(reloaded.sheet(0).list_cf_rules()[1].data_bar, got)

    def test_icon_set_floor_survives_save_and_load(self) -> None:
        icons = CfIconSetInput(
            2, [CfValueObject(0, "7"), CfValueObject(0, "9")], floor=CfValueObject(0, "5", gte=False)
        )
        with Workbook.create_default() as wb:
            wb.sheet(0).add_cf_rule(CfIconSetRuleInput(ranges=[CellRect(0, 0, 9, 0)], icon_set=icons))
            saved = wb.save()
        with Workbook.load(saved) as reloaded:
            got = reloaded.sheet(0).list_cf_rules()[0].icon_set
            self.assertEqual(got.floor, CfValueObject(0, "5", gte=False))

    def test_omitted_data_bar_x14_fields_keep_the_model_defaults(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).add_cf_rule(
                CfDataBarRuleInput(
                    ranges=[CellRect(0, 0, 2, 0)],
                    data_bar=CfDataBarInput(CfValueObject(3), CfValueObject(4), _rgb(0, 0, 255)),
                ),
            )
            got = wb.sheet(0).list_cf_rules()[0].data_bar
            # The getter engages all six, so they read back as the defaults
            # rather than as `None`.
            self.assertIs(got.gradient, True)
            self.assertEqual(got.axis_position, 0)
            self.assertEqual(got.negative_fill, _rgb(0, 0, 255))


if __name__ == "__main__":
    unittest.main()
