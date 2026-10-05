"""Conditional-format surface tests."""

from __future__ import annotations

import unittest

from formulon import (
    CfColor,
    CfValueObject,
    ColorScale,
    ConditionalFormatInput,
    DataBar,
    DifferentialFormat,
    IconSet,
    MergeRange,
    Workbook,
)


class ConditionalFormatTests(unittest.TestCase):
    def test_cf_add_get_evaluate_clear(self) -> None:
        with Workbook.create_default() as wb:
            for r in range(3):
                wb.set_number(0, r, 0, float(r * 5))  # A1=0, A2=5, A3=10
            # A rule's dxf_id must resolve against a registered dxf.
            dxf_index = wb.add_dxf(DifferentialFormat())
            wb.add_conditional_format(
                0,
                ConditionalFormatInput(
                    sqref=[MergeRange(0, 0, 2, 0)],
                    type=1,  # cellIs
                    op_engaged=True,
                    op=4,  # greaterThan
                    formula1="4",
                    dxf_id_engaged=True,
                    dxf_id=dxf_index,
                ),
            )
            rules = wb.get_conditional_formats(0)
            self.assertEqual(len(rules), 1)
            self.assertEqual(rules[0].type, 1)
            self.assertEqual(rules[0].formula1, "4")

            wb.recalc()
            cells = wb.evaluate_cf_range(0, 0, 0, 2, 0)
            matched = {(c.row, c.col) for c in cells if c.matches}
            # A2 (=5) and A3 (=10) exceed 4; A1 (=0) does not.
            self.assertIn((1, 0), matched)
            self.assertIn((2, 0), matched)
            self.assertNotIn((0, 0), matched)

            wb.clear_conditional_formats(0)
            self.assertEqual(wb.get_conditional_formats(0), [])

    def test_visual_rules_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            rules = [
                ConditionalFormatInput(
                    sqref=[MergeRange(0, 0, 2, 0)],
                    type=2,
                    color_scale=ColorScale(
                        [CfValueObject(3), CfValueObject(4)],
                        [CfColor(255, 0, 0), CfColor(0, 255, 0)],
                    ),
                ),
                ConditionalFormatInput(
                    sqref=[MergeRange(0, 1, 2, 1)],
                    type=3,
                    data_bar=DataBar(CfValueObject(3), CfValueObject(4), CfColor(0, 0, 255)),
                ),
                ConditionalFormatInput(
                    sqref=[MergeRange(0, 2, 2, 2)],
                    type=4,
                    icon_set=IconSet(0, [CfValueObject(1, "33"), CfValueObject(1, "67")]),
                ),
            ]
            for rule in rules:
                wb.add_conditional_format(0, rule)
            got = wb.get_conditional_formats(0)
            self.assertEqual(got[0].color_scale.colors[1], CfColor(0, 255, 0))
            self.assertEqual(got[1].data_bar.fill, CfColor(0, 0, 255))
            self.assertEqual(got[2].icon_set.thresholds[1].value, "67")

    # `CfValueObject.type`: 0 num, 1 percent, 2 percentile, 3 min, 4 max,
    # 5 formula, 6 autoMin, 7 autoMax. Every threshold below states a
    # `value`, which is what puts a borrowed string behind each CFVO.
    EXPECTED_CFVO_VALUES = ["0", "50", "100", "5", "95", "0", "33", "67"]

    @staticmethod
    def _cfvo_rules() -> list[ConditionalFormatInput]:
        """Build one color-scale, one data-bar and one icon-set rule."""
        return [
            ConditionalFormatInput(
                sqref=[MergeRange(0, 0, 4, 0)],
                type=2,
                color_scale=ColorScale(
                    [CfValueObject(0, "0"), CfValueObject(1, "50"), CfValueObject(0, "100")],
                    [CfColor(248, 105, 107), CfColor(255, 235, 132), CfColor(99, 190, 123)],
                ),
            ),
            ConditionalFormatInput(
                sqref=[MergeRange(0, 1, 4, 1)],
                type=3,
                data_bar=DataBar(CfValueObject(0, "5"), CfValueObject(0, "95"), CfColor(0, 112, 192)),
            ),
            ConditionalFormatInput(
                sqref=[MergeRange(0, 2, 4, 2)],
                type=4,
                icon_set=IconSet(0, [CfValueObject(1, "0"), CfValueObject(1, "33"), CfValueObject(1, "67")]),
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
            + [data_bar.minimum.value, data_bar.maximum.value]
            + [t.value for t in icon_set.thresholds]
        )

    def test_cfvo_value_strings_survive_read_back_and_a_save_load_cycle(self) -> None:
        # Each CFVO `value` crosses the C ABI as a borrowed `const char*`.
        # A store that relocates while the later thresholds are pulled
        # publishes a pointer into freed bytes, so the strings come back
        # wrong or empty even though the rule count and colors look right.
        with Workbook.create_default() as wb:
            for rule in self._cfvo_rules():
                wb.add_conditional_format(0, rule)

            in_session = wb.get_conditional_formats(0)
            self.assertEqual(len(in_session), 3)
            self.assertEqual(self._cfvo_values(in_session), self.EXPECTED_CFVO_VALUES)
            by_type = {rule.type: rule for rule in in_session}
            # The type travels with the value: a percent threshold read
            # back as type 0 would mean an absolute number instead.
            self.assertEqual([t.type for t in by_type[2].color_scale.thresholds], [0, 1, 0])
            self.assertEqual([t.type for t in by_type[4].icon_set.thresholds], [1, 1, 1])
            self.assertEqual(by_type[2].color_scale.colors[2], CfColor(99, 190, 123))
            self.assertEqual(by_type[3].data_bar.fill, CfColor(0, 112, 192))
            saved = wb.save()

        with Workbook.load(saved) as reloaded:
            rules = reloaded.get_conditional_formats(0)
            self.assertEqual(len(rules), 3)
            self.assertEqual(self._cfvo_values(rules), self.EXPECTED_CFVO_VALUES)

    def test_data_bar_x14_fields_survive_save_and_load(self) -> None:
        # These six live in the `x14` extension, not the legacy `<dataBar>`
        # element. An in-session round-trip alone would not catch a writer
        # that never emits the extension, which is how they were lost
        # before: the values came back from the model and disappeared on
        # the way through the file.
        bar = DataBar(
            CfValueObject(3),
            CfValueObject(4),
            CfColor(0, 0, 255),
            gradient=False,
            axis_position=1,
            negative_fill=CfColor(255, 0, 0),
            border=CfColor(9, 9, 9),
            negative_border=CfColor(8, 8, 8),
            axis_color=CfColor(1, 2, 3),
            direction=2,
        )
        with Workbook.create_default() as wb:
            wb.add_conditional_format(0, ConditionalFormatInput(sqref=[MergeRange(0, 0, 2, 0)], type=3, data_bar=bar))
            saved = wb.save()

        with Workbook.load(saved) as reloaded:
            got = reloaded.get_conditional_formats(0)[0].data_bar
            self.assertIs(got.gradient, False)
            self.assertEqual(got.axis_position, 1)
            self.assertEqual(got.negative_fill, CfColor(255, 0, 0))
            self.assertEqual(got.border, CfColor(9, 9, 9))
            self.assertEqual(got.negative_border, CfColor(8, 8, 8))
            self.assertEqual(got.axis_color, CfColor(1, 2, 3))
            self.assertEqual(got.direction, 2)
            # Feeding the decoded bar straight back must reproduce it.
            reloaded.add_conditional_format(
                0, ConditionalFormatInput(sqref=[MergeRange(4, 0, 6, 0)], type=3, data_bar=got)
            )
            self.assertEqual(reloaded.get_conditional_formats(0)[1].data_bar, got)

    def test_icon_set_floor_survives_save_and_load(self) -> None:
        icons = IconSet(2, [CfValueObject(0, "7"), CfValueObject(0, "9")], floor=CfValueObject(0, "5", gte=False))
        with Workbook.create_default() as wb:
            wb.add_conditional_format(0, ConditionalFormatInput(sqref=[MergeRange(0, 0, 9, 0)], type=4, icon_set=icons))
            saved = wb.save()
        with Workbook.load(saved) as reloaded:
            got = reloaded.get_conditional_formats(0)[0].icon_set
            self.assertEqual(got.floor, CfValueObject(0, "5", gte=False))

    def test_omitted_data_bar_x14_fields_keep_the_model_defaults(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_conditional_format(
                0,
                ConditionalFormatInput(
                    sqref=[MergeRange(0, 0, 2, 0)],
                    type=3,
                    data_bar=DataBar(CfValueObject(3), CfValueObject(4), CfColor(0, 0, 255)),
                ),
            )
            got = wb.get_conditional_formats(0)[0].data_bar
            # The getter engages all six, so they read back as the defaults
            # rather than as `None`.
            self.assertIs(got.gradient, True)
            self.assertEqual(got.axis_position, 0)
            self.assertEqual(got.negative_fill, CfColor(0, 0, 255))


if __name__ == "__main__":
    unittest.main()
