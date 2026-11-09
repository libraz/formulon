"""Pivot references and readers, validation / CF updates, style record setters and record names."""

from __future__ import annotations

import dataclasses
import unittest
from pathlib import Path

import formulon
from formulon import (
    AlreadyExistsError,
    BorderSide,
    BorderStyle,
    CellXfInput,
    CfCellIsRuleInput,
    CfIconSetInput,
    CfIconSetRuleInput,
    CfValueObject,
    Color,
    ColorKind,
    ComparisonOperator,
    DataValidationInput,
    DxfInput,
    ErrorCode,
    ExcelError,
    Fill,
    FillPattern,
    FontScheme,
    HorizontalAlign,
    IconSetName,
    IndexOutOfRangeError,
    InvalidArgumentError,
    NotFoundError,
    PivotAggregation,
    PivotAxis,
    PivotDataFieldInput,
    PivotFieldInput,
    PivotFilterInput,
    PivotFilterKind,
    PivotInput,
    PivotReportLayout,
    PivotWorksheetSourceInput,
    ThemeScriptFont,
    ValidationErrorStyle,
    ValidationType,
    Workbook,
)

_NAMES_YAML = Path(__file__).resolve().parents[3] / "tools" / "api" / "names.yaml"

_RED = Color(kind=ColorKind.RGB, argb=0xFFFF0000)
_GREEN = Color(kind=ColorKind.RGB, argb=0xFF00FF00)


def _book() -> Workbook:
    """A workbook with A1:B3 filled and a pivot cache over it, two records stored."""
    wb = Workbook.create_default()
    sheet = wb.sheet(0)
    sheet.set_text(0, 0, "Region")
    sheet.set_text(0, 1, "Amount")
    for row, (region, amount) in enumerate((("East", 10.0), ("West", 30.0)), start=1):
        sheet.set_text(row, 0, region)
        sheet.set_number(row, 1, amount)
    return wb


def _cache(wb: Workbook) -> formulon.PivotCache:
    cache = wb.add_pivot_cache()
    cache.set_worksheet_source(PivotWorksheetSourceInput(range="A1:B3", sheet="Sheet1"))
    cache.add_field("Region")
    cache.add_field("Amount")
    for region, amount in (("East", 10.0), ("West", 30.0)):
        record = cache.add_record()
        cache.set_record_value(record, 0, region)
        cache.set_record_value(record, 1, amount)
    return cache


class PivotCacheRefTests(unittest.TestCase):
    def test_a_cache_handle_stands_for_its_id_everywhere(self) -> None:
        with _book() as wb:
            cache = _cache(wb)
            pivot = wb.sheet(0).add_pivot(PivotInput(name="P", cache_id=cache, anchor="E1"))
            self.assertEqual(pivot.get_properties().cache_id, cache.cache_id)
            self.assertIs(wb.pivot_cache(cache), cache)
            self.assertEqual(wb.pivot_cache(cache.cache_id).cache_id, cache.cache_id)
            other = wb.sheet(0).add_pivot(PivotInput(name="Q", cache_id=cache.cache_id, anchor="H1"))
            self.assertEqual(other.get_properties().cache_id, cache.cache_id)

    def test_a_cache_handle_of_another_workbook_is_invalid_argument(self) -> None:
        with _book() as wb, Workbook.create_default() as foreign:
            foreign_cache = foreign.add_pivot_cache()
            with self.assertRaises(InvalidArgumentError):
                wb.pivot_cache(foreign_cache)
            with self.assertRaises(InvalidArgumentError):
                wb.sheet(0).add_pivot(PivotInput(name="P", cache_id=foreign_cache, anchor="E1"))
            self.assertEqual(wb.sheet(0).list_pivots(), [])

    def test_a_cache_ref_that_is_neither_an_id_nor_a_handle_is_a_type_error(self) -> None:
        with _book() as wb:
            with self.assertRaises(TypeError):
                wb.pivot_cache("1")  # type: ignore[arg-type]


class PivotFieldRefTests(unittest.TestCase):
    @staticmethod
    def _pivot(wb: Workbook) -> formulon.Pivot:
        cache = _cache(wb)
        pivot = wb.sheet(0).add_pivot(PivotInput(name="P", cache_id=cache, anchor="E1"))
        pivot.add_field(PivotFieldInput("Region", PivotAxis.ROW))
        pivot.add_field(PivotFieldInput("Amount", PivotAxis.VALUE))
        return pivot

    def test_a_filter_names_its_field_by_name_or_index_and_reads_back_the_index(self) -> None:
        with _book() as wb:
            pivot = self._pivot(wb)
            pivot.add_filter(PivotFilterInput(field_index="region", type=PivotFilterKind.LABEL_EQUAL, value="East"))
            pivot.add_filter(PivotFilterInput(field_index=0, type=PivotFilterKind.LABEL_EQUAL, value="East"))
            by_name, by_index = pivot.list_filters()
            self.assertEqual(by_name, by_index)
            self.assertEqual(by_name.field_index, 0)

    def test_a_filter_on_a_field_the_pivot_lacks_is_refused_and_adds_nothing(self) -> None:
        with _book() as wb:
            pivot = self._pivot(wb)
            with self.assertRaises(NotFoundError):
                pivot.add_filter(PivotFilterInput(field_index="Nope", type=PivotFilterKind.LABEL_EQUAL, value="x"))
            with self.assertRaises(IndexOutOfRangeError):
                pivot.add_filter(PivotFilterInput(field_index=9, type=PivotFilterKind.LABEL_EQUAL, value="x"))
            self.assertEqual(pivot.list_filters(), [])

    def test_field_input_takes_name_or_index_and_a_required_axis(self) -> None:
        with _book() as wb:
            cache = _cache(wb)
            pivot = wb.sheet(0).add_pivot(PivotInput(name="P", cache_id=cache, anchor="E1"))
            by_name = pivot.add_field(PivotFieldInput("Region", PivotAxis.ROW))
            self.assertEqual(by_name, 0)
            self.assertIsNone(pivot.get_field(by_name).custom_name)
            with self.assertRaises(TypeError):
                PivotFieldInput("Region")  # type: ignore[call-arg]

    def test_a_data_field_defaults_to_sum_and_a_derived_name(self) -> None:
        with _book() as wb:
            pivot = self._pivot(wb)
            index = pivot.add_data_field(PivotDataFieldInput(1))
            data_field = pivot.get_data_field(index)
            self.assertEqual(data_field.aggregation, PivotAggregation.SUM)
            self.assertTrue(data_field.name.endswith("Amount"))
            self.assertEqual(PivotDataFieldInput(field_index="Amount").name, None)


class PivotSubtotalSetTests(unittest.TestCase):
    @staticmethod
    def _pivot(wb: Workbook) -> formulon.Pivot:
        return PivotFieldRefTests._pivot(wb)

    def test_the_set_replaces_the_functions_and_an_empty_list_clears_them(self) -> None:
        with _book() as wb:
            pivot = self._pivot(wb)
            pivot.set_field_subtotal_functions(0, [PivotAggregation.SUM, PivotAggregation.COUNT])
            field = pivot.get_field(0)
            self.assertEqual(field.subtotal_functions, [PivotAggregation.SUM, PivotAggregation.COUNT])
            self.assertFalse(field.default_subtotal)
            pivot.set_field_subtotal_functions("Region", [PivotAggregation.MAX])
            self.assertEqual(pivot.get_field(0).subtotal_functions, [PivotAggregation.MAX])
            pivot.set_field_subtotal_functions(0, [])
            self.assertEqual(pivot.get_field(0).subtotal_functions, [])
            self.assertTrue(pivot.get_field(0).default_subtotal)

    def test_a_function_listed_twice_is_refused_and_changes_nothing(self) -> None:
        with _book() as wb:
            pivot = self._pivot(wb)
            pivot.set_field_subtotal_functions(0, [PivotAggregation.MAX])
            with self.assertRaises(InvalidArgumentError):
                pivot.set_field_subtotal_functions(0, [PivotAggregation.SUM, PivotAggregation.SUM])
            self.assertEqual(pivot.get_field(0).subtotal_functions, [PivotAggregation.MAX])

    def test_the_add_and_clear_members_are_gone(self) -> None:
        with _book() as wb:
            pivot = self._pivot(wb)
            self.assertFalse(hasattr(pivot, "add_field_subtotal_function"))
            self.assertFalse(hasattr(pivot, "clear_field_subtotal_functions"))


class PivotReaderTests(unittest.TestCase):
    @staticmethod
    def _pivot(wb: Workbook) -> formulon.Pivot:
        pivot = PivotFieldRefTests._pivot(wb)
        pivot.add_data_field(PivotDataFieldInput(1, name="Total"))
        pivot.add_filter(PivotFilterInput(field_index=0, type=PivotFilterKind.LABEL_EQUAL, value="East"))
        return pivot

    def test_each_single_reader_equals_the_list_element_at_its_position(self) -> None:
        with _book() as wb:
            pivot = self._pivot(wb)
            for index, field in enumerate(pivot.list_fields()):
                self.assertEqual(pivot.get_field(index), field)
            self.assertEqual(pivot.get_field("Amount"), pivot.list_fields()[1])
            self.assertEqual(pivot.get_data_field(0), pivot.list_data_fields()[0])
            self.assertEqual(pivot.get_filter(0), pivot.list_filters()[0])

    def test_a_position_past_the_end_or_below_zero_is_out_of_range(self) -> None:
        with _book() as wb:
            pivot = self._pivot(wb)
            for call in (
                lambda: pivot.get_field(2),
                lambda: pivot.get_data_field(1),
                lambda: pivot.get_filter(1),
                lambda: pivot.get_data_field(-1),
                lambda: pivot.get_filter(-1),
            ):
                with self.assertRaises(IndexOutOfRangeError):
                    call()

    def test_properties_compose_the_separate_getters(self) -> None:
        with _book() as wb:
            pivot = self._pivot(wb)
            pivot.set_report_layout(PivotReportLayout.TABULAR)
            pivot.update_grand_totals(rows=False)
            pivot.update_style_options(show_row_stripes=True)
            properties = pivot.get_properties()
            self.assertEqual(properties.name, "P")
            self.assertEqual(properties.report_layout, pivot.get_report_layout())
            self.assertEqual(properties.report_layout, PivotReportLayout.TABULAR)
            self.assertEqual(properties.grand_totals, pivot.get_grand_totals())
            self.assertFalse(properties.grand_totals.rows)
            self.assertEqual(properties.style_options, pivot.get_style_options())
            self.assertTrue(properties.style_options.show_row_stripes)

    def test_repr_names_the_pivot_and_survives_removal(self) -> None:
        with _book() as wb:
            pivot = self._pivot(wb)
            self.assertIn("'P'", repr(pivot))
            wb.sheet(0).remove_pivot("P")
            self.assertIsInstance(repr(pivot), str)

    def test_a_duplicate_pivot_name_is_refused(self) -> None:
        with _book() as wb:
            cache = _cache(wb)
            wb.sheet(0).add_pivot(PivotInput(name="P", cache_id=cache, anchor="E1"))
            with self.assertRaises(AlreadyExistsError):
                wb.sheet(0).add_pivot(PivotInput(name="p", cache_id=cache, anchor="H1"))


class PivotCacheRecordTests(unittest.TestCase):
    def test_a_stored_value_reads_back_in_its_own_type(self) -> None:
        with _book() as wb:
            cache = _cache(wb)
            record = cache.add_record()
            cache.set_record_value(record, 0, True)
            cache.set_record_value(record, 1, ExcelError(ErrorCode.NA))
            self.assertEqual(cache.get_record_value(0, 0), "East")
            self.assertEqual(cache.get_record_value(0, 1), 10.0)
            self.assertIs(cache.get_record_value(record, 0), True)
            self.assertEqual(cache.get_record_value(record, 1), ExcelError(ErrorCode.NA))
            blank = cache.add_record()
            self.assertIsNone(cache.get_record_value(blank, 0))

    def test_a_position_past_the_end_is_out_of_range(self) -> None:
        with _book() as wb:
            cache = _cache(wb)
            for args in ((2, 0), (0, 2), (-1, 0)):
                with self.assertRaises(IndexOutOfRangeError):
                    cache.get_record_value(*args)


class ValidationAndCfUpdateTests(unittest.TestCase):
    @staticmethod
    def _validation() -> DataValidationInput:
        return DataValidationInput(
            type=ValidationType.LIST,
            ranges="A5:A6",
            formula1='"a,b"',
            error_message="bad",
            error_title="Title",
        )

    def test_an_update_keeps_the_fields_it_does_not_name_and_none_clears_a_text(self) -> None:
        with _book() as wb:
            sheet = wb.sheet(0)
            rule_id = sheet.add_validation(self._validation())
            sheet.update_validation(rule_id, error_message="changed", error_style=ValidationErrorStyle.WARNING)
            rule = sheet.list_validations()[0]
            self.assertEqual((rule.error_message, rule.error_style), ("changed", ValidationErrorStyle.WARNING))
            self.assertEqual((rule.error_title, rule.formula1, rule.type), ("Title", '="a,b"', ValidationType.LIST))
            sheet.update_validation(rule_id, error_title=None, formula2=None)
            self.assertIsNone(sheet.list_validations()[0].error_title)
            before = sheet.list_validations()[0]
            sheet.update_validation(rule_id)
            self.assertEqual(sheet.list_validations()[0], before)

    def test_a_refused_validation_update_changes_nothing(self) -> None:
        with _book() as wb:
            sheet = wb.sheet(0)
            rule_id = sheet.add_validation(self._validation())
            before = sheet.list_validations()[0]
            with self.assertRaises(TypeError):
                sheet.update_validation(rule_id, error_message="x", type=None)  # type: ignore[arg-type]
            with self.assertRaises(NotFoundError):
                sheet.update_validation("nope", error_message="x")
            with self.assertRaises(IndexOutOfRangeError):
                sheet.update_validation(5, error_message="x")
            self.assertEqual(sheet.list_validations()[0], before)

    def test_a_cf_rule_update_by_id_or_position_keeps_its_type(self) -> None:
        with _book() as wb:
            sheet = wb.sheet(0)
            dxf = wb.ooxml.add_dxf(DxfInput())
            sheet.add_cf_rule(
                CfCellIsRuleInput(ranges="A1:A3", operator=ComparisonOperator.GREATER_THAN, formula1="5", dxf_index=dxf)
            )
            rule = sheet.list_cf_rules()[0]
            sheet.update_cf_rule(0, formula1="7", operator=ComparisonOperator.LESS_THAN)
            sheet.update_cf_rule(rule.id, stop_if_true=True)
            updated = sheet.list_cf_rules()[0]
            self.assertEqual(updated.formula1, "=7")
            self.assertEqual(updated.operator, ComparisonOperator.LESS_THAN)
            self.assertTrue(updated.stop_if_true)
            self.assertEqual((updated.id, updated.type, updated.dxf_index), (rule.id, rule.type, dxf))
            sheet.update_cf_rule(0, dxf_index=None)
            self.assertIsNone(sheet.list_cf_rules()[0].dxf_index)

    def test_a_cf_rule_update_refuses_a_field_of_another_type_and_changes_nothing(self) -> None:
        with _book() as wb:
            sheet = wb.sheet(0)
            sheet.add_cf_rule(CfCellIsRuleInput(ranges="A1:A3", operator=ComparisonOperator.EQUAL, formula1="1"))
            before = sheet.list_cf_rules()[0]
            with self.assertRaises(InvalidArgumentError):
                sheet.update_cf_rule(0, rank=3)
            with self.assertRaises(TypeError):
                sheet.update_cf_rule(0, priority=None)  # type: ignore[arg-type]
            with self.assertRaises(IndexOutOfRangeError):
                sheet.update_cf_rule(4, formula1="2")
            self.assertEqual(sheet.list_cf_rules()[0], before)

    def test_a_cf_payload_update_replaces_the_whole_payload(self) -> None:
        with _book() as wb:
            sheet = wb.sheet(0)
            sheet.add_cf_rule(
                CfIconSetRuleInput(
                    ranges="A1:A3",
                    icon_set=CfIconSetInput(IconSetName.THREE_ARROWS, [CfValueObject(1, "33"), CfValueObject(1, "67")]),
                )
            )
            sheet.update_cf_rule(
                0,
                icon_set=CfIconSetInput(IconSetName.THREE_FLAGS, [CfValueObject(1, "10"), CfValueObject(1, "90")]),
            )
            icon_set = sheet.list_cf_rules()[0].icon_set
            self.assertEqual(icon_set.name, IconSetName.THREE_FLAGS)
            self.assertEqual([t.value for t in icon_set.thresholds], ["10", "90"])


class IdAddressedCollectionTests(unittest.TestCase):
    def test_a_validation_id_survives_a_row_insert_and_another_removal(self) -> None:
        with _book() as wb:
            sheet = wb.sheet(0)
            first = sheet.add_validation(
                DataValidationInput(type=ValidationType.WHOLE, ranges="C5:C6", formula1="1", formula2="9")
            )
            second = sheet.add_validation(
                DataValidationInput(type=ValidationType.WHOLE, ranges="D5:D6", formula1="1", formula2="9")
            )
            self.assertNotEqual(first, second)
            sheet.insert_rows(0, 0)
            sheet.remove_validation(first)
            remaining = sheet.list_validations()
            self.assertEqual([rule.id for rule in remaining], [second])
            self.assertEqual(remaining[0].ranges[0].first_row, 5)
            with self.assertRaises(NotFoundError):
                sheet.remove_validation(first)
            with self.assertRaises(NotFoundError):
                sheet.remove_validation("nope")
            with self.assertRaises(IndexOutOfRangeError):
                sheet.remove_validation(3)
            sheet.remove_validation(0)
            self.assertEqual(sheet.list_validations(), [])

    def test_ids_after_a_load_are_unique(self) -> None:
        with _book() as wb:
            sheet = wb.sheet(0)
            for column in ("C", "D", "E"):
                sheet.add_validation(
                    DataValidationInput(type=ValidationType.WHOLE, ranges=f"{column}5", formula1="1", formula2="9")
                )
            saved = wb.save()
        with Workbook.load(saved) as loaded:
            ids = [rule.id for rule in loaded.sheet(0).list_validations()]
            self.assertEqual(len(ids), 3)
            self.assertEqual(len(set(ids)), 3)
            self.assertTrue(all(ids))

    def test_a_sparkline_group_is_addressed_by_the_id_add_returns(self) -> None:
        with _book() as wb:
            sheet = wb.sheet(0)
            sheet.set_values("A5", [[1, 2, 3]])
            group_id = sheet.add_sparkline_group(
                formulon.SparklineGroupInput(sparklines=[formulon.SparklineInput("Sheet1!A5:C5", "D5")])
            )
            other = sheet.add_sparkline_group(
                formulon.SparklineGroupInput(sparklines=[formulon.SparklineInput("Sheet1!A5:C5", "D6")])
            )
            self.assertNotEqual(group_id, other)
            self.assertEqual([group.id for group in sheet.list_sparkline_groups()], [group_id, other])
            sheet.remove_sparkline_group(group_id)
            self.assertEqual([group.id for group in sheet.list_sparkline_groups()], [other])
            with self.assertRaises(NotFoundError):
                sheet.remove_sparkline_group(group_id)
            with self.assertRaises(IndexOutOfRangeError):
                sheet.remove_sparkline_group(4)


class StyleRecordSetterTests(unittest.TestCase):
    def test_a_fill_is_read_modified_and_written_back(self) -> None:
        with Workbook.create_default() as wb:
            index = wb.ooxml.add_fill(Fill(FillPattern.SOLID, _RED, Color(kind=ColorKind.NONE), None))
            fill = wb.ooxml.get_fill(index)
            wb.ooxml.set_fill(index, dataclasses.replace(fill, foreground=_GREEN))
            self.assertEqual(wb.ooxml.get_fill(index).foreground, _GREEN)
            self.assertEqual(wb.ooxml.get_fill(0), wb.ooxml.get_fill(0))

    def test_a_border_is_read_modified_and_written_back(self) -> None:
        with Workbook.create_default() as wb:
            border = wb.ooxml.get_border(0)
            index = wb.ooxml.add_border(
                dataclasses.replace(border, left=BorderSide(style=BorderStyle.THIN, color=_RED))
            )
            changed = dataclasses.replace(wb.ooxml.get_border(index), right=BorderSide(BorderStyle.THICK, _GREEN))
            wb.ooxml.set_border(index, changed)
            self.assertEqual(wb.ooxml.get_border(index), changed)

    def test_a_cell_xf_is_read_modified_and_written_back(self) -> None:
        with Workbook.create_default() as wb:
            index = wb.ooxml.add_cell_xf(
                CellXfInput(font_index=0, fill_index=0, border_index=0, num_fmt_id=0, wrap_text=True)
            )
            changed = dataclasses.replace(wb.ooxml.get_cell_xf(index), horizontal=HorizontalAlign.CENTER)
            wb.ooxml.set_cell_xf(index, changed)
            self.assertEqual(wb.ooxml.get_cell_xf(index), changed)
            self.assertEqual(wb.ooxml.get_cell_xf(index).horizontal, HorizontalAlign.CENTER)

    def test_a_dxf_is_read_modified_and_written_back(self) -> None:
        with Workbook.create_default() as wb:
            index = wb.ooxml.add_dxf(DxfInput())
            dxf = wb.ooxml.get_dxf(index)
            wb.ooxml.set_dxf(index, dxf)
            self.assertEqual(wb.ooxml.get_dxf(index), dxf)

    def test_an_index_past_the_end_is_out_of_range_for_every_setter(self) -> None:
        with Workbook.create_default() as wb:
            ooxml = wb.ooxml
            for call in (
                lambda: ooxml.set_fill(999, ooxml.get_fill(0)),
                lambda: ooxml.set_border(999, ooxml.get_border(0)),
                lambda: ooxml.set_cell_xf(999, ooxml.get_cell_xf(0)),
                lambda: ooxml.set_dxf(999, DxfInput()),
            ):
                with self.assertRaises(IndexOutOfRangeError):
                    call()


class ThemeScriptFontListTests(unittest.TestCase):
    def test_a_set_script_font_appears_in_its_scheme_list_only(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_theme_script_font(ThemeScriptFont(scheme=FontScheme.MAJOR, script="Jpan", typeface="Test Face"))
            major = wb.list_theme_script_fonts(FontScheme.MAJOR)
            self.assertIn(ThemeScriptFont(FontScheme.MAJOR, "Jpan", "Test Face"), major)
            self.assertTrue(all(font.scheme == FontScheme.MAJOR for font in major))
            minor = wb.list_theme_script_fonts(FontScheme.MINOR)
            self.assertTrue(all(font.typeface != "Test Face" for font in minor))

    def test_passthrough_parts_list_as_names(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_theme_script_font(ThemeScriptFont(scheme=FontScheme.MINOR, script="Zzzz", typeface="Face"))
            names = wb.ooxml.list_passthroughs()
            self.assertTrue(all(isinstance(name, str) for name in names))
            for name in names:
                self.assertTrue(wb.ooxml.has_passthrough(name))
            self.assertFalse(hasattr(formulon, "PassthroughPart"))


class EffectiveStyleBordersTests(unittest.TestCase):
    def test_borders_are_a_record_keyed_by_side(self) -> None:
        with Workbook.create_default() as wb:
            style = wb.sheet(0).ooxml.get_effective_style(0, 0)
            self.assertIsInstance(style.borders, formulon.EffectiveBorders)
            self.assertEqual(style.borders.left, style.borders.right)
            self.assertEqual(
                {f.name for f in dataclasses.fields(style.borders)}, {"left", "right", "top", "bottom", "diagonal"}
            )


def _retired_python_type_names() -> list:
    """The former Python names of every renamed record, typedef and enum in the name ledger, less any name an entry
    publishes today."""
    import yaml

    ledger = yaml.safe_load(_NAMES_YAML.read_text(encoding="utf-8"))["entries"]
    kinds = {"record", "typedef", "enum"}
    current = {e.get("py") for e in ledger.values() if isinstance(e, dict)}
    retired = set()
    for entry in ledger.values():
        if not isinstance(entry, dict) or entry.get("kind") not in kinds:
            continue
        was = entry.get("was")
        former = was.get("py") if isinstance(was, dict) else None
        if isinstance(former, str) and former != entry.get("py") and former not in current:
            retired.add(former)
    return sorted(retired)


class RecordNameTests(unittest.TestCase):
    def test_records_carry_the_names_the_other_bindings_use(self) -> None:
        for name in ("LocaleErrorName", "PaginationMargins", "PaginationPage", "PaginationPaper", "CfValueObject"):
            self.assertTrue(hasattr(formulon, name), name)
        # Names a post-release rename retired before they shipped are not in the ledger's `was` keys.
        for name in (*_retired_python_type_names(), "Paper", "Cfvo", "ColorScale", "DataBar", "IconSet"):
            self.assertFalse(hasattr(formulon, name), name)

    def test_one_input_dataclass_per_cf_rule_type(self) -> None:
        for name in (
            "CfExpressionRuleInput",
            "CfCellIsRuleInput",
            "CfColorScaleRuleInput",
            "CfDataBarRuleInput",
            "CfIconSetRuleInput",
            "CfTop10RuleInput",
            "CfAboveAverageRuleInput",
            "CfTextRuleInput",
            "CfTimePeriodRuleInput",
            "CfPlainRuleInput",
        ):
            self.assertTrue(dataclasses.is_dataclass(getattr(formulon, name)), name)

    def test_a_field_of_another_rule_type_is_a_type_error(self) -> None:
        with self.assertRaises(TypeError):
            CfCellIsRuleInput(ranges="A1", operator=ComparisonOperator.EQUAL, formula1="1", rank=3)  # type: ignore[call-arg]

    def test_a_read_rule_adds_back_without_conversion(self) -> None:
        with _book() as wb:
            sheet = wb.sheet(0)
            sheet.add_cf_rule(CfCellIsRuleInput(ranges="A1:A3", operator=ComparisonOperator.EQUAL, formula1="5"))
            rule = sheet.list_cf_rules()[0]
            sheet.add_cf_rule(rule)
            first, second = sheet.list_cf_rules()
            self.assertEqual((first.formula1, first.operator), (second.formula1, second.operator))
            self.assertNotEqual(first.id, second.id)


if __name__ == "__main__":
    unittest.main()
