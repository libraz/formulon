"""Enum tables, typed record fields and value-input unions, through ``import formulon``."""

from __future__ import annotations

import datetime
import enum
import unittest
from typing import List, Tuple

import _chart_input

import formulon
from formulon import (
    AnchorKind,
    AutoFilter,
    AutoFilterColumn,
    AutoFilterCondition,
    BlanksAs,
    CellAddress,
    CellRect,
    CfCellIsRuleInput,
    CfIcon,
    CfIconSetInput,
    CfIconSetRuleInput,
    CfTimePeriodRuleInput,
    CfValueObject,
    ChartLegend,
    ChartLegendPosition,
    ChartPointState,
    ComparisonOperator,
    DataValidationInput,
    DrawingAnchorInput,
    DxfInput,
    EffectiveStyleSource,
    ErrorCode,
    ExcelError,
    FilterKind,
    FontScheme,
    FormulonError,
    FunctionAvailability,
    GeometryMode,
    IconSetName,
    Person,
    PhoneticProperties,
    PivotAxis,
    PivotDateGroupInput,
    PivotDatePeriod,
    PivotFieldInput,
    PivotFilterKind,
    PivotInput,
    PivotSortOrder,
    SlicerSortOrder,
    TableColumnTotalsInput,
    TableInput,
    TableStyle,
    TableStyleElement,
    ThemeScriptFont,
    ThemeSource,
    ThreadedCommentInput,
    ValidationOutcome,
    ValidationType,
    Value,
    ValueKind,
    Workbook,
)


class _Absent:
    """Stands in for a public name the package does not export yet; any use fails the test that makes it."""

    def __init__(self, name: str) -> None:
        self._name = name

    def _fail(self) -> "AssertionError":
        return AssertionError(f"formulon.{self._name} is not exported")

    def __getattr__(self, attr: str):
        raise self._fail()

    def __call__(self, *args, **kwargs):
        raise self._fail()

    def __iter__(self):
        raise self._fail()


ChartDomain = getattr(formulon, "ChartDomain", None) or _Absent("ChartDomain")
ChartTickLabelPosition = getattr(formulon, "ChartTickLabelPosition", None) or _Absent("ChartTickLabelPosition")
DatePeriod = getattr(formulon, "DatePeriod", None) or _Absent("DatePeriod")
PageOrder = getattr(formulon, "PageOrder", None) or _Absent("PageOrder")
PaperSize = getattr(formulon, "PaperSize", None) or _Absent("PaperSize")
PhoneticAlignment = getattr(formulon, "PhoneticAlignment", None) or _Absent("PhoneticAlignment")
PhoneticType = getattr(formulon, "PhoneticType", None) or _Absent("PhoneticType")
TableStyleElementType = getattr(formulon, "TableStyleElementType", None) or _Absent("TableStyleElementType")
TableTotalsRowFunction = getattr(formulon, "TableTotalsRowFunction", None) or _Absent("TableTotalsRowFunction")
WidthPlatform = getattr(formulon, "WidthPlatform", None) or _Absent("WidthPlatform")

ALICE = "{11111111-1111-4111-8111-111111111111}"
THREAD = "{33333333-3333-4333-8333-333333333333}"


def _members(enum_type: "type[enum.Enum]") -> List[Tuple[str, int]]:
    return [(m.name, int(m.value)) for m in enum_type]  # type: ignore[call-overload]


def _serial(day: datetime.date) -> float:
    return float((day - datetime.date(1899, 12, 30)).days)


def _chart_workbook(kind: _chart_input.Kind = _chart_input.COLUMN, source: str = "Sheet1!$A$1:$C$4"):
    wb = Workbook.create_default()
    sheet = wb.sheet(0)
    sheet.set_text(0, 1, "S1")
    sheet.set_text(0, 2, "S2")
    for r, cat in enumerate("abc"):
        sheet.set_text(r + 1, 0, cat)
        sheet.set_number(r + 1, 1, float(r + 1))
        sheet.set_number(r + 1, 2, float((r + 1) * 2))
    oid = _chart_input.add_chart(
        sheet, kind, source, anchor=DrawingAnchorInput(kind=AnchorKind.TWO_CELL, from_row=6)
    ).object_id
    return wb, oid


class EnumTableValueTests(unittest.TestCase):
    def test_comparison_operator_follows_the_validation_order(self) -> None:
        self.assertEqual(
            _members(ComparisonOperator),
            [
                ("BETWEEN", 0),
                ("NOT_BETWEEN", 1),
                ("EQUAL", 2),
                ("NOT_EQUAL", 3),
                ("LESS_THAN", 4),
                ("LESS_THAN_OR_EQUAL", 5),
                ("GREATER_THAN", 6),
                ("GREATER_THAN_OR_EQUAL", 7),
            ],
        )

    def test_replaced_operator_and_blanks_tables_are_gone(self) -> None:
        for name in (
            "CfOperator",
            "ValidationOperator",
            "SparklineEmptyCells",
            "ChartBlanksAs",
            "CfTimePeriod",
            "FilterOperator",
            "PivotFilterOperandKind",
        ):
            with self.subTest(name=name):
                self.assertFalse(hasattr(formulon, name))
                self.assertNotIn(name, formulon.__all__)

    def test_date_period_follows_the_time_period_schema_order(self) -> None:
        self.assertEqual(
            _members(DatePeriod),
            [
                ("TODAY", 0),
                ("YESTERDAY", 1),
                ("TOMORROW", 2),
                ("LAST_7_DAYS", 3),
                ("THIS_MONTH", 4),
                ("LAST_MONTH", 5),
                ("NEXT_MONTH", 6),
                ("THIS_WEEK", 7),
                ("LAST_WEEK", 8),
                ("NEXT_WEEK", 9),
            ],
        )

    def test_blanks_as_puts_zero_first_and_has_no_absent_member(self) -> None:
        self.assertEqual(_members(BlanksAs), [("ZERO", 0), ("GAP", 1), ("SPAN", 2)])

    def test_legend_position_puts_the_default_right_at_zero(self) -> None:
        self.assertEqual(
            _members(ChartLegendPosition),
            [("RIGHT", 0), ("BOTTOM", 1), ("TOP_RIGHT", 2), ("LEFT", 3), ("TOP", 4)],
        )

    def test_pivot_and_slicer_sort_put_the_default_at_zero(self) -> None:
        self.assertEqual(
            [(PivotSortOrder.MANUAL, 0), (PivotSortOrder.ASCENDING, 1), (PivotSortOrder.DESCENDING, 2)],
            [(m, int(m)) for m in PivotSortOrder],
        )
        self.assertEqual(_members(SlicerSortOrder), [("ASCENDING", 0), ("DESCENDING", 1)])

    def test_table_totals_row_function_follows_the_schema_order(self) -> None:
        self.assertTrue(issubclass(TableTotalsRowFunction, enum.IntEnum))
        self.assertEqual(
            _members(TableTotalsRowFunction),
            [
                ("NONE", 0),
                ("SUM", 1),
                ("MIN", 2),
                ("MAX", 3),
                ("AVERAGE", 4),
                ("COUNT", 5),
                ("COUNT_NUMBERS", 6),
                ("STD_DEV", 7),
                ("VAR", 8),
                ("CUSTOM", 9),
            ],
        )

    def test_table_style_element_type_follows_the_schema_order(self) -> None:
        names = [
            "WHOLE_TABLE",
            "HEADER_ROW",
            "TOTAL_ROW",
            "FIRST_COLUMN",
            "LAST_COLUMN",
            "FIRST_ROW_STRIPE",
            "SECOND_ROW_STRIPE",
            "FIRST_COLUMN_STRIPE",
            "SECOND_COLUMN_STRIPE",
            "FIRST_HEADER_CELL",
            "LAST_HEADER_CELL",
            "FIRST_TOTAL_CELL",
            "LAST_TOTAL_CELL",
            "FIRST_SUBTOTAL_COLUMN",
            "SECOND_SUBTOTAL_COLUMN",
            "THIRD_SUBTOTAL_COLUMN",
            "FIRST_SUBTOTAL_ROW",
            "SECOND_SUBTOTAL_ROW",
            "THIRD_SUBTOTAL_ROW",
            "BLANK_ROW",
            "FIRST_COLUMN_SUBHEADING",
            "SECOND_COLUMN_SUBHEADING",
            "THIRD_COLUMN_SUBHEADING",
            "FIRST_ROW_SUBHEADING",
            "SECOND_ROW_SUBHEADING",
            "THIRD_ROW_SUBHEADING",
            "PAGE_FIELD_LABELS",
            "PAGE_FIELD_VALUES",
        ]
        self.assertTrue(issubclass(TableStyleElementType, enum.IntEnum))
        self.assertEqual(_members(TableStyleElementType), list(zip(names, range(len(names)))))

    def test_chart_tick_label_position_and_domain(self) -> None:
        self.assertEqual(_members(ChartTickLabelPosition), [("NEXT_TO", 0), ("HIGH", 1), ("LOW", 2), ("NONE", 3)])
        self.assertEqual(_members(ChartDomain), [("CATEGORY", 0), ("XY", 1)])

    def test_width_platform_puts_windows_first(self) -> None:
        self.assertEqual(_members(WidthPlatform), [("WIN", 0), ("MAC", 1)])

    def test_phonetic_enums_keep_the_absent_element_state_at_zero(self) -> None:
        self.assertEqual(
            _members(PhoneticType),
            [("HALFWIDTH_KATAKANA", 0), ("FULLWIDTH_KATAKANA", 1), ("HIRAGANA", 2), ("NO_CONVERSION", 3)],
        )
        self.assertEqual(
            _members(PhoneticAlignment),
            [("NO_CONTROL", 0), ("LEFT", 1), ("CENTER", 2), ("DISTRIBUTED", 3)],
        )

    def test_page_order_and_paper_size_constants(self) -> None:
        self.assertEqual(_members(PageOrder), [("DOWN_THEN_OVER", 0), ("OVER_THEN_DOWN", 1)])
        self.assertFalse(isinstance(PaperSize, enum.EnumMeta))
        self.assertEqual(
            {name: getattr(PaperSize, name) for name in ("LETTER", "LEGAL", "A3", "A4", "A5", "B4", "B5", "A6")},
            {"LETTER": 1, "LEGAL": 5, "A3": 8, "A4": 9, "A5": 11, "B4": 12, "B5": 13, "A6": 70},
        )
        for name in ("LETTER", "A4", "A6"):
            self.assertIs(type(getattr(PaperSize, name)), int)

    def test_function_availability_and_style_source_values(self) -> None:
        self.assertEqual(
            _members(FunctionAvailability),
            [
                ("IMPLEMENTED", 0),
                ("IMPLEMENTED_UNVERIFIED", 1),
                ("ENVIRONMENT_BOUND", 2),
                ("UNAVAILABLE_STUB", 3),
            ],
        )
        self.assertEqual(_members(EffectiveStyleSource), [("CELL", 0), ("ROW", 1), ("COLUMN", 2), ("DEFAULT", 3)])
        self.assertEqual(_members(ThemeSource), [("PART", 0), ("DEFAULT", 1), ("UNPARSEABLE", 2)])

    def test_chart_point_state_values(self) -> None:
        self.assertEqual(_members(ChartPointState), [("VALUE", 0), ("GAP", 1), ("CONNECT", 2)])

    def test_icon_set_name_gains_the_custom_only_sets_after_the_builtins(self) -> None:
        members = dict(_members(IconSetName))
        self.assertEqual(len(members), 21)
        self.assertEqual(members["FIVE_QUARTERS"], 16)
        self.assertEqual(
            (members["THREE_STARS"], members["THREE_TRIANGLES"], members["FIVE_BOXES"], members["NO_ICONS"]),
            (17, 18, 19, 20),
        )

    def test_pivot_date_period_is_a_bit_flag(self) -> None:
        self.assertTrue(issubclass(PivotDatePeriod, enum.IntFlag))
        self.assertEqual(
            [(m.name, int(m)) for m in PivotDatePeriod],
            [
                ("SECONDS", 1),
                ("MINUTES", 2),
                ("HOURS", 4),
                ("DAYS", 8),
                ("MONTHS", 16),
                ("QUARTERS", 32),
                ("YEARS", 64),
            ],
        )
        self.assertEqual(int(PivotDatePeriod.MONTHS | PivotDatePeriod.YEARS), 80)

    def test_pivot_filter_relative_periods_drop_the_date_prefix(self) -> None:
        for name in ("TODAY", "YESTERDAY", "TOMORROW", "THIS_WEEK", "LAST_MONTH", "NEXT_QUARTER", "YEAR_TO_DATE"):
            with self.subTest(name=name):
                self.assertTrue(hasattr(PivotFilterKind, name))
                self.assertFalse(hasattr(PivotFilterKind, "DATE_" + name))
        for name in ("Q1", "Q4", "M1", "M12"):
            with self.subTest(name=name):
                self.assertTrue(hasattr(PivotFilterKind, name))
        self.assertTrue(hasattr(PivotFilterKind, "DATE_BETWEEN"))

    def test_shared_period_spellings_match_the_date_period_table(self) -> None:
        shared = {m.name for m in DatePeriod} - {"LAST_7_DAYS"}
        self.assertTrue(shared <= {m.name for m in PivotFilterKind})
        self.assertTrue(shared <= {m.name for m in formulon.DynamicFilterType})


class EnumRecordFieldTests(unittest.TestCase):
    def test_cf_operator_and_time_period_round_trip_as_enums(self) -> None:
        with Workbook.create_default() as wb:
            sheet = wb.sheet(0)
            sheet.add_cf_rule(
                CfCellIsRuleInput(ranges=[CellRect(0, 0, 2, 0)], operator=ComparisonOperator.GREATER_THAN, formula1="4")
            )
            sheet.add_cf_rule(CfTimePeriodRuleInput(ranges=[CellRect(0, 1, 2, 1)], time_period=DatePeriod.THIS_WEEK))
            saved = wb.save()
        with Workbook.load(saved) as reloaded:
            rules = reloaded.sheet(0).list_cf_rules()
            self.assertIs(rules[0].operator, ComparisonOperator.GREATER_THAN)
            self.assertEqual(int(rules[0].operator), 6)
            self.assertIs(rules[1].time_period, DatePeriod.THIS_WEEK)
            self.assertEqual(int(rules[1].time_period), 7)

    def test_validation_operator_uses_the_shared_table(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).add_validation(
                DataValidationInput(
                    type=ValidationType.WHOLE,
                    ranges=[CellRect(0, 0, 4, 0)],
                    operator=ComparisonOperator.LESS_THAN,
                    formula1="10",
                )
            )
            saved = wb.save()
        with Workbook.load(saved) as reloaded:
            rule = reloaded.sheet(0).list_validations()[0]
            self.assertIs(rule.operator, ComparisonOperator.LESS_THAN)
            self.assertEqual(int(rule.operator), 4)
            self.assertTrue(reloaded.sheet(0).validate_value(0, 0, 3).valid)
            self.assertFalse(reloaded.sheet(0).validate_value(0, 0, 30).valid)

    def test_autofilter_reads_the_operator_it_stated_and_refuses_range_members(self) -> None:
        with Workbook.create_default() as wb:
            sheet = wb.sheet(0)
            sheet.set_text(0, 0, "n")
            for i in range(1, 6):
                sheet.set_number(i, 0, float(i))

            def custom(op: ComparisonOperator) -> AutoFilter:
                return AutoFilter(
                    range=CellRect(0, 0, 5, 0),
                    columns=[
                        AutoFilterColumn(
                            column_offset=0, kind=FilterKind.CUSTOM, conditions=[AutoFilterCondition(op, "2")]
                        )
                    ],
                )

            sheet.set_auto_filter(custom(ComparisonOperator.LESS_THAN_OR_EQUAL))
            got = sheet.get_auto_filter()
            assert got is not None
            self.assertIs(got.columns[0].conditions[0].operator, ComparisonOperator.LESS_THAN_OR_EQUAL)
            for op in (ComparisonOperator.BETWEEN, ComparisonOperator.NOT_BETWEEN):
                with self.subTest(op=op), self.assertRaises(FormulonError):
                    sheet.set_auto_filter(custom(op))

    def test_legend_position_and_blanks_round_trip_by_member(self) -> None:
        wb, oid = _chart_workbook()
        with wb:
            chart = wb.sheet(0).chart(oid)
            for position in ChartLegendPosition:
                with self.subTest(position=position):
                    chart.set_legend(ChartLegend(position, False, []))
                    self.assertIs(chart.get_properties().legend.position, position)
            chart.set_legend(None)
            self.assertIsNone(chart.get_properties().legend)
            chart.update_display(blanks_as=BlanksAs.ZERO)
            self.assertIs(chart.get_properties().blanks_as, BlanksAs.ZERO)
            chart.update_display(blanks_as=None)
            self.assertIsNone(chart.get_properties().blanks_as)

    def test_chart_data_reports_enum_domain_and_point_states(self) -> None:
        wb, oid = _chart_workbook()
        with wb:
            data = wb.sheet(0).chart(oid).get_data()
            self.assertIs(data.domain, ChartDomain.CATEGORY)
            states = data.series[0].states
            self.assertTrue(states)
            self.assertTrue(all(isinstance(s, ChartPointState) for s in states))
        wb, oid = _chart_workbook(_chart_input.SCATTER, "Sheet1!$B$1:$C$4")
        with wb:
            self.assertIs(wb.sheet(0).chart(oid).get_data().domain, ChartDomain.XY)

    def test_axis_tick_label_position_is_an_enum_and_none_is_a_value(self) -> None:
        wb, oid = _chart_workbook()
        with wb:
            chart = wb.sheet(0).chart(oid)
            axis_id = next(a.id for a in chart.get_properties().axes if a.kind == formulon.ChartAxisKind.VALUE)
            for position in ChartTickLabelPosition:
                with self.subTest(position=position):
                    chart.update_axis(axis_id, tick_label_position=position)
                    axis = next(a for a in chart.get_properties().axes if a.id == axis_id)
                    self.assertIs(axis.tick_label_position, position)
            with self.assertRaises(TypeError):
                chart.update_axis(axis_id, tick_label_position="low")  # type: ignore[arg-type]

    def test_width_model_platform_is_an_enum(self) -> None:
        with Workbook.create_default() as wb:
            for mode in (GeometryMode.DISPLAY, GeometryMode.PRINT):
                self.assertIs(wb.sheet(0).get_width_model(mode).platform, WidthPlatform.WIN)

    def test_table_totals_row_function_takes_and_returns_the_enum(self) -> None:
        with Workbook.create_default() as wb:
            sheet = wb.sheet(0)
            sheet.set_text(0, 0, "Product")
            sheet.set_text(0, 1, "Amount")
            sheet.set_number(1, 1, 10.0)
            sheet.set_number(2, 1, 20.0)
            table = sheet.add_table(
                TableInput(range=CellRect(0, 0, 3, 1), name="Sales", columns=["Product", "Amount"], totals_row=True)
            )
            table.set_column_totals(1, TableColumnTotalsInput(totals_row_function=TableTotalsRowFunction.SUM))
            with self.assertRaises(TypeError):
                table.set_column_totals(1, TableColumnTotalsInput(totals_row_function="sum"))  # type: ignore[arg-type]
            with self.assertRaises(FormulonError):
                table.set_column_totals(1, TableColumnTotalsInput(totals_row_function=99))  # type: ignore[arg-type]
            saved = wb.save()
        with Workbook.load(saved) as reloaded:
            columns = reloaded.sheet(0).list_tables()[0].get_properties().columns
            self.assertIs(columns[1].totals_row_function, TableTotalsRowFunction.SUM)
            self.assertEqual(int(columns[1].totals_row_function), 1)

    def test_table_style_element_type_round_trips_as_the_enum(self) -> None:
        with Workbook.create_default() as wb:
            dxf = wb.ooxml.add_dxf(DxfInput())
            wb.set_table_style(
                TableStyle(
                    name="Corporate",
                    pivot=False,
                    elements=[
                        TableStyleElement(TableStyleElementType.WHOLE_TABLE, dxf_index=dxf),
                        TableStyleElement(TableStyleElementType.FIRST_ROW_STRIPE, size=2, dxf_index=dxf),
                        TableStyleElement(TableStyleElementType.TOTAL_ROW),
                    ],
                )
            )
            saved = wb.save()
        with Workbook.load(saved) as reloaded:
            elements = reloaded.list_table_styles()[0].elements
            self.assertEqual(
                [e.type for e in elements],
                [
                    TableStyleElementType.WHOLE_TABLE,
                    TableStyleElementType.FIRST_ROW_STRIPE,
                    TableStyleElementType.TOTAL_ROW,
                ],
            )
            self.assertTrue(all(isinstance(e.type, TableStyleElementType) for e in elements))

    def test_custom_icon_level_takes_the_icon_set_name_table(self) -> None:
        icons = CfIconSetInput(
            IconSetName.THREE_ARROWS,
            [CfValueObject(1, "33"), CfValueObject(1, "67")],
            icons=[
                CfIcon(IconSetName.THREE_STARS, 0),
                CfIcon(IconSetName.FIVE_BOXES, 2),
                CfIcon(IconSetName.NO_ICONS, 0),
            ],
        )
        with Workbook.create_default() as wb:
            wb.sheet(0).add_cf_rule(CfIconSetRuleInput(ranges=[CellRect(0, 0, 9, 0)], icon_set=icons))
            saved = wb.save()
        with Workbook.load(saved) as reloaded:
            got = reloaded.sheet(0).list_cf_rules()[0].icon_set
            self.assertEqual(
                [i.icon_set for i in got.icons],
                [IconSetName.THREE_STARS, IconSetName.FIVE_BOXES, IconSetName.NO_ICONS],
            )
            self.assertEqual([i.icon_index for i in got.icons], [0, 2, 0])

    def test_phonetic_properties_take_and_return_enums(self) -> None:
        with Workbook.create_default() as wb:
            sheet = wb.sheet(0)
            sheet.set_text(0, 0, "x")
            sheet.set_cell_phonetic_properties(
                0, 0, PhoneticProperties(font_index=0, type=PhoneticType.HIRAGANA, alignment=PhoneticAlignment.CENTER)
            )
            props = sheet.get_cell_phonetic_properties(0, 0)
            self.assertIs(props.type, PhoneticType.HIRAGANA)
            self.assertIs(props.alignment, PhoneticAlignment.CENTER)
            self.assertTrue(props.stated)
            sheet.set_cell_phonetic_properties(0, 0, PhoneticProperties(font_index=0, type=3, alignment=1))
            props = sheet.get_cell_phonetic_properties(0, 0)
            self.assertIs(props.type, PhoneticType.NO_CONVERSION)
            self.assertIs(props.alignment, PhoneticAlignment.LEFT)

    def test_pagination_page_order_is_an_enum(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_number(0, 0, 1.0)
            self.assertIs(wb.sheet(0).get_pagination().page_order, PageOrder.DOWN_THEN_OVER)

    def test_pivot_sort_order_round_trips_by_member(self) -> None:
        with Workbook.create_default() as wb:
            sheet = wb.sheet(0)
            sheet.set_text(0, 0, "Region")
            sheet.set_text(0, 1, "Amount")
            sheet.set_text(1, 0, "East")
            sheet.set_number(1, 1, 10.0)
            sheet.set_text(2, 0, "West")
            sheet.set_number(2, 1, 30.0)
            cache = wb.add_pivot_cache()
            cache.set_worksheet_source(formulon.PivotWorksheetSourceInput(range="A1:B3", sheet="Sheet1"))
            cache.add_field("Region")
            cache.add_field("Amount")
            for region, amount in (("East", 10.0), ("West", 30.0)):
                record = cache.add_record()
                cache.set_record_value(record, 0, region)
                cache.set_record_value(record, 1, amount)
            pivot = sheet.pivot(
                sheet.add_pivot(PivotInput(name="P", cache_id=cache.cache_id, anchor=CellAddress(0, 4)))
            )
            pivot.add_field(PivotFieldInput(field_index="Region", axis=PivotAxis.ROW))
            pivot.update_field(0, sort_order=PivotSortOrder.DESCENDING)
            saved = wb.save()
        with Workbook.load(saved) as reloaded:
            sheet = reloaded.sheet(0)
            fields = sheet.pivot(0).list_fields()
            self.assertIs(fields[0].sort_order, PivotSortOrder.DESCENDING)

    def test_theme_script_font_takes_a_font_scheme(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_theme_script_font(ThemeScriptFont(scheme=FontScheme.MAJOR, script="Jpan", typeface="Head Face"))
            wb.set_theme_script_font(ThemeScriptFont(scheme=FontScheme.MINOR, script="Jpan", typeface="Body Face"))
            major = wb.get_theme_script_font(scheme=FontScheme.MAJOR, script="Jpan")
            minor = wb.get_theme_script_font(scheme=FontScheme.MINOR, script="Jpan")
            assert major is not None and minor is not None
            self.assertEqual((major.scheme, major.typeface), (FontScheme.MAJOR, "Head Face"))
            self.assertEqual((minor.scheme, minor.typeface), (FontScheme.MINOR, "Body Face"))
            with self.assertRaises(FormulonError):
                wb.set_theme_script_font(ThemeScriptFont(scheme=FontScheme.NONE, script="Jpan", typeface="x"))

    def test_merged_function_metadata_availability_is_an_enum(self) -> None:
        base = formulon.get_function_metadata("XLOOKUP")
        assert base is not None
        self.assertIsInstance(base.availability, FunctionAvailability)
        merged = formulon.merge_function_metadata(base, {"signature": "XLOOKUP(a, b, c)"}, "en-US")
        self.assertIsInstance(merged.availability, FunctionAvailability)
        self.assertIs(merged.availability, base.availability)


class ValueInputUnionTests(unittest.TestCase):
    def test_threaded_comment_created_takes_and_returns_a_datetime(self) -> None:
        stamp = datetime.datetime(2026, 1, 1, 10, 0, 5)
        with Workbook.create_default() as wb:
            wb.add_person(Person(id=ALICE, display_name="Alice"))
            wb.sheet(0).add_threaded_comment(
                ThreadedCommentInput(id=THREAD, cell=CellAddress(2, 1), person_id=ALICE, created=stamp, text="Check")
            )
            comment = wb.sheet(0).list_threaded_comments()[0]
            self.assertIsInstance(comment.created, datetime.datetime)
            self.assertEqual(comment.created, stamp)

    def test_threaded_comment_created_accepts_a_date_as_midnight(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_person(Person(id=ALICE, display_name="Alice"))
            wb.sheet(0).add_threaded_comment(
                ThreadedCommentInput(
                    id=THREAD, cell=CellAddress(0, 0), person_id=ALICE, created=datetime.date(2026, 3, 4), text="d"
                )
            )
            self.assertEqual(wb.sheet(0).list_threaded_comments()[0].created, datetime.datetime(2026, 3, 4))

    def test_threaded_comment_created_refuses_an_aware_datetime(self) -> None:
        aware = datetime.datetime(2026, 6, 1, 12, 0, 0, tzinfo=datetime.timezone.utc)
        with Workbook.create_default() as wb:
            wb.add_person(Person(id=ALICE, display_name="Alice"))
            with self.assertRaises(ValueError):
                wb.sheet(0).add_threaded_comment(
                    ThreadedCommentInput(id=THREAD, cell=CellAddress(0, 0), person_id=ALICE, created=aware, text="z")
                )
            self.assertEqual(wb.sheet(0).list_threaded_comments(), [])

    def test_pinned_now_round_trips_a_naive_datetime(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_pinned_now(datetime.datetime(2026, 4, 23, 15, 30, 45, 999))
            self.assertEqual(wb.get_pinned_now(), datetime.datetime(2026, 4, 23, 15, 30, 45))
            wb.set_pinned_now(None)
            self.assertIsNone(wb.get_pinned_now())

    def test_cf_range_today_accepts_date_and_datetime(self) -> None:
        inside = datetime.date(2024, 3, 20)
        outside = datetime.date(2024, 4, 2)
        with Workbook.create_default() as wb:
            sheet = wb.sheet(0)
            sheet.set_number(0, 0, _serial(inside))
            sheet.set_number(1, 0, _serial(outside))
            sheet.add_cf_rule(CfTimePeriodRuleInput(ranges=[CellRect(0, 0, 1, 0)], time_period=DatePeriod.THIS_MONTH))
            wb.recalc()

            def matched(today: object) -> set:
                cells = sheet.evaluate_cf_range(CellRect(0, 0, 1, 0), today=today)  # type: ignore[call-arg]
                return {(c.row, c.col) for c in cells if c.matches}

            self.assertEqual(matched(datetime.date(2024, 3, 15)), {(0, 0)})
            self.assertEqual(matched(datetime.datetime(2024, 3, 15, 23, 59, 59)), {(0, 0)})
            self.assertEqual(matched(datetime.date(2024, 4, 15)), {(1, 0)})
            with self.assertRaises(TypeError):
                matched("2024-03-15")

    def test_pivot_cache_values_accept_date_and_datetime(self) -> None:
        with Workbook.create_default() as wb:
            cache = wb.add_pivot_cache()
            cache.add_field("When")
            cache.add_field_shared_item(0, datetime.date(2024, 1, 31))
            cache.add_field_shared_item(0, datetime.datetime(2024, 1, 31, 12, 0, 0))
            record = cache.add_record()
            cache.set_record_value(record, 0, datetime.date(2024, 2, 1))
            items = cache.list_field_shared_items(0)
            self.assertTrue(all(i.is_date for i in items))
            self.assertEqual([i.value for i in items], [45322.0, 45322.5])
            self.assertFalse(hasattr(cache.add_field_shared_item, "date"))
            with self.assertRaises(TypeError):
                cache.add_field_shared_item(0, "x", date=True)  # type: ignore[call-arg]

    def test_pivot_date_group_bounds_accept_dates(self) -> None:
        with Workbook.create_default() as wb:
            sheet = wb.sheet(0)
            cache = wb.add_pivot_cache()
            cache.add_field("When")
            cache.add_field("Amount")
            for day in (datetime.date(2024, 1, 5), datetime.date(2024, 3, 9)):
                record = cache.add_record()
                cache.set_record_value(record, 0, day)
                cache.set_record_value(record, 1, 1.0)
            pivot = sheet.pivot(
                sheet.add_pivot(PivotInput(name="P", cache_id=cache.cache_id, anchor=CellAddress(0, 4)))
            )
            pivot.add_field(PivotFieldInput(field_index="When", axis=PivotAxis.ROW))
            pivot.set_field_date_group(
                0,
                PivotDateGroupInput(
                    periods=PivotDatePeriod.MONTHS,
                    start=datetime.date(2024, 1, 1),
                    end=datetime.datetime(2024, 12, 31, 0, 0, 0),
                ),
            )

    def test_set_error_accepts_an_excel_error(self) -> None:
        with Workbook.create_default() as wb:
            sheet = wb.sheet(0)
            sheet.set_error(0, 0, ErrorCode.DIV0)
            sheet.set_error(1, 0, ExcelError(ErrorCode.NA))
            self.assertEqual(sheet.get_value(0, 0), ExcelError(ErrorCode.DIV0))
            self.assertEqual(sheet.get_value(1, 0), ExcelError(ErrorCode.NA))
            sheet.set_error(2, 0, sheet.get_value(1, 0))
            self.assertEqual(sheet.get_value(2, 0), ExcelError(ErrorCode.NA))
            sheet.set_error("A4", ExcelError(ErrorCode.NUM))
            self.assertEqual(sheet.get_value("A4"), ExcelError(ErrorCode.NUM))

    def test_validate_value_accepts_the_set_value_union(self) -> None:
        with Workbook.create_default() as wb:
            sheet = wb.sheet(0)
            sheet.add_validation(
                DataValidationInput(type=ValidationType.LIST, ranges=[CellRect(0, 3, 9, 3)], formula1='"Open,Closed"')
            )
            ok = sheet.validate_value("D1", "Open")
            self.assertEqual((ok.valid, ok.rule_id), (True, sheet.list_validations()[0].id))
            self.assertFalse(sheet.validate_value("D1", "Pending").valid)
            self.assertFalse(sheet.validate_value(0, 3, 42).valid)
            self.assertFalse(sheet.validate_value(0, 3, True).valid)
            self.assertTrue(sheet.validate_value(0, 3, None).valid)
            self.assertEqual(
                sheet.validate_value(0, 3, Value(kind=ValueKind.TEXT, text="Closed")),
                ValidationOutcome(
                    valid=True, rule_id=sheet.list_validations()[0].id, error_style=formulon.ValidationErrorStyle.STOP
                ),
            )
            self.assertTrue(sheet.validate_value(0, 0, 1.5).valid)


class FormulaLeadingEqualsTests(unittest.TestCase):
    def test_formula_reads_start_with_equals_for_either_input_spelling(self) -> None:
        with Workbook.create_default() as wb:
            sheet = wb.sheet(0)
            sheet.set_formula(0, 0, "=1+2")
            sheet.set_formula(1, 0, "1+2")
            self.assertEqual(sheet.get_formula(0, 0), "=1+2")
            self.assertEqual(sheet.get_formula(1, 0), "=1+2")
            wb.recalc()
            self.assertEqual((sheet.get_value(0, 0), sheet.get_value(1, 0)), (3.0, 3.0))

    def test_cf_and_validation_formulas_read_with_equals_and_accept_both_spellings(self) -> None:
        with Workbook.create_default() as wb:
            sheet = wb.sheet(0)
            for formula in ("5", "=5"):
                sheet.add_cf_rule(
                    CfCellIsRuleInput(
                        ranges=[CellRect(0, 0, 2, 0)], operator=ComparisonOperator.GREATER_THAN, formula1=formula
                    )
                )
                sheet.add_validation(
                    DataValidationInput(
                        type=ValidationType.WHOLE,
                        ranges=[CellRect(0, 1, 2, 1)],
                        operator=ComparisonOperator.GREATER_THAN,
                        formula1=formula,
                    )
                )
            self.assertEqual([r.formula1 for r in sheet.list_cf_rules()], ["=5", "=5"])
            self.assertEqual([v.formula1 for v in sheet.list_validations()], ["=5", "=5"])


class ToNativeTests(unittest.TestCase):
    def test_to_native_is_exported_and_replaces_to_python(self) -> None:
        self.assertIn("to_native", formulon.__all__)
        self.assertTrue(callable(formulon.to_native))
        self.assertFalse(hasattr(Value, "to_python"))

    def test_to_native_maps_each_value_kind(self) -> None:
        cases = [
            (Value(kind=ValueKind.NUMBER, number=3.0), 3.0),
            (Value(kind=ValueKind.TEXT, text="abc"), "abc"),
            (Value(kind=ValueKind.BOOL, boolean=True), True),
            (Value(kind=ValueKind.BLANK), None),
        ]
        for value, expected in cases:
            with self.subTest(kind=value.kind):
                got = formulon.to_native(value)
                self.assertEqual(got, expected)
                self.assertIs(type(got), type(expected))

    def test_to_native_returns_excel_error_and_nested_row_tuples(self) -> None:
        self.assertEqual(formulon.to_native(formulon.evaluate_typed_formula("=1/0")), ExcelError(ErrorCode.DIV0))
        self.assertEqual(formulon.to_native(formulon.evaluate_typed_formula("=NA()")), ExcelError(ErrorCode.NA))
        grid = formulon.to_native(formulon.evaluate_typed_formula("=SEQUENCE(2,2)"))
        self.assertEqual([list(row) for row in grid], [[1.0, 2.0], [3.0, 4.0]])


if __name__ == "__main__":
    unittest.main()
