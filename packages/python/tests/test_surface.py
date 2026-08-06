"""Surface tests for the expanded Formulon Python Workbook API.

These exercise a representative sample of every method group added on top
of the load/edit/recalc/save core: pivots, conditional formats, styles,
matrix edits, merges/comments/hyperlinks/validations, defined names,
partial recalc, precedents/dependents, spill, function catalog, sheet
view/protection, and external links.

Run via ``make python-test`` (or ``python -m unittest`` with
``PYTHONPATH=packages/python``). Uses stdlib ``unittest`` -- no pytest.
"""

from __future__ import annotations

import subprocess
import unittest
from pathlib import Path

import formulon
from formulon import (
    CalcMode,
    CfColor,
    CfValueObject,
    ColorScale,
    ConditionalFormatInput,
    DataBar,
    DataValidationInput,
    DifferentialFormat,
    FillRecord,
    FontRecord,
    IconSet,
    MergeRange,
    PivotAggregation,
    PivotAxis,
    PivotDataFieldSpec,
    PivotFieldSpec,
    PivotReportLayout,
    PivotWorksheetSource,
    SheetProtection,
    ValueKind,
    Workbook,
)
from formulon import _structs as S


class StructLayoutTests(unittest.TestCase):
    """Guards against silent field-reorder breakage in _structs.py."""

    EXPECTED_SIZES = {
        "MERGE_RANGE": 16,
        "HYPERLINK": 24,
        "COMMENT": 16,
        "DATA_VALIDATION": 52,
        "SHEET_PROTECTION": 88,
        "VIEWPORT": 20,
        "CF_MATCH": 72,
        "CF_RULE": 168,
        "PIVOT_CELL": 40,
        "PIVOT_FIELD_SPEC": 20,
        "PIVOT_DATA_FIELD_SPEC": 28,
        "PIVOT_FILTER_SPEC": 56,
        "SPILL_INFO": 20,
        "FUNCTION_METADATA": 24,
        "SHEET_VIEW": 16,
        "COLUMN_LAYOUT": 24,
        "ROW_LAYOUT": 24,
        "CELL_XF": 20,
        "FONT_RECORD": 40,
        "FILL_RECORD": 12,
        "BORDER_RECORD": 48,
        "DXF_RECORD": 128,
        "CELL_STYLE_RECORD": 24,
        "EXTERNAL_LINK_RECORD": 24,
    }

    def test_struct_sizes(self) -> None:
        for name, expected in self.EXPECTED_SIZES.items():
            layout = getattr(S, name)
            self.assertEqual(layout.size, expected, f"{name} size drifted to {layout.size}")

    def test_pivot_cell_value_offset(self) -> None:
        self.assertEqual(S.PIVOT_CELL_VALUE_OFFSET, 8)

    def test_layouts_match_c_header(self) -> None:
        root = Path(__file__).resolve().parents[3]
        result = subprocess.run(
            ["python3", str(root / "tools" / "dev" / "check_binding_drift.py"), "python-struct-layouts"],
            cwd=root,
            text=True,
            capture_output=True,
            check=False,
        )
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)


class SheetStructureTests(unittest.TestCase):
    def test_add_rename_remove_move_sheet(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_sheet("Second")
            wb.add_sheet("Third")
            self.assertEqual(wb.sheet_count(), 3)
            wb.rename_sheet(1, "Renamed")
            self.assertEqual(wb.sheet_name(1), "Renamed")
            wb.move_sheet(2, 0)
            self.assertEqual(wb.sheet_name(0), "Third")
            wb.remove_sheet(0)
            self.assertEqual(wb.sheet_count(), 2)


class MatrixEditTests(unittest.TestCase):
    def test_insert_row_shifts_value_down(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_number(0, 0, 0, 11.0)  # A1
            wb.insert_rows(0, 0, 1)
            wb.recalc()
            # A1 is now blank; the value moved to A2.
            self.assertEqual(wb.get_value(0, 0, 0).kind, ValueKind.BLANK)
            self.assertEqual(wb.get_value(0, 1, 0).to_python(), 11.0)

    def test_delete_col_shifts_value_left(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_number(0, 0, 0, 1.0)  # A1
            wb.set_number(0, 0, 1, 2.0)  # B1
            wb.delete_cols(0, 0, 1)
            wb.recalc()
            # The old B1 collapses into A1.
            self.assertEqual(wb.get_value(0, 0, 0).to_python(), 2.0)


class AdHocArrayEvalTests(unittest.TestCase):
    def test_sequence_matrix_preserves_all_cells(self) -> None:
        with Workbook.create_default() as wb:
            grid = wb.evaluate_formula_array(0, 0, 0, "=SEQUENCE(2,3)")
            self.assertEqual(len(grid), 2)
            self.assertEqual(len(grid[0]), 3)
            # Row-major 1..6.
            self.assertEqual(grid[0][0].to_python(), 1.0)
            self.assertEqual(grid[1][2].to_python(), 6.0)

    def test_scalar_reported_as_one_by_one(self) -> None:
        with Workbook.create_default() as wb:
            grid = wb.evaluate_formula_array(0, 0, 0, "=1+2")
            self.assertEqual(len(grid), 1)
            self.assertEqual(len(grid[0]), 1)
            self.assertEqual(grid[0][0].kind, ValueKind.NUMBER)
            self.assertEqual(grid[0][0].to_python(), 3.0)


class DefinedNameTests(unittest.TestCase):
    def test_defined_name_set_get_remove(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_defined_name("MyRef", "Sheet1!$A$1")
            names = {dn.name: dn.formula for dn in wb.iter_defined_names()}
            self.assertIn("MyRef", names)
            self.assertEqual(names["MyRef"], "Sheet1!$A$1")

            # Updating in place replaces the formula text.
            wb.set_defined_name("MyRef", "Sheet1!$B$2")
            names = {dn.name: dn.formula for dn in wb.iter_defined_names()}
            self.assertEqual(names["MyRef"], "Sheet1!$B$2")

            # An empty formula removes the entry.
            wb.set_defined_name("MyRef", "")
            self.assertNotIn("MyRef", {dn.name for dn in wb.iter_defined_names()})

    def test_scoped_defined_name_set_get_remove(self) -> None:
        with Workbook.create_empty() as wb:
            wb.add_sheet("Sheet1")
            wb.add_sheet("Sheet2")
            wb.set_defined_name("Rate", "=1")
            wb.set_defined_name_scoped("Rate", "=2", 1)
            names = {(dn.name, dn.local_sheet_id): dn.formula for dn in wb.iter_defined_names()}
            self.assertEqual(names[("Rate", -1)], "=1")
            self.assertEqual(names[("Rate", 1)], "=2")

            wb.set_defined_name_scoped("Rate", "", 1)
            self.assertNotIn(
                ("Rate", 1),
                {(dn.name, dn.local_sheet_id) for dn in wb.iter_defined_names()},
            )


class CalcPolicyTests(unittest.TestCase):
    def test_calc_mode_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(wb.calc_mode(), CalcMode.AUTO)
            wb.set_calc_mode(CalcMode.MANUAL)
            self.assertEqual(wb.calc_mode(), CalcMode.MANUAL)

    def test_excel_profile_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            self.assertTrue(wb.excel_profile_id())
            wb.set_excel_profile_id("mac-365-ja_JP")
            self.assertEqual(wb.excel_profile_id(), "mac-365-ja_JP")


class PartialRecalcTests(unittest.TestCase):
    def test_partial_recalc_recomputes_chain(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_number(0, 0, 0, 10.0)  # A1
            wb.set_formula(0, 0, 1, "=A1*2")  # B1
            recomputed = wb.partial_recalc(0, 0, 0, 1, 1)
            self.assertGreater(recomputed, 0)
            self.assertEqual(wb.get_value(0, 0, 1).to_python(), 20.0)


class TraceTests(unittest.TestCase):
    def test_precedents_and_dependents(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_number(0, 0, 0, 1.0)  # A1
            wb.set_formula(0, 0, 1, "=A1")  # B1
            wb.recalc()
            prec = wb.precedents(0, 0, 1, 1)
            self.assertIn((0, 0, 0), [(p.sheet, p.row, p.col) for p in prec])
            deps = wb.dependents(0, 0, 0, 1)
            self.assertIn((0, 0, 1), [(d.sheet, d.row, d.col) for d in deps])


class SpillTests(unittest.TestCase):
    def test_sequence_spills(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_formula(0, 0, 0, "=SEQUENCE(3)")
            wb.recalc()
            info = wb.spill_info(0, 0, 0)
            self.assertTrue(info.engaged)
            self.assertEqual(info.rows, 3)
            self.assertEqual((info.anchor_row, info.anchor_col), (0, 0))


class MergeCommentHyperlinkTests(unittest.TestCase):
    def test_merge_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_merge(0, MergeRange(0, 0, 1, 1))
            merges = wb.get_merges(0)
            self.assertEqual(len(merges), 1)
            self.assertEqual((merges[0].first_row, merges[0].last_col), (0, 1))
            wb.clear_merges(0)
            self.assertEqual(wb.get_merges(0), [])

    def test_comment_roundtrip_and_absent(self) -> None:
        with Workbook.create_default() as wb:
            self.assertIsNone(wb.get_comment(0, 0, 0))
            wb.set_comment(0, 0, 0, "alice", "see note")
            c = wb.get_comment(0, 0, 0)
            self.assertIsNotNone(c)
            self.assertEqual(c.author, "alice")
            self.assertEqual(c.text, "see note")

    def test_comment_invalid_sheet_raises_instead_of_looking_absent(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(formulon.FormulonError) as ctx:
                wb.get_comment(99, 0, 0)
            self.assertEqual(ctx.exception.status, 2)  # kInvalidArgument

    def test_hyperlink_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_hyperlink(0, 0, 0, "https://example.com", "Example", "tip")
            links = wb.get_hyperlinks(0)
            self.assertEqual(len(links), 1)
            self.assertEqual(links[0].target, "https://example.com")
            self.assertEqual(links[0].display, "Example")


class ValidationTests(unittest.TestCase):
    def test_list_validation_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_validation(
                0,
                DataValidationInput(
                    type=3,  # list
                    ranges=[MergeRange(0, 0, 4, 0)],
                    allow_blank=True,
                    formula1='"a,b,c"',
                ),
            )
            self.assertEqual(wb.validation_count(0), 1)
            dv = wb.get_validation_at(0, 0)
            self.assertEqual(dv.type, 3)
            self.assertTrue(dv.allow_blank)
            self.assertEqual(dv.formula1, '"a,b,c"')
            self.assertEqual(len(dv.ranges), 1)
            wb.clear_validations(0)
            self.assertEqual(wb.validation_count(0), 0)


class SheetViewProtectionTests(unittest.TestCase):
    def test_view_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_sheet_zoom(0, 150)
            wb.set_sheet_freeze(0, 1, 2)
            view = wb.get_sheet_view(0)
            self.assertEqual(view.zoom_scale, 150)
            self.assertEqual((view.freeze_rows, view.freeze_cols), (1, 2))

    def test_view_display_defaults(self) -> None:
        with Workbook.create_default() as wb:
            view = wb.get_sheet_view(0)
            self.assertTrue(view.show_grid_lines)
            self.assertTrue(view.show_row_col_headers)
            self.assertTrue(view.show_zeros)
            self.assertFalse(view.right_to_left)
            self.assertFalse(view.tab_selected)
            self.assertEqual(view.view_mode, "")

    def test_view_display_setters_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_sheet_show_grid_lines(0, False)
            wb.set_sheet_show_row_col_headers(0, False)
            wb.set_sheet_show_zeros(0, False)
            wb.set_sheet_right_to_left(0, True)
            wb.set_sheet_tab_selected(0, True)
            wb.set_sheet_view_mode(0, "pageBreakPreview")
            view = wb.get_sheet_view(0)
            self.assertFalse(view.show_grid_lines)
            self.assertFalse(view.show_row_col_headers)
            self.assertFalse(view.show_zeros)
            self.assertTrue(view.right_to_left)
            self.assertTrue(view.tab_selected)
            self.assertEqual(view.view_mode, "pageBreakPreview")
            # An empty mode is a meaningful value (the OOXML-default
            # "normal" view), not a no-op -- it must round-trip too.
            wb.set_sheet_view_mode(0, "")
            self.assertEqual(wb.get_sheet_view(0).view_mode, "")

    def test_column_row_overrides(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_column_width(0, 0, 2, 18.5)
            cols = wb.get_sheet_columns(0)
            self.assertTrue(any(abs(c.width - 18.5) < 1e-9 for c in cols))
            wb.set_row_height(0, 0, 30.0)
            rows = wb.get_sheet_row_overrides(0)
            self.assertTrue(any(abs(r.height - 30.0) < 1e-9 for r in rows))

    def test_protection_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            prot = SheetProtection(enabled=True, sheet=True, format_cells=True)
            wb.set_sheet_protection(0, prot)
            got = wb.get_sheet_protection(0)
            self.assertTrue(got.enabled)
            self.assertTrue(got.sheet)
            self.assertTrue(got.format_cells)


class StyleTests(unittest.TestCase):
    def test_font_numfmt_xf_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            fi = wb.add_font(FontRecord(name="Calibri", size=12.0, bold=True))
            self.assertGreaterEqual(fi, 0)
            font = wb.get_font(fi)
            self.assertEqual(font.name, "Calibri")
            self.assertTrue(font.bold)

            fill = wb.add_fill(FillRecord(pattern=1, fg_argb=0xFFFF0000))
            nf = wb.add_num_fmt("0.00")
            self.assertGreater(nf, 0)
            self.assertEqual(wb.get_num_fmt(nf), "0.00")

            # add_cell_xf validates indices against the parallel tables, so a
            # border must exist before it can be referenced.
            border = wb.add_border({"left": {"style": 1, "color_argb": 0xFF000000}})

            from formulon import CellXf

            xf = wb.add_cell_xf(
                CellXf(
                    font_index=fi,
                    fill_index=fill,
                    border_index=border,
                    num_fmt_id=nf,
                    horizontal_align=0,
                    vertical_align=0,
                    wrap_text=False,
                )
            )
            wb.set_cell_xf_index(0, 0, 0, xf)
            self.assertEqual(wb.get_cell_xf_index(0, 0, 0), xf)
            resolved = wb.get_cell_xf(xf)
            self.assertEqual(resolved.font_index, fi)
            self.assertEqual(resolved.num_fmt_id, nf)

    def test_dxf_roundtrip_and_dedup(self) -> None:
        with Workbook.create_default() as wb:
            record = DifferentialFormat(
                font=FontRecord(name="Arial", size=12.0, bold=True, color_argb=0xFFFF0000),
                fill=FillRecord(pattern=1, fg_argb=0xFFFFFF00),
                num_fmt_id=164,
                num_fmt_code="0.00",
            )
            index = wb.add_dxf(record)
            self.assertEqual(wb.add_dxf(record), index)
            self.assertEqual(wb.dxf_count(), 1)
            got = wb.get_dxf(index)
            self.assertIsNotNone(got.font)
            self.assertEqual(got.font.name, "Arial")
            self.assertTrue(got.font.bold)
            self.assertIsNotNone(got.fill)
            self.assertEqual(got.fill.fg_argb, 0xFFFFFF00)
            self.assertEqual(got.num_fmt_id, 164)
            self.assertEqual(got.num_fmt_code, "0.00")


class ConditionalFormatTests(unittest.TestCase):
    def test_cf_add_get_evaluate_clear(self) -> None:
        with Workbook.create_default() as wb:
            for r in range(3):
                wb.set_number(0, r, 0, float(r * 5))  # A1=0, A2=5, A3=10
            wb.add_conditional_format(
                0,
                ConditionalFormatInput(
                    sqref=[MergeRange(0, 0, 2, 0)],
                    type=1,  # cellIs
                    op_engaged=True,
                    op=4,  # greaterThan
                    formula1="4",
                    dxf_id_engaged=True,
                    dxf_id=0,
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


class PivotTests(unittest.TestCase):
    def test_build_and_project_pivot(self) -> None:
        with Workbook.create_default() as wb:
            cache_id = wb.pivot_cache_create()
            wb.pivot_cache_field_add(cache_id, "Region")
            wb.pivot_cache_field_add(cache_id, "Amount")
            for region, amount in [("East", 10.0), ("East", 20.0), ("West", 30.0)]:
                rec = wb.pivot_cache_record_add(cache_id)
                wb.pivot_cache_record_set_text(cache_id, rec, 0, region)
                wb.pivot_cache_record_set_number(cache_id, rec, 1, amount)
            self.assertEqual(wb.pivot_cache_record_count(cache_id), 3)

            pivot = wb.pivot_create(0, "Pivot1", cache_id, 0, 4)
            region_field = wb.pivot_field_add(0, pivot, PivotFieldSpec(source_name="Region", axis=PivotAxis.ROW))
            amount_field = wb.pivot_field_add(0, pivot, PivotFieldSpec(source_name="Amount", axis=PivotAxis.VALUE))
            self.assertEqual(region_field, 0)
            wb.pivot_data_field_add(
                0,
                pivot,
                PivotDataFieldSpec(
                    name="Sum of Amount",
                    field_index=amount_field,
                    aggregation=PivotAggregation.SUM,
                ),
            )

            layout = wb.pivot_layout(0, pivot)
            self.assertGreater(len(layout.cells), 0)
            numbers = [c.value.to_python() for c in layout.cells if c.value.kind == ValueKind.NUMBER]
            # The single grand total of 10 + 20 + 30 must appear.
            self.assertIn(60.0, numbers)

    def test_cache_source_and_report_layout_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            cache = wb.pivot_cache_create()
            wb.set_pivot_cache_worksheet_source(cache, PivotWorksheetSource(ref="A1:B9", sheet="Sheet1"))
            self.assertEqual(wb.get_pivot_cache_worksheet_source(cache).sheet, "Sheet1")
            pivot = wb.pivot_create(0, "Layout", cache, 0, 0)
            wb.set_pivot_report_layout(0, pivot, PivotReportLayout.TABULAR)
            self.assertEqual(wb.get_pivot_report_layout(0, pivot), PivotReportLayout.TABULAR)
            wb.set_pivot_cache_worksheet_source(cache, None)
            self.assertIsNone(wb.get_pivot_cache_worksheet_source(cache))


class FunctionCatalogTests(unittest.TestCase):
    def test_catalog_metadata(self) -> None:
        self.assertGreater(Workbook.function_count(), 0)
        meta = Workbook.function_metadata("SUM", 0)
        self.assertIsNotNone(meta)
        self.assertEqual(meta.name, "SUM")
        self.assertGreaterEqual(meta.min_arity, 1)
        # SUM is an unbounded variadic; the sentinel is normalized to None.
        self.assertIsNone(meta.max_arity)
        # Lazy-dispatch forms (not in the eager registry) still resolve.
        xlookup = Workbook.function_metadata("XLOOKUP", 0)
        self.assertIsNotNone(xlookup)
        self.assertEqual(xlookup.name, "XLOOKUP")
        names = {Workbook.function_name_at(i) for i in range(Workbook.function_count())}
        self.assertIn("XLOOKUP", names)
        self.assertIsNone(Workbook.function_metadata("NOT_A_REAL_FUNCTION", 0))

    def test_function_name_at(self) -> None:
        name = Workbook.function_name_at(0)
        self.assertIsInstance(name, str)
        self.assertGreater(len(name), 0)

    def test_merge_function_metadata(self) -> None:
        base = Workbook.function_metadata("XLOOKUP", 0)
        self.assertIsNotNone(base)
        # The engine leaves display metadata empty.
        self.assertIsNone(base.signature_template)
        self.assertIsNone(base.description)

        entry = {
            "signature": "XLOOKUP(lookup_value, lookup_array, return_array)",
            "description": "Searches a range or an array.",
            "aliases": {"fr-FR": "RECHERCHEX"},
            "localized": {"fr-FR": {"signature": "RECHERCHEX(...)", "description": "Recherche."}},
        }

        # Localized override wins for the matching locale.
        fr = formulon.merge_function_metadata(base, entry, "fr-FR")
        self.assertEqual(fr.signature_template, "RECHERCHEX(...)")
        self.assertEqual(fr.description, "Recherche.")
        self.assertEqual(fr.localized_name, "RECHERCHEX")
        # Structural fields survive the merge.
        self.assertEqual(fr.name, "XLOOKUP")

        # A locale with no localized/alias entry falls back to the default
        # signature/description and the canonical display name.
        de = formulon.merge_function_metadata(base, entry, "de-DE")
        self.assertEqual(
            de.signature_template,
            "XLOOKUP(lookup_value, lookup_array, return_array)",
        )
        self.assertEqual(de.description, "Searches a range or an array.")
        self.assertEqual(de.localized_name, "XLOOKUP")

        # No provider entry -> base returned verbatim; display metadata NULL.
        none = formulon.merge_function_metadata(base, None, "fr-FR")
        self.assertIs(none, base)
        self.assertIsNone(none.signature_template)
        self.assertIsNone(none.description)


class ExternalLinkTests(unittest.TestCase):
    def test_fresh_workbook_has_no_external_links(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(wb.external_link_count(), 0)
            self.assertEqual(wb.get_external_links(), [])


class SurfaceParityTests(unittest.TestCase):
    """Every advertised method must exist on the Workbook class."""

    METHODS = [
        "move_sheet",
        "remove_sheet",
        "rename_sheet",
        "set_defined_name",
        "set_defined_name_scoped",
        "set_error",
        "insert_rows",
        "delete_rows",
        "insert_cols",
        "delete_cols",
        "calc_mode",
        "set_calc_mode",
        "excel_profile_id",
        "set_excel_profile_id",
        "partial_recalc",
        "lambda_text_at",
        "evaluate_formula_array",
        "add_merge",
        "remove_merge",
        "remove_merge_at",
        "clear_merges",
        "get_merges",
        "merge_count",
        "add_hyperlink",
        "remove_hyperlink",
        "remove_hyperlink_at",
        "clear_hyperlinks",
        "get_hyperlinks",
        "hyperlink_count",
        "get_comment",
        "set_comment",
        "validation_count",
        "get_validation_at",
        "get_validations",
        "add_validation",
        "remove_validation_at",
        "clear_validations",
        "get_sheet_protection",
        "set_sheet_protection",
        "get_sheet_view",
        "set_sheet_zoom",
        "set_sheet_freeze",
        "set_sheet_tab_hidden",
        "set_sheet_show_grid_lines",
        "set_sheet_show_row_col_headers",
        "set_sheet_show_zeros",
        "set_sheet_right_to_left",
        "set_sheet_tab_selected",
        "set_sheet_view_mode",
        "get_sheet_columns",
        "set_column_width",
        "set_column_hidden",
        "set_column_outline",
        "get_sheet_row_overrides",
        "set_row_height",
        "set_row_hidden",
        "set_row_outline",
        "evaluate_cf_range",
        "cf_count",
        "get_conditional_format_at",
        "get_conditional_formats",
        "add_conditional_format",
        "remove_conditional_format_at",
        "clear_conditional_formats",
        "get_cell_xf_index",
        "set_cell_xf_index",
        "get_cell_xf",
        "get_font",
        "get_fill",
        "get_border",
        "get_num_fmt",
        "font_count",
        "fill_count",
        "border_count",
        "cell_xf_count",
        "cell_style_count",
        "cell_style_xf_count",
        "add_font",
        "add_fill",
        "add_border",
        "add_num_fmt",
        "add_cell_xf",
        "get_cell_style",
        "get_cell_style_xf",
        "pivot_count",
        "pivot_layout",
        "pivot_cache_count",
        "pivot_cache_id_at",
        "pivot_cache_create",
        "pivot_cache_remove",
        "pivot_cache_field_count",
        "pivot_cache_field_name",
        "pivot_cache_field_add",
        "pivot_cache_field_clear",
        "pivot_cache_record_add",
        "pivot_cache_record_set_number",
        "pivot_cache_record_set_text",
        "pivot_create",
        "pivot_remove",
        "pivot_set_name",
        "pivot_set_anchor",
        "pivot_set_grand_totals",
        "pivot_field_add",
        "pivot_field_set_axis",
        "pivot_data_field_add",
        "pivot_data_field_set",
        "pivot_filter_add",
        "pivot_set_row_field_order",
        "pivot_set_col_field_order",
        "precedents",
        "dependents",
        "spill_info",
        "function_count",
        "function_name_at",
        "function_metadata",
        "localize_function_name",
        "canonicalize_function_name",
        "external_link_count",
        "get_external_link_at",
        "get_external_links",
    ]

    def test_all_methods_present(self) -> None:
        for name in self.METHODS:
            self.assertTrue(hasattr(Workbook, name), f"Workbook missing method: {name}")

    def test_public_types_exported(self) -> None:
        for name in (
            "MergeRange",
            "Comment",
            "Hyperlink",
            "DataValidation",
            "DataValidationInput",
            "SheetProtection",
            "SheetView",
            "ConditionalFormat",
            "ConditionalFormatInput",
            "CfMatch",
            "CfCellResult",
            "CellNode",
            "SpillInfo",
            "FunctionMetadata",
            "CellXf",
            "FontRecord",
            "FillRecord",
            "CellStyle",
            "ExternalLink",
            "PivotCell",
            "PivotLayout",
            "PivotFieldSpec",
            "PivotDataFieldSpec",
            "PivotFilterSpec",
            "CalcMode",
            "PivotAxis",
            "PivotAggregation",
            "MergedFunctionMetadata",
            "merge_function_metadata",
        ):
            self.assertTrue(hasattr(formulon, name), f"formulon missing: {name}")


if __name__ == "__main__":
    unittest.main()
