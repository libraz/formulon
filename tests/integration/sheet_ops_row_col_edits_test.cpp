// Structural workbook mutation tests grouped by public surface.
#include "sheet_ops_test_support.h"

namespace formulon {
namespace {
using namespace sheet_ops_test;

TEST(WorkbookSheetOps, InsertRowsRekeysAllDependenciesBeforeRegisteringNewCoordinates) {
  Workbook wb = Workbook::create();
  constexpr std::uint32_t kRows = 60U;
  for (std::uint32_t row = 0; row < kRows; ++row) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, row, 0U, Value::number(static_cast<double>(row)))));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, row, 1U, "=A" + std::to_string(row + 1U))));
  }
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0U, 0U, 1U)));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 1U, 0U, Value::number(999.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  for (std::uint32_t row = 0; row < kRows; ++row) {
    const Cell* dependent = wb.sheet(0).cell_at(row + 1U, 1U);
    ASSERT_NE(dependent, nullptr) << "missing formula at row " << row + 2U;
    ASSERT_TRUE(dependent->cached_value.is_number());
    const double expected = (row == 0U) ? 999.0 : static_cast<double>(row);
    EXPECT_DOUBLE_EQ(dependent->cached_value.as_number(), expected) << "stale dependent at row " << row + 2U;
  }
}

TEST(WorkbookSheetOps, ConsecutiveRowInsertsKeepShiftedDependenciesLive) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 0, Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 1, 0, Value::number(2.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 2, 1, "=SUM(A1:A2)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_DOUBLE_EQ(wb.sheet(0).cell_at(2, 1)->cached_value.as_number(), 3.0);

  // Inserting inside the range expands it; the original second input moves
  // to A3 and the dependent moves to B4.
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, 1, 1)));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 1, 0, Value::number(5.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_NE(wb.sheet(0).cell_at(3, 1), nullptr);
  EXPECT_EQ(wb.sheet(0).cell_at(3, 1)->formula_text, "=SUM(A1:A3)");
  EXPECT_DOUBLE_EQ(wb.sheet(0).cell_at(3, 1)->cached_value.as_number(), 8.0);

  // A second insert must re-key the already-shifted dependency. Updating
  // the moved original A1 (now A2) must still invalidate the formula B5.
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, 0, 1)));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 1, 0, Value::number(10.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_NE(wb.sheet(0).cell_at(4, 1), nullptr);
  EXPECT_EQ(wb.sheet(0).cell_at(4, 1)->formula_text, "=SUM(A2:A4)");
  EXPECT_DOUBLE_EQ(wb.sheet(0).cell_at(4, 1)->cached_value.as_number(), 17.0);
}

TEST(WorkbookRowColEdits, InsertRowsShiftsCellsAndFormulas) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 0, Value::number(10.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 5, 0, Value::number(20.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 6, 0, "=A1+A6")));

  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, /*row=*/3, /*count=*/2)));

  // Cells at row 0 stayed put; cell at row 5 moved to row 7; the formula
  // cell at row 6 moved to row 8 and its A6 ref shifted to A8.
  const Sheet& s = wb.sheet(0);
  ASSERT_NE(s.cell_at(0, 0), nullptr);
  EXPECT_EQ(s.cell_at(0, 0)->cached_value.as_number(), 10.0);
  EXPECT_EQ(s.cell_at(5, 0), nullptr);
  ASSERT_NE(s.cell_at(7, 0), nullptr);
  EXPECT_EQ(s.cell_at(7, 0)->cached_value.as_number(), 20.0);
  ASSERT_NE(s.cell_at(8, 0), nullptr);
  EXPECT_EQ(s.cell_at(8, 0)->formula_text, "=A1+A8");
}

TEST(WorkbookRowColEdits, InsertRowsShiftsConditionalFormatsLayoutAndBreaks) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  cf::ConditionalFormat format;
  format.sqref.push_back(cf::CFCellRange{CellAddress{1, 0}, CellAddress{2, 1}});
  cf::CFRule rule;
  rule.formula1 = "A2";
  format.rules.push_back(std::move(rule));
  sheet.mutable_conditional_formats().push_back(std::move(format));
  sheet.mutable_layout().row_overrides.push_back(RowLayout{2, 24.0, true, 1, true});
  sheet.mutable_layout().columns.push_back(ColumnLayout{1, 2, 18.0, true, 2});
  sheet.mutable_print_settings().manual_row_breaks.push_back(ManualBreak{3, 0, 10, true});
  sheet.mutable_print_settings().manual_col_breaks.push_back(ManualBreak{2, 0, 10, true});
  TableMetadata table;
  table.id = 1;
  table.name = "Table1";
  table.display_name = "Table1";
  table.sheet_index = 0;
  table.ref = "A2:B3";
  table.columns.push_back(TableColumn{1, "Value", {}, {}, "A2"});
  wb.set_tables({table});
  auto pivot = std::make_unique<pivot::PivotTable>();
  pivot->set_anchor(2, 1, 3, 2);
  sheet.add_pivot_table(std::move(pivot));
  ASSERT_TRUE(sheet.commit_spill(4, 4, 2, 1, std::vector<Value>{Value::number(1), Value::number(2)}));

  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, 1, 1)));
  const cf::CFCellRange row_shifted = sheet.conditional_formats()[0].sqref[0];
  EXPECT_EQ(row_shifted.first.row, 2U);
  EXPECT_EQ(row_shifted.last.row, 3U);
  ASSERT_TRUE(sheet.conditional_formats()[0].rules[0].formula1.has_value());
  EXPECT_EQ(*sheet.conditional_formats()[0].rules[0].formula1, "A3");
  EXPECT_EQ(sheet.layout().row_overrides[0].row, 3U);
  EXPECT_EQ(sheet.print_settings().manual_row_breaks[0].id, 4U);
  ASSERT_EQ(sheet.pivot_tables().size(), 1U);
  EXPECT_EQ(sheet.pivot_tables()[0]->anchor_row(), 3U);
  EXPECT_EQ(sheet.spill_region_at_anchor(5, 4), nullptr);
  ASSERT_EQ(wb.tables().size(), 1U);
  EXPECT_EQ(wb.tables()[0].ref, "A3:B4");
  EXPECT_EQ(wb.tables()[0].columns[0].calculated_column_formula, "A3");
  // Column-bound metadata is not touched by a row edit.
  EXPECT_EQ(sheet.layout().columns[0].first, 1U);
  EXPECT_EQ(sheet.print_settings().manual_col_breaks[0].id, 2U);

  ASSERT_TRUE(static_cast<bool>(wb.insert_cols(0, 0, 1)));
  const cf::CFCellRange col_shifted = sheet.conditional_formats()[0].sqref[0];
  EXPECT_EQ(col_shifted.first.col, 1U);
  EXPECT_EQ(col_shifted.last.col, 2U);
  EXPECT_EQ(sheet.layout().columns[0].first, 2U);
  EXPECT_EQ(sheet.layout().columns[0].last, 3U);
  EXPECT_EQ(sheet.print_settings().manual_col_breaks[0].id, 3U);
  EXPECT_EQ(sheet.pivot_tables()[0]->anchor_col(), 2U);
  EXPECT_EQ(wb.tables()[0].ref, "B3:C4");
}

TEST(WorkbookRowColEdits, DeleteRowsRemovesConditionalFormatWhoseRangeIsDeleted) {
  Workbook wb = Workbook::create();
  cf::ConditionalFormat format;
  format.sqref.push_back(cf::CFCellRange{CellAddress{1, 0}, CellAddress{2, 1}});
  cf::CFRule rule;
  rule.formula1 = "A2";
  format.rules.push_back(std::move(rule));
  wb.sheet(0).mutable_conditional_formats().push_back(std::move(format));

  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, 1, 2)));
  EXPECT_TRUE(wb.sheet(0).conditional_formats().empty());
}

TEST(WorkbookRowColEdits, DeleteRowsCollapsesReferencesInsideInterval) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 0, "=A5")));     // A5 lives inside the deletion.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 1, "=A8+A2")));  // A8 trails the deletion; A2 is below.

  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, /*row=*/3, /*count=*/3)));

  const Sheet& s = wb.sheet(0);
  // A5 falls inside [3, 6); collapses to #REF!.
  EXPECT_EQ(s.cell_at(0, 0)->formula_text, "=#REF!");
  // A8 -> A5 (shifted up by 3); A2 unchanged.
  EXPECT_EQ(s.cell_at(0, 1)->formula_text, "=A5+A2");
}

TEST(WorkbookRowColEdits, DeleteRowsShrinksRangeAtFirstLastAndMiddle) {
  {
    Workbook wb = Workbook::create();
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 10, 0, "=SUM(A1:A3)")));
    ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, /*row=*/0, /*count=*/1)));
    ASSERT_NE(wb.sheet(0).cell_at(9, 0), nullptr);
    EXPECT_EQ(wb.sheet(0).cell_at(9, 0)->formula_text, "=SUM(A1:A2)");
  }
  {
    Workbook wb = Workbook::create();
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 10, 0, "=SUM(A1:A3)")));
    ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, /*row=*/2, /*count=*/1)));
    ASSERT_NE(wb.sheet(0).cell_at(9, 0), nullptr);
    EXPECT_EQ(wb.sheet(0).cell_at(9, 0)->formula_text, "=SUM(A1:A2)");
  }
  {
    Workbook wb = Workbook::create();
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 10, 0, "=SUM(A1:A5)")));
    ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, /*row=*/2, /*count=*/1)));
    ASSERT_NE(wb.sheet(0).cell_at(9, 0), nullptr);
    EXPECT_EQ(wb.sheet(0).cell_at(9, 0)->formula_text, "=SUM(A1:A4)");
  }
}

TEST(WorkbookRowColEdits, DeleteRowsCollapsesSingletonAndShrinksWholeRowRange) {
  {
    Workbook wb = Workbook::create();
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 10, 0, "=SUM(A2:A2)")));
    ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, /*row=*/1, /*count=*/1)));
    ASSERT_NE(wb.sheet(0).cell_at(9, 0), nullptr);
    EXPECT_EQ(wb.sheet(0).cell_at(9, 0)->formula_text, "=SUM(#REF!)");
  }
  {
    Workbook wb = Workbook::create();
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 10, 0, "=SUM(1:4)")));
    ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, /*row=*/1, /*count=*/1)));
    ASSERT_NE(wb.sheet(0).cell_at(9, 0), nullptr);
    EXPECT_EQ(wb.sheet(0).cell_at(9, 0)->formula_text, "=SUM(1:3)");
  }
}

TEST(WorkbookRowColEdits, DeleteRowsCollapsesFullyDeletedMultiCellRange) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 10, 0, "=SUM(A2:A4)")));
  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, /*row=*/1, /*count=*/3)));
  ASSERT_NE(wb.sheet(0).cell_at(7, 0), nullptr);
  EXPECT_EQ(wb.sheet(0).cell_at(7, 0)->formula_text, "=SUM(#REF!)");
}

TEST(WorkbookRowColEdits, DeleteRowsShrunkRangeRecalculatesAfterSurvivorChange) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 0, Value::number(10.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 1, 0, Value::number(20.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 2, 0, Value::number(30.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 3, 0, Value::number(40.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 6, 3, "=SUM(A1:A4)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_EQ(wb.sheet(0).cell_at(6, 3)->cached_value.as_number(), 100.0);

  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, /*row=*/0, /*count=*/1)));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_NE(wb.sheet(0).cell_at(5, 3), nullptr);
  EXPECT_EQ(wb.sheet(0).cell_at(5, 3)->formula_text, "=SUM(A1:A3)");
  EXPECT_EQ(wb.sheet(0).cell_at(5, 3)->cached_value.as_number(), 90.0);

  // The surviving dependency at A2 must still point at the moved formula;
  // changing it and recalculating proves the structural edit rebuilt the
  // dependency graph rather than merely changing formula text.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 1, 0, Value::number(50.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_EQ(wb.sheet(0).cell_at(5, 3)->cached_value.as_number(), 110.0);
}

TEST(WorkbookRowColEdits, InsertColsShiftsCellsAcrossRow) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 0, Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 3, Value::number(4.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 1, 0, "=A1+D1")));

  ASSERT_TRUE(static_cast<bool>(wb.insert_cols(0, /*col=*/1, /*count=*/2)));

  const Sheet& s = wb.sheet(0);
  EXPECT_EQ(s.cell_at(0, 0)->cached_value.as_number(), 1.0);
  // Slot at (0, 3) exists in the row vector but is a default Cell — the
  // value that used to live there migrated forward.
  ASSERT_NE(s.cell_at(0, 3), nullptr);
  EXPECT_TRUE(s.cell_at(0, 3)->cached_value.is_blank());
  ASSERT_NE(s.cell_at(0, 5), nullptr);
  EXPECT_EQ(s.cell_at(0, 5)->cached_value.as_number(), 4.0);
  // The formula cell sat at (1, 0), which is below the insert origin
  // (col 1) and so does not move. Its references shift: D1 -> F1.
  ASSERT_NE(s.cell_at(1, 0), nullptr);
  EXPECT_EQ(s.cell_at(1, 0)->formula_text, "=A1+F1");
}

TEST(WorkbookRowColEdits, DeleteColsRewritesFormulaAndDropsCell) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 5, "=B1+C1+E1")));

  // Delete cols B..C (col indices 1..2).
  ASSERT_TRUE(static_cast<bool>(wb.delete_cols(0, /*col=*/1, /*count=*/2)));

  const Sheet& s = wb.sheet(0);
  // Source cell moved from F1 (col 5) to D1 (col 3).
  ASSERT_NE(s.cell_at(0, 3), nullptr);
  // B1 and C1 are inside the deleted span: collapse the entire formula
  // to #REF! because every range / sum endpoint that lands inside the
  // deletion poisons its enclosing expression.
  EXPECT_EQ(s.cell_at(0, 3)->formula_text, "=#REF!+#REF!+C1");
}

TEST(WorkbookRowColEdits, DeleteColsShrinksRangeAtFirstLastAndMiddle) {
  {
    Workbook wb = Workbook::create();
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 10, "=SUM(A1:D1)")));
    ASSERT_TRUE(static_cast<bool>(wb.delete_cols(0, /*col=*/0, /*count=*/1)));
    ASSERT_NE(wb.sheet(0).cell_at(0, 9), nullptr);
    EXPECT_EQ(wb.sheet(0).cell_at(0, 9)->formula_text, "=SUM(A1:C1)");
  }
  {
    Workbook wb = Workbook::create();
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 10, "=SUM(A1:D1)")));
    ASSERT_TRUE(static_cast<bool>(wb.delete_cols(0, /*col=*/3, /*count=*/1)));
    ASSERT_NE(wb.sheet(0).cell_at(0, 9), nullptr);
    EXPECT_EQ(wb.sheet(0).cell_at(0, 9)->formula_text, "=SUM(A1:C1)");
  }
  {
    Workbook wb = Workbook::create();
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 10, "=SUM(A1:F1)")));
    ASSERT_TRUE(static_cast<bool>(wb.delete_cols(0, /*col=*/2, /*count=*/2)));
    ASSERT_NE(wb.sheet(0).cell_at(0, 8), nullptr);
    EXPECT_EQ(wb.sheet(0).cell_at(0, 8)->formula_text, "=SUM(A1:D1)");
  }
}

TEST(WorkbookRowColEdits, DeleteColsCollapsesSingletonAndShrinksWholeColumnRange) {
  {
    Workbook wb = Workbook::create();
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 10, "=SUM(B1:B1)")));
    ASSERT_TRUE(static_cast<bool>(wb.delete_cols(0, /*col=*/1, /*count=*/1)));
    ASSERT_NE(wb.sheet(0).cell_at(0, 9), nullptr);
    EXPECT_EQ(wb.sheet(0).cell_at(0, 9)->formula_text, "=SUM(#REF!)");
  }
  {
    Workbook wb = Workbook::create();
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 10, "=SUM(A:D)")));
    ASSERT_TRUE(static_cast<bool>(wb.delete_cols(0, /*col=*/1, /*count=*/1)));
    ASSERT_NE(wb.sheet(0).cell_at(0, 9), nullptr);
    EXPECT_EQ(wb.sheet(0).cell_at(0, 9)->formula_text, "=SUM(A:C)");
  }
}

TEST(WorkbookRowColEdits, DeleteColsCollapsesFullyDeletedMultiCellRange) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 10, "=SUM(B1:D1)")));
  ASSERT_TRUE(static_cast<bool>(wb.delete_cols(0, /*col=*/1, /*count=*/3)));
  ASSERT_NE(wb.sheet(0).cell_at(0, 7), nullptr);
  EXPECT_EQ(wb.sheet(0).cell_at(0, 7)->formula_text, "=SUM(#REF!)");
}

TEST(WorkbookRowColEdits, DeleteColsShrunkRangeRecalculatesAfterSurvivorChange) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 0, Value::number(10.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 1, Value::number(20.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 2, Value::number(30.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 3, Value::number(40.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 5, 6, "=SUM(A1:D1)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_EQ(wb.sheet(0).cell_at(5, 6)->cached_value.as_number(), 100.0);

  ASSERT_TRUE(static_cast<bool>(wb.delete_cols(0, /*col=*/0, /*count=*/1)));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_NE(wb.sheet(0).cell_at(5, 5), nullptr);
  EXPECT_EQ(wb.sheet(0).cell_at(5, 5)->formula_text, "=SUM(A1:C1)");
  EXPECT_EQ(wb.sheet(0).cell_at(5, 5)->cached_value.as_number(), 90.0);

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 1, Value::number(50.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_EQ(wb.sheet(0).cell_at(5, 5)->cached_value.as_number(), 110.0);
}

TEST(WorkbookRowColEdits, InsertRowsShiftsMergeRanges) {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.mutable_merges() = {MergeRange{0, 0, 0, 1}, MergeRange{5, 0, 7, 1}};

  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, /*row=*/3, /*count=*/2)));

  ASSERT_EQ(s.merges().size(), 2U);
  // First merge sits entirely below the insert; unchanged.
  EXPECT_EQ(s.merges()[0].first_row, 0U);
  EXPECT_EQ(s.merges()[0].last_row, 0U);
  // Second merge shifted forward by 2.
  EXPECT_EQ(s.merges()[1].first_row, 7U);
  EXPECT_EQ(s.merges()[1].last_row, 9U);
}

TEST(WorkbookRowColEdits, DeleteRowsClampsStraddlingMerge) {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.mutable_merges() = {MergeRange{2, 0, 8, 1}};

  // Delete rows 4..6.
  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, /*row=*/4, /*count=*/3)));

  ASSERT_EQ(s.merges().size(), 1U);
  EXPECT_EQ(s.merges()[0].first_row, 2U);
  // Last row was 8 (past the deletion); shifts up by 3 to 5.
  EXPECT_EQ(s.merges()[0].last_row, 5U);
}

TEST(WorkbookRowColEdits, InsertRowsRewritesDefinedName) {
  Workbook wb = Workbook::create();
  std::vector<DefinedName> names;
  DefinedName dn;
  dn.name = "Region";
  dn.formula = "Sheet1!$A$5:$A$10";
  dn.local_sheet_id = -1;
  names.push_back(dn);
  wb.set_defined_names(std::move(names));

  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, /*row=*/3, /*count=*/2)));

  EXPECT_EQ(wb.defined_names()[0].formula, "Sheet1!$A$7:$A$12");
}

TEST(WorkbookRowColEdits, InsertRowsReindexesFormulasThatReachRangesViaDefinedName) {
  Workbook wb = Workbook::create();
  std::vector<DefinedName> names;
  DefinedName dn;
  dn.name = "Region";
  dn.formula = "Sheet1!$A$2:$A$3";
  dn.local_sheet_id = -1;
  names.push_back(dn);
  wb.set_defined_names(std::move(names));

  // A2 = 10, A3 = 20 (0-based rows 1 and 2).
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 1, 0, Value::number(10.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 2, 0, Value::number(20.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 2, "=SUM(Region)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_NE(wb.sheet(0).cell_at(0, 2), nullptr);
  EXPECT_EQ(wb.sheet(0).cell_at(0, 2)->cached_value.as_number(), 30.0);

  // Insert a row above the named range: the definition moves to $A$3:$A$4 and
  // the two values move with it. `=SUM(Region)` is unchanged as text.
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, /*row=*/1, /*count=*/1)));
  EXPECT_EQ(wb.defined_names()[0].formula, "Sheet1!$A$3:$A$4");

  // Write to A4 (0-based row 3). That row is inside the *new* definition and
  // outside the old one, so the edit only marks `=SUM(Region)` dirty if the
  // dep graph was re-indexed against the rewritten definition.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 3, 0, Value::number(25.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_NE(wb.sheet(0).cell_at(0, 2), nullptr);
  EXPECT_EQ(wb.sheet(0).cell_at(0, 2)->cached_value.as_number(), 35.0)
      << "SUM over the defined name did not follow the shifted definition";
}

TEST(WorkbookRowColEdits, InsertRowsRecomputesShiftedFormulasAndAggregates) {
  Workbook wb = Workbook::create();
  // Items: qty (col 1), unit (col 2), subtotal=qty*unit (col 3), tax=sub*0.08 (col 4).
  // Row 0 reserved for headers (left empty).
  // Item rows 1..5.
  const double qty[] = {24, 30, 4, 12, 5};
  const double unit[] = {0.42, 2.5, 4.45, 2.0, 0.75};
  for (std::uint32_t i = 0; i < 5; ++i) {
    const std::uint32_t r = i + 1;
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, r, 1, Value::number(qty[i]))));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, r, 2, Value::number(unit[i]))));
    // Excel-text uses 1-based addressing; the formula at cell row=i+1 references row=i+1 (1-based).
    const std::string subtotal_formula = "=B" + std::to_string(r + 1) + "*C" + std::to_string(r + 1);
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, r, 3, subtotal_formula)));
    const std::string tax_formula = "=D" + std::to_string(r + 1) + "*0.08";
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, r, 4, tax_formula)));
  }
  // Totals at row 7 (1-based row 8). =SUM(D2:D6) covers the entire item range.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 7, 3, "=SUM(D2:D6)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 7, 4, "=SUM(E2:E6)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 8, 3, "=D8+E8")));

  // Initial recalc establishes the baseline.
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  {
    const Sheet& s = wb.sheet(0);
    EXPECT_EQ(s.cell_at(5, 3)->cached_value.as_number(), 5.0 * 0.75);  // eraser subtotal pre-shift
    EXPECT_EQ(s.cell_at(7, 3)->cached_value.as_number(), 24 * 0.42 + 30 * 2.5 + 4 * 4.45 + 12 * 2.0 + 5 * 0.75);
  }

  // Insert one blank row at row index 1, push every item down by 1.
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, /*row=*/1, /*count=*/1)));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  const Sheet& s = wb.sheet(0);
  // Items now at rows 2..6, totals at row 8, with-tax at row 9.
  for (std::uint32_t i = 0; i < 5; ++i) {
    const std::uint32_t r = i + 2;
    ASSERT_NE(s.cell_at(r, 3), nullptr) << "subtotal cell missing at row " << r;
    ASSERT_NE(s.cell_at(r, 4), nullptr) << "tax cell missing at row " << r;
    EXPECT_EQ(s.cell_at(r, 3)->cached_value.as_number(), qty[i] * unit[i])
        << "subtotal mismatch at post-shift row " << r;
    EXPECT_EQ(s.cell_at(r, 4)->cached_value.as_number(), qty[i] * unit[i] * 0.08)
        << "tax mismatch at post-shift row " << r;
  }
  ASSERT_NE(s.cell_at(8, 3), nullptr);
  EXPECT_EQ(s.cell_at(8, 3)->cached_value.as_number(), 24 * 0.42 + 30 * 2.5 + 4 * 4.45 + 12 * 2.0 + 5 * 0.75)
      << "post-shift SUM mismatch (should NOT be #REF!)";
  ASSERT_NE(s.cell_at(8, 4), nullptr);
  EXPECT_NEAR(s.cell_at(8, 4)->cached_value.as_number(), (24 * 0.42 + 30 * 2.5 + 4 * 4.45 + 12 * 2.0 + 5 * 0.75) * 0.08,
              1e-9);
  ASSERT_NE(s.cell_at(9, 3), nullptr);
  EXPECT_NEAR(s.cell_at(9, 3)->cached_value.as_number(), (24 * 0.42 + 30 * 2.5 + 4 * 4.45 + 12 * 2.0 + 5 * 0.75) * 1.08,
              1e-9);
}

TEST(WorkbookRowColEdits, InsertColsRecomputesShiftedFormulasAndAggregates) {
  Workbook wb = Workbook::create();
  // Layout (1-based): row 3 has qty/unit per item; row 4 has the
  // subtotal formulas; H4 = SUM(B4:F4).
  const double qty[] = {2, 3, 4, 5, 6};
  const double unit[] = {1.5, 2.5, 3.5, 4.5, 5.5};
  for (std::uint32_t i = 0; i < 5; ++i) {
    const std::uint32_t c = i + 1;  // cols 1..5 (B..F in 1-based)
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 2, c, Value::number(qty[i]))));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 3, c, Value::number(unit[i]))));
    const std::string col_letter(1, static_cast<char>('A' + c));
    const std::string sub_formula = "=" + col_letter + "3*" + col_letter + "4";
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 4, c, sub_formula)));
  }
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 4, 7, "=SUM(B5:F5)")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  double expected_sum = 0.0;
  for (std::uint32_t i = 0; i < 5; ++i) {
    expected_sum += qty[i] * unit[i];
  }
  EXPECT_EQ(wb.sheet(0).cell_at(4, 7)->cached_value.as_number(), expected_sum);

  ASSERT_TRUE(static_cast<bool>(wb.insert_cols(0, /*col=*/1, /*count=*/1)));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  const Sheet& s = wb.sheet(0);
  for (std::uint32_t i = 0; i < 5; ++i) {
    const std::uint32_t c = i + 2;  // cols 2..6 post-shift
    ASSERT_NE(s.cell_at(4, c), nullptr);
    EXPECT_EQ(s.cell_at(4, c)->cached_value.as_number(), qty[i] * unit[i])
        << "post-shift subtotal mismatch at col " << c;
  }
  ASSERT_NE(s.cell_at(4, 8), nullptr);
  EXPECT_EQ(s.cell_at(4, 8)->cached_value.as_number(), expected_sum) << "post-shift SUM stale or #REF!";
}

TEST(WorkbookRowColEdits, DeleteRowsRecomputesShiftedFormulasAndAggregates) {
  Workbook wb = Workbook::create();
  const double qty[] = {2, 3, 4, 5, 6};
  for (std::uint32_t i = 0; i < 5; ++i) {
    const std::uint32_t r = i + 1;
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, r, 1, Value::number(qty[i]))));
    const std::string subtotal_formula = "=B" + std::to_string(r + 1) + "*2";
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, r, 3, subtotal_formula)));
  }
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 7, 3, "=SUM(D2:D6)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  double expected_pre = 0.0;
  for (std::uint32_t i = 0; i < 5; ++i) {
    expected_pre += qty[i] * 2;
  }
  EXPECT_EQ(wb.sheet(0).cell_at(7, 3)->cached_value.as_number(), expected_pre);

  // Delete rows 2..3 (0-based), removing the second and third items.
  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, /*row=*/2, /*count=*/2)));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  const Sheet& s = wb.sheet(0);
  ASSERT_NE(s.cell_at(5, 3), nullptr);  // SUM moved from row 7 → 5
  // SUM range was D2:D6 covering rows 1..5; after deleting rows 2..3
  // the range refs collapse: D2 (row 1, untouched) + D5..D6 (rows 4..5
  // pre-delete, now rows 2..3) survive; rows 2..3 (qty[1], qty[2]) drop.
  const double expected_post = qty[0] * 2 + qty[3] * 2 + qty[4] * 2;
  EXPECT_EQ(s.cell_at(5, 3)->cached_value.as_number(), expected_post)
      << "post-delete SUM mismatch (range refs should collapse, not stale)";
}

TEST(WorkbookRowColEdits, RejectsZeroCount) {
  Workbook wb = Workbook::create();
  auto r = wb.insert_rows(0, /*row=*/0, /*count=*/0);
  ASSERT_FALSE(static_cast<bool>(r));
  EXPECT_EQ(r.error().code, FormulonErrorCode::kInvalidArgument);
}

TEST(WorkbookRowColEdits, RejectsOutOfRangeSheetIndex) {
  Workbook wb = Workbook::create();
  auto r = wb.delete_cols(99, /*col=*/0, /*count=*/1);
  ASSERT_FALSE(static_cast<bool>(r));
  EXPECT_EQ(r.error().code, FormulonErrorCode::kInvalidArgument);
}

TEST(WorkbookSheetOps, RowInsertKeepsWrittenParentheses) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 4U, 1U, "=(A1)+1")));
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0U, 0U, 1U)));
  const Cell* cell = wb.sheet(0).cell_at(5U, 1U);
  ASSERT_NE(cell, nullptr);
  EXPECT_EQ(cell->formula_text, "=(A2)+1");
}

}  // namespace
}  // namespace formulon
