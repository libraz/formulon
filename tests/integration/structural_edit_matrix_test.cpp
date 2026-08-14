//
// Matrix coverage for structural edits: every sheet-attached structure has to
// follow the same coordinate mapping the edit defines.
//
// The individual structures each grew their own shift helper, so a new one is
// easy to add and easy to forget to wire in — and a forgotten one is silent:
// the workbook still loads, the numbers still look plausible, and the
// conditional format simply colours the wrong band. These tests fix one
// workbook that carries every structure at a known coordinate and assert the
// post-edit coordinate for all of them under each operation, so a structure
// that is not wired into an edit fails here rather than in a user's file.
//
// Row/column axis (insert / delete): cells, formula text, defined names,
// hyperlinks (single-cell and range), merges, comments, data-validation
// ranges, conditional-format sqref, spill regions, table refs, column and row
// layout, manual page breaks, pivot anchors, and the dependency graph
// (observed through a recalc rather than by inspecting the graph directly).
//
// Eight of those — hyperlinks, comments, merges, validation ranges,
// conditional-format sqref, row/column layout, manual breaks, and pivot
// anchors — are exactly what `Sheet::shift_sheet_metadata` enumerates, and
// each is asserted under all four operations. Adding a structure there
// without adding it here puts the header's coverage claim back out of date.
//
// Sheet axis (rename / move / remove): the structures keyed by sheet name or
// sheet index rather than by coordinate.

#include <cstdint>
#include <memory>
#include <string>
#include <vector>

#include "cf/cf_types.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/defined_names.h"
#include "io/tables_reader.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

// Anchor coordinates every structure is placed at, all 0-based. Row 5 /
// column 3 sit strictly after the edit origin below, so an edit that fails to
// shift a structure leaves it at a coordinate this file can name exactly.
constexpr std::uint32_t kAnchorRow = 5;
constexpr std::uint32_t kAnchorCol = 3;
// Edits happen at row/col 1, before every anchor, and move one band.
constexpr std::uint32_t kEditOrigin = 1;
constexpr std::uint32_t kEditCount = 1;

/// Builds a workbook whose single sheet carries one instance of every
/// structure a row/column edit has to move.
Workbook MakeWorkbook() {
  Workbook wb = Workbook::create();

  // Values feeding the formula and aggregate targets. D6 / D7 (0-based
  // (5,3) / (6,3)).
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_value(0, kAnchorRow, kAnchorCol, Value::number(10.0))));
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_value(0, kAnchorRow + 1U, kAnchorCol, Value::number(20.0))));
  // A formula referencing them, parked out of the way at (0, 0).
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 0, "=SUM(D6:D7)")));

  std::vector<io::DefinedName> names;
  io::DefinedName dn;
  dn.name = "Anchor";
  dn.formula = "Sheet1!$D$6:$D$7";
  dn.local_sheet_id = -1;
  names.push_back(dn);
  wb.set_defined_names(std::move(names));

  Sheet& sheet = wb.sheet(0);

  // Hyperlinks are rectangle metadata, not just a left-top anchor. Include
  // both a single-cell link and a two-by-two range in the same edit matrix.
  Hyperlink single;
  single.row = kAnchorRow;
  single.col = kAnchorCol;
  single.last_row = kAnchorRow;
  single.last_col = kAnchorCol;
  single.target = "https://single.example";
  sheet.mutable_hyperlinks().push_back(single);
  Hyperlink hyperlink_range;
  hyperlink_range.row = kAnchorRow;
  hyperlink_range.col = kAnchorCol;
  hyperlink_range.last_row = kAnchorRow + 1U;
  hyperlink_range.last_col = kAnchorCol + 1U;
  hyperlink_range.location = "Sheet1!A1";
  sheet.mutable_hyperlinks().push_back(hyperlink_range);

  // Merges are rectangle metadata like hyperlinks; a two-by-two block
  // exercises both corners under one edit.
  sheet.mutable_merges().push_back(MergeRange{kAnchorRow, kAnchorCol, kAnchorRow + 1U, kAnchorCol + 1U});

  // Comments are single-cell anchors and shift on both axes.
  CellComment comment;
  comment.row = kAnchorRow;
  comment.col = kAnchorCol;
  comment.author = "Reviewer";
  comment.text = "anchored note";
  sheet.mutable_comments().push_back(comment);

  // A validation carries its own range list, shifted by the same
  // rectangle helper the merges use.
  DataValidation validation;
  validation.ranges.push_back(MergeRange{kAnchorRow, kAnchorCol, kAnchorRow + 1U, kAnchorCol + 1U});
  validation.type = 3;  // list
  validation.formula1 = "$D$6:$D$7";
  sheet.mutable_validations().push_back(validation);

  cf::ConditionalFormat format;
  cf::CFCellRange range;
  range.first = CellAddress{kAnchorRow, kAnchorCol};
  range.last = CellAddress{kAnchorRow + 1U, kAnchorCol};
  format.sqref.push_back(range);
  cf::CFRule rule;
  rule.type = cf::RuleType::Expression;
  rule.priority = 1;
  rule.formula1 = "D6>5";
  format.rules.push_back(rule);
  sheet.mutable_conditional_formats().push_back(format);

  ColumnLayout col;
  col.first = kAnchorCol;
  col.last = kAnchorCol;
  col.width = 33.5;
  sheet.mutable_layout().columns.push_back(col);

  RowLayout row;
  row.row = kAnchorRow;
  row.height = 42.0;
  row.has_height = true;
  sheet.mutable_layout().row_overrides.push_back(row);

  ManualBreak row_break;
  row_break.id = kAnchorRow;
  row_break.manual = true;
  sheet.mutable_print_settings().manual_row_breaks.push_back(row_break);

  ManualBreak col_break;
  col_break.id = kAnchorCol;
  col_break.manual = true;
  sheet.mutable_print_settings().manual_col_breaks.push_back(col_break);

  auto pivot = std::make_unique<pivot::PivotTable>();
  pivot->set_anchor(kAnchorRow, kAnchorCol, /*rows=*/2, /*cols=*/2);
  sheet.mutable_pivot_tables().push_back(std::move(pivot));

  io::TableMetadata table;
  table.id = 1;
  table.name = "Table1";
  table.display_name = "Table1";
  table.ref = "D6:D7";
  table.sheet_index = 0;
  wb.mutable_tables().push_back(table);

  return wb;
}

/// The single conditional-format range on the sheet.
const cf::CFCellRange& OnlyCfRange(const Workbook& wb) {
  return wb.sheet(0).conditional_formats().at(0).sqref.at(0);
}

/// Post-edit coordinates of the three rectangle/anchor structures after a
/// *row* edit: `moved_row` is where `kAnchorRow` ended up, and every column
/// coordinate must be untouched. Kept as one helper so the four operations
/// cannot drift into checking different structures.
void ExpectRowEditMovedRectangles(const Sheet& sheet, std::uint32_t moved_row) {
  ASSERT_EQ(sheet.merges().size(), 1U);
  EXPECT_EQ(sheet.merges().at(0).first_row, moved_row);
  EXPECT_EQ(sheet.merges().at(0).last_row, moved_row + 1U);
  EXPECT_EQ(sheet.merges().at(0).first_col, kAnchorCol);
  EXPECT_EQ(sheet.merges().at(0).last_col, kAnchorCol + 1U);

  ASSERT_EQ(sheet.comments().size(), 1U);
  EXPECT_EQ(sheet.comments().at(0).row, moved_row);
  EXPECT_EQ(sheet.comments().at(0).col, kAnchorCol);

  ASSERT_EQ(sheet.validations().size(), 1U);
  ASSERT_EQ(sheet.validations().at(0).ranges.size(), 1U);
  EXPECT_EQ(sheet.validations().at(0).ranges.at(0).first_row, moved_row);
  EXPECT_EQ(sheet.validations().at(0).ranges.at(0).last_row, moved_row + 1U);
  EXPECT_EQ(sheet.validations().at(0).ranges.at(0).first_col, kAnchorCol);
  EXPECT_EQ(sheet.validations().at(0).ranges.at(0).last_col, kAnchorCol + 1U);
}

/// Column-axis counterpart: `moved_col` is where `kAnchorCol` ended up and
/// every row coordinate must be untouched.
void ExpectColEditMovedRectangles(const Sheet& sheet, std::uint32_t moved_col) {
  ASSERT_EQ(sheet.merges().size(), 1U);
  EXPECT_EQ(sheet.merges().at(0).first_col, moved_col);
  EXPECT_EQ(sheet.merges().at(0).last_col, moved_col + 1U);
  EXPECT_EQ(sheet.merges().at(0).first_row, kAnchorRow);
  EXPECT_EQ(sheet.merges().at(0).last_row, kAnchorRow + 1U);

  ASSERT_EQ(sheet.comments().size(), 1U);
  EXPECT_EQ(sheet.comments().at(0).col, moved_col);
  EXPECT_EQ(sheet.comments().at(0).row, kAnchorRow);

  ASSERT_EQ(sheet.validations().size(), 1U);
  ASSERT_EQ(sheet.validations().at(0).ranges.size(), 1U);
  EXPECT_EQ(sheet.validations().at(0).ranges.at(0).first_col, moved_col);
  EXPECT_EQ(sheet.validations().at(0).ranges.at(0).last_col, moved_col + 1U);
  EXPECT_EQ(sheet.validations().at(0).ranges.at(0).first_row, kAnchorRow);
  EXPECT_EQ(sheet.validations().at(0).ranges.at(0).last_row, kAnchorRow + 1U);
}

/// Adds a second sheet whose metadata formulas all point at the edited sheet.
/// The edited sheet's own validation comes from `MakeWorkbook`.
///
/// A row/column edit on Sheet1 is not a Sheet1-only event: every qualified
/// reference to Sheet1 shifts wherever it lives, while an unqualified one on
/// Sheet2 names Sheet2 and must be left alone. Sheet2 therefore carries both
/// kinds, and its own sqref stays at the Sheet1 anchor coordinates so a
/// coordinate shift leaking across sheets shows up here too.
Workbook MakeCrossSheetWorkbook() {
  Workbook wb = MakeWorkbook();

  Sheet& other = wb.add_sheet("Sheet2");

  cf::ConditionalFormat cross_sheet_format;
  cf::CFCellRange cross_sheet_range;
  cross_sheet_range.first = CellAddress{kAnchorRow, kAnchorCol};
  cross_sheet_range.last = CellAddress{kAnchorRow + 1U, kAnchorCol};
  cross_sheet_format.sqref.push_back(cross_sheet_range);
  cf::CFRule cross_sheet_rule;
  cross_sheet_rule.type = cf::RuleType::Expression;
  cross_sheet_rule.priority = 1;
  cross_sheet_rule.formula1 = "Sheet1!D6>0";
  cross_sheet_format.rules.push_back(cross_sheet_rule);
  other.mutable_conditional_formats().push_back(cross_sheet_format);

  cf::ConditionalFormat local_format;
  cf::CFCellRange local_range;
  local_range.first = CellAddress{0, 0};
  local_range.last = CellAddress{0, 0};
  local_format.sqref.push_back(local_range);
  cf::CFRule local_rule;
  local_rule.type = cf::RuleType::Expression;
  local_rule.priority = 2;
  local_rule.formula1 = "D6>5";
  local_format.rules.push_back(local_rule);
  other.mutable_conditional_formats().push_back(local_format);

  Hyperlink qualified;
  qualified.row = 0;
  qualified.col = 0;
  qualified.last_row = 0;
  qualified.last_col = 0;
  qualified.location = "Sheet1!D6";
  other.mutable_hyperlinks().push_back(qualified);
  Hyperlink fragment;
  fragment.row = 1;
  fragment.col = 0;
  fragment.last_row = 1;
  fragment.last_col = 0;
  fragment.location = "#Sheet1!D6";
  other.mutable_hyperlinks().push_back(fragment);

  DataValidation cross_sheet_validation;
  cross_sheet_validation.ranges.push_back(MergeRange{0, 0, 0, 0});
  cross_sheet_validation.formula1 = "Sheet1!$D$6:$D$7";
  other.mutable_validations().push_back(cross_sheet_validation);

  io::TableMetadata table;
  table.id = 2;
  table.name = "Table2";
  table.display_name = "Table2";
  table.ref = "A1:B2";
  table.sheet_index = 1;
  table.columns.push_back(io::TableColumn{1, "Value", {}, {}, "Sheet1!D6*2"});
  wb.mutable_tables().push_back(table);

  auto cache = std::make_unique<pivot::PivotCache>();
  cache->set_cache_id(1U);
  cache->mutable_worksheet_source() = {true, "$D$6:$D$7", "Sheet1", ""};
  wb.add_pivot_cache(std::move(cache));

  return wb;
}

/// The cross-sheet rule on Sheet2 — the one qualified with the edited sheet.
const std::string& CrossSheetCfFormula(const Workbook& wb) {
  return *wb.sheet(1).conditional_formats().at(0).rules.at(0).formula1;
}

// ---------------------------------------------------------------------------
// Row axis
// ---------------------------------------------------------------------------

TEST(StructuralEditMatrix, InsertRowsMovesEveryRowAnchoredStructure) {
  Workbook wb = MakeWorkbook();
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, kEditOrigin, kEditCount)));
  const Sheet& sheet = wb.sheet(0);
  const std::uint32_t moved = kAnchorRow + kEditCount;

  // 1. Cells.
  ASSERT_NE(sheet.cell_at(moved, kAnchorCol), nullptr);
  EXPECT_EQ(sheet.cell_at(moved, kAnchorCol)->cached_value.as_number(), 10.0);
  EXPECT_EQ(sheet.cell_at(kAnchorRow, kAnchorCol), nullptr) << "value left behind at the pre-edit row";

  // 2. Formula text.
  ASSERT_NE(sheet.cell_at(0, 0), nullptr);
  EXPECT_EQ(sheet.cell_at(0, 0)->formula_text, "=SUM(D7:D8)");

  // 3. Defined names.
  EXPECT_EQ(wb.defined_names().at(0).formula, "Sheet1!$D$7:$D$8");

  // 4. Conditional-format sqref.
  EXPECT_EQ(OnlyCfRange(wb).first.row, moved);
  EXPECT_EQ(OnlyCfRange(wb).last.row, moved + 1U);

  ASSERT_EQ(sheet.hyperlinks().size(), 2U);
  EXPECT_EQ(sheet.hyperlinks()[0].row, moved);
  EXPECT_EQ(sheet.hyperlinks()[0].col, kAnchorCol);
  EXPECT_EQ(sheet.hyperlinks()[0].last_row, moved);
  EXPECT_EQ(sheet.hyperlinks()[0].last_col, kAnchorCol);
  EXPECT_EQ(sheet.hyperlinks()[1].row, moved);
  EXPECT_EQ(sheet.hyperlinks()[1].col, kAnchorCol);
  EXPECT_EQ(sheet.hyperlinks()[1].last_row, moved + 1U);
  EXPECT_EQ(sheet.hyperlinks()[1].last_col, kAnchorCol + 1U);

  ExpectRowEditMovedRectangles(sheet, moved);

  // 5. Row layout.
  ASSERT_EQ(sheet.layout().row_overrides.size(), 1U);
  EXPECT_EQ(sheet.layout().row_overrides.at(0).row, moved);

  // 6. Manual row breaks.
  ASSERT_EQ(sheet.print_settings().manual_row_breaks.size(), 1U);
  EXPECT_EQ(sheet.print_settings().manual_row_breaks.at(0).id, moved);

  // 7. Pivot anchor.
  ASSERT_EQ(sheet.pivot_tables().size(), 1U);
  EXPECT_EQ(sheet.pivot_tables().at(0)->anchor_row(), moved);

  // 8. Table ref.
  ASSERT_EQ(wb.tables().size(), 1U);
  EXPECT_EQ(wb.tables().at(0).ref, "D7:D8");

  // 9. Dependency graph: writing to a shifted cell must reach the formula.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, moved, kAnchorCol, Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_EQ(wb.sheet(0).cell_at(0, 0)->cached_value.as_number(), 21.0);
}

TEST(StructuralEditMatrix, DeleteRowsMovesEveryRowAnchoredStructure) {
  Workbook wb = MakeWorkbook();
  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, kEditOrigin, kEditCount)));
  const Sheet& sheet = wb.sheet(0);
  const std::uint32_t moved = kAnchorRow - kEditCount;

  ASSERT_NE(sheet.cell_at(moved, kAnchorCol), nullptr);
  EXPECT_EQ(sheet.cell_at(moved, kAnchorCol)->cached_value.as_number(), 10.0);

  ASSERT_NE(sheet.cell_at(0, 0), nullptr);
  EXPECT_EQ(sheet.cell_at(0, 0)->formula_text, "=SUM(D5:D6)");

  EXPECT_EQ(wb.defined_names().at(0).formula, "Sheet1!$D$5:$D$6");

  EXPECT_EQ(OnlyCfRange(wb).first.row, moved);
  EXPECT_EQ(OnlyCfRange(wb).last.row, moved + 1U);

  ASSERT_EQ(sheet.hyperlinks().size(), 2U);
  EXPECT_EQ(sheet.hyperlinks()[0].row, moved);
  EXPECT_EQ(sheet.hyperlinks()[0].col, kAnchorCol);
  EXPECT_EQ(sheet.hyperlinks()[0].last_row, moved);
  EXPECT_EQ(sheet.hyperlinks()[0].last_col, kAnchorCol);
  EXPECT_EQ(sheet.hyperlinks()[1].row, moved);
  EXPECT_EQ(sheet.hyperlinks()[1].col, kAnchorCol);
  EXPECT_EQ(sheet.hyperlinks()[1].last_row, moved + 1U);
  EXPECT_EQ(sheet.hyperlinks()[1].last_col, kAnchorCol + 1U);

  ExpectRowEditMovedRectangles(sheet, moved);

  ASSERT_EQ(sheet.layout().row_overrides.size(), 1U);
  EXPECT_EQ(sheet.layout().row_overrides.at(0).row, moved);

  ASSERT_EQ(sheet.print_settings().manual_row_breaks.size(), 1U);
  EXPECT_EQ(sheet.print_settings().manual_row_breaks.at(0).id, moved);

  ASSERT_EQ(sheet.pivot_tables().size(), 1U);
  EXPECT_EQ(sheet.pivot_tables().at(0)->anchor_row(), moved);

  EXPECT_EQ(wb.tables().at(0).ref, "D5:D6");

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, moved, kAnchorCol, Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_EQ(wb.sheet(0).cell_at(0, 0)->cached_value.as_number(), 21.0);
}

// ---------------------------------------------------------------------------
// Column axis
// ---------------------------------------------------------------------------

TEST(StructuralEditMatrix, InsertColsMovesEveryColumnAnchoredStructure) {
  Workbook wb = MakeWorkbook();
  ASSERT_TRUE(static_cast<bool>(wb.insert_cols(0, kEditOrigin, kEditCount)));
  const Sheet& sheet = wb.sheet(0);
  const std::uint32_t moved = kAnchorCol + kEditCount;

  ASSERT_NE(sheet.cell_at(kAnchorRow, moved), nullptr);
  EXPECT_EQ(sheet.cell_at(kAnchorRow, moved)->cached_value.as_number(), 10.0);
  // A column insert shifts within each row's dense cell vector, so the vacated
  // slot stays addressable and must read back blank rather than keeping the
  // old value (a row insert erases the whole row, hence the null check there).
  const Cell* vacated = sheet.cell_at(kAnchorRow, kAnchorCol);
  EXPECT_TRUE(vacated == nullptr || vacated->cached_value.is_blank()) << "value left behind at the pre-edit column";

  // The formula itself sits at (0, 0), which the insert at column 1 does not
  // move; only its references shift from column D to column E.
  ASSERT_NE(sheet.cell_at(0, 0), nullptr);
  EXPECT_EQ(sheet.cell_at(0, 0)->formula_text, "=SUM(E6:E7)");

  EXPECT_EQ(wb.defined_names().at(0).formula, "Sheet1!$E$6:$E$7");

  EXPECT_EQ(OnlyCfRange(wb).first.col, moved);
  EXPECT_EQ(OnlyCfRange(wb).last.col, moved);

  ASSERT_EQ(sheet.hyperlinks().size(), 2U);
  EXPECT_EQ(sheet.hyperlinks()[0].row, kAnchorRow);
  EXPECT_EQ(sheet.hyperlinks()[0].col, moved);
  EXPECT_EQ(sheet.hyperlinks()[0].last_row, kAnchorRow);
  EXPECT_EQ(sheet.hyperlinks()[0].last_col, moved);
  EXPECT_EQ(sheet.hyperlinks()[1].row, kAnchorRow);
  EXPECT_EQ(sheet.hyperlinks()[1].col, moved);
  EXPECT_EQ(sheet.hyperlinks()[1].last_row, kAnchorRow + 1U);
  EXPECT_EQ(sheet.hyperlinks()[1].last_col, moved + 1U);

  ExpectColEditMovedRectangles(sheet, moved);

  ASSERT_EQ(sheet.layout().columns.size(), 1U);
  EXPECT_EQ(sheet.layout().columns.at(0).first, moved);
  EXPECT_EQ(sheet.layout().columns.at(0).last, moved);

  ASSERT_EQ(sheet.print_settings().manual_col_breaks.size(), 1U);
  EXPECT_EQ(sheet.print_settings().manual_col_breaks.at(0).id, moved);

  ASSERT_EQ(sheet.pivot_tables().size(), 1U);
  EXPECT_EQ(sheet.pivot_tables().at(0)->anchor_col(), moved);

  EXPECT_EQ(wb.tables().at(0).ref, "E6:E7");

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, kAnchorRow, moved, Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_EQ(wb.sheet(0).cell_at(0, 0)->cached_value.as_number(), 21.0);
}

TEST(StructuralEditMatrix, DeleteColsMovesEveryColumnAnchoredStructure) {
  Workbook wb = MakeWorkbook();
  ASSERT_TRUE(static_cast<bool>(wb.delete_cols(0, kEditOrigin, kEditCount)));
  const Sheet& sheet = wb.sheet(0);
  const std::uint32_t moved = kAnchorCol - kEditCount;

  ASSERT_NE(sheet.cell_at(kAnchorRow, moved), nullptr);
  EXPECT_EQ(sheet.cell_at(kAnchorRow, moved)->cached_value.as_number(), 10.0);

  ASSERT_NE(sheet.cell_at(0, 0), nullptr);
  EXPECT_EQ(sheet.cell_at(0, 0)->formula_text, "=SUM(C6:C7)");

  EXPECT_EQ(wb.defined_names().at(0).formula, "Sheet1!$C$6:$C$7");

  EXPECT_EQ(OnlyCfRange(wb).first.col, moved);
  EXPECT_EQ(OnlyCfRange(wb).last.col, moved);

  ASSERT_EQ(sheet.hyperlinks().size(), 2U);
  EXPECT_EQ(sheet.hyperlinks()[0].row, kAnchorRow);
  EXPECT_EQ(sheet.hyperlinks()[0].col, moved);
  EXPECT_EQ(sheet.hyperlinks()[0].last_row, kAnchorRow);
  EXPECT_EQ(sheet.hyperlinks()[0].last_col, moved);
  EXPECT_EQ(sheet.hyperlinks()[1].row, kAnchorRow);
  EXPECT_EQ(sheet.hyperlinks()[1].col, moved);
  EXPECT_EQ(sheet.hyperlinks()[1].last_row, kAnchorRow + 1U);
  EXPECT_EQ(sheet.hyperlinks()[1].last_col, moved + 1U);

  ExpectColEditMovedRectangles(sheet, moved);

  ASSERT_EQ(sheet.layout().columns.size(), 1U);
  EXPECT_EQ(sheet.layout().columns.at(0).first, moved);

  ASSERT_EQ(sheet.print_settings().manual_col_breaks.size(), 1U);
  EXPECT_EQ(sheet.print_settings().manual_col_breaks.at(0).id, moved);

  ASSERT_EQ(sheet.pivot_tables().size(), 1U);
  EXPECT_EQ(sheet.pivot_tables().at(0)->anchor_col(), moved);

  EXPECT_EQ(wb.tables().at(0).ref, "C6:C7");

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, kAnchorRow, moved, Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_EQ(wb.sheet(0).cell_at(0, 0)->cached_value.as_number(), 21.0);
}

TEST(StructuralEditMatrix, DeleteRowsShrinksHyperlinkRange) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Hyperlink single;
  single.row = 5;
  single.col = 3;
  single.last_row = 5;
  single.last_col = 3;
  Hyperlink range;
  range.row = 5;
  range.col = 3;
  range.last_row = 6;
  range.last_col = 4;
  sheet.mutable_hyperlinks() = {single, range};

  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, 6, 1)));
  ASSERT_EQ(sheet.hyperlinks().size(), 2U);
  EXPECT_EQ(sheet.hyperlinks()[0].row, 5U);
  EXPECT_EQ(sheet.hyperlinks()[0].col, 3U);
  EXPECT_EQ(sheet.hyperlinks()[0].last_row, 5U);
  EXPECT_EQ(sheet.hyperlinks()[0].last_col, 3U);
  EXPECT_EQ(sheet.hyperlinks()[1].row, 5U);
  EXPECT_EQ(sheet.hyperlinks()[1].col, 3U);
  EXPECT_EQ(sheet.hyperlinks()[1].last_row, 5U);
  EXPECT_EQ(sheet.hyperlinks()[1].last_col, 4U);
}

TEST(StructuralEditMatrix, DeleteRowsDropsFullyConsumedHyperlinkRange) {
  Workbook wb = Workbook::create();
  Hyperlink range;
  range.row = 5;
  range.col = 3;
  range.last_row = 6;
  range.last_col = 4;
  wb.sheet(0).mutable_hyperlinks().push_back(range);

  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, 5, 2)));
  EXPECT_TRUE(wb.sheet(0).hyperlinks().empty());
}

TEST(StructuralEditMatrix, DeleteColsShrinksHyperlinkRange) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Hyperlink single;
  single.row = 5;
  single.col = 3;
  single.last_row = 5;
  single.last_col = 3;
  Hyperlink range;
  range.row = 5;
  range.col = 3;
  range.last_row = 6;
  range.last_col = 4;
  sheet.mutable_hyperlinks() = {single, range};

  ASSERT_TRUE(static_cast<bool>(wb.delete_cols(0, 4, 1)));
  ASSERT_EQ(sheet.hyperlinks().size(), 2U);
  EXPECT_EQ(sheet.hyperlinks()[0].col, 3U);
  EXPECT_EQ(sheet.hyperlinks()[0].row, 5U);
  EXPECT_EQ(sheet.hyperlinks()[0].last_col, 3U);
  EXPECT_EQ(sheet.hyperlinks()[0].last_row, 5U);
  EXPECT_EQ(sheet.hyperlinks()[1].col, 3U);
  EXPECT_EQ(sheet.hyperlinks()[1].row, 5U);
  EXPECT_EQ(sheet.hyperlinks()[1].last_col, 3U);
  EXPECT_EQ(sheet.hyperlinks()[1].last_row, 6U);
}

TEST(StructuralEditMatrix, DeleteColsDropsFullyConsumedHyperlinkRange) {
  Workbook wb = Workbook::create();
  Hyperlink range;
  range.row = 5;
  range.col = 3;
  range.last_row = 6;
  range.last_col = 4;
  wb.sheet(0).mutable_hyperlinks().push_back(range);

  ASSERT_TRUE(static_cast<bool>(wb.delete_cols(0, 3, 2)));
  EXPECT_TRUE(wb.sheet(0).hyperlinks().empty());
}

// ---------------------------------------------------------------------------
// Cross-sheet metadata formulas
// ---------------------------------------------------------------------------
//
// The coordinate half of a structure is sheet-local, but its formula half is
// not: a rule, link, validation, table column or pivot source on any sheet may
// name the edited sheet, and those references shift exactly like a cell
// formula's do. Text and coordinates are separate concerns here — the same
// edit that shifts Sheet2's reference to Sheet1 must leave Sheet2's own sqref
// and its own unqualified references alone.

TEST(StructuralEditMatrix, CrossSheetConditionalFormatFormulaShiftsOnEveryEdit) {
  Workbook inserted_rows = MakeCrossSheetWorkbook();
  ASSERT_TRUE(static_cast<bool>(inserted_rows.insert_rows(0, kEditOrigin, kEditCount)));
  EXPECT_EQ(CrossSheetCfFormula(inserted_rows), "Sheet1!D7>0");

  Workbook deleted_rows = MakeCrossSheetWorkbook();
  ASSERT_TRUE(static_cast<bool>(deleted_rows.delete_rows(0, kEditOrigin, kEditCount)));
  EXPECT_EQ(CrossSheetCfFormula(deleted_rows), "Sheet1!D5>0");

  Workbook inserted_cols = MakeCrossSheetWorkbook();
  ASSERT_TRUE(static_cast<bool>(inserted_cols.insert_cols(0, kEditOrigin, kEditCount)));
  EXPECT_EQ(CrossSheetCfFormula(inserted_cols), "Sheet1!E6>0");

  Workbook deleted_cols = MakeCrossSheetWorkbook();
  ASSERT_TRUE(static_cast<bool>(deleted_cols.delete_cols(0, kEditOrigin, kEditCount)));
  EXPECT_EQ(CrossSheetCfFormula(deleted_cols), "Sheet1!C6>0");
}

TEST(StructuralEditMatrix, CrossSheetConditionalFormatFormulaBecomesRefWhenTargetDeleted) {
  Workbook wb = MakeCrossSheetWorkbook();
  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, kAnchorRow, kEditCount)));
  EXPECT_EQ(CrossSheetCfFormula(wb), "#REF!>0");
}

TEST(StructuralEditMatrix, LocalConditionalFormatOnOtherSheetIsNotShifted) {
  Workbook wb = MakeCrossSheetWorkbook();
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, kEditOrigin, kEditCount)));
  ASSERT_TRUE(wb.sheet(1).conditional_formats().at(1).rules.at(0).formula1.has_value());
  EXPECT_EQ(*wb.sheet(1).conditional_formats().at(1).rules.at(0).formula1, "D6>5")
      << "an unqualified reference on Sheet2 names Sheet2, which the edit did not touch";
}

TEST(StructuralEditMatrix, TargetSheetConditionalFormatFormulaStillShifts) {
  Workbook wb = MakeCrossSheetWorkbook();
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, 0, 2)));
  ASSERT_TRUE(wb.sheet(0).conditional_formats().at(0).rules.at(0).formula1.has_value());
  EXPECT_EQ(*wb.sheet(0).conditional_formats().at(0).rules.at(0).formula1, "D8>5");
}

TEST(StructuralEditMatrix, CrossSheetConditionalFormatSqrefIsNotShifted) {
  Workbook wb = MakeCrossSheetWorkbook();
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, kEditOrigin, kEditCount)));
  const cf::CFCellRange& range = wb.sheet(1).conditional_formats().at(0).sqref.at(0);
  EXPECT_EQ(range.first.row, kAnchorRow);
  EXPECT_EQ(range.last.row, kAnchorRow + 1U);
  EXPECT_EQ(range.first.col, kAnchorCol);
}

TEST(StructuralEditMatrix, CrossSheetHyperlinkLocationShiftsOnRowInsert) {
  Workbook wb = MakeCrossSheetWorkbook();
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, kEditOrigin, kEditCount)));
  ASSERT_EQ(wb.sheet(1).hyperlinks().size(), 2U);
  EXPECT_EQ(wb.sheet(1).hyperlinks()[0].location, "Sheet1!D7");
  EXPECT_EQ(wb.sheet(1).hyperlinks()[1].location, "#Sheet1!D7");
  EXPECT_EQ(wb.sheet(1).hyperlinks()[0].row, 0U) << "the link anchor lives on Sheet2 and does not move";
}

TEST(StructuralEditMatrix, TargetSheetValidationListFormulaShiftsOnRowInsert) {
  Workbook wb = MakeCrossSheetWorkbook();
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, kEditOrigin, kEditCount)));
  ASSERT_EQ(wb.sheet(0).validations().size(), 1U);
  EXPECT_EQ(wb.sheet(0).validations()[0].formula1, "$D$7:$D$8");
}

TEST(StructuralEditMatrix, CrossSheetValidationFormulaShiftsOnRowInsert) {
  Workbook wb = MakeCrossSheetWorkbook();
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, kEditOrigin, kEditCount)));
  ASSERT_EQ(wb.sheet(1).validations().size(), 1U);
  EXPECT_EQ(wb.sheet(1).validations()[0].formula1, "Sheet1!$D$7:$D$8");
}

TEST(StructuralEditMatrix, CrossSheetTableCalculatedColumnFormulaShifts) {
  Workbook wb = MakeCrossSheetWorkbook();
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, kEditOrigin, kEditCount)));
  ASSERT_EQ(wb.tables().size(), 2U);
  ASSERT_EQ(wb.tables().at(1).columns.size(), 1U);
  EXPECT_EQ(wb.tables().at(1).columns.at(0).calculated_column_formula, "Sheet1!D7*2");
  EXPECT_EQ(wb.tables().at(1).ref, "A1:B2") << "the table's own range lives on Sheet2 and does not move";
}

TEST(StructuralEditMatrix, PivotWorksheetSourceRefShiftsOnRowInsert) {
  Workbook wb = MakeCrossSheetWorkbook();
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, kEditOrigin, kEditCount)));
  ASSERT_EQ(wb.pivot_caches().size(), 1U);
  ASSERT_NE(wb.pivot_caches()[0], nullptr);
  EXPECT_EQ(wb.pivot_caches()[0]->worksheet_source().ref, "$D$7:$D$8");
  EXPECT_EQ(wb.pivot_caches()[0]->worksheet_source().sheet, "Sheet1");
}

// ---------------------------------------------------------------------------
// Spill regions
// ---------------------------------------------------------------------------
//
// Spill is the one structure that is deliberately *not* remapped: the edit
// drops every region and the next recalc re-spills from the anchor formula.
// Remapping would have to decide what a spill that straddles a deleted band
// means, and the recalc already knows.

TEST(StructuralEditMatrix, RowEditDropsSpillAndRecalcRebuildsIt) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, kAnchorRow, kAnchorCol, "=SEQUENCE(3)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_NE(wb.sheet(0).spill_region_at_anchor(kAnchorRow, kAnchorCol), nullptr);

  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, kEditOrigin, kEditCount)));
  const std::uint32_t moved = kAnchorRow + kEditCount;
  EXPECT_EQ(wb.sheet(0).spill_region_at_anchor(kAnchorRow, kAnchorCol), nullptr)
      << "a stale region at the pre-edit anchor would hide real cells";

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  const SpillRegion* region = wb.sheet(0).spill_region_at_anchor(moved, kAnchorCol);
  ASSERT_NE(region, nullptr) << "recalc did not re-spill at the shifted anchor";
  EXPECT_EQ(region->rows, 3U);
  EXPECT_EQ(region->cols, 1U);
}

// ---------------------------------------------------------------------------
// Compact range dependencies
// ---------------------------------------------------------------------------
//
// A rectangle too wide to be flattened into per-cell edges is remapped by
// rewriting the formula text and re-registering it. The rewrite is only half
// the contract: the re-registered rectangle has to watch the shifted
// coordinates, which nothing else in this file would notice.

TEST(StructuralEditMatrix, StructuralEditRemapsCompactBoundedRange) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Sheet2");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, kAnchorRow, 0, Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(1, 0, 2, "=SUM(Sheet1!A1:A60000)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_NE(wb.sheet(1).cell_at(0, 2), nullptr);
  EXPECT_DOUBLE_EQ(wb.sheet(1).cell_at(0, 2)->cached_value.as_number(), 1.0);

  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, 0, 2)));
  EXPECT_EQ(wb.sheet(1).cell_at(0, 2)->formula_text, "=SUM(Sheet1!A3:A60002)");

  // A write at the rectangle's new first row must wake the aggregate; the
  // pre-edit rectangle would have covered it too, so the payload also has to
  // change to prove the shifted literal is still inside.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 2, 0, Value::number(5.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_DOUBLE_EQ(wb.sheet(1).cell_at(0, 2)->cached_value.as_number(), 6.0);
}

// ---------------------------------------------------------------------------
// Sheet axis
// ---------------------------------------------------------------------------

TEST(StructuralEditMatrix, RenameSheetRewritesEverySheetNameQualifiedReference) {
  Workbook wb = MakeWorkbook();
  ASSERT_TRUE(static_cast<bool>(wb.rename_sheet(0, "Renamed")));

  EXPECT_EQ(wb.sheet(0).name(), "Renamed");
  EXPECT_EQ(wb.defined_names().at(0).formula, "Renamed!$D$6:$D$7");
}

TEST(StructuralEditMatrix, MoveSheetKeepsTablesAttachedToTheirOwnSheet) {
  Workbook wb = MakeWorkbook();
  wb.add_sheet("Second");
  ASSERT_EQ(wb.sheet_count(), 2U);

  // Move Sheet1 (index 0, the one carrying the table) to the end.
  ASSERT_TRUE(static_cast<bool>(wb.move_sheet(0, 1)));
  ASSERT_EQ(wb.sheet(1).name(), "Sheet1");

  ASSERT_EQ(wb.tables().size(), 1U);
  EXPECT_EQ(wb.tables().at(0).sheet_index, 1U) << "the table stayed pinned to the old index, attaching it to Second";
}

TEST(StructuralEditMatrix, RemoveSheetDropsItsTables) {
  Workbook wb = MakeWorkbook();
  wb.add_sheet("Second");
  ASSERT_TRUE(static_cast<bool>(wb.remove_sheet(0)));

  EXPECT_TRUE(wb.tables().empty()) << "a table left behind would be written against a sheet that no longer exists";
}

}  // namespace
}  // namespace formulon
