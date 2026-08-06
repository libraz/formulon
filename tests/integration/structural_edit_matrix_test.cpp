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
// conditional-format sqref, spill regions, table refs, column and row layout,
// manual page breaks, pivot anchors, and the dependency graph (observed
// through a recalc rather than by inspecting the graph directly).
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
