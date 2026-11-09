//
// C ABI tests for Insert / Delete Cells: the band that moves, the references
// that follow it, the structures that split or unmerge, and the status of
// every refusal.

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"

namespace {

// The last 0-based row and column of the Excel grid.
constexpr uint32_t kLastRow = 1048575U;
constexpr uint32_t kLastCol = 16383U;

double Number(const fm_workbook_t* wb, uint32_t row, uint32_t col) {
  fm_value_t v{};
  EXPECT_EQ(fm_sheet_get_value(wb, 0, row, col, &v), 0);
  EXPECT_EQ(v.kind, FM_VALUE_KIND_NUMBER) << row << "," << col;
  return v.u.number;
}

int32_t Kind(const fm_workbook_t* wb, uint32_t row, uint32_t col) {
  fm_value_t v{};
  EXPECT_EQ(fm_sheet_get_value(wb, 0, row, col, &v), 0);
  return v.kind;
}

std::string Formula(const fm_workbook_t* wb, uint32_t row, uint32_t col) {
  int32_t present = 0;
  const char* text = nullptr;
  EXPECT_EQ(fm_sheet_get_formula(wb, 0, row, col, &present, &text), 0);
  return present != 0 && text != nullptr ? std::string(text) : std::string();
}

std::size_t MergeCount(const fm_workbook_t* wb) {
  std::size_t count = 0;
  EXPECT_EQ(fm_sheet_get_merge_count(wb, 0, &count), 0);
  return count;
}

bool SameRect(const fm_cell_rect_t& a, const fm_cell_rect_t& b) {
  return a.first_row == b.first_row && a.first_col == b.first_col && a.last_row == b.last_row &&
         a.last_col == b.last_col;
}

TEST(FormulonCApiInsertDeleteCells, InsertDownMovesBandAndReferences) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 1, 1, 5.0), 0);     // B2
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 1, 0, 7.0), 0);     // A2, outside the band
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 3, "=B2"), 0);  // D1
  const fm_cell_rect_t range{1U, 1U, 2U, 1U};                     // B2:B3

  ASSERT_EQ(fm_sheet_insert_cells(wb.handle, 0, &range, FM_CELL_SHIFT_DOWN), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  EXPECT_DOUBLE_EQ(Number(wb.handle, 3, 1), 5.0);  // B4
  EXPECT_EQ(Kind(wb.handle, 1, 1), FM_VALUE_KIND_BLANK);
  EXPECT_DOUBLE_EQ(Number(wb.handle, 1, 0), 7.0);
  const std::string formula = Formula(wb.handle, 0, 3);
  EXPECT_NE(formula.find("B4"), std::string::npos) << formula;
  EXPECT_DOUBLE_EQ(Number(wb.handle, 0, 3), 5.0);
}

TEST(FormulonCApiInsertDeleteCells, DeleteLeftDeletedRefIsRefError) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 1, 1, 5.0), 0);     // B2
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 1, 2, 6.0), 0);     // C2
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 9, 0, "=B2"), 0);  // A10
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 9, 1, "=C2"), 0);  // B10
  const fm_cell_rect_t range{1U, 1U, 1U, 1U};                     // B2

  ASSERT_EQ(fm_sheet_delete_cells(wb.handle, 0, &range, FM_CELL_SHIFT_LEFT), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  EXPECT_DOUBLE_EQ(Number(wb.handle, 1, 1), 6.0);  // C2 moved to B2
  const std::string deleted = Formula(wb.handle, 9, 0);
  EXPECT_NE(deleted.find("#REF!"), std::string::npos) << deleted;
  EXPECT_EQ(Kind(wb.handle, 9, 0), FM_VALUE_KIND_ERROR);
  const std::string moved = Formula(wb.handle, 9, 1);
  EXPECT_NE(moved.find("B2"), std::string::npos) << moved;
  EXPECT_DOUBLE_EQ(Number(wb.handle, 9, 1), 6.0);
}

TEST(FormulonCApiInsertDeleteCells, StraddlingRangeUnchanged) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  for (uint32_t row = 0; row < 5U; ++row) {
    ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, row, 0, 1.0), 0);
    ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, row, 1, 1.0), 0);
  }
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 4, "=SUM(A1:B5)"), 0);  // E1
  const fm_cell_rect_t range{2U, 1U, 2U, 1U};                             // B3

  ASSERT_EQ(fm_sheet_insert_cells(wb.handle, 0, &range, FM_CELL_SHIFT_DOWN), 0) << fm_get_last_error_message();
  const std::string formula = Formula(wb.handle, 0, 4);
  EXPECT_NE(formula.find("A1:B5"), std::string::npos) << formula;
}

TEST(FormulonCApiInsertDeleteCells, CfRangeSplits) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_cell_rect_t sqref{1U, 0U, 1U, 2U};  // A2:C2
  fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
  rule.type = 0;  // Expression
  rule.formula1 = "TRUE";
  rule.ranges = &sqref;
  rule.range_count = 1;
  const char* rule_id = nullptr;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &rule_id), 0) << fm_get_last_error_message();
  const fm_cell_rect_t range{0U, 1U, 0U, 1U};  // B1

  ASSERT_EQ(fm_sheet_insert_cells(wb.handle, 0, &range, FM_CELL_SHIFT_DOWN), 0) << fm_get_last_error_message();
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_cf_rule_count(wb.handle, 0, &count), 0);
  std::vector<fm_cell_rect_t> ranges;
  for (std::size_t i = 0; i < count; ++i) {
    fm_cf_rule_t out = FM_STRUCT_INIT(fm_cf_rule_t);
    ASSERT_EQ(fm_sheet_get_cf_rule_at(wb.handle, 0, i, &out), 0);
    ranges.insert(ranges.end(), out.ranges, out.ranges + out.range_count);
  }
  const fm_cell_rect_t moved{2U, 1U, 2U, 1U};  // B3
  bool has_moved = false;
  bool has_original = false;
  for (const fm_cell_rect_t& r : ranges) {
    has_moved = has_moved || SameRect(r, moved);
    has_original = has_original || SameRect(r, sqref);
  }
  EXPECT_TRUE(has_moved);
  EXPECT_FALSE(has_original);
  EXPECT_GE(ranges.size(), 2U);
}

TEST(FormulonCApiInsertDeleteCells, PartialMergeUnmerged) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_cell_rect_t merge{1U, 0U, 1U, 2U};  // A2:C2
  ASSERT_EQ(fm_sheet_add_merge(wb.handle, 0, &merge, nullptr), 0);
  const fm_cell_rect_t range{0U, 1U, 0U, 1U};  // B1

  ASSERT_EQ(fm_sheet_insert_cells(wb.handle, 0, &range, FM_CELL_SHIFT_DOWN), 0) << fm_get_last_error_message();
  EXPECT_EQ(MergeCount(wb.handle), 0U);
}

TEST(FormulonCApiInsertDeleteCells, PartialTableRefusedPreconditionFailed) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 0, "a"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 1, "b"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 2, "c"), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 1, 1, 5.0), 0);  // B2
  fm_table_spec_t spec = FM_STRUCT_INIT(fm_table_spec_t);
  spec.range = fm_cell_rect_t{0U, 0U, 2U, 2U};  // A1:C3
  spec.name = "Data";
  std::size_t index = 0;
  ASSERT_EQ(fm_sheet_add_table(wb.handle, 0, &spec, &index), 0) << fm_get_last_error_message();
  const fm_cell_rect_t range{1U, 1U, 1U, 1U};  // B2

  EXPECT_EQ(fm_sheet_insert_cells(wb.handle, 0, &range, FM_CELL_SHIFT_DOWN), FM_STATUS_PRECONDITION_FAILED);
  EXPECT_EQ(fm_sheet_delete_cells(wb.handle, 0, &range, FM_CELL_SHIFT_UP), FM_STATUS_PRECONDITION_FAILED);
  // The workbook is unchanged.
  EXPECT_DOUBLE_EQ(Number(wb.handle, 1, 1), 5.0);
  EXPECT_EQ(Kind(wb.handle, 2, 1), FM_VALUE_KIND_BLANK);
}

TEST(FormulonCApiInsertDeleteCells, WrongDirectionInvalidArgument) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 1, 1, 5.0), 0);  // B2
  const fm_cell_rect_t range{1U, 1U, 1U, 1U};                  // B2
  EXPECT_EQ(fm_sheet_insert_cells(wb.handle, 0, &range, FM_CELL_SHIFT_UP), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(fm_sheet_insert_cells(wb.handle, 0, &range, FM_CELL_SHIFT_LEFT), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(fm_sheet_delete_cells(wb.handle, 0, &range, FM_CELL_SHIFT_RIGHT), FM_STATUS_INVALID_ARGUMENT);
  for (const int32_t shift : {-1, 4}) {
    EXPECT_EQ(fm_sheet_insert_cells(wb.handle, 0, &range, shift), FM_STATUS_INVALID_ARGUMENT) << shift;
    EXPECT_EQ(fm_sheet_delete_cells(wb.handle, 0, &range, shift), FM_STATUS_INVALID_ARGUMENT) << shift;
  }
  const fm_cell_rect_t inverted{2U, 1U, 1U, 1U};
  EXPECT_EQ(fm_sheet_insert_cells(wb.handle, 0, &inverted, FM_CELL_SHIFT_DOWN), FM_STATUS_INVALID_ARGUMENT);
  const fm_cell_rect_t off_grid{1U, 1U, 1U, kLastCol + 1U};
  EXPECT_EQ(fm_sheet_insert_cells(wb.handle, 0, &off_grid, FM_CELL_SHIFT_RIGHT), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_DOUBLE_EQ(Number(wb.handle, 1, 1), 5.0);
}

TEST(FormulonCApiInsertDeleteCells, ZeroShiftOnDeleteInvalidArgument) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 1, 1, 5.0), 0);  // B2
  const fm_cell_rect_t range{0U, 1U, 0U, 1U};                  // B1
  EXPECT_EQ(fm_sheet_delete_cells(wb.handle, 0, &range, 0), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_DOUBLE_EQ(Number(wb.handle, 1, 1), 5.0);
}

TEST(FormulonCApiInsertDeleteCells, WholeColumnDownInvalidArgument) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_cell_rect_t column{0U, 1U, kLastRow, 1U};  // B:B
  EXPECT_EQ(fm_sheet_insert_cells(wb.handle, 0, &column, FM_CELL_SHIFT_DOWN), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(fm_sheet_delete_cells(wb.handle, 0, &column, FM_CELL_SHIFT_UP), FM_STATUS_INVALID_ARGUMENT);
  const fm_cell_rect_t row{1U, 0U, 1U, kLastCol};  // 2:2
  EXPECT_EQ(fm_sheet_insert_cells(wb.handle, 0, &row, FM_CELL_SHIFT_RIGHT), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(fm_sheet_delete_cells(wb.handle, 0, &row, FM_CELL_SHIFT_LEFT), FM_STATUS_INVALID_ARGUMENT);
}

TEST(FormulonCApiInsertDeleteCells, NullRangeIsNullPointer) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_cell_rect_t range{1U, 1U, 1U, 1U};
  EXPECT_EQ(fm_sheet_insert_cells(wb.handle, 0, nullptr, FM_CELL_SHIFT_DOWN), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_delete_cells(wb.handle, 0, nullptr, FM_CELL_SHIFT_UP), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_insert_cells(nullptr, 0, &range, FM_CELL_SHIFT_DOWN), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_delete_cells(nullptr, 0, &range, FM_CELL_SHIFT_UP), FM_STATUS_BINDING_NULL_POINTER);
}

TEST(FormulonCApiInsertDeleteCells, SheetIndexOutOfRange) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_cell_rect_t range{1U, 1U, 1U, 1U};
  EXPECT_EQ(fm_sheet_insert_cells(wb.handle, 9, &range, FM_CELL_SHIFT_DOWN), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_delete_cells(wb.handle, 9, &range, FM_CELL_SHIFT_UP), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
}

}  // namespace
