//
// C ABI tests for What-If data tables: creating, reading and removing one,
// the write and structural-edit refusals, the `TABLE()` formula its result
// cells report, and the explicit table calculation.

#include <cstddef>
#include <cstdint>

#include "c_api/formulon_c.h"
#include "formulon_c_pivot_test_helpers.h"
#include "gtest/gtest.h"
#include "utils/error.h"

namespace {

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kBindingNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);
constexpr fm_status_t kPreconditionFailed = static_cast<fm_status_t>(formulon::FormulonErrorCode::kPreconditionFailed);
constexpr fm_status_t kDataTableInvalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kDataTableInvalid);

double Number(const fm_workbook_t* wb, uint32_t row, uint32_t col) {
  fm_value_t v{};
  EXPECT_EQ(fm_workbook_get_value(wb, 0, row, col, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_NUMBER) << row << "," << col;
  return v.u.number;
}

// B1 = 3 (the column input), B2 = B1 * B1 over A3:A5 = {1, 2, 3}: a squares table at B3:B5.
void BuildSquares(WorkbookGuard& wb) {
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 1, 3.0), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 1, 1, "=B1*B1"), 0);
  for (uint32_t i = 0; i < 3; ++i) {
    ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 2 + i, 0, static_cast<double>(i + 1)), 0);
  }
}

fm_data_table SquaresSpec() {
  fm_data_table spec{};
  spec.first_row = 2;
  spec.first_col = 1;
  spec.last_row = 4;
  spec.last_col = 1;
  spec.has_col_input = 1;
  spec.col_input_row = 0;
  spec.col_input_col = 1;
  return spec;
}

TEST(FormulonCApiDataTable, AddComputesReadsAndRefusesWrites) {
  WorkbookGuard wb;
  BuildSquares(wb);
  const fm_data_table spec = SquaresSpec();
  ASSERT_EQ(fm_sheet_add_data_table(wb.handle, 0, &spec), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  EXPECT_DOUBLE_EQ(Number(wb.handle, 2, 1), 1.0);
  EXPECT_DOUBLE_EQ(Number(wb.handle, 3, 1), 4.0);
  EXPECT_DOUBLE_EQ(Number(wb.handle, 4, 1), 9.0);
  EXPECT_DOUBLE_EQ(Number(wb.handle, 0, 1), 3.0);
  EXPECT_DOUBLE_EQ(Number(wb.handle, 1, 1), 9.0);

  int32_t present = 0;
  fm_data_table info{};
  ASSERT_EQ(fm_sheet_data_table_at(wb.handle, 0, 3, 1, &present, &info), 0);
  EXPECT_EQ(present, 1);
  EXPECT_EQ(info.first_row, 2U);
  EXPECT_EQ(info.last_row, 4U);
  EXPECT_EQ(info.first_col, 1U);
  EXPECT_EQ(info.last_col, 1U);
  EXPECT_EQ(info.has_col_input, 1);
  EXPECT_EQ(info.has_row_input, 0);
  EXPECT_EQ(info.col_input_row, 0U);
  EXPECT_EQ(info.col_input_col, 1U);
  EXPECT_EQ(info.two_dimensional, 0);
  EXPECT_EQ(info.row_oriented, 0);
  EXPECT_EQ(info.col_input_deleted, 0);

  // A cell outside the range reports no table and a zeroed record.
  ASSERT_EQ(fm_sheet_data_table_at(wb.handle, 0, 1, 1, &present, &info), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(info.last_row, 0U);

  EXPECT_EQ(fm_workbook_set_number(wb.handle, 0, 3, 1, 7.0), kPreconditionFailed);
  EXPECT_EQ(fm_workbook_set_text(wb.handle, 0, 2, 1, "x"), kPreconditionFailed);
  EXPECT_EQ(fm_workbook_set_formula(wb.handle, 0, 4, 1, "=1"), kPreconditionFailed);
  EXPECT_DOUBLE_EQ(Number(wb.handle, 3, 1), 4.0);
  // The header cells stay editable and the table follows them.
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 3, 0, 10.0), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  EXPECT_DOUBLE_EQ(Number(wb.handle, 3, 1), 100.0);
}

TEST(FormulonCApiDataTable, TwoVariableTableReadsBothInputs) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 5, 1.0), 0);  // F1: row input.
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 1, 5, 1.0), 0);  // F2: column input.
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 2, 2, "=F1*F2"), 0);
  for (uint32_t j = 0; j < 3; ++j) {
    ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 2, 3 + j, static_cast<double>(j + 1)), 0);
    ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 3 + j, 2, static_cast<double>(10 * (j + 1))), 0);
  }
  fm_data_table spec{};
  spec.first_row = 3;
  spec.first_col = 3;
  spec.last_row = 5;
  spec.last_col = 5;
  spec.has_row_input = 1;
  spec.row_input_row = 0;
  spec.row_input_col = 5;
  spec.has_col_input = 1;
  spec.col_input_row = 1;
  spec.col_input_col = 5;
  ASSERT_EQ(fm_sheet_add_data_table(wb.handle, 0, &spec), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  EXPECT_DOUBLE_EQ(Number(wb.handle, 3, 3), 10.0);
  EXPECT_DOUBLE_EQ(Number(wb.handle, 4, 4), 40.0);
  EXPECT_DOUBLE_EQ(Number(wb.handle, 5, 5), 90.0);

  int32_t present = 0;
  fm_data_table info{};
  ASSERT_EQ(fm_sheet_data_table_at(wb.handle, 0, 5, 5, &present, &info), 0);
  EXPECT_EQ(present, 1);
  EXPECT_EQ(info.two_dimensional, 1);
  EXPECT_EQ(info.has_row_input, 1);
  EXPECT_EQ(info.has_col_input, 1);
  EXPECT_EQ(info.row_input_row, 0U);
  EXPECT_EQ(info.col_input_row, 1U);

  // Deleting the row input's row flags it deleted and turns the table into #REF!.
  ASSERT_EQ(fm_workbook_delete_rows(wb.handle, 0, 0, 1), 0);
  ASSERT_EQ(fm_sheet_data_table_at(wb.handle, 0, 4, 3, &present, &info), 0);
  ASSERT_EQ(present, 1);
  EXPECT_EQ(info.row_input_deleted, 1);
  EXPECT_EQ(info.col_input_deleted, 0);
  EXPECT_EQ(info.first_row, 2U);
  EXPECT_EQ(info.col_input_row, 0U);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 2, 3, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_ERROR);
}

TEST(FormulonCApiDataTable, RemoveBlanksTheResultCells) {
  WorkbookGuard wb;
  BuildSquares(wb);
  const fm_data_table spec = SquaresSpec();
  ASSERT_EQ(fm_sheet_add_data_table(wb.handle, 0, &spec), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  ASSERT_EQ(fm_sheet_remove_data_table(wb.handle, 0, 4, 1), 0);
  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 3, 1, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_BLANK);
  int32_t present = 1;
  fm_data_table info{};
  ASSERT_EQ(fm_sheet_data_table_at(wb.handle, 0, 3, 1, &present, &info), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(fm_sheet_remove_data_table(wb.handle, 0, 4, 1), kDataTableInvalid);
  EXPECT_EQ(fm_workbook_set_number(wb.handle, 0, 3, 1, 7.0), 0);
}

TEST(FormulonCApiDataTable, AutoNoTableComputesOnlyOnTheExplicitCall) {
  WorkbookGuard wb;
  BuildSquares(wb);
  const fm_data_table spec = SquaresSpec();
  ASSERT_EQ(fm_sheet_add_data_table(wb.handle, 0, &spec), 0);
  ASSERT_EQ(fm_workbook_set_calc_mode(wb.handle, FM_CALC_MODE_AUTO_NO_TABLE), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  EXPECT_DOUBLE_EQ(Number(wb.handle, 4, 1), 0.0);
  EXPECT_DOUBLE_EQ(Number(wb.handle, 1, 1), 9.0);
  ASSERT_EQ(fm_workbook_recalc_data_tables(wb.handle), 0);
  EXPECT_DOUBLE_EQ(Number(wb.handle, 4, 1), 9.0);
  EXPECT_DOUBLE_EQ(Number(wb.handle, 0, 1), 3.0);
}

TEST(FormulonCApiDataTable, SurvivesSaveAndLoad) {
  WorkbookGuard wb;
  BuildSquares(wb);
  const fm_data_table spec = SquaresSpec();
  ASSERT_EQ(fm_sheet_add_data_table(wb.handle, 0, &spec), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);

  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &loaded.handle), 0);
  int32_t present = 0;
  fm_data_table info{};
  ASSERT_EQ(fm_sheet_data_table_at(loaded.handle, 0, 3, 1, &present, &info), 0);
  EXPECT_EQ(present, 1);
  EXPECT_EQ(info.has_col_input, 1);
  EXPECT_EQ(info.col_input_col, 1U);
  EXPECT_DOUBLE_EQ(Number(loaded.handle, 4, 1), 9.0);
  EXPECT_EQ(fm_workbook_set_number(loaded.handle, 0, 4, 1, 1.0), kPreconditionFailed);
}

TEST(FormulonCApiDataTable, ResultCellsReportTheTableFormula) {
  WorkbookGuard wb;
  BuildSquares(wb);
  const fm_data_table spec = SquaresSpec();
  ASSERT_EQ(fm_sheet_add_data_table(wb.handle, 0, &spec), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  const char* formula = nullptr;
  ASSERT_EQ(fm_workbook_get_formula(wb.handle, 0, 3, 1, &formula), 0);
  EXPECT_STREQ(formula, "=TABLE(,B1)");
  ASSERT_EQ(fm_workbook_get_formula_r1c1(wb.handle, 0, 3, 1, &formula), 0);
  EXPECT_STREQ(formula, "TABLE(,R[-3]C)");
  ASSERT_EQ(fm_workbook_get_formula(wb.handle, 0, 3, 0, &formula), 0);
  EXPECT_STREQ(formula, "");  // A header value.

  fm_cell_range_t* range = nullptr;
  ASSERT_EQ(fm_sheet_cells_in_range(wb.handle, 0, 2, 1, 4, 1, 0, 0, &range), 0);
  size_t count = 0;
  ASSERT_EQ(fm_cell_range_count(range, &count), 0);
  EXPECT_EQ(count, 3U);
  for (size_t i = 0; i < count; ++i) {
    uint32_t row = 0;
    uint32_t col = 0;
    fm_value_t v{};
    ASSERT_EQ(fm_cell_range_at(range, i, &row, &col, &formula, &v), 0);
    EXPECT_STREQ(formula, "=TABLE(,B1)") << row;
  }
  fm_cell_range_destroy(range);

  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 6, 0, "=FORMULATEXT(B4)"), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 6, 1, "=ISFORMULA(B4)"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 6, 0, &v), 0);
  ASSERT_EQ(v.kind, FM_VAL_TEXT);
  EXPECT_STREQ(v.u.text, "{=TABLE(,B1)}");
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 6, 1, &v), 0);
  ASSERT_EQ(v.kind, FM_VAL_BOOL);
  EXPECT_NE(v.u.boolean, 0);
}

TEST(FormulonCApiDataTable, StructuralEditsMoveOrRemoveTheTableWhole) {
  WorkbookGuard wb;
  BuildSquares(wb);
  const fm_data_table spec = SquaresSpec();
  ASSERT_EQ(fm_sheet_add_data_table(wb.handle, 0, &spec), 0);
  // The table's block is A2:B5: the results B3:B5, the formula row 2 and the value column A.
  EXPECT_EQ(fm_workbook_delete_rows(wb.handle, 0, 3, 1), kPreconditionFailed);
  EXPECT_EQ(fm_workbook_delete_rows(wb.handle, 0, 2, 3), kPreconditionFailed);
  EXPECT_EQ(fm_workbook_insert_rows(wb.handle, 0, 4, 1), kPreconditionFailed);
  EXPECT_EQ(fm_workbook_delete_cols(wb.handle, 0, 0, 1), kPreconditionFailed);
  EXPECT_EQ(fm_workbook_insert_cols(wb.handle, 0, 1, 1), kPreconditionFailed);
  int32_t present = 0;
  fm_data_table info{};
  ASSERT_EQ(fm_sheet_data_table_at(wb.handle, 0, 2, 1, &present, &info), 0);
  EXPECT_EQ(present, 1);
  EXPECT_EQ(info.last_row, 4U);
  // Rows inserted above the block move it; deleting the whole block removes it.
  ASSERT_EQ(fm_workbook_insert_rows(wb.handle, 0, 1, 1), 0);
  ASSERT_EQ(fm_sheet_data_table_at(wb.handle, 0, 5, 1, &present, &info), 0);
  EXPECT_EQ(present, 1);
  EXPECT_EQ(info.first_row, 3U);
  ASSERT_EQ(fm_workbook_delete_rows(wb.handle, 0, 2, 4), 0);
  ASSERT_EQ(fm_sheet_data_table_at(wb.handle, 0, 3, 1, &present, &info), 0);
  EXPECT_EQ(present, 0);
}

TEST(FormulonCApiDataTable, RejectsBadArguments) {
  WorkbookGuard wb;
  BuildSquares(wb);
  fm_data_table spec = SquaresSpec();
  int32_t present = 0;
  fm_data_table info{};
  EXPECT_EQ(fm_sheet_add_data_table(wb.handle, 0, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_add_data_table(wb.handle, 9, &spec), kInvalidArgument);
  EXPECT_EQ(fm_sheet_data_table_at(wb.handle, 0, 0, 0, nullptr, &info), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_data_table_at(wb.handle, 0, 0, 0, &present, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_data_table_at(wb.handle, 9, 0, 0, &present, &info), kInvalidArgument);
  EXPECT_EQ(fm_sheet_remove_data_table(wb.handle, 9, 0, 0), kInvalidArgument);
  EXPECT_EQ(fm_workbook_recalc_data_tables(nullptr), kBindingNullPointer);

  fm_data_table none = spec;
  none.has_col_input = 0;
  EXPECT_EQ(fm_sheet_add_data_table(wb.handle, 0, &none), kDataTableInvalid);
  fm_data_table no_header = spec;
  no_header.first_row = 0;
  EXPECT_EQ(fm_sheet_add_data_table(wb.handle, 0, &no_header), kDataTableInvalid);
  fm_data_table input_inside = spec;
  input_inside.col_input_row = 3;
  input_inside.col_input_col = 1;
  EXPECT_EQ(fm_sheet_add_data_table(wb.handle, 0, &input_inside), kDataTableInvalid);
  fm_data_table off_grid = spec;
  off_grid.last_row = 0xFFFFFFFFU;
  EXPECT_EQ(fm_sheet_add_data_table(wb.handle, 0, &off_grid), kInvalidArgument);
  ASSERT_EQ(fm_sheet_add_data_table(wb.handle, 0, &spec), 0);
  EXPECT_EQ(fm_sheet_add_data_table(wb.handle, 0, &spec), kDataTableInvalid);
}

}  // namespace
