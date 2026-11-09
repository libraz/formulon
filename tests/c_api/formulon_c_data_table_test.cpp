//
// C ABI tests for What-If data tables: creating, reading and removing one,
// the write and structural-edit refusals, the `TABLE()` formula its result
// cells report, and the explicit table calculation.

#include <cstddef>
#include <cstdint>

#include "c_api/formulon_c.h"
#include "formulon_c_pivot_test_helpers.h"
#include "formulon_c_record_arg.h"
#include "gtest/gtest.h"
#include "utils/error.h"

namespace {

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kBindingNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);
constexpr fm_status_t kPreconditionFailed = static_cast<fm_status_t>(formulon::FormulonErrorCode::kPreconditionFailed);
constexpr fm_status_t kDataTableInvalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kDataTableInvalid);

double Number(const fm_workbook_t* wb, uint32_t row, uint32_t col) {
  fm_value_t v{};
  EXPECT_EQ(fm_sheet_get_value(wb, 0, row, col, &v), 0);
  EXPECT_EQ(v.kind, FM_VALUE_KIND_NUMBER) << row << "," << col;
  return v.u.number;
}

// B1 = 3 (the column input), B2 = B1 * B1 over A3:A5 = {1, 2, 3}: a squares table at B3:B5.
void BuildSquares(WorkbookGuard& wb) {
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 1, 3.0), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 1, 1, "=B1*B1"), 0);
  for (uint32_t i = 0; i < 3; ++i) {
    ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 2 + i, 0, static_cast<double>(i + 1)), 0);
  }
}

fm_data_table_t SquaresSpec() {
  fm_data_table_t spec = FM_STRUCT_INIT(fm_data_table_t);
  spec.rect.first_row = 2;
  spec.rect.first_col = 1;
  spec.rect.last_row = 4;
  spec.rect.last_col = 1;
  spec.has_column_input = 1;
  spec.column_input_row = 0;
  spec.column_input_col = 1;
  return spec;
}

TEST(FormulonCApiDataTable, AddComputesReadsAndRefusesWrites) {
  WorkbookGuard wb;
  BuildSquares(wb);
  const fm_data_table_t spec = SquaresSpec();
  ASSERT_EQ(fm_sheet_add_data_table(wb.handle, 0, &spec, nullptr), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  EXPECT_DOUBLE_EQ(Number(wb.handle, 2, 1), 1.0);
  EXPECT_DOUBLE_EQ(Number(wb.handle, 3, 1), 4.0);
  EXPECT_DOUBLE_EQ(Number(wb.handle, 4, 1), 9.0);
  EXPECT_DOUBLE_EQ(Number(wb.handle, 0, 1), 3.0);
  EXPECT_DOUBLE_EQ(Number(wb.handle, 1, 1), 9.0);

  int32_t present = 0;
  fm_data_table_t info = FM_STRUCT_INIT(fm_data_table_t);
  ASSERT_EQ(fm_sheet_get_data_table(wb.handle, 0, 3, 1, &present, &info), 0);
  EXPECT_EQ(present, 1);
  EXPECT_EQ(info.rect.first_row, 2U);
  EXPECT_EQ(info.rect.last_row, 4U);
  EXPECT_EQ(info.rect.first_col, 1U);
  EXPECT_EQ(info.rect.last_col, 1U);
  EXPECT_EQ(info.has_column_input, 1);
  EXPECT_EQ(info.has_row_input, 0);
  EXPECT_EQ(info.column_input_row, 0U);
  EXPECT_EQ(info.column_input_col, 1U);
  EXPECT_EQ(info.two_dimensional, 0);
  EXPECT_EQ(info.row_oriented, 0);
  EXPECT_EQ(info.column_input_deleted, 0);

  // A cell outside the range reports no table and a zeroed record.
  ASSERT_EQ(fm_sheet_get_data_table(wb.handle, 0, 1, 1, &present, &info), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(info.rect.last_row, 0U);

  EXPECT_EQ(fm_sheet_set_number(wb.handle, 0, 3, 1, 7.0), kPreconditionFailed);
  EXPECT_EQ(fm_sheet_set_text(wb.handle, 0, 2, 1, "x"), kPreconditionFailed);
  EXPECT_EQ(fm_sheet_set_formula(wb.handle, 0, 4, 1, "=1"), kPreconditionFailed);
  EXPECT_DOUBLE_EQ(Number(wb.handle, 3, 1), 4.0);
  // The header cells stay editable and the table follows them.
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 3, 0, 10.0), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  EXPECT_DOUBLE_EQ(Number(wb.handle, 3, 1), 100.0);
}

TEST(FormulonCApiDataTable, TwoVariableTableReadsBothInputs) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 5, 1.0), 0);  // F1: row input.
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 1, 5, 1.0), 0);  // F2: column input.
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 2, 2, "=F1*F2"), 0);
  for (uint32_t j = 0; j < 3; ++j) {
    ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 2, 3 + j, static_cast<double>(j + 1)), 0);
    ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 3 + j, 2, static_cast<double>(10 * (j + 1))), 0);
  }
  fm_data_table_t spec = FM_STRUCT_INIT(fm_data_table_t);
  spec.rect.first_row = 3;
  spec.rect.first_col = 3;
  spec.rect.last_row = 5;
  spec.rect.last_col = 5;
  spec.has_row_input = 1;
  spec.row_input_row = 0;
  spec.row_input_col = 5;
  spec.has_column_input = 1;
  spec.column_input_row = 1;
  spec.column_input_col = 5;
  ASSERT_EQ(fm_sheet_add_data_table(wb.handle, 0, &spec, nullptr), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  EXPECT_DOUBLE_EQ(Number(wb.handle, 3, 3), 10.0);
  EXPECT_DOUBLE_EQ(Number(wb.handle, 4, 4), 40.0);
  EXPECT_DOUBLE_EQ(Number(wb.handle, 5, 5), 90.0);

  int32_t present = 0;
  fm_data_table_t info = FM_STRUCT_INIT(fm_data_table_t);
  ASSERT_EQ(fm_sheet_get_data_table(wb.handle, 0, 5, 5, &present, &info), 0);
  EXPECT_EQ(present, 1);
  EXPECT_EQ(info.two_dimensional, 1);
  EXPECT_EQ(info.has_row_input, 1);
  EXPECT_EQ(info.has_column_input, 1);
  EXPECT_EQ(info.row_input_row, 0U);
  EXPECT_EQ(info.column_input_row, 1U);

  // Deleting the row input's row flags it deleted and turns the table into #REF!.
  ASSERT_EQ(fm_sheet_delete_rows(wb.handle, 0, 0, 0), 0);
  ASSERT_EQ(fm_sheet_get_data_table(wb.handle, 0, 4, 3, &present, &info), 0);
  ASSERT_EQ(present, 1);
  EXPECT_EQ(info.row_input_deleted, 1);
  EXPECT_EQ(info.column_input_deleted, 0);
  EXPECT_EQ(info.rect.first_row, 2U);
  EXPECT_EQ(info.column_input_row, 0U);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  fm_value_t v{};
  ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, 2, 3, &v), 0);
  EXPECT_EQ(v.kind, FM_VALUE_KIND_ERROR);
}

TEST(FormulonCApiDataTable, RemoveBlanksTheResultCells) {
  WorkbookGuard wb;
  BuildSquares(wb);
  const fm_data_table_t spec = SquaresSpec();
  ASSERT_EQ(fm_sheet_add_data_table(wb.handle, 0, &spec, nullptr), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  ASSERT_EQ(fm_sheet_remove_data_table(wb.handle, 0, 4, 1), 0);
  fm_value_t v{};
  ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, 3, 1, &v), 0);
  EXPECT_EQ(v.kind, FM_VALUE_KIND_BLANK);
  int32_t present = 1;
  fm_data_table_t info = FM_STRUCT_INIT(fm_data_table_t);
  ASSERT_EQ(fm_sheet_get_data_table(wb.handle, 0, 3, 1, &present, &info), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(fm_sheet_remove_data_table(wb.handle, 0, 4, 1), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_sheet_set_number(wb.handle, 0, 3, 1, 7.0), 0);
}

TEST(FormulonCApiDataTable, AutoNoTableComputesOnlyOnTheExplicitCall) {
  WorkbookGuard wb;
  BuildSquares(wb);
  const fm_data_table_t spec = SquaresSpec();
  ASSERT_EQ(fm_sheet_add_data_table(wb.handle, 0, &spec, nullptr), 0);
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
  const fm_data_table_t spec = SquaresSpec();
  ASSERT_EQ(fm_sheet_add_data_table(wb.handle, 0, &spec, nullptr), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len), 0);

  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &loaded.handle), 0);
  int32_t present = 0;
  fm_data_table_t info = FM_STRUCT_INIT(fm_data_table_t);
  ASSERT_EQ(fm_sheet_get_data_table(loaded.handle, 0, 3, 1, &present, &info), 0);
  EXPECT_EQ(present, 1);
  EXPECT_EQ(info.has_column_input, 1);
  EXPECT_EQ(info.column_input_col, 1U);
  EXPECT_DOUBLE_EQ(Number(loaded.handle, 4, 1), 9.0);
  EXPECT_EQ(fm_sheet_set_number(loaded.handle, 0, 4, 1, 1.0), kPreconditionFailed);
}

TEST(FormulonCApiDataTable, ResultCellsReportTheTableFormula) {
  WorkbookGuard wb;
  BuildSquares(wb);
  const fm_data_table_t spec = SquaresSpec();
  ASSERT_EQ(fm_sheet_add_data_table(wb.handle, 0, &spec, nullptr), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  int32_t present = 0;
  const char* formula = nullptr;
  ASSERT_EQ(fm_sheet_get_formula(wb.handle, 0, 3, 1, &present, &formula), 0);
  EXPECT_EQ(present, 1);
  EXPECT_STREQ(formula, "=TABLE(,B1)");
  ASSERT_EQ(fm_sheet_get_formula_r1c1(wb.handle, 0, 3, 1, &present, &formula), 0);
  EXPECT_EQ(present, 1);
  EXPECT_STREQ(formula, "=TABLE(,R[-3]C)");
  ASSERT_EQ(fm_sheet_get_formula(wb.handle, 0, 3, 0, &present, &formula), 0);
  EXPECT_EQ(present, 0);  // A header value.
  EXPECT_EQ(formula, nullptr);

  fm_cell_range_t* range = nullptr;
  ASSERT_EQ(fm_sheet_get_cells_in_range(wb.handle, 0, record_arg(fm_cell_rect_t{2, 1, 4, 1}), 0, 0, &range), 0);
  size_t count = 0;
  ASSERT_EQ(fm_cell_range_get_cell_count(range, &count), 0);
  EXPECT_EQ(count, 3U);
  for (size_t i = 0; i < count; ++i) {
    uint32_t row = 0;
    uint32_t col = 0;
    fm_value_t v{};
    ASSERT_EQ(fm_cell_range_get_cell_at(range, i, &row, &col, &formula, &v), 0);
    EXPECT_STREQ(formula, "=TABLE(,B1)") << row;
  }
  fm_cell_range_destroy(range);

  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 6, 0, "=FORMULATEXT(B4)"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 6, 1, "=ISFORMULA(B4)"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  fm_value_t v{};
  ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, 6, 0, &v), 0);
  ASSERT_EQ(v.kind, FM_VALUE_KIND_TEXT);
  EXPECT_STREQ(v.u.text, "{=TABLE(,B1)}");
  ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, 6, 1, &v), 0);
  ASSERT_EQ(v.kind, FM_VALUE_KIND_BOOL);
  EXPECT_NE(v.u.boolean, 0);
}

TEST(FormulonCApiDataTable, StructuralEditsMoveOrRemoveTheTableWhole) {
  WorkbookGuard wb;
  BuildSquares(wb);
  const fm_data_table_t spec = SquaresSpec();
  ASSERT_EQ(fm_sheet_add_data_table(wb.handle, 0, &spec, nullptr), 0);
  // The table's block is A2:B5: the results B3:B5, the formula row 2 and the value column A.
  EXPECT_EQ(fm_sheet_delete_rows(wb.handle, 0, 3, 3), kPreconditionFailed);
  EXPECT_EQ(fm_sheet_delete_rows(wb.handle, 0, 2, 4), kPreconditionFailed);
  EXPECT_EQ(fm_sheet_insert_rows(wb.handle, 0, 4, 4), kPreconditionFailed);
  EXPECT_EQ(fm_sheet_delete_columns(wb.handle, 0, 0, 0), kPreconditionFailed);
  EXPECT_EQ(fm_sheet_insert_columns(wb.handle, 0, 1, 1), kPreconditionFailed);
  int32_t present = 0;
  fm_data_table_t info = FM_STRUCT_INIT(fm_data_table_t);
  ASSERT_EQ(fm_sheet_get_data_table(wb.handle, 0, 2, 1, &present, &info), 0);
  EXPECT_EQ(present, 1);
  EXPECT_EQ(info.rect.last_row, 4U);
  // Rows inserted above the block move it; deleting the whole block removes it.
  ASSERT_EQ(fm_sheet_insert_rows(wb.handle, 0, 1, 1), 0);
  ASSERT_EQ(fm_sheet_get_data_table(wb.handle, 0, 5, 1, &present, &info), 0);
  EXPECT_EQ(present, 1);
  EXPECT_EQ(info.rect.first_row, 3U);
  ASSERT_EQ(fm_sheet_delete_rows(wb.handle, 0, 2, 5), 0);
  ASSERT_EQ(fm_sheet_get_data_table(wb.handle, 0, 3, 1, &present, &info), 0);
  EXPECT_EQ(present, 0);
}

TEST(FormulonCApiDataTable, RejectsBadArguments) {
  WorkbookGuard wb;
  BuildSquares(wb);
  fm_data_table_t spec = SquaresSpec();
  int32_t present = 0;
  fm_data_table_t info = FM_STRUCT_INIT(fm_data_table_t);
  EXPECT_EQ(fm_sheet_add_data_table(wb.handle, 0, nullptr, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_add_data_table(wb.handle, 9, &spec, nullptr), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_get_data_table(wb.handle, 0, 0, 0, nullptr, &info), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_get_data_table(wb.handle, 0, 0, 0, &present, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_get_data_table(wb.handle, 9, 0, 0, &present, &info), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_remove_data_table(wb.handle, 9, 0, 0), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_workbook_recalc_data_tables(nullptr), kBindingNullPointer);

  fm_data_table_t none = spec;
  none.has_column_input = 0;
  EXPECT_EQ(fm_sheet_add_data_table(wb.handle, 0, &none, nullptr), kDataTableInvalid);
  fm_data_table_t no_header = spec;
  no_header.rect.first_row = 0;
  EXPECT_EQ(fm_sheet_add_data_table(wb.handle, 0, &no_header, nullptr), kDataTableInvalid);
  fm_data_table_t input_inside = spec;
  input_inside.column_input_row = 3;
  input_inside.column_input_col = 1;
  EXPECT_EQ(fm_sheet_add_data_table(wb.handle, 0, &input_inside, nullptr), kDataTableInvalid);
  fm_data_table_t off_grid = spec;
  off_grid.rect.last_row = 0xFFFFFFFFU;
  EXPECT_EQ(fm_sheet_add_data_table(wb.handle, 0, &off_grid, nullptr), kInvalidArgument);
  ASSERT_EQ(fm_sheet_add_data_table(wb.handle, 0, &spec, nullptr), 0);
  EXPECT_EQ(fm_sheet_add_data_table(wb.handle, 0, &spec, nullptr), kDataTableInvalid);
}

// Asserts a refused add left `*cell` zeroed.
void ExpectZeroCell(const fm_cell_address_t& cell) {
  EXPECT_EQ(cell.row, 0U);
  EXPECT_EQ(cell.col, 0U);
}

TEST(FormulonCApiDataTableAdd, SuccessWritesTopLeftAndReadsBack) {
  WorkbookGuard wb;
  BuildSquares(wb);
  const fm_data_table_t spec = SquaresSpec();
  fm_cell_address_t cell{9U, 9U};
  ASSERT_EQ(fm_sheet_add_data_table(wb.handle, 0, &spec, &cell), 0) << fm_get_last_error_message();
  EXPECT_EQ(cell.row, 2U);
  EXPECT_EQ(cell.col, 1U);

  int32_t present = 0;
  fm_data_table_t info = FM_STRUCT_INIT(fm_data_table_t);
  ASSERT_EQ(fm_sheet_get_data_table(wb.handle, 0, cell.row, cell.col, &present, &info), 0);
  ASSERT_EQ(present, 1);
  EXPECT_EQ(info.rect.first_row, 2U);
  EXPECT_EQ(info.rect.first_col, 1U);
  EXPECT_EQ(info.rect.last_row, 4U);
  EXPECT_EQ(info.rect.last_col, 1U);
  EXPECT_EQ(info.has_row_input, 0);
  EXPECT_EQ(info.has_column_input, 1);
  EXPECT_EQ(info.column_input_row, 0U);
  EXPECT_EQ(info.column_input_col, 1U);
  EXPECT_EQ(info.two_dimensional, 0);
  EXPECT_EQ(info.row_oriented, 0);
  EXPECT_EQ(info.row_input_deleted, 0);
  EXPECT_EQ(info.column_input_deleted, 0);

  // A two-variable table over C3:D4, inputs F1 (row) and G1 (column).
  fm_data_table_t two = FM_STRUCT_INIT(fm_data_table_t);
  two.rect = fm_cell_rect_t{2U, 2U, 3U, 3U};
  two.has_row_input = 1;
  two.row_input_row = 0;
  two.row_input_col = 5;
  two.has_column_input = 1;
  two.column_input_row = 0;
  two.column_input_col = 6;
  ASSERT_EQ(fm_sheet_add_data_table(wb.handle, 0, &two, &cell), 0) << fm_get_last_error_message();
  EXPECT_EQ(cell.row, 2U);
  EXPECT_EQ(cell.col, 2U);
  ASSERT_EQ(fm_sheet_get_data_table(wb.handle, 0, cell.row, cell.col, &present, &info), 0);
  ASSERT_EQ(present, 1);
  EXPECT_EQ(info.two_dimensional, 1);
  EXPECT_EQ(info.has_row_input, 1);
  EXPECT_EQ(info.row_input_col, 5U);
  EXPECT_EQ(info.has_column_input, 1);
  EXPECT_EQ(info.column_input_col, 6U);

  // An inverted range is stored normalised; the written cell is its top-left.
  fm_data_table_t inverted = SquaresSpec();
  inverted.rect = fm_cell_rect_t{8U, 4U, 6U, 4U};
  ASSERT_EQ(fm_sheet_add_data_table(wb.handle, 0, &inverted, &cell), 0) << fm_get_last_error_message();
  EXPECT_EQ(cell.row, 6U);
  EXPECT_EQ(cell.col, 4U);
}

TEST(FormulonCApiDataTableAdd, NullSpecIsNullPointer) {
  WorkbookGuard wb;
  BuildSquares(wb);
  fm_cell_address_t cell{9U, 9U};
  EXPECT_EQ(fm_sheet_add_data_table(wb.handle, 0, nullptr, &cell), kBindingNullPointer);
  ExpectZeroCell(cell);
}

TEST(FormulonCApiDataTableAdd, SheetOutOfRange) {
  WorkbookGuard wb;
  BuildSquares(wb);
  const fm_data_table_t spec = SquaresSpec();
  fm_cell_address_t cell{9U, 9U};
  EXPECT_EQ(fm_sheet_add_data_table(wb.handle, 9, &spec, &cell), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  ExpectZeroCell(cell);
}

TEST(FormulonCApiDataTableAdd, OffGridCoordinateInvalidArgument) {
  WorkbookGuard wb;
  BuildSquares(wb);
  fm_data_table_t spec = SquaresSpec();
  spec.rect.last_col = 0xFFFFFFFFU;
  fm_cell_address_t cell{9U, 9U};
  EXPECT_EQ(fm_sheet_add_data_table(wb.handle, 0, &spec, &cell), kInvalidArgument);
  ExpectZeroCell(cell);
  spec = SquaresSpec();
  spec.column_input_row = 0xFFFFFFFFU;
  EXPECT_EQ(fm_sheet_add_data_table(wb.handle, 0, &spec, &cell), kInvalidArgument);
}

TEST(FormulonCApiDataTableAdd, NoInputIsDataTableInvalid) {
  WorkbookGuard wb;
  BuildSquares(wb);
  fm_data_table_t spec = SquaresSpec();
  spec.has_column_input = 0;
  fm_cell_address_t cell{9U, 9U};
  EXPECT_EQ(fm_sheet_add_data_table(wb.handle, 0, &spec, &cell), kDataTableInvalid);
  ExpectZeroCell(cell);
}

TEST(FormulonCApiDataTableAdd, NoRowAboveOrColumnLeftIsDataTableInvalid) {
  WorkbookGuard wb;
  BuildSquares(wb);
  fm_data_table_t no_row_above = SquaresSpec();
  no_row_above.rect.first_row = 0;
  fm_cell_address_t cell{9U, 9U};
  EXPECT_EQ(fm_sheet_add_data_table(wb.handle, 0, &no_row_above, &cell), kDataTableInvalid);
  ExpectZeroCell(cell);
  fm_data_table_t no_column_left = SquaresSpec();
  no_column_left.rect.first_col = 0;
  no_column_left.rect.last_col = 0;
  EXPECT_EQ(fm_sheet_add_data_table(wb.handle, 0, &no_column_left, nullptr), kDataTableInvalid);
}

TEST(FormulonCApiDataTableAdd, InputInsideRangeIsDataTableInvalid) {
  WorkbookGuard wb;
  BuildSquares(wb);
  fm_data_table_t spec = SquaresSpec();
  spec.column_input_row = 3;
  spec.column_input_col = 1;
  fm_cell_address_t cell{9U, 9U};
  EXPECT_EQ(fm_sheet_add_data_table(wb.handle, 0, &spec, &cell), kDataTableInvalid);
  ExpectZeroCell(cell);
}

TEST(FormulonCApiDataTableAdd, OverlapIsDataTableInvalid) {
  WorkbookGuard wb;
  BuildSquares(wb);
  const fm_data_table_t spec = SquaresSpec();
  ASSERT_EQ(fm_sheet_add_data_table(wb.handle, 0, &spec, nullptr), 0) << fm_get_last_error_message();
  fm_data_table_t overlapping = SquaresSpec();
  overlapping.rect.first_row = 3;
  overlapping.rect.last_row = 6;
  fm_cell_address_t cell{9U, 9U};
  EXPECT_EQ(fm_sheet_add_data_table(wb.handle, 0, &overlapping, &cell), kDataTableInvalid);
  ExpectZeroCell(cell);
}

}  // namespace
