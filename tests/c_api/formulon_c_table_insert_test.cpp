// C ABI tests for `fm_sheet_insert_table_rows` / `fm_sheet_insert_table_columns`:
// the (sheet_index, per-sheet table_index) mapping, append and middle inserts,
// calculated-column fill, column naming and the argument and engine refusals.

#include <cstddef>
#include <cstdint>
#include <string>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "cell.h"
#include "formulon_c_out_index.h"
#include "formulon_c_record_arg.h"
#include "formulon_c_setters.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "table.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace {

constexpr fm_status_t kBindingNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);

/// Sheet1 with table tblSales over A1:C3 (Item, Qty, Double; two data rows holding 1/2 and 10/20), plus a
/// second sheet holding table tblOther over A1:B2 (Name, Val). `profile` NULL creates the default profile.
void BuildSales(WorkbookGuard& wb, const char* profile = nullptr) {
  if (profile == nullptr) {
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  } else {
    ASSERT_EQ(fm_workbook_create_for_profile(profile, &wb.handle), 0) << fm_get_last_error_message();
  }
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Sheet2", append_position(wb.handle), ignored_index()), 0);
  const char* columns[] = {"Item", "Qty", "Double"};
  for (std::uint32_t c = 0; c < 3U; ++c) {
    ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, c, columns[c]), 0);
  }
  for (std::uint32_t r = 1; r <= 2U; ++r) {
    ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, r, 0, static_cast<double>(r)), 0);
    ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, r, 1, static_cast<double>(10U * r)), 0);
  }
  std::size_t index = 0;
  ASSERT_EQ(AddTable(wb.handle, 0, record_arg(fm_cell_rect_t{0, 0, 2, 2}), "tblSales", columns, 3, "", 1, 0, &index), 0)
      << fm_get_last_error_message();
  const char* other[] = {"Name", "Val"};
  for (std::uint32_t c = 0; c < 2U; ++c) {
    ASSERT_EQ(fm_sheet_set_text(wb.handle, 1, 0, c, other[c]), 0);
  }
  ASSERT_EQ(AddTable(wb.handle, 1, record_arg(fm_cell_rect_t{0, 0, 1, 1}), "tblOther", other, 2, "", 1, 0, &index), 0)
      << fm_get_last_error_message();
}

fm_table_info_t Info(const WorkbookGuard& wb, std::size_t sheet, std::size_t table) {
  fm_table_info_t info = FM_STRUCT_INIT(fm_table_info_t);
  EXPECT_EQ(fm_sheet_get_table_at(wb.handle, sheet, table, &info), 0);
  return info;
}

std::string ColumnName(const WorkbookGuard& wb, std::size_t column) {
  fm_table_column_info_t info = FM_STRUCT_INIT(fm_table_column_info_t);
  EXPECT_EQ(fm_sheet_get_table_column_at(wb.handle, 0, 0, column, &info), 0);
  return info.name != nullptr ? info.name : "";
}

std::string Names(const WorkbookGuard& wb) {
  std::string out;
  const std::uint32_t count = Info(wb, 0, 0).column_count;
  for (std::size_t c = 0; c < count; ++c) {
    out += out.empty() ? "" : ",";
    out += ColumnName(wb, c);
  }
  return out;
}

/// The cell as text: "" blank, the integer for a number, the text otherwise.
std::string Show(const WorkbookGuard& wb, std::size_t sheet, std::uint32_t row, std::uint32_t col) {
  const formulon::Value value = wb.handle->workbook().sheet(sheet).resolve_cell_value(row, col);
  if (value.is_blank()) {
    return {};
  }
  if (value.is_number()) {
    return std::to_string(static_cast<long long>(value.as_number()));
  }
  return value.is_text() ? std::string(value.as_text()) : std::string("<other>");
}

std::string FormulaAt(const WorkbookGuard& wb, std::uint32_t row, std::uint32_t col) {
  const formulon::Cell* cell = wb.handle->workbook().sheet(0).cell_at(row, col);
  return cell != nullptr ? cell->formula_text : "";
}

TEST(CApiTableInsert, AppendRowGrowsTheTableAndMovesOnlyItsColumns) {
  WorkbookGuard wb;
  BuildSales(wb);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 3, 0, 99), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 3, 3, 77), 0);
  ASSERT_EQ(fm_sheet_insert_table_rows(wb.handle, 0, 0, 2, 2), 0) << fm_get_last_error_message();
  EXPECT_EQ(Info(wb, 0, 0).range.last_row, 3U);
  EXPECT_EQ(Show(wb, 0, 3, 0), "");
  EXPECT_EQ(Show(wb, 0, 4, 0), "99");
  EXPECT_EQ(Show(wb, 0, 3, 3), "77");
}

TEST(CApiTableInsert, MiddleSpanShiftsTheRowsBelow) {
  WorkbookGuard wb;
  BuildSales(wb);
  ASSERT_EQ(fm_sheet_insert_table_rows(wb.handle, 0, 0, 1, 2), 0) << fm_get_last_error_message();
  EXPECT_EQ(Info(wb, 0, 0).range.last_row, 4U);
  EXPECT_EQ(Show(wb, 0, 1, 0), "1");
  EXPECT_EQ(Show(wb, 0, 2, 0), "");
  EXPECT_EQ(Show(wb, 0, 3, 0), "");
  EXPECT_EQ(Show(wb, 0, 4, 0), "2");
}

TEST(CApiTableInsert, CalculatedColumnIsFilledAndOtherColumnsStayBlank) {
  WorkbookGuard wb;
  BuildSales(wb);
  wb.handle->workbook().mutable_tables()[0].columns[2].calculated_column_formula = "B2*2";
  ASSERT_EQ(fm_sheet_insert_table_rows(wb.handle, 0, 0, 1, 1), 0) << fm_get_last_error_message();
  EXPECT_EQ(FormulaAt(wb, 2, 2), "=B3*2");
  EXPECT_EQ(FormulaAt(wb, 2, 0), "");
  EXPECT_EQ(Show(wb, 0, 2, 1), "");
}

TEST(CApiTableInsert, TableIndexIsPerSheet) {
  WorkbookGuard wb;
  BuildSales(wb);
  ASSERT_EQ(fm_sheet_insert_table_rows(wb.handle, 1, 0, 1, 1), 0) << fm_get_last_error_message();
  EXPECT_EQ(Info(wb, 1, 0).range.last_row, 2U);
  EXPECT_EQ(Info(wb, 0, 0).range.last_row, 2U);
}

TEST(CApiTableInsert, ColumnNamesFollowTheLeftNeighbourAndTheProfileWord) {
  WorkbookGuard wb;
  BuildSales(wb, "mac-365-ja_JP");
  ASSERT_EQ(fm_sheet_insert_table_columns(wb.handle, 0, 0, 1, 1), 0) << fm_get_last_error_message();
  EXPECT_EQ(Names(wb), "Item,列1,Qty,Double");
  EXPECT_EQ(Show(wb, 0, 0, 1), "列1");
  EXPECT_EQ(Show(wb, 0, 1, 2), "10");
  ASSERT_EQ(fm_sheet_insert_table_columns(wb.handle, 0, 0, 0, 0), 0) << fm_get_last_error_message();
  EXPECT_EQ(ColumnName(wb, 0), "列2");
  EXPECT_EQ(Show(wb, 0, 0, 0), "列2");
  ASSERT_EQ(fm_sheet_insert_table_columns(wb.handle, 0, 0, 2, 2), 0) << fm_get_last_error_message();
  EXPECT_EQ(ColumnName(wb, 2), "列3");
}

TEST(CApiTableInsert, AppendColumnTakesTheNumberAfterItsLeftNeighbour) {
  WorkbookGuard wb;
  BuildSales(wb);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 2, "Q2"), 0);
  ASSERT_EQ(fm_sheet_insert_table_columns(wb.handle, 0, 0, 3, 4), 0) << fm_get_last_error_message();
  EXPECT_EQ(Names(wb), "Item,Qty,Q2,Q3,Q4");
  EXPECT_EQ(Info(wb, 0, 0).range.last_col, 4U);
  EXPECT_EQ(Show(wb, 0, 0, 4), "Q4");
}

TEST(CApiTableInsert, ColumnInsertLeavesCellsOfOtherRowsInPlace) {
  WorkbookGuard wb;
  BuildSales(wb);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 3, 1, 55), 0);
  ASSERT_EQ(fm_sheet_insert_table_columns(wb.handle, 0, 0, 1, 1), 0) << fm_get_last_error_message();
  EXPECT_EQ(Show(wb, 0, 3, 1), "55");
  EXPECT_EQ(Show(wb, 0, 1, 2), "10");
}

TEST(CApiTableInsert, ArgumentErrors) {
  WorkbookGuard wb;
  BuildSales(wb);
  EXPECT_EQ(fm_sheet_insert_table_rows(nullptr, 0, 0, 0, 0), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_insert_table_columns(nullptr, 0, 0, 0, 0), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_insert_table_rows(wb.handle, 9, 0, 0, 0), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_insert_table_rows(wb.handle, 0, 1, 0, 0), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_insert_table_columns(wb.handle, 0, 1, 0, 0), FM_STATUS_INDEX_OUT_OF_RANGE);
  // first past the count, last below first, span past the grid.
  const auto message = [] {
    const char* text = fm_get_last_error_message();
    return text == nullptr ? std::string() : std::string(text);
  };
  EXPECT_EQ(fm_sheet_insert_table_rows(wb.handle, 0, 0, 3, 3), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(message(), "first 3 is out of range (2 rows)");
  EXPECT_EQ(fm_sheet_insert_table_rows(wb.handle, 0, 0, 0xFFFFFFFFU, 0xFFFFFFFFU), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_insert_table_columns(wb.handle, 0, 0, 0xFFFFFFFFU, 0xFFFFFFFFU), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_insert_table_rows(wb.handle, 0, 0, 2, 1), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(message(), "last 1 is before first 2");
  EXPECT_EQ(fm_sheet_insert_table_rows(wb.handle, 0, 0, 2, 1048576U), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(message(), "last 1048576 is outside the grid");
  EXPECT_EQ(fm_sheet_insert_table_columns(wb.handle, 0, 0, 4, 4), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(message(), "first 4 is out of range (3 columns)");
  EXPECT_EQ(fm_sheet_insert_table_columns(wb.handle, 0, 0, 3, 2), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(message(), "last 2 is before first 3");
  EXPECT_EQ(fm_sheet_insert_table_columns(wb.handle, 0, 0, 3, 16384U), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(message(), "last 16384 is outside the grid");
  EXPECT_EQ(Info(wb, 0, 0).range.last_row, 2U);
  EXPECT_EQ(Info(wb, 0, 0).range.last_col, 2U);
}

TEST(CApiTableInsert, ContentPushedOffTheGridIsRefused) {
  WorkbookGuard wb;
  BuildSales(wb);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 1048575, 1, 1), 0);
  EXPECT_EQ(fm_sheet_insert_table_rows(wb.handle, 0, 0, 2, 2), FM_STATUS_PRECONDITION_FAILED);
  EXPECT_STREQ(fm_get_last_error_message(), "Cannot shift a value or formula off the sheet");
  EXPECT_EQ(Info(wb, 0, 0).range.last_row, 2U);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 1, 16383, 1), 0);
  EXPECT_EQ(fm_sheet_insert_table_columns(wb.handle, 0, 0, 3, 3), FM_STATUS_PRECONDITION_FAILED);
  EXPECT_EQ(Info(wb, 0, 0).range.last_col, 2U);
}

}  // namespace
