//
// C ABI tests for the AutoFilter of a worksheet table, addressed as
// (sheet_index, table_index): removal through a NULL record, clearing the
// criteria, and the common table failures.

#include <cstddef>
#include <cstdint>
#include <set>
#include <vector>

#include "c_api/formulon_c.h"
#include "formulon_c_record_arg.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"
#include "utils/error.h"

namespace {

constexpr fm_status_t kBindingNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);
constexpr fm_status_t kNotFound = static_cast<fm_status_t>(formulon::FormulonErrorCode::kNotFound);

std::set<std::uint32_t> HiddenRows(const fm_workbook_t* wb, std::size_t sheet = 0U) {
  std::set<std::uint32_t> rows;
  std::size_t count = 0;
  EXPECT_EQ(fm_sheet_get_row_layout_count(wb, sheet, &count), 0);
  for (std::size_t i = 0; i < count; ++i) {
    fm_row_layout_t layout = FM_STRUCT_INIT(fm_row_layout_t);
    EXPECT_EQ(fm_sheet_get_row_layout_at(wb, sheet, i, &layout), 0);
    if (layout.hidden != 0) {
      rows.insert(layout.row);
    }
  }
  return rows;
}

/// Sheet1 holding header `N` and 1..5, with table `Nums` over A1:A6.
void CreateNumbersTable(WorkbookGuard& wb) {
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 0, "N"), 0);
  for (std::uint32_t i = 1; i <= 5; ++i) {
    ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, i, 0, static_cast<double>(i)), 0);
  }
  const char* names[] = {"N"};
  std::size_t table = 99;
  ASSERT_EQ(
      fm_sheet_add_table(wb.handle, 0, record_arg(fm_cell_rect_t{0, 0, 5, 0}), "Nums", names, 1, nullptr, 1, 0, &table),
      0)
      << fm_get_last_error_message();
  ASSERT_EQ(table, 0U);
}

/// Keeps the rows greater than 3 (custom filter `greaterThan 3` on column 0).
fm_auto_filter_t GreaterThanThree(fm_filter_column_t* column) {
  *column = fm_filter_column_t{};
  column->kind = 2;
  column->show_button = 1;
  column->custom_count = 1;
  column->op1 = 5;
  column->val1 = "3";
  fm_auto_filter_t filter = FM_STRUCT_INIT(fm_auto_filter_t);
  filter.range = fm_cell_rect_t{0, 0, 5, 0};
  filter.columns = column;
  filter.column_count = 1;
  return filter;
}

TEST(FormulonCApiTableAutoFilter, SetNullRemovesTheFilterAndIsANoOpWhenAbsent) {
  WorkbookGuard wb;
  CreateNumbersTable(wb);
  fm_auto_filter_t out = FM_STRUCT_INIT(fm_auto_filter_t);
  int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_table_auto_filter(wb.handle, 0, 0, &out, &present), 0);
  EXPECT_EQ(present, 1) << "a created table carries its filter definition";

  ASSERT_EQ(fm_sheet_set_table_auto_filter(wb.handle, 0, 0, nullptr), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_table_auto_filter(wb.handle, 0, 0, &out, &present), 0);
  EXPECT_EQ(present, 0);

  EXPECT_EQ(fm_sheet_set_table_auto_filter(wb.handle, 0, 0, nullptr), 0) << "no-op when the table has none";
  ASSERT_EQ(fm_sheet_get_table_auto_filter(wb.handle, 0, 0, &out, &present), 0);
  EXPECT_EQ(present, 0);
}

TEST(FormulonCApiTableAutoFilter, SetNullLeavesTheRowVisibilityAlone) {
  WorkbookGuard wb;
  CreateNumbersTable(wb);
  fm_filter_column_t column;
  const fm_auto_filter_t filter = GreaterThanThree(&column);
  ASSERT_EQ(fm_sheet_set_table_auto_filter(wb.handle, 0, 0, &filter), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_apply_table_auto_filter(wb.handle, 0, 0), 0);
  ASSERT_EQ(HiddenRows(wb.handle), (std::set<std::uint32_t>{1, 2, 3}));

  ASSERT_EQ(fm_sheet_set_table_auto_filter(wb.handle, 0, 0, nullptr), 0);
  EXPECT_EQ(HiddenRows(wb.handle), (std::set<std::uint32_t>{1, 2, 3}));
}

TEST(FormulonCApiTableAutoFilter, ClearCriteriaShowsTheRowsAndKeepsTheFilter) {
  WorkbookGuard wb;
  CreateNumbersTable(wb);
  fm_filter_column_t column;
  const fm_auto_filter_t filter = GreaterThanThree(&column);
  ASSERT_EQ(fm_sheet_set_table_auto_filter(wb.handle, 0, 0, &filter), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_apply_table_auto_filter(wb.handle, 0, 0), 0);
  ASSERT_EQ(HiddenRows(wb.handle), (std::set<std::uint32_t>{1, 2, 3}));

  ASSERT_EQ(fm_sheet_clear_table_auto_filter_criteria(wb.handle, 0, 0), 0) << fm_get_last_error_message();
  EXPECT_TRUE(HiddenRows(wb.handle).empty());
  fm_auto_filter_t out = FM_STRUCT_INIT(fm_auto_filter_t);
  int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_table_auto_filter(wb.handle, 0, 0, &out, &present), 0);
  EXPECT_EQ(present, 1);
  EXPECT_EQ(out.column_count, 0U);
}

TEST(FormulonCApiTableAutoFilter, ClearCriteriaIsANoOpOnAFilterWithoutCriteria) {
  WorkbookGuard wb;
  CreateNumbersTable(wb);
  EXPECT_EQ(fm_sheet_clear_table_auto_filter_criteria(wb.handle, 0, 0), 0);
  EXPECT_EQ(fm_sheet_clear_table_auto_filter_criteria(wb.handle, 0, 0), 0);
  EXPECT_TRUE(HiddenRows(wb.handle).empty());
  fm_auto_filter_t out = FM_STRUCT_INIT(fm_auto_filter_t);
  int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_table_auto_filter(wb.handle, 0, 0, &out, &present), 0);
  EXPECT_EQ(present, 1);
}

TEST(FormulonCApiTableAutoFilter, ClearCriteriaOnATableWithoutAFilterIsNotFound) {
  WorkbookGuard wb;
  CreateNumbersTable(wb);
  ASSERT_EQ(fm_sheet_set_table_auto_filter(wb.handle, 0, 0, nullptr), 0);
  EXPECT_EQ(fm_sheet_clear_table_auto_filter_criteria(wb.handle, 0, 0), kNotFound);
  EXPECT_EQ(fm_sheet_apply_table_auto_filter(wb.handle, 0, 0), kNotFound);
}

TEST(FormulonCApiTableAutoFilter, TableFailuresAreSheetThenTableIndexOutOfRange) {
  WorkbookGuard wb;
  CreateNumbersTable(wb);
  fm_filter_column_t column;
  const fm_auto_filter_t filter = GreaterThanThree(&column);
  fm_auto_filter_t out = FM_STRUCT_INIT(fm_auto_filter_t);
  int32_t present = 0;
  std::size_t count = 0;
  std::uint32_t first_row = 0;

  EXPECT_EQ(fm_sheet_get_table_auto_filter(wb.handle, 0, 4, &out, &present), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_set_table_auto_filter(wb.handle, 0, 4, &filter), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_set_table_auto_filter(wb.handle, 0, 4, nullptr), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_apply_table_auto_filter(wb.handle, 0, 4), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_clear_table_auto_filter_criteria(wb.handle, 0, 4), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_evaluate_table_auto_filter(wb.handle, 0, 4, nullptr, 0, &count, &first_row),
            FM_STATUS_INDEX_OUT_OF_RANGE);

  EXPECT_EQ(fm_sheet_get_table_auto_filter(wb.handle, 9, 0, &out, &present), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_set_table_auto_filter(wb.handle, 9, 0, &filter), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_apply_table_auto_filter(wb.handle, 9, 0), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_clear_table_auto_filter_criteria(wb.handle, 9, 0), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_evaluate_table_auto_filter(wb.handle, 9, 0, nullptr, 0, &count, &first_row),
            FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);

  EXPECT_EQ(fm_sheet_clear_table_auto_filter_criteria(nullptr, 0, 0), kBindingNullPointer);
}

TEST(FormulonCApiTableAutoFilter, FilterOfATableOnASecondSheetHidesThatSheetsRows) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t second = 99;
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Second", &second), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, second, 0, 0, "N"), 0);
  for (std::uint32_t i = 1; i <= 5; ++i) {
    ASSERT_EQ(fm_sheet_set_number(wb.handle, second, i, 0, static_cast<double>(i)), 0);
  }
  const char* names[] = {"N"};
  std::size_t table = 99;
  ASSERT_EQ(fm_sheet_add_table(wb.handle, second, record_arg(fm_cell_rect_t{0, 0, 5, 0}), "OnSecond", names, 1, nullptr,
                               1, 0, &table),
            0)
      << fm_get_last_error_message();
  ASSERT_EQ(table, 0U);

  fm_filter_column_t column;
  const fm_auto_filter_t filter = GreaterThanThree(&column);
  ASSERT_EQ(fm_sheet_set_table_auto_filter(wb.handle, second, table, &filter), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_apply_table_auto_filter(wb.handle, second, table), 0);
  EXPECT_EQ(HiddenRows(wb.handle, second), (std::set<std::uint32_t>{1, 2, 3}));
  EXPECT_TRUE(HiddenRows(wb.handle, 0).empty());

  std::size_t count = 0;
  std::uint32_t first_row = 0;
  ASSERT_EQ(fm_sheet_evaluate_table_auto_filter(wb.handle, second, table, nullptr, 0, &count, &first_row), 0);
  EXPECT_EQ(count, 5U);
  std::vector<std::uint8_t> match(count, 9U);
  ASSERT_EQ(
      fm_sheet_evaluate_table_auto_filter(wb.handle, second, table, match.data(), match.size(), &count, &first_row), 0);
  EXPECT_EQ(match, (std::vector<std::uint8_t>{0, 0, 0, 1, 1}));
}

}  // namespace
