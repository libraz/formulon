// C ABI input records: a defined name read and written as one record, the
// table, pivot and reference specs, a pivot cache's worksheet source as one
// record, and the AutoFilter custom-condition list.

#include <cstddef>
#include <cstdint>
#include <string>

#include "c_api/formulon_c.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"

namespace {

std::size_t TableCount(const fm_workbook_t* wb) {
  std::size_t count = 99;
  EXPECT_EQ(fm_sheet_get_table_count(wb, 0, &count), FM_STATUS_OK);
  return count;
}

/// Sheet1 with headers Product / Amount and two data rows, plus a total row slot.
void FillSales(fm_workbook_t* wb) {
  ASSERT_EQ(fm_sheet_set_text(wb, 0, 0, 0, "Product"), FM_STATUS_OK);
  ASSERT_EQ(fm_sheet_set_text(wb, 0, 0, 1, "Amount"), FM_STATUS_OK);
  ASSERT_EQ(fm_sheet_set_text(wb, 0, 1, 0, "Widget"), FM_STATUS_OK);
  ASSERT_EQ(fm_sheet_set_number(wb, 0, 1, 1, 10.0), FM_STATUS_OK);
  ASSERT_EQ(fm_sheet_set_text(wb, 0, 2, 0, "Gadget"), FM_STATUS_OK);
  ASSERT_EQ(fm_sheet_set_number(wb, 0, 2, 1, 20.0), FM_STATUS_OK);
}

fm_table_spec_t SalesSpec() {
  fm_table_spec_t spec = FM_STRUCT_INIT(fm_table_spec_t);
  spec.range = fm_cell_rect_t{0, 0, 2, 1};
  spec.name = "Sales";
  return spec;
}

fm_table_style_options_t TableStyle(const fm_workbook_t* wb, std::size_t table = 0, int32_t* present = nullptr) {
  fm_table_style_options_t style = FM_STRUCT_INIT(fm_table_style_options_t);
  int32_t stated = 0;
  EXPECT_EQ(fm_sheet_get_table_style_options(wb, 0, table, &stated, &style), FM_STATUS_OK);
  if (present != nullptr) {
    *present = stated;
  }
  return style;
}

int32_t TableStylePresent(const fm_workbook_t* wb, std::size_t table = 0) {
  int32_t present = 0;
  TableStyle(wb, table, &present);
  return present;
}

fm_table_info_t TableInfo(const fm_workbook_t* wb, std::size_t table = 0) {
  fm_table_info_t info = FM_STRUCT_INIT(fm_table_info_t);
  EXPECT_EQ(fm_sheet_get_table_at(wb, 0, table, &info), FM_STATUS_OK);
  return info;
}

// ---------------------------------------------------------------------------
// Defined name
// ---------------------------------------------------------------------------

TEST(FormulonCApiInputRecords, DefinedNameReadRecordGoesBackThroughTheSetterUnchanged) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  const fm_defined_name_t original{sizeof(fm_defined_name_t), "Rate", "=Sheet1!$A$1", 0, "the rate", 1};
  ASSERT_EQ(fm_workbook_set_defined_name(wb.handle, &original), FM_STATUS_OK) << fm_get_last_error_message();

  fm_defined_name_t read = FM_STRUCT_INIT(fm_defined_name_t);
  ASSERT_EQ(fm_workbook_get_defined_name_at(wb.handle, 0, &read), FM_STATUS_OK);
  EXPECT_STREQ(read.name, "Rate");
  EXPECT_STREQ(read.formula, "=Sheet1!$A$1");
  EXPECT_EQ(read.local_sheet_id, 0);
  EXPECT_STREQ(read.comment, "the rate");
  EXPECT_EQ(read.hidden, 1);

  // The record as read is a valid input.
  ASSERT_EQ(fm_workbook_set_defined_name(wb.handle, &read), FM_STATUS_OK) << fm_get_last_error_message();
  std::size_t count = 0;
  ASSERT_EQ(fm_workbook_get_defined_name_count(wb.handle, &count), FM_STATUS_OK);
  EXPECT_EQ(count, 1U);
  fm_defined_name_t again = FM_STRUCT_INIT(fm_defined_name_t);
  ASSERT_EQ(fm_workbook_get_defined_name_at(wb.handle, 0, &again), FM_STATUS_OK);
  EXPECT_STREQ(again.name, "Rate");
  EXPECT_STREQ(again.formula, "=Sheet1!$A$1");
  EXPECT_EQ(again.local_sheet_id, 0);
  EXPECT_STREQ(again.comment, "the rate");
  EXPECT_EQ(again.hidden, 1);
}

TEST(FormulonCApiInputRecords, DefinedNameByNameFillsTheRecordTheIndexedReadFills) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  const fm_defined_name_t rate{sizeof(fm_defined_name_t), "Rate", "0.25", -1, "", 0};
  ASSERT_EQ(fm_workbook_set_defined_name(wb.handle, &rate), FM_STATUS_OK) << fm_get_last_error_message();

  fm_defined_name_t by_index = FM_STRUCT_INIT(fm_defined_name_t);
  ASSERT_EQ(fm_workbook_get_defined_name_at(wb.handle, 0, &by_index), FM_STATUS_OK);
  const std::string index_formula = by_index.formula;
  // An empty comment is no comment, which reads NULL.
  EXPECT_EQ(by_index.comment, nullptr);

  std::int32_t present = 0;
  fm_defined_name_t by_name = FM_STRUCT_INIT(fm_defined_name_t);
  ASSERT_EQ(fm_workbook_get_defined_name(wb.handle, "RATE", FM_DEFINED_NAME_SCOPE_WORKBOOK, &present, &by_name),
            FM_STATUS_OK);
  ASSERT_EQ(present, 1);
  EXPECT_STREQ(by_name.name, "Rate");
  EXPECT_EQ(index_formula, by_name.formula);
  EXPECT_EQ(by_name.comment, nullptr);
  EXPECT_EQ(by_name.local_sheet_id, by_index.local_sheet_id);
  EXPECT_EQ(by_name.hidden, by_index.hidden);

  // An absent name is a successful read with nothing in the record.
  by_name = FM_STRUCT_INIT(fm_defined_name_t);
  ASSERT_EQ(fm_workbook_get_defined_name(wb.handle, "Absent", FM_DEFINED_NAME_SCOPE_WORKBOOK, &present, &by_name),
            FM_STATUS_OK);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(by_name.name, nullptr);
  EXPECT_EQ(by_name.formula, nullptr);
  EXPECT_EQ(by_name.hidden, 0);
  // The same name in another scope is absent too.
  ASSERT_EQ(fm_workbook_get_defined_name(wb.handle, "Rate", 0, &present, &by_name), FM_STATUS_OK);
  EXPECT_EQ(present, 0);
}

TEST(FormulonCApiInputRecords, DefinedNameSetterRejectsMalformedRecords) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  EXPECT_EQ(fm_workbook_set_defined_name(wb.handle, nullptr), FM_STATUS_BINDING_NULL_POINTER);
  fm_defined_name_t name = FM_STRUCT_INIT(fm_defined_name_t);
  name.name = nullptr;
  name.formula = "1";
  name.local_sheet_id = FM_DEFINED_NAME_SCOPE_WORKBOOK;
  EXPECT_EQ(fm_workbook_set_defined_name(wb.handle, &name), FM_STATUS_BINDING_NULL_POINTER);
  name.name = "Rate";
  name.formula = nullptr;
  EXPECT_EQ(fm_workbook_set_defined_name(wb.handle, &name), FM_STATUS_BINDING_NULL_POINTER);
  name.formula = "";
  EXPECT_EQ(fm_workbook_set_defined_name(wb.handle, &name), FM_STATUS_INVALID_ARGUMENT);
  name.formula = "1";
  name.local_sheet_id = 5;
  EXPECT_EQ(fm_workbook_set_defined_name(wb.handle, &name), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  name.local_sheet_id = -2;
  EXPECT_EQ(fm_workbook_set_defined_name(wb.handle, &name), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  std::size_t count = 9;
  ASSERT_EQ(fm_workbook_get_defined_name_count(wb.handle, &count), FM_STATUS_OK);
  EXPECT_EQ(count, 0U);
}

// ---------------------------------------------------------------------------
// Table spec
// ---------------------------------------------------------------------------

TEST(FormulonCApiInputRecords, TableSpecWithRangeAndNameGivesAHeaderRowAndTheDefaultStyle) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  FillSales(wb.handle);
  const fm_table_spec_t spec = SalesSpec();
  std::size_t index = 99;
  ASSERT_EQ(fm_sheet_add_table(wb.handle, 0, &spec, &index), FM_STATUS_OK) << fm_get_last_error_message();
  EXPECT_EQ(index, 0U);

  const fm_table_info_t info = TableInfo(wb.handle);
  EXPECT_EQ(info.header_row_count, 1U);
  EXPECT_EQ(info.totals_row_count, 0U);
  const fm_table_style_options_t style = TableStyle(wb.handle);
  EXPECT_EQ(TableStylePresent(wb.handle), 1);
  ASSERT_NE(style.style_name, nullptr);
  EXPECT_EQ(style.show_row_stripes, 1);
  EXPECT_EQ(style.show_first_column, 0);
  EXPECT_EQ(style.show_last_column, 0);
  EXPECT_EQ(style.show_column_stripes, 0);
}

TEST(FormulonCApiInputRecords, TableSpecWritesAllFourStyleTogglesInOneCall) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  FillSales(wb.handle);
  fm_table_spec_t spec = SalesSpec();
  spec.style_name = "TableStyleMedium9";
  spec.show_first_column = FM_TRISTATE_TRUE;
  spec.show_last_column = FM_TRISTATE_TRUE;
  spec.show_row_stripes = FM_TRISTATE_FALSE;
  spec.show_column_stripes = FM_TRISTATE_TRUE;
  std::size_t index = 99;
  ASSERT_EQ(fm_sheet_add_table(wb.handle, 0, &spec, &index), FM_STATUS_OK) << fm_get_last_error_message();

  const fm_table_style_options_t style = TableStyle(wb.handle);
  EXPECT_EQ(TableStylePresent(wb.handle), 1);
  EXPECT_STREQ(style.style_name, "TableStyleMedium9");
  EXPECT_EQ(style.show_first_column, 1);
  EXPECT_EQ(style.show_last_column, 1);
  EXPECT_EQ(style.show_row_stripes, 0);
  EXPECT_EQ(style.show_column_stripes, 1);

  // The explicit values survive a save and load.
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), FM_STATUS_OK) << fm_get_last_error_message();
  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), FM_STATUS_OK) << fm_get_last_error_message();
  const fm_table_style_options_t again = TableStyle(reloaded.handle);
  EXPECT_EQ(again.show_first_column, 1);
  EXPECT_EQ(again.show_last_column, 1);
  EXPECT_EQ(again.show_row_stripes, 0);
  EXPECT_EQ(again.show_column_stripes, 1);
}

TEST(FormulonCApiInputRecords, TableSpecStatesTheHeaderAndTotalsRowsAndAnEmptyStyle) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  FillSales(wb.handle);
  fm_table_spec_t spec = SalesSpec();
  spec.range = fm_cell_rect_t{0, 0, 3, 1};
  spec.header_row = FM_TRISTATE_TRUE;
  spec.totals_row = FM_TRISTATE_TRUE;
  spec.style_name = "";
  std::size_t index = 99;
  ASSERT_EQ(fm_sheet_add_table(wb.handle, 0, &spec, &index), FM_STATUS_OK) << fm_get_last_error_message();
  const fm_table_info_t info = TableInfo(wb.handle);
  EXPECT_EQ(info.header_row_count, 1U);
  EXPECT_EQ(info.totals_row_count, 1U);
  EXPECT_EQ(TableStylePresent(wb.handle), 0);
}

TEST(FormulonCApiInputRecords, TableSpecRefusesAFlagOutsideTheTristateAndCreatesNothing) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  FillSales(wb.handle);
  std::size_t index = 99;
  for (int member = 0; member < 6; ++member) {
    for (const std::int32_t bad : {3, -1}) {
      fm_table_spec_t spec = SalesSpec();
      switch (member) {
        case 0:
          spec.header_row = bad;
          break;
        case 1:
          spec.totals_row = bad;
          break;
        case 2:
          spec.show_first_column = bad;
          break;
        case 3:
          spec.show_last_column = bad;
          break;
        case 4:
          spec.show_row_stripes = bad;
          break;
        default:
          spec.show_column_stripes = bad;
          break;
      }
      EXPECT_EQ(fm_sheet_add_table(wb.handle, 0, &spec, &index), FM_STATUS_INVALID_ARGUMENT)
          << "member " << member << " value " << bad;
      EXPECT_EQ(TableCount(wb.handle), 0U);
    }
  }
}

TEST(FormulonCApiInputRecords, TableSpecRefusesMissingPartsAndATakenName) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  FillSales(wb.handle);
  std::size_t index = 99;
  EXPECT_EQ(fm_sheet_add_table(wb.handle, 0, nullptr, &index), FM_STATUS_BINDING_NULL_POINTER);
  fm_table_spec_t spec = SalesSpec();
  EXPECT_EQ(fm_sheet_add_table(wb.handle, 0, &spec, nullptr), FM_STATUS_BINDING_NULL_POINTER);
  spec.name = nullptr;
  EXPECT_EQ(fm_sheet_add_table(wb.handle, 0, &spec, &index), FM_STATUS_BINDING_NULL_POINTER);
  spec = SalesSpec();
  EXPECT_EQ(fm_sheet_add_table(wb.handle, 9, &spec, &index), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  spec.range = fm_cell_rect_t{2, 0, 0, 1};
  EXPECT_EQ(fm_sheet_add_table(wb.handle, 0, &spec, &index), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(TableCount(wb.handle), 0U);

  spec = SalesSpec();
  ASSERT_EQ(fm_sheet_add_table(wb.handle, 0, &spec, &index), FM_STATUS_OK) << fm_get_last_error_message();
  const char* const other_columns[] = {"P", "Q"};
  spec.range = fm_cell_rect_t{0, 4, 2, 5};
  spec.name = "sales";
  spec.column_names = other_columns;
  spec.column_count = 2;
  EXPECT_EQ(fm_sheet_add_table(wb.handle, 0, &spec, &index), FM_STATUS_ALREADY_EXISTS);
  EXPECT_EQ(TableCount(wb.handle), 1U);
}

// ---------------------------------------------------------------------------
// Pivot spec
// ---------------------------------------------------------------------------

TEST(FormulonCApiInputRecords, PivotSpecCreatesATableAndRefusesMissingParts) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  std::uint32_t cache_id = 0;
  ASSERT_EQ(fm_workbook_add_pivot_cache(wb.handle, 0, &cache_id), FM_STATUS_OK) << fm_get_last_error_message();

  std::size_t pivot_index = 99;
  EXPECT_EQ(fm_sheet_add_pivot(wb.handle, 0, nullptr, &pivot_index), FM_STATUS_BINDING_NULL_POINTER);
  fm_pivot_spec_t spec = FM_STRUCT_INIT(fm_pivot_spec_t);
  spec.name = nullptr;
  spec.cache_id = cache_id;
  EXPECT_EQ(fm_sheet_add_pivot(wb.handle, 0, &spec, &pivot_index), FM_STATUS_BINDING_NULL_POINTER);
  spec.name = "PT";
  EXPECT_EQ(fm_sheet_add_pivot(wb.handle, 0, &spec, nullptr), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_add_pivot(wb.handle, 9, &spec, &pivot_index), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  std::size_t count = 99;
  ASSERT_EQ(fm_sheet_get_pivot_count(wb.handle, 0, &count), FM_STATUS_OK);
  EXPECT_EQ(count, 0U);

  spec.anchor_row = 2;
  spec.anchor_col = 5;
  ASSERT_EQ(fm_sheet_add_pivot(wb.handle, 0, &spec, &pivot_index), FM_STATUS_OK) << fm_get_last_error_message();
  EXPECT_EQ(pivot_index, 0U);
  ASSERT_EQ(fm_sheet_get_pivot_count(wb.handle, 0, &count), FM_STATUS_OK);
  EXPECT_EQ(count, 1U);
}

// ---------------------------------------------------------------------------
// Pivot cache worksheet source
// ---------------------------------------------------------------------------

TEST(FormulonCApiInputRecords, WorksheetSourceReadsPresenceAndSetsOrClearsAsOneRecord) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  std::uint32_t cache_id = 0;
  ASSERT_EQ(fm_workbook_add_pivot_cache(wb.handle, 0, &cache_id), FM_STATUS_OK) << fm_get_last_error_message();

  std::int32_t present = -1;
  fm_pivot_worksheet_source_t source = FM_STRUCT_INIT(fm_pivot_worksheet_source_t);
  ASSERT_EQ(fm_workbook_get_pivot_cache_worksheet_source(wb.handle, cache_id, &present, &source), FM_STATUS_OK);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(source.ref, nullptr);
  EXPECT_EQ(source.sheet, nullptr);
  EXPECT_EQ(source.name, nullptr);

  // A sheet the workbook does not have is stored as written.
  fm_pivot_worksheet_source_t set = FM_STRUCT_INIT(fm_pivot_worksheet_source_t);
  set.ref = "$A$1:$C$5";
  set.sheet = "Elsewhere";
  ASSERT_EQ(fm_workbook_set_pivot_cache_worksheet_source(wb.handle, cache_id, &set), FM_STATUS_OK)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_get_pivot_cache_worksheet_source(wb.handle, cache_id, &present, &source), FM_STATUS_OK);
  EXPECT_EQ(present, 1);
  EXPECT_STREQ(source.ref, "$A$1:$C$5");
  EXPECT_STREQ(source.sheet, "Elsewhere");
  EXPECT_EQ(source.name, nullptr);

  // The record as read is a valid input; a name replaces the reference.
  fm_pivot_worksheet_source_t named = FM_STRUCT_INIT(fm_pivot_worksheet_source_t);
  named.name = "Sales";
  ASSERT_EQ(fm_workbook_set_pivot_cache_worksheet_source(wb.handle, cache_id, &named), FM_STATUS_OK)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_get_pivot_cache_worksheet_source(wb.handle, cache_id, &present, &source), FM_STATUS_OK);
  EXPECT_EQ(present, 1);
  EXPECT_EQ(source.ref, nullptr);
  EXPECT_EQ(source.sheet, nullptr);
  EXPECT_STREQ(source.name, "Sales");
  ASSERT_EQ(fm_workbook_set_pivot_cache_worksheet_source(wb.handle, cache_id, &source), FM_STATUS_OK)
      << fm_get_last_error_message();

  // NULL clears it.
  ASSERT_EQ(fm_workbook_set_pivot_cache_worksheet_source(wb.handle, cache_id, nullptr), FM_STATUS_OK);
  ASSERT_EQ(fm_workbook_get_pivot_cache_worksheet_source(wb.handle, cache_id, &present, &source), FM_STATUS_OK);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(source.ref, nullptr);
  EXPECT_EQ(source.name, nullptr);
}

TEST(FormulonCApiInputRecords, WorksheetSourceRefusesMissingOutputsAndAnUnknownCache) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  std::uint32_t cache_id = 0;
  ASSERT_EQ(fm_workbook_add_pivot_cache(wb.handle, 0, &cache_id), FM_STATUS_OK) << fm_get_last_error_message();
  std::int32_t present = 0;
  fm_pivot_worksheet_source_t source = FM_STRUCT_INIT(fm_pivot_worksheet_source_t);
  EXPECT_EQ(fm_workbook_get_pivot_cache_worksheet_source(wb.handle, cache_id, nullptr, &source),
            FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_workbook_get_pivot_cache_worksheet_source(wb.handle, cache_id, &present, nullptr),
            FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_workbook_get_pivot_cache_worksheet_source(wb.handle, cache_id + 100U, &present, &source),
            FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_workbook_set_pivot_cache_worksheet_source(wb.handle, cache_id + 100U, nullptr), FM_STATUS_NOT_FOUND);
}

// ---------------------------------------------------------------------------
// AutoFilter custom conditions
// ---------------------------------------------------------------------------

fm_auto_filter_t CustomFilter(fm_filter_column_t* column) {
  fm_auto_filter_t filter = FM_STRUCT_INIT(fm_auto_filter_t);
  filter.range = fm_cell_rect_t{0, 0, 5, 1};
  filter.columns = column;
  filter.column_count = 1;
  return filter;
}

TEST(FormulonCApiInputRecords, CustomFilterReadsBackOneOrTwoConditions) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  for (const std::uint32_t count : {1U, 2U}) {
    const fm_filter_condition_t conditions[2] = {{FM_COMPARISON_OPERATOR_GREATER_THAN, "3"},
                                                 {FM_COMPARISON_OPERATOR_LESS_THAN_OR_EQUAL, "9"}};
    fm_filter_column_t column{};
    column.column_offset = 1;
    column.show_button = 1;
    column.kind = 2;
    column.custom_and = 1;
    column.conditions = conditions;
    column.condition_count = count;
    const fm_auto_filter_t filter = CustomFilter(&column);
    ASSERT_EQ(fm_sheet_set_auto_filter(wb.handle, 0, &filter), FM_STATUS_OK) << fm_get_last_error_message();

    fm_auto_filter_t out = FM_STRUCT_INIT(fm_auto_filter_t);
    std::int32_t present = 0;
    ASSERT_EQ(fm_sheet_get_auto_filter(wb.handle, 0, &present, &out), FM_STATUS_OK);
    ASSERT_EQ(present, 1);
    ASSERT_EQ(out.column_count, 1U);
    const fm_filter_column_t& read = out.columns[0];
    EXPECT_EQ(read.kind, 2);
    EXPECT_EQ(read.custom_and, 1);
    ASSERT_EQ(read.condition_count, count);
    EXPECT_EQ(read.conditions[0].op, FM_COMPARISON_OPERATOR_GREATER_THAN);
    EXPECT_STREQ(read.conditions[0].value, "3");
    if (count == 2U) {
      EXPECT_EQ(read.conditions[1].op, FM_COMPARISON_OPERATOR_LESS_THAN_OR_EQUAL);
      EXPECT_STREQ(read.conditions[1].value, "9");
    }
  }
}

TEST(FormulonCApiInputRecords, CustomFilterRefusesZeroOrThreeConditionsAndARangeOperator) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  const fm_filter_condition_t conditions[3] = {
      {FM_COMPARISON_OPERATOR_EQUAL, "1"}, {FM_COMPARISON_OPERATOR_EQUAL, "2"}, {FM_COMPARISON_OPERATOR_EQUAL, "3"}};
  fm_filter_column_t column{};
  column.show_button = 1;
  column.kind = 2;
  column.conditions = conditions;
  const fm_auto_filter_t filter = CustomFilter(&column);
  for (const std::uint32_t count : {0U, 3U}) {
    column.condition_count = count;
    EXPECT_EQ(fm_sheet_set_auto_filter(wb.handle, 0, &filter), FM_STATUS_AUTO_FILTER_INVALID) << count;
  }
  for (const std::int32_t op : {static_cast<std::int32_t>(FM_COMPARISON_OPERATOR_BETWEEN),
                                static_cast<std::int32_t>(FM_COMPARISON_OPERATOR_NOT_BETWEEN)}) {
    const fm_filter_condition_t range_operator[1] = {{op, "1"}};
    column.conditions = range_operator;
    column.condition_count = 1;
    EXPECT_EQ(fm_sheet_set_auto_filter(wb.handle, 0, &filter), FM_STATUS_INVALID_ARGUMENT) << op;
  }
  std::int32_t present = 0;
  fm_auto_filter_t out = FM_STRUCT_INIT(fm_auto_filter_t);
  ASSERT_EQ(fm_sheet_get_auto_filter(wb.handle, 0, &present, &out), FM_STATUS_OK);
  EXPECT_EQ(present, 0);
}

TEST(FormulonCApiInputRecords, ColumnsOfOtherKindsReadNoConditions) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  const char* values[] = {"x"};
  fm_filter_column_t column{};
  column.show_button = 1;
  column.kind = 1;
  column.values = values;
  column.value_count = 1;
  // The count is ignored for a kind that holds no conditions.
  column.condition_count = 7;
  const fm_auto_filter_t filter = CustomFilter(&column);
  ASSERT_EQ(fm_sheet_set_auto_filter(wb.handle, 0, &filter), FM_STATUS_OK) << fm_get_last_error_message();
  fm_auto_filter_t out = FM_STRUCT_INIT(fm_auto_filter_t);
  std::int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_auto_filter(wb.handle, 0, &present, &out), FM_STATUS_OK);
  ASSERT_EQ(present, 1);
  ASSERT_EQ(out.column_count, 1U);
  EXPECT_EQ(out.columns[0].kind, 1);
  EXPECT_EQ(out.columns[0].condition_count, 0U);
}

}  // namespace
