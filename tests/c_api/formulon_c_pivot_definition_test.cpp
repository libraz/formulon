// Stable C ABI PivotTable definition read-back and data-field naming tests.

#include "formulon_c_out_index.h"
#include "formulon_c_pivot_test_helpers.h"
#include "pivot/pivot_cache.h"

namespace {

// Name of the sole data field of the scratch pivot, rebuilt with an empty name.
std::string DerivedDataFieldName(const char* profile_id, fm_pivot_aggregation_t aggregation, const char* name) {
  WorkbookGuard wb;
  EXPECT_EQ(fm_workbook_create_for_profile(profile_id, &wb.handle), 0) << fm_get_last_error_message();
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  EXPECT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();
  fm_pivot_data_field_t spec = FM_STRUCT_INIT(fm_pivot_data_field_t);
  spec.name = name;
  spec.field_index = 1;  // "Amount"
  spec.aggregation = aggregation;
  spec.has_show_as_base_field = 0;
  spec.has_show_as_base_item = 0;
  std::size_t idx = 99;
  EXPECT_EQ(fm_sheet_add_pivot_data_field(wb.handle, 0, pivot_index, &spec, &idx), 0) << fm_get_last_error_message();
  fm_pivot_data_field_t read = FM_STRUCT_INIT(fm_pivot_data_field_t);
  EXPECT_EQ(fm_sheet_get_pivot_data_field_at(wb.handle, 0, pivot_index, idx, &read), 0) << fm_get_last_error_message();
  return read.name != nullptr ? read.name : "";
}

/// Sets both grand total bands.
fm_status_t SetGrandTotals(fm_workbook_t* wb, std::size_t pivot_index, std::int32_t rows, std::int32_t columns) {
  fm_pivot_grand_totals_patch_t patch = FM_STRUCT_INIT(fm_pivot_grand_totals_patch_t);
  patch.rows_op = FM_PATCH_OP_SET;
  patch.rows = rows;
  patch.columns_op = FM_PATCH_OP_SET;
  patch.columns = columns;
  return fm_sheet_update_pivot_grand_totals(wb, 0, pivot_index, &patch);
}

}  // namespace

TEST(FormulonCApiPivot, ReadBackRoundTripsWhatTheSettersWrote) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();

  std::uint32_t read_cache = 99;
  ASSERT_EQ(fm_sheet_get_pivot_cache_id(wb.handle, 0, pivot_index, &read_cache), 0);
  EXPECT_EQ(read_cache, cache_id);
  const char* name = nullptr;
  ASSERT_EQ(fm_sheet_get_pivot_name(wb.handle, 0, pivot_index, &name), 0);
  EXPECT_STREQ(name, "PT");

  ASSERT_EQ(fm_sheet_rename_pivot(wb.handle, 0, pivot_index, "Renamed"), 0);
  ASSERT_EQ(fm_sheet_get_pivot_name(wb.handle, 0, pivot_index, &name), 0);
  EXPECT_STREQ(name, "Renamed");

  ASSERT_EQ(SetGrandTotals(wb.handle, pivot_index, 0, 1), 0);
  std::int32_t grand_rows = -1;
  std::int32_t grand_cols = -1;
  ASSERT_EQ(fm_sheet_get_pivot_grand_totals(wb.handle, 0, pivot_index, &grand_rows, &grand_cols), 0);
  EXPECT_EQ(grand_rows, 0);
  EXPECT_EQ(grand_cols, 1);

  // Field 0 is "Region" (row), field 1 is "Amount" (value).
  ASSERT_EQ(SetPivotFieldSort(wb.handle, 0, pivot_index, 0, FM_PIVOT_SORT_ORDER_DESCENDING, "Sum of Amount"), 0);
  // Region holds text, which no date group takes.
  EXPECT_EQ(SetPivotFieldDateGroup(wb.handle, 0, pivot_index, 0, FM_PIVOT_PERIOD_DAYS, FM_PIVOT_CALENDAR_GREGORIAN, 7U,
                                   45000.0, 45100.0),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  fm_pivot_field_info_t region = FM_STRUCT_INIT(fm_pivot_field_info_t);
  ASSERT_EQ(fm_sheet_get_pivot_field_at(wb.handle, 0, pivot_index, 0, &region), 0) << fm_get_last_error_message();
  EXPECT_STREQ(region.source_name, "Region");
  EXPECT_EQ(region.axis, FM_PIVOT_AXIS_ROW);
  EXPECT_EQ(region.subtotal_top, FM_TRISTATE_FALSE);
  EXPECT_EQ(region.cache_field_index, 0);
  EXPECT_EQ(region.sort_order, FM_PIVOT_SORT_ORDER_DESCENDING);
  EXPECT_STREQ(region.sort_by_field, "Sum of Amount");
  EXPECT_EQ(region.date_periods, 0U);
  EXPECT_EQ(region.date_calendar, FM_PIVOT_CALENDAR_GREGORIAN);
  EXPECT_EQ(region.has_group_base_index, 0);
  EXPECT_EQ(region.insert_blank_row, 0);
  EXPECT_EQ(region.fill_down_labels, 0);

  ASSERT_EQ(SetPivotFieldSort(wb.handle, 0, pivot_index, 0, FM_PIVOT_SORT_ORDER_MANUAL, nullptr), 0);
  ASSERT_EQ(fm_sheet_set_pivot_field_date_group(wb.handle, 0, pivot_index, 0, nullptr), 0);
  ASSERT_EQ(fm_sheet_get_pivot_field_at(wb.handle, 0, pivot_index, 0, &region), 0);
  EXPECT_EQ(region.sort_order, FM_PIVOT_SORT_ORDER_MANUAL);
  EXPECT_EQ(region.date_periods, 0U);

  // Items: the scratch pivot added North and South by label, bound to the shared items.
  std::size_t item_count = 99;
  ASSERT_EQ(fm_sheet_get_pivot_field_item_count(wb.handle, 0, pivot_index, 0, &item_count), 0);
  ASSERT_EQ(item_count, 2U);
  ASSERT_EQ(SetPivotFieldItemVisible(wb.handle, 0, pivot_index, 0, 1, 0), 0);
  fm_pivot_item_info_t item = FM_STRUCT_INIT(fm_pivot_item_info_t);
  ASSERT_EQ(fm_sheet_get_pivot_field_item_at(wb.handle, 0, pivot_index, 0, 0, &item), 0);
  EXPECT_STREQ(item.name, "North");
  EXPECT_EQ(item.cache_index, 0);
  EXPECT_EQ(item.visible, 1);
  ASSERT_EQ(fm_sheet_get_pivot_field_item_at(wb.handle, 0, pivot_index, 0, 1, &item), 0);
  EXPECT_STREQ(item.name, "South");
  EXPECT_EQ(item.cache_index, 1);
  EXPECT_EQ(item.visible, 0);
  ASSERT_EQ(fm_sheet_add_pivot_field_item_by_cache_index(wb.handle, 0, pivot_index, 0, 1U, 1, ignored_index()), 0);
  ASSERT_EQ(fm_sheet_get_pivot_field_item_at(wb.handle, 0, pivot_index, 0, 2, &item), 0);
  EXPECT_EQ(item.name, nullptr) << "an item bound only by cache index has no label";
  EXPECT_EQ(item.cache_index, 1);
  EXPECT_EQ(item.visible, 1);

  // Data field written with a base field / item.
  fm_pivot_data_field_t written = FM_STRUCT_INIT(fm_pivot_data_field_t);
  written.name = "Share";
  written.field_index = 1;
  written.aggregation = FM_PIVOT_AGGREGATION_AVERAGE;
  written.number_format = "";
  written.show_as = FM_PIVOT_SHOW_VALUES_AS_DIFFERENCE_FROM;
  written.has_show_as_base_field = 1;
  written.show_as_base_field = 0;
  written.has_show_as_base_item = 1;
  written.show_as_base_item = 1;
  ASSERT_EQ(fm_sheet_set_pivot_data_field(wb.handle, 0, pivot_index, 0, &written), 0) << fm_get_last_error_message();
  fm_pivot_data_field_t data_field = FM_STRUCT_INIT(fm_pivot_data_field_t);
  ASSERT_EQ(fm_sheet_get_pivot_data_field_at(wb.handle, 0, pivot_index, 0, &data_field), 0);
  EXPECT_STREQ(data_field.name, "Share");
  EXPECT_EQ(data_field.field_index, 1U);
  EXPECT_EQ(data_field.aggregation, FM_PIVOT_AGGREGATION_AVERAGE);
  EXPECT_EQ(data_field.show_as, FM_PIVOT_SHOW_VALUES_AS_DIFFERENCE_FROM);
  EXPECT_EQ(data_field.has_show_as_base_field, 1);
  EXPECT_EQ(data_field.show_as_base_field, 0U);
  EXPECT_EQ(data_field.has_show_as_base_item, 1);
  EXPECT_EQ(data_field.show_as_base_item, 1U);

  // Field orders follow the setters.
  fm_pivot_field_spec_t extra = FM_STRUCT_INIT(fm_pivot_field_spec_t);
  extra.source_name = "Amount";
  extra.axis = FM_PIVOT_AXIS_COL;
  extra.subtotal_top = FM_TRISTATE_TRUE;
  std::size_t extra_index = 99;
  ASSERT_EQ(fm_sheet_add_pivot_field(wb.handle, 0, pivot_index, &extra, &extra_index), 0);
  std::uint32_t order[8] = {};
  std::size_t order_count = 99;
  ASSERT_EQ(fm_sheet_get_pivot_column_field_order(wb.handle, 0, pivot_index, order, 8U, &order_count), 0);
  ASSERT_EQ(order_count, 1U);
  EXPECT_EQ(order[0], static_cast<std::uint32_t>(extra_index));
  const std::uint32_t swapped[] = {static_cast<std::uint32_t>(extra_index), 0U};
  ASSERT_EQ(fm_sheet_set_pivot_row_field_order(wb.handle, 0, pivot_index, swapped, 2U), 0);
  ASSERT_EQ(fm_sheet_get_pivot_row_field_order(wb.handle, 0, pivot_index, order, 8U, &order_count), 0);
  ASSERT_EQ(order_count, 2U);
  EXPECT_EQ(order[0], static_cast<std::uint32_t>(extra_index));
  EXPECT_EQ(order[1], 0U);
  ASSERT_EQ(fm_sheet_get_pivot_column_field_order(wb.handle, 0, pivot_index, order, 8U, &order_count), 0);
  EXPECT_EQ(order_count, 0U);
  fm_pivot_field_info_t moved = FM_STRUCT_INIT(fm_pivot_field_info_t);
  ASSERT_EQ(fm_sheet_get_pivot_field_at(wb.handle, 0, pivot_index, extra_index, &moved), 0);
  EXPECT_EQ(moved.axis, FM_PIVOT_AXIS_ROW);

  // A field left out of an axis order sits on no axis.
  const std::uint32_t only_extra[] = {static_cast<std::uint32_t>(extra_index)};
  ASSERT_EQ(fm_sheet_set_pivot_row_field_order(wb.handle, 0, pivot_index, only_extra, 1U), 0);
  fm_pivot_field_info_t dropped = FM_STRUCT_INIT(fm_pivot_field_info_t);
  ASSERT_EQ(fm_sheet_get_pivot_field_at(wb.handle, 0, pivot_index, 0, &dropped), 0);
  EXPECT_EQ(dropped.axis, FM_PIVOT_AXIS_ABSENT);
}

TEST(FormulonCApiPivot, ReadBackSurvivesSaveAndReload) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();
  ASSERT_EQ(SetGrandTotals(wb.handle, pivot_index, 1, 0), 0);
  ASSERT_EQ(SetPivotFieldItemVisible(wb.handle, 0, pivot_index, 0, 1, 0), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0) << fm_get_last_error_message();
  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0) << fm_get_last_error_message();

  std::uint32_t read_cache = 99;
  ASSERT_EQ(fm_sheet_get_pivot_cache_id(reloaded.handle, 0, 0, &read_cache), 0) << fm_get_last_error_message();
  const char* name = nullptr;
  ASSERT_EQ(fm_sheet_get_pivot_name(reloaded.handle, 0, 0, &name), 0);
  EXPECT_STREQ(name, "PT");
  std::int32_t grand_rows = -1;
  std::int32_t grand_cols = -1;
  ASSERT_EQ(fm_sheet_get_pivot_grand_totals(reloaded.handle, 0, 0, &grand_rows, &grand_cols), 0);
  EXPECT_EQ(grand_rows, 1);
  EXPECT_EQ(grand_cols, 0);

  fm_pivot_field_info_t region = FM_STRUCT_INIT(fm_pivot_field_info_t);
  ASSERT_EQ(fm_sheet_get_pivot_field_at(reloaded.handle, 0, 0, 0, &region), 0);
  EXPECT_STREQ(region.source_name, "Region");
  EXPECT_EQ(region.axis, FM_PIVOT_AXIS_ROW);
  EXPECT_EQ(region.cache_field_index, 0);
  fm_pivot_item_info_t item = FM_STRUCT_INIT(fm_pivot_item_info_t);
  ASSERT_EQ(fm_sheet_get_pivot_field_item_at(reloaded.handle, 0, 0, 0, 1, &item), 0);
  EXPECT_EQ(item.cache_index, 1);
  EXPECT_EQ(item.visible, 0);

  fm_pivot_data_field_t data_field = FM_STRUCT_INIT(fm_pivot_data_field_t);
  ASSERT_EQ(fm_sheet_get_pivot_data_field_at(reloaded.handle, 0, 0, 0, &data_field), 0);
  EXPECT_STREQ(data_field.name, "Sum of Amount");
  EXPECT_EQ(data_field.aggregation, FM_PIVOT_AGGREGATION_SUM);
  EXPECT_EQ(data_field.field_index, 1U);

  std::uint32_t order[8] = {};
  std::size_t order_count = 99;
  ASSERT_EQ(fm_sheet_get_pivot_row_field_order(reloaded.handle, 0, 0, order, 8U, &order_count), 0);
  ASSERT_EQ(order_count, 1U);
  EXPECT_EQ(order[0], 0U);
}

TEST(FormulonCApiPivot, ReadBackRejectsBadArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();

  std::uint32_t out_cache = 7;
  EXPECT_EQ(fm_sheet_get_pivot_cache_id(nullptr, 0, pivot_index, &out_cache),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_sheet_get_pivot_cache_id(wb.handle, 0, pivot_index, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_sheet_get_pivot_cache_id(wb.handle, 0, pivot_index + 1U, &out_cache),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kIndexOutOfRange));
  EXPECT_EQ(out_cache, 7U);

  fm_pivot_field_info_t field = FM_STRUCT_INIT(fm_pivot_field_info_t);
  EXPECT_EQ(fm_sheet_get_pivot_field_at(wb.handle, 0, pivot_index, 99, &field),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kIndexOutOfRange));
  EXPECT_EQ(fm_sheet_get_pivot_field_at(wb.handle, 0, pivot_index, 0, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  fm_pivot_item_info_t item = FM_STRUCT_INIT(fm_pivot_item_info_t);
  EXPECT_EQ(fm_sheet_get_pivot_field_item_at(wb.handle, 0, pivot_index, 0, 99, &item),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kIndexOutOfRange));
  fm_pivot_data_field_t data_field = FM_STRUCT_INIT(fm_pivot_data_field_t);
  EXPECT_EQ(fm_sheet_get_pivot_data_field_at(wb.handle, 0, pivot_index, 99, &data_field),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kIndexOutOfRange));
  std::uint32_t order[8] = {};
  std::size_t count = 0;
  EXPECT_EQ(fm_sheet_get_pivot_row_field_order(wb.handle, 0, pivot_index, nullptr, 1U, &count),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer))
      << "a NULL buffer with a non-zero capacity";
  EXPECT_EQ(fm_sheet_get_pivot_row_field_order(wb.handle, 0, pivot_index, order, 8U, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_sheet_get_pivot_row_field_order(wb.handle, 0, pivot_index, nullptr, 0U, &count), 0)
      << "capacity 0 sizes the buffer";
  EXPECT_EQ(fm_sheet_get_pivot_column_field_order(wb.handle, 9, pivot_index, order, 8U, &count),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kSheetIndexOutOfRange));
}

TEST(FormulonCApiPivot, EmptyDataFieldNameIsDerivedFromTheProfile) {
  // The scratch pivot already holds "Sum of Amount", so Excel numbers the new one.
  EXPECT_EQ(DerivedDataFieldName("win-365-en_US", FM_PIVOT_AGGREGATION_SUM, nullptr), "Sum of Amount2");
  EXPECT_EQ(DerivedDataFieldName("win-365-en_US", FM_PIVOT_AGGREGATION_AVERAGE, ""), "Average of Amount");
  EXPECT_EQ(DerivedDataFieldName("win-365-ja_JP", FM_PIVOT_AGGREGATION_SUM, nullptr), "合計 / Amount");
  EXPECT_EQ(DerivedDataFieldName("win-365-ja_JP", FM_PIVOT_AGGREGATION_COUNT, ""), "個数 / Amount");
  // An explicit name is kept as given.
  EXPECT_EQ(DerivedDataFieldName("win-365-ja_JP", FM_PIVOT_AGGREGATION_SUM, "Total"), "Total");
}

TEST(FormulonCApiPivot, DerivedDataFieldNameUsesTheFieldDisplayNameAndSetPath) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_for_profile("win-365-en_US", &wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();

  // A renamed source field is named by its caption.
  fm_pivot_field_spec_t caption = FM_STRUCT_INIT(fm_pivot_field_spec_t);
  caption.source_name = "Amount";
  caption.custom_name = "Sales";
  caption.axis = FM_PIVOT_AXIS_VALUE;
  caption.subtotal_top = FM_TRISTATE_TRUE;
  std::size_t caption_index = 99;
  ASSERT_EQ(fm_sheet_add_pivot_field(wb.handle, 0, pivot_index, &caption, &caption_index), 0);

  fm_pivot_data_field_t spec = FM_STRUCT_INIT(fm_pivot_data_field_t);
  spec.field_index = static_cast<std::uint32_t>(caption_index);
  spec.aggregation = FM_PIVOT_AGGREGATION_MAX;
  spec.has_show_as_base_field = 0;
  spec.has_show_as_base_item = 0;
  ASSERT_EQ(fm_sheet_set_pivot_data_field(wb.handle, 0, pivot_index, 0, &spec), 0) << fm_get_last_error_message();
  fm_pivot_data_field_t read = FM_STRUCT_INIT(fm_pivot_data_field_t);
  ASSERT_EQ(fm_sheet_get_pivot_data_field_at(wb.handle, 0, pivot_index, 0, &read), 0);
  EXPECT_STREQ(read.name, "Max of Sales");

  // A derived name needs a field to name it after.
  spec.field_index = 99;
  EXPECT_EQ(fm_sheet_set_pivot_data_field(wb.handle, 0, pivot_index, 0, &spec),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kIndexOutOfRange));
  std::size_t added = 0;
  EXPECT_EQ(fm_sheet_add_pivot_data_field(wb.handle, 0, pivot_index, &spec, &added),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kIndexOutOfRange));
  ASSERT_EQ(fm_sheet_get_pivot_data_field_at(wb.handle, 0, pivot_index, 0, &read), 0);
  EXPECT_STREQ(read.name, "Max of Sales");
}

namespace {

// A cache [Region, Product, Sales, Qty] over three records, and a pivot on it with no fields yet.
void BuildWideCachePivot(fm_workbook_t* wb, std::uint32_t* cache_id, std::size_t* pivot_index) {
  ASSERT_EQ(fm_workbook_add_pivot_cache(wb, 0, cache_id), 0);
  ASSERT_EQ(SetPivotCacheWorksheetSource(wb, *cache_id, 1, "A1:D4", "Sheet1", nullptr), 0);
  for (const char* name : {"Region", "Product", "Sales", "Qty"}) {
    std::size_t idx = 0;
    ASSERT_EQ(fm_workbook_add_pivot_cache_field(wb, *cache_id, name, &idx), 0);
  }
  ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_text(wb, *cache_id, 0, "North"), 0);
  ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_text(wb, *cache_id, 0, "South"), 0);
  ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_text(wb, *cache_id, 1, "Pen"), 0);
  const double rows[][4] = {{0, 0, 100, 1}, {1, 0, 200, 2}, {0, 0, 300, 3}};
  for (const auto& row : rows) {
    std::size_t rec = 0;
    ASSERT_EQ(fm_workbook_add_pivot_cache_record(wb, *cache_id, &rec), 0);
    for (std::size_t col = 0; col < 4; ++col) {
      ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb, *cache_id, rec, col, row[col]), 0);
    }
  }
  ASSERT_EQ(AddPivot(wb, 0, "PT", *cache_id, 0, 6, pivot_index), 0);
}

fm_pivot_field_spec_t FieldSpec(const char* name, fm_pivot_axis_t axis) {
  fm_pivot_field_spec_t spec = FM_STRUCT_INIT(fm_pivot_field_spec_t);
  spec.source_name = name;
  spec.axis = axis;
  spec.subtotal_top = FM_TRISTATE_TRUE;
  return spec;
}

}  // namespace

TEST(FormulonCApiPivot, FieldAddBindsToTheCacheColumnByName) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot = 0;
  BuildWideCachePivot(wb.handle, &cache_id, &pivot);

  std::size_t region = 99;
  std::size_t sales = 99;
  const fm_pivot_field_spec_t region_spec = FieldSpec("Region", FM_PIVOT_AXIS_ROW);
  const fm_pivot_field_spec_t sales_spec = FieldSpec("Sales", FM_PIVOT_AXIS_VALUE);
  ASSERT_EQ(fm_sheet_add_pivot_field(wb.handle, 0, pivot, &region_spec, &region), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_add_pivot_field(wb.handle, 0, pivot, &sales_spec, &sales), 0) << fm_get_last_error_message();
  // The index is the cache column's position, not the order the fields were added in.
  EXPECT_EQ(region, 0U);
  EXPECT_EQ(sales, 2U);
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_pivot_field_count(wb.handle, 0, pivot, &count), 0);
  EXPECT_EQ(count, 4U);

  fm_pivot_field_info_t product = FM_STRUCT_INIT(fm_pivot_field_info_t);
  ASSERT_EQ(fm_sheet_get_pivot_field_at(wb.handle, 0, pivot, 1, &product), 0);
  EXPECT_STREQ(product.source_name, "Product");
  EXPECT_EQ(product.axis, FM_PIVOT_AXIS_ABSENT);

  fm_pivot_data_field_t df = FM_STRUCT_INIT(fm_pivot_data_field_t);
  df.name = "Total";
  df.field_index = static_cast<std::uint32_t>(sales);
  df.aggregation = FM_PIVOT_AGGREGATION_SUM;
  df.has_show_as_base_field = 0;
  df.has_show_as_base_item = 0;
  std::size_t df_index = 0;
  ASSERT_EQ(fm_sheet_add_pivot_data_field(wb.handle, 0, pivot, &df, &df_index), 0);
  PivotCellsGuard projected;
  ASSERT_EQ(fm_sheet_get_pivot_layout(wb.handle, 0, pivot, &projected.handle), 0) << fm_get_last_error_message();
  double grand_total = 0.0;
  for (const fm_pivot_cell_info_t& cell : CollectCells(projected.handle)) {
    if (cell.kind == FM_PIVOT_CELL_KIND_GRAND_TOTAL && cell.value.kind == FM_VALUE_KIND_NUMBER) {
      grand_total = cell.value.u.number;
    }
  }
  EXPECT_EQ(grand_total, 600.0);

  // Adding a field again updates it in place.
  const fm_pivot_field_spec_t renamed = [] {
    fm_pivot_field_spec_t spec = FieldSpec("Region", FM_PIVOT_AXIS_ROW);
    spec.custom_name = "Area";
    return spec;
  }();
  std::size_t again = 99;
  ASSERT_EQ(fm_sheet_add_pivot_field(wb.handle, 0, pivot, &renamed, &again), 0);
  EXPECT_EQ(again, 0U);
  ASSERT_EQ(fm_sheet_get_pivot_field_count(wb.handle, 0, pivot, &count), 0);
  EXPECT_EQ(count, 4U);
  std::uint32_t order[8] = {};
  std::size_t order_count = 0;
  ASSERT_EQ(fm_sheet_get_pivot_row_field_order(wb.handle, 0, pivot, order, 8U, &order_count), 0);
  EXPECT_EQ(order_count, 1U);

  // A name the cache lacks is appended after the materialised fields.
  const fm_pivot_field_spec_t extra = FieldSpec("Notes", FM_PIVOT_AXIS_ROW);
  std::size_t appended = 0;
  ASSERT_EQ(fm_sheet_add_pivot_field(wb.handle, 0, pivot, &extra, &appended), 0);
  EXPECT_EQ(appended, 4U);
}

TEST(FormulonCApiPivot, FieldAxisNoneTakesTheFieldOffTheReport) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot = 0;
  BuildWideCachePivot(wb.handle, &cache_id, &pivot);
  std::size_t idx = 0;
  const fm_pivot_field_spec_t region = FieldSpec("Region", FM_PIVOT_AXIS_ROW);
  const fm_pivot_field_spec_t sales = FieldSpec("Sales", FM_PIVOT_AXIS_VALUE);
  const fm_pivot_field_spec_t qty = FieldSpec("Qty", FM_PIVOT_AXIS_VALUE);
  ASSERT_EQ(fm_sheet_add_pivot_field(wb.handle, 0, pivot, &region, &idx), 0);
  ASSERT_EQ(fm_sheet_add_pivot_field(wb.handle, 0, pivot, &sales, &idx), 0);
  ASSERT_EQ(fm_sheet_add_pivot_field(wb.handle, 0, pivot, &qty, &idx), 0);
  for (const std::uint32_t field : {2U, 3U}) {
    fm_pivot_data_field_t df = FM_STRUCT_INIT(fm_pivot_data_field_t);
    df.field_index = field;
    df.aggregation = FM_PIVOT_AGGREGATION_SUM;
    df.has_show_as_base_field = 0;
    df.has_show_as_base_item = 0;
    ASSERT_EQ(fm_sheet_add_pivot_data_field(wb.handle, 0, pivot, &df, &idx), 0) << fm_get_last_error_message();
  }
  fm_pivot_filter_t filter = FM_STRUCT_INIT(fm_pivot_filter_t);
  filter.field_index = 3;
  filter.type = FM_PIVOT_FILTER_KIND_VALUE_GREATER_THAN;
  filter.value_kind = FM_PIVOT_FILTER_OPERAND_KIND_DOUBLE;
  filter.value_double = 0.0;
  filter.data_field_index = 1;
  std::size_t filter_index = 99;
  ASSERT_EQ(fm_sheet_add_pivot_filter(wb.handle, 0, pivot, &filter, &filter_index), 0) << fm_get_last_error_message();
  EXPECT_EQ(filter_index, 0U);

  // Taking Sales off removes its data field; Qty's data field and filter move down one.
  ASSERT_EQ(fm_sheet_set_pivot_field_axis(wb.handle, 0, pivot, 2, FM_PIVOT_AXIS_ABSENT), 0)
      << fm_get_last_error_message();
  fm_pivot_field_info_t info = FM_STRUCT_INIT(fm_pivot_field_info_t);
  ASSERT_EQ(fm_sheet_get_pivot_field_at(wb.handle, 0, pivot, 2, &info), 0);
  EXPECT_EQ(info.axis, FM_PIVOT_AXIS_ABSENT);
  std::size_t data_fields = 0;
  ASSERT_EQ(fm_sheet_get_pivot_data_field_count(wb.handle, 0, pivot, &data_fields), 0);
  ASSERT_EQ(data_fields, 1U);
  fm_pivot_data_field_t remaining = FM_STRUCT_INIT(fm_pivot_data_field_t);
  ASSERT_EQ(fm_sheet_get_pivot_data_field_at(wb.handle, 0, pivot, 0, &remaining), 0);
  EXPECT_EQ(remaining.field_index, 3U);
  fm_pivot_filter_t kept = FM_STRUCT_INIT(fm_pivot_filter_t);
  ASSERT_EQ(fm_sheet_get_pivot_filter_at(wb.handle, 0, pivot, 0, &kept), 0);
  EXPECT_EQ(kept.data_field_index, 0U);

  // A row field leaves its axis order too; Qty then takes its filter along.
  ASSERT_EQ(fm_sheet_set_pivot_field_axis(wb.handle, 0, pivot, 0, FM_PIVOT_AXIS_ABSENT), 0);
  std::uint32_t order[8] = {};
  std::size_t order_count = 9;
  ASSERT_EQ(fm_sheet_get_pivot_row_field_order(wb.handle, 0, pivot, order, 8U, &order_count), 0);
  EXPECT_EQ(order_count, 0U);
  ASSERT_EQ(fm_sheet_set_pivot_field_axis(wb.handle, 0, pivot, 3, FM_PIVOT_AXIS_ABSENT), 0);
  std::size_t filters = 9;
  ASSERT_EQ(fm_sheet_get_pivot_filter_count(wb.handle, 0, pivot, &filters), 0);
  EXPECT_EQ(filters, 0U);

  EXPECT_NE(fm_sheet_set_pivot_field_axis(wb.handle, 0, pivot, 0, -2), 0);
  EXPECT_NE(fm_sheet_set_pivot_field_axis(wb.handle, 0, pivot, 0, 4), 0);
}

TEST(FormulonCApiPivot, SaveDropsPivotCachesNoPivotUses) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot = 0;
  BuildWideCachePivot(wb.handle, &cache_id, &pivot);
  const auto saved_with_cache = [&] {
    BufferGuard saved;
    EXPECT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0) << fm_get_last_error_message();
    return std::string(reinterpret_cast<const char*>(saved.data), saved.len).find("pivotCacheDefinition1.xml") !=
           std::string::npos;
  };
  EXPECT_TRUE(saved_with_cache());

  // Removing the pivot keeps the cache in the model, so a recreated pivot can reuse its id ...
  ASSERT_EQ(fm_sheet_remove_pivot(wb.handle, 0, pivot), 0);
  std::size_t caches = 0;
  ASSERT_EQ(fm_workbook_get_pivot_cache_count(wb.handle, &caches), 0);
  EXPECT_EQ(caches, 1U);
  // ... but nothing draws from it, so a save leaves it out.
  EXPECT_FALSE(saved_with_cache());
  std::size_t again = 0;
  ASSERT_EQ(AddPivot(wb.handle, 0, "PT2", cache_id, 0, 6, &again), 0);
  EXPECT_TRUE(saved_with_cache());
}

namespace {

// Rebuilds `archive` with the pivot table part replaced by `pivot_xml`.
std::vector<std::uint8_t> WithPivotTablePart(const std::vector<std::uint8_t>& archive, const std::string& pivot_xml) {
  constexpr std::string_view kPivotPart = "xl/pivotTables/pivotTable1.xml";
  mz_zip_archive reader{};
  EXPECT_NE(mz_zip_reader_init_mem(&reader, archive.data(), archive.size(), 0), MZ_FALSE);
  const mz_uint count = mz_zip_reader_get_num_files(&reader);
  std::vector<std::string> names;
  std::vector<std::string> bodies;
  for (mz_uint i = 0; i < count; ++i) {
    char name[256] = {};
    mz_zip_reader_get_filename(&reader, i, name, sizeof(name));
    if (mz_zip_reader_is_file_a_directory(&reader, i) != MZ_FALSE) {
      continue;
    }
    names.emplace_back(name);
    if (names.back() == kPivotPart) {
      bodies.push_back(pivot_xml);
      continue;
    }
    std::size_t size = 0;
    void* data = mz_zip_reader_extract_to_heap(&reader, i, &size, 0);
    EXPECT_NE(data, nullptr);
    bodies.emplace_back(static_cast<const char*>(data), size);
    mz_free(data);
  }
  mz_zip_reader_end(&reader);
  std::vector<PartFile> parts;
  for (std::size_t i = 0; i < names.size(); ++i) {
    parts.push_back({names[i].c_str(), bodies[i]});
  }
  return BuildZip(parts);
}

// Saves `wb`, swaps everything after `</dataFields>` in its pivot part for
// `tail`, and loads the result into `out`.
void LoadWithPivotTail(fm_workbook_t* wb, const std::string& tail, WorkbookGuard* out) {
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb, &saved.data, &saved.len), 0) << fm_get_last_error_message();
  const std::vector<std::uint8_t> archive(saved.data, saved.data + saved.len);
  std::string xml = ExtractZipEntry(archive, "xl/pivotTables/pivotTable1.xml");
  const std::size_t cut = xml.find("</dataFields>");
  ASSERT_NE(cut, std::string::npos) << xml;
  xml.resize(cut + std::string_view("</dataFields>").size());
  xml.append(tail).append("</pivotTableDefinition>");
  const std::vector<std::uint8_t> patched = WithPivotTablePart(archive, xml);
  ASSERT_EQ(fm_workbook_load(patched.data(), patched.size(), &out->handle), 0) << fm_get_last_error_message();
}

/// The style record together with the `out_present` flag of the getter.
struct PivotStyleRead : fm_pivot_style_options_t {
  std::int32_t present = 0;
};

PivotStyleRead ReadStyleOptions(const fm_workbook_t* wb) {
  PivotStyleRead read;
  static_cast<fm_pivot_style_options_t&>(read) = FM_STRUCT_INIT(fm_pivot_style_options_t);
  EXPECT_EQ(fm_sheet_get_pivot_style_options(wb, 0, 0, &read.present, &read), 0) << fm_get_last_error_message();
  return read;
}

}  // namespace

TEST(FormulonCApiPivot, CreatedPivotCarriesTheDefaultTableStyle) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();

  const auto style = ReadStyleOptions(wb.handle);
  EXPECT_EQ(style.present, 1);
  EXPECT_STREQ(style.style_name, "PivotStyleLight16");
  EXPECT_EQ(style.show_row_headers, 1);
  EXPECT_EQ(style.show_column_headers, 1);
  EXPECT_EQ(style.show_row_stripes, 0);
  EXPECT_EQ(style.show_column_stripes, 0);
  EXPECT_EQ(style.show_last_column, 1);

  const std::string xml = SavedPivotXml(wb.handle);
  EXPECT_NE(xml.find("<pivotTableStyleInfo name=\"PivotStyleLight16\" showRowHeaders=\"1\" showColHeaders=\"1\" "
                     "showRowStripes=\"0\" showColStripes=\"0\" showLastColumn=\"1\"/>"),
            std::string::npos)
      << xml;
  EXPECT_LT(xml.find("</dataFields>"), xml.find("<pivotTableStyleInfo"));
  EXPECT_LT(xml.find("<pivotTableStyleInfo"), xml.find("<extLst"));
}

TEST(FormulonCApiPivot, StyleOptionsRoundTripsThroughSetSaveAndLoad) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();

  ASSERT_EQ(UpdatePivotStyleOptions(wb.handle, 0, pivot_index, "PivotStyleMedium9", 0, 1, 1, 1, 0), 0)
      << fm_get_last_error_message();
  auto style = ReadStyleOptions(wb.handle);
  EXPECT_STREQ(style.style_name, "PivotStyleMedium9");
  EXPECT_EQ(style.show_row_headers, 0);
  EXPECT_EQ(style.show_column_headers, 1);
  EXPECT_EQ(style.show_row_stripes, 1);
  EXPECT_EQ(style.show_column_stripes, 1);
  EXPECT_EQ(style.show_last_column, 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0) << fm_get_last_error_message();
  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0) << fm_get_last_error_message();
  style = ReadStyleOptions(reloaded.handle);
  EXPECT_EQ(style.present, 1);
  EXPECT_STREQ(style.style_name, "PivotStyleMedium9");
  EXPECT_EQ(style.show_row_headers, 0);
  EXPECT_EQ(style.show_column_headers, 1);
  EXPECT_EQ(style.show_row_stripes, 1);
  EXPECT_EQ(style.show_column_stripes, 1);
  EXPECT_EQ(style.show_last_column, 0);

  // A NULL name keeps the element but drops the style name.
  ASSERT_EQ(UpdatePivotStyleOptions(wb.handle, 0, pivot_index, nullptr, 1, 1, 0, 0, 0), 0);
  style = ReadStyleOptions(wb.handle);
  EXPECT_EQ(style.present, 1);
  EXPECT_EQ(style.style_name, nullptr);
  const std::string xml = SavedPivotXml(wb.handle);
  EXPECT_NE(xml.find("<pivotTableStyleInfo showRowHeaders=\"1\""), std::string::npos) << xml;
}

TEST(FormulonCApiPivot, StyleOptionsRejectsBadArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();

  fm_pivot_style_options_t style = FM_STRUCT_INIT(fm_pivot_style_options_t);
  std::int32_t present = 0;
  EXPECT_EQ(fm_sheet_get_pivot_style_options(nullptr, 0, pivot_index, &present, &style),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_sheet_get_pivot_style_options(wb.handle, 0, pivot_index, &present, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_sheet_get_pivot_style_options(wb.handle, 0, pivot_index, nullptr, &style),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_sheet_get_pivot_style_options(wb.handle, 0, pivot_index + 1U, &present, &style),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kIndexOutOfRange));
  EXPECT_NE(UpdatePivotStyleOptions(wb.handle, 0, pivot_index + 1U, "X", 1, 1, 0, 0, 0), 0);
  EXPECT_NE(UpdatePivotStyleOptions(nullptr, 0, pivot_index, "X", 1, 1, 0, 0, 0), 0);
}

TEST(FormulonCApiPivot, LoadedStyleInfoKeepsSchemaOrderBetweenFormatsAndExtLst) {
  WorkbookGuard source;
  ASSERT_EQ(fm_workbook_create_default(&source.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(source.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();

  WorkbookGuard loaded;
  LoadWithPivotTail(source.handle,
                    "<formats count=\"1\"><format action=\"formatting\"><pivotArea outline=\"0\" "
                    "collapsedLevelsAreSubtotals=\"1\" fieldPosition=\"0\"/></format></formats>"
                    "<pivotTableStyleInfo name=\"PivotStyleDark3\" showRowHeaders=\"0\" showColHeaders=\"1\" "
                    "showRowStripes=\"1\" showColStripes=\"0\" showLastColumn=\"0\"/>"
                    "<extLst><ext uri=\"{FF0F0A60-0000-4000-8000-000000000001}\" "
                    "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"/></extLst>",
                    &loaded);
  const auto style = ReadStyleOptions(loaded.handle);
  EXPECT_EQ(style.present, 1);
  EXPECT_STREQ(style.style_name, "PivotStyleDark3");
  EXPECT_EQ(style.show_row_headers, 0);
  EXPECT_EQ(style.show_column_headers, 1);
  EXPECT_EQ(style.show_row_stripes, 1);
  EXPECT_EQ(style.show_column_stripes, 0);
  EXPECT_EQ(style.show_last_column, 0);

  const std::string xml = SavedPivotXml(loaded.handle);
  const std::size_t data_fields = xml.find("</dataFields>");
  const std::size_t formats = xml.find("<formats");
  const std::size_t style_info = xml.find("<pivotTableStyleInfo");
  const std::size_t ext_lst = xml.find("<extLst");
  ASSERT_NE(data_fields, std::string::npos) << xml;
  ASSERT_NE(formats, std::string::npos) << xml;
  ASSERT_NE(style_info, std::string::npos) << xml;
  ASSERT_NE(ext_lst, std::string::npos) << xml;
  EXPECT_LT(data_fields, formats);
  EXPECT_LT(formats, style_info);
  EXPECT_LT(style_info, ext_lst);
  EXPECT_EQ(CountOccurrences(xml, "<pivotTableStyleInfo"), 1U) << xml;
  EXPECT_EQ(CountOccurrences(xml, "<extLst"), 1U) << xml;
}

TEST(FormulonCApiPivot, LoadedPivotWithoutStyleInfoStaysWithoutIt) {
  WorkbookGuard source;
  ASSERT_EQ(fm_workbook_create_default(&source.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(source.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();

  WorkbookGuard loaded;
  LoadWithPivotTail(source.handle, "", &loaded);
  const auto style = ReadStyleOptions(loaded.handle);
  EXPECT_EQ(style.present, 0);
  EXPECT_EQ(style.style_name, nullptr);
  EXPECT_EQ(style.show_row_headers, 0);
  EXPECT_EQ(style.show_last_column, 0);
  EXPECT_EQ(SavedPivotXml(loaded.handle).find("<pivotTableStyleInfo"), std::string::npos);
}

TEST(FormulonCApiPivot, SettingAWorksheetSourceDropsTheCapturedSourceKind) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();
  formulon::pivot::PivotCache* cache = nullptr;
  for (const auto& candidate : wb.handle->workbook().mutable_pivot_caches()) {
    if (candidate->cache_id() == cache_id) {
      cache = candidate.get();
    }
  }
  ASSERT_NE(cache, nullptr);
  // As read from a consolidation-sourced cache.
  cache->mutable_cache_source_attrs() = {{"type", "consolidation"}};
  cache->mutable_cache_source_children_xml() = "<consolidation><rangeSets count=\"0\"/></consolidation>";

  ASSERT_EQ(SetPivotCacheWorksheetSource(wb.handle, cache_id, 1, "A1:B5", "Sheet1", nullptr), 0)
      << fm_get_last_error_message();
  EXPECT_TRUE(cache->cache_source_attrs().empty());
  EXPECT_TRUE(cache->cache_source_children_xml().empty());

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0) << fm_get_last_error_message();
  const std::string definition = ExtractZipEntry(std::vector<std::uint8_t>(saved.data, saved.data + saved.len),
                                                 "xl/pivotCache/pivotCacheDefinition1.xml");
  EXPECT_NE(definition.find("<cacheSource type=\"worksheet\"><worksheetSource"), std::string::npos) << definition;
  EXPECT_EQ(definition.find("consolidation"), std::string::npos) << definition;
}
