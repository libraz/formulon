// Stable C ABI PivotTable definition read-back and data-field naming tests.

#include "formulon_c_pivot_test_helpers.h"
#include "pivot/pivot_cache.h"

namespace {

// Name of the sole data field of the scratch pivot, rebuilt with an empty name.
std::string DerivedDataFieldName(const char* profile_id, fm_pivot_aggregation_t aggregation, const char* name) {
  WorkbookGuard wb;
  EXPECT_EQ(fm_workbook_create_for_profile(profile_id, &wb.handle), 0) << fm_last_error_message();
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  EXPECT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();
  fm_pivot_data_field_spec_t spec{};
  spec.name = name;
  spec.field_index = 1;  // "Amount"
  spec.aggregation = aggregation;
  spec.show_as_base_field = -1;
  spec.show_as_base_item = -1;
  std::size_t idx = 99;
  EXPECT_EQ(fm_workbook_pivot_data_field_add(wb.handle, 0, pivot_idx, &spec, &idx), 0) << fm_last_error_message();
  fm_pivot_data_field_spec_t read{};
  EXPECT_EQ(fm_workbook_pivot_data_field_at(wb.handle, 0, pivot_idx, idx, &read), 0) << fm_last_error_message();
  return read.name != nullptr ? read.name : "";
}

}  // namespace

TEST(FormulonCApiPivot, ReadBackRoundTripsWhatTheSettersWrote) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  std::uint32_t read_cache = 99;
  ASSERT_EQ(fm_workbook_pivot_get_cache_id(wb.handle, 0, pivot_idx, &read_cache), 0);
  EXPECT_EQ(read_cache, cache_id);
  const char* name = nullptr;
  ASSERT_EQ(fm_workbook_pivot_get_name(wb.handle, 0, pivot_idx, &name), 0);
  EXPECT_STREQ(name, "PT");

  ASSERT_EQ(fm_workbook_pivot_set_name(wb.handle, 0, pivot_idx, "Renamed"), 0);
  ASSERT_EQ(fm_workbook_pivot_get_name(wb.handle, 0, pivot_idx, &name), 0);
  EXPECT_STREQ(name, "Renamed");

  ASSERT_EQ(fm_workbook_pivot_set_grand_totals(wb.handle, 0, pivot_idx, 0, 1), 0);
  std::int32_t grand_rows = -1;
  std::int32_t grand_cols = -1;
  ASSERT_EQ(fm_workbook_pivot_get_grand_totals(wb.handle, 0, pivot_idx, &grand_rows, &grand_cols), 0);
  EXPECT_EQ(grand_rows, 0);
  EXPECT_EQ(grand_cols, 1);

  // Field 0 is "Region" (row), field 1 is "Amount" (value).
  ASSERT_EQ(fm_workbook_pivot_field_set_sort(wb.handle, 0, pivot_idx, 0, FM_PIVOT_SORT_DESCENDING, "Sum of Amount"), 0);
  ASSERT_EQ(fm_workbook_pivot_field_set_date_group(wb.handle, 0, pivot_idx, 0, FM_PIVOT_DATE_DAYS,
                                                   FM_PIVOT_CALENDAR_GREGORIAN, -1, -1, 7U, 45000.0, 45100.0),
            0)
      << fm_last_error_message();
  fm_pivot_field_info_t region{};
  ASSERT_EQ(fm_workbook_pivot_field_at(wb.handle, 0, pivot_idx, 0, &region), 0) << fm_last_error_message();
  EXPECT_STREQ(region.source_name, "Region");
  EXPECT_EQ(region.axis, FM_PIVOT_AXIS_ROW);
  EXPECT_EQ(region.subtotal_top, 0);
  EXPECT_EQ(region.cache_field_index, 0);
  EXPECT_EQ(region.sort_order, FM_PIVOT_SORT_DESCENDING);
  EXPECT_STREQ(region.sort_by_field, "Sum of Amount");
  EXPECT_EQ(region.has_date_group, 1);
  EXPECT_EQ(region.date_granularity, FM_PIVOT_DATE_DAYS);
  EXPECT_EQ(region.date_calendar, FM_PIVOT_CALENDAR_GREGORIAN);
  EXPECT_EQ(region.date_start_year, -1);
  EXPECT_EQ(region.date_interval_days, 7U);
  EXPECT_EQ(region.date_start_serial, 45000.0);
  EXPECT_EQ(region.date_end_serial, 45100.0);

  ASSERT_EQ(fm_workbook_pivot_field_set_sort(wb.handle, 0, pivot_idx, 0, FM_PIVOT_SORT_MANUAL, nullptr), 0);
  ASSERT_EQ(fm_workbook_pivot_field_clear_date_group(wb.handle, 0, pivot_idx, 0), 0);
  ASSERT_EQ(fm_workbook_pivot_field_at(wb.handle, 0, pivot_idx, 0, &region), 0);
  EXPECT_EQ(region.sort_order, FM_PIVOT_SORT_MANUAL);
  EXPECT_EQ(region.has_date_group, 0);

  // Items: the scratch pivot added North and South by label, bound to the shared items.
  std::size_t item_count = 99;
  ASSERT_EQ(fm_workbook_pivot_field_item_count(wb.handle, 0, pivot_idx, 0, &item_count), 0);
  ASSERT_EQ(item_count, 2U);
  ASSERT_EQ(fm_workbook_pivot_field_set_item_visible(wb.handle, 0, pivot_idx, 0, 1, 0), 0);
  fm_pivot_item_info_t item{};
  ASSERT_EQ(fm_workbook_pivot_field_item_at(wb.handle, 0, pivot_idx, 0, 0, &item), 0);
  EXPECT_STREQ(item.name, "North");
  EXPECT_EQ(item.cache_index, 0);
  EXPECT_EQ(item.visible, 1);
  ASSERT_EQ(fm_workbook_pivot_field_item_at(wb.handle, 0, pivot_idx, 0, 1, &item), 0);
  EXPECT_STREQ(item.name, "South");
  EXPECT_EQ(item.cache_index, 1);
  EXPECT_EQ(item.visible, 0);
  ASSERT_EQ(fm_workbook_pivot_field_add_item_at(wb.handle, 0, pivot_idx, 0, 1U, 1), 0);
  ASSERT_EQ(fm_workbook_pivot_field_item_at(wb.handle, 0, pivot_idx, 0, 2, &item), 0);
  EXPECT_STREQ(item.name, "");
  EXPECT_EQ(item.cache_index, 1);
  EXPECT_EQ(item.visible, 1);

  // Data field written with a base field / item.
  fm_pivot_data_field_spec_t written{};
  written.name = "Share";
  written.field_index = 1;
  written.aggregation = FM_PIVOT_AGG_AVERAGE;
  written.number_format = "";
  written.show_as = FM_PIVOT_SHOW_AS_DIFFERENCE_FROM;
  written.show_as_base_field = 0;
  written.show_as_base_item = 1;
  ASSERT_EQ(fm_workbook_pivot_data_field_set(wb.handle, 0, pivot_idx, 0, &written), 0) << fm_last_error_message();
  fm_pivot_data_field_spec_t data_field{};
  ASSERT_EQ(fm_workbook_pivot_data_field_at(wb.handle, 0, pivot_idx, 0, &data_field), 0);
  EXPECT_STREQ(data_field.name, "Share");
  EXPECT_EQ(data_field.field_index, 1U);
  EXPECT_EQ(data_field.aggregation, FM_PIVOT_AGG_AVERAGE);
  EXPECT_EQ(data_field.show_as, FM_PIVOT_SHOW_AS_DIFFERENCE_FROM);
  EXPECT_EQ(data_field.show_as_base_field, 0);
  EXPECT_EQ(data_field.show_as_base_item, 1);

  // Field orders follow the setters.
  fm_pivot_field_spec_t extra{};
  extra.source_name = "Amount";
  extra.axis = FM_PIVOT_AXIS_COL;
  extra.subtotal_top = 1;
  std::size_t extra_idx = 99;
  ASSERT_EQ(fm_workbook_pivot_field_add(wb.handle, 0, pivot_idx, &extra, &extra_idx), 0);
  const std::uint32_t* order = nullptr;
  std::size_t order_count = 99;
  ASSERT_EQ(fm_workbook_pivot_get_col_field_order(wb.handle, 0, pivot_idx, &order, &order_count), 0);
  ASSERT_EQ(order_count, 1U);
  EXPECT_EQ(order[0], static_cast<std::uint32_t>(extra_idx));
  const std::uint32_t swapped[] = {static_cast<std::uint32_t>(extra_idx), 0U};
  ASSERT_EQ(fm_workbook_pivot_set_row_field_order(wb.handle, 0, pivot_idx, swapped, 2U), 0);
  ASSERT_EQ(fm_workbook_pivot_get_row_field_order(wb.handle, 0, pivot_idx, &order, &order_count), 0);
  ASSERT_EQ(order_count, 2U);
  EXPECT_EQ(order[0], static_cast<std::uint32_t>(extra_idx));
  EXPECT_EQ(order[1], 0U);
  ASSERT_EQ(fm_workbook_pivot_get_col_field_order(wb.handle, 0, pivot_idx, &order, &order_count), 0);
  EXPECT_EQ(order_count, 0U);
  fm_pivot_field_info_t moved{};
  ASSERT_EQ(fm_workbook_pivot_field_at(wb.handle, 0, pivot_idx, extra_idx, &moved), 0);
  EXPECT_EQ(moved.axis, FM_PIVOT_AXIS_ROW);

  // A field left out of an axis order sits on no axis.
  const std::uint32_t only_extra[] = {static_cast<std::uint32_t>(extra_idx)};
  ASSERT_EQ(fm_workbook_pivot_set_row_field_order(wb.handle, 0, pivot_idx, only_extra, 1U), 0);
  fm_pivot_field_info_t dropped{};
  ASSERT_EQ(fm_workbook_pivot_field_at(wb.handle, 0, pivot_idx, 0, &dropped), 0);
  EXPECT_EQ(dropped.axis, -1);
}

TEST(FormulonCApiPivot, ReadBackSurvivesSaveAndReload) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();
  ASSERT_EQ(fm_workbook_pivot_set_grand_totals(wb.handle, 0, pivot_idx, 1, 0), 0);
  ASSERT_EQ(fm_workbook_pivot_field_set_item_visible(wb.handle, 0, pivot_idx, 0, 1, 0), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0) << fm_last_error_message();
  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0) << fm_last_error_message();

  std::uint32_t read_cache = 99;
  ASSERT_EQ(fm_workbook_pivot_get_cache_id(reloaded.handle, 0, 0, &read_cache), 0) << fm_last_error_message();
  const char* name = nullptr;
  ASSERT_EQ(fm_workbook_pivot_get_name(reloaded.handle, 0, 0, &name), 0);
  EXPECT_STREQ(name, "PT");
  std::int32_t grand_rows = -1;
  std::int32_t grand_cols = -1;
  ASSERT_EQ(fm_workbook_pivot_get_grand_totals(reloaded.handle, 0, 0, &grand_rows, &grand_cols), 0);
  EXPECT_EQ(grand_rows, 1);
  EXPECT_EQ(grand_cols, 0);

  fm_pivot_field_info_t region{};
  ASSERT_EQ(fm_workbook_pivot_field_at(reloaded.handle, 0, 0, 0, &region), 0);
  EXPECT_STREQ(region.source_name, "Region");
  EXPECT_EQ(region.axis, FM_PIVOT_AXIS_ROW);
  EXPECT_EQ(region.cache_field_index, 0);
  fm_pivot_item_info_t item{};
  ASSERT_EQ(fm_workbook_pivot_field_item_at(reloaded.handle, 0, 0, 0, 1, &item), 0);
  EXPECT_EQ(item.cache_index, 1);
  EXPECT_EQ(item.visible, 0);

  fm_pivot_data_field_spec_t data_field{};
  ASSERT_EQ(fm_workbook_pivot_data_field_at(reloaded.handle, 0, 0, 0, &data_field), 0);
  EXPECT_STREQ(data_field.name, "Sum of Amount");
  EXPECT_EQ(data_field.aggregation, FM_PIVOT_AGG_SUM);
  EXPECT_EQ(data_field.field_index, 1U);

  const std::uint32_t* order = nullptr;
  std::size_t order_count = 99;
  ASSERT_EQ(fm_workbook_pivot_get_row_field_order(reloaded.handle, 0, 0, &order, &order_count), 0);
  ASSERT_EQ(order_count, 1U);
  EXPECT_EQ(order[0], 0U);
}

TEST(FormulonCApiPivot, ReadBackRejectsBadArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  std::uint32_t out_cache = 7;
  EXPECT_EQ(fm_workbook_pivot_get_cache_id(nullptr, 0, pivot_idx, &out_cache),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_workbook_pivot_get_cache_id(wb.handle, 0, pivot_idx, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_workbook_pivot_get_cache_id(wb.handle, 0, pivot_idx + 1U, &out_cache),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_EQ(out_cache, 7U);

  fm_pivot_field_info_t field{};
  EXPECT_EQ(fm_workbook_pivot_field_at(wb.handle, 0, pivot_idx, 99, &field),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_EQ(fm_workbook_pivot_field_at(wb.handle, 0, pivot_idx, 0, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  fm_pivot_item_info_t item{};
  EXPECT_EQ(fm_workbook_pivot_field_item_at(wb.handle, 0, pivot_idx, 0, 99, &item),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  fm_pivot_data_field_spec_t data_field{};
  EXPECT_EQ(fm_workbook_pivot_data_field_at(wb.handle, 0, pivot_idx, 99, &data_field),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  const std::uint32_t* order = nullptr;
  std::size_t count = 0;
  EXPECT_EQ(fm_workbook_pivot_get_row_field_order(wb.handle, 0, pivot_idx, nullptr, &count),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_workbook_pivot_get_col_field_order(wb.handle, 9, pivot_idx, &order, &count),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
}

TEST(FormulonCApiPivot, EmptyDataFieldNameIsDerivedFromTheProfile) {
  EXPECT_EQ(DerivedDataFieldName("win-365-en_US", FM_PIVOT_AGG_SUM, nullptr), "Sum of Amount");
  EXPECT_EQ(DerivedDataFieldName("win-365-en_US", FM_PIVOT_AGG_AVERAGE, ""), "Average of Amount");
  EXPECT_EQ(DerivedDataFieldName("win-365-ja_JP", FM_PIVOT_AGG_SUM, nullptr), "合計 / Amount");
  EXPECT_EQ(DerivedDataFieldName("win-365-ja_JP", FM_PIVOT_AGG_COUNT, ""), "個数 / Amount");
  // An explicit name is kept as given.
  EXPECT_EQ(DerivedDataFieldName("win-365-ja_JP", FM_PIVOT_AGG_SUM, "Total"), "Total");
}

TEST(FormulonCApiPivot, DerivedDataFieldNameUsesTheFieldDisplayNameAndSetPath) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_for_profile("win-365-en_US", &wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  // A renamed source field is named by its caption.
  fm_pivot_field_spec_t caption{};
  caption.source_name = "Amount";
  caption.custom_name = "Sales";
  caption.axis = FM_PIVOT_AXIS_VALUE;
  caption.subtotal_top = 1;
  std::size_t caption_idx = 99;
  ASSERT_EQ(fm_workbook_pivot_field_add(wb.handle, 0, pivot_idx, &caption, &caption_idx), 0);

  fm_pivot_data_field_spec_t spec{};
  spec.field_index = static_cast<std::uint32_t>(caption_idx);
  spec.aggregation = FM_PIVOT_AGG_MAX;
  spec.show_as_base_field = -1;
  spec.show_as_base_item = -1;
  ASSERT_EQ(fm_workbook_pivot_data_field_set(wb.handle, 0, pivot_idx, 0, &spec), 0) << fm_last_error_message();
  fm_pivot_data_field_spec_t read{};
  ASSERT_EQ(fm_workbook_pivot_data_field_at(wb.handle, 0, pivot_idx, 0, &read), 0);
  EXPECT_STREQ(read.name, "Max of Sales");

  // A derived name needs a field to name it after.
  spec.field_index = 99;
  EXPECT_EQ(fm_workbook_pivot_data_field_set(wb.handle, 0, pivot_idx, 0, &spec),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  std::size_t added = 0;
  EXPECT_EQ(fm_workbook_pivot_data_field_add(wb.handle, 0, pivot_idx, &spec, &added),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  ASSERT_EQ(fm_workbook_pivot_data_field_at(wb.handle, 0, pivot_idx, 0, &read), 0);
  EXPECT_STREQ(read.name, "Max of Sales");
}

namespace {

// A cache [Region, Product, Sales, Qty] over three records, and a pivot on it with no fields yet.
void BuildWideCachePivot(fm_workbook_t* wb, std::uint32_t* cache_id, std::size_t* pivot_idx) {
  ASSERT_EQ(fm_workbook_pivot_cache_create(wb, 0, cache_id), 0);
  ASSERT_EQ(fm_workbook_pivot_cache_set_worksheet_source(wb, *cache_id, 1, "A1:D4", "Sheet1", nullptr), 0);
  for (const char* name : {"Region", "Product", "Sales", "Qty"}) {
    std::size_t idx = 0;
    ASSERT_EQ(fm_workbook_pivot_cache_field_add(wb, *cache_id, name, &idx), 0);
  }
  ASSERT_EQ(fm_workbook_pivot_cache_field_add_shared_item_text(wb, *cache_id, 0, "North"), 0);
  ASSERT_EQ(fm_workbook_pivot_cache_field_add_shared_item_text(wb, *cache_id, 0, "South"), 0);
  ASSERT_EQ(fm_workbook_pivot_cache_field_add_shared_item_text(wb, *cache_id, 1, "Pen"), 0);
  const double rows[][4] = {{0, 0, 100, 1}, {1, 0, 200, 2}, {0, 0, 300, 3}};
  for (const auto& row : rows) {
    std::size_t rec = 0;
    ASSERT_EQ(fm_workbook_pivot_cache_record_add(wb, *cache_id, &rec), 0);
    for (std::size_t col = 0; col < 4; ++col) {
      ASSERT_EQ(fm_workbook_pivot_cache_record_set_number(wb, *cache_id, rec, col, row[col]), 0);
    }
  }
  ASSERT_EQ(fm_workbook_pivot_create(wb, 0, "PT", *cache_id, 0, 6, pivot_idx), 0);
}

fm_pivot_field_spec_t FieldSpec(const char* name, fm_pivot_axis_t axis) {
  fm_pivot_field_spec_t spec{};
  spec.source_name = name;
  spec.axis = axis;
  spec.subtotal_top = 1;
  return spec;
}

}  // namespace

TEST(FormulonCApiPivot, FieldAddBindsToTheCacheColumnByName) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot = 0;
  BuildWideCachePivot(wb.handle, &cache_id, &pivot);

  std::size_t region = 99;
  std::size_t sales = 99;
  const fm_pivot_field_spec_t region_spec = FieldSpec("Region", FM_PIVOT_AXIS_ROW);
  const fm_pivot_field_spec_t sales_spec = FieldSpec("Sales", FM_PIVOT_AXIS_VALUE);
  ASSERT_EQ(fm_workbook_pivot_field_add(wb.handle, 0, pivot, &region_spec, &region), 0) << fm_last_error_message();
  ASSERT_EQ(fm_workbook_pivot_field_add(wb.handle, 0, pivot, &sales_spec, &sales), 0) << fm_last_error_message();
  // The index is the cache column's position, not the order the fields were added in.
  EXPECT_EQ(region, 0U);
  EXPECT_EQ(sales, 2U);
  std::size_t count = 0;
  ASSERT_EQ(fm_workbook_pivot_field_count(wb.handle, 0, pivot, &count), 0);
  EXPECT_EQ(count, 4U);

  fm_pivot_field_info_t product{};
  ASSERT_EQ(fm_workbook_pivot_field_at(wb.handle, 0, pivot, 1, &product), 0);
  EXPECT_STREQ(product.source_name, "Product");
  EXPECT_EQ(product.axis, -1);

  fm_pivot_data_field_spec_t df{};
  df.name = "Total";
  df.field_index = static_cast<std::uint32_t>(sales);
  df.aggregation = FM_PIVOT_AGG_SUM;
  df.show_as_base_field = -1;
  df.show_as_base_item = -1;
  std::size_t df_idx = 0;
  ASSERT_EQ(fm_workbook_pivot_data_field_add(wb.handle, 0, pivot, &df, &df_idx), 0);
  PivotCellsGuard projected;
  ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 0, pivot, &projected.handle), 0) << fm_last_error_message();
  double grand_total = 0.0;
  for (const fm_pivot_cell_t& cell : CollectCells(projected.handle)) {
    if (cell.kind == FM_PIVOT_CELL_GRAND_TOTAL && cell.value.kind == FM_VAL_NUMBER) {
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
  ASSERT_EQ(fm_workbook_pivot_field_add(wb.handle, 0, pivot, &renamed, &again), 0);
  EXPECT_EQ(again, 0U);
  ASSERT_EQ(fm_workbook_pivot_field_count(wb.handle, 0, pivot, &count), 0);
  EXPECT_EQ(count, 4U);
  const std::uint32_t* order = nullptr;
  std::size_t order_count = 0;
  ASSERT_EQ(fm_workbook_pivot_get_row_field_order(wb.handle, 0, pivot, &order, &order_count), 0);
  EXPECT_EQ(order_count, 1U);

  // A name the cache lacks is appended after the materialised fields.
  const fm_pivot_field_spec_t extra = FieldSpec("Notes", FM_PIVOT_AXIS_ROW);
  std::size_t appended = 0;
  ASSERT_EQ(fm_workbook_pivot_field_add(wb.handle, 0, pivot, &extra, &appended), 0);
  EXPECT_EQ(appended, 4U);
}

TEST(FormulonCApiPivot, FieldAxisNoneTakesTheFieldOffTheReport) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot = 0;
  BuildWideCachePivot(wb.handle, &cache_id, &pivot);
  std::size_t idx = 0;
  const fm_pivot_field_spec_t region = FieldSpec("Region", FM_PIVOT_AXIS_ROW);
  const fm_pivot_field_spec_t sales = FieldSpec("Sales", FM_PIVOT_AXIS_VALUE);
  const fm_pivot_field_spec_t qty = FieldSpec("Qty", FM_PIVOT_AXIS_VALUE);
  ASSERT_EQ(fm_workbook_pivot_field_add(wb.handle, 0, pivot, &region, &idx), 0);
  ASSERT_EQ(fm_workbook_pivot_field_add(wb.handle, 0, pivot, &sales, &idx), 0);
  ASSERT_EQ(fm_workbook_pivot_field_add(wb.handle, 0, pivot, &qty, &idx), 0);
  for (const std::uint32_t field : {2U, 3U}) {
    fm_pivot_data_field_spec_t df{};
    df.field_index = field;
    df.aggregation = FM_PIVOT_AGG_SUM;
    df.show_as_base_field = -1;
    df.show_as_base_item = -1;
    ASSERT_EQ(fm_workbook_pivot_data_field_add(wb.handle, 0, pivot, &df, &idx), 0) << fm_last_error_message();
  }
  fm_pivot_filter_spec_t filter{};
  filter.axis = FM_PIVOT_AXIS_VALUE;
  filter.field_name = "Qty";
  filter.type = FM_PIVOT_FILTER_VALUE_GREATER_THAN;
  filter.value_kind = FM_PIVOT_FILTER_VALUE_DOUBLE;
  filter.value_double = 0.0;
  filter.data_field_index = 1;
  ASSERT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, pivot, &filter), 0) << fm_last_error_message();

  // Taking Sales off removes its data field; Qty's data field and filter move down one.
  ASSERT_EQ(fm_workbook_pivot_field_set_axis(wb.handle, 0, pivot, 2, -1), 0) << fm_last_error_message();
  fm_pivot_field_info_t info{};
  ASSERT_EQ(fm_workbook_pivot_field_at(wb.handle, 0, pivot, 2, &info), 0);
  EXPECT_EQ(info.axis, -1);
  std::size_t data_fields = 0;
  ASSERT_EQ(fm_workbook_pivot_data_field_count(wb.handle, 0, pivot, &data_fields), 0);
  ASSERT_EQ(data_fields, 1U);
  fm_pivot_data_field_spec_t remaining{};
  ASSERT_EQ(fm_workbook_pivot_data_field_at(wb.handle, 0, pivot, 0, &remaining), 0);
  EXPECT_EQ(remaining.field_index, 3U);
  fm_pivot_filter_spec_t kept{};
  ASSERT_EQ(fm_workbook_pivot_filter_at(wb.handle, 0, pivot, 0, &kept), 0);
  EXPECT_EQ(kept.data_field_index, 0U);

  // A row field leaves its axis order too; Qty then takes its filter along.
  ASSERT_EQ(fm_workbook_pivot_field_set_axis(wb.handle, 0, pivot, 0, -1), 0);
  const std::uint32_t* order = nullptr;
  std::size_t order_count = 9;
  ASSERT_EQ(fm_workbook_pivot_get_row_field_order(wb.handle, 0, pivot, &order, &order_count), 0);
  EXPECT_EQ(order_count, 0U);
  ASSERT_EQ(fm_workbook_pivot_field_set_axis(wb.handle, 0, pivot, 3, -1), 0);
  std::size_t filters = 9;
  ASSERT_EQ(fm_workbook_pivot_filter_count(wb.handle, 0, pivot, &filters), 0);
  EXPECT_EQ(filters, 0U);

  EXPECT_NE(fm_workbook_pivot_field_set_axis(wb.handle, 0, pivot, 0, -2), 0);
  EXPECT_NE(fm_workbook_pivot_field_set_axis(wb.handle, 0, pivot, 0, 4), 0);
}

TEST(FormulonCApiPivot, SaveDropsPivotCachesNoPivotUses) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot = 0;
  BuildWideCachePivot(wb.handle, &cache_id, &pivot);
  const auto saved_with_cache = [&] {
    BufferGuard saved;
    EXPECT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0) << fm_last_error_message();
    return std::string(reinterpret_cast<const char*>(saved.data), saved.len).find("pivotCacheDefinition1.xml") !=
           std::string::npos;
  };
  EXPECT_TRUE(saved_with_cache());

  // Removing the pivot keeps the cache in the model, so a recreated pivot can reuse its id ...
  ASSERT_EQ(fm_workbook_pivot_remove(wb.handle, 0, pivot), 0);
  std::size_t caches = 0;
  ASSERT_EQ(fm_workbook_pivot_cache_count(wb.handle, &caches), 0);
  EXPECT_EQ(caches, 1U);
  // ... but nothing draws from it, so a save leaves it out.
  EXPECT_FALSE(saved_with_cache());
  std::size_t again = 0;
  ASSERT_EQ(fm_workbook_pivot_create(wb.handle, 0, "PT2", cache_id, 0, 6, &again), 0);
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
  ASSERT_EQ(fm_workbook_save(wb, &saved.data, &saved.len), 0) << fm_last_error_message();
  const std::vector<std::uint8_t> archive(saved.data, saved.data + saved.len);
  std::string xml = ExtractZipEntry(archive, "xl/pivotTables/pivotTable1.xml");
  const std::size_t cut = xml.find("</dataFields>");
  ASSERT_NE(cut, std::string::npos) << xml;
  xml.resize(cut + std::string_view("</dataFields>").size());
  xml.append(tail).append("</pivotTableDefinition>");
  const std::vector<std::uint8_t> patched = WithPivotTablePart(archive, xml);
  ASSERT_EQ(fm_workbook_load(patched.data(), patched.size(), &out->handle), 0) << fm_last_error_message();
}

fm_pivot_style_info_t ReadStyleInfo(const fm_workbook_t* wb) {
  fm_pivot_style_info_t style{};
  EXPECT_EQ(fm_workbook_pivot_get_style_info(wb, 0, 0, &style), 0) << fm_last_error_message();
  return style;
}

}  // namespace

TEST(FormulonCApiPivot, CreatedPivotCarriesTheDefaultTableStyle) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  const fm_pivot_style_info_t style = ReadStyleInfo(wb.handle);
  EXPECT_EQ(style.present, 1);
  EXPECT_STREQ(style.name, "PivotStyleLight16");
  EXPECT_EQ(style.show_row_headers, 1);
  EXPECT_EQ(style.show_col_headers, 1);
  EXPECT_EQ(style.show_row_stripes, 0);
  EXPECT_EQ(style.show_col_stripes, 0);
  EXPECT_EQ(style.show_last_column, 1);

  const std::string xml = SavedPivotXml(wb.handle);
  EXPECT_NE(xml.find("<pivotTableStyleInfo name=\"PivotStyleLight16\" showRowHeaders=\"1\" showColHeaders=\"1\" "
                     "showRowStripes=\"0\" showColStripes=\"0\" showLastColumn=\"1\"/>"),
            std::string::npos)
      << xml;
  EXPECT_LT(xml.find("</dataFields>"), xml.find("<pivotTableStyleInfo"));
  EXPECT_LT(xml.find("<pivotTableStyleInfo"), xml.find("<extLst"));
}

TEST(FormulonCApiPivot, StyleInfoRoundTripsThroughSetSaveAndLoad) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  ASSERT_EQ(fm_workbook_pivot_set_style_info(wb.handle, 0, pivot_idx, "PivotStyleMedium9", 0, 1, 1, 1, 0), 0)
      << fm_last_error_message();
  fm_pivot_style_info_t style = ReadStyleInfo(wb.handle);
  EXPECT_STREQ(style.name, "PivotStyleMedium9");
  EXPECT_EQ(style.show_row_headers, 0);
  EXPECT_EQ(style.show_col_headers, 1);
  EXPECT_EQ(style.show_row_stripes, 1);
  EXPECT_EQ(style.show_col_stripes, 1);
  EXPECT_EQ(style.show_last_column, 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0) << fm_last_error_message();
  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0) << fm_last_error_message();
  style = ReadStyleInfo(reloaded.handle);
  EXPECT_EQ(style.present, 1);
  EXPECT_STREQ(style.name, "PivotStyleMedium9");
  EXPECT_EQ(style.show_row_headers, 0);
  EXPECT_EQ(style.show_col_headers, 1);
  EXPECT_EQ(style.show_row_stripes, 1);
  EXPECT_EQ(style.show_col_stripes, 1);
  EXPECT_EQ(style.show_last_column, 0);

  // A NULL name keeps the element but drops the style name.
  ASSERT_EQ(fm_workbook_pivot_set_style_info(wb.handle, 0, pivot_idx, nullptr, 1, 1, 0, 0, 0), 0);
  style = ReadStyleInfo(wb.handle);
  EXPECT_EQ(style.present, 1);
  EXPECT_STREQ(style.name, "");
  const std::string xml = SavedPivotXml(wb.handle);
  EXPECT_NE(xml.find("<pivotTableStyleInfo showRowHeaders=\"1\""), std::string::npos) << xml;
}

TEST(FormulonCApiPivot, StyleInfoRejectsBadArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  fm_pivot_style_info_t style{};
  EXPECT_EQ(fm_workbook_pivot_get_style_info(nullptr, 0, pivot_idx, &style),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_workbook_pivot_get_style_info(wb.handle, 0, pivot_idx, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_workbook_pivot_get_style_info(wb.handle, 0, pivot_idx + 1U, &style),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_NE(fm_workbook_pivot_set_style_info(wb.handle, 0, pivot_idx + 1U, "X", 1, 1, 0, 0, 0), 0);
  EXPECT_NE(fm_workbook_pivot_set_style_info(nullptr, 0, pivot_idx, "X", 1, 1, 0, 0, 0), 0);
}

TEST(FormulonCApiPivot, LoadedStyleInfoKeepsSchemaOrderBetweenFormatsAndExtLst) {
  WorkbookGuard source;
  ASSERT_EQ(fm_workbook_create(&source.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(source.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  WorkbookGuard loaded;
  LoadWithPivotTail(source.handle,
                    "<formats count=\"1\"><format action=\"formatting\"><pivotArea outline=\"0\" "
                    "collapsedLevelsAreSubtotals=\"1\" fieldPosition=\"0\"/></format></formats>"
                    "<pivotTableStyleInfo name=\"PivotStyleDark3\" showRowHeaders=\"0\" showColHeaders=\"1\" "
                    "showRowStripes=\"1\" showColStripes=\"0\" showLastColumn=\"0\"/>"
                    "<extLst><ext uri=\"{FF0F0A60-0000-4000-8000-000000000001}\" "
                    "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"/></extLst>",
                    &loaded);
  const fm_pivot_style_info_t style = ReadStyleInfo(loaded.handle);
  EXPECT_EQ(style.present, 1);
  EXPECT_STREQ(style.name, "PivotStyleDark3");
  EXPECT_EQ(style.show_row_headers, 0);
  EXPECT_EQ(style.show_col_headers, 1);
  EXPECT_EQ(style.show_row_stripes, 1);
  EXPECT_EQ(style.show_col_stripes, 0);
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
  ASSERT_EQ(fm_workbook_create(&source.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(source.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  WorkbookGuard loaded;
  LoadWithPivotTail(source.handle, "", &loaded);
  const fm_pivot_style_info_t style = ReadStyleInfo(loaded.handle);
  EXPECT_EQ(style.present, 0);
  EXPECT_STREQ(style.name, "");
  EXPECT_EQ(style.show_row_headers, 0);
  EXPECT_EQ(style.show_last_column, 0);
  EXPECT_EQ(SavedPivotXml(loaded.handle).find("<pivotTableStyleInfo"), std::string::npos);
}

TEST(FormulonCApiPivot, SettingAWorksheetSourceDropsTheCapturedSourceKind) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();
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

  ASSERT_EQ(fm_workbook_pivot_cache_set_worksheet_source(wb.handle, cache_id, 1, "A1:B5", "Sheet1", nullptr), 0)
      << fm_last_error_message();
  EXPECT_TRUE(cache->cache_source_attrs().empty());
  EXPECT_TRUE(cache->cache_source_children_xml().empty());

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0) << fm_last_error_message();
  const std::string definition = ExtractZipEntry(std::vector<std::uint8_t>(saved.data, saved.data + saved.len),
                                                 "xl/pivotCache/pivotCacheDefinition1.xml");
  EXPECT_NE(definition.find("<cacheSource type=\"worksheet\"><worksheetSource"), std::string::npos) << definition;
  EXPECT_EQ(definition.find("consolidation"), std::string::npos) << definition;
}
