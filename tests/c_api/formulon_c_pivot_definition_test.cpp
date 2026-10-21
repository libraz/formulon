// Stable C ABI PivotTable definition read-back and data-field naming tests.

#include "formulon_c_pivot_test_helpers.h"

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
