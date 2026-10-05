// Stable C ABI PivotTable lifecycle and number-format tests.

#include "formulon_c_pivot_test_helpers.h"

TEST(FormulonCApiPivot, RemovePivotAndCache) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  std::size_t before_pivots = 0;
  std::size_t before_caches = 0;
  ASSERT_EQ(fm_workbook_pivot_count(wb.handle, 0, &before_pivots), 0);
  ASSERT_EQ(fm_workbook_pivot_cache_count(wb.handle, &before_caches), 0);
  EXPECT_EQ(before_pivots, 1U);
  EXPECT_EQ(before_caches, 1U);

  // Remove the pivot first, then the cache.
  ASSERT_EQ(fm_workbook_pivot_remove(wb.handle, 0, pivot_idx), 0);
  std::size_t after_remove_pivots = 99;
  ASSERT_EQ(fm_workbook_pivot_count(wb.handle, 0, &after_remove_pivots), 0);
  EXPECT_EQ(after_remove_pivots, 0U);

  ASSERT_EQ(fm_workbook_pivot_cache_remove(wb.handle, cache_id), 0);
  std::size_t after_caches = 99;
  ASSERT_EQ(fm_workbook_pivot_cache_count(wb.handle, &after_caches), 0);
  EXPECT_EQ(after_caches, 0U);
}

TEST(FormulonCApiPivot, CacheRemoveBlockedByPivot) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  EXPECT_EQ(fm_workbook_pivot_cache_remove(wb.handle, cache_id),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));

  // Removing a non-existent cache id must also surface kInvalidArgument.
  EXPECT_EQ(fm_workbook_pivot_cache_remove(wb.handle, 9999U),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
}

TEST(FormulonCApiPivot, NullPointerArgumentsRejected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  ASSERT_EQ(fm_workbook_pivot_cache_create(wb.handle, 0, &cache_id), 0);
  std::size_t field_idx = 99;
  ASSERT_EQ(fm_workbook_pivot_cache_field_add(wb.handle, cache_id, "F", &field_idx), 0);

  std::size_t out_count = 0;
  EXPECT_EQ(fm_workbook_pivot_cache_count(nullptr, &out_count),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_workbook_pivot_cache_count(wb.handle, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));

  std::size_t fc = 0;
  EXPECT_EQ(fm_workbook_pivot_cache_field_count(nullptr, cache_id, &fc),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));

  EXPECT_EQ(fm_workbook_pivot_cache_field_add(wb.handle, cache_id, nullptr, &field_idx),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));

  EXPECT_EQ(fm_workbook_pivot_cache_field_add_shared_item_text(wb.handle, cache_id, field_idx, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));

  std::size_t pivot_idx = 99;
  EXPECT_EQ(fm_workbook_pivot_create(wb.handle, 0, nullptr, cache_id, 0, 0, &pivot_idx),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));

  // spec arguments
  EXPECT_EQ(fm_workbook_pivot_field_add(wb.handle, 0, 0, nullptr, &pivot_idx),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));

  fm_pivot_filter_spec_t bad_filter{};
  bad_filter.field_name = nullptr;
  bad_filter.type = FM_PIVOT_FILTER_LABEL_CONTAINS;
  bad_filter.value_kind = FM_PIVOT_FILTER_VALUE_TEXT;
  bad_filter.value_text = "X";
  bad_filter.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  EXPECT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, 0, &bad_filter),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
}

TEST(FormulonCApiPivot, NullHandleStatusMatchesRecordedError) {
  // The shared-item and record setters resolve their target through a
  // common lookup helper. The status each returns must be the code that
  // helper recorded, so a caller branching on the return value and a
  // caller reading `fm_last_error_message` reach the same conclusion.
  const auto null_pointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);
  const auto invalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);

  const auto expect_null_handle = [&](fm_status_t status) {
    EXPECT_EQ(status, null_pointer);
    EXPECT_NE(std::string(fm_last_error_message()).find("wb is NULL"), std::string::npos)
        << "message: " << fm_last_error_message();
  };

  expect_null_handle(fm_workbook_pivot_cache_field_add_shared_item_number(nullptr, 0U, 0U, 1.0));
  expect_null_handle(fm_workbook_pivot_cache_field_add_shared_item_bool(nullptr, 0U, 0U, 1));
  expect_null_handle(fm_workbook_pivot_cache_field_add_shared_item_blank(nullptr, 0U, 0U));
  expect_null_handle(fm_workbook_pivot_cache_field_add_shared_item_error(nullptr, 0U, 0U, /*error=*/1));
  expect_null_handle(fm_workbook_pivot_cache_field_clear_shared_items(nullptr, 0U, 0U));
  expect_null_handle(fm_workbook_pivot_cache_record_set_number(nullptr, 0U, 0U, 0U, 1.0));
  expect_null_handle(fm_workbook_pivot_cache_record_set_text(nullptr, 0U, 0U, 0U, "x"));
  expect_null_handle(fm_workbook_pivot_cache_record_set_bool(nullptr, 0U, 0U, 0U, 1));
  expect_null_handle(fm_workbook_pivot_cache_record_set_blank(nullptr, 0U, 0U, 0U));
  expect_null_handle(fm_workbook_pivot_cache_record_set_error(nullptr, 0U, 0U, 0U, /*error=*/1));

  // A live handle with an out-of-range coordinate keeps reporting
  // kInvalidArgument through the same helpers.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  ASSERT_EQ(fm_workbook_pivot_cache_create(wb.handle, 0, &cache_id), 0);
  std::size_t field_idx = 99;
  ASSERT_EQ(fm_workbook_pivot_cache_field_add(wb.handle, cache_id, "F", &field_idx), 0);

  EXPECT_EQ(fm_workbook_pivot_cache_field_add_shared_item_number(wb.handle, cache_id, field_idx + 1U, 1.0), invalid);
  EXPECT_EQ(fm_workbook_pivot_cache_field_add_shared_item_number(wb.handle, cache_id + 100U, field_idx, 1.0), invalid);
  EXPECT_EQ(fm_workbook_pivot_cache_record_set_number(wb.handle, cache_id, /*record_idx=*/0U, field_idx, 1.0), invalid);
}

TEST(FormulonCApiPivot, OutOfGridAnchorRejected) {
  // Excel grid ceilings; kept as literals so this test needs no core
  // header. Mirrors Sheet::kMaxRows / kMaxCols.
  constexpr std::uint32_t kMaxRows = 1'048'576U;
  constexpr std::uint32_t kMaxCols = 16'384U;

  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  ASSERT_EQ(fm_workbook_pivot_cache_create(wb.handle, 0, &cache_id), 0);
  std::size_t field_idx = 99;
  ASSERT_EQ(fm_workbook_pivot_cache_field_add(wb.handle, cache_id, "F", &field_idx), 0);

  // create() with an anchor past the grid must be rejected, not stored.
  std::size_t pivot_idx = 99;
  EXPECT_EQ(fm_workbook_pivot_create(wb.handle, 0, "PT", cache_id, kMaxRows, 0U, &pivot_idx),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_EQ(fm_workbook_pivot_create(wb.handle, 0, "PT", cache_id, 0U, kMaxCols, &pivot_idx),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));

  // A valid create succeeds, then set_anchor must reject a span whose far
  // corner leaves the grid (the wrap the audit flagged) and a zero span.
  ASSERT_EQ(fm_workbook_pivot_create(wb.handle, 0, "PT", cache_id, 0U, 0U, &pivot_idx), 0);
  EXPECT_EQ(
      fm_workbook_pivot_set_anchor(wb.handle, 0, pivot_idx, kMaxRows - 1U, 0U, /*span_rows=*/2U, /*span_cols=*/1U),
      static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_EQ(fm_workbook_pivot_set_anchor(wb.handle, 0, pivot_idx, 0U, 0U, /*span_rows=*/0U, /*span_cols=*/1U),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  // An in-grid span still succeeds.
  EXPECT_EQ(fm_workbook_pivot_set_anchor(wb.handle, 0, pivot_idx, 0U, 0U, /*span_rows=*/3U, /*span_cols=*/2U), 0);
}

TEST(FormulonCApiPivot, MutationAndClearExportsCompleteTheirLifecycles) {
  // Exercise the C ABI setters that are not needed by the projected-layout
  // examples above.  Each mutation is followed by the corresponding clear
  // where one exists, which also verifies that the handle remains usable
  // across each invalidation boundary.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  std::uint32_t id_at_zero = 0;
  ASSERT_EQ(fm_workbook_pivot_cache_id_at(wb.handle, 0, &id_at_zero), 0);
  EXPECT_EQ(id_at_zero, cache_id);

  const char* field_name = nullptr;
  ASSERT_EQ(fm_workbook_pivot_cache_field_name(wb.handle, cache_id, 0, &field_name), 0);
  EXPECT_STREQ(field_name, "Region");

  ASSERT_EQ(fm_workbook_pivot_cache_field_add_shared_item_number(wb.handle, cache_id, 0, 7.0), 0);
  ASSERT_EQ(fm_workbook_pivot_cache_field_add_shared_item_bool(wb.handle, cache_id, 0, 1), 0);
  ASSERT_EQ(fm_workbook_pivot_cache_field_add_shared_item_blank(wb.handle, cache_id, 0), 0);
  ASSERT_EQ(fm_workbook_pivot_cache_field_clear_shared_items(wb.handle, cache_id, 0), 0);
  std::size_t shared_count = 99;
  ASSERT_EQ(fm_workbook_pivot_cache_field_shared_item_count(wb.handle, cache_id, 0, &shared_count), 0);
  EXPECT_EQ(shared_count, 0U);

  std::size_t record_idx = 99;
  ASSERT_EQ(fm_workbook_pivot_cache_record_add(wb.handle, cache_id, &record_idx), 0);
  ASSERT_EQ(fm_workbook_pivot_cache_record_set_text(wb.handle, cache_id, record_idx, 0, "North"), 0);
  ASSERT_EQ(fm_workbook_pivot_cache_record_set_bool(wb.handle, cache_id, record_idx, 0, 1), 0);
  ASSERT_EQ(fm_workbook_pivot_cache_record_set_blank(wb.handle, cache_id, record_idx, 0), 0);
  ASSERT_EQ(fm_workbook_pivot_cache_record_set_error(wb.handle, cache_id, record_idx, 0, 1), 0);
  ASSERT_EQ(fm_workbook_pivot_cache_record_clear(wb.handle, cache_id), 0);
  std::size_t record_count = 99;
  ASSERT_EQ(fm_workbook_pivot_cache_record_count(wb.handle, cache_id, &record_count), 0);
  EXPECT_EQ(record_count, 0U);

  ASSERT_EQ(fm_workbook_pivot_set_name(wb.handle, 0, pivot_idx, "Renamed"), 0);
  ASSERT_EQ(fm_workbook_pivot_set_grand_totals(wb.handle, 0, pivot_idx, 0, 1), 0);
  std::size_t field_count = 99;
  ASSERT_EQ(fm_workbook_pivot_field_count(wb.handle, 0, pivot_idx, &field_count), 0);
  ASSERT_EQ(field_count, 2U);

  ASSERT_EQ(fm_workbook_pivot_field_set_axis(wb.handle, 0, pivot_idx, 0, FM_PIVOT_AXIS_ROW), 0);
  ASSERT_EQ(fm_workbook_pivot_field_set_sort(wb.handle, 0, pivot_idx, 0, 0, "Amount"), 0);
  ASSERT_EQ(fm_workbook_pivot_field_set_subtotal_top(wb.handle, 0, pivot_idx, 0, 1), 0);
  ASSERT_EQ(fm_workbook_pivot_field_set_item_visible(wb.handle, 0, pivot_idx, 0, 0, 0), 0);
  ASSERT_EQ(fm_workbook_pivot_field_clear_items(wb.handle, 0, pivot_idx, 0), 0);
  ASSERT_EQ(fm_workbook_pivot_field_add_subtotal_fn(wb.handle, 0, pivot_idx, 0, FM_PIVOT_AGG_COUNT), 0);
  ASSERT_EQ(fm_workbook_pivot_field_clear_subtotal_fns(wb.handle, 0, pivot_idx, 0), 0);
  ASSERT_EQ(fm_workbook_pivot_field_set_date_group(wb.handle, 0, pivot_idx, 0, FM_PIVOT_DATE_YEAR,
                                                   FM_PIVOT_CALENDAR_GREGORIAN, 2020, 2026, 1, -1.0, -1.0),
            0);
  ASSERT_EQ(fm_workbook_pivot_field_clear_date_group(wb.handle, 0, pivot_idx, 0), 0);
  ASSERT_EQ(fm_workbook_pivot_field_set_number_format(wb.handle, 0, pivot_idx, 1, "4"), 0);
  const std::uint32_t col_order[] = {1U};
  ASSERT_EQ(fm_workbook_pivot_set_col_field_order(wb.handle, 0, pivot_idx, col_order, 1U), 0);

  ASSERT_EQ(fm_workbook_pivot_data_field_clear(wb.handle, 0, pivot_idx), 0);
  std::size_t data_field_count = 99;
  ASSERT_EQ(fm_workbook_pivot_data_field_count(wb.handle, 0, pivot_idx, &data_field_count), 0);
  EXPECT_EQ(data_field_count, 0U);

  fm_pivot_filter_spec_t filter{};
  filter.axis = FM_PIVOT_AXIS_ROW;
  filter.field_name = "Region";
  filter.type = FM_PIVOT_FILTER_LABEL_BEGINS_WITH;
  filter.value_kind = FM_PIVOT_FILTER_VALUE_TEXT;
  filter.value_text = "N";
  filter.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  ASSERT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, pivot_idx, &filter), 0);
  ASSERT_EQ(fm_workbook_pivot_filter_clear(wb.handle, 0, pivot_idx), 0);

  ASSERT_EQ(fm_workbook_pivot_field_clear(wb.handle, 0, pivot_idx), 0);
  ASSERT_EQ(fm_workbook_pivot_field_count(wb.handle, 0, pivot_idx, &field_count), 0);
  EXPECT_EQ(field_count, 0U);
  ASSERT_EQ(fm_workbook_pivot_cache_field_clear(wb.handle, cache_id), 0);
  ASSERT_EQ(fm_workbook_pivot_cache_field_count(wb.handle, cache_id, &field_count), 0);
  EXPECT_EQ(field_count, 0U);
}

TEST(FormulonCApiPivot, FieldNumberFormatRoundTripsAsNumFmtId) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();
  std::uint16_t custom_id = 0;
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "0.000", &custom_id), 0) << fm_last_error_message();
  const std::string custom = std::to_string(custom_id);
  ASSERT_EQ(fm_workbook_pivot_field_set_number_format(wb.handle, 0, pivot_idx, 1, custom.c_str()), 0)
      << fm_last_error_message();

  const std::string first = SavedPivotXml(wb.handle);
  const std::string attr = "numFmtId=\"" + custom + "\"";
  EXPECT_EQ(CountOccurrences(first, attr), 1U) << first;

  // A reload reads the attribute into the model rather than into the
  // passthrough bin, so the next save still emits it exactly once.
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0) << fm_last_error_message();
  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0) << fm_last_error_message();
  const std::string second = SavedPivotXml(reloaded.handle);
  EXPECT_EQ(CountOccurrences(second, attr), 1U) << second;
  EXPECT_NE(second.find("<pivotField dataField=\"1\" " + attr), std::string::npos) << second;
}

TEST(FormulonCApiPivot, NumberFormatRejectsAnythingButAKnownNumFmtId) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();
  const fm_status_t invalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);

  // A format code, a signed id, an id nothing registered, one past uint16.
  for (const char* bad : {"0.00", "-4", "200", "65536"}) {
    EXPECT_EQ(fm_workbook_pivot_field_set_number_format(wb.handle, 0, pivot_idx, 1, bad), invalid) << bad;

    fm_pivot_field_spec_t field_spec{};
    field_spec.source_name = "Amount";
    field_spec.axis = FM_PIVOT_AXIS_VALUE;
    field_spec.number_format = bad;
    std::size_t field_idx = 99;
    EXPECT_EQ(fm_workbook_pivot_field_add(wb.handle, 0, pivot_idx, &field_spec, &field_idx), invalid) << bad;

    fm_pivot_data_field_spec_t df_spec{};
    df_spec.name = "Sum of Amount";
    df_spec.field_index = 1U;
    df_spec.aggregation = FM_PIVOT_AGG_SUM;
    df_spec.number_format = bad;
    df_spec.show_as_base_field = -1;
    df_spec.show_as_base_item = -1;
    std::size_t df_idx = 99;
    EXPECT_EQ(fm_workbook_pivot_data_field_add(wb.handle, 0, pivot_idx, &df_spec, &df_idx), invalid) << bad;
    EXPECT_EQ(fm_workbook_pivot_data_field_set(wb.handle, 0, pivot_idx, 0, &df_spec), invalid) << bad;
  }
  std::size_t field_count = 0;
  ASSERT_EQ(fm_workbook_pivot_field_count(wb.handle, 0, pivot_idx, &field_count), 0);
  EXPECT_EQ(field_count, 2U);
  std::size_t df_count = 0;
  ASSERT_EQ(fm_workbook_pivot_data_field_count(wb.handle, 0, pivot_idx, &df_count), 0);
  EXPECT_EQ(df_count, 1U);

  // A built-in id and the empty string (clear) are accepted.
  EXPECT_EQ(fm_workbook_pivot_field_set_number_format(wb.handle, 0, pivot_idx, 1, "4"), 0) << fm_last_error_message();
  EXPECT_EQ(fm_workbook_pivot_field_set_number_format(wb.handle, 0, pivot_idx, 1, ""), 0) << fm_last_error_message();
}
