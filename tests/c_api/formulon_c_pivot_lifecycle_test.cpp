// Stable C ABI PivotTable lifecycle, number-format and runtime-id tests.

#include <limits>
#include <string>

#include "formulon_c_out_index.h"
#include "formulon_c_pivot_test_helpers.h"
#include "formulon_c_record_arg.h"

TEST(FormulonCApiPivot, RemovePivotAndCache) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();

  std::size_t before_pivots = 0;
  std::size_t before_caches = 0;
  ASSERT_EQ(fm_sheet_get_pivot_count(wb.handle, 0, &before_pivots), 0);
  ASSERT_EQ(fm_workbook_get_pivot_cache_count(wb.handle, &before_caches), 0);
  EXPECT_EQ(before_pivots, 1U);
  EXPECT_EQ(before_caches, 1U);

  // Remove the pivot first, then the cache.
  ASSERT_EQ(fm_sheet_remove_pivot(wb.handle, 0, pivot_index), 0);
  std::size_t after_remove_pivots = 99;
  ASSERT_EQ(fm_sheet_get_pivot_count(wb.handle, 0, &after_remove_pivots), 0);
  EXPECT_EQ(after_remove_pivots, 0U);

  ASSERT_EQ(fm_workbook_remove_pivot_cache(wb.handle, cache_id), 0);
  std::size_t after_caches = 99;
  ASSERT_EQ(fm_workbook_get_pivot_cache_count(wb.handle, &after_caches), 0);
  EXPECT_EQ(after_caches, 0U);
}

TEST(FormulonCApiPivot, AutoPivotCacheIdRejectsOverflowAndPreservesOutput) {
  const std::uint32_t max_id = std::numeric_limits<std::uint32_t>::max();
  const fm_status_t invalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);

  WorkbookGuard at_ceiling;
  ASSERT_EQ(fm_workbook_create_default(&at_ceiling.handle), 0);
  std::uint32_t assigned = 0;
  ASSERT_EQ(fm_workbook_add_pivot_cache(at_ceiling.handle, max_id, &assigned), 0) << fm_get_last_error_message();
  EXPECT_EQ(assigned, max_id);
  std::size_t cache_count = 0;
  ASSERT_EQ(fm_workbook_get_pivot_cache_count(at_ceiling.handle, &cache_count), 0);
  EXPECT_EQ(cache_count, 1U);

  assigned = 0xA5A5A5A5U;
  EXPECT_EQ(fm_workbook_add_pivot_cache(at_ceiling.handle, 0U, &assigned), invalid);
  EXPECT_EQ(assigned, 0xA5A5A5A5U);
  ASSERT_EQ(fm_workbook_get_pivot_cache_count(at_ceiling.handle, &cache_count), 0);
  EXPECT_EQ(cache_count, 1U);

  // A free explicit id below the ceiling remains valid, and auto-assignment
  // can still allocate the next id when max(existing_ids) is max-1.
  WorkbookGuard below_ceiling;
  ASSERT_EQ(fm_workbook_create_default(&below_ceiling.handle), 0);
  assigned = 0;
  ASSERT_EQ(fm_workbook_add_pivot_cache(below_ceiling.handle, max_id - 1U, &assigned), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(assigned, max_id - 1U);
  assigned = 0;
  ASSERT_EQ(fm_workbook_add_pivot_cache(below_ceiling.handle, 0U, &assigned), 0) << fm_get_last_error_message();
  EXPECT_EQ(assigned, max_id);
}

TEST(FormulonCApiPivot, RecreatedPivotWakesGetPivotDataAfterRef) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=GETPIVOTDATA(\"Sum of Amount\",D1,\"Region\",\"North\")"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  ASSERT_EQ(fm_sheet_remove_pivot(wb.handle, 0, pivot_index), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  fm_value_t value{};
  ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, 0, 0, &value), 0);
  ASSERT_EQ(value.kind, FM_VALUE_KIND_ERROR);
  EXPECT_EQ(value.u.error_code, static_cast<std::int32_t>(formulon::ErrorCode::Ref));

  // Recreating and configuring a pivot at the same anchor must dirty the
  // already existing formula, so a subsequent recalc recovers from #REF!.
  std::uint32_t recreated_cache_id = 0;
  std::size_t recreated_pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &recreated_cache_id, &recreated_pivot_index), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, 0, 0, &value), 0);
  ASSERT_EQ(value.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_DOUBLE_EQ(value.u.number, 400.0);
}

TEST(FormulonCApiPivot, CacheRemoveBlockedByPivot) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();

  EXPECT_EQ(fm_workbook_remove_pivot_cache(wb.handle, cache_id),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));

  // Removing a non-existent cache id is a missing key.
  EXPECT_EQ(fm_workbook_remove_pivot_cache(wb.handle, 9999U), FM_STATUS_NOT_FOUND);
}

TEST(FormulonCApiPivot, NullPointerArgumentsRejected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  ASSERT_EQ(fm_workbook_add_pivot_cache(wb.handle, 0, &cache_id), 0);
  std::size_t field_index = 99;
  ASSERT_EQ(fm_workbook_add_pivot_cache_field(wb.handle, cache_id, "F", &field_index), 0);

  std::size_t out_count = 0;
  EXPECT_EQ(fm_workbook_get_pivot_cache_count(nullptr, &out_count),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_workbook_get_pivot_cache_count(wb.handle, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));

  std::size_t fc = 0;
  EXPECT_EQ(fm_workbook_get_pivot_cache_field_count(nullptr, cache_id, &fc),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));

  EXPECT_EQ(fm_workbook_add_pivot_cache_field(wb.handle, cache_id, nullptr, &field_index),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));

  EXPECT_EQ(fm_workbook_add_pivot_cache_field_shared_item_text(wb.handle, cache_id, field_index, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));

  std::size_t pivot_index = 99;
  EXPECT_EQ(fm_sheet_add_pivot(wb.handle, 0, nullptr, cache_id, 0, 0, &pivot_index),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));

  // spec arguments
  EXPECT_EQ(fm_sheet_add_pivot_field(wb.handle, 0, 0, nullptr, &pivot_index),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));

  fm_pivot_filter_spec_t bad_filter = FM_STRUCT_INIT(fm_pivot_filter_spec_t);
  bad_filter.field_name = nullptr;
  bad_filter.type = FM_PIVOT_FILTER_KIND_LABEL_CONTAINS;
  bad_filter.value_kind = FM_PIVOT_FILTER_OPERAND_KIND_TEXT;
  bad_filter.value_text = "X";
  bad_filter.value_high_kind = FM_PIVOT_FILTER_OPERAND_KIND_NONE;
  EXPECT_EQ(fm_sheet_add_pivot_filter(wb.handle, 0, 0, &bad_filter, ignored_index()),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
}

TEST(FormulonCApiPivot, NullHandleStatusMatchesRecordedError) {
  // The shared-item and record setters resolve their target through a
  // common lookup helper. The status each returns must be the code that
  // helper recorded, so a caller branching on the return value and a
  // caller reading `fm_get_last_error_message` reach the same conclusion.
  const auto null_pointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);

  const auto expect_null_handle = [&](fm_status_t status) {
    EXPECT_EQ(status, null_pointer);
    EXPECT_NE(std::string(fm_get_last_error_message()).find("wb is NULL"), std::string::npos)
        << "message: " << fm_get_last_error_message();
  };

  expect_null_handle(fm_workbook_add_pivot_cache_field_shared_item_number(nullptr, 0U, 0U, 1.0));
  expect_null_handle(fm_workbook_add_pivot_cache_field_shared_item_bool(nullptr, 0U, 0U, 1));
  expect_null_handle(fm_workbook_add_pivot_cache_field_shared_item_blank(nullptr, 0U, 0U));
  expect_null_handle(fm_workbook_add_pivot_cache_field_shared_item_error(nullptr, 0U, 0U, /*error=*/1));
  expect_null_handle(fm_workbook_clear_pivot_cache_field_shared_items(nullptr, 0U, 0U));
  expect_null_handle(fm_workbook_set_pivot_cache_record_number(nullptr, 0U, 0U, 0U, 1.0));
  expect_null_handle(fm_workbook_set_pivot_cache_record_text(nullptr, 0U, 0U, 0U, "x"));
  expect_null_handle(fm_workbook_set_pivot_cache_record_bool(nullptr, 0U, 0U, 0U, 1));
  expect_null_handle(fm_workbook_set_pivot_cache_record_blank(nullptr, 0U, 0U, 0U));
  expect_null_handle(fm_workbook_set_pivot_cache_record_error(nullptr, 0U, 0U, 0U, /*error=*/1));

  // A live handle with an out-of-range field index reports kIndexOutOfRange and
  // an unknown cache id kNotFound, through the same helpers.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  ASSERT_EQ(fm_workbook_add_pivot_cache(wb.handle, 0, &cache_id), 0);
  std::size_t field_index = 99;
  ASSERT_EQ(fm_workbook_add_pivot_cache_field(wb.handle, cache_id, "F", &field_index), 0);

  EXPECT_EQ(fm_workbook_add_pivot_cache_field_shared_item_number(wb.handle, cache_id, field_index + 1U, 1.0),
            FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_workbook_add_pivot_cache_field_shared_item_number(wb.handle, cache_id + 100U, field_index, 1.0),
            FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache_id, /*record_index=*/0U, field_index, 1.0),
            FM_STATUS_INDEX_OUT_OF_RANGE);
}

TEST(FormulonCApiPivot, MissingCacheIdIsNotFound) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t field_count = 0;
  EXPECT_EQ(fm_workbook_get_pivot_cache_field_count(wb.handle, 9999U, &field_count), FM_STATUS_NOT_FOUND);
  std::size_t field_index = 0;
  EXPECT_EQ(fm_workbook_add_pivot_cache_field(wb.handle, 9999U, "F", &field_index), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_workbook_clear_pivot_cache_records(wb.handle, 9999U), FM_STATUS_NOT_FOUND);
  std::size_t pivot_index = 0;
  EXPECT_EQ(fm_sheet_add_pivot(wb.handle, 0, "PT", 9999U, 0U, 0U, &pivot_index), FM_STATUS_NOT_FOUND);
}

TEST(FormulonCApiPivot, ShowDetailsRejectsAnOutOfGridCell) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();
  const auto invalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  std::size_t created = 0;
  EXPECT_EQ(fm_sheet_show_pivot_details(wb.handle, 0, pivot_index, 1'048'576U, 0U, &created), invalid);
  EXPECT_EQ(fm_sheet_show_pivot_details(wb.handle, 0, pivot_index, 0U, 16'384U, &created), invalid);
}

TEST(FormulonCApiPivot, OutOfGridAnchorRejected) {
  // Excel grid ceilings; kept as literals so this test needs no core
  // header. Mirrors Sheet::kMaxRows / kMaxCols.
  constexpr std::uint32_t kMaxRows = 1'048'576U;
  constexpr std::uint32_t kMaxCols = 16'384U;

  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  ASSERT_EQ(fm_workbook_add_pivot_cache(wb.handle, 0, &cache_id), 0);
  std::size_t field_index = 99;
  ASSERT_EQ(fm_workbook_add_pivot_cache_field(wb.handle, cache_id, "F", &field_index), 0);

  // create() with an anchor past the grid must be rejected, not stored.
  std::size_t pivot_index = 99;
  EXPECT_EQ(fm_sheet_add_pivot(wb.handle, 0, "PT", cache_id, kMaxRows, 0U, &pivot_index),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_EQ(fm_sheet_add_pivot(wb.handle, 0, "PT", cache_id, 0U, kMaxCols, &pivot_index),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));

  // A valid create succeeds, then set_anchor must reject a cell outside the grid.
  ASSERT_EQ(fm_sheet_add_pivot(wb.handle, 0, "PT", cache_id, 0U, 0U, &pivot_index), 0);
  EXPECT_EQ(fm_sheet_set_pivot_anchor(wb.handle, 0, pivot_index, kMaxRows, 0U),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_EQ(fm_sheet_set_pivot_anchor(wb.handle, 0, pivot_index, 0U, kMaxCols),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_EQ(fm_sheet_set_pivot_anchor(nullptr, 0, pivot_index, 0U, 0U),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  // An in-grid cell still succeeds.
  EXPECT_EQ(fm_sheet_set_pivot_anchor(wb.handle, 0, pivot_index, 1U, 2U), 0);
}

TEST(FormulonCApiPivot, MutationAndClearExportsCompleteTheirLifecycles) {
  // Exercise the C ABI setters that are not needed by the projected-layout
  // examples above.  Each mutation is followed by the corresponding clear
  // where one exists, which also verifies that the handle remains usable
  // across each invalidation boundary.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();

  std::uint32_t id_at_zero = 0;
  ASSERT_EQ(fm_workbook_get_pivot_cache_id_at(wb.handle, 0, &id_at_zero), 0);
  EXPECT_EQ(id_at_zero, cache_id);

  const char* field_name = nullptr;
  ASSERT_EQ(fm_workbook_get_pivot_cache_field_name(wb.handle, cache_id, 0, &field_name), 0);
  EXPECT_STREQ(field_name, "Region");

  ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_number(wb.handle, cache_id, 0, 7.0), 0);
  ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_bool(wb.handle, cache_id, 0, 1), 0);
  ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_blank(wb.handle, cache_id, 0), 0);
  ASSERT_EQ(fm_workbook_clear_pivot_cache_field_shared_items(wb.handle, cache_id, 0), 0);
  std::size_t shared_count = 99;
  ASSERT_EQ(fm_workbook_get_pivot_cache_field_shared_item_count(wb.handle, cache_id, 0, &shared_count), 0);
  EXPECT_EQ(shared_count, 0U);

  std::size_t record_index = 99;
  ASSERT_EQ(fm_workbook_add_pivot_cache_record(wb.handle, cache_id, &record_index), 0);
  ASSERT_EQ(fm_workbook_set_pivot_cache_record_text(wb.handle, cache_id, record_index, 0, "North"), 0);
  ASSERT_EQ(fm_workbook_set_pivot_cache_record_bool(wb.handle, cache_id, record_index, 0, 1), 0);
  ASSERT_EQ(fm_workbook_set_pivot_cache_record_blank(wb.handle, cache_id, record_index, 0), 0);
  ASSERT_EQ(fm_workbook_set_pivot_cache_record_error(wb.handle, cache_id, record_index, 0, 1), 0);
  ASSERT_EQ(fm_workbook_clear_pivot_cache_records(wb.handle, cache_id), 0);
  std::size_t record_count = 99;
  ASSERT_EQ(fm_workbook_get_pivot_cache_record_count(wb.handle, cache_id, &record_count), 0);
  EXPECT_EQ(record_count, 0U);

  ASSERT_EQ(fm_sheet_rename_pivot(wb.handle, 0, pivot_index, "Renamed"), 0);
  ASSERT_EQ(fm_sheet_update_pivot_grand_totals(wb.handle, 0, pivot_index, 0, 1), 0);
  std::size_t field_count = 99;
  ASSERT_EQ(fm_sheet_get_pivot_field_count(wb.handle, 0, pivot_index, &field_count), 0);
  ASSERT_EQ(field_count, 2U);

  ASSERT_EQ(fm_sheet_set_pivot_field_axis(wb.handle, 0, pivot_index, 0, FM_PIVOT_AXIS_ROW), 0);
  ASSERT_EQ(fm_sheet_set_pivot_field_sort(wb.handle, 0, pivot_index, 0, FM_PIVOT_SORT_ORDER_ASCENDING, "Amount"), 0);
  ASSERT_EQ(fm_sheet_set_pivot_field_subtotal_top(wb.handle, 0, pivot_index, 0, 1), 0);
  ASSERT_EQ(fm_sheet_set_pivot_field_item_visible(wb.handle, 0, pivot_index, 0, 0, 0), 0);
  ASSERT_EQ(fm_sheet_clear_pivot_field_items(wb.handle, 0, pivot_index, 0), 0);
  ASSERT_EQ(fm_sheet_add_pivot_field_subtotal_function(wb.handle, 0, pivot_index, 0, FM_PIVOT_AGGREGATION_COUNT), 0);
  ASSERT_EQ(fm_sheet_clear_pivot_field_subtotal_functions(wb.handle, 0, pivot_index, 0), 0);
  // Region holds text: grouping it by date is refused, and clearing a
  // grouping it does not have changes nothing.
  ASSERT_EQ(fm_sheet_set_pivot_field_date_group(
                wb.handle, 0, pivot_index, 0, FM_PIVOT_PERIOD_YEARS, FM_PIVOT_CALENDAR_GREGORIAN, 1,
                std::numeric_limits<double>::quiet_NaN(), std::numeric_limits<double>::quiet_NaN()),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  ASSERT_EQ(fm_sheet_clear_pivot_field_date_group(wb.handle, 0, pivot_index, 0), 0);
  ASSERT_EQ(fm_sheet_set_pivot_field_number_format(wb.handle, 0, pivot_index, 1, "#,##0.00"), 0);
  const std::uint32_t col_order[] = {1U};
  ASSERT_EQ(fm_sheet_set_pivot_column_field_order(wb.handle, 0, pivot_index, col_order, 1U), 0);

  ASSERT_EQ(fm_sheet_clear_pivot_data_fields(wb.handle, 0, pivot_index), 0);
  std::size_t data_field_count = 99;
  ASSERT_EQ(fm_sheet_get_pivot_data_field_count(wb.handle, 0, pivot_index, &data_field_count), 0);
  EXPECT_EQ(data_field_count, 0U);

  fm_pivot_filter_spec_t filter = FM_STRUCT_INIT(fm_pivot_filter_spec_t);
  filter.axis = FM_PIVOT_AXIS_ROW;
  filter.field_name = "Region";
  filter.type = FM_PIVOT_FILTER_KIND_LABEL_BEGINS_WITH;
  filter.value_kind = FM_PIVOT_FILTER_OPERAND_KIND_TEXT;
  filter.value_text = "N";
  filter.value_high_kind = FM_PIVOT_FILTER_OPERAND_KIND_NONE;
  ASSERT_EQ(fm_sheet_add_pivot_filter(wb.handle, 0, pivot_index, &filter, ignored_index()), 0);
  ASSERT_EQ(fm_sheet_clear_pivot_filters(wb.handle, 0, pivot_index), 0);

  ASSERT_EQ(fm_sheet_clear_pivot_fields(wb.handle, 0, pivot_index), 0);
  ASSERT_EQ(fm_sheet_get_pivot_field_count(wb.handle, 0, pivot_index, &field_count), 0);
  EXPECT_EQ(field_count, 0U);
  ASSERT_EQ(fm_workbook_clear_pivot_cache_fields(wb.handle, cache_id), 0);
  ASSERT_EQ(fm_workbook_get_pivot_cache_field_count(wb.handle, cache_id, &field_count), 0);
  EXPECT_EQ(field_count, 0U);
}

namespace {

constexpr fm_status_t kNotFound = FM_STATUS_NOT_FOUND;
constexpr fm_status_t kPreconditionFailed = static_cast<fm_status_t>(formulon::FormulonErrorCode::kPreconditionFailed);

std::string FieldNumberFormat(const fm_workbook_t* wb, std::size_t pivot_index, std::size_t field_index) {
  fm_pivot_field_info_t info = FM_STRUCT_INIT(fm_pivot_field_info_t);
  EXPECT_EQ(fm_sheet_get_pivot_field(wb, 0, pivot_index, field_index, &info), 0) << fm_get_last_error_message();
  return info.number_format != nullptr ? info.number_format : "<null>";
}

std::string DataFieldNumberFormat(const fm_workbook_t* wb, std::size_t pivot_index, std::size_t data_field_index) {
  fm_pivot_data_field_spec_t spec = FM_STRUCT_INIT(fm_pivot_data_field_spec_t);
  EXPECT_EQ(fm_sheet_get_pivot_data_field(wb, 0, pivot_index, data_field_index, &spec), 0)
      << fm_get_last_error_message();
  return spec.number_format != nullptr ? spec.number_format : "<null>";
}

fm_pivot_data_field_spec_t SumOfAmount(const char* number_format) {
  fm_pivot_data_field_spec_t spec = FM_STRUCT_INIT(fm_pivot_data_field_spec_t);
  spec.name = "Sum of Amount";
  spec.field_index = 1U;
  spec.aggregation = FM_PIVOT_AGGREGATION_SUM;
  spec.number_format = number_format;
  spec.show_as = FM_PIVOT_SHOW_VALUES_AS_NORMAL;
  spec.show_as_base_field = -1;
  spec.show_as_base_item = -1;
  return spec;
}

std::uint32_t PivotRuntimeId(const fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index) {
  std::uint32_t id = 0;
  EXPECT_EQ(fm_sheet_get_pivot_runtime_id(wb, sheet_index, pivot_index, &id), 0) << fm_get_last_error_message();
  return id;
}

std::string PivotName(const fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index) {
  const char* name = nullptr;
  EXPECT_EQ(fm_sheet_get_pivot_name(wb, sheet_index, pivot_index, &name), 0) << fm_get_last_error_message();
  return name != nullptr ? name : "";
}

}  // namespace

TEST(FormulonCApiPivot, NumberFormatCodeRoundTripsThroughEveryInput) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();
  EXPECT_EQ(FieldNumberFormat(wb.handle, pivot_index, 1), "");
  EXPECT_EQ(DataFieldNumberFormat(wb.handle, pivot_index, 0), "");

  // fm_sheet_set_pivot_field_number_format.
  ASSERT_EQ(fm_sheet_set_pivot_field_number_format(wb.handle, 0, pivot_index, 1, "0.0%"), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(FieldNumberFormat(wb.handle, pivot_index, 1), "0.0%");

  // fm_sheet_add_pivot_field, configuring a cache field in place.
  fm_pivot_field_spec_t field_spec = FM_STRUCT_INIT(fm_pivot_field_spec_t);
  field_spec.source_name = "Region";
  field_spec.custom_name = "";
  field_spec.axis = FM_PIVOT_AXIS_ROW;
  field_spec.subtotal_top = FM_TRISTATE_FALSE;
  field_spec.number_format = "0.0%";
  std::size_t field_index = 99;
  ASSERT_EQ(fm_sheet_add_pivot_field(wb.handle, 0, pivot_index, &field_spec, &field_index), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(FieldNumberFormat(wb.handle, pivot_index, field_index), "0.0%");

  // fm_sheet_set_pivot_data_field and fm_sheet_add_pivot_data_field.
  const fm_pivot_data_field_spec_t percent = SumOfAmount("0.0%");
  ASSERT_EQ(fm_sheet_set_pivot_data_field(wb.handle, 0, pivot_index, 0, &percent), 0) << fm_get_last_error_message();
  EXPECT_EQ(DataFieldNumberFormat(wb.handle, pivot_index, 0), "0.0%");
  fm_pivot_data_field_spec_t count = SumOfAmount("#,##0");
  count.name = "Count of Amount";
  count.aggregation = FM_PIVOT_AGGREGATION_COUNT;
  std::size_t df_index = 99;
  ASSERT_EQ(fm_sheet_add_pivot_data_field(wb.handle, 0, pivot_index, &count, &df_index), 0)
      << fm_get_last_error_message();
  // "#,##0" is the ECMA-376 code of built-in id 3.
  EXPECT_EQ(DataFieldNumberFormat(wb.handle, pivot_index, df_index), "#,##0");

  // The file stores the id the code was registered under, once per carrier.
  std::uint32_t percent_id = 0;
  ASSERT_EQ(fm_workbook_add_number_format(wb.handle, "0.0%", &percent_id), 0);
  EXPECT_GE(percent_id, 164U);
  const std::string saved = SavedPivotXml(wb.handle);
  EXPECT_NE(saved.find("<pivotField dataField=\"1\" numFmtId=\"" + std::to_string(percent_id) + "\""),
            std::string::npos)
      << saved;
  EXPECT_NE(saved.find("numFmtId=\"3\""), std::string::npos) << saved;

  // A string of digits is a code, not an id; NULL and "" clear.
  ASSERT_EQ(fm_sheet_set_pivot_field_number_format(wb.handle, 0, pivot_index, 1, "10"), 0);
  EXPECT_EQ(FieldNumberFormat(wb.handle, pivot_index, 1), "10");
  ASSERT_EQ(fm_sheet_set_pivot_field_number_format(wb.handle, 0, pivot_index, 1, ""), 0);
  EXPECT_EQ(FieldNumberFormat(wb.handle, pivot_index, 1), "");
  const fm_pivot_data_field_spec_t cleared = SumOfAmount(nullptr);
  ASSERT_EQ(fm_sheet_set_pivot_data_field(wb.handle, 0, pivot_index, 0, &cleared), 0) << fm_get_last_error_message();
  EXPECT_EQ(DataFieldNumberFormat(wb.handle, pivot_index, 0), "");
}

TEST(FormulonCApiPivot, NumberFormatRejectsMalformedUtf8OnEveryInput) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();
  const char* bad = "0.0\xC3";

  EXPECT_EQ(fm_sheet_set_pivot_field_number_format(wb.handle, 0, pivot_index, 1, bad),
            FM_STATUS_BINDING_UTF8ENCODING_ERROR);
  fm_pivot_field_spec_t field_spec = FM_STRUCT_INIT(fm_pivot_field_spec_t);
  field_spec.source_name = "Amount";
  field_spec.axis = FM_PIVOT_AXIS_VALUE;
  field_spec.number_format = bad;
  std::size_t field_index = 99;
  EXPECT_EQ(fm_sheet_add_pivot_field(wb.handle, 0, pivot_index, &field_spec, &field_index),
            FM_STATUS_BINDING_UTF8ENCODING_ERROR);
  const fm_pivot_data_field_spec_t df_spec = SumOfAmount(bad);
  std::size_t df_index = 99;
  EXPECT_EQ(fm_sheet_add_pivot_data_field(wb.handle, 0, pivot_index, &df_spec, &df_index),
            FM_STATUS_BINDING_UTF8ENCODING_ERROR);
  EXPECT_EQ(fm_sheet_set_pivot_data_field(wb.handle, 0, pivot_index, 0, &df_spec),
            FM_STATUS_BINDING_UTF8ENCODING_ERROR);
  std::size_t df_count = 0;
  ASSERT_EQ(fm_sheet_get_pivot_data_field_count(wb.handle, 0, pivot_index, &df_count), 0);
  EXPECT_EQ(df_count, 1U);
  EXPECT_EQ(FieldNumberFormat(wb.handle, pivot_index, 1), "");
}

TEST(FormulonCApiPivot, ProfileRenderedBuiltinCodeKeepsItsId) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_profile_id(wb.handle, "mac-365-ja_JP"), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();

  // Id 27 has no ECMA-376 code; ja-JP renders it as the Japanese-era date.
  int32_t present = 0;
  const char* code27 = nullptr;
  ASSERT_EQ(fm_workbook_get_number_format(wb.handle, 27U, &present, &code27), 0);
  ASSERT_EQ(present, 1);
  const std::string code(code27);
  ASSERT_EQ(fm_sheet_set_pivot_field_number_format(wb.handle, 0, pivot_index, 1, code.c_str()), 0)
      << fm_get_last_error_message();
  EXPECT_NE(SavedPivotXml(wb.handle).find("numFmtId=\"27\""), std::string::npos);

  // Reload (the profile is not saved), read the code back and write it unchanged: still 27.
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0) << fm_get_last_error_message();
  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_set_profile_id(reloaded.handle, "mac-365-ja_JP"), 0);
  const std::string read_back = FieldNumberFormat(reloaded.handle, 0, 1);
  EXPECT_EQ(read_back, code);
  ASSERT_EQ(fm_sheet_set_pivot_field_number_format(reloaded.handle, 0, 0, 1, read_back.c_str()), 0)
      << fm_get_last_error_message();
  const std::string resaved = SavedPivotXml(reloaded.handle);
  EXPECT_NE(resaved.find("numFmtId=\"27\""), std::string::npos) << resaved;
  // No custom format was registered for it: the first custom id is still free.
  const char* custom = nullptr;
  ASSERT_EQ(fm_workbook_get_number_format(reloaded.handle, 164U, &present, &custom), 0);
  EXPECT_EQ(present, 0);
}

TEST(FormulonCApiPivot, LayoutValueCellsCarryTheDataFieldFormatCode) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();
  const fm_pivot_data_field_spec_t percent = SumOfAmount("0.0%");
  ASSERT_EQ(fm_sheet_set_pivot_data_field(wb.handle, 0, pivot_index, 0, &percent), 0) << fm_get_last_error_message();

  PivotCellsGuard projected;
  ASSERT_EQ(fm_sheet_get_pivot_layout(wb.handle, 0, pivot_index, &projected.handle), 0) << fm_get_last_error_message();
  std::size_t values = 0;
  for (const fm_pivot_cell_t& cell : CollectCells(projected.handle)) {
    if (cell.value.kind != FM_VALUE_KIND_NUMBER ||
        (cell.kind != FM_PIVOT_CELL_KIND_DATA && cell.kind != FM_PIVOT_CELL_KIND_GRAND_TOTAL)) {
      continue;
    }
    ++values;
    ASSERT_NE(cell.number_format, nullptr);
    EXPECT_STREQ(cell.number_format, "0.0%") << "row=" << cell.row << " col=" << cell.col;
  }
  // North, South and the grand total.
  EXPECT_EQ(values, 3U);
}

TEST(FormulonCApiPivot, RuntimeIdFollowsThePivotAcrossEditsAndIsNeverReused) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();
  std::size_t second = 99;
  ASSERT_EQ(fm_sheet_add_pivot(wb.handle, 0, "PT2", cache_id, 0U, 7U, &second), 0) << fm_get_last_error_message();
  ASSERT_EQ(second, 1U);

  const std::uint32_t first_id = PivotRuntimeId(wb.handle, 0, 0);
  const std::uint32_t second_id = PivotRuntimeId(wb.handle, 0, 1);
  EXPECT_NE(first_id, 0U);
  EXPECT_NE(second_id, 0U);
  EXPECT_NE(first_id, second_id);
  EXPECT_EQ(PivotRuntimeId(wb.handle, 0, 0), first_id);

  std::size_t index = 99;
  EXPECT_EQ(fm_sheet_resolve_pivot_runtime_id(wb.handle, 0, 0U, &index), kNotFound);
  EXPECT_EQ(index, 99U);

  // A pivot edit the projection refuses is rolled back from a clone; the id stays.
  ASSERT_EQ(fm_sheet_add_merge(wb.handle, 0U, record_arg(fm_cell_rect_t{0U, 5U, 1U, 5U})), 0)
      << fm_get_last_error_message();
  fm_pivot_data_field_spec_t widen = SumOfAmount("");
  widen.name = "Count of Amount";
  widen.aggregation = FM_PIVOT_AGGREGATION_COUNT;
  std::size_t df_index = 99;
  ASSERT_EQ(fm_sheet_add_pivot_data_field(wb.handle, 0U, 0U, &widen, &df_index), kPreconditionFailed);
  ASSERT_EQ(fm_sheet_resolve_pivot_runtime_id(wb.handle, 0, first_id, &index), 0) << fm_get_last_error_message();
  EXPECT_EQ(index, 0U);
  EXPECT_EQ(PivotRuntimeId(wb.handle, 0, 0), first_id);

  // A copied sheet's pivots are new objects with ids of their own.
  std::size_t copy = 99;
  ASSERT_EQ(fm_workbook_copy_sheet(wb.handle, 0, 1, "Copy", &copy), 0) << fm_get_last_error_message();
  const std::uint32_t copy_first = PivotRuntimeId(wb.handle, copy, 0);
  const std::uint32_t copy_second = PivotRuntimeId(wb.handle, copy, 1);
  for (const std::uint32_t id : {copy_first, copy_second}) {
    EXPECT_NE(id, first_id);
    EXPECT_NE(id, second_id);
  }
  EXPECT_NE(copy_first, copy_second);
  EXPECT_EQ(fm_sheet_resolve_pivot_runtime_id(wb.handle, copy, first_id, &index), kNotFound);

  // Removing the preceding pivot moves the second one down; the id follows it.
  ASSERT_EQ(fm_sheet_remove_pivot(wb.handle, 0, 0), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_resolve_pivot_runtime_id(wb.handle, 0, second_id, &index), 0) << fm_get_last_error_message();
  EXPECT_EQ(index, 0U);
  EXPECT_EQ(PivotName(wb.handle, 0, index), "PT2");
  EXPECT_EQ(fm_sheet_resolve_pivot_runtime_id(wb.handle, 0, first_id, &index), kNotFound);

  // Another sheet's row deletion and removal leave it alone.
  std::size_t other = 99;
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Other", &other), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_delete_rows(wb.handle, other, 0U, 2U), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_remove_sheet(wb.handle, copy), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_resolve_pivot_runtime_id(wb.handle, 0, second_id, &index), 0) << fm_get_last_error_message();
  EXPECT_EQ(PivotName(wb.handle, 0, index), "PT2");

  // A pivot added after a removal gets an id never handed out before.
  std::size_t added = 99;
  ASSERT_EQ(fm_sheet_add_pivot(wb.handle, 0, "PT3", cache_id, 20U, 0U, &added), 0) << fm_get_last_error_message();
  const std::uint32_t added_id = PivotRuntimeId(wb.handle, 0, added);
  for (const std::uint32_t id : {first_id, second_id, copy_first, copy_second}) {
    EXPECT_NE(added_id, id);
  }

  std::uint32_t out = 0;
  EXPECT_EQ(fm_sheet_get_pivot_runtime_id(wb.handle, 0, 9, &out), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_get_pivot_runtime_id(wb.handle, 0, 0, nullptr), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_resolve_pivot_runtime_id(wb.handle, 0, added_id, nullptr), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_resolve_pivot_runtime_id(wb.handle, 9, added_id, &index), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
}
