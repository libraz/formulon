// Stable C ABI PivotTable filter tests.

#include <limits>
#include <vector>

#include "formulon_c_out_index.h"
#include "formulon_c_pivot_test_helpers.h"

TEST(FormulonCApiPivot, MutateExistingPivotFilter) {
  // Build a scratch pivot — the OOXML reader populates only `custom_name`
  // from `<pivotField name="...">`, while the filter resolver matches on
  // `source_name`, so a filter applied to a loaded workbook would silently
  // no-op. Building from scratch lets us exercise the same surface against
  // a pivot whose `source_name` is set the way the resolver expects.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();

  // Baseline projection contains "South" because no filter is active.
  {
    PivotCellsGuard projected;
    ASSERT_EQ(fm_sheet_get_pivot_layout(wb.handle, 0, pivot_index, &projected.handle), 0);
    bool saw_south = false;
    for (const fm_pivot_cell_t& c : CollectCells(projected.handle)) {
      if (c.kind == FM_PIVOT_CELL_ROW_LABEL && c.value.kind == FM_VAL_TEXT && c.value.u.text != nullptr &&
          std::string_view(c.value.u.text) == "South") {
        saw_south = true;
      }
    }
    EXPECT_TRUE(saw_south);
  }

  // Add a label filter that only retains "North".
  fm_pivot_filter_spec_t spec{};
  spec.axis = FM_PIVOT_AXIS_ROW;
  spec.field_name = "Region";
  spec.type = FM_PIVOT_FILTER_LABEL_BEGINS_WITH;
  spec.value_kind = FM_PIVOT_FILTER_VALUE_TEXT;
  spec.value_text = "N";
  spec.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  ASSERT_EQ(fm_sheet_add_pivot_filter(wb.handle, 0, pivot_index, &spec, ignored_index()), 0)
      << fm_get_last_error_message();

  std::size_t filter_count = 0;
  ASSERT_EQ(fm_sheet_get_pivot_filter_count(wb.handle, 0, pivot_index, &filter_count), 0);
  EXPECT_EQ(filter_count, 1U);

  // Re-project — "South" must be dropped.
  PivotCellsGuard projected;
  ASSERT_EQ(fm_sheet_get_pivot_layout(wb.handle, 0, pivot_index, &projected.handle), 0);
  bool saw_south = false;
  bool saw_north = false;
  for (const fm_pivot_cell_t& c : CollectCells(projected.handle)) {
    if (c.kind == FM_PIVOT_CELL_ROW_LABEL && c.value.kind == FM_VAL_TEXT && c.value.u.text != nullptr) {
      if (std::string_view(c.value.u.text) == "South") {
        saw_south = true;
      }
      if (std::string_view(c.value.u.text) == "North") {
        saw_north = true;
      }
    }
  }
  EXPECT_TRUE(saw_north);
  EXPECT_FALSE(saw_south);

  // Verify remove_at clears the filter.
  ASSERT_EQ(fm_sheet_remove_pivot_filter(wb.handle, 0, pivot_index, 0), 0);
  std::size_t after_count = 99;
  ASSERT_EQ(fm_sheet_get_pivot_filter_count(wb.handle, 0, pivot_index, &after_count), 0);
  EXPECT_EQ(after_count, 0U);
}

TEST(FormulonCApiPivot, InspectPivotFiltersByIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();

  fm_pivot_filter_spec_t int_filter{};
  int_filter.axis = FM_PIVOT_AXIS_VALUE;
  int_filter.field_name = "Amount";
  int_filter.type = FM_PIVOT_FILTER_VALUE_TOP_10;
  int_filter.value_kind = FM_PIVOT_FILTER_VALUE_INT;
  int_filter.value_int = 3;
  int_filter.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  int_filter.data_field_index = 0;
  ASSERT_EQ(fm_sheet_add_pivot_filter(wb.handle, 0, pivot_index, &int_filter, ignored_index()), 0)
      << fm_get_last_error_message();

  fm_pivot_filter_spec_t double_filter{};
  double_filter.axis = FM_PIVOT_AXIS_VALUE;
  double_filter.field_name = "Amount";
  double_filter.type = FM_PIVOT_FILTER_VALUE_GREATER_THAN;
  double_filter.value_kind = FM_PIVOT_FILTER_VALUE_DOUBLE;
  double_filter.value_double = 50.5;
  double_filter.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  double_filter.data_field_index = 0;
  ASSERT_EQ(fm_sheet_add_pivot_filter(wb.handle, 0, pivot_index, &double_filter, ignored_index()), 0)
      << fm_get_last_error_message();

  fm_pivot_filter_spec_t text_filter{};
  text_filter.axis = FM_PIVOT_AXIS_ROW;
  text_filter.field_name = "Region";
  text_filter.type = FM_PIVOT_FILTER_LABEL_CONTAINS;
  text_filter.value_kind = FM_PIVOT_FILTER_VALUE_TEXT;
  text_filter.value_text = "o";
  text_filter.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  ASSERT_EQ(fm_sheet_add_pivot_filter(wb.handle, 0, pivot_index, &text_filter, ignored_index()), 0)
      << fm_get_last_error_message();

  fm_pivot_filter_spec_t range_int_double{};
  range_int_double.axis = FM_PIVOT_AXIS_VALUE;
  range_int_double.field_name = "Amount";
  range_int_double.type = FM_PIVOT_FILTER_VALUE_BETWEEN;
  range_int_double.value_kind = FM_PIVOT_FILTER_VALUE_INT;
  range_int_double.value_int = 100;
  range_int_double.value_high_kind = FM_PIVOT_FILTER_VALUE_DOUBLE;
  range_int_double.value_high_double = 500.25;
  range_int_double.data_field_index = 0;
  ASSERT_EQ(fm_sheet_add_pivot_filter(wb.handle, 0, pivot_index, &range_int_double, ignored_index()), 0)
      << fm_get_last_error_message();

  fm_pivot_filter_spec_t range_double_int{};
  range_double_int.axis = FM_PIVOT_AXIS_VALUE;
  range_double_int.field_name = "Amount";
  range_double_int.type = FM_PIVOT_FILTER_VALUE_BETWEEN;
  range_double_int.value_kind = FM_PIVOT_FILTER_VALUE_DOUBLE;
  range_double_int.value_double = 99.75;
  range_double_int.value_high_kind = FM_PIVOT_FILTER_VALUE_INT;
  range_double_int.value_high_int = 450;
  range_double_int.data_field_index = 0;
  ASSERT_EQ(fm_sheet_add_pivot_filter(wb.handle, 0, pivot_index, &range_double_int, ignored_index()), 0)
      << fm_get_last_error_message();

  std::size_t filter_count = 0;
  ASSERT_EQ(fm_sheet_get_pivot_filter_count(wb.handle, 0, pivot_index, &filter_count), 0);
  ASSERT_EQ(filter_count, 5U);

  // Establish a projected result before inspection; a read must not change
  // either the active-filter count or the memoised evaluation state.
  PivotCellsGuard before_projection;
  ASSERT_EQ(fm_sheet_get_pivot_layout(wb.handle, 0, pivot_index, &before_projection.handle), 0)
      << fm_get_last_error_message();
  const std::vector<fm_pivot_cell_t> before_cells = CollectCells(before_projection.handle);

  fm_pivot_filter_spec_t got{};
  ASSERT_EQ(fm_sheet_get_pivot_filter(wb.handle, 0, pivot_index, 0, &got), 0);
  EXPECT_EQ(got.axis, FM_PIVOT_AXIS_VALUE);
  ASSERT_NE(got.field_name, nullptr);
  EXPECT_STREQ(got.field_name, "Amount");
  EXPECT_EQ(got.type, FM_PIVOT_FILTER_VALUE_TOP_10);
  EXPECT_EQ(got.value_kind, FM_PIVOT_FILTER_VALUE_INT);
  EXPECT_EQ(got.value_int, 3);
  EXPECT_DOUBLE_EQ(got.value_double, 0.0);
  EXPECT_EQ(got.value_text, nullptr);
  EXPECT_EQ(got.value_high_kind, FM_PIVOT_FILTER_VALUE_NONE);
  EXPECT_EQ(got.value_high_int, 0);
  EXPECT_DOUBLE_EQ(got.value_high_double, 0.0);
  EXPECT_EQ(got.data_field_index, 0U);

  got = {};
  ASSERT_EQ(fm_sheet_get_pivot_filter(wb.handle, 0, pivot_index, 1, &got), 0);
  EXPECT_EQ(got.axis, FM_PIVOT_AXIS_VALUE);
  ASSERT_NE(got.field_name, nullptr);
  EXPECT_STREQ(got.field_name, "Amount");
  EXPECT_EQ(got.type, FM_PIVOT_FILTER_VALUE_GREATER_THAN);
  EXPECT_EQ(got.value_kind, FM_PIVOT_FILTER_VALUE_DOUBLE);
  EXPECT_EQ(got.value_int, 0);
  EXPECT_DOUBLE_EQ(got.value_double, 50.5);
  EXPECT_EQ(got.value_text, nullptr);
  EXPECT_EQ(got.value_high_kind, FM_PIVOT_FILTER_VALUE_NONE);
  EXPECT_EQ(got.value_high_int, 0);
  EXPECT_DOUBLE_EQ(got.value_high_double, 0.0);
  EXPECT_EQ(got.data_field_index, 0U);

  got = {};
  ASSERT_EQ(fm_sheet_get_pivot_filter(wb.handle, 0, pivot_index, 2, &got), 0);
  EXPECT_EQ(got.axis, FM_PIVOT_AXIS_ROW);
  ASSERT_NE(got.field_name, nullptr);
  EXPECT_STREQ(got.field_name, "Region");
  EXPECT_EQ(got.type, FM_PIVOT_FILTER_LABEL_CONTAINS);
  EXPECT_EQ(got.value_kind, FM_PIVOT_FILTER_VALUE_TEXT);
  EXPECT_EQ(got.value_int, 0);
  EXPECT_DOUBLE_EQ(got.value_double, 0.0);
  ASSERT_NE(got.value_text, nullptr);
  EXPECT_STREQ(got.value_text, "o");
  EXPECT_EQ(got.value_high_kind, FM_PIVOT_FILTER_VALUE_NONE);
  EXPECT_EQ(got.value_high_int, 0);
  EXPECT_DOUBLE_EQ(got.value_high_double, 0.0);
  EXPECT_EQ(got.data_field_index, 0U);

  got = {};
  ASSERT_EQ(fm_sheet_get_pivot_filter(wb.handle, 0, pivot_index, 3, &got), 0);
  EXPECT_EQ(got.axis, FM_PIVOT_AXIS_VALUE);
  ASSERT_NE(got.field_name, nullptr);
  EXPECT_STREQ(got.field_name, "Amount");
  EXPECT_EQ(got.type, FM_PIVOT_FILTER_VALUE_BETWEEN);
  EXPECT_EQ(got.value_kind, FM_PIVOT_FILTER_VALUE_INT);
  EXPECT_EQ(got.value_int, 100);
  EXPECT_DOUBLE_EQ(got.value_double, 0.0);
  EXPECT_EQ(got.value_text, nullptr);
  EXPECT_EQ(got.value_high_kind, FM_PIVOT_FILTER_VALUE_DOUBLE);
  EXPECT_EQ(got.value_high_int, 0);
  EXPECT_DOUBLE_EQ(got.value_high_double, 500.25);
  EXPECT_EQ(got.data_field_index, 0U);

  got = {};
  ASSERT_EQ(fm_sheet_get_pivot_filter(wb.handle, 0, pivot_index, 4, &got), 0);
  EXPECT_EQ(got.axis, FM_PIVOT_AXIS_VALUE);
  ASSERT_NE(got.field_name, nullptr);
  EXPECT_STREQ(got.field_name, "Amount");
  EXPECT_EQ(got.type, FM_PIVOT_FILTER_VALUE_BETWEEN);
  EXPECT_EQ(got.value_kind, FM_PIVOT_FILTER_VALUE_DOUBLE);
  EXPECT_EQ(got.value_int, 0);
  EXPECT_DOUBLE_EQ(got.value_double, 99.75);
  EXPECT_EQ(got.value_text, nullptr);
  EXPECT_EQ(got.value_high_kind, FM_PIVOT_FILTER_VALUE_INT);
  EXPECT_EQ(got.value_high_int, 450);
  EXPECT_DOUBLE_EQ(got.value_high_double, 0.0);
  EXPECT_EQ(got.data_field_index, 0U);

  std::size_t unchanged_count = 0;
  ASSERT_EQ(fm_sheet_get_pivot_filter_count(wb.handle, 0, pivot_index, &unchanged_count), 0);
  EXPECT_EQ(unchanged_count, filter_count);
  PivotCellsGuard after_projection;
  ASSERT_EQ(fm_sheet_get_pivot_layout(wb.handle, 0, pivot_index, &after_projection.handle), 0)
      << fm_get_last_error_message();
  const std::vector<fm_pivot_cell_t> after_cells = CollectCells(after_projection.handle);
  ASSERT_EQ(after_cells.size(), before_cells.size());
  for (std::size_t i = 0; i < before_cells.size(); ++i) {
    EXPECT_EQ(after_cells[i].row, before_cells[i].row);
    EXPECT_EQ(after_cells[i].col, before_cells[i].col);
    EXPECT_EQ(after_cells[i].kind, before_cells[i].kind);
    EXPECT_EQ(after_cells[i].value.kind, before_cells[i].value.kind);
    if (before_cells[i].value.kind == FM_VAL_NUMBER) {
      EXPECT_DOUBLE_EQ(after_cells[i].value.u.number, before_cells[i].value.u.number);
    } else if (before_cells[i].value.kind == FM_VAL_TEXT) {
      ASSERT_NE(before_cells[i].value.u.text, nullptr);
      ASSERT_NE(after_cells[i].value.u.text, nullptr);
      EXPECT_STREQ(after_cells[i].value.u.text, before_cells[i].value.u.text);
    }
  }

  // Adding a filter invalidates the old model-backed views. Do not inspect
  // those pointers after the mutation; reacquire the entry instead.
  fm_pivot_filter_spec_t mutation{};
  mutation.axis = FM_PIVOT_AXIS_ROW;
  mutation.field_name = "Region";
  mutation.type = FM_PIVOT_FILTER_LABEL_BEGINS_WITH;
  mutation.value_kind = FM_PIVOT_FILTER_VALUE_TEXT;
  mutation.value_text = "N";
  mutation.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  ASSERT_EQ(fm_sheet_add_pivot_filter(wb.handle, 0, pivot_index, &mutation, ignored_index()), 0)
      << fm_get_last_error_message();
  fm_pivot_filter_spec_t reacquired{};
  ASSERT_EQ(fm_sheet_get_pivot_filter(wb.handle, 0, pivot_index, 2, &reacquired), 0);
  EXPECT_EQ(reacquired.value_kind, FM_PIVOT_FILTER_VALUE_TEXT);
  ASSERT_NE(reacquired.value_text, nullptr);
  EXPECT_STREQ(reacquired.value_text, "o");

  const auto make_sentinel = [] {
    fm_pivot_filter_spec_t sentinel{};
    sentinel.axis = FM_PIVOT_AXIS_PAGE;
    sentinel.field_name = "sentinel-field";
    sentinel.type = FM_PIVOT_FILTER_LABEL_DATE;
    sentinel.value_kind = FM_PIVOT_FILTER_VALUE_DOUBLE;
    sentinel.value_int = -123456789;
    sentinel.value_double = -9876.5;
    sentinel.value_text = "sentinel-text";
    sentinel.value_high_kind = FM_PIVOT_FILTER_VALUE_INT;
    sentinel.value_high_int = 13579;
    sentinel.value_high_double = 24680.5;
    sentinel.data_field_index = 0xDEADBEEFU;
    sentinel.value_high_text = "sentinel-high";
    return sentinel;
  };
  const auto expect_unchanged = [](const fm_pivot_filter_spec_t& expected, const fm_pivot_filter_spec_t& actual) {
    EXPECT_EQ(actual.axis, expected.axis);
    EXPECT_EQ(actual.field_name, expected.field_name);
    EXPECT_EQ(actual.type, expected.type);
    EXPECT_EQ(actual.value_kind, expected.value_kind);
    EXPECT_EQ(actual.value_int, expected.value_int);
    EXPECT_DOUBLE_EQ(actual.value_double, expected.value_double);
    EXPECT_EQ(actual.value_text, expected.value_text);
    EXPECT_EQ(actual.value_high_kind, expected.value_high_kind);
    EXPECT_EQ(actual.value_high_int, expected.value_high_int);
    EXPECT_DOUBLE_EQ(actual.value_high_double, expected.value_high_double);
    EXPECT_EQ(actual.data_field_index, expected.data_field_index);
    EXPECT_EQ(actual.value_high_text, expected.value_high_text);
  };
  const auto expect_error_preserves_output = [&](const auto& invoke, fm_status_t expected_status) {
    fm_pivot_filter_spec_t invalid_out = make_sentinel();
    const fm_pivot_filter_spec_t before = invalid_out;
    EXPECT_EQ(invoke(&invalid_out), expected_status);
    expect_unchanged(before, invalid_out);
  };
  expect_error_preserves_output(
      [&](fm_pivot_filter_spec_t* out) { return fm_sheet_get_pivot_filter(wb.handle, 0, pivot_index, 6, out); },
      FM_STATUS_INDEX_OUT_OF_RANGE);
  expect_error_preserves_output(
      [&](fm_pivot_filter_spec_t* out) { return fm_sheet_get_pivot_filter(wb.handle, 1, pivot_index, 0, out); },
      FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  expect_error_preserves_output(
      [&](fm_pivot_filter_spec_t* out) { return fm_sheet_get_pivot_filter(wb.handle, 0, pivot_index + 1, 0, out); },
      FM_STATUS_INDEX_OUT_OF_RANGE);
  expect_error_preserves_output(
      [&](fm_pivot_filter_spec_t* out) { return fm_sheet_get_pivot_filter(nullptr, 0, pivot_index, 0, out); },
      static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_sheet_get_pivot_filter(wb.handle, 0, pivot_index, 0, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
}

TEST(FormulonCApiPivot, GetterRejectsUnrepresentableModelEnumsWithoutMutation) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();

  auto* table = wb.handle->workbook().sheet(0).mutable_pivot_tables()[pivot_index].get();
  ASSERT_NE(table, nullptr);
  const auto invalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  const auto make_sentinel = [] {
    fm_pivot_filter_spec_t sentinel{};
    sentinel.axis = FM_PIVOT_AXIS_PAGE;
    sentinel.field_name = "sentinel-field";
    sentinel.type = FM_PIVOT_FILTER_LABEL_DATE;
    sentinel.value_kind = FM_PIVOT_FILTER_VALUE_DOUBLE;
    sentinel.value_int = -123456789;
    sentinel.value_double = -9876.5;
    sentinel.value_text = "sentinel-text";
    sentinel.value_high_kind = FM_PIVOT_FILTER_VALUE_INT;
    sentinel.value_high_int = 13579;
    sentinel.value_high_double = 24680.5;
    sentinel.data_field_index = 0xDEADBEEFU;
    return sentinel;
  };
  const auto inject_and_expect_rejected = [&](formulon::pivot::PivotAxis axis, formulon::pivot::FilterType type) {
    table->mutable_active_filters().clear();
    formulon::pivot::PivotFilter filter;
    filter.axis = axis;
    filter.field_name = "Region";
    filter.type = type;
    filter.value = std::string("N");
    table->mutable_active_filters().push_back(filter);

    fm_pivot_filter_spec_t out = make_sentinel();
    const fm_pivot_filter_spec_t before = out;
    EXPECT_EQ(fm_sheet_get_pivot_filter(wb.handle, 0, pivot_index, 0, &out), invalid);
    EXPECT_EQ(std::memcmp(&out, &before, sizeof(out)), 0);
  };

  inject_and_expect_rejected(formulon::pivot::PivotAxis::None, formulon::pivot::FilterType::LabelContains);
  inject_and_expect_rejected(formulon::pivot::PivotAxis::Row, static_cast<formulon::pivot::FilterType>(0xFF));
}

TEST(FormulonCApiPivot, PivotFilterAddRejectsRawAxisAndTypeWithoutMutation) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();

  fm_pivot_filter_spec_t spec{};
  spec.axis = RawEnumValue<fm_pivot_axis_t>(99);
  spec.field_name = "Region";
  spec.type = FM_PIVOT_FILTER_LABEL_BEGINS_WITH;
  spec.value_kind = FM_PIVOT_FILTER_VALUE_TEXT;
  spec.value_text = "N";
  spec.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  const auto invalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(fm_sheet_add_pivot_filter(wb.handle, 0, pivot_index, &spec, ignored_index()), invalid);

  std::size_t count = 99;
  ASSERT_EQ(fm_sheet_get_pivot_filter_count(wb.handle, 0, pivot_index, &count), 0);
  EXPECT_EQ(count, 0U);

  spec.axis = FM_PIVOT_AXIS_ROW;
  spec.type = RawEnumValue<fm_pivot_filter_type_t>(99);
  EXPECT_EQ(fm_sheet_add_pivot_filter(wb.handle, 0, pivot_index, &spec, ignored_index()), invalid);
  ASSERT_EQ(fm_sheet_get_pivot_filter_count(wb.handle, 0, pivot_index, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST(FormulonCApiPivot, PivotFilterRejectsInvalidDataFieldWithoutMutation) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();

  fm_pivot_filter_spec_t spec{};
  spec.axis = FM_PIVOT_AXIS_ROW;
  spec.field_name = "Region";
  spec.type = FM_PIVOT_FILTER_VALUE_TOP_10;
  spec.value_kind = FM_PIVOT_FILTER_VALUE_INT;
  spec.value_int = 1;
  spec.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  spec.data_field_index = 1;  // BuildScratchPivot has only slot 0.
  EXPECT_EQ(fm_sheet_add_pivot_filter(wb.handle, 0, pivot_index, &spec, ignored_index()), FM_STATUS_INDEX_OUT_OF_RANGE);

  std::size_t count = 99;
  ASSERT_EQ(fm_sheet_get_pivot_filter_count(wb.handle, 0, pivot_index, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST(FormulonCApiPivot, PivotFilterRejectsSelectorWhenNoDataFields) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_clear_pivot_data_fields(wb.handle, 0, pivot_index), 0) << fm_get_last_error_message();

  fm_pivot_filter_spec_t spec{};
  spec.axis = FM_PIVOT_AXIS_ROW;
  spec.field_name = "Region";
  spec.type = FM_PIVOT_FILTER_VALUE_TOP_10;
  spec.value_kind = FM_PIVOT_FILTER_VALUE_INT;
  spec.value_int = 1;
  spec.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  spec.data_field_index = 0;
  EXPECT_EQ(fm_sheet_add_pivot_filter(wb.handle, 0, pivot_index, &spec, ignored_index()), FM_STATUS_INDEX_OUT_OF_RANGE);

  std::size_t count = 99;
  ASSERT_EQ(fm_sheet_get_pivot_filter_count(wb.handle, 0, pivot_index, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST(FormulonCApiPivot, PivotFilterRejectsUnknownFieldNameWithoutMutation) {
  // A label filter naming no pivot field would be a silent no-op inside the
  // engine, so the mutator rejects it instead of reporting success.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();

  fm_pivot_filter_spec_t spec{};
  spec.axis = FM_PIVOT_AXIS_ROW;
  spec.field_name = "NoSuchField";
  spec.type = FM_PIVOT_FILTER_LABEL_BEGINS_WITH;
  spec.value_kind = FM_PIVOT_FILTER_VALUE_TEXT;
  spec.value_text = "N";
  spec.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  const auto invalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(fm_sheet_add_pivot_filter(wb.handle, 0, pivot_index, &spec, ignored_index()), invalid);

  std::size_t count = 99;
  ASSERT_EQ(fm_sheet_get_pivot_filter_count(wb.handle, 0, pivot_index, &count), 0);
  EXPECT_EQ(count, 0U);

  // The data field's display name resolves through to its source field, so
  // the same call shape is accepted for it.
  spec.field_name = "Sum of Amount";
  ASSERT_EQ(fm_sheet_add_pivot_filter(wb.handle, 0, pivot_index, &spec, ignored_index()), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_pivot_filter_count(wb.handle, 0, pivot_index, &count), 0);
  EXPECT_EQ(count, 1U);
}

TEST(FormulonCApiPivot, ValueFilterRejectsUnknownFieldNameWithoutMutation) {
  // The engine skips a filter whose field does not resolve, so a value filter
  // naming no pivot field must be rejected rather than accepted as a no-op.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();

  fm_pivot_filter_spec_t spec{};
  spec.axis = FM_PIVOT_AXIS_ROW;
  spec.field_name = "NoSuchField";
  spec.type = FM_PIVOT_FILTER_VALUE_TOP_10;
  spec.value_kind = FM_PIVOT_FILTER_VALUE_INT;
  spec.value_int = 1;
  spec.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  spec.data_field_index = 0;
  const auto invalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(fm_sheet_add_pivot_filter(wb.handle, 0, pivot_index, &spec, ignored_index()), invalid);

  std::size_t count = 99;
  ASSERT_EQ(fm_sheet_get_pivot_filter_count(wb.handle, 0, pivot_index, &count), 0);
  EXPECT_EQ(count, 0U);

  spec.field_name = "Region";
  ASSERT_EQ(fm_sheet_add_pivot_filter(wb.handle, 0, pivot_index, &spec, ignored_index()), 0)
      << fm_get_last_error_message();
}

TEST(FormulonCApiPivot, FilterRejectsAmbiguousFieldNameWithoutMutation) {
  // A caption shared by two distinct fields names neither of them, exactly
  // like an unknown name.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();

  fm_pivot_field_spec_t dup_spec{};
  dup_spec.source_name = "Amount";
  dup_spec.custom_name = "Region";
  dup_spec.axis = FM_PIVOT_AXIS_VALUE;
  dup_spec.number_format = "";
  std::size_t dup_field = 99;
  ASSERT_EQ(fm_sheet_add_pivot_field(wb.handle, 0, pivot_index, &dup_spec, &dup_field), 0)
      << fm_get_last_error_message();

  fm_pivot_filter_spec_t spec{};
  spec.axis = FM_PIVOT_AXIS_ROW;
  spec.field_name = "Region";
  spec.type = FM_PIVOT_FILTER_LABEL_BEGINS_WITH;
  spec.value_kind = FM_PIVOT_FILTER_VALUE_TEXT;
  spec.value_text = "N";
  spec.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  const auto invalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(fm_sheet_add_pivot_filter(wb.handle, 0, pivot_index, &spec, ignored_index()), invalid);

  std::size_t count = 99;
  ASSERT_EQ(fm_sheet_get_pivot_filter_count(wb.handle, 0, pivot_index, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST(FormulonCApiPivot, FilterErrorUsesApiName) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();

  fm_pivot_filter_spec_t spec{};
  spec.axis = FM_PIVOT_AXIS_ROW;
  spec.field_name = "Region";
  spec.type = FM_PIVOT_FILTER_LABEL_BEGINS_WITH;
  spec.value_kind = FM_PIVOT_FILTER_VALUE_NONE;
  spec.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  const auto invalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(fm_sheet_add_pivot_filter(wb.handle, 0, pivot_index, &spec, ignored_index()), invalid);
  EXPECT_STRNE(fm_get_last_error_message(), "");
}

// The spec of a well-formed filter of `type` on the scratch pivot's Region
// field: text for a label type, a number for a value or date type, nothing
// for a relative period or calendar position.
fm_pivot_filter_spec_t FilterSpecFor(int type) {
  fm_pivot_filter_spec_t spec{};
  spec.axis = FM_PIVOT_AXIS_ROW;
  spec.field_name = "Region";
  spec.type = static_cast<fm_pivot_filter_type_t>(type);
  spec.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  const bool label = type == FM_PIVOT_FILTER_LABEL_CONTAINS || type == FM_PIVOT_FILTER_LABEL_BEGINS_WITH ||
                     (type >= FM_PIVOT_FILTER_LABEL_EQUAL && type <= FM_PIVOT_FILTER_LABEL_NOT_BETWEEN);
  const bool numeric_range = type == FM_PIVOT_FILTER_VALUE_BETWEEN || type == FM_PIVOT_FILTER_VALUE_NOT_BETWEEN ||
                             type == FM_PIVOT_FILTER_LABEL_DATE || type == FM_PIVOT_FILTER_DATE_NOT_BETWEEN;
  const bool amount = type == FM_PIVOT_FILTER_VALUE_TOP_10 ||
                      (type >= FM_PIVOT_FILTER_VALUE_BOTTOM_COUNT && type <= FM_PIVOT_FILTER_VALUE_BOTTOM_SUM);
  if (type >= FM_PIVOT_FILTER_DATE_TOMORROW) {
    spec.value_kind = FM_PIVOT_FILTER_VALUE_NONE;
  } else if (label) {
    spec.value_kind = FM_PIVOT_FILTER_VALUE_TEXT;
    spec.value_text = "N";
    if (type == FM_PIVOT_FILTER_LABEL_BETWEEN || type == FM_PIVOT_FILTER_LABEL_NOT_BETWEEN) {
      spec.value_high_kind = FM_PIVOT_FILTER_VALUE_TEXT;
      spec.value_high_text = "S";
    }
  } else if (amount) {
    spec.value_kind = FM_PIVOT_FILTER_VALUE_INT;
    spec.value_int = 2;
  } else {
    spec.value_kind = FM_PIVOT_FILTER_VALUE_DOUBLE;
    spec.value_double = 45000.5;
    if (numeric_range) {
      spec.value_high_kind = FM_PIVOT_FILTER_VALUE_DOUBLE;
      spec.value_high_double = 45100.25;
    }
  }
  return spec;
}

// Every type is accepted, saved, and listed again by `filter_at` after a
// reload, with the payload it was added with.
TEST(FormulonCApiPivot, EveryFilterTypeRoundTripsThroughSaveAndLoad) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();
  for (int type = 0; type <= FM_PIVOT_FILTER_DATE_M12; ++type) {
    const fm_pivot_filter_spec_t spec = FilterSpecFor(type);
    ASSERT_EQ(fm_sheet_add_pivot_filter(wb.handle, 0, pivot_index, &spec, ignored_index()), 0)
        << "type=" << type << " " << fm_get_last_error_message();
  }

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0) << fm_get_last_error_message();
  WorkbookGuard back;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &back.handle), 0) << fm_get_last_error_message();
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_pivot_filter_count(back.handle, 0, 0, &count), 0);
  ASSERT_EQ(count, static_cast<std::size_t>(FM_PIVOT_FILTER_DATE_M12) + 1U);
  for (int type = 0; type <= FM_PIVOT_FILTER_DATE_M12; ++type) {
    SCOPED_TRACE(type);
    const fm_pivot_filter_spec_t want = FilterSpecFor(type);
    fm_pivot_filter_spec_t got{};
    ASSERT_EQ(fm_sheet_get_pivot_filter(back.handle, 0, 0, static_cast<std::size_t>(type), &got), 0)
        << fm_get_last_error_message();
    EXPECT_EQ(got.type, want.type);
    EXPECT_EQ(got.axis, FM_PIVOT_AXIS_ROW);
    ASSERT_NE(got.field_name, nullptr);
    EXPECT_STREQ(got.field_name, "Region");
    EXPECT_EQ(got.value_kind, want.value_kind);
    EXPECT_EQ(got.value_int, want.value_int);
    EXPECT_DOUBLE_EQ(got.value_double, want.value_double);
    if (want.value_text != nullptr) {
      ASSERT_NE(got.value_text, nullptr);
      EXPECT_STREQ(got.value_text, want.value_text);
    } else {
      EXPECT_EQ(got.value_text, nullptr);
    }
    EXPECT_EQ(got.value_high_kind, want.value_high_kind);
    EXPECT_DOUBLE_EQ(got.value_high_double, want.value_high_double);
    if (want.value_high_text != nullptr) {
      ASSERT_NE(got.value_high_text, nullptr);
      EXPECT_STREQ(got.value_high_text, want.value_high_text);
    } else {
      EXPECT_EQ(got.value_high_text, nullptr);
    }
  }
}

// Payloads that do not fit the type are refused and leave the list alone.
TEST(FormulonCApiPivot, FilterAddRejectsPayloadsThatDoNotFitTheType) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();
  const double nan = std::numeric_limits<double>::quiet_NaN();
  const double inf = std::numeric_limits<double>::infinity();

  struct Case {
    const char* label;
    fm_pivot_filter_spec_t spec;
    bool names_collection_element = false;
  };
  std::vector<Case> cases;
  const auto with = [&](const char* label, int type, auto edit) {
    fm_pivot_filter_spec_t spec = FilterSpecFor(type);
    edit(spec);
    cases.push_back(Case{label, spec});
  };
  with("operandless with a value", FM_PIVOT_FILTER_DATE_TODAY, [](fm_pivot_filter_spec_t& s) {
    s.value_kind = FM_PIVOT_FILTER_VALUE_INT;
    s.value_int = 1;
  });
  with("label with a number", FM_PIVOT_FILTER_LABEL_EQUAL, [](fm_pivot_filter_spec_t& s) {
    s.value_kind = FM_PIVOT_FILTER_VALUE_DOUBLE;
    s.value_double = 1.0;
  });
  with("value with text", FM_PIVOT_FILTER_VALUE_EQUAL, [](fm_pivot_filter_spec_t& s) {
    s.value_kind = FM_PIVOT_FILTER_VALUE_TEXT;
    s.value_text = "1";
  });
  with("date with text", FM_PIVOT_FILTER_DATE_EQUAL, [](fm_pivot_filter_spec_t& s) {
    s.value_kind = FM_PIVOT_FILTER_VALUE_TEXT;
    s.value_text = "2024-01-01";
  });
  with("value without a value", FM_PIVOT_FILTER_VALUE_EQUAL,
       [](fm_pivot_filter_spec_t& s) { s.value_kind = FM_PIVOT_FILTER_VALUE_NONE; });
  with("label range without its upper text", FM_PIVOT_FILTER_LABEL_BETWEEN,
       [](fm_pivot_filter_spec_t& s) { s.value_high_text = nullptr; });
  with("label range with a numeric upper bound", FM_PIVOT_FILTER_LABEL_NOT_BETWEEN, [](fm_pivot_filter_spec_t& s) {
    s.value_high_kind = FM_PIVOT_FILTER_VALUE_DOUBLE;
    s.value_high_double = 2.0;
  });
  with("upper text on another type", FM_PIVOT_FILTER_LABEL_EQUAL,
       [](fm_pivot_filter_spec_t& s) { s.value_high_text = "Z"; });
  with("NaN", FM_PIVOT_FILTER_VALUE_GREATER_THAN, [&](fm_pivot_filter_spec_t& s) { s.value_double = nan; });
  with("infinite upper bound", FM_PIVOT_FILTER_VALUE_BETWEEN,
       [&](fm_pivot_filter_spec_t& s) { s.value_high_double = inf; });
  with("percent above 100", FM_PIVOT_FILTER_VALUE_TOP_PERCENT, [](fm_pivot_filter_spec_t& s) { s.value_int = 101; });
  with("percent below 0", FM_PIVOT_FILTER_VALUE_BOTTOM_PERCENT, [](fm_pivot_filter_spec_t& s) {
    s.value_kind = FM_PIVOT_FILTER_VALUE_DOUBLE;
    s.value_double = -0.5;
  });
  with("count of 0", FM_PIVOT_FILTER_VALUE_TOP_10, [](fm_pivot_filter_spec_t& s) { s.value_int = 0; });
  with("negative sum", FM_PIVOT_FILTER_VALUE_BOTTOM_SUM, [](fm_pivot_filter_spec_t& s) { s.value_int = -5; });
  with("data field out of range", FM_PIVOT_FILTER_VALUE_BOTTOM_COUNT,
       [](fm_pivot_filter_spec_t& s) { s.data_field_index = 3; });
  cases.back().names_collection_element = true;
  with("type past the last", FM_PIVOT_FILTER_DATE_M12,
       [](fm_pivot_filter_spec_t& s) { s.type = RawEnumValue<fm_pivot_filter_type_t>(68); });

  const auto invalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  for (const Case& c : cases) {
    SCOPED_TRACE(c.label);
    EXPECT_EQ(fm_sheet_add_pivot_filter(wb.handle, 0, pivot_index, &c.spec, ignored_index()),
              c.names_collection_element ? FM_STATUS_INDEX_OUT_OF_RANGE : invalid);
    std::size_t count = 99;
    ASSERT_EQ(fm_sheet_get_pivot_filter_count(wb.handle, 0, pivot_index, &count), 0);
    EXPECT_EQ(count, 0U);
  }

  // The percent bounds are inclusive.
  fm_pivot_filter_spec_t edge = FilterSpecFor(FM_PIVOT_FILTER_VALUE_TOP_PERCENT);
  edge.value_int = 100;
  EXPECT_EQ(fm_sheet_add_pivot_filter(wb.handle, 0, pivot_index, &edge, ignored_index()), 0)
      << fm_get_last_error_message();
  edge.value_int = 0;
  EXPECT_EQ(fm_sheet_add_pivot_filter(wb.handle, 0, pivot_index, &edge, ignored_index()), 0)
      << fm_get_last_error_message();
}

// A label filter's wildcard operand reaches the evaluation as entered.
TEST(FormulonCApiPivot, LabelFiltersEvaluateTheirOperands) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();

  fm_pivot_filter_spec_t spec = FilterSpecFor(FM_PIVOT_FILTER_LABEL_NOT_EQUAL);
  spec.value_text = "n*";
  ASSERT_EQ(fm_sheet_add_pivot_filter(wb.handle, 0, pivot_index, &spec, ignored_index()), 0)
      << fm_get_last_error_message();
  EXPECT_FALSE(LayoutHasText(wb.handle, pivot_index, "North"));
  EXPECT_TRUE(LayoutHasText(wb.handle, pivot_index, "South"));

  ASSERT_EQ(fm_sheet_clear_pivot_filters(wb.handle, 0, pivot_index), 0);
  spec = FilterSpecFor(FM_PIVOT_FILTER_LABEL_BETWEEN);
  spec.value_text = "A";
  spec.value_high_text = "O";
  ASSERT_EQ(fm_sheet_add_pivot_filter(wb.handle, 0, pivot_index, &spec, ignored_index()), 0)
      << fm_get_last_error_message();
  EXPECT_TRUE(LayoutHasText(wb.handle, pivot_index, "North"));
  EXPECT_FALSE(LayoutHasText(wb.handle, pivot_index, "South"));
}

// Taking a data field's source off the report drops the value filters that
// score it and re-points the others, for every value type.
TEST(FormulonCApiPivot, RemovingADataFieldRepointsOrDropsEveryValueFilterType) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  ASSERT_EQ(fm_workbook_add_pivot_cache(wb.handle, 0, &cache_id), 0);
  ASSERT_EQ(fm_workbook_set_pivot_cache_worksheet_source(wb.handle, cache_id, 1, "A1:C4", "Sheet1", nullptr), 0);
  for (const char* name : {"Region", "Sales", "Qty"}) {
    std::size_t idx = 0;
    ASSERT_EQ(fm_workbook_add_pivot_cache_field(wb.handle, cache_id, name, &idx), 0);
  }
  ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_text(wb.handle, cache_id, 0, "North"), 0);
  ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_text(wb.handle, cache_id, 0, "South"), 0);
  const double rows[][3] = {{0, 100, 1}, {1, 200, 2}, {0, 300, 3}};
  for (const auto& row : rows) {
    std::size_t rec = 0;
    ASSERT_EQ(fm_workbook_add_pivot_cache_record(wb.handle, cache_id, &rec), 0);
    for (std::size_t col = 0; col < 3; ++col) {
      ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache_id, rec, col, row[col]), 0);
    }
  }
  std::size_t pivot = 0;
  ASSERT_EQ(fm_sheet_add_pivot(wb.handle, 0, "PT", cache_id, 0, 6, &pivot), 0) << fm_get_last_error_message();
  const struct {
    const char* name;
    fm_pivot_axis_t axis;
  } fields[] = {{"Region", FM_PIVOT_AXIS_ROW}, {"Sales", FM_PIVOT_AXIS_VALUE}, {"Qty", FM_PIVOT_AXIS_VALUE}};
  for (const auto& field : fields) {
    fm_pivot_field_spec_t spec{};
    spec.source_name = field.name;
    spec.axis = field.axis;
    spec.subtotal_top = 1;
    std::size_t idx = 0;
    ASSERT_EQ(fm_sheet_add_pivot_field(wb.handle, 0, pivot, &spec, &idx), 0) << fm_get_last_error_message();
  }
  for (const std::uint32_t field : {1U, 2U}) {
    fm_pivot_data_field_spec_t df{};
    df.field_index = field;
    df.aggregation = FM_PIVOT_AGG_SUM;
    df.show_as_base_field = -1;
    df.show_as_base_item = -1;
    std::size_t idx = 0;
    ASSERT_EQ(fm_sheet_add_pivot_data_field(wb.handle, 0, pivot, &df, &idx), 0) << fm_get_last_error_message();
  }
  // Top sum scores Sales (data field 0); the range scores Qty (data field 1).
  fm_pivot_filter_spec_t top_sum = FilterSpecFor(FM_PIVOT_FILTER_VALUE_TOP_SUM);
  top_sum.data_field_index = 0;
  ASSERT_EQ(fm_sheet_add_pivot_filter(wb.handle, 0, pivot, &top_sum, ignored_index()), 0)
      << fm_get_last_error_message();
  fm_pivot_filter_spec_t range = FilterSpecFor(FM_PIVOT_FILTER_VALUE_NOT_BETWEEN);
  range.data_field_index = 1;
  ASSERT_EQ(fm_sheet_add_pivot_filter(wb.handle, 0, pivot, &range, ignored_index()), 0) << fm_get_last_error_message();

  ASSERT_EQ(fm_sheet_set_pivot_field_axis(wb.handle, 0, pivot, 1, -1), 0) << fm_get_last_error_message();
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_pivot_filter_count(wb.handle, 0, pivot, &count), 0);
  ASSERT_EQ(count, 1U);
  fm_pivot_filter_spec_t kept{};
  ASSERT_EQ(fm_sheet_get_pivot_filter(wb.handle, 0, pivot, 0, &kept), 0);
  EXPECT_EQ(kept.type, FM_PIVOT_FILTER_VALUE_NOT_BETWEEN);
  EXPECT_EQ(kept.data_field_index, 0U);
}

TEST(FormulonCApiPivot, MutateShowValuesAs) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();

  // Re-set the data field to PercentOfRow.
  std::size_t df_count = 0;
  ASSERT_EQ(fm_sheet_get_pivot_data_field_count(wb.handle, 0, pivot_index, &df_count), 0);
  ASSERT_EQ(df_count, 1U);

  fm_pivot_data_field_spec_t df_spec{};
  df_spec.name = "Sum of Amount";
  df_spec.field_index = 1U;  // Amount is the second pivot field.
  df_spec.aggregation = FM_PIVOT_AGG_SUM;
  df_spec.number_format = "";
  df_spec.show_as = FM_PIVOT_SHOW_VALUES_AS_PERCENT_OF_ROW;
  df_spec.show_as_base_field = -1;
  df_spec.show_as_base_item = -1;
  ASSERT_EQ(fm_sheet_set_pivot_data_field(wb.handle, 0, pivot_index, 0, &df_spec), 0) << fm_get_last_error_message();

  PivotCellsGuard projected;
  ASSERT_EQ(fm_sheet_get_pivot_layout(wb.handle, 0, pivot_index, &projected.handle), 0) << fm_get_last_error_message();

  // Each FM_PIVOT_CELL_DATA cell should equal 1.0 (single data column ->
  // each row's percent-of-row is 100% of itself).
  std::size_t data_count = 0;
  for (const fm_pivot_cell_t& c : CollectCells(projected.handle)) {
    if (c.kind == FM_PIVOT_CELL_DATA && c.value.kind == FM_VAL_NUMBER) {
      EXPECT_DOUBLE_EQ(c.value.u.number, 1.0);
      ++data_count;
    }
  }
  EXPECT_GT(data_count, 0U);
}
