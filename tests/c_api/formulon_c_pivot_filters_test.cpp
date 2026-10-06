// Stable C ABI PivotTable filter tests.

#include "formulon_c_pivot_test_helpers.h"

TEST(FormulonCApiPivot, MutateExistingPivotFilter) {
  // Build a scratch pivot — the OOXML reader populates only `custom_name`
  // from `<pivotField name="...">`, while the filter resolver matches on
  // `source_name`, so a filter applied to a loaded workbook would silently
  // no-op. Building from scratch lets us exercise the same surface against
  // a pivot whose `source_name` is set the way the resolver expects.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  // Baseline projection contains "South" because no filter is active.
  {
    PivotCellsGuard projected;
    ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0);
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
  ASSERT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, pivot_idx, &spec), 0) << fm_last_error_message();

  std::size_t filter_count = 0;
  ASSERT_EQ(fm_workbook_pivot_filter_count(wb.handle, 0, pivot_idx, &filter_count), 0);
  EXPECT_EQ(filter_count, 1U);

  // Re-project — "South" must be dropped.
  PivotCellsGuard projected;
  ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0);
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
  ASSERT_EQ(fm_workbook_pivot_filter_remove_at(wb.handle, 0, pivot_idx, 0), 0);
  std::size_t after_count = 99;
  ASSERT_EQ(fm_workbook_pivot_filter_count(wb.handle, 0, pivot_idx, &after_count), 0);
  EXPECT_EQ(after_count, 0U);
}

TEST(FormulonCApiPivot, InspectPivotFiltersByIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  fm_pivot_filter_spec_t int_filter{};
  int_filter.axis = FM_PIVOT_AXIS_VALUE;
  int_filter.field_name = "Amount";
  int_filter.type = FM_PIVOT_FILTER_VALUE_TOP_10;
  int_filter.value_kind = FM_PIVOT_FILTER_VALUE_INT;
  int_filter.value_int = 3;
  int_filter.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  int_filter.data_field_index = 0;
  ASSERT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, pivot_idx, &int_filter), 0) << fm_last_error_message();

  fm_pivot_filter_spec_t double_filter{};
  double_filter.axis = FM_PIVOT_AXIS_VALUE;
  double_filter.field_name = "Amount";
  double_filter.type = FM_PIVOT_FILTER_VALUE_GREATER_THAN;
  double_filter.value_kind = FM_PIVOT_FILTER_VALUE_DOUBLE;
  double_filter.value_double = 50.5;
  double_filter.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  double_filter.data_field_index = 0;
  ASSERT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, pivot_idx, &double_filter), 0) << fm_last_error_message();

  fm_pivot_filter_spec_t text_filter{};
  text_filter.axis = FM_PIVOT_AXIS_ROW;
  text_filter.field_name = "Region";
  text_filter.type = FM_PIVOT_FILTER_LABEL_CONTAINS;
  text_filter.value_kind = FM_PIVOT_FILTER_VALUE_TEXT;
  text_filter.value_text = "o";
  text_filter.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  ASSERT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, pivot_idx, &text_filter), 0) << fm_last_error_message();

  fm_pivot_filter_spec_t range_int_double{};
  range_int_double.axis = FM_PIVOT_AXIS_VALUE;
  range_int_double.field_name = "Amount";
  range_int_double.type = FM_PIVOT_FILTER_VALUE_BETWEEN;
  range_int_double.value_kind = FM_PIVOT_FILTER_VALUE_INT;
  range_int_double.value_int = 100;
  range_int_double.value_high_kind = FM_PIVOT_FILTER_VALUE_DOUBLE;
  range_int_double.value_high_double = 500.25;
  range_int_double.data_field_index = 0;
  ASSERT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, pivot_idx, &range_int_double), 0) << fm_last_error_message();

  fm_pivot_filter_spec_t range_double_int{};
  range_double_int.axis = FM_PIVOT_AXIS_VALUE;
  range_double_int.field_name = "Amount";
  range_double_int.type = FM_PIVOT_FILTER_VALUE_BETWEEN;
  range_double_int.value_kind = FM_PIVOT_FILTER_VALUE_DOUBLE;
  range_double_int.value_double = 99.75;
  range_double_int.value_high_kind = FM_PIVOT_FILTER_VALUE_INT;
  range_double_int.value_high_int = 450;
  range_double_int.data_field_index = 0;
  ASSERT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, pivot_idx, &range_double_int), 0) << fm_last_error_message();

  std::size_t filter_count = 0;
  ASSERT_EQ(fm_workbook_pivot_filter_count(wb.handle, 0, pivot_idx, &filter_count), 0);
  ASSERT_EQ(filter_count, 5U);

  // Establish a projected result before inspection; a read must not change
  // either the active-filter count or the memoised evaluation state.
  PivotCellsGuard before_projection;
  ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 0, pivot_idx, &before_projection.handle), 0) << fm_last_error_message();
  const std::vector<fm_pivot_cell_t> before_cells = CollectCells(before_projection.handle);

  fm_pivot_filter_spec_t got{};
  ASSERT_EQ(fm_workbook_pivot_filter_at(wb.handle, 0, pivot_idx, 0, &got), 0);
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
  ASSERT_EQ(fm_workbook_pivot_filter_at(wb.handle, 0, pivot_idx, 1, &got), 0);
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
  ASSERT_EQ(fm_workbook_pivot_filter_at(wb.handle, 0, pivot_idx, 2, &got), 0);
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
  ASSERT_EQ(fm_workbook_pivot_filter_at(wb.handle, 0, pivot_idx, 3, &got), 0);
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
  ASSERT_EQ(fm_workbook_pivot_filter_at(wb.handle, 0, pivot_idx, 4, &got), 0);
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
  ASSERT_EQ(fm_workbook_pivot_filter_count(wb.handle, 0, pivot_idx, &unchanged_count), 0);
  EXPECT_EQ(unchanged_count, filter_count);
  PivotCellsGuard after_projection;
  ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 0, pivot_idx, &after_projection.handle), 0) << fm_last_error_message();
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
  ASSERT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, pivot_idx, &mutation), 0) << fm_last_error_message();
  fm_pivot_filter_spec_t reacquired{};
  ASSERT_EQ(fm_workbook_pivot_filter_at(wb.handle, 0, pivot_idx, 2, &reacquired), 0);
  EXPECT_EQ(reacquired.value_kind, FM_PIVOT_FILTER_VALUE_TEXT);
  ASSERT_NE(reacquired.value_text, nullptr);
  EXPECT_STREQ(reacquired.value_text, "o");

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
  };
  const auto expect_error_preserves_output = [&](const auto& invoke, fm_status_t expected_status) {
    fm_pivot_filter_spec_t invalid_out = make_sentinel();
    const fm_pivot_filter_spec_t before = invalid_out;
    EXPECT_EQ(invoke(&invalid_out), expected_status);
    expect_unchanged(before, invalid_out);
  };
  expect_error_preserves_output(
      [&](fm_pivot_filter_spec_t* out) { return fm_workbook_pivot_filter_at(wb.handle, 0, pivot_idx, 6, out); },
      invalid);
  expect_error_preserves_output(
      [&](fm_pivot_filter_spec_t* out) { return fm_workbook_pivot_filter_at(wb.handle, 1, pivot_idx, 0, out); },
      invalid);
  expect_error_preserves_output(
      [&](fm_pivot_filter_spec_t* out) { return fm_workbook_pivot_filter_at(wb.handle, 0, pivot_idx + 1, 0, out); },
      invalid);
  expect_error_preserves_output(
      [&](fm_pivot_filter_spec_t* out) { return fm_workbook_pivot_filter_at(nullptr, 0, pivot_idx, 0, out); },
      static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_workbook_pivot_filter_at(wb.handle, 0, pivot_idx, 0, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
}

TEST(FormulonCApiPivot, GetterRejectsUnrepresentableModelEnumsWithoutMutation) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  auto* table = wb.handle->workbook().sheet(0).mutable_pivot_tables()[pivot_idx].get();
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
    EXPECT_EQ(fm_workbook_pivot_filter_at(wb.handle, 0, pivot_idx, 0, &out), invalid);
    EXPECT_EQ(std::memcmp(&out, &before, sizeof(out)), 0);
  };

  inject_and_expect_rejected(formulon::pivot::PivotAxis::None, formulon::pivot::FilterType::LabelContains);
  inject_and_expect_rejected(formulon::pivot::PivotAxis::Row, static_cast<formulon::pivot::FilterType>(0xFF));
}

TEST(FormulonCApiPivot, PivotFilterAddRejectsRawAxisAndTypeWithoutMutation) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  fm_pivot_filter_spec_t spec{};
  spec.axis = RawEnumValue<fm_pivot_axis_t>(99);
  spec.field_name = "Region";
  spec.type = FM_PIVOT_FILTER_LABEL_BEGINS_WITH;
  spec.value_kind = FM_PIVOT_FILTER_VALUE_TEXT;
  spec.value_text = "N";
  spec.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  const auto invalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, pivot_idx, &spec), invalid);

  std::size_t count = 99;
  ASSERT_EQ(fm_workbook_pivot_filter_count(wb.handle, 0, pivot_idx, &count), 0);
  EXPECT_EQ(count, 0U);

  spec.axis = FM_PIVOT_AXIS_ROW;
  spec.type = RawEnumValue<fm_pivot_filter_type_t>(99);
  EXPECT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, pivot_idx, &spec), invalid);
  ASSERT_EQ(fm_workbook_pivot_filter_count(wb.handle, 0, pivot_idx, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST(FormulonCApiPivot, PivotFilterRejectsInvalidDataFieldWithoutMutation) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  fm_pivot_filter_spec_t spec{};
  spec.axis = FM_PIVOT_AXIS_ROW;
  spec.field_name = "Region";
  spec.type = FM_PIVOT_FILTER_VALUE_TOP_10;
  spec.value_kind = FM_PIVOT_FILTER_VALUE_INT;
  spec.value_int = 1;
  spec.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  spec.data_field_index = 1;  // BuildScratchPivot has only slot 0.
  const auto invalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, pivot_idx, &spec), invalid);

  std::size_t count = 99;
  ASSERT_EQ(fm_workbook_pivot_filter_count(wb.handle, 0, pivot_idx, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST(FormulonCApiPivot, PivotFilterRejectsSelectorWhenNoDataFields) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();
  ASSERT_EQ(fm_workbook_pivot_data_field_clear(wb.handle, 0, pivot_idx), 0) << fm_last_error_message();

  fm_pivot_filter_spec_t spec{};
  spec.axis = FM_PIVOT_AXIS_ROW;
  spec.field_name = "Region";
  spec.type = FM_PIVOT_FILTER_VALUE_TOP_10;
  spec.value_kind = FM_PIVOT_FILTER_VALUE_INT;
  spec.value_int = 1;
  spec.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  spec.data_field_index = 0;
  const auto invalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, pivot_idx, &spec), invalid);

  std::size_t count = 99;
  ASSERT_EQ(fm_workbook_pivot_filter_count(wb.handle, 0, pivot_idx, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST(FormulonCApiPivot, PivotFilterRejectsUnknownFieldNameWithoutMutation) {
  // A label filter naming no pivot field would be a silent no-op inside the
  // engine, so the mutator rejects it instead of reporting success.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  fm_pivot_filter_spec_t spec{};
  spec.axis = FM_PIVOT_AXIS_ROW;
  spec.field_name = "NoSuchField";
  spec.type = FM_PIVOT_FILTER_LABEL_BEGINS_WITH;
  spec.value_kind = FM_PIVOT_FILTER_VALUE_TEXT;
  spec.value_text = "N";
  spec.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  const auto invalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, pivot_idx, &spec), invalid);

  std::size_t count = 99;
  ASSERT_EQ(fm_workbook_pivot_filter_count(wb.handle, 0, pivot_idx, &count), 0);
  EXPECT_EQ(count, 0U);

  // The data field's display name resolves through to its source field, so
  // the same call shape is accepted for it.
  spec.field_name = "Sum of Amount";
  ASSERT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, pivot_idx, &spec), 0) << fm_last_error_message();
  ASSERT_EQ(fm_workbook_pivot_filter_count(wb.handle, 0, pivot_idx, &count), 0);
  EXPECT_EQ(count, 1U);
}

TEST(FormulonCApiPivot, ValueFilterRejectsUnknownFieldNameWithoutMutation) {
  // The engine skips a filter whose field does not resolve, so a value filter
  // naming no pivot field must be rejected rather than accepted as a no-op.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  fm_pivot_filter_spec_t spec{};
  spec.axis = FM_PIVOT_AXIS_ROW;
  spec.field_name = "NoSuchField";
  spec.type = FM_PIVOT_FILTER_VALUE_TOP_10;
  spec.value_kind = FM_PIVOT_FILTER_VALUE_INT;
  spec.value_int = 1;
  spec.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  spec.data_field_index = 0;
  const auto invalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, pivot_idx, &spec), invalid);

  std::size_t count = 99;
  ASSERT_EQ(fm_workbook_pivot_filter_count(wb.handle, 0, pivot_idx, &count), 0);
  EXPECT_EQ(count, 0U);

  spec.field_name = "Region";
  ASSERT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, pivot_idx, &spec), 0) << fm_last_error_message();
}

TEST(FormulonCApiPivot, FilterErrorUsesApiName) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  fm_pivot_filter_spec_t spec{};
  spec.axis = FM_PIVOT_AXIS_ROW;
  spec.field_name = "Region";
  spec.type = FM_PIVOT_FILTER_LABEL_BEGINS_WITH;
  spec.value_kind = FM_PIVOT_FILTER_VALUE_NONE;
  spec.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  const auto invalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, pivot_idx, &spec), invalid);
  EXPECT_EQ(std::string_view(fm_last_error_message()).find("fm_workbook_pivot_filter_add:"), 0U);
}

TEST(FormulonCApiPivot, MutateShowValuesAs) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  // Re-set the data field to PercentOfRow.
  std::size_t df_count = 0;
  ASSERT_EQ(fm_workbook_pivot_data_field_count(wb.handle, 0, pivot_idx, &df_count), 0);
  ASSERT_EQ(df_count, 1U);

  fm_pivot_data_field_spec_t df_spec{};
  df_spec.name = "Sum of Amount";
  df_spec.field_index = 1U;  // Amount is the second pivot field.
  df_spec.aggregation = FM_PIVOT_AGG_SUM;
  df_spec.number_format = "";
  df_spec.show_as = FM_PIVOT_SHOW_AS_PERCENT_OF_ROW;
  df_spec.show_as_base_field = -1;
  df_spec.show_as_base_item = -1;
  ASSERT_EQ(fm_workbook_pivot_data_field_set(wb.handle, 0, pivot_idx, 0, &df_spec), 0) << fm_last_error_message();

  PivotCellsGuard projected;
  ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0) << fm_last_error_message();

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
