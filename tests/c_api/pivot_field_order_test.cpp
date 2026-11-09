// C ABI PivotTable field-order replacement behavior.

#include <array>
#include <cstdint>
#include <string>
#include <vector>

#include "formulon_c_pivot_test_helpers.h"

namespace {

constexpr std::size_t kFieldCount = 5U;

fm_status_t BuildFieldOrderPivot(fm_workbook_t* wb, std::size_t* out_pivot_index) {
  std::uint32_t cache_id = 0;
  if (const fm_status_t status = fm_workbook_add_pivot_cache(wb, 0U, &cache_id); status != 0) {
    return status;
  }
  constexpr std::array<const char*, kFieldCount> kFieldNames = {"A", "B", "C", "D", "E"};
  for (const char* name : kFieldNames) {
    std::size_t cache_field_index = 0;
    if (const fm_status_t status = fm_workbook_add_pivot_cache_field(wb, cache_id, name, &cache_field_index);
        status != 0) {
      return status;
    }
  }
  if (const fm_status_t status = AddPivot(wb, 0U, "FieldOrder", cache_id, 0U, 0U, out_pivot_index); status != 0) {
    return status;
  }
  for (const char* name : kFieldNames) {
    fm_pivot_field_spec_t spec = FM_STRUCT_INIT(fm_pivot_field_spec_t);
    spec.source_name = name;
    spec.axis = FM_PIVOT_AXIS_VALUE;
    spec.subtotal_top = FM_TRISTATE_TRUE;
    std::size_t field_index = 0;
    if (const fm_status_t status = fm_sheet_add_pivot_field(wb, 0U, *out_pivot_index, &spec, &field_index);
        status != 0) {
      return status;
    }
  }
  return 0;
}

std::vector<std::uint32_t> ReadFieldOrder(const fm_workbook_t* wb, std::size_t pivot_index, bool row) {
  const auto read = [&](std::uint32_t* out, std::size_t capacity, std::size_t* count) {
    return row ? fm_sheet_get_pivot_row_field_order(wb, 0U, pivot_index, out, capacity, count)
               : fm_sheet_get_pivot_column_field_order(wb, 0U, pivot_index, out, capacity, count);
  };
  std::size_t count = 0;
  fm_status_t status = read(nullptr, 0U, &count);
  EXPECT_EQ(status, 0) << fm_get_last_error_message();
  std::vector<std::uint32_t> indices(count);
  if (status == 0 && count != 0U) {
    status = read(indices.data(), indices.size(), &count);
    EXPECT_EQ(status, 0) << fm_get_last_error_message();
  }
  return status == 0 ? indices : std::vector<std::uint32_t>{};
}

formulon::pivot::PivotTable& GetPivot(WorkbookGuard& wb, std::size_t pivot_index) {
  return *wb.handle->wb->sheet(0U).mutable_pivot_tables()[pivot_index];
}

void ExpectAxes(const fm_workbook_t* wb, std::size_t pivot_index, const std::array<std::int32_t, kFieldCount>& axes) {
  for (std::size_t i = 0; i < axes.size(); ++i) {
    fm_pivot_field_info_t field = FM_STRUCT_INIT(fm_pivot_field_info_t);
    ASSERT_EQ(fm_sheet_get_pivot_field_at(wb, 0U, pivot_index, i, &field), 0) << fm_get_last_error_message();
    EXPECT_EQ(field.axis, axes[i]) << "field=" << i;
  }
}

fm_status_t SetFieldOrder(fm_workbook_t* wb, std::size_t pivot_index, bool row, const std::uint32_t* indices,
                          std::size_t count) {
  return row ? fm_sheet_set_pivot_row_field_order(wb, 0U, pivot_index, indices, count)
             : fm_sheet_set_pivot_column_field_order(wb, 0U, pivot_index, indices, count);
}

}  // namespace

TEST(FormulonCApiPivotFieldOrder, ReplacesMembershipAndMovesMultipleFieldsBetweenAxes) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildFieldOrderPivot(wb.handle, &pivot_index), 0) << fm_get_last_error_message();

  const std::uint32_t row_order[] = {0U, 1U, 2U};
  const std::uint32_t col_order[] = {3U, 4U};
  ASSERT_EQ(SetFieldOrder(wb.handle, pivot_index, true, row_order, 3U), 0) << fm_get_last_error_message();
  ASSERT_EQ(SetFieldOrder(wb.handle, pivot_index, false, col_order, 2U), 0) << fm_get_last_error_message();

  // Reordering columns moves two fields from rows, retains field 4, and drops
  // the omitted field 3 from the column axis.
  const std::uint32_t reordered_cols[] = {2U, 4U, 0U};
  ASSERT_EQ(SetFieldOrder(wb.handle, pivot_index, false, reordered_cols, 3U), 0) << fm_get_last_error_message();
  EXPECT_EQ(ReadFieldOrder(wb.handle, pivot_index, true), (std::vector<std::uint32_t>{1U}));
  EXPECT_EQ(ReadFieldOrder(wb.handle, pivot_index, false), (std::vector<std::uint32_t>{2U, 4U, 0U}));
  ExpectAxes(wb.handle, pivot_index,
             {FM_PIVOT_AXIS_COLUMN, FM_PIVOT_AXIS_ROW, FM_PIVOT_AXIS_COLUMN, -1, FM_PIVOT_AXIS_COLUMN});

  // A later row replacement moves field 3 from no axis and field 1 from rows;
  // the column membership and order remain intact.
  const std::uint32_t reordered_rows[] = {3U, 1U};
  ASSERT_EQ(SetFieldOrder(wb.handle, pivot_index, true, reordered_rows, 2U), 0) << fm_get_last_error_message();
  EXPECT_EQ(ReadFieldOrder(wb.handle, pivot_index, true), (std::vector<std::uint32_t>{3U, 1U}));
  EXPECT_EQ(ReadFieldOrder(wb.handle, pivot_index, false), (std::vector<std::uint32_t>{2U, 4U, 0U}));
  ExpectAxes(wb.handle, pivot_index,
             {FM_PIVOT_AXIS_COLUMN, FM_PIVOT_AXIS_ROW, FM_PIVOT_AXIS_COLUMN, FM_PIVOT_AXIS_ROW, FM_PIVOT_AXIS_COLUMN});
}

TEST(FormulonCApiPivotFieldOrder, AcceptsGetterOrderAsTheOtherAxisInput) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildFieldOrderPivot(wb.handle, &pivot_index), 0) << fm_get_last_error_message();

  const std::uint32_t original[] = {0U, 1U};
  ASSERT_EQ(fm_sheet_set_pivot_row_field_order(wb.handle, 0U, pivot_index, original, 2U), 0)
      << fm_get_last_error_message();
  std::uint32_t borrowed_row_order[8] = {};
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_pivot_row_field_order(wb.handle, 0U, pivot_index, borrowed_row_order, 8U, &count), 0);
  ASSERT_EQ(count, 2U);

  // The getter's listing is a valid input for the opposite axis, even though
  // the operation removes those fields from the axis it was read from.
  ASSERT_EQ(fm_sheet_set_pivot_column_field_order(wb.handle, 0U, pivot_index, borrowed_row_order, count), 0)
      << fm_get_last_error_message();
  EXPECT_TRUE(ReadFieldOrder(wb.handle, pivot_index, true).empty());
  EXPECT_EQ(ReadFieldOrder(wb.handle, pivot_index, false), (std::vector<std::uint32_t>{0U, 1U}));
  ExpectAxes(
      wb.handle, pivot_index,
      {FM_PIVOT_AXIS_COLUMN, FM_PIVOT_AXIS_COLUMN, FM_PIVOT_AXIS_VALUE, FM_PIVOT_AXIS_VALUE, FM_PIVOT_AXIS_VALUE});
}

TEST(FormulonCApiPivotFieldOrder, ASmallerBufferTakesThePrefixAndReportsTheFullCount) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildFieldOrderPivot(wb.handle, &pivot_index), 0) << fm_get_last_error_message();
  const std::uint32_t order[] = {1U, 0U};
  ASSERT_EQ(fm_sheet_set_pivot_row_field_order(wb.handle, 0U, pivot_index, order, 2U), 0)
      << fm_get_last_error_message();

  std::uint32_t buffer[2] = {77U, 77U};
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_pivot_row_field_order(wb.handle, 0U, pivot_index, buffer, 1U, &count), 0);
  EXPECT_EQ(count, 2U);
  EXPECT_EQ(buffer[0], 1U);
  EXPECT_EQ(buffer[1], 77U) << "nothing is written past the capacity";
}

TEST(FormulonCApiPivotFieldOrder, CompactsOppositeAxisAndAdjustsValuesMarker) {
  for (const std::size_t marker_position : {1U, 2U, 3U}) {
    for (const bool row : {true, false}) {
      WorkbookGuard wb;
      ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
      std::size_t pivot_index = 0;
      ASSERT_EQ(BuildFieldOrderPivot(wb.handle, &pivot_index), 0) << fm_get_last_error_message();

      const std::uint32_t full_order[] = {0U, 1U, 2U, 3U, 4U};
      ASSERT_EQ(SetFieldOrder(wb.handle, pivot_index, row, full_order, 5U), 0) << fm_get_last_error_message();
      formulon::pivot::PivotTable& table = GetPivot(wb, pivot_index);
      if (row) {
        table.set_row_values_position(marker_position);
      } else {
        table.set_col_values_position(marker_position);
      }

      const std::uint32_t moved[] = {2U};
      ASSERT_EQ(SetFieldOrder(wb.handle, pivot_index, !row, moved, 1U), 0) << fm_get_last_error_message();
      const std::size_t expected_position = marker_position - (marker_position > 2U ? 1U : 0U);
      EXPECT_EQ(row ? table.row_values_position() : table.col_values_position(), expected_position);
      EXPECT_EQ(ReadFieldOrder(wb.handle, pivot_index, row), (std::vector<std::uint32_t>{0U, 1U, 3U, 4U}));
      EXPECT_EQ(ReadFieldOrder(wb.handle, pivot_index, !row), (std::vector<std::uint32_t>{2U}));
    }
  }
}

TEST(FormulonCApiPivotFieldOrder, CompactsMultipleMovedFieldsBeforeMarkerOnBothAxes) {
  for (const bool row : {true, false}) {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    std::size_t pivot_index = 0;
    ASSERT_EQ(BuildFieldOrderPivot(wb.handle, &pivot_index), 0) << fm_get_last_error_message();

    const std::uint32_t full_order[] = {0U, 1U, 2U, 3U, 4U};
    ASSERT_EQ(SetFieldOrder(wb.handle, pivot_index, row, full_order, 5U), 0) << fm_get_last_error_message();
    formulon::pivot::PivotTable& table = GetPivot(wb, pivot_index);
    if (row) {
      table.set_row_values_position(4U);
    } else {
      table.set_col_values_position(4U);
    }

    // Remove positions 3, 0, and 1 in an order different from their document
    // order. All three precede the marker, so it shifts from 4 to 1.
    const std::uint32_t moved[] = {3U, 0U, 1U};
    ASSERT_EQ(SetFieldOrder(wb.handle, pivot_index, !row, moved, 3U), 0) << fm_get_last_error_message();
    EXPECT_EQ(row ? table.row_values_position() : table.col_values_position(), 1U);
    EXPECT_EQ(ReadFieldOrder(wb.handle, pivot_index, row), (std::vector<std::uint32_t>{2U, 4U}));
    EXPECT_EQ(ReadFieldOrder(wb.handle, pivot_index, !row), (std::vector<std::uint32_t>{3U, 0U, 1U}));
  }
}

TEST(FormulonCApiPivotFieldOrder, ClampsTargetValuesMarkerToShorterOrder) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildFieldOrderPivot(wb.handle, &pivot_index), 0) << fm_get_last_error_message();

  const std::uint32_t full_order[] = {0U, 1U, 2U, 3U};
  ASSERT_EQ(fm_sheet_set_pivot_row_field_order(wb.handle, 0U, pivot_index, full_order, 4U), 0);
  formulon::pivot::PivotTable& table = GetPivot(wb, pivot_index);
  table.set_row_values_position(4U);
  const std::uint32_t shortened[] = {3U};
  ASSERT_EQ(fm_sheet_set_pivot_row_field_order(wb.handle, 0U, pivot_index, shortened, 1U), 0);
  EXPECT_EQ(table.row_values_position(), 1U);
}

TEST(FormulonCApiPivotFieldOrder, InvalidAndDuplicateOrdersLeaveAllStateUntouched) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildFieldOrderPivot(wb.handle, &pivot_index), 0) << fm_get_last_error_message();

  const std::uint32_t rows[] = {0U, 1U};
  const std::uint32_t cols[] = {2U, 3U, 4U};
  ASSERT_EQ(SetFieldOrder(wb.handle, pivot_index, true, rows, 2U), 0);
  ASSERT_EQ(SetFieldOrder(wb.handle, pivot_index, false, cols, 3U), 0);
  formulon::pivot::PivotTable& table = GetPivot(wb, pivot_index);
  table.set_row_values_position(1U);
  table.set_col_values_position(2U);

  const std::uint32_t out_of_range[] = {3U, 99U, 0U};
  EXPECT_EQ(SetFieldOrder(wb.handle, pivot_index, true, out_of_range, 3U),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kIndexOutOfRange));
  std::string error = fm_get_last_error_message();
  EXPECT_NE(error.find("indices[i] 99 is out of range"), std::string::npos) << error;
  EXPECT_NE(std::string(fm_get_last_error_context()).find("indices[i]=99 count="), std::string::npos);
  EXPECT_EQ(ReadFieldOrder(wb.handle, pivot_index, true), (std::vector<std::uint32_t>{0U, 1U}));
  EXPECT_EQ(ReadFieldOrder(wb.handle, pivot_index, false), (std::vector<std::uint32_t>{2U, 3U, 4U}));
  EXPECT_EQ(table.row_values_position(), 1U);
  EXPECT_EQ(table.col_values_position(), 2U);

  const std::uint32_t duplicate[] = {4U, 4U};
  EXPECT_EQ(SetFieldOrder(wb.handle, pivot_index, false, duplicate, 2U),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  error = fm_get_last_error_message();
  EXPECT_NE(error.find("indices lists a field twice"), std::string::npos) << error;
  EXPECT_STREQ(fm_get_last_error_context(), "i=1 value=4");
  EXPECT_EQ(ReadFieldOrder(wb.handle, pivot_index, true), (std::vector<std::uint32_t>{0U, 1U}));
  EXPECT_EQ(ReadFieldOrder(wb.handle, pivot_index, false), (std::vector<std::uint32_t>{2U, 3U, 4U}));
  EXPECT_EQ(table.row_values_position(), 1U);
  EXPECT_EQ(table.col_values_position(), 2U);
  ExpectAxes(wb.handle, pivot_index,
             {FM_PIVOT_AXIS_ROW, FM_PIVOT_AXIS_ROW, FM_PIVOT_AXIS_COLUMN, FM_PIVOT_AXIS_COLUMN, FM_PIVOT_AXIS_COLUMN});
}

TEST(FormulonCApiPivotFieldOrder, NullPointerWithEmptyOrderClearsBothAxesSafely) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildFieldOrderPivot(wb.handle, &pivot_index), 0) << fm_get_last_error_message();

  const std::uint32_t rows[] = {0U, 1U};
  const std::uint32_t cols[] = {2U, 3U, 4U};
  ASSERT_EQ(SetFieldOrder(wb.handle, pivot_index, true, rows, 2U), 0);
  ASSERT_EQ(SetFieldOrder(wb.handle, pivot_index, false, cols, 3U), 0);

  ASSERT_EQ(fm_sheet_set_pivot_row_field_order(wb.handle, 0U, pivot_index, nullptr, 0U), 0)
      << fm_get_last_error_message();
  EXPECT_TRUE(ReadFieldOrder(wb.handle, pivot_index, true).empty());
  EXPECT_EQ(ReadFieldOrder(wb.handle, pivot_index, false), (std::vector<std::uint32_t>{2U, 3U, 4U}));
  ExpectAxes(wb.handle, pivot_index, {-1, -1, FM_PIVOT_AXIS_COLUMN, FM_PIVOT_AXIS_COLUMN, FM_PIVOT_AXIS_COLUMN});

  ASSERT_EQ(fm_sheet_set_pivot_column_field_order(wb.handle, 0U, pivot_index, nullptr, 0U), 0)
      << fm_get_last_error_message();
  EXPECT_TRUE(ReadFieldOrder(wb.handle, pivot_index, false).empty());
  ExpectAxes(wb.handle, pivot_index, {-1, -1, -1, -1, -1});
}

TEST(FormulonCApiPivotFieldOrder, MovesOnlyFirstOccurrenceFromMalformedOppositeOrder) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildFieldOrderPivot(wb.handle, &pivot_index), 0) << fm_get_last_error_message();

  const std::uint32_t rows[] = {0U, 1U, 2U, 3U};
  ASSERT_EQ(fm_sheet_set_pivot_row_field_order(wb.handle, 0U, pivot_index, rows, 4U), 0);
  GetPivot(wb, pivot_index).mutable_row_field_order() = {0U, 1U, 2U, 1U, 3U};

  const std::uint32_t moved[] = {1U};
  ASSERT_EQ(fm_sheet_set_pivot_column_field_order(wb.handle, 0U, pivot_index, moved, 1U), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(ReadFieldOrder(wb.handle, pivot_index, true), (std::vector<std::uint32_t>{0U, 2U, 1U, 3U}));
  EXPECT_EQ(ReadFieldOrder(wb.handle, pivot_index, false), (std::vector<std::uint32_t>{1U}));
}

namespace {

std::vector<std::uint32_t> ReadPageOrder(const fm_workbook_t* wb, std::size_t pivot_index) {
  std::size_t count = 0;
  EXPECT_EQ(fm_sheet_get_pivot_page_field_order(wb, 0U, pivot_index, nullptr, 0U, &count), 0)
      << fm_get_last_error_message();
  std::vector<std::uint32_t> indices(count);
  if (count != 0U) {
    EXPECT_EQ(fm_sheet_get_pivot_page_field_order(wb, 0U, pivot_index, indices.data(), indices.size(), &count), 0)
        << fm_get_last_error_message();
  }
  return indices;
}

}  // namespace

TEST(FormulonCApiPivotFieldOrder, PageOrderReadsSetsAndReadsBackReversed) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildFieldOrderPivot(wb.handle, &pivot_index), 0) << fm_get_last_error_message();
  EXPECT_TRUE(ReadPageOrder(wb.handle, pivot_index).empty());

  const std::uint32_t order[] = {1U, 3U};
  ASSERT_EQ(fm_sheet_set_pivot_page_field_order(wb.handle, 0U, pivot_index, order, 2U), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(ReadPageOrder(wb.handle, pivot_index), (std::vector<std::uint32_t>{1U, 3U}));
  ExpectAxes(wb.handle, pivot_index,
             {FM_PIVOT_AXIS_VALUE, FM_PIVOT_AXIS_PAGE, FM_PIVOT_AXIS_VALUE, FM_PIVOT_AXIS_PAGE, FM_PIVOT_AXIS_VALUE});

  const std::uint32_t reversed[] = {3U, 1U};
  ASSERT_EQ(fm_sheet_set_pivot_page_field_order(wb.handle, 0U, pivot_index, reversed, 2U), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(ReadPageOrder(wb.handle, pivot_index), (std::vector<std::uint32_t>{3U, 1U}));

  // A field left out comes off the page axis.
  const std::uint32_t only[] = {3U};
  ASSERT_EQ(fm_sheet_set_pivot_page_field_order(wb.handle, 0U, pivot_index, only, 1U), 0);
  EXPECT_EQ(ReadPageOrder(wb.handle, pivot_index), (std::vector<std::uint32_t>{3U}));
  ExpectAxes(wb.handle, pivot_index,
             {FM_PIVOT_AXIS_VALUE, -1, FM_PIVOT_AXIS_VALUE, FM_PIVOT_AXIS_PAGE, FM_PIVOT_AXIS_VALUE});
}

TEST(FormulonCApiPivotFieldOrder, ListingARowFieldOnThePageMovesItOffTheRowAxis) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildFieldOrderPivot(wb.handle, &pivot_index), 0) << fm_get_last_error_message();
  const std::uint32_t rows[] = {0U, 1U};
  ASSERT_EQ(SetFieldOrder(wb.handle, pivot_index, true, rows, 2U), 0);
  const std::uint32_t page[] = {1U};
  ASSERT_EQ(fm_sheet_set_pivot_page_field_order(wb.handle, 0U, pivot_index, page, 1U), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(ReadFieldOrder(wb.handle, pivot_index, true), (std::vector<std::uint32_t>{0U}));
  EXPECT_EQ(ReadPageOrder(wb.handle, pivot_index), (std::vector<std::uint32_t>{1U}));
}

TEST(FormulonCApiPivotFieldOrder, PageOrderRefusesTheValuesFieldDuplicatesAndPastTheEnd) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildFieldOrderPivot(wb.handle, &pivot_index), 0) << fm_get_last_error_message();
  const std::uint32_t seed[] = {2U};
  ASSERT_EQ(fm_sheet_set_pivot_page_field_order(wb.handle, 0U, pivot_index, seed, 1U), 0);

  const std::uint32_t values[] = {0U, FM_PIVOT_VALUES_FIELD};
  EXPECT_EQ(fm_sheet_set_pivot_page_field_order(wb.handle, 0U, pivot_index, values, 2U), FM_STATUS_INVALID_ARGUMENT);
  const std::uint32_t twice[] = {0U, 0U};
  EXPECT_EQ(fm_sheet_set_pivot_page_field_order(wb.handle, 0U, pivot_index, twice, 2U), FM_STATUS_INVALID_ARGUMENT);
  const std::uint32_t past[] = {0U, static_cast<std::uint32_t>(kFieldCount)};
  EXPECT_EQ(fm_sheet_set_pivot_page_field_order(wb.handle, 0U, pivot_index, past, 2U), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(ReadPageOrder(wb.handle, pivot_index), (std::vector<std::uint32_t>{2U})) << "a refusal changes nothing";
}

TEST(FormulonCApiPivotFieldOrder, AFieldMovedOffThePageAxisAndBackJoinsThePageOrderLast) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildFieldOrderPivot(wb.handle, &pivot_index), 0) << fm_get_last_error_message();
  const std::uint32_t page[] = {1U, 2U, 3U};
  ASSERT_EQ(fm_sheet_set_pivot_page_field_order(wb.handle, 0U, pivot_index, page, 3U), 0)
      << fm_get_last_error_message();

  ASSERT_EQ(fm_sheet_set_pivot_field_axis(wb.handle, 0U, pivot_index, 1U, FM_PIVOT_AXIS_ROW), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(ReadPageOrder(wb.handle, pivot_index), (std::vector<std::uint32_t>{2U, 3U}));
  ASSERT_EQ(fm_sheet_set_pivot_field_axis(wb.handle, 0U, pivot_index, 1U, FM_PIVOT_AXIS_PAGE), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(ReadPageOrder(wb.handle, pivot_index), (std::vector<std::uint32_t>{2U, 3U, 1U}));
}
