#include "sheet_layout_edit.h"

#include <cstdint>
#include <vector>

#include "gtest/gtest.h"
#include "sheet.h"
#include "style_region.h"
#include "utils/error.h"
#include "workbook.h"

namespace formulon {
namespace {

std::vector<std::uint32_t> Rows(const SheetLayout& layout) {
  std::vector<std::uint32_t> out;
  for (const RowLayout& entry : layout.row_overrides) {
    out.push_back(entry.row);
  }
  return out;
}

TEST(SheetLayoutEdit, InsertKeepsRowOrderAndTheFirstDuplicate) {
  SheetLayout layout;
  for (const std::uint32_t row : {5U, 9U, 2U, 7U}) {
    RowLayout entry;
    entry.row = row;
    entry.height = static_cast<double>(row);
    insert_row_override(layout, entry);
  }
  RowLayout duplicate;
  duplicate.row = 7U;
  duplicate.height = 99.0;
  insert_row_override(layout, duplicate);

  EXPECT_EQ(Rows(layout), (std::vector<std::uint32_t>{2U, 5U, 7U, 9U}));
  ASSERT_NE(find_row_override(layout, 7U), nullptr);
  EXPECT_DOUBLE_EQ(find_row_override(layout, 7U)->height, 7.0);
  EXPECT_EQ(find_row_override(layout, 6U), nullptr);
}

TEST(SheetLayoutEdit, UpsertAndRowStyleSpansStaySorted) {
  SheetLayout layout;
  upsert_row_override(layout, 40U, false).hidden = true;
  upsert_row_override(layout, 3U, false).outline_level = 2U;
  set_row_style_span(layout, 1U, 5U, 4U, false);
  set_row_style_span(layout, 38U, 41U, 6U, false);

  EXPECT_EQ(Rows(layout), (std::vector<std::uint32_t>{1U, 2U, 3U, 4U, 5U, 38U, 39U, 40U, 41U}));
  EXPECT_EQ(find_row_override(layout, 3U)->outline_level, 2U);
  EXPECT_EQ(find_row_override(layout, 40U)->style_xf, 6U);

  clear_row_style_span(layout, 0U, 100U, false);
  EXPECT_EQ(Rows(layout), (std::vector<std::uint32_t>{3U, 40U}));
  EXPECT_TRUE(find_row_override(layout, 40U)->hidden);
}

TEST(SheetLayoutEdit, VisibilityAwareStyleClearStaysSparse) {
  SheetLayout layout;
  RowLayout styled;
  styled.row = 3U;
  styled.has_style = true;
  styled.style_xf = 7U;
  RowLayout visible;
  visible.row = 11U;
  RowLayout hidden_styled;
  hidden_styled.row = Sheet::kMaxRows - 1U;
  hidden_styled.hidden = true;
  hidden_styled.has_style = true;
  hidden_styled.style_xf = 9U;
  layout.row_overrides = {styled, visible, hidden_styled};

  clear_row_style_span(layout, 0U, Sheet::kMaxRows - 1U, true);

  ASSERT_EQ(layout.row_overrides.size(), 3U);
  EXPECT_EQ(layout.row_overrides[0].row, 3U);
  EXPECT_FALSE(layout.row_overrides[0].hidden);
  EXPECT_FALSE(layout.row_overrides[0].has_style);
  EXPECT_EQ(layout.row_overrides[1].row, 11U);
  EXPECT_FALSE(layout.row_overrides[1].hidden);
  EXPECT_FALSE(layout.row_overrides[1].has_style);
  EXPECT_EQ(layout.row_overrides[2].row, Sheet::kMaxRows - 1U);
  EXPECT_TRUE(layout.row_overrides[2].hidden);
  EXPECT_FALSE(layout.row_overrides[2].has_style);

  clear_row_style_span(layout, 0U, Sheet::kMaxRows - 1U, false);
  ASSERT_EQ(layout.row_overrides.size(), 1U);
  EXPECT_EQ(layout.row_overrides[0].row, Sheet::kMaxRows - 1U);
}

TEST(SheetLayoutEdit, NullStyleCaptureRejectsBeforeChangingSheet) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 2U, 3U, Value::number(42.0))));
  RowLayout row;
  row.row = 8U;
  row.hidden = false;
  row.has_style = true;
  row.style_xf = 0U;
  wb.sheet(0U).mutable_layout().row_overrides.push_back(row);

  const std::vector<RowLayout> rows_before = wb.sheet(0U).layout().row_overrides;
  for (const std::size_t len : {0U, 4U}) {
    const auto restored = wb.restore_range_style(0U, nullptr, len);
    ASSERT_FALSE(static_cast<bool>(restored));
    EXPECT_EQ(restored.error().code, FormulonErrorCode::kInvalidArgument);
    ASSERT_NE(wb.sheet(0U).cell_at(2U, 3U), nullptr);
    EXPECT_DOUBLE_EQ(wb.sheet(0U).cell_at(2U, 3U)->cached_value.as_number(), 42.0);
    ASSERT_EQ(wb.sheet(0U).layout().row_overrides.size(), rows_before.size());
    EXPECT_EQ(wb.sheet(0U).layout().row_overrides[0].row, rows_before[0].row);
    EXPECT_EQ(wb.sheet(0U).layout().row_overrides[0].style_xf, rows_before[0].style_xf);
  }
}

}  // namespace
}  // namespace formulon
