#include "sheet_layout_edit.h"

#include <cstdint>
#include <vector>

#include "gtest/gtest.h"
#include "style_region.h"

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
  upsert_row_override(layout, 40U).hidden = true;
  upsert_row_override(layout, 3U).outline_level = 2U;
  set_row_style_span(layout, 1U, 5U, 4U);
  set_row_style_span(layout, 38U, 41U, 6U);

  EXPECT_EQ(Rows(layout), (std::vector<std::uint32_t>{1U, 2U, 3U, 4U, 5U, 38U, 39U, 40U, 41U}));
  EXPECT_EQ(find_row_override(layout, 3U)->outline_level, 2U);
  EXPECT_EQ(find_row_override(layout, 40U)->style_xf, 6U);

  clear_row_style_span(layout, 0U, 100U);
  EXPECT_EQ(Rows(layout), (std::vector<std::uint32_t>{3U, 40U}));
  EXPECT_TRUE(find_row_override(layout, 40U)->hidden);
}

}  // namespace
}  // namespace formulon
