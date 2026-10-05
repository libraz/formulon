// Sheet spill and range lifetime tests.

#include <algorithm>
#include <chrono>
#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "cell.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace {
// Overwrites enough unrelated cells to make an allocator hand the freed bytes
// out again, so a stale view is unlikely to still find its old contents.
void ChurnTextAllocations(Sheet& sheet) {
  for (std::uint32_t row = 100U; row < 200U; ++row) {
    sheet.set_cell_text(row, 0U, "filler payload of about the same length as the original");
  }
}
TEST(SheetReadRange, TextOutlivesTheWriteThatFreesTheCellsOwnBytes) {
  Sheet s("Sheet1");
  const std::string original = "a text payload long enough to live on the heap";
  s.set_cell_text(0U, 0U, original);

  Arena arena;
  std::vector<Value> bulk;
  std::vector<std::size_t> formula_indices;
  s.read_range(0U, 0U, 0U, 0U, arena, bulk, formula_indices);
  ASSERT_EQ(bulk.size(), 1U);
  ASSERT_TRUE(bulk[0].is_text());

  const Cell* cell = s.cell_at(0U, 0U);
  ASSERT_NE(cell, nullptr);
  ASSERT_NE(cell->cached_text_owned, nullptr);
  EXPECT_NE(bulk[0].as_text().data(), cell->cached_text_owned->data())
      << "the read handed back a view into the cell's own allocation";

  // The next cached write replaces that allocation.
  s.set_cell_cached_value(0U, 0U, Value::text("a replacement payload of a comparable length"));
  ChurnTextAllocations(s);
  EXPECT_EQ(bulk[0].as_text(), original);
}
TEST(SheetReadRange, TextOutlivesTheClearThatFreesTheSpillRegionsBytes) {
  Sheet s("Sheet1");
  const std::string original = "a spilled text payload long enough to live on the heap";
  std::vector<Value> cells = {Value::number(1.0), Value::text(original)};
  ASSERT_TRUE(s.commit_spill(0U, 0U, 2U, 1U, std::move(cells)));

  Arena arena;
  std::vector<Value> bulk;
  std::vector<std::size_t> formula_indices;
  s.read_range(0U, 1U, 0U, 0U, arena, bulk, formula_indices);
  ASSERT_EQ(bulk.size(), 2U);
  ASSERT_TRUE(bulk[1].is_text());

  const SpillRegion* region = s.spill_region_at_anchor(0U, 0U);
  ASSERT_NE(region, nullptr);
  ASSERT_EQ(region->owned_strings.size(), 1U);
  EXPECT_NE(bulk[1].as_text().data(), region->owned_strings.front().data())
      << "the read handed back a view into the region's own storage";

  s.clear_spill(0U, 0U);
  ChurnTextAllocations(s);
  EXPECT_EQ(bulk[1].as_text(), original);
}
TEST(SheetSpillTest, ReadSpillRegionAtAnchorOutlivesTheRegionItCopied) {
  Sheet s("Sheet1");
  const std::string original = "a spilled text payload long enough to live on the heap";
  std::vector<Value> cells = {Value::number(1.0), Value::text(original)};
  ASSERT_TRUE(s.commit_spill(0U, 0U, 2U, 1U, std::move(cells)));

  Arena arena;
  std::vector<Value> copied;
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
  ASSERT_TRUE(s.read_spill_region_at_anchor(0U, 0U, arena, copied, &rows, &cols));
  EXPECT_EQ(rows, 2U);
  EXPECT_EQ(cols, 1U);
  ASSERT_EQ(copied.size(), 2U);
  EXPECT_EQ(copied[0], Value::number(1.0));
  ASSERT_TRUE(copied[1].is_text());

  s.clear_spill(0U, 0U);
  ChurnTextAllocations(s);
  EXPECT_EQ(copied[1].as_text(), original);

  // A coordinate that anchors nothing appends nothing and says so.
  EXPECT_FALSE(s.read_spill_region_at_anchor(9U, 9U, arena, copied, &rows, &cols));
  EXPECT_EQ(copied.size(), 2U);
}

}  // namespace
}  // namespace formulon
