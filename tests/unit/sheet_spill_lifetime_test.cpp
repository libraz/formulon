// Sheet spill and range lifetime tests.

#include <algorithm>
#include <array>
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
constexpr std::array<std::string_view, 2> kFirstSpillTexts = {
    "SSO", "a spill anchor payload long enough to require heap storage"};

// Overwrites enough unrelated cells to make an allocator hand the freed bytes
// out again, so a stale view is unlikely to still find its old contents.
void ChurnTextAllocations(Sheet& sheet) {
  for (std::uint32_t row = 100U; row < 200U; ++row) {
    sheet.set_cell_text(row, 0U, "filler payload of about the same length as the original");
  }
}

bool CommitTextAnchorSpill(Sheet& sheet, std::string_view text, std::uint32_t row, std::uint32_t col) {
  std::vector<Value> cells = {Value::text(text), Value::number(42.0)};
  if (!sheet.commit_spill(row, col, 1U, 2U, std::move(cells))) {
    ADD_FAILURE() << "could not commit the text spill";
    return false;
  }

  const SpillRegion* region = sheet.spill_region_at_anchor(row, col);
  if (region == nullptr || region->owned_strings.size() != 1U) {
    ADD_FAILURE() << "committed region does not own its first text value";
    return false;
  }
  const Cell* anchor = sheet.cell_at(row, col);
  if (anchor == nullptr || !anchor->cached_value.is_text() || anchor->cached_text_owned == nullptr) {
    ADD_FAILURE() << "the anchor must own its cached text independently of the spill region";
    return false;
  }
  EXPECT_NE(anchor->cached_text_owned->data(), region->owned_strings.front().data())
      << "the anchor must not point into the region's storage";
  EXPECT_EQ(anchor->cached_value.as_text(), text);
  return true;
}

void ExpectAnchorRetainsText(const Sheet& sheet, std::uint32_t row, std::uint32_t col, std::string_view text) {
  const Cell* anchor = sheet.cell_at(row, col);
  ASSERT_NE(anchor, nullptr);
  ASSERT_TRUE(anchor->cached_value.is_text());
  ASSERT_NE(anchor->cached_text_owned, nullptr);
  EXPECT_EQ(anchor->cached_value.as_text(), text);
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

TEST(SheetSpillLifetime, AnchorTextSurvivesClearSpill) {
  for (const std::string_view text : kFirstSpillTexts) {
    Sheet sheet("Sheet1");
    if (!CommitTextAnchorSpill(sheet, text, 4U, 4U)) {
      continue;
    }

    sheet.clear_spill(4U, 4U);
    EXPECT_EQ(sheet.spill_region_at_anchor(4U, 4U), nullptr);
    ChurnTextAllocations(sheet);
    ExpectAnchorRetainsText(sheet, 4U, 4U, text);
  }
}

TEST(SheetSpillLifetime, AnchorTextSurvivesClearAllSpills) {
  for (const std::string_view text : kFirstSpillTexts) {
    Sheet sheet("Sheet1");
    if (!CommitTextAnchorSpill(sheet, text, 4U, 4U)) {
      continue;
    }

    sheet.clear_all_spills();
    EXPECT_EQ(sheet.spill_region_at_anchor(4U, 4U), nullptr);
    ChurnTextAllocations(sheet);
    ExpectAnchorRetainsText(sheet, 4U, 4U, text);
  }
}

TEST(SheetSpillLifetime, AnchorTextSurvivesPhantomOverwrite) {
  for (const std::string_view text : kFirstSpillTexts) {
    Sheet sheet("Sheet1");
    if (!CommitTextAnchorSpill(sheet, text, 4U, 4U)) {
      continue;
    }

    sheet.set_cell_value(4U, 5U, Value::number(7.0));
    EXPECT_EQ(sheet.spill_region_at_anchor(4U, 4U), nullptr);
    EXPECT_EQ(sheet.resolve_cell_value(4U, 5U), Value::number(7.0));
    ChurnTextAllocations(sheet);
    ExpectAnchorRetainsText(sheet, 4U, 4U, text);
  }
}

TEST(SheetSpillLifetime, AnchorTextSurvivesRowInsertion) {
  for (const std::string_view text : kFirstSpillTexts) {
    Sheet sheet("Sheet1");
    if (!CommitTextAnchorSpill(sheet, text, 4U, 4U)) {
      continue;
    }

    sheet.insert_rows(0U, 1U);
    EXPECT_EQ(sheet.spill_region_at_anchor(5U, 4U), nullptr);
    ChurnTextAllocations(sheet);
    ExpectAnchorRetainsText(sheet, 5U, 4U, text);
  }
}

TEST(SheetSpillLifetime, AnchorTextSurvivesColumnInsertion) {
  for (const std::string_view text : kFirstSpillTexts) {
    Sheet sheet("Sheet1");
    if (!CommitTextAnchorSpill(sheet, text, 4U, 4U)) {
      continue;
    }

    sheet.insert_cols(0U, 1U);
    EXPECT_EQ(sheet.spill_region_at_anchor(4U, 5U), nullptr);
    ChurnTextAllocations(sheet);
    ExpectAnchorRetainsText(sheet, 4U, 5U, text);
  }
}

TEST(SheetSpillLifetime, RecommitCopiedCellsRetainsTextAfterClearingNewRegion) {
  for (const std::string_view text : kFirstSpillTexts) {
    Sheet sheet("Sheet1");
    ASSERT_TRUE(CommitTextAnchorSpill(sheet, text, 4U, 4U));

    const SpillRegion* original_region = sheet.spill_region_at_anchor(4U, 4U);
    ASSERT_NE(original_region, nullptr);
    std::vector<Value> copied_cells = original_region->cells;
    ASSERT_EQ(copied_cells.size(), 2U);

    ASSERT_TRUE(sheet.commit_spill(4U, 4U, 1U, 2U, std::move(copied_cells)));
    const SpillRegion* recommitted_region = sheet.spill_region_at_anchor(4U, 4U);
    ASSERT_NE(recommitted_region, nullptr);
    ASSERT_EQ(recommitted_region->cells.size(), 2U);
    ASSERT_TRUE(recommitted_region->cells[0].is_text());
    EXPECT_EQ(recommitted_region->cells[0].as_text(), text);
    EXPECT_EQ(recommitted_region->cells[1], Value::number(42.0));

    const Cell* anchor = sheet.cell_at(4U, 4U);
    ASSERT_NE(anchor, nullptr);
    ASSERT_NE(anchor->cached_text_owned, nullptr);
    EXPECT_NE(anchor->cached_text_owned->data(), recommitted_region->owned_strings.front().data());

    sheet.clear_spill(4U, 4U);
    EXPECT_EQ(sheet.spill_region_at_anchor(4U, 4U), nullptr);
    ChurnTextAllocations(sheet);
    ExpectAnchorRetainsText(sheet, 4U, 4U, text);
  }
}

TEST(SheetSpillLifetime, RecommitAliasedCellsThatExpandIntoBlockerRecordsFailure) {
  for (const std::string_view text : kFirstSpillTexts) {
    Sheet sheet("Sheet1");
    ASSERT_TRUE(CommitTextAnchorSpill(sheet, text, 4U, 4U));
    sheet.set_cell_value(4U, 6U, Value::number(7.0));

    const SpillRegion* original_region = sheet.spill_region_at_anchor(4U, 4U);
    ASSERT_NE(original_region, nullptr);
    std::vector<Value> expanded_cells = original_region->cells;
    expanded_cells.push_back(Value::number(99.0));

    EXPECT_FALSE(sheet.commit_spill(4U, 4U, 1U, 3U, std::move(expanded_cells)));
    EXPECT_EQ(sheet.spill_region_at_anchor(4U, 4U), nullptr);
    EXPECT_EQ(sheet.resolve_cell_value(4U, 4U), Value::error(ErrorCode::Spill));
    EXPECT_EQ(sheet.resolve_cell_value(4U, 5U), Value::blank());
    EXPECT_EQ(sheet.resolve_cell_value(4U, 6U), Value::number(7.0));

    const std::vector<BlockedSpillFootprint> blocked = sheet.blocked_spill_footprints();
    ASSERT_EQ(blocked.size(), 1U);
    EXPECT_EQ(blocked[0].anchor_row, 4U);
    EXPECT_EQ(blocked[0].anchor_col, 4U);
    EXPECT_EQ(blocked[0].rows, 1U);
    EXPECT_EQ(blocked[0].cols, 3U);
  }
}

}  // namespace
}  // namespace formulon
