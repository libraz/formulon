// Tests for the hyperlink list of `Sheet`: identity by exact rectangle, list
// order, containment reads and whole-link removal by intersection.

#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "merge_range.h"
#include "sheet.h"
#include "sheet_metadata.h"

namespace formulon {
namespace {

Hyperlink MakeLink(std::uint32_t row, std::uint32_t col, std::uint32_t last_row, std::uint32_t last_col,
                   std::string target) {
  Hyperlink link;
  link.row = row;
  link.col = col;
  link.last_row = last_row;
  link.last_col = last_col;
  link.target = std::move(target);
  return link;
}

std::vector<std::string> Targets(const Sheet& sheet) {
  std::vector<std::string> out;
  for (const Hyperlink& link : sheet.hyperlinks()) {
    out.push_back(link.target);
  }
  return out;
}

TEST(SheetHyperlinks, CoversAndIntersectsAreInclusive) {
  const Hyperlink link = MakeLink(1, 1, 3, 3, "x");
  EXPECT_TRUE(link.covers(1, 1));
  EXPECT_TRUE(link.covers(3, 3));
  EXPECT_FALSE(link.covers(0, 1));
  EXPECT_FALSE(link.covers(1, 4));
  EXPECT_TRUE(link.intersects(3, 3, 9, 9));
  EXPECT_TRUE(link.intersects(0, 0, 1, 1));
  EXPECT_FALSE(link.intersects(4, 0, 9, 9));
  EXPECT_FALSE(link.intersects(0, 4, 9, 9));
  EXPECT_TRUE(link.same_range(MakeLink(1, 1, 3, 3, "other")));
  EXPECT_FALSE(link.same_range(MakeLink(1, 1, 3, 2, "x")));
}

TEST(SheetHyperlinks, SetAppendsANewRangeAndKeepsOverlappingOnesWhole) {
  Sheet sheet("Sheet1");
  sheet.set_hyperlink(MakeLink(0, 0, 2, 2, "wide"));
  sheet.set_hyperlink(MakeLink(1, 1, 1, 1, "cell"));
  sheet.set_hyperlink(MakeLink(1, 1, 3, 3, "partial"));
  sheet.set_hyperlink(MakeLink(0, 0, 1, 1, "sub"));
  EXPECT_EQ(Targets(sheet), (std::vector<std::string>{"wide", "cell", "partial", "sub"}));
  EXPECT_EQ(sheet.hyperlinks()[0].last_row, 2U);
  EXPECT_EQ(sheet.hyperlinks()[2].last_col, 3U);
}

TEST(SheetHyperlinks, SetReplacesTheSameRangeInPlaceWithAllFields) {
  Sheet sheet("Sheet1");
  sheet.set_hyperlink(MakeLink(0, 0, 0, 0, "a"));
  Hyperlink original = MakeLink(1, 0, 1, 0, "b");
  original.tooltip = "old tip";
  original.rid = "rId7";
  sheet.set_hyperlink(std::move(original));
  sheet.set_hyperlink(MakeLink(2, 0, 2, 0, "c"));

  Hyperlink replacement = MakeLink(1, 0, 1, 0, "");
  replacement.location = "Sheet1!A1";
  sheet.set_hyperlink(std::move(replacement));

  ASSERT_EQ(sheet.hyperlinks().size(), 3U);
  EXPECT_EQ(Targets(sheet), (std::vector<std::string>{"a", "", "c"}));
  const Hyperlink& replaced = sheet.hyperlinks()[1];
  EXPECT_EQ(replaced.location, "Sheet1!A1");
  EXPECT_TRUE(replaced.tooltip.empty());
  EXPECT_TRUE(replaced.rid.empty());
}

TEST(SheetHyperlinks, SetReplacesTheFirstOfADuplicatePairAndDropsTheLaterOne) {
  Sheet sheet("Sheet1");
  sheet.mutable_hyperlinks().push_back(MakeLink(0, 0, 0, 0, "dup1"));
  sheet.mutable_hyperlinks().push_back(MakeLink(4, 4, 4, 4, "mid"));
  sheet.mutable_hyperlinks().push_back(MakeLink(0, 0, 0, 0, "dup2"));
  sheet.mutable_hyperlinks().push_back(MakeLink(0, 0, 0, 0, "dup3"));
  ASSERT_EQ(sheet.hyperlinks().size(), 4U);

  sheet.set_hyperlink(MakeLink(0, 0, 0, 0, "new"));
  EXPECT_EQ(Targets(sheet), (std::vector<std::string>{"new", "mid"}));
}

TEST(SheetHyperlinks, MutableListKeepsEveryEntryAsGiven) {
  Sheet sheet("Sheet1");
  sheet.mutable_hyperlinks().push_back(MakeLink(0, 0, 0, 0, "a"));
  sheet.mutable_hyperlinks().push_back(MakeLink(0, 0, 0, 0, "b"));
  sheet.mutable_hyperlinks().insert(sheet.mutable_hyperlinks().begin(), MakeLink(0, 0, 0, 0, "c"));
  EXPECT_EQ(Targets(sheet), (std::vector<std::string>{"c", "a", "b"}));
}

TEST(SheetHyperlinks, HyperlinkAtReadsTheEarliestListEntryContainingTheCell) {
  Sheet sheet("Sheet1");
  sheet.set_hyperlink(MakeLink(0, 0, 2, 2, "wide"));
  sheet.set_hyperlink(MakeLink(1, 1, 1, 1, "cell"));
  ASSERT_NE(sheet.hyperlink_at(1, 1), nullptr);
  EXPECT_EQ(sheet.hyperlink_at(1, 1)->target, "wide");
  ASSERT_NE(sheet.hyperlink_at(2, 2), nullptr);
  EXPECT_EQ(sheet.hyperlink_at(2, 2)->target, "wide");
  EXPECT_EQ(sheet.hyperlink_at(3, 3), nullptr);
  EXPECT_EQ(sheet.hyperlink_at(0, 3), nullptr);

  Sheet reversed("Sheet1");
  reversed.set_hyperlink(MakeLink(1, 1, 1, 1, "cell"));
  reversed.set_hyperlink(MakeLink(0, 0, 2, 2, "wide"));
  EXPECT_EQ(reversed.hyperlink_at(1, 1)->target, "cell");
  EXPECT_EQ(reversed.hyperlink_at(0, 0)->target, "wide");
}

TEST(SheetHyperlinks, HyperlinkAtOnAnEmptyListIsNull) {
  const Sheet sheet("Sheet1");
  EXPECT_EQ(sheet.hyperlink_at(0, 0), nullptr);
  EXPECT_EQ(sheet.hyperlink_at(Sheet::kMaxRows - 1U, Sheet::kMaxCols - 1U), nullptr);
}

TEST(SheetHyperlinks, RemoveIntersectingRemovesEveryTouchedLinkWholeInListOrder) {
  Sheet sheet("Sheet1");
  sheet.set_hyperlink(MakeLink(0, 0, 2, 2, "wide"));
  sheet.set_hyperlink(MakeLink(1, 1, 1, 1, "cell"));
  sheet.set_hyperlink(MakeLink(9, 9, 9, 9, "far"));

  const std::vector<Hyperlink> removed = sheet.remove_hyperlinks_intersecting(MergeRange{1, 1, 1, 1});
  ASSERT_EQ(removed.size(), 2U);
  EXPECT_EQ(removed[0].target, "wide");
  EXPECT_EQ(removed[0].last_row, 2U);
  EXPECT_EQ(removed[1].target, "cell");
  EXPECT_EQ(Targets(sheet), (std::vector<std::string>{"far"}));
}

TEST(SheetHyperlinks, RemoveIntersectingKeepsTheOtherLinkWhenOnlyOneIsTouched) {
  Sheet sheet("Sheet1");
  sheet.set_hyperlink(MakeLink(0, 0, 2, 2, "wide"));
  sheet.set_hyperlink(MakeLink(1, 1, 1, 1, "cell"));
  const std::vector<Hyperlink> removed = sheet.remove_hyperlinks_intersecting(MergeRange{0, 0, 0, 0});
  ASSERT_EQ(removed.size(), 1U);
  EXPECT_EQ(removed[0].target, "wide");
  EXPECT_EQ(Targets(sheet), (std::vector<std::string>{"cell"}));
}

TEST(SheetHyperlinks, RemoveIntersectingNothingReturnsEmptyAndChangesNothing) {
  Sheet sheet("Sheet1");
  sheet.set_hyperlink(MakeLink(0, 0, 1, 1, "a"));
  sheet.set_hyperlink(MakeLink(5, 5, 5, 5, "b"));
  EXPECT_TRUE(sheet.remove_hyperlinks_intersecting(MergeRange{2, 2, 4, 4}).empty());
  EXPECT_EQ(Targets(sheet), (std::vector<std::string>{"a", "b"}));
  EXPECT_TRUE(Sheet("Empty").remove_hyperlinks_intersecting(MergeRange{0, 0, 9, 9}).empty());
}

TEST(SheetHyperlinks, RemoveIntersectingTakesEveryDuplicateOfARange) {
  Sheet sheet("Sheet1");
  sheet.mutable_hyperlinks().push_back(MakeLink(0, 0, 0, 0, "dup1"));
  sheet.mutable_hyperlinks().push_back(MakeLink(3, 3, 3, 3, "keep"));
  sheet.mutable_hyperlinks().push_back(MakeLink(0, 0, 0, 0, "dup2"));
  const std::vector<Hyperlink> removed = sheet.remove_hyperlinks_intersecting(MergeRange{0, 0, 0, 0});
  ASSERT_EQ(removed.size(), 2U);
  EXPECT_EQ(removed[0].target, "dup1");
  EXPECT_EQ(removed[1].target, "dup2");
  EXPECT_EQ(Targets(sheet), (std::vector<std::string>{"keep"}));
}

TEST(SheetHyperlinks, RemoveIntersectingTouchesOnlyAtTheSharedEdge) {
  Sheet sheet("Sheet1");
  sheet.set_hyperlink(MakeLink(0, 0, 2, 2, "a"));
  // The range starts one row below the link: no shared cell.
  EXPECT_TRUE(sheet.remove_hyperlinks_intersecting(MergeRange{3, 0, 5, 2}).empty());
  // The range shares exactly the corner cell.
  EXPECT_EQ(sheet.remove_hyperlinks_intersecting(MergeRange{2, 2, 5, 5}).size(), 1U);
  EXPECT_TRUE(sheet.hyperlinks().empty());
}

}  // namespace
}  // namespace formulon
