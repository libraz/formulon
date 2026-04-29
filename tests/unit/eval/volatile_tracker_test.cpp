// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for the per-workbook volatile-cell tracker and the
// `is_volatile_function` static classifier.

#include "eval/volatile_tracker.h"

#include <cstdint>
#include <string_view>
#include <unordered_set>
#include <vector>

#include "eval/dep_graph.h"
#include "gtest/gtest.h"

namespace formulon::eval {
namespace {

CellNodeId Make(std::uint16_t sheet, std::uint32_t row, std::uint32_t col) {
  return CellNodeId{sheet, row, col};
}

TEST(VolatileTracker, EmptyByDefault) {
  VolatileTracker v;
  EXPECT_TRUE(v.empty());
  EXPECT_EQ(v.size(), 0u);
  EXPECT_FALSE(v.contains(Make(0, 0, 0)));
}

TEST(VolatileTracker, RegisterAndContains) {
  VolatileTracker v;
  CellNodeId a = Make(0, 0, 0);
  CellNodeId b = Make(0, 0, 1);
  v.register_cell(a);
  EXPECT_TRUE(v.contains(a));
  EXPECT_FALSE(v.contains(b));
  EXPECT_EQ(v.size(), 1u);
}

TEST(VolatileTracker, RegisterIsIdempotent) {
  VolatileTracker v;
  CellNodeId a = Make(0, 1, 1);
  v.register_cell(a);
  v.register_cell(a);
  v.register_cell(a);
  EXPECT_EQ(v.size(), 1u);
}

TEST(VolatileTracker, UnregisterDropsEntry) {
  VolatileTracker v;
  CellNodeId a = Make(0, 0, 0);
  v.register_cell(a);
  v.unregister_cell(a);
  EXPECT_FALSE(v.contains(a));
  EXPECT_TRUE(v.empty());
  // Unregistering an absent id is a no-op.
  v.unregister_cell(Make(9, 9, 9));
  EXPECT_TRUE(v.empty());
}

TEST(VolatileTracker, ClearEmptiesSet) {
  VolatileTracker v;
  v.register_cell(Make(0, 0, 0));
  v.register_cell(Make(0, 0, 1));
  v.register_cell(Make(1, 5, 5));
  EXPECT_EQ(v.size(), 3u);
  v.clear();
  EXPECT_TRUE(v.empty());
}

TEST(VolatileTracker, ForEachVisitsEveryRegisteredCell) {
  VolatileTracker v;
  std::vector<CellNodeId> ids = {
      Make(0, 0, 0),
      Make(0, 1, 2),
      Make(1, 5, 5),
      Make(2, 9, 0),
  };
  for (CellNodeId id : ids) {
    v.register_cell(id);
  }

  std::unordered_set<CellNodeId, CellNodeIdHash> visited;
  v.for_each([&](CellNodeId id) { visited.insert(id); });

  EXPECT_EQ(visited.size(), ids.size());
  for (CellNodeId id : ids) {
    EXPECT_TRUE(visited.find(id) != visited.end());
  }
}

TEST(VolatileTracker, IsVolatileFunctionMatchesAllNine) {
  // Excel's nine volatile functions per backup/plans/02-calc-engine.md §2.7.2.
  EXPECT_TRUE(VolatileTracker::is_volatile_function("NOW"));
  EXPECT_TRUE(VolatileTracker::is_volatile_function("TODAY"));
  EXPECT_TRUE(VolatileTracker::is_volatile_function("RAND"));
  EXPECT_TRUE(VolatileTracker::is_volatile_function("RANDBETWEEN"));
  EXPECT_TRUE(VolatileTracker::is_volatile_function("RANDARRAY"));
  EXPECT_TRUE(VolatileTracker::is_volatile_function("OFFSET"));
  EXPECT_TRUE(VolatileTracker::is_volatile_function("INDIRECT"));
  EXPECT_TRUE(VolatileTracker::is_volatile_function("INFO"));
  EXPECT_TRUE(VolatileTracker::is_volatile_function("CELL"));
}

TEST(VolatileTracker, IsVolatileFunctionRejectsNonVolatiles) {
  EXPECT_FALSE(VolatileTracker::is_volatile_function("SUM"));
  EXPECT_FALSE(VolatileTracker::is_volatile_function("ABS"));
  EXPECT_FALSE(VolatileTracker::is_volatile_function("IF"));
  EXPECT_FALSE(VolatileTracker::is_volatile_function("XLOOKUP"));
  EXPECT_FALSE(VolatileTracker::is_volatile_function("VLOOKUP"));
}

TEST(VolatileTracker, IsVolatileFunctionIsCaseSensitive) {
  // Names are expected to already be canonicalised to ASCII uppercase by
  // the caller — lowercase / mixed-case must not match.
  EXPECT_FALSE(VolatileTracker::is_volatile_function("now"));
  EXPECT_FALSE(VolatileTracker::is_volatile_function("Now"));
  EXPECT_FALSE(VolatileTracker::is_volatile_function("rand"));
  EXPECT_FALSE(VolatileTracker::is_volatile_function("Today"));
}

TEST(VolatileTracker, IsVolatileFunctionRejectsEmptyAndJunk) {
  EXPECT_FALSE(VolatileTracker::is_volatile_function(""));
  EXPECT_FALSE(VolatileTracker::is_volatile_function(" "));
  EXPECT_FALSE(VolatileTracker::is_volatile_function("NOWX"));
  EXPECT_FALSE(VolatileTracker::is_volatile_function("RAN"));
  EXPECT_FALSE(VolatileTracker::is_volatile_function("OFFSE"));
  EXPECT_FALSE(VolatileTracker::is_volatile_function("X"));
}

TEST(VolatileTracker, IsVolatileFunctionRejectsPrefixesAndSuffixes) {
  // Defense against substring-match bugs.
  EXPECT_FALSE(VolatileTracker::is_volatile_function("RANDOM"));
  EXPECT_FALSE(VolatileTracker::is_volatile_function("RANDXYZ"));
  EXPECT_FALSE(VolatileTracker::is_volatile_function("CELLA"));
  EXPECT_FALSE(VolatileTracker::is_volatile_function("INFOX"));
}

TEST(VolatileTracker, CrossSheetCellsAreDistinct) {
  VolatileTracker v;
  CellNodeId a = Make(0, 1, 1);
  CellNodeId b = Make(1, 1, 1);  // same row/col, different sheet
  v.register_cell(a);
  EXPECT_TRUE(v.contains(a));
  EXPECT_FALSE(v.contains(b));
  v.register_cell(b);
  EXPECT_EQ(v.size(), 2u);
}

}  // namespace
}  // namespace formulon::eval
