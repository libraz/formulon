//
// Unit tests for the per-workbook volatile-cell tracker and its two static
// name classifiers.

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
  v.register_cell(a, VolatileKind::kValue);
  EXPECT_TRUE(v.contains(a));
  EXPECT_FALSE(v.contains(b));
  EXPECT_EQ(v.size(), 1u);
}

TEST(VolatileTracker, RegisterIsIdempotent) {
  VolatileTracker v;
  CellNodeId a = Make(0, 1, 1);
  v.register_cell(a, VolatileKind::kValue);
  v.register_cell(a, VolatileKind::kValue);
  v.register_cell(a, VolatileKind::kValue);
  EXPECT_EQ(v.size(), 1u);
}

TEST(VolatileTracker, UnregisterDropsEntry) {
  VolatileTracker v;
  CellNodeId a = Make(0, 0, 0);
  v.register_cell(a, VolatileKind::kValue);
  v.unregister_cell(a);
  EXPECT_FALSE(v.contains(a));
  EXPECT_TRUE(v.empty());
  // Unregistering an absent id is a no-op.
  v.unregister_cell(Make(9, 9, 9));
  EXPECT_TRUE(v.empty());
}

TEST(VolatileTracker, ClearEmptiesSet) {
  VolatileTracker v;
  v.register_cell(Make(0, 0, 0), VolatileKind::kValue);
  v.register_cell(Make(0, 0, 1), VolatileKind::kValue);
  v.register_cell(Make(1, 5, 5), VolatileKind::kValue);
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
    v.register_cell(id, VolatileKind::kValue);
  }

  std::unordered_set<CellNodeId, CellNodeIdHash> visited;
  v.for_each([&](CellNodeId id) { visited.insert(id); });

  EXPECT_EQ(visited.size(), ids.size());
  for (CellNodeId id : ids) {
    EXPECT_TRUE(visited.find(id) != visited.end());
  }
}

TEST(VolatileTracker, IsVolatileFunctionMatchesAllNine) {
  // Excel's nine volatile functions.
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

TEST(VolatileTracker, IsVolatileFunctionIsCaseInsensitive) {
  // A hand-typed `=now()` keeps its lowercase lexeme; the classifier must
  // still recognise it so the cell re-fires on recalc.
  EXPECT_TRUE(VolatileTracker::is_volatile_function("now"));
  EXPECT_TRUE(VolatileTracker::is_volatile_function("Now"));
  EXPECT_TRUE(VolatileTracker::is_volatile_function("rand"));
  EXPECT_TRUE(VolatileTracker::is_volatile_function("Today"));
  EXPECT_TRUE(VolatileTracker::is_volatile_function("Offset"));
  EXPECT_TRUE(VolatileTracker::is_volatile_function("indirect"));
  // Mixed-case non-volatiles still reject.
  EXPECT_FALSE(VolatileTracker::is_volatile_function("sum"));
  EXPECT_FALSE(VolatileTracker::is_volatile_function("nowx"));
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

TEST(VolatileTracker, DynamicReferenceClassIsSeparateFromValueClass) {
  VolatileTracker v;
  const CellNodeId value_cell = Make(0, 0, 0);
  const CellNodeId dynamic_cell = Make(0, 0, 1);
  v.register_cell(value_cell, VolatileKind::kValue);
  v.register_cell(dynamic_cell, VolatileKind::kDynamicReference);

  EXPECT_TRUE(v.contains(value_cell));
  EXPECT_TRUE(v.contains(dynamic_cell));
  EXPECT_FALSE(v.contains_dynamic_reference(value_cell));
  EXPECT_TRUE(v.contains_dynamic_reference(dynamic_cell));
  EXPECT_EQ(v.size(), 2u);
}

TEST(VolatileTracker, UnregisteredCellHasNoDynamicReference) {
  VolatileTracker v;
  EXPECT_FALSE(v.contains_dynamic_reference(Make(3, 3, 3)));
}

TEST(VolatileTracker, ReregisteringReplacesTheRecordedClass) {
  VolatileTracker v;
  const CellNodeId cell = Make(0, 2, 2);
  v.register_cell(cell, VolatileKind::kDynamicReference);
  ASSERT_TRUE(v.contains_dynamic_reference(cell));
  // A rewrite that drops the dynamic call must not leave the cell pinned
  // to the serial path forever.
  v.register_cell(cell, VolatileKind::kValue);
  EXPECT_TRUE(v.contains(cell));
  EXPECT_FALSE(v.contains_dynamic_reference(cell));
  EXPECT_EQ(v.size(), 1u);
}

TEST(VolatileTracker, DynamicReferenceClassifierMatchesOnlyResolvedAtEvalTime) {
  EXPECT_TRUE(VolatileTracker::is_dynamic_reference_function("INDIRECT"));
  EXPECT_TRUE(VolatileTracker::is_dynamic_reference_function("OFFSET"));
  EXPECT_TRUE(VolatileTracker::is_dynamic_reference_function("indirect"));
  EXPECT_TRUE(VolatileTracker::is_dynamic_reference_function("Offset"));
  // The value-volatile functions read the host environment or a spelled-out
  // reference; both are covered by the recorded edges.
  EXPECT_FALSE(VolatileTracker::is_dynamic_reference_function("NOW"));
  EXPECT_FALSE(VolatileTracker::is_dynamic_reference_function("TODAY"));
  EXPECT_FALSE(VolatileTracker::is_dynamic_reference_function("RAND"));
  EXPECT_FALSE(VolatileTracker::is_dynamic_reference_function("RANDBETWEEN"));
  EXPECT_FALSE(VolatileTracker::is_dynamic_reference_function("RANDARRAY"));
  EXPECT_FALSE(VolatileTracker::is_dynamic_reference_function("INFO"));
  EXPECT_FALSE(VolatileTracker::is_dynamic_reference_function("CELL"));
  EXPECT_FALSE(VolatileTracker::is_dynamic_reference_function("SUM"));
  EXPECT_FALSE(VolatileTracker::is_dynamic_reference_function(""));
  EXPECT_FALSE(VolatileTracker::is_dynamic_reference_function("INDIRECTX"));
  EXPECT_FALSE(VolatileTracker::is_dynamic_reference_function("OFFSE"));
}

TEST(VolatileTracker, EveryDynamicReferenceFunctionIsAlsoVolatile) {
  for (const std::string_view name : {"INDIRECT", "OFFSET", "indirect", "Offset"}) {
    EXPECT_TRUE(VolatileTracker::is_dynamic_reference_function(name)) << name;
    EXPECT_TRUE(VolatileTracker::is_volatile_function(name)) << name;
  }
}

TEST(VolatileTracker, CrossSheetCellsAreDistinct) {
  VolatileTracker v;
  CellNodeId a = Make(0, 1, 1);
  CellNodeId b = Make(1, 1, 1);  // same row/col, different sheet
  v.register_cell(a, VolatileKind::kValue);
  EXPECT_TRUE(v.contains(a));
  EXPECT_FALSE(v.contains(b));
  v.register_cell(b, VolatileKind::kValue);
  EXPECT_EQ(v.size(), 2u);
}

}  // namespace
}  // namespace formulon::eval
