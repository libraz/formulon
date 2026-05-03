// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Unit tests for the recalc-pass-scoped dirty-cell set.

#include "eval/dirty_set.h"

#include <cstdint>
#include <unordered_set>
#include <vector>

#include "eval/dep_graph.h"
#include "gtest/gtest.h"

namespace formulon::eval {
namespace {

CellNodeId Make(std::uint16_t sheet, std::uint32_t row, std::uint32_t col) {
  return CellNodeId{sheet, row, col};
}

TEST(DirtySet, EmptyByDefault) {
  DirtySet d;
  EXPECT_TRUE(d.empty());
  EXPECT_EQ(d.size(), 0u);
  EXPECT_FALSE(d.contains(Make(0, 0, 0)));
}

TEST(DirtySet, MarkAndContains) {
  DirtySet d;
  CellNodeId a = Make(0, 0, 0);
  CellNodeId b = Make(0, 0, 1);
  d.mark(a);
  EXPECT_TRUE(d.contains(a));
  EXPECT_FALSE(d.contains(b));
  EXPECT_EQ(d.size(), 1u);
  EXPECT_FALSE(d.empty());
}

TEST(DirtySet, MarkIsIdempotent) {
  DirtySet d;
  CellNodeId a = Make(2, 3, 4);
  d.mark(a);
  d.mark(a);
  d.mark(a);
  EXPECT_EQ(d.size(), 1u);
  EXPECT_TRUE(d.contains(a));
}

TEST(DirtySet, MultipleMarksThenClear) {
  DirtySet d;
  d.mark(Make(0, 0, 0));
  d.mark(Make(0, 0, 1));
  d.mark(Make(1, 5, 5));
  EXPECT_EQ(d.size(), 3u);

  d.clear();
  EXPECT_TRUE(d.empty());
  EXPECT_EQ(d.size(), 0u);
  EXPECT_FALSE(d.contains(Make(0, 0, 0)));
}

TEST(DirtySet, Unmark) {
  DirtySet d;
  CellNodeId a = Make(0, 0, 0);
  CellNodeId b = Make(0, 0, 1);
  d.mark(a);
  d.mark(b);
  d.unmark(a);
  EXPECT_FALSE(d.contains(a));
  EXPECT_TRUE(d.contains(b));
  EXPECT_EQ(d.size(), 1u);
  // Unmarking an absent id is a no-op.
  d.unmark(Make(7, 7, 7));
  EXPECT_EQ(d.size(), 1u);
}

TEST(DirtySet, ForEachVisitsEveryMarkedCellExactlyOnce) {
  DirtySet d;
  std::vector<CellNodeId> ids = {
      Make(0, 0, 0),
      Make(0, 1, 2),
      Make(1, 5, 5),
      Make(2, 9, 0),
  };
  for (CellNodeId id : ids) {
    d.mark(id);
  }

  std::unordered_set<CellNodeId, CellNodeIdHash> visited;
  d.for_each([&](CellNodeId id) { visited.insert(id); });

  EXPECT_EQ(visited.size(), ids.size());
  for (CellNodeId id : ids) {
    EXPECT_TRUE(visited.find(id) != visited.end());
  }
}

TEST(DirtySet, CrossSheetCellsAreDistinct) {
  DirtySet d;
  CellNodeId a = Make(0, 1, 1);
  CellNodeId b = Make(1, 1, 1);  // same row/col, different sheet
  d.mark(a);
  EXPECT_TRUE(d.contains(a));
  EXPECT_FALSE(d.contains(b));
  d.mark(b);
  EXPECT_EQ(d.size(), 2u);
}

}  // namespace
}  // namespace formulon::eval
