// Unit tests for the compact rectangle dependency index.

#include "eval/range_dep_index.h"

#include <algorithm>
#include <cstdint>
#include <vector>

#include "gtest/gtest.h"

namespace formulon::eval {
namespace {

CellNodeId Make(std::uint16_t sheet, std::uint32_t row, std::uint32_t col) {
  return CellNodeId{sheet, row, col};
}

std::vector<std::uint32_t> RangesAt(const RangeDepIndex& index, CellNodeId cell) {
  std::vector<std::uint32_t> ranges;
  index.for_each_range_covering(cell, [&](std::uint32_t range_id) { ranges.push_back(range_id); });
  std::sort(ranges.begin(), ranges.end());
  return ranges;
}

std::vector<CellNodeId> OwnersAt(const RangeDepIndex& index, CellNodeId cell) {
  std::vector<CellNodeId> owners;
  index.for_each_owner_covering(cell, [&](CellNodeId owner) { owners.push_back(owner); });
  std::sort(owners.begin(), owners.end(), CellNodeIdOrder{});
  return owners;
}

std::vector<CellNodeId> NewOwnersAt(RangeDepIndex& index, CellNodeId cell, std::uint64_t generation) {
  std::vector<CellNodeId> owners;
  index.for_each_new_owner_covering(cell, generation, [&](CellNodeId owner) { owners.push_back(owner); });
  return owners;
}

TEST(RangeDepIndex, ReleasingMiddleMultibandRangeKeepsMovedPositionsValid) {
  RangeDepIndex index;
  const CellNodeId owner_a = Make(0, 1, 0);
  const CellNodeId owner_b = Make(0, 2, 0);
  const CellNodeId owner_c = Make(0, 3, 0);
  const CellRangeDependency range_a{1U, 0U, 8191U, 0U, 4U};
  const CellRangeDependency range_b{1U, 4096U, 12287U, 0U, 4U};
  const CellRangeDependency range_c{1U, 0U, 16383U, 0U, 4U};

  const auto a = index.add(owner_a, range_a);
  const auto b = index.add(owner_b, range_b);
  const auto c = index.add(owner_c, range_c);
  ASSERT_TRUE(a.is_new);
  ASSERT_TRUE(b.is_new);
  ASSERT_TRUE(c.is_new);
  ASSERT_EQ(index.distinct_range_count(), 3U);

  EXPECT_EQ(index.erase_owner(owner_b), std::vector<std::uint32_t>{b.range_id});
  EXPECT_EQ(index.distinct_range_count(), 2U);
  EXPECT_EQ(RangesAt(index, Make(1, 7000, 2)), (std::vector<std::uint32_t>{a.range_id, c.range_id}));
  EXPECT_EQ(RangesAt(index, Make(1, 9000, 2)), std::vector<std::uint32_t>{c.range_id});
  EXPECT_EQ(RangesAt(index, Make(1, 15000, 2)), std::vector<std::uint32_t>{c.range_id});
  EXPECT_TRUE(RangesAt(index, Make(1, 20000, 2)).empty());
  EXPECT_EQ(OwnersAt(index, Make(1, 9000, 2)), std::vector<CellNodeId>{owner_c});

  // Releasing the range swapped into the removed middle slots exercises the
  // updated per-band positions in both shared buckets.
  EXPECT_EQ(index.erase_owner(owner_c), std::vector<std::uint32_t>{c.range_id});
  EXPECT_EQ(index.distinct_range_count(), 1U);
  EXPECT_EQ(RangesAt(index, Make(1, 7000, 2)), std::vector<std::uint32_t>{a.range_id});
  EXPECT_TRUE(RangesAt(index, Make(1, 9000, 2)).empty());
}

TEST(RangeDepIndex, ReusedRangeIdResetsCoverageAndOwnerGeneration) {
  RangeDepIndex index;
  const CellNodeId old_owner = Make(0, 10, 0);
  const CellNodeId new_owner = Make(0, 20, 0);
  const CellNodeId second_owner = Make(0, 21, 0);
  const CellNodeId third_owner = Make(0, 22, 0);
  const CellRangeDependency old_range{3U, 0U, 8191U, 1U, 1U};
  const CellRangeDependency new_range{3U, 16384U, 20479U, 1U, 1U};
  constexpr std::uint64_t kGeneration = 17U;

  const auto old = index.add(old_owner, old_range);
  ASSERT_TRUE(old.is_new);
  EXPECT_EQ(NewOwnersAt(index, Make(3, 5, 1), kGeneration), std::vector<CellNodeId>{old_owner});
  EXPECT_TRUE(NewOwnersAt(index, Make(3, 5, 1), kGeneration).empty());
  EXPECT_EQ(index.erase_owner(old_owner), std::vector<std::uint32_t>{old.range_id});
  EXPECT_TRUE(index.empty());

  const auto reused = index.add(new_owner, new_range);
  EXPECT_TRUE(reused.is_new);
  EXPECT_EQ(reused.range_id, old.range_id);
  EXPECT_TRUE(RangesAt(index, Make(3, 5, 1)).empty());
  EXPECT_EQ(RangesAt(index, Make(3, 17000, 1)), std::vector<std::uint32_t>{reused.range_id});
  EXPECT_EQ(NewOwnersAt(index, Make(3, 17000, 1), kGeneration), std::vector<CellNodeId>{new_owner});

  const auto shared = index.add(second_owner, new_range);
  EXPECT_FALSE(shared.is_new);
  EXPECT_EQ(shared.range_id, reused.range_id);
  EXPECT_EQ(NewOwnersAt(index, Make(3, 17000, 1), kGeneration), std::vector<CellNodeId>{second_owner});
  EXPECT_TRUE(NewOwnersAt(index, Make(3, 17000, 1), kGeneration).empty());

  // Swap-removing an already visited owner leaves the remaining prefix
  // visited, while an owner added afterward is still surfaced this generation.
  EXPECT_TRUE(index.erase_owner(new_owner).empty());
  EXPECT_TRUE(NewOwnersAt(index, Make(3, 17000, 1), kGeneration).empty());
  EXPECT_FALSE(index.add(third_owner, new_range).is_new);
  EXPECT_EQ(NewOwnersAt(index, Make(3, 17000, 1), kGeneration), std::vector<CellNodeId>{third_owner});
  EXPECT_EQ(OwnersAt(index, Make(3, 17000, 1)), (std::vector<CellNodeId>{second_owner, third_owner}));
}

}  // namespace
}  // namespace formulon::eval
