// Sheet spill cost-shape and population-count tests.

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
/// Fastest of `kTrials` runs of `work`, in nanoseconds.
template <typename Fn>
double FastestRunNanos(Fn&& work) {
  constexpr int kTrials = 5;
  double best = 0.0;
  for (int trial = 0; trial < kTrials; ++trial) {
    const auto start = std::chrono::steady_clock::now();
    work();
    const double elapsed = std::chrono::duration<double, std::nano>(std::chrono::steady_clock::now() - start).count();
    if (trial == 0 || elapsed < best) {
      best = elapsed;
    }
  }
  return best;
}

/// How much slower the loaded shape may be before the cost is following the
/// wrong quantity. The shapes below differ by two orders of magnitude in the
/// quantity under test, so anything near 1 passes and a linear dependence
/// fails by a wide margin — the gap is what keeps this from flaking.
constexpr double kMaxCostRatio = 10.0;
TEST(SheetReadRange, CostDoesNotFollowTheNumberOfSpillRegionsOnTheSheet) {
  // A rectangle of empty cells, read on two sheets that differ only in how
  // many spill regions sit far away from it. Consulting the spill table per
  // coordinate makes the second sheet hundreds of times slower for a result
  // that is identical.
  constexpr std::uint32_t kLastRow = 199U;
  constexpr std::uint32_t kLastCol = 99U;
  constexpr std::uint32_t kManyRegions = 400U;

  const auto build = [](std::uint32_t regions) {
    Sheet sheet("Sheet1");
    for (std::uint32_t i = 0; i < regions; ++i) {
      // Parked well below and to the right of the read rectangle.
      EXPECT_TRUE(sheet.commit_spill(1000U + i, 500U, 1U, 1U, {Value::number(i)}));
    }
    return sheet;
  };
  Sheet few = build(1U);
  Sheet many = build(kManyRegions);

  const auto read = [&](Sheet& sheet) {
    return [&sheet]() {
      Arena arena;
      std::vector<Value> out;
      std::vector<std::size_t> formula_indices;
      sheet.read_range(0U, kLastRow, 0U, kLastCol, arena, out, formula_indices);
      EXPECT_EQ(out.size(), static_cast<std::size_t>(kLastRow + 1U) * (kLastCol + 1U));
    };
  };

  const double one_region = FastestRunNanos(read(few));
  const double many_regions = FastestRunNanos(read(many));
  EXPECT_LT(many_regions, one_region * kMaxCostRatio)
      << "read_range cost follows the spill-table size: " << one_region << "ns with one region, " << many_regions
      << "ns with " << kManyRegions;
}
TEST(SheetSpillTest, CellCountDoesNotFollowTheSpillArea) {
  // A spill is one rectangle with one payload, so counting the coordinates it
  // covers is arithmetic. Walking them instead puts a lookup per spilled cell
  // under the sheet lock, which a whole-column spill turns into a million.
  constexpr std::uint32_t kSmallSide = 4U;
  constexpr std::uint32_t kLargeSide = 400U;

  const auto build = [](std::uint32_t side) {
    Sheet sheet("Sheet1");
    std::vector<Value> cells;
    cells.reserve(static_cast<std::size_t>(side) * side);
    for (std::size_t i = 0; i < static_cast<std::size_t>(side) * side; ++i) {
      cells.push_back(Value::number(static_cast<double>(i)));
    }
    EXPECT_TRUE(sheet.commit_spill(0U, 0U, side, side, std::move(cells)));
    return sheet;
  };
  Sheet small = build(kSmallSide);
  Sheet large = build(kLargeSide);

  EXPECT_EQ(small.cell_count(), static_cast<std::size_t>(kSmallSide) * kSmallSide);
  EXPECT_EQ(large.cell_count(), static_cast<std::size_t>(kLargeSide) * kLargeSide);

  const auto count = [](Sheet& sheet) {
    return [&sheet]() {
      for (int i = 0; i < 20; ++i) {
        EXPECT_GT(sheet.cell_count(), 0U);
      }
    };
  };

  const double small_area = FastestRunNanos(count(small));
  const double large_area = FastestRunNanos(count(large));
  EXPECT_LT(large_area, small_area * kMaxCostRatio)
      << "cell_count cost follows the spill area: " << small_area << "ns over " << kSmallSide * kSmallSide << " cells, "
      << large_area << "ns over " << kLargeSide * kLargeSide;
}
TEST(SheetSpillTest, CellCountKeepsCountingPhantomsThatShareAMaterialisedSlot) {
  Sheet s("Sheet1");
  // Writing A1 and then E1 materialises the whole run A1..E1, three of them
  // implicitly blank. A spill over A1:C1 then covers slots that the stored
  // count already includes.
  s.set_cell_value(0U, 0U, Value::blank());
  s.set_cell_value(0U, 4U, Value::number(9.0));
  ASSERT_EQ(s.cell_count(), 5U);
  ASSERT_TRUE(s.commit_spill(0U, 0U, 1U, 3U, {Value::number(1.0), Value::number(2.0), Value::number(3.0)}));
  EXPECT_EQ(s.cell_count(), 5U) << "phantoms that coincide with materialised slots must not be counted twice";

  // A spill running past the end of the row's materialised run adds only the
  // coordinates that are not already slots.
  Sheet t("Sheet1");
  ASSERT_TRUE(t.commit_spill(
      0U, 0U, 1U, 5U,
      {Value::number(1.0), Value::number(2.0), Value::number(3.0), Value::number(4.0), Value::number(5.0)}));
  EXPECT_EQ(t.cell_count(), 5U);
}

}  // namespace
}  // namespace formulon
