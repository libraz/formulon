// Pivot SUM / AVERAGE apply Excel's cancellation snap per step, PRODUCT
// flushes a subnormal step to 0, and a Sum total is a running sum of the
// rendered cells (snapped along one axis, plain across both). Expected values
// were read from Mac Excel 365 after refreshing a pivot over the same records.

#include <cstddef>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_evaluator.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot_evaluator_fixtures.h"
#include "value.h"

namespace formulon::pivot {
namespace {

using test::build_sum_amount_table;
using test::owned_text;
using test::row_index;

struct Record {
  const char* region;
  const char* product;
  double amount;
};

PivotCache build_cache(const std::vector<Record>& records) {
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Region", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Product", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});
  for (const Record& r : records) {
    PivotCacheRecord rec;
    rec.cells.push_back(owned_text(cache, r.region));
    rec.cells.push_back(owned_text(cache, r.product));
    rec.cells.push_back(Value::number(r.amount));
    cache.mutable_records().push_back(std::move(rec));
  }
  return cache;
}

PivotResult evaluate_with(const PivotCache& cache, PivotTable table, Aggregation aggregation) {
  table.mutable_data_fields()[0].aggregation = aggregation;
  table.set_grand_totals(/*rows=*/true, /*cols=*/true);
  auto r_or = evaluate(table, cache);
  EXPECT_TRUE(static_cast<bool>(r_or));
  return std::move(r_or).value();
}

const std::vector<Record> kRowOnlyRecords = {
    {"North", "A", 0.5}, {"North", "B", -0.4}, {"North", "A", -0.1}, {"South", "A", 1e-300}, {"East", "A", 1e-10},
};

TEST(PivotSumSnap, CancellingGroupSumsToZero) {
  const PivotCache cache = build_cache(kRowOnlyRecords);
  const PivotResult r = evaluate_with(cache, build_sum_amount_table({0}, {}), Aggregation::Sum);
  const std::size_t north = row_index(r, "North");
  ASSERT_LT(north, r.values.size());
  ASSERT_TRUE(r.values[north][0][0].is_number());
  EXPECT_EQ(r.values[north][0][0].as_number(), 0.0);
}

TEST(PivotSumSnap, CancellingGroupAveragesToZero) {
  const PivotCache cache = build_cache(kRowOnlyRecords);
  const PivotResult r = evaluate_with(cache, build_sum_amount_table({0}, {}), Aggregation::Average);
  const std::size_t north = row_index(r, "North");
  ASSERT_LT(north, r.values.size());
  ASSERT_TRUE(r.values[north][0][0].is_number());
  EXPECT_EQ(r.values[north][0][0].as_number(), 0.0);
}

TEST(PivotSumSnap, GrandTotalAddsRenderedLeafCells) {
  // East 1e-10, North 0 (snapped), South 1e-300: the cancelled North residue never reaches the total.
  const PivotCache cache = build_cache(kRowOnlyRecords);
  const PivotResult r = evaluate_with(cache, build_sum_amount_table({0}, {}), Aggregation::Sum);
  ASSERT_TRUE(r.grand_total.is_number());
  EXPECT_EQ(r.grand_total.as_number(), 1e-10);
}

TEST(PivotSumSnap, RowOnlyGrandTotalSnapsEachLeafCellStep) {
  // Records in item order give 0.20000000010000002, in source order 1e-10, raw cells 2.2.
  const PivotCache cache = build_cache({
      {"East", "A", -0.3},
      {"North", "A", 0.5},
      {"South", "A", 0.1},
      {"East", "A", -1e16},
      {"North", "A", 0.7},
      {"South", "A", 0.1},
      {"North", "A", 1e16},
      {"North", "A", 1e-10},
  });
  const PivotResult r = evaluate_with(cache, build_sum_amount_table({0}, {}), Aggregation::Sum);
  ASSERT_TRUE(r.grand_total.is_number());
  EXPECT_EQ(r.grand_total.as_number(), 0.2);
}

TEST(PivotSumSnap, NestedRowGrandTotalAddsAllLeafCells) {
  // Regions R1..R3 sort in source order. Not via the region subtotal cells, and not the records in either order.
  const PivotCache cache = build_cache({
      {"R1", "A", 1e16},
      {"R2", "B", 1.9999999999999996},
      {"R2", "A", -0.1},
      {"R1", "B", -1e16},
      {"R2", "A", 1e-17},
      {"R1", "B", 0.1},
      {"R2", "A", 0.3},
      {"R1", "A", -2.000000000000002},
      {"R3", "B", -0.3},
      {"R3", "A", -2.000000000000002},
  });
  const PivotResult r = evaluate_with(cache, build_sum_amount_table({0, 1}, {}), Aggregation::Sum);
  ASSERT_TRUE(r.grand_total.is_number());
  EXPECT_EQ(r.grand_total.as_number(), -0.10000000000000248);
}

TEST(PivotSumSnap, LeafCellAddsRecordsInSourceOrder) {
  // Source order gives -0.0999999999999999; the reverse order gives -0.09999999999999994.
  const PivotCache cache = build_cache({
      {"North", "A", 1e-16},
      {"North", "A", 0.2},
      {"North", "A", -0.4},
      {"North", "A", 0.1},
  });
  const PivotResult r = evaluate_with(cache, build_sum_amount_table({0}, {}), Aggregation::Sum);
  const std::size_t north = row_index(r, "North");
  ASSERT_LT(north, r.values.size());
  ASSERT_TRUE(r.values[north][0][0].is_number());
  EXPECT_EQ(r.values[north][0][0].as_number(), -0.0999999999999999);
}

TEST(PivotSumSnap, TwoAxisGrandTotalIsPlainSumOfRowTotals) {
  // The grand total row snaps 1.3 + 1e16 - 1e16 to 0; the corner adds the row totals without a snap.
  const PivotCache cache = build_cache({
      {"R1", "A", 1.3},
      {"R2", "A", 1e16},
      {"R3", "A", -1e16},
  });
  const PivotResult r = evaluate_with(cache, build_sum_amount_table({0}, {1}), Aggregation::Sum);
  ASSERT_EQ(r.col_leaf_totals.size(), 1U);
  ASSERT_TRUE(r.col_leaf_totals[0][0].is_number());
  EXPECT_EQ(r.col_leaf_totals[0][0].as_number(), 0.0);
  ASSERT_TRUE(r.grand_total.is_number());
  EXPECT_EQ(r.grand_total.as_number(), 2.0);
}

TEST(PivotSumSnap, TwoAxisCornerKeepsResidueTheColumnTotalSnaps) {
  const PivotCache cache = build_cache({
      {"R1", "A", 0.7},
      {"R2", "A", -0.4},
      {"R3", "A", -0.3},
  });
  const PivotResult r = evaluate_with(cache, build_sum_amount_table({0}, {1}), Aggregation::Sum);
  ASSERT_EQ(r.col_leaf_totals.size(), 1U);
  ASSERT_TRUE(r.col_leaf_totals[0][0].is_number());
  EXPECT_EQ(r.col_leaf_totals[0][0].as_number(), 0.0);
  ASSERT_TRUE(r.grand_total.is_number());
  EXPECT_EQ(r.grand_total.as_number(), -5.551115123125783e-17);
}

TEST(PivotSumSnap, RowTotalAddsColumnLeavesLeftToRight) {
  constexpr double kBig = 9007199254740992.0;  // 2^53
  const PivotCache cache = build_cache({
      {"North", "A", kBig},
      {"North", "B", 4.0},
      {"North", "C", -kBig},
      {"North", "D", 100.0},
  });
  const PivotResult r = evaluate_with(cache, build_sum_amount_table({0}, {1}), Aggregation::Sum);
  const std::size_t north = row_index(r, "North");
  ASSERT_LT(north, r.row_leaf_totals.size());
  ASSERT_TRUE(r.row_leaf_totals[north][0].is_number());
  EXPECT_EQ(r.row_leaf_totals[north][0].as_number(), 100.0);
  ASSERT_TRUE(r.grand_total.is_number());
  EXPECT_EQ(r.grand_total.as_number(), 100.0);
}

TEST(PivotSumSnap, TwoAxisGrandTotalAddsColumnLeavesOuter) {
  const PivotCache cache = build_cache({
      {"North", "A", 0.5},
      {"South", "A", -0.4},
      {"North", "B", -0.1},
      {"East", "A", 1e-10},
      {"East", "B", 0.3},
  });
  const PivotResult r = evaluate_with(cache, build_sum_amount_table({0}, {1}), Aggregation::Sum);
  ASSERT_TRUE(r.grand_total.is_number());
  EXPECT_EQ(r.grand_total.as_number(), 0.30000000009999994);
}

TEST(PivotProductFlush, SubnormalProductIsZero) {
  const PivotCache cache = build_cache({
      {"North", "A", 3.0},
      {"South", "A", 1e-300},
      {"South", "B", 1e-10},
      {"East", "A", 2.0},
      {"East", "B", 4.0},
  });
  const PivotResult r = evaluate_with(cache, build_sum_amount_table({0}, {}), Aggregation::Product);
  const std::size_t south = row_index(r, "South");
  ASSERT_LT(south, r.values.size());
  ASSERT_TRUE(r.values[south][0][0].is_number());
  EXPECT_EQ(r.values[south][0][0].as_number(), 0.0);
  ASSERT_TRUE(r.grand_total.is_number());
  EXPECT_EQ(r.grand_total.as_number(), 0.0);
}

}  // namespace
}  // namespace formulon::pivot
