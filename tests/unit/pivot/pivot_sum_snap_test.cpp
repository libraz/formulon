// Pivot SUM / AVERAGE apply Excel's cancellation snap per step, PRODUCT
// flushes a subnormal step to 0, and the grand total adds records in leaf
// order (column leaves outer). Expected values were read from Mac Excel 365
// after refreshing a pivot over the same records.

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

TEST(PivotSumSnap, GrandTotalAddsInItemOrder) {
  // East, North, South: 1e-10 + 0.5 - 0.4 - 0.1 + 1e-300 keeps its residue.
  const PivotCache cache = build_cache(kRowOnlyRecords);
  const PivotResult r = evaluate_with(cache, build_sum_amount_table({0}, {}), Aggregation::Sum);
  ASSERT_TRUE(r.grand_total.is_number());
  EXPECT_EQ(r.grand_total.as_number(), 9.999998051846148e-11);
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
