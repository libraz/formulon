// Grand totals and subtotal aggregation metadata for `formulon::pivot::evaluate`.

#include <cstddef>
#include <string>
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

using test::build_basic_cache;
using test::build_sum_amount_table;
using test::row_index;

TEST(PivotEvaluator, GrandTotalWhenFlagged) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/true, /*cols=*/false);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  ASSERT_TRUE(r.grand_total.is_number());
  EXPECT_DOUBLE_EQ(r.grand_total.as_number(), 675.0);  // 100 + 50 + 200 + 300 + 25
  ASSERT_EQ(r.grand_totals.size(), 1U);
  ASSERT_TRUE(r.grand_totals[0].is_number());
  EXPECT_DOUBLE_EQ(r.grand_totals[0].as_number(), 675.0);
}

TEST(PivotEvaluator, GrandTotalsCarryOneValuePerDataField) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  PivotDataField count_amount;
  count_amount.name = "Count of Amount";
  count_amount.field_index = 2;
  count_amount.aggregation = Aggregation::Count;
  table.mutable_data_fields().push_back(std::move(count_amount));
  table.set_grand_totals(/*rows=*/true, /*cols=*/true);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  ASSERT_EQ(r.grand_totals.size(), 2U);
  ASSERT_TRUE(r.grand_totals[0].is_number());
  EXPECT_DOUBLE_EQ(r.grand_totals[0].as_number(), 675.0);
  ASSERT_TRUE(r.grand_totals[1].is_number());
  EXPECT_DOUBLE_EQ(r.grand_totals[1].as_number(), 5.0);
  ASSERT_TRUE(r.grand_total.is_number());
  EXPECT_DOUBLE_EQ(r.grand_total.as_number(), r.grand_totals[0].as_number());
}

TEST(PivotEvaluator, GrandTotalBlankWhenDisabled) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  EXPECT_TRUE(r_or.value().grand_total.is_blank());
}

TEST(PivotEvaluator, RowSubtotalEmittedWhenRequested) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0, 1}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  // Region declares subtotal_top so Region-level subtotals appear.
  table.mutable_fields()[0].subtotal_top = true;

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Two regions -> two subtotal rows, each with one data field slot.
  ASSERT_EQ(r.subtotals.size(), 2U);
  ASSERT_EQ(r.row_subtotals.size(), 2U);
  ASSERT_EQ(r.subtotals[0].size(), 1U);
  EXPECT_EQ(r.row_subtotals[0].depth, 0U);
  ASSERT_EQ(r.row_subtotals[0].labels.size(), 1U);
  // Subtotals appear in row-hierarchy DFS order (post-order at each
  // non-leaf), matching how Excel walks the tree to position them.
  std::vector<double> totals{r.subtotals[0][0].as_number(), r.subtotals[1][0].as_number()};
  std::sort(totals.begin(), totals.end());
  EXPECT_DOUBLE_EQ(totals[0], 175.0);  // North subtotal: 100 + 50 + 25
  EXPECT_DOUBLE_EQ(totals[1], 500.0);  // South subtotal: 200 + 300
}

TEST(PivotEvaluator, ColSubtotalEmittedWhenRequested) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{}, /*col=*/{0, 1});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  table.mutable_fields()[0].subtotal_top = true;

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  ASSERT_EQ(r.col_subtotals.size(), 2U);
  ASSERT_EQ(r.col_subtotals[0].labels.size(), 1U);
  EXPECT_EQ(r.col_subtotals[0].depth, 0U);
  ASSERT_EQ(r.col_subtotals[0].values.size(), 1U);
  ASSERT_EQ(r.col_subtotals[0].values[0].size(), 1U);

  std::vector<double> totals{r.col_subtotals[0].values[0][0].as_number(), r.col_subtotals[1].values[0][0].as_number()};
  std::sort(totals.begin(), totals.end());
  EXPECT_DOUBLE_EQ(totals[0], 175.0);
  EXPECT_DOUBLE_EQ(totals[1], 500.0);
}

TEST(PivotEvaluator, OuterSubtotalEmittedByDefaultWithoutSubtotalTop) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0, 1}, /*col=*/{});
  // Do not touch subtotal_top; default_subtotal defaults to true.
  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  // Region-level (depth 0) subtotals for North and South are present.
  ASSERT_EQ(r.row_subtotals.size(), 2U);
  EXPECT_EQ(r.row_subtotals[0].depth, 0U);
}

TEST(PivotEvaluator, DefaultSubtotalOffSuppressesOuterSubtotal) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0, 1}, /*col=*/{});
  table.mutable_fields()[0].default_subtotal = false;  // Region: no subtotal.
  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  EXPECT_TRUE(r_or.value().row_subtotals.empty());
}

TEST(PivotEvaluator, CustomSubtotalFunctionUsesSelectedAggregation) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0, 1}, /*col=*/{});
  // Region subtotal computed with Average instead of the data field's Sum.
  table.mutable_fields()[0].subtotal_fns = {SubtotalFn::Average};
  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.row_subtotals.size(), 2U);
  // North amounts: 100, 50, 25 -> average 58.333..., not the Sum 175.
  std::size_t north = r.row_subtotals[0].labels[0] == "North" ? 0 : 1;
  ASSERT_LT(north, r.row_subtotals.size());
  ASSERT_FALSE(r.row_subtotals[north].values.empty());
  ASSERT_TRUE(r.row_subtotals[north].values[0].is_number());
  EXPECT_NEAR(r.row_subtotals[north].values[0].as_number(), 175.0 / 3.0, 1e-9);
}

TEST(PivotEvaluator, ColumnCustomSubtotalFunctionUsesSelectedAggregation) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{}, /*col=*/{0, 1});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  table.mutable_fields()[0].subtotal_fns = {SubtotalFn::Average};

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.col_subtotals.size(), 2U);

  const std::size_t north = r.col_subtotals[0].labels[0] == "North" ? 0 : 1;
  ASSERT_LT(north, r.col_subtotals.size());
  ASSERT_TRUE(r.col_subtotals[north].aggregation.has_value());
  EXPECT_EQ(*r.col_subtotals[north].aggregation, Aggregation::Average);
  ASSERT_EQ(r.col_subtotals[north].values.size(), 1U);
  ASSERT_EQ(r.col_subtotals[north].values[0].size(), 1U);
  ASSERT_TRUE(r.col_subtotals[north].values[0][0].is_number());
  EXPECT_NEAR(r.col_subtotals[north].values[0][0].as_number(), 175.0 / 3.0, 1e-9);
}

TEST(PivotEvaluator, NonAdditiveRowTotalReaggregates) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{1});
  table.mutable_data_fields()[0].aggregation = Aggregation::Average;
  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  const std::size_t north = row_index(r, "North");
  ASSERT_LT(north, r.row_leaf_totals.size());
  ASSERT_FALSE(r.row_leaf_totals[north].empty());
  ASSERT_TRUE(r.row_leaf_totals[north][0].is_number());
  // North across all products: mean(100, 50, 25) = 58.333..., NOT the mean
  // of the per-cell averages (62.5, 50) = 56.25.
  EXPECT_NEAR(r.row_leaf_totals[north][0].as_number(), 175.0 / 3.0, 1e-9);
}

}  // namespace
}  // namespace formulon::pivot
