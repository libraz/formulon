//
// Unit tests for the show-values-as transforms `formulon::pivot::evaluate`
// applies: percent of row / column / total, running total, index,
// difference from and percent of parent, including their subtotals and
// grand totals.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_evaluator.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot_evaluator_fixtures.h"
#include "utils/error.h"
#include "value.h"

namespace formulon::pivot {
namespace {

using test::build_basic_cache;
using test::build_sum_amount_table;
using test::owned_text;
using test::row_index;

// ---------------------------------------------------------------------------
// 8e. Show values as (% of row / col / total / running total / index)
// ---------------------------------------------------------------------------

TEST(PivotEvaluator, ShowAsPercentOfRow) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{1});  // Region x Product.
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  table.mutable_data_fields()[0].show_as = ShowValuesAs::PercentOfRow;

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // North row totals: Gadget=50, Widget=125, sum=175.
  // South row totals: Gadget=300, Widget=200, sum=500.
  const std::size_t north = row_index(r, "North");
  const std::size_t south = row_index(r, "South");
  const std::size_t gadget = (r.cols[0].label == "Gadget") ? 0U : 1U;
  const std::size_t widget = 1U - gadget;
  EXPECT_DOUBLE_EQ(r.values[north][gadget][0].as_number(), 50.0 / 175.0);
  EXPECT_DOUBLE_EQ(r.values[north][widget][0].as_number(), 125.0 / 175.0);
  EXPECT_DOUBLE_EQ(r.values[south][gadget][0].as_number(), 300.0 / 500.0);
  EXPECT_DOUBLE_EQ(r.values[south][widget][0].as_number(), 200.0 / 500.0);
}

TEST(PivotEvaluator, ShowAsPercentOfRowReaggregatesNonAdditiveDenominator) {
  // AVERAGE is non-additive: the row denominator must be
  // AVERAGE(every North record) = (100+25+50)/3, not the sum of each
  // product's own already-computed average (62.5 + 50 = 112.5), which
  // is what summing per-cell aggregates would produce.
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{1});
  table.mutable_data_fields()[0].aggregation = Aggregation::Average;
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  table.mutable_data_fields()[0].show_as = ShowValuesAs::PercentOfRow;

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  const std::size_t north = row_index(r, "North");
  const std::size_t gadget = (r.cols[0].label == "Gadget") ? 0U : 1U;
  const std::size_t widget = 1U - gadget;
  const double north_avg = (100.0 + 25.0 + 50.0) / 3.0;
  EXPECT_DOUBLE_EQ(r.values[north][widget][0].as_number(), 62.5 / north_avg);
  EXPECT_DOUBLE_EQ(r.values[north][gadget][0].as_number(), 50.0 / north_avg);
}

TEST(PivotEvaluator, ShowAsPercentOfTotalReaggregatesNonAdditiveDenominatorWithTotalsHidden) {
  // The denominator is AVERAGE over every record = 675 / 5, not the sum of the
  // leaf-cell averages, and the hidden grand totals stay absent from the result.
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{1});
  table.mutable_data_fields()[0].aggregation = Aggregation::Average;
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  table.mutable_data_fields()[0].show_as = ShowValuesAs::PercentOfTotal;

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  const std::size_t north = row_index(r, "North");
  const std::size_t gadget = (r.cols[0].label == "Gadget") ? 0U : 1U;
  const std::size_t widget = 1U - gadget;
  EXPECT_DOUBLE_EQ(r.values[north][widget][0].as_number(), 62.5 / 135.0);
  EXPECT_DOUBLE_EQ(r.values[north][gadget][0].as_number(), 50.0 / 135.0);
  EXPECT_TRUE(r.grand_totals.empty());
  EXPECT_TRUE(r.grand_total.is_blank());
}

TEST(PivotEvaluator, ShowAsPercentOfCol) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{1});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  table.mutable_data_fields()[0].show_as = ShowValuesAs::PercentOfCol;

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Column totals: Gadget = 50 + 300 = 350, Widget = 125 + 200 = 325.
  const std::size_t north = row_index(r, "North");
  const std::size_t south = row_index(r, "South");
  const std::size_t gadget = (r.cols[0].label == "Gadget") ? 0U : 1U;
  const std::size_t widget = 1U - gadget;
  EXPECT_DOUBLE_EQ(r.values[north][gadget][0].as_number(), 50.0 / 350.0);
  EXPECT_DOUBLE_EQ(r.values[south][gadget][0].as_number(), 300.0 / 350.0);
  EXPECT_DOUBLE_EQ(r.values[north][widget][0].as_number(), 125.0 / 325.0);
  EXPECT_DOUBLE_EQ(r.values[south][widget][0].as_number(), 200.0 / 325.0);
}

TEST(PivotEvaluator, ShowAsPercentOfTotal) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{1});
  table.set_grand_totals(/*rows=*/true, /*cols=*/true);
  table.mutable_data_fields()[0].show_as = ShowValuesAs::PercentOfTotal;

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Grand total = 100+50+200+300+25 = 675.
  const std::size_t north = row_index(r, "North");
  const std::size_t gadget = (r.cols[0].label == "Gadget") ? 0U : 1U;
  EXPECT_DOUBLE_EQ(r.values[north][gadget][0].as_number(), 50.0 / 675.0);
}

TEST(PivotEvaluator, ShowAsRunningTotalInRow) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{1});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  table.mutable_data_fields()[0].show_as = ShowValuesAs::RunningTotalInRow;

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  const std::size_t north = row_index(r, "North");
  const std::size_t gadget = (r.cols[0].label == "Gadget") ? 0U : 1U;
  const std::size_t widget = 1U - gadget;
  // Cumulative across cols in display order (Gadget < Widget).
  EXPECT_DOUBLE_EQ(r.values[north][0][0].as_number(), gadget == 0 ? 50.0 : 125.0);
  EXPECT_DOUBLE_EQ(r.values[north][1][0].as_number(), 50.0 + 125.0);
  (void)widget;
}

TEST(PivotEvaluator, ShowAsIndex) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{1});
  table.set_grand_totals(/*rows=*/true, /*cols=*/true);
  table.mutable_data_fields()[0].show_as = ShowValuesAs::Index;

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Index(N, Gadget) = (cell * total) / (row_sum * col_sum)
  //                  = (50 * 675) / (175 * 350)
  const std::size_t north = row_index(r, "North");
  const std::size_t gadget = (r.cols[0].label == "Gadget") ? 0U : 1U;
  EXPECT_DOUBLE_EQ(r.values[north][gadget][0].as_number(), (50.0 * 675.0) / (175.0 * 350.0));
}

TEST(PivotEvaluator, ShowAsPercentRowZeroSumYieldsDiv0) {
  // Construct a degenerate row whose data field sums to 0 so the
  // % of row transform must surface Div0 rather than NaN.
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"R", {}});
  cache.mutable_fields().push_back(PivotCacheField{"C", {}});
  cache.mutable_fields().push_back(PivotCacheField{"V", {}});
  auto add = [&](const char* row_label, const char* col_label, double v) {
    PivotCacheRecord rec;
    rec.cells.push_back(owned_text(cache, row_label));
    rec.cells.push_back(owned_text(cache, col_label));
    rec.cells.push_back(Value::number(v));
    cache.mutable_records().push_back(std::move(rec));
  };
  add("X", "P", 5.0);
  add("X", "Q", -5.0);  // Row X sums to 0.

  PivotTable table;
  table.set_pivot_cache_id(1);
  PivotField rf;
  rf.source_name = "R";
  rf.axis = PivotAxis::Row;
  PivotField cf;
  cf.source_name = "C";
  cf.axis = PivotAxis::Col;
  PivotField vf;
  vf.source_name = "V";
  vf.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(rf));
  table.mutable_fields().push_back(std::move(cf));
  table.mutable_fields().push_back(std::move(vf));
  table.mutable_row_field_order() = {0};
  table.mutable_col_field_order() = {1};
  PivotDataField sum;
  sum.name = "Sum of V";
  sum.field_index = 2;
  sum.aggregation = Aggregation::Sum;
  sum.show_as = ShowValuesAs::PercentOfRow;
  table.mutable_data_fields().push_back(std::move(sum));
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.rows.size(), 1U);
  ASSERT_EQ(r.cols.size(), 2U);
  EXPECT_TRUE(r.values[0][0][0].is_error());
  EXPECT_EQ(r.values[0][0][0].as_error(), ErrorCode::Div0);
  EXPECT_TRUE(r.values[0][1][0].is_error());
}

// ---------------------------------------------------------------------------
// 8f. Show-values-as transforms propagate to subtotals + grand totals
// (Percent* ratio modes only; Running / Index keep raw aggregates)
// ---------------------------------------------------------------------------

// Build a 4-field cache (Region, Product, Channel, Amount) so a single
// PivotTable can present both a row hierarchy (Region/Product) and a
// column hierarchy (Channel/Amount-not-needed -> we use Channel only).
// The data is small and hand-summable.
PivotCache build_show_as_cache() {
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Region", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Product", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Channel", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});
  auto add = [&](const char* region, const char* product, const char* channel, double amount) {
    PivotCacheRecord rec;
    rec.cells.push_back(owned_text(cache, region));
    rec.cells.push_back(owned_text(cache, product));
    rec.cells.push_back(owned_text(cache, channel));
    rec.cells.push_back(Value::number(amount));
    cache.mutable_records().push_back(std::move(rec));
  };
  // Region North:
  //   Widget:  Online=10, Store=20  -> row=30, by-channel(Online)=10, (Store)=20
  //   Gadget:  Online=30, Store=40  -> row=70, by-channel(Online)=30, (Store)=40
  // Region South:
  //   Widget:  Online=50, Store=60  -> row=110
  //   Gadget:  Online=70, Store=80  -> row=150
  // Grand total = 30 + 70 + 110 + 150 = 360.
  // Per-col-leaf totals: Online = 10+30+50+70 = 160, Store = 20+40+60+80 = 200.
  add("North", "Widget", "Online", 10.0);
  add("North", "Widget", "Store", 20.0);
  add("North", "Gadget", "Online", 30.0);
  add("North", "Gadget", "Store", 40.0);
  add("South", "Widget", "Online", 50.0);
  add("South", "Widget", "Store", 60.0);
  add("South", "Gadget", "Online", 70.0);
  add("South", "Gadget", "Store", 80.0);
  return cache;
}

// Builds a pivot over `build_show_as_cache()` with row hierarchy
// (Region, Product) and column axis (Channel). The Region row field
// emits subtotals so `row_subtotals` is populated and carries
// `col_values`.
PivotTable build_show_as_table(ShowValuesAs mode, bool grand_totals) {
  PivotTable table;
  table.set_pivot_cache_id(1);
  PivotField region_f;
  region_f.source_name = "Region";
  region_f.axis = PivotAxis::Row;
  region_f.subtotal_top = true;
  PivotField product_f;
  product_f.source_name = "Product";
  product_f.axis = PivotAxis::Row;
  PivotField channel_f;
  channel_f.source_name = "Channel";
  channel_f.axis = PivotAxis::Col;
  PivotField amount_f;
  amount_f.source_name = "Amount";
  amount_f.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(region_f));
  table.mutable_fields().push_back(std::move(product_f));
  table.mutable_fields().push_back(std::move(channel_f));
  table.mutable_fields().push_back(std::move(amount_f));
  table.mutable_row_field_order() = {0, 1};
  table.mutable_col_field_order() = {2};

  PivotDataField sum_amount;
  sum_amount.name = "Sum of Amount";
  sum_amount.field_index = 3;
  sum_amount.aggregation = Aggregation::Sum;
  sum_amount.show_as = mode;
  table.mutable_data_fields().push_back(std::move(sum_amount));
  table.set_grand_totals(grand_totals, grand_totals);
  return table;
}

TEST(PivotEvaluator, ShowAsPercentOfRowTransformsRowSubtotal) {
  PivotCache cache = build_show_as_cache();
  PivotTable table = build_show_as_table(ShowValuesAs::PercentOfRow, /*grand_totals=*/false);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  ASSERT_EQ(r.row_subtotals.size(), 2U);
  // Each leaf row should sum to 1.0 across cols.
  ASSERT_EQ(r.values.size(), 4U);
  for (const auto& row_slot : r.values) {
    double row_sum = 0.0;
    for (const auto& col_slot : row_slot) {
      ASSERT_FALSE(col_slot.empty());
      ASSERT_TRUE(col_slot[0].is_number());
      row_sum += col_slot[0].as_number();
    }
    EXPECT_NEAR(row_sum, 1.0, 1e-9);
  }
  // Each row_subtotal's col_values should sum to 1.0; values[0] should be 1.0.
  for (const RowSubtotal& sub : r.row_subtotals) {
    double sub_row_sum = 0.0;
    for (const auto& col_slot : sub.col_values) {
      ASSERT_FALSE(col_slot.empty());
      ASSERT_TRUE(col_slot[0].is_number());
      sub_row_sum += col_slot[0].as_number();
    }
    EXPECT_NEAR(sub_row_sum, 1.0, 1e-9);
    ASSERT_FALSE(sub.values.empty());
    ASSERT_TRUE(sub.values[0].is_number());
    EXPECT_NEAR(sub.values[0].as_number(), 1.0, 1e-9);
  }
}

TEST(PivotEvaluator, ShowAsPercentOfRowGrandTotalIsOne) {
  PivotCache cache = build_show_as_cache();
  PivotTable table = build_show_as_table(ShowValuesAs::PercentOfRow, /*grand_totals=*/true);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  ASSERT_EQ(r.grand_totals.size(), 1U);
  ASSERT_TRUE(r.grand_totals[0].is_number());
  EXPECT_NEAR(r.grand_totals[0].as_number(), 1.0, 1e-9);
  // Legacy mirror also re-synced.
  ASSERT_TRUE(r.grand_total.is_number());
  EXPECT_NEAR(r.grand_total.as_number(), 1.0, 1e-9);
}

TEST(PivotEvaluator, ShowAsPercentOfRowTransformsMarginTotals) {
  // The right-hand "Grand Total" column (row_leaf_totals) normalizes to
  // 1.0 against itself; the bottom "Grand Total" row (col_leaf_totals)
  // is itself just another row, so it divides by the overall grand
  // total instead. Region totals: North=175, South=500; Product totals:
  // Gadget=350, Widget=325; grand=675.
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{1});
  table.set_grand_totals(/*rows=*/true, /*cols=*/true);
  table.mutable_data_fields()[0].show_as = ShowValuesAs::PercentOfRow;

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  const std::size_t north = row_index(r, "North");
  const std::size_t south = row_index(r, "South");
  ASSERT_EQ(r.row_leaf_totals.size(), 2U);
  EXPECT_DOUBLE_EQ(r.row_leaf_totals[north][0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(r.row_leaf_totals[south][0].as_number(), 1.0);

  const std::size_t gadget = (r.cols[0].label == "Gadget") ? 0U : 1U;
  const std::size_t widget = 1U - gadget;
  ASSERT_EQ(r.col_leaf_totals.size(), 2U);
  EXPECT_DOUBLE_EQ(r.col_leaf_totals[gadget][0].as_number(), 350.0 / 675.0);
  EXPECT_DOUBLE_EQ(r.col_leaf_totals[widget][0].as_number(), 325.0 / 675.0);
}

TEST(PivotEvaluator, ShowAsPercentOfColTransformsMarginTotals) {
  // Mirror of ShowAsPercentOfRowTransformsMarginTotals: col_leaf_totals
  // normalizes to 1.0 against itself, row_leaf_totals divides by the
  // overall grand total.
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{1});
  table.set_grand_totals(/*rows=*/true, /*cols=*/true);
  table.mutable_data_fields()[0].show_as = ShowValuesAs::PercentOfCol;

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  ASSERT_EQ(r.col_leaf_totals.size(), 2U);
  EXPECT_DOUBLE_EQ(r.col_leaf_totals[0][0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(r.col_leaf_totals[1][0].as_number(), 1.0);

  const std::size_t north = row_index(r, "North");
  const std::size_t south = row_index(r, "South");
  ASSERT_EQ(r.row_leaf_totals.size(), 2U);
  EXPECT_DOUBLE_EQ(r.row_leaf_totals[north][0].as_number(), 175.0 / 675.0);
  EXPECT_DOUBLE_EQ(r.row_leaf_totals[south][0].as_number(), 500.0 / 675.0);
}

TEST(PivotEvaluator, ShowAsPercentOfColTransformsColSubtotal) {
  // Use a col hierarchy (Region/Product) and a row axis (Channel) so
  // col_subtotals is populated.
  PivotCache cache = build_show_as_cache();
  PivotTable table;
  table.set_pivot_cache_id(1);
  PivotField region_f;
  region_f.source_name = "Region";
  region_f.axis = PivotAxis::Col;
  region_f.subtotal_top = true;
  PivotField product_f;
  product_f.source_name = "Product";
  product_f.axis = PivotAxis::Col;
  PivotField channel_f;
  channel_f.source_name = "Channel";
  channel_f.axis = PivotAxis::Row;
  PivotField amount_f;
  amount_f.source_name = "Amount";
  amount_f.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(region_f));
  table.mutable_fields().push_back(std::move(product_f));
  table.mutable_fields().push_back(std::move(channel_f));
  table.mutable_fields().push_back(std::move(amount_f));
  table.mutable_row_field_order() = {2};
  table.mutable_col_field_order() = {0, 1};
  PivotDataField sum_amount;
  sum_amount.name = "Sum of Amount";
  sum_amount.field_index = 3;
  sum_amount.aggregation = Aggregation::Sum;
  sum_amount.show_as = ShowValuesAs::PercentOfCol;
  table.mutable_data_fields().push_back(std::move(sum_amount));
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  ASSERT_EQ(r.col_subtotals.size(), 2U);
  // Each col_subtotal's column should sum to 1.0 across the row leaves.
  for (const ColSubtotal& csub : r.col_subtotals) {
    double col_sum = 0.0;
    for (const auto& row_slot : csub.values) {
      ASSERT_FALSE(row_slot.empty());
      ASSERT_TRUE(row_slot[0].is_number());
      col_sum += row_slot[0].as_number();
    }
    EXPECT_NEAR(col_sum, 1.0, 1e-9);
  }
  // And every leaf column should also sum to 1.0.
  ASSERT_FALSE(r.values.empty());
  const std::size_t n_cols = r.values[0].size();
  for (std::size_t c = 0; c < n_cols; ++c) {
    double col_sum = 0.0;
    for (const auto& row_slot : r.values) {
      ASSERT_FALSE(row_slot[c].empty());
      ASSERT_TRUE(row_slot[c][0].is_number());
      col_sum += row_slot[c][0].as_number();
    }
    EXPECT_NEAR(col_sum, 1.0, 1e-9);
  }
}

TEST(PivotEvaluator, ShowAsPercentOfColReaggregatesNonAdditiveColSubtotalDenominator) {
  // AVERAGE is non-additive: the North col-subtotal's own denominator
  // must be AVERAGE(10,20,30,40) = 25 (every North record, re-
  // aggregated), not the sum of its per-row-leaf averages
  // (AVERAGE(10,30)=20 for Online + AVERAGE(20,40)=30 for Store = 50),
  // which is what summing `ColSubtotal::values` would give.
  PivotCache cache = build_show_as_cache();
  PivotTable table;
  table.set_pivot_cache_id(1);
  PivotField region_f;
  region_f.source_name = "Region";
  region_f.axis = PivotAxis::Col;
  PivotField product_f;
  product_f.source_name = "Product";
  product_f.axis = PivotAxis::Col;
  PivotField channel_f;
  channel_f.source_name = "Channel";
  channel_f.axis = PivotAxis::Row;
  PivotField amount_f;
  amount_f.source_name = "Amount";
  amount_f.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(region_f));
  table.mutable_fields().push_back(std::move(product_f));
  table.mutable_fields().push_back(std::move(channel_f));
  table.mutable_fields().push_back(std::move(amount_f));
  table.mutable_row_field_order() = {2};
  table.mutable_col_field_order() = {0, 1};
  PivotDataField avg_amount;
  avg_amount.name = "Average of Amount";
  avg_amount.field_index = 3;
  avg_amount.aggregation = Aggregation::Average;
  avg_amount.show_as = ShowValuesAs::PercentOfCol;
  table.mutable_data_fields().push_back(std::move(avg_amount));
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  ASSERT_EQ(r.col_subtotals.size(), 2U);
  const ColSubtotal* north_subtotal = nullptr;
  for (const ColSubtotal& csub : r.col_subtotals) {
    if (!csub.labels.empty() && csub.labels[0] == "North") {
      north_subtotal = &csub;
    }
  }
  ASSERT_NE(north_subtotal, nullptr);
  const std::size_t online = row_index(r, "Online");
  const std::size_t store = row_index(r, "Store");
  ASSERT_TRUE(north_subtotal->values[online][0].is_number());
  ASSERT_TRUE(north_subtotal->values[store][0].is_number());
  EXPECT_DOUBLE_EQ(north_subtotal->values[online][0].as_number(), 20.0 / 25.0);
  EXPECT_DOUBLE_EQ(north_subtotal->values[store][0].as_number(), 30.0 / 25.0);
}

TEST(PivotEvaluator, ShowAsPercentOfTotalAppliesToSubtotalsAndGrandTotal) {
  // Row hierarchy (Region/Product) with subtotals, single-level col
  // (Channel). Grand totals on. Verifies PercentOfTotal propagates to
  // every slot: leaves, row_subtotals.values, row_subtotals.col_values,
  // and grand_totals. (Col-subtotal-specific coverage lives in
  // `ShowAsPercentOfColTransformsColSubtotal`.)
  PivotCache cache = build_show_as_cache();
  PivotTable table = build_show_as_table(ShowValuesAs::PercentOfTotal, /*grand_totals=*/true);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Grand total = 360 originally, after PercentOfTotal => 1.0.
  ASSERT_EQ(r.grand_totals.size(), 1U);
  ASSERT_TRUE(r.grand_totals[0].is_number());
  EXPECT_NEAR(r.grand_totals[0].as_number(), 1.0, 1e-9);

  // Row subtotals: North = 100/360, South = 260/360.
  ASSERT_EQ(r.row_subtotals.size(), 2U);
  std::vector<double> row_subtotal_vals;
  for (const RowSubtotal& sub : r.row_subtotals) {
    ASSERT_FALSE(sub.values.empty());
    ASSERT_TRUE(sub.values[0].is_number());
    row_subtotal_vals.push_back(sub.values[0].as_number());
  }
  std::sort(row_subtotal_vals.begin(), row_subtotal_vals.end());
  EXPECT_NEAR(row_subtotal_vals[0], 100.0 / 360.0, 1e-9);
  EXPECT_NEAR(row_subtotal_vals[1], 260.0 / 360.0, 1e-9);

  // Every leaf cell must equal its raw / 360. Cells sum to 1.0.
  double leaf_sum = 0.0;
  for (const auto& row_slot : r.values) {
    for (const auto& col_slot : row_slot) {
      ASSERT_FALSE(col_slot.empty());
      ASSERT_TRUE(col_slot[0].is_number());
      leaf_sum += col_slot[0].as_number();
    }
  }
  EXPECT_NEAR(leaf_sum, 1.0, 1e-9);

  // Every row_subtotal.col_values cell is raw/360, so all
  // row_subtotal col_values together sum to (sum of row subtotal
  // raws)/360 = 360/360 = 1.0.
  double row_sub_col_values_sum = 0.0;
  for (const RowSubtotal& sub : r.row_subtotals) {
    for (const auto& col_slot : sub.col_values) {
      ASSERT_FALSE(col_slot.empty());
      ASSERT_TRUE(col_slot[0].is_number());
      row_sub_col_values_sum += col_slot[0].as_number();
    }
  }
  EXPECT_NEAR(row_sub_col_values_sum, 1.0, 1e-9);

  // Legacy mirrors are kept in sync.
  ASSERT_TRUE(r.grand_total.is_number());
  EXPECT_NEAR(r.grand_total.as_number(), 1.0, 1e-9);
  ASSERT_EQ(r.subtotals.size(), 2U);
  for (std::size_t i = 0; i < 2; ++i) {
    ASSERT_FALSE(r.subtotals[i].empty());
    ASSERT_TRUE(r.subtotals[i][0].is_number());
    EXPECT_NEAR(r.subtotals[i][0].as_number(), r.row_subtotals[i].values[0].as_number(), 1e-12);
  }
}

TEST(PivotEvaluator, ShowAsRunningTotalInRowAccumulatesSubtotalsAndBlanksTotalsAlongBaseField) {
  PivotCache cache = build_show_as_cache();
  PivotTable table = build_show_as_table(ShowValuesAs::RunningTotalInRow, /*grand_totals=*/true);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // North / Gadget: Online=30, Store=40 -> 30, 70.
  ASSERT_EQ(r.values.size(), 4U);
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 30.0);
  EXPECT_DOUBLE_EQ(r.values[0][1][0].as_number(), 70.0);
  // The Region subtotal rows accumulate across the channels too:
  // North Online=40, Store=60 -> 40, 100.
  ASSERT_EQ(r.row_subtotals.size(), 2U);
  const RowSubtotal& north = r.row_subtotals[0];
  EXPECT_DOUBLE_EQ(north.col_values[0][0].as_number(), 40.0);
  EXPECT_DOUBLE_EQ(north.col_values[1][0].as_number(), 100.0);
  // The totals along the base field render blank.
  EXPECT_TRUE(north.values[0].is_blank());
  EXPECT_TRUE(r.row_leaf_totals[0][0].is_blank());
  ASSERT_EQ(r.grand_totals.size(), 1U);
  EXPECT_TRUE(r.grand_totals[0].is_blank());
  // The Grand Total row accumulates across the channels: 160, 360.
  EXPECT_DOUBLE_EQ(r.col_leaf_totals[0][0].as_number(), 160.0);
  EXPECT_DOUBLE_EQ(r.col_leaf_totals[1][0].as_number(), 360.0);
}

// ---------------------------------------------------------------------------
// 8g. Show-values-as: Difference From / % Difference From / % Of Parent
// ---------------------------------------------------------------------------

// Builds a 2-column cache (Region, Amount) with three rows so the
// row-axis transforms have three positions to operate on.
//   Region  Amount
//   ------  ------
//   A       10
//   B       25
//   C       40
PivotCache build_diff_from_cache() {
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Region", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});
  auto add = [&](const char* region, double amount) {
    PivotCacheRecord rec;
    rec.cells.push_back(owned_text(cache, region));
    rec.cells.push_back(Value::number(amount));
    cache.mutable_records().push_back(std::move(rec));
  };
  add("A", 10.0);
  add("B", 25.0);
  add("C", 40.0);
  return cache;
}

// Builds a single-row-field pivot whose row field declares items
// [A, B, C] so the "specific item" lookup can resolve a base item
// index to a row label. The data field's show-as configuration is the
// caller's responsibility.
PivotTable build_diff_from_table() {
  PivotTable table;
  table.set_pivot_cache_id(1);
  PivotField region_f;
  region_f.source_name = "Region";
  region_f.axis = PivotAxis::Row;
  region_f.items.push_back(PivotItem{"A", true});
  region_f.items.push_back(PivotItem{"B", true});
  region_f.items.push_back(PivotItem{"C", true});
  PivotField amount_f;
  amount_f.source_name = "Amount";
  amount_f.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(region_f));
  table.mutable_fields().push_back(std::move(amount_f));
  table.mutable_row_field_order() = {0};

  PivotDataField sum_amount;
  sum_amount.name = "Sum of Amount";
  sum_amount.field_index = 1;
  sum_amount.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum_amount));
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  return table;
}

TEST(PivotEvaluator, ShowAsDifferenceFromPreviousRowAxis) {
  PivotCache cache = build_diff_from_cache();
  PivotTable table = build_diff_from_table();
  auto& df = table.mutable_data_fields()[0];
  df.show_as = ShowValuesAs::DifferenceFrom;
  df.show_as_base_field = 0U;  // Region
  df.show_as_base_item = kShowAsBasePrev;

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.values.size(), 3U);
  ASSERT_EQ(r.values[0][0].size(), 1U);
  // Row 0 (A) has no previous -> blank.
  EXPECT_TRUE(r.values[0][0][0].is_blank());
  // Row 1 (B): 25 - 10 = 15.
  ASSERT_TRUE(r.values[1][0][0].is_number());
  EXPECT_DOUBLE_EQ(r.values[1][0][0].as_number(), 15.0);
  // Row 2 (C): 40 - 25 = 15.
  ASSERT_TRUE(r.values[2][0][0].is_number());
  EXPECT_DOUBLE_EQ(r.values[2][0][0].as_number(), 15.0);
}

TEST(PivotEvaluator, ShowAsPercentDifferenceFromPrevious) {
  PivotCache cache = build_diff_from_cache();
  PivotTable table = build_diff_from_table();
  auto& df = table.mutable_data_fields()[0];
  df.show_as = ShowValuesAs::PercentDifferenceFrom;
  df.show_as_base_field = 0U;
  df.show_as_base_item = kShowAsBasePrev;

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.values.size(), 3U);
  EXPECT_TRUE(r.values[0][0][0].is_blank());
  ASSERT_TRUE(r.values[1][0][0].is_number());
  EXPECT_DOUBLE_EQ(r.values[1][0][0].as_number(), 25.0 / 10.0 - 1.0);  // 1.5
  ASSERT_TRUE(r.values[2][0][0].is_number());
  EXPECT_DOUBLE_EQ(r.values[2][0][0].as_number(), 40.0 / 25.0 - 1.0);  // 0.6
}

TEST(PivotEvaluator, ShowAsDifferenceFromSpecificItem) {
  PivotCache cache = build_diff_from_cache();
  PivotTable table = build_diff_from_table();
  auto& df = table.mutable_data_fields()[0];
  df.show_as = ShowValuesAs::DifferenceFrom;
  df.show_as_base_field = 0U;
  df.show_as_base_item = 0U;  // -> field.items[0].name = "A"

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.values.size(), 3U);
  EXPECT_TRUE(r.values[0][0][0].is_blank());  // The base item itself is blank.
  ASSERT_TRUE(r.values[1][0][0].is_number());
  EXPECT_DOUBLE_EQ(r.values[1][0][0].as_number(), 15.0);  // B - A = 25 - 10.
  ASSERT_TRUE(r.values[2][0][0].is_number());
  EXPECT_DOUBLE_EQ(r.values[2][0][0].as_number(), 30.0);  // C - A = 40 - 10.
}

TEST(PivotEvaluator, ShowAsDifferenceFromUnnamedItemMatchesItsBoundLabel) {
  PivotCache cache = build_diff_from_cache();
  for (const char* label : {"A", "B", "C"}) {
    cache.mutable_fields()[0].shared_items.push_back(owned_text(cache, label));
  }
  PivotTable table = build_diff_from_table();
  std::vector<PivotItem>& items = table.mutable_fields()[0].items;
  for (std::uint32_t i = 0; i < items.size(); ++i) {
    items[i].name.clear();
    items[i].has_cache_index = true;
    items[i].cache_index = i;
  }
  auto& df = table.mutable_data_fields()[0];
  df.show_as = ShowValuesAs::DifferenceFrom;
  df.show_as_base_field = 0U;
  df.show_as_base_item = 1U;  // -> shared_items[1] = "B"

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.values.size(), 3U);
  ASSERT_TRUE(r.values[0][0][0].is_number());
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), -15.0);  // A - B = 10 - 25.
  EXPECT_TRUE(r.values[1][0][0].is_blank());
  ASSERT_TRUE(r.values[2][0][0].is_number());
  EXPECT_DOUBLE_EQ(r.values[2][0][0].as_number(), 15.0);  // C - B = 40 - 25.
}

// 2-level row hierarchy (Region -> Product) with subtotal_top on Region.
// Records:
//   N/A = 10, N/B = 20, S/A = 30, S/B = 60.
// Region subtotals: North = 30, South = 90.
PivotCache build_parent_row_cache() {
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Region", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Product", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});
  auto add = [&](const char* region, const char* product, double amount) {
    PivotCacheRecord rec;
    rec.cells.push_back(owned_text(cache, region));
    rec.cells.push_back(owned_text(cache, product));
    rec.cells.push_back(Value::number(amount));
    cache.mutable_records().push_back(std::move(rec));
  };
  add("North", "A", 10.0);
  add("North", "B", 20.0);
  add("South", "A", 30.0);
  add("South", "B", 60.0);
  return cache;
}

TEST(PivotEvaluator, ShowAsPercentOfParentRowSingleParent) {
  PivotCache cache = build_parent_row_cache();
  PivotTable table;
  table.set_pivot_cache_id(1);
  PivotField region_f;
  region_f.source_name = "Region";
  region_f.axis = PivotAxis::Row;
  region_f.subtotal_top = true;  // emit Region-level subtotals.
  PivotField product_f;
  product_f.source_name = "Product";
  product_f.axis = PivotAxis::Row;
  PivotField amount_f;
  amount_f.source_name = "Amount";
  amount_f.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(region_f));
  table.mutable_fields().push_back(std::move(product_f));
  table.mutable_fields().push_back(std::move(amount_f));
  table.mutable_row_field_order() = {0, 1};

  PivotDataField sum_amount;
  sum_amount.name = "Sum of Amount";
  sum_amount.field_index = 2;
  sum_amount.aggregation = Aggregation::Sum;
  sum_amount.show_as = ShowValuesAs::PercentOfParentRow;
  sum_amount.show_as_base_field = 0U;  // Region
  table.mutable_data_fields().push_back(std::move(sum_amount));
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  // 4 leaves in row-axis DFS order: N/A, N/B, S/A, S/B.
  ASSERT_EQ(r.values.size(), 4U);
  ASSERT_EQ(r.values[0][0].size(), 1U);
  ASSERT_TRUE(r.values[0][0][0].is_number());
  EXPECT_NEAR(r.values[0][0][0].as_number(), 10.0 / 30.0, 1e-9);
  ASSERT_TRUE(r.values[1][0][0].is_number());
  EXPECT_NEAR(r.values[1][0][0].as_number(), 20.0 / 30.0, 1e-9);
  ASSERT_TRUE(r.values[2][0][0].is_number());
  EXPECT_NEAR(r.values[2][0][0].as_number(), 30.0 / 90.0, 1e-9);
  ASSERT_TRUE(r.values[3][0][0].is_number());
  EXPECT_NEAR(r.values[3][0][0].as_number(), 60.0 / 90.0, 1e-9);
}

TEST(PivotEvaluator, ShowAsPercentOfParentColSingleParent) {
  // Column-axis equivalent of ShowAsPercentOfParentRowSingleParent.
  // Same data; rotate axes so Region/Product become column fields.
  PivotCache cache = build_parent_row_cache();
  PivotTable table;
  table.set_pivot_cache_id(1);
  PivotField region_f;
  region_f.source_name = "Region";
  region_f.axis = PivotAxis::Col;
  region_f.subtotal_top = true;
  PivotField product_f;
  product_f.source_name = "Product";
  product_f.axis = PivotAxis::Col;
  PivotField amount_f;
  amount_f.source_name = "Amount";
  amount_f.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(region_f));
  table.mutable_fields().push_back(std::move(product_f));
  table.mutable_fields().push_back(std::move(amount_f));
  table.mutable_col_field_order() = {0, 1};

  PivotDataField sum_amount;
  sum_amount.name = "Sum of Amount";
  sum_amount.field_index = 2;
  sum_amount.aggregation = Aggregation::Sum;
  sum_amount.show_as = ShowValuesAs::PercentOfParentCol;
  sum_amount.show_as_base_field = 0U;  // Region
  table.mutable_data_fields().push_back(std::move(sum_amount));
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  // 1 implicit row leaf, 4 col leaves: N/A, N/B, S/A, S/B.
  ASSERT_EQ(r.values.size(), 1U);
  ASSERT_EQ(r.values[0].size(), 4U);
  ASSERT_TRUE(r.values[0][0][0].is_number());
  EXPECT_NEAR(r.values[0][0][0].as_number(), 10.0 / 30.0, 1e-9);
  ASSERT_TRUE(r.values[0][1][0].is_number());
  EXPECT_NEAR(r.values[0][1][0].as_number(), 20.0 / 30.0, 1e-9);
  ASSERT_TRUE(r.values[0][2][0].is_number());
  EXPECT_NEAR(r.values[0][2][0].as_number(), 30.0 / 90.0, 1e-9);
  ASSERT_TRUE(r.values[0][3][0].is_number());
  EXPECT_NEAR(r.values[0][3][0].as_number(), 60.0 / 90.0, 1e-9);
}

// ---------------------------------------------------------------------------
// H-25: Index still produces non-zero results when grand totals are off.
// ---------------------------------------------------------------------------

TEST(PivotEvaluator, IndexWorksWithoutGrandTotals) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{1});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  table.mutable_data_fields()[0].show_as = ShowValuesAs::Index;
  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  // At least one leaf cell must be a non-zero index (the pre-fix behaviour
  // collapsed every cell to 0 because the total denominator was 0).
  bool saw_nonzero = false;
  for (const auto& row_slot : r.values) {
    for (const auto& cell_slot : row_slot) {
      if (!cell_slot.empty() && cell_slot[0].is_number() && cell_slot[0].as_number() != 0.0) {
        saw_nonzero = true;
      }
    }
  }
  EXPECT_TRUE(saw_nonzero);
}

// ---------------------------------------------------------------------------
// runTotal accumulation direction follows the data field's baseField, not the
// enum spelling. A base field on the column axis accumulates across the
// columns even when the mode is RunningTotalInCol.
// ---------------------------------------------------------------------------

TEST(PivotEvaluator, RunTotalDirectionFollowsBaseFieldAxis) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{1});
  table.mutable_data_fields()[0].show_as = ShowValuesAs::RunningTotalInCol;
  table.mutable_data_fields()[0].show_as_base_field = 1;  // Product = a column field.
  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  const std::size_t south = row_index(r, "South");
  ASSERT_LT(south, r.values.size());
  ASSERT_EQ(r.values[south].size(), 2U);
  const PivotResult plain = evaluate(build_sum_amount_table({0}, {1}), cache).value();
  // Across the Product columns: Gadget keeps its own value, Widget adds it.
  const double gadget = plain.values[south][0][0].as_number();
  const double widget = plain.values[south][1][0].as_number();
  EXPECT_NEAR(r.values[south][0][0].as_number(), gadget, 1e-9);
  EXPECT_NEAR(r.values[south][1][0].as_number(), gadget + widget, 1e-9);
}

TEST(PivotEvaluator, ShowAsPercentOfBaseItem) {
  PivotCache cache = build_diff_from_cache();
  PivotTable table = build_diff_from_table();
  auto& df = table.mutable_data_fields()[0];
  df.show_as = ShowValuesAs::PercentOf;
  df.show_as_base_field = 0U;
  df.show_as_base_item = 1U;  // items[1] = "B"

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.values.size(), 3U);
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 10.0 / 25.0);
  EXPECT_DOUBLE_EQ(r.values[1][0][0].as_number(), 1.0);  // The base item is c / c.
  EXPECT_DOUBLE_EQ(r.values[2][0][0].as_number(), 40.0 / 25.0);
}

TEST(PivotEvaluator, ShowAsPercentDifferenceFromFirstItemPreviousIsItself) {
  PivotCache cache = build_diff_from_cache();
  PivotTable table = build_diff_from_table();
  auto& df = table.mutable_data_fields()[0];
  df.show_as = ShowValuesAs::PercentOf;
  df.show_as_base_field = 0U;
  df.show_as_base_item = kShowAsBasePrev;

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 1.0);  // The first item's previous is itself.
  EXPECT_DOUBLE_EQ(r.values[1][0][0].as_number(), 25.0 / 10.0);
}

TEST(PivotEvaluator, ShowAsPercentOfRunningTotalDividesByTheTotalOverTheBaseField) {
  PivotCache cache = build_diff_from_cache();
  PivotTable table = build_diff_from_table();
  auto& df = table.mutable_data_fields()[0];
  df.show_as = ShowValuesAs::PercentOfRunningTotal;
  df.show_as_base_field = 0U;

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 10.0 / 75.0);
  EXPECT_DOUBLE_EQ(r.values[1][0][0].as_number(), 35.0 / 75.0);
  EXPECT_DOUBLE_EQ(r.values[2][0][0].as_number(), 1.0);
}

TEST(PivotEvaluator, ShowAsRankIsDenseAndTiesShareARank) {
  PivotCache cache = build_diff_from_cache();
  // A and B tie at 10; C is 40.
  cache.mutable_records()[1].cells[1] = Value::number(10.0);
  PivotTable table = build_diff_from_table();
  auto& df = table.mutable_data_fields()[0];
  df.show_as = ShowValuesAs::RankDescending;
  df.show_as_base_field = 0U;

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  EXPECT_DOUBLE_EQ(r_or.value().values[0][0][0].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(r_or.value().values[1][0][0].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(r_or.value().values[2][0][0].as_number(), 1.0);

  table.mutable_data_fields()[0].show_as = ShowValuesAs::RankAscending;
  auto asc_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(asc_or)) << asc_or.error().message;
  EXPECT_DOUBLE_EQ(asc_or.value().values[0][0][0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(asc_or.value().values[1][0][0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(asc_or.value().values[2][0][0].as_number(), 2.0);
}

TEST(PivotEvaluator, ShowAsPercentOfParentRowIsBlankWhenTheParentIsZero) {
  PivotCache cache = build_show_as_cache();
  // North sums to 100; zero out its records so the parent is 0.
  for (std::size_t i = 0; i < 4; ++i) {
    cache.mutable_records()[i].cells[3] = Value::number(0.0);
  }
  PivotTable table = build_show_as_table(ShowValuesAs::PercentOfParentRow, /*grand_totals=*/true);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  // North / Gadget: the North subtotal in the same column is 0 -> blank (not #DIV/0!).
  EXPECT_TRUE(r.values[0][0][0].is_blank());
  // South / Gadget Online: 70 / South Online subtotal 120.
  EXPECT_DOUBLE_EQ(r.values[2][0][0].as_number(), 70.0 / 120.0);
}

}  // namespace
}  // namespace formulon::pivot
