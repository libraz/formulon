// Evaluator value-filter compaction and leaf-index invariants.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_evaluator.h"
#include "pivot/pivot_layout.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot/record_access.h"
#include "pivot_evaluator_fixtures.h"
#include "utils/checked_index.h"
#include "utils/error.h"
#include "value.h"

namespace formulon::pivot {
namespace {

using test::build_basic_cache;
using test::build_sum_amount_table;
using test::owned_text;
using test::row_index;

PivotCache build_subtotal_grid_cache() {
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Region", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Product", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Year", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Quarter", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});

  const struct {
    const char* region;
    const char* product;
    double base;
  } rows[] = {{"North", "A", 1.0}, {"North", "B", 2.0}, {"South", "A", 10.0}, {"South", "B", 20.0}};
  const struct {
    const char* year;
    const char* quarter;
    double multiplier;
  } cols[] = {{"2024", "Q1", 1.0}, {"2024", "Q2", 2.0}, {"2025", "Q1", 10.0}, {"2025", "Q2", 20.0}};

  for (const auto& row : rows) {
    for (const auto& col : cols) {
      PivotCacheRecord rec;
      rec.cells.push_back(owned_text(cache, row.region));
      rec.cells.push_back(owned_text(cache, row.product));
      rec.cells.push_back(owned_text(cache, col.year));
      rec.cells.push_back(owned_text(cache, col.quarter));
      rec.cells.push_back(Value::number(row.base * col.multiplier));
      cache.mutable_records().push_back(std::move(rec));
    }
  }
  return cache;
}

// Table over `build_subtotal_grid_cache()`: rows Region -> Product, columns
// Year -> Quarter, one SUM(Amount) data field.
PivotTable build_subtotal_grid_table() {
  PivotTable table;
  table.set_pivot_cache_id(1);
  const char* names[] = {"Region", "Product", "Year", "Quarter", "Amount"};
  for (std::size_t i = 0; i < 5; ++i) {
    PivotField field;
    field.source_name = names[i];
    field.axis = i == 4 ? PivotAxis::Value : PivotAxis::Row;
    table.mutable_fields().push_back(std::move(field));
  }
  table.mutable_row_field_order() = {0, 1};
  table.mutable_col_field_order() = {2, 3};
  PivotDataField sum;
  sum.name = "Sum of Amount";
  sum.field_index = 4;
  sum.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum));
  return table;
}

TEST(PivotEvaluator, ValueTop10FilterSaturatesOutOfDomainCounts) {
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
  add("B", 50.0);
  add("C", 30.0);
  add("D", 20.0);

  // Returns how many row leaves survive a Top-N filter asking for `requested`.
  auto rows_kept = [&cache](double requested) -> std::size_t {
    PivotTable table;
    table.set_pivot_cache_id(1);
    PivotField rf;
    rf.source_name = "Region";
    rf.axis = PivotAxis::Row;
    PivotField af;
    af.source_name = "Amount";
    af.axis = PivotAxis::Value;
    table.mutable_fields().push_back(std::move(rf));
    table.mutable_fields().push_back(std::move(af));
    table.mutable_row_field_order() = {0};
    PivotDataField sum;
    sum.name = "Sum of Amount";
    sum.field_index = 1;
    sum.aggregation = Aggregation::Sum;
    table.mutable_data_fields().push_back(std::move(sum));
    table.set_grand_totals(/*rows=*/false, /*cols=*/false);

    PivotFilter f;
    f.axis = PivotAxis::Row;
    f.field_name = "Region";
    f.type = FilterType::ValueTop10;
    f.value = requested;
    table.mutable_active_filters().push_back(f);

    auto r_or = evaluate(table, cache);
    EXPECT_TRUE(static_cast<bool>(r_or)) << "requested=" << requested;
    if (!r_or) {
      return 0;
    }
    return r_or.value().rows.size();
  };

  // NaN loses every comparison and a negative count asks for nothing.
  EXPECT_EQ(rows_kept(std::numeric_limits<double>::quiet_NaN()), 0U);
  EXPECT_EQ(rows_kept(-1.0), 0U);
  EXPECT_EQ(rows_kept(0.0), 0U);
  // A count at or beyond the axis keeps every leaf.
  EXPECT_EQ(rows_kept(1e30), 4U);
  EXPECT_EQ(rows_kept(std::numeric_limits<double>::infinity()), 4U);
  EXPECT_EQ(rows_kept(4.0), 4U);
  // And an ordinary count still ranks.
  EXPECT_EQ(rows_kept(2.0), 2U);
}

TEST(PivotEvaluator, RowValueFilterCompactsEveryLeafIndexedStructure) {
  PivotCache cache = build_subtotal_grid_cache();
  PivotTable table = build_subtotal_grid_table();

  // Region scores are North=99 and South=990, so Top-1 keeps both South
  // leaves and prunes the whole North group.
  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Region";
  f.type = FilterType::ValueTop10;
  f.value = 1;
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Two surviving row leaves; every row-leaf-indexed structure agrees.
  const std::size_t surviving_rows = 2U;
  ASSERT_EQ(r.values.size(), surviving_rows);
  EXPECT_EQ(r.row_leaf_totals.size(), surviving_rows);
  ASSERT_EQ(r.col_subtotals.size(), 2U);  // One per Year.
  for (const ColSubtotal& col_subtotal : r.col_subtotals) {
    EXPECT_EQ(col_subtotal.values.size(), surviving_rows);
  }

  // The pruned group's subtotal is gone from both the metadata-rich list
  // and the compact compatibility surface.
  ASSERT_EQ(r.row_subtotals.size(), 1U);
  ASSERT_EQ(r.row_subtotals[0].labels.size(), 1U);
  EXPECT_EQ(r.row_subtotals[0].labels[0], "South");
  EXPECT_EQ(r.subtotals.size(), 1U);
  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, "South");

  // Cross-axis structures on the surviving subtotal keep the column index
  // space untouched.
  EXPECT_EQ(r.row_subtotals[0].col_values.size(), 4U);
  EXPECT_EQ(r.row_subtotals[0].col_subtotal_values.size(), r.col_subtotals.size());

  // A column subtotal now reads the surviving rows: South/A over 2024 is
  // 10*(1+2)=30 and South/B is 20*(1+2)=60. Reading the pre-filter index
  // space would surface the pruned North rows (3 and 6).
  ASSERT_EQ(r.col_subtotals[0].values[0].size(), 1U);
  EXPECT_DOUBLE_EQ(r.col_subtotals[0].values[0][0].as_number(), 30.0);
  EXPECT_DOUBLE_EQ(r.col_subtotals[0].values[1][0].as_number(), 60.0);
  EXPECT_DOUBLE_EQ(r.col_subtotals[1].values[0][0].as_number(), 300.0);
  EXPECT_DOUBLE_EQ(r.col_subtotals[1].values[1][0].as_number(), 600.0);
}

TEST(PivotEvaluator, ColValueFilterCompactsEveryLeafIndexedStructure) {
  PivotCache cache = build_subtotal_grid_cache();
  PivotTable table = build_subtotal_grid_table();

  // Year scores are 2024=99 and 2025=990, so Top-1 keeps the 2025 quarters
  // and prunes the whole 2024 group.
  PivotFilter f;
  f.axis = PivotAxis::Col;
  f.field_name = "Year";
  f.type = FilterType::ValueTop10;
  f.value = 1;
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Two surviving column leaves; every column-leaf-indexed structure agrees.
  const std::size_t surviving_cols = 2U;
  ASSERT_EQ(r.values.size(), 4U);  // The row axis is untouched.
  for (const auto& row_slot : r.values) {
    EXPECT_EQ(row_slot.size(), surviving_cols);
  }
  EXPECT_EQ(r.col_leaf_totals.size(), surviving_cols);
  ASSERT_EQ(r.row_subtotals.size(), 2U);  // One per Region.
  for (const RowSubtotal& row_subtotal : r.row_subtotals) {
    EXPECT_EQ(row_subtotal.col_values.size(), surviving_cols);
  }

  // The pruned group's column subtotal is gone, and so is its slot in every
  // row x column subtotal intersection.
  ASSERT_EQ(r.col_subtotals.size(), 1U);
  ASSERT_EQ(r.col_subtotals[0].labels.size(), 1U);
  EXPECT_EQ(r.col_subtotals[0].labels[0], "2025");
  for (const RowSubtotal& row_subtotal : r.row_subtotals) {
    EXPECT_EQ(row_subtotal.col_subtotal_values.size(), r.col_subtotals.size());
  }
  ASSERT_EQ(r.cols.size(), 1U);
  EXPECT_EQ(r.cols[0].label, "2025");

  // North's subtotal row now reads the surviving columns: (1+2)*10=30 at
  // 2025/Q1 and (1+2)*20=60 at 2025/Q2, with 90 at the 2025 subtotal column.
  const RowSubtotal& north = r.row_subtotals[0];
  ASSERT_EQ(north.labels.size(), 1U);
  ASSERT_EQ(north.labels[0], "North");
  ASSERT_EQ(north.col_values[0].size(), 1U);
  EXPECT_DOUBLE_EQ(north.col_values[0][0].as_number(), 30.0);
  EXPECT_DOUBLE_EQ(north.col_values[1][0].as_number(), 60.0);
  EXPECT_DOUBLE_EQ(north.col_subtotal_values[0][0].as_number(), 90.0);
}

TEST(PivotEvaluator, ThreeLevelRowAxisNumbersLeavesAndSubtotalsPerDepth) {
  PivotCache cache = build_subtotal_grid_cache();
  PivotTable table = build_subtotal_grid_table();
  table.mutable_row_field_order() = {0, 1, 2};  // Region / Product / Year.
  table.mutable_col_field_order() = {3};        // Quarter.
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // 2 regions x 2 products x 2 years = 8 leaves, 2 quarter columns.
  ASSERT_EQ(r.rows.size(), 2U);
  ASSERT_EQ(r.rows[0].label, "North");
  ASSERT_EQ(r.rows[0].children.size(), 2U);
  ASSERT_EQ(r.rows[0].children[0].label, "A");
  ASSERT_EQ(r.rows[0].children[0].children.size(), 2U);
  EXPECT_EQ(r.rows[0].children[0].children[0].label, "2024");
  EXPECT_EQ(r.rows[0].children[0].children[1].label, "2025");
  ASSERT_EQ(r.values.size(), 8U);
  ASSERT_EQ(r.values[0].size(), 2U);

  // Leaf order is DFS pre-order over the three levels; the record buckets
  // have to follow the same numbering or the values land on the wrong rows.
  // Cell value is `row base * quarter multiplier` (North/A base 1,
  // North/B 2, South/A 10, South/B 20; 2024 Q1/Q2 = 1/2, 2025 = 10/20).
  const std::vector<std::vector<double>> expected = {
      {1.0, 2.0}, {10.0, 20.0}, {2.0, 4.0}, {20.0, 40.0}, {10.0, 20.0}, {100.0, 200.0}, {20.0, 40.0}, {200.0, 400.0},
  };
  for (std::size_t leaf = 0; leaf < expected.size(); ++leaf) {
    for (std::size_t col = 0; col < 2U; ++col) {
      EXPECT_DOUBLE_EQ(r.values[leaf][col][0].as_number(), expected[leaf][col]) << "leaf=" << leaf << " col=" << col;
    }
  }

  // One subtotal per interior owner: four at Region/Product, two at Region,
  // emitted in the post-order the row walk produces.
  ASSERT_EQ(r.row_subtotals.size(), 6U);
  const std::vector<std::vector<std::string>> expected_labels = {
      {"North", "A"}, {"North", "B"}, {"North"}, {"South", "A"}, {"South", "B"}, {"South"},
  };
  const std::vector<double> expected_totals = {33.0, 66.0, 99.0, 330.0, 660.0, 990.0};
  for (std::size_t i = 0; i < r.row_subtotals.size(); ++i) {
    EXPECT_EQ(r.row_subtotals[i].labels, expected_labels[i]) << "subtotal=" << i;
    EXPECT_EQ(r.row_subtotals[i].depth, expected_labels[i].size() - 1U) << "subtotal=" << i;
    ASSERT_EQ(r.row_subtotals[i].values.size(), 1U);
    EXPECT_DOUBLE_EQ(r.row_subtotals[i].values[0].as_number(), expected_totals[i]) << "subtotal=" << i;
  }
}

TEST(PivotEvaluator, RowValueFilterCollapsesEmptiedBranchesOfAThreeLevelAxis) {
  PivotCache cache = build_subtotal_grid_cache();
  PivotTable table = build_subtotal_grid_table();
  table.mutable_row_field_order() = {0, 1, 2};
  table.mutable_col_field_order() = {3};
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  // Year scores across the two quarter columns are 3, 30, 6, 60, 30, 300,
  // 60, 600. Greater-than-100 keeps South/A/2025 and South/B/2025 only, so
  // all of North and both 2024 leaves under South disappear.
  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Year";
  f.type = FilterType::ValueGreaterThan;
  f.value = 100.0;
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // The North branch is gone at every depth; South keeps both products but
  // only their 2025 child.
  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, "South");
  ASSERT_EQ(r.rows[0].children.size(), 2U);
  for (const AxisHierarchyNode& product : r.rows[0].children) {
    ASSERT_EQ(product.children.size(), 1U);
    EXPECT_EQ(product.children[0].label, "2025");
  }

  ASSERT_EQ(r.values.size(), 2U);
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 100.0);  // South/A/2025 Q1.
  EXPECT_DOUBLE_EQ(r.values[0][1][0].as_number(), 200.0);  // South/A/2025 Q2.
  EXPECT_DOUBLE_EQ(r.values[1][0][0].as_number(), 200.0);  // South/B/2025 Q1.
  EXPECT_DOUBLE_EQ(r.values[1][1][0].as_number(), 400.0);  // South/B/2025 Q2.
  EXPECT_EQ(r.row_leaf_totals.size(), 2U);

  // Only the owners that still cover a surviving leaf keep a subtotal, and
  // both the metadata-rich list and the compact surface agree.
  ASSERT_EQ(r.row_subtotals.size(), 3U);
  const std::vector<std::vector<std::string>> expected_labels = {{"South", "A"}, {"South", "B"}, {"South"}};
  for (std::size_t i = 0; i < r.row_subtotals.size(); ++i) {
    EXPECT_EQ(r.row_subtotals[i].labels, expected_labels[i]) << "subtotal=" << i;
  }
  EXPECT_EQ(r.subtotals.size(), 3U);
  // A surviving subtotal is re-aggregated from just its surviving leaves,
  // matching Excel's own Top-N grand total (verified against
  // pivot_value_date_filters.xlsx / pivot_recurring_period_filter.xlsx):
  // South/A = 100 + 200 = 300; South (both products) = 300 + 200 + 400 = 900.
  EXPECT_DOUBLE_EQ(r.row_subtotals[0].values[0].as_number(), 300.0);
  EXPECT_DOUBLE_EQ(r.row_subtotals[2].values[0].as_number(), 900.0);
}

TEST(PivotEvaluator, ValueFilterRecalculatesGrandTotalFromSurvivingLeavesOnly) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/true, /*cols=*/true);

  // North totals 175 (100+50+25), South totals 500 (200+300). Top-1
  // keeps South only.
  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Region";
  f.type = FilterType::ValueTop10;
  f.value = 1;
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, "South");
  ASSERT_FALSE(r.grand_totals.empty());
  EXPECT_DOUBLE_EQ(r.grand_totals[0].as_number(), 500.0);
  ASSERT_TRUE(r.grand_total.is_number());
  EXPECT_DOUBLE_EQ(r.grand_total.as_number(), 500.0);
}

TEST(CheckedIndex, RejectsEveryDoubleOutsideTheContainerBound) {
  constexpr double kNaN = std::numeric_limits<double>::quiet_NaN();
  constexpr double kInf = std::numeric_limits<double>::infinity();
  constexpr std::size_t kLimit = 4;

  EXPECT_FALSE(index_from_double(kNaN, kLimit).has_value());
  EXPECT_FALSE(index_from_double(kInf, kLimit).has_value());
  EXPECT_FALSE(index_from_double(-kInf, kLimit).has_value());
  EXPECT_FALSE(index_from_double(-1.0, kLimit).has_value());
  EXPECT_FALSE(index_from_double(static_cast<double>(kLimit), kLimit).has_value());
  EXPECT_FALSE(index_from_double(static_cast<double>(kLimit) + 1.0, kLimit).has_value());
  EXPECT_FALSE(index_from_double(1e30, kLimit).has_value());
  EXPECT_FALSE(index_from_double(9007199254740992.0, kLimit).has_value());  // 2^53.

  // In-domain values, including the negative zero that compares equal to 0.
  EXPECT_EQ(index_from_double(0.0, kLimit), std::optional<std::size_t>{0});
  EXPECT_EQ(index_from_double(-0.0, kLimit), std::optional<std::size_t>{0});
  EXPECT_EQ(index_from_double(0.5, kLimit), std::optional<std::size_t>{0});
  EXPECT_EQ(index_from_double(static_cast<double>(kLimit) - 1.0, kLimit), std::optional<std::size_t>{kLimit - 1});

  // An empty container has no valid index at all.
  EXPECT_FALSE(index_from_double(0.0, 0U).has_value());
  EXPECT_FALSE(index_from_double(kNaN, 0U).has_value());
}

}  // namespace
}  // namespace formulon::pivot
