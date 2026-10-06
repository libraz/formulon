// Aggregate value filters, including axis recovery and multi-level pruning.

#include <cstddef>
#include <cstdint>
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
#include "pivot_evaluator_fixtures.h"
#include "value.h"

namespace formulon::pivot {
namespace {

using test::build_basic_cache;
using test::build_sum_amount_table;
using test::owned_text;
using test::row_index;

// Helper: builds a 2-column cache where each row label has a distinct
// numeric amount. With one record per row the post-aggregation row score
// equals the record's Amount.
PivotCache build_amount_cache_for_between() {
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
  add("D", 75.0);
  return cache;
}

struct RegionProductAmount {
  const char* region;
  const char* product;
  double amount;
};

PivotCache build_region_product_cache(const std::vector<RegionProductAmount>& records) {
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Region", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Product", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});
  for (const RegionProductAmount& record : records) {
    PivotCacheRecord rec;
    rec.cells.push_back(owned_text(cache, record.region));
    rec.cells.push_back(owned_text(cache, record.product));
    rec.cells.push_back(Value::number(record.amount));
    cache.mutable_records().push_back(std::move(rec));
  }
  return cache;
}

PivotTable build_region_product_table(Aggregation aggregation) {
  PivotTable table;
  table.set_pivot_cache_id(1);
  const char* names[] = {"Region", "Product", "Amount"};
  for (std::size_t i = 0; i < 3; ++i) {
    PivotField field;
    field.source_name = names[i];
    field.axis = i == 2 ? PivotAxis::Value : PivotAxis::Row;
    table.mutable_fields().push_back(std::move(field));
  }
  table.mutable_row_field_order() = {0, 1};
  PivotDataField data;
  data.name = "Amount";
  data.field_index = 2;
  data.aggregation = aggregation;
  table.mutable_data_fields().push_back(std::move(data));
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  return table;
}

// A filter on the outer field ranks whole regions by their own total and

TEST(PivotEvaluator, AuthoredTopCountKeepsTheHighestScoringLeaf) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  AuthoredValueFilter f;
  f.field_index = 0;
  f.type = FilterType::ValueTop10;
  f.value = 1.0;
  table.mutable_authored_value_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().rows.size(), 1U);
  EXPECT_EQ(r_or.value().rows[0].label, "South");
}

TEST(PivotEvaluator, AuthoredTopCountWithTopFalseKeepsTheLowestScoringLeaf) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  AuthoredValueFilter f;
  f.field_index = 0;
  f.type = FilterType::ValueTop10;
  f.value = 1.0;
  f.top = false;
  table.mutable_authored_value_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().rows.size(), 1U);
  EXPECT_EQ(r_or.value().rows[0].label, "North");
}

TEST(PivotEvaluator, AuthoredGreaterThanIsStrictOnTheThreshold) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  AuthoredValueFilter f;
  f.field_index = 0;
  f.type = FilterType::ValueGreaterThan;
  f.value = 175.0;  // exactly North's total
  table.mutable_authored_value_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().rows.size(), 1U);
  EXPECT_EQ(r_or.value().rows[0].label, "South");
}

TEST(PivotEvaluator, AuthoredBetweenIncludesBothBounds) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  AuthoredValueFilter f;
  f.field_index = 0;
  f.type = FilterType::ValueBetween;
  f.value = 175.0;
  f.value_high = 500.0;
  table.mutable_authored_value_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  EXPECT_EQ(r_or.value().rows.size(), 2U);
}

TEST(PivotEvaluator, AuthoredValueFilterOnAnOffAxisFieldIsInert) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  AuthoredValueFilter f;
  f.field_index = 2;  // Amount: a data field, on no axis
  f.type = FilterType::ValueTop10;
  f.value = 1.0;
  table.mutable_authored_value_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  EXPECT_EQ(r_or.value().rows.size(), 2U);
}

TEST(PivotEvaluator, AuthoredValueFilterPrunesTheColumnAxisToo) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{}, /*col=*/{0});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  AuthoredValueFilter f;
  f.field_index = 0;
  f.type = FilterType::ValueTop10;
  f.value = 1.0;
  table.mutable_authored_value_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().cols.size(), 1U);
  EXPECT_EQ(r_or.value().cols[0].label, "South");
}

TEST(PivotEvaluator, ValueTop10FilterKeepsTopRows) {
  // Use a richer cache so ranking is meaningful.
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
  f.value = 2;  // Top 2.
  table.mutable_active_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Top 2 by amount: B (50), C (30).
  ASSERT_EQ(r.rows.size(), 2U);
  // Rows are still in alphabetical order after pruning (we kept only the
  // surviving leaves' positions); order: B, C.
  std::vector<std::string> labels;
  std::vector<double> totals;
  for (std::size_t i = 0; i < r.rows.size(); ++i) {
    labels.push_back(r.rows[i].label);
    totals.push_back(r.values[i][0][0].as_number());
  }
  // Sort to make assertions order-independent.
  std::sort(labels.begin(), labels.end());
  std::sort(totals.begin(), totals.end());
  EXPECT_EQ(labels[0], "B");
  EXPECT_EQ(labels[1], "C");
  EXPECT_DOUBLE_EQ(totals[0], 30.0);
  EXPECT_DOUBLE_EQ(totals[1], 50.0);
}

TEST(PivotEvaluator, ValueGreaterThanFilterKeepsAboveThreshold) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Region";
  f.type = FilterType::ValueGreaterThan;
  f.value = 200.0;  // Threshold; North (175) drops, South (500) survives.
  table.mutable_active_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, "South");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 500.0);
}

TEST(PivotEvaluator, ValueTop10FilterOnColAxis) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{1});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  PivotFilter f;
  f.axis = PivotAxis::Col;
  f.field_name = "Product";
  f.type = FilterType::ValueTop10;
  f.value = 1;  // Keep 1 column with the highest column total.
  table.mutable_active_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Widget total = 100 + 200 + 25 = 325; Gadget total = 50 + 300 = 350.
  // Top-1 by col total -> Gadget survives.
  ASSERT_EQ(r.cols.size(), 1U);
  EXPECT_EQ(r.cols[0].label, "Gadget");
  // Each row keeps a single col slot.
  ASSERT_EQ(r.values.size(), 2U);
  for (const auto& row_slot : r.values) {
    ASSERT_EQ(row_slot.size(), 1U);
  }
}

TEST(PivotEvaluator, ValueFilterSelectsDataFieldAndInvalidSelectorIsNoOp) {
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Region", {}});
  cache.mutable_fields().push_back(PivotCacheField{"MeasureA", {}});
  cache.mutable_fields().push_back(PivotCacheField{"MeasureB", {}});
  auto add = [&](const char* region, double a, double b) {
    PivotCacheRecord rec;
    rec.cells.push_back(owned_text(cache, region));
    rec.cells.push_back(Value::number(a));
    rec.cells.push_back(Value::number(b));
    cache.mutable_records().push_back(std::move(rec));
  };
  add("A", 100.0, 1.0);
  add("B", 50.0, 100.0);

  PivotTable table;
  table.set_pivot_cache_id(1);
  PivotField region_field;
  region_field.source_name = "Region";
  region_field.axis = PivotAxis::Row;
  PivotField measure_a_field;
  measure_a_field.source_name = "MeasureA";
  measure_a_field.axis = PivotAxis::Value;
  PivotField measure_b_field;
  measure_b_field.source_name = "MeasureB";
  measure_b_field.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(region_field));
  table.mutable_fields().push_back(std::move(measure_a_field));
  table.mutable_fields().push_back(std::move(measure_b_field));
  table.mutable_row_field_order() = {0};
  PivotDataField measure_a;
  measure_a.name = "A";
  measure_a.field_index = 1;
  measure_a.aggregation = Aggregation::Sum;
  PivotDataField measure_b;
  measure_b.name = "B";
  measure_b.field_index = 2;
  measure_b.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(measure_a));
  table.mutable_data_fields().push_back(std::move(measure_b));

  PivotFilter selected;
  selected.axis = PivotAxis::Row;
  selected.field_name = "Region";
  selected.type = FilterType::ValueTop10;
  selected.value = 1;
  selected.data_field_index = 1;
  table.mutable_active_filters().push_back(selected);

  auto selected_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(selected_or)) << selected_or.error().message;
  ASSERT_EQ(selected_or.value().rows.size(), 1U);
  EXPECT_EQ(selected_or.value().rows[0].label, "B");

  // A direct C++ model can still contain a stale selector; evaluation must
  // leave the report untouched rather than dropping every leaf.
  table.mutable_active_filters()[0].data_field_index = 2;
  auto invalid_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(invalid_or)) << invalid_or.error().message;
  ASSERT_EQ(invalid_or.value().rows.size(), 2U);
  EXPECT_EQ(invalid_or.value().rows[0].label, "A");
  EXPECT_EQ(invalid_or.value().rows[1].label, "B");
}

TEST(PivotEvaluator, ValueBetweenFilterRowAxis) {
  PivotCache cache = build_amount_cache_for_between();
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
  f.type = FilterType::ValueBetween;
  f.value = 20.0;       // inclusive low bound
  f.value_high = 50.0;  // inclusive high bound
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // 25 and 40 land in [20, 50]; 10 and 75 fall outside.
  ASSERT_EQ(r.rows.size(), 2U);
  std::vector<std::string> labels;
  std::vector<double> totals;
  for (std::size_t i = 0; i < r.rows.size(); ++i) {
    labels.push_back(r.rows[i].label);
    totals.push_back(r.values[i][0][0].as_number());
  }
  std::sort(labels.begin(), labels.end());
  std::sort(totals.begin(), totals.end());
  EXPECT_EQ(labels[0], "B");
  EXPECT_EQ(labels[1], "C");
  EXPECT_DOUBLE_EQ(totals[0], 25.0);
  EXPECT_DOUBLE_EQ(totals[1], 40.0);
}

TEST(PivotEvaluator, ValueBetweenFilterColAxis) {
  // Mirror of the row-axis test along the column axis: each Region
  // becomes a column with its single Amount as the column total.
  PivotCache cache = build_amount_cache_for_between();
  PivotTable table;
  table.set_pivot_cache_id(1);
  PivotField rf;
  rf.source_name = "Region";
  rf.axis = PivotAxis::Col;
  PivotField af;
  af.source_name = "Amount";
  af.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(rf));
  table.mutable_fields().push_back(std::move(af));
  table.mutable_col_field_order() = {0};
  PivotDataField sum;
  sum.name = "Sum of Amount";
  sum.field_index = 1;
  sum.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum));
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  PivotFilter f;
  f.axis = PivotAxis::Col;
  f.field_name = "Region";
  f.type = FilterType::ValueBetween;
  f.value = 20.0;
  f.value_high = 50.0;
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  ASSERT_EQ(r.cols.size(), 2U);
  std::vector<std::string> labels{r.cols[0].label, r.cols[1].label};
  std::sort(labels.begin(), labels.end());
  EXPECT_EQ(labels[0], "B");
  EXPECT_EQ(labels[1], "C");
  // The single implicit row should retain exactly the two surviving cols.
  ASSERT_EQ(r.values.size(), 1U);
  ASSERT_EQ(r.values[0].size(), 2U);
}

TEST(PivotEvaluator, ValueBetweenFilterUnboundedHighIsNoOp) {
  PivotCache cache = build_amount_cache_for_between();
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
  f.type = FilterType::ValueBetween;
  f.value = 20.0;
  // value_high left as default monostate -> filter degrades to no-op.
  f.value_high = std::monostate{};
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // No upper bound -> all four rows survive.
  ASSERT_EQ(r.rows.size(), 4U);
}

TEST(PivotEvaluator, ValueTop10MultiLevelRowAxis) {
  // Two row fields: Region (North/South) -> Product (A/B/C). Each
  // (region, product) leaf has a single record so the post-aggregation
  // leaf score equals its Amount.
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
  add("North", "B", 50.0);
  add("North", "C", 20.0);
  add("South", "A", 80.0);
  add("South", "B", 30.0);
  add("South", "C", 5.0);

  PivotTable table;
  table.set_pivot_cache_id(1);
  PivotField region_f;
  region_f.source_name = "Region";
  region_f.axis = PivotAxis::Row;
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
  PivotDataField sum;
  sum.name = "Sum of Amount";
  sum.field_index = 2;
  sum.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum));
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Product";
  f.type = FilterType::ValueTop10;
  f.value = 2;  // Top 2 products under each region.
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Ranked within each region, not across the flattened axis: North keeps
  // B=50 and C=20 even though South/B=30 outranks North/C globally, and
  // South keeps A=80 and B=30.
  ASSERT_EQ(r.rows.size(), 2U);
  ASSERT_EQ(r.rows[0].label, "North");
  ASSERT_EQ(r.rows[0].children.size(), 2U);
  EXPECT_EQ(r.rows[0].children[0].label, "B");
  EXPECT_EQ(r.rows[0].children[1].label, "C");
  ASSERT_EQ(r.rows[1].label, "South");
  ASSERT_EQ(r.rows[1].children.size(), 2U);
  EXPECT_EQ(r.rows[1].children[0].label, "A");
  EXPECT_EQ(r.rows[1].children[1].label, "B");

  ASSERT_EQ(r.values.size(), 4U);
  const std::vector<double> expected = {50.0, 20.0, 80.0, 30.0};
  for (std::size_t leaf = 0; leaf < expected.size(); ++leaf) {
    ASSERT_EQ(r.values[leaf].size(), 1U);
    EXPECT_DOUBLE_EQ(r.values[leaf][0][0].as_number(), expected[leaf]) << "leaf=" << leaf;
  }
}

TEST(PivotEvaluator, ValueFilterWithUnknownFieldIsNoOpOnNestedAxis) {
  PivotCache cache = build_region_product_cache(
      {{"North", "A", 10.0}, {"North", "B", 50.0}, {"South", "A", 80.0}, {"South", "B", 30.0}});
  PivotTable table = build_region_product_table(Aggregation::Sum);
  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Missing";
  f.type = FilterType::ValueTop10;
  f.value = 1;
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.rows.size(), 2U);
  ASSERT_EQ(r.rows[0].children.size(), 2U);
  ASSERT_EQ(r.rows[1].children.size(), 2U);
  EXPECT_EQ(r.values.size(), 4U);
  const std::vector<double> expected = {10.0, 50.0, 80.0, 30.0};
  for (std::size_t leaf = 0; leaf < expected.size(); ++leaf) {
    ASSERT_EQ(r.values[leaf].size(), 1U);
    EXPECT_DOUBLE_EQ(r.values[leaf][0][0].as_number(), expected[leaf]) << "leaf=" << leaf;
  }
}

TEST(PivotEvaluator, ValueTop10OnOuterFieldKeepsWholeGroups) {
  PivotCache cache = build_region_product_cache({{"North", "A", 10.0},
                                                 {"North", "B", 50.0},
                                                 {"North", "C", 20.0},
                                                 {"South", "A", 80.0},
                                                 {"South", "B", 30.0},
                                                 {"South", "C", 5.0}});
  PivotTable table = build_region_product_table(Aggregation::Sum);
  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Region";
  f.type = FilterType::ValueTop10;
  f.value = 1;  // North=80, South=115.
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, "South");
  EXPECT_EQ(r.rows[0].children.size(), 3U);
  EXPECT_EQ(r.values.size(), 3U);
}

TEST(PivotEvaluator, ValueTop10OnInnerFieldRanksPerParentWithoutSubtotals) {
  PivotCache cache = build_region_product_cache(
      {{"North", "A", 10.0}, {"North", "B", 50.0}, {"South", "A", 80.0}, {"South", "B", 30.0}});
  PivotTable table = build_region_product_table(Aggregation::Sum);
  table.mutable_fields()[0].default_subtotal = false;
  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Product";
  f.type = FilterType::ValueTop10;
  f.value = 1;
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  EXPECT_TRUE(r.row_subtotals.empty());
  ASSERT_EQ(r.rows.size(), 2U);
  ASSERT_EQ(r.rows[0].children.size(), 1U);
  EXPECT_EQ(r.rows[0].children[0].label, "B");  // North's best, below South/B.
  ASSERT_EQ(r.rows[1].children.size(), 1U);
  EXPECT_EQ(r.rows[1].children[0].label, "A");
}

TEST(PivotEvaluator, AuthoredTopCountOnInnerFieldRanksPerParent) {
  PivotCache cache = build_region_product_cache(
      {{"North", "A", 10.0}, {"North", "B", 50.0}, {"South", "A", 80.0}, {"South", "B", 30.0}});
  PivotTable table = build_region_product_table(Aggregation::Sum);
  AuthoredValueFilter f;
  f.field_index = 1;
  f.type = FilterType::ValueTop10;
  f.value = 1.0;
  table.mutable_authored_value_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.rows.size(), 2U);
  ASSERT_EQ(r.rows[0].children.size(), 1U);
  EXPECT_EQ(r.rows[0].children[0].label, "B");
  ASSERT_EQ(r.rows[1].children.size(), 1U);
  EXPECT_EQ(r.rows[1].children[0].label, "A");
}

TEST(PivotEvaluator, ValueTop10OnOuterFieldScoresNonAdditiveAggregate) {
  PivotCache cache = build_region_product_cache(
      {{"North", "A", 40.0}, {"North", "B", 40.0}, {"North", "C", 40.0}, {"South", "A", 50.0}});
  PivotTable table = build_region_product_table(Aggregation::Average);
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
}

TEST(PivotEvaluator, ValueGreaterThanMultiLevelColAxis) {
  // Two column fields: Year (2024/2025) -> Quarter (Q1/Q2). Single
  // implicit row, so each col leaf's score is just its Amount.
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Year", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Quarter", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});
  auto add = [&](const char* year, const char* quarter, double amount) {
    PivotCacheRecord rec;
    rec.cells.push_back(owned_text(cache, year));
    rec.cells.push_back(owned_text(cache, quarter));
    rec.cells.push_back(Value::number(amount));
    cache.mutable_records().push_back(std::move(rec));
  };
  // Leaf order in DFS pre-order: 2024/Q1=15, 2024/Q2=40, 2025/Q1=5, 2025/Q2=100.
  add("2024", "Q1", 15.0);
  add("2024", "Q2", 40.0);
  add("2025", "Q1", 5.0);
  add("2025", "Q2", 100.0);

  PivotTable table;
  table.set_pivot_cache_id(1);
  PivotField year_f;
  year_f.source_name = "Year";
  year_f.axis = PivotAxis::Col;
  PivotField quarter_f;
  quarter_f.source_name = "Quarter";
  quarter_f.axis = PivotAxis::Col;
  PivotField amount_f;
  amount_f.source_name = "Amount";
  amount_f.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(year_f));
  table.mutable_fields().push_back(std::move(quarter_f));
  table.mutable_fields().push_back(std::move(amount_f));
  table.mutable_col_field_order() = {0, 1};
  PivotDataField sum;
  sum.name = "Sum of Amount";
  sum.field_index = 2;
  sum.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum));
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  PivotFilter f;
  f.axis = PivotAxis::Col;
  f.field_name = "Quarter";
  f.type = FilterType::ValueGreaterThan;
  f.value = 20.0;
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Surviving col leaves at indices 1 (2024/Q2=40) and 3 (2025/Q2=100).
  // Each Year keeps only Q2; both Year subtrees survive.
  ASSERT_EQ(r.cols.size(), 2U);
  EXPECT_EQ(r.cols[0].label, "2024");
  EXPECT_EQ(r.cols[1].label, "2025");
  ASSERT_EQ(r.cols[0].children.size(), 1U);
  ASSERT_EQ(r.cols[1].children.size(), 1U);
  EXPECT_EQ(r.cols[0].children[0].label, "Q2");
  EXPECT_EQ(r.cols[1].children[0].label, "Q2");

  ASSERT_EQ(r.values.size(), 1U);
  ASSERT_EQ(r.values[0].size(), 2U);
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 40.0);
  EXPECT_DOUBLE_EQ(r.values[0][1][0].as_number(), 100.0);
}

TEST(PivotEvaluator, ValueBetweenMultiLevelDropsEmptyParent) {
  // Two row fields: Region (North/South) -> Product (A/B). All North
  // leaves fall outside [10, 70], so North gets pruned entirely.
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
  add("North", "A", 5.0);
  add("North", "B", 8.0);
  add("South", "A", 50.0);
  add("South", "B", 60.0);

  PivotTable table;
  table.set_pivot_cache_id(1);
  PivotField region_f;
  region_f.source_name = "Region";
  region_f.axis = PivotAxis::Row;
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
  PivotDataField sum;
  sum.name = "Sum of Amount";
  sum.field_index = 2;
  sum.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum));
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Product";
  f.type = FilterType::ValueBetween;
  f.value = 10.0;
  f.value_high = 70.0;
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // North has no surviving children -> pruned entirely. Only South
  // remains, with both A and B.
  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, "South");
  ASSERT_EQ(r.rows[0].children.size(), 2U);
  EXPECT_EQ(r.rows[0].children[0].label, "A");
  EXPECT_EQ(r.rows[0].children[1].label, "B");

  ASSERT_EQ(r.values.size(), 2U);
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 50.0);
  EXPECT_DOUBLE_EQ(r.values[1][0][0].as_number(), 60.0);
}

}  // namespace
}  // namespace formulon::pivot
