// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Unit tests for `formulon::pivot::evaluate`. Each test hand-builds a
// `PivotCache` + `PivotTable` (no XML, no xlsx) and checks the produced
// `PivotResult` shape and per-cell values. The MVP scope mirrors the
// evaluator implementation: SUM / COUNT / AVERAGE / MAX / MIN / PRODUCT
// / CountNumbers, manual-filter visibility, hierarchy + grand total +
// row-axis subtotals.

#include "pivot/pivot_evaluator.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <string>
#include <variant>
#include <vector>

#include "eval/date_time.h"
#include "eval/groupby_pivotby/common.h"
#include "gtest/gtest.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot/value_order.h"
#include "utils/error.h"
#include "value.h"

namespace formulon::pivot {
namespace {

// ---------------------------------------------------------------------------
// Cross-subsystem comparator agreement: the pivot comparator
// (`pivot::value_less`) and the GROUPBY / SORT comparator
// (`eval::cmp_value_asc`) share one Excel kind rank, so a key column mixing
// Bool and Text must order identically in both. Excel's ascending order puts
// Text before Bool.
// ---------------------------------------------------------------------------

TEST(PivotComparatorParity, BoolVsTextMatchesGroupByOrder) {
  const Value text_val = Value::text("zebra");
  const Value bool_val = Value::boolean(false);

  // Pivot: Text sorts before Bool.
  EXPECT_TRUE(value_less(text_val, bool_val));
  EXPECT_FALSE(value_less(bool_val, text_val));

  // GROUPBY / SORT: same ordering (negative => first argument sorts first).
  EXPECT_LT(eval::cmp_value_asc(text_val, bool_val), 0);
  EXPECT_GT(eval::cmp_value_asc(bool_val, text_val), 0);
}

TEST(PivotComparatorParity, NumberBeforeTextBeforeBool) {
  const Value num = Value::number(1.0);
  const Value text_val = Value::text("a");
  const Value bool_val = Value::boolean(true);

  EXPECT_TRUE(value_less(num, text_val));
  EXPECT_TRUE(value_less(text_val, bool_val));

  EXPECT_LT(eval::cmp_value_asc(num, text_val), 0);
  EXPECT_LT(eval::cmp_value_asc(text_val, bool_val), 0);
}

// ---------------------------------------------------------------------------
// Fixture helpers
// ---------------------------------------------------------------------------

// Pushes `s` into the cache's text storage so the returned `Value::text`
// holds a stable view for the lifetime of the cache.
Value owned_text(PivotCache& cache, std::string s) {
  cache.mutable_text_storage().push_back(std::move(s));
  return Value::text(cache.text_storage().back());
}

// Builds a 3-column cache: Region (text), Product (text), Amount (number).
// Records form a small, easy-to-reason-about dataset.
//
//   Region  Product  Amount
//   ------  -------  ------
//   North   Widget    100
//   North   Gadget     50
//   South   Widget    200
//   South   Gadget    300
//   North   Widget     25
PivotCache build_basic_cache() {
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

  add("North", "Widget", 100.0);
  add("North", "Gadget", 50.0);
  add("South", "Widget", 200.0);
  add("South", "Gadget", 300.0);
  add("North", "Widget", 25.0);
  return cache;
}

// Builds a `PivotTable` matching `build_basic_cache()` with the supplied
// row/col field indices and a single SUM(Amount) data field.
PivotTable build_sum_amount_table(std::vector<std::uint32_t> row_fields, std::vector<std::uint32_t> col_fields) {
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

  PivotDataField sum_amount;
  sum_amount.name = "Sum of Amount";
  sum_amount.field_index = 2;  // Amount.
  sum_amount.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum_amount));

  table.mutable_row_field_order() = std::move(row_fields);
  table.mutable_col_field_order() = std::move(col_fields);
  return table;
}

// Convenience: locate a top-level row leaf by its label, return the
// index into `result.values`.
std::size_t row_index(const PivotResult& r, const std::string& label) {
  for (std::size_t i = 0; i < r.rows.size(); ++i) {
    if (r.rows[i].label == label) {
      return i;
    }
  }
  return static_cast<std::size_t>(-1);
}

// ---------------------------------------------------------------------------
// 1. Single SUM happy path
// ---------------------------------------------------------------------------

TEST(PivotEvaluator, SingleSumByRegion) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Two row leaves (North, South), one implicit col leaf, one data field.
  ASSERT_EQ(r.rows.size(), 2U);
  EXPECT_TRUE(r.cols.empty());
  ASSERT_EQ(r.values.size(), 2U);
  ASSERT_EQ(r.values[0].size(), 1U);
  ASSERT_EQ(r.values[0][0].size(), 1U);

  const std::size_t north = row_index(r, "North");
  const std::size_t south = row_index(r, "South");
  ASSERT_NE(north, static_cast<std::size_t>(-1));
  ASSERT_NE(south, static_cast<std::size_t>(-1));

  EXPECT_DOUBLE_EQ(r.values[north][0][0].as_number(), 175.0);  // 100 + 50 + 25
  EXPECT_DOUBLE_EQ(r.values[south][0][0].as_number(), 500.0);  // 200 + 300
  EXPECT_TRUE(r.grand_total.is_blank());
}

// ---------------------------------------------------------------------------
// 2. Two row fields
// ---------------------------------------------------------------------------

TEST(PivotEvaluator, TwoRowFieldsHierarchy) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0, 1}, /*col=*/{});

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Row hierarchy: North { Gadget, Widget }, South { Gadget, Widget }.
  ASSERT_EQ(r.rows.size(), 2U);
  for (const auto& region : r.rows) {
    EXPECT_FALSE(region.children.empty());
    EXPECT_EQ(region.children.size(), 2U);
  }
  // Each region should have both products.
  const std::size_t north = row_index(r, "North");
  const std::size_t south = row_index(r, "South");
  ASSERT_NE(north, static_cast<std::size_t>(-1));
  ASSERT_NE(south, static_cast<std::size_t>(-1));

  // Verify children are alphabetically ordered (Gadget < Widget).
  EXPECT_EQ(r.rows[north].children[0].label, "Gadget");
  EXPECT_EQ(r.rows[north].children[1].label, "Widget");

  // Four leaves in total: walk values matrix and check totals.
  // Leaves are ordered: North/Gadget, North/Widget, South/Gadget, South/Widget.
  ASSERT_EQ(r.values.size(), 4U);
  // We don't depend on the exact leaf ordering across regions; just
  // sum and check totals match {50, 125, 300, 200}.
  std::vector<double> got;
  got.reserve(4);
  for (const auto& row_slot : r.values) {
    ASSERT_EQ(row_slot.size(), 1U);
    ASSERT_EQ(row_slot[0].size(), 1U);
    got.push_back(row_slot[0][0].as_number());
  }
  std::sort(got.begin(), got.end());
  EXPECT_DOUBLE_EQ(got[0], 50.0);   // North/Gadget
  EXPECT_DOUBLE_EQ(got[1], 125.0);  // North/Widget (100 + 25)
  EXPECT_DOUBLE_EQ(got[2], 200.0);  // South/Widget
  EXPECT_DOUBLE_EQ(got[3], 300.0);  // South/Gadget
}

// ---------------------------------------------------------------------------
// 3. One row + one col field
// ---------------------------------------------------------------------------

TEST(PivotEvaluator, OneRowOneColField) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{1});

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  ASSERT_EQ(r.rows.size(), 2U);
  ASSERT_EQ(r.cols.size(), 2U);
  // 2 row leaves x 2 col leaves x 1 data field.
  ASSERT_EQ(r.values.size(), 2U);
  ASSERT_EQ(r.values[0].size(), 2U);
  ASSERT_EQ(r.values[0][0].size(), 1U);

  // Column ordering: Gadget < Widget alphabetically.
  EXPECT_EQ(r.cols[0].label, "Gadget");
  EXPECT_EQ(r.cols[1].label, "Widget");

  const std::size_t north = row_index(r, "North");
  const std::size_t south = row_index(r, "South");
  EXPECT_DOUBLE_EQ(r.values[north][0][0].as_number(), 50.0);   // North/Gadget
  EXPECT_DOUBLE_EQ(r.values[north][1][0].as_number(), 125.0);  // North/Widget
  EXPECT_DOUBLE_EQ(r.values[south][0][0].as_number(), 300.0);  // South/Gadget
  EXPECT_DOUBLE_EQ(r.values[south][1][0].as_number(), 200.0);  // South/Widget
}

// ---------------------------------------------------------------------------
// 4. Multiple aggregations on the same source
// ---------------------------------------------------------------------------

TEST(PivotEvaluator, SumAndAverageOnSameField) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  // Add an Average data field on the same source column.
  PivotDataField avg_amount;
  avg_amount.name = "Average of Amount";
  avg_amount.field_index = 2;
  avg_amount.aggregation = Aggregation::Average;
  table.mutable_data_fields().push_back(std::move(avg_amount));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  ASSERT_EQ(r.values.size(), 2U);
  ASSERT_EQ(r.values[0][0].size(), 2U);  // Sum + Average.

  const std::size_t north = row_index(r, "North");
  const std::size_t south = row_index(r, "South");
  EXPECT_DOUBLE_EQ(r.values[north][0][0].as_number(), 175.0);
  EXPECT_DOUBLE_EQ(r.values[north][0][1].as_number(), 175.0 / 3.0);
  EXPECT_DOUBLE_EQ(r.values[south][0][0].as_number(), 500.0);
  EXPECT_DOUBLE_EQ(r.values[south][0][1].as_number(), 250.0);
}

// ---------------------------------------------------------------------------
// 5. COUNT counts non-blank cells (including text and booleans)
// ---------------------------------------------------------------------------

TEST(PivotEvaluator, CountIncludesNonBlanks) {
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Group", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Note", {}});

  auto add = [&](const char* group, Value note) {
    PivotCacheRecord rec;
    rec.cells.push_back(owned_text(cache, group));
    rec.cells.push_back(note);
    cache.mutable_records().push_back(std::move(rec));
  };

  add("A", owned_text(cache, "ok"));   // counted
  add("A", Value::blank());            // NOT counted
  add("A", Value::number(7.0));        // counted
  add("B", owned_text(cache, "yes"));  // counted
  add("B", Value::blank());            // NOT counted

  PivotTable table;
  table.set_pivot_cache_id(1);
  PivotField group_f;
  group_f.source_name = "Group";
  group_f.axis = PivotAxis::Row;
  PivotField note_f;
  note_f.source_name = "Note";
  note_f.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(group_f));
  table.mutable_fields().push_back(std::move(note_f));
  table.mutable_row_field_order() = {0};
  PivotDataField cnt;
  cnt.name = "Count of Note";
  cnt.field_index = 1;
  cnt.aggregation = Aggregation::Count;
  table.mutable_data_fields().push_back(std::move(cnt));
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  ASSERT_EQ(r.rows.size(), 2U);
  const std::size_t a = row_index(r, "A");
  const std::size_t b = row_index(r, "B");
  EXPECT_DOUBLE_EQ(r.values[a][0][0].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(r.values[b][0][0].as_number(), 1.0);
}

// ---------------------------------------------------------------------------
// 6. CountNumbers excludes text
// ---------------------------------------------------------------------------

TEST(PivotEvaluator, CountNumbersExcludesText) {
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Group", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Mixed", {}});

  auto add = [&](const char* group, Value v) {
    PivotCacheRecord rec;
    rec.cells.push_back(owned_text(cache, group));
    rec.cells.push_back(v);
    cache.mutable_records().push_back(std::move(rec));
  };
  add("A", Value::number(1.0));
  add("A", owned_text(cache, "skip"));
  add("A", Value::number(2.0));
  add("A", owned_text(cache, "skip2"));

  PivotTable table;
  table.set_pivot_cache_id(1);
  PivotField group_f;
  group_f.source_name = "Group";
  group_f.axis = PivotAxis::Row;
  PivotField mixed_f;
  mixed_f.source_name = "Mixed";
  mixed_f.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(group_f));
  table.mutable_fields().push_back(std::move(mixed_f));
  table.mutable_row_field_order() = {0};
  PivotDataField cn;
  cn.name = "CountNumbers of Mixed";
  cn.field_index = 1;
  cn.aggregation = Aggregation::CountNumbers;
  table.mutable_data_fields().push_back(std::move(cn));
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 2.0);
}

// ---------------------------------------------------------------------------
// 7. AVERAGE on no-numbers returns #DIV/0!
// ---------------------------------------------------------------------------

TEST(PivotEvaluator, AverageNoNumbersReturnsDiv0) {
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Group", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Note", {}});

  auto add = [&](const char* g, const char* n) {
    PivotCacheRecord rec;
    rec.cells.push_back(owned_text(cache, g));
    rec.cells.push_back(owned_text(cache, n));
    cache.mutable_records().push_back(std::move(rec));
  };
  add("A", "x");
  add("A", "y");

  PivotTable table;
  table.set_pivot_cache_id(1);
  PivotField gf;
  gf.source_name = "Group";
  gf.axis = PivotAxis::Row;
  PivotField nf;
  nf.source_name = "Note";
  nf.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(gf));
  table.mutable_fields().push_back(std::move(nf));
  table.mutable_row_field_order() = {0};
  PivotDataField avg;
  avg.name = "Average of Note";
  avg.field_index = 1;
  avg.aggregation = Aggregation::Average;
  table.mutable_data_fields().push_back(std::move(avg));
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  ASSERT_EQ(r.rows.size(), 1U);
  ASSERT_TRUE(r.values[0][0][0].is_error());
  EXPECT_EQ(r.values[0][0][0].as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// 8. MAX / MIN / PRODUCT arithmetic
// ---------------------------------------------------------------------------

TEST(PivotEvaluator, MaxMinProduct) {
  PivotCache cache = build_basic_cache();

  auto run = [&](Aggregation agg) {
    PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
    table.mutable_data_fields()[0].aggregation = agg;
    table.set_grand_totals(/*rows=*/false, /*cols=*/false);
    auto r_or = evaluate(table, cache);
    EXPECT_TRUE(static_cast<bool>(r_or));
    return r_or.value();
  };

  {
    PivotResult r = run(Aggregation::Max);
    const std::size_t north = row_index(r, "North");
    const std::size_t south = row_index(r, "South");
    EXPECT_DOUBLE_EQ(r.values[north][0][0].as_number(), 100.0);
    EXPECT_DOUBLE_EQ(r.values[south][0][0].as_number(), 300.0);
  }
  {
    PivotResult r = run(Aggregation::Min);
    const std::size_t north = row_index(r, "North");
    const std::size_t south = row_index(r, "South");
    EXPECT_DOUBLE_EQ(r.values[north][0][0].as_number(), 25.0);
    EXPECT_DOUBLE_EQ(r.values[south][0][0].as_number(), 200.0);
  }
  {
    PivotResult r = run(Aggregation::Product);
    const std::size_t north = row_index(r, "North");
    const std::size_t south = row_index(r, "South");
    // North: 100 * 50 * 25 = 125_000
    EXPECT_DOUBLE_EQ(r.values[north][0][0].as_number(), 125000.0);
    // South: 200 * 300 = 60_000
    EXPECT_DOUBLE_EQ(r.values[south][0][0].as_number(), 60000.0);
  }
}

// ---------------------------------------------------------------------------
// 8b. StdDev / StdDevP / Var / VarP arithmetic
// ---------------------------------------------------------------------------
//
// `build_basic_cache()` per region:
//   North: {100, 50, 25}            mean = 175 / 3
//   South: {200, 300}               mean = 250

TEST(PivotEvaluator, StdDevAndVarFamily) {
  PivotCache cache = build_basic_cache();

  auto run = [&](Aggregation agg) {
    PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
    table.mutable_data_fields()[0].aggregation = agg;
    table.set_grand_totals(/*rows=*/false, /*cols=*/false);
    auto r_or = evaluate(table, cache);
    EXPECT_TRUE(static_cast<bool>(r_or));
    return r_or.value();
  };

  // North n=3, mean = 175/3, ss = sum((x - mean)^2)
  //   = (100 - 175/3)^2 + (50 - 175/3)^2 + (25 - 175/3)^2
  //   = (125/3)^2 + (-25/3)^2 + (-100/3)^2
  //   = 15625/9 + 625/9 + 10000/9
  //   = 26250/9
  const double north_ss =
      (125.0 / 3.0) * (125.0 / 3.0) + (-25.0 / 3.0) * (-25.0 / 3.0) + (-100.0 / 3.0) * (-100.0 / 3.0);
  // South n=2, mean = 250, ss = (200-250)^2 + (300-250)^2 = 5000
  const double south_ss = 5000.0;

  {
    PivotResult r = run(Aggregation::Var);  // sample, divisor n-1
    const std::size_t north = row_index(r, "North");
    const std::size_t south = row_index(r, "South");
    EXPECT_DOUBLE_EQ(r.values[north][0][0].as_number(), north_ss / 2.0);
    EXPECT_DOUBLE_EQ(r.values[south][0][0].as_number(), south_ss / 1.0);
  }
  {
    PivotResult r = run(Aggregation::VarP);  // population, divisor n
    const std::size_t north = row_index(r, "North");
    const std::size_t south = row_index(r, "South");
    EXPECT_DOUBLE_EQ(r.values[north][0][0].as_number(), north_ss / 3.0);
    EXPECT_DOUBLE_EQ(r.values[south][0][0].as_number(), south_ss / 2.0);
  }
  {
    PivotResult r = run(Aggregation::StdDev);
    const std::size_t north = row_index(r, "North");
    const std::size_t south = row_index(r, "South");
    EXPECT_DOUBLE_EQ(r.values[north][0][0].as_number(), std::sqrt(north_ss / 2.0));
    EXPECT_DOUBLE_EQ(r.values[south][0][0].as_number(), std::sqrt(south_ss / 1.0));
  }
  {
    PivotResult r = run(Aggregation::StdDevP);
    const std::size_t north = row_index(r, "North");
    const std::size_t south = row_index(r, "South");
    EXPECT_DOUBLE_EQ(r.values[north][0][0].as_number(), std::sqrt(north_ss / 3.0));
    EXPECT_DOUBLE_EQ(r.values[south][0][0].as_number(), std::sqrt(south_ss / 2.0));
  }
}

TEST(PivotEvaluator, VarAndStdDevSampleNeedTwoValuesElseDiv0) {
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"G", {}});
  cache.mutable_fields().push_back(PivotCacheField{"V", {}});

  auto add = [&](const char* g, double v) {
    PivotCacheRecord rec;
    rec.cells.push_back(owned_text(cache, g));
    rec.cells.push_back(Value::number(v));
    cache.mutable_records().push_back(std::move(rec));
  };
  add("solo", 42.0);  // single record, sample stats undefined.
  add("pair", 10.0);
  add("pair", 20.0);

  auto run = [&](Aggregation agg) {
    PivotTable table;
    table.set_pivot_cache_id(1);
    PivotField gf;
    gf.source_name = "G";
    gf.axis = PivotAxis::Row;
    PivotField vf;
    vf.source_name = "V";
    vf.axis = PivotAxis::Value;
    table.mutable_fields().push_back(std::move(gf));
    table.mutable_fields().push_back(std::move(vf));
    table.mutable_row_field_order() = {0};
    PivotDataField df;
    df.name = "f";
    df.field_index = 1;
    df.aggregation = agg;
    table.mutable_data_fields().push_back(std::move(df));
    table.set_grand_totals(/*rows=*/false, /*cols=*/false);
    auto r_or = evaluate(table, cache);
    EXPECT_TRUE(static_cast<bool>(r_or));
    return r_or.value();
  };

  for (Aggregation agg : {Aggregation::Var, Aggregation::StdDev}) {
    PivotResult r = run(agg);
    const std::size_t solo = row_index(r, "solo");
    const std::size_t pair = row_index(r, "pair");
    ASSERT_NE(solo, static_cast<std::size_t>(-1));
    ASSERT_NE(pair, static_cast<std::size_t>(-1));
    ASSERT_TRUE(r.values[solo][0][0].is_error());
    EXPECT_EQ(r.values[solo][0][0].as_error(), ErrorCode::Div0);
    EXPECT_TRUE(r.values[pair][0][0].is_number());
  }
  for (Aggregation agg : {Aggregation::VarP, Aggregation::StdDevP}) {
    PivotResult r = run(agg);
    const std::size_t solo = row_index(r, "solo");
    const std::size_t pair = row_index(r, "pair");
    // Population variant accepts n=1 and yields 0.
    EXPECT_TRUE(r.values[solo][0][0].is_number());
    EXPECT_DOUBLE_EQ(r.values[solo][0][0].as_number(), 0.0);
    EXPECT_TRUE(r.values[pair][0][0].is_number());
  }
}

// ---------------------------------------------------------------------------
// 8c. Date grouping (Gregorian + Japanese)
// ---------------------------------------------------------------------------
//
// Cache schema: Date (number, Excel serial) + Amount (number).
// Records cover three calendar years and span the Heisei -> Reiwa boundary
// so the Japanese-calendar bucket boundaries are exercised.
PivotCache build_date_cache() {
  // Excel serials (1900 leap-bug aware):
  //   2018-12-31 -> 43465  (Heisei)
  //   2019-04-30 -> 43585  (Heisei, last day of era)
  //   2019-05-01 -> 43586  (Reiwa, first day of era)
  //   2019-12-31 -> 43830  (Reiwa)
  //   2024-03-15 -> 45366  (Reiwa)
  //   2024-07-04 -> 45477  (Reiwa)
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Date", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});
  auto add = [&](double serial, double amount) {
    PivotCacheRecord rec;
    rec.cells.push_back(Value::number(serial));
    rec.cells.push_back(Value::number(amount));
    cache.mutable_records().push_back(std::move(rec));
  };
  add(43465.0, 10.0);
  add(43585.0, 20.0);
  add(43586.0, 40.0);
  add(43830.0, 80.0);
  add(45366.0, 160.0);
  add(45477.0, 320.0);
  return cache;
}

PivotTable build_date_grouped_table(DateGrouping g, CalendarSystem cal) {
  PivotTable table;
  table.set_pivot_cache_id(1);
  PivotField date_f;
  date_f.source_name = "Date";
  date_f.axis = PivotAxis::Row;
  PivotDateGroup dg;
  dg.granularity = g;
  dg.calendar = cal;
  date_f.date_group = dg;
  PivotField amount_f;
  amount_f.source_name = "Amount";
  amount_f.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(date_f));
  table.mutable_fields().push_back(std::move(amount_f));
  table.mutable_row_field_order() = {0};
  PivotDataField sum;
  sum.name = "Sum of Amount";
  sum.field_index = 1;
  sum.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum));
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  return table;
}

TEST(PivotEvaluator, DateGroupingByYearGregorian) {
  PivotCache cache = build_date_cache();
  PivotTable table = build_date_grouped_table(DateGrouping::Year, CalendarSystem::Gregorian);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Three Gregorian years: 2018, 2019, 2024 (chronological).
  ASSERT_EQ(r.rows.size(), 3U);
  EXPECT_EQ(r.rows[0].label, "2018");
  EXPECT_EQ(r.rows[1].label, "2019");
  EXPECT_EQ(r.rows[2].label, "2024");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 10.0);
  EXPECT_DOUBLE_EQ(r.values[1][0][0].as_number(), 20.0 + 40.0 + 80.0);
  EXPECT_DOUBLE_EQ(r.values[2][0][0].as_number(), 160.0 + 320.0);
}

TEST(PivotEvaluator, DateGroupingByQuarterGregorian) {
  PivotCache cache = build_date_cache();
  PivotTable table = build_date_grouped_table(DateGrouping::Quarter, CalendarSystem::Gregorian);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // 2018-Q4 (Dec), 2019-Q2 (Apr+May), 2019-Q4 (Dec), 2024-Q1 (Mar), 2024-Q3 (Jul).
  ASSERT_EQ(r.rows.size(), 5U);
  EXPECT_EQ(r.rows[0].label, "2018-Q4");
  EXPECT_EQ(r.rows[1].label, "2019-Q2");
  EXPECT_EQ(r.rows[2].label, "2019-Q4");
  EXPECT_EQ(r.rows[3].label, "2024-Q1");
  EXPECT_EQ(r.rows[4].label, "2024-Q3");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 10.0);
  EXPECT_DOUBLE_EQ(r.values[1][0][0].as_number(), 20.0 + 40.0);
  EXPECT_DOUBLE_EQ(r.values[2][0][0].as_number(), 80.0);
  EXPECT_DOUBLE_EQ(r.values[3][0][0].as_number(), 160.0);
  EXPECT_DOUBLE_EQ(r.values[4][0][0].as_number(), 320.0);
}

TEST(PivotEvaluator, DateGroupingByMonthGregorian) {
  PivotCache cache = build_date_cache();
  PivotTable table = build_date_grouped_table(DateGrouping::Month, CalendarSystem::Gregorian);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Six distinct months across the dataset.
  ASSERT_EQ(r.rows.size(), 6U);
  EXPECT_EQ(r.rows[0].label, "2018-12");
  EXPECT_EQ(r.rows[1].label, "2019-04");
  EXPECT_EQ(r.rows[2].label, "2019-05");
  EXPECT_EQ(r.rows[3].label, "2019-12");
  EXPECT_EQ(r.rows[4].label, "2024-03");
  EXPECT_EQ(r.rows[5].label, "2024-07");
}

TEST(PivotEvaluator, DateGroupingByDay) {
  PivotCache cache = build_date_cache();
  PivotTable table = build_date_grouped_table(DateGrouping::Day, CalendarSystem::Gregorian);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Each record is a distinct day -> 6 leaves in chronological order.
  ASSERT_EQ(r.rows.size(), 6U);
  EXPECT_EQ(r.rows[0].label, "2018-12-31");
  EXPECT_EQ(r.rows[1].label, "2019-04-30");
  EXPECT_EQ(r.rows[2].label, "2019-05-01");
  EXPECT_EQ(r.rows[3].label, "2019-12-31");
  EXPECT_EQ(r.rows[4].label, "2024-03-15");
  EXPECT_EQ(r.rows[5].label, "2024-07-04");
}

TEST(PivotEvaluator, DateGroupingByYearJapaneseCalendar) {
  PivotCache cache = build_date_cache();
  PivotTable table = build_date_grouped_table(DateGrouping::Year, CalendarSystem::Japanese);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // 2018 -> Heisei 30, 2019-04-30 -> Heisei 31, 2019-05-01 -> Reiwa 1,
  // 2019-12-31 -> Reiwa 1, 2024 -> Reiwa 6.
  // Buckets in chronological order: 平成30 / 平成31 / 令和1 / 令和6.
  ASSERT_EQ(r.rows.size(), 4U);
  // 平成 = E5 B9 B3 E6 88 90, 令和 = E4 BB A4 E5 92 8C, 年 = E5 B9 B4
  EXPECT_EQ(r.rows[0].label, std::string("\xE5\xB9\xB3\xE6\x88\x90") + "30" + "\xE5\xB9\xB4");
  EXPECT_EQ(r.rows[1].label, std::string("\xE5\xB9\xB3\xE6\x88\x90") + "31" + "\xE5\xB9\xB4");
  EXPECT_EQ(r.rows[2].label, std::string("\xE4\xBB\xA4\xE5\x92\x8C") + "1" + "\xE5\xB9\xB4");
  EXPECT_EQ(r.rows[3].label, std::string("\xE4\xBB\xA4\xE5\x92\x8C") + "6" + "\xE5\xB9\xB4");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 10.0);
  EXPECT_DOUBLE_EQ(r.values[1][0][0].as_number(), 20.0);
  EXPECT_DOUBLE_EQ(r.values[2][0][0].as_number(), 40.0 + 80.0);
  EXPECT_DOUBLE_EQ(r.values[3][0][0].as_number(), 160.0 + 320.0);
}

TEST(PivotEvaluator, DateGroupingNonNumericPassesThrough) {
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Date", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});
  // A blank (not a serial) should not crash; bucket_date is only invoked
  // for numeric values.
  PivotCacheRecord rec;
  rec.cells.push_back(Value::blank());
  rec.cells.push_back(Value::number(42.0));
  cache.mutable_records().push_back(std::move(rec));

  PivotTable table = build_date_grouped_table(DateGrouping::Year, CalendarSystem::Gregorian);
  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  // The single blank-keyed bucket survives without exploding; label is the
  // empty string (per `display_string` for Blank).
  ASSERT_EQ(r_or.value().rows.size(), 1U);
}

// ---------------------------------------------------------------------------
// Sub-day / week granularities (Week / Hour / Minute / Second).
// ---------------------------------------------------------------------------
//
// These build their own caches so the records can carry sub-day fractions
// or weekday-precise dates that don't fit `build_date_cache()`'s shape.

PivotCache build_two_field_cache() {
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Date", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});
  return cache;
}

void push_record(PivotCache& cache, double serial, double amount) {
  PivotCacheRecord rec;
  rec.cells.push_back(Value::number(serial));
  rec.cells.push_back(Value::number(amount));
  cache.mutable_records().push_back(std::move(rec));
}

TEST(PivotEvaluator, DateGroupingByWeekGregorian) {
  // Excel serials for ja-JP-friendly Gregorian dates in 2024:
  //   2024-03-11 (Mon) = 45362, 2024-03-13 (Wed) = 45364,
  //   2024-03-15 (Fri) = 45366  -> all in the week starting 2024-03-10 (Sun).
  //   2024-03-18 (Mon) = 45369  -> in the next week, starting 2024-03-17 (Sun).
  PivotCache cache = build_two_field_cache();
  push_record(cache, 45362.0, 1.0);
  push_record(cache, 45364.0, 2.0);
  push_record(cache, 45366.0, 4.0);
  push_record(cache, 45369.0, 8.0);

  PivotTable table = build_date_grouped_table(DateGrouping::Week, CalendarSystem::Gregorian);
  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.rows.size(), 2U);
  EXPECT_EQ(r.rows[0].label, "2024-03-10");
  EXPECT_EQ(r.rows[1].label, "2024-03-17");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 1.0 + 2.0 + 4.0);
  EXPECT_DOUBLE_EQ(r.values[1][0][0].as_number(), 8.0);
}

TEST(PivotEvaluator, DateGroupingByWeekJapaneseUsesGregorianLabel) {
  // The Japanese-calendar selector is ignored for Week buckets; Mac Excel
  // ja-JP renders weekly labels as Gregorian YYYY-MM-DD.
  PivotCache cache = build_two_field_cache();
  push_record(cache, 45362.0, 1.0);  // 2024-03-11 Mon
  push_record(cache, 45369.0, 2.0);  // 2024-03-18 Mon

  PivotTable table = build_date_grouped_table(DateGrouping::Week, CalendarSystem::Japanese);
  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.rows.size(), 2U);
  EXPECT_EQ(r.rows[0].label, "2024-03-10");
  EXPECT_EQ(r.rows[1].label, "2024-03-17");
}

TEST(PivotEvaluator, DateGroupingByHour) {
  // Two records on 2024-03-15 within different hours, plus a third in the
  // same hour as the second so the second bucket sums two rows.
  const double base = 45366.0;  // 2024-03-15
  PivotCache cache = build_two_field_cache();
  push_record(cache, base + 9.0 / 24.0 + 30.0 / 1440.0, 1.0);   // 09:30
  push_record(cache, base + 10.0 / 24.0 + 5.0 / 1440.0, 2.0);   // 10:05
  push_record(cache, base + 10.0 / 24.0 + 45.0 / 1440.0, 4.0);  // 10:45

  PivotTable table = build_date_grouped_table(DateGrouping::Hour, CalendarSystem::Gregorian);
  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.rows.size(), 2U);
  EXPECT_EQ(r.rows[0].label, "2024-03-15 09");
  EXPECT_EQ(r.rows[1].label, "2024-03-15 10");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(r.values[1][0][0].as_number(), 2.0 + 4.0);
}

TEST(PivotEvaluator, DateGroupingByMinute) {
  // Two minutes within the same hour; chronological ordering is preserved.
  const double base = 45366.0;  // 2024-03-15
  PivotCache cache = build_two_field_cache();
  push_record(cache, base + 10.0 / 24.0 + 5.0 / 1440.0, 1.0);   // 10:05
  push_record(cache, base + 10.0 / 24.0 + 45.0 / 1440.0, 2.0);  // 10:45
  push_record(cache, base + 10.0 / 24.0 + 5.0 / 1440.0, 4.0);   // 10:05 again

  PivotTable table = build_date_grouped_table(DateGrouping::Minute, CalendarSystem::Gregorian);
  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.rows.size(), 2U);
  EXPECT_EQ(r.rows[0].label, "2024-03-15 10:05");
  EXPECT_EQ(r.rows[1].label, "2024-03-15 10:45");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 1.0 + 4.0);
  EXPECT_DOUBLE_EQ(r.values[1][0][0].as_number(), 2.0);
}

TEST(PivotEvaluator, DateGroupingBySecond) {
  // Two seconds within the same minute.
  const double base = 45366.0;  // 2024-03-15
  PivotCache cache = build_two_field_cache();
  push_record(cache, base + 10.0 / 24.0 + 5.0 / 1440.0 + 12.0 / 86400.0, 1.0);  // 10:05:12
  push_record(cache, base + 10.0 / 24.0 + 5.0 / 1440.0 + 47.0 / 86400.0, 2.0);  // 10:05:47
  push_record(cache, base + 10.0 / 24.0 + 5.0 / 1440.0 + 12.0 / 86400.0, 4.0);  // 10:05:12 again

  PivotTable table = build_date_grouped_table(DateGrouping::Second, CalendarSystem::Gregorian);
  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.rows.size(), 2U);
  EXPECT_EQ(r.rows[0].label, "2024-03-15 10:05:12");
  EXPECT_EQ(r.rows[1].label, "2024-03-15 10:05:47");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 1.0 + 4.0);
  EXPECT_DOUBLE_EQ(r.values[1][0][0].as_number(), 2.0);
}

// ---------------------------------------------------------------------------
// 8d. PivotFilter (LabelContains / LabelBeginsWith / ValueTop10 / ValueGreaterThan)
// ---------------------------------------------------------------------------

TEST(PivotEvaluator, LabelContainsFilterDropsRecords) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  // Drop records whose Region label contains "outh" (i.e. South).
  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Region";
  f.type = FilterType::LabelContains;
  f.value = std::string("outh");
  // PivotFilter pre-aggregation reject: every match drops the record.
  // Wait — LabelContains with payload "outh" *passes* records whose label
  // contains it. To drop South we instead use LabelBeginsWith("North")
  // below; LabelContains here keeps only South.
  table.mutable_active_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, "South");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 500.0);  // 200 + 300
}

TEST(PivotEvaluator, LabelBeginsWithFilterDropsRecords) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Region";
  f.type = FilterType::LabelBeginsWith;
  f.value = std::string("Nor");
  table.mutable_active_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, "North");
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

TEST(PivotEvaluator, ShowAsRunningTotalInRowKeepsRawSubtotals) {
  PivotCache cache = build_show_as_cache();
  PivotTable table = build_show_as_table(ShowValuesAs::RunningTotalInRow, /*grand_totals=*/true);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // North row leaf sums: Widget=30 (Online=10, Store=20),
  // Gadget=70 (Online=30, Store=40); the running-total transform
  // rewrites each row so the LAST col cell equals the row's raw sum.
  // With 2 col leaves, last col = full row sum.
  ASSERT_EQ(r.values.size(), 4U);
  for (const auto& row_slot : r.values) {
    ASSERT_EQ(row_slot.size(), 2U);
    ASSERT_TRUE(row_slot[0][0].is_number());
    ASSERT_TRUE(row_slot[1][0].is_number());
    // Strictly non-decreasing along the row (running sum of
    // non-negative cells).
    EXPECT_GE(row_slot[1][0].as_number(), row_slot[0][0].as_number());
  }
  // Row subtotals must remain at raw aggregate (North=100, South=260).
  ASSERT_EQ(r.row_subtotals.size(), 2U);
  std::vector<double> raw_subtotals;
  for (const RowSubtotal& sub : r.row_subtotals) {
    ASSERT_FALSE(sub.values.empty());
    ASSERT_TRUE(sub.values[0].is_number());
    raw_subtotals.push_back(sub.values[0].as_number());
  }
  std::sort(raw_subtotals.begin(), raw_subtotals.end());
  EXPECT_NEAR(raw_subtotals[0], 100.0, 1e-9);
  EXPECT_NEAR(raw_subtotals[1], 260.0, 1e-9);

  // Grand total stays at raw 360.
  ASSERT_EQ(r.grand_totals.size(), 1U);
  ASSERT_TRUE(r.grand_totals[0].is_number());
  EXPECT_NEAR(r.grand_totals[0].as_number(), 360.0, 1e-9);
}

// ---------------------------------------------------------------------------
// 9. Manual filter (PivotItem::visible == false) hides records
// ---------------------------------------------------------------------------

TEST(PivotEvaluator, ManualFilterHidesItem) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  // Hide all records with Region == "South" via the Region field's items.
  table.mutable_fields()[0].items = {
      PivotItem{"North", /*visible=*/true},
      PivotItem{"South", /*visible=*/false},
  };

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Only North survives.
  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, "North");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 175.0);
}

// ---------------------------------------------------------------------------
// 10. Grand total when flagged
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// 11. Grand total when not flagged stays Blank
// ---------------------------------------------------------------------------

TEST(PivotEvaluator, GrandTotalBlankWhenDisabled) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  EXPECT_TRUE(r_or.value().grand_total.is_blank());
}

// ---------------------------------------------------------------------------
// 12. Cache id mismatch
// ---------------------------------------------------------------------------

TEST(PivotEvaluator, CacheIdMismatchYieldsMissing) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_pivot_cache_id(99);  // Cache has id 1.

  auto r_or = evaluate(table, cache);
  ASSERT_FALSE(static_cast<bool>(r_or));
  EXPECT_EQ(r_or.error().code, FormulonErrorCode::kEvalPivotMissing);
}

// ---------------------------------------------------------------------------
// 13. Out-of-bounds field_index
// ---------------------------------------------------------------------------

TEST(PivotEvaluator, OutOfRangeFieldIndexYieldsInvalid) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.mutable_data_fields()[0].field_index = 99;

  auto r_or = evaluate(table, cache);
  ASSERT_FALSE(static_cast<bool>(r_or));
  EXPECT_EQ(r_or.error().code, FormulonErrorCode::kEvalPivotInvalid);
}

// ---------------------------------------------------------------------------
// 14. Error in source value propagates through SUM
// ---------------------------------------------------------------------------

TEST(PivotEvaluator, ErrorPropagatesThroughSum) {
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Group", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});

  auto add = [&](const char* g, Value v) {
    PivotCacheRecord rec;
    rec.cells.push_back(owned_text(cache, g));
    rec.cells.push_back(v);
    cache.mutable_records().push_back(std::move(rec));
  };
  add("A", Value::number(10.0));
  add("A", Value::error(ErrorCode::Div0));
  add("A", Value::number(20.0));

  PivotTable table;
  table.set_pivot_cache_id(1);
  PivotField gf;
  gf.source_name = "Group";
  gf.axis = PivotAxis::Row;
  PivotField af;
  af.source_name = "Amount";
  af.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(gf));
  table.mutable_fields().push_back(std::move(af));
  table.mutable_row_field_order() = {0};
  PivotDataField sum;
  sum.name = "Sum of Amount";
  sum.field_index = 1;
  sum.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum));
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  ASSERT_EQ(r.rows.size(), 1U);
  ASSERT_TRUE(r.values[0][0][0].is_error());
  EXPECT_EQ(r.values[0][0][0].as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// Bonus: subtotal at Region level when subtotal_top is set
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// LabelDate filter (date-range, pre-aggregation)
// ---------------------------------------------------------------------------

// Helper: builds a 2-column cache (Date as numeric serial, Amount). The
// records are picked so that a closed [2024-01-01, 2024-12-31] window
// keeps the first three and drops the fourth.
PivotCache build_label_date_filter_cache() {
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Date", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});
  auto add = [&](double serial, double amount) {
    PivotCacheRecord rec;
    rec.cells.push_back(Value::number(serial));
    rec.cells.push_back(Value::number(amount));
    cache.mutable_records().push_back(std::move(rec));
  };
  add(formulon::eval::date_time::serial_from_ymd(2024, 1, 1), 10.0);
  add(formulon::eval::date_time::serial_from_ymd(2024, 6, 15), 20.0);
  add(formulon::eval::date_time::serial_from_ymd(2024, 12, 31), 30.0);
  add(formulon::eval::date_time::serial_from_ymd(2025, 3, 1), 40.0);
  return cache;
}

// Helper: builds a single-row-axis pivot over `Date` with SUM(Amount).
// The Date field becomes the row axis so each surviving record forms its
// own row leaf, simplifying assertions.
PivotTable build_label_date_filter_table() {
  PivotTable table;
  table.set_pivot_cache_id(1);
  PivotField date_f;
  date_f.source_name = "Date";
  date_f.axis = PivotAxis::Row;
  PivotField amount_f;
  amount_f.source_name = "Amount";
  amount_f.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(date_f));
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

TEST(PivotEvaluator, LabelDateFilterIncludesInRange) {
  PivotCache cache = build_label_date_filter_cache();
  PivotTable table = build_label_date_filter_table();

  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Date";
  f.type = FilterType::LabelDate;
  f.value = formulon::eval::date_time::serial_from_ymd(2024, 1, 1);
  f.value_high = formulon::eval::date_time::serial_from_ymd(2024, 12, 31);
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Three of the four records fall inside [2024-01-01, 2024-12-31].
  ASSERT_EQ(r.rows.size(), 3U);
  std::vector<double> totals;
  totals.reserve(r.values.size());
  for (const auto& row_slot : r.values) {
    ASSERT_EQ(row_slot.size(), 1U);
    ASSERT_EQ(row_slot[0].size(), 1U);
    totals.push_back(row_slot[0][0].as_number());
  }
  std::sort(totals.begin(), totals.end());
  EXPECT_DOUBLE_EQ(totals[0], 10.0);
  EXPECT_DOUBLE_EQ(totals[1], 20.0);
  EXPECT_DOUBLE_EQ(totals[2], 30.0);
}

TEST(PivotEvaluator, LabelDateFilterUnboundedHighIsNoOp) {
  PivotCache cache = build_label_date_filter_cache();
  PivotTable table = build_label_date_filter_table();

  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Date";
  f.type = FilterType::LabelDate;
  f.value = formulon::eval::date_time::serial_from_ymd(2024, 1, 1);
  // value_high left as default monostate -> filter degrades to no-op.
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // No upper bound -> all four records survive.
  ASSERT_EQ(r.rows.size(), 4U);
}

// ---------------------------------------------------------------------------
// ValueBetween filter (post-aggregation, parallel to Top-N / GreaterThan)
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// Multi-level value filters (Top-N / GreaterThan / Between) over hierarchies
// ---------------------------------------------------------------------------

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
  f.field_name = "Region";
  f.type = FilterType::ValueTop10;
  f.value = 3;  // Top 3 leaves.
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Top-3 by Amount: South/A=80, North/B=50, South/B=30. The "C" leaves
  // and North/A are pruned. After pruning North loses A and C (only B
  // remains); South loses C (A and B remain).
  ASSERT_EQ(r.rows.size(), 2U);

  // Find Region nodes; ordering is alphabetical (North < South).
  const RowHierarchyNode* north = nullptr;
  const RowHierarchyNode* south = nullptr;
  for (const auto& node : r.rows) {
    if (node.label == "North") {
      north = &node;
    } else if (node.label == "South") {
      south = &node;
    }
  }
  ASSERT_NE(north, nullptr);
  ASSERT_NE(south, nullptr);

  // North keeps only B.
  ASSERT_EQ(north->children.size(), 1U);
  EXPECT_EQ(north->children[0].label, "B");
  // South keeps A and B.
  ASSERT_EQ(south->children.size(), 2U);
  EXPECT_EQ(south->children[0].label, "A");
  EXPECT_EQ(south->children[1].label, "B");

  // Three surviving leaves -> three rows in `values`.
  ASSERT_EQ(r.values.size(), 3U);
  std::vector<double> totals;
  totals.reserve(3);
  for (const auto& row_slot : r.values) {
    ASSERT_EQ(row_slot.size(), 1U);
    ASSERT_EQ(row_slot[0].size(), 1U);
    totals.push_back(row_slot[0][0].as_number());
  }
  std::sort(totals.begin(), totals.end());
  EXPECT_DOUBLE_EQ(totals[0], 30.0);
  EXPECT_DOUBLE_EQ(totals[1], 50.0);
  EXPECT_DOUBLE_EQ(totals[2], 80.0);
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
  f.field_name = "Year";
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
  f.field_name = "Region";
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
  ASSERT_TRUE(r.values[0][0][0].is_number());
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 0.0);  // A - A = 0.
  ASSERT_TRUE(r.values[1][0][0].is_number());
  EXPECT_DOUBLE_EQ(r.values[1][0][0].as_number(), 15.0);  // B - A = 25 - 10.
  ASSERT_TRUE(r.values[2][0][0].is_number());
  EXPECT_DOUBLE_EQ(r.values[2][0][0].as_number(), 30.0);  // C - A = 40 - 10.
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
// Regression: subtotals are gated by default_subtotal, not subtotal_top.
// A multi-level row hierarchy emits outer-field subtotals by default even
// when subtotal_top was never set (subtotal_top is only the position flag).
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// H-23: a custom subtotal function replaces the default aggregation for the
// group's subtotal row.
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// H-24: per-leaf row/col totals re-aggregate the underlying records so a
// non-additive function (Average) is not computed as an average of the
// per-cell averages.
// ---------------------------------------------------------------------------

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
// M-19: runTotal accumulation direction follows the data field's baseField,
// not the enum spelling. A base field on the column axis accumulates down
// columns even when the mode is RunningTotalInRow.
// ---------------------------------------------------------------------------

TEST(PivotEvaluator, RunTotalDirectionFollowsBaseFieldAxis) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{1});
  table.mutable_data_fields()[0].show_as = ShowValuesAs::RunningTotalInRow;
  table.mutable_data_fields()[0].show_as_base_field = 1;  // Product = a column field.
  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  // Row leaves: North(0), South(1). Col leaves: Gadget(0), Widget(1).
  // Column-direction running total down the Gadget column: North=50,
  // South=50+300=350. Row-direction would have left South/Gadget at 300.
  const std::size_t south = row_index(r, "South");
  ASSERT_LT(south, r.values.size());
  ASSERT_GE(r.values[south].size(), 1U);
  ASSERT_FALSE(r.values[south][0].empty());
  ASSERT_TRUE(r.values[south][0][0].is_number());
  EXPECT_NEAR(r.values[south][0][0].as_number(), 350.0, 1e-9);
}

// ---------------------------------------------------------------------------
// M-22: the pivot comparator folds Japanese text (half-width katakana to
// full-width) exactly like the GROUPBY / SORT comparator, so the two cannot
// diverge on kana collation.
// ---------------------------------------------------------------------------

TEST(PivotComparatorParity, FoldsHalfWidthKatakana) {
  // U+FF76 (halfwidth ｶ) folds to U+30AB (fullwidth カ); after folding the
  // two are equal, so neither orders before the other.
  const Value full = Value::text("\xE3\x82\xAB");  // カ
  const Value half = Value::text("\xEF\xBD\xB6");  // ｶ
  EXPECT_FALSE(value_less(full, half));
  EXPECT_FALSE(value_less(half, full));
  // GROUPBY / SORT agrees they are equal.
  EXPECT_EQ(eval::cmp_value_asc(full, half), 0);
}

}  // namespace
}  // namespace formulon::pivot
