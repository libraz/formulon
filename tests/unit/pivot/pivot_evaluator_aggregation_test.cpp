// Aggregate functions and error handling for `formulon::pivot::evaluate`.

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <limits>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "pivot/aggregator.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_evaluator.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot/record_access.h"
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
// 14b. Count / CountNumbers classify cells, so a source error never becomes
//      the count -- at the leaf, at the subtotal, or at the grand total
// ---------------------------------------------------------------------------

// Builds a Group x Sub hierarchy over a value column that carries one
// error, so a single `#N/A` reaches the leaf cell, its group's subtotal and
// the grand total through the same aggregation.
PivotTable build_count_over_error_table(PivotCache& cache, Aggregation aggregation) {
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Group", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Sub", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});

  auto add = [&](const char* group, const char* sub, Value v) {
    PivotCacheRecord rec;
    rec.cells.push_back(owned_text(cache, group));
    rec.cells.push_back(owned_text(cache, sub));
    rec.cells.push_back(v);
    cache.mutable_records().push_back(std::move(rec));
  };
  add("A", "x", Value::number(10.0));
  add("A", "x", Value::error(ErrorCode::NA));  // e.g. a lookup that missed
  add("A", "y", Value::number(20.0));
  add("B", "x", Value::number(30.0));

  PivotTable table;
  table.set_pivot_cache_id(1);
  PivotField group_f;
  group_f.source_name = "Group";
  group_f.axis = PivotAxis::Row;
  group_f.subtotal_top = true;
  PivotField sub_f;
  sub_f.source_name = "Sub";
  sub_f.axis = PivotAxis::Row;
  PivotField amount_f;
  amount_f.source_name = "Amount";
  amount_f.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(group_f));
  table.mutable_fields().push_back(std::move(sub_f));
  table.mutable_fields().push_back(std::move(amount_f));
  table.mutable_row_field_order() = {0, 1};
  PivotDataField df;
  df.name = "Count of Amount";
  df.field_index = 2;
  df.aggregation = aggregation;
  table.mutable_data_fields().push_back(std::move(df));
  table.set_grand_totals(/*rows=*/true, /*cols=*/true);
  return table;
}

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

TEST(PivotEvaluator, ArithmeticAggregatesMapOverflowToNumError) {
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
  const double huge = std::numeric_limits<double>::max();
  add("North", huge);
  add("North", huge);  // huge + huge overflows a double to +inf.

  auto run = [&](Aggregation agg) {
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
    PivotDataField df;
    df.name = "Agg";
    df.field_index = 1;
    df.aggregation = agg;
    table.mutable_data_fields().push_back(std::move(df));
    table.set_grand_totals(/*rows=*/false, /*cols=*/false);
    auto r_or = evaluate(table, cache);
    EXPECT_TRUE(static_cast<bool>(r_or));
    return r_or.value();
  };

  for (Aggregation agg : {Aggregation::Sum, Aggregation::Average, Aggregation::Product}) {
    PivotResult r = run(agg);
    ASSERT_TRUE(r.values[0][0][0].is_error()) << static_cast<int>(agg);
    EXPECT_EQ(r.values[0][0][0].as_error(), ErrorCode::Num) << static_cast<int>(agg);
  }
}

TEST(PivotEvaluator, AverageKeepsPlainSumRoundingAndOverflow) {
  const Value last_bit =
      apply_aggregation(Aggregation::Average, {Value::number(0.1), Value::number(0.2), Value::number(0.3)});
  ASSERT_TRUE(last_bit.is_number()) << last_bit.debug_to_string();
  EXPECT_DOUBLE_EQ(last_bit.as_number(), (0.1 + 0.2 + 0.3) / 3.0);

  const Value absorbed =
      apply_aggregation(Aggregation::Average, {Value::number(1.0e16), Value::number(1.0), Value::number(-1.0e16)});
  ASSERT_TRUE(absorbed.is_number()) << absorbed.debug_to_string();
  EXPECT_DOUBLE_EQ(absorbed.as_number(), 0.0);

  const Value overflow = apply_aggregation(Aggregation::Average, {Value::number(1.0e308), Value::number(1.0e308)});
  ASSERT_TRUE(overflow.is_error()) << overflow.debug_to_string();
  EXPECT_EQ(overflow.as_error(), ErrorCode::Num);
}

TEST(PivotEvaluator, MaxMinMapAStoredNonFiniteValueToNumError) {
  auto run = [&](Aggregation agg, double extreme) {
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
    add("North", extreme);
    add("North", 1.0);

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
    PivotDataField df;
    df.name = "Agg";
    df.field_index = 1;
    df.aggregation = agg;
    table.mutable_data_fields().push_back(std::move(df));
    table.set_grand_totals(/*rows=*/false, /*cols=*/false);
    auto r_or = evaluate(table, cache);
    EXPECT_TRUE(static_cast<bool>(r_or));
    return r_or.value();
  };

  {
    // +Infinity beats every finite value, so it is the one MAX picks.
    PivotResult r = run(Aggregation::Max, std::numeric_limits<double>::infinity());
    ASSERT_TRUE(r.values[0][0][0].is_error());
    EXPECT_EQ(r.values[0][0][0].as_error(), ErrorCode::Num);
  }
  {
    // -Infinity is smaller than every finite value, so it is the one MIN picks.
    PivotResult r = run(Aggregation::Min, -std::numeric_limits<double>::infinity());
    ASSERT_TRUE(r.values[0][0][0].is_error());
    EXPECT_EQ(r.values[0][0][0].as_error(), ErrorCode::Num);
  }
}

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

TEST(PivotEvaluator, ExtremeDispersionOverflowIsNumError) {
  const std::vector<Value> equal = {Value::number(1.0e308), Value::number(1.0e308)};
  for (Aggregation agg : {Aggregation::Var, Aggregation::VarP, Aggregation::StdDev, Aggregation::StdDevP}) {
    const Value result = apply_aggregation(agg, equal);
    ASSERT_TRUE(result.is_error()) << static_cast<int>(agg) << ": " << result.debug_to_string();
    EXPECT_EQ(result.as_error(), ErrorCode::Num) << static_cast<int>(agg);
  }

  const double max = std::numeric_limits<double>::max();
  const Value opposing = apply_aggregation(Aggregation::StdDevP, {Value::number(max), Value::number(-max)});
  ASSERT_TRUE(opposing.is_error()) << opposing.debug_to_string();
  EXPECT_EQ(opposing.as_error(), ErrorCode::Num);
}

TEST(PivotEvaluator, AdjacentLargeValuesMatchExcelPopulationVariance) {
  constexpr double first = 1.0e150;
  const Value variance =
      apply_aggregation(Aggregation::VarP, {Value::number(first), Value::number(std::nextafter(first, 0.0))});
  ASSERT_TRUE(variance.is_number()) << variance.debug_to_string();
  // One-pass (n*S2 - S1*S1) / n^2, as Excel's pivot computes it, not the two-pass worksheet VARP.
  EXPECT_DOUBLE_EQ(variance.as_number(), 1.4870169084777831e+284);
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

TEST(PivotEvaluator, CountCountsAnErrorCellInsteadOfReturningIt) {
  PivotCache cache;
  PivotTable table = build_count_over_error_table(cache, Aggregation::Count);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Row leaves are A/x, A/y, B/x in hierarchy order; A/x is the one
  // holding a number and the error, and COUNTA counts both.
  ASSERT_EQ(r.rows.size(), 2U);
  EXPECT_EQ(r.rows[0].label, "A");
  ASSERT_EQ(r.values.size(), 3U);
  ASSERT_TRUE(r.values[0][0][0].is_number()) << "Count returned " << r.values[0][0][0].debug_to_string();
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 2.0);

  // The group subtotal and the grand total run the same aggregation over a
  // wider record set, so they must not turn into the error either.
  ASSERT_EQ(r.row_subtotals.size(), 2U);
  ASSERT_TRUE(r.row_subtotals[0].values[0].is_number());
  EXPECT_DOUBLE_EQ(r.row_subtotals[0].values[0].as_number(), 3.0);
  ASSERT_FALSE(r.grand_totals.empty());
  ASSERT_TRUE(r.grand_totals[0].is_number());
  EXPECT_DOUBLE_EQ(r.grand_totals[0].as_number(), 4.0);
}

TEST(PivotEvaluator, CountNumbersPassesOverAnErrorCellInsteadOfReturningIt) {
  PivotCache cache;
  PivotTable table = build_count_over_error_table(cache, Aggregation::CountNumbers);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // COUNT admits only numeric cells, so the error is passed over rather
  // than counted -- and still never becomes the result.
  ASSERT_EQ(r.values.size(), 3U);
  ASSERT_TRUE(r.values[0][0][0].is_number()) << "CountNumbers returned " << r.values[0][0][0].debug_to_string();
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 1.0);

  ASSERT_EQ(r.row_subtotals.size(), 2U);
  ASSERT_TRUE(r.row_subtotals[0].values[0].is_number());
  EXPECT_DOUBLE_EQ(r.row_subtotals[0].values[0].as_number(), 2.0);
  ASSERT_FALSE(r.grand_totals.empty());
  ASSERT_TRUE(r.grand_totals[0].is_number());
  EXPECT_DOUBLE_EQ(r.grand_totals[0].as_number(), 3.0);
}

TEST(PivotEvaluator, SumOverTheSameFixtureStillPropagatesTheError) {
  PivotCache cache;
  PivotTable table = build_count_over_error_table(cache, Aggregation::Sum);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  ASSERT_EQ(r.values.size(), 3U);
  ASSERT_TRUE(r.values[0][0][0].is_error());
  EXPECT_EQ(r.values[0][0][0].as_error(), ErrorCode::NA);
}

void expect_bitwise_value_equal(const Value& expected, const Value& actual) {
  ASSERT_EQ(expected.kind(), actual.kind());
  if (expected.is_number()) {
    std::uint64_t expected_bits = 0;
    std::uint64_t actual_bits = 0;
    const double expected_number = expected.as_number();
    const double actual_number = actual.as_number();
    std::memcpy(&expected_bits, &expected_number, sizeof(expected_bits));
    std::memcpy(&actual_bits, &actual_number, sizeof(actual_bits));
    EXPECT_EQ(expected_bits, actual_bits);
    return;
  }
  EXPECT_EQ(expected, actual);
}

TEST(PivotAggregator, BorrowedRecordViewsMatchMaterializedValuesForEveryAggregation) {
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(
      PivotCacheField{"Value", {Value::number(1.0e16), Value::number(1.0), Value::number(-1.0e16)}});
  auto add = [&](Value value, bool is_index) {
    PivotCacheRecord record;
    record.cells.push_back(value);
    record.cell_is_index.push_back(is_index);
    cache.mutable_records().push_back(std::move(record));
  };
  add(Value::number(0.0), true);   // 1e16 from shared items.
  add(Value::number(1.0), false);  // Inline 1, despite the shared field.
  add(Value::number(2.0), true);   // -1e16 from shared items.
  add(Value::boolean(true), false);
  add(Value::text("text"), false);
  add(Value::blank(), false);

  const std::vector<std::size_t> direct_indices{0, 1, 2, 3, 4, 5};
  const std::vector<Value> materialized = {
      cell_value(cache, cache.records()[0], 0), cell_value(cache, cache.records()[1], 0),
      cell_value(cache, cache.records()[2], 0), cell_value(cache, cache.records()[3], 0),
      cell_value(cache, cache.records()[4], 0), cell_value(cache, cache.records()[5], 0)};
  RecordBuckets buckets(2, std::vector<std::vector<std::size_t>>(2));
  buckets[0][0] = {0, 1};
  buckets[0][1] = {2, 3};
  buckets[1][0] = {4, 5};
  const std::vector<std::size_t> rows{0, 1};
  const std::vector<std::size_t> cols{0, 1};
  std::vector<RecordIndexSpan> spans;
  const RecordIndexView multi = make_leaf_set_record_view(buckets, rows, cols, spans);

  const Aggregation aggregations[] = {Aggregation::Sum,          Aggregation::Count,  Aggregation::Average,
                                      Aggregation::Max,          Aggregation::Min,    Aggregation::Product,
                                      Aggregation::CountNumbers, Aggregation::StdDev, Aggregation::StdDevP,
                                      Aggregation::Var,          Aggregation::VarP};
  for (const Aggregation aggregation : aggregations) {
    const Value expected = apply_aggregation(aggregation, materialized);
    const Value direct = apply_record_aggregation(aggregation, cache, direct_indices, 0);
    const Value borrowed = apply_record_aggregation(aggregation, cache, multi, 0);
    expect_bitwise_value_equal(expected, direct);
    expect_bitwise_value_equal(expected, borrowed);
  }

  // Pinned bucket-order results: 1e16 + 1 rounds back to 1e16 before -1e16 cancels it; the boolean is non-numeric.
  expect_bitwise_value_equal(apply_record_aggregation(Aggregation::Sum, cache, multi, 0), Value::number(0.0));
  expect_bitwise_value_equal(apply_record_aggregation(Aggregation::Average, cache, multi, 0), Value::number(0.0));
  expect_bitwise_value_equal(apply_record_aggregation(Aggregation::Count, cache, multi, 0), Value::number(5.0));
  expect_bitwise_value_equal(apply_record_aggregation(Aggregation::CountNumbers, cache, multi, 0), Value::number(3.0));
  expect_bitwise_value_equal(apply_record_aggregation(Aggregation::Max, cache, multi, 0), Value::number(1.0e16));
  expect_bitwise_value_equal(apply_record_aggregation(Aggregation::Min, cache, multi, 0), Value::number(-1.0e16));
}

TEST(PivotAggregator, BorrowedRecordViewsPreserveFirstErrorsAndSparseBlankSemantics) {
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Value", {Value::number(7.0)}});
  auto add = [&](Value value, bool is_index) {
    PivotCacheRecord record;
    record.cells.push_back(value);
    record.cell_is_index.push_back(is_index);
    cache.mutable_records().push_back(std::move(record));
  };
  add(Value::number(0.0), true);  // shared numeric
  add(Value::boolean(false), false);
  add(Value::text("inline"), false);
  add(Value::blank(), false);
  add(Value::error(ErrorCode::NA), false);

  const std::vector<std::size_t> direct_indices{0, 1, 2, 3, 4};
  RecordBuckets buckets(2, std::vector<std::vector<std::size_t>>(1));
  buckets[0][0] = {0, 1, 2};
  buckets[1][0] = {3, 4};
  const std::vector<std::size_t> rows{0, 1};
  const std::vector<std::size_t> cols{0};
  std::vector<RecordIndexSpan> spans;
  const RecordIndexView multi = make_leaf_set_record_view(buckets, rows, cols, spans);
  const std::vector<Value> materialized = {
      cell_value(cache, cache.records()[0], 0), cell_value(cache, cache.records()[1], 0),
      cell_value(cache, cache.records()[2], 0), cell_value(cache, cache.records()[3], 0),
      cell_value(cache, cache.records()[4], 0)};

  for (const Aggregation aggregation :
       {Aggregation::Sum, Aggregation::Average, Aggregation::Var, Aggregation::StdDev, Aggregation::Product}) {
    const Value expected = apply_aggregation(aggregation, materialized);
    expect_bitwise_value_equal(expected, apply_record_aggregation(aggregation, cache, direct_indices, 0));
    expect_bitwise_value_equal(expected, apply_record_aggregation(aggregation, cache, multi, 0));
  }
  for (const Aggregation aggregation : {Aggregation::Count, Aggregation::CountNumbers}) {
    const Value expected = apply_aggregation(aggregation, materialized);
    expect_bitwise_value_equal(expected, apply_record_aggregation(aggregation, cache, direct_indices, 0));
    expect_bitwise_value_equal(expected, apply_record_aggregation(aggregation, cache, multi, 0));
  }

  PivotResult result;
  const std::vector<std::size_t> empty_indices;
  const Value sparse_blank = aggregate_records_or_blank(Aggregation::Sum, cache, empty_indices, 0, result);
  EXPECT_TRUE(sparse_blank.is_blank());
  const std::vector<std::size_t> one_blank{3};
  const Value one_blank_result = aggregate_records_or_blank(Aggregation::Sum, cache, one_blank, 0, result);
  EXPECT_TRUE(one_blank_result.is_blank());
  const std::vector<std::size_t> blank_and_text{3, 2};
  const Value text_result = aggregate_records_or_blank(Aggregation::Sum, cache, blank_and_text, 0, result);
  ASSERT_TRUE(text_result.is_number());
  EXPECT_DOUBLE_EQ(text_result.as_number(), 0.0);
}

}  // namespace
}  // namespace formulon::pivot
