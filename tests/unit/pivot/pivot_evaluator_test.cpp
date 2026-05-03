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
#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "utils/error.h"
#include "value.h"

namespace formulon::pivot {
namespace {

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
  ASSERT_EQ(r.subtotals[0].size(), 1U);
  // Subtotals appear in row-hierarchy DFS order (post-order at each
  // non-leaf), matching how Excel walks the tree to position them.
  std::vector<double> totals{r.subtotals[0][0].as_number(), r.subtotals[1][0].as_number()};
  std::sort(totals.begin(), totals.end());
  EXPECT_DOUBLE_EQ(totals[0], 175.0);  // North subtotal: 100 + 50 + 25
  EXPECT_DOUBLE_EQ(totals[1], 500.0);  // South subtotal: 200 + 300
}

}  // namespace
}  // namespace formulon::pivot
