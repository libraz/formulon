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
#include <limits>
#include <optional>
#include <string>
#include <variant>
#include <vector>

#include "eval/date_time.h"
#include "eval/groupby_pivotby/common.h"
#include "gtest/gtest.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_index.h"
#include "pivot/pivot_layout.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot/record_access.h"
#include "pivot/value_order.h"
#include "utils/checked_index.h"
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

// Same three-column shape as `build_basic_cache()`, but the amounts are
// deliberately non-integral / past the 64-bit integer range so a record's
// label goes through the general numeric rendering path rather than the
// integral one.
//
//   Region  Product  Amount
//   ------  -------  ------
//   North   Widget      1.5
//   North   Gadget      2
//   South   Widget     1e20
//   South   Gadget      4
PivotCache build_non_integral_amount_cache() {
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

  add("North", "Widget", 1.5);
  add("North", "Gadget", 2.0);
  add("South", "Widget", 1e20);
  add("South", "Gadget", 4.0);
  return cache;
}

// Index into `build_shared_region_cache()`'s Region `shared_items` of the
// entry with no value, and of the text entry deliberately spelled like the
// default blank placeholder.
constexpr std::uint32_t kSharedRegionBlank = 1U;
constexpr std::uint32_t kSharedRegionPlaceholderText = 2U;

// Same three-column shape as `build_basic_cache()`, but Region is a *shared*
// field the way a discrete axis column arrives from OOXML: each record stores
// an index into `shared_items`, one of which has no value at all.
//
// The third shared item is a genuine text value spelled exactly like the
// default blank placeholder, so a filter that identified the empty item by its
// rendered label would catch this row too.
//
//   Region     Product  Amount
//   ---------  -------  ------
//   North      Widget    100
//   <no value> Widget     10
//   "(blank)"  Widget      7
PivotCache build_shared_region_cache() {
  PivotCache cache;
  cache.set_cache_id(1);
  PivotCacheField region;
  region.name = "Region";
  region.shared_items.push_back(owned_text(cache, "North"));
  region.shared_items.push_back(Value::blank());
  region.shared_items.push_back(owned_text(cache, std::string{PivotLayoutOptions{}.blank_item_label}));
  cache.mutable_fields().push_back(std::move(region));
  cache.mutable_fields().push_back(PivotCacheField{"Product", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});

  auto add = [&](std::uint32_t region_index, const char* product, double amount) {
    PivotCacheRecord rec;
    rec.cells = {Value::number(region_index), owned_text(cache, product), Value::number(amount)};
    rec.cell_is_index = {true, false, false};
    cache.mutable_records().push_back(std::move(rec));
  };

  add(0U, "Widget", 100.0);
  add(kSharedRegionBlank, "Widget", 10.0);
  add(kSharedRegionPlaceholderText, "Widget", 7.0);
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

TEST(PivotEvaluator, FieldDescendingSortReversesHierarchyAndValues) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{1});
  table.mutable_fields()[0].sort.ascending = false;
  table.mutable_fields()[1].sort.ascending = false;

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.rows.size(), 2U);
  ASSERT_EQ(r.cols.size(), 2U);
  EXPECT_EQ(r.rows[0].label, "South");
  EXPECT_EQ(r.rows[1].label, "North");
  EXPECT_EQ(r.cols[0].label, "Widget");
  EXPECT_EQ(r.cols[1].label, "Gadget");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 200.0);  // South/Widget
  EXPECT_DOUBLE_EQ(r.values[0][1][0].as_number(), 300.0);  // South/Gadget
  EXPECT_DOUBLE_EQ(r.values[1][0][0].as_number(), 125.0);  // North/Widget
  EXPECT_DOUBLE_EQ(r.values[1][1][0].as_number(), 50.0);   // North/Gadget
}

TEST(PivotEvaluator, FieldSortByValueFieldReordersAxisAndValues) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{1});
  table.mutable_fields()[0].sort.by_field = "Amount";
  table.mutable_fields()[0].sort.ascending = false;
  table.mutable_fields()[1].sort.by_field = "Sum of Amount";
  table.mutable_fields()[1].sort.ascending = true;

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.rows.size(), 2U);
  ASSERT_EQ(r.cols.size(), 2U);
  // Row totals: South=500, North=175, so descending by Amount puts South first.
  EXPECT_EQ(r.rows[0].label, "South");
  EXPECT_EQ(r.rows[1].label, "North");
  // Column totals: Widget=325, Gadget=350, so ascending by Sum of Amount puts Widget first.
  EXPECT_EQ(r.cols[0].label, "Widget");
  EXPECT_EQ(r.cols[1].label, "Gadget");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 200.0);  // South/Widget
  EXPECT_DOUBLE_EQ(r.values[0][1][0].as_number(), 300.0);  // South/Gadget
  EXPECT_DOUBLE_EQ(r.values[1][0][0].as_number(), 125.0);  // North/Widget
  EXPECT_DOUBLE_EQ(r.values[1][1][0].as_number(), 50.0);   // North/Gadget
}

TEST(PivotEvaluator, SortedSubtotalWalkPreservesTypedDuplicateDisplayLabels) {
  PivotCache cache;
  cache.set_cache_id(17);
  cache.mutable_fields().push_back(PivotCacheField{"Group", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Kind", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Item", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});

  auto add = [&](const char* group, Value kind, double amount) {
    PivotCacheRecord record;
    record.cells.push_back(owned_text(cache, group));
    record.cells.push_back(std::move(kind));
    record.cells.push_back(owned_text(cache, "Item"));
    record.cells.push_back(Value::number(amount));
    cache.mutable_records().push_back(std::move(record));
  };
  // Number(1) and Text("1") render to the same label, but remain distinct
  // hierarchy owners. The same is repeated under two sorted groups.
  add("S", Value::number(3.0), 10.0);
  add("S", owned_text(cache, "3"), 20.0);
  add("D", Value::number(1.0), 30.0);
  add("D", owned_text(cache, "1"), 40.0);

  PivotTable table;
  table.set_pivot_cache_id(17);
  for (const char* name : {"Group", "Kind", "Item"}) {
    PivotField field;
    field.source_name = name;
    field.axis = PivotAxis::Row;
    table.mutable_fields().push_back(std::move(field));
  }
  PivotField amount_field;
  amount_field.source_name = "Amount";
  amount_field.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(amount_field));
  table.mutable_row_field_order() = {0, 1, 2};
  PivotDataField sum_amount;
  sum_amount.name = "Sum of Amount";
  sum_amount.field_index = 3;
  sum_amount.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum_amount));
  table.mutable_fields()[0].sort.ascending = false;
  table.mutable_fields()[1].subtotal_fns = {SubtotalFn::Sum, SubtotalFn::Average};
  table.mutable_fields()[0].subtotal_top = true;
  table.set_grand_totals(/*rows=*/true, /*cols=*/true);
  table.set_anchor(0, 0, 1, 1);

  auto result_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  const PivotResult& result = result_or.value();
  ASSERT_EQ(result.rows.size(), 2U);
  ASSERT_EQ(result.rows[0].label, "S");
  ASSERT_EQ(result.rows[1].label, "D");
  ASSERT_EQ(result.row_subtotals.size(), 10U);

  PivotLayoutOptions options;
  options.row_labels_label = "Row Labels";
  auto cells_or = layout(table, result, options);
  ASSERT_TRUE(static_cast<bool>(cells_or)) << cells_or.error().message << " " << cells_or.error().context;
  const PivotCells& cells = cells_or.value();

  std::vector<std::pair<PivotCellKind, double>> cells_in_render_order;
  for (const PivotCell& cell : cells.cells) {
    if (cell.col != cells.left + 1U || !cell.value.is_number()) {
      continue;
    }
    if (cell.kind == PivotCellKind::Data || cell.kind == PivotCellKind::RowSubtotal ||
        cell.kind == PivotCellKind::GrandTotal) {
      cells_in_render_order.emplace_back(cell.kind, cell.value.as_number());
    }
  }
  EXPECT_EQ(cells_in_render_order, (std::vector<std::pair<PivotCellKind, double>>{
                                       {PivotCellKind::RowSubtotal, 30.0},  // S subtotal.
                                       {PivotCellKind::RowSubtotal, 10.0},  // S / Number(3) SUM.
                                       {PivotCellKind::RowSubtotal, 10.0},  // S / Number(3) AVERAGE.
                                       {PivotCellKind::Data, 10.0},         // S / Number(3) detail.
                                       {PivotCellKind::RowSubtotal, 20.0},  // S / Text("3") SUM.
                                       {PivotCellKind::RowSubtotal, 20.0},  // S / Text("3") AVERAGE.
                                       {PivotCellKind::Data, 20.0},         // S / Text("3") detail.
                                       {PivotCellKind::RowSubtotal, 70.0},  // D subtotal.
                                       {PivotCellKind::RowSubtotal, 30.0},  // D / Number(1) SUM.
                                       {PivotCellKind::RowSubtotal, 30.0},  // D / Number(1) AVERAGE.
                                       {PivotCellKind::Data, 30.0},         // D / Number(1) detail.
                                       {PivotCellKind::RowSubtotal, 40.0},  // D / Text("1") SUM.
                                       {PivotCellKind::RowSubtotal, 40.0},  // D / Text("1") AVERAGE.
                                       {PivotCellKind::Data, 40.0},         // D / Text("1") detail.
                                       {PivotCellKind::GrandTotal, 100.0},
                                   }));
  EXPECT_EQ(cells.cols, 2U);
  EXPECT_GT(cells.rows, result.row_subtotals.size());
}

TEST(PivotEvaluator, SparseRowColumnIntersectionIsBlank) {
  PivotCache cache = build_basic_cache();
  auto& records = cache.mutable_records();
  records.erase(std::remove_if(records.begin(), records.end(),
                               [](const PivotCacheRecord& record) {
                                 return record.cells[0].as_text() == "North" && record.cells[1].as_text() == "Gadget";
                               }),
                records.end());
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{1});
  auto result_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  const PivotResult& result = result_or.value();
  const std::size_t north = row_index(result, "North");
  ASSERT_NE(north, static_cast<std::size_t>(-1));
  ASSERT_EQ(result.cols.size(), 2U);
  ASSERT_EQ(result.cols[0].label, "Gadget");
  EXPECT_TRUE(result.values[north][0][0].is_blank());
}

TEST(PivotEvaluator, UnresolvedHiddenItemDoesNotFilterBlankRecords) {
  PivotCache cache = build_basic_cache();
  PivotCacheRecord blank_region;
  blank_region.cells = {Value::blank(), owned_text(cache, "Widget"), Value::number(75.0)};
  cache.mutable_records().push_back(std::move(blank_region));
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  PivotItem malformed_hidden;
  malformed_hidden.visible = false;
  malformed_hidden.has_cache_index = true;
  malformed_hidden.cache_index = 999U;
  table.mutable_fields()[0].items.push_back(std::move(malformed_hidden));
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  auto result_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  const PivotResult& result = result_or.value();
  const std::size_t blank_leaf = row_index(result, PivotLayoutOptions{}.blank_item_label);
  ASSERT_NE(blank_leaf, static_cast<std::size_t>(-1));
  EXPECT_DOUBLE_EQ(result.values[blank_leaf][0][0].as_number(), 75.0);
}

// ---------------------------------------------------------------------------
// 3c. Blank source values still name their axis group
// ---------------------------------------------------------------------------

// A source row whose row-field cell is empty groups under a placeholder, not
// under a nameless node: the label is what the grid draws, and GETPIVOTDATA
// walks labels by exact match, so an unnamed group could not be addressed at
// all. What the placeholder spells is the vocabulary's business — these tests
// pin the mechanism, never a particular text.
TEST(PivotEvaluator, BlankAxisItemTakesThePlaceholderLabel) {
  PivotCache cache = build_basic_cache();
  PivotCacheRecord blank_region;
  blank_region.cells = {Value::blank(), owned_text(cache, "Widget"), Value::number(75.0)};
  cache.mutable_records().push_back(std::move(blank_region));
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  auto result_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  const PivotResult& result = result_or.value();

  for (const RowHierarchyNode& row : result.rows) {
    EXPECT_FALSE(row.label.empty());
  }
  const std::size_t blank_leaf = row_index(result, PivotLayoutOptions{}.blank_item_label);
  ASSERT_NE(blank_leaf, static_cast<std::size_t>(-1));
  EXPECT_DOUBLE_EQ(result.values[blank_leaf][0][0].as_number(), 75.0);
  // The named groups are untouched.
  EXPECT_NE(row_index(result, "North"), static_cast<std::size_t>(-1));
  EXPECT_NE(row_index(result, "South"), static_cast<std::size_t>(-1));
}

// The placeholder is part of the locale's label vocabulary, so it travels
// with the rest of it rather than being fixed in the evaluator.
TEST(PivotEvaluator, BlankAxisItemPlaceholderComesFromTheSuppliedVocabulary) {
  PivotCache cache = build_basic_cache();
  PivotCacheRecord blank_region;
  blank_region.cells = {Value::blank(), owned_text(cache, "Widget"), Value::number(75.0)};
  cache.mutable_records().push_back(std::move(blank_region));
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});

  PivotLayoutOptions options;
  options.blank_item_label = "<none>";
  auto result_or = evaluate(table, cache, options);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  EXPECT_NE(row_index(result_or.value(), "<none>"), static_cast<std::size_t>(-1));
}

// Subtotal rows address their group by the same label path, so an interior
// blank group has to carry the placeholder there too.
TEST(PivotEvaluator, BlankGroupSubtotalCarriesThePlaceholderLabel) {
  PivotCache cache = build_basic_cache();
  PivotCacheRecord blank_widget;
  blank_widget.cells = {Value::blank(), owned_text(cache, "Widget"), Value::number(75.0)};
  cache.mutable_records().push_back(std::move(blank_widget));
  PivotCacheRecord blank_gadget;
  blank_gadget.cells = {Value::blank(), owned_text(cache, "Gadget"), Value::number(25.0)};
  cache.mutable_records().push_back(std::move(blank_gadget));
  PivotTable table = build_sum_amount_table(/*row=*/{0, 1}, /*col=*/{});

  PivotLayoutOptions options;
  options.blank_item_label = "<no-value>";
  auto result_or = evaluate(table, cache, options);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  const PivotResult& result = result_or.value();

  bool saw_blank_subtotal = false;
  for (const RowSubtotal& subtotal : result.row_subtotals) {
    ASSERT_FALSE(subtotal.labels.empty());
    if (subtotal.labels[0] == options.blank_item_label) {
      saw_blank_subtotal = true;
      EXPECT_DOUBLE_EQ(subtotal.values[0].as_number(), 100.0);  // 75 + 25.
    }
    EXPECT_FALSE(subtotal.labels[0].empty());
  }
  EXPECT_TRUE(saw_blank_subtotal);
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

// ---------------------------------------------------------------------------
// 8e. AuthoredCaptionFilter (decoded from an OOXML `<filters>` block)
// ---------------------------------------------------------------------------

TEST(PivotEvaluator, AuthoredCaptionEqualFilterDropsRecords) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  AuthoredCaptionFilter f;
  f.field_index = 0;  // Region
  f.predicate = CaptionPredicate::Equal;
  f.value = "North";
  table.mutable_authored_caption_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, "North");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 175.0);  // 100 + 50 + 25
}

// Excel matches a caption filter the way it matches an AutoFilter
// criterion, so a criterion that differs only in case still selects.
TEST(PivotEvaluator, AuthoredCaptionFilterIgnoresCase) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  AuthoredCaptionFilter f;
  f.field_index = 0;
  f.predicate = CaptionPredicate::Equal;
  f.value = "nOrTh";
  table.mutable_authored_caption_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().rows.size(), 1U);
  EXPECT_EQ(r_or.value().rows[0].label, "North");
}

TEST(PivotEvaluator, AuthoredCaptionNotContainsFilterDropsMatchingRecords) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  AuthoredCaptionFilter f;
  f.field_index = 0;
  f.predicate = CaptionPredicate::NotContains;
  f.value = "outh";
  table.mutable_authored_caption_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().rows.size(), 1U);
  EXPECT_EQ(r_or.value().rows[0].label, "North");
}

TEST(PivotEvaluator, AuthoredCaptionBetweenFilterKeepsTheInclusiveRange) {
  PivotCache cache = build_basic_cache();
  // Row on Product so the range spans "Gadget" / "Widget".
  PivotTable table = build_sum_amount_table(/*row=*/{1}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  AuthoredCaptionFilter f;
  f.field_index = 1;  // Product
  f.predicate = CaptionPredicate::Between;
  f.value = "A";
  f.value_high = "M";
  table.mutable_authored_caption_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().rows.size(), 1U);
  EXPECT_EQ(r_or.value().rows[0].label, "Gadget");
  EXPECT_DOUBLE_EQ(r_or.value().values[0][0][0].as_number(), 350.0);  // 50 + 300
}

// A `fld` past the end of `<pivotFields>` is a malformed definition; it
// must filter nothing rather than drop every record.
TEST(PivotEvaluator, AuthoredCaptionFilterWithOutOfRangeFieldIsInert) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  AuthoredCaptionFilter f;
  f.field_index = 99;
  f.predicate = CaptionPredicate::Equal;
  f.value = "North";
  table.mutable_authored_caption_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  EXPECT_EQ(r_or.value().rows.size(), 2U);
}

// ---------------------------------------------------------------------------
// 8f. AuthoredValueFilter (the value and date half of the same block)
//
// The basic cache totals North 175 and South 500 by Region.
// ---------------------------------------------------------------------------

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

// A file names only the field; the axis it prunes has to be recovered
// from that field's place in the row / column order. A field on neither
// axis leaves nothing to prune.
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

// Ranked by aggregate, so it cannot be decided until the aggregates
// exist -- unlike its caption and date siblings, which prune records.
// Pinning it on the column axis proves the axis recovery is not
// hard-coded to rows.
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

TEST(PivotEvaluator, AuthoredDateFilterPrunesRecordsByTheirSerial) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  // Field 2 holds the amounts, which are plain numbers -- the date pass
  // reads whatever serial the bound cell carries, so a range that spans
  // only North's two smaller amounts drops South entirely.
  AuthoredValueFilter f;
  f.field_index = 2;
  f.type = FilterType::LabelDate;
  f.value = 0.0;
  f.value_high = 100.0;
  table.mutable_authored_value_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().rows.size(), 1U);
  EXPECT_EQ(r_or.value().rows[0].label, "North");
}

// ---------------------------------------------------------------------------
// 8g. Relative-period filters
//
// These carry no bounds in the file, so the substance is the calendar
// arithmetic that turns a period name plus a clock reading into a window.
// The boundaries are asserted directly against `serial_from_ymd` rather
// than against literal serials: the point under test is which calendar day
// each window starts and ends on, not the serial encoding, which
// `DateTimeDate` already pins.
// ---------------------------------------------------------------------------

namespace {

// 2026-05-20 12:00:00. Mid-month, mid-quarter, so every window's two
// boundaries are distinct from the reading itself.
constexpr eval::date_time::CivilTime kMay20{{2026, 5U, 20U}, {12U, 0U, 0U}};

double Serial(int year, unsigned month, unsigned day) {
  return eval::date_time::serial_from_ymd(year, month, day, /*date1904=*/false);
}

void ExpectWindow(RelativePeriod period, const eval::date_time::CivilTime& now, int low_y, unsigned low_m,
                  unsigned low_d, int high_y, unsigned high_m, unsigned high_d) {
  const DateWindow window = resolve_relative_period(period, now, /*date1904=*/false);
  EXPECT_DOUBLE_EQ(window.low, Serial(low_y, low_m, low_d));
  EXPECT_DOUBLE_EQ(window.high, Serial(high_y, high_m, high_d));
}

}  // namespace

TEST(RelativePeriod, DayWindowsAreSingleDays) {
  ExpectWindow(RelativePeriod::Today, kMay20, 2026, 5, 20, 2026, 5, 20);
  ExpectWindow(RelativePeriod::Yesterday, kMay20, 2026, 5, 19, 2026, 5, 19);
  ExpectWindow(RelativePeriod::Tomorrow, kMay20, 2026, 5, 21, 2026, 5, 21);
}

TEST(RelativePeriod, MonthWindowsEndOnTheirOwnLastDay) {
  // May has 31 days, April 30 -- the window end is derived from the next
  // month's first day, so a wrong month length would show up here.
  ExpectWindow(RelativePeriod::ThisMonth, kMay20, 2026, 5, 1, 2026, 5, 31);
  ExpectWindow(RelativePeriod::LastMonth, kMay20, 2026, 4, 1, 2026, 4, 30);
  ExpectWindow(RelativePeriod::NextMonth, kMay20, 2026, 6, 1, 2026, 6, 30);
}

TEST(RelativePeriod, QuarterWindowsSnapToTheCalendarQuarter) {
  // May sits in Q2, so "this quarter" starts in April rather than in May.
  ExpectWindow(RelativePeriod::ThisQuarter, kMay20, 2026, 4, 1, 2026, 6, 30);
  ExpectWindow(RelativePeriod::LastQuarter, kMay20, 2026, 1, 1, 2026, 3, 31);
  ExpectWindow(RelativePeriod::NextQuarter, kMay20, 2026, 7, 1, 2026, 9, 30);
}

TEST(RelativePeriod, YearWindowsSpanTheWholeCalendarYear) {
  ExpectWindow(RelativePeriod::ThisYear, kMay20, 2026, 1, 1, 2026, 12, 31);
  ExpectWindow(RelativePeriod::LastYear, kMay20, 2025, 1, 1, 2025, 12, 31);
  ExpectWindow(RelativePeriod::NextYear, kMay20, 2027, 1, 1, 2027, 12, 31);
}

TEST(RelativePeriod, YearToDateStopsAtTheReadingNotAtYearEnd) {
  ExpectWindow(RelativePeriod::YearToDate, kMay20, 2026, 1, 1, 2026, 5, 20);
}

TEST(RelativePeriod, WindowsRollOverTheYearBoundary) {
  // January is the case the explicit month arithmetic exists for: naive
  // subtraction would produce month 0 rather than the previous December.
  constexpr eval::date_time::CivilTime kJan10{{2026, 1U, 10U}, {0U, 0U, 0U}};
  ExpectWindow(RelativePeriod::LastMonth, kJan10, 2025, 12, 1, 2025, 12, 31);
  ExpectWindow(RelativePeriod::LastQuarter, kJan10, 2025, 10, 1, 2025, 12, 31);
  constexpr eval::date_time::CivilTime kDec10{{2026, 12U, 10U}, {0U, 0U, 0U}};
  ExpectWindow(RelativePeriod::NextMonth, kDec10, 2027, 1, 1, 2027, 1, 31);
  ExpectWindow(RelativePeriod::NextQuarter, kDec10, 2027, 1, 1, 2027, 3, 31);
}

TEST(RelativePeriod, LeapFebruaryKeepsItsTwentyNinthDay) {
  constexpr eval::date_time::CivilTime kFeb2024{{2024, 2U, 5U}, {0U, 0U, 0U}};
  ExpectWindow(RelativePeriod::ThisMonth, kFeb2024, 2024, 2, 1, 2024, 2, 29);
}

TEST(RelativePeriod, TheWindowFollowsTheWorkbookEpoch) {
  // A 1904-system workbook stores every date 1462 lower, so the resolved
  // window has to shift with it or it would select the wrong records.
  const DateWindow window = resolve_relative_period(RelativePeriod::ThisMonth, kMay20, /*date1904=*/true);
  EXPECT_DOUBLE_EQ(window.low, Serial(2026, 5, 1) - eval::date_time::kDate1904EpochGap);
  EXPECT_DOUBLE_EQ(window.high, Serial(2026, 5, 31) - eval::date_time::kDate1904EpochGap);
}

TEST(PivotEvaluator, AuthoredPeriodFilterPrunesRecordsAgainstThePinnedClock) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  // Field 2 carries the amounts (North 100 / 50 / 25, South 200 / 300),
  // read here as date serials. Pinning to 1900-02-15 makes "this quarter"
  // 1900-01-01..1900-03-31, i.e. serials 1..91: it admits North's 50 and
  // 25 and excludes everything else, South included.
  AuthoredPeriodFilter f;
  f.field_index = 2;
  f.period = RelativePeriod::ThisQuarter;
  table.mutable_authored_period_filters().push_back(f);

  PivotFilterEnv env;
  env.pinned_now = eval::date_time::CivilTime{{1900, 2U, 15U}, {0U, 0U, 0U}};
  auto r_or = evaluate(table, cache, PivotLayoutOptions{}, env);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().rows.size(), 1U);
  EXPECT_EQ(r_or.value().rows[0].label, "North");
  ASSERT_EQ(r_or.value().values.size(), 1U);
  ASSERT_EQ(r_or.value().values[0].size(), 1U);
  ASSERT_EQ(r_or.value().values[0][0].size(), 1U);
  EXPECT_DOUBLE_EQ(r_or.value().values[0][0][0].as_number(), 75.0);
}

TEST(PivotEvaluator, MovingThePinnedClockMovesWhichRecordsSurvive) {
  // Same filter, a quarter later: the window becomes serials 92..181,
  // which admits North's 100 alone and still excludes South.
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  AuthoredPeriodFilter f;
  f.field_index = 2;
  f.period = RelativePeriod::ThisQuarter;
  table.mutable_authored_period_filters().push_back(f);

  PivotFilterEnv env;
  env.pinned_now = eval::date_time::CivilTime{{1900, 5U, 15U}, {0U, 0U, 0U}};
  auto r_or = evaluate(table, cache, PivotLayoutOptions{}, env);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().rows.size(), 1U);
  EXPECT_EQ(r_or.value().rows[0].label, "North");
  ASSERT_EQ(r_or.value().values.size(), 1U);
  ASSERT_EQ(r_or.value().values[0].size(), 1U);
  ASSERT_EQ(r_or.value().values[0][0].size(), 1U);
  EXPECT_DOUBLE_EQ(r_or.value().values[0][0][0].as_number(), 100.0);
}

TEST(PivotEvaluator, APeriodFilterOnAnOutOfRangeFieldIsInert) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  AuthoredPeriodFilter f;
  f.field_index = 99;
  f.period = RelativePeriod::Today;
  table.mutable_authored_period_filters().push_back(f);

  PivotFilterEnv env;
  env.pinned_now = eval::date_time::CivilTime{{1900, 2U, 15U}, {0U, 0U, 0U}};
  auto r_or = evaluate(table, cache, PivotLayoutOptions{}, env);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  EXPECT_EQ(r_or.value().rows.size(), 2U);
}

// An unbounded range is a no-op rather than a half-open filter, matching
// how `PivotFilter` treats a missing upper bound.
TEST(PivotEvaluator, AuthoredDateFilterWithNoUpperBoundIsInert) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  AuthoredValueFilter f;
  f.field_index = 2;
  f.type = FilterType::LabelDate;
  f.value = 1000.0;
  table.mutable_authored_value_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  EXPECT_EQ(r_or.value().rows.size(), 2U);
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

// A source cell with no value is an axis item like any other, so the manual
// filter has to be able to hide it. It is the one item that cannot be named:
// its label is the locale's placeholder, which a genuine text value is free to
// spell identically, so the filter identifies it by the cache value the item
// binds to instead.
TEST(PivotEvaluator, ManualFilterHidesTheBlankItem) {
  PivotCache cache = build_shared_region_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  PivotItem hidden_blank;
  hidden_blank.visible = false;
  hidden_blank.has_cache_index = true;
  hidden_blank.cache_index = kSharedRegionBlank;
  table.mutable_fields()[0].items.push_back(hidden_blank);
  // The load path leaves the item unnamed, because the value it binds to
  // renders to nothing. Running it here pins that the filter works on the
  // shape a workbook actually arrives in.
  resolve_pivot_names(table, cache);
  ASSERT_TRUE(table.fields()[0].items[0].name.empty());

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  const PivotLayoutOptions defaults;
  ASSERT_EQ(r.rows.size(), 2U);
  EXPECT_DOUBLE_EQ(r.values[row_index(r, "North")][0][0].as_number(), 100.0);
  // The text value spelled like the placeholder is untouched; only the group
  // with no value is gone, and its 10 is out of every aggregate.
  const std::size_t placeholder_text = row_index(r, defaults.blank_item_label);
  ASSERT_NE(placeholder_text, static_cast<std::size_t>(-1));
  EXPECT_DOUBLE_EQ(r.values[placeholder_text][0][0].as_number(), 7.0);
}

// The mirror direction: hiding the text item that happens to be spelled like
// the placeholder must not take the blank group with it.
TEST(PivotEvaluator, HidingPlaceholderSpelledTextKeepsTheBlankItem) {
  PivotCache cache = build_shared_region_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  PivotItem hidden_text;
  hidden_text.visible = false;
  hidden_text.has_cache_index = true;
  hidden_text.cache_index = kSharedRegionPlaceholderText;
  table.mutable_fields()[0].items.push_back(hidden_text);
  resolve_pivot_names(table, cache);
  const PivotLayoutOptions defaults;
  ASSERT_EQ(table.fields()[0].items[0].name, defaults.blank_item_label);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  ASSERT_EQ(r.rows.size(), 2U);
  EXPECT_DOUBLE_EQ(r.values[row_index(r, "North")][0][0].as_number(), 100.0);
  const std::size_t blank_leaf = row_index(r, defaults.blank_item_label);
  ASSERT_NE(blank_leaf, static_cast<std::size_t>(-1));
  EXPECT_DOUBLE_EQ(r.values[blank_leaf][0][0].as_number(), 10.0);
}

TEST(PivotEvaluator, ManualFilterMatchesNumericDisplayLabel) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  // The numeric Amount field is rendered as its integral display label.
  // Hiding 300 removes only South/Gadget without allocating a label string
  // per record while the filter is evaluated.
  table.mutable_fields()[2].items = {PivotItem{"300", /*visible=*/false}};

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  const std::size_t north = row_index(r, "North");
  const std::size_t south = row_index(r, "South");
  ASSERT_NE(north, static_cast<std::size_t>(-1));
  ASSERT_NE(south, static_cast<std::size_t>(-1));
  EXPECT_DOUBLE_EQ(r.values[north][0][0].as_number(), 175.0);
  EXPECT_DOUBLE_EQ(r.values[south][0][0].as_number(), 200.0);
}

// A hidden item is matched against the record's label, so the two sides must
// spell a non-integral number the same way. A record-side renderer of its own
// would spell 1.5 as "1.500000" and 1e20 as its 27-character fixed-point form,
// neither of which any authored item can name — the records would stay in
// every aggregate no matter what the filter says.
TEST(PivotEvaluator, ManualFilterHidesNonIntegralNumericItem) {
  PivotCache cache = build_non_integral_amount_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  table.mutable_fields()[2].items = {PivotItem{display_string(Value::number(1.5)), /*visible=*/false}};

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  const std::size_t north = row_index(r, "North");
  ASSERT_NE(north, static_cast<std::size_t>(-1));
  // North keeps only the 2.0 record; the 1.5 one is gone from the aggregate.
  EXPECT_DOUBLE_EQ(r.values[north][0][0].as_number(), 2.0);
}

TEST(PivotEvaluator, ManualFilterHidesNumericItemBeyondIntegerRange) {
  PivotCache cache = build_non_integral_amount_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  table.mutable_fields()[2].items = {PivotItem{display_string(Value::number(1e20)), /*visible=*/false}};

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  const std::size_t south = row_index(r, "South");
  ASSERT_NE(south, static_cast<std::size_t>(-1));
  EXPECT_DOUBLE_EQ(r.values[south][0][0].as_number(), 4.0);
}

// A label filter reads the same rendering, so a needle taken from the axis
// label of a numeric item selects the record that produced it.
TEST(PivotEvaluator, LabelContainsFilterMatchesNumericItemLabel) {
  PivotCache cache = build_non_integral_amount_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Amount";
  f.type = FilterType::LabelContains;
  f.value = display_string(Value::number(1e20));  // "1E+20"
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, "South");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 1e20);
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

// ---------------------------------------------------------------------------
// Narrowing an embedder-supplied double to an index.
//
// A hand-built cache stores whatever number the embedder passed, and a value
// filter stores whatever count it was given. Both reach a `std::size_t`
// narrowing, which is undefined for NaN, for either infinity, and for any
// magnitude the destination cannot represent — and a trap, not a wrong
// answer, once the same code is compiled to wasm32.
// ---------------------------------------------------------------------------

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

TEST(PivotRecordAccess, OutOfDomainSharedItemIndexCollapsesToBlank) {
  // One shared item, so index 0 is the only resolvable reference. A cache
  // built through the mutation API leaves `cell_is_index` empty, which is
  // what makes a numeric cell in a shared field be read as an index.
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {Value::number(42.0)}});

  const struct {
    double stored;
    bool resolves;
  } cases[] = {
      {std::numeric_limits<double>::quiet_NaN(), false},
      {std::numeric_limits<double>::infinity(), false},
      {-std::numeric_limits<double>::infinity(), false},
      {-1.0, false},
      {1e30, false},
      {4.3e9, false},  // Past the wasm32 `size_t` range as well as past the container.
      {1.0, false},    // One past the only shared item.
      {0.0, true},
      {0.5, true},
  };

  for (const auto& c : cases) {
    PivotCacheRecord record;
    record.cells.push_back(Value::number(c.stored));
    ASSERT_TRUE(record.cell_is_index.empty());
    const Value v = cell_value(cache, record, 0);
    if (c.resolves) {
      ASSERT_TRUE(v.is_number()) << "stored=" << c.stored;
      EXPECT_DOUBLE_EQ(v.as_number(), 42.0) << "stored=" << c.stored;
    } else {
      EXPECT_TRUE(v.is_blank()) << "stored=" << c.stored;
    }
  }
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

// ---------------------------------------------------------------------------
// Value filters must leave every leaf-indexed structure in one index space
// ---------------------------------------------------------------------------

// Builds a cache with a two-level row axis (Region -> Product) and a two-level
// column axis (Year -> Quarter), so the evaluator emits both row and column
// subtotals. Each cell holds `region_product_base * year_quarter_multiplier`,
// which makes the leaf scores easy to rank:
//
//   base:       North/A=1  North/B=2  South/A=10  South/B=20
//   multiplier: 2024/Q1=1  2024/Q2=2  2025/Q1=10  2025/Q2=20
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

TEST(PivotEvaluator, RowValueFilterCompactsEveryLeafIndexedStructure) {
  PivotCache cache = build_subtotal_grid_cache();
  PivotTable table = build_subtotal_grid_table();

  // Row leaf scores are base * 33: North/A=33, North/B=66, South/A=330,
  // South/B=660. Top-2 therefore keeps both South leaves and prunes the
  // whole North group.
  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Region";
  f.type = FilterType::ValueTop10;
  f.value = 2;
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

  // Column leaf scores are multiplier * 33: 2024/Q1=33, 2024/Q2=66,
  // 2025/Q1=330, 2025/Q2=660. Top-2 keeps the 2025 quarters and prunes the
  // whole 2024 group.
  PivotFilter f;
  f.axis = PivotAxis::Col;
  f.field_name = "Year";
  f.type = FilterType::ValueTop10;
  f.value = 2;
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

// ---------------------------------------------------------------------------
// Three-level row axis, built from a cache rather than hand-assembled
// ---------------------------------------------------------------------------

// Two nesting levels only ever exercise one interior depth, so a subtotal
// owner's label path is never longer than one entry and the leaf enumeration
// never descends twice. A third level puts both under load: leaves are still
// numbered in DFS pre-order, and each owner emits its subtotal at its own
// depth with the full path to it.
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

// Pruning a three-level axis has to collapse whole branches: a group whose
// every descendant leaf is filtered away leaves no interior node and no
// subtotal behind, while the survivors are re-expressed in the compacted
// leaf-index space.
TEST(PivotEvaluator, RowValueFilterCollapsesEmptiedBranchesOfAThreeLevelAxis) {
  PivotCache cache = build_subtotal_grid_cache();
  PivotTable table = build_subtotal_grid_table();
  table.mutable_row_field_order() = {0, 1, 2};
  table.mutable_col_field_order() = {3};
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  // Leaf scores across the two quarter columns are 3, 30, 6, 60, 30, 300,
  // 60, 600. Top-2 keeps South/A/2025 and South/B/2025 only, so all of
  // North and both 2024 leaves under South disappear.
  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Region";
  f.type = FilterType::ValueTop10;
  f.value = 2;
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // The North branch is gone at every depth; South keeps both products but
  // only their 2025 child.
  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, "South");
  ASSERT_EQ(r.rows[0].children.size(), 2U);
  for (const RowHierarchyNode& product : r.rows[0].children) {
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
  // A surviving subtotal keeps its pre-filter aggregate, which is what lets
  // a Top-N report still frame a leaf against its whole group.
  EXPECT_DOUBLE_EQ(r.row_subtotals[0].values[0].as_number(), 330.0);
  EXPECT_DOUBLE_EQ(r.row_subtotals[2].values[0].as_number(), 990.0);
}

// ---------------------------------------------------------------------------
// Label filters resolve a field by its display name
// ---------------------------------------------------------------------------

TEST(PivotEvaluator, LabelFilterResolvesFieldByCustomName) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.mutable_fields()[0].custom_name = "Area";

  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Area";  // The display name, not the source name.
  f.type = FilterType::LabelBeginsWith;
  f.value = std::string("N");
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Only the North records survive, so both the axis and the aggregate move.
  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, "North");
  ASSERT_EQ(r.values.size(), 1U);
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 175.0);  // 100 + 50 + 25
  ASSERT_TRUE(r.grand_total.is_number());
  EXPECT_DOUBLE_EQ(r.grand_total.as_number(), 175.0);
}

TEST(PivotEvaluator, LabelFilterResolvesFieldByDataFieldName) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});

  // "Sum of Amount" is the data field's display name for the Amount field,
  // so the filter applies to Amount's own values.
  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Sum of Amount";
  f.type = FilterType::LabelBeginsWith;
  f.value = std::string("1");
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Only the Amount=100 record starts with "1".
  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, "North");
  ASSERT_TRUE(r.grand_total.is_number());
  EXPECT_DOUBLE_EQ(r.grand_total.as_number(), 100.0);
}

}  // namespace
}  // namespace formulon::pivot
