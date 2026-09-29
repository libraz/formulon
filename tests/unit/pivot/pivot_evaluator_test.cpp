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
#include <vector>

#include "eval/groupby_pivotby/common.h"
#include "gtest/gtest.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_layout.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot/record_access.h"
#include "pivot/value_order.h"
#include "pivot_evaluator_fixtures.h"
#include "utils/checked_index.h"
#include "utils/date_time.h"
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

using test::build_basic_cache;
using test::build_sum_amount_table;
using test::owned_text;
using test::row_index;

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

// `SortSpec::manual` (OOXML `sortType="manual"`) orders siblings by the
// field's `<items>` document position, resolved through the bound
// cache's `shared_items` -- not by display label. "South" sorts after
// "North" alphabetically, so this only passes if manual order actually
// overrides the default ascending-by-label sort.
TEST(PivotEvaluator, ManualSortOrdersByItemDocumentPosition) {
  PivotCache cache;
  cache.set_cache_id(1);
  PivotCacheField region;
  region.name = "Region";
  region.shared_items.push_back(owned_text(cache, "South"));
  region.shared_items.push_back(owned_text(cache, "North"));
  cache.mutable_fields().push_back(std::move(region));
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});

  auto add = [&](std::uint32_t region_index, double amount) {
    PivotCacheRecord rec;
    rec.cells = {Value::number(region_index), Value::number(amount)};
    rec.cell_is_index = {true, false};
    cache.mutable_records().push_back(std::move(rec));
  };
  add(0U, 200.0);  // South
  add(1U, 100.0);  // North

  PivotTable table;
  table.set_pivot_cache_id(1);
  PivotField region_f;
  region_f.source_name = "Region";
  region_f.axis = PivotAxis::Row;
  region_f.sort.manual = true;
  region_f.items.push_back(PivotItem{"South", true, /*has_cache_index=*/true, /*cache_index=*/0U});
  region_f.items.push_back(PivotItem{"North", true, /*has_cache_index=*/true, /*cache_index=*/1U});
  PivotField amount_f;
  amount_f.source_name = "Amount";
  amount_f.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(region_f));
  table.mutable_fields().push_back(std::move(amount_f));

  PivotDataField sum_amount;
  sum_amount.name = "Sum of Amount";
  sum_amount.field_index = 1;
  sum_amount.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum_amount));
  table.mutable_row_field_order() = {0};

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.rows.size(), 2U);
  EXPECT_EQ(r.rows[0].label, "South");
  EXPECT_EQ(r.rows[1].label, "North");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 200.0);
  EXPECT_DOUBLE_EQ(r.values[1][0][0].as_number(), 100.0);
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

  for (const AxisHierarchyNode& row : result.rows) {
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

// The variance/stddev family already maps a non-finite intermediate result
// to `#NUM!` (`variance_helper`); the arithmetic aggregates share the same
// obligation and, until fixed, returned `Value::number(inf)` instead.
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

// Max/Min pick an existing value rather than compute one, so they only
// see a non-finite result when the source cell itself already carries
// one (a cached value from outside Excel's own write path, since Excel
// never stores `Infinity`).
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

// A 1904-epoch workbook stores its cache serials on that epoch's scale.
// Bucketing them as if they were 1900-epoch (the pre-fix default) would
// misread the civil year by roughly four years; passing the workbook's
// actual epoch through `PivotFilterEnv` keeps the bucket on the year the
// record was actually authored under.
TEST(PivotEvaluator, DateGroupingByYearHonorsDate1904Epoch) {
  PivotCache cache = build_two_field_cache();
  push_record(cache, date_time::serial_from_ymd(2024, 3, 15, /*date1904=*/true), 10.0);
  push_record(cache, date_time::serial_from_ymd(2023, 6, 1, /*date1904=*/true), 20.0);

  PivotTable table = build_date_grouped_table(DateGrouping::Year, CalendarSystem::Gregorian);
  PivotFilterEnv env;
  env.date1904 = true;
  auto r_or = evaluate(table, cache, PivotLayoutOptions{}, env);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.rows.size(), 2U);
  EXPECT_EQ(r.rows[0].label, "2023");
  EXPECT_EQ(r.rows[1].label, "2024");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 20.0);
  EXPECT_DOUBLE_EQ(r.values[1][0][0].as_number(), 10.0);
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
// 13b. An axis field the reader could not decode grouping for
// ---------------------------------------------------------------------------

// The reader has no structural model for Excel's date/number grouping
// (`databaseField="0"` cache fields backed by a `<fieldGroup>`); every
// record's value for such a field is the placeholder `Value::blank()`.
// Placing one on an axis must refuse evaluation rather than silently
// collapse that axis to a single blank item.
TEST(PivotEvaluator, GroupingDerivedAxisFieldWithNoDateGroupIsRefused) {
  PivotCache cache = build_basic_cache();
  cache.mutable_fields().push_back(PivotCacheField{"Region Years", {}});
  cache.mutable_fields().back().is_database_field = false;
  cache.mutable_fields().back().field_group_xml = "<fieldGroup base=\"0\"><rangePr groupBy=\"years\"/></fieldGroup>";

  PivotTable table = build_sum_amount_table(/*row=*/{3}, /*col=*/{});
  PivotField grouped_f;
  grouped_f.source_name = "Region Years";
  grouped_f.axis = PivotAxis::Row;
  table.mutable_fields().push_back(std::move(grouped_f));
  table.mutable_row_field_order() = {3};

  auto r_or = evaluate(table, cache);
  ASSERT_FALSE(static_cast<bool>(r_or));
  EXPECT_EQ(r_or.error().code, FormulonErrorCode::kEvalPivotInvalid);
}

// A caller that has supplied its own `date_group` (via the C API setter)
// is exempt: evaluation proceeds using that grouping instead of refusing.
TEST(PivotEvaluator, GroupingDerivedAxisFieldWithDateGroupSetEvaluatesNormally) {
  PivotCache cache = build_basic_cache();
  cache.mutable_fields().push_back(PivotCacheField{"Region Years", {}});
  cache.mutable_fields().back().is_database_field = false;
  cache.mutable_fields().back().field_group_xml = "<fieldGroup base=\"0\"><rangePr groupBy=\"years\"/></fieldGroup>";

  PivotTable table = build_sum_amount_table(/*row=*/{3}, /*col=*/{});
  PivotField grouped_f;
  grouped_f.source_name = "Region Years";
  grouped_f.axis = PivotAxis::Row;
  grouped_f.date_group = PivotDateGroup{};
  table.mutable_fields().push_back(std::move(grouped_f));
  table.mutable_row_field_order() = {3};

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
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

// The arithmetic aggregations keep the short-circuit: the same fixture
// under SUM still reports the error, so the change above is scoped to the
// COUNT family rather than to error handling in general.
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

// Excel's own Top-N grand total covers only the visible rows, not the
// pre-filter set -- verified against two real Excel-authored fixtures
// (tests/fixtures/excel/pivot_value_date_filters.xlsx and
// pivot_recurring_period_filter.xlsx): both cache a Top-2 filter whose
// rendered Grand Total cell (675) equals the sum of the two surviving
// rows (500 + 175), not all four source rows.
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

}  // namespace
}  // namespace formulon::pivot
