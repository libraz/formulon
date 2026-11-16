//
// An item caption set to "" is stored as `<item n="" .../>` and Excel draws the
// item with an empty label: the row-label cell is empty, GETPIVOTDATA answers
// to "" and a show-values-as base item resolves to it. The model holds that as
// `PivotItem::blank_caption` with an empty caption, which must not fall back to
// the item's own label (or to the locale's "(blank)" placeholder).

#include <cstddef>
#include <cstdint>
#include <memory>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_evaluator.h"
#include "pivot/pivot_index.h"
#include "pivot/pivot_layout.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "sheet.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

namespace formulon::pivot {
namespace {

Value owned_text(PivotCache& cache, std::string s) {
  cache.mutable_text_storage().push_back(std::move(s));
  return Value::text(cache.text_storage().back());
}

// Region / Amount records: East 100, East 50, South 200, West 300. Region
// shares items [East, South, West].
PivotCache build_cache() {
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Region", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});
  for (const char* region : {"East", "South", "West"}) {
    cache.mutable_fields()[0].shared_items.push_back(owned_text(cache, region));
  }
  const auto add = [&](const char* region, double amount) {
    PivotCacheRecord rec;
    rec.cells.push_back(owned_text(cache, region));
    rec.cells.push_back(Value::number(amount));
    cache.mutable_records().push_back(std::move(rec));
  };
  add("East", 100.0);
  add("East", 50.0);
  add("South", 200.0);
  add("West", 300.0);
  return cache;
}

// Region on the rows, Sum of Qty; `east_blank` sets the empty caption on East, West is
// renamed "X" as the control.
PivotTable build_table(bool east_blank) {
  PivotTable table;
  table.set_name("Pivot1");
  table.set_pivot_cache_id(1);
  table.set_anchor(2, 3, 1, 1);

  PivotField region;
  region.source_name = "Region";
  region.axis = PivotAxis::Row;
  for (std::uint32_t x = 0; x < 3U; ++x) {
    PivotItem item;
    item.has_cache_index = true;
    item.cache_index = x;
    if (x == 0U) {
      item.blank_caption = east_blank;
    } else if (x == 2U) {
      item.caption = "X";
    }
    region.items.push_back(std::move(item));
  }
  PivotField amount;
  amount.source_name = "Amount";
  amount.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(region));
  table.mutable_fields().push_back(std::move(amount));
  table.mutable_row_field_order() = {0};

  PivotDataField sum;
  sum.name = "Sum of Qty";
  sum.field_index = 1;
  sum.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum));
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  return table;
}

// 1a. The evaluated row labels carry the empty caption: no fallback to
// "East", and no "(blank)" placeholder (which names a source cell with no
// value, a different thing).
TEST(PivotEmptyItemLabel, EvaluatedRowLabelIsEmptyNotTheItemName) {
  PivotCache cache = build_cache();
  PivotTable table = build_table(/*east_blank=*/true);
  resolve_pivot_names(table, cache);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.rows.size(), 3U);
  // Rows keep the source-value order: East, South, West.
  EXPECT_EQ(r.rows[0].label, "");
  EXPECT_EQ(r.rows[1].label, "South");
  EXPECT_EQ(r.rows[2].label, "X");
  ASSERT_TRUE(r.rows[0].identity.has_value());
  ASSERT_TRUE(r.rows[0].identity->is_text());
  EXPECT_EQ(r.rows[0].identity->as_text(), "East");
}

// 1b. The rendered row-label cell. An empty label is emitted as a blank value
// (the layout writes `Value::blank()` for an empty label rather than empty
// text), so the cell exists as a RowLabel and is blank; the "X" control still
// shows its caption.
TEST(PivotEmptyItemLabel, LayoutRowLabelCellIsBlank) {
  PivotCache cache = build_cache();
  PivotTable table = build_table(/*east_blank=*/true);
  resolve_pivot_names(table, cache);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  auto cells_or = layout(table, r_or.value());
  ASSERT_TRUE(static_cast<bool>(cells_or)) << cells_or.error().message;

  std::vector<const PivotCell*> labels;
  for (const PivotCell& cell : cells_or.value().cells) {
    if (cell.kind == PivotCellKind::RowLabel) {
      labels.push_back(&cell);
    }
  }
  ASSERT_EQ(labels.size(), 3U);
  EXPECT_TRUE(labels[0]->value.is_blank());
  ASSERT_TRUE(labels[1]->value.is_text());
  EXPECT_EQ(labels[1]->value.as_text(), "South");
  ASSERT_TRUE(labels[2]->value.is_text());
  EXPECT_EQ(labels[2]->value.as_text(), "X");
}

// 3. A percent-of whose base item is the empty-labelled one resolves to that
// item: its own cell is 100% and the others are relative to it. The label
// check pins the premise that the base is drawn empty.
TEST(PivotEmptyItemLabel, ShowAsPercentOfBaseItemResolvesTheEmptyLabelledItem) {
  PivotCache cache = build_cache();
  PivotTable table = build_table(/*east_blank=*/true);
  resolve_pivot_names(table, cache);
  PivotDataField& df = table.mutable_data_fields()[0];
  df.show_as = ShowValuesAs::PercentOf;
  df.show_as_base_field = 0U;
  df.show_as_base_item = 0U;  // East, drawn as "".

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.rows.size(), 3U);
  EXPECT_EQ(r.rows[0].label, "");
  ASSERT_EQ(r.values.size(), 3U);
  ASSERT_TRUE(r.values[0][0][0].is_number());
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 1.0);
  ASSERT_TRUE(r.values[1][0][0].is_number());
  EXPECT_DOUBLE_EQ(r.values[1][0][0].as_number(), 200.0 / 150.0);
  ASSERT_TRUE(r.values[2][0][0].is_number());
  EXPECT_DOUBLE_EQ(r.values[2][0][0].as_number(), 300.0 / 150.0);
}

// ---------------------------------------------------------------------------
// GETPIVOTDATA
// ---------------------------------------------------------------------------

Workbook build_workbook() {
  PivotCache cache = build_cache();
  PivotTable table = build_table(/*east_blank=*/true);
  table.set_anchor(2U, 0U, 5U, 2U);  // A3.
  resolve_pivot_names(table, cache);
  Workbook wb = Workbook::create();
  wb.add_pivot_cache(std::make_unique<PivotCache>(std::move(cache)));
  wb.sheet(0).add_pivot_table(std::make_unique<PivotTable>(std::move(table)));
  return wb;
}

Value eval_formula(std::string_view src, const eval::EvalContext& ctx) {
  static thread_local Arena arena;
  arena.reset();
  parser::Parser p(src, arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return eval::evaluate(*root, arena, eval::default_registry(), ctx);
}

// 2a. The empty label addresses the renamed item.
TEST(PivotEmptyItemLabel, GetPivotDataAddressesTheItemByItsEmptyLabel) {
  Workbook wb = build_workbook();
  eval::EvalState state;
  const eval::EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = eval_formula("=GETPIVOTDATA(\"Sum of Qty\", A3, \"Region\", \"\")", ctx);
  ASSERT_TRUE(v.is_number()) << "got error " << (v.is_error() ? static_cast<int>(v.as_error()) : -1);
  EXPECT_DOUBLE_EQ(v.as_number(), 150.0);
  // Neighbours are unaffected by the empty label.
  const Value x = eval_formula("=GETPIVOTDATA(\"Sum of Qty\", A3, \"Region\", \"X\")", ctx);
  ASSERT_TRUE(x.is_number());
  EXPECT_DOUBLE_EQ(x.as_number(), 300.0);
}

// 2b. Existing rule (getpivotdata_lazy_test RenamedItemMatchesItsCaptionAndItsSourceName):
// a renamed text item still answers to the source text it renames, so "East"
// keeps returning the value. This holds today and must keep holding.
// ASSUMPTION: the task brief expects #REF! for the old label in Excel; the
// repo's existing rule is followed instead, as instructed.
TEST(PivotEmptyItemLabel, GetPivotDataStillMatchesTheSourceTextOfTheEmptyLabelledItem) {
  Workbook wb = build_workbook();
  eval::EvalState state;
  const eval::EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = eval_formula("=GETPIVOTDATA(\"Sum of Qty\", A3, \"Region\", \"East\")", ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 150.0);
}

}  // namespace
}  // namespace formulon::pivot
