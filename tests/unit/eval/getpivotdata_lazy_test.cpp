//
// Unit tests for the GETPIVOTDATA lazy form. The Ref-anchor path
// requires a workbook + pivot fixture so it cannot be expressed via
// the formula-only `EvalSource` helper. This file builds the workbook
// directly via the storage-layer + pivot-data-model APIs (no xlsx /
// no OOXML reader) and exercises the lookup path:
//
//   * data-field name + anchor only -> grand total
//   * single (row-axis field, item) pair -> per-leaf aggregation
//   * unknown field / item -> #REF!
//   * anchor not over a pivot -> #REF!
//   * arity / shape errors -> #REF!
//   * argument errors propagate verbatim
//
// All values are constructed in-memory; the OOXML wiring lands in a
// follow-up PR (the pivot reader will populate the same `PivotCache`
// and `PivotTable` structures these tests build by hand).

#include "eval/getpivotdata_lazy.h"

#include <cstdint>
#include <memory>
#include <string>
#include <string_view>
#include <utility>

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/pivot_locale.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

// Parses `src` and evaluates against `ctx`.
Value EvalWith(std::string_view src, const EvalContext& ctx) {
  static thread_local Arena arena;
  arena.reset();
  parser::Parser p(src, arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return evaluate(*root, arena, default_registry(), ctx);
}

// Pushes `s` into the cache's text storage so the returned `Value::text`
// holds a stable view for the lifetime of the cache.
Value owned_text(pivot::PivotCache& cache, std::string s) {
  cache.mutable_text_storage().push_back(std::move(s));
  return Value::text(cache.text_storage().back());
}

// ---------------------------------------------------------------------------
// Fixture: a one-sheet workbook with a Region/Amount pivot anchored at
// A3:B7.
//
//   Cache (cache_id=1):
//     Field 0 = "Region"  (shared_items: ["North", "South"])
//     Field 1 = "Amount"  (range-typed)
//
//   Records:
//     North 100
//     North 200
//     South 300
//     South 400
//
//   Table (cache_id=1, anchor 2..6 row, 0..1 col):
//     row_field_order = [0]      // Region
//     col_field_order = []        // no col axis
//     data_fields     = [{"Sum of Amount", 1, Sum}]
//
//   Expected aggregates:
//     North leaf    = 300
//     South leaf    = 700
//     grand total   = 1000
// ---------------------------------------------------------------------------

// Builds the cache (cache_id = 1).
std::unique_ptr<pivot::PivotCache> BuildBasicCache() {
  auto cache = std::make_unique<pivot::PivotCache>();
  cache->set_cache_id(1U);

  pivot::PivotCacheField region_field;
  region_field.name = "Region";
  region_field.shared_items.push_back(owned_text(*cache, "North"));
  region_field.shared_items.push_back(owned_text(*cache, "South"));
  cache->mutable_fields().push_back(std::move(region_field));

  pivot::PivotCacheField amount_field;
  amount_field.name = "Amount";
  cache->mutable_fields().push_back(std::move(amount_field));

  auto add = [&](double region_idx, double amount) {
    pivot::PivotCacheRecord rec;
    rec.cells.push_back(Value::number(region_idx));
    rec.cells.push_back(Value::number(amount));
    cache->mutable_records().push_back(std::move(rec));
  };
  add(0.0, 100.0);  // North 100
  add(0.0, 200.0);  // North 200
  add(1.0, 300.0);  // South 300
  add(1.0, 400.0);  // South 400
  return cache;
}

// Builds the pivot table (cache_id = 1) anchored at (2, 0)..(6, 1).
std::unique_ptr<pivot::PivotTable> BuildBasicTable() {
  auto table = std::make_unique<pivot::PivotTable>();
  table->set_name("PivotTable1");
  table->set_pivot_cache_id(1U);

  pivot::PivotField region_f;
  region_f.source_name = "Region";
  region_f.axis = pivot::PivotAxis::Row;
  table->mutable_fields().push_back(std::move(region_f));

  pivot::PivotField amount_f;
  amount_f.source_name = "Amount";
  amount_f.axis = pivot::PivotAxis::Value;
  table->mutable_fields().push_back(std::move(amount_f));

  pivot::PivotDataField sum_amount;
  sum_amount.name = "Sum of Amount";
  sum_amount.field_index = 1U;  // Amount in `fields()`.
  sum_amount.aggregation = pivot::Aggregation::Sum;
  table->mutable_data_fields().push_back(std::move(sum_amount));

  table->mutable_row_field_order().push_back(0U);  // Region.
  table->set_anchor(2U, 0U, 5U, 2U);
  return table;
}

// Builds a single-sheet workbook with the basic pivot wired in. The
// pivot anchor is at row 2 col 0 (A3 in 1-based). Cells are not
// populated; the storage layer treats unset coordinates as blank.
Workbook BuildBasicWorkbook() {
  Workbook wb = Workbook::create();
  wb.add_pivot_cache(BuildBasicCache());
  wb.sheet(0).add_pivot_table(BuildBasicTable());
  return wb;
}

// The basic workbook plus one record whose Region cell carries no value —
// the shape a source row with an empty row-field cell produces.
Workbook BuildBlankRegionWorkbook() {
  Workbook wb = Workbook::create();
  auto cache = BuildBasicCache();
  pivot::PivotCacheRecord blank_region;
  blank_region.cells.push_back(Value::blank());
  blank_region.cells.push_back(Value::number(500.0));
  cache->mutable_records().push_back(std::move(blank_region));
  wb.add_pivot_cache(std::move(cache));
  wb.sheet(0).add_pivot_table(BuildBasicTable());
  return wb;
}

Workbook BuildMultiDataWorkbook() {
  Workbook wb = Workbook::create();
  wb.add_pivot_cache(BuildBasicCache());
  auto table = BuildBasicTable();
  pivot::PivotDataField count_amount;
  count_amount.name = "Count of Amount";
  count_amount.field_index = 1U;
  count_amount.aggregation = pivot::Aggregation::Count;
  table->mutable_data_fields().push_back(std::move(count_amount));
  wb.sheet(0).add_pivot_table(std::move(table));
  return wb;
}

// ---------------------------------------------------------------------------
// Happy path
// ---------------------------------------------------------------------------

TEST(GetPivotDataLazy, AnchorOnlyReturnsGrandTotal) {
  Workbook wb = BuildBasicWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=GETPIVOTDATA(\"Sum of Amount\", A3)", ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1000.0);
}

TEST(GetPivotDataLazy, AnchorOnlyReturnsGrandTotalForSecondDataField) {
  Workbook wb = BuildMultiDataWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=GETPIVOTDATA(\"Count of Amount\", A3)", ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 4.0);
}

TEST(GetPivotDataLazy, RowFieldNorthReturnsLeafSum) {
  Workbook wb = BuildBasicWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=GETPIVOTDATA(\"Sum of Amount\", A3, \"Region\", \"North\")", ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 300.0);
}

TEST(GetPivotDataLazy, RowFieldSouthReturnsLeafSum) {
  Workbook wb = BuildBasicWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=GETPIVOTDATA(\"Sum of Amount\", A3, \"Region\", \"South\")", ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 700.0);
}

// The lookup matches axis labels exactly, so the group formed by blank source
// cells is only reachable if it carries the workbook locale's placeholder
// label. An unnamed group would leave those records addressable by nothing at
// all. The expected literal is read back out of the locale vocabulary rather
// than spelled here, so the test pins addressability, not the placeholder text.
TEST(GetPivotDataLazy, BlankRowFieldGroupIsAddressableByItsPlaceholderLabel) {
  Workbook wb = BuildBlankRegionWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const std::string placeholder = pivot_layout_options_for(wb.excel_profile()).blank_item_label;
  ASSERT_FALSE(placeholder.empty());
  const std::string formula = "=GETPIVOTDATA(\"Sum of Amount\", A3, \"Region\", \"" + placeholder + "\")";
  const Value v = EvalWith(formula, ctx);
  ASSERT_TRUE(v.is_number()) << "the blank group has no label a formula can name";
  EXPECT_EQ(v.as_number(), 500.0);
}

// ---------------------------------------------------------------------------
// Negative paths
// ---------------------------------------------------------------------------

TEST(GetPivotDataLazy, UnknownItemReturnsRef) {
  Workbook wb = BuildBasicWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  // "East" is not in the Region shared_items.
  const Value v = EvalWith("=GETPIVOTDATA(\"Sum of Amount\", A3, \"Region\", \"East\")", ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(GetPivotDataLazy, UnknownDataFieldReturnsRef) {
  Workbook wb = BuildBasicWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=GETPIVOTDATA(\"MissingField\", A3)", ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(GetPivotDataLazy, AnchorOutsidePivotReturnsRef) {
  Workbook wb = BuildBasicWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  // Z99 is well outside the pivot bounds (anchor at A3 spans 5x2).
  const Value v = EvalWith("=GETPIVOTDATA(\"Sum of Amount\", Z99)", ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(GetPivotDataLazy, OddFieldItemArityReturnsRef) {
  Workbook wb = BuildBasicWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  // Odd number of trailing args (3 args total -> 1 trailing, not paired).
  const Value v = EvalWith("=GETPIVOTDATA(\"Sum of Amount\", A3, \"Region\")", ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(GetPivotDataLazy, ZeroArityReturnsRef) {
  Workbook wb = BuildBasicWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=GETPIVOTDATA()", ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(GetPivotDataLazy, OneArityReturnsRef) {
  Workbook wb = BuildBasicWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=GETPIVOTDATA(\"Sum of Amount\")", ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

// ---------------------------------------------------------------------------
// Error propagation
// ---------------------------------------------------------------------------

TEST(GetPivotDataLazy, FirstArgErrorPropagates) {
  Workbook wb = BuildBasicWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  // The data-field arg is evaluated first; #DIV/0! must surface even
  // though the call also has an unrecognised anchor target downstream.
  const Value v = EvalWith("=GETPIVOTDATA(1/0, A3)", ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(GetPivotDataLazy, FieldItemErrorPropagates) {
  Workbook wb = BuildBasicWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  // An error inside a (field, item) pair must propagate before the
  // structural #REF! reject fires.
  const Value v = EvalWith("=GETPIVOTDATA(\"Sum of Amount\", A3, 1/0, \"North\")", ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(GetPivotDataLazy, NonRefAnchorWithErrorPropagates) {
  Workbook wb = BuildBasicWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  // Arg 1 is not a Ref; the impl eagerly evaluates it for error
  // propagation, so an embedded error must surface instead of #REF!.
  const Value v = EvalWith("=GETPIVOTDATA(\"Sum of Amount\", 1/0)", ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// Anchor variants
// ---------------------------------------------------------------------------

TEST(GetPivotDataLazy, RangeAnchorUsesTopLeftCell) {
  // RangeOp `A3:B7` parses as a range; the impl uses the leftmost
  // descendant Ref (`A3`) as the anchor and resolves the same pivot.
  Workbook wb = BuildBasicWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=GETPIVOTDATA(\"Sum of Amount\", A3:B7)", ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1000.0);
}

TEST(GetPivotDataLazy, NonRefAnchorReturnsRef) {
  // Arg 1 is a literal text (not a Ref / RangeOp); the structural
  // check rejects it as #REF!.
  Workbook wb = BuildBasicWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=GETPIVOTDATA(\"Sum of Amount\", \"PivotTable1\")", ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
