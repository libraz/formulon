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
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_locale.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/date_time.h"
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

// A one-field pivot whose two source values render to the same caption. The
// underlying types remain distinct so GETPIVOTDATA must use typed identity
// before it considers a display-label fallback.
Workbook BuildNumericTextCollisionWorkbook() {
  auto cache = std::make_unique<pivot::PivotCache>();
  cache->set_cache_id(1U);
  pivot::PivotCacheField region_field;
  region_field.name = "Region";
  cache->mutable_fields().push_back(std::move(region_field));
  pivot::PivotCacheField amount_field;
  amount_field.name = "Amount";
  cache->mutable_fields().push_back(std::move(amount_field));

  pivot::PivotCacheRecord numeric;
  numeric.cells = {Value::number(1.0), Value::number(30.0)};
  cache->mutable_records().push_back(std::move(numeric));
  pivot::PivotCacheRecord text;
  text.cells = {owned_text(*cache, "1"), Value::number(40.0)};
  cache->mutable_records().push_back(std::move(text));

  auto table = BuildBasicTable();
  Workbook wb = Workbook::create();
  wb.add_pivot_cache(std::move(cache));
  wb.sheet(0).add_pivot_table(std::move(table));
  return wb;
}

// A two-axis pivot where both distinct fields expose the same custom caption.
// A GETPIVOTDATA field name matching that caption is ambiguous and must not
// silently bind to whichever axis happens to appear first in the model.
Workbook BuildDuplicateCaptionWorkbook() {
  auto cache = std::make_unique<pivot::PivotCache>();
  cache->set_cache_id(1U);
  for (const char* name : {"Region", "Product", "Amount"}) {
    pivot::PivotCacheField field;
    field.name = name;
    cache->mutable_fields().push_back(std::move(field));
  }
  auto add = [&](const char* region, const char* product, double amount) {
    pivot::PivotCacheRecord record;
    record.cells = {owned_text(*cache, region), owned_text(*cache, product), Value::number(amount)};
    cache->mutable_records().push_back(std::move(record));
  };
  add("North", "A", 10.0);
  add("North", "B", 20.0);
  add("South", "A", 30.0);
  add("South", "B", 40.0);

  auto table = std::make_unique<pivot::PivotTable>();
  table->set_name("PivotTable1");
  table->set_pivot_cache_id(1U);
  pivot::PivotField region;
  region.source_name = "Region";
  region.custom_name = "Axis";
  region.axis = pivot::PivotAxis::Row;
  table->mutable_fields().push_back(std::move(region));
  pivot::PivotField product;
  product.source_name = "Product";
  product.custom_name = "Axis";
  product.axis = pivot::PivotAxis::Col;
  table->mutable_fields().push_back(std::move(product));
  pivot::PivotField amount;
  amount.source_name = "Amount";
  amount.axis = pivot::PivotAxis::Value;
  table->mutable_fields().push_back(std::move(amount));
  pivot::PivotDataField sum_amount;
  sum_amount.name = "Sum of Amount";
  sum_amount.field_index = 2U;
  sum_amount.aggregation = pivot::Aggregation::Sum;
  table->mutable_data_fields().push_back(std::move(sum_amount));
  table->mutable_row_field_order().push_back(0U);
  table->mutable_col_field_order().push_back(1U);
  table->set_anchor(2U, 0U, 6U, 4U);

  Workbook wb = Workbook::create();
  wb.add_pivot_cache(std::move(cache));
  wb.sheet(0).add_pivot_table(std::move(table));
  return wb;
}

// A two-level row-axis pivot keeps the numeric/text collision under two
// different parents so the lookup must preserve both typed identity and the
// preceding-parent leaf offset.
Workbook BuildNestedCollisionWorkbook() {
  auto cache = std::make_unique<pivot::PivotCache>();
  cache->set_cache_id(1U);
  for (const char* name : {"Region", "Product", "Amount"}) {
    pivot::PivotCacheField field;
    field.name = name;
    cache->mutable_fields().push_back(std::move(field));
  }
  auto add = [&](const char* region, Value product, double amount) {
    pivot::PivotCacheRecord record;
    record.cells = {owned_text(*cache, region), product, Value::number(amount)};
    cache->mutable_records().push_back(std::move(record));
  };
  add("North", Value::number(1.0), 30.0);
  add("North", owned_text(*cache, "1"), 40.0);
  add("South", Value::number(1.0), 50.0);
  add("South", owned_text(*cache, "1"), 60.0);

  auto table = std::make_unique<pivot::PivotTable>();
  table->set_name("PivotTable1");
  table->set_pivot_cache_id(1U);
  for (const char* name : {"Region", "Product", "Amount"}) {
    pivot::PivotField field;
    field.source_name = name;
    field.axis = std::string_view(name) == "Amount" ? pivot::PivotAxis::Value : pivot::PivotAxis::Row;
    table->mutable_fields().push_back(std::move(field));
  }
  pivot::PivotDataField sum_amount;
  sum_amount.name = "Sum of Amount";
  sum_amount.field_index = 2U;
  sum_amount.aggregation = pivot::Aggregation::Sum;
  table->mutable_data_fields().push_back(std::move(sum_amount));
  table->mutable_row_field_order() = {0U, 1U};
  table->set_anchor(2U, 0U, 8U, 2U);

  Workbook wb = Workbook::create();
  wb.add_pivot_cache(std::move(cache));
  wb.sheet(0).add_pivot_table(std::move(table));
  return wb;
}

Workbook BuildGroupedDateWorkbook(pivot::DateGrouping granularity = pivot::DateGrouping::Year) {
  auto cache = std::make_unique<pivot::PivotCache>();
  cache->set_cache_id(1U);
  cache->mutable_fields().push_back(pivot::PivotCacheField{"Date", {}});
  cache->mutable_fields().push_back(pivot::PivotCacheField{"Amount", {}});
  auto add = [&](int year, unsigned month, unsigned day, double amount) {
    pivot::PivotCacheRecord record;
    record.cells = {Value::number(date_time::serial_from_ymd(year, month, day, false)), Value::number(amount)};
    cache->mutable_records().push_back(std::move(record));
  };
  add(2024, 1U, 1U, 10.0);
  add(2024, 6U, 15U, 20.0);
  add(2025, 3U, 1U, 40.0);

  auto table = std::make_unique<pivot::PivotTable>();
  table->set_name("PivotTable1");
  table->set_pivot_cache_id(1U);
  pivot::PivotField date;
  date.source_name = "Date";
  date.axis = pivot::PivotAxis::Row;
  pivot::PivotDateGroup date_group;
  date_group.granularity = granularity;
  date_group.calendar = pivot::CalendarSystem::Gregorian;
  date.date_group = date_group;
  table->mutable_fields().push_back(std::move(date));
  pivot::PivotField amount;
  amount.source_name = "Amount";
  amount.axis = pivot::PivotAxis::Value;
  table->mutable_fields().push_back(std::move(amount));
  pivot::PivotDataField sum_amount;
  sum_amount.name = "Sum of Amount";
  sum_amount.field_index = 1U;
  sum_amount.aggregation = pivot::Aggregation::Sum;
  table->mutable_data_fields().push_back(std::move(sum_amount));
  table->mutable_row_field_order() = {0U};
  table->set_anchor(2U, 0U, 5U, 2U);

  Workbook wb = Workbook::create();
  wb.add_pivot_cache(std::move(cache));
  wb.sheet(0).add_pivot_table(std::move(table));
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

TEST(GetPivotDataLazy, NumericAndTextItemsWithSameLabelUseTypedIdentity) {
  Workbook wb = BuildNumericTextCollisionWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value numeric = EvalWith("=GETPIVOTDATA(\"Sum of Amount\", A3, \"Region\", 1)", ctx);
  const Value text = EvalWith("=GETPIVOTDATA(\"Sum of Amount\", A3, \"Region\", \"1\")", ctx);
  ASSERT_TRUE(numeric.is_number());
  ASSERT_TRUE(text.is_number());
  EXPECT_EQ(numeric.as_number(), 30.0);
  EXPECT_EQ(text.as_number(), 40.0);
}

TEST(GetPivotDataLazy, DuplicateCustomCaptionAcrossAxesReturnsRef) {
  Workbook wb = BuildDuplicateCaptionWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=GETPIVOTDATA(\"Sum of Amount\", A3, \"Axis\", \"North\")", ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(GetPivotDataLazy, DistinctRowAndColumnFieldsUseTheirLeafOffsets) {
  Workbook wb = BuildDuplicateCaptionWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=GETPIVOTDATA(\"Sum of Amount\", A3, \"Region\", \"South\", \"Product\", \"B\")", ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 40.0);
}

TEST(GetPivotDataLazy, NestedParentsPreserveTypedLeafIdentity) {
  Workbook wb = BuildNestedCollisionWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value numeric = EvalWith("=GETPIVOTDATA(\"Sum of Amount\", A3, \"Region\", \"South\", \"Product\", 1)", ctx);
  const Value text = EvalWith("=GETPIVOTDATA(\"Sum of Amount\", A3, \"Region\", \"South\", \"Product\", \"1\")", ctx);
  ASSERT_TRUE(numeric.is_number());
  ASSERT_TRUE(text.is_number());
  EXPECT_EQ(numeric.as_number(), 50.0);
  EXPECT_EQ(text.as_number(), 60.0);
}

TEST(GetPivotDataLazy, GroupedDisplayLabelIsUniqueFallback) {
  Workbook wb = BuildGroupedDateWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=GETPIVOTDATA(\"Sum of Amount\", A3, \"Date\", \"2024\")", ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 30.0);
}

TEST(GetPivotDataLazy, DateGroupSyntheticSortKeyIsNotAnItem) {
  // Month buckets sort on y*100+m; that key is not a source value and must
  // not address the group, while the group's label still does.
  Workbook wb = BuildGroupedDateWorkbook(pivot::DateGrouping::Month);
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value by_key = EvalWith("=GETPIVOTDATA(\"Sum of Amount\", A3, \"Date\", 202401)", ctx);
  ASSERT_TRUE(by_key.is_error());
  EXPECT_EQ(by_key.as_error(), ErrorCode::Ref);
  const Value by_label = EvalWith("=GETPIVOTDATA(\"Sum of Amount\", A3, \"Date\", \"2024-01\")", ctx);
  ASSERT_TRUE(by_label.is_number());
  EXPECT_EQ(by_label.as_number(), 10.0);
}

TEST(GetPivotDataLazy, AxisIdentityOwnsCopiedText) {
  Workbook wb = BuildBasicWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value first = EvalWith("=GETPIVOTDATA(\"Sum of Amount\", A3, \"Region\", \"North\")", ctx);
  ASSERT_TRUE(first.is_number());
  EXPECT_EQ(first.as_number(), 300.0);

  const auto& table = *wb.sheet(0).pivot_tables().front();
  const std::shared_ptr<const pivot::PivotResult> snapshot = table.last_result();
  ASSERT_NE(snapshot, nullptr);
  ASSERT_EQ(snapshot->rows.size(), 2U);
  ASSERT_TRUE(snapshot->rows[0].identity.has_value());
  ASSERT_TRUE(snapshot->rows[0].identity->is_text());
  EXPECT_EQ(snapshot->rows[0].identity->as_text(), "North");

  auto& cache = *wb.mutable_pivot_caches().front();
  ASSERT_FALSE(cache.mutable_text_storage().empty());
  cache.mutable_text_storage().front() = "Changed";
  EXPECT_EQ(snapshot->rows[0].identity->as_text(), "North");
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
  const std::string placeholder = pivot::pivot_layout_options_for(wb.excel_profile()).blank_item_label;
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
