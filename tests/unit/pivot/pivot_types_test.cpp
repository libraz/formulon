// Copyright 2026 libraz. Licensed under the MIT License.
//
// Default-construction and basic invariant tests for the pivot data
// model. These structures are header-only and behaviour-free at this
// stage; the suite locks in the contract that subsequent PRs (cache,
// evaluator, GETPIVOTDATA) build on.

#include "pivot/pivot_types.h"

#include "gtest/gtest.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "value.h"

namespace formulon::pivot {
namespace {

TEST(PivotTypes, EnumDefaults) {
  // Spot-check the numeric assignments stay where the design pins them.
  EXPECT_EQ(static_cast<std::uint8_t>(PivotAxis::Row), 0u);
  EXPECT_EQ(static_cast<std::uint8_t>(PivotAxis::Col), 1u);
  EXPECT_EQ(static_cast<std::uint8_t>(PivotAxis::Value), 2u);
  EXPECT_EQ(static_cast<std::uint8_t>(PivotAxis::Page), 3u);

  EXPECT_EQ(static_cast<std::uint8_t>(Aggregation::Sum), 0u);
  EXPECT_EQ(static_cast<std::uint8_t>(SubtotalFn::Sum), 0u);

  EXPECT_EQ(static_cast<std::uint8_t>(PivotLayout::Compact), 0u);
  EXPECT_EQ(static_cast<std::uint8_t>(FilterType::ValueTop10), 0u);
  EXPECT_EQ(static_cast<std::uint8_t>(DateGrouping::Day), 0u);
  EXPECT_EQ(static_cast<std::uint8_t>(CalendarSystem::Gregorian), 0u);
}

TEST(PivotTypes, PivotItemDefault) {
  PivotItem item;
  EXPECT_TRUE(item.name.empty());
  EXPECT_TRUE(item.visible);
}

TEST(PivotTypes, PivotDateGroupDefault) {
  PivotDateGroup g;
  EXPECT_EQ(g.granularity, DateGrouping::Year);
  EXPECT_EQ(g.calendar, CalendarSystem::Gregorian);
  EXPECT_FALSE(g.start_year.has_value());
  EXPECT_FALSE(g.end_year.has_value());
}

TEST(PivotTypes, PivotFilterDefault) {
  PivotFilter f;
  EXPECT_EQ(f.axis, PivotAxis::Row);
  EXPECT_TRUE(f.field_name.empty());
  EXPECT_EQ(f.type, FilterType::ValueTop10);
  // Variant default-constructs to its first alternative (int = 0).
  ASSERT_EQ(f.value.index(), 0u);
  EXPECT_EQ(std::get<int>(f.value), 0);
}

TEST(PivotTypes, SortSpecDefault) {
  SortSpec s;
  EXPECT_TRUE(s.ascending);
  EXPECT_TRUE(s.by_field.empty());
}

TEST(PivotTypes, PivotFieldDefault) {
  PivotField f;
  EXPECT_TRUE(f.source_name.empty());
  EXPECT_TRUE(f.custom_name.empty());
  EXPECT_EQ(f.axis, PivotAxis::Row);
  EXPECT_TRUE(f.aggregations.empty());
  EXPECT_TRUE(f.sort.ascending);
  EXPECT_TRUE(f.items.empty());
  EXPECT_FALSE(f.subtotal_top);
  EXPECT_TRUE(f.subtotal_fns.empty());
  EXPECT_TRUE(f.number_format.empty());
  EXPECT_FALSE(f.date_group.has_value());
}

TEST(PivotCache, DefaultsAndAccessors) {
  PivotCache cache;
  EXPECT_EQ(cache.cache_id(), 0u);
  EXPECT_TRUE(cache.fields().empty());
  EXPECT_TRUE(cache.records().empty());
}

TEST(PivotCache, CacheIdRoundTrip) {
  PivotCache cache;
  cache.set_cache_id(7u);
  EXPECT_EQ(cache.cache_id(), 7u);
}

TEST(PivotCache, MutableFieldsAndRecords) {
  PivotCache cache;
  cache.mutable_fields().push_back(PivotCacheField{"region", {}});
  cache.mutable_records().push_back(PivotCacheRecord{});
  EXPECT_EQ(cache.fields().size(), 1u);
  EXPECT_EQ(cache.fields()[0].name, "region");
  EXPECT_EQ(cache.records().size(), 1u);
}

TEST(PivotResult, Defaults) {
  PivotResult r;
  EXPECT_TRUE(r.rows.empty());
  EXPECT_TRUE(r.cols.empty());
  EXPECT_TRUE(r.values.empty());
  EXPECT_TRUE(r.subtotals.empty());
  EXPECT_EQ(r.grand_total.kind(), ValueKind::Blank);
}

TEST(PivotTable, IdentityDefaults) {
  PivotTable t;
  EXPECT_TRUE(t.name().empty());
  EXPECT_EQ(t.pivot_cache_id(), 0u);
  EXPECT_TRUE(t.fields().empty());
  EXPECT_EQ(t.layout(), PivotLayout::Compact);
  EXPECT_EQ(t.anchor_row(), 0u);
  EXPECT_EQ(t.anchor_col(), 0u);
  EXPECT_EQ(t.span_rows(), 0u);
  EXPECT_EQ(t.span_cols(), 0u);
  EXPECT_TRUE(t.grand_totals_rows());
  EXPECT_TRUE(t.grand_totals_cols());
  EXPECT_TRUE(t.active_filters().empty());
  EXPECT_FALSE(t.last_result().has_value());
}

TEST(PivotTable, IdentityAccessors) {
  PivotTable t;
  t.set_name("PivotTable1");
  t.set_pivot_cache_id(5u);
  EXPECT_EQ(t.name(), "PivotTable1");
  EXPECT_EQ(t.pivot_cache_id(), 5u);
}

TEST(PivotTable, ContainsRespectsAnchorAndSpan) {
  PivotTable t;
  t.set_anchor(/*row=*/3, /*col=*/2, /*rows=*/4, /*cols=*/5);

  EXPECT_EQ(t.anchor_row(), 3u);
  EXPECT_EQ(t.anchor_col(), 2u);
  EXPECT_EQ(t.span_rows(), 4u);
  EXPECT_EQ(t.span_cols(), 5u);

  // Inside the bounds.
  EXPECT_TRUE(t.contains(3, 2));  // Top-left corner.
  EXPECT_TRUE(t.contains(6, 6));  // Bottom-right inclusive corner.
  EXPECT_TRUE(t.contains(4, 4));  // Interior.

  // Outside the bounds.
  EXPECT_FALSE(t.contains(2, 2));  // Above.
  EXPECT_FALSE(t.contains(3, 1));  // Left.
  EXPECT_FALSE(t.contains(7, 6));  // Below (one past last row).
  EXPECT_FALSE(t.contains(6, 7));  // Right (one past last col).
}

TEST(PivotTable, EmptySpanContainsNothing) {
  PivotTable t;
  // Default anchor/span is (0, 0, 0, 0): no cell is inside.
  EXPECT_FALSE(t.contains(0, 0));
  EXPECT_FALSE(t.contains(1, 1));
}

TEST(PivotTable, LayoutAccessor) {
  PivotTable t;
  t.set_layout(PivotLayout::Tabular);
  EXPECT_EQ(t.layout(), PivotLayout::Tabular);
}

TEST(PivotTable, GrandTotalsAccessor) {
  PivotTable t;
  t.set_grand_totals(/*rows=*/false, /*cols=*/false);
  EXPECT_FALSE(t.grand_totals_rows());
  EXPECT_FALSE(t.grand_totals_cols());
  t.set_grand_totals(/*rows=*/true, /*cols=*/false);
  EXPECT_TRUE(t.grand_totals_rows());
  EXPECT_FALSE(t.grand_totals_cols());
}

TEST(PivotTable, MutableActiveFilters) {
  PivotTable t;
  PivotFilter f;
  f.axis = PivotAxis::Col;
  f.field_name = "Region";
  f.type = FilterType::LabelContains;
  f.value = std::string{"North"};
  t.mutable_active_filters().push_back(std::move(f));

  ASSERT_EQ(t.active_filters().size(), 1u);
  EXPECT_EQ(t.active_filters()[0].axis, PivotAxis::Col);
  EXPECT_EQ(t.active_filters()[0].field_name, "Region");
  EXPECT_EQ(t.active_filters()[0].type, FilterType::LabelContains);
  ASSERT_EQ(t.active_filters()[0].value.index(), 2u);
  EXPECT_EQ(std::get<std::string>(t.active_filters()[0].value), "North");
}

TEST(PivotTable, MutableLastResultRoundTrip) {
  PivotTable t;
  EXPECT_FALSE(t.last_result().has_value());

  PivotResult r;
  r.grand_total = Value::number(42.0);
  t.mutable_last_result() = std::move(r);

  ASSERT_TRUE(t.last_result().has_value());
  EXPECT_EQ(t.last_result()->grand_total.kind(), ValueKind::Number);
  EXPECT_EQ(t.last_result()->grand_total.as_number(), 42.0);
}

}  // namespace
}  // namespace formulon::pivot
