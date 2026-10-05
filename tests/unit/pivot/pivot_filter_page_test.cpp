// Page (report-filter) field selection vocabulary and visibility.

#include <algorithm>
#include <cstddef>
#include <optional>
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
#include "pivot_filter_test_helpers.h"
#include "value.h"

namespace formulon::pivot {
namespace {

using filter_test::build_shared_north_south_cache;
using test::build_basic_cache;
using test::build_sum_amount_table;

PivotTable build_page_axis_table() {
  PivotTable table = build_sum_amount_table(/*row=*/{1}, /*col=*/{});
  table.mutable_fields()[0].axis = PivotAxis::Page;
  return table;
}

// Lists Region's two shared items, hiding whichever ones are named.
void set_region_items(PivotTable& table, const std::vector<std::string>& hidden) {
  std::vector<PivotItem>& items = table.mutable_fields()[0].items;
  items.clear();
  std::uint32_t index = 0;
  for (const char* name : {"North", "South"}) {
    PivotItem item;
    item.name = name;
    item.visible = std::find(hidden.begin(), hidden.end(), name) == hidden.end();
    item.has_cache_index = true;
    item.cache_index = index++;
    items.push_back(std::move(item));
  }
}

TEST(PivotPageSelection, UnfilteredFieldShowsTheAllPlaceholder) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_page_axis_table();

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& result = r_or.value();

  ASSERT_EQ(result.page_selections.size(), 1U);
  EXPECT_EQ(result.page_selections[0].field_index, 0U);
  EXPECT_EQ(result.page_selections[0].field_label, "Region");
  EXPECT_EQ(result.page_selections[0].item_label, "(All)");
}

TEST(PivotPageSelection, AnItemListThatHidesNothingIsStillUnfiltered) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_page_axis_table();
  set_region_items(table, /*hidden=*/{});

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().page_selections.size(), 1U);
  EXPECT_EQ(r_or.value().page_selections[0].item_label, "(All)");
}

TEST(PivotPageSelection, SoleVisibleItemNamesItself) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_page_axis_table();
  set_region_items(table, /*hidden=*/{"South"});

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().page_selections.size(), 1U);
  EXPECT_EQ(r_or.value().page_selections[0].item_label, "North");

  // The hidden item also prunes records, which is the pre-existing manual
  // filter path rather than anything the selection label drives.
  ASSERT_TRUE(r_or.value().grand_total.is_number());
  EXPECT_DOUBLE_EQ(r_or.value().grand_total.as_number(), 175.0);
}

TEST(PivotPageSelection, SoleVisibleUnnamedItemIsLabelledFromItsBinding) {
  PivotCache cache = build_shared_north_south_cache();
  PivotTable table = build_page_axis_table();
  for (std::uint32_t index : {0U, 1U}) {
    PivotItem item;
    item.visible = index == 1U;
    item.has_cache_index = true;
    item.cache_index = index;
    table.mutable_fields()[0].items.push_back(item);
  }

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().page_selections.size(), 1U);
  EXPECT_EQ(r_or.value().page_selections[0].item_label, "South");
  ASSERT_TRUE(r_or.value().grand_total.is_number());
  EXPECT_DOUBLE_EQ(r_or.value().grand_total.as_number(), 500.0);
}

TEST(PivotPageSelection, SeveralVisibleItemsUseTheMultiplePlaceholder) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_page_axis_table();
  // A third item nobody selected leaves two of three showing.
  set_region_items(table, /*hidden=*/{});
  PivotItem east;
  east.name = "East";
  east.visible = false;
  east.has_cache_index = true;
  east.cache_index = 2;
  table.mutable_fields()[0].items.push_back(std::move(east));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().page_selections.size(), 1U);
  EXPECT_EQ(r_or.value().page_selections[0].item_label, "(Multiple Items)");
}

TEST(PivotPageSelection, ExplicitPageFieldItemWinsOverVisibility) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_page_axis_table();
  set_region_items(table, /*hidden=*/{});
  table.mutable_page_fields().push_back(PivotPageField{0, std::optional<std::uint32_t>{1}});

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().page_selections.size(), 1U);
  EXPECT_EQ(r_or.value().page_selections[0].item_label, "South");
}

TEST(PivotPageSelection, OutOfRangeItemFallsBackToVisibility) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_page_axis_table();
  set_region_items(table, /*hidden=*/{});
  table.mutable_page_fields().push_back(PivotPageField{0, std::optional<std::uint32_t>{9}});

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().page_selections.size(), 1U);
  EXPECT_EQ(r_or.value().page_selections[0].item_label, "(All)");
}

TEST(PivotPageSelection, LocaleVocabularyNamesTheSelection) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_page_axis_table();
  PivotLayoutOptions options;
  options.all_pages_label = "(すべて)";

  auto r_or = evaluate(table, cache, options);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().page_selections.size(), 1U);
  EXPECT_EQ(r_or.value().page_selections[0].item_label, "(すべて)");
}

}  // namespace
}  // namespace formulon::pivot
