// Format and chart-format areas across a refresh, per the rules Excel applies:
// an area's <x v> follows its item's position in the field's items (data
// pseudo-field: its data field's index), and a format naming a field, item or
// data field the refresh removed is removed whole.

#include <cstddef>
#include <cstdint>
#include <initializer_list>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "eval/eval_profile_scope.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "pivot/field_remap.h"
#include "pivot/filter_engine.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_locale.h"
#include "pivot/pivot_refresh.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "value.h"

namespace formulon {
namespace {

using pivot::PivotAxis;
using pivot::PivotCache;
using pivot::PivotCacheField;
using pivot::PivotChartFormat;
using pivot::PivotDataField;
using pivot::PivotField;
using pivot::PivotFormat;
using pivot::PivotItem;
using pivot::PivotTable;
using pivot::PivotTableArea;
using pivot::PivotTableAreaReference;
using pivot::RefreshContext;
using pivot::RefreshPlan;
using pivot::RefreshSource;

constexpr std::uint32_t kValues = static_cast<std::uint32_t>(pivot::kValuesFieldPosition);
constexpr std::uint32_t kItemField = 0U;

// From a range: GCC 14 -Werror=stringop-overflow misjudges a vector built from a one-element braced list.
std::vector<std::uint32_t> U32s(std::initializer_list<std::uint32_t> values) {
  return std::vector<std::uint32_t>(values.begin(), values.end());
}

PivotTableAreaReference Ref(std::uint32_t field, std::initializer_list<std::uint32_t> items) {
  PivotTableAreaReference ref;
  ref.field = field;
  ref.items = U32s(items);
  return ref;
}

PivotTableArea Area(std::vector<PivotTableAreaReference> refs) {
  PivotTableArea area;
  area.references = std::move(refs);
  return area;
}

PivotFormat Format(std::uint32_t dxf, std::vector<PivotTableAreaReference> refs) {
  PivotFormat format;
  format.dxf_id = dxf;
  format.area = Area(std::move(refs));
  return format;
}

PivotChartFormat ChartFormat(std::uint32_t id, bool series, std::vector<PivotTableAreaReference> refs) {
  PivotChartFormat format;
  format.format = id;
  format.series = series;
  format.area = Area(std::move(refs));
  return format;
}

PivotItem Item(std::uint32_t x) {
  PivotItem item;
  item.has_cache_index = true;
  item.cache_index = x;
  return item;
}

PivotDataField DataField(std::string name, std::uint32_t field) {
  PivotDataField data;
  data.name = std::move(name);
  data.field_index = field;
  return data;
}

// Item / Amt / Qty, Item's shared items B C D.
PivotCache OldCache() {
  PivotCache cache;
  cache.set_cache_id(1U);
  cache.mutable_fields().push_back(PivotCacheField{"Item", {Value::text("B"), Value::text("C"), Value::text("D")}});
  cache.mutable_fields().push_back(PivotCacheField{"Amt", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Qty", {}});
  return cache;
}

// Item on rows sorted ascending, Sum of Amt and Sum of Qty; formats on B's label,
// on B x Sum of Amt, and on the whole Sum of Qty.
PivotTable Table() {
  PivotTable table;
  PivotField item;
  item.source_name = "Item";
  item.axis = PivotAxis::Row;
  item.items = {Item(0U), Item(1U), Item(2U)};
  item.sort.ascending = true;
  item.sort.stated = true;
  PivotField amt;
  amt.source_name = "Amt";
  amt.axis = PivotAxis::Value;
  PivotField qty;
  qty.source_name = "Qty";
  qty.axis = PivotAxis::Value;
  table.mutable_fields() = {item, amt, qty};
  table.mutable_row_field_order() = {kItemField};
  table.mutable_data_fields() = {DataField("Sum of Amt", 1U), DataField("Sum of Qty", 2U)};
  table.mutable_formats() = {
      Format(3U, {Ref(kItemField, {0U})}),
      Format(2U, {Ref(kItemField, {0U}), Ref(kValues, {0U})}),
      Format(1U, {Ref(kValues, {1U})}),
  };
  return table;
}

RefreshSource Source(std::vector<std::string> headers, const std::vector<std::vector<Value>>& rows) {
  RefreshSource source;
  source.columns = static_cast<std::uint32_t>(headers.size());
  source.rows = static_cast<std::uint32_t>(rows.size());
  source.headers = std::move(headers);
  source.num_fmt_ids.assign(source.columns, 0U);
  for (const std::vector<Value>& row : rows) {
    for (const Value& v : row) {
      source.cells.push_back(v);
      source.cell_is_date.push_back(false);
    }
  }
  return source;
}

std::vector<Value> Row(const char* item, double amt, double qty) {
  return {Value::text(item), Value::number(amt), Value::number(qty)};
}

// Refreshes `table` over `cache` from `source` and remaps it onto the new cache.
void Refresh(PivotTable& table, const PivotCache& cache, const RefreshSource& source) {
  const ExcelProfile profile = mac_365_ja_jp_profile();
  RefreshContext context;
  context.axis_fields = {true, false, false};
  context.blank_header_word = pivot::blank_header_field_word(profile);
  context.profile = profile;
  auto plan = pivot::build_refreshed_cache(cache, source, context);
  ASSERT_TRUE(static_cast<bool>(plan)) << plan.error().message;
  pivot::remap_table(table, plan.value().cache, plan.value().remap);
}

std::vector<std::uint32_t> Dxfs(const PivotTable& table) {
  std::vector<std::uint32_t> out;
  for (const PivotFormat& format : table.formats()) {
    out.push_back(format.dxf_id.value_or(UINT32_MAX));
  }
  return out;
}

const std::vector<std::uint32_t>& ItemsOf(const PivotTableArea& area, std::size_t ref) {
  return area.references.at(ref).items;
}

TEST(PivotFormatRemap, ItemSortedInFrontMovesTheFormatsToItsPosition) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  PivotTable table = Table();
  Refresh(table, OldCache(),
          Source({"Item", "Amt", "Qty"}, {Row("A", 1, 1), Row("B", 2, 2), Row("C", 3, 3), Row("D", 4, 4)}));
  ASSERT_EQ(Dxfs(table), U32s({3U, 2U, 1U}));
  EXPECT_EQ(ItemsOf(table.formats()[0].area, 0), U32s({1U}));
  EXPECT_EQ(ItemsOf(table.formats()[1].area, 0), U32s({1U}));
  EXPECT_EQ(ItemsOf(table.formats()[1].area, 1), U32s({0U}));
  EXPECT_EQ(ItemsOf(table.formats()[2].area, 0), U32s({1U}));
}

TEST(PivotFormatRemap, AreaFollowsTheItemPositionNotTheCacheIndex) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  PivotTable table = Table();
  // A comes last in the source, so it takes cache index 3 but sorts first.
  Refresh(table, OldCache(),
          Source({"Item", "Amt", "Qty"}, {Row("B", 2, 2), Row("C", 3, 3), Row("D", 4, 4), Row("A", 1, 1)}));
  ASSERT_EQ(table.fields()[kItemField].items.front().cache_index, 3U);
  ASSERT_EQ(Dxfs(table), U32s({3U, 2U, 1U}));
  EXPECT_EQ(ItemsOf(table.formats()[0].area, 0), U32s({1U}));
  EXPECT_EQ(ItemsOf(table.formats()[1].area, 0), U32s({1U}));
}

TEST(PivotFormatRemap, ItemPurgedUnderLimitZeroRemovesItsFormats) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  PivotCache cache = OldCache();
  cache.set_missing_items_limit(0U);
  PivotTable table = Table();
  Refresh(table, cache, Source({"Item", "Amt", "Qty"}, {Row("C", 3, 3), Row("D", 4, 4)}));
  ASSERT_EQ(Dxfs(table), U32s({1U}));
  EXPECT_EQ(ItemsOf(table.formats()[0].area, 0), U32s({1U}));
}

TEST(PivotFormatRemap, ItemKeptAsMissingKeepsItsFormats) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  PivotTable table = Table();
  Refresh(table, OldCache(), Source({"Item", "Amt", "Qty"}, {Row("C", 3, 3), Row("D", 4, 4)}));
  ASSERT_EQ(Dxfs(table), U32s({3U, 2U, 1U}));
  EXPECT_EQ(ItemsOf(table.formats()[0].area, 0), U32s({0U}));
  EXPECT_EQ(ItemsOf(table.formats()[1].area, 0), U32s({0U}));
}

TEST(PivotFormatRemap, DeletedFieldRemovesEveryFormatOnIt) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  PivotTable table = Table();
  Refresh(table, OldCache(), Source({"Amt", "Qty"}, {{Value::number(1), Value::number(2)}}));
  ASSERT_EQ(Dxfs(table), U32s({1U}));
  EXPECT_EQ(table.formats()[0].area.references[0].field, std::optional<std::uint32_t>(kValues));
  EXPECT_EQ(ItemsOf(table.formats()[0].area, 0), U32s({1U}));
}

TEST(PivotFormatRemap, DeletedDataFieldRemovesItsFormatsAndShiftsLaterOnes) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  PivotTable table = Table();
  Refresh(table, OldCache(), Source({"Item", "Qty"}, {{Value::text("B"), Value::number(2)}}));
  ASSERT_EQ(Dxfs(table), U32s({3U, 1U}));
  EXPECT_EQ(ItemsOf(table.formats()[0].area, 0), U32s({0U}));
  EXPECT_EQ(ItemsOf(table.formats()[1].area, 0), U32s({0U}));
}

TEST(PivotFormatRemap, DataFieldLeavingTheValuesAreaRemovesItsFormatsAndShiftsLaterOnes) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  PivotTable table = Table();
  // Amt leaves the values area; the field itself stays.
  pivot::drop_field_aggregates(table, 1U);
  ASSERT_EQ(table.data_fields().size(), 1U);
  ASSERT_EQ(table.formats().size(), 2U);
  EXPECT_EQ(table.formats()[0].dxf_id, std::optional<std::uint32_t>(3U));
  EXPECT_EQ(table.formats()[1].dxf_id, std::optional<std::uint32_t>(1U));
  EXPECT_EQ(ItemsOf(table.formats()[0].area, 0).at(0), 0U);
  EXPECT_EQ(ItemsOf(table.formats()[1].area, 0).at(0), 0U);
}

TEST(PivotFormatRemap, ClearingTheValuesAreaRemovesEveryDataFieldFormat) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  PivotTable table = Table();
  pivot::clear_data_fields(table);
  EXPECT_TRUE(table.data_fields().empty());
  ASSERT_EQ(table.formats().size(), 1U);
  EXPECT_EQ(table.formats()[0].dxf_id, std::optional<std::uint32_t>(3U));
  EXPECT_EQ(ItemsOf(table.formats()[0].area, 0).at(0), 0U);
}

// Series Amt, series Qty, and point B of series Amt.
PivotTable ChartTable() {
  PivotTable table = Table();
  table.mutable_formats().clear();
  table.mutable_chart_formats() = {
      ChartFormat(0U, true, {Ref(kValues, {0U})}),
      ChartFormat(1U, true, {Ref(kValues, {1U})}),
      ChartFormat(2U, false, {Ref(kValues, {0U}), Ref(kItemField, {0U})}),
  };
  return table;
}

std::vector<std::uint32_t> ChartIds(const PivotTable& table) {
  std::vector<std::uint32_t> out;
  for (const PivotChartFormat& format : table.chart_formats()) {
    out.push_back(format.format);
  }
  return out;
}

TEST(PivotFormatRemap, ChartPointOnAPurgedItemIsRemoved) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  PivotCache cache = OldCache();
  cache.set_missing_items_limit(0U);
  PivotTable table = ChartTable();
  Refresh(table, cache, Source({"Item", "Amt", "Qty"}, {Row("C", 3, 3), Row("D", 4, 4)}));
  EXPECT_EQ(ChartIds(table), U32s({0U, 1U}));
}

TEST(PivotFormatRemap, ChartPointOnAMissingItemStays) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  PivotTable table = ChartTable();
  Refresh(table, OldCache(), Source({"Item", "Amt", "Qty"}, {Row("C", 3, 3), Row("D", 4, 4)}));
  ASSERT_EQ(ChartIds(table), U32s({0U, 1U, 2U}));
  EXPECT_EQ(ItemsOf(table.chart_formats()[2].area, 1), U32s({0U}));
}

TEST(PivotFormatRemap, ChartPointOnADeletedFieldIsRemoved) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  PivotTable table = ChartTable();
  Refresh(table, OldCache(), Source({"Amt", "Qty"}, {{Value::number(1), Value::number(2)}}));
  EXPECT_EQ(ChartIds(table), U32s({0U, 1U}));
}

TEST(PivotFormatRemap, ChartFormatsOnADeletedDataFieldGoWithoutRenumbering) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  PivotTable table = ChartTable();
  Refresh(table, OldCache(), Source({"Item", "Qty"}, {{Value::text("B"), Value::number(2)}}));
  ASSERT_EQ(ChartIds(table), U32s({1U}));
  EXPECT_EQ(ItemsOf(table.chart_formats()[0].area, 0), U32s({0U}));
}

}  // namespace
}  // namespace formulon
