// Numeric pivot grouping: the group field it appends, its range items, the
// pivot fields and axes it rewires, ungrouping, and reading a date or numeric
// grouping back.

#include <cstddef>
#include <cstdint>
#include <initializer_list>
#include <memory>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/eval_profile_scope.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "pivot/field_grouping.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "sheet.h"
#include "utils/date_time.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon::pivot {
namespace {

constexpr std::size_t kRegion = 0U;
constexpr std::size_t kAmt = 1U;
constexpr std::size_t kQty = 2U;
constexpr std::size_t kGroup = 3U;

// Region (shared) / Amt (inline, `amounts`) / Qty (inline); each table puts
// Amt then Region on rows and sums Qty.
Workbook AmountWorkbook(const std::vector<Value>& amounts, std::size_t tables = 1U) {
  Workbook wb = Workbook::create_for_profile(mac_365_ja_jp_profile());
  auto cache = std::make_unique<PivotCache>();
  cache->set_cache_id(1U);
  cache->mutable_fields().push_back(PivotCacheField{"Region", {Value::text("North"), Value::text("South")}});
  cache->mutable_fields().push_back(PivotCacheField{"Amt", {}});
  cache->mutable_fields().push_back(PivotCacheField{"Qty", {}});
  for (std::size_t i = 0; i < amounts.size(); ++i) {
    PivotCacheRecord record;
    record.cells = {Value::number(static_cast<double>(i % 2U)), amounts[i], Value::number(1.0)};
    record.cell_is_index = {true, false, false};
    cache->mutable_records().push_back(std::move(record));
  }
  wb.add_pivot_cache(std::move(cache));
  for (std::size_t t = 0; t < tables; ++t) {
    auto table = std::make_unique<PivotTable>();
    table->set_name("PT" + std::to_string(t + 1U));
    table->set_pivot_cache_id(1U);
    for (const char* name : {"Region", "Amt", "Qty"}) {
      PivotField field;
      field.source_name = name;
      table->mutable_fields().push_back(std::move(field));
    }
    table->mutable_fields()[kAmt].axis = PivotAxis::Row;
    table->mutable_fields()[kRegion].axis = PivotAxis::Row;
    table->mutable_fields()[kQty].axis = PivotAxis::Value;
    PivotDataField sum;
    sum.name = "合計 / Qty";
    sum.field_index = static_cast<std::uint32_t>(kQty);
    table->mutable_data_fields().push_back(std::move(sum));
    table->mutable_row_field_order() = {static_cast<std::uint32_t>(kAmt), static_cast<std::uint32_t>(kRegion)};
    table->set_anchor(2U, 6U * static_cast<std::uint32_t>(t), 10U, 2U);
    wb.sheet(0).add_pivot_table(std::move(table));
  }
  return wb;
}

std::vector<Value> Numbers(std::initializer_list<double> values) {
  std::vector<Value> out;
  for (const double v : values) {
    out.push_back(Value::number(v));
  }
  return out;
}

const PivotCache& CacheOf(const Workbook& wb) {
  return *wb.pivot_caches().front();
}

std::vector<std::string> GroupTexts(const PivotCacheField& field) {
  std::vector<std::string> texts;
  for (const Value& item : field.field_group->group_items) {
    texts.emplace_back(item.is_text() ? item.as_text() : std::string_view("?"));
  }
  return texts;
}

std::vector<double> GroupNumbers(const PivotCache& cache) {
  std::vector<double> numbers;
  for (const PivotCacheRecord& record : cache.records()) {
    numbers.push_back(
        kGroup < record.cells.size() && record.cells[kGroup].is_number() ? record.cells[kGroup].as_number() : -1.0);
  }
  return numbers;
}

TEST(NumericGroup, AutoBoundsTakeTheDataAndLabelEachRange) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  Workbook wb = AmountWorkbook(Numbers({20, 55, 140, 300}), 2U);
  ASSERT_TRUE(static_cast<bool>(set_numeric_group(wb, 0U, 0U, kAmt, std::nullopt, std::nullopt, 100.0)));
  const PivotCache& cache = CacheOf(wb);
  ASSERT_EQ(cache.fields().size(), 4U);
  const PivotCacheField& group = cache.fields()[kGroup];
  EXPECT_EQ(group.name, "Amt2");
  EXPECT_FALSE(group.is_database_field);
  EXPECT_EQ(group.field_group->base.value_or(99U), kAmt);
  EXPECT_EQ(cache.fields()[kAmt].field_group->par.value_or(99U), kGroup);
  const RangeGroup& range = group.field_group->range;
  EXPECT_FALSE(range.is_date);
  EXPECT_EQ(range.by, GroupBy::Range);
  EXPECT_EQ(range.start_num, 20.0);
  EXPECT_EQ(range.end_num, 300.0);
  EXPECT_EQ(range.interval, 100.0);
  EXPECT_TRUE(range.auto_start);
  EXPECT_TRUE(range.auto_end);
  EXPECT_EQ(GroupTexts(group), (std::vector<std::string>{"<20", "20-119", "120-219", "220-319", ">320"}));
  EXPECT_EQ(GroupNumbers(cache), (std::vector<double>{1, 1, 2, 3}));

  const PivotTable& grouped = *wb.sheet(0).pivot_tables()[0];
  const PivotTable& sibling = *wb.sheet(0).pivot_tables()[1];
  for (const PivotTable* table : {&grouped, &sibling}) {
    ASSERT_EQ(table->fields().size(), 4U);
    const PivotField& field = table->fields()[kGroup];
    EXPECT_EQ(field.source_name, "Amt2");
    EXPECT_TRUE(field.sort.manual);
    ASSERT_EQ(field.items.size(), 5U);
    for (std::uint32_t x = 0; x < 5U; ++x) {
      EXPECT_EQ(field.items[x].cache_index, x);
    }
    EXPECT_TRUE(table->layout_dirty());
  }
  EXPECT_EQ(grouped.row_field_order(), (std::vector<std::uint32_t>{3U, 0U}));
  EXPECT_EQ(grouped.fields()[kAmt].axis, PivotAxis::None);
  EXPECT_EQ(sibling.row_field_order(), (std::vector<std::uint32_t>{3U, 1U, 0U}));
  EXPECT_EQ(sibling.fields()[kAmt].axis, PivotAxis::Row);
}

TEST(NumericGroup, NegativeStartLabelsCountFromIt) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  Workbook wb = AmountWorkbook(Numbers({-30, 400, 950}));
  ASSERT_TRUE(static_cast<bool>(set_numeric_group(wb, 0U, 0U, kAmt, std::nullopt, std::nullopt, 100.0)));
  EXPECT_EQ(GroupTexts(CacheOf(wb).fields()[kGroup]),
            (std::vector<std::string>{"<-30", "-30-69", "70-169", "170-269", "270-369", "370-469", "470-569", "570-669",
                                      "670-769", "770-869", "870-969", ">970"}));
  EXPECT_EQ(GroupNumbers(CacheOf(wb)), (std::vector<double>{1, 5, 10}));
}

TEST(NumericGroup, FixedBoundsPutOutliersInTheEndItems) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  Workbook wb = AmountWorkbook(Numbers({-5, 50, 150, 250}));
  ASSERT_TRUE(static_cast<bool>(set_numeric_group(wb, 0U, 0U, kAmt, 0.0, 200.0, 100.0)));
  const PivotCache& cache = CacheOf(wb);
  const RangeGroup& range = cache.fields()[kGroup].field_group->range;
  EXPECT_FALSE(range.auto_start);
  EXPECT_FALSE(range.auto_end);
  EXPECT_EQ(GroupTexts(cache.fields()[kGroup]), (std::vector<std::string>{"<0", "0-99", "100-199", "200-299", ">300"}));
  EXPECT_EQ(GroupNumbers(cache), (std::vector<double>{0, 1, 2, 4}));
}

TEST(NumericGroup, RegroupingReplacesTheGroupField) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  Workbook wb = AmountWorkbook(Numbers({20, 300}));
  ASSERT_TRUE(static_cast<bool>(set_numeric_group(wb, 0U, 0U, kAmt, std::nullopt, std::nullopt, 100.0)));
  // Through the group field, as on a group cell.
  ASSERT_TRUE(static_cast<bool>(set_numeric_group(wb, 0U, 0U, kGroup, std::nullopt, std::nullopt, 50.0)));
  const PivotCache& cache = CacheOf(wb);
  ASSERT_EQ(cache.fields().size(), 4U);
  EXPECT_EQ(cache.fields()[kGroup].name, "Amt2");
  EXPECT_EQ(cache.fields()[kGroup].field_group->group_items.size(), 8U);
  const PivotTable& table = *wb.sheet(0).pivot_tables()[0];
  EXPECT_EQ(table.row_field_order(), (std::vector<std::uint32_t>{3U, 0U}));
  EXPECT_EQ(table.fields()[kGroup].items.size(), 8U);
}

TEST(NumericGroup, ClearRestoresTheBaseWhereItsGroupWas) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  Workbook wb = AmountWorkbook(Numbers({20, 300}), 2U);
  ASSERT_TRUE(static_cast<bool>(set_numeric_group(wb, 0U, 0U, kAmt, std::nullopt, std::nullopt, 100.0)));
  ASSERT_TRUE(static_cast<bool>(clear_numeric_group(wb, 0U, 1U, kGroup)));
  const PivotCache& cache = CacheOf(wb);
  ASSERT_EQ(cache.fields().size(), 3U);
  EXPECT_FALSE(cache.fields()[kAmt].field_group.has_value());
  for (const auto& table : wb.sheet(0).pivot_tables()) {
    ASSERT_EQ(table->fields().size(), 3U);
    EXPECT_EQ(table->row_field_order(), (std::vector<std::uint32_t>{1U, 0U}));
    EXPECT_EQ(table->fields()[kAmt].axis, PivotAxis::Row);
  }
  // Clearing what is not grouped changes nothing; an index out of range fails.
  EXPECT_TRUE(static_cast<bool>(clear_numeric_group(wb, 0U, 0U, kAmt)));
  EXPECT_FALSE(static_cast<bool>(clear_numeric_group(wb, 0U, 0U, 9U)));
}

TEST(NumericGroup, NonIntegerSettingsAreNotImplementedAndChangeNothing) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  Workbook wb = AmountWorkbook(Numbers({20, 300}));
  ASSERT_TRUE(static_cast<bool>(set_numeric_group(wb, 0U, 0U, kAmt, std::nullopt, std::nullopt, 100.0)));
  const std::vector<std::string> before = GroupTexts(CacheOf(wb).fields()[kGroup]);
  const std::vector<Expected<void, Error>> refused = {
      set_numeric_group(wb, 0U, 0U, kAmt, std::nullopt, std::nullopt, 0.5),
      set_numeric_group(wb, 0U, 0U, kAmt, 0.5, std::nullopt, 100.0),
      set_numeric_group(wb, 0U, 0U, kAmt, std::nullopt, 299.5, 100.0),
  };
  for (const Expected<void, Error>& result : refused) {
    ASSERT_FALSE(static_cast<bool>(result));
    EXPECT_EQ(result.error().code, FormulonErrorCode::kNotImplemented);
  }
  ASSERT_EQ(CacheOf(wb).fields().size(), 4U);
  EXPECT_EQ(GroupTexts(CacheOf(wb).fields()[kGroup]), before);
  EXPECT_EQ(wb.sheet(0).pivot_tables()[0]->fields().size(), 4U);

  // Data that is not whole leaves an automatic bound non-integer.
  Workbook fractional = AmountWorkbook(Numbers({20.5, 300}));
  const Expected<void, Error> result = set_numeric_group(fractional, 0U, 0U, kAmt, std::nullopt, std::nullopt, 100.0);
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kNotImplemented);
  EXPECT_EQ(CacheOf(fractional).fields().size(), 3U);
}

TEST(NumericGroup, InvalidSettingsAndBasesAreRefused) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  Workbook wb = AmountWorkbook(Numbers({20, 300}));
  for (const Expected<void, Error>& result : {
           set_numeric_group(wb, 0U, 0U, kAmt, std::nullopt, std::nullopt, 0.0),
           set_numeric_group(wb, 0U, 0U, kAmt, std::nullopt, std::nullopt, -100.0),
           set_numeric_group(wb, 0U, 0U, kAmt, 300.0, 20.0, 100.0),
           set_numeric_group(wb, 0U, 0U, kRegion, std::nullopt, std::nullopt, 100.0),
       }) {
    ASSERT_FALSE(static_cast<bool>(result));
    EXPECT_EQ(result.error().code, FormulonErrorCode::kInvalidArgument);
  }
  EXPECT_EQ(set_numeric_group(wb, 1U, 0U, kAmt, std::nullopt, std::nullopt, 100.0).error().code,
            FormulonErrorCode::kSheetIndexOutOfRange);
  EXPECT_EQ(set_numeric_group(wb, 0U, 3U, kAmt, std::nullopt, std::nullopt, 100.0).error().code,
            FormulonErrorCode::kIndexOutOfRange);

  // A missing shared item does not count as a present value.
  Workbook stale = AmountWorkbook({});
  PivotCacheField& amt = (*stale.mutable_pivot_caches().front()).mutable_fields()[kAmt];
  amt.shared_items = {Value::number(10.0), Value::text("gone")};
  amt.shared_item_is_missing = {false, true};
  EXPECT_TRUE(static_cast<bool>(set_numeric_group(stale, 0U, 0U, kAmt, std::nullopt, std::nullopt, 100.0)));
}

TEST(NumericGroup, StateReadsTheGroupingBack) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  Workbook wb = AmountWorkbook(Numbers({20, 300}));
  EXPECT_FALSE(numeric_group_state(CacheOf(wb), kAmt).has_value());
  ASSERT_TRUE(static_cast<bool>(set_numeric_group(wb, 0U, 0U, kAmt, std::nullopt, 400.0, 100.0)));
  for (const std::size_t field : {kAmt, kGroup}) {
    const std::optional<NumericGroupState> state = numeric_group_state(CacheOf(wb), field);
    ASSERT_TRUE(state.has_value());
    EXPECT_EQ(state->start, 20.0);
    EXPECT_EQ(state->end, 400.0);
    EXPECT_EQ(state->interval, 100.0);
    EXPECT_TRUE(state->auto_start);
    EXPECT_FALSE(state->auto_end);
  }
  EXPECT_FALSE(numeric_group_state(CacheOf(wb), kRegion).has_value());
  EXPECT_FALSE(date_group_state(CacheOf(wb), kAmt, /*date1904=*/false).has_value());
}

double Day(int y, unsigned m, unsigned d) {
  return date_time::serial_from_ymd(y, m, d);
}

PivotCache DateCache() {
  PivotCache cache;
  cache.set_cache_id(1U);
  cache.mutable_fields().push_back(PivotCacheField{"Date", {}});
  for (const double serial : {Day(2024, 1U, 15U), Day(2024, 3U, 10U)}) {
    PivotCacheRecord record;
    record.cells = {Value::number(serial)};
    cache.mutable_records().push_back(std::move(record));
  }
  return cache;
}

TEST(NumericGroup, DateStateReadsTheDateGroupingBack) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  PivotCache cache = DateCache();
  EXPECT_FALSE(date_group_state(cache, 0U, /*date1904=*/false).has_value());
  DateGroupSpec spec;
  spec.periods = date_period_bit(GroupBy::Months) | date_period_bit(GroupBy::Years);
  spec.end = Day(2024, 12U, 31U);
  ASSERT_TRUE(static_cast<bool>(group_cache_by_date(cache, 0U, spec, /*date1904=*/false, mac_365_ja_jp_profile())));
  for (const std::size_t field : {0U, 1U, 2U}) {
    const std::optional<DateGroupState> state = date_group_state(cache, field, /*date1904=*/false);
    ASSERT_TRUE(state.has_value());
    EXPECT_EQ(state->periods, spec.periods);
    EXPECT_EQ(state->calendar, CalendarSystem::Gregorian);
    EXPECT_EQ(state->interval_days, 1U);
    EXPECT_EQ(state->start, Day(2024, 1U, 15U));
    EXPECT_EQ(state->end, Day(2024, 12U, 31U));
    EXPECT_TRUE(state->auto_start);
    EXPECT_FALSE(state->auto_end);
  }
  EXPECT_FALSE(numeric_group_state(cache, 0U).has_value());

  PivotCache days = DateCache();
  spec = DateGroupSpec{};
  spec.periods = date_period_bit(GroupBy::Days);
  spec.interval_days = 7U;
  ASSERT_TRUE(static_cast<bool>(group_cache_by_date(days, 0U, spec, /*date1904=*/false, mac_365_ja_jp_profile())));
  const std::optional<DateGroupState> state = date_group_state(days, 0U, /*date1904=*/false);
  ASSERT_TRUE(state.has_value());
  EXPECT_EQ(state->interval_days, 7U);
  EXPECT_EQ(state->end, Day(2024, 3U, 11U));
  EXPECT_TRUE(state->auto_end);
}

}  // namespace
}  // namespace formulon::pivot
