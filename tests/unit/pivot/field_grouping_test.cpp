// Group-item numbering of pivot group fields, its materialisation into cache
// records, and the date groups Excel creates.

#include "pivot/field_grouping.h"

#include <array>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <memory>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "excel_profile.h"
#include "gtest/gtest.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot/record_access.h"
#include "sheet.h"
#include "utils/date_time.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon::pivot {
namespace {

double Day(int y, unsigned m, unsigned d) {
  return date_time::serial_from_ymd(y, m, d);
}

Value DateValue(int y, unsigned m, unsigned d) {
  return Value::number(Day(y, m, d));
}

RangeGroup DateRange(GroupBy by, const char* start, const char* end, double interval = 1.0) {
  RangeGroup range;
  range.by = by;
  range.is_date = true;
  range.start_date = start;
  range.end_date = end;
  range.interval = interval;
  return range;
}

std::uint32_t Number(const RangeGroup& range, std::size_t item_count, const Value& value) {
  const std::optional<GroupBounds> bounds = stated_group_bounds(range, /*date1904=*/false);
  EXPECT_TRUE(bounds.has_value());
  return range_group_item_number(range, bounds.value_or(GroupBounds{}), item_count, value, /*date1904=*/false);
}

TEST(FieldGrouping, GroupDateSerialParsesDateAndDateTime) {
  EXPECT_EQ(group_date_serial("2024-01-15T00:00:00", false).value_or(-1.0), Day(2024, 1U, 15U));
  EXPECT_EQ(group_date_serial("2024-01-15", false).value_or(-1.0), Day(2024, 1U, 15U));
  EXPECT_EQ(group_date_serial("2024-01-15T12:00:00", false).value_or(-1.0), Day(2024, 1U, 15U) + 0.5);
  EXPECT_EQ(group_date_serial("2024-01-15T00:00:00", true).value_or(-1.0),
            date_time::serial_from_ymd(2024, 1U, 15U, true));
  EXPECT_FALSE(group_date_serial("", false).has_value());
  EXPECT_FALSE(group_date_serial("2024-13-01", false).has_value());
  EXPECT_FALSE(group_date_serial("2023-02-29", false).has_value());
  EXPECT_FALSE(group_date_serial("2024-01-15X", false).has_value());
}

TEST(FieldGrouping, StatedBoundsNeedBothDates) {
  EXPECT_FALSE(stated_group_bounds(DateRange(GroupBy::Months, "", "2024-07-23T00:00:00"), false).has_value());
  RangeGroup numeric;
  numeric.start_num = 5.0;
  numeric.end_num = 50.0;
  const std::optional<GroupBounds> bounds = stated_group_bounds(numeric, false);
  ASSERT_TRUE(bounds.has_value());
  EXPECT_EQ(bounds->start, 5.0);
  EXPECT_EQ(bounds->end, 50.0);
}

TEST(FieldGrouping, MonthsNumberByMonthWithOutOfRangeCatchAlls) {
  const RangeGroup range = DateRange(GroupBy::Months, "2024-01-15T00:00:00", "2024-07-23T00:00:00");
  EXPECT_EQ(Number(range, 14U, DateValue(2024, 1U, 14U)), 0U);
  EXPECT_EQ(Number(range, 14U, DateValue(2024, 1U, 15U)), 1U);
  EXPECT_EQ(Number(range, 14U, DateValue(2024, 3U, 10U)), 3U);
  EXPECT_EQ(Number(range, 14U, DateValue(2024, 7U, 22U)), 7U);
  // The end day itself is inside the range, a time on it too.
  EXPECT_EQ(Number(range, 14U, Value::number(Day(2024, 7U, 23U) + 0.75)), 7U);
  EXPECT_EQ(Number(range, 14U, DateValue(2024, 7U, 24U)), 13U);
}

TEST(FieldGrouping, QuartersAndYears) {
  const RangeGroup quarters = DateRange(GroupBy::Quarters, "2024-01-15T00:00:00", "2024-12-31T00:00:00");
  EXPECT_EQ(Number(quarters, 6U, DateValue(2024, 3U, 31U)), 1U);
  EXPECT_EQ(Number(quarters, 6U, DateValue(2024, 4U, 1U)), 2U);
  EXPECT_EQ(Number(quarters, 6U, DateValue(2024, 12U, 31U)), 4U);
  EXPECT_EQ(Number(quarters, 6U, DateValue(2025, 1U, 1U)), 5U);

  const RangeGroup years = DateRange(GroupBy::Years, "2023-06-01T00:00:00", "2025-06-01T00:00:00");
  EXPECT_EQ(Number(years, 5U, DateValue(2023, 5U, 31U)), 0U);
  EXPECT_EQ(Number(years, 5U, DateValue(2023, 6U, 1U)), 1U);
  EXPECT_EQ(Number(years, 5U, DateValue(2024, 2U, 29U)), 2U);
  EXPECT_EQ(Number(years, 5U, DateValue(2025, 6U, 1U)), 3U);
  EXPECT_EQ(Number(years, 5U, DateValue(2025, 6U, 2U)), 4U);
}

TEST(FieldGrouping, DaysNumberByDayOfALeapYear) {
  const RangeGroup days = DateRange(GroupBy::Days, "2023-01-01T00:00:00", "2025-01-01T00:00:00");
  EXPECT_EQ(Number(days, 368U, DateValue(2024, 1U, 15U)), 15U);
  EXPECT_EQ(Number(days, 368U, DateValue(2024, 2U, 29U)), 60U);
  EXPECT_EQ(Number(days, 368U, DateValue(2024, 3U, 1U)), 61U);
  // A year without 2/29 skips day 60.
  EXPECT_EQ(Number(days, 368U, DateValue(2023, 2U, 28U)), 59U);
  EXPECT_EQ(Number(days, 368U, DateValue(2023, 3U, 1U)), 61U);
  EXPECT_EQ(Number(days, 368U, DateValue(2023, 12U, 31U)), 366U);
  EXPECT_EQ(Number(days, 368U, DateValue(2024, 12U, 31U)), 366U);
  EXPECT_EQ(Number(days, 368U, DateValue(2025, 1U, 2U)), 367U);
}

TEST(FieldGrouping, DaysIntervalCountsFromTheStart) {
  const RangeGroup week = DateRange(GroupBy::Days, "2024-01-15T00:00:00", "2024-07-23T00:00:00", 7.0);
  EXPECT_EQ(Number(week, 30U, DateValue(2024, 1U, 15U)), 1U);
  EXPECT_EQ(Number(week, 30U, DateValue(2024, 1U, 25U)), 2U);
  EXPECT_EQ(Number(week, 30U, DateValue(2024, 2U, 8U)), 4U);
  EXPECT_EQ(Number(week, 30U, DateValue(2024, 2U, 20U)), 6U);
  EXPECT_EQ(Number(week, 30U, DateValue(2024, 3U, 10U)), 8U);
  // 190 days do not divide by 7: the last bucket is the short remainder.
  EXPECT_EQ(Number(week, 30U, DateValue(2024, 7U, 22U)), 28U);
  EXPECT_EQ(Number(week, 30U, DateValue(2024, 7U, 23U)), 28U);
  EXPECT_EQ(Number(week, 30U, DateValue(2024, 7U, 24U)), 29U);
}

TEST(FieldGrouping, DaysIntervalFloorsFractionalBounds) {
  RangeGroup range;
  range.by = GroupBy::Days;
  range.is_date = true;
  range.interval = 7.0;
  const GroupBounds bounds{1.9, 60.9};
  EXPECT_EQ(range_group_item_number(range, bounds, 11U, Value::number(0.0), false), 0U);
  EXPECT_EQ(range_group_item_number(range, bounds, 11U, Value::number(1.0), false), 1U);
  EXPECT_EQ(range_group_item_number(range, bounds, 11U, Value::number(7.5), false), 1U);
  EXPECT_EQ(range_group_item_number(range, bounds, 11U, Value::number(8.0), false), 2U);
  EXPECT_EQ(range_group_item_number(range, bounds, 11U, Value::number(60.5), false), 9U);
  EXPECT_EQ(range_group_item_number(range, bounds, 11U, Value::number(61.0), false), 10U);
}

TEST(FieldGrouping, InRangeNumberIsClampedBelowTheEndCatchAll) {
  // Too few items for the range: the in-range number never reaches the last item.
  const RangeGroup week = DateRange(GroupBy::Days, "2024-01-15T00:00:00", "2024-07-23T00:00:00", 7.0);
  EXPECT_EQ(Number(week, 4U, DateValue(2024, 3U, 10U)), 2U);
  EXPECT_EQ(Number(week, 0U, DateValue(2024, 3U, 10U)), 0U);
}

TEST(FieldGrouping, TimeUnitsNumberByTheirValuePlusOne) {
  const RangeGroup hours = DateRange(GroupBy::Hours, "2024-01-15T00:00:00", "2024-01-16T00:00:00");
  const Value at = Value::number(Day(2024, 1U, 15U) + (13.0 * 3600.0 + 45.0 * 60.0 + 30.0) / 86400.0);
  EXPECT_EQ(Number(hours, 26U, at), 14U);
  EXPECT_EQ(Number(DateRange(GroupBy::Minutes, "2024-01-15T00:00:00", "2024-01-16T00:00:00"), 62U, at), 46U);
  EXPECT_EQ(Number(DateRange(GroupBy::Seconds, "2024-01-15T00:00:00", "2024-01-16T00:00:00"), 62U, at), 31U);
  EXPECT_EQ(Number(hours, 26U, Value::number(Day(2024, 1U, 14U) + 0.5)), 0U);
  EXPECT_EQ(Number(hours, 26U, Value::number(Day(2024, 1U, 16U) + 0.5)), 25U);
}

TEST(FieldGrouping, BlankAndNonDateBaseValuesTakeTheFirstItem) {
  const RangeGroup range = DateRange(GroupBy::Months, "2024-01-15T00:00:00", "2024-07-23T00:00:00");
  EXPECT_EQ(Number(range, 14U, Value::blank()), 0U);
  EXPECT_EQ(Number(range, 14U, Value::text("2024-03-01")), 0U);
  EXPECT_EQ(Number(range, 14U, Value::boolean(true)), 0U);
  EXPECT_EQ(Number(range, 14U, Value::number(-1.0)), 0U);
  EXPECT_EQ(Number(range, 14U, Value::number(std::numeric_limits<double>::quiet_NaN())), 0U);
}

TEST(FieldGrouping, NumericRangeNumbersByInterval) {
  RangeGroup range;
  range.start_num = 0.0;
  range.end_num = 100.0;
  range.interval = 10.0;
  const GroupBounds bounds{0.0, 100.0};
  EXPECT_EQ(range_group_item_number(range, bounds, 13U, Value::number(-1.0), false), 0U);
  EXPECT_EQ(range_group_item_number(range, bounds, 13U, Value::number(0.0), false), 1U);
  EXPECT_EQ(range_group_item_number(range, bounds, 13U, Value::number(9.99), false), 1U);
  EXPECT_EQ(range_group_item_number(range, bounds, 13U, Value::number(10.0), false), 2U);
  EXPECT_EQ(range_group_item_number(range, bounds, 13U, Value::number(100.0), false), 11U);
  EXPECT_EQ(range_group_item_number(range, bounds, 13U, Value::number(100.5), false), 12U);
  EXPECT_EQ(range_group_item_number(range, bounds, 13U, Value::text("x"), false), 0U);
}

TEST(FieldGrouping, DiscreteMapsBaseItems) {
  FieldGroup group;
  group.base = 0U;
  group.kind = GroupKind::Discrete;
  group.discrete_map = {1U, 0U, 1U};
  EXPECT_EQ(discrete_group_item_number(group, 0U).value_or(99U), 1U);
  EXPECT_EQ(discrete_group_item_number(group, 1U).value_or(99U), 0U);
  EXPECT_FALSE(discrete_group_item_number(group, 3U).has_value());
}

PivotCacheField GroupField(const char* name, FieldGroup group) {
  PivotCacheField field{name, {}};
  field.is_database_field = false;
  field.field_group = std::move(group);
  return field;
}

TEST(FieldGrouping, MaterializeWritesRangeNumbersIntoRecords) {
  PivotCache cache;
  cache.mutable_fields().push_back(PivotCacheField{"Date", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amt", {}});
  FieldGroup months;
  months.base = 0U;
  months.kind = GroupKind::Range;
  months.range = DateRange(GroupBy::Months, "2024-01-15T00:00:00", "2024-07-23T00:00:00");
  months.group_items.assign(14U, Value::text("m"));
  cache.mutable_fields().push_back(GroupField("Months", months));
  for (const double serial : {Day(2024, 1U, 15U), Day(2024, 3U, 10U), Day(2024, 7U, 22U)}) {
    PivotCacheRecord record;
    // As the reader leaves them: a slot per field and the encoding flags.
    record.cells = {Value::number(serial), Value::number(1.0), Value::blank()};
    record.cell_is_index = {false, false, false};
    record.cell_is_date = {true, false, false};
    cache.mutable_records().push_back(std::move(record));
  }
  PivotCacheRecord blank;
  blank.cells = {Value::blank(), Value::number(1.0)};
  cache.mutable_records().push_back(std::move(blank));

  materialize_group_fields(cache, /*date1904=*/false);

  const std::vector<double> expected = {1.0, 3.0, 7.0, 0.0};
  for (std::size_t r = 0; r < expected.size(); ++r) {
    const PivotCacheRecord& record = cache.records()[r];
    ASSERT_EQ(record.cells.size(), 3U) << r;
    const Value v = cell_value(cache, record, 2U);
    ASSERT_TRUE(v.is_number()) << r;
    EXPECT_EQ(v.as_number(), expected[r]) << r;
  }
  EXPECT_FALSE(cache.records()[0].cell_is_index[2]);
  EXPECT_FALSE(cache.records()[0].cell_is_date[2]);
  EXPECT_TRUE(cache.records()[0].cell_is_date[0]);
  EXPECT_EQ(cell_value(cache, cache.records()[0], 0U).as_number(), Day(2024, 1U, 15U));
}

TEST(FieldGrouping, MaterializeResolvesUnstatedDateBoundsFromTheData) {
  PivotCache cache;
  cache.mutable_fields().push_back(PivotCacheField{"Date", {}});
  FieldGroup week;
  week.base = 0U;
  week.kind = GroupKind::Range;
  week.range = DateRange(GroupBy::Days, "", "", 7.0);
  week.group_items.assign(6U, Value::text("w"));
  cache.mutable_fields().push_back(GroupField("Week", week));
  for (const double serial : {Day(2024, 1U, 10U) + 0.5, Day(2024, 1U, 17U), Day(2024, 1U, 30U)}) {
    PivotCacheRecord record;
    record.cells = {Value::number(serial)};
    cache.mutable_records().push_back(std::move(record));
  }

  materialize_group_fields(cache, /*date1904=*/false);

  // Start is the first day (1/10); 1/17 opens the second week, 1/30 the third.
  EXPECT_EQ(cell_value(cache, cache.records()[0], 1U).as_number(), 1.0);
  EXPECT_EQ(cell_value(cache, cache.records()[1], 1U).as_number(), 2.0);
  EXPECT_EQ(cell_value(cache, cache.records()[2], 1U).as_number(), 3.0);
}

TEST(FieldGrouping, MaterializeHonoursTheDateEpoch) {
  PivotCache cache;
  cache.mutable_fields().push_back(PivotCacheField{"Date", {}});
  FieldGroup months;
  months.base = 0U;
  months.kind = GroupKind::Range;
  months.range = DateRange(GroupBy::Months, "2024-01-01T00:00:00", "2025-01-01T00:00:00");
  months.group_items.assign(14U, Value::text("m"));
  cache.mutable_fields().push_back(GroupField("Months", months));
  PivotCacheRecord record;
  record.cells = {Value::number(date_time::serial_from_ymd(2024, 5U, 31U, /*date1904=*/true))};
  cache.mutable_records().push_back(std::move(record));

  materialize_group_fields(cache, /*date1904=*/true);

  EXPECT_EQ(cell_value(cache, cache.records()[0], 1U).as_number(), 5.0);
}

TEST(FieldGrouping, MaterializeMapsDiscreteGroupsThroughSharedItems) {
  PivotCache cache;
  cache.mutable_fields().push_back(
      PivotCacheField{"Region", {Value::text("North"), Value::text("South"), Value::text("East")}});
  FieldGroup group;
  group.base = 0U;
  group.kind = GroupKind::Discrete;
  group.discrete_map = {0U, 0U, 1U};
  group.group_items = {Value::text("Group1"), Value::text("East")};
  cache.mutable_fields().push_back(GroupField("Region2", group));
  // An index cell, an inline value of the same field, and an unknown value.
  PivotCacheRecord indexed;
  indexed.cells = {Value::number(2.0)};
  indexed.cell_is_index = {true};
  cache.mutable_records().push_back(std::move(indexed));
  PivotCacheRecord inline_value;
  inline_value.cells = {Value::text("South")};
  inline_value.cell_is_index = {false};
  cache.mutable_records().push_back(std::move(inline_value));
  PivotCacheRecord unknown;
  unknown.cells = {Value::text("West")};
  unknown.cell_is_index = {false};
  cache.mutable_records().push_back(std::move(unknown));

  materialize_group_fields(cache, /*date1904=*/false);

  EXPECT_EQ(cell_value(cache, cache.records()[0], 1U).as_number(), 1.0);
  EXPECT_EQ(cell_value(cache, cache.records()[1], 1U).as_number(), 0.0);
  EXPECT_TRUE(cell_value(cache, cache.records()[2], 1U).is_blank());
  ASSERT_EQ(cache.records()[0].cell_is_index.size(), 2U);
  EXPECT_FALSE(cache.records()[0].cell_is_index[1]);
}

TEST(FieldGrouping, MaterializeLeavesDatabaseAndBaseFieldsAlone) {
  PivotCache cache;
  cache.mutable_fields().push_back(PivotCacheField{"Date", {}});
  // The base field's own `<fieldGroup par>` makes it no group field.
  FieldGroup par_only;
  par_only.par = 1U;
  cache.mutable_fields()[0].field_group = par_only;
  PivotCacheRecord record;
  record.cells = {Value::number(Day(2024, 1U, 15U))};
  cache.mutable_records().push_back(std::move(record));

  materialize_group_fields(cache, /*date1904=*/false);

  ASSERT_EQ(cache.records()[0].cells.size(), 1U);
  EXPECT_EQ(cache.records()[0].cells[0].as_number(), Day(2024, 1U, 15U));
  EXPECT_FALSE(is_group_field(cache.fields()[0]));
}

constexpr std::uint32_t kSeconds = date_period_bit(GroupBy::Seconds);
constexpr std::uint32_t kMinutes = date_period_bit(GroupBy::Minutes);
constexpr std::uint32_t kHours = date_period_bit(GroupBy::Hours);
constexpr std::uint32_t kDays = date_period_bit(GroupBy::Days);
constexpr std::uint32_t kMonths = date_period_bit(GroupBy::Months);
constexpr std::uint32_t kQuarters = date_period_bit(GroupBy::Quarters);
constexpr std::uint32_t kYears = date_period_bit(GroupBy::Years);

constexpr std::size_t kDateField = 2;

// The dates of the measured fixture's Date column, one per record.
std::vector<double> MeasuredDates() {
  return {Day(2024, 1U, 15U), Day(2024, 2U, 20U), Day(2024, 3U, 10U), Day(2024, 5U, 5U),
          Day(2024, 1U, 25U), Day(2024, 6U, 18U), Day(2024, 2U, 8U),  Day(2024, 7U, 22U)};
}

// The same records with a time of day on each date.
std::vector<double> MeasuredDateTimes() {
  const auto at = [](double day, int h, int m, int sec) { return day + (h * 3600.0 + m * 60.0 + sec) / 86400.0; };
  return {at(Day(2024, 1U, 15U), 9, 30, 15), at(Day(2024, 2U, 20U), 14, 5, 40), at(Day(2024, 3U, 10U), 23, 59, 59),
          at(Day(2024, 5U, 5U), 0, 0, 1),    at(Day(2024, 1U, 25U), 12, 0, 0),  at(Day(2024, 6U, 18U), 7, 45, 30),
          at(Day(2024, 2U, 8U), 18, 20, 5),  at(Day(2024, 7U, 22U), 3, 10, 50)};
}

// Region, Product, Date and Amt over `dates`. `as_saved` lays Date out as
// Excel saves it (dated shared items, their hints, index cells); otherwise
// the dates sit inline in the records, as a cache built in memory holds them.
PivotCache DateCache(const std::vector<double>& dates, bool as_saved) {
  PivotCache cache;
  cache.set_cache_id(1U);
  cache.mutable_fields().push_back(PivotCacheField{"Region", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Product", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Date", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amt", {}});
  PivotCacheField& date = cache.mutable_fields()[kDateField];
  if (as_saved) {
    for (const double serial : dates) {
      date.shared_items.push_back(Value::number(serial));
    }
    date.shared_item_is_date.assign(dates.size(), true);
    SharedItemsHints& h = date.shared_items_hints;
    h.present = true;
    h.has_contains_date = true;
    h.contains_date = true;
    h.has_contains_non_date = true;
    h.has_min_date = true;
    h.min_date = "2024-01-15T00:00:00";
    h.has_max_date = true;
    h.max_date = "2024-07-23T00:00:00";
  }
  for (std::size_t i = 0; i < dates.size(); ++i) {
    PivotCacheRecord record;
    record.cells = {Value::text("North"), Value::text("A"),
                    as_saved ? Value::number(static_cast<double>(i)) : Value::number(dates[i]),
                    Value::number(10.0 * static_cast<double>(i + 1U))};
    record.cell_is_index = {false, false, as_saved, false};
    cache.mutable_records().push_back(std::move(record));
  }
  return cache;
}

DateGroupSpec Periods(std::uint32_t periods) {
  DateGroupSpec spec;
  spec.periods = periods;
  return spec;
}

std::vector<std::string> ItemTexts(const PivotCacheField& field) {
  std::vector<std::string> texts;
  for (const Value& item : field.field_group->group_items) {
    texts.emplace_back(item.is_text() ? item.as_text() : std::string_view("?"));
  }
  return texts;
}

std::vector<std::string> Bracketed(std::string_view before, std::vector<std::string> inner, std::string_view after) {
  inner.insert(inner.begin(), std::string(before));
  inner.emplace_back(after);
  return inner;
}

const std::vector<std::string> kJaMonths = {"1月", "2月", "3月", "4月",  "5月",  "6月",
                                            "7月", "8月", "9月", "10月", "11月", "12月"};

TEST(FieldGrouping, MonthsQuartersYearsMatchExcelWithAndWithoutSavedHints) {
  for (const bool as_saved : {true, false}) {
    PivotCache cache = DateCache(MeasuredDates(), as_saved);
    const Expected<std::size_t, Error> first = group_cache_by_date(
        cache, kDateField, Periods(kMonths | kQuarters | kYears), /*date1904=*/false, mac_365_ja_jp_profile());
    ASSERT_TRUE(static_cast<bool>(first)) << first.error().message;
    EXPECT_EQ(first.value(), 4U);
    ASSERT_EQ(cache.fields().size(), 7U) << as_saved;
    EXPECT_EQ(cache.fields()[kDateField].field_group->par.value_or(99U), 6U);
    const char* names[] = {"月 (Date)", "四半期 (Date)", "年 (Date)"};
    const GroupBy bys[] = {GroupBy::Months, GroupBy::Quarters, GroupBy::Years};
    for (std::size_t k = 0; k < 3U; ++k) {
      const PivotCacheField& field = cache.fields()[4U + k];
      EXPECT_EQ(field.name, names[k]);
      EXPECT_FALSE(field.is_database_field);
      EXPECT_EQ(field.num_fmt_id.value_or(99U), 0U);
      ASSERT_TRUE(is_group_field(field));
      const FieldGroup& group = *field.field_group;
      EXPECT_EQ(*group.base, kDateField);
      EXPECT_FALSE(group.par.has_value());
      EXPECT_EQ(group.kind, GroupKind::Range);
      EXPECT_EQ(group.range.by, bys[k]);
      EXPECT_EQ(group.range.start_date, "2024-01-15T00:00:00") << as_saved;
      EXPECT_EQ(group.range.end_date, "2024-07-23T00:00:00") << as_saved;
      EXPECT_TRUE(group.range.auto_start);
      EXPECT_TRUE(group.range.auto_end);
      EXPECT_EQ(group.range.interval, 1.0);
    }
    EXPECT_EQ(ItemTexts(cache.fields()[4]), Bracketed("<2024/1/15", kJaMonths, ">2024/7/23"));
    EXPECT_EQ(ItemTexts(cache.fields()[5]),
              Bracketed("<2024/1/15", {"第 1 四半期", "第 2 四半期", "第 3 四半期", "第 4 四半期"}, ">2024/7/23"));
    EXPECT_EQ(ItemTexts(cache.fields()[6]), Bracketed("<2024/1/15", {"2024年"}, ">2024/7/23"));
    // 2024-03-10 is March, the first quarter and the first year.
    EXPECT_EQ(cell_value(cache, cache.records()[2], 4U).as_number(), 3.0);
    EXPECT_EQ(cell_value(cache, cache.records()[2], 5U).as_number(), 1.0);
    EXPECT_EQ(cell_value(cache, cache.records()[2], 6U).as_number(), 1.0);
    EXPECT_EQ(date_group_periods(cache, kDateField, false), kMonths | kQuarters | kYears);
    EXPECT_EQ(date_group_calendar(cache, kDateField, false), CalendarSystem::Gregorian);
    EXPECT_EQ(group_base_index(cache, 5U), kDateField);
    EXPECT_EQ(group_base_index(cache, 3U), 3U);
  }
}

TEST(FieldGrouping, DaysListEveryDayOfALeapYear) {
  PivotCache cache = DateCache(MeasuredDates(), /*as_saved=*/true);
  ASSERT_TRUE(static_cast<bool>(
      group_cache_by_date(cache, kDateField, Periods(kDays), /*date1904=*/false, mac_365_ja_jp_profile())));
  const std::vector<std::string> items = ItemTexts(cache.fields()[4]);
  EXPECT_EQ(cache.fields()[4].name, "日 (Date)");
  ASSERT_EQ(items.size(), 368U);
  EXPECT_EQ(items[0], "<2024/1/15");
  EXPECT_EQ(items[1], "1月1日");
  EXPECT_EQ(items[15], "1月15日");
  EXPECT_EQ(items[60], "2月29日");
  EXPECT_EQ(items[61], "3月1日");
  EXPECT_EQ(items[366], "12月31日");
  EXPECT_EQ(items[367], ">2024/7/23");
  EXPECT_EQ(cell_value(cache, cache.records()[0], 4U).as_number(), 15.0);
}

TEST(FieldGrouping, DaysIntervalListsRunsClippedToTheEnd) {
  PivotCache cache = DateCache(MeasuredDates(), /*as_saved=*/true);
  DateGroupSpec spec = Periods(kDays);
  spec.interval_days = 7U;
  ASSERT_TRUE(
      static_cast<bool>(group_cache_by_date(cache, kDateField, spec, /*date1904=*/false, mac_365_ja_jp_profile())));
  const FieldGroup& group = *cache.fields()[4].field_group;
  EXPECT_EQ(group.range.interval, 7.0);
  const std::vector<std::string> items = ItemTexts(cache.fields()[4]);
  ASSERT_EQ(items.size(), 30U);
  EXPECT_EQ(items[1], "2024/1/15 - 2024/1/21");
  EXPECT_EQ(items[3], "2024/1/29 - 2024/2/4");
  EXPECT_EQ(items[28], "2024/7/22 - 2024/7/23");
  EXPECT_EQ(items[29], ">2024/7/23");
  // 2024-02-08 falls in the fourth week.
  EXPECT_EQ(cell_value(cache, cache.records()[6], 4U).as_number(), 4.0);
}

TEST(FieldGrouping, TimeUnitsKeepTheDataTimesInTheirBounds) {
  PivotCache cache = DateCache(MeasuredDateTimes(), /*as_saved=*/false);
  ASSERT_TRUE(static_cast<bool>(group_cache_by_date(cache, kDateField, Periods(kSeconds | kMinutes | kHours),
                                                    /*date1904=*/false, mac_365_ja_jp_profile())));
  ASSERT_EQ(cache.fields().size(), 7U);
  EXPECT_EQ(cache.fields()[4].name, "秒 (Date)");
  EXPECT_EQ(cache.fields()[5].name, "分 (Date)");
  EXPECT_EQ(cache.fields()[6].name, "時間 (Date)");
  EXPECT_EQ(cache.fields()[kDateField].field_group->par.value_or(99U), 6U);
  for (std::size_t fi = 4; fi < 7U; ++fi) {
    EXPECT_EQ(cache.fields()[fi].field_group->range.start_date, "2024-01-15T09:30:15");
    EXPECT_EQ(cache.fields()[fi].field_group->range.end_date, "2024-07-22T03:10:50");
  }
  const std::vector<std::string> seconds = ItemTexts(cache.fields()[4]);
  const std::vector<std::string> hours = ItemTexts(cache.fields()[6]);
  ASSERT_EQ(seconds.size(), 62U);
  ASSERT_EQ(ItemTexts(cache.fields()[5]).size(), 62U);
  ASSERT_EQ(hours.size(), 26U);
  // The bounds show their date only.
  EXPECT_EQ(seconds.front(), "<2024/1/15");
  EXPECT_EQ(seconds.back(), ">2024/7/22");
  EXPECT_EQ(seconds[1], "0秒");
  EXPECT_EQ(hours[1], "0時");
  EXPECT_EQ(hours[24], "23時");
  // 9:30:15 is hour 9, minute 30, second 15.
  EXPECT_EQ(cell_value(cache, cache.records()[0], 4U).as_number(), 16.0);
  EXPECT_EQ(cell_value(cache, cache.records()[0], 5U).as_number(), 31.0);
  EXPECT_EQ(cell_value(cache, cache.records()[0], 6U).as_number(), 10.0);
}

struct LocaleDateWords {
  ExcelLocale locale;
  std::array<const char*, 12> months;
  std::array<const char*, 4> quarters;
  const char* year;
  std::array<const char*, 7> field_names;
  const char* before;
  const char* after;
  const char* after_time_data;
  // Jan 1, Jan 15, Feb 29, Mar 1, Dec 31.
  std::array<const char*, 5> days;
  const char* first_week;
  const char* last_week;
  std::array<const char*, 24> hours;
  // 0, 5, 59.
  std::array<const char*, 3> minutes;
  std::array<const char*, 3> seconds;
};

// Mac Excel 365's groups of the measured data in each UI locale.
const LocaleDateWords kLocaleDateWords[] = {
    {ExcelLocale::kJaJP,
     {"1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"},
     {"第 1 四半期", "第 2 四半期", "第 3 四半期", "第 4 四半期"},
     "2024年",
     {"秒 (Date)", "分 (Date)", "時間 (Date)", "日 (Date)", "月 (Date)", "四半期 (Date)", "年 (Date)"},
     "<2024/1/15",
     ">2024/7/23",
     ">2024/7/22",
     {"1月1日", "1月15日", "2月29日", "3月1日", "12月31日"},
     "2024/1/15 - 2024/1/21",
     "2024/7/22 - 2024/7/23",
     {"0時",  "1時",  "2時",  "3時",  "4時",  "5時",  "6時",  "7時",  "8時",  "9時",  "10時", "11時",
      "12時", "13時", "14時", "15時", "16時", "17時", "18時", "19時", "20時", "21時", "22時", "23時"},
     {"0分", "5分", "59分"},
     {"0秒", "5秒", "59秒"}},
    {ExcelLocale::kEnUS,
     {"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"},
     {"Qtr1", "Qtr2", "Qtr3", "Qtr4"},
     "2024",
     {"Seconds (Date)", "Minutes (Date)", "Hours (Date)", "Days (Date)", "Months (Date)", "Quarters (Date)",
      "Years (Date)"},
     "<1/15/24",
     ">7/23/24",
     ">7/22/24",
     {"1-Jan", "15-Jan", "29-Feb", "1-Mar", "31-Dec"},
     "1/15/24 - 1/21/24",
     "7/22/24 - 7/23/24",
     {"12 AM", "1 AM", "2 AM", "3 AM", "4 AM", "5 AM", "6 AM", "7 AM", "8 AM", "9 AM", "10 AM", "11 AM",
      "12 PM", "1 PM", "2 PM", "3 PM", "4 PM", "5 PM", "6 PM", "7 PM", "8 PM", "9 PM", "10 PM", "11 PM"},
     {":00", ":05", ":59"},
     {":00", ":05", ":59"}},
    {ExcelLocale::kDeDE,
     {"Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"},
     {"Qrtl1", "Qrtl2", "Qrtl3", "Qrtl4"},
     "2024",
     {"Sekunden (Date)", "Minuten (Date)", "Stunden (Date)", "Tage (Date)", "Monate (Date)", "Quartale (Date)",
      "Jahre (Date)"},
     "<15.01.24",
     ">23.07.24",
     ">22.07.24",
     {"01. Jan", "15. Jan", "29. Feb", "01. Mär", "31. Dez"},
     "15.01.24 - 21.01.24",
     "22.07.24 - 23.07.24",
     {"00", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11",
      "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23"},
     {":00", ":05", ":59"},
     {":00", ":05", ":59"}},
    {ExcelLocale::kFrFR,
     {"janv", "févr", "mars", "avr", "mai", "juin", "juil", "août", "sept", "oct", "nov", "déc"},
     {"Trimestre1", "Trimestre2", "Trimestre3", "Trimestre4"},
     "2024",
     {"Secondes (Date)", "Minutes (Date)", "Heures (Date)", "Jours (Date)", "Mois (Date)", "Trimestres (Date)",
      "Années (Date)"},
     "<15/01/2024",
     ">23/07/2024",
     ">22/07/2024",
     {"01-janv", "15-janv", "29-févr", "01-mars", "31-déc"},
     "15/01/2024 - 21/01/2024",
     "22/07/2024 - 23/07/2024",
     {"00", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11",
      "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23"},
     {":00", ":05", ":59"},
     {":00", ":05", ":59"}},
    {ExcelLocale::kZhCN,
     {"1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"},
     {"第一季", "第二季", "第三季", "第四季"},
     "2024年",
     {"秒(Date)", "分钟(Date)", "小时(Date)", "天(Date)", "月(Date)", "季度(Date)", "年(Date)"},
     "<2024/1/15",
     ">2024/7/23",
     ">2024/7/22",
     {"1月1日", "1月15日", "2月29日", "3月1日", "12月31日"},
     "2024/1/15 - 2024/1/21",
     "2024/7/22 - 2024/7/23",
     {"0时",  "1时",  "2时",  "3时",  "4时",  "5时",  "6时",  "7时",  "8时",  "9时",  "10时", "11时",
      "12时", "13时", "14时", "15时", "16时", "17时", "18时", "19时", "20时", "21时", "22时", "23时"},
     {"0分", "5分", "59分"},
     {"0秒", "5秒", "59秒"}},
    {ExcelLocale::kKoKR,
     {"1월", "2월", "3월", "4월", "5월", "6월", "7월", "8월", "9월", "10월", "11월", "12월"},
     {"1사분기", "2사분기", "3사분기", "4사분기"},
     "2024년",
     {"초(Date)", "분(Date)", "시간(Date)", "일(Date)", "개월(Date)", "분기(Date)", "년(Date)"},
     "<2024.1.15",
     ">2024.7.23",
     ">2024.7.22",
     {"1월1일", "1월15일", "2월29일", "3월1일", "12월31일"},
     "2024.1.15 - 2024.1.21",
     "2024.7.22 - 2024.7.23",
     {"0시",  "1시",  "2시",  "3시",  "4시",  "5시",  "6시",  "7시",  "8시",  "9시",  "10시", "11시",
      "12시", "13시", "14시", "15시", "16시", "17시", "18시", "19시", "20시", "21시", "22시", "23시"},
     {"0분", "5분", "59분"},
     {"0초", "5초", "59초"}},
    {ExcelLocale::kThTH,
     {"ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."},
     {"ไตรมาส1", "ไตรมาส2", "ไตรมาส3", "ไตรมาส4"},
     "2024",
     {"วินาที (Date)", "นาที (Date)", "ชั่วโมง (Date)", "วัน (Date)", "เดือน (Date)", "ไตรมาส (Date)", "ปี (Date)"},
     "<15/1/24",
     ">23/7/24",
     ">22/7/24",
     {"1-ม.ค.", "15-ม.ค.", "29-ก.พ.", "1-มี.ค.", "31-ธ.ค."},
     "15/1/24 - 21/1/24",
     "22/7/24 - 23/7/24",
     {"00", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11",
      "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23"},
     {":00", ":05", ":59"},
     {":00", ":05", ":59"}},
    {ExcelLocale::kEsES,
     {"ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sept", "oct", "nov", "dic"},
     {"Trim.1", "Trim.2", "Trim.3", "Trim.4"},
     "2024",
     {"Segundos (Date)", "Minutos (Date)", "Horas (Date)", "Días (Date)", "Meses (Date)", "Trimestres (Date)",
      "Años (Date)"},
     "<15/1/24",
     ">23/7/24",
     ">22/7/24",
     {"1-ene", "15-ene", "29-feb", "1-mar", "31-dic"},
     "15/1/24 - 21/1/24",
     "22/7/24 - 23/7/24",
     {"0",  "1",  "2",  "3",  "4",  "5",  "6",  "7",  "8",  "9",  "10", "11",
      "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23"},
     {":00", ":05", ":59"},
     {":00", ":05", ":59"}},
    {ExcelLocale::kEsMX,
     {"ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"},
     {"Trim.1", "Trim.2", "Trim.3", "Trim.4"},
     "2024",
     {"Segundos (Date)", "Minutos (Date)", "Horas (Date)", "Días (Date)", "Meses (Date)", "Trimestres (Date)",
      "Años (Date)"},
     "<15/01/24",
     ">23/07/24",
     ">22/07/24",
     {"01-ene", "15-ene", "29-feb", "01-mar", "31-dic"},
     "15/01/24 - 21/01/24",
     "22/07/24 - 23/07/24",
     {"12 a.m.", "1 a.m.", "2 a.m.",  "3 a.m.",  "4 a.m.",  "5 a.m.", "6 a.m.",  "7 a.m.",
      "8 a.m.",  "9 a.m.", "10 a.m.", "11 a.m.", "12 p.m.", "1 p.m.", "2 p.m.",  "3 p.m.",
      "4 p.m.",  "5 p.m.", "6 p.m.",  "7 p.m.",  "8 p.m.",  "9 p.m.", "10 p.m.", "11 p.m."},
     {":00", ":05", ":59"},
     {":00", ":05", ":59"}},
    {ExcelLocale::kPtBR,
     {"jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"},
     {"Trim1", "Trim2", "Trim3", "Trim4"},
     "2024",
     {"Segundos (Date)", "Minutos (Date)", "Horas (Date)", "Dias (Date)", "Meses (Date)", "Trimestres (Date)",
      "Anos (Date)"},
     "<15/01/2024",
     ">23/07/2024",
     ">22/07/2024",
     {"01/jan", "15/jan", "29/fev", "01/mar", "31/dez"},
     "15/01/2024 - 21/01/2024",
     "22/07/2024 - 23/07/2024",
     {"00", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11",
      "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23"},
     {":00", ":05", ":59"},
     {":00", ":05", ":59"}},
    {ExcelLocale::kRuRU,
     {"янв", "февр", "март", "апр", "май", "июнь", "июль", "авг", "сент", "окт", "нояб", "дек"},
     {"Кв-л1", "Кв-л2", "Кв-л3", "Кв-л4"},
     "2024",
     {"Секунды (Date)", "Минуты (Date)", "Часы (Date)", "Дни (Date)", "Месяцы (Date)", "Кварталы (Date)",
      "Годы (Date)"},
     "<15.01.2024",
     ">23.07.2024",
     ">22.07.2024",
     {"01.янв", "15.янв", "29.февр", "01.март", "31.дек"},
     "15.01.2024 - 21.01.2024",
     "22.07.2024 - 23.07.2024",
     {"00", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11",
      "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23"},
     {":00", ":05", ":59"},
     {":00", ":05", ":59"}},
    {ExcelLocale::kZhTW,
     {"1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"},
     {"第一季", "第二季", "第三季", "第四季"},
     "2024年",
     {"秒 (Date)", "分鐘 (Date)", "小時 (Date)", "天 (Date)", "月 (Date)", "季 (Date)", "年 (Date)"},
     "<2024/1/15",
     ">2024/7/23",
     ">2024/7/22",
     {"1月1日", "1月15日", "2月29日", "3月1日", "12月31日"},
     "2024/1/15 - 2024/1/21",
     "2024/7/22 - 2024/7/23",
     {"0時",  "1時",  "2時",  "3時",  "4時",  "5時",  "6時",  "7時",  "8時",  "9時",  "10時", "11時",
      "12時", "13時", "14時", "15時", "16時", "17時", "18時", "19時", "20時", "21時", "22時", "23時"},
     {"0分", "5分", "59分"},
     {"0秒", "5秒", "59秒"}},
    {ExcelLocale::kItIT,
     {"gen", "feb", "mar", "apr", "mag", "giu", "lug", "ago", "set", "ott", "nov", "dic"},
     {"Trim1", "Trim2", "Trim3", "Trim4"},
     "2024",
     {"Secondi (Date)", "Minuti (Date)", "Ore (Date)", "Giorni (Date)", "Mesi (Date)", "Trimestri (Date)",
      "Anni (Date)"},
     "<15/01/24",
     ">23/07/24",
     ">22/07/24",
     {"01-gen", "15-gen", "29-feb", "01-mar", "31-dic"},
     "15/01/24 - 21/01/24",
     "22/07/24 - 23/07/24",
     {"00", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11",
      "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23"},
     {":00", ":05", ":59"},
     {":00", ":05", ":59"}},
    {ExcelLocale::kNlNL,
     {"jan", "feb", "mrt", "apr", "mei", "jun", "jul", "aug", "sep", "okt", "nov", "dec"},
     {"Kwrt1", "Kwrt2", "Kwrt3", "Kwrt4"},
     "2024",
     {"Seconden (Date)", "Minuten (Date)", "Uren (Date)", "Dagen (Date)", "Maanden (Date)", "Kwartalen (Date)",
      "Jaren (Date)"},
     "<15-01-2024",
     ">23-07-2024",
     ">22-07-2024",
     {"01-jan", "15-jan", "29-feb", "01-mrt", "31-dec"},
     "15-01-2024 - 21-01-2024",
     "22-07-2024 - 23-07-2024",
     {"00", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11",
      "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23"},
     {":00", ":05", ":59"},
     {":00", ":05", ":59"}},

};

TEST(FieldGrouping, GroupItemsAndNamesMatchExcelInEveryLocale) {
  for (const LocaleDateWords& want : kLocaleDateWords) {
    const ExcelProfile profile{ExcelHost::kMac365, want.locale};
    const int locale = static_cast<int>(want.locale);
    PivotCache dates = DateCache(MeasuredDates(), /*as_saved=*/true);
    ASSERT_TRUE(static_cast<bool>(group_cache_by_date(
        dates, kDateField, Periods(kAllDatePeriods & ~(kSeconds | kMinutes | kHours)), false, profile)))
        << locale;
    // Days, months, quarters, years.
    for (std::size_t k = 0; k < 4U; ++k) {
      EXPECT_EQ(dates.fields()[4U + k].name, want.field_names[3U + k]) << locale;
    }
    const std::vector<std::string> days = ItemTexts(dates.fields()[4]);
    ASSERT_EQ(days.size(), 368U) << locale;
    EXPECT_EQ(days.front(), want.before) << locale;
    EXPECT_EQ(days.back(), want.after) << locale;
    const std::size_t day_index[] = {1U, 15U, 60U, 61U, 366U};
    for (std::size_t k = 0; k < 5U; ++k) {
      EXPECT_EQ(days[day_index[k]], want.days[k]) << locale << " day " << k;
    }
    const std::vector<std::string> months = ItemTexts(dates.fields()[5]);
    ASSERT_EQ(months.size(), 14U) << locale;
    for (std::size_t m = 0; m < 12U; ++m) {
      EXPECT_EQ(months[m + 1U], want.months[m]) << locale << " month " << m;
    }
    const std::vector<std::string> quarters = ItemTexts(dates.fields()[6]);
    ASSERT_EQ(quarters.size(), 6U) << locale;
    for (std::size_t q = 0; q < 4U; ++q) {
      EXPECT_EQ(quarters[q + 1U], want.quarters[q]) << locale << " quarter " << q;
    }
    EXPECT_EQ(ItemTexts(dates.fields()[7]), Bracketed(want.before, {want.year}, want.after)) << locale;

    PivotCache weeks = DateCache(MeasuredDates(), /*as_saved=*/true);
    DateGroupSpec by_week = Periods(kDays);
    by_week.interval_days = 7U;
    ASSERT_TRUE(static_cast<bool>(group_cache_by_date(weeks, kDateField, by_week, false, profile))) << locale;
    const std::vector<std::string> runs = ItemTexts(weeks.fields()[4]);
    ASSERT_EQ(runs.size(), 30U) << locale;
    EXPECT_EQ(runs[1], want.first_week) << locale;
    EXPECT_EQ(runs[28], want.last_week) << locale;

    PivotCache times = DateCache(MeasuredDateTimes(), /*as_saved=*/false);
    ASSERT_TRUE(static_cast<bool>(
        group_cache_by_date(times, kDateField, Periods(kSeconds | kMinutes | kHours), false, profile)))
        << locale;
    for (std::size_t k = 0; k < 3U; ++k) {
      EXPECT_EQ(times.fields()[4U + k].name, want.field_names[k]) << locale;
    }
    const std::vector<std::string> seconds = ItemTexts(times.fields()[4]);
    const std::vector<std::string> minutes = ItemTexts(times.fields()[5]);
    const std::vector<std::string> hours = ItemTexts(times.fields()[6]);
    ASSERT_EQ(hours.size(), 26U) << locale;
    EXPECT_EQ(hours.back(), want.after_time_data) << locale;
    for (std::size_t h = 0; h < 24U; ++h) {
      EXPECT_EQ(hours[h + 1U], want.hours[h]) << locale << " hour " << h;
    }
    const std::size_t unit_index[] = {0U, 5U, 59U};
    for (std::size_t k = 0; k < 3U; ++k) {
      EXPECT_EQ(minutes[unit_index[k] + 1U], want.minutes[k]) << locale;
      EXPECT_EQ(seconds[unit_index[k] + 1U], want.seconds[k]) << locale;
    }
  }
}

TEST(FieldGrouping, AuthoringRefusesWhatExcelCannotGroup) {
  const auto code = [](PivotCache& cache, const DateGroupSpec& spec) {
    const Expected<std::size_t, Error> r =
        group_cache_by_date(cache, kDateField, spec, /*date1904=*/false, mac_365_ja_jp_profile());
    return r ? FormulonErrorCode::kOk : r.error().code;
  };
  const FormulonErrorCode invalid = FormulonErrorCode::kInvalidArgument;
  PivotCache cache = DateCache(MeasuredDates(), /*as_saved=*/false);
  EXPECT_EQ(code(cache, Periods(0U)), invalid);
  EXPECT_EQ(code(cache, Periods(kYears | 128U)), invalid);
  DateGroupSpec spec = Periods(kDays);
  spec.interval_days = 0U;
  EXPECT_EQ(code(cache, spec), invalid);
  spec = Periods(kDays | kMonths);
  spec.interval_days = 7U;
  EXPECT_EQ(code(cache, spec), invalid);
  spec = Periods(kMonths);
  spec.start = Day(2024, 3U, 1U);
  spec.end = Day(2024, 2U, 1U);
  EXPECT_EQ(code(cache, spec), invalid);
  spec.start = std::numeric_limits<double>::infinity();
  spec.end.reset();
  EXPECT_EQ(code(cache, spec), invalid);
  // An explicit start past the data's end.
  spec.start = Day(2025, 1U, 1U);
  EXPECT_EQ(code(cache, spec), invalid);
  EXPECT_EQ(cache.fields().size(), 4U);

  // A blank, a text and a saved number are no dates; no records, no dates.
  for (const Value& bad : {Value::blank(), Value::text("n/a")}) {
    PivotCache with_bad = DateCache(MeasuredDates(), /*as_saved=*/false);
    with_bad.mutable_records()[3].cells[kDateField] = bad;
    EXPECT_EQ(code(with_bad, Periods(kMonths)), invalid);
  }
  PivotCache saved_number = DateCache(MeasuredDates(), /*as_saved=*/true);
  saved_number.mutable_fields()[kDateField].shared_item_is_date[4] = false;
  EXPECT_EQ(code(saved_number, Periods(kMonths)), invalid);
  PivotCache empty = DateCache({}, /*as_saved=*/false);
  EXPECT_EQ(code(empty, Periods(kMonths)), invalid);
  // A group field is no base to group.
  PivotCache grouped = DateCache(MeasuredDates(), /*as_saved=*/false);
  ASSERT_EQ(code(grouped, Periods(kMonths)), FormulonErrorCode::kOk);
  EXPECT_EQ(code(grouped, Periods(kYears)), FormulonErrorCode::kPreconditionFailed);
  EXPECT_EQ(group_cache_by_date(grouped, 4U, Periods(kYears), false, mac_365_ja_jp_profile()).error().code, invalid);
}

TEST(FieldGrouping, ExplicitBoundsAreKeptAndMarkedAsSet) {
  PivotCache cache = DateCache(MeasuredDates(), /*as_saved=*/true);
  DateGroupSpec spec = Periods(kMonths);
  spec.start = Day(2024, 2U, 1U);
  spec.end = Day(2024, 3U, 31U) + 0.5;
  ASSERT_TRUE(
      static_cast<bool>(group_cache_by_date(cache, kDateField, spec, /*date1904=*/false, mac_365_ja_jp_profile())));
  const RangeGroup& range = cache.fields()[4].field_group->range;
  EXPECT_FALSE(range.auto_start);
  EXPECT_FALSE(range.auto_end);
  EXPECT_EQ(range.start_date, "2024-02-01T00:00:00");
  EXPECT_EQ(range.end_date, "2024-03-31T12:00:00");
  EXPECT_EQ(ItemTexts(cache.fields()[4]), Bracketed("<2024/2/1", kJaMonths, ">2024/3/31"));
  // 1/15 falls before the start, 5/5 after the end.
  EXPECT_EQ(cell_value(cache, cache.records()[0], 4U).as_number(), 0.0);
  EXPECT_EQ(cell_value(cache, cache.records()[3], 4U).as_number(), 13.0);
}

TEST(FieldGrouping, AutoBoundsKeepAMinimumAndATimedMaximum) {
  const GroupBounds date_only = auto_date_group_bounds(Day(2024, 1U, 15U) + 0.25, Day(2024, 7U, 22U), false);
  EXPECT_EQ(date_only.start, Day(2024, 1U, 15U) + 0.25);
  EXPECT_EQ(date_only.end, Day(2024, 7U, 23U));
  const GroupBounds timed = auto_date_group_bounds(Day(2024, 1U, 15U), Day(2024, 7U, 22U) + 0.5, false);
  EXPECT_EQ(timed.end, Day(2024, 7U, 22U) + 0.5);
  const double last = date_time::serial_from_ymd(9999, 12U, 31U);
  EXPECT_EQ(auto_date_group_bounds(last, last, false).end, last);
  EXPECT_EQ(group_date_iso(Day(2024, 1U, 15U) + (9 * 3600.0 + 30 * 60.0 + 15.0) / 86400.0, false),
            "2024-01-15T09:30:15");
  EXPECT_EQ(group_date_iso(date_time::serial_from_ymd(2024, 1U, 15U, true), true), "2024-01-15T00:00:00");
}

TEST(FieldGrouping, JapaneseCalendarGroupsYearsByEraYear) {
  const std::vector<double> dates = {Day(2018, 12U, 31U), Day(2019, 4U, 30U), Day(2019, 5U, 1U),
                                     Day(2019, 12U, 31U), Day(2024, 3U, 15U), Day(2019, 4U, 30U)};
  PivotCache cache = DateCache(dates, /*as_saved=*/false);
  DateGroupSpec spec = Periods(kMonths | kYears);
  spec.calendar = CalendarSystem::Japanese;
  ASSERT_TRUE(
      static_cast<bool>(group_cache_by_date(cache, kDateField, spec, /*date1904=*/false, mac_365_ja_jp_profile())));
  ASSERT_EQ(cache.fields().size(), 6U);
  EXPECT_EQ(cache.fields()[4].field_group->kind, GroupKind::Range);
  const PivotCacheField& years = cache.fields()[5];
  EXPECT_EQ(years.name, "年 (Date)");
  ASSERT_EQ(years.field_group->kind, GroupKind::Discrete);
  EXPECT_EQ(*years.field_group->base, kDateField);
  EXPECT_EQ(cache.fields()[kDateField].field_group->par.value_or(99U), 5U);
  EXPECT_EQ(ItemTexts(years), (std::vector<std::string>{"平成30年", "平成31年", "令和1年", "令和6年"}));
  // The base now names its distinct dates as shared items.
  const PivotCacheField& base = cache.fields()[kDateField];
  ASSERT_EQ(base.shared_items.size(), 5U);
  EXPECT_EQ(years.field_group->discrete_map, (std::vector<std::uint32_t>{0U, 1U, 2U, 2U, 3U}));
  EXPECT_TRUE(base.shared_items_hints.present);
  EXPECT_EQ(base.shared_items_hints.min_date, "2018-12-31T00:00:00");
  EXPECT_EQ(base.shared_items_hints.max_date, "2024-03-16T00:00:00");
  const std::vector<double> want_era = {0.0, 1.0, 2.0, 2.0, 3.0, 1.0};
  for (std::size_t r = 0; r < dates.size(); ++r) {
    EXPECT_EQ(cell_value(cache, cache.records()[r], kDateField).as_number(), dates[r]) << r;
    EXPECT_EQ(cell_value(cache, cache.records()[r], 5U).as_number(), want_era[r]) << r;
    EXPECT_EQ(cell_value(cache, cache.records()[r], 3U).as_number(), 10.0 * static_cast<double>(r + 1U)) << r;
  }
  EXPECT_EQ(date_group_periods(cache, kDateField, false), kMonths | kYears);
  EXPECT_EQ(date_group_calendar(cache, kDateField, false), CalendarSystem::Japanese);
}

TEST(FieldGrouping, UngroupRemovesOnlyTrailingGroupFields) {
  PivotCache cache = DateCache(MeasuredDates(), /*as_saved=*/true);
  ASSERT_TRUE(static_cast<bool>(
      group_cache_by_date(cache, kDateField, Periods(kMonths | kYears), false, mac_365_ja_jp_profile())));
  cache.mutable_fields().push_back(PivotCacheField{"Extra", {}});
  const Expected<std::size_t, Error> refused = ungroup_cache_field(cache, kDateField);
  ASSERT_FALSE(static_cast<bool>(refused));
  EXPECT_EQ(refused.error().code, FormulonErrorCode::kPreconditionFailed);
  cache.mutable_fields().pop_back();

  const Expected<std::size_t, Error> removed = ungroup_cache_field(cache, kDateField);
  ASSERT_TRUE(static_cast<bool>(removed));
  EXPECT_EQ(removed.value(), 2U);
  ASSERT_EQ(cache.fields().size(), 4U);
  EXPECT_FALSE(cache.fields()[kDateField].field_group.has_value());
  EXPECT_EQ(cache.records()[0].cells.size(), 4U);
  EXPECT_EQ(ungroup_cache_field(cache, kDateField).value(), 0U);
}

// A ja-JP workbook whose cache holds the measured dates, with `tables`
// pivots that each show Date on rows and sum Amt.
Workbook DateWorkbook(std::size_t tables) {
  Workbook wb = Workbook::create_for_profile(mac_365_ja_jp_profile());
  wb.add_pivot_cache(std::make_unique<PivotCache>(DateCache(MeasuredDates(), /*as_saved=*/true)));
  for (std::size_t t = 0; t < tables; ++t) {
    auto table = std::make_unique<PivotTable>();
    table->set_name("PT" + std::to_string(t + 1U));
    table->set_pivot_cache_id(1U);
    for (const char* name : {"Region", "Product", "Date", "Amt"}) {
      PivotField field;
      field.source_name = name;
      field.axis = PivotAxis::None;
      table->mutable_fields().push_back(std::move(field));
    }
    PivotField& date = table->mutable_fields()[kDateField];
    date.axis = PivotAxis::Row;
    date.number_format = "14";
    for (std::uint32_t x = 0; x < 8U; ++x) {
      PivotItem item;
      item.has_cache_index = true;
      item.cache_index = x;
      date.items.push_back(item);
    }
    table->mutable_fields()[3].axis = PivotAxis::Value;
    PivotDataField sum;
    sum.name = "合計 / Amt";
    sum.field_index = 3U;
    table->mutable_data_fields().push_back(std::move(sum));
    table->mutable_row_field_order() = {static_cast<std::uint32_t>(kDateField)};
    table->set_anchor(2U, 6U * static_cast<std::uint32_t>(t), 10U, 2U);
    wb.sheet(0).add_pivot_table(std::move(table));
  }
  return wb;
}

TEST(FieldGrouping, SetDateGroupReplacesTheBaseHereAndPrecedesItOnSiblings) {
  Workbook wb = DateWorkbook(2U);
  ASSERT_TRUE(static_cast<bool>(set_date_group(wb, 0U, 0U, kDateField, kMonths | kYears, CalendarSystem::Gregorian, 1U,
                                               std::nullopt, std::nullopt)));
  const PivotTable& grouped = *wb.sheet(0).pivot_tables()[0];
  const PivotTable& sibling = *wb.sheet(0).pivot_tables()[1];
  for (const PivotTable* table : {&grouped, &sibling}) {
    ASSERT_EQ(table->fields().size(), 6U);
    EXPECT_EQ(table->fields()[4].source_name, "月 (Date)");
    EXPECT_EQ(table->fields()[4].items.size(), 14U);
    EXPECT_EQ(table->fields()[5].items.size(), 3U);
    EXPECT_EQ(table->fields()[5].items[1].name, "2024年");
    EXPECT_TRUE(table->fields()[5].sort.manual);
    EXPECT_TRUE(table->layout_dirty());
    // The base keeps its items and format.
    EXPECT_EQ(table->fields()[kDateField].items.size(), 8U);
    EXPECT_EQ(table->fields()[kDateField].number_format, "14");
  }
  EXPECT_EQ(grouped.row_field_order(), (std::vector<std::uint32_t>{5U, 4U}));
  EXPECT_EQ(grouped.fields()[kDateField].axis, PivotAxis::None);
  EXPECT_EQ(sibling.row_field_order(), (std::vector<std::uint32_t>{5U, 4U, 2U}));
  EXPECT_EQ(sibling.fields()[kDateField].axis, PivotAxis::Row);
  EXPECT_EQ(sibling.fields()[4].axis, PivotAxis::Row);
}

TEST(FieldGrouping, SetDateGroupKeepsTheValuesPosition) {
  Workbook wb = DateWorkbook(1U);
  PivotTable& table = *wb.sheet(0).mutable_pivot_tables()[0];
  table.set_row_values_position(1U);
  ASSERT_TRUE(static_cast<bool>(set_date_group(wb, 0U, 0U, kDateField, kMonths | kYears, CalendarSystem::Gregorian, 1U,
                                               std::nullopt, std::nullopt)));
  EXPECT_EQ(table.row_field_order(), (std::vector<std::uint32_t>{5U, 4U}));
  EXPECT_EQ(table.row_values_position().value_or(99U), 2U);
}

TEST(FieldGrouping, ClearDateGroupRestoresTheBaseWhereItsGroupsWere) {
  Workbook wb = DateWorkbook(2U);
  ASSERT_TRUE(static_cast<bool>(set_date_group(wb, 0U, 0U, kDateField, kMonths | kYears, CalendarSystem::Gregorian, 1U,
                                               std::nullopt, std::nullopt)));
  // Through a group field, as Excel's Ungroup on a year cell.
  ASSERT_TRUE(static_cast<bool>(clear_date_group(wb, 0U, 0U, 5U)));
  const PivotCache& cache = *wb.pivot_caches().front();
  ASSERT_EQ(cache.fields().size(), 4U);
  EXPECT_FALSE(cache.fields()[kDateField].field_group.has_value());
  for (const auto& table : wb.sheet(0).pivot_tables()) {
    ASSERT_EQ(table->fields().size(), 4U);
    EXPECT_EQ(table->row_field_order(), (std::vector<std::uint32_t>{2U}));
    EXPECT_EQ(table->fields()[kDateField].axis, PivotAxis::Row);
  }
  // Clearing what is not grouped changes nothing.
  EXPECT_TRUE(static_cast<bool>(clear_date_group(wb, 0U, 0U, kDateField)));
  EXPECT_FALSE(static_cast<bool>(clear_date_group(wb, 0U, 1U, 9U)));
}

TEST(FieldGrouping, RegroupingReplacesTheGroupFields) {
  Workbook wb = DateWorkbook(1U);
  ASSERT_TRUE(static_cast<bool>(set_date_group(wb, 0U, 0U, kDateField, kMonths | kYears, CalendarSystem::Gregorian, 1U,
                                               std::nullopt, std::nullopt)));
  ASSERT_TRUE(static_cast<bool>(
      set_date_group(wb, 0U, 0U, 4U, kDays, CalendarSystem::Gregorian, 7U, std::nullopt, std::nullopt)));
  const PivotCache& cache = *wb.pivot_caches().front();
  ASSERT_EQ(cache.fields().size(), 5U);
  EXPECT_EQ(cache.fields()[4].name, "日 (Date)");
  EXPECT_EQ(cache.fields()[kDateField].field_group->par.value_or(99U), 4U);
  const PivotTable& table = *wb.sheet(0).pivot_tables()[0];
  ASSERT_EQ(table.fields().size(), 5U);
  EXPECT_EQ(table.row_field_order(), (std::vector<std::uint32_t>{4U}));
  // A failed regroup leaves the grouping as it was.
  EXPECT_FALSE(static_cast<bool>(
      set_date_group(wb, 0U, 0U, kDateField, 0U, CalendarSystem::Gregorian, 1U, std::nullopt, std::nullopt)));
  EXPECT_EQ(cache.fields().size(), 5U);
}

TEST(FieldGrouping, GroupDateIsoSaturatesAtTheLastInstant) {
  EXPECT_EQ(group_date_iso(2958465.99999, false), "9999-12-31T23:59:59");
  EXPECT_EQ(group_date_iso(2958465.0, false), "9999-12-31T00:00:00");
  EXPECT_EQ(group_date_iso(2957003.99999, true), "9999-12-31T23:59:59");
}

// A date mark declares the field typed, so a plain number beside it refuses
// the grouping whether or not hints were captured; an unmarked field groups.
TEST(FieldGrouping, MarkedDateBesideAPlainNumberRefusesGrouping) {
  const std::vector<double> dates = {Day(2024, 1U, 15U), Day(2024, 2U, 1U), Day(2024, 3U, 10U)};
  PivotCache plain = DateCache(dates, /*as_saved=*/false);
  EXPECT_TRUE(static_cast<bool>(
      group_cache_by_date(plain, kDateField, Periods(kMonths), /*date1904=*/false, mac_365_ja_jp_profile())));

  PivotCache mixed = DateCache(dates, /*as_saved=*/false);
  mixed.mutable_records()[0].cell_is_date = {false, false, true, false};
  const Expected<std::size_t, Error> refused =
      group_cache_by_date(mixed, kDateField, Periods(kMonths), /*date1904=*/false, mac_365_ja_jp_profile());
  ASSERT_FALSE(static_cast<bool>(refused));
  EXPECT_EQ(refused.error().code, FormulonErrorCode::kInvalidArgument);

  PivotCache marked = DateCache(dates, /*as_saved=*/false);
  for (PivotCacheRecord& record : marked.mutable_records()) {
    record.cell_is_date = {false, false, true, false};
  }
  EXPECT_TRUE(static_cast<bool>(
      group_cache_by_date(marked, kDateField, Periods(kMonths), /*date1904=*/false, mac_365_ja_jp_profile())));

  PivotCache shared = DateCache(dates, /*as_saved=*/false);
  shared.mutable_fields()[kDateField].shared_items = {Value::number(dates[0]), Value::number(dates[1])};
  shared.mutable_fields()[kDateField].shared_item_is_date = {true, false};
  EXPECT_FALSE(static_cast<bool>(
      group_cache_by_date(shared, kDateField, Periods(kMonths), /*date1904=*/false, mac_365_ja_jp_profile())));
}

TEST(FieldGrouping, CellIsSharedIndexInfersFromTheFieldAndTheDateMark) {
  PivotCacheField shared{"S", {Value::number(7)}};
  const PivotCacheField inline_field{"I", {}};
  PivotCacheRecord record;
  record.cells = {Value::number(0), Value::text("x")};
  EXPECT_TRUE(cell_is_shared_index(record, shared, 0U));
  EXPECT_FALSE(cell_is_shared_index(record, shared, 1U));
  EXPECT_FALSE(cell_is_shared_index(record, inline_field, 0U));
  record.cell_is_date = {true};
  EXPECT_FALSE(cell_is_shared_index(record, shared, 0U));
  record.cell_is_index = {true, false};
  EXPECT_TRUE(cell_is_shared_index(record, inline_field, 0U));
}

}  // namespace
}  // namespace formulon::pivot
