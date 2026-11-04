// Group-item numbering of pivot group fields and its materialisation into
// cache records.

#include "pivot/field_grouping.h"

#include <cstddef>
#include <cstdint>
#include <limits>
#include <optional>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "pivot/pivot_cache.h"
#include "pivot/record_access.h"
#include "utils/date_time.h"
#include "value.h"

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

}  // namespace
}  // namespace formulon::pivot
