// Calendar and date-window filters for pivot fields.

#include <algorithm>
#include <chrono>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>
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
#include "utils/date_time.h"
#include "value.h"

namespace formulon::pivot {
namespace {

using test::build_basic_cache;
using test::build_sum_amount_table;
using test::owned_text;
using test::row_index;

// Relative-period boundary helper.
namespace {
// 2026-05-20 12:00:00. Mid-month, mid-quarter, so every window's two
// boundaries are distinct from the reading itself.
constexpr date_time::CivilTime kMay20{{2026, 5U, 20U}, {12U, 0U, 0U}};

double Serial(int year, unsigned month, unsigned day) {
  return date_time::serial_from_ymd(year, month, day, /*date1904=*/false);
}

void ExpectWindow(RelativePeriod period, const date_time::CivilTime& now, int low_y, unsigned low_m, unsigned low_d,
                  int high_y, unsigned high_m, unsigned high_d) {
  const DateWindow window = resolve_relative_period(period, now, /*date1904=*/false);
  EXPECT_DOUBLE_EQ(window.low, Serial(low_y, low_m, low_d));
  EXPECT_DOUBLE_EQ(window.high, Serial(high_y, high_m, high_d));
}

}  // namespace

PivotCache build_label_date_filter_cache() {
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
  add(formulon::date_time::serial_from_ymd(2024, 1, 1), 10.0);
  add(formulon::date_time::serial_from_ymd(2024, 6, 15), 20.0);
  add(formulon::date_time::serial_from_ymd(2024, 12, 31), 30.0);
  add(formulon::date_time::serial_from_ymd(2025, 3, 1), 40.0);
  return cache;
}

// Helper: builds a single-row-axis pivot over `Date` with SUM(Amount).
// The Date field becomes the row axis so each surviving record forms its
// own row leaf, simplifying assertions.
PivotTable build_label_date_filter_table() {
  PivotTable table;
  table.set_pivot_cache_id(1);
  PivotField date_f;
  date_f.source_name = "Date";
  date_f.axis = PivotAxis::Row;
  PivotField amount_f;
  amount_f.source_name = "Amount";
  amount_f.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(date_f));
  table.mutable_fields().push_back(std::move(amount_f));
  table.mutable_row_field_order() = {0};

  PivotDataField sum_amount;
  sum_amount.name = "Sum of Amount";
  sum_amount.field_index = 1;
  sum_amount.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum_amount));
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  return table;
}
TEST(RelativePeriod, DayWindowsAreSingleDays) {
  ExpectWindow(RelativePeriod::Today, kMay20, 2026, 5, 20, 2026, 5, 20);
  ExpectWindow(RelativePeriod::Yesterday, kMay20, 2026, 5, 19, 2026, 5, 19);
  ExpectWindow(RelativePeriod::Tomorrow, kMay20, 2026, 5, 21, 2026, 5, 21);
}

TEST(RelativePeriod, MonthWindowsEndOnTheirOwnLastDay) {
  // May has 31 days, April 30 -- the window end is derived from the next
  // month's first day, so a wrong month length would show up here.
  ExpectWindow(RelativePeriod::ThisMonth, kMay20, 2026, 5, 1, 2026, 5, 31);
  ExpectWindow(RelativePeriod::LastMonth, kMay20, 2026, 4, 1, 2026, 4, 30);
  ExpectWindow(RelativePeriod::NextMonth, kMay20, 2026, 6, 1, 2026, 6, 30);
}

TEST(RelativePeriod, QuarterWindowsSnapToTheCalendarQuarter) {
  // May sits in Q2, so "this quarter" starts in April rather than in May.
  ExpectWindow(RelativePeriod::ThisQuarter, kMay20, 2026, 4, 1, 2026, 6, 30);
  ExpectWindow(RelativePeriod::LastQuarter, kMay20, 2026, 1, 1, 2026, 3, 31);
  ExpectWindow(RelativePeriod::NextQuarter, kMay20, 2026, 7, 1, 2026, 9, 30);
}

TEST(RelativePeriod, YearWindowsSpanTheWholeCalendarYear) {
  ExpectWindow(RelativePeriod::ThisYear, kMay20, 2026, 1, 1, 2026, 12, 31);
  ExpectWindow(RelativePeriod::LastYear, kMay20, 2025, 1, 1, 2025, 12, 31);
  ExpectWindow(RelativePeriod::NextYear, kMay20, 2027, 1, 1, 2027, 12, 31);
}

TEST(RelativePeriod, YearToDateStopsAtTheReadingNotAtYearEnd) {
  ExpectWindow(RelativePeriod::YearToDate, kMay20, 2026, 1, 1, 2026, 5, 20);
}

TEST(RelativePeriod, WindowsRollOverTheYearBoundary) {
  // January is the case the explicit month arithmetic exists for: naive
  // subtraction would produce month 0 rather than the previous December.
  constexpr date_time::CivilTime kJan10{{2026, 1U, 10U}, {0U, 0U, 0U}};
  ExpectWindow(RelativePeriod::LastMonth, kJan10, 2025, 12, 1, 2025, 12, 31);
  ExpectWindow(RelativePeriod::LastQuarter, kJan10, 2025, 10, 1, 2025, 12, 31);
  constexpr date_time::CivilTime kDec10{{2026, 12U, 10U}, {0U, 0U, 0U}};
  ExpectWindow(RelativePeriod::NextMonth, kDec10, 2027, 1, 1, 2027, 1, 31);
  ExpectWindow(RelativePeriod::NextQuarter, kDec10, 2027, 1, 1, 2027, 3, 31);
}

TEST(RelativePeriod, LeapFebruaryKeepsItsTwentyNinthDay) {
  constexpr date_time::CivilTime kFeb2024{{2024, 2U, 5U}, {0U, 0U, 0U}};
  ExpectWindow(RelativePeriod::ThisMonth, kFeb2024, 2024, 2, 1, 2024, 2, 29);
}

TEST(RelativePeriod, TheWindowFollowsTheWorkbookEpoch) {
  // A 1904-system workbook stores every date 1462 lower, so the resolved
  // window has to shift with it or it would select the wrong records.
  const DateWindow window = resolve_relative_period(RelativePeriod::ThisMonth, kMay20, /*date1904=*/true);
  EXPECT_DOUBLE_EQ(window.low, Serial(2026, 5, 1) - date_time::kDate1904EpochGap);
  EXPECT_DOUBLE_EQ(window.high, Serial(2026, 5, 31) - date_time::kDate1904EpochGap);
}

TEST(RelativePeriod, WeekWindowsRunSundayThroughSaturday) {
  // Excel anchors the week group on the calendar week, so a reading taken
  // mid-week still starts the window on the preceding Sunday rather than
  // seven days back from the reading. 2026-08-21 is a Friday.
  constexpr date_time::CivilTime kAug21{{2026, 8U, 21U}, {9U, 30U, 0U}};
  ExpectWindow(RelativePeriod::ThisWeek, kAug21, 2026, 8, 16, 2026, 8, 22);
  ExpectWindow(RelativePeriod::LastWeek, kAug21, 2026, 8, 9, 2026, 8, 15);
  ExpectWindow(RelativePeriod::NextWeek, kAug21, 2026, 8, 23, 2026, 8, 29);
}

TEST(RelativePeriod, WeekWindowsTileWithoutGapOrOverlap) {
  // The three windows are adjacent by construction; asserting it directly
  // is what rules out an off-by-one that a single window would hide.
  const DateWindow last = resolve_relative_period(RelativePeriod::LastWeek, kMay20, /*date1904=*/false);
  const DateWindow current = resolve_relative_period(RelativePeriod::ThisWeek, kMay20, /*date1904=*/false);
  const DateWindow next = resolve_relative_period(RelativePeriod::NextWeek, kMay20, /*date1904=*/false);
  EXPECT_DOUBLE_EQ(current.low, last.high + 1.0);
  EXPECT_DOUBLE_EQ(next.low, current.high + 1.0);
  EXPECT_DOUBLE_EQ(current.high - current.low, 6.0);
}

TEST(RelativePeriod, AWeekWindowStartingOnSundayDoesNotShiftBack) {
  // A Sunday reading is the boundary case: the week it belongs to is its
  // own, not the one that just ended.
  constexpr date_time::CivilTime kSunday{{2025, 12U, 28U}, {0U, 0U, 0U}};
  ExpectWindow(RelativePeriod::ThisWeek, kSunday, 2025, 12, 28, 2026, 1, 3);
}

TEST(RelativePeriod, WeekWindowsCrossTheYearBoundary) {
  constexpr date_time::CivilTime kNewYear{{2026, 1U, 1U}, {0U, 0U, 0U}};
  ExpectWindow(RelativePeriod::ThisWeek, kNewYear, 2025, 12, 28, 2026, 1, 3);
  ExpectWindow(RelativePeriod::LastWeek, kNewYear, 2025, 12, 21, 2025, 12, 27);
}

TEST(RelativePeriod, TheWeekWindowFollowsTheWorkbookEpoch) {
  const DateWindow window = resolve_relative_period(RelativePeriod::ThisWeek, kMay20, /*date1904=*/true);
  EXPECT_DOUBLE_EQ(window.low, Serial(2026, 5, 17) - date_time::kDate1904EpochGap);
  EXPECT_DOUBLE_EQ(window.high, Serial(2026, 5, 23) - date_time::kDate1904EpochGap);
}

TEST(PivotEvaluator, RecurringMonthFilterKeepsEveryYearsMatchingMonth) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  // Field 2 holds North 100 / 50 / 25 and South 200 / 300, read here as
  // date serials: 25 is 1900-01-25, 50 is 1900-02-19, 100 is 1900-04-09,
  // 200 is 1900-07-18 and 300 is 1900-10-26. January therefore admits
  // North's 25 alone -- and it does so with no clock reading at all,
  // which is what separates this family from the relative periods.
  AuthoredRecurringFilter f;
  f.field_index = 2;
  f.month_low = 1;
  f.month_high = 1;
  table.mutable_authored_recurring_filters().push_back(f);

  auto r_or = evaluate(table, cache, PivotLayoutOptions{}, PivotFilterEnv{});
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().rows.size(), 1U);
  EXPECT_EQ(r_or.value().rows[0].label, "North");
  EXPECT_DOUBLE_EQ(r_or.value().values[0][0][0].as_number(), 25.0);
}

TEST(PivotEvaluator, ARecurringQuarterSpansItsThreeMonths) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  // Q2 is months 4..6; serial 100 is 1900-04-09, so North's 100 survives
  // and its January pair does not.
  AuthoredRecurringFilter f;
  f.field_index = 2;
  f.month_low = 4;
  f.month_high = 6;
  table.mutable_authored_recurring_filters().push_back(f);

  auto r_or = evaluate(table, cache, PivotLayoutOptions{}, PivotFilterEnv{});
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().rows.size(), 1U);
  EXPECT_EQ(r_or.value().rows[0].label, "North");
  EXPECT_DOUBLE_EQ(r_or.value().values[0][0][0].as_number(), 100.0);
}

TEST(PivotEvaluator, ARecurringFilterIsIndependentOfTheClock) {
  // The whole point of the family: the same table evaluated against two
  // very different readings selects the same records.
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  AuthoredRecurringFilter f;
  f.field_index = 2;
  f.month_low = 1;
  f.month_high = 1;
  table.mutable_authored_recurring_filters().push_back(f);

  PivotFilterEnv early;
  early.pinned_now = date_time::CivilTime{{1900, 2U, 15U}, {0U, 0U, 0U}};
  PivotFilterEnv late;
  late.pinned_now = date_time::CivilTime{{2099, 11U, 3U}, {0U, 0U, 0U}};
  auto early_or = evaluate(table, cache, PivotLayoutOptions{}, early);
  auto late_or = evaluate(table, cache, PivotLayoutOptions{}, late);
  ASSERT_TRUE(static_cast<bool>(early_or)) << early_or.error().message;
  ASSERT_TRUE(static_cast<bool>(late_or)) << late_or.error().message;
  EXPECT_DOUBLE_EQ(early_or.value().values[0][0][0].as_number(), 25.0);
  EXPECT_DOUBLE_EQ(late_or.value().values[0][0][0].as_number(), 25.0);
}

TEST(PivotEvaluator, AuthoredPeriodFilterPrunesRecordsAgainstThePinnedClock) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  // Field 2 carries the amounts (North 100 / 50 / 25, South 200 / 300),
  // read here as date serials. Pinning to 1900-02-15 makes "this quarter"
  // 1900-01-01..1900-03-31, i.e. serials 1..91: it admits North's 50 and
  // 25 and excludes everything else, South included.
  AuthoredPeriodFilter f;
  f.field_index = 2;
  f.period = RelativePeriod::ThisQuarter;
  table.mutable_authored_period_filters().push_back(f);

  PivotFilterEnv env;
  env.pinned_now = date_time::CivilTime{{1900, 2U, 15U}, {0U, 0U, 0U}};
  auto r_or = evaluate(table, cache, PivotLayoutOptions{}, env);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().rows.size(), 1U);
  EXPECT_EQ(r_or.value().rows[0].label, "North");
  ASSERT_EQ(r_or.value().values.size(), 1U);
  ASSERT_EQ(r_or.value().values[0].size(), 1U);
  ASSERT_EQ(r_or.value().values[0][0].size(), 1U);
  EXPECT_DOUBLE_EQ(r_or.value().values[0][0][0].as_number(), 75.0);
}

TEST(PivotEvaluator, MovingThePinnedClockMovesWhichRecordsSurvive) {
  // Same filter, a quarter later: the window becomes serials 92..181,
  // which admits North's 100 alone and still excludes South.
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  AuthoredPeriodFilter f;
  f.field_index = 2;
  f.period = RelativePeriod::ThisQuarter;
  table.mutable_authored_period_filters().push_back(f);

  PivotFilterEnv env;
  env.pinned_now = date_time::CivilTime{{1900, 5U, 15U}, {0U, 0U, 0U}};
  auto r_or = evaluate(table, cache, PivotLayoutOptions{}, env);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().rows.size(), 1U);
  EXPECT_EQ(r_or.value().rows[0].label, "North");
  ASSERT_EQ(r_or.value().values.size(), 1U);
  ASSERT_EQ(r_or.value().values[0].size(), 1U);
  ASSERT_EQ(r_or.value().values[0][0].size(), 1U);
  EXPECT_DOUBLE_EQ(r_or.value().values[0][0][0].as_number(), 100.0);
}

TEST(PivotEvaluator, APeriodFilterOnAnOutOfRangeFieldIsInert) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  AuthoredPeriodFilter f;
  f.field_index = 99;
  f.period = RelativePeriod::Today;
  table.mutable_authored_period_filters().push_back(f);

  PivotFilterEnv env;
  env.pinned_now = date_time::CivilTime{{1900, 2U, 15U}, {0U, 0U, 0U}};
  auto r_or = evaluate(table, cache, PivotLayoutOptions{}, env);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  EXPECT_EQ(r_or.value().rows.size(), 2U);
}

TEST(PivotEvaluator, AuthoredDateFilterWithNoUpperBoundIsInert) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  AuthoredValueFilter f;
  f.field_index = 2;
  f.type = FilterType::LabelDate;
  f.value = 1000.0;
  table.mutable_authored_value_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  EXPECT_EQ(r_or.value().rows.size(), 2U);
}

TEST(PivotEvaluator, LabelDateFilterIncludesInRange) {
  PivotCache cache = build_label_date_filter_cache();
  PivotTable table = build_label_date_filter_table();

  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Date";
  f.type = FilterType::LabelDate;
  f.value = formulon::date_time::serial_from_ymd(2024, 1, 1);
  f.value_high = formulon::date_time::serial_from_ymd(2024, 12, 31);
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Three of the four records fall inside [2024-01-01, 2024-12-31].
  ASSERT_EQ(r.rows.size(), 3U);
  std::vector<double> totals;
  totals.reserve(r.values.size());
  for (const auto& row_slot : r.values) {
    ASSERT_EQ(row_slot.size(), 1U);
    ASSERT_EQ(row_slot[0].size(), 1U);
    totals.push_back(row_slot[0][0].as_number());
  }
  std::sort(totals.begin(), totals.end());
  EXPECT_DOUBLE_EQ(totals[0], 10.0);
  EXPECT_DOUBLE_EQ(totals[1], 20.0);
  EXPECT_DOUBLE_EQ(totals[2], 30.0);
}

TEST(PivotEvaluator, LabelDateFilterUnboundedHighIsNoOp) {
  PivotCache cache = build_label_date_filter_cache();
  PivotTable table = build_label_date_filter_table();

  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Date";
  f.type = FilterType::LabelDate;
  f.value = formulon::date_time::serial_from_ymd(2024, 1, 1);
  // value_high left as default monostate -> filter degrades to no-op.
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // No upper bound -> all four records survive.
  ASSERT_EQ(r.rows.size(), 4U);
}

TEST(PivotEvaluator, LabelDateFiltersWithInvalidNumericCriteriaAreNoOp) {
  const double invalid[] = {
      std::numeric_limits<double>::quiet_NaN(),
      std::numeric_limits<double>::infinity(),
      -std::numeric_limits<double>::infinity(),
      std::numeric_limits<double>::max(),
  };
  for (const double raw : invalid) {
    PivotCache active_cache = build_label_date_filter_cache();
    PivotTable active_table = build_label_date_filter_table();
    PivotFilter active;
    active.axis = PivotAxis::Row;
    active.field_name = "Date";
    active.type = FilterType::LabelDate;
    active.value = raw;
    active.value_high = date_time::serial_from_ymd(2025, 12U, 31U);
    active_table.mutable_active_filters().push_back(active);
    auto active_or = evaluate(active_table, active_cache);
    ASSERT_TRUE(static_cast<bool>(active_or)) << active_or.error().message << " low=" << raw;
    EXPECT_EQ(active_or.value().rows.size(), 4U) << "invalid low=" << raw;

    active_cache = build_label_date_filter_cache();
    active_table = build_label_date_filter_table();
    active.value = 0.0;
    active.value_high = raw;
    active_table.mutable_active_filters().push_back(active);
    active_or = evaluate(active_table, active_cache);
    ASSERT_TRUE(static_cast<bool>(active_or)) << active_or.error().message << " high=" << raw;
    EXPECT_EQ(active_or.value().rows.size(), 4U) << "invalid high=" << raw;

    PivotCache authored_cache = build_label_date_filter_cache();
    PivotTable authored_table = build_label_date_filter_table();
    AuthoredValueFilter authored;
    authored.field_index = 0;
    authored.type = FilterType::LabelDate;
    authored.value = raw;
    authored.value_high = date_time::serial_from_ymd(2025, 12U, 31U);
    authored_table.mutable_authored_value_filters().push_back(authored);
    auto authored_or = evaluate(authored_table, authored_cache);
    ASSERT_TRUE(static_cast<bool>(authored_or)) << authored_or.error().message << " low=" << raw;
    EXPECT_EQ(authored_or.value().rows.size(), 4U) << "invalid authored low=" << raw;

    authored_cache = build_label_date_filter_cache();
    authored_table = build_label_date_filter_table();
    authored.value = 0.0;
    authored.value_high = raw;
    authored_table.mutable_authored_value_filters().push_back(authored);
    authored_or = evaluate(authored_table, authored_cache);
    ASSERT_TRUE(static_cast<bool>(authored_or)) << authored_or.error().message << " high=" << raw;
    EXPECT_EQ(authored_or.value().rows.size(), 4U) << "invalid authored high=" << raw;
  }
}

TEST(PivotEvaluator, RelativeAndRecurringDateFiltersPassInvalidNumericRecords) {
  const double invalid[] = {
      std::numeric_limits<double>::quiet_NaN(),
      std::numeric_limits<double>::infinity(),
      -std::numeric_limits<double>::infinity(),
      std::numeric_limits<double>::max(),
  };

  PivotCache period_cache = build_label_date_filter_cache();
  for (std::size_t i = 0; i < std::size(invalid); ++i) {
    PivotCacheRecord record;
    record.cells = {Value::number(invalid[i]), Value::number(100.0 + static_cast<double>(i))};
    period_cache.mutable_records().push_back(std::move(record));
  }
  PivotTable period_table = build_label_date_filter_table();
  AuthoredPeriodFilter period;
  period.field_index = 0;
  period.period = RelativePeriod::ThisYear;
  period_table.mutable_authored_period_filters().push_back(period);
  PivotFilterEnv env;
  env.pinned_now = date_time::CivilTime{{2024, 6U, 15U}, {0U, 0U, 0U}};
  auto period_or = evaluate(period_table, period_cache, PivotLayoutOptions{}, env);
  ASSERT_TRUE(static_cast<bool>(period_or)) << period_or.error().message;
  ASSERT_EQ(period_or.value().rows.size(), 7U);
  double period_total = 0.0;
  for (const auto& row_slot : period_or.value().values) {
    period_total += row_slot[0][0].as_number();
  }
  EXPECT_DOUBLE_EQ(period_total, 10.0 + 20.0 + 30.0 + 100.0 + 101.0 + 102.0 + 103.0);

  PivotCache recurring_cache = build_label_date_filter_cache();
  for (std::size_t i = 0; i < std::size(invalid); ++i) {
    PivotCacheRecord record;
    record.cells = {Value::number(invalid[i]), Value::number(100.0 + static_cast<double>(i))};
    recurring_cache.mutable_records().push_back(std::move(record));
  }
  PivotTable recurring_table = build_label_date_filter_table();
  AuthoredRecurringFilter recurring;
  recurring.field_index = 0;
  recurring.month_low = 1;
  recurring.month_high = 1;
  recurring_table.mutable_authored_recurring_filters().push_back(recurring);
  auto recurring_or = evaluate(recurring_table, recurring_cache);
  ASSERT_TRUE(static_cast<bool>(recurring_or)) << recurring_or.error().message;
  ASSERT_EQ(recurring_or.value().rows.size(), 5U);
  double recurring_total = 0.0;
  for (const auto& row_slot : recurring_or.value().values) {
    recurring_total += row_slot[0][0].as_number();
  }
  EXPECT_DOUBLE_EQ(recurring_total, 10.0 + 100.0 + 101.0 + 102.0 + 103.0);
}

TEST(PivotEvaluator, ValueFiltersRetainNonFiniteComparisonSemantics) {
  const double invalid[] = {
      std::numeric_limits<double>::quiet_NaN(),
      std::numeric_limits<double>::infinity(),
      -std::numeric_limits<double>::infinity(),
  };
  for (const double raw : invalid) {
    PivotCache active_cache = build_basic_cache();
    PivotTable active_table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
    PivotFilter active;
    active.axis = PivotAxis::Row;
    active.field_name = "Region";
    active.type = FilterType::ValueGreaterThan;
    active.value = raw;
    active_table.mutable_active_filters().push_back(active);
    auto active_or = evaluate(active_table, active_cache);
    ASSERT_TRUE(static_cast<bool>(active_or)) << active_or.error().message;
    const std::size_t expected_active = std::isinf(raw) && std::signbit(raw) ? 2U : 0U;
    EXPECT_EQ(active_or.value().rows.size(), expected_active) << "active invalid=" << raw;

    PivotCache authored_cache = build_basic_cache();
    PivotTable authored_table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
    AuthoredValueFilter authored;
    authored.field_index = 0;
    authored.type = FilterType::ValueTop10;
    authored.value = raw;
    authored_table.mutable_authored_value_filters().push_back(authored);
    auto authored_or = evaluate(authored_table, authored_cache);
    ASSERT_TRUE(static_cast<bool>(authored_or)) << authored_or.error().message;
    const std::size_t expected_authored = std::isinf(raw) && raw > 0.0 ? 2U : 0U;
    EXPECT_EQ(authored_or.value().rows.size(), expected_authored) << "authored invalid=" << raw;

    PivotCache between_cache = build_basic_cache();
    PivotTable between_table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
    PivotFilter between;
    between.axis = PivotAxis::Row;
    between.field_name = "Region";
    between.type = FilterType::ValueBetween;
    between.value = 0.0;
    between.value_high = raw;
    between_table.mutable_active_filters().push_back(between);
    auto between_or = evaluate(between_table, between_cache);
    ASSERT_TRUE(static_cast<bool>(between_or)) << between_or.error().message;
    const std::size_t expected_between = std::isinf(raw) && raw > 0.0 ? 2U : 0U;
    EXPECT_EQ(between_or.value().rows.size(), expected_between) << "active between invalid=" << raw;
  }
}

TEST(PivotEvaluator, AuthoredDateFilterPrunesRecordsByTheirSerial) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  // Field 2 holds the amounts, which are plain numbers -- the date pass
  // reads whatever serial the bound cell carries, so a range that spans
  // only North's two smaller amounts drops South entirely.
  AuthoredValueFilter f;
  f.field_index = 2;
  f.type = FilterType::LabelDate;
  f.value = 0.0;
  f.value_high = 100.0;
  table.mutable_authored_value_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().rows.size(), 1U);
  EXPECT_EQ(r_or.value().rows[0].label, "North");
}

}  // namespace
}  // namespace formulon::pivot
