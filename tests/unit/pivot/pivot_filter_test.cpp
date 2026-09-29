//
// Unit tests for the filters `formulon::pivot::evaluate` applies: label,
// caption, value, date and relative-period filters, manual item visibility,
// and the page (report filter) axis selection.

#include <algorithm>
#include <chrono>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <utility>
#include <variant>
#include <vector>

#include "gtest/gtest.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_evaluator.h"
#include "pivot/pivot_index.h"
#include "pivot/pivot_layout.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot_evaluator_fixtures.h"
#include "utils/date_time.h"
#include "utils/error.h"
#include "value.h"

namespace formulon::pivot {
namespace {

using test::build_basic_cache;
using test::build_sum_amount_table;
using test::owned_text;
using test::row_index;

// Same three-column shape as `build_basic_cache()`, but the amounts are
// deliberately non-integral / past the 64-bit integer range so a record's
// label goes through the general numeric rendering path rather than the
// integral one.
//
//   Region  Product  Amount
//   ------  -------  ------
//   North   Widget      1.5
//   North   Gadget      2
//   South   Widget     1e20
//   South   Gadget      4
PivotCache build_non_integral_amount_cache() {
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Region", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Product", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});

  auto add = [&](const char* region, const char* product, double amount) {
    PivotCacheRecord rec;
    rec.cells.push_back(owned_text(cache, region));
    rec.cells.push_back(owned_text(cache, product));
    rec.cells.push_back(Value::number(amount));
    cache.mutable_records().push_back(std::move(rec));
  };

  add("North", "Widget", 1.5);
  add("North", "Gadget", 2.0);
  add("South", "Widget", 1e20);
  add("South", "Gadget", 4.0);
  return cache;
}

// Index into `build_shared_region_cache()`'s Region `shared_items` of the
// entry with no value, and of the text entry deliberately spelled like the
// default blank placeholder.
constexpr std::uint32_t kSharedRegionBlank = 1U;
constexpr std::uint32_t kSharedRegionPlaceholderText = 2U;

// Same three-column shape as `build_basic_cache()`, but Region is a *shared*
// field the way a discrete axis column arrives from OOXML: each record stores
// an index into `shared_items`, one of which has no value at all.
//
// The third shared item is a genuine text value spelled exactly like the
// default blank placeholder, so a filter that identified the empty item by its
// rendered label would catch this row too.
//
//   Region     Product  Amount
//   ---------  -------  ------
//   North      Widget    100
//   <no value> Widget     10
//   "(blank)"  Widget      7
PivotCache build_shared_region_cache() {
  PivotCache cache;
  cache.set_cache_id(1);
  PivotCacheField region;
  region.name = "Region";
  region.shared_items.push_back(owned_text(cache, "North"));
  region.shared_items.push_back(Value::blank());
  region.shared_items.push_back(owned_text(cache, std::string{PivotLayoutOptions{}.blank_item_label}));
  cache.mutable_fields().push_back(std::move(region));
  cache.mutable_fields().push_back(PivotCacheField{"Product", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});

  auto add = [&](std::uint32_t region_index, const char* product, double amount) {
    PivotCacheRecord rec;
    rec.cells = {Value::number(region_index), owned_text(cache, product), Value::number(amount)};
    rec.cell_is_index = {true, false, false};
    cache.mutable_records().push_back(std::move(rec));
  };

  add(0U, "Widget", 100.0);
  add(kSharedRegionBlank, "Widget", 10.0);
  add(kSharedRegionPlaceholderText, "Widget", 7.0);
  return cache;
}

// ---------------------------------------------------------------------------
// 8d. PivotFilter (LabelContains / LabelBeginsWith / ValueTop10 / ValueGreaterThan)
// ---------------------------------------------------------------------------

TEST(PivotEvaluator, LabelContainsFilterDropsRecords) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  // Drop records whose Region label contains "outh" (i.e. South).
  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Region";
  f.type = FilterType::LabelContains;
  f.value = std::string("outh");
  // PivotFilter pre-aggregation reject: every match drops the record.
  // Wait — LabelContains with payload "outh" *passes* records whose label
  // contains it. To drop South we instead use LabelBeginsWith("North")
  // below; LabelContains here keeps only South.
  table.mutable_active_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, "South");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 500.0);  // 200 + 300
}

TEST(PivotEvaluator, LabelBeginsWithFilterDropsRecords) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Region";
  f.type = FilterType::LabelBeginsWith;
  f.value = std::string("Nor");
  table.mutable_active_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, "North");
}

// ---------------------------------------------------------------------------
// 8e. AuthoredCaptionFilter (decoded from an OOXML `<filters>` block)
// ---------------------------------------------------------------------------

TEST(PivotEvaluator, AuthoredCaptionEqualFilterDropsRecords) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  AuthoredCaptionFilter f;
  f.field_index = 0;  // Region
  f.predicate = CaptionPredicate::Equal;
  f.value = "North";
  table.mutable_authored_caption_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, "North");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 175.0);  // 100 + 50 + 25
}

// Excel matches a caption filter the way it matches an AutoFilter
// criterion, so a criterion that differs only in case still selects.
TEST(PivotEvaluator, AuthoredCaptionFilterIgnoresCase) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  AuthoredCaptionFilter f;
  f.field_index = 0;
  f.predicate = CaptionPredicate::Equal;
  f.value = "nOrTh";
  table.mutable_authored_caption_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().rows.size(), 1U);
  EXPECT_EQ(r_or.value().rows[0].label, "North");
}

TEST(PivotEvaluator, AuthoredCaptionNotContainsFilterDropsMatchingRecords) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  AuthoredCaptionFilter f;
  f.field_index = 0;
  f.predicate = CaptionPredicate::NotContains;
  f.value = "outh";
  table.mutable_authored_caption_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().rows.size(), 1U);
  EXPECT_EQ(r_or.value().rows[0].label, "North");
}

TEST(PivotEvaluator, AuthoredCaptionBetweenFilterKeepsTheInclusiveRange) {
  PivotCache cache = build_basic_cache();
  // Row on Product so the range spans "Gadget" / "Widget".
  PivotTable table = build_sum_amount_table(/*row=*/{1}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  AuthoredCaptionFilter f;
  f.field_index = 1;  // Product
  f.predicate = CaptionPredicate::Between;
  f.value = "A";
  f.value_high = "M";
  table.mutable_authored_caption_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().rows.size(), 1U);
  EXPECT_EQ(r_or.value().rows[0].label, "Gadget");
  EXPECT_DOUBLE_EQ(r_or.value().values[0][0][0].as_number(), 350.0);  // 50 + 300
}

// A `fld` past the end of `<pivotFields>` is a malformed definition; it
// must filter nothing rather than drop every record.
TEST(PivotEvaluator, AuthoredCaptionFilterWithOutOfRangeFieldIsInert) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  AuthoredCaptionFilter f;
  f.field_index = 99;
  f.predicate = CaptionPredicate::Equal;
  f.value = "North";
  table.mutable_authored_caption_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  EXPECT_EQ(r_or.value().rows.size(), 2U);
}

// ---------------------------------------------------------------------------
// 8f. AuthoredValueFilter (the value and date half of the same block)
//
// The basic cache totals North 175 and South 500 by Region.
// ---------------------------------------------------------------------------

TEST(PivotEvaluator, AuthoredTopCountKeepsTheHighestScoringLeaf) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  AuthoredValueFilter f;
  f.field_index = 0;
  f.type = FilterType::ValueTop10;
  f.value = 1.0;
  table.mutable_authored_value_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().rows.size(), 1U);
  EXPECT_EQ(r_or.value().rows[0].label, "South");
}

// Excel's "Bottom N" dialog option writes the identically shaped
// `<top10>` with `top="0"`. Keeping the wrong direction here would keep
// South (175 < 500 makes North the actual bottom) instead of North.
TEST(PivotEvaluator, AuthoredTopCountWithTopFalseKeepsTheLowestScoringLeaf) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  AuthoredValueFilter f;
  f.field_index = 0;
  f.type = FilterType::ValueTop10;
  f.value = 1.0;
  f.top = false;
  table.mutable_authored_value_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().rows.size(), 1U);
  EXPECT_EQ(r_or.value().rows[0].label, "North");
}

TEST(PivotEvaluator, AuthoredGreaterThanIsStrictOnTheThreshold) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  AuthoredValueFilter f;
  f.field_index = 0;
  f.type = FilterType::ValueGreaterThan;
  f.value = 175.0;  // exactly North's total
  table.mutable_authored_value_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().rows.size(), 1U);
  EXPECT_EQ(r_or.value().rows[0].label, "South");
}

TEST(PivotEvaluator, AuthoredBetweenIncludesBothBounds) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  AuthoredValueFilter f;
  f.field_index = 0;
  f.type = FilterType::ValueBetween;
  f.value = 175.0;
  f.value_high = 500.0;
  table.mutable_authored_value_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  EXPECT_EQ(r_or.value().rows.size(), 2U);
}

// A file names only the field; the axis it prunes has to be recovered
// from that field's place in the row / column order. A field on neither
// axis leaves nothing to prune.
TEST(PivotEvaluator, AuthoredValueFilterOnAnOffAxisFieldIsInert) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  AuthoredValueFilter f;
  f.field_index = 2;  // Amount: a data field, on no axis
  f.type = FilterType::ValueTop10;
  f.value = 1.0;
  table.mutable_authored_value_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  EXPECT_EQ(r_or.value().rows.size(), 2U);
}

// Ranked by aggregate, so it cannot be decided until the aggregates
// exist -- unlike its caption and date siblings, which prune records.
// Pinning it on the column axis proves the axis recovery is not
// hard-coded to rows.
TEST(PivotEvaluator, AuthoredValueFilterPrunesTheColumnAxisToo) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{}, /*col=*/{0});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  AuthoredValueFilter f;
  f.field_index = 0;
  f.type = FilterType::ValueTop10;
  f.value = 1.0;
  table.mutable_authored_value_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().cols.size(), 1U);
  EXPECT_EQ(r_or.value().cols[0].label, "South");
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

// ---------------------------------------------------------------------------
// 8g. Relative-period filters
//
// These carry no bounds in the file, so the substance is the calendar
// arithmetic that turns a period name plus a clock reading into a window.
// The boundaries are asserted directly against `serial_from_ymd` rather
// than against literal serials: the point under test is which calendar day
// each window starts and ends on, not the serial encoding, which
// `DateTimeDate` already pins.
// ---------------------------------------------------------------------------

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

// An unbounded range is a no-op rather than a half-open filter, matching
// how `PivotFilter` treats a missing upper bound.
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

TEST(PivotEvaluator, ValueTop10FilterKeepsTopRows) {
  // Use a richer cache so ranking is meaningful.
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Region", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});
  auto add = [&](const char* region, double amount) {
    PivotCacheRecord rec;
    rec.cells.push_back(owned_text(cache, region));
    rec.cells.push_back(Value::number(amount));
    cache.mutable_records().push_back(std::move(rec));
  };
  add("A", 10.0);
  add("B", 50.0);
  add("C", 30.0);
  add("D", 20.0);

  PivotTable table;
  table.set_pivot_cache_id(1);
  PivotField rf;
  rf.source_name = "Region";
  rf.axis = PivotAxis::Row;
  PivotField af;
  af.source_name = "Amount";
  af.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(rf));
  table.mutable_fields().push_back(std::move(af));
  table.mutable_row_field_order() = {0};
  PivotDataField sum;
  sum.name = "Sum of Amount";
  sum.field_index = 1;
  sum.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum));
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Region";
  f.type = FilterType::ValueTop10;
  f.value = 2;  // Top 2.
  table.mutable_active_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Top 2 by amount: B (50), C (30).
  ASSERT_EQ(r.rows.size(), 2U);
  // Rows are still in alphabetical order after pruning (we kept only the
  // surviving leaves' positions); order: B, C.
  std::vector<std::string> labels;
  std::vector<double> totals;
  for (std::size_t i = 0; i < r.rows.size(); ++i) {
    labels.push_back(r.rows[i].label);
    totals.push_back(r.values[i][0][0].as_number());
  }
  // Sort to make assertions order-independent.
  std::sort(labels.begin(), labels.end());
  std::sort(totals.begin(), totals.end());
  EXPECT_EQ(labels[0], "B");
  EXPECT_EQ(labels[1], "C");
  EXPECT_DOUBLE_EQ(totals[0], 30.0);
  EXPECT_DOUBLE_EQ(totals[1], 50.0);
}

TEST(PivotEvaluator, ValueGreaterThanFilterKeepsAboveThreshold) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Region";
  f.type = FilterType::ValueGreaterThan;
  f.value = 200.0;  // Threshold; North (175) drops, South (500) survives.
  table.mutable_active_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, "South");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 500.0);
}

TEST(PivotEvaluator, ValueTop10FilterOnColAxis) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{1});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  PivotFilter f;
  f.axis = PivotAxis::Col;
  f.field_name = "Product";
  f.type = FilterType::ValueTop10;
  f.value = 1;  // Keep 1 column with the highest column total.
  table.mutable_active_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Widget total = 100 + 200 + 25 = 325; Gadget total = 50 + 300 = 350.
  // Top-1 by col total -> Gadget survives.
  ASSERT_EQ(r.cols.size(), 1U);
  EXPECT_EQ(r.cols[0].label, "Gadget");
  // Each row keeps a single col slot.
  ASSERT_EQ(r.values.size(), 2U);
  for (const auto& row_slot : r.values) {
    ASSERT_EQ(row_slot.size(), 1U);
  }
}

TEST(PivotEvaluator, ValueFilterSelectsDataFieldAndInvalidSelectorIsNoOp) {
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Region", {}});
  cache.mutable_fields().push_back(PivotCacheField{"MeasureA", {}});
  cache.mutable_fields().push_back(PivotCacheField{"MeasureB", {}});
  auto add = [&](const char* region, double a, double b) {
    PivotCacheRecord rec;
    rec.cells.push_back(owned_text(cache, region));
    rec.cells.push_back(Value::number(a));
    rec.cells.push_back(Value::number(b));
    cache.mutable_records().push_back(std::move(rec));
  };
  add("A", 100.0, 1.0);
  add("B", 50.0, 100.0);

  PivotTable table;
  table.set_pivot_cache_id(1);
  PivotField region_field;
  region_field.source_name = "Region";
  region_field.axis = PivotAxis::Row;
  PivotField measure_a_field;
  measure_a_field.source_name = "MeasureA";
  measure_a_field.axis = PivotAxis::Value;
  PivotField measure_b_field;
  measure_b_field.source_name = "MeasureB";
  measure_b_field.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(region_field));
  table.mutable_fields().push_back(std::move(measure_a_field));
  table.mutable_fields().push_back(std::move(measure_b_field));
  table.mutable_row_field_order() = {0};
  PivotDataField measure_a;
  measure_a.name = "A";
  measure_a.field_index = 1;
  measure_a.aggregation = Aggregation::Sum;
  PivotDataField measure_b;
  measure_b.name = "B";
  measure_b.field_index = 2;
  measure_b.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(measure_a));
  table.mutable_data_fields().push_back(std::move(measure_b));

  PivotFilter selected;
  selected.axis = PivotAxis::Row;
  selected.field_name = "Region";
  selected.type = FilterType::ValueTop10;
  selected.value = 1;
  selected.data_field_index = 1;
  table.mutable_active_filters().push_back(selected);

  auto selected_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(selected_or)) << selected_or.error().message;
  ASSERT_EQ(selected_or.value().rows.size(), 1U);
  EXPECT_EQ(selected_or.value().rows[0].label, "B");

  // A direct C++ model can still contain a stale selector; evaluation must
  // leave the report untouched rather than dropping every leaf.
  table.mutable_active_filters()[0].data_field_index = 2;
  auto invalid_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(invalid_or)) << invalid_or.error().message;
  ASSERT_EQ(invalid_or.value().rows.size(), 2U);
  EXPECT_EQ(invalid_or.value().rows[0].label, "A");
  EXPECT_EQ(invalid_or.value().rows[1].label, "B");
}

// ---------------------------------------------------------------------------
// 9. Manual filter (PivotItem::visible == false) hides records
// ---------------------------------------------------------------------------

TEST(PivotEvaluator, ManualFilterHidesItem) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  // Hide all records with Region == "South" via the Region field's items.
  table.mutable_fields()[0].items = {
      PivotItem{"North", /*visible=*/true},
      PivotItem{"South", /*visible=*/false},
  };

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Only North survives.
  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, "North");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 175.0);
}

// A source cell with no value is an axis item like any other, so the manual
// filter has to be able to hide it. It is the one item that cannot be named:
// its label is the locale's placeholder, which a genuine text value is free to
// spell identically, so the filter identifies it by the cache value the item
// binds to instead.
TEST(PivotEvaluator, ManualFilterHidesTheBlankItem) {
  PivotCache cache = build_shared_region_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  PivotItem hidden_blank;
  hidden_blank.visible = false;
  hidden_blank.has_cache_index = true;
  hidden_blank.cache_index = kSharedRegionBlank;
  table.mutable_fields()[0].items.push_back(hidden_blank);
  // The load path leaves the item unnamed, because the value it binds to
  // renders to nothing. Running it here pins that the filter works on the
  // shape a workbook actually arrives in.
  resolve_pivot_names(table, cache);
  ASSERT_TRUE(table.fields()[0].items[0].name.empty());

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  const PivotLayoutOptions defaults;
  ASSERT_EQ(r.rows.size(), 2U);
  EXPECT_DOUBLE_EQ(r.values[row_index(r, "North")][0][0].as_number(), 100.0);
  // The text value spelled like the placeholder is untouched; only the group
  // with no value is gone, and its 10 is out of every aggregate.
  const std::size_t placeholder_text = row_index(r, defaults.blank_item_label);
  ASSERT_NE(placeholder_text, static_cast<std::size_t>(-1));
  EXPECT_DOUBLE_EQ(r.values[placeholder_text][0][0].as_number(), 7.0);
}

// The mirror direction: hiding the text item that happens to be spelled like
// the placeholder must not take the blank group with it.
TEST(PivotEvaluator, HidingPlaceholderSpelledTextKeepsTheBlankItem) {
  PivotCache cache = build_shared_region_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  PivotItem hidden_text;
  hidden_text.visible = false;
  hidden_text.has_cache_index = true;
  hidden_text.cache_index = kSharedRegionPlaceholderText;
  table.mutable_fields()[0].items.push_back(hidden_text);
  resolve_pivot_names(table, cache);
  const PivotLayoutOptions defaults;
  ASSERT_EQ(table.fields()[0].items[0].name, defaults.blank_item_label);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  ASSERT_EQ(r.rows.size(), 2U);
  EXPECT_DOUBLE_EQ(r.values[row_index(r, "North")][0][0].as_number(), 100.0);
  const std::size_t blank_leaf = row_index(r, defaults.blank_item_label);
  ASSERT_NE(blank_leaf, static_cast<std::size_t>(-1));
  EXPECT_DOUBLE_EQ(r.values[blank_leaf][0][0].as_number(), 10.0);
}

// `build_basic_cache()` with Region shared the way it arrives from a file:
// records store an index into `shared_items` {North, South}.
PivotCache build_shared_north_south_cache() {
  PivotCache cache;
  cache.set_cache_id(1);
  PivotCacheField region;
  region.name = "Region";
  region.shared_items.push_back(owned_text(cache, "North"));
  region.shared_items.push_back(owned_text(cache, "South"));
  cache.mutable_fields().push_back(std::move(region));
  cache.mutable_fields().push_back(PivotCacheField{"Product", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});
  auto add = [&](std::uint32_t region_index, const char* product, double amount) {
    PivotCacheRecord rec;
    rec.cells = {Value::number(region_index), owned_text(cache, product), Value::number(amount)};
    rec.cell_is_index = {true, false, false};
    cache.mutable_records().push_back(std::move(rec));
  };
  add(0U, "Widget", 100.0);
  add(0U, "Gadget", 50.0);
  add(1U, "Widget", 200.0);
  add(1U, "Gadget", 300.0);
  add(0U, "Widget", 25.0);
  return cache;
}

// An item built by cache index carries no name until a load resolves it; the
// filter reads its label from the binding, so it hides the bound value
// without `resolve_pivot_names` ever running.
TEST(PivotEvaluator, UnnamedItemBoundByCacheIndexHidesItsValue) {
  PivotCache cache = build_shared_north_south_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  PivotItem hidden_south;
  hidden_south.visible = false;
  hidden_south.has_cache_index = true;
  hidden_south.cache_index = 1U;
  table.mutable_fields()[0].items.push_back(hidden_south);
  ASSERT_TRUE(table.fields()[0].items[0].name.empty());

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, "North");
  ASSERT_TRUE(r.grand_total.is_number());
  EXPECT_DOUBLE_EQ(r.grand_total.as_number(), 175.0);
}

TEST(PivotItemLabel, NameWinsThenBindingThenEmpty) {
  const PivotCache cache = build_shared_north_south_cache();
  PivotItem named;
  named.name = "Renamed";
  named.cache_index = 1U;
  EXPECT_EQ(pivot_item_label(cache, 0U, named), "Renamed");
  PivotItem bound;
  bound.has_cache_index = true;
  bound.cache_index = 1U;
  EXPECT_EQ(pivot_item_label(cache, 0U, bound), "South");
  PivotItem dangling;
  dangling.has_cache_index = true;
  dangling.cache_index = 9U;
  EXPECT_EQ(pivot_item_label(cache, 0U, dangling), "");
  EXPECT_EQ(pivot_item_label(cache, 7U, bound), "");
}

TEST(PivotEvaluator, ManualFilterMatchesNumericDisplayLabel) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  // The numeric Amount field is rendered as its integral display label.
  // Hiding 300 removes only South/Gadget without allocating a label string
  // per record while the filter is evaluated.
  table.mutable_fields()[2].items = {PivotItem{"300", /*visible=*/false}};

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  const std::size_t north = row_index(r, "North");
  const std::size_t south = row_index(r, "South");
  ASSERT_NE(north, static_cast<std::size_t>(-1));
  ASSERT_NE(south, static_cast<std::size_t>(-1));
  EXPECT_DOUBLE_EQ(r.values[north][0][0].as_number(), 175.0);
  EXPECT_DOUBLE_EQ(r.values[south][0][0].as_number(), 200.0);
}

// A hidden item is matched against the record's label, so the two sides must
// spell a non-integral number the same way. A record-side renderer of its own
// would spell 1.5 as "1.500000" and 1e20 as its 27-character fixed-point form,
// neither of which any authored item can name — the records would stay in
// every aggregate no matter what the filter says.
TEST(PivotEvaluator, ManualFilterHidesNonIntegralNumericItem) {
  PivotCache cache = build_non_integral_amount_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  table.mutable_fields()[2].items = {PivotItem{display_string(Value::number(1.5)), /*visible=*/false}};

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  const std::size_t north = row_index(r, "North");
  ASSERT_NE(north, static_cast<std::size_t>(-1));
  // North keeps only the 2.0 record; the 1.5 one is gone from the aggregate.
  EXPECT_DOUBLE_EQ(r.values[north][0][0].as_number(), 2.0);
}

TEST(PivotEvaluator, ManualFilterHidesNumericItemBeyondIntegerRange) {
  PivotCache cache = build_non_integral_amount_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  table.mutable_fields()[2].items = {PivotItem{display_string(Value::number(1e20)), /*visible=*/false}};

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  const std::size_t south = row_index(r, "South");
  ASSERT_NE(south, static_cast<std::size_t>(-1));
  EXPECT_DOUBLE_EQ(r.values[south][0][0].as_number(), 4.0);
}

// A label filter reads the same rendering, so a needle taken from the axis
// label of a numeric item selects the record that produced it.
TEST(PivotEvaluator, LabelContainsFilterMatchesNumericItemLabel) {
  PivotCache cache = build_non_integral_amount_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Amount";
  f.type = FilterType::LabelContains;
  f.value = display_string(Value::number(1e20));  // "1E+20"
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, "South");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 1e20);
}

// Excel authors an `items[]` list for every field it places on an axis, so the
// list's presence says nothing about whether a filter is set. A list that hides
// nothing must leave the result exactly as it was.
TEST(PivotEvaluator, AnItemListThatHidesNothingDoesNotChangeTheResult) {
  PivotCache cache = build_basic_cache();
  PivotTable unfiltered = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  unfiltered.set_grand_totals(/*rows=*/false, /*cols=*/false);
  auto baseline_or = evaluate(unfiltered, cache);
  ASSERT_TRUE(static_cast<bool>(baseline_or)) << baseline_or.error().message;

  PivotTable listed = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  listed.set_grand_totals(/*rows=*/false, /*cols=*/false);
  // Far more items than the field has values, none of them hidden.
  for (std::size_t i = 0; i < 500; ++i) {
    listed.mutable_fields()[0].items.push_back(PivotItem{"item_" + std::to_string(i), /*visible=*/true});
  }
  listed.mutable_fields()[0].items.push_back(PivotItem{"North", /*visible=*/true});
  listed.mutable_fields()[0].items.push_back(PivotItem{"South", /*visible=*/true});
  auto listed_or = evaluate(listed, cache);
  ASSERT_TRUE(static_cast<bool>(listed_or)) << listed_or.error().message;

  const PivotResult& baseline = baseline_or.value();
  const PivotResult& with_list = listed_or.value();
  ASSERT_EQ(with_list.rows.size(), baseline.rows.size());
  for (std::size_t i = 0; i < baseline.rows.size(); ++i) {
    EXPECT_EQ(with_list.rows[i].label, baseline.rows[i].label);
    EXPECT_DOUBLE_EQ(with_list.values[i][0][0].as_number(), baseline.values[i][0][0].as_number());
  }
}

// One field can hide the blank and a named value at once. The two are decided
// by different rules — the blank by the cache value its unlabelled item binds
// to, a named item by the label the grid draws — so a field carrying both has
// to apply both.
TEST(PivotEvaluator, OneFieldHidesTheBlankAndANamedItemTogether) {
  PivotCache cache = build_shared_region_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  PivotItem hidden_blank;
  hidden_blank.visible = false;
  hidden_blank.has_cache_index = true;
  hidden_blank.cache_index = kSharedRegionBlank;
  table.mutable_fields()[0].items.push_back(hidden_blank);
  table.mutable_fields()[0].items.push_back(PivotItem{"North", /*visible=*/false});
  resolve_pivot_names(table, cache);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Only the text value spelled like the placeholder survives.
  const PivotLayoutOptions defaults;
  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, defaults.blank_item_label);
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 7.0);
}

// The blank rule keys off the binding, not off the missing label: an unlabelled
// hidden item bound to a value that is not blank hides that value, labelled
// from its binding, and leaves the blank group alone.
TEST(PivotEvaluator, AnUnlabelledHiddenItemBoundToAValueHidesThatValue) {
  PivotCache cache = build_shared_region_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  PivotItem unlabelled;
  unlabelled.visible = false;
  unlabelled.has_cache_index = true;
  unlabelled.cache_index = 0U;  // "North", a value that renders to a label.
  table.mutable_fields()[0].items.push_back(std::move(unlabelled));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // North is gone; the blank group and the text value spelled like the
  // placeholder both survive. They draw the same label, so they are counted
  // through the total rather than looked up by it.
  ASSERT_EQ(r.rows.size(), 2U);
  EXPECT_EQ(row_index(r, "North"), static_cast<std::size_t>(-1));
  double total = 0.0;
  for (std::size_t i = 0; i < r.rows.size(); ++i) {
    total += r.values[i][0][0].as_number();
  }
  EXPECT_DOUBLE_EQ(total, 10.0 + 7.0);
}

// A hidden numeric item is matched by the label the grid draws for it, and that
// stays true when the field's item list is long: the record's value is rendered
// against the set of hidden labels, not against each item in turn.
TEST(PivotEvaluator, ManualFilterMatchesANumericLabelAmongManyItems) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  // 2,000 visible numeric items the cache has no record for, plus the one that
  // matches. Only the last one may prune anything.
  for (std::size_t i = 1000; i < 3000; ++i) {
    table.mutable_fields()[2].items.push_back(PivotItem{std::to_string(i), /*visible=*/true});
  }
  table.mutable_fields()[2].items.push_back(PivotItem{"300", /*visible=*/false});

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  const std::size_t south = row_index(r, "South");
  ASSERT_NE(south, static_cast<std::size_t>(-1));
  // South loses its 300 record and keeps the 200 one.
  EXPECT_DOUBLE_EQ(r.values[south][0][0].as_number(), 200.0);
}

// What a field hides is a property of the table, so deciding one record costs
// one lookup per field that hides something — never one comparison per item.
//
// Pinned by comparing two evaluations that differ only in how long the item
// list is. Both do the same hierarchy and aggregation work, so the difference
// between them is the filter pass alone, and reading it as a ratio keeps the
// assertion independent of how fast the host is. Matching each record against
// each item would render every record's label once per item, which is three
// orders of magnitude more work on the long list; building the hidden-label set
// once is paid per table, so the two runs stay within a small factor.
TEST(PivotEvaluator, ManualFilterCostDoesNotFollowTheItemListLength) {
  constexpr std::size_t kRecords = 10000;

  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Serial", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});
  for (std::size_t i = 0; i < kRecords; ++i) {
    PivotCacheRecord rec;
    // Non-integral so the label goes through the general numeric rendering.
    rec.cells = {Value::number(static_cast<double>(i) + 0.25), Value::number(1.0)};
    cache.mutable_records().push_back(std::move(rec));
  }

  // Every item is hidden and none of them names a value any record carries, so
  // the item list prunes nothing however long it is.
  const auto build = [](std::size_t item_count) {
    PivotTable table;
    table.set_pivot_cache_id(1);
    PivotField serial_field;
    serial_field.source_name = "Serial";
    serial_field.axis = PivotAxis::Row;
    for (std::size_t i = 0; i < item_count; ++i) {
      serial_field.items.push_back(
          PivotItem{display_string(Value::number(-static_cast<double>(i) - 0.5)), /*visible=*/false});
    }
    table.mutable_fields().push_back(std::move(serial_field));
    PivotField amount_field;
    amount_field.source_name = "Amount";
    amount_field.axis = PivotAxis::Value;
    table.mutable_fields().push_back(std::move(amount_field));
    PivotDataField count;
    count.name = "Count of Amount";
    count.field_index = 1;
    count.aggregation = Aggregation::Count;
    table.mutable_data_fields().push_back(std::move(count));
    table.mutable_row_field_order() = {0};
    table.set_grand_totals(/*rows=*/false, /*cols=*/false);
    return table;
  };
  const PivotTable short_list = build(20);
  const PivotTable long_list = build(20000);

  const auto timed = [&](const PivotTable& table) {
    const auto started = std::chrono::steady_clock::now();
    auto r_or = evaluate(table, cache);
    const double seconds = std::chrono::duration<double>(std::chrono::steady_clock::now() - started).count();
    EXPECT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
    EXPECT_EQ(r_or.value().rows.size(), kRecords);
    return seconds;
  };
  const double short_seconds = timed(short_list);
  const double long_seconds = timed(long_list);

  // The allowance is deliberately loose — a factor of ten plus half a second —
  // because it is separating a constant from a thousandfold, not measuring the
  // machine.
  EXPECT_LT(long_seconds, short_seconds * 10.0 + 0.5)
      << "short list " << short_seconds << "s, long list " << long_seconds << "s";
}

// ---------------------------------------------------------------------------
// LabelDate filter (date-range, pre-aggregation)
// ---------------------------------------------------------------------------

// Helper: builds a 2-column cache (Date as numeric serial, Amount). The
// records are picked so that a closed [2024-01-01, 2024-12-31] window
// keeps the first three and drops the fourth.
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

// ---------------------------------------------------------------------------
// ValueBetween filter (post-aggregation, parallel to Top-N / GreaterThan)
// ---------------------------------------------------------------------------

// Helper: builds a 2-column cache where each row label has a distinct
// numeric amount. With one record per row the post-aggregation row score
// equals the record's Amount.
PivotCache build_amount_cache_for_between() {
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Region", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});
  auto add = [&](const char* region, double amount) {
    PivotCacheRecord rec;
    rec.cells.push_back(owned_text(cache, region));
    rec.cells.push_back(Value::number(amount));
    cache.mutable_records().push_back(std::move(rec));
  };
  add("A", 10.0);
  add("B", 25.0);
  add("C", 40.0);
  add("D", 75.0);
  return cache;
}

TEST(PivotEvaluator, ValueBetweenFilterRowAxis) {
  PivotCache cache = build_amount_cache_for_between();
  PivotTable table;
  table.set_pivot_cache_id(1);
  PivotField rf;
  rf.source_name = "Region";
  rf.axis = PivotAxis::Row;
  PivotField af;
  af.source_name = "Amount";
  af.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(rf));
  table.mutable_fields().push_back(std::move(af));
  table.mutable_row_field_order() = {0};
  PivotDataField sum;
  sum.name = "Sum of Amount";
  sum.field_index = 1;
  sum.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum));
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Region";
  f.type = FilterType::ValueBetween;
  f.value = 20.0;       // inclusive low bound
  f.value_high = 50.0;  // inclusive high bound
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // 25 and 40 land in [20, 50]; 10 and 75 fall outside.
  ASSERT_EQ(r.rows.size(), 2U);
  std::vector<std::string> labels;
  std::vector<double> totals;
  for (std::size_t i = 0; i < r.rows.size(); ++i) {
    labels.push_back(r.rows[i].label);
    totals.push_back(r.values[i][0][0].as_number());
  }
  std::sort(labels.begin(), labels.end());
  std::sort(totals.begin(), totals.end());
  EXPECT_EQ(labels[0], "B");
  EXPECT_EQ(labels[1], "C");
  EXPECT_DOUBLE_EQ(totals[0], 25.0);
  EXPECT_DOUBLE_EQ(totals[1], 40.0);
}

TEST(PivotEvaluator, ValueBetweenFilterColAxis) {
  // Mirror of the row-axis test along the column axis: each Region
  // becomes a column with its single Amount as the column total.
  PivotCache cache = build_amount_cache_for_between();
  PivotTable table;
  table.set_pivot_cache_id(1);
  PivotField rf;
  rf.source_name = "Region";
  rf.axis = PivotAxis::Col;
  PivotField af;
  af.source_name = "Amount";
  af.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(rf));
  table.mutable_fields().push_back(std::move(af));
  table.mutable_col_field_order() = {0};
  PivotDataField sum;
  sum.name = "Sum of Amount";
  sum.field_index = 1;
  sum.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum));
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  PivotFilter f;
  f.axis = PivotAxis::Col;
  f.field_name = "Region";
  f.type = FilterType::ValueBetween;
  f.value = 20.0;
  f.value_high = 50.0;
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  ASSERT_EQ(r.cols.size(), 2U);
  std::vector<std::string> labels{r.cols[0].label, r.cols[1].label};
  std::sort(labels.begin(), labels.end());
  EXPECT_EQ(labels[0], "B");
  EXPECT_EQ(labels[1], "C");
  // The single implicit row should retain exactly the two surviving cols.
  ASSERT_EQ(r.values.size(), 1U);
  ASSERT_EQ(r.values[0].size(), 2U);
}

TEST(PivotEvaluator, ValueBetweenFilterUnboundedHighIsNoOp) {
  PivotCache cache = build_amount_cache_for_between();
  PivotTable table;
  table.set_pivot_cache_id(1);
  PivotField rf;
  rf.source_name = "Region";
  rf.axis = PivotAxis::Row;
  PivotField af;
  af.source_name = "Amount";
  af.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(rf));
  table.mutable_fields().push_back(std::move(af));
  table.mutable_row_field_order() = {0};
  PivotDataField sum;
  sum.name = "Sum of Amount";
  sum.field_index = 1;
  sum.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum));
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Region";
  f.type = FilterType::ValueBetween;
  f.value = 20.0;
  // value_high left as default monostate -> filter degrades to no-op.
  f.value_high = std::monostate{};
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // No upper bound -> all four rows survive.
  ASSERT_EQ(r.rows.size(), 4U);
}

// ---------------------------------------------------------------------------
// Multi-level value filters (Top-N / GreaterThan / Between) over hierarchies
// ---------------------------------------------------------------------------

TEST(PivotEvaluator, ValueTop10MultiLevelRowAxis) {
  // Two row fields: Region (North/South) -> Product (A/B/C). Each
  // (region, product) leaf has a single record so the post-aggregation
  // leaf score equals its Amount.
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Region", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Product", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});
  auto add = [&](const char* region, const char* product, double amount) {
    PivotCacheRecord rec;
    rec.cells.push_back(owned_text(cache, region));
    rec.cells.push_back(owned_text(cache, product));
    rec.cells.push_back(Value::number(amount));
    cache.mutable_records().push_back(std::move(rec));
  };
  add("North", "A", 10.0);
  add("North", "B", 50.0);
  add("North", "C", 20.0);
  add("South", "A", 80.0);
  add("South", "B", 30.0);
  add("South", "C", 5.0);

  PivotTable table;
  table.set_pivot_cache_id(1);
  PivotField region_f;
  region_f.source_name = "Region";
  region_f.axis = PivotAxis::Row;
  PivotField product_f;
  product_f.source_name = "Product";
  product_f.axis = PivotAxis::Row;
  PivotField amount_f;
  amount_f.source_name = "Amount";
  amount_f.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(region_f));
  table.mutable_fields().push_back(std::move(product_f));
  table.mutable_fields().push_back(std::move(amount_f));
  table.mutable_row_field_order() = {0, 1};
  PivotDataField sum;
  sum.name = "Sum of Amount";
  sum.field_index = 2;
  sum.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum));
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Product";
  f.type = FilterType::ValueTop10;
  f.value = 2;  // Top 2 products under each region.
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Ranked within each region, not across the flattened axis: North keeps
  // B=50 and C=20 even though South/B=30 outranks North/C globally, and
  // South keeps A=80 and B=30.
  ASSERT_EQ(r.rows.size(), 2U);
  ASSERT_EQ(r.rows[0].label, "North");
  ASSERT_EQ(r.rows[0].children.size(), 2U);
  EXPECT_EQ(r.rows[0].children[0].label, "B");
  EXPECT_EQ(r.rows[0].children[1].label, "C");
  ASSERT_EQ(r.rows[1].label, "South");
  ASSERT_EQ(r.rows[1].children.size(), 2U);
  EXPECT_EQ(r.rows[1].children[0].label, "A");
  EXPECT_EQ(r.rows[1].children[1].label, "B");

  ASSERT_EQ(r.values.size(), 4U);
  const std::vector<double> expected = {50.0, 20.0, 80.0, 30.0};
  for (std::size_t leaf = 0; leaf < expected.size(); ++leaf) {
    ASSERT_EQ(r.values[leaf].size(), 1U);
    EXPECT_DOUBLE_EQ(r.values[leaf][0][0].as_number(), expected[leaf]) << "leaf=" << leaf;
  }
}

// Region -> Product -> Amount, one record per `(region, product, amount)`,
// for the value-filter depth tests below.
struct RegionProductAmount {
  const char* region;
  const char* product;
  double amount;
};

PivotCache build_region_product_cache(const std::vector<RegionProductAmount>& records) {
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Region", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Product", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});
  for (const RegionProductAmount& record : records) {
    PivotCacheRecord rec;
    rec.cells.push_back(owned_text(cache, record.region));
    rec.cells.push_back(owned_text(cache, record.product));
    rec.cells.push_back(Value::number(record.amount));
    cache.mutable_records().push_back(std::move(rec));
  }
  return cache;
}

PivotTable build_region_product_table(Aggregation aggregation) {
  PivotTable table;
  table.set_pivot_cache_id(1);
  const char* names[] = {"Region", "Product", "Amount"};
  for (std::size_t i = 0; i < 3; ++i) {
    PivotField field;
    field.source_name = names[i];
    field.axis = i == 2 ? PivotAxis::Value : PivotAxis::Row;
    table.mutable_fields().push_back(std::move(field));
  }
  table.mutable_row_field_order() = {0, 1};
  PivotDataField data;
  data.name = "Amount";
  data.field_index = 2;
  data.aggregation = aggregation;
  table.mutable_data_fields().push_back(std::move(data));
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  return table;
}

// A filter on the outer field ranks whole regions by their own total and
// keeps every product under a surviving region.
TEST(PivotEvaluator, ValueTop10OnOuterFieldKeepsWholeGroups) {
  PivotCache cache = build_region_product_cache({{"North", "A", 10.0},
                                                 {"North", "B", 50.0},
                                                 {"North", "C", 20.0},
                                                 {"South", "A", 80.0},
                                                 {"South", "B", 30.0},
                                                 {"South", "C", 5.0}});
  PivotTable table = build_region_product_table(Aggregation::Sum);
  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Region";
  f.type = FilterType::ValueTop10;
  f.value = 1;  // North=80, South=115.
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, "South");
  EXPECT_EQ(r.rows[0].children.size(), 3U);
  EXPECT_EQ(r.values.size(), 3U);
}

// The per-parent partition is structural: turning the outer field's
// subtotals off must not merge its groups back into one global ranking.
TEST(PivotEvaluator, ValueTop10OnInnerFieldRanksPerParentWithoutSubtotals) {
  PivotCache cache = build_region_product_cache(
      {{"North", "A", 10.0}, {"North", "B", 50.0}, {"South", "A", 80.0}, {"South", "B", 30.0}});
  PivotTable table = build_region_product_table(Aggregation::Sum);
  table.mutable_fields()[0].default_subtotal = false;
  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Product";
  f.type = FilterType::ValueTop10;
  f.value = 1;
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  EXPECT_TRUE(r.row_subtotals.empty());
  ASSERT_EQ(r.rows.size(), 2U);
  ASSERT_EQ(r.rows[0].children.size(), 1U);
  EXPECT_EQ(r.rows[0].children[0].label, "B");  // North's best, below South/B.
  ASSERT_EQ(r.rows[1].children.size(), 1U);
  EXPECT_EQ(r.rows[1].children[0].label, "A");
}

// An authored `<filters>` entry names its field by index; the pass ranks at
// that field's depth the same way it does for a named filter.
TEST(PivotEvaluator, AuthoredTopCountOnInnerFieldRanksPerParent) {
  PivotCache cache = build_region_product_cache(
      {{"North", "A", 10.0}, {"North", "B", 50.0}, {"South", "A", 80.0}, {"South", "B", 30.0}});
  PivotTable table = build_region_product_table(Aggregation::Sum);
  AuthoredValueFilter f;
  f.field_index = 1;
  f.type = FilterType::ValueTop10;
  f.value = 1.0;
  table.mutable_authored_value_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.rows.size(), 2U);
  ASSERT_EQ(r.rows[0].children.size(), 1U);
  EXPECT_EQ(r.rows[0].children[0].label, "B");
  ASSERT_EQ(r.rows[1].children.size(), 1U);
  EXPECT_EQ(r.rows[1].children[0].label, "A");
}

// A group is scored by its own aggregate, not by summing its leaves: under
// Average, South (50) outranks North (40) although North's leaves sum higher.
TEST(PivotEvaluator, ValueTop10OnOuterFieldScoresNonAdditiveAggregate) {
  PivotCache cache = build_region_product_cache(
      {{"North", "A", 40.0}, {"North", "B", 40.0}, {"North", "C", 40.0}, {"South", "A", 50.0}});
  PivotTable table = build_region_product_table(Aggregation::Average);
  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Region";
  f.type = FilterType::ValueTop10;
  f.value = 1;
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, "South");
}

TEST(PivotEvaluator, ValueGreaterThanMultiLevelColAxis) {
  // Two column fields: Year (2024/2025) -> Quarter (Q1/Q2). Single
  // implicit row, so each col leaf's score is just its Amount.
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Year", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Quarter", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});
  auto add = [&](const char* year, const char* quarter, double amount) {
    PivotCacheRecord rec;
    rec.cells.push_back(owned_text(cache, year));
    rec.cells.push_back(owned_text(cache, quarter));
    rec.cells.push_back(Value::number(amount));
    cache.mutable_records().push_back(std::move(rec));
  };
  // Leaf order in DFS pre-order: 2024/Q1=15, 2024/Q2=40, 2025/Q1=5, 2025/Q2=100.
  add("2024", "Q1", 15.0);
  add("2024", "Q2", 40.0);
  add("2025", "Q1", 5.0);
  add("2025", "Q2", 100.0);

  PivotTable table;
  table.set_pivot_cache_id(1);
  PivotField year_f;
  year_f.source_name = "Year";
  year_f.axis = PivotAxis::Col;
  PivotField quarter_f;
  quarter_f.source_name = "Quarter";
  quarter_f.axis = PivotAxis::Col;
  PivotField amount_f;
  amount_f.source_name = "Amount";
  amount_f.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(year_f));
  table.mutable_fields().push_back(std::move(quarter_f));
  table.mutable_fields().push_back(std::move(amount_f));
  table.mutable_col_field_order() = {0, 1};
  PivotDataField sum;
  sum.name = "Sum of Amount";
  sum.field_index = 2;
  sum.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum));
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  PivotFilter f;
  f.axis = PivotAxis::Col;
  f.field_name = "Quarter";
  f.type = FilterType::ValueGreaterThan;
  f.value = 20.0;
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Surviving col leaves at indices 1 (2024/Q2=40) and 3 (2025/Q2=100).
  // Each Year keeps only Q2; both Year subtrees survive.
  ASSERT_EQ(r.cols.size(), 2U);
  EXPECT_EQ(r.cols[0].label, "2024");
  EXPECT_EQ(r.cols[1].label, "2025");
  ASSERT_EQ(r.cols[0].children.size(), 1U);
  ASSERT_EQ(r.cols[1].children.size(), 1U);
  EXPECT_EQ(r.cols[0].children[0].label, "Q2");
  EXPECT_EQ(r.cols[1].children[0].label, "Q2");

  ASSERT_EQ(r.values.size(), 1U);
  ASSERT_EQ(r.values[0].size(), 2U);
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 40.0);
  EXPECT_DOUBLE_EQ(r.values[0][1][0].as_number(), 100.0);
}

TEST(PivotEvaluator, ValueBetweenMultiLevelDropsEmptyParent) {
  // Two row fields: Region (North/South) -> Product (A/B). All North
  // leaves fall outside [10, 70], so North gets pruned entirely.
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Region", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Product", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});
  auto add = [&](const char* region, const char* product, double amount) {
    PivotCacheRecord rec;
    rec.cells.push_back(owned_text(cache, region));
    rec.cells.push_back(owned_text(cache, product));
    rec.cells.push_back(Value::number(amount));
    cache.mutable_records().push_back(std::move(rec));
  };
  add("North", "A", 5.0);
  add("North", "B", 8.0);
  add("South", "A", 50.0);
  add("South", "B", 60.0);

  PivotTable table;
  table.set_pivot_cache_id(1);
  PivotField region_f;
  region_f.source_name = "Region";
  region_f.axis = PivotAxis::Row;
  PivotField product_f;
  product_f.source_name = "Product";
  product_f.axis = PivotAxis::Row;
  PivotField amount_f;
  amount_f.source_name = "Amount";
  amount_f.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(region_f));
  table.mutable_fields().push_back(std::move(product_f));
  table.mutable_fields().push_back(std::move(amount_f));
  table.mutable_row_field_order() = {0, 1};
  PivotDataField sum;
  sum.name = "Sum of Amount";
  sum.field_index = 2;
  sum.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum));
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Product";
  f.type = FilterType::ValueBetween;
  f.value = 10.0;
  f.value_high = 70.0;
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // North has no surviving children -> pruned entirely. Only South
  // remains, with both A and B.
  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, "South");
  ASSERT_EQ(r.rows[0].children.size(), 2U);
  EXPECT_EQ(r.rows[0].children[0].label, "A");
  EXPECT_EQ(r.rows[0].children[1].label, "B");

  ASSERT_EQ(r.values.size(), 2U);
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 50.0);
  EXPECT_DOUBLE_EQ(r.values[1][0][0].as_number(), 60.0);
}

// ---------------------------------------------------------------------------
// Label filters resolve a field by its display name
// ---------------------------------------------------------------------------

TEST(PivotEvaluator, LabelFilterResolvesFieldByCustomName) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.mutable_fields()[0].custom_name = "Area";

  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Area";  // The display name, not the source name.
  f.type = FilterType::LabelBeginsWith;
  f.value = std::string("N");
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Only the North records survive, so both the axis and the aggregate move.
  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, "North");
  ASSERT_EQ(r.values.size(), 1U);
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 175.0);  // 100 + 50 + 25
  ASSERT_TRUE(r.grand_total.is_number());
  EXPECT_DOUBLE_EQ(r.grand_total.as_number(), 175.0);
}

TEST(PivotEvaluator, LabelFilterResolvesFieldByDataFieldName) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});

  // "Sum of Amount" is the data field's display name for the Amount field,
  // so the filter applies to Amount's own values.
  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Sum of Amount";
  f.type = FilterType::LabelBeginsWith;
  f.value = std::string("1");
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Only the Amount=100 record starts with "1".
  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, "North");
  ASSERT_TRUE(r.grand_total.is_number());
  EXPECT_DOUBLE_EQ(r.grand_total.as_number(), 100.0);
}

// ---------------------------------------------------------------------------
// Page (report filter) axis selection.
// ---------------------------------------------------------------------------
//
// The selection is named during evaluation rather than by the projection
// because it comes from the bound cache, which `layout` is not handed. Only
// the unfiltered spelling is oracle-measured (`(すべて)` under the ja-JP
// profile); these pin the state machine that picks between the three
// outcomes.

// Region on the page axis, Product on the row axis.
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
