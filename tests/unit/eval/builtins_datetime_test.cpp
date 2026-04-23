// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for the date/time built-ins: DATE, TIME, YEAR, MONTH,
// DAY, HOUR, MINUTE, SECOND, WEEKDAY, EDATE, EOMONTH, DAYS. Each test
// parses a formula source, evaluates the AST through the default
// registry, and asserts the resulting Value.
//
// Serial values used in the assertions below were cross-checked by
// independently running Howard Hinnant's `days_from_civil` algorithm and
// adding the +25569 post-ghost / +25568 pre-ghost offsets (see
// `src/eval/date_time.cpp`). Example landmarks:
//
//   DATE(1900, 1, 1)  =     1   (first Excel serial, pre-ghost)
//   DATE(1900, 2, 28) =    59   (last pre-ghost real day)
//   DATE(1900, 3, 1)  =    61   (first post-ghost day -- serial 60 is the
//                                fictional 1900-02-29 preserved from
//                                Lotus 1-2-3's leap-year bug)
//   DATE(2024, 2, 29) = 45351   (actual leap day, 2024 leap year)
//   DATE(2026, 4, 23) = 46135   (today's date per CLAUDE.md currentDate)

#include <cmath>
#include <string_view>

#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Parses `src` and evaluates it via the default function registry. The
// thread-local arenas keep text payloads readable for the immediately
// following EXPECT_*. Each call resets the arenas to avoid cross-test
// contamination.
Value EvalSource(std::string_view src) {
  static thread_local Arena parse_arena;
  static thread_local Arena eval_arena;
  parse_arena.reset();
  eval_arena.reset();
  parser::Parser p(src, parse_arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return evaluate(*root, eval_arena);
}

// ---------------------------------------------------------------------------
// DATE
// ---------------------------------------------------------------------------

TEST(DateTimeDate, CurrentDateSerial) {
  const Value v = EvalSource("=DATE(2026, 4, 23)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 46135.0);
}

TEST(DateTimeDate, FirstSerial) {
  const Value v = EvalSource("=DATE(1900, 1, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(DateTimeDate, LastPreGhostDay) {
  const Value v = EvalSource("=DATE(1900, 2, 28)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 59.0);
}

TEST(DateTimeDate, FirstPostGhostDay) {
  // Serial 60 is Excel's fictional 1900-02-29 (Lotus leap-year bug);
  // the next real day, 1900-03-01, is serial 61.
  const Value v = EvalSource("=DATE(1900, 3, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 61.0);
}

TEST(DateTimeDate, ActualLeapDay2024) {
  const Value v = EvalSource("=DATE(2024, 2, 29)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 45351.0);
}

TEST(DateTimeDate, TwoDigitYearExpansion) {
  // Excel's documented rule: `0 <= year < 1900` adds 1900. So 26 -> 1926,
  // NOT 2026 (Excel does not infer a 20xx pivot for two-digit years inside
  // DATE, unlike some parsers). See Excel function reference for DATE.
  const Value a = EvalSource("=DATE(26, 4, 23)");
  const Value b = EvalSource("=DATE(1926, 4, 23)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_EQ(a.as_number(), b.as_number());
}

TEST(DateTimeDate, ZeroYearExpandsTo1900) {
  // `DATE(0, 1, 1)` is `DATE(1900, 1, 1)` = 1.
  const Value v = EvalSource("=DATE(0, 1, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(DateTimeDate, MonthOverflowRollsYear) {
  const Value a = EvalSource("=DATE(2026, 13, 1)");
  const Value b = EvalSource("=DATE(2027, 1, 1)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_EQ(a.as_number(), b.as_number());
}

TEST(DateTimeDate, MonthUnderflowRollsYear) {
  // Month 0 -> December of previous year.
  const Value a = EvalSource("=DATE(2026, 0, 15)");
  const Value b = EvalSource("=DATE(2025, 12, 15)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_EQ(a.as_number(), b.as_number());
}

TEST(DateTimeDate, DayOverflowRollsMonth) {
  // Feb 30 in a non-leap year -> March 2.
  const Value a = EvalSource("=DATE(2026, 2, 30)");
  const Value b = EvalSource("=DATE(2026, 3, 2)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_EQ(a.as_number(), b.as_number());
}

TEST(DateTimeDate, NegativeMonthSubtracts) {
  const Value a = EvalSource("=DATE(2026, -1, 1)");
  const Value b = EvalSource("=DATE(2025, 11, 1)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_EQ(a.as_number(), b.as_number());
}

TEST(DateTimeDate, Year9999Accepted) {
  const Value v = EvalSource("=DATE(9999, 12, 31)");
  ASSERT_TRUE(v.is_number());
  EXPECT_GT(v.as_number(), 0.0);
}

TEST(DateTimeDate, Year10000Rejected) {
  const Value v = EvalSource("=DATE(10000, 1, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(DateTimeDate, NegativeYearRejected) {
  const Value v = EvalSource("=DATE(-1, 1, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(DateTimeDate, NonNumericYearPropagatesValueError) {
  const Value v = EvalSource("=DATE(\"abc\", 1, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// TIME
// ---------------------------------------------------------------------------

TEST(DateTimeTime, NoonIsHalf) {
  const Value v = EvalSource("=TIME(12, 0, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.5);
}

TEST(DateTimeTime, OneSecond) {
  const Value v = EvalSource("=TIME(0, 0, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0 / 86400.0, 1e-12);
}

TEST(DateTimeTime, EndOfDay) {
  const Value v = EvalSource("=TIME(23, 59, 59)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 86399.0 / 86400.0, 1e-12);
}

TEST(DateTimeTime, HourOverflowWrapsModuloDay) {
  // TIME(25, 0, 0) == TIME(1, 0, 0) modulo the last ULP: fmod on the
  // scaled-seconds path introduces a few ULPs of rounding versus the
  // direct 3600/86400 division, so compare with a small tolerance.
  const Value a = EvalSource("=TIME(25, 0, 0)");
  const Value b = EvalSource("=TIME(1, 0, 0)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_NEAR(a.as_number(), b.as_number(), 1e-12);
}

TEST(DateTimeTime, MinuteOverflowNormalises) {
  // TIME(1, 60, 0) == TIME(2, 0, 0) (same ULP caveat as above).
  const Value a = EvalSource("=TIME(1, 60, 0)");
  const Value b = EvalSource("=TIME(2, 0, 0)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_NEAR(a.as_number(), b.as_number(), 1e-12);
}

TEST(DateTimeTime, NegativeComponentProducesNum) {
  const Value v = EvalSource("=TIME(-1, 0, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// YEAR / MONTH / DAY
// ---------------------------------------------------------------------------

TEST(DateTimeYear, ExtractsFromKnownSerial) {
  const Value v = EvalSource("=YEAR(DATE(2026, 4, 23))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2026.0);
}

TEST(DateTimeMonth, ExtractsFromKnownSerial) {
  const Value v = EvalSource("=MONTH(DATE(2026, 4, 23))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 4.0);
}

TEST(DateTimeDay, ExtractsFromKnownSerial) {
  const Value v = EvalSource("=DAY(DATE(2026, 4, 23))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 23.0);
}

TEST(DateTimeYear, GhostDaySerial60) {
  // Serial 60 is the fictional 1900-02-29: Excel still reports year 1900,
  // month 2, day 29, so the engine matches.
  const Value y = EvalSource("=YEAR(60)");
  const Value m = EvalSource("=MONTH(60)");
  const Value d = EvalSource("=DAY(60)");
  ASSERT_TRUE(y.is_number());
  ASSERT_TRUE(m.is_number());
  ASSERT_TRUE(d.is_number());
  EXPECT_EQ(y.as_number(), 1900.0);
  EXPECT_EQ(m.as_number(), 2.0);
  EXPECT_EQ(d.as_number(), 29.0);
}

TEST(DateTimeYear, IgnoresTimeFraction) {
  const Value v = EvalSource("=YEAR(DATE(2026, 4, 23) + 0.75)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2026.0);
}

TEST(DateTimeYear, NegativeSerialIsNum) {
  const Value v = EvalSource("=YEAR(-1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(DateTimeYear, NonNumericIsValueError) {
  const Value v = EvalSource("=YEAR(\"abc\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// HOUR / MINUTE / SECOND
// ---------------------------------------------------------------------------

TEST(DateTimeHour, ExtractsFromDatePlusTime) {
  const Value v = EvalSource("=HOUR(DATE(2026, 4, 23) + TIME(14, 30, 15))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 14.0);
}

TEST(DateTimeMinute, ExtractsFromDatePlusTime) {
  const Value v = EvalSource("=MINUTE(DATE(2026, 4, 23) + TIME(14, 30, 15))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 30.0);
}

TEST(DateTimeSecond, ExtractsFromDatePlusTime) {
  const Value v = EvalSource("=SECOND(DATE(2026, 4, 23) + TIME(14, 30, 15))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 15.0);
}

TEST(DateTimeHour, NoonIs12) {
  const Value v = EvalSource("=HOUR(0.5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 12.0);
}

TEST(DateTimeHour, MidnightIs0) {
  const Value v = EvalSource("=HOUR(0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(DateTimeSecond, NearFullDayRoundsToZero) {
  // 0.999999999 of a day is 86399.9999136 seconds; rounded to the nearest
  // second this becomes 86400, which wraps back to 0:0:0 of the next day.
  const Value h = EvalSource("=HOUR(0.999999999)");
  const Value m = EvalSource("=MINUTE(0.999999999)");
  const Value s = EvalSource("=SECOND(0.999999999)");
  ASSERT_TRUE(h.is_number());
  ASSERT_TRUE(m.is_number());
  ASSERT_TRUE(s.is_number());
  EXPECT_EQ(h.as_number(), 0.0);
  EXPECT_EQ(m.as_number(), 0.0);
  EXPECT_EQ(s.as_number(), 0.0);
}

TEST(DateTimeMinute, NegativeSerialIsNum) {
  const Value v = EvalSource("=MINUTE(-1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// WEEKDAY
// ---------------------------------------------------------------------------

TEST(DateTimeWeekday, DefaultTypeSundayBased) {
  // 2026-04-23 is a Thursday; type 1 returns Sun=1..Sat=7, so Thu = 5.
  const Value v = EvalSource("=WEEKDAY(DATE(2026, 4, 23))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 5.0);
}

TEST(DateTimeWeekday, Type2MondayBased) {
  const Value v = EvalSource("=WEEKDAY(DATE(2026, 4, 23), 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 4.0);
}

TEST(DateTimeWeekday, Type3MondayZero) {
  const Value v = EvalSource("=WEEKDAY(DATE(2026, 4, 23), 3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 3.0);
}

TEST(DateTimeWeekday, Type11MondayStart) {
  // Type 11 starts the week on Monday, returns 1..7. Thursday -> 4.
  const Value v = EvalSource("=WEEKDAY(DATE(2026, 4, 23), 11)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 4.0);
}

TEST(DateTimeWeekday, Type14ThursdayStart) {
  // Type 14 starts the week on Thursday, so Thursday itself -> 1.
  const Value v = EvalSource("=WEEKDAY(DATE(2026, 4, 23), 14)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(DateTimeWeekday, Type16SaturdayStart) {
  // Type 16 starts on Saturday. Thursday is 6 days after Saturday -> 6.
  const Value v = EvalSource("=WEEKDAY(DATE(2026, 4, 23), 16)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 6.0);
}

TEST(DateTimeWeekday, Type17SundayStart) {
  // Type 17 starts on Sunday. Thursday is 4 days after Sunday -> 5.
  const Value v = EvalSource("=WEEKDAY(DATE(2026, 4, 23), 17)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 5.0);
}

TEST(DateTimeWeekday, UnsupportedTypeIsNum) {
  const Value v = EvalSource("=WEEKDAY(DATE(2026, 4, 23), 5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// EDATE
// ---------------------------------------------------------------------------

TEST(DateTimeEdate, ForwardMonths) {
  // 2026-04-23 + 2 months -> 2026-06-23.
  const Value a = EvalSource("=EDATE(DATE(2026, 4, 23), 2)");
  const Value b = EvalSource("=DATE(2026, 6, 23)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_EQ(a.as_number(), b.as_number());
}

TEST(DateTimeEdate, BackwardMonths) {
  // 2026-04-23 - 5 months -> 2025-11-23.
  const Value a = EvalSource("=EDATE(DATE(2026, 4, 23), -5)");
  const Value b = EvalSource("=DATE(2025, 11, 23)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_EQ(a.as_number(), b.as_number());
}

TEST(DateTimeEdate, Jan31PlusOneMonthInLeapYear) {
  // 2024 is a leap year: Jan 31 + 1 month -> Feb 29.
  const Value a = EvalSource("=EDATE(DATE(2024, 1, 31), 1)");
  const Value b = EvalSource("=DATE(2024, 2, 29)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_EQ(a.as_number(), b.as_number());
}

TEST(DateTimeEdate, Jan31PlusOneMonthInNonLeapYear) {
  // 2023 is not a leap year: Jan 31 + 1 month -> Feb 28.
  const Value a = EvalSource("=EDATE(DATE(2023, 1, 31), 1)");
  const Value b = EvalSource("=DATE(2023, 2, 28)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_EQ(a.as_number(), b.as_number());
}

TEST(DateTimeEdate, CrossYearBackwards) {
  // 2026-01-15 - 13 months -> 2024-12-15.
  const Value a = EvalSource("=EDATE(DATE(2026, 1, 15), -13)");
  const Value b = EvalSource("=DATE(2024, 12, 15)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_EQ(a.as_number(), b.as_number());
}

TEST(DateTimeEdate, TruncatesMonthArgument) {
  // `months = 2.9` truncates to 2.
  const Value a = EvalSource("=EDATE(DATE(2026, 4, 23), 2.9)");
  const Value b = EvalSource("=DATE(2026, 6, 23)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_EQ(a.as_number(), b.as_number());
}

// ---------------------------------------------------------------------------
// EOMONTH
// ---------------------------------------------------------------------------

TEST(DateTimeEomonth, SameMonth) {
  // EOMONTH(2026-04-10, 0) -> 2026-04-30.
  const Value a = EvalSource("=EOMONTH(DATE(2026, 4, 10), 0)");
  const Value b = EvalSource("=DATE(2026, 4, 30)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_EQ(a.as_number(), b.as_number());
}

TEST(DateTimeEomonth, LeapYearFebruary) {
  // EOMONTH(2024-02-15, 0) -> 2024-02-29 (leap year).
  const Value a = EvalSource("=EOMONTH(DATE(2024, 2, 15), 0)");
  const Value b = EvalSource("=DATE(2024, 2, 29)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_EQ(a.as_number(), b.as_number());
}

TEST(DateTimeEomonth, NonLeapYearFebruary) {
  const Value a = EvalSource("=EOMONTH(DATE(2023, 2, 15), 0)");
  const Value b = EvalSource("=DATE(2023, 2, 28)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_EQ(a.as_number(), b.as_number());
}

TEST(DateTimeEomonth, ShiftsForward) {
  // EOMONTH(2026-04-10, 2) -> 2026-06-30.
  const Value a = EvalSource("=EOMONTH(DATE(2026, 4, 10), 2)");
  const Value b = EvalSource("=DATE(2026, 6, 30)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_EQ(a.as_number(), b.as_number());
}

TEST(DateTimeEomonth, ShiftsBackward) {
  // EOMONTH(2026-04-10, -1) -> 2026-03-31.
  const Value a = EvalSource("=EOMONTH(DATE(2026, 4, 10), -1)");
  const Value b = EvalSource("=DATE(2026, 3, 31)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_EQ(a.as_number(), b.as_number());
}

TEST(DateTimeEomonth, Century2100NotLeap) {
  // 2100 is divisible by 100 but not 400 -> NOT a leap year.
  const Value a = EvalSource("=EOMONTH(DATE(2100, 2, 15), 0)");
  const Value b = EvalSource("=DATE(2100, 2, 28)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_EQ(a.as_number(), b.as_number());
}

TEST(DateTimeEomonth, Century2000IsLeap) {
  // 2000 is divisible by 400 -> IS a leap year.
  const Value a = EvalSource("=EOMONTH(DATE(2000, 2, 15), 0)");
  const Value b = EvalSource("=DATE(2000, 2, 29)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_EQ(a.as_number(), b.as_number());
}

// ---------------------------------------------------------------------------
// DAYS
// ---------------------------------------------------------------------------

TEST(DateTimeDays, PositiveDiff) {
  const Value v = EvalSource("=DAYS(DATE(2026, 4, 23), DATE(2026, 4, 20))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 3.0);
}

TEST(DateTimeDays, ZeroDiff) {
  const Value v = EvalSource("=DAYS(DATE(2026, 4, 23), DATE(2026, 4, 23))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(DateTimeDays, NegativeDiff) {
  const Value v = EvalSource("=DAYS(DATE(2026, 4, 20), DATE(2026, 4, 23))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -3.0);
}

TEST(DateTimeDays, CrossYear) {
  // 2026-01-01 minus 2025-12-31 = 1 day.
  const Value v = EvalSource("=DAYS(DATE(2026, 1, 1), DATE(2025, 12, 31))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(DateTimeDays, IgnoresFractionalPart) {
  // DAYS floors both operands, so a fractional serial should not leak in.
  const Value v = EvalSource("=DAYS(DATE(2026, 4, 23) + 0.9, DATE(2026, 4, 20) + 0.1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 3.0);
}

TEST(DateTimeDays, ErrorPropagates) {
  const Value v = EvalSource("=DAYS(\"abc\", DATE(2026, 4, 23))");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// Round-trip sanity checks combining DATE with extractors.
// ---------------------------------------------------------------------------

TEST(DateTimeRoundTrip, DateExtractsItself) {
  const Value y = EvalSource("=YEAR(DATE(1995, 7, 4))");
  const Value m = EvalSource("=MONTH(DATE(1995, 7, 4))");
  const Value d = EvalSource("=DAY(DATE(1995, 7, 4))");
  ASSERT_TRUE(y.is_number());
  ASSERT_TRUE(m.is_number());
  ASSERT_TRUE(d.is_number());
  EXPECT_EQ(y.as_number(), 1995.0);
  EXPECT_EQ(m.as_number(), 7.0);
  EXPECT_EQ(d.as_number(), 4.0);
}

TEST(DateTimeRoundTrip, TimeExtractsItself) {
  const Value h = EvalSource("=HOUR(TIME(7, 15, 30))");
  const Value m = EvalSource("=MINUTE(TIME(7, 15, 30))");
  const Value s = EvalSource("=SECOND(TIME(7, 15, 30))");
  ASSERT_TRUE(h.is_number());
  ASSERT_TRUE(m.is_number());
  ASSERT_TRUE(s.is_number());
  EXPECT_EQ(h.as_number(), 7.0);
  EXPECT_EQ(m.as_number(), 15.0);
  EXPECT_EQ(s.as_number(), 30.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
