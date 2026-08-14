//
// End-to-end tests for the date/time built-ins: DATE, TIME, YEAR, MONTH,
// DAY, HOUR, MINUTE, SECOND, WEEKDAY, EDATE, EOMONTH, DAYS, WEEKNUM,
// ISOWEEKNUM, YEARFRAC, DATEDIF, NETWORKDAYS, WORKDAY. Each test parses a
// formula source, evaluates the AST through the default registry, and
// asserts the resulting Value. NETWORKDAYS and WORKDAY tests use a bound
// workbook so they can exercise range-sourced holiday lists.
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
#include <string>
#include <string_view>

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "test_eval_helpers.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

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
  return evaluate(*root, eval_arena, default_registry(), test::mac_context());
}

// Bound-workbook variant used by NETWORKDAYS / WORKDAY tests: lets A1-style
// refs in the formula resolve against the sheet's live cells so range-
// sourced holiday lists exercise the same path as real workbooks.
Value EvalSourceIn(std::string_view src, const Workbook& wb, const Sheet& current) {
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
  EvalState state;
  const EvalContext ctx(wb, current, state);
  return evaluate(*root, eval_arena, default_registry(), ctx);
}

// Evaluates `src` under the 1904 date system (the calendar family reads the
// epoch from `EvalContext::date1904()`).
Value EvalSource1904(std::string_view src) {
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
  return evaluate(*root, eval_arena, default_registry(), test::mac_context().with_date1904(true));
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
  EXPECT_DOUBLE_EQ(v.as_number(), 2958465.0);
}

TEST(DateTimeDate, DayOverflowNormalisesToUpperEndpoint) {
  const Value v = EvalSource("=DATE(9999, 11, 61)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2958465.0);
}

TEST(DateTimeDate, DayOverflowPastUpperEndpointIsNum) {
  const Value v = EvalSource("=DATE(9999, 12, 32)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(DateTimeDate, NegativeDayNormalisesToLowerBoundary) {
  const Value serial_one = EvalSource("=DATE(1901, 1, -364)");
  ASSERT_TRUE(serial_one.is_number());
  EXPECT_DOUBLE_EQ(serial_one.as_number(), 1.0);

  const Value serial_zero = EvalSource("=DATE(1901, 1, -365)");
  ASSERT_TRUE(serial_zero.is_number());
  EXPECT_DOUBLE_EQ(serial_zero.as_number(), 0.0);
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

TEST(DateTimeYear, BoundedRangeSpillsElementwise) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0U, 0U, Value::number(46135.0))));  // 2026-04-23
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 1U, 0U, Value::number(45351.0))));  // 2024-02-29
  const Value result = EvalSourceIn("=YEAR(A1:A2)", wb, wb.sheet(0));
  ASSERT_TRUE(result.is_array());
  ASSERT_EQ(result.as_array_rows(), 2U);
  ASSERT_EQ(result.as_array_cols(), 1U);
  EXPECT_EQ(result.as_array()->cells[0].as_number(), 2026.0);
  EXPECT_EQ(result.as_array()->cells[1].as_number(), 2024.0);
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

TEST(DateTimeUpperEndpoint, AcceptsFinalDayTimeFraction) {
  const Value year = EvalSource("=YEAR(2958465.999988426)");
  const Value month = EvalSource("=MONTH(2958465.999988426)");
  const Value day = EvalSource("=DAY(2958465.999988426)");
  ASSERT_TRUE(year.is_number());
  ASSERT_TRUE(month.is_number());
  ASSERT_TRUE(day.is_number());
  EXPECT_DOUBLE_EQ(year.as_number(), 9999.0);
  EXPECT_DOUBLE_EQ(month.as_number(), 12.0);
  EXPECT_DOUBLE_EQ(day.as_number(), 31.0);
}

TEST(DateTimeUpperEndpoint, RejectsExactNextDay) {
  const Value year = EvalSource("=YEAR(2958466)");
  const Value month = EvalSource("=MONTH(2958466)");
  const Value day = EvalSource("=DAY(2958466)");
  ASSERT_TRUE(year.is_error());
  ASSERT_TRUE(month.is_error());
  ASSERT_TRUE(day.is_error());
  EXPECT_EQ(year.as_error(), ErrorCode::Num);
  EXPECT_EQ(month.as_error(), ErrorCode::Num);
  EXPECT_EQ(day.as_error(), ErrorCode::Num);
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

TEST(DateTimeWeekday, PreGhostDaysUseGregorianWeekdays) {
  // Excel's serial weekday arithmetic deliberately reports the weekday one
  // day earlier before its fictitious 1900-02-29. Formulon keeps the real
  // Gregorian weekday instead; the verified intentional divergence lives in
  // tests/divergence.yaml as weekday_pre_1900_excel_serial_bug.
  const Value jan_1 = EvalSource("=WEEKDAY(DATE(1900, 1, 1), 2)");
  const Value feb_28 = EvalSource("=WEEKDAY(DATE(1900, 2, 28), 2)");
  ASSERT_TRUE(jan_1.is_number());
  ASSERT_TRUE(feb_28.is_number());
  EXPECT_EQ(jan_1.as_number(), 1.0);   // Monday
  EXPECT_EQ(feb_28.as_number(), 3.0);  // Wednesday
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

TEST(DateTimeEdate, UpperEndpointWithZeroMonths) {
  const Value v = EvalSource("=EDATE(DATE(9999, 12, 31), 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2958465.0);
}

TEST(DateTimeEdate, CrossingUpperEndpointIsNum) {
  const Value v = EvalSource("=EDATE(DATE(9999, 12, 1), 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(DateTimeEdate, InputOnePastUpperEndpointIsNum) {
  const Value v = EvalSource("=EDATE(DATE(9999, 12, 31) + 1, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(DateTimeEdate, AcceptsFinalDayTimeFraction) {
  const Value endpoint = EvalSource("=EDATE(2958465.999988426, 0)");
  ASSERT_TRUE(endpoint.is_number());
  EXPECT_DOUBLE_EQ(endpoint.as_number(), 2958465.0);
}

TEST(DateTimeEdate, FinalDayTimeFractionCannotCrossUpperEndpoint) {
  const Value overflow = EvalSource("=EDATE(2958465.999988426, 1)");
  ASSERT_TRUE(overflow.is_error());
  EXPECT_EQ(overflow.as_error(), ErrorCode::Num);
}

TEST(DateTimeEdate, GiganticMonthOffsetIsNum) {
  const Value v = EvalSource("=EDATE(DATE(2024, 1, 1), 1E20)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(DateTimeEdate, LeftmostErrorPrecedesGiganticMonthOffset) {
  const Value v = EvalSource("=EDATE(1/0, 1E20)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(DateTimeEdate, LowerBoundaryWithZeroMonths) {
  const Value v = EvalSource("=EDATE(DATE(1900, 1, 1), 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
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

TEST(DateTimeEomonth, RejectsBooleanArgument) {
  // Excel 365 rejects Booleans in EOMONTH with #VALUE! (unlike EDATE, which
  // coerces them to 0/1). Guard both positional arguments.
  const Value v_start = EvalSource("=EOMONTH(TRUE, 0)");
  ASSERT_TRUE(v_start.is_error());
  EXPECT_EQ(v_start.as_error(), ErrorCode::Value);

  const Value v_months = EvalSource("=EOMONTH(44987, TRUE)");
  ASSERT_TRUE(v_months.is_error());
  EXPECT_EQ(v_months.as_error(), ErrorCode::Value);
}

TEST(DateTimeEomonth, UpperEndpointWithZeroMonths) {
  const Value v = EvalSource("=EOMONTH(DATE(9999, 12, 1), 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2958465.0);
}

TEST(DateTimeEomonth, CrossingUpperEndpointIsNum) {
  const Value v = EvalSource("=EOMONTH(DATE(9999, 12, 1), 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(DateTimeEomonth, AcceptsFinalDayTimeFraction) {
  const Value endpoint = EvalSource("=EOMONTH(2958465.999988426, 0)");
  ASSERT_TRUE(endpoint.is_number());
  EXPECT_DOUBLE_EQ(endpoint.as_number(), 2958465.0);
}

TEST(DateTimeEomonth, FinalDayTimeFractionCannotCrossUpperEndpoint) {
  const Value overflow = EvalSource("=EOMONTH(2958465.999988426, 1)");
  ASSERT_TRUE(overflow.is_error());
  EXPECT_EQ(overflow.as_error(), ErrorCode::Num);
}

TEST(DateTimeEomonth, LowerBoundaryPreviousMonthIsZero) {
  const Value v = EvalSource("=EOMONTH(1, -1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(DateTimeEomonth, GiganticMonthOffsetIsNum) {
  const Value v = EvalSource("=EOMONTH(DATE(2024, 1, 1), 1E20)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(DateTimeEomonth, BooleanPrecedenceBeatsGiganticMonthOffset) {
  const Value v = EvalSource("=EOMONTH(TRUE, 1E20)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
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

// ---------------------------------------------------------------------------
// WEEKNUM / ISOWEEKNUM
// ---------------------------------------------------------------------------

TEST(DateTimeWeeknum, DefaultSundayType) {
  // 2024-01-01 is a Monday. Default return_type=1 (Sunday start); week 1
  // contains Jan 1, Jan 1 itself is the second day of that week.
  const Value v = EvalSource("=WEEKNUM(DATE(2024,1,1))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(DateTimeWeeknum, MidYearDefault) {
  // 2024-07-04 is a Thursday of ISO week 27; return_type=1 (Sun start)
  // yields 27 as well.
  const Value v = EvalSource("=WEEKNUM(DATE(2024,7,4))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 27.0);
}

TEST(DateTimeWeeknum, MondayTypeTwo) {
  // 2024-01-01 is a Monday; with return_type=2 (Mon start) it is the
  // first day of week 1.
  const Value v = EvalSource("=WEEKNUM(DATE(2024,1,1),2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(DateTimeWeeknum, Iso21Matches2024Jan1) {
  // ISO 8601: 2024-01-01 is a Monday -> week 1 of 2024.
  const Value v = EvalSource("=WEEKNUM(DATE(2024,1,1),21)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(DateTimeWeeknum, Iso21_2023Jan1_RollsToPrevYear) {
  // 2023-01-01 is a Sunday -> ISO week 52 of 2022.
  const Value v = EvalSource("=WEEKNUM(DATE(2023,1,1),21)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 52.0);
}

TEST(DateTimeWeeknum, InvalidReturnTypeUsesDefault) {
  const Value v = EvalSource("=WEEKNUM(DATE(2024,1,1),99)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(DateTimeWeeknum, NegativeSerialIsNum) {
  const Value v = EvalSource("=WEEKNUM(-1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(DateTimeIsoWeeknum, MidYear) {
  // 2024-07-04 -> ISO week 27.
  const Value v = EvalSource("=ISOWEEKNUM(DATE(2024,7,4))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 27.0);
}

TEST(DateTimeIsoWeeknum, Jan1_2023_IsWeek52OfPrevYear) {
  const Value v = EvalSource("=ISOWEEKNUM(DATE(2023,1,1))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 52.0);
}

TEST(DateTimeIsoWeeknum, Dec31_2024_IsWeek1OfNextYear) {
  // 2024-12-31 is a Tuesday -> ISO week 1 of 2025.
  const Value v = EvalSource("=ISOWEEKNUM(DATE(2024,12,31))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

// ---------------------------------------------------------------------------
// YEARFRAC
// ---------------------------------------------------------------------------

TEST(DateTimeYearfrac, Basis0_US30_360_HalfYear) {
  // 2024-01-01 -> 2024-07-01 under 30/360 = 180/360 = 0.5 exactly.
  const Value v = EvalSource("=YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.5, 1e-12);
}

TEST(DateTimeYearfrac, Basis1_ActualActual_OneFullLeapYear) {
  // 2023-01-01 -> 2024-01-01 exact anniversary spanning 365 days across
  // two calendar years -> 1.0 under the avg-year rule.
  const Value v = EvalSource("=YEARFRAC(DATE(2023,1,1),DATE(2024,1,1),1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0, 1e-12);
}

TEST(DateTimeYearfrac, Basis1_ActualActual_OneYearSpanEnclosingFeb29) {
  // 2020-01-01 -> 2021-01-01 is exactly one year and encloses 2020-02-29, so
  // the denominator is the single leap year length 366 (not the multi-year
  // average 365.5): 366/366 = 1.0 exactly. Regression for the actual/actual
  // leap-year bug (previously 1.0013679890560876).
  const Value v = EvalSource("=YEARFRAC(DATE(2020,1,1),DATE(2021,1,1),1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0, 1e-12);
}

TEST(DateTimeYearfrac, Basis1_ActualActual_OneYearSpanAnniversaryEnclosingFeb29) {
  // 2019-03-01 -> 2020-03-01 is a one-year span whose interval includes
  // 2020-02-29, so the denominator is 366 and the fraction is exactly 1.0.
  const Value v = EvalSource("=YEARFRAC(DATE(2019,3,1),DATE(2020,3,1),1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0, 1e-12);
}

TEST(DateTimeYearfrac, Basis1_ActualActual_NonLeapOneYearSpan) {
  // 2021-01-01 -> 2022-01-01: neither year is leap, span is one year -> 1.0.
  const Value v = EvalSource("=YEARFRAC(DATE(2021,1,1),DATE(2022,1,1),1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0, 1e-12);
}

TEST(DateTimeYearfrac, Basis2_Actual360_HalfLeapYear) {
  // 2024-01-01 -> 2024-07-01 is 182 days in 2024; 182/360.
  const Value v = EvalSource("=YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 182.0 / 360.0, 1e-12);
}

TEST(DateTimeYearfrac, Basis3_Actual365_HalfLeapYear) {
  const Value v = EvalSource("=YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 182.0 / 365.0, 1e-12);
}

TEST(DateTimeYearfrac, Basis4_EU30_360_HalfYear) {
  const Value v = EvalSource("=YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.5, 1e-12);
}

TEST(DateTimeYearfrac, NegativeIntervalSwaps) {
  // Oracle-documented behaviour: YEARFRAC is symmetric in its first two
  // args (positive regardless of order).
  const Value a = EvalSource("=YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),0)");
  const Value b = EvalSource("=YEARFRAC(DATE(2024,7,1),DATE(2024,1,1),0)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_EQ(a.as_number(), b.as_number());
}

TEST(DateTimeYearfrac, InvalidBasisIsNum) {
  const Value v = EvalSource("=YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// DATEDIF
// ---------------------------------------------------------------------------

TEST(DateTimeDatedif, YearsBetween) {
  const Value v = EvalSource("=DATEDIF(DATE(2020,3,15),DATE(2024,7,1),\"Y\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 4.0);
}

TEST(DateTimeDatedif, MonthsBetween) {
  // 2024-01-15 -> 2024-07-01: 5 complete months (d2 < d1 shaves one off).
  const Value v = EvalSource("=DATEDIF(DATE(2024,1,15),DATE(2024,7,1),\"M\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 5.0);
}

TEST(DateTimeDatedif, DaysBetween) {
  const Value v = EvalSource("=DATEDIF(DATE(2024,1,1),DATE(2024,1,15),\"D\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 14.0);
}

TEST(DateTimeDatedif, YmIgnoresYears) {
  // 4 years 3 months 16 days -> YM = 3.
  const Value v = EvalSource("=DATEDIF(DATE(2020,3,15),DATE(2024,7,1),\"YM\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 3.0);
}

TEST(DateTimeDatedif, YdIgnoresYearsSameCalendarPosition) {
  // Same month/day pair: YD = 0 when the calendar positions align.
  const Value v = EvalSource("=DATEDIF(DATE(2020,3,15),DATE(2024,3,15),\"YD\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(DateTimeDatedif, MdDayOnlyDiff) {
  // Same day-of-month -> MD = 0.
  const Value v = EvalSource("=DATEDIF(DATE(2024,1,15),DATE(2024,7,15),\"MD\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(DateTimeDatedif, EndBeforeStartIsNum) {
  const Value v = EvalSource("=DATEDIF(DATE(2024,7,1),DATE(2020,3,15),\"Y\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(DateTimeDatedif, UnknownUnitIsNum) {
  const Value v = EvalSource("=DATEDIF(DATE(2024,1,1),DATE(2024,12,31),\"X\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(DateTimeDatedif, UnitTokensAreCaseInsensitive) {
  struct Case {
    std::string_view unit;
    std::string_view start;
    std::string_view end;
    double expected;
  };
  const Case cases[] = {
      {"y", "DATE(2020,3,15)", "DATE(2024,7,1)", 4.0},   {"m", "DATE(2024,1,15)", "DATE(2024,7,1)", 5.0},
      {"d", "DATE(2024,1,1)", "DATE(2024,1,15)", 14.0},  {"yM", "DATE(2020,3,15)", "DATE(2024,7,1)", 3.0},
      {"Yd", "DATE(2020,3,15)", "DATE(2024,3,15)", 0.0}, {"mD", "DATE(2024,1,15)", "DATE(2024,7,15)", 0.0},
  };
  for (const Case& test : cases) {
    const std::string formula =
        "=DATEDIF(" + std::string(test.start) + "," + std::string(test.end) + ",\"" + std::string(test.unit) + "\")";
    const Value v = EvalSource(formula);
    ASSERT_TRUE(v.is_number()) << formula << ": " << v.debug_to_string();
    EXPECT_DOUBLE_EQ(v.as_number(), test.expected) << formula;
  }
}

// ---------------------------------------------------------------------------
// NETWORKDAYS / WORKDAY
// ---------------------------------------------------------------------------

TEST(DateTimeNetworkdays, OneCompleteWeek) {
  // 2024-01-01 (Mon) .. 2024-01-05 (Fri) inclusive -> 5 business days.
  const Value v = EvalSource("=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,5))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 5.0);
}

TEST(DateTimeNetworkdays, TwoCompleteWeeks) {
  // 2024-01-01 .. 2024-01-12 (Fri) inclusive -> 10 business days.
  const Value v = EvalSource("=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,12))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 10.0);
}

TEST(DateTimeNetworkdays, SingleWeekdayIsOne) {
  // Mon .. Mon same day -> 1.
  const Value v = EvalSource("=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,1))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(DateTimeNetworkdays, SingleSaturdayIsZero) {
  // 2024-01-06 is a Saturday. A weekend-only interval counts as zero.
  const Value v = EvalSource("=NETWORKDAYS(DATE(2024,1,6),DATE(2024,1,6))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(DateTimeNetworkdays, ArrayLiteralHolidays) {
  // Drop Jan 1 and Jan 2 via an inline array literal; 2024-01-01..01-05
  // has 5 business days, minus 2 holidays -> 3.
  const Value v = EvalSource("=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,5),{45292,45293})");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 3.0);
}

TEST(DateTimeNetworkdays, RangeHolidays) {
  // Same setup as above but sourced from A1:A2 on a bound sheet.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(45292.0));  // 2024-01-01
  wb.sheet(0).set_cell_value(1, 0, Value::number(45293.0));  // 2024-01-02
  const Value v = EvalSourceIn("=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,5),A1:A2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 3.0);
}

TEST(DateTimeNetworkdays, ReversedRangeNegates) {
  // Excel 365 returns the negated business-day count when start > end.
  const Value v = EvalSource("=NETWORKDAYS(DATE(2024,1,5),DATE(2024,1,1))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -5.0);
}

TEST(DateTimeNetworkdays, ErrorPropagates) {
  const Value v = EvalSource("=NETWORKDAYS(\"abc\",DATE(2024,1,5))");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(DateTimeNetworkdays, RejectsOutOfRangeSerialBeforeIteration) {
  const Value v = EvalSource("=NETWORKDAYS(1,1000000000)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(DateTimeWorkday, RejectsOutOfRangeDayCountBeforeIteration) {
  const Value v = EvalSource("=WORKDAY(1,1000000000)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(DateTimeWorkday, ForwardFiveDays) {
  // 2024-01-01 (Mon). +5 business days -> 2024-01-08 (the next Monday,
  // because Jan 6/7 are Sat/Sun and skip without counting).
  const Value v = EvalSource("=WORKDAY(DATE(2024,1,1),5)");
  const Value exp = EvalSource("=DATE(2024,1,8)");
  ASSERT_TRUE(v.is_number());
  ASSERT_TRUE(exp.is_number());
  EXPECT_EQ(v.as_number(), exp.as_number());
}

TEST(DateTimeWorkday, BackwardFiveDays) {
  // 2024-01-15 (Mon) - 5 business days -> 2024-01-08 (Mon).
  const Value v = EvalSource("=WORKDAY(DATE(2024,1,15),-5)");
  const Value exp = EvalSource("=DATE(2024,1,8)");
  ASSERT_TRUE(v.is_number());
  ASSERT_TRUE(exp.is_number());
  EXPECT_EQ(v.as_number(), exp.as_number());
}

TEST(DateTimeWorkday, ZeroDaysReturnsStart) {
  // Excel: WORKDAY(start, 0) returns start regardless of weekday status.
  const Value v = EvalSource("=WORKDAY(DATE(2024,1,3),0)");
  const Value exp = EvalSource("=DATE(2024,1,3)");
  ASSERT_TRUE(v.is_number());
  ASSERT_TRUE(exp.is_number());
  EXPECT_EQ(v.as_number(), exp.as_number());
}

TEST(DateTimeWorkday, HolidaysExtendForward) {
  // +5 business days from Mon with Tue / Wed as holidays -> Fri of next
  // week (the 5 working days skipped the two holidays AND the weekend).
  const Value v = EvalSource("=WORKDAY(DATE(2024,1,1),5,{45293,45294})");
  const Value exp = EvalSource("=DATE(2024,1,10)");
  ASSERT_TRUE(v.is_number());
  ASSERT_TRUE(exp.is_number());
  EXPECT_EQ(v.as_number(), exp.as_number());
}

TEST(DateTimeWorkday, RangeHolidays) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(45293.0));  // 2024-01-02
  wb.sheet(0).set_cell_value(1, 0, Value::number(45294.0));  // 2024-01-03
  const Value v = EvalSourceIn("=WORKDAY(DATE(2024,1,1),5,A1:A2)", wb, wb.sheet(0));
  const Value exp = EvalSource("=DATE(2024,1,10)");
  ASSERT_TRUE(v.is_number());
  ASSERT_TRUE(exp.is_number());
  EXPECT_EQ(v.as_number(), exp.as_number());
}

TEST(DateTimeWorkday, ErrorPropagates) {
  const Value v = EvalSource("=WORKDAY(\"abc\",5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// 1904 date system. The 1904 epoch is 1462 days after the 1900 epoch, so a
// given calendar date has a serial 1462 smaller. The calendar family reads
// the epoch from `EvalContext::date1904()`.
// ---------------------------------------------------------------------------

TEST(DateTime1904, DateSerialIs1462Less) {
  // DATE(2020,1,1) is serial 43831 under the 1900 system, 42369 under 1904.
  const Value v1900 = EvalSource("=DATE(2020,1,1)");
  ASSERT_TRUE(v1900.is_number());
  EXPECT_DOUBLE_EQ(v1900.as_number(), 43831.0);
  const Value v1904 = EvalSource1904("=DATE(2020,1,1)");
  ASSERT_TRUE(v1904.is_number());
  EXPECT_DOUBLE_EQ(v1904.as_number(), 42369.0);
}

TEST(DateTime1904, TextRendersDateTextInTheWorkbookEpoch) {
  // TEXT coerces its first argument through the shared numeric ladder, whose
  // date fallback always produces a 1900-system serial. Rendering it with
  // 1904 format codes must therefore round-trip to the same calendar day
  // under both epochs, not shift by the 1462-day gap.
  const Value v1900 = EvalSource("=TEXT(\"2024-03-15\", \"yyyy/m/d\")");
  ASSERT_TRUE(v1900.is_text());
  EXPECT_EQ(v1900.as_text(), "2024/3/15");

  const Value v1904 = EvalSource1904("=TEXT(\"2024-03-15\", \"yyyy/m/d\")");
  ASSERT_TRUE(v1904.is_text());
  EXPECT_EQ(v1904.as_text(), "2024/3/15");
}

TEST(DateTime1904, TextShiftsOnlyDateDerivedSerials) {
  // The epoch shift is keyed on the coercion ladder's date rung, not on the
  // format codes. A value that already is a serial — whether it arrives as a
  // number or as numeric text — is read in the workbook's own epoch and must
  // not be moved again.
  const std::string number_1900 = std::string(EvalSource("=TEXT(45366,\"yyyy/m/d\")").as_text());
  EXPECT_EQ(number_1900, "2024/3/15");

  const std::string number_1904 = std::string(EvalSource1904("=TEXT(45366,\"yyyy/m/d\")").as_text());
  EXPECT_EQ(number_1904, "2028/3/16") << "a bare serial follows the workbook epoch";

  const std::string numeric_text_1904 = std::string(EvalSource1904("=TEXT(\"45366\",\"yyyy/m/d\")").as_text());
  EXPECT_EQ(numeric_text_1904, number_1904) << "numeric text must not pick up the date rung's epoch shift";

  // A non-date rung of the ladder renders identically under both epochs.
  const std::string percent_1900 = std::string(EvalSource("=TEXT(\"50%\",\"0.00\")").as_text());
  const std::string percent_1904 = std::string(EvalSource1904("=TEXT(\"50%\",\"0.00\")").as_text());
  EXPECT_EQ(percent_1900, "0.50");
  EXPECT_EQ(percent_1904, percent_1900);

  // Time-only text is a day fraction: the same number, and the same
  // rendering, under either epoch.
  const std::string time_1900 = std::string(EvalSource("=TEXT(\"13:30\",\"h:mm\")").as_text());
  const std::string time_1904 = std::string(EvalSource1904("=TEXT(\"13:30\",\"h:mm\")").as_text());
  EXPECT_EQ(time_1900, "13:30");
  EXPECT_EQ(time_1904, time_1900);
}

TEST(DateTime1904, SerialZeroExtractorsFollowTheWorkbookEpoch) {
  // Serial 0 is the fictitious "1900-01-00" origin under the 1900 system and
  // a real 1904-01-01 under the 1904 one, so the day component differs while
  // the month does not. Two branches of `coerce_serial_ymd` meet here: the
  // explicit day-0 alias, which only applies under the 1900 epoch, and the
  // epoch shift that rebases a 1904 serial before decoding. Expected values
  // come from the `datetime` / `datetime_1904_text` oracle suites, not from
  // this implementation.
  const Value year_1900 = EvalSource("=YEAR(0)");
  const Value month_1900 = EvalSource("=MONTH(0)");
  const Value day_1900 = EvalSource("=DAY(0)");
  ASSERT_TRUE(year_1900.is_number());
  ASSERT_TRUE(month_1900.is_number());
  ASSERT_TRUE(day_1900.is_number());
  EXPECT_DOUBLE_EQ(year_1900.as_number(), 1900.0);
  EXPECT_DOUBLE_EQ(month_1900.as_number(), 1.0);
  EXPECT_DOUBLE_EQ(day_1900.as_number(), 0.0);

  const Value year_1904 = EvalSource1904("=YEAR(0)");
  const Value month_1904 = EvalSource1904("=MONTH(0)");
  const Value day_1904 = EvalSource1904("=DAY(0)");
  ASSERT_TRUE(year_1904.is_number());
  ASSERT_TRUE(month_1904.is_number());
  ASSERT_TRUE(day_1904.is_number());
  EXPECT_DOUBLE_EQ(year_1904.as_number(), 1904.0);
  EXPECT_DOUBLE_EQ(month_1904.as_number(), 1.0);
  EXPECT_DOUBLE_EQ(day_1904.as_number(), 1.0);
}

TEST(DateTime1904, UpperEndpointIsEpochSpecific) {
  const Value endpoint = EvalSource1904("=DATE(9999, 12, 31)");
  ASSERT_TRUE(endpoint.is_number());
  EXPECT_DOUBLE_EQ(endpoint.as_number(), 2957003.0);

  const Value normalised = EvalSource1904("=DATE(9999, 11, 61)");
  ASSERT_TRUE(normalised.is_number());
  EXPECT_DOUBLE_EQ(normalised.as_number(), 2957003.0);
}

TEST(DateTime1904, DateOverflowPastUpperEndpointIsNum) {
  const Value v = EvalSource1904("=DATE(9999, 12, 32)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(DateTime1904, EdateUpperEndpointAndOverflow) {
  const Value endpoint = EvalSource1904("=EDATE(DATE(9999, 12, 31), 0)");
  ASSERT_TRUE(endpoint.is_number());
  EXPECT_DOUBLE_EQ(endpoint.as_number(), 2957003.0);

  const Value overflow = EvalSource1904("=EDATE(DATE(9999, 12, 1), 1)");
  ASSERT_TRUE(overflow.is_error());
  EXPECT_EQ(overflow.as_error(), ErrorCode::Num);
}

TEST(DateTime1904, EdateInputOnePastUpperEndpointIsNum) {
  const Value v = EvalSource1904("=EDATE(DATE(9999, 12, 31) + 1, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(DateTime1904, ExtractorRejectsOnePastEpochSpecificUpperEndpoint) {
  const Value v = EvalSource1904("=YEAR(DATE(9999, 12, 31) + 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(DateTime1904, AcceptsFinalDayTimeFraction) {
  const Value year = EvalSource1904("=YEAR(2957003.999988426)");
  const Value month = EvalSource1904("=MONTH(2957003.999988426)");
  const Value day = EvalSource1904("=DAY(2957003.999988426)");
  ASSERT_TRUE(year.is_number());
  ASSERT_TRUE(month.is_number());
  ASSERT_TRUE(day.is_number());
  EXPECT_DOUBLE_EQ(year.as_number(), 9999.0);
  EXPECT_DOUBLE_EQ(month.as_number(), 12.0);
  EXPECT_DOUBLE_EQ(day.as_number(), 31.0);
}

TEST(DateTime1904, RejectsExactNextDay) {
  const Value year = EvalSource1904("=YEAR(2957004)");
  const Value month = EvalSource1904("=MONTH(2957004)");
  const Value day = EvalSource1904("=DAY(2957004)");
  ASSERT_TRUE(year.is_error());
  ASSERT_TRUE(month.is_error());
  ASSERT_TRUE(day.is_error());
  EXPECT_EQ(year.as_error(), ErrorCode::Num);
  EXPECT_EQ(month.as_error(), ErrorCode::Num);
  EXPECT_EQ(day.as_error(), ErrorCode::Num);
}

TEST(DateTime1904, EdateGiganticMonthOffsetIsNum) {
  const Value v = EvalSource1904("=EDATE(DATE(2024, 1, 1), 1E20)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(DateTime1904, EdateLeftmostErrorPrecedesGiganticMonthOffset) {
  const Value v = EvalSource1904("=EDATE(1/0, 1E20)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(DateTime1904, EdateLowerBoundaryWithZeroMonths) {
  const Value v = EvalSource1904("=EDATE(DATE(1904, 1, 1), 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(DateTime1904, EomonthUpperEndpointAndOverflow) {
  const Value endpoint = EvalSource1904("=EOMONTH(DATE(9999, 12, 1), 0)");
  ASSERT_TRUE(endpoint.is_number());
  EXPECT_DOUBLE_EQ(endpoint.as_number(), 2957003.0);

  const Value overflow = EvalSource1904("=EOMONTH(DATE(9999, 12, 1), 1)");
  ASSERT_TRUE(overflow.is_error());
  EXPECT_EQ(overflow.as_error(), ErrorCode::Num);
}

TEST(DateTime1904, EdateAcceptsFinalDayTimeFraction) {
  const Value endpoint = EvalSource1904("=EDATE(2957003.999988426, 0)");
  ASSERT_TRUE(endpoint.is_number());
  EXPECT_DOUBLE_EQ(endpoint.as_number(), 2957003.0);
}

TEST(DateTime1904, EdateFinalDayTimeFractionCannotCrossUpperEndpoint) {
  const Value overflow = EvalSource1904("=EDATE(2957003.999988426, 1)");
  ASSERT_TRUE(overflow.is_error());
  EXPECT_EQ(overflow.as_error(), ErrorCode::Num);
}

TEST(DateTime1904, EomonthAcceptsFinalDayTimeFraction) {
  const Value endpoint = EvalSource1904("=EOMONTH(2957003.999988426, 0)");
  ASSERT_TRUE(endpoint.is_number());
  EXPECT_DOUBLE_EQ(endpoint.as_number(), 2957003.0);
}

TEST(DateTime1904, EomonthFinalDayTimeFractionCannotCrossUpperEndpoint) {
  const Value overflow = EvalSource1904("=EOMONTH(2957003.999988426, 1)");
  ASSERT_TRUE(overflow.is_error());
  EXPECT_EQ(overflow.as_error(), ErrorCode::Num);
}

TEST(DateTime1904, EomonthGiganticMonthOffsetIsNum) {
  const Value v = EvalSource1904("=EOMONTH(DATE(2024, 1, 1), 1E20)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(DateTime1904, EomonthBooleanPrecedenceBeatsGiganticMonthOffset) {
  const Value v = EvalSource1904("=EOMONTH(TRUE, 1E20)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(DateTime1904, EomonthLowerBoundaryControl) {
  const Value v = EvalSource1904("=EOMONTH(DATE(1904, 1, 1), 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 30.0);
}

TEST(DateTime1904, YearMonthDayInvertUnder1904) {
  // Serial 42369 is 2020-01-01 in the 1904 system.
  const Value y = EvalSource1904("=YEAR(42369)");
  ASSERT_TRUE(y.is_number());
  EXPECT_DOUBLE_EQ(y.as_number(), 2020.0);
  const Value m = EvalSource1904("=MONTH(42369)");
  ASSERT_TRUE(m.is_number());
  EXPECT_DOUBLE_EQ(m.as_number(), 1.0);
  const Value d = EvalSource1904("=DAY(42369)");
  ASSERT_TRUE(d.is_number());
  EXPECT_DOUBLE_EQ(d.as_number(), 1.0);
}

TEST(DateTime1904, EdateShiftsWithin1904Epoch) {
  // EDATE(2020-01-01, 1) -> 2020-02-01; serial 42400 in the 1904 system
  // (43862 - 1462).
  const Value v = EvalSource1904("=EDATE(42369,1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 42400.0);
}

TEST(DateTime1904, DateValueUsesWorkbookEpoch) {
  const Value v = EvalSource1904("=DATEVALUE(\"2020-01-01\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 42369.0);
}

TEST(DateTime1904, WeekdayMatchesCalendarDateUnder1904) {
  // 2020-01-01 is a Wednesday -> WEEKDAY type 1 (Sun=1) returns 4.
  const Value v = EvalSource1904("=WEEKDAY(42369)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 4.0);
}

TEST(DateTime1904, DefaultSystemUnaffected) {
  // Sanity: the 1900 default path is unchanged (YEAR of the 1900 serial).
  const Value v = EvalSource("=YEAR(43831)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2020.0);
}

TEST(DateTime1904, TodayIsPositiveIntegerUnder1904) {
  // TODAY() reads the wall clock; assert only that the 1904 epoch yields a
  // positive integer serial (deterministic value would depend on the date).
  const Value v = EvalSource1904("=TODAY()");
  ASSERT_TRUE(v.is_number());
  EXPECT_GT(v.as_number(), 0.0);
  EXPECT_DOUBLE_EQ(v.as_number(), std::floor(v.as_number()));
}

TEST(DateTime1904, TextDateRenderingUsesWorkbookEpoch) {
  // Serial 42369 is 2015-12-31 in the 1900 system and 2020-01-01 in the 1904
  // system. TEXT's date format must render against the workbook epoch.
  const Value v1900 = EvalSource("=TEXT(42369,\"yyyy-mm-dd\")");
  ASSERT_TRUE(v1900.is_text());
  EXPECT_EQ(v1900.as_text(), "2015-12-31");
  const Value v1904 = EvalSource1904("=TEXT(42369,\"yyyy-mm-dd\")");
  ASSERT_TRUE(v1904.is_text());
  EXPECT_EQ(v1904.as_text(), "2020-01-01");
}

TEST(DateTime1904, TextRoundTripsDateBuiltinUnder1904) {
  // DATE and TEXT share the workbook epoch, so DATE(2020,3,15) formatted back
  // reads "2020-03-15" under both systems.
  const Value v = EvalSource1904("=TEXT(DATE(2020,3,15),\"yyyy-mm-dd\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "2020-03-15");
}

TEST(DateTime1904, TextNumericFormatUnaffectedByEpoch) {
  // A pure-numeric format has no date tokens, so the epoch is irrelevant.
  const Value v = EvalSource1904("=TEXT(1234.5,\"0.00\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "1234.50");
}

TEST(DateTimeSerialZero, 1900SystemKeepsFictitiousDayAlias) {
  const Value year = EvalSource("=YEAR(0)");
  const Value month = EvalSource("=MONTH(0)");
  const Value day = EvalSource("=DAY(0)");
  ASSERT_TRUE(year.is_number());
  ASSERT_TRUE(month.is_number());
  ASSERT_TRUE(day.is_number());
  EXPECT_DOUBLE_EQ(year.as_number(), 1900.0);
  EXPECT_DOUBLE_EQ(month.as_number(), 1.0);
  EXPECT_DOUBLE_EQ(day.as_number(), 0.0);
}

TEST(DateTimeSerialZero, 1904SystemUsesRealEpochDay) {
  const Value year = EvalSource1904("=YEAR(0)");
  const Value month = EvalSource1904("=MONTH(0)");
  const Value day = EvalSource1904("=DAY(0)");
  ASSERT_TRUE(year.is_number());
  ASSERT_TRUE(month.is_number());
  ASSERT_TRUE(day.is_number());
  EXPECT_DOUBLE_EQ(year.as_number(), 1904.0);
  EXPECT_DOUBLE_EQ(month.as_number(), 1.0);
  EXPECT_DOUBLE_EQ(day.as_number(), 1.0);
}

TEST(DateTime1904, RawTextUsesCalendarDateAcrossExtractors) {
  const Value year = EvalSource1904("=YEAR(\"2024-03-15\")");
  const Value month = EvalSource1904("=MONTH(\"2024-03-15\")");
  const Value day = EvalSource1904("=DAY(\"2024-03-15\")");
  const Value weekday = EvalSource1904("=WEEKDAY(\"2024-03-15\")");
  const Value weeknum = EvalSource1904("=WEEKNUM(\"2024-03-15\")");
  const Value isoweeknum = EvalSource1904("=ISOWEEKNUM(\"2024-03-15\")");
  ASSERT_TRUE(year.is_number());
  ASSERT_TRUE(month.is_number());
  ASSERT_TRUE(day.is_number());
  ASSERT_TRUE(weekday.is_number());
  ASSERT_TRUE(weeknum.is_number());
  ASSERT_TRUE(isoweeknum.is_number());
  EXPECT_DOUBLE_EQ(year.as_number(), 2024.0);
  EXPECT_DOUBLE_EQ(month.as_number(), 3.0);
  EXPECT_DOUBLE_EQ(day.as_number(), 15.0);
  EXPECT_DOUBLE_EQ(weekday.as_number(), 6.0);  // Friday, Sunday-first.
  EXPECT_DOUBLE_EQ(weeknum.as_number(), 11.0);
  EXPECT_DOUBLE_EQ(isoweeknum.as_number(), 11.0);
}

TEST(DateTime1904, RawTextDayCountsMatchDatevalueNumbers) {
  const Value datedif_raw = EvalSource1904("=DATEDIF(\"2024-01-01\",DATE(2024,12,31),\"D\")");
  const Value datedif_datevalue = EvalSource1904("=DATEDIF(DATEVALUE(\"2024-01-01\"),DATE(2024,12,31),\"D\")");
  const Value days360_raw = EvalSource1904("=DAYS360(\"2024-01-01\",DATE(2024,12,31))");
  const Value days360_datevalue = EvalSource1904("=DAYS360(DATEVALUE(\"2024-01-01\"),DATE(2024,12,31))");
  const Value days_raw = EvalSource1904("=DAYS(\"2024-03-15\",DATE(2024,3,14))");
  const Value days_datevalue = EvalSource1904("=DAYS(DATEVALUE(\"2024-03-15\"),DATE(2024,3,14))");
  ASSERT_TRUE(datedif_raw.is_number());
  ASSERT_TRUE(datedif_datevalue.is_number());
  ASSERT_TRUE(days360_raw.is_number());
  ASSERT_TRUE(days360_datevalue.is_number());
  ASSERT_TRUE(days_raw.is_number());
  ASSERT_TRUE(days_datevalue.is_number());
  EXPECT_DOUBLE_EQ(datedif_raw.as_number(), 365.0);
  EXPECT_DOUBLE_EQ(datedif_datevalue.as_number(), 365.0);
  EXPECT_DOUBLE_EQ(days360_raw.as_number(), 360.0);
  EXPECT_DOUBLE_EQ(days360_datevalue.as_number(), 360.0);
  EXPECT_DOUBLE_EQ(days_raw.as_number(), 1.0);
  EXPECT_DOUBLE_EQ(days_datevalue.as_number(), 1.0);
}

TEST(DateTime1904, DaysBroadcastsArrayArguments) {
  const Value result = EvalSource1904("=DAYS({\"2024-03-15\",\"2024-03-16\"},DATE(2024,3,14))");
  ASSERT_TRUE(result.is_array());
  ASSERT_EQ(result.as_array_rows(), 1U);
  ASSERT_EQ(result.as_array_cols(), 2U);
  ASSERT_TRUE(result.as_array()->cells[0].is_number());
  ASSERT_TRUE(result.as_array()->cells[1].is_number());
  EXPECT_DOUBLE_EQ(result.as_array()->cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(result.as_array()->cells[1].as_number(), 2.0);
}

TEST(DateTime1904, LegacyMonthFunctionsUseRawTextInputCoordinateAfterValidation) {
  const Value edate_raw = EvalSource1904("=EDATE(\"2024-03-15\",1)");
  const Value edate_datevalue = EvalSource1904("=EDATE(DATEVALUE(\"2024-03-15\"),1)");
  const Value eomonth_raw = EvalSource1904("=EOMONTH(\"2024-03-15\",0)");
  const Value eomonth_datevalue = EvalSource1904("=EOMONTH(DATEVALUE(\"2024-03-15\"),0)");
  ASSERT_TRUE(edate_raw.is_number());
  ASSERT_TRUE(edate_datevalue.is_number());
  ASSERT_TRUE(eomonth_raw.is_number());
  ASSERT_TRUE(eomonth_datevalue.is_number());
  EXPECT_DOUBLE_EQ(edate_raw.as_number(), 42473.0);
  EXPECT_DOUBLE_EQ(edate_datevalue.as_number(), 43935.0);
  EXPECT_DOUBLE_EQ(eomonth_raw.as_number(), 42459.0);
  EXPECT_DOUBLE_EQ(eomonth_datevalue.as_number(), 43920.0);
}

TEST(DateTime1904, YearfracRawTextUsesLegacyEndpoints) {
  const Value raw_start = EvalSource1904("=YEARFRAC(\"2024-01-01\",DATE(2024,12,31),3)");
  const Value raw_end = EvalSource1904("=YEARFRAC(DATE(2024,1,1),\"2024-12-31\",3)");
  const Value datevalue_start = EvalSource1904("=YEARFRAC(DATEVALUE(\"2024-01-01\"),DATE(2024,12,31),3)");
  const Value datevalue_end = EvalSource1904("=YEARFRAC(DATE(2024,1,1),DATEVALUE(\"2024-12-31\"),3)");
  const Value both_raw_basis3 = EvalSource1904("=YEARFRAC(\"2024-01-01\",\"2024-12-31\",3)");
  const Value both_raw_basis1 = EvalSource1904("=YEARFRAC(\"2024-01-01\",\"2024-12-31\",1)");
  ASSERT_TRUE(raw_start.is_number());
  ASSERT_TRUE(raw_end.is_number());
  ASSERT_TRUE(datevalue_start.is_number());
  ASSERT_TRUE(datevalue_end.is_number());
  ASSERT_TRUE(both_raw_basis3.is_number());
  ASSERT_TRUE(both_raw_basis1.is_number());
  EXPECT_DOUBLE_EQ(raw_start.as_number(), 1827.0 / 365.0);
  EXPECT_DOUBLE_EQ(raw_end.as_number(), 1097.0 / 365.0);
  EXPECT_DOUBLE_EQ(datevalue_start.as_number(), 1.0);
  EXPECT_DOUBLE_EQ(datevalue_end.as_number(), 1.0);
  EXPECT_DOUBLE_EQ(both_raw_basis3.as_number(), 1.0);
  EXPECT_DOUBLE_EQ(both_raw_basis1.as_number(), 365.0 / 366.0);
}

TEST(DateTime1904, LegacyEdateTextBoundaryProjections) {
  struct Case {
    const char* date;
    int months;
    double expected;
  };
  constexpr Case cases[] = {
      {"1904-01-01", 0, -1462.0}, {"1907-12-31", 0, -2.0},   {"1908-01-01", 0, -1.0},
      {"1908-01-02", 0, 0.0},     {"1908-01-01", -1, -32.0}, {"2024-02-29", 0, 42427.0},
  };
  for (const Case& tc : cases) {
    const std::string formula = "=EDATE(\"" + std::string(tc.date) + "\"," + std::to_string(tc.months) + ")";
    const Value v = EvalSource1904(formula);
    ASSERT_TRUE(v.is_number()) << formula;
    EXPECT_DOUBLE_EQ(v.as_number(), tc.expected) << formula;
  }
}

TEST(DateTime1904, LegacyEomonthTextBoundaryProjections) {
  struct Case {
    const char* date;
    double expected;
  };
  constexpr Case cases[] = {
      {"1904-01-01", -1431.0}, {"1907-12-31", -1.0},    {"1908-01-01", -1.0},
      {"1908-01-02", 30.0},    {"2024-02-29", 42428.0},
  };
  for (const Case& tc : cases) {
    const std::string formula = "=EOMONTH(\"" + std::string(tc.date) + "\",0)";
    const Value v = EvalSource1904(formula);
    ASSERT_TRUE(v.is_number()) << formula;
    EXPECT_DOUBLE_EQ(v.as_number(), tc.expected) << formula;
  }
}

TEST(DateTime1904, PreEpochRawDateTextIsValueError) {
  const Value year = EvalSource1904("=YEAR(\"1903-12-31\")");
  const Value edate = EvalSource1904("=EDATE(\"1903-12-31\",0)");
  const Value eomonth = EvalSource1904("=EOMONTH(\"1903-12-31\",0)");
  ASSERT_TRUE(year.is_error());
  ASSERT_TRUE(edate.is_error());
  ASSERT_TRUE(eomonth.is_error());
  EXPECT_EQ(year.as_error(), ErrorCode::Value);
  EXPECT_EQ(edate.as_error(), ErrorCode::Value);
  EXPECT_EQ(eomonth.as_error(), ErrorCode::Value);
}

TEST(DateTime1904, NumericTextWhitespaceAndTimeOnlyStayNonDirect) {
  const Value numeric_text = EvalSource1904("=YEAR(\"42369\")");
  const Value whitespace = EvalSource1904("=YEAR(\" 2024-03-15 \")");
  const Value time_value = EvalSource1904("=YEAR(\"12:00\")");
  ASSERT_TRUE(numeric_text.is_number());
  ASSERT_TRUE(whitespace.is_error());
  ASSERT_TRUE(time_value.is_number());
  EXPECT_DOUBLE_EQ(numeric_text.as_number(), 2020.0);
  EXPECT_EQ(whitespace.as_error(), ErrorCode::Value);
  EXPECT_DOUBLE_EQ(time_value.as_number(), 1904.0);
}

TEST(DateTime1904, NumericAndDatevalueMonthBoundsStayCanonical) {
  const Value numeric_lower = EvalSource1904("=EDATE(0,0)");
  const Value datevalue_lower = EvalSource1904("=EDATE(DATEVALUE(\"1904-01-01\"),0)");
  const Value numeric_upper = EvalSource1904("=EDATE(DATE(9999,12,31),0)");
  const Value datevalue_upper = EvalSource1904("=EDATE(DATEVALUE(\"9999-12-31\"),0)");
  ASSERT_TRUE(numeric_lower.is_number());
  ASSERT_TRUE(datevalue_lower.is_number());
  ASSERT_TRUE(numeric_upper.is_number());
  ASSERT_TRUE(datevalue_upper.is_number());
  EXPECT_DOUBLE_EQ(numeric_lower.as_number(), 0.0);
  EXPECT_DOUBLE_EQ(datevalue_lower.as_number(), 0.0);
  EXPECT_DOUBLE_EQ(numeric_upper.as_number(), 2957003.0);
  EXPECT_DOUBLE_EQ(datevalue_upper.as_number(), 2957003.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
