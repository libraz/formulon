// End-to-end date/time built-in tests: constructors and extractors.

#include <limits>
#include <utility>

#include "builtins_datetime_test_helpers.h"
namespace formulon {
namespace eval {
namespace {

// ---------------------------------------------------------------------------
// DATE
// ---------------------------------------------------------------------------

TEST(DateTimeTime, OverflowingComponentsReturnNumError) {
  for (const char* formula : {"=TIME(1E308,0,0)", "=TIME(0,1E308,0)", "=TIME(1E308,-1E308,0)"}) {
    SCOPED_TRACE(formula);
    const Value value = EvalSource(formula);
    ASSERT_TRUE(value.is_error());
    EXPECT_EQ(value.as_error(), ErrorCode::Num);
  }
  const Value carry = EvalSource("=TIME(25,0,0)");
  ASSERT_TRUE(carry.is_number());
  EXPECT_DOUBLE_EQ(carry.as_number(), 1.0 / 24.0);
}

TEST(DateTimeTime, ComponentUpperBoundsApplyBeforeDayWrapping) {
  for (const char* formula : {"=TIME(32768,0,0)", "=TIME(0,32768,0)", "=TIME(0,0,32768)", "=TIME(0,0,1E308)"}) {
    SCOPED_TRACE(formula);
    const Value value = EvalSource(formula);
    ASSERT_TRUE(value.is_error());
    EXPECT_EQ(value.as_error(), ErrorCode::Num);
  }
  for (const auto& row :
       {std::pair{"=TIME(32767,0,0)", 7.0 / 24.0},
        std::pair{"=TIME(0,32767,0)", (32767.0 * 60.0 - 22.0 * 86400.0) / 86400.0},
        std::pair{"=TIME(0,0,32767)", 32767.0 / 86400.0}, std::pair{"=TIME(32767.9,0,0)", 7.0 / 24.0}}) {
    SCOPED_TRACE(row.first);
    const Value value = EvalSource(row.first);
    ASSERT_TRUE(value.is_number());
    EXPECT_DOUBLE_EQ(value.as_number(), row.second);
  }
}

TEST(DateTimeNumericSafety, NonFiniteCellTimesReturnNumError) {
  Workbook workbook = Workbook::create();
  for (double input : {std::numeric_limits<double>::quiet_NaN(), std::numeric_limits<double>::infinity(),
                       -std::numeric_limits<double>::infinity()}) {
    workbook.sheet(0).set_cell_value(0, 0, Value::number(input));
    for (const char* formula : {"=HOUR(A1)", "=MINUTE(A1)", "=SECOND(A1)"}) {
      const Value value = EvalSourceIn(formula, workbook, workbook.sheet(0));
      ASSERT_TRUE(value.is_error()) << formula;
      EXPECT_EQ(value.as_error(), ErrorCode::Num);
    }
  }
}

TEST(DateTimeNumericSafety, HugeSelectorsReturnNumError) {
  for (const char* formula : {"=WEEKDAY(1,1E300)", "=WEEKDAY(1,-1E300)", "=WEEKNUM(1,1E300)", "=WEEKNUM(1,-1E300)",
                              "=YEARFRAC(1,2,1E300)", "=YEARFRAC(1,2,-1E300)"}) {
    SCOPED_TRACE(formula);
    const Value value = EvalSource(formula);
    ASSERT_TRUE(value.is_error());
    EXPECT_EQ(value.as_error(), ErrorCode::Num);
  }
}

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
  // TIME(25, 0, 0) == TIME(1, 0, 0): the mod runs on raw total seconds
  // before the division by 86400, so equal totals are bit-identical.
  const Value a = EvalSource("=TIME(25, 0, 0)");
  const Value b = EvalSource("=TIME(1, 0, 0)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_EQ(a.as_number(), b.as_number());
}

TEST(DateTimeTime, MinuteOverflowNormalises) {
  // TIME(1, 60, 0) == TIME(2, 0, 0), bit-identical (same reasoning as above).
  const Value a = EvalSource("=TIME(1, 60, 0)");
  const Value b = EvalSource("=TIME(2, 0, 0)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_EQ(a.as_number(), b.as_number());
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

TEST(DateTimeWeekday, PreGhostDaysUseExcelSerialWeekdays) {
  // Excel counts weekdays off the serial, so days before the fictitious
  // 1900-02-29 fall one day earlier than their Gregorian weekday.
  const Value jan_1 = EvalSource("=WEEKDAY(DATE(1900, 1, 1), 2)");
  const Value feb_28 = EvalSource("=WEEKDAY(DATE(1900, 2, 28), 2)");
  ASSERT_TRUE(jan_1.is_number());
  ASSERT_TRUE(feb_28.is_number());
  EXPECT_EQ(jan_1.as_number(), 7.0);   // Sunday
  EXPECT_EQ(feb_28.as_number(), 2.0);  // Tuesday
}

TEST(DateTimeWeekday, BlankCellIsSerialZeroSaturday) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=WEEKDAY(A1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 7.0);
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

}  // namespace
}  // namespace eval
}  // namespace formulon
