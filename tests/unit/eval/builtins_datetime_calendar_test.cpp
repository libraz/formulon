// End-to-end date/time built-in tests: week numbering, year fractions, date differences, and workdays.

#include <string>

#include "builtins_datetime_test_helpers.h"
namespace formulon {
namespace eval {
namespace {

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

// Windows Excel 365 (the primary reference) rejects an unsupported
// return_type with #NUM!; only 1, 2, 11..17 and 21 are accepted.
TEST(DateTimeWeeknum, InvalidReturnTypeIsNum) {
  for (const char* src :
       {"=WEEKNUM(DATE(2024,1,1),99)", "=WEEKNUM(DATE(2024,1,1),0)", "=WEEKNUM(DATE(2024,1,1),3)",
        "=WEEKNUM(DATE(2024,1,1),10)", "=WEEKNUM(DATE(2024,1,1),18)", "=WEEKNUM(DATE(2024,1,1),22)"}) {
    const Value v = EvalSource(src);
    ASSERT_TRUE(v.is_error()) << src;
    EXPECT_EQ(v.as_error(), ErrorCode::Num) << src;
  }
}

TEST(DateTimeWeeknum, EveryAcceptedReturnTypeComputes) {
  for (const char* src :
       {"=WEEKNUM(DATE(2024,1,1),1)", "=WEEKNUM(DATE(2024,1,1),2)", "=WEEKNUM(DATE(2024,1,1),11)",
        "=WEEKNUM(DATE(2024,1,1),17)", "=WEEKNUM(DATE(2024,1,1),21)", "=WEEKNUM(DATE(2024,1,1),11.9)"}) {
    EXPECT_TRUE(EvalSource(src).is_number()) << src;
  }
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

TEST(DateTimeYearfrac, Basis1_ActualActual_MultiYearAveragesFullYearLengths) {
  // Spans longer than one year divide by the average full length of the
  // years y1..y2, even when the leap year's Feb 29 lies outside the span.
  // Values captured from Excel 365.
  const Value leap_start = EvalSource("=YEARFRAC(DATE(2024,3,1),DATE(2026,6,1),1)");
  ASSERT_TRUE(leap_start.is_number());
  EXPECT_EQ(leap_start.as_number(), 2.25);
  const Value leap_end = EvalSource("=YEARFRAC(DATE(2022,6,1),DATE(2024,2,1),1)");
  ASSERT_TRUE(leap_end.is_number());
  EXPECT_EQ(leap_end.as_number(), 1.6697080291970803);
  const Value contains = EvalSource("=YEARFRAC(DATE(2023,6,1),DATE(2025,6,1),1)");
  ASSERT_TRUE(contains.is_number());
  EXPECT_EQ(contains.as_number(), 2.0009124087591244);
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

}  // namespace
}  // namespace eval
}  // namespace formulon
