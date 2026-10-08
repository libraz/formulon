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

// ---------------------------------------------------------------------------
// Excel-measured near-integer snap, second rounding, DATE limits, serial 0 and
// the Analysis-ToolPak argument rule for the calendar family.
// ---------------------------------------------------------------------------

void ExpectNumber(std::string_view src, double expected) {
  const Value v = EvalSource(src);
  ASSERT_TRUE(v.is_number()) << src;
  EXPECT_DOUBLE_EQ(v.as_number(), expected) << src;
}

void ExpectError(std::string_view src, ErrorCode code) {
  const Value v = EvalSource(src);
  ASSERT_TRUE(v.is_error()) << src;
  EXPECT_EQ(v.as_error(), code) << src;
}

TEST(DateTimeSnap, DateParts) {
  ExpectNumber("=DATE(2025,1,14.9999997616)", 45672.0);
  ExpectNumber("=DATE(2025,1,14.999999761)", 45671.0);
  ExpectNumber("=DATE(2025,1,10000.9999997616)", 55658.0);
  ExpectNumber("=DATE(2025,1,0.9999999)", 45658.0);
  ExpectNumber("=DATE(2025,3,-1.9999999)", 45714.0);
  ExpectNumber("=DATE(2025,3,-1.0000001)", 45715.0);
  ExpectNumber("=DATE(1899.9999999,1,1)", 1.0);
  ExpectNumber("=DATE(2025,12.9999999,1)", 46023.0);
}

TEST(DateTimeSnap, TimeParts) {
  ExpectNumber("=TIME(0,0,0.99999995)*86400", 1.0);
  ExpectNumber("=TIME(1,-0.9999999,0)*1440", 59.0);
  ExpectNumber("=TIME(0,0,0.5)*86400", 0.0);
}

TEST(DateTimeSnap, ReturnTypes) {
  ExpectNumber("=WEEKDAY(1,1.9999999)", 7.0);
  ExpectNumber("=WEEKDAY(1,1.999999)", 1.0);
  ExpectNumber("=WEEKNUM(10,1.9999999)", 3.0);
}

TEST(DateTimeSnap, DoesNotSnapElsewhere) {
  ExpectNumber("=EOMONTH(DATE(2025,1,15),0.9999999)", 45688.0);
  ExpectNumber("=EDATE(DATE(2025,1,15),1.9999999)", 45703.0);
}

TEST(DateTimeSecondRounding, DayMonthYearWeekday) {
  ExpectNumber("=DAY(40000.9999942)", 6.0);
  ExpectNumber("=DAY(40000.9999943)", 7.0);
  ExpectNumber("=MONTH(45688.9999999)", 2.0);
  ExpectNumber("=YEAR(DATE(2025,12,31)+0.9999999)", 2026.0);
  ExpectNumber("=WEEKDAY(40000.9999999)", 3.0);
  ExpectNumber("=WEEKDAY(40000.9999942)", 2.0);
  ExpectError("=DAY(2958465.999999)", ErrorCode::Num);
  ExpectNumber("=DAY(2958465.99998)", 31.0);
}

TEST(DateTimeDateLimits, MonthBounds) {
  ExpectError("=DATE(2025,32767,1)", ErrorCode::Num);
  ExpectError("=DATE(5000,-32768,1)", ErrorCode::Num);
  ExpectError("=DATE(2025,32766.9999999,1)", ErrorCode::Num);
  ExpectError("=DATE(5000,-32767.9999999,1)", ErrorCode::Num);
  EXPECT_TRUE(EvalSource("=DATE(1900,32766,1)").is_number());
}

TEST(DateTimeDateLimits, DaySaturates) {
  ExpectNumber("=DATE(2025,1,32767)", 78424.0);
  ExpectNumber("=DATE(2025,1,32768)", 78424.0);
  ExpectNumber("=DATE(2025,1,1000000)", 78424.0);
  ExpectNumber("=DATE(2025,1,32767.5)", 78424.0);
  ExpectNumber("=DATE(2025,1,-32768)", 12889.0);
  ExpectNumber("=DATE(2025,1,-32769)", 78424.0);
  ExpectNumber("=DATE(2025,1,-1000000)", 78424.0);
  ExpectError("=DATE(10000,1,1)", ErrorCode::Num);
}

TEST(DateTimeDateLimits, TimeLimits) {
  ExpectNumber("=TIME(32767,0,0)*24", 7.0);
  ExpectError("=TIME(32768,0,0)", ErrorCode::Num);
  ExpectError("=TIME(0,0,100000)", ErrorCode::Num);
}

TEST(DateTimeSerialZero, EomonthAndEdate) {
  ExpectNumber("=EOMONTH(0,0)", 31.0);
  ExpectNumber("=EOMONTH(0,7)", 244.0);
  ExpectError("=EOMONTH(0,-1)", ErrorCode::Num);
  Workbook wb = Workbook::create();
  const Value blank_start = EvalSourceIn("=EOMONTH(A1,0)", wb, wb.sheet(0));
  ASSERT_TRUE(blank_start.is_number());
  EXPECT_EQ(blank_start.as_number(), 31.0);
  ExpectNumber("=EDATE(0,1)", 31.0);
  ExpectNumber("=EDATE(0,0)", 0.0);
  ExpectNumber("=EOMONTH(1,-1)", 0.0);
}

TEST(DateTimeWeekday, TextReturnTypeCheckedBeforeSerialRange) {
  ExpectError("=WEEKDAY(-1,\"abc\")", ErrorCode::Value);
  ExpectError("=WEEKDAY(-1,1)", ErrorCode::Num);
}

TEST(DateTimeAnalysisToolPak, BooleansAreValue) {
  ExpectError("=EDATE(TRUE,1)", ErrorCode::Value);
  ExpectError("=EDATE(1,TRUE)", ErrorCode::Value);
  ExpectError("=EOMONTH(TRUE,1)", ErrorCode::Value);
  ExpectError("=NETWORKDAYS(TRUE,10)", ErrorCode::Value);
  ExpectError("=NETWORKDAYS(1,TRUE)", ErrorCode::Value);
  ExpectError("=WORKDAY(TRUE,1)", ErrorCode::Value);
  ExpectError("=WORKDAY(1,TRUE)", ErrorCode::Value);
  ExpectError("=WORKDAY(1,5,TRUE)", ErrorCode::Value);
  ExpectError("=WORKDAY.INTL(TRUE,5)", ErrorCode::Value);
  ExpectError("=NETWORKDAYS.INTL(TRUE,5)", ErrorCode::Value);
  ExpectError("=YEARFRAC(TRUE,10)", ErrorCode::Value);
  ExpectError("=YEARFRAC(1,10,TRUE)", ErrorCode::Value);
  ExpectError("=WEEKNUM(TRUE)", ErrorCode::Value);
  ExpectError("=WEEKNUM(10,FALSE)", ErrorCode::Value);
}

TEST(DateTimeAnalysisToolPak, OmittedRequiredIsNA) {
  ExpectError("=EDATE(,1)", ErrorCode::NA);
  ExpectError("=EDATE(1,)", ErrorCode::NA);
  ExpectError("=EDATE(,)", ErrorCode::NA);
  ExpectError("=EOMONTH(1,)", ErrorCode::NA);
  ExpectError("=NETWORKDAYS(1,)", ErrorCode::NA);
  ExpectError("=NETWORKDAYS(,)", ErrorCode::NA);
  ExpectError("=WORKDAY(,)", ErrorCode::NA);
  ExpectError("=WORKDAY(1,)", ErrorCode::NA);
  ExpectError("=WORKDAY.INTL(,5)", ErrorCode::NA);
  ExpectError("=NETWORKDAYS.INTL(,5)", ErrorCode::NA);
  ExpectError("=YEARFRAC(,1)", ErrorCode::NA);
  ExpectError("=YEARFRAC(1,)", ErrorCode::NA);
  ExpectError("=WEEKNUM(,1)", ErrorCode::NA);
  ExpectError("=WEEKNUM(,)", ErrorCode::NA);
}

TEST(DateTimeAnalysisToolPak, OmittedOptionalTakesDefault) {
  ExpectNumber("=WEEKNUM(10,)", 2.0);
  ExpectNumber("=YEARFRAC(1,400,)", 392.0 / 360.0);
  ExpectNumber("=WORKDAY.INTL(DATE(2025,1,1),10,,)", 45672.0);
  ExpectNumber("=WORKDAY.INTL(DATE(2025,1,1),10,,{45660})", 45673.0);
  ExpectNumber("=WORKDAY(DATE(2025,1,1),10,)", 45672.0);
  ExpectNumber("=NETWORKDAYS(DATE(2025,1,1),DATE(2025,1,31),)", 23.0);
}

TEST(DateTimeAnalysisToolPak, NonMembersKeepCoercion) {
  ExpectNumber("=DAYS(,5)", -5.0);
  ExpectNumber("=WEEKDAY(,1)", 7.0);
  ExpectError("=DATEDIF(1,,\"d\")", ErrorCode::Num);
  ExpectNumber("=DAY(TRUE)", 1.0);
  ExpectNumber("=HOUR(TRUE)", 0.0);
}

TEST(DateTimeIntlWeekend, Selectors) {
  ExpectNumber("=WORKDAY.INTL(DATE(2025,1,1),10,TRUE)", 45672.0);
  ExpectNumber("=NETWORKDAYS.INTL(DATE(2025,1,1),DATE(2025,12,31),TRUE)", 261.0);
  ExpectError("=WORKDAY.INTL(DATE(2025,1,1),10,FALSE)", ErrorCode::Num);
  ExpectError("=NETWORKDAYS.INTL(DATE(2025,1,1),DATE(2025,12,31),FALSE)", ErrorCode::Num);
  ExpectError("=WORKDAY.INTL(DATE(2025,1,1),10,\"1\")", ErrorCode::Value);
}

TEST(DateTimeIntlWeekend, BlankCellIsNumAndBlankHolidaysPass) {
  Workbook wb = Workbook::create();
  const Value blank_weekend = EvalSourceIn("=WORKDAY.INTL(DATE(2025,1,1),10,A1)", wb, wb.sheet(0));
  ASSERT_TRUE(blank_weekend.is_error());
  EXPECT_EQ(blank_weekend.as_error(), ErrorCode::Num);
  const Value blank_holidays = EvalSourceIn("=WORKDAY.INTL(DATE(2025,1,1),10,1,A1)", wb, wb.sheet(0));
  ASSERT_TRUE(blank_holidays.is_number());
  EXPECT_EQ(blank_holidays.as_number(), 45672.0);
}

TEST(DateTimeAnalysisToolPak, BooleanCellsAreValue) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::boolean(true));
  for (const char* formula : {"=EDATE(10,A1)", "=NETWORKDAYS(1,A1)", "=WORKDAY(1,5,A1)", "=YEARFRAC(1,10,A1)",
                              "=WEEKNUM(A1)", "=WORKDAY.INTL(A1,5)"}) {
    const Value v = EvalSourceIn(formula, wb, wb.sheet(0));
    ASSERT_TRUE(v.is_error()) << formula;
    EXPECT_EQ(v.as_error(), ErrorCode::Value) << formula;
  }
}

}  // namespace
}  // namespace eval
}  // namespace formulon
