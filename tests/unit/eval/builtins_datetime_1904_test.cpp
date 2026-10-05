// End-to-end date/time built-in tests: 1904 epoch and raw text behavior.

#include <cmath>
#include <string>

#include "builtins_datetime_test_helpers.h"
namespace formulon {
namespace eval {
namespace {

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

TEST(DateTime1904, NetworkdaysCountsTheSameCalendarSpanUnderBothEpochs) {
  // 2024-01-01 (Mon) .. 2024-01-12 (Fri) is 10 business days regardless of
  // the workbook epoch. The 1900/1904 offset of 1462 days is not a multiple
  // of 7, so reading a 1904 serial as a 1900 one moves every weekend by two
  // days and the span answers 9 instead.
  const Value v1900 = EvalSource("=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,12))");
  ASSERT_TRUE(v1900.is_number());
  EXPECT_DOUBLE_EQ(v1900.as_number(), 10.0);

  const Value v1904 = EvalSource1904("=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,12))");
  ASSERT_TRUE(v1904.is_number());
  EXPECT_DOUBLE_EQ(v1904.as_number(), 10.0);
}

TEST(DateTime1904, NetworkdaysIntlWeekendMaskFollowsTheWorkbookEpoch) {
  // Weekend selector 11 is "Sunday only", so 2024-01-01 .. 2024-01-13
  // (Mon .. Sat) drops just 2024-01-07 and leaves 12 working days. Under an
  // epoch-blind weekday test the mask lands on the Saturdays instead and the
  // span answers 11.
  const Value v1900 = EvalSource("=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,13),11)");
  ASSERT_TRUE(v1900.is_number());
  EXPECT_DOUBLE_EQ(v1900.as_number(), 12.0);

  const Value v1904 = EvalSource1904("=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,13),11)");
  ASSERT_TRUE(v1904.is_number());
  EXPECT_DOUBLE_EQ(v1904.as_number(), 12.0);
}

TEST(DateTime1904, WorkdayLandsOnTheSameCalendarDateUnderBothEpochs) {
  // One working day after 2024-01-05 (Fri) is 2024-01-08 (Mon): the weekend
  // in between must be recognised in the workbook's own epoch.
  const Value v1900 = EvalSource("=WORKDAY(DATE(2024,1,5),1)");
  ASSERT_TRUE(v1900.is_number());
  const Value expected_1900 = EvalSource("=DATE(2024,1,8)");
  ASSERT_TRUE(expected_1900.is_number());
  EXPECT_DOUBLE_EQ(v1900.as_number(), expected_1900.as_number());

  const Value v1904 = EvalSource1904("=WORKDAY(DATE(2024,1,5),1)");
  ASSERT_TRUE(v1904.is_number());
  EXPECT_DOUBLE_EQ(v1904.as_number(), v1900.as_number() - 1462.0);
}

TEST(DateTime1904, WorkdayIntlWeekendMaskFollowsTheWorkbookEpoch) {
  // Weekend selector 11 is "Sunday only", so the day after 2024-01-05 (Fri)
  // is the working Saturday 2024-01-06.
  const Value v1900 = EvalSource("=WORKDAY.INTL(DATE(2024,1,5),1,11)");
  ASSERT_TRUE(v1900.is_number());
  const Value expected_1900 = EvalSource("=DATE(2024,1,6)");
  ASSERT_TRUE(expected_1900.is_number());
  EXPECT_DOUBLE_EQ(v1900.as_number(), expected_1900.as_number());

  const Value v1904 = EvalSource1904("=WORKDAY.INTL(DATE(2024,1,5),1,11)");
  ASSERT_TRUE(v1904.is_number());
  EXPECT_DOUBLE_EQ(v1904.as_number(), v1900.as_number() - 1462.0);
}

TEST(DateTime1904, WorkdayHolidaysAreReadInTheWorkbookEpoch) {
  // 2024-01-08 (Mon) as a holiday pushes the result to 2024-01-09, and the
  // holiday serial itself is expressed in the workbook's epoch.
  const Value v1900 = EvalSource("=WORKDAY(DATE(2024,1,5),1,DATE(2024,1,8))");
  ASSERT_TRUE(v1900.is_number());
  const Value expected_1900 = EvalSource("=DATE(2024,1,9)");
  ASSERT_TRUE(expected_1900.is_number());
  EXPECT_DOUBLE_EQ(v1900.as_number(), expected_1900.as_number());

  const Value v1904 = EvalSource1904("=WORKDAY(DATE(2024,1,5),1,DATE(2024,1,8))");
  ASSERT_TRUE(v1904.is_number());
  EXPECT_DOUBLE_EQ(v1904.as_number(), v1900.as_number() - 1462.0);
}

TEST(DateTime1904, WorkdayUpperEndpointIsEpochSpecific) {
  // 9999-12-31 is serial 2957003 under the 1904 system, so walking forward
  // from 2957000 runs off the calendar. The same serial is an ordinary day
  // under the 1900 system and the walk succeeds there.
  const Value over = EvalSource1904("=WORKDAY(2957000,10)");
  ASSERT_TRUE(over.is_error());
  EXPECT_EQ(over.as_error(), ErrorCode::Num);

  const Value within = EvalSource("=WORKDAY(2957000,10)");
  ASSERT_TRUE(within.is_number());
}

TEST(DateTime1904, NetworkdaysRejectsSerialsPastTheEpochUpperEndpoint) {
  const Value v = EvalSource1904("=NETWORKDAYS(1,2957004)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
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
