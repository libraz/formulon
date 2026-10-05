// End-to-end date/time built-in tests: month shifts and day differences.

#include "builtins_datetime_test_helpers.h"
namespace formulon {
namespace eval {
namespace {

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

TEST(DateTimeDays, OutOfRangeSerialIsNum) {
  // A serial far beyond Excel's max date (9999-12-31 = 2958465) must be
  // rejected with #NUM! rather than reaching the int64 cast in
  // ymd_from_serial with undefined behavior.
  const Value v = EvalSource("=DAYS(1E300, DATE(2026, 4, 23))");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
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
