// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for DATEVALUE / TIMEVALUE. These cover the hand-rolled
// text parsers that convert Excel-style date/time strings into serial
// numbers. The expected values are hand-computed from the Excel 1900 epoch
// conventions (see `eval/date_time.cpp`) and cross-referenced against the
// assertions in builtins_datetime_test.cpp; the oracle suite below
// (tests/oracle/cases/datevalue_timevalue.yaml) exists as the second line
// of defence against locale-specific drift.
//
// Landmark serials used throughout:
//
//   2024-01-01 = 45292
//   2024-02-29 = 45351
//   2024-03-01 = 45352
//   2024-03-15 = 45366
//   2024-07-04 = 45477
//   2026-04-23 = 46135

#include <cmath>
#include <string_view>

#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

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
// DATEVALUE
// ---------------------------------------------------------------------------

TEST(Datevalue, IsoDash) {
  const Value v = EvalSource("=DATEVALUE(\"2024-03-15\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 45366.0);
}

TEST(Datevalue, IsoDashSingleDigitMonthDay) {
  // "2024-3-5" should parse identically to "2024-03-05".
  const Value v = EvalSource("=DATEVALUE(\"2024-3-5\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 45356.0);  // 45352 (Mar 1) + 4 = 45356
}

TEST(Datevalue, IsoWithTSeparatorIsValueError) {
  // Mac Excel 365 rejects the ISO 8601 'T' date/time separator with #VALUE!,
  // even when the underlying date would be valid. Formulon mirrors that:
  // use a space between date and time instead.
  const Value v = EvalSource("=DATEVALUE(\"2024-03-15T00:00:00\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(Datevalue, IsoWithSpaceAndTime) {
  // "2024-03-15 13:30" -> 45366 (time stripped).
  const Value v = EvalSource("=DATEVALUE(\"2024-03-15 13:30\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 45366.0);
}

TEST(Datevalue, SlashForm) {
  const Value v = EvalSource("=DATEVALUE(\"2024/3/15\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 45366.0);
}

TEST(Datevalue, SlashFormPadded) {
  const Value v = EvalSource("=DATEVALUE(\"2024/03/15\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 45366.0);
}

TEST(Datevalue, KanjiFormWithTerminator) {
  // UTF-8 for 2024年3月15日.
  const Value v = EvalSource(
      "=DATEVALUE(\"2024\xE5\xB9\xB4"
      "3\xE6\x9C\x88"
      "15\xE6\x97\xA5\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 45366.0);
}

TEST(Datevalue, KanjiFormRequiresTerminator) {
  // Mac Excel 365 rejects the kanji form when the trailing 日 is missing.
  // "2024年03月15" is #VALUE!; use "2024年03月15日" instead.
  const Value v = EvalSource(
      "=DATEVALUE(\"2024\xE5\xB9\xB4"
      "03\xE6\x9C\x88"
      "15\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(Datevalue, LeadingAsciiWhitespaceTrimmed) {
  const Value v = EvalSource("=DATEVALUE(\"   2024-03-15\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 45366.0);
}

TEST(Datevalue, TrailingAsciiWhitespaceTrimmed) {
  const Value v = EvalSource("=DATEVALUE(\"2024-03-15   \")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 45366.0);
}

TEST(Datevalue, FullwidthSpaceNotTrimmed) {
  // Mac Excel 365 rejects strings surrounded by U+3000 fullwidth spaces
  // with #VALUE!. Formulon trims only ASCII whitespace to match.
  const Value v = EvalSource(
      "=DATEVALUE(\"\xE3\x80\x80"
      "2024-03-15\xE3\x80\x80\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(Datevalue, TwoDigitYearPivotBefore30IsTwentyFirstCentury) {
  // "29-03-15" -> 2029-03-15 (Excel DATEVALUE pivots 00..29 -> 2000..2029).
  // 2029-03-15 = DATE(2029,3,15). Serial: 2024-03-15 = 45366, plus
  // 365*5 + 2 leap-adjust (2024 leap)... easier: compute offsets from
  // 2024-01-01 = 45292. 2025 is not leap, 2026 not, 2027 not, 2028 leap.
  // 2024 had 366, 2025 has 365, 2026 365, 2027 365, 2028 366. So
  // 2029-01-01 = 45292 + 366+365+365+365+366 = 45292 + 1827 = 47119.
  // 2029-03-15 = 47119 + 31 + 28 + 14 = 47192.
  const Value v = EvalSource("=DATEVALUE(\"29-03-15\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 47192.0);
}

TEST(Datevalue, TwoDigitYearPivotAfter30IsNineteenthCentury) {
  // "95-07-04" -> 1995-07-04. Cross-checked with DATE(1995, 7, 4):
  // days(1900-01-01 -> 1995-01-01) = 95*365 + 23 leap days + 1 ghost day
  // (Excel 1900 leap-year bug) = 34699, so serial(1995-01-01) = 34700 and
  // serial(1995-07-04) = 34700 + 184 = 34884.
  const Value a = EvalSource("=DATEVALUE(\"95-07-04\")");
  const Value b = EvalSource("=DATE(1995,7,4)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_EQ(a.as_number(), b.as_number());
  EXPECT_EQ(a.as_number(), 34884.0);
}

TEST(Datevalue, EmptyStringIsValueError) {
  const Value v = EvalSource("=DATEVALUE(\"\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(Datevalue, WhitespaceOnlyIsValueError) {
  const Value v = EvalSource("=DATEVALUE(\"   \")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(Datevalue, MalformedStringIsValueError) {
  const Value v = EvalSource("=DATEVALUE(\"hello world\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(Datevalue, MonthOutOfRangeIsValueError) {
  const Value v = EvalSource("=DATEVALUE(\"2024-13-01\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(Datevalue, DayOutOfRangeIsValueError) {
  const Value v = EvalSource("=DATEVALUE(\"2024-02-30\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(Datevalue, MixedAsciiSeparatorsAccepted) {
  // Mac Excel is lenient about mixing '-' and '/' between the two date
  // positions. We match that leniency: "2024-03/15" parses the same as
  // "2024-03-15" or "2024/03/15".
  const Value v = EvalSource("=DATEVALUE(\"2024-03/15\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 45366.0);
}

TEST(Datevalue, TimeOnlyInputIsValueError) {
  // DATEVALUE with a time-only string: Excel would default to today's
  // date. We reject without a clock; see datetime.cpp banner comment and
  // the divergence note in tests/divergence.yaml.
  const Value v = EvalSource("=DATEVALUE(\"13:30\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(Datevalue, BooleanCoercesAndFails) {
  // Boolean coerces to text "TRUE" which isn't a date -> #VALUE!.
  const Value v = EvalSource("=DATEVALUE(TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(Datevalue, ErrorArgumentPropagates) {
  const Value v = EvalSource("=DATEVALUE(#DIV/0!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(Datevalue, LeapDayParses) {
  // 2024-02-29 is a real leap day.
  const Value v = EvalSource("=DATEVALUE(\"2024-02-29\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 45351.0);
}

TEST(Datevalue, NonLeapDayFeb29IsValueError) {
  // 2023-02-29 does not exist.
  const Value v = EvalSource("=DATEVALUE(\"2023-02-29\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// TIMEVALUE
// ---------------------------------------------------------------------------

TEST(Timevalue, HoursAndMinutes) {
  // 13:30 = (13*3600 + 30*60) / 86400 = 48600/86400 = 0.5625.
  const Value v = EvalSource("=TIMEVALUE(\"13:30\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.5625);
}

TEST(Timevalue, HoursMinutesSeconds) {
  // 13:30:45 = 48645 / 86400.
  const Value v = EvalSource("=TIMEVALUE(\"13:30:45\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 48645.0 / 86400.0, 1e-12);
}

TEST(Timevalue, Midnight) {
  const Value v = EvalSource("=TIMEVALUE(\"00:00\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(Timevalue, Noon) {
  const Value v = EvalSource("=TIMEVALUE(\"12:00\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.5);
}

TEST(Timevalue, AmPmOneThirtyPm) {
  const Value v = EvalSource("=TIMEVALUE(\"1:30 PM\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.5625);
}

TEST(Timevalue, AmPmTwelveAmIsMidnight) {
  // Excel: 12:00 AM -> 0.0.
  const Value v = EvalSource("=TIMEVALUE(\"12:00 AM\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(Timevalue, AmPmTwelvePmIsNoon) {
  // Excel: 12:00 PM -> 0.5.
  const Value v = EvalSource("=TIMEVALUE(\"12:00 PM\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.5);
}

TEST(Timevalue, AmPmLowercase) {
  const Value v = EvalSource("=TIMEVALUE(\"1:30 pm\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.5625);
}

TEST(Timevalue, AmPmWithoutSpaceIsValueError) {
  // Mac Excel 365 requires at least one ASCII whitespace between the time
  // digits and the AM/PM marker. "1:30PM" is #VALUE!; "1:30 PM" works.
  const Value v = EvalSource("=TIMEVALUE(\"1:30PM\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(Timevalue, HoursAbove24WrapModulo) {
  // "25:00" -> 25h * 3600 / 86400 = 1.04166..., reduced modulo 1.0 -> 0.04166...
  const Value v = EvalSource("=TIMEVALUE(\"25:00\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0 / 24.0, 1e-12);
}

TEST(Timevalue, EmbeddedInDateReturnsFractionalPart) {
  // "2024-03-15 13:30" -> TIMEVALUE returns only the fractional part.
  const Value v = EvalSource("=TIMEVALUE(\"2024-03-15 13:30\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.5625);
}

TEST(Timevalue, IsoTSeparatorIsValueError) {
  // Mac Excel 365 rejects the ISO 8601 'T' date/time separator for both
  // DATEVALUE and TIMEVALUE. Use a space separator instead.
  const Value v = EvalSource("=TIMEVALUE(\"2024-03-15T06:00:00\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(Timevalue, FractionalSeconds) {
  // "00:00:00.5" -> 0.5 seconds of one day = 0.5 / 86400.
  const Value v = EvalSource("=TIMEVALUE(\"00:00:00.5\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.5 / 86400.0, 1e-15);
}

TEST(Timevalue, LeadingAsciiWhitespaceTrimmed) {
  const Value v = EvalSource("=TIMEVALUE(\"   13:30\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.5625);
}

TEST(Timevalue, TrailingFullwidthSpaceAccepted) {
  // Mac Excel 365 tolerates a trailing U+3000 (fullwidth space) AFTER a
  // successfully parsed time token: `TIMEVALUE("13:30　")` returns 0.5625.
  // Note the asymmetry with DATEVALUE, which rejects strings SURROUNDED by
  // U+3000 (leading U+3000 is not tokenisable as the start of a date).
  const Value v = EvalSource("=TIMEVALUE(\"13:30\xE3\x80\x80\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.5625);
}

TEST(Timevalue, EmptyStringIsValueError) {
  const Value v = EvalSource("=TIMEVALUE(\"\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(Timevalue, MinutesOver59CarryIntoHours) {
  // Excel applies TIME()-style carry semantics for out-of-range minutes:
  // "12:60" becomes 13:00 = 13/24 = 0.5416666...
  const Value v = EvalSource("=TIMEVALUE(\"12:60\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 13.0 / 24.0, 1e-12);
}

TEST(Timevalue, SecondsOver59CarryIntoMinutes) {
  // "12:30:60" carries the overflow second into minutes -> 12:31:00
  // = (12*3600 + 31*60) / 86400 = 45060/86400 = 0.52152777...
  const Value v = EvalSource("=TIMEVALUE(\"12:30:60\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), (12.0 * 3600.0 + 31.0 * 60.0) / 86400.0, 1e-12);
}

TEST(Timevalue, MalformedTimeIsValueError) {
  const Value v = EvalSource("=TIMEVALUE(\"not a time\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(Timevalue, AmPmWithHourZeroAccepted) {
  // Mac Excel 365 accepts "0:30 AM" as 00:30 = 30/1440 ≈ 0.02083...
  // (the AM/PM modifier is ignored when the hour is already 0).
  const Value v = EvalSource("=TIMEVALUE(\"0:30 AM\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 30.0 / 1440.0, 1e-12);
}

TEST(Timevalue, HourZeroWithoutAmPmAgreesWithAmVariant) {
  // "0:30 AM" and "0:30" should give the same answer.
  const Value a = EvalSource("=TIMEVALUE(\"0:30 AM\")");
  const Value b = EvalSource("=TIMEVALUE(\"0:30\")");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_EQ(a.as_number(), b.as_number());
}

TEST(Timevalue, AmPmWithHourAbove12IsValueError) {
  // 13:30 PM is invalid in 12-hour mode.
  const Value v = EvalSource("=TIMEVALUE(\"13:30 PM\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(Timevalue, DateOnlyInputIsValueError) {
  // TIMEVALUE with a date-only string: must have an explicit time.
  const Value v = EvalSource("=TIMEVALUE(\"2024-03-15\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(Timevalue, BooleanCoercesAndFails) {
  const Value v = EvalSource("=TIMEVALUE(TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(Timevalue, ErrorArgumentPropagates) {
  const Value v = EvalSource("=TIMEVALUE(#N/A)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(Timevalue, JustBeforeMidnight) {
  // 23:59:59 = 86399 / 86400.
  const Value v = EvalSource("=TIMEVALUE(\"23:59:59\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 86399.0 / 86400.0, 1e-12);
}

// ---------------------------------------------------------------------------
// Round-trip: DATEVALUE / TIMEVALUE agree with DATE / TIME.
// ---------------------------------------------------------------------------

TEST(DatevalueRoundTrip, AgreesWithDate) {
  const Value a = EvalSource("=DATEVALUE(\"2026-04-23\")");
  const Value b = EvalSource("=DATE(2026,4,23)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_EQ(a.as_number(), b.as_number());
}

TEST(TimevalueRoundTrip, AgreesWithTime) {
  const Value a = EvalSource("=TIMEVALUE(\"14:30:15\")");
  const Value b = EvalSource("=TIME(14,30,15)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_NEAR(a.as_number(), b.as_number(), 1e-12);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
