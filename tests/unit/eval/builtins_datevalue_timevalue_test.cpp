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

#include <array>
#include <cmath>
#include <string>
#include <string_view>

#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "util/test_arena.h"
#include "util/test_eval_helpers.h"
#include "utils/arena.h"
#include "utils/date_time.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

using formulon::test::EvalSource;

Value EvalSourceWithProfile(std::string_view src, ExcelProfile profile);

// ---------------------------------------------------------------------------
// DATEVALUE
// ---------------------------------------------------------------------------

TEST(Datevalue, IsoDash) {
  const Value v = EvalSource("=DATEVALUE(\"2024-03-15\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 45366.0);
}

TEST(Datevalue, AcceptsExcelGhostFeb29_1900) {
  // Excel preserves Lotus 1-2-3's fictitious 1900-02-29 as serial 60. The
  // Gregorian calendar says 1900 is not a leap year, but DATEVALUE must
  // still accept "29-Feb-1900" for parity with Excel 365.
  const Value v = EvalSource("=DATEVALUE(\"1900-2-29\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 60.0);
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
  const Value v = EvalSourceWithProfile(
      "=DATEVALUE(\"2024\xE5\xB9\xB4"
      "3\xE6\x9C\x88"
      "15\xE6\x97\xA5\")",
      mac_365_ja_jp_profile());
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
  const Value v = EvalSourceWithProfile("=DATEVALUE(\"29-03-15\")", mac_365_ja_jp_profile());
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 47192.0);
}

TEST(Datevalue, TwoDigitYearPivotAfter30IsNineteenthCentury) {
  // "95-07-04" -> 1995-07-04. Cross-checked with DATE(1995, 7, 4):
  // days(1900-01-01 -> 1995-01-01) = 95*365 + 23 leap days + 1 ghost day
  // (Excel 1900 leap-year bug) = 34699, so serial(1995-01-01) = 34700 and
  // serial(1995-07-04) = 34700 + 184 = 34884.
  const Value a = EvalSourceWithProfile("=DATEVALUE(\"95-07-04\")", mac_365_ja_jp_profile());
  const Value b = EvalSourceWithProfile("=DATE(1995,7,4)", mac_365_ja_jp_profile());
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

TEST(Datevalue, TimeOnlyInputIsWholeDays) {
  // DATEVALUE is the integer part of the recognised value, so a time-only
  // string counts whole days from 0.
  const Value v = EvalSource("=DATEVALUE(\"13:30\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
  const Value overflow = EvalSource("=DATEVALUE(\"25:00\")");
  ASSERT_TRUE(overflow.is_number());
  EXPECT_EQ(overflow.as_number(), 1.0);
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

TEST(Timevalue, DateOnlyInputIsZero) {
  // TIMEVALUE is the fractional part of the recognised value; a date-only
  // string has none.
  const Value v = EvalSource("=TIMEVALUE(\"2024-03-15\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
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

// ---------------------------------------------------------------------------
// DATEVALUE: d-mmm-yyyy / d mmm yyyy / d/mmm/yyyy forms.
//
// These unblock IronCalc oracle cases that feed Excel-style textual dates
// through DATEVALUE / VALUE. The month must be spelled with English letters
// (3-letter abbreviation or full name, case-insensitive). The separator must
// be consistent on both sides of the month word.
// ---------------------------------------------------------------------------

TEST(DatevalueDmmm, DashGhostFeb29_1900) {
  // Excel's Lotus-1-2-3 compatibility ghost day: 1900-02-29 = serial 60.
  const Value v = EvalSource("=DATEVALUE(\"29-Feb-1900\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 60.0);
}

TEST(DatevalueDmmm, DashMar1_1900) {
  // Day immediately after the ghost day.
  const Value v = EvalSource("=DATEVALUE(\"1-Mar-1900\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 61.0);
}

TEST(DatevalueDmmm, DashFeb2_2026) {
  // IronCalc VALUE("2-Feb-2026") -> 46055.
  const Value v = EvalSource("=DATEVALUE(\"2-Feb-2026\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 46055.0);
}

TEST(DatevalueDmmm, SpaceSeparatedShortMonth) {
  // "15 Dec 2024" = 2024-12-15 = 45641.
  const Value v = EvalSource("=DATEVALUE(\"15 Dec 2024\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 45641.0);
}

TEST(DatevalueDmmm, DashFullMonthLongestMatchWins) {
  // "1-January-2024" = 2024-01-01 = 45292. Ensures the full name is
  // preferred over the 3-letter prefix ("Jan").
  const Value v = EvalSource("=DATEVALUE(\"1-January-2024\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 45292.0);
}

TEST(DatevalueDmmm, SlashSeparatorAllowed) {
  // Slashes must match on both sides of the month word.
  const Value v = EvalSource("=DATEVALUE(\"15/Dec/2024\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 45641.0);
}

TEST(DatevalueDmmm, CaseInsensitiveMonthName) {
  const Value v = EvalSource("=DATEVALUE(\"1-MAR-1900\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 61.0);
}

TEST(DatevalueDmmm, PurelyNumericDmyRejected) {
  // "10-12-1999" has no letter month token and must NOT be accepted via the
  // d-mmm-yyyy fall-back. The yyyy-first path also cannot consume it
  // cleanly (the trailing "99" bytes leak through), so DATEVALUE returns
  // #VALUE!.
  const Value v = EvalSourceWithProfile("=DATEVALUE(\"10-12-1999\")", mac_365_ja_jp_profile());
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(DatevalueDmmm, MixedSeparatorsAccepted) {
  // The gaps around the month name take `-` and `/` independently.
  const Value v = EvalSource("=DATEVALUE(\"29-Feb/1900\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 60.0);
}

TEST(DatevalueDmmm, MonthFirstOrderRejected) {
  // Formulon intentionally does not accept "Feb-29-1900" (month-first).
  const Value v = EvalSource("=DATEVALUE(\"Feb-29-1900\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(DatevalueDmmm, IsoPathStillWorks) {
  // Regression guard: the primary yyyy-first path remains untouched.
  const Value v = EvalSource("=DATEVALUE(\"2024-03-15\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 45366.0);
}

// ---------------------------------------------------------------------------
// DATEVALUE year-less date_text ("3/15", "3月15日"): Microsoft's documented
// rule is "If the year portion of the date_text argument is omitted,
// DATEVALUE uses the current year". Assertable only with a pinned clock
// (see builtins_datetime_test.cpp's ClockSeam section for the same seam).
// ---------------------------------------------------------------------------

constexpr date_time::CivilTime kPinnedNowForDatevalue{{2026, 4U, 23U}, {15U, 30U, 45U}};

Value EvalSourcePinned(std::string_view src, date_time::CivilTime pinned = kPinnedNowForDatevalue,
                       ExcelProfile profile = default_excel_profile()) {
  Arena& parse_arena = formulon::test::test_parse_arena();
  Arena& eval_arena = formulon::test::test_eval_arena();
  parser::Parser p(src, parse_arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return evaluate(*root, eval_arena, default_registry(),
                  EvalContext().with_pinned_now(pinned).with_excel_profile(profile));
}

struct DateProfileCase {
  ExcelProfile profile;
  const char* id;
  bool japanese;
};

constexpr std::array<DateProfileCase, 4> kDateProfiles = {{
    {mac_365_ja_jp_profile(), "mac-365-ja_JP", true},
    {win_365_ja_jp_profile(), "win-365-ja_JP", true},
    {mac_365_en_us_profile(), "mac-365-en_US", false},
    {win_365_en_us_profile(), "win-365-en_US", false},
}};

Value EvalSourceWithProfile(std::string_view src, ExcelProfile profile) {
  Arena& parse_arena = formulon::test::test_parse_arena();
  Arena& eval_arena = formulon::test::test_eval_arena();
  parser::Parser p(src, parse_arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return evaluate(*root, eval_arena, default_registry(), EvalContext().with_excel_profile(profile));
}

TEST(DatevalueProfile, TwoDigitYearUsesLocaleOrderForDatevalueAndValue) {
  for (const DateProfileCase& profile : kDateProfiles) {
    SCOPED_TRACE(profile.id);
    const Value date = EvalSourceWithProfile(profile.japanese ? "=DATEVALUE(\"29/3/15\")" : "=DATEVALUE(\"3/15/29\")",
                                             profile.profile);
    ASSERT_TRUE(date.is_number());
    EXPECT_DOUBLE_EQ(date.as_number(), 47192.0);

    const Value value =
        EvalSourceWithProfile(profile.japanese ? "=VALUE(\"29-3-15\")" : "=VALUE(\"3-15-29\")", profile.profile);
    ASSERT_TRUE(value.is_number());
    EXPECT_DOUBLE_EQ(value.as_number(), 47192.0);
  }
}

TEST(DatevalueProfile, EnglishDateOrderDoesNotFallBackToYearFirst) {
  for (const DateProfileCase& profile : kDateProfiles) {
    if (profile.japanese) {
      continue;
    }
    SCOPED_TRACE(profile.id);
    const Value date = EvalSourceWithProfile("=DATEVALUE(\"29/3/15\")", profile.profile);
    ASSERT_TRUE(date.is_error());
    EXPECT_EQ(date.as_error(), ErrorCode::Value);

    const Value value = EvalSourceWithProfile("=VALUE(\"29-3-15\")", profile.profile);
    ASSERT_TRUE(value.is_error());
    EXPECT_EQ(value.as_error(), ErrorCode::Value);
  }
}

TEST(DatevalueProfile, FourDigitIsoDatesRemainYearFirst) {
  for (const DateProfileCase& profile : kDateProfiles) {
    SCOPED_TRACE(profile.id);
    const Value value = EvalSourceWithProfile("=DATEVALUE(\"2024/3/15\")", profile.profile);
    ASSERT_TRUE(value.is_number());
    EXPECT_DOUBLE_EQ(value.as_number(), 45366.0);
  }
}

TEST(DatevalueProfile, JapaneseDateTimeFormsAreLocaleGated) {
  struct Case {
    const char* formula;
    double expected;
  };
  constexpr Case cases[] = {
      {"=DATEVALUE(\"2024年3月15日\")", 45366.0},
      {"=DATEVALUE(\"令和6年4月1日\")", 45383.0},
      {"=DATEVALUE(\"２０２４/３/１５\")", 45366.0},
      {"=TIMEVALUE(\"13時30分\")", 0.5625},
  };
  for (const DateProfileCase& profile : kDateProfiles) {
    for (const Case& test_case : cases) {
      SCOPED_TRACE(profile.id);
      SCOPED_TRACE(test_case.formula);
      const Value value = EvalSourceWithProfile(test_case.formula, profile.profile);
      if (profile.japanese) {
        ASSERT_TRUE(value.is_number());
        EXPECT_DOUBLE_EQ(value.as_number(), test_case.expected);
      } else {
        ASSERT_TRUE(value.is_error());
        EXPECT_EQ(value.as_error(), ErrorCode::Value);
      }
    }
  }
}

TEST(DatevalueYearLess, SlashSeparatedUsesPinnedCurrentYear) {
  const Value v = EvalSourcePinned("=DATEVALUE(\"3/15\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 46096.0);  // 2026-03-15
}

TEST(DatevalueYearLess, DashSeparatedUsesPinnedCurrentYear) {
  const Value v = EvalSourcePinned("=DATEVALUE(\"3-15\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 46096.0);
}

TEST(DatevalueYearLess, KanjiFormUsesPinnedCurrentYear) {
  const Value v = EvalSourcePinned(
      "=DATEVALUE(\"3\xE6\x9C\x88"
      "15\xE6\x97\xA5\")",
      kPinnedNowForDatevalue, mac_365_ja_jp_profile());
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 46096.0);
}

TEST(DatevalueYearLess, InvalidMonthDayStillValueError) {
  const Value v = EvalSourcePinned("=DATEVALUE(\"13/40\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(DatevalueYearLess, FullYearFormStillTakesPriorityOverMonthDay) {
  // "2024/3/15" must not be reinterpreted as month/day off a truncated
  // read; the 3-component yyyy-first parser still wins.
  const Value v = EvalSourcePinned("=DATEVALUE(\"2024/3/15\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 45366.0);  // 2024-03-15, independent of the pinned year
}

TEST(DatevalueYearLess, PinnedClockReachesContextFreeCoercion) {
  const Value plus = EvalSourcePinned("=\"3/15\"+0");
  ASSERT_TRUE(plus.is_number());
  EXPECT_EQ(plus.as_number(), 46096.0);
  const Value abs = EvalSourcePinned("=ABS(\"3/15\")");
  ASSERT_TRUE(abs.is_number());
  EXPECT_EQ(abs.as_number(), 46096.0);

  // Feb 29 exists only in a leap pinned year, so TIMEVALUE shows which year it
  // read; a Year-Month-Day profile has no m/y reading to fall back to.
  constexpr date_time::CivilTime kLeapYear{{2024, 4U, 23U}, {15U, 30U, 45U}};
  const Value leap = EvalSourcePinned("=TIMEVALUE(\"2/29 6:00\")", kLeapYear, mac_365_ja_jp_profile());
  ASSERT_TRUE(leap.is_number());
  EXPECT_EQ(leap.as_number(), 0.25);
  const Value common = EvalSourcePinned("=TIMEVALUE(\"2/29 6:00\")", kPinnedNowForDatevalue, mac_365_ja_jp_profile());
  ASSERT_TRUE(common.is_error());
  EXPECT_EQ(common.as_error(), ErrorCode::Value);
}

TEST(DatevalueYearLess, WithoutAPinnedClockStillUsesTheHostYear) {
  // The contextless entry point (no EvalContext offered at all) falls back
  // to the host clock rather than rejecting the year-less form outright.
  const Value v = EvalSource("=YEAR(DATEVALUE(\"3/15\"))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), static_cast<double>(date_time::host_civil_time().date.y));
}

// ---------------------------------------------------------------------------
// Hour-only and dotted 12-hour forms, shared by TIMEVALUE / VALUE / DATEVALUE /
// arithmetic coercion.
// ---------------------------------------------------------------------------

TEST(TimeTextMeridiem, AcceptedForms) {
  struct Case {
    const char* formula;
    double expected;
  };
  const Case cases[] = {
      {"=TIMEVALUE(\"6 PM\")", 0.75},
      {"=TIMEVALUE(\"6 AM\")", 0.25},
      {"=TIMEVALUE(\"6 P\")", 0.75},
      {"=TIMEVALUE(\"12 PM\")", 0.5},
      {"=TIMEVALUE(\"12 AM\")", 0.0},
      {"=TIMEVALUE(\"0 PM\")", 0.5},
      {"=TIMEVALUE(\"4 pm\")", 4.0 / 24.0 + 0.5},
      {"=TIMEVALUE(\"6:00 PM.\")", 0.75},
      {"=TIMEVALUE(\"6:30 pm.\")", 0.7708333333333334},
      {"=TIMEVALUE(\"6 pm .\")", 0.75},
      {"=TIMEVALUE(\"2026/7/7 6 PM\")", 0.75},
      {"=TIMEVALUE(\"2026/7/7 6 pm.\")", 0.75},
      {"=TIMEVALUE(\"6 PM 2026/7/7\")", 0.75},
      {"=DATEVALUE(\"2026/7/7 6 PM\")", 46210.0},
      {"=DATEVALUE(\"6 PM 2026/7/7\")", 46210.0},
      {"=VALUE(\"6 PM\")", 0.75},
      {"=VALUE(\"6:30 PM.\")", 0.7708333333333334},
      {"=\"6 PM\"+0", 0.75},
  };
  for (const Case& c : cases) {
    const Value v = EvalSource(c.formula);
    ASSERT_TRUE(v.is_number()) << c.formula;
    EXPECT_NEAR(v.as_number(), c.expected, 1e-15) << c.formula;
  }
}

TEST(TimeTextMeridiem, RejectedForms) {
  const char* const texts[] = {"6PM", "18 PM", "13:00 PM", "6.", "6", "6 a.m.", "6 p.m.", "AM 6", "6 pmx", "6 Px"};
  for (const char* text : texts) {
    const std::string formula = std::string("=TIMEVALUE(\"") + text + "\")";
    const Value v = EvalSource(formula);
    ASSERT_TRUE(v.is_error()) << formula;
    EXPECT_EQ(v.as_error(), ErrorCode::Value) << formula;
  }
}

TEST(TimeTextMeridiem, MonthWordAfterTimeIsNotAMarker) {
  const Value v = EvalSource("=TIMEVALUE(\"10:00 Apr 5, 2024\")");
  ASSERT_TRUE(v.is_number());
  // TIMEVALUE is v - INT(v) of the date-time serial, last bits included.
  const double serial = 45387.0 + 10.0 / 24.0;
  EXPECT_EQ(v.as_number(), serial - std::floor(serial));
}

TEST(DateTextLenientShapes, HourBeyondTwentyThreeRollsTheDateForward) {
  for (const char* text : {"2026-01-01 24:00:00", "2026-01-01 25:00:00"}) {
    const Value v = EvalSource(std::string("=DATEVALUE(\"") + text + "\")");
    ASSERT_TRUE(v.is_number()) << text;
    EXPECT_EQ(v.as_number(), 46024.0) << text;
  }
}

TEST(DateTextLenientShapes, BlanksAroundDateSeparators) {
  const Value v = EvalSource("=DATEVALUE(\"2026 - 01 - 01\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 46023.0);
}

TEST(DateTextLenientShapes, BlanksAroundTimeSeparator) {
  const Value v = EvalSource("=TIMEVALUE(\"4 : 35\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.1909722222222222, 1e-15);
}

// ---------------------------------------------------------------------------
// Mac Excel ja-JP recogniser, clock pinned to 2026-10-15: a non-leap year,
// which decides `Feb 29` and `29/2`. Expected values are Excel's.
// ---------------------------------------------------------------------------

constexpr date_time::CivilTime kPinnedJaNow{{2026, 10U, 15U}, {12U, 0U, 0U}};

struct JaNumberCase {
  const char* formula;
  double expected;
};

void ExpectJaNumbers(const JaNumberCase* cases, std::size_t count) {
  for (std::size_t i = 0; i < count; ++i) {
    const Value v = EvalSourcePinned(cases[i].formula, kPinnedJaNow, mac_365_ja_jp_profile());
    ASSERT_TRUE(v.is_number()) << cases[i].formula;
    EXPECT_EQ(v.as_number(), cases[i].expected) << cases[i].formula;
  }
}

void ExpectJaValueErrors(const char* const* formulas, std::size_t count) {
  for (std::size_t i = 0; i < count; ++i) {
    const Value v = EvalSourcePinned(formulas[i], kPinnedJaNow, mac_365_ja_jp_profile());
    ASSERT_TRUE(v.is_error()) << formulas[i];
    EXPECT_EQ(v.as_error(), ErrorCode::Value) << formulas[i];
  }
}

TEST(DateTextJaJp, MonthNameForms) {
  const JaNumberCase cases[] = {
      {"=VALUE(\"Sept 1\")", 46266.0},  {"=VALUE(\"Jan1\")", 46023.0},       {"=VALUE(\"1-Jan\")", 46023.0},
      {"=VALUE(\"Jan - 1\")", 46023.0}, {"=VALUE(\"Jan 32\")", 11689.0},     {"=VALUE(\"Feb 29\")", 47150.0},
      {"=VALUE(\"Jan 0\")", 36526.0},   {"=VALUE(\"Jan 2026\")", 46023.0},   {"=VALUE(\"Jan 1 , 2026\")", 46023.0},
      {"=VALUE(\"1 Jan 1\")", 36892.0}, {"=VALUE(\"1-Jan/2026\")", 46023.0}, {"=VALUE(\"1Jan2026\")", 46023.0},
  };
  ExpectJaNumbers(cases, sizeof(cases) / sizeof(cases[0]));
  const char* const rejected[] = {
      "=VALUE(\"Ma 1\")",       "=VALUE(\"Jan.1\")",    "=VALUE(\"Jan 100\")",     "=VALUE(\"Jan 1 2026\")",
      "=VALUE(\"Jan 1,2026\")", "=VALUE(\"32 Jan\")",   "=VALUE(\"29 Feb\")",      "=VALUE(\"Mon Jan 1\")",
      "=VALUE(\"Junee 1\")",    "=VALUE(\"2026 Jan\")", "=VALUE(\"1 Jan, 2026\")",
  };
  ExpectJaValueErrors(rejected, sizeof(rejected) / sizeof(rejected[0]));
}

TEST(DateTextJaJp, NumericPairsAndTriples) {
  const JaNumberCase cases[] = {
      {"=VALUE(\"13/1\")", 46035.0},   {"=VALUE(\"31/4\")", 11414.0},   {"=VALUE(\"32/1\")", 11689.0},
      {"=VALUE(\"29/2\")", 47150.0},   {"=VALUE(\"2026/1\")", 46023.0}, {"=VALUE(\"1 / 2\")", 46024.0},
      {"=VALUE(\"29/1/1\")", 47119.0}, {"=VALUE(\"30/1/1\")", 10959.0}, {"=VALUE(\"1900/2/29\")", 60.0},
      {"=VALUE(\"1 -1/2\")", 36893.0},
  };
  ExpectJaNumbers(cases, sizeof(cases) / sizeof(cases[0]));
  const char* const rejected[] = {
      "=VALUE(\"1/2026\")", "=VALUE(\"0/1\")",      "=VALUE(\"2/29\")", "=VALUE(\"2026/1/001\")",
      "=VALUE(\"123/1\")",  "=VALUE(\"1/2/0026\")", "=VALUE(\"1/2-\")", "=VALUE(\"2026.1.1\")",
  };
  ExpectJaValueErrors(rejected, sizeof(rejected) / sizeof(rejected[0]));
}

TEST(DateTextJaJp, KanjiAndEraForms) {
  const JaNumberCase cases[] = {
      {"=VALUE(\"2026\xE5\xB9\xB4"
       "1\xE6\x9C\x88\")",
       46023.0},
      {"=VALUE(\"026\xE5\xB9\xB4"
       "001\xE6\x9C\x88"
       "0001\xE6\x97\xA5\")",
       46023.0},
      {"=VALUE(\"R8/1/1\")", 46023.0},
      {"=VALUE(\"R0008.01.001\")", 46023.0},
      {"=VALUE(\"\xEF\xBD\x88"
       "30.1.1\")",
       43101.0},
      {"=VALUE(\"001\xE6\x99\x82"
       "030\xE5\x88\x86"
       "005\xE7\xA7\x92\")",
       0.062557870370370375},
  };
  ExpectJaNumbers(cases, sizeof(cases) / sizeof(cases[0]));
  const char* const rejected[] = {
      "=VALUE(\"\xE4\xBB\xA4\xE5\x92\x8C"
      "8\xE5\xB9\xB4"
      "1\xE6\x9C\x88\")",
      "=VALUE(\"\xE6\x98\x8E\xE6\xB2\xBB"
      "6\xE5\xB9\xB4"
      "1\xE6\x9C\x88"
      "1\xE6\x97\xA5\")",
      "=VALUE(\"R100.1.1\")",
      "=VALUE(\"M6.1.1\")",
      "=VALUE(\"R8.1.1 1:30\")",
      "=VALUE(\"2026\xE5\xB9\xB4"
      "1\xE6\x9C\x88"
      "1\xE6\x97\xA5 1:30\")",
      "=VALUE(\"1/1 1\xE6\x99\x82"
      "30\xE5\x88\x86\")",
      "=VALUE(\"24\xE6\x99\x82"
      "00\xE5\x88\x86\")",
  };
  ExpectJaValueErrors(rejected, sizeof(rejected) / sizeof(rejected[0]));
}

TEST(DateTextJaJp, TimeForms) {
  const JaNumberCase cases[] = {
      {"=VALUE(\"1:\")", 0.041666666666666664},
      {"=VALUE(\"1:30.5\")", 0.0010474537037037037},
      {"=VALUE(\"1:30:00.0005\")", 0.062500011574074074},
      {"=VALUE(\"1:30:00.0004\")", 0.0625},
      {"=VALUE(\"1:30:00.0035\")", 0.062500046296296297},
      {"=VALUE(\"23:59:59.9995\")", 1.0},
      {"=VALUE(\"1:99\")", 0.11041666666666666},
      {"=VALUE(\"9999:59\")", 416.66597222222219},
      {"=VALUE(\"0100:0\")", 4.166666666666667},
      {"=VALUE(\"1:30-\")", 0.0625},
      {"=VALUE(\"1:30 PM -\")", 0.5625},
      {"=VALUE(\"1:30:00 .5\")", 0.062505787037037033},
      {"=VALUE(\"1:30.5:00\")", 0.062563657407407408},
      {"=VALUE(\"1:30.5 PM\")", 0.50104745370370374},
  };
  ExpectJaNumbers(cases, sizeof(cases) / sizeof(cases[0]));
  const char* const rejected[] = {
      "=VALUE(\"99:99\")", "=VALUE(\"001:30\")",     "=VALUE(\"1:000\")",     "=VALUE(\"10000:0\")",
      "=VALUE(\"24:60\")", "=VALUE(\"12:60 PM\")",   "=VALUE(\"1:30 P.M.\")", "=VALUE(\"1:30--\")",
      "=VALUE(\"1 P.-\")", "=VALUE(\"1:30:00.5.\")", "=VALUE(\"1:30PM\")",    "=VALUE(\"1:30,5\")",
  };
  ExpectJaValueErrors(rejected, sizeof(rejected) / sizeof(rejected[0]));
}

TEST(DateTextJaJp, DateAndTimeComposition) {
  const JaNumberCase cases[] = {
      {"=VALUE(\"Jan 1 1:30 PM\")", 46023.5625},           {"=VALUE(\"1:30 Aug 1\")", 46235.0625},
      {"=VALUE(\"1 Jan 1:30 2025\")", 46023.0625},         {"=VALUE(\"2026/1/1 1:30 1/2\")", 46023.0625},
      {"=VALUE(\"2026/1/1 1:30-1-\")", 46023.0625},        {"=VALUE(\"2026/1/1 1:30 1-1-1-1\")", 46023.0625},
      {"=VALUE(\"1:30 - 2026/1/1\")", 46023.0625},         {"=VALUE(\"2026/1/1 1:30:00.9999\")", 46023.062511574077},
      {"=VALUE(\"9999/12/31 23:59:59.9995\")", 2958466.0},
  };
  ExpectJaNumbers(cases, sizeof(cases) / sizeof(cases[0]));
  const char* const rejected[] = {
      "=VALUE(\"2026/1/1 1:30 12345\")",
      "=VALUE(\"2026/1/1 1:30 1 2 3 4 5\")",
      "=VALUE(\"1:30 2026/1/1 1\")",
      "=VALUE(\"2026/1/1 1\")",
      "=VALUE(\"1:30 1\")",
      "=VALUE(\"2026/1/1 1:30 1 PM\")",
  };
  ExpectJaValueErrors(rejected, sizeof(rejected) / sizeof(rejected[0]));
}

TEST(DateTextJaJp, PreProcessing) {
  const JaNumberCase cases[] = {
      {"=VALUE(\"\xEF\xBC\x92\xEF\xBC\x90\xEF\xBC\x92\xEF\xBC\x96\xEF\xBC\x8F\xEF\xBC\x91\xEF\xBC\x8F\xEF\xBC\x91"
       "\xE3\x80\x80\xEF\xBC\x91\xEF\xBC\x9A\xEF\xBC\x93\xEF\xBC\x90\")",
       46023.0625},
      {"=VALUE(\"\xEF\xBC\xAA\xEF\xBD\x81\xEF\xBD\x8E 1, 2026\")", 46023.0},
      {"=VALUE(\"1/2\xE3\x80\x80\")", 46024.0},
      {"=VALUE(\" 1/2\")", 46024.0},
      {"=VALUE(\"1:30\"&REPT(\" \",251))", 0.0625},
  };
  ExpectJaNumbers(cases, sizeof(cases) / sizeof(cases[0]));
  const char* const rejected[] = {
      "=VALUE(\"\u30001/2\")",
      "=VALUE(CHAR(9)&\"1/2\")",
      "=VALUE(\"1/2\"&UNICHAR(160))",
      "=VALUE(\"1:30\"&REPT(\" \",252))",
  };
  ExpectJaValueErrors(rejected, sizeof(rejected) / sizeof(rejected[0]));
}

TEST(DateTextJaJp, DatevalueAndTimevalueSplitOneValue) {
  const JaNumberCase cases[] = {
      {"=DATEVALUE(\"1:30 PM\")", 0.0},
      {"=DATEVALUE(\"9999:59\")", 416.0},
      {"=DATEVALUE(\"2026/1/1 23:59:59.9999\")", 46024.0},
      {"=TIMEVALUE(\"2026/1/1\")", 0.0},
      {"=TIMEVALUE(\"2026/1/1 23:59:59.9999\")", 0.0},
      {"=TIMEVALUE(\"25:00\")", 1.0416666666666667 - 1.0},
  };
  ExpectJaNumbers(cases, sizeof(cases) / sizeof(cases[0]));
  const char* const rejected[] = {"=DATEVALUE(\"46023\")", "=TIMEVALUE(\"0.5\")", "=DATEVALUE(46023.5)",
                                  "=DATEVALUE(\"1 1/2\")"};
  ExpectJaValueErrors(rejected, sizeof(rejected) / sizeof(rejected[0]));
}

TEST(DateTextJaJp, ImplicitCoercionKeepsALeadingBlank) {
  // Coercion reads the host clock's year, so the year-less forms are
  // compared against DATE(YEAR(TODAY()), ...) on the same clock.
  const Value pair = EvalSourceWithProfile("=\"13/1\"+0-DATE(YEAR(TODAY()),1,13)", mac_365_ja_jp_profile());
  ASSERT_TRUE(pair.is_number());
  EXPECT_EQ(pair.as_number(), 0.0);
  const Value trailing = EvalSourceWithProfile("=\"1/2 \"+0-DATE(YEAR(TODAY()),1,2)", mac_365_ja_jp_profile());
  ASSERT_TRUE(trailing.is_number());
  EXPECT_EQ(trailing.as_number(), 0.0);
  const Value leading = EvalSourceWithProfile("=\" 1/2\"+0", mac_365_ja_jp_profile());
  ASSERT_TRUE(leading.is_error());
  EXPECT_EQ(leading.as_error(), ErrorCode::Value);
}

TEST(DateTextJaJp, MixedFractions) {
  const JaNumberCase cases[] = {
      {"=VALUE(\"1 1/2\")", 1.5},
      {"=VALUE(\"2026 1/1\")", 2027.0},
      {"=VALUE(\"-1 1/2\")", -1.5},
      {"=VALUE(\"(1 1/2)%\")", -0.015},
      {"=VALUE(\"1 1/2%\")", 0.015},
      {"=VALUE(\"1,000 1/2\")", 1000.5},
      {"=VALUE(\"1 12345/1\")", 12346.0},
      {"=VALUE(\"1 32767/32767\")", 2.0},
      {"=VALUE(\"1 00001/2\")", 1.5},
      {"=VALUE(\"\xEF\xBC\x91\xE3\x80\x80\xEF\xBC\x91\xEF\xBC\x8F\xEF\xBC\x92\")", 1.5},
      {"=\"1 1/2\"+0", 1.5},
  };
  ExpectJaNumbers(cases, sizeof(cases) / sizeof(cases[0]));
  const char* const rejected[] = {"=VALUE(\"1 1/0\")",      "=VALUE(\"1  1/2\")",    "=VALUE(\"1.5 1/2\")",
                                  "=VALUE(\"1 1/2 1:30\")", "=VALUE(\"1 1 /2\")",    "=VALUE(\"1 1/2e1\")",
                                  "=VALUE(\"1 32768/1\")",  "=VALUE(\"1 1/32768\")", "=VALUE(\"1 000001/2\")"};
  ExpectJaValueErrors(rejected, sizeof(rejected) / sizeof(rejected[0]));
}

}  // namespace
}  // namespace eval
}  // namespace formulon
