//
// Unit tests for the date/time side of the TEXT() format-string engine.
// Serials used below match the Excel 1900 epoch (see `eval/date_time.cpp`):
//
//   2024-01-01 = 45292
//   2024-03-15 = 45366
//   2024-07-04 = 45477
//   2026-04-23 = 46135
//
// Time-of-day serials are expressed as `serial + fractional_day`. The
// fractional helpers in date_time.cpp are exercised transitively here.

#include <string>
#include <string_view>
#include <utility>

#include "eval/text_format/display_text.h"
#include "eval/text_format/number_format.h"
#include "gtest/gtest.h"
#include "util/test_eval_helpers.h"

namespace formulon {
namespace text_format {
namespace {

std::string Render(double value, std::string_view format) {
  std::string out;
  const FormatStatus s = apply_format(value, format, out);
  EXPECT_EQ(s, FormatStatus::kOk);
  const auto display = format_value_for_display(Value::number(value), format, false);
  EXPECT_EQ(display.status, DisplayStatus::kOk);
  EXPECT_EQ(display.text, out);
  return out;
}

// ---------------------------------------------------------------------------
// Year / month / day
// ---------------------------------------------------------------------------

TEST(DateFormatYear, FourDigit) {
  EXPECT_EQ(Render(45366.0, "yyyy"), "2024");
}

TEST(DateFormatYear, TwoDigit) {
  EXPECT_EQ(Render(45366.0, "yy"), "24");
}

TEST(DateFormatMonth, NumericOne) {
  EXPECT_EQ(Render(45292.0, "m"), "1");
}

TEST(DateFormatMonth, NumericTwoDigit) {
  EXPECT_EQ(Render(45292.0, "mm"), "01");
}

TEST(DateFormatMonth, EnglishShort) {
  // Mac Excel ja-JP surprisingly renders `mmm` as English 3-letter names
  // (Jan / Feb / ...). The Japanese `N月` form is reserved for `[DBNum2]`
  // and friends, which we do not implement.
  EXPECT_EQ(Render(45366.0, "mmm"), "Mar");
}

TEST(DateFormatMonth, EnglishLong) {
  // `mmmm` in Mac Excel ja-JP renders the English full month name.
  EXPECT_EQ(Render(45292.0, "mmmm"), "January");
}

TEST(DateFormatDay, NumericOne) {
  EXPECT_EQ(Render(45292.0, "d"), "1");
}

TEST(DateFormatDay, TwoDigit) {
  EXPECT_EQ(Render(45292.0, "dd"), "01");
}

TEST(DateFormatDay, WeekdayShort) {
  // 2024-01-01 is a Monday. Mac Excel ja-JP renders `ddd` as `Mon` etc.
  EXPECT_EQ(Render(45292.0, "ddd"), "Mon");
}

TEST(DateFormatDay, WeekdayLong) {
  // `dddd` renders the English full weekday name in ja-JP.
  EXPECT_EQ(Render(45292.0, "dddd"), "Monday");
}

// ---------------------------------------------------------------------------
// Composite date formats
// ---------------------------------------------------------------------------

TEST(DateFormatComposite, Iso) {
  EXPECT_EQ(Render(45366.0, "yyyy-mm-dd"), "2024-03-15");
}

TEST(DateFormatComposite, Slash) {
  EXPECT_EQ(Render(45366.0, "yyyy/m/d"), "2024/3/15");
}

TEST(DateFormatComposite, Kanji) {
  EXPECT_EQ(Render(45366.0, "yyyy年m月d日"), "2024年3月15日");
}

// ---------------------------------------------------------------------------
// Hours / minutes / seconds
// ---------------------------------------------------------------------------

TEST(DateFormatTime, TwentyFourHour) {
  // 13:30:00 -> fractional 0.5625
  EXPECT_EQ(Render(0.5625, "h:mm"), "13:30");
}

TEST(DateFormatTime, TwoDigitHour) {
  EXPECT_EQ(Render(0.5625, "hh:mm"), "13:30");
}

TEST(DateFormatTime, HourMinuteSecond) {
  // 13:30:45 -> fraction = (13*3600+30*60+45)/86400
  const double frac = (13.0 * 3600.0 + 30.0 * 60.0 + 45.0) / 86400.0;
  EXPECT_EQ(Render(frac, "h:mm:ss"), "13:30:45");
}

TEST(DateFormatTime, AmPmMorning) {
  const double frac = 9.5 / 24.0;  // 09:30:00
  EXPECT_EQ(Render(frac, "h:mm AM/PM"), "9:30 AM");
}

TEST(DateFormatTime, AmPmAfternoon) {
  const double frac = 13.5 / 24.0;  // 13:30:00
  EXPECT_EQ(Render(frac, "h:mm AM/PM"), "1:30 PM");
}

TEST(DateFormatTime, AShortMarker) {
  const double frac = 9.5 / 24.0;
  EXPECT_EQ(Render(frac, "h:mm A/P"), "9:30 A");
}

// Mac Excel 16.111.3 (ja-JP) accepts full-width date/time syntax while
// keeping quoted/escaped text opaque. The implementation normalises the
// syntax before tokenisation so date disambiguation remains unchanged.
TEST(DateFormatJaJpFullWidth, DateAndTimeRuns) {
  EXPECT_EQ(Render(45366.0, "ｙｙｙｙ－ｍｍ－ｄｄ"), "2024-03-15");
  const double frac = (13.0 * 3600.0 + 30.0 * 60.0 + 45.0) / 86400.0;
  EXPECT_EQ(Render(frac, "ｈ：ｍｍ：ｓｓ"), "13:30:45");
  EXPECT_EQ(Render(45366.0, "ｙｙｙｙ‘ｍｍ’ｄｄ"), "2024‘03’15");
}

TEST(DateFormatJaJpFullWidth, AmPmAndWeekdayMarkers) {
  const double morning = 9.5 / 24.0;
  const double afternoon = 13.5 / 24.0;
  EXPECT_EQ(Render(morning, "ｈ：ｍｍ ＡＭ／ＰＭ"), "9:30 AM");
  EXPECT_EQ(Render(afternoon, "ｈ：ｍｍ Ａ／Ｐ"), "1:30 P");
  EXPECT_EQ(Render(45292.0, "ｙｙｙｙ－ｍｍ－ｄｄ ａａａ"), "2024-01-01 月");
}

// ---------------------------------------------------------------------------
// Minute disambiguation
// ---------------------------------------------------------------------------

TEST(DateFormatMinuteDisambiguation, AfterHour) {
  const double frac = 13.5 / 24.0;  // 13:30
  EXPECT_EQ(Render(frac, "h:m"), "13:30");
}

TEST(DateFormatMinuteDisambiguation, BeforeSecond) {
  const double frac = (13.0 * 3600.0 + 30.0 * 60.0 + 45.0) / 86400.0;
  EXPECT_EQ(Render(frac, "m:ss"), "30:45");
}

TEST(DateFormatMinuteDisambiguation, StandaloneIsMonth) {
  // Without adjacent h: or :s, `m` is the month.
  EXPECT_EQ(Render(45366.0, "m"), "3");
}

// ---------------------------------------------------------------------------
// Elapsed-time brackets
// ---------------------------------------------------------------------------

TEST(DateFormatElapsed, ElapsedHours) {
  // 1.5 days = 36 hours.
  EXPECT_EQ(Render(1.5, "[h]"), "36");
  // Repeated elapsed-hour tokens request a minimum field width.
  EXPECT_EQ(Render(1.0 / 24.0, "[hh]"), "01");
}

TEST(DateFormatElapsed, ElapsedMinutes) {
  EXPECT_EQ(Render(1.0 / 24.0, "[m]"), "60");
}

TEST(DateFormatElapsed, ElapsedSeconds) {
  EXPECT_EQ(Render(1.0 / 86400.0, "[s]"), "1");
}

// ---------------------------------------------------------------------------
// Fractional seconds
// ---------------------------------------------------------------------------

TEST(DateFormatFractionalSeconds, TwoDigits) {
  // 12:00:00.5 -> 0.5 + (12 hours in fraction).
  const double frac = 0.5 + 0.5 / 86400.0;  // +0.5s
  const std::string out = Render(frac, "h:mm:ss.00");
  EXPECT_EQ(out, "12:00:00.50");
}

TEST(DateFormatFractionalSeconds, DecimalDigitsSurviveRounding) {
  EXPECT_EQ(Render(0.3 / 86400.0, "ss.00"), "00.30");
  EXPECT_EQ(Render(0.29 / 86400.0, "ss.000"), "00.290");
  EXPECT_EQ(Render(0.58 / 86400.0, "ss.00"), "00.58");
  EXPECT_EQ(Render(0.5 + 0.3 / 86400.0, "hh:mm:ss.00"), "12:00:00.30");
}

TEST(DateFormatFractionalSeconds, WidePrecisionRemainsFiniteAndKeepsAllPlaceholders) {
  const std::string format = "ss." + std::string(400, '0');
  EXPECT_EQ(Render(0.5 / 86400.0, format), "00.5" + std::string(399, '0'));
}

TEST(DateFormatElapsed, WideFieldsKeepTheirPadding) {
  for (char token : {'h', 'm', 's'}) {
    const std::string format = "[" + std::string(256, token) + "]";
    EXPECT_EQ(Render(1.0 / 86400.0, format), std::string(255, '0') + (token == 's' ? "1" : "0"));
  }
}

TEST(DateFormatElapsed, RoundedComponentsCarryTogether) {
  EXPECT_EQ(Render(3599.6 / 86400.0, "[h]:mm:ss"), "1:00:00");
  EXPECT_EQ(Render(59.6 / 86400.0, "[m]:ss"), "1:00");
  EXPECT_EQ(Render(1.6 / 86400.0, "[s]"), "2");
  EXPECT_EQ(Render(1.0 + 86399.6 / 86400.0, "[h]:mm:ss"), "48:00:00");
  EXPECT_EQ(Render(59.996 / 86400.0, "[m]:ss.00"), "1:00.00");
  EXPECT_EQ(Render(59.994 / 86400.0, "[m]:ss.00"), "0:59.99");
}

TEST(DateFormatTime, TextBuiltinUsesTheSameRounding) {
  for (const auto& row :
       {std::pair{"=TEXT(0.3/86400,\"ss.00\")", "00.30"}, std::pair{"=TEXT(3599.6/86400,\"[h]:mm:ss\")", "1:00:00"},
        std::pair{"=TEXT(DATE(9999,12,31)+0.5,\"yyyy/mm/dd hh:mm:ss\")", "9999/12/31 12:00:00"}}) {
    const Value value = test::EvalSource(row.first);
    ASSERT_TRUE(value.is_text()) << row.first;
    EXPECT_EQ(value.as_text(), row.second);
  }
}

TEST(DateFormatTime, MidnightRoundingCarriesIntoCalendarFieldsInBothEpochs) {
  for (bool date1904 : {false, true}) {
    const double day = date1904 ? 43830.0 : 45292.0;
    for (const char* code : {"yyyy/mm/dd hh:mm:ss", "yyyy/mm/dd"}) {
      std::string out;
      EXPECT_EQ(apply_format(day + 86399.6 / 86400.0, code, out, date1904), FormatStatus::kOk);
      EXPECT_EQ(out, std::string("2024/01/02") + (std::string_view(code).size() > 10 ? " 00:00:00" : ""));
    }
  }
}

TEST(DateFormatRange, RoundedCalendarCannotAdvanceBeyondYear9999) {
  for (bool date1904 : {false, true}) {
    const double last_day = date1904 ? 2957003.0 : 2958465.0;
    for (const char* code : {"yyyy/mm/dd hh:mm:ss", "yyyy/mm/dd"}) {
      std::string out = "prefix";
      EXPECT_EQ(apply_format(last_day + 86399.6 / 86400.0, code, out, date1904), FormatStatus::kOverflow);
      EXPECT_EQ(out, "prefix");
      const auto display = format_value_for_display(Value::number(last_day + 86399.6 / 86400.0), code, date1904);
      EXPECT_EQ(display.status, DisplayStatus::kOverflow);
      if (date1904) {
        EXPECT_EQ(apply_format(-(last_day + 86399.6 / 86400.0), code, out, true), FormatStatus::kOverflow);
        EXPECT_EQ(out, "prefix");
      }
    }
  }
}

TEST(DateFormatRange, LastCalendarDayIncludesTimeInBothEpochs) {
  for (bool date1904 : {false, true}) {
    const double last_day = date1904 ? 2957003.0 : 2958465.0;
    std::string out;
    EXPECT_EQ(apply_format(last_day + 0.5, "yyyy/mm/dd hh:mm:ss", out, date1904), FormatStatus::kOk);
    EXPECT_EQ(out, "9999/12/31 12:00:00");
    const auto display = format_value_for_display(Value::number(last_day + 0.5), "yyyy/mm/dd hh:mm:ss", date1904);
    EXPECT_EQ(display.status, DisplayStatus::kOk);
    EXPECT_EQ(display.text, "9999/12/31 12:00:00");
    if (date1904) {
      out.clear();
      EXPECT_EQ(apply_format(-(last_day + 0.5), "yyyy/mm/dd hh:mm:ss", out, true), FormatStatus::kOk);
      EXPECT_EQ(out, "-9999/12/31 12:00:00");
      out.clear();
      EXPECT_EQ(apply_format(-(last_day + 1.0), "yyyy/mm/dd", out, true), FormatStatus::kOverflow);
    }
    out.clear();
    EXPECT_EQ(apply_format(last_day + 1.0, "yyyy/mm/dd", out, date1904), FormatStatus::kOverflow);
  }
}

TEST(DateFormatFractionalSeconds, OneDigit) {
  // 12:00:00.3 -> rendered at .0 precision rounds to .3.
  const double frac = 0.5 + 0.3 / 86400.0;
  EXPECT_EQ(Render(frac, "h:mm:ss.0"), "12:00:00.3");
}

// ---------------------------------------------------------------------------
// Single-letter month (`mmmmm`, run length >= 5)
// ---------------------------------------------------------------------------

TEST(DateFormatMonth, SingleLetterSeptember) {
  // 2024-09-15: the 259th day of 2024 -> serial 45292 + 258 = 45550.
  EXPECT_EQ(Render(45550.0, "mmmmm"), "S");
}

TEST(DateFormatMonth, SingleLetterJanuary) {
  // 2024-01-01 -> serial 45292.
  EXPECT_EQ(Render(45292.0, "mmmmm"), "J");
}

TEST(DateFormatMonth, SingleLetterLongRunStillFirstLetter) {
  // Runs longer than 5 `m` characters are also first-letter.
  EXPECT_EQ(Render(45292.0, "mmmmmm"), "J");
}

TEST(DateFormatMonth, FullAndSingleLetterComposite) {
  // `ddddd/mmmmm/yyyyy` on a Saturday in 2012.
  // 2012-03-31 is a Saturday. Serial = 40999.
  EXPECT_EQ(Render(40999.0, "ddddd/mmmmm/yyyyy"), "Saturday/M/2012");
}

// ---------------------------------------------------------------------------
// Mixed date + number-digit tokens must surface `#VALUE!`.
// ---------------------------------------------------------------------------

TEST(DateFormatMixedRejection, MonthAndOptionalDigits) {
  std::string out;
  const FormatStatus s = apply_format(45292.0, "mm###", out);
  EXPECT_EQ(s, FormatStatus::kValueError);
}

TEST(DateFormatMixedRejection, MonthAndZeroDigit) {
  std::string out;
  const FormatStatus s = apply_format(45292.0, "mm000", out);
  EXPECT_EQ(s, FormatStatus::kValueError);
}

// ---------------------------------------------------------------------------
// Calendar range (Excel Range.Text, ja-JP).
// ---------------------------------------------------------------------------

TEST(DateFormatRange, SerialZeroIsDayZeroOf1900) {
  EXPECT_EQ(Render(0.0, "yyyy/m/d"), "1900/1/0");
  EXPECT_EQ(Render(0.5, "yyyy/m/d"), "1900/1/0");
  EXPECT_EQ(Render(0.0, R"(ggge"年"m"月"d"日")"), "明治33年1月0日");
}

TEST(DateFormatRange, OutOfRangeSerialIsOverflow) {
  std::string out;
  EXPECT_EQ(apply_format(-1.0, "yyyy/m/d", out), FormatStatus::kOverflow);
  EXPECT_EQ(apply_format(2958466.0, "h:mm:ss", out), FormatStatus::kOverflow);
  EXPECT_EQ(apply_format(2958465.0, "yyyy/m/d", out, /*date1904=*/true), FormatStatus::kOverflow);
}

TEST(DateFormatRange, Negative1904SerialShowsALeadingMinus) {
  std::string out;
  EXPECT_EQ(apply_format(-1.0, "yyyy/m/d", out, /*date1904=*/true), FormatStatus::kOk);
  EXPECT_EQ(out, "-1904/1/2");
  out.clear();
  EXPECT_EQ(apply_format(-1234.5678, "[h]:mm", out, /*date1904=*/true), FormatStatus::kOk);
  EXPECT_EQ(out, "-29629:37");
}

}  // namespace
}  // namespace text_format
}  // namespace formulon
