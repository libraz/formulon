// Tests for the locale-aware number, boolean and error text helpers.

#include "eval/locale_text.h"

#include "eval/date_text_parse.h"
#include "eval/eval_profile_scope.h"
#include "eval/logical_coerce.h"
#include "eval/number_parse.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "value.h"

namespace formulon::eval {
namespace {

TEST(LocaleText, EnglishAndJapaneseKeepTheInvariantSpelling) {
  for (const ExcelProfile profile :
       {mac_365_en_us_profile(), win_365_en_us_profile(), mac_365_ja_jp_profile(), win_365_ja_jp_profile()}) {
    SCOPED_TRACE(excel_profile_id(profile));
    const ScopedEvalProfile scope(profile);
    EXPECT_EQ(locale_number_text(1.5E-20), "1.5E-20");
    EXPECT_EQ(locale_number_text(-1234.5), "-1234.5");
    EXPECT_EQ(locale_bool_text(true), "TRUE");
    EXPECT_EQ(locale_bool_text(false), "FALSE");
    EXPECT_EQ(locale_error_text(ErrorCode::NA), "#N/A");
    EXPECT_EQ(locale_error_text(ErrorCode::Div0), "#DIV/0!");
  }
}

TEST(LocaleText, ExponentOf99OrMoreKeeps14Digits) {
  const ScopedEvalProfile scope(mac_365_ja_jp_profile());
  EXPECT_EQ(locale_number_text(1.23456789012345E+98), "1.23456789012345E+98");
  EXPECT_EQ(locale_number_text(1.23456789012345E+99), "1.2345678901235E+99");
  EXPECT_EQ(locale_number_text(1.23456789012345E-99), "1.2345678901235E-99");
  EXPECT_EQ(locale_number_text(9.99999999999999E+99), "1E+100");
}

TEST(LocaleText, GermanUsesItsSeparatorAndNames) {
  const ScopedEvalProfile scope(ExcelProfile{ExcelHost::kMac365, ExcelLocale::kDeDE});
  // locale_tokens.general_tiny_number_text, bool_text_true / bool_text_false
  EXPECT_EQ(locale_number_text(1.5E-20), "1,5E-20");
  EXPECT_EQ(locale_bool_text(true), "WAHR");
  EXPECT_EQ(locale_bool_text(false), "FALSCH");
  // arraytotext.arraytotext_only_error_cells_default
  EXPECT_EQ(locale_error_text(ErrorCode::NA), "#NV");
  EXPECT_EQ(locale_error_text(ErrorCode::Div0), "#DIV/0!");
}

ExcelProfile mac_profile(ExcelLocale locale) {
  return ExcelProfile{ExcelHost::kMac365, locale};
}

bool parse_date(std::string_view text, double* serial, int current_year = 0) {
  date_parse::DateTimeText parsed{};
  if (!date_parse::parse_date_time_text(text, date_parse::TextEdges::kTrim, current_year, &parsed)) {
    return false;
  }
  *serial = parsed.value;
  return true;
}

TEST(LocaleNumberInput, DecimalAndGroupFollowTheProfile) {
  double v = 0.0;
  {
    const ScopedEvalProfile scope(mac_profile(ExcelLocale::kDeDE));
    // value_numbervalue.value_thousands, value_coercion_probes.value_comma_decimal
    ASSERT_TRUE(parse_excel_number("1,234", &v));
    EXPECT_EQ(v, 1.234);
    ASSERT_TRUE(parse_excel_number("1.234,5", &v));
    EXPECT_EQ(v, 1234.5);
    EXPECT_FALSE(parse_excel_number("1.5e2", &v));
    EXPECT_FALSE(parse_excel_number("$100", &v));
    // locale_tokens.value_euro_suffix_de
    ASSERT_TRUE(parse_excel_number("1.234,50 \xE2\x82\xAC", &v));
    EXPECT_EQ(v, 1234.5);
  }
  {
    const ScopedEvalProfile scope(mac_profile(ExcelLocale::kFrFR));
    // locale_tokens.value_euro_suffix_fr
    ASSERT_TRUE(parse_excel_number("1 234,50 \xE2\x82\xAC", &v));
    EXPECT_EQ(v, 1234.5);
    EXPECT_FALSE(parse_excel_number("3.14", &v));
  }
  {
    const ScopedEvalProfile scope(mac_profile(ExcelLocale::kThTH));
    // locale_tokens.value_baht_prefix, value_dollar_prefix
    ASSERT_TRUE(
        parse_excel_number("\xE0\xB8\xBF"
                           "1,234",
                           &v));
    EXPECT_EQ(v, 1234.0);
    EXPECT_FALSE(parse_excel_number("$1234", &v));
  }
}

TEST(LocaleDateInput, DayFirstDottedAndMonthNames) {
  double serial = 0.0;
  {
    const ScopedEvalProfile scope(mac_profile(ExcelLocale::kDeDE));
    // value_coercion_probes.value_slash_date_short, locale_tokens.value_dotted_dmy
    ASSERT_TRUE(parse_date("2/3/2023", &serial));
    EXPECT_EQ(serial, 44987.0);
    ASSERT_TRUE(parse_date("15.3.2024", &serial));
    EXPECT_EQ(serial, 45366.0);
    // locale_tokens.value_month_name_de, value_month_name_en
    ASSERT_TRUE(parse_date("15. M\xC3\xA4rz 2024", &serial));
    EXPECT_EQ(serial, 45366.0);
    EXPECT_FALSE(parse_date("15 March 2024", &serial));
    // value_coercion_probes.value_d_mmm_yy
    ASSERT_TRUE(parse_date("1-Jan-24", &serial));
    EXPECT_EQ(serial, 45292.0);
    // value_numbervalue.value_decimal reads month.year
    ASSERT_TRUE(parse_date("3.14", &serial));
    EXPECT_EQ(serial, 41699.0);
  }
  {
    const ScopedEvalProfile scope(mac_profile(ExcelLocale::kFrFR));
    ASSERT_TRUE(parse_date("15 mars 2024", &serial));
    EXPECT_EQ(serial, 45366.0);
    EXPECT_FALSE(parse_date("2024.3.15", &serial));
  }
  {
    const ScopedEvalProfile scope(mac_profile(ExcelLocale::kKoKR));
    // locale_tokens.value_dotted_ymd, value_numbervalue.value_multiple_decimals_is_value
    ASSERT_TRUE(parse_date("2024.3.15", &serial));
    EXPECT_EQ(serial, 45366.0);
    ASSERT_TRUE(parse_date("1.2.3", &serial));
    EXPECT_EQ(serial, 36925.0);
  }
  {
    const ScopedEvalProfile scope(mac_profile(ExcelLocale::kThTH));
    // locale_tokens.value_thai_slash_dmy; value_english_month_abbrev
    ASSERT_TRUE(parse_date("15/3/2567", &serial));
    EXPECT_EQ(serial, 243692.0);
    EXPECT_FALSE(parse_date("Jan 1, 2024", &serial));
    ASSERT_TRUE(parse_date("15 March 2024", &serial));
    EXPECT_EQ(serial, 45366.0);
  }
}

constexpr int kYear = 2026;

bool rejects(std::string_view text) {
  double serial = 0.0;
  return !parse_date(text, &serial, kYear);
}

double date_of(std::string_view text) {
  double serial = -1.0;
  EXPECT_TRUE(parse_date(text, &serial, kYear)) << text;
  return serial;
}

TEST(LocaleDateInput, MonthFirstOrder) {
  const ScopedEvalProfile scope(mac_profile(ExcelLocale::kEnUS));
  // A pair is m/d in the current year, else m/y; never d/m or y/m.
  EXPECT_EQ(date_of("1/2"), 46024.0);
  EXPECT_EQ(date_of("2/30"), 10990.0);
  EXPECT_EQ(date_of("1/0"), 36526.0);
  EXPECT_EQ(date_of("1/2026"), 46023.0);
  EXPECT_TRUE(rejects("13/1"));
  EXPECT_TRUE(rejects("2026/1"));
  // Triples follow the order with no fallback; a four-digit lead is the year.
  EXPECT_EQ(date_of("1/2/3"), 37623.0);
  EXPECT_EQ(date_of("2026/1/1"), 46023.0);
  EXPECT_TRUE(rejects("13/2/2026"));
  EXPECT_TRUE(rejects("2026/1/001"));
  // Month names: `MMM D` falls back to a year; `D MMM`; `MMM D, Y`.
  EXPECT_EQ(date_of("Jan 32"), 11689.0);
  EXPECT_EQ(date_of("1 Jan"), 46023.0);
  EXPECT_EQ(date_of("Sept 1, 2026"), 46266.0);
  // Full-width letters fold, full-width digits do not.
  EXPECT_EQ(date_of("\xEF\xBC\xAA\xEF\xBD\x81\xEF\xBD\x8E 1"), 46023.0);
  EXPECT_TRUE(rejects("\xEF\xBC\x91/2"));
  // U+3000 is a blank only after a time and in a long month-name date.
  EXPECT_EQ(date_of("1:30\xE3\x80\x80"), 0.0625);
  EXPECT_EQ(date_of("Jan 1\xE3\x80\x80"
                    "2026"),
            46023.0);
  EXPECT_TRUE(rejects("1/2\xE3\x80\x80"));
  EXPECT_TRUE(
      rejects("1/2\xE3\x80\x80"
              "1:30"));
  EXPECT_TRUE(
      rejects("1 Jan\xE3\x80\x80"
              "2026"));
}

TEST(LocaleDateInput, DayFirstOrder) {
  const ScopedEvalProfile scope(mac_profile(ExcelLocale::kFrFR));
  // A pair is d/m in the current year, else m/y with a year of at least 1.
  EXPECT_EQ(date_of("1/13"), 41275.0);
  EXPECT_EQ(date_of("2/29"), 47150.0);
  EXPECT_TRUE(rejects("1/0"));
  EXPECT_EQ(date_of("1/2/3"), 37653.0);
  // `MMM N` is month and year; `D MMM, Y`; local name prefixes only.
  EXPECT_EQ(date_of("Jan 1"), 36892.0);
  EXPECT_TRUE(rejects("Jan 1, 2026"));
  EXPECT_EQ(date_of("1 Jan, 2026"), 46023.0);
  EXPECT_EQ(date_of("1\xE3\x80\x80Jan\xE3\x80\x80"
                    "2026"),
            46023.0);
  EXPECT_TRUE(rejects("Feb 1"));
  // Extra whole-token spellings on top of the prefix rule; ambiguous prefixes take the first month.
  EXPECT_EQ(date_of("Jun 1"), 37043.0);
  EXPECT_EQ(date_of("1 Jul 2026"), 46204.0);
  EXPECT_EQ(date_of("Jui 1"), 37043.0);
  EXPECT_TRUE(rejects("Jn 1"));
  EXPECT_TRUE(rejects("Fev 1"));
  // `,` closes a time, `.` does not (and is no fraction).
  EXPECT_EQ(date_of("1:30 PM,"), 0.5625);
  EXPECT_TRUE(rejects("1:30."));
  EXPECT_TRUE(rejects("2026/1/1 1:30:00.5"));
  EXPECT_NEAR(date_of("1:30,5"), 90.5 / 86400.0, 1e-15);
}

TEST(LocaleDateInput, ExtraMonthSpellingsAreLocaleData) {
  const ScopedEvalProfile scope(mac_profile(ExcelLocale::kItIT));
  EXPECT_TRUE(rejects("Jun 1"));
  EXPECT_TRUE(rejects("Jul 1"));
  EXPECT_EQ(date_of("Nov 1"), 37196.0);
}

TEST(LocaleDateInput, DottedDayFirstOrder) {
  const ScopedEvalProfile scope(mac_profile(ExcelLocale::kDeDE));
  EXPECT_EQ(date_of("1.13"), 41275.0);
  EXPECT_TRUE(rejects("2026.1"));
  EXPECT_EQ(date_of("1.2.2026"), 46054.0);
  EXPECT_EQ(date_of("Janu 1"), 36892.0);
  EXPECT_TRUE(rejects("May 1"));
  // `.` closes a time, joins it to a following date and separates the tail.
  EXPECT_EQ(date_of("1:30.5.5"), 46147.0625);
  EXPECT_EQ(date_of("2026/1/1 1:30 1.5"), 46023.0625);
  EXPECT_EQ(date_of("2026/1/1 1:30:00.9999"), 46023.0625);
  EXPECT_TRUE(rejects("2026/1/1 1:30:00.0004"));
  EXPECT_TRUE(rejects("1:30:00.5"));
}

TEST(LocaleDateInput, YearFirstOrder) {
  {
    const ScopedEvalProfile scope(mac_profile(ExcelLocale::kKoKR));
    EXPECT_EQ(date_of("Jan.1"), 46023.0);
    EXPECT_EQ(date_of("Jan. 1, 2026"), 46023.0);
    EXPECT_TRUE(rejects("1:30:00.5"));
  }
  {
    const ScopedEvalProfile scope(mac_profile(ExcelLocale::kZhTW));
    EXPECT_EQ(date_of("/1/2"), 46024.0);
    EXPECT_EQ(date_of("2026\xE5\xB9\xB4"
                      "1\xE6\x9C\x88"
                      "1\xE6\x97\xA5"),
              46023.0);
    EXPECT_TRUE(
        rejects("1\xE6\x9C\x88"
                "1\xE6\x97\xA5"));
  }
  {
    const ScopedEvalProfile scope(mac_365_ja_jp_profile());
    // A time's trailing separator joins a following date without a blank.
    EXPECT_EQ(date_of("1:30-5-5"), 46147.0625);
    // Tail runs follow the time-field digit rule.
    EXPECT_EQ(date_of("2026/1/1 1:30 9999"), 46023.0625);
    EXPECT_TRUE(rejects("2026/1/1 1:30 0004"));
    // Ethiopic digits are numbers, never date fields.
    EXPECT_TRUE(rejects("\xE1\x8D\xA9/\xE1\x8D\xAA"));
  }
}

TEST(LocaleDateInput, MarkerSpellings) {
  {
    const ScopedEvalProfile scope(mac_profile(ExcelLocale::kThTH));
    EXPECT_EQ(date_of("Sep 1"), 37135.0);
    EXPECT_TRUE(rejects("Sept 1"));
    EXPECT_TRUE(rejects("1:30 PM."));
    EXPECT_EQ(date_of("1:30 PM ."), 0.5625);
  }
  {
    const ScopedEvalProfile scope(mac_profile(ExcelLocale::kEsES));
    EXPECT_EQ(date_of("1 P. M."), 13.0 / 24.0);
  }
}

TEST(LocaleDateInput, ExtraMonthSpellingInGerman) {
  const ScopedEvalProfile scope(mac_profile(ExcelLocale::kDeDE));
  EXPECT_EQ(date_of("2 Mrz 2026"), 46083.0);
  EXPECT_EQ(date_of("Mrz 2"), 37316.0);
}

TEST(LocaleDateInput, HangulDateAndTime) {
  {
    const ScopedEvalProfile scope(mac_profile(ExcelLocale::kKoKR));
    EXPECT_EQ(date_of("2026\xEB\x85\x84 1\xEC\x9B\x94 2\xEC\x9D\xBC"), 46024.0);
    // A year below 100 lands in 4263..4362; 100..1899 is no year.
    EXPECT_EQ(date_of("26\xEB\x85\x84 1\xEC\x9B\x94 2\xEC\x9D\xBC"), 886081.0);
    EXPECT_EQ(date_of("68\xEB\x85\x84 2\xEC\x9B\x94 29\xEC\x9D\xBC"), 864955.0);
    EXPECT_EQ(date_of("0059\xEB\x85\x84 1\xEC\x9B\x94 1\xEC\x9D\xBC"), 898133.0);
    EXPECT_TRUE(rejects("150\xEB\x85\x84 1\xEC\x9B\x94 1\xEC\x9D\xBC"));
    // Exactly one blank between the parts, and no time after it.
    EXPECT_TRUE(
        rejects("2026\xEB\x85\x84"
                "1\xEC\x9B\x94 2\xEC\x9D\xBC"));
    EXPECT_TRUE(rejects("2026\xEB\x85\x84  1\xEC\x9B\x94 2\xEC\x9D\xBC"));
    EXPECT_TRUE(rejects("2026\xEB\x85\x84 1\xEC\x9B\x94 2\xEC\x9D\xBC 1:30"));
    EXPECT_EQ(date_of("1\xEC\x8B\x9C 30\xEB\xB6\x84"), 0.0625);
    EXPECT_EQ(date_of("1\xEC\x8B\x9C 30\xEB\xB6\x84 5\xEC\xB4\x88"), (5400.0 + 5.0) / 86400.0);
    EXPECT_TRUE(
        rejects("1\xEC\x8B\x9C"
                "30\xEB\xB6\x84"));
    EXPECT_TRUE(rejects("1\xEC\x8B\x9C"));
    EXPECT_TRUE(rejects("1:30 \xEC\x98\xA4\xED\x9B\x84"));
  }
  {
    const ScopedEvalProfile scope(mac_365_ja_jp_profile());
    EXPECT_TRUE(rejects("1\xEC\x8B\x9C 30\xEB\xB6\x84"));
  }
}

TEST(LocaleDateInput, ChineseMeridiemMarkers) {
  for (const ExcelLocale locale : {ExcelLocale::kZhCN, ExcelLocale::kZhTW}) {
    const ScopedEvalProfile scope(mac_profile(locale));
    EXPECT_EQ(date_of("1:30 \xE4\xB8\x8B\xE5\x8D\x88"), 0.5625);
    EXPECT_EQ(date_of("12:00 \xE4\xB8\x8A\xE5\x8D\x88"), 0.0);
    EXPECT_EQ(date_of("1:30 \xE4\xB8\x8B\xE5\x8D\x88."), 0.5625);
    EXPECT_EQ(date_of("1:30 \xE4\xB8\x8B\xE5\x8D\x88 2026/1/2"), 46024.5625);
    EXPECT_EQ(date_of("2026/1/2 1:30 \xE4\xB8\x8B\xE5\x8D\x88"), 46024.5625);
    EXPECT_TRUE(rejects("1:30\xE4\xB8\x8B\xE5\x8D\x88"));
    EXPECT_TRUE(rejects("13:30 \xE4\xB8\x8B\xE5\x8D\x88"));
    EXPECT_TRUE(rejects("\xE4\xB8\x8B\xE5\x8D\x88 1:30"));
  }
  const ScopedEvalProfile scope(mac_365_ja_jp_profile());
  EXPECT_TRUE(rejects("1:30 \xE4\xB8\x8B\xE5\x8D\x88"));
}

TEST(LocaleNumberInput, CurrencySymbolsAreLocaleData) {
  double v = 0.0;
  {
    const ScopedEvalProfile scope(mac_365_ja_jp_profile());
    ASSERT_TRUE(parse_excel_number("US$1,000", &v));
    EXPECT_EQ(v, 1000.0);
  }
  {
    const ScopedEvalProfile scope(mac_profile(ExcelLocale::kKoKR));
    ASSERT_TRUE(
        parse_excel_number("\xEF\xBF\xA6"
                           "1.5",
                           &v));
    EXPECT_EQ(v, 1.5);
  }
  {
    const ScopedEvalProfile scope(mac_profile(ExcelLocale::kZhTW));
    ASSERT_TRUE(parse_excel_number("NT$-1", &v));
    EXPECT_EQ(v, -1.0);
  }
  {
    const ScopedEvalProfile scope(mac_profile(ExcelLocale::kPtBR));
    ASSERT_TRUE(parse_excel_number("(R$1)", &v));
    EXPECT_EQ(v, -1.0);
    ASSERT_TRUE(parse_excel_number("R$1,5", &v));
    EXPECT_EQ(v, 1.5);
  }
  {
    const ScopedEvalProfile scope(mac_profile(ExcelLocale::kRuRU));
    ASSERT_TRUE(parse_excel_number("1,5 \xE2\x82\xBD", &v));
    EXPECT_EQ(v, 1.5);
    EXPECT_FALSE(
        parse_excel_number("\xE2\x82\xBD"
                           "1",
                           &v));
  }
  {
    const ScopedEvalProfile scope(mac_profile(ExcelLocale::kEnUS));
    EXPECT_FALSE(parse_excel_number("US$1", &v));
  }
}

TEST(LocaleNumberInput, SpaceGroupCapsTheMixedFractionNumerator) {
  double v = 0.0;
  const ScopedEvalProfile scope(mac_profile(ExcelLocale::kFrFR));
  ASSERT_TRUE(parse_excel_number("1 12/345", &v));
  EXPECT_DOUBLE_EQ(v, 1.0 + 12.0 / 345.0);
  EXPECT_FALSE(parse_excel_number("1 100/1", &v));
}

TEST(LocaleBooleanInput, OnlyTheLocaleNamesCoerce) {
  bool out = false;
  ErrorCode err = ErrorCode::Value;
  const ScopedEvalProfile scope(mac_profile(ExcelLocale::kDeDE));
  // text_to_bool_probes.text_bool_and_true_false: no logical value at all
  EXPECT_EQ(logical_coerce(Value::text("TRUE"), &out, &err), LogicalCoerce::Skip);
  EXPECT_EQ(logical_coerce(Value::text("WAHR"), &out, &err), LogicalCoerce::HasValue);
  EXPECT_TRUE(out);
  EXPECT_EQ(logical_coerce(Value::text("falsch"), &out, &err), LogicalCoerce::HasValue);
  EXPECT_FALSE(out);
  EXPECT_EQ(logical_coerce(Value::text("1"), &out, &err), LogicalCoerce::Error);
}

}  // namespace
}  // namespace formulon::eval
