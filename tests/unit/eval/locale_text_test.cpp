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
  double frac = 0.0;
  bool has_date = false;
  bool has_time = false;
  return date_parse::parse_date_time_text(text, serial, &frac, &has_date, &has_time, current_year);
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

TEST(LocaleNumberInput, LenientGroupsDropSeparatorsAnywhere) {
  double v = 0.0;
  // value_numbervalue.numbervalue_whitespace_trimmed (de)
  ASSERT_TRUE(parse_numeric("3.14", ',', '.', &v, /*strict_groups=*/false));
  EXPECT_EQ(v, 314.0);
  EXPECT_FALSE(parse_numeric("3.14", ',', '.', &v));
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
