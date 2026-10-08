#include "eval/text_format/format_localize.h"

#include <string>
#include <string_view>

#include "eval/eval_profile_scope.h"
#include "eval/text_format/number_format.h"
#include "excel_profile.h"
#include "gtest/gtest.h"

namespace formulon {
namespace text_format {
namespace {

using number_format_detail::localize_format;
using number_format_detail::LocalizedFormat;

constexpr ExcelProfile kDe{ExcelHost::kMac365, ExcelLocale::kDeDE};
constexpr ExcelProfile kFr{ExcelHost::kMac365, ExcelLocale::kFrFR};
constexpr ExcelProfile kZh{ExcelHost::kMac365, ExcelLocale::kZhCN};
constexpr ExcelProfile kTh{ExcelHost::kMac365, ExcelLocale::kThTH};

LocalizedFormat Localize(std::string_view fmt, ExcelProfile profile) {
  return localize_format(fmt, FormatDialect::kLocalized, profile);
}

struct Rendered {
  FormatStatus status;
  std::string text;
};

Rendered Render(double value, std::string_view fmt, ExcelProfile profile,
                FormatDialect dialect = FormatDialect::kLocalized) {
  const eval::ScopedEvalProfile scope(profile);
  Rendered r{FormatStatus::kOk, {}};
  r.status = apply_format(value, fmt, r.text, false, dialect);
  return r;
}

TEST(FormatLocalize, SeparatorsMapOntoTheInvariantSpelling) {
  EXPECT_EQ(Localize("#.##0,00", kDe).text, "#,##0.00");
  EXPECT_EQ(Localize("0.0E+00", kDe).text, "0,0E+00");
  EXPECT_EQ(Localize("# ##0,00", kFr).text, "#,##0.00");
  // fr-FR has no `.` separator, and a blank outside digit placeholders is text.
  EXPECT_EQ(Localize("0.00", kFr).text, "0\\.00");
  EXPECT_EQ(Localize("h:mm AM/PM", kFr).text, "h:mm AM/PM");
  // Quoted text and escape payloads are copied unchanged.
  EXPECT_EQ(Localize("0,0\"a,b.\"\\.", kDe).text, "0.0\"a,b.\"\\.");
  EXPECT_EQ(Localize("#,##0.00", mac_365_en_us_profile()).text, "#,##0.00");
}

TEST(FormatLocalize, ColourNamesFollowTheLocale) {
  EXPECT_EQ(Localize("0;[Rot]0", kDe).text, "0;[Red]0");    // locale_tokens.text_color_rot
  EXPECT_EQ(Localize("0;[Rouge]0", kFr).text, "0;[Red]0");  // locale_tokens.text_color_rouge
  EXPECT_EQ(Localize("[红色]0", kZh).text, "[Red]0");       // locale_tokens.text_color_hongse
  EXPECT_EQ(Localize("[赤]0", mac_365_ja_jp_profile()).text, "[Red]0");
  EXPECT_EQ(Localize("[色12]0", mac_365_ja_jp_profile()).text, "[Color12]0");
  EXPECT_FALSE(Localize("0;[Red]0", kDe).valid);  // locale_tokens.text_color_red
  EXPECT_FALSE(Localize("0;[Red]0", kTh).valid);
  EXPECT_FALSE(Localize("[Red]0", mac_365_ja_jp_profile()).valid);
  EXPECT_TRUE(Localize("[Red]0", mac_365_en_us_profile()).valid);
}

TEST(FormatLocalize, GeneralKeywordFollowsTheLocale) {
  EXPECT_EQ(Localize("Standard", kDe).text, "General");    // locale_tokens.text_general_standard
  EXPECT_EQ(Localize("G/通用格式", kZh).text, "General");  // locale_tokens.text_general_g_tongyong
  EXPECT_EQ(Localize("G/標準", mac_365_ja_jp_profile()).text, "General");
  EXPECT_FALSE(Localize("General", kDe).valid);  // locale_tokens.text_general_english
  EXPECT_FALSE(Localize("General", kZh).valid);
  EXPECT_TRUE(Localize("General", kTh).valid);
}

TEST(FormatLocalize, StoredCodesAreAlreadyInvariant) {
  const LocalizedFormat stored = localize_format("#,##0.00;[Red]General", FormatDialect::kStored, kDe);
  EXPECT_TRUE(stored.valid);
  EXPECT_EQ(stored.text, "#,##0.00;[Red]General");
}

TEST(FormatLocalizeRender, LocaleDateLetters) {
  // text_format.text_era_ggge_reiwa: de-DE `m` is a minute and `d` is text.
  EXPECT_EQ(Render(45383, "ggge年m月d日", kDe).text, "2024年0月d日");
  EXPECT_EQ(Render(45383, "ggge年m月d日", kFr).text, "2024年4月d日");
  EXPECT_EQ(Render(45366, "yyyy-mm-dd h:mm", kDe).text, "yyyy-00-dd 0:00");
  EXPECT_EQ(Render(45292, "TTTT", kDe).text, "Montag");   // locale_tokens.weekdays_utututut
  EXPECT_EQ(Render(45292, "JJJJ", kDe).text, "2024");     // locale_tokens.weekdays_ujujujuj
  EXPECT_EQ(Render(45292, "jjjj", kFr).text, "lundi");    // locale_tokens.weekdays_jjjj
  EXPECT_EQ(Render(45292, "mmmm", kFr).text, "janvier");  // locale_tokens.weekdays_mmmm
  EXPECT_EQ(Render(45292, "mmmm", kDe).text, "00");       // locale_tokens.weekdays_mmmm
  EXPECT_EQ(Render(45383, "aaaa", kFr).text, "2024");     // text_format.text_dow_aaaa
  EXPECT_EQ(Render(45383, "MMM", kDe).text, "Apr");       // locale_tokens.months_umumum
}

TEST(FormatLocalizeRender, SlashBetweenDateLettersIsRejected) {
  EXPECT_EQ(Render(45366, "yyyy/m/d", kDe).status, FormatStatus::kValueError);  // text_format.text_date_slash
  EXPECT_EQ(Render(45366, "yyyy/m/d", kFr).status, FormatStatus::kValueError);
  EXPECT_EQ(Render(45366, "yyyy/m/d", kZh).text, "2024/3/15");
  EXPECT_EQ(Render(0.75, "h:mm AM/PM", kDe).text, "6:00 PM");
}

TEST(FormatLocalizeRender, LocaleNumberSyntax) {
  EXPECT_EQ(Render(150000000000.0, "0.0E+00", kDe).text, "1.500E+08");  // text_format.text_sci_one_digit_mantissa
  EXPECT_EQ(Render(150000000000.0, "0.0E+00", kFr).text, "1.5E+10");
  EXPECT_EQ(Render(1234567, "#,##0", kDe).text, "1234567,0");     // text_format.text_thousands_basic
  EXPECT_EQ(Render(-1234.5, "# ##0,00", kFr).text, "-1 234,50");  // locale_tokens.text_group_space_decimal_comma
  EXPECT_EQ(Render(-1234.5, "0.00", kFr).text, "-12.35");         // locale_tokens.text_decimal_point
  EXPECT_EQ(Render(-1234.5, "Standard", kDe).text, "-1234,5");    // locale_tokens.text_general_standard
  // text_format.text_fraction_single_digit: the fr-FR blank groups digits, so `# ?/?` is malformed.
  EXPECT_EQ(Render(0.5, "# ?/?", kFr).status, FormatStatus::kValueError);
  EXPECT_EQ(Render(0.5, "# ?/?", kDe).text, " 1/2");
  // A stored code renders with the locale's separators.
  EXPECT_EQ(Render(1234.5, "#,##0.00", kDe, FormatDialect::kStored).text, "1.234,50");
}

TEST(FormatLocalizeRender, InvariantGapsMeasuredAcrossLocales) {
  const ExcelProfile en = mac_365_en_us_profile();
  EXPECT_EQ(Render(-1234.5, "0,0E+00", en).text, "-1,235E+00");     // locale_tokens.text_sci_comma
  EXPECT_EQ(Render(-1234.5, "#.##0,00", en).text, "-1234.5000");    // locale_tokens.text_group_dot_decimal_comma
  EXPECT_EQ(Render(45383, "bbbb", kDe).text, "2567");               // locale_tokens.text_buddhist_year_bbbb
  EXPECT_EQ(Render(45383, "bb", en).text, "67");                    // locale_tokens.text_buddhist_year_bb
  EXPECT_EQ(Render(0.75, "h:mm 上午/下午", en).text, "6:00 下午");  // locale_tokens.text_ampm_zh
  EXPECT_EQ(Render(0.25, "h:mm 上午/下午", kFr).text, "6:00 上午");
  EXPECT_EQ(Render(0.75, "h:mm 午前/午後", en).text, "18:00 午前/午後");  // locale_tokens.text_ampm_ja
}

TEST(FormatLocalizeRender, DbNumDigitsFollowTheLocale) {
  // text_format.text_dbnum1_with_era
  EXPECT_EQ(Render(45383, "[DBNum1]ggge年m月d日", kZh).text, "二○二四年四月一日");
  EXPECT_EQ(Render(1234, "[DBNum2]0", kZh).text, "壹贰叁肆");  // text_format.text_dbnum2
}

}  // namespace
}  // namespace text_format
}  // namespace formulon
