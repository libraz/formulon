// Tests for the locale-aware number, boolean and error text helpers.

#include "eval/locale_text.h"

#include "eval/eval_profile_scope.h"
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

}  // namespace
}  // namespace formulon::eval
