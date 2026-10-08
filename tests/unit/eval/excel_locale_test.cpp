#include "excel_locale.h"

#include <array>
#include <string_view>

#include "excel_profile.h"
#include "gtest/gtest.h"
#include "workbook.h"

namespace formulon {
namespace {

TEST(ExcelLocale, ProfileIdsRoundTrip) {
  for (const char* id : {"mac-365-ja_JP", "win-365-ja_JP", "mac-365-en_US", "win-365-en_US"}) {
    ExcelProfile profile = default_excel_profile();
    ASSERT_TRUE(parse_excel_profile_id(id, &profile)) << id;
    EXPECT_STREQ(excel_profile_id(profile), id);
  }
}

TEST(ExcelLocale, InvalidIdsLeaveTheProfileUnchanged) {
  for (const char* id :
       {"", "mac-365-en_GB", "mac-365-de_DE", "win-365-en-US", "MAC-365-ja_JP", "win-365-en_US-extra"}) {
    ExcelProfile profile = mac_365_ja_jp_profile();
    EXPECT_FALSE(parse_excel_profile_id(id, &profile)) << id;
    EXPECT_TRUE(same_profile(profile, mac_365_ja_jp_profile())) << id;
  }
}

TEST(ExcelLocale, RuntimeDefaultIsWindowsEnglish) {
  EXPECT_STREQ(excel_profile_id(default_excel_profile()), "win-365-en_US");
  EXPECT_STREQ(excel_profile_id(Workbook::create().excel_profile()), "win-365-en_US");
}

void ExpectFactsEqual(const LocaleFacts& lhs, const LocaleFacts& rhs) {
  EXPECT_EQ(lhs.grand_total, rhs.grand_total);
  EXPECT_EQ(lhs.hierarchy_grand_total, rhs.hierarchy_grand_total);
  EXPECT_EQ(lhs.row_field_prefix, rhs.row_field_prefix);
  EXPECT_EQ(lhs.column_field_prefix, rhs.column_field_prefix);
  EXPECT_EQ(lhs.value_field_prefix, rhs.value_field_prefix);
  EXPECT_EQ(lhs.dollar_format, rhs.dollar_format);
  EXPECT_EQ(lhs.dollar_default_decimals, rhs.dollar_default_decimals);
  EXPECT_EQ(lhs.currency_symbol, rhs.currency_symbol);
  EXPECT_EQ(lhs.date_order, rhs.date_order);
  EXPECT_EQ(lhs.dbcs, rhs.dbcs);
  EXPECT_EQ(lhs.fullwidth_numeric_text, rhs.fullwidth_numeric_text);
  EXPECT_EQ(lhs.kanji_date_text, rhs.kanji_date_text);
  EXPECT_EQ(lhs.japanese_era, rhs.japanese_era);
  EXPECT_EQ(lhs.ja_format_syntax, rhs.ja_format_syntax);
  EXPECT_EQ(lhs.weekday_long, rhs.weekday_long);
  EXPECT_EQ(lhs.weekday_short, rhs.weekday_short);
  EXPECT_EQ(lhs.char_snaps_near_integer, rhs.char_snaps_near_integer);
  EXPECT_EQ(lhs.phonetic, rhs.phonetic);
  EXPECT_EQ(lhs.criteria_header_keeps_halfwidth_kana, rhs.criteria_header_keeps_halfwidth_kana);
}

void ExpectJapaneseFacts(const LocaleFacts& facts) {
  EXPECT_EQ(facts.grand_total, "合計");
  EXPECT_EQ(facts.hierarchy_grand_total, "総計");
  EXPECT_EQ(facts.row_field_prefix, "行フィールド ");
  EXPECT_EQ(facts.column_field_prefix, "列フィールド ");
  EXPECT_EQ(facts.value_field_prefix, "値 ");
  EXPECT_EQ(facts.dollar_format, "¥#,##0;¥-#,##0");
  EXPECT_EQ(facts.dollar_default_decimals, 0U);
  EXPECT_EQ(facts.currency_symbol, "¥");
  EXPECT_EQ(facts.date_order, DateOrder::kYMD);
  EXPECT_TRUE(facts.dbcs);
  EXPECT_TRUE(facts.fullwidth_numeric_text);
  EXPECT_TRUE(facts.kanji_date_text);
  EXPECT_TRUE(facts.japanese_era);
  EXPECT_TRUE(facts.ja_format_syntax);
  EXPECT_EQ(facts.weekday_long,
            (std::array<std::string_view, 7>{"日曜日", "月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日"}));
  EXPECT_EQ(facts.weekday_short, (std::array<std::string_view, 7>{"日", "月", "火", "水", "木", "金", "土"}));
  EXPECT_FALSE(facts.char_snaps_near_integer);
  EXPECT_TRUE(facts.phonetic);
  EXPECT_TRUE(facts.criteria_header_keeps_halfwidth_kana);
}

void ExpectEnglishFacts(const LocaleFacts& facts) {
  EXPECT_EQ(facts.grand_total, "Total");
  EXPECT_EQ(facts.hierarchy_grand_total, "Grand Total");
  EXPECT_EQ(facts.row_field_prefix, "Row Field ");
  EXPECT_EQ(facts.column_field_prefix, "Column Field ");
  EXPECT_EQ(facts.value_field_prefix, "Value ");
  EXPECT_EQ(facts.dollar_format, "$#,##0.00_);($#,##0.00)");
  EXPECT_EQ(facts.dollar_default_decimals, 2U);
  EXPECT_EQ(facts.currency_symbol, "");
  EXPECT_EQ(facts.date_order, DateOrder::kMDY);
  EXPECT_FALSE(facts.dbcs);
  EXPECT_FALSE(facts.fullwidth_numeric_text);
  EXPECT_FALSE(facts.kanji_date_text);
  EXPECT_FALSE(facts.japanese_era);
  EXPECT_FALSE(facts.ja_format_syntax);
  EXPECT_EQ(facts.weekday_long, (std::array<std::string_view, 7>{"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday",
                                                                 "Friday", "Saturday"}));
  EXPECT_EQ(facts.weekday_short, (std::array<std::string_view, 7>{"Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"}));
  EXPECT_TRUE(facts.char_snaps_near_integer);
  EXPECT_FALSE(facts.phonetic);
  EXPECT_FALSE(facts.criteria_header_keeps_halfwidth_kana);
}

TEST(ExcelLocale, LocaleFactsHaveExactJapaneseAndEnglishValues) {
  const LocaleFacts& mac_ja = locale_facts(mac_365_ja_jp_profile());
  const LocaleFacts& win_ja = locale_facts(win_365_ja_jp_profile());
  const LocaleFacts& mac_en = locale_facts(mac_365_en_us_profile());
  const LocaleFacts& win_en = locale_facts(win_365_en_us_profile());

  ExpectJapaneseFacts(mac_ja);
  ExpectJapaneseFacts(win_ja);
  ExpectEnglishFacts(mac_en);
  ExpectEnglishFacts(win_en);
  ExpectFactsEqual(mac_ja, win_ja);
  ExpectFactsEqual(mac_en, win_en);
}

TEST(ExcelLocale, CodePageDependsOnLocaleAndHost) {
  EXPECT_EQ(sbcs_codepage(mac_365_ja_jp_profile()), SbcsCodepage::kCp932);
  EXPECT_EQ(sbcs_codepage(win_365_ja_jp_profile()), SbcsCodepage::kCp932);
  EXPECT_EQ(sbcs_codepage(mac_365_en_us_profile()), SbcsCodepage::kMacRoman);
  EXPECT_EQ(sbcs_codepage(win_365_en_us_profile()), SbcsCodepage::kWindows1252);
}

TEST(ExcelLocale, WidthFoldingDependsOnHostOnly) {
  EXPECT_EQ(width_folding(mac_365_ja_jp_profile()), WidthFolding::kMac);
  EXPECT_EQ(width_folding(mac_365_en_us_profile()), WidthFolding::kMac);
  EXPECT_EQ(width_folding(win_365_ja_jp_profile()), WidthFolding::kWin);
  EXPECT_EQ(width_folding(win_365_en_us_profile()), WidthFolding::kWin);
}

}  // namespace
}  // namespace formulon
