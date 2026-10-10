#include "excel_locale.h"

#include <array>
#include <cstddef>
#include <string_view>

#include "excel_profile.h"
#include "gtest/gtest.h"
#include "workbook.h"

namespace formulon {
namespace {

constexpr std::array<const char*, 28> kAllProfileIds = {
    "mac-365-ja_JP", "win-365-ja_JP", "mac-365-en_US", "win-365-en_US", "mac-365-de_DE", "win-365-de_DE",
    "mac-365-fr_FR", "win-365-fr_FR", "mac-365-zh_CN", "win-365-zh_CN", "mac-365-ko_KR", "win-365-ko_KR",
    "mac-365-th_TH", "win-365-th_TH", "mac-365-es_ES", "win-365-es_ES", "mac-365-es_MX", "win-365-es_MX",
    "mac-365-pt_BR", "win-365-pt_BR", "mac-365-ru_RU", "win-365-ru_RU", "mac-365-zh_TW", "win-365-zh_TW",
    "mac-365-it_IT", "win-365-it_IT", "mac-365-nl_NL", "win-365-nl_NL",
};

constexpr ExcelProfile profile_of(ExcelHost host, ExcelLocale locale) noexcept {
  return ExcelProfile{host, locale};
}

TEST(ExcelLocale, ProfileIdsRoundTrip) {
  ASSERT_EQ(detail::kExcelProfileIds.size(), kAllProfileIds.size());
  for (const char* id : kAllProfileIds) {
    ExcelProfile profile = default_excel_profile();
    ASSERT_TRUE(parse_excel_profile_id(id, &profile)) << id;
    EXPECT_STREQ(excel_profile_id(profile), id);
  }
}

TEST(ExcelLocale, InvalidIdsLeaveTheProfileUnchanged) {
  for (const char* id :
       {"", "mac-365-en_GB", "mac-365-pt_PT", "win-365-en-US", "MAC-365-ja_JP", "win-365-en_US-extra"}) {
    ExcelProfile profile = mac_365_ja_jp_profile();
    EXPECT_FALSE(parse_excel_profile_id(id, &profile)) << id;
    EXPECT_TRUE(same_profile(profile, mac_365_ja_jp_profile())) << id;
  }
}

TEST(ExcelLocale, RuntimeDefaultIsWindowsEnglish) {
  EXPECT_STREQ(excel_profile_id(default_excel_profile()), "win-365-en_US");
  EXPECT_STREQ(excel_profile_id(Workbook::create().excel_profile()), "win-365-en_US");
}

TEST(ExcelLocale, UnknownProfileFallsBackToTheDefaultId) {
  const ExcelProfile unknown{ExcelHost::kMac365, static_cast<ExcelLocale>(0x7F)};
  EXPECT_STREQ(excel_profile_id(unknown), "win-365-en_US");
}

void ExpectFactsEqual(const LocaleFacts& lhs, const LocaleFacts& rhs) {
  EXPECT_EQ(lhs.decimal_separator, rhs.decimal_separator);
  EXPECT_EQ(lhs.group_separator, rhs.group_separator);
  EXPECT_EQ(lhs.list_separator, rhs.list_separator);
  EXPECT_EQ(lhs.array_column_separator, rhs.array_column_separator);
  EXPECT_EQ(lhs.array_row_separator, rhs.array_row_separator);
  EXPECT_EQ(lhs.true_name, rhs.true_name);
  EXPECT_EQ(lhs.false_name, rhs.false_name);
  EXPECT_EQ(lhs.logical_skips_english_bool_text, rhs.logical_skips_english_bool_text);
  EXPECT_EQ(lhs.error_names, rhs.error_names);
  EXPECT_EQ(lhs.r1c1_row, rhs.r1c1_row);
  EXPECT_EQ(lhs.r1c1_col, rhs.r1c1_col);
  EXPECT_EQ(lhs.r1c1_open, rhs.r1c1_open);
  EXPECT_EQ(lhs.r1c1_close, rhs.r1c1_close);
  EXPECT_EQ(lhs.cell_type_blank, rhs.cell_type_blank);
  EXPECT_EQ(lhs.cell_type_label, rhs.cell_type_label);
  EXPECT_EQ(lhs.cell_type_value, rhs.cell_type_value);
  EXPECT_EQ(lhs.cell_format_general, rhs.cell_format_general);
  EXPECT_EQ(lhs.month_long, rhs.month_long);
  EXPECT_EQ(lhs.month_short, rhs.month_short);
  EXPECT_EQ(lhs.day_long, rhs.day_long);
  EXPECT_EQ(lhs.day_short, rhs.day_short);
  EXPECT_EQ(lhs.format_letters.year, rhs.format_letters.year);
  EXPECT_EQ(lhs.format_letters.month, rhs.format_letters.month);
  EXPECT_EQ(lhs.format_letters.day, rhs.format_letters.day);
  EXPECT_EQ(lhs.format_letters.hour, rhs.format_letters.hour);
  EXPECT_EQ(lhs.format_letters.minute, rhs.format_letters.minute);
  EXPECT_EQ(lhs.format_letters.second, rhs.format_letters.second);
  EXPECT_EQ(lhs.format_letters.case_sensitive, rhs.format_letters.case_sensitive);
  EXPECT_EQ(lhs.format_letters.minute_unconditional, rhs.format_letters.minute_unconditional);
  EXPECT_EQ(lhs.format_letters.month_contextual, rhs.format_letters.month_contextual);
  for (std::size_t i = 0; i < lhs.format_letter_aliases.size(); ++i) {
    EXPECT_EQ(lhs.format_letter_aliases[i].spelling, rhs.format_letter_aliases[i].spelling);
    EXPECT_EQ(lhs.format_letter_aliases[i].letter, rhs.format_letter_aliases[i].letter);
  }
  EXPECT_EQ(lhs.format_letters.weekday, rhs.format_letters.weekday);
  EXPECT_EQ(lhs.currency.symbol, rhs.currency.symbol);
  EXPECT_EQ(lhs.currency.suffix, rhs.currency.suffix);
  EXPECT_EQ(lhs.currency.space, rhs.currency.space);
  EXPECT_EQ(lhs.currency.negative_parens, rhs.currency.negative_parens);
  EXPECT_EQ(lhs.currency.minus_after_symbol, rhs.currency.minus_after_symbol);
  EXPECT_EQ(lhs.currency.negative_zero_signed, rhs.currency.negative_zero_signed);
  EXPECT_EQ(lhs.currency.default_decimals, rhs.currency.default_decimals);
  EXPECT_EQ(lhs.accepted_currency, rhs.accepted_currency);
  EXPECT_EQ(lhs.usdollar_symbol, rhs.usdollar_symbol);
  EXPECT_EQ(lhs.date_order, rhs.date_order);
  EXPECT_EQ(lhs.dotted_date, rhs.dotted_date);
  EXPECT_EQ(lhs.kanji_ymd_text, rhs.kanji_ymd_text);
  EXPECT_EQ(lhs.kanji_time_text, rhs.kanji_time_text);
  EXPECT_EQ(lhs.japanese_era, rhs.japanese_era);
  EXPECT_EQ(lhs.r_letter, rhs.r_letter);
  EXPECT_EQ(lhs.hangul_ymd_text, rhs.hangul_ymd_text);
  EXPECT_EQ(lhs.english_month_names, rhs.english_month_names);
  EXPECT_EQ(lhs.hyphen_english_months, rhs.hyphen_english_months);
  EXPECT_EQ(lhs.fractional_seconds, rhs.fractional_seconds);
  EXPECT_EQ(lhs.meridiem_dot_attached, rhs.meridiem_dot_attached);
  EXPECT_EQ(lhs.meridiem_dot_spaced, rhs.meridiem_dot_spaced);
  EXPECT_EQ(lhs.am_name, rhs.am_name);
  EXPECT_EQ(lhs.pm_name, rhs.pm_name);
  EXPECT_EQ(lhs.short_meridiem, rhs.short_meridiem);
  EXPECT_EQ(lhs.dbcs_codepage, rhs.dbcs_codepage);
  EXPECT_EQ(lhs.halfwidth_kana_single_byte, rhs.halfwidth_kana_single_byte);
  EXPECT_EQ(lhs.fullwidth_numeric_text, rhs.fullwidth_numeric_text);
  EXPECT_EQ(lhs.char_snaps_near_integer, rhs.char_snaps_near_integer);
  EXPECT_EQ(lhs.phonetic, rhs.phonetic);
  EXPECT_EQ(lhs.criteria_header_keeps_halfwidth_kana, rhs.criteria_header_keeps_halfwidth_kana);
  EXPECT_EQ(lhs.bang_escape, rhs.bang_escape);
  EXPECT_EQ(lhs.dbnum, rhs.dbnum);
  EXPECT_EQ(lhs.thai_digit_letter, rhs.thai_digit_letter);
  EXPECT_EQ(lhs.blank_date_letter, rhs.blank_date_letter);
  EXPECT_EQ(lhs.fullwidth_syntax_fold, rhs.fullwidth_syntax_fold);
  EXPECT_EQ(lhs.english_general, rhs.english_general);
  EXPECT_EQ(lhs.format_rejects_dot, rhs.format_rejects_dot);
  EXPECT_EQ(lhs.general_alias, rhs.general_alias);
  EXPECT_EQ(lhs.color_names, rhs.color_names);
  EXPECT_EQ(lhs.color_index_prefix, rhs.color_index_prefix);
  EXPECT_EQ(lhs.grand_total, rhs.grand_total);
  EXPECT_EQ(lhs.hierarchy_grand_total, rhs.hierarchy_grand_total);
  EXPECT_EQ(lhs.row_field_prefix, rhs.row_field_prefix);
  EXPECT_EQ(lhs.column_field_prefix, rhs.column_field_prefix);
  EXPECT_EQ(lhs.value_field_prefix, rhs.value_field_prefix);
  EXPECT_EQ(lhs.weekday_long, rhs.weekday_long);
  EXPECT_EQ(lhs.weekday_short, rhs.weekday_short);
}

using Names7 = std::array<std::string_view, 7>;
using Names12 = std::array<std::string_view, 12>;

constexpr std::array<std::string_view, kErrorNameCount> kEnglishErrors{
    "#NULL!", "#DIV/0!", "#VALUE!",   "#REF!",     "#NAME?",     "#NUM!",  "#N/A",     "#GETTING_DATA", "#SPILL!",
    "#CALC!", "#FIELD!", "#BLOCKED!", "#CONNECT!", "#EXTERNAL!", "#BUSY!", "#PYTHON!", "#UNKNOWN!",
};
constexpr Names12 kEnglishMonthsLong{"January", "February", "March",     "April",   "May",      "June",
                                     "July",    "August",   "September", "October", "November", "December"};
constexpr Names12 kEnglishMonthsShort{"Jan", "Feb", "Mar", "Apr", "May", "Jun",
                                      "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"};
constexpr Names7 kEnglishDaysLong{"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"};
constexpr Names7 kEnglishDaysShort{"Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"};

void ExpectInvariantFormatLetters(const FormatLetters& letters) {
  EXPECT_EQ(letters.year, 'y');
  EXPECT_EQ(letters.month, 'm');
  EXPECT_EQ(letters.day, 'd');
  EXPECT_EQ(letters.hour, 'h');
  EXPECT_EQ(letters.minute, 'm');
  EXPECT_EQ(letters.second, 's');
  EXPECT_FALSE(letters.case_sensitive);
  EXPECT_FALSE(letters.minute_unconditional);
  EXPECT_FALSE(letters.month_contextual);
  EXPECT_EQ(letters.weekday, 'a');
}

// Shared by the ja-JP and en-US rows: the invariant separators and names.
void ExpectInvariantTokens(const LocaleFacts& facts) {
  EXPECT_EQ(facts.decimal_separator, '.');
  EXPECT_EQ(facts.group_separator, ',');
  EXPECT_EQ(facts.list_separator, ',');
  EXPECT_EQ(facts.array_column_separator, ',');
  EXPECT_EQ(facts.array_row_separator, ';');
  EXPECT_EQ(facts.true_name, "TRUE");
  EXPECT_EQ(facts.false_name, "FALSE");
  EXPECT_EQ(facts.r1c1_row, 'R');
  EXPECT_EQ(facts.r1c1_col, 'C');
  EXPECT_EQ(facts.r1c1_open, '[');
  EXPECT_EQ(facts.r1c1_close, ']');
  EXPECT_EQ(facts.cell_type_blank, 'b');
  EXPECT_EQ(facts.cell_type_label, 'l');
  EXPECT_EQ(facts.cell_type_value, 'v');
  EXPECT_EQ(facts.cell_format_general, 'G');
  EXPECT_EQ(facts.month_long, kEnglishMonthsLong);
  EXPECT_EQ(facts.month_short, kEnglishMonthsShort);
  EXPECT_EQ(facts.day_long, kEnglishDaysLong);
  EXPECT_EQ(facts.day_short, kEnglishDaysShort);
  ExpectInvariantFormatLetters(facts.format_letters);
  EXPECT_FALSE(facts.dotted_date);
}

void ExpectJapaneseFacts(const LocaleFacts& facts) {
  ExpectInvariantTokens(facts);
  std::array<std::string_view, kErrorNameCount> japanese_errors = kEnglishErrors;
  japanese_errors[static_cast<std::size_t>(ErrorCode::Spill)] = "#スピル!";
  EXPECT_EQ(facts.error_names, japanese_errors);
  EXPECT_EQ(facts.currency.symbol, "¥");
  EXPECT_FALSE(facts.currency.suffix);
  EXPECT_FALSE(facts.currency.space);
  EXPECT_FALSE(facts.currency.negative_parens);
  EXPECT_TRUE(facts.currency.negative_zero_signed);
  EXPECT_EQ(facts.currency.default_decimals, 0U);
  EXPECT_EQ(facts.accepted_currency, (std::array<std::string_view, 4>{"$", "€", "¥", "￥"}));
  EXPECT_EQ(facts.date_order, DateOrder::kYMD);
  EXPECT_TRUE(facts.kanji_ymd_text);
  EXPECT_TRUE(facts.kanji_time_text);
  EXPECT_TRUE(facts.japanese_era);
  EXPECT_EQ(facts.dbcs_codepage, DbcsCodepage::kJis0208);
  EXPECT_TRUE(facts.halfwidth_kana_single_byte);
  EXPECT_TRUE(facts.fullwidth_numeric_text);
  EXPECT_FALSE(facts.char_snaps_near_integer);
  EXPECT_TRUE(facts.phonetic);
  EXPECT_TRUE(facts.criteria_header_keeps_halfwidth_kana);
  EXPECT_TRUE(facts.bang_escape);
  ASSERT_NE(facts.dbnum, nullptr);
  EXPECT_TRUE(facts.fullwidth_syntax_fold);
  EXPECT_EQ(facts.general_alias, "G/標準");
  EXPECT_EQ(facts.color_names, (std::array<std::string_view, 8>{"黒", "青", "水", "緑", "紫", "赤", "白", "黄"}));
  EXPECT_EQ(facts.color_index_prefix, "色");
  EXPECT_EQ((*facts.dbnum)[0].digits,
            (std::array<std::string_view, 10>{"〇", "一", "二", "三", "四", "五", "六", "七", "八", "九"}));
  EXPECT_EQ((*facts.dbnum)[1].digits,
            (std::array<std::string_view, 10>{"〇", "壱", "弐", "参", "四", "伍", "六", "七", "八", "九"}));
  EXPECT_EQ((*facts.dbnum)[1].place_units, (std::array<std::string_view, 3>{"拾", "百", "阡"}));
  EXPECT_FALSE((*facts.dbnum)[0].place_one);
  EXPECT_TRUE((*facts.dbnum)[1].place_one);
  EXPECT_FALSE((*facts.dbnum)[0].zero_filler);
  EXPECT_TRUE((*facts.dbnum)[3].place_units[0].empty());
  EXPECT_EQ(facts.grand_total, "合計");
  EXPECT_EQ(facts.hierarchy_grand_total, "総計");
  EXPECT_EQ(facts.row_field_prefix, "行フィールド ");
  EXPECT_EQ(facts.column_field_prefix, "列フィールド ");
  EXPECT_EQ(facts.value_field_prefix, "値 ");
  EXPECT_EQ(facts.weekday_long, (Names7{"日曜日", "月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日"}));
  EXPECT_EQ(facts.weekday_short, (Names7{"日", "月", "火", "水", "木", "金", "土"}));
}

void ExpectEnglishFacts(const LocaleFacts& facts) {
  ExpectInvariantTokens(facts);
  EXPECT_EQ(facts.error_names, kEnglishErrors);
  EXPECT_EQ(facts.currency.symbol, "$");
  EXPECT_FALSE(facts.currency.suffix);
  EXPECT_FALSE(facts.currency.space);
  EXPECT_TRUE(facts.currency.negative_parens);
  EXPECT_TRUE(facts.currency.negative_zero_signed);
  EXPECT_EQ(facts.currency.default_decimals, 2U);
  EXPECT_EQ(facts.accepted_currency, (std::array<std::string_view, 4>{"$", "€", "", ""}));
  EXPECT_EQ(facts.date_order, DateOrder::kMDY);
  EXPECT_FALSE(facts.kanji_ymd_text);
  EXPECT_FALSE(facts.kanji_time_text);
  EXPECT_FALSE(facts.japanese_era);
  EXPECT_EQ(facts.dbcs_codepage, DbcsCodepage::kNone);
  EXPECT_FALSE(facts.halfwidth_kana_single_byte);
  EXPECT_FALSE(facts.fullwidth_numeric_text);
  EXPECT_TRUE(facts.char_snaps_near_integer);
  EXPECT_FALSE(facts.phonetic);
  EXPECT_FALSE(facts.criteria_header_keeps_halfwidth_kana);
  EXPECT_FALSE(facts.bang_escape);
  EXPECT_EQ(facts.dbnum, nullptr);
  EXPECT_FALSE(facts.fullwidth_syntax_fold);
  EXPECT_EQ(facts.general_alias, "");
  EXPECT_EQ(facts.color_names,
            (std::array<std::string_view, 8>{"Black", "Blue", "Cyan", "Green", "Magenta", "Red", "White", "Yellow"}));
  EXPECT_EQ(facts.color_index_prefix, "Color");
  EXPECT_EQ(facts.grand_total, "Total");
  EXPECT_EQ(facts.hierarchy_grand_total, "Grand Total");
  EXPECT_EQ(facts.row_field_prefix, "Row Field ");
  EXPECT_EQ(facts.column_field_prefix, "Column Field ");
  EXPECT_EQ(facts.value_field_prefix, "Value ");
  EXPECT_EQ(facts.weekday_long, kEnglishDaysLong);
  EXPECT_EQ(facts.weekday_short, kEnglishDaysShort);
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

TEST(ExcelLocale, EveryLocaleNamesEveryErrorMonthAndDay) {
  for (const char* id : kAllProfileIds) {
    ExcelProfile profile = default_excel_profile();
    ASSERT_TRUE(parse_excel_profile_id(id, &profile)) << id;
    const LocaleFacts& facts = locale_facts(profile);
    for (std::size_t i = 0; i < facts.error_names.size(); ++i) {
      EXPECT_FALSE(facts.error_names[i].empty()) << id << " error " << i;
    }
    // Names are reached through a month or day letter; a locale without one
    // (ru-RU reads no Latin date letter) names none.
    const bool has_month = facts.format_letters.month != '\0';
    const bool has_day = facts.format_letters.day != '\0';
    for (const std::string_view name : facts.month_long) {
      EXPECT_EQ(name.empty(), !has_month) << id;
    }
    for (const std::string_view name : facts.month_short) {
      EXPECT_EQ(name.empty(), !has_month) << id;
    }
    for (const std::string_view name : facts.day_long) {
      EXPECT_EQ(name.empty(), !has_day) << id;
    }
    for (const std::string_view name : facts.day_short) {
      EXPECT_EQ(name.empty(), !has_day) << id;
    }
    for (const std::string_view name : facts.weekday_long) {
      EXPECT_FALSE(name.empty()) << id;
    }
    for (const std::string_view name : facts.weekday_short) {
      EXPECT_FALSE(name.empty()) << id;
    }
    if (facts.dbnum != nullptr) {
      for (const DbnumStyle& style : *facts.dbnum) {
        for (const std::string_view digit : style.digits) {
          EXPECT_FALSE(digit.empty()) << id;
        }
      }
    }
    for (const std::string_view name : facts.color_names) {
      EXPECT_FALSE(name.empty()) << id;
    }
    EXPECT_FALSE(facts.color_index_prefix.empty()) << id;
    EXPECT_FALSE(facts.true_name.empty()) << id;
    EXPECT_FALSE(facts.false_name.empty()) << id;
    EXPECT_FALSE(facts.currency.symbol.empty()) << id;
  }
}

// Values below are the Mac Excel 365 goldens under tests/oracle/targets/mac-365-<locale>/golden.
TEST(ExcelLocale, NewLocaleSeparatorsAndBooleansMatchTheGoldens) {
  struct Case {
    ExcelLocale locale;
    char decimal;
    char group;
    char list;
    char array_column;
    char array_row;
    std::string_view true_name;
    std::string_view false_name;
    DateOrder date_order;
  };
  // locale_tokens.fixed_negative, formulatext_bool_literal, formulatext_array_constant,
  // bool_text_true / bool_text_false; value_coercion_probes.value_slash_date_short.
  constexpr std::array<Case, 12> kCases{{
      {ExcelLocale::kDeDE, ',', '.', ';', '.', ';', "WAHR", "FALSCH", DateOrder::kDMY},
      {ExcelLocale::kFrFR, ',', ' ', ';', '.', ';', "VRAI", "FAUX", DateOrder::kDMY},
      {ExcelLocale::kZhCN, '.', ',', ',', ',', ';', "TRUE", "FALSE", DateOrder::kYMD},
      {ExcelLocale::kKoKR, '.', ',', ',', ',', ';', "TRUE", "FALSE", DateOrder::kYMD},
      {ExcelLocale::kThTH, '.', ',', ',', ',', ';', "TRUE", "FALSE", DateOrder::kDMY},
      {ExcelLocale::kRuRU, ',', ' ', ';', ';', ':', "ИСТИНА", "ЛОЖЬ", DateOrder::kDMY},
      {ExcelLocale::kZhTW, '.', ',', ',', ',', ';', "TRUE", "FALSE", DateOrder::kYMD},
      {ExcelLocale::kItIT, ',', '.', ';', '\\', '.', "VERO", "FALSO", DateOrder::kDMY},
      {ExcelLocale::kNlNL, ',', '.', ';', '\\', ';', "WAAR", "ONWAAR", DateOrder::kDMY},
      {ExcelLocale::kEsES, ',', '.', ';', '\\', ';', "VERDADERO", "FALSO", DateOrder::kDMY},
      {ExcelLocale::kEsMX, '.', ',', ',', ',', ';', "VERDADERO", "FALSO", DateOrder::kDMY},
      {ExcelLocale::kPtBR, ',', '.', ';', '\\', ';', "VERDADEIRO", "FALSO", DateOrder::kDMY},
  }};
  for (const Case& c : kCases) {
    for (const ExcelHost host : {ExcelHost::kMac365, ExcelHost::kWin365}) {
      const ExcelProfile profile = profile_of(host, c.locale);
      SCOPED_TRACE(excel_profile_id(profile));
      const LocaleFacts& facts = locale_facts(profile);
      EXPECT_EQ(facts.decimal_separator, c.decimal);
      EXPECT_EQ(facts.group_separator, c.group);
      EXPECT_EQ(facts.list_separator, c.list);
      EXPECT_EQ(facts.array_column_separator, c.array_column);
      EXPECT_EQ(facts.array_row_separator, c.array_row);
      EXPECT_EQ(facts.true_name, c.true_name);
      EXPECT_EQ(facts.false_name, c.false_name);
      EXPECT_EQ(facts.date_order, c.date_order);
    }
  }
  // locale_tokens.formulatext_error_literal, arraytotext_error_literal_in_array_default
  EXPECT_EQ(locale_facts(profile_of(ExcelHost::kMac365, ExcelLocale::kDeDE))
                .error_names[static_cast<std::size_t>(ErrorCode::NA)],
            "#NV");
  EXPECT_EQ(locale_facts(profile_of(ExcelHost::kMac365, ExcelLocale::kFrFR))
                .error_names[static_cast<std::size_t>(ErrorCode::NA)],
            "#N/A");
  // arraytotext.arraytotext_error_literal_in_array_default, arraytotext_only_error_cells_default
  for (const ExcelLocale locale : {ExcelLocale::kEsES, ExcelLocale::kEsMX, ExcelLocale::kPtBR}) {
    EXPECT_EQ(locale_facts(profile_of(ExcelHost::kMac365, locale)).error_names[static_cast<std::size_t>(ErrorCode::NA)],
              "#N/D");
  }
  EXPECT_EQ(locale_facts(profile_of(ExcelHost::kMac365, ExcelLocale::kEsES))
                .error_names[static_cast<std::size_t>(ErrorCode::Div0)],
            "#¡DIV/0!");
  EXPECT_EQ(locale_facts(profile_of(ExcelHost::kMac365, ExcelLocale::kEsMX))
                .error_names[static_cast<std::size_t>(ErrorCode::Div0)],
            "#DIV/0!");
}

TEST(ExcelLocale, CodePageDependsOnLocaleAndHost) {
  EXPECT_EQ(sbcs_codepage(mac_365_ja_jp_profile()), SbcsCodepage::kCp932);
  EXPECT_EQ(sbcs_codepage(win_365_ja_jp_profile()), SbcsCodepage::kCp932);
  EXPECT_EQ(sbcs_codepage(mac_365_en_us_profile()), SbcsCodepage::kMacRoman);
  EXPECT_EQ(sbcs_codepage(win_365_en_us_profile()), SbcsCodepage::kWindows1252);
  for (const ExcelHost host : {ExcelHost::kMac365, ExcelHost::kWin365}) {
    const SbcsCodepage host_page = host == ExcelHost::kMac365 ? SbcsCodepage::kMacRoman : SbcsCodepage::kWindows1252;
    EXPECT_EQ(sbcs_codepage(profile_of(host, ExcelLocale::kDeDE)), host_page);
    EXPECT_EQ(sbcs_codepage(profile_of(host, ExcelLocale::kFrFR)), host_page);
    EXPECT_EQ(sbcs_codepage(profile_of(host, ExcelLocale::kZhCN)), SbcsCodepage::kDbcsHighBlank);
    EXPECT_EQ(sbcs_codepage(profile_of(host, ExcelLocale::kKoKR)), SbcsCodepage::kDbcsHighBlank);
    EXPECT_EQ(sbcs_codepage(profile_of(host, ExcelLocale::kThTH)), SbcsCodepage::kMacThai);
    EXPECT_EQ(sbcs_codepage(profile_of(host, ExcelLocale::kRuRU)), SbcsCodepage::kMacCyrillic);
    EXPECT_EQ(sbcs_codepage(profile_of(host, ExcelLocale::kZhTW)), SbcsCodepage::kDbcsHighBlank);
    EXPECT_EQ(sbcs_codepage(profile_of(host, ExcelLocale::kItIT)), host_page);
    EXPECT_EQ(sbcs_codepage(profile_of(host, ExcelLocale::kNlNL)), host_page);
  }
}

TEST(ExcelLocale, DbcsCodepageFollowsTheLocale) {
  EXPECT_EQ(locale_facts(profile_of(ExcelHost::kMac365, ExcelLocale::kDeDE)).dbcs_codepage, DbcsCodepage::kNone);
  EXPECT_EQ(locale_facts(profile_of(ExcelHost::kMac365, ExcelLocale::kFrFR)).dbcs_codepage, DbcsCodepage::kNone);
  EXPECT_EQ(locale_facts(profile_of(ExcelHost::kMac365, ExcelLocale::kZhCN)).dbcs_codepage, DbcsCodepage::kGb2312);
  EXPECT_EQ(locale_facts(profile_of(ExcelHost::kMac365, ExcelLocale::kKoKR)).dbcs_codepage, DbcsCodepage::kKsX1001);
  EXPECT_EQ(locale_facts(profile_of(ExcelHost::kMac365, ExcelLocale::kThTH)).dbcs_codepage, DbcsCodepage::kNone);
  EXPECT_EQ(locale_facts(profile_of(ExcelHost::kMac365, ExcelLocale::kZhTW)).dbcs_codepage, DbcsCodepage::kBig5);
  EXPECT_EQ(locale_facts(profile_of(ExcelHost::kMac365, ExcelLocale::kRuRU)).dbcs_codepage, DbcsCodepage::kNone);
}

TEST(ExcelLocale, WidthFoldingDependsOnHostExceptForChinese) {
  EXPECT_EQ(width_folding(mac_365_ja_jp_profile()), WidthFolding::kMac);
  EXPECT_EQ(width_folding(mac_365_en_us_profile()), WidthFolding::kMac);
  EXPECT_EQ(width_folding(win_365_ja_jp_profile()), WidthFolding::kWin);
  EXPECT_EQ(width_folding(win_365_en_us_profile()), WidthFolding::kWin);
  EXPECT_EQ(width_folding(profile_of(ExcelHost::kMac365, ExcelLocale::kKoKR)), WidthFolding::kMac);
  EXPECT_EQ(width_folding(profile_of(ExcelHost::kWin365, ExcelLocale::kKoKR)), WidthFolding::kWin);
  EXPECT_EQ(width_folding(profile_of(ExcelHost::kMac365, ExcelLocale::kZhCN)), WidthFolding::kNone);
  EXPECT_EQ(width_folding(profile_of(ExcelHost::kWin365, ExcelLocale::kZhCN)), WidthFolding::kNone);
  EXPECT_EQ(width_folding(profile_of(ExcelHost::kMac365, ExcelLocale::kZhTW)), WidthFolding::kNone);
  EXPECT_EQ(width_folding(profile_of(ExcelHost::kWin365, ExcelLocale::kZhTW)), WidthFolding::kNone);
}

}  // namespace
}  // namespace formulon
