#include "excel_locale.h"

namespace formulon {
namespace {

// Trailing comments name the measuring case: `locale_tokens.<id>` unless a
// suite is given. `// unmeasured` slots hold the en-US value.

using Names7 = std::array<std::string_view, 7>;
using Names12 = std::array<std::string_view, 12>;
using ErrorNames = std::array<std::string_view, kErrorNameCount>;
using DbnumDigits = std::array<std::array<std::string_view, 10>, 2>;

constexpr ErrorNames kEnglishErrorNames{
    "#NULL!", "#DIV/0!", "#VALUE!",   "#REF!",     "#NAME?",     "#NUM!",  "#N/A",     "#GETTING_DATA", "#SPILL!",
    "#CALC!", "#FIELD!", "#BLOCKED!", "#CONNECT!", "#EXTERNAL!", "#BUSY!", "#PYTHON!", "#UNKNOWN!",
};

// Only #N/A is localized by a capture (arraytotext.arraytotext_error_literal_in_array_default);
// #DIV/0! is measured unchanged and every other slot is unmeasured.
constexpr ErrorNames kGermanErrorNames{
    "#NULL!", "#DIV/0!", "#VALUE!",   "#REF!",     "#NAME?",     "#NUM!",  "#NV",      "#GETTING_DATA", "#SPILL!",
    "#CALC!", "#FIELD!", "#BLOCKED!", "#CONNECT!", "#EXTERNAL!", "#BUSY!", "#PYTHON!", "#UNKNOWN!",
};

constexpr Names12 kEnglishMonthsLong{"January", "February", "March",     "April",   "May",      "June",
                                     "July",    "August",   "September", "October", "November", "December"};
constexpr Names12 kEnglishMonthsShort{"Jan", "Feb", "Mar", "Apr", "May", "Jun",
                                      "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"};
constexpr Names7 kEnglishDaysLong{"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"};
constexpr Names7 kEnglishDaysShort{"Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"};

constexpr std::array<FormatLetterAlias, 10> kNoLetterAliases{};
// date_cyr_d_mm_yyyy, date_cyr_dd_m_yy, date_cyr_lower, time_cyr_lower, time_cyr_upper, elapsed_cyr_hours
constexpr std::array<FormatLetterAlias, 10> kRussianLetterAliases{{
    {"Г", 'y'},
    {"г", 'y'},
    {"М", 'M'},
    {"м", 'm'},
    {"Д", 'd'},
    {"д", 'd'},
    {"Ч", 'h'},
    {"ч", 'h'},
    {"С", 's'},
    {"с", 's'},
}};
constexpr FormatLetters kInvariantLetters{'y', 'm', 'd', 'h', 'm', 's', false, false, false, 'a'};

constexpr std::array<std::string_view, 8> kEnglishColorNames{"Black",   "Blue", "Cyan",  "Green",
                                                             "Magenta", "Red",  "White", "Yellow"};

constexpr std::array<std::string_view, 10> kAsciiDigits{"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"};

// @size-budget: 12 KB
constexpr LocaleFacts kJapaneseFacts{
    '.',                 // locale_tokens.fixed_negative
    ',',                 // fixed_negative
    ',',                 // formulatext_bool_literal
    ',',                 // formulatext_array_constant
    ';',                 // formulatext_array_constant
    "TRUE",              // bool_text_true
    "FALSE",             // bool_text_false
    false,               // text_to_bool_probes.text_bool_and_whitespace, text_bool_and_two_true
    kEnglishErrorNames,  // arraytotext.arraytotext_only_error_cells_default; others unmeasured
    'R',
    'C',
    '[',
    ']',                  // references.address_r1c1_row_abs_col_rel
    'b',                  // cell_type_blank
    'l',                  // cell_type_text
    'v',                  // cell_type_number
    'G',                  // cell.cell_format_general
    kEnglishMonthsLong,   // months_mmmm
    kEnglishMonthsShort,  // months_mmm
    kEnglishDaysLong,     // weekdays_dddd
    kEnglishDaysShort,    // weekdays_ddd
    kInvariantLetters,    // months_yyyy, months_d, weekdays_aaaa
    kNoLetterAliases,
    {"¥", false, false, false, true, true, 0U},  // text.dollar_negative_minus_sign, text.dollar_rounds_to_negative_zero
    {"$", "€", "¥", "￥"},                       // value_dollar_prefix, value_euro_prefix, value_yen_prefix
    "$",                                         // usdollar.usdollar_two_decimals
    DateOrder::kYMD,
    false,                  // value_dotted_ymd
    true,                   // datevalue_timevalue.datevalue_kanji_with_terminator
    true,                   // datevalue_timevalue.timevalue_jp_kanji_units
    true,                   // datevalue_timevalue.datevalue_era_reiwa_full
    RLetter::kJapaneseEra,  // locale_tokens.text_letter_r_date, text_letter_rr_date
    false,                  // locale_tokens.value_korean_ymd
    true,                   // value_month_name_en
    true,                   // value_coercion_probes.value_d_mmm_yy
    true,                   // datevalue_timevalue.timevalue_fractional_seconds, timevalue_comma_fraction
    true,                   // datevalue_timevalue.timevalue_trailing_dot
    true,                   // datevalue_timevalue.timevalue_space_dot
    "AM",                   // text_format.text_time_am
    "PM",                   // text_format.text_time_pm
    true,                   // text_format.text_time_a_p
    DbcsCodepage::kJis0208,
    true,      // code_char_jp_probes.char_halfwidth_kata_177
    true,      // value_numbervalue.value_fullwidth_digits
    false,     // code_char_jp_probes.near_char
    true,      // lazy_forms.lazy_phonetic_unannotated_cell
    true,      // dfunc_kana_folding_probes.dsum_criteria_header_halfwidth_vs_fullwidth_db_header
    true,      // text_format.text_four_section_text
    true,      // text_format.text_dbnum1
    false,     // locale_tokens.text_letter_t_digits
    '\0',      // locale_tokens.text_letter_x_positive, text_letter_x_negative, text_slash_letters
    true,      // existing ja behaviour; unmeasured
    false,     // text_general_english
    false,     // text_decimal_point
    "G/標準",  // text_general_g_hyojun
    {"黒", "青", "水", "緑", "紫", "赤", "白", "黄"},  // text_color_aka
    DbnumDigits{{
        {"〇", "一", "二", "三", "四", "五", "六", "七", "八", "九"},
        {"零", "壱", "弐", "参", "四", "伍", "六", "七", "捌", "玖"},
    }},  // text_format.text_dbnum1, text_format.text_dbnum2
    "合計",
    "総計",
    "行フィールド ",
    "列フィールド ",
    "値 ",
    {"日曜日", "月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日"},  // weekdays_aaaa
    {"日", "月", "火", "水", "木", "金", "土"},                              // weekdays_aaa
};

constexpr LocaleFacts kEnglishFacts{
    '.',      // fixed_negative
    ',',      // fixed_negative
    ',',      // formulatext_bool_literal
    ',',      // formulatext_array_constant
    ';',      // formulatext_array_constant
    "TRUE",   // bool_text_true
    "FALSE",  // bool_text_false
    false,    // text_to_bool_probes.text_bool_and_whitespace, text_bool_and_two_true
    kEnglishErrorNames,
    'R',
    'C',
    '[',
    ']',                  // references.address_r1c1_row_abs_col_rel
    'b',                  // cell_type_blank
    'l',                  // cell_type_text
    'v',                  // cell_type_number
    'G',                  // cell.cell_format_general
    kEnglishMonthsLong,   // months_mmmm
    kEnglishMonthsShort,  // months_mmm
    kEnglishDaysLong,     // weekdays_dddd
    kEnglishDaysShort,    // weekdays_ddd
    kInvariantLetters,    // months_yyyy, months_d, weekdays_aaaa
    kNoLetterAliases,
    {"$", false, false, true, false, true, 2U},  // text.dollar_negative_minus_sign, text.dollar_rounds_to_negative_zero
    {"$", "€", "", ""},                          // value_dollar_prefix, value_euro_prefix
    "$",                                         // usdollar.usdollar_two_decimals
    DateOrder::kMDY,
    false,  // value_dotted_ymd
    false,
    false,
    false,
    RLetter::kLiteral,  // locale_tokens.text_letter_r_date, text_letter_rr_date
    false,              // locale_tokens.value_korean_ymd
    true,               // value_month_name_en
    true,               // value_coercion_probes.value_d_mmm_yy
    true,               // datevalue_timevalue.timevalue_fractional_seconds, timevalue_comma_fraction
    true,               // datevalue_timevalue.timevalue_trailing_dot
    true,               // datevalue_timevalue.timevalue_space_dot
    "AM",               // text_format.text_time_am
    "PM",               // text_format.text_time_pm
    true,               // text_format.text_time_a_p
    DbcsCodepage::kNone,
    false,
    false,
    true,  // code_char_jp_probes.near_char
    false,
    false,
    false,
    false,
    false,  // locale_tokens.text_letter_t_digits
    '\0',   // locale_tokens.text_letter_x_positive, text_letter_x_negative, text_slash_letters
    false,
    true,                                       // text_general_english
    false,                                      // text_decimal_point
    "",                                         // text_general_english
    kEnglishColorNames,                         // text_color_red
    DbnumDigits{{kAsciiDigits, kAsciiDigits}},  // text_format.text_dbnum1
    "Total",
    "Grand Total",
    "Row Field ",
    "Column Field ",
    "Value ",
    kEnglishDaysLong,   // weekdays_aaaa
    kEnglishDaysShort,  // weekdays_aaa
};

constexpr LocaleFacts kGermanFacts{
    ',',       // fixed_negative
    '.',       // fixed_negative
    ';',       // formulatext_bool_literal
    '.',       // formulatext_array_constant
    ';',       // formulatext_array_constant
    "WAHR",    // bool_text_true
    "FALSCH",  // bool_text_false
    true,      // text_to_bool_probes.text_bool_and_whitespace, text_bool_and_two_true
    kGermanErrorNames,
    'Z',
    'S',
    '(',
    ')',  // references.address_r1c1_row_abs_col_rel
    'b',  // cell_type_blank
    'l',  // cell_type_text
    'w',  // cell_type_number
    'S',  // cell.cell_format_general
    {"Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November",
     "Dezember"},                                                                          // months_umumumum
    {"Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"},  // months_umumum
    {"Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"},     // weekdays_utututut
    {"So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"},                                            // weekdays_ututut
    // months_ujujujuj, months_umumumum, months_utut, text_format.text_datetime_combined,
    // text_format.text_time_hh_mm_ss, weekdays_aaaa
    {'J', 'M', 'T', 'h', 'm', 's', true, true, false, 'a'},
    kNoLetterAliases,
    {"€", true, true, false, false, true, 2U},  // text.dollar_negative_minus_sign, text.dollar_rounds_to_negative_zero
    {"€", "", "", ""},                          // value_euro_prefix, value_dollar_prefix
    "",                                         // usdollar.usdollar_two_decimals
    DateOrder::kDMY,                            // value_coercion_probes.value_slash_date_short
    true,                                       // value_dotted_dmy
    false,                                      // datevalue_timevalue.datevalue_kanji_with_terminator
    false,                                      // datevalue_timevalue.timevalue_jp_kanji_units
    false,                                      // datevalue_timevalue.datevalue_era_reiwa_full
    RLetter::kLiteral,                          // locale_tokens.text_letter_r_date, text_letter_rr_date
    false,                                      // locale_tokens.value_korean_ymd
    false,                                      // value_month_name_en
    true,                                       // value_coercion_probes.value_d_mmm_yy
    true,                 // datevalue_timevalue.timevalue_fractional_seconds, timevalue_comma_fraction
    true,                 // datevalue_timevalue.timevalue_trailing_dot
    true,                 // datevalue_timevalue.timevalue_space_dot
    "AM",                 // text_format.text_time_am
    "PM",                 // text_format.text_time_pm
    true,                 // text_format.text_time_a_p
    DbcsCodepage::kNone,  // lenb_hangul
    false,
    false,       // value_numbervalue.value_fullwidth_digits
    true,        // code_char_jp_probes.near_char
    false,       // lazy_forms.lazy_phonetic_unannotated_cell
    false,       // dfunc_kana_folding_probes.dsum_criteria_header_halfwidth_vs_fullwidth_db_header
    false,       // text_format.text_four_section_text
    false,       // text_format.text_dbnum1
    false,       // locale_tokens.text_letter_t_digits
    '\0',        // locale_tokens.text_letter_x_positive, text_letter_x_negative, text_slash_letters
    false,       // unmeasured
    false,       // text_general_english
    false,       // text_decimal_point
    "Standard",  // text_general_standard
    {"", "", "", "", "", "Rot", "", ""},        // text_color_rot; other colours unmeasured
    DbnumDigits{{kAsciiDigits, kAsciiDigits}},  // text_format.text_dbnum1
    "Gesamt",                                   // groupby.groupby_measured_fh2_hdr
    "Gesamtergebnis",                           // pivotby.pivotby_row_subtotal_depth_two
    "Zeilenfeld ",                              // groupby.groupby_measured_fh2_hdr
    "Spaltenfeld ",                             // pivotby.pivotby_measured_fh2_hdr
    "Wert ",                                    // groupby.groupby_measured_fh2_hdr
    {"Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"},  // weekdays_aaaa
    {"So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"},                                         // weekdays_aaa
};

constexpr LocaleFacts kFrenchFacts{
    ',',                 // fixed_negative
    ' ',                 // fixed_negative
    ';',                 // formulatext_bool_literal
    '.',                 // formulatext_array_constant
    ';',                 // formulatext_array_constant
    "VRAI",              // bool_text_true
    "FAUX",              // bool_text_false
    true,                // text_to_bool_probes.text_bool_and_whitespace, text_bool_and_two_true
    kEnglishErrorNames,  // arraytotext.arraytotext_error_literal_in_array_default; others unmeasured
    'L',
    'C',
    '(',
    ')',  // references.address_r1c1_row_abs_col_rel
    'i',  // cell_type_blank
    'l',  // cell_type_text
    'v',  // cell_type_number
    'S',  // cell.cell_format_general
    {"janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre",
     "décembre"},                                                                                 // months_mmmm
    {"janv", "févr", "mars", "avr", "mai", "juin", "juil", "août", "sept", "oct", "nov", "déc"},  // months_mmm
    {"dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"},                    // weekdays_jjjj
    {"dim", "lun", "mar", "mer", "jeu", "ven", "sam"},                                            // weekdays_jjj
    // months_aaaa, months_mmmm, months_jj, text_format.text_datetime_combined,
    // text_format.text_time_hh_mm_ss, weekdays_aaaa (aaaa is a year)
    {'a', 'm', 'j', 'h', 'm', 's', false, false, false, '\0'},
    kNoLetterAliases,
    {"€", true, true, true, false, true, 2U},  // text.dollar_negative_minus_sign, text.dollar_rounds_to_negative_zero
    {"€", "", "", ""},                         // value_euro_prefix, value_dollar_prefix
    "",                                        // usdollar.usdollar_two_decimals
    DateOrder::kDMY,                           // value_coercion_probes.value_slash_date_short
    false,                                     // value_dotted_dmy
    false,                                     // datevalue_timevalue.datevalue_kanji_with_terminator
    false,                                     // datevalue_timevalue.timevalue_jp_kanji_units
    false,                                     // datevalue_timevalue.datevalue_era_reiwa_full
    RLetter::kLiteral,                         // locale_tokens.text_letter_r_date, text_letter_rr_date
    false,                                     // locale_tokens.value_korean_ymd
    false,                                     // value_month_name_en
    true,                                      // value_coercion_probes.value_d_mmm_yy
    true,                 // datevalue_timevalue.timevalue_fractional_seconds, timevalue_comma_fraction
    false,                // datevalue_timevalue.timevalue_trailing_dot
    false,                // datevalue_timevalue.timevalue_space_dot
    "AM",                 // text_format.text_time_am
    "PM",                 // text_format.text_time_pm
    true,                 // text_format.text_time_a_p
    DbcsCodepage::kNone,  // lenb_hangul
    false,
    false,       // value_numbervalue.value_fullwidth_digits
    true,        // code_char_jp_probes.near_char
    false,       // lazy_forms.lazy_phonetic_unannotated_cell
    false,       // dfunc_kana_folding_probes.dsum_criteria_header_halfwidth_vs_fullwidth_db_header
    false,       // text_format.text_four_section_text
    false,       // text_format.text_dbnum1
    false,       // locale_tokens.text_letter_t_digits
    '\0',        // locale_tokens.text_letter_x_positive, text_letter_x_negative, text_slash_letters
    false,       // unmeasured
    false,       // text_general_english
    false,       // text_decimal_point
    "Standard",  // text_general_standard
    {"", "", "", "", "", "Rouge", "", ""},      // text_color_rouge; other colours unmeasured
    DbnumDigits{{kAsciiDigits, kAsciiDigits}},  // text_format.text_dbnum1
    "Total",                                    // groupby.groupby_measured_fh2_hdr
    "Total général",                            // pivotby.pivotby_row_subtotal_depth_two
    "Champ de ligne ",                          // groupby.groupby_measured_fh2_hdr
    "Champ de colonne ",                        // pivotby.pivotby_measured_fh2_hdr
    "Valeur ",                                  // groupby.groupby_measured_fh2_hdr
    kEnglishDaysLong,                           // unmeasured: aaaa is a year token (weekdays_aaaa)
    kEnglishDaysShort,                          // unmeasured: aaa is a year token (weekdays_aaa)
};

constexpr LocaleFacts kChineseFacts{
    '.',                 // fixed_negative
    ',',                 // fixed_negative
    ',',                 // formulatext_bool_literal
    ',',                 // formulatext_array_constant
    ';',                 // formulatext_array_constant
    "TRUE",              // bool_text_true
    "FALSE",             // bool_text_false
    false,               // text_to_bool_probes.text_bool_and_whitespace, text_bool_and_two_true
    kEnglishErrorNames,  // arraytotext.arraytotext_error_literal_in_array_default; others unmeasured
    'R',
    'C',
    '[',
    ']',                  // references.address_r1c1_row_abs_col_rel
    'b',                  // cell_type_blank
    'l',                  // cell_type_text
    'v',                  // cell_type_number
    'G',                  // cell.cell_format_general
    kEnglishMonthsLong,   // months_mmmm
    kEnglishMonthsShort,  // months_mmm
    kEnglishDaysLong,     // weekdays_dddd
    kEnglishDaysShort,    // weekdays_ddd
    kInvariantLetters,    // months_yyyy, months_d, weekdays_aaaa
    kNoLetterAliases,
    {"¥", false, false, true, false, true, 2U},  // text.dollar_negative_minus_sign, text.dollar_rounds_to_negative_zero
    {"$", "€", "¥", ""},    // value_dollar_prefix, value_euro_prefix, value_yen_prefix; ￥ unmeasured
    "$",                    // usdollar.usdollar_two_decimals
    DateOrder::kYMD,        // locale_profile_measurements.datevalue_two_digit_year
    false,                  // value_dotted_ymd
    true,                   // datevalue_timevalue.datevalue_kanji_with_terminator
    false,                  // datevalue_timevalue.timevalue_jp_kanji_units
    false,                  // datevalue_timevalue.datevalue_era_reiwa_full
    RLetter::kLiteral,      // locale_tokens.text_letter_r_date, text_letter_rr_date
    false,                  // locale_tokens.value_korean_ymd
    true,                   // value_month_name_en
    true,                   // value_coercion_probes.value_d_mmm_yy
    true,                   // datevalue_timevalue.timevalue_fractional_seconds, timevalue_comma_fraction
    true,                   // datevalue_timevalue.timevalue_trailing_dot
    true,                   // datevalue_timevalue.timevalue_space_dot
    "AM",                   // text_format.text_time_am
    "PM",                   // text_format.text_time_pm
    true,                   // text_format.text_time_a_p
    DbcsCodepage::kGb2312,  // lenb_kanji_not_in_gb2312
    false,                  // code_char_jp_probes.char_halfwidth_kata_177
    true,                   // value_numbervalue.value_fullwidth_digits
    false,                  // code_char_jp_probes.near_char
    true,                   // lazy_forms.lazy_phonetic_unannotated_cell
    false,         // no width folding (dfunc_kana_folding_probes.dsum_criteria_header_halfwidth_vs_fullwidth_db_header)
    true,          // text_format.text_four_section_text
    true,          // text_format.text_dbnum1
    false,         // locale_tokens.text_letter_t_digits
    '\0',          // locale_tokens.text_letter_x_positive, text_letter_x_negative, text_slash_letters
    false,         // unmeasured
    false,         // text_general_english
    false,         // text_decimal_point
    "G/通用格式",  // text_general_g_tongyong
    {"", "", "", "", "", "红色", "", ""},  // text_color_hongse; other colours unmeasured
    DbnumDigits{{
        {"○", "一", "二", "三", "四", "5", "6", "7", "8", "9"},  // 5-9 unmeasured
        {"0", "壹", "贰", "叁", "肆", "5", "6", "7", "8", "9"},  // 0 and 5-9 unmeasured
    }},         // text_format.text_dbnum1_with_era, text_format.text_dbnum1, text_format.text_dbnum2
    "总计",     // groupby.groupby_measured_fh2_hdr
    "总计",     // pivotby.pivotby_row_subtotal_depth_two
    "行字段 ",  // groupby.groupby_measured_fh2_hdr
    "列字段 ",  // pivotby.pivotby_measured_fh2_hdr
    "值 ",      // groupby.groupby_measured_fh2_hdr
    {"星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"},  // weekdays_aaaa
    {"周日", "周一", "周二", "周三", "周四", "周五", "周六"},                // weekdays_aaa
};

constexpr LocaleFacts kKoreanFacts{
    '.',                 // fixed_negative
    ',',                 // fixed_negative
    ',',                 // formulatext_bool_literal
    ',',                 // formulatext_array_constant
    ';',                 // formulatext_array_constant
    "TRUE",              // bool_text_true
    "FALSE",             // bool_text_false
    false,               // text_to_bool_probes.text_bool_and_whitespace, text_bool_and_two_true
    kEnglishErrorNames,  // arraytotext.arraytotext_error_literal_in_array_default; others unmeasured
    'R',
    'C',
    '[',
    ']',                  // references.address_r1c1_row_abs_col_rel
    'b',                  // cell_type_blank
    'l',                  // cell_type_text
    'v',                  // cell_type_number
    'G',                  // cell.cell_format_general
    kEnglishMonthsLong,   // months_mmmm
    kEnglishMonthsShort,  // months_mmm
    kEnglishDaysLong,     // weekdays_dddd
    kEnglishDaysShort,    // weekdays_ddd
    kInvariantLetters,    // months_yyyy, months_d, weekdays_aaaa
    kNoLetterAliases,
    {"₩", false, false, true, false, true, 0U},  // text.dollar_negative_minus_sign, text.dollar_zero
    {"$", "€", "₩", ""},                         // value_dollar_prefix, value_euro_prefix, value_won_prefix
    "$",                                         // usdollar.usdollar_two_decimals
    DateOrder::kYMD,                             // locale_profile_measurements.datevalue_two_digit_year
    true,                                        // value_dotted_ymd
    false,                                       // datevalue_timevalue.datevalue_kanji_with_terminator
    false,                                       // datevalue_timevalue.timevalue_jp_kanji_units
    false,                                       // datevalue_timevalue.datevalue_era_reiwa_full
    RLetter::kLiteral,                           // locale_tokens.text_letter_r_date, text_letter_rr_date
    true,                                        // locale_tokens.value_korean_ymd
    true,                                        // value_month_name_en
    true,                                        // value_coercion_probes.value_d_mmm_yy
    false,                   // datevalue_timevalue.timevalue_fractional_seconds, timevalue_comma_fraction
    true,                    // datevalue_timevalue.timevalue_trailing_dot
    true,                    // datevalue_timevalue.timevalue_space_dot
    "AM",                    // text_format.text_time_am
    "PM",                    // text_format.text_time_pm
    true,                    // text_format.text_time_a_p
    DbcsCodepage::kKsX1001,  // code_hangul
    false,                   // code_char_jp_probes.char_halfwidth_kata_177
    true,                    // value_numbervalue.value_fullwidth_digits
    false,                   // code_char_jp_probes.near_char
    true,                    // lazy_forms.lazy_phonetic_unannotated_cell
    false,                   // dfunc_kana_folding_probes.dsum_criteria_header_halfwidth_vs_fullwidth_db_header
    true,                    // text_format.text_four_section_text
    true,                    // text_format.text_dbnum1
    false,                   // locale_tokens.text_letter_t_digits
    '\0',                    // locale_tokens.text_letter_x_positive, text_letter_x_negative, text_slash_letters
    false,                   // unmeasured
    false,                   // text_general_english
    false,                   // text_decimal_point
    "G/표준",                // text_general_g_pyojun
    {"", "", "", "", "", "빨강", "", ""},  // text_color_ppalgang; other colours unmeasured
    DbnumDigits{{
        {"０", "一", "二", "三", "四", "5", "6", "7", "8", "9"},           // 5-9 unmeasured
        {"0", "壹", "貳", "\xEF\xA5\xAB", "四", "5", "6", "7", "8", "9"},  // 3 is U+F96B; 0 and 5-9 unmeasured
    }},          // text_format.text_dbnum1_with_era, text_format.text_dbnum1, text_format.text_dbnum2
    "합계",      // groupby.groupby_measured_fh2_hdr
    "총합계",    // pivotby.pivotby_row_subtotal_depth_two
    "행 필드 ",  // groupby.groupby_measured_fh2_hdr
    "열 필드 ",  // pivotby.pivotby_measured_fh2_hdr
    "값 ",       // groupby.groupby_measured_fh2_hdr
    {"일요일", "월요일", "화요일", "수요일", "목요일", "금요일", "토요일"},  // weekdays_aaaa
    {"일", "월", "화", "수", "목", "금", "토"},                              // weekdays_aaa
};

constexpr LocaleFacts kThaiFacts{
    '.',                 // fixed_negative
    ',',                 // fixed_negative
    ',',                 // formulatext_bool_literal
    ',',                 // formulatext_array_constant
    ';',                 // formulatext_array_constant
    "TRUE",              // bool_text_true
    "FALSE",             // bool_text_false
    false,               // text_to_bool_probes.text_bool_and_whitespace, text_bool_and_two_true
    kEnglishErrorNames,  // arraytotext.arraytotext_error_literal_in_array_default; others unmeasured
    'R',
    'C',
    '[',
    ']',  // references.address_r1c1_row_abs_col_rel
    'b',  // cell_type_blank
    'l',  // cell_type_text
    'v',  // cell_type_number
    'G',  // cell.cell_format_general
    {"มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน",
     "ธันวาคม"},                                                                                         // months_mmmm
    {"ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."},  // months_mmm
    {"วันอาทิตย์", "วันจันทร์", "วันอังคาร", "วันพุธ", "วันพฤหัสบดี", "วันศุกร์", "วันเสาร์"},                            // weekdays_dddd
    {"อาทิตย์", "จันทร์", "อังคาร", "พุธ", "พฤหัส", "ศุกร์", "เสาร์"},                                            // weekdays_ddd
    kInvariantLetters,  // months_yyyy, months_d, weekdays_aaaa
    kNoLetterAliases,
    {"฿", false, false, true, false, true, 2U},  // text.dollar_negative_minus_sign, text.dollar_rounds_to_negative_zero
    {"€", "฿", "", ""},                          // value_euro_prefix, value_baht_prefix, value_dollar_prefix
    "",                                          // usdollar.usdollar_two_decimals
    DateOrder::kDMY,                             // value_coercion_probes.value_slash_date_short
    false,                                       // value_dotted_dmy
    false,                                       // datevalue_timevalue.datevalue_kanji_with_terminator
    false,                                       // datevalue_timevalue.timevalue_jp_kanji_units
    false,                                       // datevalue_timevalue.datevalue_era_reiwa_full
    RLetter::kLiteral,                           // locale_tokens.text_letter_r_date, text_letter_rr_date
    false,                                       // locale_tokens.value_korean_ymd
    true,                                        // value_month_name_en
    true,                                        // value_coercion_probes.value_d_mmm_yy
    true,                 // datevalue_timevalue.timevalue_fractional_seconds, timevalue_comma_fraction
    false,                // datevalue_timevalue.timevalue_trailing_dot
    true,                 // datevalue_timevalue.timevalue_space_dot
    "AM",                 // text_format.text_time_am
    "PM",                 // text_format.text_time_pm
    true,                 // text_format.text_time_a_p
    DbcsCodepage::kNone,  // lenb_hangul
    false,
    false,  // value_numbervalue.value_fullwidth_digits
    true,   // code_char_jp_probes.near_char
    false,  // lazy_forms.lazy_phonetic_unannotated_cell
    false,  // dfunc_kana_folding_probes.dsum_criteria_header_halfwidth_vs_fullwidth_db_header
    false,  // text_format.text_four_section_text
    false,  // text_format.text_dbnum1
    true,   // locale_tokens.text_letter_t_digits
    '\0',   // locale_tokens.text_letter_x_positive, text_letter_x_negative, text_slash_letters
    false,  // unmeasured
    true,   // text_general_english
    false,  // text_decimal_point
    "",     // text_general_english
    {"", "", "", "", "", "", "", ""},           // text_color_red, text_color_thai_red; other colours unmeasured
    DbnumDigits{{kAsciiDigits, kAsciiDigits}},  // text_format.text_dbnum1
    "ผลรวม",                                    // groupby.groupby_measured_fh2_hdr
    "ผลรวมทั้งหมด",                               // pivotby.pivotby_row_subtotal_depth_two
    "เขตข้อมูลแถว ",                              // groupby.groupby_measured_fh2_hdr
    "เขตข้อมูลคอลัมน์ ",                            // pivotby.pivotby_measured_fh2_hdr
    "ค่า ",                                      // groupby.groupby_measured_fh2_hdr
    {"วันอาทิตย์", "วันจันทร์", "วันอังคาร", "วันพุธ", "วันพฤหัสบดี", "วันศุกร์", "วันเสาร์"},  // weekdays_aaaa
    {"อาทิตย์", "จันทร์", "อังคาร", "พุธ", "พฤหัส", "ศุกร์", "เสาร์"},                  // weekdays_aaa
};

// ru-RU: #N/A and #DIV/0! are localized by captures
// (arraytotext.arraytotext_error_literal_in_array_default, arraytotext.arraytotext_only_error_cells_default);
// every other slot is unmeasured.
constexpr ErrorNames kRussianErrorNames{
    "#NULL!", "#ДЕЛ/0!", "#VALUE!",   "#REF!",     "#NAME?",     "#NUM!",  "#Н/Д",     "#GETTING_DATA", "#SPILL!",
    "#CALC!", "#FIELD!", "#BLOCKED!", "#CONNECT!", "#EXTERNAL!", "#BUSY!", "#PYTHON!", "#UNKNOWN!",
};

constexpr LocaleFacts kRussianFacts{
    ',',                 // fixed_negative
    ' ',                 // fixed_negative
    ';',                 // formulatext_bool_literal
    ';',                 // formulatext_array_constant
    ':',                 // formulatext_array_constant
    "ИСТИНА",            // bool_text_true
    "ЛОЖЬ",              // bool_text_false
    true,                // text_to_bool_probes.text_bool_and_whitespace, text_bool_and_two_true
    kRussianErrorNames,  // arraytotext.arraytotext_error_literal_in_array_default,
                         // arraytotext.arraytotext_only_error_cells_default; others unmeasured
    'R',
    'C',
    '[',
    ']',  // references.address_r1c1_row_abs_col_rel
    'b',  // cell_type_blank
    'l',  // cell_type_text
    'v',  // cell_type_number
    'G',  // cell.cell_format_general
    {"январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август", "сентябрь", "октябрь", "ноябрь",
     "декабрь"},                                                                                 // months_cyr_mmmm
    {"янв", "февр", "март", "апр", "май", "июнь", "июль", "авг", "сент", "окт", "нояб", "дек"},  // months_cyr_mmm
    {"воскресенье", "понедельник", "вторник", "среда", "четверг", "пятница", "суббота"},         // weekdays_cyr_dddd
    {"Вс", "Пн", "Вт", "Ср", "Чт", "Пт", "Сб"},                                                  // weekdays_cyr_ddd
    // date_cyr_d_mm_yyyy, date_cyr_lower, time_cyr_lower, time_cyr_upper, weekdays_aaaa; the Latin letters
    // are plain text (months_yyyy, months_d, text_format.text_time_hh_mm_ss)
    {'y', 'M', 'd', 'h', 'm', 's', true, true, true, 'a'},
    kRussianLetterAliases,
    {"₽", true, true, false, false, true,
     2U},                 // text.dollar_negative_minus_sign, usdollar.dollar_rounds_to_negative_zero
    {"€", "", "", ""},    // value_euro_prefix, value_dollar_prefix, value_yen_prefix
    "",                   // usdollar.usdollar_two_decimals
    DateOrder::kDMY,      // value_coercion_probes.value_slash_date_short
    true,                 // value_dotted_dmy, value_dotted_ymd
    false,                // datevalue_timevalue.datevalue_kanji_with_terminator
    false,                // datevalue_timevalue.timevalue_jp_kanji_units
    false,                // datevalue_timevalue.datevalue_era_reiwa_full
    RLetter::kLiteral,    // locale_tokens.text_letter_r_date, text_letter_rr_date
    false,                // locale_tokens.value_korean_ymd
    false,                // value_month_name_en
    false,                // value_coercion_probes.value_d_mmm_yy
    true,                 // timevalue_comma_fraction
    true,                 // datevalue_timevalue.timevalue_trailing_dot
    true,                 // datevalue_timevalue.timevalue_space_dot
    "AM",                 // text_format.text_time_am
    "PM",                 // text_format.text_time_pm
    true,                 // text_format.text_time_a_p
    DbcsCodepage::kNone,  // lenb_hangul
    false,                // code_char_jp_probes.char_halfwidth_kata_177
    false,                // value_numbervalue.value_fullwidth_digits
    true,                 // code_char_jp_probes.near_char
    false,                // lazy_forms.lazy_phonetic_unannotated_cell
    false,                // dfunc_kana_folding_probes.dsum_criteria_header_halfwidth_vs_fullwidth_db_header
    false,                // text_format.text_four_section_text
    false,                // text_format.text_dbnum1
    false,                // locale_tokens.text_letter_t_digits
    '\0',                 // locale_tokens.text_letter_x_positive, text_letter_x_negative, text_slash_letters
    false,                // unmeasured
    false,                // text_general_english
    true,                 // text_decimal_point
    "Основной",           // text_general_osnovnoy; General and Standard are rejected (text_general_english,
                          // text_general_standard)
    {"", "", "", "", "", "Красный", "",
     ""},  // text_color_krasny; English Red is rejected (text_color_red), others unmeasured
    DbnumDigits{{kAsciiDigits, kAsciiDigits}},  // text_format.text_dbnum1
    "Итого",                                    // groupby.groupby_measured_fh2_hdr
    "Общий итог",                               // pivotby.pivotby_row_subtotal_depth_two
    "Поле строки ",                             // groupby.groupby_measured_fh2_hdr
    "Поле столбца ",                            // pivotby.pivotby_measured_fh2_hdr
    "Значение ",                                // groupby.groupby_measured_fh2_hdr
    {"воскресенье", "понедельник", "вторник", "среда", "четверг", "пятница", "суббота"},  // weekdays_aaaa
    {"Вс", "Пн", "Вт", "Ср", "Чт", "Пт", "Сб"},                                           // weekdays_aaa
};

constexpr LocaleFacts kTraditionalChineseFacts{
    '.',                 // fixed_negative
    ',',                 // fixed_negative
    ',',                 // formulatext_bool_literal
    ',',                 // formulatext_array_constant
    ';',                 // formulatext_array_constant
    "TRUE",              // bool_text_true
    "FALSE",             // bool_text_false
    false,               // text_to_bool_probes.text_bool_and_whitespace, text_bool_and_two_true
    kEnglishErrorNames,  // arraytotext.arraytotext_error_literal_in_array_default; others unmeasured
    'R',
    'C',
    '[',
    ']',                  // references.address_r1c1_row_abs_col_rel
    'b',                  // cell_type_blank
    'l',                  // cell_type_text
    'v',                  // cell_type_number
    'G',                  // cell.cell_format_general
    kEnglishMonthsLong,   // months_mmmm
    kEnglishMonthsShort,  // months_mmm
    kEnglishDaysLong,     // weekdays_dddd
    kEnglishDaysShort,    // weekdays_ddd
    kInvariantLetters,    // months_yyyy, months_d, weekdays_aaaa
    kNoLetterAliases,
    {"$", false, false, true, false, true,
     2U},                 // text.dollar_negative_minus_sign, usdollar.dollar_rounds_to_negative_zero
    {"$", "€", "", ""},   // value_dollar_prefix, value_euro_prefix, value_yen_prefix
    "US$",                // usdollar.usdollar_two_decimals
    DateOrder::kYMD,      // locale_profile_measurements.datevalue_two_digit_year
    false,                // value_dotted_ymd
    true,                 // datevalue_timevalue.datevalue_kanji_with_terminator
    true,                 // datevalue_timevalue.timevalue_jp_kanji_units
    false,                // datevalue_timevalue.datevalue_era_reiwa_full
    RLetter::kYear,       // locale_tokens.text_letter_r_date, text_letter_rr_date
    false,                // locale_tokens.value_korean_ymd
    true,                 // value_month_name_en
    true,                 // value_coercion_probes.value_d_mmm_yy
    true,                 // datevalue_timevalue.timevalue_fractional_seconds
    true,                 // datevalue_timevalue.timevalue_trailing_dot
    true,                 // datevalue_timevalue.timevalue_space_dot
    "AM",                 // text_format.text_time_am
    "PM",                 // text_format.text_time_pm
    true,                 // text_format.text_time_a_p
    DbcsCodepage::kBig5,  // code_char_jp_probes (Big5, no ETEN rows), lenb_kanji_not_in_gb2312
    false,                // code_char_jp_probes.char_halfwidth_kata_177
    true,                 // value_numbervalue.value_fullwidth_digits
    false,                // code_char_jp_probes.near_char
    true,                 // lazy_forms.lazy_phonetic_unannotated_cell
    false,                // dfunc_kana_folding_probes.dsum_criteria_header_halfwidth_vs_fullwidth_db_header
    true,                 // text_format.text_four_section_text
    true,                 // text_format.text_dbnum1
    false,                // locale_tokens.text_letter_t_digits
    '\0',                 // locale_tokens.text_letter_x_positive, text_letter_x_negative, text_slash_letters
    false,                // unmeasured
    false,                // text_general_english
    false,                // text_decimal_point
    "G/通用格式",         // text_general_g_tongyong
    {"", "", "", "", "", "紅色", "",
     ""},  // text_color_hongse_traditional; English Red is rejected (text_color_red), others unmeasured
    DbnumDigits{{
        {"○", "一", "二", "三", "四", "5", "6", "7", "8",
         "9"},  // 0 from text_format.text_dbnum1_with_era; 5-9 unmeasured
        {"0", "壹", "貳", "參", "肆", "5", "6", "7", "8", "9"},  // text_format.text_dbnum2; 0 and 5-9 unmeasured
    }},         // text_format.text_dbnum1_with_era, text_format.text_dbnum1, text_format.text_dbnum2
    "總計",     // groupby.groupby_measured_fh2_hdr
    "總計",     // pivotby.pivotby_row_subtotal_depth_two
    "列欄位 ",  // groupby.groupby_measured_fh2_hdr
    "欄欄位 ",  // pivotby.pivotby_measured_fh2_hdr
    "值 ",      // groupby.groupby_measured_fh2_hdr
    {"星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"},  // weekdays_aaaa
    {"週日", "週一", "週二", "週三", "週四", "週五", "週六"},                // weekdays_aaa
};

// it-IT: #N/A is localized by captures
// (arraytotext.arraytotext_error_literal_in_array_default); #DIV/0! is measured unchanged
// (arraytotext.arraytotext_only_error_cells_default); every other slot is unmeasured.
constexpr ErrorNames kItalianErrorNames{
    "#NULL!", "#DIV/0!", "#VALUE!",   "#REF!",     "#NAME?",     "#NUM!",  "#N/D",     "#GETTING_DATA", "#SPILL!",
    "#CALC!", "#FIELD!", "#BLOCKED!", "#CONNECT!", "#EXTERNAL!", "#BUSY!", "#PYTHON!", "#UNKNOWN!",
};

constexpr LocaleFacts kItalianFacts{
    ',',                 // fixed_negative
    '.',                 // fixed_negative
    ';',                 // formulatext_bool_literal
    '\\',                // formulatext_array_constant
    '.',                 // formulatext_array_constant
    "VERO",              // bool_text_true
    "FALSO",             // bool_text_false
    true,                // text_to_bool_probes.text_bool_and_whitespace, text_bool_and_two_true
    kItalianErrorNames,  // arraytotext.arraytotext_error_literal_in_array_default; others unmeasured
    'R',
    'C',
    '[',
    ']',  // references.address_r1c1_row_abs_col_rel
    'b',  // cell_type_blank
    'l',  // cell_type_text
    'v',  // cell_type_number
    'G',  // cell.cell_format_general
    {"gennaio", "febbraio", "marzo", "aprile", "maggio", "giugno", "luglio", "agosto", "settembre", "ottobre",
     "novembre", "dicembre"},                                                              // months_mmmm
    {"gen", "feb", "mar", "apr", "mag", "giu", "lug", "ago", "set", "ott", "nov", "dic"},  // months_mmm
    {"domenica", "lunedì", "martedì", "mercoledì", "giovedì", "venerdì", "sabato"},        // weekdays_gggg
    {"dom", "lun", "mar", "mer", "gio", "ven", "sab"},                                     // weekdays_ggg
    // months_aaaa, months_mmmm, text_format.text_era_gge_reiwa (g is the day letter),
    // text_format.text_datetime_combined, text_format.text_time_hh_mm_ss; aaaa is a year (weekdays_aaaa)
    {'a', 'm', 'g', 'h', 'm', 's', false, false, false, '\0'},
    kNoLetterAliases,
    {"€", true, true, false, false, true,
     2U},                 // text.dollar_negative_minus_sign, usdollar.dollar_rounds_to_negative_zero
    {"€", "", "", ""},    // value_euro_prefix, value_dollar_prefix, value_yen_prefix
    "",                   // usdollar.usdollar_two_decimals
    DateOrder::kDMY,      // value_coercion_probes.value_slash_date_short
    false,                // value_dotted_dmy, value_dotted_ymd
    false,                // datevalue_timevalue.datevalue_kanji_with_terminator
    false,                // datevalue_timevalue.timevalue_jp_kanji_units
    false,                // datevalue_timevalue.datevalue_era_reiwa_full
    RLetter::kLiteral,    // locale_tokens.text_letter_r_date, text_letter_rr_date
    false,                // locale_tokens.value_korean_ymd
    false,                // value_month_name_en
    false,                // value_coercion_probes.value_d_mmm_yy
    true,                 // timevalue_comma_fraction
    false,                // datevalue_timevalue.timevalue_trailing_dot
    false,                // datevalue_timevalue.timevalue_space_dot
    "AM",                 // text_format.text_time_am
    "PM",                 // text_format.text_time_pm
    true,                 // text_format.text_time_a_p
    DbcsCodepage::kNone,  // lenb_hangul
    false,                // code_char_jp_probes.char_halfwidth_kata_177
    false,                // value_numbervalue.value_fullwidth_digits
    true,                 // code_char_jp_probes.near_char
    false,                // lazy_forms.lazy_phonetic_unannotated_cell
    false,                // dfunc_kana_folding_probes.dsum_criteria_header_halfwidth_vs_fullwidth_db_header
    false,                // text_format.text_four_section_text
    false,                // text_format.text_dbnum1
    false,                // locale_tokens.text_letter_t_digits
    'x',                  // locale_tokens.text_letter_x_positive, text_letter_x_negative, text_slash_letters
    false,                // unmeasured
    false,                // text_general_english
    false,                // text_decimal_point
    "Standard",           // text_general_standard
    {"", "", "", "", "", "Rosso", "",
     ""},  // text_color_rosso; English Red is rejected (text_color_red), others unmeasured
    DbnumDigits{{kAsciiDigits, kAsciiDigits}},  // text_format.text_dbnum1
    "Totale",                                   // groupby.groupby_measured_fh2_hdr
    "Totale complessivo",                       // pivotby.pivotby_row_subtotal_depth_two
    "Campo riga ",                              // groupby.groupby_measured_fh2_hdr
    "Campo colonna ",                           // pivotby.pivotby_measured_fh2_hdr
    "Valore ",                                  // groupby.groupby_measured_fh2_hdr
    kEnglishDaysLong,                           // unmeasured: aaaa is a year token (weekdays_aaaa)
    kEnglishDaysShort,                          // unmeasured: aaa is a year token (weekdays_aaa)
};

// nl-NL: #N/A and #DIV/0! are localized by captures
// (arraytotext.arraytotext_error_literal_in_array_default, arraytotext.arraytotext_only_error_cells_default);
// every other slot is unmeasured.
constexpr ErrorNames kDutchErrorNames{
    "#NULL!",    "#DELING.DOOR.0!", "#VALUE!", "#REF!",    "#NAME?",    "#NUM!",
    "#N/B",      "#GETTING_DATA",   "#SPILL!", "#CALC!",   "#FIELD!",   "#BLOCKED!",
    "#CONNECT!", "#EXTERNAL!",      "#BUSY!",  "#PYTHON!", "#UNKNOWN!",
};

constexpr LocaleFacts kDutchFacts{
    ',',               // fixed_negative
    '.',               // fixed_negative
    ';',               // formulatext_bool_literal
    '\\',              // formulatext_array_constant
    ';',               // formulatext_array_constant
    "WAAR",            // bool_text_true
    "ONWAAR",          // bool_text_false
    true,              // text_to_bool_probes.text_bool_and_whitespace, text_bool_and_two_true
    kDutchErrorNames,  // arraytotext.arraytotext_error_literal_in_array_default,
                       // arraytotext.arraytotext_only_error_cells_default; others unmeasured
    'R',
    'K',
    '[',
    ']',  // references.address_r1c1_row_abs_col_rel
    'g',  // cell_type_blank
    'l',  // cell_type_text
    'w',  // cell_type_number
    'S',  // cell.cell_format_general
    {"januari", "februari", "maart", "april", "mei", "juni", "juli", "augustus", "september", "oktober", "november",
     "december"},                                                                          // months_mmmm
    {"jan", "feb", "mrt", "apr", "mei", "jun", "jul", "aug", "sep", "okt", "nov", "dec"},  // months_mmm
    {"zondag", "maandag", "dinsdag", "woensdag", "donderdag", "vrijdag", "zaterdag"},      // weekdays_dddd
    {"zo", "ma", "di", "wo", "do", "vr", "za"},                                            // weekdays_ddd
    // months_jjjj, months_mmmm, months_dd, text_format.text_time_hh_mm_ss, weekdays_aaaa, time_u_hours
    {'j', 'm', 'd', 'u', 'm', 's', false, false, false, 'a'},
    kNoLetterAliases,
    {"€", false, true, true, false, true,
     2U},                 // text.dollar_negative_minus_sign, usdollar.dollar_rounds_to_negative_zero
    {"€", "", "", ""},    // value_euro_prefix, value_dollar_prefix, value_yen_prefix
    "",                   // usdollar.usdollar_two_decimals
    DateOrder::kDMY,      // value_coercion_probes.value_slash_date_short
    false,                // value_dotted_dmy, value_dotted_ymd
    false,                // datevalue_timevalue.datevalue_kanji_with_terminator
    false,                // datevalue_timevalue.timevalue_jp_kanji_units
    false,                // datevalue_timevalue.datevalue_era_reiwa_full
    RLetter::kLiteral,    // locale_tokens.text_letter_r_date, text_letter_rr_date
    false,                // locale_tokens.value_korean_ymd
    false,                // value_month_name_en
    true,                 // value_coercion_probes.value_d_mmm_yy
    true,                 // timevalue_comma_fraction
    false,                // datevalue_timevalue.timevalue_trailing_dot
    false,                // datevalue_timevalue.timevalue_space_dot
    "AM",                 // text_format.text_time_am
    "PM",                 // text_format.text_time_pm
    true,                 // text_format.text_time_a_p
    DbcsCodepage::kNone,  // lenb_hangul
    false,                // code_char_jp_probes.char_halfwidth_kata_177
    false,                // value_numbervalue.value_fullwidth_digits
    true,                 // code_char_jp_probes.near_char
    false,                // lazy_forms.lazy_phonetic_unannotated_cell
    false,                // dfunc_kana_folding_probes.dsum_criteria_header_halfwidth_vs_fullwidth_db_header
    false,                // text_format.text_four_section_text
    false,                // text_format.text_dbnum1
    false,                // locale_tokens.text_letter_t_digits
    '\0',                 // locale_tokens.text_letter_x_positive, text_letter_x_negative, text_slash_letters
    false,                // unmeasured
    false,                // text_general_english
    false,                // text_decimal_point
    "Standaard",          // text_general_standaard; General and Standard are rejected (text_general_english,
                          // text_general_standard)
    {"", "", "", "", "", "Rood", "",
     ""},  // text_color_rood; English Red is rejected (text_color_red), others unmeasured
    DbnumDigits{{kAsciiDigits, kAsciiDigits}},  // text_format.text_dbnum1
    "Totaal",                                   // groupby.groupby_measured_fh2_hdr
    "Eindtotaal",                               // pivotby.pivotby_row_subtotal_depth_two
    "Rijveld ",                                 // groupby.groupby_measured_fh2_hdr
    "Kolomveld ",                               // pivotby.pivotby_measured_fh2_hdr
    "Waarde ",                                  // groupby.groupby_measured_fh2_hdr
    {"zondag", "maandag", "dinsdag", "woensdag", "donderdag", "vrijdag", "zaterdag"},  // weekdays_aaaa
    {"zo", "ma", "di", "wo", "do", "vr", "za"},                                        // weekdays_aaa
};

// es-ES: #N/A and #DIV/0! are localized by captures
// (arraytotext.arraytotext_error_literal_in_array_default, arraytotext.arraytotext_only_error_cells_default);
// every other slot is unmeasured.
constexpr ErrorNames kSpanishErrorNames{
    "#NULL!", "#¡DIV/0!", "#VALUE!",   "#REF!",     "#NAME?",     "#NUM!",  "#N/D",     "#GETTING_DATA", "#SPILL!",
    "#CALC!", "#FIELD!",  "#BLOCKED!", "#CONNECT!", "#EXTERNAL!", "#BUSY!", "#PYTHON!", "#UNKNOWN!",
};

// pt-BR and es-MX: #N/A is localized by captures (arraytotext.arraytotext_error_literal_in_array_default);
// #DIV/0! is measured unchanged (arraytotext.arraytotext_only_error_cells_default) under both, so es-MX
// shares this array; every other slot is unmeasured.
constexpr ErrorNames kPortugueseErrorNames{
    "#NULL!", "#DIV/0!", "#VALUE!",   "#REF!",     "#NAME?",     "#NUM!",  "#N/D",     "#GETTING_DATA", "#SPILL!",
    "#CALC!", "#FIELD!", "#BLOCKED!", "#CONNECT!", "#EXTERNAL!", "#BUSY!", "#PYTHON!", "#UNKNOWN!",
};

constexpr Names12 kSpanishMonthsLong{"enero", "febrero", "marzo",      "abril",   "mayo",      "junio",
                                     "julio", "agosto",  "septiembre", "octubre", "noviembre", "diciembre"};
constexpr Names7 kSpanishDaysLong{"domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"};
constexpr Names7 kSpanishDaysShort{"dom", "lun", "mar", "mié", "jue", "vie", "sáb"};

// a is the year letter and aaa / aaaa are year tokens, so there is no weekday letter.
constexpr FormatLetters kIberianLetters{'a', 'm', 'd', 'h', 'm', 's', false, false, false, '\0'};

constexpr LocaleFacts kSpanishFacts{
    ',',                 // fixed_negative
    '.',                 // fixed_negative
    ';',                 // formulatext_bool_literal
    '\\',                // formulatext_array_constant
    ';',                 // formulatext_array_constant
    "VERDADERO",         // bool_text_true
    "FALSO",             // bool_text_false
    true,                // text_to_bool_probes.text_bool_and_whitespace, text_bool_or_whitespace
    kSpanishErrorNames,  // arraytotext.arraytotext_error_literal_in_array_default,
                         // arraytotext.arraytotext_only_error_cells_default; others unmeasured
    'F',
    'C',
    '[',
    ']',                 // references.address_r1c1_row_abs_col_rel
    'b',                 // cell_type_blank
    'r',                 // cell_type_text
    'v',                 // cell_type_number
    'G',                 // cell.cell_format_general
    kSpanishMonthsLong,  // months_mmmm
    {"ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sept", "oct", "nov", "dic"},  // months_mmm
    kSpanishDaysLong,                                                                       // weekdays_dddd
    kSpanishDaysShort,                                                                      // weekdays_ddd
    kIberianLetters,  // months_aaaa, months_mmmm, months_dd, text_format.text_time_hh_mm_ss, weekdays_aaaa
    kNoLetterAliases,
    {"€", true, true, false, false, true,
     2U},                 // text.dollar_negative_minus_sign, usdollar.dollar_rounds_to_negative_zero
    {"€", "", "", ""},    // value_euro_prefix, value_dollar_prefix, value_yen_prefix
    "",                   // usdollar.usdollar_two_decimals
    DateOrder::kDMY,      // value_coercion_probes.value_slash_date_short
    false,                // value_dotted_dmy, value_dotted_ymd
    false,                // datevalue_timevalue.datevalue_kanji_with_terminator
    false,                // datevalue_timevalue.timevalue_jp_kanji_units
    false,                // datevalue_timevalue.datevalue_era_reiwa_full
    RLetter::kLiteral,    // locale_tokens.text_letter_r_date, text_letter_rr_date
    false,                // locale_tokens.value_korean_ymd
    false,                // value_month_name_en
    false,                // value_coercion_probes.value_d_mmm_yy
    true,                 // timevalue_comma_fraction
    false,                // datevalue_timevalue.timevalue_trailing_dot
    false,                // datevalue_timevalue.timevalue_space_dot
    "a.\u202Fm.",         // text_format.text_time_am
    "p.\u202Fm.",         // text_format.text_time_pm
    false,                // text_format.text_time_a_p
    DbcsCodepage::kNone,  // lenb_hangul
    false,                // code_char_jp_probes.char_halfwidth_kata_177
    false,                // value_numbervalue.value_fullwidth_digits
    true,                 // code_char_jp_probes.near_char
    false,                // lazy_forms.lazy_phonetic_unannotated_cell
    false,                // dfunc_kana_folding_probes.dsum_criteria_header_halfwidth_vs_fullwidth_db_header
    false,                // text_format.text_four_section_text
    false,                // text_format.text_dbnum1
    false,                // locale_tokens.text_letter_t_digits
    '\0',                 // locale_tokens.text_letter_x_positive, text_letter_x_negative, text_slash_letters
    false,                // unmeasured
    false,                // text_general_english
    false,                // text_decimal_point
    "Estándar",           // text_general_estandar; General and Standard are rejected (text_general_english,
                          // text_general_standard)
    {"", "", "", "", "", "Rojo", "",
     ""},  // text_color_rojo; English Red is rejected (text_color_red), others unmeasured
    DbnumDigits{{kAsciiDigits, kAsciiDigits}},  // text_format.text_dbnum1
    "Total",                                    // groupby.groupby_measured_fh2_hdr
    "Total general",                            // pivotby.pivotby_row_subtotal_depth_two
    "Campo de fila ",                           // groupby.groupby_measured_fh2_hdr
    "Campo de columna ",                        // pivotby.pivotby_measured_fh2_hdr
    "Valor ",                                   // groupby.groupby_measured_fh2_hdr
    kEnglishDaysLong,                           // unmeasured: aaaa is a year token (weekdays_aaaa)
    kEnglishDaysShort,                          // unmeasured: aaa is a year token (weekdays_aaa)
};

// es-MX keeps the en-US separators and number-format dialect; month and weekday names, bool names, currency
// and date order are its own.
constexpr LocaleFacts kMexicanSpanishFacts{
    '.',                    // fixed_negative
    ',',                    // fixed_negative
    ',',                    // formulatext_bool_literal
    ',',                    // formulatext_array_constant
    ';',                    // formulatext_array_constant
    "VERDADERO",            // bool_text_true
    "FALSO",                // bool_text_false
    true,                   // text_to_bool_probes.text_bool_or_whitespace
    kPortugueseErrorNames,  // arraytotext.arraytotext_error_literal_in_array_default,
                            // arraytotext.arraytotext_only_error_cells_default; others unmeasured
    'F',
    'C',
    '[',
    ']',                 // references.address_r1c1_row_abs_col_rel
    'b',                 // cell_type_blank
    'r',                 // cell_type_text
    'v',                 // cell_type_number
    'G',                 // cell.cell_format_general
    kSpanishMonthsLong,  // months_mmmm
    {"ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"},  // months_mmm
    kSpanishDaysLong,                                                                      // weekdays_dddd
    kSpanishDaysShort,                                                                     // weekdays_ddd
    kIberianLetters,  // months_aaaa, months_mmmm, months_dd, text_format.text_time_hh_mm_ss, weekdays_aaaa
    kNoLetterAliases,
    {"$", false, false, false, false, true,
     2U},                 // text.dollar_negative_minus_sign, usdollar.dollar_rounds_to_negative_zero
    {"$", "€", "", ""},   // value_dollar_prefix, value_euro_prefix, value_yen_prefix
    "",                   // usdollar.usdollar_two_decimals
    DateOrder::kDMY,      // value_coercion_probes.value_slash_date_short
    false,                // value_dotted_dmy, value_dotted_ymd
    false,                // datevalue_timevalue.datevalue_kanji_with_terminator
    false,                // datevalue_timevalue.timevalue_jp_kanji_units
    false,                // datevalue_timevalue.datevalue_era_reiwa_full
    RLetter::kLiteral,    // locale_tokens.text_letter_r_date, text_letter_rr_date
    false,                // locale_tokens.value_korean_ymd
    false,                // value_month_name_en
    false,                // value_coercion_probes.value_d_mmm_yy
    true,                 // datevalue_timevalue.timevalue_fractional_seconds
    true,                 // datevalue_timevalue.timevalue_trailing_dot
    true,                 // datevalue_timevalue.timevalue_space_dot
    "a.m.",               // text_format.text_time_am
    "p.m.",               // text_format.text_time_pm
    false,                // text_format.text_time_a_p
    DbcsCodepage::kNone,  // lenb_hangul
    false,                // code_char_jp_probes.char_halfwidth_kata_177
    false,                // value_numbervalue.value_fullwidth_digits
    true,                 // code_char_jp_probes.near_char
    false,                // lazy_forms.lazy_phonetic_unannotated_cell
    false,                // dfunc_kana_folding_probes.dsum_criteria_header_halfwidth_vs_fullwidth_db_header
    false,                // text_format.text_four_section_text
    false,                // text_format.text_dbnum1
    false,                // locale_tokens.text_letter_t_digits
    '\0',                 // locale_tokens.text_letter_x_positive, text_letter_x_negative, text_slash_letters
    false,                // unmeasured
    false,                // text_general_english
    false,                // text_decimal_point
    "Estándar",           // text_general_estandar; General and Standard are rejected (text_general_english,
                          // text_general_standard)
    {"", "", "", "", "", "Rojo", "",
     ""},  // text_color_rojo; English Red is rejected (text_color_red), others unmeasured
    DbnumDigits{{kAsciiDigits, kAsciiDigits}},  // text_format.text_dbnum1
    "Total",                                    // groupby.groupby_measured_fh2_hdr
    "Total general",                            // pivotby.pivotby_row_subtotal_depth_two
    "Campo de fila ",                           // groupby.groupby_measured_fh2_hdr
    "Campo de columna ",                        // pivotby.pivotby_measured_fh2_hdr
    "Valor ",                                   // groupby.groupby_measured_fh2_hdr
    kEnglishDaysLong,                           // unmeasured: aaaa is a year token (weekdays_aaaa)
    kEnglishDaysShort,                          // unmeasured: aaa is a year token (weekdays_aaa)
};

constexpr LocaleFacts kPortugueseFacts{
    ',',                    // fixed_negative
    '.',                    // fixed_negative
    ';',                    // formulatext_bool_literal
    '\\',                   // formulatext_array_constant
    ';',                    // formulatext_array_constant
    "VERDADEIRO",           // bool_text_true
    "FALSO",                // bool_text_false
    true,                   // text_to_bool_probes.text_bool_or_whitespace
    kPortugueseErrorNames,  // arraytotext.arraytotext_error_literal_in_array_default,
                            // arraytotext.arraytotext_only_error_cells_default; others unmeasured
    'L',
    'C',
    '[',
    ']',  // references.address_r1c1_row_abs_col_rel
    'b',  // cell_type_blank
    'l',  // cell_type_text
    'v',  // cell_type_number
    'G',  // cell.cell_format_general
    {"janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro",
     "dezembro"},                                                                          // months_mmmm
    {"jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"},  // months_mmm
    {"domingo", "segunda-feira", "terça-feira", "quarta-feira", "quinta-feira", "sexta-feira",
     "sábado"},                                         // weekdays_dddd
    {"dom", "seg", "ter", "qua", "qui", "sex", "sáb"},  // weekdays_ddd
    kIberianLetters,  // months_aaaa, months_mmmm, months_dd, text_format.text_time_hh_mm_ss, weekdays_aaaa
    kNoLetterAliases,
    {"R$", false, true, false, false, true,
     2U},                 // text.dollar_negative_minus_sign, usdollar.dollar_rounds_to_negative_zero
    {"€", "", "", ""},    // value_euro_prefix, value_dollar_prefix, value_yen_prefix
    "",                   // usdollar.usdollar_two_decimals
    DateOrder::kDMY,      // value_coercion_probes.value_slash_date_short
    false,                // value_dotted_dmy, value_dotted_ymd
    false,                // datevalue_timevalue.datevalue_kanji_with_terminator
    false,                // datevalue_timevalue.timevalue_jp_kanji_units
    false,                // datevalue_timevalue.datevalue_era_reiwa_full
    RLetter::kLiteral,    // locale_tokens.text_letter_r_date, text_letter_rr_date
    false,                // locale_tokens.value_korean_ymd
    false,                // value_month_name_en
    true,                 // value_coercion_probes.value_d_mmm_yy
    true,                 // timevalue_comma_fraction
    false,                // datevalue_timevalue.timevalue_trailing_dot
    false,                // datevalue_timevalue.timevalue_space_dot
    "AM",                 // text_format.text_time_am
    "PM",                 // text_format.text_time_pm
    true,                 // text_format.text_time_a_p
    DbcsCodepage::kNone,  // lenb_hangul
    false,                // code_char_jp_probes.char_halfwidth_kata_177
    false,                // value_numbervalue.value_fullwidth_digits
    true,                 // code_char_jp_probes.near_char
    false,                // lazy_forms.lazy_phonetic_unannotated_cell
    false,                // dfunc_kana_folding_probes.dsum_criteria_header_halfwidth_vs_fullwidth_db_header
    false,                // text_format.text_four_section_text
    false,                // text_format.text_dbnum1
    false,                // locale_tokens.text_letter_t_digits
    '\0',                 // locale_tokens.text_letter_x_positive, text_letter_x_negative, text_slash_letters
    false,                // unmeasured
    false,                // text_general_english
    false,                // text_decimal_point
    "Geral",  // text_general_geral; General and Standard are rejected (text_general_english, text_general_standard)
    {"", "", "", "", "", "Vermelho", "",
     ""},  // text_color_vermelho; English Red is rejected (text_color_red), others unmeasured
    DbnumDigits{{kAsciiDigits, kAsciiDigits}},  // text_format.text_dbnum1
    "Total",                                    // groupby.groupby_measured_fh2_hdr
    "Total Geral",                              // pivotby.pivotby_row_subtotal_depth_two
    "Campo de linha ",                          // groupby.groupby_measured_fh2_hdr
    "Campo de coluna ",                         // pivotby.pivotby_measured_fh2_hdr
    "Valor ",                                   // groupby.groupby_measured_fh2_hdr
    kEnglishDaysLong,                           // unmeasured: aaaa is a year token (weekdays_aaaa)
    kEnglishDaysShort,                          // unmeasured: aaa is a year token (weekdays_aaa)
};

}  // namespace

const LocaleFacts& locale_facts(ExcelProfile profile) noexcept {
  switch (profile.locale) {
    case ExcelLocale::kJaJP:
      return kJapaneseFacts;
    case ExcelLocale::kEnUS:
      return kEnglishFacts;
    case ExcelLocale::kDeDE:
      return kGermanFacts;
    case ExcelLocale::kFrFR:
      return kFrenchFacts;
    case ExcelLocale::kZhCN:
      return kChineseFacts;
    case ExcelLocale::kKoKR:
      return kKoreanFacts;
    case ExcelLocale::kThTH:
      return kThaiFacts;
    case ExcelLocale::kEsES:
      return kSpanishFacts;
    case ExcelLocale::kEsMX:
      return kMexicanSpanishFacts;
    case ExcelLocale::kPtBR:
      return kPortugueseFacts;
    case ExcelLocale::kRuRU:
      return kRussianFacts;
    case ExcelLocale::kZhTW:
      return kTraditionalChineseFacts;
    case ExcelLocale::kItIT:
      return kItalianFacts;
    case ExcelLocale::kNlNL:
      return kDutchFacts;
  }
  return kEnglishFacts;
}

bool error_name_measured(std::size_t error_ordinal) noexcept {
  // #DIV/0! and #N/A: arraytotext.arraytotext_error_literal_in_array_default and
  // arraytotext.arraytotext_only_error_cells_default run under every Mac locale.
  return error_ordinal == static_cast<std::size_t>(ErrorCode::Div0) ||
         error_ordinal == static_cast<std::size_t>(ErrorCode::NA);
}

SbcsCodepage sbcs_codepage(ExcelProfile profile) noexcept {
  switch (profile.locale) {
    case ExcelLocale::kJaJP:
      return SbcsCodepage::kCp932;
    case ExcelLocale::kZhCN:
    case ExcelLocale::kKoKR:
    case ExcelLocale::kZhTW:
      return SbcsCodepage::kDbcsHighBlank;
    case ExcelLocale::kThTH:
      return SbcsCodepage::kMacThai;
    case ExcelLocale::kRuRU:
      return SbcsCodepage::kMacCyrillic;
    case ExcelLocale::kEnUS:
    case ExcelLocale::kDeDE:
    case ExcelLocale::kFrFR:
    case ExcelLocale::kEsES:
    case ExcelLocale::kEsMX:
    case ExcelLocale::kPtBR:
    case ExcelLocale::kItIT:
    case ExcelLocale::kNlNL:
      break;
  }
  return profile.host == ExcelHost::kMac365 ? SbcsCodepage::kMacRoman : SbcsCodepage::kWindows1252;
}

WidthFolding width_folding(ExcelProfile profile) noexcept {
  if (profile.locale == ExcelLocale::kZhCN || profile.locale == ExcelLocale::kZhTW) {
    return WidthFolding::kNone;
  }
  return profile.host == ExcelHost::kMac365 ? WidthFolding::kMac : WidthFolding::kWin;
}

}  // namespace formulon
