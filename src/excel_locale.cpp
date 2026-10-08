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

constexpr FormatLetters kInvariantLetters{'y', 'm', 'd', 'h', 'm', 's', false, false, 'a'};

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
    ']',                                   // references.address_r1c1_row_abs_col_rel
    'b',                                   // cell_type_blank
    'l',                                   // cell_type_text
    'v',                                   // cell_type_number
    'G',                                   // cell.cell_format_general
    kEnglishMonthsLong,                    // months_mmmm
    kEnglishMonthsShort,                   // months_mmm
    kEnglishDaysLong,                      // weekdays_dddd
    kEnglishDaysShort,                     // weekdays_ddd
    kInvariantLetters,                     // months_yyyy, months_d, weekdays_aaaa
    {"¥", false, false, false, true, 0U},  // text.dollar_negative_minus_sign, text.dollar_rounds_to_negative_zero
    {"$", "€", "¥", "￥"},                 // value_dollar_prefix, value_euro_prefix, value_yen_prefix
    true,                                  // usdollar.usdollar_two_decimals
    DateOrder::kYMD,
    false,  // value_dotted_ymd
    true,   // datevalue_timevalue.datevalue_kanji_with_terminator
    true,   // datevalue_timevalue.timevalue_jp_kanji_units
    true,   // datevalue_timevalue.datevalue_era_reiwa_full
    false,  // locale_tokens.value_korean_ymd
    true,   // value_month_name_en
    true,   // datevalue_timevalue.timevalue_fractional_seconds, timevalue_comma_fraction
    true,   // datevalue_timevalue.timevalue_trailing_dot
    true,   // datevalue_timevalue.timevalue_space_dot
    DbcsCodepage::kJis0208,
    true,      // code_char_jp_probes.char_halfwidth_kata_177
    true,      // value_numbervalue.value_fullwidth_digits
    false,     // code_char_jp_probes.near_char
    true,      // lazy_forms.lazy_phonetic_unannotated_cell
    true,      // dfunc_kana_folding_probes.dsum_criteria_header_halfwidth_vs_fullwidth_db_header
    true,      // text_format.text_four_section_text
    true,      // text_format.text_dbnum1
    true,      // existing ja behaviour; unmeasured
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
    ']',                                  // references.address_r1c1_row_abs_col_rel
    'b',                                  // cell_type_blank
    'l',                                  // cell_type_text
    'v',                                  // cell_type_number
    'G',                                  // cell.cell_format_general
    kEnglishMonthsLong,                   // months_mmmm
    kEnglishMonthsShort,                  // months_mmm
    kEnglishDaysLong,                     // weekdays_dddd
    kEnglishDaysShort,                    // weekdays_ddd
    kInvariantLetters,                    // months_yyyy, months_d, weekdays_aaaa
    {"$", false, false, true, true, 2U},  // text.dollar_negative_minus_sign, text.dollar_rounds_to_negative_zero
    {"$", "€", "", ""},                   // value_dollar_prefix, value_euro_prefix
    true,                                 // usdollar.usdollar_two_decimals
    DateOrder::kMDY,
    false,  // value_dotted_ymd
    false,
    false,
    false,
    false,  // locale_tokens.value_korean_ymd
    true,   // value_month_name_en
    true,   // datevalue_timevalue.timevalue_fractional_seconds, timevalue_comma_fraction
    true,   // datevalue_timevalue.timevalue_trailing_dot
    true,   // datevalue_timevalue.timevalue_space_dot
    DbcsCodepage::kNone,
    false,
    false,
    true,  // code_char_jp_probes.near_char
    false,
    false,
    false,
    false,
    false,
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
    {'J', 'M', 'T', 'h', 'm', 's', true, true, 'a'},
    {"€", true, true, false, true, 2U},  // text.dollar_negative_minus_sign, text.dollar_rounds_to_negative_zero
    {"€", "", "", ""},                   // value_euro_prefix, value_dollar_prefix
    false,                               // usdollar.usdollar_two_decimals
    DateOrder::kDMY,                     // value_coercion_probes.value_slash_date_short
    true,                                // value_dotted_dmy
    false,                               // datevalue_timevalue.datevalue_kanji_with_terminator
    false,                               // datevalue_timevalue.timevalue_jp_kanji_units
    false,                               // datevalue_timevalue.datevalue_era_reiwa_full
    false,                               // locale_tokens.value_korean_ymd
    false,                               // value_month_name_en
    true,                                // datevalue_timevalue.timevalue_fractional_seconds, timevalue_comma_fraction
    true,                                // datevalue_timevalue.timevalue_trailing_dot
    true,                                // datevalue_timevalue.timevalue_space_dot
    DbcsCodepage::kNone,                 // lenb_hangul
    false,
    false,       // value_numbervalue.value_fullwidth_digits
    true,        // code_char_jp_probes.near_char
    false,       // lazy_forms.lazy_phonetic_unannotated_cell
    false,       // dfunc_kana_folding_probes.dsum_criteria_header_halfwidth_vs_fullwidth_db_header
    false,       // text_format.text_four_section_text
    false,       // text_format.text_dbnum1
    false,       // unmeasured
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
    {'a', 'm', 'j', 'h', 'm', 's', false, false, '\0'},
    {"€", true, true, true, true, 2U},  // text.dollar_negative_minus_sign, text.dollar_rounds_to_negative_zero
    {"€", "", "", ""},                  // value_euro_prefix, value_dollar_prefix
    false,                              // usdollar.usdollar_two_decimals
    DateOrder::kDMY,                    // value_coercion_probes.value_slash_date_short
    false,                              // value_dotted_dmy
    false,                              // datevalue_timevalue.datevalue_kanji_with_terminator
    false,                              // datevalue_timevalue.timevalue_jp_kanji_units
    false,                              // datevalue_timevalue.datevalue_era_reiwa_full
    false,                              // locale_tokens.value_korean_ymd
    false,                              // value_month_name_en
    true,                               // datevalue_timevalue.timevalue_fractional_seconds, timevalue_comma_fraction
    false,                              // datevalue_timevalue.timevalue_trailing_dot
    false,                              // datevalue_timevalue.timevalue_space_dot
    DbcsCodepage::kNone,                // lenb_hangul
    false,
    false,       // value_numbervalue.value_fullwidth_digits
    true,        // code_char_jp_probes.near_char
    false,       // lazy_forms.lazy_phonetic_unannotated_cell
    false,       // dfunc_kana_folding_probes.dsum_criteria_header_halfwidth_vs_fullwidth_db_header
    false,       // text_format.text_four_section_text
    false,       // text_format.text_dbnum1
    false,       // unmeasured
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
    ']',                                  // references.address_r1c1_row_abs_col_rel
    'b',                                  // cell_type_blank
    'l',                                  // cell_type_text
    'v',                                  // cell_type_number
    'G',                                  // cell.cell_format_general
    kEnglishMonthsLong,                   // months_mmmm
    kEnglishMonthsShort,                  // months_mmm
    kEnglishDaysLong,                     // weekdays_dddd
    kEnglishDaysShort,                    // weekdays_ddd
    kInvariantLetters,                    // months_yyyy, months_d, weekdays_aaaa
    {"¥", false, false, true, true, 2U},  // text.dollar_negative_minus_sign, text.dollar_rounds_to_negative_zero
    {"$", "€", "¥", ""},                  // value_dollar_prefix, value_euro_prefix, value_yen_prefix; ￥ unmeasured
    true,                                 // usdollar.usdollar_two_decimals
    DateOrder::kYMD,                      // locale_profile_measurements.datevalue_two_digit_year
    false,                                // value_dotted_ymd
    true,                                 // datevalue_timevalue.datevalue_kanji_with_terminator
    false,                                // datevalue_timevalue.timevalue_jp_kanji_units
    false,                                // datevalue_timevalue.datevalue_era_reiwa_full
    false,                                // locale_tokens.value_korean_ymd
    true,                                 // value_month_name_en
    true,                                 // datevalue_timevalue.timevalue_fractional_seconds, timevalue_comma_fraction
    true,                                 // datevalue_timevalue.timevalue_trailing_dot
    true,                                 // datevalue_timevalue.timevalue_space_dot
    DbcsCodepage::kGb2312,                // lenb_kanji_not_in_gb2312
    false,                                // code_char_jp_probes.char_halfwidth_kata_177
    true,                                 // value_numbervalue.value_fullwidth_digits
    false,                                // code_char_jp_probes.near_char
    true,                                 // lazy_forms.lazy_phonetic_unannotated_cell
    false,         // no width folding (dfunc_kana_folding_probes.dsum_criteria_header_halfwidth_vs_fullwidth_db_header)
    true,          // text_format.text_four_section_text
    true,          // text_format.text_dbnum1
    false,         // unmeasured
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
    ']',                                  // references.address_r1c1_row_abs_col_rel
    'b',                                  // cell_type_blank
    'l',                                  // cell_type_text
    'v',                                  // cell_type_number
    'G',                                  // cell.cell_format_general
    kEnglishMonthsLong,                   // months_mmmm
    kEnglishMonthsShort,                  // months_mmm
    kEnglishDaysLong,                     // weekdays_dddd
    kEnglishDaysShort,                    // weekdays_ddd
    kInvariantLetters,                    // months_yyyy, months_d, weekdays_aaaa
    {"₩", false, false, true, true, 0U},  // text.dollar_negative_minus_sign, text.dollar_zero
    {"$", "€", "₩", ""},                  // value_dollar_prefix, value_euro_prefix, value_won_prefix
    true,                                 // usdollar.usdollar_two_decimals
    DateOrder::kYMD,                      // locale_profile_measurements.datevalue_two_digit_year
    true,                                 // value_dotted_ymd
    false,                                // datevalue_timevalue.datevalue_kanji_with_terminator
    false,                                // datevalue_timevalue.timevalue_jp_kanji_units
    false,                                // datevalue_timevalue.datevalue_era_reiwa_full
    true,                                 // locale_tokens.value_korean_ymd
    true,                                 // value_month_name_en
    false,                                // datevalue_timevalue.timevalue_fractional_seconds, timevalue_comma_fraction
    true,                                 // datevalue_timevalue.timevalue_trailing_dot
    true,                                 // datevalue_timevalue.timevalue_space_dot
    DbcsCodepage::kKsX1001,               // code_hangul
    false,                                // code_char_jp_probes.char_halfwidth_kata_177
    true,                                 // value_numbervalue.value_fullwidth_digits
    false,                                // code_char_jp_probes.near_char
    true,                                 // lazy_forms.lazy_phonetic_unannotated_cell
    false,     // dfunc_kana_folding_probes.dsum_criteria_header_halfwidth_vs_fullwidth_db_header
    true,      // text_format.text_four_section_text
    true,      // text_format.text_dbnum1
    false,     // unmeasured
    "G/표준",  // text_general_g_pyojun
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
    kInvariantLetters,                    // months_yyyy, months_d, weekdays_aaaa
    {"฿", false, false, true, true, 2U},  // text.dollar_negative_minus_sign, text.dollar_rounds_to_negative_zero
    {"€", "฿", "", ""},                   // value_euro_prefix, value_baht_prefix, value_dollar_prefix
    false,                                // usdollar.usdollar_two_decimals
    DateOrder::kDMY,                      // value_coercion_probes.value_slash_date_short
    false,                                // value_dotted_dmy
    false,                                // datevalue_timevalue.datevalue_kanji_with_terminator
    false,                                // datevalue_timevalue.timevalue_jp_kanji_units
    false,                                // datevalue_timevalue.datevalue_era_reiwa_full
    false,                                // locale_tokens.value_korean_ymd
    true,                                 // value_month_name_en
    true,                                 // datevalue_timevalue.timevalue_fractional_seconds, timevalue_comma_fraction
    false,                                // datevalue_timevalue.timevalue_trailing_dot
    true,                                 // datevalue_timevalue.timevalue_space_dot
    DbcsCodepage::kNone,                  // lenb_hangul
    false,
    false,  // value_numbervalue.value_fullwidth_digits
    true,   // code_char_jp_probes.near_char
    false,  // lazy_forms.lazy_phonetic_unannotated_cell
    false,  // dfunc_kana_folding_probes.dsum_criteria_header_halfwidth_vs_fullwidth_db_header
    false,  // text_format.text_four_section_text
    false,  // text_format.text_dbnum1
    false,  // unmeasured
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
  }
  return kEnglishFacts;
}

SbcsCodepage sbcs_codepage(ExcelProfile profile) noexcept {
  switch (profile.locale) {
    case ExcelLocale::kJaJP:
      return SbcsCodepage::kCp932;
    case ExcelLocale::kZhCN:
    case ExcelLocale::kKoKR:
      return SbcsCodepage::kDbcsHighBlank;
    case ExcelLocale::kThTH:
      return SbcsCodepage::kMacThai;
    case ExcelLocale::kEnUS:
    case ExcelLocale::kDeDE:
    case ExcelLocale::kFrFR:
      break;
  }
  return profile.host == ExcelHost::kMac365 ? SbcsCodepage::kMacRoman : SbcsCodepage::kWindows1252;
}

WidthFolding width_folding(ExcelProfile profile) noexcept {
  if (profile.locale == ExcelLocale::kZhCN) {
    return WidthFolding::kNone;
  }
  return profile.host == ExcelHost::kMac365 ? WidthFolding::kMac : WidthFolding::kWin;
}

}  // namespace formulon
