#include "excel_locale.h"

#include "utils/strings.h"

namespace formulon {
namespace {

// Trailing comments name the measuring case: `locale_tokens.<id>` unless a
// suite is given. `// unmeasured` slots hold the en-US value.

using Names7 = std::array<std::string_view, 7>;
using Names12 = std::array<std::string_view, 12>;
using ErrorNames = std::array<std::string_view, kErrorNameCount>;

// Error spellings, in `kErrorTable` order: the classic errors and #GETTING_DATA
// as formula text writes them, #SPILL! and #CALC! as ARRAYTOTEXT does
// (locale_tokens.arraytotext_range_error_*, arraytotext_error_literal_getting_data).
// The errors after #CALC! cannot be produced on Mac and keep the English spelling.
constexpr ErrorNames kEnglishErrorNames{
    "#NULL!", "#DIV/0!", "#VALUE!",   "#REF!",     "#NAME?",     "#NUM!",  "#N/A",     "#GETTING_DATA", "#SPILL!",
    "#CALC!", "#FIELD!", "#BLOCKED!", "#CONNECT!", "#EXTERNAL!", "#BUSY!", "#PYTHON!", "#UNKNOWN!",
};
constexpr ErrorNames kJapaneseErrorNames{
    "#NULL!", "#DIV/0!", "#VALUE!",   "#REF!",     "#NAME?",     "#NUM!",  "#N/A",     "#GETTING_DATA", "#スピル!",
    "#CALC!", "#FIELD!", "#BLOCKED!", "#CONNECT!", "#EXTERNAL!", "#BUSY!", "#PYTHON!", "#UNKNOWN!",
};
constexpr ErrorNames kChineseErrorNames{
    "#NULL!", "#DIV/0!", "#VALUE!",   "#REF!",     "#NAME?",     "#NUM!",  "#N/A",     "#GETTING_DATA", "#溢出!",
    "#CALC!", "#FIELD!", "#BLOCKED!", "#CONNECT!", "#EXTERNAL!", "#BUSY!", "#PYTHON!", "#UNKNOWN!",
};
constexpr ErrorNames kKoreanErrorNames{
    "#NULL!", "#DIV/0!", "#VALUE!",   "#REF!",     "#NAME?",     "#NUM!",  "#N/A",     "#GETTING_DATA", "#분산!",
    "#CALC!", "#FIELD!", "#BLOCKED!", "#CONNECT!", "#EXTERNAL!", "#BUSY!", "#PYTHON!", "#UNKNOWN!",
};
constexpr ErrorNames kThaiErrorNames{
    "#NULL!", "#DIV/0!", "#VALUE!",   "#REF!",     "#NAME?",     "#NUM!",  "#N/A",     "#GETTING_DATA", "#สปิลล์!",
    "#CALC!", "#FIELD!", "#BLOCKED!", "#CONNECT!", "#EXTERNAL!", "#BUSY!", "#PYTHON!", "#UNKNOWN!",
};
constexpr ErrorNames kGermanErrorNames{
    "#NULL!", "#DIV/0!", "#WERT!",    "#BEZUG!",   "#NAME?",     "#ZAHL!", "#NV",      "#DATEN_ABRUFEN", "#ÜBERLAUF!",
    "#CALC!", "#FIELD!", "#BLOCKED!", "#CONNECT!", "#EXTERNAL!", "#BUSY!", "#PYTHON!", "#UNKNOWN!",
};
constexpr ErrorNames kFrenchErrorNames{
    "#NUL!",         "#DIV/0!", "#VALEUR!", "#REF!",     "#NOM?",     "#NOMBRE!",   "#N/A",   "#CHARGEMENT_DONNEES",
    "#PROPAGATION!", "#CALC!",  "#FIELD!",  "#BLOCKED!", "#CONNECT!", "#EXTERNAL!", "#BUSY!", "#PYTHON!",
    "#UNKNOWN!",
};
constexpr ErrorNames kSpanishErrorNames{
    "#¡NULO!",           "#¡DIV/0!", "#¡VALOR!", "#¡REF!",    "#¿NOMBRE?", "#¡NUM!",     "#N/D",   "#OBTENIENDO_DATOS",
    "#¡DESBORDAMIENTO!", "#CALC!",   "#FIELD!",  "#BLOCKED!", "#CONNECT!", "#EXTERNAL!", "#BUSY!", "#PYTHON!",
    "#UNKNOWN!",
};
constexpr ErrorNames kMexicanSpanishErrorNames{
    "#NULO!",    "#DIV/0!",           "#VALOR!",
    "#REF!",     "#NOMBRE?",          "#N¡NUM!",
    "#N/D",      "#OBTENIENDO_DATOS", "#¡DESBORDAMIENTO!",
    "#¡CALC!",   "#FIELD!",           "#BLOCKED!",
    "#CONNECT!", "#EXTERNAL!",        "#BUSY!",
    "#PYTHON!",  "#UNKNOWN!",
};
constexpr ErrorNames kPortugueseErrorNames{
    "#NULO!", "#DIV/0!", "#VALOR!",   "#REF!",     "#NOME?",     "#NÚM!",  "#N/D",     "#OBTENDO_DADOS", "#DESPEJAR!",
    "#CALC!", "#FIELD!", "#BLOCKED!", "#CONNECT!", "#EXTERNAL!", "#BUSY!", "#PYTHON!", "#UNKNOWN!",
};
constexpr ErrorNames kRussianErrorNames{
    "#ПУСТО!",   "#ДЕЛ/0!",          "#ЗНАЧ!",    "#ССЫЛКА!", "#ИМЯ?",     "#ЧИСЛО!",
    "#Н/Д",      "#ОЖИДАНИЕ_ДАННЫХ", "#ПЕРЕНОС!", "#ВЫЧИСЛ!", "#FIELD!",   "#BLOCKED!",
    "#CONNECT!", "#EXTERNAL!",       "#BUSY!",    "#PYTHON!", "#UNKNOWN!",
};
constexpr ErrorNames kItalianErrorNames{
    "#NULL!",       "#DIV/0!",    "#VALORE!", "#RIF!",
    "#NOME?",       "#NUM!",      "#N/D",     "#ESTRAZIONE_DATI_IN_CORSO",
    "#ESPANSIONE!", "#CALC!",     "#FIELD!",  "#BLOCKED!",
    "#CONNECT!",    "#EXTERNAL!", "#BUSY!",   "#PYTHON!",
    "#UNKNOWN!",
};
constexpr ErrorNames kDutchErrorNames{
    "#LEEG!",    "#DELING.DOOR.0!",   "#WAARDE!",    "#VERW!",      "#NAAM?",    "#GETAL!",
    "#N/B",      "#GEGEVENS.OPHALEN", "#OVERLOPEN!", "#BEREKENEN!", "#FIELD!",   "#BLOCKED!",
    "#CONNECT!", "#EXTERNAL!",        "#BUSY!",      "#PYTHON!",    "#UNKNOWN!",
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
constexpr std::array<std::string_view, 8> kEnglishColorNames{"Black",   "Blue", "Cyan",  "Green",
                                                             "Magenta", "Red",  "White", "Yellow"};

// [DBNum1]-[DBNum4]: dbnum_digits, dbnum_place_units, dbnum_place_one_filler, dbnum_general_*,
// dbnum_date_fields.
// A directive the locale accepts and leaves in ASCII digits (dbnum_digits).
constexpr DbnumStyle kAsciiDbnumStyle{
    {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}, {}, {}, false, false, false, false};
constexpr DbnumStyles kJapaneseDbnumStyles{{
    DbnumStyle{{"〇", "一", "二", "三", "四", "五", "六", "七", "八", "九"},
               {"十", "百", "千"},
               {"万", "億", "兆"},
               false,
               false,
               true,
               false},
    DbnumStyle{{"〇", "壱", "弐", "参", "四", "伍", "六", "七", "八", "九"},
               {"拾", "百", "阡"},
               {"萬", "億", "兆"},
               true,
               false,
               true,
               true},
    DbnumStyle{{"０", "１", "２", "３", "４", "５", "６", "７", "８", "９"},
               {"十", "百", "千"},
               {"万", "億", "兆"},
               false,
               false,
               false,
               false},
    kAsciiDbnumStyle,
}};

constexpr DbnumStyles kChineseDbnumStyles{{
    DbnumStyle{{"○", "一", "二", "三", "四", "五", "六", "七", "八", "九"},
               {"十", "百", "千"},
               {"万", "亿", "兆"},
               true,
               true,
               true,
               false},
    DbnumStyle{{"零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖"},
               {"拾", "佰", "仟"},
               {"万", "亿", "兆"},
               true,
               true,
               true,
               true},
    DbnumStyle{{"０", "１", "２", "３", "４", "５", "６", "７", "８", "９"},
               {"十", "百", "千"},
               {"万", "亿", "兆"},
               true,
               true,
               false,
               false},
    kAsciiDbnumStyle,
}};

constexpr DbnumStyles kKoreanDbnumStyles{{
    DbnumStyle{{"０", "一", "二", "三", "四", "五", "六", "七", "八", "九"},
               {"十", "百", "千"},
               {"万", "億", "兆"},
               true,
               false,
               true,
               false},
    DbnumStyle{{"零", "壹", "貳", "參", "四", "伍", "六", "七", "八", "九"},
               {"拾", "百", "阡"},
               {"萬", "億", "兆"},
               true,
               false,
               true,
               true},
    DbnumStyle{{"０", "１", "２", "３", "４", "５", "６", "７", "８", "９"},
               {"十", "百", "千"},
               {"万", "億", "兆"},
               false,
               false,
               false,
               false},
    DbnumStyle{{"영", "일", "이", "삼", "사", "오", "육", "칠", "팔", "구"},
               {"십", "백", "천"},
               {"만", "억", "조"},
               true,
               false,
               true,
               false},
}};

constexpr DbnumStyles kTraditionalChineseDbnumStyles{{
    DbnumStyle{{"○", "一", "二", "三", "四", "五", "六", "七", "八", "九"},
               {"十", "百", "千"},
               {"萬", "億", "兆"},
               true,
               true,
               true,
               false},
    DbnumStyle{{"零", "壹", "貳", "參", "肆", "伍", "陸", "柒", "捌", "玖"},
               {"拾", "佰", "仟"},
               {"萬", "億", "兆"},
               true,
               true,
               true,
               true},
    DbnumStyle{{"０", "１", "２", "３", "４", "５", "６", "７", "８", "９"},
               {"十", "百", "千"},
               {"萬", "億", "兆"},
               true,
               true,
               false,
               false},
    kAsciiDbnumStyle,
}};

// `[DBNum4]` stays in ASCII digits in the section carrying a ko-KR tag (locale_tokens.lcid_dbnum4_sections).
// Digit sets of `[$-02000000]`-`[$-13000000]` in code order; Tamil and Ethiopic
// write an ASCII 0 (locale_tokens.lcid_numeral_systems).
constexpr std::array<DbnumStyle, 18> kNumeralSystemStyles{{
    {{"٠", "١", "٢", "٣", "٤", "٥", "٦", "٧", "٨", "٩"}, {}, {}, false, false, false, false},
    {{"۰", "۱", "۲", "۳", "۴", "۵", "۶", "۷", "۸", "۹"}, {}, {}, false, false, false, false},
    {{"०", "१", "२", "३", "४", "५", "६", "७", "८", "९"}, {}, {}, false, false, false, false},
    {{"০", "১", "২", "৩", "৪", "৫", "৬", "৭", "৮", "৯"}, {}, {}, false, false, false, false},
    {{"੦", "੧", "੨", "੩", "੪", "੫", "੬", "੭", "੮", "੯"}, {}, {}, false, false, false, false},
    {{"૦", "૧", "૨", "૩", "૪", "૫", "૬", "૭", "૮", "૯"}, {}, {}, false, false, false, false},
    {{"୦", "୧", "୨", "୩", "୪", "୫", "୬", "୭", "୮", "୯"}, {}, {}, false, false, false, false},
    {{"0", "௧", "௨", "௩", "௪", "௫", "௬", "௭", "௮", "௯"}, {}, {}, false, false, false, false},
    {{"౦", "౧", "౨", "౩", "౪", "౫", "౬", "౭", "౮", "౯"}, {}, {}, false, false, false, false},
    {{"೦", "೧", "೨", "೩", "೪", "೫", "೬", "೭", "೮", "೯"}, {}, {}, false, false, false, false},
    {{"൦", "൧", "൨", "൩", "൪", "൫", "൬", "൭", "൮", "൯"}, {}, {}, false, false, false, false},
    {{"๐", "๑", "๒", "๓", "๔", "๕", "๖", "๗", "๘", "๙"}, {}, {}, false, false, false, false},
    {{"໐", "໑", "໒", "໓", "໔", "໕", "໖", "໗", "໘", "໙"}, {}, {}, false, false, false, false},
    {{"༠", "༡", "༢", "༣", "༤", "༥", "༦", "༧", "༨", "༩"}, {}, {}, false, false, false, false},
    {{"၀", "၁", "၂", "၃", "၄", "၅", "၆", "၇", "၈", "၉"}, {}, {}, false, false, false, false},
    {{"0", "፩", "፪", "፫", "፬", "፭", "፮", "፯", "፰", "፱"}, {}, {}, false, false, false, false},
    {{"០", "១", "២", "៣", "៤", "៥", "៦", "៧", "៨", "៩"}, {}, {}, false, false, false, false},
    {{"᠐", "᠑", "᠒", "᠓", "᠔", "᠕", "᠖", "᠗", "᠘", "᠙"}, {}, {}, false, false, false, false},
}};

// locale_tokens.lcid_names_411
constexpr TagLanguage kTagJapanese{
    {"1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"},
    {"1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"},
    {"日曜日", "月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日"},
    {"日", "月", "火", "水", "木", "金", "土"},
    "午前",
    "午後",
    &kJapaneseDbnumStyles,
    true,
};

// locale_tokens.lcid_names_409
constexpr TagLanguage kTagEnglish{
    {"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November",
     "December"},
    {"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"},
    {"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"},
    {"Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"},
    "AM",
    "PM",
    nullptr,
    false,
};

// locale_tokens.lcid_names_407
constexpr TagLanguage kTagGerman{
    {"Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November",
     "Dezember"},
    {"Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"},
    {"Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"},
    {"So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"},
    "AM",
    "PM",
    nullptr,
    false,
};

// locale_tokens.lcid_names_40c
constexpr TagLanguage kTagFrench{
    {"janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre",
     "décembre"},
    {"janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc."},
    {"dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"},
    {"dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."},
    "AM",
    "PM",
    nullptr,
    false,
};

// locale_tokens.lcid_names_804
constexpr TagLanguage kTagChinese{
    {"一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月"},
    {"1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"},
    {"星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"},
    {"周日", "周一", "周二", "周三", "周四", "周五", "周六"},
    "上午",
    "下午",
    &kChineseDbnumStyles,
    false,
};

// locale_tokens.lcid_names_412
constexpr TagLanguage kTagKorean{
    {"1월", "2월", "3월", "4월", "5월", "6월", "7월", "8월", "9월", "10월", "11월", "12월"},
    {"1월", "2월", "3월", "4월", "5월", "6월", "7월", "8월", "9월", "10월", "11월", "12월"},
    {"일요일", "월요일", "화요일", "수요일", "목요일", "금요일", "토요일"},
    {"일", "월", "화", "수", "목", "금", "토"},
    "오전",
    "오후",
    &kKoreanDbnumStyles,
    false,
};

// locale_tokens.lcid_names_41e
constexpr TagLanguage kTagThai{
    {"มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน",
     "ธันวาคม"},
    {"ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."},
    {"วันอาทิตย์", "วันจันทร์", "วันอังคาร", "วันพุธ", "วันพฤหัสบดี", "วันศุกร์", "วันเสาร์"},
    {"อาทิตย์", "จันทร์", "อังคาร", "พุธ", "พฤหัส", "ศุกร์", "เสาร์"},
    "AM",
    "PM",
    nullptr,
    false,
};

// locale_tokens.lcid_names_c0a
constexpr TagLanguage kTagSpanish{
    {"enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre",
     "diciembre"},
    {"ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic."},
    {"domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"},
    {"do.", "lu.", "ma.", "mi.", "ju.", "vi.", "sá."},
    "a. m.",
    "p. m.",
    nullptr,
    false,
};

// locale_tokens.lcid_names_80a
constexpr TagLanguage kTagMexicanSpanish{
    {"enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre",
     "diciembre"},
    {"ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic."},
    {"domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"},
    {"dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."},
    "a. m.",
    "p. m.",
    nullptr,
    false,
};

// locale_tokens.lcid_names_416
constexpr TagLanguage kTagPortuguese{
    {"janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro",
     "dezembro"},
    {"jan.", "fev.", "mar.", "abr.", "mai.", "jun.", "jul.", "ago.", "set.", "out.", "nov.", "dez."},
    {"domingo", "segunda-feira", "terça-feira", "quarta-feira", "quinta-feira", "sexta-feira", "sábado"},
    {"dom.", "seg.", "ter.", "qua.", "qui.", "sex.", "sáb."},
    "AM",
    "PM",
    nullptr,
    false,
};

// locale_tokens.lcid_names_419
constexpr TagLanguage kTagRussian{
    {"январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август", "сентябрь", "октябрь", "ноябрь",
     "декабрь"},
    {"янв.", "февр.", "март", "апр.", "май", "июнь", "июль", "авг.", "сент.", "окт.", "нояб.", "дек."},
    {"воскресенье", "понедельник", "вторник", "среда", "четверг", "пятница", "суббота"},
    {"Вс", "Пн", "Вт", "Ср", "Чт", "Пт", "Сб"},
    "AM",
    "PM",
    nullptr,
    false,
};

// 0xFC19 writes genitive month names, as the ru-RU system long date does (locale_tokens.text_system_date_tag).
constexpr TagLanguage kTagRussianGenitive{
    {"января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября",
     "декабря"},
    kTagRussian.month_short,
    kTagRussian.day_long,
    kTagRussian.day_short,
    kTagRussian.am_name,
    kTagRussian.pm_name,
    nullptr,
    false,
};

// locale_tokens.lcid_names_404
constexpr TagLanguage kTagTraditionalChinese{
    {"1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"},
    {"1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"},
    {"星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"},
    {"週日", "週一", "週二", "週三", "週四", "週五", "週六"},
    "上午",
    "下午",
    &kTraditionalChineseDbnumStyles,
    false,
};

// locale_tokens.lcid_names_410
constexpr TagLanguage kTagItalian{
    {"gennaio", "febbraio", "marzo", "aprile", "maggio", "giugno", "luglio", "agosto", "settembre", "ottobre",
     "novembre", "dicembre"},
    {"gen", "feb", "mar", "apr", "mag", "giu", "lug", "ago", "set", "ott", "nov", "dic"},
    {"domenica", "lunedì", "martedì", "mercoledì", "giovedì", "venerdì", "sabato"},
    {"dom", "lun", "mar", "mer", "gio", "ven", "sab"},
    "AM",
    "PM",
    nullptr,
    false,
};

// locale_tokens.lcid_names_413
constexpr TagLanguage kTagDutch{
    {"januari", "februari", "maart", "april", "mei", "juni", "juli", "augustus", "september", "oktober", "november",
     "december"},
    {"jan", "feb", "mrt", "apr", "mei", "jun", "jul", "aug", "sep", "okt", "nov", "dec"},
    {"zondag", "maandag", "dinsdag", "woensdag", "donderdag", "vrijdag", "zaterdag"},
    {"zo", "ma", "di", "wo", "do", "vr", "za"},
    "AM",
    "PM",
    nullptr,
    false,
};

// @size-budget: 12 KB
constexpr LocaleFacts kJapaneseFacts{
    '.',                  // locale_tokens.fixed_negative
    ',',                  // fixed_negative
    ',',                  // formulatext_bool_literal
    ',',                  // formulatext_array_constant
    ';',                  // formulatext_array_constant
    "TRUE",               // bool_text_true
    "FALSE",              // bool_text_false
    false,                // text_to_bool_probes.text_bool_and_whitespace, text_bool_and_two_true
    kJapaneseErrorNames,  // locale_tokens.arraytotext_range_error_*
    'R',
    'C',
    '[',
    ']',                      // references.address_r1c1_row_abs_col_rel
    'b',                      // cell_type_blank
    'l',                      // cell_type_text
    'v',                      // cell_type_number
    'G',                      // cell.cell_format_general
    kEnglishMonthsLong,       // months_mmmm
    kEnglishMonthsShort,      // months_mmm
    kEnglishDaysLong,         // weekdays_dddd
    kEnglishDaysShort,        // weekdays_ddd
    kInvariantFormatLetters,  // months_yyyy, months_d, weekdays_aaaa
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
    true,                   // code_char_jp_probes.char_halfwidth_kata_177
    true,                   // value_numbervalue.value_fullwidth_digits
    false,                  // code_char_jp_probes.near_char
    true,                   // lazy_forms.lazy_phonetic_unannotated_cell
    true,                   // dfunc_kana_folding_probes.dsum_criteria_header_halfwidth_vs_fullwidth_db_header
    true,                   // text_format.text_four_section_text
    &kJapaneseDbnumStyles,  // locale_tokens.dbnum_digits
    false,                  // locale_tokens.lcid_dbnum4_sections
    false,                  // locale_tokens.text_letter_t_digits
    '\0',                   // locale_tokens.text_letter_x_positive, text_letter_x_negative, text_slash_letters
    false,                  // locale_tokens.text_fullwidth_number_syntax, text_fullwidth_percent_syntax
    false,                  // text_general_english
    false,                  // text_decimal_point
    "G/標準",               // text_general_g_hyojun
    {"黒", "青", "水", "緑", "紫", "赤", "白", "黄"},  // text_colors_ja_jp
    "色",                                              // text_color_index_prefixes
    "合計",
    "総計",
    "行フィールド ",
    "列フィールド ",
    "値 ",
    {"日曜日", "月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日"},  // weekdays_aaaa
    {"日", "月", "火", "水", "木", "金", "土"},                              // weekdays_aaa
};

constexpr LocaleFacts kEnglishFacts{
    '.',                 // fixed_negative
    ',',                 // fixed_negative
    ',',                 // formulatext_bool_literal
    ',',                 // formulatext_array_constant
    ';',                 // formulatext_array_constant
    "TRUE",              // bool_text_true
    "FALSE",             // bool_text_false
    false,               // text_to_bool_probes.text_bool_and_whitespace, text_bool_and_two_true
    kEnglishErrorNames,  // locale_tokens.arraytotext_range_error_*
    'R',
    'C',
    '[',
    ']',                      // references.address_r1c1_row_abs_col_rel
    'b',                      // cell_type_blank
    'l',                      // cell_type_text
    'v',                      // cell_type_number
    'G',                      // cell.cell_format_general
    kEnglishMonthsLong,       // months_mmmm
    kEnglishMonthsShort,      // months_mmm
    kEnglishDaysLong,         // weekdays_dddd
    kEnglishDaysShort,        // weekdays_ddd
    kInvariantFormatLetters,  // months_yyyy, months_d, weekdays_aaaa
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
    nullptr,             // locale_tokens.dbnum_digits
    false,               // locale_tokens.lcid_dbnum4_sections
    false,               // locale_tokens.text_letter_t_digits
    '\0',                // locale_tokens.text_letter_x_positive, text_letter_x_negative, text_slash_letters
    true,                // locale_tokens.text_fullwidth_number_syntax, text_fullwidth_percent_syntax
    true,                // text_general_english
    false,               // text_decimal_point
    "",                  // text_general_english
    kEnglishColorNames,  // text_colors_en_us
    "Color",             // text_color_index_prefixes
    "Total",
    "Grand Total",
    "Row Field ",
    "Column Field ",
    "Value ",
    kEnglishDaysLong,   // weekdays_aaaa
    kEnglishDaysShort,  // weekdays_aaa
};

constexpr LocaleFacts kGermanFacts{
    ',',                // fixed_negative
    '.',                // fixed_negative
    ';',                // formulatext_bool_literal
    '.',                // formulatext_array_constant
    ';',                // formulatext_array_constant
    "WAHR",             // bool_text_true
    "FALSCH",           // bool_text_false
    true,               // text_to_bool_probes.text_bool_and_whitespace, text_bool_and_two_true
    kGermanErrorNames,  // locale_tokens.arraytotext_range_error_*
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
    nullptr,     // locale_tokens.dbnum_digits
    false,       // locale_tokens.lcid_dbnum4_sections
    false,       // locale_tokens.text_letter_t_digits
    '\0',        // locale_tokens.text_letter_x_positive, text_letter_x_negative, text_slash_letters
    true,        // locale_tokens.text_fullwidth_number_syntax, text_fullwidth_percent_syntax
    false,       // text_general_english
    false,       // text_decimal_point
    "Standard",  // text_general_standard
    {"Schwarz", "Blau", "Zyan", "Grün", "Magenta", "Rot", "Weiß", "Gelb"},  // text_colors_de_de
    "Farbe",                                                                // text_color_index_prefixes
    "Gesamt",                                                               // groupby.groupby_measured_fh2_hdr
    "Gesamtergebnis",                                                       // pivotby.pivotby_row_subtotal_depth_two
    "Zeilenfeld ",                                                          // groupby.groupby_measured_fh2_hdr
    "Spaltenfeld ",                                                         // pivotby.pivotby_measured_fh2_hdr
    "Wert ",                                                                // groupby.groupby_measured_fh2_hdr
    {"Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"},  // weekdays_aaaa
    {"So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"},                                         // weekdays_aaa
};

constexpr LocaleFacts kFrenchFacts{
    ',',                // fixed_negative
    ' ',                // fixed_negative
    ';',                // formulatext_bool_literal
    '.',                // formulatext_array_constant
    ';',                // formulatext_array_constant
    "VRAI",             // bool_text_true
    "FAUX",             // bool_text_false
    true,               // text_to_bool_probes.text_bool_and_whitespace, text_bool_and_two_true
    kFrenchErrorNames,  // locale_tokens.arraytotext_range_error_*
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
    nullptr,     // locale_tokens.dbnum_digits
    false,       // locale_tokens.lcid_dbnum4_sections
    false,       // locale_tokens.text_letter_t_digits
    '\0',        // locale_tokens.text_letter_x_positive, text_letter_x_negative, text_slash_letters
    true,        // locale_tokens.text_fullwidth_number_syntax, text_fullwidth_percent_syntax
    false,       // text_general_english
    false,       // text_decimal_point
    "Standard",  // text_general_standard
    {"Noir", "Bleu", "Cyan", "Vert", "Magenta", "Rouge", "Blanc", "Jaune"},  // text_colors_fr_fr
    "Couleur",                                                               // text_color_index_prefixes
    "Total",                                                                 // groupby.groupby_measured_fh2_hdr
    "Total général",                                                         // pivotby.pivotby_row_subtotal_depth_two
    "Champ de ligne ",                                                       // groupby.groupby_measured_fh2_hdr
    "Champ de colonne ",                                                     // pivotby.pivotby_measured_fh2_hdr
    "Valeur ",                                                               // groupby.groupby_measured_fh2_hdr
    kEnglishDaysLong,   // unmeasured: aaaa is a year token (weekdays_aaaa)
    kEnglishDaysShort,  // unmeasured: aaa is a year token (weekdays_aaa)
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
    kChineseErrorNames,  // locale_tokens.arraytotext_range_error_*
    'R',
    'C',
    '[',
    ']',                      // references.address_r1c1_row_abs_col_rel
    'b',                      // cell_type_blank
    'l',                      // cell_type_text
    'v',                      // cell_type_number
    'G',                      // cell.cell_format_general
    kEnglishMonthsLong,       // months_mmmm
    kEnglishMonthsShort,      // months_mmm
    kEnglishDaysLong,         // weekdays_dddd
    kEnglishDaysShort,        // weekdays_ddd
    kInvariantFormatLetters,  // months_yyyy, months_d, weekdays_aaaa
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
    false,  // no width folding (dfunc_kana_folding_probes.dsum_criteria_header_halfwidth_vs_fullwidth_db_header)
    true,   // text_format.text_four_section_text
    &kChineseDbnumStyles,  // locale_tokens.dbnum_digits
    false,                 // locale_tokens.lcid_dbnum4_sections
    false,                 // locale_tokens.text_letter_t_digits
    '\0',                  // locale_tokens.text_letter_x_positive, text_letter_x_negative, text_slash_letters
    false,                 // locale_tokens.text_fullwidth_number_syntax, text_fullwidth_percent_syntax
    false,                 // text_general_english
    false,                 // text_decimal_point
    "G/通用格式",          // text_general_g_tongyong
    {"黑色", "蓝色", "蓝绿色", "绿色", "洋红", "红色", "白色", "黄色"},      // text_colors_zh_cn
    "颜色",                                                                  // text_color_index_prefixes
    "总计",                                                                  // groupby.groupby_measured_fh2_hdr
    "总计",                                                                  // pivotby.pivotby_row_subtotal_depth_two
    "行字段 ",                                                               // groupby.groupby_measured_fh2_hdr
    "列字段 ",                                                               // pivotby.pivotby_measured_fh2_hdr
    "值 ",                                                                   // groupby.groupby_measured_fh2_hdr
    {"星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"},  // weekdays_aaaa
    {"周日", "周一", "周二", "周三", "周四", "周五", "周六"},                // weekdays_aaa
};

constexpr LocaleFacts kKoreanFacts{
    '.',                // fixed_negative
    ',',                // fixed_negative
    ',',                // formulatext_bool_literal
    ',',                // formulatext_array_constant
    ';',                // formulatext_array_constant
    "TRUE",             // bool_text_true
    "FALSE",            // bool_text_false
    false,              // text_to_bool_probes.text_bool_and_whitespace, text_bool_and_two_true
    kKoreanErrorNames,  // locale_tokens.arraytotext_range_error_*
    'R',
    'C',
    '[',
    ']',                      // references.address_r1c1_row_abs_col_rel
    'b',                      // cell_type_blank
    'l',                      // cell_type_text
    'v',                      // cell_type_number
    'G',                      // cell.cell_format_general
    kEnglishMonthsLong,       // months_mmmm
    kEnglishMonthsShort,      // months_mmm
    kEnglishDaysLong,         // weekdays_dddd
    kEnglishDaysShort,        // weekdays_ddd
    kInvariantFormatLetters,  // months_yyyy, months_d, weekdays_aaaa
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
    &kKoreanDbnumStyles,     // locale_tokens.dbnum_digits
    false,                   // locale_tokens.lcid_dbnum4_sections
    false,                   // locale_tokens.text_letter_t_digits
    '\0',                    // locale_tokens.text_letter_x_positive, text_letter_x_negative, text_slash_letters
    false,                   // locale_tokens.text_fullwidth_number_syntax, text_fullwidth_percent_syntax
    false,                   // text_general_english
    false,                   // text_decimal_point
    "G/표준",                // text_general_g_pyojun
    {"검정", "파랑", "녹청", "녹색", "자홍", "빨강", "흰색", "노랑"},        // text_colors_ko_kr
    "색",                                                                    // text_color_index_prefixes
    "합계",                                                                  // groupby.groupby_measured_fh2_hdr
    "총합계",                                                                // pivotby.pivotby_row_subtotal_depth_two
    "행 필드 ",                                                              // groupby.groupby_measured_fh2_hdr
    "열 필드 ",                                                              // pivotby.pivotby_measured_fh2_hdr
    "값 ",                                                                   // groupby.groupby_measured_fh2_hdr
    {"일요일", "월요일", "화요일", "수요일", "목요일", "금요일", "토요일"},  // weekdays_aaaa
    {"일", "월", "화", "수", "목", "금", "토"},                              // weekdays_aaa
};

constexpr LocaleFacts kThaiFacts{
    '.',              // fixed_negative
    ',',              // fixed_negative
    ',',              // formulatext_bool_literal
    ',',              // formulatext_array_constant
    ';',              // formulatext_array_constant
    "TRUE",           // bool_text_true
    "FALSE",          // bool_text_false
    false,            // text_to_bool_probes.text_bool_and_whitespace, text_bool_and_two_true
    kThaiErrorNames,  // locale_tokens.arraytotext_range_error_*
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
    kInvariantFormatLetters,  // months_yyyy, months_d, weekdays_aaaa
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
    false,    // value_numbervalue.value_fullwidth_digits
    true,     // code_char_jp_probes.near_char
    false,    // lazy_forms.lazy_phonetic_unannotated_cell
    false,    // dfunc_kana_folding_probes.dsum_criteria_header_halfwidth_vs_fullwidth_db_header
    false,    // text_format.text_four_section_text
    nullptr,  // locale_tokens.dbnum_digits
    true,     // locale_tokens.lcid_dbnum4_sections
    true,     // locale_tokens.text_letter_t_digits
    '\0',     // locale_tokens.text_letter_x_positive, text_letter_x_negative, text_slash_letters
    true,     // locale_tokens.text_fullwidth_number_syntax, text_fullwidth_percent_syntax
    true,     // text_general_english
    false,    // text_decimal_point
    "",       // text_general_english
    {"ดำ", "น้ำเงิน", "ฟ้า", "เขียว", "ม่วงมาเจนต้า", "แดง", "ขาว", "เหลือง"},       // text_colors_th_th
    "สี",                                                                      // text_color_index_prefixes
    "ผลรวม",                                                                  // groupby.groupby_measured_fh2_hdr
    "ผลรวมทั้งหมด",                                                             // pivotby.pivotby_row_subtotal_depth_two
    "เขตข้อมูลแถว ",                                                            // groupby.groupby_measured_fh2_hdr
    "เขตข้อมูลคอลัมน์ ",                                                          // pivotby.pivotby_measured_fh2_hdr
    "ค่า ",                                                                    // groupby.groupby_measured_fh2_hdr
    {"วันอาทิตย์", "วันจันทร์", "วันอังคาร", "วันพุธ", "วันพฤหัสบดี", "วันศุกร์", "วันเสาร์"},  // weekdays_aaaa
    {"อาทิตย์", "จันทร์", "อังคาร", "พุธ", "พฤหัส", "ศุกร์", "เสาร์"},                  // weekdays_aaa
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
    kRussianErrorNames,  // locale_tokens.arraytotext_range_error_*
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
    nullptr,              // locale_tokens.dbnum_digits
    false,                // locale_tokens.lcid_dbnum4_sections
    false,                // locale_tokens.text_letter_t_digits
    '\0',                 // locale_tokens.text_letter_x_positive, text_letter_x_negative, text_slash_letters
    true,                 // locale_tokens.text_fullwidth_number_syntax, text_fullwidth_percent_syntax
    false,                // text_general_english
    true,                 // text_decimal_point
    "Основной",           // text_general_osnovnoy; General and Standard are rejected (text_general_english,
                          // text_general_standard)
    {"Черный", "Синий", "Голубой", "Зеленый", "Фиолетовый", "Красный", "Белый", "Желтый"},  // text_colors_ru_ru
    "Цвет",                                                                                 // text_color_index_prefixes
    "Итого",          // groupby.groupby_measured_fh2_hdr
    "Общий итог",     // pivotby.pivotby_row_subtotal_depth_two
    "Поле строки ",   // groupby.groupby_measured_fh2_hdr
    "Поле столбца ",  // pivotby.pivotby_measured_fh2_hdr
    "Значение ",      // groupby.groupby_measured_fh2_hdr
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
    kChineseErrorNames,  // locale_tokens.arraytotext_range_error_*
    'R',
    'C',
    '[',
    ']',                      // references.address_r1c1_row_abs_col_rel
    'b',                      // cell_type_blank
    'l',                      // cell_type_text
    'v',                      // cell_type_number
    'G',                      // cell.cell_format_general
    kEnglishMonthsLong,       // months_mmmm
    kEnglishMonthsShort,      // months_mmm
    kEnglishDaysLong,         // weekdays_dddd
    kEnglishDaysShort,        // weekdays_ddd
    kInvariantFormatLetters,  // months_yyyy, months_d, weekdays_aaaa
    kNoLetterAliases,
    {"$", false, false, true, false, true,
     2U},                             // text.dollar_negative_minus_sign, usdollar.dollar_rounds_to_negative_zero
    {"$", "€", "", ""},               // value_dollar_prefix, value_euro_prefix, value_yen_prefix
    "US$",                            // usdollar.usdollar_two_decimals
    DateOrder::kYMD,                  // locale_profile_measurements.datevalue_two_digit_year
    false,                            // value_dotted_ymd
    true,                             // datevalue_timevalue.datevalue_kanji_with_terminator
    true,                             // datevalue_timevalue.timevalue_jp_kanji_units
    false,                            // datevalue_timevalue.datevalue_era_reiwa_full
    RLetter::kYear,                   // locale_tokens.text_letter_r_date, text_letter_rr_date
    false,                            // locale_tokens.value_korean_ymd
    true,                             // value_month_name_en
    true,                             // value_coercion_probes.value_d_mmm_yy
    true,                             // datevalue_timevalue.timevalue_fractional_seconds
    true,                             // datevalue_timevalue.timevalue_trailing_dot
    true,                             // datevalue_timevalue.timevalue_space_dot
    "AM",                             // text_format.text_time_am
    "PM",                             // text_format.text_time_pm
    true,                             // text_format.text_time_a_p
    DbcsCodepage::kBig5,              // code_char_jp_probes (Big5, no ETEN rows), lenb_kanji_not_in_gb2312
    false,                            // code_char_jp_probes.char_halfwidth_kata_177
    true,                             // value_numbervalue.value_fullwidth_digits
    false,                            // code_char_jp_probes.near_char
    true,                             // lazy_forms.lazy_phonetic_unannotated_cell
    false,                            // dfunc_kana_folding_probes.dsum_criteria_header_halfwidth_vs_fullwidth_db_header
    true,                             // text_format.text_four_section_text
    &kTraditionalChineseDbnumStyles,  // locale_tokens.dbnum_digits
    false,                            // locale_tokens.lcid_dbnum4_sections
    false,                            // locale_tokens.text_letter_t_digits
    '\0',          // locale_tokens.text_letter_x_positive, text_letter_x_negative, text_slash_letters
    false,         // locale_tokens.text_fullwidth_number_syntax, text_fullwidth_percent_syntax
    false,         // text_general_english
    false,         // text_decimal_point
    "G/通用格式",  // text_general_g_tongyong
    {"黑色", "藍色", "青色", "綠色", "洋紅", "紅色", "白色", "黃色"},        // text_colors_zh_tw
    "色彩",                                                                  // text_color_index_prefixes
    "總計",                                                                  // groupby.groupby_measured_fh2_hdr
    "總計",                                                                  // pivotby.pivotby_row_subtotal_depth_two
    "列欄位 ",                                                               // groupby.groupby_measured_fh2_hdr
    "欄欄位 ",                                                               // pivotby.pivotby_measured_fh2_hdr
    "值 ",                                                                   // groupby.groupby_measured_fh2_hdr
    {"星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"},  // weekdays_aaaa
    {"週日", "週一", "週二", "週三", "週四", "週五", "週六"},                // weekdays_aaa
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
    kItalianErrorNames,  // locale_tokens.arraytotext_range_error_*
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
    nullptr,              // locale_tokens.dbnum_digits
    false,                // locale_tokens.lcid_dbnum4_sections
    false,                // locale_tokens.text_letter_t_digits
    'x',                  // locale_tokens.text_letter_x_positive, text_letter_x_negative, text_slash_letters
    true,                 // locale_tokens.text_fullwidth_number_syntax, text_fullwidth_percent_syntax
    false,                // text_general_english
    false,                // text_decimal_point
    "Standard",           // text_general_standard
    {"Nero", "Blu", "Celeste", "Verde", "Fucsia", "Rosso", "Bianco", "Giallo"},  // text_colors_it_it
    "Colore",                                                                    // text_color_index_prefixes
    "Totale",                                                                    // groupby.groupby_measured_fh2_hdr
    "Totale complessivo",  // pivotby.pivotby_row_subtotal_depth_two
    "Campo riga ",         // groupby.groupby_measured_fh2_hdr
    "Campo colonna ",      // pivotby.pivotby_measured_fh2_hdr
    "Valore ",             // groupby.groupby_measured_fh2_hdr
    kEnglishDaysLong,      // unmeasured: aaaa is a year token (weekdays_aaaa)
    kEnglishDaysShort,     // unmeasured: aaa is a year token (weekdays_aaa)
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
    kDutchErrorNames,  // locale_tokens.arraytotext_range_error_*
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
    nullptr,              // locale_tokens.dbnum_digits
    false,                // locale_tokens.lcid_dbnum4_sections
    false,                // locale_tokens.text_letter_t_digits
    '\0',                 // locale_tokens.text_letter_x_positive, text_letter_x_negative, text_slash_letters
    true,                 // locale_tokens.text_fullwidth_number_syntax, text_fullwidth_percent_syntax
    false,                // text_general_english
    false,                // text_decimal_point
    "Standaard",          // text_general_standaard; General and Standard are rejected (text_general_english,
                          // text_general_standard)
    {"Zwart", "Blauw", "Cyaan", "Groen", "Magenta", "Rood", "Wit", "Geel"},  // text_colors_nl_nl
    "Kleur",                                                                 // text_color_index_prefixes
    "Totaal",                                                                // groupby.groupby_measured_fh2_hdr
    "Eindtotaal",                                                            // pivotby.pivotby_row_subtotal_depth_two
    "Rijveld ",                                                              // groupby.groupby_measured_fh2_hdr
    "Kolomveld ",                                                            // pivotby.pivotby_measured_fh2_hdr
    "Waarde ",                                                               // groupby.groupby_measured_fh2_hdr
    {"zondag", "maandag", "dinsdag", "woensdag", "donderdag", "vrijdag", "zaterdag"},  // weekdays_aaaa
    {"zo", "ma", "di", "wo", "do", "vr", "za"},                                        // weekdays_aaa
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
    kSpanishErrorNames,  // locale_tokens.arraytotext_range_error_*
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
    nullptr,              // locale_tokens.dbnum_digits
    false,                // locale_tokens.lcid_dbnum4_sections
    false,                // locale_tokens.text_letter_t_digits
    '\0',                 // locale_tokens.text_letter_x_positive, text_letter_x_negative, text_slash_letters
    true,                 // locale_tokens.text_fullwidth_number_syntax, text_fullwidth_percent_syntax
    false,                // text_general_english
    false,                // text_decimal_point
    "Estándar",           // text_general_estandar; General and Standard are rejected (text_general_english,
                          // text_general_standard)
    {"Negro", "Azul", "Cian", "Verde", "Magenta", "Rojo", "Blanco", "Amarillo"},  // text_colors_es_es
    "Color",                                                                      // text_color_index_prefixes
    "Total",                                                                      // groupby.groupby_measured_fh2_hdr
    "Total general",      // pivotby.pivotby_row_subtotal_depth_two
    "Campo de fila ",     // groupby.groupby_measured_fh2_hdr
    "Campo de columna ",  // pivotby.pivotby_measured_fh2_hdr
    "Valor ",             // groupby.groupby_measured_fh2_hdr
    kEnglishDaysLong,     // unmeasured: aaaa is a year token (weekdays_aaaa)
    kEnglishDaysShort,    // unmeasured: aaa is a year token (weekdays_aaa)
};

// es-MX keeps the en-US separators and number-format dialect; month and weekday names, bool names, currency
// and date order are its own.
constexpr LocaleFacts kMexicanSpanishFacts{
    '.',                        // fixed_negative
    ',',                        // fixed_negative
    ',',                        // formulatext_bool_literal
    ',',                        // formulatext_array_constant
    ';',                        // formulatext_array_constant
    "VERDADERO",                // bool_text_true
    "FALSO",                    // bool_text_false
    true,                       // text_to_bool_probes.text_bool_or_whitespace
    kMexicanSpanishErrorNames,  // locale_tokens.arraytotext_range_error_*
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
    nullptr,              // locale_tokens.dbnum_digits
    false,                // locale_tokens.lcid_dbnum4_sections
    false,                // locale_tokens.text_letter_t_digits
    '\0',                 // locale_tokens.text_letter_x_positive, text_letter_x_negative, text_slash_letters
    true,                 // locale_tokens.text_fullwidth_number_syntax, text_fullwidth_percent_syntax
    false,                // text_general_english
    false,                // text_decimal_point
    "Estándar",           // text_general_estandar; General and Standard are rejected (text_general_english,
                          // text_general_standard)
    {"Negro", "Azul", "Cian", "Verde", "Magenta", "Rojo", "Blanco", "Amarillo"},  // text_colors_es_mx
    "Color",                                                                      // text_color_index_prefixes
    "Total",                                                                      // groupby.groupby_measured_fh2_hdr
    "Total general",      // pivotby.pivotby_row_subtotal_depth_two
    "Campo de fila ",     // groupby.groupby_measured_fh2_hdr
    "Campo de columna ",  // pivotby.pivotby_measured_fh2_hdr
    "Valor ",             // groupby.groupby_measured_fh2_hdr
    kEnglishDaysLong,     // unmeasured: aaaa is a year token (weekdays_aaaa)
    kEnglishDaysShort,    // unmeasured: aaa is a year token (weekdays_aaa)
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
    kPortugueseErrorNames,  // locale_tokens.arraytotext_range_error_*
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
    nullptr,              // locale_tokens.dbnum_digits
    false,                // locale_tokens.lcid_dbnum4_sections
    false,                // locale_tokens.text_letter_t_digits
    '\0',                 // locale_tokens.text_letter_x_positive, text_letter_x_negative, text_slash_letters
    true,                 // locale_tokens.text_fullwidth_number_syntax, text_fullwidth_percent_syntax
    false,                // text_general_english
    false,                // text_decimal_point
    "Geral",  // text_general_geral; General and Standard are rejected (text_general_english, text_general_standard)
    {"Preto", "Azul", "Ciano", "Verde", "Magenta", "Vermelho", "Branco", "Amarelo"},  // text_colors_pt_br
    "Cor",                                                                            // text_color_index_prefixes
    "Total",             // groupby.groupby_measured_fh2_hdr
    "Total Geral",       // pivotby.pivotby_row_subtotal_depth_two
    "Campo de linha ",   // groupby.groupby_measured_fh2_hdr
    "Campo de coluna ",  // pivotby.pivotby_measured_fh2_hdr
    "Valor ",            // groupby.groupby_measured_fh2_hdr
    kEnglishDaysLong,    // unmeasured: aaaa is a year token (weekdays_aaaa)
    kEnglishDaysShort,   // unmeasured: aaa is a year token (weekdays_aaa)
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
  // #NULL! through #CALC!: locale_tokens.arraytotext_range_error_* and
  // arraytotext_error_literal_getting_data run under every Mac locale.
  return error_ordinal <= static_cast<std::size_t>(ErrorCode::Calc);
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

const TagLanguage* tag_language_for_lcid(std::uint16_t lcid) noexcept {
  // Regional LCIDs measured to share the base spellings (locale_tokens.lcid_regional_variants).
  switch (lcid) {
    case 0x0411:
      return &kTagJapanese;
    case 0x0409:
      return &kTagEnglish;
    case 0x0407:
    case 0x0807:
      return &kTagGerman;
    case 0x040C:
    case 0x080C:
      return &kTagFrench;
    case 0x0804:
      return &kTagChinese;
    case 0x0412:
      return &kTagKorean;
    case 0x041E:
      return &kTagThai;
    case 0x0C0A:
    case 0x040A:
      return &kTagSpanish;
    case 0x080A:
      return &kTagMexicanSpanish;
    case 0x0416:
      return &kTagPortuguese;
    case 0x0419:
      return &kTagRussian;
    case 0xFC19:
      return &kTagRussianGenitive;
    case 0x0404:
      return &kTagTraditionalChinese;
    case 0x0410:
    case 0x0810:
      return &kTagItalian;
    case 0x0413:
    case 0x0813:
      return &kTagDutch;
    default:
      return nullptr;
  }
}

const TagLanguage* tag_language_for_name(std::string_view name) noexcept {
  // Bare language names measured to match the regional tables (locale_tokens.lcid_language_names).
  struct NamedLcid {
    std::string_view name;
    std::uint16_t lcid;
  };
  static constexpr NamedLcid kNames[] = {
      {"ja-JP", 0x0411}, {"ja", 0x0411},    {"en-US", 0x0409}, {"en", 0x0409},    {"de-DE", 0x0407},
      {"fr-FR", 0x040C}, {"zh-CN", 0x0804}, {"ko-KR", 0x0412}, {"th-TH", 0x041E}, {"es-ES", 0x0C0A},
      {"es-MX", 0x080A}, {"pt-BR", 0x0416}, {"ru-RU", 0x0419}, {"zh-TW", 0x0404}, {"zh-Hant", 0x0404},
      {"it-IT", 0x0410}, {"nl-NL", 0x0413}, {"de", 0x0407},    {"fr", 0x040C},    {"ko", 0x0412},
      {"th", 0x041E},    {"ru", 0x0419},    {"it", 0x0410},    {"nl", 0x0413},
  };
  for (const NamedLcid& entry : kNames) {
    if (strings::case_insensitive_eq(entry.name, name)) {
      return tag_language_for_lcid(entry.lcid);
    }
  }
  return nullptr;
}

const TagLanguage& native_tag_language(ExcelLocale locale) noexcept {
  switch (locale) {
    case ExcelLocale::kJaJP:
      return kTagJapanese;
    case ExcelLocale::kEnUS:
      return kTagEnglish;
    case ExcelLocale::kDeDE:
      return kTagGerman;
    case ExcelLocale::kFrFR:
      return kTagFrench;
    case ExcelLocale::kZhCN:
      return kTagChinese;
    case ExcelLocale::kKoKR:
      return kTagKorean;
    case ExcelLocale::kThTH:
      return kTagThai;
    case ExcelLocale::kEsES:
      return kTagSpanish;
    case ExcelLocale::kEsMX:
      return kTagMexicanSpanish;
    case ExcelLocale::kPtBR:
      return kTagPortuguese;
    case ExcelLocale::kRuRU:
      return kTagRussian;
    case ExcelLocale::kZhTW:
      return kTagTraditionalChinese;
    case ExcelLocale::kItIT:
      return kTagItalian;
    case ExcelLocale::kNlNL:
      return kTagDutch;
  }
  return kTagEnglish;
}

const DbnumStyle* numeral_system_style(std::uint8_t code) noexcept {
  // 1B-27 are the ja, zh-CN, zh-TW and ko DBNum styles in turn (locale_tokens.lcid_numeral_systems).
  constexpr std::uint8_t kFirstDigitSet = 0x02;
  constexpr std::uint8_t kFirstCjk = 0x1B;
  if (code >= kFirstDigitSet && code < kFirstDigitSet + kNumeralSystemStyles.size()) {
    return &kNumeralSystemStyles[code - kFirstDigitSet];
  }
  if (code >= kFirstCjk && code < kFirstCjk + 3U) {
    return &kJapaneseDbnumStyles[code - kFirstCjk];
  }
  if (code >= kFirstCjk + 3U && code < kFirstCjk + 6U) {
    return &kChineseDbnumStyles[code - kFirstCjk - 3U];
  }
  if (code >= kFirstCjk + 6U && code < kFirstCjk + 9U) {
    return &kTraditionalChineseDbnumStyles[code - kFirstCjk - 6U];
  }
  if (code >= kFirstCjk + 9U && code < kFirstCjk + 13U) {
    return &kKoreanDbnumStyles[code - kFirstCjk - 9U];
  }
  return nullptr;
}

std::string_view system_date_format(ExcelProfile profile) noexcept {
  // locale_tokens.text_system_date_tag; ru-RU writes genitive month names through 0xFC19.
  switch (profile.locale) {
    case ExcelLocale::kJaJP:
      return "[$-411]yyyy\"年\"m\"月\"d\"日\" dddd";
    case ExcelLocale::kEnUS:
      return "[$-409]dddd\", \"mmmm\" \"d\", \"yyyy";
    case ExcelLocale::kDeDE:
      return "[$-407]dddd\", \"d\". \"mmmm\" \"yyyy";
    case ExcelLocale::kFrFR:
      return "[$-40C]dddd\" \"d\" \"mmmm\" \"yyyy";
    case ExcelLocale::kZhCN:
      return "[$-804]yyyy\"年\"m\"月\"d\"日\" dddd";
    case ExcelLocale::kKoKR:
      return "[$-412]yyyy\"년\" m\"월\" d\"일\" dddd";
    case ExcelLocale::kThTH:
      return "[$-41E]dddd\"ที่ \"d mmmm\"  \"yyyy";
    case ExcelLocale::kEsES:
      return "[$-C0A]dddd\", \"d\" de \"mmmm\" de \"yyyy";
    case ExcelLocale::kEsMX:
      return "[$-80A]dddd\", \"d\" de \"mmmm\" de \"yyyy";
    case ExcelLocale::kPtBR:
      return "[$-416]dddd\", \"d\" de \"mmmm\" de \"yyyy";
    case ExcelLocale::kRuRU:
      return "[$-FC19]dddd\", \"d\" \"mmmm\" \"yyyy\" г.\"";
    case ExcelLocale::kZhTW:
      return "[$-404]yyyy\"年\"m\"月\"d\"日\" dddd";
    case ExcelLocale::kItIT:
      return "[$-410]dddd\" \"d\" \"mmmm\" \"yyyy";
    case ExcelLocale::kNlNL:
      return "[$-413]dddd\" \"d\" \"mmmm\" \"yyyy";
  }
  return {};
}

std::string_view system_time_format(ExcelProfile profile) noexcept {
  // locale_tokens.text_system_time_tag.
  switch (profile.locale) {
    case ExcelLocale::kJaJP:
    case ExcelLocale::kEsES:
      return "h:mm:ss";
    case ExcelLocale::kEnUS:
      return "[$-409]h:mm:ss AM/PM";
    case ExcelLocale::kEsMX:
      return "[$-80A]h:mm:ss AM/PM";
    case ExcelLocale::kZhCN:
      return "\"z\"hh:mm:ss";
    case ExcelLocale::kZhTW:
      return "\"z B\"h:mm:ss";
    case ExcelLocale::kKoKR:
      return "[$-412]AM/PM h\"시\" m\"분\" s\"초\"";
    case ExcelLocale::kThTH:
      return "h\" นาฬิกา \"mm\" นาที \"ss\" วินาที\"";
    case ExcelLocale::kDeDE:
    case ExcelLocale::kFrFR:
    case ExcelLocale::kPtBR:
    case ExcelLocale::kRuRU:
    case ExcelLocale::kItIT:
    case ExcelLocale::kNlNL:
      return "hh:mm:ss";
  }
  return {};
}

}  // namespace formulon
