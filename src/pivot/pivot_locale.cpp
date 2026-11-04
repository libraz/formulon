// @size-budget: 18 KB
//
// Locale-driven labelling for the pivot-grid layout layer. See the
// header for the contract.

#include "pivot/pivot_locale.h"

#include <array>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>

#include "eval/collation.h"
#include "excel_profile.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_layout.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"

namespace formulon::pivot {
namespace {

constexpr std::size_t kAggregationCount = 11;

using AggregationLabels = std::array<std::string_view, kAggregationCount>;

// Labels per `pivot::Aggregation`, in enum order, as Mac Excel 365 names a
// data field after its function is switched in each UI locale. Only ja-JP
// gives CountNumbers its own label; every other locale reuses Count's.
// es-MX shares the es-ES labels.
constexpr AggregationLabels kJaJpAggregationLabels{
    "合計", "個数", "平均", "最大", "最小", "積", "数値の個数", "標本標準偏差", "標準偏差", "標本分散", "分散",
};
constexpr AggregationLabels kEnUsAggregationLabels{
    "Sum", "Count", "Average", "Max", "Min", "Product", "Count", "StdDev", "StdDevp", "Var", "Varp",
};
constexpr AggregationLabels kDeDeAggregationLabels{
    "Summe",
    "Anzahl",
    "Mittelwert",
    "Max.",
    "Min.",
    "Produkt",
    "Anzahl",
    "STABW",
    "Standardabweichung (Grundgesamtheit)",
    "Var",
    "Varianz (Grundgesamtheit)",
};
constexpr AggregationLabels kFrFrAggregationLabels{
    "Somme", "Nombre", "Moyenne", "Max.", "Min.", "Produit", "Nombre", "Écartype", "Écartypep", "Var", "Varp",
};
constexpr AggregationLabels kZhCnAggregationLabels{
    "求和项", "计数项",     "平均值项",       "最大值项", "最小值项",   "乘积项",
    "计数项", "标准偏差项", "总体标准偏差项", "方差项",   "总体方差项",
};
// A custom subtotal row names its function without the data-field caption's
// trailing 项.
constexpr AggregationLabels kZhCnSubtotalFunctionLabels{
    "求和", "计数", "平均值", "最大值", "最小值", "乘积", "计数", "标准偏差", "总体标准偏差", "方差", "总体方差",
};
constexpr AggregationLabels kKoKrAggregationLabels{
    "합계", "개수", "평균", "최대", "최소", "곱", "개수", "표본 표준 편차", "표준 편차", "표본 분산", "분산",
};
constexpr AggregationLabels kThThAggregationLabels{
    "ผลรวม",
    "นับจำนวน",
    "ค่าเฉลี่ย",
    "สูงสุด",
    "ต่ำสุด",
    "ผลคูณ",
    "นับจำนวน",
    "ส่วนเบี่ยงเบนมาตรฐาน",
    "ส่วนเบี่ยงเบนมาตรฐานของประชากร",
    "ค่าความแปรปรวน",
    "ค่าความแปรปรวนของประชากร",
};

constexpr AggregationLabels kEsEsAggregationLabels{
    "Suma", "Cuenta", "Promedio", "Máx.", "Mín.", "Producto", "Cuenta", "Desvest", "Desvestp", "Var", "Varp",
};
constexpr AggregationLabels kPtBrAggregationLabels{
    "Soma", "Contagem", "Média", "Máx.", "Mín.", "Produto", "Contagem", "DesvPad", "DesvPadp", "Var.", "Varp",
};
constexpr AggregationLabels kRuRuAggregationLabels{
    "Сумма",
    "Количество",
    "Среднее",
    "Максимум",
    "Минимум",
    "Произведение",
    "Количество",
    "Стандартное отклонение",
    "Несмещенное отклонение",
    "Дисперсия",
    "Несмещенная дисперсия",
};
constexpr AggregationLabels kZhTwAggregationLabels{
    "加總", "計數", "平均值", "最大", "最小", "乘積", "計數", "標準差", "母體標準差", "變異數", "母體變異值",
};
constexpr AggregationLabels kItItAggregationLabels{
    "Somma",         "Conteggio",          "Media",    "Max",           "Min", "Prodotto", "Conteggio",
    "Dev. standard", "Dev. standard pop.", "Varianza", "Varianza pop.",
};
constexpr AggregationLabels kNlNlAggregationLabels{
    "Som", "Aantal", "Gemiddelde", "Max", "Min", "Product", "Aantal", "Stdev", "Stdevp", "Var", "VARP",
};

struct PivotLocaleLabels {
  AggregationLabels aggregation_labels;
  std::string_view grand_total_label;
  std::string_view values_label;
  std::string_view row_labels_label;
  std::string_view column_labels_label;
  std::string_view subtotal_prefix;
  std::string_view subtotal_suffix;
  std::string_view blank_item_label;
  std::string_view all_pages_label;
  std::string_view multiple_items_label;
  std::string_view data_field_separator;
  std::string_view data_total_prefix;
  std::string_view data_total_suffix;
};

// Indexed by `ExcelLocale`. Mac Excel 365 values per UI locale: what Excel
// renders and names. `values_label` is the dataCaption Excel stores when a
// pivot is created in that UI.
constexpr std::array<PivotLocaleLabels, 14> kLocaleLabels{{
    {kJaJpAggregationLabels, "総計", "値", "行ラベル", "列ラベル", "", " 集計", "(空白)", "(すべて)",
     "(複数のアイテム)", " / ", "全体の ", ""},
    {kEnUsAggregationLabels, "Grand Total", "Values", "Row Labels", "Column Labels", "", " Total", "(blank)", "(All)",
     "(Multiple Items)", " of ", "Total ", ""},
    {kDeDeAggregationLabels, "Gesamtergebnis", "Werte", "Zeilenbeschriftungen", "Spaltenbeschriftungen", "",
     " Ergebnis", "(Leer)", "(Alle)", "(Mehrere Elemente)", " von ", "Gesamt: ", ""},
    {kFrFrAggregationLabels, "Total général", "Valeurs", "Étiquettes de lignes", "Étiquettes de colonnes", "Total ", "",
     "(vide)", "(Tous)", "(Plusieurs éléments)", " de ", "Total ", ""},
    {kZhCnAggregationLabels, "总计", "值", "行标签", "列标签", "", " 汇总", "(空白)", "(全部)", "(多项)", ":", "",
     "汇总"},
    {kKoKrAggregationLabels, "총합계", "값", "행 레이블", "열 레이블", "", " 요약", "(비어 있음)", "(모두)",
     "(다중 항목)", " : ", "전체 ", ""},
    {kThThAggregationLabels, "ผลรวมทั้งหมด", "ค่า", "ป้ายชื่อแถว", "ป้ายชื่อคอลัมน์", "", " ผลรวม", "(ว่าง)", "(ทั้งหมด)",
     "(หลายรายการ)", " ของ ", "ผลรวม ", ""},
    {kEsEsAggregationLabels, "Total general", "Valores", "Etiquetas de fila", "Etiquetas de columna", "Total ", "",
     "(en blanco)", "(Todas)", "(Varios elementos)", " de ", "Total ", ""},
    {kEsEsAggregationLabels, "Total general", "Valores", "Etiquetas de fila", "Etiquetas de columna", "Total ", "",
     "(en blanco)", "(Todas)", "(Varios elementos)", " de ", "Total ", ""},
    {kPtBrAggregationLabels, "Total Geral", "Valores", "Rótulos de Linha", "Rótulos de Coluna", "", " Total", "(vazio)",
     "(Tudo)", "(Vários itens)", " de ", "Total ", ""},
    {kRuRuAggregationLabels, "Общий итог", "Значения", "Названия строк", "Названия столбцов", "", " Итог", "(пусто)",
     "(Все)", "(несколько элементов)", " по полю ", "Итог ", ""},
    {kZhTwAggregationLabels, "總計", "值", "列標籤", "欄標籤", "", " 合計", "(空白)", "(全部)", "(多重項目)", " - ", "",
     " 的加總"},
    {kItItAggregationLabels, "Totale complessivo", "Valori", "Etichette di riga", "Etichette di colonna", "", " Totale",
     "(vuoto)", "(Tutto)", "(più elementi)", " di ", "", " totale"},
    {kNlNlAggregationLabels, "Eindtotaal", "Waarden", "Rijlabels", "Kolomlabels", "Totaal ", "", "(leeg)", "(Alle)",
     "(Meerdere items)", " van ", "Totaal ", ""},
}};

static_assert(kLocaleLabels.size() == static_cast<std::size_t>(ExcelLocale::kNlNL) + 1U,
              "one label set per ExcelLocale");

// How a group item of the hour unit is written.
enum class ClockStyle : std::uint8_t {
  kSuffix,      // "9時"
  kTwelveHour,  // "9 AM", "12 PM"
  kPlain,       // "9"
  kPadded,      // "09"
};

// Words Mac Excel 365 writes into the group items and group-field names it
// creates when a date field is grouped in each UI locale. A minute or second
// item without a suffix is written ":05".
struct PivotDateWords {
  std::array<std::string_view, 12> months;
  std::array<std::string_view, 4> quarters;
  std::string_view year_suffix;
  bool day_first;
  bool pad_day;
  std::string_view day_separator;
  std::string_view day_suffix;
  ClockStyle hour_style;
  std::string_view hour_suffix;
  std::string_view am;
  std::string_view pm;
  std::string_view minute_suffix;
  std::string_view second_suffix;
  // Seconds through years, as `GroupBy` numbers them from 1.
  std::array<std::string_view, 7> period_words;
  std::string_view field_name_open;
};

// Indexed by `ExcelLocale`.
constexpr std::array<PivotDateWords, 14> kDateWords{{
    // ja_JP
    {{"1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"},
     {"第 1 四半期", "第 2 四半期", "第 3 四半期", "第 4 四半期"},
     "年",
     false,
     false,
     "",
     "日",
     ClockStyle::kSuffix,
     "時",
     "",
     "",
     "分",
     "秒",
     {"秒", "分", "時間", "日", "月", "四半期", "年"},
     " ("},
    // en_US
    {{"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"},
     {"Qtr1", "Qtr2", "Qtr3", "Qtr4"},
     "",
     true,
     false,
     "-",
     "",
     ClockStyle::kTwelveHour,
     "",
     " AM",
     " PM",
     "",
     "",
     {"Seconds", "Minutes", "Hours", "Days", "Months", "Quarters", "Years"},
     " ("},
    // de_DE
    {{"Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"},
     {"Qrtl1", "Qrtl2", "Qrtl3", "Qrtl4"},
     "",
     true,
     true,
     ". ",
     "",
     ClockStyle::kPadded,
     "",
     "",
     "",
     "",
     "",
     {"Sekunden", "Minuten", "Stunden", "Tage", "Monate", "Quartale", "Jahre"},
     " ("},
    // fr_FR
    {{"janv", "févr", "mars", "avr", "mai", "juin", "juil", "août", "sept", "oct", "nov", "déc"},
     {"Trimestre1", "Trimestre2", "Trimestre3", "Trimestre4"},
     "",
     true,
     true,
     "-",
     "",
     ClockStyle::kPadded,
     "",
     "",
     "",
     "",
     "",
     {"Secondes", "Minutes", "Heures", "Jours", "Mois", "Trimestres", "Années"},
     " ("},
    // zh_CN
    {{"1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"},
     {"第一季", "第二季", "第三季", "第四季"},
     "年",
     false,
     false,
     "",
     "日",
     ClockStyle::kSuffix,
     "时",
     "",
     "",
     "分",
     "秒",
     {"秒", "分钟", "小时", "天", "月", "季度", "年"},
     "("},
    // ko_KR
    {{"1월", "2월", "3월", "4월", "5월", "6월", "7월", "8월", "9월", "10월", "11월", "12월"},
     {"1사분기", "2사분기", "3사분기", "4사분기"},
     "년",
     false,
     false,
     "",
     "일",
     ClockStyle::kSuffix,
     "시",
     "",
     "",
     "분",
     "초",
     {"초", "분", "시간", "일", "개월", "분기", "년"},
     "("},
    // th_TH
    {{"ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."},
     {"ไตรมาส1", "ไตรมาส2", "ไตรมาส3", "ไตรมาส4"},
     "",
     true,
     false,
     "-",
     "",
     ClockStyle::kPadded,
     "",
     "",
     "",
     "",
     "",
     {"วินาที", "นาที", "ชั่วโมง", "วัน", "เดือน", "ไตรมาส", "ปี"},
     " ("},
    // es_ES
    {{"ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sept", "oct", "nov", "dic"},
     {"Trim.1", "Trim.2", "Trim.3", "Trim.4"},
     "",
     true,
     false,
     "-",
     "",
     ClockStyle::kPlain,
     "",
     "",
     "",
     "",
     "",
     {"Segundos", "Minutos", "Horas", "Días", "Meses", "Trimestres", "Años"},
     " ("},
    // es_MX
    {{"ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"},
     {"Trim.1", "Trim.2", "Trim.3", "Trim.4"},
     "",
     true,
     true,
     "-",
     "",
     ClockStyle::kTwelveHour,
     "",
     " a.m.",
     " p.m.",
     "",
     "",
     {"Segundos", "Minutos", "Horas", "Días", "Meses", "Trimestres", "Años"},
     " ("},
    // pt_BR
    {{"jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"},
     {"Trim1", "Trim2", "Trim3", "Trim4"},
     "",
     true,
     true,
     "/",
     "",
     ClockStyle::kPadded,
     "",
     "",
     "",
     "",
     "",
     {"Segundos", "Minutos", "Horas", "Dias", "Meses", "Trimestres", "Anos"},
     " ("},
    // ru_RU
    {{"янв", "февр", "март", "апр", "май", "июнь", "июль", "авг", "сент", "окт", "нояб", "дек"},
     {"Кв-л1", "Кв-л2", "Кв-л3", "Кв-л4"},
     "",
     true,
     true,
     ".",
     "",
     ClockStyle::kPadded,
     "",
     "",
     "",
     "",
     "",
     {"Секунды", "Минуты", "Часы", "Дни", "Месяцы", "Кварталы", "Годы"},
     " ("},
    // zh_TW
    {{"1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"},
     {"第一季", "第二季", "第三季", "第四季"},
     "年",
     false,
     false,
     "",
     "日",
     ClockStyle::kSuffix,
     "時",
     "",
     "",
     "分",
     "秒",
     {"秒", "分鐘", "小時", "天", "月", "季", "年"},
     " ("},
    // it_IT
    {{"gen", "feb", "mar", "apr", "mag", "giu", "lug", "ago", "set", "ott", "nov", "dic"},
     {"Trim1", "Trim2", "Trim3", "Trim4"},
     "",
     true,
     true,
     "-",
     "",
     ClockStyle::kPadded,
     "",
     "",
     "",
     "",
     "",
     {"Secondi", "Minuti", "Ore", "Giorni", "Mesi", "Trimestri", "Anni"},
     " ("},
    // nl_NL
    {{"jan", "feb", "mrt", "apr", "mei", "jun", "jul", "aug", "sep", "okt", "nov", "dec"},
     {"Kwrt1", "Kwrt2", "Kwrt3", "Kwrt4"},
     "",
     true,
     true,
     "-",
     "",
     ClockStyle::kPadded,
     "",
     "",
     "",
     "",
     "",
     {"Seconden", "Minuten", "Uren", "Dagen", "Maanden", "Kwartalen", "Jaren"},
     " ("},
}};

static_assert(kDateWords.size() == static_cast<std::size_t>(ExcelLocale::kNlNL) + 1U, "one word set per ExcelLocale");

const PivotDateWords& date_words(ExcelProfile profile) noexcept {
  // NOLINTNEXTLINE(cppcoreguidelines-pro-bounds-constant-array-index)
  return kDateWords[static_cast<std::size_t>(profile.locale)];
}

// Days before each month in a leap year.
constexpr std::array<unsigned, 13> kLeapDaysBefore{0U,   31U,  60U,  91U,  121U, 152U, 182U,
                                                   213U, 244U, 274U, 305U, 335U, 366U};

void append_two_digits(std::string& out, unsigned n) {
  out.push_back(static_cast<char>('0' + (n / 10U) % 10U));
  out.push_back(static_cast<char>('0' + n % 10U));
}

std::string clock_label(unsigned n, std::string_view suffix) {
  std::string out;
  if (suffix.empty()) {
    out.push_back(':');
    append_two_digits(out, n);
    return out;
  }
  out = std::to_string(n);
  out.append(suffix);
  return out;
}

std::string hour_label(unsigned hour, const PivotDateWords& words) {
  std::string out;
  switch (words.hour_style) {
    case ClockStyle::kSuffix:
      out = std::to_string(hour);
      out.append(words.hour_suffix);
      break;
    case ClockStyle::kTwelveHour:
      out = std::to_string(hour % 12U == 0U ? 12U : hour % 12U);
      out.append(hour < 12U ? words.am : words.pm);
      break;
    case ClockStyle::kPlain:
      out = std::to_string(hour);
      break;
    case ClockStyle::kPadded:
      append_two_digits(out, hour);
      break;
  }
  return out;
}

std::string day_label(unsigned day_of_leap_year, const PivotDateWords& words) {
  unsigned month = 0;
  while (month + 1U < 12U && day_of_leap_year >= kLeapDaysBefore[month + 1U]) {
    ++month;
  }
  // NOLINTNEXTLINE(cppcoreguidelines-pro-bounds-constant-array-index)
  const unsigned day = day_of_leap_year - kLeapDaysBefore[month] + 1U;
  std::string day_text;
  if (words.pad_day) {
    append_two_digits(day_text, day);
  } else {
    day_text = std::to_string(day);
  }
  // NOLINTNEXTLINE(cppcoreguidelines-pro-bounds-constant-array-index)
  const std::string_view month_text = words.months[month];
  std::string out;
  if (words.day_first) {
    out = std::move(day_text);
    out.append(words.day_separator);
    out.append(month_text);
  } else {
    out.assign(month_text);
    out.append(day_text);
    out.append(words.day_suffix);
  }
  return out;
}

const PivotLocaleLabels& locale_labels(ExcelLocale locale) noexcept {
  // NOLINTNEXTLINE(cppcoreguidelines-pro-bounds-constant-array-index)
  return kLocaleLabels[static_cast<std::size_t>(locale)];
}

std::string_view label_at(const std::array<std::string_view, kAggregationCount>& table, pivot::Aggregation agg) {
  const auto idx = static_cast<std::size_t>(agg);
  if (idx >= table.size()) {
    return table.front();
  }
  // The branch above proves `idx < table.size()`, so this subscript is
  // in-range. clang-tidy's constant-array-index check can't see across
  // the guard; suppress just this line rather than reach for `gsl::at`.
  // NOLINTNEXTLINE(cppcoreguidelines-pro-bounds-constant-array-index)
  return table[idx];
}

}  // namespace

pivot::PivotLayoutOptions pivot_layout_options_for(ExcelProfile profile) {
  const PivotLocaleLabels& labels = locale_labels(profile.locale);
  pivot::PivotLayoutOptions options;
  options.grand_total_label = labels.grand_total_label;
  options.values_label = labels.values_label;
  options.row_labels_label = labels.row_labels_label;
  options.column_labels_label = labels.column_labels_label;
  options.subtotal_prefix = labels.subtotal_prefix;
  options.subtotal_suffix = labels.subtotal_suffix;
  options.blank_item_label = labels.blank_item_label;
  options.all_pages_label = labels.all_pages_label;
  options.multiple_items_label = labels.multiple_items_label;
  options.data_total_prefix = labels.data_total_prefix;
  options.data_total_suffix = labels.data_total_suffix;
  options.profile = profile;
  return options;
}

std::string_view aggregation_label(pivot::Aggregation agg, ExcelProfile profile) {
  return label_at(locale_labels(profile.locale).aggregation_labels, agg);
}

std::string_view subtotal_function_label(pivot::Aggregation agg, ExcelProfile profile) {
  if (profile.locale == ExcelLocale::kZhCN) {
    return label_at(kZhCnSubtotalFunctionLabels, agg);
  }
  return aggregation_label(agg, profile);
}

std::string_view data_field_separator(ExcelProfile profile) {
  return locale_labels(profile.locale).data_field_separator;
}

std::string data_field_display_name(pivot::Aggregation agg, std::string_view field_name, ExcelProfile profile) {
  std::string out;
  const std::string_view label = aggregation_label(agg, profile);
  const std::string_view sep = data_field_separator(profile);
  out.reserve(label.size() + sep.size() + field_name.size());
  out.append(label.data(), label.size());
  out.append(sep.data(), sep.size());
  out.append(field_name.data(), field_name.size());
  return out;
}

std::string data_field_default_name(const PivotTable& table, const PivotCache& cache, pivot::Aggregation agg,
                                    std::size_t field_index, ExcelProfile profile,
                                    std::optional<std::size_t> exclude_data_field) {
  std::string_view field_name;
  if (field_index < table.fields().size()) {
    const PivotField& field = table.fields()[field_index];
    field_name = field.custom_name.empty() ? field.source_name : field.custom_name;
  } else if (field_index < cache.fields().size()) {
    field_name = cache.fields()[field_index].name;
  }
  const std::string base = data_field_display_name(agg, field_name, profile);
  const auto taken = [&](std::string_view name) {
    const auto same = [&](std::string_view other) {
      return eval::collate(name, other, profile, eval::CollationStrength::kDynamic) == 0;
    };
    for (std::size_t i = 0; i < table.data_fields().size(); ++i) {
      if (i != exclude_data_field && same(table.data_fields()[i].name)) {
        return true;
      }
    }
    for (std::size_t i = 0; i < cache.fields().size(); ++i) {
      const bool renamed = i < table.fields().size() && !table.fields()[i].custom_name.empty();
      if (same(renamed ? table.fields()[i].custom_name : cache.fields()[i].name)) {
        return true;
      }
    }
    return false;
  };
  std::string name = base;
  for (std::size_t n = 2; taken(name); ++n) {
    name = base + std::to_string(n);
  }
  return name;
}

std::string date_group_field_name(GroupBy period, std::string_view base_name, ExcelProfile profile) {
  const PivotDateWords& words = date_words(profile);
  const auto slot = static_cast<std::size_t>(period);
  std::string out;
  if (slot >= 1U && slot <= words.period_words.size()) {
    // NOLINTNEXTLINE(cppcoreguidelines-pro-bounds-constant-array-index)
    out.assign(words.period_words[slot - 1U]);
  }
  out.append(words.field_name_open);
  out.append(base_name);
  out.push_back(')');
  return out;
}

std::string date_group_unit_label(GroupBy period, unsigned index, ExcelProfile profile) {
  const PivotDateWords& words = date_words(profile);
  switch (period) {
    case GroupBy::Seconds:
      return index < 60U ? clock_label(index, words.second_suffix) : std::string();
    case GroupBy::Minutes:
      return index < 60U ? clock_label(index, words.minute_suffix) : std::string();
    case GroupBy::Hours:
      return index < 24U ? hour_label(index, words) : std::string();
    case GroupBy::Days:
      return index < kLeapDaysBefore.back() ? day_label(index, words) : std::string();
    case GroupBy::Months:
      // NOLINTNEXTLINE(cppcoreguidelines-pro-bounds-constant-array-index)
      return index < words.months.size() ? std::string(words.months[index]) : std::string();
    case GroupBy::Quarters:
      // NOLINTNEXTLINE(cppcoreguidelines-pro-bounds-constant-array-index)
      return index < words.quarters.size() ? std::string(words.quarters[index]) : std::string();
    default:
      return {};
  }
}

std::string date_group_year_label(int year, ExcelProfile profile) {
  std::string out = std::to_string(year);
  out.append(date_words(profile).year_suffix);
  return out;
}

// Every measured locale writes the boundary and interval items alike; only
// the short date inside them is localized.
std::string date_group_before_label(std::string_view start_text) {
  std::string out = "<";
  out.append(start_text);
  return out;
}

std::string date_group_after_label(std::string_view end_text) {
  std::string out = ">";
  out.append(end_text);
  return out;
}

std::string date_group_interval_label(std::string_view first_text, std::string_view last_text) {
  std::string out(first_text);
  out.append(" - ");
  out.append(last_text);
  return out;
}

}  // namespace formulon::pivot
