//
// Locale-driven labelling for the pivot-grid layout layer. See the
// header for the contract.

#include "pivot/pivot_locale.h"

#include <array>
#include <cstddef>
#include <string>
#include <string_view>

#include "excel_profile.h"
#include "pivot/pivot_layout.h"
#include "pivot/pivot_types.h"

namespace formulon::pivot {
namespace {

constexpr std::size_t kAggregationCount = 11;

using AggregationLabels = std::array<std::string_view, kAggregationCount>;

// Labels per `pivot::Aggregation`, in enum order, as Mac Excel 365 names a
// data field after its function is switched in each UI locale. Only ja-JP
// gives CountNumbers its own label; every other locale reuses Count's.
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

// Indexed by `ExcelLocale`. Measured on Mac Excel 365 by switching the UI
// locale and reading what Excel renders and names. `values_label` is the
// dataCaption Excel stores when a pivot is created in that UI.
constexpr std::array<PivotLocaleLabels, 7> kLocaleLabels{{
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
}};

static_assert(kLocaleLabels.size() == static_cast<std::size_t>(ExcelLocale::kThTH) + 1U,
              "one label set per ExcelLocale");

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
  return options;
}

std::string_view aggregation_label(pivot::Aggregation agg, ExcelProfile profile) {
  return label_at(locale_labels(profile.locale).aggregation_labels, agg);
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

}  // namespace formulon::pivot
