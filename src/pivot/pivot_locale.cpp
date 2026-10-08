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

/// Excel 365 ja-JP labels per `pivot::Aggregation`, in enum order.
///
/// These follow Mac/Win Excel 365 ja-JP exactly: in particular `StdDev`
/// (sample) localises as "標本標準偏差" rather than the dictionary
/// "標準偏差", and `VarP` (population variance) renders as plain "分散"
/// matching the observed UI. Keep in sync with `pivot::Aggregation`.
constexpr std::array<std::string_view, kAggregationCount> kJaJpAggregationLabels{
    "合計",          // Sum
    "個数",          // Count
    "平均",          // Average
    "最大",          // Max
    "最小",          // Min
    "積",            // Product
    "数値の個数",    // CountNumbers
    "標本標準偏差",  // StdDev (sample)
    "標準偏差",      // StdDevP (population)
    "標本分散",      // Var (sample)
    "分散",          // VarP (population)
};

/// English labels mirroring Excel's pivot UI ("Sum", "Count", ...).
/// Combined with `data_field_separator` (" of ") this reproduces the
/// historical "Sum of <field>" / "CountNumbers of <field>" wording
/// used by the workbook-oracle harness and the OOXML round-trip
/// fixtures, so default-locale callers see no churn.
constexpr std::array<std::string_view, kAggregationCount> kEnglishAggregationLabels{
    "Sum",      // Sum
    "Count",    // Count
    "Average",  // Average
    "Max",      // Max
    "Min",      // Min
    "Product",  // Product
    "CountNumbers", "StdDev", "StdDevP", "Var", "VarP",
};

struct PivotLocaleLabels {
  std::array<std::string_view, kAggregationCount> aggregation_labels;
  std::string_view grand_total_label;
  std::string_view values_label;
  std::string_view row_labels_label;
  std::string_view column_labels_label;
  std::string_view subtotal_suffix;
  std::string_view blank_item_label;
  std::string_view all_pages_label;
  std::string_view multiple_items_label;
  std::string_view data_field_separator;
};

constexpr std::array<PivotLocaleLabels, 2> kLocaleLabels{{
    {kJaJpAggregationLabels, "総計", "値", "行ラベル", "列ラベル", " 集計", "(空白)", "(すべて)", "(複数のアイテム)",
     " / "},
    {kEnglishAggregationLabels, "Grand Total", "Values", "", "", "", "(blank)", "(All)", "(Multiple Items)", " of "},
}};

const PivotLocaleLabels& locale_labels(ExcelLocale locale) noexcept {
  const auto index = static_cast<std::size_t>(locale);
  if (index < kLocaleLabels.size()) {
    return kLocaleLabels[index];
  }
  return kLocaleLabels[static_cast<std::size_t>(ExcelLocale::kEnUS)];
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
  options.subtotal_suffix = labels.subtotal_suffix;
  options.blank_item_label = labels.blank_item_label;
  options.all_pages_label = labels.all_pages_label;
  options.multiple_items_label = labels.multiple_items_label;
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
