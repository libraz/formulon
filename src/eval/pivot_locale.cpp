//
// Locale-driven labelling for the pivot-grid layout layer. See the
// header for the contract.

#include "eval/pivot_locale.h"

#include <array>
#include <cstddef>
#include <string>
#include <string_view>

#include "eval/compat.h"
#include "pivot/pivot_layout.h"
#include "pivot/pivot_types.h"

namespace formulon::eval {
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

bool is_ja_jp(ExcelProfile profile) noexcept {
  return profile.locale == ExcelLocale::kJaJP;
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
  pivot::PivotLayoutOptions options;
  if (is_ja_jp(profile)) {
    options.grand_total_label = "総計";
    options.values_label = "値";
    options.row_labels_label = "行ラベル";
    options.column_labels_label = "列ラベル";
    options.subtotal_suffix = " 集計";
    // Provisional spelling: unlike the labels above, this one is not backed
    // by an Excel observation yet. Pinning it needs a ja-JP capture of a
    // pivot whose row field contains empty source cells.
    options.blank_item_label = "(空白)";
  }
  return options;
}

std::string_view aggregation_label(pivot::Aggregation agg, ExcelProfile profile) {
  return is_ja_jp(profile) ? label_at(kJaJpAggregationLabels, agg) : label_at(kEnglishAggregationLabels, agg);
}

std::string_view data_field_separator(ExcelProfile profile) {
  return is_ja_jp(profile) ? " / " : " of ";
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

}  // namespace formulon::eval
