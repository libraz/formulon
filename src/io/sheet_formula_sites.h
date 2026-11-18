//
// The formulas a sheet part and a workbook's tables carry, visited in the order Excel numbers ExternSheet entries.

#ifndef FORMULON_IO_SHEET_FORMULA_SITES_H_
#define FORMULON_IO_SHEET_FORMULA_SITES_H_

#include <algorithm>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "cf/cf_types.h"
#include "io/sparkline_groups.h"
#include "sheet.h"
#include "table.h"
#include "utils/formula_text.h"
#include "utils/index_sort.h"
#include "workbook.h"

namespace formulon {
namespace io {

/// Where a formula `ForEachSheetFormula` visits lives.
enum class SheetFormulaSite : std::uint8_t {
  kCell,        // a cell formula in a dynamic-array, CSE or spill form, or a sparkline formula
  kLegacyCell,  // a cell formula the writer encodes in the legacy form
  kRule,        // a conditional-format or data-validation formula
  kOwnRange,    // no formula: the point where the sheet's own range joins the ExternSheet table
};

/// Whether a threshold is measured against the values of the rule's range (min, max, percent, percentile) rather
/// than stated (a number or a formula).
inline bool ThresholdReadsRange(const cf::CfValueObject& v) {
  return v.type != cf::CfvoType::Number && v.type != cf::CfvoType::Formula;
}

/// Whether a conditional-format rule adds its sheet's own range after its formulas: a scale, bar or icon set does
/// when a threshold reads the range, every other kind but an expression or cell-value rule always does.
inline bool RuleAddsOwnRange(const cf::CFRule& rule) {
  if (rule.color_scale) {
    return std::any_of(rule.color_scale->thresholds.begin(), rule.color_scale->thresholds.end(), ThresholdReadsRange);
  }
  if (rule.data_bar) {
    return ThresholdReadsRange(rule.data_bar->min) || ThresholdReadsRange(rule.data_bar->max);
  }
  if (rule.icon_set) {
    return ThresholdReadsRange(rule.icon_set->floor) ||
           std::any_of(rule.icon_set->thresholds.begin(), rule.icon_set->thresholds.end(), ThresholdReadsRange);
  }
  return rule.type != cf::RuleType::Expression && rule.type != cf::RuleType::CellIs;
}

/// Whether a data validation adds its sheet's own range after its formulas: any with a formula, except a list of
/// literal items.
inline bool ValidationAddsOwnRange(const DataValidation& dv) {
  constexpr std::uint8_t kListType = 3;
  if (dv.formula1.empty() && dv.formula2.empty()) {
    return false;
  }
  return dv.type != kListType || strip_formula_prefix(dv.formula1).substr(0, 1) != "\"";
}

/// Calls `visit(text, site)` with the text (no leading `=`) of every formula a
/// sheet part carries, in the order Excel numbers ExternSheet entries: cell
/// formulas in row-major order, then each conditional-format rule's formulas
/// and thresholds, then data-validation formulas, then sparkline formulas.
/// A `kOwnRange` visit (empty text) follows a rule or validation that adds the
/// sheet's own range. The name and ExternSheet tables are both built over this
/// one walk, so every formula the sheet writer encodes finds its entries.
template <typename Visit>
void ForEachSheetFormula(const Sheet& sheet, Visit&& visit) {
  std::vector<std::uint32_t> rows;
  rows.reserve(sheet.rows().size());
  for (const auto& kv : sheet.rows()) {
    rows.push_back(kv.first);
  }
  sort_ascending(rows);
  for (const std::uint32_t row : rows) {
    for (const auto& [col, cell] : sheet.rows().at(row)) {
      const std::string_view body = strip_formula_prefix(cell.formula_text);
      if (!body.empty()) {
        const bool legacy =
            !cell.dynamic_array && cell.array_block_rows == 0U && sheet.spill_region_at_anchor(row, col) == nullptr;
        visit(body, legacy ? SheetFormulaSite::kLegacyCell : SheetFormulaSite::kCell);
      }
    }
  }
  const auto visit_text = [&visit](const std::string& text, SheetFormulaSite site) {
    if (!text.empty()) {
      visit(std::string_view(text), site);
    }
  };
  const auto visit_cfvo = [&visit_text](const cf::CfValueObject& v) {
    if (v.type == cf::CfvoType::Formula) {
      visit_text(v.value, SheetFormulaSite::kRule);
    }
  };
  for (const cf::ConditionalFormat& format : sheet.conditional_formats()) {
    for (const cf::CFRule& rule : format.rules) {
      visit_text(rule.formula1.value_or(std::string()), SheetFormulaSite::kRule);
      visit_text(rule.formula2.value_or(std::string()), SheetFormulaSite::kRule);
      if (rule.color_scale) {
        std::for_each(rule.color_scale->thresholds.begin(), rule.color_scale->thresholds.end(), visit_cfvo);
      }
      if (rule.data_bar) {
        visit_cfvo(rule.data_bar->min);
        visit_cfvo(rule.data_bar->max);
      }
      if (rule.icon_set) {
        visit_cfvo(rule.icon_set->floor);
        std::for_each(rule.icon_set->thresholds.begin(), rule.icon_set->thresholds.end(), visit_cfvo);
      }
      if (RuleAddsOwnRange(rule)) {
        visit(std::string_view(), SheetFormulaSite::kOwnRange);
      }
    }
  }
  for (const DataValidation& dv : sheet.validations()) {
    visit_text(dv.formula1, SheetFormulaSite::kRule);
    visit_text(dv.formula2, SheetFormulaSite::kRule);
    if (ValidationAddsOwnRange(dv)) {
      visit(std::string_view(), SheetFormulaSite::kOwnRange);
    }
  }
  for (const SparklineGroup& group : read_sparkline_groups(sheet.ext_lst_xml())) {
    visit_text(group.date_axis_ref, SheetFormulaSite::kCell);
    for (const SparklineItem& item : group.items) {
      visit_text(item.formula, SheetFormulaSite::kCell);
    }
  }
}

/// Calls `visit` with the text and sheet of every calculated-column and
/// totals-row formula of `wb`'s tables.
template <typename Visit>
void ForEachTableFormula(const Workbook& wb, Visit&& visit) {
  for (const TableMetadata& table : wb.tables()) {
    for (const TableColumn& column : table.columns) {
      for (const std::string* text : {&column.calculated_column_formula, &column.totals_row_formula}) {
        if (!text->empty()) {
          visit(std::string_view(*text), table.sheet_index);
        }
      }
    }
  }
}

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_SHEET_FORMULA_SITES_H_
