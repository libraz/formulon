//
// Implementation of the conditional-format rule properties. See cf_rule_traits.h.

#include "cf/cf_rule_traits.h"

#include <initializer_list>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "cf/cf_types.h"
#include "parser/parser.h"

namespace formulon::cf {
namespace {

bool CfvoReferencesOtherSheet(const CfValueObject& v, std::string_view own_sheet,
                              const std::vector<std::string>* sheet_list) {
  return v.type == CfvoType::Formula && parser::formula_references_other_sheet(v.value, own_sheet, sheet_list);
}

}  // namespace

bool rule_is_x14_only(const CFRule& rule, std::string_view own_sheet, const std::vector<std::string>* sheet_list) {
  if (own_sheet.empty()) {
    return false;
  }
  for (const std::optional<std::string>* formula : {&rule.formula1, &rule.formula2}) {
    if (formula->has_value() && parser::formula_references_other_sheet(**formula, own_sheet, sheet_list)) {
      return true;
    }
  }
  if (rule.color_scale.has_value()) {
    for (const CfValueObject& v : rule.color_scale->thresholds) {
      if (CfvoReferencesOtherSheet(v, own_sheet, sheet_list)) {
        return true;
      }
    }
  }
  if (rule.data_bar.has_value() && (CfvoReferencesOtherSheet(rule.data_bar->min, own_sheet, sheet_list) ||
                                    CfvoReferencesOtherSheet(rule.data_bar->max, own_sheet, sheet_list))) {
    return true;
  }
  if (rule.icon_set.has_value()) {
    if (CfvoReferencesOtherSheet(rule.icon_set->floor, own_sheet, sheet_list)) {
      return true;
    }
    for (const CfValueObject& v : rule.icon_set->thresholds) {
      if (CfvoReferencesOtherSheet(v, own_sheet, sheet_list)) {
        return true;
      }
    }
  }
  return false;
}

}  // namespace formulon::cf
