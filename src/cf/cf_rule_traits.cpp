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

bool CfvoIsRangeRelative(const CfValueObject& v) {
  return v.type == CfvoType::Min || v.type == CfvoType::Max || v.type == CfvoType::Percent ||
         v.type == CfvoType::AutoMin || v.type == CfvoType::AutoMax;
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

bool x14_only_rule_carries_stop_if_true(const CFRule& rule) {
  if (rule.type == RuleType::DataBar) {
    return false;
  }
  return !(rule.type == RuleType::IconSet && rule.icon_set.has_value() && rule.icon_set->custom_icons.empty());
}

bool far_edge_edit_splits_rule(const CFRule& rule) {
  switch (rule.type) {
    case RuleType::Top10:
    case RuleType::AboveAverage:
    case RuleType::DuplicateValues:
    case RuleType::UniqueValues:
      return true;
    default:
      break;
  }
  if (rule.color_scale.has_value()) {
    for (const CfValueObject& v : rule.color_scale->thresholds) {
      if (CfvoIsRangeRelative(v)) {
        return true;
      }
    }
  }
  if (rule.data_bar.has_value() &&
      (CfvoIsRangeRelative(rule.data_bar->min) || CfvoIsRangeRelative(rule.data_bar->max))) {
    return true;
  }
  if (rule.icon_set.has_value()) {
    if (CfvoIsRangeRelative(rule.icon_set->floor)) {
      return true;
    }
    for (const CfValueObject& v : rule.icon_set->thresholds) {
      if (CfvoIsRangeRelative(v)) {
        return true;
      }
    }
  }
  return false;
}

}  // namespace formulon::cf
