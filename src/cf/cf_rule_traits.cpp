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

bool x14_only_rule_carries_stop_if_true(const CFRule& rule) {
  if (rule.type == RuleType::DataBar) {
    return false;
  }
  return !(rule.type == RuleType::IconSet && rule.icon_set.has_value() && rule.icon_set->custom_icons.empty());
}

bool far_edge_edit_splits_type(RuleType type) {
  return type == RuleType::Top10 || type == RuleType::AboveAverage || type == RuleType::DuplicateValues ||
         type == RuleType::UniqueValues;
}

bool far_edge_edit_splits_threshold(CfvoType type) {
  return type == CfvoType::Min || type == CfvoType::Max || type == CfvoType::Percent || type == CfvoType::AutoMin ||
         type == CfvoType::AutoMax;
}

bool far_edge_edit_splits_rule(const CFRule& rule) {
  if (far_edge_edit_splits_type(rule.type)) {
    return true;
  }
  if (rule.color_scale.has_value()) {
    for (const CfValueObject& v : rule.color_scale->thresholds) {
      if (far_edge_edit_splits_threshold(v.type)) {
        return true;
      }
    }
  }
  if (rule.data_bar.has_value() && (far_edge_edit_splits_threshold(rule.data_bar->min.type) ||
                                    far_edge_edit_splits_threshold(rule.data_bar->max.type))) {
    return true;
  }
  if (rule.icon_set.has_value()) {
    if (far_edge_edit_splits_threshold(rule.icon_set->floor.type)) {
      return true;
    }
    for (const CfValueObject& v : rule.icon_set->thresholds) {
      if (far_edge_edit_splits_threshold(v.type)) {
        return true;
      }
    }
  }
  return false;
}

}  // namespace formulon::cf
