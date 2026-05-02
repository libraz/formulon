// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the CF evaluator. See cf_evaluator.h for the
// scoped contract and the staged-PR roadmap.

#include "cf/cf_evaluator.h"

#include "cf/cf_match.h"
#include "cf/cf_types.h"
#include "value.h"

namespace formulon::cf {

bool match_rule(const CFRule& rule, const Value& cell_value) {
  switch (rule.type) {
    case RuleType::ContainsBlanks:
      return cell_value.is_blank();
    case RuleType::NotContainsBlanks:
      return !cell_value.is_blank();
    case RuleType::ContainsErrors:
      return cell_value.is_error();
    case RuleType::NotContainsErrors:
      return !cell_value.is_error();
    // Rule types whose evaluator lands in subsequent PRs return false
    // here so a caller that walks the full rule list does not mis-fire
    // on a partially-implemented engine. The UI is expected to gate on
    // the engine version it links against.
    case RuleType::Expression:
    case RuleType::CellIs:
    case RuleType::ColorScale:
    case RuleType::DataBar:
    case RuleType::IconSet:
    case RuleType::Top10:
    case RuleType::AboveAverage:
    case RuleType::ContainsText:
    case RuleType::NotContainsText:
    case RuleType::BeginsWith:
    case RuleType::EndsWith:
    case RuleType::TimePeriod:
    case RuleType::DuplicateValues:
    case RuleType::UniqueValues:
      return false;
  }
  return false;
}

CFMatch make_match(const CFRule& rule) {
  CFMatch m;
  m.rule_id = rule.id;
  m.priority = rule.priority;
  // PR6 covers only dxf-driven rule types; later PRs widen this
  // dispatch to assign `kind = ColorScale / DataBar / IconSet` and
  // populate the corresponding render payload.
  m.kind = CFMatchKind::DifferentialFormat;
  m.dxf_id = rule.dxf_id;
  return m;
}

}  // namespace formulon::cf
