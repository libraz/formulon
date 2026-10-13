//
// The formula Excel stores, when a conditional format is authored, for the
// rule types whose decision lives in side fields (text family, blanks, errors,
// timePeriod). Readers surface that stored formula, so authoring produces the
// same one and the model is identical whichever way a rule arrived.

#ifndef FORMULON_CF_CF_AUTHORING_H_
#define FORMULON_CF_CF_AUTHORING_H_

#include <optional>
#include <string>

#include "cell.h"
#include "cf/cf_types.h"

namespace formulon::cf {

/// Excel's stored formula for `rule`, relative to `anchor` (the top-left corner
/// of the bounding box of the rule's sqref ranges), with `rule.text` quoted
/// as a string literal. `nullopt` for a type Excel stores no derived formula for, or when
/// the side field it needs (`text` / `time_period`) is absent.
std::optional<std::string> derived_rule_formula(const CFRule& rule, CellAddress anchor);

}  // namespace formulon::cf

#endif  // FORMULON_CF_CF_AUTHORING_H_
