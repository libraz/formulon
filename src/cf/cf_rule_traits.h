//
// Properties of a conditional-format rule that the structural edits and the
// readers and writers all decide by, kept in one definition.

#ifndef FORMULON_CF_CF_RULE_TRAITS_H_
#define FORMULON_CF_CF_RULE_TRAITS_H_

#include <string>
#include <string_view>
#include <vector>

#include "cf/cf_types.h"

namespace formulon::cf {

/// True when `rule` is stored only in the x14 extension: a formula or `cfvo`
/// of it names a sheet other than `own_sheet`, spellings resolved among
/// `sheet_list`. An empty `own_sheet` places no rule there.
bool rule_is_x14_only(const CFRule& rule, std::string_view own_sheet, const std::vector<std::string>* sheet_list);

/// True when an x14-only `rule` keeps `stopIfTrue` across an Excel save. Excel drops it from a data bar and from a
/// stock icon set stored only in the x14 extension, so both writers leave it out for those.
bool x14_only_rule_carries_stop_if_true(const CFRule& rule);

}  // namespace formulon::cf

#endif  // FORMULON_CF_CF_RULE_TRAITS_H_
