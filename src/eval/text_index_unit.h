//
// The position unit text functions count in under the active profile.

#ifndef FORMULON_EVAL_TEXT_INDEX_UNIT_H_
#define FORMULON_EVAL_TEXT_INDEX_UNIT_H_

#include "eval/eval_profile_scope.h"
#include "excel_profile.h"
#include "utils/text_ops.h"

namespace formulon::eval {

/// Mac Excel counts a supplementary-plane character as one position in LEN,
/// LEFT, RIGHT, MID, FIND, SEARCH and REPLACE; Windows counts UTF-16 units.
inline TextIndexUnit text_index_unit() noexcept {
  return current_eval_profile().host == ExcelHost::kMac365 ? TextIndexUnit::kCodePoints : TextIndexUnit::kUtf16Units;
}

}  // namespace formulon::eval

#endif  // FORMULON_EVAL_TEXT_INDEX_UNIT_H_
