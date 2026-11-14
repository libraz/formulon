//
// The position unit text functions count in under the evaluated workbook.

#ifndef FORMULON_EVAL_TEXT_INDEX_UNIT_H_
#define FORMULON_EVAL_TEXT_INDEX_UNIT_H_

#include "eval/eval_profile_scope.h"
#include "excel_profile.h"
#include "utils/text_ops.h"

namespace formulon::eval {

/// The position unit of LEN, MID, FIND, SEARCH and REPLACE: the evaluated
/// workbook's Compatibility Version (version 1 counts UTF-16 units, every
/// other version code points), set by `ScopedEvalProfile`.
inline TextIndexUnit text_index_unit() noexcept {
  return current_text_index_unit();
}

}  // namespace formulon::eval

#endif  // FORMULON_EVAL_TEXT_INDEX_UNIT_H_
