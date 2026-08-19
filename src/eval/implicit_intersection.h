//
// Shared scalar reduction for the `@` operator and `_xlfn.SINGLE`.

#ifndef FORMULON_EVAL_IMPLICIT_INTERSECTION_H_
#define FORMULON_EVAL_IMPLICIT_INTERSECTION_H_

#include <cstdint>
#include <optional>

#include "parser/reference.h"
#include "value.h"

namespace formulon::eval {

/// Applies implicit intersection to an already-evaluated value that has no
/// static range coordinates left to project. Scalars pass through unchanged;
/// dynamic arrays use their top-left element. Empty/corrupt arrays are not a
/// valid scalar and surface `#VALUE!`.
inline Value implicit_intersect_value(Value value) {
  if (!value.is_array()) {
    return value;
  }
  if (value.as_array_rows() == 0U || value.as_array_cols() == 0U || value.as_array_cells() == nullptr) {
    return Value::error(ErrorCode::Value);
  }
  return value.as_array_cells()[0];
}

/// Projects the formula cell onto a static `Ref:Ref` range, yielding the single
/// cell implicit intersection selects.
///
/// A single-column range takes the formula row, a single-row range takes the
/// formula column, and a 2-D range intersects on both axes. A formula cell
/// outside the relevant span has nothing to project onto and yields
/// `std::nullopt`, which both call sites surface as `#VALUE!`. The result
/// inherits `lhs.sheet`, matching Excel's treatment of the left endpoint as the
/// qualifier for the whole range.
std::optional<parser::Reference> project_implicit_intersection(const parser::Reference& lhs,
                                                               const parser::Reference& rhs, std::uint32_t formula_row,
                                                               std::uint32_t formula_col);

}  // namespace formulon::eval

#endif  // FORMULON_EVAL_IMPLICIT_INTERSECTION_H_
