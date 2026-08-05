//
// Shared scalar reduction for the `@` operator and `_xlfn.SINGLE`.

#ifndef FORMULON_EVAL_IMPLICIT_INTERSECTION_H_
#define FORMULON_EVAL_IMPLICIT_INTERSECTION_H_

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

}  // namespace formulon::eval

#endif  // FORMULON_EVAL_IMPLICIT_INTERSECTION_H_
