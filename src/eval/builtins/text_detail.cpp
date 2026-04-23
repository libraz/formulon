// Copyright 2026 libraz. Licensed under the MIT License.
//
// Shared definition of `read_int_arg` for the text builtin family. Hoisted
// out of `text.cpp` into its own TU because the DBCS family (`text_dbcs.cpp`)
// and the modern TEXTBEFORE/TEXTAFTER family (`text_modern.cpp`) both
// reference it; keeping a single definition here avoids ODR violations.

#include "eval/builtins/text_detail.h"

#include <cmath>

#include "eval/coerce.h"

namespace formulon {
namespace eval {
namespace text_detail {

Expected<int, ErrorCode> read_int_arg(const Value& v) {
  auto coerced = coerce_to_number(v);
  if (!coerced) {
    return coerced.error();
  }
  const double d = coerced.value();
  if (std::isnan(d) || std::isinf(d)) {
    return ErrorCode::Num;
  }
  return static_cast<int>(std::trunc(d));
}

}  // namespace text_detail
}  // namespace eval
}  // namespace formulon
