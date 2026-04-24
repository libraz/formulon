// Copyright 2026 libraz. Licensed under the MIT License.
//
// Strict logical coercion shared by the AND / OR / XOR eager builtins and
// the IFS lazy impl.
//
// Mac Excel 365 uses a *stricter* coercion for these four functions than
// the generic `coerce_to_bool` helper in `eval/coerce.h` (which is used by
// IF / IFERROR / NOT and accepts numeric strings like "0" / "1"):
//
//   * A bare Bool is returned as-is.
//   * A Number coerces as `d != 0` (NaN / Inf -> `#NUM!`).
//   * Blank is Skip for AND / OR / XOR (the per-arg loop skips it).
//     Callers that want "Blank is FALSE" (IFS) treat Skip as false.
//   * Text (trimmed, ASCII case-insensitive):
//       - ""          -> Skip
//       - "TRUE"      -> true
//       - "FALSE"     -> false
//       - anything else (including "0" / "1") -> `#VALUE!`
//   * Error propagates its code.
//   * Array / Ref / Lambda -> `#VALUE!`.
//
// The `out_skip`-style three-state return ({HasValue, Skip, Error}) lets
// AND / OR / XOR distinguish "no value here; don't count this argument"
// from "saw false / saw true". IFS callers collapse Skip to false.

#ifndef FORMULON_EVAL_LOGICAL_COERCE_H_
#define FORMULON_EVAL_LOGICAL_COERCE_H_

#include <cmath>
#include <cstdint>
#include <string_view>

#include "utils/strings.h"
#include "value.h"

namespace formulon {
namespace eval {

enum class LogicalCoerce : std::uint8_t {
  HasValue,
  Skip,
  Error,
};

/// Strict logical coercion for AND / OR / XOR / IFS. See file-level comment
/// for the full rule table. On `HasValue` the coerced bool is written to
/// `*out_bool`; on `Error` the matching code is written to `*out_err`.
inline LogicalCoerce logical_coerce(const Value& v, bool* out_bool, ErrorCode* out_err) noexcept {
  switch (v.kind()) {
    case ValueKind::Bool:
      *out_bool = v.as_boolean();
      return LogicalCoerce::HasValue;
    case ValueKind::Number: {
      const double d = v.as_number();
      if (std::isnan(d) || std::isinf(d)) {
        *out_err = ErrorCode::Num;
        return LogicalCoerce::Error;
      }
      *out_bool = d != 0.0;
      return LogicalCoerce::HasValue;
    }
    case ValueKind::Blank:
      return LogicalCoerce::Skip;
    case ValueKind::Text: {
      const std::string_view trimmed = strings::trim(v.as_text());
      if (trimmed.empty()) {
        return LogicalCoerce::Skip;
      }
      if (strings::case_insensitive_eq(trimmed, "TRUE")) {
        *out_bool = true;
        return LogicalCoerce::HasValue;
      }
      if (strings::case_insensitive_eq(trimmed, "FALSE")) {
        *out_bool = false;
        return LogicalCoerce::HasValue;
      }
      *out_err = ErrorCode::Value;
      return LogicalCoerce::Error;
    }
    case ValueKind::Error:
      *out_err = v.as_error();
      return LogicalCoerce::Error;
    case ValueKind::Array:
    case ValueKind::Ref:
    case ValueKind::Lambda:
      *out_err = ErrorCode::Value;
      return LogicalCoerce::Error;
  }
  *out_err = ErrorCode::Value;
  return LogicalCoerce::Error;
}

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_LOGICAL_COERCE_H_
