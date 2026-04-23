// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of Formulon's built-in functions. Each function follows
// the same recipe:
//
//   1. Coerce each argument via the helpers in `eval/coerce.h`.
//   2. Propagate the left-most coercion error.
//   3. Compute the result and finalize (interning text, checking finiteness).
//
// `IF`, `IFERROR`, and `IFNA` are intentionally absent: they short-circuit
// and are special-cased in the tree walker before the registry is consulted.

#include "eval/builtins.h"

#include <cmath>
#include <cstdint>

#include "eval/builtins_aggregate.h"
#include "eval/builtins_datetime.h"
#include "eval/builtins_logical.h"
#include "eval/builtins_math.h"
#include "eval/builtins_math_trig.h"
#include "eval/builtins_stats.h"
#include "eval/builtins_text.h"
#include "eval/coerce.h"
#include "eval/function_registry.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// --- Type predicates ----------------------------------------------------
//
// These functions are registered with `propagate_errors = false` so the
// dispatcher hands them error-typed inputs verbatim. Each returns a Bool
// based purely on the input's `ValueKind` (and, for ISERR / ISNA, the
// specific `ErrorCode`). None of them ever surface a formula error of
// their own beyond the registry's arity check.

// ISNUMBER(value) - true iff the input is a Number cell.
Value IsNumber(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return Value::boolean(args[0].kind() == ValueKind::Number);
}

// ISTEXT(value) - true iff the input is a Text cell.
Value IsText(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return Value::boolean(args[0].kind() == ValueKind::Text);
}

// ISBLANK(value) - true iff the input is the Blank scalar. Empty text
// (`""`) is NOT blank in Excel - `ISBLANK("")` returns FALSE.
Value IsBlank(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return Value::boolean(args[0].kind() == ValueKind::Blank);
}

// ISLOGICAL(value) - true iff the input is a Bool. Numeric 0/1 do not
// count: only the actual TRUE/FALSE booleans qualify.
Value IsLogical(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return Value::boolean(args[0].kind() == ValueKind::Bool);
}

// ISERROR(value) - true iff the input is any formula error, including
// `#N/A`. Combined with the cleared `propagate_errors` flag this lets
// callers branch on errors without first wrapping them in IFERROR.
Value IsError(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return Value::boolean(args[0].kind() == ValueKind::Error);
}

// ISERR(value) - true iff the input is a formula error OTHER than `#N/A`.
// `ISERR(#N/A)` is FALSE; `ISERR(#DIV/0!)`, `ISERR(#REF!)`, etc. are TRUE.
Value IsErr(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  const Value& v = args[0];
  if (v.kind() != ValueKind::Error) {
    return Value::boolean(false);
  }
  return Value::boolean(v.as_error() != ErrorCode::NA);
}

// ISNA(value) - true iff the input is exactly `#N/A`. All other errors
// (and all non-error values) yield FALSE.
Value IsNa(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  const Value& v = args[0];
  if (v.kind() != ValueKind::Error) {
    return Value::boolean(false);
  }
  return Value::boolean(v.as_error() == ErrorCode::NA);
}

// --- Coercion-style info functions --------------------------------------
//
// `N` and `T` propagate errors via the dispatcher's default short-circuit
// (no flag override needed). They differ from `VALUE` / generic
// `coerce_to_*` in that they NEVER fail: any non-matching input maps to
// the function's neutral element (0 for `N`, "" for `T`).

// N(value) - coerce to a Number with Excel's narrow rules:
//   - Number          -> the number unchanged
//   - Bool            -> 1.0 for TRUE, 0.0 for FALSE
//   - Text            -> 0.0 ALWAYS (N intentionally does not parse text;
//                        contrast with VALUE, which does)
//   - Blank           -> 0.0
//   - Error           -> propagated by the dispatcher before this body runs
Value N(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  const Value& v = args[0];
  switch (v.kind()) {
    case ValueKind::Number:
      return v;
    case ValueKind::Bool:
      return Value::number(v.as_boolean() ? 1.0 : 0.0);
    case ValueKind::Text:
    case ValueKind::Blank:
      return Value::number(0.0);
    case ValueKind::Error:
      // Unreachable in practice: dispatcher short-circuits errors. Defensive
      // fall-through keeps the switch exhaustive.
      return v;
    case ValueKind::Array:
    case ValueKind::Ref:
    case ValueKind::Lambda:
      return Value::error(ErrorCode::Value);
  }
  return Value::error(ErrorCode::Value);
}

// T(value) - coerce to Text with Excel's narrow rules:
//   - Text            -> the text unchanged
//   - Number / Bool   -> empty string ""
//   - Blank           -> empty string ""
//   - Error           -> propagated by the dispatcher before this body runs
Value T(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  const Value& v = args[0];
  if (v.kind() == ValueKind::Text) {
    return v;
  }
  if (v.kind() == ValueKind::Error) {
    // Unreachable in practice: dispatcher short-circuits errors.
    return v;
  }
  return Value::text({});
}

}  // namespace

void register_builtins(FunctionRegistry& registry) {
  register_aggregate_builtins(registry);
  register_logical_builtins(registry);
  register_math_builtins(registry);
  register_stats_builtins(registry);

  // Exponential / logarithmic / trigonometric.
  register_math_trig_builtins(registry);

  // Text manipulation.
  register_text_builtins(registry);

  // Info / type queries.
  // The IS* family must inspect error-typed arguments verbatim, so each
  // entry clears `propagate_errors` to opt out of the dispatcher's default
  // left-most-error short-circuit. `N` and `T` use the default (errors
  // propagate before the body runs).
  registry.register_function(FunctionDef{"ISNUMBER", 1u, 1u, &IsNumber, /*propagate_errors=*/false});
  registry.register_function(FunctionDef{"ISTEXT", 1u, 1u, &IsText, /*propagate_errors=*/false});
  registry.register_function(FunctionDef{"ISBLANK", 1u, 1u, &IsBlank, /*propagate_errors=*/false});
  registry.register_function(FunctionDef{"ISLOGICAL", 1u, 1u, &IsLogical, /*propagate_errors=*/false});
  registry.register_function(FunctionDef{"ISERROR", 1u, 1u, &IsError, /*propagate_errors=*/false});
  registry.register_function(FunctionDef{"ISERR", 1u, 1u, &IsErr, /*propagate_errors=*/false});
  registry.register_function(FunctionDef{"ISNA", 1u, 1u, &IsNa, /*propagate_errors=*/false});
  registry.register_function(FunctionDef{"N", 1u, 1u, &N});
  registry.register_function(FunctionDef{"T", 1u, 1u, &T});

  // Date / time.
  register_datetime_builtins(registry);
}

}  // namespace eval
}  // namespace formulon
