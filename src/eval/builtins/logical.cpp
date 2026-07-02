// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of Formulon's logical built-in functions: TRUE, FALSE, NOT,
// AND, OR, and XOR. Each impl follows the same recipe as the rest of the
// builtin catalog: coerce arguments via `eval/coerce.h`, propagate the
// left-most coercion error, and return a `Value`.

#include "eval/builtins/logical.h"

#include <cstdint>

#include "eval/builtins/registration_helpers.h"
#include "eval/coerce.h"
#include "eval/function_registry.h"
#include "eval/logical_coerce.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// TRUE() / FALSE() -------------------------------------------------------
// Both are zero-argument constants. Excel rejects any argument with #VALUE!,
// which the registry's arity check enforces (min=max=0). The body simply
// returns the corresponding boolean.
Value True_(const Value* /*args*/, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return Value::boolean(true);
}

Value False_(const Value* /*args*/, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return Value::boolean(false);
}

// NOT(value) -------------------------------------------------------------
// Coerces the single argument to bool and negates. Errors propagate (the
// dispatcher already short-circuits on argument errors before invoking
// this body, so by the time we run the input is non-error). A coercion
// failure (e.g. non-numeric text) surfaces as #VALUE!.
Value Not(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto coerced = coerce_to_bool(args[0]);
  if (!coerced) {
    return Value::error(coerced.error());
  }
  return Value::boolean(!coerced.value());
}

// XOR(value, ...) --------------------------------------------------------
// Excel's stricter logical coercion (see `logical_coerce` above):
//
//   * The literal strings "TRUE" / "FALSE" (case-insensitive, trimmed)
//     and bool / finite-number arguments carry a bool value.
//   * An empty / whitespace-only string is skipped ("no value here").
//   * Any other text — including numeric strings "0" / "1" — surfaces
//     `#VALUE!`.
//   * Errors propagate from the left-most failure (the dispatcher
//     short-circuits before entering this body for range-provided errors;
//     scalar errors arrive as `Error` values handled by `logical_coerce`).
//
// When every argument is skipped (all blanks / "") the result is `#VALUE!`
// rather than the neutral default, matching `XOR("")` -> #VALUE!.
//
// AND / OR are NOT registered here: they ride the lazy dispatch table
// (`eval_and_lazy` / `eval_or_lazy` in `special_forms_lazy.cpp`), which the
// call dispatcher consults before the eager registry, so any eager AND / OR
// body would be unreachable. XOR has no lazy entry and remains eager.
Value Xor_(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  bool result = false;
  bool any_value = false;
  for (std::uint32_t i = 0; i < arity; ++i) {
    bool v = false;
    ErrorCode err = ErrorCode::Value;
    const LogicalCoerce lc = logical_coerce(args[i], &v, &err);
    if (lc == LogicalCoerce::Error) {
      return Value::error(err);
    }
    if (lc == LogicalCoerce::Skip) {
      continue;
    }
    any_value = true;
    result ^= v;
  }
  if (!any_value) {
    return Value::error(ErrorCode::Value);
  }
  return Value::boolean(result);
}

}  // namespace

void register_logical_builtins(FunctionRegistry& registry) {
  // AND / OR are intentionally absent: they are served by the lazy dispatch
  // table (see `special_forms_lazy.cpp`), which unifies the range / array /
  // spill argument gate that the eager path could not express. XOR stays
  // eager and range-aware: Text and Blank cells inside a range are silently
  // skipped rather than surfacing #VALUE!.
  static constexpr builtins_detail::BuiltinRegistration functions[] = {
      {"TRUE", 0u, 0u, &True_},
      {"FALSE", 0u, 0u, &False_},
      {"NOT", 1u, 1u, &Not},
      {"XOR", 1u, kVariadic, &Xor_, true, true, false, true},
  };
  builtins_detail::register_builtin_functions(registry, functions, sizeof(functions) / sizeof(functions[0]));
}

}  // namespace eval
}  // namespace formulon
