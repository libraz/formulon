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

// AND(value, ...) / OR(value, ...) / XOR(value, ...) ---------------------
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
// rather than the neutral default, matching `AND("")` -> #VALUE!.
Value And_(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  bool result = true;
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
    if (!v) {
      result = false;
    }
  }
  if (!any_value) {
    return Value::error(ErrorCode::Value);
  }
  return Value::boolean(result);
}

Value Or_(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
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
    if (v) {
      result = true;
    }
  }
  if (!any_value) {
    return Value::error(ErrorCode::Value);
  }
  return Value::boolean(result);
}

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
  // AND / OR are range-aware so `=AND(A1:A3)` expands the rectangle. The
  // `range_filter_bool_coercible` flag silently drops Text / Blank cells
  // inside a range (Excel skips them rather than surfacing #VALUE!), while
  // direct scalar arguments still flow through `coerce_to_bool` and surface
  // #VALUE! for non-coercible text literals.
  static constexpr builtins_detail::BuiltinRegistration functions[] = {
      {"TRUE", 0u, 0u, &True_},
      {"FALSE", 0u, 0u, &False_},
      {"NOT", 1u, 1u, &Not},
      {"AND", 1u, kVariadic, &And_, true, true, false, true},
      {"OR", 1u, kVariadic, &Or_, true, true, false, true},
      // XOR shares AND / OR's range-aware surface: Text and Blank cells
      // inside a range are silently skipped rather than surfacing #VALUE!.
      {"XOR", 1u, kVariadic, &Xor_, true, true, false, true},
  };
  builtins_detail::register_builtin_functions(registry, functions, sizeof(functions) / sizeof(functions[0]));
}

}  // namespace eval
}  // namespace formulon
