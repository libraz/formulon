// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of Formulon's logical built-in functions: TRUE, FALSE, NOT,
// AND, OR, and XOR. Each impl follows the same recipe as the rest of the
// builtin catalog: coerce arguments via `eval/coerce.h`, propagate the
// left-most coercion error, and return a `Value`.

#include "eval/builtins/logical.h"

#include <cstdint>

#include "eval/coerce.h"
#include "eval/function_registry.h"
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

// AND(value, ...) / OR(value, ...) ---------------------------------------
// Both functions evaluate every argument (Excel does not logically
// short-circuit AND / OR; the only short-circuit is the dispatcher's
// left-most-error rule, which fires before this body runs). Each argument
// is coerced via `coerce_to_bool`; the first coercion failure surfaces as
// #VALUE! (or #NUM! for non-finite numeric inputs). AND returns true iff
// every argument coerces to true; OR returns true iff any does.
Value And_(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  bool result = true;
  for (std::uint32_t i = 0; i < arity; ++i) {
    auto coerced = coerce_to_bool(args[i]);
    if (!coerced) {
      return Value::error(coerced.error());
    }
    if (!coerced.value()) {
      result = false;
    }
  }
  return Value::boolean(result);
}

Value Or_(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  bool result = false;
  for (std::uint32_t i = 0; i < arity; ++i) {
    auto coerced = coerce_to_bool(args[i]);
    if (!coerced) {
      return Value::error(coerced.error());
    }
    if (coerced.value()) {
      result = true;
    }
  }
  return Value::boolean(result);
}

/// XOR(value, ...) -- exclusive-or across any number of boolean-coercible
/// arguments. Returns TRUE iff an odd number of the coerced arguments are
/// TRUE. Shares AND / OR's range-aware surface: range-sourced Text / Blank
/// cells are silently skipped via `range_filter_bool_coercible`; direct
/// scalar arguments that fail `coerce_to_bool` surface `#VALUE!`.
Value Xor_(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  bool result = false;
  for (std::uint32_t i = 0; i < arity; ++i) {
    auto coerced = coerce_to_bool(args[i]);
    if (!coerced) {
      return Value::error(coerced.error());
    }
    result ^= coerced.value();
  }
  return Value::boolean(result);
}

}  // namespace

void register_logical_builtins(FunctionRegistry& registry) {
  registry.register_function(FunctionDef{"TRUE", 0u, 0u, &True_});
  registry.register_function(FunctionDef{"FALSE", 0u, 0u, &False_});
  registry.register_function(FunctionDef{"NOT", 1u, 1u, &Not});
  // AND / OR are range-aware so `=AND(A1:A3)` expands the rectangle. The
  // `range_filter_bool_coercible` flag silently drops Text / Blank cells
  // inside a range (Excel skips them rather than surfacing #VALUE!), while
  // direct scalar arguments still flow through `coerce_to_bool` and surface
  // #VALUE! for non-coercible text literals.
  {
    FunctionDef def{"AND", 1u, kVariadic, &And_};
    def.accepts_ranges = true;
    def.range_filter_bool_coercible = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"OR", 1u, kVariadic, &Or_};
    def.accepts_ranges = true;
    def.range_filter_bool_coercible = true;
    registry.register_function(def);
  }
  {
    // XOR shares AND / OR's range-aware surface: a trailing range argument
    // such as `=XOR(A1:A3)` expands in the dispatcher; Text and Blank cells
    // inside the range are silently skipped rather than surfacing #VALUE!.
    FunctionDef def{"XOR", 1u, kVariadic, &Xor_};
    def.accepts_ranges = true;
    def.range_filter_bool_coercible = true;
    registry.register_function(def);
  }
}

}  // namespace eval
}  // namespace formulon
