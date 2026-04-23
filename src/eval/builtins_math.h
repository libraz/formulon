// Copyright 2026 libraz. Licensed under the MIT License.
//
// Registers Excel's arithmetic and rounding built-ins (ABS/SIGN/INT/TRUNC/
// SQRT/MOD/POWER/ROUND/ROUNDDOWN/ROUNDUP) into a FunctionRegistry. Kept in
// its own translation unit so the math/rounding family can evolve
// independently of the rest of the builtin catalog.

#ifndef FORMULON_EVAL_BUILTINS_MATH_H_
#define FORMULON_EVAL_BUILTINS_MATH_H_

namespace formulon {
namespace eval {

class FunctionRegistry;

/// Registers the math and rounding built-in functions (ABS, SIGN, INT, TRUNC,
/// SQRT, MOD, POWER, ROUND, ROUNDDOWN, ROUNDUP) into `registry`. Intended to
/// be invoked from `register_builtins`.
void register_math_builtins(FunctionRegistry& registry);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_MATH_H_
