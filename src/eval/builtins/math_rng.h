// Copyright 2026 libraz. Licensed under the MIT License.
//
// Registers Excel's random-number built-ins (RAND, RANDBETWEEN) into a
// FunctionRegistry. Kept in its own translation unit so the RNG family can
// evolve independently of the rest of the math catalog in `math.cpp`.
//
// Both functions are volatile in Excel's sense: they re-evaluate on every
// call. Formulon enforces this implicitly because `FunctionDef` has no
// cached-result slot -- every dispatch re-runs `impl`.

#ifndef FORMULON_EVAL_BUILTINS_MATH_RNG_H_
#define FORMULON_EVAL_BUILTINS_MATH_RNG_H_

namespace formulon {
namespace eval {

class FunctionRegistry;

/// Registers the random-number scalar built-ins (RAND, RANDBETWEEN) into
/// `registry`. Intended to be invoked from `register_builtins`.
void register_math_rng_builtins(FunctionRegistry& registry);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_MATH_RNG_H_
