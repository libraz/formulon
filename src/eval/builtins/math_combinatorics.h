// Copyright 2026 libraz. Licensed under the MIT License.
//
// Registers Excel's combinatorial, numeral-system, precise-rounding, and
// miscellaneous scalar math built-ins into a `FunctionRegistry`. Kept in
// its own translation unit so the combinatorics family can evolve
// independently of the rest of the math catalog in `math.cpp`.
//
// Registered names: ARABIC, ROMAN, BASE, DECIMAL, CEILING.PRECISE,
// FLOOR.PRECISE, ISO.CEILING, COMBIN, COMBINA, FACT, FACTDOUBLE, GCD, LCM,
// MULTINOMIAL, SQRTPI.

#ifndef FORMULON_EVAL_BUILTINS_MATH_COMBINATORICS_H_
#define FORMULON_EVAL_BUILTINS_MATH_COMBINATORICS_H_

namespace formulon {
namespace eval {

class FunctionRegistry;

/// Registers the combinatorial / numeral-system / precise-rounding scalar
/// math built-ins (ARABIC, ROMAN, BASE, DECIMAL, CEILING.PRECISE,
/// FLOOR.PRECISE, ISO.CEILING, COMBIN, COMBINA, FACT, FACTDOUBLE, GCD, LCM,
/// MULTINOMIAL, SQRTPI) into `registry`. Intended to be invoked from
/// `register_builtins`.
void register_math_combinatorics_builtins(FunctionRegistry& registry);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_MATH_COMBINATORICS_H_
