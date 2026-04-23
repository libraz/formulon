// Copyright 2026 libraz. Licensed under the MIT License.
//
// Registers Excel's transcendental math built-ins (EXP/LN/LOG/LOG10/PI/
// RADIANS/DEGREES/SIN/COS/TAN/ASIN/ACOS/ATAN/ATAN2) into a FunctionRegistry.
// Kept in its own translation unit so the transcendental family can evolve
// independently of the rest of the builtin catalog.

#ifndef FORMULON_EVAL_BUILTINS_MATH_TRIG_H_
#define FORMULON_EVAL_BUILTINS_MATH_TRIG_H_

namespace formulon {
namespace eval {

class FunctionRegistry;

/// Registers the exponential, logarithmic, and trigonometric built-in
/// functions (EXP, LN, LOG, LOG10, PI, RADIANS, DEGREES, SIN, COS, TAN,
/// ASIN, ACOS, ATAN, ATAN2) into `registry`. Intended to be invoked from
/// `register_builtins`.
void register_math_trig_builtins(FunctionRegistry& registry);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_MATH_TRIG_H_
