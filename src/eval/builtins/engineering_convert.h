// Copyright 2026 libraz. Licensed under the MIT License.
//
// Registers Excel's CONVERT built-in. CONVERT is large enough -- it carries
// ~80 unit entries plus SI / binary prefix rules -- that it lives in its
// own translation unit to keep `engineering.cpp` focused on the simple
// integer-only engineering subset (base conversion / bit ops / DELTA /
// GESTEP).

#ifndef FORMULON_EVAL_BUILTINS_ENGINEERING_CONVERT_H_
#define FORMULON_EVAL_BUILTINS_ENGINEERING_CONVERT_H_

namespace formulon {
namespace eval {

class FunctionRegistry;

/// Registers the CONVERT built-in into `registry`. Intended to be invoked
/// from `register_builtins`.
void register_engineering_convert_builtins(FunctionRegistry& registry);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_ENGINEERING_CONVERT_H_
