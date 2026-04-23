// Copyright 2026 libraz. Licensed under the MIT License.
//
// Registers Excel's "special function" engineering built-ins into a
// FunctionRegistry. This family covers the continuous-math engineering
// functions that do not fit the integer-oriented base-conversion / bit
// manipulation family in `engineering.{h,cpp}` nor the complex-number
// family in `complex_num.{h,cpp}`:
//
//   * Error function family (4): ERF, ERF.PRECISE, ERFC, ERFC.PRECISE.
//   * Bessel function family (4): BESSELJ, BESSELY, BESSELI, BESSELK.
//
// All functions in this file are eager scalar (dispatched via the default
// `FunctionDef` path); no tree-walker changes are needed.

#ifndef FORMULON_EVAL_BUILTINS_ENGINEERING_SPECIAL_H_
#define FORMULON_EVAL_BUILTINS_ENGINEERING_SPECIAL_H_

namespace formulon {
namespace eval {

class FunctionRegistry;

/// Registers the ERF / BESSEL engineering built-ins into `registry`.
///
/// Included functions:
///   * ERF family (4): ERF, ERF.PRECISE, ERFC, ERFC.PRECISE.
///   * BESSEL family (4): BESSELJ, BESSELY, BESSELI, BESSELK.
///
/// Intended to be invoked from `register_builtins`.
void register_engineering_special_builtins(FunctionRegistry& registry);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_ENGINEERING_SPECIAL_H_
