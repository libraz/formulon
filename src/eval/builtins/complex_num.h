// Copyright 2026 libraz. Licensed under the MIT License.
//
// Registers Excel's complex-number built-ins into a FunctionRegistry.
//
// Excel represents a complex number as a text value of the form "a+bi" (or
// "a+bj"); the real part is elided when zero, the imaginary coefficient is
// elided when +/-1, and the suffix is always `i` or `j`. See
// backup/plans/02-calc-engine.md for the authoritative formatting table.
//
// The family covers:
//   * Construction / inspection (6): COMPLEX, IMABS, IMAGINARY, IMREAL,
//     IMCONJUGATE, IMARGUMENT.
//   * Arithmetic (5): IMSUM, IMSUB, IMPRODUCT, IMDIV, IMPOWER.
//   * Exponentials (4): IMEXP, IMLN, IMLOG10, IMLOG2, IMSQRT.
//   * Trigonometric (7): IMSIN, IMCOS, IMTAN, IMSEC, IMCSC, IMCOT.
//   * Hyperbolic (4): IMSINH, IMCOSH, IMSECH, IMCSCH.
//
// All entries are eager scalar (dispatched via the default `FunctionDef`
// path); no tree-walker changes are needed.

#ifndef FORMULON_EVAL_BUILTINS_COMPLEX_NUM_H_
#define FORMULON_EVAL_BUILTINS_COMPLEX_NUM_H_

namespace formulon {
namespace eval {

class FunctionRegistry;

/// Registers the complex-number engineering built-ins into `registry`.
///
/// Intended to be invoked from `register_builtins`.
void register_complex_num_builtins(FunctionRegistry& registry);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_COMPLEX_NUM_H_
