// Copyright 2026 libraz. Licensed under the MIT License.
//
// Registers Excel's engineering built-ins into a FunctionRegistry. The
// current scope covers the "simple integer" subset: base conversion
// (BIN/OCT/HEX/DEC mutual converters), bit manipulation (BITAND / BITOR /
// BITXOR / BITLSHIFT / BITRSHIFT), and the two comparators DELTA and
// GESTEP. Heavier engineering families (BESSEL*, IM* complex arithmetic,
// CONVERT, ERF*) live in their own future translation units.
//
// All functions in this file are eager scalar (dispatched via the default
// `FunctionDef` path); no tree-walker changes are needed.

#ifndef FORMULON_EVAL_BUILTINS_ENGINEERING_H_
#define FORMULON_EVAL_BUILTINS_ENGINEERING_H_

namespace formulon {
namespace eval {

class FunctionRegistry;

/// Registers the simple integer-only engineering built-ins into `registry`.
///
/// Included functions:
///   * Base conversion (12): BIN2DEC, BIN2OCT, BIN2HEX, OCT2DEC, OCT2BIN,
///     OCT2HEX, HEX2DEC, HEX2BIN, HEX2OCT, DEC2BIN, DEC2OCT, DEC2HEX.
///   * Bit operations (5): BITAND, BITOR, BITXOR, BITLSHIFT, BITRSHIFT.
///   * Comparators (2): DELTA, GESTEP.
///
/// Intended to be invoked from `register_builtins`.
void register_engineering_builtins(FunctionRegistry& registry);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_ENGINEERING_H_
