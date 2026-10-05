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

#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {

class FunctionRegistry;

namespace builtins_detail {

/// Coerces `v` to a finite number truncated toward zero and within
/// [lo, hi]. A direct Bool is `#VALUE!` when `reject_bool`; non-finite or
/// out-of-range values are `#NUM!`. Shared by the engineering integer
/// arguments and the BESSEL order.
Expected<double, ErrorCode> coerce_truncated_in_range(const Value& v, double lo, double hi, bool reject_bool);

}  // namespace builtins_detail

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
