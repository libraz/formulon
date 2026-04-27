// Copyright 2026 libraz. Licensed under the MIT License.
//
// Internal header -- do not include outside `src/eval/builtins/financial*`.
//
// Shared closed-form clean-price computation used by both PRICE (which
// returns the price directly) and YIELD (which inverts this function via
// Newton-Raphson to recover the yield-to-maturity from a target market
// price). Keeping the helper in a header lets the YIELD translation unit
// call into the same numerical kernel without re-implementing the
// two-branch (n==1 simple-interest vs n>1 v^(t1+i) discounting) logic.
//
// The function takes the same `(args, arity)` shape as the Value-returning
// builtins -- it parses settlement / maturity / rate / yld / redemption /
// frequency / [basis] from positions 0..6 of `args` and validates them in
// the documented order. Any validation failure surfaces as `#NUM!` via
// the returned `Expected`. See `financial_price.cpp` for the file-level
// derivation comment that documents the closed form.

#ifndef FORMULON_EVAL_BUILTINS_FINANCIAL_CLEAN_PRICE_H_
#define FORMULON_EVAL_BUILTINS_FINANCIAL_CLEAN_PRICE_H_

#include <cstdint>

#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {

/// Computes the clean price per 100 face for a regular-period bond.
///
/// `args` must point to at least 6 (or 7 when `arity == 7`) Values laid
/// out as PRICE / YIELD's positional contract:
///
///   args[0] = settlement (Excel serial)
///   args[1] = maturity   (Excel serial)
///   args[2] = rate       (annual coupon rate, decimal)
///   args[3] = yld        (annual yield to maturity, decimal)
///   args[4] = redemption (per 100 face)
///   args[5] = frequency  (1, 2, or 4)
///   args[6] = basis      (0..4, optional; only consulted when arity==7)
///
/// Returns the clean price on success or `ErrorCode::Num` on any
/// validation / numerical failure (date ordering, frequency / basis
/// domain, negative rate / yld, non-positive redemption, coupon-schedule
/// failure, non-finite intermediate or final value).
Expected<double, ErrorCode> compute_clean_price(const Value* args, std::uint32_t arity);

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_FINANCIAL_CLEAN_PRICE_H_
