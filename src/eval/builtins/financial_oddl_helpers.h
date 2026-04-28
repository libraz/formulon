// Copyright 2026 libraz. Licensed under the MIT License.
//
// Internal header -- do not include outside `src/eval/builtins/financial*`.
//
// Shared closed-form schedule + clean-price helpers used by both
// ODDLPRICE (which returns the price directly) and ODDLYIELD (which
// inverts ODDLPRICE in closed form to recover the yield-to-maturity from
// a target market price). The "odd last" family handles a final coupon
// period that does not align with the regular `12/freq`-month grid: the
// bond pays its periodic coupons up to `last_interest` and then a single
// irregular coupon + redemption at `maturity`. The interval
// (last_interest, maturity] is composed of `NC` quasi-coupon
// sub-periods of identical normal length `E` (basis-adjusted days,
// e.g. 360/freq for basis 0/2/4); the irregular bit is just that
// `maturity` itself is not a regular coupon date relative to a backward
// walk from any subsequent date.
//
// `compute_odd_last_schedule` walks the quasi-coupon dates forward from
// `last_interest` by `12/freq` months per step (preserving the
// last_interest day-of-month, clamped to month-end as needed) and
// reports:
//
//   * `dc_total` -- basis-adjusted days from last_interest to maturity
//                   (sum of per-quasi-period day counts; equals NC*E for
//                   the typical non-basis-1 case).
//   * `a_total`  -- basis-adjusted days from last_interest to settlement
//                   (sum across quasi-periods up to the one containing
//                   settlement).
//   * `dsc`      -- basis-adjusted days from settlement to maturity
//                   (= dc_total - a_total in the clean case; computed
//                   independently as the sum from the settlement-bearing
//                   quasi-period through maturity, which agrees with
//                   `dc_total - a_total` to floating-point precision).
//   * `e`        -- normal-period length in basis-adjusted days
//                   (360/freq for basis 0/2/4, 365/freq for basis 3,
//                   actual NCD-PCD gap for basis 1's first quasi-period;
//                   for basis 1 each quasi-period contributes its own
//                   actual length to dc/a/dsc but `e` is reported as the
//                   first quasi-period's length, which is what
//                   Microsoft's documented formula uses for the
//                   simple-interest residual discount factor).
//
// `compute_oddl_clean_price` evaluates ODDLPRICE's closed form:
//
//   cf   = 100 * rate / freq * (DC_total / E)
//   ai   = 100 * rate / freq * (A_total  / E)
//   disc = 1 + DSC * yld / freq / E
//   ODDLPRICE = (redemption + cf) / disc - ai
//
// `compute_oddl_yield` inverts the above for `yld`:
//
//   yld = (freq * E / DSC) * ((redemption + cf - pr - ai) / (pr + ai))
//
// All three helpers return `Expected<...>` and surface any
// validation/numerical failure as `ErrorCode::Num`.
//
// The shared header pattern mirrors `financial_clean_price.h` (which
// links PRICE / YIELD); see that file for the precedent.

#ifndef FORMULON_EVAL_BUILTINS_FINANCIAL_ODDL_HELPERS_H_
#define FORMULON_EVAL_BUILTINS_FINANCIAL_ODDL_HELPERS_H_

#include <cstdint>

#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {

/// Schedule context for an ODDLPRICE / ODDLYIELD evaluation.
struct OddLastSchedule {
  double dc_total;  ///< Basis-adjusted days from last_interest to maturity.
  double a_total;   ///< Basis-adjusted days from last_interest to settlement.
  double dsc;       ///< Basis-adjusted days from settlement to maturity.
  double e;         ///< Normal coupon-period length in basis-adjusted days.
};

/// Builds the OddLastSchedule for `(last_interest, settlement, maturity,
/// frequency, basis)`. Callers must have already validated
/// `last_interest < settlement < maturity`, `frequency` in {1, 2, 4}, and
/// `basis` in {0..4}.
///
/// Returns `ErrorCode::Num` on any internal failure (date decomposition,
/// non-finite intermediate value, schedule walk that overshoots maturity
/// without ever covering settlement). Given valid pre-validated inputs
/// this path is not reachable in practice; the failure modes are
/// pathological-input defenses.
Expected<OddLastSchedule, ErrorCode> compute_odd_last_schedule(double settlement, double maturity, double last_interest,
                                                               int frequency, int basis) noexcept;

/// Computes the ODDLPRICE clean price per 100 face. Performs the same
/// argument validation order as PRICE (date ordering -> frequency domain
/// -> basis domain -> rate / yld sign -> redemption sign), additionally
/// rejecting `last_interest >= settlement`. Surfaces any failure as
/// `ErrorCode::Num`.
///
/// `args` layout is ODDLPRICE's positional contract:
///
///   args[0] = settlement     (Excel serial)
///   args[1] = maturity       (Excel serial)
///   args[2] = last_interest  (Excel serial)
///   args[3] = rate           (annual coupon rate, decimal)
///   args[4] = yld            (annual yield to maturity, decimal)
///   args[5] = redemption     (per 100 face, > 0)
///   args[6] = frequency      (1, 2, or 4)
///   args[7] = basis          (0..4, optional; only consulted when arity == 8)
Expected<double, ErrorCode> compute_oddl_clean_price(const Value* args, std::uint32_t arity);

/// Computes the ODDLYIELD yield-to-maturity (decimal). Same arg layout as
/// ODDLPRICE except slot 4 holds `pr` (clean market price, > 0) instead
/// of `yld`. Returns `ErrorCode::Num` on any validation / numerical
/// failure (including the closed-form denominator `pr + ai == 0`).
Expected<double, ErrorCode> compute_oddl_yield(const Value* args, std::uint32_t arity);

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_FINANCIAL_ODDL_HELPERS_H_
