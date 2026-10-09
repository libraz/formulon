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
// (last_interest, maturity] is split into quasi-coupon periods walked
// forward from `last_interest` by `12/freq` months (day of month kept,
// clamped to shorter months); the last one is cut short by `maturity`.
//
// `compute_odd_last_schedule` sums, over those periods, each period's
// days divided by its normal length NL_i (360/freq for bases 0 / 4, the
// period's actual length for bases 1 / 2 / 3):
//
//   * `dc_units`  -- sum of DC_i / NL_i, days from last_interest to maturity.
//   * `a_units`   -- sum of A_i / NL_i, days from last_interest to settlement.
//   * `dsc_units` -- sum of DSC_i / NL_i, days from settlement to maturity.
//
// `compute_oddl_clean_price` evaluates ODDLPRICE's closed form:
//
//   cf   = 100 * rate / freq * dc_units
//   ai   = 100 * rate / freq * a_units
//   disc = 1 + dsc_units * yld / freq
//   ODDLPRICE = (redemption + cf) / disc - ai
//
// `compute_oddl_yield` inverts the above for `yld`:
//
//   yld = (freq / dsc_units) * ((redemption + cf - pr - ai) / (pr + ai))
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
  double dc_units;   ///< Coupon periods from last_interest to maturity.
  double a_units;    ///< Coupon periods from last_interest to settlement.
  double dsc_units;  ///< Coupon periods from settlement to maturity.
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
                                                               int frequency, int basis, bool date1904) noexcept;

/// Everything ODDLPRICE and ODDLYIELD read and derive before their closed
/// forms diverge. Slot 4 is the one argument whose meaning differs between the
/// two — `yld` for the price spelling, `pr` for the yield spelling — so it is
/// carried unnamed.
struct OddLastInputs {
  double slot4;       ///< args[4]: yld for ODDLPRICE, pr for ODDLYIELD.
  double redemption;  ///< args[5], per 100 face.
  double cf;          ///< Irregular final coupon per 100 face.
  double ai;          ///< Accrued interest per 100 face.
  double dsc_units;   ///< Coupon periods from settlement to maturity.
  double freq_d;      ///< Coupon frequency as a double.
};

/// Reads and validates the eight-slot ODDL argument list and walks the coupon
/// schedule, leaving only the closed form itself to the caller.
///
/// Validation order is PRICE's: date ordering -> frequency -> basis -> rate
/// sign -> slot-4 sign -> redemption sign, plus the odd-last constraint
/// `last_interest < settlement`. `slot4_must_be_positive` is the one
/// difference between the two spellings: ODDLPRICE accepts a zero `yld`, while
/// ODDLYIELD rejects a non-positive `pr` because a zero price implies an
/// infinite yield. Any failure surfaces as `ErrorCode::Num`.
Expected<OddLastInputs, ErrorCode> read_odd_last_inputs(const Value* args, std::uint32_t arity,
                                                        bool slot4_must_be_positive, bool date1904);

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
Expected<double, ErrorCode> compute_oddl_clean_price(const Value* args, std::uint32_t arity, bool date1904);

/// Computes the ODDLYIELD yield-to-maturity (decimal). Same arg layout as
/// ODDLPRICE except slot 4 holds `pr` (clean market price, > 0) instead
/// of `yld`. Returns `ErrorCode::Num` on any validation / numerical
/// failure (including the closed-form denominator `pr + ai == 0`).
Expected<double, ErrorCode> compute_oddl_yield(const Value* args, std::uint32_t arity, bool date1904);

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_FINANCIAL_ODDL_HELPERS_H_
