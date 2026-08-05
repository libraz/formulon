//
// Internal header -- do not include outside `src/eval/builtins/financial*`.
//
// Shared schedule walker + clean-price kernel used by both ODDFPRICE
// (clean price per 100 face) and ODDFYIELD (yield-to-maturity inferred
// by Newton-Raphson on the price kernel). The "odd first" family
// handles a bond whose initial coupon period does not align with the
// regular `12/freq`-month grid: the issue date precedes the first
// coupon date by an arbitrary span which may be shorter than, equal
// to, or longer than a normal period. After `first_coupon` the bond
// pays coupons on the regular schedule up to `maturity`.
//
// `compute_odd_first_schedule` walks the quasi-coupon dates BACKWARD
// from `first_coupon` by `12/freq` months per step (preserving
// `first_coupon`'s day-of-month and clamping to month-end as needed)
// until the candidate date is at or before `issue`. It records a
// vector of quasi-periods plus aggregate day counts:
//
//   * `nc`               -- number of quasi-periods covering the
//                           irregular first span [issue, first_coupon].
//   * `qp[0..nc-1]`      -- per-quasi-period start/end serials, sorted
//                           oldest -> newest. `qp[0].start <= issue`,
//                           `qp[nc-1].end == first_coupon`.
//   * `nl[i]`            -- basis-adjusted length of quasi-period i.
//   * `dc[i]`            -- days within quasi-period i that contribute
//                           to coupon income (issue-to-end on the
//                           first; full nl[i] on subsequent ones).
//   * `a[i]`             -- days within quasi-period i that contribute
//                           to accrued interest (start-or-issue to
//                           settlement when settlement falls in i; 0
//                           for quasi-periods after the
//                           settlement-bearing one; full nl[i] for
//                           quasi-periods strictly before it). The
//                           first quasi-period substitutes `issue` for
//                           its start.
//   * `dsc`              -- basis-adjusted days from settlement to
//                           `first_coupon` summed over the
//                           settlement-bearing quasi-period plus all
//                           subsequent ones in the irregular span.
//   * `e`                -- normal coupon-period length in
//                           basis-adjusted days (360/freq for basis
//                           0/2/4, 365/freq for basis 3, the actual
//                           length of the most-recent quasi-period
//                           anchored on `first_coupon` for basis 1).
//   * `n_regular`        -- number of regular coupon periods between
//                           `first_coupon` and `maturity` (>= 1, from
//                           `compute_coupon_dates(first_coupon,
//                           maturity, ...)`).
//
// `compute_oddf_clean_price` evaluates Microsoft's documented
// closed form:
//
//   cf  = 100 * rate / freq
//   ai  = cf * sum_i (a[i] / nl[i])
//   v   = 1 / (1 + yld / freq)
//   first_period_pv = cf * sum_i ((dc[i] / nl[i]) * v^(NQ_i + dsc/E))
//                    where NQ_i = nc - i (number of quasi-periods
//                    between i and first_coupon, exclusive on i and
//                    inclusive on the discount steps).
//   reg_coupons_pv  = cf * sum_{j=1..n_regular} v^(dsc/E + j)
//   redemption_pv   = redemption * v^(dsc/E + n_regular)
//   ODDFPRICE       = first_period_pv + reg_coupons_pv + redemption_pv - ai
//
// When NC == 1 (short first period) the first_period_pv collapses to
// the standard short-first-period formula; this implementation walks
// the same code path for every NC.
//
// `compute_oddf_yield` inverts the above for `yld` via Newton-Raphson
// on `price(yld) - pr` with a central-difference derivative. The
// initial guess is Microsoft's documented "approximate yield"
// formula. There is no closed-form inverse: the price kernel is a
// polynomial in `v` of degree `nc + n_regular`, which has no
// algebraic solution in `yld` for the general case.

#ifndef FORMULON_EVAL_BUILTINS_FINANCIAL_ODDF_HELPERS_H_
#define FORMULON_EVAL_BUILTINS_FINANCIAL_ODDF_HELPERS_H_

#include <cstdint>

#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {

/// Maximum number of quasi-periods we record for the irregular first
/// span. A bond with semi-annual coupons would need NC = 200 to span a
/// 100-year-old issue; that vastly exceeds Excel's documented bond
/// horizon. The schedule walker surfaces `ErrorCode::Num` if the cap is
/// exceeded, matching Excel's behaviour for pathological inputs.
constexpr int kMaxOddFirstQuasiPeriods = 200;

/// One quasi-coupon period in the irregular first span.
struct OddFirstQuasiPeriod {
  double start;  ///< Earliest serial of the quasi-period (oldest = i=0).
  double end;    ///< Latest serial; equal to `first_coupon` for i=nc-1.
  double nl;     ///< Basis-adjusted period length.
  double dc;     ///< Basis-adjusted days contributing to coupon income.
  double a;      ///< Basis-adjusted days contributing to accrued interest.
};

/// Schedule context for an ODDFPRICE / ODDFYIELD evaluation.
struct OddFirstSchedule {
  int nc;                                            ///< Number of quasi-periods in the irregular first span.
  OddFirstQuasiPeriod qp[kMaxOddFirstQuasiPeriods];  ///< Per-quasi-period detail, oldest first.
  double dsc;                                        ///< Days from settlement to first_coupon (basis-adjusted).
  double e;                                          ///< Normal coupon-period length in basis-adjusted days.
  std::int32_t n_regular;                            ///< Regular coupon periods between first_coupon and maturity.
};

/// Builds the OddFirstSchedule for `(settlement, maturity, issue,
/// first_coupon, frequency, basis)`. Callers must have already
/// validated `issue < settlement`, `settlement < first_coupon`,
/// `first_coupon < maturity`, `frequency` in {1, 2, 4}, and `basis`
/// in {0..4}.
///
/// Returns `ErrorCode::Num` on any internal failure (date
/// decomposition, non-finite intermediate value, schedule walk that
/// exceeds `kMaxOddFirstQuasiPeriods`, or `compute_coupon_dates`
/// failure). Given valid pre-validated inputs the failure modes are
/// pathological-input defenses.
Expected<OddFirstSchedule, ErrorCode> compute_odd_first_schedule(double settlement, double maturity, double issue,
                                                                 double first_coupon, int frequency,
                                                                 int basis) noexcept;

/// Computes the ODDFPRICE clean price per 100 face. Performs the same
/// argument validation order as PRICE / ODDLPRICE (date ordering ->
/// frequency domain -> basis domain -> rate / yld sign -> redemption
/// sign), additionally requiring `issue < settlement < first_coupon <
/// maturity`. Surfaces any failure as `ErrorCode::Num`.
///
/// `args` layout is ODDFPRICE's positional contract:
///
///   args[0] = settlement     (Excel serial)
///   args[1] = maturity       (Excel serial)
///   args[2] = issue          (Excel serial)
///   args[3] = first_coupon   (Excel serial)
///   args[4] = rate           (annual coupon rate, decimal)
///   args[5] = yld            (annual yield to maturity, decimal)
///   args[6] = redemption     (per 100 face, > 0)
///   args[7] = frequency      (1, 2, or 4)
///   args[8] = basis          (0..4, optional; only consulted when arity == 9)
Expected<double, ErrorCode> compute_oddf_clean_price(const Value* args, std::uint32_t arity);

/// Computes the ODDFYIELD yield-to-maturity (decimal). Same arg layout
/// as ODDFPRICE except slot 5 holds `pr` (clean market price, > 0)
/// instead of `yld`. Returns `ErrorCode::Num` on any validation /
/// numerical failure (degenerate derivative, iteration cap reached,
/// non-finite intermediate).
Expected<double, ErrorCode> compute_oddf_yield(const Value* args, std::uint32_t arity);

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_FINANCIAL_ODDF_HELPERS_H_
