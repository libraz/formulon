// Copyright 2026 libraz. Licensed under the MIT License.
//
// Shared coupon-schedule engine used by the COUP* family (COUPPCD /
// COUPNCD / COUPNUM / COUPDAYBS / COUPDAYSNC / COUPDAYS) and — when the
// bond-pricing functions land — by PRICE / YIELD / DURATION /
// MDURATION / ACCRINT. Kept in `eval/` (not under `builtins/`) because
// several translation units will need it.
//
// The engine walks the coupon dates backwards from maturity by
// `12/frequency` months per step, preserving the maturity's
// day-of-month and clamping to the last day of the target month when
// the month has fewer days (Aug-31 -> May-31, but Feb-29 in a non-leap
// year -> Feb-28). The previous-coupon-date (PCD) is the last such
// candidate `<= settlement`; the next-coupon-date (NCD) is the step
// forward from that.
//
// Day counts are basis-adjusted following Excel's five-basis ruleset
// (see `eval/date_time.h`'s YEARFRAC helpers); the COUPDAYS period
// length uses the fixed `360 / frequency` or `365 / frequency`
// shortcuts for bases 0/2/4 and 3 respectively, and the actual PCD-NCD
// day gap for basis 1.

#ifndef FORMULON_EVAL_COUPON_SCHEDULE_H_
#define FORMULON_EVAL_COUPON_SCHEDULE_H_

#include <cstdint>

namespace formulon {
namespace eval {

/// Coupon-schedule context for a single (settlement, maturity,
/// frequency, basis) tuple. All day counts are Excel-compatible
/// doubles; `pcd` / `ncd` are Excel serial numbers (integer-valued).
struct CouponDates {
  double pcd;                      ///< Previous coupon date on or before settlement.
  double ncd;                      ///< Next coupon date strictly after settlement.
  double days_bs;                  ///< Basis-adjusted days from PCD to settlement.
  double days_nc;                  ///< Basis-adjusted days from settlement to NCD.
  double period_days;              ///< Basis-adjusted total days in the coupon period.
  std::int32_t coupons_remaining;  ///< Coupons strictly after settlement and <= maturity.
};

/// Computes the full coupon-schedule context.
///
/// Callers must have already validated:
///   - `settlement < maturity`
///   - `frequency` in {1, 2, 4}
///   - `basis` in {0..4}
///
/// Returns `false` on an unrecoverable internal error (e.g. date
/// decomposition failure). Given valid pre-validated inputs this path
/// is not reachable in practice; callers translate a `false` return
/// into `#NUM!` defensively.
bool compute_coupon_dates(double settlement, double maturity, int frequency, int basis, CouponDates* out) noexcept;

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_COUPON_SCHEDULE_H_
