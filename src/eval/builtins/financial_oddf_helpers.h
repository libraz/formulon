//
// Internal header -- do not include outside `src/eval/builtins/financial*`.
//
// Shared schedule and price kernel of ODDFPRICE and ODDFYIELD. The "odd
// first" family prices a bond whose first coupon period is irregular: issued
// at `issue`, first coupon at `first_coupon`, regular coupons to `maturity`.
//
// The coupon schedule is anchored at maturity and keeps its end-of-month code
// (1: last day of the month, 29 / 30: that fixed day, otherwise the plain day
// clamped to short months); `first_coupon` must lie on it, else #NUM!.
//
//   E   = length of the quasi period holding settlement, walking back from
//         first_coupon (360/freq, 365/freq for basis 3, actual for basis 1)
//   N   = coupons between settlement and maturity
//   DFC = days(issue, first_coupon); DFC == E is #NUM!
//
// Short first period (DFC < E), with DSC = days(settlement, first_coupon),
// A = days(issue, settlement), t1 = DSC / E, cf = (100 * rate) / freq:
//
//   price = ((red / q^((N-1) + t1) + ((cf * DFC) / E) / q^t1)
//            + sum_{k=2..N} cf / q^((k-1) + t1)) - ((A / E) * (rate / freq)) * 100
//
// Long first period (DFC > E): the quasi periods before first_coupon are
// stepped back with plain month steps; each contributes DC / NL (every one but
// the earliest counts in full) to `sum_dcnl` and A_i / NL to `sum_a`. With
// t = DSC / E, Nq whole quasi periods between settlement's quasi date and
// first_coupon and N2 coupons from first_coupon to maturity:
//
//   price = ((red / q^((t + Nq) + N2) + (cf * sum_dcnl) / q^(Nq + t))
//            + sum_{k=1..N2} cf / q^((k + Nq) + t)) - cf * sum_a
//
// ODDFYIELD runs YIELD's Newton iteration (same start, derivative and
// stopping rule) on this price.

#ifndef FORMULON_EVAL_BUILTINS_FINANCIAL_ODDF_HELPERS_H_
#define FORMULON_EVAL_BUILTINS_FINANCIAL_ODDF_HELPERS_H_

#include <cstdint>

#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {

/// A validated ODDFPRICE / ODDFYIELD call with every yield-independent schedule fact resolved.
struct OddFirstBond {
  double settlement;
  double maturity;
  double issue;
  double first_coupon;
  double rate;
  double slot5;  ///< args[5]: yld for ODDFPRICE, pr for ODDFYIELD.
  double redemption;
  double freq_d;
  int basis;
  double e;          ///< Quasi-period length E.
  std::int32_t n;    ///< Coupons between settlement and maturity.
  double dsc_first;  ///< days(settlement, first_coupon).
  bool long_first;   ///< DFC > E.
  // Short first period.
  double dfc;
  double accrued_days;  ///< A.
  // Long first period.
  double sum_dcnl;
  double sum_a;
  double dsc;  ///< DSC of the long form.
  double nq;
  std::int32_t n2;
};

/// Reads and validates the nine-slot ODDF argument list (date ordering, frequency, basis, signs,
/// first_coupon on the maturity schedule, DFC != E) and resolves the schedule. `slot5_must_be_positive`
/// is the one difference between the two spellings: ODDFPRICE accepts a zero `yld`, ODDFYIELD rejects
/// a non-positive `pr`. Any failure is `ErrorCode::Num`.
Expected<OddFirstBond, ErrorCode> read_odd_first_bond(const Value* args, std::uint32_t arity,
                                                      bool slot5_must_be_positive, bool date1904);

/// ODDFPRICE's closed form at yield `yld`; `bond.slot5` is not read. `ErrorCode::Num` on a non-finite result.
Expected<double, ErrorCode> oddf_price_at(const OddFirstBond& bond, double yld);

/// Computes the ODDFPRICE clean price per 100 face.
///
///   args[0] = settlement, args[1] = maturity, args[2] = issue, args[3] = first_coupon,
///   args[4] = rate, args[5] = yld, args[6] = redemption, args[7] = frequency (1, 2, 4),
///   args[8] = basis (0..4, optional)
Expected<double, ErrorCode> compute_oddf_clean_price(const Value* args, std::uint32_t arity, bool date1904);

/// Computes the ODDFYIELD yield-to-maturity (decimal). Same layout as ODDFPRICE with `pr` in slot 5.
Expected<double, ErrorCode> compute_oddf_yield(const Value* args, std::uint32_t arity, bool date1904);

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_FINANCIAL_ODDF_HELPERS_H_
