//
// Implementation of the irregular-first-period bond yield-to-maturity
// built-in:
//
//   * ODDFYIELD(settlement, maturity, issue, first_coupon, rate, pr,
//               redemption, frequency, [basis=0])
//
// Returns the annual yield-to-maturity (decimal) implied by the given
// clean market price `pr`. ODDFYIELD numerically inverts the
// ODDFPRICE kernel: feeding ODDFYIELD's output back into ODDFPRICE
// (with the same arguments) recovers `pr` to ~1e-9 across all five
// bases.
//
// The numerical strategy mirrors YIELD's: Newton-Raphson with a
// central-difference derivative on
//
//   f(yld) = compute_oddf_clean_price(yld) - pr
//
// with
//
//   f'(yld) ~= (f(yld + h) - f(yld - h)) / (2h)
//   h        = 1e-7 * max(1, |yld|)
//
// Convergence: |f(yld)| < 1e-12 * (|pr| + 1) OR |delta| < 1e-15
// within `kMaxIter` (= 100) iterations.
//
// Initial guess: Microsoft's documented "approximate yield" formula:
//
//   yld_0 = ((100*rate + (redemption - pr)/N_total)
//            / ((redemption + pr)/2)) * frequency
//
// where `N_total` is the approximate number of coupon periods from
// settlement to maturity. We take `N_total = nc + n_regular - dsc/E`,
// which is the floating-point period count from settlement through
// the irregular span and on to maturity (consistent with how the
// price kernel's discount exponents add up). This guess is within a
// few percent of the true yield across the entire valid bond domain;
// non-convergence (degenerate derivative or 100-iter timeout)
// surfaces as `#NUM!`.
//
// Why no closed-form inverse: ODDFPRICE is a polynomial in
// `v = 1 / (1 + yld/freq)` of degree `nc + n_regular`, with
// generally-unequal coefficients (the dc[i]/nl[i] ratios differ
// across quasi-periods in the long-first-period case, and even when
// they're equal there are still `n_regular + 1` distinct discount
// powers). No algebraic root extraction exists for arbitrary degree.
// Even ODDLPRICE, which DOES have a closed form, only gets one
// because its residual period uses simple-interest discounting on a
// single cash flow -- a structural simplification ODDFPRICE cannot
// share.

#include "eval/builtins/financial_oddfyield.h"

#include <algorithm>
#include <cmath>
#include <cstdint>

#include "eval/builtins/financial_helpers.h"
#include "eval/builtins/financial_oddf_helpers.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {
namespace {

// Evaluates `compute_oddf_clean_price` at the candidate yld by reusing
// ODDFYIELD's own argument vector with slot 5 swapped for the iterate.
//
// ODDFPRICE's arg layout is `(settlement, maturity, issue,
// first_coupon, rate, yld, redemption, frequency, [basis])`;
// ODDFYIELD's layout is identical except slot 5 holds `pr`. Only that
// one slot differs, so we copy the (already-validated) Values
// verbatim and substitute the yld slot.
//
// `Value`'s default constructor is private, so we initialise the
// buffer from existing valid Values (slot 0) before assigning into
// each slot. Buffer sized 9 to cover the optional basis without
// branching; `arity` is forwarded unchanged so
// `read_optional_number` inside `compute_oddf_clean_price` selects
// the correct value. Mirrors the `price_at` helper in
// `financial_yield.cpp`.
Expected<double, ErrorCode> price_at(const Value* yield_args, std::uint32_t arity, double yld_iterate) {
  Value buf[9] = {yield_args[0], yield_args[0], yield_args[0], yield_args[0], yield_args[0],
                  yield_args[0], yield_args[0], yield_args[0], yield_args[0]};
  buf[1] = yield_args[1];
  buf[2] = yield_args[2];
  buf[3] = yield_args[3];
  buf[4] = yield_args[4];
  buf[5] = Value::number(yld_iterate);
  buf[6] = yield_args[6];
  buf[7] = yield_args[7];
  if (arity == 9) {
    buf[8] = yield_args[8];
  }
  return compute_oddf_clean_price(buf, arity);
}

}  // namespace

Expected<double, ErrorCode> compute_oddf_yield(const Value* args, std::uint32_t arity) {
  // The schedule is built once here for the initial-guess heuristic. The
  // Newton iterate calls `price_at` (which rebuilds the schedule each time)
  // -- the schedule's structure is independent of yld so recomputing it per
  // iteration is cheap relative to the sum-of-powers in the price kernel.
  OddFirstSchedule sched{};
  auto in = read_odd_first_inputs(args, arity, /*slot5_must_be_positive=*/true, sched);
  if (!in) {
    return in.error();
  }

  const double freq_d = in.value().freq_d;
  const double red = in.value().redemption;
  const double pr_v = in.value().slot5;

  // Initial guess: Microsoft's documented approximate-yield formula.
  // `n_total` is the total period count from settlement to maturity
  // (irregular + regular), expressed as a floating-point fraction so
  // the heuristic stays well-conditioned on long-first-period cases
  // where settlement falls deep inside the irregular span.
  const double n_total = static_cast<double>(sched.nc) + static_cast<double>(sched.n_regular) - sched.dsc / sched.e;
  const double avg_capital = (red + pr_v) / 2.0;
  if (avg_capital == 0.0 || n_total <= 0.0) {
    return ErrorCode::Num;
  }
  double yld = ((100.0 * in.value().rate + (red - pr_v) / n_total) / avg_capital) * freq_d;
  if (std::isnan(yld) || std::isinf(yld)) {
    return ErrorCode::Num;
  }
  // ODDFPRICE has the same yld >= 0 domain as PRICE. Premium-bond
  // heuristics can begin below zero, so clamp to the boundary and use the
  // one-sided derivative below rather than evaluating an invalid iterate.
  yld = std::max(0.0, yld);

  return solve_yield_by_newton(price_at, args, arity, pr_v, yld);
}

// --- ODDFYIELD(settlement, maturity, issue, first_coupon, rate, pr,
//              redemption, frequency, [basis=0]) ----------------------------
//
// Annual yield-to-maturity (decimal) for a security whose first
// coupon period is irregular. Numerically inverts ODDFPRICE via
// Newton-Raphson.
Value OddfYield(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto y = compute_oddf_yield(args, arity);
  if (!y) {
    return Value::error(y.error());
  }
  return finalize(y.value());
}

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon
