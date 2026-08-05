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
  auto settlement = read_financial_date(args, 0);
  if (!settlement) {
    return settlement.error();
  }
  auto maturity = read_financial_date(args, 1);
  if (!maturity) {
    return maturity.error();
  }
  auto issue = read_financial_date(args, 2);
  if (!issue) {
    return issue.error();
  }
  auto first_coupon = read_financial_date(args, 3);
  if (!first_coupon) {
    return first_coupon.error();
  }
  auto rate = read_required_number(args, 4);
  if (!rate) {
    return rate.error();
  }
  auto pr = read_required_number(args, 5);
  if (!pr) {
    return pr.error();
  }
  auto redemption = read_required_number(args, 6);
  if (!redemption) {
    return redemption.error();
  }
  auto frequency_e = read_coupon_frequency(args, 7);
  if (!frequency_e) {
    return frequency_e.error();
  }
  auto basis_e = read_day_count_basis(args, arity, 8);
  if (!basis_e) {
    return basis_e.error();
  }

  // Validation order mirrors ODDFPRICE: date ordering -> frequency ->
  // basis -> rate sign -> pr sign (replaces yld) -> redemption sign.
  // The `pr <= 0` check is the YIELD-family-specific addition;
  // ODDFPRICE accepts a zero yld but ODDFYIELD rejects a non-positive
  // market price (zero would imply infinite yield).
  if (issue.value() >= settlement.value()) {
    return ErrorCode::Num;
  }
  if (settlement.value() >= first_coupon.value()) {
    return ErrorCode::Num;
  }
  if (first_coupon.value() >= maturity.value()) {
    return ErrorCode::Num;
  }
  const int frequency = frequency_e.value();
  const int basis = basis_e.value();
  if (rate.value() < 0.0) {
    return ErrorCode::Num;
  }
  if (pr.value() <= 0.0) {
    return ErrorCode::Num;
  }
  if (redemption.value() <= 0.0) {
    return ErrorCode::Num;
  }

  // Build the schedule once for the initial-guess heuristic. The
  // Newton iterate calls `price_at` (which rebuilds the schedule
  // each time) -- the schedule's structure is independent of yld so
  // recomputing it per iteration is cheap relative to the
  // sum-of-powers in the price kernel.
  auto sched_e = compute_odd_first_schedule(settlement.value(), maturity.value(), issue.value(), first_coupon.value(),
                                            frequency, basis);
  if (!sched_e) {
    return sched_e.error();
  }
  const OddFirstSchedule& sched = sched_e.value();

  const double freq_d = static_cast<double>(frequency);
  const double red = redemption.value();
  const double pr_v = pr.value();

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
  double yld = ((100.0 * rate.value() + (red - pr_v) / n_total) / avg_capital) * freq_d;
  if (std::isnan(yld) || std::isinf(yld)) {
    return ErrorCode::Num;
  }

  constexpr int kMaxIter = 100;
  // Convergence: function tolerance scales with |pr|+1 so the
  // criterion remains meaningful across the full bond price range;
  // step tolerance catches the case where Newton has flat-lined at
  // floating-point resolution.
  const double f_tol = 1.0e-12 * (std::fabs(pr_v) + 1.0);
  constexpr double kStepTol = 1.0e-15;
  for (int iter = 0; iter < kMaxIter; ++iter) {
    auto f0 = price_at(args, arity, yld);
    if (!f0) {
      return f0.error();
    }
    const double residual = f0.value() - pr_v;
    if (std::fabs(residual) < f_tol) {
      if (std::isnan(yld) || std::isinf(yld)) {
        return ErrorCode::Num;
      }
      return yld;
    }
    // Central-difference derivative. Step size scales with |yld| to
    // keep relative precision near the limit of double-precision
    // arithmetic without underflowing for tiny yields.
    const double h = 1.0e-7 * std::max(1.0, std::fabs(yld));
    auto f_plus = price_at(args, arity, yld + h);
    if (!f_plus) {
      return f_plus.error();
    }
    auto f_minus = price_at(args, arity, yld - h);
    if (!f_minus) {
      return f_minus.error();
    }
    const double df = (f_plus.value() - f_minus.value()) / (2.0 * h);
    if (df == 0.0 || std::isnan(df) || std::isinf(df)) {
      return ErrorCode::Num;
    }
    const double delta = residual / df;
    const double new_yld = yld - delta;
    if (std::isnan(new_yld) || std::isinf(new_yld)) {
      return ErrorCode::Num;
    }
    if (std::fabs(delta) < kStepTol) {
      return new_yld;
    }
    yld = new_yld;
  }
  // Iteration cap reached without convergence. Excel surfaces this as
  // `#NUM!` (caller can retry with a different price if needed).
  return ErrorCode::Num;
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
