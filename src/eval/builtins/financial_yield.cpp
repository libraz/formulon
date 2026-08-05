//
// Implementation of the regular-period bond yield-to-maturity built-in:
//
//   * YIELD(settlement, maturity, rate, pr, redemption, frequency,
//           [basis=0])
//
// Returns the annual yield (decimal) implied by the given clean market
// price `pr`. YIELD is the analytic inverse of PRICE: feeding YIELD's
// output back into PRICE (with the same settlement/maturity/rate/...
// arguments) recovers `pr` to ~1e-9.
//
// The numerical strategy mirrors Microsoft's documented YIELD:
//
//   * n == 1 (only one coupon left -- settlement falls inside the final
//     coupon period). PRICE in this branch uses simple-interest
//     discounting on the single remaining cash flow, which inverts in
//     closed form. Solving
//
//       pr = (redemption + cf) / (1 + t1*(yld/freq)) - AI
//
//     for yld gives
//
//       yld = ((redemption + cf - pr - AI) / (pr + AI)) * (E / DSC) * freq
//
//     where cf = 100*rate/freq, AI = 100*rate*A/(E*freq),
//     t1 = DSC/E. (Notation: A = days_bs, E = period_days, DSC = days_nc.)
//     A zero `pr + AI` denominator surfaces as `#NUM!`.
//
//   * n > 1. The PRICE branch sums n discounted cash flows by v^(t1+i)
//     and has no analytic inverse. We use Newton-Raphson on
//
//       f(yld) = compute_clean_price(yld) - pr
//
//     with a numerical central-difference derivative
//
//       f'(yld) ~= (f(yld + h) - f(yld - h)) / (2h)
//
//     where h = 1e-7 * max(1, |yld|). Convergence: |f(yld)| <
//     1e-12 * (|pr| + 1) OR |delta| < 1e-15 within `kMaxIter` (= 100)
//     iterations. The initial guess is Microsoft's documented "approximate
//     yield" formula
//
//       yld_0 = ((100*rate + (redemption - pr)/n) / ((redemption + pr)/2))
//                * frequency
//
//     which is well-conditioned across the entire valid bond domain --
//     premiums (yld < rate -> pr > redemption), discounts, and zero
//     coupons all converge in fewer than 10 iterations on every observed
//     case. Non-convergence (degenerate derivative or 100-iter timeout)
//     surfaces as `#NUM!`.
//
// The clean-price kernel itself is shared with PRICE via
// `financial_clean_price.h` -- both functions agree by construction.

#include "eval/builtins/financial_yield.h"

#include <algorithm>
#include <cmath>
#include <cstdint>

#include "eval/builtins/financial_clean_price.h"
#include "eval/builtins/financial_helpers.h"
#include "eval/coupon_schedule.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {
namespace {

// Evaluates `compute_clean_price` at the candidate yld by reusing
// YIELD's own argument vector with slot 3 swapped for the iterate.
//
// PRICE's arg layout is `(settlement, maturity, rate, yld, redemption,
// frequency, [basis])`; YIELD's layout is `(settlement, maturity, rate,
// pr, redemption, frequency, [basis])`. Only slot 3 differs, so we copy
// the (already-validated) Values verbatim and substitute the yld slot.
//
// `Value`'s default constructor is private, so we initialise the buffer
// from existing valid Values (slot 0) before assigning into each slot;
// this avoids any blank / uninitialised states. The buffer is sized 7
// to cover the optional basis without branching; `arity` is forwarded
// unchanged so `read_optional_number` inside `compute_clean_price`
// selects the correct value.
Expected<double, ErrorCode> price_at(const Value* yield_args, std::uint32_t arity, double yld_iterate) {
  Value buf[7] = {yield_args[0], yield_args[0], yield_args[0], yield_args[0],
                  yield_args[0], yield_args[0], yield_args[0]};
  buf[1] = yield_args[1];
  buf[2] = yield_args[2];
  buf[3] = Value::number(yld_iterate);
  buf[4] = yield_args[4];
  buf[5] = yield_args[5];
  if (arity == 7) {
    buf[6] = yield_args[6];
  }
  return compute_clean_price(buf, arity);
}

// Computes the YIELD result. Returns the yield-to-maturity (decimal) on
// success or `ErrorCode::Num` on any validation / numerical failure.
Expected<double, ErrorCode> compute_yield(const Value* args, std::uint32_t arity) {
  auto settlement = read_financial_date(args, 0);
  if (!settlement) {
    return settlement.error();
  }
  auto maturity = read_financial_date(args, 1);
  if (!maturity) {
    return maturity.error();
  }
  auto rate = read_required_number(args, 2);
  if (!rate) {
    return rate.error();
  }
  auto pr = read_required_number(args, 3);
  if (!pr) {
    return pr.error();
  }
  auto redemption = read_required_number(args, 4);
  if (!redemption) {
    return redemption.error();
  }
  auto frequency_e = read_coupon_frequency(args, 5);
  if (!frequency_e) {
    return frequency_e.error();
  }
  auto basis_e = read_day_count_basis(args, arity, 6);
  if (!basis_e) {
    return basis_e.error();
  }

  // Validation order mirrors PRICE: date ordering -> frequency domain ->
  // basis domain -> rate sign -> pr sign (replaces yld) -> redemption
  // sign. The `pr <= 0` check is the YIELD-specific addition; PRICE
  // accepts a zero yld but YIELD rejects a non-positive market price
  // (zero would mean "infinite yield").
  if (settlement.value() >= maturity.value()) {
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

  CouponDates cd{};
  if (!compute_coupon_dates(settlement.value(), maturity.value(), frequency, basis, &cd)) {
    return ErrorCode::Num;
  }
  if (cd.coupons_remaining <= 0 || cd.period_days <= 0.0) {
    return ErrorCode::Num;
  }

  const double freq_d = static_cast<double>(frequency);
  const double cf = 100.0 * rate.value() / freq_d;
  const double ai = 100.0 * rate.value() * cd.days_bs / (cd.period_days * freq_d);
  const std::int32_t n = cd.coupons_remaining;
  const double red = redemption.value();
  const double pr_v = pr.value();

  // --- n == 1: closed-form analytic inversion of the simple-interest
  // discount branch. See the file-level comment for the derivation.
  // We use `bond_dsc` (not the raw `days_nc`) because for bases 2 / 3
  // the bond formula needs `days_bs + DSC == period_days`; see
  // `coupon_schedule.h` for the rationale.
  if (n == 1) {
    const double denom = pr_v + ai;
    if (denom == 0.0) {
      return ErrorCode::Num;
    }
    if (cd.bond_dsc <= 0.0) {
      return ErrorCode::Num;
    }
    const double yld = ((red + cf - pr_v - ai) / denom) * (cd.period_days / cd.bond_dsc) * freq_d;
    if (std::isnan(yld) || std::isinf(yld)) {
      return ErrorCode::Num;
    }
    return yld;
  }

  // --- n > 1: Newton-Raphson on f(yld) = price(yld) - pr.
  //
  // Initial guess uses Microsoft's documented approximate-yield formula:
  //
  //   yld_0 = ((100*rate + (redemption - pr)/n) / ((redemption + pr)/2))
  //            * frequency
  //
  // The numerator is the per-period income (coupon + amortised
  // capital gain/loss); the denominator is the average invested
  // capital. Multiplying by `frequency` annualises a per-period rate.
  // This guess is within a few percent of the true yield across the
  // entire bond domain, including premiums and discounts, so the
  // Newton iteration converges fast.
  const double avg_capital = (red + pr_v) / 2.0;
  if (avg_capital == 0.0) {
    return ErrorCode::Num;
  }
  double yld = ((100.0 * rate.value() + (red - pr_v) / static_cast<double>(n)) / avg_capital) * freq_d;
  if (std::isnan(yld) || std::isinf(yld)) {
    return ErrorCode::Num;
  }

  constexpr int kMaxIter = 100;
  // Convergence: function tolerance is scaled by |pr|+1 so the criterion
  // remains meaningful across the full bond price range; step tolerance
  // catches the case where Newton has flat-lined at floating-point
  // resolution.
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
    // Central-difference derivative. The step size is scaled with |yld|
    // to keep relative precision near the limit of double-precision
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

}  // namespace

// --- YIELD(settlement, maturity, rate, pr, redemption, frequency,
//           [basis=0]) -----------------------------------------------------
//
// Annual yield-to-maturity (decimal) for a security paying periodic
// interest. The analytic inverse of PRICE.
Value Yield(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto y = compute_yield(args, arity);
  if (!y) {
    return Value::error(y.error());
  }
  return finalize(y.value());
}

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon
