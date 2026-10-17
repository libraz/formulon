//
// Implementation of the eager financial built-ins: PV, FV, PMT, NPER,
// NPV, RATE, IPMT, PPMT, CUMIPMT, CUMPRINC. The depreciation family
// (SLN / SYD / DDB / DB) lives in `financial_depreciation.cpp`; the
// DOLLARDE / DOLLARFR / EFFECT / NOMINAL / FVSCHEDULE / PDURATION /
// RRI / ISPMT group lives in `financial_misc.cpp`. All three
// translation units share the `read_required_number` / `finalize` /
// `pmt_scalar` / ... helpers defined inline in
// `builtins/financial_helpers.h`.
//
// Each function follows Excel's time-value-of-money sign convention
// (cash out = negative, cash in = positive) and defaults to
// end-of-period payments (`type = 0`).
//
// IRR is intentionally absent — it needs the un-flattened AST of its
// values argument to walk range/Ref/ArrayLiteral shapes, so it lives on
// the lazy-dispatch seam in `eval/financial_lazy.cpp`. The date1904-
// sensitive bond family (AMORDEGRC/AMORLINC, ACCRINT/ACCRINTM, DISC/
// INTRATE/RECEIVED/TBILL*, PRICE*/YIELD*, DURATION/MDURATION, ODDF*/
// ODDL*, and COUPPCD..COUPDAYS) is absent for the same reason `IRR` is:
// the eager calling convention here cannot carry
// `EvalContext::date1904()`, so those route through the lazy dispatch
// table too (`eval/financial_lazy.h`'s `find_financial_date_entry`).

#include "eval/builtins/financial.h"

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <utility>

#include "eval/builtins/financial_bond_simple.h"
#include "eval/builtins/financial_helpers.h"
#include "eval/builtins/registration_helpers.h"
#include "eval/coerce.h"
#include "eval/function_registry.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

using financial_detail::annuity_h;
using financial_detail::annuity_k;
using financial_detail::annuity_v;
using financial_detail::finalize;
using financial_detail::ipmt_scalar;
using financial_detail::normalize_type;
using financial_detail::payment_factor;
using financial_detail::pmt_scalar;
using financial_detail::ppmt_scalar;
using financial_detail::read_optional_number;
using financial_detail::read_required_number;
using financial_detail::read_required_numbers;

struct TvmArgs {
  double first;
  double second;
  double third;
  double fourth;
  double type;
};

struct RateArgs {
  double nper;
  double pmt;
  double pv;
  double fv;
  double type;
  double guess;
};

struct PaymentArgs {
  double rate;
  double per;
  double nper;
  double pv;
  double fv;
  double type;
};

struct CumPaymentArgs {
  double rate;
  double nper;
  double pv;
  double start;
  double end;
  double type;
};

Expected<TvmArgs, ErrorCode> read_tvm_args(const Value* args, std::uint32_t arity, double fourth_default) {
  double v[3];
  if (auto read = read_required_numbers(args, "nnn", v); !read) {
    return std::move(read.error());
  }
  auto fourth = read_optional_number(args, arity, 3, fourth_default);
  if (!fourth) {
    return std::move(fourth.error());
  }
  auto type = read_optional_number(args, arity, 4, 0.0);
  if (!type) {
    return std::move(type.error());
  }
  return TvmArgs{v[0], v[1], v[2], fourth.value(), normalize_type(type.value())};
}

Expected<RateArgs, ErrorCode> read_rate_args(const Value* args, std::uint32_t arity) {
  auto tvm = read_tvm_args(args, arity, 0.0);
  if (!tvm) {
    return std::move(tvm.error());
  }
  auto guess = read_optional_number(args, arity, 5, 0.1);
  if (!guess) {
    return std::move(guess.error());
  }
  return RateArgs{tvm.value().first,  tvm.value().second, tvm.value().third,
                  tvm.value().fourth, tvm.value().type,   guess.value()};
}

Expected<PaymentArgs, ErrorCode> read_payment_args(const Value* args, std::uint32_t arity) {
  double v[4];
  if (auto read = read_required_numbers(args, "nnnn", v); !read) {
    return std::move(read.error());
  }
  auto fv = read_optional_number(args, arity, 4, 0.0);
  if (!fv) {
    return std::move(fv.error());
  }
  auto type = read_optional_number(args, arity, 5, 0.0);
  if (!type) {
    return std::move(type.error());
  }
  return PaymentArgs{v[0], v[1], v[2], v[3], fv.value(), normalize_type(type.value())};
}

Expected<CumPaymentArgs, ErrorCode> read_cum_payment_args(const Value* args) {
  for (std::uint32_t i = 0; i < 6; ++i) {
    if (args[i].kind() == ValueKind::Bool) {
      return ErrorCode::Value;
    }
  }
  double v[6];
  if (auto read = read_required_numbers(args, "nnnnnn", v); !read) {
    return std::move(read.error());
  }
  if (v[5] != 0.0 && v[5] != 1.0) {
    return ErrorCode::Num;
  }
  return CumPaymentArgs{v[0], v[1], v[2], v[3], v[4], normalize_type(v[5])};
}

Expected<std::pair<std::int64_t, std::int64_t>, ErrorCode> cum_period_bounds(double start, double end, double nper) {
  if (start < 1.0 || end < 1.0 || start > end || start > nper || end > nper) {
    return ErrorCode::Num;
  }
  // Mac Excel 365 rounds `start` UP (ceil) and `end` DOWN (floor) before
  // iterating; see CUMIPMT / CUMPRINC oracle fixtures.
  // Clamped to 2^53 so a huge finite bound narrows without overflow.
  constexpr double kMaxExactPeriod = 9007199254740992.0;
  const auto start_i = static_cast<std::int64_t>(std::ceil(std::min(start, kMaxExactPeriod)));
  const auto end_i = static_cast<std::int64_t>(std::floor(std::min(end, kMaxExactPeriod)));
  if (start_i > end_i) {
    return ErrorCode::Num;
  }
  return std::pair<std::int64_t, std::int64_t>{start_i, end_i};
}

// --- PV(rate, nper, pmt, [fv=0], [type=0]) ------------------------------
//
// Matches Excel's documented formula:
//
//   rate == 0  ->  PV = -(pmt * nper + fv)
//   rate != 0  ->  f  = (1 + rate)^nper   (POWER's integer squaring)
//                  PV = (-fv - pmt * (1 + rate*type) * ((f - 1) / rate)) / f
//
// Sign convention follows Excel: cash-out values are negative, cash-in
// values are positive. Non-finite intermediate or final values surface as
// `#NUM!`.
Value Pv(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto input = read_tvm_args(args, arity, 0.0);
  if (!input) {
    return Value::error(input.error());
  }
  const auto [r, n, p, f, t] = input.value();
  if (r == -1.0) {
    // (1+r)^n is 0 for n>0 (yields division by 0 inside the closed form)
    // or infinite for n<=0; Excel collapses both to #DIV/0!.
    return Value::error(ErrorCode::Div0);
  }
  if (r == 0.0) {
    return finalize(-(p * n + f));
  }
  auto grown = apply_pow(1.0 + r, n);
  if (!grown) {
    return Value::error(grown.error() == ErrorCode::Div0 ? ErrorCode::Div0 : ErrorCode::Num);
  }
  const double pow_term = grown.value();
  if (pow_term == 0.0) {
    // An underflowed discount factor divides by zero like rate == -1 does.
    return Value::error(ErrorCode::Div0);
  }
  const double result = (-f - p * (1.0 + r * t) * ((pow_term - 1.0) / r)) / pow_term;
  return finalize(result);
}

// --- FV(rate, nper, pmt, [pv=0], [type=0]) ------------------------------
//
//   rate == 0  ->  FV = -(pv + pmt * nper)
//   rate != 0  ->  f  = (1 + rate)^nper
//                  FV = -(pv * f + pmt * (1 + rate*type) * ((f - 1) / rate))
Value Fv(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto input = read_tvm_args(args, arity, 0.0);
  if (!input) {
    return Value::error(input.error());
  }
  const auto [r, n, p, v, t] = input.value();
  if (r == -1.0) {
    // (1+r)^n is 0 for n>0 (zero payment discount factor) or infinite for
    // n<=0; Excel collapses both to #DIV/0!.
    return Value::error(ErrorCode::Div0);
  }
  if (r == 0.0) {
    return finalize(-(v + p * n));
  }
  auto grown = apply_pow(1.0 + r, n);
  if (!grown) {
    return Value::error(grown.error() == ErrorCode::Div0 ? ErrorCode::Div0 : ErrorCode::Num);
  }
  const double pow_term = grown.value();
  const double result = -(v * pow_term + p * (1.0 + r * t) * ((pow_term - 1.0) / r));
  return finalize(result);
}

// --- PMT(rate, nper, pv, [fv=0], [type=0]) ------------------------------
//
// See `pmt_scalar`; nper == 0 surfaces as `#NUM!`.
Value Pmt(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto input = read_tvm_args(args, arity, 0.0);
  if (!input) {
    return Value::error(input.error());
  }
  const auto [r, n, v, f, t] = input.value();
  if (r <= -1.0) {
    // Mac Excel 365 rejects rate <= -1 for PMT outright with #NUM!, even
    // when the closed form would evaluate to a finite value (e.g. rate=-3
    // with integer nper). This differs from FV/PV which do compute at
    // rate < -1 when nper is an integer (verified via IronCalc xlsx
    // oracle).
    return Value::error(ErrorCode::Num);
  }
  return finalize(pmt_scalar(r, n, v, f, t));
}

// --- NPER(rate, pmt, pv, [fv=0], [type=0]) ------------------------------
//
//   rate == 0  ->  pmt == 0 -> #NUM!, otherwise NPER = -(pv + fv) / pmt
//
//   rate != 0  ->  type == 0:  num = pmt - fv * rate
//                               den = pmt + pv * rate
//                  type == 1:  num = pmt * (1 + rate) - fv * rate
//                               den = pmt * (1 + rate) + pv * rate
//                  NPER = log(num / den) / log(1 + rate)
//                  The quotient num/den must be strictly positive (log
//                  domain); a non-positive quotient surfaces as `#NUM!`.
//
// Note: Excel 365 Mac does *not* clamp NPER to the non-negative half line —
// a sign-mismatched loan such as `=NPER(0.05/12, 500, 25000)` (both pmt
// and pv positive) produces the raw negative algebraic answer (~-45.51),
// not `#NUM!`. The oracle confirmed this. We therefore return whatever
// the log formula yields as long as the input domain is valid.
Value Nper(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto input = read_tvm_args(args, arity, 0.0);
  if (!input) {
    return Value::error(input.error());
  }
  const auto [r, p, v, f, t] = input.value();
  if (r == 0.0) {
    if (p == 0.0) {
      return Value::error(ErrorCode::Num);
    }
    return finalize(-(v + f) / p);
  }
  if (r <= -1.0) {
    // log(1 + rate) is undefined when rate <= -1; Excel returns #NUM!.
    return Value::error(ErrorCode::Num);
  }
  const double pmt_scaled = t == 0.0 ? p : p * (1.0 + r);
  const double num = pmt_scaled - f * r;
  const double den = pmt_scaled + v * r;
  if (den == 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double ratio = num / den;
  if (ratio <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double result = std::log(ratio) / std::log(1.0 + r);
  return finalize(result);
}

// --- NPV(rate, value1, value2, ...) -------------------------------------
//
// Discounts each value by its 1-based positional index `i` (NPV's rate is
// applied from period 1, unlike IRR which uses period 0..n-1):
//
//   NPV = sum_{i=1..n} value_i / (1 + rate)^i
//
// When `accepts_ranges = true` with `range_filter_numeric_only = true`,
// the dispatcher flattens each RangeOp into consecutive numeric cells in
// row-major order; the positional counter runs across the concatenated
// flat sequence (scalars in argument order, ranges row-major, continuing
// the index). Text / Bool / Blank cells inside a range are filtered out
// by the dispatcher before reaching this impl. Direct scalar arguments
// still coerce through `coerce_to_number`, so a direct bool argument
// becomes 1.0 / 0.0 and a direct blank becomes 0.0 — matching Excel.
Value Npv(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  // arity guaranteed >= 2 by min_arity in register_financial_builtins.
  auto rate_e = read_required_number(args, 0);
  if (!rate_e) {
    return Value::error(rate_e.error());
  }
  const double rate = rate_e.value();
  const double discount = 1.0 + rate;
  double total = 0.0;
  // Walk value1..valueN. Range-sourced bool/text/blank cells were already
  // dropped by the dispatcher's numeric-only filter; a direct scalar
  // logical / numeric-text argument is coerced and contributes at its
  // period (see the per-iteration comment). `period_discount` only steps
  // forward when a cash flow is actually consumed. Each flow is divided by the
  // growth factor of the period before it; the sum is discounted once at the end.
  double period_discount = 1.0;
  for (std::uint32_t i = 1; i < arity; ++i) {
    const Value& v = args[i];
    if (v.is_error()) {
      return v;
    }
    // The dispatcher already dropped range-sourced Bool / Text / Blank
    // cells (range_filter_numeric_only), so any non-Number reaching this
    // impl is a DIRECT scalar argument. Excel counts a directly-passed
    // logical / numeric-text argument (TRUE -> 1, "5" -> 5) at its period;
    // only non-numeric text is ignored. Coerce here so the period counter
    // advances for those, matching Excel.
    double cash = 0.0;
    if (v.is_number()) {
      cash = v.as_number();
    } else {
      auto coerced = coerce_to_number(v);
      if (!coerced) {
        continue;  // non-numeric text: ignored, period unchanged.
      }
      cash = coerced.value();
    }
    total = total + cash / period_discount;
    period_discount = period_discount * discount;
  }
  return finalize(total / discount);
}

// --- RATE(nper, pmt, pv, [fv=0], [type=0], [guess=0.1]) -----------------
//
// Solves f(r) = pv*q + pmt*(((q-1)/r)*(1+r*type)) + fv = 0 with q = (1+r)^nper
// by the secant method, step for step as Excel does:
//   - q is `financial_detail::excel_power(1+r, nper)`; f(0) is the linear form;
//   - the second point is guess -/+ 0.001 against the sign of f(guess);
//   - at most 200 secant steps, accepted when both |dx| and |f| are < 1e-7;
//     after the cap the last iterate is kept only if |f| < 1e-4;
//   - iterates are not domain-guarded; a repeated f, or a root with
//     1 + r <= 1e-7, is #NUM!.
double rate_residual(double r, double nper, double pmt, double pv, double fv, double type) {
  if (r == 0.0) {
    return pv + pmt * nper + fv;
  }
  const double q = financial_detail::excel_power(1.0 + r, nper);
  return pv * q + pmt * (((q - 1.0) / r) * (1.0 + r * type)) + fv;
}

Value Rate(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto input = read_rate_args(args, arity);
  if (!input) {
    return Value::error(input.error());
  }
  const double nper = input.value().nper;
  const double pmt = input.value().pmt;
  const double pv = input.value().pv;
  const double fv = input.value().fv;
  const double type = input.value().type;

  if (nper < 1.0) {
    return Value::error(ErrorCode::Num);
  }

  constexpr int kMaxSteps = 200;
  constexpr double kConvergence = 1.0e-7;
  constexpr double kCapResidual = 1.0e-4;
  constexpr double kRateFloor = 1.0e-7;
  constexpr double kSecondPointOffset = 1.0e-3;
  const bool fractional = std::trunc(nper) != nper;
  const auto residual = [&](double r) { return rate_residual(r, nper, pmt, pv, fv, type); };
  const auto out_of_domain = [&](double r) { return fractional && !(1.0 + r > 0.0); };

  double x0 = input.value().guess;
  if (out_of_domain(x0)) {
    return Value::error(ErrorCode::Num);
  }
  double f0 = residual(x0);
  double x1 = f0 > 0.0 ? x0 - kSecondPointOffset : x0 + kSecondPointOffset;
  if (out_of_domain(x1)) {
    return Value::error(ErrorCode::Num);
  }
  double f1 = residual(x1);
  double x2 = x1;
  double f2 = f1;
  bool converged = false;
  for (int step = 0; step < kMaxSteps; ++step) {
    if (f1 == f0) {
      return Value::error(ErrorCode::Num);
    }
    x2 = x1 - f1 * (x1 - x0) / (f1 - f0);
    if (out_of_domain(x2)) {
      return Value::error(ErrorCode::Num);
    }
    f2 = residual(x2);
    if (std::fabs(x2 - x1) < kConvergence && std::fabs(f2) < kConvergence) {
      converged = true;
      break;
    }
    x0 = x1;
    f0 = f1;
    x1 = x2;
    f1 = f2;
  }
  if (!converged && !(std::fabs(f2) < kCapResidual)) {
    return Value::error(ErrorCode::Num);
  }
  if (!(1.0 + x2 > kRateFloor)) {
    return Value::error(ErrorCode::Num);
  }
  return finalize(x2);
}

// --- IPMT(rate, per, nper, pv, [fv=0], [type=0]) ------------------------
//
// Interest portion of the period-`per` payment on a standard annuity
// schedule. `per` is 1-based and must fall within `[1, nper]`; otherwise
// `#NUM!`. See `ipmt_scalar` above for the formula details.
Value Ipmt(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto input = read_payment_args(args, arity);
  if (!input) {
    return Value::error(input.error());
  }
  const double result = ipmt_scalar(input.value().rate, input.value().per, input.value().nper, input.value().pv,
                                    input.value().fv, input.value().type);
  return finalize(result);
}

// --- PPMT(rate, per, nper, pv, [fv=0], [type=0]) ------------------------
//
// Principal portion of the period-`per` payment; see `ppmt_scalar`. PPMT
// shares IPMT's `per` domain.
Value Ppmt(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto input = read_payment_args(args, arity);
  if (!input) {
    return Value::error(input.error());
  }
  const auto& in = input.value();
  return finalize(ppmt_scalar(in.rate, in.per, in.nper, in.pv, in.fv, in.type));
}

// Principal sum S over the periods [start, end] (`principal`) or the interest
// sum (CUMIPMT = payment * d - S, d the signed period count). Type 1 starting
// at period 1 takes the first payment out of S.
double cum_payment_closed(double rate, double nper, double pv, double start, double end, double type, bool principal) {
  const bool first = type != 0.0 && start == 1.0;
  const double f = payment_factor(rate, type);
  const double k = annuity_k(rate, nper, pv, 0.0);
  const double payment = k * f;
  double d = start - end;
  if (!first) {
    d = d - 1.0;
  }
  const double g = type != 0.0 ? 1.0 : 1.0 + rate;
  const double sum = ((annuity_h(rate, d) * annuity_v(rate, (end - nper) - 1.0)) * k) * g;
  if (!principal) {
    return payment * d - sum;
  }
  return first ? sum - payment : sum;
}

// Shared body of CUMIPMT / CUMPRINC: validates the common domain and totals
// either the interest or the principal portion over [start, end] in closed form.
Value cum_payment_sum(const Value* args, bool principal) {
  auto input = read_cum_payment_args(args);
  if (!input) {
    return Value::error(input.error());
  }
  const double rate = input.value().rate;
  const double nper = input.value().nper;
  const double pv = input.value().pv;
  if (rate <= 0.0 || nper <= 0.0 || pv <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  auto bounds = cum_period_bounds(input.value().start, input.value().end, nper);
  if (!bounds) {
    return Value::error(bounds.error());
  }
  const std::int64_t start_i = bounds.value().first;
  const std::int64_t end_i = bounds.value().second;
  if (static_cast<double>(end_i - start_i) + 1.0 > financial_detail::kMaxDepreciationPeriods) {
    return Value::error(ErrorCode::Num);
  }
  return finalize(cum_payment_closed(rate, nper, pv, static_cast<double>(start_i), static_cast<double>(end_i),
                                     input.value().type, principal));
}

// --- CUMIPMT(rate, nper, pv, start, end, type) --------------------------
//
// Sum of interest paid from period `start` to `end` inclusive. Unlike the
// PMT family, `type` is REQUIRED here (6-arity, not 5+optional) —
// matching Excel's signature.
//
// Domain checks per Microsoft docs:
//   - rate  > 0
//   - nper  > 0
//   - pv    > 0
//   - 1 <= start <= end <= nper
// Violating any check returns `#NUM!`.
Value Cumipmt(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return cum_payment_sum(args, /*principal=*/false);
}

// --- CUMPRINC(rate, nper, pv, start, end, type) -------------------------
//
// Sum of principal paid from period `start` to `end` inclusive. Same
// domain contract as CUMIPMT.
Value Cumprinc(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return cum_payment_sum(args, /*principal=*/true);
}

}  // namespace

void register_financial_builtins(FunctionRegistry& registry) {
  // AMORDEGRC / AMORLINC, ACCRINT / ACCRINTM, DISC / INTRATE / RECEIVED /
  // TBILL*, PRICE* / YIELD* (regular, discount, maturity, and odd-period
  // variants), and DURATION / MDURATION are NOT registered here: they
  // decompose a date serial and so need the workbook's
  // `EvalContext::date1904()`, unreachable through this eager calling
  // convention. They route through the lazy dispatch table instead (see
  // `eval/financial_lazy.h`'s `find_financial_date_entry`).
  static constexpr builtins_detail::BuiltinRegistration functions[] = {
      {"PV", 3u, 5u, &Pv},
      {"FV", 3u, 5u, &Fv},
      {"PMT", 3u, 5u, &Pmt},
      {"NPER", 3u, 5u, &Nper},
      {"NPV", 2u, kVariadic, &Npv, true, true, true},
      {"RATE", 3u, 6u, &Rate},
      {"IPMT", 4u, 6u, &Ipmt},
      {"PPMT", 4u, 6u, &Ppmt},
      builtins_detail::analysis_toolpak({"CUMIPMT", 6u, 6u, &Cumipmt}),
      builtins_detail::analysis_toolpak({"CUMPRINC", 6u, 6u, &Cumprinc}),
      {"SLN", 3u, 3u, &financial_detail::Sln},
      {"SYD", 4u, 4u, &financial_detail::Syd},
      {"DDB", 4u, 5u, &financial_detail::Ddb},
      {"DB", 4u, 5u, &financial_detail::Db},
      {"VDB", 5u, 7u, &financial_detail::Vdb},
      builtins_detail::analysis_toolpak({"DOLLARDE", 2u, 2u, &financial_detail::DollarDe}),
      builtins_detail::analysis_toolpak({"DOLLARFR", 2u, 2u, &financial_detail::DollarFr}),
      builtins_detail::analysis_toolpak({"EFFECT", 2u, 2u, &financial_detail::Effect}),
      builtins_detail::analysis_toolpak({"NOMINAL", 2u, 2u, &financial_detail::Nominal}),
      {"PDURATION", 3u, 3u, &financial_detail::PDuration},
      {"RRI", 3u, 3u, &financial_detail::Rri},
      {"ISPMT", 4u, 4u, &financial_detail::IsPmt},
      builtins_detail::analysis_toolpak({"FVSCHEDULE", 2u, kVariadic, &financial_detail::FvSchedule, true, true, true}),
      {"STOCKHISTORY", 2u, kVariadic, &financial_detail::StockHistory},
  };
  builtins_detail::register_builtin_functions(registry, functions, sizeof(functions) / sizeof(functions[0]));
}

}  // namespace eval
}  // namespace formulon
