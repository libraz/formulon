//
// Implementation of the IRR / MIRR / XIRR / XNPV lazy impls. See
// `eval/financial_lazy.h` for the dispatch-table contract and
// `eval/lazy_impls.h` for the shared `eval_node` / `LazyImpl` vocabulary.

#include "eval/financial_lazy.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/builtins/financial_bond_simple.h"
#include "eval/builtins/financial_coupon.h"
#include "eval/builtins/financial_duration.h"
#include "eval/builtins/financial_helpers.h"
#include "eval/builtins/financial_oddfprice.h"
#include "eval/builtins/financial_oddfyield.h"
#include "eval/builtins/financial_oddlprice.h"
#include "eval/builtins/financial_oddlyield.h"
#include "eval/builtins/financial_price.h"
#include "eval/builtins/financial_yield.h"
#include "eval/builtins/numeric_helpers.h"
#include "eval/coerce.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "eval/omitted_arg.h"
#include "eval/range_args.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/index_sort.h"
#include "utils/strings.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

using builtins_detail::to_finite_value;

// Collects the numeric cash flows from IRR's first argument. Every shape
// `resolve_range_arg` understands is accepted — `Ref` / `RangeOp` /
// `SpillRef` / `ArrayLiteral` and dynamic-array producers such as
// `SEQUENCE` — while a bare scalar is rejected with #VALUE!, mirroring
// Excel's requirement that IRR's first argument be a reference or array.
//
// The per-cell rule then splits on where the cells came from:
//
//   - inline `ArrayLiteral` -> each element is coerced to a number and a
//                              non-numeric element yields #VALUE!
//                              (literals are direct scalars, not
//                              range-sourced cells);
//   - everything else       -> Text / Bool / Blank cells are silently
//                              skipped, matching SUM / AVERAGE on
//                              range-sourced inputs.
//
// Returns `true` on success with the cash flows written to `*out_flows`.
// On failure writes an Excel error into `*out_err` and returns `false`.
// An error cell encountered during resolution propagates as the IRR
// call's result.
bool collect_cash_flows(const parser::AstNode& arg, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx, std::vector<double>* out_flows, Value* out_err) {
  out_flows->clear();
  auto resolved = resolve_range_arg_no_scalar(arg, arena, registry, ctx, ErrorCode::Value);
  if (!resolved) {
    *out_err = Value::error(resolved.error());
    return false;
  }
  const bool literal = arg.kind() == parser::NodeKind::ArrayLiteral;
  for (const Value& v : resolved.value().cells) {
    if (v.is_error()) {
      *out_err = v;
      return false;
    }
    if (!literal) {
      if (v.is_number()) {
        out_flows->push_back(v.as_number());
      }
      continue;
    }
    auto coerced = coerce_to_number(v);
    if (!coerced) {
      *out_err = Value::error(coerced.error());
      return false;
    }
    out_flows->push_back(coerced.value());
  }
  return true;
}

struct CashFlowSigns {
  bool has_positive = false;
  bool has_negative = false;
};

CashFlowSigns scan_cash_flow_signs(const std::vector<double>& flows) noexcept {
  CashFlowSigns signs;
  for (double v : flows) {
    if (v > 0.0) {
      signs.has_positive = true;
    } else if (v < 0.0) {
      signs.has_negative = true;
    }
  }
  return signs;
}

bool has_opposite_cash_flow_signs(const std::vector<double>& flows) noexcept {
  const CashFlowSigns signs = scan_cash_flow_signs(flows);
  return signs.has_positive && signs.has_negative;
}

// Evaluates IRR's NPV function at `rate` for the cash flow sequence.
// IRR uses period 0..n-1 indexing: the first cash flow is at time 0,
// undiscounted, and the last is at time n-1.
double irr_npv(const std::vector<double>& flows, double rate) noexcept {
  const double base = 1.0 + rate;
  double total = 0.0;
  double discount = 1.0;  // (1+rate)^0
  for (double v : flows) {
    total += v / discount;
    discount *= base;
  }
  return total;
}

// Derivative of irr_npv with respect to rate. With f(r) = sum v[i] /
// (1+r)^i, f'(r) = sum -i * v[i] / (1+r)^(i+1).
double irr_dnpv(const std::vector<double>& flows, double rate) noexcept {
  const double base = 1.0 + rate;
  double total = 0.0;
  double discount = base;  // (1+rate)^1, used for i=0: derivative term is 0
  for (std::size_t i = 0; i < flows.size(); ++i) {
    if (i > 0) {
      total += -static_cast<double>(i) * flows[i] / discount;
    }
    discount *= base;
  }
  return total;
}

// Runs Newton-Raphson on irr_npv starting from `guess`. Returns NaN on
// failure (the iterate wanders past rate <= -1, derivative collapses, or
// the iteration cap is exhausted without converging); otherwise the
// converged rate. Used as the fast path for IRR; a bracket + bisection
// fallback (`irr_bracket`) picks up cases where Newton diverges.
double irr_newton(const std::vector<double>& flows, double guess) noexcept {
  constexpr int kMaxIter = 100;
  constexpr double kTolerance = 1.0e-10;
  // Residual gate on the accepted root. Looser than the step tolerance
  // because irr_npv's magnitude scales with the cashflow magnitudes; a
  // legitimate root still leaves a small absolute residual.
  constexpr double kResidualTolerance = 1.0e-6;
  double rate = guess;
  for (int iter = 0; iter < kMaxIter; ++iter) {
    if (rate <= -1.0) {
      return std::numeric_limits<double>::quiet_NaN();
    }
    const double npv = irr_npv(flows, rate);
    const double dnpv = irr_dnpv(flows, rate);
    if (dnpv == 0.0 || std::isnan(dnpv) || std::isinf(dnpv)) {
      return std::numeric_limits<double>::quiet_NaN();
    }
    const double new_rate = rate - npv / dnpv;
    if (std::isnan(new_rate) || std::isinf(new_rate)) {
      return std::numeric_limits<double>::quiet_NaN();
    }
    if (std::fabs(new_rate - rate) < kTolerance) {
      // Step convergence alone is not sufficient: on pathological
      // cashflows the iterate can stall at a point where NPV is far from
      // zero. Accept only when the residual is also near zero; otherwise
      // report non-convergence so the bracketed fallback / #NUM! path
      // takes over.
      if (std::fabs(irr_npv(flows, new_rate)) < kResidualTolerance) {
        return new_rate;
      }
      return std::numeric_limits<double>::quiet_NaN();
    }
    rate = new_rate;
  }
  return std::numeric_limits<double>::quiet_NaN();
}

// Bracketed fallback for IRR when Newton diverges or can't
// cross the rate = -1 singularity. Samples f(rate) on a log-linear grid
// spanning the domain rate > -1, looks for an adjacent sign change, and
// bisects. The final midpoint is optionally polished by the caller's
// Newton implementation for full double precision.
template <typename EvalFn, typename PolishFn>
double bracket_rate_root(EvalFn eval_at_rate, PolishFn polish_midpoint) noexcept {
  std::vector<double> grid;
  grid.reserve(40);
  for (int k = -10; k <= -1; ++k) {
    grid.push_back(-1.0 + std::pow(10.0, static_cast<double>(k)));
  }
  for (double m : {-0.8, -0.5, -0.2, 0.0, 0.1, 0.2, 0.5, 1.0, 2.0, 5.0}) {
    grid.push_back(m);
  }
  for (int k = 1; k <= 6; ++k) {
    grid.push_back(std::pow(10.0, static_cast<double>(k)));
  }
  std::sort(grid.begin(), grid.end());
  // Scan for an adjacent sign change.
  double prev_r = grid[0];
  double prev_f = eval_at_rate(prev_r);
  if (std::isfinite(prev_f) && std::fabs(prev_f) < 1.0e-10) {
    return prev_r;
  }
  for (std::size_t i = 1; i < grid.size(); ++i) {
    const double r = grid[i];
    const double f = eval_at_rate(r);
    if (!std::isfinite(f)) {
      prev_r = r;
      prev_f = f;
      continue;
    }
    if (std::fabs(f) < 1.0e-10) {
      return r;
    }
    if (std::isfinite(prev_f) && ((prev_f < 0.0 && f > 0.0) || (prev_f > 0.0 && f < 0.0))) {
      double lo = prev_r;
      double hi = r;
      double f_lo = prev_f;
      for (int iter = 0; iter < 200; ++iter) {
        const double mid = 0.5 * (lo + hi);
        const double f_mid = eval_at_rate(mid);
        if (!std::isfinite(f_mid)) {
          hi = mid;
          continue;
        }
        if (std::fabs(f_mid) < 1.0e-12) {
          return mid;
        }
        if ((f_lo < 0.0 && f_mid < 0.0) || (f_lo > 0.0 && f_mid > 0.0)) {
          lo = mid;
          f_lo = f_mid;
        } else {
          hi = mid;
        }
        if (std::fabs(hi - lo) < 1.0e-14) {
          return 0.5 * (lo + hi);
        }
      }
      const double refined = polish_midpoint(0.5 * (lo + hi));
      if (std::isfinite(refined)) {
        return refined;
      }
      return 0.5 * (lo + hi);
    }
    prev_r = r;
    prev_f = f;
  }
  return std::numeric_limits<double>::quiet_NaN();
}

double irr_bracket(const std::vector<double>& flows) noexcept {
  return bracket_rate_root([&](double rate) noexcept { return irr_npv(flows, rate); },
                           [&](double guess) noexcept { return irr_newton(flows, guess); });
}

// Closed-form MIRR helper. Positive flows discount at the reinvestment rate and
// negative ones at the finance rate (the first flow undiscounted), each into
// its own accumulator:
//
//   MIRR = POWER(-((A * F) / q) / B, 1 / (n - 1)) - 1
//
// with F = q^n the reinvestment growth factor after the walk. Returns NaN when
// `flows` holds fewer than two values or the negative accumulator is zero.
double mirr_closed_form(const std::vector<double>& flows, double finance_rate, double reinvest_rate) noexcept {
  const std::size_t n = flows.size();
  if (n < 2) {
    return std::numeric_limits<double>::quiet_NaN();
  }
  const double q = 1.0 + reinvest_rate;
  const double p = 1.0 + finance_rate;
  double npv_pos = 0.0;
  double npv_neg = 0.0;
  double pos_factor = 1.0;
  double neg_factor = 1.0;
  for (std::size_t i = 0; i < n; ++i) {
    if (flows[i] > 0.0) {
      npv_pos = npv_pos + flows[i] / pos_factor;
    } else if (flows[i] < 0.0) {
      npv_neg = npv_neg + flows[i] / neg_factor;
    }
    pos_factor = pos_factor * q;
    neg_factor = neg_factor * p;
  }
  if (npv_neg == 0.0) {
    return std::numeric_limits<double>::quiet_NaN();
  }
  const double base = (-((npv_pos * pos_factor) / q)) / npv_neg;
  return financial_detail::excel_power(base, 1.0 / static_cast<double>(n - 1)) - 1.0;
}

// Reads an optional guess argument into `*guess`; a blank leaves the caller's default so a trailing
// comma matches an omitted argument. `analysis_toolpak` (XIRR) rejects a boolean guess with #VALUE!.
// Returns false with the error in `*out_err` on error or non-numeric.
bool read_optional_guess(const parser::AstNode& arg, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx, bool analysis_toolpak, double* guess, Value* out_err) {
  const Value guess_v = eval_node(arg, arena, registry, ctx);
  if (guess_v.is_error()) {
    *out_err = guess_v;
    return false;
  }
  if (analysis_toolpak) {
    if (const Value atp = atp_arg_error(arg, guess_v, /*required=*/false); atp.is_error()) {
      *out_err = atp;
      return false;
    }
  }
  if (!guess_v.is_blank()) {
    auto coerced = coerce_to_number(guess_v);
    if (!coerced) {
      *out_err = Value::error(coerced.error());
      return false;
    }
    *guess = coerced.value();
  }
  return true;
}

}  // namespace

Value eval_irr_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                    const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 1U || arity > 2U) {
    return Value::error(ErrorCode::Value);
  }

  std::vector<double> flows;
  Value collect_err = Value::blank();
  if (!collect_cash_flows(call.as_call_arg(0), arena, registry, ctx, &flows, &collect_err)) {
    return collect_err;
  }

  // IRR requires at least one positive and one negative cash flow so the
  // NPV(r) function has a root. Otherwise the iteration would either
  // diverge or converge to a degenerate boundary (rate == -1).
  if (!has_opposite_cash_flow_signs(flows)) {
    return Value::error(ErrorCode::Num);
  }

  double rate = 0.1;  // default guess per Excel.
  Value guess_err = Value::blank();
  if (arity == 2U &&
      !read_optional_guess(call.as_call_arg(1), arena, registry, ctx, /*analysis_toolpak=*/false, &rate, &guess_err)) {
    return guess_err;
  }
  // An explicit out-of-domain guess (the NPV expansion needs 1 + rate > 0) is an input error, so
  // Excel surfaces #VALUE! rather than the #NUM! the in-loop boundary check uses.
  if (rate <= -1.0) {
    return Value::error(ErrorCode::Value);
  }

  // Try Newton-Raphson first; it converges quickly for the well-posed
  // schedules (single positive root, guess near the answer). When it
  // fails — typically because the true root lies near the rate = -1
  // singularity and Newton walks past it — fall back to a bracket +
  // bisection scan over the full rate > -1 domain. Mirrors the XIRR
  // hybrid strategy used below.
  const double newton_rate = irr_newton(flows, rate);
  if (std::isfinite(newton_rate)) {
    return Value::number(newton_rate);
  }
  const double bracket_rate = irr_bracket(flows);
  if (std::isfinite(bracket_rate)) {
    return Value::number(bracket_rate);
  }
  return Value::error(ErrorCode::Num);
}

Value eval_mirr_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity != 3U) {
    return Value::error(ErrorCode::Value);
  }

  std::vector<double> flows;
  Value collect_err = Value::blank();
  if (!collect_cash_flows(call.as_call_arg(0), arena, registry, ctx, &flows, &collect_err)) {
    return collect_err;
  }

  // All-positive schedules drive `npv_neg` to 0, so the closed-form
  // division would divide by zero — Excel surfaces this as `#DIV/0!`,
  // not `#NUM!` as for IRR. All-*negative* schedules, by contrast, are
  // well-defined: `npv_pos == 0` makes `ratio == 0` and `pow(0, x) - 1`
  // yields exactly -1.0 for any positive x. We therefore pre-check only
  // the all-positive case and let the math carry for all-negative.
  if (!scan_cash_flow_signs(flows).has_negative) {
    return Value::error(ErrorCode::Div0);
  }

  // Evaluate the two trailing rate arguments. Errors propagate directly
  // (we bypass the eager dispatcher, so the short-circuit is manual).
  const Value finance_v = eval_node(call.as_call_arg(1), arena, registry, ctx);
  if (finance_v.is_error()) {
    return finance_v;
  }
  auto finance = coerce_to_number(finance_v);
  if (!finance) {
    return Value::error(finance.error());
  }
  const Value reinvest_v = eval_node(call.as_call_arg(2), arena, registry, ctx);
  if (reinvest_v.is_error()) {
    return reinvest_v;
  }
  auto reinvest = coerce_to_number(reinvest_v);
  if (!reinvest) {
    return Value::error(reinvest.error());
  }

  const double result = mirr_closed_form(flows, finance.value(), reinvest.value());
  return to_finite_value(result);
}

// ---------------------------------------------------------------------------
// XIRR / XNPV
// ---------------------------------------------------------------------------
//
// These two share three concerns that IRR / MIRR do not have:
//
//   1. A parallel `dates` range that must iterate in lockstep with
//      `values`. Excel enforces equal total cell counts between the two;
//      we implement this as a row-major flatten of each argument and a
//      hard size check before any math runs.
//   2. A per-pair filter whose rules diverge between the two functions.
//      The asymmetry is empirical, measured against the IronCalc oracle:
//        * XIRR silently drops a pair whose value cell is Blank, sorts
//          the surviving pairs by date ascending (so the schedule is
//          normalised even when the input columns are shuffled), and
//          rejects non-numeric value cells (Bool, Text) with #VALUE!.
//        * XNPV rejects any Blank or non-numeric value cell with #NUM!,
//          matching Excel's habit of folding malformed XNPV inputs into
//          a single catch-all error. It also rejects rates < 0 and any
//          schedule with a negative date.
//      A blank date cell disqualifies the pair in both functions.
//   3. A date-to-offset transform: `t_i = (dates[i] - dates[0]) / 365`,
//      with `dates[i]` truncated toward zero (`std::trunc`) to match
//      Excel's serial-date conversion rule. Dates are taken from the
//      sorted schedule, so `dates[0]` is always the earliest surviving
//      date.
//
// The two impls share the collection pipeline so a single bug in the
// filter cannot drift XIRR away from XNPV.

namespace {

// Which of the two functions is driving the (value, date) collection.
// The per-cell rules diverge in Bool / Text / Blank handling and XNPV
// additionally enforces a rate >= 0 check upstream.
enum class XFinKind { Xirr, Xnpv };

// Row-major flatten of a single range-shaped argument into a vector of
// cells, accepting everything `resolve_range_arg` resolves (`Ref` /
// `RangeOp` / `SpillRef` / `ArrayLiteral` / dynamic-array producers). A
// bare scalar is not a valid XIRR / XNPV range argument and surfaces
// #VALUE!. `out_cells` is cleared on entry; on failure returns false
// with the Excel-visible error written to `*out_err`.
bool collect_range_cells(const parser::AstNode& arg, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx, std::vector<Value>* out_cells, Value* out_err) {
  out_cells->clear();
  auto resolved = resolve_range_arg_no_scalar(arg, arena, registry, ctx, ErrorCode::Value);
  if (!resolved) {
    *out_err = Value::error(resolved.error());
    return false;
  }
  *out_cells = std::move(resolved.value().cells);
  return true;
}

// Collects filtered (value, date) pairs from the two parallel ranges
// with the kind-specific rule set described at the top of this section.
// The resulting pairs are sorted by date ascending so downstream code
// can treat `dates[0]` as the schedule anchor unconditionally.
//
// Returns false with `*out_err` set when the call should surface an
// error; on success `*out_values` and `*out_dates` hold the filtered,
// sorted schedule (possibly empty).
bool collect_xpairs(XFinKind which, const parser::AstNode& values_arg, const parser::AstNode& dates_arg, Arena& arena,
                    const FunctionRegistry& registry, const EvalContext& ctx, std::vector<double>* out_values,
                    std::vector<double>* out_dates, Value* out_err) {
  std::vector<Value> value_cells;
  std::vector<Value> date_cells;
  if (!collect_range_cells(values_arg, arena, registry, ctx, &value_cells, out_err)) {
    return false;
  }
  if (!collect_range_cells(dates_arg, arena, registry, ctx, &date_cells, out_err)) {
    return false;
  }
  // Excel requires the two ranges to have the same total cell count
  // (dimensionality differences that share a count — e.g. 1x4 vs 4x1 —
  // are tolerated: both flatten to a 4-element vector).
  if (value_cells.size() != date_cells.size()) {
    *out_err = Value::error(ErrorCode::Num);
    return false;
  }
  // Collect into a paired vector first so the final sort does not
  // desynchronise values and dates.
  std::vector<std::pair<double, double>> pairs;  // (date, value)
  pairs.reserve(value_cells.size());
  for (std::size_t i = 0; i < value_cells.size(); ++i) {
    const Value& v = value_cells[i];
    const Value& d = date_cells[i];
    if (v.is_error()) {
      *out_err = v;
      return false;
    }
    if (d.is_error()) {
      *out_err = d;
      return false;
    }
    // Blank on the date side always skips the pair — an undated cash
    // flow has no schedule slot. The value-side rule differs between
    // XIRR (skip) and XNPV (#NUM!).
    if (d.is_blank()) {
      continue;
    }
    if (v.is_blank()) {
      if (which == XFinKind::Xnpv) {
        *out_err = Value::error(ErrorCode::Num);
        return false;
      }
      continue;  // XIRR: drop the pair.
    }
    // Non-numeric value cells: XIRR surfaces #VALUE!, XNPV surfaces
    // #NUM! — matches Mac Excel 365 / IronCalc oracle observations.
    if (v.is_boolean() || v.is_text()) {
      *out_err = Value::error(which == XFinKind::Xirr ? ErrorCode::Value : ErrorCode::Num);
      return false;
    }
    auto v_num = coerce_to_number(v);
    if (!v_num) {
      *out_err = Value::error(v_num.error());
      return false;
    }
    // Date cells must coerce cleanly; Bool / Text on the date side is
    // #VALUE! for both functions (`std::pow` has nothing useful to do
    // with a string schedule entry).
    if (d.is_boolean() || d.is_text()) {
      *out_err = Value::error(ErrorCode::Value);
      return false;
    }
    auto d_num = coerce_to_number(d);
    if (!d_num) {
      *out_err = Value::error(d_num.error());
      return false;
    }
    pairs.emplace_back(std::trunc(d_num.value()), v_num.value());
  }
  // Excel rejects pre-1900 / negative serial dates; the oracle treats
  // this as #NUM! for both functions.
  for (const auto& p : pairs) {
    if (p.first < 0.0) {
      *out_err = Value::error(ErrorCode::Num);
      return false;
    }
  }
  // Mac Excel requires the first date in an XNPV schedule to be the
  // earliest and rejects any later entry that precedes it. XIRR accepts
  // earlier dates after the first one.
  if (which == XFinKind::Xnpv && !pairs.empty()) {
    const double anchor = pairs.front().first;
    for (std::size_t i = 1; i < pairs.size(); ++i) {
      if (pairs[i].first < anchor) {
        *out_err = Value::error(ErrorCode::Num);
        return false;
      }
    }
  }
  // XNPV sorts by date ascending (stable, so same-day entries keep their row
  // order). XIRR keeps the listed order: Excel anchors it at the first listed
  // pair, earlier dates discounting with negative exponents.
  if (which == XFinKind::Xnpv) {
    sort_by_index(pairs, [](const std::pair<double, double>& a, const std::pair<double, double>& b) {
      return a.first < b.first;
    });
  }
  out_values->clear();
  out_dates->clear();
  out_values->reserve(pairs.size());
  out_dates->reserve(pairs.size());
  for (const auto& p : pairs) {
    out_dates->push_back(p.first);
    out_values->push_back(p.second);
  }
  return true;
}

// XNPV-style sum over a (value, date) schedule at `rate`. At `rate == -1`
// every later flow divides by zero, so the sum is +-inf, or NaN when later
// flows of both signs meet.
double xnpv_sum(const std::vector<double>& values, const std::vector<double>& dates, double rate) noexcept {
  const double base = 1.0 + rate;
  const double anchor = dates.empty() ? 0.0 : dates[0];
  double total = 0.0;
  for (std::size_t i = 0; i < values.size(); ++i) {
    const double t = (dates[i] - anchor) / 365.0;
    total = total + values[i] / std::exp(std::log(base) * t);
  }
  return total;
}

// The later flows discounted at `x`, summed in listed order. A discount
// factor that is NaN, infinite or zero makes the whole evaluation fail.
bool xirr_rest(const std::vector<double>& values, const std::vector<double>& dates, double x, double* out) noexcept {
  const double base = x + 1.0;
  double total = 0.0;
  for (std::size_t i = 1; i < values.size(); ++i) {
    const double p = std::pow(base, (dates[i] - dates[0]) / 365.0);
    if (!std::isfinite(p) || p == 0.0) {
      return false;
    }
    total = values[i] / p + total;
  }
  *out = total;
  return true;
}

// The side of zero `f` must lie on to be opposite to the first flow.
bool xirr_opposite(double first_value, double f) noexcept {
  return first_value != 0.0 && ((f > 0.0) - (f < 0.0)) == -((first_value > 0.0) - (first_value < 0.0));
}

// XIRR's search from the guess `g` (non-zero, above -1). The guess is negated
// when the first flow outweighs the rest and it is positive. A negative guess
// bisects [0, -1] until both the residual and the bracket are small; a
// positive one doubles until the sign flips, then bisects a bracket. 100
// passes without convergence give NaN.
double xirr_solve(const std::vector<double>& values, const std::vector<double>& dates, double g) noexcept {
  constexpr int kMaxPasses = 100;
  const double first_value = values[0];
  double later_sum = 0.0;
  for (std::size_t i = 1; i < values.size(); ++i) {
    later_sum = values[i] + later_sum;
  }
  if (std::fabs(first_value) > std::fabs(later_sum) && g > 0.0) {
    g = -g;
  }
  double rest = 0.0;
  if (!xirr_rest(values, dates, g, &rest)) {
    return std::numeric_limits<double>::quiet_NaN();
  }
  if (g < 0.0) {
    double lo = 0.0;
    double hi = -1.0;
    for (int pass = 0; pass < kMaxPasses; ++pass) {
      const double f = first_value + rest;
      const double residual_limit = (std::fabs(rest) + std::fabs(first_value)) * 1.0e-8;
      if (std::fabs(f) < residual_limit && std::fabs(hi - lo) < (std::fabs(lo + hi) + 1.0) * 1.0e-8) {
        return g;
      }
      if (xirr_opposite(first_value, f)) {
        hi = g;
      } else {
        lo = g;
      }
      g = (hi - lo) / 2.0 + lo;
      if (!xirr_rest(values, dates, g, &rest)) {
        return std::numeric_limits<double>::quiet_NaN();
      }
    }
    return std::numeric_limits<double>::quiet_NaN();
  }
  bool bracketed = false;
  double lo = 0.0;
  double hi = 0.0;
  for (int pass = 0; pass < kMaxPasses; ++pass) {
    if (bracketed && std::fabs(hi - lo) < (std::fabs(lo + hi) + 1.0) * 1.0e-8) {
      return g;
    }
    const double f = first_value + rest;
    if (xirr_opposite(first_value, f)) {
      if (bracketed) {
        lo = g;
        g = (hi - lo) / 2.0 + lo;
      } else {
        g = g * 2.0;
      }
    } else if (bracketed) {
      hi = g;
      g = (hi - lo) / 2.0 + lo;
    } else {
      bracketed = true;
      hi = g;
      g = g / 2.0;
      lo = 0.0;
    }
    if (!xirr_rest(values, dates, g, &rest)) {
      return std::numeric_limits<double>::quiet_NaN();
    }
  }
  return std::numeric_limits<double>::quiet_NaN();
}

}  // namespace

Value eval_xirr_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 2U || arity > 3U) {
    return Value::error(ErrorCode::Value);
  }

  for (std::uint32_t i = 0; i < 2U; ++i) {
    const Value omitted = atp_arg_error(call.as_call_arg(i), Value::blank(), /*required=*/true);
    if (omitted.is_error()) {
      return omitted;
    }
  }
  std::vector<double> values;
  std::vector<double> dates;
  Value err = Value::blank();
  if (!collect_xpairs(XFinKind::Xirr, call.as_call_arg(0), call.as_call_arg(1), arena, registry, ctx, &values, &dates,
                      &err)) {
    return err;
  }

  double guess = 0.1;  // default guess per Excel.
  if (arity == 3U &&
      !read_optional_guess(call.as_call_arg(2), arena, registry, ctx, /*analysis_toolpak=*/true, &guess, &err)) {
    return err;
  }
  // A guess of -1 or below is outside the rate domain; a single flow has no
  // rate to find, and the flows must include both signs.
  if (!(guess > -1.0)) {
    return Value::error(ErrorCode::Num);
  }
  if (values.size() < 2U) {
    return Value::error(ErrorCode::NA);
  }
  if (!has_opposite_cash_flow_signs(values)) {
    return Value::error(ErrorCode::Num);
  }
  const double rate = xirr_solve(values, dates, guess == 0.0 ? 1.0e-5 : guess);
  if (std::isfinite(rate)) {
    return Value::number(rate);
  }
  return Value::error(ErrorCode::Num);
}

Value eval_xnpv_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity != 3U) {
    return Value::error(ErrorCode::Value);
  }

  // Rate is a scalar numeric argument; evaluate it first so a rate-side
  // error short-circuits before we flatten the parallel ranges.
  const Value rate_v = eval_node(call.as_call_arg(0), arena, registry, ctx);
  if (rate_v.is_error()) {
    return rate_v;
  }
  if (const Value atp = atp_arg_error(call.as_call_arg(0), rate_v, /*required=*/true); atp.is_error()) {
    return atp;
  }
  auto rate = coerce_to_number(rate_v);
  if (!rate) {
    return Value::error(rate.error());
  }
  // Mac Excel 365 rejects rate <= 0 outright. Mathematically rate == 0
  // would just collapse XNPV to sum(values), but Excel treats it as an
  // input error so we mirror that for 1-bit parity.
  if (rate.value() <= 0.0) {
    return Value::error(ErrorCode::Num);
  }

  for (std::uint32_t i = 1; i < 3U; ++i) {
    const Value omitted = atp_arg_error(call.as_call_arg(i), Value::blank(), /*required=*/true);
    if (omitted.is_error()) {
      return omitted;
    }
  }
  std::vector<double> values;
  std::vector<double> dates;
  Value err = Value::blank();
  if (!collect_xpairs(XFinKind::Xnpv, call.as_call_arg(1), call.as_call_arg(2), arena, registry, ctx, &values, &dates,
                      &err)) {
    return err;
  }

  const double result = xnpv_sum(values, dates, rate.value());
  return to_finite_value(result);
}

const FinancialDateEntry* find_financial_date_entry(std::string_view name) noexcept {
  // Strip a leading future-function prefix, mirroring `find_date_entry` --
  // none of these names ship `_xlfn.`-prefixed in practice, but a caller
  // that already stripped or didn't strip a prefix must not matter here.
  for (const std::string_view prefix :
       {std::string_view("_xlfn."), std::string_view("_xlpm."), std::string_view("_xlws.")}) {
    if (name.size() >= prefix.size() && strings::case_insensitive_eq(name.substr(0, prefix.size()), prefix)) {
      name.remove_prefix(prefix.size());
      break;
    }
  }
  struct NamedEntry {
    std::string_view name;
    FinancialDateEntry entry;
  };
  static constexpr NamedEntry kEntries[] = {
      {"COUPPCD", {&financial_detail::CoupPcd, 3u, 4u}},
      {"COUPNCD", {&financial_detail::CoupNcd, 3u, 4u}},
      {"COUPNUM", {&financial_detail::CoupNum, 3u, 4u}},
      {"COUPDAYBS", {&financial_detail::CoupDayBs, 3u, 4u}},
      {"COUPDAYSNC", {&financial_detail::CoupDaysNc, 3u, 4u}},
      {"COUPDAYS", {&financial_detail::CoupDays, 3u, 4u}},
      {"ACCRINT", {&financial_detail::Accrint, 6u, 8u, 1u << 7}},
      {"ACCRINTM", {&financial_detail::Accrintm, 4u, 5u}},
      {"DISC", {&financial_detail::Disc, 4u, 5u}},
      {"INTRATE", {&financial_detail::Intrate, 4u, 5u}},
      {"RECEIVED", {&financial_detail::Received, 4u, 5u}},
      {"TBILLPRICE", {&financial_detail::TBillPrice, 3u, 3u}},
      {"TBILLYIELD", {&financial_detail::TBillYield, 3u, 3u}},
      {"TBILLEQ", {&financial_detail::TBillEq, 3u, 3u}},
      {"PRICEDISC", {&financial_detail::PriceDisc, 4u, 5u}},
      {"PRICEMAT", {&financial_detail::PriceMat, 5u, 6u}},
      {"YIELDDISC", {&financial_detail::YieldDisc, 4u, 5u}},
      {"YIELDMAT", {&financial_detail::YieldMat, 5u, 6u}},
      {"DURATION", {&financial_detail::Duration, 5u, 6u}},
      {"MDURATION", {&financial_detail::MDuration, 5u, 6u}},
      {"PRICE", {&financial_detail::Price, 6u, 7u}},
      {"YIELD", {&financial_detail::Yield, 6u, 7u}},
      {"ODDLPRICE", {&financial_detail::OddlPrice, 7u, 8u}},
      {"ODDLYIELD", {&financial_detail::OddlYield, 7u, 8u}},
      {"ODDFPRICE", {&financial_detail::OddfPrice, 8u, 9u}},
      {"ODDFYIELD", {&financial_detail::OddfYield, 8u, 9u}},
      {"AMORDEGRC", {&financial_detail::Amordegrc, 6u, 7u}},
      {"AMORLINC", {&financial_detail::Amorlinc, 6u, 7u}},
  };
  for (const auto& e : kEntries) {
    if (strings::case_insensitive_eq(e.name, name)) {
      return &e.entry;
    }
  }
  return nullptr;
}

Value eval_financial_date_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                               const EvalContext& ctx) {
  const FinancialDateEntry* entry = find_financial_date_entry(call.as_call_name());
  if (entry == nullptr) {
    // Registered names always resolve; a miss can only mean a table drift.
    return Value::error(ErrorCode::Name);
  }
  const std::uint32_t arity = call.as_call_arity();
  if (arity < entry->min_arity || arity > entry->max_arity) {
    return Value::error(ErrorCode::Value);
  }
  // This family is scalar-only, follows the Analysis-ToolPak argument
  // rule, and propagates the left-most argument error (none opt out of
  // that rule), matching the eager dispatcher's pre-evaluation contract.
  const std::uint32_t evaluated = atp_evaluated_arity(call, entry->min_arity, entry->max_arity);
  std::vector<Value> args;
  args.reserve(evaluated);
  for (std::uint32_t i = 0; i < evaluated; ++i) {
    Value v = eval_node(call.as_call_arg(i), arena, registry, ctx);
    if (v.is_error()) {
      return v;
    }
    const bool logical = ((entry->logical_args >> i) & 1U) != 0U;
    const Value atp =
        atp_arg_error(call.as_call_arg(i), v, atp_slot_required(i, entry->min_arity, entry->max_arity), logical);
    if (atp.is_error()) {
      return atp;
    }
    args.push_back(v);
  }
  return entry->impl(args.empty() ? nullptr : args.data(), evaluated, arena, ctx.date1904());
}

}  // namespace eval
}  // namespace formulon
