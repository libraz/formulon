// Copyright 2026 libraz. Licensed under the MIT License.
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
#include <utility>
#include <vector>

#include "eval/coerce.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "eval/range_args.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Collects the numeric cash flows from IRR's first argument. Accepts:
//
//   - RangeOp / Ref   -> resolved via resolve_range_arg; Text / Bool /
//                         Blank cells are silently skipped (Excel rule).
//   - ArrayLiteral    -> walked cell by cell; each element is evaluated
//                         and coerced to a number. Non-numeric elements
//                         yield #VALUE! (literals are treated as direct
//                         scalars, not filtered like range cells).
//   - anything else   -> rejected with #VALUE!, mirroring Excel's
//                         requirement that IRR's first argument be a
//                         reference or array.
//
// Returns `true` on success with the cash flows written to `*out_flows`.
// On failure writes an Excel error into `*out_err` and returns `false`.
// An error cell encountered during resolution propagates as the IRR
// call's result.
bool collect_cash_flows(const parser::AstNode& arg, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx, std::vector<double>* out_flows, Value* out_err) {
  out_flows->clear();
  const parser::NodeKind k = arg.kind();
  if (k == parser::NodeKind::Ref || k == parser::NodeKind::RangeOp) {
    std::vector<Value> cells;
    ErrorCode range_err = ErrorCode::Value;
    if (!resolve_range_arg(arg, arena, registry, ctx, &cells, &range_err)) {
      *out_err = Value::error(range_err);
      return false;
    }
    for (const Value& v : cells) {
      if (v.is_error()) {
        *out_err = v;
        return false;
      }
      // Excel IRR silently skips non-numeric cells inside a range
      // (matches SUM / AVERAGE behaviour on range-sourced inputs).
      if (!v.is_number()) {
        continue;
      }
      out_flows->push_back(v.as_number());
    }
    return true;
  }
  if (k == parser::NodeKind::ArrayLiteral) {
    const std::uint32_t rows = arg.as_array_rows();
    const std::uint32_t cols = arg.as_array_cols();
    for (std::uint32_t r = 0; r < rows; ++r) {
      for (std::uint32_t c = 0; c < cols; ++c) {
        const Value v = eval_node(arg.as_array_element(r, c), arena, registry, ctx);
        if (v.is_error()) {
          *out_err = v;
          return false;
        }
        auto coerced = coerce_to_number(v);
        if (!coerced) {
          *out_err = Value::error(coerced.error());
          return false;
        }
        out_flows->push_back(coerced.value());
      }
    }
    return true;
  }
  // Any other shape (a scalar literal, a function call, etc.) is not a
  // valid IRR values argument.
  *out_err = Value::error(ErrorCode::Value);
  return false;
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

// Closed-form MIRR helper. `flows` must already have been validated to
// contain at least one positive and at least one negative value. Returns
// NaN on a degenerate ratio (e.g. all-positive or all-negative schedule
// bypasses this function via the caller's pre-check; this guard is
// defensive). `n` is `flows.size()` cached by the caller.
double mirr_closed_form(const std::vector<double>& flows, double finance_rate, double reinvest_rate) noexcept {
  const std::size_t n = flows.size();
  if (n < 2) {
    return std::numeric_limits<double>::quiet_NaN();
  }
  // Walk once: period index runs 0..n-1. Positive flows discount at
  // reinvest_rate; negative flows discount at finance_rate. We sum each
  // into its own accumulator so the final ratio is a single division.
  double npv_pos = 0.0;
  double npv_neg = 0.0;
  double pos_factor = 1.0;  // (1 + reinvest_rate)^0
  double neg_factor = 1.0;  // (1 + finance_rate)^0
  const double pos_base = 1.0 + reinvest_rate;
  const double neg_base = 1.0 + finance_rate;
  for (std::size_t i = 0; i < n; ++i) {
    if (flows[i] > 0.0) {
      npv_pos += flows[i] / pos_factor;
    } else if (flows[i] < 0.0) {
      npv_neg += flows[i] / neg_factor;
    }
    pos_factor *= pos_base;
    neg_factor *= neg_base;
  }
  if (npv_neg == 0.0) {
    return std::numeric_limits<double>::quiet_NaN();
  }
  // Canonical Excel form: (-npv_pos * (1+reinvest)^(n-1) / npv_neg)^(1/(n-1)) - 1.
  // We already have (1+reinvest)^n cached in `pos_factor` at loop exit,
  // so divide once by pos_base to get (1+reinvest)^(n-1).
  const double pos_n_minus_1 = pos_factor / pos_base;
  const double ratio = -npv_pos * pos_n_minus_1 / npv_neg;
  // `ratio < 0.0` keeps the usual NaN guard (a real-valued `(n-1)`th root
  // of a negative number is undefined in the reals). `ratio == 0.0` is
  // valid and arises when npv_pos == 0 (all-negative schedule): IEEE-754
  // `pow(0, 1/(n-1)) = 0` for positive exponents, so the result is -1.0,
  // which matches Mac Excel 365.
  if (ratio < 0.0) {
    return std::numeric_limits<double>::quiet_NaN();
  }
  return std::pow(ratio, 1.0 / static_cast<double>(n - 1)) - 1.0;
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
  bool has_positive = false;
  bool has_negative = false;
  for (double v : flows) {
    if (v > 0.0) {
      has_positive = true;
    } else if (v < 0.0) {
      has_negative = true;
    }
  }
  if (!has_positive || !has_negative) {
    return Value::error(ErrorCode::Num);
  }

  double rate = 0.1;  // default guess per Excel.
  if (arity == 2U) {
    const Value guess_v = eval_node(call.as_call_arg(1), arena, registry, ctx);
    if (guess_v.is_error()) {
      return guess_v;
    }
    // Blank guess falls back to 0.1, matching Excel's behaviour when the
    // second argument is elided with a trailing comma. `coerce_to_number`
    // returns 0 for blank, which is a valid (but different) starting
    // guess — we therefore only apply the coerced value when the cell is
    // not blank, so `=IRR(A1:A4,)` and `=IRR(A1:A4)` behave identically.
    if (!guess_v.is_blank()) {
      auto coerced = coerce_to_number(guess_v);
      if (!coerced) {
        return Value::error(coerced.error());
      }
      rate = coerced.value();
    }
  }

  // Newton-Raphson with a fixed iteration cap. The Excel specification
  // uses 20 iterations; we use 100 to match common open-source
  // implementations and give the method a wider convergence basin. The
  // tolerance 1e-10 on |delta rate| matches the oracle's expectations.
  constexpr int kMaxIter = 100;
  constexpr double kTolerance = 1.0e-10;
  for (int iter = 0; iter < kMaxIter; ++iter) {
    if (rate <= -1.0) {
      // (1 + rate) must stay positive for the NPV expansion to be
      // meaningful; Excel returns #NUM! when the iterate wanders past
      // that boundary.
      return Value::error(ErrorCode::Num);
    }
    const double npv = irr_npv(flows, rate);
    const double dnpv = irr_dnpv(flows, rate);
    if (dnpv == 0.0 || std::isnan(dnpv) || std::isinf(dnpv)) {
      return Value::error(ErrorCode::Num);
    }
    const double new_rate = rate - npv / dnpv;
    if (std::isnan(new_rate) || std::isinf(new_rate)) {
      return Value::error(ErrorCode::Num);
    }
    if (std::fabs(new_rate - rate) < kTolerance) {
      return Value::number(new_rate);
    }
    rate = new_rate;
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
  bool has_negative = false;
  for (double v : flows) {
    if (v < 0.0) {
      has_negative = true;
      break;
    }
  }
  if (!has_negative) {
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
  if (std::isnan(result) || std::isinf(result)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(result);
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

// Row-major flatten of a single range/Ref/ArrayLiteral argument into a
// vector of cells. `out_cells` is cleared on entry; on failure returns
// false with the Excel-visible error written to `*out_err`.
bool collect_range_cells(const parser::AstNode& arg, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx, std::vector<Value>* out_cells, Value* out_err) {
  out_cells->clear();
  const parser::NodeKind k = arg.kind();
  if (k == parser::NodeKind::Ref || k == parser::NodeKind::RangeOp) {
    ErrorCode range_err = ErrorCode::Value;
    if (!resolve_range_arg(arg, arena, registry, ctx, out_cells, &range_err)) {
      *out_err = Value::error(range_err);
      return false;
    }
    return true;
  }
  if (k == parser::NodeKind::ArrayLiteral) {
    const std::uint32_t rows = arg.as_array_rows();
    const std::uint32_t cols = arg.as_array_cols();
    out_cells->reserve(static_cast<std::size_t>(rows) * cols);
    for (std::uint32_t r = 0; r < rows; ++r) {
      for (std::uint32_t c = 0; c < cols; ++c) {
        out_cells->push_back(eval_node(arg.as_array_element(r, c), arena, registry, ctx));
      }
    }
    return true;
  }
  // Any other shape (scalar literal, arithmetic expression, etc.) is
  // not a valid XIRR / XNPV range argument.
  *out_err = Value::error(ErrorCode::Value);
  return false;
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
  // Sort by date ascending. Stable-sort keeps the original row order
  // for same-day entries, which matches Excel's behaviour when two
  // cash flows share a serial.
  std::stable_sort(
      pairs.begin(), pairs.end(),
      [](const std::pair<double, double>& a, const std::pair<double, double>& b) { return a.first < b.first; });
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

// XNPV-style sum over a (value, date) schedule at `rate`. Caller must
// have ensured `rate > -1`.
double xnpv_sum(const std::vector<double>& values, const std::vector<double>& dates, double rate) noexcept {
  const double base = 1.0 + rate;
  const double anchor = dates.empty() ? 0.0 : dates[0];
  double total = 0.0;
  for (std::size_t i = 0; i < values.size(); ++i) {
    const double t = (dates[i] - anchor) / 365.0;
    total += values[i] / std::pow(base, t);
  }
  return total;
}

// d/dr of xnpv_sum: f'(r) = -sum_i values[i] * t_i / (1+r)^(t_i + 1).
double xnpv_dsum(const std::vector<double>& values, const std::vector<double>& dates, double rate) noexcept {
  const double base = 1.0 + rate;
  const double anchor = dates.empty() ? 0.0 : dates[0];
  double total = 0.0;
  for (std::size_t i = 0; i < values.size(); ++i) {
    const double t = (dates[i] - anchor) / 365.0;
    total += -values[i] * t / std::pow(base, t + 1.0);
  }
  return total;
}

// Runs Newton-Raphson on xnpv_sum starting from `guess`. Returns NaN on
// failure (iterate leaves the domain rate > -1, derivative collapses,
// or the iteration cap is exhausted without converging); otherwise the
// converged rate.
double xirr_newton(const std::vector<double>& values, const std::vector<double>& dates, double guess) noexcept {
  constexpr int kMaxIter = 100;
  constexpr double kTolerance = 1.0e-10;
  double rate = guess;
  for (int iter = 0; iter < kMaxIter; ++iter) {
    if (rate <= -1.0) {
      return std::numeric_limits<double>::quiet_NaN();
    }
    const double fval = xnpv_sum(values, dates, rate);
    if (std::isnan(fval) || std::isinf(fval)) {
      return std::numeric_limits<double>::quiet_NaN();
    }
    if (std::fabs(fval) < kTolerance) {
      return rate;
    }
    const double dval = xnpv_dsum(values, dates, rate);
    if (dval == 0.0 || std::isnan(dval) || std::isinf(dval)) {
      return std::numeric_limits<double>::quiet_NaN();
    }
    const double new_rate = rate - fval / dval;
    if (std::isnan(new_rate) || std::isinf(new_rate)) {
      return std::numeric_limits<double>::quiet_NaN();
    }
    if (std::fabs(new_rate - rate) < kTolerance) {
      return new_rate;
    }
    rate = new_rate;
  }
  return std::numeric_limits<double>::quiet_NaN();
}

// Bracketed fallback for XIRR when Newton diverges or can't cross the
// r=-1 singularity. Samples f(rate) on a log-linear grid spanning the
// domain rate > -1, looks for an adjacent sign change, and bisects.
// The grid pattern is:
//
//   grid_near_minus_1 = { -1 + 10^k for k in [-10, -1] }   // [~-1, -0.9]
//   grid_mid           = { -0.8, -0.5, -0.2, 0.0, 0.1, 0.2, 0.5, 1.0 }
//   grid_large         = { 10^k for k in [1, 6] }          // [10, 1e6]
//
// Combined (sorted ascending), this produces ~30 samples which are
// enough to bracket the XIRR root on every fixture in the oracle
// corpus. Bisection then drives the bracket width below 1e-12 or
// |f(mid)| below 1e-10, whichever comes first.
double xirr_bracket(const std::vector<double>& values, const std::vector<double>& dates) noexcept {
  // Build the sample grid once; numbers chosen so neighbouring samples
  // are close enough to catch even the R17-style root at r ~= -0.912.
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
  double prev_f = xnpv_sum(values, dates, prev_r);
  if (std::isfinite(prev_f) && std::fabs(prev_f) < 1.0e-10) {
    return prev_r;
  }
  for (std::size_t i = 1; i < grid.size(); ++i) {
    const double r = grid[i];
    const double f = xnpv_sum(values, dates, r);
    if (!std::isfinite(f)) {
      prev_r = r;
      prev_f = f;
      continue;
    }
    if (std::fabs(f) < 1.0e-10) {
      return r;
    }
    if (std::isfinite(prev_f) && ((prev_f < 0.0 && f > 0.0) || (prev_f > 0.0 && f < 0.0))) {
      // Bisect on [prev_r, r].
      double lo = prev_r;
      double hi = r;
      double f_lo = prev_f;
      for (int iter = 0; iter < 200; ++iter) {
        const double mid = 0.5 * (lo + hi);
        const double f_mid = xnpv_sum(values, dates, mid);
        if (!std::isfinite(f_mid)) {
          // Numerical blow-up mid-bracket; narrow from the other end.
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
      // Refine the final bracket midpoint with Newton; usually just
      // polishes the converged root to full double precision.
      const double refined = xirr_newton(values, dates, 0.5 * (lo + hi));
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

}  // namespace

Value eval_xirr_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 2U || arity > 3U) {
    return Value::error(ErrorCode::Value);
  }

  std::vector<double> values;
  std::vector<double> dates;
  Value err = Value::blank();
  if (!collect_xpairs(XFinKind::Xirr, call.as_call_arg(0), call.as_call_arg(1), arena, registry, ctx, &values, &dates,
                      &err)) {
    return err;
  }

  // XIRR needs at least two cash flows with opposite sign to root-find
  // meaningfully; otherwise Newton either diverges or converges to the
  // boundary `rate == -1`. Mac Excel 365 reports #NUM! here.
  if (values.size() < 2U) {
    return Value::error(ErrorCode::Num);
  }
  bool has_positive = false;
  bool has_negative = false;
  for (double v : values) {
    if (v > 0.0) {
      has_positive = true;
    } else if (v < 0.0) {
      has_negative = true;
    }
  }
  if (!has_positive || !has_negative) {
    return Value::error(ErrorCode::Num);
  }

  double guess = 0.1;  // default guess per Excel.
  if (arity == 3U) {
    const Value guess_v = eval_node(call.as_call_arg(2), arena, registry, ctx);
    if (guess_v.is_error()) {
      return guess_v;
    }
    // Blank guess falls back to 0.1 so `=XIRR(v, d,)` matches `=XIRR(v, d)`.
    if (!guess_v.is_blank()) {
      auto coerced = coerce_to_number(guess_v);
      if (!coerced) {
        return Value::error(coerced.error());
      }
      guess = coerced.value();
    }
  }
  if (guess <= -1.0) {
    return Value::error(ErrorCode::Num);
  }

  // Try Newton-Raphson first; it converges quickly for the well-posed
  // schedules (positive roots, guess near the answer). When it fails —
  // typically because the true root is near the rate == -1 singularity
  // and Newton can't cross it — fall back to a bracket + bisection
  // scan over the full rate > -1 domain. LibreOffice and gnumeric use
  // the same hybrid strategy.
  const double newton_rate = xirr_newton(values, dates, guess);
  if (std::isfinite(newton_rate)) {
    return Value::number(newton_rate);
  }
  const double bracket_rate = xirr_bracket(values, dates);
  if (std::isfinite(bracket_rate)) {
    return Value::number(bracket_rate);
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
  auto rate = coerce_to_number(rate_v);
  if (!rate) {
    return Value::error(rate.error());
  }
  // Excel / IronCalc reject any negative rate for XNPV. rate == 0 is
  // still allowed — the discount factors collapse to 1 and the sum
  // degenerates to the undiscounted total, which is well-defined.
  if (rate.value() < 0.0) {
    return Value::error(ErrorCode::Num);
  }

  std::vector<double> values;
  std::vector<double> dates;
  Value err = Value::blank();
  if (!collect_xpairs(XFinKind::Xnpv, call.as_call_arg(1), call.as_call_arg(2), arena, registry, ctx, &values, &dates,
                      &err)) {
    return err;
  }

  const double result = xnpv_sum(values, dates, rate.value());
  if (std::isnan(result) || std::isinf(result)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(result);
}

}  // namespace eval
}  // namespace formulon
