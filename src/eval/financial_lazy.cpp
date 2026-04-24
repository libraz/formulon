// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the IRR / MIRR / XIRR / XNPV lazy impls. See
// `eval/financial_lazy.h` for the dispatch-table contract and
// `eval/lazy_impls.h` for the shared `eval_node` / `LazyImpl` vocabulary.

#include "eval/financial_lazy.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>
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
//   2. A per-pair filter that is more permissive than IRR's: Excel
//      silently SKIPS any pair where `values[i]` is blank (not coerced
//      to zero), matching the way SUM / NPV treat blank cells in a
//      range. A pair where `dates[i]` is blank is also skipped because
//      an undated cash flow has no place in the schedule.
//   3. A date-to-offset transform: `t_i = (dates[i] - dates[0]) / 365`,
//      with `dates[i]` truncated toward zero (`std::trunc`) to match
//      Excel's serial-date conversion rule.
//
// The two impls share the collection pipeline so a single bug in the
// filter cannot drift XIRR away from XNPV.

namespace {

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

// Collects filtered (value, date) pairs from the two parallel ranges.
// Rules:
//   * Error cells on either side propagate as the call's result.
//   * Pairs where either side is Blank are silently skipped.
//   * Pairs where either side fails numeric coercion surface #VALUE!.
//   * Dates are truncated toward zero (Excel serial rule).
// Writes `*out_err` and returns false on any failure that should
// propagate out as the final result.
bool collect_xpairs(const parser::AstNode& values_arg, const parser::AstNode& dates_arg, Arena& arena,
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
  out_values->clear();
  out_dates->clear();
  out_values->reserve(value_cells.size());
  out_dates->reserve(date_cells.size());
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
    // A blank cell on either side disqualifies the pair; matches
    // Excel's treatment of blank cells in cash-flow ranges.
    if (v.is_blank() || d.is_blank()) {
      continue;
    }
    auto v_num = coerce_to_number(v);
    if (!v_num) {
      *out_err = Value::error(v_num.error());
      return false;
    }
    auto d_num = coerce_to_number(d);
    if (!d_num) {
      *out_err = Value::error(d_num.error());
      return false;
    }
    out_values->push_back(v_num.value());
    out_dates->push_back(std::trunc(d_num.value()));
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
  if (!collect_xpairs(call.as_call_arg(0), call.as_call_arg(1), arena, registry, ctx, &values, &dates, &err)) {
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

  double rate = 0.1;  // default guess per Excel.
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
      rate = coerced.value();
    }
  }
  if (rate <= -1.0) {
    return Value::error(ErrorCode::Num);
  }

  // Newton-Raphson, same iteration cap and tolerance as IRR so oracle
  // expectations on the two impls stay aligned.
  constexpr int kMaxIter = 100;
  constexpr double kTolerance = 1.0e-10;
  for (int iter = 0; iter < kMaxIter; ++iter) {
    if (rate <= -1.0) {
      return Value::error(ErrorCode::Num);
    }
    const double fval = xnpv_sum(values, dates, rate);
    if (std::isnan(fval) || std::isinf(fval)) {
      return Value::error(ErrorCode::Num);
    }
    if (std::fabs(fval) < kTolerance) {
      return Value::number(rate);
    }
    const double dval = xnpv_dsum(values, dates, rate);
    if (dval == 0.0 || std::isnan(dval) || std::isinf(dval)) {
      return Value::error(ErrorCode::Num);
    }
    const double new_rate = rate - fval / dval;
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
  if (rate.value() <= -1.0) {
    return Value::error(ErrorCode::Num);
  }

  std::vector<double> values;
  std::vector<double> dates;
  Value err = Value::blank();
  if (!collect_xpairs(call.as_call_arg(1), call.as_call_arg(2), arena, registry, ctx, &values, &dates, &err)) {
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
