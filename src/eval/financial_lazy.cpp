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
#include <optional>
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
#include "eval/checked_arith.h"
#include "eval/coerce.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "eval/omitted_arg.h"
#include "eval/range_args.h"
#include "eval/rate_secant.h"
#include "eval/scalar_args.h"
#include "eval/tree_walker/dispatch.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/elementary_math.h"
#include "utils/error.h"
#include "utils/index_sort.h"
#include "utils/strings.h"
#include "utils/subnormal_flush.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

using builtins_detail::to_finite_value;

// Collects the numeric cash flows from IRR's first argument. Every shape
// `resolve_range_arg` understands is accepted — `Ref` / `RangeOp` /
// `SpillRef` / `ArrayLiteral` and dynamic-array producers such as
// `SEQUENCE`. A lone value is #VALUE!, answered by the lazy table's
// lone-value policy before the impl runs.
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
// An error cell encountered during resolution is written to `*out_err`;
// `*out_cell_error` (when given) tells whether it came from a reference.
bool collect_cash_flows(const parser::AstNode& arg, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx, std::vector<double>* out_flows, Value* out_err,
                        bool* out_cell_error = nullptr) {
  out_flows->clear();
  auto resolved = resolve_range_arg(arg, arena, registry, ctx);
  if (!resolved) {
    *out_err = Value::error(resolved.error());
    return false;
  }
  const bool literal = arg.kind() == parser::NodeKind::ArrayLiteral;
  for (const Value& v : resolved.value().cells) {
    if (v.is_error()) {
      *out_err = v;
      if (out_cell_error != nullptr) {
        *out_cell_error = !literal && !resolved.value().from_scalar;
      }
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

// The residual IRR's secant search drives to zero, as a function of
// u = g / (-1 - g): `F = F * B; S = S + v / F` with `B = 1 / (1 + u)`, the
// first flow undiscounted and no factor update after the last flow. The
// arithmetic is checked (see CheckedArith).
class IrrResidual {
 public:
  static constexpr int kFaultDivide = CheckedArith::kFaultDivide;

  explicit IrrResidual(const std::vector<double>& flows) noexcept : flows_(flows) {}

  double operator()(double u) noexcept {
    const double b = math_.div(1.0, math_.add(1.0, u));
    double sum = 0.0;
    double factor = 1.0;
    bool first = true;
    for (const double v : flows_) {
      if (!first) {
        factor = math_.mul(factor, b);
      }
      first = false;
      sum = math_.add(sum, math_.div(v, factor));
    }
    return sum;
  }

  /// Last fault code left by the evaluations so far, or 0.
  int fault() const noexcept { return math_.fault(); }

  bool failed() const noexcept { return math_.failed(); }

 private:
  const std::vector<double>& flows_;
  CheckedArith math_;
};

// IRR from `guess`: the secant search in u, restarted from 0.1, -1/11 and 2
// when it fails (the first restart is skipped when the guess is omitted).
// A start whose 1 + u is not a positive normal number is #VALUE!; running out
// of restarts is #NUM!.
Value irr_solve(const std::vector<double>& flows, double guess, bool guess_given) noexcept {
  constexpr double kRestarts[] = {0.1, -0.09090909090909091, 2.0};
  std::size_t restart = guess_given ? 0U : 1U;
  if (-1.0 - guess == 0.0) {
    return Value::error(ErrorCode::Value);
  }
  double u = guess / (-1.0 - guess);
  IrrResidual residual(flows);
  while (true) {
    const double base = 1.0 + u;
    if (std::isnan(base) || std::isinf(base) || !(base > rate_secant::kTiny)) {
      return Value::error(ErrorCode::Value);
    }
    const rate_secant::Result found = rate_secant::search(residual, u);
    if (residual.fault() != 0) {
      return Value::error(residual.fault() == IrrResidual::kFaultDivide ? ErrorCode::Div0 : ErrorCode::Num);
    }
    if (found.converged && (found.x - found.tolerance) > -1.0) {
      return Value::number(1.0 / (found.x + 1.0) - 1.0);
    }
    if (restart > 2U) {
      return Value::error(ErrorCode::Num);
    }
    u = kRestarts[restart] / (-1.0 - kRestarts[restart]);
    ++restart;
  }
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

// Reads one guess or rate value into `*out`: a blank cell is 0. `arg` is the
// argument it came from; with `analysis_toolpak` (XIRR, XNPV) a boolean is
// #VALUE!. Returns false with the error in `*out_err` on an error or a
// non-numeric value.
bool read_rate_value(const parser::AstNode& arg, const Value& v, bool analysis_toolpak, double* out, Value* out_err) {
  if (v.is_error()) {
    *out_err = v;
    return false;
  }
  if (analysis_toolpak) {
    if (const Value atp = atp_arg_error(arg, v, /*required=*/false); atp.is_error()) {
      *out_err = atp;
      return false;
    }
  }
  auto coerced = coerce_to_number(v);
  if (!coerced) {
    *out_err = Value::error(coerced.error());
    return false;
  }
  *out = coerced.value();
  return true;
}

// Evaluates `rate_args` scalar rate arguments of `call` starting at `first` and
// runs `cell` on their numbers, element by element when one is an array. An
// omitted slot reads `defaults[i]`.
template <std::uint32_t kCount, typename Cell>
Value lift_rates(const parser::AstNode& call, std::uint32_t first, const double (&defaults)[kCount],
                 bool analysis_toolpak, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                 const Cell& cell) {
  std::vector<Value> args(kCount, Value::blank());
  for (std::uint32_t i = 0; i < kCount; ++i) {
    args[i] = call_arg_supplied(call, first + i) ? eval_node(call.as_call_arg(first + i), arena, registry, ctx)
                                                 : Value::number(defaults[i]);
  }
  return lift_over_array_args(args.data(), kCount, arena, [&](const Value* vals, std::uint32_t) {
    double rates[kCount];
    for (std::uint32_t i = 0; i < kCount; ++i) {
      Value err = Value::blank();
      const parser::AstNode& arg = call.as_call_arity() > first + i ? call.as_call_arg(first + i) : call;
      if (!read_rate_value(arg, vals[i], analysis_toolpak, &rates[i], &err)) {
        return err;
      }
    }
    return cell(rates);
  });
}

}  // namespace

Value eval_irr_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                    const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 1U || arity > 2U) {
    return Value::error(ErrorCode::Value);
  }

  // An error cell read through a reference is #VALUE!, whatever the error.
  std::vector<double> flows;
  Value collect_err = Value::blank();
  bool cell_error = false;
  if (!collect_cash_flows(call.as_call_arg(0), arena, registry, ctx, &flows, &collect_err, &cell_error)) {
    return cell_error ? Value::error(ErrorCode::Value) : collect_err;
  }

  // IRR requires at least one positive and one negative cash flow so the
  // NPV(r) function has a root. Otherwise the iteration would either
  // diverge or converge to a degenerate boundary (rate == -1).
  if (!has_opposite_cash_flow_signs(flows)) {
    return Value::error(ErrorCode::Num);
  }

  // An omitted guess starts from 0.1 and skips the first restart.
  if (!call_arg_supplied(call, 1U)) {
    return irr_solve(flows, 0.1, /*guess_given=*/false);
  }
  constexpr double kNoDefault[1] = {0.0};
  return lift_rates(call, 1U, kNoDefault, /*analysis_toolpak=*/false, arena, registry, ctx,
                    [&](const double* rates) { return irr_solve(flows, rates[0], /*guess_given=*/true); });
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

  // The finance and reinvestment rates lift element by element.
  constexpr double kNoDefaults[2] = {0.0, 0.0};
  return lift_rates(call, 1U, kNoDefaults, /*analysis_toolpak=*/false, arena, registry, ctx,
                    [&](const double* rates) { return to_finite_value(mirr_closed_form(flows, rates[0], rates[1])); });
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
//      The two functions differ as follows:
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
  auto values = resolve_range_arg(values_arg, arena, registry, ctx);
  if (!values) {
    *out_err = Value::error(values.error());
    return false;
  }
  auto dates = resolve_range_arg(dates_arg, arena, registry, ctx);
  if (!dates) {
    *out_err = Value::error(dates.error());
    return false;
  }
  const RangeResult value_range = std::move(values.value());
  const RangeResult date_range = std::move(dates.value());
  const std::vector<Value>& value_cells = value_range.cells;
  const std::vector<Value>& date_cells = date_range.cells;
  // A lone value (a scalar or a one-cell reference) answers its error before
  // the count check; an array's error element, computed or read from a range,
  // only after it.
  const auto lone_error = [&](const parser::AstNode& arg, const RangeResult& resolved) -> const Value* {
    if (resolved.cells.size() == 1U && resolved.cells.front().is_error() &&
        (resolved.from_scalar || resolve_binding_reference(arg, arena, registry, ctx) != nullptr)) {
      return &resolved.cells.front();
    }
    return nullptr;
  };
  for (const Value* error : {lone_error(values_arg, value_range), lone_error(dates_arg, date_range)}) {
    if (error != nullptr) {
      *out_err = *error;
      return false;
    }
  }
  // The two must hold the same number of cells (1x4 against 4x1 is fine).
  if (value_cells.size() != date_cells.size()) {
    *out_err = Value::error(ErrorCode::Num);
    return false;
  }
  // The earliest error position answers; at one position XIRR takes the
  // date's error and XNPV the value's.
  const std::vector<Value>* first = which == XFinKind::Xirr ? &date_cells : &value_cells;
  const std::vector<Value>* second = which == XFinKind::Xirr ? &value_cells : &date_cells;
  for (std::size_t i = 0; i < value_cells.size(); ++i) {
    for (const std::vector<Value>* cells : {first, second}) {
      if ((*cells)[i].is_error()) {
        *out_err = (*cells)[i];
        return false;
      }
    }
  }
  // Collect into a paired vector first so the final sort does not
  // desynchronise values and dates.
  std::vector<std::pair<double, double>> pairs;  // (date, value)
  pairs.reserve(value_cells.size());
  for (std::size_t i = 0; i < value_cells.size(); ++i) {
    const Value& v = value_cells[i];
    const Value& d = date_cells[i];
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
    total = total + values[i] / elementary::exp(elementary::log(base) * t);
  }
  return total;
}

// The later flows discounted at `x`, summed in listed order. A discount
// factor that is NaN, infinite or zero makes the whole evaluation fail.
bool xirr_rest(const std::vector<double>& values, const std::vector<double>& dates, double x, double* out) noexcept {
  const double base = x + 1.0;
  double total = 0.0;
  for (std::size_t i = 1; i < values.size(); ++i) {
    const double p = elementary::pow(base, (dates[i] - dates[0]) / 365.0);
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

// Returns the first required-argument error among args [begin, end), else blank.
Value first_omitted_arg_error(const parser::AstNode& call, std::uint32_t begin, std::uint32_t end) {
  for (std::uint32_t i = begin; i < end; ++i) {
    const Value omitted = atp_arg_error(call.as_call_arg(i), Value::blank(), /*required=*/true);
    if (omitted.is_error()) {
      return omitted;
    }
  }
  return Value::blank();
}

}  // namespace

Value eval_xirr_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 2U || arity > 3U) {
    return Value::error(ErrorCode::Value);
  }

  if (const Value omitted = first_omitted_arg_error(call, 0U, 2U); omitted.is_error()) {
    return omitted;
  }
  std::vector<double> values;
  std::vector<double> dates;
  Value err = Value::blank();
  if (!collect_xpairs(XFinKind::Xirr, call.as_call_arg(0), call.as_call_arg(1), arena, registry, ctx, &values, &dates,
                      &err)) {
    return err;
  }

  // The guess defaults to 0.1 and lifts element by element.
  constexpr double kDefaultGuess[1] = {0.1};
  return lift_rates(call, 2U, kDefaultGuess, /*analysis_toolpak=*/true, arena, registry, ctx, [&](const double* g) {
    const double guess = g[0];
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
    return std::isfinite(rate) ? Value::number(rate) : Value::error(ErrorCode::Num);
  });
}

Value eval_xnpv_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity != 3U) {
    return Value::error(ErrorCode::Value);
  }

  // The rate is read first, so a rate error answers before the schedule's;
  // an array rate lifts element by element.
  const Value rate_v = eval_node(call.as_call_arg(0), arena, registry, ctx);
  if (rate_v.is_error()) {
    return rate_v;
  }
  if (const Value atp = atp_arg_error(call.as_call_arg(0), rate_v, /*required=*/true); atp.is_error()) {
    return atp;
  }
  return lift_over_array_args(&rate_v, 1U, arena, [&](const Value* args, std::uint32_t) {
    double rate = 0.0;
    Value err = Value::blank();
    if (!read_rate_value(call.as_call_arg(0), args[0], /*analysis_toolpak=*/true, &rate, &err)) {
      return err;
    }
    // Mac Excel 365 rejects rate <= 0 outright, though rate 0 would only
    // collapse XNPV to sum(values).
    if (rate <= 0.0) {
      return Value::error(ErrorCode::Num);
    }
    if (const Value omitted = first_omitted_arg_error(call, 1U, 3U); omitted.is_error()) {
      return omitted;
    }
    std::vector<double> values;
    std::vector<double> dates;
    if (!collect_xpairs(XFinKind::Xnpv, call.as_call_arg(1), call.as_call_arg(2), arena, registry, ctx, &values, &dates,
                        &err)) {
      return err;
    }
    return to_finite_value(xnpv_sum(values, dates, rate));
  });
}

const FinancialDateEntry* find_financial_date_entry(std::string_view name) noexcept {
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
  const FinancialDateEntry* entry = find_financial_date_entry(strip_future_prefix(call.as_call_name()));
  if (entry == nullptr) {
    // Registered names always resolve; a miss can only mean a table drift.
    return Value::error(ErrorCode::Name);
  }
  const std::uint32_t arity = call.as_call_arity();
  if (arity < entry->min_arity || arity > entry->max_arity) {
    return Value::error(ErrorCode::Value);
  }
  // This family follows the Analysis-ToolPak argument rule, reports the first
  // failing argument, and lifts element by element over array arguments, as
  // the calendar family does.
  std::vector<Value> args;
  std::optional<ErrorOriginSnapshot> error_arg_origin;
  eval_scalar_call_args(call, entry->min_arity, entry->max_arity, /*atp_rule=*/true, entry->logical_args, arena,
                        registry, ctx, args, error_arg_origin);
  const bool date1904 = ctx.date1904();
  const auto evaluated = static_cast<std::uint32_t>(args.size());
  const Value result = lift_over_array_args(args.empty() ? nullptr : args.data(), evaluated, arena,
                                            [&](const Value* a, std::uint32_t n) {
                                              return report_error_arg(entry->impl(a, n, arena, date1904), a, n, false);
                                            });
  restore_reported_origin(result, args.data(), evaluated, false, error_arg_origin);
  return result;
}

}  // namespace eval
}  // namespace formulon
