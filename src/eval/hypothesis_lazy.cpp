// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the hypothesis-test / probability lazy impls: T.TEST
// (TTEST), F.TEST (FTEST), CHISQ.TEST (CHITEST), Z.TEST (ZTEST), PROB.
//
// The argument-resolution machinery mirrors `eval/regression_lazy.cpp`:
// a `ResolvedArray` struct captures `(rows, cols, cells)` for a single
// array argument, `resolve_array_arg` walks a `Ref` / `RangeOp` /
// `ArrayLiteral` / scalar subtree into that struct (remapping the
// non-array rejection to `#N/A` to match Excel's hypothesis-test
// conventions), and `collect_numeric_pairs` pairs two resolved arrays
// for the shape-matched paths (T.TEST type==1, CHISQ.TEST, PROB). The
// independent-collection paths (T.TEST types 2 and 3, F.TEST, Z.TEST)
// use `collect_single_array_numeric`, which does NOT shape-match and
// silently drops non-numeric cells from whichever array it is handed.
//
// The helper duplication vs. `regression_lazy.cpp` is intentional: each
// TU keeps its own in-`namespace`-anonymous copy so neither pays for the
// other's specialisation, and promoting the helpers to a shared module is
// deferred until a third family needs them.
//
// Distribution primitives come from `eval/stats/special_functions.h`
// (`regularized_incomplete_beta`, `q_gamma`) and from `<cmath>`
// (`std::erfc` for the standard-normal tail used by Z.TEST).

#include "eval/hypothesis_lazy.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <variant>
#include <vector>

#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "eval/range_args.h"
#include "eval/stats/special_functions.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// One resolved array argument: flat row-major cells plus the rectangle
// shape used by the shape-match check. Mirrors the struct of the same
// name in `regression_lazy.cpp`; kept in-TU to avoid a shared header.
struct ResolvedArray {
  std::uint32_t rows;
  std::uint32_t cols;
  std::vector<Value> cells;
};

// Resolves a single array argument. Accepts `Ref`, `RangeOp`, and
// `ArrayLiteral`; any other shape (a scalar literal, a function call,
// arithmetic) yields `#N/A` because Excel's hypothesis-test family uses
// `#N/A` for shape errors rather than the `#VALUE!` used by
// SUMPRODUCT. A pre-evaluated error subtree propagates with its real
// code. Returns `true` on success; on failure writes the Excel error
// into `*out_err` and returns `false`.
bool resolve_array_arg(const parser::AstNode& arg_node, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx, ResolvedArray* out, Value* out_err) {
  const parser::NodeKind k = arg_node.kind();
  if (k == parser::NodeKind::Ref || k == parser::NodeKind::RangeOp) {
    ErrorCode err_code = ErrorCode::NA;
    if (!resolve_range_arg(arg_node, arena, registry, ctx, &out->cells, &err_code, &out->rows, &out->cols)) {
      // `resolve_range_arg` reports `#VALUE!` for non-Ref / non-RangeOp
      // shapes and `#REF!` for expansion failures. Remap the shape
      // rejection to `#N/A` to match Excel's hypothesis family.
      *out_err = Value::error(err_code == ErrorCode::Value ? ErrorCode::NA : err_code);
      return false;
    }
    return true;
  }
  if (k == parser::NodeKind::ArrayLiteral) {
    out->rows = arg_node.as_array_rows();
    out->cols = arg_node.as_array_cols();
    const std::size_t total = static_cast<std::size_t>(out->rows) * out->cols;
    out->cells.clear();
    out->cells.reserve(total);
    for (std::uint32_t r = 0; r < out->rows; ++r) {
      for (std::uint32_t c = 0; c < out->cols; ++c) {
        Value v = eval_node(arg_node.as_array_element(r, c), arena, registry, ctx);
        out->cells.push_back(v);
      }
    }
    return true;
  }
  // A scalar / call / arithmetic subtree is not a valid array here.
  // Evaluate it first so an error in the subtree propagates with its
  // real code; otherwise reject with `#N/A`.
  const Value v = eval_node(arg_node, arena, registry, ctx);
  if (v.is_error()) {
    *out_err = v;
    return false;
  }
  *out_err = Value::error(ErrorCode::NA);
  return false;
}

// Paired (x, y) numeric samples distilled from two resolved arrays.
struct NumericPairs {
  std::vector<double> x;
  std::vector<double> y;
};

// Resolves both array arguments, enforces shape match, propagates errors
// in row-major scan order (first-argument array first), and collects
// every pair whose *both* cells are numeric. If either cell in a pair is
// non-numeric the whole pair is dropped — used by T.TEST's paired mode
// and by PROB.
//
// Returns the error `Value` to propagate on the left side of the variant;
// otherwise a populated `NumericPairs` on the right.
std::variant<Value, NumericPairs> collect_numeric_pairs(const parser::AstNode& a_arg, const parser::AstNode& b_arg,
                                                        Arena& arena, const FunctionRegistry& registry,
                                                        const EvalContext& ctx) {
  ResolvedArray a{};
  Value err = Value::blank();
  if (!resolve_array_arg(a_arg, arena, registry, ctx, &a, &err)) {
    return err;
  }
  ResolvedArray b{};
  if (!resolve_array_arg(b_arg, arena, registry, ctx, &b, &err)) {
    return err;
  }
  if (a.rows != b.rows || a.cols != b.cols) {
    return Value{Value::error(ErrorCode::NA)};
  }
  // Error propagation runs over every cell in both arrays, even cells
  // that would be dropped by the text-pair rule. Scan the first argument
  // first so the leftmost-argument error wins.
  for (const Value& v : a.cells) {
    if (v.is_error()) {
      return v;
    }
  }
  for (const Value& v : b.cells) {
    if (v.is_error()) {
      return v;
    }
  }
  NumericPairs pairs;
  const std::size_t n = a.cells.size();
  pairs.x.reserve(n);
  pairs.y.reserve(n);
  for (std::size_t i = 0; i < n; ++i) {
    const Value& av = a.cells[i];
    const Value& bv = b.cells[i];
    if (!av.is_number() || !bv.is_number()) {
      continue;
    }
    // `pairs.x` carries the first argument, `pairs.y` the second.
    pairs.x.push_back(av.as_number());
    pairs.y.push_back(bv.as_number());
  }
  return pairs;
}

// Resolves a single array argument and returns the numeric cells
// (silently dropping Blank/Bool/Text). Error cells propagate via
// `*out_err`. Used by the independent-collection paths (T.TEST types
// 2/3, F.TEST, Z.TEST) where there is no pairing constraint. Returns
// `true` on success; `false` if `*out_err` was written.
bool collect_single_array_numeric(const parser::AstNode& arg_node, Arena& arena, const FunctionRegistry& registry,
                                  const EvalContext& ctx, std::vector<double>* out, Value* out_err) {
  ResolvedArray r{};
  if (!resolve_array_arg(arg_node, arena, registry, ctx, &r, out_err)) {
    return false;
  }
  for (const Value& v : r.cells) {
    if (v.is_error()) {
      *out_err = v;
      return false;
    }
  }
  out->clear();
  out->reserve(r.cells.size());
  for (const Value& v : r.cells) {
    if (v.is_number()) {
      out->push_back(v.as_number());
    }
  }
  return true;
}

// Mean and sample variance (divisor `n - 1`). Returns `false` if
// `samples.size() < 2` — the caller decides how to surface that (T.TEST
// routes it to `#DIV/0!`; Z.TEST routes it to `#DIV/0!` only when sigma
// is not supplied).
bool mean_and_sample_variance(const std::vector<double>& samples, double* out_mean, double* out_var) noexcept {
  const std::size_t n = samples.size();
  if (n < 2U) {
    return false;
  }
  double sum = 0.0;
  for (double v : samples) {
    sum += v;
  }
  const double dn = static_cast<double>(n);
  const double mean = sum / dn;
  double ss = 0.0;
  for (double v : samples) {
    const double d = v - mean;
    ss += d * d;
  }
  *out_mean = mean;
  *out_var = ss / (dn - 1.0);
  return true;
}

// Guards the final numeric result. Any NaN / infinity becomes `#NUM!`.
Value finite_number(double r) {
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// Two-tailed half-probability for Student's t: given `t` and `df`, the
// symmetric one-tailed area in the upper tail beyond |t| equals
// `0.5 * I(df/2, 1/2, df/(df+t²))`. T.TEST multiplies this by 1 or 2
// depending on `tails`.
double t_half_tail(double t, double df) noexcept {
  const double y = df / (df + t * t);
  return 0.5 * stats::regularized_incomplete_beta(df / 2.0, 0.5, y);
}

}  // namespace

Value eval_t_test_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  if (call.as_call_arity() != 4U) {
    return Value::error(ErrorCode::Value);
  }
  // Evaluate `tails` and `type` scalars first; their validity decides
  // whether we even need to touch the arrays. Errors propagate with
  // their real code.
  const Value tails_v = eval_node(call.as_call_arg(2), arena, registry, ctx);
  if (tails_v.is_error()) {
    return tails_v;
  }
  if (!tails_v.is_number()) {
    return Value::error(ErrorCode::Value);
  }
  const Value type_v = eval_node(call.as_call_arg(3), arena, registry, ctx);
  if (type_v.is_error()) {
    return type_v;
  }
  if (!type_v.is_number()) {
    return Value::error(ErrorCode::Value);
  }
  // Excel truncates (not rounds) `tails` and `type` toward zero before
  // the domain check, matching the documented behaviour.
  const int tails = static_cast<int>(tails_v.as_number());
  const int type = static_cast<int>(type_v.as_number());
  if (tails != 1 && tails != 2) {
    return Value::error(ErrorCode::Num);
  }
  if (type != 1 && type != 2 && type != 3) {
    return Value::error(ErrorCode::Num);
  }

  double t = 0.0;
  double df = 0.0;

  if (type == 1) {
    // Paired t-test: require shape match, then compute on pairwise
    // differences for every (a, b) pair where both sides are numeric.
    auto prepared = collect_numeric_pairs(call.as_call_arg(0), call.as_call_arg(1), arena, registry, ctx);
    if (std::holds_alternative<Value>(prepared)) {
      return std::get<Value>(prepared);
    }
    const NumericPairs& pairs = std::get<NumericPairs>(prepared);
    std::vector<double> diffs;
    diffs.reserve(pairs.x.size());
    for (std::size_t i = 0; i < pairs.x.size(); ++i) {
      diffs.push_back(pairs.x[i] - pairs.y[i]);
    }
    double mean_d = 0.0;
    double var_d = 0.0;
    if (!mean_and_sample_variance(diffs, &mean_d, &var_d)) {
      return Value::error(ErrorCode::Div0);
    }
    if (var_d == 0.0) {
      // Zero within-pair variance: the null hypothesis cannot be tested
      // (either every difference is identical to the sample mean, so
      // s_d == 0 and t is 0/0, or every pair is identical so mean_d == 0
      // too). Excel surfaces this as `#DIV/0!`.
      return Value::error(ErrorCode::Div0);
    }
    const double n = static_cast<double>(diffs.size());
    t = mean_d / std::sqrt(var_d / n);
    df = n - 1.0;
  } else {
    // Two-sample paths collect each array independently with no pairing.
    // Error propagation runs left-to-right (array1 first).
    std::vector<double> a;
    std::vector<double> b;
    Value err = Value::blank();
    if (!collect_single_array_numeric(call.as_call_arg(0), arena, registry, ctx, &a, &err)) {
      return err;
    }
    if (!collect_single_array_numeric(call.as_call_arg(1), arena, registry, ctx, &b, &err)) {
      return err;
    }
    double m1 = 0.0;
    double m2 = 0.0;
    double v1 = 0.0;
    double v2 = 0.0;
    if (!mean_and_sample_variance(a, &m1, &v1) || !mean_and_sample_variance(b, &m2, &v2)) {
      return Value::error(ErrorCode::Div0);
    }
    const double n1 = static_cast<double>(a.size());
    const double n2 = static_cast<double>(b.size());
    if (type == 2) {
      // Pooled-variance equal-variance t-test.
      const double sp2 = ((n1 - 1.0) * v1 + (n2 - 1.0) * v2) / (n1 + n2 - 2.0);
      if (sp2 <= 0.0) {
        return Value::error(ErrorCode::Div0);
      }
      t = (m1 - m2) / std::sqrt(sp2 * (1.0 / n1 + 1.0 / n2));
      df = n1 + n2 - 2.0;
    } else {
      // Welch's t-test: separate per-sample variance estimates.
      const double u1 = v1 / n1;
      const double u2 = v2 / n2;
      const double denom = u1 + u2;
      if (denom <= 0.0) {
        return Value::error(ErrorCode::Div0);
      }
      t = (m1 - m2) / std::sqrt(denom);
      const double df_denom = (u1 * u1) / (n1 - 1.0) + (u2 * u2) / (n2 - 1.0);
      if (df_denom <= 0.0) {
        return Value::error(ErrorCode::Div0);
      }
      df = (denom * denom) / df_denom;
    }
  }

  const double half = t_half_tail(t, df);
  const double p = (tails == 1) ? half : 2.0 * half;
  return finite_number(p);
}

Value eval_f_test_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  if (call.as_call_arity() != 2U) {
    return Value::error(ErrorCode::Value);
  }
  std::vector<double> a;
  std::vector<double> b;
  Value err = Value::blank();
  if (!collect_single_array_numeric(call.as_call_arg(0), arena, registry, ctx, &a, &err)) {
    return err;
  }
  if (!collect_single_array_numeric(call.as_call_arg(1), arena, registry, ctx, &b, &err)) {
    return err;
  }
  double m1 = 0.0;
  double m2 = 0.0;
  double v1 = 0.0;
  double v2 = 0.0;
  if (!mean_and_sample_variance(a, &m1, &v1) || !mean_and_sample_variance(b, &m2, &v2)) {
    return Value::error(ErrorCode::Div0);
  }
  if (v1 == 0.0 || v2 == 0.0) {
    return Value::error(ErrorCode::Div0);
  }
  const double df1 = static_cast<double>(a.size()) - 1.0;
  const double df2 = static_cast<double>(b.size()) - 1.0;
  const double f = v1 / v2;
  // CDF of F(df1, df2) at `f` is `I(df1/2, df2/2, df1*f / (df1*f + df2))`.
  const double x = (df1 * f) / (df1 * f + df2);
  const double cdf = stats::regularized_incomplete_beta(df1 / 2.0, df2 / 2.0, x);
  if (std::isnan(cdf)) {
    return Value::error(ErrorCode::Num);
  }
  const double lo = cdf < (1.0 - cdf) ? cdf : (1.0 - cdf);
  return finite_number(2.0 * lo);
}

Value eval_chisq_test_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx) {
  if (call.as_call_arity() != 2U) {
    return Value::error(ErrorCode::Value);
  }
  ResolvedArray actual{};
  Value err = Value::blank();
  if (!resolve_array_arg(call.as_call_arg(0), arena, registry, ctx, &actual, &err)) {
    return err;
  }
  ResolvedArray expected{};
  if (!resolve_array_arg(call.as_call_arg(1), arena, registry, ctx, &expected, &err)) {
    return err;
  }
  if (actual.rows != expected.rows || actual.cols != expected.cols) {
    return Value::error(ErrorCode::NA);
  }
  // Errors propagate, left-to-right (actual first).
  for (const Value& v : actual.cells) {
    if (v.is_error()) {
      return v;
    }
  }
  for (const Value& v : expected.cells) {
    if (v.is_error()) {
      return v;
    }
  }
  // CHISQ.TEST expects a clean contingency table: any non-numeric cell
  // is a structural problem, not a silent-drop case. Return `#VALUE!`.
  for (const Value& v : actual.cells) {
    if (!v.is_number()) {
      return Value::error(ErrorCode::Value);
    }
  }
  for (const Value& v : expected.cells) {
    if (!v.is_number()) {
      return Value::error(ErrorCode::Value);
    }
  }
  // Degrees of freedom: (r-1)(c-1) for a true 2-D contingency table;
  // `n - 1` for a 1-D sequence (a single row or a single column). Fewer
  // than one degree of freedom leaves no test to run; Mac Excel surfaces
  // the degenerate-shape (e.g. 1x1) case as `#N/A`, distinct from the
  // `#DIV/0!` it returns when an expected count is exactly zero.
  double df = 0.0;
  if (actual.rows > 1U && actual.cols > 1U) {
    df = static_cast<double>((actual.rows - 1U) * (actual.cols - 1U));
  } else {
    const std::size_t n = actual.cells.size();
    df = static_cast<double>(n) - 1.0;
  }
  if (df < 1.0) {
    return Value::error(ErrorCode::NA);
  }
  // Any expected cell of exactly zero divides the chi-squared term by
  // zero. Mac Excel reports this as `#DIV/0!` (strict-zero check; a tiny
  // positive expected count is fine).
  for (const Value& v : expected.cells) {
    if (v.as_number() == 0.0) {
      return Value::error(ErrorCode::Div0);
    }
  }
  double chi2 = 0.0;
  for (std::size_t i = 0; i < actual.cells.size(); ++i) {
    const double e = expected.cells[i].as_number();
    if (e < 0.0) {
      return Value::error(ErrorCode::Num);
    }
    const double diff = actual.cells[i].as_number() - e;
    chi2 += (diff * diff) / e;
  }
  const double p = stats::q_gamma(df / 2.0, chi2 / 2.0);
  if (std::isnan(p)) {
    return Value::error(ErrorCode::Num);
  }
  return finite_number(p);
}

Value eval_z_test_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 2U || arity > 3U) {
    return Value::error(ErrorCode::Value);
  }
  // Scalar `x` is evaluated eagerly so its error (if any) takes
  // precedence over anything in the array argument.
  const Value x_v = eval_node(call.as_call_arg(1), arena, registry, ctx);
  if (x_v.is_error()) {
    return x_v;
  }
  if (!x_v.is_number()) {
    return Value::error(ErrorCode::Value);
  }
  const double x = x_v.as_number();

  std::vector<double> samples;
  Value err = Value::blank();
  if (!collect_single_array_numeric(call.as_call_arg(0), arena, registry, ctx, &samples, &err)) {
    return err;
  }
  if (samples.empty()) {
    return Value::error(ErrorCode::NA);
  }
  double sum = 0.0;
  for (double v : samples) {
    sum += v;
  }
  const double n = static_cast<double>(samples.size());
  const double mean = sum / n;

  double sigma = 0.0;
  if (arity == 3U) {
    const Value sigma_v = eval_node(call.as_call_arg(2), arena, registry, ctx);
    if (sigma_v.is_error()) {
      return sigma_v;
    }
    if (!sigma_v.is_number()) {
      return Value::error(ErrorCode::Value);
    }
    sigma = sigma_v.as_number();
    if (sigma <= 0.0) {
      return Value::error(ErrorCode::Num);
    }
  } else {
    // No sigma supplied: use the sample standard deviation. Excel needs
    // n >= 2 in that case (the sample SD is undefined at n == 1) and
    // errors with `#DIV/0!` if the sample SD comes out exactly zero.
    double m_ignored = 0.0;
    double var = 0.0;
    if (!mean_and_sample_variance(samples, &m_ignored, &var)) {
      return Value::error(ErrorCode::Div0);
    }
    if (var <= 0.0) {
      return Value::error(ErrorCode::Div0);
    }
    sigma = std::sqrt(var);
  }

  // Z-statistic on the sample mean under the null hypothesis μ == x.
  const double z = (mean - x) / (sigma / std::sqrt(n));
  // One-tailed upper area: `1 - Φ(z) = 0.5 * erfc(z / sqrt(2))`.
  const double p = 0.5 * std::erfc(z / std::sqrt(2.0));
  return finite_number(p);
}

Value eval_prob_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 3U || arity > 4U) {
    return Value::error(ErrorCode::Value);
  }
  // Paired collection enforces shape match + shared error-scan order
  // (x_range first, then prob_range).
  auto prepared = collect_numeric_pairs(call.as_call_arg(0), call.as_call_arg(1), arena, registry, ctx);
  if (std::holds_alternative<Value>(prepared)) {
    return std::get<Value>(prepared);
  }
  const NumericPairs& pairs = std::get<NumericPairs>(prepared);
  // `pairs.x` holds the x-values (first arg), `pairs.y` the probs.
  if (pairs.x.empty()) {
    return Value::error(ErrorCode::NA);
  }
  double prob_sum = 0.0;
  for (double p : pairs.y) {
    if (p < 0.0 || p > 1.0) {
      return Value::error(ErrorCode::Num);
    }
    prob_sum += p;
  }
  // Excel accepts a tiny amount of floating-point slop in the sum; an
  // absolute tolerance of 1e-12 is well inside the round-off noise for
  // any realistic probability vector.
  if (std::fabs(prob_sum - 1.0) > 1e-12) {
    return Value::error(ErrorCode::Num);
  }

  const Value lower_v = eval_node(call.as_call_arg(2), arena, registry, ctx);
  if (lower_v.is_error()) {
    return lower_v;
  }
  if (!lower_v.is_number()) {
    return Value::error(ErrorCode::Value);
  }
  const double lower = lower_v.as_number();

  if (arity == 3U) {
    // Upper omitted: degenerate single-point probability — sum of the
    // prob cells whose matched x equals the lower_limit exactly.
    double total = 0.0;
    for (std::size_t i = 0; i < pairs.x.size(); ++i) {
      if (pairs.x[i] == lower) {
        total += pairs.y[i];
      }
    }
    return finite_number(total);
  }

  const Value upper_v = eval_node(call.as_call_arg(3), arena, registry, ctx);
  if (upper_v.is_error()) {
    return upper_v;
  }
  if (!upper_v.is_number()) {
    return Value::error(ErrorCode::Value);
  }
  const double upper = upper_v.as_number();
  if (upper < lower) {
    // Empty interval (lower_limit > upper_limit): Mac Excel treats this
    // as zero probability mass rather than a parameter error.
    return Value::number(0.0);
  }
  double total = 0.0;
  for (std::size_t i = 0; i < pairs.x.size(); ++i) {
    const double xi = pairs.x[i];
    if (xi >= lower && xi <= upper) {
      total += pairs.y[i];
    }
  }
  return finite_number(total);
}

}  // namespace eval
}  // namespace formulon
