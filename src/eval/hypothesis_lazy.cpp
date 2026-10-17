//
// Implementation of the hypothesis-test / probability lazy impls: T.TEST
// (TTEST), F.TEST (FTEST), CHISQ.TEST (CHITEST), Z.TEST (ZTEST), PROB.
//
// The argument-resolution machinery is shared with `eval/regression_lazy.cpp`
// via `resolve_array_arg_na` (in `eval/range_args.{h,cpp}`), which walks a
// `Ref` / `RangeOp` / `ArrayLiteral` / scalar subtree into a `RangeResult`
// (remapping the non-array rejection to `#N/A` to match Excel's
// hypothesis-test conventions). `collect_paired_samples` pairs two resolved
// arrays for the shape-matched paths (T.TEST type==1, CHISQ.TEST, PROB);
// the independent-collection paths (T.TEST types 2 and 3, F.TEST, Z.TEST)
// use `collect_single_array_numeric`, which does NOT shape-match and
// silently drops non-numeric cells from whichever array it is handed.
//
// T.TEST, F.TEST and CHISQ.TEST evaluate through `eval/stats/hypothesis_tests.h`;
// the standard-normal tail used by Z.TEST comes from `eval/stats/special_functions.h`.

#include "eval/hypothesis_lazy.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <variant>
#include <vector>

#include "eval/builtins/numeric_helpers.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "eval/numeric_pairs.h"
#include "eval/range_args.h"
#include "eval/stats/hypothesis_tests.h"
#include "eval/stats/special_functions.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

using builtins_detail::to_finite_value;

// Paired collection for the shape-matched paths. The hypothesis family
// requires both dimensions to agree, so the transpose the regression family
// tolerates is refused here.
std::variant<Value, NumericPairs> collect_paired_samples(const parser::AstNode& a_arg, const parser::AstNode& b_arg,
                                                         Arena& arena, const FunctionRegistry& registry,
                                                         const EvalContext& ctx) {
  return collect_numeric_pairs(a_arg, b_arg, arena, registry, ctx, /*allow_transpose=*/false);
}

// Resolves a single array argument and returns the numeric cells
// (silently dropping Blank/Bool/Text). Error cells propagate via
// `*out_err`. Used by the independent-collection paths (T.TEST types
// 2/3, F.TEST, Z.TEST) where there is no pairing constraint. Returns
// `true` on success; `false` if `*out_err` was written.
bool collect_single_array_numeric(const parser::AstNode& arg_node, Arena& arena, const FunctionRegistry& registry,
                                  const EvalContext& ctx, std::vector<double>* out, Value* out_err) {
  auto resolved = resolve_array_arg_na(arg_node, arena, registry, ctx);
  if (!resolved) {
    *out_err = Value::error(resolved.error());
    return false;
  }
  RangeResult r = std::move(resolved.value());
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

// Collects the two independent samples T.TEST (types 2 and 3) and F.TEST
// both take. Error propagation runs left to right, array1 first.
//
// Returns `false` with the propagating error in `*out_err`.
bool two_samples(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                 std::vector<double>* first, std::vector<double>* second, Value* out_err) {
  return collect_single_array_numeric(call.as_call_arg(0), arena, registry, ctx, first, out_err) &&
         collect_single_array_numeric(call.as_call_arg(1), arena, registry, ctx, second, out_err);
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
  const double tails_d = std::trunc(tails_v.as_number());
  const double type_d = std::trunc(type_v.as_number());
  if (!std::isfinite(tails_d) || (tails_d != 1.0 && tails_d != 2.0)) {
    return Value::error(ErrorCode::Num);
  }
  if (!std::isfinite(type_d) || (type_d != 1.0 && type_d != 2.0 && type_d != 3.0)) {
    return Value::error(ErrorCode::Num);
  }
  const int tails = static_cast<int>(tails_d);
  const int type = static_cast<int>(type_d);

  stats::TestOutcome outcome{};
  if (type == 1) {
    // Paired t-test: require shape match, then work on every (a, b) pair
    // where both sides are numeric.
    auto prepared = collect_paired_samples(call.as_call_arg(0), call.as_call_arg(1), arena, registry, ctx);
    if (std::holds_alternative<Value>(prepared)) {
      return std::get<Value>(prepared);
    }
    const NumericPairs& pairs = std::get<NumericPairs>(prepared);
    outcome = stats::t_test_paired_excel(pairs.first.data(), pairs.second.data(), pairs.first.size(), tails);
  } else {
    // Two-sample paths collect each array independently with no pairing.
    std::vector<double> first;
    std::vector<double> second;
    Value err = Value::blank();
    if (!two_samples(call, arena, registry, ctx, &first, &second, &err)) {
      return err;
    }
    outcome = stats::t_test_two_sample_excel(first.data(), first.size(), second.data(), second.size(), tails, type);
  }
  if (outcome.div0) {
    return Value::error(ErrorCode::Div0);
  }
  return to_finite_value(outcome.p);
}

Value eval_f_test_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  if (call.as_call_arity() != 2U) {
    return Value::error(ErrorCode::Value);
  }
  std::vector<double> first;
  std::vector<double> second;
  Value err = Value::blank();
  if (!two_samples(call, arena, registry, ctx, &first, &second, &err)) {
    return err;
  }
  const stats::TestOutcome outcome = stats::f_test_excel(first.data(), first.size(), second.data(), second.size());
  if (outcome.div0) {
    return Value::error(ErrorCode::Div0);
  }
  return to_finite_value(outcome.p);
}

Value eval_chisq_test_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx) {
  if (call.as_call_arity() != 2U) {
    return Value::error(ErrorCode::Value);
  }
  auto actual_resolved = resolve_array_arg_na(call.as_call_arg(0), arena, registry, ctx);
  if (!actual_resolved) {
    return Value::error(actual_resolved.error());
  }
  RangeResult actual = std::move(actual_resolved.value());
  auto expected_resolved = resolve_array_arg_na(call.as_call_arg(1), arena, registry, ctx);
  if (!expected_resolved) {
    return Value::error(expected_resolved.error());
  }
  RangeResult expected = std::move(expected_resolved.value());
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
  // A pair with a non-numeric side (Bool, Text, Blank) is dropped, though the
  // degrees of freedom above still follow the declared shape. An expected
  // count of exactly zero divides the term by zero; negative counts are
  // allowed and only a negative statistic is rejected.
  double chi2 = 0.0;
  std::size_t used = 0;
  for (std::size_t i = 0; i < actual.cells.size(); ++i) {
    if (!actual.cells[i].is_number() || !expected.cells[i].is_number()) {
      continue;
    }
    const double e = expected.cells[i].as_number();
    if (e == 0.0) {
      return Value::error(ErrorCode::Div0);
    }
    chi2 = stats::chisq_accumulate_excel(chi2, actual.cells[i].as_number(), e);
    ++used;
  }
  if (used == 0U) {
    return Value::error(ErrorCode::Div0);
  }
  if (chi2 < 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double p = stats::chisq_test_p_excel(chi2, df);
  if (std::isnan(p)) {
    return Value::error(ErrorCode::Num);
  }
  return to_finite_value(p);
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
    // Sample SD (two-pass) needs n >= 2; a zero SD is #NUM!.
    if (samples.size() < 2U) {
      return Value::error(ErrorCode::Div0);
    }
    double ss = 0.0;
    for (double v : samples) {
      ss += (v - mean) * (v - mean);
    }
    sigma = std::sqrt(ss / (n - 1.0));
    if (sigma == 0.0) {
      return Value::error(ErrorCode::Num);
    }
  }

  // Upper-tail area of the standard normal at z.
  const double z = (std::sqrt(n) * (mean - x)) / sigma;
  const double p = stats::erfc_excel(z * 0.7071067811865476) * 0.5;
  return to_finite_value(p);
}

Value eval_prob_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 3U || arity > 4U) {
    return Value::error(ErrorCode::Value);
  }
  // Paired collection enforces shape match + shared error-scan order
  // (x_range first, then prob_range).
  auto prepared = collect_paired_samples(call.as_call_arg(0), call.as_call_arg(1), arena, registry, ctx);
  if (std::holds_alternative<Value>(prepared)) {
    return std::get<Value>(prepared);
  }
  const NumericPairs& pairs = std::get<NumericPairs>(prepared);
  // `pairs.first` holds the x-values (first arg), `pairs.second` the probs.
  // No numeric (x, prob) pair survives, e.g. text-only ranges.
  if (pairs.first.empty()) {
    return Value::error(ErrorCode::Div0);
  }
  double prob_sum = 0.0;
  for (double p : pairs.second) {
    if (p < 0.0 || p > 1.0) {
      return Value::error(ErrorCode::Num);
    }
    prob_sum += p;
  }
  // Excel accepts probability vectors that only sum to 1 within ordinary
  // rounding — e.g. three cells of 0.333333 sum to 0.999999, whose distance
  // from 1.0 is a hair over 1e-6 in IEEE-754 — yet Excel returns a number.
  // A 1e-12 absolute tolerance rejected those; the bound is loosened to
  // 1e-5 so a few decimal-rounded cells comfortably qualify while coarse
  // errors (the 1.1 / 0.9 range) still surface #NUM!.
  // NOTE: Excel's exact threshold is unverified; confirm on a live oracle.
  if (std::fabs(prob_sum - 1.0) > 1e-5) {
    return Value::error(ErrorCode::Num);
  }

  const Value lower_v = eval_node(call.as_call_arg(2), arena, registry, ctx);
  if (lower_v.is_error()) {
    return lower_v;
  }
  if (!lower_v.is_number() && !lower_v.is_blank()) {
    return Value::error(ErrorCode::Value);
  }
  const double lower = lower_v.is_blank() ? 0.0 : lower_v.as_number();

  if (arity == 3U) {
    // Upper omitted: degenerate single-point probability — sum of the
    // prob cells whose matched x equals the lower_limit exactly.
    double total = 0.0;
    for (std::size_t i = 0; i < pairs.first.size(); ++i) {
      if (pairs.first[i] == lower) {
        total += pairs.second[i];
      }
    }
    return to_finite_value(total);
  }

  const Value upper_v = eval_node(call.as_call_arg(3), arena, registry, ctx);
  if (upper_v.is_error()) {
    return upper_v;
  }
  if (!upper_v.is_number() && !upper_v.is_blank()) {
    return Value::error(ErrorCode::Value);
  }
  const double upper = upper_v.is_blank() ? 0.0 : upper_v.as_number();
  if (upper < lower) {
    // Empty interval (lower_limit > upper_limit): Mac Excel treats this
    // as zero probability mass rather than a parameter error.
    return Value::number(0.0);
  }
  double total = 0.0;
  for (std::size_t i = 0; i < pairs.first.size(); ++i) {
    const double xi = pairs.first[i];
    if (xi >= lower && xi <= upper) {
      total += pairs.second[i];
    }
  }
  return to_finite_value(total);
}

}  // namespace eval
}  // namespace formulon
