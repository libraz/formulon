// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the rank / percentile-rank lazy impls:
// RANK (legacy), RANK.EQ, RANK.AVG, PERCENTRANK (legacy),
// PERCENTRANK.INC, PERCENTRANK.EXC.
//
// Every function shares the same front-end work: resolve the array
// argument — which may be a `Ref`, a `RangeOp`, or an inline
// `ArrayLiteral` — into a flat vector of `Value`s, propagate any error
// cell in scan order, then filter down to the numeric cells only
// (Text / Bool / Blank are skipped, matching MEDIAN / LARGE / SMALL
// semantics in `src/eval/builtins/stats.cpp`). The scalar arguments
// (`number` / `x` / `order` / `significance`) are evaluated through
// `eval_node` and checked for numeric kind directly — `Bool` is
// rejected with `#VALUE!` in the slots where Excel does so, rather
// than being coerced through `coerce_to_number`.

#include "eval/rank_lazy.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <variant>
#include <vector>

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

// Resolves the array / Ref / RangeOp / ArrayLiteral argument into a
// vector of raw `Value`s in row-major order. Accepts any shape accepted
// by `resolve_range_arg` plus inline ArrayLiterals; on an unsupported
// shape (bare scalar, function call other than OFFSET, arithmetic) we
// evaluate the subtree to allow its own error to propagate and
// otherwise surface `#VALUE!`. Returns `true` on success; on failure
// writes the Excel error into `*out_err` and returns `false`.
bool resolve_array_cells(const parser::AstNode& arg_node, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx, std::vector<Value>* out_cells, Value* out_err) {
  const parser::NodeKind k = arg_node.kind();
  if (k == parser::NodeKind::Ref || k == parser::NodeKind::RangeOp) {
    ErrorCode err_code = ErrorCode::Value;
    if (!resolve_range_arg(arg_node, arena, registry, ctx, out_cells, &err_code, nullptr, nullptr)) {
      *out_err = Value::error(err_code);
      return false;
    }
    return true;
  }
  if (k == parser::NodeKind::ArrayLiteral) {
    const std::uint32_t rows = arg_node.as_array_rows();
    const std::uint32_t cols = arg_node.as_array_cols();
    const std::size_t total = static_cast<std::size_t>(rows) * cols;
    out_cells->clear();
    out_cells->reserve(total);
    for (std::uint32_t r = 0; r < rows; ++r) {
      for (std::uint32_t c = 0; c < cols; ++c) {
        out_cells->push_back(eval_node(arg_node.as_array_element(r, c), arena, registry, ctx));
      }
    }
    return true;
  }
  // Scalar / call / arithmetic subtree: evaluate so a genuine error
  // inside the subtree propagates with its real code; otherwise reject
  // with `#VALUE!` because a bare scalar is not a valid array argument
  // for RANK / PERCENTRANK.
  const Value v = eval_node(arg_node, arena, registry, ctx);
  if (v.is_error()) {
    *out_err = v;
    return false;
  }
  *out_err = Value::error(ErrorCode::Value);
  return false;
}

// Collects the numeric cells of the resolved array in scan order,
// propagating any error cell first (matching Excel's left-to-right
// error precedence). Returns the error `Value` on the left of the
// variant, otherwise the numeric samples on the right.
std::variant<Value, std::vector<double>> collect_rank_array(const parser::AstNode& arg_node, Arena& arena,
                                                            const FunctionRegistry& registry, const EvalContext& ctx) {
  std::vector<Value> cells;
  Value err = Value::blank();
  if (!resolve_array_cells(arg_node, arena, registry, ctx, &cells, &err)) {
    return err;
  }
  for (const Value& v : cells) {
    if (v.is_error()) {
      return v;
    }
  }
  std::vector<double> nums;
  nums.reserve(cells.size());
  for (const Value& v : cells) {
    // Skip Text / Bool / Blank silently — matches MEDIAN / LARGE /
    // SMALL semantics on mixed ranges.
    if (v.is_number()) {
      nums.push_back(v.as_number());
    }
  }
  return nums;
}

// Evaluates one AST arg as a scalar number. Errors propagate; Bool,
// Text, Blank, and any non-numeric kind surface as
// `Value::error(on_non_numeric)` — callers pass `#VALUE!` for slots
// where Excel rejects bool coercion (the `number` / `x` arguments),
// which is the only usage today.
std::variant<Value, double> read_scalar_number(const parser::AstNode& arg, Arena& arena,
                                               const FunctionRegistry& registry, const EvalContext& ctx,
                                               ErrorCode on_non_numeric) {
  const Value v = eval_node(arg, arena, registry, ctx);
  if (v.is_error()) {
    return v;
  }
  if (!v.is_number()) {
    return Value{Value::error(on_non_numeric)};
  }
  return v.as_number();
}

// Truncates `raw` to `significance` fractional digits (>= 1). Excel
// truncates toward zero rather than rounding; implemented as
// `trunc(raw * 10^sig) / 10^sig`.
double truncate_to_significance(double raw, std::int64_t significance) {
  const double mult = std::pow(10.0, static_cast<double>(significance));
  return std::trunc(raw * mult) / mult;
}

// Extracts and validates the optional `significance` argument shared by
// PERCENTRANK.INC and PERCENTRANK.EXC. Default is 3; non-numeric ->
// `#VALUE!`; values < 1 (after truncation toward zero) -> `#NUM!`.
// Returns the error `Value` on the left of the variant, otherwise the
// truncated integer significance on the right.
std::variant<Value, std::int64_t> read_significance(const parser::AstNode& call, Arena& arena,
                                                    const FunctionRegistry& registry, const EvalContext& ctx) {
  if (call.as_call_arity() < 3U) {
    return static_cast<std::int64_t>(3);
  }
  auto raw = read_scalar_number(call.as_call_arg(2), arena, registry, ctx, ErrorCode::Value);
  if (std::holds_alternative<Value>(raw)) {
    return std::get<Value>(raw);
  }
  const double sig_d = std::trunc(std::get<double>(raw));
  if (sig_d < 1.0) {
    return Value{Value::error(ErrorCode::Num)};
  }
  return static_cast<std::int64_t>(sig_d);
}

// Shared RANK front-end: decode (number, ref, [order]) arguments,
// propagate errors, and collect the array. On any failure returns the
// error `Value` on the left of the variant; on success the three
// pieces are laid out on the right.
struct RankInputs {
  double number;
  bool descending;
  std::vector<double> values;
};

std::variant<Value, RankInputs> prepare_rank(const parser::AstNode& call, Arena& arena,
                                             const FunctionRegistry& registry, const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 2U || arity > 3U) {
    return Value{Value::error(ErrorCode::Value)};
  }
  // Argument 0: `number`. Bool / Text / Blank -> #VALUE! (Excel rejects
  // a bool in this slot).
  auto number = read_scalar_number(call.as_call_arg(0), arena, registry, ctx, ErrorCode::Value);
  if (std::holds_alternative<Value>(number)) {
    return std::get<Value>(number);
  }
  // Argument 2 (optional): `order`. Any nonzero value -> ascending;
  // 0 or omitted -> descending.
  bool descending = true;
  if (arity == 3U) {
    auto order = read_scalar_number(call.as_call_arg(2), arena, registry, ctx, ErrorCode::Value);
    if (std::holds_alternative<Value>(order)) {
      return std::get<Value>(order);
    }
    descending = std::get<double>(order) == 0.0;
  }
  // Argument 1: the `ref` array.
  auto arr = collect_rank_array(call.as_call_arg(1), arena, registry, ctx);
  if (std::holds_alternative<Value>(arr)) {
    return std::get<Value>(arr);
  }
  return RankInputs{std::get<double>(number), descending, std::move(std::get<std::vector<double>>(arr))};
}

// Counts strictly-better and equal values for `number` within `values`
// according to the chosen order. "Strictly better" means strictly
// greater when descending, strictly less when ascending.
struct RankCounts {
  std::size_t greater;  // strictly better (see above)
  std::size_t equal;    // exact FP equality to `number`
};

RankCounts count_rank(const std::vector<double>& values, double number, bool descending) {
  RankCounts c{0U, 0U};
  for (const double v : values) {
    if (v == number) {
      ++c.equal;
      continue;
    }
    if (descending ? v > number : v < number) {
      ++c.greater;
    }
  }
  return c;
}

}  // namespace

Value eval_rank_eq_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx) {
  auto prepared = prepare_rank(call, arena, registry, ctx);
  if (std::holds_alternative<Value>(prepared)) {
    return std::get<Value>(prepared);
  }
  const RankInputs& in = std::get<RankInputs>(prepared);
  if (in.values.empty()) {
    return Value::error(ErrorCode::NA);
  }
  const RankCounts c = count_rank(in.values, in.number, in.descending);
  // Excel returns #N/A when `number` is not present in the filtered
  // numeric array.
  if (c.equal == 0U) {
    return Value::error(ErrorCode::NA);
  }
  return Value::number(static_cast<double>(c.greater + 1U));
}

Value eval_rank_avg_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx) {
  auto prepared = prepare_rank(call, arena, registry, ctx);
  if (std::holds_alternative<Value>(prepared)) {
    return std::get<Value>(prepared);
  }
  const RankInputs& in = std::get<RankInputs>(prepared);
  if (in.values.empty()) {
    return Value::error(ErrorCode::NA);
  }
  const RankCounts c = count_rank(in.values, in.number, in.descending);
  if (c.equal == 0U) {
    return Value::error(ErrorCode::NA);
  }
  // Average of the `equal` contiguous rank slots starting at position
  // `greater + 1` (1-based). Closed-form: midpoint = greater + 1 +
  // (equal - 1) / 2 = greater + (equal + 1) / 2.
  const double avg = static_cast<double>(c.greater) + (static_cast<double>(c.equal) + 1.0) / 2.0;
  return Value::number(avg);
}

namespace {

// Shared PERCENTRANK front-end: decode (array, x, [significance]),
// propagate errors, collect the array, sort ascending, validate the
// significance, and hand back the pieces the two variants need.
struct PercentRankInputs {
  double x;
  std::int64_t significance;
  std::vector<double> sorted;
};

std::variant<Value, PercentRankInputs> prepare_percentrank(const parser::AstNode& call, Arena& arena,
                                                           const FunctionRegistry& registry, const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 2U || arity > 3U) {
    return Value{Value::error(ErrorCode::Value)};
  }
  // Argument 0: the array.
  auto arr = collect_rank_array(call.as_call_arg(0), arena, registry, ctx);
  if (std::holds_alternative<Value>(arr)) {
    return std::get<Value>(arr);
  }
  // Argument 1: `x`. Non-numeric -> #VALUE!.
  auto x = read_scalar_number(call.as_call_arg(1), arena, registry, ctx, ErrorCode::Value);
  if (std::holds_alternative<Value>(x)) {
    return std::get<Value>(x);
  }
  // Argument 2 (optional): `significance`.
  auto sig = read_significance(call, arena, registry, ctx);
  if (std::holds_alternative<Value>(sig)) {
    return std::get<Value>(sig);
  }
  PercentRankInputs out{std::get<double>(x), std::get<std::int64_t>(sig),
                        std::move(std::get<std::vector<double>>(arr))};
  std::sort(out.sorted.begin(), out.sorted.end());
  return out;
}

}  // namespace

Value eval_percentrank_inc_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                                const EvalContext& ctx) {
  auto prepared = prepare_percentrank(call, arena, registry, ctx);
  if (std::holds_alternative<Value>(prepared)) {
    return std::get<Value>(prepared);
  }
  const PercentRankInputs& in = std::get<PercentRankInputs>(prepared);
  const std::size_t n = in.sorted.size();
  // Excel returns #N/A for an array with fewer than two numeric cells
  // (the `(N - 1)` divisor collapses).
  if (n < 2U) {
    return Value::error(ErrorCode::NA);
  }
  if (in.x < in.sorted.front() || in.x > in.sorted.back()) {
    return Value::error(ErrorCode::NA);
  }
  // Find the lowest index k such that sorted[k] <= x < sorted[k + 1].
  // When x equals an element exactly, we want the lowest index of that
  // element (Excel reports the lowest rank for duplicates).
  std::size_t k = 0U;
  for (std::size_t i = 0; i < n; ++i) {
    if (in.sorted[i] <= in.x) {
      k = i;
    } else {
      break;
    }
  }
  // Scan backwards over an equal run so k points at the first occurrence
  // of the value, matching Excel's observed "lowest rank wins" rule.
  while (k > 0U && in.sorted[k - 1U] == in.sorted[k]) {
    --k;
  }
  double raw = 0.0;
  if (in.sorted[k] == in.x) {
    raw = static_cast<double>(k) / static_cast<double>(n - 1U);
  } else {
    // Interpolate between sorted[k] and sorted[k + 1]. The outer range
    // check above guarantees k + 1 < n here.
    const double span = in.sorted[k + 1U] - in.sorted[k];
    raw = (static_cast<double>(k) + (in.x - in.sorted[k]) / span) / static_cast<double>(n - 1U);
  }
  return Value::number(truncate_to_significance(raw, in.significance));
}

Value eval_percentrank_exc_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                                const EvalContext& ctx) {
  auto prepared = prepare_percentrank(call, arena, registry, ctx);
  if (std::holds_alternative<Value>(prepared)) {
    return std::get<Value>(prepared);
  }
  const PercentRankInputs& in = std::get<PercentRankInputs>(prepared);
  const std::size_t n = in.sorted.size();
  if (n < 1U) {
    return Value::error(ErrorCode::NA);
  }
  if (in.x < in.sorted.front() || in.x > in.sorted.back()) {
    return Value::error(ErrorCode::NA);
  }
  // Exclusive uses divisor (N + 1) and 1-based positions, so an exact
  // match at sorted[k] yields raw = (k + 1) / (N + 1).
  std::size_t k = 0U;
  for (std::size_t i = 0; i < n; ++i) {
    if (in.sorted[i] <= in.x) {
      k = i;
    } else {
      break;
    }
  }
  while (k > 0U && in.sorted[k - 1U] == in.sorted[k]) {
    --k;
  }
  const double denom = static_cast<double>(n + 1U);
  double raw = 0.0;
  if (in.sorted[k] == in.x) {
    raw = static_cast<double>(k + 1U) / denom;
  } else {
    // k + 1 < n because x < sorted.back() when no exact match is found.
    const double span = in.sorted[k + 1U] - in.sorted[k];
    raw = (static_cast<double>(k + 1U) + (in.x - in.sorted[k]) / span) / denom;
  }
  return Value::number(truncate_to_significance(raw, in.significance));
}

}  // namespace eval
}  // namespace formulon
