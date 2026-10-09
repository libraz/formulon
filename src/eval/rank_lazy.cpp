//
// Implementation of the rank / percentile-rank lazy impls:
// RANK (legacy), RANK.EQ, RANK.AVG, PERCENTRANK (legacy),
// PERCENTRANK.INC, PERCENTRANK.EXC.
//
// Every function shares the same front-end work: resolve the array
// argument — any range, spilled range, array literal or dynamic-array
// expression — into a flat vector of `Value`s, propagate any error
// cell in scan order, then filter down to the numeric cells only
// (Text / Bool / Blank are skipped, matching MEDIAN / LARGE / SMALL
// semantics in `src/eval/builtins/stats.cpp`). The scalar arguments
// split by slot:
//   * RANK's `number` slot goes through `coerce_scalar_number` so Mac
//     Excel 365's lenient coercion applies (Number passes through;
//     Bool -> 1 / 0; Blank -> 0; text-numeric parsed; non-numeric
//     text -> `#VALUE!`).
//   * RANK's `order` slot is hand-coded: Number / Bool / Blank are
//     accepted (Bool -> 1 / 0, Blank -> 0) but ANY text -- numeric or
//     not -- yields `#VALUE!`. Mac Excel rejects `=RANK(20, A1:A3,
//     "0")` even though `"0"` is text-numeric.
//   * PERCENTRANK's `x` and `significance` coerce as RANK's `number`
//     does, and an array in either evaluates the function per element.

#include "eval/rank_lazy.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <variant>
#include <vector>

#include "eval/array_alloc.h"
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

// Resolves the array argument into a vector of raw `Value`s in row-major
// order. Accepts every shape `resolve_range_arg` resolves — `Ref` /
// `RangeOp` / `SpillRef` / `ArrayLiteral` and dynamic-array producers
// such as `SEQUENCE`. A subtree that collapses to a bare scalar is
// rejected with `#VALUE!` because a lone value is not a valid array
// argument for RANK / PERCENTRANK; an error inside the subtree
// propagates with its own code. Returns `true` on success; on failure
// writes the Excel error into `*out_err` and returns `false`.
bool resolve_array_cells(const parser::AstNode& arg_node, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx, std::vector<Value>* out_cells, Value* out_err) {
  auto resolved = resolve_range_arg_no_scalar(arg_node, arena, registry, ctx, ErrorCode::Value);
  if (!resolved) {
    *out_err = Value::error(resolved.error());
    return false;
  }
  *out_cells = std::move(resolved.value().cells);
  return true;
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

/// Evaluates one AST arg as a scalar number using Excel's lenient
/// coercion rules: Number passes through; Bool -> 1.0 / 0.0; Blank ->
/// 0.0; text-numeric is parsed; non-numeric text yields `#VALUE!`;
/// non-finite numbers yield `#NUM!`. Errors propagate. Used by the
/// RANK `number` and `order` slots so probes like
/// `=RANK(TRUE, A1:A3, 1)` and `=RANK("20", A1:A3, 0)` match Mac
/// Excel 365 ja-JP behaviour.
std::variant<Value, double> coerce_scalar_number(const parser::AstNode& arg, Arena& arena,
                                                 const FunctionRegistry& registry, const EvalContext& ctx) {
  const Value v = eval_node(arg, arena, registry, ctx);
  if (v.is_error()) {
    return v;
  }
  auto coerced = coerce_to_number(v);
  if (!coerced) {
    return Value{Value::error(coerced.error())};
  }
  return coerced.value();
}

// Truncates `raw` to `significance` fractional digits (>= 1). Excel
// truncates toward zero rather than rounding; implemented as
// `trunc(raw * 10^sig) / 10^sig`.
double truncate_to_significance(double raw, std::int64_t significance) {
  const double mult = std::pow(10.0, static_cast<double>(significance));
  return std::trunc(raw * mult) / mult;
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
  // Argument 0: `number`. Mac Excel 365 coerces Bool (TRUE -> 1, FALSE
  // -> 0), text-numeric (`"20"` -> 20), and Blank (-> 0); only truly
  // non-numeric text yields `#VALUE!`. Verified by the
  // `rank_number_text_numeric`, `rank_number_bool_true`, and
  // `rank_number_blank_arg` probes in
  // `tests/oracle/cases/comparison_rank_probes.yaml`.
  auto number = coerce_scalar_number(call.as_call_arg(0), arena, registry, ctx);
  if (std::holds_alternative<Value>(number)) {
    return std::get<Value>(number);
  }
  // Argument 2 (optional): `order`. Mac Excel 365 accepts Number / Bool
  // / Blank here but rejects ANY text (numeric or not) with `#VALUE!`.
  // That asymmetry vs the `number` slot is real: probes
  // `rank_order_true_bool` / `rank_order_false_bool` succeed but
  // `rank_order_text_zero` / `rank_order_text_one` /
  // `rank_order_text_nonnumeric` all return `#VALUE!`.
  // Accordingly we evaluate the arg, propagate errors, accept Number
  // as-is, coerce Bool (TRUE -> 1, FALSE -> 0) and Blank (-> 0), and
  // reject everything else with `#VALUE!`. Any nonzero value ->
  // ascending; 0 or omitted -> descending.
  bool descending = true;
  if (arity == 3U) {
    const Value v = eval_node(call.as_call_arg(2), arena, registry, ctx);
    if (v.is_error()) {
      return v;
    }
    double order_d = 0.0;
    switch (v.kind()) {
      case ValueKind::Number:
        order_d = v.as_number();
        break;
      case ValueKind::Bool:
        order_d = v.as_boolean() ? 1.0 : 0.0;
        break;
      case ValueKind::Blank:
        order_d = 0.0;
        break;
      default:
        return Value::error(ErrorCode::Value);
    }
    descending = order_d == 0.0;
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

namespace {

// RANK.EQ and RANK.AVG share the lookup; they differ only in how ties rank.
Value eval_rank_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx, bool average_ties) {
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
  if (!average_ties) {
    return Value::number(static_cast<double>(c.greater + 1U));
  }
  // Average of the `equal` contiguous rank slots starting at position
  // `greater + 1` (1-based). Closed-form: midpoint = greater + 1 +
  // (equal - 1) / 2 = greater + (equal + 1) / 2.
  const double avg = static_cast<double>(c.greater) + (static_cast<double>(c.equal) + 1.0) / 2.0;
  return Value::number(avg);
}

}  // namespace

Value eval_rank_eq_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx) {
  return eval_rank_lazy(call, arena, registry, ctx, /*average_ties=*/false);
}

Value eval_rank_avg_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx) {
  return eval_rank_lazy(call, arena, registry, ctx, /*average_ties=*/true);
}

namespace {

// Finds the highest index k with sorted[k] <= x. Requires a non-empty
// `sorted` with `sorted.front() <= x`.
std::size_t percentrank_floor_index(const std::vector<double>& sorted, double x) {
  std::size_t k = 0U;
  for (std::size_t i = 0; i < sorted.size(); ++i) {
    if (sorted[i] <= x) {
      k = i;
    } else {
      break;
    }
  }
  return k;
}

// PERCENTRANK.INC / .EXC of one `x` at one `significance` over the sorted
// numeric array. The inclusive form ranks exact matches at k / (N - 1); the
// exclusive one uses 1-based positions over (N + 1), so an exact match at
// sorted[k] yields (k + 1) / (N + 1).
Value percentrank_cell(const std::vector<double>& sorted, const Value& x_v, const Value& sig_v, bool exclusive) {
  if (x_v.is_error()) {
    return x_v;
  }
  const auto x_or = coerce_to_number(x_v);
  if (!x_or) {
    return Value::error(x_or.error());
  }
  if (sig_v.is_error()) {
    return sig_v;
  }
  const auto sig_or = coerce_to_number(sig_v);
  if (!sig_or) {
    return Value::error(sig_or.error());
  }
  // Significance truncates toward zero and must be at least 1.
  const double sig_d = std::trunc(sig_or.value());
  if (!(sig_d >= 1.0) || sig_d > static_cast<double>(std::numeric_limits<std::int64_t>::max())) {
    return Value::error(ErrorCode::Num);
  }
  const auto significance = static_cast<std::int64_t>(sig_d);
  const double x = x_or.value();
  const std::size_t n = sorted.size();
  // Excel returns #N/A for an empty array, and for the inclusive form also
  // for a single numeric cell (the `(N - 1)` divisor collapses).
  if (n < (exclusive ? 1U : 2U)) {
    return Value::error(ErrorCode::NA);
  }
  if (x < sorted.front() || x > sorted.back()) {
    return Value::error(ErrorCode::NA);
  }
  std::size_t k = percentrank_floor_index(sorted, x);
  const bool exact = sorted[k] == x;
  // An exact match reports the lowest rank of a duplicate run; an
  // interpolated x keeps the run's last index as its anchor.
  if (exact) {
    while (k > 0U && sorted[k - 1U] == sorted[k]) {
      --k;
    }
  }
  const std::size_t pos = exclusive ? k + 1U : k;
  const double denom = static_cast<double>(exclusive ? n + 1U : n - 1U);
  double raw = 0.0;
  if (exact) {
    raw = static_cast<double>(pos) / denom;
  } else {
    // Interpolate between sorted[k] and sorted[k + 1]. The outer range
    // check above guarantees k + 1 < n here.
    const double span = sorted[k + 1U] - sorted[k];
    raw = (static_cast<double>(pos) + (x - sorted[k]) / span) / denom;
  }
  const double result = truncate_to_significance(raw, significance);
  return std::isfinite(result) ? Value::number(result) : Value::error(ErrorCode::Num);
}

// Element (r, c) of `v` under 1xN / Nx1 broadcasting; a scalar supplies
// every cell and an array too small for the cell gives #N/A.
Value broadcast_at(const Value& v, std::uint32_t r, std::uint32_t c) {
  if (!v.is_array()) {
    return v;
  }
  const ArrayValue* a = v.as_array();
  const std::uint32_t ri = a->rows == 1U ? 0U : r;
  const std::uint32_t ci = a->cols == 1U ? 0U : c;
  if (ri >= a->rows || ci >= a->cols) {
    return Value::error(ErrorCode::NA);
  }
  return a->cells[static_cast<std::size_t>(ri) * a->cols + ci];
}

// Shared body of PERCENTRANK.INC / .EXC: (array, x, [significance]), with
// an array `x` or `significance` evaluated per element.
Value percentrank_impl(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx, bool exclusive) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 2U || arity > 3U) {
    return Value::error(ErrorCode::Value);
  }
  auto arr = collect_rank_array(call.as_call_arg(0), arena, registry, ctx);
  if (std::holds_alternative<Value>(arr)) {
    return std::get<Value>(arr);
  }
  std::vector<double>& sorted = std::get<std::vector<double>>(arr);
  std::sort(sorted.begin(), sorted.end());
  const Value x_v = eval_node(call.as_call_arg(1), arena, registry, ctx);
  const Value sig_v = arity == 3U ? eval_node(call.as_call_arg(2), arena, registry, ctx) : Value::number(3.0);
  if (!x_v.is_array() && !sig_v.is_array()) {
    return percentrank_cell(sorted, x_v, sig_v, exclusive);
  }
  std::uint32_t rows = 1U;
  std::uint32_t cols = 1U;
  for (const Value* v : {&x_v, &sig_v}) {
    if (v->is_array()) {
      rows = std::max(rows, v->as_array()->rows);
      cols = std::max(cols, v->as_array()->cols);
    }
  }
  Value* cells = nullptr;
  ArrayValue* out = allocate_array_value(rows, cols, arena, cells, kMaxDerivedArrayCells);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  for (std::uint32_t r = 0; r < rows; ++r) {
    for (std::uint32_t c = 0; c < cols; ++c) {
      cells[static_cast<std::size_t>(r) * cols + c] =
          percentrank_cell(sorted, broadcast_at(x_v, r, c), broadcast_at(sig_v, r, c), exclusive);
    }
  }
  return rows == 1U && cols == 1U ? cells[0] : Value::array(out);
}

}  // namespace

Value eval_percentrank_inc_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                                const EvalContext& ctx) {
  return percentrank_impl(call, arena, registry, ctx, /*exclusive=*/false);
}

Value eval_percentrank_exc_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                                const EvalContext& ctx) {
  return percentrank_impl(call, arena, registry, ctx, /*exclusive=*/true);
}

}  // namespace eval
}  // namespace formulon
