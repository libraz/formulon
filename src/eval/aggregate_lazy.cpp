
#include "eval/aggregate_lazy.h"

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <iterator>
#include <utility>
#include <vector>

#include "eval/aggregate_kernels.h"
#include "eval/coerce.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "eval/name_env_resolve.h"
#include "eval/range_args.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// 1..13 are the SUBTOTAL-aligned modes; 14..19 are the AGGREGATE-only "k-
// arg" modes. Storing the integer code rather than an enum keeps the
// k-arg validation switch easy to read against the Excel docs.
constexpr int kCodeAverage = 1;
constexpr int kCodeCount = 2;
constexpr int kCodeCountA = 3;
constexpr int kCodeMax = 4;
constexpr int kCodeMin = 5;
constexpr int kCodeProduct = 6;
constexpr int kCodeStdevS = 7;
constexpr int kCodeStdevP = 8;
constexpr int kCodeSum = 9;
constexpr int kCodeVarS = 10;
constexpr int kCodeVarP = 11;
constexpr int kCodeMedian = 12;
constexpr int kCodeModeSngl = 13;
constexpr int kCodeLarge = 14;
constexpr int kCodeSmall = 15;
constexpr int kCodePercentileInc = 16;
constexpr int kCodeQuartileInc = 17;
constexpr int kCodePercentileExc = 18;
constexpr int kCodeQuartileExc = 19;

constexpr int kFnMin = 1;
constexpr int kFnMax = 19;
constexpr int kFnKArgFirst = 14;  // 14..19 take a trailing k arg.

// Reads a required scalar metadata argument (function_num / options /
// k). Errors propagate verbatim; non-coercible values surface
// `#VALUE!`; non-finite results (e.g. coercion overflow) surface
// `#NUM!`. The Expected return type avoids the previous in-band
// `0.0`-on-error sentinel, which collided with legitimate zero
// arguments (e.g. `AGGREGATE(2, 0, range)` -> COUNT mode + clear-flags
// option) and risked silent-wrong-result bugs.
Expected<double, ErrorCode> read_scalar(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                                        const EvalContext& ctx) {
  const Value v = eval_node(node, arena, registry, ctx);
  if (v.is_error()) {
    return v.as_error();
  }
  auto coerced = coerce_to_number(v);
  if (!coerced) {
    return coerced.error();
  }
  const double x = coerced.value();
  if (!std::isfinite(x)) {
    return ErrorCode::Num;
  }
  return x;
}

// Appends every scalar Value produced by `arg_node` to `out_cells`, mirroring
// PERCENTOF's `sum_arg_for_percentof` provenance walk. LET-bound NameRefs
// resolve to their bound AST when range-shaped. On any expansion failure
// (e.g. `#REF!` from a missing sheet) returns false with the propagating
// error in `*out_err`. Returns true on a clean walk.
//
// Unlike PERCENTOF this helper does NOT filter by Value kind: AGGREGATE's
// per-mode rules (numeric branches drop non-numerics; COUNTA counts them;
// the error-ignore bit decides whether errors short-circuit) are applied
// later by `apply_filters`.
bool collect_arg(const parser::AstNode& arg_node, Arena& arena, const FunctionRegistry& registry,
                 const EvalContext& ctx, std::vector<Value>* out_cells, Value* out_err) {
  const parser::AstNode* effective = &arg_node;
  if (arg_node.kind() == parser::NodeKind::NameRef) {
    const parser::AstNode& resolved = resolve_name_ast(arg_node, ctx.name_env());
    if (&resolved != &arg_node && is_range_shaped_ast(resolved)) {
      effective = &resolved;
    }
  }
  const parser::AstNode& node = *effective;
  const parser::NodeKind k = node.kind();

  // Range / Ref / SpillRef / RangeOp -> use the canonical resolver.
  if (k == parser::NodeKind::Ref || k == parser::NodeKind::RangeOp || k == parser::NodeKind::SpillRef) {
    auto resolved = resolve_range_arg(node, arena, registry, ctx);
    if (!resolved) {
      *out_err = Value::error(resolved.error());
      return false;
    }
    auto& rr = resolved.value();
    out_cells->insert(out_cells->end(), std::make_move_iterator(rr.cells.begin()),
                      std::make_move_iterator(rr.cells.end()));
    return true;
  }

  // Inline array literal `{a;b;c}` walked in row-major order.
  if (k == parser::NodeKind::ArrayLiteral) {
    const std::uint32_t rows = node.as_array_rows();
    const std::uint32_t cols = node.as_array_cols();
    for (std::uint32_t r = 0; r < rows; ++r) {
      for (std::uint32_t c = 0; c < cols; ++c) {
        const Value v = eval_node(node.as_array_element(r, c), arena, registry, ctx);
        out_cells->push_back(v);
      }
    }
    return true;
  }

  // Anything else (literal scalar, arithmetic expression, function call) is
  // evaluated normally. Array-valued calls are flattened in row-major order
  // so AGGREGATE's code-3 COUNTA path sees the same marker-bearing cells as
  // the eager COUNTA dispatcher; scalar values remain one direct argument.
  const Value v = eval_node(node, arena, registry, ctx);
  if (v.is_array()) {
    const ArrayValue* array = v.as_array();
    const std::size_t n = static_cast<std::size_t>(array->rows) * static_cast<std::size_t>(array->cols);
    out_cells->insert(out_cells->end(), array->cells, array->cells + n);
    return true;
  }
  out_cells->push_back(v);
  return true;
}

// Filters `cells` in place according to the options bit and the function
// code's expected provenance:
//
//   * Errors: dropped silently when ignore_errors == true; the first error
//     short-circuits the call and is written to `*out_err` otherwise.
//   * COUNTA (code 3): keep every non-blank, plus blanks owned by a derived
//     value array; raw-reference blanks are dropped.
//   * Numeric modes (everything else): keep only Numbers; drop Bool / Text /
//     Blank.
//
// Returns true on a clean filter, false (with `*out_err` populated) when an
// un-ignored error short-circuits.
bool apply_filters(std::vector<Value>* cells, int code, bool ignore_errors, Value* out_err) {
  std::vector<Value> kept;
  kept.reserve(cells->size());
  for (const Value& v : *cells) {
    if (v.is_error()) {
      if (ignore_errors) {
        continue;
      }
      *out_err = v;
      return false;
    }
    if (code == kCodeCountA) {
      if (!v.is_blank() || v.blank_counts_for_counta()) {
        kept.push_back(v);
      }
      continue;
    }
    // Numeric branches: drop everything that is not a Number.
    if (v.is_number()) {
      kept.push_back(v);
    }
  }
  *cells = std::move(kept);
  return true;
}

// Helper: extract the numeric slice once we know every kept cell is a
// Number (true for codes 1, 2, 4..19).
std::vector<double> to_numbers(const std::vector<Value>& cells) {
  std::vector<double> out;
  out.reserve(cells.size());
  for (const Value& v : cells) {
    if (v.is_number()) {
      out.push_back(v.as_number());
    }
  }
  return out;
}

// ---------------------------------------------------------------------------
// Mode runners (codes 1..13). The numeric-aggregator slots (SUM / PRODUCT /
// MIN / MAX / AVERAGE / VAR.* / STDEV.*) all delegate to the shared kernels
// in `aggregate_kernels.h` so SUBTOTAL and AGGREGATE cannot drift. Empty-
// range behaviour matches Excel's convention for SUBTOTAL / AGGREGATE:
// SUM/PRODUCT/MIN/MAX -> 0, AVERAGE/VAR/STDEV -> #DIV/0!, MEDIAN -> #DIV/0!,
// COUNT/COUNTA -> 0, MODE.SNGL -> #N/A.

// Lifts an `Expected<double, ErrorCode>` kernel result into the `Value`
// shape AGGREGATE's dispatcher expects.
Value lift_kernel_result(Expected<double, ErrorCode> result) {
  if (!result) {
    return Value::error(result.error());
  }
  return Value::number(result.value());
}

Value run_count(const std::vector<Value>& cells) {
  // After `apply_filters`, numeric branches retain only Numbers; this gives
  // the same answer as iterating the post-filter `cells` directly.
  std::uint32_t n = 0;
  for (const Value& v : cells) {
    if (v.is_number()) {
      ++n;
    }
  }
  return Value::number(static_cast<double>(n));
}

Value run_counta(const std::vector<Value>& cells) {
  // The COUNTA branch of `apply_filters` already dropped plain and
  // raw-reference Blanks; everything remaining contributes 1.
  return Value::number(static_cast<double>(cells.size()));
}

Value run_median(std::vector<double> xs) {
  if (xs.empty()) {
    return Value::error(ErrorCode::Div0);
  }
  std::sort(xs.begin(), xs.end());
  const std::size_t n = xs.size();
  const double m = (n % 2 == 1U) ? xs[n / 2] : 0.5 * (xs[n / 2 - 1U] + xs[n / 2]);
  if (!std::isfinite(m)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(m);
}

// LARGE / SMALL — k must be a positive integer in [1, n]. k is truncated.
Value run_large_small(std::vector<double> xs, double k_raw, bool want_large) {
  if (xs.empty()) {
    return Value::error(ErrorCode::Num);
  }
  const double k_trunc = std::trunc(k_raw);
  if (!std::isfinite(k_trunc) || k_trunc < 1.0 || k_trunc > static_cast<double>(xs.size())) {
    return Value::error(ErrorCode::Num);
  }
  const auto k = static_cast<std::size_t>(k_trunc);
  std::sort(xs.begin(), xs.end());
  // LARGE: k-th largest = xs[n - k]. SMALL: k-th smallest = xs[k - 1].
  const double picked = want_large ? xs[xs.size() - k] : xs[k - 1];
  if (!std::isfinite(picked)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(picked);
}

// PERCENTILE.INC. Domain / position-formula logic lives in
// `aggregate_kernels::percentile_sorted_inc`; this wrapper sorts in place
// and lifts the kernel result to `Value`.
Value run_percentile_inc(std::vector<double> xs, double p) {
  std::sort(xs.begin(), xs.end());
  return lift_kernel_result(aggregate_kernels::percentile_sorted_inc(xs, p));
}

// PERCENTILE.EXC. Domain / position-formula logic lives in
// `aggregate_kernels::percentile_sorted_exc`; this wrapper sorts in place
// and lifts the kernel result to `Value`.
Value run_percentile_exc(std::vector<double> xs, double p) {
  std::sort(xs.begin(), xs.end());
  return lift_kernel_result(aggregate_kernels::percentile_sorted_exc(xs, p));
}

// QUARTILE.INC delegates to PERCENTILE.INC at p ∈ {0, 0.25, 0.5, 0.75, 1.0}.
// `quart` must be an integer in [0, 4]; truncated like the rest.
Value run_quartile_inc(std::vector<double> xs, double quart_raw) {
  const double q_trunc = std::trunc(quart_raw);
  if (!std::isfinite(q_trunc) || q_trunc < 0.0 || q_trunc > 4.0) {
    return Value::error(ErrorCode::Num);
  }
  static constexpr double kProb[] = {0.0, 0.25, 0.5, 0.75, 1.0};
  const auto idx = static_cast<std::size_t>(q_trunc);
  return run_percentile_inc(std::move(xs), kProb[idx]);
}

// QUARTILE.EXC delegates to PERCENTILE.EXC at p ∈ {0.25, 0.5, 0.75}.
// `quart` must be an integer in {1, 2, 3}; 0 and 4 are rejected.
Value run_quartile_exc(std::vector<double> xs, double quart_raw) {
  const double q_trunc = std::trunc(quart_raw);
  if (!std::isfinite(q_trunc) || q_trunc < 1.0 || q_trunc > 3.0) {
    return Value::error(ErrorCode::Num);
  }
  static constexpr double kProb[] = {0.25, 0.5, 0.75};
  const auto idx = static_cast<std::size_t>(q_trunc) - 1U;
  return run_percentile_exc(std::move(xs), kProb[idx]);
}

}  // namespace

Value eval_aggregate_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  // function_num + options + at least one data arg.
  if (arity < 3U) {
    return Value::error(ErrorCode::Value);
  }

  Value err = Value::blank();
  auto fn_raw_or = read_scalar(call.as_call_arg(0), arena, registry, ctx);
  if (!fn_raw_or) {
    return Value::error(fn_raw_or.error());
  }
  const int code = static_cast<int>(std::trunc(fn_raw_or.value()));
  if (code < kFnMin || code > kFnMax) {
    return Value::error(ErrorCode::Value);
  }

  auto opts_raw_or = read_scalar(call.as_call_arg(1), arena, registry, ctx);
  if (!opts_raw_or) {
    return Value::error(opts_raw_or.error());
  }
  const int options = static_cast<int>(std::trunc(opts_raw_or.value()));
  if (options < 0 || options > 7) {
    return Value::error(ErrorCode::Value);
  }
  // Bit 1 (mask 2) of the options byte is the error-ignore bit. The other
  // bits (hidden rows, nested SUBTOTAL/AGGREGATE) are not yet observable;
  // see the file header.
  const bool ignore_errors = (options & 2) != 0;

  std::vector<Value> cells;

  if (code >= kFnKArgFirst) {
    // 14..19 — Excel requires exactly one data range plus a trailing k.
    // Anything other than `(fn, options, data, k)` -> #VALUE!.
    if (arity != 4U) {
      return Value::error(ErrorCode::Value);
    }
    if (!collect_arg(call.as_call_arg(2), arena, registry, ctx, &cells, &err)) {
      return err;
    }
    if (!apply_filters(&cells, code, ignore_errors, &err)) {
      return err;
    }
    // k is a scalar metadata arg: errors propagate regardless of the
    // options bit (matches the function_num / options contract).
    auto k_raw_or = read_scalar(call.as_call_arg(3), arena, registry, ctx);
    if (!k_raw_or) {
      return Value::error(k_raw_or.error());
    }
    const double k_raw = k_raw_or.value();
    std::vector<double> xs = to_numbers(cells);
    switch (code) {
      case kCodeLarge:
        return run_large_small(std::move(xs), k_raw, /*want_large=*/true);
      case kCodeSmall:
        return run_large_small(std::move(xs), k_raw, /*want_large=*/false);
      case kCodePercentileInc:
        return run_percentile_inc(std::move(xs), k_raw);
      case kCodeQuartileInc:
        return run_quartile_inc(std::move(xs), k_raw);
      case kCodePercentileExc:
        return run_percentile_exc(std::move(xs), k_raw);
      case kCodeQuartileExc:
        return run_quartile_exc(std::move(xs), k_raw);
      default:
        // Unreachable: code is constrained to [14, 19] in this branch.
        return Value::error(ErrorCode::Value);
    }
  }

  // Codes 1..13 — every remaining positional arg is data.
  for (std::uint32_t i = 2; i < arity; ++i) {
    if (!collect_arg(call.as_call_arg(i), arena, registry, ctx, &cells, &err)) {
      return err;
    }
  }
  if (!apply_filters(&cells, code, ignore_errors, &err)) {
    return err;
  }

  switch (code) {
    case kCodeAverage:
      return lift_kernel_result(aggregate_kernels::run_average(to_numbers(cells)));
    case kCodeCount:
      return run_count(cells);
    case kCodeCountA:
      return run_counta(cells);
    case kCodeMax:
      return lift_kernel_result(aggregate_kernels::run_max(to_numbers(cells)));
    case kCodeMin:
      return lift_kernel_result(aggregate_kernels::run_min(to_numbers(cells)));
    case kCodeProduct:
      return lift_kernel_result(aggregate_kernels::run_product(to_numbers(cells)));
    case kCodeStdevS:
      return lift_kernel_result(aggregate_kernels::run_stdev(to_numbers(cells), /*sample=*/true));
    case kCodeStdevP:
      return lift_kernel_result(aggregate_kernels::run_stdev(to_numbers(cells), /*sample=*/false));
    case kCodeSum:
      return lift_kernel_result(aggregate_kernels::run_sum(to_numbers(cells)));
    case kCodeVarS:
      return lift_kernel_result(aggregate_kernels::run_variance(to_numbers(cells), /*sample=*/true));
    case kCodeVarP:
      return lift_kernel_result(aggregate_kernels::run_variance(to_numbers(cells), /*sample=*/false));
    case kCodeMedian:
      return run_median(to_numbers(cells));
    case kCodeModeSngl:
      // First-occurrence tie-break (Excel MODE.SNGL): the shared kernel
      // consumes the cells in input order, so do NOT sort first.
      return lift_kernel_result(aggregate_kernels::mode_first_occurrence(to_numbers(cells)));
    default:
      // Unreachable: code is constrained to [1, 13] in this branch.
      return Value::error(ErrorCode::Value);
  }
}

}  // namespace eval
}  // namespace formulon
