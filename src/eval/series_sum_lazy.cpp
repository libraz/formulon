//
// Implementation of `SERIESSUM(x, n, m, coefficients)`. Strict arity 4:
// three scalar numeric arguments followed by an array of coefficients.
// See `eval/series_sum_lazy.h` for the public contract and the
// design-space discussion.

#include "eval/series_sum_lazy.h"

#include <cmath>
#include <cstddef>
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

// Evaluates a scalar numeric argument. Returns `true` and writes the
// numeric value to `*out` on success; on failure writes the Excel error
// to `*out_err` and returns `false`. Bool / Text / Blank all resolve to
// `#VALUE!` - SERIESSUM's three scalar arguments are documented as
// numeric and Excel rejects non-numeric scalars.
bool eval_scalar_numeric(const parser::AstNode& arg_node, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx, double* out, Value* out_err) {
  const Value v = eval_node(arg_node, arena, registry, ctx);
  if (v.is_error()) {
    *out_err = v;
    return false;
  }
  if (!v.is_number()) {
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  *out = v.as_number();
  return true;
}

// Resolves the coefficients argument through the canonical
// `resolve_range_arg`, which covers `Ref` / `RangeOp` / `SpillRef` /
// `ArrayLiteral` and dynamic-array producers such as `SEQUENCE`. The
// rectangle shape is discarded because SERIESSUM treats the coefficient
// list as a flat 1-D sequence in row-major order, and a bare scalar is
// accepted as a 1x1 range: Excel treats `=SERIESSUM(2, 0, 1, 5)` as a
// single-coefficient series.
//
// Returns `true` on success and writes the flat cell vector to `*out`,
// with `*out_direct_scalar` recording whether the slot held a directly
// supplied scalar rather than a rectangle. On failure writes the Excel
// error to `*out_err` and returns `false`.
bool resolve_coefficients(const parser::AstNode& arg_node, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx, std::vector<Value>* out, bool* out_direct_scalar, Value* out_err) {
  auto resolved = resolve_range_arg(arg_node, arena, registry, ctx);
  if (!resolved) {
    *out_err = Value::error(resolved.error());
    return false;
  }
  *out_direct_scalar = resolved.value().from_scalar;
  *out = std::move(resolved.value().cells);
  return true;
}

}  // namespace

Value eval_series_sum_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx) {
  if (call.as_call_arity() != 4U) {
    return Value::error(ErrorCode::Value);
  }

  // Evaluate the scalars left-to-right so errors propagate in Excel's
  // documented argument order (x, n, m).
  double x = 0.0;
  double n = 0.0;
  double m = 0.0;
  Value err = Value::blank();
  if (!eval_scalar_numeric(call.as_call_arg(0), arena, registry, ctx, &x, &err)) {
    return err;
  }
  if (!eval_scalar_numeric(call.as_call_arg(1), arena, registry, ctx, &n, &err)) {
    return err;
  }
  if (!eval_scalar_numeric(call.as_call_arg(2), arena, registry, ctx, &m, &err)) {
    return err;
  }

  std::vector<Value> coefficients;
  bool direct_scalar = false;
  if (!resolve_coefficients(call.as_call_arg(3), arena, registry, ctx, &coefficients, &direct_scalar, &err)) {
    return err;
  }
  if (coefficients.empty()) {
    return Value::error(ErrorCode::Value);
  }

  // Any error cell in the coefficient range propagates. Scan in
  // row-major order so the left-most error wins.
  for (const Value& v : coefficients) {
    if (v.is_error()) {
      return v;
    }
  }

  // A directly supplied scalar coefficient (`=SERIESSUM(2, 0, 1, 5)`) is
  // coerced, so a non-numeric one is #VALUE! rather than a term silently
  // dropped from the series. Range-sourced cells keep the opposite rule
  // below. This is the same direct-scalar / range-cell split IRR applies
  // to its cash flows.
  if (direct_scalar) {
    auto coerced = coerce_to_number(coefficients.front());
    if (!coerced) {
      return Value::error(coerced.error());
    }
    coefficients.front() = Value::number(coerced.value());
  }

  // Accumulate Σᵢ coeff_i · x^(n + i·m) for i = 0..k-1. Range-sourced
  // cells that are not numbers (Blank, Bool, Text) are skipped, matching
  // Excel's tolerance of mixed columns; the power index still advances so
  // the i-th coefficient is always paired with the i-th term, matching
  // Excel's 1-based enumeration (first coefficient gets x^n).
  double total = 0.0;
  for (std::size_t i = 0; i < coefficients.size(); ++i) {
    const Value& v = coefficients[i];
    if (!v.is_number()) {
      continue;
    }
    const double power = n + static_cast<double>(i) * m;
    total += v.as_number() * std::pow(x, power);
  }
  if (std::isnan(total) || std::isinf(total)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(total);
}

}  // namespace eval
}  // namespace formulon
