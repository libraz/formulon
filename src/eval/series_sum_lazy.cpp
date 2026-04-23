// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of `SERIESSUM(x, n, m, coefficients)`. Strict arity 4:
// three scalar numeric arguments followed by an array of coefficients.
// See `eval/series_sum_lazy.h` for the public contract and the
// design-space discussion.

#include "eval/series_sum_lazy.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
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

// Resolves the coefficients argument. Accepts `Ref`, `RangeOp`, and
// `ArrayLiteral` in the same style as `regression_lazy.cpp`'s
// `resolve_array_arg`, but we do not need the rectangle shape because
// SERIESSUM treats the coefficient list as a flat 1-D sequence in
// row-major order.
//
// Returns `true` on success and writes the flat cell vector to `*out`.
// On failure writes the Excel error to `*out_err` and returns `false`.
bool resolve_coefficients(const parser::AstNode& arg_node, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx, std::vector<Value>* out, Value* out_err) {
  const parser::NodeKind k = arg_node.kind();
  if (k == parser::NodeKind::Ref || k == parser::NodeKind::RangeOp) {
    ErrorCode err_code = ErrorCode::Value;
    if (!resolve_range_arg(arg_node, arena, registry, ctx, out, &err_code)) {
      *out_err = Value::error(err_code);
      return false;
    }
    return true;
  }
  if (k == parser::NodeKind::ArrayLiteral) {
    const std::uint32_t rows = arg_node.as_array_rows();
    const std::uint32_t cols = arg_node.as_array_cols();
    const std::size_t total = static_cast<std::size_t>(rows) * cols;
    out->clear();
    out->reserve(total);
    for (std::uint32_t r = 0; r < rows; ++r) {
      for (std::uint32_t c = 0; c < cols; ++c) {
        Value v = eval_node(arg_node.as_array_element(r, c), arena, registry, ctx);
        out->push_back(v);
      }
    }
    return true;
  }
  // A bare scalar in the coefficients slot is acceptable: Excel treats
  // `=SERIESSUM(2, 0, 1, 5)` as a single-coefficient series. Evaluate the
  // subtree and wrap the result in a 1-element vector, propagating any
  // error.
  const Value v = eval_node(arg_node, arena, registry, ctx);
  if (v.is_error()) {
    *out_err = v;
    return false;
  }
  out->clear();
  out->push_back(v);
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
  if (!resolve_coefficients(call.as_call_arg(3), arena, registry, ctx, &coefficients, &err)) {
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

  // Accumulate Σᵢ coeff_i · x^(n + i·m) for i = 0..k-1. Non-numeric cells
  // (Blank, Bool, Text) are skipped; the power index still advances so
  // the i-th numeric coefficient is always paired with the i-th term,
  // matching Excel's 1-based enumeration (first coefficient gets x^n).
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
