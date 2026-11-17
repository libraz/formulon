//
// Implementation of `SERIESSUM(x, n, m, coefficients)`. Strict arity 4:
// three scalar numeric arguments followed by an array of coefficients.
// See `eval/series_sum_lazy.h` for the public contract and the
// design-space discussion.

#include "eval/series_sum_lazy.h"

#include <cstddef>
#include <utility>
#include <vector>

#include "eval/builtins/numeric_helpers.h"
#include "eval/coerce.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "eval/omitted_arg.h"
#include "eval/range_args.h"
#include "eval/scalar_args.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/subnormal_flush.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

using builtins_detail::to_finite_value;

// Reads one of the three scalar arguments: a blank cell is 0, text coerces
// as a number ("2" is 2), and a Boolean is #VALUE!. Returns false with the
// error in `*out_err`.
bool read_scalar_numeric(const Value& v, double* out, Value* out_err) {
  if (v.is_boolean()) {
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  return number_from_value(v, *out, *out_err);
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

  // Evaluate the scalars left to right so errors answer in argument order
  // (x, n, m), ahead of the coefficients'. An omitted slot is #N/A under the
  // Analysis-ToolPak rule.
  Value scalars[3] = {Value::blank(), Value::blank(), Value::blank()};
  for (std::uint32_t i = 0; i < 3U; ++i) {
    const parser::AstNode& arg = call.as_call_arg(i);
    scalars[i] = eval_node(arg, arena, registry, ctx);
    if (scalars[i].is_error()) {
      return scalars[i];
    }
    if (const Value atp = atp_arg_error(arg, scalars[i], /*required=*/true); atp.is_error()) {
      return atp;
    }
    double unused = 0.0;
    if (!scalars[i].is_array() && !read_scalar_numeric(scalars[i], &unused, &scalars[i])) {
      return scalars[i];
    }
  }

  Value err = Value::blank();
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
  if (const Value* error_cell = first_error(coefficients)) {
    return *error_cell;
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

  // x, n and m lift element by element.
  return lift_over_array_args(scalars, 3U, arena, [&](const Value* args, std::uint32_t) {
    double xnm[3] = {0.0, 0.0, 0.0};
    Value scalar_err = Value::blank();
    for (std::uint32_t i = 0; i < 3U; ++i) {
      if (!read_scalar_numeric(args[i], &xnm[i], &scalar_err)) {
        return scalar_err;
      }
    }
    const double x = xnm[0];
    const double n = xnm[1];
    const double m = xnm[2];
    // Accumulate sum(coeff_i * x^e_i) with e_0 = n and e_{i+1} = e_i + m, so the
    // exponent advances by repeated addition. Range-sourced cells that are not
    // numbers (Blank, Bool, Text) are skipped, but the exponent still advances.
    // A negative base with a non-integral exponent contributes 0 to the sum.
    // Otherwise each power follows the `^` operator: 0^0 is #NUM! and 0^negative
    // is #DIV/0!.
    double total = 0.0;
    double e = n;
    for (const Value& v : coefficients) {
      const double power = e;
      e += m;
      if (!v.is_number()) {
        continue;
      }
      double term_value = 0.0;
      if (!(x < 0.0 && (power < 0.0 ? std::ceil(power) : std::floor(power)) != power)) {
        auto term = apply_pow(x, power);
        if (!term) {
          return Value::error(term.error());
        }
        term_value = term.value();
      }
      total += flush_subnormal(v.as_number() * term_value);
    }
    return to_finite_value(total);
  });
}

}  // namespace eval
}  // namespace formulon
