// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the IRR lazy impl. See `eval/financial_lazy.h` for
// the dispatch-table contract and `eval/lazy_impls.h` for the shared
// `eval_node` / `LazyImpl` vocabulary.

#include "eval/financial_lazy.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
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

}  // namespace eval
}  // namespace formulon
