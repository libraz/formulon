// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the tree-walk evaluator. See `tree_walker.h` for the
// public contract and the design references.

#include "eval/tree_walker.h"

#include <cmath>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "eval/coerce.h"
#include "eval/conditional_aggregates.h"
#include "eval/eval_context.h"
#include "eval/financial_lazy.h"
#include "eval/function_registry.h"
#include "eval/lazy_impls.h"
#include "eval/lookups/classic.h"
#include "eval/lookups/xlookup.h"
#include "eval/range_args.h"
#include "eval/regression_lazy.h"
#include "eval/shape_ops_lazy.h"
#include "eval/special_forms_lazy.h"
#include "eval/workdays_lazy.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/expected.h"  // FM_CHECK
#include "utils/strings.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// ---------------------------------------------------------------------------
// Cross-type comparison
// ---------------------------------------------------------------------------

// Excel cross-type comparison order: Number < Text < Bool. Blank coerces to
// numeric zero. Text equality and ordering are case-insensitive over ASCII
// letters; locale-aware comparison is deferred. NaN compares as "unordered":
// every relational operator returns FALSE except `<>`.
//
// `out_unordered` is set to true iff one of the operands is NaN; the caller
// uses it to short-circuit relational operators to FALSE while still
// returning TRUE for `<>`. The integer return value (-1/0/+1) is meaningful
// only when `out_unordered` is false.
int compare_values(const Value& lhs, const Value& rhs, bool* out_unordered) {
  *out_unordered = false;

  auto rank = [](ValueKind k) -> int {
    // Blank is treated as numeric (0) for comparison purposes.
    switch (k) {
      case ValueKind::Number:
      case ValueKind::Blank:
        return 0;
      case ValueKind::Text:
        return 1;
      case ValueKind::Bool:
        return 2;
      default:
        return 3;
    }
  };

  const int lr = rank(lhs.kind());
  const int rr = rank(rhs.kind());
  if (lr != rr) {
    return lr < rr ? -1 : 1;
  }

  switch (lr) {
    case 0: {
      const double a = lhs.is_blank() ? 0.0 : lhs.as_number();
      const double b = rhs.is_blank() ? 0.0 : rhs.as_number();
      if (std::isnan(a) || std::isnan(b)) {
        *out_unordered = true;
        return 0;
      }
      if (a < b) {
        return -1;
      }
      if (a > b) {
        return 1;
      }
      return 0;
    }
    case 1: {
      const std::string_view a = lhs.as_text();
      const std::string_view b = rhs.as_text();
      const std::size_t n = a.size() < b.size() ? a.size() : b.size();
      for (std::size_t i = 0; i < n; ++i) {
        const char ca = strings::ascii_to_lower(a[i]);
        const char cb = strings::ascii_to_lower(b[i]);
        if (ca != cb) {
          return ca < cb ? -1 : 1;
        }
      }
      if (a.size() != b.size()) {
        return a.size() < b.size() ? -1 : 1;
      }
      return 0;
    }
    case 2: {
      const bool a = lhs.as_boolean();
      const bool b = rhs.as_boolean();
      if (a == b) {
        return 0;
      }
      // FALSE < TRUE.
      return a ? 1 : -1;
    }
    default:
      return 0;
  }
}

// ---------------------------------------------------------------------------
// Per-operator helpers
// ---------------------------------------------------------------------------

Value finalize_arithmetic(double r) {
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

Value apply_unary(parser::UnaryOp op, const Value& operand) {
  if (operand.is_error()) {
    return operand;
  }
  auto coerced = coerce_to_number(operand);
  if (!coerced) {
    return Value::error(coerced.error());
  }
  const double x = coerced.value();
  switch (op) {
    case parser::UnaryOp::Plus:
      return finalize_arithmetic(x);
    case parser::UnaryOp::Minus:
      return finalize_arithmetic(-x);
    case parser::UnaryOp::Percent:
      return finalize_arithmetic(x / 100.0);
  }
  return Value::error(ErrorCode::Value);
}

Value apply_arithmetic(parser::BinOp op, double lhs, double rhs) {
  switch (op) {
    case parser::BinOp::Add:
      return finalize_arithmetic(lhs + rhs);
    case parser::BinOp::Sub:
      return finalize_arithmetic(lhs - rhs);
    case parser::BinOp::Mul:
      return finalize_arithmetic(lhs * rhs);
    case parser::BinOp::Div:
      // Excel reports #DIV/0! for any division whose divisor is exactly
      // zero, including 0/0 (no #NUM! tie-break).
      if (rhs == 0.0) {
        return Value::error(ErrorCode::Div0);
      }
      return finalize_arithmetic(lhs / rhs);
    case parser::BinOp::Pow: {
      // Delegates to the shared `apply_pow` helper so the `^` operator and
      // the `POWER()` builtin cannot drift apart on edge cases.
      auto r = apply_pow(lhs, rhs);
      if (!r) {
        return Value::error(r.error());
      }
      return Value::number(r.value());
    }
    default:
      // Caller guarantees op is arithmetic.
      FM_CHECK(false, "apply_arithmetic called with non-arithmetic op");
      return Value::error(ErrorCode::Value);
  }
}

Value apply_concat(const Value& lhs, const Value& rhs, Arena& arena) {
  auto lhs_text = coerce_to_text(lhs);
  if (!lhs_text) {
    return Value::error(lhs_text.error());
  }
  auto rhs_text = coerce_to_text(rhs);
  if (!rhs_text) {
    return Value::error(rhs_text.error());
  }
  std::string joined;
  joined.reserve(lhs_text.value().size() + rhs_text.value().size());
  joined.append(lhs_text.value());
  joined.append(rhs_text.value());
  const std::string_view interned = arena.intern(joined);
  // Empty input is fine: Arena::intern returns an empty view that is still
  // a valid Text payload.
  return Value::text(interned);
}

Value apply_comparison(parser::BinOp op, const Value& lhs, const Value& rhs) {
  bool unordered = false;
  const int cmp = compare_values(lhs, rhs, &unordered);
  switch (op) {
    case parser::BinOp::Eq:
      return Value::boolean(!unordered && cmp == 0);
    case parser::BinOp::NotEq:
      // NaN != anything is TRUE, matching IEEE-754 semantics.
      return Value::boolean(unordered || cmp != 0);
    case parser::BinOp::Lt:
      return Value::boolean(!unordered && cmp < 0);
    case parser::BinOp::LtEq:
      return Value::boolean(!unordered && cmp <= 0);
    case parser::BinOp::Gt:
      return Value::boolean(!unordered && cmp > 0);
    case parser::BinOp::GtEq:
      return Value::boolean(!unordered && cmp >= 0);
    default:
      FM_CHECK(false, "apply_comparison called with non-comparison op");
      return Value::error(ErrorCode::Value);
  }
}

// ---------------------------------------------------------------------------
// Recursive evaluator
// ---------------------------------------------------------------------------
//
// `eval_node` is declared in `eval/lazy_impls.h` with external linkage so
// lazy-impl translation units (e.g. `special_forms_lazy.cpp`) can reach
// it. Its definition lives at the bottom of this file, outside the
// anonymous namespace.

// ---------------------------------------------------------------------------
// Lazy (short-circuit) function impls
// ---------------------------------------------------------------------------
//
// Each lazy impl receives the full `Call` AST node so it can pull arguments
// out by index and decide which subtrees to evaluate. The eager path in
// `dispatch_call` is bypassed entirely: arity checks and error propagation
// belong inside each impl. On arity mismatch the impls return #VALUE! to
// match the eager dispatcher's behaviour.
//
// Current entries:
//   IF          - short-circuit branch: only the taken side is evaluated.
//   IFERROR     - evaluates fallback only when primary is any error.
//   IFNA        - evaluates fallback only when primary is exactly #N/A.
//   COUNTIF     - range-aware: arg 0 must be a range/Ref, arg 1 is a scalar
//                 criterion evaluated once; counts matching cells.
//   SUMIF       - range-aware: arg 0 is the criteria range, arg 2 (optional)
//                 is the parallel sum range; sums matching numeric cells.
//   AVERAGEIF   - like SUMIF, but returns the mean of matching numeric
//                 cells or #DIV/0! when nothing matches.
//   COUNTIFS    - multi-criteria AND across N (range, criterion) pairs.
//   SUMIFS      - like COUNTIFS, but with a result range as the leading arg.
//   AVERAGEIFS  - like SUMIFS, returns mean or #DIV/0! when no matches.
//   MAXIFS      - like SUMIFS, returns max of numerics (0 if no matches).
//   MINIFS      - like SUMIFS, returns min of numerics (0 if no matches).
//   CHOOSE      - index-selected argument; only the chosen subtree runs.
//   INDEX       - range-aware: shape (rows,cols) of arg 0 is used to pick
//                 a single cell by (row_num, col_num).
//   MATCH       - range-aware: lookup_array (arg 1) must be a 1-D range/Ref.
//
// The conditional aggregators (`*IF`/`*IFS`) cannot ride on the eager
// `accepts_ranges` path because arg 0 must reach the impl as AST (so a
// bare single-cell Ref can be treated as a 1-cell range) AND the parallel
// result / additional criteria ranges must iterate in lockstep rather than
// being flattened into a single values vector alongside the first.

// The lazy impls themselves live in per-family translation units:
//   IF / IFERROR / IFNA                        -> src/eval/special_forms_lazy.cpp
//   COUNTIF / SUMIF / AVERAGEIF / *IFS         -> src/eval/conditional_aggregates.cpp
//   CHOOSE / INDEX / MATCH / VLOOKUP / HLOOKUP -> src/eval/lookups/classic.cpp
//   XLOOKUP / XMATCH                           -> src/eval/lookups/xlookup.cpp
//   ROWS / COLUMNS / ROW / COLUMN / SUMPRODUCT -> src/eval/shape_ops_lazy.cpp
//   NETWORKDAYS / WORKDAY                      -> src/eval/workdays_lazy.cpp
//   CORREL / COVARIANCE.P / COVARIANCE.S /
//   SLOPE / INTERCEPT / RSQ / FORECAST.LINEAR  -> src/eval/regression_lazy.cpp
// Each family publishes its externs via its own header
// (`eval/special_forms_lazy.h`, `eval/conditional_aggregates.h`,
// `eval/lookups/classic.h`, `eval/lookups/xlookup.h`,
// `eval/shape_ops_lazy.h`, `eval/workdays_lazy.h`), which the dispatch
// table below includes.

// `LazyImpl` is declared in `eval/lazy_impls.h` so translation units that
// own individual lazy impls can publish matching function pointers.
struct LazyEntry {
  const char* name;  // canonical UPPERCASE
  LazyImpl impl;
};

constexpr LazyEntry kLazyDispatch[] = {
    {"AVERAGEIF", &eval_averageif_lazy},
    {"AVERAGEIFS", &eval_averageifs_lazy},
    {"CHOOSE", &eval_choose_lazy},
    {"COLUMN", &eval_column_lazy},
    {"COLUMNS", &eval_columns_lazy},
    {"CORREL", &eval_correl_lazy},
    {"COUNT", &eval_count_lazy},
    {"COUNTIF", &eval_countif_lazy},
    {"COUNTIFS", &eval_countifs_lazy},
    {"COVARIANCE.P", &eval_covariance_p_lazy},
    {"COVARIANCE.S", &eval_covariance_s_lazy},
    // FORECAST is the legacy spelling kept by Excel for back-compat;
    // its impl and arity are identical to FORECAST.LINEAR.
    {"FORECAST", &eval_forecast_linear_lazy},
    {"FORECAST.LINEAR", &eval_forecast_linear_lazy},
    {"HLOOKUP", &eval_hlookup_lazy},
    {"IF", &eval_if_lazy},
    {"IFERROR", &eval_iferror_lazy},
    {"IFNA", &eval_ifna_lazy},
    {"IFS", &eval_ifs_lazy},
    {"INDEX", &eval_index_lazy},
    {"INTERCEPT", &eval_intercept_lazy},
    {"IRR", &eval_irr_lazy},
    {"MATCH", &eval_match_lazy},
    {"MAXIFS", &eval_maxifs_lazy},
    {"MINIFS", &eval_minifs_lazy},
    {"NETWORKDAYS", &eval_networkdays_lazy},
    {"ROW", &eval_row_lazy},
    {"ROWS", &eval_rows_lazy},
    {"RSQ", &eval_rsq_lazy},
    {"SLOPE", &eval_slope_lazy},
    {"SUMIF", &eval_sumif_lazy},
    {"SUMIFS", &eval_sumifs_lazy},
    {"SUMPRODUCT", &eval_sumproduct_lazy},
    {"SWITCH", &eval_switch_lazy},
    {"VLOOKUP", &eval_vlookup_lazy},
    {"WORKDAY", &eval_workday_lazy},
    {"XLOOKUP", &eval_xlookup_lazy},
    {"XMATCH", &eval_xmatch_lazy},
};

const LazyEntry* find_lazy(std::string_view name) noexcept {
  for (const auto& e : kLazyDispatch) {
    if (strings::case_insensitive_eq(name, std::string_view(e.name))) {
      return &e;
    }
  }
  return nullptr;
}

// Special-cased function-call dispatch.
//
// Lazy entries (`IF`, `IFERROR`, `IFNA`, the `*IF`/`*IFS` aggregators) are
// routed through the table above;
// each impl owns its own arity check and chooses which subtrees to evaluate.
//
// All other names are routed through `registry`: unknown name -> #NAME?,
// arity violation -> #VALUE!, otherwise every argument is pre-evaluated in
// order. By default the left-most error short-circuits before the impl
// runs, but an entry whose `propagate_errors` flag is `false` (the IS*
// type-predicate family) opts out of that short-circuit and receives raw
// error values among its arguments.
Value dispatch_call(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                    const EvalContext& ctx) {
  const std::string_view name = node.as_call_name();
  const std::uint32_t arity = node.as_call_arity();

  if (const LazyEntry* lazy = find_lazy(name); lazy != nullptr) {
    return lazy->impl(node, arena, registry, ctx);
  }

  const FunctionDef* def = registry.lookup(name);
  if (def == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  // The pre-expansion arity guards min_arity / max_arity. This happens to
  // align with Excel's behaviour for the range-aware aggregators:
  // `=SUM()` is rejected at parse time, and `=SUM(A1:A1)` passes the
  // `min_arity = 1` check even though its expansion might be empty (which
  // cannot happen with a finite valid rectangle today).
  if (arity < def->min_arity || arity > def->max_arity) {
    return Value::error(ErrorCode::Value);
  }

  // Pre-evaluate arguments left-to-right. By default the first error wins
  // and the impl is never invoked; functions that need to inspect error
  // arguments (e.g. `ISERROR`) clear `propagate_errors` to opt out. When
  // the function is range-aware (`accepts_ranges`), any argument whose AST
  // node is a simple RangeOp (Ref:Ref) is flattened into the values vector
  // in row-major order.
  std::vector<Value> values;
  values.reserve(arity);
  for (std::uint32_t i = 0; i < arity; ++i) {
    const parser::AstNode& arg_node = node.as_call_arg(i);
    // Minimal array-literal support: when a range-aware function receives
    // a `{a;b;c}` style literal, flatten it in row-major order exactly like
    // a RangeOp argument. This is just enough to let LARGE / SMALL /
    // PERCENTILE.INC / QUARTILE.INC accept brace literals as their "array"
    // input; full array-aware evaluation (broadcasting, spilled output)
    // stays deferred — a bare `={1;2;3}` outside a function call still
    // surfaces as #VALUE! via `eval_node`'s ArrayLiteral case.
    if (def->accepts_ranges && arg_node.kind() == parser::NodeKind::ArrayLiteral) {
      const std::uint32_t rows = arg_node.as_array_rows();
      const std::uint32_t cols = arg_node.as_array_cols();
      bool short_circuit = false;
      Value propagated_err = Value::blank();
      for (std::uint32_t r = 0; r < rows && !short_circuit; ++r) {
        for (std::uint32_t c = 0; c < cols; ++c) {
          Value v = eval_node(arg_node.as_array_element(r, c), arena, registry, ctx);
          if (def->propagate_errors && v.is_error()) {
            propagated_err = v;
            short_circuit = true;
            break;
          }
          values.push_back(v);
        }
      }
      if (short_circuit) {
        return propagated_err;
      }
      continue;
    }
    if (def->accepts_ranges && arg_node.kind() == parser::NodeKind::RangeOp) {
      const parser::AstNode& lhs_ast = arg_node.as_range_lhs();
      const parser::AstNode& rhs_ast = arg_node.as_range_rhs();
      // Only the simplest form -- literal A1:B2 where both endpoints are
      // Refs -- is expanded. Anything else (INDIRECT(...), named ranges,
      // etc.) surfaces as #REF! here; full dynamic range resolution is
      // deferred.
      if (lhs_ast.kind() != parser::NodeKind::Ref || rhs_ast.kind() != parser::NodeKind::Ref) {
        const Value err = Value::error(ErrorCode::Ref);
        if (def->propagate_errors) {
          return err;
        }
        values.push_back(err);
        continue;
      }
      auto expanded = ctx.expand_range(lhs_ast.as_ref(), rhs_ast.as_ref(), arena, registry);
      if (!expanded) {
        const Value err = Value::error(expanded.error());
        if (def->propagate_errors) {
          return err;
        }
        values.push_back(err);
        continue;
      }
      for (const Value& v : expanded.value()) {
        if (def->propagate_errors && v.is_error()) {
          return v;
        }
        // Provenance-aware filtering for range-sourced values. Excel skips
        // Bool / Text / Blank cells inside a range for SUM / AVERAGE /
        // MIN / MAX / PRODUCT, and skips Text / Blank for AND / OR.
        // Direct scalar arguments (handled in the `else` branch below) are
        // not affected by either filter.
        if (def->range_filter_numeric_only && v.kind() != ValueKind::Number) {
          continue;
        }
        if (def->range_filter_bool_coercible && v.kind() != ValueKind::Number && v.kind() != ValueKind::Bool) {
          continue;
        }
        values.push_back(v);
      }
      continue;
    }
    Value v = eval_node(arg_node, arena, registry, ctx);
    if (def->propagate_errors && v.is_error()) {
      return v;
    }
    values.push_back(v);
  }
  // Hand the post-expansion size to the impl; aggregator bodies walk the
  // flattened vector directly.
  return def->impl(values.data(), static_cast<std::uint32_t>(values.size()), arena);
}

}  // namespace

// Defined with external linkage (declared in `eval/lazy_impls.h`) so the
// per-family lazy-impl TUs can recurse into the evaluator. The helpers it
// calls below — `apply_unary`, `apply_arithmetic`, `apply_concat`,
// `apply_comparison`, `dispatch_call` — live in the anonymous namespace
// above and remain reachable via ordinary unqualified lookup because that
// anonymous namespace is nested inside `formulon::eval`.
Value eval_node(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx) {
  switch (node.kind()) {
    case parser::NodeKind::Literal:
      return node.as_literal();

    case parser::NodeKind::ErrorLiteral:
      return Value::error(node.as_error_literal());

    case parser::NodeKind::ErrorPlaceholder:
      // Panic-mode skipped this subtree at parse time; we cannot do better
      // than #NAME? since the original tokens are unavailable.
      return Value::error(ErrorCode::Name);

    case parser::NodeKind::ImplicitIntersection:
      // Identity for scalars. Once arrays land this becomes the contraction
      // operator (1x1 selection from a column / row at the call site).
      return eval_node(node.as_implicit_intersection_operand(), arena, registry, ctx);

    case parser::NodeKind::UnaryOp:
      return apply_unary(node.as_unary_op(), eval_node(node.as_unary_operand(), arena, registry, ctx));

    case parser::NodeKind::BinaryOp: {
      const parser::BinOp op = node.as_binary_op();
      // Evaluate left first so error propagation honours the documented
      // left-most-wins rule from backup/plans/02-calc-engine.md §2.1.1.
      const Value lhs = eval_node(node.as_binary_lhs(), arena, registry, ctx);
      if (lhs.is_error()) {
        return lhs;
      }
      const Value rhs = eval_node(node.as_binary_rhs(), arena, registry, ctx);
      if (rhs.is_error()) {
        return rhs;
      }

      switch (op) {
        case parser::BinOp::Add:
        case parser::BinOp::Sub:
        case parser::BinOp::Mul:
        case parser::BinOp::Div:
        case parser::BinOp::Pow: {
          auto lhs_n = coerce_to_number(lhs);
          if (!lhs_n) {
            return Value::error(lhs_n.error());
          }
          auto rhs_n = coerce_to_number(rhs);
          if (!rhs_n) {
            return Value::error(rhs_n.error());
          }
          return apply_arithmetic(op, lhs_n.value(), rhs_n.value());
        }
        case parser::BinOp::Concat:
          return apply_concat(lhs, rhs, arena);
        case parser::BinOp::Eq:
        case parser::BinOp::NotEq:
        case parser::BinOp::Lt:
        case parser::BinOp::LtEq:
        case parser::BinOp::Gt:
        case parser::BinOp::GtEq:
          return apply_comparison(op, lhs, rhs);
      }
      return Value::error(ErrorCode::Value);
    }

    case parser::NodeKind::Call:
      return dispatch_call(node, arena, registry, ctx);

    case parser::NodeKind::Ref:
      return ctx.resolve_ref(node.as_ref(), arena, registry);

    // -- Unsupported: name resolution / closures --------------------------
    case parser::NodeKind::ExternalRef:
    case parser::NodeKind::StructuredRef:
    case parser::NodeKind::NameRef:
    case parser::NodeKind::LambdaCall:
    case parser::NodeKind::Lambda:
    case parser::NodeKind::LetBinding:
      return Value::error(ErrorCode::Name);

    // -- Unsupported: range-producing operators / array literals ----------
    case parser::NodeKind::RangeOp:
    case parser::NodeKind::UnionOp:
    case parser::NodeKind::IntersectOp:
    case parser::NodeKind::ArrayLiteral:
      return Value::error(ErrorCode::Value);
  }
  return Value::error(ErrorCode::Value);
}

Value evaluate(const parser::AstNode& node, Arena& arena) {
  return evaluate(node, arena, default_registry(), EvalContext{});
}

Value evaluate(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry) {
  return evaluate(node, arena, registry, EvalContext{});
}

Value evaluate(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx) {
  return eval_node(node, arena, registry, ctx);
}

}  // namespace eval
}  // namespace formulon
