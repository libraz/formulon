// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the tree-walk evaluator. See `tree_walker.h` for the
// public contract and the design references.

#include "eval/tree_walker.h"

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "eval/areas_lazy.h"
#include "eval/coerce.h"
#include "eval/conditional_aggregates.h"
#include "eval/database_lazy.h"
#include "eval/eval_context.h"
#include "eval/financial_lazy.h"
#include "eval/function_registry.h"
#include "eval/hypothesis_lazy.h"
#include "eval/info_lazy.h"
#include "eval/lazy_impls.h"
#include "eval/lookups/classic.h"
#include "eval/lookups/xlookup.h"
#include "eval/name_env.h"
#include "eval/name_env_resolve.h"
#include "eval/range_args.h"
#include "eval/rank_lazy.h"
#include "eval/reference_lazy.h"
#include "eval/regression_lazy.h"
#include "eval/series_sum_lazy.h"
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

  // Blank is chameleonic in Excel equality comparisons: `blank = 0` is
  // TRUE (already handled by the numeric rank above), and `blank = ""` is
  // also TRUE. For the text-side comparison we treat a Blank operand as
  // the empty string so `IF(A1="","Empty",...)` fires on unset cells.
  const ValueKind lk = lhs.kind();
  const ValueKind rk = rhs.kind();
  if (lk == ValueKind::Blank && rk == ValueKind::Text) {
    const std::string_view b = rhs.as_text();
    if (b.empty()) {
      return 0;
    }
    return -1;  // Blank (as "") sorts before any non-empty string.
  }
  if (rk == ValueKind::Blank && lk == ValueKind::Text) {
    const std::string_view a = lhs.as_text();
    if (a.empty()) {
      return 0;
    }
    return 1;
  }

  const int lr = rank(lk);
  const int rr = rank(rk);
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
  // Unary `+` in Excel 365 is an identity operation — it does NOT coerce.
  // `=+""` evaluates to the empty string (not #VALUE!); `=+TRUE` stays
  // TRUE; numbers are returned unchanged. This diverges from Minus and
  // Percent, which explicitly require a numeric coercion.
  if (op == parser::UnaryOp::Plus) {
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
//   INDIRECT / OFFSET                          -> src/eval/reference_lazy.cpp
//   CORREL / COVARIANCE.P / COVARIANCE.S /
//   SLOPE / INTERCEPT / RSQ / FORECAST.LINEAR /
//   STEYX / SUMX2PY2 / SUMX2MY2 / SUMXMY2      -> src/eval/regression_lazy.cpp
//   SERIESSUM                                  -> src/eval/series_sum_lazy.cpp
//   RANK / RANK.EQ / RANK.AVG /
//   PERCENTRANK / PERCENTRANK.INC /
//   PERCENTRANK.EXC                            -> src/eval/rank_lazy.cpp
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
    {"AREAS", &eval_areas_lazy},
    {"AVERAGEIF", &eval_averageif_lazy},
    {"AVERAGEIFS", &eval_averageifs_lazy},
    // CHITEST is the pre-2010 legacy spelling of CHISQ.TEST; same impl.
    {"CHISQ.TEST", &eval_chisq_test_lazy},
    {"CHITEST", &eval_chisq_test_lazy},
    {"CHOOSE", &eval_choose_lazy},
    {"COLUMN", &eval_column_lazy},
    {"COLUMNS", &eval_columns_lazy},
    {"CORREL", &eval_correl_lazy},
    {"COUNT", &eval_count_lazy},
    {"COUNTIF", &eval_countif_lazy},
    {"COUNTIFS", &eval_countifs_lazy},
    // COVAR is the pre-2010 legacy spelling of COVARIANCE.P; both compute
    // the population covariance with identical semantics.
    {"COVAR", &eval_covariance_p_lazy},
    {"COVARIANCE.P", &eval_covariance_p_lazy},
    {"COVARIANCE.S", &eval_covariance_s_lazy},
    {"DAVERAGE", &eval_daverage_lazy},
    {"DCOUNT", &eval_dcount_lazy},
    {"DCOUNTA", &eval_dcounta_lazy},
    {"DGET", &eval_dget_lazy},
    {"DMAX", &eval_dmax_lazy},
    {"DMIN", &eval_dmin_lazy},
    {"DPRODUCT", &eval_dproduct_lazy},
    {"DSTDEV", &eval_dstdev_lazy},
    {"DSTDEVP", &eval_dstdevp_lazy},
    {"DSUM", &eval_dsum_lazy},
    {"DVAR", &eval_dvar_lazy},
    {"DVARP", &eval_dvarp_lazy},
    {"F.TEST", &eval_f_test_lazy},
    // FORECAST is the legacy spelling kept by Excel for back-compat;
    // its impl and arity are identical to FORECAST.LINEAR.
    {"FORECAST", &eval_forecast_linear_lazy},
    {"FORECAST.LINEAR", &eval_forecast_linear_lazy},
    // FORMULATEXT returns the source text of the referenced cell's formula,
    // so it must inspect the un-evaluated Ref AST and the bound Sheet's
    // `formula_text` directly — the eager path would flatten the argument
    // to a Value before we could see the reference.
    {"FORMULATEXT", &eval_formulatext_lazy},
    // FTEST is the pre-2010 legacy spelling of F.TEST; same impl.
    {"FTEST", &eval_f_test_lazy},
    {"HLOOKUP", &eval_hlookup_lazy},
    {"IF", &eval_if_lazy},
    {"IFERROR", &eval_iferror_lazy},
    {"IFNA", &eval_ifna_lazy},
    {"IFS", &eval_ifs_lazy},
    {"INDEX", &eval_index_lazy},
    {"INDIRECT", &eval_indirect_lazy},
    {"INTERCEPT", &eval_intercept_lazy},
    // ISFORMULA / ISREF inspect the un-evaluated AST of their argument;
    // they cannot ride the eager path because it flattens references to
    // `Value` before the impl runs.
    {"ISFORMULA", &eval_isformula_lazy},
    {"ISREF", &eval_isref_lazy},
    {"IRR", &eval_irr_lazy},
    {"LOOKUP", &eval_lookup_lazy},
    {"MATCH", &eval_match_lazy},
    {"MAXIFS", &eval_maxifs_lazy},
    {"MINIFS", &eval_minifs_lazy},
    {"MIRR", &eval_mirr_lazy},
    {"NETWORKDAYS", &eval_networkdays_lazy},
    {"NETWORKDAYS.INTL", &eval_networkdays_intl_lazy},
    {"OFFSET", &eval_offset_lazy},
    // PEARSON is mathematically identical to CORREL (Pearson product-moment
    // correlation coefficient); Excel keeps both names for back-compat.
    {"PEARSON", &eval_correl_lazy},
    {"PERCENTRANK", &eval_percentrank_inc_lazy},
    {"PERCENTRANK.EXC", &eval_percentrank_exc_lazy},
    {"PERCENTRANK.INC", &eval_percentrank_inc_lazy},
    {"PROB", &eval_prob_lazy},
    {"RANK", &eval_rank_eq_lazy},
    {"RANK.AVG", &eval_rank_avg_lazy},
    {"RANK.EQ", &eval_rank_eq_lazy},
    {"ROW", &eval_row_lazy},
    {"ROWS", &eval_rows_lazy},
    {"RSQ", &eval_rsq_lazy},
    {"SERIESSUM", &eval_series_sum_lazy},
    // SHEET / SHEETS consult the bound Workbook + current Sheet on the
    // EvalContext; AST introspection of an optional reference argument
    // tells them which sheet to answer for.
    {"SHEET", &eval_sheet_lazy},
    {"SHEETS", &eval_sheets_lazy},
    {"SLOPE", &eval_slope_lazy},
    {"STEYX", &eval_steyx_lazy},
    {"SUMIF", &eval_sumif_lazy},
    {"SUMIFS", &eval_sumifs_lazy},
    {"SUMPRODUCT", &eval_sumproduct_lazy},
    {"SUMX2MY2", &eval_sumx2my2_lazy},
    {"SUMX2PY2", &eval_sumx2py2_lazy},
    {"SUMXMY2", &eval_sumxmy2_lazy},
    {"SWITCH", &eval_switch_lazy},
    {"T.TEST", &eval_t_test_lazy},
    // TTEST is the pre-2010 legacy spelling of T.TEST; same impl.
    {"TTEST", &eval_t_test_lazy},
    {"VLOOKUP", &eval_vlookup_lazy},
    {"WORKDAY", &eval_workday_lazy},
    {"WORKDAY.INTL", &eval_workday_intl_lazy},
    {"XIRR", &eval_xirr_lazy},
    {"XLOOKUP", &eval_xlookup_lazy},
    {"XMATCH", &eval_xmatch_lazy},
    {"XNPV", &eval_xnpv_lazy},
    {"Z.TEST", &eval_z_test_lazy},
    // ZTEST is the pre-2010 legacy spelling of Z.TEST; same impl.
    {"ZTEST", &eval_z_test_lazy},
};

const LazyEntry* find_lazy(std::string_view name) noexcept {
  for (const auto& e : kLazyDispatch) {
    if (strings::case_insensitive_eq(name, std::string_view(e.name))) {
      return &e;
    }
  }
  return nullptr;
}

// Strips the xlsx-only `_xlfn.` and `_xlfn._xlws.` prefixes from a function
// name. These prefixes are a storage artifact: xlsx tags post-2007 functions
// with `_xlfn.` and modern worksheet-only ones (FILTER, XLOOKUP, LET, ...)
// with `_xlfn._xlws.` so older Excel versions don't accidentally try to
// evaluate them. Excel itself transparently strips the tag on load, so the
// canonical name is the only thing the registry knows about.
//
// Matches ASCII case-insensitively to tolerate `_xlfn.` vs `_XLFN.` casing.
std::string_view strip_future_prefix(std::string_view name) noexcept {
  constexpr std::string_view kXlws = "_xlfn._xlws.";
  constexpr std::string_view kXlfn = "_xlfn.";
  if (name.size() > kXlws.size() && strings::case_insensitive_eq(name.substr(0, kXlws.size()), kXlws)) {
    return name.substr(kXlws.size());
  }
  if (name.size() > kXlfn.size() && strings::case_insensitive_eq(name.substr(0, kXlfn.size()), kXlfn)) {
    return name.substr(kXlfn.size());
  }
  return name;
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
  const std::string_view name = strip_future_prefix(node.as_call_name());
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
  // Tracks whether any argument slot was range-shaped (RangeOp / OFFSET-call
  // / ArrayLiteral). Used by the deferred `RejectAnyScalar` blank-scalar
  // policy: Mac Excel only surfaces #VALUE! for `=GCD(A1,B1,C1)` (all blank
  // scalar refs) when there is NO range-shaped arg in the call. The mixed
  // form `=GCD(A1:B1, C1)` over the same blank cells still returns 0,
  // because the range arg "rescues" the policy.
  bool had_range_shaped_arg = false;
  // Tracks whether at least one scalar arg slot satisfied the
  // `RejectAnyScalar` policy (a Blank-valued scalar that is not a literal
  // zero). Combined with `had_range_shaped_arg` to decide whether to fire
  // the deferred error after the loop completes.
  bool any_scalar_blank_for_reject_any = false;
  for (std::uint32_t i = 0; i < arity; ++i) {
    const parser::AstNode& raw_arg = node.as_call_arg(i);
    // LET-binding passthrough: when the caller wrote `SUM(r)` where `r` is
    // bound to a RangeOp / ArrayLiteral / OFFSET / CHOOSE / INDIRECT, the
    // dispatcher must see the underlying range AST or it would fall back
    // to the scalar path and collapse the binding to its spill anchor.
    // Substitute only when the resolved AST is genuinely range-shaped so
    // that a NameRef bound to a scalar (or a single-cell Ref) continues to
    // flow through the existing scalar branch with its original provenance.
    const parser::AstNode* effective = &raw_arg;
    if (raw_arg.kind() == parser::NodeKind::NameRef) {
      const parser::AstNode& resolved = resolve_name_ast(raw_arg, ctx.name_env());
      if (&resolved != &raw_arg && is_range_shaped_ast(resolved)) {
        effective = &resolved;
      }
    }
    const parser::AstNode& arg_node = *effective;
    // Minimal array-literal support: when a range-aware function receives
    // a `{a;b;c}` style literal, flatten it in row-major order exactly like
    // a RangeOp argument. This is just enough to let LARGE / SMALL /
    // PERCENTILE.INC / QUARTILE.INC accept brace literals as their "array"
    // input; full array-aware evaluation (broadcasting, spilled output)
    // stays deferred — a bare `={1;2;3}` outside a function call still
    // surfaces as #VALUE! via `eval_node`'s ArrayLiteral case.
    if (def->accepts_ranges && arg_node.kind() == parser::NodeKind::ArrayLiteral) {
      had_range_shaped_arg = true;
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
          // Provenance-aware filtering for array-literal-sourced values.
          // An element inside a `{...}` literal is treated the same way as
          // a cell inside a range: non-numeric elements are dropped for
          // `range_filter_numeric_only`, coerced for the A-family, etc.
          // Direct scalar arguments (handled in the fall-through branch
          // below) still coerce through the impl's strict rules.
          if (def->range_filter_numeric_only && v.kind() != ValueKind::Number) {
            continue;
          }
          if (def->range_filter_bool_coercible && v.kind() != ValueKind::Number && v.kind() != ValueKind::Bool) {
            continue;
          }
          if (def->range_filter_a_coerce) {
            if (v.kind() == ValueKind::Blank) {
              continue;
            }
            if (v.kind() == ValueKind::Bool) {
              values.push_back(Value::number(v.as_boolean() ? 1.0 : 0.0));
              continue;
            }
            if (v.kind() == ValueKind::Text) {
              values.push_back(Value::number(0.0));
              continue;
            }
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
      had_range_shaped_arg = true;
      const parser::AstNode& lhs_ast = arg_node.as_range_lhs();
      const parser::AstNode& rhs_ast = arg_node.as_range_rhs();
      // Endpoints may be plain Refs (the simple `A1:B2` form) or
      // reference-producing calls (`OFFSET(...)` / `INDIRECT(...)`);
      // `resolve_range_endpoint` normalises both to a rectangle so we
      // can union them and feed `expand_range` two synthetic Refs.
      // Anything else (literals, arithmetic, named ranges, …) surfaces
      // as #REF! / #VALUE! per the helper's error code.
      std::string_view lhs_sheet;
      std::string_view rhs_sheet;
      std::uint32_t lhs_top = 0;
      std::uint32_t lhs_left = 0;
      std::uint32_t lhs_bottom = 0;
      std::uint32_t lhs_right = 0;
      std::uint32_t rhs_top = 0;
      std::uint32_t rhs_left = 0;
      std::uint32_t rhs_bottom = 0;
      std::uint32_t rhs_right = 0;
      ErrorCode endpoint_err = ErrorCode::Ref;
      if (!resolve_range_endpoint(lhs_ast, arena, registry, ctx, &lhs_sheet, &lhs_top, &lhs_left, &lhs_bottom,
                                  &lhs_right, &endpoint_err) ||
          !resolve_range_endpoint(rhs_ast, arena, registry, ctx, &rhs_sheet, &rhs_top, &rhs_left, &rhs_bottom,
                                  &rhs_right, &endpoint_err)) {
        const Value err = Value::error(endpoint_err);
        if (def->propagate_errors) {
          return err;
        }
        values.push_back(err);
        continue;
      }
      // Build the union rectangle's two corner Refs and let
      // `expand_range` validate sheet equality (mismatched qualifiers
      // surface as #REF!).
      parser::Reference union_lhs{};
      parser::Reference union_rhs{};
      union_lhs.sheet = lhs_sheet;
      union_lhs.row = std::min(lhs_top, rhs_top);
      union_lhs.col = std::min(lhs_left, rhs_left);
      union_rhs.sheet = rhs_sheet;
      union_rhs.row = std::max(lhs_bottom, rhs_bottom);
      union_rhs.col = std::max(lhs_right, rhs_right);
      auto expanded = ctx.expand_range(union_lhs, union_rhs, arena, registry);
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
        if (def->range_filter_a_coerce) {
          // A-family (AVERAGEA / MAXA / MINA / VAR{A,PA} / STDEV{A,PA}).
          // Bool and Text cells inside a range are coerced to numbers rather
          // than dropped: TRUE -> 1, FALSE -> 0, any text (including the
          // empty string and numeric-looking strings like "3.14") -> 0.
          // Blank cells are still skipped.
          if (v.kind() == ValueKind::Blank) {
            continue;
          }
          if (v.kind() == ValueKind::Bool) {
            values.push_back(Value::number(v.as_boolean() ? 1.0 : 0.0));
            continue;
          }
          if (v.kind() == ValueKind::Text) {
            values.push_back(Value::number(0.0));
            continue;
          }
        }
        values.push_back(v);
      }
      continue;
    }
    // Range-aware functions that receive `OFFSET(...)` as an argument see
    // the rectangle OFFSET would synthesize, not the `#VALUE!` OFFSET
    // itself returns in scalar context. We share the expansion helper
    // with `resolve_range_arg` (lazy family) so the two paths cannot
    // drift on cross-sheet / cycle / bounds semantics. Other calls (e.g.
    // `INDIRECT("A1:B2")`) would need a `Value::Array` runtime to
    // expand; they fall through to `eval_node` and surface whatever
    // scalar result OFFSET / INDIRECT produces today.
    if (def->accepts_ranges && arg_node.kind() == parser::NodeKind::Call &&
        strings::case_insensitive_eq(arg_node.as_call_name(), "OFFSET")) {
      had_range_shaped_arg = true;
      std::vector<Value> off_cells;
      ErrorCode off_err = ErrorCode::Value;
      if (!expand_offset_call(arg_node, arena, registry, ctx, &off_cells, &off_err, nullptr, nullptr)) {
        const Value err = Value::error(off_err);
        if (def->propagate_errors) {
          return err;
        }
        values.push_back(err);
        continue;
      }
      for (const Value& v : off_cells) {
        if (def->propagate_errors && v.is_error()) {
          return v;
        }
        if (def->range_filter_numeric_only && v.kind() != ValueKind::Number) {
          continue;
        }
        if (def->range_filter_bool_coercible && v.kind() != ValueKind::Number && v.kind() != ValueKind::Bool) {
          continue;
        }
        if (def->range_filter_a_coerce) {
          // A-family (AVERAGEA / MAXA / MINA / VAR{A,PA} / STDEV{A,PA}).
          // Bool and Text cells inside a range are coerced to numbers rather
          // than dropped: TRUE -> 1, FALSE -> 0, any text (including the
          // empty string and numeric-looking strings like "3.14") -> 0.
          // Blank cells are still skipped.
          if (v.kind() == ValueKind::Blank) {
            continue;
          }
          if (v.kind() == ValueKind::Bool) {
            values.push_back(Value::number(v.as_boolean() ? 1.0 : 0.0));
            continue;
          }
          if (v.kind() == ValueKind::Text) {
            values.push_back(Value::number(0.0));
            continue;
          }
        }
        values.push_back(v);
      }
      continue;
    }
    // CHOOSE-as-range-producer mirrors the OFFSET branch above: when an
    // aggregator receives `CHOOSE(idx, range1, range2, ...)`, the chosen
    // child must be flattened to a vector of cells (recursively, if it is
    // itself an OFFSET / CHOOSE call). `expand_choose_call` shares the
    // same evaluation, range-resolution, and filter contracts so SUM /
    // AVERAGE / MIN / MAX / AVERAGEA all behave consistently.
    if (def->accepts_ranges && arg_node.kind() == parser::NodeKind::Call &&
        strings::case_insensitive_eq(arg_node.as_call_name(), "CHOOSE")) {
      had_range_shaped_arg = true;
      std::vector<Value> ch_cells;
      ErrorCode ch_err = ErrorCode::Value;
      if (!expand_choose_call(arg_node, arena, registry, ctx, &ch_cells, &ch_err, nullptr, nullptr)) {
        const Value err = Value::error(ch_err);
        if (def->propagate_errors) {
          return err;
        }
        values.push_back(err);
        continue;
      }
      for (const Value& v : ch_cells) {
        if (def->propagate_errors && v.is_error()) {
          return v;
        }
        if (def->range_filter_numeric_only && v.kind() != ValueKind::Number) {
          continue;
        }
        if (def->range_filter_bool_coercible && v.kind() != ValueKind::Number && v.kind() != ValueKind::Bool) {
          continue;
        }
        if (def->range_filter_a_coerce) {
          // A-family (AVERAGEA / MAXA / MINA / VAR{A,PA} / STDEV{A,PA}).
          // Bool and Text cells inside a range are coerced to numbers rather
          // than dropped: TRUE -> 1, FALSE -> 0, any text (including the
          // empty string and numeric-looking strings like "3.14") -> 0.
          // Blank cells are still skipped.
          if (v.kind() == ValueKind::Blank) {
            continue;
          }
          if (v.kind() == ValueKind::Bool) {
            values.push_back(Value::number(v.as_boolean() ? 1.0 : 0.0));
            continue;
          }
          if (v.kind() == ValueKind::Text) {
            values.push_back(Value::number(0.0));
            continue;
          }
        }
        values.push_back(v);
      }
      continue;
    }
    Value v = eval_node(arg_node, arena, registry, ctx);
    if (def->propagate_errors && v.is_error()) {
      return v;
    }
    // Blank-scalar policy. RangeOp / OFFSET-call / ArrayLiteral args were
    // handled above and `continue`'d, so reaching this point implies a
    // scalar arg slot (Literal, Ref, BinaryOp, ...).
    //
    //   * `RejectLiteralEmpty` (MROUND) fires eagerly: only a parser-injected
    //     `Literal(blank)` for an empty arg slot triggers it. A Ref to a
    //     blank cell still flows through to the impl as 0, matching Mac.
    //   * `RejectAnyScalar` (GCD / LCM) defers the decision to end-of-args.
    //     Mac surfaces #VALUE! for `=GCD(A1,B1,C1)` (all blank scalar refs)
    //     but returns 0 for the mixed form `=GCD(A1:B1, C1)` over the same
    //     blank cells — the range arg "rescues" the call. The flag is
    //     consulted after the loop in conjunction with
    //     `had_range_shaped_arg`. Direct numeric literals (including
    //     `=GCD(0,0,0)`) are not Blank and do not set the flag.
    if (def->blank_scalar_policy == FunctionDef::BlankScalarPolicy::RejectLiteralEmpty &&
        v.kind() == ValueKind::Blank && arg_node.kind() == parser::NodeKind::Literal) {
      const Value err = Value::error(def->blank_scalar_error);
      if (def->propagate_errors) {
        return err;
      }
      values.push_back(err);
      continue;
    }
    if (def->blank_scalar_policy == FunctionDef::BlankScalarPolicy::RejectAnyScalar && v.kind() == ValueKind::Blank) {
      any_scalar_blank_for_reject_any = true;
    }
    // Provenance-aware filter also applies to a single-cell `Ref` argument
    // for range-aware aggregators. Excel treats `MIN(A1, A2, A3)` the same
    // way it treats `MIN(A1:A3)`: Text / Bool / Blank cells are silently
    // skipped for `range_filter_numeric_only`, coerced for the A-family,
    // etc. Direct scalar literals (numbers, bool literals, text literals)
    // still use strict coercion in the impl.
    if (def->accepts_ranges && arg_node.kind() == parser::NodeKind::Ref) {
      if (def->range_filter_numeric_only && v.kind() != ValueKind::Number) {
        continue;
      }
      if (def->range_filter_bool_coercible && v.kind() != ValueKind::Number && v.kind() != ValueKind::Bool) {
        continue;
      }
      if (def->range_filter_a_coerce) {
        if (v.kind() == ValueKind::Blank) {
          continue;
        }
        if (v.kind() == ValueKind::Bool) {
          values.push_back(Value::number(v.as_boolean() ? 1.0 : 0.0));
          continue;
        }
        if (v.kind() == ValueKind::Text) {
          values.push_back(Value::number(0.0));
          continue;
        }
      }
    }
    values.push_back(v);
  }
  // Deferred fire-point for `RejectAnyScalar`: only surface the error when
  // there is no range-shaped arg in the call. The mixed form
  // `=GCD(A1:B1, C1)` keeps `had_range_shaped_arg = true` and so falls
  // through to the impl, where the lone blank scalar coerces to 0 and the
  // result matches Mac. Pure-scalar all-blank shapes
  // (`=GCD(A1,B1,C1)`) surface the policy error here.
  if (def->blank_scalar_policy == FunctionDef::BlankScalarPolicy::RejectAnyScalar && any_scalar_blank_for_reject_any &&
      !had_range_shaped_arg) {
    return Value::error(def->blank_scalar_error);
  }
  // Hand the post-expansion size to the impl; aggregator bodies walk the
  // flattened vector directly.
  return def->impl(values.data(), static_cast<std::uint32_t>(values.size()), arena);
}

}  // namespace

// Enumerates the canonical UPPERCASE names in `kLazyDispatch`. Kept in the
// same translation unit as the dispatch table so the array it hands out
// cannot drift from the actual routing decisions: the names are copied
// once from `kLazyDispatch[i].name` into a nullptr-terminated buffer with
// program lifetime. Declared in `eval/tree_walker.h`.
const char* const* lazy_form_names() {
  static constexpr std::size_t kCount = sizeof(kLazyDispatch) / sizeof(kLazyDispatch[0]);
  // The storage itself has static duration and is initialized exactly once
  // (C++11 magic statics); the lambda builds the nullptr-terminated view
  // from the adjacent dispatch table. No allocation, no ordering concerns.
  static const char* const* kTable = [] {
    static const char* storage[kCount + 1] = {};
    for (std::size_t i = 0; i < kCount; ++i) {
      storage[i] = kLazyDispatch[i].name;
    }
    storage[kCount] = nullptr;
    return static_cast<const char* const*>(storage);
  }();
  return kTable;
}

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

    case parser::NodeKind::RangeOp: {
      // Excel 365 dynamic-array spill vs. legacy implicit intersection.
      // Both behaviors map a bare range used in scalar context onto a
      // single cell; the difference is which cell:
      //
      //   * Mac Excel 365 with a fresh-typed formula spills the array
      //     and a single-cell reader (xlwings, Formulon's evaluator
      //     entry) sees the spill anchor = top-left of the source range.
      //   * Pre-365 / @-prefix auto-promoted formulas use implicit
      //     intersection: the formula cell's row (vertical range) or
      //     column (horizontal range) selects the aligned cell, with
      //     #VALUE! when the formula cell is outside the range.
      //
      // We split the difference observationally: try row/col alignment
      // first when the formula cell is INSIDE the range -- this matches
      // legacy II for the cases IronCalc fixtures cache. When the
      // formula cell is OUTSIDE the range (or the range is 2D, or no
      // formula cell is bound), fall back to the top-left, matching
      // Mac's spill anchor. At the Z1 anchor the Mac oracle uses, the
      // two behaviors are identical because Z1's row=0/col=25 either
      // matches the range top-left or sits outside the range entirely.
      //
      // Verified Mac semantics: tests/oracle/cases/implicit_intersection.yaml.
      const auto& lhs = node.as_range_lhs();
      const auto& rhs = node.as_range_rhs();
      if (lhs.kind() != parser::NodeKind::Ref || rhs.kind() != parser::NodeKind::Ref) {
        // Complex range expressions (OFFSET-based, INDIRECT, named ranges)
        // are not yet resolved at this level; surface #VALUE! as before.
        return Value::error(ErrorCode::Value);
      }
      const auto& lhs_ref = lhs.as_ref();
      const auto& rhs_ref = rhs.as_ref();
      const std::uint32_t r1 = std::min(lhs_ref.row, rhs_ref.row);
      const std::uint32_t r2 = std::max(lhs_ref.row, rhs_ref.row);
      const std::uint32_t c1 = std::min(lhs_ref.col, rhs_ref.col);
      const std::uint32_t c2 = std::max(lhs_ref.col, rhs_ref.col);

      if (ctx.has_formula_cell()) {
        const std::uint32_t fr = ctx.formula_row();
        const std::uint32_t fc = ctx.formula_col();
        parser::Reference target{};
        target.sheet = lhs_ref.sheet;
        if (c1 == c2) {
          if (fr >= r1 && fr <= r2) {
            target.row = fr;
            target.col = c1;
            return ctx.resolve_ref(target, arena, registry);
          }
        } else if (r1 == r2) {
          if (fc >= c1 && fc <= c2) {
            target.row = r1;
            target.col = fc;
            return ctx.resolve_ref(target, arena, registry);
          }
        }
        // 2D range: alignment requires both axes; fall through to top-left.
      }

      parser::Reference top_left{};
      top_left.sheet = lhs_ref.sheet;
      top_left.row = r1;
      top_left.col = c1;
      return ctx.resolve_ref(top_left, arena, registry);
    }

    case parser::NodeKind::ImplicitIntersection: {
      const auto& operand = node.as_implicit_intersection_operand();
      if (operand.kind() == parser::NodeKind::RangeOp) {
        // Implicit intersection on a range: project onto the formula cell's
        // row or column. Single-column range -> requires formula row in
        // range; single-row range -> requires formula col in range. 2D
        // ranges and any non-aligned cases return #VALUE!.
        const auto& lhs_ast = operand.as_range_lhs();
        const auto& rhs_ast = operand.as_range_rhs();
        if (lhs_ast.kind() != parser::NodeKind::Ref || rhs_ast.kind() != parser::NodeKind::Ref) {
          return Value::error(ErrorCode::Value);
        }
        if (!ctx.has_formula_cell()) {
          // No formula-cell context (top-level evaluator entry) -> degrade to
          // top-left, matching the bare-range fallback. Production calls
          // through Workbook always supply a formula cell, so this branch
          // only fires for parser-driven smoke tests.
          return eval_node(operand, arena, registry, ctx);
        }
        const auto& lhs_ref = lhs_ast.as_ref();
        const auto& rhs_ref = rhs_ast.as_ref();
        const std::uint32_t r1 = std::min(lhs_ref.row, rhs_ref.row);
        const std::uint32_t r2 = std::max(lhs_ref.row, rhs_ref.row);
        const std::uint32_t c1 = std::min(lhs_ref.col, rhs_ref.col);
        const std::uint32_t c2 = std::max(lhs_ref.col, rhs_ref.col);
        const std::uint32_t fr = ctx.formula_row();
        const std::uint32_t fc = ctx.formula_col();
        parser::Reference target{};
        target.sheet = lhs_ref.sheet;
        if (c1 == c2) {
          // Single-column range: project formula row.
          if (fr < r1 || fr > r2) {
            return Value::error(ErrorCode::Value);
          }
          target.row = fr;
          target.col = c1;
        } else if (r1 == r2) {
          // Single-row range: project formula column.
          if (fc < c1 || fc > c2) {
            return Value::error(ErrorCode::Value);
          }
          target.row = r1;
          target.col = fc;
        } else {
          // 2D range: implicit intersection requires both axes to align,
          // and Excel returns #VALUE! when the formula cell isn't covered
          // by both spans. Verified Mac behavior pending; conservative
          // default for now.
          return Value::error(ErrorCode::Value);
        }
        return ctx.resolve_ref(target, arena, registry);
      }
      // Non-range operands: identity (current pass-through behavior).
      return eval_node(operand, arena, registry, ctx);
    }

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

    case parser::NodeKind::NameRef: {
      // Lexical-scope lookup for LET (and, eventually, LAMBDA) bindings.
      // When the name is not in scope we surface `#NAME?`; defined-name
      // resolution at workbook scope is a separate infrastructure pass and
      // intentionally not handled here.
      const NameEnv* env = ctx.name_env();
      if (env != nullptr) {
        if (const Value* bound = env->lookup(node.as_name()); bound != nullptr) {
          return *bound;
        }
      }
      return Value::error(ErrorCode::Name);
    }

    case parser::NodeKind::LetBinding: {
      // Sequential (left-to-right) bind-then-body. Excel semantics:
      //   * Each binding initialiser evaluates in the scope of previously
      //     bound names, so `LET(x, 1, y, x+2, y)` returns 3.
      //   * Error values DO flow into the environment -- downstream
      //     expressions (including `IFERROR` inside the body) may catch
      //     them: `LET(x, 1/0, IFERROR(x, 99))` returns 99.
      //   * Names are ASCII-case-insensitive and a later binding with the
      //     same name shadows earlier ones in subsequent expressions.
      //   * Range-shaped initialisers (RangeOp, ArrayLiteral, OFFSET/CHOOSE/
      //     INDIRECT calls, single-cell Refs, or NameRefs that resolve to
      //     such) keep a pointer to their source AST in the binding so that
      //     range-aware consumers (SUM, COUNT, VLOOKUP, ...) can re-dispatch
      //     on the underlying shape rather than seeing only the spill-anchor
      //     scalar that `eval_node` collapses a bare RangeOp to.
      NameEnv env;
      const NameEnv* parent = ctx.name_env();
      // Start from whatever the caller supplied; extending `NameEnv` makes
      // `env` point at a new head frame while preserving the parent chain.
      if (parent != nullptr) {
        env = *parent;
      }
      const std::uint32_t count = node.as_let_binding_count();
      for (std::uint32_t i = 0; i < count; ++i) {
        const parser::AstNode& expr_node = node.as_let_binding_expr(i);
        const EvalContext inner_ctx = ctx.with_name_env(&env);
        const Value v = eval_node(expr_node, arena, registry, inner_ctx);
        // Record the AST source for genuinely range-shaped bindings only:
        // `RangeOp`, `ArrayLiteral`, and the reference-producing calls
        // (`OFFSET`, `CHOOSE`, `INDIRECT`). Single-cell `Ref` and scalar
        // shapes bind by Value -- recording their AST would either be
        // unused (the resolve helper skips non-range shapes) or risk
        // re-evaluating side-effecting / expensive sub-expressions on
        // every NameRef read. NameRef-on-NameRef is transitive: if the
        // RHS already resolves to a range-shaped AST in the (possibly
        // outer) scope, we inherit that AST so `=LET(s, r, SUM(s))` works
        // when `r` is itself a range binding.
        const parser::AstNode* expr_for_binding = nullptr;
        if (is_range_shaped_ast(expr_node)) {
          expr_for_binding = &expr_node;
        } else if (expr_node.kind() == parser::NodeKind::NameRef) {
          // `env` already reflects every previously bound name in this LET
          // (and, via `parent`, any outer LETs); a single lookup walks the
          // whole chain.
          expr_for_binding = env.lookup_ast(expr_node.as_name());
        }
        env = env.extend(node.as_let_binding_name(i), v, expr_for_binding,
                         arena);
      }
      const EvalContext body_ctx = ctx.with_name_env(&env);
      return eval_node(node.as_let_body(), arena, registry, body_ctx);
    }

    // -- Unsupported: closures / external names ---------------------------
    case parser::NodeKind::ExternalRef:
    case parser::NodeKind::StructuredRef:
    case parser::NodeKind::LambdaCall:
    case parser::NodeKind::Lambda:
      // LAMBDA requires closure capture which is not yet modelled by the
      // scalar `Value` variant; defer to follow-up work.
      return Value::error(ErrorCode::Name);

    // -- Unsupported: range-producing operators / array literals ----------
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
  Value v = eval_node(node, arena, registry, ctx);
  // Mac Excel 365 displays a blank-cell-resolved formula result as 0 in
  // numeric column rendering, and the oracle pipeline reads it back as
  // number(0.0). We mirror that here so plain `=A1` (A1 blank) and other
  // top-level reference paths agree with Mac. The Literal-Blank case from
  // the AST factory (used in unit tests like `BlankFromFactory` to verify
  // the value variant) is preserved by gating on node kind: a literal
  // Blank node remains Blank to keep the variant inspectable from tests.
  if (v.is_blank() && node.kind() != parser::NodeKind::Literal) {
    return Value::number(0.0);
  }
  return v;
}

}  // namespace eval
}  // namespace formulon
