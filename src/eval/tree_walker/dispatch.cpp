// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Function-call dispatch path of the tree-walk evaluator. Split out of
// `tree_walker.cpp` to keep the recursive walker compile unit small;
// see `tree_walker/dispatch.h` for the public contract.
//
// Lazy entries (`IF`, `IFERROR`, `IFNA`, the `*IF`/`*IFS` aggregators,
// the lookup family, ...) are routed through the central lazy dispatch
// table in `tree_walker_lazy_table.cpp`; each impl owns its own arity
// check and chooses which subtrees to evaluate.
//
// All other names are routed through `FunctionRegistry`:
//   * Unknown name -> #NAME?
//   * Arity violation -> #VALUE!
//   * Otherwise every argument is pre-evaluated in order; by default the
//     left-most error short-circuits before the impl runs, but an entry
//     whose `propagate_errors` flag is `false` (the IS* type-predicate
//     family) opts out and receives raw error values among its arguments.
//
// Range-aware `accepts_ranges` entries expand range-shaped arguments
// (RangeOp, SpillRef, Ref3D, IntersectOp, StructuredRef, ArrayLiteral,
// and range-producing calls like OFFSET / CHOOSE / IF / ROW / COLUMN)
// into a flat values vector before invoking the impl. The expansion
// applies the per-cell `range_filter_*` rules so blank / text / bool
// cells are dropped or coerced consistently across input shapes.

#include "eval/tree_walker/dispatch.h"

#include <algorithm>
#include <cstdint>
#include <string_view>
#include <vector>

#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/lambda_value.h"
#include "eval/lazy_impls.h"
#include "eval/name_env.h"
#include "eval/name_env_resolve.h"
#include "eval/range_expanders.h"
#include "eval/range_resolvers.h"
#include "eval/tree_walker/depth_guard.h"
#include "eval/tree_walker_lazy_table.h"
#include "parser/ast.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/strings.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

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

bool append_range_sourced_value(const FunctionDef& def, const Value& v, std::vector<Value>* values, Value* out_err) {
  if (def.propagate_errors && v.is_error()) {
    *out_err = v;
    return false;
  }
  if (def.range_filter_numeric_only && v.kind() != ValueKind::Number) {
    return true;
  }
  if (def.range_filter_bool_coercible && v.kind() != ValueKind::Number && v.kind() != ValueKind::Bool) {
    return true;
  }
  if (def.range_filter_a_coerce) {
    if (v.kind() == ValueKind::Blank) {
      return true;
    }
    if (v.kind() == ValueKind::Bool) {
      values->push_back(Value::number(v.as_boolean() ? 1.0 : 0.0));
      return true;
    }
    if (v.kind() == ValueKind::Text) {
      values->push_back(Value::number(0.0));
      return true;
    }
  }
  values->push_back(v);
  return true;
}

bool append_range_sourced_values(const FunctionDef& def, const Value* cells, std::size_t count,
                                 std::vector<Value>* values, Value* out_err) {
  for (std::size_t i = 0; i < count; ++i) {
    if (!append_range_sourced_value(def, cells[i], values, out_err)) {
      return false;
    }
  }
  return true;
}

using RangeCallExpander = bool (*)(const parser::AstNode&, Arena&, const FunctionRegistry&, const EvalContext&,
                                   std::vector<Value>*, ErrorCode*, std::uint32_t*, std::uint32_t*);

bool append_expanded_call_argument(const FunctionDef& def, const parser::AstNode& arg_node, std::string_view call_name,
                                   RangeCallExpander expand, Arena& arena, const FunctionRegistry& registry,
                                   const EvalContext& ctx, std::vector<Value>* values, bool* handled,
                                   Value* immediate_return) {
  *handled = false;
  if (!def.accepts_ranges || arg_node.kind() != parser::NodeKind::Call ||
      !strings::case_insensitive_eq(arg_node.as_call_name(), call_name)) {
    return true;
  }
  *handled = true;
  std::vector<Value> cells;
  ErrorCode err_code = ErrorCode::Value;
  if (!expand(arg_node, arena, registry, ctx, &cells, &err_code, nullptr, nullptr)) {
    const Value err = Value::error(err_code);
    if (def.propagate_errors) {
      *immediate_return = err;
      return false;
    }
    values->push_back(err);
    return true;
  }
  Value range_err = Value::blank();
  if (!append_range_sourced_values(def, cells.data(), cells.size(), values, &range_err)) {
    *immediate_return = range_err;
    return false;
  }
  return true;
}

}  // namespace

Value invoke_lambda(const LambdaValue* lv, std::uint32_t arity, const parser::AstNode* const* call_args, Arena& arena,
                    const FunctionRegistry& registry, const EvalContext& ctx) {
  // Lambda-depth cap fires before the arity check so a runaway
  // self-recursion (e.g. `LAMBDA(n, f(n+1))`) cannot keep extending the
  // call stack on its own dime. See `kMaxLambdaDepth` for rationale.
  EvalDepthGuard lambda_guard(ctx.lambda_depth_counter(), kMaxLambdaDepth);
  if (lambda_guard.exceeded()) {
    return Value::error(ErrorCode::Calc);
  }
  const std::uint32_t required = lv->param_count - lv->optional_count;
  if (arity < required || arity > lv->param_count) {
    return Value::error(ErrorCode::Value);
  }
  NameEnv env;
  if (lv->captured_env != nullptr) {
    env = *lv->captured_env;
  }
  for (std::uint32_t i = 0; i < arity; ++i) {
    const Value arg = eval_node(*call_args[i], arena, registry, ctx);
    if (arg.is_error()) {
      return arg;
    }
    env = env.extend(lv->params[i], arg, arena);
  }
  for (std::uint32_t i = arity; i < lv->param_count; ++i) {
    env = env.extend_omitted(lv->params[i], arena);
  }
  const EvalContext body_ctx = ctx.with_name_env(&env);
  return eval_node(*lv->body, arena, registry, body_ctx);
}

Value dispatch_call(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                    const EvalContext& ctx) {
  const std::string_view name = strip_future_prefix(node.as_call_name());
  const std::uint32_t arity = node.as_call_arity();

  // Name-bound lambda dispatch: when the formula text reads `f(x)` and `f`
  // resolves to a runtime `LambdaValue` via the lexical name environment
  // (LET-bound or LAMBDA-parameter), invoke it as if the user had written
  // an explicit IIFE. The lookup runs *before* the registry / lazy table so
  // a LET binding can shadow a built-in name (matching Excel's semantics).
  // A bound non-Lambda value is `#VALUE!` (calling a non-callable); a
  // bound error propagates verbatim. Unbound names fall through to the
  // existing registry path.
  if (const NameEnv* env = ctx.name_env(); env != nullptr) {
    if (const Value* bound = env->lookup(name); bound != nullptr) {
      if (bound->is_lambda()) {
        // Build a flat argv pointer array from the call's child slots.
        std::vector<const parser::AstNode*> argv;
        argv.reserve(arity);
        for (std::uint32_t i = 0; i < arity; ++i) {
          argv.push_back(&node.as_call_arg(i));
        }
        return invoke_lambda(bound->as_lambda(), arity, argv.empty() ? nullptr : argv.data(), arena, registry, ctx);
      }
      if (bound->is_error()) {
        return *bound;
      }
      return Value::error(ErrorCode::Value);
    }
  }

  if (LazyImpl lazy = find_lazy_impl(name); lazy != nullptr) {
    return lazy(node, arena, registry, ctx);
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
          if (!append_range_sourced_value(*def, v, &values, &propagated_err)) {
            short_circuit = true;
            break;
          }
        }
      }
      if (short_circuit) {
        return propagated_err;
      }
      continue;
    }
    // Spilled-range `A1#` argument: resolve the spill region anchored at
    // the reference and flatten its row-major cells into the values
    // vector. Mirrors the RangeOp branch below; the same provenance-aware
    // filters apply because cells inside a spill region behave the same
    // way as cells inside an ordinary range when consumed by SUM /
    // AVERAGE / MIN / MAX / etc. A missing spill yields `#REF!`; an
    // unbound context yields `#NAME?`. Errors short-circuit per
    // `propagate_errors`.
    if (def->accepts_ranges && arg_node.kind() == parser::NodeKind::SpillRef) {
      had_range_shaped_arg = true;
      const parser::Reference& sr = arg_node.as_spill_ref();
      const Sheet* current = ctx.current_sheet();
      const Sheet* target = current;
      ErrorCode spill_err = ErrorCode::Ref;
      if (current == nullptr) {
        spill_err = ErrorCode::Name;
      } else if (!sr.sheet.empty()) {
        const Workbook* wb = ctx.workbook();
        if (wb == nullptr) {
          target = nullptr;
        } else {
          target = wb->sheet_by_name(sr.sheet);
        }
      }
      if (target == nullptr || sr.row >= Sheet::kMaxRows || sr.col >= Sheet::kMaxCols) {
        const Value err = Value::error(spill_err);
        if (def->propagate_errors) {
          return err;
        }
        values.push_back(err);
        continue;
      }
      const SpillRegion* region = target->spill_region_at_anchor(sr.row, sr.col);
      if (region == nullptr) {
        const Value err = Value::error(ErrorCode::Ref);
        if (def->propagate_errors) {
          return err;
        }
        values.push_back(err);
        continue;
      }
      Value range_err = Value::blank();
      if (!append_range_sourced_values(*def, region->cells.data(), region->cells.size(), &values, &range_err)) {
        return range_err;
      }
      continue;
    }
    // 3-D reference argument (`SUM(Sheet2:Sheet3!A1)`): resolve the inclusive
    // sheet span by workbook order, read the referenced cell from each sheet,
    // and flatten the resulting cells into the values vector. Mirrors the
    // RangeOp / SpillRef branches; the same provenance filters apply because
    // a 3-D ref is a range shape. A span endpoint that names a missing sheet
    // surfaces `#REF!`. Errors short-circuit per `propagate_errors`.
    if (def->accepts_ranges && arg_node.kind() == parser::NodeKind::Ref3D) {
      const Workbook* wb = ctx.workbook();
      const parser::Reference& cell = arg_node.as_ref3d_cell();
      ErrorCode ref3d_err = ErrorCode::Ref;
      std::size_t begin_idx = static_cast<std::size_t>(-1);
      std::size_t end_idx = static_cast<std::size_t>(-1);
      if (wb != nullptr) {
        begin_idx = wb->sheet_index_by_name(arg_node.as_ref3d_sheet_begin());
        end_idx = wb->sheet_index_by_name(arg_node.as_ref3d_sheet_end());
      }
      if (wb == nullptr || begin_idx == static_cast<std::size_t>(-1) || end_idx == static_cast<std::size_t>(-1) ||
          cell.is_full_col || cell.is_full_row || cell.row >= Sheet::kMaxRows || cell.col >= Sheet::kMaxCols) {
        const Value err = Value::error(ref3d_err);
        if (def->propagate_errors) {
          return err;
        }
        values.push_back(err);
        continue;
      }
      had_range_shaped_arg = true;
      const std::size_t lo = std::min(begin_idx, end_idx);
      const std::size_t hi = std::max(begin_idx, end_idx);
      std::vector<Value> ref3d_cells;
      ref3d_cells.reserve(hi - lo + 1);
      for (std::size_t s = lo; s <= hi; ++s) {
        parser::Reference per_sheet = cell;
        per_sheet.sheet = wb->sheet(s).name();
        ref3d_cells.push_back(ctx.resolve_ref(per_sheet, arena, registry));
      }
      Value range_err = Value::blank();
      if (!append_range_sourced_values(*def, ref3d_cells.data(), ref3d_cells.size(), &values, &range_err)) {
        return range_err;
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
      Value range_err = Value::blank();
      const std::vector<Value>& expanded_values = expanded.value();
      if (!append_range_sourced_values(*def, expanded_values.data(), expanded_values.size(), &values, &range_err)) {
        return range_err;
      }
      continue;
    }
    // Intersection operator as a range-aware function argument: Excel's
    // space operator (`A1:C3 B1:B5`) yields the overlapping rectangle,
    // and an aggregator must see every cell of that rectangle rather
    // than the single anchor `eval_node` would collapse it to. Compute
    // the clipped intersection rectangle and flatten it row-major,
    // mirroring the RangeOp branch above. Disjoint operands -> #NULL!.
    if (def->accepts_ranges && arg_node.kind() == parser::NodeKind::IntersectOp) {
      std::string_view isect_sheet;
      std::uint32_t isect_r1 = 0;
      std::uint32_t isect_c1 = 0;
      std::uint32_t isect_r2 = 0;
      std::uint32_t isect_c2 = 0;
      bool isect_disjoint = false;
      ErrorCode isect_err = ErrorCode::Value;
      if (!compute_intersect_rect(arg_node.as_intersect_lhs(), arg_node.as_intersect_rhs(), arena, registry, ctx,
                                  &isect_sheet, &isect_r1, &isect_c1, &isect_r2, &isect_c2, &isect_disjoint,
                                  &isect_err)) {
        const Value err = Value::error(isect_err);
        if (def->propagate_errors) {
          return err;
        }
        values.push_back(err);
        continue;
      }
      if (isect_disjoint) {
        const Value err = Value::error(ErrorCode::Null);
        if (def->propagate_errors) {
          return err;
        }
        values.push_back(err);
        continue;
      }
      had_range_shaped_arg = true;
      parser::Reference isect_lhs{};
      parser::Reference isect_rhs{};
      isect_lhs.sheet = isect_sheet;
      isect_lhs.row = isect_r1;
      isect_lhs.col = isect_c1;
      isect_rhs.sheet = isect_sheet;
      isect_rhs.row = isect_r2;
      isect_rhs.col = isect_c2;
      auto isect_expanded = ctx.expand_range(isect_lhs, isect_rhs, arena, registry);
      if (!isect_expanded) {
        const Value err = Value::error(isect_expanded.error());
        if (def->propagate_errors) {
          return err;
        }
        values.push_back(err);
        continue;
      }
      Value range_err = Value::blank();
      const std::vector<Value>& isect_values = isect_expanded.value();
      if (!append_range_sourced_values(*def, isect_values.data(), isect_values.size(), &values, &range_err)) {
        return range_err;
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
    bool expanded_call_handled = false;
    Value expanded_call_return = Value::blank();
    if (!append_expanded_call_argument(*def, arg_node, "OFFSET", expand_offset_call, arena, registry, ctx, &values,
                                       &expanded_call_handled, &expanded_call_return)) {
      return expanded_call_return;
    }
    if (expanded_call_handled) {
      had_range_shaped_arg = true;
      continue;
    }
    // IF-as-range-producer mirrors the CHOOSE / OFFSET branches: when an
    // aggregator receives `IF(cond, range1, range2)`, the picked branch
    // must be flattened to a vector of cells. This is what makes
    // `=LET(r, IF(TRUE, A1:A3, B1:B3), SUM(r))` aggregate the 3-cell
    // range rather than collapse `r` to a scalar. `expand_if_call` shares
    // the same evaluation / range-resolution / filter contracts as the
    // CHOOSE / OFFSET expanders.
    if (!append_expanded_call_argument(*def, arg_node, "IF", expand_if_call, arena, registry, ctx, &values,
                                       &expanded_call_handled, &expanded_call_return)) {
      return expanded_call_return;
    }
    if (expanded_call_handled) {
      had_range_shaped_arg = true;
      continue;
    }
    // CHOOSE-as-range-producer mirrors the OFFSET branch above: when an
    // aggregator receives `CHOOSE(idx, range1, range2, ...)`, the chosen
    // child must be flattened to a vector of cells (recursively, if it is
    // itself an OFFSET / CHOOSE call). `expand_choose_call` shares the
    // same evaluation, range-resolution, and filter contracts so SUM /
    // AVERAGE / MIN / MAX / AVERAGEA all behave consistently.
    if (!append_expanded_call_argument(*def, arg_node, "CHOOSE", expand_choose_call, arena, registry, ctx, &values,
                                       &expanded_call_handled, &expanded_call_return)) {
      return expanded_call_return;
    }
    if (expanded_call_handled) {
      had_range_shaped_arg = true;
      continue;
    }
    // ROW(range) / COLUMN(range) as an aggregator argument: Excel 365 spills
    // them to `{1;2;3;...}` / `{1,2,3,...}` and the surrounding aggregator
    // iterates the spill. Without a `Value::Array` runtime the scalar path
    // collapses to the rectangle's first row / column; this branch unpacks
    // the indices directly so `=SUM(ROW(A1:A5))` aggregates 15 rather than
    // the scalar 1. Mirrors the OFFSET / CHOOSE / IF branches above; the
    // emitted cells are always `Number`, so `range_filter_*` rules pass
    // them through unchanged.
    if (!append_expanded_call_argument(*def, arg_node, "ROW", expand_row_call, arena, registry, ctx, &values,
                                       &expanded_call_handled, &expanded_call_return)) {
      return expanded_call_return;
    }
    if (expanded_call_handled) {
      had_range_shaped_arg = true;
      continue;
    }
    if (def->accepts_ranges && arg_node.kind() == parser::NodeKind::StructuredRef) {
      // Structured (table) reference argument: evaluate it and unpack the
      // resulting Array row-major into the values vector. Mirrors the
      // SpillRef / RangeOp branches so SUM / AVERAGE / COUNTIF / ... see
      // the rectangle's cells exactly as if `Table[Col]` had been written
      // as a literal `A2:A10`. Errors short-circuit per `propagate_errors`.
      // Scalar (single-cell) results fall through to the generic argument
      // path; the only effect is that the per-cell `range_filter_*` rules
      // are not applied, which matches Mac for a single-cell `Sales[@Col]`
      // (Excel never broadcasts a single-cell structured ref through the
      // range-filter pipeline). The array-shaped path applies the filters
      // exactly like the RangeOp / SpillRef branches.
      Value sr_val = eval_node(arg_node, arena, registry, ctx);
      if (def->propagate_errors && sr_val.is_error()) {
        return sr_val;
      }
      if (sr_val.is_array()) {
        had_range_shaped_arg = true;
        const ArrayValue* a = sr_val.as_array();
        const std::size_t sr_total = static_cast<std::size_t>(a->rows) * static_cast<std::size_t>(a->cols);
        Value range_err = Value::blank();
        if (!append_range_sourced_values(*def, a->cells, sr_total, &values, &range_err)) {
          return range_err;
        }
        continue;
      }
      // Scalar single-cell result: push without range-filtering. Falling
      // through to the bottom-of-loop scalar handling would re-eval the
      // node; the value we already have is correct.
      values.push_back(sr_val);
      continue;
    }
    if (!append_expanded_call_argument(*def, arg_node, "COLUMN", expand_column_call, arena, registry, ctx, &values,
                                       &expanded_call_handled, &expanded_call_return)) {
      return expanded_call_return;
    }
    if (expanded_call_handled) {
      had_range_shaped_arg = true;
      continue;
    }
    Value v = eval_node(arg_node, arena, registry, ctx);
    if (def->propagate_errors && v.is_error()) {
      return v;
    }
    // Generic Array-result flatten for range-aware aggregators. Lazy
    // builtins (`ANCHORARRAY`, `SEQUENCE`, `TRANSPOSE`, `MUNIT`, ...)
    // return `Value::Array`; without this branch SUM / AVERAGE / MIN /
    // MAX would receive the Array as a single scalar slot and fail with
    // `#VALUE!` from `coerce_to_number`. Flattening row-major mirrors the
    // SpillRef / RangeOp branches above, applying the same provenance
    // filters so blank / text / bool cells are dropped or coerced
    // consistently. Calls whose specific shape we already special-cased
    // (OFFSET, IF, CHOOSE, INDEX, INDIRECT, ROW, COLUMN, TRANSPOSE-via-
    // SUMPRODUCT) reach `continue` before getting here, so this branch
    // only catches the otherwise-uncovered Array-returning calls.
    if (def->accepts_ranges && v.is_array()) {
      had_range_shaped_arg = true;
      const ArrayValue* arr = v.as_array();
      const std::size_t n = static_cast<std::size_t>(arr->rows) * static_cast<std::size_t>(arr->cols);
      Value range_err = Value::blank();
      if (!append_range_sourced_values(*def, arr->cells, n, &values, &range_err)) {
        return range_err;
      }
      continue;
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
      Value range_err = Value::blank();
      if (!append_range_sourced_value(*def, v, &values, &range_err)) {
        return range_err;
      }
      continue;
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

}  // namespace eval
}  // namespace formulon
