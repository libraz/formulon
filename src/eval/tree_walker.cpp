// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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

#include "eval/coerce.h"
#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/iterative_solver.h"
#include "eval/lambda_value.h"
#include "eval/lazy_impls.h"
#include "eval/name_env.h"
#include "eval/name_env_resolve.h"
#include "eval/range_args.h"
#include "eval/range_expanders.h"
#include "eval/range_resolvers.h"
#include "eval/scalar_ops.h"
#include "eval/structured_ref.h"
#include "eval/tree_walker_lazy_table.h"
#include "parser/ast.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/strings.h"
#include "value.h"
#include "workbook.h"

#ifdef FORMULON_VM_PARITY_CHECK
#include <cstdio>
#include <cstring>

#include "eval/compiler.h"
#include "eval/vm.h"
#endif

namespace formulon {
namespace eval {
namespace {

// Returns true when a dynamic-array result anchored at `(anchor_row,
// anchor_col)` with the given shape would collide with an already-occupied
// cell, in which case Excel surfaces `#SPILL!` at the anchor instead of
// spilling. A cell other than the anchor is "occupied" when it carries a
// formula (even one evaluating to `""`) or a non-blank cached value; a
// genuinely blank cell does not block.
//
// This is the read-only counterpart of `Sheet::commit_spill`'s collision
// scan: the production recalc path commits through `commit_spill` (which
// also clears any stale phantom region first), so this helper only runs on
// the read-only evaluation path where no spill is committed.
bool spill_would_collide(const Sheet& sheet, std::uint32_t anchor_row, std::uint32_t anchor_col, std::uint32_t rows,
                         std::uint32_t cols) {
  if (static_cast<std::uint64_t>(anchor_row) + rows > Sheet::kMaxRows ||
      static_cast<std::uint64_t>(anchor_col) + cols > Sheet::kMaxCols) {
    return true;
  }
  for (std::uint32_t r = 0; r < rows; ++r) {
    for (std::uint32_t c = 0; c < cols; ++c) {
      const std::uint32_t row = anchor_row + r;
      const std::uint32_t col = anchor_col + c;
      if (row == anchor_row && col == anchor_col) {
        continue;
      }
      const Cell* cell = sheet.cell_at(row, col);
      if (cell == nullptr) {
        continue;
      }
      if (!cell->formula_text.empty() || !cell->cached_value.is_blank()) {
        return true;
      }
    }
  }
  return false;
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
//   INDIRECT                                   -> src/eval/reference/indirect.cpp
//   OFFSET / expand_*_call                     -> src/eval/reference/offset.cpp
//   resolve_*_call / compute_intersect_rect    -> src/eval/reference/intersection.cpp
//   parse_a1_ref + shared helpers              -> src/eval/reference/common.cpp
//   CORREL / COVARIANCE.P / COVARIANCE.S /
//   SLOPE / INTERCEPT / RSQ / FORECAST.LINEAR /
//   STEYX / SUMX2PY2 / SUMX2MY2 / SUMXMY2      -> src/eval/regression_lazy.cpp
//   SERIESSUM                                  -> src/eval/series_sum_lazy.cpp
//   RANK / RANK.EQ / RANK.AVG /
//   PERCENTRANK / PERCENTRANK.INC /
//   PERCENTRANK.EXC                            -> src/eval/rank_lazy.cpp
//   PERCENTOF                                  -> src/eval/builtins/aggregate.cpp
//   AGGREGATE                                  -> src/eval/aggregate_lazy.cpp
//   REGEXTEST / REGEXEXTRACT / REGEXREPLACE    -> src/eval/regex_lazy.cpp
// Each family publishes its externs via its own header
// (`eval/special_forms_lazy.h`, `eval/conditional_aggregates.h`,
// `eval/lookups/classic.h`, `eval/lookups/xlookup.h`,
// `eval/shape_ops_lazy.h`, `eval/workdays_lazy.h`), which the dispatch
// table below includes.

// The lazy dispatch table itself lives in `tree_walker_lazy_table.cpp` so
// adding a new lazy entry does not force a rebuild of this evaluator TU.
// `find_lazy_impl` and `lazy_table_names` are the only seam this file
// uses to reach the per-family routing decisions.

// Hard caps on recursion depth. The parser already enforces a parse-depth
// limit of 128 (see `parser::ParserOptions::max_parse_depth`), which bounds
// stack growth from a single formula. These two evaluator-side caps defend
// against the orthogonal vectors that bypass that bound:
//
//   * `kMaxEvalDepth` — bounds linear cell-chain recursion through
//     `EvalContext::resolve_ref`. A workbook of `A1=A2, A2=A3, ..., A1000=1`
//     is not a cycle (so `EvalState::push_cell` does not flag it), and each
//     resolved formula spawns a fresh `eval_node` recursion. Without this
//     cap a chain of ~1000 cells overflows the WASM 256-512 KB stack.
//
//   * `kMaxLambdaDepth` — bounds runtime recursion through user-defined
//     LAMBDA closures (e.g. `LET(f, LAMBDA(n, f(n+1)), f(0))`). The body
//     AST stays small so `kMaxEvalDepth` does not trigger; the recursion
//     lives in `invoke_lambda` re-entering itself.
//
// On overflow the offending sub-expression returns `#CALC!` (the same
// Excel-visible code Mac Excel surfaces for indeterminate / runaway lambda
// recursion). The internal `kEvalStackOverflow` code is reserved for the
// `Expected<T, Error>` plumbing (currently unused on this path).
constexpr std::uint32_t kMaxEvalDepth = 512;
constexpr std::uint32_t kMaxLambdaDepth = 256;

// RAII guard: bumps `*p` on construction (when `*p < cap`) and decrements
// on destruction. When the cap was already reached, `exceeded()` reports
// `true` and the counter is left untouched so a later sibling in the same
// frame does not double-decrement past zero. Null `p` disables tracking
// entirely — `exceeded()` always returns `false` — which preserves
// behaviour for ad-hoc callers that bypass `evaluate()`.
class EvalDepthGuard {
 public:
  EvalDepthGuard(std::uint32_t* p, std::uint32_t cap) noexcept : p_(p), exceeded_(p != nullptr && *p >= cap) {
    if (p_ != nullptr && !exceeded_) {
      ++(*p_);
    }
  }
  ~EvalDepthGuard() noexcept {
    if (p_ != nullptr && !exceeded_) {
      --(*p_);
    }
  }
  EvalDepthGuard(const EvalDepthGuard&) = delete;
  EvalDepthGuard& operator=(const EvalDepthGuard&) = delete;
  EvalDepthGuard(EvalDepthGuard&&) = delete;
  EvalDepthGuard& operator=(EvalDepthGuard&&) = delete;

  bool exceeded() const noexcept { return exceeded_; }

 private:
  std::uint32_t* p_;
  bool exceeded_;
};

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

// Invokes a runtime LambdaValue with the given argument-AST accessor and
// arity. Shared between the `LambdaCall` AST case (a parser-emitted IIFE or
// curried call) and the name-bound dispatch path in `dispatch_call` (where
// the user wrote `f(x)` and `f` resolves through `NameEnv` to a Lambda).
//
// Arity check: required slots = `param_count - optional_count`; the call
// must satisfy `required <= arity <= param_count`. Anything else surfaces
// `#VALUE!`. Trailing optional slots that the caller did not supply bind
// to an "omitted" sentinel that ISOMITTED detects via `lookup_omitted`.
// Argument evaluation is eager and left-to-right in the *caller's* scope;
// the first error short-circuits. Bindings extend a fresh `NameEnv` rooted
// at the lambda's captured environment so closure capture continues to
// work across both invocation paths.
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

// ---------------------------------------------------------------------------
// Top-level array broadcasting for BinaryOp / UnaryOp
// ---------------------------------------------------------------------------
//
// `eval_binop_array_ctx` in shape_ops_lazy.cpp re-walks the AST under SUMPRODUCT
// and other array-context callers; the helpers below take *already evaluated*
// `Value`s so the top-level dispatch can broadcast over arrays produced by
// SpillRef / TRANSPOSE / SEQUENCE without a second AST walk. Shape rules
// mirror eval_binop_array_ctx exactly:
//
//   * Both operands 1x1                 -> 1x1 result
//   * Either operand 1x1, other R x C   -> R x C result, scalar broadcasts
//   * Otherwise dimensions must match   -> mismatch surfaces scalar #VALUE!
//     (matches Mac Excel's whole-expression short-circuit; it does NOT spill
//     a sea of #VALUE! cells)
//
// Per-cell error short-circuit: if either operand cell is an Error, that
// error is written verbatim into the result cell (left-most wins on ties via
// the lhs-first check in the loop). The caller decides whether the resulting
// Value::Array spills (via `EvalContext::dispatch_array_result`) or feeds
// into a downstream operator that consumes it as an Array.
struct ArrayView {
  std::uint32_t rows;
  std::uint32_t cols;
  const Value* cells;
};

// Resolves `v` to an ArrayView. For an Array the view aliases the existing
// cells buffer (no copy). For a scalar the caller-supplied 1-element backing
// slot `scalar_slot` is populated and aliased. Lifetime: the view is valid
// as long as either the source Array or `scalar_slot` outlives it.
ArrayView as_array_view(const Value& v, Value* scalar_slot) {
  if (v.is_array()) {
    const ArrayValue* a = v.as_array();
    return {a->rows, a->cols, a->cells};
  }
  *scalar_slot = v;
  return {1U, 1U, scalar_slot};
}

Value apply_binop_per_cell(parser::BinOp op, const Value& lhs, const Value& rhs, Arena& arena) {
  if (lhs.is_error()) {
    return lhs;
  }
  if (rhs.is_error()) {
    return rhs;
  }
  switch (op) {
    case parser::BinOp::Add:
    case parser::BinOp::Sub:
    case parser::BinOp::Mul:
    case parser::BinOp::Div:
    case parser::BinOp::Pow: {
      auto ln = coerce_to_number(lhs);
      if (!ln) {
        return Value::error(ln.error());
      }
      auto rn = coerce_to_number(rhs);
      if (!rn) {
        return Value::error(rn.error());
      }
      return apply_arithmetic(op, ln.value(), rn.value());
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

Value broadcast_binop(parser::BinOp op, const Value& lhs, const Value& rhs, Arena& arena) {
  Value l_slot = Value::blank();
  Value r_slot = Value::blank();
  const ArrayView la = as_array_view(lhs, &l_slot);
  const ArrayView ra = as_array_view(rhs, &r_slot);

  std::uint32_t out_rows = la.rows;
  std::uint32_t out_cols = la.cols;
  bool l_broadcast = false;
  bool r_broadcast = false;
  if (la.rows == 1U && la.cols == 1U) {
    out_rows = ra.rows;
    out_cols = ra.cols;
    l_broadcast = true;
  } else if (ra.rows == 1U && ra.cols == 1U) {
    r_broadcast = true;
  } else if (la.rows != ra.rows || la.cols != ra.cols) {
    return Value::error(ErrorCode::Value);
  }

  const std::size_t n = static_cast<std::size_t>(out_rows) * static_cast<std::size_t>(out_cols);
  Value* buf = arena.create_array<Value>(n);
  if (buf == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  for (std::size_t i = 0; i < n; ++i) {
    const Value& lv = l_broadcast ? la.cells[0] : la.cells[i];
    const Value& rv = r_broadcast ? ra.cells[0] : ra.cells[i];
    buf[i] = apply_binop_per_cell(op, lv, rv, arena);
  }
  ArrayValue* out = arena.create<ArrayValue>();
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  out->rows = out_rows;
  out->cols = out_cols;
  out->cells = buf;
  return Value::array(out);
}

Value broadcast_unary(parser::UnaryOp op, const Value& operand, Arena& arena) {
  if (!operand.is_array()) {
    return apply_unary(op, operand);
  }
  const ArrayValue* in = operand.as_array();
  const std::size_t n = static_cast<std::size_t>(in->rows) * static_cast<std::size_t>(in->cols);
  Value* buf = arena.create_array<Value>(n);
  if (buf == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  for (std::size_t i = 0; i < n; ++i) {
    const Value& cell = in->cells[i];
    buf[i] = cell.is_error() ? cell : apply_unary(op, cell);
  }
  ArrayValue* out = arena.create<ArrayValue>();
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  out->rows = in->rows;
  out->cols = in->cols;
  out->cells = buf;
  return Value::array(out);
}

}  // namespace

// Public entry point declared in `eval/tree_walker.h`. Routes through
// the lazy-table seam; the actual array is owned by
// `tree_walker_lazy_table.cpp`.
const char* const* lazy_form_names() {
  return lazy_table_names();
}

// Defined with external linkage (declared in `eval/lazy_impls.h`) so the
// per-family lazy-impl TUs can recurse into the evaluator. The scalar
// operator helpers it calls below — `apply_unary`, `apply_arithmetic`,
// `apply_concat`, `apply_comparison` — live in `eval/scalar_ops.h` and are
// reachable via ordinary unqualified lookup. `dispatch_call` lives in the
// anonymous namespace above and remains reachable for the same reason
// (the anonymous namespace is nested inside `formulon::eval`).
Value eval_node(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx) {
  // Bounds linear cell-chain recursion through `EvalContext::resolve_ref`
  // and any other re-entrant evaluator path. See `kMaxEvalDepth`.
  EvalDepthGuard depth_guard(ctx.eval_depth_counter(), kMaxEvalDepth);
  if (depth_guard.exceeded()) {
    return Value::error(ErrorCode::Calc);
  }
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

    case parser::NodeKind::UnaryOp: {
      // Eager scalar unary; broadcast cellwise when the operand evaluates to
      // a Value::Array (e.g. `=-A1#`). The Array result then bubbles up to
      // the cell entry point where dispatch_array_result decides whether to
      // commit a spill.
      const Value operand = eval_node(node.as_unary_operand(), arena, registry, ctx);
      if (operand.is_error()) {
        return operand;
      }
      return broadcast_unary(node.as_unary_op(), operand, arena);
    }

    case parser::NodeKind::BinaryOp: {
      const parser::BinOp op = node.as_binary_op();
      // Evaluate left first so error propagation honours the documented
      // left-most-wins rule.
      const Value lhs = eval_node(node.as_binary_lhs(), arena, registry, ctx);
      if (lhs.is_error()) {
        return lhs;
      }
      const Value rhs = eval_node(node.as_binary_rhs(), arena, registry, ctx);
      if (rhs.is_error()) {
        return rhs;
      }
      // Cellwise broadcast when either operand is an Array (SpillRef #,
      // TRANSPOSE, SEQUENCE, or any future array-producing builtin). Pure
      // scalar operands take a 1x1 fast path inside `broadcast_binop` and
      // return an Array of shape 1x1; we unwrap that to a scalar so the
      // common case keeps the same surface.
      if (lhs.is_array() || rhs.is_array()) {
        return broadcast_binop(op, lhs, rhs, arena);
      }
      return apply_binop_per_cell(op, lhs, rhs, arena);
    }

    case parser::NodeKind::Call:
      return dispatch_call(node, arena, registry, ctx);

    case parser::NodeKind::Ref:
      return ctx.resolve_ref(node.as_ref(), arena, registry);

    case parser::NodeKind::SpillRef: {
      // Excel's `=A1#` operator: yields the entire spill region anchored at
      // the referenced cell as a `Value::Array`. Resolution rules:
      //   * Unbound context (no current sheet)             -> #NAME?
      //   * Qualified anchor with no workbook bound        -> #REF!
      //   * Qualified anchor with missing target sheet     -> #REF!
      //   * Anchor row/col >= Sheet::kMax{Rows,Cols}       -> #REF!
      //   * No spill region anchored at this address       -> #REF!
      //
      // The returned ArrayValue header lives in the eval `arena`; its cells
      // buffer holds shallow copies of the SpillRegion's `Value`s. Text
      // payloads point into `SpillRegion::owned_strings`, which lives on
      // Sheet and outlives any single evaluation arena (zero-copy reuse).
      const parser::Reference& r = node.as_spill_ref();
      const Sheet* current = ctx.current_sheet();
      if (current == nullptr) {
        return Value::error(ErrorCode::Name);
      }
      const Sheet* target = current;
      if (!r.sheet.empty()) {
        const Workbook* wb = ctx.workbook();
        if (wb == nullptr) {
          return Value::error(ErrorCode::Ref);
        }
        target = wb->sheet_by_name(r.sheet);
        if (target == nullptr) {
          return Value::error(ErrorCode::Ref);
        }
      }
      if (r.row >= Sheet::kMaxRows || r.col >= Sheet::kMaxCols) {
        return Value::error(ErrorCode::Ref);
      }
      const SpillRegion* region = target->spill_region_at_anchor(r.row, r.col);
      if (region == nullptr) {
        return Value::error(ErrorCode::Ref);
      }
      const std::size_t n = static_cast<std::size_t>(region->rows) * static_cast<std::size_t>(region->cols);
      Value* buffer = arena.create_array<Value>(n);
      if (buffer == nullptr) {
        return Value::error(ErrorCode::Num);
      }
      for (std::size_t i = 0; i < n; ++i) {
        buffer[i] = region->cells[i];
      }
      ArrayValue* arr = arena.create<ArrayValue>();
      if (arr == nullptr) {
        return Value::error(ErrorCode::Num);
      }
      arr->rows = region->rows;
      arr->cols = region->cols;
      arr->cells = buffer;
      return Value::array(arr);
    }

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
        // Record the AST source for reference-shaped bindings: a bare
        // `Ref` (`=LET(r, A5, ROW(r))`), a `RangeOp`, an `ArrayLiteral`,
        // or one of the reference-producing calls (`OFFSET`, `CHOOSE`,
        // `INDIRECT`, `IF`). The `Ref` case keeps `ROW` / `COLUMN` able
        // to introspect a single-cell binding without re-evaluation;
        // SUM-family callers continue to skip `Ref` via the narrower
        // `is_range_shaped_ast` predicate so a scalar-bound name still
        // flows through the existing scalar branch. Truly scalar shapes
        // (literals, arithmetic) still bind by Value only — recording
        // their AST would risk re-evaluating side-effecting or expensive
        // sub-expressions on every NameRef read. NameRef-on-NameRef is
        // transitive: if the RHS already resolves to a reference-shaped
        // AST in the (possibly outer) scope, we inherit that AST so
        // `=LET(s, r, SUM(s))` works when `r` is itself a range binding.
        const parser::AstNode* expr_for_binding = nullptr;
        if (expr_node.kind() == parser::NodeKind::Ref || is_range_shaped_ast(expr_node)) {
          expr_for_binding = &expr_node;
        } else if (expr_node.kind() == parser::NodeKind::NameRef) {
          // `env` already reflects every previously bound name in this LET
          // (and, via `parent`, any outer LETs); a single lookup walks the
          // whole chain.
          expr_for_binding = env.lookup_ast(expr_node.as_name());
        }
        env = env.extend(node.as_let_binding_name(i), v, expr_for_binding, arena);
      }
      const EvalContext body_ctx = ctx.with_name_env(&env);
      return eval_node(node.as_let_body(), arena, registry, body_ctx);
    }

    // -- Unsupported: external names --------------------------------------
    case parser::NodeKind::ExternalRef:
      return Value::error(ErrorCode::Name);

    case parser::NodeKind::Ref3D: {
      // A 3-D reference (`Sheet2:Sheet3!A1`) denotes one cell across a span
      // of sheets — a range shape. In scalar context Excel cannot collapse
      // it to a single value, so it surfaces `#VALUE!`. A span endpoint that
      // names a missing sheet is `#REF!`, which takes priority. Range-aware
      // aggregators intercept this node in `dispatch_call` before reaching
      // here, so this branch only fires for true scalar context.
      const Workbook* wb = ctx.workbook();
      if (wb == nullptr) {
        return Value::error(ErrorCode::Ref);
      }
      const std::size_t begin_idx = wb->sheet_index_by_name(node.as_ref3d_sheet_begin());
      const std::size_t end_idx = wb->sheet_index_by_name(node.as_ref3d_sheet_end());
      if (begin_idx == static_cast<std::size_t>(-1) || end_idx == static_cast<std::size_t>(-1)) {
        return Value::error(ErrorCode::Ref);
      }
      return Value::error(ErrorCode::Value);
    }

    case parser::NodeKind::StructuredRef: {
      // Resolve the table reference (`Table[Col]`, `Table[#All]`, ...) to a
      // concrete rectangle on the table's home sheet, then read the cells
      // through the existing reference-resolution machinery. The bracket
      // payload was captured verbatim by the parser into the node's
      // `column` slot; the resolver re-parses it here so multi-specifier
      // and column-range forms can flow through a single AST shape.
      const Workbook* wb = ctx.workbook();
      const Sheet* current = ctx.current_sheet();
      if (wb == nullptr || current == nullptr) {
        return Value::error(ErrorCode::Name);
      }
      auto sel_or = parse_structured_ref_payload(node.as_structured_ref_column());
      if (!sel_or) {
        return Value::error(sel_or.error());
      }
      StructuredRefSelector sel = std::move(sel_or).value();
      sel.table_name = node.as_structured_ref_table();
      // Locate the current sheet's workbook index; falling back to 0 is fine
      // because the resolver only consults `current_sheet_index` for
      // future cross-sheet contracts (the row-implicit form uses the
      // formula cell's row directly).
      std::uint32_t current_sheet_index = 0;
      for (std::size_t i = 0; i < wb->sheet_count(); ++i) {
        if (&wb->sheet(i) == current) {
          current_sheet_index = static_cast<std::uint32_t>(i);
          break;
        }
      }
      const std::uint32_t current_row = ctx.has_formula_cell() ? ctx.formula_row() : EvalContext::kNoFormulaCell;
      auto rect_or = resolve_structured_ref(sel, *wb, current_sheet_index, current_row);
      if (!rect_or) {
        return Value::error(rect_or.error());
      }
      const StructuredRefRange rect = std::move(rect_or).value();
      // Build A1 references for the rectangle's two corners and let
      // `expand_range` do the actual cell fetch. This funnels structured
      // refs through the same cross-sheet / cycle-detecting path as a
      // literal `Sheet1!A1:C10` reference.
      parser::Reference lhs{};
      lhs.sheet = rect.sheet_name;
      lhs.row = rect.row_first;
      lhs.col = rect.col_first;
      parser::Reference rhs{};
      rhs.sheet = rect.sheet_name;
      rhs.row = rect.row_last;
      rhs.col = rect.col_last;
      // Single-cell rectangle: read the value directly, mirroring `Ref`.
      if (rect.row_first == rect.row_last && rect.col_first == rect.col_last) {
        return ctx.resolve_ref(lhs, arena, registry);
      }
      auto cells = ctx.expand_range(lhs, rhs, arena, registry);
      if (!cells) {
        return Value::error(cells.error());
      }
      const std::uint32_t rows = rect.row_last - rect.row_first + 1u;
      const std::uint32_t cols = rect.col_last - rect.col_first + 1u;
      const std::size_t total = static_cast<std::size_t>(rows) * static_cast<std::size_t>(cols);
      Value* buffer = arena.create_array<Value>(total);
      if (buffer == nullptr) {
        return Value::error(ErrorCode::Num);
      }
      for (std::size_t i = 0; i < total && i < cells.value().size(); ++i) {
        buffer[i] = cells.value()[i];
      }
      ArrayValue* arr = arena.create<ArrayValue>();
      if (arr == nullptr) {
        return Value::error(ErrorCode::Num);
      }
      arr->rows = rows;
      arr->cols = cols;
      arr->cells = buffer;
      return Value::array(arr);
    }

    case parser::NodeKind::Lambda: {
      // Build a runtime closure capturing the current name environment so a
      // body using outer LET-bound names (e.g. `LET(y, 100, LAMBDA(x, x+y))`)
      // sees `y` at call time even though the LET frame has gone out of
      // lexical scope. Param string_views are re-copied into the eval arena
      // even though the parser-arena view they reference also outlives the
      // value: the copy keeps the lifetime story uniform with the rest of
      // the LambdaValue payload.
      auto* lv = arena.create<LambdaValue>();
      if (lv == nullptr) {
        return Value::error(ErrorCode::Num);
      }
      const std::uint32_t n = node.as_lambda_param_count();
      std::string_view* params = nullptr;
      if (n > 0) {
        params = arena.create_array<std::string_view>(n);
        if (params == nullptr) {
          return Value::error(ErrorCode::Num);
        }
        for (std::uint32_t i = 0; i < n; ++i) {
          params[i] = node.as_lambda_param(i);
        }
      }
      lv->params = params;
      lv->param_count = n;
      lv->optional_count = node.as_lambda_optional_count();
      lv->body = &node.as_lambda_body();
      // Copy the caller's NameEnv into the arena: the live `NameEnv` value at
      // `ctx.name_env()` typically lives on a parent eval_node frame that
      // disappears once that frame returns, but every `Binding*` it points
      // at is arena-allocated and survives. Cloning the small wrapper struct
      // into the arena keeps the closure's reach into the binding chain
      // valid for the lifetime of the LambdaValue.
      if (const NameEnv* parent = ctx.name_env(); parent != nullptr) {
        auto* env_copy = arena.create<NameEnv>();
        if (env_copy == nullptr) {
          return Value::error(ErrorCode::Num);
        }
        *env_copy = *parent;
        lv->captured_env = env_copy;
      } else {
        lv->captured_env = nullptr;
      }
      return Value::lambda(lv);
    }

    case parser::NodeKind::LambdaCall: {
      // Evaluate the callee expression. Excel rejects calling a non-lambda
      // with #VALUE! (e.g. `(1+2)(3)` — when the parser admits the form).
      const Value callee = eval_node(node.as_lambda_call_callee(), arena, registry, ctx);
      if (callee.is_error()) {
        return callee;
      }
      if (!callee.is_lambda()) {
        return Value::error(ErrorCode::Value);
      }
      // Hand off to the shared invoker. Argument-AST accessors differ
      // between the `LambdaCall` case (here) and the name-bound dispatch
      // path in `dispatch_call`, so we materialise a flat pointer array
      // before invoking. ISOMITTED would change the eager-evaluation
      // contract once it lands (currently a service stub), but every
      // non-omitted argument case matches Excel.
      const std::uint32_t arity = node.as_lambda_call_arity();
      std::vector<const parser::AstNode*> argv;
      argv.reserve(arity);
      for (std::uint32_t i = 0; i < arity; ++i) {
        argv.push_back(&node.as_lambda_call_arg(i));
      }
      return invoke_lambda(callee.as_lambda(), arity, argv.empty() ? nullptr : argv.data(), arena, registry, ctx);
    }

    case parser::NodeKind::IntersectOp: {
      // Excel's space-as-intersection operator: `A1:C3 B2:D4` -> the
      // overlapping rectangle (here B2:C3). When the operands do not
      // overlap, Excel returns `#NULL!`. In scalar context (no formula
      // cell binding, or 2-D intersection rectangle) we collapse to the
      // top-left of the intersection rectangle, mirroring the RangeOp
      // case above.
      std::string_view sheet;
      std::uint32_t r1 = 0;
      std::uint32_t c1 = 0;
      std::uint32_t r2 = 0;
      std::uint32_t c2 = 0;
      bool disjoint = false;
      ErrorCode err = ErrorCode::Value;
      if (!compute_intersect_rect(node.as_intersect_lhs(), node.as_intersect_rhs(), arena, registry, ctx, &sheet, &r1,
                                  &c1, &r2, &c2, &disjoint, &err)) {
        return Value::error(err);
      }
      if (disjoint) {
        return Value::error(ErrorCode::Null);
      }
      // Implicit-intersection style alignment when the formula cell sits
      // inside the intersection rectangle's row or column band; otherwise
      // the spill anchor (top-left). Mirrors the RangeOp scalar policy.
      if (ctx.has_formula_cell()) {
        const std::uint32_t fr = ctx.formula_row();
        const std::uint32_t fc = ctx.formula_col();
        parser::Reference target{};
        target.sheet = sheet;
        if (c1 == c2 && fr >= r1 && fr <= r2) {
          target.row = fr;
          target.col = c1;
          return ctx.resolve_ref(target, arena, registry);
        }
        if (r1 == r2 && fc >= c1 && fc <= c2) {
          target.row = r1;
          target.col = fc;
          return ctx.resolve_ref(target, arena, registry);
        }
      }
      parser::Reference top_left{};
      top_left.sheet = sheet;
      top_left.row = r1;
      top_left.col = c1;
      return ctx.resolve_ref(top_left, arena, registry);
    }

    // -- Unsupported: range-producing operators / array literals ----------
    case parser::NodeKind::UnionOp:
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
  // Allocate the depth counters on this stack frame iff the inbound
  // context does not already carry them. `EvalContext::resolve_ref`
  // recursively re-enters `evaluate()` when a referenced cell is a
  // formula; if that re-entry reset the counters, a linear cell chain
  // (`A1=A2, A2=A3, ..., A1000=1`) would never trip the cap because
  // each link would start from zero. Preserving inherited counters lets
  // `kMaxEvalDepth` bound the cumulative recursion across the chain.
  // See `kMaxEvalDepth` / `kMaxLambdaDepth` for the policy.
  std::uint32_t eval_depth = 0;
  std::uint32_t lambda_depth = 0;
  const bool is_top_level = ctx.eval_depth_counter() == nullptr;
  EvalContext ctx_with_counters = ctx;
  if (is_top_level) {
    ctx_with_counters = ctx_with_counters.with_depth_counters(&eval_depth, &lambda_depth);
  }

  // Iterative-calculation driver. When the bound workbook has Excel's
  // "Enable iterative calculation" option on, a formula anchored at a known
  // cell is evaluated as a fixed-point iteration rather than once: each pass
  // runs against a fresh `EvalState` whose memo is pre-seeded with the
  // anchor cell's value from the previous pass, so a self-referential read
  // (`=IF(Z1>=5,5,Z1+1)` at Z1) resolves to that prior value instead of
  // recursing into a circular-reference `#REF!`. The loop stops as soon as
  // the absolute change between two passes drops below `max_change`, or
  // after `max_iterations` passes. A non-circular formula simply converges
  // on the second pass (zero delta). Only the top-level `evaluate()` call
  // drives the loop; nested re-entry (`resolve_ref`) keeps the ordinary
  // single-pass behaviour.
  if (is_top_level && !ctx.iterative_driver_suppressed() && ctx.has_formula_cell() && ctx.current_sheet() != nullptr &&
      ctx.workbook() != nullptr && ctx.workbook()->iterative_options().enabled) {
    const IterativeOptions& iopts = ctx.workbook()->iterative_options();
    const std::uint32_t max_iter = iopts.max_iterations == 0U ? 1U : iopts.max_iterations;
    const Sheet* anchor_sheet = ctx.current_sheet();
    const std::uint32_t anchor_row = ctx.formula_row();
    const std::uint32_t anchor_col = ctx.formula_col();
    // Excel seeds a fresh circular cell at 0 before the first pass.
    Value current = Value::number(0.0);
    for (std::uint32_t pass = 0; pass < max_iter; ++pass) {
      EvalState pass_state;
      // Seed the anchor cell so any self-reference resolves to the previous
      // pass's value without re-entrant evaluation.
      pass_state.memoize(anchor_sheet, anchor_row, anchor_col, current);
      EvalContext pass_ctx = ctx_with_counters.with_state(pass_state);
      Value next = eval_node(node, arena, registry, pass_ctx);
      // Apply the blank -> 0 surface contract so a blank-resolving pass is
      // comparable to a numeric one (matches the contract applied below for
      // the non-iterative path).
      if (next.is_blank() && node.kind() != parser::NodeKind::Literal) {
        next = Value::number(0.0);
      }
      // Convergence test: absolute change of the numeric value. A pass that
      // produces a non-number (or flips kind) is treated as not-yet-converged
      // so the loop keeps running until the cap.
      bool converged = false;
      if (next.is_number() && current.is_number()) {
        const double delta = std::fabs(next.as_number() - current.as_number());
        converged = delta < iopts.max_change;
      }
      current = next;
      if (converged) {
        break;
      }
    }
    return current;
  }

  Value v = eval_node(node, arena, registry, ctx_with_counters);
  // Dynamic-array spill-collision surface contract. When a 365-era formula
  // produces a multi-cell array and is anchored at a known formula cell on
  // a resolvable sheet, Excel reports `#SPILL!` at the anchor if any cell
  // the result would occupy (other than the anchor) is already non-empty.
  // The mutable-sheet recalc path commits through `Sheet::commit_spill`,
  // which runs the authoritative collision scan (and clears stale phantom
  // regions first); this read-only check covers the path where no spill is
  // committed (ad-hoc evaluation, the oracle harness) so a blocked spill
  // still surfaces `#SPILL!` rather than the bare anchor scalar.
  if (v.is_array() && ctx.mutable_sheet() == nullptr && ctx.has_formula_cell() && ctx.current_sheet() != nullptr) {
    const std::uint32_t rows = v.as_array_rows();
    const std::uint32_t cols = v.as_array_cols();
    if (rows > 0U && cols > 0U &&
        spill_would_collide(*ctx.current_sheet(), ctx.formula_row(), ctx.formula_col(), rows, cols)) {
      return Value::error(ErrorCode::Spill);
    }
  }
  // Excel surfaces a non-IIFE LAMBDA expression sitting at the top of a cell
  // formula as `#CALC!`: `=LAMBDA(x, x+1)` is a closure value with no
  // application site, so the cell renderer cannot project it onto a scalar.
  // The internal evaluator still produces (and consumes) Lambda values
  // happily — IIFE and LET-bound dispatch both rely on them — so we only
  // gate this surface contract at the top-level `evaluate()` boundary.
  // Sub-expression Lambdas (a callee subtree, a LET initialiser) do not
  // pass through here.
  if (v.is_lambda()) {
    return Value::error(ErrorCode::Calc);
  }
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
#ifdef FORMULON_VM_PARITY_CHECK
  // Bundle 5.2 parity harness. Compile the same AST through the bytecode
  // pipeline, run it through the VM, and compare the result. Mismatches are
  // surfaced as a diagnostic on stderr; we deliberately do NOT abort or
  // mutate the returned value so the production code path stays unchanged.
  // The parity sweep test reads the same `evaluate()` entry point and asserts
  // separately on its own corpus; this hook is the in-flight cross-check.
  {
    auto bc_or = compile(node, arena);
    if (bc_or) {
      auto vm_or = execute(bc_or.value(), arena, registry, ctx);
      if (vm_or) {
        const Value vm_v = vm_or.value();
        // Apply the same Lambda-at-top / blank->0 surface contracts the
        // tree-walker applies above so the two paths are compared on equal
        // terms (the VM does not re-apply these wrappers).
        Value vm_final = vm_v;
        if (vm_final.is_lambda()) {
          vm_final = Value::error(ErrorCode::Calc);
        }
        if (vm_final.is_blank() && node.kind() != parser::NodeKind::Literal) {
          vm_final = Value::number(0.0);
        }
        // Bit-exact equality: same kind, same payload. For Number we use
        // raw bit comparison so NaN payloads must agree exactly.
        bool eq = (v.kind() == vm_final.kind());
        if (eq) {
          switch (v.kind()) {
            case ValueKind::Number: {
              const double a = v.as_number();
              const double b = vm_final.as_number();
              std::uint64_t ua = 0;
              std::uint64_t ub = 0;
              std::memcpy(&ua, &a, sizeof(ua));
              std::memcpy(&ub, &b, sizeof(ub));
              eq = (ua == ub);
              break;
            }
            case ValueKind::Bool:
              eq = (v.as_boolean() == vm_final.as_boolean());
              break;
            case ValueKind::Error:
              eq = (v.as_error() == vm_final.as_error());
              break;
            case ValueKind::Text:
              eq = (v.as_text() == vm_final.as_text());
              break;
            case ValueKind::Array: {
              const ArrayValue* la = v.as_array();
              const ArrayValue* ra = vm_final.as_array();
              eq = (la->rows == ra->rows && la->cols == ra->cols);
              break;
            }
            case ValueKind::Blank:
            case ValueKind::Ref:
            case ValueKind::Lambda:
              break;
          }
        }
        if (!eq) {
          // Best-effort diagnostic: the harness does not have access to the
          // formula text here, so we only print the divergent kind / payload.
          std::fprintf(stderr, "[FORMULON_VM_PARITY] mismatch: tree=%s vm=%s\n", v.debug_to_string().c_str(),
                       vm_final.debug_to_string().c_str());
        }
      }
    }
  }
#endif  // FORMULON_VM_PARITY_CHECK
  return v;
}

}  // namespace eval
}  // namespace formulon
