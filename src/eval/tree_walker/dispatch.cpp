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
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "eval/array_alloc.h"
#include "eval/coerce.h"
#include "eval/declared_rect.h"
#include "eval/defined_name_resolve.h"
#include "eval/dynamic_array/anchor.h"
#include "eval/eval_context.h"
#include "eval/external_ref.h"
#include "eval/function_registry.h"
#include "eval/lambda_value.h"
#include "eval/lazy_impls.h"
#include "eval/name_env.h"
#include "eval/name_env_resolve.h"
#include "eval/omitted_arg.h"
#include "eval/range_args.h"
#include "eval/range_expanders.h"
#include "eval/range_resolvers.h"
#include "eval/tail_array.h"
#include "eval/tree_walker/broadcast.h"
#include "eval/tree_walker/depth_guard.h"
#include "eval/tree_walker/dispatch_internal.h"
#include "eval/tree_walker_lazy_table.h"
#include "parser/ast.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/strings.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {

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

bool is_reference_shape(const parser::AstNode& node) noexcept {
  switch (node.kind()) {
    case parser::NodeKind::Ref:
    case parser::NodeKind::Ref3D:
    case parser::NodeKind::RangeOp:
    case parser::NodeKind::UnionOp:
    case parser::NodeKind::IntersectOp:
    case parser::NodeKind::SpillRef:
      return true;
    default:
      return false;
  }
}

bool is_full_axis_range(const parser::Reference& lhs, const parser::Reference& rhs) {
  const Expected<DeclaredRect, ErrorCode> rect = declared_rect(lhs, rhs);
  return rect && (rect.value().rows() == Sheet::kMaxRows || rect.value().cols() == Sheet::kMaxCols);
}

namespace {

// The body of `dispatch_call`. A scalar function applied over a whole column /
// row returns through `*tail_out` (the returned value is then unused).
Value dispatch_call_impl(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx, const TailArray** tail_out) {
  const std::string_view name = strip_future_prefix(node.as_call_name());
  const std::uint32_t arity = node.as_call_arity();

  // Name-bound lambda dispatch: when the formula text reads `f(x)` and `f`
  // resolves to a runtime `LambdaValue` via the lexical name environment
  // (LET-bound or LAMBDA-parameter), invoke it as if the user had written
  // an explicit IIFE. The lookup runs *before* the registry / lazy table so
  // a LET binding can shadow a built-in name (matching Excel's semantics).
  // A name bound to a reference is `#REF!`, as calling a cell is; any other
  // bound non-Lambda value is `#VALUE!`; a bound error propagates verbatim.
  // Unbound names fall through to the existing registry path.
  if (const NameEnv* env = ctx.name_env(); env != nullptr) {
    if (const auto* binding = env->lookup(name); binding != nullptr) {
      if (const parser::AstNode* ast = env->lookup_ast(name); ast != nullptr && is_reference_shape(*ast)) {
        return Value::error(ErrorCode::Ref);
      }
      const auto read = [&](const parser::AstNode& ref) { return eval_node_shaped(ref, arena, registry, ctx); };
      const Value bound = NameEnv::binding_value(*binding, arena, read);
      if (bound.is_lambda()) {
        return invoke_lambda(bound.as_lambda(), arity, node.as_call_args(), arena, registry, ctx);
      }
      if (bound.is_error()) {
        return bound;
      }
      return Value::error(ErrorCode::Value);
    }
  }

  // A visible defined name has precedence over the lazy table and the
  // function registry. This is the call-shaped counterpart of NameRef
  // resolution: a definition may evaluate to a LambdaValue and is then
  // invoked with the ordinary AST-backed lambda path; an error value is
  // propagated, while a successfully evaluated non-lambda is not callable.
  // `resolve_defined_name` returns the LambdaValue after its expansion frame
  // has been removed, so a named lambda may recurse by resolving itself on
  // each nested call without being mistaken for a circular defined-name
  // reference.
  if (find_defined_name(ctx, name) != nullptr) {
    const Value defined = resolve_defined_name(name, arena, registry, ctx);
    if (defined.is_error()) {
      return defined;
    }
    if (!defined.is_lambda()) {
      return Value::error(ErrorCode::Value);
    }
    return invoke_lambda(defined.as_lambda(), arity, node.as_call_args(), arena, registry, ctx);
  }

  if (LazyImpl lazy = find_lazy_impl(name); lazy != nullptr) {
    // Lazy callees read their arguments as rectangles, so whole-axis
    // references among them are walked to one shared length and agree on
    // shape however far each is populated. Eager callees flatten ranges and
    // have no shape to agree on.
    WholeAxisScope scope;
    scope.call = &node;
    return lazy(node, arena, registry, ctx.with_whole_axis_scope(&scope));
  }

  const FunctionDef* def = lookup_on_current_host(registry, name);
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
  // Parallel to `values`: the whole column / row a scalar function's argument
  // evaluated to, whose `values` slot is then a placeholder.
  std::vector<const TailArray*> tails;
  // Derived whole column / row arguments of a reducer with `FoldOps`.
  std::vector<StreamedTail> streamed;
  // Tracks whether any argument slot was range-shaped (RangeOp / OFFSET-call
  // / ArrayLiteral). The deferred `RejectAllScalarsBlank` policy surfaces
  // #VALUE! for `=GCD(A1,B1,C1)` (all blank
  // scalar refs) when there is NO range-shaped arg in the call. The mixed
  // form `=GCD(A1:B1, C1)` over the same blank cells still returns 0,
  // because the range arg "rescues" the policy.
  bool had_range_shaped_arg = false;
  // Tracks whether a scalar reference resolved to a blank under the
  // `RejectAllScalarsBlank` policy. Such a reference is omitted from the
  // callee's values; the deferred error fires only when it was the entire
  // scalar argument set.
  bool saw_blank_scalar_ref = false;
  // Set once a direct text literal that does not coerce to a number has been
  // seen by a numeric or A-family aggregator.
  bool literal_text_not_numeric = false;
  // The array a `last_arg_scalar` function's last argument evaluated to; the
  // call is then evaluated once per element.
  const ArrayValue* lifted_last_arg = nullptr;
  const std::uint32_t evaluated_arity = def->analysis_toolpak_args && !def->atp_omitted_optional_is_na
                                            ? atp_evaluated_arity(node, def->min_arity, def->max_arity)
                                            : arity;
  for (std::uint32_t i = 0; i < evaluated_arity; ++i) {
    const parser::AstNode& raw_arg = node.as_call_arg(i);
    if (def->last_arg_scalar && i > 0U && i + 1U == evaluated_arity) {
      const Value v = densify(eval_node_shaped(raw_arg, arena, registry, ctx), arena);
      if (def->propagate_errors && v.is_error()) {
        return literal_text_not_numeric ? Value::error(ErrorCode::Value) : v;
      }
      if (v.is_array()) {
        lifted_last_arg = v.as_array();
      }
      values.push_back(v);
      continue;
    }
    // LET-binding passthrough: when the caller wrote `SUM(r)` where `r` is
    // bound to a RangeOp / ArrayLiteral / OFFSET / CHOOSE / INDIRECT, the
    // dispatcher must see the underlying range AST or it would fall back
    // to the scalar path and collapse the binding to its spill anchor.
    // Substitute only when the resolved AST is genuinely range-shaped so
    // that a NameRef bound to a scalar (or a single-cell Ref) continues to
    // flow through the existing scalar branch with its original provenance.
    const parser::AstNode& arg_node = resolve_range_binding(raw_arg, ctx.name_env(), /*accept_ref=*/false);
    Value range_arg_result = Value::blank();
    switch (collect_range_arg(*def, arg_node, arena, registry, ctx, values, had_range_shaped_arg, range_arg_result)) {
      case RangeArgOutcome::kCollected:
        continue;
      case RangeArgOutcome::kReturn:
        return range_arg_result;
      case RangeArgOutcome::kNotRange:
        break;
    }
    // A scalar function keeps a whole column / row at its declared size for
    // the element-wise broadcast below; a range-aware one takes it expanded.
    const Shaped shaped_arg = eval_node_shaped(arg_node, arena, registry, ctx);
    if (shaped_arg.tail_array != nullptr && !def->accepts_ranges) {
      tails.resize(values.size() + 1U, nullptr);
      tails[values.size()] = shaped_arg.tail_array;
      values.push_back(Value::blank());
      continue;
    }
    // A whole column / row read straight from a reference (a defined name
    // over `A:A`) passes only its populated head, as `SUM(A:A)` does; a
    // derived one (`A:A*1`) is expanded below.
    if (shaped_arg.tail_array != nullptr && shaped_arg.tail_array->from_reference) {
      had_range_shaped_arg = true;
      const TailArray& ta = *shaped_arg.tail_array;
      const std::size_t span = ta.axis == TailAxis::kRows ? ta.cols : ta.rows;
      Value range_err = Value::blank();
      if (!append_range_sourced_values(*def, ta.cells, static_cast<std::size_t>(ta.head) * span, &values, &range_err)) {
        return range_err;
      }
      continue;
    }
    // A reducer that takes values one at a time streams a derived one
    // (`-A:BB`), which can hold more cells than an array may.
    if (shaped_arg.tail_array != nullptr && def->accepts_ranges && def->fold != nullptr) {
      had_range_shaped_arg = true;
      if (def->propagate_errors) {
        const Value err = first_tail_error(*shaped_arg.tail_array);
        if (err.is_error()) {
          return err;
        }
      }
      streamed.push_back(StreamedTail{values.size(), shaped_arg.tail_array});
      continue;
    }
    Value v = densify(shaped_arg, arena);
    if (def->propagate_errors && v.is_error()) {
      // A text literal that cannot be a number already failed ahead of this
      // error, and the earlier failure is the one reported.
      return literal_text_not_numeric ? Value::error(ErrorCode::Value) : v;
    }
    if ((def->range_filter_numeric_only || def->range_filter_a_coerce) &&
        arg_node.kind() == parser::NodeKind::Literal && v.is_text() && !coerce_to_number(v)) {
      literal_text_not_numeric = true;
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
      if (!append_array_sourced_values(*def, arr->cells, n, &values, &range_err)) {
        return range_err;
      }
      continue;
    }
    // RangeOp / OFFSET-call / ArrayLiteral args were handled above and
    // `continue`'d, so reaching this point implies a scalar arg slot
    // (Literal, Ref, BinaryOp, ...). The Analysis-ToolPak rule fires
    // eagerly so the left-most offending slot wins.
    if (def->analysis_toolpak_args) {
      const bool required = def->atp_omitted_optional_is_na || atp_slot_required(i, def->min_arity, def->max_arity);
      const Value err = atp_arg_error(arg_node, v, required);
      if (err.is_error()) {
        if (!absorb_arg_error(*def, err, &values)) {
          return err;
        }
        continue;
      }
    }
    // `RejectAllScalarsBlank` (GCD / LCM) omits blank scalar references so
    // a blank mixed with a numeric literal or reference cannot become zero
    // inside the callee. Range-shaped blanks are handled above and remain
    // zero, preserving the range form's semantics.
    if (def->blank_scalar_policy == FunctionDef::BlankScalarPolicy::RejectAllScalarsBlank &&
        arg_node.kind() == parser::NodeKind::Ref && v.kind() == ValueKind::Blank) {
      saw_blank_scalar_ref = true;
      continue;
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
    // An omitted slot of a range-aware aggregator is a counted 0 (COUNTA(3,,4) = 3).
    if (def->accepts_ranges && v.kind() == ValueKind::Blank && is_omitted_arg(arg_node)) {
      values.push_back(Value::blank(BlankGridProjection::kValueArrayZero));
      continue;
    }
    values.push_back(v);
  }
  // Deferred fire-point for `RejectAllScalarsBlank`: only surface the error
  // when blank scalar references were the entire scalar argument set. A
  // numeric literal or reference mixed with a blank therefore proceeds with
  // the blank omitted, while any range-shaped argument keeps its existing
  // blank-as-zero behavior.
  if (def->blank_scalar_policy == FunctionDef::BlankScalarPolicy::RejectAllScalarsBlank && saw_blank_scalar_ref &&
      values.empty() && !had_range_shaped_arg) {
    return Value::error(def->blank_scalar_error);
  }
  // Dynamic-array spill for scalar (non-range-aware) functions: when any
  // collected argument is an Array — a bounded range materialised above,
  // or a nested array-returning call (`=ROUND(SEQUENCE(3),0)`) — Excel 365
  // evaluates the function element-wise over the broadcast rectangle and
  // spills the result. Range-aware aggregators already flattened their
  // array arguments into `values`, so they never take this path.
  if (!def->accepts_ranges) {
    if (!tails.empty()) {
      tails.resize(values.size(), nullptr);
      const Shaped result = broadcast_scalar_call_shaped(*def, values, tails, arena);
      *tail_out = result.tail_array;
      return result.value;
    }
    std::uint32_t out_rows = 1;
    std::uint32_t out_cols = 1;
    bool any_array = false;
    for (const Value& v : values) {
      if (v.is_array()) {
        any_array = true;
        const ArrayValue* a = v.as_array();
        out_rows = std::max(out_rows, a->rows);
        out_cols = std::max(out_cols, a->cols);
      }
    }
    if (any_array) {
      return broadcast_scalar_call(*def, values, out_rows, out_cols, arena);
    }
  }
  if (!streamed.empty()) {
    return fold_with_tails(*def, values, streamed);
  }
  if (lifted_last_arg != nullptr) {
    return lift_last_arg_call(*def, std::move(values), *lifted_last_arg, arena);
  }
  // Hand the post-expansion size to the impl; aggregator bodies walk the
  // flattened vector directly.
  return def->impl(values.data(), static_cast<std::uint32_t>(values.size()), arena);
}

}  // namespace

Shaped dispatch_call(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  // A whole-axis scope belongs to the arguments of the call that installed
  // it; a nested call starts from its own references.
  if (ctx.whole_axis_scope() != nullptr) {
    return dispatch_call(node, arena, registry, ctx.with_whole_axis_scope(nullptr));
  }
  Shaped out;
  out.value = dispatch_call_impl(node, arena, registry, ctx, &out.tail_array);
  return out;
}

}  // namespace eval
}  // namespace formulon
