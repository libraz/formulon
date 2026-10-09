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

namespace {

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

// Expands the rectangle [`lhs`, `rhs`] into `values` through `def`'s range
// filters. An expansion error is pushed, or returned via `immediate_return`
// (with `false`) when `def` propagates errors.
bool append_expanded_range(const FunctionDef& def, const parser::Reference& lhs, const parser::Reference& rhs,
                           Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                           std::vector<Value>* values, Value* immediate_return) {
  auto expanded = ctx.expand_range(lhs, rhs, arena, registry);
  if (!expanded) {
    const Value err = Value::error(expanded.error());
    if (def.propagate_errors) {
      *immediate_return = err;
      return false;
    }
    values->push_back(err);
    return true;
  }
  Value range_err = Value::blank();
  const std::vector<Value>& cells = expanded.value();
  if (!append_range_sourced_values(def, cells.data(), cells.size(), values, &range_err)) {
    *immediate_return = range_err;
    return false;
  }
  return true;
}

}  // namespace

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

// One argument slot for element-wise broadcasting of a scalar function
// over array-shaped arguments: either a scalar (`scalar != nullptr`) or
// an ArrayValue (`array != nullptr`).
struct BroadcastArg {
  const Value* scalar;
  const ArrayValue* array;
  std::uint32_t rows;
  std::uint32_t cols;
  const TailArray* tail = nullptr;  ///< Set instead of `array` for a whole column / row.
};

// Element of `arg` at output position (r, c) under Excel's 1xN / Nx1
// broadcast rules. A scalar supplies every position; an array whose
// non-1 extent is smaller than the output cannot supply the position and
// yields `#N/A`, matching Excel's ragged-broadcast fill.
Value broadcast_element(const BroadcastArg& arg, std::uint32_t r, std::uint32_t c) {
  if (arg.scalar != nullptr) {
    return *arg.scalar;
  }
  const std::uint32_t ri = arg.rows == 1U ? 0U : r;
  const std::uint32_t ci = arg.cols == 1U ? 0U : c;
  if (ri >= arg.rows || ci >= arg.cols) {
    return Value::error(ErrorCode::NA);
  }
  if (arg.tail != nullptr) {
    return tail_array_at(*arg.tail, ri, ci);
  }
  return arg.array->cells[static_cast<std::size_t>(ri) * arg.cols + ci];
}

// Evaluates a scalar (non-range-aware) function element-wise across the
// broadcast rectangle of its already-collected arguments and returns a
// spilled `Value::Array`. `out_rows` / `out_cols` are the max extents
// over the array arguments; scalars broadcast to every cell. A 1x1
// result unwraps to a plain scalar so a single-cell range argument does
// not spill. Per-cell error propagation mirrors the scalar dispatch
// path: with `propagate_errors`, the first error argument short-circuits
// that cell before the impl runs.
Value broadcast_scalar_call(const FunctionDef& def, const std::vector<Value>& args, std::uint32_t out_rows,
                            std::uint32_t out_cols, Arena& arena) {
  std::vector<BroadcastArg> views;
  views.reserve(args.size());
  for (const Value& v : args) {
    if (v.is_array()) {
      const ArrayValue* a = v.as_array();
      views.push_back({nullptr, a, a->rows, a->cols});
    } else {
      views.push_back({&v, nullptr, 1U, 1U});
    }
  }
  Value* cells = nullptr;
  ArrayValue* out = allocate_array_value(out_rows, out_cols, arena, cells, kMaxDerivedArrayCells);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  std::vector<Value> cell_args(views.size(), Value::blank());
  std::size_t idx = 0;
  for (std::uint32_t r = 0; r < out_rows; ++r) {
    for (std::uint32_t c = 0; c < out_cols; ++c, ++idx) {
      Value result = Value::blank();
      bool short_circuited = false;
      for (std::size_t a = 0; a < views.size(); ++a) {
        cell_args[a] = broadcast_element(views[a], r, c);
        if (def.propagate_errors && cell_args[a].is_error()) {
          result = cell_args[a];
          short_circuited = true;
          break;
        }
      }
      if (!short_circuited) {
        result = def.impl(cell_args.data(), static_cast<std::uint32_t>(cell_args.size()), arena);
      }
      cells[idx] = result;
    }
  }
  if (out_rows == 1U && out_cols == 1U) {
    return cells[0];
  }
  return Value::array(out);
}

// `broadcast_scalar_call` for arguments of which `tails[i]` (when non-null) is
// a whole column / row standing in for `args[i]`. Arguments on one axis give a
// `TailArray`: the function is applied to the head cells and once to the tail
// row / column. Whole columns mixed with whole rows are expanded and take the
// dense route.
Shaped broadcast_scalar_call_shaped(const FunctionDef& def, const std::vector<Value>& args,
                                    const std::vector<const TailArray*>& tails, Arena& arena) {
  bool by_rows = false;
  bool by_cols = false;
  for (const TailArray* ta : tails) {
    if (ta != nullptr) {
      (ta->axis == TailAxis::kRows ? by_rows : by_cols) = true;
    }
  }
  std::vector<BroadcastArg> views;
  views.reserve(args.size());
  std::uint32_t out_rows = 1;
  std::uint32_t out_cols = 1;
  std::vector<Value> dense_args;
  const bool mixed = by_rows && by_cols;
  if (mixed) {
    dense_args = args;
    for (std::size_t i = 0; i < args.size(); ++i) {
      if (tails[i] != nullptr) {
        Shaped s;
        s.tail_array = tails[i];
        dense_args[i] = densify(s, arena);
        if (dense_args[i].is_error()) {
          return Shaped{dense_args[i], nullptr};
        }
      }
    }
  }
  const std::vector<Value>& source = mixed ? dense_args : args;
  for (std::size_t i = 0; i < source.size(); ++i) {
    if (!mixed && tails[i] != nullptr) {
      views.push_back({nullptr, nullptr, tails[i]->rows, tails[i]->cols, tails[i]});
    } else if (source[i].is_array()) {
      const ArrayValue* a = source[i].as_array();
      views.push_back({nullptr, a, a->rows, a->cols});
    } else {
      views.push_back({&source[i], nullptr, 1U, 1U});
    }
    if (views.back().scalar == nullptr) {
      out_rows = std::max(out_rows, views.back().rows);
      out_cols = std::max(out_cols, views.back().cols);
    }
  }
  const auto eval_cell = [&](std::vector<Value>& cell_args, std::uint32_t r, std::uint32_t c) {
    for (std::size_t a = 0; a < views.size(); ++a) {
      cell_args[a] = broadcast_element(views[a], r, c);
      if (def.propagate_errors && cell_args[a].is_error()) {
        return cell_args[a];
      }
    }
    return def.impl(cell_args.data(), static_cast<std::uint32_t>(cell_args.size()), arena);
  };
  std::vector<Value> cell_args(views.size(), Value::blank());
  if (mixed) {
    Value* cells = nullptr;
    ArrayValue* out = allocate_array_value(out_rows, out_cols, arena, cells, kMaxDerivedArrayCells);
    if (out == nullptr) {
      return Shaped{Value::error(ErrorCode::Num), nullptr};
    }
    std::size_t idx = 0;
    for (std::uint32_t r = 0; r < out_rows; ++r) {
      for (std::uint32_t c = 0; c < out_cols; ++c, ++idx) {
        cells[idx] = eval_cell(cell_args, r, c);
      }
    }
    return Shaped{Value::array(out), nullptr};
  }

  // Positions past `head` read a constant from every argument: a tail, a
  // stretched size-1 axis, or the `#N/A` beyond a shorter dense array.
  const TailAxis axis = by_rows ? TailAxis::kRows : TailAxis::kCols;
  const bool along_rows = axis == TailAxis::kRows;
  const std::uint32_t extent = along_rows ? out_rows : out_cols;
  std::uint32_t head = 0;
  for (const BroadcastArg& v : views) {
    if (v.tail != nullptr) {
      head = std::max(head, v.tail->head);
    } else if (v.scalar == nullptr) {
      const std::uint32_t ext = along_rows ? v.rows : v.cols;
      if (ext > 1U) {
        head = std::max(head, ext);
      }
    }
  }
  head = std::min(head, extent);
  const std::size_t line = along_rows ? out_cols : out_rows;
  const std::uint64_t head_cells = static_cast<std::uint64_t>(head) * line;
  if (head_cells > kMaxDerivedArrayCells) {
    return Shaped{Value::error(ErrorCode::Num), nullptr};
  }
  Value* cells = head_cells == 0U ? nullptr : arena.create_array<Value>(static_cast<std::size_t>(head_cells));
  Value* tail = arena.create_array<Value>(line);
  if ((head_cells != 0U && cells == nullptr) || tail == nullptr) {
    return Shaped{Value::error(ErrorCode::Num), nullptr};
  }
  std::size_t idx = 0;
  for (std::uint32_t r = 0; r < (along_rows ? head : out_rows); ++r) {
    for (std::uint32_t c = 0; c < (along_rows ? out_cols : head); ++c, ++idx) {
      cells[idx] = eval_cell(cell_args, r, c);
    }
  }
  const std::uint32_t tail_at = std::min(head, extent - 1U);
  for (std::uint32_t k = 0; k < line; ++k) {
    tail[k] = along_rows ? eval_cell(cell_args, tail_at, k) : eval_cell(cell_args, k, tail_at);
  }
  Shaped out;
  out.tail_array = make_tail_array(arena, out_rows, out_cols, head, axis, cells, tail, /*from_reference=*/false);
  if (out.tail_array == nullptr) {
    return Shaped{Value::error(ErrorCode::Num), nullptr};
  }
  return out;
}

// Resolves a bounded `RangeOp`'s two endpoints and hands back the union
// rectangle as the corner Refs `expand_range` takes. Endpoints may be
// plain Refs (the simple `A1:B2` form), reference-producing calls
// (`OFFSET(...)` / `INDIRECT(...)`) or names standing for either; anything
// else surfaces per `resolve_range_endpoint`'s error code. Both corners
// carry the sheet `merge_range_endpoint_sheets` settles.
//
// Returns `false` with the endpoint's error code in `*out_err`.
bool union_endpoint_refs(const parser::AstNode& lhs_ast, const parser::AstNode& rhs_ast, Arena& arena,
                         const FunctionRegistry& registry, const EvalContext& ctx, parser::Reference* out_lhs,
                         parser::Reference* out_rhs, ErrorCode* out_err) {
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
  if (!resolve_range_endpoint(lhs_ast, arena, registry, ctx, &lhs_sheet, &lhs_top, &lhs_left, &lhs_bottom, &lhs_right,
                              out_err) ||
      !resolve_range_endpoint(rhs_ast, arena, registry, ctx, &rhs_sheet, &rhs_top, &rhs_left, &rhs_bottom, &rhs_right,
                              out_err)) {
    return false;
  }
  std::string_view sheet;
  if (!merge_range_endpoint_sheets(lhs_ast, lhs_sheet, rhs_ast, rhs_sheet, ctx, &sheet, out_err)) {
    return false;
  }
  out_lhs->sheet = sheet;
  out_lhs->row = std::min(lhs_top, rhs_top);
  out_lhs->col = std::min(lhs_left, rhs_left);
  out_rhs->sheet = sheet;
  out_rhs->row = std::max(lhs_bottom, rhs_bottom);
  out_rhs->col = std::max(lhs_right, rhs_right);
  return true;
}

}  // namespace

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
        // Build a flat argv pointer array from the call's child slots.
        std::vector<const parser::AstNode*> argv;
        argv.reserve(arity);
        for (std::uint32_t i = 0; i < arity; ++i) {
          argv.push_back(&node.as_call_arg(i));
        }
        return invoke_lambda(bound.as_lambda(), arity, argv.empty() ? nullptr : argv.data(), arena, registry, ctx);
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
    std::vector<const parser::AstNode*> argv;
    argv.reserve(arity);
    for (std::uint32_t i = 0; i < arity; ++i) {
      argv.push_back(&node.as_call_arg(i));
    }
    return invoke_lambda(defined.as_lambda(), arity, argv.empty() ? nullptr : argv.data(), arena, registry, ctx);
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
  // Parallel to `values`: the whole column / row a scalar function's argument
  // evaluated to, whose `values` slot is then a placeholder.
  std::vector<const TailArray*> tails;
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
  const std::uint32_t evaluated_arity = def->analysis_toolpak_args && !def->atp_omitted_optional_is_na
                                            ? atp_evaluated_arity(node, def->min_arity, def->max_arity)
                                            : arity;
  for (std::uint32_t i = 0; i < evaluated_arity; ++i) {
    const parser::AstNode& raw_arg = node.as_call_arg(i);
    // LET-binding passthrough: when the caller wrote `SUM(r)` where `r` is
    // bound to a RangeOp / ArrayLiteral / OFFSET / CHOOSE / INDIRECT, the
    // dispatcher must see the underlying range AST or it would fall back
    // to the scalar path and collapse the binding to its spill anchor.
    // Substitute only when the resolved AST is genuinely range-shaped so
    // that a NameRef bound to a scalar (or a single-cell Ref) continues to
    // flow through the existing scalar branch with its original provenance.
    const parser::AstNode& arg_node = resolve_range_binding(raw_arg, ctx.name_env(), /*accept_ref=*/false);
    // Minimal array-literal support: when a range-aware function receives
    // a `{a;b;c}` style literal, flatten it in row-major order exactly like
    // a RangeOp argument. This is just enough to let LARGE / SMALL /
    // PERCENTILE.INC / QUARTILE.INC accept brace literals as their "array"
    // input. The same literals are also first-class dynamic arrays when
    // evaluated directly; flattening here preserves this range-aware
    // function dispatch path.
    if (def->accepts_ranges && arg_node.kind() == parser::NodeKind::ArrayLiteral) {
      had_range_shaped_arg = true;
      auto resolved = resolve_range_arg(arg_node, arena, registry, ctx);
      if (!resolved) {
        return Value::error(resolved.error());
      }
      bool short_circuit = false;
      Value propagated_err = Value::blank();
      for (const Value& value : resolved.value().cells) {
        // A boolean inside an array literal is #VALUE! under the
        // Analysis-ToolPak rule (`GCD({TRUE,2})`).
        if (def->analysis_toolpak_args && value.is_boolean()) {
          return Value::error(ErrorCode::Value);
        }
        if (!append_range_sourced_value(*def, value, &values, &propagated_err)) {
          short_circuit = true;
          break;
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
      std::string_view anchor_sheet;
      std::uint32_t anchor_row = 0;
      std::uint32_t anchor_col = 0;
      ErrorCode spill_err = ErrorCode::Ref;
      if (!resolve_spill_anchor_node(arg_node, arena, registry, ctx, &anchor_sheet, &anchor_row, &anchor_col,
                                     &spill_err)) {
        const Value err = Value::error(spill_err);
        if (def->propagate_errors) {
          return err;
        }
        values.push_back(err);
        continue;
      }
      const Sheet* current = ctx.current_sheet();
      const Sheet* target = current;
      if (current == nullptr) {
        spill_err = ErrorCode::Name;
      } else if (!anchor_sheet.empty()) {
        const Workbook* wb = ctx.workbook();
        if (wb == nullptr) {
          target = nullptr;
        } else {
          target = wb->sheet_by_name(anchor_sheet);
        }
      }
      if (target == nullptr || anchor_row >= Sheet::kMaxRows || anchor_col >= Sheet::kMaxCols) {
        const Value err = Value::error(spill_err);
        if (def->propagate_errors) {
          return err;
        }
        values.push_back(err);
        continue;
      }
      // Copied out under the sheet lock, with Text payloads re-homed into
      // `arena`: the flattened values feed the whole call and outlive the
      // region they came from.
      std::vector<Value> region_cells;
      if (!ctx.read_spill_region(*target, anchor_row, anchor_col, arena, region_cells, nullptr, nullptr)) {
        const Value err = Value::error(ErrorCode::Ref);
        if (def->propagate_errors) {
          return err;
        }
        values.push_back(err);
        continue;
      }
      Value range_err = Value::blank();
      if (!append_range_sourced_values(*def, region_cells.data(), region_cells.size(), &values, &range_err)) {
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
      std::vector<Value> ref3d_cells;
      if (!expand_ref3d_cells(arg_node, arena, registry, ctx, &ref3d_cells)) {
        const Value err = Value::error(ErrorCode::Ref);
        if (def->propagate_errors) {
          return err;
        }
        values.push_back(err);
        continue;
      }
      had_range_shaped_arg = true;
      Value range_err = Value::blank();
      if (!append_range_sourced_values(*def, ref3d_cells.data(), ref3d_cells.size(), &values, &range_err)) {
        return range_err;
      }
      continue;
    }
    // Cross-workbook 3-D reference (`SUM([Book.xlsx]S1:S2!A1)`): the same
    // flattening over the link's sheet order, read from its cache.
    if (def->accepts_ranges && arg_node.kind() == parser::NodeKind::ExternalRef &&
        !arg_node.as_external_ref_sheet_end().empty()) {
      std::vector<Value> ref3d_cells;
      if (!collect_external_ref3d_cells(arg_node, arena, ctx, &ref3d_cells)) {
        const Value err = Value::error(ErrorCode::Ref);
        if (def->propagate_errors) {
          return err;
        }
        values.push_back(err);
        continue;
      }
      had_range_shaped_arg = true;
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
      // Multi-column (`A:C`) / multi-row (`1:3`) whole references parse as a
      // RangeOp over two whole-column / whole-row Refs. `resolve_range_endpoint`
      // rejects whole references (no bounded rectangle on their own), so route
      // the pair straight to `expand_range`, which clamps the unbounded axis to
      // the sheet's used range.
      //
      // A `:` chain of three or more endpoints nests, so neither side is a
      // bare `Ref`; `declared_rect_endpoint_pair` reduces the whole chain
      // to the endpoint pair that bounds it. Without it a chain falls into
      // the endpoint union below, which rejects whole references and turns
      // `=SUM(A:C:E:G)` into `#VALUE!` where Excel answers.
      parser::Reference chain_lhs{};
      parser::Reference chain_rhs{};
      const bool is_bare_pair = lhs_ast.kind() == parser::NodeKind::Ref && rhs_ast.kind() == parser::NodeKind::Ref;
      const bool whole_axis_pair = is_bare_pair && (lhs_ast.as_ref().is_full_col || lhs_ast.as_ref().is_full_row ||
                                                    rhs_ast.as_ref().is_full_col || rhs_ast.as_ref().is_full_row);
      if (whole_axis_pair) {
        chain_lhs = lhs_ast.as_ref();
        chain_rhs = rhs_ast.as_ref();
      } else if (is_bare_pair || !declared_rect_endpoint_pair(arg_node, &chain_lhs, &chain_rhs)) {
        // Union the two endpoints into one rectangle and let `expand_range`
        // validate sheet equality (mismatched qualifiers surface as #REF!).
        ErrorCode endpoint_err = ErrorCode::Ref;
        if (!union_endpoint_refs(lhs_ast, rhs_ast, arena, registry, ctx, &chain_lhs, &chain_rhs, &endpoint_err)) {
          const Value err = Value::error(endpoint_err);
          if (def->propagate_errors) {
            return err;
          }
          values.push_back(err);
          continue;
        }
      }
      Value range_return = Value::blank();
      if (!append_expanded_range(*def, chain_lhs, chain_rhs, arena, registry, ctx, &values, &range_return)) {
        return range_return;
      }
      continue;
    }
    // Scalar (non-range-aware) function receiving a bounded `A1:B2`
    // RangeOp: Excel 365 evaluates the function element-wise over the
    // rectangle and spills the result rather than collapsing the range
    // to its implicit-intersection anchor. Materialise the rectangle as a
    // `Value::Array` here; the post-loop broadcast step then lifts the
    // scalar impl over it. Whole-column / whole-row endpoints (`A:A`,
    // `1:1`) keep the legacy anchor projection (they fall through to the
    // scalar `eval_node` path below) to avoid spilling an unbounded
    // rectangle through a scalar function. `@` / SINGLE-wrapped args are
    // Call nodes, not RangeOp, so they never reach this branch and retain
    // implicit-intersection semantics. A rectangle spanning a whole grid axis
    // (`A:C`, `1:3`, `A1:A1048576`) is not expanded here: it falls through to
    // `eval_node_shaped`, which reads it at its declared size without
    // materialising the unused part.
    if (!def->accepts_ranges && arg_node.kind() == parser::NodeKind::RangeOp) {
      const parser::AstNode& lhs_ast = arg_node.as_range_lhs();
      const parser::AstNode& rhs_ast = arg_node.as_range_rhs();
      const bool whole =
          (lhs_ast.kind() == parser::NodeKind::Ref && (lhs_ast.as_ref().is_full_col || lhs_ast.as_ref().is_full_row)) ||
          (rhs_ast.kind() == parser::NodeKind::Ref && (rhs_ast.as_ref().is_full_col || rhs_ast.as_ref().is_full_row)) ||
          (lhs_ast.kind() == parser::NodeKind::Ref && rhs_ast.kind() == parser::NodeKind::Ref &&
           is_full_axis_range(lhs_ast.as_ref(), rhs_ast.as_ref()));
      if (!whole) {
        parser::Reference union_lhs{};
        parser::Reference union_rhs{};
        ErrorCode endpoint_err = ErrorCode::Ref;
        if (!union_endpoint_refs(lhs_ast, rhs_ast, arena, registry, ctx, &union_lhs, &union_rhs, &endpoint_err)) {
          const Value err = Value::error(endpoint_err);
          if (def->propagate_errors) {
            return err;
          }
          values.push_back(err);
          continue;
        }
        auto expanded = ctx.expand_range(union_lhs, union_rhs, arena, registry);
        if (!expanded) {
          const Value err = Value::error(expanded.error());
          if (def->propagate_errors) {
            return err;
          }
          values.push_back(err);
          continue;
        }
        const std::uint32_t rrows = union_rhs.row - union_lhs.row + 1u;
        const std::uint32_t rcols = union_rhs.col - union_lhs.col + 1u;
        const std::vector<Value>& ev = expanded.value();
        ArrayValue* arr = array_from_values(rrows, rcols, ev.data(), ev.size(), arena);
        if (arr == nullptr) {
          return Value::error(ErrorCode::Num);
        }
        values.push_back(Value::array(arr));
        continue;
      }
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
      Value range_return = Value::blank();
      if (!append_expanded_range(*def, isect_lhs, isect_rhs, arena, registry, ctx, &values, &range_return)) {
        return range_return;
      }
      continue;
    }
    // Union operator (`(A1:A2,B1:B2)`) as a range-aware function argument:
    // Excel's comma operator concatenates the areas, and an aggregator must
    // see every cell of every area. Resolve each area through
    // `resolve_range_arg` (which handles Ref / RangeOp / range-shaped calls)
    // and flatten them in order. Overlapping areas are intentionally counted
    // more than once — Excel's union does NOT de-duplicate, so
    // `SUM((A1:A2,A1:A2))` doubles the A1:A2 total.
    if (def->accepts_ranges && arg_node.kind() == parser::NodeKind::UnionOp) {
      had_range_shaped_arg = true;
      const std::uint32_t area_count = arg_node.as_union_arity();
      Value range_err = Value::blank();
      bool short_circuit = false;
      for (std::uint32_t area = 0; area < area_count; ++area) {
        auto resolved = resolve_range_arg(arg_node.as_union_child(area), arena, registry, ctx);
        if (!resolved) {
          const Value err = Value::error(resolved.error());
          if (def->propagate_errors) {
            return err;
          }
          values.push_back(err);
          continue;
        }
        const RangeResult& rr = resolved.value();
        if (!append_range_sourced_values(*def, rr.cells.data(), rr.cells.size(), &values, &range_err)) {
          short_circuit = true;
          break;
        }
      }
      if (short_circuit) {
        return range_err;
      }
      continue;
    }
    // Range-aware functions that receive `OFFSET(...)` as an argument see
    // the rectangle OFFSET would synthesize, not the `#VALUE!` OFFSET
    // itself returns in scalar context. We share the expansion helper
    // with `resolve_range_arg` (lazy family) so the two paths cannot
    // drift on cross-sheet / cycle / bounds semantics.
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
    // INDIRECT names a reference, so its cells take the range filters even
    // when it resolves to one cell: SUM(INDIRECT("A2")) skips text as SUM(A2) does.
    if (!append_expanded_call_argument(*def, arg_node, "INDIRECT", expand_indirect_call, arena, registry, ctx, &values,
                                       &expanded_call_handled, &expanded_call_return)) {
      return expanded_call_return;
    }
    if (expanded_call_handled) {
      had_range_shaped_arg = true;
      continue;
    }
    // IF as a range producer: a picked range branch is flattened through the
    // range filters, so `=LET(r, IF(TRUE, A1:A3, B1:B3), SUM(r))` aggregates
    // three cells. A picked plain value (`IF(TRUE, "", 1)`) is a direct
    // argument, not a range cell, so the provenance filter must not drop it.
    if (def->accepts_ranges && arg_node.kind() == parser::NodeKind::Call &&
        strings::case_insensitive_eq(arg_node.as_call_name(), "IF")) {
      auto picked = resolve_range_arg(arg_node, arena, registry, ctx);
      if (!picked) {
        const Value err = Value::error(picked.error());
        if (def->propagate_errors) {
          return err;
        }
        values.push_back(err);
        continue;
      }
      if (picked.value().from_scalar && picked.value().cells.size() == 1U) {
        const Value& direct = picked.value().cells[0];
        if (def->propagate_errors && direct.is_error()) {
          return direct;
        }
        values.push_back(direct);
        continue;
      }
      had_range_shaped_arg = true;
      Value range_err = Value::blank();
      if (!append_range_sourced_values(*def, picked.value().cells.data(), picked.value().cells.size(), &values,
                                       &range_err)) {
        return range_err;
      }
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
    // Whole-column (`A:A`) / whole-row (`1:1`) single reference: expand
    // against the sheet's used range so range-aware aggregators see the
    // populated cells instead of the `#VALUE!` a scalar `resolve_ref`
    // returns for an unbounded reference. Multi-span (`A:C` / `1:3`) is
    // handled by the RangeOp branch above.
    if (def->accepts_ranges && arg_node.kind() == parser::NodeKind::Ref &&
        (arg_node.as_ref().is_full_col || arg_node.as_ref().is_full_row)) {
      had_range_shaped_arg = true;
      const parser::Reference& ref = arg_node.as_ref();
      Value range_return = Value::blank();
      if (!append_expanded_range(*def, ref, ref, arena, registry, ctx, &values, &range_return)) {
        return range_return;
      }
      continue;
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
      if (!append_range_sourced_values(*def, arr->cells, n, &values, &range_err)) {
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
        if (def->propagate_errors) {
          return err;
        }
        values.push_back(err);
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
