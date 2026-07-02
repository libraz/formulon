// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of the shape / geometry-inspection lazy builtins
// (`ROWS`, `COLUMNS`, `ROW`, `COLUMN`, `SUMPRODUCT`). Each dispatches on
// the raw AST of its argument(s) rather than a flattened `Value`, which
// is why these functions ride the lazy-dispatch seam rather than the
// eager path in `tree_walker::dispatch_call`.
//
// See `eval/shape_ops_lazy.h` for the dispatch-table contract and
// `eval/lazy_impls.h` for the shared `eval_node` / `LazyImpl` vocabulary.

#include "eval/shape_ops_lazy.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <vector>

#include "eval/coerce.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "eval/name_env_resolve.h"
#include "eval/range_args.h"
#include "eval/range_expanders.h"
#include "eval/range_resolvers.h"
#include "eval/scalar_ops.h"
#include "eval/tree_walker/broadcast.h"
#include "parser/ast.h"
#include "parser/reference.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/strings.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Fills `out_rows` / `out_cols` with the rectangle shape implied by
// `arg_node`. Returns `true` on success; on failure sets `*out_err` to
// the appropriate Excel error (e.g. expansion error -> `#REF!`) and
// returns `false`. Scalar / non-reference arguments degenerate to 1x1,
// matching Excel's treatment of `ROWS(scalar)` and SUMPRODUCT's scalar
// broadcast rule (where "broadcast" here means only "a scalar is 1x1",
// not the full v-array broadcasting Excel 365 implements for `--`).
//
// For an `ArrayLiteral` both dimensions come from the AST directly. For
// any other kind we simply evaluate it to determine whether it is an
// error (and propagate if so); the shape is 1x1 otherwise.
bool resolve_shape(const parser::AstNode& raw_arg, Arena& arena, const FunctionRegistry& registry,
                   const EvalContext& ctx, std::uint32_t* out_rows, std::uint32_t* out_cols, Value* out_err) {
  // LET-binding passthrough: `LET(x, A1:C3, ROWS(x))` parses `x` as a
  // NameRef; we want the bound RangeOp / ArrayLiteral / OFFSET-call AST
  // so the kind dispatch below sees the same shape it would for a literal
  // `=ROWS(A1:C3)`. Single-cell Refs and scalar bindings are left as-is
  // (the scalar fallback already returns 1x1 for them).
  const parser::AstNode* effective = &raw_arg;
  if (raw_arg.kind() == parser::NodeKind::NameRef) {
    const parser::AstNode& resolved = resolve_name_ast(raw_arg, ctx.name_env());
    if (&resolved != &raw_arg && is_range_shaped_ast(resolved)) {
      effective = &resolved;
    }
  }
  const parser::AstNode& arg_node = *effective;
  const parser::NodeKind k = arg_node.kind();
  if (k == parser::NodeKind::Ref || k == parser::NodeKind::RangeOp) {
    auto resolved = resolve_range_arg(arg_node, arena, registry, ctx);
    if (!resolved) {
      *out_err = Value::error(resolved.error());
      return false;
    }
    *out_rows = resolved.value().rows;
    *out_cols = resolved.value().cols;
    return true;
  }
  if (k == parser::NodeKind::ArrayLiteral) {
    *out_rows = arg_node.as_array_rows();
    *out_cols = arg_node.as_array_cols();
    return true;
  }
  if (k == parser::NodeKind::Call) {
    // Recognise reference-returning builtins (INDIRECT, OFFSET) so
    // `ROWS(INDIRECT("A1:B3"))` reports 3 rather than degrading to the
    // 1x1 scalar fallback.
    std::string_view sheet;
    std::uint32_t top = 0;
    std::uint32_t left = 0;
    std::uint32_t bottom = 0;
    std::uint32_t right = 0;
    bool is_range = false;
    ErrorCode err = ErrorCode::Value;
    if (resolve_reference_call(arg_node, arena, registry, ctx, &sheet, &top, &left, &bottom, &right, &is_range, &err)) {
      *out_rows = bottom - top + 1U;
      *out_cols = right - left + 1U;
      return true;
    }
    // Fall through to the scalar fallback below so subtree errors
    // propagate and non-reference calls degrade to 1x1.
  }
  // Scalar fallback. We still evaluate so an error argument propagates
  // with the correct code instead of silently producing 1.
  const Value v = eval_node(arg_node, arena, registry, ctx);
  if (v.is_error()) {
    *out_err = v;
    return false;
  }
  if (v.is_array()) {
    *out_rows = v.as_array_rows();
    *out_cols = v.as_array_cols();
    return true;
  }
  *out_rows = 1U;
  *out_cols = 1U;
  return true;
}

// Walks `arg_node` as an AST array literal, evaluating each element into
// `out_cells` in row-major order. Records the literal's `(rows, cols)`
// shape. On the first evaluated element that turns out to be an error,
// returns `false` and reports the error via `*out_err` (callers treat
// the error as the overall result of SUMPRODUCT).
bool flatten_array_literal(const parser::AstNode& arg_node, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx, std::vector<Value>* out_cells, std::uint32_t* out_rows,
                           std::uint32_t* out_cols, Value* out_err) {
  const std::uint32_t rows = arg_node.as_array_rows();
  const std::uint32_t cols = arg_node.as_array_cols();
  *out_rows = rows;
  *out_cols = cols;
  out_cells->clear();
  out_cells->reserve(static_cast<std::size_t>(rows) * cols);
  for (std::uint32_t r = 0; r < rows; ++r) {
    for (std::uint32_t c = 0; c < cols; ++c) {
      Value v = eval_node(arg_node.as_array_element(r, c), arena, registry, ctx);
      if (v.is_error()) {
        *out_err = v;
        return false;
      }
      out_cells->push_back(v);
    }
  }
  return true;
}

// SUMPRODUCT's per-cell coercion rule. `Number` contributes its value;
// every other type (including `Text` that happens to be numeric-looking,
// per Excel's long-standing behaviour for this function) contributes 0.
// Errors are handled one level up (they short-circuit the whole call in
// scan order); this helper must only be called on non-error values.
double sumproduct_coerce(const Value& v) {
  if (v.is_number()) {
    return v.as_number();
  }
  return 0.0;
}

// Builds a freshly arena-allocated `ArrayValue` from a flat row-major cell
// vector and shape. The cells are copied into a fresh arena buffer so the
// caller's vector can go out of scope safely; both the buffer and the
// `ArrayValue` header live in `arena` for the same lifetime contract as
// `Value::Text`.
const ArrayValue* make_array_value(Arena& arena, std::uint32_t rows, std::uint32_t cols,
                                   const std::vector<Value>& cells) {
  const std::size_t n = static_cast<std::size_t>(rows) * static_cast<std::size_t>(cols);
  Value* buffer = arena.create_array<Value>(n);
  for (std::size_t i = 0; i < n; ++i) {
    buffer[i] = cells[i];
  }
  ArrayValue* arr = arena.create<ArrayValue>();
  arr->rows = rows;
  arr->cols = cols;
  arr->cells = buffer;
  return arr;
}

// True for the reference-shaped `Call` names that `resolve_range_arg`
// already knows how to expand into a flat row-major rectangle. Kept
// alongside the array-context dispatch so the two seams stay in sync
// when a new reference-producing builtin (XLOOKUP-style spilling, etc.)
// is added.
bool is_range_producing_call(const parser::AstNode& call_node) {
  if (call_node.kind() != parser::NodeKind::Call) {
    return false;
  }
  const std::string_view name = call_node.as_call_name();
  return strings::case_insensitive_eq(name, "OFFSET") || strings::case_insensitive_eq(name, "CHOOSE") ||
         strings::case_insensitive_eq(name, "IF") || strings::case_insensitive_eq(name, "ROW") ||
         strings::case_insensitive_eq(name, "COLUMN");
}

}  // namespace

Value eval_rows_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  if (call.as_call_arity() != 1U) {
    return Value::error(ErrorCode::Value);
  }
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
  Value err = Value::blank();
  if (!resolve_shape(call.as_call_arg(0), arena, registry, ctx, &rows, &cols, &err)) {
    return err;
  }
  return Value::number(static_cast<double>(rows));
}

Value eval_columns_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx) {
  if (call.as_call_arity() != 1U) {
    return Value::error(ErrorCode::Value);
  }
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
  Value err = Value::blank();
  if (!resolve_shape(call.as_call_arg(0), arena, registry, ctx, &rows, &cols, &err)) {
    return err;
  }
  return Value::number(static_cast<double>(cols));
}

// Shared helper for ROW / COLUMN. Picks the row or column axis via
// `want_row`. `References` are stored 0-based internally; Excel exposes
// them 1-based, hence the `+1` on the way out. An `ArrayLiteral` (or
// anything else that isn't a reference / RangeOp of refs) surfaces as
// `#VALUE!`: array literals in Excel are not references, and a bare
// `=ROW({1,2,3})` yields `#VALUE!`.
namespace {

Value eval_row_or_column(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx, bool want_row) {
  if (call.as_call_arity() == 0U) {
    // ROW() / COLUMN() with no argument returns the formula cell's own
    // row / column (1-based). When no formula cell is bound — e.g. ad-hoc
    // `eval` CLI invocations with no anchoring address — surface `#VALUE!`
    // because the result is genuinely undefined in that context.
    if (!ctx.has_formula_cell()) {
      return Value::error(ErrorCode::Value);
    }
    const std::uint32_t idx = want_row ? ctx.formula_row() : ctx.formula_col();
    return Value::number(static_cast<double>(idx + 1U));
  }
  if (call.as_call_arity() != 1U) {
    return Value::error(ErrorCode::Value);
  }
  const parser::AstNode& raw_arg = call.as_call_arg(0);
  // LET-binding passthrough: `=LET(r, A1:A3, ROW(r))` parses `r` as a
  // NameRef, but Mac Excel returns the first row of the bound rectangle
  // rather than `#VALUE!`. Look through the NameRef to the bound AST and
  // accept the broader "Ref OR range-shaped" set: ROW(single-cell Ref)
  // is meaningful (returns its row), so unlike the SUM-style passthrough
  // we do not exclude bare Refs here.
  const parser::AstNode* effective = &raw_arg;
  if (raw_arg.kind() == parser::NodeKind::NameRef) {
    const parser::AstNode& resolved = resolve_name_ast(raw_arg, ctx.name_env());
    if (&resolved != &raw_arg && (resolved.kind() == parser::NodeKind::Ref || is_range_shaped_ast(resolved))) {
      effective = &resolved;
    }
  }
  const parser::AstNode& arg = *effective;
  const parser::NodeKind k = arg.kind();
  if (k == parser::NodeKind::Ref) {
    const parser::Reference& r = arg.as_ref();
    const std::uint32_t idx = want_row ? r.row : r.col;
    return Value::number(static_cast<double>(idx + 1U));
  }
  if (k == parser::NodeKind::RangeOp) {
    const parser::AstNode& lhs_ast = arg.as_range_lhs();
    const parser::AstNode& rhs_ast = arg.as_range_rhs();
    if (lhs_ast.kind() != parser::NodeKind::Ref || rhs_ast.kind() != parser::NodeKind::Ref) {
      return Value::error(ErrorCode::Value);
    }
    const parser::Reference& lhs = lhs_ast.as_ref();
    const parser::Reference& rhs = rhs_ast.as_ref();
    // Excel returns the first-row / first-column of the rectangle, which
    // is the smaller of the two endpoints after normalisation.
    const std::uint32_t a = want_row ? lhs.row : lhs.col;
    const std::uint32_t b = want_row ? rhs.row : rhs.col;
    const std::uint32_t lo = a < b ? a : b;
    return Value::number(static_cast<double>(lo + 1U));
  }
  if (k == parser::NodeKind::Call) {
    // Reference-returning builtins (INDIRECT, OFFSET) nested inside
    // ROW/COLUMN must be inspected without dereferencing so
    // `ROW(INDIRECT("A5"))` returns 5 rather than the value at A5.
    std::string_view sheet;
    std::uint32_t top = 0;
    std::uint32_t left = 0;
    std::uint32_t bottom = 0;
    std::uint32_t right = 0;
    bool is_range = false;
    ErrorCode err = ErrorCode::Value;
    if (resolve_reference_call(arg, arena, registry, ctx, &sheet, &top, &left, &bottom, &right, &is_range, &err)) {
      const std::uint32_t idx = want_row ? top : left;
      return Value::number(static_cast<double>(idx + 1U));
    }
    // Fall through to the scalar fallback so subtree errors propagate
    // and non-reference calls still surface `#VALUE!`.
  }
  // Evaluate to surface errors from the subtree verbatim; otherwise
  // reject non-references as `#VALUE!` (matches Excel for ROW(literal)
  // and ROW({...}) array-literal forms).
  const Value v = eval_node(arg, arena, registry, ctx);
  if (v.is_error()) {
    return v;
  }
  return Value::error(ErrorCode::Value);
}

}  // namespace

Value eval_row_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                    const EvalContext& ctx) {
  return eval_row_or_column(call, arena, registry, ctx, /*want_row=*/true);
}

Value eval_column_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  return eval_row_or_column(call, arena, registry, ctx, /*want_row=*/false);
}

// SUMPRODUCT(arr1, arr2, ...)
//
// Excel semantics implemented here:
//   * Every argument must resolve to the same (rows, cols) rectangle
//     (including the "all args are 1x1 scalars" edge). Any mismatch
//     returns `#VALUE!`.
//   * RangeOp / Ref args are expanded via `resolve_range_arg`.
//   * ArrayLiteral args are walked by `flatten_array_literal`, which
//     evaluates each element through `eval_node` so nested calls /
//     literals behave correctly.
//   * BinaryOp / UnaryOp args are evaluated in array context via
//     `eval_node_as_array`, so `(A1:A5>2)*1` and `(A>2)*(B<10)` produce
//     a cellwise rectangle that participates in the SUMPRODUCT rather
//     than collapsing to a scalar at the top-level operator.
//   * Scalar (any other) args evaluate to a 1x1 vector.
//   * Errors propagate in row-major scan order: arrays are inspected
//     left-to-right, and within each array top-to-bottom row-major.
//   * Non-numeric cells (Bool, Text, Blank) contribute zero. This is
//     the long-standing Excel SUMPRODUCT rule: non-numerics are *not*
//     coerced, they simply drop out of the product as 0.
Value eval_sumproduct_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 1U) {
    return Value::error(ErrorCode::Value);
  }

  // Materialise every argument as a flat row-major vector + shape.
  // `all_args[i]` is (rows, cols, cells) for the i-th argument.
  struct ArgArray {
    std::uint32_t rows;
    std::uint32_t cols;
    std::vector<Value> cells;
  };
  std::vector<ArgArray> all_args;
  all_args.reserve(arity);

  for (std::uint32_t i = 0; i < arity; ++i) {
    const parser::AstNode& raw_arg = call.as_call_arg(i);
    // LET-binding passthrough: `=LET(r, A1:A3, SUMPRODUCT(r))` parses `r`
    // as a NameRef; we want the bound RangeOp / ArrayLiteral / OFFSET-call
    // / CHOOSE-call / IF-call AST so the kind dispatch below sees the same
    // shape it would for a literal `=SUMPRODUCT(A1:A3)`. Single-cell Refs
    // and scalar bindings are left as-is (the scalar fallback already
    // returns 1x1 for them). This mirrors `resolve_shape` above.
    const parser::AstNode* effective = &raw_arg;
    if (raw_arg.kind() == parser::NodeKind::NameRef) {
      const parser::AstNode& resolved = resolve_name_ast(raw_arg, ctx.name_env());
      if (&resolved != &raw_arg && is_range_shaped_ast(resolved)) {
        effective = &resolved;
      }
    }
    const parser::AstNode& arg_node = *effective;
    const parser::NodeKind k = arg_node.kind();
    ArgArray a{};
    if (k == parser::NodeKind::Ref || k == parser::NodeKind::RangeOp) {
      auto resolved = resolve_range_arg(arg_node, arena, registry, ctx);
      if (!resolved) {
        return Value::error(resolved.error());
      }
      auto& rr = resolved.value();
      a.rows = rr.rows;
      a.cols = rr.cols;
      a.cells = std::move(rr.cells);
    } else if (k == parser::NodeKind::Call) {
      // OFFSET / CHOOSE / IF call after LET passthrough — route through
      // `resolve_range_arg` so the rectangle (and its row/col shape) is
      // expanded the same way as a literal `RangeOp` argument.
      auto resolved = resolve_range_arg(arg_node, arena, registry, ctx);
      if (!resolved) {
        return Value::error(resolved.error());
      }
      auto& rr = resolved.value();
      a.rows = rr.rows;
      a.cols = rr.cols;
      a.cells = std::move(rr.cells);
    } else if (k == parser::NodeKind::ArrayLiteral) {
      Value err = Value::blank();
      if (!flatten_array_literal(arg_node, arena, registry, ctx, &a.cells, &a.rows, &a.cols, &err)) {
        return err;
      }
    } else if (k == parser::NodeKind::BinaryOp || k == parser::NodeKind::UnaryOp) {
      // Array-context evaluation: BinaryOp / UnaryOp args carry range-shaped
      // subexpressions that must be broadcast cellwise. `eval_node_as_array`
      // recurses through scalar_ops to produce an ArrayValue (or scalar error
      // on shape mismatch / left-most-error short-circuit). This is what makes
      // `=SUMPRODUCT((A1:A5>2)*1)` and `=SUMPRODUCT((A>2)*(B<10), C)` compute
      // the cellwise product instead of collapsing to scalar.
      const Value arr_v = eval_node_as_array(arg_node, arena, registry, ctx);
      if (arr_v.is_error()) {
        return arr_v;
      }
      // `eval_node_as_array` is contracted to return either an Array or a
      // scalar Error; the is_array() check is defensive against future API
      // drift.
      if (!arr_v.is_array()) {
        return Value::error(ErrorCode::Value);
      }
      const ArrayValue* arr = arr_v.as_array();
      a.rows = arr->rows;
      a.cols = arr->cols;
      const std::size_t n = static_cast<std::size_t>(arr->rows) * static_cast<std::size_t>(arr->cols);
      a.cells.assign(arr->cells, arr->cells + n);
    } else {
      // Scalar argument: evaluate and treat as 1x1.
      const Value v = eval_node(arg_node, arena, registry, ctx);
      if (v.is_error()) {
        return v;
      }
      a.rows = 1U;
      a.cols = 1U;
      a.cells.push_back(v);
    }
    all_args.push_back(std::move(a));
  }

  // Shape check: all arrays must share the reference shape taken from
  // the first argument. Scalars (1x1) ride on this rule too — any
  // mismatch (e.g. 1x1 vs 3x1) is `#VALUE!`.
  const std::uint32_t ref_rows = all_args.front().rows;
  const std::uint32_t ref_cols = all_args.front().cols;
  for (std::size_t i = 1; i < all_args.size(); ++i) {
    if (all_args[i].rows != ref_rows || all_args[i].cols != ref_cols) {
      return Value::error(ErrorCode::Value);
    }
  }

  // Scan for errors in canonical Excel order: for each argument
  // (left-to-right), walk its cells in row-major order. The first
  // error encountered wins. This runs before the numeric accumulation
  // so the returned code matches Excel's leftmost-wins rule even if a
  // later numeric overflow would otherwise upstage it.
  for (const ArgArray& a : all_args) {
    for (const Value& v : a.cells) {
      if (v.is_error()) {
        return v;
      }
    }
  }

  // Element-wise product accumulated into total. The element index
  // `idx` walks `ref_rows * ref_cols` positions in row-major order.
  const std::size_t n = static_cast<std::size_t>(ref_rows) * static_cast<std::size_t>(ref_cols);
  double total = 0.0;
  for (std::size_t idx = 0; idx < n; ++idx) {
    double product = 1.0;
    for (const ArgArray& a : all_args) {
      product *= sumproduct_coerce(a.cells[idx]);
    }
    total += product;
  }
  if (std::isnan(total) || std::isinf(total)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(total);
}

Value eval_node_as_array(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx) {
  // LET-binding NameRef passthrough: mirror the pattern used by the other
  // shape-aware seams so `=LET(r, A1:A3, r+1)` sees the bound RangeOp / Call
  // AST rather than the NameRef wrapper.
  const parser::AstNode* effective = &node;
  if (node.kind() == parser::NodeKind::NameRef) {
    const parser::AstNode& resolved = resolve_name_ast(node, ctx.name_env());
    if (&resolved != &node && (resolved.kind() == parser::NodeKind::Ref || is_range_shaped_ast(resolved))) {
      effective = &resolved;
    }
  }
  const parser::AstNode& target = *effective;
  const parser::NodeKind k = target.kind();

  // BinaryOp -> recurse into the cellwise broadcaster.
  if (k == parser::NodeKind::BinaryOp) {
    return eval_binop_array_ctx(target, arena, registry, ctx);
  }

  // UnaryOp -> evaluate operand as array, then apply cellwise.
  if (k == parser::NodeKind::UnaryOp) {
    const Value inner = eval_node_as_array(target.as_unary_operand(), arena, registry, ctx);
    if (inner.is_error()) {
      return inner;
    }
    // `eval_node_as_array` is contracted to return either an Array or a
    // scalar Error; any other shape would be a bug in this seam.
    if (!inner.is_array()) {
      return Value::error(ErrorCode::Value);
    }
    const ArrayValue* in_arr = inner.as_array();
    const std::size_t n = static_cast<std::size_t>(in_arr->rows) * static_cast<std::size_t>(in_arr->cols);
    std::vector<Value> out;
    out.reserve(n);
    for (std::size_t i = 0; i < n; ++i) {
      out.push_back(apply_unary(target.as_unary_op(), in_arr->cells[i]));
    }
    return Value::array(make_array_value(arena, in_arr->rows, in_arr->cols, out));
  }

  // Range-shaped AST: Ref / RangeOp / one of the reference-producing Calls.
  // `resolve_range_arg` already knows how to expand each of these into a
  // flat row-major vector + shape, so reuse it to keep the expansion path
  // identical to the conditional-aggregator and lookup families.
  if (k == parser::NodeKind::Ref || k == parser::NodeKind::RangeOp || is_range_producing_call(target)) {
    auto resolved = resolve_range_arg(target, arena, registry, ctx);
    if (!resolved) {
      return Value::error(resolved.error());
    }
    auto& rr = resolved.value();
    return Value::array(make_array_value(arena, rr.rows, rr.cols, rr.cells));
  }

  // ArrayLiteral in array context preserves per-cell errors instead of
  // short-circuiting the whole argument.
  if (k == parser::NodeKind::ArrayLiteral) {
    std::vector<Value> cells;
    const std::uint32_t rows = target.as_array_rows();
    const std::uint32_t cols = target.as_array_cols();
    cells.reserve(static_cast<std::size_t>(rows) * cols);
    for (std::uint32_t r = 0; r < rows; ++r) {
      for (std::uint32_t c = 0; c < cols; ++c) {
        cells.push_back(eval_node(target.as_array_element(r, c), arena, registry, ctx));
      }
    }
    return Value::array(make_array_value(arena, rows, cols, cells));
  }

  // Scalar fallback. Evaluate normally, then wrap into a 1x1 array. Errors
  // stay scalar so they can short-circuit the broadcaster in the caller.
  // A pre-existing Array (some future array-producing builtin) is forwarded
  // unchanged so eval_node_as_array remains idempotent on its outputs.
  const Value v = eval_node(target, arena, registry, ctx);
  if (v.is_error()) {
    return v;
  }
  if (v.is_array()) {
    return v;
  }
  const std::vector<Value> single{v};
  return Value::array(make_array_value(arena, 1U, 1U, single));
}

Value eval_binop_array_ctx(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx) {
  const Value lhs = eval_node_as_array(node.as_binary_lhs(), arena, registry, ctx);
  if (lhs.is_error()) {
    return lhs;
  }
  const Value rhs = eval_node_as_array(node.as_binary_rhs(), arena, registry, ctx);
  if (rhs.is_error()) {
    return rhs;
  }
  // Both operands are guaranteed Array post the array-context contract.
  // Delegate to the single Excel-broadcasting implementation shared with the
  // top-level BinaryOp dispatch (`tree_walker/broadcast.cpp`), which handles
  // size-1 broadcast, the outer product, RxC/Rx1/1xC combinations, and #N/A
  // padding for mismatched non-1 axes.
  return broadcast_binop(node.as_binary_op(), lhs, rhs, arena);
}

Value eval_transpose_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx) {
  if (call.as_call_arity() != 1U) {
    return Value::error(ErrorCode::Value);
  }
  const parser::AstNode& arg = call.as_call_arg(0);

  // ArrayLiteral arg: materialise cells one-by-one preserving error cells
  // verbatim. The `flatten_array_literal` helper shared with SUMPRODUCT
  // short-circuits on the first error to match SUMPRODUCT semantics, but
  // TRANSPOSE must keep error cells in place — Mac Excel returns
  // `=TRANSPOSE({1,#N/A;2,3})` as an array containing #N/A at the
  // corresponding transposed position.
  if (arg.kind() == parser::NodeKind::ArrayLiteral) {
    const std::uint32_t in_rows = arg.as_array_rows();
    const std::uint32_t in_cols = arg.as_array_cols();
    const std::uint32_t out_rows = in_cols;
    const std::uint32_t out_cols = in_rows;
    const std::size_t n = static_cast<std::size_t>(out_rows) * static_cast<std::size_t>(out_cols);
    Value* buffer = arena.create_array<Value>(n);
    if (buffer == nullptr) {
      return Value::error(ErrorCode::Num);
    }
    for (std::uint32_t r = 0; r < in_rows; ++r) {
      for (std::uint32_t c = 0; c < in_cols; ++c) {
        const Value cell = eval_node(arg.as_array_element(r, c), arena, registry, ctx);
        // Per-cell error pass-through: errors land at their transposed
        // position rather than short-circuiting the whole call.
        buffer[static_cast<std::size_t>(c) * in_rows + r] = cell;
      }
    }
    ArrayValue* arr = arena.create<ArrayValue>();
    if (arr == nullptr) {
      return Value::error(ErrorCode::Num);
    }
    arr->rows = out_rows;
    arr->cols = out_cols;
    arr->cells = buffer;
    return Value::array(arr);
  }

  // For range-shaped args (Ref / RangeOp / reference-producing calls)
  // route through `eval_node_as_array` so the underlying 2D shape is
  // preserved. For every other arg shape (scalar literals, arithmetic,
  // calls returning scalars), evaluate directly via `eval_node` so a
  // scalar error short-circuits the whole TRANSPOSE — matching the eager
  // dispatcher's `propagate_errors=true` rule that would otherwise apply.
  Value v = Value::blank();
  if (arg.kind() == parser::NodeKind::Ref || arg.kind() == parser::NodeKind::RangeOp ||
      arg.kind() == parser::NodeKind::SpillRef || is_range_producing_call(arg)) {
    v = eval_node_as_array(arg, arena, registry, ctx);
  } else {
    v = eval_node(arg, arena, registry, ctx);
  }
  if (v.is_error()) {
    return v;
  }
  if (!v.is_array()) {
    // Defensive: the contract of `eval_node_as_array` is "Array or Error".
    // Treat any other shape as a 1x1 fall-through so TRANSPOSE remains
    // total over its input domain.
    Value* one = arena.create_array<Value>(1);
    if (one == nullptr) {
      return Value::error(ErrorCode::Num);
    }
    one[0] = v;
    ArrayValue* arr = arena.create<ArrayValue>();
    if (arr == nullptr) {
      return Value::error(ErrorCode::Num);
    }
    arr->rows = 1U;
    arr->cols = 1U;
    arr->cells = one;
    return Value::array(arr);
  }
  const ArrayValue* src = v.as_array();
  const std::uint32_t out_rows = src->cols;
  const std::uint32_t out_cols = src->rows;
  const std::size_t n = static_cast<std::size_t>(out_rows) * static_cast<std::size_t>(out_cols);
  Value* buffer = arena.create_array<Value>(n);
  if (buffer == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  // Row-major fill: dst[c * src->rows + r] = src[r * src->cols + c]. The
  // scalar 1x1 case is degenerate (both dimensions are 1) so the loop runs
  // once and yields the same single cell back, consistent with Mac Excel.
  for (std::uint32_t r = 0; r < src->rows; ++r) {
    for (std::uint32_t c = 0; c < src->cols; ++c) {
      buffer[static_cast<std::size_t>(c) * src->rows + r] = src->cells[static_cast<std::size_t>(r) * src->cols + c];
    }
  }
  ArrayValue* arr = arena.create<ArrayValue>();
  if (arr == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  arr->rows = out_rows;
  arr->cols = out_cols;
  arr->cells = buffer;
  return Value::array(arr);
}

}  // namespace eval
}  // namespace formulon
