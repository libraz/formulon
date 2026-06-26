// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Recursive node visitor for the tree-walk evaluator. Holds the public
// `evaluate()` overloads, the `eval_node` switch (declared with external
// linkage in `eval/lazy_impls.h` so lazy-impl TUs can recurse back into
// it), the read-only spill-collision detector, and — when the
// `FORMULON_VM_PARITY_CHECK` build option is on — the bytecode-VM
// parity diagnostic at the tail of `evaluate()`.
//
// The function-call dispatch path (`dispatch_call`, `invoke_lambda`,
// range-argument expansion) lives in `tree_walker/dispatch.cpp`; the
// array-broadcasting helpers (`broadcast_binop`, `broadcast_unary`,
// `apply_binop_per_cell`) live in `tree_walker/broadcast.cpp`. The
// three TUs split the original monolithic `tree_walker.cpp` while
// preserving file-local helpers and the existing `formulon::eval`
// namespace shape (anonymous helpers are local to each TU).
//
// See `tree_walker.h` for the public contract.

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <string_view>
#include <vector>

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/iterative_solver.h"
#include "eval/lambda_value.h"
#include "eval/lazy_impls.h"
#include "eval/name_env.h"
#include "eval/name_env_resolve.h"
#include "eval/range_resolvers.h"
#include "eval/structured_ref.h"
#include "eval/tree_walker.h"
#include "eval/tree_walker/broadcast.h"
#include "eval/tree_walker/depth_guard.h"
#include "eval/tree_walker/dispatch.h"
#include "eval/tree_walker_lazy_table.h"
#include "parser/ast.h"
#include "sheet.h"
#include "utils/arena.h"
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

#ifdef FORMULON_VM_PARITY_CHECK
// Parity-harness honesty filter.
//
// Returns true when any function call reachable from `node` resolves through
// the lazy-dispatch table (`find_lazy_impl`) but has NO eager `FunctionDef`
// in `registry`. Such "lazy-only" functions — IRR / MIRR / XIRR / XNPV,
// NETWORKDAYS / WORKDAY / REGEX* / TEXTSPLIT / PHONETIC, the higher-order
// array forms MAP / REDUCE / SCAN / BYROW / BYCOL / MAKEARRAY, and the
// AST-introspecting info functions — cannot be executed by the bytecode VM:
// the IR carries no AST at runtime, so the VM has no eager registry impl to
// call and surfaces `#NAME?`. That is a documented structural limitation of
// the bytecode IR, not a tree-walker / VM divergence, so the parity
// comparison must skip these formulas rather than report a false mismatch.
//
// IF / IFERROR / IFNA / AND / OR and the conditional aggregators are also in
// the lazy table, but they additionally have eager registry entries (or are
// lowered to dedicated opcodes), so they are NOT skipped — the VM evaluates
// them and parity is meaningful.
bool has_lazy_only_call(const parser::AstNode& node, const FunctionRegistry& registry) {
  switch (node.kind()) {
    case parser::NodeKind::Call: {
      const std::string_view name = node.as_call_name();
      if (find_lazy_impl(name) != nullptr && registry.lookup(name) == nullptr) {
        return true;
      }
      const std::uint32_t arity = node.as_call_arity();
      for (std::uint32_t i = 0; i < arity; ++i) {
        if (has_lazy_only_call(node.as_call_arg(i), registry)) {
          return true;
        }
      }
      return false;
    }
    case parser::NodeKind::UnaryOp:
      return has_lazy_only_call(node.as_unary_operand(), registry);
    case parser::NodeKind::BinaryOp:
      return has_lazy_only_call(node.as_binary_lhs(), registry) || has_lazy_only_call(node.as_binary_rhs(), registry);
    case parser::NodeKind::RangeOp:
      return has_lazy_only_call(node.as_range_lhs(), registry) || has_lazy_only_call(node.as_range_rhs(), registry);
    case parser::NodeKind::IntersectOp:
      return has_lazy_only_call(node.as_intersect_lhs(), registry) ||
             has_lazy_only_call(node.as_intersect_rhs(), registry);
    case parser::NodeKind::ImplicitIntersection:
      return has_lazy_only_call(node.as_implicit_intersection_operand(), registry);
    case parser::NodeKind::UnionOp: {
      const std::uint32_t arity = node.as_union_arity();
      for (std::uint32_t i = 0; i < arity; ++i) {
        if (has_lazy_only_call(node.as_union_child(i), registry)) {
          return true;
        }
      }
      return false;
    }
    case parser::NodeKind::ArrayLiteral: {
      const std::uint32_t rows = node.as_array_rows();
      const std::uint32_t cols = node.as_array_cols();
      for (std::uint32_t r = 0; r < rows; ++r) {
        for (std::uint32_t c = 0; c < cols; ++c) {
          if (has_lazy_only_call(node.as_array_element(r, c), registry)) {
            return true;
          }
        }
      }
      return false;
    }
    case parser::NodeKind::Lambda:
      return has_lazy_only_call(node.as_lambda_body(), registry);
    case parser::NodeKind::LetBinding: {
      const std::uint32_t count = node.as_let_binding_count();
      for (std::uint32_t i = 0; i < count; ++i) {
        if (has_lazy_only_call(node.as_let_binding_expr(i), registry)) {
          return true;
        }
      }
      return has_lazy_only_call(node.as_let_body(), registry);
    }
    case parser::NodeKind::LambdaCall: {
      if (has_lazy_only_call(node.as_lambda_call_callee(), registry)) {
        return true;
      }
      const std::uint32_t arity = node.as_lambda_call_arity();
      for (std::uint32_t i = 0; i < arity; ++i) {
        if (has_lazy_only_call(node.as_lambda_call_arg(i), registry)) {
          return true;
        }
      }
      return false;
    }
    default:
      // Leaf nodes (Literal, Ref, NameRef, error literals, etc.) carry no
      // nested calls.
      return false;
  }
}
#endif  // FORMULON_VM_PARITY_CHECK

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
// reachable via ordinary unqualified lookup. `dispatch_call` lives in
// `tree_walker/dispatch.cpp` and is declared in
// `eval/tree_walker/dispatch.h`; the broadcast helpers live in
// `tree_walker/broadcast.cpp` and are declared in
// `eval/tree_walker/broadcast.h`.
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
  // Parity harness. Compile the same AST through the bytecode pipeline,
  // run it through the VM, and compare the result. Mismatches are
  // surfaced as a diagnostic on stderr; we deliberately do NOT abort or
  // mutate the returned value so the production code path stays unchanged.
  // The parity sweep test reads the same `evaluate()` entry point and asserts
  // separately on its own corpus; this hook is the in-flight cross-check.
  //
  // Two classes of formula are skipped because the bytecode IR structurally
  // cannot reproduce the tree-walker result, so comparing would emit a false
  // mismatch:
  //
  //   1. Lazy-only calls (`has_lazy_only_call`): functions present in the
  //      lazy-dispatch table but absent from the eager registry. The VM has
  //      no AST at runtime and no eager impl to call, so it surfaces
  //      `#NAME?`. This covers IRR / MIRR / XIRR / XNPV, NETWORKDAYS /
  //      WORKDAY / REGEX* / TEXTSPLIT / PHONETIC, the higher-order array
  //      forms (MAP / REDUCE / SCAN / BYROW / BYCOL / MAKEARRAY), and the
  //      AST-introspecting info functions. Documented IR limitation.
  //
  //   2. IFERROR / IFNA eager-fallback drift: the bytecode lowers these as an
  //      eager `Call` with both arguments pre-evaluated (see
  //      `compiler.cpp::compile_iferror_or_ifna`), so the fallback is always
  //      evaluated even when the primary succeeds. True short-circuit would
  //      need a new "jump-if-not-error(-of-kind-NA)" opcode the IR does not
  //      have; that is a deferred IR change, not a correctness bug in the
  //      diagnostic-only VM. When the fallback raises a different error than
  //      the primary, the two paths legitimately diverge, so we skip those
  //      formulas here rather than chase a known limitation.
  if (has_lazy_only_call(node, registry)) {
    return v;
  }
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
