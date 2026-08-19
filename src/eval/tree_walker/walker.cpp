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
#include <limits>
#include <optional>
#include <string_view>
#include <vector>

#include "eval/array_alloc.h"
#include "eval/defined_name_resolve.h"
#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/implicit_intersection.h"
#include "eval/iterative_solver.h"
#include "eval/lambda_value.h"
#include "eval/lazy_impls.h"
#include "eval/name_env.h"
#include "eval/name_env_resolve.h"
#include "eval/range_resolvers.h"
#include "eval/spill_anchor.h"
#include "eval/structured_ref.h"
#include "eval/structured_ref_project.h"
#include "eval/tree_walker.h"
#include "eval/tree_walker/broadcast.h"
#include "eval/tree_walker/depth_guard.h"
#include "eval/tree_walker/dispatch.h"
#include "eval/tree_walker_lazy_table.h"
#include "parser/ast.h"
#include "sheet.h"
#include "sheet_name.h"
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

// Materialises the inclusive rectangle [`top_left`, `bottom_right`] as a
// `Value::Array`. Both endpoints must be bounded (no `is_full_col` /
// `is_full_row`) and already normalised so that `top_left` is the smaller
// corner on both axes. Cells the expansion did not produce are padded with
// Blank, which the top-level surface contract later projects to 0.
Value materialize_rectangle(const parser::Reference& top_left, const parser::Reference& bottom_right, Arena& arena,
                            const FunctionRegistry& registry, const EvalContext& ctx) {
  auto expanded = ctx.expand_range(top_left, bottom_right, arena, registry);
  if (!expanded) {
    return Value::error(expanded.error());
  }
  const std::uint32_t rows = bottom_right.row - top_left.row + 1U;
  const std::uint32_t cols = bottom_right.col - top_left.col + 1U;
  Value* buffer = nullptr;
  ArrayValue* arr = allocate_array_value(rows, cols, arena, buffer, kMaxDerivedArrayCells);
  if (arr == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  const std::size_t total = static_cast<std::size_t>(rows) * static_cast<std::size_t>(cols);
  const std::vector<Value>& ev = expanded.value();
  for (std::size_t k = 0; k < total; ++k) {
    buffer[k] = k < ev.size() ? ev[k] : Value::blank();
  }
  return Value::array(arr);
}

// Recovers the rectangle a bare whole-axis reference declares, so it can be
// evaluated through the same path as any other bare range.
//
// Excel 365 gives `A:A` / `A:C` / `1:2` the array of the *declared*
// rectangle — rows 1..1048576 for a column span, columns A..XFD for a row
// span — not the populated extent. Nothing is trimmed, which is why an
// occupied cell far below the data still blocks the spill.
//
// `A:A` and `1:1` parse as a single `Ref` carrying `is_full_col` /
// `is_full_row`; the multi-span forms parse as a `RangeOp` over two
// same-axis whole `Ref`s. Both shapes are recognised here. A mixed-axis
// pair (`A:1`) has no rectangle and is rejected, as is an endpoint outside
// the grid; the caller then keeps its existing scalar degradation.
//
// Returns false when `node` is not a bare whole-axis reference, leaving the
// out-params untouched.
bool whole_axis_declared_rect(const parser::AstNode& node, parser::Reference* top_left,
                              parser::Reference* bottom_right) {
  const parser::Reference* lhs = nullptr;
  const parser::Reference* rhs = nullptr;
  if (node.kind() == parser::NodeKind::Ref) {
    lhs = &node.as_ref();
    rhs = lhs;
  } else if (node.kind() == parser::NodeKind::RangeOp) {
    const parser::AstNode& lhs_ast = node.as_range_lhs();
    const parser::AstNode& rhs_ast = node.as_range_rhs();
    if (lhs_ast.kind() != parser::NodeKind::Ref || rhs_ast.kind() != parser::NodeKind::Ref) {
      return false;
    }
    lhs = &lhs_ast.as_ref();
    rhs = &rhs_ast.as_ref();
  } else {
    return false;
  }

  const bool full_col = lhs->is_full_col && rhs->is_full_col;
  const bool full_row = lhs->is_full_row && rhs->is_full_row;
  // Same-axis pairs only. A reference flagged on both axes is malformed and
  // an axis mismatch has no rectangle; either way this is not our shape.
  if (full_col == full_row) {
    return false;
  }

  parser::Reference first{};
  parser::Reference last{};
  // The parser keeps the sheet qualifier on the left endpoint, matching how
  // `Sheet1!A1:B2` parses; `expand_range` inherits it for the rectangle.
  first.sheet = lhs->sheet;
  first.sheet_quoted = lhs->sheet_quoted;
  if (full_col) {
    if (lhs->col >= Sheet::kMaxCols || rhs->col >= Sheet::kMaxCols) {
      return false;
    }
    first.row = 0U;
    first.col = std::min(lhs->col, rhs->col);
    last.row = Sheet::kMaxRows - 1U;
    last.col = std::max(lhs->col, rhs->col);
  } else {
    if (lhs->row >= Sheet::kMaxRows || rhs->row >= Sheet::kMaxRows) {
      return false;
    }
    first.row = std::min(lhs->row, rhs->row);
    first.col = 0U;
    last.row = std::max(lhs->row, rhs->row);
    last.col = Sheet::kMaxCols - 1U;
  }
  *top_left = first;
  *bottom_right = last;
  return true;
}

// Recovers the rectangle a bare bounded range declares (`A1:C10`), so the
// spelling that names both corners can be measured before it is built.
//
// This is the same rectangle `eval_node`'s `RangeOp` case materialises,
// recovered one level up because only the whole-formula position can spill
// and therefore only that position may consult a footprint. The endpoints
// are normalised so the top-left corner comes first, exactly as there.
//
// Three shapes are deliberately left to `eval_node`: a whole-axis endpoint,
// which `whole_axis_declared_rect` recognises instead; a single-cell range
// (`A1:A1`), which keeps its scalar degradation; and an endpoint outside the
// grid, which must keep surfacing the `#REF!` range expansion gives it
// rather than being answered by the footprint.
//
// Returns false for anything else, leaving the out-params untouched.
bool bounded_declared_rect(const parser::AstNode& node, parser::Reference* top_left, parser::Reference* bottom_right) {
  if (node.kind() != parser::NodeKind::RangeOp) {
    return false;
  }
  const parser::AstNode& lhs_ast = node.as_range_lhs();
  const parser::AstNode& rhs_ast = node.as_range_rhs();
  if (lhs_ast.kind() != parser::NodeKind::Ref || rhs_ast.kind() != parser::NodeKind::Ref) {
    return false;
  }
  const parser::Reference& lhs = lhs_ast.as_ref();
  const parser::Reference& rhs = rhs_ast.as_ref();
  if (lhs.is_full_col || lhs.is_full_row || rhs.is_full_col || rhs.is_full_row) {
    return false;
  }
  if (lhs.row >= Sheet::kMaxRows || rhs.row >= Sheet::kMaxRows || lhs.col >= Sheet::kMaxCols ||
      rhs.col >= Sheet::kMaxCols) {
    return false;
  }
  const std::uint32_t r1 = std::min(lhs.row, rhs.row);
  const std::uint32_t r2 = std::max(lhs.row, rhs.row);
  const std::uint32_t c1 = std::min(lhs.col, rhs.col);
  const std::uint32_t c2 = std::max(lhs.col, rhs.col);
  if (r1 == r2 && c1 == c2) {
    return false;
  }
  parser::Reference first{};
  parser::Reference last{};
  // `eval_node` carries the left endpoint's qualifier onto both corners, and
  // `expand_range` lets the right corner inherit it either way; reproducing
  // it keeps the two entry points describing one rectangle.
  first.sheet = lhs.sheet;
  first.sheet_quoted = lhs.sheet_quoted;
  first.row = r1;
  first.col = c1;
  last.sheet = lhs.sheet;
  last.sheet_quoted = lhs.sheet_quoted;
  last.row = r2;
  last.col = c2;
  *top_left = first;
  *bottom_right = last;
  return true;
}

// Evaluates a bare range standing as the entire formula — the only position
// in which a range spills, and so the only one where a footprint may be
// consulted.
//
// The outcomes are tried in the order Excel decides them, which is also the
// only order that is affordable:
//
//   1. The rectangle covers the formula's own cell. The formula then reads
//      its own result, which is a circular reference and never spills —
//      whatever else occupies the rectangle. This is settled first, so a
//      self-referential formula does not get answered by the footprint.
//      `settle_circularity` selects it; see the parameter note below.
//   2. The footprint cannot be placed, either because the rectangle leaves
//      the grid measured from the anchor or because something occupies it.
//      `Sheet::probe_spill_footprint` decides that without building any
//      values, which is what keeps `=A:C` — and equally `=A1:C1048576` — at
//      Z2 from allocating three million cells to return `#SPILL!`.
//   3. Otherwise the rectangle is materialised and spills.
//
// `settle_circularity` decides whether the cycle is pre-empted here or left
// to be reported per cell by `resolve_ref` inside the expansion. It is on
// for the whole-axis spelling and for a bounded rectangle spanning a full
// grid axis; it is off for every smaller rectangle. The line is drawn where
// Excel draws it: a full-height or full-width bounded range is rewritten
// into the whole-axis spelling on entry, so `=A1:XFD1048576` and `=A:XFD`
// are not two ways of writing one rectangle but literally one formula, and
// two answers for one formula are indefensible whatever the answers are. A
// rectangle Excel does not canonicalise (`=A1:C3`) has no twin to disagree
// with and keeps the per-cell route.
//
// For the committing driver this is a difference in cost, not in verdict:
// the materialised footprint feeds a self-edge back into the dependency
// graph and the engine's cycle policy reaches `#REF!` on its own, after
// building the rectangle. Pre-empting reaches the same answer without
// building it — the same shape as the footprint pre-check above. The
// read-only driver has no graph to fall back on, which is where the two
// spellings visibly disagreed.
//
// The pre-emption does not introduce a divergence from Excel. Excel
// abandons the closure for a self-covering rectangle and leaves 0, while
// Formulon reports every cycle member as `#REF!`; that is the registered
// engine-wide policy, and this extends it to a spelling Excel treats as
// identical to one already covered by it.
//
// Without a formula cell to anchor against there is no footprint to
// measure and nothing that could spill, so the rectangle is materialised
// directly; that is the shape ad-hoc parser-level evaluation sees.
Value evaluate_bare_range_spill(const parser::Reference& top_left, const parser::Reference& bottom_right, Arena& arena,
                                const FunctionRegistry& registry, const EvalContext& ctx, bool settle_circularity) {
  const Sheet* sheet = ctx.current_sheet();
  if (sheet == nullptr || !ctx.has_formula_cell()) {
    return materialize_rectangle(top_left, bottom_right, arena, registry, ctx);
  }
  const std::uint32_t anchor_row = ctx.formula_row();
  const std::uint32_t anchor_col = ctx.formula_col();

  // The rectangle is on the formula's own sheet when it carries no
  // qualifier, or when the qualifier names that sheet.
  const bool same_sheet = top_left.sheet.empty() || sheet_names::equal(top_left.sheet, sheet->name());
  if (settle_circularity && same_sheet && anchor_row >= top_left.row && anchor_row <= bottom_right.row &&
      anchor_col >= top_left.col && anchor_col <= bottom_right.col) {
    // Circular, and reported as the engine reports every other cycle
    // rather than as anything specific to this shape: `#REF!` with
    // iterative calculation off, the cell's last computed value when it is
    // on. That mirrors `EvalContext::resolve_ref`'s back-edge branch and
    // the sentinel `RecalcEngine` writes for each member of a cyclic
    // component. It cannot be delegated to `resolve_ref` here, because the
    // dependency-ordered recalc path evaluates with no `EvalState` and
    // short-circuits formula references to their cached values — asking it
    // would yield the anchor's stale value instead of a cycle. Excel
    // leaves 0 in the cell, which Formulon deliberately does not match.
    const Workbook* workbook = ctx.workbook();
    if (workbook != nullptr && workbook->iterative_options().enabled) {
      return sheet->resolve_cell_value(anchor_row, anchor_col);
    }
    return Value::error(ErrorCode::Ref);
  }

  const std::uint32_t rows = bottom_right.row - top_left.row + 1U;
  const std::uint32_t cols = bottom_right.col - top_left.col + 1U;
  if (sheet->probe_spill_footprint(anchor_row, anchor_col, rows, cols) != Sheet::SpillAdmission::kAdmissible) {
    // A refusal has to be recorded, not just returned: the remembered
    // rectangle is what lets the release machinery retry this anchor once
    // the blocker goes away. A read-only context cannot record anything,
    // and commits nothing either, so it simply reports the error.
    if (Sheet* target = ctx.mutable_sheet(); target == sheet) {
      target->reject_spill_footprint(anchor_row, anchor_col, rows, cols);
    }
    return Value::error(ErrorCode::Spill);
  }
  return materialize_rectangle(top_left, bottom_right, arena, registry, ctx);
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
      // Excel 365 dynamic-array semantics: a bare bounded range used in a
      // value context spills. It evaluates to the whole rectangle as a
      // Value::Array, which bubbles up to the cell entry point (committing a
      // spill) or to an enclosing operator's cellwise broadcast, rather than
      // collapsing to a single implicit-intersection cell. Legacy implicit
      // intersection is reached only through the explicit `@` / SINGLE
      // wrapper (the ImplicitIntersection case below), never here.
      //
      // Two shapes keep the top-left anchor projection here:
      //   * A whole-column / whole-row endpoint (`A:C`, `1:3`): its declared
      //     rectangle spans a whole grid axis, which is only materialised
      //     when the reference is the entire formula and can therefore
      //     spill. `evaluate()` intercepts that position; every other value
      //     context (an operand, a scalar function argument) keeps the
      //     anchor projection so an unbounded rectangle is not conjured
      //     behind an operator. See `whole_axis_declared_rect`.
      //   * A single-cell (`A1:A1`) range: degrades to the scalar so the
      //     degenerate surface is unchanged.
      //
      // A bounded rectangle reaches the materialisation below only from a
      // nested position. In the spilling position `evaluate()` intercepts it
      // too, so that the footprint refusing it is measured rather than
      // discovered by building the rectangle; the rectangle it materialises
      // when admitted is the one built here. See `bounded_declared_rect`.
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

      parser::Reference top_left{};
      top_left.sheet = lhs_ref.sheet;
      top_left.row = r1;
      top_left.col = c1;

      const bool whole = lhs_ref.is_full_col || lhs_ref.is_full_row || rhs_ref.is_full_col || rhs_ref.is_full_row;
      if (whole || (r1 == r2 && c1 == c2)) {
        return ctx.resolve_ref(top_left, arena, registry);
      }

      parser::Reference bottom_right{};
      bottom_right.sheet = lhs_ref.sheet;
      bottom_right.row = r2;
      bottom_right.col = c2;
      return materialize_rectangle(top_left, bottom_right, arena, registry, ctx);
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
        const std::optional<parser::Reference> target =
            project_implicit_intersection(lhs_ast.as_ref(), rhs_ast.as_ref(), ctx.formula_row(), ctx.formula_col());
        if (!target.has_value()) {
          return Value::error(ErrorCode::Value);
        }
        return ctx.resolve_ref(*target, arena, registry);
      }
      // Dynamic arrays produced by a call, spill reference, or expression no
      // longer retain static range coordinates. Excel's `@` takes their
      // top-left element instead of allowing the value to spill.
      return implicit_intersect_value(eval_node(operand, arena, registry, ctx));
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
      ErrorCode spill_err = ErrorCode::Ref;
      ArrayValue* arr = project_spill_at_anchor(r.sheet, r.row, r.col, arena, ctx, &spill_err);
      if (arr == nullptr) {
        return Value::error(spill_err);
      }
      return Value::array(arr);
    }

    case parser::NodeKind::NameRef: {
      // Resolution order: lexical scope (LET / LAMBDA bindings) wins over a
      // workbook / sheet-scoped defined name, so `=LET(Rate, 2, Rate)` reads
      // the binding, not a `Rate` defined name. When no binding matches, fall
      // through to defined-name resolution, which returns `#NAME?` itself when
      // the name is undefined in scope.
      const NameEnv* env = ctx.name_env();
      if (env != nullptr) {
        if (const Value* bound = env->lookup(node.as_name()); bound != nullptr) {
          return *bound;
        }
      }
      return resolve_defined_name(node.as_name(), arena, registry, ctx);
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
      // Resolve the table reference (`Table[Col]`, `Table[#All]`, ...) to
      // a concrete rectangle and read it through the shared projection,
      // which the bytecode VM's `StructRef` opcode also runs.
      bool arena_exhausted = false;
      return project_structured_ref(node.as_structured_ref_table(), node.as_structured_ref_column(), arena, registry,
                                    ctx, &arena_exhausted);
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

    case parser::NodeKind::ArrayLiteral: {
      // A brace literal is a first-class dynamic array in modern Excel. The
      // bytecode VM already lowers it to MakeArray; mirror that shape here so
      // the tree walker can spill it, broadcast it through an operator, and
      // let `@` reduce it through the common implicit-intersection path.
      const std::uint32_t rows = node.as_array_rows();
      const std::uint32_t cols = node.as_array_cols();
      Value* cells = nullptr;
      ArrayValue* array = allocate_array_value(rows, cols, arena, cells, kMaxDerivedArrayCells);
      if (array == nullptr) {
        return Value::error(ErrorCode::Num);
      }
      for (std::uint32_t row = 0; row < rows; ++row) {
        for (std::uint32_t col = 0; col < cols; ++col) {
          cells[static_cast<std::size_t>(row) * cols + col] =
              eval_node(node.as_array_element(row, col), arena, registry, ctx);
        }
      }
      return Value::array(array);
    }

    // -- Unsupported range-producing operator ------------------------------
    case parser::NodeKind::UnionOp:
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

  // The value of the whole formula is produced by exactly one of three
  // branches: the bare-range spilling position, the iterative-calculation
  // driver, or the ordinary single-pass walk.
  //
  // A bare range standing as the entire formula — whole-axis (`=A:A`,
  // `=A:C`, `=1:2`) or bounded (`=A1:C10`) — is a spilling expression: its
  // value is the array of the declared rectangle, anchored at the formula
  // cell. Both spellings are intercepted here rather than in `eval_node`
  // because this is the only position where the result can actually spill,
  // and therefore the only one that may weigh the rectangle against the
  // anchor's footprint. An operand or a scalar function argument reaches
  // `eval_node` instead and is unaffected: the whole-axis form keeps its
  // top-left anchor projection there, the bounded form its plain
  // materialisation.
  //
  // Routing both through one place is what keeps them from disagreeing on
  // cost. Two spellings of one rectangle already returned one answer, but
  // only the whole-axis form measured the footprint before building it, so
  // the bounded form paid a full materialisation — 3.1 million cells for
  // `=A1:C1048576` — to arrive at the same `#SPILL!`. The rectangle is
  // bounded by the same range-expansion ceiling either way.
  //
  // Excel's observed consequences — the rectangle must fit measured from
  // the anchor, an occupied cell anywhere inside the declared rectangle
  // blocks it, and unpopulated cells spill as 0 — are pinned by
  // tests/oracle/cases/whole_axis_spill.yaml.
  //
  // One consequence in that suite is deliberately NOT matched. A formula
  // sitting inside the axis it references is circular; Excel resolves the
  // cycle to 0 with iterative calculation off, while Formulon's engine-wide
  // policy reports `#REF!` for every cycle member. That difference is a
  // registered divergence, not an observation this code reproduces.
  // `evaluate_bare_range_spill` settles circularity before the footprint so
  // the cycle cannot be pre-empted by a blocker.
  //
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
  Value v = Value::blank();
  parser::Reference bare_range_top{};
  parser::Reference bare_range_bottom{};
  if (is_top_level && whole_axis_declared_rect(node, &bare_range_top, &bare_range_bottom)) {
    v = evaluate_bare_range_spill(bare_range_top, bare_range_bottom, arena, registry, ctx_with_counters,
                                  /*settle_circularity=*/true);
  } else if (is_top_level && bounded_declared_rect(node, &bare_range_top, &bare_range_bottom)) {
    // Excel rewrites a bounded range spanning a full grid axis into the
    // whole-axis spelling on entry, so exactly this set has a twin above
    // whose answer it must match. Anything narrower has no twin.
    const bool full_height = bare_range_top.row == 0U && bare_range_bottom.row == Sheet::kMaxRows - 1U;
    const bool full_width = bare_range_top.col == 0U && bare_range_bottom.col == Sheet::kMaxCols - 1U;
    v = evaluate_bare_range_spill(bare_range_top, bare_range_bottom, arena, registry, ctx_with_counters,
                                  /*settle_circularity=*/full_height || full_width);
  } else if (is_top_level && !ctx.iterative_driver_suppressed() && ctx.has_formula_cell() &&
             ctx.current_sheet() != nullptr && ctx.workbook() != nullptr &&
             ctx.workbook()->iterative_options().enabled) {
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
      // Fixed-point convergence is defined only for numeric values. A
      // nonnumeric result cannot become convergent by repeating an otherwise
      // independent top-level formula, so retain its first result and let the
      // shared surface contract below render it (#CALC!, #SPILL!, etc.).
      if (!next.is_number()) {
        current = next;
        break;
      }
      // Convergence test: absolute change of the numeric value.
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
    // Fall through to the shared top-level surface contract below.
    v = current;
  } else {
    v = eval_node(node, arena, registry, ctx_with_counters);
  }
  if (arena.exhausted()) {
    if (EvalState* state = ctx.state(); state != nullptr) {
      state->mark_out_of_memory();
    }
    return Value::error(ErrorCode::Num);
  }
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
    // `probe_spill_footprint` gives the same verdict as the committing path
    // at a cost proportional to what the sheet stores rather than to the
    // rectangle's area, which matters now that a grid-axis result can reach
    // here with 1,048,576 cells.
    if (ctx.current_sheet()->probe_spill_footprint(ctx.formula_row(), ctx.formula_col(), rows, cols) !=
        Sheet::SpillAdmission::kAdmissible) {
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
  // The same blank -> 0 grid contract applies per cell to a spilled raw
  // range: Excel renders a blank source cell inside a spilled `=A1:A3` as
  // 0. Raw-grid ingress marks those copied cells in the shared range seam;
  // other array-producing functions project only blanks explicitly marked by
  // their producer (e.g. generated EXPAND pads). Genuine source blanks remain
  // Blank for nested consumers and neutral producers. Rebuild only when a
  // conversion is required.
  if (v.is_array()) {
    const ArrayValue* arr = v.as_array();
    const std::size_t n = static_cast<std::size_t>(arr->rows) * static_cast<std::size_t>(arr->cols);
    bool any_projection = false;
    for (std::size_t i = 0; i < n; ++i) {
      const Value& cell = arr->cells[i];
      if (cell.blank_projects_to_zero()) {
        any_projection = true;
        break;
      }
    }
    if (any_projection) {
      Value* buffer = nullptr;
      ArrayValue* promoted = allocate_array_value(arr->rows, arr->cols, arena, buffer, kMaxDerivedArrayCells);
      if (promoted == nullptr) {
        if (arena.exhausted()) {
          if (EvalState* state = ctx.state(); state != nullptr) {
            state->mark_out_of_memory();
          }
        }
        return Value::error(ErrorCode::Num);
      }
      for (std::size_t i = 0; i < n; ++i) {
        const Value& cell = arr->cells[i];
        buffer[i] = cell.blank_projects_to_zero() ? Value::number(0.0) : cell;
      }
      v = Value::array(promoted);
    }
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
