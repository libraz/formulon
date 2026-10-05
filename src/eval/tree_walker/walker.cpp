//
// Recursive node visitor for the tree-walk evaluator. Holds the public
// `evaluate()` overloads, the `eval_node` switch (declared with external
// linkage in `eval/lazy_impls.h` so lazy-impl TUs can recurse back into
// it), and the read-only spill-collision detector.
//
// The function-call dispatch path (`dispatch_call`, range-argument
// expansion) lives in `tree_walker/dispatch.cpp`, the runtime lambda
// invocation (`invoke_lambda`) in `tree_walker/lambda_invoke.cpp`; the
// array-broadcasting helpers (`broadcast_binop`, `broadcast_unary`,
// `apply_binop_per_cell`) live in `tree_walker/broadcast.cpp`. The
// TUs split the original monolithic `tree_walker.cpp` while
// preserving file-local helpers and the existing `formulon::eval`
// namespace shape (anonymous helpers are local to each TU).
//
// See `tree_walker.h` for the public contract.

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <limits>
#include <string_view>
#include <vector>

#include "eval/array_alloc.h"
#include "eval/declared_rect.h"
#include "eval/defined_name_resolve.h"
#include "eval/dynamic_array/anchor.h"
#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/external_ref.h"
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

namespace formulon {
namespace eval {
namespace {

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
  const std::vector<Value>& ev = expanded.value();
  ArrayValue* arr = array_from_values(rows, cols, ev.data(), ev.size(), arena);
  if (arr == nullptr) {
    return Value::error(ErrorCode::Num);
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
  parser::Reference lhs{};
  parser::Reference rhs{};
  if (!declared_rect_endpoint_pair(node, &lhs, &rhs)) {
    return false;
  }
  const Expected<DeclaredRect, ErrorCode> rect = declared_rect(lhs, rhs);
  if (!rect || !rect.value().whole_axis) {
    return false;
  }
  parser::Reference first{};
  parser::Reference last{};
  // The parser keeps the sheet qualifier on the left endpoint, matching how
  // `Sheet1!A1:B2` parses; `expand_range` inherits it for the rectangle.
  first.sheet = lhs.sheet;
  first.sheet_quoted = lhs.sheet_quoted;
  first.row = rect.value().row_first;
  first.col = rect.value().col_first;
  last.row = rect.value().row_last;
  last.col = rect.value().col_last;
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
  parser::Reference lhs{};
  parser::Reference rhs{};
  if (!declared_rect_endpoint_pair(node, &lhs, &rhs)) {
    return false;
  }
  const Expected<DeclaredRect, ErrorCode> rect = declared_rect(lhs, rhs);
  if (!rect || rect.value().whole_axis || rect.value().single_cell()) {
    return false;
  }
  parser::Reference first{};
  parser::Reference last{};
  // `eval_node` carries the left endpoint's qualifier onto both corners, and
  // `expand_range` lets the right corner inherit it either way; reproducing
  // it keeps the two entry points describing one rectangle.
  first.sheet = lhs.sheet;
  first.sheet_quoted = lhs.sheet_quoted;
  first.row = rect.value().row_first;
  first.col = rect.value().col_first;
  last.sheet = lhs.sheet;
  last.sheet_quoted = lhs.sheet_quoted;
  last.row = rect.value().row_last;
  last.col = rect.value().col_last;
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

// Evaluates the space-as-intersection operator: `A1:C3 B2:D4` denotes the
// overlapping rectangle (here B2:C3), and non-overlapping operands are
// `#NULL!`.
//
// The overlap is a reference like any other, so it materialises wherever
// `RangeOp` does. A 1x1 overlap degrades to its scalar exactly as `A1:A1`
// does; anything larger is the rectangle, which is what makes
// `=C1:D3 D1:E2` a 3x1 array rather than the single cell at its top-left.
//
// `spill_position` says whether the caller is the whole-formula position.
// There the rectangle is weighed against the anchor's footprint before it is
// built, and a refusal is recorded for the release machinery — the same
// treatment a bare range gets, for the same reason. Circularity is left to
// `resolve_ref` to report per cell: an intersection rectangle is bounded by
// its operands and has no whole-axis spelling Excel would canonicalise it
// into, so it is in the class that keeps the per-cell route.
//
// A nested position also keeps the `RangeOp` case's one exception: an
// overlap that still spans a whole grid axis (`A:C B:B`) collapses to its
// top-left there, so an operator cannot conjure a million cells per side.
Value eval_intersect_op(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx, bool spill_position) {
  std::string_view sheet;
  std::uint32_t r1 = 0;
  std::uint32_t c1 = 0;
  std::uint32_t r2 = 0;
  std::uint32_t c2 = 0;
  bool disjoint = false;
  ErrorCode err = ErrorCode::Value;
  if (!compute_intersect_rect(node.as_intersect_lhs(), node.as_intersect_rhs(), arena, registry, ctx, &sheet, &r1, &c1,
                              &r2, &c2, &disjoint, &err)) {
    return Value::error(err);
  }
  if (disjoint) {
    return Value::error(ErrorCode::Null);
  }
  parser::Reference top_left{};
  top_left.sheet = sheet;
  top_left.row = r1;
  top_left.col = c1;
  const bool full_height = r1 == 0U && r2 == Sheet::kMaxRows - 1U;
  const bool full_width = c1 == 0U && c2 == Sheet::kMaxCols - 1U;
  if ((r1 == r2 && c1 == c2) || (!spill_position && (full_height || full_width))) {
    return ctx.resolve_ref(top_left, arena, registry);
  }
  parser::Reference bottom_right{};
  bottom_right.sheet = sheet;
  bottom_right.row = r2;
  bottom_right.col = c2;
  if (spill_position) {
    return evaluate_bare_range_spill(top_left, bottom_right, arena, registry, ctx, /*settle_circularity=*/false);
  }
  return materialize_rectangle(top_left, bottom_right, arena, registry, ctx);
}

// True when `node` reads no lexical binding, so it evaluates the same in any
// scope. Conservative: a shape it does not know is not scope-free.
bool is_scope_free(const parser::AstNode& node, const NameEnv* env) {
  switch (node.kind()) {
    case parser::NodeKind::Literal:
    case parser::NodeKind::ErrorLiteral:
    case parser::NodeKind::Ref:
    case parser::NodeKind::Ref3D:
    case parser::NodeKind::ArrayLiteral:
      return true;
    case parser::NodeKind::SpillRef:
      return node.as_spill_ref_anchor_expr() == nullptr;
    case parser::NodeKind::UnaryOp:
      return is_scope_free(node.as_unary_operand(), env);
    case parser::NodeKind::BinaryOp:
      return is_scope_free(node.as_binary_lhs(), env) && is_scope_free(node.as_binary_rhs(), env);
    case parser::NodeKind::RangeOp:
      return is_scope_free(node.as_range_lhs(), env) && is_scope_free(node.as_range_rhs(), env);
    case parser::NodeKind::IntersectOp:
      return is_scope_free(node.as_intersect_lhs(), env) && is_scope_free(node.as_intersect_rhs(), env);
    case parser::NodeKind::UnionOp:
      for (std::uint32_t i = 0; i < node.as_union_arity(); ++i) {
        if (!is_scope_free(node.as_union_child(i), env)) {
          return false;
        }
      }
      return true;
    case parser::NodeKind::Call:
      if (env != nullptr && env->lookup(node.as_call_name()) != nullptr) {
        return false;
      }
      for (std::uint32_t i = 0; i < node.as_call_arity(); ++i) {
        if (!is_scope_free(node.as_call_arg(i), env)) {
          return false;
        }
      }
      return true;
    default:
      return false;
  }
}

// Resolves a binding source that returns a reference (a reference-returning
// call, a `:` over one, or an intersection) to the static `RangeOp` it names,
// so the binding carries that rectangle rather than the call. Returns nullptr
// for any other shape or when the call picks a non-reference (`INDEX({1,2},1)`).
const parser::AstNode* resolve_computed_reference(const parser::AstNode& expr, Arena& arena,
                                                  const FunctionRegistry& registry, const EvalContext& ctx) {
  if (expr.kind() == parser::NodeKind::Call) {
    const NameEnv* env = ctx.name_env();
    if (!is_reference_call_name(expr.as_call_name()) ||
        (env != nullptr && env->lookup(expr.as_call_name()) != nullptr)) {
      return nullptr;
    }
  } else if (expr.kind() == parser::NodeKind::RangeOp) {
    parser::Reference lhs{};
    parser::Reference rhs{};
    if (declared_rect_endpoint_pair(expr, &lhs, &rhs)) {
      return nullptr;
    }
  } else if (expr.kind() != parser::NodeKind::IntersectOp) {
    return nullptr;
  }
  std::string_view sheet;
  std::uint32_t top = 0;
  std::uint32_t left = 0;
  std::uint32_t bottom = 0;
  std::uint32_t right = 0;
  ErrorCode err = ErrorCode::Value;
  if (!resolve_reference_rect(expr, arena, registry, ctx, &sheet, &top, &left, &bottom, &right, &err)) {
    return nullptr;
  }
  parser::Reference first{};
  parser::Reference last{};
  first.sheet = sheet;
  last.sheet = sheet;
  if (top == 0U && bottom == Sheet::kMaxRows - 1U) {
    first.is_full_col = true;
    last.is_full_col = true;
    first.col = left;
    last.col = right;
  } else if (left == 0U && right == Sheet::kMaxCols - 1U) {
    first.is_full_row = true;
    last.is_full_row = true;
    first.row = top;
    last.row = bottom;
  } else {
    first.row = top;
    first.col = left;
    last.row = bottom;
    last.col = right;
  }
  // Always a `RangeOp`, even for one cell (`D1:D1`): a bare `Ref` binding is
  // the scalar-provenance shape, and a reference-returning call is not.
  parser::AstNode* lhs = parser::make_ref(arena, first);
  parser::AstNode* rhs = parser::make_ref(arena, last);
  if (lhs == nullptr || rhs == nullptr) {
    return nullptr;
  }
  return parser::make_range_op(arena, lhs, rhs);
}

}  // namespace

const parser::AstNode* resolve_binding_reference(const parser::AstNode& expr, Arena& arena,
                                                 const FunctionRegistry& registry, const EvalContext& ctx) {
  parser::Reference lhs{};
  parser::Reference rhs{};
  switch (expr.kind()) {
    case parser::NodeKind::Ref:
      return &expr;
    case parser::NodeKind::RangeOp:
      if (declared_rect_endpoint_pair(expr, &lhs, &rhs)) {
        return &expr;
      }
      break;
    case parser::NodeKind::NameRef: {
      const NameEnv* env = ctx.name_env();
      const parser::AstNode* bound =
          (env != nullptr && expr.as_name_sheet().empty()) ? env->lookup_ast(expr.as_name()) : nullptr;
      if (bound != nullptr && (bound->kind() == parser::NodeKind::Ref || bound->kind() == parser::NodeKind::RangeOp)) {
        return bound;
      }
      return nullptr;
    }
    default:
      break;
  }
  return resolve_computed_reference(expr, arena, registry, ctx);
}

Value eval_binding_source(const parser::AstNode& expr, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx, const parser::AstNode** out_ast) {
  const parser::AstNode* ref = resolve_binding_reference(expr, arena, registry, ctx);
  *out_ast = ref;
  // A union, or a range-shaped call naming no single rectangle (one picking a
  // union), is recorded as written when it reads the same in any scope.
  const bool unresolved_range =
      ref == nullptr && (expr.kind() == parser::NodeKind::UnionOp ||
                         (expr.kind() == parser::NodeKind::Call && is_range_shaped_ast(expr)));
  if (expr.kind() == parser::NodeKind::ArrayLiteral ||
      (expr.kind() == parser::NodeKind::SpillRef && expr.as_spill_ref_anchor_expr() == nullptr) ||
      (unresolved_range && is_scope_free(expr, ctx.name_env()))) {
    *out_ast = &expr;
  } else if (expr.kind() == parser::NodeKind::NameRef && expr.as_name_sheet().empty() && ctx.name_env() != nullptr) {
    *out_ast = ctx.name_env()->lookup_ast(expr.as_name());
  }
  // A reference is read where the bound name is used, not here.
  if (ref != nullptr) {
    return Value::blank();
  }
  return eval_node(expr, arena, registry, ctx);
}

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
        // An endpoint that is a reference-returning call (`A1:INDEX(...)`,
        // `A1:OFFSET(...)`) names its rectangle only once resolved; the
        // union then spills like any bounded range.
        std::string_view lhs_sheet;
        std::string_view rhs_sheet;
        std::uint32_t lt = 0;
        std::uint32_t ll = 0;
        std::uint32_t lb = 0;
        std::uint32_t lr = 0;
        std::uint32_t rt = 0;
        std::uint32_t rl = 0;
        std::uint32_t rb = 0;
        std::uint32_t rr = 0;
        ErrorCode err = ErrorCode::Value;
        if (!resolve_range_endpoint(lhs, arena, registry, ctx, &lhs_sheet, &lt, &ll, &lb, &lr, &err) ||
            !resolve_range_endpoint(rhs, arena, registry, ctx, &rhs_sheet, &rt, &rl, &rb, &rr, &err)) {
          return Value::error(err);
        }
        parser::Reference top_left{};
        if (!merge_range_endpoint_sheets(lhs, lhs_sheet, rhs, rhs_sheet, ctx, &top_left.sheet, &err)) {
          return Value::error(err);
        }
        top_left.row = std::min(lt, rt);
        top_left.col = std::min(ll, rl);
        parser::Reference bottom_right{};
        bottom_right.sheet = top_left.sheet;
        bottom_right.row = std::max(lb, rb);
        bottom_right.col = std::max(lr, rr);
        if (top_left.row == bottom_right.row && top_left.col == bottom_right.col) {
          return ctx.resolve_ref(top_left, arena, registry);
        }
        return materialize_rectangle(top_left, bottom_right, arena, registry, ctx);
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
      // Implicit intersection on a reference: project the formula cell onto
      // the declared rectangle. All three spellings that declare one —
      // bounded `Ref:Ref`, full-axis `Ref:Ref`, and the single `Ref` the
      // parser folds `A:A` / `1:1` into — go through the projection shared
      // with `_xlfn.SINGLE`, so `=@A:B` and `=@A1:B3` agree wherever they
      // denote the same rectangle.
      if (ctx.has_formula_cell()) {
        parser::Reference target{};
        switch (project_implicit_intersection(operand, ctx.formula_row(), ctx.formula_col(), &target)) {
          case IntersectionProjection::kCell:
            return ctx.resolve_ref(target, arena, registry);
          case IntersectionProjection::kNoCell:
            return Value::error(ErrorCode::Value);
          case IntersectionProjection::kNotStaticReference:
            break;
        }
        if (Value projected = Value::blank(); project_reference_result(operand, arena, registry, ctx, &projected)) {
          return projected;
        }
      } else if (operand.kind() == parser::NodeKind::RangeOp) {
        // No formula-cell context (top-level evaluator entry) -> degrade to
        // top-left, matching the bare-range fallback. Production calls
        // through Workbook always supply a formula cell, so this branch
        // only fires for parser-driven smoke tests.
        return eval_node(operand, arena, registry, ctx);
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
      //   * Computed anchor that names no single cell      -> #REF!
      //
      // The returned ArrayValue header lives in the eval `arena`; its cells
      // buffer holds shallow copies of the SpillRegion's `Value`s. Text
      // payloads point into `SpillRegion::owned_strings`, which lives on
      // Sheet and outlives any single evaluation arena (zero-copy reuse).
      std::string_view anchor_sheet;
      std::uint32_t anchor_row = 0;
      std::uint32_t anchor_col = 0;
      ErrorCode spill_err = ErrorCode::Ref;
      if (!resolve_spill_anchor_node(node, arena, registry, ctx, &anchor_sheet, &anchor_row, &anchor_col, &spill_err)) {
        return Value::error(spill_err);
      }
      ArrayValue* arr = project_spill_at_anchor(anchor_sheet, anchor_row, anchor_col, arena, ctx, &spill_err);
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
      // the name is undefined in scope. `Sheet1!Name` is never a lexical
      // binding and is looked up in Sheet1's scope.
      if (const std::string_view sheet = node.as_name_sheet(); !sheet.empty()) {
        return resolve_sheet_defined_name(sheet, node.as_name(), arena, registry, ctx);
      }
      const NameEnv* env = ctx.name_env();
      if (env != nullptr) {
        const auto read = [&](const parser::AstNode& ref) { return eval_node(ref, arena, registry, ctx); };
        if (const Value* bound = env->lookup_or_read(node.as_name(), read); bound != nullptr) {
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
      //   * A reference initialiser binds as a reference, by the rule every
      //     LAMBDA argument follows too (see `eval_binding_source`).
      NameEnv env;
      const NameEnv* parent = ctx.name_env();
      // Start from whatever the caller supplied; extending `NameEnv` makes
      // `env` point at a new head frame while preserving the parent chain.
      if (parent != nullptr) {
        env = *parent;
      }
      const std::uint32_t count = node.as_let_binding_count();
      for (std::uint32_t i = 0; i < count; ++i) {
        const parser::AstNode* bound_ast = nullptr;
        const Value v =
            eval_binding_source(node.as_let_binding_expr(i), arena, registry, ctx.with_name_env(&env), &bound_ast);
        env = env.extend(node.as_let_binding_name(i), v, bound_ast, arena);
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

    case parser::NodeKind::ExternalRef:
      // `[0]!Name` names a defined name of this workbook itself.
      if (parser::is_self_book_name_ref(node)) {
        return resolve_self_book_defined_name(node.as_external_ref_name(), arena, registry, ctx);
      }
      // Read straight out of the external-link cache. Unlike `Ref3D` this
      // needs no sheet resolution against the workbook: the target lives
      // in another file whose grid the cache already holds, so a
      // rectangle materialises here rather than being routed through
      // `expand_range`.
      return resolve_external_ref(node, arena, ctx);

    case parser::NodeKind::StructuredRef: {
      // Resolve the table reference (`Table[Col]`, `Table[#All]`, ...) to
      // a concrete rectangle and read it through the shared projection.
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
      lv->name_scope_sheet = ctx.name_scope_sheet();
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
      // Calling a reference (`A1(1)`, `Sheet1!LOG10(100)`, `(A1:A2)(1)`) is
      // #REF! in Excel, whatever the cells hold.
      if (is_reference_shape(node.as_lambda_call_callee())) {
        return Value::error(ErrorCode::Ref);
      }
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
      // before invoking.
      //
      // Only the arguments the call actually writes are passed. Trailing
      // optional parameters the call omits are bound to the omitted
      // sentinel by `invoke_lambda` itself, which is what `ISOMITTED` reads
      // inside the body; nothing about omission is decided here.
      const std::uint32_t arity = node.as_lambda_call_arity();
      std::vector<const parser::AstNode*> argv;
      argv.reserve(arity);
      for (std::uint32_t i = 0; i < arity; ++i) {
        argv.push_back(&node.as_lambda_call_arg(i));
      }
      return invoke_lambda(callee.as_lambda(), arity, argv.empty() ? nullptr : argv.data(), arena, registry, ctx);
    }

    case parser::NodeKind::IntersectOp:
      // A nested position, so the rectangle is materialised rather than
      // weighed against a spill footprint — the same split the `RangeOp`
      // case above draws against `evaluate()`.
      return eval_intersect_op(node, arena, registry, ctx, /*spill_position=*/false);

    case parser::NodeKind::ArrayLiteral: {
      // A brace literal is a first-class dynamic array in modern Excel, so
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

  // The value of the whole formula is produced by exactly one of four
  // branches: the bare-range spilling position, the intersect operator's
  // spilling position, the iterative-calculation driver, or the ordinary
  // single-pass walk.
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
  } else if (is_top_level && node.kind() == parser::NodeKind::IntersectOp) {
    // The intersect operator names a rectangle just as `:` does, so the
    // spelling that produces one is intercepted here for the same reason:
    // this is the only position it can spill from, and therefore the only
    // one that may weigh the rectangle against the anchor's footprint.
    v = eval_intersect_op(node, arena, registry, ctx_with_counters, /*spill_position=*/true);
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
  return v;
}

}  // namespace eval
}  // namespace formulon
