//
// Top-level evaluation entry of the tree-walk evaluator: the public
// `evaluate*` overloads, the iterative-calculation fixed-point driver, and the
// surface contract applied to the root value (spill footprint, blank -> 0,
// lambda -> #CALC!).

#include <cmath>
#include <cstddef>
#include <cstdint>

#include "eval/array_alloc.h"
#include "eval/dynamic_array/anchor.h"
#include "eval/eval_clock_scope.h"
#include "eval/eval_context.h"
#include "eval/eval_profile_scope.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/iterative_solver.h"
#include "eval/lazy_impls.h"
#include "eval/spill_anchor.h"
#include "eval/tail_array.h"
#include "eval/tree_walker.h"
#include "eval/tree_walker/dispatch.h"
#include "eval/tree_walker/walker_internal.h"
#include "parser/ast.h"
#include "sheet.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {

namespace {

// Evaluates `node` as a fixed-point iteration against the bound workbook's
// iterative-calculation options. The result goes through the shared top-level
// surface contract in the caller.
Value run_iterative_fixed_point(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                                const EvalContext& ctx) {
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
    EvalContext pass_ctx = ctx.with_state(pass_state);
    Value next = densify(eval_root_shaped(node, arena, registry, pass_ctx), arena);
    // Apply the blank -> 0 surface contract so a blank-resolving pass is
    // comparable to a numeric one.
    if (next.is_blank() && node.kind() != parser::NodeKind::Literal) {
      next = Value::number(0.0);
    }
    // Fixed-point convergence is defined only for numeric values. A
    // nonnumeric result cannot become convergent by repeating an otherwise
    // independent top-level formula, so retain its first result.
    if (!next.is_number()) {
      return next;
    }
    const bool converged = std::fabs(next.as_number() - current.as_number()) < iopts.max_change;
    current = next;
    if (converged) {
      break;
    }
  }
  return current;
}

// What the top-level evaluation makes of an array result.
enum class TopResult : std::uint8_t {
  kSpill,         // weighed against the anchor's spill footprint
  kFirstElement,  // reduced to its first element
  kWholeArray,    // returned whole, with no footprint check
};

// The one top-level evaluation behind `evaluate`, `evaluate_first_element` and
// `evaluate_array_formula`. They differ only in what an array result becomes.
Value evaluate_top_body(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx, TopResult mode) {
  const bool first_element_only = mode == TopResult::kFirstElement;
  const ScopedEvalProfile profile_scope(ctx.excel_profile());
  const ScopedEvalClock clock_scope(ctx.pinned_now());

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
  const bool spill_position = is_top_level && mode == TopResult::kSpill;

  // The value of the whole formula is produced by exactly one of four
  // branches: the bare-range spilling position, the intersect operator's
  // spilling position, the iterative-calculation driver, or the ordinary
  // single-pass walk.
  //
  // A bare range standing as the entire formula -- whole-axis (`=A:A`,
  // `=A:C`, `=1:2`) or bounded (`=A1:C10`) -- is a spilling expression: its
  // value is the array of the declared rectangle, anchored at the formula
  // cell. Both spellings are intercepted here rather than in `eval_node`
  // because this is the only position where the result can actually spill,
  // and therefore the only one that may weigh the rectangle against the
  // anchor's footprint before building it.
  //
  // Excel's observed consequences -- the rectangle must fit measured from
  // the anchor, an occupied cell anywhere inside the declared rectangle
  // blocks it, and unpopulated cells spill as 0 -- are pinned by
  // tests/oracle/cases/whole_axis_spill.yaml.
  //
  // A formula sitting inside the axis it references is circular, and
  // answered by the engine's cycle policy: 0 when the formula cell is the
  // cycle's only member, as Excel leaves it with iterative calculation off.
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
  Shaped sv;
  bool footprint_probed = false;
  parser::Reference bare_range_top{};
  parser::Reference bare_range_bottom{};
  if (spill_position && whole_axis_declared_rect(node, &bare_range_top, &bare_range_bottom, ctx.workbook())) {
    sv = evaluate_bare_range_spill(bare_range_top, bare_range_bottom, arena, registry, ctx_with_counters,
                                   BareRangeSpillOptions{/*settle_circularity=*/true, /*allow_tail=*/true},
                                   &footprint_probed);
  } else if (spill_position && bounded_declared_rect(node, &bare_range_top, &bare_range_bottom, ctx.workbook())) {
    // Excel rewrites a bounded range spanning a full grid axis into the
    // whole-axis spelling on entry, so exactly this set has a twin above
    // whose answer it must match. Anything narrower has no twin.
    const bool full_height = spans_all_rows(bare_range_top.row, bare_range_bottom.row);
    const bool full_width = spans_all_cols(bare_range_top.col, bare_range_bottom.col);
    sv = evaluate_bare_range_spill(bare_range_top, bare_range_bottom, arena, registry, ctx_with_counters,
                                   BareRangeSpillOptions{/*settle_circularity=*/full_height || full_width,
                                                         /*allow_tail=*/true},
                                   &footprint_probed);
  } else if (spill_position && node.kind() == parser::NodeKind::IntersectOp) {
    // The intersect operator names a rectangle just as `:` does, so the
    // spelling that produces one is intercepted here for the same reason:
    // this is the only position it can spill from, and therefore the only
    // one that may weigh the rectangle against the anchor's footprint.
    sv.value = eval_intersect_op(node, arena, registry, ctx_with_counters, /*spill_position=*/true);
  } else if (is_top_level && !ctx.iterative_driver_suppressed() && ctx.has_formula_cell() &&
             ctx.current_sheet() != nullptr && ctx.workbook() != nullptr &&
             ctx.workbook()->iterative_options().enabled) {
    sv.value = run_iterative_fixed_point(node, arena, registry, ctx_with_counters);
  } else {
    sv = eval_root_shaped(node, arena, registry, ctx_with_counters);
  }
  const auto out_of_memory = [&]() {
    if (EvalState* state = ctx.state(); state != nullptr) {
      state->mark_out_of_memory();
    }
    return Value::error(ErrorCode::Num);
  };
  if (arena.exhausted()) {
    return out_of_memory();
  }

  Value v = Value::blank();
  if (first_element_only) {
    v = first_element(sv);
  } else {
    // A whole column or row spills at its declared size, though only its
    // populated head is stored. The footprint is decided before the array is
    // expanded: a refusal is recorded on the recalc path for the release
    // machinery to retry, as `evaluate_bare_range_spill` records one.
    std::uint32_t footprint_rows = 0;
    std::uint32_t footprint_cols = 0;
    bool has_footprint = false;
    if (sv.tail_array != nullptr) {
      footprint_rows = sv.tail_array->rows;
      footprint_cols = sv.tail_array->cols;
      has_footprint = true;
    }
    const bool spills = mode == TopResult::kSpill;
    if (has_footprint && !footprint_probed && spills && ctx.has_formula_cell() && ctx.current_sheet() != nullptr) {
      if (ctx.current_sheet()->probe_spill_footprint(ctx.formula_row(), ctx.formula_col(), footprint_rows,
                                                     footprint_cols) != Sheet::SpillAdmission::kAdmissible) {
        if (Sheet* target = ctx.mutable_sheet(); target == ctx.current_sheet()) {
          target->reject_spill_footprint(ctx.formula_row(), ctx.formula_col(), footprint_rows, footprint_cols);
        }
        return Value::error(ErrorCode::Spill);
      }
    }
    v = densify(sv, arena);
    if (arena.exhausted()) {
      return out_of_memory();
    }
    // Dynamic-array spill-collision surface contract. When a 365-era formula
    // produces a multi-cell array and is anchored at a known formula cell on
    // a resolvable sheet, Excel reports `#SPILL!` at the anchor if any cell
    // the result would occupy (other than the anchor) is already non-empty.
    // The mutable-sheet recalc path commits through `Sheet::commit_spill`,
    // which runs the authoritative collision scan (and clears stale phantom
    // regions first); this read-only check covers the path where no spill is
    // committed (ad-hoc evaluation, the oracle harness) so a blocked spill
    // still surfaces `#SPILL!` rather than the bare anchor scalar. A
    // footprint already measured above is not measured again.
    if (v.is_array() && !footprint_probed && spills && sv.tail_array == nullptr && ctx.mutable_sheet() == nullptr &&
        ctx.has_formula_cell() && ctx.current_sheet() != nullptr) {
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
  if (first_element_only) {
    return v;
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
      ArrayValue* promoted = allocate_derived_array_value(arr->rows, arr->cols, arena, buffer);
      if (promoted == nullptr) {
        return arena.exhausted() ? out_of_memory() : Value::error(ErrorCode::Num);
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

Value evaluate_top(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                   TopResult mode) {
  // An array past Excel's element bound abandons the whole formula, which
  // then holds 0 whatever encloses the array.
  const ScopedFormulaAbandon abandon_scope;
  const Value v = evaluate_top_body(node, arena, registry, ctx, mode);
  return abandon_scope.abandoned() ? Value::number(0.0) : v;
}

}  // namespace

Value evaluate(const parser::AstNode& node, Arena& arena) {
  return evaluate(node, arena, default_registry(), EvalContext{});
}

Value evaluate(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry) {
  return evaluate(node, arena, registry, EvalContext{});
}

Value evaluate(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx) {
  return evaluate_top(node, arena, registry, ctx, TopResult::kSpill);
}

Value evaluate_first_element(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                             const EvalContext& ctx) {
  return evaluate_top(node, arena, registry, ctx, TopResult::kFirstElement);
}

Value evaluate_array_formula(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                             const EvalContext& ctx) {
  return evaluate_top(node, arena, registry, ctx, TopResult::kWholeArray);
}

}  // namespace eval
}  // namespace formulon
