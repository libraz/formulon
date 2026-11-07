//
// Spill positions of the tree-walk evaluator: a bare range or an
// intersection standing as the entire formula, read at its declared size and
// weighed against the anchor cell's spill footprint before it is built.

#include <cstddef>
#include <cstdint>
#include <string_view>
#include <vector>

#include "eval/array_alloc.h"
#include "eval/declared_rect.h"
#include "eval/dep_graph.h"
#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/lazy_impls.h"
#include "eval/range_resolvers.h"
#include "eval/recalc_engine.h"
#include "eval/spill_anchor.h"
#include "eval/tail_array.h"
#include "eval/tree_walker/dispatch.h"
#include "eval/tree_walker/walker_internal.h"
#include "parser/ast.h"
#include "sheet.h"
#include "sheet_name.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {

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

namespace {

// The corner `Reference`s of `rect`. The first carries `lhs`'s sheet
// qualifier, as `Sheet1!A1:B2` parses; the last does too when `qualify_last`
// is set, and otherwise inherits it in `expand_range`.
void rect_corner_refs(const DeclaredRect& rect, const parser::Reference& lhs, bool qualify_last,
                      parser::Reference* top_left, parser::Reference* bottom_right) {
  parser::Reference first{};
  parser::Reference last{};
  first.sheet = lhs.sheet;
  first.sheet_quoted = lhs.sheet_quoted;
  first.row = rect.row_first;
  first.col = rect.col_first;
  if (qualify_last) {
    last.sheet = lhs.sheet;
    last.sheet_quoted = lhs.sheet_quoted;
  }
  last.row = rect.row_last;
  last.col = rect.col_last;
  *top_left = first;
  *bottom_right = last;
}

}  // namespace

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
  // The parser keeps the sheet qualifier on the left endpoint, matching how
  // `Sheet1!A1:B2` parses; `expand_range` inherits it for the rectangle.
  rect_corner_refs(rect.value(), lhs, /*qualify_last=*/false, top_left, bottom_right);
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
  // `eval_node` carries the left endpoint's qualifier onto both corners;
  // reproducing it keeps the two entry points describing one rectangle.
  rect_corner_refs(rect.value(), lhs, /*qualify_last=*/true, top_left, bottom_right);
  return true;
}

// Reads the rectangle `[lhs : rhs]`, which spans a whole grid axis, at its
// declared size: the populated head as dense cells (the same walk
// `expand_range` makes) and one repeated blank row / column for everything
// past it. The endpoints may be whole-column / whole-row `Ref`s or bounded
// corners that cover the full axis. A rectangle whole on both axes has no
// compressed form and exceeds the range-expansion ceiling.
Shaped read_whole_axis(const parser::Reference& lhs, const parser::Reference& rhs, Arena& arena,
                       const FunctionRegistry& registry, const EvalContext& ctx) {
  const Expected<DeclaredRect, ErrorCode> declared = ctx.declared_range_rect(lhs, rhs);
  if (!declared) {
    return Shaped{Value::error(declared.error()), nullptr};
  }
  const DeclaredRect& rect = declared.value();
  const bool full_height = rect.rows() == Sheet::kMaxRows;
  parser::Reference lo{};
  parser::Reference hi{};
  lo.sheet = lhs.sheet;
  lo.sheet_quoted = lhs.sheet_quoted;
  hi.sheet = rhs.sheet;
  hi.sheet_quoted = rhs.sheet_quoted;
  lo.row = rect.row_first;
  hi.row = rect.row_last;
  lo.col = rect.col_first;
  hi.col = rect.col_last;
  // `expand_range` narrows only a reference flagged whole-axis.
  (full_height ? lo.is_full_col : lo.is_full_row) = true;
  (full_height ? hi.is_full_col : hi.is_full_row) = true;
  std::uint32_t walked_rows = 0;
  std::uint32_t walked_cols = 0;
  auto expanded = ctx.expand_range(lo, hi, arena, registry, &walked_rows, &walked_cols);
  if (!expanded) {
    return Shaped{Value::error(expanded.error()), nullptr};
  }
  const std::vector<Value>& walked = expanded.value();
  Value* cells = walked.empty() ? nullptr : arena.create_array<Value>(walked.size());
  if (!walked.empty() && cells == nullptr) {
    return Shaped{Value::error(ErrorCode::Num), nullptr};
  }
  for (std::size_t i = 0; i < walked.size(); ++i) {
    cells[i] = walked[i];
  }
  return make_reference_tail(arena, rect.rows(), rect.cols(), full_height ? walked_rows : walked_cols,
                             full_height ? TailAxis::kRows : TailAxis::kCols, cells);
}

namespace {

// Builds the value of the bounded-corner rectangle a bare range declares: at
// its declared size when it spans a whole grid axis (and `allow_tail`),
// otherwise materialised.
Shaped build_declared_rect(const parser::Reference& top_left, const parser::Reference& bottom_right, Arena& arena,
                           const FunctionRegistry& registry, const EvalContext& ctx, bool allow_tail) {
  if (allow_tail && is_full_axis_range(top_left, bottom_right)) {
    return read_whole_axis(top_left, bottom_right, arena, registry, ctx);
  }
  return Shaped{materialize_rectangle(top_left, bottom_right, arena, registry, ctx), nullptr};
}

}  // namespace

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
// graph and the engine's cycle policy answers on its own, after building
// the rectangle. Pre-empting reaches the same answer without building it —
// the same shape as the footprint pre-check above. The read-only driver has
// no graph to fall back on, which is where the two spellings visibly
// disagreed.
//
// The answer is the engine's cycle policy. When no cell inside the
// rectangle reads the formula cell back, the formula cell is the only member
// of its cycle, and Excel leaves it at 0; otherwise the cycle runs through
// several cells, which Formulon reports as `#REF!`.
//
// Without a formula cell to anchor against there is no footprint to
// measure and nothing that could spill, so the rectangle is built
// directly; that is the shape ad-hoc parser-level evaluation sees.
//
// A rectangle spanning a whole grid axis is built as a `TailArray` when
// `allow_tail`, so only the populated head is read; the caller expands it once
// the footprint is admitted. `*footprint_probed` is set when the footprint was
// measured here, so the caller does not measure it again.
Shaped evaluate_bare_range_spill(const parser::Reference& top_left, const parser::Reference& bottom_right, Arena& arena,
                                 const FunctionRegistry& registry, const EvalContext& ctx,
                                 const BareRangeSpillOptions& options, bool* footprint_probed) {
  const bool settle_circularity = options.settle_circularity;
  const bool allow_tail = options.allow_tail;
  const Sheet* sheet = ctx.current_sheet();
  if (sheet == nullptr || !ctx.has_formula_cell()) {
    return build_declared_rect(top_left, bottom_right, arena, registry, ctx, allow_tail);
  }
  const std::uint32_t anchor_row = ctx.formula_row();
  const std::uint32_t anchor_col = ctx.formula_col();

  // The rectangle is on the formula's own sheet when it carries no
  // qualifier, or when the qualifier names that sheet.
  const bool same_sheet = top_left.sheet.empty() || sheet_names::equal(top_left.sheet, sheet->name());
  if (settle_circularity && same_sheet && anchor_row >= top_left.row && anchor_row <= bottom_right.row &&
      anchor_col >= top_left.col && anchor_col <= bottom_right.col) {
    // Circular, and reported as the engine reports every other cycle
    // rather than as anything specific to this shape: with iterative
    // calculation off, 0 for a cell that is its own only cycle member and
    // `#REF!` otherwise; the cell's last computed value when it is on. That
    // mirrors `EvalContext::resolve_ref`'s back-edge branch and the value
    // `RecalcEngine` writes for each member of a cyclic component. It
    // cannot be delegated to `resolve_ref` here, because the
    // dependency-ordered recalc path evaluates with no `EvalState` and
    // short-circuits formula references to their cached values — asking it
    // would yield the anchor's stale value instead of a cycle.
    const Workbook* workbook = ctx.workbook();
    if (workbook != nullptr && workbook->iterative_options().enabled) {
      return Shaped{sheet->resolve_cell_value(anchor_row, anchor_col), nullptr};
    }
    // Another cell joins the cycle only if it reads the anchor back; a
    // formula that merely sits in the rectangle does not.
    const std::size_t sheet_index = workbook == nullptr ? 0U : workbook->sheet_index_by_name(sheet->name());
    const bool sole_member = workbook == nullptr || sheet_index >= workbook->sheet_count() ||
                             !workbook->recalc_engine().dep_graph().rect_cell_depends_on(
                                 CellNodeId{static_cast<std::uint16_t>(sheet_index), anchor_row, anchor_col},
                                 top_left.row, top_left.col, bottom_right.row, bottom_right.col);
    if (sole_member) {
      abandon_formula();
    }
    return Shaped{Value::error(ErrorCode::Ref), nullptr};
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
    return Shaped{Value::error(ErrorCode::Spill), nullptr};
  }
  if (footprint_probed != nullptr) {
    *footprint_probed = true;
  }
  return build_declared_rect(top_left, bottom_right, arena, registry, ctx, allow_tail);
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
  parser::Reference bottom_right{};
  set_rect_corner_refs(sheet, r1, c1, r2, c2, &top_left, &bottom_right);
  const bool full_height = spans_all_rows(r1, r2);
  const bool full_width = spans_all_cols(c1, c2);
  if ((r1 == r2 && c1 == c2) || (!spill_position && (full_height || full_width))) {
    return ctx.resolve_ref(top_left, arena, registry);
  }
  if (spill_position) {
    return evaluate_bare_range_spill(top_left, bottom_right, arena, registry, ctx,
                                     BareRangeSpillOptions{/*settle_circularity=*/false, /*allow_tail=*/false},
                                     /*footprint_probed=*/nullptr)
        .value;
  }
  return materialize_rectangle(top_left, bottom_right, arena, registry, ctx);
}

}  // namespace eval
}  // namespace formulon
