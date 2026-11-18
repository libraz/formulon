//
// Reference-level half of implicit intersection, shared by the `@` operator in
// the tree walker and `_xlfn.SINGLE`. Keeping one body means the two spellings
// cannot drift on which formula-cell positions are in range.

#include "eval/implicit_intersection.h"

#include <optional>
#include <string_view>

#include "eval/declared_rect.h"
#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/lazy_impls.h"
#include "eval/range_resolvers.h"
#include "eval/reference/deref.h"
#include "eval/reference/ref_value.h"
#include "eval/tree_walker/dispatch.h"
#include "eval/tree_walker_lazy_table.h"

namespace formulon::eval {

std::optional<parser::Reference> project_implicit_intersection(const parser::Reference& lhs,
                                                               const parser::Reference& rhs, std::uint32_t formula_row,
                                                               std::uint32_t formula_col) {
  const Expected<DeclaredRect, ErrorCode> rect = declared_rect(lhs, rhs);
  if (!rect) {
    return std::nullopt;
  }
  const DeclaredRect& box = rect.value();

  parser::Reference target{};
  target.sheet = lhs.sheet;
  target.sheet_quoted = lhs.sheet_quoted;
  if (box.col_first == box.col_last && box.row_first == box.row_last) {
    // A one-cell rectangle is its own intersection: `=@D1:D1` reads D1 from any formula position.
    target.row = box.row_first;
    target.col = box.col_first;
    return target;
  }
  if (box.col_first == box.col_last) {
    // Single-column rectangle: project the formula row. A whole-column
    // reference lands here with the full grid height, so every formula row
    // is in span — which is why `=@C:C` reads the formula's own row instead
    // of degrading to `#VALUE!`.
    if (formula_row < box.row_first || formula_row > box.row_last) {
      return std::nullopt;
    }
    target.row = formula_row;
    target.col = box.col_first;
    return target;
  }
  if (box.row_first == box.row_last) {
    // Single-row rectangle: project the formula column.
    if (formula_col < box.col_first || formula_col > box.col_last) {
      return std::nullopt;
    }
    target.row = box.row_first;
    target.col = formula_col;
    return target;
  }
  // 2-D rectangle: Excel intersects on the formula cell's row AND column. When
  // the cell falls inside both spans the result is the single cell at
  // (formula_row, formula_col); any other position has no intersection.
  if (formula_row < box.row_first || formula_row > box.row_last || formula_col < box.col_first ||
      formula_col > box.col_last) {
    return std::nullopt;
  }
  target.row = formula_row;
  target.col = formula_col;
  return target;
}

IntersectionProjection project_implicit_intersection(const parser::AstNode& operand, std::uint32_t formula_row,
                                                     std::uint32_t formula_col, parser::Reference* out_target,
                                                     const Workbook* wb) {
  parser::Reference lhs{};
  parser::Reference rhs{};
  if (!declared_rect_endpoint_pair(operand, &lhs, &rhs, wb)) {
    // A `RangeOp` over reference-returning calls is still a range, but not
    // one with static coordinates; Excel rejects `@` on it rather than
    // reducing the evaluated rectangle.
    return operand.kind() == parser::NodeKind::RangeOp ? IntersectionProjection::kNoCell
                                                       : IntersectionProjection::kNotStaticReference;
  }
  // A bounded single `Ref` is already a scalar. Projecting it as a 1x1
  // rectangle would reject `=@A1` from every cell outside row 1.
  if (operand.kind() == parser::NodeKind::Ref && !lhs.is_full_col && !lhs.is_full_row) {
    return IntersectionProjection::kNotStaticReference;
  }
  const std::optional<parser::Reference> target = project_implicit_intersection(lhs, rhs, formula_row, formula_col);
  if (!target.has_value()) {
    return IntersectionProjection::kNoCell;
  }
  *out_target = *target;
  return IntersectionProjection::kCell;
}

namespace {

// Whether `operand` can yield a reference here: a name, an intersection, a spill, a reference-returning call, or a
// LET or lambda call returning one.
bool may_yield_reference(const parser::AstNode& operand, const FunctionRegistry& registry) {
  switch (operand.kind()) {
    case parser::NodeKind::NameRef:
    case parser::NodeKind::IntersectOp:
    case parser::NodeKind::SpillRef:
    case parser::NodeKind::LetBinding:
    case parser::NodeKind::LambdaCall:
      return true;
    case parser::NodeKind::Literal:
      return operand.as_literal().is_ref();
    case parser::NodeKind::ExternalRef:
      return parser::is_self_book_name_ref(operand);
    case parser::NodeKind::Call: {
      // A reference-returning built-in, or a call through a name bound to a LAMBDA.
      const std::string_view name = operand.as_call_name();
      return is_reference_call_name(name) || (lookup_on_current_host(registry, name) == nullptr &&
                                              find_lazy_impl(strip_future_prefix(name)) == nullptr);
    }
    default:
      return false;
  }
}

}  // namespace

bool project_reference_result(const parser::AstNode& operand, Arena& arena, const FunctionRegistry& registry,
                              const EvalContext& ctx, Value* out) {
  // A call keeping a grid of references (`@SORT(OFFSET(A1,{0;3},{3;1}))`, `@+OFFSET(A1,{0;2},0)`) is projected in a
  // dynamic-array cell.
  const bool keeps_grid = !ctx.legacy_formula_cell() &&
                          (operand.kind() == parser::NodeKind::Call || operand.kind() == parser::NodeKind::UnaryOp) &&
                          !may_yield_reference(operand, registry) && evaluates_to_reference(operand);
  if (!keeps_grid && !may_yield_reference(operand, registry)) {
    return false;
  }
  Value ref = eval_node_ref(operand, arena, registry, ctx);
  if (!ref.is_ref()) {
    if (keeps_grid) {
      *out = implicit_intersect_value(ref);
      return true;
    }
    return false;
  }
  // A grid gives its first element. A dynamic-array cell reads that reference whole (`@OFFSET(A1:A3,,{1,2,3})` is
  // B1:B3); a legacy cell intersects it with the formula cell like any area.
  if (ref.as_ref().form == RefForm::kGrid) {
    const Value first = grid_element(ref.as_ref(), 0U, arena);
    if (!first.is_ref() || !ctx.legacy_formula_cell()) {
      *out = first.is_ref() ? deref_shaped(first, arena, registry, ctx, ReadCount::kCounted).value : first;
      return true;
    }
    ref = first;
  }
  const RefView view = ref.as_ref();
  const RefArea& area = view.areas[0];
  parser::Reference lhs{};
  if (view.form != RefForm::kArea || view.book != 0U || !ref_sheet_name(ctx, area.sheet_first, &lhs.sheet)) {
    return false;
  }
  lhs.row = area.row_first;
  lhs.col = area.col_first;
  parser::Reference rhs = lhs;
  rhs.row = area.row_last;
  rhs.col = area.col_last;
  const std::optional<parser::Reference> cell =
      area.row_first == area.row_last && area.col_first == area.col_last
          ? std::optional<parser::Reference>(lhs)
          : project_implicit_intersection(lhs, rhs, ctx.formula_row(), ctx.formula_col());
  if (!cell.has_value()) {
    *out = Value::error(ErrorCode::Value);
    return true;
  }
  // A dynamic reference reads only the cell it projects to.
  if (ref.ref_is_dynamic()) {
    ctx.note_dynamic_read(cell->sheet, DeclaredRect{cell->row, cell->row, cell->col, cell->col, /*whole_axis=*/false});
  }
  *out = ctx.resolve_ref(*cell, arena, registry);
  return true;
}

}  // namespace formulon::eval
