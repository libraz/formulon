//
// Reference-level half of implicit intersection, shared by the `@` operator in
// the tree walker and `_xlfn.SINGLE`. Keeping one body means the two spellings
// cannot drift on which formula-cell positions are in range.

#include "eval/implicit_intersection.h"

#include <string_view>

#include "eval/declared_rect.h"
#include "eval/eval_context.h"
#include "eval/name_env_resolve.h"
#include "eval/range_resolvers.h"

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
                                                     std::uint32_t formula_col, parser::Reference* out_target) {
  parser::Reference lhs{};
  parser::Reference rhs{};
  if (!declared_rect_endpoint_pair(operand, &lhs, &rhs)) {
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

bool project_reference_result(const parser::AstNode& operand, Arena& arena, const FunctionRegistry& registry,
                              const EvalContext& ctx, Value* out) {
  const parser::AstNode& target = resolve_name_ast(operand, ctx.name_env());
  const bool yields_reference =
      target.kind() == parser::NodeKind::NameRef || target.kind() == parser::NodeKind::IntersectOp ||
      (target.kind() == parser::NodeKind::ExternalRef && parser::is_self_book_name_ref(target)) ||
      (target.kind() == parser::NodeKind::Call && is_reference_call_name(target.as_call_name()));
  if (!yields_reference) {
    return false;
  }
  std::string_view sheet;
  std::uint32_t top = 0;
  std::uint32_t left = 0;
  std::uint32_t bottom = 0;
  std::uint32_t right = 0;
  ErrorCode err = ErrorCode::Value;
  if (!resolve_reference_rect(target, arena, registry, ctx, &sheet, &top, &left, &bottom, &right, &err)) {
    return false;
  }
  parser::Reference lhs{};
  lhs.sheet = sheet;
  lhs.row = top;
  lhs.col = left;
  parser::Reference rhs = lhs;
  rhs.row = bottom;
  rhs.col = right;
  if (top == bottom && left == right) {
    *out = ctx.resolve_ref(lhs, arena, registry);
    return true;
  }
  const std::optional<parser::Reference> cell =
      project_implicit_intersection(lhs, rhs, ctx.formula_row(), ctx.formula_col());
  *out = cell.has_value() ? ctx.resolve_ref(*cell, arena, registry) : Value::error(ErrorCode::Value);
  return true;
}

}  // namespace formulon::eval
