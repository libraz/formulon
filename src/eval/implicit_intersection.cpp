//
// Reference-level half of implicit intersection, shared by the `@` operator in
// the tree walker and `_xlfn.SINGLE`. Keeping one body means the two spellings
// cannot drift on which formula-cell positions are in range.

#include "eval/implicit_intersection.h"

#include <string_view>
#include <vector>

#include "eval/declared_rect.h"
#include "eval/dynamic_array/anchor.h"
#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/lambda_value.h"
#include "eval/lazy_impls.h"
#include "eval/name_env_resolve.h"
#include "eval/range_resolvers.h"
#include "eval/spill_anchor.h"
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

namespace {

// The rectangle `target` names when it yields a reference: a defined name,
// an intersection, a spill, a reference-returning call or a lambda call
// returning one. False when it yields none here.
bool ReferenceResultRect(const parser::AstNode& target, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx, std::string_view* sheet, std::uint32_t* top, std::uint32_t* left,
                         std::uint32_t* bottom, std::uint32_t* right) {
  ErrorCode err = ErrorCode::Value;
  switch (target.kind()) {
    case parser::NodeKind::NameRef:
    case parser::NodeKind::IntersectOp:
      return resolve_reference_rect(target, arena, registry, ctx, sheet, top, left, bottom, right, &err);
    case parser::NodeKind::ExternalRef:
      return parser::is_self_book_name_ref(target) &&
             resolve_reference_rect(target, arena, registry, ctx, sheet, top, left, bottom, right, &err);
    case parser::NodeKind::SpillRef: {
      if (!resolve_spill_anchor_node(target, arena, registry, ctx, sheet, top, left, &err)) {
        return false;
      }
      const ArrayValue* spill = project_spill_at_anchor(*sheet, *top, *left, arena, ctx, &err);
      if (spill == nullptr || spill->rows == 0U || spill->cols == 0U) {
        return false;
      }
      *bottom = *top + spill->rows - 1U;
      *right = *left + spill->cols - 1U;
      return true;
    }
    case parser::NodeKind::Call: {
      const std::string_view name = target.as_call_name();
      if (is_reference_call_name(name)) {
        return resolve_reference_rect(target, arena, registry, ctx, sheet, top, left, bottom, right, &err);
      }
      if (registry.lookup(name) != nullptr || find_lazy_impl(strip_future_prefix(name)) != nullptr) {
        return false;
      }
      // A call through a name bound to a LAMBDA.
      const std::uint32_t arity = target.as_call_arity();
      Value callee_err = Value::blank();
      const parser::AstNode* callee = parser::make_name_ref(arena, name);
      const LambdaValue* lv =
          callee != nullptr ? resolve_callable(*callee, arity, arena, registry, ctx, &callee_err) : nullptr;
      std::vector<const parser::AstNode*> args;
      for (std::uint32_t i = 0; i < arity; ++i) {
        args.push_back(&target.as_call_arg(i));
      }
      return lv != nullptr && resolve_lambda_reference(lv, arity, args.empty() ? nullptr : args.data(), arena, registry,
                                                       ctx, sheet, top, left, bottom, right, &err);
    }
    case parser::NodeKind::LambdaCall: {
      const Value callee = eval_node(target.as_lambda_call_callee(), arena, registry, ctx);
      if (!callee.is_lambda()) {
        return false;
      }
      std::vector<const parser::AstNode*> args;
      for (std::uint32_t i = 0; i < target.as_lambda_call_arity(); ++i) {
        args.push_back(&target.as_lambda_call_arg(i));
      }
      return resolve_lambda_reference(callee.as_lambda(), static_cast<std::uint32_t>(args.size()),
                                      args.empty() ? nullptr : args.data(), arena, registry, ctx, sheet, top, left,
                                      bottom, right, &err);
    }
    default:
      return false;
  }
}

}  // namespace

bool project_reference_result(const parser::AstNode& operand, Arena& arena, const FunctionRegistry& registry,
                              const EvalContext& ctx, Value* out) {
  const parser::AstNode& target = resolve_name_ast(operand, ctx.name_env());
  std::string_view sheet;
  std::uint32_t top = 0;
  std::uint32_t left = 0;
  std::uint32_t bottom = 0;
  std::uint32_t right = 0;
  if (!ReferenceResultRect(target, arena, registry, ctx, &sheet, &top, &left, &bottom, &right)) {
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
