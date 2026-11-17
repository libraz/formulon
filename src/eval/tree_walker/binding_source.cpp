//
// LET / LAMBDA binding-source evaluation, and the static AST of a reference
// source for readers that dispatch on the AST. The public contract is
// declared in `tree_walker/dispatch.h`.

#include <cstddef>
#include <cstdint>
#include <string_view>

#include "eval/declared_rect.h"
#include "eval/eval_context.h"
#include "eval/external_ref.h"
#include "eval/lazy_impls.h"
#include "eval/range_resolvers.h"
#include "eval/reference/ref_value.h"
#include "eval/structured_ref_project.h"
#include "eval/tree_walker/dispatch.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

// A cell or range of a supporting workbook as a `Ref` into that book's cache; null when `expr` is no such reference
// or its book is no external link of this workbook.
const Value* external_reference(const parser::AstNode& expr, Arena& arena, const EvalContext& ctx, Value* out) {
  const parser::Reference* lhs = nullptr;
  const parser::Reference* rhs = nullptr;
  const Workbook* wb = ctx.workbook();
  if (wb == nullptr || !external_ref_declared_endpoints(expr, &lhs, &rhs)) {
    return nullptr;
  }
  const Expected<ExternalRect, ErrorCode> rect = resolve_external_rect(expr, ctx);
  if (!rect) {
    *out = Value::error(rect.error());
    return out;
  }
  const auto& links = wb->external_links();
  for (std::size_t i = 0; i < links.size(); ++i) {
    if (&links[i].book == rect.value().book) {
      const auto sheet = static_cast<std::uint16_t>(rect.value().sheet);
      const RefArea area = make_ref_area(rect.value().declared, sheet, sheet, RefOrigin::kSpelled);
      *out = make_ref(arena, RefForm::kArea, static_cast<std::uint16_t>(i + 1U), &area, 1U);
      return out;
    }
  }
  return nullptr;
}

}  // namespace

Value eval_binding_source(const parser::AstNode& expr, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx) {
  if (Value external = Value::blank(); external_reference(expr, arena, ctx, &external) != nullptr) {
    return external;
  }
  return eval_node_ref(expr, arena, registry, ctx);
}

const parser::AstNode* resolve_binding_reference(const parser::AstNode& expr, Arena& arena,
                                                 const FunctionRegistry& registry, const EvalContext& ctx) {
  parser::Reference first{};
  parser::Reference last{};
  if (expr.kind() == parser::NodeKind::Ref || reference_endpoint_pair(expr, &first, &last, ctx)) {
    return &expr;
  }
  if (expr.kind() == parser::NodeKind::ExternalRef) {
    // The defined-name form (`[Book.xlsx]!Name`, `[0]!Name`) is not a cell.
    return expr.as_external_ref_name().empty() ? &expr : nullptr;
  }
  if (!evaluates_to_reference(expr)) {
    return nullptr;
  }
  std::string_view sheet;
  RefArea area{};
  ErrorCode err = ErrorCode::Value;
  if (!eval_reference_area(expr, arena, registry, ctx, &sheet, &area, &err)) {
    return nullptr;
  }
  first = parser::Reference{};
  last = parser::Reference{};
  first.sheet = sheet;
  last.sheet = sheet;
  if (area.whole_col != 0U) {
    first.is_full_col = true;
    last.is_full_col = true;
    first.col = area.col_first;
    last.col = area.col_last;
  } else if (area.whole_row != 0U) {
    first.is_full_row = true;
    last.is_full_row = true;
    first.row = area.row_first;
    last.row = area.row_last;
  } else {
    first.row = area.row_first;
    first.col = area.col_first;
    last.row = area.row_last;
    last.col = area.col_last;
  }
  // Always a `RangeOp`, even for one cell (`D1:D1`): a bare `Ref` is the
  // scalar-provenance shape, and a reference-returning call is not.
  parser::AstNode* lhs = parser::make_ref(arena, first);
  parser::AstNode* rhs = parser::make_ref(arena, last);
  if (lhs == nullptr || rhs == nullptr) {
    return nullptr;
  }
  // A whole-axis rectangle is one `A:C` token.
  return parser::make_range_op(arena, lhs, rhs, /*spliced=*/first.is_full_col || first.is_full_row);
}

}  // namespace eval
}  // namespace formulon
