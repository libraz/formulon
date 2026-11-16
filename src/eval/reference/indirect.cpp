//
// Implementation of the INDIRECT lazy impl. The textual A1 -> rectangle
// decoder lives in `reference/common.cpp`; this TU builds the reference.

#include "eval/reference/indirect.h"

#include "eval/declared_rect.h"
#include "eval/eval_context.h"
#include "eval/reference/common.h"
#include "eval/reference/ref_value.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {

Value eval_indirect_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx) {
  refs_internal::IndirectReference indirect{};
  ErrorCode err = ErrorCode::Value;
  if (!refs_internal::resolve_indirect_reference(call, arena, registry, ctx, &indirect, &err)) {
    return Value::error(err);
  }
  // The reference itself; reading it as values is the dereference's job, which records the read.
  // Whole-column (`A:A`) and whole-row (`1:1`) text names a whole-axis rectangle.
  const DeclaredRect rect{indirect.top_row, indirect.bottom_row, indirect.left_col, indirect.right_col,
                          /*whole_axis=*/indirect.is_full_col || indirect.is_full_row};
  return make_rect_ref(indirect.sheet, rect, RefOrigin::kDerived, arena, ctx).with_ref_dynamic();
}

}  // namespace eval
}  // namespace formulon
