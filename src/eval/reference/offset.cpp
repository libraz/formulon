//
// Implementation of the OFFSET lazy impl.
//
// The rectangle-construction core (`compute_offset_rect`, `OffsetBase`)
// lives in `reference/common.cpp` because the intersection resolver
// (`reference/intersection.cpp`) reaches it too.

#include "eval/reference/offset.h"

#include <cstdint>

#include "eval/declared_rect.h"
#include "eval/eval_context.h"
#include "eval/reference/common.h"
#include "eval/reference/ref_value.h"
#include "parser/ast.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {

namespace {

// The rectangle an OFFSET call names: the base's sheet plus the 0-based
// top-left corner and the extent.
struct OffsetRect {
  refs_internal::OffsetBase base{};
  std::uint32_t top_row = 0;
  std::uint32_t left_col = 0;
  std::uint32_t height = 0;
  std::uint32_t width = 0;
};

bool resolve_offset_rect(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx, OffsetRect* out, ErrorCode* out_err) {
  return refs_internal::compute_offset_rect(call, arena, registry, ctx, &out->base, &out->top_row, &out->left_col,
                                            &out->height, &out->width, out_err);
}

// The rectangle `rect` names; a span of a whole grid axis is a whole-axis rectangle.
DeclaredRect declared_offset_rect(const OffsetRect& rect) {
  return DeclaredRect{rect.top_row, rect.top_row + rect.height - 1U, rect.left_col, rect.left_col + rect.width - 1U,
                      /*whole_axis=*/rect.height == Sheet::kMaxRows || rect.width == Sheet::kMaxCols};
}

}  // namespace

Value eval_offset_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  OffsetRect rect{};
  ErrorCode err = ErrorCode::Value;
  if (!resolve_offset_rect(call, arena, registry, ctx, &rect, &err)) {
    return Value::error(err);
  }
  // The reference itself; reading it as values is the dereference's job, which records the read.
  return make_rect_ref(rect.base.sheet, declared_offset_rect(rect), RefOrigin::kDerived, arena, ctx).with_ref_dynamic();
}

}  // namespace eval
}  // namespace formulon
