
#include "eval/dynamic_array/anchor.h"

#include <cstddef>
#include <cstdint>
#include <string_view>

#include "eval/eval_context.h"
#include "eval/range_resolvers.h"
#include "parser/ast.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {

Value eval_anchorarray_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                            const EvalContext& ctx) {
  if (call.as_call_arity() != 1U) {
    return Value::error(ErrorCode::Value);
  }
  const parser::AstNode& arg = call.as_call_arg(0);

  // Resolve the argument AST to (sheet_qualifier, top-left cell). The spill
  // anchor is always a single cell, so we ignore any range extent that
  // `resolve_reference_call` may report for OFFSET / INDIRECT — Excel treats
  // `_xlfn.ANCHORARRAY(B2:D5)` the same as `ANCHORARRAY(B2)`: the anchor
  // is whichever cell sits at the top-left.
  std::string_view sheet_name;
  std::uint32_t anchor_row = 0;
  std::uint32_t anchor_col = 0;
  switch (arg.kind()) {
    case parser::NodeKind::Ref: {
      const parser::Reference& r = arg.as_ref();
      if (r.is_full_col || r.is_full_row) {
        return Value::error(ErrorCode::Ref);
      }
      sheet_name = r.sheet;
      anchor_row = r.row;
      anchor_col = r.col;
      break;
    }
    case parser::NodeKind::SpillRef: {
      const parser::Reference& r = arg.as_spill_ref();
      sheet_name = r.sheet;
      anchor_row = r.row;
      anchor_col = r.col;
      break;
    }
    case parser::NodeKind::Call: {
      std::uint32_t bottom_row = 0;
      std::uint32_t right_col = 0;
      bool is_range = false;
      ErrorCode err = ErrorCode::Ref;
      if (!resolve_reference_call(arg, arena, registry, ctx, &sheet_name, &anchor_row, &anchor_col, &bottom_row,
                                  &right_col, &is_range, &err)) {
        return Value::error(err);
      }
      break;
    }
    default:
      return Value::error(ErrorCode::Value);
  }

  const Sheet* current = ctx.current_sheet();
  if (current == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  const Sheet* target = current;
  if (!sheet_name.empty()) {
    const Workbook* wb = ctx.workbook();
    if (wb == nullptr) {
      return Value::error(ErrorCode::Ref);
    }
    target = wb->sheet_by_name(sheet_name);
    if (target == nullptr) {
      return Value::error(ErrorCode::Ref);
    }
  }
  if (anchor_row >= Sheet::kMaxRows || anchor_col >= Sheet::kMaxCols) {
    return Value::error(ErrorCode::Ref);
  }
  const SpillRegion* region = target->spill_region_at_anchor(anchor_row, anchor_col);
  if (region == nullptr) {
    return Value::error(ErrorCode::Ref);
  }
  const std::size_t n = static_cast<std::size_t>(region->rows) * static_cast<std::size_t>(region->cols);
  Value* buffer = arena.create_array<Value>(n);
  if (buffer == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  for (std::size_t i = 0; i < n; ++i) {
    buffer[i] = region->cells[i];
  }
  ArrayValue* arr = arena.create<ArrayValue>();
  if (arr == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  arr->rows = region->rows;
  arr->cols = region->cols;
  arr->cells = buffer;
  return Value::array(arr);
}

}  // namespace eval
}  // namespace formulon
