
#include "eval/dynamic_array/indexing.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <string_view>
#include <vector>

#include "eval/builtin_names.h"
#include "eval/coerce.h"
#include "eval/dynamic_array/common.h"
#include "eval/eval_context.h"
#include "eval/name_env_resolve.h"
#include "eval/omitted_arg.h"
#include "eval/range_args.h"
#include "eval/range_resolvers.h"
#include "eval/tree_walker/dispatch.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/strings.h"
#include "value.h"

namespace formulon {
namespace eval {

namespace {

/// Maps one CHOOSECOLS / CHOOSEROWS index to a 0-based lane. The value is
/// truncated toward zero; positives address `[1, axis_size]`, negatives count
/// from the end, and zero or anything out of range is `#VALUE!`.
bool map_choose_index(const Value& cell, std::uint32_t axis_size, std::uint32_t& out, Value& error_out) {
  if (cell.is_error()) {
    error_out = cell;
    return false;
  }
  auto number = coerce_to_index_number(cell);
  if (!number) {
    error_out = Value::error(number.error());
    return false;
  }
  const double truncated = number.value();
  const double magnitude = std::fabs(truncated);
  if (truncated == 0.0 || magnitude > static_cast<double>(axis_size)) {
    error_out = Value::error(ErrorCode::Value);
    return false;
  }
  const auto lane = static_cast<std::uint32_t>(magnitude);
  out = truncated > 0.0 ? lane - 1U : axis_size - lane;
  return true;
}

/// CHOOSECOLS / CHOOSEROWS. The two differ only in which axis the indices
/// address, so `by_col` selects both the bound the indices are checked against
/// and the lane the result is assembled from.
Value eval_choose_lanes(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx, bool by_col) {
  const std::uint32_t arity = call.as_call_arity();
  // Need `(array, index1, ...)`; cap at 254 indices (Excel ceiling) +
  // 1 array slot = 255 args.
  Value err = Value::error(ErrorCode::Value);
  const ArrayValue* array = dynamic_array::resolve_array_prelude(call, 2U, kMaxCallArguments, /*reject_empty=*/false,
                                                                 arena, registry, ctx, err);
  if (array == nullptr) {
    return err;
  }

  // Each index argument is a scalar or a 1-D array; its elements are picked in
  // order. A 2-D index array is `#VALUE!`.
  const std::uint32_t lanes = by_col ? array->cols : array->rows;
  std::vector<std::uint32_t> picks;
  picks.reserve(arity - 1U);
  for (std::uint32_t i = 1; i < arity; ++i) {
    const ArrayValue* indices = nullptr;
    if (!resolve_array_value(call.as_call_arg(i), arena, registry, ctx, &indices, &err)) {
      return err;
    }
    if (indices->rows != 1U && indices->cols != 1U) {
      return Value::error(ErrorCode::Value);
    }
    const std::size_t count = static_cast<std::size_t>(indices->rows) * indices->cols;
    for (std::size_t k = 0; k < count; ++k) {
      std::uint32_t idx = 0;
      if (!map_choose_index(indices->cells[k], lanes, idx, err)) {
        return err;
      }
      picks.push_back(idx);
    }
  }

  ArrayValue* out = dynamic_array::materialise_selected_lanes(*array, picks, by_col, arena);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  return Value::array(out);
}

/// TAKE / DROP. Both read `(array, [rows], [cols])` and slice the result; the
/// only difference is whether a count names the span to keep or the span to
/// remove, which `resolve_take_drop_range` decides from `take`.
// Whether `node` denotes a reference, so TAKE / DROP over it yield a reference.
bool denotes_reference(const parser::AstNode& node, const EvalContext& ctx) {
  const parser::AstNode& source = resolve_range_binding(node, ctx.name_env(), /*accept_ref=*/true);
  switch (source.kind()) {
    case parser::NodeKind::Ref:
    case parser::NodeKind::RangeOp:
    case parser::NodeKind::SpillRef:
    case parser::NodeKind::StructuredRef:
      return true;
    default:
      return false;
  }
}

Value eval_take_drop(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx, bool take) {
  Value err = Value::error(ErrorCode::Value);
  const ArrayValue* array =
      dynamic_array::resolve_array_prelude(call, 2U, 3U, /*reject_empty=*/false, arena, registry, ctx, err);
  if (array == nullptr) {
    return err;
  }

  std::uint32_t row_lo = 0;
  std::uint32_t row_hi = 0;
  const parser::AstNode* row_arg = is_omitted_arg(call.as_call_arg(1)) ? nullptr : &call.as_call_arg(1);
  if (!dynamic_array::resolve_take_drop_range(row_arg, array->rows, take, arena, registry, ctx, row_lo, row_hi, err)) {
    return err;
  }
  std::uint32_t col_lo = 0;
  std::uint32_t col_hi = array->cols;
  const parser::AstNode* col_arg = call_arg_supplied(call, 2U) ? &call.as_call_arg(2) : nullptr;
  if (!dynamic_array::resolve_take_drop_range(col_arg, array->cols, take, arena, registry, ctx, col_lo, col_hi, err)) {
    return err;
  }

  ArrayValue* out = dynamic_array::materialise_slice(*array, row_lo, row_hi, col_lo, col_hi, arena);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  // A one-cell slice of a reference is a one-cell reference, which reads as that cell's value.
  if (out->rows == 1U && out->cols == 1U && denotes_reference(call.as_call_arg(0), ctx)) {
    return out->cells[0];
  }
  return Value::array(out);
}

}  // namespace

bool resolve_take_drop_reference(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                                 const EvalContext& ctx, std::string_view* out_sheet, std::uint32_t* out_top_row,
                                 std::uint32_t* out_left_col, std::uint32_t* out_bottom_row,
                                 std::uint32_t* out_right_col, bool* out_is_range, ErrorCode* out_err) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 2U || arity > 3U) {
    *out_err = ErrorCode::Value;
    return false;
  }
  const bool take = strings::case_insensitive_eq(strip_future_prefix(call.as_call_name()), "TAKE");
  std::string_view sheet;
  std::uint32_t top = 0;
  std::uint32_t left = 0;
  std::uint32_t bottom = 0;
  std::uint32_t right = 0;
  if (!resolve_reference_rect(call.as_call_arg(0), arena, registry, ctx, &sheet, &top, &left, &bottom, &right,
                              out_err)) {
    return false;
  }
  Value err = Value::error(ErrorCode::Value);
  std::uint32_t row_lo = 0;
  std::uint32_t row_hi = 0;
  const parser::AstNode* row_arg = is_omitted_arg(call.as_call_arg(1)) ? nullptr : &call.as_call_arg(1);
  std::uint32_t col_lo = 0;
  std::uint32_t col_hi = right - left + 1U;
  const parser::AstNode* col_arg = call_arg_supplied(call, 2U) ? &call.as_call_arg(2) : nullptr;
  if (!dynamic_array::resolve_take_drop_range(row_arg, bottom - top + 1U, take, arena, registry, ctx, row_lo, row_hi,
                                              err) ||
      !dynamic_array::resolve_take_drop_range(col_arg, right - left + 1U, take, arena, registry, ctx, col_lo, col_hi,
                                              err)) {
    *out_err = err.is_error() ? err.as_error() : ErrorCode::Value;
    return false;
  }
  if (row_lo >= row_hi || col_lo >= col_hi) {
    *out_err = ErrorCode::Calc;
    return false;
  }
  *out_sheet = sheet;
  *out_top_row = top + row_lo;
  *out_bottom_row = top + row_hi - 1U;
  *out_left_col = left + col_lo;
  *out_right_col = left + col_hi - 1U;
  *out_is_range = *out_top_row != *out_bottom_row || *out_left_col != *out_right_col;
  return true;
}

Value eval_choosecols_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx) {
  return eval_choose_lanes(call, arena, registry, ctx, /*by_col=*/true);
}

Value eval_chooserows_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx) {
  return eval_choose_lanes(call, arena, registry, ctx, /*by_col=*/false);
}

Value eval_take_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  return eval_take_drop(call, arena, registry, ctx, /*take=*/true);
}

Value eval_drop_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  return eval_take_drop(call, arena, registry, ctx, /*take=*/false);
}

}  // namespace eval
}  // namespace formulon
