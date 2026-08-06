
#include "eval/dynamic_array/indexing.h"

#include <cstdint>
#include <vector>

#include "eval/dynamic_array/common.h"
#include "eval/range_args.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {

Value eval_choosecols_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  // Need `(array, col_num1, ...)`; cap at 254 indices (Excel ceiling) +
  // 1 array slot = 255 args.
  if (arity < 2U || arity > 255U) {
    return Value::error(ErrorCode::Value);
  }

  const ArrayValue* array = nullptr;
  Value err = Value::error(ErrorCode::Value);
  if (!resolve_array_value(call.as_call_arg(0), arena, registry, ctx, &array, &err)) {
    return err;
  }

  // Resolve each index in turn. `resolve_choose_index` handles coercion,
  // truncation, sign mapping, and bounds in one place.
  std::vector<std::uint32_t> picks;
  picks.reserve(arity - 1U);
  for (std::uint32_t i = 1; i < arity; ++i) {
    std::uint32_t idx = 0;
    Value idx_err = Value::error(ErrorCode::Value);
    if (!dynamic_array::resolve_choose_index(call.as_call_arg(i), array->cols, arena, registry, ctx, idx, idx_err)) {
      return idx_err;
    }
    picks.push_back(idx);
  }

  ArrayValue* out = dynamic_array::materialise_selected_lanes(*array, picks, /*by_col=*/true, arena);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  return Value::array(out);
}

Value eval_chooserows_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 2U || arity > 255U) {
    return Value::error(ErrorCode::Value);
  }

  const ArrayValue* array = nullptr;
  Value err = Value::error(ErrorCode::Value);
  if (!resolve_array_value(call.as_call_arg(0), arena, registry, ctx, &array, &err)) {
    return err;
  }

  std::vector<std::uint32_t> picks;
  picks.reserve(arity - 1U);
  for (std::uint32_t i = 1; i < arity; ++i) {
    std::uint32_t idx = 0;
    Value idx_err = Value::error(ErrorCode::Value);
    if (!dynamic_array::resolve_choose_index(call.as_call_arg(i), array->rows, arena, registry, ctx, idx, idx_err)) {
      return idx_err;
    }
    picks.push_back(idx);
  }

  ArrayValue* out = dynamic_array::materialise_selected_lanes(*array, picks, /*by_col=*/false, arena);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  return Value::array(out);
}

Value eval_take_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 2U || arity > 3U) {
    return Value::error(ErrorCode::Value);
  }

  const ArrayValue* array = nullptr;
  Value err = Value::error(ErrorCode::Value);
  if (!resolve_array_value(call.as_call_arg(0), arena, registry, ctx, &array, &err)) {
    return err;
  }

  std::uint32_t row_lo = 0;
  std::uint32_t row_hi = 0;
  if (!dynamic_array::resolve_take_drop_range(&call.as_call_arg(1), array->rows, /*take=*/true, arena, registry, ctx,
                                              row_lo, row_hi, err)) {
    return err;
  }
  std::uint32_t col_lo = 0;
  std::uint32_t col_hi = array->cols;
  if (arity == 3U) {
    if (!dynamic_array::resolve_take_drop_range(&call.as_call_arg(2), array->cols, /*take=*/true, arena, registry, ctx,
                                                col_lo, col_hi, err)) {
      return err;
    }
  }

  ArrayValue* out = dynamic_array::materialise_slice(*array, row_lo, row_hi, col_lo, col_hi, arena);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  return Value::array(out);
}

Value eval_drop_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 2U || arity > 3U) {
    return Value::error(ErrorCode::Value);
  }

  const ArrayValue* array = nullptr;
  Value err = Value::error(ErrorCode::Value);
  if (!resolve_array_value(call.as_call_arg(0), arena, registry, ctx, &array, &err)) {
    return err;
  }

  std::uint32_t row_lo = 0;
  std::uint32_t row_hi = 0;
  if (!dynamic_array::resolve_take_drop_range(&call.as_call_arg(1), array->rows, /*take=*/false, arena, registry, ctx,
                                              row_lo, row_hi, err)) {
    return err;
  }
  std::uint32_t col_lo = 0;
  std::uint32_t col_hi = array->cols;
  if (arity == 3U) {
    if (!dynamic_array::resolve_take_drop_range(&call.as_call_arg(2), array->cols, /*take=*/false, arena, registry, ctx,
                                                col_lo, col_hi, err)) {
      return err;
    }
  }

  ArrayValue* out = dynamic_array::materialise_slice(*array, row_lo, row_hi, col_lo, col_hi, arena);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  return Value::array(out);
}

}  // namespace eval
}  // namespace formulon
