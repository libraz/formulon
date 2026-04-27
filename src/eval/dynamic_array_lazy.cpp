// Copyright 2026 libraz. Licensed under the MIT License.

#include "eval/dynamic_array_lazy.h"

#include <cstddef>
#include <cstdint>
#include <vector>

#include "eval/coerce.h"
#include "eval/lazy_impls.h"
#include "eval/shape_ops_lazy.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {

Value eval_filter_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 2U || arity > 3U) {
    return Value::error(ErrorCode::Value);
  }

  // Resolve `array` and `include` as Value::Array via the array-context seam
  // so range-shaped args (Ref / RangeOp / OFFSET / CHOOSE / IF / SpillRef)
  // keep their 2D shape. Errors short-circuit before we touch shape logic.
  const Value array_v = eval_node_as_array(call.as_call_arg(0), arena, registry, ctx);
  if (array_v.is_error()) {
    return array_v;
  }
  if (!array_v.is_array()) {
    return Value::error(ErrorCode::Value);
  }
  const Value include_v = eval_node_as_array(call.as_call_arg(1), arena, registry, ctx);
  if (include_v.is_error()) {
    return include_v;
  }
  if (!include_v.is_array()) {
    return Value::error(ErrorCode::Value);
  }

  const ArrayValue* array = array_v.as_array();
  const ArrayValue* include = include_v.as_array();

  // Axis decision. `include` must be a 1D vector matching one of `array`'s
  // dimensions. Both shapes equal -> filter rows (the column-vector form is
  // the dominant Excel idiom). Otherwise either an exact column-vector
  // match (rows-axis) or row-vector match (cols-axis); anything else is
  // ambiguous / mismatched.
  enum class Axis { Rows, Cols };
  Axis axis = Axis::Rows;
  if (include->rows == array->rows && include->cols == 1U) {
    axis = Axis::Rows;
  } else if (include->cols == array->cols && include->rows == 1U) {
    axis = Axis::Cols;
  } else {
    return Value::error(ErrorCode::Value);
  }

  // Coerce each include cell to bool, recording the kept indices. Coercion
  // errors and explicit Error cells short-circuit the whole call -- Mac
  // Excel's conservative behaviour for mixed-shape spill candidates.
  const std::uint32_t mask_n = (axis == Axis::Rows) ? include->rows : include->cols;
  std::vector<std::uint32_t> kept;
  kept.reserve(mask_n);
  for (std::uint32_t i = 0; i < mask_n; ++i) {
    const Value& cell = include->cells[i];
    if (cell.is_error()) {
      return cell;
    }
    auto coerced = coerce_to_bool(cell);
    if (!coerced) {
      return Value::error(coerced.error());
    }
    if (coerced.value()) {
      kept.push_back(i);
    }
  }

  // Empty-result handling: return `if_empty` if provided (scalar, NOT
  // spilled), otherwise #CALC! per Mac Excel's documented surface.
  if (kept.empty()) {
    if (arity == 3U) {
      return eval_node(call.as_call_arg(2), arena, registry, ctx);
    }
    return Value::error(ErrorCode::Calc);
  }

  // Build the output array. For axis=Rows the output is (kept.size(),
  // array->cols); for axis=Cols it is (array->rows, kept.size()). Cells are
  // copied row-major.
  std::uint32_t out_rows = 0;
  std::uint32_t out_cols = 0;
  if (axis == Axis::Rows) {
    out_rows = static_cast<std::uint32_t>(kept.size());
    out_cols = array->cols;
  } else {
    out_rows = array->rows;
    out_cols = static_cast<std::uint32_t>(kept.size());
  }
  const std::size_t total = static_cast<std::size_t>(out_rows) * static_cast<std::size_t>(out_cols);
  Value* buffer = arena.create_array<Value>(total);
  if (buffer == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  if (axis == Axis::Rows) {
    for (std::size_t i = 0; i < kept.size(); ++i) {
      const std::uint32_t src_row = kept[i];
      for (std::uint32_t c = 0; c < array->cols; ++c) {
        buffer[i * static_cast<std::size_t>(out_cols) + c] =
            array->cells[static_cast<std::size_t>(src_row) * array->cols + c];
      }
    }
  } else {
    for (std::uint32_t r = 0; r < array->rows; ++r) {
      for (std::size_t i = 0; i < kept.size(); ++i) {
        const std::uint32_t src_col = kept[i];
        buffer[static_cast<std::size_t>(r) * out_cols + i] =
            array->cells[static_cast<std::size_t>(r) * array->cols + src_col];
      }
    }
  }

  ArrayValue* out = arena.create<ArrayValue>();
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  out->rows = out_rows;
  out->cols = out_cols;
  out->cells = buffer;
  return Value::array(out);
}

}  // namespace eval
}  // namespace formulon
