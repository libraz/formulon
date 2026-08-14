
#include "eval/dynamic_array/reshape.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <vector>

#include "eval/dynamic_array/common.h"
#include "eval/lazy_impls.h"
#include "eval/omitted_arg.h"
#include "eval/range_args.h"
#include "eval/shape_ops_lazy.h"
#include "parser/ast.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {

namespace {

/// Resolve every call argument into an `ArrayValue` for stacking. Unlike
/// the generic `resolve_array_value`, a scalar argument is treated as a
/// 1x1 block regardless of its kind: a scalar Number, Text, Bool, Blank,
/// or Error all become a single cell that participates in the stack and is
/// NA-padded against taller / wider neighbours. In Excel an error value
/// passed to HSTACK / VSTACK (e.g. `=HSTACK({1;2}, NA())`) is a normal
/// cell, not a propagating error, so it must not short-circuit the result.
///
/// Returns true on success and populates `out`. The returned `ArrayValue`
/// pointers borrow from the caller arena, which must outlive their use.
/// The only failure path is arena exhaustion while wrapping a scalar.
bool resolve_stack_args(const parser::AstNode& call, std::uint32_t arity, Arena& arena,
                        const FunctionRegistry& registry, const EvalContext& ctx, std::vector<const ArrayValue*>& out,
                        Value& error_out) {
  out.reserve(arity);
  for (std::uint32_t i = 0; i < arity; ++i) {
    // `eval_node_as_array` already wraps non-error scalars into 1x1 arrays
    // and expands ranges / array literals; it only returns a scalar when
    // the argument evaluated to an error. Wrap that error into a 1x1 block
    // so it stacks like any other cell.
    const Value v = eval_node_as_array(call.as_call_arg(i), arena, registry, ctx);
    if (v.is_array()) {
      out.push_back(v.as_array());
      continue;
    }
    Value* buffer = nullptr;
    ArrayValue* cell = dynamic_array::allocate_array_value(1U, 1U, arena, buffer, kMaxDerivedArrayCells);
    if (cell == nullptr) {
      error_out = Value::error(ErrorCode::Num);
      return false;
    }
    buffer[0] = v;
    out.push_back(cell);
  }
  return true;
}

Value materialise_stacked_arrays(const std::vector<const ArrayValue*>& arrays, bool horizontal, Arena& arena) {
  std::size_t out_rows_sz = 0;
  std::size_t out_cols_sz = 0;
  if (horizontal) {
    for (const ArrayValue* a : arrays) {
      out_rows_sz = std::max(out_rows_sz, static_cast<std::size_t>(a->rows));
      out_cols_sz += static_cast<std::size_t>(a->cols);
    }
  } else {
    for (const ArrayValue* a : arrays) {
      out_rows_sz += static_cast<std::size_t>(a->rows);
      out_cols_sz = std::max(out_cols_sz, static_cast<std::size_t>(a->cols));
    }
  }
  if (out_rows_sz > static_cast<std::size_t>(0xFFFFFFFFU) || out_cols_sz > static_cast<std::size_t>(0xFFFFFFFFU)) {
    return Value::error(ErrorCode::Num);
  }

  const auto out_rows = static_cast<std::uint32_t>(out_rows_sz);
  const auto out_cols = static_cast<std::uint32_t>(out_cols_sz);
  Value* buffer = nullptr;
  ArrayValue* out = dynamic_array::allocate_array_value(out_rows, out_cols, arena, buffer, kMaxDerivedArrayCells);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }

  if (horizontal) {
    std::uint32_t col_off = 0;
    for (const ArrayValue* a : arrays) {
      for (std::uint32_t r = 0; r < out_rows; ++r) {
        for (std::uint32_t c = 0; c < a->cols; ++c) {
          const std::size_t dst = static_cast<std::size_t>(r) * out_cols + col_off + c;
          buffer[dst] =
              (r < a->rows)
                  ? a->cells[static_cast<std::size_t>(r) * a->cols + c].promote_reference_blank_to_value_array()
                  : Value::error(ErrorCode::NA);
        }
      }
      col_off += a->cols;
    }
  } else {
    std::uint32_t row_off = 0;
    for (const ArrayValue* a : arrays) {
      for (std::uint32_t r = 0; r < a->rows; ++r) {
        for (std::uint32_t c = 0; c < out_cols; ++c) {
          const std::size_t dst = static_cast<std::size_t>(row_off + r) * out_cols + c;
          buffer[dst] =
              (c < a->cols)
                  ? a->cells[static_cast<std::size_t>(r) * a->cols + c].promote_reference_blank_to_value_array()
                  : Value::error(ErrorCode::NA);
        }
      }
      row_off += a->rows;
    }
  }

  return Value::array(out);
}

}  // namespace

Value eval_hstack_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  // Excel's documented HSTACK accepts up to 254 array references; reject
  // 0-arg explicitly (no array context to spill into) and use 254 as the
  // safety ceiling.
  if (arity < 1U || arity > 254U) {
    return Value::error(ErrorCode::Value);
  }

  std::vector<const ArrayValue*> arrays;
  Value err = Value::error(ErrorCode::Value);
  if (!resolve_stack_args(call, arity, arena, registry, ctx, arrays, err)) {
    return err;
  }
  return materialise_stacked_arrays(arrays, /*horizontal=*/true, arena);
}

Value eval_vstack_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 1U || arity > 254U) {
    return Value::error(ErrorCode::Value);
  }

  std::vector<const ArrayValue*> arrays;
  Value err = Value::error(ErrorCode::Value);
  if (!resolve_stack_args(call, arity, arena, registry, ctx, arrays, err)) {
    return err;
  }
  return materialise_stacked_arrays(arrays, /*horizontal=*/false, arena);
}

Value eval_expand_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 2U || arity > 4U) {
    return Value::error(ErrorCode::Value);
  }

  const ArrayValue* array = nullptr;
  Value err = Value::error(ErrorCode::Value);
  if (!resolve_array_value(call.as_call_arg(0), arena, registry, ctx, &array, &err)) {
    return err;
  }

  // Resolve the new row count. An omitted `rows` slot keeps the source
  // height, matching Excel's optional-argument default.
  double rows_d = static_cast<double>(array->rows);
  if (!is_omitted_arg(call.as_call_arg(1)) &&
      !dynamic_array::eval_truncated_number_arg(call.as_call_arg(1), arena, registry, ctx, rows_d, err)) {
    return err;
  }
  if (rows_d < static_cast<double>(array->rows)) {
    return Value::error(ErrorCode::Value);
  }
  if (rows_d > static_cast<double>(Sheet::kMaxRows)) {
    return Value::error(ErrorCode::Num);
  }

  // Resolve the new column count. Optional; defaults to the existing
  // column count (no horizontal expansion).
  std::uint32_t out_cols = array->cols;
  if (arity >= 3U && !is_omitted_arg(call.as_call_arg(2))) {
    double cols_d = 0.0;
    if (!dynamic_array::eval_truncated_number_arg(call.as_call_arg(2), arena, registry, ctx, cols_d, err)) {
      return err;
    }
    if (cols_d < static_cast<double>(array->cols)) {
      return Value::error(ErrorCode::Value);
    }
    if (cols_d > static_cast<double>(Sheet::kMaxCols)) {
      return Value::error(ErrorCode::Num);
    }
    out_cols = static_cast<std::uint32_t>(cols_d);
  }
  const auto out_rows = static_cast<std::uint32_t>(rows_d);

  // Resolve the pad value. An absent `pad_with` defaults to #N/A. A present
  // empty slot evaluates to Blank, which remains Blank for nested consumers;
  // the same applies to a blank first cell selected from an array pad.
  // Explicit empty text remains text. If `pad_with` is supplied as an array
  // (range / range literal), Excel takes the scalar at cell 0 and pads with
  // that. Errors at the argument level propagate.
  Value pad = Value::error(ErrorCode::NA);
  if (arity == 4U) {
    const Value pad_v = eval_node(call.as_call_arg(3), arena, registry, ctx);
    if (pad_v.is_error()) {
      return pad_v;
    }
    if (pad_v.is_array()) {
      const ArrayValue* pa = pad_v.as_array();
      if (pa->rows == 0U || pa->cols == 0U) {
        return Value::error(ErrorCode::Value);
      }
      pad = pa->cells[0];
    } else {
      pad = pad_v;
    }
  }

  Value* buffer = nullptr;
  ArrayValue* out = dynamic_array::allocate_array_value(out_rows, out_cols, arena, buffer);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  for (std::uint32_t r = 0; r < out_rows; ++r) {
    for (std::uint32_t c = 0; c < out_cols; ++c) {
      const std::size_t dst = static_cast<std::size_t>(r) * out_cols + c;
      if (r < array->rows && c < array->cols) {
        buffer[dst] =
            array->cells[static_cast<std::size_t>(r) * array->cols + c].promote_reference_blank_to_value_array();
      } else {
        // Generated pad cells get the direct-grid zero projection, while a
        // source Blank copied from the input array remains untouched.
        buffer[dst] = pad.is_blank() ? Value::blank(BlankGridProjection::kValueArrayZero)
                                     : pad.promote_reference_blank_to_value_array();
      }
    }
  }

  return Value::array(out);
}

}  // namespace eval
}  // namespace formulon
