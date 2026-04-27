// Copyright 2026 libraz. Licensed under the MIT License.

#include "eval/trimrange_lazy.h"

#include <cmath>
#include <cstddef>
#include <cstdint>

#include "eval/coerce.h"
#include "eval/lazy_impls.h"
#include "eval/shape_ops_lazy.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {

namespace {

// Reads an optional integer trim-mode argument with a default. Validates that
// the coerced integer lies in [0, 3]; anything else surfaces #VALUE! through
// *out_err. Argument errors propagate verbatim.
int read_mode_opt(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                  int default_value, Value* out_err) {
  const Value v = eval_node(node, arena, registry, ctx);
  if (v.is_error()) {
    *out_err = v;
    return default_value;
  }
  auto coerced = coerce_to_number(v);
  if (!coerced) {
    *out_err = Value::error(coerced.error());
    return default_value;
  }
  const int mode = static_cast<int>(std::trunc(coerced.value()));
  if (mode < 0 || mode > 3) {
    *out_err = Value::error(ErrorCode::Value);
    return default_value;
  }
  return mode;
}

// Returns true iff every cell in row `r` of `arr` is the Blank variant. Empty
// text / 0 / FALSE / errors all count as non-blank, matching Mac Excel.
bool row_is_blank(const ArrayValue& arr, std::uint32_t r) {
  const std::size_t cols = static_cast<std::size_t>(arr.cols);
  const std::size_t base = static_cast<std::size_t>(r) * cols;
  for (std::size_t c = 0; c < cols; ++c) {
    if (!arr.cells[base + c].is_blank()) {
      return false;
    }
  }
  return true;
}

bool col_is_blank(const ArrayValue& arr, std::uint32_t c) {
  const std::size_t cols = static_cast<std::size_t>(arr.cols);
  const std::size_t rows = static_cast<std::size_t>(arr.rows);
  for (std::size_t r = 0; r < rows; ++r) {
    if (!arr.cells[r * cols + static_cast<std::size_t>(c)].is_blank()) {
      return false;
    }
  }
  return true;
}

}  // namespace

Value eval_trimrange_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 1U || arity > 3U) {
    return Value::error(ErrorCode::Value);
  }

  // Evaluate the source in array context so 2D shape is preserved. Errors
  // propagate as a scalar.
  const Value src = eval_node_as_array(call.as_call_arg(0), arena, registry, ctx);
  if (src.is_error()) {
    return src;
  }
  // `eval_node_as_array` is contracted to return either an Array or a scalar
  // Error; the is_array() check is defensive against future API drift.
  if (!src.is_array()) {
    return Value::error(ErrorCode::Value);
  }
  const ArrayValue* in = src.as_array();
  // A scalar expression that errored (e.g. `1/0`) reaches us as a 1x1 array
  // with a single error cell because `eval_node_as_array` broadcasts arithmetic
  // cellwise. Surface that as a scalar error so callers see Mac Excel's
  // visible spill (`=TRIMRANGE(1/0)` shows `#DIV/0!`, not an array).
  if (in->rows == 1U && in->cols == 1U && in->cells[0].is_error()) {
    return in->cells[0];
  }

  Value err = Value::blank();
  int trim_rows = 3;
  if (arity >= 2U) {
    trim_rows = read_mode_opt(call.as_call_arg(1), arena, registry, ctx, 3, &err);
    if (err.is_error()) {
      return err;
    }
  }
  int trim_cols = 3;
  if (arity >= 3U) {
    trim_cols = read_mode_opt(call.as_call_arg(2), arena, registry, ctx, 3, &err);
    if (err.is_error()) {
      return err;
    }
  }

  const bool trim_leading_rows = (trim_rows & 1) != 0;
  const bool trim_trailing_rows = (trim_rows & 2) != 0;
  const bool trim_leading_cols = (trim_cols & 1) != 0;
  const bool trim_trailing_cols = (trim_cols & 2) != 0;

  const std::uint32_t rows_in = in->rows;
  const std::uint32_t cols_in = in->cols;

  // Defensive: a 0-row or 0-col input cannot be trimmed any further, and an
  // empty array would surface #CALC! anyway.
  if (rows_in == 0U || cols_in == 0U) {
    return Value::error(ErrorCode::Calc);
  }

  std::uint32_t row_start = 0;
  if (trim_leading_rows) {
    while (row_start < rows_in && row_is_blank(*in, row_start)) {
      ++row_start;
    }
  }
  std::uint32_t row_end = rows_in;  // exclusive
  if (trim_trailing_rows) {
    while (row_end > row_start && row_is_blank(*in, row_end - 1U)) {
      --row_end;
    }
  }

  std::uint32_t col_start = 0;
  if (trim_leading_cols) {
    while (col_start < cols_in && col_is_blank(*in, col_start)) {
      ++col_start;
    }
  }
  std::uint32_t col_end = cols_in;  // exclusive
  if (trim_trailing_cols) {
    while (col_end > col_start && col_is_blank(*in, col_end - 1U)) {
      --col_end;
    }
  }

  if (row_end <= row_start || col_end <= col_start) {
    return Value::error(ErrorCode::Calc);
  }

  const std::uint32_t kept_rows = row_end - row_start;
  const std::uint32_t kept_cols = col_end - col_start;
  const std::size_t total = static_cast<std::size_t>(kept_rows) * static_cast<std::size_t>(kept_cols);

  Value* buffer = arena.create_array<Value>(total);
  if (buffer == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  const std::size_t in_cols = static_cast<std::size_t>(cols_in);
  for (std::uint32_t r = 0; r < kept_rows; ++r) {
    const std::size_t src_row_base = (static_cast<std::size_t>(row_start) + r) * in_cols;
    const std::size_t dst_row_base = static_cast<std::size_t>(r) * static_cast<std::size_t>(kept_cols);
    for (std::uint32_t c = 0; c < kept_cols; ++c) {
      buffer[dst_row_base + c] = in->cells[src_row_base + col_start + c];
    }
  }

  ArrayValue* out = arena.create<ArrayValue>();
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  out->rows = kept_rows;
  out->cols = kept_cols;
  out->cells = buffer;
  return Value::array(out);
}

}  // namespace eval
}  // namespace formulon
