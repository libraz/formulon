
#include "eval/dynamic_array/common.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <string_view>
#include <vector>

#include "eval/coerce.h"
#include "eval/lazy_impls.h"
#include "eval/omitted_arg.h"
#include "parser/ast.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/checked_mul.h"
#include "utils/error.h"
#include "utils/resource_budget.h"
#include "utils/strings.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace dynamic_array {

bool eval_truncated_number_arg(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                               const EvalContext& ctx, double& out, Value& error_out) {
  const Value v = eval_node(node, arena, registry, ctx);
  if (v.is_error()) {
    error_out = v;
    return false;
  }
  auto coerced = coerce_to_number(v);
  if (!coerced) {
    error_out = Value::error(coerced.error());
    return false;
  }
  out = std::trunc(coerced.value());
  return true;
}

ArrayValue* allocate_array_value(std::uint32_t rows, std::uint32_t cols, Arena& arena, Value*& out_buffer) {
  out_buffer = nullptr;
  if (rows == 0U || cols == 0U || rows > Sheet::kMaxRows || cols > Sheet::kMaxCols) {
    return nullptr;
  }
  const auto total = checked_mul_size_t(rows, cols);
  if (!total || total.value() > kMaxDynamicArrayCells) {
    return nullptr;
  }
  Value* buffer = arena.create_array<Value>(total.value());
  if (buffer == nullptr) {
    return nullptr;
  }
  ArrayValue* out = arena.create<ArrayValue>();
  if (out == nullptr) {
    return nullptr;
  }
  out->rows = rows;
  out->cols = cols;
  out->cells = buffer;
  out_buffer = buffer;
  return out;
}

ArrayValue* materialise_selected_lanes(const ArrayValue& src, const std::vector<std::uint32_t>& indices, bool by_col,
                                       Arena& arena) {
  const std::uint32_t out_rows = by_col ? src.rows : static_cast<std::uint32_t>(indices.size());
  const std::uint32_t out_cols = by_col ? static_cast<std::uint32_t>(indices.size()) : src.cols;
  Value* buffer = nullptr;
  ArrayValue* out = allocate_array_value(out_rows, out_cols, arena, buffer);
  if (out == nullptr) {
    return nullptr;
  }
  if (by_col) {
    for (std::uint32_t r = 0; r < out_rows; ++r) {
      for (std::uint32_t i = 0; i < out_cols; ++i) {
        const std::uint32_t src_col = indices[i];
        buffer[static_cast<std::size_t>(r) * out_cols + i] =
            src.cells[static_cast<std::size_t>(r) * src.cols + src_col];
      }
    }
  } else {
    for (std::uint32_t i = 0; i < out_rows; ++i) {
      const std::uint32_t src_row = indices[i];
      for (std::uint32_t c = 0; c < out_cols; ++c) {
        buffer[static_cast<std::size_t>(i) * out_cols + c] =
            src.cells[static_cast<std::size_t>(src_row) * src.cols + c];
      }
    }
  }

  return out;
}

ArrayValue* materialise_slice(const ArrayValue& src, std::uint32_t row_lo, std::uint32_t row_hi, std::uint32_t col_lo,
                              std::uint32_t col_hi, Arena& arena) {
  const std::uint32_t out_rows = row_hi - row_lo;
  const std::uint32_t out_cols = col_hi - col_lo;
  Value* buffer = nullptr;
  ArrayValue* out = allocate_array_value(out_rows, out_cols, arena, buffer);
  if (out == nullptr) {
    return nullptr;
  }
  for (std::uint32_t r = 0; r < out_rows; ++r) {
    for (std::uint32_t c = 0; c < out_cols; ++c) {
      buffer[static_cast<std::size_t>(r) * out_cols + c] =
          src.cells[static_cast<std::size_t>(row_lo + r) * src.cols + (col_lo + c)];
    }
  }
  return out;
}

ArrayValue* materialise_vector(std::vector<Value>&& cells, bool as_column, Arena& arena) {
  const auto n = static_cast<std::uint32_t>(cells.size());
  const std::uint32_t rows = as_column ? n : 1U;
  const std::uint32_t cols = as_column ? 1U : n;
  Value* buffer = nullptr;
  ArrayValue* out = allocate_array_value(rows, cols, arena, buffer);
  if (out == nullptr) {
    return nullptr;
  }
  for (std::size_t i = 0; i < cells.size(); ++i) {
    buffer[i] = cells[i];
  }
  return out;
}

bool unique_cell_equal(const Value& a, const Value& b) {
  if (a.kind() != b.kind()) {
    return false;
  }
  switch (a.kind()) {
    case ValueKind::Blank:
      return true;
    case ValueKind::Number:
      return a.as_number() == b.as_number();
    case ValueKind::Bool:
      return a.as_boolean() == b.as_boolean();
    case ValueKind::Error:
      return a.as_error() == b.as_error();
    case ValueKind::Text:
      return strings::case_insensitive_eq(a.as_text(), b.as_text());
    default:
      // Array / Ref / Lambda are not produced by cell reads; treat any
      // hypothetical match conservatively as "not equal" to avoid silent
      // dedup of structurally complex values we have no canonical form for.
      return false;
  }
}

bool unique_lane_equal(const ArrayValue& arr, std::uint32_t i, std::uint32_t j, bool by_col) {
  if (by_col) {
    for (std::uint32_t r = 0; r < arr.rows; ++r) {
      const Value& av = arr.cells[static_cast<std::size_t>(r) * arr.cols + i];
      const Value& bv = arr.cells[static_cast<std::size_t>(r) * arr.cols + j];
      if (!unique_cell_equal(av, bv)) {
        return false;
      }
    }
    return true;
  }
  for (std::uint32_t c = 0; c < arr.cols; ++c) {
    const Value& av = arr.cells[static_cast<std::size_t>(i) * arr.cols + c];
    const Value& bv = arr.cells[static_cast<std::size_t>(j) * arr.cols + c];
    if (!unique_cell_equal(av, bv)) {
      return false;
    }
  }
  return true;
}

namespace {

// Excel-canonical kind rank for SORT. Lower rank sorts earlier under
// ascending order. Blank deliberately sits at rank 5 -- the SORT impl
// special-cases blanks so they always sink to the end regardless of
// `sort_order`. Numbers / Text / Bool / Error participate normally.
int sort_kind_rank(const Value& v) {
  switch (v.kind()) {
    case ValueKind::Number:
      return 0;
    case ValueKind::Text:
      return 1;
    case ValueKind::Bool:
      return 2;
    case ValueKind::Error:
      return 3;
    case ValueKind::Blank:
      return 5;
    default:
      // Array / Ref / Lambda are not produced by cell reads. Slot them
      // between Error and Blank so they participate predictably without
      // colliding with real-world cells.
      return 4;
  }
}

}  // namespace

bool ascii_ci_less(std::string_view a, std::string_view b) {
  const std::size_t n = std::min(a.size(), b.size());
  for (std::size_t i = 0; i < n; ++i) {
    const unsigned char ca = static_cast<unsigned char>(a[i]);
    const unsigned char cb = static_cast<unsigned char>(b[i]);
    const unsigned char la = (ca >= 'A' && ca <= 'Z') ? static_cast<unsigned char>(ca + 32) : ca;
    const unsigned char lb = (cb >= 'A' && cb <= 'Z') ? static_cast<unsigned char>(cb + 32) : cb;
    if (la != lb) {
      return la < lb;
    }
  }
  return a.size() < b.size();
}

bool sort_cell_less_asc(const Value& a, const Value& b) {
  const int ra = sort_kind_rank(a);
  const int rb = sort_kind_rank(b);
  if (ra != rb) {
    return ra < rb;
  }
  switch (a.kind()) {
    case ValueKind::Number:
      return a.as_number() < b.as_number();
    case ValueKind::Text:
      return ascii_ci_less(a.as_text(), b.as_text());
    case ValueKind::Bool:
      return !a.as_boolean() && b.as_boolean();
    case ValueKind::Error:
      return static_cast<int>(a.as_error()) < static_cast<int>(b.as_error());
    case ValueKind::Blank:
    default:
      return false;
  }
}

bool sort_lane_less(const Value& key_a, const Value& key_b, bool descending) {
  const bool a_blank = key_a.is_blank();
  const bool b_blank = key_b.is_blank();
  if (a_blank != b_blank) {
    // Non-blank wins (sorts earlier) regardless of direction.
    return !a_blank;
  }
  if (a_blank && b_blank) {
    return false;
  }
  return descending ? sort_cell_less_asc(key_b, key_a) : sort_cell_less_asc(key_a, key_b);
}

bool resolve_sort_order_arg(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                            const EvalContext& ctx, bool& descending, Value& error_out) {
  if (is_omitted_arg(node)) {
    descending = false;
    return true;
  }
  const Value v = eval_node(node, arena, registry, ctx);
  if (v.is_error()) {
    error_out = v;
    return false;
  }
  auto coerced = coerce_to_number(v);
  if (!coerced) {
    error_out = Value::error(coerced.error());
    return false;
  }
  const double n = coerced.value();
  if (n == 1.0) {
    descending = false;
    return true;
  }
  if (n == -1.0) {
    descending = true;
    return true;
  }
  error_out = Value::error(ErrorCode::Value);
  return false;
}

bool resolve_choose_index(const parser::AstNode& node, std::uint32_t axis_size, Arena& arena,
                          const FunctionRegistry& registry, const EvalContext& ctx, std::uint32_t& out,
                          Value& error_out) {
  double truncated = 0.0;
  if (!eval_truncated_number_arg(node, arena, registry, ctx, truncated, error_out)) {
    return false;
  }
  // Truncate-toward-zero on the user-supplied index. `0` after truncation
  // is invalid; positives map to `[1, axis_size]`, negatives to
  // `[-axis_size, -1]`. Anything else surfaces #VALUE!.
  if (truncated == 0.0) {
    error_out = Value::error(ErrorCode::Value);
    return false;
  }
  if (truncated > 0.0) {
    if (truncated > static_cast<double>(axis_size)) {
      error_out = Value::error(ErrorCode::Value);
      return false;
    }
    out = static_cast<std::uint32_t>(truncated) - 1U;
    return true;
  }
  // Negative index path. `-1` maps to the last element (`axis_size - 1`).
  const double abs_idx = -truncated;
  if (abs_idx > static_cast<double>(axis_size)) {
    error_out = Value::error(ErrorCode::Value);
    return false;
  }
  out = axis_size - static_cast<std::uint32_t>(abs_idx);
  return true;
}

bool resolve_take_drop_range(const parser::AstNode* node, std::uint32_t axis_size, bool take, Arena& arena,
                             const FunctionRegistry& registry, const EvalContext& ctx, std::uint32_t& lo,
                             std::uint32_t& hi, Value& error_out) {
  if (node == nullptr) {
    lo = 0;
    hi = axis_size;
    return true;
  }
  double count = 0.0;
  if (!eval_truncated_number_arg(*node, arena, registry, ctx, count, error_out)) {
    return false;
  }
  if (take) {
    if (count == 0.0) {
      error_out = Value::error(ErrorCode::Calc);
      return false;
    }
    if (count > 0.0) {
      const auto take_n = (count >= static_cast<double>(axis_size)) ? axis_size : static_cast<std::uint32_t>(count);
      lo = 0;
      hi = take_n;
      return true;
    }
    const double abs_count = -count;
    const auto take_n =
        (abs_count >= static_cast<double>(axis_size)) ? axis_size : static_cast<std::uint32_t>(abs_count);
    lo = axis_size - take_n;
    hi = axis_size;
    return true;
  }
  if (count >= 0.0) {
    if (count >= static_cast<double>(axis_size)) {
      error_out = Value::error(ErrorCode::Calc);
      return false;
    }
    lo = static_cast<std::uint32_t>(count);
    hi = axis_size;
    return true;
  }
  const double abs_count = -count;
  if (abs_count >= static_cast<double>(axis_size)) {
    error_out = Value::error(ErrorCode::Calc);
    return false;
  }
  lo = 0;
  hi = axis_size - static_cast<std::uint32_t>(abs_count);
  return true;
}

}  // namespace dynamic_array
}  // namespace eval
}  // namespace formulon
