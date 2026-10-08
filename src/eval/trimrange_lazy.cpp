#include "eval/trimrange_lazy.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string_view>

#include "eval/array_alloc.h"
#include "eval/coerce.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "eval/range_resolvers.h"
#include "eval/shape_ops_lazy.h"
#include "eval/tree_walker/dispatch.h"
#include "parser/ast.h"
#include "parser/reference.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/strings.h"
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

// The row and column trim modes of a TRIMRANGE or `_TRO_*` call; false with
// the error in `*out_err` for a bad arity or mode argument.
bool read_modes(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                int* trim_rows, int* trim_cols, Value* out_err) {
  const std::uint32_t arity = call.as_call_arity();
  const std::string_view name = strip_future_prefix(call.as_call_name());
  if (!strings::case_insensitive_eq(name, "TRIMRANGE")) {
    if (arity != 1U) {
      *out_err = Value::error(ErrorCode::Value);
      return false;
    }
    for (const parser::TrimRefMode m :
         {parser::TrimRefMode::Leading, parser::TrimRefMode::Trailing, parser::TrimRefMode::Both}) {
      if (strings::case_insensitive_eq(name, parser::trim_ref_function_name(m))) {
        *trim_rows = *trim_cols = static_cast<int>(m);
      }
    }
    return true;
  }
  if (arity < 1U || arity > 3U) {
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  *trim_rows = 3;
  *trim_cols = 3;
  if (arity >= 2U) {
    *trim_rows = read_mode_opt(call.as_call_arg(1), arena, registry, ctx, 3, out_err);
    if (out_err->is_error()) {
      return false;
    }
  }
  if (arity >= 3U) {
    *trim_cols = read_mode_opt(call.as_call_arg(2), arena, registry, ctx, 3, out_err);
    if (out_err->is_error()) {
      return false;
    }
  }
  return true;
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

// TRIMRANGE over a value that is no reference (an array literal, an
// expression): drops the Blank edges `trim_rows` / `trim_cols` select.
Value trim_array_value(const parser::AstNode& source, const parser::AstNode& call, Arena& arena,
                       const FunctionRegistry& registry, const EvalContext& ctx) {
  // Evaluate the source in array context so 2D shape is preserved. Errors
  // propagate as a scalar.
  const Value src = eval_node_as_array(source, arena, registry, ctx);
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
  int trim_cols = 3;
  if (!read_modes(call, arena, registry, ctx, &trim_rows, &trim_cols, &err)) {
    return err;
  }

  const bool trim_leading_rows = (trim_rows & 1) != 0;
  const bool trim_trailing_rows = (trim_rows & 2) != 0;
  const bool trim_leading_cols = (trim_cols & 1) != 0;
  const bool trim_trailing_cols = (trim_cols & 2) != 0;

  const std::uint32_t rows_in = in->rows;
  const std::uint32_t cols_in = in->cols;

  // Defensive: a 0-row or 0-col input cannot be trimmed any further, and an
  // empty array surfaces #REF! to match Mac Excel 365 (ja-JP).
  if (rows_in == 0U || cols_in == 0U) {
    return Value::error(ErrorCode::Ref);
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
    return Value::error(ErrorCode::Ref);
  }

  const std::uint32_t kept_rows = row_end - row_start;
  const std::uint32_t kept_cols = col_end - col_start;
  Value* buffer = nullptr;
  ArrayValue* out = allocate_array_value(kept_rows, kept_cols, arena, buffer, kMaxDerivedArrayCells);
  if (out == nullptr) {
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

  return Value::array(out);
}

// TRIMRANGE and the trim operators as a value: a reference source is read
// over its trimmed rectangle like any range, anything else is trimmed by value.
Value eval_trim(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx) {
  if (call.as_call_arity() < 1U) {
    return Value::error(ErrorCode::Value);
  }
  std::string_view sheet;
  std::uint32_t top = 0;
  std::uint32_t left = 0;
  std::uint32_t bottom = 0;
  std::uint32_t right = 0;
  ErrorCode err = ErrorCode::Value;
  if (!resolve_reference_rect(call.as_call_arg(0), arena, registry, ctx, &sheet, &top, &left, &bottom, &right, &err)) {
    return trim_array_value(call.as_call_arg(0), call, arena, registry, ctx);
  }
  bool is_range = false;
  if (!resolve_trim_reference(call, arena, registry, ctx, &sheet, &top, &left, &bottom, &right, &is_range, &err)) {
    return Value::error(err);
  }
  parser::Reference top_left{};
  top_left.sheet = sheet;
  top_left.row = top;
  top_left.col = left;
  parser::Reference bottom_right = top_left;
  bottom_right.row = bottom;
  bottom_right.col = right;
  parser::AstNode* lhs = parser::make_ref(arena, top_left);
  parser::AstNode* rhs = parser::make_ref(arena, bottom_right);
  const parser::AstNode* rect = lhs != nullptr && rhs != nullptr ? parser::make_range_op(arena, lhs, rhs) : nullptr;
  if (rect == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  return eval_node_as_array(*rect, arena, registry, ctx);
}

}  // namespace

bool resolve_trim_reference(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                            const EvalContext& ctx, std::string_view* out_sheet, std::uint32_t* out_top_row,
                            std::uint32_t* out_left_col, std::uint32_t* out_bottom_row, std::uint32_t* out_right_col,
                            bool* out_is_range, ErrorCode* out_err) {
  if (call.as_call_arity() < 1U) {
    *out_err = ErrorCode::Value;
    return false;
  }
  std::string_view sheet;
  std::uint32_t top = 0;
  std::uint32_t left = 0;
  std::uint32_t bottom = 0;
  std::uint32_t right = 0;
  if (!resolve_reference_rect(call.as_call_arg(0), arena, registry, ctx, &sheet, &top, &left, &bottom, &right,
                              out_err)) {
    return false;
  }
  int trim_rows = 3;
  int trim_cols = 3;
  Value mode_err = Value::blank();
  if (!read_modes(call, arena, registry, ctx, &trim_rows, &trim_cols, &mode_err)) {
    *out_err = mode_err.as_error();
    return false;
  }
  const Sheet* target = ctx.sheet_for_qualifier(sheet);
  if (target == nullptr) {
    *out_err = ErrorCode::Ref;
    return false;
  }
  // Blankness is whether a cell holds anything, a formula included, so no
  // cell is evaluated to find the edges: a formula inside its own range
  // (`=COLUMNS(1:.1)` in row 1) is no circular read.
  const std::optional<Sheet::PopulatedExtent> extent = target->populated_extent(top, left, bottom, right);
  if (!extent) {
    *out_err = ErrorCode::Ref;
    return false;
  }
  *out_sheet = sheet;
  *out_top_row = (trim_rows & 1) != 0 ? extent->first_row : top;
  *out_bottom_row = (trim_rows & 2) != 0 ? extent->last_row : bottom;
  *out_left_col = (trim_cols & 1) != 0 ? extent->first_col : left;
  *out_right_col = (trim_cols & 2) != 0 ? extent->last_col : right;
  *out_is_range = *out_top_row != *out_bottom_row || *out_left_col != *out_right_col;
  return true;
}

Value eval_trimrange_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx) {
  if (call.as_call_arity() < 1U || call.as_call_arity() > 3U) {
    return Value::error(ErrorCode::Value);
  }
  return eval_trim(call, arena, registry, ctx);
}

Value eval_trim_ref_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx) {
  if (call.as_call_arity() != 1U) {
    return Value::error(ErrorCode::Value);
  }
  return eval_trim(call, arena, registry, ctx);
}

}  // namespace eval
}  // namespace formulon
