
#include "eval/dynamic_array/layout.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <utility>
#include <vector>

#include "eval/coerce.h"
#include "eval/dynamic_array/common.h"
#include "eval/lazy_impls.h"
#include "eval/range_args.h"
#include "eval/shape_ops_lazy.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/checked_mul.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {

namespace {

/// Walk every cell of `array` and append it to `out` (subject to the
/// `ignore` bitmask) using the iteration order set by `scan_by_column`.
/// Returns `false` on a coercion / argument-validation failure with the
/// caller-visible error written into `error_out`.
bool collect_tocol_torow_cells(const ArrayValue& array, std::int64_t ignore_mask, bool scan_by_column,
                               std::vector<Value>& out, Value& error_out) {
  // Mac Excel restricts the bitmask to 0..3 inclusive. Any other integer
  // surfaces #VALUE!.
  if (ignore_mask < 0 || ignore_mask > 3) {
    error_out = Value::error(ErrorCode::Value);
    return false;
  }
  const bool skip_blanks = (ignore_mask & 1) != 0;
  const bool skip_errors = (ignore_mask & 2) != 0;

  auto consider = [&](const Value& v) {
    if (skip_blanks && v.is_blank()) {
      return;
    }
    if (skip_errors && v.is_error()) {
      return;
    }
    if (v.is_text() && v.as_text().empty()) {
      out.push_back(Value::blank());
      return;
    }
    out.push_back(v);
  };

  if (scan_by_column) {
    for (std::uint32_t c = 0; c < array.cols; ++c) {
      for (std::uint32_t r = 0; r < array.rows; ++r) {
        consider(array.cells[static_cast<std::size_t>(r) * array.cols + c]);
      }
    }
  } else {
    for (std::uint32_t r = 0; r < array.rows; ++r) {
      for (std::uint32_t c = 0; c < array.cols; ++c) {
        consider(array.cells[static_cast<std::size_t>(r) * array.cols + c]);
      }
    }
  }
  return true;
}

/// Resolve the (`ignore`, `scan_by_column`) optional args shared by
/// TOCOL / TOROW. `arity` is the call arity (1..3); the call's first
/// argument is the source array, hence the optional args are at indices
/// 1 and 2 here.
bool resolve_tocol_torow_options(const parser::AstNode& call, std::uint32_t arity, Arena& arena,
                                 const FunctionRegistry& registry, const EvalContext& ctx, std::int64_t& ignore_mask,
                                 bool& scan_by_column, Value& error_out) {
  ignore_mask = 0;
  scan_by_column = false;
  if (arity >= 2U) {
    double ignore_d = 0.0;
    if (!dynamic_array::eval_truncated_number_arg(call.as_call_arg(1), arena, registry, ctx, ignore_d, error_out)) {
      return false;
    }
    ignore_mask = static_cast<std::int64_t>(ignore_d);
  }
  if (arity >= 3U) {
    const Value sc_v = eval_node(call.as_call_arg(2), arena, registry, ctx);
    if (sc_v.is_error()) {
      error_out = sc_v;
      return false;
    }
    auto coerced_b = coerce_to_bool(sc_v);
    if (!coerced_b) {
      error_out = Value::error(coerced_b.error());
      return false;
    }
    scan_by_column = coerced_b.value();
  }
  return true;
}

/// WRAPCOLS / WRAPROWS share an arg layout: a 1D `vector`, an integer
/// `wrap_count >= 1`, and an optional scalar `pad_with` (default #N/A).
/// Returns true and writes `flat`, `wrap_count`, `pad` on success;
/// `false` and the caller-visible error otherwise.
bool resolve_wrap_args(const parser::AstNode& call, std::uint32_t arity, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx, const ArrayValue*& vector_arr, std::uint32_t& wrap_count, Value& pad,
                       Value& error_out) {
  if (arity < 2U || arity > 3U) {
    error_out = Value::error(ErrorCode::Value);
    return false;
  }
  const Value v = eval_node_as_array(call.as_call_arg(0), arena, registry, ctx);
  if (v.is_error()) {
    error_out = v;
    return false;
  }
  if (!v.is_array()) {
    error_out = Value::error(ErrorCode::Value);
    return false;
  }
  const ArrayValue* arr = v.as_array();
  // `vector` must be 1D in either orientation; a 2D rectangle is rejected
  // here (the WRAP family is for vectors only).
  if (arr->rows != 1U && arr->cols != 1U) {
    error_out = Value::error(ErrorCode::Value);
    return false;
  }
  vector_arr = arr;

  double wc_d = 0.0;
  if (!dynamic_array::eval_truncated_number_arg(call.as_call_arg(1), arena, registry, ctx, wc_d, error_out)) {
    return false;
  }
  if (wc_d < 1.0) {
    error_out = Value::error(ErrorCode::Num);
    return false;
  }
  wrap_count = static_cast<std::uint32_t>(wc_d);

  pad = Value::error(ErrorCode::NA);
  if (arity == 3U) {
    const Value p_v = eval_node(call.as_call_arg(2), arena, registry, ctx);
    if (p_v.is_error()) {
      error_out = p_v;
      return false;
    }
    if (p_v.is_array()) {
      const ArrayValue* pa = p_v.as_array();
      if (pa->rows == 0U || pa->cols == 0U) {
        error_out = Value::error(ErrorCode::Value);
        return false;
      }
      pad = pa->cells[0];
    } else {
      pad = p_v;
    }
  }
  return true;
}

/// TOCOL / TOROW. Identical apart from the orientation the kept cells are
/// finally laid out in; the ignore mask and scan order are arguments, not
/// properties of the spelling.
Value eval_to_vector(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx, bool as_column) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 1U || arity > 3U) {
    return Value::error(ErrorCode::Value);
  }
  const ArrayValue* array = nullptr;
  Value err = Value::error(ErrorCode::Value);
  if (!resolve_array_value(call.as_call_arg(0), arena, registry, ctx, &array, &err)) {
    return err;
  }

  std::int64_t ignore_mask = 0;
  bool scan_by_column = false;
  if (!resolve_tocol_torow_options(call, arity, arena, registry, ctx, ignore_mask, scan_by_column, err)) {
    return err;
  }

  // Defensive overflow guard on `rows * cols` before the reserve: on
  // 32-bit `size_t` (WASM) a maliciously crafted source array could wrap
  // and request a tiny capacity, causing the subsequent push_backs to
  // grow unexpectedly. Surface as `#NUM!` if the upper bound itself
  // cannot be represented.
  auto reserve_or = checked_mul_size_t(array->rows, array->cols);
  if (!reserve_or) {
    return Value::error(ErrorCode::Num);
  }
  std::vector<Value> kept;
  kept.reserve(reserve_or.value());
  if (!collect_tocol_torow_cells(*array, ignore_mask, scan_by_column, kept, err)) {
    return err;
  }
  if (kept.empty()) {
    return Value::error(ErrorCode::Calc);
  }
  ArrayValue* out = dynamic_array::materialise_vector(std::move(kept), as_column, arena);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  return Value::array(out);
}

/// WRAPROWS / WRAPCOLS. `by_row` decides which axis the wrap count sizes and,
/// with it, whether the source vector is read straight through (row-major fill)
/// or transposed as it is written (column-major fill).
Value eval_wrap(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                bool by_row) {
  const std::uint32_t arity = call.as_call_arity();
  const ArrayValue* vector_arr = nullptr;
  std::uint32_t wrap_count = 0;
  Value pad = Value::error(ErrorCode::NA);
  Value err = Value::error(ErrorCode::Value);
  if (!resolve_wrap_args(call, arity, arena, registry, ctx, vector_arr, wrap_count, pad, err)) {
    return err;
  }
  // The vector's natural flat order is row-major, which is what we want
  // both for a row vector (1xN) and a column vector (Nx1) — both have
  // their cells already laid out left-to-right / top-to-bottom in
  // `cells` (since rows or cols == 1). Total cell count is rows * cols.
  const std::size_t count = static_cast<std::size_t>(vector_arr->rows) * static_cast<std::size_t>(vector_arr->cols);
  // Excel does not extend the final shape beyond the source length when the
  // requested wrap count exceeds the vector length.
  const std::uint32_t effective_wrap =
      static_cast<std::uint32_t>(std::min<std::size_t>(static_cast<std::size_t>(wrap_count), count));
  const auto wrapped =
      static_cast<std::uint32_t>((count + static_cast<std::size_t>(effective_wrap) - 1) / effective_wrap);
  const std::uint32_t out_rows = by_row ? wrapped : effective_wrap;
  const std::uint32_t out_cols = by_row ? effective_wrap : wrapped;
  const std::size_t total = static_cast<std::size_t>(out_rows) * static_cast<std::size_t>(out_cols);
  Value* buffer = nullptr;
  ArrayValue* out = dynamic_array::allocate_array_value(out_rows, out_cols, arena, buffer, kMaxDerivedArrayCells);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  for (std::size_t i = 0; i < total; ++i) {
    // WRAPCOLS fills down each column first, so the source index walks the
    // output in column-major order.
    const std::size_t flat = by_row ? i : ((i % out_cols) * out_rows) + (i / out_cols);
    buffer[i] = (flat < count) ? vector_arr->cells[flat] : pad;
  }
  return Value::array(out);
}

}  // namespace

Value eval_tocol_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx) {
  return eval_to_vector(call, arena, registry, ctx, /*as_column=*/true);
}

Value eval_torow_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx) {
  return eval_to_vector(call, arena, registry, ctx, /*as_column=*/false);
}

Value eval_wraprows_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx) {
  return eval_wrap(call, arena, registry, ctx, /*by_row=*/true);
}

Value eval_wrapcols_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx) {
  return eval_wrap(call, arena, registry, ctx, /*by_row=*/false);
}

}  // namespace eval
}  // namespace formulon
