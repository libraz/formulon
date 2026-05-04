// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Linear-algebra builtins (MMULT, MDETERM, MINVERSE). See matrix_ops_lazy.h
// for argument and error semantics. Each function:
//   1. Materialises its array argument(s) via `eval_node_as_array` (the same
//      shape-preserving seam used by SUMPRODUCT and the dynamic-array
//      family). Range / Ref / ArrayLiteral all keep their 2D rectangle;
//      scalar args become (1, 1).
//   2. Coerces every cell to `double` via `coerce_to_number_strict`, which
//      rejects Text / Blank with `#VALUE!` per Mac Excel's matrix-cell
//      contract. Errors short-circuit in row-major scan order.
//   3. Runs the hand-rolled algorithm (triple loop / Gauss elimination /
//      Gauss-Jordan) on a flat `std::vector<double>` working buffer.
//   4. Materialises the result into the caller's arena (for MMULT /
//      MINVERSE) or returns it as a scalar `Value::number` (for MDETERM).

#include "eval/matrix_ops_lazy.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <vector>

#include "eval/coerce.h"
#include "eval/lazy_impls.h"
#include "eval/shape_ops_lazy.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/checked_mul.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

/// Materialises a single array argument as an `ArrayValue` (scalars wrap to
/// 1x1). Returns `true` on success; on failure writes the propagating error
/// into `out_err`. The returned `ArrayValue` borrows from the caller arena.
bool resolve_matrix_arg(const parser::AstNode& arg, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx, const ArrayValue** out, Value* out_err) {
  const Value v = eval_node_as_array(arg, arena, registry, ctx);
  if (v.is_error()) {
    *out_err = v;
    return false;
  }
  if (!v.is_array()) {
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  *out = v.as_array();
  return true;
}

/// Per-cell numeric coercion. MMULT / MDETERM / MINVERSE require every
/// matrix cell to be a `Number` (Booleans coerce to 1/0 via the standard
/// dispatcher rule; Blank / Text / Error never become a numeric matrix
/// cell). On the first non-numeric or error cell, returns `false` and
/// writes the propagating error into `out_err`.
bool coerce_matrix(const ArrayValue& src, std::vector<double>& out, Value& out_err) {
  // Defensive overflow check: on 32-bit `size_t` (WASM) a maliciously
  // crafted ArrayValue with `rows * cols >= 2^32` would silently wrap and
  // leave `out` shorter than the loop expects, producing an OOB read on
  // `src.cells[i]`. Surface the overflow as `#NUM!` instead.
  auto n_or = checked_mul_size_t(src.rows, src.cols);
  if (!n_or) {
    out_err = Value::error(ErrorCode::Num);
    return false;
  }
  const std::size_t n = n_or.value();
  out.resize(n);
  for (std::size_t i = 0; i < n; ++i) {
    const Value& cell = src.cells[i];
    if (cell.is_error()) {
      out_err = cell;
      return false;
    }
    if (cell.is_number()) {
      out[i] = cell.as_number();
      continue;
    }
    if (cell.is_boolean()) {
      out[i] = cell.as_boolean() ? 1.0 : 0.0;
      continue;
    }
    // Blank / Text / Lambda / Ref / Array — all rejected. Mac Excel's
    // matrix functions do not coerce Text-numeric strings ("3" -> 3); a
    // bare Text cell anywhere in the matrix surfaces `#VALUE!`.
    out_err = Value::error(ErrorCode::Value);
    return false;
  }
  return true;
}

/// Builds an arena-allocated `ArrayValue` from a flat row-major double
/// buffer + shape. Returns `nullptr` on arena OOM.
ArrayValue* make_double_array(const std::vector<double>& data, std::uint32_t rows, std::uint32_t cols, Arena& arena) {
  // Same overflow-defensive guard as `coerce_matrix`: on 32-bit `size_t`
  // an attacker-controlled `rows * cols >= 2^32` would wrap and request a
  // truncated arena buffer. Bail to nullptr; callers map that to `#NUM!`.
  auto n_or = checked_mul_size_t(rows, cols);
  if (!n_or) {
    return nullptr;
  }
  const std::size_t n = n_or.value();
  Value* buffer = arena.create_array<Value>(n);
  if (buffer == nullptr) {
    return nullptr;
  }
  for (std::size_t i = 0; i < n; ++i) {
    buffer[i] = Value::number(data[i]);
  }
  ArrayValue* arr = arena.create<ArrayValue>();
  if (arr == nullptr) {
    return nullptr;
  }
  arr->rows = rows;
  arr->cols = cols;
  arr->cells = buffer;
  return arr;
}

}  // namespace

Value eval_mmult_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx) {
  if (call.as_call_arity() != 2U) {
    return Value::error(ErrorCode::Value);
  }

  const ArrayValue* a = nullptr;
  const ArrayValue* b = nullptr;
  Value err = Value::blank();
  if (!resolve_matrix_arg(call.as_call_arg(0), arena, registry, ctx, &a, &err)) {
    return err;
  }
  if (!resolve_matrix_arg(call.as_call_arg(1), arena, registry, ctx, &b, &err)) {
    return err;
  }

  // Inner-dimension check: cols(a) must equal rows(b). Mismatch is
  // `#VALUE!`, not a per-cell error — Mac Excel collapses this to scalar.
  if (a->cols != b->rows) {
    return Value::error(ErrorCode::Value);
  }

  std::vector<double> ad;
  std::vector<double> bd;
  if (!coerce_matrix(*a, ad, err)) {
    return err;
  }
  if (!coerce_matrix(*b, bd, err)) {
    return err;
  }

  const std::uint32_t out_rows = a->rows;
  const std::uint32_t out_cols = b->cols;
  const std::uint32_t inner = a->cols;
  // Defensive overflow check on the result allocation; same rationale as
  // `coerce_matrix`. WASM 32-bit `size_t` can wrap on `out_rows * out_cols`
  // when both operands are pathologically wide.
  auto out_n_or = checked_mul_size_t(out_rows, out_cols);
  if (!out_n_or) {
    return Value::error(ErrorCode::Num);
  }
  std::vector<double> out(out_n_or.value(), 0.0);
  for (std::uint32_t i = 0; i < out_rows; ++i) {
    for (std::uint32_t j = 0; j < out_cols; ++j) {
      double s = 0.0;
      for (std::uint32_t k = 0; k < inner; ++k) {
        s += ad[static_cast<std::size_t>(i) * inner + k] * bd[static_cast<std::size_t>(k) * out_cols + j];
      }
      if (std::isnan(s) || std::isinf(s)) {
        return Value::error(ErrorCode::Num);
      }
      out[static_cast<std::size_t>(i) * out_cols + j] = s;
    }
  }

  ArrayValue* arr = make_double_array(out, out_rows, out_cols, arena);
  if (arr == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  return Value::array(arr);
}

namespace {

/// In-place LU decomposition with partial pivoting. `m` is a row-major
/// `n x n` buffer; on return `m` holds L (below diagonal) and U (on/above)
/// merged in the standard packed form. `*sign` is multiplied by `-1` for
/// each row swap. Returns `true` on success; `false` if the matrix is
/// numerically singular (zero pivot encountered) — in which case `m` is
/// in an indeterminate state and the caller should treat the result as
/// the singular case (determinant = 0; inverse = `#NUM!`).
bool lu_decompose(std::vector<double>& m, std::uint32_t n, double* sign) {
  for (std::uint32_t k = 0; k < n; ++k) {
    // Partial pivot: find the row in column k (at or below the diagonal)
    // with the largest absolute value, and swap it into row k. This is
    // the standard stability fix for naive Gaussian elimination.
    std::uint32_t pivot = k;
    double pivot_abs = std::fabs(m[static_cast<std::size_t>(k) * n + k]);
    for (std::uint32_t r = k + 1U; r < n; ++r) {
      const double v = std::fabs(m[static_cast<std::size_t>(r) * n + k]);
      if (v > pivot_abs) {
        pivot_abs = v;
        pivot = r;
      }
    }
    if (pivot_abs == 0.0) {
      return false;
    }
    if (pivot != k) {
      for (std::uint32_t c = 0; c < n; ++c) {
        const std::size_t a = static_cast<std::size_t>(k) * n + c;
        const std::size_t b = static_cast<std::size_t>(pivot) * n + c;
        const double tmp = m[a];
        m[a] = m[b];
        m[b] = tmp;
      }
      *sign = -*sign;
    }
    // Eliminate column k below the diagonal. The multiplier `f` is stored
    // in m[r,k] (the L-half of the LU layout); the U-half overwrite
    // continues on the right.
    const double diag = m[static_cast<std::size_t>(k) * n + k];
    for (std::uint32_t r = k + 1U; r < n; ++r) {
      const double f = m[static_cast<std::size_t>(r) * n + k] / diag;
      m[static_cast<std::size_t>(r) * n + k] = f;
      for (std::uint32_t c = k + 1U; c < n; ++c) {
        m[static_cast<std::size_t>(r) * n + c] -= f * m[static_cast<std::size_t>(k) * n + c];
      }
    }
  }
  return true;
}

}  // namespace

Value eval_mdeterm_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx) {
  if (call.as_call_arity() != 1U) {
    return Value::error(ErrorCode::Value);
  }
  const ArrayValue* a = nullptr;
  Value err = Value::blank();
  if (!resolve_matrix_arg(call.as_call_arg(0), arena, registry, ctx, &a, &err)) {
    return err;
  }
  if (a->rows != a->cols) {
    return Value::error(ErrorCode::Value);
  }
  std::vector<double> m;
  if (!coerce_matrix(*a, m, err)) {
    return err;
  }
  const std::uint32_t n = a->rows;
  // 0x0 matrix can only arise via `eval_node_as_array` returning a 1x1
  // wrapper; n is always >= 1 here. Nonetheless guard explicitly so the
  // determinant of a degenerate input is `#VALUE!` (Mac Excel surfaces
  // the same code for an empty array).
  if (n == 0U) {
    return Value::error(ErrorCode::Value);
  }

  double sign = 1.0;
  if (!lu_decompose(m, n, &sign)) {
    // Singular matrix -> determinant is exactly 0. (Mac Excel returns
    // 0 here rather than #NUM!.)
    return Value::number(0.0);
  }
  // Determinant = sign * product of U's diagonal.
  double det = sign;
  for (std::uint32_t i = 0; i < n; ++i) {
    det *= m[static_cast<std::size_t>(i) * n + i];
  }
  if (std::isnan(det) || std::isinf(det)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(det);
}

Value eval_minverse_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx) {
  if (call.as_call_arity() != 1U) {
    return Value::error(ErrorCode::Value);
  }
  const ArrayValue* a = nullptr;
  Value err = Value::blank();
  if (!resolve_matrix_arg(call.as_call_arg(0), arena, registry, ctx, &a, &err)) {
    return err;
  }
  if (a->rows != a->cols) {
    return Value::error(ErrorCode::Value);
  }
  std::vector<double> m;
  if (!coerce_matrix(*a, m, err)) {
    return err;
  }
  const std::uint32_t n = a->rows;
  if (n == 0U) {
    return Value::error(ErrorCode::Value);
  }
  // Hard cap to keep downstream `uint32_t` arithmetic from wrapping. The
  // augmented system is `n x 2n`, so `2U * n` and the loop indices over
  // `n * w` need to stay well below `UINT32_MAX`. Excel 365 itself caps
  // worksheet width at 16384 columns, so any matrix the engine could
  // legitimately encounter from a workbook is comfortably under this
  // limit; values above it can only come from synthetic inputs.
  constexpr std::uint32_t kMaxMinverseDimension = 16384U;
  if (n > kMaxMinverseDimension) {
    return Value::error(ErrorCode::Num);
  }

  // Gauss-Jordan elimination on the augmented `[m | I]` system. We work
  // on a single 2n-wide buffer so the inverse falls out in the right
  // half once the left half has been reduced to the identity. Partial
  // pivoting keeps the algorithm stable for ill-conditioned but
  // non-singular inputs.
  const std::uint32_t w = 2U * n;
  std::vector<double> aug(static_cast<std::size_t>(n) * w, 0.0);
  for (std::uint32_t r = 0; r < n; ++r) {
    for (std::uint32_t c = 0; c < n; ++c) {
      aug[static_cast<std::size_t>(r) * w + c] = m[static_cast<std::size_t>(r) * n + c];
    }
    aug[static_cast<std::size_t>(r) * w + (n + r)] = 1.0;
  }

  for (std::uint32_t k = 0; k < n; ++k) {
    // Partial pivot (same logic as `lu_decompose`).
    std::uint32_t pivot = k;
    double pivot_abs = std::fabs(aug[static_cast<std::size_t>(k) * w + k]);
    for (std::uint32_t r = k + 1U; r < n; ++r) {
      const double v = std::fabs(aug[static_cast<std::size_t>(r) * w + k]);
      if (v > pivot_abs) {
        pivot_abs = v;
        pivot = r;
      }
    }
    if (pivot_abs == 0.0) {
      return Value::error(ErrorCode::Num);
    }
    if (pivot != k) {
      for (std::uint32_t c = 0; c < w; ++c) {
        const std::size_t a_idx = static_cast<std::size_t>(k) * w + c;
        const std::size_t b_idx = static_cast<std::size_t>(pivot) * w + c;
        const double tmp = aug[a_idx];
        aug[a_idx] = aug[b_idx];
        aug[b_idx] = tmp;
      }
    }
    // Normalise the pivot row so the diagonal entry becomes 1.
    const double diag = aug[static_cast<std::size_t>(k) * w + k];
    for (std::uint32_t c = 0; c < w; ++c) {
      aug[static_cast<std::size_t>(k) * w + c] /= diag;
    }
    // Eliminate column `k` in every other row.
    for (std::uint32_t r = 0; r < n; ++r) {
      if (r == k) {
        continue;
      }
      const double f = aug[static_cast<std::size_t>(r) * w + k];
      if (f == 0.0) {
        continue;
      }
      for (std::uint32_t c = 0; c < w; ++c) {
        aug[static_cast<std::size_t>(r) * w + c] -= f * aug[static_cast<std::size_t>(k) * w + c];
      }
    }
  }

  std::vector<double> out(static_cast<std::size_t>(n) * static_cast<std::size_t>(n), 0.0);
  for (std::uint32_t r = 0; r < n; ++r) {
    for (std::uint32_t c = 0; c < n; ++c) {
      const double v = aug[static_cast<std::size_t>(r) * w + (n + c)];
      if (std::isnan(v) || std::isinf(v)) {
        return Value::error(ErrorCode::Num);
      }
      out[static_cast<std::size_t>(r) * n + c] = v;
    }
  }
  ArrayValue* arr = make_double_array(out, n, n, arena);
  if (arr == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  return Value::array(arr);
}

}  // namespace eval
}  // namespace formulon
