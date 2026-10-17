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
#include <utility>
#include <vector>

#include "eval/array_alloc.h"
#include "eval/builtins/numeric_helpers.h"
#include "eval/coerce.h"
#include "eval/lazy_impls.h"
#include "eval/range_args.h"
#include "eval/shape_ops_lazy.h"
#include "parser/ast.h"
#include "utils/a1_ref.h"
#include "utils/arena.h"
#include "utils/checked_mul.h"
#include "utils/error.h"
#include "utils/subnormal_flush.h"
#include "value.h"

namespace formulon {
namespace eval {

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
    // Mac Excel's matrix functions do not coerce Text-numeric strings ("3" -> 3).
    auto cell = matrix_strict_number(src.cells[i]);
    if (!cell) {
      out_err = Value::error(cell.error());
      return false;
    }
    out[i] = cell.value();
  }
  return true;
}

namespace {

using builtins_detail::to_finite_value;

/// Builds an arena-allocated `ArrayValue` from a flat row-major double
/// buffer + shape. Returns `nullptr` on arena OOM.
ArrayValue* make_double_array(const std::vector<double>& data, std::uint32_t rows, std::uint32_t cols, Arena& arena) {
  // Same overflow-defensive guard as `coerce_matrix`: on 32-bit `size_t`
  // an attacker-controlled `rows * cols >= 2^32` would wrap and request a
  // truncated arena buffer. Bail to nullptr; callers map that to `#NUM!`.
  Value* buffer = nullptr;
  ArrayValue* arr = allocate_array_value(rows, cols, arena, buffer, kMaxDerivedArrayCells);
  if (arr == nullptr) {
    return nullptr;
  }
  const std::size_t n = static_cast<std::size_t>(rows) * static_cast<std::size_t>(cols);
  for (std::size_t i = 0; i < n; ++i) {
    buffer[i] = Value::number(data[i]);
  }
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
  if (!resolve_array_value(call.as_call_arg(0), arena, registry, ctx, &a, &err)) {
    return err;
  }
  if (!resolve_array_value(call.as_call_arg(1), arena, registry, ctx, &b, &err)) {
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
        s += flush_subnormal(ad[static_cast<std::size_t>(i) * inner + k] *
                             bd[static_cast<std::size_t>(k) * out_cols + j]);
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

// Resolves the single square-matrix argument MDETERM and MINVERSE both
// take, into a row-major `n x n` buffer. A non-square or degenerate
// (0x0) input is `#VALUE!`, which is what Mac Excel surfaces for both.
//
// Returns `false` with the propagating error in `*out_err`.
bool resolve_square_matrix(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx, std::vector<double>* out_matrix, std::uint32_t* out_n,
                           Value* out_err) {
  if (call.as_call_arity() != 1U) {
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  const ArrayValue* arr = nullptr;
  if (!resolve_array_value(call.as_call_arg(0), arena, registry, ctx, &arr, out_err)) {
    return false;
  }
  if (arr->rows != arr->cols) {
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  if (!coerce_matrix(*arr, *out_matrix, *out_err)) {
    return false;
  }
  // A 0x0 matrix can only arise via `eval_node_as_array` returning a 1x1
  // wrapper; `rows` is always >= 1 here. Guard explicitly anyway so a
  // degenerate input is `#VALUE!` rather than an empty decomposition.
  if (arr->rows == 0U) {
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  *out_n = arr->rows;
  return true;
}

}  // namespace

Value eval_mdeterm_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx) {
  std::vector<double> m;
  std::uint32_t n = 0;
  Value err = Value::blank();
  if (!resolve_square_matrix(call, arena, registry, ctx, &m, &n, &err)) {
    return err;
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
  return to_finite_value(det);
}

Value eval_minverse_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx) {
  std::vector<double> m;
  std::uint32_t n = 0;
  Value err = Value::blank();
  if (!resolve_square_matrix(call, arena, registry, ctx, &m, &n, &err)) {
    return err;
  }
  // Hard cap to keep downstream `uint32_t` arithmetic from wrapping: the loop
  // indices over `n * n` need to stay well below `UINT32_MAX`. Excel 365 itself caps
  // worksheet width at 16384 columns, so any matrix the engine could
  // legitimately encounter from a workbook is comfortably under this
  // limit; values above it can only come from synthetic inputs.
  if (n > a1::kMaxCols) {
    return Value::error(ErrorCode::Num);
  }

  // LU decomposition with partial pivoting (the first row holding the largest
  // |a[i][k]| becomes the pivot row; whole rows are swapped), then one
  // permuted unit column at a time is solved by row-oriented forward and
  // backward substitution, accumulating `acc - a * y` for k ascending and
  // dividing by the diagonal last. A zero pivot is singular (#NUM!).
  const std::size_t un = n;
  std::vector<double> a = m;
  std::vector<std::uint32_t> source_row(n);
  for (std::uint32_t r = 0; r < n; ++r) {
    source_row[r] = r;
  }
  for (std::uint32_t k = 0; k < n; ++k) {
    std::uint32_t pivot = k;
    double pivot_abs = std::fabs(a[k * un + k]);
    for (std::uint32_t r = k + 1U; r < n; ++r) {
      const double v = std::fabs(a[r * un + k]);
      if (v > pivot_abs) {
        pivot_abs = v;
        pivot = r;
      }
    }
    if (pivot != k) {
      for (std::uint32_t c = 0; c < n; ++c) {
        std::swap(a[k * un + c], a[pivot * un + c]);
      }
      std::swap(source_row[k], source_row[pivot]);
    }
    const double diagonal = a[k * un + k];
    if (diagonal == 0.0) {
      return Value::error(ErrorCode::Num);
    }
    for (std::uint32_t i = k + 1U; i < n; ++i) {
      a[i * un + k] = a[i * un + k] / diagonal;
      const double factor = a[i * un + k];
      for (std::uint32_t j = k + 1U; j < n; ++j) {
        const double product = factor * a[k * un + j];
        a[i * un + j] = a[i * un + j] - product;
      }
    }
  }
  std::vector<double> inverse(un * un, 0.0);
  std::vector<double> y(un);
  for (std::uint32_t col = 0; col < n; ++col) {
    for (std::uint32_t i = 0; i < n; ++i) {
      double acc = source_row[i] == col ? 1.0 : 0.0;
      for (std::uint32_t k = 0; k < i; ++k) {
        const double product = a[i * un + k] * y[k];
        acc = acc - product;
      }
      y[i] = acc;
    }
    for (std::uint32_t i = n; i-- > 0;) {
      double acc = y[i];
      for (std::uint32_t k = i + 1U; k < n; ++k) {
        const double product = a[i * un + k] * inverse[k * un + col];
        acc = acc - product;
      }
      inverse[i * un + col] = acc / a[i * un + i];
    }
  }

  std::vector<double> out = std::move(inverse);
  for (const double v : out) {
    if (std::isnan(v) || std::isinf(v)) {
      return Value::error(ErrorCode::Num);
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
