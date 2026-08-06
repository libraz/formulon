//
// LINEST implementation. See linest_lazy.h for the user-facing
// contract; this file owns the numerical recipe (build the design
// matrix, form the normal equations, solve via Gauss-Jordan with
// partial pivoting, optionally derive the regression statistics).
//
// The matrix kernel is intentionally hand-rolled — same `~150 line`
// budget as MINVERSE — to keep the WASM size policy intact (no Eigen,
// no LAPACK).

#include "eval/linest_lazy.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <vector>

#include "eval/array_alloc.h"
#include "eval/lazy_impls.h"
#include "eval/omitted_arg.h"
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

/// Strict numeric coercion for LINEST inputs. Numbers pass through;
/// Booleans coerce to 1/0; Blank/Text/Error fail. Errors propagate
/// verbatim via `*out_err` (and `false`); non-numeric non-error cells
/// surface as `#VALUE!`. Mac Excel's matrix-strict rule.
bool coerce_cell(const Value& cell, double* out, Value* out_err) {
  if (cell.is_error()) {
    *out_err = cell;
    return false;
  }
  if (cell.is_number()) {
    *out = cell.as_number();
    return true;
  }
  if (cell.is_boolean()) {
    *out = cell.as_boolean() ? 1.0 : 0.0;
    return true;
  }
  *out_err = Value::error(ErrorCode::Value);
  return false;
}

/// Evaluates a flat-cell scan over an entire `ArrayValue`, populating
/// `out` with the coerced doubles in row-major order. Returns `false`
/// on the first non-numeric / error cell.
bool coerce_array(const ArrayValue& src, std::vector<double>& out, Value* out_err) {
  // Defensive overflow guard: on 32-bit `size_t` (WASM) a malformed
  // ArrayValue with `rows * cols >= 2^32` would silently wrap the
  // multiply, leaving `out` shorter than the loop's iteration bound and
  // producing OOB reads on `src.cells[i]`. Surface as `#NUM!` instead.
  auto n_or = checked_mul_size_t(src.rows, src.cols);
  if (!n_or) {
    *out_err = Value::error(ErrorCode::Num);
    return false;
  }
  const std::size_t n = n_or.value();
  out.resize(n);
  for (std::size_t i = 0; i < n; ++i) {
    if (!coerce_cell(src.cells[i], &out[i], out_err)) {
      return false;
    }
  }
  return true;
}

/// Flat row-major buffer + shape returned by `make_double_array`.
ArrayValue* make_double_array(const std::vector<Value>& data, std::uint32_t rows, std::uint32_t cols, Arena& arena) {
  // Same overflow-defensive guard as `coerce_array`: bail on 32-bit
  // wrap so the arena allocation request matches what the loop expects.
  Value* buffer = nullptr;
  ArrayValue* arr = allocate_array_value(rows, cols, arena, buffer, kMaxDerivedArrayCells);
  if (arr == nullptr) {
    return nullptr;
  }
  const std::size_t n = static_cast<std::size_t>(rows) * static_cast<std::size_t>(cols);
  for (std::size_t i = 0; i < n; ++i) {
    buffer[i] = data[i];
  }
  return arr;
}

/// Returns true if `v` is finite, false for NaN / +-inf. Used to gate
/// final result cells; LINEST surfaces `#NUM!` for any non-finite
/// statistic.
bool is_finite(double v) noexcept {
  return !std::isnan(v) && !std::isinf(v);
}

/// Evaluates an optional scalar Boolean argument. Numbers coerce
/// to false (zero) / true (non-zero); Booleans pass through; errors
/// propagate via `*out_err` (returning `false`); other types -> `#VALUE!`.
bool eval_bool_arg(const parser::AstNode& arg, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                   bool default_value, bool* out, Value* out_err) {
  const Value v = eval_node(arg, arena, registry, ctx);
  if (v.is_error()) {
    *out_err = v;
    return false;
  }
  if (v.is_blank()) {
    *out = default_value;
    return true;
  }
  if (v.is_boolean()) {
    *out = v.as_boolean();
    return true;
  }
  if (v.is_number()) {
    *out = v.as_number() != 0.0;
    return true;
  }
  *out_err = Value::error(ErrorCode::Value);
  return false;
}

/// Solves `A * X = B` in-place via rank-aware Gauss-Jordan with partial
/// pivoting. `aug` is a row-major `n x w` buffer with `w = n + extra`
/// and the augmented columns laid out to the right of `A`. On return,
/// the non-redundant pivot rows are normalised and back-eliminated; rows
/// corresponding to redundant (rank-deficient) columns are zeroed.
///
/// Rank-deficiency policy: at step `k`, partial pivoting picks the
/// largest |entry| in column `k` from rows `k..n-1`. If that magnitude
/// is below `kPivotEpsilon = 1e-12 * max|A_initial|`, column `k` is
/// declared redundant — row `k` is zeroed (so its b-slot becomes 0 and
/// its inverse-slot row becomes 0), no swap / divide / elimination is
/// performed for that step, and the loop continues. The non-zero
/// off-diagonal entries that remain in earlier rows do not affect the
/// solution because the corresponding `x_k` is forced to 0.
///
/// Always returns `true`; callers do not need to handle a failure path.
/// Mirrors Mac Excel 365's behaviour of returning a partial fit on
/// rank-deficient systems rather than `#NUM!`.
bool gauss_jordan(std::vector<double>& aug, std::uint32_t n, std::uint32_t w) {
  // Establish the rank-deficiency tolerance from the initial scale of A
  // (the left n columns of aug). A relative tolerance keeps the test
  // sensible across very different problem magnitudes; an absolute
  // 1e-12 floor avoids treating a genuinely zero matrix as non-trivial.
  double max_abs = 0.0;
  for (std::uint32_t r = 0; r < n; ++r) {
    for (std::uint32_t c = 0; c < n; ++c) {
      const double v = std::fabs(aug[static_cast<std::size_t>(r) * w + c]);
      if (v > max_abs) {
        max_abs = v;
      }
    }
  }
  const double kPivotEpsilon = 1e-12 * (max_abs > 0.0 ? max_abs : 1.0);

  for (std::uint32_t k = 0; k < n; ++k) {
    std::uint32_t pivot = k;
    double pivot_abs = std::fabs(aug[static_cast<std::size_t>(k) * w + k]);
    for (std::uint32_t r = k + 1U; r < n; ++r) {
      const double v = std::fabs(aug[static_cast<std::size_t>(r) * w + k]);
      if (v > pivot_abs) {
        pivot_abs = v;
        pivot = r;
      }
    }
    if (pivot_abs < kPivotEpsilon) {
      // Column k is redundant. Zero out row k entirely so the b-slot
      // becomes 0 (coefficient = 0 in the unpacked solution) and the
      // inverse-slot row becomes 0 (SE = 0 for that coefficient).
      for (std::uint32_t c = 0; c < w; ++c) {
        aug[static_cast<std::size_t>(k) * w + c] = 0.0;
      }
      continue;
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
    const double diag = aug[static_cast<std::size_t>(k) * w + k];
    for (std::uint32_t c = 0; c < w; ++c) {
      aug[static_cast<std::size_t>(k) * w + c] /= diag;
    }
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
  return true;
}

/// One observation row in the design matrix view: the `k` predictor
/// values plus an implicit constant `1` (when `with_const`). The
/// caller indexes `m` rows of these.
struct DesignMatrix {
  std::uint32_t m;           // observations
  std::uint32_t k;           // predictors (X columns)
  bool with_const;           // include intercept column
  std::vector<double> data;  // row-major, m x p (p = k or k+1)
  std::vector<double> y;     // length m
};

struct FitStats {
  double ss_resid = 0.0;
  double ss_total = 0.0;
  double mean_y = 0.0;
};

/// Forms `A = X^T X` (p x p) and `b = X^T y` (length p) given the
/// design matrix view `dm`. `p` = `k + (with_const ? 1 : 0)`.
void normal_equations(const DesignMatrix& dm, std::vector<double>& a, std::vector<double>& b) {
  const std::uint32_t p = dm.k + (dm.with_const ? 1U : 0U);
  const std::uint32_t m = dm.m;
  a.assign(static_cast<std::size_t>(p) * p, 0.0);
  b.assign(p, 0.0);
  for (std::uint32_t i = 0; i < m; ++i) {
    for (std::uint32_t r = 0; r < p; ++r) {
      const double xri = dm.data[static_cast<std::size_t>(i) * p + r];
      b[r] += xri * dm.y[i];
      for (std::uint32_t c = 0; c < p; ++c) {
        a[static_cast<std::size_t>(r) * p + c] += xri * dm.data[static_cast<std::size_t>(i) * p + c];
      }
    }
  }
}

/// End-to-end coefficient solve. Forms the normal equations
/// `A * beta = b`, then runs the rank-aware Gauss-Jordan kernel on
/// `[A | b | I]` (when `want_inv=true`) or `[A | b]` (otherwise).
/// Writes `*coeffs` (length `p`) and — when requested — `*inv_a`
/// (`p x p` row-major).
///
/// Intercept-first permutation: when `dm.with_const` is true the
/// design-matrix layout puts the intercept column at index `k` (last).
/// To make the rank-aware kernel's left-to-right "first surviving
/// column wins" rule match Mac Excel 365's behaviour on rank-deficient
/// X — where the intercept absorbs `mean(y)` and collinear predictors
/// drop to 0 — we swap row/column 0 with row/column k of A (and entry
/// 0 with entry k of b) before solving, then reverse the swap on
/// `coeffs` and on the rows + columns of `inv_a` afterwards. For
/// non-singular systems the permutation is a no-op (matrix swaps are
/// invertible), so this only changes behaviour at the rank-deficient
/// boundary.
///
/// Always returns `true` — `gauss_jordan` is rank-aware and never
/// fails. The bool return is preserved for call-site stability.
bool solve_normal_equations(const DesignMatrix& dm, bool want_inv, std::vector<double>* coeffs,
                            std::vector<double>* inv_a) {
  const std::uint32_t p = dm.k + (dm.with_const ? 1U : 0U);
  std::vector<double> a;
  std::vector<double> b;
  normal_equations(dm, a, b);

  // Intercept-first permutation: swap rows 0 and k of A (symmetric, so
  // swap matching columns too) and entries 0 and k of b. Only meaningful
  // when an intercept is present and there is at least one predictor.
  const bool permute = dm.with_const && dm.k >= 1U;
  const std::uint32_t k_idx = dm.k;  // intercept column index in design.
  if (permute) {
    for (std::uint32_t c = 0; c < p; ++c) {
      const std::size_t i0 = static_cast<std::size_t>(0) * p + c;
      const std::size_t ik = static_cast<std::size_t>(k_idx) * p + c;
      const double tmp = a[i0];
      a[i0] = a[ik];
      a[ik] = tmp;
    }
    for (std::uint32_t r = 0; r < p; ++r) {
      const std::size_t i0 = static_cast<std::size_t>(r) * p + 0U;
      const std::size_t ik = static_cast<std::size_t>(r) * p + k_idx;
      const double tmp = a[i0];
      a[i0] = a[ik];
      a[ik] = tmp;
    }
    const double tmp_b = b[0];
    b[0] = b[k_idx];
    b[k_idx] = tmp_b;
  }

  const std::uint32_t w = p + 1U + (want_inv ? p : 0U);
  std::vector<double> aug(static_cast<std::size_t>(p) * w, 0.0);
  for (std::uint32_t r = 0; r < p; ++r) {
    for (std::uint32_t c = 0; c < p; ++c) {
      aug[static_cast<std::size_t>(r) * w + c] = a[static_cast<std::size_t>(r) * p + c];
    }
    aug[static_cast<std::size_t>(r) * w + p] = b[r];
    if (want_inv) {
      aug[static_cast<std::size_t>(r) * w + (p + 1U + r)] = 1.0;
    }
  }

  (void)gauss_jordan(aug, p, w);

  coeffs->assign(p, 0.0);
  for (std::uint32_t r = 0; r < p; ++r) {
    (*coeffs)[r] = aug[static_cast<std::size_t>(r) * w + p];
  }
  if (want_inv) {
    inv_a->assign(static_cast<std::size_t>(p) * p, 0.0);
    for (std::uint32_t r = 0; r < p; ++r) {
      for (std::uint32_t c = 0; c < p; ++c) {
        (*inv_a)[static_cast<std::size_t>(r) * p + c] = aug[static_cast<std::size_t>(r) * w + (p + 1U + c)];
      }
    }
  }

  // Unpermute: undo the (0 <-> k_idx) swap on coeffs and on rows + cols
  // of inv_a so the caller sees them in the original predictor-first
  // layout.
  if (permute) {
    const double tmp_c = (*coeffs)[0];
    (*coeffs)[0] = (*coeffs)[k_idx];
    (*coeffs)[k_idx] = tmp_c;
    if (want_inv) {
      for (std::uint32_t c = 0; c < p; ++c) {
        const std::size_t i0 = static_cast<std::size_t>(0) * p + c;
        const std::size_t ik = static_cast<std::size_t>(k_idx) * p + c;
        const double tmp = (*inv_a)[i0];
        (*inv_a)[i0] = (*inv_a)[ik];
        (*inv_a)[ik] = tmp;
      }
      for (std::uint32_t r = 0; r < p; ++r) {
        const std::size_t i0 = static_cast<std::size_t>(r) * p + 0U;
        const std::size_t ik = static_cast<std::size_t>(r) * p + k_idx;
        const double tmp = (*inv_a)[i0];
        (*inv_a)[i0] = (*inv_a)[ik];
        (*inv_a)[ik] = tmp;
      }
    }
  }
  return true;
}

FitStats compute_fit_stats(const DesignMatrix& dm, const std::vector<double>& coeffs) {
  const std::uint32_t p = dm.k + (dm.with_const ? 1U : 0U);
  FitStats stats;
  double sum_y = 0.0;
  for (std::uint32_t i = 0; i < dm.m; ++i) {
    double yhat = 0.0;
    for (std::uint32_t j = 0; j < p; ++j) {
      yhat += coeffs[j] * dm.data[static_cast<std::size_t>(i) * p + j];
    }
    const double e = dm.y[i] - yhat;
    stats.ss_resid += e * e;
    sum_y += dm.y[i];
  }
  stats.mean_y = sum_y / static_cast<double>(dm.m);

  if (dm.with_const) {
    for (std::uint32_t i = 0; i < dm.m; ++i) {
      const double d = dm.y[i] - stats.mean_y;
      stats.ss_total += d * d;
    }
  } else {
    for (std::uint32_t i = 0; i < dm.m; ++i) {
      stats.ss_total += dm.y[i] * dm.y[i];
    }
  }
  return stats;
}

/// Builds the design matrix from a known_y array and an optional
/// known_x array. `known_y` may be 1 x m or m x 1; the orientation
/// determines how `known_x` is read. If `x_arr` is `nullptr`, the
/// predictor defaults to the index sequence `{1, 2, ..., m}`.
///
/// Returns `false` on shape mismatch (writes `#REF!` / `#VALUE!` into
/// `*out_err`) or on a non-numeric / error cell (propagates verbatim).
bool build_design_matrix(const ArrayValue& y_arr, const ArrayValue* x_arr, bool with_const, DesignMatrix* dm,
                         Value* out_err) {
  // y must be 1-D (column m x 1 or row 1 x m).
  const bool y_is_col = y_arr.cols == 1U && y_arr.rows >= 1U;
  const bool y_is_row = y_arr.rows == 1U && y_arr.cols >= 1U;
  if (!y_is_col && !y_is_row) {
    *out_err = Value::error(ErrorCode::Ref);
    return false;
  }
  const std::uint32_t m = y_is_col ? y_arr.rows : y_arr.cols;

  // Coerce y first; errors here propagate (left-to-right rule: y first).
  std::vector<double> y_vec;
  if (!coerce_array(y_arr, y_vec, out_err)) {
    return false;
  }

  // Determine k (number of predictors) and the predictor scan rule.
  std::uint32_t k = 0;
  std::vector<double> x_data;  // m x k row-major
  if (x_arr == nullptr) {
    // Default predictor: {1, 2, ..., m} as a single column.
    k = 1U;
    x_data.resize(m);
    for (std::uint32_t i = 0; i < m; ++i) {
      x_data[i] = static_cast<double>(i + 1U);
    }
  } else {
    std::vector<double> x_vec;
    if (!coerce_array(*x_arr, x_vec, out_err)) {
      return false;
    }
    // Match orientation against y. The "observation axis" of x must
    // equal m. Mac Excel accepts either layout; the other axis becomes k.
    if (y_is_col) {
      // y is m x 1; x should be m x k.
      if (x_arr->rows != m) {
        // Allow x as 1 x m as a fallback (single-variable, transposed).
        if (x_arr->rows == 1U && x_arr->cols == m) {
          k = 1U;
          x_data = x_vec;  // already length m
        } else if (x_arr->cols == m && x_arr->rows >= 1U) {
          // x is k x m: transpose to m x k.
          k = x_arr->rows;
          x_data.resize(static_cast<std::size_t>(m) * k);
          for (std::uint32_t i = 0; i < m; ++i) {
            for (std::uint32_t j = 0; j < k; ++j) {
              x_data[static_cast<std::size_t>(i) * k + j] = x_vec[static_cast<std::size_t>(j) * m + i];
            }
          }
        } else {
          *out_err = Value::error(ErrorCode::Ref);
          return false;
        }
      } else {
        // x is m x k (the canonical column-y layout).
        k = x_arr->cols;
        x_data = x_vec;  // already row-major m x k
      }
    } else {
      // y is 1 x m; x should be k x m (row layout, predictors stacked).
      if (x_arr->cols != m) {
        // Allow x as m x 1 as a fallback (single-variable).
        if (x_arr->cols == 1U && x_arr->rows == m) {
          k = 1U;
          x_data = x_vec;
        } else if (x_arr->rows == m && x_arr->cols >= 1U) {
          // x is m x k: take row-major m x k as-is.
          k = x_arr->cols;
          x_data = x_vec;
        } else {
          *out_err = Value::error(ErrorCode::Ref);
          return false;
        }
      } else {
        // x is k x m: transpose to m x k.
        k = x_arr->rows;
        x_data.resize(static_cast<std::size_t>(m) * k);
        for (std::uint32_t i = 0; i < m; ++i) {
          for (std::uint32_t j = 0; j < k; ++j) {
            x_data[static_cast<std::size_t>(i) * k + j] = x_vec[static_cast<std::size_t>(j) * m + i];
          }
        }
      }
    }
  }

  const std::uint32_t p = k + (with_const ? 1U : 0U);
  // Need at least p observations for the system to be solvable, plus
  // 1 more if we want a non-zero residual df (which LINEST always
  // requires for the SE statistics — but not for the bare coefficients).
  // Reject degenerate dimension up front.
  if (m == 0U || k == 0U) {
    *out_err = Value::error(ErrorCode::Ref);
    return false;
  }

  dm->m = m;
  dm->k = k;
  dm->with_const = with_const;
  dm->y = std::move(y_vec);
  dm->data.resize(static_cast<std::size_t>(m) * p);
  for (std::uint32_t i = 0; i < m; ++i) {
    for (std::uint32_t j = 0; j < k; ++j) {
      dm->data[static_cast<std::size_t>(i) * p + j] = x_data[static_cast<std::size_t>(i) * k + j];
    }
    if (with_const) {
      dm->data[static_cast<std::size_t>(i) * p + k] = 1.0;
    }
  }
  return true;
}

ArrayValue* build_prediction_output(const std::vector<double>& coeffs, const std::vector<double>& pred_design,
                                    std::uint32_t n_obs, std::uint32_t p, bool y_is_col, bool exp_result,
                                    Arena& arena) {
  std::vector<Value> cells;
  cells.reserve(n_obs);
  for (std::uint32_t i = 0; i < n_obs; ++i) {
    double s = 0.0;
    for (std::uint32_t j = 0; j < p; ++j) {
      s += coeffs[j] * pred_design[static_cast<std::size_t>(i) * p + j];
    }
    const double v = exp_result ? std::exp(s) : s;
    cells.push_back(is_finite(v) ? Value::number(v) : Value::error(ErrorCode::Num));
  }

  const std::uint32_t out_rows = y_is_col ? n_obs : 1U;
  const std::uint32_t out_cols = y_is_col ? 1U : n_obs;
  return make_double_array(cells, out_rows, out_cols, arena);
}

/// Builds the LINEST / LOGEST output array. `coeffs` is in normal
/// index order (`coeffs[j]` = coefficient on column j of the design
/// matrix); the output presents them in reverse —
/// `[b_k, b_{k-1}, ..., b_1, b_0]` — which is Excel's right-to-left
/// convention. When `log_form=true`, every coefficient cell is
/// exponentiated and the suppressed-intercept slot becomes `1.0`
/// (i.e. `exp(0)`) — this is the LOGEST output rule.
ArrayValue* build_simple_output(const std::vector<double>& coeffs, std::uint32_t k, bool with_const, Arena& arena,
                                bool log_form = false) {
  const std::uint32_t out_cols = k + 1U;
  std::vector<Value> cells(out_cols, Value::blank());
  for (std::uint32_t j = 0; j < k; ++j) {
    const double raw = coeffs[j];
    const double v = log_form ? std::exp(raw) : raw;
    cells[k - 1U - j] = is_finite(v) ? Value::number(v) : Value::error(ErrorCode::Num);
  }
  if (with_const) {
    const double raw = coeffs[k];
    const double v = log_form ? std::exp(raw) : raw;
    cells[k] = is_finite(v) ? Value::number(v) : Value::error(ErrorCode::Num);
  } else {
    cells[k] = Value::number(log_form ? 1.0 : 0.0);
  }
  return make_double_array(cells, 1U, out_cols, arena);
}

/// Builds the 5x(k+1) statistics matrix. `inv_a` is `A^{-1}` (the
/// inverse of `X^T X`) in normal index order. `coeffs` is in the same
/// order. The reverse-ordering and the `#N/A` padding rules mirror
/// `build_simple_output`. When `log_form=true`, only row 1
/// (coefficients) is exponentiated; rows 2-5 stay on the linear /
/// log-domain scale because Excel's LOGEST documents that the
/// "additional statistics" are computed on the linearised model.
ArrayValue* build_stats_output(const std::vector<double>& coeffs, const std::vector<double>& inv_a,
                               const DesignMatrix& dm, double ss_resid, double ss_total, double mean_y, Arena& arena,
                               bool log_form = false) {
  const std::uint32_t k = dm.k;
  const std::uint32_t p = k + (dm.with_const ? 1U : 0U);
  const std::uint32_t out_cols = k + 1U;
  const std::uint32_t out_rows = 5U;
  std::vector<Value> cells(static_cast<std::size_t>(out_rows) * out_cols, Value::error(ErrorCode::NA));

  const double m = static_cast<double>(dm.m);
  const double df_resid_d = m - static_cast<double>(p);
  const double k_d = static_cast<double>(k);

  // Row 1: coefficients in reverse order. Intercept slot is 0 (LINEST)
  // or 1 (LOGEST) when const=false.
  for (std::uint32_t j = 0; j < k; ++j) {
    const double raw = coeffs[j];
    const double v = log_form ? std::exp(raw) : raw;
    cells[k - 1U - j] = is_finite(v) ? Value::number(v) : Value::error(ErrorCode::Num);
  }
  if (dm.with_const) {
    const double raw = coeffs[k];
    const double v = log_form ? std::exp(raw) : raw;
    cells[k] = is_finite(v) ? Value::number(v) : Value::error(ErrorCode::Num);
  } else {
    cells[k] = Value::number(log_form ? 1.0 : 0.0);
  }

  // Row 2: standard errors. SE[j] = sqrt(inv_a[j,j] * ss_resid / df_resid).
  // df_resid <= 0 (under-determined system) makes SE undefined -> #N/A.
  // When const=false, SE for the intercept slot is #N/A (no intercept
  // was estimated).
  const bool df_ok = df_resid_d > 0.0;
  const double sigma2 = df_ok ? ss_resid / df_resid_d : 0.0;
  for (std::uint32_t j = 0; j < k; ++j) {
    if (!df_ok) {
      cells[out_cols + (k - 1U - j)] = Value::error(ErrorCode::NA);
      continue;
    }
    const double var = sigma2 * inv_a[static_cast<std::size_t>(j) * p + j];
    const double se = (var >= 0.0 && is_finite(var)) ? std::sqrt(var) : std::nan("");
    cells[out_cols + (k - 1U - j)] = is_finite(se) ? Value::number(se) : Value::error(ErrorCode::NA);
  }
  if (dm.with_const) {
    if (!df_ok) {
      cells[out_cols + k] = Value::error(ErrorCode::NA);
    } else {
      const double var = sigma2 * inv_a[static_cast<std::size_t>(k) * p + k];
      const double se = (var >= 0.0 && is_finite(var)) ? std::sqrt(var) : std::nan("");
      cells[out_cols + k] = is_finite(se) ? Value::number(se) : Value::error(ErrorCode::NA);
    }
  } else {
    cells[out_cols + k] = Value::error(ErrorCode::NA);
  }

  // Row 3: [r^2, se_y, #N/A, ...]. Suppress `mean_y` warning when
  // const=false (reference is unused on that branch by design).
  (void)mean_y;
  const double r2 = (ss_total > 0.0 && is_finite(ss_total)) ? 1.0 - ss_resid / ss_total : std::nan("");
  cells[2U * out_cols + 0U] = is_finite(r2) ? Value::number(r2) : Value::error(ErrorCode::NA);
  if (df_ok) {
    const double sey = std::sqrt(sigma2);
    cells[2U * out_cols + 1U] = is_finite(sey) ? Value::number(sey) : Value::error(ErrorCode::NA);
  } else {
    cells[2U * out_cols + 1U] = Value::error(ErrorCode::NA);
  }

  // Row 4: [F, df_resid, #N/A, ...]. F = (ss_reg / k) / (ss_resid / df_resid).
  //
  // An exact fit drives the residual to zero, which makes F infinite in the
  // limit. Excel never lands there: its own arithmetic leaves a residual a
  // few ULPs above zero and it reports a correspondingly huge finite F.
  // Flooring the residual at the same place — the squared rounding error
  // the data's own scale can still resolve — keeps a perfect fit reporting
  // "huge and finite" instead of #N/A. The magnitude at that floor carries
  // no statistical meaning and will not agree with Excel digit for digit;
  // only its finiteness is contractual. A flat response (ss_total == 0)
  // leaves the floor at zero, so F stays #N/A where it is genuinely
  // undefined.
  if (df_ok) {
    const double ss_reg = ss_total - ss_resid;
    constexpr double kEps = std::numeric_limits<double>::epsilon();
    const double resid_floor = ss_total * kEps * kEps;
    const double effective_resid = (ss_resid > resid_floor) ? ss_resid : resid_floor;
    const double f =
        (k_d > 0.0 && effective_resid > 0.0) ? (ss_reg / k_d) / (effective_resid / df_resid_d) : std::nan("");
    cells[3U * out_cols + 0U] = is_finite(f) ? Value::number(f) : Value::error(ErrorCode::NA);
  } else {
    cells[3U * out_cols + 0U] = Value::error(ErrorCode::NA);
  }
  cells[3U * out_cols + 1U] = Value::number(df_resid_d);

  // Row 5: [ss_reg, ss_resid, #N/A, ...].
  const double ss_reg = ss_total - ss_resid;
  cells[4U * out_cols + 0U] = is_finite(ss_reg) ? Value::number(ss_reg) : Value::error(ErrorCode::NA);
  cells[4U * out_cols + 1U] = is_finite(ss_resid) ? Value::number(ss_resid) : Value::error(ErrorCode::NA);

  return make_double_array(cells, out_rows, out_cols, arena);
}

}  // namespace

Value eval_linest_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 1U || arity > 4U) {
    return Value::error(ErrorCode::Value);
  }

  const ArrayValue* y_arr = nullptr;
  Value err = Value::blank();
  if (!resolve_array_value(call.as_call_arg(0), arena, registry, ctx, &y_arr, &err)) {
    return err;
  }

  const ArrayValue* x_arr = nullptr;
  if (arity >= 2U) {
    if (!resolve_array_value(call.as_call_arg(1), arena, registry, ctx, &x_arr, &err)) {
      return err;
    }
  }

  bool with_const = true;
  if (arity >= 3U) {
    if (!eval_bool_arg(call.as_call_arg(2), arena, registry, ctx, true, &with_const, &err)) {
      return err;
    }
  }

  bool want_stats = false;
  if (arity >= 4U) {
    if (!eval_bool_arg(call.as_call_arg(3), arena, registry, ctx, false, &want_stats, &err)) {
      return err;
    }
  }

  DesignMatrix dm;
  if (!build_design_matrix(*y_arr, x_arr, with_const, &dm, &err)) {
    return err;
  }

  const std::uint32_t p = dm.k + (with_const ? 1U : 0U);
  if (dm.m < p) {
    // Fewer observations than parameters -> system is under-determined.
    return Value::error(ErrorCode::Num);
  }

  std::vector<double> coeffs;
  std::vector<double> inv_a;
  if (!solve_normal_equations(dm, /*want_inv=*/want_stats, &coeffs, &inv_a)) {
    return Value::error(ErrorCode::Num);
  }

  if (!want_stats) {
    // `build_simple_output` reads coeffs[0..k-1] always and coeffs[k]
    // only on the `with_const=true` branch — so `coeffs.size() == p`
    // is correct in both branches.
    ArrayValue* arr = build_simple_output(coeffs, dm.k, with_const, arena);
    if (arr == nullptr) {
      return Value::error(ErrorCode::Num);
    }
    return Value::array(arr);
  }

  const FitStats stats = compute_fit_stats(dm, coeffs);
  ArrayValue* arr = build_stats_output(coeffs, inv_a, dm, stats.ss_resid, stats.ss_total, stats.mean_y, arena);
  if (arr == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  return Value::array(arr);
}

namespace {

/// Resolves `new_x` against an already-built design matrix, producing
/// an `N x p` row-major buffer (predictor columns followed by an
/// intercept-1 column when `with_const`). The exact axis interpretation
/// mirrors `build_design_matrix`: when `known_y` is a column, `new_x`'s
/// predictor axis is its columns; when row, its rows. A single-variable
/// fallback (`k == 1`) accepts the orthogonal orientation as a courtesy.
///
/// Returns `false` on shape mismatch (`#REF!`) or non-numeric / error
/// cells (errors propagate verbatim, non-numeric -> `#VALUE!`).
bool resolve_new_x_design(const ArrayValue& new_x, const DesignMatrix& dm, bool y_is_col, bool with_const,
                          std::vector<double>* out_design, std::uint32_t* out_n, Value* out_err) {
  std::vector<double> raw;
  if (!coerce_array(new_x, raw, out_err)) {
    return false;
  }
  const std::uint32_t k = dm.k;
  std::uint32_t n = 0;
  std::vector<double> features;  // N x k row-major
  if (y_is_col) {
    if (new_x.cols == k) {
      n = new_x.rows;
      features = std::move(raw);  // already N x k row-major
    } else if (k == 1U && new_x.rows == 1U && new_x.cols >= 1U) {
      // Single-variable fallback: 1 x N row works as N x 1 column.
      n = new_x.cols;
      features = std::move(raw);
    } else {
      *out_err = Value::error(ErrorCode::Ref);
      return false;
    }
  } else {
    if (new_x.rows == k) {
      // k x N -> transpose to N x k.
      n = new_x.cols;
      features.resize(static_cast<std::size_t>(n) * k);
      for (std::uint32_t i = 0; i < n; ++i) {
        for (std::uint32_t j = 0; j < k; ++j) {
          features[static_cast<std::size_t>(i) * k + j] = raw[static_cast<std::size_t>(j) * n + i];
        }
      }
    } else if (k == 1U && new_x.cols == 1U && new_x.rows >= 1U) {
      // Single-variable fallback: N x 1 column works as 1 x N row.
      n = new_x.rows;
      features = std::move(raw);
    } else {
      *out_err = Value::error(ErrorCode::Ref);
      return false;
    }
  }

  const std::uint32_t p = k + (with_const ? 1U : 0U);
  out_design->resize(static_cast<std::size_t>(n) * p);
  for (std::uint32_t i = 0; i < n; ++i) {
    for (std::uint32_t j = 0; j < k; ++j) {
      (*out_design)[static_cast<std::size_t>(i) * p + j] = features[static_cast<std::size_t>(i) * k + j];
    }
    if (with_const) {
      (*out_design)[static_cast<std::size_t>(i) * p + k] = 1.0;
    }
  }
  *out_n = n;
  return true;
}

}  // namespace

Value eval_trend_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 1U || arity > 4U) {
    return Value::error(ErrorCode::Value);
  }

  const ArrayValue* y_arr = nullptr;
  Value err = Value::blank();
  if (!resolve_array_value(call.as_call_arg(0), arena, registry, ctx, &y_arr, &err)) {
    return err;
  }
  const ArrayValue* x_arr = nullptr;
  if (arity >= 2U) {
    if (!resolve_array_value(call.as_call_arg(1), arena, registry, ctx, &x_arr, &err)) {
      return err;
    }
  }
  const ArrayValue* new_x_arr = nullptr;
  if (arity >= 3U && !is_omitted_arg(call.as_call_arg(2))) {
    if (!resolve_array_value(call.as_call_arg(2), arena, registry, ctx, &new_x_arr, &err)) {
      return err;
    }
  }
  bool with_const = true;
  if (arity >= 4U) {
    if (!eval_bool_arg(call.as_call_arg(3), arena, registry, ctx, true, &with_const, &err)) {
      return err;
    }
  }

  DesignMatrix dm;
  if (!build_design_matrix(*y_arr, x_arr, with_const, &dm, &err)) {
    return err;
  }
  const std::uint32_t p = dm.k + (with_const ? 1U : 0U);
  if (dm.m < p) {
    return Value::error(ErrorCode::Num);
  }

  std::vector<double> coeffs;
  std::vector<double> unused_inv;
  if (!solve_normal_equations(dm, /*want_inv=*/false, &coeffs, &unused_inv)) {
    return Value::error(ErrorCode::Num);
  }

  // Build the prediction design buffer. When new_x is omitted, this is
  // the original design matrix (returns the fitted values y_hat at the
  // training observations); otherwise we build a fresh N x p buffer.
  const bool y_is_col = y_arr->cols == 1U && y_arr->rows >= 1U;
  std::vector<double> pred_design;
  std::uint32_t n_obs = 0;
  if (new_x_arr == nullptr) {
    n_obs = dm.m;
    pred_design = dm.data;  // m x p row-major; intercept column already
                            // present when with_const
  } else {
    if (!resolve_new_x_design(*new_x_arr, dm, y_is_col, with_const, &pred_design, &n_obs, &err)) {
      return err;
    }
  }

  ArrayValue* arr = build_prediction_output(coeffs, pred_design, n_obs, p, y_is_col, /*exp_result=*/false, arena);
  if (arr == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  return Value::array(arr);
}

namespace {

/// Replaces every entry of `dm.y` with `ln(y_i)`. Returns `false` if
/// any `y_i <= 0` — `ln` is undefined there, and Mac Excel's
/// LOGEST / GROWTH surface `#NUM!` in that case.
bool log_transform_y(DesignMatrix& dm) {
  for (std::uint32_t i = 0; i < dm.m; ++i) {
    if (!(dm.y[i] > 0.0)) {
      return false;
    }
    dm.y[i] = std::log(dm.y[i]);
  }
  return true;
}

}  // namespace

Value eval_logest_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  // Same argument shape as LINEST. The only differences are the
  // log-transform on `y` before solving and the `exp()` of the row-1
  // coefficients in the output.
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 1U || arity > 4U) {
    return Value::error(ErrorCode::Value);
  }

  const ArrayValue* y_arr = nullptr;
  Value err = Value::blank();
  if (!resolve_array_value(call.as_call_arg(0), arena, registry, ctx, &y_arr, &err)) {
    return err;
  }
  const ArrayValue* x_arr = nullptr;
  if (arity >= 2U) {
    if (!resolve_array_value(call.as_call_arg(1), arena, registry, ctx, &x_arr, &err)) {
      return err;
    }
  }
  bool with_const = true;
  if (arity >= 3U) {
    if (!eval_bool_arg(call.as_call_arg(2), arena, registry, ctx, true, &with_const, &err)) {
      return err;
    }
  }
  bool want_stats = false;
  if (arity >= 4U) {
    if (!eval_bool_arg(call.as_call_arg(3), arena, registry, ctx, false, &want_stats, &err)) {
      return err;
    }
  }

  DesignMatrix dm;
  if (!build_design_matrix(*y_arr, x_arr, with_const, &dm, &err)) {
    return err;
  }
  if (!log_transform_y(dm)) {
    return Value::error(ErrorCode::Num);
  }
  const std::uint32_t p = dm.k + (with_const ? 1U : 0U);
  if (dm.m < p) {
    return Value::error(ErrorCode::Num);
  }

  std::vector<double> coeffs;
  std::vector<double> inv_a;
  if (!solve_normal_equations(dm, /*want_inv=*/want_stats, &coeffs, &inv_a)) {
    return Value::error(ErrorCode::Num);
  }

  if (!want_stats) {
    ArrayValue* arr = build_simple_output(coeffs, dm.k, with_const, arena, /*log_form=*/true);
    if (arr == nullptr) {
      return Value::error(ErrorCode::Num);
    }
    return Value::array(arr);
  }

  const FitStats stats = compute_fit_stats(dm, coeffs);
  ArrayValue* arr =
      build_stats_output(coeffs, inv_a, dm, stats.ss_resid, stats.ss_total, stats.mean_y, arena, /*log_form=*/true);
  if (arr == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  return Value::array(arr);
}

Value eval_growth_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  // Same argument shape as TREND. The body is the TREND recipe with
  // a `log()` on `y` before solving and an `exp()` on the predicted
  // log-scale `y_hat`.
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 1U || arity > 4U) {
    return Value::error(ErrorCode::Value);
  }

  const ArrayValue* y_arr = nullptr;
  Value err = Value::blank();
  if (!resolve_array_value(call.as_call_arg(0), arena, registry, ctx, &y_arr, &err)) {
    return err;
  }
  const ArrayValue* x_arr = nullptr;
  if (arity >= 2U) {
    if (!resolve_array_value(call.as_call_arg(1), arena, registry, ctx, &x_arr, &err)) {
      return err;
    }
  }
  const ArrayValue* new_x_arr = nullptr;
  if (arity >= 3U && !is_omitted_arg(call.as_call_arg(2))) {
    if (!resolve_array_value(call.as_call_arg(2), arena, registry, ctx, &new_x_arr, &err)) {
      return err;
    }
  }
  bool with_const = true;
  if (arity >= 4U) {
    if (!eval_bool_arg(call.as_call_arg(3), arena, registry, ctx, true, &with_const, &err)) {
      return err;
    }
  }

  DesignMatrix dm;
  if (!build_design_matrix(*y_arr, x_arr, with_const, &dm, &err)) {
    return err;
  }
  if (!log_transform_y(dm)) {
    return Value::error(ErrorCode::Num);
  }
  const std::uint32_t p = dm.k + (with_const ? 1U : 0U);
  if (dm.m < p) {
    return Value::error(ErrorCode::Num);
  }

  std::vector<double> coeffs;
  std::vector<double> unused_inv;
  if (!solve_normal_equations(dm, /*want_inv=*/false, &coeffs, &unused_inv)) {
    return Value::error(ErrorCode::Num);
  }

  const bool y_is_col = y_arr->cols == 1U && y_arr->rows >= 1U;
  std::vector<double> pred_design;
  std::uint32_t n_obs = 0;
  if (new_x_arr == nullptr) {
    n_obs = dm.m;
    pred_design = dm.data;
  } else {
    if (!resolve_new_x_design(*new_x_arr, dm, y_is_col, with_const, &pred_design, &n_obs, &err)) {
      return err;
    }
  }

  ArrayValue* arr = build_prediction_output(coeffs, pred_design, n_obs, p, y_is_col, /*exp_result=*/true, arena);
  if (arr == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  return Value::array(arr);
}

}  // namespace eval
}  // namespace formulon
