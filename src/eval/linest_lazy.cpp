//
// LINEST implementation. See linest_lazy.h for the user-facing
// contract; this file owns the numerical recipe (build the design
// matrix, solve by Householder QR, optionally derive the regression
// statistics). One kernel serves LINEST, LOGEST, TREND and GROWTH.
//
// The kernel is hand-rolled to keep the WASM size policy intact (no
// Eigen, no LAPACK). Operation order follows the measured Mac Excel 365
// results to the last bit.

#include "eval/linest_lazy.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <vector>

#include "eval/array_alloc.h"
#include "eval/lazy_impls.h"
#include "eval/matrix_ops_lazy.h"
#include "eval/omitted_arg.h"
#include "eval/range_args.h"
#include "eval/shape_ops_lazy.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

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

/// Least-squares result in the caller's parameter order `[b_0 .. b_{k-1}, a]`.
struct LinFit {
  std::vector<double> coeffs;     // dropped (collinear) columns are 0
  std::vector<double> var_diag;   // diag of (X'X)^-1 per parameter; 0 for dropped columns
  double ss_resid = 0.0;          // sum of squares of the reflected-y tail
  double sxx = 0.0;               // univariate: sum of squared (centred) x
  double mean_x = 0.0;            // univariate with const: mean of x
  std::size_t ratio_dropped = 0;  // columns dropped by the norm-ratio test (not exact zeros)
  double r00 = 0.0;               // univariate with const: R of [1, x - mean]
  double r01 = 0.0;
  double r11 = 0.0;
};

/// A column is dropped as collinear when its norm after the earlier
/// reflections falls below this fraction of its original norm.
constexpr double kCollinearRatio = 1e-6;

double sum_forward(const std::vector<double>& v) {
  double s = 0.0;
  for (const double x : v) {
    s += x;
  }
  return s;
}

double sum_backward(const std::vector<double>& v) {
  double s = 0.0;
  for (std::size_t i = v.size(); i > 0U; --i) {
    s += v[i - 1U];
  }
  return s;
}

/// Sum of squares of `col[from..]`, accumulated forward.
double column_norm_sq(const std::vector<double>& col, std::size_t from) {
  double s = 0.0;
  for (std::size_t i = from; i < col.size(); ++i) {
    s += col[i] * col[i];
  }
  return s;
}

/// Householder-QR least squares. With a constant, x and y are centred by
/// their means but a ones column is kept in the design; columns are
/// reflected ones-first then by descending centred norm. Parameters come
/// back in original order with the intercept recovered from the means.
void fit_least_squares(const DesignMatrix& dm, bool want_inv, LinFit* out) {
  const std::size_t m = dm.m;
  const std::size_t k = dm.k;
  const std::size_t p = k + (dm.with_const ? 1U : 0U);
  const std::size_t off = dm.with_const ? 1U : 0U;

  std::vector<std::vector<double>> xcols(k, std::vector<double>(m));
  for (std::size_t j = 0; j < k; ++j) {
    for (std::size_t i = 0; i < m; ++i) {
      xcols[j][i] = dm.data[i * p + j];
    }
  }
  std::vector<double> yy = dm.y;
  std::vector<double> mx(k, 0.0);
  double my = 0.0;
  if (dm.with_const) {
    for (std::size_t j = 0; j < k; ++j) {
      mx[j] = sum_forward(xcols[j]) / static_cast<double>(m);
      for (std::size_t i = 0; i < m; ++i) {
        xcols[j][i] -= mx[j];
      }
    }
    my = sum_forward(yy) / static_cast<double>(m);
    for (std::size_t i = 0; i < m; ++i) {
      yy[i] -= my;
    }
  }

  std::vector<std::size_t> order(k);
  std::vector<double> norm_sq(k);
  for (std::size_t j = 0; j < k; ++j) {
    order[j] = j;
    norm_sq[j] = column_norm_sq(xcols[j], 0U);
  }
  std::stable_sort(order.begin(), order.end(), [&](std::size_t a, std::size_t b) { return norm_sq[a] > norm_sq[b]; });
  // Of two dependent columns of equal norm the first survives without a constant, the last with one.
  if (dm.with_const) {
    for (std::size_t t = 0; t + 1U < k; ++t) {
      const std::size_t a = order[t];
      const std::size_t b = order[t + 1U];
      if (norm_sq[a] != norm_sq[b] || a > b) {
        continue;
      }
      double dot = 0.0;
      for (std::size_t i = 0; i < m; ++i) {
        dot += xcols[a][i] * xcols[b][i];
      }
      const double resid = norm_sq[b] - dot * dot / norm_sq[a];
      if (norm_sq[a] > 0.0 && resid < kCollinearRatio * kCollinearRatio * norm_sq[b]) {
        std::swap(order[t], order[t + 1U]);
      }
    }
  }

  // Reflected columns in processing order; `param[t]` is the slot in [c0?, x_0..].
  std::vector<std::vector<double>> cols;
  std::vector<std::size_t> param;
  if (dm.with_const) {
    cols.emplace_back(m, 1.0);
    param.push_back(0U);
  }
  for (const std::size_t j : order) {
    cols.push_back(xcols[j]);
    param.push_back(off + j);
  }
  std::vector<double> orig_norm(cols.size());
  for (std::size_t t = 0; t < cols.size(); ++t) {
    orig_norm[t] = std::sqrt(column_norm_sq(cols[t], 0U));
  }

  std::vector<std::size_t> surv;  // processing index of each kept column; row = ordinal
  std::vector<double> v(m);
  for (std::size_t t = 0; t < cols.size(); ++t) {
    const std::size_t r = surv.size();
    if (r >= m) {
      continue;
    }
    const double nrm = std::sqrt(column_norm_sq(cols[t], r));
    if (nrm == 0.0) {
      continue;
    }
    if (nrm < kCollinearRatio * orig_norm[t]) {
      ++out->ratio_dropped;
      continue;
    }
    const double x1 = cols[t][r];
    const double beta = -std::copysign(nrm, x1);
    std::vector<double> vtail;
    v[r] = x1 - beta;
    for (std::size_t i = r + 1U; i < m; ++i) {
      v[i] = cols[t][i];
    }
    vtail.assign(v.begin() + static_cast<std::ptrdiff_t>(r), v.end());
    for (double& e : vtail) {
      e *= e;
    }
    const double vn = sum_backward(vtail);
    const auto reflect = [&](std::vector<double>& col) {
      double d = 0.0;
      for (std::size_t i = r; i < m; ++i) {
        d += v[i] * col[i];
      }
      const double f = 2.0 * d / vn;
      for (std::size_t i = r; i < m; ++i) {
        col[i] -= f * v[i];
      }
    };
    for (std::size_t u = t + 1U; u < cols.size(); ++u) {
      reflect(cols[u]);
    }
    reflect(yy);
    cols[t][r] = beta;
    surv.push_back(t);
  }

  const std::size_t rank = surv.size();
  const auto r_at = [&](std::size_t i, std::size_t j) { return cols[surv[j]][i]; };
  std::vector<double> sol(rank, 0.0);
  for (std::size_t i = rank; i > 0U; --i) {
    const std::size_t row = i - 1U;
    double s = yy[row];
    for (std::size_t j = row + 1U; j < rank; ++j) {
      s -= r_at(row, j) * sol[j];
    }
    sol[row] = s / r_at(row, row);
  }

  std::vector<double> pc(p, 0.0);  // centred-parameter values, [c0?, x_0..]
  for (std::size_t s = 0; s < rank; ++s) {
    pc[param[surv[s]]] = sol[s];
  }
  out->coeffs.assign(p, 0.0);
  for (std::size_t j = 0; j < k; ++j) {
    out->coeffs[j] = pc[off + j];
  }
  if (dm.with_const) {
    double a = my;
    for (std::size_t j = 0; j < k; ++j) {
      a -= out->coeffs[j] * mx[j];
    }
    out->coeffs[k] = a;
  }

  std::vector<double> tail;
  for (std::size_t i = rank; i < m; ++i) {
    tail.push_back(yy[i] * yy[i]);
  }
  out->ss_resid = sum_forward(tail);

  if (k == 1U) {
    std::vector<double> sq(m);
    for (std::size_t i = 0; i < m; ++i) {
      sq[i] = xcols[0][i] * xcols[0][i];
    }
    out->sxx = sum_backward(sq);
    out->mean_x = mx[0];
    if (dm.with_const && rank == 2U) {
      out->r00 = r_at(0U, 0U);
      out->r01 = r_at(0U, 1U);
      out->r11 = r_at(1U, 1U);
    }
  }

  if (!want_inv) {
    return;
  }
  // (R'R)^-1 over the kept columns, scattered to centred-parameter slots.
  std::vector<double> rinv(rank * rank, 0.0);
  for (std::size_t c = 0; c < rank; ++c) {
    for (std::size_t i = c + 1U; i > 0U; --i) {
      const std::size_t row = i - 1U;
      double s = (row == c) ? 1.0 : 0.0;
      for (std::size_t j = row + 1U; j <= c; ++j) {
        s -= r_at(row, j) * rinv[j * rank + c];
      }
      rinv[row * rank + c] = s / r_at(row, row);
    }
  }
  std::vector<double> vp(p * p, 0.0);
  for (std::size_t a = 0; a < rank; ++a) {
    for (std::size_t b = 0; b < rank; ++b) {
      double s = 0.0;
      for (std::size_t l = std::max(a, b); l < rank; ++l) {
        s += rinv[a * rank + l] * rinv[b * rank + l];
      }
      vp[param[surv[a]] * p + param[surv[b]]] = s;
    }
  }
  out->var_diag.assign(p, 0.0);
  for (std::size_t j = 0; j < k; ++j) {
    out->var_diag[j] = vp[(off + j) * p + (off + j)];
  }
  if (dm.with_const) {
    double var_a = vp[0];
    for (std::size_t j = 0; j < k; ++j) {
      var_a -= 2.0 * mx[j] * vp[off + j];
      for (std::size_t l = 0; l < k; ++l) {
        var_a += mx[j] * mx[l] * vp[(off + j) * p + (off + l)];
      }
    }
    out->var_diag[k] = var_a;
  }
}

FitStats compute_fit_stats(const DesignMatrix& dm, const LinFit& fit) {
  FitStats stats;
  stats.ss_resid = fit.ss_resid;
  double sum_y = 0.0;
  for (std::uint32_t i = 0; i < dm.m; ++i) {
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
  if (!coerce_matrix(y_arr, y_vec, *out_err)) {
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
    if (!coerce_matrix(*x_arr, x_vec, *out_err)) {
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
                                    std::uint32_t n_obs, std::uint32_t p, bool with_const, bool y_is_col,
                                    bool exp_result, Arena& arena) {
  const std::uint32_t k = with_const ? p - 1U : p;
  std::vector<Value> cells;
  cells.reserve(n_obs);
  for (std::uint32_t i = 0; i < n_obs; ++i) {
    double v = 0.0;
    if (exp_result) {
      // exp(a) times one exp factor per predictor, in column order.
      v = with_const ? std::exp(coeffs[k]) : 1.0;
      for (std::uint32_t j = 0; j < k; ++j) {
        v *= std::exp(coeffs[j] * pred_design[static_cast<std::size_t>(i) * p + j]);
      }
    } else {
      v = with_const ? coeffs[k] : 0.0;
      for (std::uint32_t j = 0; j < k; ++j) {
        v += coeffs[j] * pred_design[static_cast<std::size_t>(i) * p + j];
      }
    }
    cells.push_back(is_finite(v) ? Value::number(v) : Value::error(ErrorCode::Num));
  }

  const std::uint32_t out_rows = y_is_col ? n_obs : 1U;
  const std::uint32_t out_cols = y_is_col ? 1U : n_obs;
  return array_from_values(out_rows, out_cols, cells.data(), cells.size(), arena);
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
  return array_from_values(1U, out_cols, cells.data(), cells.size(), arena);
}

/// Builds the 5x(k+1) statistics matrix. `fit.var_diag` is the diagonal
/// of `(X^T X)^{-1}` in normal index order, like `fit.coeffs`. The reverse-ordering and the `#N/A` padding rules mirror
/// `build_simple_output`. When `log_form=true`, only row 1
/// (coefficients) is exponentiated; rows 2-5 stay on the linear /
/// log-domain scale because Excel's LOGEST documents that the
/// "additional statistics" are computed on the linearised model.
ArrayValue* build_stats_output(const LinFit& fit, const DesignMatrix& dm, double ss_resid, double ss_total,
                               Arena& arena, bool log_form = false) {
  const std::vector<double>& coeffs = fit.coeffs;
  const std::uint32_t k = dm.k;
  const std::uint32_t p = k + (dm.with_const ? 1U : 0U);
  const std::uint32_t out_cols = k + 1U;
  const std::uint32_t out_rows = 5U;
  std::vector<Value> cells(static_cast<std::size_t>(out_rows) * out_cols, Value::error(ErrorCode::NA));

  const double m = static_cast<double>(dm.m);
  const double df_resid_d = std::max(0.0, m - static_cast<double>(p) + static_cast<double>(fit.ratio_dropped));
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

  // Row 2: standard errors. SE[j] = sqrt(var_diag[j] * ss_resid / df_resid);
  // the univariate case follows Excel's operation order instead.
  // df_resid == 0 (saturated fit) reports zero errors.
  // When const=false, SE for the intercept slot is #N/A (no intercept
  // was estimated).
  const bool df_ok = df_resid_d > 0.0;
  const double sigma2 = df_ok ? ss_resid / df_resid_d : 0.0;
  const double sey = df_ok ? std::sqrt(sigma2) : 0.0;
  const bool uni = k == 1U && fit.var_diag[0] > 0.0;
  for (std::uint32_t j = 0; j < k; ++j) {
    const double var = uni ? sey * sey / fit.sxx : sigma2 * fit.var_diag[j];
    const double se = (var >= 0.0 && is_finite(var)) ? std::sqrt(var) : std::nan("");
    cells[out_cols + (k - 1U - j)] = is_finite(se) ? Value::number(se) : Value::error(ErrorCode::NA);
  }
  if (dm.with_const) {
    {
      double var_a = sigma2 * fit.var_diag[k];
      if (uni) {
        // v = (R'R)^-1 of the [1, x - mean] QR, shifted back by the mean.
        const double i11 = 1.0 / fit.r00;
        const double i12 = -fit.r01 / (fit.r00 * fit.r11);
        const double i22 = 1.0 / fit.r11;
        const double v00 = i11 * i11 + i12 * i12;
        const double v01 = i12 * i22;
        const double v11 = i22 * i22;
        const double s2 = sey * sey;
        var_a = s2 * v00 + s2 * fit.mean_x * fit.mean_x * v11 - 2 * s2 * fit.mean_x * v01;
      }
      const double se = (var_a >= 0.0 && is_finite(var_a)) ? std::sqrt(var_a) : std::nan("");
      cells[out_cols + k] = is_finite(se) ? Value::number(se) : Value::error(ErrorCode::NA);
    }
  } else {
    cells[out_cols + k] = Value::error(ErrorCode::NA);
  }

  // Row 3: [r^2, se_y, #N/A, ...].
  const double ss_reg_r2 = ss_total - ss_resid;
  const double r2 = (ss_total > 0.0 && is_finite(ss_total)) ? ss_reg_r2 / (ss_reg_r2 + ss_resid) : std::nan("");
  cells[2U * out_cols + 0U] = is_finite(r2) ? Value::number(r2) : Value::error(ErrorCode::NA);
  cells[2U * out_cols + 1U] = is_finite(sey) ? Value::number(sey) : Value::error(ErrorCode::NA);

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

  return array_from_values(out_rows, out_cols, cells.data(), cells.size(), arena);
}

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

/// Reads the `known_y, [known_x]` head all four regression builtins
/// take, along with the arity bound they share. `*out_x` stays null when
/// `known_x` was omitted, which the design-matrix builder reads as the
/// implicit `{1, 2, 3, ...}` predictor.
///
/// Returns `false` with the propagating error in `*out_err`.
bool read_known_xy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                   const ArrayValue** out_y, const ArrayValue** out_x, Value* out_err) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 1U || arity > 4U) {
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  if (!resolve_array_value(call.as_call_arg(0), arena, registry, ctx, out_y, out_err)) {
    return false;
  }
  if (arity >= 2U && !resolve_array_value(call.as_call_arg(1), arena, registry, ctx, out_x, out_err)) {
    return false;
  }
  return true;
}

/// Builds the design matrix, applies the `ln(y)` transform when
/// `log_form`, and solves the least-squares fit — the step shared by the
/// fit and prediction families. Non-positive y under `log_form` or fewer
/// observations than parameters is `#NUM!`.
///
/// Returns `false` with the propagating error in `*out_err`.
bool fit_coefficients(const ArrayValue& y_arr, const ArrayValue* x_arr, bool with_const, bool log_form, bool want_inv,
                      DesignMatrix* out_dm, LinFit* out_fit, Value* out_err) {
  if (!build_design_matrix(y_arr, x_arr, with_const, out_dm, out_err)) {
    return false;
  }
  if (log_form && !log_transform_y(*out_dm)) {
    *out_err = Value::error(ErrorCode::Num);
    return false;
  }
  if (out_dm->m < out_dm->k) {
    // Fewer observations than predictors -> system is under-determined.
    *out_err = Value::error(ErrorCode::Num);
    return false;
  }
  fit_least_squares(*out_dm, want_inv, out_fit);
  return true;
}

/// LINEST and LOGEST take the same four arguments — `known_y`,
/// `known_x`, `const`, `stats` — and run the same regression. LOGEST
/// fits `ln(y)` and reports the coefficients in exponential form;
/// `log_form` selects between the two.
Value eval_fit(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
               bool log_form) {
  const ArrayValue* y_arr = nullptr;
  const ArrayValue* x_arr = nullptr;
  Value err = Value::blank();
  if (!read_known_xy(call, arena, registry, ctx, &y_arr, &x_arr, &err)) {
    return err;
  }
  const std::uint32_t arity = call.as_call_arity();

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
  LinFit fit;
  if (!fit_coefficients(*y_arr, x_arr, with_const, log_form, /*want_inv=*/want_stats, &dm, &fit, &err)) {
    return err;
  }
  const std::vector<double>& coeffs = fit.coeffs;

  if (!want_stats) {
    // `build_simple_output` reads coeffs[0..k-1] always and coeffs[k]
    // only on the `with_const=true` branch — so `coeffs.size() == p`
    // is correct in both branches.
    ArrayValue* arr = build_simple_output(coeffs, dm.k, with_const, arena, log_form);
    if (arr == nullptr) {
      return Value::error(ErrorCode::Num);
    }
    return Value::array(arr);
  }

  const FitStats stats = compute_fit_stats(dm, fit);
  ArrayValue* arr = build_stats_output(fit, dm, stats.ss_resid, stats.ss_total, arena, log_form);
  if (arr == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  return Value::array(arr);
}

}  // namespace

Value eval_linest_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  return eval_fit(call, arena, registry, ctx, /*log_form=*/false);
}

Value eval_logest_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  return eval_fit(call, arena, registry, ctx, /*log_form=*/true);
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
  if (!coerce_matrix(new_x, raw, *out_err)) {
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

/// TREND and GROWTH take the same four arguments — `known_y`,
/// `known_x`, `new_x`, `const` — and predict from the same solve.
/// GROWTH fits `ln(y)` and exponentiates the prediction; `log_form`
/// selects between the two.
Value eval_prediction(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx, bool log_form) {
  const ArrayValue* y_arr = nullptr;
  const ArrayValue* x_arr = nullptr;
  Value err = Value::blank();
  if (!read_known_xy(call, arena, registry, ctx, &y_arr, &x_arr, &err)) {
    return err;
  }
  const std::uint32_t arity = call.as_call_arity();
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
  LinFit fit;
  if (!fit_coefficients(*y_arr, x_arr, with_const, log_form, /*want_inv=*/false, &dm, &fit, &err)) {
    return err;
  }
  const std::vector<double>& coeffs = fit.coeffs;
  const std::uint32_t p = dm.k + (with_const ? 1U : 0U);

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

  ArrayValue* arr =
      build_prediction_output(coeffs, pred_design, n_obs, p, with_const, y_is_col, /*exp_result=*/log_form, arena);
  if (arr == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  return Value::array(arr);
}

}  // namespace

Value eval_trend_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx) {
  return eval_prediction(call, arena, registry, ctx, /*log_form=*/false);
}

Value eval_growth_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  return eval_prediction(call, arena, registry, ctx, /*log_form=*/true);
}

}  // namespace eval
}  // namespace formulon
