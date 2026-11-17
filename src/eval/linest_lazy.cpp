//
// LINEST implementation. See linest_lazy.h for the user-facing
// contract; this file owns the numerical recipe (build the design
// matrix, solve by Householder QR, optionally derive the regression
// statistics). One kernel serves LINEST, LOGEST, TREND and GROWTH.
//
// The kernel is hand-rolled to keep the WASM size policy intact (no
// Eigen, no LAPACK). Every rounding step is fixed: no fused multiply-add,
// rows rebuilt from the raw data at each Householder step, downdated-norm
// pivoting, forward sums in row order.

#include "eval/linest_lazy.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <utility>
#include <vector>

#include "eval/array_alloc.h"
#include "eval/lazy_impls.h"
#include "eval/matrix_ops_lazy.h"
#include "eval/omitted_arg.h"
#include "eval/range_args.h"
#include "eval/scalar_args.h"
#include "eval/shape_ops_lazy.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/elementary_math.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

/// The value of optional flag argument `index`: `default_value` when the
/// slot is absent or omitted, otherwise the evaluated argument.
Value flag_arg_value(const parser::AstNode& call, std::uint32_t index, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx, bool default_value) {
  if (!call_arg_supplied(call, index)) {
    return Value::boolean(default_value);
  }
  return eval_node(call.as_call_arg(index), arena, registry, ctx);
}

/// Reads a flag: Booleans pass through, numbers are non-zero, a blank cell is
/// FALSE. Anything else, errors included, is `#VALUE!`.
bool read_flag(const Value& v, bool* out) {
  if (v.is_boolean()) {
    *out = v.as_boolean();
    return true;
  }
  if (v.is_number()) {
    *out = v.as_number() != 0.0;
    return true;
  }
  if (v.is_blank()) {
    *out = false;
    return true;
  }
  return false;
}

/// Runs `fit` once for scalar flags, or element by element when a flag is an
/// array; each lifted element is the top-left cell of its result.
template <typename Fit>
Value lift_flags(const Value* flags, std::uint32_t count, Arena& arena, const Fit& fit) {
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
  const bool lifted = lift_extent(flags, count, nullptr, &rows, &cols);
  return lift_over_array_args(flags, count, arena, [&](const Value* args, std::uint32_t n) {
    bool values[2] = {false, false};
    for (std::uint32_t i = 0; i < n; ++i) {
      if (!read_flag(args[i], &values[i])) {
        return Value::error(ErrorCode::Value);
      }
    }
    const Value out = fit(values);
    return lifted && out.is_array() ? out.as_array()->cells[0] : out;
  });
}

/// Coerces a regression input matrix: any cell that is not a number, an
/// error included, is `#VALUE!`.
bool coerce_fit_matrix(const ArrayValue& src, std::vector<double>& out, Value& out_err) {
  if (coerce_matrix(src, out, out_err)) {
    return true;
  }
  out_err = Value::error(ErrorCode::Value);
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

/// Least-squares result in the caller's parameter order `[b_0 .. b_{k-1}, a]`.
struct LinFit {
  std::vector<double> coeffs;  // columns that never became a pivot are 0
  std::vector<double> se;      // standard errors in the same order (0 for unused columns)
  double ss_resid = 0.0;       // squared transformed y below the pivot rows
  double ss_total = 0.0;       // sum of squared (centred) y
  double df = 0.0;             // m minus the number of pivot steps
  std::size_t steps = 0;       // pivot steps taken (including the ones column)
};

/// A column stops being a pivot once its downdated norm^2 falls below this
/// fraction of its original norm^2.
constexpr double kPivotFloor = 4.547473508864641e-13;  // 2^-41

/// Division that yields 0 for a zero divisor.
double div0(double a, double b) {
  return b == 0.0 ? 0.0 : a / b;
}

/// Householder-QR least squares on the design [x_1..x_k, ones?] with the
/// right-hand side appended as a last column. With a constant, x and y are
/// centred by their means and a ones column is kept. Rows are rebuilt from the
/// raw data at every step by replaying the earlier multipliers.
void fit_least_squares(const DesignMatrix& dm, bool want_se, LinFit* out) {
  const std::size_t m = dm.m;
  const std::size_t nx = dm.k;
  const std::size_t p = nx + (dm.with_const ? 1U : 0U);
  const std::size_t ncol = p;
  const std::size_t ycol = ncol;
  const std::size_t width = ncol + 1U;
  const std::size_t ones = nx;

  std::vector<double> means(nx, 0.0);
  double ymean = 0.0;
  if (dm.with_const) {
    for (std::size_t j = 0; j < nx; ++j) {
      double s = 0.0;
      for (std::size_t r = 0; r < m; ++r) {
        s += dm.data[r * p + j];
      }
      means[j] = s / static_cast<double>(m);
    }
    double s = 0.0;
    for (std::size_t r = 0; r < m; ++r) {
      s += dm.y[r];
    }
    ymean = s / static_cast<double>(m);
  }

  std::vector<std::vector<double>> raw(m, std::vector<double>(width));
  for (std::size_t r = 0; r < m; ++r) {
    for (std::size_t j = 0; j < nx; ++j) {
      raw[r][j] = dm.data[r * p + j] - means[j];
    }
    if (dm.with_const) {
      raw[r][ones] = 1.0;
    }
    raw[r][ycol] = dm.y[r] - ymean;
  }
  std::vector<double> orig(ncol, 0.0);
  for (std::size_t c = 0; c < ncol; ++c) {
    for (std::size_t r = 0; r < m; ++r) {
      orig[c] += raw[r][c] * raw[r][c];
    }
  }
  double syy = 0.0;
  for (std::size_t r = 0; r < m; ++r) {
    const double v = dm.y[r] - ymean;
    syy += v * v;
  }

  std::vector<double> remaining = orig;
  std::vector<std::vector<std::pair<std::size_t, double>>> mult;  // per step: (column, multiplier)
  std::vector<std::vector<double>> rmat;                          // per step: R row over all columns
  std::vector<std::size_t> order;
  std::vector<std::size_t> avail(ncol);
  std::vector<std::size_t> phys(ncol);
  for (std::size_t c = 0; c < ncol; ++c) {
    avail[c] = c;
    phys[c] = c;
  }

  const auto build_row = [&](std::size_t r, std::size_t upto) {
    std::vector<double> a = raw[r];
    for (std::size_t i = 0; i < upto; ++i) {
      const double ai = a[order[i]];
      for (const auto& jm : mult[i]) {
        a[jm.first] = a[jm.first] + -(ai * jm.second);
      }
    }
    return a;
  };

  std::size_t k = 0;
  while (!avail.empty() && k < m) {
    std::size_t piv;
    if (dm.with_const && k == 0U) {
      piv = ones;
    } else {
      double best = 0.0;
      piv = phys[k];
      for (std::size_t t = k; t < ncol; ++t) {
        if (remaining[phys[t]] > best) {
          best = remaining[phys[t]];
          piv = phys[t];
        }
      }
      if (remaining[piv] < kPivotFloor * orig[piv]) {
        break;
      }
    }
    for (std::size_t t = 0; t < ncol; ++t) {
      if (phys[t] == piv) {
        std::swap(phys[k], phys[t]);
        break;
      }
    }
    avail.erase(std::find(avail.begin(), avail.end(), piv));
    order.push_back(piv);

    std::vector<std::vector<double>> rows;
    for (std::size_t r = k; r < m; ++r) {
      rows.push_back(build_row(r, k));
    }
    double ssq = 0.0;
    double tail = 0.0;
    for (std::size_t i = 0; i < rows.size(); ++i) {
      const double v = rows[i][piv];
      ssq = ssq + v * v;
      if (i != 0U) {
        tail = tail + v * v;
      }
    }
    const double akk = rows[0][piv];
    const double nrm = std::sqrt(ssq);
    const double beta = akk > 0.0 ? -nrm : nrm;
    const double vk = -beta + akk;
    const double vn = vk * vk + tail;

    std::vector<std::size_t> others = avail;
    others.push_back(ycol);
    std::vector<std::pair<std::size_t, double>> mk;
    std::vector<double> rrow(width, 0.0);
    rrow[piv] = beta;
    for (const std::size_t j : others) {
      double w = vk * rows[0][j];
      for (std::size_t i = 1; i < rows.size(); ++i) {
        w = w + rows[i][piv] * rows[i][j];
      }
      const double mj = div0(2.0 * w, vn);
      mk.emplace_back(j, mj);
      const double rkj = rows[0][j] + -(mj * vk);
      rrow[j] = rkj;
      if (j != ycol) {
        remaining[j] = remaining[j] + -(rkj * rkj);
      }
    }
    mult.push_back(std::move(mk));
    rmat.push_back(std::move(rrow));
    ++k;
  }
  const std::size_t steps = k;

  std::vector<double> b(ncol, 0.0);
  for (std::size_t i = steps; i > 0U; --i) {
    const std::size_t row = i - 1U;
    double t = 0.0;
    for (std::size_t j = row + 1U; j < steps; ++j) {
      t = t + rmat[row][order[j]] * b[order[j]];
    }
    b[order[row]] = div0(rmat[row][ycol] - t, rmat[row][order[row]]);
  }

  double ssr = 0.0;
  for (std::size_t r = steps; r < m; ++r) {
    const double v = build_row(r, steps)[ycol];
    ssr = ssr + v * v;
  }

  out->coeffs.assign(p, 0.0);
  for (std::size_t j = 0; j < nx; ++j) {
    out->coeffs[j] = b[j];
  }
  if (dm.with_const) {
    double a = ymean;
    for (std::size_t j = 0; j < nx; ++j) {
      a = a - b[j] * means[j];
    }
    out->coeffs[nx] = a;
  }
  out->ss_resid = ssr;
  out->ss_total = syy;
  out->df = static_cast<double>(m - steps);
  out->steps = steps;
  if (!want_se) {
    return;
  }

  std::vector<double> rinv(steps * steps, 0.0);
  for (std::size_t c = 0; c < steps; ++c) {
    for (std::size_t i = steps; i > 0U; --i) {
      const std::size_t row = i - 1U;
      double s = (row == c) ? 1.0 : 0.0;
      for (std::size_t j = row + 1U; j < steps; ++j) {
        s = s + -(rmat[row][order[j]] * rinv[j * steps + c]);
      }
      rinv[row * steps + c] = div0(s, rmat[row][order[row]]);
    }
  }
  std::vector<double> v(steps * steps, 0.0);
  for (std::size_t r = 0; r < steps; ++r) {
    for (std::size_t c = 0; c < steps; ++c) {
      double s = 0.0;
      for (std::size_t j = 0; j < steps; ++j) {
        s = s + rinv[r * steps + j] * rinv[c * steps + j];
      }
      v[r * steps + c] = div0(s * ssr, out->df);
    }
  }
  std::vector<int> pos(ncol, -1);
  for (std::size_t i = 0; i < steps; ++i) {
    pos[order[i]] = static_cast<int>(i);
  }
  const auto vc = [&](std::size_t a, std::size_t c) {
    return (pos[a] >= 0 && pos[c] >= 0) ? v[static_cast<std::size_t>(pos[a]) * steps + static_cast<std::size_t>(pos[c])]
                                        : 0.0;
  };
  out->se.assign(p, 0.0);
  for (std::size_t j = 0; j < nx; ++j) {
    out->se[j] = pos[j] >= 0 ? std::sqrt(vc(j, j)) : 0.0;
  }
  if (dm.with_const) {
    double voo = vc(ones, ones);
    for (std::size_t j = 0; j < nx; ++j) {
      voo = voo + (means[j] * means[j]) * vc(j, j);
    }
    for (std::size_t j = 0; j < nx; ++j) {
      for (std::size_t l = j + 1U; l < nx; ++l) {
        voo = voo + ((2.0 * means[j]) * means[l]) * vc(j, l);
      }
    }
    for (std::size_t j = 0; j < nx; ++j) {
      voo = voo - (2.0 * means[j]) * vc(j, ones);
    }
    out->se[nx] = voo >= 0.0 ? std::sqrt(voo) : std::nan("");
  }
}

/// Writes the row-major `m x k` transpose of the `k x m` matrix `src` into `dst`.
void transpose_into(const std::vector<double>& src, std::uint32_t m, std::uint32_t k, std::vector<double>& dst) {
  dst.resize(static_cast<std::size_t>(m) * k);
  for (std::uint32_t i = 0; i < m; ++i) {
    for (std::uint32_t j = 0; j < k; ++j) {
      dst[static_cast<std::size_t>(i) * k + j] = src[static_cast<std::size_t>(j) * m + i];
    }
  }
}

/// Fills the first `k + 1` cells of `cells` with the coefficient row in
/// reverse order. The intercept slot is 0 (LINEST) or 1 (LOGEST) without a
/// constant.
void fill_coeff_row(const std::vector<double>& coeffs, std::uint32_t k, bool with_const, bool log_form, Value* cells) {
  for (std::uint32_t j = 0; j < k; ++j) {
    const double raw = coeffs[j];
    const double v = log_form ? elementary::exp(raw) : raw;
    cells[k - 1U - j] = std::isfinite(v) ? Value::number(v) : Value::error(ErrorCode::Num);
  }
  if (with_const) {
    const double raw = coeffs[k];
    const double v = log_form ? elementary::exp(raw) : raw;
    cells[k] = std::isfinite(v) ? Value::number(v) : Value::error(ErrorCode::Num);
  } else {
    cells[k] = Value::number(log_form ? 1.0 : 0.0);
  }
}

/// Builds the design matrix from a known_y array and an optional
/// known_x array. `known_y` may be 1 x m or m x 1; the orientation
/// determines how `known_x` is read. If `x_arr` is `nullptr`, the
/// predictor defaults to the index sequence `{1, 2, ..., m}`.
///
/// Returns `false` on a non-numeric y cell (`#VALUE!`), then on a shape
/// mismatch (`#REF!`), then on a non-numeric x cell (`#VALUE!`).
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

  // y's content is checked before x's shape, x's shape before its content.
  std::vector<double> y_vec;
  if (!coerce_fit_matrix(y_arr, y_vec, *out_err)) {
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
    // Match orientation against y. The "observation axis" of x must
    // equal m; the other axis becomes k. Mac Excel accepts either layout.
    bool transpose = false;
    if (y_is_col && x_arr->rows == m) {
      k = x_arr->cols;  // m x k, the canonical column-y layout
    } else if (y_is_col && x_arr->cols == m) {
      k = x_arr->rows;  // k x m (1 x m for a single variable)
      transpose = k > 1U;
    } else if (!y_is_col && x_arr->cols == m) {
      k = x_arr->rows;  // k x m, predictors stacked
      transpose = true;
    } else if (!y_is_col && x_arr->rows == m) {
      k = x_arr->cols;  // m x k (m x 1 for a single variable)
    } else {
      *out_err = Value::error(ErrorCode::Ref);
      return false;
    }
    std::vector<double> x_vec;
    if (!coerce_fit_matrix(*x_arr, x_vec, *out_err)) {
      return false;
    }
    if (transpose) {
      transpose_into(x_vec, m, k, x_data);
    } else {
      x_data = std::move(x_vec);
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
    const double a = with_const ? coeffs[k] : 0.0;
    double v = exp_result ? 1.0 : 0.0;
    for (std::uint32_t j = 0; j < k; ++j) {
      const double t = coeffs[j] * pred_design[static_cast<std::size_t>(i) * p + j];
      v = exp_result ? v * elementary::exp(t) : v + t;
    }
    v = exp_result ? v * elementary::exp(a) : v + a;
    cells.push_back(std::isfinite(v) ? Value::number(v) : Value::error(ErrorCode::Num));
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
  fill_coeff_row(coeffs, k, with_const, log_form, cells.data());
  return array_from_values(1U, out_cols, cells.data(), cells.size(), arena);
}

/// Builds the 5x(k+1) statistics matrix from the fit. The reverse-ordering
/// and the `#N/A` padding rules mirror `build_simple_output`. When
/// `log_form=true`, only row 1 (coefficients) is exponentiated; rows 2-5
/// stay on the log-domain scale because they describe the linearised model.
ArrayValue* build_stats_output(const LinFit& fit, const DesignMatrix& dm, Arena& arena, bool log_form = false) {
  const std::vector<double>& coeffs = fit.coeffs;
  const std::uint32_t k = dm.k;
  const std::uint32_t out_cols = k + 1U;
  const std::uint32_t out_rows = 5U;
  std::vector<Value> cells(static_cast<std::size_t>(out_rows) * out_cols, Value::error(ErrorCode::NA));
  const auto num_or_na = [](double v) { return std::isfinite(v) ? Value::number(v) : Value::error(ErrorCode::NA); };

  // Row 1: coefficients in reverse order. Intercept slot is 0 (LINEST)
  // or 1 (LOGEST) when const=false.
  fill_coeff_row(coeffs, k, dm.with_const, log_form, cells.data());

  // Row 2: standard errors; the intercept slot is #N/A without a constant.
  for (std::uint32_t j = 0; j < k; ++j) {
    cells[out_cols + (k - 1U - j)] = num_or_na(fit.se[j]);
  }
  if (dm.with_const) {
    cells[out_cols + k] = num_or_na(fit.se[k]);
  }

  const double ss_reg = fit.ss_total - fit.ss_resid;
  const double sigma2 = div0(fit.ss_resid, fit.df);
  // Row 3: [r^2, se_y, #N/A, ...].
  if (fit.ss_total != 0.0) {
    cells[2U * out_cols + 0U] = num_or_na(ss_reg / fit.ss_total);
  }
  cells[2U * out_cols + 1U] = num_or_na(std::sqrt(sigma2));

  // Row 4: [F, df_resid, #N/A, ...]; F divides by the residual variance, so
  // a zero residual is #NUM!.
  const double kx = static_cast<double>(fit.steps) - (dm.with_const ? 1.0 : 0.0);
  if (fit.df > 0.0 && kx > 0.0) {
    cells[3U * out_cols + 0U] =
        fit.ss_resid == 0.0 ? Value::error(ErrorCode::Num) : num_or_na((ss_reg / kx) / (fit.ss_resid / fit.df));
  }
  cells[3U * out_cols + 1U] = Value::number(fit.df);

  // Row 5: [ss_reg, ss_resid, #N/A, ...].
  cells[4U * out_cols + 0U] = num_or_na(ss_reg);
  cells[4U * out_cols + 1U] = num_or_na(fit.ss_resid);

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
    dm.y[i] = elementary::log(dm.y[i]);
  }
  return true;
}

/// Reads the `known_y, [known_x]` head all four regression builtins
/// take, along with the arity bound they share. `*out_x` stays null when
/// `known_x` was omitted, which the design-matrix builder reads as the
/// implicit `{1, 2, 3, ...}` predictor.
///
/// Returns `false` with the error in `*out_err`: an argument that evaluates to
/// an error is `#VALUE!`.
bool read_known_xy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                   const ArrayValue** out_y, const ArrayValue** out_x, Value* out_err) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 1U || arity > 4U) {
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  if (!resolve_array_value(call.as_call_arg(0), arena, registry, ctx, out_y, out_err) ||
      (arity >= 2U && !resolve_array_value(call.as_call_arg(1), arena, registry, ctx, out_x, out_err))) {
    *out_err = Value::error(ErrorCode::Value);
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
  const Value flags[2] = {flag_arg_value(call, 2U, arena, registry, ctx, true),
                          flag_arg_value(call, 3U, arena, registry, ctx, false)};
  return lift_flags(flags, 2U, arena, [&](const bool* values) {
    const bool with_const = values[0];
    const bool want_stats = values[1];
    DesignMatrix dm;
    LinFit fit;
    Value fit_err = Value::blank();
    if (!fit_coefficients(*y_arr, x_arr, with_const, log_form, /*want_inv=*/want_stats, &dm, &fit, &fit_err)) {
      return fit_err;
    }
    ArrayValue* arr = want_stats ? build_stats_output(fit, dm, arena, log_form)
                                 : build_simple_output(fit.coeffs, dm.k, with_const, arena, log_form);
    return arr == nullptr ? Value::error(ErrorCode::Num) : Value::array(arr);
  });
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
/// Returns `false` on shape mismatch (`#REF!`), then on a non-numeric
/// cell, an error included (`#VALUE!`).
bool resolve_new_x_design(const ArrayValue& new_x, const DesignMatrix& dm, bool y_is_col, bool with_const,
                          std::vector<double>* out_design, std::uint32_t* out_n, Value* out_err) {
  const std::uint32_t k = dm.k;
  std::uint32_t n = 0;
  bool transpose = false;
  if (y_is_col && new_x.cols == k) {
    n = new_x.rows;  // already N x k row-major
  } else if (y_is_col && k == 1U && new_x.rows == 1U) {
    n = new_x.cols;  // single-variable fallback: 1 x N works as N x 1
  } else if (!y_is_col && new_x.rows == k) {
    n = new_x.cols;  // k x N -> transpose to N x k
    transpose = true;
  } else if (!y_is_col && k == 1U && new_x.cols == 1U) {
    n = new_x.rows;  // single-variable fallback: N x 1 works as 1 x N
  } else {
    *out_err = Value::error(ErrorCode::Ref);
    return false;
  }
  std::vector<double> raw;
  if (!coerce_fit_matrix(new_x, raw, *out_err)) {
    return false;
  }
  std::vector<double> features;  // N x k row-major
  if (transpose) {
    transpose_into(raw, n, k, features);
  } else {
    features = std::move(raw);
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
  const ArrayValue* new_x_arr = nullptr;
  if (call_arg_supplied(call, 2U) &&
      !resolve_array_value(call.as_call_arg(2), arena, registry, ctx, &new_x_arr, &err)) {
    return Value::error(ErrorCode::Value);
  }
  const Value flags[1] = {flag_arg_value(call, 3U, arena, registry, ctx, true)};
  return lift_flags(flags, 1U, arena, [&](const bool* values) {
    const bool with_const = values[0];
    DesignMatrix dm;
    LinFit fit;
    Value fit_err = Value::blank();
    if (!fit_coefficients(*y_arr, x_arr, with_const, log_form, /*want_inv=*/false, &dm, &fit, &fit_err)) {
      return fit_err;
    }
    const std::uint32_t p = dm.k + (with_const ? 1U : 0U);
    // Without new_x the prediction runs over the training design itself.
    const bool y_is_col = y_arr->cols == 1U && y_arr->rows >= 1U;
    std::vector<double> pred_design;
    std::uint32_t n_obs = 0;
    if (new_x_arr == nullptr) {
      n_obs = dm.m;
      pred_design = std::move(dm.data);
    } else if (!resolve_new_x_design(*new_x_arr, dm, y_is_col, with_const, &pred_design, &n_obs, &fit_err)) {
      return fit_err;
    }
    ArrayValue* arr = build_prediction_output(fit.coeffs, pred_design, n_obs, p, with_const, y_is_col,
                                              /*exp_result=*/log_form, arena);
    return arr == nullptr ? Value::error(ErrorCode::Num) : Value::array(arr);
  });
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
