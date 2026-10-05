//
// Holt-Winters fit for the FORECAST.ETS family; the contract is declared in
// `eval/holt_winters.h`.

#include "eval/holt_winters.h"

#include <algorithm>
#include <array>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <utility>
#include <vector>

#include "utils/error.h"
#include "utils/index_sort.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace ets {
namespace {

// Nelder-Mead bounds clamp -- coordinates are squeezed into
// (kBoundEps, 1 - kBoundEps) so the gradient never escapes the unit
// cube and the smoothing recurrences stay well-defined.
constexpr double kBoundEps = 1e-6;

// Nelder-Mead convergence: stop when (fmax - fmin) over the simplex falls
// below this, or when iteration count exceeds kMaxNelderMeadIters.
constexpr double kNelderMeadFTol = 1e-9;
constexpr int kMaxNelderMeadIters = 200;

// Runs the Holt-Winters smoothing recurrences over `y` with explicit
// initial state (`init_level`, `init_trend`, `init_season[0..m-1]`),
// populates `*out` (level / trend / season / residuals), and returns the
// in-sample SSE used by the optimiser.
//
// Recurrences (additive seasonality):
//   L_t = alpha (y_t - S_{t-m}) + (1-alpha)(L_{t-1} + B_{t-1})
//   B_t = beta  (L_t - L_{t-1}) + (1-beta)  B_{t-1}
//   S_t = gamma (y_t - L_t)     + (1-gamma) S_{t-m}
// Forecast residual at step t is `y_t - (L_{t-1} + B_{t-1} + S_{t-m})`.
//
// `init_season` may be null when `m == 1`.
double simulate_with_init(const std::vector<double>& y, std::uint32_t m, double alpha, double beta, double gamma,
                          double init_level, double init_trend, const double* init_season, HoltWintersFit* out) {
  const std::size_t n = y.size();
  double level = init_level;
  double trend = init_trend;
  std::vector<double> season(m == 1U ? 0U : static_cast<std::size_t>(m), 0.0);
  if (m != 1U && init_season != nullptr) {
    for (std::uint32_t i = 0; i < m; ++i) {
      season[i] = init_season[i];
    }
  }

  std::vector<double> residuals(n, 0.0);
  double sse = 0.0;
  // The first `start` indices have no one-step-ahead forecast (no prior
  // level / trend / seasonal estimate is available beneath them).
  const std::size_t start = (m == 1U) ? 1U : static_cast<std::size_t>(m);
  for (std::size_t t = start; t < n; ++t) {
    const double prev_level = level;
    const double prev_trend = trend;
    const double s_prev = (m == 1U) ? 0.0 : season[t % m];
    const double y_hat = prev_level + prev_trend + s_prev;
    const double err = y[t] - y_hat;
    residuals[t] = err;
    sse += err * err;
    // Update equations.
    level = alpha * (y[t] - s_prev) + (1.0 - alpha) * (prev_level + prev_trend);
    trend = beta * (level - prev_level) + (1.0 - beta) * prev_trend;
    if (m != 1U) {
      season[t % m] = gamma * (y[t] - level) + (1.0 - gamma) * s_prev;
    }
  }

  if (out != nullptr) {
    out->alpha = alpha;
    out->beta = beta;
    out->gamma = (m == 1U) ? 0.0 : gamma;
    out->level = level;
    out->trend = trend;
    out->season = season;
    out->m = m;
    out->residuals = std::move(residuals);
  }
  return sse;
}

// Computes the textbook two-season-means default initial state used as the
// starting point of the optimiser. With `m == 1` we use Holt's double
// exponential: L_0 = y_0, B_0 = y_1 - y_0 (no seasonal vector).
void compute_default_init(const std::vector<double>& y, std::uint32_t m, double* out_level, double* out_trend,
                          std::vector<double>* out_season) {
  const std::size_t n = y.size();
  out_season->assign(m == 1U ? 0U : static_cast<std::size_t>(m), 0.0);
  if (m == 1U) {
    *out_level = y[0];
    *out_trend = (n >= 2U) ? (y[1] - y[0]) : 0.0;
    return;
  }
  double mean1 = 0.0;
  for (std::uint32_t i = 0; i < m; ++i)
    mean1 += y[i];
  mean1 /= static_cast<double>(m);
  double mean2 = mean1;
  if (n >= 2U * static_cast<std::size_t>(m)) {
    mean2 = 0.0;
    for (std::uint32_t i = 0; i < m; ++i)
      mean2 += y[m + i];
    mean2 /= static_cast<double>(m);
  }
  *out_level = mean1;
  *out_trend = (mean2 - mean1) / static_cast<double>(m);
  double s_sum = 0.0;
  for (std::uint32_t i = 0; i < m; ++i) {
    (*out_season)[i] = y[i] - mean1;
    s_sum += (*out_season)[i];
  }
  const double s_mean = s_sum / static_cast<double>(m);
  for (std::uint32_t i = 0; i < m; ++i)
    (*out_season)[i] -= s_mean;
}

// Backwards-compatible wrapper: runs `simulate_with_init` from the textbook
// two-season-means default initialisation.
double simulate(const std::vector<double>& y, std::uint32_t m, double alpha, double beta, double gamma,
                HoltWintersFit* out) {
  double init_level = 0.0;
  double init_trend = 0.0;
  std::vector<double> init_season;
  compute_default_init(y, m, &init_level, &init_trend, &init_season);
  return simulate_with_init(y, m, alpha, beta, gamma, init_level, init_trend,
                            init_season.empty() ? nullptr : init_season.data(), out);
}

// ---------------------------------------------------------------------------
// Bounded Nelder-Mead in 2D / 3D
// ---------------------------------------------------------------------------

struct ParamBounds {
  double lo;
  double hi;
};

// Clamps each coordinate of `v` (length `dim`) into its per-coordinate bound.
void clamp_vertex_bounded(double* v, std::size_t dim, const ParamBounds* bounds) noexcept {
  for (std::size_t i = 0; i < dim; ++i) {
    if (v[i] < bounds[i].lo)
      v[i] = bounds[i].lo;
    if (v[i] > bounds[i].hi)
      v[i] = bounds[i].hi;
  }
}

// Expands a packed parameter vector of length `dim` into (alpha, beta, gamma).
// `dim == 2` => non-seasonal: gamma is forced to 0.
// `dim == 3` => seasonal: all three are read out.
void unpack_params(const double* p, std::size_t dim, double* alpha, double* beta, double* gamma) noexcept {
  *alpha = p[0];
  *beta = p[1];
  *gamma = (dim == 3U) ? p[2] : 0.0;
}

// Objective evaluated by the simplex search. `smooth_dim` is only meaningful
// to the joint objective; the smoothing-only one ignores it.
using SimplexObjective = double (*)(const double* params, std::size_t dim, std::size_t smooth_dim,
                                    const std::vector<double>& y, std::uint32_t m);

// SSE objective for the smoothing-only optimiser. Returns +inf for non-finite
// SSE (which guides the simplex away from pathological corners).
double objective(const double* params, std::size_t dim, std::size_t /*smooth_dim*/, const std::vector<double>& y,
                 std::uint32_t m) {
  double alpha = 0.0;
  double beta = 0.0;
  double gamma = 0.0;
  unpack_params(params, dim, &alpha, &beta, &gamma);
  const double sse = simulate(y, m, alpha, beta, gamma, nullptr);
  if (!std::isfinite(sse)) {
    return std::numeric_limits<double>::infinity();
  }
  return sse;
}

// Bounded Nelder-Mead over `dim` coordinates, minimising `evaluate`. The
// starting simplex is `seed` plus one `seed_step` perturbation per axis; every
// vertex is clamped into `bounds` after each move. Writes the best vertex into
// `*out_params` and returns its objective value.
//
// The simplex coefficients are the canonical (Nelder-Mead 1965) values:
//   reflect rho = 1, expand chi = 2, contract psi = 0.5, shrink sigma = 0.5.
double nelder_mead_search(const std::vector<double>& y, std::uint32_t m, std::size_t dim, std::size_t smooth_dim,
                          const double* seed, double seed_step, const std::vector<ParamBounds>& bounds, int iter_cap,
                          SimplexObjective evaluate, std::vector<double>* out_params) {
  const std::size_t k = dim + 1U;
  std::vector<std::vector<double>> simplex(k, std::vector<double>(dim, 0.0));
  // Vertex 0: anchor.
  for (std::size_t j = 0; j < dim; ++j) {
    simplex[0][j] = seed[j];
  }
  // Vertices 1..dim: perturb the i-th axis.
  for (std::size_t i = 1; i < k; ++i) {
    for (std::size_t j = 0; j < dim; ++j) {
      simplex[i][j] = simplex[0][j];
    }
    simplex[i][i - 1U] += seed_step;
  }
  for (auto& v : simplex) {
    clamp_vertex_bounded(v.data(), dim, bounds.data());
  }

  std::vector<double> fvals(k, 0.0);
  for (std::size_t i = 0; i < k; ++i) {
    fvals[i] = evaluate(simplex[i].data(), dim, smooth_dim, y, m);
  }

  for (int iter = 0; iter < iter_cap; ++iter) {
    // Sort vertices by function value (ascending).
    std::vector<std::size_t> order(k);
    for (std::size_t i = 0; i < k; ++i)
      order[i] = i;
    std::sort(order.begin(), order.end(),
              [&fvals](std::size_t a, std::size_t b) noexcept { return fvals[a] < fvals[b]; });
    const std::size_t best = order[0];
    const std::size_t worst = order[k - 1U];
    const std::size_t second_worst = order[k - 2U];

    // Convergence check: span of f-values across the simplex.
    if (std::isfinite(fvals[best]) && std::isfinite(fvals[worst]) && (fvals[worst] - fvals[best]) < kNelderMeadFTol) {
      break;
    }

    // Centroid of the best-`dim` vertices (i.e. all except the worst).
    std::vector<double> centroid(dim, 0.0);
    for (std::size_t i = 0; i < k; ++i) {
      if (i == worst)
        continue;
      for (std::size_t j = 0; j < dim; ++j)
        centroid[j] += simplex[i][j];
    }
    for (std::size_t j = 0; j < dim; ++j)
      centroid[j] /= static_cast<double>(dim);

    // Reflect.
    std::vector<double> reflected(dim);
    for (std::size_t j = 0; j < dim; ++j) {
      reflected[j] = centroid[j] + 1.0 * (centroid[j] - simplex[worst][j]);
    }
    clamp_vertex_bounded(reflected.data(), dim, bounds.data());
    const double f_reflected = evaluate(reflected.data(), dim, smooth_dim, y, m);

    if (f_reflected < fvals[second_worst] && f_reflected >= fvals[best]) {
      simplex[worst] = reflected;
      fvals[worst] = f_reflected;
      continue;
    }

    if (f_reflected < fvals[best]) {
      // Expand.
      std::vector<double> expanded(dim);
      for (std::size_t j = 0; j < dim; ++j) {
        expanded[j] = centroid[j] + 2.0 * (reflected[j] - centroid[j]);
      }
      clamp_vertex_bounded(expanded.data(), dim, bounds.data());
      const double f_expanded = evaluate(expanded.data(), dim, smooth_dim, y, m);
      if (f_expanded < f_reflected) {
        simplex[worst] = expanded;
        fvals[worst] = f_expanded;
      } else {
        simplex[worst] = reflected;
        fvals[worst] = f_reflected;
      }
      continue;
    }

    // Contract.
    std::vector<double> contracted(dim);
    if (f_reflected < fvals[worst]) {
      // Outside contraction.
      for (std::size_t j = 0; j < dim; ++j) {
        contracted[j] = centroid[j] + 0.5 * (reflected[j] - centroid[j]);
      }
    } else {
      // Inside contraction.
      for (std::size_t j = 0; j < dim; ++j) {
        contracted[j] = centroid[j] + 0.5 * (simplex[worst][j] - centroid[j]);
      }
    }
    clamp_vertex_bounded(contracted.data(), dim, bounds.data());
    const double f_contracted = evaluate(contracted.data(), dim, smooth_dim, y, m);
    if (f_contracted < std::min(f_reflected, fvals[worst])) {
      simplex[worst] = contracted;
      fvals[worst] = f_contracted;
      continue;
    }

    // Shrink toward best.
    for (std::size_t i = 0; i < k; ++i) {
      if (i == best)
        continue;
      for (std::size_t j = 0; j < dim; ++j) {
        simplex[i][j] = simplex[best][j] + 0.5 * (simplex[i][j] - simplex[best][j]);
      }
      clamp_vertex_bounded(simplex[i].data(), dim, bounds.data());
      fvals[i] = evaluate(simplex[i].data(), dim, smooth_dim, y, m);
    }
  }

  // Return the best vertex.
  std::size_t best_idx = 0;
  for (std::size_t i = 1; i < k; ++i) {
    if (fvals[i] < fvals[best_idx])
      best_idx = i;
  }
  *out_params = simplex[best_idx];
  return fvals[best_idx];
}

// Runs the search over the unit cube of dimension `dim` (2 or 3), tuning the
// smoothing parameters alone. Writes the best (alpha, beta, gamma) into the
// out-pointers and returns the best SSE.
double nelder_mead(const std::vector<double>& y, std::uint32_t m, std::size_t dim, double* out_alpha, double* out_beta,
                   double* out_gamma) {
  // Initial simplex: (0.2, 0.1, 0.1) plus unit-perturbations of magnitude
  // 0.1 in each axis. Comfortably inside the bounds.
  const std::array<double, 3> seed{0.2, 0.1, 0.1};
  const std::vector<ParamBounds> bounds(dim, ParamBounds{kBoundEps, 1.0 - kBoundEps});
  std::vector<double> best;
  const double sse = nelder_mead_search(y, m, dim, /*smooth_dim=*/dim, seed.data(), 0.1, bounds, kMaxNelderMeadIters,
                                        &objective, &best);
  unpack_params(best.data(), dim, out_alpha, out_beta, out_gamma);
  return sse;
}

// ---------------------------------------------------------------------------
// Joint (alpha, beta, gamma, L_0, B_0, S_0..S_{m-1}) optimisation
// ---------------------------------------------------------------------------
//
// The 3D `nelder_mead` above only tunes the smoothing parameters at a fixed
// two-season-means initial state. On strongly cyclic data Mac Excel reaches
// a much tighter fit by jointly tuning the initial state vector along with
// the smoothing parameters; the joint optimum can sit in a basin our 3D
// optimiser cannot reach. The generalised optimiser below packs:
//
//   p = [alpha, beta[, gamma], L_0, B_0, S_0, ..., S_{m-1}]
//
// where the seasonal block is omitted when `m == 1`. The smoothing block is
// clamped to (kBoundEps, 1 - kBoundEps); the initial-state block is bounded
// to a generous data-range envelope so the simplex can move freely without
// running away to non-finite SSE.

// Computes the per-coordinate bounds used by `nelder_mead_full`. The
// smoothing-param block (the first `smooth_dim` coordinates) is `(eps,
// 1 - eps)`; the initial-state block (the remaining `dim - smooth_dim`
// coordinates) is `[y_min - 5 * range, y_max + 5 * range]` so the simplex
// has room to escape its starting basin.
std::vector<ParamBounds> make_full_bounds(const std::vector<double>& y, std::size_t dim, std::size_t smooth_dim) {
  std::vector<ParamBounds> bounds(dim);
  for (std::size_t i = 0; i < smooth_dim; ++i) {
    bounds[i] = ParamBounds{kBoundEps, 1.0 - kBoundEps};
  }
  double y_min = y[0];
  double y_max = y[0];
  for (double v : y) {
    if (v < y_min)
      y_min = v;
    if (v > y_max)
      y_max = v;
  }
  const double range = std::max(y_max - y_min, 1.0);
  const double lo = y_min - 5.0 * range;
  const double hi = y_max + 5.0 * range;
  for (std::size_t i = smooth_dim; i < dim; ++i) {
    bounds[i] = ParamBounds{lo, hi};
  }
  return bounds;
}

// SSE objective for the joint optimiser. Reads packed params:
//   smooth_dim == 2 (m == 1): [alpha, beta, L_0, B_0]
//   smooth_dim == 3 (m >= 2): [alpha, beta, gamma, L_0, B_0, S_0..S_{m-1}]
double objective_full(const double* p, std::size_t /*dim*/, std::size_t smooth_dim, const std::vector<double>& y,
                      std::uint32_t m) {
  const double alpha = p[0];
  const double beta = p[1];
  const double gamma = (smooth_dim == 3U) ? p[2] : 0.0;
  const double init_level = p[smooth_dim];
  const double init_trend = p[smooth_dim + 1U];
  const double* init_season = (m >= 2U) ? &p[smooth_dim + 2U] : nullptr;
  const double sse = simulate_with_init(y, m, alpha, beta, gamma, init_level, init_trend, init_season, nullptr);
  return std::isfinite(sse) ? sse : std::numeric_limits<double>::infinity();
}

// Joint optimisation over the smoothing parameters and the initial state.
// Only the knobs differ from the smoothing-only search: the simplex is built
// around the supplied `seed` (length `dim`) instead of the fixed anchor, the
// bounds are per-coordinate, and the iteration budget grows with `dim` — a 9D
// landscape needs ~2-3x more iterations than the 3D one to converge through
// the same number of simplex collapses.
double nelder_mead_full(const std::vector<double>& y, std::uint32_t m, std::size_t dim, std::size_t smooth_dim,
                        const double* seed, double seed_step, const std::vector<ParamBounds>& bounds,
                        std::vector<double>* out_params) {
  const int iter_cap = static_cast<int>(static_cast<std::size_t>(kMaxNelderMeadIters) * std::max<std::size_t>(1U, dim));
  return nelder_mead_search(y, m, dim, smooth_dim, seed, seed_step, bounds, iter_cap, &objective_full, out_params);
}

}  // namespace

// ---------------------------------------------------------------------------
// Fit + diagnostics on a resampled, deduplicated, evenly spaced series
// ---------------------------------------------------------------------------

// Fits a Holt-Winters model to `y` with the given seasonality `m`. On
// success populates `*out` (parameters, final state, residuals, MAE,
// RMSE, MASE, SMAPE) and returns `true`. Writes `*out_err` and returns
// `false` on degenerate inputs (n < 2, or seasonal m with n < 2*m).
bool fit_holt_winters(const std::vector<double>& y, std::uint32_t m, HoltWintersFit* out, Value* out_err) {
  const std::size_t n = y.size();
  if (n < 2U) {
    *out_err = Value::error(ErrorCode::NA);
    return false;
  }
  if (m >= 2U && n < 2U * static_cast<std::size_t>(m)) {
    *out_err = Value::error(ErrorCode::NA);
    return false;
  }
  const std::size_t smooth_dim = (m >= 2U) ? 3U : 2U;
  double alpha = 0.0;
  double beta = 0.0;
  double gamma = 0.0;

  // Stage 1: tune (alpha, beta[, gamma]) at the textbook two-season-means
  // initialisation. This is the conservative path that has shipped since
  // FORECAST.ETS landed; on most series it sits at or near the global
  // optimum and is cheap to compute.
  const double sse_basic = nelder_mead(y, m, smooth_dim, &alpha, &beta, &gamma);

  // Stage 2: jointly tune the smoothing parameters and the initial state
  // (L_0, B_0, S_0..S_{m-1}). The 3D optimum can sit in a basin that the
  // smooth-only stage cannot reach because the initial state is fixed; on
  // strongly cyclic data this stage finds the basin Mac Excel converges
  // to. Multi-start over a small seed grid widens basin coverage further.
  double init_level = 0.0;
  double init_trend = 0.0;
  std::vector<double> init_season;
  compute_default_init(y, m, &init_level, &init_trend, &init_season);
  const std::size_t full_dim = smooth_dim + 2U + ((m >= 2U) ? static_cast<std::size_t>(m) : 0U);
  const std::vector<ParamBounds> full_bounds = make_full_bounds(y, full_dim, smooth_dim);

  // Seed grid. Each entry supplies a (alpha, beta[, gamma]) anchor; the
  // initial-state block is always seeded from the textbook default so the
  // simplex can move it as needed. Seeds cover low/mid/high alpha to
  // hedge against multimodal SSE landscapes.
  static constexpr double kSmoothSeeds[][3] = {
      {0.2, 0.1, 0.1},
      {0.5, 0.1, 0.1},
      {0.8, 0.05, 0.1},
      {0.5, 0.5, 0.5},
  };
  double sse_full = std::numeric_limits<double>::infinity();
  std::vector<double> full_params;
  for (const auto& smooth_seed : kSmoothSeeds) {
    std::vector<double> seed(full_dim, 0.0);
    seed[0] = smooth_seed[0];
    seed[1] = smooth_seed[1];
    if (smooth_dim == 3U)
      seed[2] = smooth_seed[2];
    seed[smooth_dim] = init_level;
    seed[smooth_dim + 1U] = init_trend;
    if (m >= 2U) {
      for (std::uint32_t i = 0; i < m; ++i) {
        seed[smooth_dim + 2U + i] = init_season[i];
      }
    }
    std::vector<double> params;
    const double sse = nelder_mead_full(y, m, full_dim, smooth_dim, seed.data(),
                                        /*seed_step=*/0.1, full_bounds, &params);
    if (sse < sse_full) {
      sse_full = sse;
      full_params = std::move(params);
    }
  }

  // Pick the better fit. The two stages share an objective (SSE on the
  // same series), so a direct comparison is sound. On weakly-structured
  // data the two-stage paths agree to numerical noise; on strongly-cyclic
  // data the joint fit is materially better.
  if (sse_full + 1e-9 < sse_basic && std::isfinite(sse_full)) {
    alpha = full_params[0];
    beta = full_params[1];
    gamma = (smooth_dim == 3U) ? full_params[2] : 0.0;
    init_level = full_params[smooth_dim];
    init_trend = full_params[smooth_dim + 1U];
    if (m >= 2U) {
      for (std::uint32_t i = 0; i < m; ++i) {
        init_season[i] = full_params[smooth_dim + 2U + i];
      }
    }
    (void)simulate_with_init(y, m, alpha, beta, gamma, init_level, init_trend,
                             init_season.empty() ? nullptr : init_season.data(), out);
  } else {
    // Re-run simulation at the fitted parameters to populate residuals
    // and final state.
    (void)simulate(y, m, alpha, beta, gamma, out);
  }

  // Diagnostic statistics. `start` = first index that has a forecast.
  const std::size_t start = (m == 1U) ? 1U : static_cast<std::size_t>(m);
  if (start >= n) {
    *out_err = Value::error(ErrorCode::NA);
    return false;
  }
  const std::size_t k = n - start;
  double sum_abs_err = 0.0;
  double sum_sq_err = 0.0;
  double smape_num = 0.0;
  std::size_t smape_n = 0U;
  for (std::size_t t = start; t < n; ++t) {
    const double err = out->residuals[t];
    sum_abs_err += std::fabs(err);
    sum_sq_err += err * err;
    const double y_hat = y[t] - err;
    const double denom = std::fabs(y[t]) + std::fabs(y_hat);
    if (denom > 0.0) {
      // SMAPE in the symmetric form: |F - y| / ((|F| + |y|) / 2).
      smape_num += std::fabs(err) / (denom * 0.5);
      ++smape_n;
    }
  }
  out->mae = sum_abs_err / static_cast<double>(k);
  out->rmse = std::sqrt(sum_sq_err / static_cast<double>(k));
  out->smape = (smape_n > 0U) ? (smape_num / static_cast<double>(smape_n)) : 0.0;

  // MASE: mean absolute error / mean absolute first-difference of the
  // training series.
  double sum_diff = 0.0;
  for (std::size_t i = 1; i < n; ++i) {
    sum_diff += std::fabs(y[i] - y[i - 1U]);
  }
  if (sum_diff > 0.0) {
    const double scale = sum_diff / static_cast<double>(n - 1U);
    out->mase = out->mae / scale;
  } else {
    out->mase = 0.0;
  }
  return true;
}

}  // namespace ets
}  // namespace eval
}  // namespace formulon
