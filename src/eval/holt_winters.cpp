//
// Holt-Winters fit for the FORECAST.ETS family; the contract is declared in
// `eval/holt_winters.h`.
//
// The smoothing parameters come from a fixed candidate table evaluated in
// order (first strict minimum wins, stopping at an exact zero SSE). Long
// series are fitted on a leading training part and scored over the rest with
// the state frozen. Every arithmetic step is a separate scalar double
// operation; multiply and divide flush subnormal results to zero.

#include "eval/holt_winters.h"

#include <algorithm>
#include <cfloat>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <utility>
#include <vector>

#include "utils/error.h"
#include "utils/subnormal_flush.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace ets {
namespace {

struct SeasonalCandidate {
  double alpha;
  double beta;
  double gamma;
};

struct TrendCandidate {
  double alpha;
  double beta;
};

// Smoothing parameter candidates for seasonal fits.
constexpr SeasonalCandidate kSeasonalCandidates[] = {
    {0.751, 0.001, 0.001},
    {0.501, 0.001, 0.001},
    {0.251, 0.001, 0.001},
    {0.126, 0.001, 0.001},
    {0.998, 0.099, 0.001},
    {0.126, 0.125, 0.001},
    {0.998, 0.249, 0.001},
    {0.5, 0.25, 0.001},
    {0.251, 0.25, 0.001},
    {0.998, 0.499, 0.001},
    {0.501, 0.5, 0.001},
    {0.998, 0.749, 0.001},
    {0.751, 0.75, 0.001},
    {0.998, 0.899, 0.001},
    {0.833333, 0.0833333, 0.0833333},
    {0.166667, 0.0833333, 0.0833333},
    {0.833333, 0.75, 0.0833333},
    {0.9, 0.001, 0.099},
    {0.9, 0.899, 0.099},
    {0.002, 0.001, 0.125},
    {0.666667, 0.166667, 0.166667},
    {0.333333, 0.166667, 0.166667},
    {0.666667, 0.5, 0.166667},
    {0.75, 0.001, 0.249},
    {0.75, 0.749, 0.249},
    {0.251, 0.001, 0.25},
    {0.002, 0.001, 0.25},
    {0.5, 0.25, 0.25},
    {0.251, 0.25, 0.25},
    {0.5, 0.001, 0.499},
    {0.5, 0.499, 0.499},
    {0.002, 0.001, 0.5},
    {0.333333, 0.166667, 0.5},
    {0.25, 0.001, 0.749},
    {0.25, 0.249, 0.749},
    {0.002, 0.001, 0.75},
    {0.166667, 0.0833333, 0.75},
    {0.1, 0.001, 0.899},
    {0.1, 0.099, 0.899},
};

// Smoothing parameter candidates for non-seasonal fits.
constexpr TrendCandidate kTrendCandidates[] = {
    {0.9, 0.001},   {0.75, 0.001},  {0.5, 0.001},          {0.25, 0.001},         {0.126, 0.001},
    {0.1, 0.001},   {0.002, 0.001}, {0.833333, 0.0833333}, {0.166667, 0.0833333}, {0.998, 0.099},
    {0.1, 0.099},   {0.126, 0.125}, {0.666667, 0.166667},  {0.333333, 0.166667},  {0.998, 0.249},
    {0.25, 0.249},  {0.5, 0.25},    {0.998, 0.499},        {0.5, 0.499},          {0.666667, 0.5},
    {0.998, 0.749}, {0.75, 0.749},  {0.833333, 0.75},      {0.998, 0.899},        {0.9, 0.899},
};

double mul(double a, double b) noexcept {
  return flush_subnormal(a * b);
}

double div(double a, double b) noexcept {
  return flush_subnormal(a / b);
}

// Round half away from zero.
double round_away(double v) noexcept {
  return v >= 0.0 ? std::floor(v + 0.5) : -std::floor(-v + 0.5);
}

// Least-squares slope of `y[0..n)` against 1..n. With `state`, the seasonal
// slot for point i is subtracted first.
double ols_slope(const double* y, std::size_t n, const double* state, std::size_t m) {
  if (n == 1U) {
    return 0.0;
  }
  double sy = 0.0;
  double sxy = 0.0;
  for (std::size_t i = 0; i < n; ++i) {
    double v = y[i];
    if (state != nullptr) {
      v = v - state[m + 1U - (i % m)];
    }
    sy = sy + v;
    sxy = mul(static_cast<double>(i + 1U), v) + sxy;
  }
  const double dn = static_cast<double>(n);
  double sx = dn + 1.0;
  sx = mul(sx, dn);
  sx = div(sx, 2.0);
  sx = round_away(sx);
  const double t1 = dn + 1.0;
  const double t2 = dn + 0.5;
  double sxx = div(dn, 3.0);
  sxx = mul(sxx, t1);
  sxx = mul(sxx, t2);
  sxx = round_away(sxx);
  double d = mul(sx, sx);
  d = mul(dn, sxx) - d;
  const double c = mul(sx, sy);
  const double num = mul(dn, sxy) - c;
  return div(num, d);
}

class Fitter {
 public:
  Fitter(const std::vector<double>& y, std::uint32_t m, HoltWintersFit* out)
      : n_(y.size()),
        seasonal_(m > 1U),
        mp_(m > 2U ? static_cast<std::size_t>(m) : 2U),
        size_(mp_ + 2U),
        dm_(static_cast<double>(mp_)),
        yw_(y),
        out_(out) {}

  void run() {
    pad_short_series();
    nw_ = yw_.size();
    split(nw_, &fit_w_);
    split(n_, &fit_n_);

    std::vector<double> st(size_, 0.0);
    st[0] = yw_[0];
    if (seasonal_) {
      seasonal_init(&st);
      st[0] = st[0] - st[mp_ + 1U];
    }
    normalize(&st);
    st[1] = ols_slope(yw_.data(), nw_, st.data(), mp_);
    st[0] = st[0] - st[1];

    search(st);
    final_pass(st);
  }

 private:
  // A short seasonal series is extended phase-wise by linear extrapolation.
  void pad_short_series() {
    if (!seasonal_ || 7U * mp_ < 2U * n_) {
      return;
    }
    std::size_t extra = (7U * mp_ - 2U * n_ + 1U) / 2U;
    const std::size_t rep = mp_ > n_ ? mp_ - n_ : 0U;
    double slope = 0.0;
    if (rep > 0U) {
      for (std::size_t i = 0; i < rep; ++i) {
        yw_.push_back(yw_[n_ - 1U]);
      }
      extra = extra > rep ? extra - rep : 0U;
    } else {
      double acc = 0.0;
      for (std::size_t p = 0; p < mp_; ++p) {
        double arr[3];
        std::size_t cnt = 0;
        for (std::size_t k = p; k < n_ && cnt < 3U; k += mp_) {
          arr[cnt++] = yw_[k];
        }
        acc = ols_slope(arr, cnt, nullptr, 0U) + acc;
      }
      slope = div(acc, dm_);
      slope = div(slope, dm_);
    }
    double pos = static_cast<double>(n_);
    for (std::size_t k = 0; k < extra; ++k) {
      const std::size_t j = (k + n_) % mp_;
      double v = pos - static_cast<double>(j);
      v = mul(v, slope);
      v = v + yw_[j];
      yw_.push_back(v);
      pos = pos + 1.0;
    }
  }

  // Splits `n` points into a training part (`*fit`) and a holdout.
  void split(std::size_t n, std::size_t* fit) const {
    const std::size_t k = static_cast<std::size_t>(mul(0.8, static_cast<double>(n)));
    std::size_t f = n > 10U ? k : n;
    std::size_t hold = n - f;
    if (hold != 0U) {
      const std::size_t cap = static_cast<std::size_t>(mul(0.4, static_cast<double>(n)));
      std::size_t h = hold > 2U * mp_ ? hold : 2U * mp_;
      if (h <= 10U) {
        h = 10U;
      }
      if (h >= cap) {
        h = cap;
      }
      f = n - h;
    }
    *fit = f;
  }

  // Seasonal factors from a centred 2m-point moving-average decomposition.
  void seasonal_init(std::vector<double>* st) const {
    const std::size_t m = mp_;
    const std::size_t lim = std::min(nw_, 5U * m);
    const std::size_t nma = lim - 2U * m;
    const double d2m = static_cast<double>(2U * m);
    std::vector<double> ma(nma, 0.0);
    double run = 0.0;
    for (std::size_t i = 0; i < lim; ++i) {
      if (i >= 2U * m) {
        ma[i - 2U * m] = div(run, d2m);
        run = run - yw_[i - 2U * m];
      }
      run = run + yw_[i];
    }
    const std::size_t half = m - (m >> 1U);
    for (std::size_t p = 0; p < m; ++p) {
      double cnt = 0.0;
      double s = 0.0;
      for (std::size_t k = p; k < nma; k += m) {
        cnt = cnt + 1.0;
        s = s + yw_[k + (m >> 1U) - 1U];
        s = s - ma[k];
      }
      s = div(s, cnt);
      const std::size_t idx = p <= half ? half + 2U - p : half + 2U - p + m;
      (*st)[idx] = s;
    }
  }

  // Makes the seasonal slots sum to zero, moving the mean into the level.
  void normalize(std::vector<double>* st) const {
    double s = 0.0;
    for (std::size_t i = 0; i < mp_; ++i) {
      s = s + (*st)[2U + i];
    }
    s = div(s, dm_);
    (*st)[0] = (*st)[0] + s;
    for (std::size_t i = 0; i < mp_; ++i) {
      (*st)[2U + i] = (*st)[2U + i] - s;
    }
  }

  // Applies one error correction and returns the next seasonal slot.
  std::size_t update(std::vector<double>* st, double e, double a, double b, double g, std::size_t idx) const {
    std::vector<double>& s = *st;
    double v = mul(e, a);
    v = v + s[1];
    s[0] = v + s[0];
    v = mul(e, b);
    s[1] = v + s[1];
    v = mul(e, g);
    v = v + s[idx];
    s[idx] = v;
    return idx == 2U ? size_ - 1U : idx - 1U;
  }

  // Scaled SSE of one candidate; stops early once it exceeds `best`.
  double sse(const std::vector<double>& st0, double a, double b, double g, double best) const {
    std::vector<double> st = st0;
    const double thr = div(best, 0.1);
    double s = 0.0;
    std::size_t idx = size_ - 1U;
    for (std::size_t t = 0; t < nw_; ++t) {
      double f = st[0] + st[1];
      f = f + st[idx];
      const double e = yw_[t] - f;
      s = mul(e, e) + s;
      if (s > thr) {
        break;
      }
      if (t < fit_w_) {
        idx = update(&st, e, a, b, g, idx);
      } else {
        st[0] = st[1] + st[0];
        idx -= 1U;
        if (idx <= 1U) {
          idx = size_ - 1U;
        }
      }
    }
    return mul(s, 0.1);
  }

  void search(const std::vector<double>& st) {
    double best = mul(DBL_MAX, 0.1);
    const auto consider = [&](double a, double b, double g) {
      const double v = sse(st, a, b, g, best);
      if (v < best) {
        best = v;
        alpha_ = a;
        beta_ = b;
        gamma_ = g;
        return best == 0.0;
      }
      return false;
    };
    if (seasonal_) {
      for (const SeasonalCandidate& c : kSeasonalCandidates) {
        if (consider(c.alpha, c.beta, c.gamma)) {
          break;
        }
      }
    } else {
      for (const TrendCandidate& c : kTrendCandidates) {
        if (consider(c.alpha, c.beta, DBL_EPSILON)) {
          break;
        }
      }
    }
  }

  // Runs the chosen parameters over the series, then derives the state and
  // the error statistics.
  void final_pass(const std::vector<double>& st0) {
    std::vector<double> st = st0;
    const double a = alpha_;
    const double b = beta_;
    const double g = gamma_;
    const std::size_t half = n_ > 10U ? n_ >> 1U : 0U;
    double sum_b = 0.0;
    double sum_b2 = 0.0;
    double naive = 0.0;
    double sse_all = 0.0;
    double sae = 0.0;
    double s2 = 0.0;
    double smape = 0.0;
    std::size_t idx = size_ - 1U;
    const std::size_t fit = fit_n_;
    for (std::size_t t = 0; t < n_; ++t) {
      double f = st[0] + st[1];
      f = f + st[idx];
      const double e = yw_[t] - f;
      sse_all = mul(e, e) + sse_all;
      idx = update(&st, e, a, b, g, idx);
      if (t > half) {
        sum_b = st[1] + sum_b;
        sum_b2 = mul(st[1], st[1]) + sum_b2;
      }
      if (t != 0U && (t < fit || fit == n_)) {
        naive = std::fabs(yw_[t] - yw_[t - 1U]) + naive;
      }
      if (t >= fit || fit == n_) {
        if (e != 0.0) {
          sae = std::fabs(e) + sae;
          s2 = mul(e, e) + s2;
        }
        if (!(f == 0.0 && yw_[t] == 0.0)) {
          const double d = std::fabs(f) + std::fabs(yw_[t]);
          smape = std::fabs(div(e, d)) + smape;
        }
      }
    }
    const std::size_t cnt = n_ > 10U ? n_ >> 1U : n_;
    const double mean_b = div(sum_b, static_cast<double>(cnt));
    const double dn1 = static_cast<double>(cnt - 1U);
    double v = mul(sum_b, sum_b);
    v = div(v, dn1);
    v = sum_b2 - v;
    v = div(v, dn1);
    v = std::fabs(v);
    const double sd = std::sqrt(v);
    const double sdv = cnt > 1U ? sd : 0.0;
    const double cv = mean_b != 0.0 ? std::fabs(div(sdv, mean_b)) : 0.0;
    if (mean_b == 0.0 || !(cv < 1.0)) {
      st[0] = yw_[n_ - 1U] - st[2];
      st[1] = st0[1];
      beta_ = 0.001;
    }
    naive = div(naive, static_cast<double>(fit - 1U));
    const double cnt2 = static_cast<double>(n_ - (fit != n_ ? fit : 0U));
    const double mae = div(sae, cnt2);
    smape = mul(smape, 2.0);
    normalize(&st);

    out_->alpha = alpha_;
    out_->beta = beta_;
    out_->gamma = gamma_;
    out_->sigma2 = div(sse_all, static_cast<double>(n_));
    out_->mae = mae;
    out_->rmse = std::sqrt(div(s2, cnt2));
    out_->smape = div(smape, cnt2);
    out_->mase = naive != 0.0 ? div(mae, naive) : mae;
    out_->period = mp_;
    out_->state = std::move(st);
  }

  std::size_t n_;
  bool seasonal_;
  std::size_t mp_;
  std::size_t size_;
  double dm_;
  std::vector<double> yw_;
  HoltWintersFit* out_;
  std::size_t nw_ = 0;
  std::size_t fit_w_ = 0;
  std::size_t fit_n_ = 0;
  double alpha_ = 0.0;
  double beta_ = 0.0;
  double gamma_ = 0.0;
};

}  // namespace

double HoltWintersFit::forecast(std::int64_t h) const {
  const std::int64_t m = static_cast<std::int64_t>(period);
  double v = mul(static_cast<double>(h), state[1]);
  v = v + state[0];
  const std::int64_t k = h >= 1 ? h : h + m * ((-h) / m + 1);
  const std::size_t idx = state.size() - 1U - static_cast<std::size_t>((k - 1) % m);
  return v + state[idx];
}

double HoltWintersFit::variance(std::int64_t h) const {
  if (h == 1) {
    return sigma2;
  }
  const double a = alpha;
  const double b = beta;
  const double g = gamma;
  const double m = static_cast<double>(period);
  const double dh = static_cast<double>(h);
  const double dh1 = static_cast<double>(h - 1);
  const double k = static_cast<double>((h - 1) / static_cast<std::int64_t>(period));
  const double d2h = static_cast<double>(2 * h - 1);
  double v = mul(dh, d2h);
  v = mul(v, b);
  v = mul(v, b);
  const double trend_term = div(v, 6.0);
  const double a2 = mul(a, a);
  v = mul(dh, a);
  v = mul(v, b);
  v = v + trend_term;
  v = v + a2;
  v = mul(v, dh1);
  const double base = 1.0 + v;
  v = k + 1.0;
  v = mul(v, m);
  v = mul(v, b);
  const double c = v + g;
  v = mul(2.0, a);
  v = v + c;
  v = mul(v, g);
  v = mul(v, k);
  v = v + base;
  return mul(v, sigma2);
}

bool fit_holt_winters(const std::vector<double>& y, std::uint32_t m, HoltWintersFit* out, Value* out_err) {
  if (y.size() < 2U) {
    *out_err = Value::error(ErrorCode::NA);
    return false;
  }
  Fitter(y, m, out).run();
  return true;
}

}  // namespace ets
}  // namespace eval
}  // namespace formulon
