//
// Seasonality detection for the FORECAST.ETS family; the contract is declared
// in `eval/ets_series.h`.

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <utility>
#include <vector>

#include "eval/ets_series.h"
#include "eval/ets_series_internal.h"
#include "eval/excel_math.h"
#include "eval/stats/numeric_constants.h"
#include "utils/index_sort.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace ets {

using namespace series_detail;

namespace {

// ---------------------------------------------------------------------------
// Seasonality detection
// ---------------------------------------------------------------------------
//
// The series is detrended by least squares and standardised. A tapered
// periodogram of the residuals gives a short-period candidate; a periodogram
// of the standardised series proposes longer candidates, each scored by the
// shape of the autocorrelation around its lag. An amplitude comparison
// arbitrates when both stages report a short period.

constexpr double kPi = 3.141592653589793;
constexpr double kInt32Min = -2147483648.0;
constexpr double kInt32Max = 2147483647.0;
// Longest series considered by the detector.
constexpr std::size_t kMaxDetectLength = 175680U;
// From this length on, only the largest power-of-two tail is analysed.
constexpr std::size_t kWindowedLength = 2049U;
// Detected periods at or above this are reported as none.
constexpr std::uint32_t kMaxDetectedPeriod = 8785U;
// Fraction of the series tapered at each end before the periodogram.
constexpr double kTaperFraction = 0.1;
// Normalisation of the periodogram power.
constexpr double kPowerScale = 0.875;
// Smallest mean for which the autocorrelation input is mean-centred.
constexpr double kCentringThreshold = 3e-7;
// Offset subtracted from 1/frequency before rounding the short period.
constexpr double kPeriodBias = 3e-7;
// Minimum autocorrelation at an accepted peak.
constexpr double kMinPeakAcf = 0.45;
// Peak-to-neighbour power ratio that marks a strong periodogram peak.
constexpr double kStrongPeakRatio = 4.0;
// Power ratio above which candidates no longer compete in the selection.
constexpr double kSelectionRatio = 2.0;
// Standard deviations above the mean a long-period peak must reach.
constexpr double kPeakSigma = 2.5;
// A period must repeat at least this many times in the series.
constexpr double kMinRepeats = 2.5;

// Negation that leaves values whose top 16 bits are zero unchanged.
double flip_sign(double v) noexcept {
  return top_bits_zero(v) ? v : -v;
}

// Least-squares line over xs[start, end); returns {slope, intercept}.
std::pair<double, double> regress(const std::vector<double>& xs, const std::vector<double>& ys, std::size_t start,
                                  std::size_t end) {
  double sx = 0.0;
  double sxx = 0.0;
  double sy = 0.0;
  double sxy = 0.0;
  for (std::size_t i = start; i < end; ++i) {
    const double x = xs[i];
    const double y = ys[i];
    sx = x + sx;
    sxx = mul_flush(x, x) + sxx;
    sy = sy + y;
    sxy = mul_flush(x, y) + sxy;
  }
  const double dn = static_cast<double>(end - start);
  const double d = mul_flush(dn, sxx) - mul_flush(sx, sx);
  const double num = mul_flush(dn, sxy) - mul_flush(sx, sy);
  const double slope = div_zero(num, d);
  const double t = div_zero(mul_flush(slope, sx), dn);
  return {slope, div_zero(sy, dn) - t};
}

struct Complex {
  double re;
  double im;
};

// In-place radix-2 transform of a real input of power-of-two length; the
// output is conjugated.
std::vector<Complex> fft_pow2(const std::vector<double>& x, std::size_t count, bool inverse) {
  std::vector<double> d(2U * count + 2U, 0.0);
  for (std::size_t k = 0; k < count; ++k) {
    d[1U + 2U * k] = x[k];
  }
  const std::size_t n2 = 2U * count;
  std::size_t j = 1;
  for (std::size_t i = 1; i < n2; i += 2U) {
    if (i < j) {
      std::swap(d[j], d[i]);
      std::swap(d[j + 1U], d[i + 1U]);
    }
    std::size_t m = count;
    while (m >= 2U && j > m) {
      j -= m;
      m >>= 1U;
    }
    j += m;
  }
  std::size_t mmax = 2;
  while (true) {
    double theta = div_zero(stats::kTwoPi, static_cast<double>(mmax));
    if (inverse) {
      theta = flip_sign(theta);
    }
    const double wpi = std::sin(theta);
    double wt = std::sin(mul_flush(theta, 0.5));
    wt = mul_flush(wt, wt);
    const double wpr = mul_flush(wt, -2.0);
    const std::size_t istep = 2U * mmax;
    double wr = 1.0;
    double wi = 0.0;
    for (std::size_t m = 1; m < mmax; m += 2U) {
      for (std::size_t i = m; i <= n2; i += istep) {
        const std::size_t jj = i + mmax;
        const double t1 = mul_flush(wi, d[jj + 1U]);
        const double t0 = mul_flush(wr, d[jj]);
        const double tempr = t0 - t1;
        const double t2 = mul_flush(wi, d[jj]);
        const double t3 = mul_flush(wr, d[jj + 1U]);
        const double tempi = t3 + t2;
        d[jj] = d[i] - tempr;
        d[jj + 1U] = d[i + 1U] - tempi;
        d[i] = d[i] + tempr;
        d[i + 1U] = d[i + 1U] + tempi;
      }
      const double wtemp = wr;
      double a = mul_flush(wpi, wi);
      double b = mul_flush(wpr, wr);
      b = b - a;
      wr = b + wr;
      a = mul_flush(wpi, wtemp);
      b = mul_flush(wpr, wi);
      b = b + a;
      wi = b + wi;
    }
    mmax = istep;
    if (!(istep < n2)) {
      break;
    }
  }
  std::vector<Complex> out(count);
  for (std::size_t k = 0; k < count; ++k) {
    out[k] = Complex{d[1U + 2U * k], flip_sign(d[2U + 2U * k])};
  }
  return out;
}

// Direct transform for lengths that are not a power of two, using a twiddle
// table built by accumulating the angle step.
std::vector<Complex> dft_direct(const std::vector<double>& x, std::size_t count) {
  const double step = div_zero(stats::kTwoPi, static_cast<double>(count));
  double ang = 0.0;
  std::vector<Complex> tw(count);
  for (std::size_t k = 0; k < count; ++k) {
    tw[k] = Complex{std::cos(ang), flip_sign(std::sin(ang))};
    ang = ang + step;
  }
  std::vector<Complex> out(count);
  for (std::size_t j = 0; j < count; ++j) {
    double re = 0.0;
    double im = 0.0;
    std::size_t idx = 0;
    for (std::size_t i = 0; i < count; ++i) {
      const std::size_t k = idx % count;
      re = re + mul_flush(tw[k].re, x[i]);
      im = im + mul_flush(tw[k].im, x[i]);
      idx += j;
    }
    out[j] = Complex{re, im};
  }
  return out;
}

std::vector<Complex> transform(const std::vector<double>& x, std::size_t count, bool inverse) {
  if ((count & (count - 1U)) == 0U) {
    return fft_pow2(x, count, inverse);
  }
  return dft_direct(x, count);
}

// Transform length used by the autocorrelation: a power of two above 2n.
std::size_t acf_length(std::size_t n) {
  if ((n & (n - 1U)) == 0U) {
    return 2U * n;
  }
  std::size_t p = 1;
  std::size_t m = n;
  while (m > 3U) {
    p <<= 1U;
    m >>= 1U;
  }
  return p << 3U;
}

struct Periodogram {
  std::vector<double> freq;
  std::vector<double> power;
  double mean = 0.0;
  double sd = 0.0;
};

Periodogram make_periodogram(const std::vector<double>& y) {
  const std::size_t n = y.size();
  const std::int64_t half = static_cast<std::int64_t>((n + 1U) >> 1U);
  const auto centre = [&](std::size_t i) { return static_cast<double>(1 - half + static_cast<std::int64_t>(i)); };
  double s = 0.0;
  double sxy = 0.0;
  for (std::size_t i = 0; i < n; ++i) {
    s = s + y[i];
    sxy = mul_flush(y[i], centre(i)) + sxy;
  }
  const double mean = div_zero(s, static_cast<double>(n));
  double den = (2.0 * static_cast<double>(n)) * static_cast<double>(static_cast<std::uint32_t>(n * n - 1U));
  den = div_zero(den, 24.0);
  const double slope = div_zero(sxy, den);
  std::vector<double> d(n);
  for (std::size_t i = 0; i < n; ++i) {
    double v = mul_flush(centre(i), slope);
    v = v + mean;
    v = flip_sign(v);
    d[i] = v + y[i];
  }
  const std::uint32_t taper = to_uint32(mul_flush(static_cast<double>(n), kTaperFraction));
  const double two_m = mul_flush(static_cast<double>(taper), 2.0);
  for (std::uint32_t i = 0; i < taper; ++i) {
    double w = mul_flush(static_cast<double>(2U * i + 1U), kPi);
    w = div_zero(w, two_m);
    w = std::cos(w);
    w = flip_sign(w);
    w = w + 1.0;
    w = mul_flush(w, 0.5);
    d[i] = mul_flush(w, d[i]);
    d[n - 1U - i] = mul_flush(w, d[n - 1U - i]);
  }
  const std::vector<Complex> out = transform(d, n, false);
  const double scale = mul_flush(static_cast<double>(n), kPowerScale);
  const std::size_t bins = n >> 1U;
  Periodogram pg;
  double sum = 0.0;
  double sum2 = 0.0;
  for (std::size_t i = 0; i < bins; ++i) {
    const double p = mul_flush(out[i + 1U].re, out[i + 1U].re);
    double q = mul_flush(out[i + 1U].im, out[i + 1U].im);
    q = q + p;
    q = div_zero(q, scale);
    sum = q + sum;
    sum2 = mul_flush(q, q) + sum2;
    pg.freq.push_back(div_zero(static_cast<double>(i + 1U), static_cast<double>(n)));
    pg.power.push_back(q);
  }
  if (n >= 2U) {
    pg.mean = div_zero(sum, static_cast<double>(bins));
    if (n >= 4U) {
      const double v = div_zero(sum2, static_cast<double>(bins - 1U)) - mul_flush(pg.mean, pg.mean);
      pg.sd = v >= 0.0 ? std::sqrt(v) : 0.0;
    }
  }
  return pg;
}

// Autocorrelation of the standardised series via zero-padded transforms,
// normalised by the overlap count at each lag.
std::vector<double> make_acf(const std::vector<double>& std_y) {
  const std::size_t n = std_y.size();
  const std::size_t len = acf_length(n);
  double s = 0.0;
  for (double v : std_y) {
    s = v + s;
  }
  const double mean = div_zero(s, static_cast<double>(n));
  std::vector<double> buf(len, 0.0);
  for (std::size_t i = 0; i < n; ++i) {
    buf[i] = bit_compare(mean, kCentringThreshold) == 1 ? std_y[i] - mean : std_y[i];
  }
  const auto correlate = [&](const std::vector<double>& b) {
    const std::vector<Complex> spec = transform(b, len, false);
    std::vector<double> pw(len);
    for (std::size_t k = 0; k < len; ++k) {
      pw[k] = mul_flush(spec[k].re, spec[k].re) + mul_flush(spec[k].im, spec[k].im);
    }
    const std::vector<Complex> inv = transform(pw, len, true);
    std::vector<double> r(n);
    for (std::size_t k = 0; k < n; ++k) {
      r[k] = inv[k].re;
    }
    return r;
  };
  const std::vector<double> raw = correlate(buf);
  std::vector<double> ones(len, 0.0);
  std::fill(ones.begin(), ones.begin() + static_cast<std::ptrdiff_t>(n), 1.0);
  const std::vector<double> cnt = correlate(ones);
  const double dn = static_cast<double>(n);
  double base = div_zero(raw[0], cnt[0]);
  base = div_zero(base, dn);
  std::vector<double> acf(n);
  acf[0] = 1.0;
  for (std::size_t k = 1; k < n; ++k) {
    double v = div_zero(raw[k], dn);
    v = div_zero(v, cnt[k]);
    acf[k] = div_zero(v, base);
  }
  return acf;
}

// Descending order by `power`; up to eight items use a selection sort, whose
// tie order differs from a stable sort.
template <typename T>
void sort_by_power_desc(std::vector<T>* items) {
  if (items->size() <= 8U) {
    for (std::size_t hi = items->size(); hi-- > 1U;) {
      std::size_t mx = 0;
      for (std::size_t p = 1; p <= hi; ++p) {
        if (bit_compare((*items)[mx].power, (*items)[p].power) == 1) {
          mx = p;
        }
      }
      if (mx != hi) {
        std::swap((*items)[mx], (*items)[hi]);
      }
    }
    return;
  }
  sort_by_index(*items, [](const T& a, const T& b) { return a.power > b.power; });
}

struct Candidate {
  std::int64_t period;
  double power;
};

struct LagRange {
  std::int64_t period;
  std::int64_t lo;
  std::int64_t hi;
};

struct Peak {
  std::int64_t period;
  double power;
  double acf;
};

class Detector {
 public:
  explicit Detector(const std::vector<double>& y) : y_(y), n_(y.size()) {
    std::vector<double> xs(n_);
    for (std::size_t i = 0; i < n_; ++i) {
      xs[i] = static_cast<double>(i + 1U);
    }
    const auto line = regress(xs, y_, 0, n_);
    std::vector<double> r(n_);
    double s = 0.0;
    double s2 = 0.0;
    for (std::size_t i = 0; i < n_; ++i) {
      double v = mul_flush(static_cast<double>(i + 1U), line.first);
      v = v + line.second;
      v = flip_sign(v);
      v = v + y_[i];
      r[i] = v;
      s = v + s;
      s2 = mul_flush(v, v) + s2;
    }
    const double mean = div_zero(s, static_cast<double>(n_));
    const double m2 = mul_flush(mean, mean);
    const double var = div_zero(s2, static_cast<double>(n_ - 1U)) - m2;
    const double sd = var >= 0.0 ? std::sqrt(var) : 0.0;
    std::vector<double> std_y(r);
    if (sign_of(sd) != 0) {
      for (std::size_t i = 0; i < n_; ++i) {
        std_y[i] = div_zero(r[i] - mean, sd);
      }
    }
    resid_ = make_periodogram(r);
    std_ = make_periodogram(std_y);
    acf_ = make_acf(std_y);
  }

  std::uint32_t run() {
    short_stage();
    long_stage();
    return arbitrate();
  }

 private:
  double acf_at(std::int64_t lag) const {
    return lag >= 0 && static_cast<std::size_t>(lag) < acf_.size() ? acf_[static_cast<std::size_t>(lag)] : 0.0;
  }

  // Short-period candidate from the residual periodogram.
  void short_stage() {
    if (n_ < 4U) {
      return;
    }
    const std::size_t cap = 2U * n_ < 25U ? (2U * n_) / 5U : 5U;
    const std::vector<double>& pw = resid_.power;
    const std::vector<double>& fr = resid_.freq;
    double best = pw[0];
    double second = pw[1];
    std::uint32_t per = to_uint32(round_away(div_zero(1.0, fr[0])));
    if (bit_compare(pw[1], best) == 1) {
      second = best;
      best = pw[1];
      per = to_uint32(round_away(div_zero(1.0, fr[1]) - kPeriodBias));
    }
    for (std::size_t i = 2; i < (n_ >> 1U); ++i) {
      const double p = pw[i];
      if (bit_compare(p, best) == 1) {
        second = best;
        best = p;
        per = to_uint32(round_away(div_zero(1.0, fr[i]) - kPeriodBias));
      } else if (bit_compare(p, second) == 1) {
        second = p;
      }
    }
    short_power_ = best;
    if (per >= 2U && per <= cap) {
      const double ratio = sign_of(second) != 0 ? div_zero(best, second) : 0.0;
      const double thr = 2.0 - div_zero(3.0, static_cast<double>(n_));
      if (bit_compare(ratio, thr) == 1 && bit_compare(best, resid_.mean) == 1) {
        short_period_ = per;
      }
    } else {
      const double thr = mul_flush(kPeakSigma, resid_.sd) + resid_.mean;
      if (bit_compare(best, thr) == 1 && static_cast<double>(per) * kMinRepeats <= static_cast<double>(n_)) {
        short_period_ = per;
      }
    }
  }

  std::vector<Candidate> candidates() const {
    const std::size_t bins = n_ >> 1U;
    std::vector<Candidate> cands;
    std::vector<Candidate> strong;
    const std::vector<double>& pw = std_.power;
    for (std::size_t i = 0; i < bins; ++i) {
      const double p = pw[i];
      if (bit_compare(p, std_.mean) != 1) {
        continue;
      }
      const std::int64_t per = static_cast<std::int64_t>(to_uint32(round_away(div_zero(1.0, std_.freq[i]))));
      if (per < 4 || static_cast<std::int64_t>(n_) <= 2 * per) {
        continue;
      }
      cands.push_back(Candidate{per, p});
      if (i == 0U || i == bins - 1U) {
        continue;
      }
      const double prev = pw[i - 1U];
      const double next = pw[i + 1U];
      if (sign_of(prev) != 0 && sign_of(next) != 0) {
        if (bit_compare(div_zero(p, prev), kStrongPeakRatio) == 1 &&
            bit_compare(div_zero(p, next), kStrongPeakRatio) == 1) {
          strong.push_back(Candidate{per, p});
        }
      } else if (sign_of(prev) == 0 && sign_of(next) == 0) {
        strong.push_back(Candidate{per, p});
      }
    }
    sort_by_power_desc(&cands);
    while (cands.size() > 3U) {
      cands.pop_back();
    }
    for (const Candidate& e : strong) {
      const bool known =
          std::any_of(cands.begin(), cands.end(), [&](const Candidate& c) { return c.period == e.period; });
      if (!known) {
        cands.push_back(e);
      }
    }
    return cands;
  }

  LagRange lag_range(const Candidate& c) const {
    const double dn = static_cast<double>(n_);
    const double d = div_zero(dn, static_cast<double>(c.period));
    const std::int64_t lo = static_cast<std::int64_t>(to_uint32(round_away(div_zero(dn, d + 1.0) - 1.0)));
    const std::int64_t hi = static_cast<std::int64_t>(to_uint32(round_away(div_zero(dn, d - 1.0) + 1.0)));
    const std::int64_t lo2 = lo < 4 ? 3 : lo;
    std::int64_t hi2 = (hi < 3 && lo < 4) ? 5 : hi;
    if (hi2 > static_cast<std::int64_t>(n_)) {
      hi2 = lo2;
    }
    return LagRange{c.period, lo2, hi2};
  }

  // A short-lag peak is accepted when the autocorrelation rises into it and
  // falls away after it.
  bool narrow_peak(std::int64_t p, std::int64_t x2, Peak* out) const {
    const std::int64_t w = p == 5 ? 4 : 2;
    const std::size_t cnt = static_cast<std::size_t>(2 * w + 1);
    std::vector<double> xs(cnt);
    std::vector<double> ys(cnt);
    for (std::size_t i = 0; i < cnt; ++i) {
      const std::int64_t x = p - w + static_cast<std::int64_t>(i);
      xs[i] = static_cast<double>(x);
      ys[i] = acf_at(x);
    }
    const double s1 = regress(xs, ys, 0, static_cast<std::size_t>(w) + 1U).first;
    const double s2 = regress(xs, ys, static_cast<std::size_t>(w) + 1U, cnt).first;
    if (sign_of(s1) == 1 && sign_of(s2) == -1) {
      *out = Peak{x2, 0.0, acf_at(x2)};
      return true;
    }
    return false;
  }

  // Two-piece linear fit of the flank-detrended autocorrelation over the lag
  // range; accepted when it rises then falls.
  bool broad_peak(const LagRange& r, std::int64_t x2, Peak* out) const {
    const std::int64_t span = r.hi - r.lo;
    const std::size_t cnt = static_cast<std::size_t>(span + 1);
    std::vector<double> fx;
    std::vector<double> fy;
    for (std::int64_t k = -3; k <= -1; ++k) {
      const std::int64_t lag = r.lo + k;
      if (lag + 1 >= 3) {
        fx.push_back(static_cast<double>(lag + 1));
        fy.push_back(acf_at(lag));
      }
    }
    for (std::int64_t k = -1; k <= 1; ++k) {
      const std::int64_t lag = r.hi + k;
      if (lag + 1 <= static_cast<std::int64_t>(n_)) {
        fx.push_back(static_cast<double>(lag + 1));
        fy.push_back(acf_at(lag));
      }
    }
    const auto line = regress(fx, fy, 0, fx.size());
    std::vector<double> xs(cnt);
    std::vector<double> ys(cnt);
    for (std::size_t i = 0; i < cnt; ++i) {
      const std::int64_t lag = r.lo + static_cast<std::int64_t>(i);
      double v = mul_flush(static_cast<double>(lag), line.first);
      v = v + line.second;
      v = flip_sign(v);
      v = v + acf_at(lag);
      xs[i] = static_cast<double>(lag);
      ys[i] = v;
    }
    const std::int64_t step =
        static_cast<std::int64_t>(to_uint32(std::ceil(div_zero(static_cast<double>(span - 3), 100.0))));
    const std::int64_t last = span - 1;
    if (last < 3) {
      return false;
    }
    double best = kInt32Max;
    bool found = false;
    for (std::int64_t i = 2; i < last; i += step) {
      const std::size_t si = static_cast<std::size_t>(i);
      const auto left = regress(xs, ys, 0, si + 1U);
      const auto right = regress(xs, ys, si, cnt);
      double m1 = 0.0;
      for (std::size_t j = 0; j <= si; ++j) {
        double v = mul_flush(xs[j], left.first);
        v = v + left.second;
        v = flip_sign(v);
        v = v + ys[j];
        v = mul_flush(v, v);
        m1 = v + m1;
      }
      m1 = div_zero(m1, static_cast<double>(i + 1));
      double m2 = 0.0;
      for (std::size_t j = si; j < cnt; ++j) {
        double v = mul_flush(xs[j], right.first);
        v = v + right.second;
        v = flip_sign(v);
        v = v + ys[j];
        v = mul_flush(v, v);
        m2 = v + m2;
      }
      m2 = div_zero(m2, static_cast<double>(cnt - si));
      const double total = m1 + m2;
      if (bit_compare(total, best) == -1) {
        found = sign_of(left.first) == 1 && sign_of(right.first) == -1;
        if (found) {
          *out = Peak{x2, 0.0, acf_at(x2)};
        }
        best = total;
      }
    }
    return found;
  }

  // Scores one lag range; `doubled` marks the retry at twice the period.
  bool score(const LagRange& r, bool doubled, Peak* out) const {
    const std::int64_t cnt = r.hi - r.lo + 1;
    double best = kInt32Min;
    std::int64_t bl = 0;
    for (std::int64_t lag = r.lo; lag <= r.hi; ++lag) {
      const double v = acf_at(lag);
      if (bit_compare(v, best) == 1) {
        bl = lag;
        best = v;
      }
    }
    std::int64_t x2 = bl;
    if (doubled) {
      x2 = static_cast<std::int64_t>(to_uint32(std::floor(div_zero(static_cast<double>(bl), 2.0))));
    }
    if (static_cast<double>(x2) * kMinRepeats > static_cast<double>(n_)) {
      return false;
    }
    if (bl == r.lo || bl == r.hi) {
      return false;
    }
    if ((bl >= 3 && bl <= 5) || (doubled && bl >= 6 && bl <= 11)) {
      return narrow_peak(bl, x2, out);
    }
    if (cnt < 5) {
      return false;
    }
    return broad_peak(r, x2, out);
  }

  // Long-period candidates from the standardised periodogram.
  void long_stage() {
    const std::vector<Candidate> cands = candidates();
    if (cands.empty()) {
      return;
    }
    std::vector<Peak> scored;
    const std::int64_t n = static_cast<std::int64_t>(n_);
    for (const Candidate& c : cands) {
      LagRange r = lag_range(c);
      const std::int64_t half =
          static_cast<std::int64_t>(to_uint32(std::floor(div_zero(static_cast<double>(r.period), 2.0))));
      if (half < r.hi) {
        r.lo = std::max(r.lo, half);
      }
      Peak peak{};
      if (!score(r, false, &peak)) {
        if (2 * r.hi >= n) {
          continue;
        }
        const LagRange twice{2 * r.period, 2 * r.lo, 2 * r.hi};
        if (!score(twice, true, &peak)) {
          continue;
        }
      }
      if (peak.acf < kMinPeakAcf) {
        continue;
      }
      const auto hit =
          std::find_if(scored.begin(), scored.end(), [&](const Peak& e) { return e.period == peak.period; });
      if (hit != scored.end()) {
        if (peak.acf > hit->acf) {
          hit->acf = peak.acf;
        }
      } else {
        scored.push_back(peak);
      }
    }
    if (scored.empty()) {
      return;
    }
    if (scored.size() == 1U) {
      long_period_ = scored[0].period;
      return;
    }
    std::vector<Peak> pool = scored;
    if (short_period_ >= 2 && short_period_ <= 5) {
      const Peak e{short_period_, short_power_, acf_at(short_period_)};
      const auto hit = std::find_if(pool.begin(), pool.end(), [&](const Peak& t) { return t.period == short_period_; });
      if (hit != pool.end()) {
        if (bit_compare(hit->power, e.power) != 1) {
          hit->power = e.power;
        }
        if (bit_compare(hit->acf, e.acf) != 1) {
          hit->acf = e.acf;
        }
      } else {
        pool.push_back(e);
      }
    }
    sort_by_power_desc(&pool);
    const std::size_t cnt = pool.size();
    std::size_t reach = 1;
    if (cnt >= 2U) {
      std::size_t i = 0;
      while (true) {
        const double ratio = div_zero(pool[i].power, pool[i + 1U].power);
        if (bit_compare(ratio, kSelectionRatio) == 1) {
          reach = i;
          break;
        }
        ++i;
        if (i + 1U >= cnt) {
          reach = cnt;
          break;
        }
      }
    }
    const std::size_t top = reach < cnt ? reach : cnt - 1U;
    double cur = kInt32Min;
    for (std::size_t k = top + 1U; k-- > 0U;) {
      if (bit_compare(pool[k].acf, cur) == 1) {
        long_period_ = pool[k].period;
        cur = pool[k].acf;
      }
    }
  }

  // Compares the amplitude of the two stage periods when both are short.
  std::uint32_t arbitrate() const {
    const std::int64_t p38 = long_period_;
    const std::int64_t p3c = short_period_;
    if (p38 == 1) {
      return static_cast<std::uint32_t>(p3c);
    }
    if (p3c == 1 || p38 > 5) {
      return static_cast<std::uint32_t>(p38);
    }
    if (static_cast<std::int64_t>(n_) < 2 * p3c) {
      return static_cast<std::uint32_t>(p38);
    }
    double amps[2];
    const std::int64_t periods[2] = {p38, p3c};
    for (int s = 0; s < 2; ++s) {
      const double period = static_cast<double>(periods[s]);
      const std::uint32_t k = to_uint32(std::floor(div_zero(static_cast<double>(n_), period)));
      const std::size_t length = static_cast<std::size_t>(k) * static_cast<std::size_t>(periods[s]);
      double c = 0.0;
      double sn = 0.0;
      std::size_t i = 0;
      for (std::size_t t = n_ - length; t < n_; ++t) {
        double v = mul_flush(static_cast<double>(i), kPi);
        v = mul_flush(v, 2.0);
        v = flip_sign(v);
        v = div_zero(v, period);
        c = mul_flush(std::cos(v), y_[t]) + c;
        sn = mul_flush(std::sin(v), y_[t]) + sn;
        ++i;
      }
      double a = mul_flush(c, c);
      double b = mul_flush(sn, sn);
      b = b + a;
      b = std::sqrt(b);
      amps[s] = div_zero(b, static_cast<double>(length));
    }
    double cur = kInt32Min;
    std::int64_t win = p38;
    for (int s = 0; s < 2; ++s) {
      if (bit_compare(amps[s], cur) == 1) {
        win = periods[s];
        cur = amps[s];
      }
    }
    return static_cast<std::uint32_t>(win);
  }

  const std::vector<double>& y_;
  std::size_t n_;
  Periodogram resid_;
  Periodogram std_;
  std::vector<double> acf_;
  std::int64_t short_period_ = 1;
  std::int64_t long_period_ = 1;
  double short_power_ = 0.0;
};

}  // namespace

// Auto-detected seasonality length: 1 means none, otherwise at least 2.
std::uint32_t detect_seasonality(const std::vector<double>& y) {
  std::size_t n = std::min(y.size(), kMaxDetectLength);
  if (n >= kWindowedLength) {
    std::size_t p = 1;
    while ((p << 1U) <= n) {
      p <<= 1U;
    }
    n = p;
  }
  if (n < 2U) {
    return 1U;
  }
  const std::vector<double> window(y.end() - static_cast<std::ptrdiff_t>(n), y.end());
  const std::uint32_t period = Detector(window).run();
  return (period >= kMaxDetectedPeriod || period <= 1U) ? 1U : period;
}

}  // namespace ets
}  // namespace eval
}  // namespace formulon
