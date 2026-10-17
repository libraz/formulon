//
// Series preparation for the FORECAST.ETS family; the contract is declared in
// `eval/ets_series.h`.

#include "eval/ets_series.h"

#include <algorithm>
#include <cfloat>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <utility>
#include <vector>

#include "utils/subnormal_flush.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace ets {
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
constexpr double kTwoPi = 6.283185307179586;
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

double div_zero(double a, double b) noexcept {
  if (std::fabs(b) < DBL_MIN) {
    return 0.0;
  }
  const double r = a / b;
  return (std::fabs(r) > DBL_MAX || std::fabs(r) < DBL_MIN) ? 0.0 : r;
}

double mul_flush(double a, double b) noexcept {
  return flush_subnormal(a * b);
}

// Negation that leaves values whose top 16 bits are zero unchanged.
double flip_sign(double v) noexcept {
  std::uint64_t bits = 0;
  std::memcpy(&bits, &v, sizeof bits);
  return (bits >> 48U) == 0U ? v : -v;
}

// Three-way compare on the sign-magnitude integer image: -0.0 sorts below
// +0.0 and NaN above every number.
int bit_compare(double a, double b) noexcept {
  std::int64_t ia = 0;
  std::int64_t ib = 0;
  std::memcpy(&ia, &a, sizeof ia);
  std::memcpy(&ib, &b, sizeof ib);
  if (ia == ib) {
    return 0;
  }
  int r = ia < ib ? -1 : 1;
  if ((ia & ib) < 0) {
    r = -r;
  }
  return r;
}

int sign_of(double v) noexcept {
  return bit_compare(v, 0.0);
}

// Magnitude truncated to an unsigned 32-bit value, saturating.
std::uint32_t to_uint32(double v) noexcept {
  v = std::fabs(v);
  if (v != v) {
    return 0U;
  }
  if (v >= 4294967295.0) {
    return 0xffffffffU;
  }
  return static_cast<std::uint32_t>(v);
}

double round_away(double v) noexcept {
  return v >= 0.0 ? std::floor(v + 0.5) : -std::floor(-v + 0.5);
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
    double theta = div_zero(kTwoPi, static_cast<double>(mmax));
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
  const double step = div_zero(kTwoPi, static_cast<double>(count));
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
  std::stable_sort(items->begin(), items->end(), [](const T& a, const T& b) { return a.power > b.power; });
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

namespace {

// ---------------------------------------------------------------------------
// Timeline normalisation
// ---------------------------------------------------------------------------
//
// Cells are paired, sorted when a negative step occurs, and the step is read
// from a histogram of the consecutive differences. Series with gaps or
// duplicate timestamps are rebuilt onto a regular grid node by node, walking
// back from the last timestamp (by months on a calendar timeline).

constexpr double kStepTolerance = 1e-5;
// Percentage of the steps that must fall in the chosen histogram window.
constexpr std::int64_t kMinWindowPercent = 50;
constexpr std::int64_t kMaxHistogramBins = 3000;
constexpr std::int64_t kMaxRefinePasses = 11;
constexpr double kSameStepTolerance = 3e-7;
constexpr std::uint32_t kMaxDateSerial = 2958465U;
// Longest run of regular steps tracked when anchoring the grid.
constexpr std::size_t kMaxAnchorRun = 20U;
// Largest share (percent) of grid nodes that may be missing.
constexpr std::int64_t kMaxMissingPercent = 30;
constexpr std::int64_t kMaxGridNodes = 1000000;

bool top_bits_zero(double v) noexcept {
  std::uint64_t bits = 0;
  std::memcpy(&bits, &v, sizeof bits);
  return (bits >> 48U) == 0U;
}

bool is_inf_or_nan(double v) noexcept {
  return std::isnan(v) || std::isinf(v);
}

bool sign_bit_set(double v) noexcept {
  std::int64_t bits = 0;
  std::memcpy(&bits, &v, sizeof bits);
  return bits < 0;
}

template <typename V>
auto& elem(V& v, std::int64_t i) {
  return v[static_cast<std::size_t>(i)];
}

// -- calendar -----------------------------------------------------------

std::int64_t floor_div(std::int64_t a, std::int64_t b) noexcept {
  std::int64_t q = a / b;
  if ((a % b != 0) && ((a < 0) != (b < 0))) {
    --q;
  }
  return q;
}

// Days since 1970-01-01 of a proleptic Gregorian date.
std::int64_t days_from_civil(std::int64_t y, std::int64_t m, std::int64_t d) noexcept {
  y -= m <= 2 ? 1 : 0;
  const std::int64_t era = floor_div(y, 400);
  const std::int64_t yoe = y - era * 400;
  const std::int64_t doy = (153 * (m + (m > 2 ? -3 : 9)) + 2) / 5 + d - 1;
  const std::int64_t doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
  return era * 146097 + doe - 719468;
}

void civil_from_days(std::int64_t z, std::int64_t* y, std::int64_t* m, std::int64_t* d) noexcept {
  z += 719468;
  const std::int64_t era = floor_div(z, 146097);
  const std::int64_t doe = z - era * 146097;
  const std::int64_t yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
  const std::int64_t doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
  const std::int64_t mp = (5 * doy + 2) / 153;
  *d = doy - (153 * mp + 2) / 5 + 1;
  *m = mp + (mp < 10 ? 3 : -9);
  *y = yoe + era * 400 + (*m <= 2 ? 1 : 0);
}

// Serial 0 is 1899-12-30.
std::int64_t spreadsheet_epoch_days() noexcept {
  return days_from_civil(1899, 12, 30);
}

bool is_date_serial(double x) noexcept {
  return sign_of(x) != -1 && to_uint32(x) <= kMaxDateSerial;
}

struct Ymd {
  std::int64_t year = 0;
  std::int64_t month = 0;
  std::int64_t day = 0;
};

Ymd ymd_of(double serial) {
  Ymd r;
  civil_from_days(static_cast<std::int64_t>(std::floor(serial)) + spreadsheet_epoch_days(), &r.year, &r.month, &r.day);
  return r;
}

// A day past the month end rolls into the next month.
double serial_of(const Ymd& v) {
  return static_cast<double>(days_from_civil(v.year, v.month, 1) - spreadsheet_epoch_days() + (v.day - 1));
}

// Moves the (year, month) pair in place, keeping the day.
double add_months(Ymd* v, std::int64_t delta) {
  const std::int64_t total = v->year * 12 + (v->month + delta) - 1;
  v->year = floor_div(total, 12);
  v->month = total - v->year * 12 + 1;
  return serial_of(*v);
}

// Nearest month-step node from `a` toward `b` and the signed step count.
bool month_walk(std::int64_t months, double a, double b, double* node, std::int64_t* count, ErrorCode* err) {
  std::int64_t delta = 0;
  std::int64_t sign = 0;
  if (bit_compare(a, b) == -1) {
    delta = months;
    sign = 1;
  } else {
    delta = -months;
    sign = -1;
  }
  if (months <= 0) {
    *err = ErrorCode::Value;
    return false;
  }
  std::int64_t cnt = 0;
  double cur = a;
  Ymd ymd = ymd_of(a);
  double d = std::fabs(a - b);
  while (true) {
    const double next = add_months(&ymd, delta);
    const double d2 = std::fabs(next - b);
    if (bit_compare(d2, d) >= 0) {
      break;
    }
    cur = next;
    cnt += sign;
    d = d2;
  }
  *node = cur;
  *count = cnt;
  return true;
}

// -- ordering -----------------------------------------------------------

// Sorts `keys` ascending over [lo, hi], permuting `vals` alongside. Runs of
// at most eight elements use a selection pass; larger ones are partitioned
// around a median-of-three pivot with an explicit stack. The order of equal
// keys is part of the contract: it fixes the summation order of grouped
// values.
void sort_pairs(std::vector<double>* keys, std::vector<double>* vals, std::int64_t lo, std::int64_t hi) {
  std::vector<double>& k = *keys;
  const auto cmp = [&](std::int64_t i, std::int64_t j) { return bit_compare(elem(k, i), elem(k, j)); };
  const auto swap_at = [&](std::int64_t i, std::int64_t j) {
    std::swap(elem(k, i), elem(k, j));
    if (vals != nullptr) {
      std::swap(elem(*vals, i), elem(*vals, j));
    }
  };
  std::vector<std::pair<std::int64_t, std::int64_t>> stack;
  while (true) {
    const std::int64_t size = hi - lo + 1;
    if (size <= 8) {
      while (hi > lo) {
        std::int64_t mx = lo;
        for (std::int64_t p = lo + 1; p <= hi; ++p) {
          if (cmp(p, mx) > 0) {
            mx = p;
          }
        }
        if (mx != hi) {
          swap_at(mx, hi);
        }
        --hi;
      }
    } else {
      std::int64_t mid = lo + size / 2;
      if (cmp(lo, mid) > 0) {
        swap_at(lo, mid);
      }
      if (cmp(lo, hi) > 0) {
        swap_at(lo, hi);
      }
      if (cmp(mid, hi) > 0) {
        swap_at(mid, hi);
      }
      std::int64_t loguy = lo;
      std::int64_t higuy = hi;
      while (true) {
        if (mid > loguy) {
          ++loguy;
          while (loguy < mid && cmp(loguy, mid) <= 0) {
            ++loguy;
          }
        }
        if (mid <= loguy) {
          ++loguy;
          while (loguy <= hi && cmp(loguy, mid) <= 0) {
            ++loguy;
          }
        }
        --higuy;
        while (higuy > mid && cmp(higuy, mid) > 0) {
          --higuy;
        }
        if (higuy < loguy) {
          break;
        }
        swap_at(loguy, higuy);
        if (higuy == mid) {
          mid = loguy;
        }
      }
      ++higuy;
      if (mid < higuy) {
        --higuy;
        while (higuy > mid && cmp(higuy, mid) == 0) {
          --higuy;
        }
      }
      if (mid >= higuy) {
        --higuy;
        while (higuy > lo && cmp(higuy, mid) == 0) {
          --higuy;
        }
      }
      if (higuy - lo >= hi - loguy) {
        if (lo < higuy) {
          stack.emplace_back(lo, higuy);
        }
        if (loguy < hi) {
          lo = loguy;
          continue;
        }
      } else {
        if (loguy < hi) {
          stack.emplace_back(loguy, hi);
        }
        if (lo < higuy) {
          hi = higuy;
          continue;
        }
      }
    }
    if (!stack.empty()) {
      lo = stack.back().first;
      hi = stack.back().second;
      stack.pop_back();
      continue;
    }
    break;
  }
}

// -- reading ------------------------------------------------------------

// Statistics of the consecutive timeline differences.
struct StepStats {
  double min = DBL_MAX;
  double max = 0.0;
  // Non-positive steps (duplicates, and negative steps when `needs_sort`).
  std::int64_t nonpositive = 0;
  bool needs_sort = false;

  void add(double d) {
    if (bit_compare(d, min) == -1) {
      if (top_bits_zero(d)) {
        ++nonpositive;
      } else if (is_inf_or_nan(d)) {
        // Not counted.
      } else if (sign_bit_set(d)) {
        ++nonpositive;
        needs_sort = true;
      } else {
        min = d;
      }
    }
    if (bit_compare(d, max) == 1) {
      max = d;
    }
  }
};

// Pairs the cells in order. A pair with a blank, text or boolean timeline or
// value cell is dropped, except that an error cell propagates, text values
// count as 1 under COUNTA and blank values read as 0 without interpolation.
// Under COUNT/COUNTA every kept value is 1.
bool read_pairs(const std::vector<SeriesCell>& values, const std::vector<SeriesCell>& times, int completion,
                AggregationMode agg, std::vector<double>* t, std::vector<double>* y, StepStats* stats, ErrorCode* err) {
  if (values.size() != times.size()) {
    *err = ErrorCode::NA;
    return false;
  }
  const bool count_mode = agg == AggregationMode::kCount || agg == AggregationMode::kCountA;
  for (std::size_t i = 0; i < values.size(); ++i) {
    const SeriesCell& tc = times[i];
    SeriesCell yc = values[i];
    if (tc.kind == SeriesCell::Kind::kError) {
      *err = tc.error;
      return false;
    }
    if (yc.kind == SeriesCell::Kind::kError) {
      *err = yc.error;
      return false;
    }
    if (tc.kind != SeriesCell::Kind::kNumber) {
      continue;
    }
    if (yc.kind == SeriesCell::Kind::kBlank && completion == 0) {
      yc.kind = SeriesCell::Kind::kNumber;
      yc.number = 0.0;
    }
    double yv = 0.0;
    if (yc.kind == SeriesCell::Kind::kText) {
      if (agg != AggregationMode::kCountA) {
        continue;
      }
      yv = 1.0;
    } else if (yc.kind == SeriesCell::Kind::kNumber) {
      yv = count_mode ? 1.0 : yc.number;
    } else {
      continue;
    }
    if (!t->empty()) {
      stats->add(tc.number - t->back());
    }
    t->push_back(tc.number);
    y->push_back(yv);
  }
  return true;
}

// -- step detection -----------------------------------------------------

// Working state of the step histogram.
struct Histogram {
  std::vector<std::int64_t> counts;
  std::vector<double> reps;
  std::vector<std::int64_t> same;
  std::int64_t nbins = 0;
  std::int64_t zeros = 0;
};

// Do [a, b] and [lo, hi] intersect (inclusive)?
bool ranges_overlap(double a, double b, double lo, double hi) {
  const auto pos = [&](double v) {
    if (bit_compare(v, a) == -1) {
      return -1;
    }
    return bit_compare(v, b) == 1 ? 1 : 0;
  };
  const int pa = pos(lo);
  const int pb = pos(hi);
  if (pa != pb) {
    return true;
  }
  return pa == 0;
}

// Number of steps whose bins intersect [lo, hi].
std::int64_t count_in(double low, double bw, const std::vector<std::int64_t>& counts, std::int64_t nbins, double lo,
                      double hi) {
  const auto index = [&](double v) -> std::int64_t {
    const double q = div_zero(v - low, bw);
    if (!top_bits_zero(q) && !is_inf_or_nan(q) && sign_bit_set(q)) {
      return -1;
    }
    return static_cast<std::int64_t>(to_uint32(q));
  };
  const std::int64_t i0 = index(lo);
  const std::int64_t i1 = index(hi);
  const std::int64_t start = i1 < 0 ? i0 : std::max<std::int64_t>(i0, 0);
  const std::int64_t last = (start < nbins && i1 >= nbins) ? nbins - 1 : i1;
  const std::int64_t first = std::max<std::int64_t>(start, 0);
  if (first > last || first >= nbins) {
    return 0;
  }
  std::int64_t sum = 0;
  for (std::int64_t j = first; j < nbins; ++j) {
    sum += elem(counts, j);
    if (j >= last) {
      break;
    }
  }
  return sum;
}

// Steps of about k years.
void classify_multiyear(Timeline* tl, double lo, double hi) {
  const std::int64_t k = static_cast<std::int64_t>(to_uint32(round_away(div_zero(lo, 365.0))));
  if (k < 2 || k > 100) {
    return;
  }
  const std::int64_t nominal = static_cast<std::int64_t>(static_cast<double>(k) * 365.0 + static_cast<double>(k >> 2));
  const std::int64_t tol = static_cast<std::int64_t>(std::log2(static_cast<double>(k)) + 2.0);
  if (ranges_overlap(lo, hi, static_cast<double>(nominal - tol), static_cast<double>(nominal + tol))) {
    tl->nominal = static_cast<double>(nominal);
    tl->months = 12 * k;
  }
}

// Calendar class of a constant step; `m` is the number of steps.
void classify_constant(Timeline* tl, double step, std::int64_t m) {
  tl->step = step;
  tl->months = 0;
  tl->count = m;
  if (!tl->is_date) {
    return;
  }
  if (ranges_overlap(step, step, 30.0, 31.0)) {
    const std::int64_t lim = step == 29.5 ? 2 : 2 + (step > 29.5 ? 1 : 0);
    if (m > lim) {
      return;
    }
    tl->nominal = 30.5;
    tl->months = 1;
  } else if (ranges_overlap(step, step, 363.0, 367.0)) {
    tl->nominal = 365.0;
    tl->months = 12;
  } else if (ranges_overlap(step, step, 90.0, 92.0)) {
    const std::int64_t lim = step == 90.5 ? 2 : 2 + (step > 90.5 ? 1 : 0);
    if (m > lim) {
      return;
    }
    tl->nominal = 91.0;
    tl->months = 3;
  } else if (ranges_overlap(step, step, 181.5, 184.5)) {
    if (m > 2) {
      return;
    }
    tl->nominal = 183.0;
    tl->months = 6;
  } else if (bit_compare(step, 365.0) == 1) {
    classify_multiyear(tl, step, step);
  }
}

// Calendar class of the chosen histogram bin [lo, hi]; `rep` is the common
// step of a pure bin.
void classify_bin(Timeline* tl, double lo, double hi, bool has_rep, double rep, double low, double bw,
                  const std::vector<std::int64_t>& counts, std::int64_t nbins) {
  tl->step = has_rep ? rep : lo;
  tl->months = 0;
  if (!tl->is_date) {
    return;
  }
  const auto cnt = [&](double a, double b) { return count_in(low, bw, counts, nbins, a, b); };
  const auto set = [&](double nominal, std::int64_t months) {
    tl->nominal = nominal;
    tl->months = months;
  };
  if (ranges_overlap(lo, hi, 30.0, 31.0)) {
    const std::int64_t c1 = cnt(27.5, 31.5);
    if (c1 < 2) {
      return set(30.5, 1);
    }
    const std::int64_t c2 = cnt(27.5, 29.5);
    const std::int64_t c3 = cnt(29.5, 30.5);
    const std::int64_t c4 = cnt(30.5, 31.5);
    const std::int64_t c5 = cnt(29.5, 31.5);
    if (c4 == 0 && (c3 > 2 || c2 > 2)) {
      return;
    }
    if (c3 == 0) {
      if (c4 > 4 || c2 > 2) {
        return;
      }
      return set(30.5, 1);
    }
    if (c2 < 3) {
      return set(30.5, 1);
    }
    if (c5 > 9) {
      return set(30.5, 1);
    }
    if (c5 + c2 != c1) {
      return set(30.5, 1);
    }
    return;
  }
  if (ranges_overlap(lo, hi, 363.0, 367.0)) {
    return set(365.0, 12);
  }
  if (ranges_overlap(lo, hi, 90.0, 92.0)) {
    const std::int64_t c1 = cnt(89.5, 92.5);
    if (c1 < 2) {
      return set(91.0, 3);
    }
    const std::int64_t c2 = cnt(89.5, 90.5);
    const std::int64_t c3 = cnt(90.5, 91.5);
    const std::int64_t c4 = cnt(91.5, 92.5);
    const std::int64_t c5 = cnt(90.5, 92.5);
    if (c3 == 0 && (c4 > 3 || c2 > 2)) {
      return;
    }
    if (c4 == 0) {
      if (c3 > 3 || c2 > 2) {
        return;
      }
      return set(91.0, 3);
    }
    if (c2 < 3) {
      return set(91.0, 3);
    }
    if (c5 > 1) {
      return set(91.0, 3);
    }
    if (c5 + c2 == c1) {
      return;
    }
    return set(91.0, 3);
  }
  if (ranges_overlap(lo, hi, 181.5, 184.5)) {
    const std::int64_t c1 = cnt(89.5, 92.5);
    if (c1 < 2) {
      return set(183.0, 6);
    }
    const std::int64_t c2 = cnt(180.5, 181.5);
    const std::int64_t c3 = cnt(181.5, 182.5);
    const std::int64_t c4 = cnt(183.5, 184.5);
    if (c2 == 0 && (c4 > 2 || c3 > 2)) {
      return;
    }
    if (c4 == 0) {
      if (c2 > 2 || c3 > 2) {
        return;
      }
      return set(183.0, 6);
    }
    if (c3 < 3) {
      return set(183.0, 6);
    }
    const std::int64_t s = c4 + c2;
    if (s + c3 != c1) {
      return set(183.0, 6);
    }
    if (s < 6) {
      return;
    }
    return set(183.0, 6);
  }
  if (bit_compare(lo, 365.0) == 1) {
    classify_multiyear(tl, lo, hi);
  }
}

struct Binning {
  double bw = 0.0;
  std::int64_t nbins = 0;
  double bw0 = 0.0;
  bool capped = false;
};

// Histogram geometry for the full range of steps.
Binning make_binning(const StepStats& st) {
  double base = 0.0;
  if (st.min == 0.0) {
    base = kStepTolerance;
  } else {
    base = (bit_compare(kStepTolerance, st.min) == 1 && bit_compare(kStepTolerance, st.max) == -1) ? kStepTolerance
                                                                                                   : st.min;
  }
  base = div_zero(base, 5.0);
  const double bw0 = div_zero(base, 5.0);
  const double range = st.max - st.min;
  double q = div_zero(range, bw0);
  q = q + 1.0;
  Binning b;
  if (bit_compare(q, static_cast<double>(kMaxHistogramBins)) == 1) {
    b.bw = div_zero(range, static_cast<double>(kMaxHistogramBins - 1));
    b.nbins = kMaxHistogramBins;
    b.bw0 = bw0;
    b.capped = true;
    return b;
  }
  const std::uint32_t nb = to_uint32(q);
  b.bw = div_zero(range, static_cast<double>(nb));
  const double q2 = div_zero(range, b.bw) + 1.5;
  b.nbins = static_cast<std::int64_t>(to_uint32(q2));
  b.bw0 = b.bw;
  return b;
}

void fill_histogram(Histogram* h, const std::vector<double>& t, std::int64_t nbins, double bw, double low, double high,
                    bool bounded) {
  h->counts.assign(static_cast<std::size_t>(nbins), 0);
  h->reps.assign(static_cast<std::size_t>(nbins), 0.0);
  h->same.assign(static_cast<std::size_t>(nbins), 0);
  h->nbins = nbins;
  h->zeros = 0;
  for (std::size_t i = 1; i < t.size(); ++i) {
    const double d = t[i] - t[i - 1];
    if (top_bits_zero(d) || is_inf_or_nan(d)) {
      ++h->zeros;
      continue;
    }
    if (bounded && (bit_compare(d, low) == -1 || bit_compare(d, high) == 1)) {
      continue;
    }
    double q = d - low;
    q = div_zero(q, bw);
    std::int64_t b = static_cast<std::int64_t>(to_uint32(q));
    if (b >= nbins) {
      b = nbins - 1;
    }
    const std::int64_t c = elem(h->counts, b);
    elem(h->counts, b) = c + 1;
    if (c == 0) {
      elem(h->reps, b) = d;
      elem(h->same, b) = 1;
    } else if (std::fabs(elem(h->reps, b) - d) <= kSameStepTolerance) {
      elem(h->same, b) += 1;
    }
  }
}

// Picks the step bin: a sliding window over the histogram for wide
// histograms, the largest bin otherwise. Returns true when the bin is pure.
bool choose_bin(const Histogram& h, double low, double bw, Timeline* tl) {
  const std::vector<std::int64_t>& counts = h.counts;
  const std::int64_t nbins = h.nbins;
  std::int64_t b = 0;
  std::int64_t wtotal = 0;
  if (nbins <= 4) {
    std::int64_t total = 0;
    for (std::int64_t i = 0; i < nbins; ++i) {
      total += elem(counts, i);
    }
    std::int64_t best_i = 0;
    if (nbins >= 1 && total > 0) {
      for (std::int64_t i = 1; i < nbins; ++i) {
        if (elem(counts, i) > elem(counts, best_i)) {
          best_i = i;
        }
      }
    }
    b = best_i;
    wtotal = nbins >= 1 ? total : 0;
  } else {
    const double scaled_low = div_zero(low, bw);
    std::int64_t w8 = 0;
    std::int64_t w11 = 0;
    std::int64_t w10 = 0;
    std::int64_t w12 = 0;
    std::int64_t w13 = 1;
    std::int64_t w14 = 10;
    std::int64_t acc = elem(counts, 0);
    std::int64_t best = elem(counts, 0);
    while (true) {
      if (w13 < nbins) {
        acc += elem(counts, w13);
      }
      const std::int64_t mid = (w12 + w13) / 2;
      const double thr = static_cast<double>(w14 - (5 * w12 + mid));
      if (bit_compare(scaled_low, thr) == -1) {
        if (w12 >= 0) {
          acc -= elem(counts, w12);
        }
        ++w12;
      } else {
        --w12;
        if (w12 >= 0) {
          acc += elem(counts, w12);
        }
      }
      const std::int64_t mid2 = (w12 + w13) / 2;
      if (mid2 >= nbins) {
        break;
      }
      if (best < acc || (best == acc && elem(counts, w8) < elem(counts, mid2))) {
        best = acc;
        w8 = mid2;
        w11 = w12;
        w10 = w13;
      }
      ++w13;
      w14 += 5;
    }
    std::int64_t w28 = std::max<std::int64_t>(w11, 0);
    std::int64_t w27 = std::min<std::int64_t>(w10, nbins - 1);
    std::int64_t w25 = 0;
    std::int64_t w9 = 0;
    if (w28 <= w27) {
      w25 = w28;
      w9 = elem(counts, w28);
      for (std::int64_t i = w28 + 1; i <= w27; ++i) {
        if (w9 < elem(counts, i)) {
          w25 = i;
          w9 = elem(counts, i);
        }
      }
    }
    if (w25 != w8 && !(w9 < best - 1 && 15 * best <= 100 * (best - w9))) {
      double v = mul_flush(static_cast<double>(w25), bw) + low;
      v = div_zero(v, 10.0);
      v = div_zero(v, bw);
      const std::int64_t r = static_cast<std::int64_t>(to_uint32(v));
      w28 = std::max<std::int64_t>(w25 - r, 0);
      w27 = std::min<std::int64_t>(w25 + r, nbins - 1);
      wtotal = 0;
      for (std::int64_t i = w28; i <= w27; ++i) {
        wtotal += elem(counts, i);
      }
      b = w25;
    } else {
      b = w25 != w8 ? w8 : w25;
      wtotal = best;
    }
  }
  bool found = false;
  double rep = 0.0;
  if (elem(counts, b) >= 1 && elem(h.same, b) == elem(counts, b)) {
    rep = elem(h.reps, b);
    found = true;
  }
  const double lo = mul_flush(static_cast<double>(b), bw) + low;
  const double hi = lo + bw;
  classify_bin(tl, lo, hi, found, rep, low, bw, counts, nbins);
  tl->count = wtotal;
  return found;
}

// 4 * max(0.4 a, b).
double window_width(double a, double b) {
  double v = div_zero(a, 5.0);
  v = mul_flush(v, 2.0);
  if (bit_compare(v, b) == -1) {
    v = b;
  }
  return mul_flush(v, 4.0);
}

struct Refined {
  double low = 0.0;
  double high = 0.0;
  double bw = 0.0;
  std::int64_t nbins = 0;
  bool capped = false;
};

// Narrower histogram around the current step estimate.
Refined refine_histogram(const StepStats& st, double bw, double bw0, double step) {
  const double lowbase = st.min;
  const double width = window_width(step, bw);
  bool first = true;
  while (true) {
    const double v = div_zero(width, bw0);
    if (bit_compare(v, 2900.0) == -1) {
      break;
    }
    bw0 = mul_flush(bw0, bit_compare(v, 48000.0) == 1 ? 16.0 : 2.0);
    first = false;
  }
  double newlow = step;
  double v = div_zero(step - lowbase, bw0);
  if (!(bit_compare(v, -1.0) == -1)) {
    const std::uint32_t n0 = to_uint32(v);
    v = mul_flush(static_cast<double>(n0), bw0);
    newlow = v + lowbase;
  }
  const double w2 = window_width(newlow, bw);
  const std::uint32_t k = to_uint32(div_zero(w2, bw0));
  const double half = div_zero(mul_flush(static_cast<double>(k), bw0), 2.0);
  Refined r;
  r.low = newlow - half;
  r.high = newlow + half;
  if (bit_compare(r.low, lowbase) == -1) {
    r.low = lowbase;
  }
  if (bit_compare(r.high, st.max) == 1) {
    r.high = st.max;
  }
  r.nbins = static_cast<std::int64_t>(to_uint32(div_zero(r.high - r.low, bw0) + 1.0));
  r.bw = bw0;
  r.capped = first && bit_compare(step, bw0) == 1;
  return r;
}

// Detects the timeline step. `*irregular` is set when the series has gaps or
// duplicate timestamps and must be rebuilt.
bool detect_step(const std::vector<double>& t, const StepStats& st, Timeline* tl, bool* irregular, ErrorCode* err) {
  const std::int64_t n = static_cast<std::int64_t>(t.size());
  tl->is_date = is_date_serial(t.front()) && is_date_serial(t.back());
  tl->months = tl->is_date ? -1 : 0;
  if (st.min == st.max || (st.nonpositive != 0 && st.max == 0.0)) {
    if (st.max == 0.0) {
      *err = ErrorCode::Num;
      return false;
    }
    classify_constant(tl, st.max, n - st.nonpositive - 1);
    *irregular = st.nonpositive != 0;
    return true;
  }
  const Binning bin = make_binning(st);
  double bw = bin.bw;
  std::int64_t nbins = bin.nbins;
  double bw0 = bin.bw0;
  bool capped = bin.capped;
  double low = st.min;
  Histogram h;
  fill_histogram(&h, t, nbins, bw, low, st.max, false);
  bool found = choose_bin(h, low, bw, tl);
  if (!found && capped) {
    std::int64_t pass = 2;
    while (true) {
      const Refined r = refine_histogram(st, bw, bw0, tl->step);
      low = r.low;
      bw = r.bw;
      nbins = r.nbins;
      capped = r.capped;
      fill_histogram(&h, t, nbins, bw, low, r.high, true);
      found = choose_bin(h, low, bw, tl);
      if (found || !capped) {
        break;
      }
      if (pass >= kMaxRefinePasses) {
        *err = ErrorCode::Num;
        return false;
      }
      bw0 = bw;
      ++pass;
    }
  }
  *irregular = tl->count != n - 1;
  const std::int64_t m = n - h.zeros;
  if (!(tl->count * 100 > (m - 1) * kMinWindowPercent)) {
    *err = ErrorCode::Num;
    return false;
  }
  return true;
}

// -- grid ---------------------------------------------------------------

struct GridContext {
  const Timeline* tl;
  double step;
  AggregationMode agg;
  bool count_mode;
  int completion;
  double tol;
  double half;
};

// Maximum number of grid nodes.
std::int64_t grid_capacity(std::int64_t n, double first, double last, double step) {
  const double span = div_zero(last - first, step);
  const std::int64_t by_count = (50 * n) / 35;
  const double c = std::ceil(span);
  std::int64_t w8 = 0;
  if (bit_compare(c, static_cast<double>(kMaxGridNodes)) == 1) {
    w8 = std::max<std::int64_t>(by_count, kMaxGridNodes);
  } else {
    w8 = static_cast<std::int64_t>(to_uint32(c)) + 1;
  }
  w8 += w8 > 54 ? w8 / 5 : 10;
  std::int64_t w9 = std::min(by_count, w8);
  if (!(w8 < kMaxGridNodes)) {
    w9 = by_count;
  }
  return w9;
}

// Aggregates the points from `end` backwards while t >= `lower`; returns the
// value and the index of the group's first point.
bool aggregate_group(const GridContext& g, double lower, std::int64_t end, const std::vector<double>& t,
                     std::vector<double>* yv, double* value, std::int64_t* first, ErrorCode* err) {
  std::vector<double>& y = *yv;
  if (end <= 0 || bit_compare(lower, elem(t, end - 1)) == 1) {
    *value = g.count_mode ? 1.0 : elem(y, end);
    *first = end;
    return true;
  }
  std::int64_t j = end;
  switch (g.agg) {
    case AggregationMode::kAverage:
    case AggregationMode::kSum: {
      double acc = elem(y, end);
      std::int64_t cnt = 1;
      while (j > 0) {
        if (bit_compare(lower, elem(t, j - 1)) == 1) {
          break;
        }
        --j;
        acc = acc + elem(y, j);
        ++cnt;
      }
      *value = g.agg == AggregationMode::kAverage ? div_zero(acc, static_cast<double>(cnt)) : acc;
      break;
    }
    case AggregationMode::kCount:
    case AggregationMode::kCountA: {
      std::int64_t cnt = 1;
      while (j > 0) {
        if (bit_compare(lower, elem(t, j - 1)) == 1) {
          break;
        }
        --j;
        ++cnt;
      }
      *value = static_cast<double>(cnt);
      break;
    }
    case AggregationMode::kMax:
    case AggregationMode::kMin: {
      const int want = g.agg == AggregationMode::kMax ? -1 : 1;
      double val = elem(y, end);
      while (j > 0) {
        if (bit_compare(lower, elem(t, j - 1)) == 1) {
          break;
        }
        --j;
        if (bit_compare(val, elem(y, j)) == want) {
          val = elem(y, j);
        }
      }
      *value = val;
      break;
    }
    case AggregationMode::kMedian: {
      while (j > 0 && bit_compare(lower, elem(t, j - 1)) != 1) {
        --j;
      }
      const std::int64_t cnt = end - j + 1;
      if (cnt > 1) {
        sort_pairs(yv, nullptr, j, end);
      }
      if ((cnt & 1) != 0) {
        *value = elem(y, j + (cnt >> 1));
      } else {
        const std::int64_t m = j + (cnt >> 1);
        const double s = elem(y, m) + elem(y, m - 1);
        *value = div_zero(s, 2.0);
      }
      break;
    }
    default:
      *err = ErrorCode::Num;
      return false;
  }
  *first = j;
  return true;
}

// Index of the first of a run of equal timestamps ending at `j`.
std::int64_t group_start(const std::vector<double>& t, std::int64_t j) {
  while (j > 0 && elem(t, j - 1) == elem(t, j)) {
    --j;
  }
  return j;
}

// Index of the point nearest `node` among the groups from `j` backwards.
std::int64_t nearest_group(const std::vector<double>& t, std::int64_t j, std::int64_t floor_index, double node,
                           double* best_distance) {
  double bestd = DBL_MAX;
  std::int64_t at = j;
  while (j >= floor_index) {
    const double d = std::fabs(node - elem(t, j));
    if (bit_compare(d, bestd) >= 0) {
      break;
    }
    bestd = d;
    at = j;
    if (j == 0) {
      break;
    }
    j = group_start(t, j - 1);
  }
  *best_distance = bestd;
  return at;
}

struct GridStart {
  double node = 0.0;
  double value = 0.0;
  std::int64_t first = 0;
  double stored = 0.0;
};

// Last grid node: anchored on the run of regular steps that covers most of
// the series.
bool first_node(const GridContext& g, const std::vector<double>& t, std::vector<double>* y, GridStart* out,
                ErrorCode* err) {
  const std::int64_t last = static_cast<std::int64_t>(t.size()) - 1;
  const std::int64_t x23 = group_start(t, last);
  std::size_t maxk = 0;
  std::int64_t table[kMaxAnchorRun + 1];
  for (std::int64_t& e : table) {
    e = -1;
  }
  const auto record = [&](std::size_t run, std::int64_t runstart) {
    const std::size_t k = std::min(run, kMaxAnchorRun);
    if (table[k] < 0) {
      table[k] = runstart;
    }
    if (k > maxk) {
      maxk = k;
    }
  };
  if (x23 > 0) {
    std::size_t run = 0;
    std::int64_t runstart = -1;
    std::int64_t cur = x23;
    while (true) {
      const std::int64_t prev = group_start(t, cur - 1);
      double d = (elem(t, cur) - elem(t, prev)) - g.step;
      d = std::fabs(d);
      if (bit_compare(d, g.tol) == 1) {
        if (run >= 2 && run > maxk) {
          record(run, runstart);
        }
        runstart = -1;
        run = 0;
      } else if (run == 0) {
        run = 1;
        runstart = cur;
      } else {
        ++run;
        if (run >= kMaxAnchorRun) {
          if (table[kMaxAnchorRun] < 0) {
            table[kMaxAnchorRun] = runstart;
          }
          if (maxk <= kMaxAnchorRun - 1U) {
            runstart = -1;
            maxk = kMaxAnchorRun;
          } else {
            runstart = -1;
          }
        }
      }
      cur = prev;
      if (prev <= 0) {
        break;
      }
    }
    if (run >= 2) {
      if (run > maxk) {
        record(run, runstart);
      }
    } else if (run == 1) {
      if (maxk == 0) {
        if (table[2] < 0) {
          table[2] = runstart;
        }
        maxk = 2;
      }
    } else if (maxk == 0) {
      *err = ErrorCode::Num;
      return false;
    }
  } else if (maxk == 0) {
    *err = ErrorCode::Num;
    return false;
  }
  const double span = elem(t, last) - t.front();
  double best = 0.0;
  std::int64_t anchor = -1;
  for (std::size_t k = maxk; k >= 2U; --k) {
    const std::int64_t p = table[k];
    if (p < 0) {
      continue;
    }
    double v = div_zero(elem(t, last) - elem(t, p), span);
    v = 1.0 - v;
    v = mul_flush(v, static_cast<double>(k));
    if (bit_compare(v, best) == 1) {
      best = v;
      anchor = p;
    }
  }
  if (anchor < 0) {
    *err = ErrorCode::Value;
    return false;
  }
  const std::int64_t lastg = group_start(t, last);
  bool has_stored = false;
  double stored = 0.0;
  double cur_t = 0.0;
  if (lastg == anchor) {
    cur_t = elem(t, last);
  } else {
    if (g.tl->months >= 1) {
      std::int64_t unused = 0;
      if (!month_walk(g.tl->months, elem(t, anchor), elem(t, lastg), &cur_t, &unused, err)) {
        return false;
      }
    } else {
      double v = div_zero(elem(t, last) - elem(t, anchor), g.step);
      v = round_away(v);
      v = mul_flush(v, g.step);
      cur_t = v + elem(t, anchor);
    }
    double bestd = 0.0;
    const std::int64_t at = nearest_group(t, lastg, anchor, cur_t, &bestd);
    if (bit_compare(bestd, g.tol) == 1) {
      has_stored = true;
      stored = elem(t, at);
    } else {
      cur_t = elem(t, at);
    }
  }
  const double lower = cur_t - g.half;
  double value = 0.0;
  std::int64_t first = 0;
  if (!aggregate_group(g, lower, last, t, y, &value, &first, err)) {
    return false;
  }
  out->node = cur_t;
  out->value = value;
  out->first = first;
  out->stored = has_stored ? stored : cur_t;
  return true;
}

struct PrevNode {
  double node = 0.0;
  std::int64_t steps = 0;
  double value = 0.0;
  std::int64_t first = 0;
  double stored = 0.0;
};

// Node preceding `cur_t`, with the number of steps between them.
bool previous_node(const GridContext& g, const std::vector<double>& t, std::vector<double>* y, std::int64_t ptr,
                   double cur_t, PrevNode* out, ErrorCode* err) {
  const std::int64_t x22 = group_start(t, ptr - 1);
  const double t_prev = elem(t, x22);
  double node = 0.0;
  std::int64_t steps = 0;
  if (g.tl->months >= 1) {
    std::int64_t cnt = 0;
    if (!month_walk(g.tl->months, cur_t, t_prev, &node, &cnt, err)) {
      return false;
    }
    steps = -cnt;
  } else {
    const double tmp = cur_t - g.step;
    const double lower = tmp - g.half;
    if (bit_compare(lower, t_prev) != 1) {
      steps = 1;
      node = tmp;
    } else {
      double v = div_zero(cur_t - t_prev, g.step);
      v = round_away(v);
      steps = static_cast<std::int64_t>(to_uint32(v));
      v = mul_flush(v, g.step);
      node = cur_t - v;
    }
  }
  double bestd = 0.0;
  const std::int64_t at = nearest_group(t, x22, 0, node, &bestd);
  bool has_stored = false;
  double stored = 0.0;
  if (bit_compare(bestd, g.tol) == 1) {
    has_stored = true;
    stored = elem(t, at);
  } else {
    node = elem(t, at);
  }
  const double lower2 = node - g.half;
  double value = 0.0;
  std::int64_t first = 0;
  if (!aggregate_group(g, lower2, ptr - 1, t, y, &value, &first, err)) {
    return false;
  }
  out->node = node;
  out->steps = steps;
  out->value = value;
  out->first = first;
  out->stored = has_stored ? stored : node;
  return true;
}

// Rebuilds the series on a regular grid, filling skipped nodes per the data
// completion option. `*missing` counts the filled nodes.
bool build_grid(std::vector<double>* tv, std::vector<double>* yv, const Timeline& tl, int completion,
                AggregationMode agg, std::vector<double>* out_t, std::vector<double>* out_y, std::int64_t* missing,
                ErrorCode* err) {
  const std::vector<double>& t = *tv;
  const std::int64_t n = static_cast<std::int64_t>(t.size());
  GridContext g{&tl,
                tl.step,
                agg,
                agg == AggregationMode::kCount || agg == AggregationMode::kCountA,
                completion,
                div_zero(tl.step, 10.0),
                div_zero(tl.step, 2.0)};
  const std::int64_t cap = grid_capacity(n, t.front(), t.back(), g.step);
  std::int64_t room = cap - 1;
  *missing = 0;
  const auto append = [&](double time, double value, bool real) {
    if (room < 0 || room > cap) {
      *err = ErrorCode::Num;
      return false;
    }
    out_t->push_back(time);
    out_y->push_back(value);
    if (!real) {
      ++*missing;
    }
    --room;
    return true;
  };
  GridStart start;
  if (!first_node(g, t, yv, &start, err)) {
    return false;
  }
  double cur_t = start.node;
  double cur_y = start.value;
  std::int64_t ptr = start.first;
  if (!append(start.stored, cur_y, true)) {
    return false;
  }
  while (ptr > 0) {
    PrevNode p;
    if (!previous_node(g, t, yv, ptr, cur_t, &p, err)) {
      return false;
    }
    if (p.steps >= 2) {
      if (tl.months > 0) {
        const double kd = static_cast<double>(p.steps);
        Ymd ymd = ymd_of(cur_t);
        std::int64_t gcount = p.steps;
        while (gcount >= 2) {
          --gcount;
          const double node_t = add_months(&ymd, -tl.months);
          double value = 0.0;
          if (completion == 1) {
            const double diff = cur_y - p.value;
            double v = static_cast<double>(gcount);
            v = mul_flush(v, diff);
            v = div_zero(v, kd);
            value = v + p.value;
          }
          if (!append(node_t, value, false)) {
            return false;
          }
        }
      } else {
        const double span = cur_t - p.node;
        std::int64_t gcount = p.steps;
        double node_t = cur_t;
        double value = 0.0;
        while (gcount >= 2) {
          node_t = node_t - g.step;
          if (completion == 1) {
            const double diff = cur_y - p.value;
            double v = node_t - p.node;
            v = mul_flush(v, diff);
            v = div_zero(v, span);
            value = v + p.value;
          }
          if (!append(node_t, value, false)) {
            return false;
          }
          --gcount;
        }
      }
    }
    if (!append(p.stored, p.value, true)) {
      return false;
    }
    cur_t = p.node;
    cur_y = p.value;
    ptr = p.first;
  }
  std::reverse(out_t->begin(), out_t->end());
  std::reverse(out_y->begin(), out_y->end());
  return true;
}

}  // namespace

bool prepare_series(const std::vector<SeriesCell>& values, const std::vector<SeriesCell>& timeline, int data_completion,
                    AggregationMode aggregation, PreparedSeries* out, ErrorCode* err) {
  std::vector<double> t;
  std::vector<double> y;
  StepStats stats;
  if (!read_pairs(values, timeline, data_completion, aggregation, &t, &y, &stats, err)) {
    return false;
  }
  if (t.size() <= 1U) {
    *err = ErrorCode::Div0;
    return false;
  }
  if (stats.needs_sort) {
    sort_pairs(&t, &y, 0, static_cast<std::int64_t>(t.size()) - 1);
    stats = StepStats();
    for (std::size_t i = 1; i < t.size(); ++i) {
      stats.add(t[i] - t[i - 1]);
    }
    if (stats.needs_sort) {
      *err = ErrorCode::Value;
      return false;
    }
  }
  Timeline tl;
  bool irregular = false;
  if (!detect_step(t, stats, &tl, &irregular, err)) {
    return false;
  }
  std::vector<double> t2;
  std::vector<double> y2;
  if (irregular) {
    std::int64_t missing = 0;
    if (!build_grid(&t, &y, tl, data_completion, aggregation, &t2, &y2, &missing, err)) {
      return false;
    }
    if (kMaxMissingPercent * static_cast<std::int64_t>(t2.size()) < 100 * missing) {
      *err = ErrorCode::Num;
      return false;
    }
  } else {
    t2 = std::move(t);
    y2 = std::move(y);
  }
  if (t2.size() <= 1U) {
    *err = ErrorCode::Div0;
    return false;
  }
  tl.last = t2.back();
  out->y = std::move(y2);
  out->timeline = tl;
  return true;
}

bool target_position(const Timeline& tl, double target, double* position, bool* interpolated, ErrorCode* err) {
  *interpolated = false;
  if (tl.step == 0.0) {
    *err = ErrorCode::Value;
    return false;
  }
  if (tl.months >= 1) {
    double v = div_zero(target - tl.last, tl.nominal);
    v = std::floor(v);
    std::int64_t k = static_cast<std::int64_t>(v);
    Ymd ymd = ymd_of(tl.last);
    double node = add_months(&ymd, tl.months * k);
    while (bit_compare(node, target) == 1) {
      --k;
      node = add_months(&ymd, -tl.months);
    }
    if (node == target) {
      *position = static_cast<double>(k);
      return true;
    }
    std::int64_t extra = 0;
    double hi = node;
    while (bit_compare(hi, target) == -1) {
      hi = add_months(&ymd, tl.months);
      --extra;
    }
    if (hi == target) {
      *position = static_cast<double>(k - extra);
      return true;
    }
    if (extra != -1) {
      k += (-extra - 1);
      node = add_months(&ymd, -tl.months);
    }
    const double span = hi - node;
    v = target - node;
    v = div_zero(v, span);
    *position = v + static_cast<double>(k);
    *interpolated = true;
    return true;
  }
  const double pos = div_zero(target - tl.last, tl.step);
  const double r = round_away(pos);
  if (r == pos) {
    *position = pos;
    return true;
  }
  double d = std::fabs(r - pos);
  d = mul_flush(d, 1000.0);
  if (bit_compare(d, 1.0) == 1) {
    *position = pos;
    *interpolated = true;
    return true;
  }
  *position = r;
  return true;
}

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
