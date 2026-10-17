//
// See inverse_distributions.h. Every arithmetic step is a separate double
// operation: multiplies and divides flush a subnormal result to zero, and
// nothing is contracted into a fused multiply-add.

#include "eval/stats/inverse_distributions.h"

#include <cmath>
#include <cstdint>
#include <limits>

#include "eval/excel_math.h"
#include "eval/stats/flushed_ops.h"
#include "eval/stats/incomplete_beta.h"
#include "eval/stats/incomplete_gamma.h"
#include "eval/stats/special_functions.h"

namespace formulon {
namespace eval {
namespace stats {
namespace {

constexpr double kNaN = std::numeric_limits<double>::quiet_NaN();
constexpr double kEps = 2.220446049250313e-16;
constexpr double kHalfEps = 1.1102230246251565e-16;
constexpr double kPi = 3.141592653589793;
constexpr double kDblMax = 1.7976931348623157e308;
constexpr double kDblMin = 2.2250738585072014e-308;
constexpr double kTinyShape = 0x1p-104;
constexpr int kGammaIterations = 106;

using flushed::compensated_sum;
using flushed::div;
using flushed::mul;
using flushed::neg;
using flushed::power;

// Seven quadratics in u = 1 / (2a), the coefficients of the initial-guess polynomial.
constexpr double kGuessPoly0[] = {
    0x1.9a573fe90a565p-10,
    -0x1.7c23b45c8ae7cp-10,
    0x1.9a394ef09be68p-14,
};
constexpr double kGuessPoly1[] = {
    0x1.15fcca1e2460ep-10,
    0x1.4b9d0c6ac7742p-9,
    -0x1.b9cec18f440fbp-11,
};
constexpr double kGuessPoly2[] = {
    -0x1.c77f9fb72ec4bp-8,
    0x1.52cc66dc94265p-8,
    0x1.99ad0fc31695ap-9,
};
constexpr double kGuessPoly3[] = {
    -0x1.b19ddcca45deep-7,
    -0x1.7a107ccfdf50ep-7,
    -0x1.184457d155d12p-7,
};
constexpr double kGuessPoly4[] = {
    0x1.752cc80a1a603p-6,
    0x1.71af19de7870dp-7,
    0x1.1ab3b3c1d0ffcp-13,
};
constexpr double kGuessPoly5[] = {
    -0x1.26741717e3656p-7,
    0x1.ab24fc8eb15e3p-6,
    0x1.e2b5228c1c029p-2,
};
constexpr double kGuessPoly6[] = {
    -0x1.f01378d8aed84p-7,
    -0x1.ca3684f74b25ap-3,
    0x1.0005ce761a389p+0,
};

// log(v), or log(1 - v) when `flag` is set.
double log_or_log1m(double v, bool flag) noexcept {
  return flag ? excel_log1p(mul(v, -1.0)) : std::log(v);
}

// Wilson-Hilferty style polynomial correction in u and the normal deviate z.
double guess_polynomial(double u, double z) noexcept {
  const double* const tabs[] = {kGuessPoly0, kGuessPoly1, kGuessPoly2, kGuessPoly3,
                                kGuessPoly4, kGuessPoly5, kGuessPoly6};
  double c[7];
  for (int i = 0; i < 7; ++i) {
    double acc = tabs[i][0];
    acc = mul(acc, u);
    acc += tabs[i][1];
    acc = mul(acc, u);
    acc += tabs[i][2];
    c[i] = acc;
  }
  const double w = mul(std::sqrt(u), z);
  double acc = c[0];
  for (int k = 1; k < 7; ++k) {
    acc = mul(acc, w);
    acc += c[k];
  }
  return mul(acc, mul(acc, acc));
}

// Starting point for the gamma inversion; `p` is the lower-tail target, or
// `q` (when `has_q`) the upper-tail target with `p = 1 - q`.
double gamma_guess(double a, double p, double q, bool has_q) noexcept {
  const double target = has_q ? q : p;
  const double lp = log_or_log1m(target, has_q);
  if (mul(lp, -0.62) > a) {
    double x = lgamma_excel(a) + lp;
    x += std::log(a);
    return std::exp(div(x, a));
  }
  const double u = div(0.5, a);
  double z = normal_quantile(target);
  if (has_q) {
    z = mul(z, -1.0);
  }
  if (a > 1.0) {
    const double x0 = mul(guess_polynomial(u, z), a);
    if (mul(x0, 2.2) + 3.0 > x0) {
      return x0;
    }
    const double lq = log_or_log1m(target, !has_q);
    const double t = mul(std::log(x0), a - 1.0);
    return (lgamma_excel(a) + lq) + t;
  }
  const double tt = div(1.0, mul(a, 9.0));
  const double om = 1.0 - tt;
  double r = mul(std::sqrt(tt), z);
  r += om;
  return mul(power(r, 3.0), a);
}

// Starting point for the beta inversion (normal-deviate based).
double beta_guess(double a, double b, double p) noexcept {
  const double y = 0.0 - normal_quantile(p);
  const double b2 = mul(b, 2.0);
  if (a > 1.0 && b > 1.0) {
    const double r = div(mul(y, y) - 3.0, 6.0);
    const double s = div(1.0, mul(a, 2.0) - 1.0);
    const double t = div(1.0, b2 - 1.0);
    const double h = div(2.0, t + s);
    const double d = t - s;
    const double w3 = div(2.0, mul(3.0, h));
    double v = (r + 0.8333333333333334) - w3;
    v = mul(v, d);
    double w = std::sqrt(r + h);
    w = mul(w, y);
    w = div(w, h) - v;
    w = mul(w, 2.0);
    w = mul(std::exp(w), b);
    return div(a, a + w);
  }
  const double tt = div(1.0, mul(b, 9.0));
  double v = mul(std::sqrt(tt), y);
  v += 1.0 - tt;
  v = power(v, 3.0);
  const double w = mul(v, b2);
  const double t = (mul(a, 4.0) + b2) - 2.0;
  if (w < 0.0) {
    const double l = excel_log1p(mul(p, -1.0));
    const double bl = beta_log_excel(a, b);
    double e = std::log(b) + bl;
    e += l;
    return mul(excel_expm1(div(e, b)), -1.0);
  }
  if (!(t > w)) {
    const double lp = std::log(p);
    const double bl = beta_log_excel(a, b);
    double e = std::log(a) + bl;
    e += lp;
    return std::exp(div(e, a));
  }
  return 1.0 - div(2.0, div(t, w) + 1.0);
}

}  // namespace

double gamma_inverse_excel(double a, double p, bool upper) noexcept {
  if (a == 1.0) {
    const double l = (upper && p > 0.5) ? log_or_log1m(1.0 - p, true) : log_or_log1m(p, !upper);
    return mul(l, -1.0);
  }
  if (a == 0.5) {
    const double v = upper ? erfinv_excel((0.5 - p) + 0.5) : erfinv_excel(p);
    return mul(v, v);
  }
  const bool flip = p > 0.9999;
  const double t = flip ? 1.0 - p : p;
  const bool tail = upper != flip;
  const double x0 = upper == flip ? gamma_guess(a, t, 0.0, false) : gamma_guess(a, 1.0 - t, t, true);
  if (x0 == 0.0) {
    return 0.0;
  }
  double best_err = kDblMax;
  double best_x = x0;
  double x = x0;
  for (int it = 0; it < kGammaIterations; ++it) {
    const double xi = x;
    const double d = gamma_ratio_excel(a, xi, tail) - t;
    const double ad = std::fabs(d);
    if (ad < best_err) {
      best_err = ad;
      best_x = xi;
    }
    double pdf = mul(div(a, xi), dpois_raw(xi, a));
    if (tail) {
      pdf = mul(pdf, -1.0);
    }
    if (pdf == 0.0) {
      return xi;
    }
    const double g = mul(div(a - 1.0, xi) - 1.0, d);
    const double den = mul(2.0, pdf) - g;
    x = xi - div(mul(2.0, d), den);
    if (!(std::fabs(mul(xi, kEps)) < std::fabs(x - xi))) {
      return xi;
    }
  }
  return div(best_err, t) > 1e-12 ? kNaN : best_x;
}

double gamma_inv_excel(double p, double a, double b) noexcept {
  return mul(gamma_inverse_excel(a, p, false), b);
}

double chisq_inv_excel(double p, double df, bool upper) noexcept {
  df = std::trunc(df);
  return mul(gamma_inverse_excel(mul(df, 0.5), p, upper), 2.0);
}

double beta_inverse_excel(double a, double b, double p, bool upper, double* complement) noexcept {
  double scratch;
  double* const comp = complement != nullptr ? complement : &scratch;
  if (p == 0.0) {
    *comp = upper ? 0.0 : 1.0;
    return upper ? 1.0 : 0.0;
  }
  if (p == 1.0) {
    *comp = upper ? 1.0 : 0.0;
    return upper ? 0.0 : 1.0;
  }
  const double q = (0.5 - p) + 0.5;
  const bool flip = p > 0.5;
  const double sa = flip ? b : a;
  const double sb = flip ? a : b;
  const double target = flip ? q : p;
  double x = beta_guess(sa, sb, target);
  if (!(x < 1.0)) {
    x = 0.999999;
  } else if (!(x > 0.0)) {
    x = 1e-06;
  }
  double e = div(-5.0, mul(sa, sa)) + -13.0;
  e -= power(target, -0.2);
  e = mul(e, 2.0);
  if (e < -300.0) {
    e = -300.0;
  }
  double acu = power(10.0, e);
  if (acu > kTinyShape) {
    acu = kTinyShape;
  }
  double yprev = 0.0;
  double sq_prev = 1.0;
  double prev = 1.0;
  double best_x = x;
  double xprev = x;
  double d = beta_ratio_excel(sa, sb, x, false) - target;
  double f = beta_density_excel(sa, sb, x);
  double best_d = std::fabs(d);
  int ndiv = 0;
  int count = 1;
  bool exhausted = false;
  if (f != 0.0) {
    while (true) {
      const double yy = div(mul(d, yprev), f);
      if (!(yy > 0.0)) {
        prev = sq_prev == 0.0 ? kDblMin : sq_prev;
      }
      if (!(prev > acu)) {
        break;
      }
      const double f2 = mul(f, f);
      if (f2 == 0.0) {
        if (!(std::log(acu) < mul(std::log(std::fabs(d)) - std::log(std::fabs(f)), 2.0))) {
          break;
        }
      } else {
        const double y2 = div(mul(d, d), f2);
        if (!(y2 > acu)) {
          if (y2 != 0.0 || d == 0.0) {
            break;
          }
          if (!(std::log(acu) < mul(std::log(std::fabs(d)) - std::log(std::fabs(f)), 2.0))) {
            break;
          }
        }
      }
      double g = 1.0;
      int inner = (ndiv > 699 ? ndiv : 699) - ndiv + 1;
      ++ndiv;
      bool accepted = false;
      double tx = x;
      while (true) {
        if (--inner == 0) {
          break;
        }
        const double adj = div(mul(g, d), f);
        sq_prev = mul(adj, adj);
        if (sq_prev < prev) {
          tx = x - adj;
          if (tx > 0.0 && tx < 1.0) {
            accepted = true;
            break;
          }
        }
        ++ndiv;
        g = div(g, 3.0);
      }
      if (!accepted) {
        exhausted = true;
        break;
      }
      if (x == tx) {
        break;
      }
      x = tx;
      yprev = div(d, f);
      if (!(std::fabs(mul(xprev, kHalfEps)) < std::fabs(x - xprev))) {
        break;
      }
      if (count > 500) {
        break;
      }
      xprev = x;
      d = beta_ratio_excel(sa, sb, x, false) - target;
      f = beta_density_excel(sa, sb, x);
      const double ad = std::fabs(d);
      if (ad < best_d) {
        best_d = ad;
        best_x = x;
      }
      ++count;
      if (f == 0.0) {
        break;
      }
    }
    if ((exhausted || count >= 501) && !(std::fabs(div(best_d, target)) < 1e-12)) {
      return kNaN;
    }
  }
  const double out = upper == flip ? best_x : (0.5 - best_x) + 0.5;
  const double out2 = upper == flip ? (0.5 - best_x) + 0.5 : best_x;
  if (out > 1.0) {
    *comp = 0.0;
    return 1.0;
  }
  if (out < 0.0) {
    *comp = 1.0;
    return 0.0;
  }
  *comp = out2;
  return out;
}

double beta_inv_excel(double p, double a, double b, double lo, double hi) noexcept {
  if (p < 0.0 || p > 1.0) {
    return kNaN;
  }
  if (p == 0.0) {
    p = 1e-09;
  } else if (p == 1.0) {
    p = 0.999999999;
  }
  return mul(hi - lo, beta_inverse_excel(a, b, p, false)) + lo;
}

double t_inv_excel(double p, double df) noexcept {
  df = std::trunc(df);
  if (df < 1.0 || df > 1e10 || !(p < 1.0)) {
    return kNaN;
  }
  if (p == 0.5) {
    return 0.0;
  }
  const double q = (0.5 - p) + 0.5;
  double out;
  if (df == 1.0) {
    out = div(1.0, std::tan(mul(p < 0.5 ? p : q, kPi)));
  } else {
    const double pp = p < q ? mul(2.0, p) : mul(2.0, q);
    const double a = mul(df, 0.5);
    if (pp > 0.9) {
      const double x = beta_inverse_excel(a, 0.5, pp, true);
      out = std::sqrt(div(mul(x, df), 1.0 - x));
    } else {
      const double x = beta_inverse_excel(a, 0.5, pp, false);
      out = std::sqrt(mul(div(1.0, x) - 1.0, df));
    }
  }
  return p < 0.5 ? neg(out) : out;
}

double t_inv_2t_excel(double p, double df) noexcept {
  return mul(t_inv_excel(div(p, 2.0), df), -1.0);
}

double f_inv_excel(double p, double d1, double d2, bool upper) noexcept {
  d1 = std::trunc(d1);
  d2 = std::trunc(d2);
  if (d1 < 1.0 || d2 < 1.0) {
    return kNaN;
  }
  if (upper ? p == 1.0 : p == 0.0) {
    return 0.0;
  }
  const double a1 = mul(d1, 0.5);
  const double a2 = mul(d2, 0.5);
  const double ld2 = std::log(d2);
  double z2;
  const double z = upper ? beta_inverse_excel(a2, a1, p, true, &z2) : beta_inverse_excel(a1, a2, p, false, &z2);
  double l1;
  double l2;
  if (z2 < z) {
    l1 = excel_log1p(mul(z2, -1.0));
    l2 = std::log(z2);
  } else {
    l1 = std::log(z);
    l2 = excel_log1p(mul(z, -1.0));
  }
  const double terms[] = {ld2, l1, mul(std::log(d1), -1.0), mul(l2, -1.0)};
  double sum;
  double err;
  compensated_sum(terms, 4, &sum, &err);
  return mul(std::exp(sum), std::exp(err));
}

double binom_inv_excel(double n, double p, double alpha) noexcept {
  const double nd = std::trunc(n);
  if (binom_cdf_excel(0.0, nd, p) > alpha) {
    return 0.0;
  }
  std::int64_t hi = static_cast<std::int64_t>(nd) + 1;
  std::int64_t mid = hi >> 1;
  if (hi < 2) {
    return static_cast<double>(mid + 1);
  }
  std::int64_t lo = 0;
  while (true) {
    if (binom_cdf_excel(static_cast<double>(mid), nd, p) < alpha) {
      lo = mid;
      mid = hi;
    }
    hi = mid;
    mid = (lo + hi) >> 1;
    if (!(mid > lo)) {
      break;
    }
  }
  return static_cast<double>(mid + 1);
}

}  // namespace stats
}  // namespace eval
}  // namespace formulon
