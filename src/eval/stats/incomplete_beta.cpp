//
// See incomplete_beta.h. Every arithmetic step is a separate double
// operation: multiplies and divides flush a subnormal result to zero, and
// nothing is contracted into a fused multiply-add.

#include "eval/stats/incomplete_beta.h"

#include <cmath>
#include <cstdint>
#include <cstring>
#include <limits>

#include "eval/excel_math.h"
#include "eval/stats/flushed_ops.h"
#include "eval/stats/incomplete_gamma.h"
#include "eval/stats/special_functions.h"
#include "utils/subnormal_flush.h"

namespace formulon {
namespace eval {
namespace stats {
namespace {

constexpr double kEps = 2.220446049250313e-16;
constexpr double kTol = 0x1.4p-49;
constexpr double kLnSqrt2Pi = 0.9189385332046728;
constexpr double kNaN = std::numeric_limits<double>::quiet_NaN();
constexpr double kLog07 = -0x1.6d3c324e13f4fp-2;
constexpr double kLog09 = -0x1.af8e8210a415dp-4;
constexpr double kUnder = 0x1p-970;
constexpr double kRescale = 0x1.502d2768807e1p+9;
constexpr double kTinyShape = 0x1p-104;
constexpr double kTwoOverSqrtPi = 0x1.20dd750429b6dp+0;
constexpr double kInvSqrtPi = 0x1.20dd750429b6dp-1;
constexpr double kQuarterSqrtPi = 0x1.c5bf891b4ef6bp-2;
constexpr double kTwoPowNeg1p5 = 0x1.6a09e667f3bcdp-2;
constexpr std::uint64_t kBfracTolBits = 0x3ce4000000000001ULL;
constexpr int kMaxTerms = 100000;

using flushed::bits_of;
using flushed::compensated_sum;
using flushed::div;
using flushed::mul;
using flushed::neg;
using flushed::power;
using flushed::top16_clear;
using flushed::two_sum;

// Coefficients (in b) of the 16 correction terms of the small-b expansion.
constexpr double kBgratPoly0[] = {
    0x1.33e031ce1f759p-60, -0x1.c244af3d7468cp-56, 0x1.2cc191eef4cfdp-52, -0x1.df46a5da467e5p-50,
    0x1.f30e70bf76972p-48, -0x1.5db2120ecb763p-46, 0x1.471a7e4469f90p-45, -0x1.85fa7550ca13bp-45,
    0x1.0983baa6d5df3p-45, -0x1.3557d72a19febp-47,
};
constexpr double kBgratPoly1[] = {
    0x1.03c52a05ea8b3p-52,  -0x1.37b965a0b30d7p-48, 0x1.5013e191410a8p-45,
    -0x1.a6c6a508593d7p-43, 0x1.5103e3760f09dp-41,  -0x1.59b873499de27p-40,
    0x1.b919e0ea983b5p-40,  -0x1.3bcc265ff5c3ap-40, 0x1.7da1e6add5669p-42,
};
constexpr double kBgratPoly2[] = {
    0x1.85a7bf08dfd0dp-45, -0x1.770b0e45576c3p-41, 0x1.3d3427ea06333p-38, -0x1.2f98a34d88c8bp-36,
    0x1.6026970d84b86p-35, -0x1.e9c88cc1cb587p-35, 0x1.74f5a204470d4p-35, -0x1.d6cc7574a6a5dp-37,
};
constexpr double kBgratPoly3[] = {
    0x1.ff6c2abba5c21p-38, -0x1.7f91200cbc519p-34, 0x1.ea3079df838fep-32, -0x1.52e8a3d0bd113p-30,
    0x1.0893b7c267922p-29, -0x1.b3ff4b022dc78p-30, 0x1.225c0d5895dfap-31,
};
constexpr double kBgratPoly4[] = {
    0x1.1facd8098d3d3p-30,  -0x1.43a2730abee4dp-27, 0x1.28687dae69c15p-25,
    -0x1.118c5b2482dfep-24, 0x1.f5963dc007e77p-25,  -0x1.65f59e4611336p-26,
};
constexpr double kBgratPoly5[] = {
    0x1.0db20a88f4696p-23, -0x1.af83440e53dbcp-21, 0x1.056976d41bb75p-19, -0x1.187bb909501bap-19, 0x1.b85fca40e852dp-21,
};
constexpr double kBgratPoly6[] = {
    0x1.948b0fcd6e9e0p-17,
    -0x1.a8c536fe1a8c5p-15,
    0x1.282ef4b667d84p-14,
    -0x1.0cbb766210cbbp-15,
};
constexpr double kBgratPoly7[] = {
    0x1.c71c71c71c71cp-11,
    -0x1.1111111111111p-9,
    0x1.3e93e93e93e94p-10,
};
constexpr double kBgratPoly8[] = {
    0x1.5555555555555p-5,
    -0x1.5555555555555p-5,
};
const double* const kBgratPolys[] = {kBgratPoly0, kBgratPoly1, kBgratPoly2, kBgratPoly3, kBgratPoly4,
                                     kBgratPoly5, kBgratPoly6, kBgratPoly7, kBgratPoly8};
constexpr int kBgratPolyLen[] = {10, 9, 8, 7, 6, 5, 4, 3, 2};

std::uint32_t to_uint32(double v) noexcept {
  const double av = std::fabs(v);
  return av < 4294967296.0 ? static_cast<std::uint32_t>(av) : 0xffffffffU;
}

// asin(v) = atan(v / sqrt((2 - (1 - v)) * (1 - v))) for |v| <= 1.
double asin_excel(double v) noexcept {
  const double t = 1.0 - v;
  const double u = std::sqrt(mul(2.0 - t, t));
  if (u == 0.0) {
    return v < 0.0 ? -1.5707963267948966 : 1.5707963267948966;
  }
  return std::atan(div(v, u));
}

double finish(double v, bool complement, bool upper) noexcept {
  return complement == upper ? v : (0.5 - v) + 0.5;
}

struct SeriesResult {
  double value;
  double log_value;
  bool has_log;
};

// Power series in x for I_x(a, b) with small b * x; keeps a log-scale value
// when the result underflows.
SeriesResult power_series(double a, double b, double x, double y) noexcept {
  double c = 1.0;
  double n = 1.0;
  double s = div(1.0, a);
  bool cont = true;
  for (int it = 0; n <= 2.0 || cont; ++it) {
    if (it >= kMaxTerms) {
      return {kNaN, 0.0, false};
    }
    double t = div(b, n);
    t = neg(t) + 0.5;
    t += 0.5;
    t = mul(t, x);
    c = mul(t, c);
    s = div(c, a + n) + s;
    cont = std::fabs(c) >= std::fabs(mul(s, kTol));
    n += 1.0;
  }
  const double nbl = mul(beta_log_excel(a, b), -1.0);
  double l = x >= 0.99 ? excel_log1p(0.0 - y) : std::log(x);
  l = mul(l, a);
  double sum;
  double err;
  two_sum(nbl, l, &sum, &err);
  const double r = mul(std::exp(sum), s);
  const double out = mul(std::exp(err), r);
  if (out == 0.0 && s > 0.0) {
    return {out, (std::log(s) + sum) + err, true};
  }
  return {out, 0.0, false};
}

// I_x(a, b) - I_x(a + n, b) by the forward recurrence on a, 1 <= n <= 39.
double upward_difference(double a, double b, double x, double y, int n) noexcept {
  if (n < 1 || n > 39) {
    return kNaN;
  }
  double r = div(mul(x, y), a);
  r = mul(r, x > 0.99 ? beta_density_excel(b, a, y) : beta_density_excel(a, b, x));
  if (n == 1) {
    return r;
  }
  double d[39];
  d[n - 1] = 1.0;
  const double ap1 = a + 1.0;
  const double apb = a + b;
  d[n - 2] = div(apb, ap1);
  for (int k = 2; k < n; ++k) {
    const double fk = static_cast<double>(k);
    const double t = (ap1 + fk) - 1.0;
    const double u = (apb + fk) - 1.0;
    d[n - 1 - k] = mul(div(u, t), d[n - k]);
  }
  int m = n - 1;
  if (b > 1.0 && !(y < 0.0001)) {
    double t = mul(b - 1.0, x);
    t = div(t, y) - a;
    const std::uint32_t k = to_uint32(t);
    if (static_cast<std::int64_t>(k) + 1 < n && k <= 0x7ffffffeU) {
      m = static_cast<int>(k) + 1;
    }
  }
  double w = d[n - m];
  for (int j = n - m + 1; j < n; ++j) {
    w = mul(w, x);
    w += d[j];
  }
  int j = m;
  while (true) {
    const double t = mul(power(x, static_cast<double>(j)), d[n - 1 - j]);
    w += t;
    if (t < std::fabs(mul(w, kTol))) {
      break;
    }
    if (j + 1 < n) {
      ++j;
    } else {
      break;
    }
  }
  return mul(r, w);
}

// Continued fraction for I_x(a, b) (modified Lentz).
double continued_fraction(double a, double b, double x, double y) noexcept {
  const double apb = a + b;
  double lam = neg(mul(apb, x));
  lam += a;
  const double bc = x > 0.99 ? beta_density_excel(b, a, y) : beta_density_excel(a, b, x);
  const double pre = mul(bc, mul(x, y));
  constexpr double kTiny = 1e-30;
  double f = kTiny;
  double c = kTiny;
  double d = 0.0;
  double m = 0.0;
  double an = 1.0;
  const double cc = lam + 1.0;
  double bn = mul(div(a, a + 1.0), cc);
  for (int it = 0; it < kMaxTerms; ++it) {
    d = mul(an, d) + bn;
    if (d == 0.0) {
      d = kTiny;
    }
    c = div(an, c) + bn;
    if (c == 0.0) {
      c = kTiny;
    }
    d = div(1.0, d);
    const double delta = mul(c, d);
    m += 1.0;
    f = mul(f, delta);
    if (bits_of(std::fabs(delta - 1.0)) < kBfracTolBits) {
      return mul(pre, f);
    }
    const double am = a + m;
    const double a2m = am + m;
    const double a2m1 = a2m - 1.0;
    an = div(am + -1.0, a2m1);
    const double t2 = div((apb + m) + -1.0, a2m1);
    an = mul(an, t2);
    an = mul(an, m);
    const double bmm = b - m;
    an = mul(an, bmm);
    an = mul(an, x);
    an = mul(an, x);
    const double t = mul(mul(m, x), bmm);
    bn = div(t, a2m1) + m;
    const double u = mul(y + 1.0, m) + cc;
    bn += mul(div(am, a2m + 1.0), u);
  }
  return kNaN;
}

// Correction terms of the large-a, small-b expansion, scaled by powers of z.
void small_b_terms(double b, double x, double y, double (&out)[16]) noexcept {
  const double z = x >= 0.99 ? 0.0 - excel_log1p(0.0 - y) : neg(std::log(x));
  const double z2 = mul(z, z);
  double c[9];
  for (int i = 0; i < 9; ++i) {
    const double* tab = kBgratPolys[i];
    double acc = tab[0];
    for (int k = 1; k < kBgratPolyLen[i]; ++k) {
      acc = mul(acc, b);
      acc += tab[k];
    }
    c[i] = acc;
  }
  double arr[17];
  arr[0] = b + 1.0;
  for (int i = 1; i < 17; ++i) {
    arr[i] = arr[i - 1] + 1.0;
  }
  const double zb1 = power(z, arr[0]);
  const double zb = power(z, b);
  for (int i = 0; i < 16; ++i) {
    out[i] = i % 2 == 0 ? zb : zb1;
  }
  double acc = c[0];
  for (int k = 1; k < 9; ++k) {
    acc = mul(acc, z2);
    acc += c[k];
  }
  out[15] = mul(out[15], acc);
  for (int step = 1; step < 16; ++step) {
    const int m = 15 - step;
    const int h = step >> 1;
    const int par = m & 1;
    for (int j = 0; j < 9 - h; ++j) {
      c[8 - h - j] = mul(c[8 - h - j], arr[par + 2 * j]);
    }
    acc = c[0];
    for (int k = 1; k < 9 - h; ++k) {
      acc = mul(acc, z2);
      acc += c[k];
    }
    out[m] = mul(out[m], acc);
  }
}

// Expansion for large a and small b in 1 / nu; returns 0 on total underflow.
double small_b_expansion(double a, double b, double x, double y) noexcept {
  const double nu = mul(b - 1.0, 0.5) + a;
  const double lnnu = neg(std::log(nu));
  double co[16];
  small_b_terms(b, x, y, co);
  const double lnx = x >= 0.99 ? excel_log1p(mul(y, -1.0)) : std::log(x);
  const double w = mul(lnx, nu);
  double s;
  double e1;
  double e2;
  two_sum(w, lnnu, &s, &e1);
  const double nbl = mul(beta_log_excel(a, b), -1.0);
  two_sum(s, nbl, &s, &e2);
  double p = std::exp(s);
  p = mul(std::exp(e1), p);
  p = mul(std::exp(e2), p);
  const double inv = div(1.0, nu);
  double acc = co[0];
  for (int k = 1; k < 16; ++k) {
    acc = mul(acc, inv);
    acc += co[k];
  }
  p = mul(acc, p);
  const double q = gamma_ratio_excel(b, mul(w, -1.0), true);
  if (p == 0.0 && !(q > kUnder)) {
    return 0.0;
  }
  return p + q;
}

// exp(x^2) erfc(x) by its 7-term asymptotic series divided by sqrt(pi).
double scaled_erfc_asymptotic(double x) noexcept {
  const double fx2 = mul(mul(x, x), 4.0);
  double term = 1.0;
  double s = 1.0;
  double p = 0.0;
  double q = 0.0;
  for (int i = 0; i < 6; ++i) {
    term = mul(term, -1.0);
    p += 1.0;
    term = mul(term, p);
    p += 1.0;
    term = mul(term, p);
    q += 1.0;
    term = div(term, q);
    term = div(term, fx2);
    s += term;
  }
  return mul(div(s, x), kInvSqrtPi);
}

// Uniform expansion for both a and b large; `lam` is the transition measure.
double uniform_expansion_beta(double a, double b, double lam) noexcept {
  const double bma = b - a;
  double h;
  double r1;
  double t;
  if (a < b) {
    h = div(a, b);
    r1 = div(bma, b);
    t = mul(h + 1.0, a);
  } else {
    h = div(b, a);
    r1 = div(bma, a);
    t = mul(h + 1.0, b);
  }
  const double w0 = div(1.0, std::sqrt(t));
  const double r0 = div(1.0, 1.0 + h);
  const double na = mul(-1.0, a);
  double f = mul(log1pmx(div(lam, na)), na);
  const double nb = mul(-1.0, b);
  f += mul(log1pmx(div(lam, b)), nb);
  const double tt = std::exp(0.0 - f);
  if (tt == 0.0) {
    return 0.0;
  }
  const double z0 = std::sqrt(f);
  const double z = mul(z0, 1.414213562373095);
  const double z2 = f + f;
  double a0[22] = {};
  double c[22] = {};
  double d[22] = {};
  double b0[22] = {};
  a0[0] = mul(r1, 0.6666666666666666);
  c[0] = mul(a0[0], -0.5);
  d[0] = 0.0 - c[0];
  double j0;
  if (z0 > 26.641747557046326) {
    j0 = scaled_erfc_asymptotic(z0);
  } else {
    j0 = mul(std::exp(mul(z0, z0)), erfc_excel(z0));
  }
  j0 = mul(j0, kQuarterSqrtPi);
  double j1 = kTwoPowNeg1p5;
  double s = mul(mul(j1, w0), d[0]) + j0;
  double sacc = 1.0;
  const double h2 = mul(h, h);
  double w = w0;
  double hn = 1.0;
  double znm1 = z;
  double zn = z2;
  double dn = 0.0;
  for (int n = 2;; n += 2) {
    dn += 2.0;
    hn = mul(hn, h2);
    t = mul(hn, h) + 1.0;
    t = mul(t, 2.0);
    t = mul(t, r0);
    a0[n - 1] = div(t, dn + 2.0);
    sacc += hn;
    t = div(sacc, dn + 3.0);
    t = mul(t, 2.0);
    a0[n] = mul(t, r1);
    for (int i = n; i <= n + 1; ++i) {
      const double fi = static_cast<double>(i);
      const double r = mul(fi + 1.0, -0.5);
      b0[0] = mul(r, a0[0]);
      for (int mm = 2; mm <= i; ++mm) {
        double bsum = 0.0;
        for (int j = 1; j < mm; ++j) {
          const int mmj = mm - j;
          double u = mul(static_cast<double>(j), r) - static_cast<double>(mmj);
          u = mul(u, a0[j - 1]);
          u = mul(u, b0[mmj - 1]);
          bsum = u + bsum;
        }
        const double u = mul(r, a0[mm - 1]);
        b0[mm - 1] = div(bsum, static_cast<double>(mm)) + u;
      }
      c[i - 1] = div(b0[i - 1], fi + 1.0);
      double dsum = 0.0;
      for (int j = 1; j < i; ++j) {
        dsum = mul(d[i - 1 - j], c[j - 1]) + dsum;
      }
      d[i - 1] = 0.0 - (dsum + c[i - 1]);
    }
    j0 = mul(dn - 1.0, j0);
    j0 = mul(kTwoPowNeg1p5, znm1) + j0;
    j1 = mul(j1, dn);
    j1 = mul(kTwoPowNeg1p5, zn) + j1;
    znm1 = mul(znm1, z2);
    zn = mul(zn, z2);
    w = mul(w, w0);
    const double t0 = mul(mul(w, j0), d[n - 1]);
    w = mul(w, w0);
    const double t1 = mul(mul(w, j1), d[n]);
    s = (t0 + t1) + s;
    const double tol = mul(s, kEps);
    if (!(std::fabs(t0) + std::fabs(t1) > tol) || n > 18) {
      break;
    }
  }
  const double ss = stirlerr(a + b);
  const double sa = stirlerr(a);
  const double sb = stirlerr(b);
  const double smax = a > b ? sa : sb;
  const double smin = a > b ? sb : sa;
  const double bc = (smax - ss) + smin;
  double u = std::exp(0.0 - bc);
  u = mul(u, s);
  u = mul(u, tt);
  return mul(u, kTwoOverSqrtPi);
}

// Hypergeometric pmf with (A, B, C) = (sample, successes, population).
double hyp_pmf(double a, double b, double c, double x) noexcept {
  a = std::trunc(a);
  b = std::trunc(b);
  c = std::trunc(c);
  double lo = (a - c) + b;
  if (!(lo > 0.0)) {
    lo = 0.0;
  }
  const double hi = b < a ? b : a;
  if (x < lo || x > hi) {
    return 0.0;
  }
  const double xx = std::trunc(x);
  if (b == c) {
    return xx == a ? 1.0 : 0.0;
  }
  if (c == a) {
    return xx == b ? 1.0 : 0.0;
  }
  const double amx = a - xx;
  double p = amx > xx ? div(amx, c) : div(xx, c);
  if (top16_clear(dbinom_raw(c, p, a))) {
    p = div(a, c);
  }
  double num = dbinom_raw(b, p, xx);
  num = div(num, dbinom_raw(c, p, a));
  const double t = dbinom_raw(c - b, p, amx);
  return mul(num, t);
}

// Sum of the hypergeometric pmf over 0..x by downward recurrence from the term at x.
double hyp_lower(double a, double b, double c, double x) noexcept {
  const double pm = hyp_pmf(a, b, c, x);
  if (top16_clear(pm)) {
    return 0.0;
  }
  double t = mul(x, ((c - a) - b) + x);
  double u = mul((b - x) + 1.0, (a - x) + 1.0);
  double r = div(t, u);
  double out = r;
  for (double k = x - 2.0; k >= 0.0; k -= 1.0) {
    const double j = k + 1.0;
    t = mul(j, (((neg(b) + c) - a) + k) + 1.0);
    u = mul(b - k, a - k);
    t = div(t, u);
    r = mul(r, t);
    out += r;
  }
  out += 1.0;
  out = mul(out, pm);
  if (out > 1.0) {
    return 1.0;
  }
  return out;
}

}  // namespace

double beta_log_excel(double a, double b) noexcept {
  if (a == 1.0) {
    return neg(std::log(b));
  }
  if (b == 1.0) {
    return neg(std::log(a));
  }
  const double p = a > b ? a : b;
  const double q = a > b ? b : a;
  const double s = q + p;
  if (p > 8.0) {
    const double r = div(q, p);
    const double l = excel_log1p(r);
    const double ds = stirlerr(p) - stirlerr(s);
    if (q > 8.0) {
      double t = excel_log1p(div(p, q));
      t = mul(t, q);
      t = mul(t, -1.0);
      t -= mul(p, l);
      t += kLnSqrt2Pi;
      const double v = mul(std::log(div(s, mul(q, p))), 0.5);
      t = v + t;
      t += stirlerr(q);
      return t + ds;
    }
    double t = lgamma_excel(q) + ds;
    t += mul(l, 0.5);
    t += q;
    t -= mul(std::log(s), q);
    return t - mul(excel_log1p(r), p);
  }
  const double lp = lgamma_excel(p);
  const double ls = neg(lgamma_excel(s));
  const double t = ls + lp;
  return lgamma_excel(q) + t;
}

double beta_density_excel(double a, double b, double x) noexcept {
  if (x == 0.0 || x == 1.0) {
    return 0.0;
  }
  double sum;
  double err;
  if (a > 1.0 && b > 1.0) {
    const double apb = a + b;
    const double n = apb - 2.0;
    const double k = a - 1.0;
    double lc = stirlerr(n) - stirlerr(k);
    lc -= bd0(k, mul(n, x));
    const double nk = n - k;
    lc -= stirlerr(nk);
    lc -= bd0(nk, mul(1.0 - x, n));
    const double hn = mul(std::log(n), 0.5);
    const double hk = mul(std::log(k), -0.5);
    two_sum(lc, hn, &sum, &err);
    double e;
    two_sum(sum, hk, &sum, &e);
    err += e;
    two_sum(sum, mul(std::log(nk), -0.5), &sum, &e);
    err += e;
    two_sum(sum, -kLnSqrt2Pi, &sum, &e);
    err += e;
    two_sum(sum, std::log(apb - 1.0), &sum, &e);
    err += e;
  } else {
    double l1 = excel_log1p(0.0 - x);
    l1 = mul(b - 1.0, l1);
    const double nbl = mul(beta_log_excel(a, b), -1.0);
    const double lx = mul(std::log(x), a - 1.0);
    two_sum(l1, lx, &sum, &err);
    double e;
    two_sum(sum, nbl, &sum, &e);
    err += e;
  }
  return mul(std::exp(err), std::exp(sum));
}

double beta_ratio_excel(double a, double b, double x, bool upper, double y) noexcept {
  if (x < 0.0 || x > 1.0) {
    return kNaN;
  }
  const double apb = a + b;
  if (x == 0.0) {
    return upper ? 1.0 : 0.0;
  }
  if (x == 1.0) {
    return upper ? 0.0 : 1.0;
  }
  if (a < kTinyShape && b < kTinyShape) {
    return div(upper ? a : b, apb);
  }
  if (b == 1.0) {
    if (x > 0.9) {
      return finish(mul(-1.0, excel_expm1(mul(std::log(x), a))), true, upper);
    }
    return finish(power(x, a), false, upper);
  }
  if (a == 1.0) {
    if (x < 0.1) {
      double v = excel_log1p(mul(x, -1.0));
      v = mul(v, b);
      return finish(mul(-1.0, excel_expm1(v)), false, upper);
    }
    return finish(power(1.0 - x, b), true, upper);
  }
  if (a == b && a == 0.5) {
    const bool low = x < 0.9;
    double v = asin_excel(std::sqrt(low ? x : 1.0 - x));
    v = mul(v, 2.0);
    v = div(v, std::exp(beta_log_excel(a, b)));
    return finish(v, !low, upper);
  }
  if (y < 0.0) {
    y = 1.0 + mul(x, -1.0);
  }
  double a0 = a;
  double b0 = b;
  double x0 = x;
  double y0 = y;
  bool comp;
  if (a > 1.0 && b > 1.0) {
    double lam = a < b ? a - mul(apb, x) : mul(apb, y) - b;
    if (lam < 0.0) {
      a0 = b;
      b0 = a;
      x0 = y;
      y0 = x;
      lam = std::fabs(lam);
      comp = true;
    } else {
      comp = false;
    }
    if (b0 < 40.0) {
      if (!(mul(b0, x0) > 0.7)) {
        return finish(power_series(a0, b0, x0, y0).value, comp, upper);
      }
      int n = static_cast<int>(b0);
      b0 += neg(static_cast<double>(n));
      if (b0 == 0.0) {
        --n;
        b0 = 1.0;
      }
      double w = upward_difference(b0, a0, y0, x0, n);
      double v;
      double lv = 0.0;
      bool has_log = false;
      bool series = false;
      if (!(x0 > 0.7)) {
        const SeriesResult sr = power_series(a0, b0, x0, y0);
        v = sr.value;
        lv = sr.log_value;
        has_log = sr.has_log;
        series = true;
      } else {
        if (!(a0 > 15.0)) {
          w += upward_difference(a0, b0, x0, y0, 20);
          a0 += 20.0;
        }
        v = small_b_expansion(a0, b0, x0, y0);
      }
      if (v != 0.0 || w > kUnder) {
        v += w;
      } else if (series && w > 0.0 && has_log) {
        const double t = std::exp(std::log(w) + kRescale);
        const double u = std::exp(lv + kRescale);
        v = mul(u + t, kUnder);
      }
      return finish(v, comp, upper);
    }
    const double mx = a0 > b0 ? a0 : b0;
    const double t = mul(mx, 0.03);
    const double v =
        (mx > 10000.0 && !(t < lam)) ? uniform_expansion_beta(a0, b0, lam) : continued_fraction(a0, b0, x0, y0);
    return finish(v, comp, upper);
  }
  bool swapped_comp;
  if (x > 0.5) {
    a0 = b;
    b0 = a;
    x0 = y;
    y0 = x;
    comp = true;
    swapped_comp = false;
  } else {
    comp = false;
    swapped_comp = true;
  }
  if (a > 1.0 || b > 1.0) {
    if (!(b0 > 1.0)) {
      return finish(power_series(a0, b0, x0, y0).value, comp, upper);
    }
    if (!(x0 < 0.3)) {
      return finish(power_series(b0, a0, y0, x0).value, swapped_comp, upper);
    }
    if (x0 < 0.1) {
      const double t = mul(std::log(mul(b0, x0)), a0);
      if (!(t > kLog07)) {
        return finish(power_series(a0, b0, x0, y0).value, comp, upper);
      }
    }
    double v;
    if (b0 > 15.0) {
      v = small_b_expansion(b0, a0, y0, x0);
    } else {
      const double w = upward_difference(b0, a0, y0, x0, 20);
      b0 += 20.0;
      v = small_b_expansion(b0, a0, y0, x0) + w;
    }
    return finish(v, swapped_comp, upper);
  }
  if (a0 < b0 && a0 < 0.2) {
    const double t = mul(std::log(x0), a0);
    if (t > kLog09) {
      double v;
      if (x0 < 0.3) {
        const double w = upward_difference(b0, a0, y0, x0, 20);
        b0 += 20.0;
        v = small_b_expansion(b0, a0, y0, x0) + w;
      } else {
        v = power_series(b0, a0, y0, x0).value;
      }
      return finish(v, swapped_comp, upper);
    }
  }
  return finish(power_series(a0, b0, x0, y0).value, comp, upper);
}

double beta_dist_cdf_excel(double x, double a, double b, double lo, double hi) noexcept {
  if (!(x > lo)) {
    return 0.0;
  }
  if (!(x < hi)) {
    return 1.0;
  }
  return beta_ratio_excel(a, b, div(x - lo, hi - lo), false);
}

double beta_dist_pdf_excel(double x, double a, double b, double lo, double hi) noexcept {
  if (x < lo || x > hi) {
    return 0.0;
  }
  const double w = hi - lo;
  return div(beta_density_excel(a, b, div(x - lo, w)), w);
}

double binom_cdf_excel(double k, double n, double p) noexcept {
  if (k > n) {
    return 1.0;
  }
  if (k < 0.0) {
    return 0.0;
  }
  const double kk = std::trunc(k);
  const double nn = std::trunc(n);
  if (!(kk < nn)) {
    return 1.0;
  }
  if (p == 0.0) {
    return 1.0;
  }
  if (p == 1.0) {
    return 0.0;
  }
  if (kk == 0.0) {
    return std::exp(mul(excel_log1p(mul(-1.0, p)), nn));
  }
  return beta_ratio_excel(1.0 + kk, nn - kk, p, true);
}

double binom_range_excel(double n, double p, double s1, double s2) noexcept {
  if (s2 == s1) {
    if (p == 0.0) {
      return s1 == 0.0 ? 1.0 : 0.0;
    }
    if (p == 1.0) {
      return s1 == n ? 1.0 : 0.0;
    }
    return dbinom_raw(n, p, s1);
  }
  if (s1 == 0.0) {
    return binom_cdf_excel(s2, n, p);
  }
  const double c1 = binom_cdf_excel(s1 - 1.0, n, p);
  if (c1 > 0.9) {
    const double yv = 1.0 + mul(p, -1.0);
    const double upper_from_s1 = beta_ratio_excel(1.0 + (n - s1), s1, yv, true, p);
    if (s2 == n) {
      return upper_from_s1;
    }
    const double upper_to_s2 = beta_ratio_excel(n - s2, 1.0 + s2, yv, true, p);
    return mul(upper_to_s2, -1.0) + upper_from_s1;
  }
  return mul(c1, -1.0) + binom_cdf_excel(s2, n, p);
}

double negbinom_pmf_excel(double k, double r, double p) noexcept {
  const double kk = std::trunc(k);
  const double rr = std::trunc(r);
  if (kk == 0.0) {
    return std::exp(mul(std::log(p), rr));
  }
  return mul(beta_density_excel(rr, kk + 1.0, p), div(p, kk + rr));
}

double negbinom_cdf_excel(double k, double r, double p) noexcept {
  const double kk = std::trunc(k);
  const double rr = std::trunc(r);
  if (kk == 0.0) {
    return std::exp(mul(std::log(p), rr));
  }
  return beta_ratio_excel(rr, kk + 1.0, p, false);
}

double f_dist_cdf_excel(double x, double d1, double d2, bool upper) noexcept {
  d1 = std::trunc(d1);
  d2 = std::trunc(d2);
  if (d1 < 1.0 || d2 < 1.0) {
    return kNaN;
  }
  if (x == 0.0) {
    return upper ? 1.0 : 0.0;
  }
  const double d1x = mul(d1, x);
  const double a1 = mul(d1, 0.5);
  const double a2 = mul(d2, 0.5);
  const bool swapped = d1x > d2;
  const double r = swapped ? div(d1x, d2) : div(d2, d1x);
  const double rp1 = r + 1.0;
  return beta_ratio_excel(swapped ? a2 : a1, swapped ? a1 : a2, div(1.0, rp1), upper != swapped, div(r, rp1));
}

double f_dist_pdf_excel(double x, double d1, double d2) noexcept {
  const double a = std::trunc(d1);
  const double b = std::trunc(d2);
  if (a < 1.0 || b < 1.0 || x < 0.0) {
    return kNaN;
  }
  if (x == 0.0) {
    return a < 2.0 ? kNaN : (a == 2.0 ? 1.0 : 0.0);
  }
  const double s1 = stirlerr(mul(a, 0.5));
  const double s2 = stirlerr(mul(b, 0.5));
  const double apb = a + b;
  const double s3 = mul(stirlerr(mul(apb, 0.5)), -1.0);
  constexpr double c4 = 1.2655121234846454;
  const double mn = a < b ? a : b;
  const double mx = a < b ? b : a;
  const double c5 = mul(std::log(mn) - excel_log1p(div(mn, mx)), -0.5);
  const bool big = mul(a, x) > b;
  const double x26 = big ? a : b;
  const double x27 = big ? b : a;
  const double c6a = mul(mul(std::log(x), x27), big ? 0.5 : -0.5);
  const double l2 = excel_log1p(div(x27, x26));
  const double ax = mul(a, x);
  const double ratio = big ? div(b, ax) : div(ax, b);
  const double c6 = mul(mul(excel_log1p(ratio) - l2, apb), 0.5) + c6a;
  const double terms[] = {s1, s2, s3, c4, c5, c6};
  double sum;
  double err;
  compensated_sum(terms, 6, &sum, &err);
  const double e = div(std::exp(neg(err)), x);
  return mul(std::exp(neg(sum)), e);
}

double t_dist_pdf_excel(double x, double df) noexcept {
  df = std::trunc(df);
  if (df < 1.0) {
    return kNaN;
  }
  const double x2 = mul(x, x);
  if (df == 1.0) {
    return div(1.0, mul(3.141592653589793, 1.0 + x2));
  }
  if (df == 2.0) {
    return std::exp(mul(std::log(x2 + 2.0), -1.5));
  }
  const double n1h = div(df + 1.0, 2.0);
  const double nh = div(df, 2.0);
  const double d = stirlerr(n1h) - stirlerr(nh);
  const double t = d - bd0(nh, n1h);
  double u;
  if (x2 < df) {
    const double h = div(x2, 2.0);
    const double m = div(x2 + df, 2.0);
    u = h - bd0(nh, m);
  } else {
    u = mul(nh, excel_log1p(div(x2, df)));
  }
  const double r = div(0.3989422804014327, std::sqrt(div(x2, df) + 1.0));
  return mul(std::exp(t - u), r);
}

double t_dist_2t_excel(double x, double df) noexcept {
  df = std::trunc(df);
  if (x < 0.0 || df < 1.0) {
    return kNaN;
  }
  const double a = div(df, 2.0);
  const double x2 = mul(x, x);
  const double s = x2 + df;
  return beta_ratio_excel(a, 0.5, div(df, s), false, div(x2, s));
}

double t_dist_tail_excel(double x, double df, bool upper) noexcept {
  df = std::trunc(df);
  if (df < 1.0 || df > 1e10) {
    return kNaN;
  }
  const bool keep = (!upper) != (x >= 0.0);
  const double v = div(t_dist_2t_excel(std::fabs(x), df), 2.0);
  return keep ? v : 1.0 - v;
}

double hypgeom_pmf_excel(double x, double n, double m, double n_pop) noexcept {
  return hyp_pmf(n, m, n_pop, x);
}

double hypgeom_cdf_excel(double x, double n, double m, double n_pop) noexcept {
  const double a = std::trunc(n);
  const double b = std::trunc(m);
  const double c = std::trunc(n_pop);
  double lo = (a - c) + b;
  if (!(lo > 0.0)) {
    lo = 0.0;
  }
  const double hi = b < a ? b : a;
  if (x < lo) {
    return 0.0;
  }
  if (!(x < hi)) {
    return 1.0;
  }
  const double xx = std::trunc(x);
  if (b == c) {
    return xx < a ? 0.0 : 1.0;
  }
  if (c == a) {
    return xx < b ? 0.0 : 1.0;
  }
  const double mean = mul(a, div(b, c));
  if (xx > mean) {
    return 1.0 - hyp_lower(a, c - b, c, (a - xx) - 1.0);
  }
  return hyp_lower(a, b, c, xx);
}

}  // namespace stats
}  // namespace eval
}  // namespace formulon
