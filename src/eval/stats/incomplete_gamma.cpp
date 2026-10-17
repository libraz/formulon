//
// See incomplete_gamma.h. Every arithmetic step is a separate double
// operation: multiplies and divides flush a subnormal result to zero, and
// nothing is contracted into a fused multiply-add.

#include "eval/stats/incomplete_gamma.h"

#include <cmath>
#include <cstdint>
#include <cstring>
#include <limits>

#include "eval/excel_math.h"
#include "eval/stats/horner.h"
#include "eval/stats/special_functions.h"
#include "utils/subnormal_flush.h"

namespace formulon {
namespace eval {
namespace stats {
namespace {

constexpr double kEps = 2.220446049250313e-16;
constexpr double kLnSqrt2Pi = 0.9189385332046728;
constexpr double kNaN = std::numeric_limits<double>::quiet_NaN();
constexpr int kMaxTerms = 100000;

// Tables: 1/sqrt(pi), 1/sqrt(2 pi), 2 pi, -ln 2, the continued-fraction
// tolerance; then the coefficient sets, highest degree first.
constexpr double kInvSqrtPi = 0x1.20dd750429b6dp-1;
constexpr double kPhi0 = 0x1.9884533d43651p-2;
constexpr double kTwoPi = 0x1.921fb54442d18p+2;
constexpr double kNegLn2 = -0x1.62e42fefa39efp-1;
constexpr double kCfTol = 0x1.cd2b297d889bcp-54;
constexpr double kHalfIntSeries[] = {
    0x1.799ec4c4db4d0p-63, 0x1.e3d36c1c38faap-59,
    0x1.26d4d5e132b8cp-54, 0x1.54e6174c62a5ap-50,
    0x1.74dba97b8be52p-46, 0x1.808286c768445p-42,
    0x1.747e72912d023p-38, 0x1.519297d390c9fp-34,
    0x1.1cd3b01a822a6p-30, 0x1.bd0ac3296b624p-27,
    0x1.3fdfbc45c52eap-23, 0x1.a3d5a71b92cd3p-20,
    0x1.f28db670be53bp-17, 0x1.08db48ebe51c7p-13,
    0x1.f09b28ba4d955p-11, 0x1.937e11175f095p-8,
    0x1.1566abc011567p-5,  0x1.3813813813814p-3,
    0x1.1111111111111p-1,  0x1.5555555555555p+0,
    0x1.0000000000000p+1,  0x0.0p+0,
};
constexpr double kRecipFactorial[] = {
    0x1.71b8ef6dcf572p-66, 0x1.e542ba4020225p-62, 0x1.2f49b46814157p-57, 0x1.6827863b97d97p-53, 0x1.952c77030ad4ap-49,
    0x1.ae7f3e733b81fp-45, 0x1.ae7f3e733b81fp-41, 0x1.93974a8c07c9dp-37, 0x1.6124613a86d09p-33, 0x1.1eed8eff8d898p-29,
    0x1.ae64567f544e4p-26, 0x1.27e4fb7789f5cp-22, 0x1.71de3a556c734p-19, 0x1.a01a01a01a01ap-16, 0x1.a01a01a01a01ap-13,
    0x1.6c16c16c16c17p-10, 0x1.1111111111111p-7,  0x1.5555555555555p-5,  0x1.5555555555555p-3,  0x1.0000000000000p-1,
    0x1.0000000000000p+0,  0x1.0000000000000p+0,
};
constexpr double kLgNum23[] = {
    0x1.5e5b7384217a4p-9,  0x1.0ee2f118852d6p-2,  0x1.572bf86842eb2p+2,  0x1.09fd1ad1d8e61p+5,  0x1.f0a8c053986d7p+5,
    -0x1.08305b4799886p+1, -0x1.0ae34eb3cbad5p+6, -0x1.de852aee8c14cp+4, -0x1.22260ed9784e4p+1,
};
constexpr double kLgDen23[] = {
    0x1.03cadf2cc735dp-11, 0x1.52fc9914bdb8bp-4, 0x1.5303658ef44abp+1, 0x1.bb43b464d3960p+4, 0x1.c515177eb1641p+6,
    0x1.7b39fc9194438p+7,  0x1.eddf2040d2bc4p+6, 0x1.9b252b2fab34ap+4, 0x1.0000000000000p+0,
};
constexpr double kLgNum15[] = {
    0x1.812a831074da4p-8,   -0x1.6a4e459cd3e24p-1,  -0x1.af42211d48214p+6,  -0x1.6c8e7686beff4p+10,
    -0x1.59f0533cbdeefp+12, -0x1.bcdc3471f69ecp+12, -0x1.77c148735cba9p+11, -0x1.38e8ce02baa26p+8,
};
constexpr double kLgDen15[] = {
    -0x1.13d5d4770ae5fp+5,  -0x1.9aaecfc22ba41p+9,  -0x1.4687e0abed9d9p+12, -0x1.6ef433dc482a1p+13,
    -0x1.30aae53746110p+13, -0x1.4be7c906ff0a9p+11, -0x1.3c4a87ae3d483p+7,  0x1.0000000000000p+0,
};
constexpr double kLgNum07[] = {
    -0x1.33ada3fabee89p-6, -0x1.c3a9d0165803ap+2, -0x1.3038da170c471p+7, -0x1.6d0d4bff8a504p+9, -0x1.59e7c382c45e3p+8,
    0x1.4c00f820ca19dp+11, 0x1.d3f1370f347b5p+11, 0x1.8558f6ad8f440p+10, 0x1.92587e5c124dap+7,  0x1.59f97cb258253p+2,
};
constexpr double kLgDen07[] = {
    0x1.ebffdda5dd26dp+0,  0x1.35397cc538025p+6, 0x1.8a4a667ddba05p+9, 0x1.6e6b1d4f636f5p+11, 0x1.18e838fb35c8fp+12,
    0x1.6b0b4a0835670p+11, 0x1.72eb27395650ep+9, 0x1.f168c3e77cf12p+5, 0x1.0000000000000p+0,
};
constexpr double kStirlingPoly[] = {
    -0x1.6476701181f35p+0, 0x1.6fe96381e0680p-3,   -0x1.e4286cb0f5398p-6, 0x1.a41a41a41a41ap-8,  -0x1.f6ab0d9993c7dp-10,
    0x1.b951e2b18ff23p-11, -0x1.3813813813814p-11, 0x1.a01a01a01a01ap-11, -0x1.6c16c16c16c17p-9, 0x1.5555555555555p-4,
};
constexpr double kStirlingSmall[] = {
    0x1.4c071bcda0a57p-4, 0x1.52a9b923ea63bp-5, 0x1.c579a268d80a8p-6,
    0x1.54a2662fd788ep-6, 0x1.10b4e513fcbd2p-6, 0x1.c6b167bebdf1ap-7,
};
constexpr double kRecipGammaNumLow[] = {
    0x1.351e6d63b8846p-11, -0x1.51702b49eea17p-8, 0x1.f6a3e892940b3p-8, 0x1.e949b72333263p-5,
    -0x1.d9099ed7b807cp-3, -0x1.a2e564d1ddc83p-2, 0x1.2788cfc6fb61ap-1,
};
constexpr double kRecipGammaDenLow[] = {
    0x1.15609c2d46a68p-8, 0x1.abd6b54cfcf21p-6, 0x1.44824f78abecbp-3, 0x1.b5d4cf04077e4p-2, 0x1.0000000000000p+0,
};
constexpr double kRecipGammaNumHigh[] = {
    -0x1.163d4cc01152cp-13, 0x1.1773a76d8dc66p-12, 0x1.245a6033d5b04p-9,  -0x1.839dad69b4a44p-7, 0x1.e7c67130498bap-11,
    0x1.e4e15e2097702p-4,   -0x1.f5438f2892723p-3, -0x1.8aebd0e9cdd50p-1, -0x1.b0ee6072093dep-2,
};
constexpr double kRecipGammaDenHigh[] = {
    0x1.ca42502a26800p-5,
    0x1.17a14539cb550p-2,
    0x1.0000000000000p+0,
};
constexpr double kTemmeSeed[] = {
    -0x1.5ff773ccd8f52p-58, 0x1.af2c06678a063p-57,  0x1.75713641cd216p-59,  -0x1.7cfbcf3db9bfcp-53,
    0x1.d2e7d5ca48b90p-51,  -0x1.1ca914d71a27cp-49, -0x1.283fe7950ad7bp-51, 0x1.0104fc4369a3cp-45,
    -0x1.3989bebb193c0p-43, 0x1.7ba0759769d7cp-42,  0x1.ef98008f5eec2p-44,  -0x1.61ca701fd754ap-38,
    0x1.ac9475c463659p-36,  -0x1.0070a87340428p-34, -0x1.c0d9b6edf2b0bp-36, 0x1.f6e66d24d5c8ap-31,
    -0x1.2d2197c7a2faap-28, 0x1.6097d55c37c1cp-27,  0x1.ccf5ceb7f0d9fp-28,  -0x1.7b5f9a2d0465cp-23,
    0x1.bd6d21e4b4109p-21,  -0x1.f1b22f594c6b5p-20, -0x1.255370652afc1p-19, 0x1.48c5892f7cd83p-15,
    -0x1.76e06fec7273bp-13, 0x1.71de3a556c734p-12,  0x1.2f684bda12f68p-10,  -0x1.e573ac901e574p-7,
    0x1.5555555555555p-4,   -0x1.5555555555555p-2,
};

double mul(double a, double b) noexcept {
  return flush_subnormal(a * b);
}

double div(double a, double b) noexcept {
  return flush_subnormal(a / b);
}

// Negation that leaves zero unchanged.
double neg_nz(double v) noexcept {
  return v != 0.0 ? -v : v;
}

// Saturating truncation to int.
int to_int(double v) noexcept {
  return std::fabs(v) < 2147483648.0 ? static_cast<int>(v) : std::numeric_limits<int>::min();
}

std::int64_t bits_of(double v) noexcept {
  std::int64_t b;
  std::memcpy(&b, &v, sizeof b);
  return b;
}

// (n + 0.5) * log1p(1/n) - 1.
double stirlerr_term(double n) noexcept {
  return mul(n + 0.5, excel_log1p(div(1.0, n))) - 1.0;
}

// prod_{i=0}^{count-1} (x - count + i).
double shifted_product(double x, int count) noexcept {
  double out = 1.0;
  double acc = x - static_cast<double>(count);
  for (int i = 0; i < count; ++i) {
    out = mul(acc, out);
    acc += 1.0;
  }
  return out;
}

void two_sum(double a, double b, double* s, double* e) noexcept {
  *s = a + b;
  const double bb = *s - a;
  *e = (a - (*s - bb)) + (b - bb);
}

// 1 / Gamma(1 + a) - 1 by rationals on a < 0.5 and a >= 0.5.
double recip_gamma_m1(double a) noexcept {
  if (a < 0.5) {
    const double q = horner(kRecipGammaDenLow, a);
    const double p = mul(horner(kRecipGammaNumLow, a), a);
    return div(p, q);
  }
  const double u = a - 1.0;
  const double v = div(u, a);
  const double q = horner(kRecipGammaDenHigh, u);
  const double p = mul(horner(kRecipGammaNumHigh, u), v);
  return div(p, q);
}

// Lower series sum_k x^k / ((a+1)...(a+k)) times the Poisson prefactor.
double series_lower(double a, double x) noexcept {
  double ap = a + 1.0;
  double term = div(x, ap);
  double s = term + 1.0;
  for (int it = 0; it < kMaxTerms; ++it) {
    ap += 1.0;
    term = mul(div(x, ap), term);
    s += term;
    if (mul(s, kEps) > term) {
      return mul(s, dpois_raw(x, a));
    }
  }
  return kNaN;
}

// Upper asymptotic series in (a - k) / x.
double asymptotic_upper(double a, double x) noexcept {
  double am = a - 1.0;
  double term = div(am, x);
  double s = term + 1.0;
  for (int it = 0; it < kMaxTerms; ++it) {
    am -= 1.0;
    term = mul(div(am, x), term);
    s += term;
    if (std::fabs(mul(s, kEps)) >= std::fabs(term)) {
      const double r = mul(div(a, x), s);
      return mul(r, dpois_raw(x, a));
    }
  }
  return kNaN;
}

// Upper tail by the Legendre continued fraction (modified Lentz).
double continued_fraction_upper(double a, double x) noexcept {
  const double d0 = dpois_raw(x, a);
  double f = 1e-30;
  double d = 0.0;
  double c = 1e-30;
  double k = 1.0;
  double b = x;
  double an = a;
  bool odd = false;
  for (int it = 0; it < kMaxTerms; ++it) {
    d = b + mul(an, d);
    if (d == 0.0) {
      d = 1e-30;
    }
    d = div(1.0, d);
    c = b + div(an, c);
    if (c == 0.0) {
      c = 1e-30;
    }
    const double delta = mul(c, d);
    f = mul(f, delta);
    if (std::fabs(delta - 1.0) < kCfTol) {
      return mul(f, d0);
    }
    if (!odd) {
      an = k - a;
      b = 1.0;
    } else {
      an = k;
      b = x;
      k += 1.0;
    }
    odd = !odd;
  }
  return kNaN;
}

// Lower tail by continued fraction for small a and small x.
double continued_fraction_lower(double a, double x) noexcept {
  const double pre = dpois_raw(x, a);
  double f = 1e-30;
  double c = 1e-30;
  double d = 0.0;
  double k = 0.0;
  double an = a;
  double b = a;
  const double nx = 0.0 - x;
  for (int it = 0; it < kMaxTerms; ++it) {
    d = b + mul(an, d);
    if (d == 0.0) {
      d = 1e-30;
    }
    d = div(1.0, d);
    c = b + div(an, c);
    if (c == 0.0) {
      c = 1e-30;
    }
    const double delta = mul(c, d);
    f = mul(f, delta);
    if (std::fabs(delta - 1.0) < kCfTol) {
      return mul(f, pre);
    }
    an = mul(nx, a + k);
    k += 1.0;
    b = k + (a + x);
  }
  return kNaN;
}

// Upper tail for a < 1: alternating series for the lower integral, then
// Q = 1 - P assembled from expm1(a log x) and 1/Gamma(1 + a) - 1.
double small_shape_upper(double a, double x) noexcept {
  const double nx = mul(x, -1.0);
  double term = mul(a, x);
  double s = div(term, a + 1.0);
  double k = 2.0;
  for (int it = 0;; ++it) {
    if (it >= kMaxTerms) {
      return kNaN;
    }
    term = mul(div(nx, k), term);
    s = div(term, a + k) + s;
    if (std::fabs(mul(s, kEps)) >= std::fabs(term)) {
      break;
    }
    k += 1.0;
  }
  const double alx = mul(std::log(x), a);
  s = mul(std::exp(alx), s);
  s = neg_nz(excel_expm1(alx)) + s;
  const double g = mul(std::exp(lgamma_excel(a)), a);
  s = div(s, g);
  return s - recip_gamma_m1(a);
}

// Taylor coefficients of the uniform expansion's correction series.
void temme_coefficients(double a, double (&out)[30]) noexcept {
  out[0] = kTemmeSeed[0];
  out[1] = kTemmeSeed[1];
  double c = 29.0;
  for (int i = 0; i < 28; ++i) {
    out[i + 2] = div(mul(out[i], c), a) + kTemmeSeed[i + 2];
    c -= 1.0;
  }
}

// Uniform asymptotic expansion in eta for large a with x near a; returns
// the upper tail when x > a and the lower tail otherwise.
double uniform_expansion(double a, double x) noexcept {
  const double y = div(x - a, a);
  const double lm = log1pmx(y);
  double eta = std::sqrt(mul(lm, -2.0));
  if (x < a) {
    eta = neg_nz(eta);
  }
  const double e18 = neg_nz(mul(a, lm));
  double cf[30];
  temme_coefficients(a, cf);
  double acc = cf[0];
  for (int i = 1; i < 30; ++i) {
    acc = mul(acc, eta);
    acc += cf[i];
  }
  const double r = div(acc, std::exp(stirlerr(a)));
  double d = std::exp(neg_nz(e18));
  d = mul(d, kPhi0);
  d = div(d, std::sqrt(a));
  d = mul(d, r);
  if (!(x > a)) {
    d = neg_nz(d);
  }
  const double ec = mul(0.5, erfc_excel(std::sqrt(e18)));
  return d + ec;
}

// Exact finite form of Q(a, x) for integer a: exp(-x) * sum_{k<a} x^k / k!.
double upper_integer_shape(double a, double x) noexcept {
  const double e = std::exp(mul(x, -1.0));
  const int n = to_int(a);
  double acc = kRecipFactorial[22 - n];
  int idx = 23 - n;
  for (int i = 0; i < n - 1; ++i) {
    acc = mul(acc, x);
    acc += kRecipFactorial[idx++];
  }
  return mul(acc, e);
}

// Exact finite form of Q(a, x) for half-integer a: erfc(sqrt x) plus an
// exp(-x) / sqrt(pi x) polynomial term.
double upper_half_integer_shape(double a, double x) noexcept {
  const double ec = erfc_excel(std::sqrt(x));
  const int k = to_int(a - 0.5) + 1;
  double acc = kHalfIntSeries[22 - k];
  int idx = 23 - k;
  for (int i = 0; i < k - 1; ++i) {
    acc = mul(acc, x);
    acc += kHalfIntSeries[idx++];
  }
  double l = mul(std::log(x), 0.5);
  l += x;
  double e = std::exp(-l);
  e = mul(e, kInvSqrtPi);
  e = mul(e, acc);
  return e + ec;
}

}  // namespace

double log1pmx(double x) noexcept {
  if (x == 0.0) {
    return 0.0;
  }
  if (!(x < 1.0) || !(x > -0.99)) {
    return std::log(x + 1.0) - x;
  }
  const double r = div(x, x + 2.0);
  const double y = mul(r, r);
  double term = div(mul(y, 2.0), 3.0);
  double s = term;
  double k = 3.0;
  for (int it = 0; it < kMaxTerms; ++it) {
    const double kold = k;
    k += 2.0;
    term = mul(div(mul(y, kold), k), term);
    s += term;
    const double acc = mul(s, kEps);
    if (acc > term) {
      return mul(s - x, r);
    }
    if (!(acc > 0.0)) {
      return kNaN;
    }
  }
  return kNaN;
}

double stirlerr(double n) noexcept {
  double shift = 0.0;
  if (!(n > 6.0)) {
    if (n >= 1.0) {
      const int k = to_int(n);
      if (static_cast<double>(k) == n) {
        return kStirlingSmall[k - 1];
      }
    }
    while (n < 6.0) {
      shift += stirlerr_term(n);
      n += 1.0;
    }
  }
  const double u = div(1.0, n);
  const double p = mul(horner(kStirlingPoly, mul(u, u)), u);
  return p + shift;
}

double bd0(double x, double np) noexcept {
  const double d = x - np;
  const double v = div(d, x + np);
  if (!(v > 0.2) && !(v < -0.2)) {
    const double t = mul(x, 2.0);
    double a = mul(excel_math_detail::atanh_series(v), t);
    a = neg_nz(mul(t, v)) + a;
    return mul(v, d) + a;
  }
  return mul(std::log(div(x, np)), x) - d;
}

double lgamma_excel(double x) noexcept {
  if (x == 2.0 || x == 1.0) {
    return 0.0;
  }
  if (x > 2.0) {
    if (x >= 8.0) {
      double t = mul(x - 0.5, std::log(x));
      t = (t - x) + kLnSqrt2Pi;
      return t + stirlerr(x);
    }
    if (x > 3.0) {
      const int n = to_int(x - 2.0);
      const double y = x - static_cast<double>(n);
      const double lp = std::log(shifted_product(x, n));
      const double p = div(horner(kLgNum23, y), horner(kLgDen23, y));
      return mul(p, y - 2.0) + lp;
    }
    return mul(div(horner(kLgNum23, x), horner(kLgDen23, x)), x - 2.0);
  }
  const std::int64_t bits = bits_of(x);
  if (bits < 0x3cb0000000000000LL) {
    return x > 0.0 ? -std::log(x) : kNaN;
  }
  double add = 0.0;
  double y = x;
  if (bits <= 0x3fe6666666666665LL) {
    add = -std::log(x);
    y = x + 1.0;
  }
  double p;
  if (y < 1.5) {
    p = mul(div(horner(kLgNum07, y), horner(kLgDen07, y)), 1.0 - y);
  } else {
    p = div(horner(kLgNum15, y), horner(kLgDen15, y));
    double t = mul(p, y);
    t = mul(y - 2.0, t);
    p = mul(y - 1.0, t);
  }
  return p + add;
}

double dpois_raw(double lambda, double k) noexcept {
  if (lambda > 1.1 && k > 1.1) {
    const double e = std::exp(neg_nz(stirlerr(k) + bd0(k, lambda)));
    return div(e, std::sqrt(mul(k, kTwoPi)));
  }
  const double nlg = neg_nz(lgamma_excel(k + 1.0));
  const double t1 = mul(std::log(lambda), k);
  const double t2 = neg_nz(lambda);
  double s1;
  double e1;
  double s2;
  double e2;
  two_sum(t1, t2, &s1, &e1);
  two_sum(s1, nlg, &s2, &e2);
  const double r = mul(std::exp(s2), std::exp(e1));
  return mul(r, std::exp(e2));
}

double dbinom_raw(double n, double p, double k) noexcept {
  if (k > n) {
    return 0.0;
  }
  k = std::trunc(k);
  n = std::trunc(n);
  if (n == 0.0) {
    return 1.0;
  }
  if (k == 0.0) {
    return std::exp(mul(n, excel_log1p(neg_nz(p))));
  }
  if (k == n) {
    return std::exp(mul(std::log(p), n));
  }
  const double nk = n - k;
  double lc = stirlerr(n) - stirlerr(k);
  lc -= bd0(k, mul(n, p));
  lc -= stirlerr(nk);
  lc -= bd0(nk, mul(n, 1.0 - p));
  const double t1 = mul(std::log(n), 0.5);
  const double t2 = neg_nz(mul(std::log(k), 0.5));
  const double t3 = neg_nz(mul(std::log(nk), 0.5));
  double sum;
  double e;
  double err;
  two_sum(lc, t1, &sum, &e);
  err = e;
  two_sum(sum, t2, &sum, &e);
  err += e;
  two_sum(sum, t3, &sum, &e);
  err += e;
  two_sum(sum, -kLnSqrt2Pi, &sum, &e);
  err += e;
  return mul(std::exp(sum), std::exp(err));
}

double gamma_ratio_excel(double a, double x, bool upper) noexcept {
  if (a == 0.0) {
    return upper ? 0.0 : 1.0;
  }
  if (x == 0.0) {
    return upper ? 1.0 : 0.0;
  }
  if (a == 1.0) {
    const double t = 0.0 - x;
    return upper ? std::exp(t) : 0.0 - excel_expm1(t);
  }
  if (a == 0.5) {
    const double s = std::sqrt(x);
    return upper ? erfc_excel(s) : erf_excel(s);
  }
  double r;
  bool complement;
  if (a <= 20.0) {
    if (a < 1.0) {
      if (x < 1.1) {
        const double t = x > 0.5 ? x + 0.25 : div(kNegLn2, std::log(x));
        if (t < a) {
          r = continued_fraction_lower(a, x);
          complement = false;
        } else {
          r = small_shape_upper(a, x);
          complement = true;
        }
      } else {
        r = continued_fraction_upper(a, x);
        complement = true;
      }
    } else if (!(x > a)) {
      r = series_lower(a, x);
      complement = false;
    } else {
      const int n = to_int(mul(a, 2.0));
      if (!(x > 30.0) && static_cast<double>(n) == mul(a, 2.0)) {
        r = (n & 1) != 0 ? upper_half_integer_shape(a, x) : upper_integer_shape(a, x);
      } else {
        r = continued_fraction_upper(a, x);
      }
      complement = true;
    }
  } else if (x < 1.1) {
    r = continued_fraction_lower(a, x);
    complement = false;
  } else {
    const bool right = x > a;
    if (std::fabs(1.0 - div(x, a)) > 0.3) {
      if (right) {
        r = asymptotic_upper(a, x);
        complement = true;
      } else {
        r = series_lower(a, x);
        complement = false;
      }
    } else {
      r = uniform_expansion(a, x);
      complement = right;
    }
  }
  return complement != upper ? (0.5 - r) + 0.5 : r;
}

}  // namespace stats
}  // namespace eval
}  // namespace formulon
