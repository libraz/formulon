//
// See hypothesis_tests.h. Every arithmetic step is a separate double
// operation: multiplies and divides flush a subnormal result to zero, and
// nothing is contracted into a fused multiply-add.

#include "eval/stats/hypothesis_tests.h"

#include <cmath>

#include "eval/stats/flushed_ops.h"
#include "eval/stats/incomplete_beta.h"
#include "eval/stats/incomplete_gamma.h"
#include "eval/stats/inverse_distributions.h"
#include "eval/stats/special_functions.h"

namespace formulon {
namespace eval {
namespace stats {
namespace {

using flushed::div;
using flushed::mul;
using flushed::neg;

constexpr TestOutcome kDiv0 = {0.0, true};

// Two-sided p of Student's t with `df` degrees of freedom; `tails == 1` halves it.
TestOutcome t_probability(double t, double df, int tails) noexcept {
  const double a = div(df, 2.0);
  const double den = mul(t, t) + df;
  const double p2 = beta_ratio_excel(a, 0.5, div(df, den), false);
  return {tails == 1 ? div(p2, 2.0) : p2, false};
}

}  // namespace

TestOutcome t_test_paired_excel(const double* x, const double* y, std::size_t n, int tails) noexcept {
  if (n == 0) {
    return kDiv0;
  }
  double s1 = 0.0;
  double s2 = 0.0;
  for (std::size_t i = 0; i < n; ++i) {
    s1 = x[i] + s1;
    s2 = y[i] + s2;
  }
  const double dn = static_cast<double>(n);
  const double m1 = div(s1, dn);
  const double m2 = div(s2, dn);
  double sxy = 0.0;
  double sxx = 0.0;
  double syy = 0.0;
  for (std::size_t i = 0; i < n; ++i) {
    const double dx = x[i] - m1;
    const double dy = y[i] - m2;
    sxy = mul(dy, dx) + sxy;
    sxx = mul(dx, dx) + sxx;
    syy = mul(dy, dy) + syy;
  }
  const double dn1 = dn - 1.0;
  if (dn1 == 0.0) {
    return kDiv0;
  }
  const double var1 = div(sxx, dn1);
  const double var2 = div(syy, dn1);
  const double cov = div(sxy, dn1);
  double s = neg(mul(2.0, cov)) + var1;
  s += var2;
  s = div(s, dn);
  if (s == 0.0) {
    return kDiv0;
  }
  const double t = div(m1 - m2, std::sqrt(s));
  return t_probability(t, dn1, tails);
}

TestOutcome t_test_two_sample_excel(const double* x, std::size_t nx, const double* y, std::size_t ny, int tails,
                                    int type) noexcept {
  if (nx == 0 || ny == 0) {
    return kDiv0;
  }
  double s1 = 0.0;
  double s2 = 0.0;
  for (std::size_t i = 0; i < nx; ++i) {
    s1 = x[i] + s1;
  }
  for (std::size_t i = 0; i < ny; ++i) {
    s2 = y[i] + s2;
  }
  const double dnx = static_cast<double>(nx);
  const double dny = static_cast<double>(ny);
  const double m1 = div(s1, dnx);
  const double m2 = div(s2, dny);
  double sxx = 0.0;
  double syy = 0.0;
  for (std::size_t i = 0; i < nx; ++i) {
    const double dx = x[i] - m1;
    sxx = mul(dx, dx) + sxx;
  }
  for (std::size_t i = 0; i < ny; ++i) {
    const double dy = y[i] - m2;
    syy = mul(dy, dy) + syy;
  }
  const double md = m1 - m2;
  const double dn1 = dnx - 1.0;
  const double dn2 = dny - 1.0;
  if (dn1 == 0.0 || dn2 == 0.0) {
    return kDiv0;
  }
  const double var1 = div(sxx, dn1);
  const double var2 = div(syy, dn2);
  double t;
  double df;
  if (type == 2) {
    df = dn1 + dn2;
    const double ssx = mul(var1, dn1);
    const double pooled = div(mul(var2, dn2) + ssx, df);
    if (pooled == 0.0) {
      return kDiv0;
    }
    const double s = mul(div(1.0, dny) + div(1.0, dnx), pooled);
    t = div(md, std::sqrt(s));
  } else {
    const double a1 = div(var1, dnx);
    const double a2 = div(var2, dny);
    const double s = a2 + a1;
    if (s == 0.0) {
      return kDiv0;
    }
    t = div(md, std::sqrt(s));
    const double d = div(mul(a1, a1), dn1) + div(mul(a2, a2), dn2);
    if (d == 0.0) {
      return kDiv0;
    }
    df = div(mul(s, s), d);
  }
  return t_probability(t, df, tails);
}

TestOutcome f_test_excel(const double* x, std::size_t nx, const double* y, std::size_t ny) noexcept {
  if (nx < 2 || ny < 2) {
    return kDiv0;
  }
  double s1 = 0.0;
  double s2 = 0.0;
  for (std::size_t i = 0; i < nx; ++i) {
    s1 = x[i] + s1;
  }
  for (std::size_t i = 0; i < ny; ++i) {
    s2 = y[i] + s2;
  }
  const double dnx = static_cast<double>(nx);
  const double dny = static_cast<double>(ny);
  const double m1 = div(s1, dnx);
  const double m2 = div(s2, dny);
  double sxx = 0.0;
  double syy = 0.0;
  for (std::size_t i = 0; i < nx; ++i) {
    const double dx = x[i] - m1;
    sxx = mul(dx, dx) + sxx;
  }
  for (std::size_t i = 0; i < ny; ++i) {
    const double dy = y[i] - m2;
    syy = mul(dy, dy) + syy;
  }
  const double dn1 = dnx - 1.0;
  const double dn2 = dny - 1.0;
  const double var1 = div(sxx, dn1);
  const double var2 = div(syy, dn2);
  double f;
  double dnum;
  double dden;
  if (var2 < var1) {
    if (var2 == 0.0) {
      return kDiv0;
    }
    f = div(var1, var2);
    dnum = dn1;
    dden = dn2;
  } else {
    if (var1 == 0.0) {
      return kDiv0;
    }
    f = div(var2, var1);
    dnum = dn2;
    dden = dn1;
  }
  const double t = mul(dnum, f);
  const double s = t + dden;
  const double p = beta_ratio_excel(div(dden, 2.0), div(dnum, 2.0), div(dden, s), false, div(t, s));
  double r = mul(p, 2.0);
  if (r > 1.0) {
    r = neg(r) + 2.0;
  }
  return {r, false};
}

double chisq_accumulate_excel(double acc, double observed, double expected) noexcept {
  const double d = observed - expected;
  return div(mul(d, d), expected) + acc;
}

double chisq_test_p_excel(double stat, double df) noexcept {
  return gamma_ratio_excel(div(df, 2.0), div(stat, 2.0), true);
}

double confidence_norm_excel(double alpha, double sd, double size) noexcept {
  const double z = normal_quantile(neg(div(alpha, 2.0)) + 1.0);
  return mul(div(sd, std::sqrt(std::trunc(size))), z);
}

double confidence_t_excel(double alpha, double sd, double size) noexcept {
  size = std::trunc(size);
  const double t = t_inv_2t_excel(alpha, size - 1.0);
  return mul(div(sd, std::sqrt(size)), t);
}

}  // namespace stats
}  // namespace eval
}  // namespace formulon
