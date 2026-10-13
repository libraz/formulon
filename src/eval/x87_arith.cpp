#include "eval/x87_arith.h"

#include <cfloat>
#include <cmath>
#include <cstdint>
#include <cstring>
#include <limits>

namespace formulon::eval {

namespace {

// ---------------------------------------------------------------------------
// Double rounding of a correctly rounded double result.
// ---------------------------------------------------------------------------

struct TieWindow {
  double neighbour;  // the other double next to the exact value
  double lo;         // smallest |error| that still rounds to the midpoint on a 64-bit grid
  double hi;         // midpoint distance; |error| must be below it
};

// Describes where the 64-bit rounding of `s + e` lands on the midpoint between
// `s` and its neighbour on the side of `e`. False when no adjustment applies.
bool tie_window(double s, bool toward_positive, TieWindow* out) {
  if (!std::isfinite(s) || s == 0.0 || std::fabs(s) < DBL_MIN) {
    return false;
  }
  constexpr double kInf = std::numeric_limits<double>::infinity();
  const double n = std::nextafter(s, toward_positive ? kInf : -kInf);
  if (std::isinf(n)) {
    return false;
  }
  const double gap = std::fabs(n - s);
  out->neighbour = n;
  out->hi = gap * 0.5;
  out->lo = out->hi - gap * 0x1p-12;
  return true;
}

bool is_even_significand(double v) {
  std::uint64_t bits = 0;
  std::memcpy(&bits, &v, sizeof bits);
  return (bits & 1U) == 0U;
}

double pick_even(double s, double n) {
  return is_even_significand(s) ? s : n;
}

// `s` is the correctly rounded result and `e` the exact residual (s + e is the exact value).
double adjust_with_residual(double s, double e) {
  if (e == 0.0) {
    return s;
  }
  TieWindow w{};
  if (!tie_window(s, e > 0.0, &w)) {
    return s;
  }
  const double ae = std::fabs(e);
  return (ae >= w.lo && ae < w.hi) ? pick_even(s, w.neighbour) : s;
}

// ---------------------------------------------------------------------------
// Double-double helpers (value = hi + lo, |lo| <= ulp(hi) / 2).
// ---------------------------------------------------------------------------

struct DD {
  double hi;
  double lo;
};

DD two_sum(double a, double b) {
  const double s = a + b;
  const double bb = s - a;
  return {s, (a - (s - bb)) + (b - bb)};
}

DD quick_two_sum(double a, double b) {
  const double s = a + b;
  return {s, b - (s - a)};
}

DD two_prod(double a, double b) {
  const double p = a * b;
  return {p, std::fma(a, b, -p)};
}

DD dd_add(DD a, DD b) {
  DD s = two_sum(a.hi, b.hi);
  const DD t = two_sum(a.lo, b.lo);
  s.lo += t.hi;
  s = quick_two_sum(s.hi, s.lo);
  s.lo += t.lo;
  return quick_two_sum(s.hi, s.lo);
}

DD dd_neg(DD a) {
  return {-a.hi, -a.lo};
}

DD dd_mul(DD a, DD b) {
  DD p = two_prod(a.hi, b.hi);
  p.lo += a.hi * b.lo + a.lo * b.hi;
  return quick_two_sum(p.hi, p.lo);
}

DD dd_div(DD a, DD b) {
  const double q1 = a.hi / b.hi;
  DD r = dd_add(a, dd_neg(dd_mul(b, {q1, 0.0})));
  const double q2 = r.hi / b.hi;
  r = dd_add(r, dd_neg(dd_mul(b, {q2, 0.0})));
  const double q3 = r.hi / b.hi;
  return dd_add(quick_two_sum(q1, q2), {q3, 0.0});
}

// Rounds the exact value hi + lo to a 64-bit significand (nearest, ties to even).
DD round_to_64_bits(DD v) {
  if (v.hi == 0.0 || !std::isfinite(v.hi)) {
    return v;
  }
  int e = std::ilogb(v.hi);
  // A power of two with a lower-signed tail belongs to the binade below.
  if (v.lo != 0.0 && (v.lo < 0.0) != (v.hi < 0.0) && std::fabs(v.hi) == std::ldexp(1.0, e)) {
    --e;
  }
  const int shift = 63 - e;
  const double h = std::ldexp(v.hi, shift);
  const double l = std::nearbyint(std::ldexp(v.lo, shift));
  const DD r = two_sum(h, l);
  return {std::ldexp(r.hi, -shift), std::ldexp(r.lo, -shift)};
}

double dd_to_double(DD v) {
  return v.hi + v.lo;
}

// log2(e) rounded to a 64-bit significand.
DD log2e_64() {
  constexpr std::uint64_t kSignificand = 0xB8AA3B295C17F0BCULL;
  const double hi = static_cast<double>(kSignificand);
  const auto lo = static_cast<std::int64_t>(kSignificand - static_cast<std::uint64_t>(hi));
  return {std::ldexp(hi, -63), std::ldexp(static_cast<double>(lo), -63)};
}

constexpr DD kLn2 = {6.931471805599452862e-01, 2.319046813846299558e-17};

// exp(z) - 1 for |z| <= 0.36 by Taylor series, accurate to about 1e-31.
DD dd_expm1_small(DD z) {
  constexpr int kTerms = 26;
  DD p = {1.0, 0.0};
  for (int k = kTerms; k >= 2; --k) {
    p = dd_add({1.0, 0.0}, dd_div(dd_mul(p, z), {static_cast<double>(k), 0.0}));
  }
  return dd_mul(z, p);
}

}  // namespace

double x87_add(double a, double b) {
  const DD t = two_sum(a, b);
  if (!std::isfinite(t.hi)) {
    return t.hi;
  }
  return adjust_with_residual(t.hi, t.lo);
}

double x87_sub(double a, double b) {
  return x87_add(a, -b);
}

double x87_mul(double a, double b) {
  const double s = a * b;
  if (!std::isfinite(s)) {
    return s;
  }
  return adjust_with_residual(s, std::fma(a, b, -s));
}

double x87_div(double a, double b) {
  const double q = a / b;
  TieWindow w{};
  if (!std::isfinite(a) || !std::isfinite(b)) {
    return q;
  }
  const double r = std::fma(-q, b, a);
  if (r == 0.0) {
    return q;
  }
  // The residual of the quotient is r / b; compare it with the window without dividing.
  const bool toward_positive = (r > 0.0) == (b > 0.0);
  if (!tie_window(q, toward_positive, &w)) {
    return q;
  }
  const double ar = std::fabs(r);
  const double ab = std::fabs(b);
  const bool at_least_lo = std::fma(-w.lo, ab, ar) >= 0.0;
  const bool below_hi = std::fma(-w.hi, ab, ar) < 0.0;
  return (at_least_lo && below_hi) ? pick_even(q, w.neighbour) : q;
}

double x87_exp(double x) {
  if (!(std::fabs(x) <= 709.0)) {
    return std::exp(x);
  }
  const DD t = round_to_64_bits(dd_mul(log2e_64(), {x, 0.0}));
  double n = std::nearbyint(t.hi);
  DD f = dd_add(t, {-n, 0.0});
  if (f.hi > 0.5 || (f.hi == 0.5 && f.lo > 0.0)) {
    n += 1.0;
    f = dd_add(t, {-n, 0.0});
  } else if (f.hi < -0.5 || (f.hi == -0.5 && f.lo < 0.0)) {
    n -= 1.0;
    f = dd_add(t, {-n, 0.0});
  }
  DD y = round_to_64_bits(dd_expm1_small(dd_mul(f, kLn2)));
  y = round_to_64_bits(dd_add(y, {1.0, 0.0}));
  // Scaling a double into the subnormal range rounds a second time; only that tail differs.
  return std::ldexp(dd_to_double(y), static_cast<int>(n));
}

double x87_ln(double x) {
  if (!(x > 0.0) || !std::isfinite(x)) {
    return std::log(x);
  }
  int k = 0;
  double m = std::frexp(x, &k);
  if (m < 0.7071067811865476) {
    m *= 2.0;
    --k;
  }
  const DD s = dd_div({m - 1.0, 0.0}, two_sum(m, 1.0));
  const DD s2 = dd_mul(s, s);
  constexpr int kTerms = 24;
  DD sum = dd_div({1.0, 0.0}, {2.0 * kTerms + 1.0, 0.0});
  for (int j = kTerms - 1; j >= 0; --j) {
    sum = dd_add(dd_mul(sum, s2), dd_div({1.0, 0.0}, {2.0 * j + 1.0, 0.0}));
  }
  const DD ln_m = dd_mul(dd_mul(s, sum), {2.0, 0.0});
  const DD total = dd_add(ln_m, dd_mul(kLn2, {static_cast<double>(k), 0.0}));
  return dd_to_double(round_to_64_bits(total));
}

}  // namespace formulon::eval
