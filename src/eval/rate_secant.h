//
// The alternating secant search behind IRR and RATE. A residual functor `f`
// maps a point to its residual and reports a failed evaluation through
// `failed()`; every comparison follows the solver's own zero and sign tests.

#ifndef FORMULON_EVAL_RATE_SECANT_H_
#define FORMULON_EVAL_RATE_SECANT_H_

#include <cmath>

namespace formulon {
namespace eval {
namespace rate_secant {

/// Magnitudes below this are zero to the solver's tests.
inline constexpr double kTiny = 0x1p-1026;

inline bool is_normal(double x) noexcept {
  return !(std::fabs(x) < kTiny || std::isinf(x) || std::isnan(x));
}

/// Sign test of the solver: NaN, infinities and tiny values count as non-negative.
inline bool is_negative(double x) noexcept {
  return x < 0.0 && !(std::fabs(x) < kTiny) && !std::isnan(x);
}

/// True while `quantity` still exceeds `tolerance` in the solver's sense.
inline bool exceeds(double quantity, double tolerance) noexcept {
  const double excess = std::fabs(quantity) - tolerance;
  return is_normal(excess) && excess > 0.0;
}

struct Result {
  bool converged;
  double x;
  double tolerance;
};

/// Alternating secant steps on `f` from `a`: the second point sits 0.001 away,
/// a zero difference nudges a point by 0.001, and the search stops when both
/// the residual and the step are within 1e-7. After 100 passes a last test at
/// 1e-4. `F` provides `double operator()(double)` and `bool failed() const`.
template <class F>
Result search(F& f, double a) noexcept {
  constexpr Result kFailed = {false, 0.0, 0.0};
  double tol = 1.0e-7;
  double fa = f(a);
  double b = is_negative(fa) ? a - 0.001 : 0.001 + a;
  double fb = f(b);
  if (f.failed()) {
    return kFailed;
  }
  double x = fb;
  for (int pass = 0; pass < 100; ++pass) {
    double d = x - fa;
    if (!is_normal(d)) {
      const double e = a - b;
      a = (!is_normal(e) || is_negative(e)) ? 0.001 + a : a - 0.001;
      fa = f(a);
      d = fb - fa;
      if (f.failed() || !is_normal(d)) {
        return kFailed;
      }
    }
    a = b - ((b - a) * fb) / d;
    fa = f(a);
    if (f.failed()) {
      return kFailed;
    }
    if (!exceeds(fa, tol) && !exceeds(a - b, tol)) {
      return {true, a, tol};
    }
    d = fa - fb;
    if (!is_normal(d)) {
      const double e = b - a;
      b = (!is_normal(e) || is_negative(e)) ? 0.001 + b : b - 0.001;
      fb = f(b);
      d = fa - fb;
      if (f.failed() || !is_normal(d)) {
        return kFailed;
      }
    }
    b = a - ((a - b) * fa) / d;
    fb = f(b);
    if (f.failed()) {
      return kFailed;
    }
    if (!exceeds(fb, tol) && !exceeds(a - b, tol)) {
      return {true, b, tol};
    }
    x = fb;
  }
  tol = 1.0e-4;
  if (exceeds(fb, tol) || exceeds(a - b, tol)) {
    return kFailed;
  }
  return {true, b, tol};
}

}  // namespace rate_secant
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_RATE_SECANT_H_
