//
// Implementation of the special functions declared in special_functions.h:
// the thread-safe log-gamma and the ERF / ERFC / normal-quantile kernels.
//

#include "eval/stats/special_functions.h"

// Apple's <math.h> (and traditionally other libcs) only declares the
// reentrant lgamma_r/lgammaf_r/lgammal_r family when _REENTRANT is defined
// before the header is first included; without it the plain (non-reentrant,
// signgam-writing) lgamma is all that's visible.
#ifndef _REENTRANT
#define _REENTRANT
#endif

#include <algorithm>
#include <cmath>
#include <limits>

#include "eval/stats/horner.h"
#include "utils/subnormal_flush.h"

namespace formulon {
namespace eval {
namespace stats {

double log_gamma(double x) noexcept {
#if defined(_WIN32)
  // The Microsoft CRT does not expose `signgam` at all, so its `lgamma`
  // has no global side effect and is already thread-safe.
  return std::lgamma(x);
#else
  int sign = 0;
  return ::lgamma_r(x, &sign);
#endif
}

namespace {

// Coefficient tables of the ERF/ERFC rational forms, highest degree first.
constexpr double kErfNumSmall[] = {
    0x1.932767df32911p-13, 0x1.d93a5813e2e53p-10, 0x1.6f289d165b309p-5, 0x1.201d8627a7d2cp-3, 0x1.1df5e70c7956bp+0,
};
constexpr double kErfDenSmall[] = {
    0x1.4f20a47119845p-16, 0x1.609751b5c02acp-11, 0x1.56de5da86ad9cp-7,
    0x1.79f85ba688040p-4,  0x1.d191aca8d2499p-2,  0x1.fada15778ece6p-1,
};
constexpr double kErfcNumMid[] = {
    0x1.decdbb07d9e2fp-16, 0x1.41cfbb83ed0e4p-11, 0x1.a016066459c2dp-8, 0x1.50b66746537eap-5, 0x1.76480f4d97fdap-3,
    0x1.27fc8b2d97507p-1,  0x1.4db9fcfb3c6e7p+0,  0x1.044811191d2cbp+1, 0x1.0223d9f70ed04p+1, 0x1.fffffffffff9fp-1,
};
constexpr double kErfcDenMid[] = {
    0x1.a8542411a723bp-15,
    0x1.1d32ae54a52b7p-10,
    0x1.71934f4e7b285p-7,
    0x1.2ca1bb9df5335p-4,
    0x1.516f295417ca8p-2,
    0x1.0f7f27e5dccd7p+0,
    0x1.3bc730d3c1c66p+1,
    0x1.05350002d8ffap+2,
    0x1.25449c578e4bdp+2,
    0x1.92929479238dbp+1,
    0x1p+0,
};
constexpr double kErfcNumBig[] = {
    0x1.4dc756d761350p-9, 0x1.6ca0f8ecfc798p-6, 0x1.0191f83634b97p-3, 0x1.d936f5138efd9p-2, 0x1.396f0fc8e5c9cp+0,
    0x1.26c409d384171p+1, 0x1.80fb2b5716d1ep+1, 0x1.3da266cd039d3p+1, 0x1.000aebf22de3bp+0,
};
constexpr double kErfcDenBig[] = {
    0x1.27cdbb941913dp-8, 0x1.4324d3ef0f238p-5, 0x1.cd27390103553p-3, 0x1.ad794d2e2c325p-1, 0x1.23c050f5fc4a3p+1,
    0x1.1e2d955b8fd9dp+2, 0x1.93da0f9daba61p+2, 0x1.850e3f0a8dba3p+2, 0x1.ce328ffeb796dp+1, 0x1p+0,
};

// Coefficient tables of the normal quantile: central ERFINV rationals, then
// the two tail rationals in sqrt(-ln q); highest degree first.
constexpr double kInvCentralNum[] = {
    -0x1.c6105b261ae5dp-12, 0x1.c034f2417c47bp-6, -0x1.5444e88a51ceep-2, 0x1.a0cbd20e4ce55p+0,
    -0x1.fce6e6f12a56bp+1,  0x1.492fc431b904dp+2, -0x1.b03599ad6a33ap+1, 0x1.c5bf891b4ef6bp-1,
};
constexpr double kInvCentralDen[] = {
    -0x1.5ae11dcadf003p-9, 0x1.38664c121c0dbp-4, -0x1.570c88781d342p-1, 0x1.59bb87268973cp+1,
    -0x1.708d997e809b9p+2, 0x1.ae761f165bea6p+2, -0x1.049a6a9451cabp+2, 0x1.0000000000000p+0,
};
constexpr double kInvMidNum[] = {
    0x1.08ad0740a6d4ap-2,  -0x1.32e5a7e827056p+1, -0x1.86cf9d01e16b2p+1, 0x1.ef9ffb45a3761p+4,
    -0x1.3aefb9a3ec59fp+5, 0x1.5a1d453954a8dp+2,  0x1.ff5f60cbe8bbap+2,  0x1.199fb95ef4e09p-2,
};
constexpr double kInvMidDen[] = {
    -0x1.ad7d1c36679b0p+3, 0x1.7ea3ac1d395ffp+5, -0x1.59120f7e33292p+5,
    -0x1.a4548a23b3e33p+2, 0x1.ccaeab51a6b87p+3, 0x1.0000000000000p+0,
};
constexpr double kInvOuterNum[] = {
    -0x1.1b8ef95ab8b0ep+11, 0x1.da1d6974302b0p+11, 0x1.ed8c72b7a0424p+12,  -0x1.ccdd3c3131b29p+12,
    -0x1.2c8f855be2320p+11, 0x1.f255c31ac92e5p+11, -0x1.7b1a5e1d9a10cp+10, 0x1.0a97c7664c331p+8,
    -0x1.6abfe9390c9e6p+4,  0x1.8242556d368b6p-1,
};
constexpr double kInvOuterDen[] = {
    -0x1.22d350f181b5ep+11, 0x1.9f372995afd52p+9,  0x1.75b4bf097172ep+13,  -0x1.5df83d4e27971p+12,
    -0x1.bb9141a0cb063p+12, 0x1.a1d6dc9f4d349p+12, -0x1.1a2a27132af74p+11, 0x1.76d842df4b9aap+8,
    -0x1.ec8f636f70044p+4,  0x1.0000000000000p+0,
};
constexpr double kTailNumNear[] = {
    0x1.9615ac0b7ace9p-11, 0x1.744eb6c45ec67p-6, 0x1.ef2abb9b85c37p-3, 0x1.453cc085375b2p+0,
    0x1.d2ecb1a3d02c4p+1,  0x1.713f71462256ap+2, 0x1.2857748cab19bp+2, 0x1.6c665fde9526ap+0,
};
constexpr double kTailDenNear[] = {
    0x1.20d3f686439e4p-30, 0x1.1f18cbfdf2728p-11, 0x1.f207a7eab17bfp-7, 0x1.2f5123394f040p-3,
    0x1.61292f23385c9p-1,  0x1.ad278e6526633p+0,  0x1.06cefbb46a449p+1, 0x1.0000000000000p+0,
};
constexpr double kTailNumFar[] = {
    0x1.afb74d693bf93p-23, 0x1.c6ec6cc59e02ap-16, 0x1.45c1908425345p-10, 0x1.b2b41193b4ee7p-6,
    0x1.2fad9315255cfp-2,  0x1.c8ea6461fa445p+0,  0x1.5daea6e875003p+2,  0x1.aa1b1c13ee526p+2,
};
constexpr double kTailDenFar[] = {
    0x1.269bff1f8c190p-49, 0x1.31446f740b9e0p-23, 0x1.35c2c496374bfp-16, 0x1.9c8bc979dc5d7p-11,
    0x1.e76f93215462ap-7,  0x1.186eb183443fbp-3,  0x1.331d34fc7d77fp-1,  0x1.0000000000000p+0,
};

double erfc_positive(double v) noexcept;

// Odd rational in v^2 below 0.5, `1 - erfc` beyond; `v` is non-negative.
double erf_positive(double v) noexcept {
  if (v < 0.5) {
    const double t = flush_subnormal(v * v);
    return flush_subnormal(flush_subnormal(horner(kErfNumSmall, t) * v) / horner(kErfDenSmall, t));
  }
  return 1.0 - erfc_positive(v);
}

// Rational in v times exp(-v^2), split at 4; `v` is non-negative.
double erfc_positive(double v) noexcept {
  if (v < 0.5) {
    return 1.0 - erf_positive(v);
  }
  if (v > 100.0) {
    return 0.0;
  }
  const double e = std::exp(-flush_subnormal(v * v));
  const double r = v > 4.0 ? flush_subnormal(horner(kErfcNumBig, v) / horner(kErfcDenBig, v))
                           : flush_subnormal(horner(kErfcNumMid, v) / horner(kErfcDenMid, v));
  return flush_subnormal(r * e);
}

// Tail rational in w = sqrt(-ln q) with q the smaller side; `a` is the lower-tail
// probability and `b` its complement.
double tail_quantile(double a, double b) noexcept {
  const bool upper = a > 0.5;
  const double w = std::sqrt(-std::log(upper ? b : a));
  double s = w - 5.0;
  double r;
  if (s > 0.0) {
    r = flush_subnormal(horner(kTailNumFar, s) / horner(kTailDenFar, s));
  } else {
    s = w - 1.6;
    r = flush_subnormal(horner(kTailNumNear, s) / horner(kTailDenNear, s));
  }
  return upper ? r : -r;
}

// Inverse error function in three regions of |y|, tail form beyond 0.9375.
double erfinv_core(double y) noexcept {
  if (y == 0.0) {
    return 0.0;
  }
  const double v = std::fabs(y);
  const double t = flush_subnormal(v * v);
  double r;
  if (t <= 0.5) {
    r = flush_subnormal(flush_subnormal(horner(kInvCentralNum, t) / horner(kInvCentralDen, t)) * v);
  } else if (v > 0.9375) {
    const double a = flush_subnormal((1.0 + v) * 0.5);
    const double b = flush_subnormal((1.0 - v) * 0.5);
    r = flush_subnormal(tail_quantile(a, b) * 0.7071067811865476);
  } else {
    const double w = std::sqrt(-std::log(1.0 - v));
    if (v > 0.85) {
      const double u = v - 0.87890625;
      r = flush_subnormal(horner(kInvOuterNum, u) / horner(kInvOuterDen, u));
    } else {
      r = flush_subnormal(horner(kInvMidNum, t) / horner(kInvMidDen, t));
    }
    r = flush_subnormal(r * w);
  }
  return y < 0.0 ? -r : r;
}

}  // namespace

double erf_excel(double x) noexcept {
  if (x == 0.0) {
    return 0.0;
  }
  const double r = erf_positive(std::fabs(x));
  return x < 0.0 ? -r : r;
}

double erfc_excel(double x) noexcept {
  if (x == 0.0) {
    return 1.0;
  }
  const double r = erfc_positive(std::fabs(x));
  return x < 0.0 ? (-r) + 2.0 : r;
}

double normal_cdf(double z) noexcept {
  return flush_subnormal(erfc_excel(flush_subnormal(z * -0.7071067811865476)) * 0.5);
}

double erfinv_excel(double y) noexcept {
  return erfinv_core(y);
}

double normal_quantile(double p) noexcept {
  if (p < 0.075 || p > 0.9375) {
    return tail_quantile(p, 1.0 - p);
  }
  const double y = flush_subnormal(p * 2.0) - 1.0;
  return flush_subnormal(erfinv_core(y) * 1.4142135623730951);
}

}  // namespace stats
}  // namespace eval
}  // namespace formulon
