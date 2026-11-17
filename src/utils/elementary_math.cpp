//
// Circular functions for `elementary_math.h`. An argument is reduced to r in [-pi/4, pi/4] plus a quadrant, by one
// of three paths chosen on |x|: none below pi/4, a three-part pi/2 with a compensated subtraction below 2^19, and a
// fixed-point product with the bits of 2/pi above. The kernels are odd and even polynomials in r^2 evaluated in
// Horner form with fused multiply-adds; tangent is their quotient. All arithmetic is IEEE double or 64-bit integer,
// so every host computes the same bits.

#include "utils/elementary_math.h"

#include <cmath>
#include <cstdint>

#include "utils/elementary_math_bits.h"

// The bits assume no contraction of a * b + c into one fused operation; explicit std::fma calls are unaffected.
#if defined(__clang__)
#pragma clang fp contract(off)
#endif

namespace formulon {
namespace elementary {
namespace {

constexpr std::uint64_t kSignBit = std::uint64_t{1} << 63;

using detail::from_bits;
using detail::to_bits;

// High 64 bits of the 128-bit product a * b.
std::uint64_t mul_high(std::uint64_t a, std::uint64_t b) noexcept {
  const std::uint64_t a_lo = a & 0xffffffffU;
  const std::uint64_t a_hi = a >> 32;
  const std::uint64_t b_lo = b & 0xffffffffU;
  const std::uint64_t b_hi = b >> 32;
  const std::uint64_t lo_lo = a_lo * b_lo;
  const std::uint64_t hi_lo = a_hi * b_lo;
  const std::uint64_t lo_hi = a_lo * b_hi;
  const std::uint64_t cross = (lo_lo >> 32) + (hi_lo & 0xffffffffU) + lo_hi;
  return a_hi * b_hi + (hi_lo >> 32) + (cross >> 32);
}

// Leading zero bits of v; 64 for zero.
unsigned leading_zeros(std::uint64_t v) noexcept {
  if (v == 0) {
    return 64;
  }
  unsigned n = 0;
  for (unsigned step = 32; step != 0; step >>= 1) {
    if ((v >> (64 - step)) == 0) {
      n += step;
      v <<= step;
    }
  }
  return n;
}

// Bits of 2/pi, most significant word last; a reduction reads four consecutive words chosen by the exponent.
constexpr std::uint64_t kTwoOverPiBits[19] = {0x35fdafd88fc6ae84,
                                              0x9e839cfbc5294975,
                                              0xba93dd63f5f2f8bd,
                                              0xa7a31fb34f2ff516,
                                              0xb69b3f6793e584db,
                                              0xf79788c5ad05368f,
                                              0x8ffc4bffef02cc07,
                                              0x4e422fc5defc941d,
                                              0x9cc8eb1cc1a99cfa,
                                              0x74ce38135a2fbf20,
                                              0x74411afa975da242,
                                              0x7f0ef58e5894d39f,
                                              0x0324977504e8c90e,
                                              0xdb92371d2126e970,
                                              0xff28b1d5ef5de2b0,
                                              0x6db14acc9e21c820,
                                              0xfe13abe8fa9a6ee0,
                                              0x517cc1b727220a94,
                                              0};

// pi/2 as a 0.64 fixed-point fraction of 2, for scaling a reduced fraction of a quadrant.
constexpr std::uint64_t kPiOverTwoFixed = 0xc90fdaa22168c235;

// Odd-kernel coefficients: sin(r) = r + r * r^2 * P(r^2), highest degree first.
constexpr std::uint64_t kSinCoefficients[6] = {0x3de5d8fd1fd19ccd, 0xbe5ae5e5a9291f5d, 0x3ec71de3567d48a1,
                                               0xbf2a01a019bfdf03, 0x3f8111111110f7d0, 0xbfc5555555555548};

// Even-kernel coefficients: cos(r) = 1 + r^2 * Q(r^2), highest degree first.
constexpr std::uint64_t kCosCoefficients[7] = {0xbda8fa49a0861a9b, 0x3e21ee9d7b4e3f05, 0xbe927e4f7eac4bc6,
                                               0x3efa01a019c844f5, 0xbf56c16c16c14f91, 0x3fa555555555554b,
                                               0xbfe0000000000000};

// Below this magnitude the argument is reduced with the three-part pi/2; above, with the bits of 2/pi.
constexpr std::uint64_t kMidReductionLimit = 0x412000013be57a40;

double sin_poly(double r2) noexcept {
  double s = std::fma(r2, from_bits(kSinCoefficients[0]), from_bits(kSinCoefficients[1]));
  for (int i = 2; i < 6; ++i) {
    s = std::fma(r2, s, from_bits(kSinCoefficients[i]));
  }
  return s;
}

double cos_poly(double r2) noexcept {
  double c = std::fma(r2, from_bits(kCosCoefficients[0]), from_bits(kCosCoefficients[1]));
  for (int i = 2; i < 7; ++i) {
    c = std::fma(r2, c, from_bits(kCosCoefficients[i]));
  }
  return std::fma(r2, c, 1.0);
}

// x - k * pi/2 with k = rint(x * 2/pi), as head + tail; returns k.
std::int64_t reduce_mid(double x, double* head, double* tail) noexcept {
  const double k = std::nearbyint(x * from_bits(0x3fe45f306dc9c883));  // 2/pi
  const double p1 = k * from_bits(0x3ff921fb54440000);                 // pi/2, leading 33 bits
  const double p2 = k * from_bits(0x3d868c234c4c0000);                 // pi/2, next 33 bits
  const double p3 = k * from_bits(0x3b298a2e03707345);                 // pi/2, remainder
  const double a = x - p1;
  *head = a - p2;
  *tail = ((a - *head) - p2) - p3;
  return static_cast<std::int64_t>(k);
}

// Large-argument reduction of |x|: the product of the significand with the bits of 2/pi. Returns the top product
// word (quadrant bits included); `fraction` is the distance to the nearest quadrant boundary, normalised by `shift`.
struct LargeReduction {
  std::uint64_t top;
  std::uint64_t fraction;
  unsigned shift;
  bool odd;
};

LargeReduction reduce_large(double ax) noexcept {
  const std::uint64_t bits = to_bits(ax);
  const std::int64_t exponent = static_cast<std::int64_t>(bits >> 52) - 0x3ff;
  const std::uint64_t significand = kSignBit | (bits << 11);
  const std::uint64_t* t = kTwoOverPiBits + (15 - (exponent >> 6));
  const unsigned s = static_cast<unsigned>(exponent & 63);
  std::uint64_t w0;
  std::uint64_t w1;
  std::uint64_t w2;
  if (s == 0) {
    w0 = t[1];
    w1 = t[2];
    w2 = t[3];
  } else {
    w2 = (t[3] << s) | (t[2] >> (64 - s));
    w1 = (t[2] << s) | (t[1] >> (64 - s));
    w0 = (t[1] << s) | (t[0] >> (64 - s));
  }
  std::uint64_t lo = mul_high(w0, significand);
  const std::uint64_t mid = w1 * significand;
  std::uint64_t hi = mul_high(w1, significand);
  lo += mid;
  hi += w2 * significand + (lo < mid ? 1U : 0U);
  LargeReduction out{};
  out.top = hi;
  out.odd = (((hi << 2) ^ (hi << 1)) >> 63) != 0;
  const std::uint64_t negate = ((hi << 2) >> 63) != 0 ? ~std::uint64_t{0} : 0;
  const std::uint64_t h = (hi ^ negate) & 0x1fffffffffffffffU;
  lo ^= negate;
  out.shift = leading_zeros(h);
  out.fraction = (h << (out.shift & 63)) | (lo >> ((64 - out.shift) & 63));
  return out;
}

// value / 2^fraction_bits / 2^shift, the shift applied to the exponent field.
double scaled_fixed(std::uint64_t value, unsigned fraction_bits, std::uint64_t shift) noexcept {
  const double v = static_cast<double>(value) * from_bits(static_cast<std::uint64_t>(1023 - fraction_bits) << 52);
  return from_bits(to_bits(v) - (shift << 52));
}

// Odd kernel on head + tail: sin(r) for r = head + tail.
double sin_kernel(double head, double tail, std::uint64_t sign) noexcept {
  const double r = head + tail;
  const double r2 = r * r;
  return from_bits(to_bits(std::fma(r, r2 * sin_poly(r2), tail) + head) ^ sign);
}

// Even kernel on head + tail: cos(r), with r^2 formed as head^2 + 2 head tail.
double cos_kernel(double head, double tail, std::uint64_t sign) noexcept {
  double cross = head * tail;
  cross = cross + cross;
  return from_bits(to_bits(cos_poly(std::fma(head, head, cross))) ^ sign);
}

// Odd kernel on a reduced fraction: the scaled value split into its top 16 bits and the rest.
double sin_kernel_fixed(const LargeReduction& red, std::uint64_t sign) noexcept {
  const std::uint64_t t = mul_high(kPiOverTwoFixed, red.fraction);
  const double r = scaled_fixed(t, 61, red.shift);
  const double hi = scaled_fixed(t & 0xffff000000000000U, 61, red.shift);
  const double lo = scaled_fixed(t & 0x0000ffffffffffffU, 61, red.shift);
  const double r2 = r * r;
  return from_bits(to_bits(std::fma(r, r2 * sin_poly(r2), lo) + hi) ^ sign);
}

// Even kernel on a reduced fraction, with r^2 squared in fixed point.
double cos_kernel_fixed(const LargeReduction& red, std::uint64_t sign) noexcept {
  std::uint64_t t = mul_high(kPiOverTwoFixed, red.fraction);
  t = mul_high(t, t);
  return from_bits(to_bits(cos_poly(scaled_fixed(t, 58, std::uint64_t{2} * red.shift))) ^ sign);
}

// Arctangent: |x| up to 2 in intervals, each a polynomial in |x| - centre.
struct ArctanInterval {
  std::uint64_t upper;             // largest |x| in the interval
  std::uint64_t centre;            // minus the expansion point
  int count;                       // coefficients in use
  std::uint64_t coefficients[13];  // highest degree first, the constant term last
};

constexpr ArctanInterval kArctanIntervals[5] = {
    {0x3fe8000000000000,
     0xbfe4000000000027,
     13,
     {0x3f710f3e20f6a2e2, 0xbf8d88d31abc3ae5, 0x3f8950e69dcdd967, 0x3f7ff3b7c531aa4a, 0xbfa09bc0ab7f914c,
      0x3fa094966be2b531, 0x3f8a759263f377f2, 0xbfb519e110f61b54, 0x3fba1247ca5d9475, 0x3f95d0b7e9e69054,
      0xbfd4af2b78215a1b, 0x3fe702e05c0b8155, 0x3fe1e00babdefed0}},
    {0x3ff0000000000000,
     0xbfec0000000f4213,
     12,
     {0x3f1ca7cde5de9bd7, 0xbf7265ef51b17db7, 0x3f8152311b180e6c, 0xbf7ac1645739e676, 0xbf791f7e2a7a338f,
      0x3f9c90e92ac8d86c, 0xbfa63b543effa4ef, 0x3f959bc93f81895a, 0x3fb41b15e5e8dcd0, 0xbfd1f6a8499714a2,
      0x3fe21fb781196ac3, 0x3fe700a7c580ea7e}},
    {0x3ff5555555555555,
     0xbff2aaaaaaaaaa96,
     13,
     {0xbf3883d8959134b3, 0x3f4a89224ff69018, 0xbf50120e602f6336, 0x3f2cbd49da316282, 0x3f62fcdacdd6e5b5,
      0xbf7c0b993a09ce31, 0x3f873a9328786665, 0xbf82b01fc60cc37a, 0xbf8bc317394714b7, 0x3fb3fd2b9f586a67,
      0xbfcac97826d58470, 0x3fdb1b1b1b1b1b3d, 0x3feb96e5a78c5c40}},
    {0x3ffaaaaaaaaaaaab,
     0xbff8000000000003,
     12,
     {0x3f00280537f097f3, 0x3f1b4e30d1ba3819, 0xbf3e4eef429eb680, 0x3f525bad4d85cbe1, 0xbf5f4ac1342182d2,
      0x3f5dfaa5e77b7375, 0x3f613a254d6e5b7c, 0xbf9135a0938ec462, 0x3fac963c83985742, 0xbfc22d719c06115e,
      0x3fd3b13b13b13b0c, 0x3fef730bd281f69d}},
    {0x4000000000000000,
     0xbffd555555461337,
     11,
     {0x3f100b6a460ac05d, 0xbf23045b70e93129, 0x3f310f884eac0e0a, 0xbf32b9ad13db35a8, 0xbf33182219e21362,
      0x3f68cb82a74e0699, 0xbf8881ec3d15241f, 0x3fa2b090aad5f9dc, 0xbfb8ad3c44f10dc3, 0x3fcd59ae78c11c49,
      0x3ff124a85750fb5c}},
};

// atan(x) = x + x * x^2 * P(x^2) for |x| <= 1/2, highest degree first.
constexpr std::uint64_t kArctanSmall[12] = {0x3f84a9f5c4724056, 0xbf9c0eb85f543412, 0x3fa56cdb5d887934,
                                            0xbfaa38cf590469ec, 0x3fadfe7b9674ae37, 0xbfb10f31279ec05d,
                                            0x3fb3b113f18ac049, 0xbfb745cf51795b21, 0x3fbc71c7096c2725,
                                            0xbfc2492492179ca3, 0x3fc99999999918d8, 0xbfd555555555551b};

// atan(t) = t + t * t^2 * P(t^2) for t = 1/|x| with |x| > 2, highest degree first.
constexpr std::uint64_t kArctanReciprocal[11] = {0xbf883dafda7bd3fd, 0x3fa007733e06ceb3, 0xbfa81d33e401836d,
                                                 0x3fad78252fa69c1c, 0xbfb104146b1a1ae8, 0x3fb3afd7a0e6eb75,
                                                 0xbfb745b7f2d72663, 0x3fbc71c5edfed480, 0xbfc249248e1422e3,
                                                 0x3fc999999989ebca, 0xbfd5555555554a51};

// x + x * x^2 * P(x^2), P in Horner form.
double odd_series(double x, const std::uint64_t* coefficients, int count) noexcept {
  const double x2 = x * x;
  double p = from_bits(coefficients[0]);
  for (int i = 1; i < count; ++i) {
    p = std::fma(p, x2, from_bits(coefficients[i]));
  }
  return std::fma(x2 * p, x, x);
}

double arctan_positive(double ax) noexcept {
  if (ax <= from_bits(0x3e4d12ed0af1a27f)) {  // below this atan(x) rounds to x
    return ax;
  }
  if (ax <= 0.5) {
    return odd_series(ax, kArctanSmall, 12);
  }
  for (const ArctanInterval& iv : kArctanIntervals) {
    if (ax <= from_bits(iv.upper)) {
      const double u = ax + from_bits(iv.centre);
      double p = from_bits(iv.coefficients[0]);
      for (int j = 1; j < iv.count; ++j) {
        p = std::fma(p, u, from_bits(iv.coefficients[j]));
      }
      return p;
    }
  }
  const double half_pi_hi = from_bits(0x3ff921fb54442d18);  // pi/2, leading part
  if (ax >= 0x1p53) {
    return half_pi_hi + from_bits(0x0010000000000000);
  }
  const double half_pi_lo = from_bits(0x3c91a62633145c07);  // pi/2 - half_pi_hi
  return (half_pi_lo - odd_series(1.0 / ax, kArctanReciprocal, 11)) + half_pi_hi;
}

constexpr double kPiOverFour = 0x1.921fb54442d18p-1;

}  // namespace

double sin(double x) noexcept {
  if (x == 0.0) {
    return x;
  }
  if (!std::isfinite(x)) {
    return x - x;
  }
  const double ax = std::fabs(x);
  if (ax <= kPiOverFour) {
    return sin_kernel(0.0, x, 0);
  }
  if (ax < from_bits(kMidReductionLimit)) {
    double head;
    double tail;
    const std::int64_t k = reduce_mid(x, &head, &tail);
    const std::uint64_t sign = (static_cast<std::uint64_t>(k) << 62) & kSignBit;
    return (k & 1) != 0 ? cos_kernel(head, tail, sign) : sin_kernel(head, tail, sign);
  }
  const LargeReduction red = reduce_large(ax);
  const std::uint64_t sign = ((to_bits(x) & kSignBit) ^ red.top) & kSignBit;
  return red.odd ? cos_kernel_fixed(red, sign) : sin_kernel_fixed(red, sign);
}

double cos(double x) noexcept {
  if (!std::isfinite(x)) {
    return x - x;
  }
  const double ax = std::fabs(x);
  if (ax <= kPiOverFour) {
    return cos_kernel(ax, 0.0, 0);
  }
  if (ax < from_bits(kMidReductionLimit)) {
    double head;
    double tail;
    const std::int64_t k = reduce_mid(ax, &head, &tail);
    const std::uint64_t sign = static_cast<std::uint64_t>(k ^ (k >> 1)) << 63;
    return (k & 1) != 0 ? sin_kernel(head, tail, sign) : cos_kernel(head, tail, sign);
  }
  const LargeReduction red = reduce_large(ax);
  const std::uint64_t sign = (red.top ^ (red.top << 1)) & kSignBit;
  return red.odd ? sin_kernel_fixed(red, sign) : cos_kernel_fixed(red, sign);
}

double atan(double x) noexcept {
  if (std::isnan(x)) {
    return x - x;
  }
  if (x == 0.0) {
    return x;
  }
  return x < 0.0 ? -arctan_positive(-x) : arctan_positive(x);
}

double tan(double x) noexcept {
  if (x == 0.0) {
    return x;
  }
  if (!std::isfinite(x)) {
    return x - x;
  }
  const double ax = std::fabs(x);
  double r;
  std::uint64_t sign = 0;
  bool odd = false;
  if (ax < kPiOverFour) {
    r = x;
  } else if (ax < from_bits(kMidReductionLimit)) {
    double head;
    double tail;
    const std::int64_t k = reduce_mid(x, &head, &tail);
    r = head + tail;
    odd = (k & 1) != 0;
    sign = static_cast<std::uint64_t>(k) << 63;
  } else {
    const LargeReduction red = reduce_large(ax);
    r = scaled_fixed(mul_high(kPiOverTwoFixed, red.fraction), 61, red.shift);
    odd = red.odd;
    sign = ((to_bits(x) & kSignBit) ^ (red.top << 1)) & kSignBit;
  }
  const double r2 = r * r;
  const double s = from_bits(to_bits(std::fma(r, r2 * sin_poly(r2), r)) ^ sign);
  const double c = cos_poly(r2);
  return odd ? c / s : s / c;
}

}  // namespace elementary
}  // namespace formulon
