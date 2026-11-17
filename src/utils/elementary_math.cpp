//
// Circular functions for `elementary_math.h`. An argument is reduced to r in [-pi/4, pi/4] plus a quadrant, by one
// of three paths chosen on |x|: none below pi/4, a three-part pi/2 with a compensated subtraction below 2^19, and a
// fixed-point product with the bits of 2/pi above. The kernels are odd and even polynomials in r^2 evaluated in
// Horner form with fused multiply-adds; tangent is their quotient. All arithmetic is IEEE double or 64-bit integer,
// so every host computes the same bits.

#include "utils/elementary_math.h"

#include <cmath>
#include <cstdint>
#include <cstring>

namespace formulon {
namespace elementary {
namespace {

constexpr std::uint64_t kSignBit = std::uint64_t{1} << 63;

double from_bits(std::uint64_t u) noexcept {
  double d;
  std::memcpy(&d, &u, sizeof(d));
  return d;
}

std::uint64_t to_bits(double d) noexcept {
  std::uint64_t u;
  std::memcpy(&u, &d, sizeof(u));
  return u;
}

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

}  // namespace

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
  if (ax < from_bits(0x3fe921fb54442d18)) {  // pi/4
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
