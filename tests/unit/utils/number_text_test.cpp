//
// Unit tests for the printf-free number formatters in `utils/number_text.h`.
//
// Every primitive is compared byte for byte with the host `snprintf` over
// hand-picked values, generated magnitudes, exact decimal ties and random
// bit patterns; ties are where double-conversion alone disagrees with printf.

#include "utils/number_text.h"

#include <cfloat>
#include <climits>
#include <cmath>
#include <cstdint>
#include <cstdio>
#include <cstring>
#include <limits>
#include <random>
#include <string>
#include <vector>

#include "gtest/gtest.h"

namespace formulon {
namespace {

constexpr std::size_t kBufSize = 2048;

std::string Fixed(double v, int frac) {
  char buf[kBufSize];
  const int n = format_fixed(buf, sizeof(buf), v, frac);
  EXPECT_EQ(static_cast<std::size_t>(n), std::strlen(buf));
  return buf;
}

std::string Expo(double v, int frac) {
  char buf[kBufSize];
  format_exponential(buf, sizeof(buf), v, frac);
  return buf;
}

std::string General(double v, int prec) {
  char buf[kBufSize];
  format_general(buf, sizeof(buf), v, prec);
  return buf;
}

// Apple's libc leaves a trailing zero in `%g` when tie rounding produces it
// (`%.8g` of 392240505 gives "3.9224050e+08"); musl and glibc strip it, as
// the standard requires, so the reference strips it too.
std::string StripGeneralZeros(std::string text) {
  const std::size_t e = text.find('e');
  const std::size_t mant_end = e == std::string::npos ? text.size() : e;
  if (text.find('.') == std::string::npos || text.find('.') > mant_end) {
    return text;
  }
  std::size_t end = mant_end;
  while (text[end - 1] == '0') {
    --end;
  }
  if (text[end - 1] == '.') {
    --end;
  }
  return text.erase(end, mant_end - end);
}

std::string Ref(const char* fmt, int prec, double v) {
  char buf[kBufSize];
  std::snprintf(buf, sizeof(buf), fmt, prec, v);
  return fmt[3] == 'g' ? StripGeneralZeros(buf) : std::string(buf);
}

void CheckFixed(double v, int frac) {
  ASSERT_EQ(Fixed(v, frac), Ref("%.*f", frac, v)) << "fixed frac=" << frac << " v=" << Ref("%.*e", 20, v);
}

void CheckExpo(double v, int frac) {
  ASSERT_EQ(Expo(v, frac), Ref("%.*e", frac, v)) << "expo frac=" << frac << " v=" << Ref("%.*e", 20, v);
}

void CheckGeneral(double v, int prec) {
  ASSERT_EQ(General(v, prec), Ref("%.*g", prec, v)) << "general prec=" << prec << " v=" << Ref("%.*e", 20, v);
}

void CheckAll(double v, bool with_fixed = true) {
  if (with_fixed) {
    for (int f = 0; f <= 20; ++f) {
      CheckFixed(v, f);
      if (::testing::Test::HasFatalFailure()) {
        return;
      }
    }
  }
  for (int f = 0; f <= 20; ++f) {
    CheckExpo(v, f);
    if (::testing::Test::HasFatalFailure()) {
      return;
    }
  }
  for (int p = 1; p <= 17; ++p) {
    CheckGeneral(v, p);
    if (::testing::Test::HasFatalFailure()) {
      return;
    }
  }
}

std::vector<double> HandPicked() {
  std::vector<double> base = {0.0,
                              0.5,
                              1.5,
                              2.5,
                              0.125,
                              0.375,
                              1234567890123445.0,
                              1234567890123455.0,
                              9.5,
                              99.5,
                              999.5,
                              0.95,
                              0.0005,
                              1e-7,
                              123.456,
                              DBL_MIN,
                              std::numeric_limits<double>::denorm_min(),
                              9007199254740992.0,
                              9007199254740994.0,
                              1e21,
                              1e22,
                              1e300,
                              0.1,
                              0.3,
                              1.0,
                              100.0,
                              1e15,
                              123456789012345678.0,
                              5e-324 * 3,
                              4.35,
                              0.045,
                              1e23,
                              9.999999999999999e22};
  std::vector<double> out;
  for (double v : base) {
    out.push_back(v);
    out.push_back(-v);
  }
  return out;
}

TEST(NumberText, HandPickedValues) {
  for (double v : HandPicked()) {
    CheckAll(v);
    if (HasFatalFailure()) {
      return;
    }
  }
}

TEST(NumberText, DblMaxAndSpecials) {
  for (double v : {DBL_MAX, -DBL_MAX}) {
    for (int f = 0; f <= 2; ++f) {
      CheckFixed(v, f);
    }
    for (int f = 0; f <= 20; ++f) {
      CheckExpo(v, f);
    }
    for (int p = 1; p <= 17; ++p) {
      CheckGeneral(v, p);
    }
  }
  const double inf = std::numeric_limits<double>::infinity();
  const double nan = std::numeric_limits<double>::quiet_NaN();
  for (double v : {inf, -inf, nan}) {
    CheckFixed(v, 3);
    CheckExpo(v, 3);
    CheckGeneral(v, 6);
  }
  // musl prints the sign of a NaN; Apple's libc does not, so pin it directly.
  EXPECT_EQ(Fixed(-nan, 2), "-nan");
  EXPECT_EQ(Expo(-nan, 2), "-nan");
  EXPECT_EQ(General(-nan, 2), "-nan");
}

TEST(NumberText, LongFractionsAreExact) {
  for (double v : {0.1, 1e-300, 123.456, DBL_MIN}) {
    for (int f : {40, 100, 400}) {
      CheckFixed(v, f);
    }
  }
}

TEST(NumberText, TruncationReturnsFullLength) {
  char small[6];
  EXPECT_EQ(format_fixed(small, sizeof(small), 123456.789, 3), 10);
  EXPECT_STREQ(small, "12345");
  EXPECT_EQ(format_signed(small, sizeof(small), 123456789, 0), 9);
  EXPECT_EQ(format_fixed(nullptr, 0, 1.5, 1), 3);
}

TEST(NumberText, RandomMagnitudes) {
  std::mt19937_64 rng(12345);
  std::uniform_real_distribution<double> sig(1.0, 10.0);
  std::uniform_int_distribution<int> mag(-30, 30);
  std::uniform_int_distribution<int> sign(0, 1);
  for (int i = 0; i < 100000; ++i) {
    double v = sig(rng) * std::pow(10.0, mag(rng));
    if (sign(rng) != 0) {
      v = -v;
    }
    const int f = static_cast<int>(rng() % 21U);
    const int p = 1 + static_cast<int>(rng() % 17U);
    CheckExpo(v, f);
    CheckGeneral(v, p);
    CheckFixed(v, f);
    if (HasFatalFailure()) {
      return;
    }
  }
}

TEST(NumberText, ExactTies) {
  std::mt19937_64 rng(777);
  const double denoms[] = {2.0, 8.0, 1024.0};
  for (int i = 0; i < 100000; ++i) {
    const double k = static_cast<double>(rng() % 4000000000ULL);
    const double v = k / denoms[i % 3];
    const int f = static_cast<int>(rng() % 21U);
    CheckFixed(v, f);
    CheckExpo(v, f);
    CheckGeneral(v, 1 + static_cast<int>(rng() % 17U));
    if (HasFatalFailure()) {
      return;
    }
  }
  // Integers below 2^53 tie when rounded left of their last digits.
  for (int i = 0; i < 20000; ++i) {
    const double v = static_cast<double>(rng() % (1ULL << 53U));
    CheckExpo(v, 14);
    CheckExpo(v, 15);
    CheckGeneral(v, 15);
    CheckGeneral(v, 16);
    CheckFixed(v, 0);
    if (HasFatalFailure()) {
      return;
    }
  }
  // Odd multiples of 5 * 10^k reach left-of-point ties at every position.
  for (int i = 0; i < 20000; ++i) {
    const double v = static_cast<double>((rng() % 1000000ULL) * 10ULL + 5ULL) * std::pow(10.0, static_cast<int>(i % 8));
    for (int p = 1; p <= 17; ++p) {
      CheckGeneral(v, p);
    }
    if (HasFatalFailure()) {
      return;
    }
  }
}

TEST(NumberText, RandomBitPatterns) {
  std::mt19937_64 rng(99);
  int checked = 0;
  while (checked < 50000) {
    const std::uint64_t bits = rng();
    double v;
    std::memcpy(&v, &bits, sizeof(v));
    if (!std::isfinite(v)) {
      continue;
    }
    ++checked;
    const int f = static_cast<int>(rng() % 21U);
    CheckExpo(v, f);
    CheckGeneral(v, 1 + static_cast<int>(rng() % 17U));
    if (std::fabs(v) < 1e30) {
      CheckFixed(v, f);
    }
    if (HasFatalFailure()) {
      return;
    }
  }
}

template <typename T>
std::string RefInt(const char* fmt, T v) {
  char buf[64];
  std::snprintf(buf, sizeof(buf), fmt, v);
  return buf;
}

TEST(NumberText, SignedIntegers) {
  const long long values[] = {0, 1, -1, 9, 10, -10, 12345, -12345, INT_MIN, INT_MAX, LLONG_MIN, LLONG_MAX};
  for (long long v : values) {
    char buf[64];
    format_signed(buf, sizeof(buf), v);
    EXPECT_EQ(std::string(buf), RefInt("%lld", v));
    if (v >= INT_MIN && v <= INT_MAX) {
      EXPECT_EQ(std::string(buf), RefInt("%d", static_cast<int>(v)));
      format_signed(buf, sizeof(buf), v, 4);
      EXPECT_EQ(std::string(buf), RefInt("%04d", static_cast<int>(v)));
      format_signed(buf, sizeof(buf), v, 12);
      EXPECT_EQ(std::string(buf), RefInt("%012d", static_cast<int>(v)));
    }
    format_signed(buf, sizeof(buf), v, 25);
    EXPECT_EQ(std::string(buf), RefInt("%025lld", v));
    format_signed(buf, sizeof(buf), v, 2);
    EXPECT_EQ(std::string(buf), RefInt("%02lld", v));
  }
}

TEST(NumberText, UnsignedAndHex) {
  const unsigned long long values[] = {
      0, 1, 9, 10, 255, 4095, 65535, 0xABCDEFULL, 0xDEADBEEFULL, UINT_MAX, ULLONG_MAX, ULLONG_MAX - 1, 1ULL << 63U};
  for (unsigned long long v : values) {
    char buf[64];
    format_unsigned(buf, sizeof(buf), v);
    EXPECT_EQ(std::string(buf), RefInt("%llu", v));
    format_unsigned(buf, sizeof(buf), v, 4);
    EXPECT_EQ(std::string(buf), RefInt("%04llu", v));
    if (v <= UINT_MAX) {
      format_unsigned(buf, sizeof(buf), v);
      EXPECT_EQ(std::string(buf), RefInt("%u", static_cast<unsigned>(v)));
      format_hex(buf, sizeof(buf), v, 4, true);
      EXPECT_EQ(std::string(buf), RefInt("%04X", static_cast<unsigned>(v)));
      format_hex(buf, sizeof(buf), v, 6, true);
      EXPECT_EQ(std::string(buf), RefInt("%06X", static_cast<unsigned>(v)));
      format_hex(buf, sizeof(buf), v, 8, true);
      EXPECT_EQ(std::string(buf), RefInt("%08X", static_cast<unsigned>(v)));
      format_hex(buf, sizeof(buf), v, 12, true);
      EXPECT_EQ(std::string(buf), RefInt("%012X", static_cast<unsigned>(v)));
      format_hex(buf, sizeof(buf), v, 2, true);
      EXPECT_EQ(std::string(buf), RefInt("%02X", static_cast<unsigned>(v)));
      format_hex(buf, sizeof(buf), v, 4, false);
      EXPECT_EQ(std::string(buf), RefInt("%04x", static_cast<unsigned>(v)));
    }
    format_hex(buf, sizeof(buf), v, 0, true);
    EXPECT_EQ(std::string(buf), RefInt("%llX", v));
  }
}

}  // namespace
}  // namespace formulon
