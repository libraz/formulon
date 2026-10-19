//
// Expectation for doubles whose last bits come from the host libm. Mac Excel
// shares Apple's libm, so the pins are exact on macOS; other hosts return the
// correctly rounded value, within a few ulp.

#ifndef FORMULON_TESTS_SUPPORT_HOST_LIBM_EXPECT_H_
#define FORMULON_TESTS_SUPPORT_HOST_LIBM_EXPECT_H_

#include <cmath>
#include <cstdint>
#include <cstdio>
#include <cstring>

#include "gtest/gtest.h"

namespace formulon {
namespace testing_support {

/// Distance in representable doubles between two finite values of any sign.
inline uint64_t HostLibmUlpDistance(double a, double b) {
  auto ordered = [](double v) {
    uint64_t bits;
    std::memcpy(&bits, &v, sizeof bits);
    return (bits >> 63) != 0 ? ~bits : bits | (uint64_t{1} << 63);
  };
  const uint64_t oa = ordered(a);
  const uint64_t ob = ordered(b);
  return oa > ob ? oa - ob : ob - oa;
}

/// Exact on macOS, within 4 ulp elsewhere; NaN matches NaN.
inline bool HostLibmMatch(double expected, double actual) {
  if (expected == actual || (std::isnan(expected) && std::isnan(actual))) {
    return true;
  }
#ifdef __APPLE__
  return false;
#else
  return std::isfinite(expected) && std::isfinite(actual) && HostLibmUlpDistance(expected, actual) <= 4;
#endif
}

inline ::testing::AssertionResult HostLibmEq(double actual, double expected) {
  if (HostLibmMatch(expected, actual)) {
    return ::testing::AssertionSuccess();
  }
  const bool finite = std::isfinite(actual) && std::isfinite(expected);
  const uint64_t ulps = finite ? HostLibmUlpDistance(actual, expected) : UINT64_MAX;
  char buf[160];
  std::snprintf(buf, sizeof buf, "actual %.17g, expected %.17g, distance %llu ulp", actual, expected,
                static_cast<unsigned long long>(ulps));
  return ::testing::AssertionFailure() << buf;
}

}  // namespace testing_support
}  // namespace formulon

#endif  // FORMULON_TESTS_SUPPORT_HOST_LIBM_EXPECT_H_
