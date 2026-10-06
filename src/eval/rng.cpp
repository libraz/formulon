
#include "eval/rng.h"

#include <cmath>
#include <cstdint>

namespace formulon {
namespace eval {

std::mt19937_64& thread_local_rng() {
  thread_local std::mt19937_64 rng{std::random_device{}()};
  return rng;
}

namespace {

double unit_sample(std::mt19937_64& rng) {
  std::uniform_real_distribution<double> dist(0.0, 1.0);
  return dist(rng);
}

// `double(INT64_MAX)` rounds to 2^63, so the upper bound is intentionally
// exclusive when deciding whether a floating-point endpoint can be converted
// to int64_t without leaving its representable domain.
constexpr double kInt64MinAsDouble = -0x1p63;
constexpr double kInt64UpperExclusiveAsDouble = 0x1p63;

}  // namespace

double sample_uniform_real(double min, double max, std::mt19937_64& rng) {
  if (min == max) {
    return min;
  }

  const double span = max - min;
  if (std::isfinite(span)) {
    std::uniform_real_distribution<double> dist(min, max);
    return dist(rng);
  }

  // The endpoints are finite, but their difference is not. Interpolating the
  // weighted endpoints avoids forming that overflowing difference.
  const double unit = unit_sample(rng);
  double sample = (1.0 - unit) * min + unit * max;
  if (!(sample >= min)) {
    sample = min;
  }
  if (sample >= max) {
    sample = std::nextafter(max, min);
  }
  return sample;
}

double sample_uniform_integer(double min, double max, std::mt19937_64& rng) {
  if (min == max) {
    return min;
  }

  if (min >= kInt64MinAsDouble && max < kInt64UpperExclusiveAsDouble) {
    const auto lo = static_cast<std::int64_t>(min);
    const auto hi = static_cast<std::int64_t>(max);
    std::uniform_int_distribution<std::int64_t> dist(lo, hi);
    return static_cast<double>(dist(rng));
  }

  // Out-of-range integral doubles cannot be represented by int64_t, but all
  // doubles at these magnitudes are already integral. Interpolate directly in
  // double space, then floor and clamp to retain an integral-valued result in
  // the requested closed interval.
  const double unit = unit_sample(rng);
  double sample = (1.0 - unit) * min + unit * max;
  if (!(sample >= min)) {
    sample = min;
  }
  if (sample > max) {
    sample = max;
  }
  sample = std::floor(sample);
  if (sample < min) {
    sample = min;
  }
  if (sample > max) {
    sample = max;
  }
  return sample;
}

}  // namespace eval
}  // namespace formulon
