//
// Implementation of the numeric aggregation kernels declared in
// `aggregate_kernels.h`. See that header for the rationale around algorithm
// choice (two-pass variance, Excel-aligned percentile position formulas).

#include "eval/aggregate_kernels.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <vector>

#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace aggregate_kernels {

Expected<double, ErrorCode> run_sum(const std::vector<double>& xs) {
  double total = 0.0;
  for (double x : xs) {
    total += x;
  }
  if (!std::isfinite(total)) {
    return ErrorCode::Num;
  }
  return total;
}

Expected<double, ErrorCode> run_product(const std::vector<double>& xs) {
  if (xs.empty()) {
    return 0.0;
  }
  double total = 1.0;
  for (double x : xs) {
    total *= x;
  }
  if (!std::isfinite(total)) {
    return ErrorCode::Num;
  }
  return total;
}

Expected<double, ErrorCode> run_average(const std::vector<double>& xs) {
  if (xs.empty()) {
    return ErrorCode::Div0;
  }
  double total = 0.0;
  for (double x : xs) {
    total += x;
  }
  const double avg = total / static_cast<double>(xs.size());
  if (!std::isfinite(avg)) {
    return ErrorCode::Num;
  }
  return avg;
}

namespace {

// Shared min / max body. Returns 0 on empty input to match the SUBTOTAL /
// AGGREGATE convention. Mirrors the iteration order of both pre-
// consolidation impls so the IEEE-754 comparison results are bit-stable.
Expected<double, ErrorCode> run_extreme(const std::vector<double>& xs, bool want_max) {
  if (xs.empty()) {
    return 0.0;
  }
  double best = xs[0];
  for (std::size_t i = 1; i < xs.size(); ++i) {
    if (want_max ? (xs[i] > best) : (xs[i] < best)) {
      best = xs[i];
    }
  }
  if (!std::isfinite(best)) {
    return ErrorCode::Num;
  }
  return best;
}

}  // namespace

Expected<double, ErrorCode> run_max(const std::vector<double>& xs) {
  return run_extreme(xs, /*want_max=*/true);
}

Expected<double, ErrorCode> run_min(const std::vector<double>& xs) {
  return run_extreme(xs, /*want_max=*/false);
}

Expected<double, ErrorCode> run_variance(const std::vector<double>& xs, bool sample) {
  const std::size_t n = xs.size();
  const std::size_t need = sample ? 2U : 1U;
  if (n < need) {
    return ErrorCode::Div0;
  }
  // Two-pass: compute the mean first, then accumulate squared
  // deviations against the captured mean. Matches the SUBTOTAL /
  // AGGREGATE pre-consolidation implementations.
  double sum = 0.0;
  for (double x : xs) {
    sum += x;
  }
  const double mean = sum / static_cast<double>(n);
  double sq = 0.0;
  for (double x : xs) {
    const double d = x - mean;
    sq += d * d;
  }
  const double denom = sample ? static_cast<double>(n - 1) : static_cast<double>(n);
  const double var = sq / denom;
  if (!std::isfinite(var)) {
    return ErrorCode::Num;
  }
  return var;
}

Expected<double, ErrorCode> run_stdev(const std::vector<double>& xs, bool sample) {
  auto var = run_variance(xs, sample);
  if (!var) {
    return var.error();
  }
  const double v = var.value();
  if (v < 0.0) {
    return ErrorCode::Num;
  }
  return std::sqrt(v);
}

Expected<double, ErrorCode> percentile_sorted_inc(const std::vector<double>& xs_sorted, double k) {
  if (xs_sorted.empty() || !std::isfinite(k) || k < 0.0 || k > 1.0) {
    return ErrorCode::Num;
  }
  const std::size_t n = xs_sorted.size();
  if (n == 1) {
    return xs_sorted[0];
  }
  // 1-based position formula `pos = 1 + k*(n-1)`. Linear interpolation
  // between `xs[floor(pos)-1]` and `xs[floor(pos)]` (both 0-based).
  const double pos = 1.0 + k * static_cast<double>(n - 1);
  const double floor_pos = std::floor(pos);
  const auto lo_index = static_cast<std::size_t>(floor_pos) - 1U;
  const double frac = pos - floor_pos;
  if (frac == 0.0 || lo_index + 1U >= n) {
    return xs_sorted[lo_index];
  }
  // Weighted endpoints avoid overflowing `high - low` for a legitimate
  // extreme pair such as {-1E308, 1E308}; the result remains inside the
  // closed interval spanned by the two finite neighbours.
  const double interpolated = (1.0 - frac) * xs_sorted[lo_index] + frac * xs_sorted[lo_index + 1U];
  if (!std::isfinite(interpolated)) {
    return ErrorCode::Num;
  }
  return interpolated;
}

Expected<double, ErrorCode> percentile_sorted_exc(const std::vector<double>& xs_sorted, double k) {
  if (xs_sorted.empty() || !std::isfinite(k) || k <= 0.0 || k >= 1.0) {
    return ErrorCode::Num;
  }
  const std::size_t n = xs_sorted.size();
  // 1-based position formula `pos = k*(n+1)`. The exclusive method
  // rejects positions whose integer floor is outside `[1, n-1]` (i.e.
  // `idx < 1 || idx >= n`), matching Mac Excel 365: at the upper
  // boundary `pos == n` (e.g. n=3, k=0.75) the integer floor is `n`
  // itself, which the strict `idx >= n` test rejects. The earlier
  // double-precision `pos > n` test let those boundary inputs through
  // and returned the last element instead of `#NUM!`.
  const double pos = k * static_cast<double>(n + 1);
  const double floor_pos = std::floor(pos);
  const auto idx = static_cast<std::int64_t>(floor_pos);  // 1-based; xs[idx-1] is the lower neighbour.
  if (idx < 1 || idx >= static_cast<std::int64_t>(n)) {
    return ErrorCode::Num;
  }
  const auto lo_index = static_cast<std::size_t>(idx - 1);
  const double frac = pos - floor_pos;
  if (frac == 0.0 || lo_index + 1U >= n) {
    return xs_sorted[lo_index];
  }
  const double interpolated = (1.0 - frac) * xs_sorted[lo_index] + frac * xs_sorted[lo_index + 1U];
  if (!std::isfinite(interpolated)) {
    return ErrorCode::Num;
  }
  return interpolated;
}

Expected<double, ErrorCode> mode_first_occurrence(const std::vector<double>& xs) {
  if (xs.empty()) {
    return ErrorCode::NA;
  }
  // Sort a copy by value while carrying source positions. This groups equal
  // values in O(n log n), then the smallest source position implements
  // Excel's first-occurrence tie rule without a quadratic frequency table.
  std::vector<std::pair<double, std::size_t>> ranked;
  ranked.reserve(xs.size());
  for (std::size_t i = 0; i < xs.size(); ++i) {
    ranked.emplace_back(xs[i], i);
  }
  std::sort(ranked.begin(), ranked.end(), ValueThenPositionOrder{});
  std::size_t best_count = 0;
  std::size_t best_first = xs.size();
  double best_value = 0.0;
  for (std::size_t begin = 0; begin < ranked.size();) {
    std::size_t end = begin + 1U;
    std::size_t first = ranked[begin].second;
    while (end < ranked.size() && ranked[end].first == ranked[begin].first) {
      first = std::min(first, ranked[end].second);
      ++end;
    }
    const std::size_t count = end - begin;
    if (count > best_count || (count == best_count && first < best_first)) {
      best_count = count;
      best_first = first;
      best_value = ranked[begin].first;
    }
    begin = end;
  }
  if (best_count < 2U) {
    return ErrorCode::NA;
  }
  return best_value;
}

}  // namespace aggregate_kernels
}  // namespace eval
}  // namespace formulon
