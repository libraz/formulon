//
// Implementation of the numeric aggregation kernels declared in
// `numeric_aggregate_kernels.h`. See that header for the rationale around algorithm
// choice (two-pass variance, Excel-aligned percentile position formulas).

#include "numeric_aggregate_kernels.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <vector>

#include "utils/cancellation_snap.h"
#include "utils/expected.h"
#include "utils/index_sort.h"
#include "utils/subnormal_flush.h"
#include "value.h"

namespace formulon {
namespace numeric_aggregate_kernels {
namespace {

// Mean of the two middle elements, computed so that a pair whose sum
// overflows still lands on a finite midpoint. Same-sign operands use
// `low + (high - low) / 2` (the difference cannot overflow); opposite-sign
// operands halve first, where no overflow is possible either.
double midpoint_without_overflow(double low, double high) {
  if (std::signbit(low) == std::signbit(high)) {
    return low + (high - low) * 0.5;
  }
  return low * 0.5 + high * 0.5;
}

// Collects the numeric elements of `view` in source order; `nullopt` elements
// are skipped.
std::vector<double> materialize(const NumericInputView& view) {
  std::vector<double> xs;
  if (view.get == nullptr) {
    return xs;
  }
  xs.reserve(view.size);
  for (std::size_t i = 0; i < view.size; ++i) {
    if (const std::optional<double> value = view.get(view.context, i); value.has_value()) {
      xs.push_back(*value);
    }
  }
  return xs;
}

}  // namespace

Expected<double, ErrorCode> run_sum(const std::vector<double>& xs) {
  FinalStepSnapSum sum;
  for (double x : xs) {
    sum.add(x);
  }
  const double total = sum.result();
  if (!std::isfinite(total)) {
    return ErrorCode::Num;
  }
  return total;
}

Expected<double, ErrorCode> run_product(const std::vector<double>& xs) {
  if (xs.empty()) {
    return 0.0;
  }
  // Excel starts from the first value, so a lone subnormal is not flushed.
  double total = xs.front();
  for (std::size_t i = 1; i < xs.size(); ++i) {
    total = flush_subnormal(total * xs[i]);
  }
  if (!std::isfinite(total)) {
    return ErrorCode::Num;
  }
  return total;
}

Expected<double, ErrorCode> run_average(const NumericInputView& values) {
  if (values.get == nullptr) {
    return ErrorCode::Div0;
  }
  std::size_t count = 0;
  FinalStepSnapSum sum;
  for (std::size_t i = 0; i < values.size; ++i) {
    const std::optional<double> value = values.get(values.context, i);
    if (!value.has_value()) {
      continue;
    }
    sum.add(*value);
    ++count;
  }
  if (count == 0U) {
    return ErrorCode::Div0;
  }
  const double avg = flush_subnormal(sum.result() / static_cast<double>(count));
  if (!std::isfinite(avg)) {
    return ErrorCode::Num;
  }
  return avg;
}

Expected<double, ErrorCode> run_average(const std::vector<double>& xs) {
  if (xs.empty()) {
    return ErrorCode::Div0;
  }
  FinalStepSnapSum sum;
  for (double x : xs) {
    sum.add(x);
  }
  const double avg = flush_subnormal(sum.result() / static_cast<double>(xs.size()));
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

Expected<double, ErrorCode> run_variance(const NumericInputView& values, bool sample) {
  // Keep the contiguous-slice kernel: getter-shaped loops change compiler FP
  // contraction decisions and can shift variance/stddev by one bit.
  return run_variance(materialize(values), sample);
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
    sq += flush_subnormal(d * d);
  }
  const double denom = sample ? static_cast<double>(n - 1) : static_cast<double>(n);
  const double var = sq / denom;
  if (!std::isfinite(var)) {
    return ErrorCode::Num;
  }
  return var;
}

Expected<double, ErrorCode> run_stdev(const NumericInputView& values, bool sample) {
  return run_stdev(materialize(values), sample);
}

Expected<double, ErrorCode> run_stdev(const std::vector<double>& xs, bool sample) {
  auto var = run_variance(xs, sample);
  if (!var) {
    return std::move(var.error());
  }
  const double v = var.value();
  if (v < 0.0) {
    return ErrorCode::Num;
  }
  return std::sqrt(v);
}

namespace {

/// Value `frac` of the way from `xs[lo_index]` to its successor. Weighted
/// endpoints avoid overflowing `high - low` for a legitimate extreme pair such
/// as {-1E308, 1E308}.
Expected<double, ErrorCode> interpolate_sorted(const std::vector<double>& xs, std::size_t lo_index, double frac) {
  if (frac == 0.0 || lo_index + 1U >= xs.size()) {
    return xs[lo_index];
  }
  const double interpolated = (1.0 - frac) * xs[lo_index] + frac * xs[lo_index + 1U];
  if (!std::isfinite(interpolated)) {
    return ErrorCode::Num;
  }
  return interpolated;
}

}  // namespace

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
  return interpolate_sorted(xs_sorted, lo_index, frac);
}

Expected<double, ErrorCode> percentile_sorted_exc(const std::vector<double>& xs_sorted, double k) {
  if (xs_sorted.empty() || !std::isfinite(k) || k <= 0.0 || k >= 1.0) {
    return ErrorCode::Num;
  }
  const std::size_t n = xs_sorted.size();
  // 1-based position `pos = k*(n+1)` must lie in [1, n]; Mac Excel 365
  // accepts both ends exactly (PERCENTILE.EXC({1,2,3},0.75) is 3).
  const double pos = k * static_cast<double>(n + 1);
  const double floor_pos = std::floor(pos);
  const auto idx = static_cast<std::int64_t>(floor_pos);  // 1-based; xs[idx-1] is the lower neighbour.
  if (idx < 1 || pos > static_cast<double>(n)) {
    return ErrorCode::Num;
  }
  const auto lo_index = static_cast<std::size_t>(idx - 1);
  const double frac = pos - floor_pos;
  return interpolate_sorted(xs_sorted, lo_index, frac);
}

Expected<double, ErrorCode> run_median(std::vector<double> xs) {
  if (xs.empty()) {
    return ErrorCode::Num;
  }
  std::sort(xs.begin(), xs.end());
  const std::size_t n = xs.size();
  const double m = (n % 2U) == 1U ? xs[n / 2U] : midpoint_without_overflow(xs[n / 2U - 1U], xs[n / 2U]);
  if (!std::isfinite(m)) {
    return ErrorCode::Num;
  }
  return m;
}

std::vector<ValueRun> group_equal_values(const std::vector<double>& xs) {
  // Sort positions by value: equal values group in O(n log n), and the index
  // sort's position tie-break starts each run at its first occurrence, which
  // is Excel's tie rule.
  std::vector<std::uint32_t> ranked;
  sorted_index_order(ranked, static_cast<std::uint32_t>(xs.size()),
                     IndexLess{&xs, [](const void* context, std::uint32_t lhs, std::uint32_t rhs) {
                                 const auto& values = *static_cast<const std::vector<double>*>(context);
                                 return values[lhs] < values[rhs];
                               }});
  std::vector<ValueRun> runs;
  for (std::size_t begin = 0; begin < ranked.size();) {
    std::size_t end = begin + 1U;
    const double value = xs[ranked[begin]];
    while (end < ranked.size() && xs[ranked[end]] == value) {
      ++end;
    }
    runs.push_back(ValueRun{value, ranked[begin], end - begin});
    begin = end;
  }
  return runs;
}

void select_kth_preserving_sort_bits(std::vector<double>& xs, std::size_t target) {
  if (xs.empty() || target >= xs.size()) {
    return;
  }
  bool has_nonfinite = false;
  bool has_negative_zero = false;
  bool has_positive_zero = false;
  for (const double value : xs) {
    has_nonfinite = has_nonfinite || !std::isfinite(value);
    if (value == 0.0) {
      has_negative_zero = has_negative_zero || std::signbit(value);
      has_positive_zero = has_positive_zero || !std::signbit(value);
    }
  }
  if (has_nonfinite || (has_negative_zero && has_positive_zero)) {
    // Non-finite values or mixed signed zeros: a full sort, whose tie order fixes which bits are picked.
    std::sort(xs.begin(), xs.end());
    return;
  }
  std::nth_element(xs.begin(), xs.begin() + static_cast<std::ptrdiff_t>(target), xs.end());
}

Expected<double, ErrorCode> mode_first_occurrence(const std::vector<double>& xs) {
  if (xs.empty()) {
    return ErrorCode::NA;
  }
  std::size_t best_count = 0;
  std::size_t best_first = xs.size();
  double best_value = 0.0;
  for (const ValueRun& run : group_equal_values(xs)) {
    if (run.count > best_count || (run.count == best_count && run.first < best_first)) {
      best_count = run.count;
      best_first = run.first;
      best_value = run.value;
    }
  }
  if (best_count < 2U) {
    return ErrorCode::NA;
  }
  return best_value;
}

}  // namespace numeric_aggregate_kernels
}  // namespace formulon
