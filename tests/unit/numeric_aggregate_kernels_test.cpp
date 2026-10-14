// Direct tests for the borrowed and vector numeric aggregation kernels.

#include "numeric_aggregate_kernels.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <limits>
#include <optional>
#include <vector>

#include "gtest/gtest.h"

namespace formulon::numeric_aggregate_kernels {
namespace {

struct GetterProbe {
  std::vector<std::optional<double>> values;
  mutable std::vector<std::size_t> calls;
};

std::optional<double> probe_get(const void* context, std::size_t index) {
  const auto& probe = *static_cast<const GetterProbe*>(context);
  probe.calls.push_back(index);
  return probe.values[index];
}

NumericInputView view_for(GetterProbe& probe) {
  return NumericInputView{&probe, probe.values.size(), &probe_get};
}

void expect_same_result(const Expected<double, ErrorCode>& expected, const Expected<double, ErrorCode>& actual) {
  ASSERT_EQ(expected.has_value(), actual.has_value());
  if (!expected) {
    EXPECT_EQ(expected.error(), actual.error());
    return;
  }
  EXPECT_EQ(std::memcmp(&expected.value(), &actual.value(), sizeof(double)), 0);
}

GetterProbe probe_for(const std::vector<double>& values) {
  GetterProbe probe;
  probe.values.reserve(values.size());
  for (const double value : values) {
    probe.values.emplace_back(value);
  }
  return probe;
}

}  // namespace

TEST(NumericAggregateKernels, ViewMatchesVectorBitwiseAcrossFiniteAndExtremeInputs) {
  const double nan = std::numeric_limits<double>::quiet_NaN();
  const double inf = std::numeric_limits<double>::infinity();
  const std::vector<std::vector<double>> cases = {
      {0.5, -0.4, -0.1, 7.0},
      {1.0e308, -1.0e308, 3.0},
      {-1.0e308, 1.0e308, -3.0},
      {-0.0, 0.0, -0.0, 0.0},
      {inf, 1.0},
      {nan, 1.0},
  };

  for (const std::vector<double>& values : cases) {
    GetterProbe probe = probe_for(values);
    const NumericInputView view = view_for(probe);
    expect_same_result(run_average(values), run_average(view));
    expect_same_result(run_variance(values, /*sample=*/false), run_variance(view, /*sample=*/false));
    expect_same_result(run_variance(values, /*sample=*/true), run_variance(view, /*sample=*/true));
    expect_same_result(run_stdev(values, /*sample=*/false), run_stdev(view, /*sample=*/false));
    expect_same_result(run_stdev(values, /*sample=*/true), run_stdev(view, /*sample=*/true));
  }
}

TEST(NumericAggregateKernels, ViewSkipsNulloptValuesLikeACompactedVector) {
  GetterProbe probe;
  probe.values = {1.0, std::nullopt, 3.0, std::nullopt, 5.0};
  const NumericInputView view = view_for(probe);
  const std::vector<double> compacted = {1.0, 3.0, 5.0};

  expect_same_result(run_average(compacted), run_average(view));
  expect_same_result(run_variance(compacted, /*sample=*/false), run_variance(view, /*sample=*/false));
  expect_same_result(run_variance(compacted, /*sample=*/true), run_variance(view, /*sample=*/true));
  expect_same_result(run_stdev(compacted, /*sample=*/false), run_stdev(view, /*sample=*/false));
  expect_same_result(run_stdev(compacted, /*sample=*/true), run_stdev(view, /*sample=*/true));
}

TEST(NumericAggregateKernels, VarianceViewRetainsContiguousKernelRounding) {
  // This input exposed a one-bit contraction difference in getter-shaped
  // variance loops on Apple clang Release builds.
  const std::uint64_t bits[] = {0x0ULL,
                                0x0ULL,
                                0xbfb999999999999aULL,
                                0xbfd999999999999aULL,
                                0x8000000000000000ULL,
                                0xbfd999999999999aULL,
                                0x0ULL,
                                0x2564274215621181ULL,
                                0x0010000000000000ULL,
                                0xbae235467f932b74ULL,
                                0x8fd37a91d578524aULL,
                                0x2c3b98ef646ba5dfULL,
                                0x0010000000000000ULL,
                                0x8000000000000000ULL,
                                0x0010000000000000ULL,
                                0x3fe0000000000000ULL,
                                0x3fe0000000000000ULL};
  GetterProbe probe;
  std::vector<double> compacted;
  for (const std::uint64_t bit_pattern : bits) {
    double value = 0.0;
    std::memcpy(&value, &bit_pattern, sizeof(value));
    compacted.push_back(value);
    probe.values.emplace_back(value);
    probe.values.emplace_back(std::nullopt);
  }
  const NumericInputView view = view_for(probe);
  for (const bool sample : {false, true}) {
    expect_same_result(run_variance(compacted, sample), run_variance(view, sample));
    expect_same_result(run_stdev(compacted, sample), run_stdev(view, sample));
  }
}

TEST(NumericAggregateKernels, ViewGetterPassesAreSourceOrderedAndMinimal) {
  GetterProbe probe;
  probe.values = {1.0, std::nullopt, 3.0, 5.0};
  const NumericInputView view = view_for(probe);

  ASSERT_TRUE(run_average(view));
  EXPECT_EQ(probe.calls, (std::vector<std::size_t>{0U, 1U, 2U, 3U}));

  probe.calls.clear();
  ASSERT_TRUE(run_variance(view, /*sample=*/false));
  EXPECT_EQ(probe.calls, (std::vector<std::size_t>{0U, 1U, 2U, 3U}));

  probe.calls.clear();
  ASSERT_TRUE(run_stdev(view, /*sample=*/true));
  EXPECT_EQ(probe.calls, (std::vector<std::size_t>{0U, 1U, 2U, 3U}));
}

TEST(NumericAggregateKernels, KthSelectionMatchesSortBitsForMixedSignedZeros) {
  const std::vector<double> input = {-0.0, 7.0, 0.0, -3.0};
  std::vector<double> expected = input;
  std::sort(expected.begin(), expected.end());
  for (std::size_t target = 0; target < input.size(); ++target) {
    std::vector<double> selected = input;
    select_kth_preserving_sort_bits(selected, target);
    EXPECT_EQ(std::memcmp(&selected[target], &expected[target], sizeof(double)), 0) << "target=" << target;
  }
}

TEST(NumericAggregateKernels, InsufficientOrNullGetterInputsRetainDiv0Errors) {
  GetterProbe probe;
  probe.values = {std::nullopt};
  const NumericInputView view = view_for(probe);

  const auto average = run_average(view);
  const auto variance_sample = run_variance(view, /*sample=*/true);
  const auto variance_population = run_variance(view, /*sample=*/false);
  const auto stdev_sample = run_stdev(view, /*sample=*/true);
  ASSERT_FALSE(average);
  ASSERT_FALSE(variance_sample);
  ASSERT_FALSE(variance_population);
  ASSERT_FALSE(stdev_sample);
  EXPECT_EQ(average.error(), ErrorCode::Div0);
  EXPECT_EQ(variance_sample.error(), ErrorCode::Div0);
  EXPECT_EQ(variance_population.error(), ErrorCode::Div0);
  EXPECT_EQ(stdev_sample.error(), ErrorCode::Div0);

  probe.values = {std::nullopt, 4.0};
  const NumericInputView one_value = view_for(probe);
  ASSERT_TRUE(run_average(one_value));
  EXPECT_DOUBLE_EQ(run_average(one_value).value(), 4.0);
  ASSERT_TRUE(run_variance(one_value, /*sample=*/false));
  EXPECT_DOUBLE_EQ(run_variance(one_value, /*sample=*/false).value(), 0.0);

  const NumericInputView null_getter{nullptr, 2U, nullptr};
  EXPECT_EQ(run_average(null_getter).error(), ErrorCode::Div0);
  EXPECT_EQ(run_variance(null_getter, /*sample=*/false).error(), ErrorCode::Div0);
  EXPECT_EQ(run_stdev(null_getter, /*sample=*/true).error(), ErrorCode::Div0);
}

}  // namespace formulon::numeric_aggregate_kernels
