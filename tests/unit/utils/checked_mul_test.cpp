//
// Unit tests for `formulon::checked_mul_size_t` and
// `formulon::checked_mul_u64`. The helpers are the defensive seam used by
// the matrix / dynamic-array / linest / pivot kernels to keep `rows * cols`
// from silently wrapping — the `size_t` form for byte sizes (32-bit on the
// WASM main build), the `uint64_t` form for cell and work-unit counts that
// are charged to a `ResourceBudget` regardless of host word size.

#include "utils/checked_mul.h"

#include <cstddef>
#include <cstdint>
#include <limits>

#include "gtest/gtest.h"
#include "utils/error.h"

namespace formulon {
namespace {

TEST(CheckedMulSizeT, Zero) {
  // Zero short-circuits to zero regardless of the other operand magnitude.
  EXPECT_EQ(checked_mul_size_t(0U, 0U).value(), 0U);
  EXPECT_EQ(checked_mul_size_t(0U, 12345U).value(), 0U);
  EXPECT_EQ(checked_mul_size_t(54321U, 0U).value(), 0U);
  EXPECT_EQ(checked_mul_size_t(0U, std::numeric_limits<std::size_t>::max()).value(), 0U);
}

TEST(CheckedMulSizeT, NoOverflow) {
  EXPECT_EQ(checked_mul_size_t(2U, 3U).value(), 6U);
  EXPECT_EQ(checked_mul_size_t(1024U, 1024U).value(), 1024U * 1024U);
  EXPECT_EQ(checked_mul_size_t(65535U, 65535U).value(),
            static_cast<std::size_t>(65535U) * static_cast<std::size_t>(65535U));
}

TEST(CheckedMulSizeT, OverflowAtBoundary) {
  // Exactly `SIZE_MAX / 2 + 1` doubled would wrap to 1 on a 2's-complement
  // multiply, so the overflow detector must catch it.
  constexpr std::size_t kMax = std::numeric_limits<std::size_t>::max();
  const std::size_t half_plus_one = (kMax / 2U) + 1U;
  auto result = checked_mul_size_t(half_plus_one, 2U);
  ASSERT_FALSE(result);
  EXPECT_EQ(result.error().code, FormulonErrorCode::kFnOverflow);
}

TEST(CheckedMulSizeT, OverflowFromLargeFactors) {
  constexpr std::size_t kMax = std::numeric_limits<std::size_t>::max();
  // (kMax / 2) * 3 strictly exceeds kMax (it is roughly 1.5 * kMax).
  auto result = checked_mul_size_t(kMax / 2U, 3U);
  ASSERT_FALSE(result);
  EXPECT_EQ(result.error().code, FormulonErrorCode::kFnOverflow);
}

TEST(CheckedMulU64, Zero) {
  // Zero short-circuits to zero regardless of the other operand magnitude.
  EXPECT_EQ(checked_mul_u64(0U, 0U).value(), 0U);
  EXPECT_EQ(checked_mul_u64(0U, 12345U).value(), 0U);
  EXPECT_EQ(checked_mul_u64(54321U, 0U).value(), 0U);
  EXPECT_EQ(checked_mul_u64(0U, std::numeric_limits<std::uint64_t>::max()).value(), 0U);
}

TEST(CheckedMulU64, ExactMaxProductIsAccepted) {
  constexpr std::uint64_t kMax = std::numeric_limits<std::uint64_t>::max();
  // A product landing exactly on the representable maximum must be accepted,
  // in both operand orders.
  EXPECT_EQ(checked_mul_u64(kMax, 1U).value(), kMax);
  EXPECT_EQ(checked_mul_u64(1U, kMax).value(), kMax);
  // 3 * 6148914691236517205 == 2^64 - 1 exactly.
  constexpr std::uint64_t kThird = 6148914691236517205ULL;
  EXPECT_EQ(checked_mul_u64(3U, kThird).value(), kMax);
  EXPECT_EQ(checked_mul_u64(kThird, 3U).value(), kMax);
}

TEST(CheckedMulU64, OnePastMaxOverflowsInEitherOperandOrder) {
  constexpr std::uint64_t kMax = std::numeric_limits<std::uint64_t>::max();
  const std::uint64_t half_plus_one = (kMax / 2U) + 1U;

  auto forward = checked_mul_u64(half_plus_one, 2U);
  ASSERT_FALSE(static_cast<bool>(forward));
  EXPECT_EQ(forward.error().code, FormulonErrorCode::kFnOverflow);

  auto reversed = checked_mul_u64(2U, half_plus_one);
  ASSERT_FALSE(static_cast<bool>(reversed));
  EXPECT_EQ(reversed.error().code, FormulonErrorCode::kFnOverflow);
}

TEST(CheckedMulU64, FullGridCellCountIsRepresentable) {
  // A full Excel grid (2^20 rows by 2^14 columns) is the largest cell count
  // the engine ever charges to a budget; it must survive the guard on every
  // target, including the 32-bit `size_t` WASM build.
  auto result = checked_mul_u64(1048576U, 16384U);
  ASSERT_TRUE(static_cast<bool>(result));
  EXPECT_EQ(result.value(), 17179869184ULL);
}

}  // namespace
}  // namespace formulon
