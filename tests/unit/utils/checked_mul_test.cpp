//
// Unit tests for `formulon::checked_mul_size_t`. The helper is the
// defensive seam used by the matrix / dynamic-array / linest / pivot
// kernels to keep `rows * cols` from silently wrapping on 32-bit
// `size_t` (the WASM main build).

#include "utils/checked_mul.h"

#include <cstddef>
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

}  // namespace
}  // namespace formulon
