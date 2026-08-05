//
// Unit tests for `ResourceBudget`, the request-scoped work-unit guard used
// to bound attacker-driven allocations (pivot result matrices, recalc
// viewport seeding, and similar). The tests pin the consume / would_exceed
// contract and the overflow-safe arithmetic.

#include "utils/resource_budget.h"

#include "gtest/gtest.h"
#include "utils/error.h"

namespace formulon {
namespace {

TEST(ResourceBudget, ConsumeWithinCeilingSucceedsAndAdvances) {
  ResourceBudget budget(100U);
  ASSERT_TRUE(static_cast<bool>(budget.consume(40U)));
  EXPECT_EQ(budget.used(), 40U);
  EXPECT_EQ(budget.remaining(), 60U);
  ASSERT_TRUE(static_cast<bool>(budget.consume(60U)));
  EXPECT_EQ(budget.used(), 100U);
  EXPECT_EQ(budget.remaining(), 0U);
}

TEST(ResourceBudget, ConsumePastCeilingFailsAndLeavesCountUntouched) {
  ResourceBudget budget(100U);
  ASSERT_TRUE(static_cast<bool>(budget.consume(90U)));
  auto over = budget.consume(11U);
  ASSERT_FALSE(static_cast<bool>(over));
  // The failed charge must not advance the counter.
  EXPECT_EQ(budget.used(), 90U);
  EXPECT_EQ(over.error().code, FormulonErrorCode::kSecResourceLimit);
}

TEST(ResourceBudget, ErrorCodeIsConfigurable) {
  ResourceBudget budget(1U, FormulonErrorCode::kFnOverflow);
  auto over = budget.consume(2U);
  ASSERT_FALSE(static_cast<bool>(over));
  EXPECT_EQ(over.error().code, FormulonErrorCode::kFnOverflow);
}

TEST(ResourceBudget, WouldExceedIsSideEffectFree) {
  ResourceBudget budget(10U);
  EXPECT_TRUE(budget.would_exceed(11U));
  EXPECT_FALSE(budget.would_exceed(10U));
  EXPECT_EQ(budget.used(), 0U);
}

TEST(ResourceBudget, HugeRequestNearUint64CeilingDoesNotWrap) {
  ResourceBudget budget(5U);
  // A request far larger than the headroom must fail, not wrap around the
  // subtraction and appear to fit.
  EXPECT_TRUE(budget.would_exceed(0xFFFFFFFFFFFFFFFFULL));
  auto over = budget.consume(0xFFFFFFFFFFFFFFFFULL);
  EXPECT_FALSE(static_cast<bool>(over));
  EXPECT_EQ(budget.used(), 0U);
}

}  // namespace
}  // namespace formulon
