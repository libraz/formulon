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
  ResourceBudget budget(100U, "test_part");
  ASSERT_TRUE(static_cast<bool>(budget.consume(40U)));
  EXPECT_EQ(budget.used(), 40U);
  EXPECT_EQ(budget.remaining(), 60U);
  ASSERT_TRUE(static_cast<bool>(budget.consume(60U)));
  EXPECT_EQ(budget.used(), 100U);
  EXPECT_EQ(budget.remaining(), 0U);
}

TEST(ResourceBudget, ConsumePastCeilingFailsAndLeavesCountUntouched) {
  ResourceBudget budget(100U, "test_part");
  ASSERT_TRUE(static_cast<bool>(budget.consume(90U)));
  auto over = budget.consume(11U);
  ASSERT_FALSE(static_cast<bool>(over));
  // The failed charge must not advance the counter.
  EXPECT_EQ(budget.used(), 90U);
  EXPECT_EQ(over.error().code, FormulonErrorCode::kSecurityResourceLimit);
}

TEST(ResourceBudget, RefusalIsAlwaysResourceLimitWithPartFirst) {
  ResourceBudget budget(10U, "test_part");
  EXPECT_EQ(budget.part(), "test_part");
  ASSERT_TRUE(static_cast<bool>(budget.consume(4U)));
  auto over = budget.consume(7U);
  ASSERT_FALSE(static_cast<bool>(over));
  EXPECT_EQ(over.error().code, FormulonErrorCode::kSecurityResourceLimit);
  EXPECT_EQ(over.error().context, "part=test_part requested=7 limit=10 used=4");
}

TEST(ResourceBudget, WouldExceedIsSideEffectFree) {
  ResourceBudget budget(10U, "test_part");
  EXPECT_TRUE(budget.would_exceed(11U));
  EXPECT_FALSE(budget.would_exceed(10U));
  EXPECT_EQ(budget.used(), 0U);
}

TEST(ResourceBudget, HugeRequestNearUint64CeilingDoesNotWrap) {
  ResourceBudget budget(5U, "test_part");
  // A request far larger than the headroom must fail, not wrap around the
  // subtraction and appear to fit.
  EXPECT_TRUE(budget.would_exceed(0xFFFFFFFFFFFFFFFFULL));
  auto over = budget.consume(0xFFFFFFFFFFFFFFFFULL);
  EXPECT_FALSE(static_cast<bool>(over));
  EXPECT_EQ(budget.used(), 0U);
}

}  // namespace
}  // namespace formulon
