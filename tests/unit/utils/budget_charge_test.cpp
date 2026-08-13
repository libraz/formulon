//
// Unit tests for the shared charge-then-allocate seam. The contract that
// matters is ordering: a rejected count must leave the container's capacity
// exactly where it was, and the resulting error must name both the caller's
// context and the budget's own accounting.

#include "utils/budget_charge.h"

#include <cstdint>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "utils/error.h"
#include "utils/resource_budget.h"

namespace formulon {
namespace {

TEST(BudgetCharge, SuccessfulChargeAdvancesBudget) {
  ResourceBudget budget(100U);
  ASSERT_TRUE(static_cast<bool>(charge(budget, 40U, "sheet=0 anchor=A1")));
  EXPECT_EQ(budget.used(), 40U);
}

TEST(BudgetCharge, OverCeilingChargeKeepsContextAndBudgetDiagnostics) {
  ResourceBudget budget(10U);
  auto over = charge(budget, 11U, "sheet=0 anchor=A1");
  ASSERT_FALSE(static_cast<bool>(over));
  EXPECT_EQ(over.error().code, FormulonErrorCode::kSecResourceLimit);
  EXPECT_EQ(budget.used(), 0U);

  const std::string& context = over.error().context;
  EXPECT_NE(context.find("sheet=0 anchor=A1"), std::string::npos);
  EXPECT_NE(context.find("used=0"), std::string::npos);
  EXPECT_NE(context.find("requested=11"), std::string::npos);
  EXPECT_NE(context.find("ceiling=10"), std::string::npos);
}

TEST(BudgetCharge, EmptyContextLeavesBudgetDiagnosticsUnpadded) {
  ResourceBudget budget(1U);
  auto over = charge(budget, 2U, std::string());
  ASSERT_FALSE(static_cast<bool>(over));
  EXPECT_EQ(over.error().context.rfind("used=", 0U), 0U);
}

TEST(BudgetCharge, ChargeUsesTheBudgetsConfiguredErrorCode) {
  ResourceBudget budget(1U, FormulonErrorCode::kFnOverflow);
  auto over = charge(budget, 2U, "ctx");
  ASSERT_FALSE(static_cast<bool>(over));
  EXPECT_EQ(over.error().code, FormulonErrorCode::kFnOverflow);
}

TEST(ChargeThenReserve, SuccessfulChargeReservesTheContainer) {
  ResourceBudget budget(100U);
  std::vector<int> cells;
  ASSERT_TRUE(static_cast<bool>(charge_then_reserve(budget, cells, 64U, "ctx")));
  EXPECT_GE(cells.capacity(), 64U);
  EXPECT_TRUE(cells.empty());
  EXPECT_EQ(budget.used(), 64U);
}

TEST(ChargeThenReserve, OverCeilingChargeLeavesCapacityUntouched) {
  ResourceBudget budget(8U);
  std::vector<int> cells;
  const std::size_t before = cells.capacity();

  auto over = charge_then_reserve(budget, cells, 9U, "sheet=0 record=BrtArrFmla");
  ASSERT_FALSE(static_cast<bool>(over));
  // The whole point of charging first: a rejected count must not have
  // committed the allocation on the way to being rejected.
  EXPECT_EQ(cells.capacity(), before);
  EXPECT_EQ(budget.used(), 0U);

  const std::string& context = over.error().context;
  EXPECT_NE(context.find("sheet=0 record=BrtArrFmla"), std::string::npos);
  EXPECT_NE(context.find("used=0"), std::string::npos);
  EXPECT_NE(context.find("requested=9"), std::string::npos);
  EXPECT_NE(context.find("ceiling=8"), std::string::npos);
}

TEST(ChargeThenReserve, ZeroCountSucceedsWithoutReserving) {
  ResourceBudget budget(8U);
  std::vector<int> cells;
  ASSERT_TRUE(static_cast<bool>(charge_then_reserve(budget, cells, 0U, "ctx")));
  EXPECT_EQ(budget.used(), 0U);
  EXPECT_TRUE(cells.empty());
}

TEST(ChargeThenReserve, HugeCountFailsInsteadOfNarrowing) {
  // A count past any addressable size must be rejected by the ceiling before
  // it can be narrowed to `size_t`, on 64-bit hosts as well as on WASM.
  ResourceBudget budget(1024U);
  std::vector<int> cells;
  auto over = charge_then_reserve(budget, cells, 0xFFFFFFFFFFFFFFFFULL, "ctx");
  ASSERT_FALSE(static_cast<bool>(over));
  EXPECT_EQ(cells.capacity(), 0U);
  EXPECT_EQ(budget.used(), 0U);
}

}  // namespace
}  // namespace formulon
