#include <cstddef>
#include <vector>

#include "eval/aggregate_filter.h"
#include "gtest/gtest.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

TEST(AggregateFilterStorage, CompactsStableAndRetainsCellBuffer) {
  std::vector<Value> cells;
  cells.reserve(8);
  for (double value : {10.0, 20.0, 30.0, 40.0, 50.0}) {
    cells.push_back(Value::number(value));
  }
  const Value* const original_data = cells.data();
  const std::size_t original_capacity = cells.capacity();
  const std::vector<bool> hidden = {false, true, false, false, true};
  const std::vector<bool> nested = {false, false, false, true, false};

  aggregate_filter_detail::compact_by_flags(&cells, hidden, /*use_hidden=*/true, nested, /*use_nested=*/true);

  EXPECT_EQ(cells.data(), original_data);
  EXPECT_EQ(cells.capacity(), original_capacity);
  ASSERT_EQ(cells.size(), 2U);
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 10.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 30.0);
}

TEST(AggregateFilterStorage, MissingFlagsKeepCellsBeyondEachMask) {
  std::vector<Value> cells = {Value::number(1.0), Value::number(2.0), Value::number(3.0), Value::number(4.0)};
  const std::vector<bool> hidden = {false, true};
  const std::vector<bool> nested = {false, false, true};

  aggregate_filter_detail::compact_by_flags(&cells, hidden, /*use_hidden=*/true, nested, /*use_nested=*/true);

  ASSERT_EQ(cells.size(), 2U);
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 4.0);
}

TEST(AggregateFilterStorage, DisabledFlagDoesNotDropItsCells) {
  std::vector<Value> cells = {Value::number(1.0), Value::number(2.0), Value::number(3.0)};
  const std::vector<bool> hidden = {false, true, false};
  const std::vector<bool> nested = {false, false, true};

  aggregate_filter_detail::compact_by_flags(&cells, hidden, /*use_hidden=*/false, nested, /*use_nested=*/true);

  ASSERT_EQ(cells.size(), 2U);
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 2.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
