// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for `EvalState`. These tests exercise the in-progress stack
// and memoisation map in isolation, without routing through the evaluator.

#include "eval/eval_state.h"

#include <cstdint>

#include "gtest/gtest.h"
#include "sheet.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

TEST(EvalState, InitialDepthZero) {
  EvalState state;
  EXPECT_EQ(state.depth(), 0U);
}

TEST(EvalState, PushPopRoundTrip) {
  Sheet s("Sheet1");
  EvalState state;
  ASSERT_TRUE(state.push_cell(&s, 1, 2));
  EXPECT_EQ(state.depth(), 1U);
  state.pop_cell(&s, 1, 2);
  EXPECT_EQ(state.depth(), 0U);
}

TEST(EvalState, PushSameKeyReturnsFalse) {
  Sheet s("Sheet1");
  EvalState state;
  ASSERT_TRUE(state.push_cell(&s, 1, 2));
  EXPECT_FALSE(state.push_cell(&s, 1, 2));
  // Depth should be unchanged by the failed push.
  EXPECT_EQ(state.depth(), 1U);
  state.pop_cell(&s, 1, 2);
}

TEST(EvalState, PushDifferentKeyReturnsTrue) {
  Sheet s("Sheet1");
  EvalState state;
  ASSERT_TRUE(state.push_cell(&s, 1, 2));
  ASSERT_TRUE(state.push_cell(&s, 1, 3));
  EXPECT_EQ(state.depth(), 2U);
  state.pop_cell(&s, 1, 3);
  state.pop_cell(&s, 1, 2);
  EXPECT_EQ(state.depth(), 0U);
}

TEST(EvalState, LookupMissReturnsNull) {
  Sheet s("Sheet1");
  EvalState state;
  EXPECT_EQ(state.lookup_memo(&s, 0, 0), nullptr);
  EXPECT_EQ(state.lookup_memo(&s, 100, 200), nullptr);
}

TEST(EvalState, MemoizeThenLookupReturnsStoredValue) {
  Sheet s("Sheet1");
  EvalState state;
  state.memoize(&s, 0, 0, Value::number(42.0));
  state.memoize(&s, 1, 2, Value::boolean(true));

  const Value* a = state.lookup_memo(&s, 0, 0);
  ASSERT_NE(a, nullptr);
  ASSERT_TRUE(a->is_number());
  EXPECT_EQ(a->as_number(), 42.0);

  const Value* b = state.lookup_memo(&s, 1, 2);
  ASSERT_NE(b, nullptr);
  ASSERT_TRUE(b->is_boolean());
  EXPECT_TRUE(b->as_boolean());

  EXPECT_EQ(state.lookup_memo(&s, 0, 1), nullptr);
}

TEST(EvalState, MemoizeOverwritesExisting) {
  Sheet s("Sheet1");
  EvalState state;
  state.memoize(&s, 5, 5, Value::number(1.0));
  state.memoize(&s, 5, 5, Value::number(2.0));

  const Value* v = state.lookup_memo(&s, 5, 5);
  ASSERT_NE(v, nullptr);
  ASSERT_TRUE(v->is_number());
  EXPECT_EQ(v->as_number(), 2.0);
}

TEST(EvalState, TextMemoizationReturnsSameView) {
  // Value::text stores a string_view over caller-owned storage. A literal
  // keeps the backing alive for the duration of the test.
  Sheet s("Sheet1");
  EvalState state;
  state.memoize(&s, 3, 4, Value::text("hi"));

  const Value* v = state.lookup_memo(&s, 3, 4);
  ASSERT_NE(v, nullptr);
  ASSERT_TRUE(v->is_text());
  EXPECT_EQ(v->as_text(), "hi");
}

TEST(EvalState, DistinctCoordinatesDoNotCollide) {
  // (0, 1) and (1, 0) are distinct even though they share a sheet pointer.
  Sheet s("Sheet1");
  EvalState state;
  state.memoize(&s, 0, 1, Value::number(10.0));
  state.memoize(&s, 1, 0, Value::number(20.0));

  const Value* a = state.lookup_memo(&s, 0, 1);
  const Value* b = state.lookup_memo(&s, 1, 0);
  ASSERT_NE(a, nullptr);
  ASSERT_NE(b, nullptr);
  EXPECT_EQ(a->as_number(), 10.0);
  EXPECT_EQ(b->as_number(), 20.0);
}

TEST(EvalState, PushSameRowColDifferentSheetsNotACycle) {
  // The sheet dimension must be part of the cycle key: the same (row, col)
  // on two different sheets is two distinct cells, not a cycle.
  Sheet s1("Sheet1");
  Sheet s2("Sheet2");
  EvalState state;
  ASSERT_TRUE(state.push_cell(&s1, 0, 0));
  ASSERT_TRUE(state.push_cell(&s2, 0, 0));
  EXPECT_EQ(state.depth(), 2U);
  state.pop_cell(&s2, 0, 0);
  state.pop_cell(&s1, 0, 0);
  EXPECT_EQ(state.depth(), 0U);
}

TEST(EvalState, MemoKeysDistinctBySheet) {
  // Memo entries for the same (row, col) on different sheets must not
  // alias. Looking up against a third sheet returns null.
  Sheet s1("Sheet1");
  Sheet s2("Sheet2");
  Sheet s3("Sheet3");
  EvalState state;
  state.memoize(&s1, 0, 0, Value::number(1.0));
  state.memoize(&s2, 0, 0, Value::number(2.0));

  const Value* a = state.lookup_memo(&s1, 0, 0);
  const Value* b = state.lookup_memo(&s2, 0, 0);
  ASSERT_NE(a, nullptr);
  ASSERT_NE(b, nullptr);
  EXPECT_EQ(a->as_number(), 1.0);
  EXPECT_EQ(b->as_number(), 2.0);

  EXPECT_EQ(state.lookup_memo(&s3, 0, 0), nullptr);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
