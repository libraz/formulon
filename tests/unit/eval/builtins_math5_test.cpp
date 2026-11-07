//
// End-to-end tests for the PERMUT / PERMUTATIONA combinatorial builtins.
// Both are scalar-only, float-truncation inputs, with `#NUM!` on the
// documented domain violations.

#include <string_view>

#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "util/test_eval_helpers.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

using formulon::test::EvalSource;

// ---------------------------------------------------------------------------
// Registry pin
// ---------------------------------------------------------------------------

TEST(BuiltinsMath5Registry, AllNamesRegistered) {
  for (const char* name : {"PERMUT", "PERMUTATIONA"}) {
    EXPECT_NE(default_registry().lookup(name), nullptr) << "missing: " << name;
  }
}

// ---------------------------------------------------------------------------
// PERMUT
// ---------------------------------------------------------------------------

TEST(BuiltinsMath5Permut, Basic) {
  // 5 * 4 = 20.
  const Value v = EvalSource("=PERMUT(5, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 20.0);
}

TEST(BuiltinsMath5Permut, KZeroIsOne) {
  const Value v = EvalSource("=PERMUT(5, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsMath5Permut, KEqualsNIsFactorial) {
  // 5! = 120.
  const Value v = EvalSource("=PERMUT(5, 5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 120.0);
}

TEST(BuiltinsMath5Permut, KGreaterThanNIsNum) {
  const Value v = EvalSource("=PERMUT(3, 5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsMath5Permut, NegativeNIsNum) {
  const Value v = EvalSource("=PERMUT(-1, 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// PERMUTATIONA
// ---------------------------------------------------------------------------

TEST(BuiltinsMath5PermutationA, Basic) {
  // 3^2 = 9.
  const Value v = EvalSource("=PERMUTATIONA(3, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 9.0);
}

TEST(BuiltinsMath5PermutationA, ZeroZeroIsOne) {
  const Value v = EvalSource("=PERMUTATIONA(0, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsMath5PermutationA, TenPowerMatches) {
  // 2^10 = 1024.
  const Value v = EvalSource("=PERMUTATIONA(2, 10)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1024.0);
}

// ---------------------------------------------------------------------------
// CEILING.MATH / FLOOR.MATH: an omitted significance is 1 and an omitted mode
// 0; a blank reference significance is 0.
// ---------------------------------------------------------------------------

void ExpectNumberResult(const Value& v, double expected, std::string_view formula) {
  ASSERT_TRUE(v.is_number()) << formula << ": " << v.debug_to_string();
  EXPECT_EQ(v.as_number(), expected) << formula;
}

TEST(BuiltinsMath5OmittedArgs, CeilingMathOmittedSignificanceAndMode) {
  ExpectNumberResult(EvalSource("=CEILING.MATH(1.2,,)"), 2.0, "=CEILING.MATH(1.2,,)");
  ExpectNumberResult(EvalSource("=CEILING.MATH(-1.5,,)"), -1.0, "=CEILING.MATH(-1.5,,)");
  ExpectNumberResult(EvalSource("=CEILING.MATH(-1.5,,1)"), -2.0, "=CEILING.MATH(-1.5,,1)");
}

TEST(BuiltinsMath5OmittedArgs, FloorMathOmittedSignificanceAndMode) {
  ExpectNumberResult(EvalSource("=FLOOR.MATH(-1.2,,1)"), -1.0, "=FLOOR.MATH(-1.2,,1)");
  ExpectNumberResult(EvalSource("=FLOOR.MATH(1.5,,)"), 1.0, "=FLOOR.MATH(1.5,,)");
}

TEST(BuiltinsMath5OmittedArgs, CeilingMathBlankReferenceSignificanceIsZero) {
  const Workbook wb = Workbook::create();
  ExpectNumberResult(formulon::test::EvalSourceIn("=CEILING.MATH(1.2,A99)", wb, wb.sheet(0)), 0.0,
                     "=CEILING.MATH(1.2,A99)");
}

}  // namespace
}  // namespace eval
}  // namespace formulon
