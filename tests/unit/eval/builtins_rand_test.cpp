// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for the random-number built-ins RAND and RANDBETWEEN.
// Both are volatile, so the tests exercise distributional properties over
// a batch of samples rather than asserting exact values: each formula is
// parsed once per iteration and evaluated through the default registry.
//
// Rationale for the 1000-sample batches: with a correctly seeded RNG the
// probability of every sample collapsing onto a single value in
// `RANDBETWEEN(1, 10)` (the distinct-values sanity check) is 10 * (1/10)^1000,
// i.e. vanishingly small. The tests are therefore effectively deterministic
// against a correct implementation and still survive if the RNG changes.
//
// No oracle fixtures: the generator cannot reproduce nondeterministic values
// from Excel, so volatility is asserted structurally here instead.
#include <set>
#include <string_view>

#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Parses `src` and evaluates it via the default function registry. The
// thread-local arenas keep text payloads readable for the immediately
// following EXPECT_*. Each call resets the arenas to avoid cross-test
// contamination.
Value EvalSource(std::string_view src) {
  static thread_local Arena parse_arena;
  static thread_local Arena eval_arena;
  parse_arena.reset();
  eval_arena.reset();
  parser::Parser p(src, parse_arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return evaluate(*root, eval_arena);
}

// ---------------------------------------------------------------------------
// Registry pin -- catches accidental drops / renames during refactors.
// ---------------------------------------------------------------------------

TEST(BuiltinsRandRegistry, NamesRegistered) {
  const FunctionRegistry& reg = default_registry();
  EXPECT_NE(reg.lookup("RAND"), nullptr);
  EXPECT_NE(reg.lookup("RANDBETWEEN"), nullptr);
}

// ---------------------------------------------------------------------------
// RAND
// ---------------------------------------------------------------------------

TEST(BuiltinsRand, SamplesWithinHalfOpenUnitInterval) {
  // 1000 samples. Any sample outside [0, 1) indicates a distribution bug.
  for (int i = 0; i < 1000; ++i) {
    const Value v = EvalSource("=RAND()");
    ASSERT_TRUE(v.is_number()) << "iter=" << i;
    const double x = v.as_number();
    EXPECT_GE(x, 0.0) << "iter=" << i;
    EXPECT_LT(x, 1.0) << "iter=" << i;
  }
}

TEST(BuiltinsRand, ProducesDistinctSamples) {
  // Sanity that the RNG is actually advancing; if the state were frozen
  // every iteration would return the same value.
  std::set<double> distinct;
  for (int i = 0; i < 50; ++i) {
    const Value v = EvalSource("=RAND()");
    ASSERT_TRUE(v.is_number());
    distinct.insert(v.as_number());
  }
  EXPECT_GT(distinct.size(), 1u);
}

// ---------------------------------------------------------------------------
// RANDBETWEEN
// ---------------------------------------------------------------------------

TEST(BuiltinsRandBetween, IntegerRangeOneToTen) {
  std::set<double> distinct;
  for (int i = 0; i < 1000; ++i) {
    const Value v = EvalSource("=RANDBETWEEN(1,10)");
    ASSERT_TRUE(v.is_number()) << "iter=" << i;
    const double x = v.as_number();
    EXPECT_GE(x, 1.0) << "iter=" << i;
    EXPECT_LE(x, 10.0) << "iter=" << i;
    // Every sample must be an exact integer.
    EXPECT_EQ(x, static_cast<double>(static_cast<long long>(x))) << "iter=" << i << " value=" << x;
    distinct.insert(x);
  }
  // Distinct-value sanity: with p ~ 1 - 10 * (1/10)^1000 this passes on a
  // working RNG but catches a frozen-state bug where every sample matches.
  EXPECT_GE(distinct.size(), 2u);
}

TEST(BuiltinsRandBetween, FractionalBoundsRoundInward) {
  // bottom=3.2 -> ceil=4, top=7.9 -> floor=7, so results must lie in [4, 7].
  for (int i = 0; i < 1000; ++i) {
    const Value v = EvalSource("=RANDBETWEEN(3.2,7.9)");
    ASSERT_TRUE(v.is_number()) << "iter=" << i;
    const double x = v.as_number();
    EXPECT_GE(x, 4.0) << "iter=" << i;
    EXPECT_LE(x, 7.0) << "iter=" << i;
    EXPECT_EQ(x, static_cast<double>(static_cast<long long>(x))) << "iter=" << i << " value=" << x;
  }
}

TEST(BuiltinsRandBetween, InvertedBoundsReturnsNum) {
  // ceil(5)=5 > floor(3)=3 -> #NUM!.
  const Value v = EvalSource("=RANDBETWEEN(5,3)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsRandBetween, NonNumericTextReturnsValue) {
  // Non-numeric text coercion fails -> #VALUE!.
  const Value v = EvalSource("=RANDBETWEEN(\"a\",5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsRandBetween, NegativeRange) {
  for (int i = 0; i < 500; ++i) {
    const Value v = EvalSource("=RANDBETWEEN(-3,-1)");
    ASSERT_TRUE(v.is_number()) << "iter=" << i;
    const double x = v.as_number();
    EXPECT_GE(x, -3.0) << "iter=" << i;
    EXPECT_LE(x, -1.0) << "iter=" << i;
    EXPECT_EQ(x, static_cast<double>(static_cast<long long>(x))) << "iter=" << i << " value=" << x;
  }
}

TEST(BuiltinsRandBetween, BoolBottomCoercesToOne) {
  // Excel coerces TRUE to 1, so the effective range is [1, 5].
  for (int i = 0; i < 500; ++i) {
    const Value v = EvalSource("=RANDBETWEEN(TRUE,5)");
    ASSERT_TRUE(v.is_number()) << "iter=" << i;
    const double x = v.as_number();
    EXPECT_GE(x, 1.0) << "iter=" << i;
    EXPECT_LE(x, 5.0) << "iter=" << i;
    EXPECT_EQ(x, static_cast<double>(static_cast<long long>(x))) << "iter=" << i << " value=" << x;
  }
}

}  // namespace
}  // namespace eval
}  // namespace formulon
