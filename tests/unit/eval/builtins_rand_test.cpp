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
#include <array>
#include <cmath>
#include <set>
#include <string_view>
#include <utility>

#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "util/test_eval_helpers.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

using formulon::test::EvalSource;

const FunctionDef* RandBetweenDef() {
  const FunctionDef* def = default_registry().lookup("RANDBETWEEN");
  EXPECT_NE(def, nullptr) << "RANDBETWEEN must be registered in the default registry";
  return def;
}

Value CallRandBetween(Arena& arena, double bottom, double top) {
  const FunctionDef* def = RandBetweenDef();
  if (def == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  Value args[2] = {Value::number(bottom), Value::number(top)};
  return def->impl(args, 2U, arena);
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

TEST(BuiltinsRandBetween, HugeIntegralSingletonsArePreserved) {
  // Bounds outside int64_t still represent exact integral doubles. A
  // singleton must be returned unchanged instead of being cast through the
  // signed integer distribution's implementation-defined domain.
  for (const auto& sample : {std::pair<std::string_view, double>("=RANDBETWEEN(1e20,1e20)", 1e20),
                             std::pair<std::string_view, double>("=RANDBETWEEN(-1e20,-1e20)", -1e20)}) {
    const Value v = EvalSource(sample.first);
    ASSERT_TRUE(v.is_number()) << sample.first;
    EXPECT_DOUBLE_EQ(v.as_number(), sample.second) << sample.first;
  }
}

TEST(BuiltinsRandBetween, Int64BoundaryInputsStayFiniteAndIntegral) {
  // -2^63 is representable in int64_t; the nearest representable double below
  // +2^63 is also safe. The exact +2^63 endpoint and a value below -2^63 are
  // intentionally unsafe integer-distribution inputs and must still remain
  // valid double-valued results.
  const double safe_lower = -9223372036854775808.0;
  const double safe_upper = 9223372036854774784.0;
  const double unsafe_lower = -9223372036854777856.0;
  const double unsafe_upper = 9223372036854777856.0;
  const std::array<std::pair<double, double>, 5> bounds = {{
      {safe_lower, safe_lower},
      {safe_upper, safe_upper},
      {0x1p63, 0x1p63},
      {unsafe_lower, unsafe_lower},
      {unsafe_upper, unsafe_upper},
  }};
  for (const auto& sample : bounds) {
    Arena arena;
    const Value v = CallRandBetween(arena, sample.first, sample.second);
    ASSERT_TRUE(v.is_number());
    const double x = v.as_number();
    EXPECT_TRUE(std::isfinite(x));
    EXPECT_DOUBLE_EQ(x, sample.first);
    EXPECT_EQ(std::trunc(x), x);
  }
}

TEST(BuiltinsRandBetween, HugeIntegralRangesStayInRange) {
  const std::array<std::pair<std::string_view, std::pair<double, double>>, 3> ranges = {{
      {"=RANDBETWEEN(1e20,1e21)", {1e20, 1e21}},
      {"=RANDBETWEEN(-1e21,-1e20)", {-1e21, -1e20}},
      {"=RANDBETWEEN(-1e20,1e20)", {-1e20, 1e20}},
  }};
  for (const auto& sample : ranges) {
    for (int i = 0; i < 100; ++i) {
      const Value v = EvalSource(sample.first);
      ASSERT_TRUE(v.is_number()) << sample.first << " iter=" << i;
      const double x = v.as_number();
      EXPECT_TRUE(std::isfinite(x)) << sample.first << " iter=" << i;
      EXPECT_GE(x, sample.second.first) << sample.first << " iter=" << i;
      EXPECT_LE(x, sample.second.second) << sample.first << " iter=" << i;
      EXPECT_EQ(std::trunc(x), x) << sample.first << " iter=" << i;
    }
  }
}

}  // namespace
}  // namespace eval
}  // namespace formulon
