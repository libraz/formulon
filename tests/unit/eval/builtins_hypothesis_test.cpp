// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for Excel's hypothesis-test / probability family:
// T.TEST (TTEST), F.TEST (FTEST), CHISQ.TEST (CHITEST), Z.TEST (ZTEST),
// PROB. All nine names ride the tree walker's lazy-dispatch table
// because their array arguments carry shape information (paired T.TEST
// / CHISQ.TEST / PROB require matching `(rows, cols)`; the others
// collect each array independently but still need to see a range/Ref
// rather than a flattened `Value`).
//
// Reference values are cross-checked against scipy.stats where the
// closed-form p-value is not trivial.

#include <cmath>
#include <string>
#include <string_view>

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

// Parses `src` and evaluates it via the default function registry with
// no bound workbook. Suitable for `ArrayLiteral`-only formulas.
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

// Parses `src` and evaluates it against a bound workbook + current
// sheet. Used whenever the formula references cell ranges.
Value EvalSourceIn(std::string_view src, const Workbook& wb, const Sheet& current) {
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
  EvalState state;
  const EvalContext ctx(wb, current, state);
  return evaluate(*root, eval_arena, default_registry(), ctx);
}

// ---------------------------------------------------------------------------
// Registry pin
// ---------------------------------------------------------------------------

TEST(BuiltinsHypothesisRegistry, AllNamesRegistered) {
  // Every hypothesis / probability name routes through the lazy-dispatch
  // table in tree_walker.cpp. Confirm each canonical and each legacy
  // spelling is reachable by round-tripping a trivial formula through the
  // evaluator — an unregistered name would surface as `#NAME?`.
  struct NameCheck {
    const char* formula;
    bool expect_number;
  };
  const NameCheck checks[] = {
      {"=T.TEST({1,2,3,4,5}, {6,7,8,9,10}, 2, 2)", true},
      {"=TTEST({1,2,3,4,5}, {6,7,8,9,10}, 2, 2)", true},
      {"=F.TEST({1,2,3,4,5}, {2,3,4,5,6})", true},
      {"=FTEST({1,2,3,4,5}, {2,3,4,5,6})", true},
      {"=CHISQ.TEST({58,11;35,25}, {45.35,23.65;47.65,24.85})", true},
      {"=CHITEST({58,11;35,25}, {45.35,23.65;47.65,24.85})", true},
      {"=Z.TEST({3,6,7,8,6,5,4,2,1,9}, 4)", true},
      {"=ZTEST({3,6,7,8,6,5,4,2,1,9}, 4)", true},
      {"=PROB({1,2,3,4}, {0.1,0.2,0.3,0.4}, 2)", true},
  };
  for (const auto& c : checks) {
    const Value v = EvalSource(c.formula);
    ASSERT_FALSE(v.is_error() && v.as_error() == ErrorCode::Name) << "unregistered: " << c.formula;
    if (c.expect_number) {
      EXPECT_TRUE(v.is_number()) << "not a number: " << c.formula;
    }
  }
}

// ---------------------------------------------------------------------------
// T.TEST
// ---------------------------------------------------------------------------

TEST(HypothesisTTest, PairedConstantDifferenceIsDiv0) {
  // Differences are all -1 so the within-pair variance is exactly 0 and
  // the statistic is 0/0. Excel reports this as #DIV/0!.
  const Value v = EvalSource("=T.TEST({1,2,3,4,5}, {2,3,4,5,6}, 2, 1)");
  ASSERT_TRUE(v.is_error()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(HypothesisTTest, PairedValidTwoTailed) {
  // Paired two-tailed, differences d = (a - b) = {-2, -2, -3, 0, 1}.
  // mean_d = -1.2, s_d² = 2.7, t = -1.2 / sqrt(2.7/5) ~= -1.633..., df=4.
  // Reference computed via the regularized incomplete beta:
  //   y = df/(df + t²) = 4 / 6.666... ~= 0.6
  //   p = I(2, 0.5, y) ~= 0.177807808...
  const Value v = EvalSource("=T.TEST({1,2,3,4,5}, {3,4,6,4,4}, 2, 1)");
  ASSERT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_NEAR(v.as_number(), 0.17780780836, 1e-6);
}

TEST(HypothesisTTest, EqualVarianceTwoTailed) {
  // m1=3, m2=8, v1=v2=2.5, sp²=2.5, t = -5/sqrt(2.5*(1/5+1/5)) = -5.
  // df = 8. scipy: 2 * (1 - t(8).cdf(5)) ~= 0.001052826...
  const Value v = EvalSource("=T.TEST({1,2,3,4,5}, {6,7,8,9,10}, 2, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.00105282581, 1e-6);
}

TEST(HypothesisTTest, EqualVarianceOneTailIsHalfOfTwoTail) {
  const Value one = EvalSource("=T.TEST({1,2,3,4,5}, {6,7,8,9,10}, 1, 2)");
  const Value two = EvalSource("=T.TEST({1,2,3,4,5}, {6,7,8,9,10}, 2, 2)");
  ASSERT_TRUE(one.is_number());
  ASSERT_TRUE(two.is_number());
  EXPECT_NEAR(one.as_number() * 2.0, two.as_number(), 1e-12);
}

TEST(HypothesisTTest, WelchTwoTailed) {
  // Unequal-variance two-sample with m1=6.2, v1=9.7, m2=3, v2=2.5.
  //   t ~= 2.0486, df ~= 5.9334 (Welch-Satterthwaite)
  //   p_two_tailed = 2 * 0.5 * I(df/2, 0.5, df/(df+t²)) ~= 0.086938...
  const Value v = EvalSource("=T.TEST({3,4,5,9,10}, {1,2,3,4,5}, 2, 3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.086938233, 1e-6);
}

TEST(HypothesisTTest, LegacyAliasMatches) {
  // TTEST is identical in every respect to T.TEST; pin the equivalence
  // with a numeric case that cannot trivially collapse to zero.
  const Value a = EvalSource("=TTEST({1,2,3,4,5}, {6,7,8,9,10}, 2, 2)");
  const Value b = EvalSource("=T.TEST({1,2,3,4,5}, {6,7,8,9,10}, 2, 2)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_DOUBLE_EQ(a.as_number(), b.as_number());
}

TEST(HypothesisTTest, PairedShapeMismatchIsNA) {
  const Value v = EvalSource("=T.TEST({1,2,3}, {1,2}, 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(HypothesisTTest, BadTypeIsNum) {
  const Value v = EvalSource("=T.TEST({1,2,3}, {4,5,6}, 2, 4)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(HypothesisTTest, BadTailsIsNum) {
  const Value v = EvalSource("=T.TEST({1,2,3}, {4,5,6}, 3, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(HypothesisTTest, TooFewPairsIsDiv0) {
  // Paired with only one pair -> s_d undefined (n < 2) -> #DIV/0!.
  const Value v = EvalSource("=T.TEST({1}, {2}, 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(HypothesisTTest, ArityWrong) {
  const Value v = EvalSource("=T.TEST({1,2,3}, {4,5,6}, 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(HypothesisTTest, TailsTruncates) {
  // Excel truncates toward zero before the {1,2} check. 1.9 -> 1.
  const Value a = EvalSource("=T.TEST({1,2,3,4,5}, {6,7,8,9,10}, 1.9, 2)");
  const Value b = EvalSource("=T.TEST({1,2,3,4,5}, {6,7,8,9,10}, 1, 2)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_DOUBLE_EQ(a.as_number(), b.as_number());
}

// ---------------------------------------------------------------------------
// F.TEST
// ---------------------------------------------------------------------------

TEST(HypothesisFTest, Basic) {
  // Verified with scipy: var([6,7,9,15,21], ddof=1) = 37.3,
  // var([20,28,31,38,40], ddof=1) = 63.2, F = 37.3/63.2 ~= 0.59019,
  // df1=df2=4, cdf = scipy.stats.f(4,4).cdf(0.59019) ~= 0.324166...,
  // p = 2*min(cdf, 1-cdf) ~= 0.648332...
  const Value v = EvalSource("=F.TEST({6,7,9,15,21}, {20,28,31,38,40})");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.648333, 1e-4);
}

TEST(HypothesisFTest, LegacyAliasMatches) {
  const Value a = EvalSource("=FTEST({6,7,9,15,21}, {20,28,31,38,40})");
  const Value b = EvalSource("=F.TEST({6,7,9,15,21}, {20,28,31,38,40})");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_DOUBLE_EQ(a.as_number(), b.as_number());
}

TEST(HypothesisFTest, EqualVariancePEqualsOne) {
  // Two identical distributions (var ratio exactly 1): cdf at 1 is
  // exactly 0.5 for F(df, df), so 2 * min(0.5, 0.5) = 1.
  const Value v = EvalSource("=F.TEST({1,2,3,4,5}, {1,2,3,4,5})");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0, 1e-12);
}

TEST(HypothesisFTest, ZeroVarianceIsDiv0) {
  const Value v = EvalSource("=F.TEST({7,7,7,7,7}, {1,2,3,4,5})");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(HypothesisFTest, TooFewSamplesIsDiv0) {
  const Value v = EvalSource("=F.TEST({1}, {2,3,4,5})");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(HypothesisFTest, ArityWrong) {
  const Value v = EvalSource("=F.TEST({1,2,3})");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(HypothesisFTest, DifferentShapesAllowed) {
  // Excel does NOT require shape match for F.TEST. Pin the 3-vs-5 case.
  const Value v = EvalSource("=F.TEST({1,2,3}, {20,28,31,38,40})");
  ASSERT_TRUE(v.is_number());
  EXPECT_GE(v.as_number(), 0.0);
  EXPECT_LE(v.as_number(), 1.0);
}

// ---------------------------------------------------------------------------
// CHISQ.TEST
// ---------------------------------------------------------------------------

TEST(HypothesisChisqTest, TwoByTwoTable) {
  // 2x2 contingency table. chi² sum = 15.042... df=1,
  // p = q_gamma(0.5, 7.521...) ~ scipy.stats.chi2.sf(15.042, 1)
  //   ~ 0.0001050...
  const Value v = EvalSource("=CHISQ.TEST({58,11;35,25}, {45.35,23.65;47.65,24.85})");
  ASSERT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_GT(v.as_number(), 0.0);
  EXPECT_LT(v.as_number(), 0.001);
}

TEST(HypothesisChisqTest, LegacyAliasMatches) {
  const Value a = EvalSource("=CHITEST({58,11;35,25}, {45.35,23.65;47.65,24.85})");
  const Value b = EvalSource("=CHISQ.TEST({58,11;35,25}, {45.35,23.65;47.65,24.85})");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_DOUBLE_EQ(a.as_number(), b.as_number());
}

TEST(HypothesisChisqTest, PerfectMatchIsOne) {
  // actual == expected exactly -> chi² = 0 -> q_gamma(df/2, 0) = 1.
  const Value v = EvalSource("=CHISQ.TEST({10,20;30,40}, {10,20;30,40})");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0, 1e-12);
}

TEST(HypothesisChisqTest, ShapeMismatchIsNA) {
  const Value v = EvalSource("=CHISQ.TEST({1,2,3}, {1,2,3,4})");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(HypothesisChisqTest, NonNumericCellIsValue) {
  // Excel inline arrays don't accept string literals, so build the
  // contingency table in a workbook with text in A2 to exercise the
  // non-numeric-cell rejection.
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.set_cell_value(0, 0, Value::number(1.0));
  s.set_cell_value(1, 0, Value::text("x"));
  s.set_cell_value(2, 0, Value::number(3.0));
  s.set_cell_value(0, 1, Value::number(1.0));
  s.set_cell_value(1, 1, Value::number(2.0));
  s.set_cell_value(2, 1, Value::number(3.0));
  const Value v = EvalSourceIn("=CHISQ.TEST(A1:A3, B1:B3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(HypothesisChisqTest, ExpectedZeroIsDiv0) {
  // Mac Excel returns #DIV/0! when any expected count is exactly zero
  // (the chi-squared term divides by zero).
  const Value v = EvalSource("=CHISQ.TEST({1,2,3}, {1,0,3})");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(HypothesisChisqTest, SinglePairHasNoDegreesOfFreedom) {
  // A 1x1 contingency table gives df = 0; Mac Excel surfaces this as
  // #N/A (degenerate shape, no test to run).
  const Value v = EvalSource("=CHISQ.TEST({5}, {5})");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(HypothesisChisqTest, OneDimensionalSequenceUsesNMinus1) {
  // 1x3 row: df = n - 1 = 2. actual == expected -> chi² = 0 -> p = 1.
  const Value v = EvalSource("=CHISQ.TEST({10,20,30}, {10,20,30})");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0, 1e-12);
}

TEST(HypothesisChisqTest, ArityWrong) {
  const Value v = EvalSource("=CHISQ.TEST({1,2,3})");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// Z.TEST
// ---------------------------------------------------------------------------

TEST(HypothesisZTest, SampleSigma) {
  // mean = 5.1, sample var = 6.7666..., sd ~= 2.6013, n = 10.
  //   z = (5.1 - 4) / (2.6013 / sqrt(10)) ~= 1.3372
  //   p = 0.5 * erfc(z/sqrt(2)) ~= 0.090574...
  const Value v = EvalSource("=Z.TEST({3,6,7,8,6,5,4,2,1,9}, 4)");
  ASSERT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_NEAR(v.as_number(), 0.090574197, 1e-6);
}

TEST(HypothesisZTest, KnownSigma) {
  // With supplied sigma = 2: z = 1.1 / (2/sqrt(10)) ~= 1.7393,
  //   p = 0.5 * erfc(z/sqrt(2)) ~= 0.040995...
  const Value v = EvalSource("=Z.TEST({3,6,7,8,6,5,4,2,1,9}, 4, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.040995160, 1e-6);
}

TEST(HypothesisZTest, LegacyAliasMatches) {
  const Value a = EvalSource("=ZTEST({3,6,7,8,6,5,4,2,1,9}, 4)");
  const Value b = EvalSource("=Z.TEST({3,6,7,8,6,5,4,2,1,9}, 4)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_DOUBLE_EQ(a.as_number(), b.as_number());
}

TEST(HypothesisZTest, NegativeSigmaIsNum) {
  const Value v = EvalSource("=Z.TEST({1,2,3,4,5}, 3, -1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(HypothesisZTest, ZeroSigmaIsNum) {
  const Value v = EvalSource("=Z.TEST({1,2,3,4,5}, 3, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(HypothesisZTest, ConstantSampleNoSigmaIsDiv0) {
  const Value v = EvalSource("=Z.TEST({7,7,7,7}, 7)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(HypothesisZTest, ArityWrong) {
  const Value under = EvalSource("=Z.TEST({1,2,3})");
  ASSERT_TRUE(under.is_error());
  EXPECT_EQ(under.as_error(), ErrorCode::Value);
  const Value over = EvalSource("=Z.TEST({1,2,3}, 2, 1, 0)");
  ASSERT_TRUE(over.is_error());
  EXPECT_EQ(over.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// PROB
// ---------------------------------------------------------------------------

TEST(HypothesisProb, SinglePoint) {
  const Value v = EvalSource("=PROB({1,2,3,4}, {0.1,0.2,0.3,0.4}, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.2, 1e-12);
}

TEST(HypothesisProb, Range) {
  const Value v = EvalSource("=PROB({1,2,3,4}, {0.1,0.2,0.3,0.4}, 2, 3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.5, 1e-12);
}

TEST(HypothesisProb, RangeAllInclusive) {
  const Value v = EvalSource("=PROB({1,2,3,4}, {0.1,0.2,0.3,0.4}, 1, 4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0, 1e-12);
}

TEST(HypothesisProb, NoMatchIsZero) {
  const Value v = EvalSource("=PROB({1,2,3,4}, {0.1,0.2,0.3,0.4}, 99)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(HypothesisProb, UpperLessThanLowerIsZero) {
  // Mac Excel returns 0 when lower > upper (empty interval has no mass),
  // not a parameter error.
  const Value v = EvalSource("=PROB({1,2,3,4}, {0.1,0.2,0.3,0.4}, 3, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(HypothesisProb, ProbsSumNotOneIsNum) {
  const Value v = EvalSource("=PROB({1,2,3,4}, {0.1,0.2,0.3,0.5}, 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(HypothesisProb, NegativeProbIsNum) {
  const Value v = EvalSource("=PROB({1,2,3,4}, {0.1,0.2,0.3,-0.1}, 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(HypothesisProb, ProbAboveOneIsNum) {
  const Value v = EvalSource("=PROB({1,2,3,4}, {0.1,0.2,0.3,1.4}, 2)");
  ASSERT_TRUE(v.is_error());
  // 1.4 > 1 triggers the per-cell check before the sum check.
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(HypothesisProb, ShapeMismatchIsNA) {
  const Value v = EvalSource("=PROB({1,2,3,4}, {0.1,0.2,0.3}, 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(HypothesisProb, ArityWrong) {
  const Value under = EvalSource("=PROB({1,2,3}, {0.2,0.3,0.5})");
  ASSERT_TRUE(under.is_error());
  EXPECT_EQ(under.as_error(), ErrorCode::Value);
  const Value over = EvalSource("=PROB({1,2,3}, {0.2,0.3,0.5}, 1, 2, 3)");
  ASSERT_TRUE(over.is_error());
  EXPECT_EQ(over.as_error(), ErrorCode::Value);
}

TEST(HypothesisProb, WithRanges) {
  // Range-argument smoke test: {1,2,3,4} in column A and matching
  // probabilities in column B. Confirms the lazy path works for the
  // non-ArrayLiteral carrier.
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.set_cell_value(0, 0, Value::number(1.0));
  s.set_cell_value(1, 0, Value::number(2.0));
  s.set_cell_value(2, 0, Value::number(3.0));
  s.set_cell_value(3, 0, Value::number(4.0));
  s.set_cell_value(0, 1, Value::number(0.1));
  s.set_cell_value(1, 1, Value::number(0.2));
  s.set_cell_value(2, 1, Value::number(0.3));
  s.set_cell_value(3, 1, Value::number(0.4));
  const Value v = EvalSourceIn("=PROB(A1:A4, B1:B4, 2, 3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_NEAR(v.as_number(), 0.5, 1e-12);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
