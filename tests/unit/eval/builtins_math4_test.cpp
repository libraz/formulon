//
// End-to-end tests for Formulon's combinatorial, numeral-system, precise-
// rounding, and miscellaneous scalar math built-ins: ARABIC, ROMAN, BASE,
// DECIMAL, CEILING.PRECISE, FLOOR.PRECISE, ISO.CEILING, COMBIN, COMBINA,
// FACT, FACTDOUBLE, GCD, LCM, MULTINOMIAL, SQRTPI. Each test parses a
// formula source, evaluates the AST through the default registry, and
// asserts the resulting Value. Reference values are computed from the
// mathematical definitions to keep these tests independent of any golden
// oracle fixture.

#include <cmath>
#include <string_view>

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "util/test_eval_helpers.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

using formulon::test::EvalSource;
using formulon::test::EvalSourceIn;

// Local pi copy; see `builtins_math2_test.cpp` for the rationale.
constexpr double kPi = 3.14159265358979323846;

// ---------------------------------------------------------------------------
// FACT
// ---------------------------------------------------------------------------

TEST(BuiltinsMath4Fact, ZeroIsOne) {
  const Value v = EvalSource("=FACT(0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsMath4Fact, FiveFactorial) {
  const Value v = EvalSource("=FACT(5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 120.0);
}

TEST(BuiltinsMath4Fact, FractionalTruncates) {
  // 5.9 trunc -> 5; 5! = 120.
  const Value v = EvalSource("=FACT(5.9)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 120.0);
}

TEST(BuiltinsMath4Fact, NegativeYieldsNum) {
  const Value v = EvalSource("=FACT(-1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsMath4Fact, OverflowYieldsNum) {
  // 171! overflows double precision.
  const Value v = EvalSource("=FACT(171)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsMath4Fact, Boundary170) {
  // 170! is the largest representable factorial.
  const Value v = EvalSource("=FACT(170)");
  ASSERT_TRUE(v.is_number());
  EXPECT_TRUE(std::isfinite(v.as_number()));
  EXPECT_GT(v.as_number(), 0.0);
}

// ---------------------------------------------------------------------------
// FACTDOUBLE
// ---------------------------------------------------------------------------

TEST(BuiltinsMath4FactDouble, ZeroIsOne) {
  const Value v = EvalSource("=FACTDOUBLE(0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsMath4FactDouble, MinusOneIsOne) {
  // Documented quirk: -1 is the other legal non-positive input.
  const Value v = EvalSource("=FACTDOUBLE(-1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsMath4FactDouble, SevenDoubleFactorial) {
  // 7!! = 7 * 5 * 3 * 1 = 105.
  const Value v = EvalSource("=FACTDOUBLE(7)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 105.0);
}

TEST(BuiltinsMath4FactDouble, EightDoubleFactorial) {
  // 8!! = 8 * 6 * 4 * 2 = 384.
  const Value v = EvalSource("=FACTDOUBLE(8)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 384.0);
}

TEST(BuiltinsMath4FactDouble, NegativeBelowMinusOneYieldsNum) {
  const Value v = EvalSource("=FACTDOUBLE(-2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// COMBIN
// ---------------------------------------------------------------------------

TEST(BuiltinsMath4Combin, FiveChooseTwo) {
  const Value v = EvalSource("=COMBIN(5, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 10.0);
}

TEST(BuiltinsMath4Combin, ZeroChooseZeroIsOne) {
  const Value v = EvalSource("=COMBIN(0, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsMath4Combin, NChooseNIsOne) {
  const Value v = EvalSource("=COMBIN(7, 7)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsMath4Combin, KGreaterThanNYieldsNum) {
  const Value v = EvalSource("=COMBIN(3, 5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsMath4Combin, NegativeYieldsNum) {
  const Value v = EvalSource("=COMBIN(-1, 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsMath4Combin, FractionalTruncates) {
  // 5.9 -> 5, 2.5 -> 2, C(5, 2) = 10.
  const Value v = EvalSource("=COMBIN(5.9, 2.5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 10.0);
}

TEST(BuiltinsMath4Combin, ExactIntegerForSmallKLargeN) {
  // A direct n!/(k!*(n-k)!) ratio loses precision here even though the
  // true result is a small integer: n! is well beyond 2^53 by n=24, so
  // dividing the (inexact) stored factorials used to yield 23.99999...
  // and 377.99999... instead of the exact integers.
  const Value a = EvalSource("=COMBIN(24, 1)");
  ASSERT_TRUE(a.is_number());
  EXPECT_EQ(a.as_number(), 24.0);
  const Value b = EvalSource("=COMBIN(28, 2)");
  ASSERT_TRUE(b.is_number());
  EXPECT_EQ(b.as_number(), 378.0);
}

// ---------------------------------------------------------------------------
// COMBINA
// ---------------------------------------------------------------------------

TEST(BuiltinsMath4CombinA, FourChooseThreeMultichoose) {
  // COMBINA(4, 3) = C(4 + 3 - 1, 3) = C(6, 3) = 20.
  const Value v = EvalSource("=COMBINA(4, 3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 20.0);
}

TEST(BuiltinsMath4CombinA, ZeroChooseZeroIsOne) {
  const Value v = EvalSource("=COMBINA(0, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsMath4CombinA, NegativeYieldsNum) {
  const Value v = EvalSource("=COMBINA(-1, 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsMath4CombinA, KGreaterThanNAllowed) {
  // Unlike COMBIN, multichoose allows k > n: C(n+k-1, k).
  // COMBINA(3, 5) = C(7, 5) = 21.
  const Value v = EvalSource("=COMBINA(3, 5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 21.0);
}

// ---------------------------------------------------------------------------
// MULTINOMIAL
// ---------------------------------------------------------------------------

TEST(BuiltinsMath4Multinomial, SingleArgIsOne) {
  // MULTINOMIAL(n) = n!/n! = 1.
  const Value v = EvalSource("=MULTINOMIAL(5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsMath4Multinomial, TwoArgsMatchBinomial) {
  // MULTINOMIAL(2, 3) = 5! / (2! 3!) = 10.
  const Value v = EvalSource("=MULTINOMIAL(2, 3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 10.0);
}

TEST(BuiltinsMath4Multinomial, ThreeArgs) {
  // MULTINOMIAL(1, 2, 3) = 6! / (1! 2! 3!) = 720 / 12 = 60.
  const Value v = EvalSource("=MULTINOMIAL(1, 2, 3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 60.0);
}

TEST(BuiltinsMath4Multinomial, NegativeYieldsNum) {
  const Value v = EvalSource("=MULTINOMIAL(2, -1, 3)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// GCD
// ---------------------------------------------------------------------------

TEST(BuiltinsMath4Gcd, TwoArgs) {
  const Value v = EvalSource("=GCD(12, 18)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 6.0);
}

TEST(BuiltinsMath4Gcd, ThreeArgs) {
  // gcd(24, 36, 60) = 12.
  const Value v = EvalSource("=GCD(24, 36, 60)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 12.0);
}

TEST(BuiltinsMath4Gcd, AllZeroIsZero) {
  const Value v = EvalSource("=GCD(0, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsMath4Gcd, ZeroWithValueIsValue) {
  // gcd(0, n) = n by convention.
  const Value v = EvalSource("=GCD(0, 7)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 7.0);
}

TEST(BuiltinsMath4Gcd, NegativeYieldsNum) {
  const Value v = EvalSource("=GCD(12, -18)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsMath4Gcd, FractionalTruncates) {
  // 12.9 -> 12, 18.2 -> 18, gcd = 6.
  const Value v = EvalSource("=GCD(12.9, 18.2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 6.0);
}

// ---------------------------------------------------------------------------
// GCD blank-scalar policy
// ---------------------------------------------------------------------------
//
// Mac Excel 365 surfaces #VALUE! for `=GCD(A1,B1,C1)` when every Ref
// resolves to a blank cell, but returns 0 for `=GCD(A1:C1)` over the same
// blank cells. The probe golden in
// `tests/oracle/targets/mac-365-ja_JP/golden/lowrisk_probes.golden.json` records the table.

TEST(BuiltinsMath4GcdBlankScalar, AllBlankRefsYieldValueError) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=GCD(A1,B1,C1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsMath4GcdBlankScalar, BlankRefIsIgnoredWhenMixedWithLiteral) {
  Workbook wb = Workbook::create();
  for (const char* formula : {"=GCD(10,A1)", "=GCD(A1,10)"}) {
    const Value v = EvalSourceIn(formula, wb, wb.sheet(0));
    ASSERT_TRUE(v.is_number()) << formula;
    EXPECT_EQ(v.as_number(), 10.0) << formula;
  }
}

TEST(BuiltinsMath4GcdBlankScalar, BlankRefIsIgnoredWhenMixedWithNumericRef) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(12.0));
  for (const char* formula : {"=GCD(A1,B1)", "=GCD(B1,A1)"}) {
    const Value v = EvalSourceIn(formula, wb, wb.sheet(0));
    ASSERT_TRUE(v.is_number()) << formula;
    EXPECT_EQ(v.as_number(), 12.0) << formula;
  }
}

TEST(BuiltinsMath4GcdBlankScalar, BlankRefWithZeroKeepsZeroResult) {
  Workbook wb = Workbook::create();
  for (const char* formula : {"=GCD(0,A1)", "=GCD(A1,0)"}) {
    const Value v = EvalSourceIn(formula, wb, wb.sheet(0));
    ASSERT_TRUE(v.is_number()) << formula;
    EXPECT_EQ(v.as_number(), 0.0) << formula;
  }
}

TEST(BuiltinsMath4GcdBlankScalar, RangeAndBlankScalarKeepRangeResult) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(6.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(9.0));
  for (const char* formula : {"=GCD(A1:A2,C1)", "=GCD(C1,A1:A2)"}) {
    const Value v = EvalSourceIn(formula, wb, wb.sheet(0));
    ASSERT_TRUE(v.is_number()) << formula;
    EXPECT_EQ(v.as_number(), 3.0) << formula;
  }
}

TEST(BuiltinsMath4GcdBlankScalar, BooleanArgumentsStillYieldValueError) {
  const Value literal = EvalSource("=GCD(TRUE,2)");
  ASSERT_TRUE(literal.is_error());
  EXPECT_EQ(literal.as_error(), ErrorCode::Value);

  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::boolean(true));
  const Value reference = EvalSourceIn("=GCD(A1,2)", wb, wb.sheet(0));
  ASSERT_TRUE(reference.is_error());
  EXPECT_EQ(reference.as_error(), ErrorCode::Value);
}

TEST(BuiltinsMath4GcdBlankScalar, AllBlankRangeYieldsZero) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=GCD(A1:C1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsMath4GcdBlankScalar, LiteralZerosStillYieldZero) {
  // Direct numeric literals are not blank; the policy must not fire.
  const Value v = EvalSource("=GCD(0,0,0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsMath4GcdBlankScalar, BasicSanityTwelveAndEighteen) {
  const Value v = EvalSource("=GCD(12,18)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 6.0);
}

TEST(BuiltinsMath4GcdBlankScalar, LiteralEmptySlotYieldsNA) {
  const Value v = EvalSource("=GCD(,5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

// ---------------------------------------------------------------------------
// LCM
// ---------------------------------------------------------------------------

TEST(BuiltinsMath4Lcm, TwoArgs) {
  const Value v = EvalSource("=LCM(4, 6)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 12.0);
}

TEST(BuiltinsMath4Lcm, ThreeArgs) {
  // lcm(2, 3, 4) = 12.
  const Value v = EvalSource("=LCM(2, 3, 4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 12.0);
}

TEST(BuiltinsMath4Lcm, AnyZeroIsZero) {
  // Excel quirk: a zero argument forces the whole LCM to 0.
  const Value v = EvalSource("=LCM(4, 0, 6)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsMath4Lcm, NegativeYieldsNum) {
  const Value v = EvalSource("=LCM(4, -6)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsMath4Lcm, FractionalTruncates) {
  // 4.9 -> 4, 6.1 -> 6, lcm = 12.
  const Value v = EvalSource("=LCM(4.9, 6.1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 12.0);
}

// ---------------------------------------------------------------------------
// LCM blank-scalar policy (symmetric to GCD)
// ---------------------------------------------------------------------------

TEST(BuiltinsMath4LcmBlankScalar, AllBlankRefsYieldValueError) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=LCM(A1,B1,C1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsMath4LcmBlankScalar, BlankRefIsIgnoredWhenMixedWithLiteral) {
  Workbook wb = Workbook::create();
  for (const char* formula : {"=LCM(10,A1)", "=LCM(A1,10)"}) {
    const Value v = EvalSourceIn(formula, wb, wb.sheet(0));
    ASSERT_TRUE(v.is_number()) << formula;
    EXPECT_EQ(v.as_number(), 10.0) << formula;
  }
}

TEST(BuiltinsMath4LcmBlankScalar, BlankRefIsIgnoredWhenMixedWithNumericRef) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(6.0));
  for (const char* formula : {"=LCM(A1,B1)", "=LCM(B1,A1)"}) {
    const Value v = EvalSourceIn(formula, wb, wb.sheet(0));
    ASSERT_TRUE(v.is_number()) << formula;
    EXPECT_EQ(v.as_number(), 6.0) << formula;
  }
}

TEST(BuiltinsMath4LcmBlankScalar, BlankRefWithZeroKeepsZeroResult) {
  Workbook wb = Workbook::create();
  for (const char* formula : {"=LCM(0,A1)", "=LCM(A1,0)"}) {
    const Value v = EvalSourceIn(formula, wb, wb.sheet(0));
    ASSERT_TRUE(v.is_number()) << formula;
    EXPECT_EQ(v.as_number(), 0.0) << formula;
  }
}

TEST(BuiltinsMath4LcmBlankScalar, RangeAndBlankScalarKeepRangeResult) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(3.0));
  for (const char* formula : {"=LCM(A1:A2,C1)", "=LCM(C1,A1:A2)"}) {
    const Value v = EvalSourceIn(formula, wb, wb.sheet(0));
    ASSERT_TRUE(v.is_number()) << formula;
    EXPECT_EQ(v.as_number(), 6.0) << formula;
  }
}

TEST(BuiltinsMath4LcmBlankScalar, BooleanArgumentsStillYieldValueError) {
  const Value literal = EvalSource("=LCM(TRUE,2)");
  ASSERT_TRUE(literal.is_error());
  EXPECT_EQ(literal.as_error(), ErrorCode::Value);

  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::boolean(true));
  const Value reference = EvalSourceIn("=LCM(A1,2)", wb, wb.sheet(0));
  ASSERT_TRUE(reference.is_error());
  EXPECT_EQ(reference.as_error(), ErrorCode::Value);
}

TEST(BuiltinsMath4LcmBlankScalar, AllBlankRangeYieldsZero) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=LCM(A1:C1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsMath4LcmBlankScalar, LiteralZerosStillYieldZero) {
  // LCM with any zero argument is zero by Excel's quirk; literal zeros
  // are not blank, so the policy must not fire.
  const Value v = EvalSource("=LCM(0,0,0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsMath4LcmBlankScalar, BasicSanityFourAndSix) {
  const Value v = EvalSource("=LCM(4,6)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 12.0);
}

TEST(BuiltinsMath4LcmBlankScalar, LiteralEmptySlotYieldsNA) {
  const Value v = EvalSource("=LCM(,5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

// ---------------------------------------------------------------------------
// ARABIC
// ---------------------------------------------------------------------------

TEST(BuiltinsMath4Arabic, SubtractiveForm) {
  const Value v = EvalSource("=ARABIC(\"MCMXCIX\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1999.0);
}

TEST(BuiltinsMath4Arabic, AdditiveForm) {
  // MDCCCCLXXXXVIIII = 1 + 9 + 90 + 900 + 1000 but with additive digits.
  // = 1000 + 500 + 400 + 80 + 19 (additive) = 1999.
  // Additive: MDCCCCLXXXXVIIII = 1000+500+100+100+100+100+50+10+10+10+10+5+1+1+1+1 = 1999.
  const Value v = EvalSource("=ARABIC(\"MDCCCCLXXXXVIIII\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1999.0);
}

TEST(BuiltinsMath4Arabic, EmptyIsZero) {
  const Value v = EvalSource("=ARABIC(\"\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsMath4Arabic, WhitespaceIsZero) {
  const Value v = EvalSource("=ARABIC(\"   \")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsMath4Arabic, NegativePrefixAllowed) {
  const Value v = EvalSource("=ARABIC(\"-MCM\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -1900.0);
}

TEST(BuiltinsMath4Arabic, InvalidCharYieldsValue) {
  const Value v = EvalSource("=ARABIC(\"FOO\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsMath4Arabic, LowercaseAccepted) {
  const Value v = EvalSource("=ARABIC(\"xiv\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 14.0);
}

// ---------------------------------------------------------------------------
// ROMAN
// ---------------------------------------------------------------------------

TEST(BuiltinsMath4Roman, ClassicForm) {
  const Value v = EvalSource("=ROMAN(1999)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "MCMXCIX");
}

TEST(BuiltinsMath4Roman, ZeroIsEmpty) {
  const Value v = EvalSource("=ROMAN(0)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "");
}

TEST(BuiltinsMath4Roman, FourIsIV) {
  const Value v = EvalSource("=ROMAN(4)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "IV");
}

TEST(BuiltinsMath4Roman, FortyNineIsXLIX) {
  const Value v = EvalSource("=ROMAN(49)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "XLIX");
}

TEST(BuiltinsMath4Roman, OutOfRangeNegativeYieldsValue) {
  const Value v = EvalSource("=ROMAN(-1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsMath4Roman, OutOfRangeAboveYieldsValue) {
  const Value v = EvalSource("=ROMAN(4000)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsMath4Roman, FormSimplifiedIs1999Mim) {
  // Documented Excel example: ROMAN(1999, 4) = "MIM".
  const Value v = EvalSource("=ROMAN(1999, 4)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "MIM");
}

TEST(BuiltinsMath4Roman, NinetyNineFormTwoIsIC) {
  // Mac Excel 365 places IC in form 2 (not 4 as the Microsoft 499/1999
  // examples might imply). Oracle fixtures confirm ROMAN(99,2) = "IC".
  const Value v = EvalSource("=ROMAN(99, 2)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "IC");
}

TEST(BuiltinsMath4Roman, NinetyNineFormOneIsVCIV) {
  // Form 1 only enables V-subtracted pairs (VC=95), so 99 = 95+4 = VC+IV.
  const Value v = EvalSource("=ROMAN(99, 1)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "VCIV");
}

TEST(BuiltinsMath4Roman, NinetyNineClassicIsXCIX) {
  const Value v = EvalSource("=ROMAN(99, 0)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "XCIX");
}

// ---------------------------------------------------------------------------
// BASE
// ---------------------------------------------------------------------------

TEST(BuiltinsMath4Base, BinaryOfSeven) {
  const Value v = EvalSource("=BASE(7, 2)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "111");
}

TEST(BuiltinsMath4Base, HexOfFifteen) {
  const Value v = EvalSource("=BASE(15, 16)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "F");
}

TEST(BuiltinsMath4Base, Base36) {
  // 35 in base 36 is "Z".
  const Value v = EvalSource("=BASE(35, 36)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "Z");
}

TEST(BuiltinsMath4Base, MinLenPads) {
  const Value v = EvalSource("=BASE(7, 2, 8)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "00000111");
}

TEST(BuiltinsMath4Base, ZeroIsZero) {
  const Value v = EvalSource("=BASE(0, 2)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "0");
}

TEST(BuiltinsMath4Base, NegativeYieldsNum) {
  const Value v = EvalSource("=BASE(-1, 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsMath4Base, BadRadixYieldsNum) {
  const Value v = EvalSource("=BASE(5, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// DECIMAL
// ---------------------------------------------------------------------------

TEST(BuiltinsMath4Decimal, BinaryOfSeven) {
  const Value v = EvalSource("=DECIMAL(\"111\", 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 7.0);
}

TEST(BuiltinsMath4Decimal, HexOfFF) {
  const Value v = EvalSource("=DECIMAL(\"FF\", 16)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 255.0);
}

TEST(BuiltinsMath4Decimal, LowercaseAccepted) {
  const Value v = EvalSource("=DECIMAL(\"ff\", 16)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 255.0);
}

TEST(BuiltinsMath4Decimal, EmptyIsZero) {
  // Mac Excel 365: DECIMAL("", 10) -> 0 (empty input yields zero).
  const Value v = EvalSource("=DECIMAL(\"\", 10)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsMath4Decimal, WhitespaceIsZero) {
  // Whitespace-only input trims to empty, which also yields 0.
  const Value v = EvalSource("=DECIMAL(\"   \", 10)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsMath4Decimal, BadDigitYieldsNum) {
  const Value v = EvalSource("=DECIMAL(\"2\", 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsMath4Decimal, BadRadixYieldsNum) {
  const Value v = EvalSource("=DECIMAL(\"10\", 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsMath4Decimal, RoundTripWithBase) {
  // DECIMAL(BASE(n, r), r) = n.
  const Value v = EvalSource("=DECIMAL(BASE(255, 16), 16)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 255.0);
}

// ---------------------------------------------------------------------------
// CEILING.PRECISE
// ---------------------------------------------------------------------------

TEST(BuiltinsMath4CeilingPrecise, PositiveDefault) {
  const Value v = EvalSource("=CEILING.PRECISE(4.3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 5.0);
}

TEST(BuiltinsMath4CeilingPrecise, NegativeTowardPlusInf) {
  // Always rounds toward +infinity, so -4.3 -> -4 (not -5).
  const Value v = EvalSource("=CEILING.PRECISE(-4.3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -4.0);
}

TEST(BuiltinsMath4CeilingPrecise, NegativeSigIgnored) {
  // Negative sig is treated as |sig|; result matches the positive case.
  const Value v = EvalSource("=CEILING.PRECISE(-4.3, -2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -4.0);
}

TEST(BuiltinsMath4CeilingPrecise, SigZeroIsZero) {
  const Value v = EvalSource("=CEILING.PRECISE(4.3, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsMath4CeilingPrecise, DecimalExactMultipleIsNotRoundedPastIt) {
  const Value v = EvalSource("=CEILING.PRECISE(7.1, 0.1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 7.1);
}

// ---------------------------------------------------------------------------
// FLOOR.PRECISE
// ---------------------------------------------------------------------------

TEST(BuiltinsMath4FloorPrecise, PositiveDefault) {
  const Value v = EvalSource("=FLOOR.PRECISE(4.7)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 4.0);
}

TEST(BuiltinsMath4FloorPrecise, NegativeTowardMinusInf) {
  const Value v = EvalSource("=FLOOR.PRECISE(-4.3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -5.0);
}

TEST(BuiltinsMath4FloorPrecise, NegativeSigIgnored) {
  // Negative sig is treated as |sig|.
  const Value v = EvalSource("=FLOOR.PRECISE(4.7, -2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 4.0);
}

TEST(BuiltinsMath4FloorPrecise, SigZeroIsZero) {
  const Value v = EvalSource("=FLOOR.PRECISE(4.7, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsMath4FloorPrecise, DecimalExactMultipleIsNotRoundedBelowIt) {
  const Value v = EvalSource("=FLOOR.PRECISE(7.1, 0.1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 7.1);
}

// ---------------------------------------------------------------------------
// ISO.CEILING
// ---------------------------------------------------------------------------

TEST(BuiltinsMath4IsoCeiling, MatchesCeilingPrecise) {
  const Value a = EvalSource("=ISO.CEILING(4.3, 2)");
  const Value b = EvalSource("=CEILING.PRECISE(4.3, 2)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_EQ(a.as_number(), b.as_number());
}

TEST(BuiltinsMath4IsoCeiling, NegativeMatchesCeilingPrecise) {
  const Value a = EvalSource("=ISO.CEILING(-4.3)");
  const Value b = EvalSource("=CEILING.PRECISE(-4.3)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_EQ(a.as_number(), b.as_number());
}

// ---------------------------------------------------------------------------
// SQRTPI
// ---------------------------------------------------------------------------

TEST(BuiltinsMath4SqrtPi, OneIsSqrtPi) {
  const Value v = EvalSource("=SQRTPI(1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), std::sqrt(kPi), 1e-12);
}

TEST(BuiltinsMath4SqrtPi, ZeroIsZero) {
  const Value v = EvalSource("=SQRTPI(0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsMath4SqrtPi, FourMatches2SqrtPi) {
  const Value v = EvalSource("=SQRTPI(4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 2.0 * std::sqrt(kPi), 1e-12);
}

TEST(BuiltinsMath4SqrtPi, NegativeYieldsNum) {
  const Value v = EvalSource("=SQRTPI(-1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// Registration sanity: every new name must be reachable via the registry.
// ---------------------------------------------------------------------------

TEST(BuiltinsMath4Registration, AllNamesRegistered) {
  const FunctionRegistry& reg = default_registry();
  for (const char* name : {"FACT", "FACTDOUBLE", "COMBIN", "COMBINA", "MULTINOMIAL", "GCD", "LCM", "ARABIC", "ROMAN",
                           "BASE", "DECIMAL", "CEILING.PRECISE", "FLOOR.PRECISE", "ISO.CEILING", "SQRTPI"}) {
    EXPECT_NE(reg.lookup(name), nullptr) << "not registered: " << name;
  }
}

// FACT snaps up within 2^-22: 10.9999997616 is inside, 10.999999761 is not.
TEST(BuiltinsMath4Fact, NearIntegerSnapBoundary) {
  const Value inside = EvalSource("=FACT(10.9999997616)");
  ASSERT_TRUE(inside.is_number());
  EXPECT_EQ(inside.as_number(), 39916800.0);
  const Value outside = EvalSource("=FACT(10.999999761)");
  ASSERT_TRUE(outside.is_number());
  EXPECT_EQ(outside.as_number(), 3628800.0);
}

TEST(BuiltinsMath4Fact, NegativeRawValueIsNumBeforeSnap) {
  const Value v = EvalSource("=FACT(-0.0000001)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// A char subtracts when it is smaller than the largest value to its right.
TEST(BuiltinsMath4Arabic, SubtractsAgainstLargestValueToTheRight) {
  const struct {
    const char* src;
    double expected;
  } cases[] = {
      {"=ARABIC(\"IIM\")", 998.0}, {"=ARABIC(\"IM\")", 999.0}, {"=ARABIC(\"IIX\")", 8.0},
      {"=ARABIC(\"IIC\")", 98.0},  {"=ARABIC(\"XXC\")", 80.0}, {"=ARABIC(\"VX\")", 5.0},
  };
  for (const auto& c : cases) {
    const Value v = EvalSource(c.src);
    ASSERT_TRUE(v.is_number()) << c.src;
    EXPECT_EQ(v.as_number(), c.expected) << c.src;
  }
}

// A boolean read from a range cell is #VALUE! for GCD / LCM / MULTINOMIAL.
TEST(BuiltinsMath4BoolRangeCell, GcdLcmMultinomialAreValue) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::boolean(true));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  for (const char* formula : {"=GCD(A1:A2)", "=LCM(A1:A2)", "=MULTINOMIAL(A1:A2)"}) {
    const Value v = EvalSourceIn(formula, wb, wb.sheet(0));
    ASSERT_TRUE(v.is_error()) << formula;
    EXPECT_EQ(v.as_error(), ErrorCode::Value) << formula;
  }
}

}  // namespace
}  // namespace eval
}  // namespace formulon
