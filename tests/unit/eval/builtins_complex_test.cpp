// Copyright 2026 libraz. Licensed under the MIT License.
//
// Tests for Excel's complex-number built-ins: COMPLEX + the 24 IM* family.
//
// Coverage:
//   * Registry pin.
//   * COMPLEX formatting rules (pure real, pure imag, +/-1 coef, full form,
//     alternate 'j' suffix, invalid suffix).
//   * parse_complex / format_complex round-trip via COMPLEX(IMREAL(z),
//     IMAGINARY(z)) for several canonical shapes.
//   * Each IM* function: happy path, error code, and one precision edge.
//   * Excel quirks: Blank -> 0, Bool -> 0/1, Number treated as pure real.

#include <cmath>
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
// Registry pin
// ---------------------------------------------------------------------------

TEST(BuiltinsComplexRegistry, NamesRegistered) {
  const FunctionRegistry& reg = default_registry();
  EXPECT_NE(reg.lookup("COMPLEX"), nullptr);
  EXPECT_NE(reg.lookup("IMABS"), nullptr);
  EXPECT_NE(reg.lookup("IMAGINARY"), nullptr);
  EXPECT_NE(reg.lookup("IMREAL"), nullptr);
  EXPECT_NE(reg.lookup("IMCONJUGATE"), nullptr);
  EXPECT_NE(reg.lookup("IMARGUMENT"), nullptr);
  EXPECT_NE(reg.lookup("IMSUM"), nullptr);
  EXPECT_NE(reg.lookup("IMSUB"), nullptr);
  EXPECT_NE(reg.lookup("IMPRODUCT"), nullptr);
  EXPECT_NE(reg.lookup("IMDIV"), nullptr);
  EXPECT_NE(reg.lookup("IMPOWER"), nullptr);
  EXPECT_NE(reg.lookup("IMEXP"), nullptr);
  EXPECT_NE(reg.lookup("IMLN"), nullptr);
  EXPECT_NE(reg.lookup("IMLOG10"), nullptr);
  EXPECT_NE(reg.lookup("IMLOG2"), nullptr);
  EXPECT_NE(reg.lookup("IMSQRT"), nullptr);
  EXPECT_NE(reg.lookup("IMSIN"), nullptr);
  EXPECT_NE(reg.lookup("IMCOS"), nullptr);
  EXPECT_NE(reg.lookup("IMTAN"), nullptr);
  EXPECT_NE(reg.lookup("IMSEC"), nullptr);
  EXPECT_NE(reg.lookup("IMCSC"), nullptr);
  EXPECT_NE(reg.lookup("IMCOT"), nullptr);
  EXPECT_NE(reg.lookup("IMSINH"), nullptr);
  EXPECT_NE(reg.lookup("IMCOSH"), nullptr);
  EXPECT_NE(reg.lookup("IMSECH"), nullptr);
  EXPECT_NE(reg.lookup("IMCSCH"), nullptr);
}

// ---------------------------------------------------------------------------
// COMPLEX formatting table (spec §§ "Formatting")
// ---------------------------------------------------------------------------

TEST(BuiltinsComplex, FormatFullDefault) {
  const Value v = EvalSource("=COMPLEX(3,4)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "3+4i");
}

TEST(BuiltinsComplex, FormatFullJSuffix) {
  const Value v = EvalSource("=COMPLEX(3,4,\"j\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "3+4j");
}

TEST(BuiltinsComplex, FormatPureImagPlusOne) {
  const Value v = EvalSource("=COMPLEX(0,1)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "i");
}

TEST(BuiltinsComplex, FormatPureImagMinusOne) {
  const Value v = EvalSource("=COMPLEX(0,-1)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "-i");
}

TEST(BuiltinsComplex, FormatPureImagFive) {
  const Value v = EvalSource("=COMPLEX(0,5)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "5i");
}

TEST(BuiltinsComplex, FormatPureReal) {
  const Value v = EvalSource("=COMPLEX(3,0)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "3");
}

TEST(BuiltinsComplex, FormatFullPlusOneCoefDropped) {
  const Value v = EvalSource("=COMPLEX(3,1)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "3+i");
}

TEST(BuiltinsComplex, FormatFullMinusOneCoefDropped) {
  const Value v = EvalSource("=COMPLEX(3,-1)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "3-i");
}

TEST(BuiltinsComplex, FormatRejectsInvalidSuffix) {
  const Value v = EvalSource("=COMPLEX(3,4,\"x\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsComplex, FormatRejectsUppercaseSuffix) {
  const Value v = EvalSource("=COMPLEX(3,4,\"I\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsComplex, FormatNegativeRealFullForm) {
  const Value v = EvalSource("=COMPLEX(-3,4)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "-3+4i");
}

TEST(BuiltinsComplex, FormatNegativeRealNegativeImag) {
  const Value v = EvalSource("=COMPLEX(-3,-4)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "-3-4i");
}

TEST(BuiltinsComplex, FormatNegativeImagOne) {
  const Value v = EvalSource("=COMPLEX(2,-1)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "2-i");
}

// ---------------------------------------------------------------------------
// Roundtrip: parse(format(r, i, s)) == (r, i, s)
// ---------------------------------------------------------------------------

// Re-extract (re, im) from a freshly-formatted COMPLEX text by feeding it to
// IMREAL / IMAGINARY. This exercises parse_complex end-to-end.
void ExpectRoundtrip(const char* formula_real, const char* formula_imag, double expected_re, double expected_im) {
  const Value re = EvalSource(formula_real);
  const Value im = EvalSource(formula_imag);
  ASSERT_TRUE(re.is_number()) << formula_real;
  ASSERT_TRUE(im.is_number()) << formula_imag;
  EXPECT_DOUBLE_EQ(re.as_number(), expected_re);
  EXPECT_DOUBLE_EQ(im.as_number(), expected_im);
}

TEST(BuiltinsComplex, RoundtripZero) {
  ExpectRoundtrip("=IMREAL(COMPLEX(0,0))", "=IMAGINARY(COMPLEX(0,0))", 0.0, 0.0);
}

TEST(BuiltinsComplex, RoundtripPureRealOne) {
  ExpectRoundtrip("=IMREAL(COMPLEX(1,0))", "=IMAGINARY(COMPLEX(1,0))", 1.0, 0.0);
}

TEST(BuiltinsComplex, RoundtripPureImagOne) {
  ExpectRoundtrip("=IMREAL(COMPLEX(0,1))", "=IMAGINARY(COMPLEX(0,1))", 0.0, 1.0);
}

TEST(BuiltinsComplex, RoundtripPureImagMinusOne) {
  ExpectRoundtrip("=IMREAL(COMPLEX(0,-1))", "=IMAGINARY(COMPLEX(0,-1))", 0.0, -1.0);
}

TEST(BuiltinsComplex, RoundtripThreeFour) {
  ExpectRoundtrip("=IMREAL(COMPLEX(3,4))", "=IMAGINARY(COMPLEX(3,4))", 3.0, 4.0);
}

TEST(BuiltinsComplex, RoundtripNegativeTinyImag) {
  // Small magnitude preserved; negative real too.
  ExpectRoundtrip("=IMREAL(COMPLEX(-2.5,0.000001))", "=IMAGINARY(COMPLEX(-2.5,0.000001))", -2.5, 0.000001);
}

TEST(BuiltinsComplex, RoundtripJSuffixPreserved) {
  // Suffix rides on text; parse restores it into IMAGINARY (which drops it).
  const Value t = EvalSource("=COMPLEX(2,3,\"j\")");
  ASSERT_TRUE(t.is_text());
  EXPECT_EQ(t.as_text(), "2+3j");
}

// ---------------------------------------------------------------------------
// Inspectors: IMABS / IMREAL / IMAGINARY / IMCONJUGATE / IMARGUMENT
// ---------------------------------------------------------------------------

TEST(BuiltinsComplex, ImAbsThreeFour) {
  const Value v = EvalSource("=IMABS(\"3+4i\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.0);
}

TEST(BuiltinsComplex, ImAbsNegativeComponents) {
  const Value v = EvalSource("=IMABS(\"-3-4i\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.0);
}

TEST(BuiltinsComplex, ImAbsPureImag) {
  const Value v = EvalSource("=IMABS(\"5i\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.0);
}

TEST(BuiltinsComplex, ImAbsInvalidIsNum) {
  const Value v = EvalSource("=IMABS(\"not a complex\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsComplex, ImAbsTrueIsOne) {
  const Value v = EvalSource("=IMABS(TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsComplex, ImAbsNumberIsMagnitude) {
  const Value v = EvalSource("=IMABS(5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.0);
}

TEST(BuiltinsComplex, ImRealExtractsRe) {
  const Value v = EvalSource("=IMREAL(\"3+4i\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsComplex, ImRealOfNumber) {
  const Value v = EvalSource("=IMREAL(5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.0);
}

TEST(BuiltinsComplex, ImRealOfPureImag) {
  const Value v = EvalSource("=IMREAL(\"5i\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsComplex, ImAginaryExtractsIm) {
  const Value v = EvalSource("=IMAGINARY(\"3+4i\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 4.0);
}

TEST(BuiltinsComplex, ImAginaryOfBareI) {
  const Value v = EvalSource("=IMAGINARY(\"i\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsComplex, ImAginaryOfMinusI) {
  const Value v = EvalSource("=IMAGINARY(\"-i\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), -1.0);
}

TEST(BuiltinsComplex, ImAginaryOfNumber) {
  const Value v = EvalSource("=IMAGINARY(5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsComplex, ImConjugate) {
  const Value v = EvalSource("=IMCONJUGATE(\"3+4i\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "3-4i");
}

TEST(BuiltinsComplex, ImConjugatePreservesJ) {
  const Value v = EvalSource("=IMCONJUGATE(\"3+4j\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "3-4j");
}

TEST(BuiltinsComplex, ImConjugateOfRealNoImag) {
  const Value v = EvalSource("=IMCONJUGATE(\"5\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "5");
}

TEST(BuiltinsComplex, ImArgument) {
  const Value v = EvalSource("=IMARGUMENT(\"3+4i\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), std::atan2(4.0, 3.0), 1e-12);
}

TEST(BuiltinsComplex, ImArgumentOfZeroIsDiv0) {
  const Value v = EvalSource("=IMARGUMENT(\"0\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// Arithmetic: IMSUM / IMSUB / IMPRODUCT / IMDIV / IMPOWER
// ---------------------------------------------------------------------------

TEST(BuiltinsComplex, ImSumBasic) {
  const Value v = EvalSource("=IMSUM(\"1+i\",\"2+3i\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "3+4i");
}

TEST(BuiltinsComplex, ImSumVariadic) {
  const Value v = EvalSource("=IMSUM(\"1+i\",\"2+i\",\"3+i\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "6+3i");
}

TEST(BuiltinsComplex, ImSumMixedSuffixIsValue) {
  const Value v = EvalSource("=IMSUM(\"1+i\",\"2+3j\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsComplex, ImSumPureRealsKeepFirstSuffix) {
  // Pure real (no imag) must not poison the suffix of subsequent args.
  const Value v = EvalSource("=IMSUM(\"2+3j\",\"5\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "7+3j");
}

TEST(BuiltinsComplex, ImSub) {
  const Value v = EvalSource("=IMSUB(\"5+3i\",\"2+i\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "3+2i");
}

TEST(BuiltinsComplex, ImSubNegativeImagResult) {
  const Value v = EvalSource("=IMSUB(\"2+i\",\"5+3i\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "-3-2i");
}

TEST(BuiltinsComplex, ImProductBasic) {
  // (2+i)(3+4i) = 6 + 8i + 3i - 4 = 2 + 11i.
  const Value v = EvalSource("=IMPRODUCT(\"2+i\",\"3+4i\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "2+11i");
}

TEST(BuiltinsComplex, ImProductChained) {
  // (1+i)*(1+i)*(1+i) = (2i)*(1+i) = 2i + 2i^2 = -2 + 2i.
  const Value v = EvalSource("=IMPRODUCT(\"1+i\",\"1+i\",\"1+i\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "-2+2i");
}

TEST(BuiltinsComplex, ImProductMixedSuffix) {
  const Value v = EvalSource("=IMPRODUCT(\"1+i\",\"2+3j\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsComplex, ImDivBasic) {
  // (1+i)/(1-i) = (1+i)(1+i)/2 = (1+2i-1)/2 = i.
  const Value v = EvalSource("=IMDIV(\"1+i\",\"1-i\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "i");
}

TEST(BuiltinsComplex, ImDivByZeroIsNum) {
  const Value v = EvalSource("=IMDIV(\"1+i\",\"0\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsComplex, ImDivMixedSuffix) {
  const Value v = EvalSource("=IMDIV(\"1+i\",\"1-j\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsComplex, ImPowerSquare) {
  // (1+i)^2 mathematically equals 2i, but IMPOWER goes through the polar
  // form (r=sqrt(2), theta=pi/4 -> doubled), and cos(pi/2) is not exactly
  // zero in IEEE 754. Mac Excel 365 surfaces the residue rather than
  // snapping it to zero, so the formatted text carries a ~1.22e-16 real
  // part. 1-bit parity (per CLAUDE.md) requires the same output.
  const Value v = EvalSource("=IMPOWER(\"1+i\",2)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "1.22464679914735E-16+2i");
}

TEST(BuiltinsComplex, ImPowerInteger) {
  // (2+0i)^3 = 8.
  const Value v = EvalSource("=IMPOWER(\"2\",3)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "8");
}

TEST(BuiltinsComplex, ImPowerZeroBaseNonPositiveExpIsNum) {
  const Value v = EvalSource("=IMPOWER(\"0\",0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsComplex, ImPowerZeroBaseNegExpIsNum) {
  const Value v = EvalSource("=IMPOWER(\"0\",-1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// Exponentials / logs / roots
// ---------------------------------------------------------------------------

TEST(BuiltinsComplex, ImExp) {
  // e^(1+i) = e*(cos 1 + i sin 1). Tolerance reflects format_double's
  // 6-significant-digit roundtrip through text.
  const double expected_re = std::exp(1.0) * std::cos(1.0);
  const double expected_im = std::exp(1.0) * std::sin(1.0);
  const Value re = EvalSource("=IMREAL(IMEXP(\"1+i\"))");
  const Value im = EvalSource("=IMAGINARY(IMEXP(\"1+i\"))");
  ASSERT_TRUE(re.is_number());
  ASSERT_TRUE(im.is_number());
  EXPECT_NEAR(re.as_number(), expected_re, 1e-5);
  EXPECT_NEAR(im.as_number(), expected_im, 1e-5);
}

TEST(BuiltinsComplex, ImExpOfZero) {
  const Value v = EvalSource("=IMEXP(\"0\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "1");
}

TEST(BuiltinsComplex, ImLn) {
  // ln(1+i) = 0.5*ln(2) + i*pi/4. Tolerance accommodates the 6-digit
  // roundtrip through format_double.
  const Value re = EvalSource("=IMREAL(IMLN(\"1+i\"))");
  const Value im = EvalSource("=IMAGINARY(IMLN(\"1+i\"))");
  ASSERT_TRUE(re.is_number());
  ASSERT_TRUE(im.is_number());
  EXPECT_NEAR(re.as_number(), 0.5 * std::log(2.0), 1e-5);
  EXPECT_NEAR(im.as_number(), std::atan2(1.0, 1.0), 1e-5);
}

TEST(BuiltinsComplex, ImLnOfZeroIsNum) {
  const Value v = EvalSource("=IMLN(\"0\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsComplex, ImLog10) {
  const Value re = EvalSource("=IMREAL(IMLOG10(\"1+i\"))");
  const Value im = EvalSource("=IMAGINARY(IMLOG10(\"1+i\"))");
  ASSERT_TRUE(re.is_number());
  ASSERT_TRUE(im.is_number());
  const double inv = 1.0 / std::log(10.0);
  EXPECT_NEAR(re.as_number(), 0.5 * std::log(2.0) * inv, 1e-5);
  EXPECT_NEAR(im.as_number(), std::atan2(1.0, 1.0) * inv, 1e-5);
}

TEST(BuiltinsComplex, ImLog2) {
  const Value re = EvalSource("=IMREAL(IMLOG2(\"1+i\"))");
  const Value im = EvalSource("=IMAGINARY(IMLOG2(\"1+i\"))");
  ASSERT_TRUE(re.is_number());
  ASSERT_TRUE(im.is_number());
  const double inv = 1.0 / std::log(2.0);
  EXPECT_NEAR(re.as_number(), 0.5 * std::log(2.0) * inv, 1e-5);
  EXPECT_NEAR(im.as_number(), std::atan2(1.0, 1.0) * inv, 1e-5);
}

TEST(BuiltinsComplex, ImSqrtOfI) {
  // sqrt(i) = (1+i)/sqrt(2).
  const Value re = EvalSource("=IMREAL(IMSQRT(\"i\"))");
  const Value im = EvalSource("=IMAGINARY(IMSQRT(\"i\"))");
  ASSERT_TRUE(re.is_number());
  ASSERT_TRUE(im.is_number());
  EXPECT_NEAR(re.as_number(), 1.0 / std::sqrt(2.0), 1e-5);
  EXPECT_NEAR(im.as_number(), 1.0 / std::sqrt(2.0), 1e-5);
}

TEST(BuiltinsComplex, ImSqrtOfFour) {
  const Value v = EvalSource("=IMSQRT(\"4\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "2");
}

// ---------------------------------------------------------------------------
// Trigonometric
// ---------------------------------------------------------------------------

TEST(BuiltinsComplex, ImSinOfRealZero) {
  const Value v = EvalSource("=IMSIN(\"0\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "0");
}

TEST(BuiltinsComplex, ImSinOfPureImag) {
  // sin(i) = i*sinh(1).
  const Value re = EvalSource("=IMREAL(IMSIN(\"i\"))");
  const Value im = EvalSource("=IMAGINARY(IMSIN(\"i\"))");
  ASSERT_TRUE(re.is_number());
  ASSERT_TRUE(im.is_number());
  EXPECT_NEAR(re.as_number(), 0.0, 1e-12);
  EXPECT_NEAR(im.as_number(), std::sinh(1.0), 1e-5);
}

TEST(BuiltinsComplex, ImCosOfRealZero) {
  const Value v = EvalSource("=IMCOS(\"0\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "1");
}

TEST(BuiltinsComplex, ImCosOfPureImag) {
  // cos(i) = cosh(1).
  const Value v = EvalSource("=IMREAL(IMCOS(\"i\"))");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), std::cosh(1.0), 1e-5);
}

TEST(BuiltinsComplex, ImTanOfZero) {
  const Value v = EvalSource("=IMTAN(\"0\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "0");
}

TEST(BuiltinsComplex, ImSecOfZero) {
  const Value v = EvalSource("=IMSEC(\"0\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "1");
}

TEST(BuiltinsComplex, ImCscOfI) {
  // 1/sin(i) = 1/(i*sinh(1)) = -i/sinh(1).
  const Value re = EvalSource("=IMREAL(IMCSC(\"i\"))");
  const Value im = EvalSource("=IMAGINARY(IMCSC(\"i\"))");
  ASSERT_TRUE(re.is_number());
  ASSERT_TRUE(im.is_number());
  EXPECT_NEAR(re.as_number(), 0.0, 1e-12);
  EXPECT_NEAR(im.as_number(), -1.0 / std::sinh(1.0), 1e-5);
}

TEST(BuiltinsComplex, ImCotOfI) {
  // cot(i) = cos(i)/sin(i) = cosh(1)/(i*sinh(1)) = -i*coth(1).
  const Value im = EvalSource("=IMAGINARY(IMCOT(\"i\"))");
  ASSERT_TRUE(im.is_number());
  EXPECT_NEAR(im.as_number(), -std::cosh(1.0) / std::sinh(1.0), 1e-5);
}

// ---------------------------------------------------------------------------
// Hyperbolic
// ---------------------------------------------------------------------------

TEST(BuiltinsComplex, ImSinhOfZero) {
  const Value v = EvalSource("=IMSINH(\"0\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "0");
}

TEST(BuiltinsComplex, ImCoshOfZero) {
  const Value v = EvalSource("=IMCOSH(\"0\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "1");
}

TEST(BuiltinsComplex, ImSinhOfReal) {
  // sinh(1) is real.
  const Value v = EvalSource("=IMREAL(IMSINH(\"1\"))");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), std::sinh(1.0), 1e-5);
}

TEST(BuiltinsComplex, ImCoshOfPureImagReal) {
  // cosh(i) = cos(1); real-only result.
  const Value v = EvalSource("=IMREAL(IMCOSH(\"i\"))");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), std::cos(1.0), 1e-5);
}

TEST(BuiltinsComplex, ImSechOfZero) {
  const Value v = EvalSource("=IMSECH(\"0\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "1");
}

TEST(BuiltinsComplex, ImCschOfI) {
  // 1/sinh(i) = 1/(i*sin(1)) = -i/sin(1).
  const Value im = EvalSource("=IMAGINARY(IMCSCH(\"i\"))");
  ASSERT_TRUE(im.is_number());
  EXPECT_NEAR(im.as_number(), -1.0 / std::sin(1.0), 1e-5);
}

// ---------------------------------------------------------------------------
// Parse edge cases
// ---------------------------------------------------------------------------

TEST(BuiltinsComplex, ParseScientificNotation) {
  const Value v = EvalSource("=IMREAL(\"1e2+3i\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 100.0);
}

TEST(BuiltinsComplex, ParseScientificInImag) {
  const Value v = EvalSource("=IMAGINARY(\"1+2e-3i\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.002);
}

TEST(BuiltinsComplex, ParseLeadingPlusReal) {
  const Value v = EvalSource("=IMREAL(\"+3+4i\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsComplex, ParseEmptyIsNum) {
  const Value v = EvalSource("=IMABS(\"\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsComplex, ParseDoubleSignIsNum) {
  const Value v = EvalSource("=IMABS(\"3+-4i\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsComplex, ParseMixedSuffixIsNum) {
  const Value v = EvalSource("=IMABS(\"i+j\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsComplex, ParseUppercaseISufIsNum) {
  const Value v = EvalSource("=IMABS(\"3+4I\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// Error propagation
// ---------------------------------------------------------------------------

TEST(BuiltinsComplex, ErrorPropagatesThroughIMABS) {
  const Value v = EvalSource("=IMABS(#N/A)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsComplex, ErrorPropagatesThroughIMSUM) {
  const Value v = EvalSource("=IMSUM(\"1+i\",#DIV/0!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
