// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for the simple integer-only engineering built-ins:
//   * Base conversion (BIN/OCT/HEX <-> DEC): happy paths, zero, maxima,
//     two's complement negatives, `places` padding (positive-only), and
//     the full battery of input-validation errors (#NUM! on invalid digit,
//     oversize input, out-of-signed-range).
//   * Bit operations (BITAND / BITOR / BITXOR / BITLSHIFT / BITRSHIFT):
//     canonical cases, negative-shift semantics, overflow-to-#NUM!, and
//     shift-magnitude cap.
//   * Comparators (DELTA, GESTEP): happy paths including the default
//     second argument (zero), plus #VALUE! propagation for non-numeric
//     input.

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

TEST(BuiltinsEngineeringRegistry, NamesRegistered) {
  const FunctionRegistry& reg = default_registry();
  EXPECT_NE(reg.lookup("BIN2DEC"), nullptr);
  EXPECT_NE(reg.lookup("BIN2OCT"), nullptr);
  EXPECT_NE(reg.lookup("BIN2HEX"), nullptr);
  EXPECT_NE(reg.lookup("OCT2DEC"), nullptr);
  EXPECT_NE(reg.lookup("OCT2BIN"), nullptr);
  EXPECT_NE(reg.lookup("OCT2HEX"), nullptr);
  EXPECT_NE(reg.lookup("HEX2DEC"), nullptr);
  EXPECT_NE(reg.lookup("HEX2BIN"), nullptr);
  EXPECT_NE(reg.lookup("HEX2OCT"), nullptr);
  EXPECT_NE(reg.lookup("DEC2BIN"), nullptr);
  EXPECT_NE(reg.lookup("DEC2OCT"), nullptr);
  EXPECT_NE(reg.lookup("DEC2HEX"), nullptr);
  EXPECT_NE(reg.lookup("BITAND"), nullptr);
  EXPECT_NE(reg.lookup("BITOR"), nullptr);
  EXPECT_NE(reg.lookup("BITXOR"), nullptr);
  EXPECT_NE(reg.lookup("BITLSHIFT"), nullptr);
  EXPECT_NE(reg.lookup("BITRSHIFT"), nullptr);
  EXPECT_NE(reg.lookup("DELTA"), nullptr);
  EXPECT_NE(reg.lookup("GESTEP"), nullptr);
}

// ---------------------------------------------------------------------------
// *2DEC positive happy paths
// ---------------------------------------------------------------------------

TEST(BuiltinsEngineering, Bin2DecPositive) {
  const Value v = EvalSource("=BIN2DEC(\"1010\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 10.0);
}

TEST(BuiltinsEngineering, Bin2DecZero) {
  const Value v = EvalSource("=BIN2DEC(\"0\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsEngineering, Bin2DecMaxPositive) {
  const Value v = EvalSource("=BIN2DEC(\"111111111\")");  // 9 ones -> 511
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 511.0);
}

TEST(BuiltinsEngineering, Bin2DecTwosComplementNegative) {
  // 10 digits with MSB=1 -> negative. 1111111110 (1022) - 1024 = -2.
  const Value v = EvalSource("=BIN2DEC(\"1111111110\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -2.0);
}

TEST(BuiltinsEngineering, Oct2DecPositive) {
  const Value v = EvalSource("=OCT2DEC(\"17\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 15.0);
}

TEST(BuiltinsEngineering, Oct2DecNegative) {
  // Top digit >= 4 at width 10 -> two's complement. "7777777777" -> -1.
  const Value v = EvalSource("=OCT2DEC(\"7777777777\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -1.0);
}

TEST(BuiltinsEngineering, Hex2DecPositive) {
  const Value v = EvalSource("=HEX2DEC(\"FF\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 255.0);
}

TEST(BuiltinsEngineering, Hex2DecLowercaseAccepted) {
  // Excel accepts lowercase hex input.
  const Value v = EvalSource("=HEX2DEC(\"ff\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 255.0);
}

TEST(BuiltinsEngineering, Hex2DecFullNegative) {
  // 10-digit hex with MSB>=8 is negative. FFFFFFFFFF -> -1.
  const Value v = EvalSource("=HEX2DEC(\"FFFFFFFFFF\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -1.0);
}

// ---------------------------------------------------------------------------
// DEC2* positive + negative
// ---------------------------------------------------------------------------

TEST(BuiltinsEngineering, Dec2BinPositive) {
  const Value v = EvalSource("=DEC2BIN(10)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "1010");
}

TEST(BuiltinsEngineering, Dec2BinZero) {
  const Value v = EvalSource("=DEC2BIN(0)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "0");
}

TEST(BuiltinsEngineering, Dec2BinMaxPositive) {
  const Value v = EvalSource("=DEC2BIN(511)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "111111111");
}

TEST(BuiltinsEngineering, Dec2BinNegative) {
  // -2 -> two's complement 10-digit: 1111111110.
  const Value v = EvalSource("=DEC2BIN(-2)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "1111111110");
}

TEST(BuiltinsEngineering, Dec2BinMinNegative) {
  // -512 is the signed minimum for a 10-digit binary two's complement.
  const Value v = EvalSource("=DEC2BIN(-512)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "1000000000");
}

TEST(BuiltinsEngineering, Dec2HexMinusOne) {
  const Value v = EvalSource("=DEC2HEX(-1)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "FFFFFFFFFF");
}

TEST(BuiltinsEngineering, Dec2HexPositiveUppercase) {
  // Hex output must be uppercase.
  const Value v = EvalSource("=DEC2HEX(255)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "FF");
}

TEST(BuiltinsEngineering, Dec2OctPositive) {
  const Value v = EvalSource("=DEC2OCT(15)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "17");
}

TEST(BuiltinsEngineering, Dec2OctNegative) {
  const Value v = EvalSource("=DEC2OCT(-1)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "7777777777");
}

TEST(BuiltinsEngineering, Dec2BinBoolInputIsValue) {
  // Mac Excel 365 rejects a direct Bool argument to DEC2BIN / DEC2OCT /
  // DEC2HEX with `#VALUE!` instead of coercing TRUE/FALSE to 1/0 as the
  // general numeric rule would. Matches the EFFECT / NOMINAL strict-Bool
  // quirk.
  const Value v = EvalSource("=DEC2BIN(TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsEngineering, Dec2OctBoolInputIsValue) {
  const Value v = EvalSource("=DEC2OCT(TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsEngineering, Dec2HexBoolInputIsValue) {
  const Value v = EvalSource("=DEC2HEX(FALSE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// `places` padding
// ---------------------------------------------------------------------------

TEST(BuiltinsEngineering, Dec2BinPlacesPads) {
  const Value v = EvalSource("=DEC2BIN(2,4)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "0010");
}

TEST(BuiltinsEngineering, Dec2BinPlacesTooSmall) {
  // 2 binary = "10" (2 digits); places=1 is insufficient -> #NUM!.
  const Value v = EvalSource("=DEC2BIN(2,1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsEngineering, Dec2BinPlacesIgnoredForNegative) {
  // `places` is ignored when the result is negative; always 10 digits.
  const Value v = EvalSource("=DEC2BIN(-2,4)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "1111111110");
}

TEST(BuiltinsEngineering, Dec2HexPlacesPadsUppercase) {
  const Value v = EvalSource("=DEC2HEX(10,4)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "000A");
}

TEST(BuiltinsEngineering, Dec2BinPlacesZeroRejected) {
  // places must be in [1, 10].
  const Value v = EvalSource("=DEC2BIN(1,0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsEngineering, Dec2BinPlacesTooLarge) {
  const Value v = EvalSource("=DEC2BIN(1,11)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// Range / input validation
// ---------------------------------------------------------------------------

TEST(BuiltinsEngineering, Dec2BinOutOfRange) {
  const Value v = EvalSource("=DEC2BIN(512)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsEngineering, Dec2BinBelowRange) {
  const Value v = EvalSource("=DEC2BIN(-513)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsEngineering, Dec2HexOutOfRange) {
  // 0x10000000000 = 2^40 is one past the signed maximum for 10-digit hex.
  const Value v = EvalSource("=DEC2HEX(1099511627776)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsEngineering, Bin2DecInvalidDigit) {
  const Value v = EvalSource("=BIN2DEC(\"102\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsEngineering, Hex2DecInvalidDigit) {
  const Value v = EvalSource("=HEX2DEC(\"XYZ\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsEngineering, Bin2DecTooManyDigits) {
  // 11 digits exceeds the 10-digit two's complement window.
  const Value v = EvalSource("=BIN2DEC(\"11111111111\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsEngineering, Bin2DecEmptyString) {
  // Empty text is rejected per the spec.
  const Value v = EvalSource("=BIN2DEC(\"\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// Cross-base conversions (spot check)
// ---------------------------------------------------------------------------

TEST(BuiltinsEngineering, Bin2OctPositive) {
  const Value v = EvalSource("=BIN2OCT(\"1010\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "12");
}

TEST(BuiltinsEngineering, Bin2HexPositive) {
  const Value v = EvalSource("=BIN2HEX(\"1111\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "F");
}

TEST(BuiltinsEngineering, Oct2BinNegativeRoundtrip) {
  // "7777777777" oct -> -1 dec -> "1111111111" bin (10-digit two's complement).
  const Value v = EvalSource("=OCT2BIN(\"7777777777\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "1111111111");
}

TEST(BuiltinsEngineering, Oct2HexPositive) {
  const Value v = EvalSource("=OCT2HEX(\"17\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "F");
}

TEST(BuiltinsEngineering, Hex2BinSmallPositive) {
  const Value v = EvalSource("=HEX2BIN(\"A\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "1010");
}

TEST(BuiltinsEngineering, Hex2OctPositive) {
  const Value v = EvalSource("=HEX2OCT(\"FF\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "377");
}

TEST(BuiltinsEngineering, Hex2BinNegativeOutOfRange) {
  // -512 fits in BIN (min); -513 does not. HEX "FFFFFFFE00" = -512 -> ok.
  const Value v = EvalSource("=HEX2BIN(\"FFFFFFFE00\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "1000000000");
}

TEST(BuiltinsEngineering, Hex2BinPositiveOutOfRange) {
  // 0x200 = 512, one past the BIN signed maximum.
  const Value v = EvalSource("=HEX2BIN(\"200\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// Bit operations
// ---------------------------------------------------------------------------

TEST(BuiltinsEngineering, BitAndBasic) {
  const Value v = EvalSource("=BITAND(5,3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsEngineering, BitOrBasic) {
  const Value v = EvalSource("=BITOR(5,3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 7.0);
}

TEST(BuiltinsEngineering, BitXorBasic) {
  const Value v = EvalSource("=BITXOR(5,3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 6.0);
}

TEST(BuiltinsEngineering, BitLShiftBasic) {
  const Value v = EvalSource("=BITLSHIFT(1,8)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 256.0);
}

TEST(BuiltinsEngineering, BitRShiftBasic) {
  const Value v = EvalSource("=BITRSHIFT(256,4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 16.0);
}

TEST(BuiltinsEngineering, BitLShiftNegativeShiftRightward) {
  // BITLSHIFT(n, -k) == BITRSHIFT(n, k).
  const Value v = EvalSource("=BITLSHIFT(8,-2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsEngineering, BitRShiftNegativeShiftLeftward) {
  const Value v = EvalSource("=BITRSHIFT(8,-2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 32.0);
}

TEST(BuiltinsEngineering, BitLShiftBoundary47Ok) {
  // 1 << 47 = 2^47, still fits in 48 bits.
  const Value v = EvalSource("=BITLSHIFT(1,47)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), static_cast<double>(1LL << 47));
}

TEST(BuiltinsEngineering, BitLShiftOverflow48) {
  // 1 << 48 = 2^48 exceeds the 48-bit window.
  const Value v = EvalSource("=BITLSHIFT(1,48)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsEngineering, BitAndNegativeInputRejected) {
  const Value v = EvalSource("=BITAND(-1,1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsEngineering, BitAndOverflow48BitInput) {
  // 2^48 exceeds the allowed input range.
  const Value v = EvalSource("=BITAND(281474976710656,1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsEngineering, BitLShiftMagnitudeCap) {
  // Shift magnitude > 53 -> #NUM!.
  const Value v = EvalSource("=BITLSHIFT(1,54)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsEngineering, BitRShiftMagnitudeCapNegative) {
  const Value v = EvalSource("=BITRSHIFT(1,-54)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsEngineering, BitAndMaxInput) {
  // 2^48 - 1 is the maximum allowed input.
  const Value v = EvalSource("=BITAND(281474976710655,281474976710655)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 281474976710655.0);
}

// ---------------------------------------------------------------------------
// DELTA
// ---------------------------------------------------------------------------

TEST(BuiltinsEngineering, DeltaEqual) {
  const Value v = EvalSource("=DELTA(5,5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsEngineering, DeltaUnequal) {
  const Value v = EvalSource("=DELTA(5,4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsEngineering, DeltaDefaultSecondArgZero) {
  // DELTA(3) -> compare 3 with 0 -> 0.
  const Value v = EvalSource("=DELTA(3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsEngineering, DeltaDefaultZeroEqual) {
  const Value v = EvalSource("=DELTA(0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsEngineering, DeltaNonNumericReturnsValue) {
  const Value v = EvalSource("=DELTA(\"a\",1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// GESTEP
// ---------------------------------------------------------------------------

TEST(BuiltinsEngineering, GestepAboveStep) {
  const Value v = EvalSource("=GESTEP(5,4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsEngineering, GestepBelowStep) {
  const Value v = EvalSource("=GESTEP(4,5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsEngineering, GestepEqualStep) {
  // Equality counts as ">=".
  const Value v = EvalSource("=GESTEP(5,5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsEngineering, GestepDefaultStepZeroAbove) {
  const Value v = EvalSource("=GESTEP(5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsEngineering, GestepDefaultStepZeroBelow) {
  const Value v = EvalSource("=GESTEP(-1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsEngineering, GestepNonNumericReturnsValue) {
  const Value v = EvalSource("=GESTEP(\"a\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
