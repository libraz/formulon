// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for the BAHTTEXT builtin: spells out a number as Thai-baht
// text. Covers each of the special Thai reading rules (the `เอ็ด` ones digit
// after a non-zero higher digit, the `ยี่` two-digit, the `สิบ` collapse for
// the tens digit `1`), the chunked `ล้าน` reading at and beyond the million
// boundary including the chained `ล้านล้าน` for `1e12`, the documented
// rejection at `|n| >= 1e15`, the rounding-half-away-from-zero of the satang
// component, and the standard argument coercions (Bool / Text / Blank /
// Error / Array first-cell).

#include <string_view>

#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Parses `src` and evaluates it via the default function registry. Mirrors
// the helper used by `builtins_text_test.cpp`; thread-local arenas keep text
// payloads readable for the immediately following EXPECT_*.
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

// UTF-8 building blocks shared with the implementation. Repeating the literals
// here (rather than `#include`-ing the impl header) keeps the test as a true
// black-box check on the public output.
constexpr const char* kZero = "\xE0\xB8\xA8\xE0\xB8\xB9\xE0\xB8\x99\xE0\xB8\xA2\xE0\xB9\x8C";  // ศูนย์
constexpr const char* kOne = "\xE0\xB8\xAB\xE0\xB8\x99\xE0\xB8\xB6\xE0\xB9\x88\xE0\xB8\x87";   // หนึ่ง
constexpr const char* kTwo = "\xE0\xB8\xAA\xE0\xB8\xAD\xE0\xB8\x87";                           // สอง
constexpr const char* kThree = "\xE0\xB8\xAA\xE0\xB8\xB2\xE0\xB8\xA1";                         // สาม
constexpr const char* kFour = "\xE0\xB8\xAA\xE0\xB8\xB5\xE0\xB9\x88";                          // สี่
constexpr const char* kFive = "\xE0\xB8\xAB\xE0\xB9\x89\xE0\xB8\xB2";                          // ห้า
constexpr const char* kSix = "\xE0\xB8\xAB\xE0\xB8\x81";                                       // หก
constexpr const char* kSeven = "\xE0\xB9\x80\xE0\xB8\x88\xE0\xB9\x87\xE0\xB8\x94";             // เจ็ด
constexpr const char* kEight = "\xE0\xB9\x81\xE0\xB8\x9B\xE0\xB8\x94";                         // แปด
constexpr const char* kNine = "\xE0\xB9\x80\xE0\xB8\x81\xE0\xB9\x89\xE0\xB8\xB2";              // เก้า

constexpr const char* kSip = "\xE0\xB8\xAA\xE0\xB8\xB4\xE0\xB8\x9A";                           // สิบ
constexpr const char* kRoi = "\xE0\xB8\xA3\xE0\xB9\x89\xE0\xB8\xAD\xE0\xB8\xA2";               // ร้อย
constexpr const char* kPan = "\xE0\xB8\x9E\xE0\xB8\xB1\xE0\xB8\x99";                           // พัน
constexpr const char* kMuen = "\xE0\xB8\xAB\xE0\xB8\xA1\xE0\xB8\xB7\xE0\xB9\x88\xE0\xB8\x99";  // หมื่น
constexpr const char* kSaen = "\xE0\xB9\x81\xE0\xB8\xAA\xE0\xB8\x99";                          // แสน
constexpr const char* kLaan = "\xE0\xB8\xA5\xE0\xB9\x89\xE0\xB8\xB2\xE0\xB8\x99";              // ล้าน
constexpr const char* kYi = "\xE0\xB8\xA2\xE0\xB8\xB5\xE0\xB9\x88";                            // ยี่
constexpr const char* kEt = "\xE0\xB9\x80\xE0\xB8\xAD\xE0\xB9\x87\xE0\xB8\x94";                // เอ็ด
constexpr const char* kBaht = "\xE0\xB8\x9A\xE0\xB8\xB2\xE0\xB8\x97";                          // บาท
constexpr const char* kSatang = "\xE0\xB8\xAA\xE0\xB8\x95\xE0\xB8\xB2\xE0\xB8\x87\xE0\xB8\x84\xE0\xB9\x8C";  // สตางค์
constexpr const char* kThuan = "\xE0\xB8\x96\xE0\xB9\x89\xE0\xB8\xA7\xE0\xB8\x99";  // ถ้วน
constexpr const char* kLop = "\xE0\xB8\xA5\xE0\xB8\x9A";                            // ลบ

// ---------------------------------------------------------------------------
// Single digits (standalone).
// ---------------------------------------------------------------------------

TEST(Bahttext, Zero) {
  const Value v = EvalSource("=BAHTTEXT(0)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kZero) + kBaht + kThuan);
}

TEST(Bahttext, OneStandalone) {
  const Value v = EvalSource("=BAHTTEXT(1)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kOne) + kBaht + kThuan);
}

TEST(Bahttext, EachSingleDigit) {
  struct Case {
    const char* src;
    std::string expected;
  };
  const Case cases[] = {
      {"=BAHTTEXT(2)", std::string(kTwo) + kBaht + kThuan},   {"=BAHTTEXT(3)", std::string(kThree) + kBaht + kThuan},
      {"=BAHTTEXT(4)", std::string(kFour) + kBaht + kThuan},  {"=BAHTTEXT(5)", std::string(kFive) + kBaht + kThuan},
      {"=BAHTTEXT(6)", std::string(kSix) + kBaht + kThuan},   {"=BAHTTEXT(7)", std::string(kSeven) + kBaht + kThuan},
      {"=BAHTTEXT(8)", std::string(kEight) + kBaht + kThuan}, {"=BAHTTEXT(9)", std::string(kNine) + kBaht + kThuan},
  };
  for (const auto& c : cases) {
    const Value v = EvalSource(c.src);
    ASSERT_TRUE(v.is_text()) << c.src;
    EXPECT_EQ(v.as_text(), c.expected) << c.src;
  }
}

// ---------------------------------------------------------------------------
// Tens special readings: `สิบ` (10), `เอ็ด` ones, `ยี่` two.
// ---------------------------------------------------------------------------

TEST(Bahttext, Ten) {
  const Value v = EvalSource("=BAHTTEXT(10)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kSip) + kBaht + kThuan);
}

TEST(Bahttext, Eleven) {
  const Value v = EvalSource("=BAHTTEXT(11)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kSip) + kEt + kBaht + kThuan);
}

TEST(Bahttext, Twelve) {
  const Value v = EvalSource("=BAHTTEXT(12)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kSip) + kTwo + kBaht + kThuan);
}

TEST(Bahttext, Twenty) {
  const Value v = EvalSource("=BAHTTEXT(20)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kYi) + kSip + kBaht + kThuan);
}

TEST(Bahttext, TwentyOne) {
  const Value v = EvalSource("=BAHTTEXT(21)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kYi) + kSip + kEt + kBaht + kThuan);
}

TEST(Bahttext, TwentyTwo) {
  const Value v = EvalSource("=BAHTTEXT(22)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kYi) + kSip + kTwo + kBaht + kThuan);
}

TEST(Bahttext, Thirty) {
  const Value v = EvalSource("=BAHTTEXT(30)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kThree) + kSip + kBaht + kThuan);
}

TEST(Bahttext, NinetyNine) {
  const Value v = EvalSource("=BAHTTEXT(99)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kNine) + kSip + kNine + kBaht + kThuan);
}

// ---------------------------------------------------------------------------
// Hundreds & thousands.
// ---------------------------------------------------------------------------

TEST(Bahttext, OneHundred) {
  const Value v = EvalSource("=BAHTTEXT(100)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kOne) + kRoi + kBaht + kThuan);
}

TEST(Bahttext, OneHundredOne) {
  // Ones-`1` after a non-zero higher digit (here, the hundreds place) reads
  // as `เอ็ด`.
  const Value v = EvalSource("=BAHTTEXT(101)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kOne) + kRoi + kEt + kBaht + kThuan);
}

TEST(Bahttext, OneHundredEleven) {
  const Value v = EvalSource("=BAHTTEXT(111)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kOne) + kRoi + kSip + kEt + kBaht + kThuan);
}

TEST(Bahttext, OneHundredTwentyOne) {
  const Value v = EvalSource("=BAHTTEXT(121)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kOne) + kRoi + kYi + kSip + kEt + kBaht + kThuan);
}

TEST(Bahttext, TwoHundred) {
  const Value v = EvalSource("=BAHTTEXT(200)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kTwo) + kRoi + kBaht + kThuan);
}

TEST(Bahttext, OneThousand) {
  const Value v = EvalSource("=BAHTTEXT(1000)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kOne) + kPan + kBaht + kThuan);
}

TEST(Bahttext, OneThousandTwoHundredThirtyFour) {
  const Value v = EvalSource("=BAHTTEXT(1234)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kOne) + kPan + kTwo + kRoi + kThree + kSip + kFour + kBaht + kThuan);
}

TEST(Bahttext, NineThousandNineHundredNinetyNine) {
  const Value v = EvalSource("=BAHTTEXT(9999)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kNine) + kPan + kNine + kRoi + kNine + kSip + kNine + kBaht + kThuan);
}

// ---------------------------------------------------------------------------
// Ten-thousands and hundred-thousands.
// ---------------------------------------------------------------------------

TEST(Bahttext, TenThousand) {
  const Value v = EvalSource("=BAHTTEXT(10000)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kOne) + kMuen + kBaht + kThuan);
}

TEST(Bahttext, HundredThousand) {
  const Value v = EvalSource("=BAHTTEXT(100000)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kOne) + kSaen + kBaht + kThuan);
}

TEST(Bahttext, NineHundredNinetyNineThousandNineHundredNinetyNine) {
  const Value v = EvalSource("=BAHTTEXT(999999)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kNine) + kSaen + kNine + kMuen + kNine + kPan + kNine + kRoi + kNine + kSip +
                             kNine + kBaht + kThuan);
}

// ---------------------------------------------------------------------------
// Million boundary and chained `ล้าน`.
// ---------------------------------------------------------------------------

TEST(Bahttext, OneMillion) {
  const Value v = EvalSource("=BAHTTEXT(1000000)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kOne) + kLaan + kBaht + kThuan);
}

TEST(Bahttext, OneMillionTwoHundredThirtyFourThousandFiveHundredSixtySeven) {
  const Value v = EvalSource("=BAHTTEXT(1234567)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kOne) + kLaan + kTwo + kSaen + kThree + kMuen + kFour + kPan + kFive + kRoi +
                             kSix + kSip + kSeven + kBaht + kThuan);
}

TEST(Bahttext, NineHundredNinetyNineMillionEtc) {
  // 999,999,999 = 999 ล้าน 999,999.
  const Value v = EvalSource("=BAHTTEXT(999999999)");
  ASSERT_TRUE(v.is_text());
  const std::string nine_block = std::string(kNine) + kRoi + kNine + kSip + kNine;
  EXPECT_EQ(v.as_text(), nine_block + kLaan + std::string(kNine) + kSaen + kNine + kMuen + kNine + kPan + nine_block +
                             kBaht + kThuan);
}

TEST(Bahttext, OneTrillionChainsLaan) {
  // 1,000,000,000,000 = `หนึ่งล้านล้าน` (chained ล้าน).
  const Value v = EvalSource("=BAHTTEXT(1000000000000)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kOne) + kLaan + kLaan + kBaht + kThuan);
}

TEST(Bahttext, OneMillionAndOne) {
  // 1,000,001: subordinate single-`1` chunk reads as `เอ็ด`.
  const Value v = EvalSource("=BAHTTEXT(1000001)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kOne) + kLaan + kEt + kBaht + kThuan);
}

// ---------------------------------------------------------------------------
// Beyond the documented ceiling.
// ---------------------------------------------------------------------------

TEST(Bahttext, AtLimitReturnsValueError) {
  const Value v = EvalSource("=BAHTTEXT(1E15)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(Bahttext, BeyondLimitReturnsValueError) {
  const Value v = EvalSource("=BAHTTEXT(1E16)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// Negative numbers.
// ---------------------------------------------------------------------------

TEST(Bahttext, NegativeOne) {
  const Value v = EvalSource("=BAHTTEXT(-1)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kLop) + kOne + kBaht + kThuan);
}

TEST(Bahttext, NegativeOnePointTwoFive) {
  const Value v = EvalSource("=BAHTTEXT(-1.25)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kLop) + kOne + kBaht + kYi + kSip + kFive + kSatang);
}

TEST(Bahttext, NegativeMillions) {
  const Value v = EvalSource("=BAHTTEXT(-1234567.89)");
  ASSERT_TRUE(v.is_text());
  // -1,234,567.89: lop + (1 ล้าน 2 แสน 3 หมื่น 4 พัน 5 ร้อย 6 สิบ 7) baht + (8 สิบ 9) สตางค์
  const std::string integer_words =
      std::string(kOne) + kLaan + kTwo + kSaen + kThree + kMuen + kFour + kPan + kFive + kRoi + kSix + kSip + kSeven;
  const std::string satang_words = std::string(kEight) + kSip + kNine;
  EXPECT_EQ(v.as_text(), std::string(kLop) + integer_words + kBaht + satang_words + kSatang);
}

// ---------------------------------------------------------------------------
// Decimals only.
// ---------------------------------------------------------------------------

TEST(Bahttext, FiftySatang) {
  // 0.50 -> ศูนย์บาทห้าสิบสตางค์ (Excel still emits the leading `ศูนย์บาท`).
  const Value v = EvalSource("=BAHTTEXT(0.5)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kZero) + kBaht + kFive + kSip + kSatang);
}

TEST(Bahttext, NinetyNineSatang) {
  const Value v = EvalSource("=BAHTTEXT(0.99)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kZero) + kBaht + kNine + kSip + kNine + kSatang);
}

TEST(Bahttext, OneSatang) {
  const Value v = EvalSource("=BAHTTEXT(0.01)");
  ASSERT_TRUE(v.is_text());
  // ones digit `1` standalone in the satang chunk -> `หนึ่ง`.
  EXPECT_EQ(v.as_text(), std::string(kZero) + kBaht + kOne + kSatang);
}

// ---------------------------------------------------------------------------
// Mixed integer + fractional.
// ---------------------------------------------------------------------------

TEST(Bahttext, OnePointTwoFive) {
  const Value v = EvalSource("=BAHTTEXT(1.25)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kOne) + kBaht + kYi + kSip + kFive + kSatang);
}

TEST(Bahttext, OneHundredPointFive) {
  const Value v = EvalSource("=BAHTTEXT(100.5)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kOne) + kRoi + kBaht + kFive + kSip + kSatang);
}

TEST(Bahttext, NearMaxIntegerWithSatang) {
  // 999,999.99 -> nine hundred ninety-nine thousand nine hundred ninety-nine
  // baht ninety-nine satang.
  const Value v = EvalSource("=BAHTTEXT(999999.99)");
  ASSERT_TRUE(v.is_text());
  const std::string integer_words =
      std::string(kNine) + kSaen + kNine + kMuen + kNine + kPan + kNine + kRoi + kNine + kSip + kNine;
  const std::string satang_words = std::string(kNine) + kSip + kNine;
  EXPECT_EQ(v.as_text(), integer_words + kBaht + satang_words + kSatang);
}

TEST(Bahttext, OneHundredTwentyThreePointFourFive) {
  // 123.45 -> one hundred twenty-three baht forty-five satang.
  const Value v = EvalSource("=BAHTTEXT(123.45)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kOne) + kRoi + kYi + kSip + kThree + kBaht + kFour + kSip + kFive + kSatang);
}

// ---------------------------------------------------------------------------
// Rounding (away from zero, two decimals).
// ---------------------------------------------------------------------------

TEST(Bahttext, RoundsHalfAwayFromZeroPositive) {
  // 0.005 -> 0.01 -> ศูนย์บาทหนึ่งสตางค์.
  const Value v = EvalSource("=BAHTTEXT(0.005)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kZero) + kBaht + kOne + kSatang);
}

TEST(Bahttext, RoundsTowardZeroBelowMidpoint) {
  // 1.234 -> 1.23.
  const Value v = EvalSource("=BAHTTEXT(1.234)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kOne) + kBaht + kYi + kSip + kThree + kSatang);
}

TEST(Bahttext, RoundsAwayFromZeroAtMidpoint) {
  // 1.235 -> 1.24.
  const Value v = EvalSource("=BAHTTEXT(1.235)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kOne) + kBaht + kYi + kSip + kFour + kSatang);
}

// ---------------------------------------------------------------------------
// Argument coercion.
// ---------------------------------------------------------------------------

TEST(Bahttext, BoolTrueIsOneBaht) {
  const Value v = EvalSource("=BAHTTEXT(TRUE)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kOne) + kBaht + kThuan);
}

TEST(Bahttext, BoolFalseIsZeroBaht) {
  const Value v = EvalSource("=BAHTTEXT(FALSE)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kZero) + kBaht + kThuan);
}

TEST(Bahttext, NumericText) {
  const Value v = EvalSource("=BAHTTEXT(\"100\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), std::string(kOne) + kRoi + kBaht + kThuan);
}

TEST(Bahttext, NonNumericTextSurfacesValueError) {
  const Value v = EvalSource("=BAHTTEXT(\"abc\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(Bahttext, ErrorPropagates) {
  const Value v = EvalSource("=BAHTTEXT(#REF!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(Bahttext, NaPropagates) {
  const Value v = EvalSource("=BAHTTEXT(#N/A)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
