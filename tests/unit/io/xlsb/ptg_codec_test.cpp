// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Round-trip tests for the MS-XLSB Ptg codec (encoder + decoder).
//
// Each case parses an A1 formula to the engine AST, encodes it to a Ptg
// (`rgce`) byte stream, decodes that stream back to an AST, and asserts
// the re-formatted formula text matches the original. This exercises the
// encode <-> decode pair as a single round-trip without needing a full
// xlsb package.

#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "io/xlsb/ptg_reader.h"
#include "io/xlsb/ptg_writer.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/parser.h"
#include "utils/arena.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

// Parses `formula` (without leading `=`), encodes to Ptg, decodes back,
// and returns the re-formatted formula text. `sheet_names` resolves
// 3-D-reference `ixti` indices on the decode side.
std::string RoundTrip(std::string_view formula, const std::vector<std::string>& sheet_names = {}) {
  Arena enc_arena;
  parser::Parser p(formula, enc_arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr);
  EXPECT_TRUE(p.errors().empty()) << "parse errors for: " << formula;

  auto encoded = encode_ptgs(*root, sheet_names);
  EXPECT_TRUE(static_cast<bool>(encoded))
      << "encode failed for: " << formula << " | " << (encoded ? "" : encoded.error().message);
  if (!encoded) {
    return "<encode-failed>";
  }

  Arena dec_arena;
  ByteSpan span{encoded.value().data(), encoded.value().size()};
  auto decoded = decode_ptgs(span, dec_arena, sheet_names);
  EXPECT_TRUE(static_cast<bool>(decoded))
      << "decode failed for: " << formula << " | " << (decoded ? "" : decoded.error().message);
  if (!decoded) {
    return "<decode-failed>";
  }
  return parser::format_formula(*decoded.value());
}

TEST(XlsbPtgCodec, ArithmeticWithPrecedence) {
  // `A1+B2*3`: PtgRef, PtgRef, PtgInt, PtgMul, PtgAdd.
  EXPECT_EQ(RoundTrip("A1+B2*3"), "A1+B2*3");
}

TEST(XlsbPtgCodec, SumOverArea) {
  EXPECT_EQ(RoundTrip("SUM(A1:A10)"), "SUM(A1:A10)");
}

TEST(XlsbPtgCodec, IfWithStrings) {
  EXPECT_EQ(RoundTrip("IF(A1>0,\"pos\",\"neg\")"), "IF(A1>0,\"pos\",\"neg\")");
}

TEST(XlsbPtgCodec, Concat) {
  EXPECT_EQ(RoundTrip("B1&\"x\""), "B1&\"x\"");
}

TEST(XlsbPtgCodec, UnaryMinus) {
  EXPECT_EQ(RoundTrip("-A1"), "-A1");
}

TEST(XlsbPtgCodec, PostfixPercent) {
  EXPECT_EQ(RoundTrip("A1%"), "A1%");
}

TEST(XlsbPtgCodec, ThreeDimensionalReference) {
  // `Sheet2!A1` resolves through the sheet-name list to a PtgRef3d.
  const std::vector<std::string> sheets = {"Sheet1", "Sheet2"};
  EXPECT_EQ(RoundTrip("Sheet2!A1", sheets), "Sheet2!A1");
}

TEST(XlsbPtgCodec, ConstantArray) {
  EXPECT_EQ(RoundTrip("{1,2;3,4}"), "{1,2;3,4}");
}

TEST(XlsbPtgCodec, ErrorLiteral) {
  EXPECT_EQ(RoundTrip("#DIV/0!"), "#DIV/0!");
}

TEST(XlsbPtgCodec, AbsoluteReference) {
  EXPECT_EQ(RoundTrip("$A$1"), "$A$1");
  EXPECT_EQ(RoundTrip("$A1"), "$A1");
  EXPECT_EQ(RoundTrip("A$1"), "A$1");
}

TEST(XlsbPtgCodec, AllComparisons) {
  EXPECT_EQ(RoundTrip("A1<B1"), "A1<B1");
  EXPECT_EQ(RoundTrip("A1<=B1"), "A1<=B1");
  EXPECT_EQ(RoundTrip("A1=B1"), "A1=B1");
  EXPECT_EQ(RoundTrip("A1>=B1"), "A1>=B1");
  EXPECT_EQ(RoundTrip("A1>B1"), "A1>B1");
  EXPECT_EQ(RoundTrip("A1<>B1"), "A1<>B1");
}

TEST(XlsbPtgCodec, NestedFunctions) {
  EXPECT_EQ(RoundTrip("ROUND(SUM(A1:A3),2)"), "ROUND(SUM(A1:A3),2)");
}

TEST(XlsbPtgCodec, PowerAndDivide) {
  EXPECT_EQ(RoundTrip("A1^2/B1"), "A1^2/B1");
}

TEST(XlsbPtgCodec, DecoderRejectsTruncatedStream) {
  // PtgInt (0x1E) needs a 2-byte operand; supply only the tag.
  Arena arena;
  const std::vector<std::uint8_t> bytes = {0x1E};
  ByteSpan span{bytes.data(), bytes.size()};
  auto decoded = decode_ptgs(span, arena, {});
  ASSERT_FALSE(static_cast<bool>(decoded));
  EXPECT_EQ(decoded.error().code, FormulonErrorCode::kIoXlsbRecordTruncated);
}

TEST(XlsbPtgCodec, DecoderRejectsUnknownPtg) {
  // 0x18 (PtgElfLel) is marked Unsupported in the dispatch table.
  Arena arena;
  const std::vector<std::uint8_t> bytes = {0x18, 0x00};
  ByteSpan span{bytes.data(), bytes.size()};
  auto decoded = decode_ptgs(span, arena, {});
  ASSERT_FALSE(static_cast<bool>(decoded));
  EXPECT_EQ(decoded.error().code, FormulonErrorCode::kIoXlsbUnsupportedPtg);
}

TEST(XlsbPtgCodec, EncoderRejectsDefinedName) {
  Arena arena;
  parser::Parser p("MyName", arena);
  parser::AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  auto encoded = encode_ptgs(*root, {});
  ASSERT_FALSE(static_cast<bool>(encoded));
  EXPECT_EQ(encoded.error().code, FormulonErrorCode::kIoXlsbUnsupportedPtg);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
