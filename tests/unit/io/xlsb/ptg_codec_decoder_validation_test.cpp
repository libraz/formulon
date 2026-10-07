// XLSB Ptg codec tests grouped by reference, function, scalar, and validation behavior.

#include <cstdio>
#include <cstring>
#include <string>
#include <string_view>
#include <unordered_set>
#include <vector>

#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "io/future_functions.h"
#include "io/xlsb/func_id_table.h"
#include "io/xlsb/ptg_reader.h"
#include "io/xlsb/ptg_writer.h"
#include "io/xlsb/record_writer.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/parser.h"
#include "ptg_codec_test_helpers.h"
#include "utils/arena.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {
using namespace ptg_codec_test_support;
TEST(XlsbPtgCodec, DecoderRejectsTruncatedStream) {
  // PtgInt (0x1E) needs a 2-byte operand; supply only the tag.
  Arena arena;
  const std::vector<std::uint8_t> bytes = {0x1E};
  ByteSpan span{bytes.data(), bytes.size()};
  auto decoded = decode_ptgs(span, ByteSpan{}, arena, {}, {}, {}, {});
  ASSERT_FALSE(static_cast<bool>(decoded));
  EXPECT_EQ(decoded.error().code, FormulonErrorCode::kIoXlsbRecordTruncated);
}
TEST(XlsbPtgCodec, DecoderRejectsUnknownPtg) {
  // 0x18 (PtgElfLel) is marked Unsupported in the dispatch table.
  Arena arena;
  const std::vector<std::uint8_t> bytes = {0x18, 0x00};
  ByteSpan span{bytes.data(), bytes.size()};
  auto decoded = decode_ptgs(span, ByteSpan{}, arena, {}, {}, {}, {});
  ASSERT_FALSE(static_cast<bool>(decoded));
  EXPECT_EQ(decoded.error().code, FormulonErrorCode::kIoXlsbUnsupportedPtg);
}
TEST(XlsbPtgCodec, PtgArrayDecodesRowsBeforeColsFromRawWireBytes) {
  // Raw-byte decode of a genuinely non-square (1 row x 3 cols) array
  // constant, independent of the encoder. `PtgExtraArray`
  // ([MS-XLSB] 2.5.97.41) is documented as `DRw` (row count) followed
  // by `DCol` (col count), each a plain u32 -- not a class-marked or
  // otherwise reordered pair, and the elements follow row-outer /
  // col-inner. A square real-Excel fixture cannot settle either point on
  // its own, so this test pins them from the wire bytes. If the reader
  // swapped the two u32 fields, this byte layout (rows=1, cols=3) would
  // either be rejected (3 rows needs 3x as many trailing elements) or,
  // for a same-total-count swap, transpose to `{1;2;3}`.
  const std::vector<std::uint8_t> rgce = {
      0x60,                                                  // PtgArray (array-class)
      0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,  // 14-byte placeholder
      0x00, 0x00, 0x00, 0x00, 0x00,
  };
  std::vector<std::uint8_t> rgcb;
  emit_u32(rgcb, 1U);  // DRw: rows = 1
  emit_u32(rgcb, 3U);  // DCol: cols = 3
  for (const double v : {1.0, 2.0, 3.0}) {
    rgcb.push_back(0x00);  // SerAr numeric tag
    std::uint8_t bytes[8];
    std::memcpy(bytes, &v, sizeof(v));
    rgcb.insert(rgcb.end(), bytes, bytes + 8);
  }
  Arena arena;
  auto decoded =
      decode_ptgs(ByteSpan{rgce.data(), rgce.size()}, ByteSpan{rgcb.data(), rgcb.size()}, arena, {}, {}, {}, {});
  ASSERT_TRUE(static_cast<bool>(decoded)) << (decoded ? "" : decoded.error().message);
  EXPECT_EQ(parser::format_formula(*decoded.value()), "{1,2,3}");
}
TEST(XlsbPtgCodec, PtgArrayElementsMatchExcelBytes) {
  struct Case {
    const char* formula;
    std::vector<std::uint8_t> rgce_head;  // opcode + the 10 bytes Excel sets
    std::vector<std::uint8_t> rgcb;
  };
  const Case cases[] = {
      {"{\"a\",\"b\";\"c\",\"d\"}",
       {0x40, 0x01, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00},
       {0x02, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x01, 0x01, 0x00, 0x61, 0x00, 0x01,
        0x01, 0x00, 0x62, 0x00, 0x01, 0x01, 0x00, 0x63, 0x00, 0x01, 0x01, 0x00, 0x64, 0x00}},
      {"{TRUE;FALSE}",
       {0x40, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00},
       {0x02, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x02, 0x01, 0x02, 0x00}},
      {"{1;#N/A;\"z\"}",
       {0x40, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00},
       {0x03, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0xF0, 0x3F, 0x04, 0x2A, 0x00, 0x00, 0x00, 0x01, 0x01, 0x00, 0x7A, 0x00}},
  };
  for (const Case& c : cases) {
    const EncodedFormula encoded = EncodeOnSheet1(c.formula, PtgRootClass::kValue);
    ASSERT_EQ(encoded.rgce.size(), 15U) << c.formula;
    EXPECT_EQ(std::vector<std::uint8_t>(encoded.rgce.begin(), encoded.rgce.begin() + 11), c.rgce_head) << c.formula;
    EXPECT_EQ(encoded.rgcb, c.rgcb) << c.formula;
    EXPECT_EQ(RoundTrip(c.formula), c.formula);
  }
}
TEST(XlsbPtgCodec, PtgRefRowAtGridBoundIsRecordCorrupt) {
  // PtgRef (0x24) with row == Sheet::kMaxRows, one past the last valid
  // row (1048575). A crafted/corrupt row must be rejected at the Ref
  // node construction site rather than silently wrapping in a later
  // A1-text re-parse (`format_a1` would otherwise print `A1048577`).
  const std::vector<std::uint8_t> rgce = {
      0x24,                    // PtgRef
      0x00, 0x00, 0x10, 0x00,  // row = 1048576 (u32 LE) == Sheet::kMaxRows
      0x00, 0x00,              // col = 0, absolute
  };
  Arena arena;
  auto decoded = decode_ptgs(ByteSpan{rgce.data(), rgce.size()}, {}, arena, {}, {}, {}, {});
  ASSERT_FALSE(static_cast<bool>(decoded));
  EXPECT_EQ(decoded.error().code, FormulonErrorCode::kIoXlsbRecordCorrupt);
}
TEST(XlsbPtgCodec, PtgAreaReversedCornersIsRecordCorrupt) {
  // PtgArea (0x25) with row1 > row2: both corners are individually
  // in-domain, but the range is not normalized.
  const std::vector<std::uint8_t> rgce = {
      0x25,                    // PtgArea
      0x05, 0x00, 0x00, 0x00,  // row1 = 5
      0x01, 0x00, 0x00, 0x00,  // row2 = 1 (< row1)
      0x00, 0x00,              // col1 = 0, absolute
      0x00, 0x00,              // col2 = 0, absolute
  };
  Arena arena;
  auto decoded = decode_ptgs(ByteSpan{rgce.data(), rgce.size()}, {}, arena, {}, {}, {}, {});
  ASSERT_FALSE(static_cast<bool>(decoded));
  EXPECT_EQ(decoded.error().code, FormulonErrorCode::kIoXlsbRecordCorrupt);
}
TEST(XlsbPtgCodec, PtgRefErrWithSentinelCoordinatesStillDecodesAsRef) {
  // PtgRefErr (0x2A) carries a `#REF!`-form single-cell reference; real
  // Excel files encode the dead payload with the maximum sentinel
  // row/col. Because this Ptg kind never materializes a `Reference`
  // (only `ErrorCode::Ref`), the domain check at Ref/Area construction
  // sites must not reject it.
  const std::vector<std::uint8_t> rgce = {
      0x2A,                    // PtgRefErr
      0xFF, 0xFF, 0xFF, 0xFF,  // sentinel row payload (discarded)
      0xFF, 0xFF,              // sentinel col payload (discarded)
  };
  Arena arena;
  auto decoded = decode_ptgs(ByteSpan{rgce.data(), rgce.size()}, {}, arena, {}, {}, {}, {});
  ASSERT_TRUE(static_cast<bool>(decoded)) << (decoded ? "" : decoded.error().message);
  EXPECT_EQ(parser::format_formula(*decoded.value()), "#REF!");
}
TEST(XlsbPtgCodec, ErrorReferenceFormsConsumeTheirPayloads) {
  // PtgAreaErr (0x2B): two locs. PtgRefErr3d (0x3C) / PtgAreaErr3d (0x3D):
  // an ixti, then one or two locs. Each is followed by PtgInt(1) and PtgAdd,
  // so a payload consumed short or long misreads the rest of the stream.
  const std::vector<std::vector<std::uint8_t>> streams = {
      {0x2B, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0x1E, 1, 0, 0x03},
      {0x3C, 0, 0, 0, 0, 0, 0, 0, 0, 0x1E, 1, 0, 0x03},
      {0x3D, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0x1E, 1, 0, 0x03},
  };
  for (const std::vector<std::uint8_t>& rgce : streams) {
    Arena arena;
    auto decoded = decode_ptgs(ByteSpan{rgce.data(), rgce.size()}, {}, arena, {}, {}, {}, {});
    ASSERT_TRUE(static_cast<bool>(decoded)) << (decoded ? "" : decoded.error().message);
    EXPECT_EQ(parser::format_formula(*decoded.value()), "#REF!+1");
  }
  // One byte short of the second loc.
  const std::vector<std::uint8_t> truncated = {0x2B, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 1};
  Arena arena;
  auto decoded = decode_ptgs(ByteSpan{truncated.data(), truncated.size()}, {}, arena, {}, {}, {}, {});
  ASSERT_FALSE(static_cast<bool>(decoded));
  EXPECT_EQ(decoded.error().code, FormulonErrorCode::kIoXlsbRecordTruncated);
}
TEST(XlsbPtgCodec, DecoderKeepsParenTokenInTheText) {
  // PtgInt(1), PtgInt(2), PtgAdd, PtgParen: the written parentheses Excel
  // shows, which leave the value stack unchanged.
  const std::vector<std::uint8_t> rgce = {0x1E, 0x01, 0x00, 0x1E, 0x02, 0x00, 0x03, 0x15};
  Arena arena;
  ByteSpan main{rgce.data(), rgce.size()};
  ByteSpan extra{nullptr, 0};
  auto decoded = decode_ptgs(main, extra, arena, {}, {}, {}, {});
  ASSERT_TRUE(static_cast<bool>(decoded));
  EXPECT_EQ(parser::format_formula(*decoded.value()), "(1+2)");
}
TEST(XlsbPtgCodec, DecoderConsumesMemoryCachePtgsWithoutChangingExpression) {
  // Each memory marker precedes an ordinary PtgInt(1) expression. PtgMemArea
  // additionally has a zero-range PtgExtraMem cache in RgbExtra.
  const std::vector<std::uint8_t> mem_area = {0x26, 0, 0, 0, 0, 3, 0, 0x1E, 1, 0};
  const std::vector<std::uint8_t> mem_err = {0x27, 0x17, 0, 0, 0, 3, 0, 0x1E, 1, 0};
  const std::vector<std::uint8_t> mem_no_mem = {0x28, 0, 0, 0, 0, 3, 0, 0x1E, 1, 0};
  const std::vector<std::uint8_t> mem_func = {0x29, 3, 0, 0x1E, 1, 0};
  const std::vector<std::uint8_t> extra_mem = {0, 0, 0, 0};

  const auto expect_expression = [&extra_mem](const std::vector<std::uint8_t>& rgce, bool has_extra_mem) {
    Arena arena;
    ByteSpan main{rgce.data(), rgce.size()};
    ByteSpan extra = has_extra_mem ? ByteSpan{extra_mem.data(), extra_mem.size()} : ByteSpan{};
    auto decoded = decode_ptgs(main, extra, arena, {}, {}, {}, {});
    ASSERT_TRUE(static_cast<bool>(decoded)) << (decoded ? "" : decoded.error().message);
    EXPECT_EQ(parser::format_formula(*decoded.value()), "1");
  };
  expect_expression(mem_area, true);
  expect_expression(mem_err, false);
  expect_expression(mem_no_mem, false);
  expect_expression(mem_func, false);
}
TEST(XlsbPtgCodec, DecoderRejectsTruncatedMemAreaExtraCache) {
  const std::vector<std::uint8_t> rgce = {0x26, 0, 0, 0, 0, 3, 0, 0x1E, 1, 0};
  // PtgExtraMem declares one 16-byte range but supplies none.
  const std::vector<std::uint8_t> extra = {1, 0, 0, 0};
  Arena arena;
  auto decoded =
      decode_ptgs(ByteSpan{rgce.data(), rgce.size()}, ByteSpan{extra.data(), extra.size()}, arena, {}, {}, {}, {});
  ASSERT_FALSE(static_cast<bool>(decoded));
  EXPECT_EQ(decoded.error().code, FormulonErrorCode::kIoXlsbRecordTruncated);
}
TEST(XlsbPtgCodec, DecoderRejectsAstDeeperThanSharedLimit) {
  Arena enc_arena;
  parser::AstNode* root = parser::make_literal(enc_arena, Value::number(1));
  for (std::uint32_t depth = 1; depth <= parser::kMaxFormulaAstDepth; ++depth) {
    root = parser::make_unary_op(enc_arena, parser::UnaryOp::Plus, root);
  }
  ASSERT_NE(root, nullptr);
  auto encoded = encode_ptgs(*root, {}, {}, {}, PtgRootClass::kValue);
  ASSERT_TRUE(static_cast<bool>(encoded));

  Arena dec_arena;
  ByteSpan rgce{encoded.value().rgce.data(), encoded.value().rgce.size()};
  ByteSpan rgcb{encoded.value().rgcb.data(), encoded.value().rgcb.size()};
  auto decoded = decode_ptgs(rgce, rgcb, dec_arena, {}, {}, {}, {});
  ASSERT_FALSE(static_cast<bool>(decoded));
  EXPECT_EQ(decoded.error().code, FormulonErrorCode::kIoXlsbCorrupt);
}
TEST(XlsbPtgCodec, EncoderUsesExcelClassesForReferenceArgumentsAndFunctionResults) {
  Arena arena;
  parser::Parser p("SUM(A1:A2,1)", arena);
  parser::AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  ASSERT_TRUE(p.errors().empty());
  auto encoded = encode_ptgs(*root, {}, {}, {}, PtgRootClass::kValue);
  ASSERT_TRUE(static_cast<bool>(encoded)) << (encoded ? "" : encoded.error().message);

  // Real Excel writes the range argument as reference-class PtgArea (0x25)
  // and the SUM result as value-class PtgFuncVar (0x42). Array-class 0x65
  // and reference-class 0x22 respectively make Excel repair the worksheet.
  ASSERT_GE(encoded.value().rgce.size(), 4U);
  EXPECT_EQ(encoded.value().rgce.front(), 0x25U);
  EXPECT_EQ(encoded.value().rgce[encoded.value().rgce.size() - 4U], 0x42U);
}
TEST(XlsbPtgCodec, EncoderRejectsUnregisteredDefinedName) {
  // A `NameRef` only lowers to `PtgName` when the caller's `name_table`
  // (built from `collect_ptg_names` across the whole workbook) already
  // carries an `ilbl` for it; an empty table means "not registered".
  Arena arena;
  parser::Parser p("MyName", arena);
  parser::AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  auto encoded = encode_ptgs(*root, {}, {}, {}, PtgRootClass::kValue);
  ASSERT_FALSE(static_cast<bool>(encoded));
  EXPECT_EQ(encoded.error().code, FormulonErrorCode::kIoXlsbUnsupportedPtg);
}
TEST(XlsbPtgCodec, EncoderLowersRegisteredDefinedName) {
  Arena arena;
  parser::Parser p("MyName", arena);
  parser::AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  const NameTable names = {{"MyName", 1U}};
  auto encoded = encode_ptgs(*root, {}, {}, names, PtgRootClass::kValue);
  ASSERT_TRUE(static_cast<bool>(encoded)) << (encoded ? "" : encoded.error().message);

  Arena dec_arena;
  ByteSpan rgce{encoded.value().rgce.data(), encoded.value().rgce.size()};
  const std::vector<XlsbName> name_table = {{-1, "MyName", false}};
  auto decoded = decode_ptgs(rgce, ByteSpan{}, dec_arena, {}, name_table, {}, {});
  ASSERT_TRUE(static_cast<bool>(decoded)) << (decoded ? "" : decoded.error().message);
  EXPECT_EQ(parser::format_formula(*decoded.value()), "MyName");
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
