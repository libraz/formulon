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
#include <unordered_set>
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
// and returns the re-formatted formula text. `sheet_names` resolves a
// qualified reference's sheet to its 0-based index on both sides; the
// `SheetRangeTable` / `XlsbSheetRange` list that actually carries the
// `ixti` numbering is built here (via `collect_ptg_sheet_ranges` on the
// encode side, mirrored 1:1 into `XlsbSheetRange`s for decode) so a
// single-sheet qualified reference and a genuine 3-D range share one
// `ixti` space exactly as the production writer does.
std::string RoundTrip(std::string_view formula, const std::vector<std::string>& sheet_names = {}) {
  Arena enc_arena;
  parser::Parser p(formula, enc_arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr);
  EXPECT_TRUE(p.errors().empty()) << "parse errors for: " << formula;

  SheetRangeTable sheet_ranges;
  std::unordered_set<std::uint64_t> seen;
  collect_ptg_sheet_ranges(*root, sheet_names, sheet_ranges, seen);

  auto encoded = encode_ptgs(*root, sheet_names, sheet_ranges, {});
  EXPECT_TRUE(static_cast<bool>(encoded))
      << "encode failed for: " << formula << " | " << (encoded ? "" : encoded.error().message);
  if (!encoded) {
    return "<encode-failed>";
  }

  std::vector<XlsbSheetRange> decode_ranges;
  decode_ranges.reserve(sheet_ranges.size());
  for (const auto& [itab_first, itab_last] : sheet_ranges) {
    decode_ranges.push_back(XlsbSheetRange{itab_first, itab_last});
  }

  Arena dec_arena;
  ByteSpan rgce{encoded.value().rgce.data(), encoded.value().rgce.size()};
  ByteSpan rgcb{encoded.value().rgcb.data(), encoded.value().rgcb.size()};
  auto decoded = decode_ptgs(rgce, rgcb, dec_arena, sheet_names, {}, decode_ranges);
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

TEST(XlsbPtgCodec, AttrChooseSkipsU32JumpOffsets) {
  // PtgInt(1), followed by PtgAttrChoose with count=1 and two 32-bit
  // jump offsets. The control-flow table is irrelevant after XLSB has
  // already selected its cached formula path, but its wire width must be
  // consumed exactly or the following Ptg stream is misaligned.
  const std::vector<std::uint8_t> rgce = {
      0x1E, 0x01, 0x00,        // PtgInt 1
      0x19, 0x04, 0x01, 0x00,  // PtgAttrChoose, u16 count=1
      0x00, 0x00, 0x00, 0x00,  // jump offset 0 (u32)
      0x04, 0x00, 0x00, 0x00,  // jump offset 1 (u32)
  };
  Arena arena;
  auto decoded = decode_ptgs(ByteSpan{rgce.data(), rgce.size()}, {}, arena, {}, {}, {});
  ASSERT_TRUE(static_cast<bool>(decoded)) << (decoded ? "" : decoded.error().message);
  EXPECT_EQ(parser::format_formula(*decoded.value()), "1");
}

TEST(XlsbPtgCodec, SumOverArea) {
  EXPECT_EQ(RoundTrip("SUM(A1:A10)"), "SUM(A1:A10)");
}

TEST(XlsbPtgCodec, WholeColumnAndRowRefsEncodeAsSentinelAreas) {
  EXPECT_EQ(RoundTrip("SUM(A:A)"), "SUM(A1:A1048576)");
  EXPECT_EQ(RoundTrip("SUM(1:1)"), "SUM(A1:XFD1)");
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

TEST(XlsbPtgCodec, GenuineThreeDimensionalRangeRoundTrips) {
  // A genuine multi-sheet range (`Sheet1:Sheet3!B2`) is hand-built here
  // rather than parsed from text: the text parser does not yet lower
  // `'Sheet1:Sheet3'!B2` (or an unquoted `Sheet1:Sheet3!B2`) to a `Ref3D`
  // node, so this exercises the encoder/decoder pair -- the actual XLSB
  // fidelity contract -- directly. `encode_ptgs` resolves the node's
  // `(begin, end)` span through a `SheetRangeTable` built the same way
  // the production writer builds one (`collect_ptg_sheet_ranges`), and
  // `decode_ptgs` resolves it back through the equivalent `XlsbSheetRange`
  // list, mirroring how a real `BrtExternSheet` record round-trips.
  const std::vector<std::string> sheets = {"Sheet1", "Sheet2", "Sheet3"};
  Arena arena;
  parser::Reference cell;
  cell.row = 1;
  cell.col = 1;  // B2
  parser::AstNode* node = parser::make_ref3d(arena, "Sheet1", "Sheet3", cell);
  ASSERT_NE(node, nullptr);

  const SheetRangeTable sheet_ranges = {{0, 2}};  // Sheet1 (itab 0) : Sheet3 (itab 2)
  auto encoded = encode_ptgs(*node, sheets, sheet_ranges, {});
  ASSERT_TRUE(static_cast<bool>(encoded)) << (encoded ? "" : encoded.error().message);

  Arena dec_arena;
  ByteSpan rgce{encoded.value().rgce.data(), encoded.value().rgce.size()};
  const std::vector<XlsbSheetRange> decode_ranges = {{0, 2}};
  auto decoded = decode_ptgs(rgce, ByteSpan{}, dec_arena, sheets, {}, decode_ranges);
  ASSERT_TRUE(static_cast<bool>(decoded)) << (decoded ? "" : decoded.error().message);
  EXPECT_EQ(parser::format_formula(*decoded.value()), "Sheet1:Sheet3!B2");
}

TEST(XlsbPtgCodec, GenuineThreeDimensionalRangeTailRoundTrips) {
  // A genuine 3-D range tail (`Sheet1:Sheet3!A1:B2`) encodes as PtgArea3d
  // (ixti + RgceArea) and decodes back to a range-tail `Ref3D`, preserving
  // both the sheet span and the cell rectangle. The parser now lowers this
  // form directly, so drive it through the text `RoundTrip` helper.
  const std::vector<std::string> sheets = {"Sheet1", "Sheet2", "Sheet3"};
  EXPECT_EQ(RoundTrip("SUM(Sheet1:Sheet3!A1:B2)", sheets), "SUM(Sheet1:Sheet3!A1:B2)");

  // Also drive the encoder/decoder directly from a hand-built node to pin
  // the PtgArea3d codec contract independent of the parser.
  Arena arena;
  parser::Reference a;  // A1
  parser::Reference b;  // B2
  b.row = 1;
  b.col = 1;
  parser::AstNode* node = parser::make_ref3d_range(arena, "Sheet1", "Sheet3", a, b);
  ASSERT_NE(node, nullptr);
  const SheetRangeTable sheet_ranges = {{0, 2}};
  auto encoded = encode_ptgs(*node, sheets, sheet_ranges, {});
  ASSERT_TRUE(static_cast<bool>(encoded)) << (encoded ? "" : encoded.error().message);
  Arena dec_arena;
  ByteSpan rgce{encoded.value().rgce.data(), encoded.value().rgce.size()};
  const std::vector<XlsbSheetRange> decode_ranges = {{0, 2}};
  auto decoded = decode_ptgs(rgce, ByteSpan{}, dec_arena, sheets, {}, decode_ranges);
  ASSERT_TRUE(static_cast<bool>(decoded)) << (decoded ? "" : decoded.error().message);
  EXPECT_EQ(parser::format_formula(*decoded.value()), "Sheet1:Sheet3!A1:B2");
}

TEST(XlsbPtgCodec, SingleAndMultiSheetReferencesShareOneIxtiSpace) {
  // Once a workbook emits any `BrtExternSheet` entry, every `PtgRef3d`
  // (single- or multi-sheet) resolves its `ixti` through that one table
  // -- see `SheetRangeTable`'s doc comment. This pins that a formula
  // mixing a plain single-sheet ref with a genuine 3-D range encodes
  // and decodes both correctly against a shared table.
  const std::vector<std::string> sheets = {"Sheet1", "Sheet2", "Sheet3"};
  EXPECT_EQ(RoundTrip("Sheet2!A1+Sheet1!A1", sheets), "Sheet2!A1+Sheet1!A1");
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

TEST(XlsbPtgCodec, Post2007BuiltinsUseNativeFunctionIds) {
  EXPECT_EQ(RoundTrip("ASC(\"Ａ\")"), "ASC(\"Ａ\")");
  EXPECT_EQ(RoundTrip("JIS(\"A\")"), "JIS(\"A\")");
  EXPECT_EQ(RoundTrip("EDATE(A1,1)"), "EDATE(A1,1)");
  EXPECT_EQ(RoundTrip("EOMONTH(A1,1)"), "EOMONTH(A1,1)");
  EXPECT_EQ(RoundTrip("WORKDAY(A1,1)"), "WORKDAY(A1,1)");
  EXPECT_EQ(RoundTrip("NETWORKDAYS(A1,B1)"), "NETWORKDAYS(A1,B1)");
  EXPECT_EQ(RoundTrip("IFERROR(A1,0)"), "IFERROR(A1,0)");
  EXPECT_EQ(RoundTrip("COUNTIFS(A1,1)"), "COUNTIFS(A1,1)");
  EXPECT_EQ(RoundTrip("SUMIFS(A1,B1,1)"), "SUMIFS(A1,B1,1)");
  EXPECT_EQ(RoundTrip("AVERAGEIF(A1,1)"), "AVERAGEIF(A1,1)");
  EXPECT_EQ(RoundTrip("AVERAGEIFS(A1,B1,1)"), "AVERAGEIFS(A1,B1,1)");
}

TEST(XlsbPtgCodec, PowerAndDivide) {
  EXPECT_EQ(RoundTrip("A1^2/B1"), "A1^2/B1");
}

TEST(XlsbPtgCodec, DecoderRejectsTruncatedStream) {
  // PtgInt (0x1E) needs a 2-byte operand; supply only the tag.
  Arena arena;
  const std::vector<std::uint8_t> bytes = {0x1E};
  ByteSpan span{bytes.data(), bytes.size()};
  auto decoded = decode_ptgs(span, ByteSpan{}, arena, {}, {}, {});
  ASSERT_FALSE(static_cast<bool>(decoded));
  EXPECT_EQ(decoded.error().code, FormulonErrorCode::kIoXlsbRecordTruncated);
}

TEST(XlsbPtgCodec, DecoderRejectsUnknownPtg) {
  // 0x18 (PtgElfLel) is marked Unsupported in the dispatch table.
  Arena arena;
  const std::vector<std::uint8_t> bytes = {0x18, 0x00};
  ByteSpan span{bytes.data(), bytes.size()};
  auto decoded = decode_ptgs(span, ByteSpan{}, arena, {}, {}, {});
  ASSERT_FALSE(static_cast<bool>(decoded));
  EXPECT_EQ(decoded.error().code, FormulonErrorCode::kIoXlsbUnsupportedPtg);
}

TEST(XlsbPtgCodec, DecoderAcceptsTransparentParenToken) {
  // PtgInt(1), PtgInt(2), PtgAdd, PtgParen. Excel emits the trailing
  // PtgParen for explicit grouping; its value stack is otherwise unchanged.
  const std::vector<std::uint8_t> rgce = {0x1E, 0x01, 0x00, 0x1E, 0x02, 0x00, 0x03, 0x15};
  Arena arena;
  ByteSpan main{rgce.data(), rgce.size()};
  ByteSpan extra{nullptr, 0};
  auto decoded = decode_ptgs(main, extra, arena, {}, {}, {});
  ASSERT_TRUE(static_cast<bool>(decoded));
  EXPECT_EQ(parser::format_formula(*decoded.value()), "1+2");
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
    auto decoded = decode_ptgs(main, extra, arena, {}, {}, {});
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
      decode_ptgs(ByteSpan{rgce.data(), rgce.size()}, ByteSpan{extra.data(), extra.size()}, arena, {}, {}, {});
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
  auto encoded = encode_ptgs(*root, {}, {}, {});
  ASSERT_TRUE(static_cast<bool>(encoded));

  Arena dec_arena;
  ByteSpan rgce{encoded.value().rgce.data(), encoded.value().rgce.size()};
  ByteSpan rgcb{encoded.value().rgcb.data(), encoded.value().rgcb.size()};
  auto decoded = decode_ptgs(rgce, rgcb, dec_arena, {}, {}, {});
  ASSERT_FALSE(static_cast<bool>(decoded));
  EXPECT_EQ(decoded.error().code, FormulonErrorCode::kIoXlsbCorrupt);
}

TEST(XlsbPtgCodec, EncoderUsesExcelClassesForReferenceArgumentsAndFunctionResults) {
  Arena arena;
  parser::Parser p("SUM(A1:A2)", arena);
  parser::AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  ASSERT_TRUE(p.errors().empty());
  auto encoded = encode_ptgs(*root, {}, {}, {});
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
  auto encoded = encode_ptgs(*root, {}, {}, {});
  ASSERT_FALSE(static_cast<bool>(encoded));
  EXPECT_EQ(encoded.error().code, FormulonErrorCode::kIoXlsbUnsupportedPtg);
}

TEST(XlsbPtgCodec, EncoderLowersRegisteredDefinedName) {
  Arena arena;
  parser::Parser p("MyName", arena);
  parser::AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  const NameTable names = {{"MyName", 1U}};
  auto encoded = encode_ptgs(*root, {}, {}, names);
  ASSERT_TRUE(static_cast<bool>(encoded)) << (encoded ? "" : encoded.error().message);

  Arena dec_arena;
  ByteSpan rgce{encoded.value().rgce.data(), encoded.value().rgce.size()};
  const std::vector<XlsbName> name_table = {{-1, "MyName", false}};
  auto decoded = decode_ptgs(rgce, ByteSpan{}, dec_arena, {}, name_table, {});
  ASSERT_TRUE(static_cast<bool>(decoded)) << (decoded ? "" : decoded.error().message);
  EXPECT_EQ(parser::format_formula(*decoded.value()), "MyName");
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
