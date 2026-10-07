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
TEST(XlsbPtgCodec, WrittenParenthesesRoundTripAsPtgParen) {
  const std::vector<std::uint8_t> excel = {0x19, 0x01, 0x00, 0x00, 0x44, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x41,
                                           0xDD, 0x00, 0x42, 0x01, 0x46, 0x00, 0x15, 0x09, 0x42, 0x01, 0x24, 0x00};
  EXPECT_EQ(EncodeOnSheet1("AND(A1<(WEEKDAY(TODAY())))", PtgRootClass::kValue).rgce, excel);
  Arena arena;
  auto decoded = decode_ptgs(ByteSpan{excel.data(), excel.size()}, ByteSpan{}, arena, {"Sheet1"}, {}, {}, {});
  ASSERT_TRUE(static_cast<bool>(decoded));
  EXPECT_EQ(parser::format_formula(*decoded.value()), "AND(A1<(WEEKDAY(TODAY())))");
  EXPECT_EQ(RoundTrip("((1+2))*3"), "((1+2))*3");
}
TEST(XlsbPtgCodec, SumOverArea) {
  EXPECT_EQ(RoundTrip("SUM(A1:A10)"), "SUM(A1:A10)");
}
TEST(XlsbPtgCodec, WholeColumnAndRowRefsEncodeAsSentinelAreas) {
  EXPECT_EQ(RoundTrip("SUM(A:A)"), "SUM(A:A)");
  EXPECT_EQ(RoundTrip("SUM(1:1)"), "SUM(1:1)");
  EXPECT_EQ(RoundTrip("SUM($A:B)"), "SUM($A:B)");
  EXPECT_EQ(RoundTrip("SUM(1:$2)"), "SUM(1:$2)");
  // Bytes as Excel 365 saved them: the spanned axis is absolute, and a
  // span of columns or rows is one area.
  const std::vector<std::uint8_t> full_col_a = {0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0x0F, 0x00, 0x00, 0x40, 0x00, 0x40};
  std::vector<std::uint8_t> excel = {0x45};
  excel.insert(excel.end(), full_col_a.begin(), full_col_a.end());
  EXPECT_EQ(EncodeOnSheet1("A:A", PtgRootClass::kValue).rgce, excel);
  EXPECT_EQ(EncodeOnSheet1("A:B", PtgRootClass::kValue).rgce,
            (std::vector<std::uint8_t>{0x45, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0x0F, 0x00, 0x00, 0x40, 0x01, 0x40}));
  EXPECT_EQ(EncodeOnSheet1("SUM(1:2)", PtgRootClass::kValue).rgce,
            (std::vector<std::uint8_t>{0x25, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x80, 0xFF, 0xBF,
                                       0x19, 0x10, 0x00, 0x00}));
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

  SheetRangeTable sheet_ranges;
  sheet_ranges.xti = {{0U, 0, 2}};  // Sheet1 (itab 0) : Sheet3 (itab 2)
  auto encoded = encode_ptgs(*node, sheets, sheet_ranges, {}, PtgRootClass::kValue);
  ASSERT_TRUE(static_cast<bool>(encoded)) << (encoded ? "" : encoded.error().message);

  Arena dec_arena;
  ByteSpan rgce{encoded.value().rgce.data(), encoded.value().rgce.size()};
  const std::vector<XlsbSheetRange> decode_ranges = {{0, 2}};
  auto decoded = decode_ptgs(rgce, ByteSpan{}, dec_arena, sheets, {}, decode_ranges, {});
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
  SheetRangeTable sheet_ranges;
  sheet_ranges.xti = {{0U, 0, 2}};
  auto encoded = encode_ptgs(*node, sheets, sheet_ranges, {}, PtgRootClass::kValue);
  ASSERT_TRUE(static_cast<bool>(encoded)) << (encoded ? "" : encoded.error().message);
  Arena dec_arena;
  ByteSpan rgce{encoded.value().rgce.data(), encoded.value().rgce.size()};
  const std::vector<XlsbSheetRange> decode_ranges = {{0, 2}};
  auto decoded = decode_ptgs(rgce, ByteSpan{}, dec_arena, sheets, {}, decode_ranges, {});
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
TEST(XlsbPtgCodec, ConstantArrayNonSquareRowVector) {
  // A 1-row, 3-column array: distinguishes which `PtgExtraArray` u32 is
  // rows vs. cols (a square fixture can't). Encoder writes rows=1,
  // cols=3; if the reader swapped the fields it would either reject the
  // dimension (1x3 -> read as 3x1 needs 3 rows worth of elements, which
  // *is* available here since count is symmetric at 3, but the element
  // shape would transpose) or -- for this asymmetric row/col case --
  // round-trip to `{1;2;3}` instead of `{1,2,3}`.
  EXPECT_EQ(RoundTrip("{1,2,3}"), "{1,2,3}");
}
TEST(XlsbPtgCodec, ConstantArrayNonSquareColumnVector) {
  // The transpose of the above: 3 rows, 1 column.
  EXPECT_EQ(RoundTrip("{1;2;3}"), "{1;2;3}");
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
  // `JIS` is the ja-JP formula-bar spelling of `DBCS`; Excel stores the
  // call as `DBCS` in both containers and has one function id (215) for
  // it. The codec therefore canonicalises the spelling rather than
  // preserving it — the Writer resolves `JIS` to id 215 through the
  // table's alias, and the Reader hands that id back as `DBCS`, which
  // is exactly what Excel's own formula text would say.
  EXPECT_EQ(RoundTrip("JIS(\"A\")"), "DBCS(\"A\")");
  EXPECT_EQ(RoundTrip("DBCS(\"A\")"), "DBCS(\"A\")");
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

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
