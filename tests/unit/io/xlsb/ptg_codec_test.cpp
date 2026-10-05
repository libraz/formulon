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
TEST(XlsbPtgCodec, ArithmeticWithPrecedence) {
  // `A1+B2*3`: PtgRef, PtgRef, PtgInt, PtgMul, PtgAdd.
  EXPECT_EQ(RoundTrip("A1+B2*3"), "A1+B2*3");
}
TEST(XlsbPtgCodec, AttrChooseSkipsU16JumpOffsets) {
  // PtgInt(1), followed by PtgAttrChoose with count=1 and two 16-bit
  // jump offsets. [MS-XLSB] 2.5.98.25 defines rgOffset as an array of
  // 2-byte unsigned integers, which every independent decoder of this
  // record agrees on, not 4-byte -- the control-flow table is irrelevant after
  // XLSB has already selected its cached formula path, but its wire
  // width must be consumed exactly or the following Ptg stream is
  // misaligned.
  const std::vector<std::uint8_t> rgce = {
      0x1E, 0x01, 0x00,  // PtgInt 1
      0x19, 0x04,        // PtgAttrChoose
      0x01, 0x00,        // u16 count=1
      0x00, 0x00,        // jump offset 0 (u16)
      0x04, 0x00,        // jump offset 1 (u16)
  };
  Arena arena;
  auto decoded = decode_ptgs(ByteSpan{rgce.data(), rgce.size()}, {}, arena, {}, {}, {}, {});
  ASSERT_TRUE(static_cast<bool>(decoded)) << (decoded ? "" : decoded.error().message);
  EXPECT_EQ(parser::format_formula(*decoded.value()), "1");
}
TEST(XlsbPtgCodec, AttrChooseWithMultipleBranchesSkipsU16JumpOffsets) {
  // Same as above but cOffset=2 (three CHOOSE branches), followed by a
  // PtgFuncVar(id=100, CHOOSE) to prove the stream realigns correctly
  // onto a real function token after a wider jump table.
  const std::vector<std::uint8_t> rgce = {
      0x1E, 0x02, 0x00,        // PtgInt 2 (index)
      0x1E, 0x0A, 0x00,        // PtgInt 10 (branch 1)
      0x1E, 0x14, 0x00,        // PtgInt 20 (branch 2)
      0x1E, 0x1E, 0x00,        // PtgInt 30 (branch 3)
      0x19, 0x04,              // PtgAttrChoose
      0x02, 0x00,              // u16 count=2 -> 3 offsets
      0x00, 0x00,              // jump offset 0 (u16)
      0x02, 0x00,              // jump offset 1 (u16)
      0x04, 0x00,              // jump offset 2 (u16)
      0x22, 0x04, 0x64, 0x00,  // PtgFuncVar, argc=4, iftab=100 (CHOOSE)
  };
  Arena arena;
  auto decoded = decode_ptgs(ByteSpan{rgce.data(), rgce.size()}, {}, arena, {}, {}, {}, {});
  ASSERT_TRUE(static_cast<bool>(decoded)) << (decoded ? "" : decoded.error().message);
  EXPECT_EQ(parser::format_formula(*decoded.value()), "CHOOSE(2,10,20,30)");
}
TEST(XlsbPtgCodec, NameLocalToAnotherSheetDecodesSheetQualified) {
  const std::vector<std::string> sheets = {"Sheet1", "My Sheet"};
  const std::vector<XlsbName> names = {XlsbName{1, "Local", false}, XlsbName{0, "Own", false},
                                       XlsbName{-1, "Global", false}};
  std::vector<XlsbSheetRange> ranges(1);
  ranges[0].itab_first = 1;
  ranges[0].itab_last = 1;
  struct Case {
    std::vector<std::uint8_t> rgce;
    const char* want;
  };
  const Case cases[] = {
      {{0x23, 0x01, 0x00, 0x00, 0x00}, "'My Sheet'!Local"},
      {{0x23, 0x02, 0x00, 0x00, 0x00}, "Own"},
      {{0x23, 0x03, 0x00, 0x00, 0x00}, "Global"},
      {{0x39, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00}, "'My Sheet'!Local"},
      {{0x39, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00}, "Sheet1!Own"},
      {{0x39, 0x00, 0x00, 0x03, 0x00, 0x00, 0x00}, "[0]!Global"},
  };
  for (const Case& c : cases) {
    Arena arena;
    auto decoded = decode_ptgs(ByteSpan{c.rgce.data(), c.rgce.size()}, {}, arena, sheets, names, ranges, {},
                               /*host_itab=*/0);
    ASSERT_TRUE(static_cast<bool>(decoded)) << (decoded ? "" : decoded.error().message);
    EXPECT_EQ(parser::format_formula(*decoded.value()), c.want);
  }
}
TEST(XlsbPtgCodec, SheetQualifiedNameEncodesPtgNameX) {
  const std::vector<std::string> sheets = {"Sheet1", "Sheet2"};
  NameTable table;
  table.emplace(sheet_scoped_name_key(1, "Local"), 1U);
  table.emplace(sheet_scoped_name_key(1, "Fn"), 2U);
  const std::vector<XlsbName> names = {XlsbName{1, "Local", false}, XlsbName{1, "Fn", false}};
  struct Case {
    const char* formula;
    std::vector<std::uint8_t> want;
  };
  const Case cases[] = {
      {"Sheet2!Local", {0x59, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00}},
      {"Sheet2!Fn(3)", {0x39, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x1E, 0x03, 0x00, 0x42, 0x02, 0xFF, 0x00}},
  };
  for (const Case& c : cases) {
    Arena arena;
    parser::Parser p(c.formula, arena);
    parser::AstNode* root = p.parse();
    ASSERT_NE(root, nullptr) << c.formula;
    SheetRangeTable ranges;
    std::unordered_set<std::uint64_t> seen;
    collect_ptg_sheet_ranges(*root, sheets, ranges, seen);
    ASSERT_EQ(ranges.size(), 1U) << c.formula;
    EXPECT_EQ(ranges[0].first, -2);
    EXPECT_EQ(ranges[0].second, -2);
    auto encoded = encode_ptgs(*root, sheets, ranges, table, PtgRootClass::kValue);
    ASSERT_TRUE(static_cast<bool>(encoded)) << c.formula << " | " << (encoded ? "" : encoded.error().message);
    EXPECT_EQ(encoded.value().rgce, c.want) << c.formula;

    const std::vector<XlsbSheetRange> decode_ranges = {XlsbSheetRange{-2, -2}};
    Arena dec_arena;
    auto decoded = decode_ptgs(ByteSpan{encoded.value().rgce.data(), encoded.value().rgce.size()}, {}, dec_arena,
                               sheets, names, decode_ranges, {}, /*host_itab=*/0);
    ASSERT_TRUE(static_cast<bool>(decoded)) << (decoded ? "" : decoded.error().message);
    EXPECT_EQ(parser::format_formula(*decoded.value()), c.formula);
  }
}
TEST(XlsbPtgCodec, CellReferenceCalleeMatchesExcelBytes) {
  const std::vector<std::string> sheets = {"Sheet1"};
  struct Case {
    const char* formula;
    std::vector<std::uint8_t> want;
  };
  const Case cases[] = {
      {"A1(1)", {0x24, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x1E, 0x01, 0x00, 0x42, 0x02, 0xFF, 0x00}},
      {"Sheet1!LOG10(100)",
       {0x3A, 0x00, 0x00, 0x09, 0x00, 0x00, 0x00, 0x3C, 0xE1, 0x1E, 0x64, 0x00, 0x42, 0x02, 0xFF, 0x00}},
  };
  for (const Case& c : cases) {
    Arena arena;
    parser::Parser p(c.formula, arena);
    parser::AstNode* root = p.parse();
    ASSERT_NE(root, nullptr) << c.formula;
    SheetRangeTable ranges;
    std::unordered_set<std::uint64_t> seen;
    collect_ptg_sheet_ranges(*root, sheets, ranges, seen);
    auto encoded = encode_ptgs(*root, sheets, ranges, {}, PtgRootClass::kValue);
    ASSERT_TRUE(static_cast<bool>(encoded)) << c.formula << " | " << (encoded ? "" : encoded.error().message);
    EXPECT_EQ(encoded.value().rgce, c.want) << c.formula;

    std::vector<XlsbSheetRange> decode_ranges;
    for (const auto& [first, last] : ranges) {
      decode_ranges.push_back(XlsbSheetRange{first, last});
    }
    Arena dec_arena;
    auto decoded = decode_ptgs(ByteSpan{c.want.data(), c.want.size()}, {}, dec_arena, sheets, {}, decode_ranges, {},
                               /*host_itab=*/0);
    ASSERT_TRUE(static_cast<bool>(decoded)) << (decoded ? "" : decoded.error().message);
    EXPECT_EQ(parser::format_formula(*decoded.value()), c.formula);
  }
}
TEST(XlsbPtgCodec, ReferenceOperationsMatchExcelBytes) {
  const std::vector<std::uint8_t> a1a2 = {0x25, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0xC0};
  const std::vector<std::uint8_t> b1b2 = {0x25, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0xC0, 0x01, 0xC0};
  const std::vector<std::uint8_t> a1b2 = {0x25, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x01, 0xC0};
  const std::vector<std::uint8_t> b1c3 = {0x25, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x01, 0xC0, 0x02, 0xC0};
  const std::vector<std::uint8_t> a1 = {0x24, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0};
  const std::vector<std::uint8_t> b1 = {0x24, 0x00, 0x00, 0x00, 0x00, 0x01, 0xC0};
  auto cat = [](std::initializer_list<std::vector<std::uint8_t>> parts) {
    std::vector<std::uint8_t> out;
    for (const auto& part : parts) {
      out.insert(out.end(), part.begin(), part.end());
    }
    return out;
  };
  const std::vector<std::uint8_t> mem_area = {0x00, 0x00, 0x00, 0x00, 0x1B, 0x00};  // unused + cce 27
  const std::vector<std::uint8_t> attr_sum = {0x19, 0x10, 0x00, 0x00};
  // rgcb: rectangle count, then rwFirst, rwLast, colFirst, colLast (u32 each).
  const std::vector<std::uint8_t> rect_b1b2 = {0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00,
                                               0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00};
  struct Case {
    const char* formula;
    std::vector<std::uint8_t> rgce;
    std::vector<std::uint8_t> rgcb;
  };
  const Case cases[] = {
      {"(A1:A2,B1:B2)", cat({{0x47, 0x0F, 0x00, 0x00, 0x00, 0x1B, 0x00}, a1a2, b1b2, {0x10, 0x15}}), {}},
      {"(A1,B1)", cat({{0x47, 0x0F, 0x00, 0x00, 0x00, 0x0F, 0x00}, a1, b1, {0x10, 0x15}}), {}},
      {"A1 B1", cat({{0x47, 0x00, 0x00, 0x00, 0x00, 0x0F, 0x00}, a1, b1, {0x0F}}), {}},
      {"SUM((A1:A2,B1:B2))",
       cat({{0x26}, mem_area, a1a2, b1b2, {0x10, 0x15}, attr_sum}),
       {0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00}},
      {"A1:B2 B1:C3", cat({{0x46}, mem_area, a1b2, b1c3, {0x0F}}), rect_b1b2},
      {"SUM(A1:B2 B1:C3)", cat({{0x26}, mem_area, a1b2, b1c3, {0x0F}, attr_sum}), rect_b1b2},
  };
  for (const Case& c : cases) {
    const EncodedFormula encoded = EncodeOnSheet1(c.formula, PtgRootClass::kValue);
    EXPECT_EQ(encoded.rgce, c.rgce) << c.formula;
    EXPECT_EQ(encoded.rgcb, c.rgcb) << c.formula;
  }
}
TEST(XlsbPtgCodec, NameBodyReferenceOperationsMatchExcelBytes) {
  const std::vector<std::uint8_t> a1a2 = {0x3B, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01,
                                          0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00};
  const std::vector<std::uint8_t> b1b2 = {0x3B, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01,
                                          0x00, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00};
  const std::vector<std::uint8_t> a1b2 = {0x3B, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01,
                                          0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00};
  const std::vector<std::uint8_t> b1c3 = {0x3B, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02,
                                          0x00, 0x00, 0x00, 0x01, 0x00, 0x02, 0x00};
  std::vector<std::uint8_t> union_want = {0x29, 0x1F, 0x00};
  for (const auto* part : {&a1a2, &b1b2}) {
    union_want.insert(union_want.end(), part->begin(), part->end());
  }
  union_want.insert(union_want.end(), {0x10, 0x15});
  std::vector<std::uint8_t> isect_want = {0x29, 0x1F, 0x00};
  for (const auto* part : {&a1b2, &b1c3}) {
    isect_want.insert(isect_want.end(), part->begin(), part->end());
  }
  isect_want.push_back(0x0F);
  EXPECT_EQ(EncodeOnSheet1("(Sheet1!$A$1:$A$2,Sheet1!$B$1:$B$2)", PtgRootClass::kReference).rgce, union_want);
  EXPECT_EQ(EncodeOnSheet1("Sheet1!$A$1:$B$2 Sheet1!$B$1:$C$3", PtgRootClass::kReference).rgce, isect_want);
}
TEST(XlsbPtgCodec, ParenthesesBecomePtgParen) {
  struct Case {
    const char* formula;
    std::vector<std::uint8_t> rgce;
  };
  const Case cases[] = {
      {"(1+2)*3", {0x1E, 0x01, 0x00, 0x1E, 0x02, 0x00, 0x03, 0x15, 0x1E, 0x03, 0x00, 0x05}},
      {"-(1+2)", {0x1E, 0x01, 0x00, 0x1E, 0x02, 0x00, 0x03, 0x15, 0x13}},
      {"2^(1+1)", {0x1E, 0x02, 0x00, 0x1E, 0x01, 0x00, 0x1E, 0x01, 0x00, 0x03, 0x15, 0x07}},
      {"(A1:A2)(1)", {0x25, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0xC0,
                      0x00, 0xC0, 0x15, 0x1E, 0x01, 0x00, 0x42, 0x02, 0xFF, 0x00}},
  };
  for (const Case& c : cases) {
    EXPECT_EQ(EncodeOnSheet1(c.formula, PtgRootClass::kValue).rgce, c.rgce) << c.formula;
    EXPECT_EQ(RoundTrip(c.formula), c.formula);
  }
}
TEST(XlsbPtgCodec, EveryBuiltinHasAnXlsbEncoding) {
  std::vector<std::string> missing;
  auto check = [&missing](std::string_view name) {
    const std::string_view stored = canonical_function_name(name);
    if (lookup_func_by_name(stored) == nullptr && !xlsb_uses_hidden_name(stored) &&
        classify_storage_prefix(stored) == parser::StoragePrefixKind::None) {
      missing.emplace_back(name);
    }
  };
  eval::default_registry().for_each_name(
      [](std::string_view name, void* ctx) { (*static_cast<decltype(check)*>(ctx))(name); }, &check);
  for (const char* const* p = eval::lazy_form_names(); *p != nullptr; ++p) {
    check(*p);
  }
  EXPECT_TRUE(missing.empty()) << (missing.empty() ? "" : missing.front());
}
TEST(XlsbPtgCodec, RelativeReferencesAgainstABaseCellMatchExcelBytes) {
  struct Case {
    const char* formula;
    PtgBaseCell base;
    std::vector<std::uint8_t> rgce;
  };
  const Case cases[] = {
      // sqref "E2:E10 G3:G5": base E2.
      {"E2>F$1", {1U, 4U}, {0x4C, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x4C, 0x00, 0x00, 0x00, 0x00, 0x01, 0x40, 0x0D}},
      // sqref "C4:D6": base C4, so A1:B2 sits at negative offsets; a fully
      // absolute area stays `PtgArea`.
      {"SUM(A1:B2)>$A$1:$B$2", {3U, 2U}, {0x2D, 0xFD, 0xFF, 0x0F, 0x00, 0xFE, 0xFF, 0x0F, 0x00, 0xFE, 0xFF,
                                          0xFF, 0xFF, 0x19, 0x10, 0x00, 0x00, 0x45, 0x00, 0x00, 0x00, 0x00,
                                          0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x0D}},
  };
  for (const Case& c : cases) {
    Arena arena;
    parser::Parser p(c.formula, arena);
    parser::AstNode* root = p.parse();
    ASSERT_NE(root, nullptr) << c.formula;
    auto encoded = encode_ptgs(*root, {}, {}, {}, PtgRootClass::kValue, c.base);
    ASSERT_TRUE(static_cast<bool>(encoded)) << c.formula << " | " << (encoded ? "" : encoded.error().message);
    EXPECT_EQ(encoded.value().rgce, c.rgce) << c.formula;

    Arena dec_arena;
    auto decoded = decode_ptgs(ByteSpan{c.rgce.data(), c.rgce.size()}, {}, dec_arena, {}, {}, {}, {}, -1, c.base);
    ASSERT_TRUE(static_cast<bool>(decoded)) << c.formula << " | " << (decoded ? "" : decoded.error().message);
    EXPECT_EQ(parser::format_formula(*decoded.value()), c.formula);
  }
}
TEST(XlsbPtgCodec, RelativeOffsetsWrapAroundTheGrid) {
  struct Case {
    const char* formula;
    PtgBaseCell base;
    std::vector<std::uint8_t> rgce;
  };
  const Case cases[] = {
      {"XFD1048576", {0U, 0U}, {0x4C, 0xFF, 0xFF, 0x0F, 0x00, 0xFF, 0xFF}},
      {"A1", {1048575U, 16383U}, {0x4C, 0x01, 0x00, 0x00, 0x00, 0x01, 0xC0}},
      {"$A1", {1048575U, 16383U}, {0x4C, 0x01, 0x00, 0x00, 0x00, 0x00, 0x80}},
      {"A$1", {1048575U, 16383U}, {0x4C, 0x00, 0x00, 0x00, 0x00, 0x01, 0x40}},
      {"A1:B2", {1U, 1U}, {0x4D, 0xFF, 0xFF, 0x0F, 0x00, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0x00, 0xC0}},
  };
  for (const Case& c : cases) {
    Arena arena;
    parser::Parser p(c.formula, arena);
    parser::AstNode* root = p.parse();
    ASSERT_NE(root, nullptr) << c.formula;
    auto encoded = encode_ptgs(*root, {}, {}, {}, PtgRootClass::kValue, c.base);
    ASSERT_TRUE(static_cast<bool>(encoded)) << c.formula;
    EXPECT_EQ(encoded.value().rgce, c.rgce) << c.formula;
    Arena dec_arena;
    auto decoded = decode_ptgs(ByteSpan{c.rgce.data(), c.rgce.size()}, {}, dec_arena, {}, {}, {}, {}, -1, c.base);
    ASSERT_TRUE(static_cast<bool>(decoded)) << c.formula;
    EXPECT_EQ(parser::format_formula(*decoded.value()), c.formula);
  }
}
TEST(XlsbPtgCodec, RelativeTokensNeedABaseCell) {
  EXPECT_EQ(EncodeOnSheet1("A1", PtgRootClass::kValue).rgce,
            (std::vector<std::uint8_t>{0x44, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0}));
  const std::vector<std::uint8_t> ref_n = {0x4C, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0};
  Arena arena;
  auto decoded = decode_ptgs(ByteSpan{ref_n.data(), ref_n.size()}, {}, arena, {}, {}, {}, {});
  ASSERT_FALSE(static_cast<bool>(decoded));
  EXPECT_EQ(decoded.error().code, FormulonErrorCode::kIoXlsbUnsupportedPtg);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
