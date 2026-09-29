//
// Round-trip tests for the MS-XLSB Ptg codec (encoder + decoder).
//
// Each case parses an A1 formula to the engine AST, encodes it to a Ptg
// (`rgce`) byte stream, decodes that stream back to an AST, and asserts
// the re-formatted formula text matches the original. This exercises the
// encode <-> decode pair as a single round-trip without needing a full
// xlsb package.

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

  auto encoded = encode_ptgs(*root, sheet_names, sheet_ranges, {}, PtgRootClass::kValue);
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
  auto decoded = decode_ptgs(rgce, rgcb, dec_arena, sheet_names, {}, decode_ranges, {});
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

// A `BrtName` record local to a sheet other than the formula's own is only
// reachable as `Sheet!Name`; one local to the host sheet, or workbook
// scoped, stays unqualified. `PtgNameX` through this workbook's own
// ExternSheet entry is the qualified spelling, so it names a local record
// with its sheet even on that sheet (Excel stores `Sheet1!Fn` on Sheet1 so).
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

// Excel 365 writes `Sheet2!Local` as `PtgNameX` through a sheetless
// book-scope ExternSheet entry (value class at a cell formula's root), and
// `Sheet2!Fn(3)` as the reference-class name-ref, the argument and
// `PtgFuncVar(255)`; both read back with the qualifier.
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

// A cell reference invoked as a callee is the reference-class `PtgRef` /
// `PtgRef3d`, the arguments and `PtgFuncVar(255)`; bytes as Excel 365
// saved `A1(1)` and `Sheet1!LOG10(100)`.
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

// Encodes `formula` on a one-sheet workbook and returns `rgce` + `rgcb`.
EncodedFormula EncodeOnSheet1(std::string_view formula, PtgRootClass root_class) {
  const std::vector<std::string> sheets = {"Sheet1"};
  Arena arena;
  parser::Parser p(formula, arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << formula;
  EXPECT_TRUE(p.errors().empty()) << formula;
  if (root == nullptr) {
    return {};
  }
  SheetRangeTable ranges;
  std::unordered_set<std::uint64_t> seen;
  collect_ptg_sheet_ranges(*root, sheets, ranges, seen);
  auto encoded = encode_ptgs(*root, sheets, ranges, {}, root_class);
  EXPECT_TRUE(static_cast<bool>(encoded)) << formula << " | " << (encoded ? "" : encoded.error().message);
  return encoded ? encoded.value() : EncodedFormula{};
}

// Bytes as Excel 365 saved each formula (backup/oracle_probe/root_ops). A
// union or intersection of plain references in a cell sits behind the memory
// token carrying Excel's precomputed result; `PtgMemArea`'s leading four
// bytes are unused, which Excel leaves uninitialised and this writes as zero.
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

// A defined name's body caches nothing: its root union or intersection sits
// behind reference-class `PtgMemFunc` (Excel's `UName` / `IName` records).
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

// Excel renders a formula from its tokens, so every parenthesis the text
// needs is a `PtgParen` right after the subexpression it closes; without
// them Excel showed `(1+2)*3` as `1+2*3`. `(A1:A2)(1)` is Excel's bytes.
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

// The writer stores a callee with no function id and no hidden-name route
// as a defined-name call, which is right only for names no one defined; so
// every built-in has to have one of the two. LET / LAMBDA are lowered by
// their own AST shapes.
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

// `PtgRefN` / `PtgAreaN` against a base cell, with the bytes Excel 365 saved
// for tests/fixtures/excel/xlsb_feature_rel.xlsb's conditional formats. The
// base is the top-left of the sqref's bounding box.
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

// A relative offset is taken modulo the grid both ways, so a reference
// across the grid's edge wraps: from A1 the last cell is one step up and
// left, and from the last cell A1 is one step down and right.
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

// Without a base cell nothing changes: relative references stay `PtgRef`,
// and a `PtgRefN` has nothing to resolve against.
TEST(XlsbPtgCodec, RelativeTokensNeedABaseCell) {
  EXPECT_EQ(EncodeOnSheet1("A1", PtgRootClass::kValue).rgce,
            (std::vector<std::uint8_t>{0x44, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0}));
  const std::vector<std::uint8_t> ref_n = {0x4C, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0};
  Arena arena;
  auto decoded = decode_ptgs(ByteSpan{ref_n.data(), ref_n.size()}, {}, arena, {}, {}, {}, {});
  ASSERT_FALSE(static_cast<bool>(decoded));
  EXPECT_EQ(decoded.error().code, FormulonErrorCode::kIoXlsbUnsupportedPtg);
}

// A cell formula calling a volatile function opens with `PtgAttrSemi`;
// bytes as Excel 365 saved them (backup/oracle_probe/volatile). Without it
// Excel did not recalculate an engine-written =RAND() on F9.
TEST(XlsbPtgCodec, VolatileFormulasOpenWithAttrSemi) {
  struct Case {
    const char* formula;
    std::vector<std::uint8_t> rgce;
  };
  const Case cases[] = {
      {"RAND()", {0x19, 0x01, 0x00, 0x00, 0x41, 0x3F, 0x00}},
      {"RAND()+A1", {0x19, 0x01, 0x00, 0x00, 0x41, 0x3F, 0x00, 0x44, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x03}},
      {"SUM(A1,RAND())",
       {0x19, 0x01, 0x00, 0x00, 0x24, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x41, 0x3F, 0x00, 0x42, 0x02, 0x04, 0x00}},
      {"OFFSET(A1,0,0)", {0x19, 0x01, 0x00, 0x00, 0x24, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0,
                          0x1E, 0x00, 0x00, 0x1E, 0x00, 0x00, 0x42, 0x03, 0x4E, 0x00}},
      {"A1+1", {0x44, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x1E, 0x01, 0x00, 0x03}},
  };
  for (const Case& c : cases) {
    EXPECT_EQ(EncodeOnSheet1(c.formula, PtgRootClass::kValue).rgce, c.rgce) << c.formula;
    EXPECT_EQ(RoundTrip(c.formula), c.formula);
  }
  // A defined name's body is marked the same way (Excel's `MyNow` = NOW()).
  const std::vector<std::uint8_t> name_body = EncodeOnSheet1("NOW()", PtgRootClass::kReference).rgce;
  ASSERT_GE(name_body.size(), 4U);
  EXPECT_EQ(std::vector<std::uint8_t>(name_body.begin(), name_body.begin() + 4),
            (std::vector<std::uint8_t>{0x19, 0x01, 0x00, 0x00}));
}

// A reference argument takes the class of the parameter it feeds, and an
// operator's operand the class that parameter gives it; bytes as Excel 365
// saved them (backup/oracle_probe/param_class). A reference-class operand
// under IF made Excel compute =IF(A1>0,A2,A3) as #VALUE!, a value-class one
// under SUMPRODUCT showed as =SUMPRODUCT(@A1:A2*2).
TEST(XlsbPtgCodec, ArgumentClassesFollowTheParameter) {
  struct Case {
    const char* formula;
    std::vector<std::uint8_t> rgce;
  };
  const Case cases[] = {
      {"ISNUMBER(A1:A3)",
       {0x45, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0xC0, 0x41, 0x80, 0x00}},
      {"INDEX(A1:A3,2)", {0x25, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x00,
                          0xC0, 0x00, 0xC0, 0x1E, 0x02, 0x00, 0x42, 0x02, 0x1D, 0x00}},
      {"ABS(A1)", {0x44, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x41, 0x18, 0x00}},
      {"SUMPRODUCT(A1:A2*2)", {0x65, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0xC0,
                               0x00, 0xC0, 0x1E, 0x02, 0x00, 0x05, 0x42, 0x01, 0xE4, 0x00}},
      {"VLOOKUP(2,A1:A3,1,0)", {0x1E, 0x02, 0x00, 0x25, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x00,
                                0xC0, 0x00, 0xC0, 0x1E, 0x01, 0x00, 0x1E, 0x00, 0x00, 0x42, 0x04, 0x66, 0x00}},
      // Excel also writes PtgAttrIf / PtgAttrGoto around IF's branches, which
      // this encoder leaves out (no observable difference).
      {"IF(A1>0,A2,A3)", {0x44, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x1E, 0x00, 0x00, 0x0D, 0x24, 0x01, 0x00, 0x00,
                          0x00, 0x00, 0xC0, 0x24, 0x02, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x42, 0x03, 0x01, 0x00}},
  };
  for (const Case& c : cases) {
    EXPECT_EQ(EncodeOnSheet1(c.formula, PtgRootClass::kValue).rgce, c.rgce) << c.formula;
  }
}

// The class list's tail repeats its last letter, or its last two when marked.
TEST(XlsbPtgCodec, ParameterClassTailRepeats) {
  EXPECT_EQ(xlsb_parameter_class("SUMPRODUCT", 7), 'F');
  const char sumifs[] = "RRVRVRV";
  for (std::uint32_t i = 0; i < 7U; ++i) {
    EXPECT_EQ(xlsb_parameter_class("sumifs", i), sumifs[i]) << i;
  }
  EXPECT_EQ(xlsb_parameter_class("ABS", 0), 'V');
  EXPECT_EQ(xlsb_parameter_class("NOT.A.FUNCTION", 3), 'V');
}

// A parenthesised union or intersection called as a function (#REF! in
// Excel) is stored like any other reference operation there: reference-class
// `PtgMemArea` caching its rectangles, `PtgParen`, the arguments and
// `PtgFuncVar(255)`; bytes as Excel 365 saved them (backup/oracle_probe/callee2).
TEST(XlsbPtgCodec, ReferenceOperationCalleeMatchesExcelBytes) {
  const EncodedFormula union_call = EncodeOnSheet1("(A1,B1)(1)", PtgRootClass::kValue);
  EXPECT_EQ(union_call.rgce, (std::vector<std::uint8_t>{0x26, 0x00, 0x00, 0x00, 0x00, 0x0F, 0x00, 0x24, 0x00, 0x00,
                                                        0x00, 0x00, 0x00, 0xC0, 0x24, 0x00, 0x00, 0x00, 0x00, 0x01,
                                                        0xC0, 0x10, 0x15, 0x1E, 0x01, 0x00, 0x42, 0x02, 0xFF, 0x00}));
  EXPECT_EQ(union_call.rgcb,
            (std::vector<std::uint8_t>{0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                                       0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                                       0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00}));
  const EncodedFormula isect_call = EncodeOnSheet1("(A1:B2 B1:B3)(1)", PtgRootClass::kValue);
  EXPECT_EQ(isect_call.rgce, (std::vector<std::uint8_t>{
                                 0x26, 0x00, 0x00, 0x00, 0x00, 0x1B, 0x00, 0x25, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00,
                                 0x00, 0x00, 0x00, 0xC0, 0x01, 0xC0, 0x25, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00,
                                 0x00, 0x01, 0xC0, 0x01, 0xC0, 0x0F, 0x15, 0x1E, 0x01, 0x00, 0x42, 0x02, 0xFF, 0x00}));
  EXPECT_EQ(isect_call.rgcb, (std::vector<std::uint8_t>{0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00,
                                                        0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00}));
  EXPECT_EQ(RoundTrip("(A1,B1)(1)"), "(A1,B1)(1)");
  EXPECT_EQ(RoundTrip("(A1:B2 B1:B3)(1)"), "(A1:B2 B1:B3)(1)");
}

// Formulas Excel 365 marks dynamic-array on entry because an area is
// evaluated as an array, and ones it does not (backup/oracle_probe/dyn_flag).
TEST(XlsbPtgCodec, ArrayEvaluationFollowsTheParameterClasses) {
  for (const char* formula : {"A1:A2", "A1:A2*2", "SUM(A1:A2*2)", "COUNTIF(A1:A2,A1:A2)", "AND(A1:A2>0)",
                              "ISNUMBER(A1:A2)", "LEN(A1:A2)", "IF(A1:A2>0,1,0)", "A1:B2 B2:C3"}) {
    Arena arena;
    parser::Parser p(formula, arena);
    const parser::AstNode* root = p.parse();
    ASSERT_NE(root, nullptr) << formula;
    EXPECT_TRUE(formula_uses_array_evaluation(*root)) << formula;
  }
  for (const char* formula : {"A1", "A1+1", "SUM(A1:A2)", "SUMPRODUCT(A1:A2)", "MATCH(1,A1:A2,0)",
                              "VLOOKUP(1,A1:B2,2,0)", "INDEX(A1:A2,1)", "(A1:A2,B1:B2)", "A1 B1", "ROWS(A1:A2)"}) {
    Arena arena;
    parser::Parser p(formula, arena);
    const parser::AstNode* root = p.parse();
    ASSERT_NE(root, nullptr) << formula;
    EXPECT_FALSE(formula_uses_array_evaluation(*root)) << formula;
  }
}

// A legacy formula (no dynamic-array mark) intersects an operator's area
// operand, so SUM(A1:A2*2) stores it value class where a dynamic-array one
// stores it array class; bytes as Excel 365 saved a legacy workbook's E1 and
// its CSE block {=A1:A2*2} (backup/oracle_probe/dyn_write/legacy_excel.xlsb).
TEST(XlsbPtgCodec, LegacyFormulaIntersectsOperatorOperands) {
  const std::vector<std::uint8_t> area = {0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0xC0};
  auto encode = [](const char* formula, PtgEvaluation evaluation) {
    Arena arena;
    parser::Parser p(formula, arena);
    parser::AstNode* root = p.parse();
    EXPECT_NE(root, nullptr) << formula;
    auto encoded = encode_ptgs(*root, {}, {}, {}, PtgRootClass::kValue, std::nullopt, evaluation);
    EXPECT_TRUE(static_cast<bool>(encoded)) << formula;
    return encoded ? encoded.value().rgce : std::vector<std::uint8_t>{};
  };
  auto with_class = [&area](std::uint8_t area_ptg, std::initializer_list<std::uint8_t> tail) {
    std::vector<std::uint8_t> out = {area_ptg};
    out.insert(out.end(), area.begin(), area.end());
    out.insert(out.end(), tail);
    return out;
  };
  EXPECT_EQ(encode("SUM(A1:A2*2)", PtgEvaluation::kLegacy),
            with_class(0x45, {0x1E, 0x02, 0x00, 0x05, 0x19, 0x10, 0x00, 0x00}));
  EXPECT_EQ(encode("A1:A2*2", PtgEvaluation::kLegacy), with_class(0x45, {0x1E, 0x02, 0x00, 0x05}));
  EXPECT_EQ(encode("SUM(A1:A2*2)", PtgEvaluation::kDynamicArray),
            with_class(0x65, {0x1E, 0x02, 0x00, 0x05, 0x19, 0x10, 0x00, 0x00}));
}

// A written `@` is a call to the hidden `_xlfn.SINGLE` name in either
// evaluation mode; bytes as Excel 365 saved `=@A1`
// (backup/oracle_probe/legacy_at/typed_at.xlsb C5).
TEST(XlsbPtgCodec, WrittenAtStoresAsSingleCall) {
  const NameTable names = {{"_xlfn.SINGLE", 2U}};
  const std::vector<std::uint8_t> expected = {0x23, 0x02, 0x00, 0x00, 0x00, 0x24, 0x00, 0x00,
                                              0x00, 0x00, 0x00, 0xC0, 0x42, 0x02, 0xFF, 0x00};
  for (const PtgEvaluation evaluation : {PtgEvaluation::kLegacy, PtgEvaluation::kDynamicArray}) {
    Arena arena;
    parser::Parser p("@A1", arena);
    parser::AstNode* root = p.parse();
    ASSERT_NE(root, nullptr);
    auto encoded = encode_ptgs(*root, {}, {}, names, PtgRootClass::kValue, std::nullopt, evaluation);
    ASSERT_TRUE(static_cast<bool>(encoded)) << encoded.error().message;
    EXPECT_EQ(encoded.value().rgce, expected);
  }
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
  const SheetRangeTable sheet_ranges = {{0, 2}};
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

TEST(XlsbPtgCodec, PtgArrayCoversNumericElementsOnlyAndSaysSoBothWays) {
  // The element tag preceding each `SerAr` value has only been verified
  // for `0x00` (number). Guessing at the string / bool / error layouts
  // would risk a silently wrong array constant, so both directions
  // refuse them -- and refuse the same set, which is what makes the
  // classification a `Partial` round-trip rather than a one-sided gap.
  const std::vector<std::uint8_t> rgce = {
      0x60,                                                  // PtgArray (array-class)
      0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,  // 14-byte placeholder
      0x00, 0x00, 0x00, 0x00, 0x00,
  };
  std::vector<std::uint8_t> rgcb;
  emit_u32(rgcb, 1U);    // DRw
  emit_u32(rgcb, 1U);    // DCol
  rgcb.push_back(0x01);  // SerAr string tag: not verified, not decoded
  emit_u32(rgcb, 1U);
  rgcb.push_back('a');
  rgcb.push_back(0x00);

  Arena arena;
  auto decoded =
      decode_ptgs(ByteSpan{rgce.data(), rgce.size()}, ByteSpan{rgcb.data(), rgcb.size()}, arena, {}, {}, {}, {});
  ASSERT_FALSE(static_cast<bool>(decoded));
  EXPECT_EQ(decoded.error().code, FormulonErrorCode::kIoXlsbUnsupportedPtg);

  Arena enc_arena;
  parser::Parser parser_with_text("{1,\"a\"}", enc_arena);
  parser::AstNode* root = parser_with_text.parse();
  ASSERT_NE(root, nullptr);
  ASSERT_TRUE(parser_with_text.errors().empty());
  auto encoded = encode_ptgs(*root, {}, {}, {}, PtgRootClass::kValue);
  ASSERT_FALSE(static_cast<bool>(encoded));
  EXPECT_EQ(encoded.error().code, FormulonErrorCode::kIoXlsbUnsupportedPtg);

  // The covered half is genuinely covered: an all-numeric constant still
  // round-trips, so the classification is partial and not unsupported.
  EXPECT_EQ(RoundTrip("SUM({1,2;3,4})"), "SUM({1,2;3,4})");
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

TEST(XlsbPtgCodec, DecoderAcceptsTransparentParenToken) {
  // PtgInt(1), PtgInt(2), PtgAdd, PtgParen. Excel emits the trailing
  // PtgParen for explicit grouping; its value stack is otherwise unchanged.
  const std::vector<std::uint8_t> rgce = {0x1E, 0x01, 0x00, 0x1E, 0x02, 0x00, 0x03, 0x15};
  Arena arena;
  ByteSpan main{rgce.data(), rgce.size()};
  ByteSpan extra{nullptr, 0};
  auto decoded = decode_ptgs(main, extra, arena, {}, {}, {}, {});
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
