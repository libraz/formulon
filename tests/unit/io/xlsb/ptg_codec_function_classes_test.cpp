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
      {"IF(A1>0,A2,A3)", {0x44, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x1E, 0x00, 0x00, 0x0D, 0x19, 0x02, 0x0B,
                          0x00, 0x24, 0x01, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x19, 0x08, 0x0E, 0x00, 0x24, 0x02,
                          0x00, 0x00, 0x00, 0x00, 0xC0, 0x19, 0x08, 0x03, 0x00, 0x42, 0x03, 0x01, 0x00}},
  };
  for (const Case& c : cases) {
    EXPECT_EQ(EncodeOnSheet1(c.formula, PtgRootClass::kValue).rgce, c.rgce) << c.formula;
  }
}
TEST(XlsbPtgCodec, ParameterClassTailRepeats) {
  EXPECT_EQ(xlsb_parameter_class("SUMPRODUCT", 7), 'F');
  const char sumifs[] = "RRVRVRV";
  for (std::uint32_t i = 0; i < 7U; ++i) {
    EXPECT_EQ(xlsb_parameter_class("sumifs", i), sumifs[i]) << i;
  }
  EXPECT_EQ(xlsb_parameter_class("ABS", 0), 'V');
  EXPECT_EQ(xlsb_parameter_class("NOT.A.FUNCTION", 3), 'V');
}
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
TEST(XlsbPtgCodec, ArrayEvaluationFollowsTheParameterClasses) {
  for (const char* formula : {"A1:A2", "A1:A2*2", "SUM(A1:A2*2)", "COUNTIF(A1:A2,A1:A2)", "AND(A1:A2>0)",
                              "ISNUMBER(A1:A2)", "LEN(A1:A2)", "IF(A1:A2>0,1,0)", "A1:B2 B2:C3"}) {
    Arena arena;
    parser::Parser p(formula, arena);
    const parser::AstNode* root = p.parse();
    ASSERT_NE(root, nullptr) << formula;
    EXPECT_TRUE(formula_is_dynamic_array(*root, NameShapes())) << formula;
  }
  for (const char* formula : {"A1", "A1+1", "SUM(A1:A2)", "SUMPRODUCT(A1:A2)", "MATCH(1,A1:A2,0)",
                              "VLOOKUP(1,A1:B2,2,0)", "INDEX(A1:A2,1)", "(A1:A2,B1:B2)", "A1 B1", "ROWS(A1:A2)"}) {
    Arena arena;
    parser::Parser p(formula, arena);
    const parser::AstNode* root = p.parse();
    ASSERT_NE(root, nullptr) << formula;
    EXPECT_FALSE(formula_is_dynamic_array(*root, NameShapes())) << formula;
  }
}
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
TEST(XlsbPtgCodec, LegacyFormulaStoresNoImpliedAt) {
  const NameTable names = {{"_xlfn.SINGLE", 2U}};
  auto encode = [&names](const char* formula, PtgEvaluation evaluation) {
    Arena arena;
    parser::Parser p(formula, arena);
    parser::AstNode* root = p.parse();
    EXPECT_NE(root, nullptr) << formula;
    auto encoded = encode_ptgs(*root, {}, {}, names, PtgRootClass::kValue, std::nullopt, evaluation);
    EXPECT_TRUE(static_cast<bool>(encoded)) << formula;
    return encoded ? encoded.value().rgce : std::vector<std::uint8_t>{};
  };
  EXPECT_EQ(encode("SUM(@A1:A2*2)", PtgEvaluation::kLegacy), encode("SUM(A1:A2*2)", PtgEvaluation::kLegacy));
  EXPECT_EQ(encode("@A1:A2", PtgEvaluation::kLegacy), encode("A1:A2", PtgEvaluation::kLegacy));
  EXPECT_NE(encode("SUM(@A1:A2)", PtgEvaluation::kLegacy), encode("SUM(A1:A2)", PtgEvaluation::kLegacy));
  EXPECT_NE(encode("@A1:A2", PtgEvaluation::kLegacyArray), encode("A1:A2", PtgEvaluation::kLegacyArray));
}
TEST(XlsbPtgCodec, BranchingCallsCarryExcelsJumpAttributes) {
  const std::vector<std::uint8_t> a1 = {0x44, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0};
  const std::vector<std::uint8_t> ref_a1 = {0x24, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0};
  auto join = [](std::initializer_list<std::vector<std::uint8_t>> parts) {
    std::vector<std::uint8_t> out;
    for (const auto& part : parts) {
      out.insert(out.end(), part.begin(), part.end());
    }
    return out;
  };
  EXPECT_EQ(EncodeOnSheet1("IF(A1,A1,A1)", PtgRootClass::kValue).rgce,
            join({a1,
                  {0x19, 0x02, 0x0B, 0x00},
                  ref_a1,
                  {0x19, 0x08, 0x0E, 0x00},
                  ref_a1,
                  {0x19, 0x08, 0x03, 0x00, 0x42, 0x03, 0x01, 0x00}}));
  EXPECT_EQ(EncodeOnSheet1("CHOOSE(A1,A1,A1,A1,A1)", PtgRootClass::kValue).rgce,
            join({a1,
                  {0x19, 0x04, 0x04, 0x00, 0x0A, 0x00, 0x15, 0x00, 0x20, 0x00, 0x2B, 0x00, 0x36, 0x00},
                  ref_a1,
                  {0x19, 0x08, 0x24, 0x00},
                  ref_a1,
                  {0x19, 0x08, 0x19, 0x00},
                  ref_a1,
                  {0x19, 0x08, 0x0E, 0x00},
                  ref_a1,
                  {0x19, 0x08, 0x03, 0x00, 0x42, 0x05, 0x64, 0x00}}));
  EXPECT_EQ(EncodeOnSheet1("IFERROR(A1,A1)", PtgRootClass::kValue).rgce,
            join({a1, {0x19, 0x80, 0x07, 0x00}, ref_a1, {0x19, 0x08, 0x02, 0x00, 0x41, 0xE0, 0x01}}));
  EXPECT_EQ(RoundTrip("IF(A1>0,CHOOSE(2,A1,B1),IFERROR(1/0,2))"), "IF(A1>0,CHOOSE(2,A1,B1),IFERROR(1/0,2))");
  EXPECT_EQ(EncodeOnSheet1("IF(A1>0,1)", PtgRootClass::kValue).rgce,
            join({a1,
                  {0x1E, 0x00, 0x00, 0x0D, 0x19, 0x02, 0x07, 0x00, 0x1E, 0x01, 0x00, 0x19, 0x08, 0x03, 0x00, 0x42, 0x02,
                   0x01, 0x00}}));
}
TEST(XlsbPtgCodec, ConditionalFormatArrayOperandsCarryArrayClassInward) {
  auto encode = [](const char* formula) {
    Arena arena;
    parser::Parser p(formula, arena);
    parser::AstNode* root = p.parse();
    EXPECT_NE(root, nullptr) << formula;
    auto encoded =
        encode_ptgs(*root, {"S"}, {}, {}, PtgRootClass::kValue, PtgBaseCell{5U, 7U}, PtgEvaluation::kConditionalFormat);
    EXPECT_TRUE(static_cast<bool>(encoded)) << formula;
    std::vector<std::uint8_t> rgce = encoded ? encoded.value().rgce : std::vector<std::uint8_t>{};
    if (rgce.size() >= 4U && rgce[0] == 0x19 && rgce[1] == 0x01) {
      rgce.erase(rgce.begin() + 2, rgce.begin() + 4);
    }
    return rgce;
  };
  const std::vector<std::uint8_t> ref_n = {0x00, 0x00, 0x00, 0x00, 0x00, 0xC0};
  auto join = [](std::initializer_list<std::vector<std::uint8_t>> parts) {
    std::vector<std::uint8_t> out;
    for (const auto& part : parts) {
      out.insert(out.end(), part.begin(), part.end());
    }
    return out;
  };
  EXPECT_EQ(
      encode("AND(H6<WEEKDAY(TODAY()))"),
      join({{0x19, 0x01, 0x6C}, ref_n, {0x61, 0xDD, 0x00, 0x62, 0x01, 0x46, 0x00, 0x09, 0x42, 0x01, 0x24, 0x00}}));
  EXPECT_EQ(encode("AND(MONTH(H6)=MONTH(TODAY()),YEAR(H6)=YEAR(TODAY()))"),
            join({{0x19, 0x01, 0x6C},
                  ref_n,
                  {0x61, 0x44, 0x00, 0x61, 0xDD, 0x00, 0x61, 0x44, 0x00, 0x0B, 0x6C},
                  ref_n,
                  {0x61, 0x45, 0x00, 0x61, 0xDD, 0x00, 0x61, 0x45, 0x00, 0x0B, 0x42, 0x02, 0x24, 0x00}}));
  EXPECT_EQ(
      encode("H6<WEEKDAY(TODAY())-1"),
      join({{0x19, 0x01, 0x4C}, ref_n, {0x41, 0xDD, 0x00, 0x42, 0x01, 0x46, 0x00, 0x1E, 0x01, 0x00, 0x04, 0x09}}));
}
TEST(XlsbPtgCodec, ReferenceReturningCallsFollowTheirSlot) {
  auto hex = [](std::string_view formula) {
    std::string out;
    for (const std::uint8_t b : EncodeOnSheet1(formula, PtgRootClass::kValue).rgce) {
      char buf[4];
      std::snprintf(buf, sizeof(buf), "%02x ", b);
      out += buf;
    }
    return out;
  };
  const std::string offset = "19 01 00 00 24 00 00 00 00 00 c0 1e 00 00 1e 00 00 ";
  EXPECT_EQ(hex("ROWS(OFFSET(A1,0,0))"), offset + "22 03 4e 00 41 4c 00 ");
  EXPECT_EQ(hex("ISREF(OFFSET(A1,0,0))"), offset + "22 03 4e 00 41 69 00 ");
  EXPECT_EQ(hex("ABS(OFFSET(A1,0,0))"), offset + "42 03 4e 00 41 18 00 ");
  EXPECT_EQ(hex("OFFSET(A1,0,0)"), offset + "42 03 4e 00 ");
  EXPECT_EQ(hex("SUM(INDEX(A1:C3,1,1))"),
            "25 00 00 00 00 02 00 00 00 00 c0 02 c0 1e 01 00 1e 01 00 22 03 1d 00 19 10 00 00 ");
  EXPECT_EQ(
      hex("AREAS(IF(TRUE,A1,B1))"),
      "1d 01 19 02 0b 00 24 00 00 00 00 00 c0 19 08 0e 00 24 00 00 00 00 01 c0 19 08 03 00 22 03 01 00 41 4b 00 ");
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
