//
// Flat operator chains through the Ptg codec and a whole `.xlsb` save and
// load. The token stream is postfix, so `1+2+3` is `1 2 + 3 +`; the decoder
// folds such a run back into one `BinaryOp` node, including when a tighter
// operand sits between two of its operators (`1 2 + 3 4 * +`).

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "cell.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/xlsb/ptg_reader.h"
#include "io/xlsb/ptg_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/parser.h"
#include "ptg_codec_test_helpers.h"
#include "sheet.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

using ptg_codec_test_support::RoundTrip;

// `A1+A2+...+An`, without the leading `=`.
std::string RefChain(std::uint32_t terms) {
  std::string out = "A1";
  for (std::uint32_t i = 2; i <= terms; ++i) {
    out.append("+A").append(std::to_string(i));
  }
  return out;
}

// `A1+A2+...+A9+A1+...` with `terms` operands: 2,700 of them fit in
// 8,192 characters.
std::string ShortRefChain(std::uint32_t terms) {
  std::string out = "A1";
  for (std::uint32_t i = 1; i < terms; ++i) {
    out.append("+A").append(std::to_string(i % 9U + 1U));
  }
  return out;
}

std::string MixedChain(std::uint32_t products) {
  std::string out = "1";
  for (std::uint32_t i = 0; i < products; ++i) {
    out.append("+").append(std::to_string(2 * i + 2)).append("*").append(std::to_string(2 * i + 3));
  }
  return out;
}

TEST(XlsbOperatorChain, DecoderFoldsAPostfixRunIntoOneNode) {
  // 1 2 + 3 4 * + : PtgInt 1, PtgInt 2, PtgAdd, PtgInt 3, PtgInt 4, PtgMul, PtgAdd.
  const std::vector<std::uint8_t> rgce = {0x1E, 1, 0, 0x1E, 2, 0, 0x03, 0x1E, 3, 0, 0x1E, 4, 0, 0x05, 0x03};
  Arena arena;
  auto decoded = decode_ptgs(ByteSpan{rgce.data(), rgce.size()}, ByteSpan{}, arena, {}, {}, {}, {});
  ASSERT_TRUE(static_cast<bool>(decoded)) << decoded.error().message;
  const parser::AstNode& root = *decoded.value();
  ASSERT_EQ(root.kind(), parser::NodeKind::BinaryOp);
  ASSERT_EQ(root.as_binary_arity(), 3U);
  EXPECT_EQ(root.as_binary_operand(2).kind(), parser::NodeKind::BinaryOp);
  EXPECT_EQ(parser::format_formula(root), "1+2+3*4");
}

TEST(XlsbOperatorChain, DecoderKeepsAnUnparenthesisedRightChainApart) {
  // 1 2 3 - - : the right operand is its own chain, printed in parentheses.
  const std::vector<std::uint8_t> rgce = {0x1E, 1, 0, 0x1E, 2, 0, 0x1E, 3, 0, 0x04, 0x04};
  Arena arena;
  auto decoded = decode_ptgs(ByteSpan{rgce.data(), rgce.size()}, ByteSpan{}, arena, {}, {}, {}, {});
  ASSERT_TRUE(static_cast<bool>(decoded)) << decoded.error().message;
  EXPECT_EQ(decoded.value()->as_binary_arity(), 2U);
  EXPECT_EQ(parser::format_formula(*decoded.value()), "1-(2-3)");
}

TEST(XlsbOperatorChain, ParenthesisedOperandStaysItsOwnNode) {
  EXPECT_EQ(RoundTrip("(1+2)+3-(4-5)"), "(1+2)+3-(4-5)");
}

TEST(XlsbOperatorChain, LongChainsRoundTrip) {
  for (const std::string& formula : {RefChain(150U), RefChain(1000U), ShortRefChain(2700U), MixedChain(700U)}) {
    SCOPED_TRACE(formula.size());
    ASSERT_LT(formula.size(), 8192U);
    EXPECT_EQ(RoundTrip(formula), formula);
  }
}

TEST(XlsbOperatorChain, WhitespaceBeforeEachOperatorRoundTrips) {
  const std::string formula = "A1 +A2  -A3\n+A4";
  Arena arena;
  parser::Parser p(formula, arena);
  parser::AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  ASSERT_TRUE(p.errors().empty());
  auto encoded = encode_ptgs(*root, {}, {}, {}, PtgRootClass::kValue, std::nullopt, PtgEvaluation::kDynamicArray,
                             NameShapes(), formula);
  ASSERT_TRUE(static_cast<bool>(encoded)) << encoded.error().message;
  Arena dec_arena;
  parser::FormulaWhitespace whitespace;
  auto decoded = decode_ptgs(ByteSpan{encoded.value().rgce.data(), encoded.value().rgce.size()},
                             ByteSpan{encoded.value().rgcb.data(), encoded.value().rgcb.size()}, dec_arena, {}, {}, {},
                             {}, -1, std::nullopt, &whitespace);
  ASSERT_TRUE(static_cast<bool>(decoded)) << decoded.error().message;
  ASSERT_EQ(decoded.value()->as_binary_arity(), 4U);
  EXPECT_EQ(parser::format_formula(*decoded.value(), &whitespace), formula);
}

TEST(XlsbOperatorChain, WorkbookSaveAndLoadKeepsALongChain) {
  const std::string formula = "=" + RefChain(1000U);
  Workbook wb = Workbook::create();
  for (std::uint32_t i = 0; i < 1000U; ++i) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, i, 0U, Value::number(static_cast<double>(i + 1U)))));
  }
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, formula)));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  auto saved = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  auto loaded = read_xlsb(ByteSpan{saved.value().data(), saved.value().size()});
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  Workbook& back = loaded.value().workbook;
  const Cell* cell = back.sheet(0).cell_at(0U, 1U);
  ASSERT_NE(cell, nullptr);
  EXPECT_EQ(cell->formula_text, formula);
  ASSERT_TRUE(static_cast<bool>(back.set_cell_value(0U, 0U, 0U, Value::number(0.0))));
  ASSERT_TRUE(static_cast<bool>(back.recalc(eval::default_registry())));
  const Cell* recalculated = back.sheet(0).cell_at(0U, 1U);
  ASSERT_NE(recalculated, nullptr);
  ASSERT_TRUE(recalculated->cached_value.is_number());
  EXPECT_EQ(recalculated->cached_value.as_number(), 1000.0 * 1001.0 / 2.0 - 1.0);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
