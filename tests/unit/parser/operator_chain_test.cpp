//
// A flat run of binary operators sharing one binding power (`1+2-3+...`) is
// one `BinaryOp` node, so its length is bounded by the 8,192-character
// formula limit rather than by the nesting cap. These tests pin the node
// shape, the formatter and R1C1 round trips and reference shifting over
// chains at both ends of that range.

#include <cstdint>
#include <string>
#include <string_view>

#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/ast_dump.h"
#include "parser/ast_format.h"
#include "parser/ast_format_r1c1.h"
#include "parser/ast_shift.h"
#include "parser/parse_error.h"
#include "parser/parser.h"
#include "parser/ref_transforms.h"
#include "utils/arena.h"

namespace formulon {
namespace parser {
namespace {

// `=A1+A2+...+An`.
std::string RefChain(std::uint32_t terms) {
  std::string out = "=A1";
  for (std::uint32_t i = 2; i <= terms; ++i) {
    out.append("+A").append(std::to_string(i));
  }
  return out;
}

// `=A1+A2+...+A9+A1+...` with `terms` operands: 2,700 of them fit in
// 8,192 characters.
std::string ShortRefChain(std::uint32_t terms) {
  std::string out = "=A1";
  for (std::uint32_t i = 1; i < terms; ++i) {
    out.append("+A").append(std::to_string(i % 9U + 1U));
  }
  return out;
}

// `=1+1+...+1` with `terms` operands, the densest chain per character.
std::string OnesChain(std::uint32_t terms) {
  std::string out = "=1";
  for (std::uint32_t i = 2; i <= terms; ++i) {
    out.append("+1");
  }
  return out;
}

// `=1+2*3+4*5+...`: an additive chain whose later operands are
// multiplicative chains.
std::string MixedChain(std::uint32_t products) {
  std::string out = "=1";
  for (std::uint32_t i = 0; i < products; ++i) {
    out.append("+").append(std::to_string(2 * i + 2)).append("*").append(std::to_string(2 * i + 3));
  }
  return out;
}

bool HasError(const std::vector<ParseError>& errors, ParseErrorCode code) {
  for (const ParseError& e : errors) {
    if (e.code == code) {
      return true;
    }
  }
  return false;
}

const AstNode* ParseClean(std::string_view src, Arena& arena) {
  Parser p(src, arena);
  const AstNode* root = p.parse();
  EXPECT_NE(root, nullptr);
  EXPECT_TRUE(p.errors().empty()) << "parse errors for a " << src.size() << "-character formula";
  return root;
}

TEST(OperatorChain, SamePowerOperatorsFormOneNode) {
  Arena arena;
  const AstNode* root = ParseClean("=1+2-3+4", arena);
  ASSERT_NE(root, nullptr);
  ASSERT_EQ(root->kind(), NodeKind::BinaryOp);
  ASSERT_EQ(root->as_binary_arity(), 4U);
  EXPECT_EQ(root->as_binary_op(0), BinOp::Add);
  EXPECT_EQ(root->as_binary_op(1), BinOp::Sub);
  EXPECT_EQ(root->as_binary_op(2), BinOp::Add);
  for (std::uint32_t i = 0; i < 4U; ++i) {
    EXPECT_EQ(root->as_binary_operand(i).kind(), NodeKind::Literal);
  }
  EXPECT_EQ(dump_sexpr(*root), "(binary + (num 1) (num 2) - (num 3) + (num 4))");
}

TEST(OperatorChain, TighterOperatorsNestAsOneOperand) {
  Arena arena;
  const AstNode* root = ParseClean("=1+2*3/4+5", arena);
  ASSERT_NE(root, nullptr);
  ASSERT_EQ(root->kind(), NodeKind::BinaryOp);
  ASSERT_EQ(root->as_binary_arity(), 3U);
  const AstNode& product = root->as_binary_operand(1);
  ASSERT_EQ(product.kind(), NodeKind::BinaryOp);
  EXPECT_EQ(product.as_binary_arity(), 3U);
  EXPECT_EQ(product.as_binary_op(0), BinOp::Mul);
  EXPECT_EQ(product.as_binary_op(1), BinOp::Div);
}

TEST(OperatorChain, LooserOperatorTakesTheChainAsItsLeadingOperand) {
  Arena arena;
  const AstNode* root = ParseClean("=1+2&3&4=5", arena);
  ASSERT_NE(root, nullptr);
  ASSERT_EQ(root->kind(), NodeKind::BinaryOp);
  ASSERT_EQ(root->as_binary_arity(), 2U);
  EXPECT_EQ(root->as_binary_op(0), BinOp::Eq);
  const AstNode& concat = root->as_binary_operand(0);
  ASSERT_EQ(concat.kind(), NodeKind::BinaryOp);
  EXPECT_EQ(concat.as_binary_arity(), 3U);
  EXPECT_EQ(concat.as_binary_operand(0).kind(), NodeKind::BinaryOp);
}

TEST(OperatorChain, ParenthesesStartANewNode) {
  Arena arena;
  const AstNode* root = ParseClean("=(1+2)+3-(4-5)", arena);
  ASSERT_NE(root, nullptr);
  ASSERT_EQ(root->kind(), NodeKind::BinaryOp);
  ASSERT_EQ(root->as_binary_arity(), 3U);
  EXPECT_EQ(root->as_binary_operand(0).paren_depth(), 1U);
  EXPECT_EQ(root->as_binary_operand(0).as_binary_arity(), 2U);
  EXPECT_EQ(root->as_binary_operand(2).paren_depth(), 1U);
  EXPECT_EQ(format_formula(*root), "(1+2)+3-(4-5)");
}

TEST(OperatorChain, FormatterKeepsAnUnparenthesisedRightOperandApart) {
  Arena arena;
  const AstNode* inner =
      make_binary_op(arena, BinOp::Sub, make_literal(arena, Value::number(2)), make_literal(arena, Value::number(3)));
  ASSERT_NE(inner, nullptr);
  const AstNode* operands[] = {make_literal(arena, Value::number(1)), inner};
  const BinOp ops[] = {BinOp::Sub};
  const AstNode* root = make_binary_chain(arena, operands, ops, 2U);
  ASSERT_NE(root, nullptr);
  EXPECT_EQ(format_formula(*root), "1-(2-3)");
}

TEST(OperatorChain, ComparisonAndPowerChainsAreLeftAssociative) {
  Arena arena;
  const AstNode* cmp = ParseClean("=1<2=TRUE<>FALSE", arena);
  ASSERT_NE(cmp, nullptr);
  ASSERT_EQ(cmp->kind(), NodeKind::BinaryOp);
  EXPECT_EQ(cmp->as_binary_arity(), 4U);
  const AstNode* pow = ParseClean("=2^3^2", arena);
  ASSERT_NE(pow, nullptr);
  ASSERT_EQ(pow->kind(), NodeKind::BinaryOp);
  EXPECT_EQ(pow->as_binary_arity(), 3U);
}

TEST(OperatorChain, LongChainsParseAndFormatBackToTheirText) {
  struct Case {
    const char* label;
    std::string text;
    std::uint32_t arity;
  };
  const Case cases[] = {
      {"150 refs", RefChain(150), 150U},         {"1000 refs", RefChain(1000), 1000U},
      {"2700 refs", ShortRefChain(2700), 2700U}, {"4095 ones", OnesChain(4095), 4095U},
      {"mixed", MixedChain(700), 701U},
  };
  for (const Case& c : cases) {
    SCOPED_TRACE(c.label);
    ASSERT_LT(c.text.size(), 8192U);
    Arena arena;
    const AstNode* root = ParseClean(c.text, arena);
    ASSERT_NE(root, nullptr);
    ASSERT_EQ(root->kind(), NodeKind::BinaryOp);
    EXPECT_EQ(root->as_binary_arity(), c.arity);
    EXPECT_TRUE(ast_depth_within_limit(*root, 3U));
    EXPECT_EQ("=" + format_formula(*root), c.text);
  }
}

TEST(OperatorChain, LongChainRoundTripsThroughR1C1) {
  Arena arena;
  const std::string text = ShortRefChain(2700);
  const AstNode* root = ParseClean(text, arena);
  ASSERT_NE(root, nullptr);
  const std::string r1c1 = format_formula_r1c1(*root, /*anchor_row=*/0U, /*anchor_col=*/1U);
  std::string expected = "RC[-1]";
  for (std::uint32_t i = 1; i < 2700U; ++i) {
    const std::uint32_t row = i % 9U;
    expected.append(row == 0U ? "+RC[-1]" : "+R[" + std::to_string(row) + "]C[-1]");
  }
  EXPECT_EQ(r1c1, expected);
}

TEST(OperatorChain, FillShiftsEveryOperandOfALongChain) {
  Arena arena;
  const std::string text = ShortRefChain(2700);
  const AstNode* root = ParseClean(text, arena);
  ASSERT_NE(root, nullptr);
  const AstNode* shifted = shift_relative_refs(*root, arena, /*row_delta=*/1, /*col_delta=*/1);
  ASSERT_NE(shifted, nullptr);
  ASSERT_EQ(shifted->kind(), NodeKind::BinaryOp);
  ASSERT_EQ(shifted->as_binary_arity(), 2700U);
  std::string expected = "B2";
  for (std::uint32_t i = 1; i < 2700U; ++i) {
    expected.append("+B").append(std::to_string(i % 9U + 2U));
  }
  EXPECT_EQ(format_formula(*shifted), expected);
}

TEST(OperatorChain, RowInsertMovesOnlyTheOperandsBelowIt) {
  Arena arena;
  const std::string text = RefChain(150);
  const AstNode* root = ParseClean(text, arena);
  ASSERT_NE(root, nullptr);
  // Two rows inserted above row 101 (0-based 100).
  const RowColShiftTransform transform("Sheet1", RowColAxis::kRow, RowColEdit::kInsert, /*index=*/100U, /*count=*/2U,
                                       /*local_means_target=*/true);
  const AstNode* shifted = shift_refs(*root, arena, transform);
  ASSERT_NE(shifted, nullptr);
  ASSERT_EQ(shifted->kind(), NodeKind::BinaryOp);
  EXPECT_EQ(format_formula(shifted->as_binary_operand(99)), "A100");
  EXPECT_EQ(format_formula(shifted->as_binary_operand(100)), "A103");
  EXPECT_EQ(format_formula(shifted->as_binary_operand(149)), "A152");
}

TEST(OperatorChain, ChainDoesNotCountTowardTheNestingCap) {
  Arena arena;
  ParserOptions opts;
  opts.max_parse_depth = 5;
  const std::string text = OnesChain(40);
  Parser parser(text, arena, opts);
  const AstNode* root = parser.parse();
  ASSERT_NE(root, nullptr);
  EXPECT_TRUE(parser.errors().empty());
  ASSERT_EQ(root->kind(), NodeKind::BinaryOp);
  EXPECT_EQ(root->as_binary_arity(), 40U);
}

TEST(OperatorChain, NestedParenthesesStillHitTheDefaultCap) {
  std::string text = "=";
  for (std::uint32_t i = 0; i < kMaxFormulaAstDepth; ++i) {
    text.append("(1+");
  }
  text.append("1");
  text.append(kMaxFormulaAstDepth, ')');
  Arena arena;
  Parser parser(text, arena);
  const AstNode* root = parser.parse();
  ASSERT_NE(root, nullptr);
  EXPECT_TRUE(HasError(parser.errors(), ParseErrorCode::NestedFormulaTooDeep));
}

TEST(OperatorChain, FormulaLengthLimitBoundsTheChain) {
  Arena arena;
  const std::string text = OnesChain(4200);
  Parser parser(text, arena);
  (void)parser.parse();
  EXPECT_TRUE(HasError(parser.errors(), ParseErrorCode::LexerExcessiveLength));
  EXPECT_FALSE(HasError(parser.errors(), ParseErrorCode::NestedFormulaTooDeep));
}

}  // namespace
}  // namespace parser
}  // namespace formulon
