// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Tests for the lazy `TEXTSPLIT(text, col_delim, [row_delim],
// [ignore_empty], [match_mode], [pad_with])` builtin.

#include <cstdint>
#include <string_view>

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "test_eval_helpers.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

Value EvalSrc(std::string_view src) {
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
  return evaluate(*root, eval_arena, default_registry(), test::mac_context());
}

Value EvalIn(std::string_view src, const Workbook& wb, const Sheet& current) {
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
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, current, state);
  return evaluate(*root, eval_arena, default_registry(), ctx);
}

// ---------------------------------------------------------------------------
// Column-only split
// ---------------------------------------------------------------------------

TEST(BuiltinsTextSplit, ScalarColDelimSplitsIntoRow) {
  const Value v = EvalSrc("=TEXTSPLIT(\"a,b,c\", \",\")");
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(v.as_array_cols(), 3U);
  EXPECT_EQ(v.as_array_cells()[0].as_text(), "a");
  EXPECT_EQ(v.as_array_cells()[1].as_text(), "b");
  EXPECT_EQ(v.as_array_cells()[2].as_text(), "c");
}

TEST(BuiltinsTextSplit, NoMatchYieldsSingleCellArray) {
  const Value v = EvalSrc("=TEXTSPLIT(\"abc\", \",\")");
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  EXPECT_EQ(v.as_array_cells()[0].as_text(), "abc");
}

TEST(BuiltinsTextSplit, EmptyTextYieldsSingleEmptyCell) {
  const Value v = EvalSrc("=TEXTSPLIT(\"\", \",\")");
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  EXPECT_EQ(v.as_array_cells()[0].as_text(), "");
}

TEST(BuiltinsTextSplit, EmptyDelimiterIsNoOp) {
  // An empty col_delimiter at the only level keeps the whole text as a
  // single token (Mac Excel: TEXTSPLIT("abc", "") -> {"abc"}).
  const Value v = EvalSrc("=TEXTSPLIT(\"abc\", \"\")");
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  EXPECT_EQ(v.as_array_cells()[0].as_text(), "abc");
}

TEST(BuiltinsTextSplit, AdjacentDelimitersProduceEmptyTokens) {
  // Without ignore_empty the empty middle token is preserved.
  const Value v = EvalSrc("=TEXTSPLIT(\"a,,b\", \",\")");
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(v.as_array_cols(), 3U);
  EXPECT_EQ(v.as_array_cells()[0].as_text(), "a");
  EXPECT_EQ(v.as_array_cells()[1].as_text(), "");
  EXPECT_EQ(v.as_array_cells()[2].as_text(), "b");
}

TEST(BuiltinsTextSplit, IgnoreEmptyDropsEmptyTokens) {
  const Value v = EvalSrc("=TEXTSPLIT(\"a,,b,\", \",\", , TRUE)");
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(v.as_array_cols(), 2U);
  EXPECT_EQ(v.as_array_cells()[0].as_text(), "a");
  EXPECT_EQ(v.as_array_cells()[1].as_text(), "b");
}

// ---------------------------------------------------------------------------
// Delimiter-as-array (any-of split)
// ---------------------------------------------------------------------------

TEST(BuiltinsTextSplit, ArrayOfColDelimitersSplitsOnAny) {
  // Split on either ',' or ';'.
  const Value v = EvalSrc("=TEXTSPLIT(\"a,b;c,d\", {\",\",\";\"})");
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(v.as_array_cols(), 4U);
  EXPECT_EQ(v.as_array_cells()[0].as_text(), "a");
  EXPECT_EQ(v.as_array_cells()[1].as_text(), "b");
  EXPECT_EQ(v.as_array_cells()[2].as_text(), "c");
  EXPECT_EQ(v.as_array_cells()[3].as_text(), "d");
}

TEST(BuiltinsTextSplit, EmptyEntriesInDelimArrayDropped) {
  // {",",""} is equivalent to {","} — empty entries are filtered.
  const Value v = EvalSrc("=TEXTSPLIT(\"a,b\", {\",\",\"\"})");
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(v.as_array_cols(), 2U);
}

TEST(BuiltinsTextSplit, LongerDelimiterTakesPrecedence) {
  // "ab" should win over "a" when both could match at i=0 of "abxc".
  const Value v = EvalSrc("=TEXTSPLIT(\"abxc\", {\"a\",\"ab\"})");
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(v.as_array_cols(), 2U);
  EXPECT_EQ(v.as_array_cells()[0].as_text(), "");
  EXPECT_EQ(v.as_array_cells()[1].as_text(), "xc");
}

// ---------------------------------------------------------------------------
// Row-only and 2D split
// ---------------------------------------------------------------------------

TEST(BuiltinsTextSplit, RowDelimiterOnlyYieldsColumn) {
  // Empty col_delimiter (no-op) + row split into 3 -> 3x1.
  const Value v = EvalSrc("=TEXTSPLIT(\"a;b;c\", \"\", \";\")");
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 3U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  EXPECT_EQ(v.as_array_cells()[0].as_text(), "a");
  EXPECT_EQ(v.as_array_cells()[1].as_text(), "b");
  EXPECT_EQ(v.as_array_cells()[2].as_text(), "c");
}

TEST(BuiltinsTextSplit, BothDelimitersYield2DGrid) {
  // 2x2 grid.
  const Value v = EvalSrc("=TEXTSPLIT(\"a,b;c,d\", \",\", \";\")");
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(v.as_array_cols(), 2U);
  EXPECT_EQ(v.as_array_cells()[0].as_text(), "a");
  EXPECT_EQ(v.as_array_cells()[1].as_text(), "b");
  EXPECT_EQ(v.as_array_cells()[2].as_text(), "c");
  EXPECT_EQ(v.as_array_cells()[3].as_text(), "d");
}

TEST(BuiltinsTextSplit, RaggedRowsPaddedWithNaByDefault) {
  // Row 1 has 2 cols, row 2 has 1 col, row 3 has 3 cols -> 3x3 with #N/A
  // padding the missing slots.
  const Value v = EvalSrc("=TEXTSPLIT(\"a,b;c;d,e,f\", \",\", \";\")");
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 3U);
  EXPECT_EQ(v.as_array_cols(), 3U);
  // Row 0: "a" "b" #N/A
  EXPECT_EQ(v.as_array_cells()[0].as_text(), "a");
  EXPECT_EQ(v.as_array_cells()[1].as_text(), "b");
  ASSERT_TRUE(v.as_array_cells()[2].is_error());
  EXPECT_EQ(v.as_array_cells()[2].as_error(), ErrorCode::NA);
  // Row 1: "c" #N/A #N/A
  EXPECT_EQ(v.as_array_cells()[3].as_text(), "c");
  ASSERT_TRUE(v.as_array_cells()[4].is_error());
  ASSERT_TRUE(v.as_array_cells()[5].is_error());
  // Row 2: "d" "e" "f"
  EXPECT_EQ(v.as_array_cells()[6].as_text(), "d");
  EXPECT_EQ(v.as_array_cells()[7].as_text(), "e");
  EXPECT_EQ(v.as_array_cells()[8].as_text(), "f");
}

TEST(BuiltinsTextSplit, CustomPadValueOverridesNa) {
  const Value v = EvalSrc("=TEXTSPLIT(\"a,b;c\", \",\", \";\", FALSE, 0, \"X\")");
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(v.as_array_cols(), 2U);
  EXPECT_EQ(v.as_array_cells()[0].as_text(), "a");
  EXPECT_EQ(v.as_array_cells()[1].as_text(), "b");
  EXPECT_EQ(v.as_array_cells()[2].as_text(), "c");
  EXPECT_EQ(v.as_array_cells()[3].as_text(), "X");
}

// ---------------------------------------------------------------------------
// Match mode (case folding)
// ---------------------------------------------------------------------------

TEST(BuiltinsTextSplit, MatchModeZeroIsCaseSensitive) {
  // "X" only matches uppercase.
  const Value v = EvalSrc("=TEXTSPLIT(\"aXbxc\", \"X\")");
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_cols(), 2U);
  EXPECT_EQ(v.as_array_cells()[0].as_text(), "a");
  EXPECT_EQ(v.as_array_cells()[1].as_text(), "bxc");
}

TEST(BuiltinsTextSplit, MatchModeOneIsCaseInsensitiveAscii) {
  const Value v = EvalSrc("=TEXTSPLIT(\"aXbxc\", \"X\", , , 1)");
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_cols(), 3U);
  EXPECT_EQ(v.as_array_cells()[0].as_text(), "a");
  EXPECT_EQ(v.as_array_cells()[1].as_text(), "b");
  EXPECT_EQ(v.as_array_cells()[2].as_text(), "c");
}

TEST(BuiltinsTextSplit, InvalidMatchModeRejected) {
  const Value v = EvalSrc("=TEXTSPLIT(\"a,b\", \",\", , , 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// ignore_empty in 2D grids
// ---------------------------------------------------------------------------

TEST(BuiltinsTextSplit, IgnoreEmptyOnEmptyTextYieldsCalc) {
  // Empty text + ignore_empty drops the only (empty) token, leaving a
  // grid with no cells -> #CALC! (Mac Excel surface).
  const Value v = EvalSrc("=TEXTSPLIT(\"\", \",\", , TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Calc);
}

TEST(BuiltinsTextSplit, IgnoreEmptyDropsAllEmptyRowsIn2D) {
  // ";a,b;" — three rows, first and last are empty. ignore_empty drops
  // them, leaving a 1x2 grid.
  const Value v = EvalSrc("=TEXTSPLIT(\";a,b;\", \",\", \";\", TRUE)");
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(v.as_array_cols(), 2U);
  EXPECT_EQ(v.as_array_cells()[0].as_text(), "a");
  EXPECT_EQ(v.as_array_cells()[1].as_text(), "b");
}

// ---------------------------------------------------------------------------
// Coercion of non-text args
// ---------------------------------------------------------------------------

TEST(BuiltinsTextSplit, NumberTextCoerces) {
  // 123 -> "123", split on "2" yields {"1","3"}.
  const Value v = EvalSrc("=TEXTSPLIT(123, \"2\")");
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_cols(), 2U);
  EXPECT_EQ(v.as_array_cells()[0].as_text(), "1");
  EXPECT_EQ(v.as_array_cells()[1].as_text(), "3");
}

// ---------------------------------------------------------------------------
// Range-sourced delimiters
// ---------------------------------------------------------------------------

TEST(BuiltinsTextSplit, ColDelimFromRange) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text(","));
  wb.sheet(0).set_cell_value(1, 0, Value::text(";"));
  const Value v = EvalIn("=TEXTSPLIT(\"a,b;c\", A1:A2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_cols(), 3U);
  EXPECT_EQ(v.as_array_cells()[0].as_text(), "a");
  EXPECT_EQ(v.as_array_cells()[1].as_text(), "b");
  EXPECT_EQ(v.as_array_cells()[2].as_text(), "c");
}

// ---------------------------------------------------------------------------
// Error propagation and arity
// ---------------------------------------------------------------------------

TEST(BuiltinsTextSplit, ErrorInTextPropagates) {
  const Value v = EvalSrc("=TEXTSPLIT(1/0, \",\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsTextSplit, ErrorInColDelimPropagates) {
  const Value v = EvalSrc("=TEXTSPLIT(\"a,b\", 1/0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsTextSplit, ZeroArgsRejected) {
  const Value v = EvalSrc("=TEXTSPLIT()");
  ASSERT_TRUE(v.is_error());
}

TEST(BuiltinsTextSplit, OneArgRejected) {
  const Value v = EvalSrc("=TEXTSPLIT(\"a,b\")");
  ASSERT_TRUE(v.is_error());
}

TEST(BuiltinsTextSplit, SevenArgsRejected) {
  const Value v = EvalSrc("=TEXTSPLIT(\"a\", \",\", \";\", FALSE, 0, \"X\", 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
