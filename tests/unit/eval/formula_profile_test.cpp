// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Formula-profile routing tests. These cases intentionally assert different
// outcomes for Mac Excel 365 ja-JP and Windows Excel 365 ja-JP so the runtime
// default and explicit profile switches stay visible outside the oracle suite.

#include <string_view>

#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "test_eval_helpers.h"
#include "gtest/gtest.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

Value EvalWithProfile(std::string_view src, ExcelProfile profile) {
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
  return evaluate(*root, eval_arena, default_registry(), test::context_with_profile(profile));
}

TEST(FormulaProfile, NewWorkbookDefaultsToWinJa) {
  const Workbook wb = Workbook::create();
  EXPECT_STREQ(excel_profile_id(wb.excel_profile()), "win-365-ja_JP");
}

TEST(FormulaProfile, DynamicArrayFunctionSwitchesByProfile) {
  Value v = EvalWithProfile("=SEQUENCE(2)", test::mac_profile());
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(v.as_array_cols(), 1U);

  v = EvalWithProfile("=SEQUENCE(2)", test::win_profile());
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(FormulaProfile, TextFunctionSwitchesByProfile) {
  Value v = EvalWithProfile("=TEXTBEFORE(\"a-b\", \"-\")", test::mac_profile());
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "a");

  v = EvalWithProfile("=TEXTBEFORE(\"a-b\", \"-\")", test::win_profile());
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
