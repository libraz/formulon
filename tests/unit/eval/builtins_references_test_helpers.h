#pragma once

#include <cstdint>
#include <string>
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

namespace formulon::eval::references_test_helpers {

inline Workbook ColumnOfTen() {
  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 10; ++r) {
    wb.sheet(0).set_cell_value(r, 0, Value::number(static_cast<double>(r + 1)));
  }
  return wb;
}

inline void ExpectNumber(const Value& v, double expected, std::string_view formula) {
  ASSERT_TRUE(v.is_number()) << formula;
  EXPECT_DOUBLE_EQ(v.as_number(), expected) << formula;
}

inline void ExpectText(const Value& v, std::string_view expected, std::string_view formula) {
  ASSERT_TRUE(v.is_text()) << formula;
  EXPECT_EQ(std::string(v.as_text()), std::string(expected)) << formula;
}

inline Value EvalSource(std::string_view src) {
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
  return evaluate(*root, eval_arena);
}

inline Value EvalSourceIn(std::string_view src, const Workbook& wb, const Sheet& current) {
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
  const EvalContext ctx = test::workbook_context(wb, current, state);
  return evaluate(*root, eval_arena, default_registry(), ctx);
}

}  // namespace formulon::eval::references_test_helpers
