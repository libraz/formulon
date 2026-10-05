#pragma once

#include <cstdint>
#include <string_view>

#include "eval/array_alloc.h"
#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon::eval::lookup_test_helpers {

inline thread_local std::uint32_t* g_tick_array_calls = nullptr;

inline Value TickArrayImpl(const Value* /*args*/, std::uint32_t /*arity*/, Arena& arena) {
  if (g_tick_array_calls != nullptr) {
    ++*g_tick_array_calls;
  }
  Value* cells = nullptr;
  ArrayValue* array = allocate_array_value(1U, 1U, arena, cells);
  if (array == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  cells[0] = Value::number(1.0);
  return Value::array(array);
}

inline Value EvalSourceInWithRegistry(std::string_view formula, const Workbook& wb, const Sheet& current,
                                      const FunctionRegistry& registry) {
  Arena parse_arena;
  Arena eval_arena;
  parser::Parser p(formula, parse_arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << formula;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  EvalState state;
  const EvalContext ctx(wb, current, state);
  return evaluate(*root, eval_arena, registry, ctx);
}

inline void ExpectArrayShape(const Value& value, std::uint32_t rows, std::uint32_t cols) {
  ASSERT_TRUE(value.is_array()) << value.debug_to_string();
  EXPECT_EQ(value.as_array_rows(), rows);
  EXPECT_EQ(value.as_array_cols(), cols);
}

}  // namespace formulon::eval::lookup_test_helpers
