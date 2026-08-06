//
// Implementation of the `EvalSource*` helpers declared in
// `test_eval_helpers.h`. Each helper resets the relevant test arenas,
// drives the parser, and forwards into the production tree-walk
// evaluator.

#include "util/test_eval_helpers.h"

#include <gtest/gtest.h>

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "util/test_arena.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace test {

namespace {

// Parses `formula` into a fresh AST owned by the test parse arena.
// Returns nullptr on parse failure and emits a gtest non-fatal failure
// so the diagnostic surfaces in the offending test case rather than
// being silently swallowed.
parser::AstNode* ParseOrFail(std::string_view formula, Arena& parse_arena) {
  parser::Parser p(formula, parse_arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << formula;
  return root;
}

}  // namespace

Value EvalSource(std::string_view formula) {
  Arena& parse_arena = test_parse_arena();
  Arena& eval_arena = test_eval_arena();
  parser::AstNode* root = ParseOrFail(formula, parse_arena);
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return eval::evaluate(*root, eval_arena);
}

Value EvalSourceIn(std::string_view formula, const Workbook& wb, const Sheet& current) {
  Arena& parse_arena = test_parse_arena();
  Arena& eval_arena = test_eval_arena();
  parser::AstNode* root = ParseOrFail(formula, parse_arena);
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  eval::EvalState state;
  const eval::EvalContext ctx(wb, current, state);
  return eval::evaluate(*root, eval_arena, eval::default_registry(), ctx);
}

Value EvalSourceAt(std::string_view formula, const Workbook& wb, const Sheet& current, std::uint32_t row,
                   std::uint32_t col) {
  Arena& parse_arena = test_parse_arena();
  Arena& eval_arena = test_eval_arena();
  parser::AstNode* root = ParseOrFail(formula, parse_arena);
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  eval::EvalState state;
  const eval::EvalContext ctx = eval::EvalContext(wb, current, state).with_formula_cell(row, col);
  return eval::evaluate(*root, eval_arena, eval::default_registry(), ctx);
}

Value EvalSourceAt(std::string_view formula, const Workbook& wb, std::string_view sheet, std::uint32_t row,
                   std::uint32_t col) {
  const Sheet* current = wb.sheet_by_name(sheet);
  if (current == nullptr) {
    return Value::error(ErrorCode::Ref);
  }
  return EvalSourceAt(formula, wb, *current, row, col);
}

}  // namespace test
}  // namespace formulon
