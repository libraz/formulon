#pragma once

#include <algorithm>
#include <string_view>
#include <vector>

#include "eval/dep_graph.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "utils/arena.h"

namespace formulon::eval::dep_extractor_test_helpers {

inline const parser::AstNode* ParseFormula(std::string_view source, Arena& arena) {
  parser::Parser parser(source, arena);
  parser::AstNode* root = parser.parse();
  EXPECT_NE(root, nullptr);
  return root;
}

inline std::vector<CellNodeId> Sorted(std::vector<CellNodeId> v) {
  std::sort(v.begin(), v.end(), [](CellNodeId a, CellNodeId b) {
    if (a.sheet_id != b.sheet_id)
      return a.sheet_id < b.sheet_id;
    if (a.row != b.row)
      return a.row < b.row;
    return a.col < b.col;
  });
  return v;
}

}  // namespace formulon::eval::dep_extractor_test_helpers
