//
// libFuzzer harness for the Formulon tree-walk evaluator.
//
// Fuzz goal: parse arbitrary bytes as a formula, then evaluate the AST
// (when parsing succeeds). Detects crashes in either layer and any
// undefined behaviour in the operator/dispatch path.

#include <cstddef>
#include <cstdint>
#include <string_view>

#include "eval/tree_walker.h"
#include "parser/parser.h"
#include "utils/arena.h"

extern "C" int LLVMFuzzerTestOneInput(const uint8_t* data, size_t size) {
  if (size > 65536) {
    return 0;
  }
  formulon::Arena arena;
  std::string_view src(reinterpret_cast<const char*>(data), size);
  formulon::parser::Parser parser(src, arena);
  formulon::parser::AstNode* root = parser.parse();
  if (!root || !parser.errors().empty()) {
    return 0;
  }
  (void)formulon::eval::evaluate(*root, arena);
  return 0;
}
