//
// Tests for built-in name resolution and the qualified-built-in-call entry
// rule. Every formula below was entered into Mac Excel 365 (ja-JP); the
// rejected set is what Excel refused at entry, the accepted set is what it
// stored.

#include "eval/builtin_names.h"

#include <string>
#include <string_view>

#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/formula_prefix.h"
#include "parser/parser.h"
#include "utils/arena.h"

namespace formulon {
namespace eval {
namespace {

// Mirrors formula entry: storage prefixes are stripped before the parse.
const parser::AstNode* ParseEntry(std::string_view formula, Arena& arena, std::string& storage) {
  storage = parser::strip_storage_prefixes(formula);
  return parser::parse_strict(storage, arena);
}

TEST(BuiltinNames, ResolvesEverySource) {
  EXPECT_STREQ(resolve_builtin_function_name("sum"), "SUM");
  EXPECT_STREQ(resolve_builtin_function_name("XLOOKUP"), "XLOOKUP");
  EXPECT_STREQ(resolve_builtin_function_name("Let"), "LET");
  EXPECT_STREQ(resolve_builtin_function_name("LAMBDA"), "LAMBDA");
  EXPECT_EQ(resolve_builtin_function_name("NOSUCH"), nullptr);
  EXPECT_EQ(resolve_builtin_function_name("EUROCONVERT"), nullptr);
}

TEST(BuiltinNames, QualifiedBuiltinCallIsFound) {
  for (const char* src :
       {"Sheet1!SUM(1)", "'My Sheet'!SUM(1)", "[0]!SUM(1)", "Sheet1!sum(1)", "Sheet1!XLOOKUP(1,A1:A5,A1:A5)",
        "Sheet1!_xlfn.XLOOKUP(1,A1:A5,A1:A5)", "Sheet1!LET(a,1,a)", "Sheet1!LAMBDA(x,x)(1)", "Sheet1!IF(1,2)",
        "Sheet1!NOW()", "Sheet1!ISOMITTED(1)", "SUM(Sheet1!SUM(1))", "Sheet1!SUM(1)+1"}) {
    Arena a;
    std::string storage;
    const parser::AstNode* root = ParseEntry(src, a, storage);
    ASSERT_NE(root, nullptr) << src;
    EXPECT_NE(find_qualified_builtin_call(*root), nullptr) << src;
  }
}

TEST(BuiltinNames, RejectedSpellingsTheParserAlreadyRefuses) {
  for (const char* src : {"Sheet1!TRUE(1)", "Sheet1!SUM (1)", "Sheet1:Sheet2!SUM(1)"}) {
    Arena a;
    std::string storage;
    EXPECT_EQ(ParseEntry(src, a, storage), nullptr) << src;
  }
}

TEST(BuiltinNames, AcceptedQualifiedFormsAreNotFlagged) {
  for (const char* src :
       {"Sheet1!MyFn(1)", "MyFn(1)", "[0]!WFn(1)", "WFn(1)", "'My Sheet'!QFn(1)", "Sheet1!NOSUCH(1)", "Sheet1!SUMX(1)",
        "Sheet1!EUROCONVERT(1,\"FRF\",\"EUR\")", "Sheet1!SUM", "Sheet1!SVal", "SUM(1)", "LAMBDA(x,x)(1)"}) {
    Arena a;
    std::string storage;
    const parser::AstNode* root = ParseEntry(src, a, storage);
    ASSERT_NE(root, nullptr) << src;
    EXPECT_EQ(find_qualified_builtin_call(*root), nullptr) << src;
  }
}

}  // namespace
}  // namespace eval
}  // namespace formulon
