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

TEST(BuiltinNames, FormulaEntryGateRejectsQualifiedBuiltinsOnly) {
  for (const char* src : {"Sheet1!SUM(1)", "[0]!LET(a,1,a)", "SUM(Sheet1!SUM(1))"}) {
    Arena arena;
    std::string storage = parser::strip_storage_prefixes(src);
    EXPECT_EQ(parse_formula_entry(storage, arena), nullptr) << src;
  }

  for (const char* src : {"Sheet1!MyFn(1)", "Sheet1!NOSUCH(1)", "SUM(1)", "A1+1"}) {
    Arena arena;
    std::string storage = parser::strip_storage_prefixes(src);
    EXPECT_NE(parse_formula_entry(storage, arena), nullptr) << src;
  }

  Arena malformed_arena;
  EXPECT_EQ(parse_formula_entry("SUM(", malformed_arena), nullptr);
}

// Mac Excel 365 at entry: what a reference-only parameter (SUMIF's range)
// accepts and refuses.
TEST(BuiltinNames, ReferenceOnlyParameterTakesReferenceExpressionsOnly) {
  for (const char* arg : {"A1",
                          "(A1)",
                          "(A1,A2)",
                          "A1:A2 A1",
                          "A:A",
                          "1:1",
                          "Sheet2!A1",
                          "Sheet2:Sheet3!A1",
                          "[Book9.xlsx]Sheet1!A1",
                          "A1#",
                          "@A1",
                          "#REF!",
                          "Nm",
                          "t[c]",
                          "INDEX(A1:A2,1):A2",
                          "A1:INDEX(A1:A2,2)",
                          "LAMBDA(x,x)(A1)",
                          "FOO(A1)",
                          "INDEX({1,2},1)",
                          "INDIRECT(1)",
                          "OFFSET(A1,0,0)",
                          "IF(1,1,1)",
                          "IF(1,A1)",
                          "CHOOSE(1,1)",
                          "XLOOKUP(1,{1},{1})",
                          "IFS(1,1)",
                          "SWITCH(1,1,1)",
                          "LET(x,1,x)",
                          "TRIMRANGE(1)",
                          "SINGLE(1)",
                          "TAKE(A1:A2,1)",
                          "DROP(A1:A2,1)"}) {
    Arena arena;
    std::string storage;
    const parser::AstNode* root = ParseEntry(std::string("SUMIF(") + arg + ",1)", arena, storage);
    ASSERT_NE(root, nullptr) << arg;
    EXPECT_EQ(find_non_reference_argument(*root), nullptr) << arg;
  }
  for (const char* arg : {"",
                          "1",
                          "\"A1\"",
                          "TRUE",
                          "{1}",
                          "#N/A",
                          "1%",
                          "(1)",
                          "(A1,1)",
                          "A1+0",
                          "-A1",
                          "+A1",
                          "A1&\"\"",
                          "SUM(A1)",
                          "IFERROR(A1,A1)",
                          "IFNA(A1,A1)",
                          "HYPERLINK(A1)",
                          "HSTACK(A1)",
                          "VSTACK(A1)",
                          "EXPAND(A1,1)",
                          "CHOOSEROWS(A1:A2,1)",
                          "FILTER(A1:A2,1)",
                          "SORT(A1:A2)",
                          "UNIQUE(A1:A2)",
                          "XMATCH(1,A1)",
                          "ROWS(A1)",
                          "CELL(\"row\",A1)"}) {
    Arena arena;
    std::string storage;
    const parser::AstNode* root = ParseEntry(std::string("SUMIF(") + arg + ",1)", arena, storage);
    ASSERT_NE(root, nullptr) << arg;
    EXPECT_NE(find_non_reference_argument(*root), nullptr) << arg;
  }
}

// Each reference-only parameter Excel has, given an empty argument, and the
// parameters beside it, which take any value.
TEST(BuiltinNames, ReferenceOnlyParametersPerFunction) {
  for (const char* src : {"AREAS(1)",           "COUNTBLANK(1)",
                          "FORMULATEXT(1)",     "ISFORMULA(1)",
                          "PHONETIC(1)",        "ROW(1)",
                          "COLUMN(1)",          "SUMIF(,1)",
                          "SUMIF(A1,1,)",       "AVERAGEIF(,1)",
                          "AVERAGEIF(A1,1,1)",  "COUNTIF(,1)",
                          "COUNTIFS(A1,1,,1)",  "SUMIFS(,A1,1)",
                          "SUMIFS(A1,,1)",      "SUMIFS(A1,A1,1,1,1)",
                          "AVERAGEIFS(A1,1,1)", "MAXIFS(,A1,1)",
                          "MINIFS(A1,A1,1,,1)", "DSUM(,1,1)",
                          "DGET(1,1,1)",        "DCOUNTA(,1,1)",
                          "CELL(\"row\",)",     "OFFSET(,1,1)",
                          "RANK(1,)",           "RANK.EQ(1,1)",
                          "RANK.AVG(1,,1)",     "SUBTOTAL(9,)",
                          "SUBTOTAL(9,A1,)",    "SUBTOTAL(9,A1,A1,1)"}) {
    Arena arena;
    std::string storage;
    const parser::AstNode* root = ParseEntry(src, arena, storage);
    ASSERT_NE(root, nullptr) << src;
    EXPECT_NE(find_non_reference_argument(*root), nullptr) << src;
  }
  for (const char* src : {"SUM(,)",
                          "IF(A1,,1)",
                          "COUNT(1,,1)",
                          "INDEX(A1,,)",
                          "SUMIF(A1,)",
                          "SUMIF(A1,1)",
                          "COUNTIFS(A1,,A1,)",
                          "SUMIFS(A1,A1,)",
                          "DSUM(A1,,)",
                          "CELL(,A1)",
                          "OFFSET(A1,,)",
                          "RANK(,A1,)",
                          "SUBTOTAL(,A1)",
                          "SUBTOTAL(1,A1)",
                          "AGGREGATE(9,4,)",
                          "AGGREGATE(9,4,1)",
                          "AGGREGATE(14,4,,1)",
                          "ROW()",
                          "COLUMN()",
                          "SUM(SUMIF(A1,1),1)",
                          "ISREF(1)",
                          "ROWS(1)"}) {
    Arena arena;
    std::string storage;
    const parser::AstNode* root = ParseEntry(src, arena, storage);
    ASSERT_NE(root, nullptr) << src;
    EXPECT_EQ(find_non_reference_argument(*root), nullptr) << src;
  }
  // Found at any depth, including inside LET and LAMBDA bodies.
  for (const char* src : {"1+SUBTOTAL(9,)", "LET(x,1,SUMIF(1,x))", "IF(A1,COUNTIF(1,1),0)"}) {
    Arena arena;
    std::string storage;
    const parser::AstNode* root = ParseEntry(src, arena, storage);
    ASSERT_NE(root, nullptr) << src;
    EXPECT_NE(find_non_reference_argument(*root), nullptr) << src;
  }
}

}  // namespace
}  // namespace eval
}  // namespace formulon
