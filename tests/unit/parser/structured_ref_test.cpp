//
// Parser-level tests for structured (table) references. The parser captures
// the bracket payload verbatim into the `column` slot of the AST node; the
// resolver in `src/eval/structured_ref.cpp` re-parses the payload at
// evaluation time. These tests cover the eight shapes Bundle 4.4 ships:
//
//   1. `Table[Col]`              — single-column data range
//   2. `Table[@Col]`              — row-implicit single cell
//   3. `Table[#All]`              — whole table
//   4. `Table[#Headers]`          — header row
//   5. `Table[#Totals]`           — totals row
//   6. `Table[#Data]`             — data rows (default)
//   7. `Table[ColA]:Table[ColB]`  — cross-column rectangle
//   8. `Table[[#Headers],[Col]]`  — multi-specifier form

#include <string>
#include <string_view>

#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "utils/arena.h"

namespace formulon {
namespace parser {
namespace {

TEST(ParserStructuredRef, SingleColumn) {
  Arena a;
  Parser p("=Table1[Region]", a);
  AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  EXPECT_TRUE(p.errors().empty());
  ASSERT_EQ(root->kind(), NodeKind::StructuredRef);
  EXPECT_EQ(root->as_structured_ref_table(), "Table1");
  EXPECT_EQ(root->as_structured_ref_column(), "Region");
}

TEST(ParserStructuredRef, RowImplicit) {
  Arena a;
  Parser p("=Sales[@Amount]", a);
  AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  EXPECT_TRUE(p.errors().empty());
  ASSERT_EQ(root->kind(), NodeKind::StructuredRef);
  EXPECT_EQ(root->as_structured_ref_table(), "Sales");
  // Parser preserves the bracket payload verbatim; the resolver splits the
  // leading `@` into the `kThisRow` selector at evaluation time.
  EXPECT_EQ(root->as_structured_ref_column(), "@Amount");
}

TEST(ParserStructuredRef, AllSpecifier) {
  Arena a;
  Parser p("=Sales[#All]", a);
  AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  EXPECT_TRUE(p.errors().empty());
  ASSERT_EQ(root->kind(), NodeKind::StructuredRef);
  EXPECT_EQ(root->as_structured_ref_table(), "Sales");
  EXPECT_EQ(root->as_structured_ref_column(), "#All");
}

TEST(ParserStructuredRef, HeadersSpecifier) {
  Arena a;
  Parser p("=Sales[#Headers]", a);
  AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  EXPECT_TRUE(p.errors().empty());
  ASSERT_EQ(root->kind(), NodeKind::StructuredRef);
  EXPECT_EQ(root->as_structured_ref_column(), "#Headers");
}

TEST(ParserStructuredRef, TotalsSpecifier) {
  Arena a;
  Parser p("=Sales[#Totals]", a);
  AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  EXPECT_TRUE(p.errors().empty());
  ASSERT_EQ(root->kind(), NodeKind::StructuredRef);
  EXPECT_EQ(root->as_structured_ref_column(), "#Totals");
}

TEST(ParserStructuredRef, DataSpecifier) {
  Arena a;
  Parser p("=Sales[#Data]", a);
  AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  EXPECT_TRUE(p.errors().empty());
  ASSERT_EQ(root->kind(), NodeKind::StructuredRef);
  EXPECT_EQ(root->as_structured_ref_column(), "#Data");
}

TEST(ParserStructuredRef, ColumnRangeAcrossTable) {
  Arena a;
  Parser p("=Sales[Region]:Sales[Amount]", a);
  AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  EXPECT_TRUE(p.errors().empty());
  // The `:` operator between two structured refs builds a RangeOp whose
  // endpoints are the two `StructuredRef` nodes. Sheet-qualified semantics
  // are not added by the parser; the evaluator decides how to combine the
  // two rectangles when this shape lands.
  ASSERT_EQ(root->kind(), NodeKind::RangeOp);
  EXPECT_EQ(root->as_range_lhs().kind(), NodeKind::StructuredRef);
  EXPECT_EQ(root->as_range_rhs().kind(), NodeKind::StructuredRef);
  EXPECT_EQ(root->as_range_lhs().as_structured_ref_column(), "Region");
  EXPECT_EQ(root->as_range_rhs().as_structured_ref_column(), "Amount");
}

TEST(ParserStructuredRef, MultiSpecifierForm) {
  // The full multi-specifier shape: `[#Headers]` and `[Region]` separated
  // by a top-level comma. The parser keeps the entire payload verbatim;
  // the resolver splits it.
  Arena a;
  Parser p("=Sales[[#Headers],[Region]]", a);
  AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  EXPECT_TRUE(p.errors().empty());
  ASSERT_EQ(root->kind(), NodeKind::StructuredRef);
  EXPECT_EQ(root->as_structured_ref_table(), "Sales");
  EXPECT_EQ(root->as_structured_ref_column(), "[#Headers],[Region]");
}

TEST(ParserStructuredRef, BracketedColumnRange) {
  // `Table[[ColA]:[ColB]]` — column slice, no specifier. Excel emits this
  // when both column names contain non-identifier characters.
  Arena a;
  Parser p("=Sales[[Region]:[Amount]]", a);
  AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  EXPECT_TRUE(p.errors().empty());
  ASSERT_EQ(root->kind(), NodeKind::StructuredRef);
  EXPECT_EQ(root->as_structured_ref_column(), "[Region]:[Amount]");
}

TEST(ParserStructuredRef, EmptyBracketsAcceptedAsWholeTableData) {
  // `Table[]` is rare in practice but Excel accepts it as `Table[#Data]`.
  Arena a;
  Parser p("=Sales[]", a);
  AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  EXPECT_TRUE(p.errors().empty());
  ASSERT_EQ(root->kind(), NodeKind::StructuredRef);
  EXPECT_EQ(root->as_structured_ref_column(), "");
}

}  // namespace
}  // namespace parser
}  // namespace formulon
