// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for `AstNode`: factory invocation, accessor round-trips, and
// the storage policy that strings and child arrays are arena-owned.

#include "parser/ast.h"

#include <cstring>
#include <string>
#include <string_view>
#include <type_traits>
#include <vector>

#include "gtest/gtest.h"
#include "parser/reference.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace parser {
namespace {

TEST(AstNodeLayout, TriviallyDestructible) {
  static_assert(std::is_trivially_destructible_v<AstNode>);
  // Runtime echo so the size shows up in the test log for visibility.
  EXPECT_LE(sizeof(AstNode), 64u) << "actual sizeof(AstNode) = " << sizeof(AstNode);
}

TEST(AstNodeLiteral, NumberRoundtrips) {
  Arena a;
  AstNode* n = make_literal(a, Value::number(42.0));
  ASSERT_NE(n, nullptr);
  EXPECT_EQ(n->kind(), NodeKind::Literal);
  EXPECT_TRUE(n->as_literal().is_number());
  EXPECT_DOUBLE_EQ(n->as_literal().as_number(), 42.0);
}

TEST(AstNodeLiteral, BlankRoundtrips) {
  Arena a;
  AstNode* n = make_literal(a, Value::blank());
  ASSERT_NE(n, nullptr);
  EXPECT_EQ(n->kind(), NodeKind::Literal);
  EXPECT_TRUE(n->as_literal().is_blank());
}

TEST(AstNodeLiteral, BoolRoundtrips) {
  Arena a;
  AstNode* n = make_literal(a, Value::boolean(true));
  ASSERT_NE(n, nullptr);
  EXPECT_TRUE(n->as_literal().as_boolean());
}

TEST(AstNodeRef, AccessorRoundtripsAndFormatsA1) {
  Arena a;
  Reference r;
  r.col = 0;
  r.row = 0;
  r.col_abs = true;
  r.row_abs = true;
  AstNode* n = make_ref(a, r);
  ASSERT_NE(n, nullptr);
  EXPECT_EQ(n->kind(), NodeKind::Ref);
  EXPECT_EQ(format_a1(n->as_ref()), "$A$1");
}

TEST(AstNodeRef, FormatsSheetQualifier) {
  Arena a;
  Reference r;
  r.sheet = "Sheet1";
  r.col = 0;
  r.row = 0;
  AstNode* n = make_ref(a, r);
  ASSERT_NE(n, nullptr);
  EXPECT_EQ(format_a1(n->as_ref()), "Sheet1!A1");
}

TEST(AstNodeRef, FormatsQuotedSheet) {
  Arena a;
  Reference r;
  r.sheet = "Sheet 1";
  r.sheet_quoted = true;
  r.col = 0;
  r.row = 0;
  AstNode* n = make_ref(a, r);
  ASSERT_NE(n, nullptr);
  EXPECT_EQ(format_a1(n->as_ref()), "'Sheet 1'!A1");
}

TEST(AstNodeRef, FormatsMultiLetterColumn) {
  Arena a;
  Reference r;
  r.col = 26;  // AA
  r.row = 0;
  AstNode* n = make_ref(a, r);
  ASSERT_NE(n, nullptr);
  EXPECT_EQ(format_a1(n->as_ref()), "AA1");

  Reference r2;
  r2.col = 16383;  // XFD - last Excel column
  r2.row = 1048575;
  AstNode* n2 = make_ref(a, r2);
  ASSERT_NE(n2, nullptr);
  EXPECT_EQ(format_a1(n2->as_ref()), "XFD1048576");
}

TEST(AstNodeRef, EscapesEmbeddedQuoteInSheetName) {
  Arena a;
  Reference r;
  r.sheet = "It's";
  r.sheet_quoted = true;
  r.col = 0;
  r.row = 0;
  AstNode* n = make_ref(a, r);
  ASSERT_NE(n, nullptr);
  EXPECT_EQ(format_a1(n->as_ref()), "'It''s'!A1");
}

TEST(AstNodeExternalRef, AccessorRoundtrips) {
  Arena a;
  Reference cell;
  cell.col = 0;
  cell.row = 0;
  AstNode* n = make_external_ref(a, 7, "Sheet1", cell);
  ASSERT_NE(n, nullptr);
  EXPECT_EQ(n->kind(), NodeKind::ExternalRef);
  EXPECT_EQ(n->as_external_ref_book_id(), 7u);
  EXPECT_EQ(n->as_external_ref_sheet(), "Sheet1");
  EXPECT_EQ(format_a1(n->as_external_ref_cell()), "A1");
}

TEST(AstNodeStructuredRef, AccessorRoundtripsWithColumn) {
  Arena a;
  AstNode* n = make_structured_ref(a, "Tbl", "Col", StructuredRefModifier::At);
  ASSERT_NE(n, nullptr);
  EXPECT_EQ(n->kind(), NodeKind::StructuredRef);
  EXPECT_EQ(n->as_structured_ref_table(), "Tbl");
  EXPECT_EQ(n->as_structured_ref_column(), "Col");
  EXPECT_EQ(n->as_structured_ref_modifier(), StructuredRefModifier::At);
}

TEST(AstNodeStructuredRef, WholeTableHasEmptyColumn) {
  Arena a;
  AstNode* n = make_structured_ref(a, "Tbl", "", StructuredRefModifier::None);
  ASSERT_NE(n, nullptr);
  EXPECT_EQ(n->as_structured_ref_table(), "Tbl");
  EXPECT_TRUE(n->as_structured_ref_column().empty());
  EXPECT_EQ(n->as_structured_ref_modifier(), StructuredRefModifier::None);
}

TEST(AstNodeNameRef, StringInternedToArena) {
  Arena a;
  std::string source("foo");
  AstNode* n = make_name_ref(a, source);
  ASSERT_NE(n, nullptr);
  // Mutate / destroy the source: arena copy must remain intact.
  source[0] = 'X';
  source.clear();
  EXPECT_EQ(n->as_name(), "foo");
}

TEST(AstNodeUnaryOp, StoresOpAndOperand) {
  Arena a;
  AstNode* operand = make_literal(a, Value::number(3.0));
  AstNode* n = make_unary_op(a, UnaryOp::Minus, operand);
  ASSERT_NE(n, nullptr);
  EXPECT_EQ(n->kind(), NodeKind::UnaryOp);
  EXPECT_EQ(n->as_unary_op(), UnaryOp::Minus);
  EXPECT_EQ(&n->as_unary_operand(), operand);
}

TEST(AstNodeBinaryOp, StoresKindAndChildren) {
  Arena a;
  AstNode* lhs = make_literal(a, Value::number(1.0));
  AstNode* rhs = make_literal(a, Value::number(2.0));
  AstNode* n = make_binary_op(a, BinOp::Add, lhs, rhs);
  ASSERT_NE(n, nullptr);
  EXPECT_EQ(n->kind(), NodeKind::BinaryOp);
  EXPECT_EQ(n->as_binary_op(), BinOp::Add);
  EXPECT_EQ(&n->as_binary_lhs(), lhs);
  EXPECT_EQ(&n->as_binary_rhs(), rhs);
}

TEST(AstNodeRangeOp, AccessorRoundtrips) {
  Arena a;
  Reference ra;
  ra.col = 0;
  ra.row = 0;
  Reference rb;
  rb.col = 1;
  rb.row = 1;
  AstNode* lhs = make_ref(a, ra);
  AstNode* rhs = make_ref(a, rb);
  AstNode* n = make_range_op(a, lhs, rhs);
  ASSERT_NE(n, nullptr);
  EXPECT_EQ(n->kind(), NodeKind::RangeOp);
  EXPECT_EQ(format_a1(n->as_range_lhs().as_ref()), "A1");
  EXPECT_EQ(format_a1(n->as_range_rhs().as_ref()), "B2");
}

TEST(AstNodeUnionOp, ChildArrayCopiedToArena) {
  Arena a;
  std::vector<const AstNode*> children;
  children.push_back(make_literal(a, Value::number(1.0)));
  children.push_back(make_literal(a, Value::number(2.0)));
  children.push_back(make_literal(a, Value::number(3.0)));
  AstNode* n = make_union_op(a, children.data(), 3);
  ASSERT_NE(n, nullptr);
  // Mutate the caller-side argv: AST must reference its own copy.
  children[0] = nullptr;
  children.clear();
  ASSERT_EQ(n->as_union_arity(), 3u);
  EXPECT_DOUBLE_EQ(n->as_union_child(0).as_literal().as_number(), 1.0);
  EXPECT_DOUBLE_EQ(n->as_union_child(2).as_literal().as_number(), 3.0);
}

TEST(AstNodeIntersectOp, AccessorRoundtrips) {
  Arena a;
  AstNode* lhs = make_literal(a, Value::number(1.0));
  AstNode* rhs = make_literal(a, Value::number(2.0));
  AstNode* n = make_intersect_op(a, lhs, rhs);
  ASSERT_NE(n, nullptr);
  EXPECT_EQ(n->kind(), NodeKind::IntersectOp);
  EXPECT_EQ(&n->as_intersect_lhs(), lhs);
  EXPECT_EQ(&n->as_intersect_rhs(), rhs);
}

TEST(AstNodeImplicitIntersection, StoresOperand) {
  Arena a;
  AstNode* op = make_literal(a, Value::number(5.0));
  AstNode* n = make_implicit_intersection(a, op);
  ASSERT_NE(n, nullptr);
  EXPECT_EQ(n->kind(), NodeKind::ImplicitIntersection);
  EXPECT_EQ(&n->as_implicit_intersection_operand(), op);
}

TEST(AstNodeCall, VariadicArgsCopiedToArena) {
  Arena a;
  std::vector<const AstNode*> args;
  args.push_back(make_literal(a, Value::number(1.0)));
  args.push_back(make_literal(a, Value::number(2.0)));
  args.push_back(make_literal(a, Value::number(3.0)));
  AstNode* n = make_call(a, "SUM", args.data(), 3);
  ASSERT_NE(n, nullptr);
  // Mutate caller's argv after the factory returns: AST must be unaffected.
  args[0] = nullptr;
  args[1] = nullptr;
  args[2] = nullptr;
  EXPECT_EQ(n->kind(), NodeKind::Call);
  EXPECT_EQ(n->as_call_name(), "SUM");
  EXPECT_EQ(n->as_call_arity(), 3u);
  EXPECT_DOUBLE_EQ(n->as_call_arg(0).as_literal().as_number(), 1.0);
  EXPECT_DOUBLE_EQ(n->as_call_arg(2).as_literal().as_number(), 3.0);
}

TEST(AstNodeCall, ZeroArityIsLegal) {
  Arena a;
  AstNode* n = make_call(a, "NOW", nullptr, 0);
  ASSERT_NE(n, nullptr);
  EXPECT_EQ(n->as_call_name(), "NOW");
  EXPECT_EQ(n->as_call_arity(), 0u);
}

TEST(AstNodeArrayLiteral, RowMajorAccessorRoundtrips) {
  Arena a;
  std::vector<const AstNode*> elems;
  elems.push_back(make_literal(a, Value::number(1.0)));  // (0,0)
  elems.push_back(make_literal(a, Value::number(2.0)));  // (0,1)
  elems.push_back(make_literal(a, Value::number(3.0)));  // (1,0)
  elems.push_back(make_literal(a, Value::number(4.0)));  // (1,1)
  AstNode* n = make_array_literal(a, 2, 2, elems.data());
  ASSERT_NE(n, nullptr);
  EXPECT_EQ(n->as_array_rows(), 2u);
  EXPECT_EQ(n->as_array_cols(), 2u);
  EXPECT_DOUBLE_EQ(n->as_array_element(0, 0).as_literal().as_number(), 1.0);
  EXPECT_DOUBLE_EQ(n->as_array_element(0, 1).as_literal().as_number(), 2.0);
  EXPECT_DOUBLE_EQ(n->as_array_element(1, 0).as_literal().as_number(), 3.0);
  EXPECT_DOUBLE_EQ(n->as_array_element(1, 1).as_literal().as_number(), 4.0);
}

TEST(AstNodeLambda, ParamsAndBodyRoundtrip) {
  Arena a;
  std::string p0("x");
  std::string p1("y");
  std::vector<std::string_view> params{p0, p1};
  AstNode* body = make_literal(a, Value::number(0.0));
  AstNode* n = make_lambda(a, params.data(), 2, body);
  ASSERT_NE(n, nullptr);
  // Mutate caller-side names; the AST should hold its own interned copies.
  p0[0] = 'Z';
  p1.clear();
  EXPECT_EQ(n->as_lambda_param_count(), 2u);
  EXPECT_EQ(n->as_lambda_param(0), "x");
  EXPECT_EQ(n->as_lambda_param(1), "y");
  EXPECT_EQ(&n->as_lambda_body(), body);
}

TEST(AstNodeLambda, ZeroParamsIsLegal) {
  Arena a;
  AstNode* body = make_literal(a, Value::number(7.0));
  AstNode* n = make_lambda(a, nullptr, 0, body);
  ASSERT_NE(n, nullptr);
  EXPECT_EQ(n->as_lambda_param_count(), 0u);
  EXPECT_EQ(&n->as_lambda_body(), body);
}

TEST(AstNodeLetBinding, BindingsAndBodyRoundtrip) {
  Arena a;
  std::string n0("x");
  std::string n1("y");
  std::vector<std::string_view> names{n0, n1};
  std::vector<const AstNode*> exprs;
  exprs.push_back(make_literal(a, Value::number(1.0)));
  exprs.push_back(make_literal(a, Value::number(2.0)));
  AstNode* body = make_name_ref(a, "x");
  AstNode* n = make_let_binding(a, names.data(), exprs.data(), 2, body);
  ASSERT_NE(n, nullptr);
  EXPECT_EQ(n->as_let_binding_count(), 2u);
  EXPECT_EQ(n->as_let_binding_name(0), "x");
  EXPECT_EQ(n->as_let_binding_name(1), "y");
  EXPECT_DOUBLE_EQ(n->as_let_binding_expr(0).as_literal().as_number(), 1.0);
  EXPECT_DOUBLE_EQ(n->as_let_binding_expr(1).as_literal().as_number(), 2.0);
  EXPECT_EQ(&n->as_let_body(), body);
}

TEST(AstNodeLambdaCall, CalleeAndArgsRoundtrip) {
  Arena a;
  AstNode* callee = make_name_ref(a, "myfn");
  std::vector<const AstNode*> args;
  args.push_back(make_literal(a, Value::number(10.0)));
  args.push_back(make_literal(a, Value::number(20.0)));
  AstNode* n = make_lambda_call(a, callee, args.data(), 2);
  ASSERT_NE(n, nullptr);
  EXPECT_EQ(n->kind(), NodeKind::LambdaCall);
  EXPECT_EQ(&n->as_lambda_call_callee(), callee);
  EXPECT_EQ(n->as_lambda_call_arity(), 2u);
  EXPECT_DOUBLE_EQ(n->as_lambda_call_arg(0).as_literal().as_number(), 10.0);
  EXPECT_DOUBLE_EQ(n->as_lambda_call_arg(1).as_literal().as_number(), 20.0);
}

TEST(AstNodeErrorLiteral, AccessorRoundtrips) {
  Arena a;
  AstNode* n = make_error_literal(a, ErrorCode::Div0);
  ASSERT_NE(n, nullptr);
  EXPECT_EQ(n->kind(), NodeKind::ErrorLiteral);
  EXPECT_EQ(n->as_error_literal(), ErrorCode::Div0);
}

TEST(AstNodeErrorPlaceholder, FactoryReturnsCorrectKind) {
  Arena a;
  AstNode* n = make_error_placeholder(a);
  ASSERT_NE(n, nullptr);
  EXPECT_EQ(n->kind(), NodeKind::ErrorPlaceholder);
  // No payload accessor; default range is zero-initialised.
  EXPECT_EQ(n->range().start, 0u);
  EXPECT_EQ(n->range().end, 0u);
}

TEST(AstNodeRange, SetRangeIsRoundTripped) {
  Arena a;
  AstNode* n = make_literal(a, Value::number(1.0));
  ASSERT_NE(n, nullptr);
  TextRange r;
  r.start = 5;
  r.end = 10;
  r.line = 2;
  r.column = 3;
  n->set_range(r);
  EXPECT_EQ(n->range().start, 5u);
  EXPECT_EQ(n->range().end, 10u);
  EXPECT_EQ(n->range().line, 2u);
  EXPECT_EQ(n->range().column, 3u);
}

}  // namespace
}  // namespace parser
}  // namespace formulon
