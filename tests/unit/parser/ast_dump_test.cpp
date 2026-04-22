// Copyright 2026 libraz. Licensed under the MIT License.
//
// Golden tests for `dump_sexpr`. These outputs become the parser corpus
// contract: every change here is a contract change.

#include "parser/ast_dump.h"

#include <cmath>
#include <limits>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/reference.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace parser {
namespace {

// ---------------------------------------------------------------------------
// Literals
// ---------------------------------------------------------------------------

TEST(AstDumpLiteral, NumberInteger) {
  Arena a;
  AstNode* n = make_literal(a, Value::number(42.0));
  EXPECT_EQ(dump_sexpr(*n), "(num 42)");
}

TEST(AstDumpLiteral, NumberNegativeInteger) {
  Arena a;
  AstNode* n = make_literal(a, Value::number(-7.0));
  EXPECT_EQ(dump_sexpr(*n), "(num -7)");
}

TEST(AstDumpLiteral, NumberFractionTrimsTrailingZeros) {
  Arena a;
  AstNode* n = make_literal(a, Value::number(3.14));
  EXPECT_EQ(dump_sexpr(*n), "(num 3.14)");
}

TEST(AstDumpLiteral, NumberTwoPointZeroIsInteger) {
  Arena a;
  AstNode* n = make_literal(a, Value::number(2.0));
  EXPECT_EQ(dump_sexpr(*n), "(num 2)");
}

TEST(AstDumpLiteral, NumberNaN) {
  Arena a;
  AstNode* n = make_literal(a, Value::number(std::numeric_limits<double>::quiet_NaN()));
  EXPECT_EQ(dump_sexpr(*n), "(num nan)");
}

TEST(AstDumpLiteral, NumberPositiveInfinity) {
  Arena a;
  AstNode* n = make_literal(a, Value::number(std::numeric_limits<double>::infinity()));
  EXPECT_EQ(dump_sexpr(*n), "(num inf)");
}

TEST(AstDumpLiteral, NumberNegativeInfinity) {
  Arena a;
  AstNode* n = make_literal(a, Value::number(-std::numeric_limits<double>::infinity()));
  EXPECT_EQ(dump_sexpr(*n), "(num -inf)");
}

TEST(AstDumpLiteral, NegativeZeroPrintsAsZero) {
  Arena a;
  AstNode* n = make_literal(a, Value::number(-0.0));
  EXPECT_EQ(dump_sexpr(*n), "(num 0)");
}

TEST(AstDumpLiteral, BoolTrue) {
  Arena a;
  AstNode* n = make_literal(a, Value::boolean(true));
  EXPECT_EQ(dump_sexpr(*n), "(bool true)");
}

TEST(AstDumpLiteral, BoolFalse) {
  Arena a;
  AstNode* n = make_literal(a, Value::boolean(false));
  EXPECT_EQ(dump_sexpr(*n), "(bool false)");
}

TEST(AstDumpLiteral, Blank) {
  Arena a;
  AstNode* n = make_literal(a, Value::blank());
  EXPECT_EQ(dump_sexpr(*n), "(blank)");
}

TEST(AstDumpLiteral, ErrorValue) {
  Arena a;
  AstNode* n = make_literal(a, Value::error(ErrorCode::Div0));
  EXPECT_EQ(dump_sexpr(*n), "(err #DIV/0!)");
}

// ---------------------------------------------------------------------------
// References
// ---------------------------------------------------------------------------

TEST(AstDumpRef, PlainA1) {
  Arena a;
  Reference r;
  r.col = 0;
  r.row = 0;
  AstNode* n = make_ref(a, r);
  EXPECT_EQ(dump_sexpr(*n), "(ref A1)");
}

TEST(AstDumpRef, FullyAbsoluteWithSheet) {
  Arena a;
  Reference r;
  r.sheet = "Sheet1";
  r.col = 0;
  r.row = 0;
  r.col_abs = true;
  r.row_abs = true;
  AstNode* n = make_ref(a, r);
  EXPECT_EQ(dump_sexpr(*n), "(ref Sheet1!$A$1)");
}

TEST(AstDumpRef, QuotedSheet) {
  Arena a;
  Reference r;
  r.sheet = "Sheet 1";
  r.sheet_quoted = true;
  r.col = 0;
  r.row = 0;
  AstNode* n = make_ref(a, r);
  EXPECT_EQ(dump_sexpr(*n), "(ref 'Sheet 1'!A1)");
}

TEST(AstDumpRef, FullColumn) {
  Arena a;
  Reference r;
  r.col = 0;
  r.is_full_col = true;
  AstNode* n = make_ref(a, r);
  EXPECT_EQ(dump_sexpr(*n), "(ref A:A)");
}

TEST(AstDumpRef, FullColumnAbsolute) {
  Arena a;
  Reference r;
  r.col = 0;
  r.col_abs = true;
  r.is_full_col = true;
  AstNode* n = make_ref(a, r);
  EXPECT_EQ(dump_sexpr(*n), "(ref $A:$A)");
}

TEST(AstDumpRef, FullRow) {
  Arena a;
  Reference r;
  r.row = 0;
  r.is_full_row = true;
  AstNode* n = make_ref(a, r);
  EXPECT_EQ(dump_sexpr(*n), "(ref 1:1)");
}

TEST(AstDumpRef, FullColumnSheetQualified) {
  Arena a;
  Reference r;
  r.sheet = "Sheet1";
  r.col = 0;
  r.is_full_col = true;
  AstNode* n = make_ref(a, r);
  EXPECT_EQ(dump_sexpr(*n), "(ref Sheet1!A:A)");
}

// ---------------------------------------------------------------------------
// External / structured / name refs
// ---------------------------------------------------------------------------

TEST(AstDumpExternalRef, FormatsBookSheetCell) {
  Arena a;
  Reference cell;
  cell.col = 0;
  cell.row = 0;
  AstNode* n = make_external_ref(a, 1, "Sheet1", cell);
  EXPECT_EQ(dump_sexpr(*n), "(ext-ref [1] Sheet1 A1)");
}

TEST(AstDumpStructuredRef, AtModifier) {
  Arena a;
  AstNode* n = make_structured_ref(a, "Tbl", "Col", StructuredRefModifier::At);
  EXPECT_EQ(dump_sexpr(*n), "(struct-ref Tbl Col @)");
}

TEST(AstDumpStructuredRef, DataModifier) {
  Arena a;
  AstNode* n = make_structured_ref(a, "Tbl", "Col", StructuredRefModifier::Data);
  EXPECT_EQ(dump_sexpr(*n), "(struct-ref Tbl Col #data)");
}

TEST(AstDumpStructuredRef, HeadersModifier) {
  Arena a;
  AstNode* n = make_structured_ref(a, "Tbl", "Col", StructuredRefModifier::Headers);
  EXPECT_EQ(dump_sexpr(*n), "(struct-ref Tbl Col #headers)");
}

TEST(AstDumpStructuredRef, TotalsModifier) {
  Arena a;
  AstNode* n = make_structured_ref(a, "Tbl", "Col", StructuredRefModifier::Totals);
  EXPECT_EQ(dump_sexpr(*n), "(struct-ref Tbl Col #totals)");
}

TEST(AstDumpStructuredRef, AllModifier) {
  Arena a;
  AstNode* n = make_structured_ref(a, "Tbl", "Col", StructuredRefModifier::All);
  EXPECT_EQ(dump_sexpr(*n), "(struct-ref Tbl Col #all)");
}

TEST(AstDumpStructuredRef, WholeTableNoColumnNoModifier) {
  Arena a;
  AstNode* n = make_structured_ref(a, "Tbl", "", StructuredRefModifier::None);
  EXPECT_EQ(dump_sexpr(*n), "(struct-ref Tbl)");
}

TEST(AstDumpStructuredRef, WholeTableAllOnly) {
  Arena a;
  AstNode* n = make_structured_ref(a, "Tbl", "", StructuredRefModifier::All);
  EXPECT_EQ(dump_sexpr(*n), "(struct-ref Tbl #all)");
}

TEST(AstDumpStructuredRef, ColumnOnlyDefaultModifier) {
  Arena a;
  AstNode* n = make_structured_ref(a, "Tbl", "Col", StructuredRefModifier::None);
  EXPECT_EQ(dump_sexpr(*n), "(struct-ref Tbl Col)");
}

TEST(AstDumpNameRef, Plain) {
  Arena a;
  AstNode* n = make_name_ref(a, "foo");
  EXPECT_EQ(dump_sexpr(*n), "(name foo)");
}

// ---------------------------------------------------------------------------
// Operators
// ---------------------------------------------------------------------------

TEST(AstDumpUnary, MinusOperator) {
  Arena a;
  AstNode* n = make_unary_op(a, UnaryOp::Minus, make_name_ref(a, "x"));
  EXPECT_EQ(dump_sexpr(*n), "(unary - (name x))");
}

TEST(AstDumpUnary, PlusOperator) {
  Arena a;
  AstNode* n = make_unary_op(a, UnaryOp::Plus, make_name_ref(a, "x"));
  EXPECT_EQ(dump_sexpr(*n), "(unary + (name x))");
}

TEST(AstDumpUnary, PercentOperator) {
  Arena a;
  AstNode* n = make_unary_op(a, UnaryOp::Percent, make_name_ref(a, "x"));
  EXPECT_EQ(dump_sexpr(*n), "(unary % (name x))");
}

TEST(AstDumpBinary, AllOperators) {
  Arena a;
  AstNode* l = make_name_ref(a, "a");
  AstNode* r = make_name_ref(a, "b");
  struct Case {
    BinOp op;
    const char* token;
  };
  const Case cases[] = {
      {BinOp::Add, "+"}, {BinOp::Sub, "-"},    {BinOp::Mul, "*"}, {BinOp::Div, "/"},
      {BinOp::Pow, "^"}, {BinOp::Concat, "&"}, {BinOp::Eq, "="},  {BinOp::NotEq, "<>"},
      {BinOp::Lt, "<"},  {BinOp::LtEq, "<="},  {BinOp::Gt, ">"},  {BinOp::GtEq, ">="},
  };
  for (const auto& c : cases) {
    AstNode* n = make_binary_op(a, c.op, l, r);
    std::string expected = std::string("(binary ") + c.token + " (name a) (name b))";
    EXPECT_EQ(dump_sexpr(*n), expected) << "op token: " << c.token;
  }
}

TEST(AstDumpRange, BasicA1B2) {
  Arena a;
  Reference ra;
  ra.col = 0;
  ra.row = 0;
  Reference rb;
  rb.col = 1;
  rb.row = 1;
  AstNode* n = make_range_op(a, make_ref(a, ra), make_ref(a, rb));
  EXPECT_EQ(dump_sexpr(*n), "(range (ref A1) (ref B2))");
}

TEST(AstDumpUnion, ThreeChildren) {
  Arena a;
  std::vector<const AstNode*> children;
  children.push_back(make_name_ref(a, "a"));
  children.push_back(make_name_ref(a, "b"));
  children.push_back(make_name_ref(a, "c"));
  AstNode* n = make_union_op(a, children.data(), 3);
  EXPECT_EQ(dump_sexpr(*n), "(union (name a) (name b) (name c))");
}

TEST(AstDumpIntersect, TwoChildren) {
  Arena a;
  AstNode* n = make_intersect_op(a, make_name_ref(a, "a"), make_name_ref(a, "b"));
  EXPECT_EQ(dump_sexpr(*n), "(intersect (name a) (name b))");
}

TEST(AstDumpImplicitIntersection, AtPrefix) {
  Arena a;
  AstNode* n = make_implicit_intersection(a, make_name_ref(a, "x"));
  EXPECT_EQ(dump_sexpr(*n), "(at (name x))");
}

// ---------------------------------------------------------------------------
// Calls / arrays
// ---------------------------------------------------------------------------

TEST(AstDumpCall, ThreeArgSum) {
  Arena a;
  std::vector<const AstNode*> args;
  args.push_back(make_literal(a, Value::number(1.0)));
  args.push_back(make_literal(a, Value::number(2.0)));
  args.push_back(make_literal(a, Value::number(3.0)));
  AstNode* n = make_call(a, "SUM", args.data(), 3);
  EXPECT_EQ(dump_sexpr(*n), "(call SUM (num 1) (num 2) (num 3))");
}

TEST(AstDumpCall, ZeroArg) {
  Arena a;
  AstNode* n = make_call(a, "NOW", nullptr, 0);
  EXPECT_EQ(dump_sexpr(*n), "(call NOW)");
}

TEST(AstDumpArrayLiteral, TwoByTwo) {
  Arena a;
  std::vector<const AstNode*> elems;
  elems.push_back(make_literal(a, Value::number(1.0)));
  elems.push_back(make_literal(a, Value::number(2.0)));
  elems.push_back(make_literal(a, Value::number(3.0)));
  elems.push_back(make_literal(a, Value::number(4.0)));
  AstNode* n = make_array_literal(a, 2, 2, elems.data());
  EXPECT_EQ(dump_sexpr(*n), "(array 2 2 (num 1) (num 2) (num 3) (num 4))");
}

// ---------------------------------------------------------------------------
// Lambda / let / lambda call
// ---------------------------------------------------------------------------

TEST(AstDumpLambda, TwoParams) {
  Arena a;
  std::vector<std::string_view> params{"x", "y"};
  AstNode* body = make_name_ref(a, "x");
  AstNode* n = make_lambda(a, params.data(), 2, body);
  EXPECT_EQ(dump_sexpr(*n), "(lambda (x y) (name x))");
}

TEST(AstDumpLambda, ZeroParams) {
  Arena a;
  AstNode* body = make_literal(a, Value::number(7.0));
  AstNode* n = make_lambda(a, nullptr, 0, body);
  EXPECT_EQ(dump_sexpr(*n), "(lambda () (num 7))");
}

TEST(AstDumpLet, SingleBinding) {
  Arena a;
  std::vector<std::string_view> names{"x"};
  std::vector<const AstNode*> exprs;
  exprs.push_back(make_literal(a, Value::number(1.0)));
  AstNode* body = make_name_ref(a, "x");
  AstNode* n = make_let_binding(a, names.data(), exprs.data(), 1, body);
  EXPECT_EQ(dump_sexpr(*n), "(let ((x (num 1))) (name x))");
}

TEST(AstDumpLet, TwoBindings) {
  Arena a;
  std::vector<std::string_view> names{"x", "y"};
  std::vector<const AstNode*> exprs;
  exprs.push_back(make_literal(a, Value::number(1.0)));
  exprs.push_back(make_literal(a, Value::number(2.0)));
  AstNode* body = make_name_ref(a, "x");
  AstNode* n = make_let_binding(a, names.data(), exprs.data(), 2, body);
  EXPECT_EQ(dump_sexpr(*n), "(let ((x (num 1)) (y (num 2))) (name x))");
}

TEST(AstDumpLambdaCall, TwoArgs) {
  Arena a;
  AstNode* callee = make_name_ref(a, "fn");
  std::vector<const AstNode*> args;
  args.push_back(make_name_ref(a, "a"));
  args.push_back(make_name_ref(a, "b"));
  AstNode* n = make_lambda_call(a, callee, args.data(), 2);
  EXPECT_EQ(dump_sexpr(*n), "(lambda-call (name fn) (name a) (name b))");
}

TEST(AstDumpLambdaCall, ZeroArgs) {
  Arena a;
  AstNode* callee = make_name_ref(a, "fn");
  AstNode* n = make_lambda_call(a, callee, nullptr, 0);
  EXPECT_EQ(dump_sexpr(*n), "(lambda-call (name fn))");
}

TEST(AstDumpErrorLiteral, Div0) {
  Arena a;
  AstNode* n = make_error_literal(a, ErrorCode::Div0);
  EXPECT_EQ(dump_sexpr(*n), "(err-lit #DIV/0!)");
}

TEST(AstDumpErrorLiteral, NA) {
  Arena a;
  AstNode* n = make_error_literal(a, ErrorCode::NA);
  EXPECT_EQ(dump_sexpr(*n), "(err-lit #N/A)");
}

TEST(AstDumpErrorPlaceholder, DumpsAsErrorSexpr) {
  Arena a;
  AstNode* n = make_error_placeholder(a);
  ASSERT_NE(n, nullptr);
  EXPECT_EQ(dump_sexpr(*n), "(error)");
}

// ---------------------------------------------------------------------------
// Composite trees
// ---------------------------------------------------------------------------

TEST(AstDumpComposite, SumPlusThree) {
  Arena a;
  std::vector<const AstNode*> args;
  args.push_back(make_literal(a, Value::number(1.0)));
  args.push_back(make_literal(a, Value::number(2.0)));
  AstNode* call = make_call(a, "SUM", args.data(), 2);
  AstNode* n = make_binary_op(a, BinOp::Add, call, make_literal(a, Value::number(3.0)));
  EXPECT_EQ(dump_sexpr(*n), "(binary + (call SUM (num 1) (num 2)) (num 3))");
}

TEST(AstDumpComposite, NestedLambdaCall) {
  // Models =LAMBDA(x, x+1)(5)
  Arena a;
  std::vector<std::string_view> params{"x"};
  AstNode* body = make_binary_op(a, BinOp::Add, make_name_ref(a, "x"), make_literal(a, Value::number(1.0)));
  AstNode* lam = make_lambda(a, params.data(), 1, body);
  std::vector<const AstNode*> args;
  args.push_back(make_literal(a, Value::number(5.0)));
  AstNode* n = make_lambda_call(a, lam, args.data(), 1);
  EXPECT_EQ(dump_sexpr(*n), "(lambda-call (lambda (x) (binary + (name x) (num 1))) (num 5))");
}

TEST(AstDumpComposite, LetWithCall) {
  // Models =LET(x, 10, SUM(x, 1))
  Arena a;
  std::vector<std::string_view> names{"x"};
  std::vector<const AstNode*> binds;
  binds.push_back(make_literal(a, Value::number(10.0)));
  std::vector<const AstNode*> args;
  args.push_back(make_name_ref(a, "x"));
  args.push_back(make_literal(a, Value::number(1.0)));
  AstNode* body = make_call(a, "SUM", args.data(), 2);
  AstNode* n = make_let_binding(a, names.data(), binds.data(), 1, body);
  EXPECT_EQ(dump_sexpr(*n), "(let ((x (num 10))) (call SUM (name x) (num 1)))");
}

TEST(AstDumpComposite, RangeUnionInsideSum) {
  // Models =SUM((A1:B2,C1:D2))
  Arena a;
  Reference a1;
  a1.col = 0;
  a1.row = 0;
  Reference b2;
  b2.col = 1;
  b2.row = 1;
  Reference c1;
  c1.col = 2;
  c1.row = 0;
  Reference d2;
  d2.col = 3;
  d2.row = 1;
  AstNode* range1 = make_range_op(a, make_ref(a, a1), make_ref(a, b2));
  AstNode* range2 = make_range_op(a, make_ref(a, c1), make_ref(a, d2));
  std::vector<const AstNode*> union_kids;
  union_kids.push_back(range1);
  union_kids.push_back(range2);
  AstNode* uni = make_union_op(a, union_kids.data(), 2);
  std::vector<const AstNode*> sum_args;
  sum_args.push_back(uni);
  AstNode* n = make_call(a, "SUM", sum_args.data(), 1);
  EXPECT_EQ(dump_sexpr(*n), "(call SUM (union (range (ref A1) (ref B2)) (range (ref C1) (ref D2))))");
}

}  // namespace
}  // namespace parser
}  // namespace formulon
