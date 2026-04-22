// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for the info / type-query built-in functions:
// ISNUMBER, ISTEXT, ISBLANK, ISLOGICAL, ISERROR, ISERR, ISNA, N, T.
//
// The IS* family is registered with `propagate_errors = false` so the
// dispatcher in `tree_walker` hands them error-typed inputs verbatim;
// these tests pin both the per-kind result table and the
// error-passthrough behaviour. `N` and `T` use the default propagation
// rule (errors short-circuit before the body runs); their tests confirm
// that as well as the per-kind coercion table.
//
// Most cases are exercised end-to-end via `EvalSource` (parse + evaluate)
// so the dispatcher path is in scope. ISBLANK's TRUE case is exercised by
// invoking the impl directly through the registry, because the formula
// grammar has no syntax for a Blank literal.

#include <cstdint>
#include <string_view>

#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Parses `src` and evaluates it via the default function registry. The
// thread-local arenas keep text payloads readable for the immediately
// following EXPECT_*. Each call resets the arenas to avoid cross-test
// contamination.
Value EvalSource(std::string_view src) {
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

// Invokes a registered function impl directly with a single argument.
// Used for cases that cannot be expressed in formula syntax (notably the
// Blank scalar, which has no literal form).
Value CallSingle(std::string_view name, const Value& arg) {
  static thread_local Arena arena;
  arena.reset();
  const FunctionDef* def = default_registry().lookup(name);
  EXPECT_NE(def, nullptr) << "function not registered: " << name;
  if (def == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return def->impl(&arg, 1u, arena);
}

// ---------------------------------------------------------------------------
// ISNUMBER
// ---------------------------------------------------------------------------

TEST(BuiltinsIsNumber, NumberLiteralIsTrue) {
  const Value v = EvalSource("=ISNUMBER(1)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsIsNumber, NegativeFractionIsTrue) {
  const Value v = EvalSource("=ISNUMBER(-3.14)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsIsNumber, TextIsFalse) {
  const Value v = EvalSource("=ISNUMBER(\"1\")");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsIsNumber, BoolIsFalse) {
  const Value v = EvalSource("=ISNUMBER(TRUE)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsIsNumber, ErrorIsFalseNotPropagated) {
  // Confirms `propagate_errors = false`: the error must reach the body and
  // surface as boolean FALSE, not as #DIV/0!.
  const Value v = EvalSource("=ISNUMBER(#DIV/0!)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsIsNumber, BlankInputIsFalse) {
  const Value v = CallSingle("ISNUMBER", Value::blank());
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsIsNumber, ZeroArgsIsArityViolation) {
  const Value v = EvalSource("=ISNUMBER()");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsIsNumber, TwoArgsIsArityViolation) {
  const Value v = EvalSource("=ISNUMBER(1, 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// ISTEXT
// ---------------------------------------------------------------------------

TEST(BuiltinsIsText, TextLiteralIsTrue) {
  const Value v = EvalSource("=ISTEXT(\"hello\")");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsIsText, EmptyTextIsTrue) {
  const Value v = EvalSource("=ISTEXT(\"\")");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsIsText, NumberIsFalse) {
  const Value v = EvalSource("=ISTEXT(1)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsIsText, BoolIsFalse) {
  const Value v = EvalSource("=ISTEXT(FALSE)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsIsText, ErrorIsFalseNotPropagated) {
  const Value v = EvalSource("=ISTEXT(#DIV/0!)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsIsText, BlankInputIsFalse) {
  const Value v = CallSingle("ISTEXT", Value::blank());
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

// ---------------------------------------------------------------------------
// ISBLANK
// ---------------------------------------------------------------------------

TEST(BuiltinsIsBlank, BlankInputIsTrue) {
  // The grammar has no Blank literal; invoke the impl directly.
  const Value v = CallSingle("ISBLANK", Value::blank());
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsIsBlank, EmptyTextIsFalse) {
  // Excel quirk: "" is Text, not Blank.
  const Value v = EvalSource("=ISBLANK(\"\")");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsIsBlank, NumberIsFalse) {
  const Value v = EvalSource("=ISBLANK(0)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsIsBlank, BoolIsFalse) {
  const Value v = EvalSource("=ISBLANK(FALSE)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsIsBlank, ErrorIsFalseNotPropagated) {
  const Value v = EvalSource("=ISBLANK(#DIV/0!)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

// ---------------------------------------------------------------------------
// ISLOGICAL
// ---------------------------------------------------------------------------

TEST(BuiltinsIsLogical, TrueLiteralIsTrue) {
  const Value v = EvalSource("=ISLOGICAL(TRUE)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsIsLogical, FalseLiteralIsTrue) {
  const Value v = EvalSource("=ISLOGICAL(FALSE)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsIsLogical, NumberZeroIsFalse) {
  // Numeric 0 is NOT Bool: only the actual TRUE/FALSE booleans qualify.
  const Value v = EvalSource("=ISLOGICAL(0)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsIsLogical, NumberOneIsFalse) {
  const Value v = EvalSource("=ISLOGICAL(1)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsIsLogical, NumericTextIsFalse) {
  const Value v = EvalSource("=ISLOGICAL(\"TRUE\")");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsIsLogical, ErrorIsFalseNotPropagated) {
  const Value v = EvalSource("=ISLOGICAL(#DIV/0!)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsIsLogical, BlankInputIsFalse) {
  const Value v = CallSingle("ISLOGICAL", Value::blank());
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

// ---------------------------------------------------------------------------
// ISERROR
// ---------------------------------------------------------------------------

TEST(BuiltinsIsError, Div0IsTrue) {
  const Value v = EvalSource("=ISERROR(#DIV/0!)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsIsError, NaIsTrue) {
  // ISERROR catches everything including #N/A; ISERR (next group) does not.
  const Value v = EvalSource("=ISERROR(#N/A)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsIsError, RefIsTrue) {
  const Value v = EvalSource("=ISERROR(#REF!)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsIsError, ValueIsTrue) {
  const Value v = EvalSource("=ISERROR(#VALUE!)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsIsError, ComputedDivZeroIsTrue) {
  // Confirms the dispatcher hands the materialised error to the body
  // rather than short-circuiting it (the smoking-gun for the
  // `propagate_errors = false` wiring on a non-literal trigger).
  const Value v = EvalSource("=ISERROR(1/0)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsIsError, NumberIsFalse) {
  const Value v = EvalSource("=ISERROR(1)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsIsError, TextIsFalse) {
  const Value v = EvalSource("=ISERROR(\"oops\")");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsIsError, BoolIsFalse) {
  const Value v = EvalSource("=ISERROR(TRUE)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

// ---------------------------------------------------------------------------
// ISERR (error EXCEPT #N/A)
// ---------------------------------------------------------------------------

TEST(BuiltinsIsErr, Div0IsTrue) {
  const Value v = EvalSource("=ISERR(#DIV/0!)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsIsErr, RefIsTrue) {
  const Value v = EvalSource("=ISERR(#REF!)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsIsErr, ValueIsTrue) {
  const Value v = EvalSource("=ISERR(#VALUE!)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsIsErr, NameIsTrue) {
  const Value v = EvalSource("=ISERR(#NAME?)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsIsErr, NumIsTrue) {
  const Value v = EvalSource("=ISERR(#NUM!)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsIsErr, NaIsFalse) {
  // The defining quirk: ISERR(#N/A) is FALSE.
  const Value v = EvalSource("=ISERR(#N/A)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsIsErr, NumberIsFalse) {
  const Value v = EvalSource("=ISERR(42)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsIsErr, TextIsFalse) {
  const Value v = EvalSource("=ISERR(\"x\")");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

// ---------------------------------------------------------------------------
// ISNA (only #N/A)
// ---------------------------------------------------------------------------

TEST(BuiltinsIsNa, NaIsTrue) {
  const Value v = EvalSource("=ISNA(#N/A)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsIsNa, Div0IsFalse) {
  const Value v = EvalSource("=ISNA(#DIV/0!)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsIsNa, RefIsFalse) {
  const Value v = EvalSource("=ISNA(#REF!)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsIsNa, ValueIsFalse) {
  const Value v = EvalSource("=ISNA(#VALUE!)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsIsNa, NumberIsFalse) {
  const Value v = EvalSource("=ISNA(0)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsIsNa, TextIsFalse) {
  const Value v = EvalSource("=ISNA(\"#N/A\")");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

// ---------------------------------------------------------------------------
// N (numeric coercion; never parses text)
// ---------------------------------------------------------------------------

TEST(BuiltinsN, NumberPassesThrough) {
  const Value v = EvalSource("=N(42)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 42.0);
}

TEST(BuiltinsN, NegativeFractionPassesThrough) {
  const Value v = EvalSource("=N(-3.5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -3.5);
}

TEST(BuiltinsN, TrueIsOne) {
  const Value v = EvalSource("=N(TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsN, FalseIsZero) {
  const Value v = EvalSource("=N(FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsN, AnyTextIsZero) {
  const Value v = EvalSource("=N(\"hello\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsN, NumericTextIsZeroNotParsed) {
  // CRITICAL: N never parses text. Contrast with VALUE, which does. This
  // is the divergence from the more permissive `coerce_to_number` path.
  const Value v = EvalSource("=N(\"123\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsN, EmptyTextIsZero) {
  const Value v = EvalSource("=N(\"\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsN, BlankIsZero) {
  const Value v = CallSingle("N", Value::blank());
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsN, ErrorPropagates) {
  const Value v = EvalSource("=N(#DIV/0!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsN, NaPropagates) {
  const Value v = EvalSource("=N(#N/A)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsN, ZeroArgsIsArityViolation) {
  const Value v = EvalSource("=N()");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// T (text coercion; non-text becomes "")
// ---------------------------------------------------------------------------

TEST(BuiltinsT, TextPassesThrough) {
  const Value v = EvalSource("=T(\"hello\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "hello");
}

TEST(BuiltinsT, EmptyTextPassesThrough) {
  const Value v = EvalSource("=T(\"\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "");
}

TEST(BuiltinsT, NumberBecomesEmptyText) {
  const Value v = EvalSource("=T(42)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "");
}

TEST(BuiltinsT, ZeroBecomesEmptyText) {
  // Important corner: T(0) is "", NOT "0". T does not stringify numbers.
  const Value v = EvalSource("=T(0)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "");
}

TEST(BuiltinsT, BoolBecomesEmptyText) {
  const Value v = EvalSource("=T(TRUE)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "");
}

TEST(BuiltinsT, BlankBecomesEmptyText) {
  const Value v = CallSingle("T", Value::blank());
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "");
}

TEST(BuiltinsT, ErrorPropagates) {
  const Value v = EvalSource("=T(#DIV/0!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsT, NaPropagates) {
  const Value v = EvalSource("=T(#N/A)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsT, ZeroArgsIsArityViolation) {
  const Value v = EvalSource("=T()");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsT, TwoArgsIsArityViolation) {
  const Value v = EvalSource("=T(\"a\", \"b\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
