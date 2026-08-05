//
// Tests for the service-stub family: IMAGE, RTD, TRANSLATE,
// DETECTLANGUAGE, COPILOT, plus PHONETIC (lazy) and ISOMITTED (lazy).
// Each returns a deterministic Excel-visible surface (see
// `src/eval/builtins/service_stubs.cpp` for the rationale on the eager
// stubs; `phonetic_lazy.cpp` for the lazy PHONETIC contract). The
// eager stubs ride `propagate_errors = true` so an error argument
// short-circuits before the fixed return fires. PHONETIC is a lazy
// form that reads `Cell::phonetic_text` off the un-evaluated Reference
// AST and falls through to a value-based passthrough surface (text ->
// text, blank -> "", non-text -> #N/A, errors propagate). ISOMITTED is
// registered as a lazy form so it can inspect any argument shape
// (including errors) and report "present".
//
// GETPIVOTDATA is no longer a stub; it is wired through
// `eval_getpivotdata_lazy` and exercised in `getpivotdata_lazy_test.cpp`
// where the workbook fixture lets the Ref-anchor lookup resolve a real
// pivot table. The registry-pin check below still confirms that the
// name is wired.
//
// The tests here pin:
//
//   * the nominal registry entries exist for every eager stub;
//   * each eager stub returns its documented fixed surface on valid
//     inputs and propagates argument errors;
//   * PHONETIC's non-Ref passthrough surface (the Ref path lives in
//     `phonetic_lazy_test.cpp` because it requires a workbook fixture);
//   * ISOMITTED's outside-lambda surface (errors must NOT propagate).

#include <string_view>

#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "test_eval_helpers.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Parses `src` and evaluates it via the default function registry.
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
  return evaluate(*root, eval_arena, default_registry(), test::mac_context());
}

// ---------------------------------------------------------------------------
// Registry pin
// ---------------------------------------------------------------------------

TEST(BuiltinsServiceStubRegistry, NamesRegistered) {
  const FunctionRegistry& reg = default_registry();
  EXPECT_NE(reg.lookup("IMAGE"), nullptr);
  EXPECT_NE(reg.lookup("RTD"), nullptr);
  EXPECT_NE(reg.lookup("TRANSLATE"), nullptr);
  EXPECT_NE(reg.lookup("DETECTLANGUAGE"), nullptr);
  EXPECT_NE(reg.lookup("COPILOT"), nullptr);
  // GETPIVOTDATA is intentionally not in the eager registry: it is a
  // lazy form (see `eval_getpivotdata_lazy`) so it can recover the
  // anchor's (sheet, row, col) from the un-evaluated Reference AST and
  // resolve a pivot table on the bound workbook.
  EXPECT_EQ(reg.lookup("GETPIVOTDATA"), nullptr);
  // PHONETIC is intentionally not in the eager registry: it is a lazy
  // form (see `eval_phonetic_lazy`) so it can read the referenced
  // cell's `<rPh>` annotation off the un-evaluated Reference AST.
  EXPECT_EQ(reg.lookup("PHONETIC"), nullptr);
  // ISOMITTED is intentionally not in the eager registry: it is a lazy
  // form (see `eval_isomitted_lazy`) so it can inspect the argument's
  // AST shape and the active NameEnv's omitted flag.
}

// ---------------------------------------------------------------------------
// IMAGE - #VALUE!
// ---------------------------------------------------------------------------

TEST(BuiltinsServiceStubImage, ReturnsValue) {
  const Value v = EvalSource("=IMAGE(\"http://x\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsServiceStubImage, ArgumentErrorPropagates) {
  // 1/0 -> #DIV/0! wins over the nominal #VALUE! stub return.
  const Value v = EvalSource("=IMAGE(1/0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// RTD - #N/A
// ---------------------------------------------------------------------------

TEST(BuiltinsServiceStubRtd, ReturnsNA) {
  const Value v = EvalSource("=RTD(\"prog\", \"srv\", \"topic\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsServiceStubRtd, ArgumentErrorPropagates) {
  const Value v = EvalSource("=RTD(1/0, \"srv\", \"topic\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// TRANSLATE - #NAME?
// ---------------------------------------------------------------------------

TEST(BuiltinsServiceStubTranslate, ReturnsName) {
  const Value v = EvalSource("=TRANSLATE(\"hello\", \"en\", \"ja\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(BuiltinsServiceStubTranslate, ArityTwoAlsoWorks) {
  // Two-arg form is accepted (min_arity = 2 for auto-detect source).
  const Value v = EvalSource("=TRANSLATE(\"hello\", \"ja\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(BuiltinsServiceStubTranslate, ArgumentErrorPropagates) {
  const Value v = EvalSource("=TRANSLATE(\"hello\", 1/0, \"ja\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// DETECTLANGUAGE - #NAME?
// ---------------------------------------------------------------------------

TEST(BuiltinsServiceStubDetectLanguage, ReturnsName) {
  // Literal Japanese input still produces #NAME? because the stub body
  // never consults its argument.
  const Value v = EvalSource("=DETECTLANGUAGE(\"\xE3\x81\x93\xE3\x82\x93\xE3\x81\xAB\xE3\x81\xA1\xE3\x81\xAF\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(BuiltinsServiceStubDetectLanguage, ArgumentErrorPropagates) {
  const Value v = EvalSource("=DETECTLANGUAGE(1/0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// COPILOT - #NAME?
// ---------------------------------------------------------------------------

TEST(BuiltinsServiceStubCopilot, ReturnsName) {
  const Value v = EvalSource("=COPILOT(\"summarize\", 1, 2, 3)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(BuiltinsServiceStubCopilot, ArgumentErrorPropagates) {
  const Value v = EvalSource("=COPILOT(\"summarize\", 1/0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// PHONETIC - non-Ref passthrough surface
// ---------------------------------------------------------------------------
//
// PHONETIC is a lazy form (see `eval_phonetic_lazy`). The Ref path —
// where the impl reads `Cell::phonetic_text` off the un-evaluated
// Reference AST — is exercised in `phonetic_lazy_test.cpp` because it
// requires a workbook fixture with annotated cells. The cases below
// cover the non-Ref arm: literal text passes through unchanged, blanks
// surface as "", numeric / boolean values surface #N/A, and an error
// argument propagates verbatim.

TEST(BuiltinsServiceStubPhonetic, TextInputReturnsSameText) {
  const Value v = EvalSource("=PHONETIC(\"hello\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "hello");
}

TEST(BuiltinsServiceStubPhonetic, EmptyTextInputReturnsEmpty) {
  const Value v = EvalSource("=PHONETIC(\"\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "");
}

TEST(BuiltinsServiceStubPhonetic, NumberInputReturnsNA) {
  const Value v = EvalSource("=PHONETIC(42)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsServiceStubPhonetic, BoolInputReturnsNA) {
  const Value v = EvalSource("=PHONETIC(TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsServiceStubPhonetic, ArgumentErrorPropagates) {
  const Value v = EvalSource("=PHONETIC(1/0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// GETPIVOTDATA - the full lazy-form contract is exercised in
// `getpivotdata_lazy_test.cpp` where a workbook + pivot fixture lets
// the Ref-anchor lookup resolve a real pivot table. Only the
// fixtureless / no-workbook surface is pinned here: any GETPIVOTDATA
// call evaluated through `EvalSource` runs against a default
// `EvalContext` with no bound workbook, so anchor resolution always
// fails with `#REF!`. Errors in argument 0 propagate before the
// anchor check fires.
// ---------------------------------------------------------------------------

TEST(BuiltinsGetPivotDataNoWorkbook, AnchorWithoutWorkbookReturnsRef) {
  // No workbook bound -> find_pivot_at_anchor cannot run -> #REF!.
  // Arg 1 is a string here (not a Ref) so the lazy form falls through
  // its non-Ref path and surfaces #REF! after eager-evaluating the
  // expression for error propagation.
  const Value v = EvalSource("=GETPIVOTDATA(\"Sales\", \"PivotTable1\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsGetPivotDataNoWorkbook, FirstArgErrorPropagates) {
  // The lazy form evaluates arg 0 first; an error there propagates
  // before any structural check on arg 1.
  const Value v = EvalSource("=GETPIVOTDATA(1/0, \"PivotTable1\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsGetPivotDataNoWorkbook, SecondArgErrorPropagates) {
  // A non-Ref second argument is eagerly evaluated for error
  // propagation; the embedded #DIV/0! surfaces before the structural
  // #REF! reject fires.
  const Value v = EvalSource("=GETPIVOTDATA(\"Sales\", 1/0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// ISOMITTED - lazy form
// ---------------------------------------------------------------------------
//
// ISOMITTED returns TRUE only when its argument is a bare name reference
// resolving to an "omitted" trailing-optional LAMBDA parameter. Everything
// else — numeric, text, or boolean literals; arithmetic; an error-producing
// expression; a regular LET-bound name — yields FALSE. The lambda-aware
// behaviour is exercised in `tests/unit/eval/lambda_test.cpp`; the cases
// below pin down the outside-lambda surface.

TEST(BuiltinsServiceStubIsOmitted, NumberInputReturnsFalse) {
  const Value v = EvalSource("=ISOMITTED(42)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsServiceStubIsOmitted, TextInputReturnsFalse) {
  const Value v = EvalSource("=ISOMITTED(\"hello\")");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsServiceStubIsOmitted, BoolInputReturnsFalse) {
  const Value v = EvalSource("=ISOMITTED(TRUE)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsServiceStubIsOmitted, ErrorInputDoesNotPropagate) {
  // ISOMITTED is a lazy form: the error-producing arg is never evaluated
  // because the impl only inspects AST shape. =ISOMITTED(1/0) yields FALSE,
  // not #DIV/0!.
  const Value v = EvalSource("=ISOMITTED(1/0)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

}  // namespace
}  // namespace eval
}  // namespace formulon
