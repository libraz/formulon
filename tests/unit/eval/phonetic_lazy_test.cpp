//
// Unit tests for the PHONETIC lazy form. The Ref path requires a
// workbook fixture with `Cell::phonetic_runs` populated, so it cannot
// be expressed via the formula-only `EvalSource` helper used by the
// other service-stub tests; this file builds the workbook directly via
// the storage-layer API and exercises both the annotated-cell and
// fallback paths.
//
// Mac Excel 365 (ja-JP) is the primary oracle. PHONETIC's observable
// behaviour:
//
//   * `=PHONETIC(A1)` with `<rPh>` annotations -> each annotated span
//     replaced by its kana, everything outside every span passed
//     through unchanged.
//   * `=PHONETIC(A1)` text cell, no annotation -> surface text.
//   * `=PHONETIC(A1)` blank cell -> "".
//   * `=PHONETIC(A1)` number / boolean cell -> #N/A.
//   * Whole-row / whole-column ref -> #VALUE!.
//   * Non-Ref args (literal text, arithmetic, ...) ride the eager arm
//     in `eval_phonetic_lazy` and surface the same passthrough rule.

#include <cstdint>
#include <string>
#include <string_view>

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "phonetic.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

// Parses `src` and evaluates against `ctx`.
Value EvalWith(std::string_view src, const EvalContext& ctx) {
  static thread_local Arena arena;
  arena.reset();
  parser::Parser p(src, arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return evaluate(*root, arena, default_registry(), ctx);
}

// Builds a single-sheet workbook anchored at "Sheet1".
Workbook MakeSingleSheetWorkbook() {
  return Workbook::create();
}

// ---------------------------------------------------------------------------
// Ref path: annotated text cell -> kana
// ---------------------------------------------------------------------------

TEST(BuiltinsPhoneticLazy, AnnotatedTextCellReturnsKana) {
  Workbook wb = MakeSingleSheetWorkbook();
  // 山田 with kana やまだ. `set_cell_phonetic` records the reading as a
  // run covering the whole surface text, so nothing is left over.
  wb.sheet(0).set_cell_value(0, 0, Value::text("\xE5\xB1\xB1\xE7\x94\xB0"));
  wb.sheet(0).set_cell_phonetic(0, 0, "\xE3\x82\x84\xE3\x81\xBE\xE3\x81\xA0");

  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=PHONETIC(A1)", ctx);
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "\xE3\x82\x84\xE3\x81\xBE\xE3\x81\xA0");
}

TEST(BuiltinsPhoneticLazy, AnnotatedAsciiCellReturnsKana) {
  // Ensures the impl reads the annotation regardless of the surface
  // text encoding; an ASCII surface with an ASCII annotation still
  // routes through the same path.
  Workbook wb = MakeSingleSheetWorkbook();
  wb.sheet(0).set_cell_value(0, 0, Value::text("kanji"));
  wb.sheet(0).set_cell_phonetic(0, 0, "furigana");

  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=PHONETIC(A1)", ctx);
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "furigana");
}

// ---------------------------------------------------------------------------
// Ref path: partial annotation keeps the unannotated remainder
// ---------------------------------------------------------------------------

TEST(BuiltinsPhoneticLazy, PartialAnnotationKeepsUnannotatedRemainder) {
  // 東京都 annotated over 東京 alone. Excel substitutes the annotated
  // span and passes 都 through, so the result mixes kana and kanji:
  // トウキョウ都. Observed on Mac Excel 365 ja-JP 16.112.1 against a
  // workbook carrying <rPh sb="0" eb="2">.
  Workbook wb = MakeSingleSheetWorkbook();
  wb.sheet(0).set_cell_value(0, 0, Value::text("\xE6\x9D\xB1\xE4\xBA\xAC\xE9\x83\xBD"));
  wb.sheet(0).set_cell_phonetic_runs(
      0, 0, {PhoneticRun{0U, 2U, "\xE3\x83\x88\xE3\x82\xA6\xE3\x82\xAD\xE3\x83\xA7\xE3\x82\xA6"}});

  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=PHONETIC(A1)", ctx);
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "\xE3\x83\x88\xE3\x82\xA6\xE3\x82\xAD\xE3\x83\xA7\xE3\x82\xA6\xE9\x83\xBD");
}

TEST(BuiltinsPhoneticLazy, AnnotationInTheMiddleKeepsBothEnds) {
  // A span that starts past the head exercises the pass-through arm on
  // both sides of the substitution.
  Workbook wb = MakeSingleSheetWorkbook();
  wb.sheet(0).set_cell_value(0, 0,
                             Value::text("ab\xE5\xB1\xB1\xE7\x94\xB0"
                                         "cd"));
  wb.sheet(0).set_cell_phonetic_runs(0, 0, {PhoneticRun{2U, 4U, "YAMADA"}});

  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=PHONETIC(A1)", ctx);
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "abYAMADAcd");
}

TEST(BuiltinsPhoneticLazy, MultipleSpansSubstituteIndependently) {
  // 山田太郎 with one block per surname / given name, the shape Excel
  // emits for an IME-typed name. Both spans are covered, so the result
  // is pure kana even though it arrives as two runs.
  Workbook wb = MakeSingleSheetWorkbook();
  wb.sheet(0).set_cell_value(0, 0, Value::text("\xE5\xB1\xB1\xE7\x94\xB0\xE5\xA4\xAA\xE9\x83\x8E"));
  wb.sheet(0).set_cell_phonetic_runs(0, 0,
                                     {PhoneticRun{0U, 2U, "\xE3\x82\x84\xE3\x81\xBE\xE3\x81\xA0"},
                                      PhoneticRun{2U, 4U, "\xE3\x81\x9F\xE3\x82\x8D\xE3\x81\x86"}});

  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=PHONETIC(A1)", ctx);
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "\xE3\x82\x84\xE3\x81\xBE\xE3\x81\xA0\xE3\x81\x9F\xE3\x82\x8D\xE3\x81\x86");
}

TEST(BuiltinsPhoneticLazy, SpanPastEndOfTextConsumesWhatIsThere) {
  // An `eb` beyond the surface text is malformed input Excel would not
  // write. It must not read past the string or drop the kana.
  Workbook wb = MakeSingleSheetWorkbook();
  wb.sheet(0).set_cell_value(0, 0, Value::text("ab"));
  wb.sheet(0).set_cell_phonetic_runs(0, 0, {PhoneticRun{0U, 99U, "kana"}});

  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=PHONETIC(A1)", ctx);
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "kana");
}

TEST(BuiltinsPhoneticLazy, SurrogatePairCountsAsTwoUnits) {
  // Span offsets are UTF-16 code units, so a supplementary-plane
  // codepoint occupies two of them. "\U0001F600ab" annotated over
  // [0, 2) covers the emoji alone and leaves "ab" in place.
  Workbook wb = MakeSingleSheetWorkbook();
  wb.sheet(0).set_cell_value(0, 0,
                             Value::text("\xF0\x9F\x98\x80"
                                         "ab"));
  wb.sheet(0).set_cell_phonetic_runs(0, 0, {PhoneticRun{0U, 2U, "smile"}});

  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=PHONETIC(A1)", ctx);
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "smileab");
}

// ---------------------------------------------------------------------------
// Ref path: fallback when no annotation is present
// ---------------------------------------------------------------------------

TEST(BuiltinsPhoneticLazy, UnannotatedTextCellReturnsSurfaceText) {
  Workbook wb = MakeSingleSheetWorkbook();
  wb.sheet(0).set_cell_value(0, 0, Value::text("plain"));
  // No `set_cell_phonetic` call: the run list stays empty.

  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=PHONETIC(A1)", ctx);
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "plain");
}

TEST(BuiltinsPhoneticLazy, BlankCellReturnsEmpty) {
  // A blank cell with no annotation produces an empty string, mirroring
  // Mac Excel's behaviour for `=PHONETIC(A1)` on an untouched cell.
  Workbook wb = MakeSingleSheetWorkbook();

  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=PHONETIC(A1)", ctx);
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "");
}

TEST(BuiltinsPhoneticLazy, NumberCellReturnsNA) {
  // A cell holding a number with no <rPh> annotation surfaces #N/A.
  // (PHONETIC's strict-text rule applies on the value-passthrough arm.)
  Workbook wb = MakeSingleSheetWorkbook();
  wb.sheet(0).set_cell_value(0, 0, Value::number(42.0));

  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=PHONETIC(A1)", ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsPhoneticLazy, BooleanCellReturnsNA) {
  Workbook wb = MakeSingleSheetWorkbook();
  wb.sheet(0).set_cell_value(0, 0, Value::boolean(true));

  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=PHONETIC(A1)", ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsPhoneticLazy, ErrorCellPropagates) {
  // A cell holding a stored error propagates that error rather than
  // surfacing #N/A.
  Workbook wb = MakeSingleSheetWorkbook();
  wb.sheet(0).set_cell_value(0, 0, Value::error(ErrorCode::Div0));

  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=PHONETIC(A1)", ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// Edge case: annotation is present even on a non-text cell (e.g. a
// number cell that was previously text and retained the <rPh> on
// round-trip). The lazy impl reads the run list first, before
// inspecting the value, so the annotation wins.
TEST(BuiltinsPhoneticLazy, AnnotatedNumberCellPrefersKana) {
  Workbook wb = MakeSingleSheetWorkbook();
  wb.sheet(0).set_cell_value(0, 0, Value::number(42.0));
  wb.sheet(0).set_cell_phonetic(0, 0, "kana");

  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=PHONETIC(A1)", ctx);
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "kana");
}

// ---------------------------------------------------------------------------
// Ref shape: full row / column / range
// ---------------------------------------------------------------------------

TEST(BuiltinsPhoneticLazy, FullColumnRefReturnsValue) {
  Workbook wb = MakeSingleSheetWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=PHONETIC(A:A)", ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsPhoneticLazy, FullRowRefReturnsValue) {
  Workbook wb = MakeSingleSheetWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=PHONETIC(1:1)", ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsPhoneticLazy, OneByOneRangeRefFallsThroughToScalar) {
  // `A1:A1` parses as a RangeOp, not a single Ref. The lazy impl falls
  // through to the eager-evaluation arm; the 1x1 range collapses to
  // the cell's scalar text under implicit intersection. The kana
  // annotation is NOT consulted because the impl never sees the
  // un-evaluated Reference shape; only the surface text is returned.
  // This matches Mac Excel's surface for `=PHONETIC(A1:A1)` where the
  // annotation is dropped by the range-evaluation arm.
  Workbook wb = MakeSingleSheetWorkbook();
  wb.sheet(0).set_cell_value(0, 0, Value::text("kanji"));
  wb.sheet(0).set_cell_phonetic(0, 0, "kana");

  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=PHONETIC(A1:A1)", ctx);
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "kanji");
}

TEST(BuiltinsPhoneticLazy, MultiCellRangeRefReturnsNA) {
  // A multi-cell range that does NOT collapse to a scalar surfaces
  // through the eager arm as an Array (or, today, the first-element
  // value depending on intersection rules). For numeric data, the
  // strict-text rule rejects with #N/A.
  Workbook wb = MakeSingleSheetWorkbook();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));

  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=PHONETIC(A1:A2)", ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

// ---------------------------------------------------------------------------
// Non-Ref arguments (literal / arithmetic / function call)
// ---------------------------------------------------------------------------

TEST(BuiltinsPhoneticLazy, LiteralTextPassesThrough) {
  Workbook wb = MakeSingleSheetWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=PHONETIC(\"literal\")", ctx);
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "literal");
}

TEST(BuiltinsPhoneticLazy, LiteralNumberReturnsNA) {
  Workbook wb = MakeSingleSheetWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=PHONETIC(42)", ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsPhoneticLazy, ArgumentErrorPropagates) {
  // The non-Ref arm eagerly evaluates the subtree; a #DIV/0! result
  // propagates instead of #N/A.
  Workbook wb = MakeSingleSheetWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=PHONETIC(1/0)", ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// Arity guards
// ---------------------------------------------------------------------------

TEST(BuiltinsPhoneticLazy, ZeroArityReturnsValue) {
  Workbook wb = MakeSingleSheetWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=PHONETIC()", ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsPhoneticLazy, TwoArityReturnsValue) {
  Workbook wb = MakeSingleSheetWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=PHONETIC(\"a\", \"b\")", ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
