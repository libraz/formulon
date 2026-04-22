// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit + integration tests for cross-sheet reference evaluation. Exercises
// `Workbook::sheet_by_name`, the `Workbook`-aware `EvalContext` constructor,
// and the full tree-walker pipeline against formulas that cross sheet
// boundaries — including the cross-sheet cycle detection case.

#include <cstdint>
#include <string>

#include "gtest/gtest.h"
#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "parser/parser.h"
#include "parser/reference.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

// Builds a local A1 reference at (row, col). Defaults match the parser's
// output for a plain, unqualified, not-absolute A1 lexeme.
parser::Reference MakeLocalRef(std::uint32_t row, std::uint32_t col) {
  parser::Reference ref;
  ref.row = row;
  ref.col = col;
  return ref;
}

// Builds a qualified A1 reference whose `sheet` names `sheet_name`. The
// view is borrowed from the caller — pass a stable string literal or a
// named `std::string` whose lifetime outlasts the test body.
parser::Reference MakeSheetRef(std::string_view sheet_name,
                               std::uint32_t row, std::uint32_t col) {
  parser::Reference ref;
  ref.sheet = sheet_name;
  ref.row = row;
  ref.col = col;
  return ref;
}

// Minimal two-sheet workbook: `Workbook::create()` seeds `Sheet1`, and we
// append `Sheet2` for cross-sheet targets.
Workbook MakeTwoSheetWorkbook() {
  Workbook wb = Workbook::create();
  wb.add_sheet("Sheet2");
  return wb;
}

// ---------------------------------------------------------------------------
// Workbook::sheet_by_name
// ---------------------------------------------------------------------------

TEST(WorkbookSheetLookup, ExactMatch) {
  Workbook wb = MakeTwoSheetWorkbook();
  const Sheet* s = wb.sheet_by_name("Sheet2");
  ASSERT_NE(s, nullptr);
  EXPECT_EQ(s->name(), "Sheet2");
}

TEST(WorkbookSheetLookup, CaseInsensitiveMatch) {
  Workbook wb = MakeTwoSheetWorkbook();
  const Sheet* upper = wb.sheet_by_name("SHEET2");
  const Sheet* mixed = wb.sheet_by_name("sHeEt2");
  ASSERT_NE(upper, nullptr);
  ASSERT_NE(mixed, nullptr);
  EXPECT_EQ(upper, mixed);
  EXPECT_EQ(upper->name(), "Sheet2");
}

TEST(WorkbookSheetLookup, MissingReturnsNull) {
  Workbook wb = MakeTwoSheetWorkbook();
  EXPECT_EQ(wb.sheet_by_name("Sheet3"), nullptr);
  EXPECT_EQ(wb.sheet_by_name(""), nullptr);
}

TEST(WorkbookSheetLookup, AfterAddSheetIsFound) {
  Workbook wb = MakeTwoSheetWorkbook();
  EXPECT_EQ(wb.sheet_by_name("Data"), nullptr);
  wb.add_sheet("Data");
  ASSERT_NE(wb.sheet_by_name("Data"), nullptr);
  EXPECT_EQ(wb.sheet_by_name("DATA")->name(), "Data");
}

TEST(WorkbookSheetLookup, Default_create_HasSheet1) {
  Workbook wb = Workbook::create();
  const Sheet* s = wb.sheet_by_name("Sheet1");
  ASSERT_NE(s, nullptr);
  EXPECT_EQ(s->name(), "Sheet1");
}

TEST(WorkbookSheetLookup, AddSheetReturnsReference) {
  Workbook wb = Workbook::create();
  Sheet& added = wb.add_sheet("Extra");
  added.set_cell_value(0, 0, Value::number(7.0));
  // Same mutation observed via sheet_by_name.
  const Sheet* via_lookup = wb.sheet_by_name("Extra");
  ASSERT_NE(via_lookup, nullptr);
  const Cell* c = via_lookup->cell_at(0, 0);
  ASSERT_NE(c, nullptr);
  ASSERT_TRUE(c->cached_value.is_number());
  EXPECT_EQ(c->cached_value.as_number(), 7.0);
}

// ---------------------------------------------------------------------------
// EvalContext cross-sheet single-reference resolution
// ---------------------------------------------------------------------------

TEST(EvalContextCrossSheet, QualifiedRefToOtherSheetResolves) {
  Workbook wb = MakeTwoSheetWorkbook();
  wb.sheet(1).set_cell_value(0, 0, Value::number(42.0));
  EvalState state;
  Arena arena;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = ctx.resolve_ref(MakeSheetRef("Sheet2", 0, 0), arena, default_registry());
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 42.0);
}

TEST(EvalContextCrossSheet, QualifiedRefCaseInsensitive) {
  Workbook wb = MakeTwoSheetWorkbook();
  wb.sheet(1).set_cell_value(2, 3, Value::number(100.0));
  EvalState state;
  Arena arena;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = ctx.resolve_ref(MakeSheetRef("SHEET2", 2, 3), arena, default_registry());
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 100.0);
}

TEST(EvalContextCrossSheet, QualifiedRefMissingSheetReturnsRef) {
  Workbook wb = MakeTwoSheetWorkbook();
  EvalState state;
  Arena arena;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = ctx.resolve_ref(MakeSheetRef("NoSuchSheet", 0, 0), arena, default_registry());
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(EvalContextCrossSheet, QualifiedRefWithoutWorkbookReturnsRef) {
  // 2-arg (Sheet-only) constructor leaves `workbook_ == nullptr`, so a
  // qualified reference cannot be resolved and surfaces as #REF!.
  Sheet sheet("Sheet1");
  EvalState state;
  Arena arena;
  const EvalContext ctx(sheet, state);
  const Value v = ctx.resolve_ref(MakeSheetRef("Sheet2", 0, 0), arena, default_registry());
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(EvalContextCrossSheet, SelfQualifiedRefMatchesLocal) {
  // `=Sheet1!A1` from Sheet1 is equivalent to `=A1` — the resolved target
  // is the current sheet.
  Workbook wb = MakeTwoSheetWorkbook();
  wb.sheet(0).set_cell_value(0, 0, Value::number(9.0));
  EvalState state;
  Arena arena;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = ctx.resolve_ref(MakeSheetRef("Sheet1", 0, 0), arena, default_registry());
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 9.0);
}

TEST(EvalContextCrossSheet, QuotedSheetNameResolves) {
  // sheet_quoted is a round-trip format hint only; lookup uses the bare
  // unquoted name either way.
  Workbook wb = Workbook::create();
  wb.add_sheet("Sheet 1");  // name with a space — parser would quote on output.
  wb.sheet(1).set_cell_value(0, 0, Value::number(11.0));
  EvalState state;
  Arena arena;
  const EvalContext ctx(wb, wb.sheet(0), state);
  parser::Reference ref = MakeSheetRef("Sheet 1", 0, 0);
  ref.sheet_quoted = true;
  const Value v = ctx.resolve_ref(ref, arena, default_registry());
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 11.0);
}

TEST(EvalContextCrossSheet, CrossSheetFormulaCellEvaluates) {
  Workbook wb = MakeTwoSheetWorkbook();
  wb.sheet(1).set_cell_formula(0, 0, "=1+2");
  EvalState state;
  Arena arena;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = ctx.resolve_ref(MakeSheetRef("Sheet2", 0, 0), arena, default_registry());
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 3.0);
}

TEST(EvalContextCrossSheet, CrossSheetCycleDirectReturnsRef) {
  // Sheet1!A1 = =Sheet1!A1 — direct self-cycle via qualifier. The
  // `(target_sheet, row, col)` key catches it.
  Workbook wb = MakeTwoSheetWorkbook();
  wb.sheet(0).set_cell_formula(0, 0, "=Sheet1!A1");
  EvalState state;
  Arena arena;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = ctx.resolve_ref(MakeLocalRef(0, 0), arena, default_registry());
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(EvalContextCrossSheet, CrossSheetCycleIndirectReturnsRef) {
  // Sheet1!A1 = =Sheet2!A1; Sheet2!A1 = =Sheet1!A1 — indirect cross-sheet
  // cycle.
  Workbook wb = MakeTwoSheetWorkbook();
  wb.sheet(0).set_cell_formula(0, 0, "=Sheet2!A1");
  wb.sheet(1).set_cell_formula(0, 0, "=Sheet1!A1");
  EvalState state;
  Arena arena;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = ctx.resolve_ref(MakeLocalRef(0, 0), arena, default_registry());
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(EvalContextCrossSheet, ChainAcrossSheets) {
  // Sheet1!A1 = 5 (literal)
  // Sheet2!A1 = =Sheet1!A1 + 1 (= 6)
  // Sheet1!B1 = =Sheet2!A1 * 2 (= 12)
  Workbook wb = MakeTwoSheetWorkbook();
  wb.sheet(0).set_cell_value(0, 0, Value::number(5.0));
  wb.sheet(1).set_cell_formula(0, 0, "=Sheet1!A1+1");
  wb.sheet(0).set_cell_formula(0, 1, "=Sheet2!A1*2");
  EvalState state;
  Arena arena;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = ctx.resolve_ref(MakeLocalRef(0, 1), arena, default_registry());
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 12.0);
}

// ---------------------------------------------------------------------------
// EvalContext cross-sheet range expansion
// ---------------------------------------------------------------------------

TEST(EvalContextCrossSheetRange, QualifiedRangeOnOtherSheet) {
  // Sheet2!A1:A3 expands across Sheet2; LHS carries the qualifier, RHS
  // inherits (as the parser emits).
  Workbook wb = MakeTwoSheetWorkbook();
  wb.sheet(1).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(1).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(1).set_cell_value(2, 0, Value::number(3.0));
  EvalState state;
  Arena arena;
  const EvalContext ctx(wb, wb.sheet(0), state);
  parser::Reference lhs = MakeSheetRef("Sheet2", 0, 0);
  parser::Reference rhs = MakeLocalRef(2, 0);  // RHS has no sheet qualifier.
  auto result = ctx.expand_range(lhs, rhs, arena, default_registry());
  ASSERT_TRUE(result);
  ASSERT_EQ(result.value().size(), 3u);
  EXPECT_EQ(result.value()[0].as_number(), 1.0);
  EXPECT_EQ(result.value()[1].as_number(), 2.0);
  EXPECT_EQ(result.value()[2].as_number(), 3.0);
}

TEST(EvalContextCrossSheetRange, BothEndpointsQualifiedMatching) {
  // Both endpoints naming the same sheet (case-insensitively) is legal.
  Workbook wb = MakeTwoSheetWorkbook();
  wb.sheet(1).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(1).set_cell_value(1, 0, Value::number(20.0));
  EvalState state;
  Arena arena;
  const EvalContext ctx(wb, wb.sheet(0), state);
  parser::Reference lhs = MakeSheetRef("Sheet2", 0, 0);
  parser::Reference rhs = MakeSheetRef("SHEET2", 1, 0);
  auto result = ctx.expand_range(lhs, rhs, arena, default_registry());
  ASSERT_TRUE(result);
  ASSERT_EQ(result.value().size(), 2u);
  EXPECT_EQ(result.value()[0].as_number(), 10.0);
  EXPECT_EQ(result.value()[1].as_number(), 20.0);
}

TEST(EvalContextCrossSheetRange, MismatchedEndpointsReturnRef) {
  Workbook wb = MakeTwoSheetWorkbook();
  EvalState state;
  Arena arena;
  const EvalContext ctx(wb, wb.sheet(0), state);
  parser::Reference lhs = MakeSheetRef("Sheet1", 0, 0);
  parser::Reference rhs = MakeSheetRef("Sheet2", 1, 1);
  auto result = ctx.expand_range(lhs, rhs, arena, default_registry());
  ASSERT_FALSE(result);
  EXPECT_EQ(result.error(), ErrorCode::Ref);
}

TEST(EvalContextCrossSheetRange, LocalRangeUnaffected) {
  // Both endpoints unqualified → current sheet.
  Workbook wb = MakeTwoSheetWorkbook();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  EvalState state;
  Arena arena;
  const EvalContext ctx(wb, wb.sheet(0), state);
  auto result = ctx.expand_range(MakeLocalRef(0, 0), MakeLocalRef(1, 0), arena,
                                 default_registry());
  ASSERT_TRUE(result);
  ASSERT_EQ(result.value().size(), 2u);
  EXPECT_EQ(result.value()[0].as_number(), 1.0);
  EXPECT_EQ(result.value()[1].as_number(), 2.0);
}

TEST(EvalContextCrossSheetRange, QualifiedRangeMissingSheetReturnsRef) {
  Workbook wb = MakeTwoSheetWorkbook();
  EvalState state;
  Arena arena;
  const EvalContext ctx(wb, wb.sheet(0), state);
  parser::Reference lhs = MakeSheetRef("NoSuchSheet", 0, 0);
  parser::Reference rhs = MakeLocalRef(1, 0);
  auto result = ctx.expand_range(lhs, rhs, arena, default_registry());
  ASSERT_FALSE(result);
  EXPECT_EQ(result.error(), ErrorCode::Ref);
}

TEST(EvalContextCrossSheetRange, QualifiedRangeWithoutWorkbookReturnsRef) {
  Sheet sheet("Sheet1");
  EvalState state;
  Arena arena;
  const EvalContext ctx(sheet, state);
  parser::Reference lhs = MakeSheetRef("Sheet2", 0, 0);
  parser::Reference rhs = MakeLocalRef(1, 0);
  auto result = ctx.expand_range(lhs, rhs, arena, default_registry());
  ASSERT_FALSE(result);
  EXPECT_EQ(result.error(), ErrorCode::Ref);
}

TEST(EvalContextCrossSheetRange, CrossSheetSumWorks) {
  // Full pipeline: parse `=SUM(Sheet2!A1:A3)` and evaluate it against a
  // two-sheet workbook. Sheet2 column A carries 1/2/3 → SUM == 6.
  Workbook wb = MakeTwoSheetWorkbook();
  wb.sheet(1).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(1).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(1).set_cell_value(2, 0, Value::number(3.0));

  EvalState state;
  Arena arena;
  parser::Parser p("=SUM(Sheet2!A1:A3)", arena);
  parser::AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);

  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = evaluate(*root, arena, default_registry(), ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 6.0);
}

TEST(EvalContextCrossSheetRange, CrossSheetCycleViaRangeReturnsRef) {
  // Sheet1!A1 = =SUM(Sheet2!A1); Sheet2!A1 = =Sheet1!A1. Resolving
  // Sheet1!A1 must surface #REF! via the cross-sheet cycle detection in
  // `EvalState`. SUM's error-propagation (propagate_errors = true) will
  // lift the per-cell #REF! out of the range expansion.
  Workbook wb = MakeTwoSheetWorkbook();
  wb.sheet(0).set_cell_formula(0, 0, "=SUM(Sheet2!A1)");
  wb.sheet(1).set_cell_formula(0, 0, "=Sheet1!A1");
  EvalState state;
  Arena arena;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = ctx.resolve_ref(MakeLocalRef(0, 0), arena, default_registry());
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
