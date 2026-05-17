// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Self-tests for the `tests/util/` infrastructure library. Exercises
// each helper at least once so that breakage in the shared utility code
// surfaces under `ctest` before propagating to the 200+ test files that
// will migrate onto it in subsequent waves.

#include <gtest/gtest.h>

#include <cstdint>
#include <string_view>

#include "cell.h"
#include "sheet.h"
#include "util/test_arena.h"
#include "util/test_eval_helpers.h"
#include "util/test_value_macros.h"
#include "util/test_workbook_builder.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace test {
namespace {

// ---------------------------------------------------------------------------
// test_arena
// ---------------------------------------------------------------------------

TEST(TestArenaUtil, TestArenaHolderConstructsEmpty) {
  TestArena holder;
  // A freshly built arena has not received any allocation requests
  // yet, so its bytes_used must be zero. The internal chunk count may
  // also still be zero (the Arena allocates lazily on first request).
  EXPECT_EQ(holder.get().bytes_used(), 0U);
}

TEST(TestArenaUtil, ThreadLocalArenasResetOnEachCall) {
  Arena& a = test_parse_arena();
  EXPECT_EQ(a.bytes_used(), 0U);
  void* p = a.allocate(64, alignof(std::max_align_t));
  ASSERT_NE(p, nullptr);
  EXPECT_GE(a.bytes_used(), 64U);

  // Re-acquiring resets the arena (so the test sees a clean slate)
  // but the chunk may be retained for reuse — that is the entire
  // point of the thread-local form.
  Arena& a2 = test_parse_arena();
  EXPECT_EQ(&a2, &a);  // same thread-local instance
  EXPECT_EQ(a2.bytes_used(), 0U);
}

TEST(TestArenaUtil, ParseAndEvalArenasAreDistinct) {
  Arena& parse_arena = test_parse_arena();
  Arena& eval_arena = test_eval_arena();
  EXPECT_NE(&parse_arena, &eval_arena);
}

TEST(TestArenaUtil, TestArenaAliasMatchesEvalArena) {
  Arena& a = test_arena();
  Arena& b = test_eval_arena();
  EXPECT_EQ(&a, &b);
}

// ---------------------------------------------------------------------------
// test_eval_helpers
// ---------------------------------------------------------------------------

TEST(TestEvalHelpersUtil, EvalSourceLiteralNumber) {
  const Value v = EvalSource("=1+2");
  EXPECT_VALUE_NUMBER(v, 3.0);
}

TEST(TestEvalHelpersUtil, EvalSourceLiteralText) {
  const Value v = EvalSource("=\"hello\"");
  EXPECT_VALUE_TEXT(v, "hello");
}

TEST(TestEvalHelpersUtil, EvalSourceDivByZeroSurfacesAsError) {
  const Value v = EvalSource("=1/0");
  EXPECT_VALUE_ERROR(v, ErrorCode::Div0);
}

TEST(TestEvalHelpersUtil, EvalSourceInResolvesLocalReference) {
  Workbook wb = WorkbookBuilder().sheet("Sheet1").cell("A1", 42.0).build();
  const Value v = EvalSourceIn("=A1+8", wb, wb.sheet(0));
  EXPECT_VALUE_NUMBER(v, 50.0);
}

TEST(TestEvalHelpersUtil, EvalSourceAtAnchorsFormulaCell) {
  // CELL("address") requires an anchored formula cell. EvalSourceAt
  // pins the anchor; the result is the absolute A1 address of the
  // anchor cell.
  Workbook wb = WorkbookBuilder().sheet("Sheet1").build();
  const Value v = EvalSourceAt("=CELL(\"address\")", wb, wb.sheet(0), /*row=*/2U, /*col=*/4U);
  // CELL("address") returns an absolute reference like "$E$3".
  ASSERT_TRUE(v.is_text()) << "expected Text, got " << v.debug_to_string();
  EXPECT_EQ(v.as_text(), "$E$3");
}

TEST(TestEvalHelpersUtil, EvalSourceAtByNameResolvesSheet) {
  Workbook wb = WorkbookBuilder().sheet("Data").cell("A1", 7.0).build();
  const Value v = EvalSourceAt("=A1*6", wb, std::string_view("Data"), 0U, 0U);
  EXPECT_VALUE_NUMBER(v, 42.0);
}

TEST(TestEvalHelpersUtil, EvalSourceAtByNameMissingSheetReturnsRef) {
  Workbook wb = WorkbookBuilder().sheet("Sheet1").build();
  const Value v = EvalSourceAt("=A1", wb, std::string_view("DoesNotExist"), 0U, 0U);
  EXPECT_VALUE_ERROR(v, ErrorCode::Ref);
}

// ---------------------------------------------------------------------------
// test_value_macros
// ---------------------------------------------------------------------------

TEST(TestValueMacrosUtil, NumberMacro) {
  EXPECT_VALUE_NUMBER(Value::number(3.14), 3.14);
}

TEST(TestValueMacrosUtil, NumberNearMacro) {
  EXPECT_VALUE_NUMBER_NEAR(Value::number(1.0 / 3.0), 0.333333, 1e-5);
}

TEST(TestValueMacrosUtil, TextMacro) {
  EXPECT_VALUE_TEXT(Value::text("abc"), "abc");
}

TEST(TestValueMacrosUtil, BoolMacro) {
  EXPECT_VALUE_BOOL(Value::boolean(true), true);
  EXPECT_VALUE_BOOL(Value::boolean(false), false);
}

TEST(TestValueMacrosUtil, ErrorMacro) {
  EXPECT_VALUE_ERROR(Value::error(ErrorCode::NA), ErrorCode::NA);
}

TEST(TestValueMacrosUtil, BlankMacro) {
  EXPECT_VALUE_BLANK(Value::blank());
}

// ---------------------------------------------------------------------------
// test_workbook_builder
// ---------------------------------------------------------------------------

TEST(TestWorkbookBuilderUtil, EmptyBuildHasNoSheets) {
  Workbook wb = WorkbookBuilder().build();
  EXPECT_EQ(wb.sheet_count(), 0U);
}

TEST(TestWorkbookBuilderUtil, SingleSheetIsAppended) {
  Workbook wb = WorkbookBuilder().sheet("MySheet").build();
  ASSERT_EQ(wb.sheet_count(), 1U);
  EXPECT_EQ(wb.sheet(0).name(), "MySheet");
}

TEST(TestWorkbookBuilderUtil, NumberCellStoresLiteral) {
  Workbook wb = WorkbookBuilder().sheet("S").cell("A1", 12.5).build();
  ASSERT_EQ(wb.sheet_count(), 1U);
  const Cell* cell = wb.sheet(0).cell_at(0, 0);
  ASSERT_NE(cell, nullptr);
  EXPECT_VALUE_NUMBER(cell->cached_value, 12.5);
}

TEST(TestWorkbookBuilderUtil, BoolCellStoresLiteral) {
  Workbook wb = WorkbookBuilder().sheet("S").cell("B2", true).build();
  const Cell* cell = wb.sheet(0).cell_at(1, 1);
  ASSERT_NE(cell, nullptr);
  EXPECT_VALUE_BOOL(cell->cached_value, true);
}

TEST(TestWorkbookBuilderUtil, TextCellStoresText) {
  Workbook wb = WorkbookBuilder().sheet("S").cell("C3", std::string_view("hello")).build();
  const Cell* cell = wb.sheet(0).cell_at(2, 2);
  ASSERT_NE(cell, nullptr);
  EXPECT_VALUE_TEXT(cell->cached_value, "hello");
}

TEST(TestWorkbookBuilderUtil, EqualsPrefixedStringStoresAsFormula) {
  Workbook wb = WorkbookBuilder()  //
                    .sheet("S")
                    .cell("A1", 4.0)
                    .cell("B1", std::string_view("=A1+1"))
                    .build();
  // The formula is stored unevaluated; cached value is blank until
  // recalc() runs. We only assert the formula_text round-tripped.
  const Cell* cell = wb.sheet(0).cell_at(0, 1);
  ASSERT_NE(cell, nullptr);
  EXPECT_EQ(cell->formula_text, "=A1+1");
}

TEST(TestWorkbookBuilderUtil, ExplicitTextCellNeverInterpretsAsFormula) {
  // text_cell() bypasses the leading-`=` heuristic.
  Workbook wb = WorkbookBuilder().sheet("S").text_cell("A1", "=NOT_A_FORMULA").build();
  const Cell* cell = wb.sheet(0).cell_at(0, 0);
  ASSERT_NE(cell, nullptr);
  EXPECT_VALUE_TEXT(cell->cached_value, "=NOT_A_FORMULA");
}

TEST(TestWorkbookBuilderUtil, FormulaCellWithoutEqualsIsAccepted) {
  // Mirrors set_cell_formula's contract: the leading `=` is the
  // parser's contract, not the storage layer's.
  Workbook wb = WorkbookBuilder().sheet("S").formula_cell("A1", "1+2").build();
  const Cell* cell = wb.sheet(0).cell_at(0, 0);
  ASSERT_NE(cell, nullptr);
  EXPECT_EQ(cell->formula_text, "1+2");
}

TEST(TestWorkbookBuilderUtil, ValueOverloadStoresArbitraryKinds) {
  Workbook wb = WorkbookBuilder()  //
                    .sheet("S")
                    .cell("A1", Value::error(ErrorCode::NA))
                    .build();
  const Cell* cell = wb.sheet(0).cell_at(0, 0);
  ASSERT_NE(cell, nullptr);
  EXPECT_VALUE_ERROR(cell->cached_value, ErrorCode::NA);
}

TEST(TestWorkbookBuilderUtil, SheetCallSelectsExistingSheetCaseInsensitive) {
  Workbook wb = WorkbookBuilder()      //
                    .sheet("Data")     //
                    .cell("A1", 1.0)   //
                    .sheet("Other")    //
                    .cell("A1", 2.0)   //
                    .sheet("DATA")     // case-insensitive re-select
                    .cell("A2", 99.0)  //
                    .build();
  ASSERT_EQ(wb.sheet_count(), 2U);
  const Cell* reselected = wb.sheet(0).cell_at(1, 0);
  ASSERT_NE(reselected, nullptr);
  EXPECT_VALUE_NUMBER(reselected->cached_value, 99.0);
}

TEST(TestWorkbookBuilderUtil, EvalSourceInRoundtripsThroughBuilder) {
  // End-to-end: the builder + EvalSourceIn pair should produce the
  // same result as the open-coded `Workbook::create()` /
  // `set_cell_value` / `EvalContext` boilerplate that 100+ existing
  // test files repeat.
  Workbook wb = WorkbookBuilder()      //
                    .sheet("Sheet1")   //
                    .cell("A1", 10.0)  //
                    .cell("A2", 20.0)  //
                    .cell("A3", 30.0)  //
                    .build();
  const Value v = EvalSourceIn("=SUM(A1:A3)", wb, wb.sheet(0));
  EXPECT_VALUE_NUMBER(v, 60.0);
}

}  // namespace
}  // namespace test
}  // namespace formulon
