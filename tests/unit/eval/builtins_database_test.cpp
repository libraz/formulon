// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for the database-aggregation family: DSUM, DCOUNT,
// DCOUNTA, DAVERAGE, DMAX, DMIN, DPRODUCT, DSTDEV, DSTDEVP, DVAR, DVARP,
// DGET. These route through the lazy dispatch table in `tree_walker.cpp`
// because both the `database` and `criteria` arguments must be seen as
// two-dimensional RangeOp / Ref shapes (header row distinguishable from
// data rows) and because field lookup by header name requires inspecting
// the header row's string values.

#include <cmath>
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
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

// Evaluates `src` against `wb`/`current`. Each call resets the shared
// arenas so tests do not contaminate each other.
Value EvalIn(std::string_view src, const Workbook& wb, const Sheet& current) {
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
  EvalState state;
  const EvalContext ctx(wb, current, state);
  return evaluate(*root, eval_arena, default_registry(), ctx);
}

// Builds a canonical 4-row × 3-column fruit-sales database in A1:C4 on
// sheet 0. Layout:
//   A1=Fruit     B1=Region    C1=Sales
//   A2=Apple     B2=East      C2=100
//   A3=Apple     B3=West      C3=200
//   A4=Pear      B4=East      C4=50
//   A5=Banana    B5=West      C5=300
Workbook MakeFruitWorkbook() {
  Workbook wb = Workbook::create();
  auto& s = wb.sheet(0);
  s.set_cell_value(0, 0, Value::text("Fruit"));
  s.set_cell_value(0, 1, Value::text("Region"));
  s.set_cell_value(0, 2, Value::text("Sales"));
  s.set_cell_value(1, 0, Value::text("Apple"));
  s.set_cell_value(1, 1, Value::text("East"));
  s.set_cell_value(1, 2, Value::number(100.0));
  s.set_cell_value(2, 0, Value::text("Apple"));
  s.set_cell_value(2, 1, Value::text("West"));
  s.set_cell_value(2, 2, Value::number(200.0));
  s.set_cell_value(3, 0, Value::text("Pear"));
  s.set_cell_value(3, 1, Value::text("East"));
  s.set_cell_value(3, 2, Value::number(50.0));
  s.set_cell_value(4, 0, Value::text("Banana"));
  s.set_cell_value(4, 1, Value::text("West"));
  s.set_cell_value(4, 2, Value::number(300.0));
  return wb;
}

// Writes a single-criterion block at E1:E2 matching `Fruit == needle`.
// `needle` must point at storage whose lifetime exceeds the workbook —
// `Value::text` does not copy, so callers should pass a string literal
// or a string held in test-scope storage.
void SetSingleFruitCriterion(Workbook& wb, const char* needle) {
  auto& s = wb.sheet(0);
  s.set_cell_value(0, 4, Value::text("Fruit"));
  s.set_cell_value(1, 4, Value::text(needle));
}

// ---------------------------------------------------------------------------
// DSUM
// ---------------------------------------------------------------------------

TEST(BuiltinsDatabase, DSumSimpleHeaderMatch) {
  Workbook wb = MakeFruitWorkbook();
  SetSingleFruitCriterion(wb, "Apple");
  const Value v = EvalIn("=DSUM(A1:C5, \"Sales\", E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 300.0);
}

TEST(BuiltinsDatabase, DSumNumericFieldSelector) {
  Workbook wb = MakeFruitWorkbook();
  SetSingleFruitCriterion(wb, "Apple");
  // Field 3 == "Sales"
  const Value v = EvalIn("=DSUM(A1:C5, 3, E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 300.0);
}

TEST(BuiltinsDatabase, DSumFieldHeaderCaseInsensitive) {
  Workbook wb = MakeFruitWorkbook();
  SetSingleFruitCriterion(wb, "Apple");
  const Value v = EvalIn("=DSUM(A1:C5, \"sAlEs\", E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 300.0);
}

TEST(BuiltinsDatabase, DSumEmptyMatchReturnsZero) {
  Workbook wb = MakeFruitWorkbook();
  SetSingleFruitCriterion(wb, "Mango");  // nothing matches
  const Value v = EvalIn("=DSUM(A1:C5, \"Sales\", E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsDatabase, HeaderOnlyDatabaseReturnsValueError) {
  // A database range that contains only the header row (no data rows)
  // surfaces `#VALUE!` — Mac Excel 365 treats the record-less range as
  // malformed rather than as "0 matches". The check short-circuits before
  // field / criteria resolution, so every D-function behaves the same way.
  Workbook wb = MakeFruitWorkbook();
  SetSingleFruitCriterion(wb, "Apple");
  for (const char* fn : {"DSUM", "DMIN", "DMAX", "DAVERAGE", "DPRODUCT"}) {
    const std::string src = std::string("=") + fn + "(A1:C1, \"Sales\", E1:E2)";
    const Value v = EvalIn(src, wb, wb.sheet(0));
    ASSERT_TRUE(v.is_error()) << "function: " << fn;
    EXPECT_EQ(v.as_error(), ErrorCode::Value) << "function: " << fn;
  }
}

TEST(BuiltinsDatabase, DSumFieldOutOfRangeReturnsValueError) {
  Workbook wb = MakeFruitWorkbook();
  SetSingleFruitCriterion(wb, "Apple");
  const Value v = EvalIn("=DSUM(A1:C5, 4, E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsDatabase, DSumFieldNotFoundReturnsValueError) {
  Workbook wb = MakeFruitWorkbook();
  SetSingleFruitCriterion(wb, "Apple");
  const Value v = EvalIn("=DSUM(A1:C5, \"Nonexistent\", E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsDatabase, DSumBoolFieldRejected) {
  Workbook wb = MakeFruitWorkbook();
  SetSingleFruitCriterion(wb, "Apple");
  const Value v = EvalIn("=DSUM(A1:C5, TRUE, E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsDatabase, DSumComparatorCriterion) {
  Workbook wb = MakeFruitWorkbook();
  auto& s = wb.sheet(0);
  s.set_cell_value(0, 4, Value::text("Sales"));
  s.set_cell_value(1, 4, Value::text(">=100"));
  const Value v = EvalIn("=DSUM(A1:C5, \"Sales\", E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 600.0);
}

TEST(BuiltinsDatabase, DSumWildcardCriterion) {
  Workbook wb = MakeFruitWorkbook();
  auto& s = wb.sheet(0);
  s.set_cell_value(0, 4, Value::text("Fruit"));
  s.set_cell_value(1, 4, Value::text("App*"));
  const Value v = EvalIn("=DSUM(A1:C5, \"Sales\", E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 300.0);
}

TEST(BuiltinsDatabase, DSumMultiColumnAndCriterion) {
  Workbook wb = MakeFruitWorkbook();
  auto& s = wb.sheet(0);
  s.set_cell_value(0, 4, Value::text("Fruit"));
  s.set_cell_value(0, 5, Value::text("Region"));
  s.set_cell_value(1, 4, Value::text("Apple"));
  s.set_cell_value(1, 5, Value::text("East"));
  const Value v = EvalIn("=DSUM(A1:C5, \"Sales\", E1:F2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 100.0);
}

TEST(BuiltinsDatabase, DSumMultiRowOrCriterion) {
  Workbook wb = MakeFruitWorkbook();
  auto& s = wb.sheet(0);
  s.set_cell_value(0, 4, Value::text("Fruit"));
  s.set_cell_value(1, 4, Value::text("Apple"));
  s.set_cell_value(2, 4, Value::text("Pear"));
  const Value v = EvalIn("=DSUM(A1:C5, \"Sales\", E1:E3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 350.0);
}

TEST(BuiltinsDatabase, DSumCriteriaHeaderOnlyReturnsValueError) {
  // Mac Excel 365 (16.108.1, ja-JP) rejects a criteria block with only a
  // header row (no criterion rows beneath it) by surfacing #VALUE!. The
  // shape is malformed: the contract requires at least one criterion row.
  Workbook wb = MakeFruitWorkbook();
  auto& s = wb.sheet(0);
  s.set_cell_value(0, 4, Value::text("Fruit"));
  const Value v = EvalIn("=DSUM(A1:C5, \"Sales\", E1:E1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsDatabase, DSumErrorCellsInDatabaseSilentlySkipped) {
  Workbook wb = MakeFruitWorkbook();
  auto& s = wb.sheet(0);
  // Corrupt one record with an error cell in the Sales column.
  s.set_cell_value(2, 2, Value::error(ErrorCode::Div0));
  SetSingleFruitCriterion(wb, "Apple");
  const Value v = EvalIn("=DSUM(A1:C5, \"Sales\", E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  // Only the remaining Apple (row 2, 100) contributes.
  EXPECT_DOUBLE_EQ(v.as_number(), 100.0);
}

TEST(BuiltinsDatabase, DSumNonNumericCellsSilentlySkipped) {
  Workbook wb = MakeFruitWorkbook();
  auto& s = wb.sheet(0);
  // Make one Sales cell text that cannot parse to a number.
  s.set_cell_value(2, 2, Value::text("not a number"));
  SetSingleFruitCriterion(wb, "Apple");
  const Value v = EvalIn("=DSUM(A1:C5, \"Sales\", E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 100.0);
}

// ---------------------------------------------------------------------------
// DCOUNT / DCOUNTA
// ---------------------------------------------------------------------------

TEST(BuiltinsDatabase, DCountCountsNumericOnly) {
  Workbook wb = MakeFruitWorkbook();
  auto& s = wb.sheet(0);
  // Replace one Sales with text; it should NOT be counted by DCOUNT.
  s.set_cell_value(4, 2, Value::text("tbd"));
  // Wildcard "*" matches every non-blank Fruit -> every data row.
  s.set_cell_value(0, 4, Value::text("Fruit"));
  s.set_cell_value(1, 4, Value::text("*"));
  const Value v = EvalIn("=DCOUNT(A1:C5, \"Sales\", E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsDatabase, DCountEmptyMatchIsZero) {
  Workbook wb = MakeFruitWorkbook();
  SetSingleFruitCriterion(wb, "Mango");
  const Value v = EvalIn("=DCOUNT(A1:C5, \"Sales\", E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsDatabase, DCountSkipsBoolFieldCells) {
  Workbook wb = MakeFruitWorkbook();
  auto& s = wb.sheet(0);
  // Replace one Sales with a Bool; DCOUNT skips non-Number cells in the
  // field column (matches SUM / COUNT's range-argument rule and IronCalc
  // oracle `DCOUNT_DCOUNTA B54`).
  s.set_cell_value(1, 2, Value::boolean(true));
  // Wildcard "*" criterion matches every non-blank Fruit -> every record.
  s.set_cell_value(0, 4, Value::text("Fruit"));
  s.set_cell_value(1, 4, Value::text("*"));
  const Value v = EvalIn("=DCOUNT(A1:C5, \"Sales\", E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsDatabase, DCountAIncludesTextAndBool) {
  Workbook wb = MakeFruitWorkbook();
  auto& s = wb.sheet(0);
  s.set_cell_value(1, 2, Value::text("pending"));
  s.set_cell_value(2, 2, Value::boolean(true));
  // Wildcard "*" criterion matches every non-blank Fruit -> every record.
  s.set_cell_value(0, 4, Value::text("Fruit"));
  s.set_cell_value(1, 4, Value::text("*"));
  const Value v = EvalIn("=DCOUNTA(A1:C5, \"Sales\", E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  // All four data rows have a non-blank Sales cell.
  EXPECT_DOUBLE_EQ(v.as_number(), 4.0);
}

TEST(BuiltinsDatabase, DCountACountsErrorCells) {
  Workbook wb = MakeFruitWorkbook();
  auto& s = wb.sheet(0);
  s.set_cell_value(2, 2, Value::error(ErrorCode::NA));
  // Wildcard "*" criterion matches every non-blank Fruit -> every record.
  s.set_cell_value(0, 4, Value::text("Fruit"));
  s.set_cell_value(1, 4, Value::text("*"));
  const Value v = EvalIn("=DCOUNTA(A1:C5, \"Sales\", E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 4.0);
}

TEST(BuiltinsDatabase, DCountASkipsBlank) {
  Workbook wb = MakeFruitWorkbook();
  auto& s = wb.sheet(0);
  s.set_cell_value(3, 2, Value::blank());
  // Wildcard "*" criterion matches every non-blank Fruit -> every record.
  s.set_cell_value(0, 4, Value::text("Fruit"));
  s.set_cell_value(1, 4, Value::text("*"));
  const Value v = EvalIn("=DCOUNTA(A1:C5, \"Sales\", E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

// ---------------------------------------------------------------------------
// DAVERAGE
// ---------------------------------------------------------------------------

TEST(BuiltinsDatabase, DAverageHappyPath) {
  Workbook wb = MakeFruitWorkbook();
  SetSingleFruitCriterion(wb, "Apple");
  const Value v = EvalIn("=DAVERAGE(A1:C5, \"Sales\", E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 150.0);
}

TEST(BuiltinsDatabase, DAverageEmptyMatchIsDiv0) {
  Workbook wb = MakeFruitWorkbook();
  SetSingleFruitCriterion(wb, "Mango");
  const Value v = EvalIn("=DAVERAGE(A1:C5, \"Sales\", E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// DMAX / DMIN
// ---------------------------------------------------------------------------

TEST(BuiltinsDatabase, DMaxHappyPath) {
  Workbook wb = MakeFruitWorkbook();
  SetSingleFruitCriterion(wb, "Apple");
  const Value v = EvalIn("=DMAX(A1:C5, \"Sales\", E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 200.0);
}

TEST(BuiltinsDatabase, DMaxEmptyMatchIsZero) {
  Workbook wb = MakeFruitWorkbook();
  SetSingleFruitCriterion(wb, "Mango");
  const Value v = EvalIn("=DMAX(A1:C5, \"Sales\", E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsDatabase, DMinHappyPath) {
  Workbook wb = MakeFruitWorkbook();
  SetSingleFruitCriterion(wb, "Apple");
  const Value v = EvalIn("=DMIN(A1:C5, \"Sales\", E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 100.0);
}

TEST(BuiltinsDatabase, DMinEmptyMatchIsZero) {
  Workbook wb = MakeFruitWorkbook();
  SetSingleFruitCriterion(wb, "Mango");
  const Value v = EvalIn("=DMIN(A1:C5, \"Sales\", E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

// ---------------------------------------------------------------------------
// DPRODUCT
// ---------------------------------------------------------------------------

TEST(BuiltinsDatabase, DProductHappyPath) {
  Workbook wb = MakeFruitWorkbook();
  SetSingleFruitCriterion(wb, "Apple");
  const Value v = EvalIn("=DPRODUCT(A1:C5, \"Sales\", E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 100.0 * 200.0);
}

TEST(BuiltinsDatabase, DProductEmptyMatchIsZero) {
  Workbook wb = MakeFruitWorkbook();
  SetSingleFruitCriterion(wb, "Mango");
  const Value v = EvalIn("=DPRODUCT(A1:C5, \"Sales\", E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

// ---------------------------------------------------------------------------
// DSTDEV / DSTDEVP / DVAR / DVARP
// ---------------------------------------------------------------------------

TEST(BuiltinsDatabase, DStdevOneMatchIsDiv0) {
  Workbook wb = MakeFruitWorkbook();
  SetSingleFruitCriterion(wb, "Pear");  // exactly one match
  const Value v = EvalIn("=DSTDEV(A1:C5, \"Sales\", E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsDatabase, DStdevTwoMatches) {
  Workbook wb = MakeFruitWorkbook();
  SetSingleFruitCriterion(wb, "Apple");  // 100 and 200
  const Value v = EvalIn("=DSTDEV(A1:C5, \"Sales\", E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  // Sample SD of {100, 200} with divisor n-1=1: sqrt((50^2 + 50^2)/1) = sqrt(5000).
  EXPECT_NEAR(v.as_number(), std::sqrt(5000.0), 1e-9);
}

TEST(BuiltinsDatabase, DStdevPOneMatchIsZero) {
  Workbook wb = MakeFruitWorkbook();
  SetSingleFruitCriterion(wb, "Pear");  // exactly one match
  const Value v = EvalIn("=DSTDEVP(A1:C5, \"Sales\", E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsDatabase, DStdevPEmptyMatchIsDiv0) {
  Workbook wb = MakeFruitWorkbook();
  SetSingleFruitCriterion(wb, "Mango");
  const Value v = EvalIn("=DSTDEVP(A1:C5, \"Sales\", E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsDatabase, DVarOneMatchIsDiv0) {
  Workbook wb = MakeFruitWorkbook();
  SetSingleFruitCriterion(wb, "Pear");
  const Value v = EvalIn("=DVAR(A1:C5, \"Sales\", E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsDatabase, DVarTwoMatches) {
  Workbook wb = MakeFruitWorkbook();
  SetSingleFruitCriterion(wb, "Apple");
  const Value v = EvalIn("=DVAR(A1:C5, \"Sales\", E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  // Sample variance of {100, 200}: ((100-150)^2 + (200-150)^2)/(2-1) = 5000.
  EXPECT_NEAR(v.as_number(), 5000.0, 1e-9);
}

TEST(BuiltinsDatabase, DVarPOneMatchIsZero) {
  Workbook wb = MakeFruitWorkbook();
  SetSingleFruitCriterion(wb, "Pear");
  const Value v = EvalIn("=DVARP(A1:C5, \"Sales\", E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsDatabase, DVarPTwoMatches) {
  Workbook wb = MakeFruitWorkbook();
  SetSingleFruitCriterion(wb, "Apple");
  const Value v = EvalIn("=DVARP(A1:C5, \"Sales\", E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  // Population variance of {100, 200}: ((100-150)^2 + (200-150)^2)/2 = 2500.
  EXPECT_NEAR(v.as_number(), 2500.0, 1e-9);
}

// ---------------------------------------------------------------------------
// DGET
// ---------------------------------------------------------------------------

TEST(BuiltinsDatabase, DGetExactlyOneMatch) {
  Workbook wb = MakeFruitWorkbook();
  SetSingleFruitCriterion(wb, "Pear");
  const Value v = EvalIn("=DGET(A1:C5, \"Sales\", E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 50.0);
}

TEST(BuiltinsDatabase, DGetZeroMatchesIsValueError) {
  Workbook wb = MakeFruitWorkbook();
  SetSingleFruitCriterion(wb, "Mango");
  const Value v = EvalIn("=DGET(A1:C5, \"Sales\", E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsDatabase, DGetMultipleMatchesIsNumError) {
  Workbook wb = MakeFruitWorkbook();
  SetSingleFruitCriterion(wb, "Apple");  // two matches
  const Value v = EvalIn("=DGET(A1:C5, \"Sales\", E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsDatabase, DGetReturnsTextValue) {
  Workbook wb = MakeFruitWorkbook();
  SetSingleFruitCriterion(wb, "Pear");
  // Ask for the "Region" column of the matching row.
  const Value v = EvalIn("=DGET(A1:C5, \"Region\", E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "East");
}

// ---------------------------------------------------------------------------
// Argument-shape errors
// ---------------------------------------------------------------------------

TEST(BuiltinsDatabase, DatabaseLiteralArgReturnsValue) {
  Workbook wb = MakeFruitWorkbook();
  SetSingleFruitCriterion(wb, "Apple");
  const Value v = EvalIn("=DSUM(5, \"Sales\", E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsDatabase, CriteriaLiteralArgReturnsValue) {
  Workbook wb = MakeFruitWorkbook();
  const Value v = EvalIn("=DSUM(A1:C5, \"Sales\", 5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsDatabase, ArityWrongReturnsValue) {
  Workbook wb = MakeFruitWorkbook();
  const Value v = EvalIn("=DSUM(A1:C5, \"Sales\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  // Pratt parser rejects arity-2 DSUM at call time through lazy dispatch
  // returning #VALUE!.
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// Criterion header not present in database header row
// ---------------------------------------------------------------------------

TEST(BuiltinsDatabase, CriterionHeaderNotFoundDisjunctFails) {
  Workbook wb = MakeFruitWorkbook();
  auto& s = wb.sheet(0);
  s.set_cell_value(0, 4, Value::text("NotAColumn"));
  s.set_cell_value(1, 4, Value::text("anything"));
  // The criterion's header does not match any database header, and its
  // cell is non-blank -> disjunct unsatisfiable -> zero matches.
  const Value v = EvalIn("=DSUM(A1:C5, \"Sales\", E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

// ---------------------------------------------------------------------------
// Field evaluation error propagation
// ---------------------------------------------------------------------------

TEST(BuiltinsDatabase, FieldErrorPropagates) {
  Workbook wb = MakeFruitWorkbook();
  SetSingleFruitCriterion(wb, "Apple");
  const Value v = EvalIn("=DSUM(A1:C5, 1/0, E1:E2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// Field-arg shape rejection (multi-cell RangeOp)
//
// Mac Excel 365 surfaces `#VALUE!` when the `field` argument is a multi-
// cell range (e.g. `$E$1:$F$1`), rather than projecting it through dynamic-
// array spill semantics to the top-left cell. These guards live in
// `resolve_common`; the regression they protect against came from the
// `RangeOp` evaluator now returning a spill anchor instead of `#VALUE!`.
//
// The shared layout for these tests:
//   A1 = "Fruit"  B1 = "Region"  C1 = "Sales"  D1 = "Qty"  E1 = "Year"
//   A2 = "Apple"  B2 = "East"    C2 = 100      D2 = 5      E2 = 2024
//   ... (4 data rows)
//   A14 = "Fruit"
//   A15 = "Apple"
// This puts the criterion block at A14:A15 and lets `$E$1:$F$1` denote
// the multi-cell field range used in the oracle case.
Workbook MakeFruitWorkbookWide() {
  Workbook wb = Workbook::create();
  auto& s = wb.sheet(0);
  s.set_cell_value(0, 0, Value::text("Fruit"));
  s.set_cell_value(0, 1, Value::text("Region"));
  s.set_cell_value(0, 2, Value::text("Sales"));
  s.set_cell_value(0, 3, Value::text("Qty"));
  s.set_cell_value(0, 4, Value::text("Year"));
  s.set_cell_value(1, 0, Value::text("Apple"));
  s.set_cell_value(1, 1, Value::text("East"));
  s.set_cell_value(1, 2, Value::number(100.0));
  s.set_cell_value(1, 3, Value::number(5.0));
  s.set_cell_value(1, 4, Value::number(2024.0));
  s.set_cell_value(2, 0, Value::text("Apple"));
  s.set_cell_value(2, 1, Value::text("West"));
  s.set_cell_value(2, 2, Value::number(200.0));
  s.set_cell_value(2, 3, Value::number(8.0));
  s.set_cell_value(2, 4, Value::number(2024.0));
  s.set_cell_value(3, 0, Value::text("Pear"));
  s.set_cell_value(3, 1, Value::text("East"));
  s.set_cell_value(3, 2, Value::number(50.0));
  s.set_cell_value(3, 3, Value::number(3.0));
  s.set_cell_value(3, 4, Value::number(2023.0));
  s.set_cell_value(4, 0, Value::text("Banana"));
  s.set_cell_value(4, 1, Value::text("West"));
  s.set_cell_value(4, 2, Value::number(300.0));
  s.set_cell_value(4, 3, Value::number(10.0));
  s.set_cell_value(4, 4, Value::number(2025.0));
  // Criterion block at A14:A15 (Fruit == Apple).
  s.set_cell_value(13, 0, Value::text("Fruit"));
  s.set_cell_value(14, 0, Value::text("Apple"));
  return wb;
}

TEST(BuiltinsDatabase, DCountFieldMultiCellRangeRejectsWithValue) {
  Workbook wb = MakeFruitWorkbookWide();
  // `$E$1:$F$1` spans two cells horizontally — Mac rejects with #VALUE!.
  const Value v = EvalIn("=DCOUNT(A1:E11, $E$1:$F$1, A14:A15)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsDatabase, DGetFieldMultiCellRangeRejectsWithValue) {
  Workbook wb = MakeFruitWorkbookWide();
  const Value v = EvalIn("=DGET(A1:F11, $E$1:$F$1, A14:A15)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsDatabase, DSumFieldMultiCellRangeRejectsWithValue) {
  Workbook wb = MakeFruitWorkbookWide();
  const Value v = EvalIn("=DSUM(A1:F11, $E$1:$F$1, A14:A15)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsDatabase, DCountFieldSingleCellRangeWorksAsScalarRef) {
  Workbook wb = MakeFruitWorkbookWide();
  // 1x1 RangeOp `$E$1:$E$1` must pass through as the scalar header "Year".
  // Two Apple rows; both have numeric Year values -> count is 2.
  const Value v = EvalIn("=DCOUNT(A1:E11, $E$1:$E$1, A14:A15)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number()) << "got error code " << static_cast<int>(v.as_error());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsDatabase, DCountFieldSingleScalarWorksAsScalarRef) {
  Workbook wb = MakeFruitWorkbookWide();
  // Plain Ref `$E$1` (no RangeOp) — canonical case, regression guard.
  const Value v = EvalIn("=DCOUNT(A1:E11, $E$1, A14:A15)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number()) << "got error code " << static_cast<int>(v.as_error());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

// ---------------------------------------------------------------------------
// Excel's documented "begins-with" semantics for plain text criteria
//
// Per Microsoft Support "Examples of criteria":
//   - "Sm" finds rows that begin with "Sm" such as "Smith" and "Smithfield".
//   - "Davolio" finds rows that contain "Davolio", "DAVOLIO", "davolio".
// Formulon's D-functions therefore interpret a plain-text criterion as a
// case-insensitive prefix match. An explicit "=Sm" opts out and requests
// exact equality (Excel strips the "="). Wildcards / comparators / numeric
// criteria keep their existing semantics.
//
// Shared layout (sheet 0):
//   A1=Tag    B1=Sales
//   A2=b      B2=10
//   A3=B      B3=20
//   A4=BB     B4=30
//   A5=balpha B5=40
//   A6=cherry B6=50
//   D1=Tag   (criterion header reused for every test)
//   D2=<per-test criterion cell>
Workbook MakeBeginsWithWorkbook() {
  Workbook wb = Workbook::create();
  auto& s = wb.sheet(0);
  s.set_cell_value(0, 0, Value::text("Tag"));
  s.set_cell_value(0, 1, Value::text("Sales"));
  s.set_cell_value(1, 0, Value::text("b"));
  s.set_cell_value(1, 1, Value::number(10.0));
  s.set_cell_value(2, 0, Value::text("B"));
  s.set_cell_value(2, 1, Value::number(20.0));
  s.set_cell_value(3, 0, Value::text("BB"));
  s.set_cell_value(3, 1, Value::number(30.0));
  s.set_cell_value(4, 0, Value::text("balpha"));
  s.set_cell_value(4, 1, Value::number(40.0));
  s.set_cell_value(5, 0, Value::text("cherry"));
  s.set_cell_value(5, 1, Value::number(50.0));
  s.set_cell_value(0, 3, Value::text("Tag"));
  return wb;
}

TEST(BuiltinsDatabase, DCountPlainTextIsPrefixMatch) {
  Workbook wb = MakeBeginsWithWorkbook();
  wb.sheet(0).set_cell_value(1, 3, Value::text("b"));
  // Prefix "b" matches b, B, BB, balpha (case-insensitive): 4 rows.
  const Value v = EvalIn("=DCOUNT(A1:B6, \"Sales\", D1:D2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number()) << "got error code " << static_cast<int>(v.as_error());
  EXPECT_DOUBLE_EQ(v.as_number(), 4.0);
}

TEST(BuiltinsDatabase, DSumPlainTextIsPrefixMatch) {
  Workbook wb = MakeBeginsWithWorkbook();
  wb.sheet(0).set_cell_value(1, 3, Value::text("b"));
  // 10 + 20 + 30 + 40 = 100 over the four prefix-matching rows.
  const Value v = EvalIn("=DSUM(A1:B6, \"Sales\", D1:D2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number()) << "got error code " << static_cast<int>(v.as_error());
  EXPECT_DOUBLE_EQ(v.as_number(), 100.0);
}

TEST(BuiltinsDatabase, DMinPlainTextIsPrefixMatch) {
  Workbook wb = MakeBeginsWithWorkbook();
  wb.sheet(0).set_cell_value(1, 3, Value::text("b"));
  const Value v = EvalIn("=DMIN(A1:B6, \"Sales\", D1:D2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number()) << "got error code " << static_cast<int>(v.as_error());
  EXPECT_DOUBLE_EQ(v.as_number(), 10.0);
}

TEST(BuiltinsDatabase, DMaxPlainTextIsPrefixMatch) {
  Workbook wb = MakeBeginsWithWorkbook();
  wb.sheet(0).set_cell_value(1, 3, Value::text("b"));
  const Value v = EvalIn("=DMAX(A1:B6, \"Sales\", D1:D2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number()) << "got error code " << static_cast<int>(v.as_error());
  EXPECT_DOUBLE_EQ(v.as_number(), 40.0);
}

TEST(BuiltinsDatabase, DSumExplicitEqualsIsExactMatch) {
  Workbook wb = MakeBeginsWithWorkbook();
  // `="b"` escapes to exact equality (case-insensitive): matches "b" and
  // "B" only (10 + 20 = 30), not "BB" or "balpha".
  wb.sheet(0).set_cell_value(1, 3, Value::text("=b"));
  const Value v = EvalIn("=DSUM(A1:B6, \"Sales\", D1:D2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number()) << "got error code " << static_cast<int>(v.as_error());
  EXPECT_DOUBLE_EQ(v.as_number(), 30.0);
}

TEST(BuiltinsDatabase, DSumWildcardIsUnchanged) {
  Workbook wb = MakeBeginsWithWorkbook();
  // "B*" is a wildcard pattern: same four prefix-matching rows as plain "b".
  wb.sheet(0).set_cell_value(1, 3, Value::text("B*"));
  const Value v = EvalIn("=DSUM(A1:B6, \"Sales\", D1:D2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number()) << "got error code " << static_cast<int>(v.as_error());
  EXPECT_DOUBLE_EQ(v.as_number(), 100.0);
}

TEST(BuiltinsDatabase, DSumLongerPrefixDoesNotMatchShorterCells) {
  Workbook wb = MakeBeginsWithWorkbook();
  // Prefix "ba" matches only "balpha" (cells "b" and "B" are shorter than
  // the prefix and cannot begin with "ba").
  wb.sheet(0).set_cell_value(1, 3, Value::text("ba"));
  const Value v = EvalIn("=DSUM(A1:B6, \"Sales\", D1:D2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number()) << "got error code " << static_cast<int>(v.as_error());
  EXPECT_DOUBLE_EQ(v.as_number(), 40.0);
}

TEST(BuiltinsDatabase, DSumNotEqualComparatorIsExact) {
  Workbook wb = MakeBeginsWithWorkbook();
  // `<>b` is a NotEq comparator — NOT a prefix test. Only "cherry" is
  // textually unequal to "b"; the other four ("b", "B", "BB", "balpha")
  // are compared exactly, and "b" / "B" are case-insensitively equal to
  // "b" so they fail the predicate. "BB" and "balpha" are unequal though,
  // so they pass.
  //   cherry (50) + BB (30) + balpha (40) = 120.
  wb.sheet(0).set_cell_value(1, 3, Value::text("<>b"));
  const Value v = EvalIn("=DSUM(A1:B6, \"Sales\", D1:D2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number()) << "got error code " << static_cast<int>(v.as_error());
  EXPECT_DOUBLE_EQ(v.as_number(), 120.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
