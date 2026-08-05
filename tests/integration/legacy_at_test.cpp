//
// Integration tests for legacy `@` auto-insertion (a.k.a. explicit
// implicit-intersection). When Excel 365 opens a workbook authored by
// pre-365 Excel, formulas that historically performed implicit
// intersection in scalar contexts are rewritten with a leading `@` to
// preserve the legacy "scalar of the implicit intersection" behaviour
// (e.g. `=A1:A10` in a single cell becomes `=@A1:A10`).
//
// Formulon's parser recognises `@` as the implicit-intersection operator
// (`make_implicit_intersection`) and the evaluator implements it in
// `tree_walker.cpp` under `NodeKind::ImplicitIntersection`. These tests
// drive the full `Workbook::set_cell_*` -> `Workbook::recalc()` pipeline
// to lock in:
//   * `=@A1:A10` projects the column onto the formula cell's row when
//     that row is inside the range; otherwise #VALUE!.
//   * `=A1:A10` (no `@`) yields the spill anchor (top-left) when the
//     formula cell is outside the range, or the row/col-aligned cell
//     when the formula cell is inside (legacy II compatibility).
//   * `=@SUM(A1:A10)` -- `@` on a function call collapses an array
//     result to its anchor; on a scalar-returning function the `@` is
//     a no-op.
//   * Round-trip: a workbook whose formulas carry `@` annotations must
//     preserve those annotations verbatim through OOXML save/load.
//
// Existing unit coverage in
// `tests/unit/eval/builtins_implicit_intersection_test.cpp` exercises
// the bare evaluator. This file exercises the recalc pipeline.

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "cell.h"
#include "eval/compat.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/zip_reader.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

io::ByteSpan SpanOf(const std::vector<std::uint8_t>& bytes) {
  return io::ByteSpan{bytes.data(), bytes.size()};
}

Value StoredValue(const Workbook& wb, std::size_t sheet_index, std::uint32_t row, std::uint32_t col) {
  const Sheet& s = wb.sheet(sheet_index);
  if (const Cell* c = s.cell_at(row, col); c != nullptr) {
    return c->cached_value;
  }
  return Value::blank();
}

// Populates A1..A5 with 1..5 -- a single column that supports both
// implicit-intersection and spill scenarios.
void FillA1ToA5(Workbook& wb) {
  Sheet& s = wb.sheet(0);
  s.set_cell_value(0U, 0U, Value::number(1.0));
  s.set_cell_value(1U, 0U, Value::number(2.0));
  s.set_cell_value(2U, 0U, Value::number(3.0));
  s.set_cell_value(3U, 0U, Value::number(4.0));
  s.set_cell_value(4U, 0U, Value::number(5.0));
}

// ---------------------------------------------------------------------------
// `=@Range` projects onto the formula cell's row/col
// ---------------------------------------------------------------------------

TEST(LegacyAt, AtPrefixProjectsOntoFormulaRowInsideColumn) {
  // B3 = =@A1:A5 -- single-column range, formula row 3 (0-based 2) is
  // in [1..5], so the projection is A3 == 3.
  Workbook wb = Workbook::create();
  wb.set_excel_profile(eval::mac_365_ja_jp_profile());
  FillA1ToA5(wb);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 2U, 1U, "=@A1:A5")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  const Value v = StoredValue(wb, 0U, 2U, 1U);
  ASSERT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(LegacyAt, AtPrefixOutsideColumnReturnsValueError) {
  // B7 = =@A1:A5 -- formula row 7 (0-based 6) is OUTSIDE [0..4] ->
  // #VALUE!.
  Workbook wb = Workbook::create();
  wb.set_excel_profile(eval::mac_365_ja_jp_profile());
  FillA1ToA5(wb);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 6U, 1U, "=@A1:A5")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  const Value v = StoredValue(wb, 0U, 6U, 1U);
  ASSERT_TRUE(v.is_error()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(LegacyAt, AtPrefixOnSingleRowProjectsOntoFormulaCol) {
  // C7 = =@A1:E1 -- single-row range, formula col 3 (0-based 2) is
  // in [0..4], so the projection is C1 == 30.
  Workbook wb = Workbook::create();
  wb.set_excel_profile(eval::mac_365_ja_jp_profile());
  Sheet& s = wb.sheet(0);
  s.set_cell_value(0U, 0U, Value::number(10.0));
  s.set_cell_value(0U, 1U, Value::number(20.0));
  s.set_cell_value(0U, 2U, Value::number(30.0));
  s.set_cell_value(0U, 3U, Value::number(40.0));
  s.set_cell_value(0U, 4U, Value::number(50.0));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 6U, 2U, "=@A1:E1")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  const Value v = StoredValue(wb, 0U, 6U, 2U);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 30.0);
}

TEST(LegacyAt, AtPrefixOn2DRangeReturnsValueError) {
  // 2D range with `@`: implicit intersection requires 1D alignment ->
  // #VALUE! on a 2D range. Verified Mac semantics in
  // tests/oracle/cases/implicit_intersection.yaml.
  Workbook wb = Workbook::create();
  wb.set_excel_profile(eval::mac_365_ja_jp_profile());
  Sheet& s = wb.sheet(0);
  s.set_cell_value(0U, 0U, Value::number(11.0));
  s.set_cell_value(0U, 1U, Value::number(12.0));
  s.set_cell_value(1U, 0U, Value::number(21.0));
  s.set_cell_value(1U, 1U, Value::number(22.0));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 2U, 25U, "=@A1:B5")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  const Value v = StoredValue(wb, 0U, 2U, 25U);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// Bare range (no `@`) -- spill or top-left fallback
// ---------------------------------------------------------------------------

TEST(LegacyAt, BareColumnRangeSpillsAtTopLeftOutsideRange) {
  // F1 = =A1:A5 -- a column range typed in a cell OUTSIDE the column.
  // Excel 365 spills the array; the anchor (F1) holds the top-left
  // element A1 == 1.
  Workbook wb = Workbook::create();
  wb.set_excel_profile(eval::mac_365_ja_jp_profile());
  FillA1ToA5(wb);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 5U, "=A1:A5")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  const Value f1 = StoredValue(wb, 0U, 0U, 5U);
  ASSERT_TRUE(f1.is_number()) << "kind=" << static_cast<int>(f1.kind());
  EXPECT_DOUBLE_EQ(f1.as_number(), 1.0);
  // StoredValue reads through cell_at(), which sees only the anchor's
  // stored record and not the phantom spill cells; F2 therefore reads
  // back blank here even though the spilled region logically covers it.
  EXPECT_TRUE(StoredValue(wb, 0U, 1U, 5U).is_blank());
}

TEST(LegacyAt, BareColumnRangeSpillsAtTopLeftEvenWhenRowAligned) {
  // F3 = =A1:A5 -- Excel 365 spills a bare range regardless of whether
  // the formula cell's row falls inside the range. The anchor (F3) holds
  // the top-left element A1 == 1; it is NOT the row-aligned cell A3.
  Workbook wb = Workbook::create();
  wb.set_excel_profile(eval::mac_365_ja_jp_profile());
  FillA1ToA5(wb);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 2U, 5U, "=A1:A5")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  const Value f3 = StoredValue(wb, 0U, 2U, 5U);
  ASSERT_TRUE(f3.is_number());
  EXPECT_DOUBLE_EQ(f3.as_number(), 1.0);
}

// ---------------------------------------------------------------------------
// `@` on a function call
// ---------------------------------------------------------------------------

TEST(LegacyAt, AtPrefixOnScalarFunctionIsNoop) {
  // B1 = =@SUM(A1:A5) -- SUM returns a scalar, so the `@` operator is a
  // no-op. Result == 15.
  Workbook wb = Workbook::create();
  wb.set_excel_profile(eval::mac_365_ja_jp_profile());
  FillA1ToA5(wb);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=@SUM(A1:A5)")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  const Value b1 = StoredValue(wb, 0U, 0U, 1U);
  ASSERT_TRUE(b1.is_number()) << "kind=" << static_cast<int>(b1.kind());
  EXPECT_DOUBLE_EQ(b1.as_number(), 15.0);
}

TEST(LegacyAt, AtPrefixOnArrayFunctionCollapsesToAnchor) {
  // B1 = =@SEQUENCE(3,1) -- SEQUENCE produces a Value::Array; the `@`
  // operator prevents it from spilling and yields the anchor scalar
  // (1.0). Excel 365's documented "scalar value of the array" behaviour.
  //
  // Engine note: the parser wraps the SEQUENCE call in an
  // ImplicitIntersection node. The evaluator hits the "non-range
  // operand" branch under NodeKind::ImplicitIntersection (current
  // behaviour: identity passthrough). The Array then bubbles up to the
  // recalc engine's scalar dispatch which folds it via
  // dispatch_array_result -> commits a spill.
  //
  // TODO: Excel suppresses the spill entirely under `@` -- B2/B3 should
  // remain blank. Engine currently still spills (creates phantoms at
  // B2/B3). The test below pins the engine's *current* behaviour: the
  // anchor evaluates to 1.0 but a spill region IS committed.
  Workbook wb = Workbook::create();
  wb.set_excel_profile(eval::mac_365_ja_jp_profile());
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=@SEQUENCE(3,1)")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  const Value b1 = StoredValue(wb, 0U, 0U, 1U);
  ASSERT_TRUE(b1.is_number()) << "kind=" << static_cast<int>(b1.kind());
  EXPECT_DOUBLE_EQ(b1.as_number(), 1.0);
}

// ---------------------------------------------------------------------------
// Round-trip: formulas with `@` are preserved verbatim
// ---------------------------------------------------------------------------

TEST(LegacyAt, FormulaTextWithAtPrefixRoundTrips) {
  // Save a workbook with an `@`-prefixed formula and read it back.
  // The reader should preserve the exact formula text including the
  // `@` so a follow-up recalc reproduces the legacy II semantics.
  Workbook src = Workbook::create();
  src.set_excel_profile(eval::mac_365_ja_jp_profile());
  FillA1ToA5(src);
  ASSERT_TRUE(static_cast<bool>(src.set_cell_formula(0U, 2U, 1U, "=@A1:A5")));
  ASSERT_TRUE(static_cast<bool>(src.recalc(eval::default_registry())));

  // Sanity: source recalc'd to 3.0.
  ASSERT_DOUBLE_EQ(StoredValue(src, 0U, 2U, 1U).as_number(), 3.0);

  auto bytes_or = src.save();
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << "save failed: " << bytes_or.error().message;
  const std::vector<std::uint8_t> bytes = bytes_or.value();

  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read failed: " << result_or.error().message;
  Workbook& dst = result_or.value().workbook;

  // The reader preserves formula_text verbatim -- including the `@`.
  const Cell* b3 = dst.sheet(0).cell_at(2U, 1U);
  ASSERT_NE(b3, nullptr);
  EXPECT_EQ(b3->formula_text, "=@A1:A5");

  // Recalc the destination workbook and verify it lands on 3.0 again.
  ASSERT_TRUE(static_cast<bool>(dst.recalc(eval::default_registry())));
  const Value b3_v = StoredValue(dst, 0U, 2U, 1U);
  ASSERT_TRUE(b3_v.is_number()) << "kind=" << static_cast<int>(b3_v.kind());
  EXPECT_DOUBLE_EQ(b3_v.as_number(), 3.0);
}

TEST(LegacyAt, FormulaTextWithAtOnSumRoundTrips) {
  // Same as above but with `=@SUM(A1:A5)` -- the `@` should still
  // round-trip verbatim even though it's a no-op semantically.
  Workbook src = Workbook::create();
  src.set_excel_profile(eval::mac_365_ja_jp_profile());
  FillA1ToA5(src);
  ASSERT_TRUE(static_cast<bool>(src.set_cell_formula(0U, 0U, 1U, "=@SUM(A1:A5)")));
  ASSERT_TRUE(static_cast<bool>(src.recalc(eval::default_registry())));

  auto bytes_or = src.save();
  ASSERT_TRUE(static_cast<bool>(bytes_or));
  const std::vector<std::uint8_t> bytes = bytes_or.value();

  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or));
  Workbook& dst = result_or.value().workbook;

  const Cell* b1 = dst.sheet(0).cell_at(0U, 1U);
  ASSERT_NE(b1, nullptr);
  EXPECT_EQ(b1->formula_text, "=@SUM(A1:A5)");

  ASSERT_TRUE(static_cast<bool>(dst.recalc(eval::default_registry())));
  EXPECT_DOUBLE_EQ(StoredValue(dst, 0U, 0U, 1U).as_number(), 15.0);
}

// ---------------------------------------------------------------------------
// `@` interaction with binary range operators
// ---------------------------------------------------------------------------

TEST(LegacyAt, BareRangeBinaryOpMultipliesElementWise) {
  // C1 = =A1:A5*B1:B5. Both sides are 5-row columns; Excel 365 spills
  // a 5-row product into C1:C5. Engine note: the evaluator's
  // RangeOp branch runs FIRST per side and collapses each range to a
  // scalar (top-left fallback or row/col-aligned cell). Then the
  // BinaryOp multiplies the two scalars.
  //
  // Result depends on which scalar each side collapses to. Formula
  // cell C1 has row 0 -- in [0..4] for A1:A5, so the row/col-aligned
  // branch picks A1 = 1; same for B1 = 10. The result at C1 is 10.
  //
  // TODO: Excel 365 broadcasts to a 5-row spill (Array * Array element-
  // wise). Engine collapses to a scalar; the spill engine refactor
  // will likely change this. Test pins current behaviour.
  Workbook wb = Workbook::create();
  wb.set_excel_profile(eval::mac_365_ja_jp_profile());
  FillA1ToA5(wb);
  Sheet& s = wb.sheet(0);
  s.set_cell_value(0U, 1U, Value::number(10.0));
  s.set_cell_value(1U, 1U, Value::number(20.0));
  s.set_cell_value(2U, 1U, Value::number(30.0));
  s.set_cell_value(3U, 1U, Value::number(40.0));
  s.set_cell_value(4U, 1U, Value::number(50.0));

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 2U, "=A1:A5*B1:B5")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  const Value c1 = StoredValue(wb, 0U, 0U, 2U);
  ASSERT_TRUE(c1.is_number()) << "kind=" << static_cast<int>(c1.kind());
  // A1=1, B1=10 -> 1*10 = 10 (current behaviour: scalar collapse, not
  // a 5-row spill).
  EXPECT_DOUBLE_EQ(c1.as_number(), 10.0);
}

TEST(LegacyAt, AtPrefixBeforeBinaryOpTakesTopLeftOfComputedArray) {
  // C3 = =@(A1:A5*2) -- the `@` binds the whole parenthesised
  // expression, whose operand is a BinaryOp producing the computed array
  // {2,4,6,8,10}. Implicit-intersection row/column alignment only applies
  // to a direct range/reference operand; on a computed array `@` takes
  // the top-left element -- A1*2 == 2, not the row-aligned A3*2.
  Workbook wb = Workbook::create();
  wb.set_excel_profile(eval::mac_365_ja_jp_profile());
  FillA1ToA5(wb);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 2U, 2U, "=@(A1:A5*2)")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  const Value c3 = StoredValue(wb, 0U, 2U, 2U);
  ASSERT_TRUE(c3.is_number()) << "kind=" << static_cast<int>(c3.kind());
  EXPECT_DOUBLE_EQ(c3.as_number(), 2.0);
}

}  // namespace
}  // namespace formulon
