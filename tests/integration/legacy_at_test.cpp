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

#include <array>
#include <cstddef>
#include <cstdint>
#include <sstream>
#include <string>
#include <string_view>
#include <vector>

#include "cell.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "sheet.h"
#include "support/ooxml_package_fixture.h"
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
  wb.set_excel_profile(mac_365_ja_jp_profile());
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
  wb.set_excel_profile(mac_365_ja_jp_profile());
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
  wb.set_excel_profile(mac_365_ja_jp_profile());
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
  wb.set_excel_profile(mac_365_ja_jp_profile());
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
  wb.set_excel_profile(mac_365_ja_jp_profile());
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
  wb.set_excel_profile(mac_365_ja_jp_profile());
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
  wb.set_excel_profile(mac_365_ja_jp_profile());
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
  wb.set_excel_profile(mac_365_ja_jp_profile());
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
  src.set_excel_profile(mac_365_ja_jp_profile());
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
  src.set_excel_profile(mac_365_ja_jp_profile());
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
  wb.set_excel_profile(mac_365_ja_jp_profile());
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
  wb.set_excel_profile(mac_365_ja_jp_profile());
  FillA1ToA5(wb);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 2U, 2U, "=@(A1:A5*2)")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  const Value c3 = StoredValue(wb, 0U, 2U, 2U);
  ASSERT_TRUE(c3.is_number()) << "kind=" << static_cast<int>(c3.kind());
  EXPECT_DOUBLE_EQ(c3.as_number(), 2.0);
}

// ---------------------------------------------------------------------------
// A loaded legacy formula (no dynamic-array mark) reads as Excel 365 shows it
// ---------------------------------------------------------------------------

struct LegacyCase {
  const char* stored;    // the formula as a legacy file stores it
  const char* formula2;  // Excel 365's formula2 text for it
  const char* value;     // Excel's value in column E of that row; "" when volatile
};

// Excel 365 (Mac, ja-JP) opening a legacy workbook: A1:B3 hold 1..3 per row,
// Rng = A1:A2, Cel = A1, Val = 5, Fn = LAMBDA(x,x*2), each formula in E<row>;
// formula2 text and values as Excel showed them.
constexpr LegacyCase kLegacy61[] = {
    {"=A1", "=A1", "1"},
    {"=A1+1", "=A1+1", "2"},
    {"=A1:A2", "=@A1:A2", "#VALUE!"},
    {"=A1:A2*2", "=@A1:A2*2", "#VALUE!"},
    {"=SUM(A1:A2)", "=SUM(A1:A2)", "3"},
    {"=SUM(A1:A2*2)", "=SUM(@A1:A2*2)", "#VALUE!"},
    {"=A1:A2", "=@A1:A2", "#VALUE!"},
    {"=Sheet1!A1", "=Sheet1!A1", "1"},
    {"=NOSUCH(1)", "=@NOSUCH(1)", "#NAME?"},
    {"=NOSUCHREF", "=@NOSUCHREF", "#NAME?"},
    {"=Fn(1)", "=@Fn(1)", "2"},
    {"=A1(1)", "=@A1(1)", "#REF!"},
    {"=A1:B2 B2:C3", "=@A1:B2 B2:C3", "2"},
    {"=A1 B1", "=A1 B1", "#NULL!"},
    {"=(A1:A2,B1:B2)", "=(A1:A2,B1:B2)", "#VALUE!"},
    {"=INDEX(A1:A2,1)", "=INDEX(A1:A2,1)", "1"},
    {"=INDEX(A1:B2,1,0)", "=@INDEX(A1:B2,1,0)", "#VALUE!"},
    {"=OFFSET(A1,0,0)", "=OFFSET(A1,0,0)", "1"},
    {"=OFFSET(A1,0,0,2)", "=@OFFSET(A1,0,0,2)", "#VALUE!"},
    {"=INDIRECT(\"A1\")", "=@INDIRECT(\"A1\")", "1"},
    {"=SEQUENCE(1)", "=@SEQUENCE(1)", "1"},
    {"=SEQUENCE(2)", "=@SEQUENCE(2)", "1"},
    {"=Rng", "=@Rng", "#VALUE!"},
    {"=Cel", "=Cel", "1"},
    {"=Val", "=Val", "5"},
    {"=IF(A1>0,1,0)", "=IF(A1>0,1,0)", "1"},
    {"=IF(A1:A2>0,1,0)", "=IF(@A1:A2>0,1,0)", "#VALUE!"},
    {"=ROW()", "=ROW()", "28"},
    {"=ROW(A1:A2)", "=@ROW(A1:A2)", "1"},
    {"=ROWS(A1:A2)", "=ROWS(A1:A2)", "2"},
    {"=VLOOKUP(1,A1:B2,2,0)", "=VLOOKUP(1,A1:B2,2,0)", "1"},
    {"=TODAY()", "=TODAY()", ""},
    {"=RAND()", "=RAND()", ""},
    {"=TRANSPOSE(A1)", "=@TRANSPOSE(A1)", "1"},
    {"=LET(x,1,x)", "=LET(x,1,x)", "1"},
    {"=LAMBDA(x,x)(1)", "=@LAMBDA(x,x)(1)", "1"},
    {"=CHOOSE(1,A1)", "=CHOOSE(1,A1)", "1"},
    {"=CHOOSE(1,A1:A2)", "=@CHOOSE(1,A1:A2)", "#VALUE!"},
    {"=IFERROR(A1,0)", "=IFERROR(A1,0)", "1"},
    {"=N(A1)", "=N(A1)", "1"},
    {"=ISNUMBER(A1:A2)", "=ISNUMBER(@A1:A2)", "FALSE"},
    {"=ABS(A1:A2)", "=ABS(@A1:A2)", "#VALUE!"},
    {"=ABS(A1)", "=ABS(A1)", "1"},
    {"=MMULT(A1,A1)", "=@MMULT(A1,A1)", "1"},
    {"=SUMPRODUCT(A1:A2)", "=SUMPRODUCT(A1:A2)", "3"},
    {"=COUNTIF(A1:A2,1)", "=COUNTIF(A1:A2,1)", "1"},
    {"=COUNTIF(A1:A2,A1:A2)", "=COUNTIF(A1:A2,@A1:A2)", "0"},
    {"=MATCH(1,A1:A2,0)", "=MATCH(1,A1:A2,0)", "1"},
    {"=A1&\"x\"", "=A1&\"x\"", "1x"},
    {"=\"x\"", "=\"x\"", "x"},
    {"=1", "=1", "1"},
    {"=TEXT(A1,\"0\")", "=TEXT(A1,\"0\")", "1"},
    {"=LEN(A1:A2)", "=LEN(@A1:A2)", "#VALUE!"},
    {"=Rng+1", "=@Rng+1", "#VALUE!"},
    {"=SUM(Rng)", "=SUM(Rng)", "3"},
    {"=1/0", "=1/0", "#DIV/0!"},
    {"=NA()", "=NA()", "#N/A"},
    {"=XLOOKUP(1,A1:A2,B1:B2)", "=XLOOKUP(1,A1:A2,B1:B2)", "1"},
    {"=UNIQUE(A1)", "=@UNIQUE(A1)", "1"},
    {"=MAX(A1:A2)", "=MAX(A1:A2)", "2"},
    {"=AND(A1:A2>0)", "=AND(@A1:A2>0)", "#VALUE!"},
};

// INDEX / OFFSET shapes on the same sheet with A5 = 0, A6 = 1,
// each written the same way.
constexpr LegacyCase kLegacyIndexOffset[] = {
    {"=INDEX(A1:B2,0,1)", "=@INDEX(A1:B2,0,1)", "1"},
    {"=INDEX(A1:B2,1)", "=INDEX(A1:B2,1)", "#REF!"},
    {"=INDEX(A1:B2,1,1)", "=INDEX(A1:B2,1,1)", "1"},
    {"=INDEX(A1:B2,,1)", "=@INDEX(A1:B2,,1)", "#VALUE!"},
    {"=INDEX(A1:B2,A5,1)", "=@INDEX(A1:B2,A5,1)", "#VALUE!"},
    {"=INDEX(A1:A2,A5)", "=@INDEX(A1:A2,A5)", "#VALUE!"},
    {"=INDEX(A1:B2,1,A5)", "=@INDEX(A1:B2,1,A5)", "#VALUE!"},
    {"=INDEX(A1:A2,0)", "=@INDEX(A1:A2,0)", "#VALUE!"},
    {"=OFFSET(A1,0,0,1)", "=OFFSET(A1,0,0,1)", "1"},
    {"=OFFSET(A1,0,0,1,1)", "=OFFSET(A1,0,0,1,1)", "1"},
    {"=OFFSET(A1,0,0,2,1)", "=@OFFSET(A1,0,0,2,1)", "#VALUE!"},
    {"=OFFSET(A1,0,0,1,2)", "=@OFFSET(A1,0,0,1,2)", "#VALUE!"},
    {"=OFFSET(A1,0,0,A5)", "=@OFFSET(A1,0,0,A5)", "#REF!"},
    {"=OFFSET(A1,0,0,,2)", "=@OFFSET(A1,0,0,,2)", "#VALUE!"},
    {"=OFFSET(A1,0,0,A6,1)", "=@OFFSET(A1,0,0,A6,1)", "1"},
    {"=OFFSET(A1:A2,0,0)", "=@OFFSET(A1:A2,0,0)", "#VALUE!"},
    {"=OFFSET(A1,A5,0)", "=OFFSET(A1,A5,0)", "1"},
    {"=INDEX(A1:B2,1,0)", "=@INDEX(A1:B2,1,0)", "#VALUE!"},
    {"=SUM(OFFSET(A1,0,0,2))", "=SUM(OFFSET(A1,0,0,2))", "3"},
    {"=INDEX(Rng,1)", "=INDEX(Rng,1)", "1"},
    {"=OFFSET(Rng,0,0)", "=@OFFSET(Rng,0,0)", "#VALUE!"},
    {"=INDEX((A1:A2,B1:B2),1,1,2)", "=INDEX((A1:A2,B1:B2),1,1,2)", "1"},
};

// The legacy workbook both tables describe, loaded with every
// formula unmarked, as a reader leaves a file without dynamic-array marks.
template <std::size_t N>
Workbook LoadedLegacyWorkbook(const LegacyCase (&cases)[N], bool index_offset_inputs) {
  Workbook wb = Workbook::create();
  wb.set_excel_profile(mac_365_ja_jp_profile());
  Sheet& s = wb.sheet(0);
  for (std::uint32_t r = 0; r < 3U; ++r) {
    s.set_cell_value(r, 0U, Value::number(r + 1.0));
    s.set_cell_value(r, 1U, Value::number(r + 1.0));
  }
  if (index_offset_inputs) {
    s.set_cell_value(4U, 0U, Value::number(0.0));
    s.set_cell_value(5U, 0U, Value::number(1.0));
  }
  EXPECT_TRUE(static_cast<bool>(wb.set_defined_name("Rng", "Sheet1!$A$1:$A$2")));
  EXPECT_TRUE(static_cast<bool>(wb.set_defined_name("Cel", "Sheet1!$A$1")));
  EXPECT_TRUE(static_cast<bool>(wb.set_defined_name("Val", "5")));
  EXPECT_TRUE(static_cast<bool>(wb.set_defined_name("Fn", "LAMBDA(x,x*2)")));
  for (std::uint32_t i = 0; i < N; ++i) {
    EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, i, 4U, cases[i].stored)));
    s.set_cell_dynamic_array(i, 4U, false);
  }
  wb.apply_legacy_implicit_intersections();
  return wb;
}

std::string DisplayOf(const Value& v) {
  if (v.is_number()) {
    std::ostringstream out;
    out << v.as_number();
    return out.str();
  }
  if (v.is_boolean()) {
    return v.as_boolean() ? "TRUE" : "FALSE";
  }
  if (v.is_error()) {
    return display_name(v.as_error());
  }
  if (v.is_text()) {
    return std::string(v.as_text());
  }
  return "";
}

template <std::size_t N>
void ExpectExcelsFormula2AndValues(const LegacyCase (&cases)[N], bool index_offset_inputs) {
  Workbook wb = LoadedLegacyWorkbook(cases, index_offset_inputs);
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  for (std::uint32_t i = 0; i < N; ++i) {
    const Cell* cell = wb.sheet(0).cell_at(i, 4U);
    ASSERT_NE(cell, nullptr) << cases[i].stored;
    EXPECT_EQ(cell->formula_text, cases[i].formula2) << "E" << (i + 1U);
    if (*cases[i].value != '\0') {
      EXPECT_EQ(DisplayOf(cell->cached_value), cases[i].value) << "E" << (i + 1U) << " " << cell->formula_text;
    }
  }
}

TEST(LegacyAt, LoadedLegacyFormulasReadAsExcelShowsThem) {
  ExpectExcelsFormula2AndValues(kLegacy61, false);
}

TEST(LegacyAt, LoadedLegacyIndexAndOffsetReadAsExcelShowsThem) {
  ExpectExcelsFormula2AndValues(kLegacyIndexOffset, true);
}

std::string Decimal(std::uint32_t n) {
  std::ostringstream out;
  out << n;
  return out.str();
}

// `text` escaped for an XML element body.
std::string XmlEscaped(const char* text) {
  std::string out;
  for (const char* p = text; *p != '\0'; ++p) {
    switch (*p) {
      case '&':
        out += "&amp;";
        break;
      case '<':
        out += "&lt;";
        break;
      case '>':
        out += "&gt;";
        break;
      case '"':
        out += "&quot;";
        break;
      default:
        out.push_back(*p);
    }
  }
  return out;
}

// Excel 365 (Mac, ja-JP) opening a legacy .xlsx: A1:A4 hold 1..4, Rng covers
// Sheet1!$A$1:$A$3, and the formulas fill C2 downwards in this order, stored
// as the file spells them; formula2 text and values as Excel showed them.
// An empty value is volatile.
constexpr LegacyCase kLegacyOperandsAndCalls[] = {
    {"={1,2}", "=@{1,2}", "1"},
    {"=A1:A3", "=@A1:A3", "3"},
    {"=SUM(A1:A3)", "=SUM(A1:A3)", "6"},
    {"=A1:A3+1", "=@A1:A3+1", "#VALUE!"},
    {"=INDEX(A1:A3,1)", "=INDEX(A1:A3,1)", "1"},
    {"={1,2}+1", "=@{1,2}+1", "2"},
    {"=-{1,2}", "=-@{1,2}", "-1"},
    {"=(A1:A3)", "=(@A1:A3)", "#VALUE!"},
    {"=A1", "=A1", "1"},
    {"=rng", "=@rng", "#VALUE!"},
    {"=rng+0", "=@rng+0", "#VALUE!"},
    {"=SUM(rng)", "=SUM(rng)", "6"},
    {"=OFFSET(A1,0,0,2,1)", "=@OFFSET(A1,0,0,2,1)", ""},
    {"=OFFSET(A1,0,0)", "=OFFSET(A1,0,0)", ""},
    {"=INDIRECT(\"A1:A2\")", "=@INDIRECT(\"A1:A2\")", ""},
    {"=INDIRECT(\"A1\")", "=@INDIRECT(\"A1\")", ""},
    {"=A:A", "=@A:A", "0"},
    {"=ROW(A1:A3)", "=@ROW(A1:A3)", "1"},
    {"=ROWS(A1:A3)", "=ROWS(A1:A3)", "3"},
    {"=SUMPRODUCT(A1:A3*1)", "=SUMPRODUCT(A1:A3*1)", "6"},
    {"=IF(A1:A3>1,1,0)", "=IF(@A1:A3>1,1,0)", "#VALUE!"},
    {"=A1:A3=1", "=@A1:A3=1", "#VALUE!"},
    {"=VLOOKUP(2,A1:A3,1,0)", "=VLOOKUP(2,A1:A3,1,0)", "2"},
    {"=MATCH(2,A1:A3,0)", "=MATCH(2,A1:A3,0)", "2"},
    {"=SUM(A1:A3*2)", "=SUM(@A1:A3*2)", "#VALUE!"},
    {"=INDEX(A1:A3,0)", "=@INDEX(A1:A3,0)", "#VALUE!"},
    {"=INDEX(A1:A3,0,1)", "=@INDEX(A1:A3,0,1)", "#VALUE!"},
    {"=ABS(A1:A3)", "=ABS(@A1:A3)", "#VALUE!"},
    {"=LEN(A1:A3)", "=LEN(@A1:A3)", "#VALUE!"},
    {"=N(A1:A3)", "=N(A1:A3)", "1"},
    {"=TRANSPOSE(A1:A3)", "=@TRANSPOSE(@A1:A3)", "#VALUE!"},
    {"=A1:A3 A2:A4", "=@A1:A3 A2:A4", "#VALUE!"},
    {"=SUM((A1,A2))", "=SUM((A1,A2))", "3"},
    {"=(A1,A2)", "=(A1,A2)", "#VALUE!"},
    {"=1+1", "=1+1", "2"},
    {"=\"x\"", "=\"x\"", "x"},
    {"=ISREF(A1:A3)", "=ISREF(A1:A3)", "TRUE"},
    {"=COLUMN(A1:B1)", "=@COLUMN(A1:B1)", "1"},
    {"=A1:A3&\"\"", "=@A1:A3&\"\"", "#VALUE!"},
    {"=MAX(A1:A3)+A1:A3", "=MAX(A1:A3)+@A1:A3", "#VALUE!"},
    {"=_xlfn.SEQUENCE(3)", "=@SEQUENCE(3)", "1"},
    {"=_xlfn.UNIQUE(A1:A3)", "=@UNIQUE(A1:A3)", "1"},
    {"=_xlfn.XLOOKUP(2,A1:A3,A1:A3)", "=XLOOKUP(2,A1:A3,A1:A3)", "2"},
    {"=_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)(2)", "=@LAMBDA(x,@x+1)(2)", "3"},
    {"=_xlfn.LET(_xlpm.a,A1:A3,_xlpm.a)", "=@LET(a,A1:A3,a)", "#VALUE!"},
    {"=_xlfn.SINGLE(A1:A3)", "=@A1:A3", "#VALUE!"},
    {"=IFERROR(A1:A3,0)", "=IFERROR(@A1:A3,0)", "0"},
    {"=CHOOSE(1,A1:A3)", "=@CHOOSE(1,A1:A3)", "#VALUE!"},
    {"=INDEX(rng,2)", "=INDEX(rng,2)", "2"},
};

constexpr std::uint32_t kOperandsAndCallsFirstRow = 1U;

TEST(LegacyAt, LoadedLegacyOperandsAndCallsReadAsExcelShowsThem) {
  Workbook wb = Workbook::create();
  wb.set_excel_profile(mac_365_ja_jp_profile());
  Sheet& s = wb.sheet(0);
  for (std::uint32_t r = 0; r < 4U; ++r) {
    s.set_cell_value(r, 0U, Value::number(r + 1.0));
  }
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Rng", "Sheet1!$A$1:$A$3")));
  constexpr std::size_t n = sizeof(kLegacyOperandsAndCalls) / sizeof(kLegacyOperandsAndCalls[0]);
  for (std::uint32_t i = 0; i < n; ++i) {
    const std::uint32_t row = kOperandsAndCallsFirstRow + i;
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, row, 2U, kLegacyOperandsAndCalls[i].stored)));
    s.set_cell_dynamic_array(row, 2U, false);
  }
  wb.apply_legacy_implicit_intersections();
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  for (std::uint32_t i = 0; i < n; ++i) {
    const LegacyCase& c = kLegacyOperandsAndCalls[i];
    const Cell* cell = s.cell_at(kOperandsAndCallsFirstRow + i, 2U);
    ASSERT_NE(cell, nullptr) << c.stored;
    EXPECT_EQ(cell->formula_text, c.formula2) << c.stored;
    if (*c.value != '\0') {
      EXPECT_EQ(DisplayOf(cell->cached_value), c.value) << c.stored << " " << cell->formula_text;
    }
  }
}

// The same formulas read from an .xlsx without dynamic-array marks show
// Excel's '@', and are saved back as the file spelled them, an authored
// `_xlfn.SINGLE` excepted, which Excel saves as its operand. A CSE block
// shows none.
TEST(LegacyAt, LegacyXlsxShowsImpliedAtAndSavesTheStoredText) {
  std::string sheet_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>";
  std::vector<std::uint32_t> rows;
  constexpr std::size_t n = sizeof(kLegacyOperandsAndCalls) / sizeof(kLegacyOperandsAndCalls[0]);
  for (std::uint32_t i = 0; i < n; ++i) {
    // The fixture package carries no defined names.
    if (std::string_view(kLegacyOperandsAndCalls[i].stored).find("rng") != std::string_view::npos) {
      continue;
    }
    rows.push_back(i);
  }
  for (std::uint32_t r = 0; r < kOperandsAndCallsFirstRow + n + 1U; ++r) {
    const std::string ref = Decimal(r + 1U);
    sheet_xml += "<row r=\"" + ref + "\">";
    if (r < 4U) {
      sheet_xml += "<c r=\"A" + ref + "\"><v>" + ref + "</v></c>";
    }
    for (const std::uint32_t i : rows) {
      if (kOperandsAndCallsFirstRow + i == r) {
        sheet_xml += "<c r=\"C" + ref + "\"><f>" + XmlEscaped(kLegacyOperandsAndCalls[i].stored + 1) + "</f></c>";
      }
    }
    if (r == kOperandsAndCallsFirstRow + n) {
      sheet_xml += "<c r=\"C" + ref + "\"><f t=\"array\" ref=\"C" + ref + "\">A1:A3+1</f></c>";
    }
    sheet_xml += "</row>";
  }
  sheet_xml += "</sheetData></worksheet>";
  const std::vector<std::uint8_t> package =
      test::BuildOoxmlPackage(test::OoxmlContentTypes(test::kXlsxWorkbookContentType), sheet_xml);
  auto loaded = io::read_ooxml(SpanOf(package));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  const Sheet& s = loaded.value().workbook.sheet(0);
  for (const std::uint32_t i : rows) {
    const Cell* cell = s.cell_at(kOperandsAndCallsFirstRow + i, 2U);
    ASSERT_NE(cell, nullptr) << kLegacyOperandsAndCalls[i].stored;
    EXPECT_EQ(cell->formula_text, kLegacyOperandsAndCalls[i].formula2) << kLegacyOperandsAndCalls[i].stored;
  }
  const Cell* cse = s.cell_at(kOperandsAndCallsFirstRow + n, 2U);
  ASSERT_NE(cse, nullptr);
  EXPECT_EQ(cse->formula_text, "=A1:A3+1");

  auto saved = loaded.value().workbook.save();
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(saved.value()))));
  auto saved_sheet = zip.read_entry("xl/worksheets/sheet1.xml");
  ASSERT_TRUE(static_cast<bool>(saved_sheet));
  const std::string saved_xml(saved_sheet.value().begin(), saved_sheet.value().end());
  for (const std::uint32_t i : rows) {
    std::string stored = kLegacyOperandsAndCalls[i].stored + 1;
    if (stored == "_xlfn.SINGLE(A1:A3)") {
      stored = "A1:A3";
    }
    EXPECT_NE(saved_xml.find(">" + XmlEscaped(stored.c_str()) + "</f>"), std::string::npos) << stored;
  }
  EXPECT_EQ(saved_xml.find("_xlfn.SINGLE"), std::string::npos) << saved_xml;
  EXPECT_EQ(saved_xml.find(" cm=\""), std::string::npos) << saved_xml;
}

// '@' over a spill or over a LAMBDA or LET returning a range intersects that
// range with the formula cell, as over a static one (Excel 365: G1 =
// SEQUENCE(3), Fn = LAMBDA(x,x); 2 on row 2, #VALUE! on row 5). A legacy cell
// holding a spill reference reads with Excel's implied '@'.
TEST(LegacyAt, SpillAndLambdaResultsProjectOntoTheFormulaCell) {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  for (std::uint32_t r = 0; r < 3U; ++r) {
    s.set_cell_value(r, 0U, Value::number(r + 1.0));
  }
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Fn", "LAMBDA(x,x)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 6U, "=SEQUENCE(3)")));
  for (const std::uint32_t row : {1U, 4U}) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, row, 7U, "=G1#")));
    s.set_cell_dynamic_array(row, 7U, false);
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, row, 8U, "=@Fn(A1:A3)")));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, row, 9U, "=@LAMBDA(x,x)(A1:A3)")));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, row, 10U, "=@LET(a,A1:A3,a)")));
  }
  wb.apply_legacy_implicit_intersections();
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_EQ(s.cell_at(1U, 7U)->formula_text, "=@G1#");
  for (const std::uint32_t col : {7U, 8U, 9U, 10U}) {
    EXPECT_EQ(DisplayOf(s.cell_at(1U, col)->cached_value), "2") << col;
    EXPECT_EQ(DisplayOf(s.cell_at(4U, col)->cached_value), "#VALUE!") << col;
  }
}

// Node ranges count UTF-16 units; the `@` still lands at the right byte after
// multi-byte text, including a supplementary-plane character.
TEST(LegacyAt, ImpliedAtLandsAfterNonAsciiText) {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  const char* const stored[] = {"=LEN(\"日本\")+ABS(A1:A2)", "=LEN(\"😀\")+ABS(A1:A2)", "=\"日本\"&LEN((A1:A2))"};
  const char* const shown[] = {"=LEN(\"日本\")+ABS(@A1:A2)", "=LEN(\"😀\")+ABS(@A1:A2)", "=\"日本\"&LEN((@A1:A2))"};
  for (std::uint32_t i = 0; i < 3U; ++i) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, i, 4U, stored[i])));
    s.set_cell_dynamic_array(i, 4U, false);
  }
  wb.apply_legacy_implicit_intersections();
  for (std::uint32_t i = 0; i < 3U; ++i) {
    EXPECT_EQ(s.cell_at(i, 4U)->formula_text, shown[i]) << stored[i];
  }
}

TEST(LegacyAt, DynamicArrayAndCseFormulasKeepTheirText) {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 4U, "=A1:A2")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 4U, "=A1:A2*2")));
  s.set_cell_dynamic_array(1U, 4U, false);
  s.set_cell_array_block(1U, 4U, 2U, 1U);
  wb.apply_legacy_implicit_intersections();
  EXPECT_EQ(s.cell_at(0U, 4U)->formula_text, "=A1:A2");
  EXPECT_EQ(s.cell_at(1U, 4U)->formula_text, "=A1:A2*2");
}

// A legacy array formula (CSE) fills exactly its own block: one cell takes
// the result's top-left element (Excel 365: 1, 2, 12, 1 for C51:C54, with
// FORMULATEXT in braces), a larger block repeats a one-cell axis and fills
// what the result cannot supply with #N/A. It is saved with its own extent.
TEST(LegacyAt, LegacyArrayFormulaFillsItsOwnBlock) {
  std::string sheet_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>";
  for (std::uint32_t r = 1; r <= 4U; ++r) {
    const std::string ref = Decimal(r);
    sheet_xml += "<row r=\"" + ref + "\"><c r=\"A" + ref + "\"><v>" + ref + "</v></c>";
    if (r == 1U) {
      sheet_xml +=
          "<c r=\"E1\"><f t=\"array\" ref=\"E1:E4\">ROW(A1:A3)</f></c>"
          "<c r=\"G1\"><f t=\"array\" ref=\"G1:H2\">{1,2}</f></c>"
          "<c r=\"J1\"><f t=\"array\" ref=\"J1:J2\">1+1</f></c>";
    }
    sheet_xml += "</row>";
  }
  const char* const cse[] = {"{1,2}", "A1:A3+1", "SUM(A1:A3*2)", "ROW(A1:A3)"};
  for (std::uint32_t i = 0; i < 4U; ++i) {
    const std::string ref = Decimal(51U + i);
    sheet_xml += "<row r=\"" + ref + "\"><c r=\"C" + ref + "\"><f t=\"array\" ref=\"C" + ref + "\">" + cse[i] +
                 "</f></c><c r=\"D" + ref + "\"><f>_xlfn.FORMULATEXT(C" + ref + ")</f></c></row>";
  }
  sheet_xml += "</sheetData></worksheet>";
  const std::vector<std::uint8_t> package =
      test::BuildOoxmlPackage(test::OoxmlContentTypes(test::kXlsxWorkbookContentType), sheet_xml);
  auto loaded = io::read_ooxml(SpanOf(package));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  Workbook& wb = loaded.value().workbook;
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  const Sheet& s = wb.sheet(0);
  const char* const values[] = {"1", "2", "12", "1"};
  const char* const texts[] = {"{={1,2}}", "{=A1:A3+1}", "{=SUM(A1:A3*2)}", "{=ROW(A1:A3)}"};
  for (std::uint32_t i = 0; i < 4U; ++i) {
    EXPECT_EQ(s.cell_at(50U + i, 2U)->formula_text, std::string("=") + cse[i]);
    EXPECT_EQ(DisplayOf(s.resolve_cell_value(50U + i, 2U)), values[i]) << cse[i];
    EXPECT_EQ(DisplayOf(s.resolve_cell_value(50U + i, 3U)), texts[i]) << cse[i];
  }
  const char* const column_e[] = {"1", "2", "3", "#N/A"};
  for (std::uint32_t r = 0; r < 4U; ++r) {
    EXPECT_EQ(DisplayOf(s.resolve_cell_value(r, 4U)), column_e[r]) << r;
  }
  EXPECT_EQ(DisplayOf(s.resolve_cell_value(4U, 4U)), "");
  for (std::uint32_t r = 0; r < 2U; ++r) {
    EXPECT_EQ(DisplayOf(s.resolve_cell_value(r, 6U)), "1") << r;
    EXPECT_EQ(DisplayOf(s.resolve_cell_value(r, 7U)), "2") << r;
    EXPECT_EQ(DisplayOf(s.resolve_cell_value(r, 9U)), "2") << r;
  }

  auto saved = wb.save();
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(saved.value()))));
  auto saved_sheet = zip.read_entry("xl/worksheets/sheet1.xml");
  ASSERT_TRUE(static_cast<bool>(saved_sheet));
  const std::string saved_xml(saved_sheet.value().begin(), saved_sheet.value().end());
  for (std::uint32_t i = 0; i < 4U; ++i) {
    const std::string ref = Decimal(51U + i);
    EXPECT_NE(saved_xml.find("<f t=\"array\" ref=\"C" + ref + "\">" + XmlEscaped(cse[i]) + "</f>"), std::string::npos)
        << saved_xml;
  }
  EXPECT_NE(saved_xml.find("<f t=\"array\" ref=\"E1:E4\">ROW(A1:A3)</f>"), std::string::npos) << saved_xml;
  EXPECT_NE(saved_xml.find("<f t=\"array\" ref=\"G1:H2\">{1,2}</f>"), std::string::npos) << saved_xml;
  EXPECT_EQ(saved_xml.find(" cm=\""), std::string::npos) << saved_xml;

  auto reread = io::read_ooxml(SpanOf(saved.value()));
  ASSERT_TRUE(static_cast<bool>(reread)) << reread.error().message;
  auto xlsb = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(xlsb)) << xlsb.error().message;
  auto xlsb_read = io::xlsb::read_xlsb(SpanOf(xlsb.value()));
  ASSERT_TRUE(static_cast<bool>(xlsb_read)) << xlsb_read.error().message;
  for (const Workbook* round_tripped : {&reread.value().workbook, &xlsb_read.value().workbook}) {
    const Sheet& rs = round_tripped->sheet(0);
    for (const auto& [row, col, rows, cols] :
         {std::array<std::uint32_t, 4>{0U, 4U, 4U, 1U}, {0U, 6U, 2U, 2U}, {0U, 9U, 2U, 1U}, {53U, 2U, 1U, 1U}}) {
      const Cell* block = rs.cell_at(row, col);
      ASSERT_NE(block, nullptr) << row << "," << col;
      EXPECT_EQ(block->array_block_rows, rows) << row << "," << col;
      EXPECT_EQ(block->array_block_cols, cols) << row << "," << col;
      EXPECT_FALSE(block->dynamic_array) << row << "," << col;
    }
  }
}

TEST(LegacyAt, LegacyImpliedAtIsNotStored) {
  // Excel stores the '@' it would add to a legacy formula as nothing, and a
  // written one elsewhere as `_xlfn.SINGLE`.
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 4U, "=SUM(@A1:A2*2)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 4U, "=SUM(@A1:A2)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Rng", "Sheet1!$A$1:$A$2")));
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Cel", "Sheet1!$A$1")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 2U, 4U, "=@Rng")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 3U, 4U, "=@Cel")));
  for (std::uint32_t row = 0; row < 4U; ++row) {
    s.set_cell_dynamic_array(row, 4U, false);
  }

  auto bytes_or = wb.save();
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message;
  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto sheet_or = zip.read_entry("xl/worksheets/sheet1.xml");
  ASSERT_TRUE(static_cast<bool>(sheet_or));
  const std::string sheet_xml(sheet_or.value().begin(), sheet_or.value().end());
  EXPECT_NE(sheet_xml.find("<f>SUM(A1:A2*2)</f>"), std::string::npos) << sheet_xml;
  EXPECT_NE(sheet_xml.find("<f>SUM(_xlfn.SINGLE(A1:A2))</f>"), std::string::npos) << sheet_xml;
  EXPECT_NE(sheet_xml.find("<f>Rng</f>"), std::string::npos) << sheet_xml;
  EXPECT_NE(sheet_xml.find("<f>_xlfn.SINGLE(Cel)</f>"), std::string::npos) << sheet_xml;

  auto result_or = io::read_ooxml(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  const char* const formulas[] = {"=SUM(@A1:A2*2)", "=SUM(@A1:A2)", "=@Rng", "=@Cel"};
  for (std::uint32_t row = 0; row < 4U; ++row) {
    EXPECT_EQ(result_or.value().workbook.sheet(0).cell_at(row, 4U)->formula_text, formulas[row]);
  }
}

// Every cell of a legacy array formula block reports the block's formula
// (FORMULATEXT in braces, ISFORMULA TRUE), and a write into a non-anchor
// cell is refused, leaving the block unchanged. A write to the anchor
// replaces the whole block.
TEST(LegacyAt, LegacyArrayFormulaBlockIsChangedOnlyAsAWhole) {
  const std::string sheet_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>"
      "<row r=\"1\"><c r=\"A1\"><v>5</v></c><c r=\"B1\"><f t=\"array\" ref=\"B1:B4\">ROW(A1:A3)</f></c>"
      "<c r=\"D1\"><f t=\"array\" ref=\"D1\">1+1</f></c>"
      "<c r=\"F1\"><f>_xlfn.FORMULATEXT(B1)</f></c><c r=\"G1\"><f>_xlfn.ISFORMULA(B1)</f></c></row>"
      "<row r=\"2\"><c r=\"A2\"><v>7</v></c>"
      "<c r=\"F2\"><f>_xlfn.FORMULATEXT(B2)</f></c><c r=\"G2\"><f>_xlfn.ISFORMULA(B2)</f></c></row>"
      "<row r=\"3\"><c r=\"F3\"><f>_xlfn.FORMULATEXT(B3)</f></c><c r=\"G3\"><f>_xlfn.ISFORMULA(B3)</f></c></row>"
      "<row r=\"4\"><c r=\"F4\"><f>_xlfn.FORMULATEXT(B4)</f></c><c r=\"G4\"><f>_xlfn.ISFORMULA(B4)</f></c>"
      "</row>"
      "<row r=\"5\"><c r=\"F5\"><f>_xlfn.FORMULATEXT(B5)</f></c><c r=\"G5\"><f>_xlfn.ISFORMULA(B5)</f></c>"
      "</row>"
      "</sheetData></worksheet>";
  const std::vector<std::uint8_t> package =
      test::BuildOoxmlPackage(test::OoxmlContentTypes(test::kXlsxWorkbookContentType), sheet_xml);
  auto loaded = io::read_ooxml(SpanOf(package));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  Workbook& wb = loaded.value().workbook;
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  const Sheet& s = wb.sheet(0);
  const char* const block[] = {"1", "2", "3", "#N/A"};
  const auto expect_block = [&]() {
    for (std::uint32_t r = 0; r < 4U; ++r) {
      EXPECT_EQ(DisplayOf(s.resolve_cell_value(r, 1U)), block[r]) << r;
      EXPECT_EQ(DisplayOf(s.resolve_cell_value(r, 5U)), "{=ROW(A1:A3)}") << r;
      EXPECT_EQ(DisplayOf(s.resolve_cell_value(r, 6U)), "TRUE") << r;
    }
    EXPECT_EQ(DisplayOf(s.resolve_cell_value(4U, 5U)), "#N/A");
    EXPECT_EQ(DisplayOf(s.resolve_cell_value(4U, 6U)), "FALSE");
    const Cell* anchor = s.cell_at(0U, 1U);
    ASSERT_NE(anchor, nullptr);
    EXPECT_EQ(anchor->formula_text, "=ROW(A1:A3)");
    EXPECT_EQ(anchor->array_block_rows, 4U);
    EXPECT_EQ(anchor->array_block_cols, 1U);
  };
  expect_block();

  const auto refused = [](const Expected<void, Error>& r) {
    return !r && r.error().code == FormulonErrorCode::kPreconditionFailed;
  };
  EXPECT_TRUE(refused(wb.set_cell_formula(0U, 1U, 1U, "=10")));
  EXPECT_TRUE(refused(wb.set_cell_value(0U, 2U, 1U, Value::number(9.0))));
  EXPECT_TRUE(refused(wb.set_cell_value(0U, 3U, 1U, Value::blank())));
  EXPECT_TRUE(refused(wb.set_cell_text(0U, 1U, 1U, "x")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  expect_block();

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 1U, Value::blank())));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  for (std::uint32_t r = 0; r < 4U; ++r) {
    EXPECT_TRUE(s.resolve_cell_value(r, 1U).is_blank()) << r;
    EXPECT_EQ(DisplayOf(s.resolve_cell_value(r, 6U)), "FALSE") << r;
  }

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 3U, Value::number(4.0))));
  const Cell* replaced = s.cell_at(0U, 3U);
  ASSERT_NE(replaced, nullptr);
  EXPECT_TRUE(replaced->formula_text.empty());
  EXPECT_EQ(replaced->array_block_rows, 0U);
}

}  // namespace
}  // namespace formulon
