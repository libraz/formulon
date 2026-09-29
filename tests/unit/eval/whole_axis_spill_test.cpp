//
// Unit tests for the bare whole-axis reference as a spilling expression
// (`=A:A`, `=A:C`, `=1:2`, `=3:3` standing as the entire formula).
//
// Excel 365 gives such a reference the array of the *declared* rectangle —
// rows 1..1048576 for a column span, columns A..XFD for a row span — and
// spills it from the formula cell. Nothing is trimmed to the populated
// extent, which is what makes three consequences observable and is what
// these tests pin:
//
//   * geometry: the rectangle must fit measured from the formula cell, so a
//     whole-column reference only spills from row 1 and a whole-row
//     reference only from column A; anywhere else is `#SPILL!`;
//   * footprint: an occupied cell anywhere inside the declared rectangle
//     blocks the spill, including far outside the populated extent of the
//     source;
//   * self-reference: a formula inside the axis it references is circular.
//     Excel leaves 0 with iterative calculation off; Formulon resolves a
//     cycle to `#REF!` engine-wide, so that is what these tests assert.
//
// The recalc path is used rather than ad-hoc evaluation because only a
// committed spill exposes the footprint and the phantom cells.
//
// Observed Excel behaviour: tests/oracle/cases/whole_axis_spill.yaml.

#include <cstdint>
#include <string>

#include "eval/adhoc_eval.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

// Seeds `wb`'s first sheet with `formula` at (`row`, `col`) and recalcs.
// Returns the sheet so callers can read the anchor and the spilled cells.
// `formula` is a `const char*` so the call creates no temporary: the returned
// reference points into `wb`, but any temporary argument makes GCC read the
// result as possibly dangling (-Wdangling-reference).
Sheet& RecalcWith(Workbook& wb, std::uint32_t row, std::uint32_t col, const char* formula) {
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, row, col, formula)));
  EXPECT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
  return wb.sheet(0);
}

void SeedNumber(Workbook& wb, std::uint32_t row, std::uint32_t col, double value) {
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, row, col, Value::number(value))));
}

// Evaluates the formula already stored at (`row`, `col`) through the
// read-only driver, which is the shape the oracle harness and every C ABI
// ad-hoc caller use. It differs from recalc in one way that matters here:
// the dependency graph is not consulted, so a cycle has to be recognised by
// the evaluator itself rather than by the engine's SCC pass. Assertions that
// only drive recalc cannot see the difference.
Value AdhocValue(const Workbook& wb, const Sheet& sheet, std::uint32_t row, std::uint32_t col,
                 const std::string& formula) {
  Arena arena;
  return evaluate_formula_text_array(wb, sheet, row, col, formula, arena, default_registry());
}

// Same driver, but also reports what the evaluation allocated. Every cell a
// whole-axis result materialises passes through this arena, so the byte
// count is a direct, non-timing-based witness of whether the rectangle was
// built. A refused footprint must decide without building it.
Value AdhocValueAndBytes(const Workbook& wb, const Sheet& sheet, std::uint32_t row, std::uint32_t col,
                         const std::string& formula, std::size_t* out_bytes) {
  Arena arena;
  const Value v = evaluate_formula_text_array(wb, sheet, row, col, formula, arena, default_registry());
  *out_bytes = arena.bytes_used();
  return v;
}

// One full grid column of `Value`s. A materialised whole-column rectangle
// cannot come in under this; a refusal that skips materialising stays orders
// of magnitude below it.
constexpr std::size_t kWholeColumnBytes = static_cast<std::size_t>(Sheet::kMaxRows) * sizeof(Value);

// ---------------------------------------------------------------------------
// Geometry: the declared rectangle has to fit from the formula cell.
// ---------------------------------------------------------------------------

TEST(WholeAxisSpill, WholeColumnBelowRowOneOverflows) {
  // `=A:A` at Z2: the 1048576-row rectangle no longer fits below row 2.
  Workbook wb = Workbook::create();
  SeedNumber(wb, 0U, 0U, 1.0);
  SeedNumber(wb, 1U, 0U, 2.0);
  SeedNumber(wb, 2U, 0U, 3.0);
  const Sheet& sheet = RecalcWith(wb, 1U, 25U, "=A:A");
  const Value v = sheet.resolve_cell_value(1U, 25U);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Spill);
  EXPECT_EQ(sheet.spill_region_at_anchor(1U, 25U), nullptr);
}

TEST(WholeAxisSpill, MultiColumnSpanBelowRowOneOverflows) {
  // `=A:C` at Z2: the same overflow with a three-column span.
  Workbook wb = Workbook::create();
  SeedNumber(wb, 0U, 0U, 1.0);
  SeedNumber(wb, 1U, 0U, 2.0);
  SeedNumber(wb, 0U, 1U, 10.0);
  SeedNumber(wb, 0U, 2U, 100.0);
  const Sheet& sheet = RecalcWith(wb, 1U, 25U, "=A:C");
  const Value v = sheet.resolve_cell_value(1U, 25U);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Spill);
}

TEST(WholeAxisSpill, WholeRowRightOfColumnAOverflows) {
  // `=1:2` at B5: the 16384-column rectangle no longer fits right of B.
  Workbook wb = Workbook::create();
  SeedNumber(wb, 0U, 0U, 1.0);
  SeedNumber(wb, 1U, 0U, 2.0);
  SeedNumber(wb, 0U, 1U, 10.0);
  SeedNumber(wb, 0U, 2U, 100.0);
  const Sheet& sheet = RecalcWith(wb, 4U, 1U, "=1:2");
  const Value v = sheet.resolve_cell_value(4U, 1U);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Spill);
}

// ---------------------------------------------------------------------------
// Footprint: a blocker anywhere in the declared rectangle stops the spill,
// including far outside the populated extent of the source.
// ---------------------------------------------------------------------------

TEST(WholeAxisSpill, BlockerFarBelowPopulatedExtentStopsColumnSpill) {
  // `=A:A` at Z1 with Z500 occupied. Column A holds three cells, so a
  // used-range footprint would clear row 500 and let the spill through;
  // the declared rectangle reaches it and Excel reports `#SPILL!`.
  Workbook wb = Workbook::create();
  SeedNumber(wb, 0U, 0U, 1.0);
  SeedNumber(wb, 1U, 0U, 2.0);
  SeedNumber(wb, 2U, 0U, 3.0);
  SeedNumber(wb, 499U, 25U, 7.0);
  const Sheet& sheet = RecalcWith(wb, 0U, 25U, "=A:A");
  const Value v = sheet.resolve_cell_value(0U, 25U);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Spill);
  EXPECT_EQ(sheet.spill_region_at_anchor(0U, 25U), nullptr);
  // The blocker keeps its own value; a refused spill writes nothing.
  EXPECT_EQ(sheet.resolve_cell_value(499U, 25U).as_number(), 7.0);
}

TEST(WholeAxisSpill, BlockerRightOfPopulatedExtentStopsRowSpill) {
  // `=3:3` at A5 with Z5 occupied. Row 3 holds two cells; column Z lies
  // inside the full-width footprint.
  Workbook wb = Workbook::create();
  SeedNumber(wb, 2U, 0U, 3.0);
  SeedNumber(wb, 2U, 1U, 30.0);
  SeedNumber(wb, 4U, 25U, 9.0);
  const Sheet& sheet = RecalcWith(wb, 4U, 0U, "=3:3");
  const Value v = sheet.resolve_cell_value(4U, 0U);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Spill);
  EXPECT_EQ(sheet.resolve_cell_value(4U, 25U).as_number(), 9.0);
}

// ---------------------------------------------------------------------------
// Self-reference: the formula cell lies inside the referenced axis. Both
// cases are laid out so the rectangle fits and the runway is clear, leaving
// circularity as the only thing that can stop the spill.
// ---------------------------------------------------------------------------

TEST(WholeAxisSpill, ColumnContainingItsOwnFormulaCellIsCircular) {
  // `=A:A` at A1, column A otherwise empty. Excel leaves 0 with iterative
  // calculation off; Formulon reports every cycle member as `#REF!`.
  //
  // Both drivers are asserted on purpose. Recalc alone proves nothing about
  // the evaluator: with a clear runway the rectangle commits, the committed
  // footprint then feeds a self-edge back into the graph, and the engine's
  // SCC pass reaches `#REF!` on its own — after materialising 1,048,576
  // cells. The read-only driver has no graph to fall back on, so it is the
  // one that shows the evaluator recognised the cycle.
  Workbook wb = Workbook::create();
  SeedNumber(wb, 0U, 1U, 10.0);
  SeedNumber(wb, 0U, 2U, 100.0);
  const Sheet& sheet = RecalcWith(wb, 0U, 0U, "=A:A");
  const Value v = sheet.resolve_cell_value(0U, 0U);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
  EXPECT_EQ(sheet.spill_region_at_anchor(0U, 0U), nullptr);

  const Value adhoc = AdhocValue(wb, sheet, 0U, 0U, "=A:A");
  ASSERT_TRUE(adhoc.is_error());
  EXPECT_EQ(adhoc.as_error(), ErrorCode::Ref);
}

TEST(WholeAxisSpill, RowContainingItsOwnFormulaCellIsCircular) {
  // `=5:5` at A5, row 5 otherwise empty. Same two drivers, same reason.
  Workbook wb = Workbook::create();
  SeedNumber(wb, 0U, 0U, 1.0);
  SeedNumber(wb, 0U, 1U, 10.0);
  const Sheet& sheet = RecalcWith(wb, 4U, 0U, "=5:5");
  const Value v = sheet.resolve_cell_value(4U, 0U);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);

  const Value adhoc = AdhocValue(wb, sheet, 4U, 0U, "=5:5");
  ASSERT_TRUE(adhoc.is_error());
  EXPECT_EQ(adhoc.as_error(), ErrorCode::Ref);
}

TEST(WholeAxisSpill, CircularityIsDecidedBeforeTheFootprint) {
  // `=A:A` at A1 with A2 and A3 occupied is circular *and* blocked at once.
  // Circularity wins: the formula reads its own result, so it never spills
  // and the footprint is never consulted. The two clear-runway cases above
  // cannot show this — they remove the blocker precisely so that only the
  // cycle can fire, which makes them blind to the ordering.
  Workbook wb = Workbook::create();
  SeedNumber(wb, 1U, 0U, 2.0);
  SeedNumber(wb, 2U, 0U, 3.0);
  const Sheet& sheet = RecalcWith(wb, 0U, 0U, "=A:A");
  const Value v = sheet.resolve_cell_value(0U, 0U);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(WholeAxisSpill, BlockedButNotCircularStillReportsSpill) {
  // The control for the case above, identical except that the referenced
  // axis does not contain the formula cell. Blocked and not circular, so
  // the footprint decides and the answer is `#SPILL!`. The pair is what
  // localises an inverted ordering: this one passes either way, the one
  // above only passes when circularity is settled first.
  Workbook wb = Workbook::create();
  SeedNumber(wb, 1U, 0U, 2.0);
  SeedNumber(wb, 2U, 0U, 3.0);
  const Sheet& sheet = RecalcWith(wb, 0U, 0U, "=B:B");
  const Value v = sheet.resolve_cell_value(0U, 0U);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Spill);
}

// ---------------------------------------------------------------------------
// The full-rectangle spills. The oracle harness cannot capture a 16384-cell
// spill, so its goldens carry the shape plus the cells Excel was sampled at;
// these tests assert exactly those samples.
// ---------------------------------------------------------------------------

TEST(WholeAxisSpill, SingleRowSpillsFullGridWidth) {
  // `=3:3` at A5 spills 1x16384: A5=3, B5=30 from row 3, C5 and XFD5 both
  // 0 for the unpopulated cells, A6 outside the footprint and so blank.
  Workbook wb = Workbook::create();
  SeedNumber(wb, 2U, 0U, 3.0);
  SeedNumber(wb, 2U, 1U, 30.0);
  const Sheet& sheet = RecalcWith(wb, 4U, 0U, "=3:3");
  const SpillRegion* region = sheet.spill_region_at_anchor(4U, 0U);
  ASSERT_NE(region, nullptr);
  EXPECT_EQ(region->rows, 1U);
  EXPECT_EQ(region->cols, Sheet::kMaxCols);
  EXPECT_EQ(sheet.resolve_cell_value(4U, 0U).as_number(), 3.0);
  EXPECT_EQ(sheet.resolve_cell_value(4U, 1U).as_number(), 30.0);
  EXPECT_EQ(sheet.resolve_cell_value(4U, 2U).as_number(), 0.0);
  EXPECT_EQ(sheet.resolve_cell_value(4U, Sheet::kMaxCols - 1U).as_number(), 0.0);
  EXPECT_TRUE(sheet.resolve_cell_value(5U, 0U).is_blank());
}

TEST(WholeAxisSpill, RowPairSpillsFullGridWidth) {
  // `=1:2` at A5 spills 2x16384 across both referenced rows.
  Workbook wb = Workbook::create();
  SeedNumber(wb, 0U, 0U, 1.0);
  SeedNumber(wb, 1U, 0U, 2.0);
  SeedNumber(wb, 0U, 1U, 10.0);
  SeedNumber(wb, 0U, 2U, 100.0);
  const Sheet& sheet = RecalcWith(wb, 4U, 0U, "=1:2");
  const SpillRegion* region = sheet.spill_region_at_anchor(4U, 0U);
  ASSERT_NE(region, nullptr);
  EXPECT_EQ(region->rows, 2U);
  EXPECT_EQ(region->cols, Sheet::kMaxCols);
  EXPECT_EQ(sheet.resolve_cell_value(4U, 0U).as_number(), 1.0);
  EXPECT_EQ(sheet.resolve_cell_value(4U, 1U).as_number(), 10.0);
  EXPECT_EQ(sheet.resolve_cell_value(4U, 2U).as_number(), 100.0);
  EXPECT_EQ(sheet.resolve_cell_value(5U, 0U).as_number(), 2.0);
  EXPECT_EQ(sheet.resolve_cell_value(5U, 1U).as_number(), 0.0);
  EXPECT_EQ(sheet.resolve_cell_value(4U, Sheet::kMaxCols - 1U).as_number(), 0.0);
  EXPECT_EQ(sheet.resolve_cell_value(5U, Sheet::kMaxCols - 1U).as_number(), 0.0);
}

TEST(WholeAxisSpill, SingleColumnSpillsFullGridHeight) {
  // `=A:A` at Z1 spills 1048576x1: the source cells, then 0 all the way to
  // the last row. AA1 sits outside the one-column footprint.
  Workbook wb = Workbook::create();
  SeedNumber(wb, 0U, 0U, 1.0);
  SeedNumber(wb, 1U, 0U, 2.0);
  SeedNumber(wb, 2U, 0U, 3.0);
  const Sheet& sheet = RecalcWith(wb, 0U, 25U, "=A:A");
  const SpillRegion* region = sheet.spill_region_at_anchor(0U, 25U);
  ASSERT_NE(region, nullptr);
  EXPECT_EQ(region->rows, Sheet::kMaxRows);
  EXPECT_EQ(region->cols, 1U);
  EXPECT_EQ(sheet.resolve_cell_value(0U, 25U).as_number(), 1.0);
  EXPECT_EQ(sheet.resolve_cell_value(1U, 25U).as_number(), 2.0);
  EXPECT_EQ(sheet.resolve_cell_value(2U, 25U).as_number(), 3.0);
  EXPECT_EQ(sheet.resolve_cell_value(3U, 25U).as_number(), 0.0);
  EXPECT_EQ(sheet.resolve_cell_value(Sheet::kMaxRows - 1U, 25U).as_number(), 0.0);
  EXPECT_TRUE(sheet.resolve_cell_value(0U, 26U).is_blank());
}

TEST(WholeAxisSpill, MultiColumnSpanSpillsFullGridHeight) {
  // `=A:C` at Z1 spills 1048576x3, carrying the column offsets across.
  Workbook wb = Workbook::create();
  SeedNumber(wb, 0U, 0U, 1.0);
  SeedNumber(wb, 1U, 0U, 2.0);
  SeedNumber(wb, 0U, 1U, 10.0);
  SeedNumber(wb, 0U, 2U, 100.0);
  const Sheet& sheet = RecalcWith(wb, 0U, 25U, "=A:C");
  const SpillRegion* region = sheet.spill_region_at_anchor(0U, 25U);
  ASSERT_NE(region, nullptr);
  EXPECT_EQ(region->rows, Sheet::kMaxRows);
  EXPECT_EQ(region->cols, 3U);
  EXPECT_EQ(sheet.resolve_cell_value(0U, 25U).as_number(), 1.0);
  EXPECT_EQ(sheet.resolve_cell_value(0U, 26U).as_number(), 10.0);
  EXPECT_EQ(sheet.resolve_cell_value(0U, 27U).as_number(), 100.0);
  EXPECT_EQ(sheet.resolve_cell_value(1U, 25U).as_number(), 2.0);
  EXPECT_EQ(sheet.resolve_cell_value(1U, 26U).as_number(), 0.0);
  EXPECT_EQ(sheet.resolve_cell_value(2U, 27U).as_number(), 0.0);
  EXPECT_TRUE(sheet.resolve_cell_value(0U, 28U).is_blank());
}

// ---------------------------------------------------------------------------
// Boundary of the change: only the bare reference standing as the entire
// formula spills. Materialising a grid-axis rectangle behind an operator or
// a scalar function argument would conjure millions of cells per operand,
// so those positions keep the top-left anchor projection.
// ---------------------------------------------------------------------------

TEST(WholeAxisSpill, RefusedFootprintIsDecidedWithoutBuildingTheRectangle) {
  // The answer for both refusal kinds is `#SPILL!` whether or not the
  // rectangle was built first, so the value alone cannot tell them apart.
  // What distinguishes them is the allocation: building a whole-column
  // result costs a full column of `Value`s, and a formula reachable from
  // untrusted input must not pay that to produce an error.
  Workbook wb = Workbook::create();
  SeedNumber(wb, 0U, 0U, 1.0);
  SeedNumber(wb, 1U, 0U, 2.0);
  SeedNumber(wb, 2U, 0U, 3.0);
  SeedNumber(wb, 499U, 25U, 7.0);
  const Sheet& sheet = wb.sheet(0);

  // Rectangle does not fit below row 2.
  std::size_t overflow_bytes = 0;
  const Value overflow = AdhocValueAndBytes(wb, sheet, 1U, 25U, "=A:A", &overflow_bytes);
  ASSERT_TRUE(overflow.is_error());
  EXPECT_EQ(overflow.as_error(), ErrorCode::Spill);
  EXPECT_LT(overflow_bytes, kWholeColumnBytes / 16U);

  // Rectangle fits, but Z500 occupies it.
  std::size_t blocked_bytes = 0;
  const Value blocked = AdhocValueAndBytes(wb, sheet, 0U, 25U, "=A:A", &blocked_bytes);
  ASSERT_TRUE(blocked.is_error());
  EXPECT_EQ(blocked.as_error(), ErrorCode::Spill);
  EXPECT_LT(blocked_bytes, kWholeColumnBytes / 16U);

  // Control: the same reference where the spill is admissible does build
  // the rectangle, which is what shows the bound above is discriminating
  // rather than trivially true for this formula.
  Workbook clear_wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(clear_wb.set_cell_value(0U, 0U, 0U, Value::number(1.0))));
  std::size_t spilled_bytes = 0;
  const Value spilled = AdhocValueAndBytes(clear_wb, clear_wb.sheet(0), 0U, 25U, "=A:A", &spilled_bytes);
  ASSERT_TRUE(spilled.is_array());
  EXPECT_GE(spilled_bytes, kWholeColumnBytes);
}

TEST(WholeAxisSpill, RefusedFootprintIsRecordedAndRetriedWhenTheBlockerGoes) {
  // The refusal is decided without materialising the rectangle, so it does
  // not pass through `commit_spill`. It still has to be recorded: the
  // remembered footprint is what lets the anchor be retried once the
  // blocker is cleared. A refusal that merely returned `#SPILL!` would
  // leave the anchor stranded, which is what this asserts is not happening.
  Workbook wb = Workbook::create();
  SeedNumber(wb, 2U, 0U, 3.0);
  SeedNumber(wb, 2U, 1U, 30.0);
  SeedNumber(wb, 4U, 25U, 9.0);
  Sheet& sheet = RecalcWith(wb, 4U, 0U, "=3:3");
  ASSERT_TRUE(sheet.resolve_cell_value(4U, 0U).is_error());
  ASSERT_EQ(sheet.resolve_cell_value(4U, 0U).as_error(), ErrorCode::Spill);

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 4U, 25U, Value::blank())));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
  const SpillRegion* region = sheet.spill_region_at_anchor(4U, 0U);
  ASSERT_NE(region, nullptr);
  EXPECT_EQ(region->rows, 1U);
  EXPECT_EQ(region->cols, Sheet::kMaxCols);
  EXPECT_EQ(sheet.resolve_cell_value(4U, 0U).as_number(), 3.0);
  EXPECT_EQ(sheet.resolve_cell_value(4U, 1U).as_number(), 30.0);
}

TEST(WholeAxisSpill, WholeAxisBehindAnOperatorKeepsAnchorProjection) {
  Workbook wb = Workbook::create();
  SeedNumber(wb, 0U, 0U, 1.0);
  SeedNumber(wb, 1U, 0U, 2.0);
  const Sheet& sheet = RecalcWith(wb, 0U, 25U, "=A:C+1");
  const Value v = sheet.resolve_cell_value(0U, 25U);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2.0);
  EXPECT_EQ(sheet.spill_region_at_anchor(0U, 25U), nullptr);
}

TEST(WholeAxisSpill, WholeAxisAsRangeAggregatorArgumentIsUnchanged) {
  // `SUM(A:A)` still reads the populated extent rather than 1,048,576
  // cells; range-aware dispatch intercepts the reference before the
  // spilling position is reached.
  Workbook wb = Workbook::create();
  SeedNumber(wb, 0U, 0U, 1.0);
  SeedNumber(wb, 1U, 0U, 2.0);
  SeedNumber(wb, 2U, 0U, 3.0);
  const Sheet& sheet = RecalcWith(wb, 0U, 25U, "=SUM(A:A)");
  const Value v = sheet.resolve_cell_value(0U, 25U);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 6.0);
}

// ---------------------------------------------------------------------------
// A whole-axis reference passed to a function has its declared shape: two
// whole columns agree however far each is populated, and a positional read
// beyond the populated extent sees a blank cell rather than falling off the
// end of a trimmed rectangle.
// ---------------------------------------------------------------------------

// Column A holds 1..5, column B holds 10, 20, 30: the two whole columns are
// populated to different depths.
Workbook UnevenColumns() {
  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 5U; ++r) {
    SeedNumber(wb, r, 0U, static_cast<double>(r + 1U));
  }
  for (std::uint32_t r = 0; r < 3U; ++r) {
    SeedNumber(wb, r, 1U, 10.0 * static_cast<double>(r + 1U));
  }
  return wb;
}

Value UnevenValue(const std::string& formula) {
  const Workbook wb = UnevenColumns();
  Arena arena;
  const Value v = evaluate_formula_text_array(wb, wb.sheet(0), 0U, 25U, formula, arena, default_registry());
  // Only scalar results are returned by value; arrays borrow the arena.
  EXPECT_FALSE(v.is_array()) << formula;
  return v;
}

void ExpectNumber(const Value& v, double expected, const std::string& formula) {
  ASSERT_TRUE(v.is_number()) << formula << " kind=" << static_cast<int>(v.kind());
  EXPECT_EQ(v.as_number(), expected) << formula;
}

TEST(WholeAxisArgumentShape, XlookupOverUnevenColumns) {
  ExpectNumber(UnevenValue("=XLOOKUP(2,A:A,B:B)"), 20.0, "XLOOKUP(2,A:A,B:B)");
  ExpectNumber(UnevenValue("=XLOOKUP(20,B:B,A:A)"), 2.0, "XLOOKUP(20,B:B,A:A)");
}

TEST(WholeAxisArgumentShape, ConditionalAggregatesOverUnevenColumns) {
  ExpectNumber(UnevenValue("=SUMIFS(B:B,A:A,\">1\")"), 50.0, "SUMIFS(B:B,A:A,\">1\")");
  ExpectNumber(UnevenValue("=SUMIFS(A:A,B:B,\">10\")"), 5.0, "SUMIFS(A:A,B:B,\">10\")");
  ExpectNumber(UnevenValue("=COUNTIFS(A:A,\">1\",B:B,\">0\")"), 2.0, "COUNTIFS(A:A,\">1\",B:B,\">0\")");
  ExpectNumber(UnevenValue("=AVERAGEIFS(B:B,A:A,\">1\")"), 25.0, "AVERAGEIFS(B:B,A:A,\">1\")");
  ExpectNumber(UnevenValue("=MAXIFS(B:B,A:A,\">1\")"), 30.0, "MAXIFS(B:B,A:A,\">1\")");
  // Column C is empty: every row of A beyond the first matches `">1"` and
  // meets a blank C cell.
  ExpectNumber(UnevenValue("=COUNTIFS(A:A,\">1\",C:C,\"\")"), 4.0, "COUNTIFS(A:A,\">1\",C:C,\"\")");
  ExpectNumber(UnevenValue("=SUMPRODUCT((A:A>1)*B:B)"), 50.0, "SUMPRODUCT((A:A>1)*B:B)");
}

TEST(WholeAxisArgumentShape, FilterOverUnevenColumns) {
  const Workbook wb = UnevenColumns();
  Arena arena;
  const Value kept =
      evaluate_formula_text_array(wb, wb.sheet(0), 0U, 25U, "=FILTER(A:A,B:B>10)", arena, default_registry());
  ASSERT_TRUE(kept.is_array());
  ASSERT_EQ(kept.as_array_rows(), 2U);
  ASSERT_EQ(kept.as_array_cols(), 1U);
  EXPECT_EQ(kept.as_array_cells()[0].as_number(), 2.0);
  EXPECT_EQ(kept.as_array_cells()[1].as_number(), 3.0);

  // Rows 2..5 of A pass; B holds values in only two of them.
  const Value padded =
      evaluate_formula_text_array(wb, wb.sheet(0), 0U, 25U, "=FILTER(B:B,A:A>1)", arena, default_registry());
  ASSERT_TRUE(padded.is_array());
  ASSERT_EQ(padded.as_array_rows(), 4U);
  EXPECT_EQ(padded.as_array_cells()[0].as_number(), 20.0);
  EXPECT_EQ(padded.as_array_cells()[1].as_number(), 30.0);
  EXPECT_FALSE(padded.as_array_cells()[2].is_error());
  EXPECT_FALSE(padded.as_array_cells()[3].is_error());
}

TEST(WholeAxisArgumentShape, IndexMatchAcrossSheetsReadsPastThePopulatedExtent) {
  Workbook wb = Workbook::create();
  const std::size_t data = wb.add_sheet("Data");
  for (std::uint32_t r = 0; r < 5U; ++r) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(data, r, 0U, Value::number(static_cast<double>(r + 1U)))));
  }
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(data, 0U, 1U, Value::number(10.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(data, 1U, 1U, Value::number(20.0))));
  Arena arena;
  const Value hit = evaluate_formula_text_array(wb, wb.sheet(0), 0U, 0U, "=INDEX(Data!B:B,MATCH(2,Data!A:A,0))", arena,
                                                default_registry());
  ExpectNumber(hit, 20.0, "INDEX(Data!B:B,MATCH(2,Data!A:A,0))");
  // Data!B4 is blank: INDEX reads it, it does not fall off a trimmed column.
  const Value blank = evaluate_formula_text_array(wb, wb.sheet(0), 0U, 0U, "=INDEX(Data!B:B,MATCH(4,Data!A:A,0))",
                                                  arena, default_registry());
  ExpectNumber(blank, 0.0, "INDEX(Data!B:B,MATCH(4,Data!A:A,0))");
  // A two-column whole reference populated in one row is still two-dimensional.
  const Value row2 =
      evaluate_formula_text_array(wb, wb.sheet(0), 0U, 0U, "=INDEX(Data!A:B,2,2)", arena, default_registry());
  ExpectNumber(row2, 20.0, "INDEX(Data!A:B,2,2)");
}

TEST(WholeAxisArgumentShape, DeclaredShapeDecidesAgreementAndOrientation) {
  // `A:A` declares 1,048,576 rows, so it never pairs with a five-row range
  // even though column A holds exactly five values.
  const Value sumifs = UnevenValue("=SUMIFS(A1:A5,B:B,\">0\")");
  ASSERT_TRUE(sumifs.is_error());
  EXPECT_EQ(sumifs.as_error(), ErrorCode::Value);
  const Value xlookup = UnevenValue("=XLOOKUP(2,A:A,A1:A5)");
  ASSERT_TRUE(xlookup.is_error());
  EXPECT_EQ(xlookup.as_error(), ErrorCode::Value);

  // Only the first two rows of A:C hold values, yet the reference is taller
  // than wide: LOOKUP scans column A and returns from column C, and XLOOKUP
  // rejects it as two-dimensional.
  Workbook wb = Workbook::create();
  SeedNumber(wb, 0U, 0U, 1.0);
  SeedNumber(wb, 1U, 0U, 2.0);
  SeedNumber(wb, 0U, 2U, 10.0);
  SeedNumber(wb, 1U, 2U, 20.0);
  Arena arena;
  ExpectNumber(evaluate_formula_text_array(wb, wb.sheet(0), 0U, 25U, "=LOOKUP(2,A:C)", arena, default_registry()), 20.0,
               "LOOKUP(2,A:C)");
  const Value flat =
      evaluate_formula_text_array(wb, wb.sheet(0), 0U, 25U, "=XLOOKUP(1,A:B,C:C)", arena, default_registry());
  ASSERT_TRUE(flat.is_error());
  EXPECT_EQ(flat.as_error(), ErrorCode::Value);
}

TEST(WholeAxisArgumentShape, EmptyWholeColumnsStillHonourIfNotFound) {
  Workbook wb = Workbook::create();
  SeedNumber(wb, 0U, 25U, 1.0);
  Arena arena;
  const Value v =
      evaluate_formula_text_array(wb, wb.sheet(0), 0U, 24U, "=XLOOKUP(1,A:A,B:B,\"none\")", arena, default_registry());
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "none");
}

// Column A holds 1..700 and the reference spans A:XFD, so the rectangle the
// lookups name holds 700 * 16384 cells -- past the range-expansion ceiling --
// while the lookups only ever read one column and one cell.
Workbook WideTable() {
  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 700U; ++r) {
    SeedNumber(wb, r, 0U, static_cast<double>(r + 1U));
    SeedNumber(wb, r, 1U, static_cast<double>(r + 1U) * 2.0);
  }
  return wb;
}

TEST(WholeAxisArgumentShape, LookupsOverAWideTableReadOnlyWhatTheyNeed) {
  const Workbook wb = WideTable();
  Arena arena;
  const auto eval = [&](const char* formula) {
    return evaluate_formula_text_array(wb, wb.sheet(0), 0U, 25U, formula, arena, default_registry());
  };
  ExpectNumber(eval("=VLOOKUP(650,A:XFD,2,FALSE)"), 1300.0, "VLOOKUP(650,A:XFD,2,FALSE)");
  ExpectNumber(eval("=VLOOKUP(650.5,A:XFD,1,TRUE)"), 650.0, "VLOOKUP(650.5,A:XFD,1,TRUE)");
  ExpectNumber(eval("=INDEX(A:XFD,650,2)"), 1300.0, "INDEX(A:XFD,650,2)");
  ExpectNumber(eval("=MATCH(650,A:A,0)"), 650.0, "MATCH(650,A:A,0)");
  // A 2-D whole reference is rejected by its shape without being read.
  const Value two_d = eval("=MATCH(650,A:XFD,0)");
  ASSERT_TRUE(two_d.is_error());
  EXPECT_EQ(two_d.as_error(), ErrorCode::NA);
  // The row-axis mirror: rows 1..1000 of a 10001-column table.
  Workbook row_wb = Workbook::create();
  for (std::uint32_t c = 0; c < 10001U; ++c) {
    ASSERT_TRUE(static_cast<bool>(row_wb.set_cell_value(0U, 0U, c, Value::number(static_cast<double>(c + 1U)))));
    ASSERT_TRUE(static_cast<bool>(row_wb.set_cell_value(0U, 1U, c, Value::number(static_cast<double>(c + 1U) * 3.0))));
  }
  const Value h = evaluate_formula_text_array(row_wb, row_wb.sheet(0), 1500U, 0U, "=HLOOKUP(7,1:1000,2,FALSE)", arena,
                                              default_registry());
  ExpectNumber(h, 21.0, "HLOOKUP(7,1:1000,2,FALSE)");
}

}  // namespace
}  // namespace eval
}  // namespace formulon
