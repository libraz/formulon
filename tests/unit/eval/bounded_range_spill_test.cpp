//
// Unit tests for a bare bounded range standing as the entire formula
// (`=A1:A1048576`, `=A1:C1048576`, `=A3:XFD3`).
//
// A bounded rectangle and the whole-axis spelling of the same rectangle are
// one expression with one answer, so the two must also cost the same to
// refuse. `=A1:A1048576` at a blocked or overflowing anchor is `#SPILL!`
// exactly as `=A:A` is, and neither has to be built to say so — the
// footprint decides that from what the sheet stores.
//
// What these tests pin, therefore, is a pair of properties:
//
//   * the verdict, which is what it already was: `#SPILL!` when the
//     rectangle leaves the grid measured from the anchor or something
//     occupies it, the full rectangle otherwise;
//   * the allocation, which is what changed: a refusal stays orders of
//     magnitude below the size of the rectangle it refused, while an
//     admissible spill still pays for the array it genuinely produces.
//
// The verdict half passes with or without the footprint pre-check; only the
// allocation half distinguishes them. Both are kept because a bound that no
// longer discriminates is invisible from the value alone.
//
// Observed Excel behaviour for the equivalent whole-axis spellings:
// tests/oracle/cases/whole_axis_spill.yaml.

#include <cstddef>
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

// Evaluates through the read-only driver and reports what the evaluation
// allocated. Every cell a materialised rectangle holds passes through this
// arena, so the byte count witnesses whether the rectangle was built — the
// returned value cannot, since both outcomes are the same `#SPILL!`.
Value AdhocValueAndBytes(const Workbook& wb, const Sheet& sheet, std::uint32_t row, std::uint32_t col,
                         const std::string& formula, std::size_t* out_bytes) {
  Arena arena;
  const Value v = evaluate_formula_text_array(wb, sheet, row, col, formula, arena, default_registry());
  *out_bytes = arena.bytes_used();
  return v;
}

// One full grid column of `Value`s: the floor for a materialised
// whole-column rectangle, and the yardstick a refusal must stay far under.
constexpr std::size_t kWholeColumnBytes = static_cast<std::size_t>(Sheet::kMaxRows) * sizeof(Value);

// ---------------------------------------------------------------------------
// Allocation: a refused rectangle is never built.
// ---------------------------------------------------------------------------

TEST(BoundedRangeSpill, RefusedFootprintIsDecidedWithoutBuildingTheRectangle) {
  Workbook wb = Workbook::create();
  SeedNumber(wb, 0U, 0U, 1.0);
  SeedNumber(wb, 1U, 0U, 2.0);
  SeedNumber(wb, 2U, 0U, 3.0);
  SeedNumber(wb, 499U, 25U, 7.0);
  const Sheet& sheet = wb.sheet(0);

  // Rectangle does not fit below row 2.
  std::size_t overflow_bytes = 0;
  const Value overflow = AdhocValueAndBytes(wb, sheet, 1U, 25U, "=A1:A1048576", &overflow_bytes);
  ASSERT_TRUE(overflow.is_error());
  EXPECT_EQ(overflow.as_error(), ErrorCode::Spill);
  EXPECT_LT(overflow_bytes, kWholeColumnBytes / 16U);

  // Rectangle fits, but Z500 occupies it.
  std::size_t blocked_bytes = 0;
  const Value blocked = AdhocValueAndBytes(wb, sheet, 0U, 25U, "=A1:A1048576", &blocked_bytes);
  ASSERT_TRUE(blocked.is_error());
  EXPECT_EQ(blocked.as_error(), ErrorCode::Spill);
  EXPECT_LT(blocked_bytes, kWholeColumnBytes / 16U);

  // The three-column span is the same refusal over three times the area, so
  // its cost must not scale with the area either.
  std::size_t wide_bytes = 0;
  const Value wide = AdhocValueAndBytes(wb, sheet, 0U, 25U, "=A1:C1048576", &wide_bytes);
  ASSERT_TRUE(wide.is_error());
  EXPECT_EQ(wide.as_error(), ErrorCode::Spill);
  EXPECT_LT(wide_bytes, kWholeColumnBytes / 16U);

  // Control: the same reference where the spill is admissible does build the
  // rectangle. Without it the bound above would be trivially satisfiable by
  // a formula that never allocates anything.
  Workbook clear_wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(clear_wb.set_cell_value(0U, 0U, 0U, Value::number(1.0))));
  std::size_t spilled_bytes = 0;
  const Value spilled = AdhocValueAndBytes(clear_wb, clear_wb.sheet(0), 0U, 25U, "=A1:A1048576", &spilled_bytes);
  ASSERT_TRUE(spilled.is_array());
  EXPECT_GE(spilled_bytes, kWholeColumnBytes);
}

TEST(BoundedRangeSpill, BothSpellingsOfOneRectangleCostTheSame) {
  // `=A:A` and `=A1:A1048576` describe one rectangle and return one answer.
  // The point of the pre-check is that they also reach it the same way, so
  // the bounded spelling is held to the whole-axis spelling's own cost
  // rather than to a fixed constant.
  Workbook wb = Workbook::create();
  SeedNumber(wb, 0U, 0U, 1.0);
  SeedNumber(wb, 499U, 25U, 7.0);
  const Sheet& sheet = wb.sheet(0);

  std::size_t whole_axis_bytes = 0;
  const Value whole_axis = AdhocValueAndBytes(wb, sheet, 0U, 25U, "=A:A", &whole_axis_bytes);
  std::size_t bounded_bytes = 0;
  const Value bounded = AdhocValueAndBytes(wb, sheet, 0U, 25U, "=A1:A1048576", &bounded_bytes);

  ASSERT_TRUE(whole_axis.is_error());
  ASSERT_TRUE(bounded.is_error());
  EXPECT_EQ(whole_axis.as_error(), ErrorCode::Spill);
  EXPECT_EQ(bounded.as_error(), ErrorCode::Spill);
  // Neither spelling parses to the same AST, so the two arenas hold slightly
  // different scratch; an order of magnitude of slack keeps that irrelevant
  // while still failing on a materialised rectangle.
  EXPECT_LT(bounded_bytes, whole_axis_bytes * 10U + 4096U);
}

// ---------------------------------------------------------------------------
// Verdict: unchanged by the pre-check, and asserted through recalc so the
// committed footprint is visible.
// ---------------------------------------------------------------------------

TEST(BoundedRangeSpill, RectangleBelowRowOneOverflows) {
  Workbook wb = Workbook::create();
  SeedNumber(wb, 0U, 0U, 1.0);
  SeedNumber(wb, 1U, 0U, 2.0);
  const Sheet& sheet = RecalcWith(wb, 1U, 25U, "=A1:A1048576");
  const Value v = sheet.resolve_cell_value(1U, 25U);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Spill);
  EXPECT_EQ(sheet.spill_region_at_anchor(1U, 25U), nullptr);
}

TEST(BoundedRangeSpill, BlockerFarBelowPopulatedExtentStopsTheSpill) {
  // Column A holds three cells; the blocker sits at Z500, far outside that
  // extent but inside the declared rectangle.
  Workbook wb = Workbook::create();
  SeedNumber(wb, 0U, 0U, 1.0);
  SeedNumber(wb, 1U, 0U, 2.0);
  SeedNumber(wb, 2U, 0U, 3.0);
  SeedNumber(wb, 499U, 25U, 7.0);
  const Sheet& sheet = RecalcWith(wb, 0U, 25U, "=A1:A1048576");
  const Value v = sheet.resolve_cell_value(0U, 25U);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Spill);
  EXPECT_EQ(sheet.spill_region_at_anchor(0U, 25U), nullptr);
  // A refused spill writes nothing: the blocker keeps its own value.
  EXPECT_EQ(sheet.resolve_cell_value(499U, 25U).as_number(), 7.0);
}

TEST(BoundedRangeSpill, RefusedFootprintIsRecordedAndRetriedWhenTheBlockerGoes) {
  // A refusal decided before materialising never reaches `commit_spill`, so
  // the rectangle has to be remembered separately; without it the anchor is
  // stranded at `#SPILL!` and never retried. `=A3:XFD3` is the bounded
  // spelling of `=3:3`.
  Workbook wb = Workbook::create();
  SeedNumber(wb, 2U, 0U, 3.0);
  SeedNumber(wb, 2U, 1U, 30.0);
  SeedNumber(wb, 4U, 25U, 9.0);
  Sheet& sheet = RecalcWith(wb, 4U, 0U, "=A3:XFD3");
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

TEST(BoundedRangeSpill, AdmissibleRectangleStillSpillsInFull) {
  // The unpopulated cells of the rectangle spill as 0 and the cell below the
  // footprint stays blank, so nothing about an accepted spill moved.
  Workbook wb = Workbook::create();
  SeedNumber(wb, 2U, 0U, 3.0);
  SeedNumber(wb, 2U, 1U, 30.0);
  const Sheet& sheet = RecalcWith(wb, 4U, 0U, "=A3:XFD3");
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

TEST(BoundedRangeSpill, SmallRectangleSpillsUnchanged) {
  Workbook wb = Workbook::create();
  SeedNumber(wb, 0U, 0U, 1.0);
  SeedNumber(wb, 1U, 0U, 2.0);
  SeedNumber(wb, 0U, 1U, 10.0);
  SeedNumber(wb, 1U, 1U, 20.0);
  const Sheet& sheet = RecalcWith(wb, 0U, 25U, "=A1:B2");
  const SpillRegion* region = sheet.spill_region_at_anchor(0U, 25U);
  ASSERT_NE(region, nullptr);
  EXPECT_EQ(region->rows, 2U);
  EXPECT_EQ(region->cols, 2U);
  EXPECT_EQ(sheet.resolve_cell_value(0U, 25U).as_number(), 1.0);
  EXPECT_EQ(sheet.resolve_cell_value(0U, 26U).as_number(), 10.0);
  EXPECT_EQ(sheet.resolve_cell_value(1U, 25U).as_number(), 2.0);
  EXPECT_EQ(sheet.resolve_cell_value(1U, 26U).as_number(), 20.0);
}

// ---------------------------------------------------------------------------
// A small rectangle containing its own formula cell. Excel abandons the
// closure and leaves 0 with no spill; Formulon reports every cycle member as
// `#REF!`, which is the registered engine-wide divergence rather than
// anything specific to this shape.
//
// The two drivers answer differently here, and neither answer is visible
// from the other, so they are asserted as two tests rather than two halves
// of one: a single test stops at the first failing assertion, which leaves
// the second driver unexercised exactly when something has moved. The
// oracle case for this formula
// (whole_axis_spill.bounded_rectangle_containing_its_own_formula_cell)
// expects a scalar 0 and passes — but it compares through a top-left
// projection, so it would equally pass against a 3x3 of zeros. It pins the
// anchor's value, not the shape. That is what these assertions add.
//
// `=A1:C3` is neither full-height nor full-width, so Excel does not
// canonicalise it and it has no whole-axis twin to agree with. It therefore
// keeps the per-cell cycle route rather than the pre-empted one.
// ---------------------------------------------------------------------------

TEST(BoundedRangeSpill, SmallRectangleContainingItsOwnFormulaCellSpillsZerosInTheReadOnlyDriver) {
  // On the empty sheet the oracle case uses, the rectangle is materialised
  // in full and every cell resolves to 0: the read-only path has no
  // committed footprint to close a cycle through, and an unpopulated cell
  // projects to 0 under the grid contract.
  Workbook wb = Workbook::create();
  Arena arena;
  const Value adhoc = evaluate_formula_text_array(wb, wb.sheet(0), 1U, 1U, "=A1:C3", arena, default_registry());
  ASSERT_TRUE(adhoc.is_array());
  EXPECT_EQ(adhoc.as_array_rows(), 3U);
  EXPECT_EQ(adhoc.as_array_cols(), 3U);
  const Value* cells = adhoc.as_array_cells();
  for (std::size_t i = 0; i < 9U; ++i) {
    EXPECT_TRUE(cells[i].is_number()) << "cell " << i;
    EXPECT_EQ(cells[i].as_number(), 0.0) << "cell " << i;
  }
}

TEST(BoundedRangeSpill, SmallRectangleContainingItsOwnFormulaCellIsCircularInTheCommittingDriver) {
  // The committed footprint feeds a self-edge back into the dependency graph
  // and the engine's cycle policy answers, so the cell is `#REF!` with
  // nothing spilled. This is the answer a workbook shows, and it is reached
  // through the graph rather than by pre-emption — which is why it is
  // unchanged by where the pre-emption line is drawn.
  Workbook wb = Workbook::create();
  const Sheet& sheet = RecalcWith(wb, 1U, 1U, "=A1:C3");
  const Value cell = sheet.resolve_cell_value(1U, 1U);
  ASSERT_TRUE(cell.is_error());
  EXPECT_EQ(cell.as_error(), ErrorCode::Ref);
  EXPECT_EQ(sheet.spill_region_at_anchor(1U, 1U), nullptr);
  EXPECT_TRUE(sheet.resolve_cell_value(1U, 2U).is_blank());
  EXPECT_TRUE(sheet.resolve_cell_value(2U, 1U).is_blank());
}

TEST(BoundedRangeSpill, SameRectangleAnchoredOutsideItselfSpills) {
  // The oracle pair's control: same rectangle, anchor outside it, so nothing
  // is circular and the 3x3 spills. It is what shows the case above is about
  // the anchor's position rather than about the rectangle.
  Workbook wb = Workbook::create();
  const Sheet& sheet = RecalcWith(wb, 0U, 4U, "=A1:C3");
  const SpillRegion* region = sheet.spill_region_at_anchor(0U, 4U);
  ASSERT_NE(region, nullptr);
  EXPECT_EQ(region->rows, 3U);
  EXPECT_EQ(region->cols, 3U);
  EXPECT_TRUE(sheet.resolve_cell_value(0U, 4U).is_number());
}

// ---------------------------------------------------------------------------
// Boundary of the pre-check: it applies to the bare rectangle standing as
// the whole formula and nowhere else.
// ---------------------------------------------------------------------------

TEST(BoundedRangeSpill, SingleCellRangeKeepsItsScalarDegradation) {
  Workbook wb = Workbook::create();
  SeedNumber(wb, 0U, 0U, 7.0);
  const Sheet& sheet = RecalcWith(wb, 0U, 25U, "=A1:A1");
  const Value v = sheet.resolve_cell_value(0U, 25U);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 7.0);
  EXPECT_EQ(sheet.spill_region_at_anchor(0U, 25U), nullptr);
}

TEST(BoundedRangeSpill, RectangleBehindAnOperatorStillBroadcasts) {
  // An operand is not a spilling position, so the pre-check must not reach
  // it: the operator broadcasts over the rectangle exactly as before.
  Workbook wb = Workbook::create();
  SeedNumber(wb, 0U, 0U, 1.0);
  SeedNumber(wb, 1U, 0U, 2.0);
  const Sheet& sheet = RecalcWith(wb, 0U, 25U, "=A1:A2+1");
  const SpillRegion* region = sheet.spill_region_at_anchor(0U, 25U);
  ASSERT_NE(region, nullptr);
  EXPECT_EQ(region->rows, 2U);
  EXPECT_EQ(region->cols, 1U);
  EXPECT_EQ(sheet.resolve_cell_value(0U, 25U).as_number(), 2.0);
  EXPECT_EQ(sheet.resolve_cell_value(1U, 25U).as_number(), 3.0);
}

// ---------------------------------------------------------------------------
// Rectangles above the range-expansion ceiling. Both cases use an anchor
// outside the rectangle, so circularity cannot fire and the footprint and
// the ceiling are the only two things that can answer.
//
// The `#SPILL!` verdict below is Formulon's choice, not a pinned Excel
// observation: Excel declines to compute a rectangle this size at all,
// leaving the cell uncomputed and returning neither `#SPILL!` nor `#CALC!`.
// What is being pinned is that geometry is decided by the anchor rather
// than by an internal expansion limit that happens to be reached first.
// ---------------------------------------------------------------------------

TEST(BoundedRangeSpill, RectangleTooLargeToExpandIsStillRefusedAsSpill) {
  // `=A1:XFC1048576` at XFD1: full-height, and XFD is the one column the
  // rectangle excludes, so the anchor sits outside it. The rectangle cannot
  // fit measured from XFD1, so it is refused on geometry — and stays refused
  // even though it is far above the ceiling and could never have been built.
  //
  // Building first reached the ceiling before the anchor was ever consulted
  // and reported `#CALC!`. Nothing here refuses a rectangle for being large;
  // the next test shows the ceiling still doing its own job.
  Workbook wb = Workbook::create();
  SeedNumber(wb, 0U, 0U, 1.0);
  const Sheet& sheet = wb.sheet(0);

  std::size_t refused_bytes = 0;
  const Value refused = AdhocValueAndBytes(wb, sheet, 0U, Sheet::kMaxCols - 1U, "=A1:XFC1048576", &refused_bytes);
  ASSERT_TRUE(refused.is_error());
  EXPECT_EQ(refused.as_error(), ErrorCode::Spill);
  EXPECT_LT(refused_bytes, kWholeColumnBytes / 16U);
}

TEST(BoundedRangeSpill, RectangleTheFootprintAdmitsStillMeetsTheExpansionCeiling) {
  // `=A1:XFD611` at A1000: full-width and over the ceiling, but the anchor
  // is below the rectangle's last row, so it is neither circular nor
  // blocked. The footprint admits it, the expansion is what refuses it, and
  // the answer is `#CALC!` — which is what shows the pre-check did not turn
  // into a size limit of its own.
  Workbook wb = Workbook::create();
  Arena arena;
  const Value admitted =
      evaluate_formula_text_array(wb, wb.sheet(0), 999U, 0U, "=A1:XFD611", arena, default_registry());
  ASSERT_TRUE(admitted.is_error());
  EXPECT_EQ(admitted.as_error(), ErrorCode::Calc);
}

TEST(BoundedRangeSpill, SpellingsExcelCanonicalisesAnswerIdentically) {
  // Excel rewrites `=A1:XFD1048576` into the whole-axis spelling on entry,
  // so after entry there is one formula. Both spellings are anchored inside
  // the rectangle — a full-grid rectangle contains every cell, so every
  // anchor is — and both therefore report the cycle.
  //
  // This is the assertion the full-height / full-width rule exists for. It
  // is `#REF!` because Formulon reports every cycle member that way; Excel
  // leaves 0, which is the registered engine-wide divergence and not
  // something this pair claims to match.
  Workbook wb = Workbook::create();
  SeedNumber(wb, 0U, 0U, 1.0);
  const Sheet& sheet = wb.sheet(0);

  std::size_t bounded_bytes = 0;
  const Value bounded = AdhocValueAndBytes(wb, sheet, 0U, 1U, "=A1:XFD1048576", &bounded_bytes);
  std::size_t whole_axis_bytes = 0;
  const Value whole_axis = AdhocValueAndBytes(wb, sheet, 0U, 1U, "=A:XFD", &whole_axis_bytes);

  ASSERT_TRUE(bounded.is_error());
  ASSERT_TRUE(whole_axis.is_error());
  EXPECT_EQ(bounded.as_error(), ErrorCode::Ref);
  EXPECT_EQ(whole_axis.as_error(), ErrorCode::Ref);
  EXPECT_EQ(bounded.as_error(), whole_axis.as_error());
  // Neither spelling builds anything to say so.
  EXPECT_LT(bounded_bytes, kWholeColumnBytes / 16U);
  EXPECT_LT(whole_axis_bytes, kWholeColumnBytes / 16U);
}

TEST(BoundedRangeSpill, RectangleAsRangeAggregatorArgumentIsUnchanged) {
  Workbook wb = Workbook::create();
  SeedNumber(wb, 0U, 0U, 1.0);
  SeedNumber(wb, 1U, 0U, 2.0);
  SeedNumber(wb, 2U, 0U, 3.0);
  const Sheet& sheet = RecalcWith(wb, 0U, 25U, "=SUM(A1:A1048576)");
  const Value v = sheet.resolve_cell_value(0U, 25U);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 6.0);
}

TEST(BoundedRangeSpill, CrossSheetRectangleIsRefusedAgainstTheAnchorSheet) {
  // The rectangle reads Sheet2 but occupies Sheet1, so the footprint that
  // refuses it is the anchor sheet's — Sheet2's contents are irrelevant to
  // whether the spill fits.
  Workbook wb = Workbook::create();
  wb.add_sheet("Sheet2");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1U, 0U, 0U, Value::number(5.0))));
  SeedNumber(wb, 499U, 25U, 7.0);
  const Sheet& sheet = RecalcWith(wb, 0U, 25U, "=Sheet2!A1:A1048576");
  const Value v = sheet.resolve_cell_value(0U, 25U);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Spill);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
