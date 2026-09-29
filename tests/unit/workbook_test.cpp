//
// Unit tests for the Workbook skeleton. Verifies the factory shape, sheet
// accessor mutation, and the basic byte-level shape of the save() output.

#include "workbook.h"

#include <cstddef>
#include <cstdint>
#include <utility>
#include <vector>

#include "cell.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "passthrough_part.h"
#include "sheet.h"
#include "value.h"

namespace formulon {
namespace {

TEST(WorkbookTest, CreateYieldsSingleSheetNamedSheet1) {
  Workbook wb = Workbook::create();
  ASSERT_EQ(wb.sheet_count(), 1u);
  EXPECT_EQ(wb.sheet(0).name(), "Sheet1");
}

TEST(WorkbookTest, SheetCountReturnsOne) {
  Workbook wb = Workbook::create();
  EXPECT_EQ(wb.sheet_count(), static_cast<std::size_t>(1));
}

TEST(WorkbookTest, MutateSheetNamePropagates) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_name("Daten");
  EXPECT_EQ(wb.sheet(0).name(), "Daten");

  // The const overload must observe the same state.
  const Workbook& const_wb = wb;
  EXPECT_EQ(const_wb.sheet(0).name(), "Daten");
}

TEST(WorkbookTest, SaveProducesNonEmptyBytes) {
  Workbook wb = Workbook::create();
  auto result = wb.save();
  ASSERT_TRUE(static_cast<bool>(result)) << "save() failed: " << result.error().message;
  const std::vector<std::uint8_t>& bytes = result.value();
  EXPECT_GT(bytes.size(), 0u);
}

TEST(WorkbookTest, SaveIsZipMagicBytes) {
  Workbook wb = Workbook::create();
  auto result = wb.save();
  ASSERT_TRUE(static_cast<bool>(result));
  const std::vector<std::uint8_t>& bytes = result.value();
  ASSERT_GE(bytes.size(), 4u);
  EXPECT_EQ(bytes[0], 0x50u);  // 'P'
  EXPECT_EQ(bytes[1], 0x4Bu);  // 'K'
  EXPECT_EQ(bytes[2], 0x03u);
  EXPECT_EQ(bytes[3], 0x04u);
}

TEST(WorkbookTest, ApproximateMemoryGrowsWithTheCellStore) {
  // The figure exists so a host runtime can size a workbook it only sees
  // as a handle, which means the one property that must hold is that it
  // moves with the content rather than staying at the empty baseline.
  Workbook wb = Workbook::create();
  const std::size_t empty = wb.approximate_memory_bytes();
  EXPECT_GT(empty, sizeof(Workbook));

  for (std::uint32_t row = 0; row < 500U; ++row) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, row, 0U, "=1+2")));
  }
  const std::size_t filled = wb.approximate_memory_bytes();
  EXPECT_GT(filled, empty);
  // 500 rows of formula text and cell slots is several tens of KB; a
  // figure that ignored the cell store would land within a few hundred
  // bytes of the empty baseline instead.
  EXPECT_GT(filled - empty, 500U * sizeof(Cell));
}

TEST(WorkbookTest, ApproximateMemoryCountsPassthroughPayloads) {
  // Passthrough carries the package's unmodelled binaries — embedded
  // images above all — so a workbook that opens a media-heavy file must
  // not look small to the host.
  Workbook wb = Workbook::create();
  const std::size_t before = wb.approximate_memory_bytes();

  constexpr std::size_t kPayloadBytes = 256U * 1024U;
  std::vector<PassthroughPart> parts;
  parts.push_back(PassthroughPart{"xl/media/image1.png", "image/png", std::vector<std::uint8_t>(kPayloadBytes, 0x7FU)});
  wb.set_passthrough_parts(std::move(parts));

  EXPECT_GE(wb.approximate_memory_bytes() - before, kPayloadBytes);
}

TEST(WorkbookTest, ApproximateMemoryIsStableWithoutMutation) {
  // Hosts report the delta against the previous reading, so a repeated
  // call on an unchanged workbook has to return the same number or the
  // accounting drifts on every poll.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(1.0))));
  const std::size_t first = wb.approximate_memory_bytes();
  EXPECT_EQ(wb.approximate_memory_bytes(), first);
}

TEST(WorkbookTest, SetDefinedNameLeavesUnrelatedSpillIntact) {
  // A defined-name edit used to route through `reindex_all_formulas`,
  // which cleared every sheet's committed spills workbook-wide regardless
  // of whether any formula actually referenced the changed name.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=SEQUENCE(3)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_NE(wb.sheet(0U).spill_region_at_anchor(0U, 0U), nullptr);

  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Unrelated", "=Sheet1!$A$1")));

  EXPECT_NE(wb.sheet(0U).spill_region_at_anchor(0U, 0U), nullptr)
      << "setting a name the spill's formula never mentions must not clear it";
}

TEST(WorkbookTest, SetDefinedNameUpdatesFormulasThatReferenceIt) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("MyName", "=1")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=MyName")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  const Cell* cell = wb.sheet(0U).cell_at(0U, 0U);
  ASSERT_NE(cell, nullptr);
  ASSERT_TRUE(cell->cached_value.is_number());
  EXPECT_DOUBLE_EQ(cell->cached_value.as_number(), 1.0);

  // Retargeting the name a formula actually depends on must still dirty
  // and re-resolve that formula, even though the reindex is now scoped to
  // formulas referencing the changed name rather than the whole workbook.
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("MyName", "=2")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  cell = wb.sheet(0U).cell_at(0U, 0U);
  ASSERT_NE(cell, nullptr);
  ASSERT_TRUE(cell->cached_value.is_number());
  EXPECT_DOUBLE_EQ(cell->cached_value.as_number(), 2.0);
}

namespace {

// Recalcs `wb` and returns A1's cached value.
Value recalc_a1(Workbook& wb) {
  EXPECT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  const Cell* cell = wb.sheet(0U).cell_at(0U, 0U);
  return cell == nullptr ? Value::blank() : cell->cached_value;
}

}  // namespace

TEST(WorkbookTest, RedefiningLambdaNameRecalcsItsCallers) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Fn", "=LAMBDA(x,x*2)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=Fn(3)")));
  Value v = recalc_a1(wb);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 6.0);

  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Fn", "=LAMBDA(x,x*10)")));
  v = recalc_a1(wb);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 30.0);
}

TEST(WorkbookTest, AddingLambdaNameResolvesAnEarlierCall) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=Fn(3)")));
  Value v = recalc_a1(wb);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);

  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Fn", "=LAMBDA(x,x*2)")));
  v = recalc_a1(wb);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 6.0);
}

TEST(WorkbookTest, RemovingLambdaNameTurnsItsCallersToNameError) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Fn", "=LAMBDA(x,x*2)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=Fn(3)")));
  ASSERT_TRUE(recalc_a1(wb).is_number());

  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Fn", "")));
  const Value v = recalc_a1(wb);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(WorkbookTest, RenamingLambdaNameRebindsItsCallers) {
  // A rename is a removal of the old spelling plus an addition of the new
  // one; a caller of the old spelling goes to #NAME?, a caller of the new
  // one resolves.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Fn", "=LAMBDA(x,x*2)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=Fn(3)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=Gn(4)")));
  ASSERT_TRUE(recalc_a1(wb).is_number());

  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Fn", "")));
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Gn", "=LAMBDA(x,x*2)")));
  const Value v = recalc_a1(wb);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
  const Cell* a2 = wb.sheet(0U).cell_at(1U, 0U);
  ASSERT_NE(a2, nullptr);
  ASSERT_TRUE(a2->cached_value.is_number());
  EXPECT_DOUBLE_EQ(a2->cached_value.as_number(), 8.0);
}

TEST(WorkbookTest, RedefiningLambdaNameRecalcsSheetQualifiedCallers) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name_scoped("Fn", "=LAMBDA(x,x*2)", 0)));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=Sheet1!Fn(3)")));
  Value v = recalc_a1(wb);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 6.0);

  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name_scoped("Fn", "=LAMBDA(x,x+1)", 0)));
  v = recalc_a1(wb);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 4.0);
}

TEST(WorkbookTest, RedefiningInnerNameRecalcsCallersOfLambdaUsingIt) {
  // Fn's body calls Inner; redefining Inner must reach a formula that only
  // mentions Fn.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Inner", "=LAMBDA(x,x*2)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Fn", "=LAMBDA(x,Inner(x)+1)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=Fn(3)")));
  Value v = recalc_a1(wb);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 7.0);

  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Inner", "=LAMBDA(x,x*3)")));
  v = recalc_a1(wb);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 10.0);
}

TEST(WorkbookTest, DefinedNameRangeEndpointTracksCellsInsideTheBox) {
  // `A1:CellNm` reads A1:C3; B2 is named by neither endpoint.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("CellNm", "=Sheet1!$C$3")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 1U, 1U, Value::number(5.0))));   // B2
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 9U, 0U, "=SUM(A1:CellNm)")));  // A10
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  const Cell* cell = wb.sheet(0U).cell_at(9U, 0U);
  ASSERT_NE(cell, nullptr);
  ASSERT_TRUE(cell->cached_value.is_number());
  EXPECT_DOUBLE_EQ(cell->cached_value.as_number(), 5.0);

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 1U, 1U, Value::number(7.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  cell = wb.sheet(0U).cell_at(9U, 0U);
  ASSERT_NE(cell, nullptr);
  ASSERT_TRUE(cell->cached_value.is_number());
  EXPECT_DOUBLE_EQ(cell->cached_value.as_number(), 7.0);
}

TEST(WorkbookTest, RedefiningWorkbookNameRecalcsSelfBookReferences) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name_scoped("G", "=100", 0)));
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("G", "=5")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=[0]!G")));
  Value v = recalc_a1(wb);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.0);

  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("G", "=6")));
  v = recalc_a1(wb);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 6.0);
}

}  // namespace
}  // namespace formulon
