//
// Unit tests for the Workbook skeleton. Verifies the factory shape, sheet
// accessor mutation, and the basic byte-level shape of the save() output.

#include "workbook.h"

#include <cstddef>
#include <cstdint>
#include <memory>
#include <utility>
#include <vector>

#include "cell.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "passthrough_part.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "sheet.h"
#include "utils/date_time.h"
#include "value.h"

namespace formulon {
namespace {

TEST(WorkbookTest, CreateYieldsSingleSheetNamedSheet1) {
  Workbook wb = Workbook::create();
  ASSERT_EQ(wb.sheet_count(), 1u);
  EXPECT_EQ(wb.sheet(0).name(), "Sheet1");
}

// A formula that may evaluate to an array is entered as a dynamic-array
// formula, and a literal written over it clears the mark.
TEST(WorkbookTest, FormulaEntryMarksDynamicArrayFormulas) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0U, 0U, "=SEQUENCE(2)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 1U, 0U, "=A5+1")));
  EXPECT_TRUE(wb.sheet(0).cell_at(0U, 0U)->dynamic_array);
  EXPECT_FALSE(wb.sheet(0).cell_at(1U, 0U)->dynamic_array);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0U, 0U, Value::number(1.0))));
  EXPECT_FALSE(wb.sheet(0).cell_at(0U, 0U)->dynamic_array);
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

TEST(WorkbookTest, QualifiedBuiltinCallIsTreatedAsAParseFailure) {
  // Excel refuses `=Sheet1!SUM(1)` at entry; the engine keeps the text and
  // evaluates it as a formula that failed to parse.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 1U, 0U, Value::number(4.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=Sheet1!SUM(A2)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  const Cell* cell = wb.sheet(0U).cell_at(0U, 0U);
  ASSERT_NE(cell, nullptr);
  ASSERT_TRUE(cell->cached_value.is_error());
  EXPECT_EQ(cell->cached_value.as_error(), ErrorCode::Name);
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

std::unique_ptr<pivot::PivotCache> BuildClockPivotCache() {
  auto cache = std::make_unique<pivot::PivotCache>();
  cache->set_cache_id(1U);
  cache->mutable_fields().push_back(pivot::PivotCacheField{"Date", {}});
  cache->mutable_fields().push_back(pivot::PivotCacheField{"Amount", {}});

  for (const auto [date, amount] : std::vector<std::pair<double, double>>{{100.0, 100.0}, {200.0, 200.0}}) {
    pivot::PivotCacheRecord record;
    record.cells.push_back(Value::number(date));
    record.cells.push_back(Value::number(amount));
    cache->mutable_records().push_back(std::move(record));
  }
  return cache;
}

std::unique_ptr<pivot::PivotTable> BuildClockPivotTable() {
  auto table = std::make_unique<pivot::PivotTable>();
  table->set_name("ClockPivot");
  table->set_pivot_cache_id(1U);
  table->set_anchor(0U, 0U, 1U, 1U);

  pivot::PivotField date_field;
  date_field.source_name = "Date";
  table->mutable_fields().push_back(std::move(date_field));

  pivot::PivotField amount_field;
  amount_field.source_name = "Amount";
  amount_field.axis = pivot::PivotAxis::Value;
  table->mutable_fields().push_back(std::move(amount_field));

  pivot::PivotDataField amount_data;
  amount_data.name = "Sum of Amount";
  amount_data.field_index = 1U;
  amount_data.aggregation = pivot::Aggregation::Sum;
  table->mutable_data_fields().push_back(std::move(amount_data));

  pivot::AuthoredPeriodFilter period;
  period.field_index = 0U;
  period.period = pivot::RelativePeriod::ThisQuarter;
  table->mutable_authored_period_filters().push_back(period);
  return table;
}

Workbook BuildClockPivotWorkbook() {
  Workbook wb = Workbook::create();
  wb.add_pivot_cache(BuildClockPivotCache());
  wb.sheet(0U).add_pivot_table(BuildClockPivotTable());
  return wb;
}

pivot::PivotTable* ClockPivot(Workbook& wb) {
  const auto& tables = wb.sheet(0U).pivot_tables();
  return tables.empty() ? nullptr : tables.front().get();
}

Value RecalcCell(Workbook& wb, std::uint32_t row, std::uint32_t col) {
  EXPECT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  const Cell* cell = wb.sheet(0U).cell_at(row, col);
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

TEST(WorkbookTest, ChangingExcelProfileRecalculatesCachedFormulasAndDependents) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::text("あ"))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 1U, 0U, Value::text("ア"))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=COUNTIF(A1:A2,\"あ\")")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 1U, "=B1*10")));

  const auto expect_values = [&](double count) {
    const Cell* root = wb.sheet(0U).cell_at(0U, 1U);
    const Cell* dependent = wb.sheet(0U).cell_at(1U, 1U);
    ASSERT_NE(root, nullptr);
    ASSERT_NE(dependent, nullptr);
    ASSERT_TRUE(root->cached_value.is_number());
    ASSERT_TRUE(dependent->cached_value.is_number());
    EXPECT_DOUBLE_EQ(root->cached_value.as_number(), count);
    EXPECT_DOUBLE_EQ(dependent->cached_value.as_number(), count * 10.0);
  };
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  expect_values(1.0);

  wb.set_excel_profile(mac_365_ja_jp_profile());
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  expect_values(2.0);

  wb.set_excel_profile(win_365_ja_jp_profile());
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  expect_values(1.0);

  wb.set_excel_profile(win_365_ja_jp_profile());
  const auto unchanged = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(unchanged));
  EXPECT_EQ(unchanged.value().cells_evaluated, 0U);
  expect_values(1.0);
}

TEST(WorkbookTest, ChangingExcelProfileInvalidatesPivotAndFormulaCaches) {
  Workbook wb = BuildClockPivotWorkbook();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=GETPIVOTDATA(\"Sum of Amount\",A1)")));
  wb.set_pinned_now(date_time::CivilTime{{1900, 5U, 15U}, {0U, 0U, 0U}});
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  pivot::PivotTable* table = ClockPivot(wb);
  ASSERT_NE(table, nullptr);
  const auto before = table->last_result();
  ASSERT_NE(before, nullptr);
  table->mark_span_authored();
  ASSERT_TRUE(table->has_authored_span());

  wb.set_excel_profile(mac_365_ja_jp_profile());
  EXPECT_EQ(table->last_result(), nullptr);
  EXPECT_FALSE(table->has_authored_span());
  auto changed = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(changed));
  EXPECT_EQ(changed.value().cells_evaluated, 1U);
  const auto after = table->last_result();
  ASSERT_NE(after, nullptr);
  EXPECT_NE(after, before);

  table->mark_span_authored();
  ASSERT_TRUE(table->has_authored_span());
  wb.set_excel_profile(mac_365_ja_jp_profile());
  auto unchanged = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(unchanged));
  EXPECT_EQ(unchanged.value().cells_evaluated, 0U);
  EXPECT_EQ(table->last_result(), after);
  EXPECT_TRUE(table->has_authored_span());
}

TEST(WorkbookTest, RePinningNowInvalidatesRelativePivotAndFormulaCaches) {
  Workbook wb = BuildClockPivotWorkbook();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=GETPIVOTDATA(\"Sum of Amount\",A1)")));
  const date_time::CivilTime may{{1900, 5U, 15U}, {6U, 7U, 8U}};
  const date_time::CivilTime august{{1900, 8U, 15U}, {6U, 7U, 8U}};
  wb.set_pinned_now(may);

  Value result = RecalcCell(wb, 0U, 1U);
  ASSERT_TRUE(result.is_number());
  EXPECT_DOUBLE_EQ(result.as_number(), 100.0);
  pivot::PivotTable* table = ClockPivot(wb);
  ASSERT_NE(table, nullptr);
  const auto may_snapshot = table->last_result();
  ASSERT_NE(may_snapshot, nullptr);
  table->mark_span_authored();
  ASSERT_TRUE(table->has_authored_span());

  wb.set_pinned_now(august);
  EXPECT_FALSE(table->has_authored_span());
  result = RecalcCell(wb, 0U, 1U);
  ASSERT_TRUE(result.is_number());
  EXPECT_DOUBLE_EQ(result.as_number(), 200.0);
  const auto august_snapshot = table->last_result();
  ASSERT_NE(august_snapshot, nullptr);
  EXPECT_NE(august_snapshot, may_snapshot);

  // All six civil fields are unchanged, so setting the same pin again is a
  // no-op: the cached pivot and its dependent formula remain clean.
  wb.set_pinned_now(august);
  auto no_op_stats = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(no_op_stats));
  EXPECT_EQ(no_op_stats.value().cells_evaluated, 0U);
  EXPECT_EQ(table->last_result(), august_snapshot);
}

TEST(WorkbookTest, DateEpochChangeInvalidatesDateFormulaAndPivotCache) {
  Workbook wb = BuildClockPivotWorkbook();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=GETPIVOTDATA(\"Sum of Amount\",A1)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 2U, "=DATE(2026,1,1)")));
  wb.set_pinned_now(date_time::CivilTime{{1900, 5U, 15U}, {0U, 0U, 0U}});

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  const Cell* date_cell = wb.sheet(0U).cell_at(0U, 2U);
  ASSERT_NE(date_cell, nullptr);
  ASSERT_TRUE(date_cell->cached_value.is_number());
  const double serial_1900 = date_cell->cached_value.as_number();
  EXPECT_DOUBLE_EQ(serial_1900, date_time::serial_from_ymd(2026, 1U, 1U, false));
  pivot::PivotTable* table = ClockPivot(wb);
  ASSERT_NE(table, nullptr);
  ASSERT_NE(table->last_result(), nullptr);
  table->mark_span_authored();
  ASSERT_TRUE(table->has_authored_span());

  wb.set_date1904(true);
  EXPECT_EQ(table->last_result(), nullptr);
  EXPECT_FALSE(table->has_authored_span());
  auto epoch_stats = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(epoch_stats));
  EXPECT_EQ(epoch_stats.value().cells_evaluated, 2U);
  date_cell = wb.sheet(0U).cell_at(0U, 2U);
  ASSERT_NE(date_cell, nullptr);
  ASSERT_TRUE(date_cell->cached_value.is_number());
  EXPECT_DOUBLE_EQ(date_cell->cached_value.as_number(), serial_1900 - date_time::kDate1904EpochGap);

  // Re-applying the effective epoch is a no-op and preserves the refreshed
  // pivot result and clean formula cache.
  const auto refreshed_snapshot = table->last_result();
  ASSERT_NE(refreshed_snapshot, nullptr);
  wb.set_date1904(true);
  auto no_op_stats = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(no_op_stats));
  EXPECT_EQ(no_op_stats.value().cells_evaluated, 0U);
  EXPECT_EQ(table->last_result(), refreshed_snapshot);
}

TEST(WorkbookTest, ClearingPinnedNowInvalidatesOnceAndDuplicateClearIsNoOp) {
  Workbook wb = BuildClockPivotWorkbook();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=GETPIVOTDATA(\"Sum of Amount\",A1)")));
  wb.set_pinned_now(date_time::CivilTime{{1900, 5U, 15U}, {0U, 0U, 0U}});
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  const pivot::PivotTable* table = ClockPivot(wb);
  ASSERT_NE(table, nullptr);
  ASSERT_NE(table->last_result(), nullptr);

  wb.clear_pinned_now();
  EXPECT_EQ(table->last_result(), nullptr);
  auto unpin_stats = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(unpin_stats));
  EXPECT_EQ(unpin_stats.value().cells_evaluated, 1U);
  const auto host_snapshot = table->last_result();
  ASSERT_NE(host_snapshot, nullptr);

  wb.clear_pinned_now();
  auto duplicate_clear_stats = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(duplicate_clear_stats));
  EXPECT_EQ(duplicate_clear_stats.value().cells_evaluated, 0U);
  EXPECT_EQ(table->last_result(), host_snapshot);
}

}  // namespace
}  // namespace formulon
