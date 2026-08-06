#include <cstdint>
#include <memory>
#include <string>
#include <utility>

#include "cell.h"
#include "eval/function_registry.h"
#include "eval/scheduler.h"
#include "gtest/gtest.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon::eval {
namespace {

Value OwnedText(pivot::PivotCache& cache, std::string text) {
  cache.mutable_text_storage().push_back(std::move(text));
  return Value::text(cache.text_storage().back());
}

std::unique_ptr<pivot::PivotCache> BuildCache() {
  auto cache = std::make_unique<pivot::PivotCache>();
  cache->set_cache_id(1U);

  pivot::PivotCacheField region;
  region.name = "Region";
  region.shared_items.push_back(OwnedText(*cache, "North"));
  region.shared_items.push_back(OwnedText(*cache, "South"));
  cache->mutable_fields().push_back(std::move(region));

  pivot::PivotCacheField amount;
  amount.name = "Amount";
  cache->mutable_fields().push_back(std::move(amount));

  for (const auto& [region_index, amount_value] :
       {std::pair{0.0, 100.0}, std::pair{0.0, 200.0}, std::pair{1.0, 300.0}, std::pair{1.0, 400.0}}) {
    pivot::PivotCacheRecord record;
    record.cells.push_back(Value::number(region_index));
    record.cells.push_back(Value::number(amount_value));
    cache->mutable_records().push_back(std::move(record));
  }
  return cache;
}

std::unique_ptr<pivot::PivotTable> BuildTable() {
  auto table = std::make_unique<pivot::PivotTable>();
  table->set_name("SharedPivot");
  table->set_pivot_cache_id(1U);

  pivot::PivotField region;
  region.source_name = "Region";
  region.axis = pivot::PivotAxis::Row;
  table->mutable_fields().push_back(std::move(region));

  pivot::PivotField amount;
  amount.source_name = "Amount";
  amount.axis = pivot::PivotAxis::Value;
  table->mutable_fields().push_back(std::move(amount));

  pivot::PivotDataField sum;
  sum.name = "Sum of Amount";
  sum.field_index = 1U;
  sum.aggregation = pivot::Aggregation::Sum;
  table->mutable_data_fields().push_back(std::move(sum));
  table->mutable_row_field_order().push_back(0U);
  table->set_anchor(2U, 0U, 5U, 2U);  // A3:B7
  return table;
}

Workbook BuildWorkbook() {
  Workbook workbook = Workbook::create();
  workbook.add_pivot_cache(BuildCache());
  workbook.sheet(0).add_pivot_table(BuildTable());
  return workbook;
}

Value CachedValue(const Workbook& workbook, std::uint32_t row, std::uint32_t col) {
  const Cell* cell = workbook.sheet(0).cell_at(row, col);
  return cell == nullptr ? Value::blank() : cell->cached_value;
}

TEST(PivotGetPivotDataConcurrency, ParallelInitialCachePublicationIsRaceFree) {
  // Every formula is independent in the dependency graph but resolves the
  // same pivot anchor. They therefore share a scheduler layer and race the
  // lazy result-cache publication on their first evaluation. Repeat with a
  // fresh workbook so each pass exercises that initially-empty-cache path.
  constexpr std::uint32_t kFormulaCount = 48U;
  for (std::uint32_t pass = 0U; pass < 12U; ++pass) {
    Workbook workbook = BuildWorkbook();
    for (std::uint32_t row = 0U; row < kFormulaCount; ++row) {
      ASSERT_TRUE(static_cast<bool>(
          workbook.set_cell_formula(0U, row, 1U, "=GETPIVOTDATA(\"Sum of Amount\",A3,\"Region\",\"North\")")));
    }

    SchedulerConfig config;
    config.num_threads = 8U;
    ASSERT_TRUE(static_cast<bool>(workbook.recalc_parallel(default_registry(), config, nullptr))) << "pass " << pass;

    for (std::uint32_t row = 0U; row < kFormulaCount; ++row) {
      const Value value = CachedValue(workbook, row, 1U);
      ASSERT_TRUE(value.is_number()) << "pass " << pass << " row " << row;
      EXPECT_DOUBLE_EQ(value.as_number(), 300.0) << "pass " << pass << " row " << row;
    }
  }
}

}  // namespace
}  // namespace formulon::eval
