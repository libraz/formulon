// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Integration test for the OOXML writer's pivot wiring. Builds a
// `Workbook` in memory, drives it through `io::write_ooxml`, then pipes
// the resulting bytes back through `io::read_ooxml` and asserts the
// pivot caches and pivot tables survive intact. The end-to-end shape
// also validates that `pivot::evaluate` produces the expected aggregates
// against the reread workbook -- mirroring the read-only integration
// test in `ooxml_pivot_test.cpp`, but going through the writer side.
//
// Companion to `tests/integration/ooxml_pivot_test.cpp` (which exercises
// hand-built bytes against the reader). The writer-built path here
// closes the `read -> write -> read` round trip on the pivot family.

#include <cstdint>
#include <memory>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_evaluator.h"
#include "pivot/pivot_layout.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

// Convenience: appends `s` to `cache.mutable_text_storage()` and returns
// a `Value::text` aliasing that storage. Mirrors how the reader hands
// out text payloads, so values built this way live as long as `cache`.
Value MakeText(pivot::PivotCache& cache, std::string s) {
  cache.mutable_text_storage().push_back(std::move(s));
  return Value::text(cache.mutable_text_storage().back());
}

io::ByteSpan SpanOf(const std::vector<std::uint8_t>& bytes) {
  return io::ByteSpan{bytes.data(), bytes.size()};
}

// Builds the canonical (Region, Amount) cache used by several tests
// below. Three records: North/100, South/200, North/300. Returns an
// owning `PivotCache`; the caller transfers it to a workbook via
// `add_pivot_cache(make_unique<PivotCache>(std::move(...)))`.
pivot::PivotCache BuildRegionAmountCache(std::uint32_t cache_id) {
  pivot::PivotCache cache;
  cache.set_cache_id(cache_id);

  // Field 0: shared-items text field "Region" with two values.
  pivot::PivotCacheField region;
  region.name = "Region";
  region.shared_items.push_back(MakeText(cache, "North"));
  region.shared_items.push_back(MakeText(cache, "South"));
  cache.mutable_fields().push_back(std::move(region));

  // Field 1: range-typed numeric field "Amount" with empty shared_items.
  pivot::PivotCacheField amount;
  amount.name = "Amount";
  cache.mutable_fields().push_back(std::move(amount));

  // Three records: index-into-shared-items for Region, inline number
  // for Amount.
  pivot::PivotCacheRecord r0;
  r0.cells.push_back(Value::number(0.0));  // index 0 -> "North"
  r0.cells.push_back(Value::number(100.0));
  cache.mutable_records().push_back(std::move(r0));
  pivot::PivotCacheRecord r1;
  r1.cells.push_back(Value::number(1.0));  // index 1 -> "South"
  r1.cells.push_back(Value::number(200.0));
  cache.mutable_records().push_back(std::move(r1));
  pivot::PivotCacheRecord r2;
  r2.cells.push_back(Value::number(0.0));  // index 0 -> "North"
  r2.cells.push_back(Value::number(300.0));
  cache.mutable_records().push_back(std::move(r2));

  return cache;
}

// Builds a Sum-of-Amount-by-Region pivot table bound to `cache_id`,
// anchored at (anchor_row, anchor_col).
std::unique_ptr<pivot::PivotTable> BuildRegionAmountTable(std::uint32_t cache_id, std::uint32_t anchor_row,
                                                          std::uint32_t anchor_col) {
  auto table = std::make_unique<pivot::PivotTable>();
  table->set_name("PivotTable1");
  table->set_pivot_cache_id(cache_id);
  table->set_anchor(anchor_row, anchor_col, 5U, 2U);

  pivot::PivotField region;
  region.axis = pivot::PivotAxis::Row;
  region.custom_name = "Region";
  region.items.push_back(pivot::PivotItem{"", true});
  region.items.push_back(pivot::PivotItem{"", true});
  table->mutable_fields().push_back(std::move(region));

  pivot::PivotField amount;
  amount.axis = pivot::PivotAxis::Value;
  amount.custom_name = "Amount";
  table->mutable_fields().push_back(std::move(amount));

  table->mutable_row_field_order().push_back(0U);

  pivot::PivotDataField sum_amount;
  sum_amount.name = "Sum of Amount";
  sum_amount.field_index = 1U;
  sum_amount.aggregation = pivot::Aggregation::Sum;
  table->mutable_data_fields().push_back(std::move(sum_amount));
  return table;
}

// ---------------------------------------------------------------------------
// 1. Pivot-free workbook produces no pivot parts on the round trip.
// ---------------------------------------------------------------------------

TEST(OoxmlWriterPivot, EmptyWorkbookHasNoPivotParts) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");

  auto bytes_or = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << "write_ooxml: " << bytes_or.error().message;

  auto read_or = io::read_ooxml(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << "read_ooxml: " << read_or.error().message;

  const Workbook& reloaded = read_or.value().workbook;
  EXPECT_EQ(reloaded.sheet_count(), 1U);
  EXPECT_TRUE(reloaded.pivot_caches().empty());
  for (std::size_t i = 0; i < reloaded.sheet_count(); ++i) {
    EXPECT_TRUE(reloaded.sheet(i).pivot_tables().empty()) << "sheet " << i << " has unexpected pivot tables";
  }
}

// ---------------------------------------------------------------------------
// 2. Single pivot cache with no pivot tables survives the round trip.
// ---------------------------------------------------------------------------

TEST(OoxmlWriterPivot, SinglePivotCacheRoundTrips) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  wb.add_pivot_cache(std::make_unique<pivot::PivotCache>(BuildRegionAmountCache(/*cache_id=*/0U)));

  auto bytes_or = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << "write_ooxml: " << bytes_or.error().message;

  auto read_or = io::read_ooxml(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << "read_ooxml: " << read_or.error().message;

  const Workbook& reloaded = read_or.value().workbook;
  ASSERT_EQ(reloaded.pivot_caches().size(), 1U);
  const pivot::PivotCache* cache = reloaded.pivot_caches()[0].get();
  ASSERT_NE(cache, nullptr);
  EXPECT_EQ(cache->cache_id(), 0U);
  ASSERT_EQ(cache->fields().size(), 2U);
  EXPECT_EQ(cache->fields()[0].name, "Region");
  ASSERT_EQ(cache->fields()[0].shared_items.size(), 2U);
  EXPECT_EQ(cache->fields()[0].shared_items[0].as_text(), "North");
  EXPECT_EQ(cache->fields()[0].shared_items[1].as_text(), "South");
  EXPECT_EQ(cache->fields()[1].name, "Amount");
  EXPECT_TRUE(cache->fields()[1].shared_items.empty());

  ASSERT_EQ(cache->records().size(), 3U);
  EXPECT_EQ(cache->records()[0].cells[0].as_text(), "North");
  EXPECT_DOUBLE_EQ(cache->records()[0].cells[1].as_number(), 100.0);
  EXPECT_EQ(cache->records()[1].cells[0].as_text(), "South");
  EXPECT_DOUBLE_EQ(cache->records()[1].cells[1].as_number(), 200.0);
  EXPECT_EQ(cache->records()[2].cells[0].as_text(), "North");
  EXPECT_DOUBLE_EQ(cache->records()[2].cells[1].as_number(), 300.0);

  EXPECT_TRUE(reloaded.sheet(0).pivot_tables().empty());
}

// ---------------------------------------------------------------------------
// 3. Single pivot table on Sheet2 survives, anchored at D1.
// ---------------------------------------------------------------------------

TEST(OoxmlWriterPivot, SinglePivotTableRoundTrips) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  wb.add_sheet("Sheet2");
  wb.add_pivot_cache(std::make_unique<pivot::PivotCache>(BuildRegionAmountCache(/*cache_id=*/0U)));
  wb.sheet(1).add_pivot_table(BuildRegionAmountTable(/*cache_id=*/0U, /*anchor_row=*/0U, /*anchor_col=*/3U));

  auto bytes_or = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << "write_ooxml: " << bytes_or.error().message;

  auto read_or = io::read_ooxml(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << "read_ooxml: " << read_or.error().message;

  const Workbook& reloaded = read_or.value().workbook;
  ASSERT_EQ(reloaded.pivot_caches().size(), 1U);
  const pivot::PivotCache* cache = reloaded.pivot_caches()[0].get();
  ASSERT_NE(cache, nullptr);
  EXPECT_EQ(cache->cache_id(), 0U);

  ASSERT_EQ(reloaded.sheet_count(), 2U);
  EXPECT_TRUE(reloaded.sheet(0).pivot_tables().empty());
  ASSERT_EQ(reloaded.sheet(1).pivot_tables().size(), 1U);
  const pivot::PivotTable* table = reloaded.sheet(1).pivot_tables()[0].get();
  ASSERT_NE(table, nullptr);
  EXPECT_EQ(table->name(), "PivotTable1");
  EXPECT_EQ(table->pivot_cache_id(), cache->cache_id());
  EXPECT_EQ(table->anchor_row(), 0U);
  EXPECT_EQ(table->anchor_col(), 3U);
  ASSERT_EQ(table->row_field_order().size(), 1U);
  EXPECT_EQ(table->row_field_order()[0], 0U);
  ASSERT_EQ(table->data_fields().size(), 1U);
  EXPECT_EQ(table->data_fields()[0].name, "Sum of Amount");
}

// ---------------------------------------------------------------------------
// 4. End-to-end: writer-produced workbook re-evaluates to the same
//    aggregates as the canonical read-only integration test.
// ---------------------------------------------------------------------------

TEST(OoxmlWriterPivot, EvaluatorRoundTripProducesSameAggregates) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  wb.add_sheet("Sheet2");
  wb.add_pivot_cache(std::make_unique<pivot::PivotCache>(BuildRegionAmountCache(/*cache_id=*/0U)));
  wb.sheet(1).add_pivot_table(BuildRegionAmountTable(/*cache_id=*/0U, /*anchor_row=*/0U, /*anchor_col=*/3U));

  auto bytes_or = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << "write_ooxml: " << bytes_or.error().message;
  auto read_or = io::read_ooxml(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << "read_ooxml: " << read_or.error().message;

  const Workbook& reloaded = read_or.value().workbook;
  ASSERT_EQ(reloaded.pivot_caches().size(), 1U);
  const pivot::PivotCache* cache = reloaded.pivot_caches()[0].get();
  ASSERT_NE(cache, nullptr);
  ASSERT_EQ(reloaded.sheet(1).pivot_tables().size(), 1U);
  const pivot::PivotTable* table = reloaded.sheet(1).pivot_tables()[0].get();
  ASSERT_NE(table, nullptr);

  auto eval_or = pivot::evaluate(*table, *cache);
  ASSERT_TRUE(static_cast<bool>(eval_or)) << "pivot::evaluate: " << eval_or.error().message;
  const pivot::PivotResult& result = eval_or.value();

  // North = 100 + 300 = 400, South = 200, grand = 600.
  ASSERT_EQ(result.rows.size(), 2U);
  EXPECT_EQ(result.rows[0].label, "North");
  EXPECT_EQ(result.rows[1].label, "South");
  ASSERT_EQ(result.values.size(), 2U);
  ASSERT_EQ(result.values[0].size(), 1U);
  ASSERT_EQ(result.values[0][0].size(), 1U);
  ASSERT_EQ(result.values[1].size(), 1U);
  ASSERT_EQ(result.values[1][0].size(), 1U);
  EXPECT_TRUE(result.values[0][0][0].is_number());
  EXPECT_DOUBLE_EQ(result.values[0][0][0].as_number(), 400.0);
  EXPECT_TRUE(result.values[1][0][0].is_number());
  EXPECT_DOUBLE_EQ(result.values[1][0][0].as_number(), 200.0);
  EXPECT_TRUE(result.grand_total.is_number());
  EXPECT_DOUBLE_EQ(result.grand_total.as_number(), 600.0);
}

// ---------------------------------------------------------------------------
// 5. The frontend-facing layout projection survives the same
//    write -> read path: absolute cells, semantic cell kinds, and
//    aggregate values remain stable after reloading from OOXML.
// ---------------------------------------------------------------------------

TEST(OoxmlWriterPivot, LayoutProjectionRoundTripsForGridRendering) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  wb.add_sheet("Sheet2");
  wb.add_pivot_cache(std::make_unique<pivot::PivotCache>(BuildRegionAmountCache(/*cache_id=*/0U)));
  wb.sheet(1).add_pivot_table(BuildRegionAmountTable(/*cache_id=*/0U, /*anchor_row=*/0U, /*anchor_col=*/3U));

  auto bytes_or = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << "write_ooxml: " << bytes_or.error().message;
  auto read_or = io::read_ooxml(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << "read_ooxml: " << read_or.error().message;

  const Workbook& reloaded = read_or.value().workbook;
  ASSERT_EQ(reloaded.pivot_caches().size(), 1U);
  const pivot::PivotCache* cache = reloaded.pivot_caches()[0].get();
  ASSERT_NE(cache, nullptr);
  ASSERT_EQ(reloaded.sheet(1).pivot_tables().size(), 1U);
  const pivot::PivotTable* table = reloaded.sheet(1).pivot_tables()[0].get();
  ASSERT_NE(table, nullptr);

  auto eval_or = pivot::evaluate(*table, *cache);
  ASSERT_TRUE(static_cast<bool>(eval_or)) << "pivot::evaluate: " << eval_or.error().message;
  auto layout_or = pivot::layout(*table, eval_or.value());
  ASSERT_TRUE(static_cast<bool>(layout_or)) << "pivot::layout: " << layout_or.error().message;
  const pivot::PivotCells& cells = layout_or.value();

  EXPECT_EQ(cells.top, 0U);
  EXPECT_EQ(cells.left, 3U);
  EXPECT_EQ(cells.rows, 5U);
  EXPECT_EQ(cells.cols, 3U);

  const auto find_cell = [&](std::uint32_t row, std::uint32_t col) -> const pivot::PivotCell* {
    for (const pivot::PivotCell& cell : cells.cells) {
      if (cell.row == row && cell.col == col) {
        return &cell;
      }
    }
    return nullptr;
  };

  const pivot::PivotCell* region = find_cell(1U, 3U);
  ASSERT_NE(region, nullptr);
  EXPECT_EQ(region->kind, pivot::PivotCellKind::Header);
  ASSERT_TRUE(region->value.is_text());
  EXPECT_EQ(region->value.as_text(), "Region");

  const pivot::PivotCell* north_sum = find_cell(2U, 4U);
  ASSERT_NE(north_sum, nullptr);
  EXPECT_EQ(north_sum->kind, pivot::PivotCellKind::Data);
  ASSERT_TRUE(north_sum->value.is_number());
  EXPECT_DOUBLE_EQ(north_sum->value.as_number(), 400.0);
  EXPECT_EQ(north_sum->field_name, "Sum of Amount");

  const pivot::PivotCell* grand_total = find_cell(4U, 5U);
  ASSERT_NE(grand_total, nullptr);
  EXPECT_EQ(grand_total->kind, pivot::PivotCellKind::GrandTotal);
  ASSERT_TRUE(grand_total->value.is_number());
  EXPECT_DOUBLE_EQ(grand_total->value.as_number(), 600.0);
}

// ---------------------------------------------------------------------------
// 6. Two pivot caches survive with distinct cache_ids and distinct
//    package paths. We verify both ends of that contract via the reader:
//    the reloaded workbook reports two caches whose cache_ids match the
//    originals, and each cache carries its own field/record payload.
//    Any path-numbering bug (collision, reuse, off-by-one) would corrupt
//    the reload because the writer emits one Override / rels-entry per
//    plan entry; the reader's strict zip-entry lookups would surface it.
// ---------------------------------------------------------------------------

TEST(OoxmlWriterPivot, MultiCacheRoundTrips) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");

  // Cache 0: the canonical Region/Amount cache.
  wb.add_pivot_cache(std::make_unique<pivot::PivotCache>(BuildRegionAmountCache(/*cache_id=*/0U)));

  // Cache 1: a different shape -- single shared-items text field
  // "Product" with three distinct values, plus a numeric "Qty" field.
  pivot::PivotCache cache_b;
  cache_b.set_cache_id(7U);
  pivot::PivotCacheField product;
  product.name = "Product";
  product.shared_items.push_back(MakeText(cache_b, "Apple"));
  product.shared_items.push_back(MakeText(cache_b, "Banana"));
  product.shared_items.push_back(MakeText(cache_b, "Cherry"));
  cache_b.mutable_fields().push_back(std::move(product));
  pivot::PivotCacheField qty;
  qty.name = "Qty";
  cache_b.mutable_fields().push_back(std::move(qty));
  pivot::PivotCacheRecord rec;
  rec.cells.push_back(Value::number(2.0));  // -> "Cherry"
  rec.cells.push_back(Value::number(42.0));
  cache_b.mutable_records().push_back(std::move(rec));
  wb.add_pivot_cache(std::make_unique<pivot::PivotCache>(std::move(cache_b)));

  auto bytes_or = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << "write_ooxml: " << bytes_or.error().message;
  auto read_or = io::read_ooxml(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << "read_ooxml: " << read_or.error().message;

  const Workbook& reloaded = read_or.value().workbook;
  ASSERT_EQ(reloaded.pivot_caches().size(), 2U);

  // Ordering follows document order in the workbook part. cache_id is
  // preserved per entry and must match the originals (0 and 7).
  const pivot::PivotCache* a = reloaded.pivot_caches()[0].get();
  const pivot::PivotCache* b = reloaded.pivot_caches()[1].get();
  ASSERT_NE(a, nullptr);
  ASSERT_NE(b, nullptr);
  EXPECT_EQ(a->cache_id(), 0U);
  EXPECT_EQ(b->cache_id(), 7U);

  // Cache A: Region/Amount as before.
  ASSERT_EQ(a->fields().size(), 2U);
  EXPECT_EQ(a->fields()[0].name, "Region");
  ASSERT_EQ(a->fields()[0].shared_items.size(), 2U);
  EXPECT_EQ(a->fields()[1].name, "Amount");
  EXPECT_EQ(a->records().size(), 3U);

  // Cache B: Product/Qty.
  ASSERT_EQ(b->fields().size(), 2U);
  EXPECT_EQ(b->fields()[0].name, "Product");
  ASSERT_EQ(b->fields()[0].shared_items.size(), 3U);
  EXPECT_EQ(b->fields()[0].shared_items[0].as_text(), "Apple");
  EXPECT_EQ(b->fields()[0].shared_items[1].as_text(), "Banana");
  EXPECT_EQ(b->fields()[0].shared_items[2].as_text(), "Cherry");
  EXPECT_EQ(b->fields()[1].name, "Qty");
  ASSERT_EQ(b->records().size(), 1U);
  EXPECT_EQ(b->records()[0].cells[0].as_text(), "Cherry");
  EXPECT_DOUBLE_EQ(b->records()[0].cells[1].as_number(), 42.0);
}

}  // namespace
}  // namespace formulon
