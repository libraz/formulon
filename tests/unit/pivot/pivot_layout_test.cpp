// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Unit tests for pivot layout projection. These tests verify the grid shape
// exposed to frontends: absolute coordinates, cell kinds, labels, data cells,
// and totals.

#include "pivot/pivot_layout.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_evaluator.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "value.h"

namespace formulon::pivot {
namespace {

Value owned_text(PivotCache& cache, std::string s) {
  cache.mutable_text_storage().push_back(std::move(s));
  return Value::text(cache.text_storage().back());
}

PivotCache build_basic_cache() {
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Region", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Product", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});

  auto add = [&](const char* region, const char* product, double amount) {
    PivotCacheRecord rec;
    rec.cells.push_back(owned_text(cache, region));
    rec.cells.push_back(owned_text(cache, product));
    rec.cells.push_back(Value::number(amount));
    cache.mutable_records().push_back(std::move(rec));
  };

  add("North", "Widget", 100.0);
  add("North", "Gadget", 50.0);
  add("South", "Widget", 200.0);
  add("South", "Gadget", 300.0);
  add("North", "Widget", 25.0);
  return cache;
}

PivotTable build_table(std::vector<std::uint32_t> row_fields, std::vector<std::uint32_t> col_fields) {
  PivotTable table;
  table.set_name("Pivot1");
  table.set_pivot_cache_id(1);
  table.set_anchor(2, 3, 1, 1);  // D3.

  PivotField region;
  region.source_name = "Region";
  region.axis = PivotAxis::Row;
  PivotField product;
  product.source_name = "Product";
  product.axis = PivotAxis::Col;
  PivotField amount;
  amount.source_name = "Amount";
  amount.axis = PivotAxis::Value;

  table.mutable_fields().push_back(std::move(region));
  table.mutable_fields().push_back(std::move(product));
  table.mutable_fields().push_back(std::move(amount));
  table.mutable_row_field_order() = std::move(row_fields);
  table.mutable_col_field_order() = std::move(col_fields);

  PivotDataField sum_amount;
  sum_amount.name = "Sum of Amount";
  sum_amount.field_index = 2;
  sum_amount.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum_amount));
  return table;
}

PivotCache build_region_year_quarter_cache() {
  PivotCache cache;
  cache.set_cache_id(2);
  cache.mutable_fields().push_back(PivotCacheField{"Region", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Year", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Quarter", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});

  auto add = [&](const char* region, const char* year, const char* quarter, double amount) {
    PivotCacheRecord rec;
    rec.cells.push_back(owned_text(cache, region));
    rec.cells.push_back(owned_text(cache, year));
    rec.cells.push_back(owned_text(cache, quarter));
    rec.cells.push_back(Value::number(amount));
    cache.mutable_records().push_back(std::move(rec));
  };

  add("North", "2025", "Q1", 10.0);
  add("North", "2025", "Q2", 20.0);
  add("North", "2026", "Q1", 30.0);
  add("South", "2025", "Q1", 40.0);
  add("South", "2025", "Q2", 50.0);
  add("South", "2026", "Q1", 60.0);
  return cache;
}

PivotTable build_region_year_quarter_table() {
  PivotTable table;
  table.set_name("PivotMixed");
  table.set_pivot_cache_id(2);
  table.set_anchor(2, 3, 1, 1);

  PivotField region;
  region.source_name = "Region";
  region.axis = PivotAxis::Row;
  PivotField year;
  year.source_name = "Year";
  year.axis = PivotAxis::Col;
  year.subtotal_top = true;
  PivotField quarter;
  quarter.source_name = "Quarter";
  quarter.axis = PivotAxis::Col;
  PivotField amount;
  amount.source_name = "Amount";
  amount.axis = PivotAxis::Value;

  table.mutable_fields().push_back(std::move(region));
  table.mutable_fields().push_back(std::move(year));
  table.mutable_fields().push_back(std::move(quarter));
  table.mutable_fields().push_back(std::move(amount));
  table.mutable_row_field_order() = {0};
  table.mutable_col_field_order() = {1, 2};
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  PivotDataField sum_amount;
  sum_amount.name = "Sum of Amount";
  sum_amount.field_index = 3;
  sum_amount.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum_amount));
  return table;
}

PivotTable build_region_quarter_by_year_table() {
  PivotTable table = build_region_year_quarter_table();
  table.mutable_row_field_order() = {0, 2};
  table.mutable_col_field_order() = {1};
  table.mutable_fields()[0].subtotal_top = true;
  table.mutable_fields()[1].subtotal_top = false;
  return table;
}

PivotCache build_region_year_quarter_channel_cache() {
  PivotCache cache;
  cache.set_cache_id(3);
  cache.mutable_fields().push_back(PivotCacheField{"Region", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Year", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Quarter", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Channel", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});

  auto add = [&](const char* region, const char* year, const char* quarter, const char* channel, double amount) {
    PivotCacheRecord rec;
    rec.cells.push_back(owned_text(cache, region));
    rec.cells.push_back(owned_text(cache, year));
    rec.cells.push_back(owned_text(cache, quarter));
    rec.cells.push_back(owned_text(cache, channel));
    rec.cells.push_back(Value::number(amount));
    cache.mutable_records().push_back(std::move(rec));
  };

  add("North", "2025", "Q1", "Retail", 10.0);
  add("North", "2025", "Q2", "Retail", 20.0);
  add("North", "2026", "Q1", "Retail", 30.0);
  add("North", "2025", "Q1", "Online", 5.0);
  add("South", "2025", "Q1", "Retail", 40.0);
  add("South", "2025", "Q2", "Retail", 50.0);
  add("South", "2026", "Q1", "Retail", 60.0);
  add("South", "2025", "Q1", "Online", 7.0);
  return cache;
}

PivotTable build_region_channel_by_year_quarter_table() {
  PivotTable table;
  table.set_name("PivotBothSubtotals");
  table.set_pivot_cache_id(3);
  table.set_anchor(2, 3, 1, 1);

  PivotField region;
  region.source_name = "Region";
  region.axis = PivotAxis::Row;
  region.subtotal_top = true;
  PivotField year;
  year.source_name = "Year";
  year.axis = PivotAxis::Col;
  year.subtotal_top = true;
  PivotField quarter;
  quarter.source_name = "Quarter";
  quarter.axis = PivotAxis::Col;
  PivotField channel;
  channel.source_name = "Channel";
  channel.axis = PivotAxis::Row;
  PivotField amount;
  amount.source_name = "Amount";
  amount.axis = PivotAxis::Value;

  table.mutable_fields().push_back(std::move(region));
  table.mutable_fields().push_back(std::move(year));
  table.mutable_fields().push_back(std::move(quarter));
  table.mutable_fields().push_back(std::move(channel));
  table.mutable_fields().push_back(std::move(amount));
  table.mutable_row_field_order() = {0, 3};
  table.mutable_col_field_order() = {1, 2};
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  PivotDataField sum_amount;
  sum_amount.name = "Sum of Amount";
  sum_amount.field_index = 4;
  sum_amount.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum_amount));
  return table;
}

const PivotCell* find_cell(const PivotCells& cells, std::uint32_t row, std::uint32_t col) {
  for (const PivotCell& cell : cells.cells) {
    if (cell.row == row && cell.col == col) {
      return &cell;
    }
  }
  return nullptr;
}

TEST(PivotLayout, ProjectsOneRowOneColumnPivotToAbsoluteGrid) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_table(/*row=*/{0}, /*col=*/{1});

  auto result_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;

  auto cells_or = layout(table, result_or.value());
  ASSERT_TRUE(static_cast<bool>(cells_or)) << cells_or.error().message;
  const PivotCells& cells = cells_or.value();

  EXPECT_EQ(cells.top, 2U);
  EXPECT_EQ(cells.left, 3U);
  EXPECT_EQ(cells.rows, 5U);  // col labels + row field header + 2 data rows + col grand total.
  EXPECT_EQ(cells.cols, 4U);  // row label + 2 product columns + row grand total.

  const PivotCell* region_header = find_cell(cells, 3, 3);
  ASSERT_NE(region_header, nullptr);
  EXPECT_EQ(region_header->kind, PivotCellKind::Header);
  ASSERT_TRUE(region_header->value.is_text());
  EXPECT_EQ(region_header->value.as_text(), "Region");

  const PivotCell* gadget_header = find_cell(cells, 2, 4);
  ASSERT_NE(gadget_header, nullptr);
  EXPECT_EQ(gadget_header->kind, PivotCellKind::ColLabel);
  ASSERT_TRUE(gadget_header->value.is_text());
  EXPECT_EQ(gadget_header->value.as_text(), "Gadget");

  const PivotCell* north_label = find_cell(cells, 4, 3);
  ASSERT_NE(north_label, nullptr);
  EXPECT_EQ(north_label->kind, PivotCellKind::RowLabel);
  ASSERT_TRUE(north_label->value.is_text());
  EXPECT_EQ(north_label->value.as_text(), "North");

  const PivotCell* north_gadget = find_cell(cells, 4, 4);
  ASSERT_NE(north_gadget, nullptr);
  EXPECT_EQ(north_gadget->kind, PivotCellKind::Data);
  ASSERT_TRUE(north_gadget->value.is_number());
  EXPECT_DOUBLE_EQ(north_gadget->value.as_number(), 50.0);

  const PivotCell* south_widget = find_cell(cells, 5, 5);
  ASSERT_NE(south_widget, nullptr);
  EXPECT_EQ(south_widget->kind, PivotCellKind::Data);
  ASSERT_TRUE(south_widget->value.is_number());
  EXPECT_DOUBLE_EQ(south_widget->value.as_number(), 200.0);

  const PivotCell* north_total = find_cell(cells, 4, 6);
  ASSERT_NE(north_total, nullptr);
  EXPECT_EQ(north_total->kind, PivotCellKind::GrandTotal);
  ASSERT_TRUE(north_total->value.is_number());
  EXPECT_DOUBLE_EQ(north_total->value.as_number(), 175.0);
}

TEST(PivotLayout, AddsDataFieldHeaderRowForMultipleValueFields) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_table(/*row=*/{0}, /*col=*/{});
  PivotDataField count_amount;
  count_amount.name = "Count of Amount";
  count_amount.field_index = 2;
  count_amount.aggregation = Aggregation::Count;
  table.mutable_data_fields().push_back(std::move(count_amount));

  auto result_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;

  auto cells_or = layout(table, result_or.value());
  ASSERT_TRUE(static_cast<bool>(cells_or)) << cells_or.error().message;
  const PivotCells& cells = cells_or.value();

  EXPECT_EQ(cells.rows, 6U);  // value label + data-field headers + row header + 2 rows + col grand total.
  EXPECT_EQ(cells.cols, 5U);  // row label + 2 data fields + 2 grand-total fields.

  const PivotCell* sum_header = find_cell(cells, 3, 4);
  ASSERT_NE(sum_header, nullptr);
  EXPECT_EQ(sum_header->kind, PivotCellKind::Header);
  ASSERT_TRUE(sum_header->value.is_text());
  EXPECT_EQ(sum_header->value.as_text(), "Sum of Amount");

  const PivotCell* count_header = find_cell(cells, 3, 5);
  ASSERT_NE(count_header, nullptr);
  EXPECT_EQ(count_header->kind, PivotCellKind::Header);
  ASSERT_TRUE(count_header->value.is_text());
  EXPECT_EQ(count_header->value.as_text(), "Count of Amount");

  const PivotCell* north_count = find_cell(cells, 5, 5);
  ASSERT_NE(north_count, nullptr);
  EXPECT_EQ(north_count->kind, PivotCellKind::Data);
  ASSERT_TRUE(north_count->value.is_number());
  EXPECT_DOUBLE_EQ(north_count->value.as_number(), 3.0);

  const PivotCell* sum_grand_total = find_cell(cells, 7, 6);
  ASSERT_NE(sum_grand_total, nullptr);
  EXPECT_EQ(sum_grand_total->kind, PivotCellKind::GrandTotal);
  ASSERT_TRUE(sum_grand_total->value.is_number());
  EXPECT_DOUBLE_EQ(sum_grand_total->value.as_number(), 675.0);

  const PivotCell* count_grand_total = find_cell(cells, 7, 7);
  ASSERT_NE(count_grand_total, nullptr);
  EXPECT_EQ(count_grand_total->kind, PivotCellKind::GrandTotal);
  ASSERT_TRUE(count_grand_total->value.is_number());
  EXPECT_DOUBLE_EQ(count_grand_total->value.as_number(), 5.0);
}

TEST(PivotLayout, InsertsRowSubtotalRowsWhenEvaluatorProvidesMetadata) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_table(/*row=*/{0, 1}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/true);
  table.mutable_fields()[0].subtotal_top = true;

  auto result_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  ASSERT_EQ(result_or.value().row_subtotals.size(), 2U);

  auto cells_or = layout(table, result_or.value());
  ASSERT_TRUE(static_cast<bool>(cells_or)) << cells_or.error().message;
  const PivotCells& cells = cells_or.value();

  EXPECT_EQ(cells.rows, 9U);  // value label + row headers + 4 leaves + 2 subtotals + grand total.
  EXPECT_EQ(cells.cols, 3U);  // two row fields + one value field.

  const PivotCell* north_subtotal_label = find_cell(cells, 6, 3);
  ASSERT_NE(north_subtotal_label, nullptr);
  EXPECT_EQ(north_subtotal_label->kind, PivotCellKind::RowSubtotal);
  ASSERT_TRUE(north_subtotal_label->value.is_text());
  EXPECT_EQ(north_subtotal_label->value.as_text(), "North Grand Total");

  const PivotCell* north_subtotal_value = find_cell(cells, 6, 5);
  ASSERT_NE(north_subtotal_value, nullptr);
  EXPECT_EQ(north_subtotal_value->kind, PivotCellKind::RowSubtotal);
  ASSERT_TRUE(north_subtotal_value->value.is_number());
  EXPECT_DOUBLE_EQ(north_subtotal_value->value.as_number(), 175.0);

  const PivotCell* south_subtotal_value = find_cell(cells, 9, 5);
  ASSERT_NE(south_subtotal_value, nullptr);
  EXPECT_EQ(south_subtotal_value->kind, PivotCellKind::RowSubtotal);
  ASSERT_TRUE(south_subtotal_value->value.is_number());
  EXPECT_DOUBLE_EQ(south_subtotal_value->value.as_number(), 500.0);
}

TEST(PivotLayout, InsertsColSubtotalColumnsWhenEvaluatorProvidesMetadata) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_table(/*row=*/{}, /*col=*/{0, 1});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  table.mutable_fields()[0].subtotal_top = true;

  auto result_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  ASSERT_EQ(result_or.value().col_subtotals.size(), 2U);

  auto cells_or = layout(table, result_or.value());
  ASSERT_TRUE(static_cast<bool>(cells_or)) << cells_or.error().message;
  const PivotCells& cells = cells_or.value();

  EXPECT_EQ(cells.rows, 4U);  // two col-header rows + row header + one implicit row leaf.
  EXPECT_EQ(cells.cols, 7U);  // implicit row header + 4 leaves + 2 subtotal columns.

  const PivotCell* north_subtotal_header = find_cell(cells, 2, 6);
  ASSERT_NE(north_subtotal_header, nullptr);
  EXPECT_EQ(north_subtotal_header->kind, PivotCellKind::ColSubtotal);
  ASSERT_TRUE(north_subtotal_header->value.is_text());
  EXPECT_EQ(north_subtotal_header->value.as_text(), "North Grand Total");

  const PivotCell* north_subtotal_value = find_cell(cells, 5, 6);
  ASSERT_NE(north_subtotal_value, nullptr);
  EXPECT_EQ(north_subtotal_value->kind, PivotCellKind::ColSubtotal);
  ASSERT_TRUE(north_subtotal_value->value.is_number());
  EXPECT_DOUBLE_EQ(north_subtotal_value->value.as_number(), 175.0);

  const PivotCell* south_subtotal_value = find_cell(cells, 5, 9);
  ASSERT_NE(south_subtotal_value, nullptr);
  EXPECT_EQ(south_subtotal_value->kind, PivotCellKind::ColSubtotal);
  ASSERT_TRUE(south_subtotal_value->value.is_number());
  EXPECT_DOUBLE_EQ(south_subtotal_value->value.as_number(), 500.0);
}

TEST(PivotLayout, InsertsColSubtotalColumnsWithRowAxis) {
  PivotCache cache = build_region_year_quarter_cache();
  PivotTable table = build_region_year_quarter_table();

  auto result_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  ASSERT_EQ(result_or.value().col_subtotals.size(), 2U);

  auto cells_or = layout(table, result_or.value());
  ASSERT_TRUE(static_cast<bool>(cells_or)) << cells_or.error().message;
  const PivotCells& cells = cells_or.value();

  EXPECT_EQ(cells.rows, 5U);  // two col-header rows + row header + two row leaves.
  EXPECT_EQ(cells.cols, 6U);  // row label + 3 leaves + 2 subtotal columns.

  const PivotCell* north_label = find_cell(cells, 5, 3);
  ASSERT_NE(north_label, nullptr);
  EXPECT_EQ(north_label->kind, PivotCellKind::RowLabel);
  ASSERT_TRUE(north_label->value.is_text());
  EXPECT_EQ(north_label->value.as_text(), "North");

  const PivotCell* y2025_header = find_cell(cells, 2, 6);
  ASSERT_NE(y2025_header, nullptr);
  EXPECT_EQ(y2025_header->kind, PivotCellKind::ColSubtotal);
  ASSERT_TRUE(y2025_header->value.is_text());
  EXPECT_EQ(y2025_header->value.as_text(), "2025 Grand Total");

  const PivotCell* north_2025_subtotal = find_cell(cells, 5, 6);
  ASSERT_NE(north_2025_subtotal, nullptr);
  EXPECT_EQ(north_2025_subtotal->kind, PivotCellKind::ColSubtotal);
  ASSERT_TRUE(north_2025_subtotal->value.is_number());
  EXPECT_DOUBLE_EQ(north_2025_subtotal->value.as_number(), 30.0);

  const PivotCell* south_2025_subtotal = find_cell(cells, 6, 6);
  ASSERT_NE(south_2025_subtotal, nullptr);
  EXPECT_EQ(south_2025_subtotal->kind, PivotCellKind::ColSubtotal);
  ASSERT_TRUE(south_2025_subtotal->value.is_number());
  EXPECT_DOUBLE_EQ(south_2025_subtotal->value.as_number(), 90.0);

  const PivotCell* north_2026_subtotal = find_cell(cells, 5, 8);
  ASSERT_NE(north_2026_subtotal, nullptr);
  EXPECT_EQ(north_2026_subtotal->kind, PivotCellKind::ColSubtotal);
  ASSERT_TRUE(north_2026_subtotal->value.is_number());
  EXPECT_DOUBLE_EQ(north_2026_subtotal->value.as_number(), 30.0);
}

TEST(PivotLayout, InsertsRowSubtotalRowsWithColumnAxis) {
  PivotCache cache = build_region_year_quarter_cache();
  PivotTable table = build_region_quarter_by_year_table();

  auto result_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  ASSERT_EQ(result_or.value().row_subtotals.size(), 2U);

  auto cells_or = layout(table, result_or.value());
  ASSERT_TRUE(static_cast<bool>(cells_or)) << cells_or.error().message;
  const PivotCells& cells = cells_or.value();

  EXPECT_EQ(cells.rows, 8U);  // one col-header row + row header + 4 leaves + 2 subtotals.
  EXPECT_EQ(cells.cols, 4U);  // two row fields + two year columns.

  const PivotCell* north_subtotal_label = find_cell(cells, 6, 3);
  ASSERT_NE(north_subtotal_label, nullptr);
  EXPECT_EQ(north_subtotal_label->kind, PivotCellKind::RowSubtotal);
  ASSERT_TRUE(north_subtotal_label->value.is_text());
  EXPECT_EQ(north_subtotal_label->value.as_text(), "North Grand Total");

  const PivotCell* north_2025_subtotal = find_cell(cells, 6, 5);
  ASSERT_NE(north_2025_subtotal, nullptr);
  EXPECT_EQ(north_2025_subtotal->kind, PivotCellKind::RowSubtotal);
  ASSERT_TRUE(north_2025_subtotal->value.is_number());
  EXPECT_DOUBLE_EQ(north_2025_subtotal->value.as_number(), 30.0);

  const PivotCell* north_2026_subtotal = find_cell(cells, 6, 6);
  ASSERT_NE(north_2026_subtotal, nullptr);
  EXPECT_EQ(north_2026_subtotal->kind, PivotCellKind::RowSubtotal);
  ASSERT_TRUE(north_2026_subtotal->value.is_number());
  EXPECT_DOUBLE_EQ(north_2026_subtotal->value.as_number(), 30.0);

  const PivotCell* south_2025_subtotal = find_cell(cells, 9, 5);
  ASSERT_NE(south_2025_subtotal, nullptr);
  EXPECT_EQ(south_2025_subtotal->kind, PivotCellKind::RowSubtotal);
  ASSERT_TRUE(south_2025_subtotal->value.is_number());
  EXPECT_DOUBLE_EQ(south_2025_subtotal->value.as_number(), 90.0);
}

TEST(PivotLayout, InsertsSubtotalIntersectionValues) {
  PivotCache cache = build_region_year_quarter_channel_cache();
  PivotTable table = build_region_channel_by_year_quarter_table();

  auto result_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  ASSERT_EQ(result_or.value().row_subtotals.size(), 2U);
  ASSERT_EQ(result_or.value().col_subtotals.size(), 2U);
  ASSERT_EQ(result_or.value().row_subtotals[0].col_subtotal_values.size(), 2U);

  auto cells_or = layout(table, result_or.value());
  ASSERT_TRUE(static_cast<bool>(cells_or)) << cells_or.error().message;
  const PivotCells& cells = cells_or.value();

  EXPECT_EQ(cells.rows, 9U);
  EXPECT_EQ(cells.cols, 7U);

  const PivotCell* north_2025_subtotal = find_cell(cells, 7, 7);
  ASSERT_NE(north_2025_subtotal, nullptr);
  EXPECT_EQ(north_2025_subtotal->kind, PivotCellKind::RowSubtotal);
  ASSERT_TRUE(north_2025_subtotal->value.is_number());
  EXPECT_DOUBLE_EQ(north_2025_subtotal->value.as_number(), 35.0);

  const PivotCell* north_2026_subtotal = find_cell(cells, 7, 9);
  ASSERT_NE(north_2026_subtotal, nullptr);
  EXPECT_EQ(north_2026_subtotal->kind, PivotCellKind::RowSubtotal);
  ASSERT_TRUE(north_2026_subtotal->value.is_number());
  EXPECT_DOUBLE_EQ(north_2026_subtotal->value.as_number(), 30.0);

  const PivotCell* south_2025_subtotal = find_cell(cells, 10, 7);
  ASSERT_NE(south_2025_subtotal, nullptr);
  EXPECT_EQ(south_2025_subtotal->kind, PivotCellKind::RowSubtotal);
  ASSERT_TRUE(south_2025_subtotal->value.is_number());
  EXPECT_DOUBLE_EQ(south_2025_subtotal->value.as_number(), 97.0);
}

TEST(PivotLayout, RejectsMismatchedResultShape) {
  PivotTable table = build_table(/*row=*/{0}, /*col=*/{});
  PivotResult result;
  result.values.resize(2);  // row hierarchy is empty, so layout expects one implicit row leaf.

  auto cells_or = layout(table, result);
  ASSERT_FALSE(static_cast<bool>(cells_or));
  EXPECT_EQ(cells_or.error().code, FormulonErrorCode::kEvalPivotInvalid);
}

}  // namespace
}  // namespace formulon::pivot
