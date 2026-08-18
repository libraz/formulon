//
// Unit tests for pivot layout projection. These tests verify the grid shape
// exposed to frontends: absolute coordinates, cell kinds, labels, data cells,
// and totals.

#include "pivot/pivot_layout.h"

#include <cstddef>
#include <cstdint>
#include <optional>
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

const PivotCell* find_cell(const PivotCells& cells, std::uint32_t row, std::uint32_t col);

PivotTable build_deep_layout_table() {
  PivotTable table;
  table.set_name("PivotDeepLayout");
  table.set_pivot_cache_id(99);
  table.set_anchor(0, 0, 1, 1);
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  for (const char* name : {"Level 1", "Level 2", "Level 3", "Level 4"}) {
    PivotField field;
    field.source_name = name;
    field.axis = PivotAxis::Row;
    table.mutable_fields().push_back(std::move(field));
  }
  PivotField amount;
  amount.source_name = "Amount";
  amount.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(amount));
  table.mutable_row_field_order() = {0, 1, 2, 3};

  PivotDataField sum_amount;
  sum_amount.name = "Sum of Amount";
  sum_amount.field_index = 4;
  sum_amount.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum_amount));
  return table;
}

PivotResult build_deep_layout_result(bool multiple_subtotals) {
  auto leaf = [](const char* label) { return RowHierarchyNode{label, {}}; };
  RowHierarchyNode a1x{"A1x", {leaf("A1x-1"), leaf("A1x-2")}};
  RowHierarchyNode a1y{"A1y", {leaf("A1y-1")}};
  RowHierarchyNode a1{"A1", {std::move(a1x), std::move(a1y)}};
  RowHierarchyNode a2x{"A2x", {leaf("A2x-1")}};
  RowHierarchyNode a2{"A2", {std::move(a2x)}};
  RowHierarchyNode b1x{"B1x", {leaf("B1x-1")}};
  RowHierarchyNode b1{"B1", {std::move(b1x)}};

  PivotResult result;
  result.rows = {RowHierarchyNode{"A", {std::move(a1), std::move(a2)}}, RowHierarchyNode{"B", {std::move(b1)}}};
  for (double value : {1.0, 2.0, 3.0, 4.0, 5.0}) {
    result.values.push_back({{Value::number(value)}});
  }

  auto add_subtotal = [&](std::vector<std::string> labels, std::uint32_t depth, double value) {
    RowSubtotal subtotal;
    subtotal.labels = std::move(labels);
    subtotal.depth = depth;
    subtotal.values.push_back(Value::number(value));
    result.row_subtotals.push_back(std::move(subtotal));
  };

  // The evaluator's contract is post-order: descendants are emitted before
  // the current node, and custom functions stay contiguous for one node.
  if (multiple_subtotals) {
    add_subtotal({"A", "A1", "A1x"}, 2, 201.0);
    add_subtotal({"A", "A1", "A1x"}, 2, 202.0);
    add_subtotal({"A", "A1", "A1y"}, 2, 203.0);
    add_subtotal({"A", "A1", "A1y"}, 2, 211.0);
    add_subtotal({"A", "A1"}, 1, 204.0);
    add_subtotal({"A", "A2", "A2x"}, 2, 205.0);
    add_subtotal({"A", "A2", "A2x"}, 2, 212.0);
    add_subtotal({"A", "A2"}, 1, 206.0);
    add_subtotal({"A"}, 0, 207.0);
    add_subtotal({"B", "B1", "B1x"}, 2, 208.0);
    add_subtotal({"B", "B1", "B1x"}, 2, 213.0);
    add_subtotal({"B", "B1"}, 1, 209.0);
    add_subtotal({"B"}, 0, 210.0);
  } else {
    add_subtotal({"A", "A1", "A1x"}, 2, 101.0);
    add_subtotal({"A", "A1", "A1y"}, 2, 102.0);
    add_subtotal({"A", "A1"}, 1, 103.0);
    add_subtotal({"A", "A2", "A2x"}, 2, 104.0);
    add_subtotal({"A", "A2"}, 1, 105.0);
    add_subtotal({"A"}, 0, 106.0);
    add_subtotal({"B", "B1", "B1x"}, 2, 107.0);
    add_subtotal({"B", "B1"}, 1, 108.0);
    add_subtotal({"B"}, 0, 109.0);
  }
  return result;
}

std::vector<double> projected_data_values(const PivotCells& cells, std::uint32_t data_col) {
  std::vector<double> values;
  for (std::uint32_t row = cells.top + 1; row < cells.top + cells.rows; ++row) {
    const PivotCell* cell = find_cell(cells, row, data_col);
    if (cell == nullptr) {
      continue;
    }
    EXPECT_TRUE(cell->value.is_number());
    if (cell->value.is_number()) {
      values.push_back(cell->value.as_number());
    }
  }
  return values;
}

std::vector<PivotCellKind> projected_data_kinds(const PivotCells& cells, std::uint32_t data_col) {
  std::vector<PivotCellKind> kinds;
  for (std::uint32_t row = cells.top + 1; row < cells.top + cells.rows; ++row) {
    const PivotCell* cell = find_cell(cells, row, data_col);
    if (cell != nullptr) {
      kinds.push_back(cell->kind);
    }
  }
  return kinds;
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

TEST(PivotLayout, CompactLayoutHonorsFieldSubtotalTop) {
  PivotCache cache = build_basic_cache();
  cache.mutable_records().erase(cache.mutable_records().begin() + 2, cache.mutable_records().begin() + 4);
  // Keep just the North group.
  PivotTable table = build_table(/*row=*/{0, 1}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  table.mutable_fields()[0].subtotal_top = false;

  auto result_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  ASSERT_EQ(result_or.value().row_subtotals.size(), 1U);

  PivotLayoutOptions options;
  options.row_labels_label = "Row Labels";
  auto cells_or = layout(table, result_or.value(), options);
  ASSERT_TRUE(static_cast<bool>(cells_or)) << cells_or.error().message;
  const PivotCells& cells = cells_or.value();

  const PivotCell* north_subtotal = nullptr;
  std::uint32_t north_last_detail_row = 0;
  for (const PivotCell& cell : cells.cells) {
    if (cell.kind != PivotCellKind::RowLabel && cell.kind != PivotCellKind::RowSubtotal) {
      continue;
    }
    if (!cell.value.is_text()) {
      continue;
    }
    if (cell.kind == PivotCellKind::RowSubtotal && cell.value.as_text() == "North") {
      north_subtotal = &cell;
    }
    if (cell.kind == PivotCellKind::RowLabel &&
        (cell.value.as_text() == "Gadget" || cell.value.as_text() == "Widget")) {
      north_last_detail_row = std::max(north_last_detail_row, cell.row);
    }
  }
  ASSERT_NE(north_subtotal, nullptr);
  EXPECT_GT(north_subtotal->row, north_last_detail_row);
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

TEST(PivotLayout, ProjectsEveryCustomColumnSubtotal) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_table(/*row=*/{}, /*col=*/{0, 1});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  table.mutable_fields()[0].subtotal_fns = {SubtotalFn::Average, SubtotalFn::Max};

  auto result_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  ASSERT_EQ(result_or.value().col_subtotals.size(), 4U);

  auto cells_or = layout(table, result_or.value());
  ASSERT_TRUE(static_cast<bool>(cells_or)) << cells_or.error().message;
  const PivotCells& cells = cells_or.value();
  EXPECT_EQ(cells.cols, 9U);  // implicit row header + 4 leaves + 4 subtotal columns.

  std::size_t subtotal_headers = 0;
  for (const PivotCell& cell : cells.cells) {
    if (cell.kind == PivotCellKind::ColSubtotal && cell.row == cells.top) {
      ++subtotal_headers;
    }
  }
  EXPECT_EQ(subtotal_headers, 4U);
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

TEST(PivotLayout, ProjectsDeepPostOrderSubtotalsAccordingToSubtotalTop) {
  PivotTable table = build_deep_layout_table();
  PivotResult result = build_deep_layout_result(/*multiple_subtotals=*/false);
  PivotLayoutOptions options;
  options.row_labels_label = "Row Labels";

  const std::vector<double> post_order = {1.0,   2.0,   101.0, 3.0, 102.0, 103.0, 4.0,
                                          104.0, 105.0, 106.0, 5.0, 107.0, 108.0, 109.0};
  // The index is (level-one-top << 2) | (level-two-top << 1) |
  // level-three-top. Compact and Outline both apply each field's flag to
  // its own node block, so every combination has a distinct exact order.
  const std::vector<std::vector<double>> expected_by_flags = {
      post_order,
      {101.0, 1.0, 2.0, 102.0, 3.0, 103.0, 104.0, 4.0, 105.0, 106.0, 107.0, 5.0, 108.0, 109.0},
      {103.0, 1.0, 2.0, 101.0, 3.0, 102.0, 105.0, 4.0, 104.0, 106.0, 108.0, 5.0, 107.0, 109.0},
      {103.0, 101.0, 1.0, 2.0, 102.0, 3.0, 105.0, 104.0, 4.0, 106.0, 108.0, 107.0, 5.0, 109.0},
      {106.0, 1.0, 2.0, 101.0, 3.0, 102.0, 103.0, 4.0, 104.0, 105.0, 109.0, 5.0, 107.0, 108.0},
      {106.0, 101.0, 1.0, 2.0, 102.0, 3.0, 103.0, 104.0, 4.0, 105.0, 109.0, 107.0, 5.0, 108.0},
      {106.0, 103.0, 1.0, 2.0, 101.0, 3.0, 102.0, 105.0, 4.0, 104.0, 109.0, 108.0, 5.0, 107.0},
      {106.0, 103.0, 101.0, 1.0, 2.0, 102.0, 3.0, 105.0, 104.0, 4.0, 109.0, 108.0, 107.0, 5.0},
  };

  auto assert_projection = [&](PivotLayout layout_mode, bool level_one_top, bool level_two_top, bool level_three_top,
                               const std::vector<double>& expected) {
    table.set_layout(layout_mode);
    table.mutable_fields()[0].subtotal_top = level_one_top;
    table.mutable_fields()[1].subtotal_top = level_two_top;
    table.mutable_fields()[2].subtotal_top = level_three_top;
    table.mutable_fields()[3].subtotal_top = false;

    auto cells_or = layout(table, result, options);
    ASSERT_TRUE(static_cast<bool>(cells_or)) << cells_or.error().message << " " << cells_or.error().context;
    const PivotCells& cells = cells_or.value();
    const std::uint32_t data_col = cells.left + (layout_mode == PivotLayout::Compact ? 1U : 4U);
    EXPECT_EQ(projected_data_values(cells, data_col), expected);
    const std::vector<PivotCellKind> kinds = projected_data_kinds(cells, data_col);
    ASSERT_EQ(kinds.size(), expected.size());
    std::size_t projected_subtotals = 0;
    for (PivotCellKind kind : kinds) {
      if (kind == PivotCellKind::RowSubtotal) {
        ++projected_subtotals;
      }
    }
    EXPECT_EQ(projected_subtotals, result.row_subtotals.size());
    EXPECT_EQ(cells.rows, expected.size() + 1U);
  };

  // Compact and Outline apply the flag at each of the three non-leaf
  // hierarchy levels. Exercise all eight combinations rather than only the
  // all-false/all-true endpoints and one mixed case.
  for (std::size_t mask = 0; mask < expected_by_flags.size(); ++mask) {
    const bool level_one_top = (mask & 4U) != 0;
    const bool level_two_top = (mask & 2U) != 0;
    const bool level_three_top = (mask & 1U) != 0;
    assert_projection(PivotLayout::Compact, level_one_top, level_two_top, level_three_top, expected_by_flags[mask]);
    assert_projection(PivotLayout::Outline, level_one_top, level_two_top, level_three_top, expected_by_flags[mask]);
  }

  // Tabular form deliberately keeps its historical below-group placement for
  // every flag combination; the metadata is still consumed exactly once.
  for (std::size_t mask = 0; mask < expected_by_flags.size(); ++mask) {
    assert_projection(PivotLayout::Tabular, (mask & 4U) != 0, (mask & 2U) != 0, (mask & 1U) != 0, post_order);
  }
}

TEST(PivotLayout, ProjectsMultipleCustomSubtotalsAsOneContiguousNodeBlock) {
  PivotTable table = build_deep_layout_table();
  PivotResult result = build_deep_layout_result(/*multiple_subtotals=*/true);
  PivotLayoutOptions options;
  options.row_labels_label = "Row Labels";
  table.mutable_fields()[0].subtotal_top = true;
  table.mutable_fields()[1].subtotal_top = true;
  table.mutable_fields()[2].subtotal_top = true;
  table.mutable_fields()[2].subtotal_fns = {SubtotalFn::Average, SubtotalFn::Max};

  auto cells_or = layout(table, result, options);
  ASSERT_TRUE(static_cast<bool>(cells_or)) << cells_or.error().message << " " << cells_or.error().context;
  const PivotCells& cells = cells_or.value();
  const std::vector<double> expected = {207.0, 204.0, 201.0, 202.0, 1.0,   2.0,   203.0, 211.0, 3.0,
                                        206.0, 205.0, 212.0, 4.0,   210.0, 209.0, 208.0, 213.0, 5.0};
  EXPECT_EQ(projected_data_values(cells, cells.left + 1U), expected);
  EXPECT_EQ(cells.rows, expected.size() + 1U);
  std::size_t subtotal_rows = 0;
  for (PivotCellKind kind : projected_data_kinds(cells, cells.left + 1U)) {
    if (kind == PivotCellKind::RowSubtotal) {
      ++subtotal_rows;
    }
  }
  EXPECT_EQ(subtotal_rows, result.row_subtotals.size());
}

TEST(PivotLayout, RejectsRowSubtotalMetadataInconsistentWithHierarchy) {
  PivotTable table = build_deep_layout_table();
  PivotLayoutOptions options;
  options.row_labels_label = "Row Labels";

  PivotResult wrong_depth = build_deep_layout_result(/*multiple_subtotals=*/false);
  wrong_depth.row_subtotals[0].depth = 1;
  auto depth_or = layout(table, wrong_depth, options);
  ASSERT_FALSE(static_cast<bool>(depth_or));
  EXPECT_EQ(depth_or.error().code, FormulonErrorCode::kEvalPivotInvalid);

  PivotResult wrong_order = build_deep_layout_result(/*multiple_subtotals=*/false);
  wrong_order.row_subtotals[0].labels[2] = "not-A1x";
  auto order_or = layout(table, wrong_order, options);
  ASSERT_FALSE(static_cast<bool>(order_or));
  EXPECT_EQ(order_or.error().code, FormulonErrorCode::kEvalPivotInvalid);
}

TEST(PivotLayout, RejectsStrayAndMalformedSubtotalMetadata) {
  PivotLayoutOptions options;
  options.row_labels_label = "Row Labels";

  auto expect_invalid = [&](const PivotTable& table, const PivotResult& result) {
    auto cells_or = layout(table, result, options);
    ASSERT_FALSE(static_cast<bool>(cells_or));
    EXPECT_EQ(cells_or.error().code, FormulonErrorCode::kEvalPivotInvalid);
  };

  // A subtotal entry cannot exist when its axis has no hierarchy.
  {
    PivotTable table = build_table(/*row=*/{}, /*col=*/{});
    PivotResult result;
    result.values = {{{Value::number(1.0)}}};
    RowSubtotal stray;
    stray.labels = {"Orphan"};
    stray.depth = 0;
    stray.values = {Value::number(1.0)};
    result.row_subtotals.push_back(std::move(stray));
    expect_invalid(table, result);
  }
  {
    PivotTable table = build_table(/*row=*/{0}, /*col=*/{});
    PivotResult result;
    result.rows = {RowHierarchyNode{"North", {}}};
    result.values = {{{Value::number(1.0)}}};
    ColSubtotal stray;
    stray.labels = {"Orphan"};
    stray.depth = 0;
    stray.values = {{Value::number(1.0)}};
    result.col_subtotals.push_back(std::move(stray));
    expect_invalid(table, result);
  }

  PivotCache cache = build_region_year_quarter_cache();
  PivotTable table = build_region_year_quarter_table();
  auto result_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  ASSERT_EQ(result_or.value().col_subtotals.size(), 2U);
  const PivotResult valid_result = result_or.value();

  // Missing one expected owner block, a bad depth, and a bad path must all
  // fail before any partially projected column layout can escape.
  {
    PivotResult insufficient = valid_result;
    insufficient.col_subtotals.pop_back();
    expect_invalid(table, insufficient);
  }
  {
    PivotResult wrong_depth = valid_result;
    wrong_depth.col_subtotals[0].depth = 1;
    expect_invalid(table, wrong_depth);
  }
  {
    PivotResult wrong_path = valid_result;
    wrong_path.col_subtotals[0].labels[0] = "Missing";
    expect_invalid(table, wrong_path);
  }
  {
    PivotResult unconsumed = valid_result;
    ColSubtotal orphan;
    orphan.labels = {"Orphan"};
    orphan.depth = 0;
    orphan.values = {{Value::number(0.0)}, {Value::number(0.0)}};
    unconsumed.col_subtotals.push_back(std::move(orphan));
    expect_invalid(table, unconsumed);
  }

  // A field-order index outside the table definition cannot silently turn a
  // valid subtotal stream into a partially consumed projection.
  {
    PivotTable bad_field = build_deep_layout_table();
    bad_field.mutable_row_field_order()[0] = 999U;
    PivotResult result = build_deep_layout_result(/*multiple_subtotals=*/false);
    expect_invalid(bad_field, result);
  }
}

TEST(PivotLayout, SortIndependentSubtotalProjectionKeepsBothAxesAndGrandTotals) {
  auto assert_sorted_layout = [](bool sort_by_value) {
    PivotCache cache = build_region_year_quarter_channel_cache();
    PivotTable table = build_region_channel_by_year_quarter_table();
    table.set_grand_totals(/*rows=*/true, /*cols=*/true);
    for (const std::uint32_t field_index : {0U, 1U, 2U, 3U}) {
      table.mutable_fields()[field_index].sort.ascending = false;
      if (sort_by_value) {
        table.mutable_fields()[field_index].sort.by_field = "Amount";
      }
    }

    auto result_or = evaluate(table, cache);
    ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
    const PivotResult& result = result_or.value();
    ASSERT_EQ(result.row_subtotals.size(), 2U);
    ASSERT_EQ(result.col_subtotals.size(), 2U);

    PivotLayoutOptions options;
    options.row_labels_label = "Row Labels";
    auto cells_or = layout(table, result, options);
    ASSERT_TRUE(static_cast<bool>(cells_or)) << cells_or.error().message << " " << cells_or.error().context;
    const PivotCells& cells = cells_or.value();

    // Three header rows (two column levels plus the compact row-label row),
    // four leaves + two row subtotals, and a bottom grand-total row.
    EXPECT_EQ(cells.rows, 10U);
    // One row-label column, three column leaves + two column subtotals, and
    // one right-hand grand-total column.
    EXPECT_EQ(cells.cols, 7U);

    std::size_t row_subtotal_labels = 0;
    for (const PivotCell& cell : cells.cells) {
      if (cell.col == cells.left && cell.kind == PivotCellKind::RowSubtotal) {
        ++row_subtotal_labels;
      }
    }
    EXPECT_EQ(row_subtotal_labels, result.row_subtotals.size());

    const std::uint32_t first_data_row = cells.top + 3U;
    std::vector<std::uint32_t> header_col_subtotal_columns;
    for (const PivotCell& cell : cells.cells) {
      if (cell.row < first_data_row && cell.kind == PivotCellKind::ColSubtotal) {
        bool already_seen = false;
        for (const std::uint32_t column : header_col_subtotal_columns) {
          if (column == cell.col) {
            already_seen = true;
            break;
          }
        }
        if (!already_seen) {
          header_col_subtotal_columns.push_back(cell.col);
        }
      }
    }
    EXPECT_EQ(header_col_subtotal_columns.size(), result.col_subtotals.size());

    bool saw_bottom_grand_total = false;
    for (const PivotCell& cell : cells.cells) {
      if (cell.kind == PivotCellKind::GrandTotal && cell.row == cells.top + cells.rows - 1U) {
        saw_bottom_grand_total = true;
        break;
      }
    }
    EXPECT_TRUE(saw_bottom_grand_total);
  };

  // Exercise ordinary descending hierarchy sorting and value-field sorting;
  // both differ from the evaluator's raw map traversal in at least one axis.
  assert_sorted_layout(/*sort_by_value=*/false);
  assert_sorted_layout(/*sort_by_value=*/true);
}

// --- Page (report filter) axis ---------------------------------------------
//
// The rendered shape is pinned by the `getpivotdata_page_data.page_axis_field`
// workbook-oracle capture: Excel draws one row per page field carrying the
// field name and its selection, then a blank separator row, and counts both
// inside the pivot's own extent. These cover the projection at a non-zero
// anchor and the selection states the capture does not reach.

PivotTable build_page_field_table() {
  // Region moves to the page axis; Product carries the row hierarchy.
  PivotTable table = build_table(/*row=*/{1}, /*col=*/{});
  table.mutable_fields()[0].axis = PivotAxis::Page;
  return table;
}

TEST(PivotLayout, PageFieldHeaderPrecedesTheReport) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_page_field_table();

  auto result_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  auto cells_or = layout(table, result_or.value());
  ASSERT_TRUE(static_cast<bool>(cells_or)) << cells_or.error().message;
  const PivotCells& cells = cells_or.value();

  // Anchor is D3 and the page block takes the first two rows, so everything
  // the projection would otherwise put at the top moves down by two. The
  // default (English) vocabulary uses the legacy two-row header.
  EXPECT_EQ(cells.top, 2U);
  EXPECT_EQ(cells.rows, 7U);  // page field + separator + 2 header rows + 2 data rows + grand total.

  const PivotCell* field_name = find_cell(cells, 2, 3);
  ASSERT_NE(field_name, nullptr);
  EXPECT_EQ(field_name->kind, PivotCellKind::Header);
  ASSERT_TRUE(field_name->value.is_text());
  EXPECT_EQ(field_name->value.as_text(), "Region");

  const PivotCell* selection = find_cell(cells, 2, 4);
  ASSERT_NE(selection, nullptr);
  EXPECT_EQ(selection->kind, PivotCellKind::Header);
  ASSERT_TRUE(selection->value.is_text());
  EXPECT_EQ(selection->value.as_text(), "(All)");
  // Both cells name the field they belong to, so a renderer can bind the
  // selection back to its dropdown.
  EXPECT_EQ(selection->field_name, "Region");

  const PivotCell* separator = find_cell(cells, 3, 3);
  ASSERT_NE(separator, nullptr);
  EXPECT_EQ(separator->kind, PivotCellKind::Blank);
  EXPECT_TRUE(separator->value.is_blank());

  const PivotCell* row_header = find_cell(cells, 5, 3);
  ASSERT_NE(row_header, nullptr);
  EXPECT_EQ(row_header->kind, PivotCellKind::Header);
  ASSERT_TRUE(row_header->value.is_text());
  EXPECT_EQ(row_header->value.as_text(), "Product");
}

TEST(PivotLayout, PageFieldsStackInReportOrder) {
  PivotCache cache = build_region_year_quarter_cache();
  PivotTable table = build_region_year_quarter_table();
  // Year and Quarter both leave the column axis for the page axis, which
  // leaves Region alone on the row axis.
  table.mutable_fields()[1].axis = PivotAxis::Page;
  table.mutable_fields()[2].axis = PivotAxis::Page;
  table.mutable_col_field_order().clear();

  // `<pageFields>` states the order explicitly, and it need not follow
  // `<pivotFields>` document order.
  table.mutable_page_fields().push_back(PivotPageField{2, std::nullopt});
  table.mutable_page_fields().push_back(PivotPageField{1, std::nullopt});

  auto result_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  auto cells_or = layout(table, result_or.value());
  ASSERT_TRUE(static_cast<bool>(cells_or)) << cells_or.error().message;
  const PivotCells& cells = cells_or.value();

  const PivotCell* first = find_cell(cells, 2, 3);
  ASSERT_NE(first, nullptr);
  ASSERT_TRUE(first->value.is_text());
  EXPECT_EQ(first->value.as_text(), "Quarter");

  const PivotCell* second = find_cell(cells, 3, 3);
  ASSERT_NE(second, nullptr);
  ASSERT_TRUE(second->value.is_text());
  EXPECT_EQ(second->value.as_text(), "Year");

  const PivotCell* separator = find_cell(cells, 4, 3);
  ASSERT_NE(separator, nullptr);
  EXPECT_EQ(separator->kind, PivotCellKind::Blank);
}

TEST(PivotLayout, PageBlockSpansTheFullReportWidth) {
  PivotCache cache = build_basic_cache();
  // Product on the column axis widens the report past the two columns the
  // page block itself uses; the remainder of each page row is explicit
  // blanks so the projected rectangle matches the reported extent.
  PivotTable table = build_table(/*row=*/{1}, /*col=*/{});
  table.mutable_fields()[0].axis = PivotAxis::Page;
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

  ASSERT_GT(cells.cols, 2U);
  for (std::uint32_t col = 0; col < cells.cols; ++col) {
    const PivotCell* page_cell = find_cell(cells, 2, 3 + col);
    ASSERT_NE(page_cell, nullptr) << "page row missing column " << col;
    const PivotCell* separator = find_cell(cells, 3, 3 + col);
    ASSERT_NE(separator, nullptr) << "separator row missing column " << col;
    EXPECT_EQ(separator->kind, PivotCellKind::Blank);
  }
}

TEST(PivotLayout, PageFieldWithoutAPageAxisAddsNoRows) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_table(/*row=*/{0}, /*col=*/{});

  auto result_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  EXPECT_TRUE(result_or.value().page_selections.empty());

  auto cells_or = layout(table, result_or.value());
  ASSERT_TRUE(static_cast<bool>(cells_or)) << cells_or.error().message;
  // The row-field header keeps the last header row of the legacy layout,
  // undisplaced because nothing sits above it.
  const PivotCell* header = find_cell(cells_or.value(), 3, 3);
  ASSERT_NE(header, nullptr);
  ASSERT_TRUE(header->value.is_text());
  EXPECT_EQ(header->value.as_text(), "Region");
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
