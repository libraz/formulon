// Pivot layout axes, headers, and column-field projection shapes.

#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_evaluator.h"
#include "pivot/pivot_layout.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot_layout_test_helpers.h"
#include "utils/error.h"
#include "value.h"

namespace formulon::pivot {
namespace {

using layout_test::build_basic_cache;
using layout_test::build_table;
using layout_test::find_cell;
using layout_test::ja_jp_layout_options;
using layout_test::owned_text;

PivotCache build_region_product_quarter_cache() {
  PivotCache cache;
  cache.set_cache_id(4);
  cache.mutable_fields().push_back(PivotCacheField{"Region", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Product", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Quarter", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});
  auto add = [&](const char* region, const char* product, const char* quarter, double amount) {
    PivotCacheRecord rec;
    rec.cells.push_back(owned_text(cache, region));
    rec.cells.push_back(owned_text(cache, product));
    rec.cells.push_back(owned_text(cache, quarter));
    rec.cells.push_back(Value::number(amount));
    cache.mutable_records().push_back(std::move(rec));
  };
  add("North", "Widget", "Q1", 100.0);
  add("North", "Widget", "Q2", 110.0);
  add("North", "Gadget", "Q1", 50.0);
  add("South", "Widget", "Q1", 200.0);
  add("South", "Widget", "Q2", 210.0);
  add("South", "Gadget", "Q2", 300.0);
  return cache;
}

PivotTable build_region_product_by_quarter_table(PivotLayout layout_mode) {
  PivotTable table;
  table.set_pivot_cache_id(4);
  table.set_anchor(0, 0, 1, 1);
  table.set_layout(layout_mode);
  PivotField region;
  region.source_name = "Region";
  region.axis = PivotAxis::Row;
  PivotField product;
  product.source_name = "Product";
  product.axis = PivotAxis::Row;
  PivotField quarter;
  quarter.source_name = "Quarter";
  quarter.axis = PivotAxis::Col;
  PivotField amount;
  amount.source_name = "Amount";
  amount.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(region));
  table.mutable_fields().push_back(std::move(product));
  table.mutable_fields().push_back(std::move(quarter));
  table.mutable_fields().push_back(std::move(amount));
  table.mutable_row_field_order() = {0, 1};
  table.mutable_col_field_order() = {2};
  PivotDataField sum_amount;
  sum_amount.name = "合計 / Amount";  // pivot_layout is locale-agnostic; the caller supplies the display name.
  sum_amount.field_index = 3;
  sum_amount.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum_amount));
  return table;
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

  EXPECT_EQ(cells.rows, 5U);  // data-field headers + row header + 2 rows + col grand total.
  EXPECT_EQ(cells.cols, 5U);  // row label + 2 data fields + 2 grand-total fields.

  const PivotCell* sum_header = find_cell(cells, 2, 4);
  ASSERT_NE(sum_header, nullptr);
  EXPECT_EQ(sum_header->kind, PivotCellKind::Header);
  ASSERT_TRUE(sum_header->value.is_text());
  EXPECT_EQ(sum_header->value.as_text(), "Sum of Amount");

  const PivotCell* count_header = find_cell(cells, 2, 5);
  ASSERT_NE(count_header, nullptr);
  EXPECT_EQ(count_header->kind, PivotCellKind::Header);
  ASSERT_TRUE(count_header->value.is_text());
  EXPECT_EQ(count_header->value.as_text(), "Count of Amount");

  const PivotCell* north_count = find_cell(cells, 4, 5);
  ASSERT_NE(north_count, nullptr);
  EXPECT_EQ(north_count->kind, PivotCellKind::Data);
  ASSERT_TRUE(north_count->value.is_number());
  EXPECT_DOUBLE_EQ(north_count->value.as_number(), 3.0);

  const PivotCell* sum_grand_total = find_cell(cells, 6, 6);
  ASSERT_NE(sum_grand_total, nullptr);
  EXPECT_EQ(sum_grand_total->kind, PivotCellKind::GrandTotal);
  ASSERT_TRUE(sum_grand_total->value.is_number());
  EXPECT_DOUBLE_EQ(sum_grand_total->value.as_number(), 675.0);

  const PivotCell* count_grand_total = find_cell(cells, 6, 7);
  ASSERT_NE(count_grand_total, nullptr);
  EXPECT_EQ(count_grand_total->kind, PivotCellKind::GrandTotal);
  ASSERT_TRUE(count_grand_total->value.is_number());
  EXPECT_DOUBLE_EQ(count_grand_total->value.as_number(), 5.0);
}

TEST(PivotLayout, TabularLayoutWidensToColumnFieldPivot) {
  // Column fields with Tabular/Outline layout were previously
  // unsupported (`col_depth == 0` was required); this reproduces the
  // golden's grid to prove the widened layout matches.
  PivotCache cache = build_region_product_quarter_cache();
  PivotTable table = build_region_product_by_quarter_table(PivotLayout::Tabular);

  auto result_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;

  PivotLayoutOptions options = ja_jp_layout_options();
  auto cells_or = layout(table, result_or.value(), options);
  ASSERT_TRUE(static_cast<bool>(cells_or)) << cells_or.error().message << " " << cells_or.error().context;
  const PivotCells& cells = cells_or.value();

  EXPECT_EQ(cells.rows, 9U);
  EXPECT_EQ(cells.cols, 5U);

  auto text_at = [&](std::uint32_t r, std::uint32_t c) -> std::string {
    const PivotCell* cell = find_cell(cells, r, c);
    if (cell == nullptr || !cell->value.is_text()) {
      return "<missing>";
    }
    return std::string(cell->value.as_text());
  };
  auto number_at = [&](std::uint32_t r, std::uint32_t c) -> double {
    const PivotCell* cell = find_cell(cells, r, c);
    return cell != nullptr && cell->value.is_number() ? cell->value.as_number() : -1.0;
  };
  auto is_blank_at = [&](std::uint32_t r, std::uint32_t c) -> bool {
    const PivotCell* cell = find_cell(cells, r, c);
    return cell != nullptr && cell->value.is_blank();
  };

  // Row 0: data-field name in the corner, "Quarter" over the leaf
  // columns, everything else blank.
  EXPECT_EQ(text_at(0, 0), "合計 / Amount");
  EXPECT_EQ(find_cell(cells, 0, 1)->kind, PivotCellKind::Blank);
  EXPECT_EQ(text_at(0, 2), "Quarter");
  EXPECT_EQ(find_cell(cells, 0, 3)->kind, PivotCellKind::Blank);
  EXPECT_EQ(find_cell(cells, 0, 4)->kind, PivotCellKind::Blank);

  // Row 1: row-field names + column leaf labels (the last col-header
  // row doubles as the row-field-name row).
  EXPECT_EQ(text_at(1, 0), "Region");
  EXPECT_EQ(text_at(1, 1), "Product");
  EXPECT_EQ(text_at(1, 2), "Q1");
  EXPECT_EQ(text_at(1, 3), "Q2");
  EXPECT_EQ(text_at(1, 4), "総計");

  // Leaf and subtotal rows.
  EXPECT_EQ(text_at(2, 0), "North");
  EXPECT_EQ(text_at(2, 1), "Gadget");
  EXPECT_EQ(number_at(2, 2), 50.0);
  EXPECT_TRUE(is_blank_at(2, 3));
  EXPECT_EQ(number_at(2, 4), 50.0);

  EXPECT_TRUE(is_blank_at(3, 0));
  EXPECT_EQ(text_at(3, 1), "Widget");
  EXPECT_EQ(number_at(3, 2), 100.0);
  EXPECT_EQ(number_at(3, 3), 110.0);
  EXPECT_EQ(number_at(3, 4), 210.0);

  EXPECT_EQ(text_at(4, 0), "North 集計");
  EXPECT_TRUE(is_blank_at(4, 1));
  EXPECT_EQ(number_at(4, 2), 150.0);
  EXPECT_EQ(number_at(4, 3), 110.0);
  EXPECT_EQ(number_at(4, 4), 260.0);

  EXPECT_EQ(text_at(5, 0), "South");
  EXPECT_EQ(text_at(5, 1), "Gadget");
  EXPECT_TRUE(is_blank_at(5, 2));
  EXPECT_EQ(number_at(5, 3), 300.0);
  EXPECT_EQ(number_at(5, 4), 300.0);

  EXPECT_TRUE(is_blank_at(6, 0));
  EXPECT_EQ(text_at(6, 1), "Widget");
  EXPECT_EQ(number_at(6, 2), 200.0);
  EXPECT_EQ(number_at(6, 3), 210.0);
  EXPECT_EQ(number_at(6, 4), 410.0);

  EXPECT_EQ(text_at(7, 0), "South 集計");
  EXPECT_TRUE(is_blank_at(7, 1));
  EXPECT_EQ(number_at(7, 2), 200.0);
  EXPECT_EQ(number_at(7, 3), 510.0);
  EXPECT_EQ(number_at(7, 4), 710.0);

  EXPECT_EQ(text_at(8, 0), "総計");
  EXPECT_TRUE(is_blank_at(8, 1));
  EXPECT_EQ(number_at(8, 2), 350.0);
  EXPECT_EQ(number_at(8, 3), 620.0);
  EXPECT_EQ(number_at(8, 4), 970.0);
}

TEST(PivotLayout, OutlineLayoutWidensToColumnFieldPivot) {
  // Same shape as `TabularLayoutWidensToColumnFieldPivot`, but Outline:
  // subtotal rows carry the bare parent label (no " 集計" suffix, since
  // the row doubles as the parent group's own header) and leaf rows
  // never repeat a parent value in its own column.
  // `two_row_fields_by_column_field_outline` in the same golden file.
  PivotCache cache = build_region_product_quarter_cache();
  PivotTable table = build_region_product_by_quarter_table(PivotLayout::Outline);

  auto result_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;

  PivotLayoutOptions options = ja_jp_layout_options();
  auto cells_or = layout(table, result_or.value(), options);
  ASSERT_TRUE(static_cast<bool>(cells_or)) << cells_or.error().message << " " << cells_or.error().context;
  const PivotCells& cells = cells_or.value();

  EXPECT_EQ(cells.rows, 9U);
  EXPECT_EQ(cells.cols, 5U);

  auto text_at = [&](std::uint32_t r, std::uint32_t c) -> std::string {
    const PivotCell* cell = find_cell(cells, r, c);
    if (cell == nullptr || !cell->value.is_text()) {
      return "<missing>";
    }
    return std::string(cell->value.as_text());
  };
  auto number_at = [&](std::uint32_t r, std::uint32_t c) -> double {
    const PivotCell* cell = find_cell(cells, r, c);
    return cell != nullptr && cell->value.is_number() ? cell->value.as_number() : -1.0;
  };
  auto is_blank_at = [&](std::uint32_t r, std::uint32_t c) -> bool {
    const PivotCell* cell = find_cell(cells, r, c);
    return cell != nullptr && cell->value.is_blank();
  };

  EXPECT_EQ(text_at(0, 0), "合計 / Amount");
  EXPECT_EQ(text_at(0, 2), "Quarter");
  EXPECT_EQ(text_at(1, 0), "Region");
  EXPECT_EQ(text_at(1, 1), "Product");
  EXPECT_EQ(text_at(1, 2), "Q1");
  EXPECT_EQ(text_at(1, 3), "Q2");
  EXPECT_EQ(text_at(1, 4), "総計");

  // North's subtotal row (bare label, no " 集計" suffix) precedes its
  // leaf rows -- Outline's subtotal_top default.
  EXPECT_EQ(text_at(2, 0), "North");
  EXPECT_TRUE(is_blank_at(2, 1));
  EXPECT_EQ(number_at(2, 2), 150.0);
  EXPECT_EQ(number_at(2, 3), 110.0);
  EXPECT_EQ(number_at(2, 4), 260.0);

  EXPECT_TRUE(is_blank_at(3, 0));
  EXPECT_EQ(text_at(3, 1), "Gadget");
  EXPECT_EQ(number_at(3, 2), 50.0);
  EXPECT_TRUE(is_blank_at(3, 3));
  EXPECT_EQ(number_at(3, 4), 50.0);

  EXPECT_TRUE(is_blank_at(4, 0));
  EXPECT_EQ(text_at(4, 1), "Widget");
  EXPECT_EQ(number_at(4, 2), 100.0);
  EXPECT_EQ(number_at(4, 3), 110.0);
  EXPECT_EQ(number_at(4, 4), 210.0);

  EXPECT_EQ(text_at(5, 0), "South");
  EXPECT_TRUE(is_blank_at(5, 1));
  EXPECT_EQ(number_at(5, 2), 200.0);
  EXPECT_EQ(number_at(5, 3), 510.0);
  EXPECT_EQ(number_at(5, 4), 710.0);

  EXPECT_TRUE(is_blank_at(6, 0));
  EXPECT_EQ(text_at(6, 1), "Gadget");
  EXPECT_TRUE(is_blank_at(6, 2));
  EXPECT_EQ(number_at(6, 3), 300.0);
  EXPECT_EQ(number_at(6, 4), 300.0);

  EXPECT_TRUE(is_blank_at(7, 0));
  EXPECT_EQ(text_at(7, 1), "Widget");
  EXPECT_EQ(number_at(7, 2), 200.0);
  EXPECT_EQ(number_at(7, 3), 210.0);
  EXPECT_EQ(number_at(7, 4), 410.0);

  EXPECT_EQ(text_at(8, 0), "総計");
  EXPECT_TRUE(is_blank_at(8, 1));
  EXPECT_EQ(number_at(8, 2), 350.0);
  EXPECT_EQ(number_at(8, 3), 620.0);
  EXPECT_EQ(number_at(8, 4), 970.0);
}

TEST(PivotLayout, TabularLayoutTwoColumnFieldsNamesOneColumnPerDepth) {
  // Two fields on each axis (`two_row_fields_two_col_fields` in the same
  // golden file). The extra header row above the column hierarchy
  // names one column field per depth, at `data_left + depth` -- NOT
  // aligned with how wide each level's own leaf group is (Quarter's
  // name sits over Q1's group; Channel's sits one column to the
  // right, not under Store/Web).
  PivotCache cache;
  cache.set_cache_id(5);
  cache.mutable_fields().push_back(PivotCacheField{"Region", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Product", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Quarter", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Channel", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});
  auto add = [&](const char* region, const char* product, const char* quarter, const char* channel, double amount) {
    PivotCacheRecord rec;
    rec.cells.push_back(owned_text(cache, region));
    rec.cells.push_back(owned_text(cache, product));
    rec.cells.push_back(owned_text(cache, quarter));
    rec.cells.push_back(owned_text(cache, channel));
    rec.cells.push_back(Value::number(amount));
    cache.mutable_records().push_back(std::move(rec));
  };
  add("North", "Widget", "Q1", "Web", 100.0);
  add("North", "Widget", "Q2", "Store", 110.0);
  add("North", "Gadget", "Q1", "Web", 50.0);
  add("South", "Widget", "Q1", "Store", 200.0);
  add("South", "Widget", "Q2", "Web", 210.0);
  add("South", "Gadget", "Q2", "Store", 300.0);

  PivotTable table;
  table.set_pivot_cache_id(5);
  table.set_anchor(0, 0, 1, 1);
  table.set_layout(PivotLayout::Tabular);
  PivotField region;
  region.source_name = "Region";
  region.axis = PivotAxis::Row;
  PivotField product;
  product.source_name = "Product";
  product.axis = PivotAxis::Row;
  PivotField quarter;
  quarter.source_name = "Quarter";
  quarter.axis = PivotAxis::Col;
  PivotField channel;
  channel.source_name = "Channel";
  channel.axis = PivotAxis::Col;
  PivotField amount;
  amount.source_name = "Amount";
  amount.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(region));
  table.mutable_fields().push_back(std::move(product));
  table.mutable_fields().push_back(std::move(quarter));
  table.mutable_fields().push_back(std::move(channel));
  table.mutable_fields().push_back(std::move(amount));
  table.mutable_row_field_order() = {0, 1};
  table.mutable_col_field_order() = {2, 3};
  PivotDataField sum_amount;
  sum_amount.name = "合計 / Amount";
  sum_amount.field_index = 4;
  sum_amount.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum_amount));

  auto result_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;

  PivotLayoutOptions options = ja_jp_layout_options();
  auto cells_or = layout(table, result_or.value(), options);
  ASSERT_TRUE(static_cast<bool>(cells_or)) << cells_or.error().message << " " << cells_or.error().context;
  const PivotCells& cells = cells_or.value();

  EXPECT_EQ(cells.rows, 10U);
  EXPECT_EQ(cells.cols, 9U);

  auto text_at = [&](std::uint32_t r, std::uint32_t c) -> std::string {
    const PivotCell* cell = find_cell(cells, r, c);
    if (cell == nullptr || !cell->value.is_text()) {
      return "<missing>";
    }
    return std::string(cell->value.as_text());
  };
  auto is_blank_at = [&](std::uint32_t r, std::uint32_t c) -> bool {
    const PivotCell* cell = find_cell(cells, r, c);
    return cell != nullptr && cell->value.is_blank();
  };

  // Row 0: data-field name, then "Quarter" / "Channel" at data_left/+1,
  // everything else (including the rest of the row-header prefix)
  // blank.
  EXPECT_EQ(text_at(0, 0), "合計 / Amount");
  EXPECT_TRUE(is_blank_at(0, 1));
  EXPECT_EQ(text_at(0, 2), "Quarter");
  EXPECT_EQ(text_at(0, 3), "Channel");
  for (std::uint32_t c = 4; c <= 8; ++c) {
    EXPECT_TRUE(is_blank_at(0, c)) << "col " << c;
  }

  // Row 1: Quarter's own leaf/subtotal labels (row-header columns
  // blank -- this isn't the last col-header row).
  EXPECT_TRUE(is_blank_at(1, 0));
  EXPECT_TRUE(is_blank_at(1, 1));
  EXPECT_EQ(text_at(1, 2), "Q1");
  EXPECT_EQ(text_at(1, 4), "Q1 集計");
  EXPECT_EQ(text_at(1, 5), "Q2");
  EXPECT_EQ(text_at(1, 7), "Q2 集計");
  EXPECT_EQ(text_at(1, 8), "総計");

  // Row 2: Channel's own leaf labels, and the row-field names (this IS
  // the last col-header row).
  EXPECT_EQ(text_at(2, 0), "Region");
  EXPECT_EQ(text_at(2, 1), "Product");
  EXPECT_EQ(text_at(2, 2), "Store");
  EXPECT_EQ(text_at(2, 3), "Web");
  // The grand-total strip's own header sits beside "Q1 集計"/"Q2 集計"
  // (row 1, the outermost column field's row), not beside the row-field
  // names -- row 2's total-column cell stays blank.
  EXPECT_TRUE(is_blank_at(2, 8));
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
