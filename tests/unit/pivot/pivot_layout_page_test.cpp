// Page-field rows and explicit blank cells in pivot layout projection.

#include <cstdint>
#include <optional>
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
using layout_test::build_region_year_quarter_cache;
using layout_test::build_region_year_quarter_table;
using layout_test::build_table;
using layout_test::find_cell;

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

}  // namespace
}  // namespace formulon::pivot
