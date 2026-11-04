// Report options in the layout projection: the grand-total caption, blank
// rows after item blocks, repeated (filled-down) labels, collapsed items,
// number-formatted item labels, the display options (headers, drop zones,
// error and missing captions, indent, merged labels), cell roles, and the
// item path behind a value cell.

#include <cstddef>
#include <cstdint>
#include <optional>
#include <sstream>
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
using layout_test::ja_jp_layout_options;
using layout_test::owned_text;

constexpr std::uint32_t kCacheId = 7;

PivotCache build_region_product_cache() {
  PivotCache cache;
  cache.set_cache_id(kCacheId);
  for (const char* name : {"Region", "Product", "Amt", "Qty"}) {
    cache.mutable_fields().push_back(PivotCacheField{name, {}});
  }
  auto add = [&](const char* region, const char* product, double amt, double qty) {
    PivotCacheRecord rec;
    rec.cells.push_back(owned_text(cache, region));
    rec.cells.push_back(owned_text(cache, product));
    rec.cells.push_back(Value::number(amt));
    rec.cells.push_back(Value::number(qty));
    cache.mutable_records().push_back(std::move(rec));
  };
  add("North", "A", 100.0, 1.0);
  add("North", "B", 200.0, 2.0);
  add("South", "A", 300.0, 3.0);
  add("South", "B", 400.0, 4.0);
  add("East", "A", 500.0, 5.0);
  return cache;
}

/// Region (and Product, unless `single_row_field`) on the rows, Sum of Amt and
/// Sum of Qty on the columns, anchored at A3.
PivotTable build_region_product_table(PivotLayout layout, bool single_row_field = false) {
  PivotTable table;
  table.set_name("PivotTable1");
  table.set_pivot_cache_id(kCacheId);
  table.set_anchor(2, 0, 1, 1);
  for (const char* name : {"Region", "Product", "Amt", "Qty"}) {
    PivotField field;
    field.source_name = name;
    field.axis = PivotAxis::None;
    table.mutable_fields().push_back(std::move(field));
  }
  table.mutable_fields()[0].axis = PivotAxis::Row;
  table.mutable_fields()[2].axis = PivotAxis::Value;
  table.mutable_fields()[3].axis = PivotAxis::Value;
  for (const char* name : {"East", "North", "South"}) {
    PivotItem item;
    item.name = name;
    table.mutable_fields()[0].items.push_back(std::move(item));
  }
  if (single_row_field) {
    table.mutable_row_field_order() = {0};
  } else {
    table.mutable_fields()[1].axis = PivotAxis::Row;
    table.mutable_row_field_order() = {0, 1};
  }
  for (const auto& [name, field_index] :
       {std::pair<const char*, std::uint32_t>{"Sum of Amt", 2U}, {"Sum of Qty", 3U}}) {
    PivotDataField data;
    data.name = name;
    data.field_index = field_index;
    data.aggregation = Aggregation::Sum;
    table.mutable_data_fields().push_back(std::move(data));
  }
  table.set_layout(layout);
  return table;
}

std::string CellText(const Value& v) {
  if (v.is_text()) {
    return std::string(v.as_text());
  }
  if (v.is_number()) {
    std::ostringstream os;
    os << v.as_number();
    return os.str();
  }
  return v.is_blank() ? std::string() : std::string("?");
}

/// The projected grid, one `|`-joined line per row.
std::vector<std::string> Grid(const PivotCells& cells) {
  std::vector<std::vector<std::string>> grid(cells.rows, std::vector<std::string>(cells.cols));
  for (const PivotCell& cell : cells.cells) {
    const std::uint32_t r = cell.row - cells.top;
    const std::uint32_t c = cell.col - cells.left;
    if (r < cells.rows && c < cells.cols) {
      grid[r][c] = CellText(cell.value);
    }
  }
  std::vector<std::string> lines;
  for (const std::vector<std::string>& row : grid) {
    std::string line;
    for (std::size_t c = 0; c < row.size(); ++c) {
      line += (c == 0 ? "" : "|") + row[c];
    }
    lines.push_back(std::move(line));
  }
  return lines;
}

std::vector<std::string> RenderGrid(const PivotTable& table, const PivotCache& cache) {
  const PivotLayoutOptions options = ja_jp_layout_options();
  auto result_or = evaluate(table, cache, options);
  EXPECT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  if (!result_or) {
    return {};
  }
  auto cells_or = layout(table, result_or.value(), options);
  EXPECT_TRUE(static_cast<bool>(cells_or)) << cells_or.error().message << " " << cells_or.error().context;
  return cells_or ? Grid(cells_or.value()) : std::vector<std::string>{};
}

// "集計" / "総計".
#define SUBTOTAL " \xE9\x9B\x86\xE8\xA8\x88"
#define TOTAL "\xE7\xB7\x8F\xE8\xA8\x88"
#define ROW_LABELS "\xE8\xA1\x8C\xE3\x83\xA9\xE3\x83\x99\xE3\x83\xAB"
#define COL_LABELS "\xE5\x88\x97\xE3\x83\xA9\xE3\x83\x99\xE3\x83\xAB"

const std::vector<std::string> kTabular = {
    "Region|Product|Sum of Amt|Sum of Qty",
    "East|A|500|5",
    "East" SUBTOTAL "||500|5",
    "North|A|100|1",
    "|B|200|2",
    "North" SUBTOTAL "||300|3",
    "South|A|300|3",
    "|B|400|4",
    "South" SUBTOTAL "||700|7",
    TOTAL "||1500|15",
};

const std::vector<std::string> kOutline = {
    "Region|Product|Sum of Amt|Sum of Qty",
    "East||500|5",
    "|A|500|5",
    "North||300|3",
    "|A|100|1",
    "|B|200|2",
    "South||700|7",
    "|A|300|3",
    "|B|400|4",
    TOTAL "||1500|15",
};

const std::vector<std::string> kCompactSingle = {
    ROW_LABELS "|Sum of Amt|Sum of Qty", "East|500|5", "North|300|3", "South|700|7", TOTAL "|1500|15",
};

TEST(PivotLayoutOptions, BaselineGridsMatchExcelDefaults) {
  const PivotCache cache = build_region_product_cache();
  EXPECT_EQ(RenderGrid(build_region_product_table(PivotLayout::Tabular), cache), kTabular);
  EXPECT_EQ(RenderGrid(build_region_product_table(PivotLayout::Outline), cache), kOutline);
  EXPECT_EQ(RenderGrid(build_region_product_table(PivotLayout::Compact, true), cache), kCompactSingle);
}

// ---------------------------------------------------------------------------
// Grand-total caption
// ---------------------------------------------------------------------------

TEST(PivotLayoutOptions, GrandTotalCaptionRenamesOnlyTheGrandTotalRow) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Tabular);
  table.set_grand_total_caption(std::string("TOTAL"));
  std::vector<std::string> expected = kTabular;
  expected.back() = "TOTAL||1500|15";
  EXPECT_EQ(RenderGrid(table, cache), expected);
}

TEST(PivotLayoutOptions, EmptyGrandTotalCaptionKeepsTheLocaleWord) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Tabular);
  table.set_grand_total_caption(std::string());
  EXPECT_EQ(RenderGrid(table, cache), kTabular);
}

TEST(PivotLayoutOptions, GrandTotalCaptionKeepsDataTotalLabels) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Compact, true);
  table.set_row_values_position(1);
  table.set_grand_total_caption(std::string("TOTAL"));
  const std::vector<std::string> expected = {
      ROW_LABELS "|", "East|", "Sum of Amt|500", "Sum of Qty|5", "North|", "Sum of Amt|300", "Sum of Qty|3", "South|",
      "Sum of Amt|700", "Sum of Qty|7",
      // "全体の Sum of Amt" / "全体の Sum of Qty".
      "\xE5\x85\xA8\xE4\xBD\x93\xE3\x81\xAE Sum of Amt|1500", "\xE5\x85\xA8\xE4\xBD\x93\xE3\x81\xAE Sum of Qty|15"};
  PivotLayoutOptions options = ja_jp_layout_options();
  options.data_total_prefix = "\xE5\x85\xA8\xE4\xBD\x93\xE3\x81\xAE ";
  auto result_or = evaluate(table, cache, options);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  auto cells_or = layout(table, result_or.value(), options);
  ASSERT_TRUE(static_cast<bool>(cells_or)) << cells_or.error().message;
  EXPECT_EQ(Grid(cells_or.value()), expected);
}

TEST(PivotLayoutOptions, GrandTotalCaptionRenamesTheGrandTotalColumn) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_table({0}, {1});
  table.set_grand_total_caption(std::string("ALL"));
  const PivotLayoutOptions options = ja_jp_layout_options();
  auto result_or = evaluate(table, cache, options);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  auto cells_or = layout(table, result_or.value(), options);
  ASSERT_TRUE(static_cast<bool>(cells_or)) << cells_or.error().message;
  const std::vector<std::string> grid = Grid(cells_or.value());
  ASSERT_GE(grid.size(), 2U);
  EXPECT_EQ(grid[1], ROW_LABELS "|Gadget|Widget|ALL");
  EXPECT_EQ(grid.back().substr(0, 4), "ALL|");
}

// ---------------------------------------------------------------------------
// Blank rows
// ---------------------------------------------------------------------------

TEST(PivotLayoutOptions, BlankRowFollowsEachOuterBlockInTabular) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Tabular);
  table.mutable_fields()[0].insert_blank_row = true;
  const std::vector<std::string> expected = {
      "Region|Product|Sum of Amt|Sum of Qty",
      "East|A|500|5",
      "East" SUBTOTAL "||500|5",
      "|||",
      "North|A|100|1",
      "|B|200|2",
      "North" SUBTOTAL "||300|3",
      "|||",
      "South|A|300|3",
      "|B|400|4",
      "South" SUBTOTAL "||700|7",
      "|||",
      TOTAL "||1500|15",
  };
  EXPECT_EQ(RenderGrid(table, cache), expected);
}

TEST(PivotLayoutOptions, BlankRowFollowsEachOuterBlockInOutline) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Outline);
  table.mutable_fields()[0].insert_blank_row = true;
  const std::vector<std::string> expected = {
      "Region|Product|Sum of Amt|Sum of Qty",
      "East||500|5",
      "|A|500|5",
      "|||",
      "North||300|3",
      "|A|100|1",
      "|B|200|2",
      "|||",
      "South||700|7",
      "|A|300|3",
      "|B|400|4",
      "|||",
      TOTAL "||1500|15",
  };
  EXPECT_EQ(RenderGrid(table, cache), expected);
}

TEST(PivotLayoutOptions, BlankRowIsIgnoredOnTheInnermostField) {
  const PivotCache cache = build_region_product_cache();
  PivotTable single = build_region_product_table(PivotLayout::Compact, true);
  single.mutable_fields()[0].insert_blank_row = true;
  EXPECT_EQ(RenderGrid(single, cache), kCompactSingle);

  PivotTable inner = build_region_product_table(PivotLayout::Tabular);
  inner.mutable_fields()[1].insert_blank_row = true;
  EXPECT_EQ(RenderGrid(inner, cache), kTabular);
}

TEST(PivotLayoutOptions, BlankRowCellsAreBlankKind) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Tabular);
  table.mutable_fields()[0].insert_blank_row = true;
  const PivotLayoutOptions options = ja_jp_layout_options();
  auto result_or = evaluate(table, cache, options);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  auto cells_or = layout(table, result_or.value(), options);
  ASSERT_TRUE(static_cast<bool>(cells_or)) << cells_or.error().message;
  // A3 header, then East A and East 集計: the blank row is row 6 (0-based 5).
  for (std::uint32_t col = 0; col < 4; ++col) {
    const PivotCell* cell = layout_test::find_cell(cells_or.value(), 5, col);
    ASSERT_NE(cell, nullptr) << col;
    EXPECT_EQ(cell->kind, PivotCellKind::Blank) << col;
  }
}

TEST(PivotLayoutOptions, TabularShowsAnInnerLabelUnderANewOuterItem) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Tabular);
  table.mutable_fields()[3].axis = PivotAxis::Row;
  table.mutable_row_field_order() = {0, 1, 3};
  table.mutable_data_fields().pop_back();
  const std::vector<std::string> expected = {
      "Region|Product|Qty|Sum of Amt",
      "East|A|5|500",
      "|A" SUBTOTAL "||500",
      "East" SUBTOTAL "|||500",
      "North|A|1|100",
      "|A" SUBTOTAL "||100",
      "|B|2|200",
      "|B" SUBTOTAL "||200",
      "North" SUBTOTAL "|||300",
      "South|A|3|300",
      "|A" SUBTOTAL "||300",
      "|B|4|400",
      "|B" SUBTOTAL "||400",
      "South" SUBTOTAL "|||700",
      TOTAL "|||1500",
  };
  EXPECT_EQ(RenderGrid(table, cache), expected);
}

// ---------------------------------------------------------------------------
// Repeated labels
// ---------------------------------------------------------------------------

TEST(PivotLayoutOptions, FillDownRepeatsTheOuterLabelInTabular) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Tabular);
  table.mutable_fields()[0].fill_down_labels = true;
  std::vector<std::string> expected = kTabular;
  expected[4] = "North|B|200|2";
  expected[7] = "South|B|400|4";
  EXPECT_EQ(RenderGrid(table, cache), expected);
}

TEST(PivotLayoutOptions, FillDownRepeatsTheOuterLabelInOutline) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Outline);
  table.mutable_fields()[0].fill_down_labels = true;
  const std::vector<std::string> expected = {
      "Region|Product|Sum of Amt|Sum of Qty",
      "East||500|5",
      "East|A|500|5",
      "North||300|3",
      "North|A|100|1",
      "North|B|200|2",
      "South||700|7",
      "South|A|300|3",
      "South|B|400|4",
      TOTAL "||1500|15",
  };
  EXPECT_EQ(RenderGrid(table, cache), expected);
}

TEST(PivotLayoutOptions, FillDownIsIgnoredInCompact) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Compact, true);
  table.mutable_fields()[0].fill_down_labels = true;
  EXPECT_EQ(RenderGrid(table, cache), kCompactSingle);
}

// ---------------------------------------------------------------------------
// Collapsed items
// ---------------------------------------------------------------------------

void CollapseItem(PivotTable& table, std::size_t item) {
  table.mutable_fields()[0].items[item].show_detail = false;
}

TEST(PivotLayoutOptions, CollapsedItemInTabularIsOneTotalRow) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Tabular);
  CollapseItem(table, 0);
  const std::vector<std::string> expected = {
      "Region|Product|Sum of Amt|Sum of Qty",
      "East||500|5",
      "North|A|100|1",
      "|B|200|2",
      "North" SUBTOTAL "||300|3",
      "South|A|300|3",
      "|B|400|4",
      "South" SUBTOTAL "||700|7",
      TOTAL "||1500|15",
  };
  EXPECT_EQ(RenderGrid(table, cache), expected);
}

TEST(PivotLayoutOptions, CollapsedItemInOutlineKeepsOnlyItsHeaderRow) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Outline);
  CollapseItem(table, 0);
  std::vector<std::string> expected = kOutline;
  expected.erase(expected.begin() + 2);
  EXPECT_EQ(RenderGrid(table, cache), expected);
}

TEST(PivotLayoutOptions, CollapsedItemOnTheInnermostFieldIsIgnored) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Compact, true);
  CollapseItem(table, 0);
  EXPECT_EQ(RenderGrid(table, cache), kCompactSingle);
}

TEST(PivotLayoutOptions, CollapsedFieldShowsOnlyItsItems) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Outline);
  for (std::size_t i = 0; i < 3; ++i) {
    CollapseItem(table, i);
  }
  table.mutable_fields()[0].marker_items.push_back(
      PivotMarkerItem{"<item t=\"default\" sd=\"0\"/>", "default", 3, false});
  const std::vector<std::string> expected = {
      "Region|Product|Sum of Amt|Sum of Qty", "East||500|5", "North||300|3", "South||700|7", TOTAL "||1500|15",
  };
  EXPECT_EQ(RenderGrid(table, cache), expected);
}

// ---------------------------------------------------------------------------
// Number-formatted item labels
// ---------------------------------------------------------------------------

TEST(PivotLayoutOptions, LabelCellCarriesTheItemNumberFormat) {
  PivotCache cache;
  cache.set_cache_id(kCacheId);
  cache.mutable_fields().push_back(PivotCacheField{"Date", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amt", {}});
  cache.mutable_fields()[0].num_fmt_id = 14;
  for (const double serial : {45306.0, 45316.0}) {
    PivotCacheRecord rec;
    rec.cells.push_back(Value::number(serial));
    rec.cells.push_back(Value::number(10.0));
    cache.mutable_records().push_back(std::move(rec));
  }
  PivotTable table;
  table.set_pivot_cache_id(kCacheId);
  table.set_anchor(2, 0, 1, 1);
  PivotField date;
  date.source_name = "Date";
  date.axis = PivotAxis::Row;
  PivotField amt;
  amt.source_name = "Amt";
  amt.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(date));
  table.mutable_fields().push_back(std::move(amt));
  table.mutable_row_field_order() = {0};
  PivotDataField sum;
  sum.name = "Sum of Amt";
  sum.field_index = 1;
  sum.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum));

  const PivotLayoutOptions options = ja_jp_layout_options();
  auto result_or = evaluate(table, cache, options);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  const PivotResult& result = result_or.value();
  ASSERT_EQ(result.rows.size(), 2U);
  ASSERT_FALSE(result.rows[0].number_format.empty());
  auto cells_or = layout(table, result, options);
  ASSERT_TRUE(static_cast<bool>(cells_or)) << cells_or.error().message;
  for (std::uint32_t i = 0; i < 2; ++i) {
    const PivotCell* label = layout_test::find_cell(cells_or.value(), 3 + i, 0);
    ASSERT_NE(label, nullptr) << i;
    ASSERT_TRUE(label->value.is_text()) << i;
    EXPECT_EQ(label->value.as_text(), result.rows[i].label) << i;
    EXPECT_EQ(label->number_format, result.rows[i].number_format) << i;
  }
  // The grand total label is not an item.
  const PivotCell* total = layout_test::find_cell(cells_or.value(), 5, 0);
  ASSERT_NE(total, nullptr);
  EXPECT_TRUE(total->number_format.empty());
}

// ---------------------------------------------------------------------------
// Display options
// ---------------------------------------------------------------------------

PivotCells ProjectCells(const PivotTable& table, const PivotCache& cache) {
  const PivotLayoutOptions options = ja_jp_layout_options();
  auto result_or = evaluate(table, cache, options);
  EXPECT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  if (!result_or) {
    return {};
  }
  auto cells_or = layout(table, result_or.value(), options);
  EXPECT_TRUE(static_cast<bool>(cells_or)) << cells_or.error().message;
  return cells_or ? std::move(cells_or.value()) : PivotCells{};
}

/// Region on the rows with the data fields below each item, Product on the columns.
PivotTable build_cross_table_values_on_rows() {
  PivotTable table = build_region_product_table(PivotLayout::Compact, true);
  table.mutable_fields()[1].axis = PivotAxis::Col;
  table.mutable_col_field_order() = {1};
  table.set_row_values_position(1);
  return table;
}

TEST(PivotLayoutOptions, HiddenHeadersDropTheColumnLabelsRow) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_cross_table_values_on_rows();
  const std::vector<std::string> shown = RenderGrid(table, cache);
  ASSERT_GE(shown.size(), 3U);
  EXPECT_EQ(shown[0], "|" COL_LABELS "||");
  EXPECT_EQ(shown[1], ROW_LABELS "|A|B|" TOTAL);
  table.set_show_headers(false);
  std::vector<std::string> expected(shown.begin() + 1, shown.end());
  expected[0] = "|A|B|" TOTAL;
  EXPECT_EQ(RenderGrid(table, cache), expected);
}

TEST(PivotLayoutOptions, HiddenHeadersKeepALoneDataFieldCaption) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_cross_table_values_on_rows();
  table.mutable_data_fields().pop_back();
  table.set_show_headers(false);
  const std::vector<std::string> grid = RenderGrid(table, cache);
  ASSERT_GE(grid.size(), 2U);
  EXPECT_EQ(grid[0], "Sum of Amt|||");
  EXPECT_EQ(grid[1], "|A|B|" TOTAL);
}

TEST(PivotLayoutOptions, HiddenHeadersBlankTabularFieldNames) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Tabular);
  table.set_show_headers(false);
  std::vector<std::string> expected = kTabular;
  expected[0] = "||Sum of Amt|Sum of Qty";
  EXPECT_EQ(RenderGrid(table, cache), expected);
}

TEST(PivotLayoutOptions, GridDropZonesAddARowWithoutColumnFields) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Compact, true);
  table.set_grid_drop_zones(true);
  std::vector<std::string> expected = kCompactSingle;
  expected.insert(expected.begin(), "|Values|");
  EXPECT_EQ(RenderGrid(table, cache), expected);
}

TEST(PivotLayoutOptions, GridDropZonesLeaveAColumnFieldReportAlone) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_cross_table_values_on_rows();
  const std::vector<std::string> base = RenderGrid(table, cache);
  table.set_grid_drop_zones(true);
  EXPECT_EQ(RenderGrid(table, cache), base);
}

TEST(PivotLayoutOptions, MissingCaptionDashIsWrittenAsZero) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_cross_table_values_on_rows();
  table.set_missing_caption(std::string("-"));
  const PivotCells cells = ProjectCells(table, cache);
  // Rows: 2 caption, 3 column items, 4 East, 5 East Sum of Amt; East has no B.
  const PivotCell* hole = layout_test::find_cell(cells, 5, 2);
  ASSERT_NE(hole, nullptr);
  ASSERT_TRUE(hole->value.is_number());
  EXPECT_EQ(hole->value.as_number(), 0.0);
  EXPECT_EQ(hole->role, PivotCellRole::DataValue);
  // The heading row of an item stays empty.
  for (std::uint32_t col = 1; col < 4; ++col) {
    const PivotCell* heading = layout_test::find_cell(cells, 4, col);
    ASSERT_NE(heading, nullptr) << col;
    EXPECT_TRUE(heading->value.is_blank()) << col;
    EXPECT_EQ(heading->role, PivotCellRole::DataValue) << col;
  }
}

TEST(PivotLayoutOptions, MissingCaptionTextFillsEmptyCellsWhateverShowMissingSays) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_cross_table_values_on_rows();
  table.set_missing_caption(std::string("N/A"));
  table.set_show_missing(false);
  const PivotCells cells = ProjectCells(table, cache);
  for (const std::uint32_t row : {5U, 6U}) {
    const PivotCell* hole = layout_test::find_cell(cells, row, 2);
    ASSERT_NE(hole, nullptr) << row;
    ASSERT_TRUE(hole->value.is_text()) << row;
    EXPECT_EQ(hole->value.as_text(), "N/A") << row;
  }
  const PivotCell* filled = layout_test::find_cell(cells, 5, 1);
  ASSERT_NE(filled, nullptr);
  ASSERT_TRUE(filled->value.is_number());
  EXPECT_EQ(filled->value.as_number(), 500.0);
}

/// Region rows, Product columns, sample variance of Amt: East's single record
/// gives #DIV/0!.
PivotTable build_variance_table() {
  PivotTable table = build_cross_table_values_on_rows();
  table.mutable_data_fields().pop_back();
  table.mutable_data_fields()[0].aggregation = Aggregation::Var;
  table.mutable_data_fields()[0].name = "Var of Amt";
  return table;
}

TEST(PivotLayoutOptions, ErrorCaptionReplacesErrorCellsOnlyWhileShown) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_variance_table();
  // Rows: 2 caption, 3 column items, 4 East; column 1 is item A.
  const PivotCells plain = ProjectCells(table, cache);
  const PivotCell* error = layout_test::find_cell(plain, 4, 1);
  ASSERT_NE(error, nullptr);
  ASSERT_TRUE(error->value.is_error());

  table.set_error_caption(std::string("ERR"));
  const PivotCells hidden = ProjectCells(table, cache);
  ASSERT_NE(layout_test::find_cell(hidden, 4, 1), nullptr);
  EXPECT_TRUE(layout_test::find_cell(hidden, 4, 1)->value.is_error());

  table.set_show_error(true);
  const PivotCells shown = ProjectCells(table, cache);
  const PivotCell* caption = layout_test::find_cell(shown, 4, 1);
  ASSERT_NE(caption, nullptr);
  ASSERT_TRUE(caption->value.is_text());
  EXPECT_EQ(caption->value.as_text(), "ERR");
  for (const PivotCell& cell : shown.cells) {
    EXPECT_FALSE(cell.value.is_error()) << cell.row << "," << cell.col;
  }

  table.set_error_caption(std::nullopt);
  const PivotCells blank = ProjectCells(table, cache);
  ASSERT_NE(layout_test::find_cell(blank, 4, 1), nullptr);
  EXPECT_TRUE(layout_test::find_cell(blank, 4, 1)->value.is_blank());
}

TEST(PivotLayoutOptions, CompactLabelIndentGrowsWithTheIndentOption) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Compact);
  // Rows: 2 header, 3 East, 4 A, ..., 11 grand total.
  for (const std::uint32_t indent : {0U, 1U, 2U}) {
    table.set_indent(indent);
    const PivotCells cells = ProjectCells(table, cache);
    const PivotCell* outer = layout_test::find_cell(cells, 3, 0);
    const PivotCell* inner = layout_test::find_cell(cells, 4, 0);
    const PivotCell* total = layout_test::find_cell(cells, 11, 0);
    ASSERT_NE(outer, nullptr);
    ASSERT_NE(inner, nullptr);
    ASSERT_NE(total, nullptr);
    EXPECT_EQ(outer->label_indent, std::optional<std::uint32_t>(0U)) << indent;
    EXPECT_EQ(inner->label_indent, std::optional<std::uint32_t>(indent + 1U)) << indent;
    EXPECT_EQ(total->label_indent, std::optional<std::uint32_t>(0U)) << indent;
  }
  const PivotCells tabular = ProjectCells(build_region_product_table(PivotLayout::Tabular), cache);
  for (const PivotCell& cell : tabular.cells) {
    EXPECT_FALSE(cell.label_indent.has_value()) << cell.row << "," << cell.col;
  }
}

std::vector<PivotRect> Merges(const PivotTable& table, const PivotCache& cache) {
  return ProjectCells(table, cache).merges;
}

TEST(PivotLayoutOptions, MergeItemInTabularMergesBlocksAndTotalLabels) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Tabular);
  EXPECT_TRUE(Merges(table, cache).empty());
  table.set_merge_item(true);
  // Rows: 3 East|A, 4 East 集計, 5-6 North, 7 North 集計, 8-9 South,
  // 10 South 集計, 11 総計.
  const std::vector<PivotRect> expected = {{4, 0, 4, 1}, {5, 0, 6, 0},   {7, 0, 7, 1},
                                           {8, 0, 9, 0}, {10, 0, 10, 1}, {11, 0, 11, 1}};
  EXPECT_EQ(Merges(table, cache), expected);
}

TEST(PivotLayoutOptions, MergeItemOutsideTabularMergesNothing) {
  const PivotCache cache = build_region_product_cache();
  for (const PivotLayout form : {PivotLayout::Outline, PivotLayout::Compact}) {
    PivotTable table = build_region_product_table(form);
    table.set_merge_item(true);
    EXPECT_TRUE(Merges(table, cache).empty());
  }
}

// ---------------------------------------------------------------------------
// Cell roles
// ---------------------------------------------------------------------------

PivotCellRole RoleAt(const PivotCells& cells, std::uint32_t row, std::uint32_t col) {
  const PivotCell* cell = layout_test::find_cell(cells, row, col);
  EXPECT_NE(cell, nullptr) << row << "," << col;
  return cell != nullptr ? cell->role : PivotCellRole::None;
}

TEST(PivotLayoutOptions, RolesOfACompactCrossTabWithAPageField) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_cross_table_values_on_rows();
  table.set_row_values_position(std::nullopt);
  table.mutable_data_fields().pop_back();
  table.mutable_fields()[3].axis = PivotAxis::Page;
  table.mutable_page_fields().push_back(PivotPageField{3, std::nullopt});
  const PivotCells cells = ProjectCells(table, cache);
  ASSERT_EQ(cells.top, 0U);
  // Row 0 page field, row 1 gap, row 2 caption row, row 3 column items,
  // rows 4-6 East / North / South, row 7 grand total.
  EXPECT_EQ(RoleAt(cells, 0, 0), PivotCellRole::PageCaption);
  EXPECT_EQ(RoleAt(cells, 0, 1), PivotCellRole::PageValue);
  EXPECT_EQ(RoleAt(cells, 1, 0), PivotCellRole::None);
  EXPECT_EQ(RoleAt(cells, 2, 0), PivotCellRole::FieldHeader);  // Sum of Amt
  EXPECT_EQ(RoleAt(cells, 2, 1), PivotCellRole::FieldHeader);  // 列ラベル
  EXPECT_EQ(RoleAt(cells, 2, 2), PivotCellRole::None);
  EXPECT_EQ(RoleAt(cells, 3, 0), PivotCellRole::FieldHeader);  // 行ラベル
  EXPECT_EQ(RoleAt(cells, 3, 1), PivotCellRole::ColItem);
  EXPECT_EQ(RoleAt(cells, 3, 3), PivotCellRole::GrandTotalLabel);
  EXPECT_EQ(RoleAt(cells, 4, 0), PivotCellRole::RowItem);
  EXPECT_EQ(RoleAt(cells, 4, 1), PivotCellRole::DataValue);
  EXPECT_EQ(RoleAt(cells, 4, 2), PivotCellRole::DataValue);  // East has no B
  EXPECT_EQ(RoleAt(cells, 4, 3), PivotCellRole::GrandTotalValue);
  EXPECT_EQ(RoleAt(cells, 7, 0), PivotCellRole::GrandTotalLabel);
  EXPECT_EQ(RoleAt(cells, 7, 1), PivotCellRole::GrandTotalValue);
}

TEST(PivotLayoutOptions, RolesOfATabularReportWithBlankRows) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Tabular);
  table.mutable_fields()[0].insert_blank_row = true;
  const PivotCells cells = ProjectCells(table, cache);
  // Rows: 2 header, 3 East|A, 4 East 集計, 5 blank, 6 North|A, 7 |B, ..., 14 総計.
  EXPECT_EQ(RoleAt(cells, 2, 0), PivotCellRole::FieldHeader);
  EXPECT_EQ(RoleAt(cells, 2, 1), PivotCellRole::FieldHeader);
  EXPECT_EQ(RoleAt(cells, 2, 2), PivotCellRole::DataHeader);
  EXPECT_EQ(RoleAt(cells, 3, 0), PivotCellRole::RowItem);
  EXPECT_EQ(RoleAt(cells, 3, 1), PivotCellRole::RowItem);
  EXPECT_EQ(RoleAt(cells, 3, 2), PivotCellRole::DataValue);
  EXPECT_EQ(RoleAt(cells, 4, 0), PivotCellRole::SubtotalLabel);
  EXPECT_EQ(RoleAt(cells, 4, 1), PivotCellRole::None);
  EXPECT_EQ(RoleAt(cells, 4, 2), PivotCellRole::SubtotalValue);
  EXPECT_EQ(RoleAt(cells, 5, 0), PivotCellRole::None);
  EXPECT_EQ(RoleAt(cells, 5, 2), PivotCellRole::DataValue);
  EXPECT_EQ(RoleAt(cells, 7, 0), PivotCellRole::None);  // North is not repeated
  EXPECT_EQ(RoleAt(cells, 7, 1), PivotCellRole::RowItem);
  EXPECT_EQ(RoleAt(cells, 14, 0), PivotCellRole::GrandTotalLabel);
  EXPECT_EQ(RoleAt(cells, 14, 1), PivotCellRole::None);
  EXPECT_EQ(RoleAt(cells, 14, 2), PivotCellRole::GrandTotalValue);
}

TEST(PivotLayoutOptions, RolesOfCompactRowsWithValuesAndBlankRows) {
  const PivotCache cache = build_region_product_cache();
  PivotTable values_on_rows = build_region_product_table(PivotLayout::Compact, true);
  values_on_rows.set_row_values_position(1);
  values_on_rows.set_indent(0);
  const PivotCells cells = ProjectCells(values_on_rows, cache);
  // Rows: 2 header, 3 East, 4 Sum of Amt, 5 Sum of Qty, ...
  EXPECT_EQ(RoleAt(cells, 2, 1), PivotCellRole::None);
  EXPECT_EQ(RoleAt(cells, 3, 0), PivotCellRole::RowItem);
  EXPECT_EQ(RoleAt(cells, 3, 1), PivotCellRole::DataValue);
  const PivotCell* data_label = layout_test::find_cell(cells, 4, 0);
  ASSERT_NE(data_label, nullptr);
  EXPECT_EQ(data_label->role, PivotCellRole::DataRowLabel);
  EXPECT_EQ(data_label->label_indent, std::optional<std::uint32_t>(1U));

  PivotTable blank_rows = build_region_product_table(PivotLayout::Compact);
  blank_rows.mutable_fields()[0].insert_blank_row = true;
  const PivotCells blanks = ProjectCells(blank_rows, cache);
  // Rows: 2 header, 3 East, 4 A, 5 blank.
  const PivotCell* blank_label = layout_test::find_cell(blanks, 5, 0);
  ASSERT_NE(blank_label, nullptr);
  EXPECT_EQ(blank_label->role, PivotCellRole::BlankRow);
  EXPECT_EQ(blank_label->label_indent, std::optional<std::uint32_t>(0U));
  EXPECT_EQ(RoleAt(blanks, 5, 1), PivotCellRole::DataValue);
}

// ---------------------------------------------------------------------------
// Item path of a value cell
// ---------------------------------------------------------------------------

/// `field:identity` per item, outermost first.
std::string Items(const std::vector<PivotItemRef>& items) {
  std::string out;
  for (const PivotItemRef& item : items) {
    out += (out.empty() ? "" : ",") + std::to_string(item.field_index) + ":" + CellText(item.identity) +
           (item.is_group_item ? "#" : "");
  }
  return out;
}

class PivotLayoutCellItemPathTest : public ::testing::Test {
 protected:
  void SetUp() override {
    cache_ = build_region_product_cache();
    table_ = build_region_product_table(PivotLayout::Tabular);
    table_.mutable_fields()[0].insert_blank_row = true;
    auto result_or = evaluate(table_, cache_, options_);
    ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
    result_ = std::move(result_or.value());
  }

  Expected<CellItemPath, Error> At(std::uint32_t row, std::uint32_t col) {
    return cell_item_path(table_, result_, options_, row, col);
  }

  // Grid rows (0-based sheet rows): 2 header, 3 East|A, 4 East 集計, 5 blank,
  // 6 North|A, 7 |B, 8 North 集計, 9 blank, 10 South|A, 11 |B, 12 South 集計,
  // 13 blank, 14 総計. Columns: 0 Region, 1 Product, 2 Sum of Amt, 3 Sum of Qty.
  PivotLayoutOptions options_ = ja_jp_layout_options();
  PivotCache cache_;
  PivotTable table_;
  PivotResult result_;
};

TEST_F(PivotLayoutCellItemPathTest, LeafCellListsEveryRowItem) {
  auto path_or = At(7, 2);
  ASSERT_TRUE(static_cast<bool>(path_or)) << path_or.error().message;
  EXPECT_EQ(path_or.value().kind, PivotCellKind::Data);
  EXPECT_EQ(Items(path_or.value().row_items), "0:North,1:B");
  EXPECT_TRUE(path_or.value().col_items.empty());

  auto qty_or = At(7, 3);
  ASSERT_TRUE(static_cast<bool>(qty_or)) << qty_or.error().message;
  EXPECT_EQ(Items(qty_or.value().row_items), "0:North,1:B");
}

TEST_F(PivotLayoutCellItemPathTest, SubtotalCellStopsAtItsField) {
  auto path_or = At(8, 3);
  ASSERT_TRUE(static_cast<bool>(path_or)) << path_or.error().message;
  EXPECT_EQ(path_or.value().kind, PivotCellKind::RowSubtotal);
  EXPECT_EQ(Items(path_or.value().row_items), "0:North");
}

TEST_F(PivotLayoutCellItemPathTest, GrandTotalCellListsNoItems) {
  auto path_or = At(14, 2);
  ASSERT_TRUE(static_cast<bool>(path_or)) << path_or.error().message;
  EXPECT_EQ(path_or.value().kind, PivotCellKind::GrandTotal);
  EXPECT_TRUE(path_or.value().row_items.empty());
  EXPECT_TRUE(path_or.value().col_items.empty());
}

TEST_F(PivotLayoutCellItemPathTest, CollapsedItemCellIsItsSubtotal) {
  table_.mutable_fields()[0].items[0].show_detail = false;
  // 3 East (collapsed), 4 blank, 5 North|A.
  auto path_or = At(3, 2);
  ASSERT_TRUE(static_cast<bool>(path_or)) << path_or.error().message;
  EXPECT_EQ(path_or.value().kind, PivotCellKind::RowSubtotal);
  EXPECT_EQ(Items(path_or.value().row_items), "0:East");
  auto next_or = At(5, 2);
  ASSERT_TRUE(static_cast<bool>(next_or)) << next_or.error().message;
  EXPECT_EQ(Items(next_or.value().row_items), "0:North,1:A");
}

TEST_F(PivotLayoutCellItemPathTest, NonValueCellsAreRejected) {
  for (const auto& [row, col] : std::vector<std::pair<std::uint32_t, std::uint32_t>>{
           {2, 2},   // header
           {7, 1},   // row label
           {5, 2},   // blank row
           {1, 2},   // above the report
           {15, 2},  // below the report
           {7, 4},   // right of the report
       }) {
    auto path_or = At(row, col);
    ASSERT_FALSE(static_cast<bool>(path_or)) << row << "," << col;
    EXPECT_EQ(path_or.error().code, FormulonErrorCode::kInvalidArgument) << row << "," << col;
  }
}

TEST(PivotLayoutCellItemPath, ColumnItemsAndTheTotalsStrip) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_table({0}, {1});
  const PivotLayoutOptions options = ja_jp_layout_options();
  auto result_or = evaluate(table, cache, options);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  auto cells_or = layout(table, result_or.value(), options);
  ASSERT_TRUE(static_cast<bool>(cells_or)) << cells_or.error().message;
  // Find the North row and the Widget / grand-total columns from the grid itself.
  std::uint32_t north_row = 0;
  std::uint32_t widget_col = 0;
  std::uint32_t total_col = 0;
  std::uint32_t total_row = 0;
  for (const PivotCell& cell : cells_or.value().cells) {
    if (!cell.value.is_text()) {
      continue;
    }
    const std::string text(cell.value.as_text());
    if (text == "North") {
      north_row = cell.row;
    } else if (text == "Widget") {
      widget_col = cell.col;
    } else if (text == TOTAL && cell.kind == PivotCellKind::Header) {
      total_col = cell.col;
    } else if (text == TOTAL) {
      total_row = cell.row;
    }
  }
  ASSERT_NE(north_row, 0U);
  ASSERT_NE(widget_col, 0U);
  ASSERT_NE(total_col, 0U);
  ASSERT_NE(total_row, 0U);

  auto leaf_or = cell_item_path(table, result_or.value(), options, north_row, widget_col);
  ASSERT_TRUE(static_cast<bool>(leaf_or)) << leaf_or.error().message;
  EXPECT_EQ(leaf_or.value().kind, PivotCellKind::Data);
  EXPECT_EQ(Items(leaf_or.value().row_items), "0:North");
  EXPECT_EQ(Items(leaf_or.value().col_items), "1:Widget");

  auto row_total_or = cell_item_path(table, result_or.value(), options, north_row, total_col);
  ASSERT_TRUE(static_cast<bool>(row_total_or)) << row_total_or.error().message;
  EXPECT_EQ(row_total_or.value().kind, PivotCellKind::GrandTotal);
  EXPECT_EQ(Items(row_total_or.value().row_items), "0:North");
  EXPECT_TRUE(row_total_or.value().col_items.empty());

  auto col_total_or = cell_item_path(table, result_or.value(), options, total_row, widget_col);
  ASSERT_TRUE(static_cast<bool>(col_total_or)) << col_total_or.error().message;
  EXPECT_EQ(col_total_or.value().kind, PivotCellKind::GrandTotal);
  EXPECT_TRUE(col_total_or.value().row_items.empty());
  EXPECT_EQ(Items(col_total_or.value().col_items), "1:Widget");

  auto corner_or = cell_item_path(table, result_or.value(), options, total_row, total_col);
  ASSERT_TRUE(static_cast<bool>(corner_or)) << corner_or.error().message;
  EXPECT_EQ(corner_or.value().kind, PivotCellKind::GrandTotal);
  EXPECT_TRUE(corner_or.value().row_items.empty());
  EXPECT_TRUE(corner_or.value().col_items.empty());
}

TEST(PivotLayoutCellItemPath, PathOutlivesTheResult) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_table({0}, {1});
  const PivotLayoutOptions options = ja_jp_layout_options();
  std::optional<CellItemPath> path;
  {
    auto result_or = evaluate(table, cache, options);
    ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
    // ja-JP compact with a column field: corner row, column-label row, then North.
    auto path_or = cell_item_path(table, result_or.value(), options, 4, 4);
    ASSERT_TRUE(static_cast<bool>(path_or)) << path_or.error().message;
    path = std::move(path_or.value());
  }
  EXPECT_EQ(Items(path->row_items), "0:North");
  EXPECT_EQ(Items(path->col_items), "1:Gadget");
}

}  // namespace
}  // namespace formulon::pivot
