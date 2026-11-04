// Report options in the layout projection: the grand-total caption, blank
// rows after item blocks, repeated (filled-down) labels, collapsed items,
// number-formatted item labels, and the item path behind a value cell.

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
