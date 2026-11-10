// `<rowItems>` / `<colItems>` lines read off the projected layout. Each
// expected string is the element body Excel saved for the same report, with
// `v` taken as the item's position in the field's written `<items>`.

#include <cstddef>
#include <cstdint>
#include <memory>
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

using layout_test::ja_jp_layout_options;
using layout_test::owned_text;

constexpr std::uint32_t kCacheId = 7;

/// Labels of each field's written `<items>`, by field index.
using WrittenItems = std::vector<std::vector<std::string>>;

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

/// Region items as Excel wrote them (East, North, South), Product (A, B), Qty (1..5).
const WrittenItems kRegionProductItems = {{"East", "North", "South"}, {"A", "B"}, {}, {"1", "2", "3", "4", "5"}};

/// Region (and Product) on the rows, Sum of Amt and Sum of Qty on the columns.
PivotTable build_region_product_table(PivotLayout layout, std::vector<std::uint32_t> rows = {0, 1}) {
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
  for (const char* name : {"East", "North", "South"}) {
    PivotItem item;
    item.name = name;
    table.mutable_fields()[0].items.push_back(std::move(item));
  }
  for (const char* name : {"A", "B"}) {
    PivotItem item;
    item.name = name;
    table.mutable_fields()[1].items.push_back(std::move(item));
  }
  for (const std::uint32_t field : rows) {
    table.mutable_fields()[field].axis = PivotAxis::Row;
  }
  table.mutable_row_field_order() = std::move(rows);
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

/// Moves the row fields `cols` onto the column axis.
void move_to_columns(PivotTable& table, std::vector<std::uint32_t> cols) {
  for (const std::uint32_t field : cols) {
    table.mutable_fields()[field].axis = PivotAxis::Col;
  }
  table.mutable_col_field_order() = std::move(cols);
}

const char* token(PivotLineType type) {
  switch (type) {
    case PivotLineType::kData:
      return "";
    case PivotLineType::kDefault:
      return "default";
    case PivotLineType::kSum:
      return "sum";
    case PivotLineType::kCountA:
      return "countA";
    case PivotLineType::kAverage:
      return "avg";
    case PivotLineType::kMax:
      return "max";
    case PivotLineType::kMin:
      return "min";
    case PivotLineType::kProduct:
      return "product";
    case PivotLineType::kCount:
      return "count";
    case PivotLineType::kStdDev:
      return "stdDev";
    case PivotLineType::kStdDevP:
      return "stdDevP";
    case PivotLineType::kVar:
      return "var";
    case PivotLineType::kVarP:
      return "varP";
    case PivotLineType::kGrand:
      return "grand";
    case PivotLineType::kBlank:
      return "blank";
  }
  return "?";
}

std::string encode_x(const PivotLineX& x, const WrittenItems& written) {
  std::string v;
  if (x.values_level) {
    v = std::to_string(x.data_index);
  } else if (x.node == nullptr) {
    v = "0";
  } else {
    v = "?";
    if (x.field_index < written.size()) {
      const std::vector<std::string>& labels = written[x.field_index];
      for (std::size_t p = 0; p < labels.size(); ++p) {
        if (labels[p] == x.node->label) {
          v = std::to_string(p);
          break;
        }
      }
    }
  }
  return v == "0" ? std::string("<x/>") : "<x v=\"" + v + "\"/>";
}

/// The lines as the body of `<rowItems>` / `<colItems>`.
std::string encode(const std::vector<PivotLine>& lines, const WrittenItems& written) {
  std::string out;
  for (const PivotLine& line : lines) {
    out += "<i";
    if (line.type != PivotLineType::kData) {
      out += std::string(" t=\"") + token(line.type) + "\"";
    }
    if (line.r != 0) {
      out += " r=\"" + std::to_string(line.r) + "\"";
    }
    if (line.i != 0) {
      out += " i=\"" + std::to_string(line.i) + "\"";
    }
    if (line.xs.empty()) {
      out += "/>";
      continue;
    }
    out += ">";
    for (const PivotLineX& x : line.xs) {
      out += encode_x(x, written);
    }
    out += "</i>";
  }
  return out;
}

/// The evaluated result is kept beside the lines, whose nodes point into it.
struct Projection {
  PivotResult result;
  PivotLineItems items;
};

std::unique_ptr<Projection> project(const PivotTable& table, const PivotCache& cache,
                                    const PivotLayoutOptions& options = ja_jp_layout_options()) {
  auto result_or = evaluate(table, cache, options);
  EXPECT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  if (!result_or) {
    return nullptr;
  }
  auto projection = std::make_unique<Projection>();
  projection->result = std::move(result_or.value());
  auto cells_or = layout(table, projection->result, options, &projection->items);
  EXPECT_TRUE(static_cast<bool>(cells_or)) << cells_or.error().message << " " << cells_or.error().context;
  if (!cells_or) {
    return nullptr;
  }
  return projection;
}

void expect_refused(const std::optional<Error>& refusal) {
  ASSERT_TRUE(refusal.has_value());
  EXPECT_EQ(refusal->code, FormulonErrorCode::kNotImplemented);
  EXPECT_FALSE(refusal->message.empty());
}

constexpr const char* kTwoDataFieldCols = "<i><x/></i><i i=\"1\"><x v=\"1\"/></i>";
constexpr const char* kRegionRows = "<i><x/></i><i><x v=\"1\"/></i><i><x v=\"2\"/></i><i t=\"grand\"><x/></i>";
constexpr const char* kOutlineTwoRows =
    "<i><x/></i><i r=\"1\"><x/></i><i><x v=\"1\"/></i><i r=\"1\"><x/></i><i r=\"1\"><x v=\"1\"/></i>"
    "<i><x v=\"2\"/></i><i r=\"1\"><x/></i><i r=\"1\"><x v=\"1\"/></i><i t=\"grand\"><x/></i>";

// ---------------------------------------------------------------------------
// Row forms
// ---------------------------------------------------------------------------

TEST(PivotLineItems, OneRowFieldIsOneLinePerItemInCompactAndOutline) {
  const PivotCache cache = build_region_product_cache();
  for (const PivotLayout form : {PivotLayout::Compact, PivotLayout::Outline}) {
    auto p = project(build_region_product_table(form, {0}), cache);
    ASSERT_NE(p, nullptr);
    EXPECT_FALSE(p->items.row_refusal.has_value());
    EXPECT_FALSE(p->items.col_refusal.has_value());
    EXPECT_EQ(encode(p->items.rows, kRegionProductItems), kRegionRows);
    EXPECT_EQ(encode(p->items.cols, kRegionProductItems), kTwoDataFieldCols);
  }
}

TEST(PivotLineItems, GroupRowsGiveEachLineOneItemAtItsLevel) {
  const PivotCache cache = build_region_product_cache();
  for (const PivotLayout form : {PivotLayout::Outline, PivotLayout::Compact}) {
    auto p = project(build_region_product_table(form), cache);
    ASSERT_NE(p, nullptr);
    EXPECT_EQ(encode(p->items.rows, kRegionProductItems), kOutlineTwoRows);
    EXPECT_EQ(encode(p->items.cols, kRegionProductItems), kTwoDataFieldCols);
  }
}

TEST(PivotLineItems, TabularRowsListThePathFromTheFirstNewLevel) {
  const PivotCache cache = build_region_product_cache();
  auto p = project(build_region_product_table(PivotLayout::Tabular), cache);
  ASSERT_NE(p, nullptr);
  EXPECT_EQ(encode(p->items.rows, kRegionProductItems),
            "<i><x/><x/></i><i t=\"default\"><x/></i><i><x v=\"1\"/><x/></i><i r=\"1\"><x v=\"1\"/></i>"
            "<i t=\"default\"><x v=\"1\"/></i><i><x v=\"2\"/><x/></i><i r=\"1\"><x v=\"1\"/></i>"
            "<i t=\"default\"><x v=\"2\"/></i><i t=\"grand\"><x/></i>");
  EXPECT_EQ(encode(p->items.cols, kRegionProductItems), kTwoDataFieldCols);
}

// ---------------------------------------------------------------------------
// The Values level
// ---------------------------------------------------------------------------

TEST(PivotLineItems, ValuesInnermostOnRowsTagsEveryLineOfADataField) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Outline);
  table.set_row_values_position(2);
  auto p = project(table, cache);
  ASSERT_NE(p, nullptr);
  EXPECT_EQ(encode(p->items.rows, kRegionProductItems),
            "<i><x/></i><i r=\"1\"><x/></i><i r=\"2\"><x/></i><i r=\"2\" i=\"1\"><x v=\"1\"/></i>"
            "<i t=\"default\"><x/></i><i t=\"default\" i=\"1\"><x/></i><i><x v=\"1\"/></i><i r=\"1\"><x/></i>"
            "<i r=\"2\"><x/></i><i r=\"2\" i=\"1\"><x v=\"1\"/></i><i r=\"1\"><x v=\"1\"/></i><i r=\"2\"><x/></i>"
            "<i r=\"2\" i=\"1\"><x v=\"1\"/></i><i t=\"default\"><x v=\"1\"/></i>"
            "<i t=\"default\" i=\"1\"><x v=\"1\"/></i><i><x v=\"2\"/></i><i r=\"1\"><x/></i><i r=\"2\"><x/></i>"
            "<i r=\"2\" i=\"1\"><x v=\"1\"/></i><i r=\"1\"><x v=\"1\"/></i><i r=\"2\"><x/></i>"
            "<i r=\"2\" i=\"1\"><x v=\"1\"/></i><i t=\"default\"><x v=\"2\"/></i>"
            "<i t=\"default\" i=\"1\"><x v=\"2\"/></i><i t=\"grand\"><x/></i><i t=\"grand\" i=\"1\"><x/></i>");
  EXPECT_EQ(encode(p->items.cols, kRegionProductItems), "<i/>");
}

TEST(PivotLineItems, ValuesFirstOnRowsHeadsEachDataFieldBlock) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Outline, {0});
  table.set_row_values_position(0);
  auto p = project(table, cache);
  ASSERT_NE(p, nullptr);
  EXPECT_EQ(encode(p->items.rows, kRegionProductItems),
            "<i><x/></i><i r=\"1\"><x/></i><i r=\"1\"><x v=\"1\"/></i><i r=\"1\"><x v=\"2\"/></i>"
            "<i i=\"1\"><x v=\"1\"/></i><i r=\"1\" i=\"1\"><x/></i><i r=\"1\" i=\"1\"><x v=\"1\"/></i>"
            "<i r=\"1\" i=\"1\"><x v=\"2\"/></i><i t=\"grand\"><x/></i><i t=\"grand\" i=\"1\"><x/></i>");
  EXPECT_EQ(encode(p->items.cols, kRegionProductItems), "<i/>");
}

TEST(PivotLineItems, ValuesAfterAColumnField) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Outline, {0});
  move_to_columns(table, {1});
  table.set_col_values_position(1);
  auto p = project(table, cache);
  ASSERT_NE(p, nullptr);
  EXPECT_EQ(encode(p->items.rows, kRegionProductItems), kRegionRows);
  EXPECT_EQ(encode(p->items.cols, kRegionProductItems),
            "<i><x/><x/></i><i r=\"1\" i=\"1\"><x v=\"1\"/></i><i><x v=\"1\"/><x/></i>"
            "<i r=\"1\" i=\"1\"><x v=\"1\"/></i><i t=\"grand\"><x/></i><i t=\"grand\" i=\"1\"><x/></i>");
}

TEST(PivotLineItems, ValuesBeforeAColumnField) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Outline, {0});
  move_to_columns(table, {1});
  table.set_col_values_position(0);
  auto p = project(table, cache);
  ASSERT_NE(p, nullptr);
  EXPECT_EQ(encode(p->items.rows, kRegionProductItems), kRegionRows);
  EXPECT_EQ(encode(p->items.cols, kRegionProductItems),
            "<i><x/><x/></i><i r=\"1\"><x v=\"1\"/></i><i i=\"1\"><x v=\"1\"/><x/></i>"
            "<i r=\"1\" i=\"1\"><x v=\"1\"/></i><i t=\"grand\"><x/></i><i t=\"grand\" i=\"1\"><x/></i>");
}

TEST(PivotLineItems, TwoColumnFieldsWithValuesInnermost) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Outline, {});
  move_to_columns(table, {0, 1});
  table.set_col_values_position(2);
  auto p = project(table, cache);
  ASSERT_NE(p, nullptr);
  EXPECT_EQ(encode(p->items.rows, kRegionProductItems), "<i/>");
  EXPECT_EQ(encode(p->items.cols, kRegionProductItems),
            "<i><x/><x/><x/></i><i r=\"2\" i=\"1\"><x v=\"1\"/></i><i t=\"default\"><x/></i>"
            "<i t=\"default\" i=\"1\"><x/></i><i><x v=\"1\"/><x/><x/></i><i r=\"2\" i=\"1\"><x v=\"1\"/></i>"
            "<i r=\"1\"><x v=\"1\"/><x/></i><i r=\"2\" i=\"1\"><x v=\"1\"/></i><i t=\"default\"><x v=\"1\"/></i>"
            "<i t=\"default\" i=\"1\"><x v=\"1\"/></i><i><x v=\"2\"/><x/><x/></i>"
            "<i r=\"2\" i=\"1\"><x v=\"1\"/></i><i r=\"1\"><x v=\"1\"/><x/></i><i r=\"2\" i=\"1\"><x v=\"1\"/></i>"
            "<i t=\"default\"><x v=\"2\"/></i><i t=\"default\" i=\"1\"><x v=\"2\"/></i><i t=\"grand\"><x/></i>"
            "<i t=\"grand\" i=\"1\"><x/></i>");
}

TEST(PivotLineItems, EmptyAxesAreOneEmptyLineOrOneLinePerDataField) {
  const PivotCache cache = build_region_product_cache();
  auto p = project(build_region_product_table(PivotLayout::Outline, {}), cache);
  ASSERT_NE(p, nullptr);
  EXPECT_FALSE(p->items.row_refusal.has_value());
  EXPECT_FALSE(p->items.col_refusal.has_value());
  EXPECT_EQ(encode(p->items.rows, kRegionProductItems), "<i/>");
  EXPECT_EQ(encode(p->items.cols, kRegionProductItems), kTwoDataFieldCols);
}

// ---------------------------------------------------------------------------
// Subtotals, blank rows, grand totals
// ---------------------------------------------------------------------------

PivotCache build_c9_cache() {
  PivotCache cache;
  cache.set_cache_id(kCacheId);
  for (const char* name : {"Region", "Product", "Grp", "Q", "Amt"}) {
    cache.mutable_fields().push_back(PivotCacheField{name, {}});
  }
  auto add = [&](const char* region, const char* product, const char* grp, double q) {
    PivotCacheRecord rec;
    rec.cells.push_back(owned_text(cache, region));
    rec.cells.push_back(owned_text(cache, product));
    rec.cells.push_back(owned_text(cache, grp));
    rec.cells.push_back(Value::number(q));
    rec.cells.push_back(Value::number(1.0));
    cache.mutable_records().push_back(std::move(rec));
  };
  add("N", "A", "X", 1);
  add("N", "B", "X", 1);
  add("S", "A", "Y", 2);
  add("S", "B", "X", 1);
  add("S", "B", "X", 1);
  add("S", "A", "Y", 1);
  add("S", "A", "X", 1);
  add("S", "B", "Y", 1);
  add("S", "A", "Y", 1);
  add("N", "B", "Y", 2);
  add("N", "A", "Y", 1);
  add("N", "B", "Y", 1);
  add("S", "A", "Y", 1);
  add("N", "A", "X", 2);
  add("S", "A", "X", 2);
  add("S", "A", "X", 2);
  add("S", "B", "X", 1);
  return cache;
}

TEST(PivotLineItems, CustomSubtotalsAreOneTypedLineEachInMarkerOrder) {
  const PivotCache cache = build_c9_cache();
  PivotTable table;
  table.set_name("PivotTable1");
  table.set_pivot_cache_id(kCacheId);
  table.set_anchor(2, 0, 1, 1);
  for (const char* name : {"Region", "Product", "Grp", "Q", "Amt"}) {
    PivotField field;
    field.source_name = name;
    field.axis = PivotAxis::Row;
    table.mutable_fields().push_back(std::move(field));
  }
  table.mutable_fields()[3].axis = PivotAxis::Col;
  table.mutable_fields()[4].axis = PivotAxis::Value;
  table.mutable_fields()[0].subtotal_fns = {SubtotalFn::Average, SubtotalFn::Product, SubtotalFn::CountNumbers};
  table.mutable_fields()[1].subtotal_fns = {SubtotalFn::Sum, SubtotalFn::Var};
  table.mutable_row_field_order() = {0, 1, 2};
  table.mutable_col_field_order() = {3};
  PivotDataField sum;
  sum.name = "Sum of Amt";
  sum.field_index = 4;
  sum.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum));
  table.set_layout(PivotLayout::Outline);
  // Excel lists three data-less items before N and S.
  const WrittenItems written = {{"North", "South", "", "N", "S"}, {"A", "B"}, {"X", "Y"}, {"1", "2"}};

  auto p = project(table, cache);
  ASSERT_NE(p, nullptr);
  EXPECT_FALSE(p->items.row_refusal.has_value());
  EXPECT_EQ(encode(p->items.rows, written),
            "<i><x v=\"3\"/></i><i r=\"1\"><x/></i><i r=\"2\"><x/></i><i r=\"2\"><x v=\"1\"/></i>"
            "<i t=\"sum\" r=\"1\"><x/></i><i t=\"var\" r=\"1\"><x/></i><i r=\"1\"><x v=\"1\"/></i>"
            "<i r=\"2\"><x/></i><i r=\"2\"><x v=\"1\"/></i><i t=\"sum\" r=\"1\"><x v=\"1\"/></i>"
            "<i t=\"var\" r=\"1\"><x v=\"1\"/></i><i t=\"avg\"><x v=\"3\"/></i><i t=\"product\"><x v=\"3\"/></i>"
            "<i t=\"count\"><x v=\"3\"/></i><i><x v=\"4\"/></i><i r=\"1\"><x/></i><i r=\"2\"><x/></i>"
            "<i r=\"2\"><x v=\"1\"/></i><i t=\"sum\" r=\"1\"><x/></i><i t=\"var\" r=\"1\"><x/></i>"
            "<i r=\"1\"><x v=\"1\"/></i><i r=\"2\"><x/></i><i r=\"2\"><x v=\"1\"/></i>"
            "<i t=\"sum\" r=\"1\"><x v=\"1\"/></i><i t=\"var\" r=\"1\"><x v=\"1\"/></i><i t=\"avg\"><x v=\"4\"/></i>"
            "<i t=\"product\"><x v=\"4\"/></i><i t=\"count\"><x v=\"4\"/></i><i t=\"grand\"><x/></i>");
  EXPECT_EQ(encode(p->items.cols, written), "<i><x/></i><i><x v=\"1\"/></i><i t=\"grand\"><x/></i>");
}

TEST(PivotLineItems, BlankLineNamesTheOuterItemItCloses) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Compact);
  table.mutable_fields()[0].insert_blank_row = true;
  auto p = project(table, cache);
  ASSERT_NE(p, nullptr);
  EXPECT_FALSE(p->items.row_refusal.has_value());
  EXPECT_EQ(encode(p->items.rows, kRegionProductItems),
            "<i><x/></i><i r=\"1\"><x/></i><i t=\"blank\"><x/></i><i><x v=\"1\"/></i><i r=\"1\"><x/></i>"
            "<i r=\"1\"><x v=\"1\"/></i><i t=\"blank\"><x v=\"1\"/></i><i><x v=\"2\"/></i><i r=\"1\"><x/></i>"
            "<i r=\"1\"><x v=\"1\"/></i><i t=\"blank\"><x v=\"2\"/></i><i t=\"grand\"><x/></i>");
}

TEST(PivotLineItems, TabularBlankLinesCloseOuterAndMiddleBlocks) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Tabular, {0, 1, 3});
  table.mutable_fields()[0].insert_blank_row = true;
  table.mutable_fields()[1].insert_blank_row = true;
  auto p = project(table, cache);
  ASSERT_NE(p, nullptr);
  EXPECT_FALSE(p->items.row_refusal.has_value());
  EXPECT_EQ(encode(p->items.rows, kRegionProductItems),
            "<i><x/><x/><x v=\"4\"/></i><i t=\"default\" r=\"1\"><x/></i><i t=\"blank\" r=\"1\"><x/></i>"
            "<i t=\"default\"><x/></i><i t=\"blank\"><x/></i><i><x v=\"1\"/><x/><x/></i>"
            "<i t=\"default\" r=\"1\"><x/></i><i t=\"blank\" r=\"1\"><x/></i><i r=\"1\"><x v=\"1\"/><x v=\"1\"/></i>"
            "<i t=\"default\" r=\"1\"><x v=\"1\"/></i><i t=\"blank\" r=\"1\"><x v=\"1\"/></i>"
            "<i t=\"default\"><x v=\"1\"/></i><i t=\"blank\"><x v=\"1\"/></i><i><x v=\"2\"/><x/><x v=\"2\"/></i>"
            "<i t=\"default\" r=\"1\"><x/></i><i t=\"blank\" r=\"1\"><x/></i><i r=\"1\"><x v=\"1\"/><x v=\"3\"/></i>"
            "<i t=\"default\" r=\"1\"><x v=\"1\"/></i><i t=\"blank\" r=\"1\"><x v=\"1\"/></i>"
            "<i t=\"default\"><x v=\"2\"/></i><i t=\"blank\"><x v=\"2\"/></i><i t=\"grand\"><x/></i>");
  EXPECT_EQ(encode(p->items.cols, kRegionProductItems), kTwoDataFieldCols);
}

TEST(PivotLineItems, RowGrandTotalsOffDropsTheGrandLine) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Tabular);
  table.mutable_fields()[0].insert_blank_row = true;
  table.set_grand_totals(/*rows=*/true, /*cols=*/false);
  auto p = project(table, cache);
  ASSERT_NE(p, nullptr);
  EXPECT_EQ(encode(p->items.rows, kRegionProductItems),
            "<i><x/><x/></i><i t=\"default\"><x/></i><i t=\"blank\"><x/></i><i><x v=\"1\"/><x/></i>"
            "<i r=\"1\"><x v=\"1\"/></i><i t=\"default\"><x v=\"1\"/></i><i t=\"blank\"><x v=\"1\"/></i>"
            "<i><x v=\"2\"/><x/></i><i r=\"1\"><x v=\"1\"/></i><i t=\"default\"><x v=\"2\"/></i>"
            "<i t=\"blank\"><x v=\"2\"/></i>");
  EXPECT_FALSE(p->items.col_refusal.has_value());
  EXPECT_EQ(encode(p->items.cols, kRegionProductItems), kTwoDataFieldCols);
}

// ---------------------------------------------------------------------------
// Collapsed items
// ---------------------------------------------------------------------------

TEST(PivotLineItems, CollapsedOuterItemInOutlineIsItsHeaderLineAlone) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Outline);
  table.mutable_fields()[0].items[0].show_detail = false;
  auto p = project(table, cache);
  ASSERT_NE(p, nullptr);
  EXPECT_FALSE(p->items.row_refusal.has_value());
  EXPECT_EQ(encode(p->items.rows, kRegionProductItems),
            "<i><x/></i><i><x v=\"1\"/></i><i r=\"1\"><x/></i><i r=\"1\"><x v=\"1\"/></i><i><x v=\"2\"/></i>"
            "<i r=\"1\"><x/></i><i r=\"1\"><x v=\"1\"/></i><i t=\"grand\"><x/></i>");
}

TEST(PivotLineItems, CollapsedOuterItemInTabularStopsAtItsLevel) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Tabular);
  table.mutable_fields()[0].items[0].show_detail = false;
  auto p = project(table, cache);
  ASSERT_NE(p, nullptr);
  EXPECT_FALSE(p->items.row_refusal.has_value());
  EXPECT_EQ(encode(p->items.rows, kRegionProductItems),
            "<i><x/></i><i><x v=\"1\"/><x/></i><i r=\"1\"><x v=\"1\"/></i><i t=\"default\"><x v=\"1\"/></i>"
            "<i><x v=\"2\"/><x/></i><i r=\"1\"><x v=\"1\"/></i><i t=\"default\"><x v=\"2\"/></i>"
            "<i t=\"grand\"><x/></i>");
}

TEST(PivotLineItems, CollapsedItemOfTheOnlyRowFieldChangesNothing) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Compact, {0});
  table.mutable_fields()[0].items[0].show_detail = false;
  auto p = project(table, cache);
  ASSERT_NE(p, nullptr);
  EXPECT_EQ(encode(p->items.rows, kRegionProductItems), kRegionRows);
}

// ---------------------------------------------------------------------------
// Refused shapes: the axis is refused, the other axis is still produced
// ---------------------------------------------------------------------------

TEST(PivotLineItems, RefusesACollapsedRowItemWithValuesBelow) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Outline);
  table.set_row_values_position(2);
  table.mutable_fields()[0].items[0].show_detail = false;
  auto p = project(table, cache);
  ASSERT_NE(p, nullptr);
  expect_refused(p->items.row_refusal);
  EXPECT_TRUE(p->items.rows.empty());
  EXPECT_FALSE(p->items.col_refusal.has_value());
  EXPECT_EQ(encode(p->items.cols, kRegionProductItems), "<i/>");
}

TEST(PivotLineItems, RefusesACollapsedColumnItemWithValuesBelow) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Outline, {});
  move_to_columns(table, {0, 1});
  table.set_col_values_position(2);
  table.mutable_fields()[0].items[1].show_detail = false;
  auto p = project(table, cache);
  ASSERT_NE(p, nullptr);
  expect_refused(p->items.col_refusal);
  EXPECT_FALSE(p->items.row_refusal.has_value());
  EXPECT_EQ(encode(p->items.rows, kRegionProductItems), "<i/>");
}

TEST(PivotLineItems, RefusesACollapsedInnerItemInTabular) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Tabular, {0, 1, 3});
  table.mutable_fields()[1].items[0].show_detail = false;
  auto p = project(table, cache);
  ASSERT_NE(p, nullptr);
  expect_refused(p->items.row_refusal);
  EXPECT_EQ(encode(p->items.cols, kRegionProductItems), kTwoDataFieldCols);
}

TEST(PivotLineItems, RefusesRealFieldsWithChildrenBelowTheValuesLevel) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Outline);
  table.set_row_values_position(0);
  auto p = project(table, cache);
  ASSERT_NE(p, nullptr);
  expect_refused(p->items.row_refusal);
  EXPECT_EQ(encode(p->items.cols, kRegionProductItems), "<i/>");
}

TEST(PivotLineItems, RefusesCustomSubtotalsWithValuesBelow) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Outline);
  table.set_row_values_position(2);
  table.mutable_fields()[0].subtotal_fns = {SubtotalFn::Sum, SubtotalFn::Average};
  auto p = project(table, cache);
  ASSERT_NE(p, nullptr);
  expect_refused(p->items.row_refusal);
  EXPECT_EQ(encode(p->items.cols, kRegionProductItems), "<i/>");
}

TEST(PivotLineItems, RefusesCustomSubtotalsOnTheColumnAxis) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Outline, {});
  move_to_columns(table, {0, 1});
  table.mutable_data_fields().pop_back();
  table.mutable_fields()[0].subtotal_fns = {SubtotalFn::Max};
  auto p = project(table, cache);
  ASSERT_NE(p, nullptr);
  expect_refused(p->items.col_refusal);
  EXPECT_FALSE(p->items.row_refusal.has_value());
  EXPECT_EQ(encode(p->items.rows, kRegionProductItems), "<i/>");
}

TEST(PivotLineItems, RefusesARowFieldWithItsOwnForm) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Compact);
  table.mutable_fields()[0].passthrough_attrs.emplace_back("compact", "0");
  auto p = project(table, cache);
  ASSERT_NE(p, nullptr);
  expect_refused(p->items.row_refusal);
  EXPECT_EQ(encode(p->items.cols, kRegionProductItems), kTwoDataFieldCols);
}

TEST(PivotLineItems, RefusesAReportWithNoDataFieldsAndNoFields) {
  PivotTable table;
  table.set_name("PivotTable1");
  table.set_pivot_cache_id(kCacheId);
  table.set_anchor(2, 0, 1, 1);
  PivotResult result;
  result.values = {{{}}};
  PivotLineItems items;
  auto cells_or = layout(table, result, ja_jp_layout_options(), &items);
  ASSERT_TRUE(static_cast<bool>(cells_or)) << cells_or.error().message;
  expect_refused(items.row_refusal);
  expect_refused(items.col_refusal);
}

TEST(PivotLineItems, RefusesAValuesPositionWithOneDataField) {
  const PivotCache cache = build_region_product_cache();
  PivotTable rows = build_region_product_table(PivotLayout::Outline, {0});
  rows.mutable_data_fields().pop_back();
  rows.set_row_values_position(1);
  auto p = project(rows, cache);
  ASSERT_NE(p, nullptr);
  expect_refused(p->items.row_refusal);
  EXPECT_FALSE(p->items.col_refusal.has_value());

  PivotTable cols = build_region_product_table(PivotLayout::Outline, {0});
  cols.mutable_data_fields().pop_back();
  cols.set_col_values_position(0);
  auto q = project(cols, cache);
  ASSERT_NE(q, nullptr);
  expect_refused(q->items.col_refusal);
  EXPECT_FALSE(q->items.row_refusal.has_value());
  EXPECT_EQ(encode(q->items.rows, kRegionProductItems), kRegionRows);
}

TEST(PivotLineItems, RefusesTheLegacyProjection) {
  const PivotCache cache = build_region_product_cache();
  const PivotLayoutOptions legacy{};
  auto p = project(build_region_product_table(PivotLayout::Compact), cache, legacy);
  ASSERT_NE(p, nullptr);
  expect_refused(p->items.row_refusal);
  expect_refused(p->items.col_refusal);
}

TEST(PivotLineItems, RefusesASubtotalBelowItsGroupInCompactAndOutline) {
  const PivotCache cache = build_region_product_cache();
  for (const PivotLayout form : {PivotLayout::Compact, PivotLayout::Outline}) {
    PivotTable table = build_region_product_table(form);
    table.mutable_fields()[0].subtotal_top = false;
    auto p = project(table, cache);
    ASSERT_NE(p, nullptr);
    expect_refused(p->items.row_refusal);
    EXPECT_EQ(encode(p->items.cols, kRegionProductItems), kTwoDataFieldCols);
  }
}

TEST(PivotLineItems, RefusesAnInnerBlankRowInCompactAndOutline) {
  const PivotCache cache = build_region_product_cache();
  for (const PivotLayout form : {PivotLayout::Compact, PivotLayout::Outline}) {
    PivotTable table = build_region_product_table(form, {0, 1, 3});
    table.mutable_fields()[1].insert_blank_row = true;
    auto p = project(table, cache);
    ASSERT_NE(p, nullptr);
    expect_refused(p->items.row_refusal);
    EXPECT_EQ(encode(p->items.cols, kRegionProductItems), kTwoDataFieldCols);
  }
}

TEST(PivotLineItems, RefusesAFieldListingItemsWithoutData) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Outline);
  table.mutable_fields()[1].show_all = true;
  auto p = project(table, cache);
  ASSERT_NE(p, nullptr);
  expect_refused(p->items.row_refusal);
  EXPECT_FALSE(p->items.col_refusal.has_value());
  EXPECT_EQ(encode(p->items.cols, kRegionProductItems), kTwoDataFieldCols);
}

TEST(PivotLineItems, RefusesColumnGrandTotalsOff) {
  const PivotCache cache = build_region_product_cache();
  PivotTable table = build_region_product_table(PivotLayout::Outline);
  table.set_grand_totals(/*rows=*/false, /*cols=*/true);
  auto p = project(table, cache);
  ASSERT_NE(p, nullptr);
  expect_refused(p->items.col_refusal);
  EXPECT_FALSE(p->items.row_refusal.has_value());
  EXPECT_EQ(encode(p->items.rows, kRegionProductItems), kOutlineTwoRows);
}

// ---------------------------------------------------------------------------
// Definition source
// ---------------------------------------------------------------------------

TEST(PivotLineItems, DefinitionSourceDefaultsToBuiltAndSurvivesAClone) {
  PivotTable table;
  EXPECT_EQ(table.definition_source(), PivotDefinitionSource::kBuilt);
  table.set_definition_source(PivotDefinitionSource::kXlsb);
  EXPECT_EQ(table.clone()->definition_source(), PivotDefinitionSource::kXlsb);
}

}  // namespace
}  // namespace formulon::pivot
