#include <cstddef>
#include <cstdint>
#include <memory>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "cell.h"
#include "eval/eval_profile_scope.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "pivot/field_grouping.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_evaluator.h"
#include "pivot/pivot_layout.h"
#include "pivot/pivot_locale.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "sheet.h"
#include "table.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

using pivot::PivotCache;
using pivot::PivotCacheField;
using pivot::PivotCacheRecord;
using pivot::PivotCellKind;
using pivot::PivotField;
using pivot::PivotTable;

using Grid = std::vector<std::vector<std::string>>;

constexpr std::size_t kPivotSheet = 1U;
constexpr std::uint32_t kRegion = 0U;
constexpr std::uint32_t kProduct = 1U;
constexpr std::uint32_t kAmt = 2U;

const Grid kHeader = {{"Region", "Product", "Amt", "Qty"}};

// The source rows, in source order, with Region and Product as shared items.
PivotCache SourceCache() {
  PivotCache cache;
  cache.set_cache_id(1U);
  cache.mutable_fields().push_back(
      PivotCacheField{"Region", {Value::text("North"), Value::text("South"), Value::text("East")}});
  cache.mutable_fields().push_back(PivotCacheField{"Product", {Value::text("A"), Value::text("B")}});
  cache.mutable_fields().push_back(PivotCacheField{"Amt", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Qty", {}});
  const struct {
    double region;
    double product;
    double amt;
    double qty;
  } rows[] = {{0, 0, 100, 1}, {0, 1, 200, 2}, {1, 0, 300, 3}, {1, 1, 400, 4}, {2, 0, 500, 5}};
  for (const auto& row : rows) {
    PivotCacheRecord record;
    record.cells = {Value::number(row.region), Value::number(row.product), Value::number(row.amt),
                    Value::number(row.qty)};
    record.cell_is_index = {true, true, false, false};
    cache.mutable_records().push_back(std::move(record));
  }
  return cache;
}

// Region and Product items in the label order Excel writes them.
std::unique_ptr<PivotTable> SourcePivot(std::size_t field_count) {
  auto table = std::make_unique<PivotTable>();
  table->set_name("PT1");
  table->set_pivot_cache_id(1U);
  for (std::size_t f = 0; f < field_count; ++f) {
    PivotField field;
    field.axis = pivot::PivotAxis::None;
    field.subtotal_top = false;
    table->mutable_fields().push_back(std::move(field));
  }
  for (const std::uint32_t x : {2U, 0U, 1U}) {
    pivot::PivotItem item;
    item.has_cache_index = true;
    item.cache_index = x;
    table->mutable_fields()[kRegion].items.push_back(item);
  }
  for (const std::uint32_t x : {0U, 1U}) {
    pivot::PivotItem item;
    item.has_cache_index = true;
    item.cache_index = x;
    table->mutable_fields()[kProduct].items.push_back(item);
  }
  table->mutable_fields()[kAmt].axis = pivot::PivotAxis::Value;
  pivot::PivotDataField sum;
  sum.name = "合計 / Amt";
  sum.field_index = kAmt;
  table->mutable_data_fields().push_back(std::move(sum));
  table->set_anchor(2U, 0U, 10U, 3U);
  return table;
}

std::unique_ptr<PivotTable> RowsPivot(std::vector<std::uint32_t> rows, std::vector<std::uint32_t> cols = {}) {
  auto table = SourcePivot(4U);
  for (const std::uint32_t f : rows) {
    table->mutable_fields()[f].axis = pivot::PivotAxis::Row;
  }
  for (const std::uint32_t f : cols) {
    table->mutable_fields()[f].axis = pivot::PivotAxis::Col;
  }
  table->mutable_row_field_order() = std::move(rows);
  table->mutable_col_field_order() = std::move(cols);
  return table;
}

// A ja-JP book: the data sheet `Sheet1`, then `Pivot` holding `table`.
Workbook PivotBook(std::unique_ptr<PivotTable> table, PivotCache cache = SourceCache(),
                   ExcelProfile profile = mac_365_ja_jp_profile()) {
  Workbook wb = Workbook::create_for_profile(profile);
  wb.add_pivot_cache(std::make_unique<PivotCache>(std::move(cache)));
  EXPECT_TRUE(static_cast<bool>(wb.add_sheet_entered("Pivot")));
  wb.sheet(kPivotSheet).add_pivot_table(std::move(table));
  return wb;
}

PivotTable& PivotOf(Workbook& wb, std::size_t sheet = kPivotSheet) {
  return *wb.sheet(sheet).mutable_pivot_tables()[0];
}

struct CellAt {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
};

// The grid cell of `kind` showing `value`. Without a value, a value kind
// still skips its text label cells.
CellAt FindCell(Workbook& wb, PivotCellKind kind, std::optional<double> value = std::nullopt,
                std::size_t sheet = kPivotSheet) {
  const bool numeric = kind == PivotCellKind::Data || kind == PivotCellKind::RowSubtotal ||
                       kind == PivotCellKind::ColSubtotal || kind == PivotCellKind::GrandTotal;
  const PivotTable& table = PivotOf(wb, sheet);
  const pivot::PivotLayoutOptions options = pivot::pivot_layout_options_for(wb.excel_profile());
  auto result = pivot::evaluate(table, *wb.find_pivot_cache(1U), options);
  EXPECT_TRUE(static_cast<bool>(result));
  const eval::ScopedEvalProfile scope(wb.excel_profile());
  auto cells = pivot::layout(table, result.value(), options);
  EXPECT_TRUE(static_cast<bool>(cells));
  for (const pivot::PivotCell& cell : cells.value().cells) {
    if (cell.kind != kind || (numeric && !cell.value.is_number())) {
      continue;
    }
    if (!value.has_value() || cell.value.as_number() == *value) {
      return CellAt{cell.row, cell.col};
    }
  }
  ADD_FAILURE() << "no such pivot cell";
  return {};
}

Grid Rows(const Workbook& wb, std::size_t sheet) {
  Grid out;
  const Sheet& s = wb.sheet(sheet);
  for (std::uint32_t r = 0;; ++r) {
    if (s.cell_at(r, 0U) == nullptr) {
      return out;
    }
    std::vector<std::string> row;
    for (std::uint32_t c = 0; c < 4U; ++c) {
      const Cell* cell = s.cell_at(r, c);
      if (cell == nullptr || cell->cached_value.is_blank()) {
        row.emplace_back();
      } else if (cell->cached_value.is_text()) {
        row.emplace_back(cell->cached_value.as_text());
      } else {
        row.push_back(std::to_string(static_cast<long long>(cell->cached_value.as_number())));
      }
    }
    out.push_back(std::move(row));
  }
}

Grid WithHeader(const Grid& body) {
  Grid out = kHeader;
  out.insert(out.end(), body.begin(), body.end());
  return out;
}

std::size_t ShowDetails(Workbook& wb, CellAt at, std::size_t sheet = kPivotSheet) {
  auto shown = wb.show_pivot_details(sheet, 0U, at.row, at.col);
  EXPECT_TRUE(static_cast<bool>(shown)) << (shown ? "" : shown.error().message);
  return shown ? shown.value() : 0U;
}

TEST(PivotShowDetails, LeafCellMakesAnActiveSheetAndTableBeforeThePivot) {
  Workbook wb = PivotBook(RowsPivot({kRegion, kProduct}));
  const std::size_t index = ShowDetails(wb, FindCell(wb, PivotCellKind::Data, 100.0));

  EXPECT_EQ(index, kPivotSheet);
  ASSERT_EQ(wb.sheet_count(), 3U);
  EXPECT_EQ(wb.sheet(0).name(), "Sheet1");
  EXPECT_EQ(wb.sheet(1).name(), "Sheet2");
  EXPECT_EQ(wb.sheet(2).name(), "Pivot");
  EXPECT_EQ(wb.sheet(2).pivot_tables().size(), 1U);
  EXPECT_EQ(wb.active_sheet(), 1U);
  EXPECT_EQ(Rows(wb, 1U), WithHeader({{"North", "A", "100", "1"}}));

  ASSERT_EQ(wb.tables().size(), 1U);
  const TableMetadata& table = wb.tables()[0];
  EXPECT_EQ(table.name, "テーブル1");
  EXPECT_EQ(table.display_name, "テーブル1");
  EXPECT_EQ(table.ref, "A1:D2");
  EXPECT_EQ(table.sheet_index, 1U);
  EXPECT_TRUE(table.header_row);
  EXPECT_FALSE(table.totals_row);
  EXPECT_EQ(table.totals_row_shown, std::optional<bool>(false));
  ASSERT_EQ(table.columns.size(), 4U);
  EXPECT_EQ(table.columns[0].name, "Region");
  EXPECT_EQ(table.columns[3].name, "Qty");
  EXPECT_TRUE(table.style_info.present);
  EXPECT_EQ(table.style_info.name, "TableStyleMedium9");
  EXPECT_TRUE(table.style_info.show_row_stripes);
  EXPECT_FALSE(table.style_info.show_column_stripes);
  EXPECT_FALSE(table.style_info.show_first_column);
  EXPECT_FALSE(table.style_info.show_last_column);
}

TEST(PivotShowDetails, SubtotalCellListsTheItemsRecords) {
  Workbook wb = PivotBook(RowsPivot({kRegion, kProduct}));
  ShowDetails(wb, FindCell(wb, PivotCellKind::RowSubtotal, 300.0));
  EXPECT_EQ(Rows(wb, 1U), WithHeader({{"North", "A", "100", "1"}, {"North", "B", "200", "2"}}));
  EXPECT_EQ(wb.tables()[0].ref, "A1:D3");
}

TEST(PivotShowDetails, GrandTotalOrdersByTheInnerRowFieldFirst) {
  Workbook wb = PivotBook(RowsPivot({kRegion, kProduct}));
  ShowDetails(wb, FindCell(wb, PivotCellKind::GrandTotal, 1500.0));
  EXPECT_EQ(Rows(wb, 1U), WithHeader({{"East", "A", "500", "5"},
                                      {"North", "A", "100", "1"},
                                      {"South", "A", "300", "3"},
                                      {"North", "B", "200", "2"},
                                      {"South", "B", "400", "4"}}));
}

TEST(PivotShowDetails, ManualItemOrderSetsTheRank) {
  {
    auto table = RowsPivot({kRegion, kProduct});
    PivotField& region = table->mutable_fields()[kRegion];
    region.sort.manual = true;
    region.items[0].cache_index = 0U;  // North, South, East
    region.items[1].cache_index = 1U;
    region.items[2].cache_index = 2U;
    Workbook wb = PivotBook(std::move(table));
    ShowDetails(wb, FindCell(wb, PivotCellKind::GrandTotal, 1500.0));
    EXPECT_EQ(Rows(wb, 1U), WithHeader({{"North", "A", "100", "1"},
                                        {"South", "A", "300", "3"},
                                        {"East", "A", "500", "5"},
                                        {"North", "B", "200", "2"},
                                        {"South", "B", "400", "4"}}));
  }
  {
    auto table = RowsPivot({kRegion, kProduct});
    PivotField& product = table->mutable_fields()[kProduct];
    product.sort.manual = true;
    product.items[0].cache_index = 1U;  // B, A
    product.items[1].cache_index = 0U;
    Workbook wb = PivotBook(std::move(table));
    ShowDetails(wb, FindCell(wb, PivotCellKind::GrandTotal, 1500.0));
    EXPECT_EQ(Rows(wb, 1U), WithHeader({{"North", "B", "200", "2"},
                                        {"South", "B", "400", "4"},
                                        {"East", "A", "500", "5"},
                                        {"North", "A", "100", "1"},
                                        {"South", "A", "300", "3"}}));
  }
}

TEST(PivotShowDetails, DescendingFieldReversesItsRank) {
  auto table = RowsPivot({kRegion});
  table->mutable_fields()[kRegion].sort.ascending = false;
  Workbook wb = PivotBook(std::move(table));
  ShowDetails(wb, FindCell(wb, PivotCellKind::GrandTotal, 1500.0));
  EXPECT_EQ(Rows(wb, 1U), WithHeader({{"South", "A", "300", "3"},
                                      {"South", "B", "400", "4"},
                                      {"North", "A", "100", "1"},
                                      {"North", "B", "200", "2"},
                                      {"East", "A", "500", "5"}}));
}

TEST(PivotShowDetails, ColumnFieldRanksAfterTheRowFields) {
  Workbook wb = PivotBook(RowsPivot({kRegion}, {kProduct}));
  ShowDetails(wb, FindCell(wb, PivotCellKind::GrandTotal, 1500.0));
  EXPECT_EQ(Rows(wb, 1U), WithHeader({{"East", "A", "500", "5"},
                                      {"North", "A", "100", "1"},
                                      {"North", "B", "200", "2"},
                                      {"South", "A", "300", "3"},
                                      {"South", "B", "400", "4"}}));

  // A column item's data cell matches on both axes.
  ShowDetails(wb, FindCell(wb, PivotCellKind::Data, 400.0, 2U), 2U);
  EXPECT_EQ(wb.sheet(2).name(), "Sheet3");
  EXPECT_EQ(Rows(wb, 2U), WithHeader({{"South", "B", "400", "4"}}));
}

TEST(PivotShowDetails, CollapsedItemListsEveryRecordBelowIt) {
  auto table = RowsPivot({kRegion, kProduct});
  table->mutable_fields()[kRegion].items[2].show_detail = false;  // South
  Workbook wb = PivotBook(std::move(table));
  const CellAt south = FindCell(wb, PivotCellKind::RowSubtotal, 700.0);
  ShowDetails(wb, south);
  EXPECT_EQ(Rows(wb, 1U), WithHeader({{"South", "A", "300", "3"}, {"South", "B", "400", "4"}}));
}

TEST(PivotShowDetails, ValueFilterDoesNotNarrowTheRecords) {
  auto table = RowsPivot({kRegion});
  pivot::PivotFilter filter;
  filter.axis = pivot::PivotAxis::Row;
  filter.field_name = "Region";
  filter.type = pivot::FilterType::ValueGreaterThan;
  filter.value = 400.0;
  table->mutable_active_filters().push_back(std::move(filter));
  Workbook wb = PivotBook(std::move(table));
  ShowDetails(wb, FindCell(wb, PivotCellKind::GrandTotal));
  EXPECT_EQ(Rows(wb, 1U), WithHeader({{"East", "A", "500", "5"},
                                      {"North", "A", "100", "1"},
                                      {"North", "B", "200", "2"},
                                      {"South", "A", "300", "3"},
                                      {"South", "B", "400", "4"}}));
}

TEST(PivotShowDetails, HiddenItemNarrowsTheRecords) {
  auto table = RowsPivot({kRegion, kProduct});
  table->mutable_fields()[kRegion].items[2].visible = false;  // South
  Workbook wb = PivotBook(std::move(table));
  ShowDetails(wb, FindCell(wb, PivotCellKind::GrandTotal));
  EXPECT_EQ(Rows(wb, 1U),
            WithHeader({{"East", "A", "500", "5"}, {"North", "A", "100", "1"}, {"North", "B", "200", "2"}}));
}

TEST(PivotShowDetails, GroupFieldCellMatchesByGroupItemAndListsDatabaseFieldsOnly) {
  PivotCache cache = SourceCache();
  pivot::FieldGroup group;
  group.base = kRegion;
  group.kind = pivot::GroupKind::Discrete;
  group.discrete_map = {0U, 0U, 1U};  // North, South -> Group1; East -> East
  group.group_items = {Value::text("Group1"), Value::text("East")};
  PivotCacheField grouped{"Region2", {}};
  grouped.is_database_field = false;
  grouped.field_group = std::move(group);
  cache.mutable_fields().push_back(std::move(grouped));
  pivot::materialize_group_fields(cache, /*date1904=*/false);

  auto table = SourcePivot(5U);
  table->mutable_fields()[4].axis = pivot::PivotAxis::Row;
  table->mutable_row_field_order() = {4U};
  Workbook wb = PivotBook(std::move(table), std::move(cache));

  ShowDetails(wb, FindCell(wb, PivotCellKind::Data, 1000.0));
  EXPECT_EQ(Rows(wb, 1U), WithHeader({{"North", "A", "100", "1"},
                                      {"North", "B", "200", "2"},
                                      {"South", "A", "300", "3"},
                                      {"South", "B", "400", "4"}}));
  EXPECT_EQ(wb.tables()[0].columns.size(), 4U);
  EXPECT_EQ(wb.sheet(1).cell_at(0U, 4U), nullptr);
}

TEST(PivotShowDetails, ColumnsTakeTheCacheFieldNumberFormat) {
  PivotCache cache = SourceCache();
  cache.mutable_fields()[kAmt].num_fmt_id = 14U;
  cache.mutable_fields()[3].num_fmt_id = 0U;
  Workbook wb = PivotBook(RowsPivot({kRegion}), std::move(cache));
  ShowDetails(wb, FindCell(wb, PivotCellKind::Data, 500.0));
  const Sheet& details = wb.sheet(1);
  const Cell* amt = details.cell_at(1U, kAmt);
  ASSERT_NE(amt, nullptr);
  EXPECT_EQ(wb.styles().cell_xfs[amt->xf_index].num_fmt_id, 14U);
  EXPECT_EQ(details.cell_at(0U, kAmt)->xf_index, 0U);
  EXPECT_EQ(details.cell_at(1U, 3U)->xf_index, 0U);
  EXPECT_EQ(details.cell_at(1U, kRegion)->xf_index, 0U);
}

TEST(PivotShowDetails, RejectsCellsThatAreNotValueCells) {
  Workbook wb = PivotBook(RowsPivot({kRegion, kProduct}));
  const CellAt label = FindCell(wb, PivotCellKind::RowLabel);
  const CellAt header = FindCell(wb, PivotCellKind::Header);
  for (const CellAt at : {label, header, CellAt{0U, 0U}, CellAt{500U, 500U}}) {
    auto shown = wb.show_pivot_details(kPivotSheet, 0U, at.row, at.col);
    ASSERT_FALSE(static_cast<bool>(shown));
    EXPECT_EQ(shown.error().code, FormulonErrorCode::kInvalidArgument);
  }
  const CellAt data = FindCell(wb, PivotCellKind::Data, 100.0);
  EXPECT_EQ(wb.show_pivot_details(0U, 0U, data.row, data.col).error().code, FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(wb.show_pivot_details(kPivotSheet, 1U, data.row, data.col).error().code,
            FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(wb.show_pivot_details(9U, 0U, data.row, data.col).error().code, FormulonErrorCode::kInvalidArgument);

  PivotOf(wb).set_enable_drill(false);
  EXPECT_EQ(wb.show_pivot_details(kPivotSheet, 0U, data.row, data.col).error().code,
            FormulonErrorCode::kPreconditionFailed);
  EXPECT_EQ(wb.sheet_count(), 2U);
  EXPECT_TRUE(wb.tables().empty());
}

TEST(PivotShowDetails, NameTakesTheSmallestFreeNumber) {
  Workbook wb = PivotBook(RowsPivot({kRegion}));
  ASSERT_TRUE(static_cast<bool>(wb.add_sheet_entered("Sheet3")));
  const CellAt east = FindCell(wb, PivotCellKind::Data, 500.0);

  ShowDetails(wb, east);
  EXPECT_EQ(wb.sheet(1).name(), "Sheet2");
  EXPECT_EQ(wb.tables()[0].name, "テーブル1");

  ShowDetails(wb, east, 2U);
  EXPECT_EQ(wb.sheet(2).name(), "Sheet4");
  EXPECT_EQ(wb.sheet(3).name(), "Pivot");
  EXPECT_EQ(wb.active_sheet(), 2U);
  EXPECT_EQ(wb.tables()[1].name, "テーブル2");
  EXPECT_EQ(wb.tables()[1].sheet_index, 2U);
  EXPECT_EQ(wb.tables()[0].sheet_index, 1U);
}

TEST(PivotShowDetails, NameComparesSheetsCaseInsensitively) {
  Workbook wb = PivotBook(RowsPivot({kRegion}));
  ASSERT_TRUE(static_cast<bool>(wb.add_sheet_entered("SHEET2")));
  ShowDetails(wb, FindCell(wb, PivotCellKind::Data, 500.0));
  EXPECT_EQ(wb.sheet(1).name(), "Sheet3");
}

TEST(PivotShowDetails, LocaleWordStartsAtOneWhenFree) {
  Workbook wb = PivotBook(RowsPivot({kRegion}), SourceCache(), ExcelProfile{ExcelHost::kMac365, ExcelLocale::kFrFR});
  ASSERT_EQ(wb.sheet(0).name(), "Feuil1");
  ASSERT_TRUE(static_cast<bool>(wb.rename_sheet(0U, "Données")));
  ShowDetails(wb, FindCell(wb, PivotCellKind::Data, 500.0));
  EXPECT_EQ(wb.sheet(1).name(), "Feuil1");
  EXPECT_EQ(wb.tables()[0].name, "Tableau1");
}

}  // namespace
}  // namespace formulon
