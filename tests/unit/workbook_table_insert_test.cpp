// `Workbook::insert_table_rows` / `insert_table_columns`: which cells move, how
// the table grows, calculated-column fill and column naming. Expected names
// are Excel's `ListColumns.Add` results (left name counted on, else the
// profile's column word); cell layouts follow the band-insert rule.

#include <cstddef>
#include <cstdint>
#include <memory>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "cell.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "merge_range.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "table.h"
#include "utils/a1_ref.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

CellAddress Addr(std::string_view a1) {
  CellAddress out{0U, 0U};
  EXPECT_TRUE(a1::parse_a1_ref(a1, &out.row, &out.col)) << a1;
  return out;
}

void SetNumber(Workbook& wb, std::string_view a1, double value) {
  const CellAddress at = Addr(a1);
  ASSERT_TRUE(wb.set_cell_value(0U, at.row, at.col, Value::number(value)).has_value()) << a1;
}

void SetText(Workbook& wb, std::string_view a1, std::string_view text) {
  const CellAddress at = Addr(a1);
  ASSERT_TRUE(wb.set_cell_text(0U, at.row, at.col, std::string(text)).has_value()) << a1;
}

// The value at `a1` as text: "" for a blank cell, the integer for a number.
std::string Show(const Workbook& wb, std::string_view a1, std::size_t sheet = 0U) {
  const CellAddress at = Addr(a1);
  const Value value = wb.sheet(sheet).resolve_cell_value(at.row, at.col);
  if (value.is_blank()) {
    return {};
  }
  if (value.is_number()) {
    return std::to_string(static_cast<long long>(value.as_number()));
  }
  return value.is_text() ? std::string(value.as_text()) : std::string("<other>");
}

std::string FormulaAt(const Workbook& wb, std::string_view a1) {
  const CellAddress at = Addr(a1);
  const Cell* cell = wb.sheet(0U).cell_at(at.row, at.col);
  return cell == nullptr ? std::string() : cell->formula_text;
}

std::string Names(const Workbook& wb, std::size_t table = 0U) {
  std::string out;
  for (const TableColumn& column : wb.tables()[table].columns) {
    out += out.empty() ? "" : ",";
    out += column.name;
  }
  return out;
}

void ExpectCode(const Expected<void, Error>& result, FormulonErrorCode code) {
  ASSERT_FALSE(result.has_value());
  EXPECT_EQ(result.error().code, code);
}

ExcelProfile ProfileOf(const char* id) {
  ExcelProfile profile;
  EXPECT_TRUE(parse_excel_profile_id(id, &profile)) << id;
  return profile;
}

// A table over `ref` named `name`, with `columns` as its column names.
TableMetadata MakeTable(std::string name, std::string ref, const std::vector<std::string>& columns, bool header = true,
                        bool totals = false) {
  TableMetadata table;
  table.sheet_index = 0U;
  table.name = name;
  table.display_name = std::move(name);
  table.ref = std::move(ref);
  table.header_row = header;
  table.totals_row = totals;
  for (const std::string& column_name : columns) {
    TableColumn column;
    column.name = column_name;
    table.columns.push_back(std::move(column));
  }
  return table;
}

// Adds the table and writes its names into the header row when it has one.
void AddTable(Workbook& wb, TableMetadata table) {
  const std::vector<TableColumn> columns = table.columns;
  const bool header = table.header_row;
  const CellAddress first = Addr(table.ref.substr(0U, table.ref.find(':')));
  ASSERT_TRUE(wb.add_table(std::move(table)).has_value());
  for (std::size_t i = 0; header && i < columns.size(); ++i) {
    ASSERT_TRUE(
        wb.set_cell_text(0U, first.row, first.col + static_cast<std::uint32_t>(i), columns[i].name).has_value());
  }
}

// Item, Qty over A1:B3 (header and two data rows) holding 1/2 and 10/20.
Workbook SalesBook() {
  Workbook wb = Workbook::create();
  AddTable(wb, MakeTable("tbl", "A1:B3", {"Item", "Qty"}));
  SetNumber(wb, "A2", 1);
  SetNumber(wb, "A3", 2);
  SetNumber(wb, "B2", 10);
  SetNumber(wb, "B3", 20);
  return wb;
}

// ---------------------------------------------------------------------------
// Rows
// ---------------------------------------------------------------------------

TEST(WorkbookTableInsert, AppendWithoutTotalsGrowsTheTableAndShiftsOnlyItsColumns) {
  Workbook wb = SalesBook();
  SetNumber(wb, "A4", 99);
  SetNumber(wb, "C4", 77);
  ASSERT_TRUE(wb.insert_table_rows(0U, 0U, 2U, 2U).has_value());
  EXPECT_EQ(wb.tables()[0].ref, "A1:B4");
  EXPECT_EQ(Show(wb, "A3"), "2");
  EXPECT_EQ(Show(wb, "A4"), "");
  EXPECT_EQ(Show(wb, "A5"), "99");
  EXPECT_EQ(Show(wb, "C4"), "77");
  EXPECT_EQ(Show(wb, "C5"), "");
}

TEST(WorkbookTableInsert, AppendWithTotalsInsertsAboveTheTotalsRow) {
  Workbook wb = Workbook::create();
  AddTable(wb, MakeTable("tbl", "A1:B4", {"Item", "Qty"}, true, true));
  SetNumber(wb, "A2", 1);
  SetNumber(wb, "A3", 2);
  SetText(wb, "A4", "Total");
  ASSERT_TRUE(wb.insert_table_rows(0U, 0U, 2U, 2U).has_value());
  EXPECT_EQ(wb.tables()[0].ref, "A1:B5");
  EXPECT_TRUE(wb.tables()[0].totals_row);
  EXPECT_EQ(Show(wb, "A3"), "2");
  EXPECT_EQ(Show(wb, "A4"), "");
  EXPECT_EQ(Show(wb, "A5"), "Total");
}

TEST(WorkbookTableInsert, MiddleInsertMovesTheRowsBelowInTheTableColumnsOnly) {
  Workbook wb = SalesBook();
  SetNumber(wb, "C3", 7);
  ASSERT_TRUE(wb.insert_table_rows(0U, 0U, 1U, 1U).has_value());
  EXPECT_EQ(wb.tables()[0].ref, "A1:B4");
  EXPECT_EQ(Show(wb, "A2"), "1");
  EXPECT_EQ(Show(wb, "B2"), "10");
  EXPECT_EQ(Show(wb, "A3"), "");
  EXPECT_EQ(Show(wb, "B3"), "");
  EXPECT_EQ(Show(wb, "A4"), "2");
  EXPECT_EQ(Show(wb, "B4"), "20");
  EXPECT_EQ(Show(wb, "C3"), "7");
}

TEST(WorkbookTableInsert, InsertAtZeroPutsTheNewRowsFirst) {
  Workbook wb = SalesBook();
  ASSERT_TRUE(wb.insert_table_rows(0U, 0U, 0U, 1U).has_value());
  EXPECT_EQ(wb.tables()[0].ref, "A1:B5");
  EXPECT_EQ(Show(wb, "A1"), "Item");
  EXPECT_EQ(Show(wb, "A2"), "");
  EXPECT_EQ(Show(wb, "A3"), "");
  EXPECT_EQ(Show(wb, "A4"), "1");
  EXPECT_EQ(Show(wb, "A5"), "2");
}

TEST(WorkbookTableInsert, CalculatedColumnsFillTheNewRowsAndOtherColumnsStayBlank) {
  for (const std::uint32_t at : {0U, 1U, 2U}) {  // first row, middle, append
    Workbook wb = Workbook::create();
    TableMetadata table = MakeTable("tbl", "A1:C3", {"Item", "Qty", "Double"});
    table.columns[2].calculated_column_formula = "B2*2";
    AddTable(wb, std::move(table));
    SetNumber(wb, "B2", 10);
    SetNumber(wb, "B3", 20);
    ASSERT_TRUE(wb.insert_table_rows(0U, 0U, at, at + 1U).has_value()) << at;
    EXPECT_EQ(wb.tables()[0].ref, "A1:C5") << at;
    for (std::uint32_t row = at + 2U; row <= at + 3U; ++row) {
      const std::string r = std::to_string(row);
      EXPECT_EQ(FormulaAt(wb, "C" + r), "=B" + r + "*2") << at;
      EXPECT_EQ(FormulaAt(wb, "A" + r), "") << at;
      EXPECT_EQ(Show(wb, "B" + r), "") << at;
    }
  }
}

TEST(WorkbookTableInsert, CalculatedColumnsFillAnAppendAboveTheTotalsRow) {
  Workbook wb = Workbook::create();
  TableMetadata table = MakeTable("tbl", "A1:B4", {"Qty", "Double"}, true, true);
  table.columns[1].calculated_column_formula = "A2*2";
  AddTable(wb, std::move(table));
  ASSERT_TRUE(wb.insert_table_rows(0U, 0U, 2U, 2U).has_value());
  EXPECT_EQ(FormulaAt(wb, "B4"), "=A4*2");
  EXPECT_EQ(FormulaAt(wb, "B5"), "");
}

TEST(WorkbookTableInsert, CountAboveOneGrowsByTheSpan) {
  Workbook wb = SalesBook();
  ASSERT_TRUE(wb.insert_table_rows(0U, 0U, 1U, 3U).has_value());
  EXPECT_EQ(wb.tables()[0].ref, "A1:B6");
  EXPECT_EQ(Show(wb, "A2"), "1");
  EXPECT_EQ(Show(wb, "A3"), "");
  EXPECT_EQ(Show(wb, "A5"), "");
  EXPECT_EQ(Show(wb, "A6"), "2");
}

TEST(WorkbookTableInsert, AnotherTableBelowMovesAlong) {
  Workbook wb = SalesBook();
  AddTable(wb, MakeTable("other", "A6:B8", {"Name", "Val"}));
  ASSERT_TRUE(wb.insert_table_rows(0U, 0U, 2U, 3U).has_value());
  EXPECT_EQ(wb.tables()[0].ref, "A1:B5");
  EXPECT_EQ(wb.tables()[1].ref, "A8:B10");
  EXPECT_EQ(Show(wb, "A8"), "Name");
}

TEST(WorkbookTableInsert, HeaderRowOffStartsTheBandAtTheFirstDataRow) {
  Workbook wb = Workbook::create();
  AddTable(wb, MakeTable("tbl", "A1:B2", {"Item", "Qty"}, false));
  SetNumber(wb, "A1", 1);
  SetNumber(wb, "A2", 2);
  ASSERT_TRUE(wb.insert_table_rows(0U, 0U, 0U, 0U).has_value());
  EXPECT_EQ(wb.tables()[0].ref, "A1:B3");
  EXPECT_EQ(Show(wb, "A1"), "");
  EXPECT_EQ(Show(wb, "A2"), "1");
  EXPECT_EQ(Show(wb, "A3"), "2");
  ASSERT_TRUE(wb.insert_table_rows(0U, 0U, 3U, 3U).has_value());
  EXPECT_EQ(wb.tables()[0].ref, "A1:B4");
  EXPECT_EQ(Show(wb, "A3"), "2");
  EXPECT_EQ(Show(wb, "A4"), "");
}

TEST(WorkbookTableInsert, RowInsertRefusesAPivotTableBelowInTheTableColumns) {
  Workbook wb = SalesBook();
  auto pivot = std::make_unique<pivot::PivotTable>();
  pivot->set_name("PivotTable1");
  pivot->set_anchor(5U, 0U, 3U, 2U);
  wb.sheet(0U).mutable_pivot_tables().push_back(std::move(pivot));
  ExpectCode(wb.insert_table_rows(0U, 0U, 2U, 2U), FormulonErrorCode::kPreconditionFailed);
  EXPECT_EQ(wb.tables()[0].ref, "A1:B3");
}

TEST(WorkbookTableInsert, RowInsertRefusesAnArrayBlockCutByTheBand) {
  Workbook wb = SalesBook();
  ASSERT_TRUE(wb.set_array_formula(0U, 3U, 1U, 2U, 2U, "=1").has_value());  // B4:C5
  ExpectCode(wb.insert_table_rows(0U, 0U, 2U, 2U), FormulonErrorCode::kPreconditionFailed);
  EXPECT_EQ(wb.tables()[0].ref, "A1:B3");
}

TEST(WorkbookTableInsert, RowInsertRefusesPushingContentOffTheGrid) {
  Workbook wb = SalesBook();
  SetNumber(wb, "A1048576", 1);
  ExpectCode(wb.insert_table_rows(0U, 0U, 1U, 1U), FormulonErrorCode::kPreconditionFailed);
  ExpectCode(wb.insert_table_rows(0U, 0U, 2U, 2U), FormulonErrorCode::kPreconditionFailed);
  EXPECT_EQ(wb.tables()[0].ref, "A1:B3");
  EXPECT_EQ(Show(wb, "A1048576"), "1");
}

TEST(WorkbookTableInsert, AppendOnATableEndingOnTheLastRowIsRefused) {
  Workbook wb = Workbook::create();
  AddTable(wb, MakeTable("tbl", "A1048575:B1048576", {"Item", "Qty"}));
  ExpectCode(wb.insert_table_rows(0U, 0U, 1U, 1U), FormulonErrorCode::kPreconditionFailed);
  EXPECT_EQ(wb.tables()[0].ref, "A1048575:B1048576");
}

TEST(WorkbookTableInsert, RowArgumentsAreChecked) {
  Workbook wb = SalesBook();
  ExpectCode(wb.insert_table_rows(0U, 0U, 3U, 3U), FormulonErrorCode::kInvalidArgument);           // first > count
  ExpectCode(wb.insert_table_rows(0U, 0U, 1U, 0U), FormulonErrorCode::kInvalidArgument);           // last < first
  ExpectCode(wb.insert_table_rows(0U, 0U, 2U, 0xFFFFFFFFU), FormulonErrorCode::kInvalidArgument);  // past the grid
  ExpectCode(wb.insert_table_rows(0U, 1U, 0U, 0U), FormulonErrorCode::kInvalidArgument);           // no such table
  ExpectCode(wb.insert_table_rows(1U, 0U, 0U, 0U), FormulonErrorCode::kInvalidArgument);           // no such sheet
  EXPECT_EQ(wb.tables()[0].ref, "A1:B3");
  EXPECT_EQ(Show(wb, "A3"), "2");
}

TEST(WorkbookTableInsert, ATableOfAnotherSheetIsNotAddressedThroughThisOne) {
  Workbook wb = SalesBook();
  const std::size_t second = wb.add_sheet("Sheet2");
  ExpectCode(wb.insert_table_rows(second, 0U, 0U, 0U), FormulonErrorCode::kInvalidArgument);
  ExpectCode(wb.insert_table_columns(second, 0U, 0U, 0U), FormulonErrorCode::kInvalidArgument);
}

// ---------------------------------------------------------------------------
// Columns
// ---------------------------------------------------------------------------

TEST(WorkbookTableInsert, ColumnNameCountsTheLeftNameOn) {
  Workbook wb = Workbook::create();
  AddTable(wb, MakeTable("tbl", "A1:C3", {"Item", "Q2", "Qty"}));
  ASSERT_TRUE(wb.insert_table_columns(0U, 0U, 2U, 2U).has_value());
  EXPECT_EQ(Names(wb), "Item,Q2,Q3,Qty");
  EXPECT_EQ(Show(wb, "C1"), "Q3");
  EXPECT_EQ(Show(wb, "D1"), "Qty");
  EXPECT_EQ(wb.tables()[0].ref, "A1:D3");
}

TEST(WorkbookTableInsert, ColumnWithoutTrailingDigitsTakesTheColumnWord) {
  Workbook wb = Workbook::create_for_profile(ProfileOf("mac-365-ja_JP"));
  AddTable(wb, MakeTable("tbl", "A1:B3", {"Item", "Qty"}));
  ASSERT_TRUE(wb.insert_table_columns(0U, 0U, 1U, 1U).has_value());
  EXPECT_EQ(Names(wb), "Item,列1,Qty");
  EXPECT_EQ(Show(wb, "B1"), "列1");
}

TEST(WorkbookTableInsert, ColumnAtZeroHasNoLeftNameAndTakesTheColumnWord) {
  Workbook wb = Workbook::create_for_profile(ProfileOf("mac-365-ja_JP"));
  AddTable(wb, MakeTable("tbl", "A1:B3", {"Item", "Qty"}));
  ASSERT_TRUE(wb.insert_table_columns(0U, 0U, 0U, 0U).has_value());
  EXPECT_EQ(Names(wb), "列1,Item,Qty");
  EXPECT_EQ(Show(wb, "A1"), "列1");
  EXPECT_EQ(Show(wb, "B1"), "Item");
}

TEST(WorkbookTableInsert, ColumnWordFollowsTheProfile) {
  Workbook wb = Workbook::create_for_profile(ProfileOf("mac-365-en_US"));
  AddTable(wb, MakeTable("tbl", "A1:B3", {"Item", "Qty"}));
  ASSERT_TRUE(wb.insert_table_columns(0U, 0U, 2U, 2U).has_value());
  EXPECT_EQ(Names(wb), "Item,Qty,Column1");
}

TEST(WorkbookTableInsert, ColumnNameSkipsAWordNumberAlreadyTaken) {
  Workbook wb = Workbook::create_for_profile(ProfileOf("mac-365-ja_JP"));
  AddTable(wb, MakeTable("tbl", "A1:C3", {"Item", "列1", "Qty"}));
  ASSERT_TRUE(wb.insert_table_columns(0U, 0U, 2U, 2U).has_value());
  EXPECT_EQ(Names(wb), "Item,列1,列2,Qty");
}

TEST(WorkbookTableInsert, ColumnNameAfterATakenCountedNameKeepsTheFullLeftName) {
  Workbook wb = Workbook::create_for_profile(ProfileOf("mac-365-ja_JP"));
  AddTable(wb, MakeTable("tbl", "A1:B3", {"列2", "Qty"}));
  ASSERT_TRUE(wb.insert_table_columns(0U, 0U, 2U, 2U).has_value());
  EXPECT_EQ(Names(wb), "列2,Qty,列1");
  ASSERT_TRUE(wb.insert_table_columns(0U, 0U, 3U, 3U).has_value());
  EXPECT_EQ(Names(wb), "列2,Qty,列1,列12");
  EXPECT_EQ(Show(wb, "D1"), "列12");
}

TEST(WorkbookTableInsert, ColumnNameBesideATakenNeighbourNumberAppendsTwo) {
  Workbook wb = Workbook::create();
  AddTable(wb, MakeTable("tbl", "A1:B3", {"Q3", "Q2"}));
  ASSERT_TRUE(wb.insert_table_columns(0U, 0U, 2U, 2U).has_value());
  EXPECT_EQ(Names(wb), "Q3,Q2,Q22");
  ASSERT_TRUE(wb.insert_table_columns(0U, 0U, 3U, 3U).has_value());
  EXPECT_EQ(Names(wb), "Q3,Q2,Q22,Q23");
}

TEST(WorkbookTableInsert, TrailingZerosKeepTheirWidth) {
  Workbook wb = Workbook::create();
  AddTable(wb, MakeTable("tbl", "A1:B3", {"Q09", "Qty"}));
  ASSERT_TRUE(wb.insert_table_columns(0U, 0U, 1U, 1U).has_value());
  EXPECT_EQ(Names(wb), "Q09,Q10,Qty");
}

TEST(WorkbookTableInsert, SeveralColumnsAreNamedLeftToRight) {
  Workbook wb = Workbook::create_for_profile(ProfileOf("mac-365-ja_JP"));
  AddTable(wb, MakeTable("tbl", "A1:B3", {"Item", "Qty"}));
  ASSERT_TRUE(wb.insert_table_columns(0U, 0U, 1U, 2U).has_value());
  EXPECT_EQ(Names(wb), "Item,列1,列2,Qty");
  EXPECT_EQ(wb.tables()[0].ref, "A1:D3");
  EXPECT_EQ(Show(wb, "B1"), "列1");
  EXPECT_EQ(Show(wb, "C1"), "列2");
  EXPECT_EQ(Show(wb, "D1"), "Qty");
}

TEST(WorkbookTableInsert, MiddleColumnInsertMovesOnlyTheTableRows) {
  Workbook wb = SalesBook();
  SetNumber(wb, "A4", 5);
  SetNumber(wb, "C2", 6);
  ASSERT_TRUE(wb.insert_table_columns(0U, 0U, 1U, 1U).has_value());
  EXPECT_EQ(wb.tables()[0].ref, "A1:C3");
  EXPECT_EQ(Show(wb, "A2"), "1");
  EXPECT_EQ(Show(wb, "B2"), "");
  EXPECT_EQ(Show(wb, "C2"), "10");
  EXPECT_EQ(Show(wb, "D2"), "6");
  EXPECT_EQ(Show(wb, "A4"), "5");
  EXPECT_EQ(Show(wb, "B4"), "");
}

TEST(WorkbookTableInsert, AppendedColumnShiftsTheCellsRightOfTheTableInItsRows) {
  Workbook wb = SalesBook();
  SetNumber(wb, "C2", 6);
  SetNumber(wb, "C4", 8);
  ASSERT_TRUE(wb.insert_table_columns(0U, 0U, 2U, 2U).has_value());
  EXPECT_EQ(wb.tables()[0].ref, "A1:C3");
  EXPECT_EQ(Show(wb, "C2"), "");
  EXPECT_EQ(Show(wb, "D2"), "6");
  EXPECT_EQ(Show(wb, "C4"), "8");
}

TEST(WorkbookTableInsert, ColumnInsertCarriesTheTotalsRowAndLeavesTheNewTotalsCellBlank) {
  Workbook wb = Workbook::create();
  AddTable(wb, MakeTable("tbl", "A1:B4", {"Item", "Qty"}, true, true));
  SetText(wb, "A4", "Total");
  SetNumber(wb, "B4", 30);
  ASSERT_TRUE(wb.insert_table_columns(0U, 0U, 1U, 1U).has_value());
  EXPECT_EQ(wb.tables()[0].ref, "A1:C4");
  EXPECT_EQ(Show(wb, "A4"), "Total");
  EXPECT_EQ(Show(wb, "B4"), "");
  EXPECT_EQ(Show(wb, "C4"), "30");
}

TEST(WorkbookTableInsert, ColumnInsertWithoutAHeaderRowWritesNoHeaderCell) {
  Workbook wb = Workbook::create();
  AddTable(wb, MakeTable("tbl", "A1:B2", {"Item", "Qty"}, false));
  ASSERT_TRUE(wb.insert_table_columns(0U, 0U, 2U, 2U).has_value());
  EXPECT_EQ(wb.tables()[0].ref, "A1:C2");
  EXPECT_EQ(wb.tables()[0].columns.size(), 3U);
  EXPECT_FALSE(wb.tables()[0].columns[2].name.empty());
  EXPECT_EQ(Show(wb, "C1"), "");
}

TEST(WorkbookTableInsert, ColumnInsertRefusesAPivotTableToTheRight) {
  Workbook wb = SalesBook();
  auto pivot = std::make_unique<pivot::PivotTable>();
  pivot->set_name("PivotTable1");
  pivot->set_anchor(0U, 4U, 3U, 2U);
  wb.sheet(0U).mutable_pivot_tables().push_back(std::move(pivot));
  ExpectCode(wb.insert_table_columns(0U, 0U, 1U, 1U), FormulonErrorCode::kPreconditionFailed);
  EXPECT_EQ(wb.tables()[0].ref, "A1:B3");
}

TEST(WorkbookTableInsert, ColumnInsertRefusesPushingContentOffTheGrid) {
  Workbook wb = SalesBook();
  SetNumber(wb, "XFD2", 1);
  ExpectCode(wb.insert_table_columns(0U, 0U, 1U, 1U), FormulonErrorCode::kPreconditionFailed);
  EXPECT_EQ(wb.tables()[0].ref, "A1:B3");
  EXPECT_EQ(Names(wb), "Item,Qty");
}

TEST(WorkbookTableInsert, AppendOnATableEndingOnTheLastColumnIsRefused) {
  Workbook wb = Workbook::create();
  AddTable(wb, MakeTable("tbl", "XFC1:XFD3", {"Item", "Qty"}));
  ExpectCode(wb.insert_table_columns(0U, 0U, 2U, 2U), FormulonErrorCode::kPreconditionFailed);
  EXPECT_EQ(wb.tables()[0].ref, "XFC1:XFD3");
}

TEST(WorkbookTableInsert, ColumnArgumentsAreChecked) {
  Workbook wb = SalesBook();
  ExpectCode(wb.insert_table_columns(0U, 0U, 3U, 3U), FormulonErrorCode::kInvalidArgument);  // first > count
  ExpectCode(wb.insert_table_columns(0U, 0U, 1U, 0U), FormulonErrorCode::kInvalidArgument);  // last < first
  ExpectCode(wb.insert_table_columns(0U, 0U, 2U, 0xFFFFFFFFU), FormulonErrorCode::kInvalidArgument);
  ExpectCode(wb.insert_table_columns(0U, 1U, 0U, 0U), FormulonErrorCode::kInvalidArgument);
  ExpectCode(wb.insert_table_columns(1U, 0U, 0U, 0U), FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(wb.tables()[0].ref, "A1:B3");
  EXPECT_EQ(Names(wb), "Item,Qty");
}

}  // namespace
}  // namespace formulon
