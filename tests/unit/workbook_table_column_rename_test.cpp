// `Workbook::rename_table_column` and the entered-formula bind of structured
// references, checked against Mac Excel 365 16.114 (ja-JP, plus the en-US
// default column word). Expected strings are Excel's stored spellings, never
// engine output.

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "cell.h"
#include "defined_name.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "merge_range.h"
#include "sheet_metadata.h"
#include "table.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

ExcelProfile ProfileOf(const char* id) {
  ExcelProfile profile;
  EXPECT_TRUE(parse_excel_profile_id(id, &profile)) << id;
  return profile;
}

TableMetadata MakeTable(std::string name, std::string ref, const std::vector<std::string>& columns,
                        std::size_t sheet_index = 0) {
  TableMetadata table;
  table.sheet_index = sheet_index;
  table.name = name;
  table.display_name = std::move(name);
  table.ref = std::move(ref);
  for (const std::string& column_name : columns) {
    TableColumn column;
    column.name = column_name;
    table.columns.push_back(std::move(column));
  }
  return table;
}

std::string FormulaAt(const Workbook& wb, std::size_t sheet, std::uint32_t row, std::uint32_t col) {
  const Cell* cell = wb.sheet(sheet).cell_at(row, col);
  return cell == nullptr ? std::string() : cell->formula_text;
}

std::string TextAt(const Workbook& wb, std::size_t sheet, std::uint32_t row, std::uint32_t col) {
  const Value value = wb.sheet(sheet).resolve_cell_value(row, col);
  return value.is_text() ? std::string(value.as_text()) : std::string("<not text>");
}

std::string NameFormula(const Workbook& wb, std::string_view name) {
  for (const DefinedName& entry : wb.defined_names()) {
    if (entry.name == name) {
      return entry.formula;
    }
  }
  return "<missing>";
}

void SetFormula(Workbook& wb, std::size_t sheet, std::uint32_t row, std::uint32_t col, std::string formula) {
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(sheet, row, col, std::move(formula))));
}

std::string Rename(Workbook& wb, std::size_t table, std::size_t column, std::string_view name) {
  Expected<std::string, Error> renamed = wb.rename_table_column(table, column, name);
  EXPECT_TRUE(renamed.has_value());
  return renamed.has_value() ? renamed.value() : std::string("<error>");
}

// tblSales on Sheet1!A1:E5 (header, data rows 2-4, totals row 5), columns
// Qty, Unit Cost, Price, a#b, Calc; tblOther on Sheet2!C1:C3 with its own Qty.
struct SalesBook {
  Workbook wb;
  std::size_t sheet2 = 0;
};

SalesBook MakeSalesBook(const char* profile_id) {
  SalesBook book{Workbook::create_for_profile(ProfileOf(profile_id))};
  Workbook& wb = book.wb;
  book.sheet2 = wb.add_sheet("Sheet2");
  const std::vector<std::string> columns = {"Qty", "Unit Cost", "Price", "a#b", "Calc"};
  TableMetadata sales = MakeTable("tblSales", "A1:E5", columns);
  sales.totals_row = true;
  sales.columns[0].totals_function = "custom";
  sales.columns[0].totals_row_formula = "SUBTOTAL(109,tblSales[Qty])";
  sales.columns[4].calculated_column_formula = "tblSales[[#This Row],[Qty]]*tblSales[[#This Row],[Price]]";
  EXPECT_TRUE(wb.add_table(std::move(sales)).has_value());
  EXPECT_TRUE(wb.add_table(MakeTable("tblOther", "C1:C3", {"Qty"}, book.sheet2)).has_value());
  for (std::uint32_t c = 0; c < columns.size(); ++c) {
    EXPECT_TRUE(static_cast<bool>(wb.set_cell_text(0, 0, c, columns[c])));
  }
  SetFormula(wb, 0, 1, 4, "=tblSales[[#This Row],[Qty]]*tblSales[[#This Row],[Price]]");
  SetFormula(wb, 0, 4, 0, "=SUBTOTAL(109,tblSales[Qty])");
  SetFormula(wb, 0, 1, 7, "=SUM(tblSales[Qty])");
  SetFormula(wb, 0, 2, 7, "=SUM(tblSales[[#This Row],[Qty]:[Price]])");
  SetFormula(wb, 0, 3, 7, "=tblSales[[#This Row],[Unit Cost]]");
  SetFormula(wb, 0, 4, 7, "=tblSales[[#This Row],[a'#b]]");
  SetFormula(wb, book.sheet2, 0, 0, "=SUM(tblSales[Qty])+SUM(tblOther[Qty])");
  EXPECT_TRUE(static_cast<bool>(wb.set_defined_name("NmQ", "tblSales[Qty]")));
  EXPECT_TRUE(static_cast<bool>(wb.set_defined_name("NmTR", "tblSales[[#Totals],[Price]]")));
  return book;
}

TEST(WorkbookTableColumnRename, TypedNameRewritesEverySite) {
  SalesBook book = MakeSalesBook("mac-365-ja_JP");
  Workbook& wb = book.wb;
  EXPECT_EQ(Rename(wb, 0, 0, "Amount"), "Amount");
  EXPECT_EQ(wb.tables()[0].columns[0].name, "Amount");
  EXPECT_EQ(TextAt(wb, 0, 0, 0), "Amount");
  EXPECT_EQ(FormulaAt(wb, 0, 1, 4), "=tblSales[[#This Row],[Amount]]*tblSales[[#This Row],[Price]]");
  EXPECT_EQ(FormulaAt(wb, 0, 4, 0), "=SUBTOTAL(109,tblSales[Amount])");
  EXPECT_EQ(FormulaAt(wb, 0, 1, 7), "=SUM(tblSales[Amount])");
  EXPECT_EQ(FormulaAt(wb, 0, 2, 7), "=SUM(tblSales[[#This Row],[Amount]:[Price]])");
  EXPECT_EQ(FormulaAt(wb, book.sheet2, 0, 0), "=SUM(tblSales[Amount])+SUM(tblOther[Qty])");
  EXPECT_EQ(wb.tables()[0].columns[0].totals_row_formula, "SUBTOTAL(109,tblSales[Amount])");
  EXPECT_EQ(wb.tables()[0].columns[4].calculated_column_formula,
            "tblSales[[#This Row],[Amount]]*tblSales[[#This Row],[Price]]");
  EXPECT_EQ(NameFormula(wb, "NmQ"), "tblSales[Amount]");
  EXPECT_EQ(wb.tables()[1].columns[0].name, "Qty");
}

TEST(WorkbookTableColumnRename, SpaceToHyphenRewritesBracketedColumn) {
  SalesBook book = MakeSalesBook("mac-365-ja_JP");
  Workbook& wb = book.wb;
  EXPECT_EQ(Rename(wb, 0, 1, "Unit-Cost"), "Unit-Cost");
  EXPECT_EQ(FormulaAt(wb, 0, 3, 7), "=tblSales[[#This Row],[Unit-Cost]]");
}

TEST(WorkbookTableColumnRename, ClearedHeaderTakesLowestUnusedJapaneseWord) {
  SalesBook book = MakeSalesBook("mac-365-ja_JP");
  Workbook& wb = book.wb;
  EXPECT_EQ(Rename(wb, 0, 2, ""), "列1");
  EXPECT_EQ(TextAt(wb, 0, 0, 2), "列1");
  EXPECT_EQ(NameFormula(wb, "NmTR"), "tblSales[[#Totals],[列1]]");
  EXPECT_EQ(FormulaAt(wb, 0, 2, 7), "=SUM(tblSales[[#This Row],[Qty]:[列1]])");
  EXPECT_EQ(FormulaAt(wb, 0, 1, 4), "=tblSales[[#This Row],[Qty]]*tblSales[[#This Row],[列1]]");
}

TEST(WorkbookTableColumnRename, ClearedHeaderTakesEnglishColumnWord) {
  SalesBook book = MakeSalesBook("mac-365-en_US");
  Workbook& wb = book.wb;
  EXPECT_EQ(Rename(wb, 0, 2, ""), "Column1");
  EXPECT_EQ(TextAt(wb, 0, 0, 2), "Column1");
  EXPECT_EQ(NameFormula(wb, "NmTR"), "tblSales[[#Totals],[Column1]]");
  EXPECT_EQ(Rename(wb, 0, 3, ""), "Column2");
}

TEST(WorkbookTableColumnRename, DuplicateNameGainsLowestSuffix) {
  SalesBook book = MakeSalesBook("mac-365-ja_JP");
  Workbook& wb = book.wb;
  ASSERT_EQ(Rename(wb, 0, 0, "Amount"), "Amount");
  EXPECT_EQ(Rename(wb, 0, 3, "Amount"), "Amount2");
  EXPECT_EQ(TextAt(wb, 0, 0, 3), "Amount2");
  EXPECT_EQ(FormulaAt(wb, 0, 4, 7), "=tblSales[[#This Row],[Amount2]]");
}

TEST(WorkbookTableColumnRename, NumberTextBecomesColumnName) {
  SalesBook book = MakeSalesBook("mac-365-ja_JP");
  Workbook& wb = book.wb;
  EXPECT_EQ(Rename(wb, 0, 0, "2024"), "2024");
  EXPECT_EQ(TextAt(wb, 0, 0, 0), "2024");
  EXPECT_EQ(FormulaAt(wb, 0, 1, 7), "=SUM(tblSales[2024])");
}

TEST(WorkbookTableColumnRename, EnglishProfileRewritesEverySite) {
  SalesBook book = MakeSalesBook("mac-365-en_US");
  Workbook& wb = book.wb;
  EXPECT_EQ(Rename(wb, 0, 0, "Amount"), "Amount");
  EXPECT_EQ(FormulaAt(wb, 0, 4, 0), "=SUBTOTAL(109,tblSales[Amount])");
  EXPECT_EQ(FormulaAt(wb, book.sheet2, 0, 0), "=SUM(tblSales[Amount])+SUM(tblOther[Qty])");
  EXPECT_EQ(wb.tables()[0].columns[0].totals_row_formula, "SUBTOTAL(109,tblSales[Amount])");
  EXPECT_EQ(NameFormula(wb, "NmQ"), "tblSales[Amount]");
}

// R2: columns Qty, Amount2, 列1, x, y, Calc.
Workbook MakeNumberingBook(const char* profile_id) {
  Workbook wb = Workbook::create_for_profile(ProfileOf(profile_id));
  const std::vector<std::string> columns = {"Qty", "Amount2", "列1", "x", "y", "Calc"};
  EXPECT_TRUE(wb.add_table(MakeTable("tblSales", "A1:F4", columns)).has_value());
  for (std::uint32_t c = 0; c < columns.size(); ++c) {
    EXPECT_TRUE(static_cast<bool>(wb.set_cell_text(0, 0, c, columns[c])));
  }
  SetFormula(wb, 0, 1, 7, "=tblSales[[#This Row],[x]]");
  SetFormula(wb, 0, 2, 7, "=SUM(tblSales[Qty])");
  return wb;
}

TEST(WorkbookTableColumnRename, DuplicateSuffixAppendsToWholeName) {
  Workbook wb = MakeNumberingBook("mac-365-ja_JP");
  EXPECT_EQ(Rename(wb, 0, 3, "Amount2"), "Amount22");
  EXPECT_EQ(FormulaAt(wb, 0, 1, 7), "=tblSales[[#This Row],[Amount22]]");
}

TEST(WorkbookTableColumnRename, ClearedHeaderSkipsUsedDefaultName) {
  Workbook wb = MakeNumberingBook("mac-365-ja_JP");
  EXPECT_EQ(Rename(wb, 0, 4, ""), "列2");
  EXPECT_EQ(TextAt(wb, 0, 0, 4), "列2");
}

TEST(WorkbookTableColumnRename, SurroundingSpacesAreKept) {
  Workbook wb = MakeNumberingBook("mac-365-ja_JP");
  EXPECT_EQ(Rename(wb, 0, 3, "  pad  "), "  pad  ");
  EXPECT_EQ(wb.tables()[0].columns[3].name, "  pad  ");
  EXPECT_EQ(FormulaAt(wb, 0, 1, 7), "=tblSales[[#This Row],[  pad  ]]");
}

TEST(WorkbookTableColumnRename, CaseOnlyChangeRenamesAndRewrites) {
  Workbook wb = MakeNumberingBook("mac-365-ja_JP");
  ASSERT_EQ(Rename(wb, 0, 1, "Amt"), "Amt");
  EXPECT_EQ(Rename(wb, 0, 0, "AMOUNT2"), "AMOUNT2");
  EXPECT_EQ(Rename(wb, 0, 0, "qty"), "qty");
  EXPECT_EQ(wb.tables()[0].columns[0].name, "qty");
  EXPECT_EQ(FormulaAt(wb, 0, 2, 7), "=SUM(tblSales[qty])");
}

TEST(WorkbookTableColumnRename, CaseInsensitiveDuplicateGainsSuffix) {
  Workbook wb = MakeNumberingBook("mac-365-ja_JP");
  EXPECT_EQ(Rename(wb, 0, 0, "AMOUNT2"), "AMOUNT22");
}

TEST(WorkbookTableColumnRename, IdenticalNameRewritesNothingButWritesHeader) {
  Workbook wb = MakeNumberingBook("mac-365-ja_JP");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 0, Value::number(7.0))));
  EXPECT_EQ(Rename(wb, 0, 0, "Qty"), "Qty");
  EXPECT_EQ(TextAt(wb, 0, 0, 0), "Qty");
  EXPECT_EQ(FormulaAt(wb, 0, 2, 7), "=SUM(tblSales[Qty])");
}

TEST(WorkbookTableColumnRename, OnlyTheTargetTablesReferencesChange) {
  SalesBook book = MakeSalesBook("mac-365-ja_JP");
  Workbook& wb = book.wb;
  EXPECT_EQ(Rename(wb, 1, 0, "Units"), "Units");
  EXPECT_EQ(FormulaAt(wb, book.sheet2, 0, 0), "=SUM(tblSales[Qty])+SUM(tblOther[Units])");
  EXPECT_EQ(FormulaAt(wb, 0, 1, 7), "=SUM(tblSales[Qty])");
  EXPECT_EQ(wb.tables()[0].columns[0].name, "Qty");
}

TEST(WorkbookTableColumnRename, OutOfRangeIndexIsInvalidArgument) {
  SalesBook book = MakeSalesBook("mac-365-ja_JP");
  Expected<std::string, Error> bad_table = book.wb.rename_table_column(9, 0, "X");
  ASSERT_FALSE(bad_table.has_value());
  EXPECT_EQ(bad_table.error().code, FormulonErrorCode::kInvalidArgument);
  Expected<std::string, Error> bad_column = book.wb.rename_table_column(0, 5, "X");
  ASSERT_FALSE(bad_column.has_value());
  EXPECT_EQ(bad_column.error().code, FormulonErrorCode::kInvalidArgument);
}

TEST(WorkbookTableColumnRename, InsertedColumnTakesProfileWord) {
  Workbook wb = MakeNumberingBook("mac-365-ja_JP");
  ASSERT_TRUE(static_cast<bool>(wb.insert_cols(0, 1, 1)));
  ASSERT_EQ(wb.tables()[0].columns.size(), 7U);
  EXPECT_EQ(wb.tables()[0].columns[1].name, "列2");
  EXPECT_EQ(TextAt(wb, 0, 0, 1), "列2");
}

// A ja-JP book holding table tblN on B1 with `columns` as its headers and
// three data rows.
Workbook MakeInsertBook(const std::vector<std::string>& columns) {
  Workbook wb = Workbook::create_for_profile(ProfileOf("mac-365-ja_JP"));
  const std::string last = std::string(1, static_cast<char>('A' + columns.size())) + "4";
  EXPECT_TRUE(wb.add_table(MakeTable("tblN", "B1:" + last, columns)).has_value());
  for (std::uint32_t c = 0; c < columns.size(); ++c) {
    EXPECT_TRUE(static_cast<bool>(wb.set_cell_text(0, 0, c + 1U, columns[c])));
  }
  return wb;
}

std::vector<std::string> ColumnNames(const Workbook& wb) {
  std::vector<std::string> names;
  for (const TableColumn& column : wb.tables()[0].columns) {
    names.push_back(column.name);
  }
  return names;
}

TEST(WorkbookTableColumnRename, InsertedColumnCountsOnFromItsLeftNeighbour) {
  struct Case {
    std::vector<std::string> columns;
    std::uint32_t col;
    std::uint32_t count;
    std::vector<std::string> expected;
  };
  const Case cases[] = {
      {{"Item", "Q2", "Qty"}, 3U, 1U, {"Item", "Q2", "Q3", "Qty"}},
      {{"Item", "Qty"}, 2U, 1U, {"Item", "列1", "Qty"}},
      {{"Item", "Q2", "Q3"}, 3U, 1U, {"Item", "Q2", "Q22", "Q3"}},
      {{"Item", "列1", "Qty"}, 3U, 1U, {"Item", "列1", "列2", "Qty"}},
      {{"Item", "Q09", "Qty"}, 3U, 1U, {"Item", "Q09", "Q10", "Qty"}},
      {{"Item", "q3", "Q2", "Qty"}, 4U, 1U, {"Item", "q3", "Q2", "Q22", "Qty"}},
      {{"Item", "Q2", "Qty"}, 3U, 2U, {"Item", "Q2", "Q3", "Q4", "Qty"}},
      {{"Q3", "Q2", "Qty"}, 3U, 2U, {"Q3", "Q2", "Q22", "Q23", "Qty"}},
  };
  for (const Case& c : cases) {
    SCOPED_TRACE(c.columns[1] + " at " + std::to_string(c.col));
    Workbook wb = MakeInsertBook(c.columns);
    ASSERT_TRUE(static_cast<bool>(wb.insert_cols(0, c.col, c.count)));
    EXPECT_EQ(ColumnNames(wb), c.expected);
    for (std::uint32_t i = 0; i < c.count; ++i) {
      EXPECT_EQ(TextAt(wb, 0, 0, c.col + i), c.expected[c.col - 1U + i]);
    }
  }
}

TEST(WorkbookTableColumnRename, SecondInsertAtTheSameColumnTakesTheDuplicateForm) {
  Workbook wb = MakeInsertBook({"Item", "Q2", "Qty"});
  ASSERT_TRUE(static_cast<bool>(wb.insert_cols(0, 3, 1)));
  ASSERT_TRUE(static_cast<bool>(wb.insert_cols(0, 3, 1)));
  EXPECT_EQ(ColumnNames(wb), (std::vector<std::string>{"Item", "Q2", "Q22", "Q3", "Qty"}));
}

TEST(WorkbookTableColumnRename, InsertCellsRightInsideATableNamesTheColumnFromItsLeft) {
  Workbook wb = MakeInsertBook({"Item", "Q2", "Qty"});
  ASSERT_TRUE(static_cast<bool>(wb.insert_cells(0, MergeRange{0U, 3U, 3U, 3U}, CellShift::kRight)));
  EXPECT_EQ(wb.tables()[0].ref, "B1:E4");
  EXPECT_EQ(ColumnNames(wb), (std::vector<std::string>{"Item", "Q2", "Q3", "Qty"}));
  EXPECT_EQ(TextAt(wb, 0, 0, 3), "Q3");
  EXPECT_EQ(TextAt(wb, 0, 0, 4), "Qty");
}

TEST(WorkbookTableColumnRename, IngestBindsThisRowAndBareTableName) {
  SalesBook book = MakeSalesBook("mac-365-ja_JP");
  std::string this_row = "=tblSales[@Qty]";
  book.wb.ingest_entered_formula(this_row);
  EXPECT_EQ(this_row, "=tblSales[[#This Row],[Qty]]");
  std::string bare = "=ROWS(tblSales)";
  book.wb.ingest_entered_formula(bare);
  EXPECT_EQ(bare, "=ROWS(tblSales[])");
}

// A defined name spelled like a table: Excel refuses to create one, so only a
// loaded file carries it.
void LoadNameLikeTable(Workbook& wb, std::int32_t local_sheet_id) {
  DefinedName name;
  name.name = "tblSales";
  name.formula = "Sheet1!$A$1";
  name.local_sheet_id = local_sheet_id;
  wb.set_defined_names({name});
}

TEST(WorkbookTableColumnRename, IngestKeepsDefinedNameNamedLikeTable) {
  SalesBook book = MakeSalesBook("mac-365-ja_JP");
  LoadNameLikeTable(book.wb, -1);
  std::string text = "=ROWS(tblSales)";
  book.wb.ingest_entered_formula(text);
  EXPECT_EQ(text, "=ROWS(tblSales)");
}

TEST(WorkbookTableColumnRename, IngestIgnoresSheetLocalNames) {
  SalesBook book = MakeSalesBook("mac-365-ja_JP");
  LoadNameLikeTable(book.wb, 0);
  std::string text = "=ROWS(tblSales)";
  book.wb.ingest_entered_formula(text);
  EXPECT_EQ(text, "=ROWS(tblSales[])");
}

TEST(WorkbookTableColumnRename, EnteredBindInsideTableQualifiesBareReference) {
  SalesBook book = MakeSalesBook("mac-365-ja_JP");
  EXPECT_EQ(book.wb.bind_entered_structured_refs("=[@Qty]", 0, 1, 4), "=tblSales[[#This Row],[Qty]]");
  EXPECT_EQ(book.wb.bind_entered_structured_refs("=SUM([Qty])", 0, 2, 4), "=SUM(tblSales[Qty])");
  EXPECT_EQ(book.wb.bind_entered_structured_refs("=tblSales[@Qty]", 0, 1, 7), "=tblSales[[#This Row],[Qty]]");
}

TEST(WorkbookTableColumnRename, EnteredBindHonoursSheetLocalNames) {
  SalesBook book = MakeSalesBook("mac-365-ja_JP");
  LoadNameLikeTable(book.wb, 0);
  EXPECT_EQ(book.wb.bind_entered_structured_refs("=ROWS(tblSales)", 0, 1, 7), "=ROWS(tblSales)");
  EXPECT_EQ(book.wb.bind_entered_structured_refs("=ROWS(tblSales)", book.sheet2, 5, 5), "=ROWS(tblSales[])");
}

}  // namespace
}  // namespace formulon
