#include <string>
#include <utility>

#include "gtest/gtest.h"
#include "table.h"
#include "utils/error.h"
#include "workbook.h"

namespace formulon {
namespace {

TableMetadata make_table(std::string name, std::string ref, std::size_t columns, std::size_t sheet_index = 0) {
  TableMetadata table;
  table.sheet_index = sheet_index;
  table.name = name;
  table.display_name = std::move(name);
  table.ref = std::move(ref);
  for (std::size_t i = 0; i < columns; ++i) {
    TableColumn column;
    column.name = "C" + std::to_string(i + 1);
    table.columns.push_back(std::move(column));
  }
  return table;
}

TEST(WorkbookTables, AddTableAssignsIdsAndReturnsIndex) {
  Workbook wb = Workbook::create();
  auto first = wb.add_table(make_table("Sales", "A1:B5", 2));
  ASSERT_TRUE(first.has_value());
  EXPECT_EQ(first.value(), 0U);
  auto second = wb.add_table(make_table("Costs", "D1:D5", 1));
  ASSERT_TRUE(second.has_value());
  EXPECT_EQ(second.value(), 1U);
  ASSERT_EQ(wb.tables().size(), 2U);
  EXPECT_EQ(wb.tables()[0].id, 1U);
  EXPECT_EQ(wb.tables()[1].id, 2U);
  EXPECT_EQ(wb.tables()[0].columns[1].id, 2U);
  EXPECT_FALSE(wb.tables()[0].auto_filter_xml.empty());
}

TEST(WorkbookTables, AddTableRejectsNameClashCaseInsensitively) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(wb.add_table(make_table("Sales", "A1:B5", 2)).has_value());
  auto clash = wb.add_table(make_table("SALES", "D1:E5", 2));
  ASSERT_FALSE(clash.has_value());
  EXPECT_EQ(clash.error().code, FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(clash.error().message, "table name already exists");
  EXPECT_EQ(wb.tables().size(), 1U);
}

TEST(WorkbookTables, AddTableRejectsRangeColumnMismatch) {
  Workbook wb = Workbook::create();
  auto bad = wb.add_table(make_table("Sales", "A1:C5", 2));
  ASSERT_FALSE(bad.has_value());
  EXPECT_EQ(bad.error().code, FormulonErrorCode::kInvalidArgument);
  EXPECT_TRUE(wb.tables().empty());
}

TEST(WorkbookTables, AddTableRejectsBadSheet) {
  Workbook wb = Workbook::create();
  auto bad = wb.add_table(make_table("Sales", "A1:B5", 2, 99));
  ASSERT_FALSE(bad.has_value());
  EXPECT_EQ(bad.error().code, FormulonErrorCode::kInvalidArgument);
}

TEST(WorkbookTables, AddTableRejectsDuplicateColumnNames) {
  Workbook wb = Workbook::create();
  TableMetadata table = make_table("Sales", "A1:B5", 2);
  table.columns[1].name = "c1";
  auto bad = wb.add_table(std::move(table));
  ASSERT_FALSE(bad.has_value());
  EXPECT_EQ(bad.error().message, "column name is duplicated");
}

TEST(WorkbookTables, NextDefaultTableNameFillsGaps) {
  Workbook wb = Workbook::create();
  EXPECT_EQ(wb.next_default_table_name("Table"), "Table1");
  ASSERT_TRUE(wb.add_table(make_table("Table1", "A1:A3", 1)).has_value());
  ASSERT_TRUE(wb.add_table(make_table("Table3", "C1:C3", 1)).has_value());
  EXPECT_EQ(wb.next_default_table_name("Table"), "Table2");
  ASSERT_TRUE(wb.add_table(make_table("Table2", "E1:E3", 1)).has_value());
  EXPECT_EQ(wb.next_default_table_name("Table"), "Table4");
}

TEST(WorkbookTables, NextDefaultTableNameIsCaseInsensitive) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(wb.add_table(make_table("TABLE1", "A1:A3", 1)).has_value());
  EXPECT_EQ(wb.next_default_table_name("table"), "table2");
}

}  // namespace
}  // namespace formulon
