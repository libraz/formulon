#include <cstdint>
#include <memory>
#include <string>
#include <utility>
#include <vector>

#include "data_table.h"
#include "defined_name.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "merge_range.h"
#include "pivot/pivot_table.h"
#include "table.h"
#include "utils/error.h"
#include "value.h"
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

TEST(WorkbookTables, AddTableRejectsADisplayNameExcelRefuses) {
  Workbook wb = Workbook::create();
  for (const char* name : {"T1", "R1C1", "R2D2", "TRUE", "a b", "1a"}) {
    auto bad = wb.add_table(make_table(name, "A1:B5", 2));
    ASSERT_FALSE(bad.has_value()) << name;
    EXPECT_EQ(bad.error().code, FormulonErrorCode::kInvalidArgument);
    EXPECT_EQ(bad.error().message, "invalid table name");
  }
  // Excel keeps an off-grid `name` and repairs only the display name.
  TableMetadata table = make_table("T1", "A1:B5", 2);
  table.display_name = "T1_";
  EXPECT_TRUE(wb.add_table(std::move(table)).has_value());
}

TEST(WorkbookTables, AddTableRejectsTheNameOfADefinedName) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(wb.set_defined_name_scoped("Local", "Sheet1!$D$1", 0).has_value());
  ASSERT_TRUE(wb.set_defined_name("Global", "Sheet1!$D$1").has_value());
  EXPECT_FALSE(wb.add_table(make_table("local", "A1:B5", 2)).has_value());
  EXPECT_FALSE(wb.add_table(make_table("GLOBAL", "A1:B5", 2)).has_value());
  EXPECT_EQ(wb.next_default_table_name("Glob"), "Glob1");
  ASSERT_TRUE(wb.set_defined_name("Table1", "Sheet1!$D$1").has_value());
  EXPECT_EQ(wb.next_default_table_name("Table"), "Table2");
}

TEST(WorkbookTables, AddTableRejectsAnOverlappingRange) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(wb.add_table(make_table("One", "A1:B4", 2)).has_value());
  auto overlap = wb.add_table(make_table("Two", "B3:C6", 2));
  ASSERT_FALSE(overlap.has_value());
  EXPECT_EQ(overlap.error().message, "table range overlaps another table");
  EXPECT_TRUE(wb.add_table(make_table("Two", "C1:D4", 2)).has_value());
  wb.add_sheet("Sheet2");
  EXPECT_TRUE(wb.add_table(make_table("Three", "A1:B4", 2, 1)).has_value());
}

TEST(WorkbookTables, AddTableRefusesWhatExcelKeepsATableOffAndUnmergesTheRest) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(wb.set_array_formula(0, 9U, 1U, 1U, 2U, "=1").has_value());   // B10:C10
  ASSERT_TRUE(wb.set_array_formula(0, 12U, 1U, 1U, 1U, "=1").has_value());  // B13
  ASSERT_TRUE(wb.add_merge(0, MergeRange{20U, 1U, 20U, 2U}).has_value());   // B21:C21
  ASSERT_TRUE(wb.add_merge(0, MergeRange{30U, 0U, 30U, 1U}).has_value());   // A31:B31
  DataTableInput input;
  input.present = true;
  ASSERT_TRUE(wb.add_data_table(0, MergeRange{41U, 1U, 43U, 1U}, DataTableInput{}, input).has_value());  // B42:B44
  auto pivot = std::make_unique<pivot::PivotTable>();
  pivot->set_anchor(/*row=*/50U, /*col=*/0U, /*rows=*/3U, /*cols=*/2U);  // A51:B53
  wb.sheet(0).add_pivot_table(std::move(pivot));

  const auto refused = [&](const char* ref) {
    auto r = wb.add_table(make_table("T_x", ref, 2));
    return r.has_value() ? std::string("accepted") : r.error().message;
  };
  EXPECT_EQ(refused("A9:B11"), "multi-cell array formulas are not valid in a table");
  EXPECT_EQ(refused("A20:B22"), "table range cuts through merged cells");
  EXPECT_EQ(refused("B43:C44"), "table range overlaps a data table");
  EXPECT_EQ(refused("B52:C53"), "table range overlaps a PivotTable");
  EXPECT_EQ(wb.sheet(0).merges().size(), 2U);

  EXPECT_TRUE(wb.add_table(make_table("Single", "A12:B14", 2)).has_value());
  EXPECT_TRUE(wb.add_table(make_table("Merged", "A30:B32", 2)).has_value());
  ASSERT_EQ(wb.sheet(0).merges().size(), 1U);
  EXPECT_EQ(wb.sheet(0).merges()[0].first_row, 20U);
}

TEST(WorkbookTables, RenameTableChecksTheNewNames) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(wb.add_table(make_table("One", "A1:B4", 2)).has_value());
  ASSERT_TRUE(wb.add_table(make_table("Two", "D1:E4", 2)).has_value());
  ASSERT_TRUE(wb.set_defined_name("Taken", "Sheet1!$D$1").has_value());
  EXPECT_FALSE(wb.rename_table(0, "A1", "A1").has_value());
  EXPECT_FALSE(wb.rename_table(0, "two", "two").has_value());
  EXPECT_FALSE(wb.rename_table(0, "Fresh", "taken").has_value());
  EXPECT_FALSE(wb.rename_table(0, "", "Fresh").has_value());
  EXPECT_TRUE(wb.rename_table(0, "ONE", "ONE").has_value());
  EXPECT_EQ(wb.tables()[0].display_name, "ONE");
}

TEST(WorkbookTables, NewDefinedNameFollowsNameSyntaxAndAvoidsTables) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(wb.add_table(make_table("Sales", "A1:B4", 2)).has_value());
  for (const char* name : {"A1", "R1C1", "\\a", "sales"}) {
    auto bad = wb.set_defined_name(name, "Sheet1!$D$1");
    ASSERT_FALSE(bad.has_value()) << name;
    EXPECT_EQ(bad.error().code, FormulonErrorCode::kInvalidArgument);
  }
  EXPECT_TRUE(wb.defined_names().empty());
  // A loaded entry Excel would refuse to create still updates and removes.
  wb.set_defined_names({DefinedName("A1", "Sheet1!$D$1", -1, false, "")});
  EXPECT_TRUE(wb.set_defined_name("A1", "Sheet1!$D$2").has_value());
  EXPECT_EQ(wb.defined_names()[0].formula, "Sheet1!$D$2");
  EXPECT_TRUE(wb.set_defined_name("A1", "").has_value());
  EXPECT_TRUE(wb.defined_names().empty());
}

void Recalc(Workbook& wb) {
  ASSERT_TRUE(wb.recalc(eval::default_registry()));
}

Value At(const Workbook& wb, std::uint32_t row, std::uint32_t col) {
  return wb.sheet(0).resolve_cell_value(row, col);
}

// Sales on A1:B5: header row, data rows 2-4 (B = 1, 2, 3), totals row 5.
void FillSales(Workbook& wb) {
  for (std::uint32_t r = 1; r <= 3; ++r) {
    ASSERT_TRUE(wb.set_cell_value(0, r, 1, Value::number(static_cast<double>(r))));
  }
}

TableMetadata sales_with_totals() {
  TableMetadata table = make_table("Sales", "A1:B5", 2);
  table.id = 1;
  table.columns[0].id = 1;
  table.columns[1].id = 2;
  table.totals_row = true;
  return table;
}

TEST(WorkbookTables, SetTablesRegistersFormulasReadBeforeTheirTable) {
  // A reader stores cell formulas before the table list, as a file load does.
  Workbook wb = Workbook::create();
  FillSales(wb);
  ASSERT_TRUE(wb.set_cell_formula(0, 4, 1, "=SUBTOTAL(109,Sales[C2])"));
  ASSERT_TRUE(wb.set_cell_formula(0, 0, 3, "=SUM(Sales[C2])"));
  ASSERT_TRUE(wb.set_cell_formula(0, 1, 3, "=Sales[[#Totals],[C2]]"));
  ASSERT_TRUE(wb.set_cell_formula(0, 2, 3, "=Sales[[#Headers],[C1]]"));
  ASSERT_TRUE(wb.set_cell_value(0, 0, 0, Value::text("C1")));
  std::vector<TableMetadata> tables;
  tables.push_back(sales_with_totals());
  wb.set_tables(std::move(tables));
  Recalc(wb);
  EXPECT_EQ(At(wb, 4, 1), Value::number(6.0));
  EXPECT_EQ(At(wb, 0, 3), Value::number(6.0));
  EXPECT_EQ(At(wb, 1, 3), Value::number(6.0));

  ASSERT_TRUE(wb.set_cell_value(0, 2, 1, Value::number(100.0)));
  ASSERT_TRUE(wb.set_cell_value(0, 0, 0, Value::text("Region")));
  Recalc(wb);
  EXPECT_EQ(At(wb, 4, 1), Value::number(104.0));
  EXPECT_EQ(At(wb, 0, 3), Value::number(104.0));
  EXPECT_EQ(At(wb, 1, 3), Value::number(104.0));
  EXPECT_EQ(At(wb, 2, 3), Value::text("Region"));
}

TEST(WorkbookTables, ThisRowReferenceFollowsItsOwnRow) {
  Workbook wb = Workbook::create();
  FillSales(wb);
  ASSERT_TRUE(wb.add_table(make_table("Sales", "A1:B4", 2)).has_value());
  ASSERT_TRUE(wb.set_cell_formula(0, 1, 2, "=Sales[[#This Row],[C2]]*10"));
  ASSERT_TRUE(wb.set_cell_formula(0, 2, 2, "=Sales[@C2]*10"));
  Recalc(wb);
  EXPECT_EQ(At(wb, 1, 2), Value::number(10.0));
  EXPECT_EQ(At(wb, 2, 2), Value::number(20.0));

  ASSERT_TRUE(wb.set_cell_value(0, 1, 1, Value::number(7.0)));
  ASSERT_TRUE(wb.set_cell_value(0, 2, 1, Value::number(8.0)));
  Recalc(wb);
  EXPECT_EQ(At(wb, 1, 2), Value::number(70.0));
  EXPECT_EQ(At(wb, 2, 2), Value::number(80.0));
}

}  // namespace
}  // namespace formulon
