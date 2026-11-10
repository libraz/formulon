#include <cstdint>
#include <memory>
#include <string>
#include <utility>
#include <vector>

#include "data_table.h"
#include "defined_name.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "excel_profile.h"
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
  EXPECT_EQ(clash.error().code, FormulonErrorCode::kAlreadyExists);
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

TableMetadata table_with_columns(std::string ref, const std::vector<std::string>& names) {
  TableMetadata table = make_table("Sales", std::move(ref), names.size());
  for (std::size_t i = 0; i < names.size(); ++i) {
    table.columns[i].name = names[i];
  }
  return table;
}

std::string text_at(const Workbook& wb, std::uint32_t row, std::uint32_t col) {
  const Value value = wb.sheet(0).resolve_cell_value(row, col);
  return value.is_text() ? std::string(value.as_text()) : std::string();
}

std::vector<std::string> column_names(const TableMetadata& table) {
  std::vector<std::string> names;
  for (const TableColumn& column : table.columns) {
    names.push_back(column.name);
  }
  return names;
}

TEST(WorkbookTables, AddTableNumbersDuplicateColumnsPastNamesLaterInTheRow) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(wb.add_table(table_with_columns("A1:D5", {"Qty", "qty", "QTY", "Qty2"})).has_value());
  EXPECT_EQ(column_names(wb.tables()[0]), (std::vector<std::string>{"Qty", "qty3", "QTY4", "Qty2"}));
  EXPECT_EQ(wb.tables()[0].columns[3].id, 4U);
}

TEST(WorkbookTables, AddTableNumbersRepeatedColumnsFromTwo) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(wb.add_table(table_with_columns("A1:C5", {"Qty", "Qty", "Qty"})).has_value());
  EXPECT_EQ(column_names(wb.tables()[0]), (std::vector<std::string>{"Qty", "Qty2", "Qty3"}));
}

TEST(WorkbookTables, AddTableFoldsWidthKanaAndDakutenAndKeepsTheRepeatSpelling) {
  const std::vector<std::pair<std::vector<std::string>, std::string>> cases = {
      {{"QTY", "\xEF\xBC\xB1\xEF\xBC\xB4\xEF\xBC\xB9"},
       "\xEF\xBC\xB1\xEF\xBC\xB4\xEF\xBC\xB9"
       "2"},  // ＱＴＹ
      {{"\xE3\x81\x8B\xE3\x81\xAA", "\xE3\x82\xAB\xE3\x83\x8A"},
       "\xE3\x82\xAB\xE3\x83\x8A"
       "2"},  // かな / カナ
      {{"\xE3\x82\xAC", "\xEF\xBD\xB6\xEF\xBE\x9E"},
       "\xEF\xBD\xB6\xEF\xBE\x9E"
       "2"},  // ガ / ｶﾞ
  };
  for (const auto& [given, second] : cases) {
    SCOPED_TRACE(given[0] + " / " + given[1]);
    Workbook wb = Workbook::create();
    ASSERT_TRUE(wb.add_table(table_with_columns("A1:B5", given)).has_value());
    EXPECT_EQ(wb.tables()[0].columns[0].name, given[0]);
    EXPECT_EQ(wb.tables()[0].columns[1].name, second);
  }
}

TEST(WorkbookTables, AddTableNamesBlankColumnsWithTheProfileColumnWord) {
  ExcelProfile profile;
  ASSERT_TRUE(parse_excel_profile_id("mac-365-ja_JP", &profile));
  Workbook wb = Workbook::create_for_profile(profile);
  ASSERT_TRUE(wb.add_table(table_with_columns("A1:C5", {"", "列1", ""})).has_value());
  EXPECT_EQ(column_names(wb.tables()[0]), (std::vector<std::string>{"列2", "列1", "列3"}));
}

TEST(WorkbookTables, AddTableWritesRenamedColumnsIntoTheHeaderRow) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(wb.set_cell_text(0, 0, 0, "Qty").has_value());
  ASSERT_TRUE(wb.set_cell_text(0, 0, 1, "qty").has_value());
  ASSERT_TRUE(wb.add_table(table_with_columns("A1:C5", {"Qty", "qty", ""})).has_value());
  EXPECT_EQ(text_at(wb, 0, 0), "Qty");
  EXPECT_EQ(text_at(wb, 0, 1), "qty2");
  EXPECT_EQ(text_at(wb, 0, 2), "Column1");
}

TEST(WorkbookTables, AddTableWithoutHeaderRowWritesNoHeaderCell) {
  Workbook wb = Workbook::create();
  TableMetadata table = table_with_columns("A1:B5", {"", ""});
  table.header_row = false;
  ASSERT_TRUE(wb.add_table(std::move(table)).has_value());
  EXPECT_EQ(column_names(wb.tables()[0]), (std::vector<std::string>{"Column1", "Column2"}));
  EXPECT_TRUE(text_at(wb, 0, 0).empty());
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
    EXPECT_EQ(bad.error().message, "name is not a valid table name");
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
  EXPECT_EQ(overlap.error().code, FormulonErrorCode::kPreconditionFailed);
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
  EXPECT_EQ(refused("A9:B11"), "range must not hold a multi-cell array formula");
  EXPECT_EQ(refused("A20:B22"), "table range cuts through merged cells");
  EXPECT_EQ(refused("B43:C44"), "table range overlaps a data table");
  EXPECT_EQ(refused("B52:C53"), "table range overlaps a PivotTable");
  for (const char* ref : {"A9:B11", "A20:B22", "B43:C44", "B52:C53"}) {
    auto r = wb.add_table(make_table("T_x", ref, 2));
    ASSERT_FALSE(r.has_value()) << ref;
    EXPECT_EQ(r.error().code, FormulonErrorCode::kPreconditionFailed) << ref;
  }
  EXPECT_EQ(wb.sheet(0).merges().size(), 2U);

  EXPECT_TRUE(wb.add_table(make_table("Single", "A12:B14", 2)).has_value());
  EXPECT_TRUE(wb.add_table(make_table("Merged", "A30:B32", 2)).has_value());
  ASSERT_EQ(wb.sheet(0).merges().size(), 1U);
  EXPECT_EQ(wb.sheet(0).merges()[0].first_row, 20U);
}

TEST(WorkbookTables, ResizeChecksWhatExcelsResizeRefuses) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(wb.add_table(make_table("One", "A1:B4", 2)).has_value());
  ASSERT_TRUE(wb.add_table(make_table("Two", "D1:E4", 2)).has_value());
  ASSERT_TRUE(wb.add_merge(0, MergeRange{5U, 0U, 5U, 1U}).has_value());    // A6:B6
  ASSERT_TRUE(wb.set_array_formula(0, 9U, 0U, 1U, 1U, "=1").has_value());  // A10
  const auto resized = [&](std::uint32_t first_row, std::uint32_t last_row, std::uint32_t last_col) {
    auto r = wb.check_table_resize(0, MergeRange{first_row, 0U, last_row, last_col});
    return r.has_value() ? std::string("accepted") : r.error().message;
  };
  EXPECT_EQ(resized(0U, 3U, 1U), "accepted");
  EXPECT_EQ(resized(1U, 4U, 1U), "table range must keep its first row");
  EXPECT_EQ(resized(0U, 0U, 1U), "table range has no data row");
  EXPECT_EQ(resized(0U, 3U, 3U), "table range overlaps another table");
  EXPECT_EQ(resized(0U, 6U, 1U), "table range overlaps merged cells");
  EXPECT_EQ(wb.check_table_resize(0, MergeRange{0U, 0U, 3U, 3U}).error().code, FormulonErrorCode::kPreconditionFailed);
  EXPECT_EQ(wb.check_table_resize(0, MergeRange{1U, 0U, 4U, 1U}).error().code, FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(resized(0U, 4U, 1U), "accepted");
  ASSERT_TRUE(wb.remove_merges_intersecting(0, MergeRange{5U, 0U, 5U, 1U}).has_value());
  EXPECT_EQ(resized(0U, 11U, 1U), "accepted");
  EXPECT_FALSE(wb.check_table_resize(2, MergeRange{0U, 0U, 3U, 1U}).has_value());
}

TEST(WorkbookTables, RenameTableChecksTheNewNames) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(wb.add_table(make_table("One", "A1:B4", 2)).has_value());
  ASSERT_TRUE(wb.add_table(make_table("Two", "D1:E4", 2)).has_value());
  ASSERT_TRUE(wb.set_defined_name("Taken", "Sheet1!$D$1").has_value());
  EXPECT_FALSE(wb.rename_table(0, "A1", "A1").has_value());
  EXPECT_EQ(wb.rename_table(0, "two", "two").error().code, FormulonErrorCode::kAlreadyExists);
  EXPECT_EQ(wb.rename_table(0, "Fresh", "taken").error().code, FormulonErrorCode::kAlreadyExists);
  EXPECT_FALSE(wb.rename_table(0, "", "Fresh").has_value());
  EXPECT_TRUE(wb.rename_table(0, "ONE", "ONE").has_value());
  EXPECT_EQ(wb.tables()[0].display_name, "ONE");
}

TEST(WorkbookTables, NewDefinedNameFollowsNameSyntaxAndAvoidsTables) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(wb.add_table(make_table("Sales", "A1:B4", 2)).has_value());
  for (const char* name : {"A1", "R1C1", "\\a"}) {
    auto bad = wb.set_defined_name(name, "Sheet1!$D$1");
    ASSERT_FALSE(bad.has_value()) << name;
    EXPECT_EQ(bad.error().code, FormulonErrorCode::kInvalidArgument);
  }
  auto taken = wb.set_defined_name("sales", "Sheet1!$D$1");
  ASSERT_FALSE(taken.has_value());
  EXPECT_EQ(taken.error().code, FormulonErrorCode::kAlreadyExists);
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

// Excel refuses a spill of more than one cell that touches a table, anchor
// included, and lets it through again once the table is gone.
TEST(WorkbookTables, ASpillTouchingATableIsRefused) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(wb.set_cell_formula(0, 9U, 0U, "=SEQUENCE(3)").has_value());     // A10:A12
  ASSERT_TRUE(wb.set_cell_formula(0, 1U, 0U, "=SEQUENCE(1,4)").has_value());   // A2:D2
  ASSERT_TRUE(wb.set_cell_formula(0, 1U, 6U, "=SEQUENCE(1)").has_value());     // G2
  ASSERT_TRUE(wb.set_cell_formula(0, 2U, 6U, "=SEQUENCE(2)").has_value());     // G3:G4
  ASSERT_TRUE(wb.set_cell_formula(0, 20U, 0U, "=SEQUENCE(1,2)").has_value());  // A21:B21
  Recalc(wb);
  EXPECT_EQ(At(wb, 10U, 0U), Value::number(2.0));

  ASSERT_TRUE(wb.add_table(make_table("Below", "A11:B13", 2)).has_value());
  ASSERT_TRUE(wb.add_table(make_table("Right", "C1:D4", 2)).has_value());
  ASSERT_TRUE(wb.add_table(make_table("Inner", "G1:G5", 1)).has_value());
  ASSERT_TRUE(wb.add_table(make_table("Next", "A23:B24", 2)).has_value());
  Recalc(wb);
  EXPECT_EQ(At(wb, 9U, 0U), Value::error(ErrorCode::Spill));
  EXPECT_EQ(At(wb, 1U, 0U), Value::error(ErrorCode::Spill));
  EXPECT_EQ(At(wb, 1U, 6U), Value::number(1.0));
  EXPECT_EQ(At(wb, 2U, 6U), Value::error(ErrorCode::Spill));
  EXPECT_EQ(At(wb, 20U, 0U), Value::number(1.0));

  // Rows inserted above both move them together: the table's old rectangle
  // must not stay behind on the spill's new rows.
  ASSERT_TRUE(wb.insert_rows(0, 14U, 3U).has_value());
  Recalc(wb);
  EXPECT_EQ(At(wb, 23U, 0U), Value::number(1.0));
  EXPECT_EQ(At(wb, 23U, 1U), Value::number(2.0));

  std::vector<TableMetadata> kept(wb.tables().begin() + 1, wb.tables().end());
  wb.set_tables(std::move(kept));
  Recalc(wb);
  EXPECT_EQ(At(wb, 9U, 0U), Value::number(1.0));
  EXPECT_EQ(At(wb, 11U, 0U), Value::number(3.0));
  EXPECT_EQ(At(wb, 1U, 0U), Value::error(ErrorCode::Spill));
}

// East / West on A1:B3 under a header row; Amount totals with SUM, Region with a label.
void AddRegions(Workbook& wb) {
  TableMetadata table = make_table("Regions", "A1:B3", 2);
  table.columns[0].name = "Region";
  table.columns[0].totals_label = "Total";
  table.columns[1].name = "Amount";
  table.columns[1].totals_function = "sum";
  ASSERT_TRUE(wb.set_cell_text(0, 1U, 0U, "East").has_value());
  ASSERT_TRUE(wb.set_cell_value(0, 1U, 1U, Value::number(10.0)).has_value());
  ASSERT_TRUE(wb.set_cell_text(0, 2U, 0U, "West").has_value());
  ASSERT_TRUE(wb.set_cell_value(0, 2U, 1U, Value::number(30.0)).has_value());
  ASSERT_TRUE(wb.add_table(std::move(table)).has_value());
}

void ExpectTotalsRowRefused(Workbook& wb, const std::string& ref) {
  const auto shown = wb.set_table_totals_row(0, true);
  ASSERT_FALSE(shown.has_value());
  EXPECT_EQ(shown.error().code, FormulonErrorCode::kPreconditionFailed);
  EXPECT_EQ(wb.tables()[0].ref, ref);
  EXPECT_FALSE(wb.tables()[0].totals_row);
}

TEST(WorkbookTableTotalsRow, ShowingGrowsTheTableByTheRowBelowAndKeepsTheData) {
  Workbook wb = Workbook::create();
  AddRegions(wb);
  ASSERT_TRUE(wb.set_table_totals_row(0, true).has_value());
  EXPECT_EQ(wb.tables()[0].ref, "A1:B4");
  EXPECT_TRUE(wb.tables()[0].totals_row);
  Recalc(wb);
  EXPECT_EQ(At(wb, 2U, 1U), Value::number(30.0));
  EXPECT_EQ(At(wb, 3U, 1U), Value::number(40.0));
  EXPECT_EQ(At(wb, 3U, 0U).as_text(), "Total");
  ASSERT_NE(wb.sheet(0).cell_at(3U, 1U), nullptr);
  EXPECT_NE(wb.sheet(0).cell_at(3U, 1U)->formula_text.find("SUBTOTAL(109,"), std::string::npos);
  // Showing it again changes nothing.
  ASSERT_TRUE(wb.set_table_totals_row(0, true).has_value());
  EXPECT_EQ(wb.tables()[0].ref, "A1:B4");
}

TEST(WorkbookTableTotalsRow, AnOccupiedRowBelowShiftsDownWithinTheTableColumns) {
  {
    Workbook wb = Workbook::create();
    AddRegions(wb);
    ASSERT_TRUE(wb.set_cell_text(0, 3U, 1U, "x").has_value());  // B4
    ASSERT_TRUE(wb.set_table_totals_row(0, true).has_value());
    EXPECT_EQ(wb.tables()[0].ref, "A1:B4");
    EXPECT_EQ(At(wb, 4U, 1U).as_text(), "x");
    EXPECT_EQ(At(wb, 3U, 0U).as_text(), "Total");
    EXPECT_EQ(At(wb, 2U, 1U), Value::number(30.0));
  }
  {
    Workbook wb = Workbook::create();
    AddRegions(wb);
    ASSERT_TRUE(wb.set_cell_formula(0, 3U, 0U, "=1+1").has_value());  // A4
    ASSERT_TRUE(wb.set_table_totals_row(0, true).has_value());
    EXPECT_EQ(wb.tables()[0].ref, "A1:B4");
    ASSERT_NE(wb.sheet(0).cell_at(4U, 0U), nullptr);
    EXPECT_EQ(wb.sheet(0).cell_at(4U, 0U)->formula_text, "=1+1");
  }
}

TEST(WorkbookTableTotalsRow, AStraddlingObjectOnTheRowBelowIsRefusedAndNothingChanges) {
  {
    // B4:C5 straddles the table's columns, so nothing shifts, and the row
    // below is part of the merge.
    Workbook wb = Workbook::create();
    AddRegions(wb);
    ASSERT_TRUE(wb.add_merge(0, MergeRange{3U, 1U, 4U, 2U}).has_value());  // B4:C5
    ExpectTotalsRowRefused(wb, "A1:B3");
    ASSERT_EQ(wb.sheet(0).merges().size(), 1U);
    EXPECT_EQ(wb.sheet(0).merges()[0], (MergeRange{3U, 1U, 4U, 2U}));
  }
  {
    Workbook wb = Workbook::create();
    AddRegions(wb);
    ASSERT_TRUE(wb.add_table(make_table("Below", "B4:C6", 2)).has_value());
    ExpectTotalsRowRefused(wb, "A1:B3");
    EXPECT_EQ(wb.tables()[1].ref, "B4:C6");
  }
}

TEST(WorkbookTableTotalsRow, ASpillCrossingTheNewTotalsRowIsBlocked) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(wb.add_table(make_table("Right", "C1:D3", 2)).has_value());
  ASSERT_TRUE(wb.set_cell_formula(0, 3U, 0U, "=SEQUENCE(1,4)").has_value());  // A4 spills to D4
  Recalc(wb);
  ASSERT_TRUE(wb.set_table_totals_row(0, true).has_value());
  EXPECT_EQ(wb.tables()[0].ref, "C1:D4");
  Recalc(wb);
  EXPECT_EQ(At(wb, 3U, 0U), Value::error(ErrorCode::Spill));
}

TEST(WorkbookTableTotalsRow, ATableOnTheLastGridRowCannotShowIt) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(wb.add_table(make_table("Bottom", "A1048574:B1048576", 2)).has_value());
  ExpectTotalsRowRefused(wb, "A1048574:B1048576");
}

TEST(WorkbookTableTotalsRow, HidingShrinksTheTableAndClearsTheRow) {
  Workbook wb = Workbook::create();
  AddRegions(wb);
  ASSERT_TRUE(wb.set_table_totals_row(0, true).has_value());
  ASSERT_TRUE(wb.set_table_totals_row(0, false).has_value());
  EXPECT_EQ(wb.tables()[0].ref, "A1:B3");
  EXPECT_FALSE(wb.tables()[0].totals_row);
  Recalc(wb);
  EXPECT_EQ(At(wb, 2U, 1U), Value::number(30.0));
  EXPECT_TRUE(At(wb, 3U, 0U).is_blank());
  EXPECT_TRUE(At(wb, 3U, 1U).is_blank());
  // Hiding it again changes nothing.
  ASSERT_TRUE(wb.set_table_totals_row(0, false).has_value());
  EXPECT_EQ(wb.tables()[0].ref, "A1:B3");
}

TEST(WorkbookTableTotalsRow, HidingKeepsADataRowUnderTheHeader) {
  Workbook wb = Workbook::create();
  TableMetadata table = make_table("Thin", "A1:B2", 2);
  table.totals_row = true;
  ASSERT_TRUE(wb.add_table(std::move(table)).has_value());
  const auto hidden = wb.set_table_totals_row(0, false);
  ASSERT_FALSE(hidden.has_value());
  EXPECT_EQ(hidden.error().code, FormulonErrorCode::kPreconditionFailed);
  EXPECT_EQ(wb.tables()[0].ref, "A1:B2");
  EXPECT_TRUE(wb.tables()[0].totals_row);
}

TEST(WorkbookTableTotalsRow, AnOutOfRangeIndexIsAnInvalidArgument) {
  Workbook wb = Workbook::create();
  const auto shown = wb.set_table_totals_row(0, true);
  ASSERT_FALSE(shown.has_value());
  EXPECT_EQ(shown.error().code, FormulonErrorCode::kInvalidArgument);
}

TEST(WorkbookMerges, AMergeIntersectingAnotherIsAlreadyExists) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(wb.add_merge(0, MergeRange{0U, 0U, 1U, 1U}).has_value());                      // A1:B2
  for (const MergeRange merge : {MergeRange{1U, 1U, 2U, 2U}, MergeRange{0U, 0U, 1U, 1U}}) {  // B2:C3, A1:B2
    const auto added = wb.add_merge(0, merge);
    ASSERT_FALSE(added.has_value());
    EXPECT_EQ(added.error().code, FormulonErrorCode::kAlreadyExists);
  }
  EXPECT_TRUE(wb.add_merge(0, MergeRange{0U, 2U, 1U, 3U}).has_value());  // C1:D2, adjacent
  EXPECT_EQ(wb.sheet(0).merges().size(), 2U);
}

TEST(WorkbookMerges, AMergeOverATablePivotArrayFormulaOrDataTableIsPreconditionFailed) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(wb.add_table(make_table("Sales", "A1:B4", 2)).has_value());
  ASSERT_TRUE(wb.set_array_formula(0, 9U, 1U, 1U, 1U, "=1").has_value());  // B10
  DataTableInput input;
  input.present = true;
  ASSERT_TRUE(wb.add_data_table(0, MergeRange{41U, 1U, 43U, 1U}, DataTableInput{}, input).has_value());  // B42:B44
  auto pivot = std::make_unique<pivot::PivotTable>();
  pivot->set_anchor(/*row=*/50U, /*col=*/0U, /*rows=*/3U, /*cols=*/2U);  // A51:B53
  wb.sheet(0).add_pivot_table(std::move(pivot));

  for (const MergeRange merge : {MergeRange{3U, 1U, 4U, 2U}, MergeRange{9U, 0U, 9U, 1U}, MergeRange{43U, 1U, 44U, 1U},
                                 MergeRange{52U, 1U, 53U, 2U}}) {
    const auto added = wb.add_merge(0, merge);
    ASSERT_FALSE(added.has_value());
    EXPECT_EQ(added.error().code, FormulonErrorCode::kPreconditionFailed);
  }
  EXPECT_TRUE(wb.sheet(0).merges().empty());
  EXPECT_TRUE(wb.add_merge(0, MergeRange{4U, 0U, 5U, 1U}).has_value());  // A5:B6, below the table
}

}  // namespace
}  // namespace formulon
