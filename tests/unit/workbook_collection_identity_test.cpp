#include <cstddef>
#include <cstdint>
#include <memory>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/sparkline_groups.h"
#include "pivot/pivot_index.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "sheet_metadata.h"
#include "table.h"
#include "utils/error.h"
#include "workbook.h"

namespace formulon {
namespace {

FormulonErrorCode Code(const Expected<void, Error>& r) {
  return r ? FormulonErrorCode::kOk : r.error().code;
}

TableMetadata Table(std::string name, std::string ref) {
  TableMetadata table;
  table.sheet_index = 0;
  table.name = name;
  table.display_name = std::move(name);
  table.ref = std::move(ref);
  for (const char* column : {"C1", "C2"}) {
    TableColumn c;
    c.name = column;
    table.columns.push_back(std::move(c));
  }
  return table;
}

DataValidation Validation(std::uint32_t row) {
  DataValidation dv;
  dv.type = 1;
  dv.formula1 = "1";
  dv.ranges.push_back(MergeRange{row, 0, row, 0});
  return dv;
}

io::SparklineGroup Group(const std::string& location) {
  io::SparklineGroup group;
  group.items.push_back(io::SparklineItem{"Sheet1!A1:C1", location});
  return group;
}

Workbook ForProfile(const char* id) {
  ExcelProfile profile;
  EXPECT_TRUE(parse_excel_profile_id(id, &profile));
  return Workbook::create_for_profile(profile);
}

// --- Collection names fold like sheet names ---

TEST(CollectionIdentity, TableNamesCollideUnderUnicodeCaseFolding) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(wb.add_table(Table("Äpfel", "A1:B5")).has_value());
  auto clash = wb.add_table(Table("äPFEL", "D1:E5"));
  ASSERT_FALSE(clash.has_value());
  EXPECT_EQ(clash.error().code, FormulonErrorCode::kAlreadyExists);
  EXPECT_EQ(Code(wb.set_defined_name_scoped("äpfel", "1", -1)), FormulonErrorCode::kAlreadyExists);
}

TEST(CollectionIdentity, DefinedNamesFoldBeyondAscii) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name_scoped("Ölpreis", "1", -1)));
  // The other casing names the same entry: it is updated, not appended.
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name_scoped("ölPREIS", "2", -1)));
  ASSERT_EQ(wb.defined_names().size(), 1U);
  EXPECT_EQ(wb.defined_names()[0].name, "Ölpreis");
  EXPECT_EQ(wb.defined_names()[0].formula, "2");
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name_hidden("ÖLPREIS", -1, true)));
  EXPECT_TRUE(wb.defined_names()[0].hidden);
  auto table = wb.add_table(Table("ölpreis", "A1:B5"));
  ASSERT_FALSE(table.has_value());
  EXPECT_EQ(table.error().code, FormulonErrorCode::kAlreadyExists);
}

TEST(CollectionIdentity, PivotNamesFoldOnTheirSheet) {
  Workbook wb = Workbook::create();
  auto pivot = std::make_unique<pivot::PivotTable>();
  pivot->set_name("Übersicht");
  wb.sheet(0).add_pivot_table(std::move(pivot));
  const Sheet& sheet = wb.sheet(0);
  EXPECT_TRUE(pivot::pivot_name_taken(sheet, "üBERSICHT", sheet.pivot_tables().size()));
  EXPECT_FALSE(pivot::pivot_name_taken(sheet, "üBERSICHT", 0));
  EXPECT_FALSE(pivot::pivot_name_taken(sheet, "Other", sheet.pivot_tables().size()));
}

// --- Default sheet names ---

TEST(CollectionIdentity, DefaultSheetNameTakesTheSmallestFreeNumber) {
  Workbook en = Workbook::create();
  EXPECT_EQ(en.next_default_sheet_name(), "Sheet2");

  Workbook de = ForProfile("win-365-de_DE");
  ASSERT_EQ(de.sheet(0).name(), "Tabelle1");
  EXPECT_EQ(de.next_default_sheet_name(), "Tabelle2");
  ASSERT_TRUE(static_cast<bool>(de.add_sheet_entered("tabelle2")));
  EXPECT_EQ(de.next_default_sheet_name(), "Tabelle3");
  ASSERT_TRUE(static_cast<bool>(de.add_sheet_entered("Tabelle4")));
  ASSERT_TRUE(static_cast<bool>(de.remove_sheet(0)));
  EXPECT_EQ(de.next_default_sheet_name(), "Tabelle1");
}

// --- Data-validation ids ---

TEST(CollectionIdentity, ValidationIdsSurviveEditsAndRemovals) {
  Workbook wb = Workbook::create();
  auto first = wb.add_validation(0, Validation(4));
  auto second = wb.add_validation(0, Validation(6));
  ASSERT_TRUE(first.has_value());
  ASSERT_TRUE(second.has_value());
  EXPECT_NE(first.value(), second.value());

  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, 0, 2)));
  EXPECT_EQ(wb.find_validation(0, first.value()), std::optional<std::size_t>(0));
  EXPECT_EQ(wb.sheet(0).validations()[0].ranges[0].first_row, 6U);

  ASSERT_TRUE(static_cast<bool>(wb.remove_validation(0, first.value())));
  EXPECT_EQ(wb.find_validation(0, second.value()), std::optional<std::size_t>(0));
  EXPECT_EQ(Code(wb.remove_validation(0, first.value())), FormulonErrorCode::kNotFound);
  EXPECT_EQ(Code(wb.remove_validation(5, second.value())), FormulonErrorCode::kSheetIndexOutOfRange);

  // A fresh id is never one handed out before.
  auto third = wb.add_validation(0, Validation(9));
  ASSERT_TRUE(third.has_value());
  EXPECT_NE(third.value(), first.value());
  EXPECT_NE(third.value(), second.value());
}

TEST(CollectionIdentity, ValidationIdsAreAssignedOnRequestAndFreshOnACopy) {
  Workbook wb = Workbook::create();
  wb.sheet(0).mutable_validations().push_back(Validation(0));
  const std::uint32_t id = wb.validation_runtime_id(0, 0);
  EXPECT_NE(id, 0U);
  EXPECT_EQ(wb.validation_runtime_id(0, 0), id);

  auto copy = wb.copy_sheet(0, 1, std::nullopt);
  ASSERT_TRUE(copy.has_value());
  const std::uint32_t copied = wb.validation_runtime_id(copy.value(), 0);
  EXPECT_NE(copied, id);
  EXPECT_EQ(wb.find_validation(copy.value(), id), std::nullopt);
  EXPECT_EQ(wb.find_validation(0, id), std::optional<std::size_t>(0));
}

// --- Sparkline-group ids ---

TEST(CollectionIdentity, SparklineGroupIdsFollowTheirGroups) {
  Workbook wb = Workbook::create();
  auto top = io::add_sheet_sparkline_group(wb, 0, Group("D1"));
  auto middle = io::add_sheet_sparkline_group(wb, 0, Group("D3"));
  auto bottom = io::add_sheet_sparkline_group(wb, 0, Group("D5"));
  ASSERT_TRUE(top.has_value());
  ASSERT_TRUE(middle.has_value());
  ASSERT_TRUE(bottom.has_value());
  EXPECT_EQ(wb.sparkline_group_runtime_id(0, 2), bottom.value());

  // Deleting the middle group's row drops it; the groups after it keep their ids.
  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, 2, 1)));
  ASSERT_EQ(io::read_sparkline_groups(wb.sheet(0).ext_lst_xml()).size(), 2U);
  EXPECT_EQ(wb.find_sparkline_group(0, middle.value()), std::nullopt);
  EXPECT_EQ(wb.find_sparkline_group(0, bottom.value()), std::optional<std::size_t>(1));

  ASSERT_TRUE(static_cast<bool>(io::remove_sheet_sparkline_group(wb, 0, top.value())));
  EXPECT_EQ(wb.find_sparkline_group(0, bottom.value()), std::optional<std::size_t>(0));
  EXPECT_EQ(Code(io::remove_sheet_sparkline_group(wb, 0, top.value())), FormulonErrorCode::kNotFound);

  auto copy = wb.copy_sheet(0, 1, std::nullopt);
  ASSERT_TRUE(copy.has_value());
  EXPECT_NE(wb.sparkline_group_runtime_id(copy.value(), 0), bottom.value());
  EXPECT_EQ(wb.find_sparkline_group(copy.value(), bottom.value()), std::nullopt);
}

TEST(CollectionIdentity, SparklineGroupAddRefusesWithoutChangingTheSheet) {
  Workbook wb = Workbook::create();
  auto refused = io::add_sheet_sparkline_group(wb, 0, io::SparklineGroup{});
  ASSERT_FALSE(refused.has_value());
  EXPECT_EQ(refused.error().code, FormulonErrorCode::kInvalidArgument);
  EXPECT_TRUE(wb.sheet(0).ext_lst_xml().empty());
  EXPECT_TRUE(wb.sheet(0).sparkline_group_ids().empty());
}

}  // namespace
}  // namespace formulon
