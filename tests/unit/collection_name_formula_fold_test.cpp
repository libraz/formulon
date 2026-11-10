// Names typed in formulas match defined names, tables and table columns under
// the same Unicode simple case fold that collection identity uses.

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "cell.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "sheet_name.h"
#include "table.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

void Recalc(Workbook& wb) {
  ASSERT_TRUE(wb.recalc(eval::default_registry()));
}

Value At(const Workbook& wb, std::uint32_t row, std::uint32_t col) {
  return wb.sheet(0).resolve_cell_value(row, col);
}

std::string FormulaAt(const Workbook& wb, std::uint32_t row, std::uint32_t col) {
  const Cell* cell = wb.sheet(0).cell_at(row, col);
  return cell == nullptr ? std::string() : cell->formula_text;
}

TableMetadata MakeTable(std::string name, std::string ref, const std::vector<std::string>& columns) {
  TableMetadata table;
  table.sheet_index = 0;
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

// TÄBLE1 on A1:B3: headers Ärger / Wert, data rows B2:B3 = 2, 3.
Workbook MakeTableBook() {
  Workbook wb = Workbook::create();
  EXPECT_TRUE(wb.set_cell_text(0, 0, 0, "Ärger"));
  EXPECT_TRUE(wb.set_cell_text(0, 0, 1, "Wert"));
  for (std::uint32_t r = 1; r <= 2; ++r) {
    EXPECT_TRUE(wb.set_cell_value(0, r, 0, Value::number(10.0 * r)));
    EXPECT_TRUE(wb.set_cell_value(0, r, 1, Value::number(1.0 + r)));
  }
  EXPECT_TRUE(wb.add_table(MakeTable("TÄBLE1", "A1:B3", {"Ärger", "Wert"})).has_value());
  return wb;
}

TEST(CollectionNameFormulaFold, FoldKeyAgreesWithEqual) {
  using std::string_view_literals::operator""sv;
  const std::vector<std::string_view> names = {
      "Abc"sv,      "aBC"sv,      "abd"sv,          ""sv,         "a\0b"sv,    "A\0B"sv, "Äpfel"sv,
      "äPFEL"sv,    "ÄBC"sv,      "äbc"sv,          "\u212A"sv,   "k"sv,       "K"sv,    "\u1E9E"sv,
      "\u00DF"sv,   "ss"sv,       "\u03A3"sv,       "\u03C3"sv,   "\u03C2"sv,  "\xC3"sv, "\xC3\xA4\xC3"sv,
      "\xFF\x41"sv, "\xFF\x61"sv, "\xED\xA0\x80"sv, "\xC3\x84"sv, "\xC3\xA4"sv};
  for (std::string_view a : names) {
    for (std::string_view b : names) {
      EXPECT_EQ(collection_names::equal(a, b), collection_names::fold_key(a) == collection_names::fold_key(b))
          << testing::PrintToString(std::string(a)) << " vs " << testing::PrintToString(std::string(b));
    }
  }
  EXPECT_TRUE(collection_names::equal("ÄBC", "äbc"));
  EXPECT_TRUE(collection_names::equal("\u212A", "k"));
  EXPECT_TRUE(collection_names::equal("\xFF\x41", "\xFF\x41"));
  EXPECT_FALSE(collection_names::equal("\xFF\x41", "\xFF\x61"));
  EXPECT_FALSE(collection_names::equal("\u00DF", "ss"));
}

TEST(CollectionNameFormulaFold, DefinedNameResolvesAcrossNonAsciiCase) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(wb.set_cell_value(0, 0, 0, Value::number(5.0)));
  ASSERT_TRUE(wb.set_cell_value(0, 1, 0, Value::number(7.0)));
  ASSERT_TRUE(wb.set_defined_name("Äpfel", "Sheet1!$A$1"));
  ASSERT_TRUE(wb.set_cell_formula(0, 0, 2, "=äPFEL*2"));
  Recalc(wb);
  EXPECT_EQ(At(wb, 0, 2), Value::number(10.0));

  // Redefining through another spelling keeps the entry, takes the new spelling and dirties the cell.
  ASSERT_TRUE(wb.set_defined_name("ÄPFEL", "Sheet1!$A$2"));
  ASSERT_EQ(wb.defined_names().size(), 1U);
  EXPECT_EQ(wb.defined_names()[0].name, "ÄPFEL");
  Recalc(wb);
  EXPECT_EQ(At(wb, 0, 2), Value::number(14.0));
}

TEST(CollectionNameFormulaFold, StructuredRefResolvesNonAsciiTableAndColumnAcrossCase) {
  Workbook wb = MakeTableBook();
  ASSERT_TRUE(wb.set_cell_formula(0, 0, 4, "=SUM(täble1[äRGER])"));
  ASSERT_TRUE(wb.set_cell_formula(0, 1, 4, "=SUM(TÄBLE1[wert])"));
  Recalc(wb);
  EXPECT_EQ(At(wb, 0, 4), Value::number(30.0));
  EXPECT_EQ(At(wb, 1, 4), Value::number(5.0));
}

TEST(CollectionNameFormulaFold, TableRenameRewritesReferenceSpelledInOtherCase) {
  Workbook wb = MakeTableBook();
  ASSERT_TRUE(wb.set_cell_formula(0, 0, 4, "=SUM(täble1[Wert])"));
  ASSERT_TRUE(wb.rename_table(0, "Übersicht", "Übersicht"));
  EXPECT_EQ(FormulaAt(wb, 0, 4), "=SUM(Übersicht[Wert])");
  Recalc(wb);
  EXPECT_EQ(At(wb, 0, 4), Value::number(5.0));
}

TEST(CollectionNameFormulaFold, ColumnRenameCollidingAcrossNonAsciiCaseGainsSuffix) {
  Workbook wb = MakeTableBook();
  Expected<std::string, Error> renamed = wb.rename_table_column(0, 1, "äRGER");
  ASSERT_TRUE(renamed.has_value());
  EXPECT_EQ(renamed.value(), "äRGER2");
  EXPECT_EQ(wb.tables()[0].columns[1].name, "äRGER2");
}

TEST(CollectionNameFormulaFold, ColumnRenameRewritesReferenceSpelledInOtherCase) {
  Workbook wb = MakeTableBook();
  ASSERT_TRUE(wb.set_cell_formula(0, 0, 4, "=SUM(TÄBLE1[äRGER])"));
  Expected<std::string, Error> renamed = wb.rename_table_column(0, 0, "Öl");
  ASSERT_TRUE(renamed.has_value());
  EXPECT_EQ(FormulaAt(wb, 0, 4), "=SUM(TÄBLE1[Öl])");
}

TEST(CollectionNameFormulaFold, NewTableNumbersColumnsEqualAcrossNonAsciiCase) {
  Workbook wb = Workbook::create();
  Expected<std::size_t, Error> added = wb.add_table(MakeTable("Tbl", "A1:B3", {"Ärger", "äRGER"}));
  ASSERT_TRUE(added.has_value());
  EXPECT_EQ(wb.tables()[0].columns[1].name, "äRGER2");
}

}  // namespace
}  // namespace formulon
