//
// A structured reference reads as the A1 range it covers: reference-taking
// functions see its rectangle, and a lookup over a table holding the formula
// cell reads only the cells the lookup touches.

#include <cstdint>
#include <sstream>
#include <string>
#include <utility>
#include <vector>

#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "table.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

std::string DisplayOf(const Sheet& s, std::uint32_t row, std::uint32_t col) {
  const Value v = s.resolve_cell_value(row, col);
  if (v.is_number()) {
    std::ostringstream out;
    out << v.as_number();
    return out.str();
  }
  if (v.is_error()) {
    return std::string(display_name(v.as_error()));
  }
  if (v.is_text()) {
    return std::string(v.as_text());
  }
  return "";
}

TableColumn Column(std::uint32_t id, std::string name) {
  TableColumn c;
  c.id = id;
  c.name = std::move(name);
  return c;
}

void AddTable(Workbook& wb, std::string ref, std::vector<TableColumn> columns) {
  TableMetadata table;
  table.id = 1;
  table.name = "T";
  table.display_name = "T";
  table.ref = std::move(ref);
  table.sheet_index = 0;
  table.header_row = true;
  table.columns = std::move(columns);
  wb.set_tables({table});
}

// T on A1:B4: ID = 1, 2, 3 and p = 10, 20, 30.
Workbook IdAndPTable() {
  Workbook wb = Workbook::create();
  wb.set_excel_profile(mac_365_ja_jp_profile());
  Sheet& s = wb.sheet(0);
  s.set_cell_value(0U, 0U, Value::text("ID"));
  s.set_cell_value(0U, 1U, Value::text("p"));
  for (std::uint32_t i = 0; i < 3U; ++i) {
    s.set_cell_value(i + 1U, 0U, Value::number(i + 1.0));
    s.set_cell_value(i + 1U, 1U, Value::number(10.0 * (i + 1.0)));
  }
  AddTable(wb, "A1:B4", {Column(1U, "ID"), Column(2U, "p")});
  return wb;
}

TEST(StructuredRefRange, RowOfAColumnIsItsFirstRowInALegacyCell) {
  Workbook wb = IdAndPTable();
  Sheet& s = wb.sheet(0);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 4U, "=ROW(T[p])")));
  s.set_cell_dynamic_array(0U, 4U, false);
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_EQ(DisplayOf(s, 0U, 4U), "2");
}

TEST(StructuredRefRange, RowOfAColumnSpillsItsRowsInADynamicArrayCell) {
  Workbook wb = IdAndPTable();
  const Sheet& s = wb.sheet(0);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 5U, "=ROW(T[p])")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_EQ(DisplayOf(s, 0U, 5U), "2");
  EXPECT_EQ(DisplayOf(s, 1U, 5U), "3");
  EXPECT_EQ(DisplayOf(s, 2U, 5U), "4");
}

TEST(StructuredRefRange, RowOfTheHeaderCellAndAggregateOverAColumn) {
  Workbook wb = IdAndPTable();
  const Sheet& s = wb.sheet(0);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 6U, "=ROW(T[[#Headers],[ID]])")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 7U, "=AGGREGATE(9,5,T[p])")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_EQ(DisplayOf(s, 0U, 6U), "1");
  EXPECT_EQ(DisplayOf(s, 0U, 7U), "60");
}

// Excel 365 (Mac, ja-JP): T on A1:C3 with k = x, y and v = 1, 2; the lookups
// in T's own column c read only the cells they touch, so they are no cycle.
TEST(StructuredRefRange, LookupOverTheFormulasOwnTableReadsOnlyWhatItTouches) {
  Workbook wb = Workbook::create();
  wb.set_excel_profile(mac_365_ja_jp_profile());
  Sheet& s = wb.sheet(0);
  s.set_cell_value(0U, 0U, Value::text("k"));
  s.set_cell_value(0U, 1U, Value::text("v"));
  s.set_cell_value(0U, 2U, Value::text("c"));
  s.set_cell_value(1U, 0U, Value::text("x"));
  s.set_cell_value(1U, 1U, Value::number(1.0));
  s.set_cell_value(2U, 0U, Value::text("y"));
  s.set_cell_value(2U, 1U, Value::number(2.0));
  AddTable(wb, "A1:C3", {Column(1U, "k"), Column(2U, "v"), Column(3U, "c")});
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 2U, "=VLOOKUP(D2,T[],2,FALSE)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 2U, 2U, "=VLOOKUP(D3,T[],2,FALSE)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 2U, 4U, "=VLOOKUP(\"x\",T[],2,FALSE)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_EQ(DisplayOf(s, 1U, 2U), "#N/A");
  EXPECT_EQ(DisplayOf(s, 2U, 2U), "#N/A");
  EXPECT_EQ(DisplayOf(s, 2U, 4U), "1");
}

}  // namespace
}  // namespace formulon
