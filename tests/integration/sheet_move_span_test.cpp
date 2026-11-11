// Sheet moves rewrite 3-D spans only when an endpoint crosses the other.
#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>

#include "sheet_ops_test_support.h"

namespace formulon {
namespace {
using namespace sheet_ops_test;

// "F" holds the formula under test at A1; S1..S6 follow in tab order, each
// holding its own number at A1.
Workbook MakeSpanWorkbook(const char* formula) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("F");
  for (std::uint32_t i = 1; i <= 6U; ++i) {
    wb.add_sheet("S" + std::to_string(i));
    EXPECT_TRUE(static_cast<bool>(wb.set_cell_value(i, 0, 0, Value::number(static_cast<double>(i)))));
  }
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 0, formula)));
  EXPECT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  return wb;
}

const Cell& FormulaCell(const Workbook& wb) {
  const Cell* cell = wb.sheet(0).cell_at(0, 0);
  EXPECT_NE(cell, nullptr);
  static const Cell kEmpty{};
  return cell == nullptr ? kEmpty : *cell;
}

TEST(WorkbookSheetOps, MoveSheetRewritesSpans) {
  struct Row {
    std::uint32_t from;
    std::uint32_t to;
    const char* text;
    double value;
  };
  // Indices are positions in the list "F S1 S2 S3 S4 S5 S6"; `to` is a position after the lift.
  const Row rows[] = {
      {3, 4, "=SUM('S2:S4'!A1)", 2 + 4},              // middle sheet leaves the span
      {3, 1, "=SUM('S2:S4'!A1)", 2 + 4},              // middle sheet moves out the other way
      {2, 4, "=SUM('S3:S4'!A1)", 3 + 4},              // first endpoint moved past the last
      {4, 2, "=SUM('S2:S3'!A1)", 2 + 3},              // last endpoint moved before the first
      {2, 3, "=SUM('S2:S4'!A1)", 2 + 4},              // first endpoint moved within
      {2, 1, "=SUM('S2:S4'!A1)", 1 + 2 + 3 + 4},      // first endpoint moved out the front
      {4, 3, "=SUM('S2:S4'!A1)", 2 + 4},              // last endpoint moved within
      {4, 6, "=SUM('S2:S4'!A1)", 2 + 3 + 5 + 6 + 4},  // last endpoint moved out the back
      {1, 6, "=SUM('S2:S4'!A1)", 2 + 3 + 4},          // sheet outside the span
  };
  for (const Row& row : rows) {
    SCOPED_TRACE("from=" + std::to_string(row.from) + " to=" + std::to_string(row.to));
    Workbook wb = MakeSpanWorkbook("=SUM('S2:S4'!A1)");
    ASSERT_TRUE(static_cast<bool>(wb.move_sheet(row.from, row.to)));
    ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    EXPECT_EQ(FormulaCell(wb).formula_text, row.text);
    EXPECT_EQ(FormulaCell(wb).cached_value.as_number(), row.value);
  }
}

TEST(WorkbookSheetOps, MoveSheetCollapsesASpanThatBecomesOneSheet) {
  Workbook wb = MakeSpanWorkbook("=SUM('S2:S3'!A1)");
  ASSERT_TRUE(static_cast<bool>(wb.move_sheet(2, 3)));  // F S1 S3 S2 ...
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_EQ(FormulaCell(wb).formula_text, "=SUM('S3'!A1)");
  EXPECT_EQ(FormulaCell(wb).cached_value.as_number(), 3.0);
}

TEST(WorkbookSheetOps, MoveSheetCollapsesASameEndpointSpanOnlyWhenItsSheetMoves) {
  {
    Workbook wb = MakeSpanWorkbook("=SUM('S3:S3'!A1)");
    ASSERT_TRUE(static_cast<bool>(wb.move_sheet(3, 5)));
    ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    EXPECT_EQ(FormulaCell(wb).formula_text, "=SUM('S3'!A1)");
    EXPECT_EQ(FormulaCell(wb).cached_value.as_number(), 3.0);
  }
  {
    Workbook wb = MakeSpanWorkbook("=SUM('S3:S3'!A1)");
    ASSERT_TRUE(static_cast<bool>(wb.move_sheet(1, 5)));
    EXPECT_EQ(FormulaCell(wb).formula_text, "=SUM('S3:S3'!A1)");
  }
}

TEST(WorkbookSheetOps, MoveSheetReadsAReversedSpanInTabOrder) {
  Workbook wb = MakeSpanWorkbook("=SUM('S4:S2'!A1)");
  ASSERT_TRUE(static_cast<bool>(wb.move_sheet(2, 4)));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_EQ(FormulaCell(wb).formula_text, "=SUM('S3:S4'!A1)");
  EXPECT_EQ(FormulaCell(wb).cached_value.as_number(), 3 + 4.0);
}

TEST(WorkbookSheetOps, MoveSheetLeavesASpanWithAnUnknownEndpoint) {
  Workbook wb = MakeSpanWorkbook("=SUM('S2:Gone'!A1)");
  ASSERT_TRUE(static_cast<bool>(wb.move_sheet(2, 5)));
  EXPECT_EQ(FormulaCell(wb).formula_text, "=SUM('S2:Gone'!A1)");
}

TEST(WorkbookSheetOps, MoveSheetRewritesSpanInNameAndCf) {
  Workbook wb = MakeSpanWorkbook("=1");
  wb.set_defined_names({DefinedName{"Rng", "'S2:S4'!$A$1", -1, false, ""}});
  cf::ConditionalFormat conditional_format;
  conditional_format.sqref.push_back(cf::CFCellRange{CellAddress{0, 2}, CellAddress{9, 2}});
  cf::CFRule rule;
  rule.type = cf::RuleType::Expression;
  rule.priority = 1;
  rule.formula1 = "SUM('S2:S4'!A1)>1";
  conditional_format.rules.push_back(std::move(rule));
  wb.sheet(0).mutable_conditional_formats().push_back(std::move(conditional_format));

  ASSERT_TRUE(static_cast<bool>(wb.move_sheet(2, 4)));
  ASSERT_EQ(wb.defined_names().size(), 1U);
  EXPECT_EQ(wb.defined_names()[0].formula, "'S3:S4'!$A$1");
  EXPECT_EQ(wb.sheet(0).conditional_formats()[0].rules[0].formula1.value(), "SUM('S3:S4'!A1)>1");

  // A move that keeps both endpoints on the same side leaves both holders alone.
  ASSERT_TRUE(static_cast<bool>(wb.move_sheet(1, 5)));  // S1 to after S5; F stays first
  EXPECT_EQ(wb.defined_names()[0].formula, "'S3:S4'!$A$1");
  EXPECT_EQ(wb.sheet(0).conditional_formats()[0].rules[0].formula1.value(), "SUM('S3:S4'!A1)>1");
}

TEST(WorkbookSheetOps, AddSheetAtPositionLeavesFormulasAlone) {
  Workbook wb = MakeSpanWorkbook("=SUM('S2:S4'!A1)");
  EXPECT_FALSE(static_cast<bool>(wb.add_sheet_entered("Far", wb.sheet_count() + 1U)));
  auto added = wb.add_sheet_entered("New", 3);  // lands inside the span
  ASSERT_TRUE(static_cast<bool>(added));
  EXPECT_EQ(added.value()->name(), "New");
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_EQ(FormulaCell(wb).formula_text, "=SUM('S2:S4'!A1)");
  EXPECT_EQ(FormulaCell(wb).cached_value.as_number(), 2 + 3 + 4.0);
  EXPECT_EQ(wb.sheet_index_by_name("New"), 3U);
}

}  // namespace
}  // namespace formulon
