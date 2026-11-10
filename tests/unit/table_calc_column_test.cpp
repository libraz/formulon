// Calculated columns: Excel's window rule on table creation (the measured
// table of column C sequences over a header row 2 and data from row 3), the
// set / clear / drop / re-infer entries, and the shared-formula group a set
// registers, as the `.xlsx` writer verifies it and as edits move it.

#include "table_calc_column.h"

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "cell.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/ooxml_writer_cell.h"
#include "io/zip_reader.h"
#include "merge_range.h"
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

void SetFormula(Workbook& wb, std::string_view a1, std::string formula) {
  const CellAddress at = Addr(a1);
  ASSERT_TRUE(wb.set_cell_formula(0U, at.row, at.col, std::move(formula)).has_value()) << a1;
}

void SetBlank(Workbook& wb, std::string_view a1) {
  const CellAddress at = Addr(a1);
  ASSERT_TRUE(wb.set_cell_value(0U, at.row, at.col, Value::blank()).has_value()) << a1;
}

std::string FormulaAt(const Workbook& wb, std::string_view a1, std::size_t sheet = 0U) {
  const CellAddress at = Addr(a1);
  const Cell* cell = wb.sheet(sheet).cell_at(at.row, at.col);
  return cell == nullptr ? std::string() : cell->formula_text;
}

std::string Mark(const Workbook& wb, std::size_t column, std::size_t table = 0U) {
  return wb.tables()[table].columns[column].calculated_column_formula;
}

// Header row 2 (`a`, `b`, `c`), A = 1..N and B = 5.. from row 3, and column C
// from `pattern`, one character per data row: `A` = `=A{r}*10`, `B` =
// `=A{r}*20`, `c` = the constant 99, `_` = blank.
Workbook PatternBook(std::string_view pattern) {
  Workbook wb = Workbook::create();
  SetText(wb, "A2", "a");
  SetText(wb, "B2", "b");
  SetText(wb, "C2", "c");
  for (std::size_t i = 0; i < pattern.size(); ++i) {
    const std::string r = std::to_string(3U + i);
    SetNumber(wb, "A" + r, static_cast<double>(i + 1U));
    SetNumber(wb, "B" + r, static_cast<double>(i + 5U));
    switch (pattern[i]) {
      case 'A':
        SetFormula(wb, "C" + r, "=A" + r + "*10");
        break;
      case 'B':
        SetFormula(wb, "C" + r, "=A" + r + "*20");
        break;
      case 'c':
        SetNumber(wb, "C" + r, 99);
        break;
      default:
        break;
    }
  }
  return wb;
}

TableMetadata MakeTable(std::string ref, bool totals = false) {
  TableMetadata table;
  table.sheet_index = 0U;
  table.name = "Tbl";
  table.display_name = "Tbl";
  table.ref = std::move(ref);
  table.header_row = true;
  table.totals_row = totals;
  for (const char* name : {"a", "b", "c"}) {
    TableColumn column;
    column.name = name;
    table.columns.push_back(std::move(column));
  }
  return table;
}

void AddTable(Workbook& wb, std::string ref, bool totals = false) {
  ASSERT_TRUE(wb.add_table(MakeTable(std::move(ref), totals)).has_value());
}

// The table of `pattern` over header row 2 through its last data row.
std::string PatternRef(std::string_view pattern) {
  return "A2:C" + std::to_string(2U + pattern.size());
}

void ExpectCode(const Expected<void, Error>& result, FormulonErrorCode code) {
  ASSERT_FALSE(result.has_value());
  EXPECT_EQ(result.error().code, code);
}

std::string SheetXml(const Workbook& wb, std::size_t sheet = 0U) {
  std::string out;
  io::AppendSheetDataXml(out, wb.sheet(sheet));
  return out;
}

std::size_t CountOf(std::string_view text, std::string_view needle) {
  std::size_t count = 0;
  for (std::size_t at = text.find(needle); at != std::string_view::npos; at = text.find(needle, at + 1U)) {
    ++count;
  }
  return count;
}

// ---------------------------------------------------------------------------
// The window rule
// ---------------------------------------------------------------------------

struct WindowCase {
  std::string pattern;
  /// The mark Excel gave column C; empty for none.
  std::string mark;
};

std::vector<WindowCase> WindowCases() {
  const std::string a = "A3*10";
  const std::string b = "A3*20";
  const std::string none;
  return {
      {"A", a},
      {"AA", a},
      {"AB", b},
      {"_A", a},
      {"AAA", a},
      {"ABB", b},
      {"_AA", a},
      {"AAAA", a},
      {"AABA", a},
      {"AA_A", a},
      {"AAcA", a},
      {"BAAA", a},
      {"AAAB", none},
      {"AABB", none},
      {"A__A", none},
      {"A___", none},
      {"___A", none},
      {"Accc", none},
      {"AAcc", none},
      {"ABBB", b},
      {"_AAA", a},
      {"AABBB", none},
      {"ABBBB", b},
      {"_AAAA", a},
      {"AAABBB", none},
      {"AABBBB", b},
      {std::string(5, 'A') + std::string(3, 'B'), none},
      {std::string(4, 'A') + std::string(4, 'B'), b},
      {std::string(9, 'A') + "B", none},
      {std::string(9, 'A') + "_", none},
      {std::string(9, 'A') + "c", none},
      {std::string(5, 'A') + std::string(5, 'B'), b},
      {std::string(4, 'A') + std::string(6, 'B'), b},
      {std::string(6, 'A') + std::string(4, 'B'), b},
      {std::string(4, '_') + std::string(6, 'B'), b},
      {std::string(6, '_') + std::string(4, 'B'), b},
      {std::string(7, 'A') + std::string(3, 'B'), none},
      {std::string(7, '_') + std::string(3, 'B'), none},
      {std::string(7, 'c') + std::string(3, 'B'), none},
      {std::string(8, 'A') + "_A", a},
      {std::string(6, 'A') + "_AAA", a},
      {std::string(5, '_') + "A_AAA", a},
      {std::string(7, 'A') + "__A", none},
      {std::string(6, 'A') + "___A", none},
      {std::string(5, 'A') + "__AAA", none},
      {std::string(5, 'A') + "_A_AA", none},
      {std::string(6, '_') + "AA_A", none},
      {std::string(6, 'B') + "_AAA", none},
      {std::string(5, 'A') + "B_AAA", none},
      {std::string(6, 'A') + "_BBB", none},
      {"BABABABABB", none},
      {"ABABABABAB", none},
      {std::string(16, 'A') + std::string(4, 'B'), b},
      {std::string(17, 'A') + std::string(3, 'B'), none},
  };
}

TEST(TableCalcColumnWindowRule, CreationMarksColumnsAsExcelDoes) {
  for (const WindowCase& c : WindowCases()) {
    Workbook wb = PatternBook(c.pattern);
    ASSERT_TRUE(wb.add_table(MakeTable(PatternRef(c.pattern))).has_value()) << c.pattern;
    EXPECT_EQ(Mark(wb, 2U), c.mark) << c.pattern;
    EXPECT_EQ(infer_calculated_column(wb.sheet(0U), wb.tables()[0], 2U).value_or(std::string()), c.mark) << c.pattern;
    // Constant columns take no mark.
    EXPECT_EQ(Mark(wb, 0U), "") << c.pattern;
    EXPECT_EQ(Mark(wb, 1U), "") << c.pattern;
  }
}

TEST(TableCalcColumnWindowRule, CreationLeavesMisfitCellsAsTheyAre) {
  Workbook wb = PatternBook("AABA");
  AddTable(wb, PatternRef("AABA"));
  EXPECT_EQ(Mark(wb, 2U), "A3*10");
  EXPECT_EQ(FormulaAt(wb, "C5"), "=A5*20");
  EXPECT_TRUE(wb.sheet(0U).shared_formula_groups().empty());
}

TEST(TableCalcColumnWindowRule, AbsoluteFormulaIsMarkedVerbatim) {
  Workbook wb = PatternBook("____");
  for (const char* cell : {"C3", "C4", "C5", "C6"}) {
    SetFormula(wb, cell, "=$A$3*10");
  }
  AddTable(wb, "A2:C6");
  EXPECT_EQ(Mark(wb, 2U), "$A$3*10");
}

TEST(TableCalcColumnWindowRule, NoDataRowAndOutOfRangeColumnInferNothing) {
  Workbook wb = PatternBook("");
  AddTable(wb, "A2:C2");
  EXPECT_FALSE(infer_calculated_column(wb.sheet(0U), wb.tables()[0], 2U).has_value());
  EXPECT_FALSE(infer_calculated_column(wb.sheet(0U), wb.tables()[0], 3U).has_value());
}

// ---------------------------------------------------------------------------
// Set / clear / drop / infer
// ---------------------------------------------------------------------------

TEST(TableCalcColumnSet, FillsEveryDataCellAndMarksTheFirstRowSpelling) {
  Workbook wb = PatternBook("c_cA");
  AddTable(wb, "A2:C6");
  ASSERT_TRUE(wb.set_table_calculated_column(0U, 0U, 2U, "=A3*10").has_value());
  EXPECT_EQ(Mark(wb, 2U), "A3*10");
  for (int r = 3; r <= 6; ++r) {
    const std::string row = std::to_string(r);
    EXPECT_EQ(FormulaAt(wb, "C" + row), "=A" + row + "*10") << r;
  }
  ASSERT_EQ(wb.sheet(0U).shared_formula_groups().size(), 1U);
  EXPECT_EQ(wb.sheet(0U).shared_formula_groups()[0].ref, (MergeRange{2U, 2U, 5U, 2U}));

  // Without '=', and a second set replaces the group rather than adding one.
  ASSERT_TRUE(wb.set_table_calculated_column(0U, 0U, 2U, "A3*20").has_value());
  EXPECT_EQ(Mark(wb, 2U), "A3*20");
  EXPECT_EQ(FormulaAt(wb, "C6"), "=A6*20");
  EXPECT_EQ(wb.sheet(0U).shared_formula_groups().size(), 1U);
}

TEST(TableCalcColumnSet, TotalsRowIsNotFilled) {
  Workbook wb = PatternBook("cccc");
  SetText(wb, "C7", "Total");
  AddTable(wb, "A2:C7", /*totals=*/true);
  ASSERT_TRUE(wb.set_table_calculated_column(0U, 0U, 2U, "=A3*10").has_value());
  EXPECT_EQ(FormulaAt(wb, "C6"), "=A6*10");
  EXPECT_EQ(FormulaAt(wb, "C7"), "");
  EXPECT_EQ(wb.sheet(0U).resolve_cell_value(6U, 2U).as_text(), "Total");
  EXPECT_EQ(wb.sheet(0U).shared_formula_groups()[0].ref, (MergeRange{2U, 2U, 5U, 2U}));
}

TEST(TableCalcColumnSet, RefusalsChangeNothing) {
  Workbook wb = PatternBook("cccc");
  AddTable(wb, "A2:C6");
  ExpectCode(wb.set_table_calculated_column(0U, 0U, 2U, ""), FormulonErrorCode::kInvalidArgument);
  ExpectCode(wb.set_table_calculated_column(0U, 0U, 2U, "="), FormulonErrorCode::kInvalidArgument);
  ExpectCode(wb.set_table_calculated_column(0U, 0U, 2U, "=SEQUENCE(3)"), FormulonErrorCode::kInvalidArgument);
  // The last data row would read A1048577.
  ExpectCode(wb.set_table_calculated_column(0U, 0U, 2U, "=A1048574"), FormulonErrorCode::kInvalidArgument);
  ExpectCode(wb.set_table_calculated_column(0U, 0U, 3U, "=A3"), FormulonErrorCode::kInvalidArgument);
  ExpectCode(wb.set_table_calculated_column(0U, 1U, 2U, "=A3"), FormulonErrorCode::kInvalidArgument);
  ExpectCode(wb.set_table_calculated_column(1U, 0U, 2U, "=A3"), FormulonErrorCode::kInvalidArgument);
  const Expected<void, Error> unparsable = wb.set_table_calculated_column(0U, 0U, 2U, "=A3*");
  ASSERT_FALSE(unparsable.has_value());
  EXPECT_GE(static_cast<int>(unparsable.error().code), static_cast<int>(FormulonErrorCode::kParserUnexpectedToken));
  EXPECT_LE(static_cast<int>(unparsable.error().code), static_cast<int>(FormulonErrorCode::kParserTooManyErrors));
  EXPECT_EQ(Mark(wb, 2U), "");
  EXPECT_EQ(FormulaAt(wb, "C3"), "");
  EXPECT_TRUE(wb.sheet(0U).shared_formula_groups().empty());
}

TEST(TableCalcColumnSet, ZeroDataRowsSetsOnlyTheMark) {
  Workbook wb = PatternBook("");
  AddTable(wb, "A2:C2");
  ASSERT_TRUE(wb.set_table_calculated_column(0U, 0U, 2U, "=A3*10").has_value());
  EXPECT_EQ(Mark(wb, 2U), "A3*10");
  EXPECT_EQ(FormulaAt(wb, "C3"), "");
  EXPECT_TRUE(wb.sheet(0U).shared_formula_groups().empty());
}

TEST(TableCalcColumnSet, StructuredReferenceIsFilledButNotShared) {
  Workbook wb = PatternBook("cccc");
  AddTable(wb, "A2:C6");
  ASSERT_TRUE(wb.set_table_calculated_column(0U, 0U, 2U, "=[@a]*10").has_value());
  EXPECT_NE(Mark(wb, 2U).find("Tbl["), std::string::npos) << Mark(wb, 2U);
  EXPECT_NE(FormulaAt(wb, "C6"), "");
  EXPECT_TRUE(wb.sheet(0U).shared_formula_groups().empty());
  EXPECT_EQ(CountOf(SheetXml(wb), "t=\"shared\""), 0U);
}

TEST(TableCalcColumnSet, VolatileFormulaIsNotShared) {
  Workbook wb = PatternBook("cccc");
  AddTable(wb, "A2:C6");
  ASSERT_TRUE(wb.set_table_calculated_column(0U, 0U, 2U, "=A3+RAND()").has_value());
  EXPECT_EQ(FormulaAt(wb, "C4"), "=A4+RAND()");
  EXPECT_TRUE(wb.sheet(0U).shared_formula_groups().empty());
}

TEST(TableCalcColumnClear, RemovesTheMarkAndKeepsTheCells) {
  Workbook wb = PatternBook("cccc");
  AddTable(wb, "A2:C6");
  ASSERT_TRUE(wb.set_table_calculated_column(0U, 0U, 2U, "=A3*10").has_value());
  ASSERT_TRUE(wb.clear_table_calculated_column(0U, 0U, 2U).has_value());
  EXPECT_EQ(Mark(wb, 2U), "");
  EXPECT_EQ(FormulaAt(wb, "C4"), "=A4*10");
  ExpectCode(wb.clear_table_calculated_column(0U, 0U, 3U), FormulonErrorCode::kInvalidArgument);
}

TEST(TableCalcColumnDrop, OnlyAnAllBlankColumnLosesItsMark) {
  Workbook wb = PatternBook("cccc");
  AddTable(wb, "A2:C6");
  ASSERT_TRUE(wb.set_table_calculated_column(0U, 0U, 2U, "=A3*10").has_value());
  for (const char* cell : {"C3", "C4", "C5"}) {
    SetBlank(wb, cell);
  }
  ASSERT_TRUE(wb.drop_calculated_column_if_empty(0U, 0U, 2U).has_value());
  EXPECT_EQ(Mark(wb, 2U), "A3*10");
  SetBlank(wb, "C6");
  ASSERT_TRUE(wb.drop_calculated_column_if_empty(0U, 0U, 2U).has_value());
  EXPECT_EQ(Mark(wb, 2U), "");
}

TEST(TableCalcColumnInferTable, ResizeKeepsMarkedColumnsAndFillsOnlyTheBlanksOfAddedRows) {
  Workbook wb = PatternBook("cc");
  AddTable(wb, "A2:C4");
  ASSERT_TRUE(wb.set_table_calculated_column(0U, 0U, 2U, "=A3*10").has_value());
  SetBlank(wb, "C4");
  for (int r = 3; r <= 6; ++r) {
    const std::string row = std::to_string(r);
    SetFormula(wb, "B" + row, "=A" + row + "+1");
  }
  SetNumber(wb, "C5", 99);
  wb.mutable_tables()[0].ref = "A2:C6";
  ASSERT_TRUE(wb.infer_table_calculated_columns(0U, 0U, /*first_new_row=*/4U).has_value());
  EXPECT_EQ(Mark(wb, 0U), "");
  EXPECT_EQ(Mark(wb, 1U), "A3+1");
  EXPECT_EQ(Mark(wb, 2U), "A3*10");
  EXPECT_EQ(FormulaAt(wb, "C4"), "");
  EXPECT_EQ(FormulaAt(wb, "C5"), "");
  EXPECT_EQ(wb.sheet(0U).resolve_cell_value(4U, 2U).as_number(), 99.0);
  EXPECT_EQ(FormulaAt(wb, "C6"), "=A6*10");
  ExpectCode(wb.infer_table_calculated_columns(0U, 1U, 0U), FormulonErrorCode::kInvalidArgument);
}

// ---------------------------------------------------------------------------
// Shared-formula groups
// ---------------------------------------------------------------------------

constexpr std::string_view kMaster = "<f t=\"shared\" ref=\"C3:C6\" si=\"0\">A3*10</f>";
constexpr std::string_view kMember = "<f t=\"shared\" si=\"0\"/>";

Workbook SharedBook() {
  Workbook wb = PatternBook("cccc");
  AddTable(wb, "A2:C6");
  EXPECT_TRUE(wb.set_table_calculated_column(0U, 0U, 2U, "=A3*10").has_value());
  return wb;
}

TEST(TableCalcColumnSharedXml, SetColumnIsWrittenAsOneSharedFormula) {
  const Workbook wb = SharedBook();
  const std::string xml = SheetXml(wb);
  EXPECT_NE(xml.find(std::string("<c r=\"C3\">") + std::string(kMaster)), std::string::npos) << xml;
  EXPECT_NE(xml.find(std::string("<c r=\"C4\">") + std::string(kMember) + "</c>"), std::string::npos) << xml;
  EXPECT_EQ(CountOf(xml, kMember), 3U) << xml;
}

TEST(TableCalcColumnSharedXml, SameFormulaReenteredStaysShared) {
  Workbook wb = SharedBook();
  SetFormula(wb, "C5", "=A5*10");
  EXPECT_EQ(CountOf(SheetXml(wb), kMember), 3U);
}

TEST(TableCalcColumnSharedXml, DifferentFormulaIsWrittenOnItsOwn) {
  Workbook wb = SharedBook();
  SetFormula(wb, "C5", "=A5*20");
  const std::string xml = SheetXml(wb);
  EXPECT_NE(xml.find("<c r=\"C5\"><f>A5*20</f></c>"), std::string::npos) << xml;
  EXPECT_NE(xml.find(kMaster), std::string::npos) << xml;
  EXPECT_EQ(CountOf(xml, kMember), 2U) << xml;
}

TEST(TableCalcColumnSharedXml, ConstantMasterWritesTheGroupIndividually) {
  Workbook wb = SharedBook();
  SetNumber(wb, "C3", 1);
  const std::string xml = SheetXml(wb);
  EXPECT_EQ(CountOf(xml, "t=\"shared\""), 0U) << xml;
  EXPECT_NE(xml.find("<c r=\"C4\"><f>A4*10</f></c>"), std::string::npos) << xml;
}

TEST(TableCalcColumnSharedXml, SavedBookReadsBackTheSameFormulas) {
  const Workbook wb = SharedBook();
  const Expected<std::vector<std::uint8_t>, Error> saved = io::write_ooxml(wb);
  ASSERT_TRUE(saved.has_value());
  auto loaded = io::read_ooxml(io::ByteSpan{saved.value().data(), saved.value().size()});
  ASSERT_TRUE(loaded.has_value());
  const Workbook& back = loaded.value().workbook;
  for (int r = 3; r <= 6; ++r) {
    const std::string row = std::to_string(r);
    EXPECT_EQ(FormulaAt(back, "C" + row), "=A" + row + "*10") << r;
  }
  ASSERT_EQ(back.tables().size(), 1U);
  EXPECT_EQ(Mark(back, 2U), "A3*10");
}

TEST(TableCalcColumnSharedGroup, RowInsertGrowsTheGroupAndTheNewRowJoinsIt) {
  Workbook wb = SharedBook();
  ASSERT_TRUE(wb.insert_rows(0U, 4U, 1U).has_value());
  ASSERT_EQ(wb.sheet(0U).shared_formula_groups().size(), 1U);
  EXPECT_EQ(wb.sheet(0U).shared_formula_groups()[0].ref, (MergeRange{2U, 2U, 6U, 2U}));
  const std::string xml = SheetXml(wb);
  EXPECT_NE(xml.find("<f t=\"shared\" ref=\"C3:C7\" si=\"0\">A3*10</f>"), std::string::npos) << xml;
  EXPECT_EQ(CountOf(xml, kMember), 4U) << xml;
}

TEST(TableCalcColumnSharedGroup, RowDeleteShrinksTheGroup) {
  Workbook wb = SharedBook();
  ASSERT_TRUE(wb.delete_rows(0U, 3U, 1U).has_value());
  ASSERT_EQ(wb.sheet(0U).shared_formula_groups().size(), 1U);
  EXPECT_EQ(wb.sheet(0U).shared_formula_groups()[0].ref, (MergeRange{2U, 2U, 4U, 2U}));
  EXPECT_EQ(CountOf(SheetXml(wb), kMember), 2U);
}

TEST(TableCalcColumnSharedGroup, DeletingEveryRowOfAGroupDropsIt) {
  Workbook wb = Workbook::create();
  SetFormula(wb, "E3", "=A3");
  SetFormula(wb, "E4", "=A4");
  wb.sheet(0U).mutable_shared_formula_groups().push_back(SharedFormulaGroup{MergeRange{2U, 4U, 3U, 4U}});
  ASSERT_TRUE(wb.delete_rows(0U, 2U, 2U).has_value());
  EXPECT_TRUE(wb.sheet(0U).shared_formula_groups().empty());
}

TEST(TableCalcColumnSharedGroup, CopySheetCopiesTheGroup) {
  Workbook wb = SharedBook();
  ASSERT_TRUE(wb.copy_sheet(0U, 1U, std::nullopt).has_value());
  ASSERT_EQ(wb.sheet(1U).shared_formula_groups().size(), 1U);
  EXPECT_EQ(wb.sheet(1U).shared_formula_groups()[0].ref, (MergeRange{2U, 2U, 5U, 2U}));
  EXPECT_EQ(CountOf(SheetXml(wb, 1U), kMember), 3U);
}

}  // namespace
}  // namespace formulon
