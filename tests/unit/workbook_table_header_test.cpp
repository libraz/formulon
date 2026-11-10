// Turning a table's header row off and on.

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "auto_filter.h"
#include "cell.h"
#include "cf/auto_filter_eval.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/auto_filter_xml.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/zip_reader.h"
#include "merge_range.h"
#include "sheet.h"
#include "styles.h"
#include "table.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

MergeRange Rect(const char* a1) {
  const std::optional<MergeRange> rect = parse_a1_rectangle(a1);
  EXPECT_TRUE(rect.has_value()) << a1;
  return rect.value_or(MergeRange{});
}

Workbook Book() {
  ExcelProfile profile;
  EXPECT_TRUE(parse_excel_profile_id("mac-365-ja_JP", &profile));
  return Workbook::create_for_profile(profile);
}

void Recalc(Workbook& wb) {
  ASSERT_TRUE(wb.recalc(eval::default_registry()));
}

Value At(const Workbook& wb, const char* a1) {
  const MergeRange r = Rect(a1);
  return wb.sheet(0).resolve_cell_value(r.first_row, r.first_col);
}

std::string TextAt(const Workbook& wb, const char* a1) {
  const Value v = At(wb, a1);
  return v.is_text() ? std::string(v.as_text()) : std::string();
}

const Cell* CellAt(const Workbook& wb, const char* a1) {
  const MergeRange r = Rect(a1);
  return wb.sheet(0).cell_at(r.first_row, r.first_col);
}

void PutText(Workbook& wb, const char* a1, const std::string& text) {
  const MergeRange r = Rect(a1);
  ASSERT_TRUE(wb.set_cell_text(0, r.first_row, r.first_col, text).has_value()) << a1;
}

void PutNumber(Workbook& wb, const char* a1, double number) {
  const MergeRange r = Rect(a1);
  ASSERT_TRUE(wb.set_cell_value(0, r.first_row, r.first_col, Value::number(number)).has_value()) << a1;
}

void PutFormula(Workbook& wb, const char* a1, const std::string& formula) {
  const MergeRange r = Rect(a1);
  ASSERT_TRUE(wb.set_cell_formula(0, r.first_row, r.first_col, formula).has_value()) << a1;
}

void Merge(Workbook& wb, const char* a1) {
  ASSERT_TRUE(wb.add_merge(0, Rect(a1)).has_value()) << a1;
}

bool HasMerge(const Workbook& wb, const char* a1) {
  const MergeRange want = Rect(a1);
  for (const MergeRange& merge : wb.sheet(0).merges()) {
    if (merge == want) {
      return true;
    }
  }
  return false;
}

void ExpectCode(const Expected<void, Error>& result, FormulonErrorCode code) {
  ASSERT_FALSE(result.has_value());
  EXPECT_EQ(result.error().code, code) << result.error().message;
}

std::string FilterXml(const Workbook& wb) {
  return io::auto_filter_xml(wb.tables()[0].auto_filter_xml.get());
}

// Table "S" over `ref` (three columns) with Region / Sales / Units headers and
// a bold-like cell format on the header cells; data rows hold East.. / numbers.
void AddS(Workbook& wb, const char* ref) {
  const MergeRange r = Rect(ref);
  const char* headers[] = {"Region", "Sales", "Units"};
  StylesTable& styles = wb.mutable_styles();
  if (styles.cell_xfs.empty()) {
    styles.cell_xfs.emplace_back();
  }
  styles.cell_xfs.emplace_back();
  const std::uint32_t xf = static_cast<std::uint32_t>(styles.cell_xfs.size() - 1U);
  for (std::uint32_t i = 0; i < 3U; ++i) {
    ASSERT_TRUE(wb.set_cell_text(0, r.first_row, r.first_col + i, headers[i]).has_value());
    ASSERT_TRUE(wb.set_cell_xf_index(0, r.first_row, r.first_col + i, xf).has_value());
  }
  for (std::uint32_t row = r.first_row + 1U; row <= r.last_row; ++row) {
    const std::uint32_t n = row - r.first_row;
    ASSERT_TRUE(wb.set_cell_text(0, row, r.first_col, "R" + std::to_string(n)).has_value());
    ASSERT_TRUE(wb.set_cell_value(0, row, r.first_col + 1U, Value::number(10.0 * n)).has_value());
    ASSERT_TRUE(wb.set_cell_value(0, row, r.first_col + 2U, Value::number(n)).has_value());
  }
  TableMetadata table;
  table.name = "S";
  table.display_name = "S";
  table.ref = ref;
  for (const char* header : headers) {
    TableColumn column;
    column.name = header;
    table.columns.push_back(std::move(column));
  }
  const Expected<std::size_t, Error> added = wb.add_table(std::move(table));
  ASSERT_TRUE(added.has_value()) << added.error().message;
}

// The four header-reference formulas of the measured setup, in F1:F4.
void AddHeaderFormulas(Workbook& wb) {
  PutFormula(wb, "F1", "=S[#Headers]");
  PutFormula(wb, "F2", "=S[[#Headers],[Sales]]");
  PutFormula(wb, "F3", "=ROWS(S[#Headers])");
  PutFormula(wb, "F4", "=ISERROR(S[#Headers])");
}

void ExpectHeaderFormulasResolve(const Workbook& wb) {
  EXPECT_EQ(TextAt(wb, "F1"), "Region");
  EXPECT_EQ(TextAt(wb, "G1"), "Sales") << "the header array spills";
  EXPECT_EQ(TextAt(wb, "H1"), "Units");
  EXPECT_EQ(TextAt(wb, "F2"), "Sales");
  EXPECT_EQ(At(wb, "F3"), Value::number(1.0));
  EXPECT_EQ(At(wb, "F4"), Value::boolean(false));
  EXPECT_EQ(At(wb, "H4"), Value::boolean(false));
}

void ExpectHeaderFormulasRefError(const Workbook& wb) {
  for (const char* a1 : {"F1", "F2", "F3"}) {
    const Value v = At(wb, a1);
    ASSERT_TRUE(v.is_error()) << a1;
    EXPECT_EQ(v.as_error(), ErrorCode::Ref) << a1;
  }
  EXPECT_TRUE(At(wb, "G1").is_blank()) << "the #REF! does not spill";
  EXPECT_EQ(At(wb, "F4"), Value::boolean(true));
  EXPECT_TRUE(At(wb, "G4").is_blank()) << "ISERROR is a single TRUE";
}

std::vector<std::uint8_t> SaveXlsx(const Workbook& wb) {
  auto saved = io::write_ooxml(wb);
  EXPECT_TRUE(saved.has_value()) << (saved ? "" : saved.error().message);
  return saved ? std::move(saved.value()) : std::vector<std::uint8_t>();
}

Workbook LoadXlsx(const std::vector<std::uint8_t>& bytes) {
  auto loaded = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(loaded.has_value()) << (loaded ? "" : loaded.error().message);
  return loaded ? std::move(loaded.value().workbook) : Workbook::create();
}

std::string Entry(const std::vector<std::uint8_t>& package, const std::string& prefix) {
  io::ZipReader zip;
  EXPECT_TRUE(zip.open(io::ByteSpan{package.data(), package.size()}).has_value());
  for (const std::string& name : zip.list_entries()) {
    if (name.rfind(prefix, 0) == 0) {
      auto bytes = zip.read_entry(name);
      EXPECT_TRUE(bytes.has_value());
      return bytes ? std::string(bytes.value().begin(), bytes.value().end()) : std::string();
    }
  }
  ADD_FAILURE() << "no part " << prefix;
  return {};
}

TEST(WorkbookTableHeaderToggle, HidingDeletesTheHeaderCellsAndShrinksTheRangeWithoutShifting) {
  Workbook wb = Book();
  AddS(wb, "B2:D5");
  AddHeaderFormulas(wb);
  Recalc(wb);
  ExpectHeaderFormulasResolve(wb);
  ASSERT_FALSE(FilterXml(wb).empty());

  ASSERT_TRUE(wb.set_table_header_row(0, false).has_value());
  const TableMetadata& table = wb.tables()[0];
  EXPECT_EQ(table.ref, "B3:D5");
  EXPECT_FALSE(table.header_row);
  EXPECT_TRUE(table.auto_filter_xml.empty());
  ASSERT_EQ(table.columns.size(), 3U);
  EXPECT_EQ(table.columns[1].name, "Sales") << "the column names are kept";
  for (const char* a1 : {"B2", "C2", "D2"}) {
    EXPECT_EQ(CellAt(wb, a1), nullptr) << a1 << ": value and format are gone";
  }
  EXPECT_EQ(TextAt(wb, "B3"), "R1") << "the data stays where it was";
  EXPECT_EQ(At(wb, "C5"), Value::number(30.0));
  Recalc(wb);
  ExpectHeaderFormulasRefError(wb);
  EXPECT_EQ(CellAt(wb, "F1")->formula_text, "=S[#Headers]") << "formula text is never rewritten";

  // A second hide changes nothing.
  ASSERT_TRUE(wb.set_table_header_row(0, false).has_value());
  EXPECT_EQ(wb.tables()[0].ref, "B3:D5");
}

bool RowHidden(const Workbook& wb, std::uint32_t row) {
  for (const RowLayout& r : wb.sheet(0).layout().row_overrides) {
    if (r.row == row) {
      return r.hidden;
    }
  }
  return false;
}

TEST(WorkbookTableHeaderToggle, HidingDropsTheFilterCriteriaAndShowsTheRowsTheyHid) {
  Workbook wb = Book();
  AddS(wb, "B3:D8");
  const double sales[] = {50.0, 150.0, 80.0, 200.0, 120.0};
  for (std::uint32_t i = 0; i < 5U; ++i) {
    ASSERT_TRUE(wb.set_cell_value(0, 3U + i, 2U, Value::number(sales[i])).has_value());
  }
  const AutoFilter filter = io::auto_filter_from_xml(
      "<autoFilter ref=\"B3:D8\"><filterColumn colId=\"1\"><customFilters>"
      "<customFilter operator=\"greaterThan\" val=\"100\"/></customFilters></filterColumn></autoFilter>");
  ASSERT_TRUE(wb.set_table_auto_filter(0, filter).has_value());
  ASSERT_TRUE(apply_auto_filter(wb, 0, filter).has_value());
  PutFormula(wb, "F1", "=SUBTOTAL(103,B4:B8)");
  Recalc(wb);
  ASSERT_TRUE(RowHidden(wb, 3U));
  ASSERT_TRUE(RowHidden(wb, 5U));
  ASSERT_EQ(At(wb, "F1"), Value::number(3.0));

  ASSERT_TRUE(wb.set_table_header_row(0, false).has_value());
  EXPECT_EQ(wb.tables()[0].ref, "B4:D8");
  EXPECT_TRUE(wb.tables()[0].auto_filter_xml.empty());
  for (std::uint32_t row = 3U; row <= 7U; ++row) {
    EXPECT_FALSE(RowHidden(wb, row)) << "row " << row + 1U;
  }
  Recalc(wb);
  EXPECT_EQ(At(wb, "F1"), Value::number(5.0));

  // The criteria do not come back with the header.
  ASSERT_TRUE(wb.set_table_header_row(0, true).has_value());
  EXPECT_EQ(wb.tables()[0].ref, "B3:D8");
  EXPECT_EQ(FilterXml(wb), "<autoFilter ref=\"B3:D8\"/>");
  EXPECT_FALSE(RowHidden(wb, 3U));
}

TEST(WorkbookTableHeaderToggle, HidingWithoutFilterCriteriaKeepsRowsHiddenByHand) {
  Workbook wb = Book();
  AddS(wb, "B3:D8");
  RowLayout hidden;
  hidden.row = 4U;
  hidden.hidden = true;
  wb.sheet(0).mutable_layout().row_overrides.push_back(hidden);
  ASSERT_TRUE(wb.set_table_header_row(0, false).has_value());
  EXPECT_TRUE(RowHidden(wb, 4U));
}

TEST(WorkbookTableHeaderToggle, ShowingIntoAnEmptyRowAboveGrowsTheRangeUp) {
  Workbook wb = Book();
  AddS(wb, "B2:D5");
  AddHeaderFormulas(wb);
  ASSERT_TRUE(wb.set_table_header_row(0, false).has_value());
  ASSERT_TRUE(wb.set_table_header_row(0, true).has_value());
  const TableMetadata& table = wb.tables()[0];
  EXPECT_EQ(table.ref, "B2:D5");
  EXPECT_TRUE(table.header_row);
  EXPECT_EQ(TextAt(wb, "B2"), "Region");
  EXPECT_EQ(TextAt(wb, "C2"), "Sales");
  EXPECT_EQ(TextAt(wb, "D2"), "Units");
  EXPECT_EQ(TextAt(wb, "B3"), "R1") << "nothing shifted";
  EXPECT_EQ(FilterXml(wb), "<autoFilter ref=\"B2:D5\"/>");
  Recalc(wb);
  ExpectHeaderFormulasResolve(wb);
}

TEST(WorkbookTableHeaderToggle, ShowingUnderAnOccupiedRowInsertsCellsWithinTheTableColumns) {
  Workbook wb = Book();
  AddS(wb, "B2:D5");
  AddHeaderFormulas(wb);
  PutText(wb, "A4", "left");
  PutText(wb, "E6", "right");
  ASSERT_TRUE(wb.set_table_header_row(0, false).has_value());
  PutText(wb, "C2", "X");
  ASSERT_TRUE(wb.set_table_header_row(0, true).has_value());
  const TableMetadata& table = wb.tables()[0];
  EXPECT_EQ(table.ref, "B3:D6");
  EXPECT_TRUE(table.header_row);
  EXPECT_EQ(TextAt(wb, "C2"), "X") << "the occupied cell stays put";
  EXPECT_EQ(TextAt(wb, "B3"), "Region");
  EXPECT_EQ(TextAt(wb, "C3"), "Sales");
  EXPECT_EQ(TextAt(wb, "B4"), "R1") << "the data rows moved down one";
  EXPECT_EQ(At(wb, "C6"), Value::number(30.0));
  EXPECT_EQ(TextAt(wb, "A4"), "left") << "cells outside the table's columns do not move";
  EXPECT_EQ(TextAt(wb, "E6"), "right");
  EXPECT_EQ(CellAt(wb, "F1")->formula_text, "=S[#Headers]");
  EXPECT_EQ(FilterXml(wb), "<autoFilter ref=\"B3:D6\"/>");
  Recalc(wb);
  ExpectHeaderFormulasResolve(wb);
}

TEST(WorkbookTableHeaderToggle, ATableOnRowOneGrowsBackOrShiftsByItsRowAbove) {
  {
    Workbook wb = Book();
    AddS(wb, "A1:C4");
    ASSERT_TRUE(wb.set_table_header_row(0, false).has_value());
    EXPECT_EQ(wb.tables()[0].ref, "A2:C4");
    EXPECT_EQ(CellAt(wb, "A1"), nullptr);
    ASSERT_TRUE(wb.set_table_header_row(0, true).has_value());
    EXPECT_EQ(wb.tables()[0].ref, "A1:C4");
    EXPECT_EQ(TextAt(wb, "A1"), "Region");
    EXPECT_EQ(TextAt(wb, "A2"), "R1");
  }
  {
    Workbook wb = Book();
    AddS(wb, "A1:C4");
    ASSERT_TRUE(wb.set_table_header_row(0, false).has_value());
    PutText(wb, "B1", "X");
    ASSERT_TRUE(wb.set_table_header_row(0, true).has_value());
    EXPECT_EQ(wb.tables()[0].ref, "A2:C5");
    EXPECT_EQ(TextAt(wb, "B1"), "X");
    EXPECT_EQ(TextAt(wb, "A2"), "Region");
    EXPECT_EQ(TextAt(wb, "A3"), "R1");
    EXPECT_EQ(At(wb, "B5"), Value::number(30.0));
  }
}

TEST(WorkbookTableHeaderToggle, ATableWithoutAHeaderOnTheFirstGridRowShiftsToShowIt) {
  // A table created without a header on the grid's first row has no row above it.
  Workbook top = Book();
  PutText(top, "A1", "R1");
  PutNumber(top, "B1", 10.0);
  PutNumber(top, "C1", 1.0);
  TableMetadata table;
  table.name = "T";
  table.display_name = "T";
  table.ref = "A1:C1";
  table.header_row = false;
  for (const char* name : {"Region", "Sales", "Units"}) {
    TableColumn column;
    column.name = name;
    table.columns.push_back(std::move(column));
  }
  ASSERT_TRUE(top.add_table(std::move(table)).has_value());
  ASSERT_TRUE(top.set_table_header_row(0, true).has_value());
  EXPECT_EQ(top.tables()[0].ref, "A1:C2");
  EXPECT_EQ(TextAt(top, "A1"), "Region");
  EXPECT_EQ(TextAt(top, "A2"), "R1");
}

TEST(WorkbookTableHeaderToggle, AGrownRangeOverlappingAnotherStructureShiftsInstead) {
  Workbook wb = Book();
  AddS(wb, "B2:D5");
  ASSERT_TRUE(wb.set_table_header_row(0, false).has_value());
  // The blank row above is covered by a merge reaching into the table's columns.
  Merge(wb, "A2:B2");
  ASSERT_TRUE(wb.set_table_header_row(0, true).has_value());
  EXPECT_EQ(wb.tables()[0].ref, "B3:D6");
  EXPECT_TRUE(HasMerge(wb, "A2:B2")) << "the merge above the band stays";
  EXPECT_EQ(TextAt(wb, "B3"), "Region");
}

TEST(WorkbookTableHeaderToggle, AShowingShiftUnmergesStraddlingMergesAndMovesInnerOnes) {
  Workbook wb = Book();
  PutText(wb, "A6", "a6");
  PutText(wb, "D6", "d6");
  PutText(wb, "G4", "g4");
  AddS(wb, "B2:D5");
  ASSERT_TRUE(wb.set_table_header_row(0, false).has_value());
  PutText(wb, "C2", "X");
  Merge(wb, "A6:B6");
  Merge(wb, "D6:E6");
  Merge(wb, "B7:C7");
  Merge(wb, "A9:E9");
  Merge(wb, "E4:F4");
  Merge(wb, "A3:A4");
  ASSERT_TRUE(wb.set_table_header_row(0, true).has_value());
  EXPECT_EQ(wb.tables()[0].ref, "B3:D6");
  EXPECT_FALSE(HasMerge(wb, "A6:B6")) << "across the left edge: unmerged";
  EXPECT_FALSE(HasMerge(wb, "A7:B7"));
  EXPECT_FALSE(HasMerge(wb, "D6:E6")) << "across the right edge: unmerged";
  EXPECT_FALSE(HasMerge(wb, "D7:E7"));
  EXPECT_FALSE(HasMerge(wb, "A9:E9")) << "across both edges: unmerged";
  EXPECT_FALSE(HasMerge(wb, "A10:E10"));
  EXPECT_TRUE(HasMerge(wb, "B8:C8")) << "wholly inside the columns: moved";
  EXPECT_FALSE(HasMerge(wb, "B7:C7"));
  EXPECT_TRUE(HasMerge(wb, "E4:F4")) << "beside the table: unchanged";
  EXPECT_TRUE(HasMerge(wb, "A3:A4"));
  EXPECT_EQ(TextAt(wb, "A6"), "a6");
  EXPECT_EQ(TextAt(wb, "D7"), "d6");
  EXPECT_EQ(TextAt(wb, "G4"), "g4");
}

TEST(WorkbookTableHeaderToggle, ARefusedShiftOrAnEmptyTableChangesNothing) {
  {
    Workbook wb = Book();
    AddS(wb, "B2:D5");
    ASSERT_TRUE(wb.set_table_header_row(0, false).has_value());
    PutText(wb, "C2", "X");
    PutText(wb, "C1048576", "edge");
    ExpectCode(wb.set_table_header_row(0, true), FormulonErrorCode::kPreconditionFailed);
    EXPECT_EQ(wb.tables()[0].ref, "B3:D5");
    EXPECT_FALSE(wb.tables()[0].header_row);
    EXPECT_TRUE(wb.tables()[0].auto_filter_xml.empty());
    EXPECT_EQ(TextAt(wb, "B3"), "R1");
    EXPECT_EQ(TextAt(wb, "C1048576"), "edge");
  }
  {
    Workbook wb = Book();
    AddS(wb, "B2:D5");
    ASSERT_TRUE(wb.set_table_header_row(0, false).has_value());
    PutText(wb, "C2", "X");
    // Another table below the data straddles the band's left edge.
    PutText(wb, "A8", "J");
    PutText(wb, "B8", "K");
    PutText(wb, "C8", "L");
    TableMetadata other;
    other.name = "U";
    other.display_name = "U";
    other.ref = "A8:C9";
    for (const char* name : {"J", "K", "L"}) {
      TableColumn column;
      column.name = name;
      other.columns.push_back(std::move(column));
    }
    ASSERT_TRUE(wb.add_table(std::move(other)).has_value());
    ExpectCode(wb.set_table_header_row(0, true), FormulonErrorCode::kPreconditionFailed);
    EXPECT_EQ(wb.tables()[0].ref, "B3:D5");
    EXPECT_EQ(wb.tables()[1].ref, "A8:C9");
    EXPECT_EQ(TextAt(wb, "B3"), "R1");
  }
  {
    // Header plus totals with no data row between them.
    Workbook wb = Book();
    AddS(wb, "B2:D5");
    const TableShape empty{Rect("B2:D3"), true, true};
    const auto checked = wb.check_table_header_row(0, false, empty, {});
    ASSERT_FALSE(checked.has_value());
    EXPECT_EQ(checked.error().code, FormulonErrorCode::kPreconditionFailed);
  }
  ExpectCode(Book().set_table_header_row(0, false), FormulonErrorCode::kInvalidArgument);
}

TEST(WorkbookTableHeaderToggle, TheCheckPredictsTheToggleThroughPendingShifts) {
  Workbook wb = Book();
  AddS(wb, "B2:D5");
  ASSERT_TRUE(wb.set_table_header_row(0, false).has_value());
  PutText(wb, "C2", "X");
  const TableShape shape{Rect("B3:D5"), false, false};
  auto shift = wb.check_table_header_row(0, true, shape, {});
  ASSERT_TRUE(shift.has_value());
  ASSERT_TRUE(shift.value().has_value()) << "an occupied row above takes the insert";
  EXPECT_EQ(shift.value()->range, Rect("B3:D3"));
  EXPECT_EQ(shift.value()->shift, CellShift::kDown);

  // After a pending insert at the data's first row the table sits on B4:D6 under a freshly inserted row.
  const std::vector<PendingShift> pending{PendingShift{0, Rect("B3:D3"), CellShift::kDown}};
  shift = wb.check_table_header_row(0, true, TableShape{Rect("B4:D6"), false, false}, pending);
  ASSERT_TRUE(shift.has_value());
  EXPECT_FALSE(shift.value().has_value()) << "the inserted row above is empty, so the range grows";

  shift = wb.check_table_header_row(0, false, TableShape{Rect("B2:D5"), true, false}, {});
  ASSERT_TRUE(shift.has_value());
  EXPECT_FALSE(shift.value().has_value()) << "hiding never shifts";
  EXPECT_EQ(wb.tables()[0].ref, "B3:D5") << "checks change nothing";
  EXPECT_EQ(TextAt(wb, "C2"), "X");
}

TEST(WorkbookTableHeaderToggle, AHiddenHeaderSurvivesSaveAndReload) {
  Workbook wb = Book();
  AddS(wb, "B2:D5");
  AddHeaderFormulas(wb);
  ASSERT_TRUE(wb.set_table_header_row(0, false).has_value());
  const std::vector<std::uint8_t> saved = SaveXlsx(wb);
  const std::string part = Entry(saved, "xl/tables/table");
  EXPECT_NE(part.find("ref=\"B3:D5\""), std::string::npos) << part;
  EXPECT_NE(part.find("headerRowCount=\"0\""), std::string::npos) << part;
  EXPECT_EQ(part.find("<autoFilter"), std::string::npos) << part;
  EXPECT_NE(part.find("name=\"Sales\""), std::string::npos) << "column names are kept in the part";
  const std::string sheet = Entry(saved, "xl/worksheets/sheet1.xml");
  for (const char* cell : {"r=\"B2\"", "r=\"C2\"", "r=\"D2\""}) {
    EXPECT_EQ(sheet.find(cell), std::string::npos) << cell << " is not written";
  }

  Workbook reloaded = LoadXlsx(saved);
  ASSERT_EQ(reloaded.tables().size(), 1U);
  EXPECT_EQ(reloaded.tables()[0].ref, "B3:D5");
  EXPECT_FALSE(reloaded.tables()[0].header_row);
  Recalc(reloaded);
  ExpectHeaderFormulasRefError(reloaded);

  ASSERT_TRUE(reloaded.set_table_header_row(0, true).has_value());
  EXPECT_EQ(reloaded.tables()[0].ref, "B2:D5");
  EXPECT_EQ(TextAt(reloaded, "C2"), "Sales");
  Recalc(reloaded);
  ExpectHeaderFormulasResolve(reloaded);
  const std::string shown = Entry(SaveXlsx(reloaded), "xl/tables/table");
  EXPECT_EQ(shown.find("headerRowCount"), std::string::npos) << shown;
  EXPECT_NE(shown.find("<autoFilter ref=\"B2:D5\"/>"), std::string::npos) << shown;
}

}  // namespace
}  // namespace formulon
