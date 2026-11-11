// Table totals defaults, structural deletes over tables, header-only tables and the table pre-checks.

#include <cstddef>
#include <cstdint>
#include <functional>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "auto_filter.h"
#include "c_api/formulon_c.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "merge_range.h"
#include "styles.h"
#include "table.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

MergeRange Rect(const char* a1) {
  const std::optional<MergeRange> rect = parse_a1_rectangle(a1);
  EXPECT_TRUE(rect.has_value()) << a1;
  return rect.value_or(MergeRange{});
}

Workbook Book(const char* profile_id = "mac-365-ja_JP") {
  ExcelProfile profile;
  EXPECT_TRUE(parse_excel_profile_id(profile_id, &profile)) << profile_id;
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

std::string FormulaAt(const Workbook& wb, const char* a1) {
  const MergeRange r = Rect(a1);
  const Cell* cell = wb.sheet(0).cell_at(r.first_row, r.first_col);
  return cell == nullptr ? std::string() : cell->formula_text;
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

// A cell format with number format `id`; `code` registers a custom format under it.
std::uint32_t AddNumberFormat(Workbook& wb, std::uint16_t id, const char* code = nullptr) {
  StylesTable& styles = wb.mutable_styles();
  if (styles.cell_xfs.empty()) {
    styles.cell_xfs.emplace_back();
  }
  if (code != nullptr) {
    styles.num_fmt_strings.emplace_back(code);
    styles.num_fmts.push_back(NumFmtRecord{id, static_cast<std::uint32_t>(styles.num_fmt_strings.size() - 1U)});
  }
  CellXf xf;
  xf.num_fmt_id = id;
  xf.apply_number_format = true;
  styles.cell_xfs.push_back(xf);
  return static_cast<std::uint32_t>(styles.cell_xfs.size() - 1U);
}

void Format(Workbook& wb, const char* a1, std::uint32_t xf) {
  const MergeRange r = Rect(a1);
  ASSERT_TRUE(wb.set_cell_xf_index(0, r.first_row, r.first_col, xf).has_value()) << a1;
}

TableMetadata MakeTable(const std::string& name, const char* ref, std::vector<std::string> columns) {
  TableMetadata table;
  table.name = name;
  table.display_name = name;
  table.ref = ref;
  for (std::string& column : columns) {
    TableColumn entry;
    entry.name = std::move(column);
    table.columns.push_back(std::move(entry));
  }
  return table;
}

void AddTable(Workbook& wb, TableMetadata table) {
  const std::string ref = table.ref;
  const Expected<std::size_t, Error> added = wb.add_table(std::move(table));
  ASSERT_TRUE(added.has_value()) << ref << ": " << added.error().message;
}

TableShape ShapeOf(const Workbook& wb, std::size_t index) {
  const TableMetadata& table = wb.tables()[index];
  return TableShape{Rect(table.ref.c_str()), table.header_row, table.totals_row};
}

void ExpectBare(const TableColumn& column) {
  EXPECT_TRUE(column.totals_label.empty()) << column.name;
  EXPECT_TRUE(column.totals_function.empty()) << column.name;
  EXPECT_TRUE(column.totals_row_formula.empty()) << column.name;
}

void ExpectCode(const Expected<void, Error>& result, FormulonErrorCode code) {
  ASSERT_FALSE(result.has_value());
  EXPECT_EQ(result.error().code, code) << result.error().message;
}

std::vector<std::uint8_t> SaveXlsx(const Workbook& wb) {
  auto saved = io::write_ooxml(wb);
  EXPECT_TRUE(saved.has_value()) << (saved ? "" : saved.error().message);
  return saved ? std::move(saved.value()) : std::vector<std::uint8_t>();
}

std::vector<std::uint8_t> SaveXlsb(const Workbook& wb) {
  auto saved = io::xlsb::write_xlsb(wb);
  EXPECT_TRUE(saved.has_value()) << (saved ? "" : saved.error().message);
  return saved ? std::move(saved.value()) : std::vector<std::uint8_t>();
}

Workbook LoadXlsx(const std::vector<std::uint8_t>& bytes) {
  auto loaded = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(loaded.has_value()) << (loaded ? "" : loaded.error().message);
  return loaded ? std::move(loaded.value().workbook) : Workbook::create();
}

Workbook LoadXlsb(const std::vector<std::uint8_t>& bytes) {
  auto loaded = io::xlsb::read_xlsb(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(loaded.has_value()) << (loaded ? "" : loaded.error().message);
  return loaded ? std::move(loaded.value().workbook) : Workbook::create();
}

// The first table part of an `.xlsx` package, as written.
std::string TablePart(const std::vector<std::uint8_t>& package) {
  io::ZipReader zip;
  EXPECT_TRUE(zip.open(io::ByteSpan{package.data(), package.size()}).has_value());
  for (const std::string& name : zip.list_entries()) {
    if (name.rfind("xl/tables/table", 0) == 0) {
      auto bytes = zip.read_entry(name);
      EXPECT_TRUE(bytes.has_value());
      return bytes ? std::string(bytes.value().begin(), bytes.value().end()) : std::string();
    }
  }
  ADD_FAILURE() << "no table part";
  return {};
}

// The `<table ...>` start tag of a table part.
std::string TableTag(const std::string& part) {
  const std::size_t begin = part.find("<table");
  return begin == std::string::npos ? std::string() : part.substr(begin, part.find('>', begin) - begin + 1U);
}

std::size_t Count(const std::string& text, const std::string& needle) {
  std::size_t count = 0;
  for (std::size_t at = text.find(needle); at != std::string::npos; at = text.find(needle, at + needle.size())) {
    ++count;
  }
  return count;
}

// "Sales" on A1:C5: Region (text) / Sales / Units (numbers), four data rows.
void AddSales(Workbook& wb) {
  PutText(wb, "A1", "Region");
  PutText(wb, "B1", "Sales");
  PutText(wb, "C1", "Units");
  const char* regions[] = {"East", "West", "North", "South"};
  for (int i = 0; i < 4; ++i) {
    const std::string row = std::to_string(i + 2);
    PutText(wb, ("A" + row).c_str(), regions[i]);
    PutNumber(wb, ("B" + row).c_str(), 100.0 + i);
    PutNumber(wb, ("C" + row).c_str(), 10.0 + i);
  }
  AddTable(wb, MakeTable("Sales", "A1:C5", {"Region", "Sales", "Units"}));
}

// `AddSales` with the totals row shown and Sales switched to max, as a user would set it up.
void AddSalesWithMaxTotals(Workbook& wb) {
  AddSales(wb);
  ASSERT_TRUE(wb.set_table_totals_row(0, true).has_value());
  wb.mutable_tables()[0].columns[1].totals_function = "max";
  ASSERT_TRUE(wb.write_table_totals_cell(0, 1).has_value());
}

// ---------------------------------------------------------------------------
// Totals defaults
// ---------------------------------------------------------------------------

TEST(WorkbookTableTotalsDefaults, TheLabelIsTheProfileLocaleWord) {
  struct Case {
    const char* locale;
    const char* label;
  };
  const Case cases[] = {
      {"ja_JP", "集計"},  {"en_US", "Total"},  {"de_DE", "Ergebnis"}, {"fr_FR", "Total"},  {"es_ES", "Total"},
      {"es_MX", "Total"}, {"it_IT", "Totale"}, {"pt_BR", "Total"},    {"zh_CN", "汇总"},   {"zh_TW", "合計"},
      {"ko_KR", "요약"},  {"th_TH", "ผลรวม"},  {"ru_RU", "Итог"},     {"nl_NL", "Totaal"},
  };
  for (const Case& c : cases) {
    for (const char* host : {"mac", "win"}) {
      const std::string id = std::string(host) + "-365-" + c.locale;
      SCOPED_TRACE(id);
      Workbook wb = Book(id.c_str());
      AddSales(wb);
      ASSERT_TRUE(wb.set_table_totals_row(0, true).has_value());
      const TableMetadata& table = wb.tables()[0];
      EXPECT_EQ(table.columns[0].totals_label, c.label);
      ExpectBare(table.columns[1]);
      EXPECT_EQ(table.columns[2].totals_function, "sum");
      EXPECT_EQ(TextAt(wb, "A6"), c.label);
      EXPECT_EQ(FormulaAt(wb, "B6"), "");
      EXPECT_EQ(FormulaAt(wb, "C6"), "=SUBTOTAL(109,Sales[Units])");
    }
  }
}

TEST(WorkbookTableTotalsDefaults, TheLastColumnSumsOnlyNumbersNotShownAsDates) {
  struct Case {
    const char* name;
    std::function<void(Workbook&)> fill;  // Val cells B2:B4
    const char* function;
  };
  const auto numbers = [](Workbook& wb, std::uint32_t xf) {
    for (const char* a1 : {"B2", "B3", "B4"}) {
      PutNumber(wb, a1, 0.25);
      if (xf != 0U) {
        Format(wb, a1, xf);
      }
    }
  };
  const Case cases[] = {
      {"plain numbers", [&](Workbook& wb) { numbers(wb, 0U); }, "sum"},
      {"numeric formulas",
       [](Workbook& wb) {
         for (const char* a1 : {"B2", "B3", "B4"}) {
           PutFormula(wb, a1, "=ROW()*10");
         }
       },
       "sum"},
      {"percent", [&](Workbook& wb) { numbers(wb, AddNumberFormat(wb, 9U)); }, "sum"},
      {"currency", [&](Workbook& wb) { numbers(wb, AddNumberFormat(wb, 200U, "\"¥\"#,##0")); }, "sum"},
      {"time", [&](Workbook& wb) { numbers(wb, AddNumberFormat(wb, 20U)); }, "count"},
      {"elapsed time", [&](Workbook& wb) { numbers(wb, AddNumberFormat(wb, 201U, "[h]:mm")); }, "count"},
      {"date", [&](Workbook& wb) { numbers(wb, AddNumberFormat(wb, 202U, "yyyy/m/d")); }, "count"},
      {"numbers with an error",
       [](Workbook& wb) {
         PutNumber(wb, "B2", 1.0);
         PutFormula(wb, "B3", "=NA()");
         PutNumber(wb, "B4", 3.0);
       },
       "count"},
      {"numbers stored as text",
       [](Workbook& wb) {
         PutText(wb, "B2", "1");
         PutText(wb, "B3", "2");
         PutText(wb, "B4", "3");
       },
       "count"},
      {"text",
       [](Workbook& wb) {
         PutText(wb, "B2", "x");
         PutText(wb, "B3", "y");
         PutText(wb, "B4", "z");
       },
       "count"},
      {"mixed",
       [](Workbook& wb) {
         PutNumber(wb, "B2", 1.0);
         PutText(wb, "B3", "x");
         PutNumber(wb, "B4", 3.0);
       },
       "count"},
      {"logical",
       [](Workbook& wb) {
         for (const char* a1 : {"B2", "B3", "B4"}) {
           const MergeRange r = Rect(a1);
           ASSERT_TRUE(wb.set_cell_value(0, r.first_row, r.first_col, Value::boolean(true)).has_value());
         }
       },
       "count"},
      {"empty", [](Workbook&) {}, "count"},
  };
  for (const Case& c : cases) {
    SCOPED_TRACE(c.name);
    Workbook wb = Book();
    PutText(wb, "A1", "Label");
    PutText(wb, "B1", "Val");
    PutText(wb, "A2", "a");
    PutText(wb, "A3", "b");
    PutText(wb, "A4", "c");
    c.fill(wb);
    Recalc(wb);
    AddTable(wb, MakeTable("Sales", "A1:B4", {"Label", "Val"}));
    ASSERT_TRUE(wb.set_table_totals_row(0, true).has_value());
    EXPECT_EQ(wb.tables()[0].columns[1].totals_function, c.function);
    EXPECT_EQ(FormulaAt(wb, "B5"),
              std::string("=SUBTOTAL(") + (std::string(c.function) == "sum" ? "109" : "103") + ",Sales[Val])");
    EXPECT_EQ(TextAt(wb, "A5"), "集計");
  }
}

TEST(WorkbookTableTotalsDefaults, AOneColumnTableTakesTheFunctionAlone) {
  Workbook wb = Book();
  PutText(wb, "A1", "Val");
  PutNumber(wb, "A2", 1.0);
  PutNumber(wb, "A3", 2.0);
  AddTable(wb, MakeTable("Sales", "A1:A3", {"Val"}));
  ASSERT_TRUE(wb.set_table_totals_row(0, true).has_value());
  EXPECT_TRUE(wb.tables()[0].columns[0].totals_label.empty());
  EXPECT_EQ(wb.tables()[0].columns[0].totals_function, "sum");
  EXPECT_EQ(FormulaAt(wb, "A4"), "=SUBTOTAL(109,Sales[Val])");
}

TEST(WorkbookTableTotalsDefaults, ANewTableSavesTotalsRowShownZeroAndShowingDropsIt) {
  Workbook wb = Book();
  AddSales(wb);
  EXPECT_EQ(wb.tables()[0].totals_row_shown, std::optional<bool>(false));
  const std::string created = TableTag(TablePart(SaveXlsx(wb)));
  EXPECT_NE(created.find("ref=\"A1:C5\""), std::string::npos) << created;
  EXPECT_NE(created.find("totalsRowShown=\"0\""), std::string::npos) << created;
  ASSERT_TRUE(wb.set_table_totals_row(0, true).has_value());
  const std::string shown = TableTag(TablePart(SaveXlsx(wb)));
  EXPECT_NE(shown.find("totalsRowCount=\"1\""), std::string::npos) << shown;
  EXPECT_EQ(shown.find("totalsRowShown"), std::string::npos) << shown;
}

TEST(WorkbookTableTotalsDefaults, OnceShownTheRowRestoresTheColumnsEvenWhenAllAreNone) {
  Workbook wb = Book();
  AddSales(wb);
  ASSERT_TRUE(wb.set_table_totals_row(0, true).has_value());
  // Every totals cell cleared: each column is none.
  for (std::size_t column = 0; column < 3U; ++column) {
    TableColumn& col = wb.mutable_tables()[0].columns[column];
    col.totals_label.clear();
    col.totals_function.clear();
    ASSERT_TRUE(wb.write_table_totals_cell(0, column).has_value());
  }
  ASSERT_TRUE(wb.set_table_totals_row(0, false).has_value());
  ASSERT_TRUE(wb.set_table_totals_row(0, true).has_value());
  EXPECT_EQ(wb.tables()[0].ref, "A1:C6");
  for (const TableColumn& column : wb.tables()[0].columns) {
    ExpectBare(column);
  }
  for (const char* a1 : {"A6", "B6", "C6"}) {
    EXPECT_TRUE(At(wb, a1).is_blank()) << a1;
  }
}

TEST(WorkbookTableTotalsDefaults, TogglingOffKeepsEachColumnsSettingsAcrossASave) {
  Workbook wb = Book();
  AddSalesWithMaxTotals(wb);
  ASSERT_TRUE(wb.set_table_totals_row(0, false).has_value());
  const std::vector<std::uint8_t> saved = SaveXlsx(wb);
  const std::string part = TablePart(saved);
  const std::string tag = TableTag(part);
  EXPECT_EQ(tag.find("totalsRowCount"), std::string::npos) << tag;
  EXPECT_EQ(tag.find("totalsRowShown"), std::string::npos) << tag;
  EXPECT_NE(part.find("totalsRowLabel=\"集計\""), std::string::npos) << part;
  EXPECT_NE(part.find("totalsRowFunction=\"max\""), std::string::npos) << part;
  EXPECT_NE(part.find("totalsRowFunction=\"sum\""), std::string::npos) << part;
  for (Workbook* book : {&wb}) {
    ASSERT_TRUE(book->set_table_totals_row(0, true).has_value());
    EXPECT_EQ(TextAt(*book, "A6"), "集計");
    EXPECT_EQ(FormulaAt(*book, "B6"), "=SUBTOTAL(104,Sales[Sales])");
    EXPECT_EQ(FormulaAt(*book, "C6"), "=SUBTOTAL(109,Sales[Units])");
  }
  Workbook reloaded = LoadXlsx(saved);
  ASSERT_EQ(reloaded.tables().size(), 1U);
  EXPECT_EQ(reloaded.tables()[0].columns[1].totals_function, "max");
  ASSERT_TRUE(reloaded.set_table_totals_row(0, true).has_value());
  EXPECT_EQ(TextAt(reloaded, "A6"), "集計");
  EXPECT_EQ(FormulaAt(reloaded, "B6"), "=SUBTOTAL(104,Sales[Sales])");
  EXPECT_EQ(FormulaAt(reloaded, "C6"), "=SUBTOTAL(109,Sales[Units])");
}

TEST(WorkbookTableTotalsDefaults, ATableCreatedWithItsTotalsRowStartsFromTheDefaults) {
  Workbook wb = Book();
  AddSales(wb);
  TableMetadata table = MakeTable("T", "E1:F4", {"Item", "Qty"});
  table.totals_row = true;
  PutNumber(wb, "F2", 1.0);
  PutNumber(wb, "F3", 2.0);
  AddTable(wb, std::move(table));
  const TableMetadata& created = wb.tables()[1];
  EXPECT_FALSE(created.totals_row_shown.has_value());
  EXPECT_EQ(created.columns[0].totals_label, "集計");
  EXPECT_EQ(created.columns[1].totals_function, "sum");
  EXPECT_EQ(TextAt(wb, "E4"), "集計");
  EXPECT_EQ(FormulaAt(wb, "F4"), "=SUBTOTAL(109,T[Qty])");
}

TEST(WorkbookTableTotalsDefaults, ATotalsRowWithContentAtCreationKeepsItsCells) {
  Workbook wb = Book();
  PutNumber(wb, "F2", 1.0);
  PutNumber(wb, "F3", 2.0);
  PutText(wb, "E4", "Grand");
  PutNumber(wb, "G4", 5.0);
  TableMetadata table = MakeTable("T", "E1:G4", {"Item", "Qty", "Note"});
  table.totals_row = true;
  AddTable(wb, std::move(table));
  const TableMetadata& created = wb.tables()[0];
  EXPECT_EQ(created.columns[0].totals_label, "Grand");
  ExpectBare(created.columns[1]);
  // A number in a totals cell is a label, and the cell becomes its text.
  EXPECT_EQ(created.columns[2].totals_label, "5");
  EXPECT_EQ(TextAt(wb, "E4"), "Grand");
  EXPECT_EQ(FormulaAt(wb, "F4"), "");
  EXPECT_TRUE(At(wb, "F4").is_blank());
  EXPECT_EQ(TextAt(wb, "G4"), "5");
}

// ---------------------------------------------------------------------------
// Structural deletes
// ---------------------------------------------------------------------------

void ExpectTotalsDeleted(const Workbook& wb) {
  const TableMetadata& table = wb.tables()[0];
  EXPECT_EQ(table.ref, "A1:C5");
  EXPECT_FALSE(table.totals_row);
  EXPECT_EQ(table.totals_row_shown, std::optional<bool>(false));
  for (const TableColumn& column : table.columns) {
    ExpectBare(column);
  }
  const std::string part = TablePart(SaveXlsx(wb));
  EXPECT_NE(TableTag(part).find("totalsRowShown=\"0\""), std::string::npos) << part;
  EXPECT_EQ(part.find("totalsRowLabel"), std::string::npos) << part;
  EXPECT_EQ(part.find("totalsRowFunction"), std::string::npos) << part;
  EXPECT_NE(part.find("<tableColumn id=\"2\" name=\"Sales\"/>"), std::string::npos) << part;
}

// Showing the row again after the delete gives the fresh-table defaults; max is lost.
void ExpectDefaultsOnShow(Workbook& wb) {
  ASSERT_TRUE(wb.set_table_totals_row(0, true).has_value());
  EXPECT_EQ(wb.tables()[0].columns[0].totals_label, "集計");
  ExpectBare(wb.tables()[0].columns[1]);
  EXPECT_EQ(wb.tables()[0].columns[2].totals_function, "sum");
  EXPECT_EQ(TextAt(wb, "A6"), "集計");
  EXPECT_EQ(FormulaAt(wb, "B6"), "");
  EXPECT_EQ(FormulaAt(wb, "C6"), "=SUBTOTAL(109,Sales[Units])");
}

TEST(WorkbookTableStructuralDelete, DeletingTheTotalsRowTurnsItOffAndClearsEveryColumn) {
  Workbook wb = Book();
  AddSalesWithMaxTotals(wb);
  ASSERT_TRUE(wb.delete_rows(0, 5U, 1U).has_value());
  ExpectTotalsDeleted(wb);
  ExpectDefaultsOnShow(wb);
}

TEST(WorkbookTableStructuralDelete, DeletingTheTotalsRowCellsShiftingUpDoesTheSame) {
  Workbook wb = Book();
  AddSalesWithMaxTotals(wb);
  ASSERT_TRUE(wb.delete_cells(0, Rect("A6:C6"), CellShift::kUp).has_value());
  ExpectTotalsDeleted(wb);
  ExpectDefaultsOnShow(wb);
}

TEST(WorkbookTableStructuralDelete, TheDeletedStateSurvivesASaveAndReload) {
  Workbook wb = Book();
  AddSalesWithMaxTotals(wb);
  ASSERT_TRUE(wb.delete_rows(0, 5U, 1U).has_value());
  for (bool xlsb : {false, true}) {
    SCOPED_TRACE(xlsb ? "xlsb" : "xlsx");
    Workbook reloaded = xlsb ? LoadXlsb(SaveXlsb(wb)) : LoadXlsx(SaveXlsx(wb));
    ASSERT_EQ(reloaded.tables().size(), 1U);
    EXPECT_EQ(reloaded.tables()[0].totals_row_shown, std::optional<bool>(false));
    reloaded.set_excel_profile(wb.excel_profile());
    ExpectDefaultsOnShow(reloaded);
  }
}

TEST(WorkbookTableStructuralDelete, ARowDeleteTakingTheHeaderRowIsRefused) {
  struct Case {
    const char* name;
    std::function<Expected<void, Error>(Workbook&)> edit;
  };
  const Case cases[] = {
      {"header row", [](Workbook& wb) { return wb.delete_rows(0, 0U, 1U); }},
      {"header and data rows", [](Workbook& wb) { return wb.delete_rows(0, 0U, 3U); }},
      {"header cells", [](Workbook& wb) { return wb.delete_cells(0, Rect("A1:C1"), CellShift::kUp); }},
      {"wider header cells", [](Workbook& wb) { return wb.delete_cells(0, Rect("A1:E2"), CellShift::kUp); }},
  };
  for (const Case& c : cases) {
    SCOPED_TRACE(c.name);
    Workbook wb = Book();
    AddSales(wb);
    ExpectCode(c.edit(wb), FormulonErrorCode::kPreconditionFailed);
    ASSERT_EQ(wb.tables().size(), 1U);
    EXPECT_EQ(wb.tables()[0].ref, "A1:C5");
    EXPECT_EQ(TextAt(wb, "A1"), "Region");
    EXPECT_EQ(TextAt(wb, "A2"), "East");
  }
}

TEST(WorkbookTableStructuralDelete, ADeleteCoveringTheWholeTableRemovesIt) {
  for (bool cells : {false, true}) {
    SCOPED_TRACE(cells ? "cells" : "rows");
    Workbook wb = Book();
    AddSalesWithMaxTotals(wb);
    ASSERT_TRUE((cells ? wb.delete_cells(0, Rect("A1:C6"), CellShift::kUp) : wb.delete_rows(0, 0U, 6U)).has_value());
    EXPECT_TRUE(wb.tables().empty());
  }
}

TEST(WorkbookTableStructuralDelete, AColumnDeleteTakingAHeaderCellIsAnOrdinaryColumnDelete) {
  Workbook wb = Book();
  AddSales(wb);
  ASSERT_TRUE(wb.delete_cols(0, 1U, 1U).has_value());
  ASSERT_EQ(wb.tables().size(), 1U);
  EXPECT_EQ(wb.tables()[0].ref, "A1:B5");
  ASSERT_EQ(wb.tables()[0].columns.size(), 2U);
  EXPECT_EQ(wb.tables()[0].columns[1].name, "Units");
}

// "Sales" on A1:C5 with e2..e6 in E2:E6, G1 =E2, G2 =E6 and G3 =A3, all entered before the table.
Workbook EveryDataRowBook() {
  Workbook wb = Book();
  for (int r = 2; r <= 6; ++r) {
    PutText(wb, ("E" + std::to_string(r)).c_str(), "e" + std::to_string(r));
  }
  PutFormula(wb, "G1", "=E2");
  PutFormula(wb, "G2", "=E6");
  PutFormula(wb, "G3", "=A3");
  AddSales(wb);
  return wb;
}

void ExpectFirstDataRowKept(const Workbook& wb) {
  ASSERT_EQ(wb.tables().size(), 1U);
  EXPECT_EQ(wb.tables()[0].ref, "A1:C2");
  EXPECT_TRUE(wb.tables()[0].insert_row);
  for (const char* a1 : {"A2", "B2", "C2"}) {
    EXPECT_TRUE(At(wb, a1).is_blank()) << a1;
  }
  EXPECT_EQ(TextAt(wb, "A1"), "Region");
  EXPECT_NE(TableTag(TablePart(SaveXlsx(wb))).find("insertRow=\"1\""), std::string::npos);
}

TEST(WorkbookTableStructuralDelete, DeletingEveryDataRowKeepsTheFirstAsAWholeSheetRow) {
  Workbook wb = EveryDataRowBook();
  ASSERT_TRUE(wb.delete_rows(0, 1U, 4U).has_value());
  ExpectFirstDataRowKept(wb);
  EXPECT_EQ(TextAt(wb, "E2"), "e2");
  EXPECT_EQ(TextAt(wb, "E3"), "e6");
  EXPECT_TRUE(At(wb, "E4").is_blank());
  EXPECT_EQ(FormulaAt(wb, "G1"), "=E2");
  EXPECT_EQ(FormulaAt(wb, "G2"), "=E3");
  EXPECT_EQ(FormulaAt(wb, "G3"), "");
}

TEST(WorkbookTableStructuralDelete, DeletingEveryDataRowsCellsKeepsTheFirstWithinTheTable) {
  Workbook wb = EveryDataRowBook();
  ASSERT_TRUE(wb.delete_cells(0, Rect("A2:C5"), CellShift::kUp).has_value());
  ExpectFirstDataRowKept(wb);
  for (int r = 2; r <= 6; ++r) {
    const std::string a1 = "E" + std::to_string(r);
    EXPECT_EQ(TextAt(wb, a1.c_str()), "e" + std::to_string(r));
  }
  EXPECT_EQ(FormulaAt(wb, "G1"), "=E2");
  EXPECT_EQ(FormulaAt(wb, "G2"), "=E6");
  EXPECT_EQ(FormulaAt(wb, "G3"), "=#REF!");
}

TEST(WorkbookTableStructuralDelete, DeletingEveryDataRowAndTheTotalsRowAlsoDeletesTheTotals) {
  Workbook wb = Book();
  AddSalesWithMaxTotals(wb);
  ASSERT_TRUE(wb.delete_rows(0, 1U, 5U).has_value());
  const TableMetadata& table = wb.tables()[0];
  EXPECT_EQ(table.ref, "A1:C2");
  EXPECT_TRUE(table.insert_row);
  EXPECT_FALSE(table.totals_row);
  EXPECT_EQ(table.totals_row_shown, std::optional<bool>(false));
  for (const TableColumn& column : table.columns) {
    ExpectBare(column);
  }
}

TEST(WorkbookTableStructuralDelete, ATableDeletedAlongsideMustLieWhollyBelowTheKeptRow) {
  {
    // Rows 2-5 empty S but cut into T's rows: refused, nothing changes.
    Workbook wb = Book();
    AddSales(wb);
    AddTable(wb, MakeTable("T", "E1:F8", {"P", "Q"}));
    ExpectCode(wb.delete_rows(0, 1U, 4U), FormulonErrorCode::kPreconditionFailed);
    EXPECT_EQ(wb.tables()[0].ref, "A1:C5");
    EXPECT_EQ(wb.tables()[1].ref, "E1:F8");
    EXPECT_EQ(TextAt(wb, "A2"), "East");
  }
  {
    // T lies wholly in rows 3-4, below the kept row: it goes, S keeps its first data row.
    Workbook wb = Book();
    AddSales(wb);
    AddTable(wb, MakeTable("T", "E3:F4", {"P", "Q"}));
    ASSERT_TRUE(wb.delete_rows(0, 1U, 4U).has_value());
    ASSERT_EQ(wb.tables().size(), 1U);
    EXPECT_EQ(wb.tables()[0].ref, "A1:C2");
    EXPECT_TRUE(wb.tables()[0].insert_row);
  }
}

TEST(WorkbookTableStructuralDelete, AHeaderlessTableWhoseRowsAllGoIsRemoved) {
  Workbook wb = Book();
  TableMetadata table = MakeTable("Sales", "A1:B3", {"P", "Q"});
  table.header_row = false;
  AddTable(wb, std::move(table));
  ASSERT_TRUE(wb.delete_rows(0, 0U, 3U).has_value());
  EXPECT_TRUE(wb.tables().empty());
}

// ---------------------------------------------------------------------------
// Header-only tables
// ---------------------------------------------------------------------------

struct Put {
  const char* a1;
  Value value;
};

Put Txt(const char* a1, const char* text) {
  return Put{a1, Value::text(text)};
}

Put Num(const char* a1, double number) {
  return Put{a1, Value::number(number)};
}

Put Blank(const char* a1) {
  return Put{a1, Value::blank()};
}

void Write(Workbook& wb, const Put& put) {
  if (put.value.is_text()) {
    PutText(wb, put.a1, std::string(put.value.as_text()));
  } else {
    PutNumber(wb, put.a1, put.value.as_number());
  }
}

TEST(WorkbookTableHeaderOnly, AHeaderOnlyRangeTakesABlankDataRow) {
  struct Case {
    const char* name;
    std::vector<Put> before;
    const char* range;
    std::vector<std::string> names;
    const char* ref;  // nullptr: refused
    bool shift;
    std::vector<Put> after;
  };
  const Case cases[] = {
      {"row below empty",
       {Txt("B1", "H1"), Txt("C1", "H2"), Txt("D1", "H3")},
       "B1:D1",
       {"H1", "H2", "H3"},
       "B1:D2",
       false,
       {Blank("B2"), Blank("C2"), Blank("D2")}},
      {"value under the header",
       {Txt("B1", "H1"), Txt("C1", "H2"), Txt("D1", "H3"), Num("C2", 7)},
       "B1:D1",
       {"H1", "H2", "H3"},
       "B1:D2",
       true,
       {Blank("C2"), Num("C3", 7)}},
      {"full row under the header",
       {Txt("B1", "H1"), Txt("C1", "H2"), Txt("D1", "H3"), Num("B2", 1), Num("C2", 2), Num("D2", 3)},
       "B1:D1",
       {"H1", "H2", "H3"},
       "B1:D2",
       true,
       {Blank("B2"), Num("B3", 1), Num("C3", 2), Num("D3", 3)}},
      {"value beside the columns",
       {Txt("B1", "H1"), Txt("C1", "H2"), Txt("D1", "H3"), Num("E2", 9)},
       "B1:D1",
       {"H1", "H2", "H3"},
       "B1:D2",
       false,
       {Num("E2", 9), Blank("E3")}},
      {"header on the last grid row",
       {Txt("B1048576", "H1"), Txt("C1048576", "H2"), Txt("D1048576", "H3")},
       "B1048576:D1048576",
       {"H1", "H2", "H3"},
       nullptr,
       false,
       {Txt("B1048576", "H1"), Txt("D1048576", "H3")}},
      {"one column", {Txt("B1", "H1")}, "B1:B1", {"H1"}, "B1:B2", false, {Blank("B2")}},
      {"one column over a value", {Txt("B1", "H1"), Num("B2", 5)}, "B1:B1", {"H1"}, "B1:B2", true, {Num("B3", 5)}},
      {"empty header cells",
       {Txt("B1", "H1")},
       "B1:D1",
       {"H1", "", ""},
       "B1:D2",
       false,
       {Txt("C1", "列1"), Txt("D1", "列2")}},
  };
  for (const Case& c : cases) {
    SCOPED_TRACE(c.name);
    Workbook wb = Book();
    for (const Put& put : c.before) {
      Write(wb, put);
    }
    const Expected<std::size_t, Error> added = wb.add_table(MakeTable("Sales", c.range, c.names));
    if (c.ref == nullptr) {
      ASSERT_FALSE(added.has_value());
      EXPECT_EQ(added.error().code, FormulonErrorCode::kPreconditionFailed);
      EXPECT_TRUE(wb.tables().empty());
    } else {
      ASSERT_TRUE(added.has_value()) << added.error().message;
      const TableMetadata& table = wb.tables()[0];
      EXPECT_EQ(table.ref, c.ref);
      EXPECT_TRUE(table.insert_row);
      EXPECT_EQ(table.insert_row_shift, c.shift);
      ASSERT_FALSE(table.auto_filter_xml.empty());
      std::string filter_ref;
      append_a1_rectangle(filter_ref, table.auto_filter_xml.get()->range);
      EXPECT_EQ(filter_ref, c.ref);
    }
    for (const Put& put : c.after) {
      EXPECT_EQ(At(wb, put.a1), put.value) << put.a1;
    }
  }
}

TEST(WorkbookTableHeaderOnly, AHeaderAndATotalsRowOnOneRowAreAnInvalidArgument) {
  Workbook wb = Book();
  TableMetadata table = MakeTable("Sales", "B1:D1", {"H1", "H2", "H3"});
  table.totals_row = true;
  const Expected<std::size_t, Error> added = wb.add_table(std::move(table));
  ASSERT_FALSE(added.has_value());
  EXPECT_EQ(added.error().code, FormulonErrorCode::kInvalidArgument);
  EXPECT_TRUE(wb.tables().empty());
}

// The "value under the header" case: B1:D2 with both marks, 7 moved to C3.
Workbook ShiftedHeaderOnlyBook() {
  Workbook wb = Book();
  for (const Put& put : {Txt("B1", "H1"), Txt("C1", "H2"), Txt("D1", "H3"), Num("C2", 7)}) {
    Write(wb, put);
  }
  AddTable(wb, MakeTable("Sales", "B1:D1", {"H1", "H2", "H3"}));
  return wb;
}

TEST(WorkbookTableHeaderOnly, BothMarksRoundTripThroughXlsxAndXlsb) {
  const Workbook wb = ShiftedHeaderOnlyBook();
  const std::string part = TablePart(SaveXlsx(wb));
  EXPECT_EQ(Count(part, "insertRow=\"1\""), 1U) << part;
  EXPECT_EQ(Count(part, "insertRowShift=\"1\""), 1U) << part;
  for (bool xlsb : {false, true}) {
    SCOPED_TRACE(xlsb ? "xlsb" : "xlsx");
    const Workbook back = xlsb ? LoadXlsb(SaveXlsb(wb)) : LoadXlsx(SaveXlsx(wb));
    ASSERT_EQ(back.tables().size(), 1U);
    EXPECT_TRUE(back.tables()[0].insert_row);
    EXPECT_TRUE(back.tables()[0].insert_row_shift);
    EXPECT_EQ(back.tables()[0].root_extra_attrs.find("insertRow"), std::string::npos);
  }
}

TEST(WorkbookTableHeaderOnly, AResizeEndsThePlaceholderRow) {
  Workbook wb = Book();
  PutText(wb, "A1", "H1");
  AddTable(wb, MakeTable("Sales", "A1:C1", {"H1", "H2", "H3"}));
  ASSERT_TRUE(wb.resize_table(0, Rect("A1:C3")).has_value());
  EXPECT_EQ(wb.tables()[0].ref, "A1:C3");
  EXPECT_FALSE(wb.tables()[0].insert_row);
}

TEST(WorkbookTableHeaderOnly, InsertingCellsAtThePlaceholderRowLeavesTheRange) {
  Workbook wb = Book();
  AddTable(wb, MakeTable("Sales", "A1:C1", {"H1", "H2", "H3"}));
  PutText(wb, "A3", "x");
  PutText(wb, "D2", "out");
  PutFormula(wb, "F1", "=A3");
  ASSERT_TRUE(wb.insert_cells(0, Rect("A2:C2"), CellShift::kDown).has_value());
  const TableMetadata& table = wb.tables()[0];
  EXPECT_EQ(table.ref, "A1:C2");
  EXPECT_FALSE(table.insert_row);
  ASSERT_FALSE(table.auto_filter_xml.empty());
  std::string filter_ref;
  append_a1_rectangle(filter_ref, table.auto_filter_xml.get()->range);
  EXPECT_EQ(filter_ref, "A1:C2");
  EXPECT_EQ(TextAt(wb, "A4"), "x");
  EXPECT_TRUE(At(wb, "A3").is_blank());
  EXPECT_EQ(TextAt(wb, "D2"), "out");
  EXPECT_EQ(FormulaAt(wb, "F1"), "=A4");
}

TEST(WorkbookTableHeaderOnly, AStructuralEditEndsBothMarks) {
  {
    Workbook wb = Book();
    AddTable(wb, MakeTable("Sales", "A1:C1", {"H1", "H2", "H3"}));
    ASSERT_TRUE(wb.insert_cells(0, Rect("A2:C2"), CellShift::kDown).has_value());
    EXPECT_FALSE(wb.tables()[0].insert_row);
  }
  {
    Workbook wb = ShiftedHeaderOnlyBook();
    ASSERT_TRUE(wb.insert_rows(0, 0U, 1U).has_value());
    EXPECT_EQ(wb.tables()[0].ref, "B2:D3");
    EXPECT_FALSE(wb.tables()[0].insert_row);
    EXPECT_FALSE(wb.tables()[0].insert_row_shift);
  }
  {
    Workbook wb = Book();
    AddTable(wb, MakeTable("Sales", "A1:C1", {"H1", "H2", "H3"}));
    ASSERT_TRUE(wb.set_table_totals_row(0, true).has_value());
    EXPECT_FALSE(wb.tables()[0].insert_row);
  }
}

TEST(WorkbookTableHeaderOnly, AResizeWhileTheCreationShiftIsMarkedIsRefused) {
  Workbook wb = ShiftedHeaderOnlyBook();
  ExpectCode(wb.check_table_resize(0, Rect("B1:D3")), FormulonErrorCode::kPreconditionFailed);
  ExpectCode(wb.resize_table(0, Rect("B1:D3")), FormulonErrorCode::kPreconditionFailed);
  EXPECT_EQ(wb.tables()[0].ref, "B1:D2");
  EXPECT_TRUE(wb.tables()[0].insert_row);
  EXPECT_TRUE(wb.tables()[0].insert_row_shift);
  // The current range still passes.
  EXPECT_TRUE(wb.check_table_resize(0, Rect("B1:D2")).has_value());
}

// Through the C setters: content in the placeholder row ends `insertRow`, `insertRowShift` stays, a blank keeps both.
TEST(WorkbookTableHeaderOnly, WritingThePlaceholderRowEndsOnlyThePlaceholderMark) {
  struct Case {
    const char* name;
    std::function<fm_status_t(fm_workbook_t*)> write;
    bool insert_row;
  };
  const Case cases[] = {
      {"text", [](fm_workbook_t* wb) { return fm_sheet_set_text(wb, 0, 1U, 1U, "East"); }, false},
      {"formula", [](fm_workbook_t* wb) { return fm_sheet_set_formula(wb, 0, 1U, 2U, "=1+1"); }, false},
      {"blank", [](fm_workbook_t* wb) { return fm_sheet_set_blank(wb, 0, 1U, 1U); }, true},
  };
  for (const Case& c : cases) {
    SCOPED_TRACE(c.name);
    fm_workbook_t* wb = nullptr;
    ASSERT_EQ(fm_workbook_create_default(&wb), 0);
    ASSERT_EQ(fm_sheet_set_text(wb, 0, 0U, 1U, "H1"), 0);
    ASSERT_EQ(fm_sheet_set_text(wb, 0, 0U, 2U, "H2"), 0);
    ASSERT_EQ(fm_sheet_set_text(wb, 0, 0U, 3U, "H3"), 0);
    ASSERT_EQ(fm_sheet_set_number(wb, 0, 1U, 2U, 7.0), 0);
    fm_table_spec_t spec = FM_STRUCT_INIT(fm_table_spec_t);
    spec.range = fm_cell_rect_t{0U, 1U, 0U, 3U};
    spec.name = "Sales";
    std::size_t index = 0;
    ASSERT_EQ(fm_sheet_add_table(wb, 0, &spec, &index), 0);
    ASSERT_EQ(c.write(wb), 0);
    std::uint8_t* bytes = nullptr;
    std::size_t len = 0;
    ASSERT_EQ(fm_workbook_save(wb, FM_WORKBOOK_FORMAT_XLSX, &bytes, &len), 0);
    const std::vector<std::uint8_t> package(bytes, bytes + len);
    fm_buffer_destroy(bytes);
    fm_workbook_destroy(wb);
    const Workbook back = LoadXlsx(package);
    ASSERT_EQ(back.tables().size(), 1U);
    EXPECT_EQ(back.tables()[0].ref, "B1:D2");
    EXPECT_EQ(back.tables()[0].insert_row, c.insert_row);
    EXPECT_TRUE(back.tables()[0].insert_row_shift);
  }
}

// ---------------------------------------------------------------------------
// Pre-checks
// ---------------------------------------------------------------------------

// `check_table_totals_row` answers as `set_table_totals_row` does, and changes nothing.
void ExpectTotalsCheckAgrees(Workbook& wb, bool shown, std::optional<CellShift> shift) {
  const std::string ref = wb.tables()[0].ref;
  const Expected<std::optional<PendingShift>, Error> checked = wb.check_table_totals_row(0, shown, ShapeOf(wb, 0), {});
  EXPECT_EQ(wb.tables()[0].ref, ref);
  const Expected<void, Error> applied = wb.set_table_totals_row(0, shown);
  ASSERT_EQ(checked.has_value(), applied.has_value()) << (applied ? checked.error().message : applied.error().message);
  if (!applied.has_value()) {
    EXPECT_EQ(checked.error().code, applied.error().code);
    EXPECT_EQ(wb.tables()[0].ref, ref);
    return;
  }
  ASSERT_EQ(checked.value().has_value(), shift.has_value());
  if (shift.has_value()) {
    EXPECT_EQ(checked.value()->shift, *shift);
  }
}

TEST(WorkbookTableChecks, TheTotalsRowCheckAnswersAsTheToggle) {
  struct Case {
    const char* name;
    std::function<void(Workbook&)> setup;  // after "Sales" A1:C5
    bool shown;
    std::optional<CellShift> shift;  // when accepted
  };
  const Case cases[] = {
      {"show by shifting", [](Workbook&) {}, true, CellShift::kDown},
      {"show into the free row above a straddling merge",
       [](Workbook& wb) { ASSERT_TRUE(wb.add_merge(0, Rect("C8:D8")).has_value()); }, true, std::nullopt},
      {"show into a busy row above a straddling merge",
       [](Workbook& wb) {
         PutText(wb, "A6", "x");
         ASSERT_TRUE(wb.add_merge(0, Rect("C8:D8")).has_value());
       },
       true, std::nullopt},
      {"hide by shifting back", [](Workbook& wb) { ASSERT_TRUE(wb.set_table_totals_row(0, true).has_value()); }, false,
       CellShift::kUp},
      {"hide in place above a straddling merge",
       [](Workbook& wb) {
         ASSERT_TRUE(wb.set_table_totals_row(0, true).has_value());
         ASSERT_TRUE(wb.add_merge(0, Rect("C9:D9")).has_value());
       },
       false, std::nullopt},
      {"hide a referenced row",
       [](Workbook& wb) {
         ASSERT_TRUE(wb.set_table_totals_row(0, true).has_value());
         PutFormula(wb, "H1", "=C6");
       },
       false, std::nullopt},
  };
  for (const Case& c : cases) {
    SCOPED_TRACE(c.name);
    Workbook wb = Book();
    AddSales(wb);
    c.setup(wb);
    ExpectTotalsCheckAgrees(wb, c.shown, c.shift);
  }
  {
    SCOPED_TRACE("last grid row");
    Workbook wb = Book();
    AddTable(wb, MakeTable("Sales", "A1048574:B1048576", {"P", "Q"}));
    ExpectTotalsCheckAgrees(wb, true, std::nullopt);
  }
  {
    SCOPED_TRACE("hide with no data row");
    Workbook wb = Book();
    TableMetadata table = MakeTable("Sales", "A1:B2", {"P", "Q"});
    table.totals_row = true;
    AddTable(wb, std::move(table));
    ExpectTotalsCheckAgrees(wb, false, std::nullopt);
  }
}

TEST(WorkbookTableChecks, TheBandEditCheckAnswersAsTheEdit) {
  struct Case {
    const char* name;
    std::function<void(Workbook&)> setup;
    const char* range;
    CellShift shift;
  };
  const Case cases[] = {
      {"plain insert", [](Workbook&) {}, "A7:C7", CellShift::kDown},
      {"value pushed off the grid", [](Workbook& wb) { PutText(wb, "A1048576", "z"); }, "A7:A7", CellShift::kDown},
      {"array straddling the band",
       [](Workbook& wb) { ASSERT_TRUE(wb.set_array_formula(0, 9U, 2U, 2U, 2U, "=1").has_value()); }, "A7:C7",
       CellShift::kDown},
      {"table header cells", [](Workbook&) {}, "A1:C1", CellShift::kUp},
      {"every data row beside another table", [](Workbook& wb) { AddTable(wb, MakeTable("T", "E1:F8", {"P", "Q"})); },
       "A2:F5", CellShift::kUp},
  };
  for (const Case& c : cases) {
    SCOPED_TRACE(c.name);
    Workbook wb = Book();
    AddSales(wb);
    c.setup(wb);
    const Expected<void, Error> checked = wb.check_band_edit(0, Rect(c.range), c.shift);
    EXPECT_EQ(wb.tables()[0].ref, "A1:C5");
    const bool insert = c.shift == CellShift::kDown || c.shift == CellShift::kRight;
    const Expected<void, Error> applied =
        insert ? wb.insert_cells(0, Rect(c.range), c.shift) : wb.delete_cells(0, Rect(c.range), c.shift);
    ASSERT_EQ(checked.has_value(), applied.has_value());
    if (!applied.has_value()) {
      EXPECT_EQ(checked.error().code, applied.error().code);
    }
  }
}

TEST(WorkbookTableChecks, TheResizeCheckReadsStructuresWhereAPendingShiftMovesThem) {
  Workbook wb = Book();
  AddTable(wb, MakeTable("Sales", "A1:B3", {"P", "Q"}));
  ASSERT_TRUE(wb.add_merge(0, Rect("A5:B5")).has_value());  // inside the columns: moves down
  ASSERT_TRUE(wb.add_merge(0, Rect("B7:C7")).has_value());  // across the band edge: unmerged
  const std::vector<PendingShift> pending{PendingShift{0U, Rect("A4:B4"), CellShift::kDown}};
  const TableShape shape = ShapeOf(wb, 0);
  ExpectCode(wb.check_table_resize(0, Rect("A1:B5"), shape, {}), FormulonErrorCode::kPreconditionFailed);
  EXPECT_TRUE(wb.check_table_resize(0, Rect("A1:B5"), shape, pending).has_value());
  ExpectCode(wb.check_table_resize(0, Rect("A1:B6"), shape, pending), FormulonErrorCode::kPreconditionFailed);
  ExpectCode(wb.check_table_resize(0, Rect("A1:B7"), shape, {}), FormulonErrorCode::kPreconditionFailed);
  // The assumed shape's first row and header decide, not the table's current ones.
  ExpectCode(wb.check_table_resize(0, Rect("A1:B4"), TableShape{Rect("A2:B3"), false, false}, {}),
             FormulonErrorCode::kInvalidArgument);
}

TEST(WorkbookTableChecks, TheTotalsRowCheckReadsCellsWhereAPendingShiftMovesThem) {
  Workbook wb = Book();
  AddTable(wb, MakeTable("Sales", "A1:B3", {"P", "Q"}));
  PutText(wb, "A4", "x");
  ASSERT_TRUE(wb.add_merge(0, Rect("B8:C8")).has_value());
  const TableShape shape = ShapeOf(wb, 0);
  // Under the straddling merge the busy row below refuses the totals row.
  const auto refused = wb.check_table_totals_row(0, true, shape, {});
  ASSERT_FALSE(refused.has_value());
  EXPECT_EQ(refused.error().code, FormulonErrorCode::kPreconditionFailed);
  // A pending insert at row 4 moves x down and unmerges B8:C8, so the row is shifted in.
  const std::vector<PendingShift> pending{PendingShift{0U, Rect("A4:B4"), CellShift::kDown}};
  const auto shifted = wb.check_table_totals_row(0, true, shape, pending);
  ASSERT_TRUE(shifted.has_value()) << shifted.error().message;
  ASSERT_TRUE(shifted.value().has_value());
  EXPECT_EQ(shifted.value()->range, Rect("A4:B4"));
  EXPECT_EQ(shifted.value()->shift, CellShift::kDown);
  // A delete pulling A5 up empties the row below while the merge still straddles: the free row is taken.
  const std::vector<PendingShift> pulled{PendingShift{0U, Rect("A4:A4"), CellShift::kUp}};
  const auto taken = wb.check_table_totals_row(0, true, shape, pulled);
  ASSERT_TRUE(taken.has_value()) << taken.error().message;
  EXPECT_FALSE(taken.value().has_value());
}

}  // namespace
}  // namespace formulon
