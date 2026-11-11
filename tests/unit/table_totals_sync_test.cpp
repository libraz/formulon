#include "table_totals_sync.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "cell.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "table.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

constexpr std::uint32_t kTotalsRow = 3U;

// "Sales" A1:C4: header, two data rows, totals row; columns Region, Amount, Qty.
Workbook MakeWorkbook() {
  Workbook wb = Workbook::create();
  TableMetadata table;
  table.sheet_index = 0;
  table.name = "Sales";
  table.display_name = "Sales";
  table.ref = "A1:C4";
  table.totals_row = true;
  for (const char* name : {"Region", "Amount", "Qty"}) {
    TableColumn column;
    column.name = name;
    table.columns.push_back(std::move(column));
  }
  EXPECT_TRUE(wb.set_cell_text(0, 1U, 0U, "East").has_value());
  EXPECT_TRUE(wb.set_cell_value(0, 1U, 1U, Value::number(10.0)).has_value());
  EXPECT_TRUE(wb.set_cell_value(0, 1U, 2U, Value::number(1.0)).has_value());
  EXPECT_TRUE(wb.set_cell_text(0, 2U, 0U, "West").has_value());
  EXPECT_TRUE(wb.set_cell_value(0, 2U, 1U, Value::number(30.0)).has_value());
  EXPECT_TRUE(wb.set_cell_value(0, 2U, 2U, Value::number(2.0)).has_value());
  EXPECT_TRUE(wb.add_table(std::move(table)).has_value());
  return wb;
}

const TableColumn& Column(const Workbook& wb, std::size_t column) {
  return wb.tables()[0].columns[column];
}

const Cell& TotalsCell(const Workbook& wb, std::size_t column) {
  const Cell* cell = wb.sheet(0).cell_at(kTotalsRow, static_cast<std::uint32_t>(column));
  EXPECT_NE(cell, nullptr);
  static const Cell kEmpty;
  return cell != nullptr ? *cell : kEmpty;
}

TotalsCellClass Classify(const Workbook& wb, std::size_t column) {
  return classify_totals_cell(wb, wb.tables()[0], column, TotalsCell(wb, column));
}

void ExpectNoTotals(const TableColumn& column) {
  EXPECT_TRUE(column.totals_function.empty());
  EXPECT_TRUE(column.totals_label.empty());
  EXPECT_TRUE(column.totals_row_formula.empty());
  EXPECT_FALSE(column.totals_row_formula_array);
}

TEST(TableTotalsSyncClassify, EveryBuiltinFunctionIsRecognised) {
  for (const char* token : {"sum", "min", "max", "average", "count", "countNums", "stdDev", "var"}) {
    Workbook wb = MakeWorkbook();
    wb.mutable_tables()[0].columns[1].totals_function = token;
    ASSERT_TRUE(wb.write_table_totals_cell(0, 1).has_value());
    const TotalsCellClass cls = Classify(wb, 1);
    EXPECT_EQ(cls.kind, TotalsCellClass::Kind::kFunction) << token;
    EXPECT_EQ(cls.text, token);
  }
}

TEST(TableTotalsSyncClassify, OtherFormulasAreCustom) {
  const struct {
    const char* formula;
    const char* stored;
  } cases[] = {
      {"=5", "5"},
      {"=SUBTOTAL(9,Sales[Amount])", "SUBTOTAL(9,Sales[Amount])"},
      {"=SUBTOTAL(106,Sales[Amount])", "SUBTOTAL(106,Sales[Amount])"},
      {"=SUBTOTAL(109,Sales[Qty])", "SUBTOTAL(109,Sales[Qty])"},
      {"=SUM(B2:B3)", "SUM(B2:B3)"},
  };
  for (const auto& c : cases) {
    Workbook wb = MakeWorkbook();
    ASSERT_TRUE(wb.set_cell_formula(0, kTotalsRow, 1U, c.formula).has_value());
    const TotalsCellClass cls = Classify(wb, 1);
    EXPECT_EQ(cls.kind, TotalsCellClass::Kind::kCustom) << c.formula;
    EXPECT_FALSE(cls.array);
    EXPECT_NE(cls.text.substr(0, 1), "=");
  }
  Workbook wb = MakeWorkbook();
  ASSERT_TRUE(wb.set_cell_formula(0, kTotalsRow, 1U, "=5").has_value());
  EXPECT_EQ(Classify(wb, 1).text, "5");
}

TEST(TableTotalsSyncClassify, DataStructuredReferenceIsCustom) {
  Workbook wb = MakeWorkbook();
  ASSERT_TRUE(wb.set_cell_formula(0, kTotalsRow, 1U, "=SUBTOTAL(109,Sales[[#Data],[Amount]])").has_value());
  EXPECT_EQ(Classify(wb, 1).kind, TotalsCellClass::Kind::kCustom);
}

TEST(TableTotalsSyncClassify, ArrayFormulaIsCustomWithArrayFlag) {
  Workbook wb = MakeWorkbook();
  ASSERT_TRUE(wb.set_array_formula(0, kTotalsRow, 1U, 1U, 1U, "=SUM(B2:B3)").has_value());
  const TotalsCellClass cls = Classify(wb, 1);
  EXPECT_EQ(cls.kind, TotalsCellClass::Kind::kCustom);
  EXPECT_TRUE(cls.array);
}

TEST(TableTotalsSyncClassify, ValuesAreLabels) {
  Workbook wb = MakeWorkbook();
  ASSERT_TRUE(wb.set_cell_value(0, kTotalsRow, 0U, Value::number(5.0)).has_value());
  ASSERT_TRUE(wb.set_cell_value(0, kTotalsRow, 1U, Value::number(5.5)).has_value());
  ASSERT_TRUE(wb.set_cell_text(0, kTotalsRow, 2U, "Sum").has_value());
  EXPECT_EQ(Classify(wb, 0).kind, TotalsCellClass::Kind::kLabel);
  EXPECT_EQ(Classify(wb, 0).text, "5");
  EXPECT_EQ(Classify(wb, 1).text, "5.5");
  EXPECT_EQ(Classify(wb, 2).text, "Sum");

  ASSERT_TRUE(wb.set_cell_value(0, kTotalsRow, 0U, Value::boolean(true)).has_value());
  ASSERT_TRUE(wb.set_cell_value(0, kTotalsRow, 1U, Value::boolean(false)).has_value());
  ASSERT_TRUE(wb.set_cell_value(0, kTotalsRow, 2U, Value::error(ErrorCode::Div0)).has_value());
  EXPECT_EQ(Classify(wb, 0).text, "TRUE");
  EXPECT_EQ(Classify(wb, 1).text, "FALSE");
  EXPECT_EQ(Classify(wb, 2).kind, TotalsCellClass::Kind::kLabel);
  EXPECT_EQ(Classify(wb, 2).text, "#DIV/0!");
}

TEST(TableTotalsSyncClassify, BlankIsNone) {
  Workbook wb = MakeWorkbook();
  ASSERT_TRUE(wb.set_cell_value(0, kTotalsRow, 1U, Value::blank()).has_value());
  EXPECT_EQ(wb.sheet(0).cell_at(kTotalsRow, 1U), nullptr);
  EXPECT_EQ(classify_totals_cell(wb, wb.tables()[0], 1, Cell{}).kind, TotalsCellClass::Kind::kNone);
}

TEST(TableTotalsSyncNormalize, FunctionRoundTripsThroughTheCell) {
  Workbook wb = MakeWorkbook();
  wb.mutable_tables()[0].columns[1].totals_function = "average";
  ASSERT_TRUE(wb.write_table_totals_cell(0, 1).has_value());
  wb.mutable_tables()[0].columns[1].totals_function.clear();
  wb.normalize_table_totals_cells();
  EXPECT_EQ(Column(wb, 1).totals_function, "average");
  EXPECT_TRUE(Column(wb, 1).totals_label.empty());
}

TEST(TableTotalsSyncNormalize, NumericLabelBecomesString) {
  Workbook wb = MakeWorkbook();
  ASSERT_TRUE(wb.set_cell_value(0, kTotalsRow, 0U, Value::number(5.0)).has_value());
  wb.normalize_table_totals_cells();
  EXPECT_EQ(Column(wb, 0).totals_label, "5");
  ASSERT_TRUE(TotalsCell(wb, 0).cached_value.is_text());
  EXPECT_EQ(TotalsCell(wb, 0).cached_value.as_text(), "5");
}

TEST(TableTotalsSyncNormalize, AnAbsentCellKeepsTheDeclaredSettingsOnLoad) {
  Workbook wb = MakeWorkbook();
  TableColumn& column = wb.mutable_tables()[0].columns[1];
  column.totals_function = "sum";
  wb.normalize_table_totals_cells();
  EXPECT_EQ(Column(wb, 1).totals_function, "sum");
}

TEST(TableTotalsSyncNormalize, AnAbsentCellClearsStaleSettingsOnSync) {
  Workbook wb = MakeWorkbook();
  TableColumn& column = wb.mutable_tables()[0].columns[1];
  column.totals_function = "custom";
  column.totals_row_formula = "1+1";
  column.totals_row_formula_array = true;
  column.totals_label = "stale";
  wb.sync_table_totals_cell(0U, 1U);
  ExpectNoTotals(Column(wb, 1));
}

TEST(TableTotalsSyncNormalize, CustomFormulaIsStoredWithoutEquals) {
  Workbook wb = MakeWorkbook();
  ASSERT_TRUE(wb.set_cell_formula(0, kTotalsRow, 1U, "=SUM(B2:B3)").has_value());
  wb.normalize_table_totals_cells();
  EXPECT_EQ(Column(wb, 1).totals_function, "custom");
  EXPECT_EQ(Column(wb, 1).totals_row_formula, "SUM(B2:B3)");
  EXPECT_FALSE(Column(wb, 1).totals_row_formula_array);
}

TEST(TableTotalsSyncNormalize, TableWithoutTotalsRowIsLeftAlone) {
  Workbook wb = MakeWorkbook();
  wb.mutable_tables()[0].totals_row = false;
  wb.mutable_tables()[0].columns[0].totals_label = "kept";
  wb.normalize_table_totals_cells();
  EXPECT_EQ(Column(wb, 0).totals_label, "kept");
}

// A workbook whose totals cells disagree with the column metadata: the label
// column lost its metadata, and the Amount label is a number.
Workbook DisagreeingWorkbook() {
  Workbook wb = MakeWorkbook();
  EXPECT_TRUE(wb.set_cell_text(0, kTotalsRow, 0U, "Total").has_value());
  EXPECT_TRUE(wb.set_cell_value(0, kTotalsRow, 1U, Value::number(5.0)).has_value());
  wb.mutable_tables()[0].columns[2].totals_function = "sum";
  EXPECT_TRUE(wb.write_table_totals_cell(0, 2).has_value());
  wb.mutable_tables()[0].columns[2].totals_function.clear();
  return wb;
}

void ExpectNormalized(const Workbook& wb) {
  ASSERT_EQ(wb.tables().size(), 1U);
  EXPECT_EQ(Column(wb, 0).totals_label, "Total");
  EXPECT_TRUE(Column(wb, 0).totals_function.empty());
  EXPECT_EQ(Column(wb, 1).totals_label, "5");
  ASSERT_TRUE(TotalsCell(wb, 1).cached_value.is_text());
  EXPECT_EQ(TotalsCell(wb, 1).cached_value.as_text(), "5");
  EXPECT_EQ(Column(wb, 2).totals_function, "sum");
  EXPECT_TRUE(Column(wb, 2).totals_label.empty());
}

TEST(TableTotalsSyncLoad, XlsxLoadRebuildsMetadataFromCells) {
  const Workbook wb = DisagreeingWorkbook();
  auto bytes = io::write_ooxml(wb);
  ASSERT_TRUE(bytes.has_value());
  auto loaded = io::read_ooxml(io::ByteSpan{bytes.value().data(), bytes.value().size()});
  ASSERT_TRUE(loaded.has_value());
  ExpectNormalized(loaded.value().workbook);
}

TEST(TableTotalsSyncLoad, XlsbLoadRebuildsMetadataFromCells) {
  const Workbook wb = DisagreeingWorkbook();
  auto written = io::xlsb::write_xlsb_with_result(wb);
  ASSERT_TRUE(written.has_value());
  const std::vector<std::uint8_t>& bytes = written.value().bytes;
  auto loaded = io::xlsb::read_xlsb(io::ByteSpan{bytes.data(), bytes.size()});
  ASSERT_TRUE(loaded.has_value());
  ExpectNormalized(loaded.value().workbook);
}

TEST(TableTotalsSyncLoad, AgreeingFileLoadsUnchanged) {
  Workbook wb = MakeWorkbook();
  // The intended totals row: Total / max / none, every cell written to agree.
  wb.mutable_tables()[0].columns[0].totals_label = "Total";
  wb.mutable_tables()[0].columns[1].totals_function = "max";
  wb.mutable_tables()[0].columns[2].totals_function.clear();
  ASSERT_TRUE(wb.write_table_totals_cell(0, 0).has_value());
  ASSERT_TRUE(wb.write_table_totals_cell(0, 1).has_value());
  ASSERT_TRUE(wb.write_table_totals_cell(0, 2).has_value());
  auto bytes = io::write_ooxml(wb);
  ASSERT_TRUE(bytes.has_value());
  auto loaded = io::read_ooxml(io::ByteSpan{bytes.value().data(), bytes.value().size()});
  ASSERT_TRUE(loaded.has_value());
  const Workbook& reloaded = loaded.value().workbook;
  EXPECT_EQ(Column(reloaded, 0).totals_label, "Total");
  EXPECT_EQ(Column(reloaded, 1).totals_function, "max");
  ExpectNoTotals(Column(reloaded, 2));
}

}  // namespace
}  // namespace formulon
