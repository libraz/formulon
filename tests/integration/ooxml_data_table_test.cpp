//
// Integration test: What-If data tables (`<f t="dataTable">`). A load -> save
// cycle keeps the table, row/column edits move it, a recalc fills the result
// cells by substituting into the input cells (leaving them untouched), and
// the result range refuses writes.

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "workbook.h"

namespace formulon {
namespace {

io::ByteSpan SpanOf(const std::vector<std::uint8_t>& bytes) {
  return io::ByteSpan{bytes.data(), bytes.size()};
}

constexpr std::string_view kDecl = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
constexpr std::string_view kMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
constexpr std::string_view kRelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

std::vector<std::uint8_t> BuildZip(const std::vector<std::pair<std::string, std::string>>& parts) {
  mz_zip_archive writer{};
  EXPECT_NE(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_FALSE);
  for (const auto& p : parts) {
    EXPECT_NE(mz_zip_writer_add_mem(&writer, p.first.c_str(), p.second.data(), p.second.size(),
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_FALSE);
  }
  void* archive_ptr = nullptr;
  std::size_t archive_size = 0;
  EXPECT_NE(mz_zip_writer_finalize_heap_archive(&writer, &archive_ptr, &archive_size), MZ_FALSE);
  EXPECT_NE(mz_zip_writer_end(&writer), MZ_FALSE);
  std::vector<std::uint8_t> out(static_cast<const std::uint8_t*>(archive_ptr),
                                static_cast<const std::uint8_t*>(archive_ptr) + archive_size);
  mz_free(archive_ptr);
  return out;
}

// One-sheet package whose `<sheetData>` is `rows`; `calc_pr` is spliced after `<sheets>`.
std::vector<std::uint8_t> Package(const std::string& rows, const std::string& calc_pr = {}) {
  const std::string content_types =
      std::string(kDecl) +
      "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
      "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
      "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
      "<Override PartName=\"/xl/workbook.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
      "<Override PartName=\"/xl/worksheets/sheet1.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/></Types>";
  const std::string root_rels = std::string(kDecl) +
                                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                                "<Relationship Id=\"rId1\" Type=\"" +
                                std::string(kRelNs) + "/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>";
  const std::string workbook_rels = std::string(kDecl) +
                                    "<Relationships "
                                    "xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                                    "<Relationship Id=\"rId1\" Type=\"" +
                                    std::string(kRelNs) +
                                    "/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>";
  const std::string workbook =
      std::string(kDecl) + "<workbook xmlns=\"" + std::string(kMain) + "\" xmlns:r=\"" + std::string(kRelNs) +
      "\"><sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>" + calc_pr + "</workbook>";
  const std::string sheet = std::string(kDecl) + "<worksheet xmlns=\"" + std::string(kMain) + "\"><sheetData>" + rows +
                            "</sheetData></worksheet>";
  return BuildZip({{"[Content_Types].xml", content_types},
                   {"_rels/.rels", root_rels},
                   {"xl/workbook.xml", workbook},
                   {"xl/_rels/workbook.xml.rels", workbook_rels},
                   {"xl/worksheets/sheet1.xml", sheet}});
}

Workbook LoadOrDie(const std::vector<std::uint8_t>& bytes) {
  auto load_or = io::read_ooxml(SpanOf(bytes));
  EXPECT_TRUE(static_cast<bool>(load_or)) << (load_or ? "" : load_or.error().message);
  return std::move(load_or.value().workbook);
}

std::vector<std::uint8_t> SaveOrDie(const Workbook& wb) {
  auto save_or = wb.save();
  EXPECT_TRUE(static_cast<bool>(save_or)) << (save_or ? "" : save_or.error().message);
  return save_or.value();
}

std::string ReadPart(const std::vector<std::uint8_t>& bytes, std::string_view name) {
  io::ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes))));
  auto entry_or = zip.read_entry(name);
  if (!entry_or) {
    return {};
  }
  return std::string(reinterpret_cast<const char*>(entry_or.value().data()), entry_or.value().size());
}

void Recalc(Workbook& wb) {
  auto stats = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(stats)) << stats.error().message;
}

double Number(const Workbook& wb, std::uint32_t row, std::uint32_t col) {
  const Cell* cell = wb.sheet(0).cell_at(row, col);
  EXPECT_NE(cell, nullptr) << "row=" << row << " col=" << col;
  if (cell == nullptr) {
    return std::nan("");
  }
  EXPECT_TRUE(cell->cached_value.is_number()) << "row=" << row << " col=" << col;
  return cell->cached_value.is_number() ? cell->cached_value.as_number() : std::nan("");
}

// Monthly payment of a `pv` loan over `nper` months at the annual `rate`.
double Pmt(double rate, double nper, double pv) {
  const double r = rate / 12.0;
  return pv * r / (1.0 - std::pow(1.0 + r, -nper));
}

// The `<f>` of the first cell named `ref` in the saved sheet, or an empty node.
pugi::xml_node FormulaOf(pugi::xml_document& doc, const std::vector<std::uint8_t>& saved, std::string_view ref) {
  const std::string xml = ReadPart(saved, "xl/worksheets/sheet1.xml");
  EXPECT_TRUE(doc.load_string(xml.c_str()));
  for (pugi::xml_node row : doc.child("worksheet").child("sheetData").children("row")) {
    for (pugi::xml_node c : row.children("c")) {
      if (std::string_view(c.attribute("r").value()) == ref) {
        return c.child("f");
      }
    }
  }
  return {};
}

// Loan table: rates down C5:C8, the payment formula in D4 and the interest-for-year-one formula in E4,
// both fed from the rate input B1.
std::string LoanRows() {
  return "<row r=\"1\"><c r=\"B1\"><v>0.05</v></c></row>"
         "<row r=\"2\"><c r=\"B2\"><v>12</v></c></row>"
         "<row r=\"3\"><c r=\"B3\"><v>10000</v></c></row>"
         "<row r=\"4\"><c r=\"D4\"><f>PMT(B1/12,B2,-B3)</f><v>0</v></c><c r=\"E4\"><f>B1*B3</f><v>0</v></c></row>"
         "<row r=\"5\"><c r=\"C5\"><v>0.03</v></c><c r=\"D5\"><f t=\"dataTable\" ref=\"D5:E8\" dt2D=\"0\" dtr=\"0\" "
         "r1=\"B1\"/><v>0</v></c><c r=\"E5\"><v>0</v></c></row>"
         "<row r=\"6\"><c r=\"C6\"><v>0.06</v></c><c r=\"D6\"><v>0</v></c><c r=\"E6\"><v>0</v></c></row>"
         "<row r=\"7\"><c r=\"C7\"><v>0.09</v></c><c r=\"D7\"><v>0</v></c><c r=\"E7\"><v>0</v></c></row>"
         "<row r=\"8\"><c r=\"C8\"><v>0.12</v></c><c r=\"D8\"><v>0</v></c><c r=\"E8\"><v>0</v></c></row>";
}

// Multiplication table: row values D3:F3, column values C4:C6, corner formula C3, inputs F1 (row) and F2 (column).
// `row_one_extra` is appended to row 1, after F1.
std::string MultiplicationRows(const std::string& row_one_extra = {}) {
  return "<row r=\"1\"><c r=\"F1\"><v>1</v></c>" + row_one_extra +
         "</row>"
         "<row r=\"2\"><c r=\"F2\"><v>1</v></c></row>"
         "<row r=\"3\"><c r=\"C3\"><f>F1*F2</f><v>0</v></c><c r=\"D3\"><v>1</v></c><c r=\"E3\"><v>2</v></c>"
         "<c r=\"F3\"><v>3</v></c></row>"
         "<row r=\"4\"><c r=\"C4\"><v>10</v></c><c r=\"D4\"><f t=\"dataTable\" ref=\"D4:F6\" dt2D=\"1\" dtr=\"1\" "
         "r1=\"F1\" r2=\"F2\"/><v>0</v></c><c r=\"E4\"><v>0</v></c><c r=\"F4\"><v>0</v></c></row>"
         "<row r=\"5\"><c r=\"C5\"><v>20</v></c><c r=\"D5\"><v>0</v></c><c r=\"E5\"><v>0</v></c><c "
         "r=\"F5\"><v>0</v></c>"
         "</row>"
         "<row r=\"6\"><c r=\"C6\"><v>30</v></c><c r=\"D6\"><v>0</v></c><c r=\"E6\"><v>0</v></c><c "
         "r=\"F6\"><v>0</v></c>"
         "</row>";
}

TEST(OoxmlDataTable, ReaderModelsTheTable) {
  Workbook wb = LoadOrDie(Package(MultiplicationRows()));
  ASSERT_EQ(wb.sheet(0).data_tables().size(), 1U);
  const DataTable& t = wb.sheet(0).data_tables()[0];
  EXPECT_EQ(t.ref, (MergeRange{3, 3, 5, 5}));
  EXPECT_TRUE(t.two_dimensional);
  EXPECT_TRUE(t.row_oriented);
  EXPECT_TRUE(t.r1.present);
  EXPECT_EQ(t.r1.row, 0U);
  EXPECT_EQ(t.r1.col, 5U);
  EXPECT_EQ(t.r2.row, 1U);
  EXPECT_EQ(t.r2.col, 5U);
  EXPECT_EQ(wb.data_table_at(0, 4, 4), &t);
  EXPECT_EQ(wb.data_table_at(0, 2, 3), nullptr);
}

TEST(OoxmlDataTable, OneVariableColumnTableComputesLoanPayments) {
  Workbook wb = LoadOrDie(Package(LoanRows()));
  Recalc(wb);
  const double rates[] = {0.03, 0.06, 0.09, 0.12};
  for (std::uint32_t i = 0; i < 4; ++i) {
    EXPECT_NEAR(Number(wb, 4 + i, 3), Pmt(rates[i], 12.0, 10000.0), 1e-9) << "row " << i;
    EXPECT_NEAR(Number(wb, 4 + i, 4), rates[i] * 10000.0, 1e-9) << "row " << i;
  }
  // PMT(0.5%, 12, -10000), the textbook figure.
  EXPECT_NEAR(Number(wb, 5, 3), 860.6643, 1e-4);
  // The input and everything reading it keep their real values.
  EXPECT_DOUBLE_EQ(Number(wb, 0, 1), 0.05);
  EXPECT_NEAR(Number(wb, 3, 3), Pmt(0.05, 12.0, 10000.0), 1e-9);
  EXPECT_DOUBLE_EQ(Number(wb, 3, 4), 500.0);
  EXPECT_EQ(wb.recalc_engine().dirty().size(), 0U);
}

TEST(OoxmlDataTable, OneVariableRowTableSubstitutesTheTopRow) {
  const std::string rows =
      "<row r=\"1\"><c r=\"B1\"><v>2</v></c></row>"
      "<row r=\"4\"><c r=\"D4\"><v>1</v></c><c r=\"E4\"><v>2</v></c><c r=\"F4\"><v>3</v></c></row>"
      "<row r=\"5\"><c r=\"C5\"><f>B1*10</f><v>0</v></c><c r=\"D5\"><f t=\"dataTable\" ref=\"D5:F6\" dt2D=\"0\" "
      "dtr=\"1\" r1=\"B1\"/><v>0</v></c><c r=\"E5\"><v>0</v></c><c r=\"F5\"><v>0</v></c></row>"
      "<row r=\"6\"><c r=\"C6\"><f>B1+1</f><v>0</v></c><c r=\"D6\"><v>0</v></c><c r=\"E6\"><v>0</v></c>"
      "<c r=\"F6\"><v>0</v></c></row>";
  Workbook wb = LoadOrDie(Package(rows));
  Recalc(wb);
  const double expected[2][3] = {{10, 20, 30}, {2, 3, 4}};
  for (std::uint32_t i = 0; i < 2; ++i) {
    for (std::uint32_t j = 0; j < 3; ++j) {
      EXPECT_DOUBLE_EQ(Number(wb, 4 + i, 3 + j), expected[i][j]) << i << "," << j;
    }
  }
  EXPECT_DOUBLE_EQ(Number(wb, 0, 1), 2.0);
  EXPECT_DOUBLE_EQ(Number(wb, 4, 2), 20.0);
}

TEST(OoxmlDataTable, TwoVariableTableIsAMultiplicationTable) {
  Workbook wb = LoadOrDie(Package(MultiplicationRows()));
  Recalc(wb);
  const double rows[] = {10, 20, 30};
  const double cols[] = {1, 2, 3};
  for (std::uint32_t i = 0; i < 3; ++i) {
    for (std::uint32_t j = 0; j < 3; ++j) {
      EXPECT_DOUBLE_EQ(Number(wb, 3 + i, 3 + j), rows[i] * cols[j]) << i << "," << j;
    }
  }
  EXPECT_DOUBLE_EQ(Number(wb, 0, 5), 1.0);
  EXPECT_DOUBLE_EQ(Number(wb, 1, 5), 1.0);
  EXPECT_DOUBLE_EQ(Number(wb, 2, 2), 1.0);
}

TEST(OoxmlDataTable, SubstitutionLeavesDependentsAndInputsUntouched) {
  // G1 and H1 read the input F1 directly and through a chain.
  const std::string rows =
      MultiplicationRows("<c r=\"G1\"><f>F1*100</f><v>0</v></c><c r=\"H1\"><f>G1+1</f><v>0</v></c>");
  Workbook wb = LoadOrDie(Package(rows));
  Recalc(wb);
  EXPECT_DOUBLE_EQ(Number(wb, 0, 6), 100.0);
  EXPECT_DOUBLE_EQ(Number(wb, 0, 7), 101.0);
  EXPECT_DOUBLE_EQ(Number(wb, 0, 5), 1.0);
  // A second pass with nothing changed leaves the same values.
  Recalc(wb);
  EXPECT_DOUBLE_EQ(Number(wb, 0, 6), 100.0);
  EXPECT_DOUBLE_EQ(Number(wb, 3, 5), 30.0);
  // Editing the formula's precedent recomputes the table.
  ASSERT_TRUE(wb.set_cell_formula(0, 2, 2, "F1*F2*2"));
  Recalc(wb);
  EXPECT_DOUBLE_EQ(Number(wb, 3, 5), 60.0);
  EXPECT_DOUBLE_EQ(Number(wb, 0, 5), 1.0);
}

TEST(OoxmlDataTable, ResultRangeRefusesWrites) {
  Workbook wb = LoadOrDie(Package(MultiplicationRows()));
  Recalc(wb);
  const double before = Number(wb, 4, 4);
  const std::vector<std::pair<std::uint32_t, std::uint32_t>> cells = {{3, 3}, {4, 4}, {5, 5}};
  for (const auto& [row, col] : cells) {
    auto value = wb.set_cell_value(0, row, col, Value::number(99.0));
    ASSERT_FALSE(static_cast<bool>(value)) << row << "," << col;
    EXPECT_EQ(value.error().code, FormulonErrorCode::kPreconditionFailed);
    EXPECT_NE(value.error().message.find("data table"), std::string::npos);
    EXPECT_FALSE(static_cast<bool>(wb.set_cell_text(0, row, col, "x")));
    EXPECT_FALSE(static_cast<bool>(wb.set_cell_formula(0, row, col, "1+1")));
  }
  EXPECT_DOUBLE_EQ(Number(wb, 4, 4), before);
  // The header row and column, and the input cells, stay editable.
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_value(0, 2, 3, Value::number(5.0))));
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_value(0, 3, 2, Value::number(11.0))));
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 5, Value::number(1.0))));
  Recalc(wb);
  EXPECT_DOUBLE_EQ(Number(wb, 3, 3), 55.0);
}

TEST(OoxmlDataTable, SaveKeepsTheFormulaWithEveryAttribute) {
  const std::string rows =
      "<row r=\"1\"><c r=\"F1\"><v>1</v></c></row><row r=\"2\"><c r=\"F2\"><v>1</v></c></row>"
      "<row r=\"3\"><c r=\"C3\"><f>F1*F2</f><v>1</v></c><c r=\"D3\"><v>1</v></c><c r=\"E3\"><v>2</v></c></row>"
      "<row r=\"4\"><c r=\"C4\"><v>10</v></c><c r=\"D4\"><f t=\"dataTable\" ref=\"D4:E5\" dt2D=\"1\" dtr=\"1\" "
      "r1=\"F1\" r2=\"F2\" del1=\"0\" del2=\"0\" ca=\"1\" aca=\"1\"/><v>10</v></c><c r=\"E4\"><v>20</v></c></row>"
      "<row r=\"5\"><c r=\"C5\"><v>20</v></c><c r=\"D5\"><v>20</v></c><c r=\"E5\"><v>40</v></c></row>";
  Workbook wb = LoadOrDie(Package(rows));
  const std::vector<std::uint8_t> saved = SaveOrDie(wb);

  pugi::xml_document doc;
  const pugi::xml_node f = FormulaOf(doc, saved, "D4");
  ASSERT_TRUE(f);
  EXPECT_STREQ(f.attribute("t").value(), "dataTable");
  EXPECT_STREQ(f.attribute("ref").value(), "D4:E5");
  EXPECT_STREQ(f.attribute("dt2D").value(), "1");
  EXPECT_STREQ(f.attribute("dtr").value(), "1");
  EXPECT_STREQ(f.attribute("r1").value(), "F1");
  EXPECT_STREQ(f.attribute("r2").value(), "F2");
  EXPECT_STREQ(f.attribute("del1").value(), "0");
  EXPECT_STREQ(f.attribute("del2").value(), "0");
  EXPECT_STREQ(f.attribute("ca").value(), "1");
  EXPECT_STREQ(f.attribute("aca").value(), "1");
  // The other result cells are plain values.
  for (const char* ref : {"E4", "D5", "E5"}) {
    EXPECT_FALSE(FormulaOf(doc, saved, ref)) << ref;
  }

  // It is a table again after a reload, with the cached values kept.
  Workbook again = LoadOrDie(saved);
  ASSERT_EQ(again.sheet(0).data_tables().size(), 1U);
  const DataTable& t = again.sheet(0).data_tables()[0];
  EXPECT_EQ(t.ref, (MergeRange{3, 3, 4, 4}));
  EXPECT_TRUE(t.two_dimensional && t.row_oriented && t.always_calc && t.always_calc_array);
  EXPECT_DOUBLE_EQ(Number(again, 4, 4), 40.0);
  const std::vector<std::uint8_t> resaved = SaveOrDie(again);
  EXPECT_EQ(ReadPart(resaved, "xl/worksheets/sheet1.xml"), ReadPart(saved, "xl/worksheets/sheet1.xml"));
}

TEST(OoxmlDataTable, SaveAfterRecalcWritesComputedResults) {
  Workbook wb = LoadOrDie(Package(LoanRows()));
  Recalc(wb);
  const std::vector<std::uint8_t> saved = SaveOrDie(wb);
  Workbook again = LoadOrDie(saved);
  EXPECT_NEAR(Number(again, 5, 3), 860.6643, 1e-4);
  ASSERT_EQ(again.sheet(0).data_tables().size(), 1U);
  EXPECT_FALSE(again.sheet(0).data_tables()[0].two_dimensional);
  EXPECT_FALSE(again.sheet(0).data_tables()[0].row_oriented);
}

// 2-D table B2:D4 under the header row 1 / column A, inputs in F6 and F7 below it.
std::string ShiftRows() {
  return "<row r=\"1\"><c r=\"A1\"><f>F6*F7</f><v>1</v></c><c r=\"B1\"><v>1</v></c><c r=\"C1\"><v>2</v></c>"
         "<c r=\"D1\"><v>3</v></c></row>"
         "<row r=\"2\"><c r=\"A2\"><v>10</v></c><c r=\"B2\"><f t=\"dataTable\" ref=\"B2:D4\" dt2D=\"1\" dtr=\"1\" "
         "r1=\"F6\" r2=\"F7\"/><v>0</v></c><c r=\"C2\"><v>0</v></c><c r=\"D2\"><v>0</v></c></row>"
         "<row r=\"3\"><c r=\"A3\"><v>20</v></c><c r=\"B3\"><v>0</v></c><c r=\"C3\"><v>0</v></c><c "
         "r=\"D3\"><v>0</v></c>"
         "</row>"
         "<row r=\"4\"><c r=\"A4\"><v>30</v></c><c r=\"B4\"><v>0</v></c><c r=\"C4\"><v>0</v></c><c "
         "r=\"D4\"><v>0</v></c>"
         "</row>"
         "<row r=\"6\"><c r=\"F6\"><v>1</v></c></row><row r=\"7\"><c r=\"F7\"><v>1</v></c></row>";
}

TEST(OoxmlDataTable, InsertedRowsAndColumnsShiftRangeAndInputs) {
  Workbook wb = LoadOrDie(Package(ShiftRows()));
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, 0, 2)));
  const DataTable* t = wb.data_table_at(0, 4, 2);
  ASSERT_NE(t, nullptr);
  EXPECT_EQ(t->ref, (MergeRange{3, 1, 5, 3}));
  EXPECT_EQ(t->r1.row, 7U);
  EXPECT_EQ(t->r2.row, 8U);
  EXPECT_EQ(t->r1.col, 5U);
  // A column inserted between the table and the inputs moves the inputs only.
  ASSERT_TRUE(static_cast<bool>(wb.insert_cols(0, 4, 1)));
  t = wb.data_table_at(0, 4, 2);
  ASSERT_NE(t, nullptr);
  EXPECT_EQ(t->ref, (MergeRange{3, 1, 5, 3}));
  EXPECT_EQ(t->r1.col, 6U);
  EXPECT_FALSE(t->r1.deleted);
  // One inserted before the table moves the range too.
  ASSERT_TRUE(static_cast<bool>(wb.insert_cols(0, 0, 1)));
  t = wb.data_table_at(0, 4, 3);
  ASSERT_NE(t, nullptr);
  EXPECT_EQ(t->ref, (MergeRange{3, 2, 5, 4}));
  EXPECT_EQ(t->r1.col, 7U);
  // The shifted table still computes against the shifted header and inputs.
  Recalc(wb);
  EXPECT_DOUBLE_EQ(Number(wb, 3, 2), 10.0);
  EXPECT_DOUBLE_EQ(Number(wb, 5, 4), 90.0);

  const std::vector<std::uint8_t> saved = SaveOrDie(wb);
  pugi::xml_document doc;
  const pugi::xml_node f = FormulaOf(doc, saved, "C4");
  ASSERT_TRUE(f);
  EXPECT_STREQ(f.attribute("ref").value(), "C4:E6");
  EXPECT_STREQ(f.attribute("r1").value(), "H8");
  EXPECT_STREQ(f.attribute("r2").value(), "H9");
}

TEST(OoxmlDataTable, DeletingAnInputCellMarksItDeleted) {
  Workbook wb = LoadOrDie(Package(ShiftRows()));
  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, 5, 1)));  // F6, the row input.
  const DataTable* t = wb.data_table_at(0, 1, 1);
  ASSERT_NE(t, nullptr);
  EXPECT_TRUE(t->r1.deleted);
  EXPECT_FALSE(t->r2.deleted);
  EXPECT_EQ(t->r2.row, 5U);  // F7 moved up into the deleted row's place.
  EXPECT_EQ(t->ref, (MergeRange{1, 1, 3, 3}));

  const std::vector<std::uint8_t> saved = SaveOrDie(wb);
  pugi::xml_document doc;
  const pugi::xml_node f = FormulaOf(doc, saved, "B2");
  ASSERT_TRUE(f);
  EXPECT_STREQ(f.attribute("del1").value(), "1");
  EXPECT_FALSE(f.attribute("del2"));  // Absent in the source, so not written.
  EXPECT_STREQ(f.attribute("ref").value(), "B2:D4");

  // A table with a deleted input evaluates to #REF!.
  Recalc(wb);
  const Cell* result = wb.sheet(0).cell_at(2, 2);
  ASSERT_NE(result, nullptr);
  ASSERT_TRUE(result->cached_value.is_error());
  EXPECT_EQ(result->cached_value.as_error(), ErrorCode::Ref);
}

TEST(OoxmlDataTable, AnEditCuttingIntoTheTableIsRefused) {
  Workbook wb = LoadOrDie(Package(ShiftRows()));
  // The block is A1:D4: the result range B2:D4 plus its header row 1 and column A.
  const auto refused = [&](const Expected<void, Error>& r) {
    ASSERT_FALSE(static_cast<bool>(r));
    EXPECT_EQ(r.error().code, FormulonErrorCode::kPreconditionFailed);
  };
  refused(wb.delete_cols(0, 2, 1));  // Column C, inside the range.
  refused(wb.delete_rows(0, 1, 3));  // Every result row, but not the header row.
  refused(wb.delete_rows(0, 0, 1));  // The header row alone.
  refused(wb.delete_cols(0, 0, 2));  // The header column and part of the range.
  refused(wb.insert_rows(0, 1, 1));  // Between the header row and the results.
  refused(wb.insert_cols(0, 3, 1));  // Between two result columns.
  const DataTable* t = wb.data_table_at(0, 1, 1);
  ASSERT_NE(t, nullptr);
  EXPECT_EQ(t->ref, (MergeRange{1, 1, 3, 3}));
  EXPECT_DOUBLE_EQ(Number(wb, 0, 2), 2.0);  // Nothing moved.
}

TEST(OoxmlDataTable, DeletingTheWholeTableRemovesIt) {
  Workbook rows = LoadOrDie(Package(ShiftRows()));
  ASSERT_TRUE(static_cast<bool>(rows.delete_rows(0, 0, 4)));  // Header row and every result row.
  EXPECT_TRUE(rows.sheet(0).data_tables().empty());
  Workbook cols = LoadOrDie(Package(ShiftRows()));
  ASSERT_TRUE(static_cast<bool>(cols.delete_cols(0, 0, 5)));  // Header column through past the range.
  EXPECT_TRUE(cols.sheet(0).data_tables().empty());
  ASSERT_TRUE(static_cast<bool>(cols.set_cell_value(0, 1, 0, Value::number(1.0))));  // No longer protected.
}

// Every result cell reports `=TABLE(row_input,col_input)`; FORMULATEXT braces it like a legacy array.
TEST(OoxmlDataTable, ResultCellsShowTheTableFormula) {
  Workbook wb = LoadOrDie(Package(ShiftRows()));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 9, 0, "=FORMULATEXT(C3)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 9, 1, "=ISFORMULA(D4)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 9, 2, "=ISFORMULA(A2)")));
  Recalc(wb);
  EXPECT_EQ(data_table_formula_text(*wb.data_table_at(0, 2, 2)), "=TABLE(F6,F7)");
  const Cell* text = wb.sheet(0).cell_at(9, 0);
  ASSERT_NE(text, nullptr);
  ASSERT_TRUE(text->cached_value.is_text());
  EXPECT_EQ(text->cached_value.as_text(), "{=TABLE(F6,F7)}");
  EXPECT_TRUE(wb.sheet(0).cell_at(9, 1)->cached_value.as_boolean());
  EXPECT_FALSE(wb.sheet(0).cell_at(9, 2)->cached_value.as_boolean());
}

TEST(OoxmlDataTable, AutoNoTableLeavesTablesToTheExplicitCalculation) {
  Workbook wb = LoadOrDie(Package(MultiplicationRows(), "<calcPr calcMode=\"autoNoTable\"/>"));
  ASSERT_EQ(wb.calc_mode(), CalcMode::kAutoNoTable);
  Recalc(wb);
  EXPECT_DOUBLE_EQ(Number(wb, 3, 3), 0.0);  // Cached value, not recomputed.
  EXPECT_DOUBLE_EQ(Number(wb, 2, 2), 1.0);  // Ordinary formulas still calculate.
  auto explicit_calc = wb.recalc_data_tables(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(explicit_calc)) << explicit_calc.error().message;
  EXPECT_DOUBLE_EQ(Number(wb, 3, 3), 10.0);
  EXPECT_DOUBLE_EQ(Number(wb, 5, 5), 90.0);
  EXPECT_DOUBLE_EQ(Number(wb, 0, 5), 1.0);
  // Switching back to automatic picks the tables up on the next recalc.
  wb.set_calc_mode(CalcMode::kAuto);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 2, 3, Value::number(7.0))));
  Recalc(wb);
  EXPECT_DOUBLE_EQ(Number(wb, 3, 3), 70.0);
}

TEST(OoxmlDataTable, AddAndRemoveBuildTheTableStructure) {
  Workbook wb =
      LoadOrDie(Package("<row r=\"1\"><c r=\"F1\"><v>3</v></c></row>"
                        "<row r=\"2\"><c r=\"B2\"><f>F1*F1</f><v>0</v></c></row>"
                        "<row r=\"3\"><c r=\"A3\"><v>1</v></c></row><row r=\"4\"><c r=\"A4\"><v>2</v></c></row>"));
  DataTableInput col_input;
  col_input.present = true;
  col_input.row = 0;
  col_input.col = 5;
  // Validation first: no input, no header room, input inside the range.
  EXPECT_EQ(wb.add_data_table(0, MergeRange{2, 1, 3, 1}, DataTableInput{}, DataTableInput{}).error().code,
            FormulonErrorCode::kDataTableInvalid);
  EXPECT_EQ(wb.add_data_table(0, MergeRange{0, 1, 3, 1}, DataTableInput{}, col_input).error().code,
            FormulonErrorCode::kDataTableInvalid);
  DataTableInput inside = col_input;
  inside.row = 2;
  inside.col = 1;
  EXPECT_EQ(wb.add_data_table(0, MergeRange{2, 1, 3, 1}, DataTableInput{}, inside).error().code,
            FormulonErrorCode::kDataTableInvalid);
  EXPECT_EQ(wb.add_data_table(1, MergeRange{2, 1, 3, 1}, DataTableInput{}, col_input).error().code,
            FormulonErrorCode::kInvalidArgument);

  ASSERT_TRUE(static_cast<bool>(wb.add_data_table(0, MergeRange{2, 1, 3, 1}, DataTableInput{}, col_input)));
  EXPECT_EQ(wb.add_data_table(0, MergeRange{3, 1, 3, 1}, DataTableInput{}, col_input).error().code,
            FormulonErrorCode::kDataTableInvalid);  // Overlap.
  const DataTable* t = wb.data_table_at(0, 3, 1);
  ASSERT_NE(t, nullptr);
  EXPECT_FALSE(t->two_dimensional);
  EXPECT_FALSE(t->row_oriented);
  EXPECT_EQ(t->r1.col, 5U);
  Recalc(wb);
  EXPECT_DOUBLE_EQ(Number(wb, 2, 1), 1.0);
  EXPECT_DOUBLE_EQ(Number(wb, 3, 1), 4.0);
  EXPECT_DOUBLE_EQ(Number(wb, 0, 5), 3.0);
  EXPECT_FALSE(static_cast<bool>(wb.set_cell_value(0, 3, 1, Value::number(1.0))));

  // It saves as a data table with the explicit attribute set.
  const std::vector<std::uint8_t> saved = SaveOrDie(wb);
  pugi::xml_document doc;
  const pugi::xml_node f = FormulaOf(doc, saved, "B3");
  ASSERT_TRUE(f);
  EXPECT_STREQ(f.attribute("ref").value(), "B3:B4");
  EXPECT_STREQ(f.attribute("r1").value(), "F1");
  EXPECT_STREQ(f.attribute("dt2D").value(), "0");
  EXPECT_STREQ(f.attribute("dtr").value(), "0");

  ASSERT_TRUE(static_cast<bool>(wb.remove_data_table(0, 3, 1)));
  EXPECT_TRUE(wb.sheet(0).data_tables().empty());
  EXPECT_TRUE(wb.sheet(0).cell_at(2, 1)->cached_value.is_blank());
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_value(0, 3, 1, Value::number(1.0))));
  EXPECT_EQ(wb.remove_data_table(0, 3, 1).error().code, FormulonErrorCode::kDataTableInvalid);
}

TEST(OoxmlDataTable, FormulaInputCellIsRestored) {
  // The input B1 holds a formula; the table substitutes into it and puts it back.
  const std::string rows =
      "<row r=\"1\"><c r=\"A1\"><v>5</v></c><c r=\"B1\"><f>A1+1</f><v>6</v></c></row>"
      "<row r=\"2\"><c r=\"B2\"><f>B1*2</f><v>0</v></c></row>"
      "<row r=\"3\"><c r=\"A3\"><v>10</v></c><c r=\"B3\"><f t=\"dataTable\" ref=\"B3:B4\" dt2D=\"0\" dtr=\"0\" "
      "r1=\"B1\"/><v>0</v></c></row>"
      "<row r=\"4\"><c r=\"A4\"><v>20</v></c><c r=\"B4\"><v>0</v></c></row>";
  Workbook wb = LoadOrDie(Package(rows));
  Recalc(wb);
  EXPECT_DOUBLE_EQ(Number(wb, 2, 1), 20.0);
  EXPECT_DOUBLE_EQ(Number(wb, 3, 1), 40.0);
  const Cell* input = wb.sheet(0).cell_at(0, 1);
  ASSERT_NE(input, nullptr);
  EXPECT_EQ(input->formula_text, "=A1+1");
  EXPECT_DOUBLE_EQ(Number(wb, 0, 1), 6.0);
  EXPECT_DOUBLE_EQ(Number(wb, 1, 1), 12.0);
}

}  // namespace
}  // namespace formulon
