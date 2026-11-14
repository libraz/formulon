//
// Stale formulas through `.xlsx`, as Excel (Mac 16.115) saves and opens them
// under manual calculation: a save writes `<f ca="1">` on every stale formula
// (and keeps it on volatiles) and `calcCompleted="0"` while any is left; an
// open treats exactly the `ca="1"` formulas as stale, whatever `calcCompleted`
// says, keeping every other cached value. A formula without a cached value
// is computed on the first recalc. The Excel-saved fixtures live in
// `tests/fixtures/excel/stale/`.

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>
#include <string_view>
#include <unordered_set>
#include <utility>
#include <vector>

#include "cell.h"
#include "eval/dep_graph.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/ooxml/zip_part_writer.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

using CellSet = std::unordered_set<eval::CellNodeId, eval::CellNodeIdHash>;

std::vector<std::uint8_t> ReadFixture(const std::string& name) {
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/stale/" + name;
  std::vector<std::uint8_t> out;
  FILE* file = std::fopen(path.c_str(), "rb");
  if (file == nullptr) {
    ADD_FAILURE() << "could not open fixture: " << path;
    return out;
  }
  std::fseek(file, 0, SEEK_END);
  const long size = std::ftell(file);
  std::fseek(file, 0, SEEK_SET);
  out.resize(size > 0 ? static_cast<std::size_t>(size) : 0U);
  if (std::fread(out.data(), 1, out.size(), file) != out.size()) {
    ADD_FAILURE() << "short read on fixture: " << path;
  }
  std::fclose(file);
  return out;
}

Workbook Load(const std::vector<std::uint8_t>& bytes, bool sax = false) {
  const io::ByteSpan span{bytes.data(), bytes.size()};
  auto result = sax ? io::internal::ReadOoxmlWithSaxThresholdForTesting(span, 0U) : io::read_ooxml(span);
  EXPECT_TRUE(static_cast<bool>(result)) << (result ? "" : result.error().message);
  return result ? std::move(result.value().workbook) : Workbook::create_empty();
}

std::vector<std::uint8_t> Save(const Workbook& wb) {
  auto saved = io::write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
  return saved ? std::move(saved.value()) : std::vector<std::uint8_t>{};
}

Workbook LoadXlsb(const std::vector<std::uint8_t>& bytes) {
  auto result = io::xlsb::read_xlsb(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(result)) << (result ? "" : result.error().message);
  return result ? std::move(result.value().workbook) : Workbook::create_empty();
}

std::vector<std::uint8_t> SaveXlsb(const Workbook& wb) {
  auto saved = io::xlsb::write_xlsb(wb);
  EXPECT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
  return saved ? std::move(saved.value()) : std::vector<std::uint8_t>{};
}

// The formula cells of `wb` whose value is stale.
CellSet StaleFormulas(const Workbook& wb) {
  CellSet out;
  for (const eval::CellNodeId& c : wb.recalc_engine().stale_cells()) {
    const Cell* cell = wb.sheet(c.sheet_id).cell_at(c.row, c.col);
    if (cell != nullptr && !cell->formula_text.empty()) {
      out.insert(c);
    }
  }
  return out;
}

std::string Part(const std::vector<std::uint8_t>& bytes, const char* path) {
  io::ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(io::ByteSpan{bytes.data(), bytes.size()})));
  auto part = zip.read_entry(path);
  EXPECT_TRUE(static_cast<bool>(part)) << path;
  return part ? std::string(part.value().begin(), part.value().end()) : std::string();
}

// The opening `<f ...>` tag of the cell `ref` in `sheet_xml`, or empty.
std::string FormulaTag(const std::string& sheet_xml, const std::string& ref) {
  const std::size_t cell = sheet_xml.find("<c r=\"" + ref + "\"");
  if (cell == std::string::npos) {
    return {};
  }
  const std::size_t close = sheet_xml.find("</c>", cell);
  const std::size_t f = sheet_xml.find("<f", cell);
  if (f == std::string::npos || f > close) {
    return {};
  }
  return sheet_xml.substr(f, sheet_xml.find('>', f) - f + 1U);
}

bool MarkedCa(const std::string& sheet_xml, const std::string& ref) {
  return FormulaTag(sheet_xml, ref).find("ca=\"1\"") != std::string::npos;
}

std::string CalcPr(const std::vector<std::uint8_t>& bytes) {
  const std::string xml = Part(bytes, "xl/workbook.xml");
  const std::size_t at = xml.find("<calcPr");
  return at == std::string::npos ? std::string() : xml.substr(at, xml.find('>', at) - at + 1U);
}

eval::CellNodeId Node(std::uint16_t sheet, std::uint32_t row, std::uint32_t col) {
  return eval::CellNodeId{sheet, row, col};
}

double Number(const Workbook& wb, std::size_t sheet, std::uint32_t row, std::uint32_t col) {
  const Value v = wb.sheet(sheet).resolve_cell_value(row, col);
  EXPECT_TRUE(v.is_number()) << "sheet " << sheet << " row " << row << " col " << col;
  return v.is_number() ? v.as_number() : -1.0;
}

// A manual book: A1 = 1, B1 = A1*2, C1 = B1+1, D1 = 5+1, E1 = RAND(), F1 = E1+1, recalculated.
Workbook ManualChain() {
  Workbook wb = Workbook::create();
  wb.set_calc_mode(CalcMode::kManual);
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 0, Value::number(1.0))));
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 1, "=A1*2")));
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 2, "=B1+1")));
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 3, "=5+1")));
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 4, "=RAND()")));
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 5, "=E1+1")));
  EXPECT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  return wb;
}

// A one-sheet package with `<calcPr calc_pr/>` (none when empty) and `sheet_data` as its cells.
std::vector<std::uint8_t> Package(std::string_view calc_pr, std::string_view sheet_data) {
  const std::string decl = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
  const std::string rels_ns = "http://schemas.openxmlformats.org/package/2006/relationships";
  const std::string doc_rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  io::ZipWriterGuard writer;
  EXPECT_TRUE(writer.init());
  const auto add = [&writer](const char* path, const std::string& body) {
    EXPECT_TRUE(static_cast<bool>(io::AddPart(writer.get(), path, body)));
  };
  add("[Content_Types].xml",
      decl +
          "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
          "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
          "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
          "<Override PartName=\"/xl/workbook.xml\" "
          "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
          "<Override PartName=\"/xl/worksheets/sheet1.xml\" "
          "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/></Types>");
  add("_rels/.rels", decl + "<Relationships xmlns=\"" + rels_ns + "\"><Relationship Id=\"rId1\" Type=\"" + doc_rel +
                         "/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
  add("xl/workbook.xml", decl +
                             "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
                             "xmlns:r=\"" +
                             doc_rel + "\"><sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>" +
                             (calc_pr.empty() ? std::string() : "<calcPr " + std::string(calc_pr) + "/>") +
                             "</workbook>");
  add("xl/_rels/workbook.xml.rels", decl + "<Relationships xmlns=\"" + rels_ns +
                                        "\"><Relationship Id=\"rId1\" Type=\"" + doc_rel +
                                        "/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>");
  add("xl/worksheets/sheet1.xml", decl +
                                      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                                      "<sheetData>" +
                                      std::string(sheet_data) + "</sheetData></worksheet>");
  auto out = io::FinalizeArchive(writer, "stale_roundtrip_test");
  EXPECT_TRUE(static_cast<bool>(out));
  return out ? std::move(out.value()) : std::vector<std::uint8_t>{};
}

// A1 = 2; B1 = A1*10 cached 7 (clean); C1 = A1+1 with no cached value; D1 = A1*2 cached 0, marked ca.
constexpr std::string_view kCachedRow =
    "<row r=\"1\"><c r=\"A1\"><v>2</v></c><c r=\"B1\"><f>A1*10</f><v>7</v></c><c r=\"C1\"><f>A1+1</f></c>"
    "<c r=\"D1\"><f ca=\"1\">A1*2</f><v>0</v></c></row>";

}  // namespace

// An edit leaves its dependents and every volatile with theirs stale; the save marks exactly those.
TEST(StaleRoundtrip, ManualEditSavesCaOnTheStaleFormulas) {
  Workbook wb = ManualChain();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 0, Value::number(2.0))));
  const std::vector<std::uint8_t> bytes = Save(wb);
  const std::string sheet = Part(bytes, "xl/worksheets/sheet1.xml");
  EXPECT_TRUE(MarkedCa(sheet, "B1"));
  EXPECT_TRUE(MarkedCa(sheet, "C1"));
  EXPECT_TRUE(MarkedCa(sheet, "E1"));
  EXPECT_TRUE(MarkedCa(sheet, "F1"));
  EXPECT_EQ(FormulaTag(sheet, "D1"), "<f>");
  const std::string calc_pr = CalcPr(bytes);
  EXPECT_NE(calc_pr.find("calcMode=\"manual\""), std::string::npos) << calc_pr;
  EXPECT_NE(calc_pr.find("calcCompleted=\"0\""), std::string::npos) << calc_pr;
}

// After a full recalc nothing is stale: only the volatile keeps ca, and calcCompleted is the stored setting.
TEST(StaleRoundtrip, FullRecalcLeavesCaOnVolatileChains) {
  Workbook wb = ManualChain();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 0, Value::number(2.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_TRUE(wb.recalc_engine().stale_cells().empty());
  const std::vector<std::uint8_t> bytes = Save(wb);
  const std::string sheet = Part(bytes, "xl/worksheets/sheet1.xml");
  EXPECT_EQ(FormulaTag(sheet, "B1"), "<f>");
  EXPECT_EQ(FormulaTag(sheet, "C1"), "<f>");
  EXPECT_EQ(FormulaTag(sheet, "D1"), "<f>");
  EXPECT_TRUE(MarkedCa(sheet, "E1"));
  EXPECT_TRUE(MarkedCa(sheet, "F1"));  // reads the volatile E1, as Excel keeps L3 = H8+1 marked
  EXPECT_EQ(CalcPr(bytes).find("calcCompleted"), std::string::npos);

  wb.mutable_calc_options().calc_completed = false;
  EXPECT_NE(CalcPr(Save(wb)).find("calcCompleted=\"0\""), std::string::npos);
}

// Automatic calculation never writes a formula as stale, computed or not; a volatile chain is still marked.
TEST(StaleRoundtrip, AutomaticSaveWritesNoStaleMarks) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 0, Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 1, "=A1*2")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 2, "=NOW()")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 3, "=C1+1")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 4, "=D1*2")));
  const std::vector<std::uint8_t> bytes = Save(wb);
  const std::string sheet = Part(bytes, "xl/worksheets/sheet1.xml");
  EXPECT_EQ(FormulaTag(sheet, "B1"), "<f>");
  EXPECT_TRUE(MarkedCa(sheet, "C1"));
  EXPECT_TRUE(MarkedCa(sheet, "D1"));
  EXPECT_TRUE(MarkedCa(sheet, "E1"));  // reads the volatile transitively
  EXPECT_EQ(CalcPr(bytes).find("calcCompleted"), std::string::npos);
}

// Excel's K control: st_stale.xlsx recalculated and saved keeps ca on the volatiles and what reads them only.
TEST(StaleRoundtrip, ExcelBookRecalculatedKeepsCaOnItsVolatileChain) {
  Workbook wb = Load(ReadFixture("st_stale.xlsx"));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  const std::vector<std::uint8_t> bytes = Save(wb);
  const std::string sheet1 = Part(bytes, "xl/worksheets/sheet1.xml");
  for (const char* ref : {"H8", "I8", "J8", "L3"}) {
    EXPECT_TRUE(MarkedCa(sheet1, ref)) << ref;
  }
  for (const char* ref : {"I2", "J2", "K2", "L2", "I5", "C5"}) {
    EXPECT_FALSE(MarkedCa(sheet1, ref)) << ref;
  }
  EXPECT_FALSE(MarkedCa(Part(bytes, "xl/worksheets/sheet2.xml"), "A1"));
  // The file's calcCompleted="0" no longer holds once the recalc left nothing to compute.
  EXPECT_EQ(CalcPr(bytes).find("calcCompleted"), std::string::npos);
}

// Excel's own save (st_stale.xlsx): the ca formulas open stale with their cached values, the rest clean.
TEST(StaleRoundtrip, ExcelSavedStaleBookOpensWithItsCaFormulasStale) {
  for (const bool sax : {false, true}) {
    SCOPED_TRACE(sax ? "sax" : "dom");
    Workbook wb = Load(ReadFixture("st_stale.xlsx"), sax);
    ASSERT_EQ(wb.calc_mode(), CalcMode::kManual);
    const CellSet expected = {Node(0, 1, 8), Node(0, 1, 9), Node(0, 1, 10), Node(0, 2, 11),
                              Node(0, 7, 7), Node(0, 7, 8), Node(0, 7, 9),  Node(1, 0, 0)};
    const CellSet stale = wb.recalc_engine().stale_cells();
    for (const eval::CellNodeId& c : expected) {
      EXPECT_EQ(stale.count(c), 1U) << c.sheet_id << "!" << c.row << "," << c.col;
      EXPECT_TRUE(wb.recalc_engine().awaits_recalc(c));
    }
    for (const eval::CellNodeId& c : {Node(0, 1, 11), Node(0, 4, 2), Node(0, 4, 8), Node(1, 2, 0)}) {
      EXPECT_EQ(stale.count(c), 0U) << c.sheet_id << "!" << c.row << "," << c.col;
      EXPECT_FALSE(wb.recalc_engine().awaits_recalc(c));
    }
    EXPECT_EQ(Number(wb, 0, 1, 8), 4.0);  // I2 keeps the cached value until a recalc
    EXPECT_EQ(Number(wb, 1, 0, 0), 4.0);

    auto stats = wb.recalc(eval::default_registry());
    ASSERT_TRUE(static_cast<bool>(stats));
    EXPECT_EQ(stats.value().cells_evaluated, expected.size());
    EXPECT_EQ(Number(wb, 0, 1, 8), 6.0);    // I2 = H2*2
    EXPECT_EQ(Number(wb, 0, 1, 9), 7.0);    // J2 = I2+1
    EXPECT_EQ(Number(wb, 0, 1, 10), 13.0);  // K2 = H2+H5
    EXPECT_EQ(Number(wb, 1, 0, 0), 6.0);    // Sheet2!A1 = Sheet1!I2
    EXPECT_TRUE(wb.recalc_engine().stale_cells().empty());
  }
}

// The open-time stale set follows ca, not calcCompleted (st_ca.xlsx: ca on the clean I5, calcCompleted="1").
TEST(StaleRoundtrip, OpenTimeStaleSetFollowsCaNotCalcCompleted) {
  Workbook wb = Load(ReadFixture("st_ca.xlsx"));
  ASSERT_EQ(wb.calc_mode(), CalcMode::kManual);
  EXPECT_TRUE(wb.recalc_engine().awaits_recalc(Node(0, 4, 8)));    // I5
  EXPECT_FALSE(wb.recalc_engine().awaits_recalc(Node(0, 1, 11)));  // L2
  EXPECT_FALSE(wb.recalc_engine().awaits_recalc(Node(1, 2, 0)));   // Sheet2!A3
}

// Saved stale, opened stale: the stale set survives a round trip through the file.
TEST(StaleRoundtrip, StaleFormulasSurviveASaveAndReopen) {
  Workbook wb = ManualChain();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 0, Value::number(2.0))));
  const CellSet before = wb.recalc_engine().stale_cells();
  Workbook loaded = Load(Save(wb));
  ASSERT_EQ(loaded.calc_mode(), CalcMode::kManual);
  CellSet formulas;
  for (const eval::CellNodeId& c : loaded.recalc_engine().stale_cells()) {
    const Cell* cell = loaded.sheet(c.sheet_id).cell_at(c.row, c.col);
    if (cell != nullptr && !cell->formula_text.empty()) {
      formulas.insert(c);
    }
  }
  const CellSet expected = {Node(0, 0, 1), Node(0, 0, 2), Node(0, 0, 4), Node(0, 0, 5)};
  EXPECT_EQ(formulas, expected);
  for (const eval::CellNodeId& c : expected) {
    EXPECT_EQ(before.count(c), 1U);
  }
  EXPECT_FALSE(loaded.recalc_engine().awaits_recalc(Node(0, 0, 3)));
}

// A clean formula keeps its cached value; a formula without one and a ca formula are computed by the first recalc.
TEST(StaleRoundtrip, FirstRecalcComputesUncachedAndCaFormulasOnly) {
  for (const bool sax : {false, true}) {
    SCOPED_TRACE(sax ? "sax" : "dom");
    Workbook wb = Load(Package("calcMode=\"manual\"", kCachedRow), sax);
    ASSERT_EQ(wb.calc_mode(), CalcMode::kManual);
    EXPECT_FALSE(wb.recalc_engine().awaits_recalc(Node(0, 0, 1)));
    EXPECT_TRUE(wb.recalc_engine().awaits_recalc(Node(0, 0, 2)));
    EXPECT_TRUE(wb.recalc_engine().awaits_recalc(Node(0, 0, 3)));
    ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    EXPECT_EQ(Number(wb, 0, 0, 1), 7.0);
    EXPECT_EQ(Number(wb, 0, 0, 2), 3.0);
    EXPECT_EQ(Number(wb, 0, 0, 3), 4.0);
  }
}

// Under automatic calculation a loaded workbook still recomputes every formula on its first recalc.
TEST(StaleRoundtrip, AutomaticOpenStillRecomputesEveryFormula) {
  Workbook wb = Load(Package("", kCachedRow));
  ASSERT_EQ(wb.calc_mode(), CalcMode::kAuto);
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_EQ(Number(wb, 0, 0, 1), 20.0);
  EXPECT_EQ(Number(wb, 0, 0, 2), 3.0);
  EXPECT_EQ(Number(wb, 0, 0, 3), 4.0);
}

// An .xlsb save carries the stale set as fAlwaysCalc and the book as fSomeUncalced, and reopens with it.
TEST(StaleRoundtrip, XlsbSaveKeepsTheStaleSetAndSomeUncalced) {
  Workbook wb = ManualChain();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 0, Value::number(2.0))));
  Workbook loaded = LoadXlsb(SaveXlsb(wb));
  ASSERT_EQ(loaded.calc_mode(), CalcMode::kManual);
  EXPECT_FALSE(loaded.calc_options().calc_completed);
  const CellSet expected = {Node(0, 0, 1), Node(0, 0, 2), Node(0, 0, 4), Node(0, 0, 5)};
  EXPECT_EQ(StaleFormulas(loaded), expected);
  EXPECT_FALSE(loaded.recalc_engine().awaits_recalc(Node(0, 0, 3)));
  EXPECT_EQ(Number(loaded, 0, 0, 2), 3.0);  // C1 keeps its cached value until a recalc
}

// After a full recalc the .xlsb save is complete and only the volatile chain keeps fAlwaysCalc.
TEST(StaleRoundtrip, XlsbSaveAfterAFullRecalcMarksOnlyTheVolatileChain) {
  Workbook wb = ManualChain();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 0, Value::number(2.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  Workbook loaded = LoadXlsb(SaveXlsb(wb));
  EXPECT_TRUE(loaded.calc_options().calc_completed);
  const CellSet expected = {Node(0, 0, 4), Node(0, 0, 5)};
  EXPECT_EQ(StaleFormulas(loaded), expected);
}

// Excel's stale book (st_stale.xlsx) saved as .xlsb reopens with the stale set Excel opened it with.
TEST(StaleRoundtrip, ExcelStaleBookThroughXlsbKeepsItsStaleSet) {
  Workbook wb = Load(ReadFixture("st_stale.xlsx"));
  const CellSet before = StaleFormulas(wb);
  ASSERT_FALSE(before.empty());
  Workbook loaded = LoadXlsb(SaveXlsb(wb));
  ASSERT_EQ(loaded.calc_mode(), CalcMode::kManual);
  EXPECT_FALSE(loaded.calc_options().calc_completed);
  EXPECT_EQ(StaleFormulas(loaded), before);
}

// Excel's own .xlsb save of the stale book: the fAlwaysCalc formulas open stale and fSomeUncalced reads back.
TEST(StaleRoundtrip, ExcelSavedStaleXlsbOpensWithItsAlwaysCalcFormulasStale) {
  Workbook wb = LoadXlsb(ReadFixture("st_stale.xlsb"));
  ASSERT_EQ(wb.calc_mode(), CalcMode::kManual);
  EXPECT_FALSE(wb.calc_options().calc_completed);
  const CellSet expected = {Node(0, 1, 8), Node(0, 1, 9), Node(0, 1, 10), Node(0, 2, 11),
                            Node(0, 7, 7), Node(0, 7, 8), Node(0, 7, 9),  Node(1, 0, 0)};
  EXPECT_EQ(StaleFormulas(wb), expected);
  EXPECT_EQ(Number(wb, 0, 1, 8), 4.0);  // I2 keeps the cached value until a recalc
}

// Excel recalculated the stale .xlsb and saved it: only the volatile chain keeps fAlwaysCalc; our recalc agrees.
TEST(StaleRoundtrip, XlsbRecalculatedMatchesExcelsRecalculatedSave) {
  Workbook excel = LoadXlsb(ReadFixture("st_stale_calc.xlsb"));
  EXPECT_TRUE(excel.calc_options().calc_completed);
  const CellSet volatile_chain = {Node(0, 2, 11), Node(0, 7, 7), Node(0, 7, 8), Node(0, 7, 9)};
  EXPECT_EQ(StaleFormulas(excel), volatile_chain);

  Workbook wb = LoadXlsb(ReadFixture("st_stale.xlsb"));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  Workbook loaded = LoadXlsb(SaveXlsb(wb));
  EXPECT_TRUE(loaded.calc_options().calc_completed);
  EXPECT_EQ(StaleFormulas(loaded), volatile_chain);
}

}  // namespace formulon
