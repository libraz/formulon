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
#include <optional>
#include <string>
#include <string_view>
#include <unordered_set>
#include <utility>
#include <vector>

#include "cell.h"
#include "eval/dep_graph.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "eval/scheduler.h"
#include "excel_profile.h"
#include "gtest/gtest-spi.h"
#include "gtest/gtest.h"
#include "io/ooxml/zip_part_writer.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "pugixml.hpp"
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

Workbook Load(const std::vector<std::uint8_t>& bytes, bool sax = false,
              std::optional<ExcelProfile> profile = std::nullopt) {
  const io::ByteSpan span{bytes.data(), bytes.size()};
  io::ReadOptions options;
  options.profile = profile;
  auto result = sax ? io::internal::ReadOoxmlWithSaxThresholdForTesting(span, 0U) : io::read_ooxml(span, options);
  EXPECT_TRUE(static_cast<bool>(result)) << (result ? "" : result.error().message);
  return result ? std::move(result.value().workbook) : Workbook::create_empty();
}

std::vector<std::uint8_t> Save(const Workbook& wb) {
  auto saved = io::write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
  return saved ? std::move(saved.value()) : std::vector<std::uint8_t>{};
}

Workbook LoadXlsb(const std::vector<std::uint8_t>& bytes, std::optional<ExcelProfile> profile = std::nullopt) {
  io::ReadOptions options;
  options.profile = profile;
  auto result = io::xlsb::read_xlsb(io::ByteSpan{bytes.data(), bytes.size()}, options);
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
std::vector<std::uint8_t> Package(std::string_view calc_pr, std::string_view sheet_data,
                                  std::string_view sheet_name = "Sheet1", std::string_view metadata_xml = {}) {
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
                             doc_rel + "\"><sheets><sheet name=\"" + std::string(sheet_name) +
                             "\" sheetId=\"1\" r:id=\"rId1\"/></sheets>" +
                             (calc_pr.empty() ? std::string() : "<calcPr " + std::string(calc_pr) + "/>") +
                             "</workbook>");
  add("xl/_rels/workbook.xml.rels", decl + "<Relationships xmlns=\"" + rels_ns +
                                        "\"><Relationship Id=\"rId1\" Type=\"" + doc_rel +
                                        "/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>");
  add("xl/worksheets/sheet1.xml", decl +
                                      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                                      "<sheetData>" +
                                      std::string(sheet_data) + "</sheetData></worksheet>");
  if (!metadata_xml.empty()) {
    add("xl/metadata.xml", decl + std::string(metadata_xml));
  }
  auto out = io::FinalizeArchive(writer, "stale_roundtrip_test");
  EXPECT_TRUE(static_cast<bool>(out));
  return out ? std::move(out.value()) : std::vector<std::uint8_t>{};
}

// `xl/metadata.xml` whose cell metadata 1 is the dynamic-array (XLDAPR) mark a `cm="1"` cell carries.
constexpr std::string_view kDynamicArrayMetadata =
    "<metadata xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
    "xmlns:xda=\"http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray\">"
    "<metadataTypes count=\"1\"><metadataType name=\"XLDAPR\" minSupportedVersion=\"120000\" copy=\"1\" "
    "pasteAll=\"1\" pasteValues=\"1\" merge=\"1\" splitFirst=\"1\" rowColShift=\"1\" clearFormats=\"1\" "
    "clearComments=\"1\" assign=\"1\" coerce=\"1\" cellMeta=\"1\"/></metadataTypes>"
    "<futureMetadata name=\"XLDAPR\" count=\"1\"><bk><extLst><ext uri=\"{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}\">"
    "<xda:dynamicArrayProperties fDynamic=\"1\" fCollapsed=\"0\"/></ext></extLst></bk></futureMetadata>"
    "<cellMetadata count=\"1\"><bk><rc t=\"1\" v=\"0\"/></bk></cellMetadata></metadata>";

// A1 = 2; B1 = A1*10 cached 7 (clean); C1 = A1+1 with no cached value; D1 = A1*2 cached 0, marked ca.
constexpr std::string_view kCachedRow =
    "<row r=\"1\"><c r=\"A1\"><v>2</v></c><c r=\"B1\"><f>A1*10</f><v>7</v></c><c r=\"C1\"><f>A1+1</f></c>"
    "<c r=\"D1\"><f ca=\"1\">A1*2</f><v>0</v></c></row>";

// The cached value of every formula cell of the first sheet, row-major, as text; a blank cache reads `blank`.
std::vector<std::string> CachedFormulaValues(const Workbook& wb) {
  std::vector<std::string> out;
  for (const auto& [row, cells] : wb.sheet(0).rows()) {
    for (const auto& [col, cell] : cells) {
      if (cell.formula_text.empty()) {
        continue;
      }
      out.push_back(std::to_string(row) + "," + std::to_string(col) + "=" +
                    (cell.cached_value.is_number() ? std::to_string(cell.cached_value.as_number()) : "blank"));
    }
  }
  return out;
}

bool SameProfile(ExcelProfile a, ExcelProfile b) {
  return a.host == b.host && a.locale == b.locale;
}

// A1 = 2; B1 = A1*10 cached 7 and C1 = A1+1 cached 3, both clean; D1 = A1*2 cached 0, marked ca.
constexpr std::string_view kOneCaRow =
    "<row r=\"1\"><c r=\"A1\"><v>2</v></c><c r=\"B1\"><f>A1*10</f><v>7</v></c><c r=\"C1\"><f>A1+1</f><v>3</v></c>"
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

// An automatic book stamped with the current calcId opens with its cache, as a manual one does: only the uncached
// and ca formulas compute.
TEST(StaleRoundtrip, AutomaticBookWithTheCurrentCalcIdKeepsItsCache) {
  for (const bool sax : {false, true}) {
    SCOPED_TRACE(sax ? "sax" : "dom");
    Workbook wb = Load(Package("calcId=\"191029\"", kCachedRow), sax);
    ASSERT_EQ(wb.calc_mode(), CalcMode::kAuto);
    EXPECT_FALSE(wb.opens_with_recalc());
    EXPECT_FALSE(wb.recalc_engine().awaits_recalc(Node(0, 0, 1)));
    EXPECT_TRUE(wb.recalc_engine().awaits_recalc(Node(0, 0, 2)));
    EXPECT_TRUE(wb.recalc_engine().awaits_recalc(Node(0, 0, 3)));
    ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    EXPECT_EQ(Number(wb, 0, 0, 1), 7.0);
    EXPECT_EQ(Number(wb, 0, 0, 2), 3.0);
    EXPECT_EQ(Number(wb, 0, 0, 3), 4.0);
  }
}

// fullCalcOnLoad, an older calcId or none makes an automatic book recompute every formula on open.
TEST(StaleRoundtrip, AutomaticBookRecalculatesOnFullCalcOnLoadOrAnotherCalcId) {
  for (const char* calc_pr : {"calcId=\"191029\" fullCalcOnLoad=\"1\"", "calcId=\"124519\"", ""}) {
    SCOPED_TRACE(calc_pr);
    Workbook wb = Load(Package(calc_pr, kCachedRow));
    EXPECT_TRUE(wb.opens_with_recalc());
    ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    EXPECT_EQ(Number(wb, 0, 0, 1), 20.0);
  }
}

// Manual calculation keeps the cache whatever calcId and fullCalcOnLoad say.
TEST(StaleRoundtrip, ManualBookIgnoresFullCalcOnLoad) {
  Workbook wb = Load(Package("calcId=\"124519\" calcMode=\"manual\" fullCalcOnLoad=\"1\"", kCachedRow));
  EXPECT_FALSE(wb.opens_with_recalc());
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_EQ(Number(wb, 0, 0, 1), 7.0);
}

// The xlsb reader takes the same decision from BrtCalcProp.
TEST(StaleRoundtrip, XlsbAutomaticBookWithTheCurrentCalcIdKeepsItsCache) {
  Workbook source = Load(Package("calcId=\"191029\"", kCachedRow));
  Workbook wb = LoadXlsb(SaveXlsb(source));
  EXPECT_FALSE(wb.opens_with_recalc());
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_EQ(Number(wb, 0, 0, 1), 7.0);
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

// A profile given at load leaves the stale set and the cached values exactly as the default profile has them.
TEST(StaleRoundtrip, ProfileAtLoadKeepsTheManualStaleSetXlsx) {
  const std::vector<std::vector<std::uint8_t>> books = {Package("calcMode=\"manual\"", kCachedRow),
                                                        ReadFixture("st_stale.xlsx")};
  for (const std::vector<std::uint8_t>& bytes : books) {
    Workbook plain = Load(bytes);
    Workbook profiled = Load(bytes, false, mac_365_ja_jp_profile());
    ASSERT_EQ(profiled.calc_mode(), CalcMode::kManual);
    EXPECT_TRUE(SameProfile(profiled.excel_profile(), mac_365_ja_jp_profile()));
    EXPECT_FALSE(StaleFormulas(plain).empty());
    EXPECT_EQ(StaleFormulas(profiled), StaleFormulas(plain));
    EXPECT_EQ(CachedFormulaValues(profiled), CachedFormulaValues(plain));
  }
}

// The same holds for a package an .xlsb save wrote.
TEST(StaleRoundtrip, ProfileAtLoadKeepsTheManualStaleSetXlsb) {
  Workbook wb = ManualChain();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 0, Value::number(2.0))));
  const std::vector<std::uint8_t> bytes = SaveXlsb(wb);
  Workbook plain = LoadXlsb(bytes);
  Workbook profiled = LoadXlsb(bytes, mac_365_ja_jp_profile());
  EXPECT_TRUE(SameProfile(profiled.excel_profile(), mac_365_ja_jp_profile()));
  EXPECT_FALSE(StaleFormulas(plain).empty());
  EXPECT_EQ(StaleFormulas(profiled), StaleFormulas(plain));
  EXPECT_EQ(CachedFormulaValues(profiled), CachedFormulaValues(plain));
}

// One loaded-stale (ca) formula among clean ones stays the only stale one under the profile.
TEST(StaleRoundtrip, ProfileAtLoadKeepsExactlyTheCaFormulaStale) {
  Workbook wb = Load(Package("calcMode=\"manual\"", kOneCaRow), false, mac_365_ja_jp_profile());
  const CellSet expected = {Node(0, 0, 3)};
  EXPECT_EQ(StaleFormulas(wb), expected);
  EXPECT_EQ(Number(wb, 0, 0, 1), 7.0);
  EXPECT_EQ(Number(wb, 0, 0, 2), 3.0);
}

// Under automatic calculation a workbook loaded with a profile still computes every formula.
TEST(StaleRoundtrip, ProfileAtLoadStillComputesUnderAutomaticCalculation) {
  Workbook wb = Load(Package("", kCachedRow), false, mac_365_ja_jp_profile());
  ASSERT_EQ(wb.calc_mode(), CalcMode::kAuto);
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_EQ(Number(wb, 0, 0, 1), 20.0);
  EXPECT_EQ(Number(wb, 0, 0, 2), 3.0);
  EXPECT_EQ(Number(wb, 0, 0, 3), 4.0);
}

// Sheet names are validated under the load's default host, so a name only Windows Excel refuses on entry loads
// the same under either profile; the assigned profile then governs later entries.
TEST(StaleRoundtrip, SheetNameOnlyOneHostRefusesLoadsUnderEitherProfile) {
  const std::vector<std::uint8_t> bytes = Package("calcMode=\"manual\"", kCachedRow, "History");
  Workbook plain = Load(bytes);
  Workbook mac = Load(bytes, false, mac_365_ja_jp_profile());
  Workbook win = Load(bytes, false, win_365_ja_jp_profile());
  for (const Workbook* wb : {&plain, &mac, &win}) {
    ASSERT_EQ(wb->sheet_count(), 1U);
    EXPECT_EQ(wb->sheet(0).name(), "History");
  }
  EXPECT_EQ(StaleFormulas(mac), StaleFormulas(plain));
  EXPECT_EQ(StaleFormulas(win), StaleFormulas(plain));

  const std::vector<std::uint8_t> other = Package("calcMode=\"manual\"", kCachedRow);
  Workbook win_loaded = Load(other);
  Workbook mac_loaded = Load(other, false, mac_365_ja_jp_profile());
  EXPECT_FALSE(static_cast<bool>(win_loaded.add_sheet_entered("History")));
  EXPECT_TRUE(static_cast<bool>(mac_loaded.add_sheet_entered("History")));
}

// `set_excel_profile` after a load still dirties every formula, unlike a profile given at load.
TEST(StaleRoundtrip, ProfileSetAfterLoadStillDirtiesEveryFormula) {
  Workbook wb = Load(Package("calcMode=\"manual\"", kOneCaRow));
  ASSERT_EQ(StaleFormulas(wb).size(), 1U);
  wb.set_excel_profile(mac_365_ja_jp_profile());
  const CellSet expected = {Node(0, 0, 1), Node(0, 0, 2), Node(0, 0, 3)};
  EXPECT_EQ(StaleFormulas(wb), expected);
}

// How many cells across the workbook show `text`.
static std::size_t CountText(const Workbook& wb, std::string_view text) {
  std::size_t n = 0;
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    for (const auto& [row, cells] : wb.sheet(s).rows()) {
      for (const auto& [col, cell] : cells) {
        if (cell.cached_value.is_text() && cell.cached_value.as_text() == text) {
          ++n;
        }
      }
    }
  }
  return n;
}

// A PivotTable refreshed on load is refreshed at the first settle under the profile the load gave: its captions are
// the profile's.
TEST(StaleRoundtrip, RefreshOnLoadPivotUsesTheLoadProfile) {
  std::vector<std::uint8_t> bytes;
  {
    const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/pivot_options/op_showall_only.xlsx";
    FILE* file = std::fopen(path.c_str(), "rb");
    ASSERT_NE(file, nullptr) << path;
    std::fseek(file, 0, SEEK_END);
    bytes.resize(static_cast<std::size_t>(std::ftell(file)));
    std::fseek(file, 0, SEEK_SET);
    ASSERT_EQ(std::fread(bytes.data(), 1, bytes.size(), file), bytes.size());
    std::fclose(file);
  }
  Workbook en = Load(bytes, false, win_365_en_us_profile());
  ASSERT_TRUE(static_cast<bool>(en.recalc(eval::default_registry())));
  Workbook ja = Load(bytes, false, mac_365_ja_jp_profile());
  ASSERT_TRUE(static_cast<bool>(ja.recalc(eval::default_registry())));
  EXPECT_GT(CountText(en, "Grand Total"), 0U);
  EXPECT_EQ(CountText(en, "\xE7\xB7\x8F\xE8\xA8\x88"), 0U);
  EXPECT_GT(CountText(ja, "\xE7\xB7\x8F\xE8\xA8\x88"), 0U);  // 総計
  EXPECT_EQ(CountText(ja, "Grand Total"), 0U);
}

// A one-cell dynamic array at the top-left of a merge keeps its value on load and across a save.
TEST(StaleRoundtrip, MergedOneCellDynamicArrayKeepsItsValueInManualMode) {
  const auto c3_text = [](const Workbook& wb) {
    const Value v = wb.sheet(0).resolve_cell_value(2U, 2U);
    EXPECT_TRUE(v.is_text());
    return v.is_text() ? std::string(v.as_text()) : std::string();
  };
  Workbook wb = Load(ReadFixture("../load_save/merged_one_cell_dynamic_array.xlsx"));
  EXPECT_EQ(c3_text(wb), "A");
  wb.set_calc_mode(CalcMode::kManual);
  const Workbook manual = Load(Save(wb));
  ASSERT_EQ(manual.calc_mode(), CalcMode::kManual);
  EXPECT_EQ(c3_text(manual), "A");
}

// Sheet L holds a legacy array block A5:A7 =Nm*2 and sheet D the same as a dynamic array, Nm = L!$A$1:$A$3:
// loading the names the block refers to keeps the block's cached members and the used range.
TEST(StaleRoundtrip, ArrayBlockReferencingANameKeepsItsMembersOnLoad) {
  const auto expect_block = [](const Workbook& wb) {
    for (std::size_t sheet = 0; sheet < 2U; ++sheet) {
      EXPECT_NE(wb.sheet(sheet).spill_region_at_anchor(4U, 0U), nullptr) << "sheet " << sheet;
      EXPECT_EQ(Number(wb, sheet, 4U, 0U), 2.0);
      EXPECT_EQ(Number(wb, sheet, 5U, 0U), 4.0);
      EXPECT_EQ(Number(wb, sheet, 6U, 0U), 6.0);
      const auto extent = wb.sheet(sheet).populated_extent(0U, 0U, Sheet::kMaxRows - 1U, Sheet::kMaxCols - 1U);
      ASSERT_TRUE(extent.has_value());
      EXPECT_EQ(extent->last_row, 6U) << "sheet " << sheet;
    }
  };
  Workbook wb = Load(ReadFixture("../load_save/array_block_referencing_name.xlsx"));
  expect_block(wb);
  wb.set_calc_mode(CalcMode::kManual);
  // An automatic-mode book loads with every formula pending; settle it so the manual save marks nothing.
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_TRUE(StaleFormulas(wb).empty());
  for (const bool sax : {false, true}) {
    const Workbook manual = Load(Save(wb), sax);
    ASSERT_EQ(manual.calc_mode(), CalcMode::kManual);
    expect_block(manual);
    EXPECT_TRUE(StaleFormulas(manual).empty());
    expect_block(Load(Save(manual)));
  }
}

// Excel can save a dynamic-array MAP anchor with a one-cell extent where the formula spills two rows: it keeps the
// shrunken extent through a reopen and CalculateFull while a legacy-entered lambda-helper cell sits in the same book.
// The workbook's formulas do not determine that extent, so Formulon shows the cached cell until a full recalculation
// and then spills what the formula defines.
TEST(StaleRoundtrip, ShrunkenDynamicArrayExtentHoldsUntilAFullRecalc) {
  constexpr std::string_view kRows =
      "<row r=\"22\"><c r=\"B22\"><f ca=\"1\">_xlfn.MAP(A1:A2,_xlfn.LAMBDA(_xlpm.x,ROW(_xlpm.x)))</f>"
      "<v>0</v></c><c r=\"F22\" cm=\"1\"><f t=\"array\" aca=\"1\" ref=\"F22\" ca=\"1\">"
      "_xlfn.MAP(A1:A2,_xlfn.LAMBDA(_xlpm.x,ROW(_xlpm.x)))</f><v>1</v></c></row>";
  for (const bool sax : {false, true}) {
    SCOPED_TRACE(sax ? "sax" : "dom");
    Workbook wb = Load(Package("calcId=\"191029\"", kRows, "Sheet1", kDynamicArrayMetadata), sax);
    EXPECT_EQ(Number(wb, 0, 21, 5), 1.0);
    EXPECT_TRUE(wb.sheet(0).resolve_cell_value(22U, 5U).is_blank());
    ASSERT_TRUE(static_cast<bool>(wb.recalc_full(eval::default_registry())));
    EXPECT_EQ(Number(wb, 0, 21, 5), 1.0);
    EXPECT_EQ(Number(wb, 0, 22, 5), 2.0);
  }
}

// An `@` over a grid of references takes its first reference. A legacy cell (no dynamic-array mark) intersects that
// reference with the formula cell, so `SUM(OFFSET(A1:A3,{0;1;2},0)*1)` reads A3 on row 3 and is #VALUE! on row 10; a
// dynamic-array cell reads it whole (4). The stored kind decides, on load and after our own save.
TEST(StaleRoundtrip, LegacyCellIntersectsTheFirstReferenceOfAGrid) {
  constexpr std::string_view kRows =
      "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>"
      "<row r=\"3\"><c r=\"A3\"><v>3</v></c>"
      "<c r=\"B3\"><f ca=\"1\">SUM(OFFSET(A1:A3,{0;1;2},0)*1)</f><v>0</v></c></row>"
      "<row r=\"10\"><c r=\"B10\"><f ca=\"1\">SUM(OFFSET(A1:A3,{0;1;2},0)*1)</f><v>0</v></c>"
      "<c r=\"C10\" cm=\"1\"><f t=\"array\" aca=\"1\" ref=\"C10\" ca=\"1\">"
      "SUM(_xlfn.SINGLE(OFFSET(A1:A3,{0;1;2},0))*1)</f><v>0</v></c></row>";
  const std::vector<std::uint8_t> package = Package("calcId=\"191029\"", kRows, "Sheet1", kDynamicArrayMetadata);
  const auto expect_values = [](Workbook wb) {
    ASSERT_TRUE(static_cast<bool>(wb.recalc_full(eval::default_registry())));
    EXPECT_EQ(Number(wb, 0, 2, 1), 3.0);
    const Value outside = wb.sheet(0).resolve_cell_value(9U, 1U);
    ASSERT_TRUE(outside.is_error());
    EXPECT_EQ(outside.as_error(), ErrorCode::Value);
    EXPECT_EQ(Number(wb, 0, 9, 2), 4.0);
  };
  for (const bool sax : {false, true}) {
    SCOPED_TRACE(sax ? "sax" : "dom");
    expect_values(Load(package, sax));
  }
  {
    SCOPED_TRACE("xlsx round trip");
    expect_values(Load(Save(Load(package))));
  }
  {
    SCOPED_TRACE("xlsb round trip");
    expect_values(LoadXlsb(SaveXlsb(Load(package))));
  }
}

// Each cell's `<f>` text in `sheet_xml`, keyed by cell reference.
std::vector<std::pair<std::string, std::string>> FormulaTexts(const std::string& sheet_xml) {
  pugi::xml_document doc;
  EXPECT_TRUE(static_cast<bool>(doc.load_buffer(sheet_xml.data(), sheet_xml.size())));
  std::vector<std::pair<std::string, std::string>> out;
  for (const pugi::xml_node row : doc.child("worksheet").child("sheetData").children("row")) {
    for (const pugi::xml_node cell : row.children("c")) {
      if (const pugi::xml_node f = cell.child("f")) {
        out.emplace_back(cell.attribute("r").value(), f.text().get());
      }
    }
  }
  return out;
}

// A formula re-spelled for storage keeps its spaces and line breaks, as Excel wrote them.
TEST(StaleRoundtrip, RespelledFormulasKeepTheirWhitespace) {
  const std::vector<std::uint8_t> excel = ReadFixture("../load_save/formula_whitespace_restorage.xlsx");
  const auto written = FormulaTexts(Part(excel, "xl/worksheets/sheet1.xml"));
  ASSERT_EQ(written.size(), 4U);
  const auto saved = FormulaTexts(Part(Save(Load(excel)), "xl/worksheets/sheet1.xml"));
  EXPECT_EQ(saved, written);
}

// Retargeting the name after load is an edit: the dependent blocks lose their spill and go stale until recalculated.
TEST(StaleRoundtrip, RetargetingANameAfterLoadInvalidatesTheArrayBlocks) {
  Workbook wb = Load(ReadFixture("../load_save/array_block_referencing_name.xlsx"));
  wb.set_calc_mode(CalcMode::kManual);
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  Workbook manual = Load(Save(wb));
  ASSERT_EQ(manual.calc_mode(), CalcMode::kManual);
  ASSERT_TRUE(StaleFormulas(manual).empty());
  ASSERT_TRUE(static_cast<bool>(manual.set_defined_name("Nm", "L!$A$2:$A$4")));
  const CellSet stale = StaleFormulas(manual);
  for (std::uint16_t sheet = 0; sheet < 2U; ++sheet) {
    EXPECT_EQ(manual.sheet(sheet).spill_region_at_anchor(4U, 0U), nullptr) << "sheet " << sheet;
    EXPECT_EQ(stale.count(Node(sheet, 4U, 0U)), 1U) << "sheet " << sheet;
  }
  ASSERT_TRUE(static_cast<bool>(manual.recalc(eval::default_registry())));
  for (std::size_t sheet = 0; sheet < 2U; ++sheet) {
    EXPECT_EQ(Number(manual, sheet, 4U, 0U), 4.0);
    EXPECT_EQ(Number(manual, sheet, 5U, 0U), 6.0);
    EXPECT_EQ(Number(manual, sheet, 6U, 0U), 0.0);
  }
}

namespace {

// A1 = 1; B1 = A1+1 cached 99; C1 = "a"&" " cached "a".
constexpr std::string_view kStaleCacheRow =
    "<row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"B1\"><f>A1+1</f><v>99</v></c>"
    "<c r=\"C1\" t=\"str\"><f>\"a\"&amp;\" \"</f><v>a</v></c></row>";

// Whether B1 or C1 of the stale-cache book awaits a recalc.
bool AwaitsRecalc(const Workbook& wb) {
  return wb.recalc_engine().awaits_recalc(Node(0, 0, 1)) || wb.recalc_engine().awaits_recalc(Node(0, 0, 2));
}

std::string Text(const Workbook& wb, std::uint32_t row, std::uint32_t col) {
  const Value v = wb.sheet(0).resolve_cell_value(row, col);
  EXPECT_TRUE(v.is_text()) << "row " << row << " col " << col;
  return v.is_text() ? std::string(v.as_text()) : std::string();
}

}  // namespace

// A book that opens with its cache keeps a stale number and a stale text until a full recalculation.
TEST(StaleRoundtrip, FullRecalcRecomputesTheCachedValuesABookOpensWith) {
  for (const char* calc_pr : {"calcId=\"191029\"", "calcId=\"191029\" calcMode=\"manual\""}) {
    SCOPED_TRACE(calc_pr);
    Workbook wb = Load(Package(calc_pr, kStaleCacheRow));
    EXPECT_FALSE(wb.opens_with_recalc());
    EXPECT_FALSE(AwaitsRecalc(wb));
    ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    EXPECT_EQ(Number(wb, 0, 0, 1), 99.0);
    EXPECT_EQ(Text(wb, 0, 2), "a");

    auto stats = wb.recalc_full(eval::default_registry());
    ASSERT_TRUE(static_cast<bool>(stats));
    EXPECT_EQ(stats.value().cells_evaluated, 2U);
    EXPECT_EQ(Number(wb, 0, 0, 1), 2.0);
    EXPECT_EQ(Text(wb, 0, 2), "a ");
    EXPECT_FALSE(AwaitsRecalc(wb));
  }
}

// fullCalcOnLoad makes the same book open dirty, so a plain recalc computes it.
TEST(StaleRoundtrip, FullCalcOnLoadBookIsComputedByAPlainRecalc) {
  Workbook wb = Load(Package("calcId=\"191029\" fullCalcOnLoad=\"1\"", kStaleCacheRow));
  EXPECT_TRUE(wb.opens_with_recalc());
  EXPECT_TRUE(AwaitsRecalc(wb));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_EQ(Number(wb, 0, 0, 1), 2.0);
  EXPECT_EQ(Text(wb, 0, 2), "a ");
}

// The parallel full recalculation computes the same values and counts the same cells.
TEST(StaleRoundtrip, ParallelFullRecalcMatchesTheSerialOne) {
  for (const std::uint32_t threads : {1U, 4U}) {
    SCOPED_TRACE(threads);
    Workbook wb = Load(Package("calcId=\"191029\"", kStaleCacheRow));
    eval::SchedulerConfig config;
    config.num_threads = threads;
    eval::SchedulerStats stats;
    ASSERT_TRUE(static_cast<bool>(wb.recalc_full_parallel(eval::default_registry(), config, &stats)));
    EXPECT_EQ(stats.cells_evaluated, 2U);
    EXPECT_EQ(Number(wb, 0, 0, 1), 2.0);
    EXPECT_EQ(Text(wb, 0, 2), "a ");
    EXPECT_FALSE(AwaitsRecalc(wb));
  }
}

// Called from inside a recalc on the same thread, both forms are refused without waiting on the engine.
TEST(StaleRoundtrip, FullRecalcFromInsideARecalcIsRefused) {
  // A1 = B1+1, B1 = A1/2 runs the iterative solver, whose progress callback re-enters.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=B1+1")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=A1/2")));
  IterativeOptions options;
  options.enabled = true;
  options.max_iterations = 2U;
  wb.set_iterative_options(options);
  struct Probe {
    Workbook* wb = nullptr;
    bool called = false;
    FormulonErrorCode serial = FormulonErrorCode::kOk;
    FormulonErrorCode parallel = FormulonErrorCode::kOk;
  } probe;
  probe.wb = &wb;
  wb.set_iterative_progress(
      [](std::uint32_t, double, std::uint32_t, void* data) {
        auto& p = *static_cast<Probe*>(data);
        if (p.called) {
          return true;
        }
        p.called = true;
        auto serial = p.wb->recalc_full(eval::default_registry());
        p.serial = serial ? FormulonErrorCode::kOk : serial.error().code;
        auto parallel = p.wb->recalc_full_parallel(eval::default_registry(), eval::SchedulerConfig{}, nullptr);
        p.parallel = parallel ? FormulonErrorCode::kOk : parallel.error().code;
        return true;
      },
      &probe);
  const auto outer = wb.recalc(eval::default_registry());
  wb.set_iterative_progress(nullptr, nullptr);
  ASSERT_TRUE(static_cast<bool>(outer));
  ASSERT_TRUE(probe.called);
  EXPECT_EQ(probe.serial, FormulonErrorCode::kGraphRecalcReentrant);
  EXPECT_EQ(probe.parallel, FormulonErrorCode::kGraphRecalcReentrant);
}

// Calculated-column formulas filled into the table of `<stem>.xlsx`: whether the engine's save marks the first data
// row `ca="1"`, against whether Excel's save of the same book (`<stem>_saved.xlsx`) does.
struct CaMarks {
  bool engine = false;
  bool excel = false;
};

CaMarks CalculatedColumnCaMarks(const std::string& stem, const char* formula) {
  CaMarks marks;
  const std::string dir = std::string("formula_ca_self_read/") + stem;
  Workbook wb = Load(ReadFixture("../" + dir + ".xlsx"));
  for (std::uint32_t row = 1U; row <= 3U; ++row) {
    EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(0, row, 1, formula)));
  }
  static_cast<void>(wb.recalc(eval::default_registry()));
  auto saved = io::write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
  if (!saved) {
    return marks;
  }
  marks.engine = MarkedCa(Part(saved.value(), "xl/worksheets/sheet1.xml"), "B2");
  marks.excel = MarkedCa(Part(ReadFixture("../" + dir + "_saved.xlsx"), "xl/worksheets/sheet1.xml"), "B2");
  return marks;
}

// `=ROWS(Table1[#Data])` in a calculated column reads only the shape of its own table, and neither save marks it.
TEST(CalculatedColumnCa, AFormulaTakingOnlyTheShapeOfItsOwnColumnIsNotMarked) {
  const CaMarks marks = CalculatedColumnCaMarks("shape_only_rows", "=ROWS(Table1[#Data])");
  EXPECT_FALSE(marks.excel);
  EXPECT_EQ(marks.engine, marks.excel);
}

void ExpectSelfReadMarkAsExcel() {
  const CaMarks marks = CalculatedColumnCaMarks("self_read_sum", "=SUM(Table1[X])");
  EXPECT_EQ(marks.engine, marks.excel);
}

// `=SUM(Table1[X])` in a calculated column reads its own cell. Excel's save marks it `ca="1"`; the engine marks only
// volatile, unknown-function and calculation-name formulas, so it does not. Fails once a self-reading formula is
// marked.
TEST(CalculatedColumnCa, ASelfReadingFormulaIsMarkedByExcelButNotByTheEngine) {
  const CaMarks marks = CalculatedColumnCaMarks("self_read_sum", "=SUM(Table1[X])");
  EXPECT_TRUE(marks.excel);
  EXPECT_FALSE(marks.engine);
  EXPECT_NONFATAL_FAILURE(ExpectSelfReadMarkAsExcel(), "");
}

}  // namespace formulon
