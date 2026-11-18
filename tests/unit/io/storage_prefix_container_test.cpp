//
// The storage prefix is two independent questions, and Excel answers
// them differently for the same function.
//
//   1. How is the callee spelled in the OOXML `<f>` text? A per-function
//      property of when the function was introduced
//      (`io::classify_storage_prefix`).
//   2. Does Excel's classic function table have an id for the callee, or
//      must an XLSB call reach it through the hidden-name route
//      (`PtgName` + `PtgFuncVar` with the `id == 255` sentinel)? A
//      property of the binary format's function table
//      (`io::xlsb_uses_hidden_name`).
//
// `ISO.CEILING` is the case where the two disagree, and the fixture pair
// is one workbook Excel 365 saved twice, once per container:
//
//   `storage_prefix_probe.xlsx`  A1 `<f>ISO.CEILING(4.3)</f>` -- bare,
//                                and no `<definedNames>` element at all.
//   `storage_prefix_probe.xlsb`  a `BrtName` spelling
//                                `_xlfn.ISO.CEILING`, with A1's cell
//                                stream `PtgName`, `PtgNum`,
//                                `PtgFuncVar(cparams=2, id=255)`.
//
// A2 `MROUND` (bare with a real function id in both) and A3 `XLOOKUP`
// (`_xlfn.`-prefixed in both) are the controls: they are what
// distinguishes a genuine asymmetry from a broken observation.
//
// Nothing else in the suite pins that asymmetry, so a refactor that
// "unifies" the two writers behind one classifier would silently make
// Formulon emit `_xlfn.ISO.CEILING` into `<f>`, which real Excel renders
// as `#NAME?`.

#include <cstdint>
#include <deque>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "cell.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/cf_reader.h"
#include "io/future_functions.h"
#include "io/ooxml_reader.h"
#include "io/sheet_reader.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "parser/formula_prefix.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "support/ooxml_package_fixture.h"
#include "support/roundtrip_symmetry.h"
#include "value.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace {

std::string XlsxFixturePath() {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/storage_prefix_probe.xlsx";
}

std::string XlsbFixturePath() {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/storage_prefix_probe.xlsb";
}

/// True when `haystack` contains `needle` encoded the way `BrtName`
/// stores a name: UTF-16LE, no BOM. Excel names are ASCII, so each byte
/// is followed by a zero byte.
bool ContainsUtf16Le(const std::string& haystack, std::string_view needle) {
  std::string wide;
  wide.reserve(needle.size() * 2U);
  for (const char c : needle) {
    wide.push_back(c);
    wide.push_back('\0');
  }
  return haystack.find(wide) != std::string::npos;
}

TEST(StoragePrefixContainers, ExcelSpellsIsoCeilingBareInOoxml) {
  const std::vector<std::uint8_t> bytes = test::read_file_bytes(XlsxFixturePath());
  ASSERT_FALSE(bytes.empty());

  std::string sheet;
  const ::testing::AssertionResult got_sheet =
      test::extract_part(test::span_of(bytes), "xl/worksheets/sheet1.xml", &sheet);
  ASSERT_TRUE(static_cast<bool>(got_sheet)) << got_sheet.message();

  EXPECT_NE(sheet.find("<f>ISO.CEILING(4.3)</f>"), std::string::npos) << sheet;
  EXPECT_EQ(sheet.find("_xlfn.ISO.CEILING"), std::string::npos) << sheet;
  // Controls: bare-with-an-id, and prefixed-in-both.
  EXPECT_NE(sheet.find("<f>MROUND(17,5)</f>"), std::string::npos) << sheet;
  EXPECT_EQ(sheet.find("_xlfn.MROUND"), std::string::npos) << sheet;
  EXPECT_NE(sheet.find("<f>_xlfn.XLOOKUP("), std::string::npos) << sheet;

  // In OOXML the prefix in `<f>` is the whole mechanism: Excel registers
  // no hidden defined name to back it, not even for XLOOKUP.
  std::string workbook;
  const ::testing::AssertionResult got_workbook =
      test::extract_part(test::span_of(bytes), "xl/workbook.xml", &workbook);
  ASSERT_TRUE(static_cast<bool>(got_workbook)) << got_workbook.message();
  EXPECT_EQ(workbook.find("<definedNames>"), std::string::npos) << workbook;
  EXPECT_EQ(workbook.find("_xlfn"), std::string::npos) << workbook;
}

TEST(StoragePrefixContainers, ExcelSpellsIsoCeilingPrefixedInXlsbHiddenName) {
  const std::vector<std::uint8_t> bytes = test::read_file_bytes(XlsbFixturePath());
  ASSERT_FALSE(bytes.empty());

  std::string workbook;
  const ::testing::AssertionResult got_workbook =
      test::extract_part(test::span_of(bytes), "xl/workbook.bin", &workbook);
  ASSERT_TRUE(static_cast<bool>(got_workbook)) << got_workbook.message();

  // The same workbook that spells `ISO.CEILING` bare in `<f>` carries a
  // `_xlfn.`-prefixed hidden `BrtName` for it here, because the binary
  // format has no function id to encode the call with.
  EXPECT_TRUE(ContainsUtf16Le(workbook, "_xlfn.ISO.CEILING"));
  EXPECT_TRUE(ContainsUtf16Le(workbook, "_xlfn.XLOOKUP"));
  // MROUND has a real function id (422), so it needs no hidden name.
  EXPECT_FALSE(ContainsUtf16Le(workbook, "MROUND"));
}

TEST(StoragePrefixContainers, BothContainersReadBackToTheSameBareFormulaText) {
  // Whichever spelling the container used, the formula bar text is the
  // same, so both readers must land on the bare name.
  const std::vector<std::uint8_t> xlsx = test::read_file_bytes(XlsxFixturePath());
  ASSERT_FALSE(xlsx.empty());
  auto from_xlsx = read_ooxml(test::span_of(xlsx));
  ASSERT_TRUE(static_cast<bool>(from_xlsx)) << from_xlsx.error().message;
  const Cell* xlsx_a1 = from_xlsx.value().workbook.sheet(0).cell_at(0U, 0U);
  ASSERT_NE(xlsx_a1, nullptr);
  EXPECT_EQ(xlsx_a1->formula_text, "=ISO.CEILING(4.3)");

  const std::vector<std::uint8_t> xlsb = test::read_file_bytes(XlsbFixturePath());
  ASSERT_FALSE(xlsb.empty());
  auto from_xlsb = xlsb::read_xlsb(test::span_of(xlsb));
  ASSERT_TRUE(static_cast<bool>(from_xlsb)) << from_xlsb.error().message;
  const Cell* xlsb_a1 = from_xlsb.value().workbook.sheet(0).cell_at(0U, 0U);
  ASSERT_NE(xlsb_a1, nullptr);
  EXPECT_EQ(xlsb_a1->formula_text, "=ISO.CEILING(4.3)");
}

TEST(StoragePrefixContainers, ClassifierAnswersTheTwoAxesIndependently) {
  // ISO.CEILING: bare in `<f>`, hidden-name route in XLSB.
  EXPECT_EQ(classify_storage_prefix("ISO.CEILING"), parser::StoragePrefixKind::None);
  EXPECT_EQ(storage_function_name("ISO.CEILING"), "ISO.CEILING");
  EXPECT_TRUE(xlsb_uses_hidden_name("ISO.CEILING"));
  EXPECT_EQ(xlsb_hidden_function_name("ISO.CEILING"), "_xlfn.ISO.CEILING");
  EXPECT_TRUE(xlsb_uses_hidden_name("iso.ceiling"));

  // XLOOKUP: prefixed on both axes, and the hidden name matches the
  // OOXML spelling.
  EXPECT_EQ(classify_storage_prefix("XLOOKUP"), parser::StoragePrefixKind::Xlfn);
  EXPECT_EQ(xlsb_hidden_function_name("XLOOKUP"), storage_function_name("XLOOKUP"));

  // MROUND: bare on both axes -- it has a real function id, so nothing
  // may push it onto the hidden-name route.
  EXPECT_EQ(classify_storage_prefix("MROUND"), parser::StoragePrefixKind::None);
  EXPECT_FALSE(xlsb_uses_hidden_name("MROUND"));
  // A callee whose id simply has not been harvested must not qualify
  // either: writing a hidden name Excel does not know yields `#NAME?`.
  EXPECT_FALSE(xlsb_uses_hidden_name("CUBEVALUE"));
}

TEST(StoragePrefixContainers, IsoCeilingSurvivesXlsbSaveWithoutDowngrade) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.sheet(wb.add_sheet("F"));
  s.set_cell_formula(0U, 0U, "=ISO.CEILING(4.3)");
  s.set_cell_formula(1U, 0U, "=ISO.CEILING(4.3,1)");

  auto write_or = xlsb::write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(write_or)) << write_or.error().message << " | " << write_or.error().context;
  EXPECT_EQ(write_or.value().diagnostics.downgraded_formula_count, 0U);

  std::string workbook_part;
  const ::testing::AssertionResult got_workbook =
      test::extract_part(test::span_of(write_or.value().bytes), "xl/workbook.bin", &workbook_part);
  ASSERT_TRUE(static_cast<bool>(got_workbook)) << got_workbook.message();
  EXPECT_TRUE(ContainsUtf16Le(workbook_part, "_xlfn.ISO.CEILING"));

  auto read_or = xlsb::read_xlsb(test::span_of(write_or.value().bytes));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  const Sheet& rt = read_or.value().workbook.sheet(0);
  const Cell* a1 = rt.cell_at(0U, 0U);
  ASSERT_NE(a1, nullptr);
  EXPECT_EQ(a1->formula_text, "=ISO.CEILING(4.3)");
  const Cell* a2 = rt.cell_at(1U, 0U);
  ASSERT_NE(a2, nullptr);
  EXPECT_EQ(a2->formula_text, "=ISO.CEILING(4.3,1)");
}

// One workbook Excel saved as both containers: `ECMA.CEILING` is bare in
// `<f>` and goes through a hidden `_xlfn.` name in `.xlsb`, like ISO.CEILING.
TEST(StoragePrefixContainers, EcmaCeilingReadsFromBothExcelContainersAndWritesBack) {
  const std::string dir = std::string(FORMULON_FIXTURES_DIR) + "/excel/";
  const std::vector<std::uint8_t> xlsx = test::read_file_bytes(dir + "ecma_ceiling_saved.xlsx");
  ASSERT_FALSE(xlsx.empty());
  const std::vector<std::uint8_t> xlsb = test::read_file_bytes(dir + "ecma_ceiling_saved.xlsb");
  ASSERT_FALSE(xlsb.empty());
  std::string excel_workbook;
  ASSERT_TRUE(static_cast<bool>(test::extract_part(test::span_of(xlsb), "xl/workbook.bin", &excel_workbook)));
  EXPECT_TRUE(ContainsUtf16Le(excel_workbook, "_xlfn.ECMA.CEILING"));

  auto from_xlsx = read_ooxml(test::span_of(xlsx));
  ASSERT_TRUE(static_cast<bool>(from_xlsx)) << from_xlsx.error().message;
  auto from_xlsb = xlsb::read_xlsb(test::span_of(xlsb));
  ASSERT_TRUE(static_cast<bool>(from_xlsb)) << from_xlsb.error().message;
  for (const Workbook* wb : {&from_xlsx.value().workbook, &from_xlsb.value().workbook}) {
    const Cell* a1 = wb->sheet(0).cell_at(0U, 0U);
    ASSERT_NE(a1, nullptr);
    EXPECT_EQ(a1->formula_text, "=ECMA.CEILING(4.3,-2)");
  }

  auto write_or = xlsb::write_xlsb_with_result(from_xlsx.value().workbook);
  ASSERT_TRUE(static_cast<bool>(write_or)) << write_or.error().message << " | " << write_or.error().context;
  EXPECT_EQ(write_or.value().diagnostics.downgraded_formula_count, 0U);
  std::string written_workbook;
  ASSERT_TRUE(static_cast<bool>(
      test::extract_part(test::span_of(write_or.value().bytes), "xl/workbook.bin", &written_workbook)));
  EXPECT_TRUE(ContainsUtf16Le(written_workbook, "_xlfn.ECMA.CEILING"));
  auto read_or = xlsb::read_xlsb(test::span_of(write_or.value().bytes));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message;
  const Cell* a1 = read_or.value().workbook.sheet(0).cell_at(0U, 0U);
  ASSERT_NE(a1, nullptr);
  EXPECT_EQ(a1->formula_text, "=ECMA.CEILING(4.3,-2)");
}

TEST(StoragePrefixContainers, LocalisedNameResolvesToOneStoredSpelling) {
  // `JIS` is the ja-JP formula-bar spelling of `DBCS`; Excel resolves it
  // at entry and never stores it. Both storage paths therefore consult
  // one enumeration, so neither container can store a name Excel's
  // invariant grammar lacks.
  EXPECT_EQ(canonical_function_name("JIS"), "DBCS");
  EXPECT_EQ(canonical_function_name("jis"), "DBCS");
  EXPECT_EQ(canonical_function_name("DBCS"), "DBCS");
  // Everything else is stored exactly as spelled, case included.
  EXPECT_EQ(canonical_function_name("SUM"), "SUM");
  EXPECT_EQ(canonical_function_name("sum"), "sum");

  EXPECT_EQ(storage_call_name("JIS"), "DBCS");
  EXPECT_EQ(storage_call_name("SUM"), "SUM");
  EXPECT_EQ(storage_call_name("sum"), "sum");
  EXPECT_EQ(storage_call_name("XLOOKUP"), "_xlfn.XLOOKUP");
  EXPECT_EQ(storage_call_name("FILTER"), "_xlfn._xlws.FILTER");
  EXPECT_EQ(storage_call_name("ISO.CEILING"), "ISO.CEILING");
}

TEST(StoragePrefixContainers, ModelHoldingJisSavesAsDbcsInBothContainers) {
  // The mirror of the read-side check: one model, saved twice, must not
  // produce two different callees. `JIS` in `<f>` would be `#NAME?` on
  // open, since Excel's invariant grammar has no such function.
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.sheet(wb.add_sheet("F"));
  s.set_cell_formula(0U, 0U, "=JIS(\"ABC\")");

  auto xlsx_or = wb.save();
  ASSERT_TRUE(static_cast<bool>(xlsx_or)) << xlsx_or.error().message;
  std::string sheet;
  const ::testing::AssertionResult got_sheet =
      test::extract_part(test::span_of(xlsx_or.value()), "xl/worksheets/sheet1.xml", &sheet);
  ASSERT_TRUE(static_cast<bool>(got_sheet)) << got_sheet.message();
  // The `<f>` text is XML-escaped, so the string literal's quotes appear
  // as `&quot;`.
  EXPECT_NE(sheet.find("<f>DBCS(&quot;ABC&quot;)</f>"), std::string::npos) << sheet;
  EXPECT_EQ(sheet.find("JIS("), std::string::npos) << sheet;

  auto from_xlsx = read_ooxml(test::span_of(xlsx_or.value()));
  ASSERT_TRUE(static_cast<bool>(from_xlsx)) << from_xlsx.error().message;
  const Cell* xlsx_a1 = from_xlsx.value().workbook.sheet(0).cell_at(0U, 0U);
  ASSERT_NE(xlsx_a1, nullptr);
  EXPECT_EQ(xlsx_a1->formula_text, "=DBCS(\"ABC\")");

  auto xlsb_or = xlsb::write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(xlsb_or)) << xlsb_or.error().message << " | " << xlsb_or.error().context;
  EXPECT_EQ(xlsb_or.value().diagnostics.downgraded_formula_count, 0U);
  auto from_xlsb = xlsb::read_xlsb(test::span_of(xlsb_or.value().bytes));
  ASSERT_TRUE(static_cast<bool>(from_xlsb)) << from_xlsb.error().message;
  const Cell* xlsb_a1 = from_xlsb.value().workbook.sheet(0).cell_at(0U, 0U);
  ASSERT_NE(xlsb_a1, nullptr);
  EXPECT_EQ(xlsb_a1->formula_text, "=DBCS(\"ABC\")");
}

TEST(StoragePrefixContainers, AtAndSpillOperatorsRoundTripInBothContainers) {
  // Stored as `_xlfn.SINGLE(...)` / `_xlfn.ANCHORARRAY(...)` calls, read back
  // as the operators the formula bar shows.
  const std::vector<std::string> formulas = {"=@A1", "=SUM(@A1:A2)", "=@A1:A2", "=@SUM(A1:A5)", "=SUM(A1#)", "=A1#+1"};
  // Entered as a user would, so each cell carries the dynamic-array mark
  // Excel gives it; a legacy `A1#+1` would read back as `@A1#+1`.
  Workbook wb = Workbook::create_empty();
  const std::size_t sheet = wb.add_sheet("F");
  for (std::uint32_t i = 0; i < formulas.size(); ++i) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(sheet, i, 1U, formulas[i])));
  }

  auto xlsx_or = wb.save();
  ASSERT_TRUE(static_cast<bool>(xlsx_or)) << xlsx_or.error().message;
  auto from_xlsx = read_ooxml(test::span_of(xlsx_or.value()));
  ASSERT_TRUE(static_cast<bool>(from_xlsx)) << from_xlsx.error().message;

  auto xlsb_or = xlsb::write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(xlsb_or)) << xlsb_or.error().message << " | " << xlsb_or.error().context;
  EXPECT_EQ(xlsb_or.value().diagnostics.downgraded_formula_count, 0U);
  auto from_xlsb = xlsb::read_xlsb(test::span_of(xlsb_or.value().bytes));
  ASSERT_TRUE(static_cast<bool>(from_xlsb)) << from_xlsb.error().message;
  EXPECT_EQ(from_xlsb.value().undecoded_formula_count, 0U);

  for (std::uint32_t i = 0; i < formulas.size(); ++i) {
    const Cell* xlsx_cell = from_xlsx.value().workbook.sheet(0).cell_at(i, 1U);
    const Cell* xlsb_cell = from_xlsb.value().workbook.sheet(0).cell_at(i, 1U);
    ASSERT_NE(xlsx_cell, nullptr) << formulas[i];
    ASSERT_NE(xlsb_cell, nullptr) << formulas[i];
    EXPECT_EQ(xlsx_cell->formula_text, formulas[i]);
    EXPECT_EQ(xlsb_cell->formula_text, formulas[i]);
  }
}

TEST(StoragePrefixContainers, AtAndSpillOperatorsInNamesRoundTripInBothContainers) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("F");
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("AtName", "@F!$A$1:$A$2")));
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("SpillName", "F!$A$1#")));

  auto xlsx_or = wb.save();
  ASSERT_TRUE(static_cast<bool>(xlsx_or)) << xlsx_or.error().message;
  auto from_xlsx = read_ooxml(test::span_of(xlsx_or.value()));
  ASSERT_TRUE(static_cast<bool>(from_xlsx)) << from_xlsx.error().message;
  auto xlsb_or = xlsb::write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(xlsb_or)) << xlsb_or.error().message << " | " << xlsb_or.error().context;
  auto from_xlsb = xlsb::read_xlsb(test::span_of(xlsb_or.value().bytes));
  ASSERT_TRUE(static_cast<bool>(from_xlsb)) << from_xlsb.error().message;
  EXPECT_EQ(from_xlsb.value().undecoded_defined_name_count, 0U);

  for (const Workbook* loaded : {&from_xlsx.value().workbook, &from_xlsb.value().workbook}) {
    std::vector<std::string> formulas;
    for (const DefinedName& dn : loaded->defined_names()) {
      formulas.push_back(dn.name + "=" + dn.formula);
    }
    EXPECT_EQ(formulas, (std::vector<std::string>{"AtName=@F!$A$1:$A$2", "SpillName=F!$A$1#"}));
  }
}

TEST(StoragePrefixContainers, IsoCeilingStaysBareOnOoxmlSave) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.sheet(wb.add_sheet("F"));
  s.set_cell_formula(0U, 0U, "=ISO.CEILING(4.3)");

  auto saved_or = wb.save();
  ASSERT_TRUE(static_cast<bool>(saved_or)) << saved_or.error().message;
  std::string sheet;
  const ::testing::AssertionResult got_sheet =
      test::extract_part(test::span_of(saved_or.value()), "xl/worksheets/sheet1.xml", &sheet);
  ASSERT_TRUE(static_cast<bool>(got_sheet)) << got_sheet.message();
  EXPECT_NE(sheet.find(">ISO.CEILING(4.3)<"), std::string::npos) << sheet;
  EXPECT_EQ(sheet.find("_xlfn.ISO.CEILING"), std::string::npos) << sheet;
}

// A future function spelled bare in `<f>` is a call to a user-defined
// function of that name, not the built-in: Excel 365 shows `#NAME?`, the
// formula bar shows the bare name under the implicit-intersection `@` a
// user-defined call takes, and a save writes `_xludf.NAME`. The `_xlfn.`
// spelling is the built-in.
constexpr std::string_view kBareFutureCallsSheet =
    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>"
    "<row r=\"1\">"
    "<c r=\"A1\"><f>IMAGE(\"https://example.invalid/x.png\")</f></c>"
    "<c r=\"B1\"><f>XLOOKUP(1,{1},{2})</f></c>"
    "<c r=\"C1\"><f>concat(\"a\",\"b\")</f></c>"
    "<c r=\"D1\"><f>_xlfn.CONCAT(\"a\",\"b\")</f></c>"
    "<c r=\"E1\"><f>_xludf.IMAGE(\"https://example.invalid/x.png\")</f></c>"
    "<c r=\"F1\"><f t=\"shared\" ref=\"F1:F2\" si=\"0\">CONCAT(\"a\",G1)</f></c>"
    "<c r=\"G1\"><f>SUM(1,2)&amp;\"IMAGE(\"</f></c>"
    "<c r=\"H1\"><f>_xludf.SUM(1,2)</f></c>"
    "<c r=\"I1\"><f>LET(x,1,x+1)</f></c>"
    "<c r=\"J1\"><f>NORM.DIST(1,0,1,TRUE)</f></c>"
    "</row>"
    "<row r=\"2\"><c r=\"F2\"><f t=\"shared\" si=\"0\"/></c></row>"
    "</sheetData></worksheet>";

struct BareCallRow {
  std::uint32_t row;
  std::uint32_t col;
  const char* sheet_read;  // the sheet reader's text, before the workbook places the legacy `@`
  const char* stored;
  const char* shown;
};

constexpr BareCallRow kBareCallRows[] = {
    {0U, 0U, "=_xludf.IMAGE(\"https://example.invalid/x.png\")", "=@_xludf.IMAGE(\"https://example.invalid/x.png\")",
     "=@IMAGE(\"https://example.invalid/x.png\")"},
    {0U, 1U, "=_xludf.XLOOKUP(1,{1},{2})", "=@_xludf.XLOOKUP(1,{1},{2})", "=@XLOOKUP(1,{1},{2})"},
    {0U, 2U, "=_xludf.CONCAT(\"a\",\"b\")", "=@_xludf.CONCAT(\"a\",\"b\")", "=@CONCAT(\"a\",\"b\")"},
    {0U, 3U, "=CONCAT(\"a\",\"b\")", "=CONCAT(\"a\",\"b\")", "=CONCAT(\"a\",\"b\")"},
    {0U, 4U, "=_xludf.IMAGE(\"https://example.invalid/x.png\")", "=@_xludf.IMAGE(\"https://example.invalid/x.png\")",
     "=@IMAGE(\"https://example.invalid/x.png\")"},
    {0U, 5U, "=_xludf.CONCAT(\"a\",G1)", "=@_xludf.CONCAT(\"a\",G1)", "=@CONCAT(\"a\",G1)"},
    {1U, 5U, "=_xludf.CONCAT(\"a\",G2)", "=@_xludf.CONCAT(\"a\",G2)", "=@CONCAT(\"a\",G2)"},
    {0U, 6U, "=SUM(1,2)&\"IMAGE(\"", "=SUM(1,2)&\"IMAGE(\"", "=SUM(1,2)&\"IMAGE(\""},
    {0U, 7U, "=_xludf.SUM(1,2)", "=@_xludf.SUM(1,2)", "=@SUM(1,2)"},
    {0U, 8U, "=_xludf.LET(x,1,x+1)", "=@_xludf.LET(x,1,@x+1)", "=@LET(x,1,@x+1)"},
    {0U, 9U, "=_xludf.NORM.DIST(1,0,1,TRUE)", "=@_xludf.NORM.DIST(1,0,1,TRUE)", "=@NORM.DIST(1,0,1,TRUE)"},
};

void ExpectBareCallText(const Workbook& wb, bool whole_package) {
  for (const BareCallRow& r : kBareCallRows) {
    SCOPED_TRACE(r.stored);
    const Cell* cell = wb.sheet(0).cell_at(r.row, r.col);
    ASSERT_NE(cell, nullptr);
    if (!whole_package) {
      EXPECT_EQ(cell->formula_text, r.sheet_read);
      continue;
    }
    EXPECT_EQ(cell->formula_text, r.stored);
    EXPECT_EQ(parser::display_cell_shaped_names(cell->formula_text), r.shown);
  }
}

TEST(StoragePrefixContainers, BareFutureCallReadsAsUserDefinedOnBothReadPaths) {
  const std::vector<std::uint8_t> bytes =
      test::BuildOoxmlPackage(test::OoxmlContentTypes(test::kXlsxWorkbookContentType), kBareFutureCallsSheet);
  auto loaded = read_ooxml(test::span_of(bytes));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  ExpectBareCallText(loaded.value().workbook, /*whole_package=*/true);

  Workbook sax = Workbook::create();
  SheetReadContext ctx;
  auto rs = read_sheet_data_sax(
      ByteSpan{reinterpret_cast<const std::uint8_t*>(kBareFutureCallsSheet.data()), kBareFutureCallsSheet.size()}, 0U,
      sax, ctx, sax.mutable_text_storage());
  ASSERT_TRUE(static_cast<bool>(rs)) << rs.error().message;
  ExpectBareCallText(sax, /*whole_package=*/false);
}

TEST(StoragePrefixContainers, BareFutureCallIsNameErrorAndSavesAsUserDefined) {
  const std::vector<std::uint8_t> bytes =
      test::BuildOoxmlPackage(test::OoxmlContentTypes(test::kXlsxWorkbookContentType), kBareFutureCallsSheet);
  auto loaded = read_ooxml(test::span_of(bytes));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  Workbook& wb = loaded.value().workbook;
  auto recalc = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(recalc)) << recalc.error().message;
  for (const std::uint32_t col : {0U, 1U, 2U, 4U, 5U, 7U, 8U, 9U}) {
    SCOPED_TRACE(col);
    const Value v = wb.sheet(0).resolve_cell_value(0U, col);
    ASSERT_TRUE(v.is_error());
    EXPECT_EQ(v.as_error(), ErrorCode::Name);
  }
  const Value d1 = wb.sheet(0).resolve_cell_value(0U, 3U);
  ASSERT_TRUE(d1.is_text());
  EXPECT_EQ(d1.as_text(), "ab");

  auto saved = wb.save();
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  std::string sheet;
  const ::testing::AssertionResult got_sheet =
      test::extract_part(test::span_of(saved.value()), "xl/worksheets/sheet1.xml", &sheet);
  ASSERT_TRUE(static_cast<bool>(got_sheet)) << got_sheet.message();
  EXPECT_NE(sheet.find(">_xludf.IMAGE(&quot;https://example.invalid/x.png&quot;)<"), std::string::npos) << sheet;
  EXPECT_NE(sheet.find(">_xludf.XLOOKUP(1,{1},{2})<"), std::string::npos) << sheet;
  EXPECT_NE(sheet.find(">_xludf.CONCAT(&quot;a&quot;,&quot;b&quot;)<"), std::string::npos) << sheet;
  EXPECT_NE(sheet.find(">_xlfn.CONCAT(&quot;a&quot;,&quot;b&quot;)<"), std::string::npos) << sheet;
}

TEST(StoragePrefixContainers, BareCallSpellingSkipsQuotedTextAndOperatorFunctions) {
  const auto spell = [](std::string_view f) { return parser::spell_bare_calls_as_udf(f, &bare_call_reads_as_udf); };
  EXPECT_EQ(spell("IMAGE(\"IMAGE(\")"), "_xludf.IMAGE(\"IMAGE(\")");
  EXPECT_EQ(spell("'IMAGE(x'!A1+Image(1)"), "'IMAGE(x'!A1+_xludf.IMAGE(1)");
  EXPECT_EQ(spell("T[IMAGE(]+1"), "T[IMAGE(]+1");
  EXPECT_EQ(spell("_xlfn.IMAGE(1)+MY.IMAGE(1)+IMAGE2(1)"), "_xlfn.IMAGE(1)+MY.IMAGE(1)+IMAGE2(1)");
  EXPECT_EQ(spell("SUM(1)+NORM.DIST(1,0,1,TRUE)"), "SUM(1)+_xludf.NORM.DIST(1,0,1,TRUE)");
  // LET, LAMBDA and SINGLE read as user-defined too, their parameters left
  // unprefixed; ANCHORARRAY stays bare.
  EXPECT_EQ(spell("LET(x,1,x+1)+LAMBDA(x,x)(1)+SINGLE(A1)+ANCHORARRAY(A1)"),
            "_xludf.LET(x,1,x+1)+_xludf.LAMBDA(x,x)(1)+_xludf.SINGLE(A1)+ANCHORARRAY(A1)");
  EXPECT_EQ(spell("FILTER({1},{TRUE})"), "FILTER({1},{TRUE})");
}

TEST(StoragePrefixContainers, BareFutureCallInNameAndFeatureFormulasReadsAsUserDefined) {
  EXPECT_EQ(canonical_feature_formula("IMAGE(\"https://example.invalid/x.png\")"),
            "_xludf.IMAGE(\"https://example.invalid/x.png\")");
  EXPECT_EQ(canonical_feature_formula("_xlfn.SINGLE(A1)>0"), "@A1>0");

  constexpr std::string_view kWorkbook =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
      "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
      "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>"
      "<definedNames><definedName name=\"nm\">IMAGE(\"https://example.invalid/x.png\")</definedName></definedNames>"
      "</workbook>";
  constexpr std::string_view kSheet =
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>"
      "<row r=\"1\"><c r=\"A1\"><f>nm</f></c></row></sheetData></worksheet>";
  const std::vector<std::uint8_t> bytes = test::BuildOoxmlPackage(
      test::OoxmlContentTypes(test::kXlsxWorkbookContentType), kSheet, kWorkbook, test::kOneSheetWorkbookRels);
  auto loaded = read_ooxml(test::span_of(bytes));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  Workbook& wb = loaded.value().workbook;
  ASSERT_EQ(wb.defined_names().size(), 1U);
  EXPECT_EQ(wb.defined_names()[0].formula, "_xludf.IMAGE(\"https://example.invalid/x.png\")");
  auto recalc = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(recalc)) << recalc.error().message;
  const Value a1 = wb.sheet(0).resolve_cell_value(0U, 0U);
  ASSERT_TRUE(a1.is_error());
  EXPECT_EQ(a1.as_error(), ErrorCode::Name);
}

// Excel 365 saves NETWORKDAYS.INTL and WORKDAY.INTL bare in `<f>`: every such cell of these Excel-saved books reads
// as the built-in and recalculates to the value Excel cached.
TEST(StoragePrefixContainers, ExcelSavedBareIntlWorkdayCallsAreBuiltins) {
  const std::string dir =
      std::string(FORMULON_FIXTURES_DIR) + "/../oracle/external/ironcalc/fixtures/calc_tests/DATE_AND_TIME/";
  for (const char* book : {"NETWORKDAYS_NETWORKDAYS.INTL.xlsx", "WORKDAY_WORKDAY.INTL.xlsx"}) {
    SCOPED_TRACE(book);
    const std::vector<std::uint8_t> bytes = test::read_file_bytes(dir + book);
    ASSERT_FALSE(bytes.empty());
    std::string sheet_xml;
    const ::testing::AssertionResult got_sheet =
        test::extract_part(test::span_of(bytes), "xl/worksheets/sheet1.xml", &sheet_xml);
    ASSERT_TRUE(static_cast<bool>(got_sheet)) << got_sheet.message();
    auto loaded = read_ooxml(test::span_of(bytes));
    ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
    Workbook& wb = loaded.value().workbook;
    ASSERT_TRUE(static_cast<bool>(wb.recalc_full(eval::default_registry())));
    pugi::xml_document doc;
    ASSERT_TRUE(static_cast<bool>(doc.load_buffer(sheet_xml.data(), sheet_xml.size())));
    std::size_t checked = 0U;
    for (const pugi::xml_node row : doc.child("worksheet").child("sheetData").children("row")) {
      for (const pugi::xml_node cell : row.children("c")) {
        const std::string_view formula = cell.child("f").text().get();
        if (formula.rfind("NETWORKDAYS.INTL(", 0) != 0 && formula.rfind("WORKDAY.INTL(", 0) != 0) {
          continue;
        }
        const std::string ref = cell.attribute("r").value();
        // Its date-text arguments were cached #VALUE! by the saving host's regional settings; Mac Excel parses them.
        if (ref == "D60") {
          continue;
        }
        SCOPED_TRACE(ref);
        std::uint32_t col = 0U;
        std::size_t i = 0U;
        for (; i < ref.size() && ref[i] >= 'A' && ref[i] <= 'Z'; ++i) {
          col = col * 26U + static_cast<std::uint32_t>(ref[i] - 'A' + 1);
        }
        const std::uint32_t row_index = static_cast<std::uint32_t>(std::stoul(ref.substr(i))) - 1U;
        const Value got = wb.sheet(0).resolve_cell_value(row_index, col - 1U);
        const std::string cached = cell.child("v").text().get();
        if (std::string_view(cell.attribute("t").value()) == "e") {
          ASSERT_TRUE(got.is_error());
          EXPECT_EQ(std::optional<ErrorCode>(got.as_error()), error_from_display_name(cached));
        } else {
          ASSERT_TRUE(got.is_number());
          EXPECT_EQ(got.as_number(), std::stod(cached));
        }
        ++checked;
      }
    }
    EXPECT_GT(checked, 0U);
  }
}

// The two save bare in `<f>` and by an `_xlfn.` hidden name in `.xlsb`, as Excel saves them, and read back as the
// built-ins from both containers.
TEST(StoragePrefixContainers, IntlWorkdayCallsSaveAsExcelDoes) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.sheet(wb.add_sheet("F"));
  s.set_cell_formula(0U, 0U, "=NETWORKDAYS.INTL(45658,45672)");
  s.set_cell_formula(0U, 1U, "=WORKDAY.INTL(45658,10)");

  auto xlsx = wb.save();
  ASSERT_TRUE(static_cast<bool>(xlsx)) << xlsx.error().message;
  std::string sheet;
  const ::testing::AssertionResult got_sheet =
      test::extract_part(test::span_of(xlsx.value()), "xl/worksheets/sheet1.xml", &sheet);
  ASSERT_TRUE(static_cast<bool>(got_sheet)) << got_sheet.message();
  EXPECT_NE(sheet.find(">NETWORKDAYS.INTL(45658,45672)<"), std::string::npos) << sheet;
  EXPECT_NE(sheet.find(">WORKDAY.INTL(45658,10)<"), std::string::npos) << sheet;
  EXPECT_EQ(sheet.find("_xlfn."), std::string::npos) << sheet;

  auto xlsb = xlsb::write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(xlsb)) << xlsb.error().message;
  EXPECT_EQ(xlsb.value().diagnostics.downgraded_formula_count, 0U);
  std::string workbook;
  const ::testing::AssertionResult got_workbook =
      test::extract_part(test::span_of(xlsb.value().bytes), "xl/workbook.bin", &workbook);
  ASSERT_TRUE(static_cast<bool>(got_workbook)) << got_workbook.message();
  EXPECT_TRUE(ContainsUtf16Le(workbook, "_xlfn.NETWORKDAYS.INTL"));
  EXPECT_TRUE(ContainsUtf16Le(workbook, "_xlfn.WORKDAY.INTL"));

  auto from_xlsx = read_ooxml(test::span_of(xlsx.value()));
  ASSERT_TRUE(static_cast<bool>(from_xlsx)) << from_xlsx.error().message;
  auto from_xlsb = xlsb::read_xlsb(test::span_of(xlsb.value().bytes));
  ASSERT_TRUE(static_cast<bool>(from_xlsb)) << from_xlsb.error().message;
  for (Workbook* back : {&from_xlsx.value().workbook, &from_xlsb.value().workbook}) {
    ASSERT_TRUE(static_cast<bool>(back->recalc_full(eval::default_registry())));
    const Sheet& rt = back->sheet(0);
    for (std::uint32_t col = 0U; col < 2U; ++col) {
      SCOPED_TRACE(col);
      const Cell* cell = rt.cell_at(0U, col);
      ASSERT_NE(cell, nullptr);
      EXPECT_EQ(cell->formula_text, col == 0U ? "=NETWORKDAYS.INTL(45658,45672)" : "=WORKDAY.INTL(45658,10)");
      EXPECT_TRUE(rt.resolve_cell_value(0U, col).is_number());
    }
  }
}

// Excel saves an optional LAMBDA parameter's declaration as `_xlop.b`: it reads as `[b]`, evaluates, and saves back
// as Excel stored it.
TEST(StoragePrefixContainers, ExcelSavedOptionalLambdaParameter) {
  const std::vector<std::uint8_t> bytes = test::read_file_bytes(
      FORMULON_FIXTURES_DIR "/../oracle/external/ironcalc/fixtures/calc_tests/INFORMATION/ISOMITTED.xlsx");
  ASSERT_FALSE(bytes.empty());
  auto loaded = read_ooxml(test::span_of(bytes));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  Workbook& wb = loaded.value().workbook;
  ASSERT_TRUE(static_cast<bool>(wb.recalc_full(eval::default_registry())));
  const Sheet& s = wb.sheet(0);
  ASSERT_NE(s.cell_at(5U, 0U), nullptr);
  EXPECT_EQ(s.cell_at(5U, 0U)->formula_text, "=LAMBDA(a,[b],IF(ISOMITTED(b),a/2,a*b))(A5)");
  EXPECT_EQ(s.resolve_cell_value(5U, 0U).as_number(), 6.0);
  EXPECT_EQ(s.resolve_cell_value(6U, 0U).as_number(), 36.0);

  auto saved = wb.save();
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  std::string sheet;
  const ::testing::AssertionResult got_sheet =
      test::extract_part(test::span_of(saved.value()), "xl/worksheets/sheet1.xml", &sheet);
  ASSERT_TRUE(static_cast<bool>(got_sheet)) << got_sheet.message();
  EXPECT_NE(
      sheet.find(">_xlfn.LAMBDA(_xlpm.a,_xlop.b,IF(_xlfn.ISOMITTED(_xlpm.b),_xlpm.a/2,_xlpm.a*_xlpm.b))(A5, B5)<"),
      std::string::npos)
      << sheet;
}

// Both containers of one Excel-saved book declare the optional parameter as `_xlop.b` (an `.xlsb` `PtgName` to a
// hidden `BrtName`) and read as `[b]`; a re-save of either writes the same spelling.
TEST(StoragePrefixContainers, OptionalLambdaParameterInBothContainers) {
  const std::string kLambda = "LAMBDA(a,[b],IF(ISOMITTED(b),a/2,a*b))";
  const std::vector<std::uint8_t> xlsx =
      test::read_file_bytes(FORMULON_FIXTURES_DIR "/excel/lambda_optional_param.xlsx");
  const std::vector<std::uint8_t> xlsb =
      test::read_file_bytes(FORMULON_FIXTURES_DIR "/excel/lambda_optional_param.xlsb");
  ASSERT_FALSE(xlsx.empty());
  ASSERT_FALSE(xlsb.empty());
  auto from_xlsx = read_ooxml(test::span_of(xlsx));
  ASSERT_TRUE(static_cast<bool>(from_xlsx)) << from_xlsx.error().message;
  auto from_xlsb = xlsb::read_xlsb(test::span_of(xlsb));
  ASSERT_TRUE(static_cast<bool>(from_xlsb)) << from_xlsb.error().message;
  EXPECT_EQ(from_xlsb.value().undecoded_defined_name_count, 0U);

  auto check = [&](Workbook& wb) {
    ASSERT_TRUE(static_cast<bool>(wb.recalc_full(eval::default_registry())));
    std::vector<std::string> formulas;
    for (const DefinedName& dn : wb.defined_names()) {
      formulas.push_back(dn.name + "=" + dn.formula);
    }
    EXPECT_EQ(formulas, (std::vector<std::string>{"F=" + kLambda}));
    const Sheet& s = wb.sheet(0);
    ASSERT_NE(s.cell_at(5U, 0U), nullptr);
    EXPECT_EQ(s.cell_at(5U, 0U)->formula_text, "=" + kLambda + "(A5)");
    EXPECT_EQ(s.resolve_cell_value(5U, 0U).as_number(), 6.0);
    EXPECT_EQ(s.resolve_cell_value(6U, 0U).as_number(), 36.0);
    EXPECT_EQ(s.resolve_cell_value(7U, 0U).as_number(), 6.0);
  };
  {
    SCOPED_TRACE("xlsx");
    check(from_xlsx.value().workbook);
  }
  {
    SCOPED_TRACE("xlsb");
    check(from_xlsb.value().workbook);
  }

  for (Workbook* wb : {&from_xlsx.value().workbook, &from_xlsb.value().workbook}) {
    auto saved_xlsx = wb->save();
    ASSERT_TRUE(static_cast<bool>(saved_xlsx)) << saved_xlsx.error().message;
    std::string workbook_xml;
    const ::testing::AssertionResult got_xml =
        test::extract_part(test::span_of(saved_xlsx.value()), "xl/workbook.xml", &workbook_xml);
    ASSERT_TRUE(static_cast<bool>(got_xml)) << got_xml.message();
    EXPECT_NE(
        workbook_xml.find(">_xlfn.LAMBDA(_xlpm.a,_xlop.b,IF(_xlfn.ISOMITTED(_xlpm.b),_xlpm.a/2,_xlpm.a*_xlpm.b))<"),
        std::string::npos)
        << workbook_xml;

    auto saved_xlsb = xlsb::write_xlsb_with_result(*wb);
    ASSERT_TRUE(static_cast<bool>(saved_xlsb)) << saved_xlsb.error().message << " | " << saved_xlsb.error().context;
    EXPECT_EQ(saved_xlsb.value().diagnostics.downgraded_formula_count, 0U);
    std::string workbook_bin;
    const ::testing::AssertionResult got_bin =
        test::extract_part(test::span_of(saved_xlsb.value().bytes), "xl/workbook.bin", &workbook_bin);
    ASSERT_TRUE(static_cast<bool>(got_bin)) << got_bin.message();
    EXPECT_TRUE(ContainsUtf16Le(workbook_bin, "_xlop.b"));
    auto back = xlsb::read_xlsb(test::span_of(saved_xlsb.value().bytes));
    ASSERT_TRUE(static_cast<bool>(back)) << back.error().message;
    SCOPED_TRACE("xlsb re-save");
    check(back.value().workbook);
  }
}

}  // namespace
}  // namespace io
}  // namespace formulon
