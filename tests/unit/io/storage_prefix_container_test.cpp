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
#include <string>
#include <vector>

#include "cell.h"
#include "gtest/gtest.h"
#include "io/future_functions.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "sheet.h"
#include "support/roundtrip_symmetry.h"
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

}  // namespace
}  // namespace io
}  // namespace formulon
