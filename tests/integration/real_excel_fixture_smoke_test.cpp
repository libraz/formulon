// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Smoke test against a real Mac Excel 365-produced `.xlsx`
// (`tests/fixtures/excel/xlsb_fidelity_base.xlsx`). Nearly every other
// integration test feeds the OOXML reader XML that Formulon's own
// writer produced, which cannot detect a defect the reader and writer
// share (e.g. a shape neither side emits/consumes). This test instead
// loads a file Excel itself wrote and checks the full pipeline: load,
// formula-text recovery, recalculated values, defined names, and style
// lookups -- with a save -> reload cycle to confirm none of it is an
// artifact of the first load alone.
//
// The fixture's `Data` sheet layout (see the task that added this file
// for the full authoring rationale):
//   A1:A3 = "k1"/"k2"/"k3", B1:B3 = 10/20/30
//   D1 = 44927 (yyyy/mm/dd), D2 = 1234.5 (#,##0.00), D3 = "bold-red"
//     (bold + red font), D4 = "filled" (yellow fill), D5 = 0.25 (0.0%)
//   F1 = XLOOKUP(...) -> 20, F2 = LET(...) -> 6,
//     F3 = TEXTJOIN(...) -> "k1,k2,k3", F4 = CONCAT(...) -> "k1k2k3",
//     F5 = IFS(...) -> "big", F6 = SEQUENCE(3) -> {1;2;3} spill
//   H1 = SUM({1,2;3,4}) -> 10, H2 = SUM(B1:B3) -> 60,
//     H3 = VLOOKUP(...) -> 30, H4 = 'S2'!A1*2 -> 14,
//     H5 = SUM('Data:S2'!B1) -> 10, H6 = B1*Rate -> 1
//     (defined name Rate = 0.1)
//   J1 = 1/0 -> #DIV/0!, J2 = NA() -> #N/A
// `S2` sheet: A1 = 7.

#include <cstdint>
#include <cstdio>
#include <set>
#include <string>
#include <string_view>
#include <vector>

#include "cell.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/defined_names.h"
#include "io/ooxml_reader.h"
#include "io/styles_reader.h"
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

std::string FixturePath() {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_fidelity_base.xlsx";
}

std::vector<std::uint8_t> ReadFileBytes(const std::string& path) {
  std::vector<std::uint8_t> out;
  FILE* f = std::fopen(path.c_str(), "rb");
  if (f == nullptr) {
    ADD_FAILURE() << "could not open fixture: " << path;
    return out;
  }
  std::fseek(f, 0, SEEK_END);
  const long size = std::ftell(f);
  std::fseek(f, 0, SEEK_SET);
  if (size > 0) {
    out.resize(static_cast<std::size_t>(size));
    const std::size_t read = std::fread(out.data(), 1, out.size(), f);
    if (read != out.size()) {
      ADD_FAILURE() << "short read on fixture: " << path;
      out.clear();
    }
  }
  std::fclose(f);
  return out;
}

io::ByteSpan SpanOf(const std::vector<std::uint8_t>& bytes) {
  return io::ByteSpan{bytes.data(), bytes.size()};
}

// Loads the fixture, asserting a clean (error-free) `read_ooxml`. Returns
// an invalid (0-sheet) workbook on failure so callers can still run
// (each failure is already reported via ASSERT/EXPECT inside).
Workbook LoadFixture() {
  const std::vector<std::uint8_t> bytes = ReadFileBytes(FixturePath());
  if (bytes.empty()) {
    return Workbook::create_empty();
  }
  auto result_or = io::read_ooxml(SpanOf(bytes));
  EXPECT_TRUE(static_cast<bool>(result_or)) << "read_ooxml failed: " << (result_or ? "" : result_or.error().message);
  if (!result_or) {
    return Workbook::create_empty();
  }
  return std::move(result_or.value().workbook);
}

Workbook LoadAndRecalcFixture() {
  Workbook wb = LoadFixture();
  auto recalc_or = wb.recalc(eval::default_registry());
  EXPECT_TRUE(static_cast<bool>(recalc_or)) << "recalc failed: " << (recalc_or ? "" : recalc_or.error().message);
  return wb;
}

// ---------------------------------------------------------------------------
// (a) Load succeeds, no repair; two sheets in document order.
// ---------------------------------------------------------------------------

TEST(RealExcelFixtureSmoke, LoadsSuccessfullyWithTwoSheetsInOrder) {
  Workbook wb = LoadFixture();
  ASSERT_EQ(wb.sheet_count(), 2U);
  EXPECT_EQ(wb.sheet(0).name(), "Data");
  EXPECT_EQ(wb.sheet(1).name(), "S2");
}

// ---------------------------------------------------------------------------
// (b) Formula text is restored: no `_xlfn.` / `_xlpm.` storage prefix, and
// the modern function / parameter names are exactly what Excel's formula
// bar would show.
// ---------------------------------------------------------------------------

TEST(RealExcelFixtureSmoke, FormulaTextIsRestoredWithoutXlfnPrefix) {
  Workbook wb = LoadFixture();
  ASSERT_EQ(wb.sheet_count(), 2U);
  const Sheet& data = wb.sheet(0);

  const Cell* f1 = data.cell_at(0U, 5U);  // F1
  ASSERT_NE(f1, nullptr);
  EXPECT_EQ(f1->formula_text, "=XLOOKUP(\"k2\",A1:A3,B1:B3)");

  const Cell* f2 = data.cell_at(1U, 5U);  // F2
  ASSERT_NE(f2, nullptr);
  EXPECT_EQ(f2->formula_text, "=LET(x,2,x*3)");

  const Cell* f3 = data.cell_at(2U, 5U);  // F3
  ASSERT_NE(f3, nullptr);
  EXPECT_EQ(f3->formula_text, "=TEXTJOIN(\",\",TRUE,A1:A3)");

  const Cell* f4 = data.cell_at(3U, 5U);  // F4
  ASSERT_NE(f4, nullptr);
  EXPECT_EQ(f4->formula_text, "=CONCAT(A1:A3)");

  const Cell* f5 = data.cell_at(4U, 5U);  // F5
  ASSERT_NE(f5, nullptr);
  EXPECT_EQ(f5->formula_text, "=IFS(B1>5,\"big\",TRUE,\"small\")");

  const Cell* f6 = data.cell_at(5U, 5U);  // F6
  ASSERT_NE(f6, nullptr);
  EXPECT_EQ(f6->formula_text, "=SEQUENCE(3)");
}

// ---------------------------------------------------------------------------
// (c) Recalculated values match what Excel itself computed, including the
// two error-literal formulas (down to the exact error kind) and the
// SEQUENCE(3) spill.
// ---------------------------------------------------------------------------

// Each formula cell gets its own TEST so one failing recalculation does
// not mask the pass/fail status of the others (an early fatal ASSERT_
// inside a single shared test would stop the whole function).

TEST(RealExcelFixtureSmoke, F1XlookupRecalculatesTo20) {
  Workbook wb = LoadAndRecalcFixture();
  ASSERT_EQ(wb.sheet_count(), 2U);
  const Value f1 = wb.sheet(0).resolve_cell_value(0U, 5U);
  ASSERT_TRUE(f1.is_number());
  EXPECT_DOUBLE_EQ(f1.as_number(), 20.0);
}

TEST(RealExcelFixtureSmoke, F2LetRecalculatesTo6) {
  Workbook wb = LoadAndRecalcFixture();
  ASSERT_EQ(wb.sheet_count(), 2U);
  const Value f2 = wb.sheet(0).resolve_cell_value(1U, 5U);
  ASSERT_TRUE(f2.is_number());
  EXPECT_DOUBLE_EQ(f2.as_number(), 6.0);
}

TEST(RealExcelFixtureSmoke, F3TextjoinRecalculatesToJoinedString) {
  Workbook wb = LoadAndRecalcFixture();
  ASSERT_EQ(wb.sheet_count(), 2U);
  const Value f3 = wb.sheet(0).resolve_cell_value(2U, 5U);
  ASSERT_TRUE(f3.is_text());
  EXPECT_EQ(f3.as_text(), "k1,k2,k3");
}

TEST(RealExcelFixtureSmoke, F4ConcatRecalculatesToConcatenatedString) {
  Workbook wb = LoadAndRecalcFixture();
  ASSERT_EQ(wb.sheet_count(), 2U);
  const Value f4 = wb.sheet(0).resolve_cell_value(3U, 5U);
  ASSERT_TRUE(f4.is_text());
  EXPECT_EQ(f4.as_text(), "k1k2k3");
}

TEST(RealExcelFixtureSmoke, F5IfsRecalculatesToBig) {
  Workbook wb = LoadAndRecalcFixture();
  ASSERT_EQ(wb.sheet_count(), 2U);
  const Value f5 = wb.sheet(0).resolve_cell_value(4U, 5U);
  ASSERT_TRUE(f5.is_text());
  EXPECT_EQ(f5.as_text(), "big");
}

TEST(RealExcelFixtureSmoke, F6SequenceSpillsAcrossThreeRows) {
  Workbook wb = LoadAndRecalcFixture();
  ASSERT_EQ(wb.sheet_count(), 2U);
  const Sheet& data = wb.sheet(0);
  const Value f6 = data.resolve_cell_value(5U, 5U);
  ASSERT_TRUE(f6.is_number()) << "F6 (SEQUENCE spill row 0)";
  EXPECT_DOUBLE_EQ(f6.as_number(), 1.0);
  const Value f7 = data.resolve_cell_value(6U, 5U);
  ASSERT_TRUE(f7.is_number()) << "F7 (SEQUENCE spill row 1)";
  EXPECT_DOUBLE_EQ(f7.as_number(), 2.0);
  const Value f8 = data.resolve_cell_value(7U, 5U);
  ASSERT_TRUE(f8.is_number()) << "F8 (SEQUENCE spill row 2)";
  EXPECT_DOUBLE_EQ(f8.as_number(), 3.0);
}

TEST(RealExcelFixtureSmoke, H1SumArrayLiteralIs10) {
  Workbook wb = LoadAndRecalcFixture();
  ASSERT_EQ(wb.sheet_count(), 2U);
  const Value h1 = wb.sheet(0).resolve_cell_value(0U, 7U);
  ASSERT_TRUE(h1.is_number());
  EXPECT_DOUBLE_EQ(h1.as_number(), 10.0);
}

TEST(RealExcelFixtureSmoke, H2SumRangeIs60) {
  Workbook wb = LoadAndRecalcFixture();
  ASSERT_EQ(wb.sheet_count(), 2U);
  const Value h2 = wb.sheet(0).resolve_cell_value(1U, 7U);
  ASSERT_TRUE(h2.is_number());
  EXPECT_DOUBLE_EQ(h2.as_number(), 60.0);
}

TEST(RealExcelFixtureSmoke, H3VlookupIs30) {
  Workbook wb = LoadAndRecalcFixture();
  ASSERT_EQ(wb.sheet_count(), 2U);
  const Value h3 = wb.sheet(0).resolve_cell_value(2U, 7U);
  ASSERT_TRUE(h3.is_number());
  EXPECT_DOUBLE_EQ(h3.as_number(), 30.0);
}

TEST(RealExcelFixtureSmoke, H4CrossSheetReferenceIs14) {
  Workbook wb = LoadAndRecalcFixture();
  ASSERT_EQ(wb.sheet_count(), 2U);
  const Value h4 = wb.sheet(0).resolve_cell_value(3U, 7U);
  ASSERT_TRUE(h4.is_number());
  EXPECT_DOUBLE_EQ(h4.as_number(), 14.0);
}

TEST(RealExcelFixtureSmoke, H5ThreeDReferenceSumIs10) {
  Workbook wb = LoadAndRecalcFixture();
  ASSERT_EQ(wb.sheet_count(), 2U);
  const Value h5 = wb.sheet(0).resolve_cell_value(4U, 7U);
  ASSERT_TRUE(h5.is_number());
  EXPECT_DOUBLE_EQ(h5.as_number(), 10.0);
}

TEST(RealExcelFixtureSmoke, H6DefinedNameRateAppliesIs1) {
  Workbook wb = LoadAndRecalcFixture();
  ASSERT_EQ(wb.sheet_count(), 2U);
  const Value h6 = wb.sheet(0).resolve_cell_value(5U, 7U);
  ASSERT_TRUE(h6.is_number());
  EXPECT_DOUBLE_EQ(h6.as_number(), 1.0);
}

TEST(RealExcelFixtureSmoke, J1DivisionByZeroIsDiv0Error) {
  Workbook wb = LoadAndRecalcFixture();
  ASSERT_EQ(wb.sheet_count(), 2U);
  const Value j1 = wb.sheet(0).resolve_cell_value(0U, 9U);
  ASSERT_TRUE(j1.is_error());
  EXPECT_EQ(j1.as_error(), ErrorCode::Div0);
}

TEST(RealExcelFixtureSmoke, J2NaFunctionIsNaError) {
  Workbook wb = LoadAndRecalcFixture();
  ASSERT_EQ(wb.sheet_count(), 2U);
  const Value j2 = wb.sheet(0).resolve_cell_value(1U, 9U);
  ASSERT_TRUE(j2.is_error());
  EXPECT_EQ(j2.as_error(), ErrorCode::NA);
}

// ---------------------------------------------------------------------------
// (d) The `Rate` defined name is present and actually drives evaluation
// (already exercised indirectly via H6 above; this pins the metadata
// directly).
// ---------------------------------------------------------------------------

TEST(RealExcelFixtureSmoke, DefinedNameRateIsPresentAndUsedByH6) {
  Workbook wb = LoadFixture();
  bool found = false;
  for (const io::DefinedName& dn : wb.defined_names()) {
    if (dn.name == "Rate") {
      found = true;
      EXPECT_EQ(dn.formula, "0.1");
    }
  }
  EXPECT_TRUE(found) << "defined name 'Rate' not found";

  auto recalc_or = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(recalc_or)) << "recalc failed: " << recalc_or.error().message;
  const Value h6 = wb.sheet(0).resolve_cell_value(5U, 7U);  // H1*Rate == B1*Rate == 10*0.1
  ASSERT_TRUE(h6.is_number());
  EXPECT_DOUBLE_EQ(h6.as_number(), 1.0);
}

// ---------------------------------------------------------------------------
// (e) Styles: D1/D2/D5 number-format strings, D3 bold font, D4 fill colour
// (must not have degraded to black / no-fill).
// ---------------------------------------------------------------------------

// Resolves the effective `formatCode` for `num_fmt_id`: a built-in id
// (0..163) resolves via the static builtin table; anything else is looked
// up in the styles table's own `<numFmts>` records.
std::string NumFmtStringFor(const io::StylesTable& styles, std::uint16_t num_fmt_id) {
  if (const char* builtin = io::builtin_num_fmt(num_fmt_id); builtin != nullptr && *builtin != '\0') {
    return builtin;
  }
  for (const io::NumFmtRecord& rec : styles.num_fmts) {
    if (rec.id == num_fmt_id && rec.format_string_index < styles.num_fmt_strings.size()) {
      return styles.num_fmt_strings[rec.format_string_index];
    }
  }
  return {};
}

TEST(RealExcelFixtureSmoke, StylesSurviveLoad) {
  Workbook wb = LoadFixture();
  const io::StylesTable& styles = wb.styles();
  const Sheet& data = wb.sheet(0);

  const Cell* d1 = data.cell_at(0U, 3U);
  ASSERT_NE(d1, nullptr);
  ASSERT_LT(d1->xf_index, styles.cell_xfs.size());
  EXPECT_EQ(NumFmtStringFor(styles, styles.cell_xfs[d1->xf_index].num_fmt_id), "yyyy/mm/dd");

  const Cell* d2 = data.cell_at(1U, 3U);
  ASSERT_NE(d2, nullptr);
  ASSERT_LT(d2->xf_index, styles.cell_xfs.size());
  EXPECT_EQ(NumFmtStringFor(styles, styles.cell_xfs[d2->xf_index].num_fmt_id), "#,##0.00");

  const Cell* d3 = data.cell_at(2U, 3U);
  ASSERT_NE(d3, nullptr);
  ASSERT_LT(d3->xf_index, styles.cell_xfs.size());
  const std::uint32_t d3_font = styles.cell_xfs[d3->xf_index].font_index;
  ASSERT_LT(d3_font, styles.fonts.size());
  EXPECT_TRUE(styles.fonts[d3_font].bold);
  EXPECT_EQ(styles.fonts[d3_font].color_argb, 0xFFFF0000U);

  const Cell* d4 = data.cell_at(3U, 3U);
  ASSERT_NE(d4, nullptr);
  ASSERT_LT(d4->xf_index, styles.cell_xfs.size());
  const std::uint32_t d4_fill = styles.cell_xfs[d4->xf_index].fill_index;
  ASSERT_LT(d4_fill, styles.fills.size());
  // Must be the authored yellow solid fill, not degraded to black
  // (0xFF000000) or no-fill (pattern 0 / fg_argb 0).
  EXPECT_EQ(styles.fills[d4_fill].pattern, 1U);
  EXPECT_EQ(styles.fills[d4_fill].fg_argb, 0xFFFFFF00U);

  const Cell* d5 = data.cell_at(4U, 3U);
  ASSERT_NE(d5, nullptr);
  ASSERT_LT(d5->xf_index, styles.cell_xfs.size());
  EXPECT_EQ(NumFmtStringFor(styles, styles.cell_xfs[d5->xf_index].num_fmt_id), "0.0%");
}

// ---------------------------------------------------------------------------
// (f) A save -> reload cycle must not disturb formula text or recalculated
// values.
// ---------------------------------------------------------------------------

TEST(RealExcelFixtureSmoke, SaveReloadPreservesFormulaTextAndValues) {
  Workbook wb = LoadAndRecalcFixture();
  ASSERT_EQ(wb.sheet_count(), 2U);

  auto saved_or = wb.save();
  ASSERT_TRUE(static_cast<bool>(saved_or)) << "save failed: " << saved_or.error().message;

  auto reloaded_or = io::read_ooxml(SpanOf(saved_or.value()));
  ASSERT_TRUE(static_cast<bool>(reloaded_or)) << "reload after save failed: " << reloaded_or.error().message;
  Workbook reloaded = std::move(reloaded_or.value().workbook);
  auto recalc_or = reloaded.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(recalc_or)) << "recalc after reload failed: " << recalc_or.error().message;

  ASSERT_EQ(reloaded.sheet_count(), 2U);
  const Sheet& data = reloaded.sheet(0);

  const Cell* f1 = data.cell_at(0U, 5U);
  ASSERT_NE(f1, nullptr);
  EXPECT_EQ(f1->formula_text, "=XLOOKUP(\"k2\",A1:A3,B1:B3)");

  const Value f1_value = data.resolve_cell_value(0U, 5U);
  ASSERT_TRUE(f1_value.is_number());
  EXPECT_DOUBLE_EQ(f1_value.as_number(), 20.0);

  const Value h6_value = data.resolve_cell_value(5U, 7U);
  ASSERT_TRUE(h6_value.is_number());
  EXPECT_DOUBLE_EQ(h6_value.as_number(), 1.0);

  const Value j1_value = data.resolve_cell_value(0U, 9U);
  ASSERT_TRUE(j1_value.is_error());
  EXPECT_EQ(j1_value.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// (d) Save re-applies Excel's hidden storage prefixes so the written .xlsx
// is readable by real Excel (future functions carry `_xlfn.`, LET / LAMBDA
// parameters carry `_xlpm.`). The reader strips them again, so the reloaded
// formula_text stays canonical (asserted in SaveReloadPreserves... above);
// this test inspects the saved worksheet XML directly.
// ---------------------------------------------------------------------------

TEST(RealExcelFixtureSmoke, SaveReAppliesStoragePrefixesForExcel) {
  Workbook wb = LoadFixture();
  ASSERT_EQ(wb.sheet_count(), 2U);

  auto saved_or = wb.save();
  ASSERT_TRUE(static_cast<bool>(saved_or)) << "save failed: " << saved_or.error().message;

  io::ZipReader zip;
  auto open_or = zip.open(SpanOf(saved_or.value()));
  ASSERT_TRUE(static_cast<bool>(open_or)) << "zip open failed: " << open_or.error().message;
  auto sheet_or = zip.read_entry("xl/worksheets/sheet1.xml");
  ASSERT_TRUE(static_cast<bool>(sheet_or)) << "sheet1.xml read failed: " << sheet_or.error().message;
  const std::string xml(reinterpret_cast<const char*>(sheet_or.value().data()), sheet_or.value().size());

  // Future functions are re-tagged with `_xlfn.`; LET parameters with
  // `_xlpm.`. Classic functions (SUM / VLOOKUP) stay bare.
  EXPECT_NE(xml.find("_xlfn.XLOOKUP("), std::string::npos) << xml;
  EXPECT_NE(xml.find("_xlfn.LET("), std::string::npos) << xml;
  EXPECT_NE(xml.find("_xlpm."), std::string::npos) << xml;
  EXPECT_NE(xml.find("_xlfn.SEQUENCE("), std::string::npos) << xml;
  EXPECT_NE(xml.find("_xlfn.TEXTJOIN("), std::string::npos) << xml;
  EXPECT_EQ(xml.find("_xlfn.SUM("), std::string::npos) << xml;
  EXPECT_EQ(xml.find("_xlfn.VLOOKUP("), std::string::npos) << xml;
}

// ---------------------------------------------------------------------------
// (g) The saved package must be openable by real Excel: workbook-level
// metadata round-trips, the stale calcChain cache is dropped, and no
// rich-error (`#SPILL!` / `#CALC!`) cached value is written as a bare
// `<v>` (any of which makes Excel refuse / "repair" the file).
// ---------------------------------------------------------------------------

TEST(RealExcelFixtureSmoke, SaveOutputIsExcelOpenable) {
  Workbook wb = LoadAndRecalcFixture();
  ASSERT_EQ(wb.sheet_count(), 2U);
  auto saved_or = wb.save();
  ASSERT_TRUE(static_cast<bool>(saved_or)) << "save failed: " << saved_or.error().message;

  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(saved_or.value()))));

  auto part_sv = [&zip](const char* name) -> std::string {
    auto b = zip.read_entry(name);
    EXPECT_TRUE(static_cast<bool>(b)) << "missing part: " << name;
    if (!b) {
      return {};
    }
    return std::string(reinterpret_cast<const char*>(b.value().data()), b.value().size());
  };

  // workbookPr + bookViews (activeTab) must survive the real save path
  // (not just the unit-test path — this pins the CLI regression).
  const std::string wb_xml = part_sv("xl/workbook.xml");
  EXPECT_NE(wb_xml.find("<workbookPr defaultThemeVersion=\"202300\""), std::string::npos) << wb_xml;
  EXPECT_NE(wb_xml.find("activeTab=\"1\""), std::string::npos) << wb_xml;
  // The raw <bookViews> carries `xr2:uid`; its `xr2` prefix must be declared
  // on the re-emitted <workbook> root or the file is malformed XML that
  // Excel refuses. (Namespace decls round-trip from the source root.)
  if (wb_xml.find("xr2:uid") != std::string::npos) {
    EXPECT_NE(wb_xml.find("xmlns:xr2="), std::string::npos)
        << "xr2:uid used but xr2 namespace not declared on <workbook> root: " << wb_xml;
  }

  // Stale calcChain must be dropped: no part, no relationship, no Override.
  EXPECT_FALSE(zip.has_entry("xl/calcChain.xml"));
  EXPECT_EQ(part_sv("xl/_rels/workbook.xml.rels").find("calcChain"), std::string::npos);
  EXPECT_EQ(part_sv("[Content_Types].xml").find("calcChain"), std::string::npos);

  // No rich-error cached value written anywhere (F6 SEQUENCE spill must be
  // a valid array anchor, not a `t="e">#SPILL!` cell).
  const std::string sheet_xml = part_sv("xl/worksheets/sheet1.xml");
  EXPECT_EQ(sheet_xml.find("#SPILL!"), std::string::npos) << sheet_xml;
  EXPECT_EQ(sheet_xml.find("#CALC!"), std::string::npos) << sheet_xml;
  // Legacy errors (#DIV/0!, #N/A) are still legitimately present.
  EXPECT_NE(sheet_xml.find("#DIV/0!"), std::string::npos) << "legacy error should still round-trip";
}

// Recursively collects the namespace prefixes declared (`xmlns:PREFIX`) and
// used (on element / attribute names) within `node`'s subtree.
void CollectNamespacePrefixes(const pugi::xml_node& node, std::set<std::string>& declared,
                              std::set<std::string>& used) {
  const std::string_view ename = node.name();
  const std::size_t ecolon = ename.find(':');
  if (ecolon != std::string_view::npos) {
    used.emplace(ename.substr(0, ecolon));
  }
  for (pugi::xml_attribute a = node.first_attribute(); a; a = a.next_attribute()) {
    const std::string_view an = a.name();
    if (an == "xmlns") {
      continue;  // default namespace, no prefix.
    }
    if (an.rfind("xmlns:", 0) == 0) {
      declared.emplace(an.substr(6));
      continue;
    }
    const std::size_t acolon = an.find(':');
    if (acolon != std::string_view::npos) {
      used.emplace(an.substr(0, acolon));
    }
  }
  for (pugi::xml_node ch = node.first_child(); ch; ch = ch.next_sibling()) {
    CollectNamespacePrefixes(ch, declared, used);
  }
}

// Every namespace prefix used by any element / attribute in every XML part
// Formulon writes must be declared, or the part is malformed XML that real
// Excel refuses. This is the general guard against the `xr2:uid` class of
// bug: a raw-captured fragment carrying a prefix the writer's root did not
// re-declare.
TEST(RealExcelFixtureSmoke, SavedPartsHaveAllNamespacePrefixesDeclared) {
  Workbook wb = LoadAndRecalcFixture();
  ASSERT_EQ(wb.sheet_count(), 2U);
  auto saved_or = wb.save();
  ASSERT_TRUE(static_cast<bool>(saved_or)) << "save failed: " << saved_or.error().message;

  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(saved_or.value()))));

  for (const std::string& name : zip.list_entries()) {
    if (name.size() < 4U ||
        (name.compare(name.size() - 4U, 4U, ".xml") != 0 && name.find(".rels") == std::string::npos)) {
      continue;  // only XML-shaped parts.
    }
    auto bytes_or = zip.read_entry(name);
    ASSERT_TRUE(static_cast<bool>(bytes_or)) << name;
    pugi::xml_document doc;
    ASSERT_TRUE(doc.load_buffer(bytes_or.value().data(), bytes_or.value().size())) << "unparseable: " << name;

    std::set<std::string> declared;
    std::set<std::string> used;
    for (pugi::xml_node root = doc.first_child(); root; root = root.next_sibling()) {
      CollectNamespacePrefixes(root, declared, used);
    }
    // `xml` is implicitly declared per the XML spec.
    declared.emplace("xml");
    for (const std::string& prefix : used) {
      EXPECT_NE(declared.find(prefix), declared.end())
          << "part " << name << " uses undeclared namespace prefix '" << prefix << "'";
    }
  }
}

}  // namespace
}  // namespace formulon
