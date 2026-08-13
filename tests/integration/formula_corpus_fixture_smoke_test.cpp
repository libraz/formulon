//
// Smoke test against a second real Mac Excel 365-produced fixture
// (`tests/fixtures/excel/formula_corpus.xlsx`), covering a function-family
// cross-section (text/ja, math, date, lookup, stats/financial) plus four
// dynamic-array spill shapes (SORT, UNIQUE, FILTER, TRANSPOSE).
//
// Test pattern: "cached value is the oracle". Rather than hand-transcribing
// 28+ expected values (error-prone and disconnected from what Excel
// actually computed), each formula cell's cached value -- the value Excel
// itself baked into the file on save -- is captured before `recalc()`, and
// Formulon's post-recalc value is compared against it (numeric: epsilon;
// text/error: exact).
//
// The `Corpus` sheet layout:
//   A1:A5 = 5, 3, 8, 1, 9; B1:B5 = "apples", "バナナ", "Cherry", "ﾃﾞｰﾀ", "mix"
//   E1:E28 = 28 formulas spanning text/ja, math, date, lookup, and
//     stats/financial families (see `FormulaCases()` for the exact list)
//   G1:G5 = SORT(A1:A5) spill, H1:H3 = UNIQUE({1;2;2;3;1}) spill,
//     I1:I3 = FILTER(A1:A5,A1:A5>4) spill, J1:L1 = TRANSPOSE(A1:A3) spill
//
// E4 is `JIS("ｱｲｳ123")`. Localized Excel provides JIS(), but this fixture was
// authored through Excel's invariant API with the localized name, leaving
// raw JIS and a `#NAME?` cache as evidence of invalid authoring. The direct
// E4 assertion is Formulon alias coverage, not an Excel-cache oracle.

#include <cmath>
#include <cstdint>
#include <cstdio>
#include <cstdlib>
#include <deque>
#include <map>
#include <string>
#include <vector>

#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/zip_reader.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "support/roundtrip_symmetry.h"
#include "value.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

std::string FixturePath() {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/formula_corpus.xlsx";
}

Workbook LoadFixture() {
  const std::vector<std::uint8_t> bytes = test::read_file_bytes(FixturePath());
  if (bytes.empty()) {
    return Workbook::create_empty();
  }
  auto result_or = io::read_ooxml(test::span_of(bytes));
  EXPECT_TRUE(static_cast<bool>(result_or)) << "read_ooxml failed: " << (result_or ? "" : result_or.error().message);
  if (!result_or) {
    return Workbook::create_empty();
  }
  return std::move(result_or.value().workbook);
}

// Parses `xl/worksheets/sheet1.xml` straight out of the fixture's raw zip
// bytes (independent of `LoadFixture`'s `Workbook`), for `CachedValueFromRawXml`.
pugi::xml_document LoadRawSheetXml() {
  pugi::xml_document doc;
  const std::vector<std::uint8_t> bytes = test::read_file_bytes(FixturePath());
  if (bytes.empty()) {
    return doc;
  }
  std::string xml;
  const ::testing::AssertionResult extracted =
      test::extract_part(test::span_of(bytes), "xl/worksheets/sheet1.xml", &xml);
  EXPECT_TRUE(static_cast<bool>(extracted)) << extracted.message();
  if (extracted) {
    const ::testing::AssertionResult parsed = test::parse_xml(xml, &doc);
    EXPECT_TRUE(static_cast<bool>(parsed)) << parsed.message();
  }
  return doc;
}

// `Cell::cached_value` is not populated for formula cells on load (the
// reader only calls `Workbook::set_cell_formula`, which leaves the cell
// blank until the first `recalc()`); the file's baked-in Excel cache is
// therefore read directly out of the raw part XML instead, using the
// exact same "cached value is the oracle" contract the rest of this file
// relies on. Handles the two shapes among this fixture's formula cells:
// plain numeric `<v>` (no `t=`) and `t="str"` (formula string result).
// `t="e"` (E4) is intentionally unsupported here: the fixture's invalid
// invariant-API authoring left localized `JIS` and a `#NAME?` cache. E4 is
// asserted directly for Formulon's alias behavior, not through this generic
// Excel-cache path.
//
// `Value::text` is a non-owning view: a `t="str"` result's backing bytes
// are pushed onto `text_storage` (a `deque`, so existing elements never
// move on further insertion) and the returned `Value` views that stable
// element, not the function-local decode buffer.
Value CachedValueFromRawXml(const pugi::xml_document& sheet_xml, const std::string& a1_ref,
                            std::deque<std::string>& text_storage) {
  const std::string xpath = "//c[@r='" + a1_ref + "']";
  const pugi::xpath_node xnode = sheet_xml.select_node(xpath.c_str());
  if (!xnode) {
    ADD_FAILURE() << "cell " << a1_ref << " not found in raw worksheet XML";
    return Value::blank();
  }
  const pugi::xml_node c = xnode.node();
  const pugi::xml_node v = c.child("v");
  if (!v) {
    return Value::blank();
  }
  const std::string_view t = c.attribute("t").value();
  if (t == "str") {
    text_storage.push_back(v.text().get());
    return Value::text(text_storage.back());
  }
  if (t == "e") {
    ADD_FAILURE() << "cell " << a1_ref << " is an error-typed cell; use a direct assertion instead of this helper";
    return Value::blank();
  }
  const std::string v_text = v.text().get();
  char* end = nullptr;
  const double d = std::strtod(v_text.c_str(), &end);
  return Value::number(d);
}

// Compares a cached (pre-recalc, Excel-authored) value against Formulon's
// recalculated value for the same cell: numbers within a magnitude-scaled
// epsilon (covers financial formulas carrying many decimal digits, whose
// exact bit pattern can legitimately differ by internal computation
// order), text and errors exactly.
void ExpectValuesMatch(const Value& cached, const Value& recalced) {
  if (cached.is_number()) {
    ASSERT_TRUE(recalced.is_number()) << "expected a number (cached=" << cached.as_number() << ")";
    const double c = cached.as_number();
    const double r = recalced.as_number();
    const double epsilon = std::max(1e-6, std::fabs(c) * 1e-9);
    EXPECT_NEAR(r, c, epsilon);
    return;
  }
  if (cached.is_text()) {
    ASSERT_TRUE(recalced.is_text()) << "expected text (cached=\"" << cached.as_text() << "\")";
    EXPECT_EQ(recalced.as_text(), cached.as_text());
    return;
  }
  if (cached.is_error()) {
    ASSERT_TRUE(recalced.is_error()) << "expected an error (cached=" << static_cast<int>(cached.as_error()) << ")";
    EXPECT_EQ(recalced.as_error(), cached.as_error());
    return;
  }
  if (cached.is_boolean()) {
    ASSERT_TRUE(recalced.is_boolean()) << "expected a boolean";
    EXPECT_EQ(recalced.as_boolean(), cached.as_boolean());
    return;
  }
  ADD_FAILURE() << "cached value has an unexpected kind";
}

struct FormulaCase {
  std::uint32_t row;  // 0-based
  const char* label;  // function family under test, for SCOPED_TRACE
};

// Column E (index 4), rows 1..28 (0-based 0..27). E4 (localized JIS, row 3)
// is excluded from this cached-value list because its invalid invariant-API
// cache is `#NAME?`; see the file-header comment.
const std::vector<FormulaCase>& FormulaCases() {
  static const std::vector<FormulaCase> kCases = {
      {0, "E1 LEN"},
      {1, "E2 LENB"},
      {2, "E3 ASC"},
      // row 3 (E4, localized JIS) intentionally omitted from cache comparison.
      {4, "E5 MID"},
      {5, "E6 TEXT(date)"},
      {6, "E7 TEXT(number)"},
      {7, "E8 TEXTJOIN"},
      {8, "E9 UPPER/LOWER"},
      {9, "E10 SUBSTITUTE"},
      {10, "E11 MROUND"},
      {11, "E12 ROUND"},
      {12, "E13 SUMPRODUCT"},
      {13, "E14 MOD"},
      {14, "E15 CEILING.MATH"},
      {15, "E16 DATE+1"},
      {16, "E17 EOMONTH"},
      {17, "E18 WEEKDAY"},
      {18, "E19 DATEDIF"},
      {19, "E20 YEARFRAC"},
      {20, "E21 INDEX/MATCH"},
      {21, "E22 IFERROR"},
      {22, "E23 XMATCH"},
      {23, "E24 CHOOSE"},
      {24, "E25 STDEV.P"},
      {25, "E26 PERCENTILE.INC"},
      {26, "E27 PMT"},
      {27, "E28 FV"},
  };
  return kCases;
}

// Collects `cell reference -> <f> element text` for every formula cell in
// a worksheet part. The text is pugixml's decoded form, so the comparison
// is on the formula as spelled rather than on the entity encoding: Excel
// leaves a `"` inside a string literal unescaped where Formulon's writer
// emits `&quot;`, which is the same XML text either way.
std::map<std::string, std::string> FormulaTextsByCell(const pugi::xml_document& sheet_xml) {
  std::map<std::string, std::string> out;
  for (const pugi::xpath_node& xnode : sheet_xml.select_nodes("//c[f]")) {
    const pugi::xml_node c = xnode.node();
    const std::string text = c.child("f").text().get();
    if (text.empty()) {
      continue;  // shared-formula follower: no text of its own.
    }
    out.emplace(c.attribute("r").value(), text);
  }
  return out;
}

constexpr std::uint32_t kColE = 4;
constexpr std::uint32_t kColG = 6;
constexpr std::uint32_t kColH = 7;
constexpr std::uint32_t kColI = 8;
constexpr std::uint32_t kColJ = 9;
constexpr std::uint32_t kColK = 10;
constexpr std::uint32_t kColL = 11;
constexpr std::uint32_t kRowE4 = 3;

void ExpectSpillShapesMatch(const Sheet& sheet) {
  // G1:G5 = SORT(A1:A5): 5,3,8,1,9 -> ascending 1,3,5,8,9.
  const double expected_sort[] = {1.0, 3.0, 5.0, 8.0, 9.0};
  for (std::uint32_t i = 0; i < 5; ++i) {
    SCOPED_TRACE(testing::Message() << "SORT spill row " << i);
    const Value v = sheet.resolve_cell_value(i, kColG);
    ASSERT_TRUE(v.is_number());
    EXPECT_DOUBLE_EQ(v.as_number(), expected_sort[i]);
  }

  // H1:H3 = UNIQUE({1;2;2;3;1}) -> 1,2,3.
  const double expected_unique[] = {1.0, 2.0, 3.0};
  for (std::uint32_t i = 0; i < 3; ++i) {
    SCOPED_TRACE(testing::Message() << "UNIQUE spill row " << i);
    const Value v = sheet.resolve_cell_value(i, kColH);
    ASSERT_TRUE(v.is_number());
    EXPECT_DOUBLE_EQ(v.as_number(), expected_unique[i]);
  }

  // I1:I3 = FILTER(A1:A5,A1:A5>4): of 5,3,8,1,9 the values > 4 are
  // 5, 8, 9 (three matches, not two -- confirmed against the fixture's
  // own baked-in cache: I1=5, I2=8, I3=9).
  const double expected_filter[] = {5.0, 8.0, 9.0};
  for (std::uint32_t i = 0; i < 3; ++i) {
    SCOPED_TRACE(testing::Message() << "FILTER spill row " << i);
    const Value v = sheet.resolve_cell_value(i, kColI);
    ASSERT_TRUE(v.is_number());
    EXPECT_DOUBLE_EQ(v.as_number(), expected_filter[i]);
  }

  // J1:L1 = TRANSPOSE(A1:A3): column {5,3,8} transposed to a row.
  const std::uint32_t transpose_cols[] = {kColJ, kColK, kColL};
  const double expected_transpose[] = {5.0, 3.0, 8.0};
  for (std::uint32_t i = 0; i < 3; ++i) {
    SCOPED_TRACE(testing::Message() << "TRANSPOSE spill col " << i);
    const Value v = sheet.resolve_cell_value(0U, transpose_cols[i]);
    ASSERT_TRUE(v.is_number());
    EXPECT_DOUBLE_EQ(v.as_number(), expected_transpose[i]);
  }
}

// ---------------------------------------------------------------------------
// (1)+(2) Cached (Excel-authored) values match Formulon's recalculated
// values across every formula-family cell. E4's invalid localized-JIS cache
// is not an Excel oracle; its Formulon alias semantics are asserted directly.
// ---------------------------------------------------------------------------

TEST(FormulaCorpusFixtureSmoke, CachedValuesMatchRecalcAcrossFunctionFamilies) {
  const pugi::xml_document sheet_xml = LoadRawSheetXml();
  Workbook wb = LoadFixture();
  ASSERT_EQ(wb.sheet_count(), 1U);

  std::deque<std::string> text_storage;
  std::vector<Value> cached_values;
  cached_values.reserve(FormulaCases().size());
  for (const FormulaCase& c : FormulaCases()) {
    const std::string ref = "E" + std::to_string(c.row + 1);
    cached_values.push_back(CachedValueFromRawXml(sheet_xml, ref, text_storage));
  }

  auto recalc_or = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(recalc_or)) << "recalc failed: " << recalc_or.error().message;

  const Sheet& sheet = wb.sheet(0);
  for (std::size_t i = 0; i < FormulaCases().size(); ++i) {
    SCOPED_TRACE(FormulaCases()[i].label);
    const Value recalced = sheet.resolve_cell_value(FormulaCases()[i].row, kColE);
    ExpectValuesMatch(cached_values[i], recalced);
  }
}

TEST(FormulaCorpusFixtureSmoke, E4LocalizedJisAliasConvertsHalfWidthToFullWidth) {
  // Localized Excel provides JIS(), but this fixture was authored through
  // Excel's invariant API with the localized name, leaving raw JIS and a
  // #NAME? cache. This is Formulon alias coverage, not an Excel-cache oracle.
  // The `<f t="array" aca="1" ref="E4" ca="1">` rich-array shape is also a
  // useful reader edge case, exercised by needing E4's formula to parse.
  Workbook wb = LoadFixture();
  ASSERT_EQ(wb.sheet_count(), 1U);
  auto recalc_or = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(recalc_or)) << "recalc failed: " << recalc_or.error().message;

  const Value e4 = wb.sheet(0).resolve_cell_value(kRowE4, kColE);
  ASSERT_TRUE(e4.is_text());
  EXPECT_EQ(e4.as_text(),
            "\xE3\x82\xA2\xE3\x82\xA4\xE3\x82\xA6\xEF\xBC\x91\xEF\xBC\x92\xEF\xBC\x93");  // "アイウ１２３"
}

// ---------------------------------------------------------------------------
// (3) The four spill families resolve to the correct values and shape.
// ---------------------------------------------------------------------------

TEST(FormulaCorpusFixtureSmoke, SpillFamiliesResolveWithCorrectShape) {
  Workbook wb = LoadFixture();
  ASSERT_EQ(wb.sheet_count(), 1U);
  auto recalc_or = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(recalc_or)) << "recalc failed: " << recalc_or.error().message;
  ExpectSpillShapesMatch(wb.sheet(0));
}

// ---------------------------------------------------------------------------
// (4) A save -> reload cycle does not disturb the recalculated values
// (cache-vs-recalc parity, E4, and spill shapes).
// ---------------------------------------------------------------------------

TEST(FormulaCorpusFixtureSmoke, SaveReloadPreservesRecalculatedValues) {
  Workbook wb = LoadFixture();
  ASSERT_EQ(wb.sheet_count(), 1U);
  auto recalc_or = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(recalc_or)) << "recalc failed: " << recalc_or.error().message;

  // Snapshot Formulon's own recalculated values as the post-reload ground
  // truth (independent of whether they matched Excel's cache -- that is
  // covered by the tests above; this test is purely about save/reload
  // stability).
  std::vector<Value> first_pass_values;
  first_pass_values.reserve(FormulaCases().size());
  for (const FormulaCase& c : FormulaCases()) {
    first_pass_values.push_back(wb.sheet(0).resolve_cell_value(c.row, kColE));
  }

  auto saved_or = wb.save();
  ASSERT_TRUE(static_cast<bool>(saved_or)) << "save failed: " << saved_or.error().message;

  auto reloaded_or = io::read_ooxml(test::span_of(saved_or.value()));
  ASSERT_TRUE(static_cast<bool>(reloaded_or)) << "reload after save failed: " << reloaded_or.error().message;
  Workbook reloaded = std::move(reloaded_or.value().workbook);
  auto reload_recalc_or = reloaded.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(reload_recalc_or))
      << "recalc after reload failed: " << reload_recalc_or.error().message;

  ASSERT_EQ(reloaded.sheet_count(), 1U);
  const Sheet& sheet = reloaded.sheet(0);
  for (std::size_t i = 0; i < FormulaCases().size(); ++i) {
    SCOPED_TRACE(FormulaCases()[i].label);
    const Value second_pass = sheet.resolve_cell_value(FormulaCases()[i].row, kColE);
    ExpectValuesMatch(first_pass_values[i], second_pass);
  }

  ExpectSpillShapesMatch(sheet);
}

// ---------------------------------------------------------------------------
// (5) Every `<f>` this fixture carries survives load -> recalc -> save with
// Excel's storage spelling, including E4's correction from the fixture's
// invalid invariant-API `JIS(...)` to canonical `DBCS(...)`. This corpus
// separates the two halves of the rule: MROUND / YEARFRAC / EOMONTH /
// TRANSPOSE / DATEDIF are Analysis ToolPak functions Excel 2007 made
// native and stores bare, while CEILING.MATH / STDEV.P / PERCENTILE.INC /
// XMATCH / TEXTJOIN / UNIQUE and the worksheet-only SORT / FILTER carry
// their `_xlfn.` / `_xlfn._xlws.` prefix. Prefixing a bare one, or
// dropping the prefix from a prefixed one, makes real Excel show #NAME?.
// ---------------------------------------------------------------------------

TEST(FormulaCorpusFixtureSmoke, SavedFormulaTextUsesExcelStorageSpellingForEveryCell) {
  const pugi::xml_document original_xml = LoadRawSheetXml();
  const std::map<std::string, std::string> original = FormulaTextsByCell(original_xml);
  ASSERT_FALSE(original.empty()) << "no <f> elements found in the fixture";
  const auto e4 = original.find("E4");
  ASSERT_NE(e4, original.end()) << "E4 formula missing from the fixture";
  ASSERT_EQ(e4->second, "JIS(\"ｱｲｳ123\")");
  std::map<std::string, std::string> expected = original;
  expected["E4"] = "DBCS(\"ｱｲｳ123\")";

  Workbook wb = LoadFixture();
  ASSERT_EQ(wb.sheet_count(), 1U);
  auto recalc_or = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(recalc_or)) << "recalc failed: " << recalc_or.error().message;
  auto saved_or = wb.save();
  ASSERT_TRUE(static_cast<bool>(saved_or)) << "save failed: " << saved_or.error().message;

  std::string saved_part;
  const ::testing::AssertionResult extracted =
      test::extract_part(test::span_of(saved_or.value()), "xl/worksheets/sheet1.xml", &saved_part);
  ASSERT_TRUE(static_cast<bool>(extracted)) << extracted.message();
  pugi::xml_document saved_xml;
  const ::testing::AssertionResult parsed = test::parse_xml(saved_part, &saved_xml);
  ASSERT_TRUE(static_cast<bool>(parsed)) << parsed.message();
  const std::map<std::string, std::string> saved = FormulaTextsByCell(saved_xml);

  for (const auto& entry : expected) {
    SCOPED_TRACE(entry.first);
    const auto it = saved.find(entry.first);
    ASSERT_NE(it, saved.end()) << "formula cell dropped on save";
    EXPECT_EQ(it->second, entry.second);
  }
  EXPECT_EQ(saved.size(), original.size()) << "save introduced formula cells the fixture does not have";

  // The prefixes themselves carry no escapable character, so they are also
  // assertable against the raw part bytes.
  EXPECT_NE(saved_part.find(">MROUND(17,5)<"), std::string::npos) << saved_part;
  EXPECT_EQ(saved_part.find("_xlfn.MROUND"), std::string::npos) << saved_part;
  EXPECT_EQ(saved_part.find("_xlfn.YEARFRAC"), std::string::npos) << saved_part;
  EXPECT_EQ(saved_part.find("_xlfn.TRANSPOSE"), std::string::npos) << saved_part;
  EXPECT_NE(saved_part.find("_xlfn.CEILING.MATH("), std::string::npos) << saved_part;
  EXPECT_NE(saved_part.find("_xlfn._xlws.SORT("), std::string::npos) << saved_part;
  EXPECT_NE(saved_part.find("_xlfn._xlws.FILTER("), std::string::npos) << saved_part;
  EXPECT_NE(saved_part.find("_xlfn.UNIQUE("), std::string::npos) << saved_part;
}

}  // namespace
}  // namespace formulon
