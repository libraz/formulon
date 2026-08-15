//
// Fidelity tests against a real Mac Excel 365-produced `.xlsb`
// (`tests/fixtures/excel/xlsb_fidelity_base.xlsb`). Mirrors
// `tests/integration/real_excel_fixture_smoke_test.cpp` (the `.xlsx`
// sibling of the same workbook) but exercises `io::xlsb::read_xlsb` /
// `io::xlsb::write_xlsb` instead of the OOXML reader/writer, so a
// defect specific to the binary Ptg/record path (future-function
// `id == 255` dispatch, `PtgArray` `RgbExtra`, `PtgRef3d` /
// `BrtExternSheet` resolution, `BrtName`-backed defined names, or the
// minimal `xl/styles.bin` reader) is caught even when the OOXML path
// is fine.
//
// The fixture's `Data` sheet layout (byte-level verified against the
// real package; see `real_excel_fixture_smoke_test.cpp` for the full
// authoring rationale):
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
#include <string>
#include <vector>

#include "cell.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/defined_names.h"
#include "io/format_detect.h"
#include "io/styles_reader.h"
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

std::string FixturePath() {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_fidelity_base.xlsb";
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

// Loads the fixture, asserting a clean (error-free) `read_xlsb`. Returns
// an invalid (0-sheet) workbook on failure so callers can still run
// (each failure is already reported via ASSERT/EXPECT inside).
Workbook LoadFixture() {
  const std::vector<std::uint8_t> bytes = ReadFileBytes(FixturePath());
  if (bytes.empty()) {
    return Workbook::create_empty();
  }
  auto result_or = io::xlsb::read_xlsb(SpanOf(bytes));
  EXPECT_TRUE(static_cast<bool>(result_or)) << "read_xlsb failed: " << (result_or ? "" : result_or.error().message);
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

TEST(XlsbFidelity, LoadsSuccessfullyWithTwoSheetsInOrder) {
  Workbook wb = LoadFixture();
  ASSERT_EQ(wb.sheet_count(), 2U);
  EXPECT_EQ(wb.sheet(0).name(), "Data");
  EXPECT_EQ(wb.sheet(1).name(), "S2");
}

// ---------------------------------------------------------------------------
// (b) Formula text is restored for every Ptg-encoded form this bundle
// closes: future-function `id == 255` dispatch (XLOOKUP / TEXTJOIN /
// CONCAT / IFS / SEQUENCE), `LetBinding`, `PtgArray` via `RgbExtra`,
// plain `PtgArea` ranges, `PtgRef3d` (single- and multi-sheet), and
// `PtgName` defined-name references.
// ---------------------------------------------------------------------------

TEST(XlsbFidelity, FormulaTextIsRestoredWithoutXlfnPrefix) {
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

  const Cell* h1 = data.cell_at(0U, 7U);  // H1
  ASSERT_NE(h1, nullptr);
  EXPECT_EQ(h1->formula_text, "=SUM({1,2;3,4})");

  const Cell* h2 = data.cell_at(1U, 7U);  // H2
  ASSERT_NE(h2, nullptr);
  EXPECT_EQ(h2->formula_text, "=SUM(B1:B3)");

  const Cell* h3 = data.cell_at(2U, 7U);  // H3
  ASSERT_NE(h3, nullptr);
  EXPECT_EQ(h3->formula_text, "=VLOOKUP(\"k3\",A1:B3,2,FALSE)");

  // `S2` is quoted because it is otherwise ambiguous with a cell
  // reference (column S, row 2); this matches real Excel exactly --
  // `xl/worksheets/sheet1.xml` in the sibling `.xlsx` fixture stores
  // `<f>'S2'!A1*2</f>`, never the bare `S2!A1*2` form.
  const Cell* h4 = data.cell_at(3U, 7U);  // H4
  ASSERT_NE(h4, nullptr);
  EXPECT_EQ(h4->formula_text, "='S2'!A1*2");

  // The whole `Data:S2` span quotes as one unit (not `Data:'S2'`)
  // because `S2` alone needs quoting; also verified byte-for-byte
  // against the sibling `.xlsx` fixture's `<f>SUM('Data:S2'!B1)</f>`.
  const Cell* h5 = data.cell_at(4U, 7U);  // H5
  ASSERT_NE(h5, nullptr);
  EXPECT_EQ(h5->formula_text, "=SUM('Data:S2'!B1)");

  const Cell* h6 = data.cell_at(5U, 7U);  // H6
  ASSERT_NE(h6, nullptr);
  EXPECT_EQ(h6->formula_text, "=B1*Rate");

  const Cell* j1 = data.cell_at(0U, 9U);  // J1
  ASSERT_NE(j1, nullptr);
  EXPECT_EQ(j1->formula_text, "=1/0");

  const Cell* j2 = data.cell_at(1U, 9U);  // J2
  ASSERT_NE(j2, nullptr);
  EXPECT_EQ(j2->formula_text, "=NA()");
}

// ---------------------------------------------------------------------------
// (c) Recalculated values match what Excel itself computed.
// ---------------------------------------------------------------------------

TEST(XlsbFidelity, F1XlookupRecalculatesTo20) {
  Workbook wb = LoadAndRecalcFixture();
  ASSERT_EQ(wb.sheet_count(), 2U);
  const Value f1 = wb.sheet(0).resolve_cell_value(0U, 5U);
  ASSERT_TRUE(f1.is_number());
  EXPECT_DOUBLE_EQ(f1.as_number(), 20.0);
}

TEST(XlsbFidelity, F2LetRecalculatesTo6) {
  Workbook wb = LoadAndRecalcFixture();
  ASSERT_EQ(wb.sheet_count(), 2U);
  const Value f2 = wb.sheet(0).resolve_cell_value(1U, 5U);
  ASSERT_TRUE(f2.is_number());
  EXPECT_DOUBLE_EQ(f2.as_number(), 6.0);
}

TEST(XlsbFidelity, F3TextjoinRecalculatesToJoinedString) {
  Workbook wb = LoadAndRecalcFixture();
  ASSERT_EQ(wb.sheet_count(), 2U);
  const Value f3 = wb.sheet(0).resolve_cell_value(2U, 5U);
  ASSERT_TRUE(f3.is_text());
  EXPECT_EQ(f3.as_text(), "k1,k2,k3");
}

TEST(XlsbFidelity, F4ConcatRecalculatesToConcatenatedString) {
  Workbook wb = LoadAndRecalcFixture();
  ASSERT_EQ(wb.sheet_count(), 2U);
  const Value f4 = wb.sheet(0).resolve_cell_value(3U, 5U);
  ASSERT_TRUE(f4.is_text());
  EXPECT_EQ(f4.as_text(), "k1k2k3");
}

TEST(XlsbFidelity, F5IfsRecalculatesToBig) {
  Workbook wb = LoadAndRecalcFixture();
  ASSERT_EQ(wb.sheet_count(), 2U);
  const Value f5 = wb.sheet(0).resolve_cell_value(4U, 5U);
  ASSERT_TRUE(f5.is_text());
  EXPECT_EQ(f5.as_text(), "big");
}

TEST(XlsbFidelity, F6SequenceSpillsAcrossThreeRows) {
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

TEST(XlsbFidelity, H1SumArrayLiteralIs10) {
  Workbook wb = LoadAndRecalcFixture();
  ASSERT_EQ(wb.sheet_count(), 2U);
  const Value h1 = wb.sheet(0).resolve_cell_value(0U, 7U);
  ASSERT_TRUE(h1.is_number());
  EXPECT_DOUBLE_EQ(h1.as_number(), 10.0);
}

TEST(XlsbFidelity, H2SumRangeIs60) {
  Workbook wb = LoadAndRecalcFixture();
  ASSERT_EQ(wb.sheet_count(), 2U);
  const Value h2 = wb.sheet(0).resolve_cell_value(1U, 7U);
  ASSERT_TRUE(h2.is_number());
  EXPECT_DOUBLE_EQ(h2.as_number(), 60.0);
}

TEST(XlsbFidelity, H3VlookupIs30) {
  Workbook wb = LoadAndRecalcFixture();
  ASSERT_EQ(wb.sheet_count(), 2U);
  const Value h3 = wb.sheet(0).resolve_cell_value(2U, 7U);
  ASSERT_TRUE(h3.is_number());
  EXPECT_DOUBLE_EQ(h3.as_number(), 30.0);
}

TEST(XlsbFidelity, H4CrossSheetReferenceIs14) {
  Workbook wb = LoadAndRecalcFixture();
  ASSERT_EQ(wb.sheet_count(), 2U);
  const Value h4 = wb.sheet(0).resolve_cell_value(3U, 7U);
  ASSERT_TRUE(h4.is_number());
  EXPECT_DOUBLE_EQ(h4.as_number(), 14.0);
}

TEST(XlsbFidelity, H5ThreeDReferenceSumIs10) {
  Workbook wb = LoadAndRecalcFixture();
  ASSERT_EQ(wb.sheet_count(), 2U);
  const Value h5 = wb.sheet(0).resolve_cell_value(4U, 7U);
  ASSERT_TRUE(h5.is_number());
  EXPECT_DOUBLE_EQ(h5.as_number(), 10.0);
}

TEST(XlsbFidelity, H6DefinedNameRateAppliesIs1) {
  Workbook wb = LoadAndRecalcFixture();
  ASSERT_EQ(wb.sheet_count(), 2U);
  const Value h6 = wb.sheet(0).resolve_cell_value(5U, 7U);
  ASSERT_TRUE(h6.is_number());
  EXPECT_DOUBLE_EQ(h6.as_number(), 1.0);
}

TEST(XlsbFidelity, J1DivisionByZeroIsDiv0Error) {
  Workbook wb = LoadAndRecalcFixture();
  ASSERT_EQ(wb.sheet_count(), 2U);
  const Value j1 = wb.sheet(0).resolve_cell_value(0U, 9U);
  ASSERT_TRUE(j1.is_error());
  EXPECT_EQ(j1.as_error(), ErrorCode::Div0);
}

TEST(XlsbFidelity, J2NaFunctionIsNaError) {
  Workbook wb = LoadAndRecalcFixture();
  ASSERT_EQ(wb.sheet_count(), 2U);
  const Value j2 = wb.sheet(0).resolve_cell_value(1U, 9U);
  ASSERT_TRUE(j2.is_error());
  EXPECT_EQ(j2.as_error(), ErrorCode::NA);
}

// ---------------------------------------------------------------------------
// (d) The `Rate` defined name is present (decoded from `BrtName`) and
// actually drives evaluation (already exercised indirectly via H6
// above; this pins the metadata directly).
// ---------------------------------------------------------------------------

TEST(XlsbFidelity, DefinedNameRateIsPresentAndUsedByH6) {
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
  const Value h6 = wb.sheet(0).resolve_cell_value(5U, 7U);  // H6 == B1*Rate == 10*0.1
  ASSERT_TRUE(h6.is_number());
  EXPECT_DOUBLE_EQ(h6.as_number(), 1.0);
}

// ---------------------------------------------------------------------------
// (e) Styles: D1/D2/D5 number-format strings (BrtFmt custom formats +
// the BrtXF -> numFmtId -> builtin_num_fmt fallback), D3 bold font
// index, D4 fill index (font/fill/border content itself is not
// decoded — see `io/xlsb/styles_reader.h` — only that the index is
// valid and distinct from the default).
// ---------------------------------------------------------------------------

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

TEST(XlsbFidelity, StylesSurviveLoad) {
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
  EXPECT_GT(styles.cell_xfs[d3->xf_index].font_index, 0U);

  const Cell* d4 = data.cell_at(3U, 3U);
  ASSERT_NE(d4, nullptr);
  ASSERT_LT(d4->xf_index, styles.cell_xfs.size());
  EXPECT_GT(styles.cell_xfs[d4->xf_index].fill_index, 0U);

  const Cell* d5 = data.cell_at(4U, 3U);
  ASSERT_NE(d5, nullptr);
  ASSERT_LT(d5->xf_index, styles.cell_xfs.size());
  EXPECT_EQ(NumFmtStringFor(styles, styles.cell_xfs[d5->xf_index].num_fmt_id), "0.0%");
}

TEST(XlsbFidelity, StylesBinSurvivesAsPassthroughPart) {
  Workbook wb = LoadFixture();
  bool found = false;
  for (const io::PassthroughPart& part : wb.passthrough_parts()) {
    if (part.path == "xl/styles.bin") {
      found = true;
      EXPECT_GT(part.bytes.size(), 0U);
    }
  }
  EXPECT_TRUE(found) << "xl/styles.bin not preserved as a passthrough part";
}

// ---------------------------------------------------------------------------
// (f) A save_as(Xlsb) -> reload cycle must not disturb formula text or
// recalculated values for every Ptg form this bundle closes (future-
// function dispatch, LET, PtgArray literals, plain ranges, and defined-
// name references). `H4` / `H5` are excluded — see the class comment
// above `F1XlookupRecalculatesTo20`.
// ---------------------------------------------------------------------------

TEST(XlsbFidelity, SaveReloadPreservesFormulaTextAndValues) {
  Workbook wb = LoadAndRecalcFixture();
  ASSERT_EQ(wb.sheet_count(), 2U);

  auto saved_or = wb.save_as(io::WorkbookFormat::Xlsb);
  ASSERT_TRUE(static_cast<bool>(saved_or)) << "save_as(Xlsb) failed: " << saved_or.error().message;

  const std::vector<std::uint8_t>& saved = saved_or.value();
  EXPECT_EQ(io::detect_workbook_format(SpanOf(saved)), io::WorkbookFormat::Xlsb);

  auto reloaded_or = io::xlsb::read_xlsb(SpanOf(saved));
  ASSERT_TRUE(static_cast<bool>(reloaded_or)) << "reload after save_as failed: " << reloaded_or.error().message;
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

  const Cell* f2 = data.cell_at(1U, 5U);
  ASSERT_NE(f2, nullptr);
  EXPECT_EQ(f2->formula_text, "=LET(x,2,x*3)");
  const Value f2_value = data.resolve_cell_value(1U, 5U);
  ASSERT_TRUE(f2_value.is_number());
  EXPECT_DOUBLE_EQ(f2_value.as_number(), 6.0);

  const Value h1_value = data.resolve_cell_value(0U, 7U);
  ASSERT_TRUE(h1_value.is_number());
  EXPECT_DOUBLE_EQ(h1_value.as_number(), 10.0);

  const Value h2_value = data.resolve_cell_value(1U, 7U);
  ASSERT_TRUE(h2_value.is_number());
  EXPECT_DOUBLE_EQ(h2_value.as_number(), 60.0);

  const Cell* h4 = data.cell_at(3U, 7U);
  ASSERT_NE(h4, nullptr);
  EXPECT_EQ(h4->formula_text, "='S2'!A1*2");
  const Value h4_value = data.resolve_cell_value(3U, 7U);
  ASSERT_TRUE(h4_value.is_number());
  EXPECT_DOUBLE_EQ(h4_value.as_number(), 14.0);

  const Cell* h5 = data.cell_at(4U, 7U);
  ASSERT_NE(h5, nullptr);
  EXPECT_EQ(h5->formula_text, "=SUM('Data:S2'!B1)");
  const Value h5_value = data.resolve_cell_value(4U, 7U);
  ASSERT_TRUE(h5_value.is_number());
  EXPECT_DOUBLE_EQ(h5_value.as_number(), 10.0);

  const Value h6_value = data.resolve_cell_value(5U, 7U);
  ASSERT_TRUE(h6_value.is_number());
  EXPECT_DOUBLE_EQ(h6_value.as_number(), 1.0);

  const Value j1_value = data.resolve_cell_value(0U, 9U);
  ASSERT_TRUE(j1_value.is_error());
  EXPECT_EQ(j1_value.as_error(), ErrorCode::Div0);

  bool found_rate = false;
  for (const io::DefinedName& dn : reloaded.defined_names()) {
    if (dn.name == "Rate") {
      found_rate = true;
    }
  }
  EXPECT_TRUE(found_rate) << "defined name 'Rate' lost across save_as -> reload";
}

}  // namespace
}  // namespace formulon
