//
// Cross-format read symmetry for the binary XLSB path.
//
// `tests/unit/io/xlsb_fidelity_test.cpp` (bind-a) already checks the `.xlsb`
// reader against literal expected values and covers the `write_xlsb ->
// read_xlsb` round-trip for formula text and cell values. This file adds the
// complementary angle that no existing test covers: the SAME real workbook,
// authored once by Excel and exported to both `.xlsb` and `.xlsx`, must yield
// an equivalent in-memory model regardless of which reader parsed it. A
// format-specific reader divergence (a record the XLSB reader drops but the
// OOXML reader keeps, or vice versa) surfaces as a model mismatch here even
// when each reader passes its own literal-expectation fidelity suite.
//
// XLSB is a binary record stream, so the pugixml attribute-set helpers do not
// apply; the comparison is at the `Workbook` model level.

#include <cmath>
#include <cstdint>
#include <cstring>
#include <iterator>
#include <limits>
#include <string>
#include <vector>

#include "cell.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/defined_names.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/styles_reader.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "sheet.h"
#include "support/roundtrip_symmetry.h"
#include "value.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

std::string FixturePath(const char* name) {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/" + name;
}

// Reads the shared fixture from both formats. Returns false (with a gtest
// failure) if either read fails; otherwise fills the two out-workbooks.
::testing::AssertionResult LoadBothFormats(Workbook* xlsb_out, Workbook* xlsx_out) {
  const std::vector<std::uint8_t> xlsb_bytes = test::read_file_bytes(FixturePath("xlsb_fidelity_base.xlsb"));
  const std::vector<std::uint8_t> xlsx_bytes = test::read_file_bytes(FixturePath("xlsb_fidelity_base.xlsx"));
  if (xlsb_bytes.empty() || xlsx_bytes.empty()) {
    return ::testing::AssertionFailure() << "fixture bytes empty";
  }
  auto xb = io::xlsb::read_xlsb(test::span_of(xlsb_bytes));
  if (!xb) {
    return ::testing::AssertionFailure() << "read_xlsb failed: " << xb.error().message;
  }
  auto xx = io::read_ooxml(test::span_of(xlsx_bytes));
  if (!xx) {
    return ::testing::AssertionFailure() << "read_ooxml failed: " << xx.error().message;
  }
  *xlsb_out = std::move(xb.value().workbook);
  *xlsx_out = std::move(xx.value().workbook);
  return ::testing::AssertionSuccess();
}

TEST(XlsbCrossFormatSymmetry, SheetStructureMatches) {
  Workbook xlsb = Workbook::create_empty();
  Workbook xlsx = Workbook::create_empty();
  ASSERT_TRUE(LoadBothFormats(&xlsb, &xlsx));
  ASSERT_EQ(xlsb.sheet_count(), xlsx.sheet_count());
  for (std::size_t i = 0; i < xlsb.sheet_count(); ++i) {
    EXPECT_EQ(xlsb.sheet(i).name(), xlsx.sheet(i).name()) << "sheet index " << i;
  }
}

TEST(XlsbCrossFormatSymmetry, DataSheetValuesMatch) {
  Workbook xlsb = Workbook::create_empty();
  Workbook xlsx = Workbook::create_empty();
  ASSERT_TRUE(LoadBothFormats(&xlsb, &xlsx));
  const Sheet& sb = xlsb.sheet(0);
  const Sheet& sx = xlsx.sheet(0);
  // A1:A3 are text keys; B1:B3 are the numeric column.
  for (std::uint32_t r = 0; r < 3; ++r) {
    const Cell* ab = sb.cell_at(r, 0);
    const Cell* ax = sx.cell_at(r, 0);
    ASSERT_NE(ab, nullptr);
    ASSERT_NE(ax, nullptr);
    ASSERT_TRUE(ab->cached_value.is_text());
    ASSERT_TRUE(ax->cached_value.is_text());
    EXPECT_EQ(ab->cached_value.as_text(), ax->cached_value.as_text()) << "A" << (r + 1);

    const Cell* bb = sb.cell_at(r, 1);
    const Cell* bx = sx.cell_at(r, 1);
    ASSERT_NE(bb, nullptr);
    ASSERT_NE(bx, nullptr);
    ASSERT_TRUE(bb->cached_value.is_number());
    ASSERT_TRUE(bx->cached_value.is_number());
    EXPECT_DOUBLE_EQ(bb->cached_value.as_number(), bx->cached_value.as_number()) << "B" << (r + 1);
  }
}

TEST(XlsbCrossFormatSymmetry, FormulaTextMatches) {
  Workbook xlsb = Workbook::create_empty();
  Workbook xlsx = Workbook::create_empty();
  ASSERT_TRUE(LoadBothFormats(&xlsb, &xlsx));
  // F1 holds an XLOOKUP; both readers must restore identical formula text
  // (no `_xlfn.` prefix drift between the binary and OOXML paths).
  const Cell* fb = xlsb.sheet(0).cell_at(0, 5);
  const Cell* fx = xlsx.sheet(0).cell_at(0, 5);
  ASSERT_NE(fb, nullptr);
  ASSERT_NE(fx, nullptr);
  EXPECT_EQ(fb->formula_text, fx->formula_text);
  EXPECT_FALSE(fb->formula_text.empty());
}

TEST(XlsbCrossFormatSymmetry, StyleIndexMatches) {
  Workbook xlsb = Workbook::create_empty();
  Workbook xlsx = Workbook::create_empty();
  ASSERT_TRUE(LoadBothFormats(&xlsb, &xlsx));
  // D3 is a bold-red styled cell. Both readers must resolve it to the same
  // style (xf) index -- a cross-format check on the styles.bin vs styles.xml
  // parse producing equivalent style tables.
  const Cell* db = xlsb.sheet(0).cell_at(2, 3);
  const Cell* dx = xlsx.sheet(0).cell_at(2, 3);
  ASSERT_NE(db, nullptr);
  ASSERT_NE(dx, nullptr);
  EXPECT_EQ(db->xf_index, dx->xf_index);
  EXPECT_NE(db->xf_index, 0U) << "D3 should carry a non-default style";
}

TEST(XlsbCrossFormatSymmetry, DefinedNamesMatch) {
  Workbook xlsb = Workbook::create_empty();
  Workbook xlsx = Workbook::create_empty();
  ASSERT_TRUE(LoadBothFormats(&xlsb, &xlsx));
  const std::vector<io::DefinedName>& sb = xlsb.defined_names();
  const std::vector<io::DefinedName>& sx = xlsx.defined_names();
  ASSERT_EQ(sb.size(), sx.size());
  // Field-level, not just count: the XLSB reader must fill the same
  // `io::DefinedName` field set the OOXML reader does (name, formula,
  // scope, hidden, comment) for the same source workbook, not merely
  // produce the same number of entries.
  for (std::size_t i = 0; i < sb.size(); ++i) {
    EXPECT_EQ(sb[i].name, sx[i].name) << "index " << i;
    EXPECT_EQ(sb[i].formula, sx[i].formula) << "index " << i;
    EXPECT_EQ(sb[i].local_sheet_id, sx[i].local_sheet_id) << "index " << i;
    EXPECT_EQ(sb[i].hidden, sx[i].hidden) << "index " << i;
    EXPECT_EQ(sb[i].comment, sx[i].comment) << "index " << i;
  }
}

// Resolves the numFmtId for cell (row, col) through `wb`'s style table, or
// SIZE_MAX-style 0xFFFFFFFF when the cell's xf index dangles past the table
// (which is exactly the failure mode a bare index-equality check would miss).
std::uint32_t ResolvedNumFmtId(const Workbook& wb, std::uint32_t row, std::uint32_t col) {
  const Cell* c = wb.sheet(0).cell_at(row, col);
  if (c == nullptr) {
    return 0xFFFFFFFFU;
  }
  const io::StylesTable& st = wb.styles();
  if (c->xf_index >= st.cell_xfs.size()) {
    return 0xFFFFFFFFU;  // dangling index -> style table did not round-trip
  }
  return st.cell_xfs[c->xf_index].num_fmt_id;
}

// Regression for the writer defect the cross-format check surfaced. Two halves,
// both required for a styled cell to survive `write_xlsb -> read_xlsb`:
//   1. the cell header must carry the 24-bit iStyleRef (was hardcoded to 0);
//   2. the workbook must declare the styles relationship so the reader can find
//      the (passthrough) styles.bin -- otherwise the index dangles against an
//      empty table.
// Asserting the *resolved* numFmtId (not just index equality) exercises both:
// D1 = yyyy/mm/dd (custom 179), D2 = #,##0.00 (built-in 4), D5 = 0.0% (custom
// 180). D3 keeps its non-default font xf.
TEST(XlsbWriteReadSymmetry, CellStyleAndNumberFormatSurviveRoundTrip) {
  const std::vector<std::uint8_t> bytes = test::read_file_bytes(FixturePath("xlsb_fidelity_base.xlsb"));
  ASSERT_FALSE(bytes.empty());
  auto loaded = io::xlsb::read_xlsb(test::span_of(bytes));
  ASSERT_TRUE(static_cast<bool>(loaded)) << "read_xlsb failed: " << loaded.error().message;
  const Workbook& before = loaded.value().workbook;

  // Baseline: the fixture resolves the expected number formats.
  ASSERT_EQ(ResolvedNumFmtId(before, 0, 3), 179U);  // D1 yyyy/mm/dd
  ASSERT_EQ(ResolvedNumFmtId(before, 1, 3), 4U);    // D2 #,##0.00
  ASSERT_EQ(ResolvedNumFmtId(before, 4, 3), 180U);  // D5 0.0%
  const Cell* d3_before = before.sheet(0).cell_at(2, 3);
  ASSERT_NE(d3_before, nullptr);
  ASSERT_NE(d3_before->xf_index, 0U) << "fixture D3 should carry a non-default style";

  auto saved = io::xlsb::write_xlsb(before);
  ASSERT_TRUE(static_cast<bool>(saved)) << "write_xlsb failed: " << saved.error().message;
  auto reloaded = io::xlsb::read_xlsb(test::span_of(saved.value()));
  ASSERT_TRUE(static_cast<bool>(reloaded)) << "reload failed: " << reloaded.error().message;
  const Workbook& after = reloaded.value().workbook;

  // Number formats still resolve after the round-trip (index + style table).
  EXPECT_EQ(ResolvedNumFmtId(after, 0, 3), 179U);
  EXPECT_EQ(ResolvedNumFmtId(after, 1, 3), 4U);
  EXPECT_EQ(ResolvedNumFmtId(after, 4, 3), 180U);
  // D3's font style index is preserved.
  const Cell* d3_after = after.sheet(0).cell_at(2, 3);
  ASSERT_NE(d3_after, nullptr);
  EXPECT_EQ(d3_after->xf_index, d3_before->xf_index);
}

// A workbook-scoped name and a sheet-local name may spell the same text
// (`Workbook::set_defined_name_scoped` admits the pair), and each needs
// its own `BrtName` record: the `ilbl` a cell's `PtgName` carries is a
// 1-based ordinal into the emitted record sequence, so collapsing the
// pair into one record shifts every later name's ordinal and silently
// re-points the referencing formulas at a different name.
TEST(XlsbWriteReadSymmetry, NamesSharingTextAcrossScopesKeepTheirOwnOrdinals) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name_scoped("Foo", "Sheet1!$A$1", -1)));
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name_scoped("Foo", "Sheet1!$B$1", 0)));
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name_scoped("Bar", "Sheet1!$C$1", -1)));
  // Route the edits through the workbook so the dep graph sees them.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(10.0))));  // A1
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 1U, Value::number(20.0))));  // B1
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 2U, Value::number(30.0))));  // C1
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 3U, "=Bar+1")));           // D1
  auto before_recalc = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(before_recalc)) << before_recalc.error().message;
  const Value d1_before = wb.sheet(0).resolve_cell_value(0U, 3U);
  ASSERT_TRUE(d1_before.is_number()) << d1_before.debug_to_string();
  ASSERT_EQ(d1_before.as_number(), 31.0);

  auto saved = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << "write_xlsb failed: " << saved.error().message;
  auto reloaded = io::xlsb::read_xlsb(test::span_of(saved.value()));
  ASSERT_TRUE(static_cast<bool>(reloaded)) << "read_xlsb failed: " << reloaded.error().message;
  Workbook after = std::move(reloaded.value().workbook);

  // All three names survive, in declaration order, with their scopes.
  const std::vector<io::DefinedName>& names = after.defined_names();
  ASSERT_EQ(names.size(), 3U);
  EXPECT_EQ(names[0].name, "Foo");
  EXPECT_EQ(names[0].formula, "Sheet1!$A$1");
  EXPECT_EQ(names[0].local_sheet_id, -1);
  EXPECT_EQ(names[1].name, "Foo");
  EXPECT_EQ(names[1].formula, "Sheet1!$B$1");
  EXPECT_EQ(names[1].local_sheet_id, 0);
  EXPECT_EQ(names[2].name, "Bar");
  EXPECT_EQ(names[2].formula, "Sheet1!$C$1");
  EXPECT_EQ(names[2].local_sheet_id, -1);

  // The referencing formula still names `Bar`, and recalculating it on
  // the reloaded workbook lands on C1 + 1 rather than on whatever name
  // ordinal 2 would otherwise have become.
  const Cell* d1 = after.sheet(0).cell_at(0U, 3U);
  ASSERT_NE(d1, nullptr);
  EXPECT_EQ(d1->formula_text, "=Bar+1");
  auto after_recalc = after.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(after_recalc)) << after_recalc.error().message;
  const Value d1_after = after.sheet(0).resolve_cell_value(0U, 3U);
  ASSERT_TRUE(d1_after.is_number());
  EXPECT_EQ(d1_after.as_number(), d1_before.as_number());
}

// Wire-level view of one `BrtName` record: the name text plus the scope
// it declares (`itab`, `-1` for workbook scope).
struct WireName {
  std::string name;
  std::int32_t itab = -1;
};

// Decodes `xl/workbook.bin`'s `BrtName` records in emission order, so
// entry `i` is what a `PtgName` with `ilbl == i + 1` reaches.
std::vector<WireName> ReadWireNames(const std::vector<std::uint8_t>& workbook_bin) {
  std::vector<WireName> out;
  io::ByteSpan cursor = test::span_of(workbook_bin);
  while (cursor.size > 0U) {
    auto rec_or = io::xlsb::read_record(cursor);
    if (!rec_or) {
      return out;
    }
    if (rec_or.value().type != static_cast<std::uint16_t>(io::xlsb::XlsbRecordType::BrtName)) {
      continue;
    }
    // BrtName: grbit (u32) + chKey (u8) + itab (i32) + the name string.
    io::ByteSpan p = rec_or.value().payload;
    auto grbit = io::xlsb::read_u32(p);
    auto ch_key = io::xlsb::read_u8(p);
    auto itab = io::xlsb::read_u32(p);
    if (!grbit || !ch_key || !itab) {
      return out;
    }
    auto name = io::xlsb::read_xlwidestring(p);
    if (!name) {
      return out;
    }
    out.push_back(WireName{name.value(), static_cast<std::int32_t>(itab.value())});
  }
  return out;
}

// Returns the `ilbl` encoded by the lone `PtgName` token of the formula
// stored at row 0 / `col` of `sheet_bin`, or 0 when the cell is absent
// or its token stream is not a single bare name reference.
std::uint32_t IlblOfNameOnlyFormula(const std::vector<std::uint8_t>& sheet_bin, std::uint32_t col) {
  io::ByteSpan cursor = test::span_of(sheet_bin);
  while (cursor.size > 0U) {
    auto rec_or = io::xlsb::read_record(cursor);
    if (!rec_or) {
      return 0U;
    }
    if (rec_or.value().type != static_cast<std::uint16_t>(io::xlsb::XlsbRecordType::BrtFmlaNum)) {
      continue;
    }
    // BrtFmlaNum: cell header (col u32 + iStyleRef 3B + fPhShow u8),
    // the cached double, grbitFlags (u16), then the CellParsedFormula
    // (cce + rgce + cb + rgcb).
    io::ByteSpan p = rec_or.value().payload;
    auto cell_col = io::xlsb::read_u32(p);
    if (!cell_col || cell_col.value() != col) {
      continue;
    }
    if (p.size < 14U) {
      return 0U;
    }
    p.data += 4U + 8U + 2U;  // iStyleRef + fPhShow, cached value, grbitFlags
    p.size -= 4U + 8U + 2U;
    auto cce = io::xlsb::read_u32(p);
    // `=Foo` lowers to exactly one PtgName: opcode 0x23 + a u32 ilbl.
    if (!cce || cce.value() != 5U || p.size < 5U || p.data[0] != 0x23U) {
      return 0U;
    }
    p.data += 1U;
    p.size -= 1U;
    auto ilbl = io::xlsb::read_u32(p);
    return ilbl ? ilbl.value() : 0U;
  }
  return 0U;
}

// Builds the workbook both scope-resolution cases share: a
// workbook-scoped `Foo` (Sheet1!A1 = 10) and a Sheet1-local `Foo`
// (Sheet1!B1 = 20), with `=Foo` on both sheets. `local_first` flips the
// declaration order of the two names, which is the tie-break a
// text-keyed name table falls back on.
void RunScopeResolutionCase(bool local_first) {
  SCOPED_TRACE(local_first ? "sheet-local Foo declared first" : "workbook-scoped Foo declared first");
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  wb.add_sheet("Sheet2");
  if (local_first) {
    ASSERT_TRUE(static_cast<bool>(wb.set_defined_name_scoped("Foo", "Sheet1!$B$1", 0)));
    ASSERT_TRUE(static_cast<bool>(wb.set_defined_name_scoped("Foo", "Sheet1!$A$1", -1)));
  } else {
    ASSERT_TRUE(static_cast<bool>(wb.set_defined_name_scoped("Foo", "Sheet1!$A$1", -1)));
    ASSERT_TRUE(static_cast<bool>(wb.set_defined_name_scoped("Foo", "Sheet1!$B$1", 0)));
  }
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(10.0))));  // Sheet1!A1
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 1U, Value::number(20.0))));  // Sheet1!B1
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 3U, "=Foo")));             // Sheet1!D1
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(1U, 0U, 3U, "=Foo")));             // Sheet2!D1
  auto recalc_or = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(recalc_or)) << recalc_or.error().message;
  // The engine's own resolution is the reference the encoding has to
  // agree with: Sheet1 sees the local `Foo` (B1), Sheet2 the global one.
  ASSERT_EQ(wb.sheet(0).resolve_cell_value(0U, 3U).as_number(), 20.0);
  ASSERT_EQ(wb.sheet(1).resolve_cell_value(0U, 3U).as_number(), 10.0);

  auto saved = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << "write_xlsb failed: " << saved.error().message;
  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(test::span_of(saved.value()))));
  auto workbook_bin = zip.read_entry("xl/workbook.bin");
  auto sheet1_bin = zip.read_entry("xl/worksheets/sheet1.bin");
  auto sheet2_bin = zip.read_entry("xl/worksheets/sheet2.bin");
  ASSERT_TRUE(static_cast<bool>(workbook_bin));
  ASSERT_TRUE(static_cast<bool>(sheet1_bin));
  ASSERT_TRUE(static_cast<bool>(sheet2_bin));

  const std::vector<WireName> wire_names = ReadWireNames(workbook_bin.value());
  ASSERT_EQ(wire_names.size(), 2U);

  const std::uint32_t sheet1_ilbl = IlblOfNameOnlyFormula(sheet1_bin.value(), 3U);
  const std::uint32_t sheet2_ilbl = IlblOfNameOnlyFormula(sheet2_bin.value(), 3U);
  ASSERT_GE(sheet1_ilbl, 1U);
  ASSERT_LE(sheet1_ilbl, wire_names.size());
  ASSERT_GE(sheet2_ilbl, 1U);
  ASSERT_LE(sheet2_ilbl, wire_names.size());

  // Sheet1's reference must land on the record scoped to Sheet1.
  EXPECT_EQ(wire_names[sheet1_ilbl - 1U].name, "Foo");
  EXPECT_EQ(wire_names[sheet1_ilbl - 1U].itab, 0) << "Sheet1 must resolve the sheet-local Foo";
  // Sheet2 has no local override, so it must land on the workbook one.
  EXPECT_EQ(wire_names[sheet2_ilbl - 1U].name, "Foo");
  EXPECT_EQ(wire_names[sheet2_ilbl - 1U].itab, -1) << "Sheet2 must resolve the workbook-scoped Foo";

  // The values the round trip produces are unchanged either way (the
  // reader re-resolves by name text), so they are a guard against the
  // scope fix disturbing them, not the detector for it.
  auto reloaded = io::xlsb::read_xlsb(test::span_of(saved.value()));
  ASSERT_TRUE(static_cast<bool>(reloaded)) << "read_xlsb failed: " << reloaded.error().message;
  Workbook after = std::move(reloaded.value().workbook);
  auto after_recalc = after.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(after_recalc)) << after_recalc.error().message;
  EXPECT_EQ(after.sheet(0).resolve_cell_value(0U, 3U).as_number(), 20.0);
  EXPECT_EQ(after.sheet(1).resolve_cell_value(0U, 3U).as_number(), 10.0);
}

// A `PtgName` token is a bare ordinal into the `BrtName` table -- it
// carries no scope -- so the writer, not the consumer, decides which of
// two same-named records a formula reaches. Excel resolves an
// unqualified name from the sheet the formula sits on: a sheet-local
// name shadows the workbook-scoped one there, and the workbook-scoped
// one is reached only from sheets that have no local override. Encoding
// one workbook-wide ordinal per name text silently re-points the
// reference on whichever side loses the tie-break, so both declaration
// orders are exercised: each one makes a different clause of the rule
// the one that fails. No count-based check sees this -- both records are
// emitted either way -- so the assertion has to be which record the
// ordinal lands on.
TEST(XlsbWriteReadSymmetry, UnqualifiedNameEncodesTheOrdinalOfTheScopeExcelResolves) {
  RunScopeResolutionCase(/*local_first=*/false);
}

TEST(XlsbWriteReadSymmetry, UnqualifiedNameScopeIgnoresDeclarationOrder) {
  RunScopeResolutionCase(/*local_first=*/true);
}

// Reinterprets `v`'s object representation so two doubles can be
// compared for bit equality rather than numeric equality.
std::uint64_t BitsOf(double v) {
  std::uint64_t bits = 0;
  std::memcpy(&bits, &v, sizeof(v));
  return bits;
}

// The whole point of the predicate: a `true` answer is a promise that
// `BrtCellRk` is lossless for that value. Sweep a deterministic spread
// of finite doubles -- including the currency-shaped band where the
// `x100` form is tempting and the multiplication rounds -- and hold the
// implication for every one of them.
TEST(XlsbRkEncoding, PredicateOnlyAcceptsValuesWhoseEncodingDecodesToTheSameBits) {
  std::uint64_t state = 0x9E3779B97F4A7C15ULL;  // any fixed seed; the sweep must be reproducible
  auto next = [&state]() {
    state = state * 6364136223846793005ULL + 1442695040888963407ULL;
    return state;
  };
  int accepted = 0;
  for (int i = 0; i < 20000; ++i) {
    // Currency-shaped magnitudes (cents over a 30-bit range), their two
    // immediate neighbours -- where multiplying by 100 rounds back onto
    // the exact cent count and the x100 form therefore looks applicable
    // while decoding to a different double -- and a spread of arbitrary
    // bit patterns.
    const double cents = static_cast<double>(static_cast<std::int64_t>(next() % 1000000000ULL) - 500000000);
    const double base = cents / 100.0;
    const double candidates[] = {base, std::nextafter(base, std::numeric_limits<double>::infinity()),
                                 std::nextafter(base, -std::numeric_limits<double>::infinity()), cents / 3.0, cents};
    for (const double v : candidates) {
      if (!io::xlsb::rk_round_trips_value(v)) {
        continue;
      }
      ++accepted;
      std::vector<std::uint8_t> bytes;
      io::xlsb::emit_rk_number(bytes, v);
      ASSERT_EQ(bytes.size(), 4U);
      const std::uint32_t rk = static_cast<std::uint32_t>(bytes[0]) | (static_cast<std::uint32_t>(bytes[1]) << 8) |
                               (static_cast<std::uint32_t>(bytes[2]) << 16) |
                               (static_cast<std::uint32_t>(bytes[3]) << 24);
      EXPECT_EQ(BitsOf(io::xlsb::decode_rk_number(rk)), BitsOf(v)) << "value=" << v;
    }
  }
  EXPECT_GT(accepted, 0) << "sweep never exercised the accepting branch";
}

// Values whose `x * 100` product rounds to an integer are not RK-x100
// encodable even though the product passes an integrality test: the
// decode divides by 100 again and lands on a neighbouring double. The
// cell writer must route them to `BrtCellReal`, so an `.xlsx -> .xlsb
// -> .xlsx` conversion has to preserve the bit pattern exactly.
TEST(XlsbWriteReadSymmetry, XlsxToXlsbToXlsxPreservesNumericBitPatterns) {
  const double kValues[] = {3611469.5700000003, -4123191.7399999998};
  for (const double v : kValues) {
    EXPECT_FALSE(io::xlsb::rk_round_trips_value(v)) << "value=" << v;
  }

  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  for (std::size_t i = 0; i < std::size(kValues); ++i) {
    wb.sheet(0).set_cell_value(static_cast<std::uint32_t>(i), 0U, Value::number(kValues[i]));
  }

  auto xlsx_in = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(xlsx_in)) << "write_ooxml failed: " << xlsx_in.error().message;
  auto from_xlsx = io::read_ooxml(test::span_of(xlsx_in.value()));
  ASSERT_TRUE(static_cast<bool>(from_xlsx)) << "read_ooxml failed: " << from_xlsx.error().message;

  auto xlsb = io::xlsb::write_xlsb(from_xlsx.value().workbook);
  ASSERT_TRUE(static_cast<bool>(xlsb)) << "write_xlsb failed: " << xlsb.error().message;
  auto from_xlsb = io::xlsb::read_xlsb(test::span_of(xlsb.value()));
  ASSERT_TRUE(static_cast<bool>(from_xlsb)) << "read_xlsb failed: " << from_xlsb.error().message;

  auto xlsx_out = io::write_ooxml(from_xlsb.value().workbook);
  ASSERT_TRUE(static_cast<bool>(xlsx_out)) << "write_ooxml failed: " << xlsx_out.error().message;
  auto final_wb = io::read_ooxml(test::span_of(xlsx_out.value()));
  ASSERT_TRUE(static_cast<bool>(final_wb)) << "read_ooxml failed: " << final_wb.error().message;

  const Sheet& s = final_wb.value().workbook.sheet(0);
  for (std::size_t i = 0; i < std::size(kValues); ++i) {
    const Cell* c = s.cell_at(static_cast<std::uint32_t>(i), 0U);
    ASSERT_NE(c, nullptr) << "row " << i;
    ASSERT_TRUE(c->cached_value.is_number()) << "row " << i;
    EXPECT_EQ(BitsOf(c->cached_value.as_number()), BitsOf(kValues[i])) << "row " << i;
  }
}

}  // namespace
}  // namespace formulon
