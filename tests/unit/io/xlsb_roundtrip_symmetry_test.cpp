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

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <cstring>
#include <iterator>
#include <limits>
#include <string>
#include <utility>
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

// An equal xf index says the two readers agree on which slot a cell points
// at, not on what that slot contains: styles.bin and styles.xml are decoded
// by separate code paths into the same `StylesTable`, so a slot can carry a
// different font, fill or number format on each side while every index still
// matches. The comparisons below are on the record contents, which is where
// a binary-decode defect actually lands.

/// Compares the two colour selectors a `ColorSpec` can carry.
///
/// The selector kind is the part a format conversion is most likely to lose:
/// XLSB packs automatic / indexed / rgb / theme into one flags nibble, while
/// OOXML spells each as its own `<color>` attribute. `color_argb` is only a
/// compatibility fallback for the non-RGB selectors, so it is compared
/// through the spec rather than on its own.
void ExpectColorSpecEqual(const io::ColorSpec& xlsb, const io::ColorSpec& xlsx) {
  ASSERT_EQ(static_cast<int>(xlsb.kind), static_cast<int>(xlsx.kind));
  switch (xlsb.kind) {
    case io::ColorSpec::Kind::kRgb:
      EXPECT_EQ(xlsb.rgb, xlsx.rgb);
      break;
    case io::ColorSpec::Kind::kTheme:
      EXPECT_EQ(xlsb.theme, xlsx.theme);
      EXPECT_NEAR(xlsb.tint, xlsx.tint, 1e-9);
      break;
    case io::ColorSpec::Kind::kIndexed:
      EXPECT_EQ(xlsb.indexed, xlsx.indexed);
      break;
    case io::ColorSpec::Kind::kNone:
    case io::ColorSpec::Kind::kAuto:
      break;
  }
}

TEST(XlsbCrossFormatSymmetry, FontRecordContentsMatch) {
  Workbook xlsb = Workbook::create_empty();
  Workbook xlsx = Workbook::create_empty();
  ASSERT_TRUE(LoadBothFormats(&xlsb, &xlsx));
  const std::vector<io::FontRecord>& fb = xlsb.styles().fonts;
  const std::vector<io::FontRecord>& fx = xlsx.styles().fonts;
  ASSERT_EQ(fb.size(), fx.size());

  // The fixture has to carry a font that differs from the default record in
  // more than one attribute, or every field comparison below would hold on a
  // table of blanks.
  bool saw_non_default = false;
  for (const io::FontRecord& font : fx) {
    if (font.bold && font.color.kind == io::ColorSpec::Kind::kRgb) {
      saw_non_default = true;
    }
  }
  ASSERT_TRUE(saw_non_default) << "fixture carries no bold, explicitly coloured font";

  for (std::size_t i = 0; i < fb.size(); ++i) {
    SCOPED_TRACE("font index " + std::to_string(i));
    EXPECT_EQ(fb[i].name, fx[i].name);
    EXPECT_DOUBLE_EQ(fb[i].size, fx[i].size);
    EXPECT_EQ(fb[i].bold, fx[i].bold);
    EXPECT_EQ(fb[i].has_bold, fx[i].has_bold);
    EXPECT_EQ(fb[i].italic, fx[i].italic);
    EXPECT_EQ(fb[i].has_italic, fx[i].has_italic);
    EXPECT_EQ(fb[i].strike, fx[i].strike);
    EXPECT_EQ(fb[i].has_strike, fx[i].has_strike);
    EXPECT_EQ(fb[i].underline, fx[i].underline);
    EXPECT_EQ(fb[i].vert_align, fx[i].vert_align);
    EXPECT_EQ(fb[i].has_family, fx[i].has_family);
    EXPECT_EQ(fb[i].family, fx[i].family);
    EXPECT_EQ(fb[i].has_charset, fx[i].has_charset);
    EXPECT_EQ(fb[i].charset, fx[i].charset);
    ExpectColorSpecEqual(fb[i].color, fx[i].color);
  }
}

TEST(XlsbCrossFormatSymmetry, FillRecordContentsMatch) {
  Workbook xlsb = Workbook::create_empty();
  Workbook xlsx = Workbook::create_empty();
  ASSERT_TRUE(LoadBothFormats(&xlsb, &xlsx));
  const std::vector<io::FillRecord>& fb = xlsb.styles().fills;
  const std::vector<io::FillRecord>& fx = xlsx.styles().fills;
  ASSERT_EQ(fb.size(), fx.size());

  // A patterned fill with an explicit foreground colour is the case that
  // distinguishes the two decode paths; the two placeholder fills Excel
  // always writes first would not.
  bool saw_coloured_pattern = false;
  for (const io::FillRecord& fill : fx) {
    if (fill.pattern != 0U && fill.fg.kind != io::ColorSpec::Kind::kNone) {
      saw_coloured_pattern = true;
    }
  }
  ASSERT_TRUE(saw_coloured_pattern) << "fixture carries no coloured pattern fill";

  for (std::size_t i = 0; i < fb.size(); ++i) {
    SCOPED_TRACE("fill index " + std::to_string(i));
    EXPECT_EQ(fb[i].pattern, fx[i].pattern);
    ExpectColorSpecEqual(fb[i].fg, fx[i].fg);
    ExpectColorSpecEqual(fb[i].bg, fx[i].bg);
  }
}

TEST(XlsbCrossFormatSymmetry, CellXfRecordContentsMatch) {
  Workbook xlsb = Workbook::create_empty();
  Workbook xlsx = Workbook::create_empty();
  ASSERT_TRUE(LoadBothFormats(&xlsb, &xlsx));
  const io::StylesTable& tb = xlsb.styles();
  const io::StylesTable& tx = xlsx.styles();
  ASSERT_EQ(tb.cell_xfs.size(), tx.cell_xfs.size());

  // An xf table of nothing but copies of the default record would satisfy
  // every field comparison without exercising the decode.
  bool saw_non_default = false;
  for (const io::CellXf& xf : tx.cell_xfs) {
    if (xf.font_index != 0U || xf.fill_index != 0U || xf.num_fmt_id != 0U) {
      saw_non_default = true;
    }
  }
  ASSERT_TRUE(saw_non_default) << "fixture xf table selects only default records";

  // The fixture is authored in ja-JP, where `vertical="center"` is the
  // default Excel applies to every cell, so an xf table that came back
  // bottom-aligned is the shape this comparison exists to reject.
  bool saw_alignment = false;
  bool saw_apply_flag = false;
  for (const io::CellXf& xf : tx.cell_xfs) {
    if (io::HasAlignment(xf)) {
      saw_alignment = true;
    }
    if (xf.apply_number_format || xf.apply_font || xf.apply_fill) {
      saw_apply_flag = true;
    }
  }
  ASSERT_TRUE(saw_alignment) << "fixture xf table carries no alignment";
  ASSERT_TRUE(saw_apply_flag) << "fixture xf table sets no apply flag";

  for (std::size_t i = 0; i < tb.cell_xfs.size(); ++i) {
    SCOPED_TRACE("cellXf index " + std::to_string(i));
    const io::CellXf& b = tb.cell_xfs[i];
    const io::CellXf& x = tx.cell_xfs[i];
    // The selector fields, which are what makes an xf name one font, fill,
    // border and number format rather than another.
    EXPECT_EQ(b.font_index, x.font_index);
    EXPECT_EQ(b.fill_index, x.fill_index);
    EXPECT_EQ(b.border_index, x.border_index);
    EXPECT_EQ(b.num_fmt_id, x.num_fmt_id);
    EXPECT_EQ(b.xf_id, x.xf_id);
    // The `apply*` set, which decides whether the xf's own font / format
    // wins over the named style it inherits from.
    EXPECT_EQ(b.apply_number_format, x.apply_number_format);
    EXPECT_EQ(b.apply_font, x.apply_font);
    EXPECT_EQ(b.apply_fill, x.apply_fill);
    EXPECT_EQ(b.apply_border, x.apply_border);
    EXPECT_EQ(b.apply_alignment, x.apply_alignment);
    EXPECT_EQ(b.apply_protection, x.apply_protection);
    // Alignment and protection are compared on their effective values and
    // on the presence predicates the writer consults, not on the raw
    // `has_*` bits: those record how OOXML spelled a value, and `BrtXF`
    // states every field unconditionally, so an XLSB-sourced xf derives
    // presence from the value differing from its schema default.
    EXPECT_EQ(b.horizontal_align, x.horizontal_align);
    EXPECT_EQ(b.vertical_align, x.vertical_align);
    EXPECT_EQ(b.wrap_text, x.wrap_text);
    EXPECT_EQ(b.justify_last_line, x.justify_last_line);
    EXPECT_EQ(b.shrink_to_fit, x.shrink_to_fit);
    EXPECT_EQ(b.reading_order, x.reading_order);
    EXPECT_EQ(b.text_rotation, x.text_rotation);
    EXPECT_EQ(b.indent, x.indent);
    EXPECT_EQ(b.quote_prefix, x.quote_prefix);
    EXPECT_EQ(b.locked, x.locked);
    EXPECT_EQ(b.hidden, x.hidden);
    EXPECT_EQ(b.has_protection, x.has_protection);
    EXPECT_EQ(io::HasAlignment(b), io::HasAlignment(x));
    EXPECT_EQ(io::HasHorizontalAlign(b), io::HasHorizontalAlign(x));
    EXPECT_EQ(io::HasVerticalAlign(b), io::HasVerticalAlign(x));
    EXPECT_EQ(io::HasWrapText(b), io::HasWrapText(x));
    EXPECT_EQ(io::HasJustifyLastLine(b), io::HasJustifyLastLine(x));
  }
}

/// Returns the `<cellXfs>` element of `package`'s `xl/styles.xml`, or an
/// empty string when the part or the element is missing.
std::string CellXfsBlockOfSavedPackage(const std::vector<std::uint8_t>& package) {
  std::string styles;
  if (!test::extract_part(test::span_of(package), "xl/styles.xml", &styles)) {
    return std::string();
  }
  const std::size_t begin = styles.find("<cellXfs");
  const std::size_t end = styles.find("</cellXfs>");
  if (begin == std::string::npos || end == std::string::npos || end < begin) {
    return std::string();
  }
  return styles.substr(begin, end - begin + std::strlen("</cellXfs>"));
}

/// Counts non-overlapping occurrences of `needle` in `haystack`.
std::size_t CountOccurrences(const std::string& haystack, const std::string& needle) {
  std::size_t count = 0;
  for (std::size_t pos = haystack.find(needle); pos != std::string::npos; pos = haystack.find(needle, pos + 1)) {
    ++count;
  }
  return count;
}

// The in-memory table is only half the claim: an `.xlsx` save serialises the
// model, never the retained `xl/styles.bin` bytes, so what a host or another
// spreadsheet application sees after a conversion is whatever reached
// `xl/styles.xml`. Asserting on the saved part -- and against the same part
// produced from the workbook's own `.xlsx` export, which is the identical
// writer on an identical model -- is what states that a `.xlsb` source loses
// nothing on the way out.
TEST(XlsbToOoxmlStyles, AlignmentAndApplyFlagsReachTheSavedStylesPart) {
  Workbook xlsb = Workbook::create_empty();
  Workbook xlsx = Workbook::create_empty();
  ASSERT_TRUE(LoadBothFormats(&xlsb, &xlsx));

  auto from_xlsb = io::write_ooxml(xlsb);
  ASSERT_TRUE(static_cast<bool>(from_xlsb)) << "write_ooxml failed: " << from_xlsb.error().message;
  auto from_xlsx = io::write_ooxml(xlsx);
  ASSERT_TRUE(static_cast<bool>(from_xlsx)) << "write_ooxml failed: " << from_xlsx.error().message;

  const std::string xlsb_block = CellXfsBlockOfSavedPackage(from_xlsb.value());
  const std::string xlsx_block = CellXfsBlockOfSavedPackage(from_xlsx.value());
  ASSERT_FALSE(xlsb_block.empty()) << "saved package carries no <cellXfs>";

  // The literals first: an equality against an equally empty block would
  // hold without any of this reaching the file.
  EXPECT_EQ(CountOccurrences(xlsb_block, "<alignment vertical=\"center\"/>"), 6U) << xlsb_block;
  EXPECT_EQ(CountOccurrences(xlsb_block, "applyNumberFormat=\"1\""), 3U) << xlsb_block;
  EXPECT_EQ(CountOccurrences(xlsb_block, "applyFont=\"1\""), 1U) << xlsb_block;
  EXPECT_EQ(CountOccurrences(xlsb_block, "applyFill=\"1\""), 1U) << xlsb_block;
  EXPECT_EQ(xlsb_block, xlsx_block);
}

// A custom number format is stored as an id plus an interned string, so an
// equal `num_fmt_id` on both sides still leaves the format code itself
// unchecked: the two readers intern into their own `num_fmt_strings`
// vectors, and only resolving the id back to its string compares what a
// cell would actually render with.
TEST(XlsbCrossFormatSymmetry, CustomNumberFormatCodesMatch) {
  Workbook xlsb = Workbook::create_empty();
  Workbook xlsx = Workbook::create_empty();
  ASSERT_TRUE(LoadBothFormats(&xlsb, &xlsx));
  const auto codes_by_id = [](const io::StylesTable& table) {
    std::vector<std::pair<std::uint16_t, std::string>> out;
    for (const io::NumFmtRecord& rec : table.num_fmts) {
      if (rec.format_string_index >= table.num_fmt_strings.size()) {
        ADD_FAILURE() << "numFmt id " << rec.id << " interns past the string table";
        continue;
      }
      out.emplace_back(rec.id, table.num_fmt_strings[rec.format_string_index]);
    }
    std::sort(out.begin(), out.end());
    return out;
  };
  const auto b = codes_by_id(xlsb.styles());
  const auto x = codes_by_id(xlsx.styles());
  ASSERT_FALSE(x.empty()) << "fixture declares no custom number format";
  EXPECT_EQ(b, x);
}

// `<cols>` and `BrtColInfo` describe the same span from different encodings
// (character widths versus 1/256 of a standard digit), so an entry that
// survives one reader can arrive with a shifted width or a lost outline
// level from the other.
TEST(XlsbCrossFormatSymmetry, ColumnLayoutMatches) {
  Workbook xlsb = Workbook::create_empty();
  Workbook xlsx = Workbook::create_empty();
  ASSERT_TRUE(LoadBothFormats(&xlsb, &xlsx));
  const std::vector<ColumnLayout>& cb = xlsb.sheet(0).layout().columns;
  const std::vector<ColumnLayout>& cx = xlsx.sheet(0).layout().columns;
  ASSERT_EQ(cb.size(), cx.size());
  ASSERT_FALSE(cx.empty()) << "fixture sheet declares no <cols> entry";
  for (std::size_t i = 0; i < cb.size(); ++i) {
    SCOPED_TRACE("column entry " + std::to_string(i));
    EXPECT_EQ(cb[i].first, cx[i].first);
    EXPECT_EQ(cb[i].last, cx[i].last);
    EXPECT_EQ(cb[i].hidden, cx[i].hidden);
    EXPECT_EQ(cb[i].outline_level, cx[i].outline_level);
    EXPECT_EQ(cb[i].has_width, cx[i].has_width);
    // The width survives the 1/256-digit quantisation to within one step of
    // it. `has_style` is deliberately not compared: `BrtColInfo` carries a
    // mandatory `ixfe` with no presence bit, so the XLSB side reports an
    // effective style 0 where OOXML reports none.
    EXPECT_NEAR(cb[i].width, cx[i].width, 1.0 / 256.0);
  }
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

// The Excel-authored fixture leaves two model areas with no case of their
// own: it has no custom row height, hidden row or outline level, and its one
// `<border>` is the empty default. Neither can be compared across the two
// readers on that source, so the property is pinned one step further out --
// the same in-memory workbook through each format's writer and back -- which
// still fails if either binary path drops the field.

/// Saves `wb` as `.xlsb` and reads it back, or fails the test.
::testing::AssertionResult ThroughXlsb(const Workbook& wb, Workbook* out) {
  auto saved = io::xlsb::write_xlsb(wb);
  if (!saved) {
    return ::testing::AssertionFailure() << "write_xlsb failed: " << saved.error().message;
  }
  auto reloaded = io::xlsb::read_xlsb(test::span_of(saved.value()));
  if (!reloaded) {
    return ::testing::AssertionFailure() << "read_xlsb failed: " << reloaded.error().message;
  }
  *out = std::move(reloaded.value().workbook);
  return ::testing::AssertionSuccess();
}

/// Saves `wb` as `.xlsx` and reads it back, or fails the test.
::testing::AssertionResult ThroughXlsx(const Workbook& wb, Workbook* out) {
  auto saved = io::write_ooxml(wb);
  if (!saved) {
    return ::testing::AssertionFailure() << "write_ooxml failed: " << saved.error().message;
  }
  auto reloaded = io::read_ooxml(test::span_of(saved.value()));
  if (!reloaded) {
    return ::testing::AssertionFailure() << "read_ooxml failed: " << reloaded.error().message;
  }
  *out = std::move(reloaded.value().workbook);
  return ::testing::AssertionSuccess();
}

TEST(XlsbWriteReadSymmetry, RowOverridesSurviveBothFormatsAlike) {
  Workbook source = Workbook::create_empty();
  source.add_sheet("Sheet1");
  // A cell keeps the rows from being dropped as empty, and each override
  // engages a different `BrtRowHdr` flag: a custom height (fUnsynced), the
  // hidden bit, and an outline level.
  source.sheet(0).set_cell_value(0U, 0U, Value::number(1.0));
  source.sheet(0).set_cell_value(4U, 0U, Value::number(2.0));
  RowLayout tall;
  tall.row = 0U;
  tall.height = 33.75;
  tall.has_height = true;
  RowLayout hidden;
  hidden.row = 2U;
  hidden.hidden = true;
  RowLayout grouped;
  grouped.row = 4U;
  grouped.outline_level = 2U;
  source.sheet(0).mutable_layout().row_overrides = {tall, hidden, grouped};

  Workbook via_xlsb = Workbook::create_empty();
  Workbook via_xlsx = Workbook::create_empty();
  ASSERT_TRUE(ThroughXlsb(source, &via_xlsb));
  ASSERT_TRUE(ThroughXlsx(source, &via_xlsx));

  const std::vector<RowLayout>& rb = via_xlsb.sheet(0).layout().row_overrides;
  const std::vector<RowLayout>& rx = via_xlsx.sheet(0).layout().row_overrides;
  ASSERT_EQ(rb.size(), 3U) << "the xlsb path lost a row override";
  ASSERT_EQ(rx.size(), 3U) << "the xlsx path lost a row override";
  for (std::size_t i = 0; i < rb.size(); ++i) {
    SCOPED_TRACE("row override " + std::to_string(i));
    EXPECT_EQ(rb[i].row, rx[i].row);
    EXPECT_EQ(rb[i].has_height, rx[i].has_height);
    // `miyRw` is twips, so a height survives to 1/20 of a point.
    EXPECT_NEAR(rb[i].height, rx[i].height, 1.0 / 20.0);
    EXPECT_EQ(rb[i].hidden, rx[i].hidden);
    EXPECT_EQ(rb[i].outline_level, rx[i].outline_level);
  }
  // The values themselves, not just their agreement: two equally broken
  // paths would satisfy the comparison above on their own.
  EXPECT_TRUE(rb[0].has_height);
  EXPECT_NEAR(rb[0].height, 33.75, 1.0 / 20.0);
  EXPECT_TRUE(rb[1].hidden);
  EXPECT_EQ(rb[2].outline_level, 2U);
}

TEST(XlsbWriteReadSymmetry, BorderRecordContentsSurviveBothFormatsAlike) {
  Workbook source = Workbook::create_empty();
  source.add_sheet("Sheet1");
  source.sheet(0).set_cell_value(0U, 0U, Value::number(1.0));

  io::StylesTable styles;
  styles.fonts.push_back(io::FontRecord{});
  styles.fills.push_back(io::FillRecord{});
  styles.borders.push_back(io::BorderRecord{});
  io::BorderRecord boxed;
  boxed.left.style = 1U;  // thin
  boxed.left.color.kind = io::ColorSpec::Kind::kRgb;
  boxed.left.color.rgb = 0xFF0000FFU;
  boxed.left.color_argb = 0xFF0000FFU;
  boxed.bottom.style = 2U;  // medium
  boxed.bottom.color.kind = io::ColorSpec::Kind::kRgb;
  boxed.bottom.color.rgb = 0xFFFF0000U;
  boxed.bottom.color_argb = 0xFFFF0000U;
  boxed.diagonal.style = 3U;  // dashed
  boxed.diagonal.color.kind = io::ColorSpec::Kind::kRgb;
  boxed.diagonal.color.rgb = 0xFF00FF00U;
  boxed.diagonal.color_argb = 0xFF00FF00U;
  boxed.diagonal_up = true;
  styles.borders.push_back(boxed);
  io::CellXf plain;
  io::CellXf bordered;
  bordered.border_index = 1U;
  bordered.apply_border = true;
  styles.cell_xfs = {plain, bordered};
  source.set_styles(std::move(styles));
  ASSERT_TRUE(static_cast<bool>(source.set_cell_xf_index(0U, 0U, 0U, 1U)));

  Workbook via_xlsb = Workbook::create_empty();
  Workbook via_xlsx = Workbook::create_empty();
  ASSERT_TRUE(ThroughXlsb(source, &via_xlsb));
  ASSERT_TRUE(ThroughXlsx(source, &via_xlsx));

  const std::vector<io::BorderRecord>& bb = via_xlsb.styles().borders;
  const std::vector<io::BorderRecord>& bx = via_xlsx.styles().borders;
  ASSERT_GT(bb.size(), 1U) << "the xlsb path lost the non-default border";
  ASSERT_GT(bx.size(), 1U) << "the xlsx path lost the non-default border";
  ASSERT_EQ(bb.size(), bx.size());
  for (std::size_t i = 0; i < bb.size(); ++i) {
    SCOPED_TRACE("border index " + std::to_string(i));
    const io::BorderRecord& b = bb[i];
    const io::BorderRecord& x = bx[i];
    EXPECT_EQ(b.diagonal_up, x.diagonal_up);
    EXPECT_EQ(b.diagonal_down, x.diagonal_down);
    const io::BorderSide* b_sides[] = {&b.left, &b.right, &b.top, &b.bottom, &b.diagonal};
    const io::BorderSide* x_sides[] = {&x.left, &x.right, &x.top, &x.bottom, &x.diagonal};
    const char* names[] = {"left", "right", "top", "bottom", "diagonal"};
    for (std::size_t side = 0; side < std::size(b_sides); ++side) {
      SCOPED_TRACE(names[side]);
      EXPECT_EQ(b_sides[side]->style, x_sides[side]->style);
      ExpectColorSpecEqual(b_sides[side]->color, x_sides[side]->color);
    }
  }
  // The styles, not merely their agreement: the authored record has to come
  // back with the three sides it declared.
  EXPECT_EQ(bb[1].left.style, 1U);
  EXPECT_EQ(bb[1].bottom.style, 2U);
  EXPECT_EQ(bb[1].diagonal.style, 3U);
  EXPECT_TRUE(bb[1].diagonal_up);
}

// The alignment / protection / `apply*` groups run the same risk in the
// writer as in the reader: `BrtXF` states each of them in a bit of one of
// two flag words, and a bit nobody sets is indistinguishable from a field
// nobody modelled. Authoring one xf that leaves no group at its default and
// pushing it through both writers pins the `.xlsx` -> `.xlsb` direction the
// Excel-authored fixture cannot reach -- it uses only a fraction of the set.
//
// `relativeIndent` is deliberately absent: `BrtXF` has no field for it, so
// it is the one alignment attribute an `.xlsb` save cannot carry.
TEST(XlsbWriteReadSymmetry, AlignmentAndApplyFlagsSurviveBothFormatsAlike) {
  Workbook source = Workbook::create_empty();
  source.add_sheet("Sheet1");
  source.sheet(0).set_cell_value(0U, 0U, Value::number(1.0));

  io::StylesTable styles;
  styles.fonts.push_back(io::FontRecord{});
  styles.fills.push_back(io::FillRecord{});
  styles.borders.push_back(io::BorderRecord{});
  io::CellXf plain;
  io::CellXf decorated;
  decorated.horizontal_align = 2U;  // center
  decorated.vertical_align = 0U;    // top -- not the schema default
  decorated.wrap_text = true;
  decorated.justify_last_line = true;
  decorated.shrink_to_fit = true;
  decorated.has_shrink_to_fit = true;
  decorated.reading_order = 2U;  // right-to-left
  decorated.has_reading_order = true;
  decorated.text_rotation = 45U;
  decorated.has_text_rotation = true;
  decorated.indent = 3U;
  decorated.has_indent = true;
  decorated.quote_prefix = true;
  decorated.has_protection = true;
  decorated.locked = false;
  decorated.hidden = true;
  decorated.apply_number_format = true;
  decorated.apply_font = true;
  decorated.apply_fill = true;
  decorated.apply_border = true;
  decorated.apply_alignment = true;
  decorated.apply_protection = true;
  styles.cell_xfs = {plain, decorated};
  source.set_styles(std::move(styles));
  ASSERT_TRUE(static_cast<bool>(source.set_cell_xf_index(0U, 0U, 0U, 1U)));

  Workbook via_xlsb = Workbook::create_empty();
  Workbook via_xlsx = Workbook::create_empty();
  ASSERT_TRUE(ThroughXlsb(source, &via_xlsb));
  ASSERT_TRUE(ThroughXlsx(source, &via_xlsx));

  const std::vector<io::CellXf>& xb = via_xlsb.styles().cell_xfs;
  const std::vector<io::CellXf>& xx = via_xlsx.styles().cell_xfs;
  ASSERT_GT(xb.size(), 1U) << "the xlsb path lost the decorated xf";
  ASSERT_GT(xx.size(), 1U) << "the xlsx path lost the decorated xf";
  ASSERT_EQ(xb.size(), xx.size());
  for (std::size_t i = 0; i < xb.size(); ++i) {
    SCOPED_TRACE("cellXf index " + std::to_string(i));
    const io::CellXf& b = xb[i];
    const io::CellXf& x = xx[i];
    EXPECT_EQ(b.horizontal_align, x.horizontal_align);
    EXPECT_EQ(b.vertical_align, x.vertical_align);
    EXPECT_EQ(b.wrap_text, x.wrap_text);
    EXPECT_EQ(b.justify_last_line, x.justify_last_line);
    EXPECT_EQ(b.shrink_to_fit, x.shrink_to_fit);
    EXPECT_EQ(b.reading_order, x.reading_order);
    EXPECT_EQ(b.text_rotation, x.text_rotation);
    EXPECT_EQ(b.indent, x.indent);
    EXPECT_EQ(b.quote_prefix, x.quote_prefix);
    EXPECT_EQ(b.has_protection, x.has_protection);
    EXPECT_EQ(b.locked, x.locked);
    EXPECT_EQ(b.hidden, x.hidden);
    EXPECT_EQ(b.apply_number_format, x.apply_number_format);
    EXPECT_EQ(b.apply_font, x.apply_font);
    EXPECT_EQ(b.apply_fill, x.apply_fill);
    EXPECT_EQ(b.apply_border, x.apply_border);
    EXPECT_EQ(b.apply_alignment, x.apply_alignment);
    EXPECT_EQ(b.apply_protection, x.apply_protection);
  }
  // The authored values themselves: two equally lossy paths would satisfy
  // the comparison above on their own.
  const io::CellXf& b = xb[1];
  EXPECT_EQ(b.horizontal_align, 2U);
  EXPECT_EQ(b.vertical_align, 0U);
  EXPECT_TRUE(b.wrap_text);
  EXPECT_TRUE(b.justify_last_line);
  EXPECT_TRUE(b.shrink_to_fit);
  EXPECT_EQ(b.reading_order, 2U);
  EXPECT_EQ(b.text_rotation, 45U);
  EXPECT_EQ(b.indent, 3U);
  EXPECT_TRUE(b.quote_prefix);
  EXPECT_TRUE(b.has_protection);
  EXPECT_FALSE(b.locked);
  EXPECT_TRUE(b.hidden);
  EXPECT_TRUE(b.apply_number_format);
  EXPECT_TRUE(b.apply_font);
  EXPECT_TRUE(b.apply_fill);
  EXPECT_TRUE(b.apply_border);
  EXPECT_TRUE(b.apply_alignment);
  EXPECT_TRUE(b.apply_protection);
}

}  // namespace
}  // namespace formulon
