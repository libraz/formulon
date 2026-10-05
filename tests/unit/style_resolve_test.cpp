#include "style_resolve.h"

#include <algorithm>
#include <cstdint>
#include <cstdio>
#include <set>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/styles_reader.h"
#include "io/styles_writer.h"
#include "styles.h"
#include "theme.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

std::vector<std::uint8_t> ReadFixture(const char* file) {
  std::vector<std::uint8_t> out;
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/" + file;
  FILE* f = std::fopen(path.c_str(), "rb");
  if (f == nullptr) {
    ADD_FAILURE() << "could not open fixture: " << path;
    return out;
  }
  std::fseek(f, 0, SEEK_END);
  const long size = std::ftell(f);
  std::fseek(f, 0, SEEK_SET);
  out.resize(size > 0 ? static_cast<std::size_t>(size) : 0U);
  if (std::fread(out.data(), 1, out.size(), f) != out.size()) {
    ADD_FAILURE() << "short read on fixture: " << path;
    out.clear();
  }
  std::fclose(f);
  return out;
}

Workbook LoadFixture(const char* file) {
  const std::vector<std::uint8_t> bytes = ReadFixture(file);
  auto read = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(read));
  return std::move(read.value().workbook);
}

Workbook RoundTrip(const Workbook& wb) {
  auto bytes = io::write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(bytes));
  auto read = io::read_ooxml(io::ByteSpan{bytes.value().data(), bytes.value().size()});
  EXPECT_TRUE(static_cast<bool>(read));
  return std::move(read.value().workbook);
}

const CellStyleRecord* FindStyle(const StylesTable& styles, std::uint32_t builtin_id) {
  for (const CellStyleRecord& rec : styles.cell_styles) {
    if (rec.builtin_id == builtin_id) {
      return &rec;
    }
  }
  return nullptr;
}

ColorSpec RgbSpec(std::uint32_t argb) {
  ColorSpec spec;
  spec.kind = ColorSpec::Kind::kRgb;
  spec.rgb = argb;
  return spec;
}

// ---------------------------------------------------------------------------
// StyleResolve
// ---------------------------------------------------------------------------

/// A workbook whose xfs 1..3 each carry a distinct solid fill colour.
Workbook WorkbookWithFillXfs() {
  Workbook wb = Workbook::create();
  StylesTable& styles = wb.mutable_styles();
  styles.cell_xfs.assign(4U, CellXf{});
  for (std::uint32_t i = 1; i <= 3U; ++i) {
    FillRecord fill;
    fill.pattern = 1;
    fill.fg = RgbSpec(0xFF000000U | (i * 0x10U));
    styles.fills.push_back(fill);
    styles.cell_xfs[i].fill_index = static_cast<std::uint32_t>(styles.fills.size() - 1U);
  }
  return wb;
}

TEST(StyleResolve, RowStyleBeatsColumnStyle) {
  Workbook wb = WorkbookWithFillXfs();
  Sheet& sheet = wb.sheet(0);
  ColumnLayout col;
  col.first = 0;
  col.last = 3;
  col.has_style = true;
  col.style_xf = 1;
  sheet.mutable_layout().columns.push_back(col);
  RowLayout row;
  row.row = 2;
  row.has_style = true;
  row.style_xf = 2;
  sheet.mutable_layout().row_overrides.push_back(row);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 7U, 0U, Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_xf_index(0, 7U, 0U, 3U)));

  const EffectiveStyle on_row = effective_style(wb, sheet, 2U, 1U);  // row and column both styled
  EXPECT_EQ(on_row.source, StyleSource::kRow);
  EXPECT_EQ(on_row.xf_index, 2U);
  EXPECT_EQ(on_row.fill_foreground.argb, 0xFF000020U);

  const EffectiveStyle on_col = effective_style(wb, sheet, 5U, 1U);
  EXPECT_EQ(on_col.source, StyleSource::kColumn);
  EXPECT_EQ(on_col.xf_index, 1U);

  const EffectiveStyle on_cell = effective_style(wb, sheet, 7U, 0U);  // a cell beats row and column
  EXPECT_EQ(on_cell.source, StyleSource::kCell);
  EXPECT_EQ(on_cell.xf_index, 3U);

  const EffectiveStyle outside = effective_style(wb, sheet, 5U, 9U);
  EXPECT_EQ(outside.source, StyleSource::kDefault);
  EXPECT_EQ(outside.xf_index, 0U);
}

TEST(StyleResolve, ReportsFormatProtectionAndColours) {
  Workbook wb = Workbook::create();
  StylesTable& styles = wb.mutable_styles();
  styles.cell_xfs.assign(2U, CellXf{});
  FontRecord font;
  font.color.kind = ColorSpec::Kind::kTheme;
  font.color.theme = 1;  // dk1
  styles.fonts.push_back(font);
  styles.cell_xfs[1].font_index = static_cast<std::uint32_t>(styles.fonts.size() - 1U);
  styles.cell_xfs[1].num_fmt_id = 14;
  styles.cell_xfs[1].has_protection = true;
  styles.cell_xfs[1].locked = false;
  styles.cell_xfs[1].hidden = true;
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0U, 0U, Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_xf_index(0, 0U, 0U, 1U)));

  const EffectiveStyle st = effective_style(wb, wb.sheet(0), 0U, 0U);
  EXPECT_EQ(st.num_fmt_code, "mm-dd-yy");
  EXPECT_FALSE(st.locked);
  EXPECT_TRUE(st.hidden);
  EXPECT_EQ(st.font_color.argb, 0xFF000000U);
  EXPECT_EQ(st.font_color.resolution, ColorResolution::kDefaultTheme);

  const EffectiveStyle plain = effective_style(wb, wb.sheet(0), 5U, 5U);
  EXPECT_EQ(plain.num_fmt_code, "General");
  EXPECT_TRUE(plain.locked);
  EXPECT_FALSE(plain.hidden);
  EXPECT_EQ(plain.fill_background.argb, 0xFFFFFFFFU);
  EXPECT_EQ(plain.border_colors[0].resolution, ColorResolution::kAutoContext);
}

// ---------------------------------------------------------------------------
// Styles: built-in ids, cellStyle removal, indexed colours
// ---------------------------------------------------------------------------

TEST(StylesBuiltinIds, ExcelSavedRangeZeroTo53LoadsAndRoundTrips) {
  const Workbook wb = LoadFixture("builtin_cell_styles.xlsx");
  const StylesTable& styles = wb.styles();
  std::set<std::uint32_t> loaded;
  for (const CellStyleRecord& rec : styles.cell_styles) {
    loaded.insert(rec.builtin_id);
  }
  EXPECT_EQ(styles.cell_styles.size(), 54U);
  for (std::uint32_t id = 0; id <= CellStyleRecord::kBuiltinIdMax; ++id) {
    EXPECT_EQ(loaded.count(id), 1U) << "builtinId " << id;
  }
  ASSERT_NE(FindStyle(styles, 48U), nullptr);
  EXPECT_EQ(FindStyle(styles, 48U)->name, "60% - \xE3\x82\xA2\xE3\x82\xAF\xE3\x82\xBB\xE3\x83\xB3\xE3\x83\x88 5");
  ASSERT_NE(FindStyle(styles, 53U), nullptr);
  EXPECT_EQ(FindStyle(styles, 53U)->name, "\xE8\xAA\xAC\xE6\x98\x8E\xE6\x96\x87");
  EXPECT_EQ(FindStyle(styles, 1U)->i_level, 0U);

  const Workbook reloaded = RoundTrip(wb);
  ASSERT_EQ(reloaded.styles().cell_styles.size(), 54U);
  for (std::uint32_t id = 0; id <= CellStyleRecord::kBuiltinIdMax; ++id) {
    const CellStyleRecord* before = FindStyle(styles, id);
    const CellStyleRecord* after = FindStyle(reloaded.styles(), id);
    ASSERT_NE(after, nullptr) << "builtinId " << id;
    EXPECT_EQ(after->name, before->name);
    EXPECT_EQ(after->xf_id, before->xf_id);
  }
}

TEST(StylesBuiltinIds, ExcelSavedThemeIsParsed) {
  const Workbook wb = LoadFixture("builtin_cell_styles.xlsx");
  const LoadedTheme loaded = wb.load_theme();
  EXPECT_EQ(loaded.source, ThemeSource::kPart);
  EXPECT_EQ(loaded.theme.colors[4], 0xFF4F81BDU);  // accent1 of the fixture's own theme part
}

/// The styles of the deletion measurement: a custom style (xf 1: 0.000 format,
/// bold red font, yellow fill, right alignment, unlocked) and cell xfs 1..7
/// referencing it with different apply flags.
StylesTable BuildDeletionTable() {
  StylesTable t;
  t.fonts.assign(3U, FontRecord{});  // 0 Normal, 1 bold red, 2 italic
  t.fonts[1].has_bold = true;
  t.fonts[1].bold = true;
  t.fonts[2].has_italic = true;
  t.fonts[2].italic = true;
  t.fills.assign(4U, FillRecord{});  // 0 none, 1 gray125, 2 yellow, 3 green
  t.borders.assign(1U, BorderRecord{});
  t.cell_style_xfs.assign(2U, CellXf{});
  CellXf& style = t.cell_style_xfs[1];
  style.num_fmt_id = 176;
  style.font_index = 1;
  style.fill_index = 2;
  style.horizontal_align = 3;
  style.has_alignment = true;
  style.has_protection = true;
  style.locked = false;
  CellStyleRecord normal;
  normal.name = "Normal";
  normal.xf_id = 0;
  normal.builtin_id = 0;
  CellStyleRecord probe;
  probe.name = "ProbeStyle";
  probe.xf_id = 1;
  t.cell_styles = {normal, probe};

  t.cell_xfs.assign(8U, style);
  for (CellXf& xf : t.cell_xfs) {
    xf.xf_id = 1;
  }
  t.cell_xfs[0] = CellXf{};  // the default xf
  t.cell_xfs[2].font_index = 2;
  t.cell_xfs[2].apply_font = true;
  t.cell_xfs[3].num_fmt_id = 177;
  t.cell_xfs[3].apply_number_format = true;
  t.cell_xfs[4].fill_index = 3;
  t.cell_xfs[4].apply_fill = true;
  t.cell_xfs[5].horizontal_align = 2;
  t.cell_xfs[5].apply_alignment = true;
  t.cell_xfs[6].apply_protection = true;
  t.cell_xfs[6].has_protection = false;
  t.cell_xfs[6].locked = true;
  t.cell_xfs[7] = CellXf{};
  t.cell_xfs[7].xf_id = 1;
  t.cell_xfs[7].apply_number_format = t.cell_xfs[7].apply_font = t.cell_xfs[7].apply_fill = true;
  t.cell_xfs[7].apply_border = t.cell_xfs[7].apply_alignment = t.cell_xfs[7].apply_protection = true;
  return t;
}

TEST(StylesCellStyleRemoval, RepointsXfsAndKeepsAppliedGroups) {
  StylesTable t = BuildDeletionTable();
  ASSERT_TRUE(static_cast<bool>(remove_cell_style(t, "ProbeStyle")));

  ASSERT_EQ(t.cell_styles.size(), 1U);
  EXPECT_EQ(t.cell_styles[0].name, "Normal");
  EXPECT_EQ(t.cell_style_xfs.size(), 2U);  // not compacted
  ASSERT_EQ(t.cell_xfs.size(), 8U);
  for (std::size_t i = 1; i < t.cell_xfs.size(); ++i) {
    EXPECT_EQ(t.cell_xfs[i].xf_id, 0U) << "xf " << i;
  }
  // xf 1: no apply flag, every group takes Normal's value.
  const CellXf& x1 = t.cell_xfs[1];
  EXPECT_EQ(x1.num_fmt_id, 0U);
  EXPECT_EQ(x1.font_index, 0U);
  EXPECT_EQ(x1.fill_index, 0U);
  EXPECT_EQ(x1.horizontal_align, 0U);
  EXPECT_FALSE(x1.has_alignment);
  EXPECT_FALSE(x1.has_protection);
  EXPECT_TRUE(x1.locked);
  // xf 2..6: only the applied group keeps its own value.
  EXPECT_EQ(t.cell_xfs[2].font_index, 2U);
  EXPECT_EQ(t.cell_xfs[2].fill_index, 0U);
  EXPECT_EQ(t.cell_xfs[2].num_fmt_id, 0U);
  EXPECT_EQ(t.cell_xfs[3].num_fmt_id, 177U);
  EXPECT_EQ(t.cell_xfs[3].font_index, 0U);
  EXPECT_EQ(t.cell_xfs[4].fill_index, 3U);
  EXPECT_EQ(t.cell_xfs[4].font_index, 0U);
  EXPECT_EQ(t.cell_xfs[5].horizontal_align, 2U);
  EXPECT_EQ(t.cell_xfs[5].fill_index, 0U);
  EXPECT_FALSE(t.cell_xfs[6].has_protection);
  EXPECT_TRUE(t.cell_xfs[6].locked);
  EXPECT_EQ(t.cell_xfs[6].horizontal_align, 0U);
  // xf 7 keeps its flags (not recomputed) and its own Normal-valued groups.
  EXPECT_TRUE(t.cell_xfs[7].apply_font && t.cell_xfs[7].apply_fill && t.cell_xfs[7].apply_protection);
  EXPECT_EQ(t.cell_xfs[7].font_index, 0U);
}

TEST(StylesCellStyleRemoval, KeepsXfsWhileAnotherStyleSharesTheRecord) {
  StylesTable t = BuildDeletionTable();
  CellStyleRecord twin;
  twin.name = "Twin";
  twin.xf_id = 1;
  t.cell_styles.push_back(twin);
  ASSERT_TRUE(static_cast<bool>(remove_cell_style(t, "ProbeStyle")));
  ASSERT_EQ(t.cell_styles.size(), 2U);
  EXPECT_EQ(t.cell_xfs[1].xf_id, 1U);
  EXPECT_EQ(t.cell_xfs[1].font_index, 1U);
}

TEST(StylesCellStyleRemoval, RejectsNormalAndUnknownNames) {
  StylesTable t = BuildDeletionTable();
  EXPECT_FALSE(static_cast<bool>(remove_cell_style(t, "Normal")));
  EXPECT_FALSE(static_cast<bool>(remove_cell_style(t, "NoSuchStyle")));
  EXPECT_EQ(t.cell_styles.size(), 2U);
  EXPECT_EQ(t.cell_xfs[1].xf_id, 1U);
}

TEST(StylesCellStyleRemoval, SurvivesSaveAndReload) {
  Workbook wb = LoadFixture("builtin_cell_styles.xlsx");
  ASSERT_TRUE(static_cast<bool>(remove_cell_style(wb.mutable_styles(), "\xE8\xAA\xAC\xE6\x98\x8E\xE6\x96\x87")));
  const Workbook reloaded = RoundTrip(wb);
  EXPECT_EQ(reloaded.styles().cell_styles.size(), 53U);
  EXPECT_EQ(FindStyle(reloaded.styles(), 53U), nullptr);
  EXPECT_NE(FindStyle(reloaded.styles(), 52U), nullptr);
}

TEST(StylesIndexedColors, RoundTripKeepsPaletteAndOtherColorsChildren) {
  const std::string source =
      "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
      "<colors><indexedColors><rgbColor rgb=\"FF112233\"/><rgbColor rgb=\"FF445566\"/></indexedColors>"
      "<mruColors><color rgb=\"FFAABBCC\"/></mruColors></colors></styleSheet>";
  auto read = io::read_styles(std::vector<std::uint8_t>(source.begin(), source.end()));
  ASSERT_TRUE(static_cast<bool>(read)) << read.error().message;
  EXPECT_EQ(read.value().indexed_colors, (std::vector<std::uint32_t>{0xFF112233U, 0xFF445566U}));
  EXPECT_EQ(read.value().colors_xml, "<colors><mruColors><color rgb=\"FFAABBCC\"/></mruColors></colors>");

  const std::string xml = io::write_styles(read.value());
  EXPECT_NE(xml.find("<colors><indexedColors><rgbColor rgb=\"FF112233\"/><rgbColor rgb=\"FF445566\"/></indexedColors>"
                     "<mruColors><color rgb=\"FFAABBCC\"/></mruColors></colors>"),
            std::string::npos);
  auto again = io::read_styles(std::vector<std::uint8_t>(xml.begin(), xml.end()));
  ASSERT_TRUE(static_cast<bool>(again));
  EXPECT_EQ(again.value().indexed_colors, read.value().indexed_colors);
  EXPECT_EQ(again.value().colors_xml, read.value().colors_xml);
}

TEST(StylesIndexedColors, ResolvedThroughWorkbookPalette) {
  Workbook wb = Workbook::create();
  wb.mutable_styles().indexed_colors = {0xFF010203U, 0xFF040506U};
  const Workbook reloaded = RoundTrip(wb);
  EXPECT_EQ(reloaded.styles().indexed_colors, wb.styles().indexed_colors);
  ColorSpec spec;
  spec.kind = ColorSpec::Kind::kIndexed;
  spec.indexed = 1;
  EXPECT_EQ(resolve_color(reloaded, spec, ColorContext::kFont).argb, 0xFF040506U);
  spec.indexed = 2;
  EXPECT_EQ(resolve_color(reloaded, spec, ColorContext::kFont).argb, 0xFFFF0000U);
}

TEST(StylesIndexedColors, AbsentPaletteWritesNoColorsElement) {
  const std::string xml = io::write_styles(StylesTable{});
  EXPECT_EQ(xml.find("<colors"), std::string::npos);
}

}  // namespace
}  // namespace formulon
