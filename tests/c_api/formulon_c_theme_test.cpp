// C ABI tests for the workbook theme, colour resolution and effective style.

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <map>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "utils/error.h"

namespace {

static_assert(sizeof(fm_theme_colors) == 48U, "fm_theme_colors ABI layout changed");
static_assert(sizeof(fm_theme_fonts) == 4U * sizeof(void*), "fm_theme_fonts ABI layout changed");
static_assert(offsetof(fm_effective_style, font_argb) == 20U, "fm_effective_style.font_argb offset changed");
static_assert(offsetof(fm_effective_style, border_argb) == 44U, "fm_effective_style.border_argb offset changed");
static_assert(offsetof(fm_effective_style, border_resolution) == 64U,
              "fm_effective_style.border_resolution offset changed");
static_assert(offsetof(fm_effective_style, hidden) == 88U, "fm_effective_style.hidden offset changed");
static_assert(offsetof(fm_effective_style, num_fmt_code) == (sizeof(void*) == 4U ? 92U : 96U),
              "fm_effective_style.num_fmt_code offset changed");
static_assert(sizeof(fm_effective_style) == (sizeof(void*) == 4U ? 96U : 104U),
              "fm_effective_style ABI layout changed");

using Bytes = std::vector<std::uint8_t>;

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);
constexpr fm_status_t kXmlParse = static_cast<fm_status_t>(formulon::FormulonErrorCode::kIoXmlParse);

constexpr int32_t kSourcePart = 0;
constexpr int32_t kSourceDefault = 1;
constexpr int32_t kSourceUnparseable = 2;

constexpr uint32_t kDefaultScheme[12] = {0xFF000000U, 0xFFFFFFFFU, 0xFF44546AU, 0xFFE7E6E6U, 0xFF4472C4U, 0xFFED7D31U,
                                         0xFFA5A5A5U, 0xFFFFC000U, 0xFF5B9BD5U, 0xFF70AD47U, 0xFF0563C1U, 0xFF954F72U};
constexpr const char* kYuGothic = "\xE6\xB8\xB8\xE3\x82\xB4\xE3\x82\xB7\xE3\x83\x83\xE3\x82\xAF";
constexpr const char* kYuGothicLight = "\xE6\xB8\xB8\xE3\x82\xB4\xE3\x82\xB7\xE3\x83\x83\xE3\x82\xAF Light";

Bytes Save(fm_workbook_t* wb) {
  BufferGuard buffer;
  EXPECT_EQ(fm_workbook_save(wb, &buffer.data, &buffer.len), 0);
  return Bytes(buffer.data, buffer.data + buffer.len);
}

void Load(const Bytes& bytes, WorkbookGuard* out) {
  ASSERT_EQ(fm_workbook_load(bytes.data(), bytes.size(), &out->handle), 0);
}

/// Rewrites a package, replacing the entries named in `replacements`. Every
/// replaced entry must already exist, so a misspelt path fails the test.
Bytes ReplaceEntries(const Bytes& package, const std::map<std::string, std::string>& replacements) {
  formulon::io::ZipReader input;
  EXPECT_TRUE(static_cast<bool>(input.open(formulon::io::ByteSpan{package.data(), package.size()})));
  mz_zip_archive writer{};
  EXPECT_EQ(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_TRUE);
  std::size_t replaced = 0;
  for (const std::string& name : input.list_entries()) {
    auto body = input.read_entry(name);
    EXPECT_TRUE(static_cast<bool>(body)) << name;
    if (!body) {
      continue;
    }
    Bytes content = body.value();
    if (auto it = replacements.find(name); it != replacements.end()) {
      content.assign(it->second.begin(), it->second.end());
      ++replaced;
    }
    EXPECT_EQ(mz_zip_writer_add_mem(&writer, name.c_str(), content.data(), content.size(),
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_TRUE);
  }
  EXPECT_EQ(replaced, replacements.size());
  void* archive = nullptr;
  std::size_t archive_size = 0;
  EXPECT_EQ(mz_zip_writer_finalize_heap_archive(&writer, &archive, &archive_size), MZ_TRUE);
  EXPECT_EQ(mz_zip_writer_end(&writer), MZ_TRUE);
  Bytes result(static_cast<const std::uint8_t*>(archive), static_cast<const std::uint8_t*>(archive) + archive_size);
  mz_free(archive);
  return result;
}

fm_color_spec ThemeSpec(uint32_t index, double tint) {
  fm_color_spec spec{};
  spec.kind = kFmColorTheme;
  spec.theme = index;
  spec.tint = tint;
  return spec;
}

fm_color_spec IndexedSpec(uint32_t index) {
  fm_color_spec spec{};
  spec.kind = kFmColorIndexed;
  spec.indexed = index;
  return spec;
}

struct Resolved {
  fm_status_t status = -1;
  uint32_t argb = 0;
  int32_t resolution = -1;
};

Resolved Resolve(fm_workbook_t* wb, fm_color_spec spec, int32_t context) {
  Resolved out;
  out.status = fm_workbook_resolve_color(wb, spec, context, &out.argb, &out.resolution);
  return out;
}

fm_theme_colors ReadColors(fm_workbook_t* wb, int32_t* source) {
  fm_theme_colors colors{};
  EXPECT_EQ(fm_workbook_get_theme_colors(wb, &colors, source), 0);
  return colors;
}

}  // namespace

TEST(FormulonCApiTheme, WorkbookWithoutThemePartReportsDefaultTheme) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  int32_t source = -1;
  const fm_theme_colors colors = ReadColors(wb.handle, &source);
  EXPECT_EQ(source, kSourceDefault);
  for (std::size_t i = 0; i < 12U; ++i) {
    EXPECT_EQ(colors.argb[i], kDefaultScheme[i]) << i;
  }
  fm_theme_fonts fonts{};
  ASSERT_EQ(fm_workbook_get_theme_fonts(wb.handle, &fonts), 0);
  EXPECT_STREQ(fonts.major_latin, "Calibri Light");
  EXPECT_STREQ(fonts.major_east_asian, kYuGothicLight);
  EXPECT_STREQ(fonts.minor_latin, "Calibri");
  EXPECT_STREQ(fonts.minor_east_asian, kYuGothic);
}

TEST(FormulonCApiTheme, SetColorsOnWorkbookWithoutThemeSurvivesSaveReload) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_theme_colors colors{};
  for (uint32_t i = 0; i < 12U; ++i) {
    colors.argb[i] = 0xFF101010U + i;
  }
  colors.argb[4] = 0x00123456U;  // the alpha byte is not stored
  ASSERT_EQ(fm_workbook_set_theme_colors(wb.handle, &colors), 0);

  int32_t source = -1;
  fm_theme_colors live = ReadColors(wb.handle, &source);
  EXPECT_EQ(source, kSourcePart);
  EXPECT_EQ(live.argb[0], 0xFF101010U);
  EXPECT_EQ(live.argb[4], 0xFF123456U);
  EXPECT_EQ(live.argb[11], 0xFF10101BU);
  // The generated part carries the default fonts.
  fm_theme_fonts fonts{};
  ASSERT_EQ(fm_workbook_get_theme_fonts(wb.handle, &fonts), 0);
  EXPECT_STREQ(fonts.minor_latin, "Calibri");

  WorkbookGuard loaded;
  Load(Save(wb.handle), &loaded);
  source = -1;
  const fm_theme_colors reread = ReadColors(loaded.handle, &source);
  EXPECT_EQ(source, kSourcePart);
  for (std::size_t i = 0; i < 12U; ++i) {
    EXPECT_EQ(reread.argb[i], live.argb[i]) << i;
  }
}

TEST(FormulonCApiTheme, SetFontsSurvivesSaveReload) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_theme_fonts fonts{};
  fonts.major_latin = "Georgia";
  fonts.major_east_asian = "MS Mincho";
  fonts.minor_latin = "Arial";
  fonts.minor_east_asian = "Meiryo";
  ASSERT_EQ(fm_workbook_set_theme_fonts(wb.handle, &fonts), 0);

  WorkbookGuard loaded;
  Load(Save(wb.handle), &loaded);
  fm_theme_fonts reread{};
  ASSERT_EQ(fm_workbook_get_theme_fonts(loaded.handle, &reread), 0);
  EXPECT_STREQ(reread.major_latin, "Georgia");
  EXPECT_STREQ(reread.major_east_asian, "MS Mincho");
  EXPECT_STREQ(reread.minor_latin, "Arial");
  EXPECT_STREQ(reread.minor_east_asian, "Meiryo");
  // Setting the fonts generated the part with the default colours.
  int32_t source = -1;
  const fm_theme_colors colors = ReadColors(loaded.handle, &source);
  EXPECT_EQ(source, kSourcePart);
  EXPECT_EQ(colors.argb[4], kDefaultScheme[4]);
}

TEST(FormulonCApiTheme, UnparseablePartIsReportedAndRefusesEdits) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_theme_colors colors{};
  for (uint32_t i = 0; i < 12U; ++i) {
    colors.argb[i] = 0xFF202020U;
  }
  ASSERT_EQ(fm_workbook_set_theme_colors(wb.handle, &colors), 0);
  const Bytes broken =
      ReplaceEntries(Save(wb.handle), {{"xl/theme/theme1.xml", "<a:theme>not a colour scheme</a:theme>"}});

  WorkbookGuard loaded;
  Load(broken, &loaded);
  int32_t source = -1;
  const fm_theme_colors reread = ReadColors(loaded.handle, &source);
  EXPECT_EQ(source, kSourceUnparseable);
  EXPECT_EQ(reread.argb[4], kDefaultScheme[4]);
  EXPECT_EQ(fm_workbook_set_theme_colors(loaded.handle, &colors), kXmlParse);
  fm_theme_fonts fonts{};
  fonts.major_latin = fonts.major_east_asian = fonts.minor_latin = fonts.minor_east_asian = "Arial";
  EXPECT_EQ(fm_workbook_set_theme_fonts(loaded.handle, &fonts), kXmlParse);

  const Resolved accent = Resolve(loaded.handle, ThemeSpec(4, 0.0), FM_COLOR_CONTEXT_FONT);
  ASSERT_EQ(accent.status, 0);
  EXPECT_EQ(accent.argb, kDefaultScheme[4]);
  EXPECT_EQ(accent.resolution, FM_COLOR_RESOLUTION_THEME_UNPARSEABLE);
}

TEST(FormulonCApiTheme, ThemeAccessorsRejectNullArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_theme_colors colors{};
  fm_theme_fonts fonts{};
  int32_t source = 0;
  EXPECT_EQ(fm_workbook_get_theme_colors(nullptr, &colors, &source), kNullPointer);
  EXPECT_EQ(fm_workbook_get_theme_colors(wb.handle, nullptr, &source), kNullPointer);
  EXPECT_EQ(fm_workbook_get_theme_colors(wb.handle, &colors, nullptr), kNullPointer);
  EXPECT_EQ(fm_workbook_set_theme_colors(nullptr, &colors), kNullPointer);
  EXPECT_EQ(fm_workbook_set_theme_colors(wb.handle, nullptr), kNullPointer);
  EXPECT_EQ(fm_workbook_get_theme_fonts(nullptr, &fonts), kNullPointer);
  EXPECT_EQ(fm_workbook_get_theme_fonts(wb.handle, nullptr), kNullPointer);
  fonts.major_latin = "A";
  fonts.major_east_asian = "B";
  fonts.minor_latin = "C";
  EXPECT_EQ(fm_workbook_set_theme_fonts(wb.handle, &fonts), kNullPointer);
  // A rejected edit generates no theme part.
  ReadColors(wb.handle, &source);
  EXPECT_EQ(source, kSourceDefault);
}

TEST(FormulonCApiTheme, ResolveColorReportsResolution) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_color_spec rgb{};
  rgb.kind = kFmColorRgb;
  rgb.rgb = 0xFF123456U;
  Resolved got = Resolve(wb.handle, rgb, FM_COLOR_CONTEXT_FILL_FG);
  ASSERT_EQ(got.status, 0);
  EXPECT_EQ(got.argb, 0xFF123456U);
  EXPECT_EQ(got.resolution, FM_COLOR_RESOLUTION_EXACT);

  // No theme part: theme colours resolve against the default theme.
  got = Resolve(wb.handle, ThemeSpec(4, 0.0), FM_COLOR_CONTEXT_FONT);
  EXPECT_EQ(got.argb, 0xFF4472C4U);
  EXPECT_EQ(got.resolution, FM_COLOR_RESOLUTION_DEFAULT_THEME);
  got = Resolve(wb.handle, ThemeSpec(12, 0.0), FM_COLOR_CONTEXT_FONT);
  EXPECT_EQ(got.argb, 0xFF000000U);
  EXPECT_EQ(got.resolution, FM_COLOR_RESOLUTION_INDEX_OUT_OF_RANGE);

  fm_theme_colors colors{};
  for (std::size_t i = 0; i < 12U; ++i) {
    colors.argb[i] = kDefaultScheme[i];
  }
  ASSERT_EQ(fm_workbook_set_theme_colors(wb.handle, &colors), 0);
  // Theme index 0 is lt1 and 1 is dk1: the first pairs swap against the scheme.
  got = Resolve(wb.handle, ThemeSpec(0, 0.0), FM_COLOR_CONTEXT_FONT);
  EXPECT_EQ(got.argb, 0xFFFFFFFFU);
  EXPECT_EQ(got.resolution, FM_COLOR_RESOLUTION_EXACT);
  got = Resolve(wb.handle, ThemeSpec(1, 0.0), FM_COLOR_CONTEXT_FONT);
  EXPECT_EQ(got.argb, 0xFF000000U);
  // Excel's "Accent 1, Lighter 40%" and "Darker 25%".
  got = Resolve(wb.handle, ThemeSpec(4, 0.4), FM_COLOR_CONTEXT_FILL_FG);
  EXPECT_EQ(got.argb, 0xFF8EA9DBU);
  EXPECT_EQ(got.resolution, FM_COLOR_RESOLUTION_EXACT);
  got = Resolve(wb.handle, ThemeSpec(4, -0.25), FM_COLOR_CONTEXT_FILL_FG);
  EXPECT_EQ(got.argb, 0xFF305496U);

  got = Resolve(wb.handle, IndexedSpec(2), FM_COLOR_CONTEXT_FONT);
  EXPECT_EQ(got.argb, 0xFFFF0000U);
  EXPECT_EQ(got.resolution, FM_COLOR_RESOLUTION_EXACT);
  got = Resolve(wb.handle, IndexedSpec(80), FM_COLOR_CONTEXT_FONT);
  EXPECT_EQ(got.argb, 0xFF000000U);
  EXPECT_EQ(got.resolution, FM_COLOR_RESOLUTION_INDEX_OUT_OF_RANGE);

  // System foreground (64) and automatic colours depend on the context.
  fm_color_spec automatic{};
  automatic.kind = kFmColorAuto;
  const struct {
    int32_t context;
    uint32_t argb;
  } kContexts[] = {
      {FM_COLOR_CONTEXT_FONT, 0xFF000000U},
      {FM_COLOR_CONTEXT_FILL_FG, 0xFFFFFFFFU},
      {FM_COLOR_CONTEXT_FILL_BG, 0xFFFFFFFFU},
      {FM_COLOR_CONTEXT_BORDER, 0xFF000000U},
  };
  for (const auto& c : kContexts) {
    got = Resolve(wb.handle, automatic, c.context);
    ASSERT_EQ(got.status, 0) << c.context;
    EXPECT_EQ(got.argb, c.argb) << c.context;
    EXPECT_EQ(got.resolution, FM_COLOR_RESOLUTION_AUTO_CONTEXT) << c.context;
    got = Resolve(wb.handle, IndexedSpec(64), c.context);
    EXPECT_EQ(got.argb, c.argb) << c.context;
    EXPECT_EQ(got.resolution, FM_COLOR_RESOLUTION_AUTO_CONTEXT) << c.context;
  }
}

TEST(FormulonCApiTheme, ResolveColorRejectsInvalidArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t argb = 0;
  int32_t resolution = 0;
  EXPECT_EQ(fm_workbook_resolve_color(nullptr, ThemeSpec(4, 0.0), 0, &argb, &resolution), kNullPointer);
  EXPECT_EQ(fm_workbook_resolve_color(wb.handle, ThemeSpec(4, 0.0), 0, nullptr, &resolution), kNullPointer);
  EXPECT_EQ(fm_workbook_resolve_color(wb.handle, ThemeSpec(4, 0.0), 0, &argb, nullptr), kNullPointer);
  EXPECT_EQ(Resolve(wb.handle, ThemeSpec(4, 0.0), 4).status, kInvalidArgument);
  EXPECT_EQ(Resolve(wb.handle, ThemeSpec(4, 0.0), -1).status, kInvalidArgument);
  fm_color_spec unknown{};
  unknown.kind = 5;
  EXPECT_EQ(Resolve(wb.handle, unknown, FM_COLOR_CONTEXT_FONT).status, kInvalidArgument);
  EXPECT_EQ(Resolve(wb.handle, ThemeSpec(4, std::numeric_limits<double>::quiet_NaN()), FM_COLOR_CONTEXT_FONT).status,
            kInvalidArgument);
}

namespace {

/// Style table for the precedence tests: a distinct xf per level, the cell
/// one carrying a themed font, a solid literal fill, an indexed left border,
/// a date format and explicit protection.
struct PrecedenceXfs {
  uint32_t cell = 0;
  uint32_t row = 0;
  uint32_t column = 0;
};

PrecedenceXfs AddPrecedenceXfs(fm_workbook_t* wb) {
  fm_font_record font{};
  font.name = "Calibri";
  font.size = 11.0;
  font.color.kind = kFmColorTheme;
  font.color.theme = 4;
  font.color.tint = 0.4;
  uint32_t font_index = 0;
  EXPECT_EQ(fm_styles_add_font(wb, font, &font_index), 0);
  fm_fill_record fill{};
  fill.pattern = 1;
  fill.fg.kind = kFmColorRgb;
  fill.fg.rgb = 0xFFFF0000U;
  fill.fg_argb = 0xFFFF0000U;
  uint32_t fill_index = 0;
  EXPECT_EQ(fm_styles_add_fill(wb, fill, &fill_index), 0);
  fm_border_record border{};
  border.left.style = 1;
  border.left.color.kind = kFmColorIndexed;
  border.left.color.indexed = 12;
  uint32_t border_index = 0;
  EXPECT_EQ(fm_styles_add_border(wb, border, &border_index), 0);

  PrecedenceXfs out;
  fm_cell_xf cell{};
  cell.font_index = font_index;
  cell.fill_index = fill_index;
  cell.border_index = border_index;
  cell.num_fmt_id = 14;
  cell.has_protection = 1;
  cell.locked = 0;
  cell.hidden = 1;
  EXPECT_EQ(fm_styles_add_cell_xf(wb, cell, &out.cell), 0);
  fm_cell_xf row{};
  row.num_fmt_id = 10;
  EXPECT_EQ(fm_styles_add_cell_xf(wb, row, &out.row), 0);
  fm_cell_xf column{};
  column.num_fmt_id = 3;
  EXPECT_EQ(fm_styles_add_cell_xf(wb, column, &out.column), 0);
  return out;
}

/// Saves `wb` with its only sheet replaced: column C styled, row 5 styled
/// with `customFormat`, row 7 carrying a style without it, and a cell at A5.
Bytes PrecedencePackage(fm_workbook_t* wb, const PrecedenceXfs& xfs) {
  const std::string sheet =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
      "<cols><col min=\"3\" max=\"3\" width=\"9\" style=\"" +
      std::to_string(xfs.column) +
      "\"/></cols><sheetData>"
      "<row r=\"5\" s=\"" +
      std::to_string(xfs.row) + "\" customFormat=\"1\"><c r=\"A5\" s=\"" + std::to_string(xfs.cell) +
      "\"><v>1</v></c></row>"
      "<row r=\"7\" s=\"" +
      std::to_string(xfs.row) +
      "\"><c r=\"D7\"><v>2</v></c></row>"
      "</sheetData></worksheet>";
  return ReplaceEntries(Save(wb), {{"xl/worksheets/sheet1.xml", sheet}});
}

fm_effective_style Effective(fm_workbook_t* wb, uint32_t row, uint32_t col) {
  fm_effective_style style{};
  EXPECT_EQ(fm_sheet_get_effective_style(wb, 0, row, col, &style), 0);
  return style;
}

}  // namespace

TEST(FormulonCApiEffectiveStyle, CellRowColumnDefaultPrecedence) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const PrecedenceXfs xfs = AddPrecedenceXfs(wb.handle);
  WorkbookGuard loaded;
  Load(PrecedencePackage(wb.handle, xfs), &loaded);

  const struct {
    uint32_t row;
    uint32_t col;
    int32_t source;
    uint32_t xf;
  } kCases[] = {
      {4, 0, 0, xfs.cell},    // A5: the cell's own xf
      {4, 1, 1, xfs.row},     // B5: no cell, the customFormat row
      {4, 2, 1, xfs.row},     // C5: the row wins over the column
      {0, 2, 2, xfs.column},  // C1: the column
      {6, 2, 2, xfs.column},  // C7: a row style without customFormat is inert
      {6, 0, 3, 0U},          // A7: nothing applies
      {0, 0, 3, 0U},          // A1: nothing applies
  };
  for (const auto& c : kCases) {
    const fm_effective_style style = Effective(loaded.handle, c.row, c.col);
    EXPECT_EQ(style.source, c.source) << c.row << "," << c.col;
    EXPECT_EQ(style.xf_index, c.xf) << c.row << "," << c.col;
  }
  // An existing cell without its own style still reads as the cell level.
  const fm_effective_style d7 = Effective(loaded.handle, 6, 3);
  EXPECT_EQ(d7.source, 0);
  EXPECT_EQ(d7.xf_index, 0U);
}

TEST(FormulonCApiEffectiveStyle, ResolvesColorsFormatAndProtection) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const PrecedenceXfs xfs = AddPrecedenceXfs(wb.handle);
  ASSERT_EQ(fm_cell_set_xf_index(wb.handle, 0, 0, 0, xfs.cell), 0);

  const fm_effective_style style = Effective(wb.handle, 0, 0);
  fm_cell_xf xf{};
  ASSERT_EQ(fm_styles_get_cell_xf(wb.handle, xfs.cell, &xf), 0);
  EXPECT_EQ(style.font_index, xf.font_index);
  EXPECT_EQ(style.fill_index, xf.fill_index);
  EXPECT_EQ(style.border_index, xf.border_index);
  // No theme part: the themed font resolves against the default theme.
  EXPECT_EQ(style.font_argb, 0xFF8EA9DBU);
  EXPECT_EQ(style.font_resolution, FM_COLOR_RESOLUTION_DEFAULT_THEME);
  EXPECT_EQ(style.fill_fg_argb, 0xFFFF0000U);
  EXPECT_EQ(style.fill_fg_resolution, FM_COLOR_RESOLUTION_EXACT);
  EXPECT_EQ(style.fill_bg_argb, 0xFFFFFFFFU);
  EXPECT_EQ(style.fill_bg_resolution, FM_COLOR_RESOLUTION_AUTO_CONTEXT);
  EXPECT_EQ(style.border_argb[0], 0xFF0000FFU);  // palette index 12
  EXPECT_EQ(style.border_resolution[0], FM_COLOR_RESOLUTION_EXACT);
  for (std::size_t side = 1; side < 5U; ++side) {
    EXPECT_EQ(style.border_argb[side], 0xFF000000U) << side;
    EXPECT_EQ(style.border_resolution[side], FM_COLOR_RESOLUTION_AUTO_CONTEXT) << side;
  }
  EXPECT_EQ(style.locked, 0);
  EXPECT_EQ(style.hidden, 1);
  const char* date_code = nullptr;
  ASSERT_EQ(fm_styles_get_num_fmt_string(wb.handle, 14, &date_code), 0);
  ASSERT_NE(style.num_fmt_code, nullptr);
  EXPECT_STREQ(style.num_fmt_code, date_code);

  // Once the theme part exists the same font colour is exact.
  fm_theme_colors colors{};
  for (std::size_t i = 0; i < 12U; ++i) {
    colors.argb[i] = kDefaultScheme[i];
  }
  ASSERT_EQ(fm_workbook_set_theme_colors(wb.handle, &colors), 0);
  EXPECT_EQ(Effective(wb.handle, 0, 0).font_resolution, FM_COLOR_RESOLUTION_EXACT);

  const fm_effective_style plain = Effective(wb.handle, 3, 3);
  EXPECT_EQ(plain.source, 3);
  EXPECT_EQ(plain.locked, 1);
  EXPECT_EQ(plain.hidden, 0);
  EXPECT_STREQ(plain.num_fmt_code, "General");
}

TEST(FormulonCApiEffectiveStyle, RejectsInvalidArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_effective_style style{};
  EXPECT_EQ(fm_sheet_get_effective_style(nullptr, 0, 0, 0, &style), kNullPointer);
  EXPECT_EQ(fm_sheet_get_effective_style(wb.handle, 0, 0, 0, nullptr), kNullPointer);
  EXPECT_EQ(fm_sheet_get_effective_style(wb.handle, 9, 0, 0, &style), kInvalidArgument);
  EXPECT_EQ(fm_sheet_get_effective_style(wb.handle, 0, 1048576U, 0, &style), kInvalidArgument);
  EXPECT_EQ(fm_sheet_get_effective_style(wb.handle, 0, 0, 16384U, &style), kInvalidArgument);
}
