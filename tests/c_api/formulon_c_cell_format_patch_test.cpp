// Stable C ABI cell-format patch tests (`fm_sheet_update_cell_format`).

#include <array>
#include <cmath>
#include <cstdint>
#include <limits>
#include <string>
#include <vector>

#include "formulon_c_styles_test_helpers.h"

namespace {

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);
constexpr fm_status_t kSheetIndexOutOfRange =
    static_cast<fm_status_t>(formulon::FormulonErrorCode::kSheetIndexOutOfRange);
constexpr fm_status_t kUtf8Error = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingUtf8EncodingError);
constexpr fm_status_t kResourceLimit = static_cast<fm_status_t>(formulon::FormulonErrorCode::kSecurityResourceLimit);

// 游ゴシック, 游明朝: the faces of the measured rename table that differ between the two books.
constexpr char kYuGothic[] = "\xE6\xB8\xB8\xE3\x82\xB4\xE3\x82\xB7\xE3\x83\x83\xE3\x82\xAF";
constexpr char kYuMincho[] = "\xE6\xB8\xB8\xE6\x98\x8E\xE6\x9C\x9D";

constexpr fm_cell_rect_t Cell(uint32_t row, uint32_t col) {
  return fm_cell_rect_t{row, col, row, col};
}

fm_cell_format_patch_t NewPatch() {
  return FM_STRUCT_INIT(fm_cell_format_patch_t);
}

fm_status_t Update(fm_workbook_t* wb, const fm_cell_rect_t& range, const fm_cell_format_patch_t& patch) {
  return fm_sheet_update_cell_format(wb, 0, &range, &patch);
}

struct FontView {
  std::string name;
  fm_font_t raw{};
};

uint32_t ResolvedXf(fm_workbook_t* wb, uint32_t row, uint32_t col) {
  const fm_cell_rect_t rect = Cell(row, col);
  std::array<uint32_t, 4> xfs{};
  size_t count = 0;
  EXPECT_EQ(fm_sheet_get_range_xf_bases(wb, 0, &rect, xfs.data(), xfs.size(), &count), 0);
  EXPECT_EQ(count, 1U);
  return xfs[0];
}

fm_cell_xf_t XfOf(fm_workbook_t* wb, uint32_t row, uint32_t col) {
  fm_cell_xf_t xf{};
  EXPECT_EQ(fm_workbook_get_cell_xf_at(wb, ResolvedXf(wb, row, col), &xf), 0);
  return xf;
}

FontView FontOf(fm_workbook_t* wb, uint32_t row, uint32_t col) {
  FontView view;
  EXPECT_EQ(fm_workbook_get_font_at(wb, XfOf(wb, row, col).font_index, &view.raw), 0);
  view.name = view.raw.name != nullptr ? view.raw.name : "";
  view.raw.name = nullptr;
  return view;
}

fm_fill_t FillOf(fm_workbook_t* wb, uint32_t row, uint32_t col) {
  fm_fill_t fill{};
  EXPECT_EQ(fm_workbook_get_fill_at(wb, XfOf(wb, row, col).fill_index, &fill), 0);
  return fill;
}

fm_border_t BorderOf(fm_workbook_t* wb, uint32_t row, uint32_t col) {
  fm_border_t border{};
  EXPECT_EQ(fm_workbook_get_border_at(wb, XfOf(wb, row, col).border_index, &border), 0);
  return border;
}

size_t Count(fm_status_t (*getter)(const fm_workbook_t*, size_t*), fm_workbook_t* wb) {
  size_t count = 0;
  EXPECT_EQ(getter(wb, &count), 0);
  return count;
}

/// The font record of a cell in the workbook saved from `wb`, read back from the package under `profile`.
FontView SavedFontOf(fm_workbook_t* wb, const char* profile, uint32_t row, uint32_t col) {
  BufferGuard saved;
  EXPECT_EQ(fm_workbook_save(wb, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len), 0) << fm_get_last_error_message();
  fm_load_options_spec_t options = FM_STRUCT_INIT(fm_load_options_spec_t);
  options.profile_id = profile;
  WorkbookGuard loaded;
  EXPECT_EQ(fm_workbook_load(saved.data, saved.len, &options, &loaded.handle), 0) << fm_get_last_error_message();
  return FontOf(loaded.handle, row, col);
}

void SetName(fm_cell_format_patch_t* patch, const char* name) {
  patch->font.name_op = FM_PATCH_OP_SET;
  patch->font.name = name;
}

struct RenamedFace {
  const char* face;
  int family;   // 0: no <family>
  int charset;  // 0: no <charset>
};

void ExpectRenamedFont(const char* profile, const RenamedFace& row, double normal_size) {
  SCOPED_TRACE(std::string(profile) + " -> " + row.face);
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_for_profile(profile, &wb.handle), 0);
  fm_cell_format_patch_t patch = NewPatch();
  SetName(&patch, row.face);
  ASSERT_EQ(Update(wb.handle, Cell(1, 1), patch), 0) << fm_get_last_error_message();
  const FontView font = SavedFontOf(wb.handle, profile, 1, 1);
  EXPECT_EQ(font.name, row.face);
  EXPECT_EQ(font.raw.has_family != 0, row.family != 0);
  EXPECT_EQ(font.raw.has_charset != 0, row.charset != 0);
  if (row.family != 0) {
    EXPECT_EQ(font.raw.family, row.family);
  }
  if (row.charset != 0) {
    EXPECT_EQ(font.raw.charset, row.charset);
  }
  EXPECT_EQ(font.raw.scheme, 0);
  EXPECT_DOUBLE_EQ(font.raw.size_pt, normal_size);
}

}  // namespace

// ---- font rename rules --------------------------------------------------------------------------------------------

TEST(FormulonCApiCellFormatPatch, RenamedCellFontStatesTheMeasuredFamilyAndCharsetInAJapaneseBook) {
  constexpr std::array<RenamedFace, 4> kRows = {{
      {kYuGothic, 2, 128},  // the Normal face: copies the Normal font
      {"Aptos Narrow", 0, 0},
      {kYuMincho, 1, 128},
      {"Calibri", 2, 0},
  }};
  for (const RenamedFace& row : kRows) {
    ExpectRenamedFont("mac-365-ja_JP", row, 12.0);
  }
}

TEST(FormulonCApiCellFormatPatch, RenamedCellFontStatesTheMeasuredFamilyAndCharsetInAnEnglishBook) {
  constexpr std::array<RenamedFace, 4> kRows = {{
      {kYuGothic, 0, 128},
      {"Aptos Narrow", 2, 0},  // the Normal face: copies the Normal font
      {kYuMincho, 1, 128},
      {"Calibri", 2, 0},
  }};
  for (const RenamedFace& row : kRows) {
    ExpectRenamedFont("mac-365-en_US", row, 12.0);
  }
}

TEST(FormulonCApiCellFormatPatch, TwoCellsRenamedToTheSameFaceShareOneFont) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_for_profile("mac-365-ja_JP", &wb.handle), 0);
  fm_cell_format_patch_t patch = NewPatch();
  SetName(&patch, kYuMincho);
  ASSERT_EQ(Update(wb.handle, Cell(1, 1), patch), 0);
  const size_t fonts = Count(&fm_workbook_get_font_count, wb.handle);
  ASSERT_EQ(Update(wb.handle, Cell(1, 2), patch), 0);
  EXPECT_EQ(Count(&fm_workbook_get_font_count, wb.handle), fonts);
  EXPECT_EQ(XfOf(wb.handle, 1, 1).font_index, XfOf(wb.handle, 1, 2).font_index);
}

TEST(FormulonCApiCellFormatPatch, SameFaceRenameDropsTheThemeLinkButKeepsFamilyAndCharset) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_for_profile("mac-365-ja_JP", &wb.handle), 0);
  const FontView normal = FontOf(wb.handle, 9, 9);
  ASSERT_EQ(normal.name, kYuGothic);
  ASSERT_EQ(normal.raw.scheme, FM_FONT_SCHEME_MINOR);

  fm_cell_format_patch_t patch = NewPatch();
  SetName(&patch, kYuGothic);
  ASSERT_EQ(Update(wb.handle, Cell(1, 1), patch), 0);
  const FontView renamed = FontOf(wb.handle, 1, 1);
  EXPECT_EQ(renamed.name, kYuGothic);
  EXPECT_EQ(renamed.raw.scheme, FM_FONT_SCHEME_NONE);
  EXPECT_EQ(renamed.raw.has_family, normal.raw.has_family);
  EXPECT_EQ(renamed.raw.family, normal.raw.family);
  EXPECT_EQ(renamed.raw.has_charset, normal.raw.has_charset);
  EXPECT_EQ(renamed.raw.charset, normal.raw.charset);
}

TEST(FormulonCApiCellFormatPatch, StatedSchemeBeatsTheUnlinkOfANewName) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_for_profile("mac-365-ja_JP", &wb.handle), 0);
  fm_cell_format_patch_t patch = NewPatch();
  SetName(&patch, "Arial");
  patch.font.scheme_op = FM_PATCH_OP_SET;
  patch.font.scheme = FM_FONT_SCHEME_MINOR;
  ASSERT_EQ(Update(wb.handle, Cell(1, 1), patch), 0);
  const FontView font = FontOf(wb.handle, 1, 1);
  EXPECT_EQ(font.name, "Arial");
  EXPECT_EQ(font.raw.scheme, FM_FONT_SCHEME_MINOR);
}

TEST(FormulonCApiCellFormatPatch, StatedFamilyOrCharsetSuppressesTheFaceCompletion) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_for_profile("mac-365-ja_JP", &wb.handle), 0);
  const FontView normal = FontOf(wb.handle, 9, 9);

  fm_cell_format_patch_t charset_only = NewPatch();
  SetName(&charset_only, "Arial");
  charset_only.font.charset_op = FM_PATCH_OP_SET;
  charset_only.font.charset = 5;
  ASSERT_EQ(Update(wb.handle, Cell(1, 1), charset_only), 0);
  const FontView stated = FontOf(wb.handle, 1, 1);
  EXPECT_EQ(stated.raw.has_charset, 1);
  EXPECT_EQ(stated.raw.charset, 5);
  EXPECT_EQ(stated.raw.family, normal.raw.family);  // not reset to Arial's row

  fm_cell_format_patch_t family_cleared = NewPatch();
  SetName(&family_cleared, "Arial");
  family_cleared.font.family_op = FM_PATCH_OP_CLEAR;
  ASSERT_EQ(Update(wb.handle, Cell(1, 2), family_cleared), 0);
  const FontView cleared = FontOf(wb.handle, 1, 2);
  EXPECT_EQ(cleared.raw.has_family, 0);
  EXPECT_EQ(cleared.raw.has_charset, normal.raw.has_charset);  // kept, not reset to Arial's row
  EXPECT_EQ(cleared.raw.charset, normal.raw.charset);
}

TEST(FormulonCApiCellFormatPatch, SizeOnlyPatchKeepsTheThemeLink) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_for_profile("mac-365-ja_JP", &wb.handle), 0);
  fm_cell_format_patch_t patch = NewPatch();
  patch.font.size_pt_op = FM_PATCH_OP_SET;
  patch.font.size_pt = 14.0;
  ASSERT_EQ(Update(wb.handle, Cell(1, 1), patch), 0);
  const FontView font = FontOf(wb.handle, 1, 1);
  EXPECT_DOUBLE_EQ(font.raw.size_pt, 14.0);
  EXPECT_EQ(font.raw.scheme, FM_FONT_SCHEME_MINOR);
  EXPECT_EQ(font.name, kYuGothic);
}

TEST(FormulonCApiCellFormatPatch, ClearingTheNameTakesTheNormalFaceAndLeavesTheRest) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_for_profile("mac-365-ja_JP", &wb.handle), 0);
  fm_cell_format_patch_t rename = NewPatch();
  SetName(&rename, "Arial");
  ASSERT_EQ(Update(wb.handle, Cell(1, 1), rename), 0);
  const FontView arial = FontOf(wb.handle, 1, 1);

  fm_cell_format_patch_t clear = NewPatch();
  clear.font.name_op = FM_PATCH_OP_CLEAR;
  ASSERT_EQ(Update(wb.handle, Cell(1, 1), clear), 0);
  const FontView font = FontOf(wb.handle, 1, 1);
  EXPECT_EQ(font.name, kYuGothic);
  EXPECT_EQ(font.raw.scheme, arial.raw.scheme);
  EXPECT_EQ(font.raw.has_family, arial.raw.has_family);
  EXPECT_EQ(font.raw.family, arial.raw.family);
  EXPECT_EQ(font.raw.has_charset, arial.raw.has_charset);
}

TEST(FormulonCApiCellFormatPatch, ClearingTheSizeTakesTheNormalSizeAndOtherFlagsClearToAbsent) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_for_profile("mac-365-en_US", &wb.handle), 0);
  fm_cell_format_patch_t set = NewPatch();
  set.font.size_pt_op = FM_PATCH_OP_SET;
  set.font.size_pt = 20.0;
  set.font.bold_op = FM_PATCH_OP_SET;
  set.font.bold = 1;
  set.font.italic_op = FM_PATCH_OP_SET;
  set.font.italic = 0;
  ASSERT_EQ(Update(wb.handle, Cell(0, 0), set), 0);
  const FontView before = FontOf(wb.handle, 0, 0);
  EXPECT_EQ(before.raw.has_bold, 1);
  EXPECT_EQ(before.raw.bold, 1);
  EXPECT_EQ(before.raw.has_italic, 1);  // an explicit off is a stated flag
  EXPECT_EQ(before.raw.italic, 0);

  fm_cell_format_patch_t clear = NewPatch();
  clear.font.size_pt_op = FM_PATCH_OP_CLEAR;
  clear.font.bold_op = FM_PATCH_OP_CLEAR;
  clear.font.italic_op = FM_PATCH_OP_CLEAR;
  ASSERT_EQ(Update(wb.handle, Cell(0, 0), clear), 0);
  const FontView after = FontOf(wb.handle, 0, 0);
  EXPECT_DOUBLE_EQ(after.raw.size_pt, FontOf(wb.handle, 9, 9).raw.size_pt);
  EXPECT_EQ(after.raw.has_bold, 0);
  EXPECT_EQ(after.raw.has_italic, 0);
}

TEST(FormulonCApiCellFormatPatch, EveryFontFieldMergesIntoTheCurrentFont) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_for_profile("mac-365-en_US", &wb.handle), 0);
  fm_cell_format_patch_t patch = NewPatch();
  patch.font.strike_op = FM_PATCH_OP_SET;
  patch.font.strike = 1;
  patch.font.outline_op = FM_PATCH_OP_SET;
  patch.font.outline = 1;
  patch.font.shadow_op = FM_PATCH_OP_SET;
  patch.font.shadow = 1;
  patch.font.condense_op = FM_PATCH_OP_SET;
  patch.font.condense = 1;
  patch.font.extend_op = FM_PATCH_OP_SET;
  patch.font.extend = 1;
  patch.font.underline_op = FM_PATCH_OP_SET;
  patch.font.underline = FM_UNDERLINE_STYLE_DOUBLE;
  patch.font.vert_align_op = FM_PATCH_OP_SET;
  patch.font.vert_align = FM_FONT_VERTICAL_ALIGN_SUPERSCRIPT;
  patch.font.color_op = FM_PATCH_OP_SET;
  patch.font.color = RgbColor(0xFFFF0000U);
  ASSERT_EQ(Update(wb.handle, Cell(0, 0), patch), 0);
  const FontView font = FontOf(wb.handle, 0, 0);
  EXPECT_EQ(font.raw.strike, 1);
  EXPECT_EQ(font.raw.outline, 1);
  EXPECT_EQ(font.raw.shadow, 1);
  EXPECT_EQ(font.raw.condense, 1);
  EXPECT_EQ(font.raw.extend, 1);
  EXPECT_EQ(font.raw.underline, FM_UNDERLINE_STYLE_DOUBLE);
  EXPECT_EQ(font.raw.vert_align, FM_FONT_VERTICAL_ALIGN_SUPERSCRIPT);
  EXPECT_EQ(font.raw.color.kind, FM_COLOR_KIND_RGB);
  EXPECT_EQ(font.raw.color.argb, 0xFFFF0000U);

  fm_cell_format_patch_t clear = NewPatch();
  clear.font.color_op = FM_PATCH_OP_CLEAR;
  ASSERT_EQ(Update(wb.handle, Cell(0, 0), clear), 0);
  const FontView cleared = FontOf(wb.handle, 0, 0);
  EXPECT_EQ(cleared.raw.color.kind, FM_COLOR_KIND_NONE);
  EXPECT_EQ(cleared.raw.underline, FM_UNDERLINE_STYLE_DOUBLE);  // neighbours keep
}

TEST(FormulonCApiCellFormatPatch, EachCellKeepsItsOwnOtherAttributes) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_for_profile("mac-365-en_US", &wb.handle), 0);
  fm_cell_format_patch_t bold = NewPatch();
  bold.font.bold_op = FM_PATCH_OP_SET;
  bold.font.bold = 1;
  fm_cell_format_patch_t italic = NewPatch();
  italic.font.italic_op = FM_PATCH_OP_SET;
  italic.font.italic = 1;
  ASSERT_EQ(Update(wb.handle, Cell(0, 0), bold), 0);
  ASSERT_EQ(Update(wb.handle, Cell(0, 1), italic), 0);

  fm_cell_format_patch_t size = NewPatch();
  size.font.size_pt_op = FM_PATCH_OP_SET;
  size.font.size_pt = 20.0;
  ASSERT_EQ(Update(wb.handle, fm_cell_rect_t{0, 0, 0, 1}, size), 0);
  const FontView a = FontOf(wb.handle, 0, 0);
  const FontView b = FontOf(wb.handle, 0, 1);
  EXPECT_DOUBLE_EQ(a.raw.size_pt, 20.0);
  EXPECT_DOUBLE_EQ(b.raw.size_pt, 20.0);
  EXPECT_EQ(a.raw.bold, 1);
  EXPECT_EQ(a.raw.has_italic, 0);
  EXPECT_EQ(b.raw.has_bold, 0);
  EXPECT_EQ(b.raw.italic, 1);
}

// ---- fill, border, alignment, protection, number format -----------------------------------------------------------

TEST(FormulonCApiCellFormatPatch, FillColourOnANoneFillMakesItSolid) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_cell_format_patch_t patch = NewPatch();
  patch.fill.foreground_op = FM_PATCH_OP_SET;
  patch.fill.foreground = RgbColor(0xFF00FF00U);
  ASSERT_EQ(Update(wb.handle, Cell(0, 0), patch), 0);
  fm_fill_t fill = FillOf(wb.handle, 0, 0);
  EXPECT_EQ(fill.pattern, FM_FILL_PATTERN_SOLID);
  EXPECT_EQ(fill.foreground.argb, 0xFF00FF00U);

  fm_cell_format_patch_t background = NewPatch();
  background.fill.background_op = FM_PATCH_OP_SET;
  background.fill.background = RgbColor(0xFF0000FFU);
  ASSERT_EQ(Update(wb.handle, Cell(0, 1), background), 0);
  EXPECT_EQ(FillOf(wb.handle, 0, 1).pattern, FM_FILL_PATTERN_SOLID);
}

TEST(FormulonCApiCellFormatPatch, StatedPatternBeatsTheSolidDefaultAndKeepsAnExistingPattern) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_cell_format_patch_t patch = NewPatch();
  patch.fill.pattern_op = FM_PATCH_OP_SET;
  patch.fill.pattern = FM_FILL_PATTERN_LIGHT_GRID;
  patch.fill.foreground_op = FM_PATCH_OP_SET;
  patch.fill.foreground = RgbColor(0xFF112233U);
  ASSERT_EQ(Update(wb.handle, Cell(0, 0), patch), 0);
  EXPECT_EQ(FillOf(wb.handle, 0, 0).pattern, FM_FILL_PATTERN_LIGHT_GRID);

  fm_cell_format_patch_t recolour = NewPatch();
  recolour.fill.foreground_op = FM_PATCH_OP_SET;
  recolour.fill.foreground = RgbColor(0xFF445566U);
  ASSERT_EQ(Update(wb.handle, Cell(0, 0), recolour), 0);
  const fm_fill_t fill = FillOf(wb.handle, 0, 0);
  EXPECT_EQ(fill.pattern, FM_FILL_PATTERN_LIGHT_GRID);
  EXPECT_EQ(fill.foreground.argb, 0xFF445566U);
}

TEST(FormulonCApiCellFormatPatch, AnyFillChangeDropsAGradientUnlessTheOneSetsIt) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const std::array<fm_gradient_stop_t, 2> stops = {{{0.0, RgbColor(0xFFFF0000U)}, {1.0, RgbColor(0xFF0000FFU)}}};
  fm_fill_t shape{};
  shape.gradient_type = FM_GRADIENT_TYPE_LINEAR;
  shape.gradient_angle_deg = 90.0;
  shape.gradient_stops = stops.data();
  shape.gradient_stop_count = static_cast<uint32_t>(stops.size());
  fm_cell_format_patch_t gradient = NewPatch();
  gradient.fill.gradient_op = FM_PATCH_OP_SET;
  gradient.fill.gradient = &shape;
  ASSERT_EQ(Update(wb.handle, Cell(0, 0), gradient), 0) << fm_get_last_error_message();
  fm_fill_t fill = FillOf(wb.handle, 0, 0);
  EXPECT_EQ(fill.has_gradient, 1);
  EXPECT_DOUBLE_EQ(fill.gradient_angle_deg, 90.0);
  EXPECT_EQ(fill.gradient_stop_count, 2U);

  fm_cell_format_patch_t pattern = NewPatch();
  pattern.fill.pattern_op = FM_PATCH_OP_SET;
  pattern.fill.pattern = FM_FILL_PATTERN_DARK_GRAY;
  ASSERT_EQ(Update(wb.handle, Cell(0, 0), pattern), 0);
  fill = FillOf(wb.handle, 0, 0);
  EXPECT_EQ(fill.has_gradient, 0);
  EXPECT_EQ(fill.pattern, FM_FILL_PATTERN_DARK_GRAY);

  ASSERT_EQ(Update(wb.handle, Cell(0, 0), gradient), 0);
  fm_cell_format_patch_t clear = NewPatch();
  clear.fill.gradient_op = FM_PATCH_OP_CLEAR;
  ASSERT_EQ(Update(wb.handle, Cell(0, 0), clear), 0);
  EXPECT_EQ(FillOf(wb.handle, 0, 0).has_gradient, 0);
}

TEST(FormulonCApiCellFormatPatch, BorderSidesMergeFieldByFieldAndDiagonalsAreFlags) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_cell_format_patch_t first = NewPatch();
  first.border.left.style_op = FM_PATCH_OP_SET;
  first.border.left.style = FM_BORDER_STYLE_THICK;
  first.border.left.color_op = FM_PATCH_OP_SET;
  first.border.left.color = RgbColor(0xFFFF0000U);
  first.border.diagonal_up_op = FM_PATCH_OP_SET;
  first.border.diagonal_up = 1;
  ASSERT_EQ(Update(wb.handle, Cell(0, 0), first), 0);

  fm_cell_format_patch_t second = NewPatch();
  second.border.left.style_op = FM_PATCH_OP_SET;
  second.border.left.style = FM_BORDER_STYLE_DOTTED;  // the colour stays
  second.border.bottom.style_op = FM_PATCH_OP_SET;
  second.border.bottom.style = FM_BORDER_STYLE_THIN;
  ASSERT_EQ(Update(wb.handle, Cell(0, 0), second), 0);
  const fm_border_t border = BorderOf(wb.handle, 0, 0);
  EXPECT_EQ(border.left.style, FM_BORDER_STYLE_DOTTED);
  EXPECT_EQ(border.left.color.argb, 0xFFFF0000U);
  EXPECT_EQ(border.bottom.style, FM_BORDER_STYLE_THIN);
  EXPECT_EQ(border.right.style, FM_BORDER_STYLE_NONE);
  EXPECT_EQ(border.diagonal_up, 1);
  EXPECT_EQ(border.diagonal_down, 0);
}

TEST(FormulonCApiCellFormatPatch, AlignmentAndProtectionWriteTheStatedAttributes) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_cell_format_patch_t patch = NewPatch();
  patch.alignment.horizontal_op = FM_PATCH_OP_SET;
  patch.alignment.horizontal = FM_HORIZONTAL_ALIGN_CENTER;
  patch.alignment.vertical_op = FM_PATCH_OP_SET;
  patch.alignment.vertical = FM_VERTICAL_ALIGN_TOP;
  patch.alignment.wrap_text_op = FM_PATCH_OP_SET;
  patch.alignment.wrap_text = 1;
  patch.alignment.shrink_to_fit_op = FM_PATCH_OP_SET;
  patch.alignment.shrink_to_fit = 1;
  patch.alignment.indent_op = FM_PATCH_OP_SET;
  patch.alignment.indent = 3;
  patch.alignment.text_rotation_op = FM_PATCH_OP_SET;
  patch.alignment.text_rotation = 255;
  patch.alignment.reading_order_op = FM_PATCH_OP_SET;
  patch.alignment.reading_order = FM_READING_ORDER_RIGHT_TO_LEFT;
  patch.alignment.justify_last_line_op = FM_PATCH_OP_SET;
  patch.alignment.justify_last_line = 1;
  patch.alignment.relative_indent_op = FM_PATCH_OP_SET;
  patch.alignment.relative_indent = -2;
  patch.protection.locked_op = FM_PATCH_OP_SET;
  patch.protection.locked = 0;
  ASSERT_EQ(Update(wb.handle, Cell(0, 0), patch), 0) << fm_get_last_error_message();
  const fm_cell_xf_t xf = XfOf(wb.handle, 0, 0);
  EXPECT_EQ(xf.has_alignment, 1);
  EXPECT_EQ(xf.horizontal, FM_HORIZONTAL_ALIGN_CENTER);
  EXPECT_EQ(xf.vertical, FM_VERTICAL_ALIGN_TOP);
  EXPECT_EQ(xf.wrap_text, 1);
  EXPECT_EQ(xf.shrink_to_fit, 1);
  EXPECT_EQ(xf.indent, 3U);
  EXPECT_EQ(xf.text_rotation, 255U);
  EXPECT_EQ(xf.reading_order, FM_READING_ORDER_RIGHT_TO_LEFT);
  EXPECT_EQ(xf.justify_last_line, 1);
  EXPECT_EQ(xf.relative_indent, -2);
  EXPECT_EQ(xf.apply_alignment, 1);
  EXPECT_EQ(xf.has_protection, 1);
  EXPECT_EQ(xf.locked, 0);
  EXPECT_EQ(xf.hidden, 0);
  EXPECT_EQ(xf.apply_protection, 1);

  fm_cell_format_patch_t hide = NewPatch();
  hide.protection.hidden_op = FM_PATCH_OP_SET;
  hide.protection.hidden = 1;
  ASSERT_EQ(Update(wb.handle, Cell(0, 0), hide), 0);
  const fm_cell_xf_t hidden = XfOf(wb.handle, 0, 0);
  EXPECT_EQ(hidden.locked, 0);  // the earlier flag is kept
  EXPECT_EQ(hidden.hidden, 1);
  EXPECT_EQ(hidden.horizontal, FM_HORIZONTAL_ALIGN_CENTER);
}

TEST(FormulonCApiCellFormatPatch, NumberFormatSetsAnIdAndGeneralSelectsIdZero) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_cell_format_patch_t patch = NewPatch();
  patch.number_format_op = FM_PATCH_OP_SET;
  patch.number_format = "0.000%";
  ASSERT_EQ(Update(wb.handle, Cell(0, 0), patch), 0);
  const fm_cell_xf_t xf = XfOf(wb.handle, 0, 0);
  EXPECT_NE(xf.num_fmt_id, 0U);
  EXPECT_EQ(xf.apply_number_format, 1);
  int32_t present = 0;
  const char* code = nullptr;
  ASSERT_EQ(fm_workbook_get_number_format(wb.handle, xf.num_fmt_id, &present, &code), 0);
  ASSERT_EQ(present, 1);
  EXPECT_STREQ(code, "0.000%");

  patch.number_format = "General";
  ASSERT_EQ(Update(wb.handle, Cell(0, 0), patch), 0);
  EXPECT_EQ(XfOf(wb.handle, 0, 0).num_fmt_id, 0U);
}

// ---- apply flags and no-ops ---------------------------------------------------------------------------------------

TEST(FormulonCApiCellFormatPatch, AllKeepSubPatchSetsNoApplyFlagAndAnAllKeepPatchChangesNothing) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const size_t xfs = Count(&fm_workbook_get_cell_xf_count, wb.handle);

  const fm_cell_format_patch_t empty = NewPatch();
  ASSERT_EQ(Update(wb.handle, Cell(0, 0), empty), 0);
  EXPECT_EQ(Count(&fm_workbook_get_cell_xf_count, wb.handle), xfs);
  int32_t present = 1;
  uint32_t xf_index = 0;
  ASSERT_EQ(fm_sheet_get_cell_xf_index(wb.handle, 0, 0, 0, &present, &xf_index), 0);
  EXPECT_EQ(present, 0);

  fm_cell_format_patch_t patch = NewPatch();
  patch.number_format_op = FM_PATCH_OP_SET;
  patch.number_format = "0.0";
  patch.font.name = "Ignored";  // a value without an operation states nothing
  ASSERT_EQ(Update(wb.handle, Cell(0, 0), patch), 0);
  const fm_cell_xf_t xf = XfOf(wb.handle, 0, 0);
  EXPECT_EQ(xf.apply_number_format, 1);
  EXPECT_EQ(xf.apply_font, 0);
  EXPECT_EQ(xf.apply_fill, 0);
  EXPECT_EQ(xf.apply_border, 0);
  EXPECT_EQ(xf.apply_alignment, 0);
  EXPECT_EQ(xf.apply_protection, 0);
  EXPECT_EQ(FontOf(wb.handle, 0, 0).name, FontOf(wb.handle, 9, 9).name);
}

TEST(FormulonCApiCellFormatPatch, NamedSubPatchesSetTheirApplyFlags) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_cell_format_patch_t patch = NewPatch();
  patch.font.bold_op = FM_PATCH_OP_SET;
  patch.font.bold = 1;
  patch.fill.pattern_op = FM_PATCH_OP_SET;
  patch.fill.pattern = FM_FILL_PATTERN_SOLID;
  patch.border.top.style_op = FM_PATCH_OP_SET;
  patch.border.top.style = FM_BORDER_STYLE_THIN;
  ASSERT_EQ(Update(wb.handle, Cell(0, 0), patch), 0);
  const fm_cell_xf_t xf = XfOf(wb.handle, 0, 0);
  EXPECT_EQ(xf.apply_font, 1);
  EXPECT_EQ(xf.apply_fill, 1);
  EXPECT_EQ(xf.apply_border, 1);
  EXPECT_EQ(xf.apply_alignment, 0);
  EXPECT_EQ(xf.apply_number_format, 0);
}

TEST(FormulonCApiCellFormatPatch, SecondIdenticalPatchAddsNoRecords) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_cell_format_patch_t patch = NewPatch();
  SetName(&patch, "Arial");
  patch.fill.foreground_op = FM_PATCH_OP_SET;
  patch.fill.foreground = RgbColor(0xFF123456U);
  ASSERT_EQ(Update(wb.handle, Cell(0, 0), patch), 0);
  const size_t xfs = Count(&fm_workbook_get_cell_xf_count, wb.handle);
  const size_t fonts = Count(&fm_workbook_get_font_count, wb.handle);
  const size_t fills = Count(&fm_workbook_get_fill_count, wb.handle);
  ASSERT_EQ(Update(wb.handle, Cell(0, 0), patch), 0);
  EXPECT_EQ(Count(&fm_workbook_get_cell_xf_count, wb.handle), xfs);
  EXPECT_EQ(Count(&fm_workbook_get_font_count, wb.handle), fonts);
  EXPECT_EQ(Count(&fm_workbook_get_fill_count, wb.handle), fills);
}

// ---- selections ---------------------------------------------------------------------------------------------------

TEST(FormulonCApiCellFormatPatch, WholeColumnAndRowSelectionsBecomeAxisFormatsAndAddNoCells) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_cell_format_patch_t bold = NewPatch();
  bold.font.bold_op = FM_PATCH_OP_SET;
  bold.font.bold = 1;
  ASSERT_EQ(Update(wb.handle, fm_cell_rect_t{0, 2, 1048575, 2}, bold), 0);
  EXPECT_EQ(FontOf(wb.handle, 500000, 2).raw.bold, 1);
  EXPECT_EQ(FontOf(wb.handle, 500000, 3).raw.has_bold, 0);
  int32_t present = 1;
  uint32_t xf_index = 0;
  ASSERT_EQ(fm_sheet_get_cell_xf_index(wb.handle, 0, 500000, 2, &present, &xf_index), 0);
  EXPECT_EQ(present, 0);

  fm_cell_format_patch_t italic = NewPatch();
  italic.font.italic_op = FM_PATCH_OP_SET;
  italic.font.italic = 1;
  ASSERT_EQ(Update(wb.handle, fm_cell_rect_t{4, 0, 4, 16383}, italic), 0);
  EXPECT_EQ(FontOf(wb.handle, 4, 9000).raw.italic, 1);
  EXPECT_EQ(FontOf(wb.handle, 5, 9000).raw.has_italic, 0);
  const FontView crossing = FontOf(wb.handle, 4, 2);  // the row format over the column format
  EXPECT_EQ(crossing.raw.italic, 1);
}

TEST(FormulonCApiCellFormatPatch, SelectionAboveTheCellCeilingIsRefusedAndChangesNothing) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_cell_format_patch_t bold = NewPatch();
  bold.font.bold_op = FM_PATCH_OP_SET;
  bold.font.bold = 1;
  const size_t xfs = Count(&fm_workbook_get_cell_xf_count, wb.handle);
  EXPECT_EQ(Update(wb.handle, fm_cell_rect_t{0, 0, 100000, 2000}, bold), kResourceLimit);
  EXPECT_EQ(Count(&fm_workbook_get_cell_xf_count, wb.handle), xfs);
  EXPECT_EQ(Update(wb.handle, fm_cell_rect_t{0, 0, 100000, 16382}, bold), kResourceLimit);
}

// ---- validation ---------------------------------------------------------------------------------------------------

TEST(FormulonCApiCellFormatPatch, NullAndOutOfRangeArgumentsAreRefused) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_cell_format_patch_t patch = NewPatch();
  patch.font.bold_op = FM_PATCH_OP_SET;
  const fm_cell_rect_t rect = Cell(0, 0);
  EXPECT_EQ(fm_sheet_update_cell_format(nullptr, 0, &rect, &patch), kNullPointer);
  EXPECT_EQ(fm_sheet_update_cell_format(wb.handle, 0, nullptr, &patch), kNullPointer);
  EXPECT_EQ(fm_sheet_update_cell_format(wb.handle, 0, &rect, nullptr), kNullPointer);
  EXPECT_EQ(fm_sheet_update_cell_format(wb.handle, 7, &rect, &patch), kSheetIndexOutOfRange);

  fm_cell_format_patch_t small = patch;
  small.struct_size = 4;
  EXPECT_EQ(fm_sheet_update_cell_format(wb.handle, 0, &rect, &small), kInvalidArgument);

  EXPECT_EQ(Update(wb.handle, fm_cell_rect_t{1048576, 0, 1048576, 0}, patch), kInvalidArgument);
  EXPECT_EQ(Update(wb.handle, fm_cell_rect_t{0, 16384, 0, 16384}, patch), kInvalidArgument);
  EXPECT_EQ(Update(wb.handle, fm_cell_rect_t{2, 0, 1, 0}, patch), kInvalidArgument);
}

TEST(FormulonCApiCellFormatPatch, BadOperationsAndValuesAreInvalidArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const auto refused = [&](auto&& mutate, fm_status_t expected = kInvalidArgument) {
    fm_cell_format_patch_t patch = NewPatch();
    mutate(patch);
    EXPECT_EQ(Update(wb.handle, Cell(0, 0), patch), expected) << fm_get_last_error_message();
  };
  const fm_patch_op_t bad_op = static_cast<fm_patch_op_t>(7);
  refused([&](fm_cell_format_patch_t& p) { p.font.bold_op = bad_op; });
  refused([&](fm_cell_format_patch_t& p) { p.number_format_op = FM_PATCH_OP_CLEAR; });
  refused([&](fm_cell_format_patch_t& p) {
    p.number_format_op = FM_PATCH_OP_SET;
    p.number_format = "";
  });
  refused([&](fm_cell_format_patch_t& p) {
    p.number_format_op = FM_PATCH_OP_SET;
    p.number_format = nullptr;
  });
  refused(
      [&](fm_cell_format_patch_t& p) {
        p.number_format_op = FM_PATCH_OP_SET;
        p.number_format = "\xC0\xAF";
      },
      kUtf8Error);
  refused([&](fm_cell_format_patch_t& p) { p.font.name_op = FM_PATCH_OP_SET; });
  refused([&](fm_cell_format_patch_t& p) { SetName(&p, "\xED\xA0\x80"); }, kUtf8Error);
  for (const double size :
       {0.0, 0.5, 410.0, -1.0, std::numeric_limits<double>::quiet_NaN(), std::numeric_limits<double>::infinity()}) {
    refused([&](fm_cell_format_patch_t& p) {
      p.font.size_pt_op = FM_PATCH_OP_SET;
      p.font.size_pt = size;
    });
  }
  refused([&](fm_cell_format_patch_t& p) {
    p.font.charset_op = FM_PATCH_OP_SET;
    p.font.charset = 256;
  });
  refused([&](fm_cell_format_patch_t& p) {
    p.font.charset_op = FM_PATCH_OP_SET;
    p.font.charset = -1;
  });
  refused([&](fm_cell_format_patch_t& p) {
    p.font.underline_op = FM_PATCH_OP_SET;
    p.font.underline = 5;
  });
  refused([&](fm_cell_format_patch_t& p) {
    p.font.vert_align_op = FM_PATCH_OP_SET;
    p.font.vert_align = 3;
  });
  refused([&](fm_cell_format_patch_t& p) {
    p.font.family_op = FM_PATCH_OP_SET;
    p.font.family = 6;
  });
  refused([&](fm_cell_format_patch_t& p) {
    p.font.scheme_op = FM_PATCH_OP_SET;
    p.font.scheme = 3;
  });
  refused([&](fm_cell_format_patch_t& p) {
    p.font.color_op = FM_PATCH_OP_SET;
    p.font.color.kind = 9;
  });
  refused([&](fm_cell_format_patch_t& p) {
    p.fill.pattern_op = FM_PATCH_OP_SET;
    p.fill.pattern = 19;
  });
  refused([&](fm_cell_format_patch_t& p) {
    p.fill.gradient_op = FM_PATCH_OP_SET;  // no gradient
  });
  fm_fill_t bad_type{};
  bad_type.gradient_type = 2;
  refused([&](fm_cell_format_patch_t& p) {
    p.fill.gradient_op = FM_PATCH_OP_SET;
    p.fill.gradient = &bad_type;
  });
  fm_fill_t no_array{};
  no_array.gradient_stop_count = 2;
  refused([&](fm_cell_format_patch_t& p) {
    p.fill.gradient_op = FM_PATCH_OP_SET;
    p.fill.gradient = &no_array;
  });
  refused([&](fm_cell_format_patch_t& p) {
    p.border.diagonal.style_op = FM_PATCH_OP_SET;
    p.border.diagonal.style = 14;
  });
  refused([&](fm_cell_format_patch_t& p) {
    p.border.left.color_op = FM_PATCH_OP_SET;
    p.border.left.color.kind = -1;
  });
  refused([&](fm_cell_format_patch_t& p) {
    p.alignment.horizontal_op = FM_PATCH_OP_SET;
    p.alignment.horizontal = 8;
  });
  refused([&](fm_cell_format_patch_t& p) {
    p.alignment.vertical_op = FM_PATCH_OP_SET;
    p.alignment.vertical = 5;
  });
  refused([&](fm_cell_format_patch_t& p) {
    p.alignment.reading_order_op = FM_PATCH_OP_SET;
    p.alignment.reading_order = 3;
  });
  refused([&](fm_cell_format_patch_t& p) {
    p.alignment.indent_op = FM_PATCH_OP_SET;
    p.alignment.indent = 256;
  });
  refused([&](fm_cell_format_patch_t& p) {
    p.alignment.text_rotation_op = FM_PATCH_OP_SET;
    p.alignment.text_rotation = 181;
  });
}

TEST(FormulonCApiCellFormatPatch, GradientAboveTheStopLimitIsAResourceLimit) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const std::vector<fm_gradient_stop_t> stops(1025);
  fm_fill_t shape{};
  shape.gradient_stops = stops.data();
  shape.gradient_stop_count = static_cast<uint32_t>(stops.size());
  fm_cell_format_patch_t patch = NewPatch();
  patch.fill.gradient_op = FM_PATCH_OP_SET;
  patch.fill.gradient = &shape;
  EXPECT_EQ(Update(wb.handle, Cell(0, 0), patch), kResourceLimit);
}

TEST(FormulonCApiCellFormatPatch, ClearIsAllowedOnlyOnTheClearableFields) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  // Clearable: the ledger's `font_patch` and `fill_patch` lists.
  const std::array<void (*)(fm_font_patch_t*), 12> clearable_font = {{
      [](fm_font_patch_t* p) { p->name_op = FM_PATCH_OP_CLEAR; },
      [](fm_font_patch_t* p) { p->size_pt_op = FM_PATCH_OP_CLEAR; },
      [](fm_font_patch_t* p) { p->bold_op = FM_PATCH_OP_CLEAR; },
      [](fm_font_patch_t* p) { p->italic_op = FM_PATCH_OP_CLEAR; },
      [](fm_font_patch_t* p) { p->strike_op = FM_PATCH_OP_CLEAR; },
      [](fm_font_patch_t* p) { p->outline_op = FM_PATCH_OP_CLEAR; },
      [](fm_font_patch_t* p) { p->shadow_op = FM_PATCH_OP_CLEAR; },
      [](fm_font_patch_t* p) { p->condense_op = FM_PATCH_OP_CLEAR; },
      [](fm_font_patch_t* p) { p->extend_op = FM_PATCH_OP_CLEAR; },
      [](fm_font_patch_t* p) { p->family_op = FM_PATCH_OP_CLEAR; },
      [](fm_font_patch_t* p) { p->charset_op = FM_PATCH_OP_CLEAR; },
      [](fm_font_patch_t* p) { p->color_op = FM_PATCH_OP_CLEAR; },
  }};
  for (const auto clear : clearable_font) {
    fm_cell_format_patch_t patch = NewPatch();
    clear(&patch.font);
    EXPECT_EQ(Update(wb.handle, Cell(0, 0), patch), 0) << fm_get_last_error_message();
  }
  fm_cell_format_patch_t gradient = NewPatch();
  gradient.fill.gradient_op = FM_PATCH_OP_CLEAR;
  EXPECT_EQ(Update(wb.handle, Cell(0, 0), gradient), 0);

  const std::array<void (*)(fm_cell_format_patch_t*), 28> not_clearable = {{
      [](fm_cell_format_patch_t* p) { p->font.underline_op = FM_PATCH_OP_CLEAR; },
      [](fm_cell_format_patch_t* p) { p->font.vert_align_op = FM_PATCH_OP_CLEAR; },
      [](fm_cell_format_patch_t* p) { p->font.scheme_op = FM_PATCH_OP_CLEAR; },
      [](fm_cell_format_patch_t* p) { p->fill.pattern_op = FM_PATCH_OP_CLEAR; },
      [](fm_cell_format_patch_t* p) { p->fill.foreground_op = FM_PATCH_OP_CLEAR; },
      [](fm_cell_format_patch_t* p) { p->fill.background_op = FM_PATCH_OP_CLEAR; },
      [](fm_cell_format_patch_t* p) { p->border.left.style_op = FM_PATCH_OP_CLEAR; },
      [](fm_cell_format_patch_t* p) { p->border.left.color_op = FM_PATCH_OP_CLEAR; },
      [](fm_cell_format_patch_t* p) { p->border.right.style_op = FM_PATCH_OP_CLEAR; },
      [](fm_cell_format_patch_t* p) { p->border.right.color_op = FM_PATCH_OP_CLEAR; },
      [](fm_cell_format_patch_t* p) { p->border.top.style_op = FM_PATCH_OP_CLEAR; },
      [](fm_cell_format_patch_t* p) { p->border.top.color_op = FM_PATCH_OP_CLEAR; },
      [](fm_cell_format_patch_t* p) { p->border.bottom.style_op = FM_PATCH_OP_CLEAR; },
      [](fm_cell_format_patch_t* p) { p->border.bottom.color_op = FM_PATCH_OP_CLEAR; },
      [](fm_cell_format_patch_t* p) { p->border.diagonal.style_op = FM_PATCH_OP_CLEAR; },
      [](fm_cell_format_patch_t* p) { p->border.diagonal.color_op = FM_PATCH_OP_CLEAR; },
      [](fm_cell_format_patch_t* p) { p->border.diagonal_up_op = FM_PATCH_OP_CLEAR; },
      [](fm_cell_format_patch_t* p) { p->border.diagonal_down_op = FM_PATCH_OP_CLEAR; },
      [](fm_cell_format_patch_t* p) { p->alignment.horizontal_op = FM_PATCH_OP_CLEAR; },
      [](fm_cell_format_patch_t* p) { p->alignment.vertical_op = FM_PATCH_OP_CLEAR; },
      [](fm_cell_format_patch_t* p) { p->alignment.wrap_text_op = FM_PATCH_OP_CLEAR; },
      [](fm_cell_format_patch_t* p) { p->alignment.shrink_to_fit_op = FM_PATCH_OP_CLEAR; },
      [](fm_cell_format_patch_t* p) { p->alignment.indent_op = FM_PATCH_OP_CLEAR; },
      [](fm_cell_format_patch_t* p) { p->alignment.text_rotation_op = FM_PATCH_OP_CLEAR; },
      [](fm_cell_format_patch_t* p) { p->alignment.reading_order_op = FM_PATCH_OP_CLEAR; },
      [](fm_cell_format_patch_t* p) { p->alignment.justify_last_line_op = FM_PATCH_OP_CLEAR; },
      [](fm_cell_format_patch_t* p) { p->alignment.relative_indent_op = FM_PATCH_OP_CLEAR; },
      [](fm_cell_format_patch_t* p) { p->protection.locked_op = FM_PATCH_OP_CLEAR; },
  }};
  for (const auto clear : not_clearable) {
    fm_cell_format_patch_t patch = NewPatch();
    clear(&patch);
    EXPECT_EQ(Update(wb.handle, Cell(0, 0), patch), kInvalidArgument);
  }
  fm_cell_format_patch_t hidden = NewPatch();
  hidden.protection.hidden_op = FM_PATCH_OP_CLEAR;
  EXPECT_EQ(Update(wb.handle, Cell(0, 0), hidden), kInvalidArgument);
}

TEST(FormulonCApiCellFormatPatch, InvalidPatchLeavesTheWorkbookUnchanged) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_cell_format_patch_t patch = NewPatch();
  patch.number_format_op = FM_PATCH_OP_SET;
  patch.number_format = "0.0000";
  SetName(&patch, "Arial");
  patch.fill.foreground_op = FM_PATCH_OP_SET;
  patch.fill.foreground = RgbColor(0xFF010203U);
  patch.alignment.horizontal_op = FM_PATCH_OP_SET;
  patch.alignment.horizontal = 99;  // refused last
  const size_t xfs = Count(&fm_workbook_get_cell_xf_count, wb.handle);
  const size_t fonts = Count(&fm_workbook_get_font_count, wb.handle);
  const size_t fills = Count(&fm_workbook_get_fill_count, wb.handle);
  const std::string styles_before = formulon::io::write_styles(wb.handle->workbook().styles());

  EXPECT_EQ(Update(wb.handle, Cell(0, 0), patch), kInvalidArgument);
  EXPECT_EQ(Count(&fm_workbook_get_cell_xf_count, wb.handle), xfs);
  EXPECT_EQ(Count(&fm_workbook_get_font_count, wb.handle), fonts);
  EXPECT_EQ(Count(&fm_workbook_get_fill_count, wb.handle), fills);
  EXPECT_EQ(formulon::io::write_styles(wb.handle->workbook().styles()), styles_before);
  int32_t present = 1;
  uint32_t xf_index = 0;
  ASSERT_EQ(fm_sheet_get_cell_xf_index(wb.handle, 0, 0, 0, &present, &xf_index), 0);
  EXPECT_EQ(present, 0);
}
