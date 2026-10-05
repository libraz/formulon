#include "color_resolve.h"

#include <cstdint>
#include <cstdio>
#include <string>

#include "gtest/gtest.h"
#include "theme.h"
#include "workbook.h"

namespace formulon {
namespace {

// Theme colour scheme the table was measured against (clrScheme order).
constexpr std::uint32_t kMeasuredScheme[12] = {0xFF000000U, 0xFFFFFFFFU, 0xFF44546AU, 0xFFE7E6E6U,
                                               0xFF4472C4U, 0xFFED7D31U, 0xFFA5A5A5U, 0xFFFFC000U,
                                               0xFF5B9BD5U, 0xFF70AD47U, 0xFF0563C1U, 0xFF954F72U};

// Columns of kTintTable.
constexpr double kTints[9] = {-0.5, -0.25, -0.1, 0.0, 0.1, 0.25, 0.4, 0.6, 0.8};

// Resolved RGB by `theme` attribute index (0=lt1, 1=dk1, 2=lt2, 3=dk2, 4..11) and tint.
constexpr const char* kTintTable[12][9] = {
    {"808080", "BFBFBF", "E6E6E6", "FFFFFF", "FFFFFF", "FFFFFF", "FFFFFF", "FFFFFF", "FFFFFF"},
    {"000000", "000000", "000000", "000000", "1A1A1A", "404040", "666666", "999999", "CCCCCC"},
    {"757171", "AEAAAA", "D0CECE", "E7E6E6", "E9E9E9", "ECECEC", "F0F0F0", "F4F4F4", "FAFAFA"},
    {"222B35", "333F4F", "3D4B5F", "44546A", "51647D", "657C9C", "8497B0", "ACB9CA", "D6DCE4"},
    {"203764", "305496", "3864B4", "4472C4", "557ECA", "7395D2", "8EA9DB", "B4C6E7", "D9E1F2"},
    {"833C0C", "C65911", "EB6B16", "ED7D31", "EF8945", "F19D65", "F4B084", "F8CBAD", "FCE4D6"},
    {"525252", "7B7B7B", "949494", "A5A5A5", "ADADAD", "BBBBBB", "C9C9C9", "DBDBDB", "EDEDED"},
    {"806000", "BF8F00", "E6AC00", "FFC000", "FFC61A", "FFCF40", "FFD966", "FFE699", "FFF2CC"},
    {"1F4E78", "2F75B5", "428BCE", "5B9BD5", "6AA5D9", "84B4DF", "9BC2E6", "BDD7EE", "DDEBF7"},
    {"375623", "548235", "649B40", "70AD47", "7DB955", "93C572", "A9D08E", "C6E0B4", "E2EFDA"},
    {"023160", "03498F", "0458AC", "0563C1", "0572DE", "1A89F9", "46A0FB", "85C0FC", "C1DFFD"},
    {"492738", "703A55", "864666", "954F72", "A75880", "B67495", "C58FAA", "D8B4C6", "EBDAE2"},
};

// Indexed colour by palette index: fill foreground and font colour.
constexpr const char* kIndexedFill[66] = {
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", "000000", "FFFFFF", "FF0000",
    "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", "800000", "008000", "000080", "808000", "800080", "008080",
    "C0C0C0", "808080", "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF", "000080",
    "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF", "00CCFF", "CCFFFF", "CCFFCC", "FFFF99",
    "99CCFF", "FF99CC", "CC99FF", "FFCC99", "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699",
    "969696", "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333", "FFFFFF", "FFFFFF",
};
constexpr const char* kIndexedFont[66] = {
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", "000000", "FFFFFF", "FF0000",
    "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", "800000", "008000", "000080", "808000", "800080", "008080",
    "C0C0C0", "808080", "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF", "000080",
    "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF", "00CCFF", "CCFFFF", "CCFFCC", "FFFF99",
    "99CCFF", "FF99CC", "CC99FF", "FFCC99", "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699",
    "969696", "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333", "000000", "FFFFFF",
};

std::string Hex(std::uint32_t argb) {
  char buf[12];
  std::snprintf(buf, sizeof(buf), "%06X", argb & 0xFFFFFFU);
  return buf;
}

Theme MeasuredTheme() {
  Theme theme = default_theme();
  for (std::size_t i = 0; i < kThemeColorCount; ++i) {
    theme.colors[i] = kMeasuredScheme[i];
  }
  return theme;
}

ColorSpec ThemeSpec(std::uint32_t index, double tint) {
  ColorSpec spec;
  spec.kind = ColorSpec::Kind::kTheme;
  spec.theme = index;
  spec.tint = tint;
  return spec;
}

ColorSpec IndexedSpec(std::uint32_t index) {
  ColorSpec spec;
  spec.kind = ColorSpec::Kind::kIndexed;
  spec.indexed = index;
  return spec;
}

TEST(ColorResolve, ProbeTable) {
  const Theme theme = MeasuredTheme();
  const IndexedPalette palette;
  int theme_checked = 0;
  for (std::uint32_t index = 0; index < 12U; ++index) {
    for (int t = 0; t < 9; ++t) {
      const ResolvedColor got =
          resolve_color(ThemeSpec(index, kTints[t]), theme, ThemeSource::kPart, palette, ColorContext::kFillForeground);
      EXPECT_EQ(Hex(got.argb), kTintTable[index][t]) << "theme=" << index << " tint=" << kTints[t];
      EXPECT_EQ(got.argb >> 24U, 0xFFU);
      EXPECT_EQ(got.resolution, ColorResolution::kExact);
      ++theme_checked;
    }
  }
  EXPECT_EQ(theme_checked, 108);
  for (std::uint32_t index = 0; index < 66U; ++index) {
    const ResolvedColor fill =
        resolve_color(IndexedSpec(index), theme, ThemeSource::kPart, palette, ColorContext::kFillForeground);
    const ResolvedColor font =
        resolve_color(IndexedSpec(index), theme, ThemeSource::kPart, palette, ColorContext::kFont);
    EXPECT_EQ(Hex(fill.argb), kIndexedFill[index]) << "indexed fill " << index;
    EXPECT_EQ(Hex(font.argb), kIndexedFont[index]) << "indexed font " << index;
  }
}

TEST(ColorResolve, ThemeFontColourWithoutTint) {
  const Theme theme = MeasuredTheme();
  const char* expected[4] = {"FFFFFF", "000000", "E7E6E6", "44546A"};
  for (std::uint32_t index = 0; index < 4U; ++index) {
    EXPECT_EQ(Hex(resolve_color(ThemeSpec(index, 0.0), theme, ThemeSource::kPart, {}, ColorContext::kFont).argb),
              expected[index]);
  }
}

TEST(ColorResolve, ResolutionReflectsThemeSource) {
  const Theme theme = default_theme();
  EXPECT_EQ(resolve_color(ThemeSpec(4, 0.0), theme, ThemeSource::kPart, {}, ColorContext::kFont).resolution,
            ColorResolution::kExact);
  EXPECT_EQ(resolve_color(ThemeSpec(4, 0.0), theme, ThemeSource::kDefault, {}, ColorContext::kFont).resolution,
            ColorResolution::kDefaultTheme);
  EXPECT_EQ(resolve_color(ThemeSpec(4, 0.0), theme, ThemeSource::kUnparseable, {}, ColorContext::kFont).resolution,
            ColorResolution::kThemeUnparseable);
  EXPECT_EQ(resolve_color(ThemeSpec(12, 0.0), theme, ThemeSource::kPart, {}, ColorContext::kFont).resolution,
            ColorResolution::kIndexOutOfRange);
}

TEST(ColorResolve, IndexedBeyondPaletteIsOutOfRange) {
  const ResolvedColor got =
      resolve_color(IndexedSpec(66), default_theme(), ThemeSource::kPart, {}, ColorContext::kFont);
  EXPECT_EQ(got.resolution, ColorResolution::kIndexOutOfRange);
  EXPECT_EQ(got.argb, 0xFF000000U);
}

TEST(ColorResolve, SystemIndexesAndAutoFollowContext) {
  const Theme theme = default_theme();
  ColorSpec auto_spec;
  auto_spec.kind = ColorSpec::Kind::kAuto;
  for (const ColorSpec& spec : {IndexedSpec(64), auto_spec, ColorSpec{}}) {
    const ResolvedColor font = resolve_color(spec, theme, ThemeSource::kPart, {}, ColorContext::kFont);
    const ResolvedColor border = resolve_color(spec, theme, ThemeSource::kPart, {}, ColorContext::kBorder);
    const ResolvedColor fill = resolve_color(spec, theme, ThemeSource::kPart, {}, ColorContext::kFillBackground);
    EXPECT_EQ(font.argb, 0xFF000000U);
    EXPECT_EQ(border.argb, 0xFF000000U);
    EXPECT_EQ(fill.argb, 0xFFFFFFFFU);
    EXPECT_EQ(font.resolution, ColorResolution::kAutoContext);
  }
}

TEST(ColorResolve, PaletteOverrideReplacesLeadingEntries) {
  const IndexedPalette palette = {0xFF112233U, 0x00445566U};
  const Theme theme = default_theme();
  EXPECT_EQ(resolve_color(IndexedSpec(0), theme, ThemeSource::kPart, palette, ColorContext::kFont).argb, 0xFF112233U);
  EXPECT_EQ(resolve_color(IndexedSpec(1), theme, ThemeSource::kPart, palette, ColorContext::kFont).argb, 0xFF445566U);
  EXPECT_EQ(resolve_color(IndexedSpec(2), theme, ThemeSource::kPart, palette, ColorContext::kFont).argb, 0xFFFF0000U);
}

TEST(ColorResolve, LiteralRgbIsExact) {
  ColorSpec spec;
  spec.kind = ColorSpec::Kind::kRgb;
  spec.rgb = 0xFF123456U;
  const ResolvedColor got = resolve_color(spec, default_theme(), ThemeSource::kDefault, {}, ColorContext::kFont);
  EXPECT_EQ(got.argb, 0xFF123456U);
  EXPECT_EQ(got.resolution, ColorResolution::kExact);
}

TEST(ColorResolve, WorkbookOverloadUsesWorkbookTheme) {
  Workbook wb = Workbook::create();
  EXPECT_EQ(resolve_color(wb, ThemeSpec(4, 0.0), ColorContext::kFont).resolution, ColorResolution::kDefaultTheme);
  ThemeColors colors = default_theme().colors;
  colors[4] = 0xFF010203U;
  ASSERT_TRUE(static_cast<bool>(set_theme_colors(wb, colors)));
  const ResolvedColor got = resolve_color(wb, ThemeSpec(4, 0.0), ColorContext::kFont);
  EXPECT_EQ(got.argb, 0xFF010203U);
  EXPECT_EQ(got.resolution, ColorResolution::kExact);
}

}  // namespace
}  // namespace formulon
