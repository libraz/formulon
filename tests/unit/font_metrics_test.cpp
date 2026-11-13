// Normal-font metrics per host: the Mac and Windows tables, name matching, the
// theme-minor face of a Mac Normal font, and the column-width model built on them.

#include "font_metrics.h"

#include <cstddef>
#include <cstdint>
#include <iterator>
#include <string>
#include <vector>

#include "column_width.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "styles.h"
#include "theme.h"

namespace formulon {
namespace {

constexpr double kWinRowHeight = 102.0 / 7.0;

StylesTable StylesWith(const std::string& name, double size, std::uint8_t scheme = 0) {
  StylesTable styles;
  FontRecord font;
  font.name = name;
  font.size = size;
  font.scheme = scheme;
  styles.fonts.push_back(font);
  return styles;
}

NormalFontMetrics Metrics(const std::string& name, double size, ExcelProfile profile) {
  return normal_font_metrics(StylesWith(name, size), default_theme(), profile);
}

struct MacRow {
  const char* family;
  double size;
  double mdw;
  double pad;
  double row_height;
};

constexpr MacRow kMacRows[] = {
    {"Calibri", 11, 6, 5, 15},         {"Calibri", 14, 7, 5, 19},      {"Calibri", 20, 10, 7, 26},
    {"Aptos Narrow", 11, 6, 5, 15},    {"Aptos Narrow", 12, 6, 5, 16}, {"游ゴシック", 11, 6, 5, 18},
    {"游ゴシック", 12, 7, 5, 20},      {"等线", 12, 6, 5, 16},         {"맑은 고딕", 12, 7, 5, 18},
    {"Tahoma", 12, 7, 5, 15},          {"新細明體", 12, 6, 5, 15},     {"ＭＳ Ｐゴシック", 11, 6, 5, 14},
    {"Calibri", 12, 6, 5, 16},         {"Meiryo", 11, 7, 5, 19},       {"Aptos", 11, 6, 5, 15},
    {"Aptos", 12, 6, 5, 16},           {"Arial", 10, 6, 5, 13},        {"Arial", 11, 6, 5, 14},
    {"Arial", 12, 7, 5, 16},           {"Calibri", 10, 5, 5, 14},      {"Cambria", 11, 6, 5, 14},
    {"Times New Roman", 12, 6, 5, 16}, {"Verdana", 10, 6, 5, 13},      {"Helvetica", 12, 7, 5, 16},
    {"Helvetica Neue", 12, 7, 5, 16},  {"Meiryo", 10, 6, 5, 17},       {"Meiryo", 12, 7, 5, 20},
    {"Meiryo UI", 9, 6, 5, 14},        {"Meiryo UI", 11, 7, 5, 16},    {"ＭＳ ゴシック", 11, 6, 5, 14},
    {"ＭＳ 明朝", 11, 6, 5, 14},       {"游明朝", 11, 6, 5, 18},       {"メイリオ", 11, 7, 5, 19},
    {"游ゴシック", 10, 6, 5, 17},      {"游ゴシック", 14, 8, 5, 24},   {"Aptos Narrow", 10, 5, 5, 14},
    {"Aptos Narrow", 14, 7, 5, 19},    {"等线", 11, 6, 5, 15},         {"맑은 고딕", 11, 6, 5, 17},
    {"Tahoma", 10, 5, 5, 13},          {"Tahoma", 11, 6, 5, 14},       {"Courier New", 10, 6, 5, 14},
    {"ＭＳ Ｐゴシック", 12, 6, 5, 15}, {"新細明體", 11, 5, 5, 14},
};

TEST(FontMetrics, MacTableRowsAreCalibratedWholePoints) {
  for (const MacRow& row : kMacRows) {
    const NormalFontMetrics m = Metrics(row.family, row.size, mac_365_en_us_profile());
    EXPECT_EQ(m.host, ExcelHost::kMac365) << row.family;
    EXPECT_EQ(m.family, row.family);
    EXPECT_DOUBLE_EQ(m.size_pt, row.size) << row.family;
    EXPECT_DOUBLE_EQ(m.mdw_pt, row.mdw) << row.family << row.size;
    EXPECT_DOUBLE_EQ(m.padding_pt, row.pad) << row.family << row.size;
    EXPECT_DOUBLE_EQ(m.row_height_pt, row.row_height) << row.family << row.size;
    EXPECT_TRUE(m.width_calibrated) << row.family << row.size;
    EXPECT_TRUE(m.row_height_calibrated) << row.family << row.size;
  }
}

TEST(FontMetrics, WindowsTableIsUnchangedAndNeverRowCalibrated) {
  struct Case {
    const char* family;
    double size;
    double mdw;
    double pad;
  };
  const Case cases[] = {{"Calibri", 8, 4.5, 3.75},      {"Calibri", 11, 5.25, 3.75},
                        {"Calibri", 18, 9.0, 5.25},     {"ＭＳ Ｐゴシック", 11, 6.0, 3.75},
                        {"游ゴシック", 12, 6.75, 5.25}, {"游ゴシック", 18, 9.75, 6.75},
                        {"Meiryo UI", 11, 6.75, 5.25},  {"Meiryo UI", 18, 11.25, 6.75}};
  for (const Case& c : cases) {
    const NormalFontMetrics m = Metrics(c.family, c.size, win_365_ja_jp_profile());
    EXPECT_EQ(m.host, ExcelHost::kWin365);
    EXPECT_DOUBLE_EQ(m.mdw_pt, c.mdw) << c.family << c.size;
    EXPECT_DOUBLE_EQ(m.padding_pt, c.pad) << c.family << c.size;
    EXPECT_TRUE(m.width_calibrated) << c.family << c.size;
    EXPECT_FALSE(m.row_height_calibrated) << c.family << c.size;
    EXPECT_DOUBLE_EQ(m.row_height_pt, kWinRowHeight) << c.family << c.size;
  }
}

TEST(FontMetrics, UnknownFontOrSizeFallsBackToTheHostCalibri11) {
  const NormalFontMetrics mac = Metrics("Garamond", 10, mac_365_ja_jp_profile());
  EXPECT_DOUBLE_EQ(mac.mdw_pt, 6.0);
  EXPECT_DOUBLE_EQ(mac.padding_pt, 5.0);
  EXPECT_DOUBLE_EQ(mac.row_height_pt, 15.0);
  EXPECT_FALSE(mac.width_calibrated);
  EXPECT_FALSE(mac.row_height_calibrated);
  EXPECT_EQ(mac.family, "Garamond");

  const NormalFontMetrics mac_size = Metrics("Calibri", 13, mac_365_ja_jp_profile());
  EXPECT_FALSE(mac_size.width_calibrated);

  const NormalFontMetrics win = Metrics("Arial", 10, win_365_en_us_profile());
  EXPECT_DOUBLE_EQ(win.mdw_pt, 5.25);
  EXPECT_DOUBLE_EQ(win.padding_pt, 3.75);
  EXPECT_FALSE(win.width_calibrated);
  EXPECT_FALSE(win.row_height_calibrated);
}

TEST(FontMetrics, EmptyFontTableYieldsTheUncalibratedCalibri11) {
  const StylesTable empty;
  const NormalFontMetrics mac = normal_font_metrics(empty, default_theme(), mac_365_en_us_profile());
  EXPECT_DOUBLE_EQ(mac.mdw_pt, 6.0);
  EXPECT_FALSE(mac.width_calibrated);
  EXPECT_EQ(mac.family, "");
  EXPECT_DOUBLE_EQ(mac.size_pt, 0.0);
  const NormalFontMetrics win = normal_font_metrics(empty, default_theme(), win_365_en_us_profile());
  EXPECT_DOUBLE_EQ(win.mdw_pt, 5.25);
  EXPECT_FALSE(win.width_calibrated);
}

TEST(FontMetrics, MacEnglishAliasesResolveAndYuGothicDoesNot) {
  struct Case {
    const char* alias;
    double size;
    double mdw;
    double row_height;
  };
  const Case cases[] = {
      {"MS PGothic", 11, 6, 14}, {"DengXian", 12, 6, 16}, {"Malgun Gothic", 12, 7, 18}, {"PMingLiU", 12, 6, 15}};
  for (const Case& c : cases) {
    const NormalFontMetrics m = Metrics(c.alias, c.size, mac_365_ja_jp_profile());
    EXPECT_TRUE(m.width_calibrated) << c.alias;
    EXPECT_DOUBLE_EQ(m.mdw_pt, c.mdw) << c.alias;
    EXPECT_DOUBLE_EQ(m.row_height_pt, c.row_height) << c.alias;
    EXPECT_EQ(m.family, c.alias);
    // Windows matches names exactly: an alias is not in its table.
    EXPECT_FALSE(Metrics(c.alias, c.size, win_365_ja_jp_profile()).width_calibrated) << c.alias;
  }
  const NormalFontMetrics yu = Metrics("Yu Gothic", 11, mac_365_ja_jp_profile());
  EXPECT_FALSE(yu.width_calibrated);
  EXPECT_FALSE(yu.row_height_calibrated);
  EXPECT_DOUBLE_EQ(yu.row_height_pt, 15.0);
}

// Model of the scheme=minor face test: host (mac, win) x locale (7) x scheme (none, minor) x
// theme (Office default, Jpan overridden). The four even-parity (scheme, theme, host) triples form an
// orthogonal array over every pair of the three binary parameters; each locale takes one triple and its
// complement, so every locale sees both values of each binary parameter.
struct SchemeCase {
  ExcelHost host;
  ExcelLocale locale;
  bool minor;
  bool jpan_override;
};

const char* ExpectedMinorFace(ExcelLocale locale, bool jpan_override) {
  switch (locale) {
    case ExcelLocale::kJaJP:
      return jpan_override ? "ＭＳ Ｐゴシック" : "游ゴシック";
    case ExcelLocale::kZhCN:
      return "等线";
    case ExcelLocale::kKoKR:
      return "맑은 고딕";
    case ExcelLocale::kThTH:
      return "Tahoma";
    case ExcelLocale::kZhTW:
      return "新細明體";
    default:
      return "Aptos Narrow";
  }
}

Theme JpanOverriddenTheme() {
  Theme theme = default_theme();
  theme.minor_scripts.push_back({"Jpan", "ＭＳ Ｐゴシック"});
  return theme;
}

TEST(FontMetrics, MinorSchemeFaceFollowsHostLocaleAndTheme) {
  const ExcelLocale locales[] = {ExcelLocale::kJaJP, ExcelLocale::kEnUS, ExcelLocale::kZhCN, ExcelLocale::kKoKR,
                                 ExcelLocale::kThTH, ExcelLocale::kZhTW, ExcelLocale::kDeDE};
  struct Triple {
    bool minor;
    bool jpan_override;
    bool mac;
  };
  const Triple even[] = {{false, false, false}, {false, true, true}, {true, false, true}, {true, true, false}};
  std::vector<SchemeCase> cases;
  for (std::size_t i = 0; i < std::size(locales); ++i) {
    const Triple a = even[i % 4];
    for (const Triple& t : {a, Triple{!a.minor, !a.jpan_override, !a.mac}}) {
      cases.push_back({t.mac ? ExcelHost::kMac365 : ExcelHost::kWin365, locales[i], t.minor, t.jpan_override});
    }
  }
  ASSERT_EQ(cases.size(), 14U);
  const Theme office = default_theme();
  const Theme overridden = JpanOverriddenTheme();
  for (const SchemeCase& c : cases) {
    const StylesTable styles = StylesWith("Calibri", 12, c.minor ? kFontSchemeMinor : 0);
    const NormalFontMetrics m =
        normal_font_metrics(styles, c.jpan_override ? overridden : office, ExcelProfile{c.host, c.locale});
    const bool follows_theme = c.host == ExcelHost::kMac365 && c.minor;
    EXPECT_EQ(m.family, follows_theme ? ExpectedMinorFace(c.locale, c.jpan_override) : "Calibri")
        << static_cast<int>(c.host) << ' ' << static_cast<int>(c.locale) << ' ' << c.minor << ' ' << c.jpan_override;
    EXPECT_DOUBLE_EQ(m.size_pt, 12.0);
  }
}

TEST(FontMetrics, MinorSchemeRowHeightsOfTheRowHeightProbes) {
  const StylesTable styles = StylesWith("Calibri", 11, kFontSchemeMinor);
  const NormalFontMetrics ja = normal_font_metrics(styles, default_theme(), mac_365_ja_jp_profile());
  EXPECT_EQ(ja.family, "游ゴシック");
  EXPECT_DOUBLE_EQ(ja.row_height_pt, 18.0);
  const NormalFontMetrics ja_theme = normal_font_metrics(styles, JpanOverriddenTheme(), mac_365_ja_jp_profile());
  EXPECT_EQ(ja_theme.family, "ＭＳ Ｐゴシック");
  EXPECT_DOUBLE_EQ(ja_theme.row_height_pt, 14.0);
  const NormalFontMetrics en = normal_font_metrics(styles, default_theme(), mac_365_en_us_profile());
  EXPECT_EQ(en.family, "Aptos Narrow");
  EXPECT_DOUBLE_EQ(en.row_height_pt, 15.0);
}

TEST(FontMetrics, EmptyThemeEntryTakesTheLocaleNewWorkbookFont) {
  Theme theme;
  const StylesTable styles = StylesWith("Calibri", 12, kFontSchemeMinor);
  EXPECT_EQ(normal_font_metrics(styles, theme, mac_365_ja_jp_profile()).family, "游ゴシック");
  EXPECT_EQ(normal_font_metrics(styles, theme, mac_365_en_us_profile()).family, "Aptos Narrow");
  theme.minor_scripts.push_back({"Jpan", ""});
  EXPECT_EQ(normal_font_metrics(styles, theme, mac_365_ja_jp_profile()).family, "游ゴシック");
}

ColumnWidthModel MacModel(const std::string& name, double size, GeometryMode mode = GeometryMode::kDisplay) {
  return resolve_column_width_model(StylesWith(name, size), default_theme(), mac_365_en_us_profile(), mode);
}

TEST(FontMetrics, MacStoredWidthsSnapToWholePoints) {
  struct Case {
    const char* family;
    double size;
    double base10_default;
    double default_pt;
    double ten_pt;
  };
  const Case cases[] = {{"Calibri", 11, 10.0 + 213.0 / 256.0, 65, 60},
                        {"Calibri", 14, 10.0 + 182.0 / 256.0, 75, 70},
                        {"Calibri", 20, 10.0 + 179.0 / 256.0, 107, 100}};
  for (const Case& c : cases) {
    const ColumnWidthModel model = MacModel(c.family, c.size);
    EXPECT_EQ(model.host, ExcelHost::kMac365);
    EXPECT_STREQ(model.platform, "mac");
    EXPECT_TRUE(model.calibrated);
    EXPECT_DOUBLE_EQ(column_chars_to_points(c.base10_default, model), c.default_pt) << c.size;
    EXPECT_DOUBLE_EQ(column_chars_to_points(10.0, model), c.ten_pt) << c.size;
    EXPECT_DOUBLE_EQ(column_points_to_chars(c.ten_pt, model), 10.0) << c.size;
  }
}

TEST(FontMetrics, MacDefaultStoredWidthFollowsBaseAndMdw) {
  struct Case {
    const char* family;
    double size;
    bool has_row_height;
    double base;
    double expected;
  };
  const Case cases[] = {
      {"Calibri", 11, false, 8.0, 10.0 + 213.0 / 256.0},   // no <sheetFormatPr>: base 10
      {"Calibri", 11, true, 8.0, 8.0 + 213.0 / 256.0},     // 8.83203125
      {"Calibri", 11, true, 10.0, 10.0 + 213.0 / 256.0},   // 10.83203125
      {"游ゴシック", 12, true, 8.0, 8.0 + 182.0 / 256.0},  // 8.7109375
      {"Calibri", 14, true, 10.0, 10.0 + 182.0 / 256.0},   // 10.7109375
      {"Calibri", 20, true, 10.0, 10.0 + 179.0 / 256.0},   // MDW 10, padding 7
  };
  for (const Case& c : cases) {
    Sheet sheet(std::string("Sheet1"));
    SheetFormatDefaults& defaults = sheet.mutable_format_defaults();
    defaults.has_default_row_height = c.has_row_height;
    defaults.base_col_width = c.base;
    EXPECT_DOUBLE_EQ(default_column_width_chars(sheet, MacModel(c.family, c.size)), c.expected)
        << c.family << c.size << " base " << c.base;
  }
}

TEST(FontMetrics, MacDefaultIgnoresBaseWhenDefaultColWidthIsStated) {
  Sheet sheet(std::string("Sheet1"));
  sheet.mutable_format_defaults().has_default_col_width = true;
  sheet.mutable_format_defaults().default_col_width = 12.5;
  EXPECT_DOUBLE_EQ(default_column_width_chars(sheet, MacModel("Calibri", 11)), 12.5);
}

TEST(FontMetrics, WindowsDefaultStoredWidthIsUnchanged) {
  Sheet sheet(std::string("Sheet1"));
  const ColumnWidthModel model = resolve_column_width_model(StylesWith("Calibri", 11), default_theme(),
                                                            win_365_en_us_profile(), GeometryMode::kDisplay);
  EXPECT_EQ(model.host, ExcelHost::kWin365);
  EXPECT_STREQ(model.platform, "win");
  EXPECT_DOUBLE_EQ(default_column_width_chars(sheet, model), 9.140625);
  EXPECT_DOUBLE_EQ(model.points_per_char, 5.25);
  EXPECT_DOUBLE_EQ(model.padding_pt, 3.75);
}

TEST(FontMetrics, PrintIsTheWindowsFigureOnEveryHost) {
  const ColumnWidthModel mac_print = MacModel("Calibri", 11, GeometryMode::kPrint);
  const ColumnWidthModel win_print = resolve_column_width_model(StylesWith("Calibri", 11), default_theme(),
                                                                win_365_en_us_profile(), GeometryMode::kPrint);
  EXPECT_EQ(mac_print.host, ExcelHost::kWin365);
  EXPECT_STREQ(mac_print.platform, "win");
  EXPECT_DOUBLE_EQ(mac_print.points_per_char, win_print.points_per_char);
  EXPECT_DOUBLE_EQ(mac_print.padding_pt, win_print.padding_pt);
  EXPECT_NEAR(mac_print.points_per_char, 39.0 / 7.0, 1e-12);
  EXPECT_NEAR(mac_print.padding_pt, 27.0 / 7.0, 1e-12);
  EXPECT_EQ(mac_print.max_digit_width_px, 7);
}

TEST(FontMetrics, ModelReportsTheResolvedThemeFace) {
  const ColumnWidthModel model = resolve_column_width_model(
      StylesWith("Calibri", 12, kFontSchemeMinor), default_theme(), mac_365_ja_jp_profile(), GeometryMode::kDisplay);
  EXPECT_EQ(model.normal_font_name, "游ゴシック");
  EXPECT_DOUBLE_EQ(model.normal_font_size, 12.0);
  EXPECT_DOUBLE_EQ(model.points_per_char, 7.0);
  EXPECT_DOUBLE_EQ(model.padding_pt, 5.0);
  EXPECT_TRUE(model.calibrated);
}

}  // namespace
// A Normal font set without the theme link on a ja-JP Mac book keeps its own metrics (Mac Excel 16.115:
// Meiryo 11 draws column B at stored width 10 as 70 pt, Calibri 12 and Aptos Narrow 12 as 60 pt).
TEST(FontMetrics, MacExplicitNormalFontOnAJapaneseBookUsesThatFont) {
  struct Case {
    const char* family;
    double size;
    double pt_at[3];
  };
  const Case cases[] = {
      {"Aptos Narrow", 12, {55, 64, 124}}, {"Calibri", 12, {55, 64, 124}}, {"Meiryo", 11, {64, 75, 145}}};
  const double stored[3] = {9.140625, 10.7109375, 20.7109375};
  for (const Case& c : cases) {
    const ColumnWidthModel model = resolve_column_width_model(StylesWith(c.family, c.size), default_theme(),
                                                              mac_365_ja_jp_profile(), GeometryMode::kDisplay);
    EXPECT_TRUE(model.calibrated) << c.family;
    for (int i = 0; i < 3; ++i) {
      EXPECT_DOUBLE_EQ(column_chars_to_points(stored[i], model), c.pt_at[i]) << c.family << " " << stored[i];
    }
  }
}

}  // namespace formulon
