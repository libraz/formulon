// A font linked to the theme (scheme minor or major) carries the name of the theme's face for the profile's script,
// resolved at each event that can change it: load, theme edit, profile change.

#include <cstdint>
#include <cstdio>
#include <string>
#include <utility>
#include <vector>

#include "cell_style_save_plan.h"
#include "excel_profile.h"
#include "font_metrics.h"
#include "gtest/gtest.h"
#include "io/ooxml_defs.h"
#include "io/ooxml_reader.h"
#include "io/theme_part.h"
#include "io/xlsb/reader.h"
#include "passthrough_part.h"
#include "theme.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

using Bytes = std::vector<std::uint8_t>;

constexpr const char* kMsPGothic =
    "\xEF\xBC\xAD\xEF\xBC\xB3 \xEF\xBC\xB0\xE3\x82\xB4\xE3\x82\xB7\xE3\x83\x83\xE3\x82\xAF";  // ＭＳ Ｐゴシック
constexpr const char* kYuGothic = "\xE6\xB8\xB8\xE3\x82\xB4\xE3\x82\xB7\xE3\x83\x83\xE3\x82\xAF";
constexpr const char* kYuGothicLight = "\xE6\xB8\xB8\xE3\x82\xB4\xE3\x82\xB7\xE3\x83\x83\xE3\x82\xAF Light";

Bytes ReadFixture(const std::string& name) {
  Bytes out;
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/" + name;
  FILE* file = std::fopen(path.c_str(), "rb");
  if (file == nullptr) {
    ADD_FAILURE() << "could not open fixture: " << path;
    return out;
  }
  std::fseek(file, 0, SEEK_END);
  const long size = std::ftell(file);
  std::fseek(file, 0, SEEK_SET);
  out.resize(size > 0 ? static_cast<std::size_t>(size) : 0U);
  if (std::fread(out.data(), 1, out.size(), file) != out.size()) {
    ADD_FAILURE() << "short read on fixture: " << path;
  }
  std::fclose(file);
  return out;
}

Workbook Load(const std::string& fixture, ExcelProfile profile) {
  const Bytes bytes = ReadFixture(fixture);
  ReadOptions options;
  options.profile = profile;
  auto result = read_ooxml(ByteSpan{bytes.data(), bytes.size()}, options);
  EXPECT_TRUE(static_cast<bool>(result)) << (result ? "" : result.error().message);
  return result ? std::move(result.value().workbook) : Workbook::create_empty();
}

std::vector<std::string> Names(const Workbook& wb) {
  std::vector<std::string> names;
  for (const FontRecord& font : wb.styles().fonts) {
    names.push_back(font.name);
  }
  return names;
}

FontRecord SchemeFont(std::uint8_t scheme, const std::string& name) {
  FontRecord font;
  font.name = name;
  font.size = 11.0;
  font.has_size = true;
  font.scheme = scheme;
  return font;
}

/// A theme without script entries: minor/major latin only.
std::string LatinOnlyTheme(const std::string& minor, const std::string& major) {
  std::string xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"T\"><a:themeElements>"
      "<a:clrScheme name=\"T\">"
      "<a:dk1><a:srgbClr val=\"000000\"/></a:dk1><a:lt1><a:srgbClr val=\"FFFFFF\"/></a:lt1>"
      "<a:dk2><a:srgbClr val=\"112233\"/></a:dk2><a:lt2><a:srgbClr val=\"AABBCC\"/></a:lt2>"
      "<a:accent1><a:srgbClr val=\"100000\"/></a:accent1><a:accent2><a:srgbClr val=\"200000\"/></a:accent2>"
      "<a:accent3><a:srgbClr val=\"300000\"/></a:accent3><a:accent4><a:srgbClr val=\"400000\"/></a:accent4>"
      "<a:accent5><a:srgbClr val=\"500000\"/></a:accent5><a:accent6><a:srgbClr val=\"600000\"/></a:accent6>"
      "<a:hlink><a:srgbClr val=\"0000AA\"/></a:hlink><a:folHlink><a:srgbClr val=\"0000BB\"/></a:folHlink>"
      "</a:clrScheme><a:fontScheme name=\"T\">";
  xml += "<a:majorFont><a:latin typeface=\"" + major + "\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:majorFont>";
  xml += "<a:minorFont><a:latin typeface=\"" + minor + "\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:minorFont>";
  xml += "</a:fontScheme><a:fmtScheme name=\"T\"/></a:themeElements></a:theme>";
  return xml;
}

TEST(SchemeFontNames, ExcelSavedControlFollowsTheProfileScript) {
  for (const bool ja : {true, false}) {
    SCOPED_TRACE(ja ? "ja" : "en");
    const Workbook wb = Load("xlsb_feature_flags.xlsx", ja ? mac_365_ja_jp_profile() : mac_365_en_us_profile());
    const Theme theme = wb.load_theme().theme;
    int scheme_fonts = 0;
    for (const FontRecord& font : wb.styles().fonts) {
      if (font.scheme == 0U) {
        continue;
      }
      ++scheme_fonts;
      const bool major = font.scheme == kFontSchemeMajor;
      if (!major) {
        EXPECT_EQ(font.name, ja ? kMsPGothic : "Calibri");
      } else if (!ja) {
        EXPECT_EQ(font.name, "Cambria");
      }
    }
    EXPECT_GT(scheme_fonts, 0);
    EXPECT_EQ(theme.fonts.minor_latin, "Calibri");
  }
}

TEST(SchemeFontNames, ThemeMinorGeorgiaKeepsEverythingButTheName) {
  const Workbook ja = Load("theme_minor_georgia.xlsx", mac_365_ja_jp_profile());
  const Workbook en = Load("theme_minor_georgia.xlsx", mac_365_en_us_profile());
  ASSERT_EQ(ja.styles().fonts.size(), en.styles().fonts.size());
  int minors = 0;
  int majors = 0;
  int plain = 0;
  for (std::size_t i = 0; i < ja.styles().fonts.size(); ++i) {
    const FontRecord& a = ja.styles().fonts[i];
    const FontRecord& b = en.styles().fonts[i];
    EXPECT_EQ(a.scheme, b.scheme);
    EXPECT_EQ(a.size, b.size);
    EXPECT_EQ(a.has_family, b.has_family);
    EXPECT_EQ(a.family, b.family);
    EXPECT_EQ(a.has_charset, b.has_charset);
    EXPECT_EQ(a.charset, b.charset);
    if (a.scheme == kFontSchemeMinor) {
      ++minors;
      EXPECT_EQ(a.name, kYuGothic);
      EXPECT_EQ(b.name, "Georgia");
    } else if (a.scheme == kFontSchemeMajor) {
      ++majors;
      EXPECT_EQ(a.name, kYuGothicLight);
      EXPECT_EQ(b.name, "Aptos Display");
    } else {
      ++plain;
      EXPECT_EQ(a.name, b.name);
    }
  }
  EXPECT_GT(minors, 0);
  EXPECT_GT(majors, 0);
  EXPECT_GT(plain, 0);
}

TEST(SchemeFontNames, ProfileChangeReResolvesAndRestores) {
  Workbook wb = Load("theme_minor_georgia.xlsx", mac_365_ja_jp_profile());
  const std::vector<std::string> ja = Names(wb);
  wb.set_excel_profile(mac_365_en_us_profile());
  const std::vector<std::string> en = Names(wb);
  EXPECT_NE(ja, en);
  wb.set_excel_profile(mac_365_ja_jp_profile());
  EXPECT_EQ(Names(wb), ja);
}

TEST(SchemeFontNames, ThemeEditsReResolve) {
  Workbook wb = Workbook::create_for_profile(mac_365_en_us_profile());
  wb.mutable_styles().fonts.push_back(SchemeFont(kFontSchemeMinor, "Stale"));
  wb.mutable_styles().fonts.push_back(SchemeFont(kFontSchemeMajor, "Stale"));
  const std::size_t minor = wb.styles().fonts.size() - 2;
  const std::size_t major = wb.styles().fonts.size() - 1;

  ThemeFonts fonts = wb.load_theme().theme.fonts;
  fonts.minor_latin = "Verdana";
  fonts.major_latin = "Georgia";
  ASSERT_TRUE(static_cast<bool>(wb.set_theme_fonts(fonts)));
  EXPECT_EQ(wb.styles().fonts[minor].name, "Verdana");
  EXPECT_EQ(wb.styles().fonts[major].name, "Georgia");

  wb.set_excel_profile(mac_365_ja_jp_profile());
  ASSERT_TRUE(static_cast<bool>(wb.set_theme_script_font(false, "Jpan", "Test Face")));
  EXPECT_EQ(wb.styles().fonts[minor].name, "Test Face");

  ASSERT_TRUE(static_cast<bool>(wb.reset_theme()));
  const Theme& stock = default_theme();
  EXPECT_NE(wb.styles().fonts[minor].name, "Test Face");
  EXPECT_FALSE(wb.styles().fonts[minor].name.empty());
  wb.set_excel_profile(mac_365_en_us_profile());
  EXPECT_EQ(wb.styles().fonts[minor].name, stock.fonts.minor_latin);
  EXPECT_EQ(wb.styles().fonts[major].name, stock.fonts.major_latin);
}

// Mac Excel keeps the stored name when the theme has no entry for the UI's script.
TEST(SchemeFontNames, MissingScriptEntryKeepsTheStoredName) {
  Workbook wb = Workbook::create_for_profile(mac_365_ja_jp_profile());
  const std::string xml = LatinOnlyTheme("Latin Minor", "Latin Major");
  const std::string path = theme_part_path(wb);
  PassthroughPart part(path, kThemeContentType, Bytes(xml.begin(), xml.end()));
  if (!static_cast<bool>(wb.replace_passthrough_part(path, Bytes(xml.begin(), xml.end())))) {
    ASSERT_TRUE(static_cast<bool>(wb.add_passthrough_part(std::move(part))));
  }
  wb.mutable_styles().fonts.push_back(SchemeFont(kFontSchemeMinor, "Stale"));
  wb.mutable_styles().fonts.push_back(SchemeFont(kFontSchemeMajor, "Stale"));
  resolve_scheme_font_names(wb);
  EXPECT_EQ(wb.styles().fonts[wb.styles().fonts.size() - 2].name, "Stale");
  EXPECT_EQ(wb.styles().fonts[wb.styles().fonts.size() - 1].name, "Stale");
}

TEST(SchemeFontNames, ResolvingTwiceEqualsOnceAndSkipsPlainFonts) {
  Workbook wb = Load("theme_minor_georgia.xlsx", mac_365_ja_jp_profile());
  wb.mutable_styles().fonts.push_back(SchemeFont(0, "Stale"));
  const std::vector<std::string> once = Names(wb);
  resolve_scheme_font_names(wb);
  resolve_scheme_font_names(wb);
  EXPECT_EQ(Names(wb), once);
  EXPECT_EQ(wb.styles().fonts.back().name, "Stale");
}

TEST(SchemeFontNames, SavePlanWritesTheModelNames) {
  const Workbook wb = Load("theme_minor_georgia.xlsx", mac_365_ja_jp_profile());
  const FontSavePlan plan = build_font_save_plan(wb);
  ASSERT_GE(plan.fonts().size(), wb.styles().fonts.size());
  for (std::size_t i = 0; i < wb.styles().fonts.size(); ++i) {
    EXPECT_EQ(plan.fonts()[i].name, wb.styles().fonts[i].name) << i;
  }
}

// A note run carries a copy of its table font, so it resolves with the table and still matches its entry.
TEST(SchemeFontNames, NoteRunFontsResolveWithTheTable) {
  const Bytes bytes = ReadFixture("xlsb_parity/notes/runs_fonts.xlsb");
  ReadOptions options;
  options.profile = mac_365_en_us_profile();
  auto read = xlsb::read_xlsb(ByteSpan{bytes.data(), bytes.size()}, options);
  ASSERT_TRUE(static_cast<bool>(read));
  const Workbook& wb = read.value().workbook;
  int scheme_runs = 0;
  for (const CellComment& note : wb.sheet(0).comments()) {
    for (const TextRun& run : note.runs) {
      if (run.font.has_value() && run.font->scheme == kFontSchemeMinor) {
        EXPECT_EQ(run.font->name, wb.styles().fonts[0].name);
        ++scheme_runs;
      }
    }
  }
  EXPECT_GT(scheme_runs, 0);
  EXPECT_EQ(wb.styles().fonts[0].name, "Calibri");
}

}  // namespace
}  // namespace io
}  // namespace formulon
