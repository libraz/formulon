#include "theme.h"

#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "passthrough_part.h"
#include "support/roundtrip_symmetry.h"
#include "workbook.h"

namespace formulon {
namespace {

constexpr const char* kThemeRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
constexpr const char* kThemeContentType = "application/vnd.openxmlformats-officedocument.theme+xml";

const PassthroughPart* FindPart(const Workbook& wb, const std::string& path) {
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    if (part.path == path) {
      return &part;
    }
  }
  return nullptr;
}

std::vector<std::string> PartPaths(const Workbook& wb) {
  std::vector<std::string> paths;
  paths.reserve(wb.passthrough_parts().size());
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    paths.push_back(part.path);
  }
  return paths;
}

Workbook RoundTrip(const Workbook& wb) {
  auto bytes = io::write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(bytes));
  auto read = io::read_ooxml(io::ByteSpan{bytes.value().data(), bytes.value().size()});
  EXPECT_TRUE(static_cast<bool>(read));
  return std::move(read.value().workbook);
}

Workbook WorkbookWithTheme(const std::string& xml) {
  Workbook wb = Workbook::create();
  EXPECT_TRUE(static_cast<bool>(wb.add_passthrough_part(
      PassthroughPart("xl/theme/theme1.xml", kThemeContentType, std::vector<std::uint8_t>(xml.begin(), xml.end())))));
  wb.add_workbook_relationship(kThemeRel, "xl/theme/theme1.xml");
  return wb;
}

constexpr const char* kCustomTheme =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Custom\"><a:themeElements>"
    "<a:clrScheme name=\"Custom\">"
    "<a:dk1><a:sysClr val=\"windowText\" lastClr=\"010203\"/></a:dk1><a:lt1><a:sysClr val=\"window\" "
    "lastClr=\"FEFDFC\"/></a:lt1>"
    "<a:dk2><a:srgbClr val=\"112233\"/></a:dk2><a:lt2><a:srgbClr val=\"AABBCC\"/></a:lt2>"
    "<a:accent1><a:srgbClr val=\"100000\"/></a:accent1><a:accent2><a:srgbClr val=\"200000\"/></a:accent2>"
    "<a:accent3><a:srgbClr val=\"300000\"/></a:accent3><a:accent4><a:srgbClr val=\"400000\"/></a:accent4>"
    "<a:accent5><a:srgbClr val=\"500000\"/></a:accent5><a:accent6><a:srgbClr val=\"600000\"/></a:accent6>"
    "<a:hlink><a:srgbClr val=\"0000AA\"/></a:hlink><a:folHlink><a:srgbClr val=\"0000BB\"/></a:folHlink>"
    "</a:clrScheme><a:fontScheme name=\"Custom\">"
    "<a:majorFont><a:latin typeface=\"Major Latin\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/>"
    "<a:font script=\"Jpan\" typeface=\"Major Jpan\"/></a:majorFont>"
    "<a:minorFont><a:latin typeface=\"Minor Latin\"/><a:ea typeface=\"Minor Ea\"/><a:cs typeface=\"\"/>"
    "<a:font script=\"Jpan\" typeface=\"Minor Jpan\"/></a:minorFont>"
    "</a:fontScheme><a:fmtScheme name=\"Custom\"><a:marker id=\"keep-me\"/></a:fmtScheme>"
    "</a:themeElements></a:theme>";

TEST(Theme, ReportsDefaultThemeWhenPartAbsent) {
  const Workbook wb = Workbook::create();
  const LoadedTheme loaded = wb.load_theme();
  EXPECT_EQ(loaded.source, ThemeSource::kDefault);
  EXPECT_EQ(loaded.theme.colors[4], 0xFF156082U);  // accent1 of the current Office theme
  EXPECT_EQ(loaded.theme.fonts.minor_latin, "Aptos Narrow");
  EXPECT_EQ(loaded.theme.fonts.major_latin, "Aptos Display");
}

TEST(Theme, GeneratesPartWhenAbsent) {
  Workbook wb = Workbook::create();
  ASSERT_EQ(FindPart(wb, "xl/theme/theme1.xml"), nullptr);

  ThemeColors colors = default_theme().colors;
  colors[0] = 0xFF111111U;
  colors[5] = 0xFF123456U;
  colors[11] = 0xFFABCDEFU;
  ASSERT_TRUE(static_cast<bool>(wb.set_theme_colors(colors)));
  const ThemeFonts fonts = {"Arial", "Meiryo", "Verdana", "MS Gothic"};
  ASSERT_TRUE(static_cast<bool>(wb.set_theme_fonts(fonts)));

  const PassthroughPart* part = FindPart(wb, "xl/theme/theme1.xml");
  ASSERT_NE(part, nullptr);
  EXPECT_EQ(part->content_type, kThemeContentType);

  const Workbook reloaded = RoundTrip(wb);
  const PassthroughPart* saved = FindPart(reloaded, "xl/theme/theme1.xml");
  ASSERT_NE(saved, nullptr);
  EXPECT_EQ(saved->content_type, kThemeContentType);
  bool has_rel = false;
  for (const UnknownRelationship& rel : reloaded.unknown_workbook_rels()) {
    has_rel = has_rel || (rel.type == kThemeRel && rel.target == "xl/theme/theme1.xml");
  }
  EXPECT_TRUE(has_rel);

  const LoadedTheme loaded = reloaded.load_theme();
  EXPECT_EQ(loaded.source, ThemeSource::kPart);
  EXPECT_EQ(loaded.theme.colors, colors);
  EXPECT_EQ(loaded.theme.fonts.major_latin, "Arial");
  EXPECT_EQ(loaded.theme.fonts.major_ea, "Meiryo");
  EXPECT_EQ(loaded.theme.fonts.minor_latin, "Verdana");
  EXPECT_EQ(loaded.theme.fonts.minor_ea, "MS Gothic");
}

TEST(Theme, ParsesSchemeAndFontsOfExistingPart) {
  const Workbook wb = WorkbookWithTheme(kCustomTheme);
  const LoadedTheme loaded = wb.load_theme();
  ASSERT_EQ(loaded.source, ThemeSource::kPart);
  EXPECT_EQ(loaded.theme.colors[0], 0xFF010203U);  // sysClr lastClr
  EXPECT_EQ(loaded.theme.colors[1], 0xFFFEFDFCU);
  EXPECT_EQ(loaded.theme.colors[2], 0xFF112233U);
  EXPECT_EQ(loaded.theme.colors[11], 0xFF0000BBU);
  EXPECT_EQ(loaded.theme.fonts.major_latin, "Major Latin");
  EXPECT_EQ(loaded.theme.fonts.major_ea, "Major Jpan");  // a:ea is empty: the Jpan face
  EXPECT_EQ(loaded.theme.fonts.minor_ea, "Minor Ea");    // a:ea names a face: it wins
}

TEST(Theme, EditPatchesOnlySchemeAndFontElements) {
  Workbook wb = WorkbookWithTheme(kCustomTheme);
  ThemeColors colors = wb.load_theme().theme.colors;
  colors[3] = 0xFF778899U;
  ASSERT_TRUE(static_cast<bool>(wb.set_theme_colors(colors)));
  ASSERT_TRUE(static_cast<bool>(wb.set_theme_fonts({"A", "B", "C", "D"})));

  const PassthroughPart* part = FindPart(wb, "xl/theme/theme1.xml");
  ASSERT_NE(part, nullptr);
  const std::string xml(part->bytes.begin(), part->bytes.end());
  EXPECT_NE(xml.find("<a:marker id=\"keep-me\"/>"), std::string::npos);
  EXPECT_NE(xml.find("name=\"Custom\""), std::string::npos);

  const LoadedTheme loaded = wb.load_theme();
  ASSERT_EQ(loaded.source, ThemeSource::kPart);
  EXPECT_EQ(loaded.theme.colors, colors);
  EXPECT_EQ(loaded.theme.fonts.major_ea, "B");
  EXPECT_EQ(loaded.theme.fonts.minor_ea, "D");
}

TEST(Theme, UnparseablePartReportsSourceAndRefusesEdit) {
  const std::string junk = "<a:theme>not a colour scheme</a:theme>";
  Workbook wb = WorkbookWithTheme(junk);
  const LoadedTheme loaded = wb.load_theme();
  EXPECT_EQ(loaded.source, ThemeSource::kUnparseable);
  EXPECT_EQ(loaded.theme.colors, default_theme().colors);

  EXPECT_FALSE(static_cast<bool>(wb.set_theme_colors(default_theme().colors)));
  const PassthroughPart* part = FindPart(wb, "xl/theme/theme1.xml");
  ASSERT_NE(part, nullptr);
  EXPECT_EQ(std::string(part->bytes.begin(), part->bytes.end()), junk);
}

TEST(Theme, ResetRemovesPartAndRelationship) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_theme_colors(default_theme().colors)));
  ASSERT_EQ(wb.load_theme().source, ThemeSource::kPart);

  ASSERT_TRUE(static_cast<bool>(wb.reset_theme()));
  EXPECT_EQ(wb.load_theme().source, ThemeSource::kDefault);
  EXPECT_EQ(FindPart(wb, "xl/theme/theme1.xml"), nullptr);
  EXPECT_TRUE(wb.unknown_workbook_rels().empty());
}

TEST(Theme, ResetWithoutThemeIsNoOp) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.reset_theme()));
  EXPECT_EQ(wb.load_theme().source, ThemeSource::kDefault);
  EXPECT_TRUE(wb.passthrough_parts().empty());
}

TEST(Theme, ResetRemovesUnparseablePart) {
  Workbook wb = WorkbookWithTheme("<a:theme>not a colour scheme</a:theme>");
  ASSERT_EQ(wb.load_theme().source, ThemeSource::kUnparseable);
  ASSERT_TRUE(static_cast<bool>(wb.reset_theme()));
  EXPECT_EQ(wb.load_theme().source, ThemeSource::kDefault);
  EXPECT_EQ(FindPart(wb, "xl/theme/theme1.xml"), nullptr);
  EXPECT_TRUE(wb.unknown_workbook_rels().empty());
}

TEST(Theme, ResetRemovesThemeRelsAndPartsOnlyItReferenced) {
  Workbook wb = WorkbookWithTheme(kCustomTheme);
  const std::string rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
      "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" "
      "Target=\"../media/themeimg.png\"/></Relationships>";
  ASSERT_TRUE(static_cast<bool>(wb.add_passthrough_part(
      PassthroughPart("xl/theme/_rels/theme1.xml.rels", "", std::vector<std::uint8_t>(rels.begin(), rels.end())))));
  ASSERT_TRUE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart("xl/media/themeimg.png", "", {1, 2, 3}))));
  ASSERT_TRUE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart("xl/media/other.png", "", {4}))));

  ASSERT_TRUE(static_cast<bool>(wb.reset_theme()));
  EXPECT_EQ(FindPart(wb, "xl/theme/theme1.xml"), nullptr);
  EXPECT_EQ(FindPart(wb, "xl/theme/_rels/theme1.xml.rels"), nullptr);
  EXPECT_EQ(FindPart(wb, "xl/media/themeimg.png"), nullptr);
  EXPECT_NE(FindPart(wb, "xl/media/other.png"), nullptr);
}

TEST(Theme, ResetSurvivesSaveAndLoad) {
  Workbook wb = WorkbookWithTheme(kCustomTheme);
  ASSERT_EQ(RoundTrip(wb).load_theme().source, ThemeSource::kPart);
  ASSERT_TRUE(static_cast<bool>(wb.reset_theme()));

  const Workbook reloaded = RoundTrip(wb);
  EXPECT_EQ(reloaded.load_theme().source, ThemeSource::kDefault);
  EXPECT_EQ(reloaded.load_theme().theme.colors, default_theme().colors);
  EXPECT_EQ(FindPart(reloaded, "xl/theme/theme1.xml"), nullptr);
}

/// Typeface of `script` in `scripts`, or empty when the list has no entry.
std::string ScriptFace(const std::vector<ThemeScriptFont>& scripts, const std::string& script) {
  for (const ThemeScriptFont& font : scripts) {
    if (font.script == script) {
      return font.typeface;
    }
  }
  return std::string();
}

Workbook LoadExcelNewWorkbook() {
  const std::vector<std::uint8_t> bytes =
      test::read_file_bytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/new_workbook_mac_en_us.xlsx");
  auto read = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(read));
  return std::move(read.value().workbook);
}

/// The default Office script faces from the constexpr table, in table order.
std::vector<ThemeScriptFont> OfficeScripts(bool major) {
  std::vector<ThemeScriptFont> out;
  for (std::size_t i = 0; i < kOfficeScriptCount; ++i) {
    out.push_back({office_script(i), office_script_face(i, major)});
  }
  return out;
}

TEST(Theme, DefaultThemeCarriesOfficeScriptFonts) {
  const Theme& def = default_theme();
  const std::vector<ThemeScriptFont> major_scripts = OfficeScripts(true);
  const std::vector<ThemeScriptFont> minor_scripts = OfficeScripts(false);
  ASSERT_EQ(major_scripts.size(), 47U);
  ASSERT_EQ(minor_scripts.size(), 47U);
  // Values and order from a Mac Excel 365 new workbook's theme part.
  EXPECT_EQ(major_scripts.front().script, "Jpan");
  EXPECT_EQ(major_scripts.front().typeface, "游ゴシック Light");
  EXPECT_EQ(minor_scripts.front().typeface, "游ゴシック");
  EXPECT_EQ(ScriptFace(major_scripts, "Khmr"), "MoolBoran");
  EXPECT_EQ(ScriptFace(minor_scripts, "Khmr"), "DaunPenh");
  EXPECT_EQ(ScriptFace(major_scripts, "Hans"), "等线 Light");
  EXPECT_EQ(ScriptFace(minor_scripts, "Arab"), "Arial");
  EXPECT_EQ(major_scripts.back().script, "Tfng");
  EXPECT_EQ(minor_scripts.back().typeface, "Ebrima");
  EXPECT_EQ(def.fonts.major_ea, ScriptFace(major_scripts, "Jpan"));
  EXPECT_EQ(def.fonts.minor_ea, ScriptFace(minor_scripts, "Jpan"));

  const LoadedTheme loaded = Workbook::create().load_theme();
  EXPECT_EQ(loaded.source, ThemeSource::kDefault);
  EXPECT_TRUE(loaded.theme.minor_scripts.empty());
}

TEST(Theme, ExcelSavedThemeScriptFontsMatchDefaultTable) {
  const Workbook wb = LoadExcelNewWorkbook();
  const LoadedTheme loaded = wb.load_theme();
  ASSERT_EQ(loaded.source, ThemeSource::kPart);
  ASSERT_EQ(loaded.theme.major_scripts.size(), 47U);
  ASSERT_EQ(loaded.theme.minor_scripts.size(), 47U);
  EXPECT_EQ(ScriptFace(loaded.theme.minor_scripts, "Khmr"), "DaunPenh");
  EXPECT_EQ(loaded.theme.major_scripts[1].script, "Hang");
  const std::vector<ThemeScriptFont> def_majors = OfficeScripts(true);
  const std::vector<ThemeScriptFont> def_minors = OfficeScripts(false);
  for (std::size_t i = 0; i < 47U; ++i) {
    EXPECT_EQ(loaded.theme.major_scripts[i].script, def_majors[i].script) << i;
    EXPECT_EQ(loaded.theme.major_scripts[i].typeface, def_majors[i].typeface) << i;
    EXPECT_EQ(loaded.theme.minor_scripts[i].script, def_minors[i].script) << i;
    EXPECT_EQ(loaded.theme.minor_scripts[i].typeface, def_minors[i].typeface) << i;
  }
}

TEST(Theme, ParsesScriptFontsInFileOrder) {
  const LoadedTheme loaded = WorkbookWithTheme(kCustomTheme).load_theme();
  ASSERT_EQ(loaded.theme.major_scripts.size(), 1U);
  EXPECT_EQ(loaded.theme.major_scripts[0].script, "Jpan");
  EXPECT_EQ(loaded.theme.major_scripts[0].typeface, "Major Jpan");
  ASSERT_EQ(loaded.theme.minor_scripts.size(), 1U);
  EXPECT_EQ(loaded.theme.minor_scripts[0].typeface, "Minor Jpan");
}

TEST(Theme, SetScriptFontWithoutThemePartKeepsOtherDefaults) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_theme_script_font(false, "Hang", "Batang")));

  const LoadedTheme loaded = RoundTrip(wb).load_theme();
  ASSERT_EQ(loaded.source, ThemeSource::kPart);
  ASSERT_EQ(loaded.theme.minor_scripts.size(), 47U);
  ASSERT_EQ(loaded.theme.major_scripts.size(), 47U);
  const std::vector<ThemeScriptFont> def_majors = OfficeScripts(true);
  const std::vector<ThemeScriptFont> def_minors = OfficeScripts(false);
  for (std::size_t i = 0; i < 47U; ++i) {
    const ThemeScriptFont& def_minor = def_minors[i];
    EXPECT_EQ(loaded.theme.minor_scripts[i].script, def_minor.script);
    EXPECT_EQ(loaded.theme.minor_scripts[i].typeface, def_minor.script == "Hang" ? "Batang" : def_minor.typeface);
    EXPECT_EQ(loaded.theme.major_scripts[i].typeface, def_majors[i].typeface);
  }
}

TEST(Theme, SetScriptFontOnExcelThemeChangesOnlyThatScript) {
  Workbook wb = LoadExcelNewWorkbook();
  ASSERT_TRUE(static_cast<bool>(wb.set_theme_script_font(true, "Thai", "Leelawadee UI")));

  const LoadedTheme loaded = RoundTrip(wb).load_theme();
  ASSERT_EQ(loaded.theme.major_scripts.size(), 47U);
  EXPECT_EQ(ScriptFace(loaded.theme.major_scripts, "Thai"), "Leelawadee UI");
  EXPECT_EQ(ScriptFace(loaded.theme.minor_scripts, "Thai"), "Tahoma");
  EXPECT_EQ(ScriptFace(loaded.theme.major_scripts, "Khmr"), "MoolBoran");
  EXPECT_EQ(loaded.theme.fonts.major_latin, "Aptos Display");
}

TEST(Theme, SetScriptFontAppendsAMissingScript) {
  Workbook wb = WorkbookWithTheme(kCustomTheme);
  ASSERT_TRUE(static_cast<bool>(wb.set_theme_script_font(true, "Hang", "Gulim")));
  const LoadedTheme loaded = wb.load_theme();
  ASSERT_EQ(loaded.theme.major_scripts.size(), 2U);
  EXPECT_EQ(loaded.theme.major_scripts[0].script, "Jpan");
  EXPECT_EQ(loaded.theme.major_scripts[1].script, "Hang");
  EXPECT_EQ(loaded.theme.major_scripts[1].typeface, "Gulim");
  EXPECT_EQ(loaded.theme.minor_scripts.size(), 1U);
}

TEST(Theme, SetJpanScriptFontMovesEastAsianFace) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_theme_script_font(false, "Jpan", "Meiryo")));
  const LoadedTheme loaded = RoundTrip(wb).load_theme();
  EXPECT_EQ(loaded.theme.fonts.minor_ea, "Meiryo");
  EXPECT_EQ(ScriptFace(loaded.theme.minor_scripts, "Jpan"), "Meiryo");
  EXPECT_EQ(loaded.theme.fonts.major_ea, "游ゴシック Light");
}

TEST(Theme, SetScriptFontRejectsMalformedInput) {
  Workbook wb = Workbook::create();
  for (const char* script : {"", "Jpn", "Jpann", "J1an", "Jp n", "\xC3\xA9pan"}) {
    const auto rc = wb.set_theme_script_font(true, script, "Arial");
    ASSERT_FALSE(static_cast<bool>(rc)) << script;
    EXPECT_EQ(rc.error().code, FormulonErrorCode::kInvalidArgument);
  }
  const auto empty_face = wb.set_theme_script_font(true, "Jpan", "");
  ASSERT_FALSE(static_cast<bool>(empty_face));
  EXPECT_EQ(empty_face.error().code, FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(FindPart(wb, "xl/theme/theme1.xml"), nullptr);
}

TEST(Theme, PassthroughPartApiRejectsDuplicatesAndMissing) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart("xl/x.bin", "", {1, 2}))));
  EXPECT_FALSE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart("xl/x.bin", "", {3}))));
  ASSERT_TRUE(static_cast<bool>(wb.replace_passthrough_part("xl/x.bin", {9})));
  EXPECT_EQ(FindPart(wb, "xl/x.bin")->bytes, std::vector<std::uint8_t>({9}));
  EXPECT_FALSE(static_cast<bool>(wb.replace_passthrough_part("xl/missing.bin", {1})));
  wb.add_workbook_relationship("t", "xl/x.bin");
  wb.add_workbook_relationship("t", "xl/x.bin");
  EXPECT_EQ(wb.unknown_workbook_rels().size(), 1U);
}

TEST(Theme, PassthroughPartInsertionPreservesOrderAndRecordIdentity) {
  Workbook wb = Workbook::create();
  PassthroughPart retained("xl/media/retained.bin", "application/x-retained", {1, 2, 3});
  retained.retained_origin = PassthroughPart::RetainedOrigin::kStyles;
  retained.origin_cache_id = 17U;
  retained.model_fingerprint = 0x1234U;
  const std::uint8_t* const payload = retained.bytes.data();

  ASSERT_TRUE(static_cast<bool>(wb.add_passthrough_part(std::move(retained))));
  ASSERT_TRUE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart("xl/media/z.bin", "", {9}))));
  ASSERT_TRUE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart("xl/media/a.bin", "", {0}))));
  ASSERT_TRUE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart("xl/media/m.bin", "", {8}))));

  EXPECT_EQ(PartPaths(wb),
            (std::vector<std::string>{"xl/media/a.bin", "xl/media/m.bin", "xl/media/retained.bin", "xl/media/z.bin"}));
  const PassthroughPart* found = FindPart(wb, "xl/media/retained.bin");
  ASSERT_NE(found, nullptr);
  EXPECT_EQ(found->bytes.data(), payload);
  EXPECT_EQ(found->bytes, (std::vector<std::uint8_t>{1, 2, 3}));
  EXPECT_EQ(found->content_type, "application/x-retained");
  EXPECT_EQ(found->retained_origin, PassthroughPart::RetainedOrigin::kStyles);
  EXPECT_EQ(found->origin_cache_id, 17U);
  ASSERT_TRUE(found->model_fingerprint.has_value());
  EXPECT_EQ(*found->model_fingerprint, 0x1234U);

  EXPECT_FALSE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart{})));
  EXPECT_EQ(PartPaths(wb),
            (std::vector<std::string>{"xl/media/a.bin", "xl/media/m.bin", "xl/media/retained.bin", "xl/media/z.bin"}));
}

TEST(Theme, PassthroughPartSetterRetainsUnsortedOrderUntilSuccessfulInsertion) {
  Workbook wb = Workbook::create();
  std::vector<PassthroughPart> parts;
  parts.emplace_back("xl/z.bin", "z/type", std::vector<std::uint8_t>{9});
  parts.emplace_back("xl/a.bin", "a/type", std::vector<std::uint8_t>{1});
  parts.emplace_back("xl/m.bin", "m/type", std::vector<std::uint8_t>{5});
  wb.set_passthrough_parts(std::move(parts));

  EXPECT_EQ(PartPaths(wb), (std::vector<std::string>{"xl/z.bin", "xl/a.bin", "xl/m.bin"}));
  EXPECT_FALSE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart("xl/a.bin", "changed", {7}))));
  EXPECT_EQ(PartPaths(wb), (std::vector<std::string>{"xl/z.bin", "xl/a.bin", "xl/m.bin"}));
  EXPECT_EQ(wb.passthrough_parts()[1].bytes, (std::vector<std::uint8_t>{1}));
  EXPECT_EQ(wb.passthrough_parts()[1].content_type, "a/type");

  ASSERT_TRUE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart("xl/b.bin", "b/type", {2}))));
  EXPECT_EQ(PartPaths(wb), (std::vector<std::string>{"xl/a.bin", "xl/b.bin", "xl/m.bin", "xl/z.bin"}));

  ASSERT_TRUE(static_cast<bool>(wb.replace_passthrough_part("xl/m.bin", {6, 6})));
  EXPECT_EQ(PartPaths(wb), (std::vector<std::string>{"xl/a.bin", "xl/b.bin", "xl/m.bin", "xl/z.bin"}));
  EXPECT_EQ(wb.passthrough_parts()[2].bytes, (std::vector<std::uint8_t>{6, 6}));
  EXPECT_EQ(wb.passthrough_parts()[2].content_type, "m/type");
}

TEST(Theme, PassthroughPartSortedDuplicatesAndMovesPreserveStableOrder) {
  Workbook wb = Workbook::create();
  std::vector<PassthroughPart> parts;
  parts.emplace_back("xl/a.bin", "first", std::vector<std::uint8_t>{1});
  parts.emplace_back("xl/a.bin", "second", std::vector<std::uint8_t>{2});
  parts.emplace_back("xl/c.bin", "third", std::vector<std::uint8_t>{3});
  wb.set_passthrough_parts(std::move(parts));

  ASSERT_TRUE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart("xl/b.bin", "between", {4}))));
  ASSERT_EQ(wb.passthrough_parts().size(), 4U);
  EXPECT_EQ(PartPaths(wb), (std::vector<std::string>{"xl/a.bin", "xl/a.bin", "xl/b.bin", "xl/c.bin"}));
  EXPECT_EQ(wb.passthrough_parts()[0].content_type, "first");
  EXPECT_EQ(wb.passthrough_parts()[1].content_type, "second");

  std::vector<std::uint8_t> tombstones{0, 1, 0, 0};
  ASSERT_TRUE(static_cast<bool>(wb.erase_passthrough_parts(std::move(tombstones))));
  EXPECT_EQ(PartPaths(wb), (std::vector<std::string>{"xl/a.bin", "xl/b.bin", "xl/c.bin"}));

  Workbook moved = std::move(wb);
  ASSERT_TRUE(static_cast<bool>(moved.add_passthrough_part(PassthroughPart("xl/d.bin", "last", {5}))));
  EXPECT_EQ(PartPaths(moved), (std::vector<std::string>{"xl/a.bin", "xl/b.bin", "xl/c.bin", "xl/d.bin"}));

  Workbook assigned = Workbook::create();
  assigned = std::move(moved);
  ASSERT_TRUE(static_cast<bool>(assigned.add_passthrough_part(PassthroughPart("xl/e.bin", "last", {6}))));
  EXPECT_EQ(PartPaths(assigned),
            (std::vector<std::string>{"xl/a.bin", "xl/b.bin", "xl/c.bin", "xl/d.bin", "xl/e.bin"}));
}

}  // namespace
}  // namespace formulon
