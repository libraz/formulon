// Theme packages (.thmx): applying one to a workbook and saving the current theme as one.

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "eval/dep_graph.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/ooxml/part_dom.h"
#include "io/ooxml/zip_part_writer.h"
#include "io/ooxml_defs.h"
#include "io/ooxml_writer.h"
#include "io/theme_part.h"
#include "io/zip_reader.h"
#include "passthrough_part.h"
#include "theme.h"
#include "workbook.h"
#include "workbook_theme_package.h"

namespace formulon {
namespace io {
namespace {

using Bytes = std::vector<std::uint8_t>;

constexpr const char* kRelNs = "http://schemas.openxmlformats.org/package/2006/relationships";
constexpr const char* kRelBase = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
constexpr const char* kThemePath = "xl/theme/theme1.xml";
constexpr const char* kThemeRelsPath = "xl/theme/_rels/theme1.xml.rels";
constexpr std::uint32_t kRgbMask = 0xFFFFFFU;

std::string RelsXml(const std::string& body) {
  return std::string("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"") +
         kRelNs + "\">" + body + "</Relationships>";
}

std::string Rel(const std::string& id, const std::string& type, const std::string& target) {
  return "<Relationship Id=\"" + id + "\" Type=\"" + kRelBase + type + "\" Target=\"" + target + "\"/>";
}

/// A complete theme: `dk2` is the second dark colour, `minor` the minor Latin face, `jpan` the minor Jpan entry.
std::string ThemeXml(const std::string& dk2, const std::string& minor, const std::string& jpan = std::string()) {
  const std::string jpan_font = jpan.empty() ? std::string() : "<a:font script=\"Jpan\" typeface=\"" + jpan + "\"/>";
  return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
         "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Pkg\"><a:themeElements>"
         "<a:clrScheme name=\"Pkg\">"
         "<a:dk1><a:sysClr val=\"windowText\" lastClr=\"010203\"/></a:dk1><a:lt1><a:sysClr val=\"window\" "
         "lastClr=\"FEFDFC\"/></a:lt1>"
         "<a:dk2><a:srgbClr val=\"" +
         dk2 +
         "\"/></a:dk2><a:lt2><a:srgbClr val=\"AABBCC\"/></a:lt2>"
         "<a:accent1><a:srgbClr val=\"100000\"/></a:accent1><a:accent2><a:srgbClr val=\"200000\"/></a:accent2>"
         "<a:accent3><a:srgbClr val=\"300000\"/></a:accent3><a:accent4><a:srgbClr val=\"400000\"/></a:accent4>"
         "<a:accent5><a:srgbClr val=\"500000\"/></a:accent5><a:accent6><a:srgbClr val=\"600000\"/></a:accent6>"
         "<a:hlink><a:srgbClr val=\"0000AA\"/></a:hlink><a:folHlink><a:srgbClr val=\"0000BB\"/></a:folHlink>"
         "</a:clrScheme><a:fontScheme name=\"Pkg\">"
         "<a:majorFont><a:latin typeface=\"Major Latin\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:majorFont>"
         "<a:minorFont><a:latin typeface=\"" +
         minor + "\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/>" + jpan_font +
         "</a:minorFont></a:fontScheme><a:fmtScheme name=\"Pkg\"/></a:themeElements></a:theme>";
}

/// 24 bytes of PNG header: enough for the image probe.
Bytes Png(std::uint8_t tag) {
  Bytes png = {0x89, 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 13, 'I', 'H', 'D',
               'R',  0,   0,   0,   2,    0,    0,    0,    2, 8, 6, 0,  0,   0,   tag};
  return png;
}

struct PackageSpec {
  std::string theme = ThemeXml("112233", "Minor One");
  std::string theme_rels;  // body of theme1.xml.rels; empty = no part
  std::vector<std::pair<std::string, Bytes>> media;
  bool office_document_rel = true;
  bool manager_theme_rel = true;
};

Bytes BuildPackage(const PackageSpec& spec) {
  ZipWriterGuard writer;
  EXPECT_TRUE(writer.init());
  const auto text = [&writer](const char* path, const std::string& body) {
    EXPECT_TRUE(static_cast<bool>(AddPart(writer.get(), path, body)));
  };
  text("[Content_Types].xml", "<Types/>");
  text("_rels/.rels",
       RelsXml(spec.office_document_rel ? Rel("rId1", "officeDocument", "theme/theme/themeManager.xml") : ""));
  text("theme/theme/themeManager.xml",
       "<a:themeManager xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/>");
  text("theme/theme/_rels/themeManager.xml.rels",
       RelsXml(spec.manager_theme_rel ? Rel("rId1", "theme", "theme1.xml") : ""));
  text("theme/theme/theme1.xml", spec.theme);
  if (!spec.theme_rels.empty()) {
    text("theme/theme/_rels/theme1.xml.rels", RelsXml(spec.theme_rels));
  }
  for (const auto& [name, bytes] : spec.media) {
    EXPECT_TRUE(static_cast<bool>(AddPartBytes(writer.get(), name, bytes)));
  }
  auto out = FinalizeArchive(writer, "test");
  EXPECT_TRUE(static_cast<bool>(out));
  return std::move(out.value());
}

PackageSpec ImageSpec(std::uint8_t tag = 1U) {
  PackageSpec spec;
  spec.theme_rels = Rel("rId1", "image", "../media/image1.png");
  spec.media.emplace_back("theme/media/image1.png", Png(tag));
  return spec;
}

Expected<void, Error> Apply(Workbook& wb, const Bytes& pkg) {
  return apply_theme_package(wb, pkg.data(), pkg.size());
}

std::string Entry(const Bytes& archive, const char* name) {
  ZipReader zip;
  if (!zip.open(ByteSpan{archive.data(), archive.size()}) || !zip.has_entry(name)) {
    return "<missing>";
  }
  auto bytes = zip.read_entry(name);
  return bytes ? std::string(bytes.value().begin(), bytes.value().end()) : "<unreadable>";
}

std::size_t MediaCount(const Workbook& wb) {
  std::size_t n = 0;
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    n += part.path.rfind("xl/media/", 0) == 0 ? 1U : 0U;
  }
  return n;
}

Bytes ThemeBytes(const Workbook& wb) {
  const PassthroughPart* part = ooxml::find_passthrough_part(wb, kThemePath);
  return part != nullptr ? part->bytes : Bytes();
}

TEST(ThemePackage, AppliesColoursAndFonts) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(Apply(wb, BuildPackage(PackageSpec()))));
  const LoadedTheme loaded = wb.load_theme();
  EXPECT_EQ(loaded.source, ThemeSource::kPart);
  EXPECT_EQ(loaded.theme.colors[2] & kRgbMask, 0x112233U);
  EXPECT_EQ(loaded.theme.fonts.minor_latin, "Minor One");
  EXPECT_EQ(loaded.theme.fonts.major_latin, "Major Latin");
}

TEST(ThemePackage, CarriesImageRelationshipAndMedia) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(Apply(wb, BuildPackage(ImageSpec()))));
  EXPECT_EQ(wb.load_theme().source, ThemeSource::kPart);
  EXPECT_NE(ooxml::find_passthrough_part(wb, "xl/media/image1.png"), nullptr);
  const PassthroughPart* rels = ooxml::find_passthrough_part(wb, kThemeRelsPath);
  ASSERT_NE(rels, nullptr);
  const std::string xml(rels->bytes.begin(), rels->bytes.end());
  EXPECT_NE(xml.find("Id=\"rId1\""), std::string::npos);
  EXPECT_NE(xml.find("Target=\"../media/image1.png\""), std::string::npos);
}

TEST(ThemePackage, ReapplyDropsOldMedia) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(Apply(wb, BuildPackage(ImageSpec(1U)))));
  ASSERT_TRUE(static_cast<bool>(Apply(wb, BuildPackage(ImageSpec(2U)))));
  EXPECT_EQ(MediaCount(wb), 1U);
  EXPECT_NE(ooxml::find_passthrough_part(wb, kThemeRelsPath), nullptr);
  ASSERT_TRUE(static_cast<bool>(Apply(wb, BuildPackage(PackageSpec()))));
  EXPECT_EQ(MediaCount(wb), 0U);
  EXPECT_EQ(ooxml::find_passthrough_part(wb, kThemeRelsPath), nullptr);
  EXPECT_EQ(wb.load_theme().source, ThemeSource::kPart);
}

TEST(ThemePackage, SaveThenApplyRoundTrips) {
  Workbook source = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(Apply(source, BuildPackage(ImageSpec()))));
  auto saved = save_theme_package(source);
  ASSERT_TRUE(static_cast<bool>(saved));
  EXPECT_NE(Entry(saved.value(), "theme/media/image1.png"), "<missing>");
  EXPECT_NE(Entry(saved.value(), "theme/theme/_rels/theme1.xml.rels").find("../media/image1.png"), std::string::npos);

  Workbook target = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(Apply(target, saved.value())));
  const LoadedTheme a = source.load_theme();
  const LoadedTheme b = target.load_theme();
  EXPECT_EQ(a.theme.colors, b.theme.colors);
  EXPECT_EQ(a.theme.fonts.minor_latin, b.theme.fonts.minor_latin);
  EXPECT_EQ(a.theme.fonts.major_latin, b.theme.fonts.major_latin);
  EXPECT_EQ(MediaCount(target), 1U);

  auto again = save_theme_package(target);
  ASSERT_TRUE(static_cast<bool>(again));
  EXPECT_EQ(Entry(again.value(), "theme/theme/theme1.xml"), Entry(saved.value(), "theme/theme/theme1.xml"));
}

TEST(ThemePackage, SaveWithoutThemePartYieldsDefaultTheme) {
  const Workbook wb = Workbook::create();
  auto saved = save_theme_package(wb);
  ASSERT_TRUE(static_cast<bool>(saved));
  EXPECT_EQ(Entry(saved.value(), "theme/theme/theme1.xml"), build_default_theme_xml());
  EXPECT_NE(Entry(saved.value(), "[Content_Types].xml").find("themeManager+xml"), std::string::npos);
  EXPECT_NE(Entry(saved.value(), "_rels/.rels").find("theme/theme/themeManager.xml"), std::string::npos);
}

TEST(ThemePackage, SaveRefusesUnparseableThemePart) {
  Workbook wb = Workbook::create();
  const std::string junk = "<a:theme>not a colour scheme</a:theme>";
  ASSERT_TRUE(static_cast<bool>(
      wb.add_passthrough_part(PassthroughPart(kThemePath, kThemeContentType, Bytes(junk.begin(), junk.end())))));
  wb.add_workbook_relationship(std::string(kRelTheme), kThemePath);
  auto saved = save_theme_package(wb);
  ASSERT_FALSE(static_cast<bool>(saved));
  EXPECT_EQ(saved.error().code, FormulonErrorCode::kIoXmlParse);
}

TEST(ThemePackage, FailuresLeaveTheThemeUntouched) {
  PackageSpec no_office;
  no_office.office_document_rel = false;
  PackageSpec no_theme_rel;
  no_theme_rel.manager_theme_rel = false;
  PackageSpec bad_theme;
  bad_theme.theme = "<a:theme>not a colour scheme</a:theme>";
  PackageSpec missing_media = ImageSpec();
  missing_media.media.clear();
  PackageSpec bad_image = ImageSpec();
  bad_image.media[0].second = {1, 2, 3};

  struct Case {
    const char* name;
    Bytes bytes;
    FormulonErrorCode code;
  };
  const std::vector<Case> cases = {
      {"garbage", Bytes{'n', 'o', 't', ' ', 'a', ' ', 'z', 'i', 'p'}, FormulonErrorCode::kIoZipCorrupt},
      {"no officeDocument", BuildPackage(no_office), FormulonErrorCode::kIoRelationshipBroken},
      {"no theme rel", BuildPackage(no_theme_rel), FormulonErrorCode::kIoRelationshipBroken},
      {"unparseable theme", BuildPackage(bad_theme), FormulonErrorCode::kIoXmlParse},
      {"missing media", BuildPackage(missing_media), FormulonErrorCode::kIoRelationshipBroken},
      {"unsupported media", BuildPackage(bad_image), FormulonErrorCode::kIoImageUnsupported},
  };
  for (const Case& c : cases) {
    Workbook wb = Workbook::create();
    ASSERT_TRUE(static_cast<bool>(Apply(wb, BuildPackage(ImageSpec(7U))))) << c.name;
    const Bytes theme_before = ThemeBytes(wb);
    const std::size_t parts_before = wb.passthrough_parts().size();
    auto result = Apply(wb, c.bytes);
    ASSERT_FALSE(static_cast<bool>(result)) << c.name;
    EXPECT_EQ(result.error().code, c.code) << c.name;
    EXPECT_EQ(ThemeBytes(wb), theme_before) << c.name;
    EXPECT_EQ(wb.passthrough_parts().size(), parts_before) << c.name;
  }
}

TEST(ThemePackage, RejectsNullAndEmptyInput) {
  Workbook wb = Workbook::create();
  const Bytes pkg = BuildPackage(PackageSpec());
  auto null_result = apply_theme_package(wb, nullptr, pkg.size());
  ASSERT_FALSE(static_cast<bool>(null_result));
  EXPECT_EQ(null_result.error().code, FormulonErrorCode::kBindingNullPointer);
  auto empty_result = apply_theme_package(wb, pkg.data(), 0U);
  ASSERT_FALSE(static_cast<bool>(empty_result));
  EXPECT_EQ(empty_result.error().code, FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(wb.load_theme().source, ThemeSource::kDefault);
}

TEST(ThemePackage, FontChangeDirtiesFormulasAndColourChangeDoesNot) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=1+1")));
  const eval::CellNodeId cell{0U, 1U, 0U};
  const auto settle = [&wb, &cell]() {
    ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    ASSERT_FALSE(wb.recalc_engine().awaits_recalc(cell));
  };
  ASSERT_TRUE(static_cast<bool>(Apply(wb, BuildPackage(PackageSpec()))));
  settle();

  PackageSpec colours;
  colours.theme = ThemeXml("445566", "Minor One");
  ASSERT_TRUE(static_cast<bool>(Apply(wb, BuildPackage(colours))));
  EXPECT_FALSE(wb.recalc_engine().awaits_recalc(cell));

  PackageSpec fonts;
  fonts.theme = ThemeXml("445566", "Minor Two");
  ASSERT_TRUE(static_cast<bool>(Apply(wb, BuildPackage(fonts))));
  EXPECT_TRUE(wb.recalc_engine().awaits_recalc(cell));
}

TEST(ThemePackage, SettersDirtyOnlyWhenFontsChange) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=1+1")));
  const eval::CellNodeId cell{0U, 1U, 0U};
  ASSERT_TRUE(static_cast<bool>(Apply(wb, BuildPackage(PackageSpec()))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ThemeFonts same = wb.load_theme().theme.fonts;
  ASSERT_TRUE(static_cast<bool>(wb.set_theme_fonts(same)));
  EXPECT_FALSE(wb.recalc_engine().awaits_recalc(cell));
  ThemeFonts changed = same;
  changed.minor_latin = "Another Face";
  ASSERT_TRUE(static_cast<bool>(wb.set_theme_fonts(changed)));
  EXPECT_TRUE(wb.recalc_engine().awaits_recalc(cell));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_TRUE(static_cast<bool>(wb.reset_theme()));
  EXPECT_TRUE(wb.recalc_engine().awaits_recalc(cell));
}

TEST(ThemePackage, ApplyDropsDefaultThemeVersion) {
  Workbook wb = Workbook::create();
  wb.set_workbook_pr_xml("<workbookPr defaultThemeVersion=\"166925\" filterPrivacy=\"1\"/>");
  ASSERT_TRUE(static_cast<bool>(Apply(wb, BuildPackage(PackageSpec()))));
  EXPECT_EQ(wb.workbook_pr_xml().find("defaultThemeVersion"), std::string::npos);
  EXPECT_NE(wb.workbook_pr_xml().find("filterPrivacy"), std::string::npos);

  Workbook fresh = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(Apply(fresh, BuildPackage(PackageSpec()))));
  EXPECT_EQ(fresh.workbook_pr_xml(), "<workbookPr/>");
  auto saved = write_ooxml(fresh);
  ASSERT_TRUE(static_cast<bool>(saved));
  const std::string workbook_xml = Entry(saved.value(), "xl/workbook.xml");
  EXPECT_NE(workbook_xml, "<missing>");
  EXPECT_EQ(workbook_xml.find("defaultThemeVersion"), std::string::npos);
}

TEST(ThemePackage, NormalFontFollowsTheNewThemeUiScriptFace) {
  Workbook wb = Workbook::create_for_profile(mac_365_ja_jp_profile());
  const std::string jpan = "\xE3\x83\xA1\xE3\x82\xA4\xE3\x83\xAA\xE3\x82\xAA";  // Meiryo
  PackageSpec with_jpan;
  with_jpan.theme = ThemeXml("112233", "Minor One", jpan);
  ASSERT_TRUE(static_cast<bool>(Apply(wb, BuildPackage(with_jpan))));
  auto saved = write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(saved));
  EXPECT_NE(Entry(saved.value(), "xl/styles.xml").find("<name val=\"" + jpan + "\""), std::string::npos);

  const std::string model_face = wb.styles().fonts[0].name;
  PackageSpec without_jpan;
  without_jpan.theme = ThemeXml("112233", "Minor One");
  ASSERT_TRUE(static_cast<bool>(Apply(wb, BuildPackage(without_jpan))));
  saved = write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(saved));
  const std::string styles = Entry(saved.value(), "xl/styles.xml");
  EXPECT_EQ(styles.find(jpan), std::string::npos);
  EXPECT_NE(styles.find("<name val=\"" + model_face + "\""), std::string::npos);
}

}  // namespace
}  // namespace io
}  // namespace formulon
