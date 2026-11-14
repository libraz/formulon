// The save-time font plan held against Excel 16.115's own xlsx and xlsb saves of the same sources under each of the
// fourteen Mac UI locales (tests/fixtures/excel/xlsb_parity/phonfont): a font linked to the theme is written with the
// theme's face for the locale's script, and a sheet without phonetic settings names the locale's phonetic font.

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <regex>
#include <string>
#include <utility>
#include <vector>

#include "cell_style_save_plan.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "new_workbook_defaults.h"
#include "workbook.h"
#include "writer_test_helpers.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

constexpr std::uint16_t kBrtFont = 43;
constexpr std::uint16_t kBrtPhoneticInfo = 537;

using Bytes = std::vector<std::uint8_t>;

Bytes ReadFile(const std::string& path) {
  Bytes out;
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
  }
  std::fclose(f);
  return out;
}

std::string Fixture(const std::string& name) {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_parity/phonfont/" + name;
}

Bytes Part(const Bytes& package, const std::string& path) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto bytes = zip.read_entry(path);
  EXPECT_TRUE(static_cast<bool>(bytes)) << path;
  return bytes ? std::move(bytes.value()) : Bytes{};
}

/// The theme-linked `<font>` elements of an xlsx package's `<fonts>` list, as written. A literal font is outside the
/// plan, and Excel completes its family from the host's installed face.
std::vector<std::string> XlsxFonts(const Bytes& package) {
  const Bytes part = Part(package, "xl/styles.xml");
  const std::string xml(part.begin(), part.end());
  const std::size_t begin = xml.find("<fonts");
  const std::size_t end = xml.find("</fonts>");
  std::vector<std::string> out;
  if (begin == std::string::npos || end == std::string::npos) {
    return out;
  }
  const std::string fonts = xml.substr(begin, end - begin);
  const std::regex font("<font>(.*?)</font>");
  for (auto it = std::sregex_iterator(fonts.begin(), fonts.end(), font); it != std::sregex_iterator(); ++it) {
    if ((*it)[1].str().find("<scheme ") != std::string::npos) {
      out.push_back((*it)[1].str());
    }
  }
  return out;
}

/// The `<phoneticPr>` element of each worksheet part, "" for a sheet without one.
std::vector<std::string> XlsxSheetPhonetics(const Bytes& package, std::size_t sheets) {
  std::vector<std::string> out;
  const std::regex phonetic("<phoneticPr[^>]*/>");
  for (std::size_t i = 1; i <= sheets; ++i) {
    const Bytes part = Part(package, "xl/worksheets/sheet" + std::to_string(i) + ".xml");
    const std::string xml(part.begin(), part.end());
    std::smatch match;
    out.push_back(std::regex_search(xml, match, phonetic) ? match.str() : std::string());
  }
  return out;
}

/// The payloads of every record of `type` in an xlsb part.
std::vector<Bytes> XlsbRecords(const Bytes& package, const std::string& path, std::uint16_t type) {
  std::vector<Bytes> out;
  const Bytes part = Part(package, path);
  for (const FramedRecord& rec : split_records(part)) {
    if (rec.type == type) {
      out.emplace_back(rec.payload.data, rec.payload.data + rec.payload.size);
    }
  }
  return out;
}

/// The theme-linked `BrtFont` payloads of an xlsb package (bFontScheme, byte 20, non-zero).
std::vector<Bytes> XlsbThemeFonts(const Bytes& package) {
  constexpr std::size_t kSchemeOffset = 20;
  std::vector<Bytes> out;
  for (Bytes& font : XlsbRecords(package, "xl/styles.bin", kBrtFont)) {
    if (font.size() > kSchemeOffset && font[kSchemeOffset] != 0U) {
      out.push_back(std::move(font));
    }
  }
  return out;
}

/// Each worksheet's `BrtPhoneticInfo` payload, empty for a sheet without one.
std::vector<Bytes> XlsbSheetPhonetics(const Bytes& package, std::size_t sheets) {
  std::vector<Bytes> out;
  for (std::size_t i = 1; i <= sheets; ++i) {
    std::vector<Bytes> records =
        XlsbRecords(package, "xl/worksheets/sheet" + std::to_string(i) + ".bin", kBrtPhoneticInfo);
    out.push_back(records.empty() ? Bytes{} : std::move(records.front()));
  }
  return out;
}

struct LocaleCase {
  const char* suffix;
  ExcelProfile profile;
  std::vector<const char*> cases;
};

std::vector<LocaleCase> Measured() {
  const std::vector<const char*> sweep = {"base", "no_theme", "major", "minor_second", "arial_font0"};
  const std::vector<const char*> first = {"base",       "numonly", "three_fonts", "match_existing", "no_theme",
                                          "two_sheets", "mixed",   "arial_font0", "major",          "minor_second"};
  const auto mac = [](ExcelLocale locale) { return ExcelProfile{ExcelHost::kMac365, locale}; };
  return {
      {"ja", mac(ExcelLocale::kJaJP), first},   {"en", mac(ExcelLocale::kEnUS), first},
      {"zh", mac(ExcelLocale::kZhCN), first},   {"de", mac(ExcelLocale::kDeDE), sweep},
      {"fr", mac(ExcelLocale::kFrFR), sweep},   {"ko", mac(ExcelLocale::kKoKR), sweep},
      {"th", mac(ExcelLocale::kThTH), sweep},   {"ru", mac(ExcelLocale::kRuRU), sweep},
      {"zhTW", mac(ExcelLocale::kZhTW), sweep}, {"it", mac(ExcelLocale::kItIT), sweep},
      {"nl", mac(ExcelLocale::kNlNL), sweep},   {"es", mac(ExcelLocale::kEsES), sweep},
      {"esMX", mac(ExcelLocale::kEsMX), sweep}, {"ptBR", mac(ExcelLocale::kPtBR), sweep},
  };
}

// A literal font 0 makes Excel add a literal phonetic font with a host-installed face, which the save leaves out.
bool HostFacePhoneticFont(const std::string& name, const ExcelProfile& profile) {
  return name == "arial_font0" && new_workbook_defaults(profile.locale).phonetic_size != 0U;
}

TEST(XlsbParityFontPlan, FontsAndSheetPhoneticsMatchExcel) {
  for (const LocaleCase& locale : Measured()) {
    for (const char* name : locale.cases) {
      const std::string stem = std::string("pf_") + name + "_" + locale.suffix;
      SCOPED_TRACE(stem);
      auto loaded = read_ooxml(SpanOf(ReadFile(Fixture(stem + "_src.xlsx"))));
      ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
      Workbook& wb = loaded.value().workbook;
      wb.set_excel_profile(locale.profile);
      auto xlsx = write_ooxml(wb);
      ASSERT_TRUE(static_cast<bool>(xlsx)) << xlsx.error().message;
      auto xlsb = write_xlsb_with_result(wb);
      ASSERT_TRUE(static_cast<bool>(xlsb)) << xlsb.error().message;
      const Bytes excel_xlsx = ReadFile(Fixture(stem + ".xlsx"));
      const Bytes excel_xlsb = ReadFile(Fixture(stem + ".xlsb"));
      const std::size_t sheets = wb.sheet_count();

      const std::vector<std::string> fonts = XlsxFonts(excel_xlsx);
      std::vector<std::string> phonetics = XlsxSheetPhonetics(excel_xlsx, sheets);
      const std::vector<Bytes> bin_fonts = XlsbThemeFonts(excel_xlsb);
      std::vector<Bytes> bin_phonetics = XlsbSheetPhonetics(excel_xlsb, sheets);
      if (HostFacePhoneticFont(name, locale.profile)) {
        phonetics.assign(sheets, std::string());
        bin_phonetics.assign(sheets, Bytes{});
      }
      EXPECT_EQ(XlsxFonts(xlsx.value()), fonts);
      EXPECT_EQ(XlsxSheetPhonetics(xlsx.value(), sheets), phonetics);
      EXPECT_EQ(XlsbThemeFonts(xlsb.value().bytes), bin_fonts);
      EXPECT_EQ(XlsbSheetPhonetics(xlsb.value().bytes, sheets), bin_phonetics);
    }
  }
}

// A new book of every measured profile already stores the faces the plan would write.
TEST(XlsbParityFontPlan, NewWorkbookIsIdentity) {
  for (const ::formulon::detail::ExcelProfileIdEntry& entry : ::formulon::detail::kExcelProfileIds) {
    Workbook wb = Workbook::create_for_profile(entry.profile);
    wb.add_sheet("Second");
    EXPECT_TRUE(build_font_save_plan(wb).is_identity()) << entry.id;
  }
  EXPECT_TRUE(build_font_save_plan(Workbook::create()).is_identity());
}

/// Every `<font>` element of an xlsx package's `<fonts>` list, as written.
std::vector<std::string> XlsxAllFonts(const Bytes& package) {
  const Bytes part = Part(package, "xl/styles.xml");
  const std::string xml(part.begin(), part.end());
  const std::size_t begin = xml.find("<fonts");
  const std::size_t end = xml.find("</fonts>");
  std::vector<std::string> out;
  if (begin == std::string::npos || end == std::string::npos) {
    return out;
  }
  const std::string fonts = xml.substr(begin, end - begin);
  const std::regex font("<font>(.*?)</font>");
  for (auto it = std::sregex_iterator(fonts.begin(), fonts.end(), font); it != std::sregex_iterator(); ++it) {
    out.push_back((*it)[1].str());
  }
  return out;
}

std::vector<std::string> Literal(const std::vector<std::string>& fonts) {
  std::vector<std::string> out;
  for (const std::string& font : fonts) {
    if (font.find("<scheme ") == std::string::npos) {
      out.push_back(font);
    }
  }
  return out;
}

/// The literal `BrtFont` payloads of an xlsb package (bFontScheme, byte 20, zero).
std::vector<Bytes> XlsbLiteralFonts(const Bytes& package) {
  constexpr std::size_t kSchemeOffset = 20;
  std::vector<Bytes> out;
  for (Bytes& font : XlsbRecords(package, "xl/styles.bin", kBrtFont)) {
    if (font.size() > kSchemeOffset && font[kSchemeOffset] == 0U) {
      out.push_back(std::move(font));
    }
  }
  return out;
}

// A literal font that names a face and carries neither family nor charset is saved with the ones Excel fills in, in
// every measured locale.
TEST(XlsbParityFontPlan, LiteralFontsMatchExcel) {
  for (const LocaleCase& locale : Measured()) {
    for (const char* name : locale.cases) {
      if (std::string(name) != "three_fonts" && std::string(name) != "arial_font0") {
        continue;
      }
      const std::string stem = std::string("pf_") + name + "_" + locale.suffix;
      SCOPED_TRACE(stem);
      auto loaded = read_ooxml(SpanOf(ReadFile(Fixture(stem + "_src.xlsx"))));
      ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
      Workbook& wb = loaded.value().workbook;
      wb.set_excel_profile(locale.profile);
      auto xlsx = write_ooxml(wb);
      ASSERT_TRUE(static_cast<bool>(xlsx)) << xlsx.error().message;
      auto xlsb = write_xlsb_with_result(wb);
      ASSERT_TRUE(static_cast<bool>(xlsb)) << xlsb.error().message;

      std::vector<std::string> fonts = Literal(XlsxAllFonts(ReadFile(Fixture(stem + ".xlsx"))));
      std::vector<Bytes> bin_fonts = XlsbLiteralFonts(ReadFile(Fixture(stem + ".xlsb")));
      if (HostFacePhoneticFont(name, locale.profile)) {
        // Excel's appended phonetic font names a host-installed face the save leaves out.
        fonts.pop_back();
        bin_fonts.pop_back();
      }
      EXPECT_EQ(Literal(XlsxAllFonts(xlsx.value())), fonts);
      EXPECT_EQ(XlsbLiteralFonts(xlsb.value().bytes), bin_fonts);
    }
  }
}

std::string FontFamFixture(const std::string& name) {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/fontfam/" + name;
}

// Sixteen literal faces without family or charset, saved by Excel: a known face gets its family and charset, an alias
// its canonical name, and an unknown face nothing. The model keeps what was loaded.
TEST(XlsbParityFontPlan, LiteralFamilyCompletionMatchesExcel) {
  auto loaded = read_ooxml(SpanOf(ReadFile(FontFamFixture("literal_no_family_src.xlsx"))));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  Workbook& wb = loaded.value().workbook;
  wb.set_excel_profile(ExcelProfile{ExcelHost::kMac365, ExcelLocale::kJaJP});
  const std::vector<FontRecord> before = wb.styles().fonts;
  auto xlsx = write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(xlsx)) << xlsx.error().message;
  auto xlsb = write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(xlsb)) << xlsb.error().message;

  EXPECT_EQ(Literal(XlsxAllFonts(xlsx.value())),
            Literal(XlsxAllFonts(ReadFile(FontFamFixture("literal_no_family.xlsx")))));
  EXPECT_EQ(XlsbLiteralFonts(xlsb.value().bytes), XlsbLiteralFonts(ReadFile(FontFamFixture("literal_no_family.xlsb"))));

  ASSERT_EQ(wb.styles().fonts.size(), before.size());
  for (std::size_t i = 0; i < before.size(); ++i) {
    EXPECT_EQ(wb.styles().fonts[i].name, before[i].name) << i;
    EXPECT_EQ(wb.styles().fonts[i].has_family, before[i].has_family) << i;
    EXPECT_EQ(wb.styles().fonts[i].has_charset, before[i].has_charset) << i;
  }

  // A reloaded save is completed already, so a second save writes the same fonts.
  auto reloaded = read_ooxml(SpanOf(xlsx.value()));
  ASSERT_TRUE(static_cast<bool>(reloaded)) << reloaded.error().message;
  auto again = write_ooxml(reloaded.value().workbook);
  ASSERT_TRUE(static_cast<bool>(again)) << again.error().message;
  EXPECT_EQ(Literal(XlsxAllFonts(again.value())), Literal(XlsxAllFonts(xlsx.value())));
}

// Only a font carrying neither element is completed; a partial one, a theme-linked one and an unknown face are
// written as stored.
TEST(XlsbParityFontPlan, LiteralFamilyCompletionNeedsBothAbsent) {
  Workbook wb = Workbook::create();
  std::vector<FontRecord>& fonts = wb.mutable_styles().fonts;
  FontRecord family_only;
  family_only.name = "Arial";
  family_only.has_family = true;
  family_only.family = 0;
  FontRecord charset_only;
  charset_only.name = "Malgun Gothic";
  charset_only.has_charset = true;
  charset_only.charset = 0;
  FontRecord unknown;
  unknown.name = "Aptos";
  FontRecord themed;
  themed.name = "Arial";
  themed.scheme = 2;
  const std::size_t first = fonts.size();
  fonts.push_back(family_only);
  fonts.push_back(charset_only);
  fonts.push_back(unknown);
  const FontSavePlan plan = build_font_save_plan(wb);
  ASSERT_GE(plan.fonts().size(), first + 3);
  EXPECT_TRUE(plan.fonts()[first].has_family);
  EXPECT_EQ(plan.fonts()[first].family, 0U);
  EXPECT_FALSE(plan.fonts()[first].has_charset);
  EXPECT_FALSE(plan.fonts()[first + 1].has_family);
  EXPECT_TRUE(plan.fonts()[first + 1].has_charset);
  EXPECT_EQ(plan.fonts()[first + 1].charset, 0U);
  EXPECT_FALSE(plan.fonts()[first + 2].has_family);
  EXPECT_FALSE(plan.fonts()[first + 2].has_charset);

  fonts.push_back(themed);
  const FontSavePlan themed_plan = build_font_save_plan(wb);
  EXPECT_FALSE(themed_plan.fonts()[first + 3].has_family);
  EXPECT_FALSE(themed_plan.fonts()[first + 3].has_charset);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
