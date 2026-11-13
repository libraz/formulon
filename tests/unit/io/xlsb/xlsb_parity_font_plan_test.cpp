// The save-time font plan held against Excel 16.115's own xlsx and xlsb saves of the same sources under each of the
// fourteen Mac UI locales (tests/fixtures/excel/xlsb_parity/phonfont): a font linked to the theme is written with the
// theme's face for the locale's script.

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <regex>
#include <string>
#include <vector>

#include "cell_style_save_plan.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
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
constexpr std::size_t kBrtFontNameOffset = 21;

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

/// The `<name val>` of every `<font>` in the `<fonts>` list of an xlsx package.
std::vector<std::string> XlsxFontNames(const Bytes& package) {
  const Bytes part = Part(package, "xl/styles.xml");
  const std::string xml(part.begin(), part.end());
  const std::size_t begin = xml.find("<fonts");
  const std::size_t end = xml.find("</fonts>");
  std::vector<std::string> out;
  if (begin == std::string::npos || end == std::string::npos) {
    return out;
  }
  const std::string fonts = xml.substr(begin, end - begin);
  const std::regex name("<name val=\"([^\"]*)\"/>");
  for (auto it = std::sregex_iterator(fonts.begin(), fonts.end(), name); it != std::sregex_iterator(); ++it) {
    out.push_back((*it)[1].str());
  }
  return out;
}

/// The name of every BrtFont of an xlsb package, as UTF-8.
std::vector<std::string> XlsbFontNames(const Bytes& package) {
  std::vector<std::string> out;
  for (const FramedRecord& rec : split_records(Part(package, "xl/styles.bin"))) {
    if (rec.type != kBrtFont || rec.payload.size < kBrtFontNameOffset + 4U) {
      continue;
    }
    const std::uint8_t* p = rec.payload.data + kBrtFontNameOffset;
    const auto byte = [p](std::size_t at) { return static_cast<std::uint32_t>(p[at]); };
    const std::uint32_t chars = byte(0) | (byte(1) << 8U) | (byte(2) << 16U) | (byte(3) << 24U);
    std::u16string text;
    for (std::uint32_t i = 0; i < chars; ++i) {
      text.push_back(static_cast<char16_t>(byte(4U + 2U * i) | (byte(5U + 2U * i) << 8U)));
    }
    std::string utf8;
    for (char16_t c : text) {
      if (c < 0x80U) {
        utf8.push_back(static_cast<char>(c));
      } else if (c < 0x800U) {
        utf8.push_back(static_cast<char>(0xC0U | (c >> 6U)));
        utf8.push_back(static_cast<char>(0x80U | (c & 0x3FU)));
      } else {
        utf8.push_back(static_cast<char>(0xE0U | (c >> 12U)));
        utf8.push_back(static_cast<char>(0x80U | ((c >> 6U) & 0x3FU)));
        utf8.push_back(static_cast<char>(0x80U | (c & 0x3FU)));
      }
    }
    out.push_back(utf8);
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

// Our fonts are the model's; Excel appends the phonetic font after them, so the comparison covers the model's count.
TEST(XlsbParityFontPlan, FacesMatchExcel) {
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

      const std::vector<std::string> ours = XlsxFontNames(xlsx.value());
      std::vector<std::string> excel = XlsxFontNames(ReadFile(Fixture(stem + ".xlsx")));
      ASSERT_EQ(ours.size(), wb.styles().fonts.size());
      ASSERT_GE(excel.size(), ours.size());
      excel.resize(ours.size());
      EXPECT_EQ(ours, excel);
      std::vector<std::string> excel_bin = XlsbFontNames(ReadFile(Fixture(stem + ".xlsb")));
      ASSERT_GE(excel_bin.size(), ours.size());
      excel_bin.resize(ours.size());
      EXPECT_EQ(XlsbFontNames(xlsb.value().bytes), excel_bin);
    }
  }
}

// A new book of every measured profile already stores the faces the plan would write.
TEST(XlsbParityFontPlan, NewWorkbookIsIdentity) {
  for (const ::formulon::detail::ExcelProfileIdEntry& entry : ::formulon::detail::kExcelProfileIds) {
    const Workbook wb = Workbook::create_for_profile(entry.profile);
    EXPECT_TRUE(build_font_save_plan(wb.styles(), wb.load_theme().theme, wb.excel_profile()).is_identity()) << entry.id;
  }
  const Workbook plain = Workbook::create();
  EXPECT_TRUE(build_font_save_plan(plain.styles(), plain.load_theme().theme, plain.excel_profile()).is_identity());
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
