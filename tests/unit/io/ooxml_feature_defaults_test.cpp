//
// Absent-attribute defaults of `<iconSet>` and `<dataValidation>`, pinned by
// Excel-saved .xlsx files (`tests/fixtures/excel/xlsb_feature_*.xlsx`) whose
// .xlsb twins state the same fields as explicit bits. Excel omits
// `iconSet` for 3TrafficLights1 and `allowBlank` when it is off, so the
// reader must default to those values and the writer must omit them the
// same way.

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>
#include <vector>

#include "cf/cf_types.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/zip_reader.h"
#include "sheet.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

std::vector<std::uint8_t> ReadFileBytes(const std::string& path) {
  std::vector<std::uint8_t> out;
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

Workbook ReadXlsx(const std::string& name) {
  const std::vector<std::uint8_t> bytes =
      ReadFileBytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_feature_" + name + ".xlsx");
  auto result = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(result)) << (result ? "" : result.error().message);
  return result ? std::move(result.value().workbook) : Workbook::create_empty();
}

std::string PartXml(const Workbook& wb, const char* part) {
  auto saved = io::write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
  if (!saved) {
    return {};
  }
  io::ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(io::ByteSpan{saved.value().data(), saved.value().size()})));
  auto bytes = zip.read_entry(part);
  EXPECT_TRUE(static_cast<bool>(bytes));
  return bytes ? std::string(bytes.value().begin(), bytes.value().end()) : std::string();
}

std::string SheetXml(const Workbook& wb) {
  auto saved = io::write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
  if (!saved) {
    return {};
  }
  io::ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(io::ByteSpan{saved.value().data(), saved.value().size()})));
  auto part = zip.read_entry("xl/worksheets/sheet1.xml");
  EXPECT_TRUE(static_cast<bool>(part));
  return part ? std::string(part.value().begin(), part.value().end()) : std::string();
}

TEST(OoxmlFeatureDefaults, BareIconSetIsThreeTrafficLights) {
  const Workbook wb = ReadXlsx("base");
  const cf::CFRule& rule = wb.sheet(0).conditional_formats().front().rules.back();
  ASSERT_TRUE(rule.icon_set.has_value());
  EXPECT_EQ(rule.icon_set->name, cf::IconSetName::Three_TrafficLights1);
  EXPECT_NE(SheetXml(wb).find("<iconSet><cfvo"), std::string::npos);
}

TEST(OoxmlFeatureDefaults, AbsentAllowBlankIsOff) {
  const Workbook wb = ReadXlsx("base");
  const std::vector<DataValidation>& dvs = wb.sheet(0).validations();
  ASSERT_EQ(dvs.size(), 4U);
  EXPECT_TRUE(dvs[0].allow_blank);  // allowBlank="1"
  EXPECT_FALSE(dvs[1].allow_blank);
  EXPECT_FALSE(dvs[3].allow_blank);
  const std::string xml = SheetXml(wb);
  EXPECT_EQ(xml.find("allowBlank=\"0\""), std::string::npos);
  EXPECT_NE(xml.find("allowBlank=\"1\""), std::string::npos);
}

// A dxf only changes what it states, so a dxf font without <sz> or <color>
// must not gain either on save: Excel keeps a stated size and applies it,
// and an added colour would force the formatted text black.
TEST(OoxmlFeatureDefaults, DxfFontWithoutASizeIsSavedWithoutOne) {
  const Workbook wb = ReadXlsx("base");
  ASSERT_FALSE(wb.styles().dxfs.empty());
  const std::string styles = PartXml(wb, "xl/styles.xml");
  const std::size_t dxfs = styles.find("<dxfs");
  ASSERT_NE(dxfs, std::string::npos) << styles;
  const std::string block = styles.substr(dxfs, styles.find("</dxfs>") - dxfs);
  EXPECT_NE(block.find("<b/>"), std::string::npos) << block;
  EXPECT_EQ(block.find("<sz"), std::string::npos) << block;
  EXPECT_EQ(block.find("<color"), std::string::npos) << block;
}

TEST(OoxmlFeatureDefaults, DxfFontSizeStatedInTheSourceIsKept) {
  const Workbook wb = ReadXlsx("dxf");
  const std::string styles = PartXml(wb, "xl/styles.xml");
  EXPECT_NE(styles.find("<sz val=\"14\"/>"), std::string::npos);
}

TEST(OoxmlFeatureDefaults, OtherDefaultsMatchExcel) {
  const Workbook wb = ReadXlsx("dv_all");
  const std::vector<DataValidation>& dvs = wb.sheet(0).validations();
  ASSERT_GE(dvs.size(), 35U);
  EXPECT_EQ(dvs[0].op, 0U);           // operator omitted for between
  EXPECT_EQ(dvs[0].error_style, 0U);  // errorStyle omitted for stop
  EXPECT_FALSE(dvs[0].show_input_message);
  EXPECT_FALSE(dvs[0].show_error_message);
  EXPECT_TRUE(dvs[0].show_dropdown);
  EXPECT_EQ(dvs[32].error_style, 2U);
  EXPECT_TRUE(dvs[32].show_error_message);
  EXPECT_TRUE(dvs[33].show_input_message);
  EXPECT_EQ(dvs[33].type, 0U);
}

}  // namespace
}  // namespace formulon
