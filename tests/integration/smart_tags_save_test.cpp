// A save writes no smart tags, as Excel's own save does not: the sheet `<smartTags>`, the workbook
// `smartTagPr` / `smartTagTypes` and the .xlsb smart-tag records are all left out. Measured on Mac Excel 16.115 only.

#include <algorithm>
#include <cstdint>
#include <fstream>
#include <iterator>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "workbook.h"

namespace formulon {
namespace {

std::vector<std::uint8_t> ReadFixture(const std::string& relative) {
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/smart_tags/" + relative;
  std::ifstream in(path, std::ios::binary);
  EXPECT_TRUE(in.good()) << path;
  return std::vector<std::uint8_t>((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());
}

io::ByteSpan SpanOf(const std::vector<std::uint8_t>& bytes) {
  return io::ByteSpan{bytes.data(), bytes.size()};
}

std::vector<std::string> PartNames(const std::vector<std::uint8_t>& package) {
  io::ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  std::vector<std::string> names = zip.list_entries();
  std::sort(names.begin(), names.end());
  return names;
}

std::string PartText(const std::vector<std::uint8_t>& package, const char* path) {
  io::ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto entry_or = zip.read_entry(path);
  EXPECT_TRUE(static_cast<bool>(entry_or)) << path;
  return entry_or ? std::string(entry_or.value().begin(), entry_or.value().end()) : std::string{};
}

/// The sorted relationship types of a .rels part.
std::vector<std::string> RelTypes(const std::string& rels) {
  std::vector<std::string> types;
  const std::string key = "Type=\"";
  for (std::size_t at = rels.find(key); at != std::string::npos; at = rels.find(key, at)) {
    at += key.size();
    types.push_back(rels.substr(at, rels.find('"', at) - at));
  }
  std::sort(types.begin(), types.end());
  return types;
}

std::vector<std::uint16_t> RecordTypes(const std::string& body) {
  std::vector<std::uint16_t> types;
  io::ByteSpan cursor{reinterpret_cast<const std::uint8_t*>(body.data()), body.size()};
  while (cursor.size > 0U) {
    auto record_or = io::xlsb::read_record(cursor);
    if (!record_or) {
      break;
    }
    types.push_back(record_or.value().type);
  }
  return types;
}

Workbook LoadTaggedSource() {
  const std::vector<std::uint8_t> source = ReadFixture("tagged_source.xlsx");
  auto read_or = io::read_ooxml(SpanOf(source));
  EXPECT_TRUE(static_cast<bool>(read_or)) << read_or.error().message;
  return std::move(read_or.value().workbook);
}

TEST(SmartTagsSave, XlsxSaveMatchesExcelsPackage) {
  const Workbook wb = LoadTaggedSource();
  auto saved_or = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(saved_or)) << saved_or.error().message;
  const std::vector<std::uint8_t>& saved = saved_or.value();
  const std::vector<std::uint8_t> excel = ReadFixture("excel_saved.xlsx");

  EXPECT_EQ(PartNames(saved), PartNames(excel));
  EXPECT_EQ(RelTypes(PartText(saved, "xl/_rels/workbook.xml.rels")),
            RelTypes(PartText(excel, "xl/_rels/workbook.xml.rels")));
  EXPECT_EQ(PartText(saved, "xl/workbook.xml").find("smartTag"), std::string::npos);
  EXPECT_EQ(PartText(saved, "xl/worksheets/sheet1.xml").find("smartTags"), std::string::npos);
}

TEST(SmartTagsSave, XlsbSaveMatchesExcelsPackage) {
  const Workbook wb = LoadTaggedSource();
  auto saved_or = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(saved_or)) << saved_or.error().message;
  const std::vector<std::uint8_t>& saved = saved_or.value();
  const std::vector<std::uint8_t> excel = ReadFixture("excel_saved.xlsb");

  // Excel's package also carries its cell index (binaryIndex1.bin and the sheet rels addressing it), which the
  // writer never emits; every other part matches.
  std::vector<std::string> excel_parts = PartNames(excel);
  excel_parts.erase(std::remove_if(excel_parts.begin(), excel_parts.end(),
                                   [](const std::string& name) {
                                     return name == "xl/worksheets/binaryIndex1.bin" ||
                                            name == "xl/worksheets/_rels/sheet1.bin.rels";
                                   }),
                    excel_parts.end());
  EXPECT_EQ(PartNames(saved), excel_parts);
  EXPECT_EQ(RelTypes(PartText(saved, "xl/_rels/workbook.bin.rels")),
            RelTypes(PartText(excel, "xl/_rels/workbook.bin.rels")));
  for (const std::uint16_t type : RecordTypes(PartText(saved, "xl/worksheets/sheet1.bin"))) {
    EXPECT_TRUE(type < 589U || type > 595U) << "smart-tag record " << type;
  }
}

}  // namespace
}  // namespace formulon
