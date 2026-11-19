//
// A Windows Excel workbook with one data bar on A1:A6 whose legacy rule links
// to its x14 payload through a nested `<x14:id>`, with the link and the
// payload's id both edited to lower case
// (`tests/fixtures/excel/cf_x14_linked_id_lower_case.xlsx`), and Excel's own
// re-save of it (`..._excel_saved.xlsx`). Excel opens the lower-case book
// without repair and writes the link and the payload id upper case; the
// reader takes both upper case, so a save writes what Excel writes.

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>
#include <utility>
#include <vector>

#include "cf/cf_types.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "miniz.h"
#include "sheet.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

constexpr const char* kUpperId = "{398D8C1A-AC0D-4EE9-95B0-3F8CC7E59667}";

std::vector<std::uint8_t> ReadFixture(const char* name) {
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/" + name;
  std::vector<std::uint8_t> out;
  FILE* file = std::fopen(path.c_str(), "rb");
  if (file == nullptr) {
    ADD_FAILURE() << "could not open fixture: " << path;
    return out;
  }
  std::fseek(file, 0, SEEK_END);
  const long size = std::ftell(file);
  std::fseek(file, 0, SEEK_SET);
  if (size > 0) {
    out.resize(static_cast<std::size_t>(size));
    if (std::fread(out.data(), 1, out.size(), file) != out.size()) {
      ADD_FAILURE() << "short read on fixture: " << path;
      out.clear();
    }
  }
  std::fclose(file);
  return out;
}

std::string SheetXml(const std::vector<std::uint8_t>& archive) {
  std::string out;
  mz_zip_archive reader{};
  if (mz_zip_reader_init_mem(&reader, archive.data(), archive.size(), 0) == MZ_FALSE) {
    ADD_FAILURE() << "mz_zip_reader_init_mem failed";
    return out;
  }
  std::size_t size = 0;
  void* body = mz_zip_reader_extract_file_to_heap(&reader, "xl/worksheets/sheet1.xml", &size, 0);
  if (body != nullptr) {
    out.assign(static_cast<const char*>(body), size);
    mz_free(body);
  } else {
    ADD_FAILURE() << "no xl/worksheets/sheet1.xml";
  }
  mz_zip_reader_end(&reader);
  return out;
}

// The x14 link texts, then the x14 rule ids, of a sheet part, each in document order.
std::vector<std::string> X14Ids(const std::string& xml) {
  std::vector<std::string> out;
  const auto collect = [&](const std::string& open, const std::string& close) {
    for (std::size_t at = xml.find(open); at != std::string::npos; at = xml.find(open, at)) {
      at += open.size();
      const std::size_t end = xml.find(close, at);
      if (end == std::string::npos) {
        break;
      }
      out.push_back(xml.substr(at, end - at));
    }
  };
  collect("<x14:id>", "</x14:id>");
  collect("<x14:cfRule type=\"dataBar\" id=\"", "\"");
  return out;
}

Workbook Read(const std::vector<std::uint8_t>& bytes) {
  auto read_or = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(read_or)) << (read_or ? "" : read_or.error().message);
  return read_or ? std::move(read_or.value().workbook) : Workbook::create_empty();
}

}  // namespace

TEST(CfX14LinkedIdCaseFixture, ExcelWritesTheLinkAndPayloadIdUpperCase) {
  EXPECT_EQ(X14Ids(SheetXml(ReadFixture("cf_x14_linked_id_lower_case_excel_saved.xlsx"))),
            (std::vector<std::string>{kUpperId, kUpperId}));
}

TEST(CfX14LinkedIdCaseFixture, ALowerCaseLinkIsReadAndSavedUpperCase) {
  const Workbook wb = Read(ReadFixture("cf_x14_linked_id_lower_case.xlsx"));
  ASSERT_GT(wb.sheet_count(), 0U);
  const std::vector<cf::ConditionalFormat>& blocks = wb.sheet(0).conditional_formats();
  ASSERT_EQ(blocks.size(), 1U);
  ASSERT_EQ(blocks[0].rules.size(), 1U);
  EXPECT_EQ(blocks[0].rules[0].id, kUpperId);
  EXPECT_EQ(blocks[0].rules[0].type, cf::RuleType::DataBar);

  auto bytes_or = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or));
  EXPECT_EQ(X14Ids(SheetXml(bytes_or.value())),
            X14Ids(SheetXml(ReadFixture("cf_x14_linked_id_lower_case_excel_saved.xlsx"))));
}

}  // namespace formulon
