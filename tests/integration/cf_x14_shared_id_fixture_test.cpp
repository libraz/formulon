//
// An Excel 365 workbook (`tests/fixtures/excel/cf_x14_shared_id.xlsx`) whose
// sheet holds two data-bar rules stored in the x14 extension alone, on A1:A10
// and B1:B10, under one rule id. Excel opens that state without repair and
// saves it back unchanged; the reader keeps both rules and the writer writes
// both.

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "cf/cf_types.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "miniz.h"
#include "sheet.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

constexpr const char* kSharedId = "{00000000-000E-0000-0000-000001000000}";

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

Workbook Read(const std::vector<std::uint8_t>& bytes) {
  auto read_or = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(read_or)) << (read_or ? "" : read_or.error().message);
  return read_or ? std::move(read_or.value().workbook) : Workbook::create_empty();
}

void ExpectBothRules(const Workbook& wb) {
  ASSERT_GT(wb.sheet_count(), 0U);
  const std::vector<cf::ConditionalFormat>& blocks = wb.sheet(0).conditional_formats();
  ASSERT_EQ(blocks.size(), 2U);
  for (std::size_t b = 0; b < 2U; ++b) {
    ASSERT_EQ(blocks[b].rules.size(), 1U);
    EXPECT_EQ(blocks[b].rules[0].id, kSharedId);
    EXPECT_EQ(blocks[b].rules[0].type, cf::RuleType::DataBar);
    EXPECT_EQ(blocks[b].rules[0].priority, static_cast<std::int32_t>(b + 1U));
    ASSERT_EQ(blocks[b].sqref.size(), 1U);
    EXPECT_EQ(blocks[b].sqref[0].first.col, b);
  }
}

}  // namespace

TEST(CfX14SharedIdFixture, BothRulesSharingAnIdAreRead) {
  ExpectBothRules(Read(ReadFileBytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/cf_x14_shared_id.xlsx")));
}

TEST(CfX14SharedIdFixture, BothRulesSurviveASaveAndReload) {
  const Workbook wb = Read(ReadFileBytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/cf_x14_shared_id.xlsx"));
  auto bytes_or = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or));
  ExpectBothRules(Read(bytes_or.value()));
}

}  // namespace formulon

namespace formulon {
namespace {

constexpr const char* kLinkedIdInBlocks = "{3F2A0C11-0000-4000-8000-0000000000D1}";
constexpr const char* kLinkedIdInOneBlock = "{3F2A0C11-0000-4000-8000-0000000000D2}";

std::string SharedIdFixture(const char* name) {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/cf_shared_id/" + name;
}

std::string ZipEntryText(const std::vector<std::uint8_t>& zip, const std::string& name) {
  mz_zip_archive archive{};
  if (mz_zip_reader_init_mem(&archive, zip.data(), zip.size(), 0) == MZ_FALSE) {
    ADD_FAILURE() << "not a zip package";
    return {};
  }
  std::size_t size = 0;
  void* data = mz_zip_reader_extract_file_to_heap(&archive, name.c_str(), &size, 0);
  std::string out;
  if (data != nullptr) {
    out.assign(static_cast<const char*>(data), size);
    mz_free(data);
  } else {
    ADD_FAILURE() << "missing part " << name;
  }
  mz_zip_reader_end(&archive);
  return out;
}

// The data-bar rules of every sheet, in block and rule order.
std::vector<const cf::CFRule*> DataBars(const Workbook& wb) {
  std::vector<const cf::CFRule*> out;
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    for (const cf::ConditionalFormat& block : wb.sheet(s).conditional_formats()) {
      for (const cf::CFRule& rule : block.rules) {
        if (rule.type == cf::RuleType::DataBar && rule.data_bar.has_value()) {
          out.push_back(&rule);
        }
      }
    }
  }
  return out;
}

Workbook ThroughXlsx(const Workbook& wb) {
  auto bytes_or = io::write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(bytes_or)) << (bytes_or ? "" : bytes_or.error().message);
  return bytes_or ? Read(bytes_or.value()) : Workbook::create_empty();
}

Workbook ThroughXlsb(const Workbook& wb) {
  auto bytes_or = io::xlsb::write_xlsb(wb);
  EXPECT_TRUE(static_cast<bool>(bytes_or)) << (bytes_or ? "" : bytes_or.error().message);
  if (!bytes_or) {
    return Workbook::create_empty();
  }
  auto read_or = io::xlsb::read_xlsb(io::ByteSpan{bytes_or.value().data(), bytes_or.value().size()});
  EXPECT_TRUE(static_cast<bool>(read_or)) << (read_or ? "" : read_or.error().message);
  return read_or ? std::move(read_or.value().workbook) : Workbook::create_empty();
}

// Both bars keep `id`, and both show the first payload in document order (negative fill FF0000) over the legacy
// fill 638EC6.
void ExpectSharedIdBarsShowFirstPayload(const Workbook& wb, const char* id) {
  const std::vector<const cf::CFRule*> bars = DataBars(wb);
  ASSERT_EQ(bars.size(), 2U);
  for (const cf::CFRule* bar : bars) {
    EXPECT_EQ(bar->id, id);
    const cf::DataBarSpec& spec = *bar->data_bar;
    EXPECT_EQ(spec.fill.r, 0x63);
    EXPECT_EQ(spec.fill.g, 0x8E);
    EXPECT_EQ(spec.fill.b, 0xC6);
    EXPECT_EQ(spec.negative_fill.r, 0xFF);
    EXPECT_EQ(spec.negative_fill.g, 0x00);
    EXPECT_EQ(spec.negative_fill.b, 0x00);
  }
}

std::size_t CountOf(const std::string& text, std::string_view needle) {
  std::size_t n = 0;
  for (std::size_t at = text.find(needle); at != std::string::npos; at = text.find(needle, at + needle.size())) {
    ++n;
  }
  return n;
}

// The worksheet part holding the x14 conditional-format blocks, and the priorities of its `<x14:cfRule>` elements
// in file order joined by ",".
std::string X14SheetXml(const std::vector<std::uint8_t>& zip) {
  for (const char* part : {"xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"}) {
    std::string xml = ZipEntryText(zip, part);
    if (xml.find("<x14:conditionalFormatting") != std::string::npos) {
      return xml;
    }
  }
  ADD_FAILURE() << "no worksheet with x14 conditional formats";
  return {};
}

std::string X14Priorities(const std::string& xml) {
  std::string out;
  for (std::size_t at = xml.find("<x14:cfRule"); at != std::string::npos; at = xml.find("<x14:cfRule", at + 1U)) {
    const std::size_t end = xml.find('>', at);
    const std::size_t attr = xml.find("priority=\"", at);
    if (attr == std::string::npos || attr > end) {
      continue;
    }
    const std::size_t first = attr + 10U;
    out += (out.empty() ? "" : ",") + xml.substr(first, xml.find('"', first) - first);
  }
  return out;
}

}  // namespace

// Excel instead keeps the GUID on the last rule only, gives each earlier rule a new one, and writes every payload
// with the last payload's colour. The engine's reader keeps the first payload per id and its xlsb writer emits one
// payload per id (the first rule's), so both bars reload with the first payload's colour and the shared id.
TEST(CfX14SharedIdFixture, SharedIdLinkedBarsKeepTheEngineIdAssignmentOnSave) {
  const Workbook wb = Read(ReadFileBytes(SharedIdFixture("D1.xlsx")));
  ExpectSharedIdBarsShowFirstPayload(wb, kLinkedIdInBlocks);
  ExpectSharedIdBarsShowFirstPayload(ThroughXlsx(wb), kLinkedIdInBlocks);
  ExpectSharedIdBarsShowFirstPayload(ThroughXlsb(wb), kLinkedIdInBlocks);

  auto bytes_or = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or));
  const std::string xml = X14SheetXml(bytes_or.value());
  EXPECT_EQ(CountOf(xml, "<x14:conditionalFormatting "), 2U);
  EXPECT_EQ(CountOf(xml, std::string("<x14:cfRule type=\"dataBar\" id=\"") + kLinkedIdInBlocks + "\">"), 2U);
}

// Excel instead keeps the GUID on the last rule only, gives each earlier rule a new one, and writes every payload
// with the last payload's colour. The engine writes one x14 block for the two rules, both carrying the shared id.
TEST(CfX14SharedIdFixture, SharedIdLinkedBarsInOneBlockKeepTheEngineIdAssignmentOnSave) {
  const Workbook wb = Read(ReadFileBytes(SharedIdFixture("D2.xlsx")));
  ExpectSharedIdBarsShowFirstPayload(wb, kLinkedIdInOneBlock);
  ExpectSharedIdBarsShowFirstPayload(ThroughXlsx(wb), kLinkedIdInOneBlock);
  ExpectSharedIdBarsShowFirstPayload(ThroughXlsb(wb), kLinkedIdInOneBlock);

  auto bytes_or = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or));
  const std::string xml = X14SheetXml(bytes_or.value());
  EXPECT_EQ(CountOf(xml, "<x14:conditionalFormatting "), 1U);
  EXPECT_EQ(CountOf(xml, std::string("<x14:cfRule type=\"dataBar\" id=\"") + kLinkedIdInOneBlock + "\">"), 2U);
}

// Excel instead writes one block holding both rules (ccf 2), with the rules in priority order inside it. The block
// engine writes one block per x14-only rule, in model order.
TEST(CfX14SharedIdFixture, X14OnlyRulesSharingABlockWriteTheEngineBlockLayout) {
  const Workbook wb = Read(ReadFileBytes(SharedIdFixture("B1.xlsx")));
  auto bytes_or = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or));
  const std::string xml = X14SheetXml(bytes_or.value());
  EXPECT_EQ(CountOf(xml, "<x14:conditionalFormatting "), 2U);
  EXPECT_EQ(X14Priorities(xml), "1,2");
}

// Excel instead writes one block holding both rules, saved in priority order 1, 2 although authored 2, 1. The engine
// keeps the authored order: neither reader nor writer sorts x14-only rules.
TEST(CfX14SharedIdFixture, X14OnlyRulesAuthoredOutOfPriorityOrderWriteInAuthoredOrder) {
  const Workbook wb = Read(ReadFileBytes(SharedIdFixture("B2.xlsx")));
  auto bytes_or = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or));
  const std::string xml = X14SheetXml(bytes_or.value());
  EXPECT_EQ(CountOf(xml, "<x14:conditionalFormatting "), 2U);
  EXPECT_EQ(X14Priorities(xml), "2,1");
}

}  // namespace formulon
