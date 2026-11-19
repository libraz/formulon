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
#include <utility>
#include <vector>

#include "cf/cf_types.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
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
