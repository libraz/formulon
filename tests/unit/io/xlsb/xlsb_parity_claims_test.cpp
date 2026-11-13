// Writer bytes that earlier rested on an unchecked belief about what Excel writes, each now held against an xlsb
// Excel 16.115 saved from the same source (tests/fixtures/excel/xlsb_parity/claims).

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "writer_test_helpers.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

constexpr std::uint16_t kBrtBeginList = 343;
// rfxList (16), lt, idList, crwHeader, crwTotals, flags, five dxf ids, dwConnID: the fields before stName.
constexpr std::size_t kListNamesAt = 64;

std::vector<std::uint8_t> ReadFile(const std::string& path) {
  std::vector<std::uint8_t> out;
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

std::string Claim(const std::string& name) {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_parity/claims/" + name;
}

std::vector<std::uint8_t> Converted(const std::string& stem) {
  auto loaded = read_ooxml(SpanOf(ReadFile(Claim(stem + "_src.xlsx"))));
  EXPECT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  if (!loaded) {
    return {};
  }
  auto written = write_xlsb_with_result(loaded.value().workbook);
  EXPECT_TRUE(static_cast<bool>(written)) << written.error().message;
  return written ? std::move(written.value().bytes) : std::vector<std::uint8_t>{};
}

std::vector<std::uint8_t> Part(const std::vector<std::uint8_t>& package, const std::string& path) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto bytes = zip.read_entry(path);
  EXPECT_TRUE(static_cast<bool>(bytes)) << path;
  return bytes ? std::move(bytes.value()) : std::vector<std::uint8_t>{};
}

/// BrtBeginList from stName through stDisplayName, which follows it.
std::vector<std::uint8_t> ListNames(const std::vector<std::uint8_t>& part) {
  for (const FramedRecord& rec : split_records(part)) {
    if (rec.type == kBrtBeginList && rec.payload.size > kListNamesAt) {
      return std::vector<std::uint8_t>(rec.payload.data + kListNamesAt, rec.payload.data + rec.payload.size);
    }
  }
  ADD_FAILURE() << "no BrtBeginList";
  return {};
}

// Excel writes stName in BrtBeginList whether or not the table's name differs from its display name.
TEST(XlsbParityClaims, TableNameIsWrittenLikeExcel) {
  const std::vector<std::uint8_t> ours = Converted("stname_diff");
  const std::vector<std::uint8_t> excel = ReadFile(Claim("stname_diff.xlsb"));
  for (const char* part : {"xl/tables/table1.bin", "xl/tables/table2.bin"}) {
    SCOPED_TRACE(part);
    EXPECT_EQ(ListNames(Part(ours, part)), ListNames(Part(excel, part)));
  }
}

// The legacy conditional-format rule records (cbFmla included: cce plus 2 per reference token, absolute or relative)
// equal Excel's.
TEST(XlsbParityClaims, LegacyCfRuleRecordsMatchExcel) {
  constexpr std::uint16_t kBrtBeginCFRule = 463;
  const auto rules = [](const std::vector<std::uint8_t>& part) {
    std::vector<std::vector<std::uint8_t>> out;
    for (const FramedRecord& rec : split_records(part)) {
      if (rec.type == kBrtBeginCFRule) {
        out.emplace_back(rec.payload.data, rec.payload.data + rec.payload.size);
      }
    }
    return out;
  };
  const auto ours = rules(Part(Converted("cbfmla_absolute"), "xl/worksheets/sheet1.bin"));
  const auto excel = rules(Part(ReadFile(Claim("cbfmla_absolute.xlsb")), "xl/worksheets/sheet1.bin"));
  ASSERT_FALSE(excel.empty());
  EXPECT_EQ(ours, excel);
}

// A legacy cell formula whose `@` stands where Excel's own intersection does is written without it, so the workbook
// carries no hidden SINGLE name for it; Excel's names part has the defined name alone.
TEST(XlsbParityClaims, ImpliedIntersectionDeclaresNoHiddenName) {
  constexpr std::uint16_t kBrtName = 39;
  const auto names = [](const std::vector<std::uint8_t>& part) {
    std::vector<std::vector<std::uint8_t>> out;
    for (const FramedRecord& rec : split_records(part)) {
      if (rec.type == kBrtName) {
        out.emplace_back(rec.payload.data, rec.payload.data + rec.payload.size);
      }
    }
    return out;
  };
  const auto ours = names(Part(Converted("name_scope_zero"), "xl/workbook.bin"));
  const auto excel = names(Part(ReadFile(Claim("name_scope_zero.xlsb")), "xl/workbook.bin"));
  ASSERT_EQ(excel.size(), 1U);
  EXPECT_EQ(ours, excel);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
