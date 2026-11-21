//
// Conditional-format thresholds naming another workbook, saved as xlsb, against Excel's saves (Mac 16.115; the
// source book present or absent gives the same bytes). A cell threshold loads as that book's `#REF!` and is written
// as Excel writes it, a PtgRefErr3d through the link's XTI; Excel also keeps the original coordinates in the
// token's unused bytes, which the engine writes from the `#REF!` text. A defined-name threshold is refused: Excel
// keeps it a live PtgNameX, adds a legacy block for a data bar, and cuts the link part's cache to the cells the
// formulas name, none of which is modelled.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "utils/error.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace {

constexpr std::uint16_t kBrtCfvo14 = 1050U;

std::vector<std::uint8_t> ReadFixture(const std::string& name) {
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/cf_x14_external_threshold/" + name;
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
    out.clear();
  }
  std::fclose(file);
  return out;
}

Expected<std::vector<std::uint8_t>, Error> SaveXlsb(const std::string& fixture) {
  const std::vector<std::uint8_t> bytes = ReadFixture(fixture);
  auto read = read_ooxml(ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(read)) << fixture;
  if (!read) {
    return xlsb::write_xlsb(Workbook::create_empty());
  }
  return xlsb::write_xlsb(read.value().workbook);
}

// The BrtCFVO14 payloads of the first sheet, with the six unused bytes of a PtgRefErr3d threshold cleared.
std::vector<std::vector<std::uint8_t>> Cfvo14Records(const std::vector<std::uint8_t>& package) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(ByteSpan{package.data(), package.size()})));
  auto entry = zip.read_entry("xl/worksheets/sheet1.bin");
  EXPECT_TRUE(static_cast<bool>(entry));
  std::vector<std::vector<std::uint8_t>> out;
  if (!entry) {
    return out;
  }
  // FRT header, formula count, item flags, cce and cb come before the rgce.
  constexpr std::size_t kRgceAt = 20U;
  constexpr std::uint8_t kPtgRefErr3d = 0x5CU;
  constexpr std::size_t kUnusedAt = kRgceAt + 3U;
  constexpr std::size_t kUnusedBytes = 6U;
  for (const xlsb::FramedRecord& rec : xlsb::split_records(entry.value())) {
    if (rec.type != kBrtCfvo14) {
      continue;
    }
    std::vector<std::uint8_t> payload(rec.payload.data, rec.payload.data + rec.payload.size);
    if (payload.size() > kUnusedAt + kUnusedBytes && payload[kRgceAt] == kPtgRefErr3d) {
      std::fill_n(payload.begin() + kUnusedAt, kUnusedBytes, std::uint8_t{0});
    }
    out.push_back(std::move(payload));
  }
  return out;
}

TEST(CfX14ExternalThreshold, CellThresholdOnAnotherBookIsWrittenAsExcelsRefError) {
  const std::vector<std::vector<std::uint8_t>> excel = Cfvo14Records(ReadFixture("ext_cell_excel_saved.xlsb"));
  ASSERT_EQ(excel.size(), 2U);
  for (const char* fixture : {"ext_cell_authored.xlsx", "ext_cell_excel_saved.xlsx"}) {
    auto saved = SaveXlsb(fixture);
    ASSERT_TRUE(static_cast<bool>(saved)) << fixture;
    EXPECT_EQ(Cfvo14Records(saved.value()), excel) << fixture;
  }
}

TEST(CfX14ExternalThreshold, NameThresholdOnAnotherBookRefusesTheXlsbSave) {
  for (const char* fixture : {"ext_name_authored.xlsx", "ext_name_excel_saved.xlsx"}) {
    auto saved = SaveXlsb(fixture);
    ASSERT_FALSE(static_cast<bool>(saved)) << fixture;
    EXPECT_EQ(saved.error().message, "XLSB conditional format threshold naming another workbook has no x14 record form")
        << fixture;
  }
}

}  // namespace
}  // namespace io
}  // namespace formulon
