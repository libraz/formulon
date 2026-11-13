// PtgList flag bits against the Excel-written workbooks under tests/fixtures/excel: the writer may set only bits
// some Excel-written PtgList carries, and the reader accepts a PtgList whose unused bit is set.

#include <algorithm>
#include <cstdint>
#include <cstdio>
#include <filesystem>
#include <map>
#include <string>
#include <vector>

#include "cell.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "workbook.h"
#include "writer_test_helpers.h"
#include "zip_repack_helpers.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

constexpr std::uint8_t kPtgList = 0x18;
constexpr std::uint8_t kEptgList = 0x19;
constexpr std::size_t kFlagsAt = 4;  // ptg, eptg, ixti
constexpr std::uint16_t kUnusedBit = 0x0200;
constexpr std::uint16_t kBrtFmlaString = 8;
constexpr std::uint16_t kBrtFmlaNum = 9;
constexpr std::uint16_t kBrtFmlaBool = 10;
constexpr std::uint16_t kBrtFmlaError = 11;
constexpr std::uint16_t kBrtListCCFmla = 351;

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

std::string Fixture(const std::string& relative) {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/" + relative;
}

std::uint32_t U32At(const std::uint8_t* p) {
  return static_cast<std::uint32_t>(p[0]) | (static_cast<std::uint32_t>(p[1]) << 8U) |
         (static_cast<std::uint32_t>(p[2]) << 16U) | (static_cast<std::uint32_t>(p[3]) << 24U);
}

/// Offset of `cce` in a formula record's payload; npos for a record that holds no formula.
std::size_t CceOffset(const FramedRecord& rec) {
  constexpr std::size_t kCell = 8;
  switch (rec.type) {
    case kBrtFmlaNum:
      return kCell + 8U + 2U;
    case kBrtFmlaBool:
    case kBrtFmlaError:
      return kCell + 1U + 2U;
    case kBrtFmlaString:
      return rec.payload.size >= kCell + 4U ? kCell + 4U + 2U * U32At(rec.payload.data + kCell) + 2U
                                            : std::string::npos;
    case kBrtListCCFmla:
      return 1U;
    default:
      return std::string::npos;
  }
}

/// Every PtgList in the cell and calculated-column formulas of `part`, as the payload offset of its first byte.
std::vector<std::size_t> PtgListOffsets(const std::vector<std::uint8_t>& part) {
  std::vector<std::size_t> out;
  for (const FramedRecord& rec : split_records(part)) {
    const std::size_t cce_at = CceOffset(rec);
    if (cce_at == std::string::npos || cce_at + 4U > rec.payload.size) {
      continue;
    }
    const std::size_t begin = cce_at + 4U;
    const std::size_t end = std::min<std::size_t>(rec.payload.size, begin + U32At(rec.payload.data + cce_at));
    const std::size_t base = static_cast<std::size_t>(rec.payload.data - part.data());
    for (std::size_t i = begin; i + kFlagsAt + 2U <= end; ++i) {
      if (rec.payload.data[i] == kPtgList && rec.payload.data[i + 1U] == kEptgList) {
        out.push_back(base + i);
      }
    }
  }
  return out;
}

std::uint16_t FlagsAt(const std::vector<std::uint8_t>& part, std::size_t ptg) {
  return static_cast<std::uint16_t>(part[ptg + kFlagsAt] | (part[ptg + kFlagsAt + 1U] << 8U));
}

/// The PtgList flags of every `.bin` part of `package`.
std::vector<std::uint16_t> PtgListFlags(const std::vector<std::uint8_t>& package) {
  std::vector<std::uint16_t> out;
  ZipReader zip;
  if (!zip.open(SpanOf(package))) {
    return out;
  }
  for (const std::string& name : zip.list_entries()) {
    if (name.size() < 4U || name.compare(name.size() - 4U, 4U, ".bin") != 0) {
      continue;
    }
    auto bytes = zip.read_entry(name);
    if (!bytes) {
      continue;
    }
    for (const std::size_t ptg : PtgListOffsets(bytes.value())) {
      out.push_back(FlagsAt(bytes.value(), ptg));
    }
  }
  return out;
}

std::vector<std::uint8_t> WriteXlsb(const Workbook& wb) {
  auto result = write_xlsb_with_result(wb);
  EXPECT_TRUE(static_cast<bool>(result)) << (result ? "" : result.error().message);
  return result ? std::move(result.value().bytes) : std::vector<std::uint8_t>{};
}

TEST(XlsbPtgListFlags, WriterSetsOnlyBitsExcelWrites) {
  std::uint16_t excel_bits = 0;
  std::size_t excel_lists = 0;
  std::size_t written_lists = 0;
  std::vector<std::uint16_t> written;
  for (const auto& entry : std::filesystem::recursive_directory_iterator(Fixture(""))) {
    if (entry.path().extension() != ".xlsb") {
      continue;
    }
    const std::vector<std::uint8_t> bytes = ReadFile(entry.path().string());
    const std::vector<std::uint16_t> flags = PtgListFlags(bytes);
    for (const std::uint16_t f : flags) {
      excel_bits |= f;
    }
    excel_lists += flags.size();
    if (flags.empty()) {
      continue;
    }
    auto loaded = read_xlsb(SpanOf(bytes));
    ASSERT_TRUE(static_cast<bool>(loaded)) << entry.path();
    const std::vector<std::uint16_t> ours = PtgListFlags(WriteXlsb(loaded.value().workbook));
    written.insert(written.end(), ours.begin(), ours.end());
  }
  const std::vector<std::uint8_t> xlsx = ReadFile(Fixture("table_totals/base.xlsx"));
  auto from_xlsx = read_ooxml(SpanOf(xlsx));
  ASSERT_TRUE(static_cast<bool>(from_xlsx));
  const std::vector<std::uint16_t> converted = PtgListFlags(WriteXlsb(from_xlsx.value().workbook));
  written.insert(written.end(), converted.begin(), converted.end());
  written_lists = written.size();
  ASSERT_GT(excel_lists, 0U);
  ASSERT_GT(written_lists, 0U);
  for (const std::uint16_t f : written) {
    EXPECT_EQ(f & static_cast<std::uint16_t>(~excel_bits), 0U) << std::hex << "flags 0x" << f;
  }
}

TEST(XlsbPtgListFlags, TotalsRowReferenceConvertsToExcelFlags) {
  const std::vector<std::uint8_t> xlsx = ReadFile(Fixture("table_totals/base.xlsx"));
  auto from_xlsx = read_ooxml(SpanOf(xlsx));
  ASSERT_TRUE(static_cast<bool>(from_xlsx));
  const std::vector<std::uint16_t> ours = PtgListFlags(WriteXlsb(from_xlsx.value().workbook));
  const std::vector<std::uint16_t> excel = PtgListFlags(ReadFile(Fixture("table_totals/base_excel.xlsb")));
  ASSERT_EQ(excel.size(), 2U);
  EXPECT_EQ(ours, excel);
}

TEST(XlsbPtgListFlags, ReaderAcceptsTheUnusedBit) {
  const std::vector<std::uint8_t> package = ReadFile(Fixture("table_totals/base_excel.xlsb"));
  std::map<std::string, std::vector<std::uint8_t>> patched;
  for (const char* name : {"xl/worksheets/sheet4.bin", "xl/worksheets/sheet5.bin"}) {
    ZipReader zip;
    ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
    auto part = zip.read_entry(name);
    ASSERT_TRUE(static_cast<bool>(part)) << name;
    std::vector<std::uint8_t> bytes = std::move(part.value());
    const std::vector<std::size_t> lists = PtgListOffsets(bytes);
    ASSERT_EQ(lists.size(), 1U) << name;
    bytes[lists[0] + kFlagsAt + 1U] |= static_cast<std::uint8_t>(kUnusedBit >> 8U);
    patched.emplace(name, std::move(bytes));
  }
  auto plain = read_xlsb(SpanOf(package));
  const std::vector<std::uint8_t> repacked = RepackWithEntries(package, patched);
  auto with_bit = read_xlsb(SpanOf(repacked));
  ASSERT_TRUE(static_cast<bool>(plain));
  ASSERT_TRUE(static_cast<bool>(with_bit));
  for (const std::size_t sheet : {3U, 4U}) {
    const Cell* expected = plain.value().workbook.sheet(sheet).cell_at(4, 2);
    const Cell* actual = with_bit.value().workbook.sheet(sheet).cell_at(4, 2);
    ASSERT_NE(expected, nullptr);
    ASSERT_NE(actual, nullptr);
    EXPECT_EQ(actual->formula_text, expected->formula_text);
    EXPECT_NE(actual->formula_text.find("[h3]"), std::string::npos) << actual->formula_text;
  }
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
