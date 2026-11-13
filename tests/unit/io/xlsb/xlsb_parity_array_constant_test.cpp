// Array constants whose elements are text, booleans, errors and numbers, in a cell formula and in a defined name,
// converted from xlsx to xlsb, against the xlsb Excel 16.115 saved from the same source. PtgArray's 14 unused bytes
// are left uninitialised by Excel and are masked; every other byte must match.

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

constexpr std::uint16_t kBrtFmlaString = 8;
constexpr std::uint16_t kBrtFmlaNum = 9;
constexpr std::uint16_t kBrtName = 39;
constexpr std::size_t kPtgArrayUnusedBytes = 14;
constexpr std::uint8_t kPtgFuncVar = 0x22;
constexpr std::uint8_t kPtgFuncVarValue = 0x42;
constexpr std::uint8_t kPtgFuncVarArray = 0x62;

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

std::string Case(const std::string& name) {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_parity/claims/" + name;
}

std::vector<std::uint8_t> Part(const std::vector<std::uint8_t>& package, const char* path) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto bytes = zip.read_entry(path);
  EXPECT_TRUE(static_cast<bool>(bytes)) << path;
  return bytes ? std::move(bytes.value()) : std::vector<std::uint8_t>{};
}

/// Zeroes PtgArray's unused bytes in an rgce made of the tokens these fixtures use (PtgArray, PtgFuncVar, PtgInt,
/// PtgAttr).
void MaskArrayTokens(std::vector<std::uint8_t>& payload, std::size_t rgce_at, std::size_t cce) {
  std::size_t i = rgce_at;
  while (i < rgce_at + cce) {
    const std::uint8_t ptg = payload[i];
    if (ptg == 0x20U || ptg == 0x40U || ptg == 0x60U) {  // PtgArray in reference, value or array class
      std::fill(payload.begin() + static_cast<std::ptrdiff_t>(i + 1U),
                payload.begin() + static_cast<std::ptrdiff_t>(i + 1U + kPtgArrayUnusedBytes), 0U);
      i += 1U + kPtgArrayUnusedBytes;
    } else if (ptg == kPtgFuncVar || ptg == kPtgFuncVarValue || ptg == kPtgFuncVarArray) {
      i += 4U;
    } else if (ptg == 0x1EU) {  // PtgInt
      i += 3U;
    } else if (ptg == 0x19U && payload[i + 1U] != 0x04U) {  // PtgAttr other than AttrChoose
      i += 4U;
    } else {
      ADD_FAILURE() << "unexpected token 0x" << std::hex << static_cast<int>(ptg);
      return;
    }
  }
}

std::uint32_t U32At(const std::vector<std::uint8_t>& b, std::size_t at) {
  return static_cast<std::uint32_t>(b[at]) | (static_cast<std::uint32_t>(b[at + 1U]) << 8U) |
         (static_cast<std::uint32_t>(b[at + 2U]) << 16U) | (static_cast<std::uint32_t>(b[at + 3U]) << 24U);
}

/// The formula records of `part` (cells or names), each masked.
std::vector<std::vector<std::uint8_t>> FormulaRecords(const std::vector<std::uint8_t>& part) {
  std::vector<std::vector<std::uint8_t>> out;
  for (const FramedRecord& rec : split_records(part)) {
    std::vector<std::uint8_t> payload(rec.payload.data, rec.payload.data + rec.payload.size);
    std::size_t cce_at = 0;
    if (rec.type == kBrtFmlaNum) {
      cce_at = 8U + 8U + 2U;
    } else if (rec.type == kBrtFmlaString) {
      cce_at = 8U + 4U + 2U * U32At(payload, 8U) + 2U;
    } else if (rec.type == kBrtName) {
      // grbit (4), chKey (1), itab (4), name (XLWideString), then the formula.
      cce_at = 4U + 1U + 4U + 4U + 2U * U32At(payload, 9U);
    } else {
      continue;
    }
    MaskArrayTokens(payload, cce_at + 4U, U32At(payload, cce_at));
    // The formula and its extra data only: the cached result before it is Excel's calculation, not the encoding.
    out.emplace_back(payload.begin() + static_cast<std::ptrdiff_t>(cce_at), payload.end());
  }
  return out;
}

void ExpectFormulasMatchExcel(const std::string& stem, const char* part) {
  auto loaded = read_ooxml(SpanOf(ReadFile(Case(stem + "_src.xlsx"))));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  auto written = write_xlsb_with_result(loaded.value().workbook);
  ASSERT_TRUE(static_cast<bool>(written)) << written.error().message;
  const auto ours = FormulaRecords(Part(written.value().bytes, part));
  const auto excel = FormulaRecords(Part(ReadFile(Case(stem + ".xlsb")), part));
  ASSERT_FALSE(excel.empty());
  ASSERT_EQ(ours.size(), excel.size());
  for (std::size_t i = 0; i < ours.size(); ++i) {
    EXPECT_EQ(ours[i], excel[i]) << "formula record " << i;
  }
}

TEST(XlsbParityArrayConstant, CellFormulasMatchExcel) {
  ExpectFormulasMatchExcel("array_constants_cell", "xl/worksheets/sheet1.bin");
}

TEST(XlsbParityArrayConstant, DefinedNameMatchesExcel) {
  ExpectFormulasMatchExcel("array_constants_name", "xl/workbook.bin");
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
