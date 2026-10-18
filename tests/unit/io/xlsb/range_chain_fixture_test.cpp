//
// Chained `:` through `.xlsb`, against `tests/fixtures/excel/xlsb_range_chains.xlsb`:
// column B of its one sheet holds one formula per row, rows 1 to 28 in the
// order of `kFormulas`, as Excel saved them.
//
// Excel tokenises a chain left to right: a whole-column or whole-row token
// and a cell area are one area operand each, every other endpoint stands
// alone, and `PtgRange` joins consecutive tokens. A chain that is not a
// single area sits behind a memory token: PtgMemArea with its bounding
// rectangle when every endpoint is a plain reference, PtgMemFunc otherwise.
// The decoder reads each formula back to the text it was typed as, and the
// encoder writes the same tokens.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <map>
#include <string>
#include <utility>
#include <vector>

#include "cell.h"
#include "gtest/gtest.h"
#include "io/xlsb/ptg_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "parser/parser.h"
#include "sheet.h"
#include "utils/arena.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

constexpr const char* kFormulas[] = {
    "=SUM(H:H:E:E)",
    "=SUM(A:C:E:G)",
    "=SUM(A:C)",
    "=SUM(A1:C3:E5)",
    "=SUM(A1:B2:C3:D4)",
    "=SUM(A:C:1:3)",
    "=SUM(A:C:E1)",
    "=SUM(1:3:E5)",
    "=SUM(A:A:C:C)",
    "=SUM(H:H)",
    "=A1:B2:C3",
    "=SUM($A:C:E:$G)",
    "=SUM(A1:B2)",
    "=SUM(OFFSET(A1,1,1):B2:C3)",
    "=SUM((A1:B2):C3)",
    "=SUM(A1:(B2:C3))",
    "=SUM(Sheet1!A1:B2:C3)",
    "=SUM(A1:B2:C3:D4:E5:F6)",
    "=SUM(A1:OFFSET(A1,1,1))",
    "=SUM(OFFSET(A1,1,1):B2)",
    "=SUM(H:H:E1)",
    "=ROWS(A1:INDEX(A1:A5,3))",
    "=A1:OFFSET(A1,1,1)",
    "=SUM(Sheet1!A1:B2)",
    "=SUM(A1:B2:INDEX(A1:A5,3))",
    "=ROWS(A:C:E:G)",
    "=A:C:E:G",
    "=COLUMNS(A1:C3:E5)",
};
constexpr std::uint32_t kFormulaCount = sizeof(kFormulas) / sizeof(kFormulas[0]);

std::vector<std::uint8_t> ReadFixture() {
  std::vector<std::uint8_t> out;
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_range_chains.xlsb";
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
    out.clear();
  }
  std::fclose(f);
  return out;
}

std::uint32_t U32(ByteSpan span, std::size_t at) {
  return static_cast<std::uint32_t>(span.data[at]) | (static_cast<std::uint32_t>(span.data[at + 1]) << 8) |
         (static_cast<std::uint32_t>(span.data[at + 2]) << 16) | (static_cast<std::uint32_t>(span.data[at + 3]) << 24);
}

/// `rgce` and the trailing `rgcb` per row of column B of sheet1.bin.
struct Streams {
  std::vector<std::uint8_t> rgce;
  std::vector<std::uint8_t> rgcb;
};

std::map<std::uint32_t, Streams> ReadStreams(const std::vector<std::uint8_t>& package) {
  std::map<std::uint32_t, Streams> out;
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(ByteSpan{package.data(), package.size()})));
  auto sheet = zip.read_entry("xl/worksheets/sheet1.bin");
  if (!sheet) {
    ADD_FAILURE() << "sheet1.bin missing";
    return out;
  }
  std::map<std::uint32_t, Streams> arrays;
  std::uint32_t row = 0;
  ByteSpan cursor{sheet.value().data(), sheet.value().size()};
  while (cursor.size > 0U) {
    auto record = read_record(cursor);
    if (!record) {
      ADD_FAILURE() << "unreadable record";
      break;
    }
    const XlsbRecord& r = record.value();
    const ByteSpan p = r.payload;
    if (r.type == static_cast<std::uint16_t>(XlsbRecordType::BrtRowHdr)) {
      row = U32(p, 0);
    } else if (r.type >= static_cast<std::uint16_t>(XlsbRecordType::BrtFmlaString) &&
               r.type <= static_cast<std::uint16_t>(XlsbRecordType::BrtFmlaError)) {
      std::size_t at = 8;
      if (r.type == static_cast<std::uint16_t>(XlsbRecordType::BrtFmlaString)) {
        at += 4U + 2U * U32(p, at);
      } else if (r.type == static_cast<std::uint16_t>(XlsbRecordType::BrtFmlaNum)) {
        at += 8U;
      } else {
        at += 1U;
      }
      at += 2U;
      const std::uint32_t cce = U32(p, at);
      const std::size_t cb_at = at + 4U + cce;
      const std::uint32_t cb = U32(p, cb_at);
      out[row] = Streams{std::vector<std::uint8_t>(p.data + at + 4, p.data + at + 4 + cce),
                         std::vector<std::uint8_t>(p.data + cb_at + 4, p.data + cb_at + 4 + cb)};
    } else if (r.type == static_cast<std::uint16_t>(XlsbRecordType::BrtArrFmla)) {
      const std::uint32_t cce = U32(p, 17);
      const std::size_t cb_at = 21U + cce;
      const std::uint32_t cb = U32(p, cb_at);
      arrays[U32(p, 0)] = Streams{std::vector<std::uint8_t>(p.data + 21, p.data + 21 + cce),
                                  std::vector<std::uint8_t>(p.data + cb_at + 4, p.data + cb_at + 4 + cb)};
    }
  }
  for (auto& [key, streams] : arrays) {
    out[key] = std::move(streams);
  }
  return out;
}

/// `rgce` with the fields Excel leaves uninitialised cleared: those of a
/// leading `PtgAttrSemi` and of a leading `PtgMemArea`.
std::vector<std::uint8_t> ClearUnused(std::vector<std::uint8_t> rgce) {
  std::size_t at = 0;
  if (rgce.size() >= 4U && rgce[0] == 0x19 && rgce[1] == 0x01) {
    rgce[2] = rgce[3] = 0;
    at = 4;
  }
  if (rgce.size() >= at + 5U && (rgce[at] & 0x1F) == 0x06) {
    std::fill(rgce.begin() + static_cast<std::ptrdiff_t>(at) + 1, rgce.begin() + static_cast<std::ptrdiff_t>(at) + 5,
              0);
  }
  return rgce;
}

Workbook Load(const std::vector<std::uint8_t>& package) {
  auto result = read_xlsb(ByteSpan{package.data(), package.size()});
  EXPECT_TRUE(static_cast<bool>(result)) << (result ? "" : result.error().message);
  return result ? std::move(result.value().workbook) : Workbook::create_empty();
}

TEST(XlsbRangeChainFixture, DecodesEveryChainToItsTypedText) {
  const Workbook wb = Load(ReadFixture());
  for (std::uint32_t i = 0; i < kFormulaCount; ++i) {
    const Cell* cell = wb.sheet(0).cell_at(i, 1U);
    ASSERT_NE(cell, nullptr) << kFormulas[i];
    EXPECT_EQ(cell->formula_text, kFormulas[i]);
  }
}

TEST(XlsbRangeChainFixture, SaveWritesExcelsTokens) {
  const std::vector<std::uint8_t> package = ReadFixture();
  const std::map<std::uint32_t, Streams> excel = ReadStreams(package);
  Workbook wb = Load(package);
  auto saved = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  const std::map<std::uint32_t, Streams> ours = ReadStreams(saved.value());
  for (std::uint32_t i = 0; i < kFormulaCount; ++i) {
    const auto e = excel.find(i);
    const auto o = ours.find(i);
    ASSERT_NE(e, excel.end()) << kFormulas[i];
    ASSERT_NE(o, ours.end()) << kFormulas[i];
    EXPECT_EQ(ClearUnused(o->second.rgce), ClearUnused(e->second.rgce)) << kFormulas[i];
    EXPECT_EQ(o->second.rgcb, e->second.rgcb) << kFormulas[i];
  }
}

TEST(XlsbRangeChainFixture, TypedChainsSurviveSaveAndLoad) {
  Workbook wb = Workbook::create();
  for (std::uint32_t i = 0; i < kFormulaCount; ++i) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, i, 1U, kFormulas[i]))) << kFormulas[i];
  }
  auto saved = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  const Workbook back = Load(saved.value());
  for (std::uint32_t i = 0; i < kFormulaCount; ++i) {
    const Cell* cell = back.sheet(0).cell_at(i, 1U);
    ASSERT_NE(cell, nullptr) << kFormulas[i];
    EXPECT_EQ(cell->formula_text, kFormulas[i]);
  }
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
