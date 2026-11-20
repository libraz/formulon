//
// A 3-D reference into an external workbook whose span is written against the link's stored sheet order (Z, X),
// saved as .xlsb. The engine's save is compared with Excel's save of the same book on three facts: the link's
// sheet list, the XTI list of `BrtExternSheet`, and the XTI that B2's `PtgRef3d` selects.

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "support/roundtrip_symmetry.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

using Bytes = std::vector<std::uint8_t>;

std::string FixturePath(const char* name) {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/external_link_span/" + name;
}

/// The engine's .xlsb save of Excel's .xlsx book.
Bytes EngineSave() {
  const Bytes input = test::read_file_bytes(FixturePath("reversed_span_input.xlsx"));
  auto loaded = read_ooxml(test::span_of(input));
  EXPECT_TRUE(static_cast<bool>(loaded)) << (loaded ? "" : loaded.error().message);
  if (!loaded) {
    return {};
  }
  auto saved = write_xlsb(loaded.value().workbook);
  EXPECT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
  return saved ? std::move(saved.value()) : Bytes{};
}

Bytes ExcelSave() {
  return test::read_file_bytes(FixturePath("reversed_span_saved.xlsb"));
}

std::uint32_t U32(ByteSpan span, std::size_t at) {
  return static_cast<std::uint32_t>(span.data[at]) | (static_cast<std::uint32_t>(span.data[at + 1]) << 8) |
         (static_cast<std::uint32_t>(span.data[at + 2]) << 16) | (static_cast<std::uint32_t>(span.data[at + 3]) << 24);
}

/// The first and last sheet of each `BrtExternSheet` XTI, in table order.
std::vector<std::pair<std::int32_t, std::int32_t>> XtiList(const Bytes& package) {
  std::vector<std::pair<std::int32_t, std::int32_t>> out;
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(test::span_of(package))));
  auto workbook = zip.read_entry("xl/workbook.bin");
  if (!workbook) {
    ADD_FAILURE() << "workbook.bin missing";
    return out;
  }
  ByteSpan cursor{workbook.value().data(), workbook.value().size()};
  while (cursor.size > 0U) {
    auto record = read_record(cursor);
    if (!record) {
      ADD_FAILURE() << "unreadable record";
      break;
    }
    const XlsbRecord& r = record.value();
    if (r.type == static_cast<std::uint16_t>(XlsbRecordType::BrtExternSheet)) {
      const std::uint32_t count = U32(r.payload, 0);
      for (std::uint32_t i = 0; i < count; ++i) {
        const std::size_t at = 4U + 12U * i;
        out.emplace_back(static_cast<std::int32_t>(U32(r.payload, at + 4)),
                         static_cast<std::int32_t>(U32(r.payload, at + 8)));
      }
    }
  }
  return out;
}

/// The XTI index of the leading `PtgRef3d` of the formula in `row` (0-based) of column B of sheet 1, or -1.
std::int32_t LeadingIxti(const Bytes& package, std::uint32_t row) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(test::span_of(package))));
  auto sheet = zip.read_entry("xl/worksheets/sheet1.bin");
  if (!sheet) {
    ADD_FAILURE() << "sheet1.bin missing";
    return -1;
  }
  std::uint32_t current_row = 0;
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
      current_row = U32(p, 0);
    } else if (r.type >= static_cast<std::uint16_t>(XlsbRecordType::BrtFmlaString) &&
               r.type <= static_cast<std::uint16_t>(XlsbRecordType::BrtFmlaError) && current_row == row &&
               U32(p, 0) == 1U) {
      std::size_t at = 8;
      if (r.type == static_cast<std::uint16_t>(XlsbRecordType::BrtFmlaString)) {
        at += 4U + 2U * U32(p, at);
      } else if (r.type == static_cast<std::uint16_t>(XlsbRecordType::BrtFmlaNum)) {
        at += 8U;
      } else {
        at += 1U;
      }
      at += 2U;
      const std::size_t rgce = at + 4U;
      return static_cast<std::int32_t>(p.data[rgce + 1U] | (p.data[rgce + 2U] << 8));
    }
  }
  ADD_FAILURE() << "no formula in B" << (row + 1U);
  return -1;
}

std::vector<std::string> LinkSheets(const Bytes& xlsb) {
  auto loaded = read_xlsb(test::span_of(xlsb));
  EXPECT_TRUE(static_cast<bool>(loaded)) << (loaded ? "" : loaded.error().message);
  if (!loaded || loaded.value().workbook.external_links().empty()) {
    ADD_FAILURE() << "no external link";
    return {};
  }
  return loaded.value().workbook.external_links()[0].book.sheet_names;
}

TEST(ReversedSpanXlsb, ExcelKeepsTheStoredSheetListAndWritesNoReversedXti) {
  const Bytes excel = ExcelSave();
  EXPECT_EQ(LinkSheets(excel), (std::vector<std::string>{"Z", "X"}));
  using Pair = std::pair<std::int32_t, std::int32_t>;
  EXPECT_EQ(XtiList(excel), (std::vector<Pair>{{0, 0}, {0, 1}}));
  EXPECT_EQ(LeadingIxti(excel, 0U), 0);
  EXPECT_EQ(LeadingIxti(excel, 1U), 1);
}

TEST(ReversedSpanXlsb, TheLinkSheetListIsAsExcelSaves) {
  EXPECT_EQ(LinkSheets(EngineSave()), LinkSheets(ExcelSave()));
}

TEST(ReversedSpanXlsb, TheXtiListIsAsExcelSaves) {
  EXPECT_EQ(XtiList(EngineSave()), XtiList(ExcelSave()));
}

TEST(ReversedSpanXlsb, TheSpanFormulaUsesTheXtiExcelUses) {
  EXPECT_EQ(LeadingIxti(EngineSave(), 1U), LeadingIxti(ExcelSave(), 1U));
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
