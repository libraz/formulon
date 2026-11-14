// BrtPageSetup fit-to-page values and the BrtWsProp fit bit as Excel 16.115 saves them
// (tests/fixtures/excel/xlsbprint/fit_*): fitToWidth / fitToHeight keep any value, 0 included, independent of
// fitToPage, which alone sets bit 0 of BrtWsProp's second byte.

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <map>
#include <optional>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/record_writer.h"
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

using Bytes = std::vector<std::uint8_t>;

constexpr std::size_t kPageSetupSize = 38U;
constexpr std::size_t kFitWidthOffset = 24U;
constexpr std::size_t kFitHeightOffset = 28U;
constexpr std::size_t kWsPropFlagsHighByte = 1U;
constexpr std::uint8_t kFitToPageBit = 0x01U;

struct FitCase {
  const char* name;
  std::uint32_t width;
  std::uint32_t height;
  bool fit_to_page;
};

// P2 2x3 + fitToPage, P3 0x0 + fitToPage, P4 1x0 + fitToPage, P8 2x3 without fitToPage.
const FitCase kCases[] = {{"p2", 2U, 3U, true}, {"p3", 0U, 0U, true}, {"p4", 1U, 0U, true}, {"p8", 2U, 3U, false}};

Bytes ReadFile(const std::string& path) {
  Bytes out;
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

Bytes Fixture(const std::string& name, const char* ext) {
  return ReadFile(std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsbprint/fit_" + name + ext);
}

/// The payload of the first record of `type` in sheet1.bin; empty (with a failure) when absent.
Bytes FirstRecord(const Bytes& package, std::uint16_t type) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto part = zip.read_entry("xl/worksheets/sheet1.bin");
  EXPECT_TRUE(static_cast<bool>(part));
  if (part) {
    for (const FramedRecord& rec : split_records(part.value())) {
      if (rec.type == type) {
        return Bytes(rec.payload.data, rec.payload.data + rec.payload.size);
      }
    }
  }
  ADD_FAILURE() << "no record " << type;
  return Bytes{};
}

std::uint32_t U32At(const Bytes& payload, std::size_t offset) {
  if (payload.size() < offset + 4U) {
    ADD_FAILURE() << "payload too short for offset " << offset;
    return 0xFFFFFFFFU;
  }
  return static_cast<std::uint32_t>(payload[offset]) | (static_cast<std::uint32_t>(payload[offset + 1U]) << 8U) |
         (static_cast<std::uint32_t>(payload[offset + 2U]) << 16U) |
         (static_cast<std::uint32_t>(payload[offset + 3U]) << 24U);
}

bool FitBit(const Bytes& package) {
  const Bytes ws_prop = FirstRecord(package, static_cast<std::uint16_t>(XlsbRecordType::BrtWsProp));
  return ws_prop.size() > kWsPropFlagsHighByte && (ws_prop[kWsPropFlagsHighByte] & kFitToPageBit) != 0U;
}

void ExpectFit(const Bytes& package, const FitCase& c) {
  const Bytes setup = FirstRecord(package, kBrtPageSetup);
  ASSERT_EQ(setup.size(), kPageSetupSize);
  EXPECT_EQ(U32At(setup, kFitWidthOffset), c.width);
  EXPECT_EQ(U32At(setup, kFitHeightOffset), c.height);
  EXPECT_EQ(FitBit(package), c.fit_to_page);
}

TEST(XlsbPrintFitValues, ExcelFixturesCarryTheMeasuredValues) {
  for (const FitCase& c : kCases) {
    SCOPED_TRACE(c.name);
    ExpectFit(Fixture(c.name, ".xlsb"), c);
  }
}

TEST(XlsbPrintFitValues, ModelBuiltSheetWritesTheValues) {
  for (const FitCase& c : kCases) {
    SCOPED_TRACE(c.name);
    Workbook wb = Workbook::create();
    SheetPrintSettings& print = wb.sheet(0).mutable_print_settings();
    print.page_setup_xml =
        "<pageSetup fitToWidth=\"" + std::to_string(c.width) + "\" fitToHeight=\"" + std::to_string(c.height) + "\"/>";
    if (c.fit_to_page) {
      print.sheet_pr_xml = "<sheetPr><pageSetUpPr fitToPage=\"1\"/></sheetPr>";
    }
    auto saved = write_xlsb_with_result(wb);
    ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
    ExpectFit(saved.value().bytes, c);
  }
}

TEST(XlsbPrintFitValues, ExcelXlsxReadsIntoTheModelAndSavesExcelsBytes) {
  for (const FitCase& c : kCases) {
    SCOPED_TRACE(c.name);
    auto loaded = read_ooxml(SpanOf(Fixture(c.name, ".xlsx")));
    ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
    const PageSetup& setup = loaded.value().workbook.sheet(0).print_settings().page_setup;
    EXPECT_EQ(setup.fit_to_width, c.width);
    EXPECT_EQ(setup.fit_to_height, c.height);
    EXPECT_EQ(setup.fit_to_page, c.fit_to_page);

    auto saved = write_xlsb_with_result(loaded.value().workbook);
    ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
    ExpectFit(saved.value().bytes, c);
    EXPECT_EQ(FirstRecord(saved.value().bytes, kBrtPageSetup), FirstRecord(Fixture(c.name, ".xlsb"), kBrtPageSetup));
  }
}

TEST(XlsbPrintFitValues, ExcelXlsbResaveKeepsPageSetupBytesAndFitBit) {
  for (const FitCase& c : kCases) {
    SCOPED_TRACE(c.name);
    const Bytes excel = Fixture(c.name, ".xlsb");
    auto loaded = read_xlsb(SpanOf(excel));
    ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
    auto saved = write_xlsb_with_result(loaded.value().workbook);
    ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
    EXPECT_EQ(FirstRecord(saved.value().bytes, kBrtPageSetup), FirstRecord(excel, kBrtPageSetup));
    EXPECT_EQ(FitBit(saved.value().bytes), c.fit_to_page);
    EXPECT_EQ(FirstRecord(saved.value().bytes, static_cast<std::uint16_t>(XlsbRecordType::BrtWsProp)),
              FirstRecord(excel, static_cast<std::uint16_t>(XlsbRecordType::BrtWsProp)));
  }
}

/// The page-setup fields the introspection API reports.
void ExpectSamePageSetup(const SheetPrintSettings& got, const SheetPrintSettings& want) {
  EXPECT_EQ(got.page_setup_xml, want.page_setup_xml);
  EXPECT_EQ(got.page_setup.fit_to_page, want.page_setup.fit_to_page);
  EXPECT_EQ(got.page_setup.fit_to_width, want.page_setup.fit_to_width);
  EXPECT_EQ(got.page_setup.fit_to_height, want.page_setup.fit_to_height);
  EXPECT_EQ(got.page_setup.orientation, want.page_setup.orientation);
  EXPECT_EQ(got.page_setup.scale, want.page_setup.scale);
  EXPECT_EQ(got.page_setup.paper_size, want.page_setup.paper_size);
}

TEST(XlsbPrintFitValues, ExcelXlsbLoadsThePageSetupOfItsXlsxTwinThroughEitherSave) {
  for (const FitCase& c : kCases) {
    SCOPED_TRACE(c.name);
    auto twin = read_ooxml(SpanOf(Fixture(c.name, ".xlsx")));
    ASSERT_TRUE(static_cast<bool>(twin)) << twin.error().message;
    const SheetPrintSettings& want = twin.value().workbook.sheet(0).print_settings();
    auto loaded = read_xlsb(SpanOf(Fixture(c.name, ".xlsb")));
    ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
    const Workbook& wb = loaded.value().workbook;
    ExpectSamePageSetup(wb.sheet(0).print_settings(), want);

    auto as_xlsx = write_ooxml(wb);
    ASSERT_TRUE(static_cast<bool>(as_xlsx)) << as_xlsx.error().message;
    auto from_xlsx = read_ooxml(SpanOf(as_xlsx.value()));
    ASSERT_TRUE(static_cast<bool>(from_xlsx)) << from_xlsx.error().message;
    ExpectSamePageSetup(from_xlsx.value().workbook.sheet(0).print_settings(), want);

    auto as_xlsb = write_xlsb_with_result(wb);
    ASSERT_TRUE(static_cast<bool>(as_xlsb)) << as_xlsb.error().message;
    auto from_xlsb = read_xlsb(SpanOf(as_xlsb.value().bytes));
    ASSERT_TRUE(static_cast<bool>(from_xlsb)) << from_xlsb.error().message;
    ExpectSamePageSetup(from_xlsb.value().workbook.sheet(0).print_settings(), want);
  }
}

TEST(XlsbPrintFitValues, ModelBuiltFitSurvivesAnXlsbRoundTrip) {
  Workbook wb = Workbook::create();
  SheetPrintSettings& print = wb.sheet(0).mutable_print_settings();
  print.page_setup_xml = "<pageSetup fitToWidth=\"2\" fitToHeight=\"3\" orientation=\"portrait\"/>";
  print.page_setup.fit_to_width = 2U;
  print.page_setup.fit_to_height = 3U;
  print.page_setup.orientation = Orientation::kPortrait;
  print.sheet_pr_xml = "<sheetPr><pageSetUpPr fitToPage=\"1\"/></sheetPr>";
  print.page_setup.fit_to_page = true;
  auto saved = write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  auto loaded = read_xlsb(SpanOf(saved.value().bytes));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  ExpectSamePageSetup(loaded.value().workbook.sheet(0).print_settings(), wb.sheet(0).print_settings());
}

/// `package` with sheet1's BrtPageSetup payload replaced by `edit(payload)`.
template <typename Edit>
Bytes WithPageSetup(const Bytes& package, Edit edit) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto part = zip.read_entry("xl/worksheets/sheet1.bin");
  EXPECT_TRUE(static_cast<bool>(part));
  Bytes sheet;
  if (part) {
    const Bytes& src = part.value();
    for (const FramedRecord& rec : split_records(src)) {
      if (rec.type == kBrtPageSetup) {
        emit_record(sheet, kBrtPageSetup, edit(Bytes(rec.payload.data, rec.payload.data + rec.payload.size)));
      } else {
        sheet.insert(sheet.end(), src.begin() + static_cast<std::ptrdiff_t>(rec.begin),
                     src.begin() + static_cast<std::ptrdiff_t>(rec.end));
      }
    }
  }
  return RepackWithEntries(package, {{"xl/worksheets/sheet1.bin", sheet}});
}

TEST(XlsbPrintFitValues, UnstatedPageSetupStaysPassThroughAndReportsDefaults) {
  constexpr std::size_t kFlagsOffset = 32U;
  constexpr std::uint8_t kUnmappedFlagBit = 0x40U;
  const Bytes excel = Fixture("p8", ".xlsb");
  const Bytes with_flag = WithPageSetup(excel, [](Bytes payload) {
    payload[kFlagsOffset] |= kUnmappedFlagBit;
    return payload;
  });
  const Bytes with_rel_id = WithPageSetup(excel, [](Bytes payload) {
    payload.resize(kFlagsOffset + 2U);
    emit_xlnullablewidestring(payload, std::optional<std::string_view>("rId1"));
    return payload;
  });
  for (const Bytes* source : {&with_flag, &with_rel_id}) {
    auto loaded = read_xlsb(SpanOf(*source));
    ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
    const SheetPrintSettings& print = loaded.value().workbook.sheet(0).print_settings();
    EXPECT_TRUE(print.page_setup_xml.empty());
    const PageSetup defaults;
    EXPECT_EQ(print.page_setup.fit_to_width, defaults.fit_to_width);
    EXPECT_EQ(print.page_setup.fit_to_height, defaults.fit_to_height);
    EXPECT_EQ(print.page_setup.fit_to_page, defaults.fit_to_page);
    EXPECT_EQ(print.page_setup.scale, defaults.scale);
    EXPECT_EQ(print.page_setup.orientation, defaults.orientation);
    EXPECT_EQ(print.page_setup.paper_size, defaults.paper_size);
    auto saved = write_xlsb_with_result(loaded.value().workbook);
    ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
    EXPECT_EQ(FirstRecord(saved.value().bytes, kBrtPageSetup), FirstRecord(*source, kBrtPageSetup));
  }
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
