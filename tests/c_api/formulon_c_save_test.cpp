// Stable C ABI (`src/c_api/formulon_c.h`) end-to-end tests.
//
// The test driver is C++ for gtest convenience but everything it
// touches across the boundary is the pure-C surface declared in
// `formulon_c.h`.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <iterator>
#include <limits>
#include <string>
#include <string_view>
#include <thread>
#include <utility>
#include <vector>

#include "c_api/parts/common.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"
#include "io/auto_filter_xml.h"
#include "io/format_detect.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "passthrough_part.h"
#include "sheet.h"
#include "unknown_relationship.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace {

void MarkZipEntriesEncrypted(std::vector<std::uint8_t>& bytes) {
  for (std::size_t i = 0; i + 8U <= bytes.size(); ++i) {
    const bool local = bytes[i] == 'P' && bytes[i + 1U] == 'K' && bytes[i + 2U] == 3U && bytes[i + 3U] == 4U;
    const bool central = bytes[i] == 'P' && bytes[i + 1U] == 'K' && bytes[i + 2U] == 1U && bytes[i + 3U] == 2U;
    if (local || central) {
      const std::size_t flag_offset = i + (local ? 6U : 8U);
      bytes[flag_offset] = static_cast<std::uint8_t>(bytes[flag_offset] | 0x01U);
    }
  }
}

std::vector<std::uint8_t> AppendEmptyZipEntry(const std::vector<std::uint8_t>& bytes, std::string_view name) {
  const std::uint8_t signature[] = {0x50, 0x4b, 0x05, 0x06};
  const auto eocd_it = std::find_end(bytes.begin(), bytes.end(), std::begin(signature), std::end(signature));
  EXPECT_NE(eocd_it, bytes.end());
  if (eocd_it == bytes.end()) {
    return {};
  }
  const std::size_t eocd = static_cast<std::size_t>(eocd_it - bytes.begin());
  const auto read16 = [&](std::size_t offset) {
    return static_cast<std::uint16_t>(bytes[offset] | (static_cast<std::uint16_t>(bytes[offset + 1U]) << 8U));
  };
  const auto read32 = [&](std::size_t offset) {
    return static_cast<std::uint32_t>(bytes[offset] | (static_cast<std::uint32_t>(bytes[offset + 1U]) << 8U) |
                                      (static_cast<std::uint32_t>(bytes[offset + 2U]) << 16U) |
                                      (static_cast<std::uint32_t>(bytes[offset + 3U]) << 24U));
  };
  const std::uint16_t count = read16(eocd + 10U);
  const std::uint32_t central_size = read32(eocd + 12U);
  const std::uint32_t central_offset = read32(eocd + 16U);
  const auto put16 = [](std::vector<std::uint8_t>& out, std::uint16_t value) {
    out.push_back(static_cast<std::uint8_t>(value));
    out.push_back(static_cast<std::uint8_t>(value >> 8U));
  };
  const auto put32 = [](std::vector<std::uint8_t>& out, std::uint32_t value) {
    out.push_back(static_cast<std::uint8_t>(value));
    out.push_back(static_cast<std::uint8_t>(value >> 8U));
    out.push_back(static_cast<std::uint8_t>(value >> 16U));
    out.push_back(static_cast<std::uint8_t>(value >> 24U));
  };
  const std::string encoded_name(name);
  std::vector<std::uint8_t> local;
  put32(local, 0x04034b50U);
  put16(local, 20);
  put16(local, 0);
  put16(local, 0);
  put16(local, 0);
  put16(local, 0);
  put32(local, 0);
  put32(local, 0);
  put32(local, 0);
  put16(local, static_cast<std::uint16_t>(encoded_name.size()));
  put16(local, 0);
  local.insert(local.end(), encoded_name.begin(), encoded_name.end());
  std::vector<std::uint8_t> central;
  put32(central, 0x02014b50U);
  put16(central, 20);
  put16(central, 20);
  put16(central, 0);
  put16(central, 0);
  put16(central, 0);
  put16(central, 0);
  put32(central, 0);
  put32(central, 0);
  put32(central, 0);
  put16(central, static_cast<std::uint16_t>(encoded_name.size()));
  put16(central, 0);
  put16(central, 0);
  put16(central, 0);
  put16(central, 0);
  put32(central, 0);
  put32(central, central_offset);
  central.insert(central.end(), encoded_name.begin(), encoded_name.end());
  std::vector<std::uint8_t> out;
  out.insert(out.end(), bytes.begin(), bytes.begin() + central_offset);
  out.insert(out.end(), local.begin(), local.end());
  out.insert(out.end(), bytes.begin() + central_offset, bytes.begin() + central_offset + central_size);
  out.insert(out.end(), central.begin(), central.end());
  put32(out, 0x06054b50U);
  put16(out, 0);
  put16(out, 0);
  put16(out, count + 1U);
  put16(out, count + 1U);
  put32(out, central_size + static_cast<std::uint32_t>(central.size()));
  put32(out, central_offset + static_cast<std::uint32_t>(local.size()));
  put16(out, 0);
  return out;
}

}  // namespace

TEST(FormulonCApi, LoadMapsCorruptAndEncryptedContainersToIoErrors) {
  fm_workbook_t* loaded = reinterpret_cast<fm_workbook_t*>(0x1);
  const std::vector<std::uint8_t> garbage = {0x01U, 0x02U, 0x03U, 0x04U};
  EXPECT_EQ(fm_workbook_load(garbage.data(), garbage.size(), &loaded),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kIoZipCorrupt));
  EXPECT_EQ(loaded, nullptr);

  loaded = reinterpret_cast<fm_workbook_t*>(0x1);
  const std::vector<std::uint8_t> cdfv2 = {0xD0U, 0xCFU, 0x11U, 0xE0U, 0xA1U, 0xB1U, 0x1AU, 0xE1U};
  EXPECT_EQ(fm_workbook_load(cdfv2.data(), cdfv2.size(), &loaded),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kIoZipEncrypted));
  EXPECT_EQ(loaded, nullptr);

  formulon::Workbook source = formulon::Workbook::create();
  auto saved = source.save();
  ASSERT_TRUE(static_cast<bool>(saved));
  std::vector<std::uint8_t> encrypted = saved.value();
  MarkZipEntriesEncrypted(encrypted);
  loaded = reinterpret_cast<fm_workbook_t*>(0x1);
  EXPECT_EQ(fm_workbook_load(encrypted.data(), encrypted.size(), &loaded),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kIoZipEncrypted));
  EXPECT_EQ(loaded, nullptr);
}
TEST(FormulonCApi, SaveLoadRoundTrip) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Second"), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 7.0), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 1, 0, "=A1+1"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  BufferGuard buf;
  ASSERT_EQ(fm_workbook_save(wb.handle, &buf.data, &buf.len), 0);
  ASSERT_NE(buf.data, nullptr);
  EXPECT_GT(buf.len, 0U);

  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(buf.data, buf.len, &loaded.handle), 0);
  EXPECT_EQ(fm_workbook_sheet_count(loaded.handle), 2U);

  const char* sheet0 = nullptr;
  ASSERT_EQ(fm_workbook_sheet_name(loaded.handle, 0, &sheet0), 0);
  EXPECT_STREQ(sheet0, "Sheet1");

  // The literal A1=7 must round-trip; the formula B1=A1+1 may need a
  // recalc on the loaded workbook to populate its cached value.
  ASSERT_EQ(fm_workbook_recalc(loaded.handle), 0);
  fm_value_t a1{};
  ASSERT_EQ(fm_workbook_get_value(loaded.handle, 0, 0, 0, &a1), 0);
  EXPECT_EQ(a1.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(a1.u.number, 7.0);
}
TEST(FormulonCApi, SaveExXlsxMatchesSave) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 7.0), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  BufferGuard xlsx_buf;
  ASSERT_EQ(fm_workbook_save_as(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &xlsx_buf.data, &xlsx_buf.len), 0);
  ASSERT_NE(xlsx_buf.data, nullptr);
  EXPECT_GT(xlsx_buf.len, 0U);

  // `FM_WORKBOOK_FORMAT_XLSX` must produce an OOXML container, so the
  // C ABI's own format sniff (used by `fm_workbook_load`) reports it as
  // such rather than xlsb.
  formulon::io::ByteSpan xlsx_span{xlsx_buf.data, xlsx_buf.len};
  EXPECT_EQ(formulon::io::detect_workbook_format(xlsx_span), formulon::WorkbookFormat::Ooxml);
}
TEST(FormulonCApi, SaveExXlsbProducesLoadableXlsbContainer) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 42.0), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  BufferGuard xlsb_buf;
  ASSERT_EQ(fm_workbook_save_as(wb.handle, FM_WORKBOOK_FORMAT_XLSB, &xlsb_buf.data, &xlsb_buf.len), 0);
  ASSERT_NE(xlsb_buf.data, nullptr);
  EXPECT_GT(xlsb_buf.len, 0U);

  // The bytes must be a real MS-XLSB package (declares `xl/workbook.bin`,
  // not `xl/workbook.xml`), and must load back through the byte-only
  // C ABI, which auto-detects the container from its contents.
  formulon::io::ByteSpan xlsb_span{xlsb_buf.data, xlsb_buf.len};
  EXPECT_EQ(formulon::io::detect_workbook_format(xlsb_span), formulon::WorkbookFormat::Xlsb);

  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(xlsb_buf.data, xlsb_buf.len, &loaded.handle), 0);
  ASSERT_EQ(fm_workbook_recalc(loaded.handle), 0);
  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(loaded.handle, 0, 0, 0, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 42.0);
}
TEST(FormulonCApi, SaveWithDiagnosticsReportsTheXlsbCountersAndLeavesTheXlsxOnesZero) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=T[C]"), 0);

  auto& sheet = wb.handle->workbook().sheet(0);
  formulon::Hyperlink hyperlink;
  hyperlink.target = "https://example.com";
  sheet.mutable_hyperlinks().push_back(std::move(hyperlink));
  sheet.mutable_validations().push_back(formulon::DataValidation{});
  sheet.set_auto_filter(formulon::io::auto_filter_from_xml("<autoFilter ref=\"A1:B2\"/>"));

  // The OOXML writer represents all of the above, so a clean XLSX save
  // reports nothing lost -- including on the two fields both writers own.
  BufferGuard xlsx_buf;
  fm_save_diagnostics_t xlsx{99U, 99U, 99U, 99U, 99U};
  ASSERT_EQ(fm_workbook_save_with_diagnostics(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &xlsx_buf.data, &xlsx_buf.len, &xlsx),
            0);
  EXPECT_EQ(xlsx.downgraded_formula_count, 0U);
  EXPECT_EQ(xlsx.deferred_feature_count, 0U);
  EXPECT_EQ(xlsx.dropped_part_count, 0U);
  EXPECT_EQ(xlsx.dropped_relationship_count, 0U);
  EXPECT_EQ(xlsx.renumbered_part_count, 0U);

  BufferGuard xlsb_buf;
  fm_save_diagnostics_t xlsb{};
  ASSERT_EQ(fm_workbook_save_with_diagnostics(wb.handle, FM_WORKBOOK_FORMAT_XLSB, &xlsb_buf.data, &xlsb_buf.len, &xlsb),
            0);
  EXPECT_EQ(xlsb.downgraded_formula_count, 1U);
  // Auto-filter state remains deferred. Hyperlinks and validations are
  // written from the model and are therefore not counted here.
  EXPECT_EQ(xlsb.deferred_feature_count, 1U);
  // `renumbered_part_count` has no XLSB source: the binary writer never
  // reassigns a part id.
  EXPECT_EQ(xlsb.renumbered_part_count, 0U);
}
TEST(FormulonCApi, SaveWithDiagnosticsCarriesTheOoxmlWriterCountersAcrossTheAbi) {
  // The OOXML writer's own counters are unit-tested at the io layer; this
  // pins that they survive the projection onto `fm_save_diagnostics_t`
  // instead of arriving as zero.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 1.0), 0);

  // A preserved copy of a part the writer always generates loses the
  // collision and is dropped; a relationship whose target part is absent is
  // dropped rather than left dangling.
  formulon::PassthroughPart stale;
  stale.path = "xl/styles.xml";
  stale.content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
  stale.bytes = {'<', '/', '>'};
  wb.handle->workbook().set_passthrough_parts({std::move(stale)});
  formulon::UnknownRelationship orphan;
  orphan.id = "rId9";
  orphan.type = "http://schemas.example.com/orphan";
  orphan.target = "xl/missing.xml";
  wb.handle->workbook().set_unknown_workbook_rels({std::move(orphan)});

  BufferGuard buf;
  fm_save_diagnostics_t d{99U, 99U, 99U, 99U, 99U};
  ASSERT_EQ(fm_workbook_save_with_diagnostics(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &buf.data, &buf.len, &d), 0);
  EXPECT_GT(buf.len, 0U);
  EXPECT_EQ(d.dropped_part_count, 1U);
  EXPECT_EQ(d.dropped_relationship_count, 1U);
  EXPECT_EQ(d.downgraded_formula_count, 0U);
  EXPECT_EQ(d.deferred_feature_count, 0U);
  EXPECT_EQ(d.renumbered_part_count, 0U);
}
TEST(FormulonCApi, ReadDiagnosticsAreZeroForCleanRoundTrip) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 42.0), 0);

  for (const fm_workbook_format_t format : {FM_WORKBOOK_FORMAT_XLSX, FM_WORKBOOK_FORMAT_XLSB}) {
    BufferGuard bytes;
    ASSERT_EQ(fm_workbook_save_as(wb.handle, format, &bytes.data, &bytes.len), 0);
    WorkbookGuard loaded;
    ASSERT_EQ(fm_workbook_load(bytes.data, bytes.len, &loaded.handle), 0);

    fm_read_diagnostics_t d{99U, 99U, 99U, 99U, 99U};
    ASSERT_EQ(fm_workbook_read_diagnostics(loaded.handle, &d), 0);
    EXPECT_EQ(d.undecoded_formula_count, 0U) << "format=" << format;
    EXPECT_EQ(d.undecoded_defined_name_count, 0U) << "format=" << format;
    EXPECT_EQ(d.undecoded_part_count, 0U) << "format=" << format;
    EXPECT_EQ(d.skipped_feature_count, 0U) << "format=" << format;
    EXPECT_EQ(d.unknown_content_type_count, 0U) << "format=" << format;
  }
}
TEST(FormulonCApi, DiagnosticsEntryPointsInitializeOutputsOnFailure) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  // A workbook that was never loaded reports every counter as zero rather
  // than leaving the caller's struct untouched.
  fm_read_diagnostics_t read{99U, 99U, 99U, 99U, 99U};
  ASSERT_EQ(fm_workbook_read_diagnostics(wb.handle, &read), 0);
  EXPECT_EQ(read.undecoded_formula_count, 0U);
  EXPECT_EQ(read.skipped_feature_count, 0U);

  uint8_t* bytes = reinterpret_cast<uint8_t*>(static_cast<uintptr_t>(1));
  size_t len = 99U;
  fm_save_diagnostics_t save{99U, 99U, 99U, 99U, 99U};
  EXPECT_NE(fm_workbook_save_with_diagnostics(wb.handle, FM_WORKBOOK_FORMAT_UNKNOWN, &bytes, &len, &save), 0);
  EXPECT_EQ(bytes, nullptr);
  EXPECT_EQ(len, 0U);
  EXPECT_EQ(save.downgraded_formula_count, 0U);
  EXPECT_EQ(save.deferred_feature_count, 0U);
  EXPECT_EQ(save.dropped_part_count, 0U);
  EXPECT_EQ(save.dropped_relationship_count, 0U);
  EXPECT_EQ(save.renumbered_part_count, 0U);

  read = fm_read_diagnostics_t{99U, 99U, 99U, 99U, 99U};
  EXPECT_NE(fm_workbook_read_diagnostics(nullptr, &read), 0);
  EXPECT_EQ(read.undecoded_formula_count, 0U);
  EXPECT_EQ(read.undecoded_defined_name_count, 0U);
  EXPECT_EQ(read.undecoded_part_count, 0U);
  EXPECT_EQ(read.skipped_feature_count, 0U);
  EXPECT_EQ(read.unknown_content_type_count, 0U);
}
TEST(FormulonCApi, ReadDiagnosticsPreservesUnknownPartFromDeterministicFixture) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  BufferGuard bytes;
  ASSERT_EQ(fm_workbook_save_as(wb.handle, FM_WORKBOOK_FORMAT_XLSB, &bytes.data, &bytes.len), 0);
  const std::vector<std::uint8_t> fixture =
      AppendEmptyZipEntry(std::vector<std::uint8_t>(bytes.data, bytes.data + bytes.len), "xl/dropped.bin");
  ASSERT_FALSE(fixture.empty());
  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(fixture.data(), fixture.size(), &loaded.handle), 0);
  fm_read_diagnostics_t d{};
  ASSERT_EQ(fm_workbook_read_diagnostics(loaded.handle, &d), 0);
  EXPECT_EQ(d.undecoded_formula_count, 0U);
  EXPECT_EQ(d.undecoded_defined_name_count, 0U);
  // Unknown package parts are retained as passthrough data, so loading the
  // fixture is lossless and the undecoded-part diagnostic remains zero.
  EXPECT_EQ(d.undecoded_part_count, 0U);
}
TEST(FormulonCApi, DiagnosticFailuresNameTheInvokedSymbol) {
  fm_read_diagnostics_t read{77U, 77U, 77U, 77U, 77U};
  EXPECT_NE(fm_workbook_read_diagnostics(nullptr, &read), 0);
  EXPECT_EQ(read.undecoded_formula_count, 0U);
  EXPECT_EQ(read.skipped_feature_count, 0U);
  EXPECT_STREQ(fm_last_error_message(), "fm_workbook_read_diagnostics: NULL argument");

  uint8_t* bytes = reinterpret_cast<uint8_t*>(static_cast<uintptr_t>(1));
  size_t len = 77U;
  EXPECT_NE(fm_workbook_save_as(nullptr, FM_WORKBOOK_FORMAT_XLSB, &bytes, &len), 0);
  EXPECT_EQ(bytes, nullptr);
  EXPECT_EQ(len, 0U);
  EXPECT_STREQ(fm_last_error_message(), "fm_workbook_save_as: NULL argument");

  fm_save_diagnostics_t save{77U, 77U, 77U, 77U, 77U};
  bytes = reinterpret_cast<uint8_t*>(static_cast<uintptr_t>(1));
  len = 77U;
  EXPECT_NE(fm_workbook_save_with_diagnostics(nullptr, FM_WORKBOOK_FORMAT_XLSB, &bytes, &len, &save), 0);
  EXPECT_EQ(bytes, nullptr);
  EXPECT_EQ(len, 0U);
  EXPECT_EQ(save.downgraded_formula_count, 0U);
  EXPECT_EQ(save.renumbered_part_count, 0U);
  EXPECT_STREQ(fm_last_error_message(), "fm_workbook_save_with_diagnostics: NULL argument");
}
TEST(FormulonCApi, EverySaveEntryPointZeroesOutParamsOnFailure) {
  // The whole save family shares one failure-path contract, so a caller may
  // reuse the same out variables across calls and free unconditionally. A
  // member that left a previous call's pointer in place would hand that
  // stale pointer to `fm_buffer_free` a second time.
  uint8_t* const poison = reinterpret_cast<uint8_t*>(static_cast<uintptr_t>(1));

  uint8_t* bytes = poison;
  size_t len = 77U;
  EXPECT_NE(fm_workbook_save(nullptr, &bytes, &len), 0);
  EXPECT_EQ(bytes, nullptr);
  EXPECT_EQ(len, 0U);
  EXPECT_STREQ(fm_last_error_message(), "fm_workbook_save: NULL argument");

  bytes = poison;
  len = 77U;
  EXPECT_NE(fm_workbook_save_as(nullptr, FM_WORKBOOK_FORMAT_XLSX, &bytes, &len), 0);
  EXPECT_EQ(bytes, nullptr);
  EXPECT_EQ(len, 0U);

  bytes = poison;
  len = 77U;
  fm_save_diagnostics_t save{77U, 77U, 77U, 77U, 77U};
  EXPECT_NE(fm_workbook_save_with_diagnostics(nullptr, FM_WORKBOOK_FORMAT_XLSX, &bytes, &len, &save), 0);
  EXPECT_EQ(bytes, nullptr);
  EXPECT_EQ(len, 0U);
  EXPECT_EQ(save.downgraded_formula_count, 0U);
  EXPECT_EQ(save.deferred_feature_count, 0U);
  EXPECT_EQ(save.dropped_part_count, 0U);
  EXPECT_EQ(save.dropped_relationship_count, 0U);
  EXPECT_EQ(save.renumbered_part_count, 0U);

  // A NULL out-parameter is rejected without being dereferenced, on the base
  // entry point as well as the extended ones.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  len = 77U;
  EXPECT_NE(fm_workbook_save(wb.handle, nullptr, &len), 0);
  EXPECT_EQ(len, 0U);
  bytes = poison;
  EXPECT_NE(fm_workbook_save(wb.handle, &bytes, nullptr), 0);
  EXPECT_EQ(bytes, nullptr);
}
TEST(FormulonCApi, BaseSaveProducesTheSameBytesAsTheXlsxFormatSelector) {
  // The header declares the two equivalent; delegating the base entry point
  // to the shared implementation is what keeps that true.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=1+2"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  BufferGuard base;
  ASSERT_EQ(fm_workbook_save(wb.handle, &base.data, &base.len), 0);
  BufferGuard selected;
  ASSERT_EQ(fm_workbook_save_as(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &selected.data, &selected.len), 0);

  ASSERT_EQ(base.len, selected.len);
  ASSERT_GT(base.len, 0U);
  EXPECT_EQ(std::memcmp(base.data, selected.data, base.len), 0);
}
TEST(FormulonCApi, ReservedStatusCodesKeepTheirSlotsAndSpelling) {
  // Some documented codes are allocated but no shipping path builds them.
  // The header promises numeric identity with `formulon::FormulonErrorCode`,
  // so those slots must not be reused or renumbered, and `fm_status_string`
  // must keep naming them exactly: only a value outside the enum may fall
  // back to `"kUnknownError"`.
  struct Reserved {
    fm_status_t code;
    const char* name;
  };
  const Reserved reserved[] = {
      {5008, "kIoXmlDoctype"},
      {5009, "kIoXmlEntityExplosion"},
      {5015, "kIoCsvEncodingDetect"},
      {6000, "kCryptoAgileNotSupported"},
      {6001, "kCryptoStandardNotSupported"},
      {6002, "kCryptoBadPassword"},
      {6003, "kCryptoHashMismatch"},
      {6004, "kCryptoKeyDerivationFailed"},
  };
  for (const Reserved& entry : reserved) {
    EXPECT_STREQ(fm_status_string(entry.code), entry.name) << "code=" << entry.code;
  }

  // The reserved crypto band is not what an encrypted package reports; a
  // binding offering a password prompt must branch on `kIoZipEncrypted`.
  EXPECT_STREQ(fm_status_string(static_cast<fm_status_t>(formulon::FormulonErrorCode::kIoZipEncrypted)),
               "kIoZipEncrypted");
  EXPECT_STREQ(fm_status_string(123456), "kUnknownError");
}
TEST(FormulonCApi, MemoryUsageTracksTheWorkbookAndRejectsNulls) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  size_t empty = 0U;
  ASSERT_EQ(fm_workbook_memory_usage(wb.handle, &empty), 0);
  EXPECT_GT(empty, 0U);

  for (uint32_t row = 0; row < 200U; ++row) {
    ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, row, 0, "=1+2"), 0);
  }
  size_t filled = 0U;
  ASSERT_EQ(fm_workbook_memory_usage(wb.handle, &filled), 0);
  EXPECT_GT(filled, empty);

  // Both NULL arguments are rejected rather than silently reporting 0,
  // which is what lets a binding distinguish "empty" from "broken".
  size_t scratch = 0U;
  EXPECT_NE(fm_workbook_memory_usage(nullptr, &scratch), 0);
  EXPECT_NE(fm_workbook_memory_usage(wb.handle, nullptr), 0);
}
TEST(FormulonCApi, SaveExRejectsUnknownFormat) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  const std::int32_t invalid_formats[] = {99, std::numeric_limits<std::int32_t>::min(),
                                          std::numeric_limits<std::int32_t>::max()};
  for (const std::int32_t raw : invalid_formats) {
    BufferGuard buf;
    EXPECT_EQ(fm_workbook_save_as(wb.handle, raw, &buf.data, &buf.len),
              static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
    EXPECT_EQ(buf.data, nullptr);
    EXPECT_EQ(buf.len, 0U);
    fm_save_diagnostics_t save{99U, 99U, 99U, 99U, 99U};
    EXPECT_EQ(fm_workbook_save_with_diagnostics(wb.handle, raw, &buf.data, &buf.len, &save),
              static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
    EXPECT_EQ(buf.data, nullptr);
    EXPECT_EQ(buf.len, 0U);
    EXPECT_EQ(save.downgraded_formula_count, 0U);
    EXPECT_EQ(save.deferred_feature_count, 0U);
    EXPECT_EQ(save.dropped_part_count, 0U);
    EXPECT_EQ(save.dropped_relationship_count, 0U);
    EXPECT_EQ(save.renumbered_part_count, 0U);
  }
}
TEST(FormulonCApi, LoadRoutesXlsbBytesToXlsbReader) {
  // Build a minimal `.xlsb` byte stream via the engine writer, then load
  // it through the byte-only C ABI. The load boundary must detect the
  // xlsb container and route to `read_xlsb` rather than failing in the
  // OOXML reader with a "missing xl/workbook.xml" diagnostic.
  formulon::Workbook src = formulon::Workbook::create_empty();
  formulon::Sheet& s = src.sheet(src.add_sheet("S"));
  s.set_cell_value(0U, 0U, formulon::Value::number(123.5));
  auto xlsb_or = formulon::io::xlsb::write_xlsb(src);
  ASSERT_TRUE(static_cast<bool>(xlsb_or)) << xlsb_or.error().message;
  const std::vector<std::uint8_t>& xlsb = xlsb_or.value();

  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(xlsb.data(), xlsb.size(), &loaded.handle), 0);
  EXPECT_EQ(fm_workbook_sheet_count(loaded.handle), 1U);

  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(loaded.handle, 0, 0, 0, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 123.5);
}
TEST(FormulonCApi, SaveLoadFormulaTextResult) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 1, 0, "=UPPER(\"world\")"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  BufferGuard buf;
  ASSERT_EQ(fm_workbook_save(wb.handle, &buf.data, &buf.len), 0);
  ASSERT_NE(buf.data, nullptr);
  EXPECT_GT(buf.len, 0U);

  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(buf.data, buf.len, &loaded.handle), 0);
  ASSERT_EQ(fm_workbook_recalc(loaded.handle), 0);
  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(loaded.handle, 0, 1, 0, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_TEXT);
  ASSERT_NE(v.u.text, nullptr);
  EXPECT_STREQ(v.u.text, "WORLD");
}
TEST(FormulonCApi, SaveLoadFormulaBoolResult) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 2, 0, "=TRUE()"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  BufferGuard buf;
  ASSERT_EQ(fm_workbook_save(wb.handle, &buf.data, &buf.len), 0);
  ASSERT_NE(buf.data, nullptr);
  EXPECT_GT(buf.len, 0U);

  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(buf.data, buf.len, &loaded.handle), 0);
  ASSERT_EQ(fm_workbook_recalc(loaded.handle), 0);
  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(loaded.handle, 0, 2, 0, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_BOOL);
  EXPECT_EQ(v.u.boolean, 1);
}
TEST(FormulonCApi, SaveLoadFormulaErrorResult) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 3, 0, "=1/0"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  BufferGuard buf;
  ASSERT_EQ(fm_workbook_save(wb.handle, &buf.data, &buf.len), 0);
  ASSERT_NE(buf.data, nullptr);
  EXPECT_GT(buf.len, 0U);

  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(buf.data, buf.len, &loaded.handle), 0);
  ASSERT_EQ(fm_workbook_recalc(loaded.handle), 0);
  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(loaded.handle, 0, 3, 0, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_ERROR);
  EXPECT_EQ(v.u.error_code, 1);
}

TEST(FormulonCApi, ThreadLocalLastErrorIsolation) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  // Establish a known last-error on the main thread.
  fm_value_t v{};
  ASSERT_NE(fm_workbook_get_value(nullptr, 0, 0, 0, &v), 0);
  const std::string main_msg = fm_last_error_message();
  EXPECT_GT(main_msg.size(), 0U);

  // A worker thread triggers a *different* error path; the main
  // thread's last error must remain untouched.
  std::atomic<fm_status_t> worker_rc{0};
  std::thread worker([&]() {
    // Out-of-range sheet on a freshly created workbook on this thread.
    fm_workbook_t* local = nullptr;
    if (fm_workbook_create(&local) != 0) {
      return;
    }
    worker_rc.store(fm_workbook_set_number(local, 99, 0, 0, 1.0));
    fm_workbook_destroy(local);
  });
  worker.join();
  EXPECT_NE(worker_rc.load(), 0);

  // Main thread's diagnostic survived the worker's run.
  EXPECT_EQ(std::string(fm_last_error_message()), main_msg);
}
TEST(FormulonCApi, StatusStringCoversKnownCodes) {
  // Spot-check a handful of band-spanning codes; the source-of-truth is
  // `formulon::to_cstring`, which the C API forwards to.
  EXPECT_STREQ(fm_status_string(0), "kOk");
  EXPECT_STREQ(fm_status_string(2), "kInvalidArgument");
  EXPECT_STREQ(fm_status_string(7000), "kBindingInvalidHandle");
  EXPECT_STREQ(fm_status_string(7001), "kBindingNullPointer");
  // Unknown numeric values must still return a non-NULL fallback.
  const char* unknown = fm_status_string(123456);
  ASSERT_NE(unknown, nullptr);
  EXPECT_GT(std::strlen(unknown), 0U);
}
TEST(FormulonCApi, ErrorDisplayNameUsesExcelLiterals) {
  EXPECT_STREQ(fm_error_display_name(0), "#NULL!");
  EXPECT_STREQ(fm_error_display_name(1), "#DIV/0!");
  EXPECT_STREQ(fm_error_display_name(3), "#REF!");
  EXPECT_STREQ(fm_error_display_name(-1), "#UNKNOWN!");
  EXPECT_STREQ(fm_error_display_name(999), "#UNKNOWN!");
}
TEST(FormulonCApi, VersionStringNonEmpty) {
  const char* v = fm_version_string();
  ASSERT_NE(v, nullptr);
  EXPECT_GT(std::strlen(v), 0U);
}
