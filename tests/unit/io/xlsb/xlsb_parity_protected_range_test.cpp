// Protected ranges (BrtRangeProtection / BrtRangeProtectionIso) the xlsb writer and reader handle, held against
// Excel 16.115's own saves of the same sources (tests/fixtures/excel/xlsb_parity/protection).

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/protection_xml.h"
#include "io/xlsb/protection_records.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "protection.h"
#include "sheet.h"
#include "utils/error.h"
#include "workbook.h"
#include "writer_test_helpers.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

using Bytes = std::vector<std::uint8_t>;

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

Bytes Fixture(const std::string& name) {
  return ReadFile(std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_parity/protection/" + name);
}

Workbook LoadXlsx(const Bytes& package) {
  auto loaded = read_ooxml(SpanOf(package));
  EXPECT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  return loaded ? std::move(loaded.value().workbook) : Workbook::create();
}

Workbook LoadXlsb(const Bytes& package) {
  auto loaded = read_xlsb(SpanOf(package));
  EXPECT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  return loaded ? std::move(loaded.value().workbook) : Workbook::create();
}

std::vector<ProtectedRange> Ranges(const Workbook& wb) {
  return read_protected_ranges(wb.sheet(0).raw_extensions());
}

/// The range-protection records (680, 536) of the first worksheet part, type and payload.
std::vector<std::pair<std::uint16_t, Bytes>> RangeRecords(const Bytes& package) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto part = zip.read_entry("xl/worksheets/sheet1.bin");
  EXPECT_TRUE(static_cast<bool>(part));
  std::vector<std::pair<std::uint16_t, Bytes>> out;
  if (part) {
    for (const FramedRecord& rec : split_records(part.value())) {
      if (rec.type == kBrtRangeProtection || rec.type == kBrtRangeProtectionIso) {
        out.emplace_back(rec.type, Bytes(rec.payload.data, rec.payload.data + rec.payload.size));
      }
    }
  }
  return out;
}

std::string Describe(const std::vector<std::pair<std::uint16_t, Bytes>>& records) {
  std::string out;
  for (const auto& rec : records) {
    out += std::to_string(rec.first) + "(" + std::to_string(rec.second.size()) + " B) ";
  }
  return out;
}

void ExpectSameRanges(const std::vector<ProtectedRange>& got, const std::vector<ProtectedRange>& want) {
  ASSERT_EQ(got.size(), want.size());
  for (std::size_t i = 0; i < want.size(); ++i) {
    EXPECT_EQ(got[i].name, want[i].name);
    EXPECT_EQ(got[i].sqref, want[i].sqref);
    EXPECT_EQ(got[i].password, want[i].password);
    EXPECT_EQ(got[i].algorithm_name, want[i].algorithm_name);
    EXPECT_EQ(got[i].hash_value, want[i].hash_value);
    EXPECT_EQ(got[i].salt_value, want[i].salt_value);
    EXPECT_EQ(got[i].spin_count, want[i].spin_count);
    EXPECT_EQ(got[i].security_descriptor, want[i].security_descriptor);
  }
}

void ExpectNotWritable(const Workbook& wb, const std::string& kind) {
  auto saved = write_xlsb_with_result(wb);
  ASSERT_FALSE(static_cast<bool>(saved));
  EXPECT_EQ(saved.error().code, FormulonErrorCode::kIoFeatureNotWritable) << saved.error().message;
  EXPECT_NE(saved.error().context.find("kind=" + kind), std::string::npos) << saved.error().context;
  EXPECT_NE(saved.error().context.find("sheet_index=0"), std::string::npos) << saved.error().context;
}

// Sources Excel opened and wrote without a security descriptor.
const char* const kWritableCases[] = {"hash_legacy", "hash_sha512", "title_255", "rects_8191"};

TEST(XlsbParityProtectedRange, RecordsMatchExcel) {
  for (const char* name : kWritableCases) {
    SCOPED_TRACE(name);
    const Workbook wb = LoadXlsx(Fixture(std::string(name) + "_src.xlsx"));
    auto saved = write_xlsb_with_result(wb);
    ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
    const auto ours = RangeRecords(saved.value().bytes);
    const auto theirs = RangeRecords(Fixture(std::string(name) + ".xlsb"));
    ASSERT_FALSE(theirs.empty());
    EXPECT_TRUE(ours == theirs) << "ours: " << Describe(ours) << " Excel: " << Describe(theirs);
  }
}

TEST(XlsbParityProtectedRange, ReadMatchesXlsxTwin) {
  // A malformed descriptor is dropped by Excel in both of its outputs.
  for (const char* name : {"hash_legacy", "hash_sha512", "title_255", "rects_8191", "sddl_malformed"}) {
    SCOPED_TRACE(name);
    const auto xlsb = Ranges(LoadXlsb(Fixture(std::string(name) + ".xlsb")));
    const auto xlsx = Ranges(LoadXlsx(Fixture(std::string(name) + ".xlsx")));
    ASSERT_EQ(xlsb.size(), 1U);
    ExpectSameRanges(xlsb, xlsx);
  }
}

TEST(XlsbParityProtectedRange, DescriptorRecordMatchesExcelResave) {
  // Excel encodes the SDDL text as a 72-byte self-relative descriptor; the record is kept as Excel wrote it.
  const Workbook wb = LoadXlsb(Fixture("sddl.xlsb"));
  EXPECT_TRUE(Ranges(wb).empty());
  auto saved = write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  const auto ours = RangeRecords(saved.value().bytes);
  const auto theirs = RangeRecords(Fixture("sddl.2.xlsb"));
  ASSERT_EQ(theirs.size(), 1U);
  EXPECT_EQ(theirs[0].second.size(), 32U + 72U);
  EXPECT_TRUE(ours == theirs) << "ours: " << Describe(ours) << " Excel: " << Describe(theirs);
}

TEST(XlsbParityProtectedRange, ChildDescriptorReadsAsExcel) {
  Workbook wb = LoadXlsx(Fixture("sddl_child_src.xlsx"));
  const auto ranges = Ranges(wb);
  ExpectSameRanges(ranges, Ranges(LoadXlsx(Fixture("sddl_child.xlsx"))));
  ASSERT_EQ(ranges.size(), 1U);
  EXPECT_EQ(ranges[0].security_descriptor, "O:WDG:WDD:(A;;CC;;;WD)");
  // Rewritten in Excel's spelling: the attribute, not the child.
  WorksheetRawExtensions& raw = wb.sheet(0).mutable_raw_extensions();
  write_protected_ranges(raw, ranges);
  std::string xml;
  for (const WorksheetRawChild& child : raw) {
    xml += child.xml;
  }
  EXPECT_NE(xml.find("securityDescriptor=\"O:WDG:WDD:(A;;CC;;;WD)\""), std::string::npos) << xml;
  EXPECT_EQ(xml.find("<securityDescriptor>"), std::string::npos) << xml;
}

TEST(XlsbParityProtectedRange, SddlDescriptorIsNotWritable) {
  for (const char* name : {"sddl_src.xlsx", "sddl_child_src.xlsx", "sddl_malformed_src.xlsx"}) {
    SCOPED_TRACE(name);
    const Workbook wb = LoadXlsx(Fixture(name));
    ASSERT_EQ(Ranges(wb).size(), 1U);
    ExpectNotWritable(wb, "protected_range_descriptor");
  }
}

TEST(XlsbParityProtectedRange, RefusedTitlesLoadAndAreNotWritable) {
  // Excel refuses each of these at open; the xlsx reader keeps them.
  for (const char* name : {"title_256_src.xlsx", "title_astral_128_src.xlsx", "name_empty_src.xlsx"}) {
    SCOPED_TRACE(name);
    const Workbook wb = LoadXlsx(Fixture(name));
    ASSERT_EQ(Ranges(wb).size(), 1U);
    ExpectNotWritable(wb, "protected_range_title");
  }
}

TEST(XlsbParityProtectedRange, RectsPastCapAreNotWritable) {
  // Excel writes a 8192-cell range as BrtRangeProtection14, which is not written.
  const Workbook wb = LoadXlsx(Fixture("rects_8192_src.xlsx"));
  ASSERT_EQ(Ranges(wb).size(), 1U);
  ExpectNotWritable(wb, "protected_range_rects");
}

Bytes RangeProtection(std::uint32_t count, const std::string& title, const Bytes& descriptor) {
  Bytes payload;
  emit_u16(payload, 0U);
  emit_u32(payload, count);
  for (std::uint32_t i = 0; i < count; ++i) {
    emit_rfx(payload, MergeRange{i * 2U, 0U, i * 2U, 0U});
  }
  emit_xlwidestring(payload, title);
  emit_u32(payload, static_cast<std::uint32_t>(descriptor.size()));
  payload.insert(payload.end(), descriptor.begin(), descriptor.end());
  return payload;
}

RangeProtectionRead Decode(const Bytes& payload) {
  ProtectedRange out;
  return decode_range_protection(ByteSpan{payload.data(), payload.size()}, ByteSpan{}, out);
}

TEST(XlsbParityProtectedRange, DecodeClassifiesRecords) {
  EXPECT_EQ(Decode(RangeProtection(1U, "R", {})), RangeProtectionRead::kDecoded);
  EXPECT_EQ(Decode(RangeProtection(kMaxProtectedRangeRects, "R", {})), RangeProtectionRead::kDecoded);
  EXPECT_EQ(Decode(RangeProtection(1U, std::string(kMaxProtectedRangeNameUnits, 'a'), {})),
            RangeProtectionRead::kDecoded);
  EXPECT_EQ(Decode(RangeProtection(1U, "R", {1U, 0U, 4U, 0x80U})), RangeProtectionRead::kSecurityDescriptor);
  EXPECT_EQ(Decode(RangeProtection(1U, "", {})), RangeProtectionRead::kUndecodable);
  EXPECT_EQ(Decode(RangeProtection(1U, std::string(kMaxProtectedRangeNameUnits + 1U, 'a'), {})),
            RangeProtectionRead::kUndecodable);
  EXPECT_EQ(Decode(RangeProtection(0U, "R", {})), RangeProtectionRead::kUndecodable);
  EXPECT_EQ(Decode(RangeProtection(static_cast<std::uint32_t>(kMaxProtectedRangeRects) + 1U, "R", {})),
            RangeProtectionRead::kUndecodable);
  // dwsdRel past the payload end.
  Bytes truncated = RangeProtection(1U, "R", {1U, 0U, 4U, 0x80U});
  truncated.resize(truncated.size() - 2U);
  EXPECT_EQ(Decode(truncated), RangeProtectionRead::kUndecodable);
}

TEST(XlsbParityProtectedRange, ExcelDescriptorRecordDecodesAsDescriptor) {
  const auto records = RangeRecords(Fixture("sddl.xlsb"));
  ASSERT_EQ(records.size(), 1U);
  EXPECT_EQ(Decode(records[0].second), RangeProtectionRead::kSecurityDescriptor);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
