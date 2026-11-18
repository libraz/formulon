// A chartsheet's protection is read as its part states it, from `<sheetProtection>` and from BrtCsProtection with
// an optional BrtCsProtectionIso.

#include <cstdint>
#include <fstream>
#include <iterator>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/xlsb/chartsheet_records.h"
#include "io/xlsb/record_writer.h"
#include "io/zip_reader.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

constexpr std::uint16_t kBrtCsProtection = 669;
constexpr std::uint16_t kBrtCsProtectionIso = 679;
constexpr const char* kSha512 = "SHA-512";
constexpr const char* kHashB64 =
    "FEOoOgNU+lu/gQaYkcJsMsJCAdolOUZsju7crdLGfmo6yGCo62jHPQ2LfHdvg+umnjLNWsGZoZgG16ju/+taFQ==";
constexpr const char* kSaltB64 = "BAMwdck9CSCcwdaFLEANaA==";
constexpr std::uint32_t kSpinCount = 100000;

ByteSpan SpanOf(const std::vector<std::uint8_t>& bytes) {
  return ByteSpan{bytes.data(), bytes.size()};
}

std::vector<std::uint8_t> ReadFixture(const std::string& name) {
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/chartsheet_twins/" + name;
  std::ifstream in(path, std::ios::binary);
  EXPECT_TRUE(in.good()) << path;
  return std::vector<std::uint8_t>((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());
}

std::vector<std::uint8_t> PartOf(const std::string& fixture, const char* path) {
  const std::vector<std::uint8_t> package = ReadFixture(fixture);
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto entry = zip.read_entry(path);
  EXPECT_TRUE(static_cast<bool>(entry)) << path;
  return entry ? std::move(entry.value()) : std::vector<std::uint8_t>{};
}

std::optional<ChartsheetProtection> FromXlsx(const std::string& stem) {
  const std::vector<std::uint8_t> part = PartOf(stem + ".xlsx", "xl/chartsheets/sheet1.xml");
  return chartsheet_protection_from_xml(std::string_view(reinterpret_cast<const char*>(part.data()), part.size()));
}

std::optional<ChartsheetProtection> FromXlsb(const std::string& stem) {
  const std::vector<std::uint8_t> part = PartOf(stem + ".xlsb", "xl/chartsheets/sheet1.bin");
  return chartsheet_protection_from_records(SpanOf(part));
}

std::string Chartsheet(const std::string& protection) {
  return "<chartsheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetPr/>" + protection +
         "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
         "</chartsheet>";
}

std::vector<std::uint8_t> CsProtectionRecord(std::uint16_t password, std::uint32_t content, std::uint32_t objects) {
  std::vector<std::uint8_t> payload;
  emit_u16(payload, password);
  emit_u32(payload, content);
  emit_u32(payload, objects);
  std::vector<std::uint8_t> part;
  emit_record(part, kBrtCsProtection, payload);
  return part;
}

std::vector<std::uint8_t> IsoRecord() {
  std::vector<std::uint8_t> payload;
  emit_u32(payload, kSpinCount);
  emit_u32(payload, 0U);
  emit_u32(payload, 0U);
  emit_u32(payload, 2U);
  payload.push_back(0xAA);
  payload.push_back(0xBB);
  emit_u32(payload, 1U);
  payload.push_back(0xCC);
  emit_xlwidestring(payload, kSha512);
  std::vector<std::uint8_t> part;
  emit_record(part, kBrtCsProtectionIso, payload);
  return part;
}

TEST(ChartsheetProtection, ExcelContentOnlyFixtureReadsTheSameFromBothFormats) {
  const std::optional<ChartsheetProtection> xlsx = FromXlsx("prot_contents_only");
  const std::optional<ChartsheetProtection> xlsb = FromXlsb("prot_contents_only");
  ASSERT_TRUE(xlsx.has_value());
  ASSERT_TRUE(xlsb.has_value());
  EXPECT_EQ(*xlsx, *xlsb);
  EXPECT_TRUE(xlsx->content_locked);
  EXPECT_FALSE(xlsx->objects_locked);
  EXPECT_EQ(xlsx->legacy_password_hash, "");
  EXPECT_EQ(xlsx->algorithm_name, "");
  EXPECT_EQ(xlsx->hash_value, "");
  EXPECT_EQ(xlsx->salt_value, "");
  EXPECT_EQ(xlsx->spin_count, 0U);
}

TEST(ChartsheetProtection, ExcelHashedFixtureReadsTheSameFromBothFormats) {
  const std::optional<ChartsheetProtection> xlsx = FromXlsx("prot_both_pw_abc");
  const std::optional<ChartsheetProtection> xlsb = FromXlsb("prot_both_pw_abc");
  ASSERT_TRUE(xlsx.has_value());
  ASSERT_TRUE(xlsb.has_value());
  EXPECT_EQ(*xlsx, *xlsb);
  EXPECT_EQ(xlsx->algorithm_name, kSha512);
  EXPECT_EQ(xlsx->hash_value, kHashB64);
  EXPECT_EQ(xlsx->salt_value, kSaltB64);
  EXPECT_EQ(xlsx->spin_count, kSpinCount);
  EXPECT_EQ(xlsx->legacy_password_hash, "");
  EXPECT_TRUE(xlsx->content_locked);
  EXPECT_TRUE(xlsx->objects_locked);
}

TEST(ChartsheetProtection, XmlObjectsOnly) {
  const std::optional<ChartsheetProtection> p =
      chartsheet_protection_from_xml(Chartsheet("<sheetProtection objects=\"1\"/>"));
  ASSERT_TRUE(p.has_value());
  EXPECT_FALSE(p->content_locked);
  EXPECT_TRUE(p->objects_locked);
  EXPECT_EQ(p->legacy_password_hash, "");
}

TEST(ChartsheetProtection, XmlLegacyPasswordOnly) {
  const std::optional<ChartsheetProtection> p =
      chartsheet_protection_from_xml(Chartsheet("<sheetProtection password=\"CBEB\" content=\"1\"/>"));
  ASSERT_TRUE(p.has_value());
  EXPECT_EQ(p->legacy_password_hash, "CBEB");
  EXPECT_TRUE(p->content_locked);
  EXPECT_FALSE(p->objects_locked);
  EXPECT_EQ(p->hash_value, "");
  EXPECT_EQ(p->spin_count, 0U);
}

TEST(ChartsheetProtection, XmlBareElementIsPresentAndUnlocked) {
  const std::optional<ChartsheetProtection> p = chartsheet_protection_from_xml(Chartsheet("<sheetProtection/>"));
  ASSERT_TRUE(p.has_value());
  EXPECT_FALSE(p->content_locked);
  EXPECT_FALSE(p->objects_locked);
}

TEST(ChartsheetProtection, XmlAbsent) {
  EXPECT_EQ(chartsheet_protection_from_xml(Chartsheet("")), std::nullopt);
  EXPECT_EQ(chartsheet_protection_from_xml("<chartsheet"), std::nullopt);
  EXPECT_EQ(chartsheet_protection_from_xml("<worksheet><sheetProtection sheet=\"1\"/></worksheet>"), std::nullopt);
}

TEST(ChartsheetProtection, RecordsLegacyPasswordOnly) {
  const std::vector<std::uint8_t> part = CsProtectionRecord(0xCBEB, 0U, 0U);
  const std::optional<ChartsheetProtection> p = chartsheet_protection_from_records(SpanOf(part));
  ASSERT_TRUE(p.has_value());
  EXPECT_EQ(p->legacy_password_hash, "CBEB");
  EXPECT_FALSE(p->content_locked);
  EXPECT_FALSE(p->objects_locked);
}

TEST(ChartsheetProtection, RecordsObjectsOnly) {
  const std::vector<std::uint8_t> part = CsProtectionRecord(0U, 0U, 1U);
  const std::optional<ChartsheetProtection> p = chartsheet_protection_from_records(SpanOf(part));
  ASSERT_TRUE(p.has_value());
  EXPECT_FALSE(p->content_locked);
  EXPECT_TRUE(p->objects_locked);
  EXPECT_EQ(p->legacy_password_hash, "");
}

TEST(ChartsheetProtection, RecordsAllZeroIsAbsentUnlessAHashAccompaniesIt) {
  const std::vector<std::uint8_t> zero = CsProtectionRecord(0U, 0U, 0U);
  EXPECT_EQ(chartsheet_protection_from_records(SpanOf(zero)), std::nullopt);

  std::vector<std::uint8_t> hashed = IsoRecord();
  const std::vector<std::uint8_t> tail = CsProtectionRecord(0U, 0U, 0U);
  hashed.insert(hashed.end(), tail.begin(), tail.end());
  const std::optional<ChartsheetProtection> p = chartsheet_protection_from_records(SpanOf(hashed));
  ASSERT_TRUE(p.has_value());
  EXPECT_EQ(p->algorithm_name, kSha512);
  EXPECT_EQ(p->hash_value, "qrs=");
  EXPECT_EQ(p->salt_value, "zA==");
  EXPECT_EQ(p->spin_count, kSpinCount);
}

TEST(ChartsheetProtection, RecordsReadALegacyPasswordBesideAHash) {
  std::vector<std::uint8_t> part = IsoRecord();
  const std::vector<std::uint8_t> tail = CsProtectionRecord(0xCBEB, 1U, 0U);
  part.insert(part.end(), tail.begin(), tail.end());
  const std::optional<ChartsheetProtection> p = chartsheet_protection_from_records(SpanOf(part));
  ASSERT_TRUE(p.has_value());
  EXPECT_EQ(p->legacy_password_hash, "CBEB");
  EXPECT_EQ(p->hash_value, "qrs=");
  EXPECT_TRUE(p->content_locked);
}

TEST(ChartsheetProtection, RecordsAbsent) {
  EXPECT_EQ(chartsheet_protection_from_records(ByteSpan{nullptr, 0}), std::nullopt);
  const std::vector<std::uint8_t> iso_only = IsoRecord();
  EXPECT_EQ(chartsheet_protection_from_records(SpanOf(iso_only)), std::nullopt);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
