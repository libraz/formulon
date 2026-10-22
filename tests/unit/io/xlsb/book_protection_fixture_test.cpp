//
// Workbook protection against Windows Excel 365 (16.0.20228, ja-JP) saves
// under `tests/fixtures/excel/win/book_protection/`: each case saved as both
// .xlsx and .xlsb, so the .xlsx `<workbookProtection>` is the reference for
// the .xlsb BrtBookProtection / BrtBookProtectionIso records.
//
// com_s1 / com_s2 / com_s3 were protected through COM (structure + windows,
// with password "a" for s2, windows only for s3); the hand_* cases are
// hand-written `<workbookProtection>` elements (`*.input.xlsx`) that Excel
// opened and re-saved. Excel keeps lockStructure and the SHA-512 password and
// drops lockWindows, lockRevision and revisionsPassword on both paths, in
// both formats, although BrtBookProtection defines a field for each.

#include <cstdint>
#include <cstdio>
#include <map>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/xlsb/protection_records.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

std::vector<std::uint8_t> ReadFixture(const std::string& file) {
  std::vector<std::uint8_t> out;
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/win/book_protection/" + file;
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

Workbook LoadXlsx(const std::string& file) {
  const std::vector<std::uint8_t> bytes = ReadFixture(file);
  auto r = read_ooxml(ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(r)) << file << ": " << (r ? "" : r.error().message);
  return r ? std::move(r.value().workbook) : Workbook::create_empty();
}

Workbook LoadXlsb(const std::string& file) {
  const std::vector<std::uint8_t> bytes = ReadFixture(file);
  auto r = read_xlsb(ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(r)) << file << ": " << (r ? "" : r.error().message);
  return r ? std::move(r.value().workbook) : Workbook::create_empty();
}

/// Deferred features a save reports beyond the same book saved unprotected,
/// so unrelated deferrals (the fixtures' page margins) do not count.
std::uint32_t ProtectionDeferrals(Workbook wb) {
  auto saved = write_xlsb_with_result(wb);
  EXPECT_TRUE(static_cast<bool>(saved));
  wb.set_workbook_protection_xml("");
  auto baseline = write_xlsb_with_result(wb);
  EXPECT_TRUE(static_cast<bool>(baseline));
  return saved && baseline
             ? saved.value().diagnostics.deferred_feature_count - baseline.value().diagnostics.deferred_feature_count
             : 0U;
}

/// The BrtBookProtection (534) and BrtBookProtectionIso (677) payloads of a
/// package's workbook.bin, keyed by record type.
std::map<std::uint16_t, std::vector<std::uint8_t>> BookProtectionRecords(const std::vector<std::uint8_t>& package) {
  std::map<std::uint16_t, std::vector<std::uint8_t>> out;
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(ByteSpan{package.data(), package.size()})));
  auto book = zip.read_entry("xl/workbook.bin");
  if (!book) {
    ADD_FAILURE() << "no workbook.bin";
    return out;
  }
  ByteSpan cursor{book.value().data(), book.value().size()};
  while (cursor.size > 0U) {
    auto record = read_record(cursor);
    if (!record) {
      ADD_FAILURE() << "unreadable record";
      break;
    }
    if (record.value().type == 534U || record.value().type == 677U) {
      const ByteSpan p = record.value().payload;
      out[record.value().type] = std::vector<std::uint8_t>(p.data, p.data + p.size);
    }
  }
  return out;
}

const char* const kCases[] = {"com_s1_struct_win_nopw", "com_s2_struct_win_pw_a", "com_s3_nostruct_win_nopw",
                              "hand_lockAll",           "hand_lockRevision",      "hand_lockWindows",
                              "hand_revisionsPassword"};

TEST(XlsbProtectionCodec, Base64RoundTripsRfc4648Vectors) {
  struct Vector {
    const char* plain;
    const char* encoded;
  };
  constexpr Vector vectors[] = {{"", ""},
                                {"f", "Zg=="},
                                {"fo", "Zm8="},
                                {"foo", "Zm9v"},
                                {"foob", "Zm9vYg=="},
                                {"fooba", "Zm9vYmE="},
                                {"foobar", "Zm9vYmFy"}};
  for (const Vector& vector : vectors) {
    const std::string plain(vector.plain);
    std::vector<std::uint8_t> bytes;
    bytes.reserve(plain.size());
    for (char c : plain) {
      bytes.push_back(static_cast<std::uint8_t>(c));
    }
    EXPECT_EQ(base64_encode(bytes.data(), bytes.size()), vector.encoded);

    std::vector<std::uint8_t> decoded;
    ASSERT_TRUE(base64_decode(vector.encoded, decoded)) << vector.encoded;
    EXPECT_EQ(decoded, bytes) << vector.encoded;
  }
}

TEST(XlsbProtectionCodec, RejectsNonzeroUnusedPaddingBits) {
  constexpr std::string_view kAlphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
  std::vector<std::uint8_t> decoded;

  // With two padding characters, the low four bits of the second sextet
  // do not contribute to the decoded byte and must be zero.
  for (std::uint32_t unused_bits = 1U; unused_bits < 16U; ++unused_bits) {
    std::string invalid = "Zg==";
    invalid[1] = kAlphabet[32U | unused_bits];
    EXPECT_FALSE(base64_decode(invalid, decoded)) << invalid;
  }

  // With one padding character, the low two bits of the third sextet
  // must be zero.
  for (std::uint32_t unused_bits = 1U; unused_bits < 4U; ++unused_bits) {
    std::string invalid = "Zm8=";
    invalid[2] = kAlphabet[60U | unused_bits];
    EXPECT_FALSE(base64_decode(invalid, decoded)) << invalid;
  }
}

TEST(XlsbProtectionCodec, RejectsMalformedAlphabetAndPadding) {
  for (const char* invalid : {"Zg=", "Z=g=", "====", "Zg=A", "Zg==AAAA", "!!!!"}) {
    std::vector<std::uint8_t> decoded;
    EXPECT_FALSE(base64_decode(invalid, decoded)) << invalid;
  }
}

TEST(XlsbProtectionCodec, SheetProtectionRejectsInvalidHashOrSaltWithoutAppending) {
  for (const bool invalid_hash : {true, false}) {
    SheetProtection protection;
    protection.enabled = true;
    protection.algorithm_name = "SHA-512";
    protection.hash_value = invalid_hash ? "Zh==" : "Zg==";
    protection.salt_value = invalid_hash ? "Zm8=" : "Zm9=";
    const std::vector<std::uint8_t> original = {0x12U, 0x34U};
    std::vector<std::uint8_t> output = original;

    auto emitted = emit_sheet_protection(output, protection);
    ASSERT_FALSE(static_cast<bool>(emitted));
    EXPECT_EQ(emitted.error().code, FormulonErrorCode::kInvalidArgument);
    EXPECT_EQ(output, original);
  }
}

TEST(XlsbProtectionCodec, BookProtectionRejectsInvalidHashOrSaltWithoutAppending) {
  for (const bool invalid_hash : {true, false}) {
    const std::string hash = invalid_hash ? "Zh==" : "Zg==";
    const std::string salt = invalid_hash ? "Zm8=" : "Zm9=";
    const std::string xml = "<workbookProtection workbookAlgorithmName=\"SHA-512\" workbookHashValue=\"" + hash +
                            "\" workbookSaltValue=\"" + salt + "\" workbookSpinCount=\"1\"/>";
    const std::vector<std::uint8_t> original = {0x56U, 0x78U};
    std::vector<std::uint8_t> output = original;

    auto emitted = emit_book_protection(output, xml);
    ASSERT_FALSE(static_cast<bool>(emitted));
    EXPECT_EQ(emitted.error().code, FormulonErrorCode::kInvalidArgument);
    EXPECT_EQ(output, original);
  }
}

// The .xlsb records decode to the element Excel wrote to the .xlsx twin.
TEST(XlsbBookProtectionFixture, RecordsReadAsTheXlsxTwin) {
  for (const char* name : kCases) {
    const std::string base(name);
    EXPECT_EQ(LoadXlsb(base + ".xlsb").workbook_protection_xml(), LoadXlsx(base + ".xlsx").workbook_protection_xml())
        << name;
  }
  EXPECT_EQ(LoadXlsx("com_s1_struct_win_nopw.xlsx").workbook_protection_xml(),
            "<workbookProtection lockStructure=\"1\"/>");
  EXPECT_EQ(LoadXlsx("hand_lockWindows.xlsx").workbook_protection_xml(), "");
}

// Written from the .xlsx twin, the records are the bytes Excel wrote.
TEST(XlsbBookProtectionFixture, RecordsWriteAsExcel) {
  for (const char* name : kCases) {
    const std::string base(name);
    auto saved = write_xlsb_with_result(LoadXlsx(base + ".xlsx"));
    ASSERT_TRUE(static_cast<bool>(saved)) << name;
    EXPECT_EQ(BookProtectionRecords(saved.value().bytes), BookProtectionRecords(ReadFixture(base + ".xlsb"))) << name;
    EXPECT_EQ(ProtectionDeferrals(LoadXlsx(base + ".xlsx")), 0U) << name;
  }
}

// A hand-written element carrying the fields Excel drops on its own save
// keeps them in BrtBookProtection's protpwdRev and BookProtectionFlags, and
// reads back as the element it was written from.
TEST(XlsbBookProtectionFixture, FieldsExcelDropsSurviveOurSave) {
  struct Case {
    const char* name;
    std::vector<std::uint8_t> payload;
  };
  const Case cases[] = {{"hand_lockAll", {0x00, 0x00, 0x00, 0x00, 0x07, 0x00}},
                        {"hand_lockRevision", {0x00, 0x00, 0x00, 0x00, 0x04, 0x00}},
                        {"hand_lockWindows", {0x00, 0x00, 0x00, 0x00, 0x02, 0x00}},
                        {"hand_revisionsPassword", {0x00, 0x00, 0x4B, 0xCC, 0x00, 0x00}}};
  for (const Case& c : cases) {
    const std::string base(c.name);
    const Workbook source = LoadXlsx(base + ".input.xlsx");
    auto saved = write_xlsb_with_result(source);
    ASSERT_TRUE(static_cast<bool>(saved)) << c.name;
    const auto records = BookProtectionRecords(saved.value().bytes);
    EXPECT_EQ(records.count(677U), 0U) << c.name;
    ASSERT_EQ(records.count(534U), 1U) << c.name;
    EXPECT_EQ(records.at(534U), c.payload) << c.name;
    EXPECT_EQ(ProtectionDeferrals(LoadXlsx(base + ".input.xlsx")), 0U) << c.name;
    auto back = read_xlsb(ByteSpan{saved.value().bytes.data(), saved.value().bytes.size()});
    ASSERT_TRUE(static_cast<bool>(back)) << c.name;
    EXPECT_EQ(back.value().workbook.workbook_protection_xml(), source.workbook_protection_xml()) << c.name;
  }
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
