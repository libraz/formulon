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
// both formats: no BrtBookProtection bit or field for them is ever written.

#include <cstdint>
#include <cstdio>
#include <map>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
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

// A hand-written element carrying the flags Excel drops loses them in our
// .xlsb exactly as in Excel's, and the loss is reported.
TEST(XlsbBookProtectionFixture, DroppedFlagsMatchExcelsOwnSave) {
  for (const char* name : {"hand_lockAll", "hand_lockRevision", "hand_lockWindows", "hand_revisionsPassword"}) {
    const std::string base(name);
    auto saved = write_xlsb_with_result(LoadXlsx(base + ".input.xlsx"));
    ASSERT_TRUE(static_cast<bool>(saved)) << name;
    EXPECT_EQ(BookProtectionRecords(saved.value().bytes), BookProtectionRecords(ReadFixture(base + ".xlsb"))) << name;
    EXPECT_EQ(ProtectionDeferrals(LoadXlsx(base + ".input.xlsx")), 1U) << name;
  }
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
