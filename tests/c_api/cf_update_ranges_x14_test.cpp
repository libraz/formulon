// A rule update that sets new ranges carries the rule's x14 extension entry to the same ranges, so the saved
// `<x14:conditionalFormatting>` names the cells the legacy block names. `cf_databar_union_sqref.xlsx` is an
// Excel-saved book with two single-rule data-bar blocks, each linked to its own x14 entry.

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <map>
#include <string>
#include <utility>
#include <vector>

#include "formulon_c_cf_test_helpers.h"
#include "io/zip_reader.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace {

constexpr const char* kFirstId = "{11111111-2222-4333-8444-555555555555}";
constexpr const char* kSecondId = "{66666666-7777-4888-8999-AAAAAAAAAAAA}";

std::vector<std::uint8_t> ReadFixture(const std::string& name) {
  std::vector<std::uint8_t> out;
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/" + name;
  FILE* file = std::fopen(path.c_str(), "rb");
  if (file == nullptr) {
    ADD_FAILURE() << "could not open fixture: " << path;
    return out;
  }
  std::fseek(file, 0, SEEK_END);
  const long size = std::ftell(file);
  std::fseek(file, 0, SEEK_SET);
  if (size > 0) {
    out.resize(static_cast<std::size_t>(size));
    if (std::fread(out.data(), 1, out.size(), file) != out.size()) {
      ADD_FAILURE() << "short read on fixture: " << path;
      out.clear();
    }
  }
  std::fclose(file);
  return out;
}

std::string SavedSheetXml(fm_workbook_t* handle) {
  BufferGuard saved;
  EXPECT_EQ(fm_workbook_save(handle, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len), 0)
      << fm_get_last_error_message();
  formulon::io::ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(formulon::io::ByteSpan{saved.data, saved.len})));
  auto entry = zip.read_entry("xl/worksheets/sheet1.xml");
  EXPECT_TRUE(static_cast<bool>(entry));
  return entry ? std::string(entry.value().begin(), entry.value().end()) : std::string();
}

// Each x14 rule id of the sheet XML mapped to the `xm:sqref` of the block holding it.
std::map<std::string, std::string> X14SqrefById(const std::string& xml) {
  std::map<std::string, std::string> out;
  const std::string open = "<x14:conditionalFormatting";
  const std::string close = "</x14:conditionalFormatting>";
  for (std::size_t at = xml.find(open); at != std::string::npos; at = xml.find(open, at + 1U)) {
    const std::size_t end = xml.find(close, at);
    if (end == std::string::npos) {
      break;
    }
    const std::string block = xml.substr(at, end - at);
    const std::size_t sq = block.find("<xm:sqref>");
    const std::string sqref =
        sq == std::string::npos ? std::string() : block.substr(sq + 10U, block.find("</xm:sqref>", sq) - sq - 10U);
    for (std::size_t r = block.find("<x14:cfRule"); r != std::string::npos; r = block.find("<x14:cfRule", r + 1U)) {
      const std::size_t id_at = block.find("id=\"", r);
      if (id_at == std::string::npos) {
        continue;
      }
      const std::size_t id_end = block.find('"', id_at + 4U);
      out[block.substr(id_at + 4U, id_end - id_at - 4U)] = sqref;
    }
  }
  return out;
}

fm_cf_rule_patch_t RangesPatch(const fm_cell_rect_t* ranges, std::uint32_t count) {
  fm_cf_rule_patch_t patch = FM_STRUCT_INIT(fm_cf_rule_patch_t);
  patch.ranges_op = FM_PATCH_OP_SET;
  patch.ranges = ranges;
  patch.range_count = count;
  return patch;
}

TEST(CfUpdateRangesX14, ExcelSavedDataBarEntryFollowsTheNewRanges) {
  const std::vector<std::uint8_t> bytes = ReadFixture("cf_databar_union_sqref.xlsx");
  ASSERT_FALSE(bytes.empty());
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_load(bytes.data(), bytes.size(), nullptr, &wb.handle), 0) << fm_get_last_error_message();
  ASSERT_EQ(X14SqrefById(SavedSheetXml(wb.handle)).at(kFirstId), "A1:A5 C1:C5");

  const fm_cell_rect_t moved[] = {{0, 9, 4, 9}};  // J1:J5
  const fm_cf_rule_patch_t patch = RangesPatch(moved, 1U);
  ASSERT_EQ(fm_sheet_update_cf_rule(wb.handle, 0, kFirstId, &patch), 0) << fm_get_last_error_message();

  const std::map<std::string, std::string> saved = X14SqrefById(SavedSheetXml(wb.handle));
  ASSERT_EQ(saved.count(kFirstId), 1U);
  EXPECT_EQ(saved.at(kFirstId), "J1:J5");
  ASSERT_EQ(saved.count(kSecondId), 1U);
  EXPECT_EQ(saved.at(kSecondId), "E2:E6 G2:G6");
}

TEST(CfUpdateRangesX14, RuleMovedOutOfASharedBlockTakesOnlyItsOwnEntry) {
  const std::vector<std::uint8_t> bytes = ReadFixture("cf_databar_union_sqref.xlsx");
  ASSERT_FALSE(bytes.empty());
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_load(bytes.data(), bytes.size(), nullptr, &wb.handle), 0) << fm_get_last_error_message();
  // Give the second block's range to the first rule so both rules share one block and one x14 entry list.
  const fm_cell_rect_t shared[] = {{1, 4, 5, 4}, {1, 6, 5, 6}};  // E2:E6 G2:G6
  const fm_cf_rule_patch_t join = RangesPatch(shared, 2U);
  ASSERT_EQ(fm_sheet_update_cf_rule(wb.handle, 0, kFirstId, &join), 0) << fm_get_last_error_message();
  std::map<std::string, std::string> saved = X14SqrefById(SavedSheetXml(wb.handle));
  EXPECT_EQ(saved.at(kFirstId), "E2:E6 G2:G6");

  const fm_cell_rect_t moved[] = {{0, 9, 4, 9}};  // J1:J5
  const fm_cf_rule_patch_t patch = RangesPatch(moved, 1U);
  ASSERT_EQ(fm_sheet_update_cf_rule(wb.handle, 0, kFirstId, &patch), 0) << fm_get_last_error_message();
  saved = X14SqrefById(SavedSheetXml(wb.handle));
  EXPECT_EQ(saved.at(kFirstId), "J1:J5");
  EXPECT_EQ(saved.at(kSecondId), "E2:E6 G2:G6");
}

}  // namespace
