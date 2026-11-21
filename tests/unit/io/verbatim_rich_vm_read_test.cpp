// A book whose rich-value parts stay verbatim keeps each cell's `vm` on load:
// Excel's pivot-record pictures in both formats come back with the vm Excel
// saved, a vm naming no value entry or on a cell not caching #VALUE! is
// dropped, the count survives a metadata part the model refuses, and a book
// declined for an external link's vm or an opaque sheet keeps the other
// sheets' vms.

#include <cstddef>
#include <cstdint>
#include <map>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/rich_values.h"
#include "io/xlsb/reader.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "sheet.h"
#include "support/roundtrip_symmetry.h"
#include "utils/a1_ref.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

using VmTable = std::map<std::pair<std::size_t, std::string>, std::uint32_t>;

// Per Excel pair, the vm of every cell carrying one, by (sheet index, A1), as Excel saved it in both formats.
struct Pair {
  const char* name;
  VmTable vms;
};

const std::vector<Pair>& Pairs() {
  static const std::vector<Pair> pairs = {
      {"one_picture", {{{0, "A6"}, 1}, {{1, "A3"}, 1}}},
      {"two_pictures", {{{0, "A6"}, 1}, {{0, "A7"}, 2}, {{1, "A3"}, 1}, {{1, "A5"}, 2}}},
      {"picture_and_text", {{{0, "A6"}, 1}, {{1, "A3"}, 1}}},
      {"seven_items",
       {{{0, "A6"}, 1}, {{0, "A7"}, 2}, {{0, "A10"}, 3}, {{1, "A3"}, 1}, {{1, "A5"}, 2}, {{1, "A8"}, 3}}},
      {"numeric_field", {{{0, "A6"}, 1}, {{1, "A3"}, 1}}},
  };
  return pairs;
}

std::vector<std::uint8_t> ReadFixture(const std::string& relative) {
  std::vector<std::uint8_t> bytes = test::read_file_bytes(std::string(FORMULON_FIXTURES_DIR) + "/" + relative);
  EXPECT_FALSE(bytes.empty()) << relative;
  return bytes;
}

std::vector<std::uint8_t> ReadPair(const char* name, const char* extension) {
  return ReadFixture(std::string("excel/pivot_rich_records/") + name + extension);
}

std::string Part(const std::vector<std::uint8_t>& package, std::string_view part) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(test::span_of(package))));
  auto bytes = zip.read_entry(part);
  EXPECT_TRUE(static_cast<bool>(bytes)) << part;
  return bytes ? std::string(bytes.value().begin(), bytes.value().end()) : std::string();
}

/// `package` with each part of `replaced` swapped for (or added as) its new text.
std::vector<std::uint8_t> Repack(const std::vector<std::uint8_t>& package,
                                 const std::map<std::string, std::string>& replaced) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(test::span_of(package))));
  mz_zip_archive writer{};
  EXPECT_NE(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_FALSE);
  std::map<std::string, std::string> pending = replaced;
  const auto add = [&writer](const std::string& name, const std::vector<std::uint8_t>& bytes) {
    EXPECT_NE(mz_zip_writer_add_mem(&writer, name.c_str(), bytes.data(), bytes.size(),
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_FALSE)
        << name;
  };
  for (const std::string& name : zip.list_entries()) {
    const auto it = pending.find(name);
    if (it != pending.end()) {
      add(name, std::vector<std::uint8_t>(it->second.begin(), it->second.end()));
      pending.erase(it);
      continue;
    }
    auto read = zip.read_entry(name);
    EXPECT_TRUE(static_cast<bool>(read)) << name;
    add(name, read ? read.value() : std::vector<std::uint8_t>{});
  }
  for (const auto& [name, text] : pending) {
    add(name, std::vector<std::uint8_t>(text.begin(), text.end()));
  }
  void* archive = nullptr;
  std::size_t size = 0;
  EXPECT_NE(mz_zip_writer_finalize_heap_archive(&writer, &archive, &size), MZ_FALSE);
  std::vector<std::uint8_t> out(static_cast<const std::uint8_t*>(archive),
                                static_cast<const std::uint8_t*>(archive) + size);
  mz_free(archive);
  mz_zip_writer_end(&writer);
  return out;
}

/// `text` with its one occurrence of `from` replaced by `to`.
std::string ReplaceOnce(std::string text, std::string_view from, std::string_view to) {
  const std::size_t at = text.find(from);
  EXPECT_NE(at, std::string::npos) << from;
  EXPECT_EQ(text.find(from, at + 1U), std::string::npos) << from;
  if (at != std::string::npos) {
    text.replace(at, from.size(), to);
  }
  return text;
}

Workbook Load(const std::vector<std::uint8_t>& bytes) {
  auto loaded = read_ooxml(test::span_of(bytes));
  EXPECT_TRUE(static_cast<bool>(loaded)) << (loaded ? "" : loaded.error().message);
  return loaded ? std::move(loaded.value().workbook) : Workbook::create_empty();
}

Workbook LoadXlsb(const std::vector<std::uint8_t>& bytes) {
  auto loaded = xlsb::read_xlsb(test::span_of(bytes));
  EXPECT_TRUE(static_cast<bool>(loaded)) << (loaded ? "" : loaded.error().message);
  return loaded ? std::move(loaded.value().workbook) : Workbook::create_empty();
}

/// Every opaque rich value of `wb`, by (sheet index, A1): its `verbatim_vm`, with the cell still caching #VALUE!.
VmTable VerbatimVms(const Workbook& wb) {
  VmTable out;
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    for (const CellRichValue& v : wb.sheet(s).rich_values()) {
      const Cell* cell = wb.sheet(s).cell_at(v.row, v.col);
      EXPECT_TRUE(cell != nullptr && cell->cached_value.is_error() &&
                  cell->cached_value.as_error() == ErrorCode::Value);
      EXPECT_TRUE(v.value.structure_type.empty());
      EXPECT_TRUE(v.value.keys.empty());
      out[{s, a1::encode_a1(v.row, v.col)}] = v.value.verbatim_vm;
    }
  }
  return out;
}

TEST(VerbatimRichVmRead, ExcelPivotRecordBooksKeepEveryVm) {
  for (const Pair& pair : Pairs()) {
    SCOPED_TRACE(pair.name);
    const Workbook xlsx = Load(ReadPair(pair.name, ".xlsx"));
    EXPECT_TRUE(rich_values_kept_verbatim(xlsx));
    EXPECT_EQ(VerbatimVms(xlsx), pair.vms);
    const Workbook xlsb = LoadXlsb(ReadPair(pair.name, ".xlsb"));
    EXPECT_TRUE(rich_values_kept_verbatim(xlsb));
    EXPECT_EQ(VerbatimVms(xlsb), pair.vms);
  }
}

TEST(VerbatimRichVmRead, AVmPastTheValueEntriesIsDropped) {
  const std::vector<std::uint8_t> excel = ReadPair("one_picture", ".xlsx");
  const std::string sheet = ReplaceOnce(Part(excel, "xl/worksheets/sheet2.xml"), "<c r=\"A3\" t=\"e\" vm=\"1\">",
                                        "<c r=\"A3\" t=\"e\" vm=\"2\">");
  const Workbook wb = Load(Repack(excel, {{"xl/worksheets/sheet2.xml", sheet}}));
  EXPECT_EQ(VerbatimVms(wb), (VmTable{{{0, "A6"}, 1}}));
}

TEST(VerbatimRichVmRead, AVmOnACellNotCachingValueErrorIsDropped) {
  const std::vector<std::uint8_t> excel = ReadPair("one_picture", ".xlsx");
  const std::string sheet =
      ReplaceOnce(Part(excel, "xl/worksheets/sheet2.xml"), "<c r=\"A3\" t=\"e\" vm=\"1\"><v>#VALUE!</v>",
                  "<c r=\"A3\" t=\"e\" vm=\"1\"><v>#N/A</v>");
  const Workbook wb = Load(Repack(excel, {{"xl/worksheets/sheet2.xml", sheet}}));
  EXPECT_EQ(VerbatimVms(wb), (VmTable{{{0, "A6"}, 1}}));
}

TEST(VerbatimRichVmRead, AMetadataPartTheModelRefusesStillCounts) {
  const std::vector<std::uint8_t> excel = ReadPair("two_pictures", ".xlsx");
  const std::string metadata =
      ReplaceOnce(Part(excel, "xl/metadata.xml"), "<metadataTypes count=\"1\">",
                  "<metadataTypes count=\"2\"><metadataType name=\"XLFUTURE\" minSupportedVersion=\"120000\"/>");
  const Workbook wb = Load(Repack(excel, {{"xl/metadata.xml", metadata}}));
  EXPECT_EQ(VerbatimVms(wb), Pairs()[1].vms);
}

TEST(VerbatimRichVmRead, AnUncountableMetadataPartRestoresNoVm) {
  const std::vector<std::uint8_t> excel = ReadPair("two_pictures", ".xlsx");
  const Workbook wb = Load(Repack(excel, {{"xl/metadata.xml", "<metadata"}}));
  EXPECT_TRUE(VerbatimVms(wb).empty());
}

TEST(VerbatimRichVmRead, AnExternalLinkVmKeepsTheSheetVmsVerbatim) {
  // A book the table would consume, declined because an external link's cached cells carry a vm.
  const std::vector<std::uint8_t> excel = ReadFixture("excel/cell_pictures/cell_pictures.xlsx");
  const std::string link =
      "<externalLink xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><externalBook>"
      "<sheetDataSet><sheetData sheetId=\"0\"><row r=\"1\"><cell r=\"A1\" t=\"e\" vm=\"1\"><v>#VALUE!</v></cell>"
      "</row></sheetData></sheetDataSet></externalBook></externalLink>";
  const Workbook wb = Load(Repack(excel, {{"xl/externalLinks/externalLink1.xml", link}}));
  EXPECT_TRUE(rich_values_kept_verbatim(wb));
  EXPECT_EQ(VerbatimVms(wb), (VmTable{{{0, "E2"}, 1}, {{0, "E4"}, 1}, {{0, "E6"}, 2}, {{0, "E8"}, 3}}));
}

TEST(VerbatimRichVmRead, AnOpaqueSheetKeepsNoVmWhileTheOthersDo) {
  // Sheet index 1 (`sheet2.xml`) related as a dialog sheet is kept opaque.
  const std::vector<std::uint8_t> excel = ReadPair("two_pictures", ".xlsx");
  std::string rels = Part(excel, "xl/_rels/workbook.xml.rels");
  const std::string target = "Target=\"worksheets/sheet2.xml\"";
  const std::size_t at = rels.find(target);
  ASSERT_NE(at, std::string::npos);
  const std::size_t open = rels.rfind("<Relationship ", at);
  const std::string worksheet = "relationships/worksheet\"";
  const std::size_t type = rels.find(worksheet, open);
  ASSERT_TRUE(open != std::string::npos && type != std::string::npos && type < rels.find("/>", at));
  rels.replace(type, worksheet.size(), "relationships/dialogsheet\"");
  const Workbook wb = Load(Repack(excel, {{"xl/_rels/workbook.xml.rels", rels}}));
  ASSERT_TRUE(wb.sheet(1).is_opaque_ooxml_sheet());
  EXPECT_EQ(VerbatimVms(wb), (VmTable{{{0, "A6"}, 1}, {{0, "A7"}, 2}}));
}

}  // namespace
}  // namespace io
}  // namespace formulon
