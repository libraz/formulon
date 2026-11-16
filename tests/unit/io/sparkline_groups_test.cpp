// A sparkline stored over a rectangle loads as its top-left cell; a group with a location that is not one range
// never reaches the model: the .xlsx read drops it and counts it in `ReadDiagnostics::skipped_feature_count`.

#include "io/sparkline_groups.h"

#include <cstddef>
#include <cstdint>
#include <fstream>
#include <iterator>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "sheet.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace {

constexpr std::string_view kSheetPart = "xl/worksheets/sheet1.xml";

void ReplaceFirst(std::string& text, std::string_view from, std::string_view to) {
  const std::size_t at = text.find(from);
  ASSERT_NE(at, std::string::npos) << from;
  text.replace(at, from.size(), to);
}

/// A copy of `src` whose first sheet part has each `{from, to}` pair applied once.
std::vector<std::uint8_t> RewriteSheet(const std::vector<std::uint8_t>& src,
                                       const std::vector<std::pair<std::string, std::string>>& edits) {
  mz_zip_archive reader{};
  EXPECT_NE(mz_zip_reader_init_mem(&reader, src.data(), src.size(), 0), MZ_FALSE);
  mz_zip_archive writer{};
  EXPECT_NE(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_FALSE);
  const mz_uint count = mz_zip_reader_get_num_files(&reader);
  for (mz_uint i = 0; i < count; ++i) {
    char name[256] = {};
    mz_zip_reader_get_filename(&reader, i, name, sizeof(name));
    std::size_t size = 0;
    void* data = mz_zip_reader_extract_to_heap(&reader, i, &size, 0);
    EXPECT_NE(data, nullptr) << name;
    std::string bytes(static_cast<const char*>(data), size);
    mz_free(data);
    if (std::string_view(name) == kSheetPart) {
      for (const auto& edit : edits) {
        ReplaceFirst(bytes, edit.first, edit.second);
      }
    }
    EXPECT_NE(
        mz_zip_writer_add_mem(&writer, name, bytes.data(), bytes.size(), static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
        MZ_FALSE);
  }
  mz_zip_reader_end(&reader);
  void* archive = nullptr;
  std::size_t archive_size = 0;
  EXPECT_NE(mz_zip_writer_finalize_heap_archive(&writer, &archive, &archive_size), MZ_FALSE);
  mz_zip_writer_end(&writer);
  std::vector<std::uint8_t> out(static_cast<const std::uint8_t*>(archive),
                                static_cast<const std::uint8_t*>(archive) + archive_size);
  mz_free(archive);
  return out;
}

/// A saved book with one group per entry of `locations`, in order, each over its own row of column A:D.
std::vector<std::uint8_t> SavedBook(const std::vector<std::string>& locations) {
  Workbook wb = Workbook::create();
  for (std::size_t i = 0; i < locations.size(); ++i) {
    SparklineGroup group;
    group.items.push_back({"Sheet1!A" + std::to_string(i + 2U) + ":D" + std::to_string(i + 2U), locations[i]});
    auto added = add_sheet_sparkline_group(wb, 0U, group);
    EXPECT_TRUE(static_cast<bool>(added));
  }
  auto saved = write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(saved));
  return saved ? std::move(saved.value()) : std::vector<std::uint8_t>{};
}

OoxmlReadResult Read(const std::vector<std::uint8_t>& bytes) {
  auto read = read_ooxml(ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(read));
  return std::move(read.value());
}

void ExpectLocations(const Workbook& wb, const std::vector<std::string>& expected) {
  const std::vector<SparklineGroup> groups = read_sparkline_groups(wb.sheet(0).ext_lst_xml());
  ASSERT_EQ(groups.size(), expected.size());
  for (std::size_t i = 0; i < groups.size(); ++i) {
    ASSERT_EQ(groups[i].items.size(), 1U);
    EXPECT_EQ(groups[i].items[0].location, expected[i]);
  }
}

std::vector<std::uint8_t> ReadFile(const std::string& path) {
  std::ifstream in(path, std::ios::binary);
  EXPECT_TRUE(in.is_open()) << path;
  return std::vector<std::uint8_t>(std::istreambuf_iterator<char>(in), std::istreambuf_iterator<char>());
}

TEST(SparklineGroupsReadTest, ARectangleLoadsAsItsTopLeftCellWithoutADiagnostic) {
  OoxmlReadResult read =
      Read(RewriteSheet(SavedBook({"F2", "B2"}), {{"<xm:sqref>B2</xm:sqref>", "<xm:sqref>B3:C4</xm:sqref>"}}));
  ExpectLocations(read.workbook, {"F2", "B3"});
  EXPECT_EQ(read.diagnostics.skipped_feature_count, 0U);
}

TEST(SparklineGroupsReadTest, GroupsWithALocationThatIsNotOneRangeAreDroppedAndCounted) {
  const std::vector<std::uint8_t> crafted =
      RewriteSheet(SavedBook({"F2", "B2", "C2", "D2"}), {{"<xm:sqref>B2</xm:sqref>", "<xm:sqref>A1:B2</xm:sqref>"},
                                                         {"<xm:sqref>C2</xm:sqref>", "<xm:sqref>zz</xm:sqref>"},
                                                         {"<xm:sqref>D2</xm:sqref>", "<xm:sqref>A1 C3</xm:sqref>"}});
  OoxmlReadResult read = Read(crafted);
  ExpectLocations(read.workbook, {"F2", "A1"});
  EXPECT_EQ(read.diagnostics.skipped_feature_count, 2U);

  auto saved = write_ooxml(read.workbook);
  ASSERT_TRUE(static_cast<bool>(saved));
  OoxmlReadResult reloaded = Read(saved.value());
  ExpectLocations(reloaded.workbook, {"F2", "A1"});
  EXPECT_EQ(reloaded.diagnostics.skipped_feature_count, 0U);
}

TEST(SparklineGroupsReadTest, AnOnlyUnreadableGroupLeavesNoSparklineBlock) {
  OoxmlReadResult read =
      Read(RewriteSheet(SavedBook({"B2"}), {{"<xm:sqref>B2</xm:sqref>", "<xm:sqref>zz</xm:sqref>"}}));
  EXPECT_TRUE(read_sparkline_groups(read.workbook.sheet(0).ext_lst_xml()).empty());
  EXPECT_EQ(read.workbook.sheet(0).ext_lst_xml().find("sparklineGroups"), std::string::npos);
  EXPECT_EQ(read.diagnostics.skipped_feature_count, 1U);
}

TEST(SparklineGroupsReadTest, TheSetterStillRefusesAMultiCellLocation) {
  Workbook wb = Workbook::create();
  SparklineGroup group;
  group.items.push_back({"Sheet1!A2:D2", "A1:B2"});
  EXPECT_FALSE(static_cast<bool>(add_sheet_sparkline_group(wb, 0U, group)));
}

TEST(SparklineGroupsReadTest, AnXlsbRectangleLocationLoadsAsItsTopLeftCell) {
  const std::vector<std::uint8_t> excel =
      ReadFile(std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_parity/sparklines/multicell_location.xlsb");
  auto read = xlsb::read_xlsb(ByteSpan{excel.data(), excel.size()});
  ASSERT_TRUE(static_cast<bool>(read));
  EXPECT_EQ(read.value().skipped_feature_count, 0U);
  ExpectLocations(read.value().workbook, {"G1"});
}

TEST(SparklineGroupsReadTest, AnXlsxRectangleIsWrittenToXlsbAsItsTopLeftCell) {
  OoxmlReadResult read =
      Read(RewriteSheet(SavedBook({"B2"}), {{"<xm:sqref>B2</xm:sqref>", "<xm:sqref>A1:B2</xm:sqref>"}}));
  ExpectLocations(read.workbook, {"A1"});
  auto saved = xlsb::write_xlsb(read.workbook);
  ASSERT_TRUE(static_cast<bool>(saved));
  auto back = xlsb::read_xlsb(ByteSpan{saved.value().data(), saved.value().size()});
  ASSERT_TRUE(static_cast<bool>(back));
  EXPECT_EQ(back.value().skipped_feature_count, 0U);
  ExpectLocations(back.value().workbook, {"A1"});
}

}  // namespace
}  // namespace io
}  // namespace formulon
