//
// Unit tests for `formulon::io::ooxml::load_pivot_cache_records_target`,
// which resolves a pivot cache definition's records part from its rels.

#include "io/ooxml/pivot_target_reader.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "io/zip_reader.h"
#include "miniz.h"

namespace formulon::io::ooxml {
namespace {

constexpr std::string_view kDefinitionPath = "xl/pivotCache/pivotCacheDefinition1.xml";
constexpr std::string_view kRelsPath = "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels";
constexpr std::string_view kRecordsType =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords";

std::vector<std::uint8_t> ZipWithRels(const std::string& rels_xml) {
  mz_zip_archive writer{};
  EXPECT_NE(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_FALSE);
  EXPECT_NE(mz_zip_writer_add_mem(&writer, std::string(kRelsPath).c_str(), rels_xml.data(), rels_xml.size(),
                                  static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
            MZ_FALSE);
  void* archive = nullptr;
  std::size_t archive_size = 0;
  EXPECT_NE(mz_zip_writer_finalize_heap_archive(&writer, &archive, &archive_size), MZ_FALSE);
  EXPECT_NE(mz_zip_writer_end(&writer), MZ_FALSE);
  std::vector<std::uint8_t> out(static_cast<const std::uint8_t*>(archive),
                                static_cast<const std::uint8_t*>(archive) + archive_size);
  mz_free(archive);
  return out;
}

std::string RecordsRel(std::string_view id, std::string_view target) {
  std::string xml("<Relationship Id=\"");
  xml.append(id).append("\" Type=\"").append(kRecordsType).append("\" Target=\"").append(target).append("\"/>");
  return xml;
}

std::string RelsDocument(const std::string& body) {
  return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
         "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
         body + "</Relationships>";
}

TEST(PivotTargetReader, ResolvesTheRecordsTargetBesideTheDefinition) {
  const std::vector<std::uint8_t> bytes = ZipWithRels(RelsDocument(RecordsRel("rId1", "pivotCacheRecords1.xml")));
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(ByteSpan{bytes.data(), bytes.size()})));
  auto target_or = load_pivot_cache_records_target(zip, kDefinitionPath);
  ASSERT_TRUE(static_cast<bool>(target_or)) << target_or.error().message;
  EXPECT_EQ(target_or.value(), "xl/pivotCache/pivotCacheRecords1.xml");
}

TEST(PivotTargetReader, DuplicatedRecordsRelResolvesToTheLastEntry) {
  // Matches the workbook and sheet rels readers' last-writer rule.
  const std::vector<std::uint8_t> bytes =
      ZipWithRels(RelsDocument(RecordsRel("rId1", "staleRecords.xml") + RecordsRel("rId2", "pivotCacheRecords1.xml")));
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(ByteSpan{bytes.data(), bytes.size()})));
  auto target_or = load_pivot_cache_records_target(zip, kDefinitionPath);
  ASSERT_TRUE(static_cast<bool>(target_or)) << target_or.error().message;
  EXPECT_EQ(target_or.value(), "xl/pivotCache/pivotCacheRecords1.xml");
}

}  // namespace
}  // namespace formulon::io::ooxml
