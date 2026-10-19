// Package loaders read relationship parts by namespace, so a prefixed
// `pkg:Relationships` root loads and a foreign-namespace sibling is ignored.

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "cell.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "miniz.h"
#include "sheet.h"
#include "support/roundtrip_symmetry.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

void ReplaceAll(std::string& text, std::string_view from, std::string_view to) {
  for (std::size_t at = text.find(from); at != std::string::npos; at = text.find(from, at + to.size())) {
    text.replace(at, from.size(), to);
  }
}

/// Rewrites every `.rels` entry with a `pkg:` prefix and a leading foreign
/// `Relationship` that would redirect the office document if it were read.
std::vector<std::uint8_t> PrefixRels(const std::vector<std::uint8_t>& src) {
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
    const std::string_view entry(name);
    if (entry.size() > 5U && entry.substr(entry.size() - 5U) == ".rels") {
      const std::string before = bytes;
      ReplaceAll(bytes, "<Relationships xmlns=", "<pkg:Relationships xmlns:pkg=");
      ReplaceAll(bytes, "</Relationships>", "</pkg:Relationships>");
      ReplaceAll(bytes, "<Relationship ", "<pkg:Relationship ");
      const std::size_t body = bytes.find("<pkg:Relationship ");
      EXPECT_NE(body, std::string::npos) << name;
      if (body != std::string::npos) {
        bytes.insert(body,
                     "<f:Relationship xmlns:f=\"urn:foreign\" Id=\"rIdForeign\" "
                     "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" "
                     "Target=\"missing/part.xml\"/>");
      }
      EXPECT_NE(bytes, before) << name;
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

Workbook Source() {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0U, 0U, Value::number(42.0));
  return wb;
}

void ExpectA1(const Workbook& wb) {
  ASSERT_EQ(wb.sheet_count(), 1U);
  const Cell* a1 = wb.sheet(0).cell_at(0U, 0U);
  ASSERT_NE(a1, nullptr);
  ASSERT_TRUE(a1->cached_value.is_number());
  EXPECT_EQ(a1->cached_value.as_number(), 42.0);
}

TEST(PrefixedPackageRels, XlsxLoads) {
  auto saved = Source().save();
  ASSERT_TRUE(static_cast<bool>(saved));
  const std::vector<std::uint8_t> pkg = PrefixRels(saved.value());
  auto loaded = io::read_ooxml(test::span_of(pkg));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  ExpectA1(loaded.value().workbook);
}

TEST(PrefixedPackageRels, XlsbLoads) {
  auto saved = io::xlsb::write_xlsb(Source());
  ASSERT_TRUE(static_cast<bool>(saved));
  const std::vector<std::uint8_t> pkg = PrefixRels(saved.value());
  auto loaded = io::xlsb::read_xlsb(test::span_of(pkg));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  ExpectA1(loaded.value().workbook);
}

}  // namespace
}  // namespace formulon
