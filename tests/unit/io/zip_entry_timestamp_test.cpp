//
// Saved packages carry a fixed entry timestamp, so saving an unchanged model
// twice yields identical bytes however far apart the saves are.

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <unordered_set>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml/zip_part_writer.h"
#include "io/zip_reader.h"
#include "value.h"
#include "workbook.h"
#include "workbook_format.h"

namespace formulon {
namespace {

// 1980-01-01 00:00 in MS-DOS date / time encoding.
constexpr std::uint16_t kDosDate = 0x0021U;
constexpr std::uint16_t kDosTime = 0x0000U;

std::uint16_t U16(const std::vector<std::uint8_t>& bytes, std::size_t at) {
  return static_cast<std::uint16_t>(bytes[at] | (bytes[at + 1U] << 8U));
}

std::uint32_t U32(const std::vector<std::uint8_t>& bytes, std::size_t at) {
  return static_cast<std::uint32_t>(U16(bytes, at)) | (static_cast<std::uint32_t>(U16(bytes, at + 2U)) << 16U);
}

// Checks the date and time of every central-directory entry and of the local
// header it points to; returns the number of entries seen.
std::size_t ExpectFixedTimestamps(const std::vector<std::uint8_t>& zip) {
  constexpr std::size_t kEocdSize = 22U;
  EXPECT_GE(zip.size(), kEocdSize);
  const std::size_t eocd = zip.size() - kEocdSize;
  EXPECT_EQ(U32(zip, eocd), 0x06054b50U);
  const std::size_t entries = U16(zip, eocd + 10U);
  std::size_t at = U32(zip, eocd + 16U);
  for (std::size_t i = 0; i < entries; ++i) {
    EXPECT_EQ(U32(zip, at), 0x02014b50U);
    EXPECT_EQ(U16(zip, at + 12U), kDosTime);
    EXPECT_EQ(U16(zip, at + 14U), kDosDate);
    const std::size_t local = U32(zip, at + 42U);
    EXPECT_EQ(U32(zip, local), 0x04034b50U);
    EXPECT_EQ(U16(zip, local + 10U), kDosTime);
    EXPECT_EQ(U16(zip, local + 12U), kDosDate);
    at += 46U + U16(zip, at + 28U) + U16(zip, at + 30U) + U16(zip, at + 32U);
  }
  return entries;
}

class ZipEntryTimestampTest : public ::testing::TestWithParam<WorkbookFormat> {};

TEST_P(ZipEntryTimestampTest, EveryEntryCarriesTheDosEpoch) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(wb.set_cell_value(0, 0, 0, Value::number(1.5)));
  auto saved = wb.save_as(GetParam());
  ASSERT_TRUE(saved) << saved.error().message;
  EXPECT_GT(ExpectFixedTimestamps(saved.value()), 0U);
}

TEST_P(ZipEntryTimestampTest, UnchangedModelSavesIdenticalBytes) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(wb.set_cell_value(0, 0, 0, Value::number(1.5)));
  auto first = wb.save_as(GetParam());
  auto second = wb.save_as(GetParam());
  ASSERT_TRUE(first && second);
  EXPECT_EQ(first.value(), second.value());

  ASSERT_TRUE(wb.set_cell_value(0, 0, 0, Value::number(2.5)));
  auto changed = wb.save_as(GetParam());
  ASSERT_TRUE(changed);
  EXPECT_NE(first.value(), changed.value());
}

TEST(ZipPartWriter, RejectsEmbeddedNulBeforeRecordingAndWritesValidNames) {
  io::ZipWriterGuard writer;
  ASSERT_TRUE(writer.init());

  std::unordered_set<std::string> seen_paths{"preexisting.txt"};
  const std::string text_body = "text";
  auto bad_text = io::AddPart(writer.get(), std::string_view{"plain.xml\0suffix", 16U}, text_body, &seen_paths);
  EXPECT_FALSE(static_cast<bool>(bad_text));
  if (!bad_text) {
    EXPECT_EQ(bad_text.error().code, FormulonErrorCode::kIoWriteFailed);
  }
  EXPECT_EQ(seen_paths, (std::unordered_set<std::string>{"preexisting.txt"}));

  const std::vector<std::uint8_t> binary_body{0x01U, 0x02U};
  auto bad_binary = io::AddPartBytes(writer.get(), std::string_view{"blob.bin\0tail", 13U}, binary_body, &seen_paths);
  EXPECT_FALSE(static_cast<bool>(bad_binary));
  if (!bad_binary) {
    EXPECT_EQ(bad_binary.error().code, FormulonErrorCode::kIoWriteFailed);
  }
  EXPECT_EQ(seen_paths, (std::unordered_set<std::string>{"preexisting.txt"}));

  ASSERT_TRUE(static_cast<bool>(io::AddPart(writer.get(), "plain.xml", text_body, &seen_paths)));
  ASSERT_TRUE(static_cast<bool>(io::AddPartBytes(writer.get(), "blob.bin", binary_body, &seen_paths)));
  ASSERT_TRUE(static_cast<bool>(io::AddPart(writer.get(), "Alpha.txt", text_body, &seen_paths)));
  ASSERT_TRUE(static_cast<bool>(io::AddPart(writer.get(), "alpha.txt", text_body, &seen_paths)));
  ASSERT_TRUE(static_cast<bool>(io::AddPartBytes(writer.get(), "資料/値.bin", binary_body, &seen_paths)));
  EXPECT_EQ(seen_paths.size(), 6U);
  EXPECT_NE(seen_paths.find("plain.xml"), seen_paths.end());
  EXPECT_NE(seen_paths.find("blob.bin"), seen_paths.end());
  EXPECT_NE(seen_paths.find("Alpha.txt"), seen_paths.end());
  EXPECT_NE(seen_paths.find("alpha.txt"), seen_paths.end());
  EXPECT_NE(seen_paths.find("資料/値.bin"), seen_paths.end());

  auto archive = io::FinalizeArchive(writer, "ZipPartWriter test");
  ASSERT_TRUE(static_cast<bool>(archive)) << archive.error().message;
  io::ZipReader reader;
  ASSERT_TRUE(static_cast<bool>(reader.open(io::ByteSpan{archive.value().data(), archive.value().size()})));
  EXPECT_TRUE(reader.has_entry("plain.xml"));
  EXPECT_TRUE(reader.has_entry("blob.bin"));
  EXPECT_TRUE(reader.has_entry("Alpha.txt"));
  EXPECT_TRUE(reader.has_entry("alpha.txt"));
  EXPECT_TRUE(reader.has_entry("資料/値.bin"));
  EXPECT_FALSE(reader.has_entry(std::string_view{"plain.xml\0suffix", 16U}));
  EXPECT_EQ(reader.entry_count(), 5U);
}

INSTANTIATE_TEST_SUITE_P(Formats, ZipEntryTimestampTest,
                         ::testing::Values(WorkbookFormat::Ooxml, WorkbookFormat::Xlsb));

}  // namespace
}  // namespace formulon
