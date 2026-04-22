// miniz_smoke_test.cpp
//
// Verifies that the FetchContent-vendored miniz library is reachable from
// Formulon and that the raw deflate and in-memory ZIP APIs we rely on for
// OOXML I/O are linked correctly.

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>

#include "gtest/gtest.h"
#include "miniz.h"

namespace {

TEST(MinizSmoke, DeflateInflateRoundTrip) {
  // A payload that compresses well, with some non-ASCII bytes to ensure the
  // code path does not assume text.
  std::string input;
  input.reserve(1024);
  for (int i = 0; i < 256; ++i) {
    input.append("formulon-miniz-smoke-");
    input.push_back(static_cast<char>(i & 0xFF));
  }

  const unsigned long src_len = static_cast<unsigned long>(input.size());
  unsigned long compressed_cap = compressBound(src_len);
  std::string compressed(compressed_cap, '\0');
  unsigned long compressed_len = compressed_cap;

  int status = compress(reinterpret_cast<unsigned char*>(compressed.data()), &compressed_len,
                        reinterpret_cast<const unsigned char*>(input.data()), src_len);
  ASSERT_EQ(status, Z_OK);
  ASSERT_GT(compressed_len, 0u);
  ASSERT_LE(compressed_len, compressed_cap);
  compressed.resize(compressed_len);

  std::string roundtrip(input.size(), '\0');
  unsigned long out_len = static_cast<unsigned long>(roundtrip.size());
  status = uncompress(reinterpret_cast<unsigned char*>(roundtrip.data()), &out_len,
                      reinterpret_cast<const unsigned char*>(compressed.data()), compressed_len);
  ASSERT_EQ(status, Z_OK);
  ASSERT_EQ(out_len, src_len);
  EXPECT_EQ(roundtrip, input);
}

TEST(MinizSmoke, InMemoryZipRoundTrip) {
  // Simulates the OOXML use case: build a ZIP entirely in memory, then read
  // it back and verify the entry content.
  mz_zip_archive writer{};
  ASSERT_TRUE(mz_zip_writer_init_heap(&writer, /*size_to_reserve_at_beginning=*/0,
                                      /*initial_allocation_size=*/4 * 1024));

  constexpr std::string_view kEntryName = "xl/workbook.xml";
  constexpr std::string_view kEntryBody =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
      "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>";

  ASSERT_TRUE(mz_zip_writer_add_mem(&writer, kEntryName.data(), kEntryBody.data(), kEntryBody.size(),
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)));

  void* archive_ptr = nullptr;
  std::size_t archive_size = 0;
  ASSERT_TRUE(mz_zip_writer_finalize_heap_archive(&writer, &archive_ptr, &archive_size));
  ASSERT_TRUE(mz_zip_writer_end(&writer));
  ASSERT_NE(archive_ptr, nullptr);
  ASSERT_GT(archive_size, 0u);

  mz_zip_archive reader{};
  ASSERT_TRUE(mz_zip_reader_init_mem(&reader, archive_ptr, archive_size, 0));
  const mz_uint entry_count = mz_zip_reader_get_num_files(&reader);
  ASSERT_EQ(entry_count, 1u);

  char name_buf[64] = {};
  const mz_uint name_len = mz_zip_reader_get_filename(&reader, 0, name_buf, sizeof(name_buf));
  ASSERT_GT(name_len, 0u);
  // mz_zip_reader_get_filename returns length including trailing NUL.
  EXPECT_EQ(std::string(name_buf), std::string(kEntryName));

  std::size_t uncompressed_size = 0;
  void* extracted = mz_zip_reader_extract_to_heap(&reader, 0, &uncompressed_size, /*flags=*/0);
  ASSERT_NE(extracted, nullptr);
  ASSERT_EQ(uncompressed_size, kEntryBody.size());
  EXPECT_EQ(std::memcmp(extracted, kEntryBody.data(), kEntryBody.size()), 0);

  mz_free(extracted);
  ASSERT_TRUE(mz_zip_reader_end(&reader));
  mz_free(archive_ptr);
}

}  // namespace
