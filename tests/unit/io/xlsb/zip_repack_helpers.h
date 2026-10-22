#pragma once

#include <cstdint>
#include <cstdio>
#include <map>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "miniz.h"

namespace formulon::io::xlsb {

/// Re-zips `src` with the entries named in `replaced` swapped for new bytes.
inline std::vector<std::uint8_t> RepackWithEntries(const std::vector<std::uint8_t>& src,
                                                   const std::map<std::string, std::vector<std::uint8_t>>& replaced) {
  mz_zip_archive reader{};
  EXPECT_NE(mz_zip_reader_init_mem(&reader, src.data(), src.size(), 0), MZ_FALSE);
  mz_zip_archive writer{};
  EXPECT_NE(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_FALSE);
  const mz_uint count = mz_zip_reader_get_num_files(&reader);
  for (mz_uint i = 0; i < count; ++i) {
    char name[256] = {};
    EXPECT_GT(mz_zip_reader_get_filename(&reader, i, name, sizeof(name)), 0U);
    const auto hit = replaced.find(name);
    if (hit != replaced.end()) {
      EXPECT_NE(mz_zip_writer_add_mem(&writer, name, hit->second.data(), hit->second.size(),
                                      static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
                MZ_FALSE);
      continue;
    }
    std::size_t size = 0;
    void* data = mz_zip_reader_extract_to_heap(&reader, i, &size, 0);
    EXPECT_NE(data, nullptr);
    EXPECT_NE(mz_zip_writer_add_mem(&writer, name, data, size, static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)), MZ_FALSE);
    mz_free(data);
  }
  mz_zip_reader_end(&reader);
  void* archive = nullptr;
  std::size_t archive_size = 0;
  EXPECT_NE(mz_zip_writer_finalize_heap_archive(&writer, &archive, &archive_size), MZ_FALSE);
  EXPECT_NE(mz_zip_writer_end(&writer), MZ_FALSE);
  std::vector<std::uint8_t> out(static_cast<const std::uint8_t*>(archive),
                                static_cast<const std::uint8_t*>(archive) + archive_size);
  mz_free(archive);
  return out;
}

inline std::vector<std::uint8_t> ReadFixtureFile(const std::string& path) {
  std::vector<std::uint8_t> out;
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
    out.clear();
  }
  std::fclose(f);
  return out;
}

}  // namespace formulon::io::xlsb
