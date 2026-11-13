// Formulon .xlsb outputs written by an earlier Formulon build (tests/fixtures/excel/own_output/) keep loading,
// with no feature skipped.

#include <filesystem>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "io/xlsb/reader.h"
#include "xlsb_twin_compare.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon::io::xlsb::test {
namespace {

TEST(XlsbOwnOutput, LoadsWithoutSkips) {
  const std::string dir = std::string(FORMULON_FIXTURES_DIR) + "/excel/own_output";
  std::vector<std::string> paths;
  for (const auto& entry : std::filesystem::directory_iterator(dir)) {
    if (entry.path().extension() == ".xlsb") {
      paths.push_back(entry.path().string());
    }
  }
  ASSERT_FALSE(paths.empty());
  for (const std::string& path : paths) {
    const Bytes bytes = ReadBytes(path);
    ASSERT_FALSE(bytes.empty()) << path;
    auto loaded = read_xlsb(ByteSpan{bytes.data(), bytes.size()});
    ASSERT_TRUE(static_cast<bool>(loaded)) << path << ": " << loaded.error().message;
    EXPECT_EQ(loaded.value().skipped_feature_count, 0U) << path;
  }
}

}  // namespace
}  // namespace formulon::io::xlsb::test
