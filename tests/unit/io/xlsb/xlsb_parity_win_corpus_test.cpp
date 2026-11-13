// Every Windows-Excel .xlsb under tests/fixtures (any path with a `win` directory) loads with no feature
// skipped, which backs reading the Mac-measured record set as the one Windows Excel writes too.

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

bool UnderWinDirectory(const std::filesystem::path& path) {
  for (const auto& component : path.parent_path()) {
    if (component == "win") {
      return true;
    }
  }
  return false;
}

TEST(XlsbParityWinCorpus, NoSkips) {
  const std::filesystem::path root(FORMULON_FIXTURES_DIR);
  std::vector<std::string> paths;
  for (const auto& entry : std::filesystem::recursive_directory_iterator(root)) {
    if (entry.path().extension() == ".xlsb" && UnderWinDirectory(entry.path().lexically_relative(root))) {
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
