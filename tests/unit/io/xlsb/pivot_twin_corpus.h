// The Excel-saved pivot corpus the XLSB pivot codecs are checked against:
// every `.xlsb` of tests/fixtures/excel/xlsb_pivot_twins plus the two
// Windows-saved grouping books, and the pivot parts inside each.

#ifndef FORMULON_TESTS_UNIT_IO_XLSB_PIVOT_TWIN_CORPUS_H_
#define FORMULON_TESTS_UNIT_IO_XLSB_PIVOT_TWIN_CORPUS_H_

#include <algorithm>
#include <cstdint>
#include <filesystem>
#include <string>
#include <utility>
#include <vector>

#include "io/zip_reader.h"
#include "writer_test_helpers.h"
#include "zip_repack_helpers.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace pivot_twins {

/// Paths of every Excel-saved pivot `.xlsb` in the corpus, sorted.
inline std::vector<std::string> CorpusPaths() {
  const std::string root = std::string(FORMULON_FIXTURES_DIR) + "/excel/";
  std::vector<std::string> paths;
  for (const auto& entry : std::filesystem::directory_iterator(root + "xlsb_pivot_twins")) {
    if (entry.path().extension() == ".xlsb") {
      paths.push_back(entry.path().string());
    }
  }
  paths.push_back(root + "win/pivot_days_end/pivot_days_end.xlsb");
  paths.push_back(root + "win/pivot_week_1900/pivot_week_1900.xlsb");
  std::sort(paths.begin(), paths.end());
  return paths;
}

/// (part path, bytes) of every entry of `package_path` whose path starts with
/// `prefix` and ends in `.bin` (e.g. "xl/pivotTables/pivotTable").
inline std::vector<std::pair<std::string, std::vector<std::uint8_t>>> Parts(const std::string& package_path,
                                                                            const std::string& prefix) {
  const std::vector<std::uint8_t> package = ReadFixtureFile(package_path);
  ZipReader zip;
  std::vector<std::pair<std::string, std::vector<std::uint8_t>>> out;
  if (!zip.open(SpanOf(package))) {
    return out;
  }
  for (const std::string& name : zip.list_entries()) {
    if (name.compare(0, prefix.size(), prefix) == 0 && name.size() > 4U &&
        name.compare(name.size() - 4U, 4U, ".bin") == 0) {
      auto bytes = zip.read_entry(name);
      if (bytes) {
        out.emplace_back(name, std::move(bytes.value()));
      }
    }
  }
  return out;
}

}  // namespace pivot_twins
}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_TESTS_UNIT_IO_XLSB_PIVOT_TWIN_CORPUS_H_
