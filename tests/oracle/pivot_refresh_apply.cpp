// Engine half of the pivot-refresh oracle: loads a case's before.xlsx under
// the case's Excel profile, refreshes every pivot cache (or only loads, for a
// `trigger: open` case) and writes the result as .xlsx.
// tools/oracle/pivot_refresh_golden.py compares that file with the case's
// Excel golden.
//
// Usage: formulon_pivot_refresh_apply <before.xlsx> <out.xlsx> [--profile <id>] [--open]
// Exit status: 0 when every refresh succeeded, 3 when one returned an error
// (printed as `refresh-status <cache_id> <code> <message>`; the workbook is
// still written so the caller can check that nothing changed), 1 on usage,
// I/O or load failure.

#include <cstdint>
#include <cstdio>
#include <cstring>
#include <fstream>
#include <iterator>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/function_registry.h"
#include "excel_profile.h"
#include "io/ooxml_reader.h"
#include "pivot/pivot_cache.h"
#include "workbook.h"

namespace {

bool read_file(const char* path, std::vector<std::uint8_t>* out) {
  std::ifstream file(path, std::ios::binary);
  if (!file) {
    return false;
  }
  out->assign(std::istreambuf_iterator<char>(file), std::istreambuf_iterator<char>());
  return true;
}

bool write_file(const char* path, const std::vector<std::uint8_t>& bytes) {
  std::ofstream out(path, std::ios::binary);
  out.write(reinterpret_cast<const char*>(bytes.data()), static_cast<std::streamsize>(bytes.size()));
  return static_cast<bool>(out);
}

int usage(const char* argv0) {
  std::fprintf(stderr, "usage: %s <before.xlsx> <out.xlsx> [--profile <id>] [--open]\n", argv0);
  return 1;
}

}  // namespace

int main(int argc, char** argv) {
  if (argc < 3) {
    return usage(argv[0]);
  }
  bool open_only = false;
  formulon::ExcelProfile profile = formulon::mac_365_ja_jp_profile();
  for (int i = 3; i < argc; ++i) {
    if (std::strcmp(argv[i], "--open") == 0) {
      open_only = true;
    } else if (std::strcmp(argv[i], "--profile") == 0 && i + 1 < argc &&
               formulon::parse_excel_profile_id(argv[i + 1], &profile)) {
      ++i;
    } else {
      return usage(argv[0]);
    }
  }

  std::vector<std::uint8_t> input;
  if (!read_file(argv[1], &input)) {
    std::fprintf(stderr, "cannot read %s\n", argv[1]);
    return 1;
  }
  auto read = formulon::io::read_ooxml(formulon::io::ByteSpan{input.data(), input.size()});
  if (!read) {
    std::fprintf(stderr, "load failed: %s: %s\n", argv[1], read.error().message.c_str());
    return 1;
  }
  formulon::Workbook wb = std::move(read.value().workbook);
  wb.set_excel_profile(profile);

  int exit_code = 0;
  if (!open_only) {
    std::vector<std::uint32_t> cache_ids;
    for (const auto& cache : wb.pivot_caches()) {
      cache_ids.push_back(cache->cache_id());
    }
    for (const std::uint32_t cache_id : cache_ids) {
      auto refreshed = wb.refresh_pivot_cache(cache_id, formulon::eval::default_registry());
      if (!refreshed) {
        std::printf("refresh-status %u %d %s\n", static_cast<unsigned>(cache_id),
                    static_cast<int>(refreshed.error().code), refreshed.error().message.c_str());
        exit_code = 3;
      }
    }
  }

  auto bytes = wb.save();
  if (!bytes || !write_file(argv[2], bytes.value())) {
    std::fprintf(stderr, "save failed: %s\n", argv[2]);
    return 1;
  }
  return exit_code;
}
