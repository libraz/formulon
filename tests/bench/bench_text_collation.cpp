//
// Microbenchmark: locale-aware text collation over 100k rows.
//
// Scenarios, each for mac-365-ja_JP and mac-365-zh_CN:
//   * SORT shape: build one dynamic-strength key per row, then sort row
//     indices by key.
//   * COUNTIF shape: build the criterion key once, then compare every row's
//     compare-strength key against it.
//   * Classify shape: dense dynamic-strength class ids for every row.
// Rows mix kana, ideographs, full-width and ASCII text so both the ASCII fast
// path and the general key path are exercised.

#define ANKERL_NANOBENCH_IMPLEMENT
#include <nanobench.h>

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <fstream>
#include <numeric>
#include <random>
#include <string>
#include <vector>

#include "eval/collation.h"
#include "eval/dynamic_array/common.h"
#include "excel_profile.h"
#include "value.h"

namespace {

constexpr std::size_t kRows = 100000;

std::vector<std::string> BuildRows() {
  static const char* kPieces[] = {"かー", "カ",   "ｶﾞ", "漢字", "東京", "abc", "Abc", "Ａ",
                                  "123",  "１２", "ß",  "さゝ", "ー",   "x-y", "東",  "ﾊﾟ"};
  std::mt19937 rng(42u);
  std::vector<std::string> rows;
  rows.reserve(kRows);
  for (std::size_t i = 0; i < kRows; ++i) {
    std::string s;
    const std::size_t pieces = 1 + rng() % 4;
    for (std::size_t k = 0; k < pieces; ++k) {
      s += kPieces[rng() % (sizeof(kPieces) / sizeof(kPieces[0]))];
    }
    rows.push_back(std::move(s));
  }
  return rows;
}

void SortScenario(ankerl::nanobench::Bench& bench, const char* name, const std::vector<std::string>& rows,
                  formulon::ExcelProfile profile) {
  bench.run(name, [&]() {
    std::vector<std::string> keys;
    keys.reserve(rows.size());
    for (const std::string& r : rows) {
      keys.push_back(formulon::eval::collation_key(r, profile, formulon::eval::CollationStrength::kDynamic));
    }
    std::vector<std::uint32_t> order(rows.size());
    std::iota(order.begin(), order.end(), 0u);
    std::stable_sort(order.begin(), order.end(), [&](std::uint32_t a, std::uint32_t b) {
      return formulon::eval::collate_key_compare(keys[a], keys[b]) < 0;
    });
    ankerl::nanobench::doNotOptimizeAway(order.front());
  });
}

void CountifScenario(ankerl::nanobench::Bench& bench, const char* name, const std::vector<std::string>& rows,
                     formulon::ExcelProfile profile) {
  bench.run(name, [&]() {
    const std::string criterion =
        formulon::eval::collation_key("かあ", profile, formulon::eval::CollationStrength::kCompare);
    std::size_t hits = 0;
    for (const std::string& r : rows) {
      const std::string key = formulon::eval::collation_key(r, profile, formulon::eval::CollationStrength::kCompare);
      hits += formulon::eval::collate_key_compare(key, criterion) == 0;
    }
    ankerl::nanobench::doNotOptimizeAway(hits);
  });
}

void ClassifyScenario(ankerl::nanobench::Bench& bench, const char* name, const std::vector<std::string>& rows,
                      formulon::ExcelProfile profile) {
  std::vector<formulon::Value> values;
  values.reserve(rows.size());
  for (const std::string& r : rows) {
    values.push_back(formulon::Value::text(r));
  }
  std::vector<std::uint32_t> ids(values.size());
  bench.run(name, [&]() {
    const std::uint32_t classes = formulon::eval::dynamic_array::collation_classes(
        values.data(), values.size(), profile, formulon::eval::CollationStrength::kDynamic, ids.data());
    ankerl::nanobench::doNotOptimizeAway(classes);
  });
}

}  // namespace

int main(int argc, char* argv[]) {
  std::string json_path;
  for (int i = 1; i < argc; ++i) {
    if (std::strcmp(argv[i], "--json") == 0 && i + 1 < argc) {
      json_path = argv[i + 1];
      ++i;
    }
  }

  const std::vector<std::string> rows = BuildRows();
  const formulon::ExcelProfile ja{formulon::ExcelHost::kMac365, formulon::ExcelLocale::kJaJP};
  const formulon::ExcelProfile zh{formulon::ExcelHost::kMac365, formulon::ExcelLocale::kZhCN};

  // Decode the lazily built tables before timing.
  (void)formulon::eval::collation_key("漢", ja, formulon::eval::CollationStrength::kCompare);
  (void)formulon::eval::collation_key("漢", zh, formulon::eval::CollationStrength::kCompare);

  ankerl::nanobench::Bench bench;
  bench.title("text_collation").unit("100k rows").warmup(1).minEpochIterations(3);

  SortScenario(bench, "SORT 100k rows ja_JP", rows, ja);
  SortScenario(bench, "SORT 100k rows zh_CN", rows, zh);
  CountifScenario(bench, "COUNTIF 100k rows ja_JP", rows, ja);
  CountifScenario(bench, "COUNTIF 100k rows zh_CN", rows, zh);
  ClassifyScenario(bench, "CLASSIFY 100k rows ja_JP", rows, ja);
  ClassifyScenario(bench, "CLASSIFY 100k rows zh_CN", rows, zh);

  if (!json_path.empty()) {
    std::ofstream out(json_path);
    bench.render(ankerl::nanobench::templates::json(), out);
  }
  return 0;
}
