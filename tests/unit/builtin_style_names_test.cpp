// Built-in cell style names checked against the table extracted from Mac
// Excel 365 (tests/fixtures/cell_style_save/builtin_style_names.json).

#include "builtin_style_names.h"

#include <cstddef>
#include <cstdint>
#include <fstream>
#include <iterator>
#include <map>
#include <string>
#include <vector>

#include "excel_profile.h"
#include "gtest/gtest.h"

namespace formulon {
namespace {

constexpr std::size_t kLocaleCount = 14U;

const char* const kLocaleKeys[kLocaleCount] = {"ja_JP", "en_US", "de_DE", "fr_FR", "zh_CN", "ko_KR", "th_TH",
                                               "es_ES", "es_MX", "pt_BR", "ru_RU", "zh_TW", "it_IT", "nl_NL"};

using Table = std::map<std::string, std::vector<std::string>>;

// Reads the fixed shape: `"<locale>": { ... "names": [ "...", ... ] }`. Strings carry no escapes.
Table LoadTable() {
  std::ifstream in(std::string(FORMULON_FIXTURES_DIR) + "/cell_style_save/builtin_style_names.json", std::ios::binary);
  const std::string text((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());
  Table table;
  for (const char* key : kLocaleKeys) {
    const std::size_t locale_at = text.find(std::string("\"") + key + "\": {");
    if (locale_at == std::string::npos) {
      continue;
    }
    std::size_t cur = text.find('[', text.find("\"names\"", locale_at));
    std::vector<std::string>& names = table[key];
    for (;;) {
      cur = text.find_first_of("\"]", cur + 1U);
      if (cur == std::string::npos || text[cur] == ']') {
        break;
      }
      const std::size_t close = text.find('"', cur + 1U);
      names.push_back(text.substr(cur + 1U, close - cur - 1U));
      cur = close;
    }
  }
  return table;
}

TEST(BuiltinStyleNames, EveryLocaleAndIdMatchesTheExcelTable) {
  const Table table = LoadTable();
  ASSERT_EQ(table.size(), kLocaleCount);
  std::size_t compared = 0;
  for (std::size_t li = 0; li < kLocaleCount; ++li) {
    const auto it = table.find(kLocaleKeys[li]);
    ASSERT_NE(it, table.end()) << kLocaleKeys[li];
    ASSERT_EQ(it->second.size(), kBuiltinStyleCount) << kLocaleKeys[li];
    for (std::uint32_t id = 0; id < kBuiltinStyleCount; ++id) {
      const std::string got = builtin_style_name(id, 0U, static_cast<ExcelLocale>(li));
      const std::string want = id == 1U || id == 2U ? it->second[id] + "1" : it->second[id];
      EXPECT_EQ(got, want) << kLocaleKeys[li] << " id " << id;
      ++compared;
    }
  }
  EXPECT_EQ(compared, 756U);
}

TEST(BuiltinStyleNames, LevelStemsTakeTheOneBasedLevel) {
  const Table table = LoadTable();
  for (std::size_t li = 0; li < kLocaleCount; ++li) {
    const auto& names = table.at(kLocaleKeys[li]);
    for (std::uint32_t level = 0; level < 7U; ++level) {
      const std::string digit(1, static_cast<char>('1' + level));
      EXPECT_EQ(builtin_style_name(1U, level, static_cast<ExcelLocale>(li)), names[1] + digit) << kLocaleKeys[li];
      EXPECT_EQ(builtin_style_name(2U, level, static_cast<ExcelLocale>(li)), names[2] + digit) << kLocaleKeys[li];
    }
    // Other ids ignore the level.
    EXPECT_EQ(builtin_style_name(0U, 5U, static_cast<ExcelLocale>(li)), names[0]);
  }
}

TEST(BuiltinStyleNames, OutOfRangeIdIsEmpty) {
  EXPECT_TRUE(builtin_style_name(kBuiltinStyleCount, 0U, ExcelLocale::kEnUS).empty());
  EXPECT_TRUE(builtin_style_name(0xFFFFFFFFU, 0U, ExcelLocale::kJaJP).empty());
}

}  // namespace
}  // namespace formulon
