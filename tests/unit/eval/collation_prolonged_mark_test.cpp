//
// Mac Excel text comparison around the prolonged sound mark (U+30FC and its
// halfwidth form U+FF70) on mac-365-ja_JP and mac-365-en_US. The tables are the
// `=` and `<` results Excel returns for every ordered pair of the probe strings
// below; cell [i][j] is `S_i op S_j`, '1' for TRUE.

#include <array>
#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "eval/collation.h"
#include "eval/eval_profile_scope.h"
#include "excel_locale.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "utils/text_ops.h"

namespace formulon {
namespace eval {
namespace {

constexpr std::size_t kCount = 37;

/// Code points of one probe string; unused slots are zero.
struct ProbeString {
  std::array<std::uint32_t, 5> cps;
};

constexpr ProbeString kStrings[kCount] = {
    {{0xFF83, 0xFF9E, 0xFF70, 0xFF80}},          // 0 ﾃﾞｰﾀ
    {{0x30C7, 0xFF70, 0x30BF}},                  // 1 デｰタ
    {{0xFF83, 0xFF9E, 0x30FC, 0xFF80}},          // 2 ﾃﾞーﾀ
    {{0x30C7, 0x30FC, 0x30BF}},                  // 3 データ
    {{0xFF83, 0xFF70, 0xFF80}},                  // 4 ﾃｰﾀ
    {{0x30C6, 0xFF70, 0x30BF}},                  // 5 テｰタ
    {{0xFF83, 0x30FC, 0xFF80}},                  // 6 ﾃーﾀ
    {{0x30C6, 0x30FC, 0x30BF}},                  // 7 テータ
    {{0xFF76, 0xFF70}},                          // 8 ｶｰ
    {{0x30AB, 0xFF70}},                          // 9 カｰ
    {{0xFF76, 0x30FC}},                          // 10 ｶー
    {{0x30AB, 0x30FC}},                          // 11 カー
    {{0xFF76, 0xFF9E, 0xFF70}},                  // 12 ｶﾞｰ
    {{0x30AC, 0xFF70}},                          // 13 ガｰ
    {{0xFF76, 0xFF9E, 0x30FC}},                  // 14 ｶﾞー
    {{0x30AC, 0x30FC}},                          // 15 ガー
    {{0xFF8A, 0xFF9F, 0xFF70}},                  // 16 ﾊﾟｰ
    {{0x30D1, 0x30FC}},                          // 17 パー
    {{0xFF71, 0xFF70}},                          // 18 ｱｰ
    {{0x30A2, 0x30FC}},                          // 19 アー
    {{0xFF70}},                                  // 20 ｰ
    {{0x30FC}},                                  // 21 ー
    {{0xFF83, 0xFF70}},                          // 22 ﾃｰ
    {{0x30C6, 0x30FC}},                          // 23 テー
    {{0xFF83, 0xFF9E}},                          // 24 ﾃﾞ
    {{0x30C7}},                                  // 25 デ
    {{0xFF83, 0xFF9E, 0xFF70}},                  // 26 ﾃﾞｰ
    {{0x30C7, 0x30FC}},                          // 27 デー
    {{0xFF70, 0xFF80}},                          // 28 ｰﾀ
    {{0x30FC, 0xFF80}},                          // 29 ーﾀ
    {{0xFF76, 0xFF70, 0xFF76}},                  // 30 ｶｰｶ
    {{0x30AB, 0x30FC, 0x30AB}},                  // 31 カーカ
    {{0xFF83, 0xFF9E, 0xFF70, 0xFF80, 0x0061}},  // 32 ﾃﾞｰﾀa
    {{0x30C7, 0x30FC, 0x30BF, 0x0061}},          // 33 データa
    {{0x304B, 0x30FC}},                          // 34 かー
    {{0x304B, 0xFF70}},                          // 35 かｰ
    {{0xFF83, 0xFF9E, 0x30FC, 0xFF80}},          // 36 ﾃﾞーﾀ
};
constexpr const char* kJaJpEqual[kCount] = {
    "1110000000000000000000000000000000001",  // 0
    "1110000000000000000000000000000000001",  // 1
    "1110000000000000000000000000000000001",  // 2
    "0001000000000000000000000000000000000",  // 3
    "0000110000000000000000000000000000000",  // 4
    "0000110000000000000000000000000000000",  // 5
    "0000001100000000000000000000000000000",  // 6
    "0000001100000000000000000000000000000",  // 7
    "0000000011000000000000000000000000010",  // 8
    "0000000011000000000000000000000000010",  // 9
    "0000000000110000000000000000000000100",  // 10
    "0000000000110000000000000000000000100",  // 11
    "0000000000001110000000000000000000000",  // 12
    "0000000000001110000000000000000000000",  // 13
    "0000000000001110000000000000000000000",  // 14
    "0000000000000001000000000000000000000",  // 15
    "0000000000000000100000000000000000000",  // 16
    "0000000000000000010000000000000000000",  // 17
    "0000000000000000001000000000000000000",  // 18
    "0000000000000000000100000000000000000",  // 19
    "0000000000000000000011000000000000000",  // 20
    "0000000000000000000011000000000000000",  // 21
    "0000000000000000000000100000000000000",  // 22
    "0000000000000000000000010000000000000",  // 23
    "0000000000000000000000001100000000000",  // 24
    "0000000000000000000000001100000000000",  // 25
    "0000000000000000000000000010000000000",  // 26
    "0000000000000000000000000001000000000",  // 27
    "0000000000000000000000000000110000000",  // 28
    "0000000000000000000000000000110000000",  // 29
    "0000000000000000000000000000001000000",  // 30
    "0000000000000000000000000000000100000",  // 31
    "0000000000000000000000000000000010000",  // 32
    "0000000000000000000000000000000001000",  // 33
    "0000000000110000000000000000000000100",  // 34
    "0000000011000000000000000000000000010",  // 35
    "1110000000000000000000000000000000001",  // 36
};
constexpr const char* kJaJpLess[kCount] = {
    "0001001100000000110000010001000011000",  // 0
    "0001001100000000110000010001000011000",  // 1
    "0001001100000000110000010001000011000",  // 2
    "0000000000000000110000000000000001000",  // 3
    "1111001100000000110000010001000011001",  // 4
    "1111001100000000110000010001000011001",  // 5
    "0001000000000000110000000000000001000",  // 6
    "0001000000000000110000000000000001000",  // 7
    "1111111100111111110000111111001111101",  // 8
    "1111111100111111110000111111001111101",  // 9
    "1111111100000001110000111111000111001",  // 10
    "1111111100000001110000111111000111001",  // 11
    "1111111100110001110000111111001111101",  // 12
    "1111111100110001110000111111001111101",  // 13
    "1111111100110001110000111111001111101",  // 14
    "1111111100000000110000111111000111001",  // 15
    "0000000000000000010000000000000000000",  // 16
    "0000000000000000000000000000000000000",  // 17
    "1111111111111111110100111111001111111",  // 18
    "1111111111111111110000111111001111111",  // 19
    "1111111111111111111100111111111111111",  // 20
    "1111111111111111111100111111111111111",  // 21
    "1111111100000000110000010011000011001",  // 22
    "0001001100000000110000000001000001000",  // 23
    "1111111100000000110000110011000011001",  // 24
    "1111111100000000110000110011000011001",  // 25
    "1111111100000000110000010001000011001",  // 26
    "0001001100000000110000000000000001000",  // 27
    "1111111111111111111100111111001111111",  // 28
    "1111111111111111111100111111001111111",  // 29
    "1111111100110001110000111111000111101",  // 30
    "1111111100000000110000111111000011001",  // 31
    "0001001100000000110000010001000001000",  // 32
    "0000000000000000110000000000000000000",  // 33
    "1111111100000001110000111111000111001",  // 34
    "1111111100111111110000111111001111101",  // 35
    "0001001100000000110000010001000011000",  // 36
};
constexpr const char* kEnUsEqual[kCount] = {
    "1111000000000000000000000000000000001",  // 0
    "1111000000000000000000000000000000001",  // 1
    "1111000000000000000000000000000000001",  // 2
    "1111000000000000000000000000000000001",  // 3
    "0000111100000000000000000000000000000",  // 4
    "0000111100000000000000000000000000000",  // 5
    "0000111100000000000000000000000000000",  // 6
    "0000111100000000000000000000000000000",  // 7
    "0000000011110000000000000000000000110",  // 8
    "0000000011110000000000000000000000110",  // 9
    "0000000011110000000000000000000000110",  // 10
    "0000000011110000000000000000000000110",  // 11
    "0000000000001111000000000000000000000",  // 12
    "0000000000001111000000000000000000000",  // 13
    "0000000000001111000000000000000000000",  // 14
    "0000000000001111000000000000000000000",  // 15
    "0000000000000000110000000000000000000",  // 16
    "0000000000000000110000000000000000000",  // 17
    "0000000000000000001100000000000000000",  // 18
    "0000000000000000001100000000000000000",  // 19
    "0000000000000000000011000000000000000",  // 20
    "0000000000000000000011000000000000000",  // 21
    "0000000000000000000000110000000000000",  // 22
    "0000000000000000000000110000000000000",  // 23
    "0000000000000000000000001100000000000",  // 24
    "0000000000000000000000001100000000000",  // 25
    "0000000000000000000000000011000000000",  // 26
    "0000000000000000000000000011000000000",  // 27
    "0000000000000000000000000000110000000",  // 28
    "0000000000000000000000000000110000000",  // 29
    "0000000000000000000000000000001100000",  // 30
    "0000000000000000000000000000001100000",  // 31
    "0000000000000000000000000000000011000",  // 32
    "0000000000000000000000000000000011000",  // 33
    "0000000011110000000000000000000000110",  // 34
    "0000000011110000000000000000000000110",  // 35
    "1111000000000000000000000000000000001",  // 36
};

std::string Utf8(std::size_t i) {
  std::string out;
  for (const std::uint32_t cp : kStrings[i].cps) {
    if (cp == 0) {
      break;
    }
    out += encode_utf8_codepoint(cp);
  }
  return out;
}

bool Bit(const char* const* table, std::size_t i, std::size_t j) {
  return table[i][j] == '1';
}

std::string Join(const std::vector<std::string>& lines) {
  std::string out;
  for (const std::string& line : lines) {
    out += line;
    out += '\n';
  }
  return out;
}

enum class Relation { kEqual, kLess };

// Every ordered pair whose engine result differs from the measured table.
std::vector<std::string> Mismatches(ExcelLocale locale, const char* const* table, Relation relation) {
  const ExcelProfile profile{ExcelHost::kMac365, locale};
  ScopedEvalProfile scope(profile);
  std::vector<std::string> misses;
  for (std::size_t i = 0; i < kCount; ++i) {
    for (std::size_t j = 0; j < kCount; ++j) {
      const int order = collate(Utf8(i), Utf8(j), profile, CollationStrength::kCompare);
      const bool got = relation == Relation::kEqual ? order == 0 : order < 0;
      if (got != Bit(table, i, j)) {
        misses.push_back(std::to_string(i) + " " + Utf8(i) + " | " + std::to_string(j) + " " + Utf8(j) +
                         " want=" + table[i][j] + " collate=" + std::to_string(order));
      }
    }
  }
  return misses;
}

TEST(CollationProlongedMark, JaJpEqualityMatchesExcel) {
  const std::vector<std::string> misses = Mismatches(ExcelLocale::kJaJP, kJaJpEqual, Relation::kEqual);
  EXPECT_TRUE(misses.empty()) << misses.size() << " pairs differ:\n" << Join(misses);
}

TEST(CollationProlongedMark, JaJpOrderMatchesExcel) {
  const std::vector<std::string> misses = Mismatches(ExcelLocale::kJaJP, kJaJpLess, Relation::kLess);
  EXPECT_TRUE(misses.empty()) << misses.size() << " pairs differ:\n" << Join(misses);
}

TEST(CollationProlongedMark, EnUsEqualityMatchesExcel) {
  const std::vector<std::string> misses = Mismatches(ExcelLocale::kEnUS, kEnUsEqual, Relation::kEqual);
  EXPECT_TRUE(misses.empty()) << misses.size() << " pairs differ:\n" << Join(misses);
}

// A kana followed by the combining voiced (U+3099) or semi-voiced (U+309A) mark
// equals the precomposed kana under `=` in both UIs.
TEST(CollationProlongedMark, CombiningVoicingMarksCompose) {
  const std::string ga = encode_utf8_codepoint(0x30AC);                                         // ガ
  const std::string ka_voiced = encode_utf8_codepoint(0x30AB) + encode_utf8_codepoint(0x3099);  // カ + U+3099
  const std::string pa = encode_utf8_codepoint(0x30D1);                                         // パ
  const std::string ha_semi = encode_utf8_codepoint(0x30CF) + encode_utf8_codepoint(0x309A);    // ハ + U+309A
  for (const ExcelLocale locale : {ExcelLocale::kJaJP, ExcelLocale::kEnUS}) {
    const ExcelProfile profile{ExcelHost::kMac365, locale};
    ScopedEvalProfile scope(profile);
    EXPECT_EQ(collate(ga, ka_voiced, profile, CollationStrength::kCompare), 0) << static_cast<int>(locale);
    EXPECT_EQ(collate(pa, ha_semi, profile, CollationStrength::kCompare), 0) << static_cast<int>(locale);
  }
}

// Guards the tables themselves: `=` is an equivalence, `<` a strict order
// consistent with it, exactly one of `a<b`, `a=b`, `b<a` holds, and ja-JP
// equality refines en-US equality.
TEST(CollationProlongedMark, MeasuredTablesAreConsistent) {
  for (std::size_t i = 0; i < kCount; ++i) {
    ASSERT_EQ(std::string(kJaJpEqual[i]).size(), kCount) << "row " << i;
    ASSERT_EQ(std::string(kJaJpLess[i]).size(), kCount) << "row " << i;
    ASSERT_EQ(std::string(kEnUsEqual[i]).size(), kCount) << "row " << i;
  }
  for (const char* const* eq : {kJaJpEqual, kEnUsEqual}) {
    for (std::size_t i = 0; i < kCount; ++i) {
      EXPECT_TRUE(Bit(eq, i, i)) << i;
      for (std::size_t j = 0; j < kCount; ++j) {
        EXPECT_EQ(Bit(eq, i, j), Bit(eq, j, i)) << i << "," << j;
        for (std::size_t k = 0; k < kCount; ++k) {
          if (Bit(eq, i, j) && Bit(eq, j, k)) {
            EXPECT_TRUE(Bit(eq, i, k)) << i << "," << j << "," << k;
          }
        }
      }
    }
  }
  for (std::size_t i = 0; i < kCount; ++i) {
    for (std::size_t j = 0; j < kCount; ++j) {
      const int held = static_cast<int>(Bit(kJaJpEqual, i, j)) + static_cast<int>(Bit(kJaJpLess, i, j)) +
                       static_cast<int>(Bit(kJaJpLess, j, i));
      EXPECT_EQ(held, 1) << i << "," << j;
      if (Bit(kJaJpEqual, i, j)) {
        EXPECT_TRUE(Bit(kEnUsEqual, i, j)) << i << "," << j;
      }
      for (std::size_t k = 0; k < kCount; ++k) {
        if (Bit(kJaJpLess, i, j) && Bit(kJaJpLess, j, k)) {
          EXPECT_TRUE(Bit(kJaJpLess, i, k)) << i << "," << j << "," << k;
        }
        if (Bit(kJaJpEqual, i, k)) {
          EXPECT_EQ(Bit(kJaJpLess, i, j), Bit(kJaJpLess, k, j)) << i << "," << k << " vs " << j;
        }
      }
    }
  }
  // Strings 2 and 36 are the same code points, so their rows must agree.
  EXPECT_EQ(Utf8(2), Utf8(36));
  EXPECT_EQ(std::string(kJaJpLess[2]), std::string(kJaJpLess[36]));
}

}  // namespace
}  // namespace eval
}  // namespace formulon
