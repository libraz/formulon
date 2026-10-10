//
// Unit tests for the text collation keys. Expectations come from the Mac Excel
// measurements under tools/collation/data (FORMULON_COLLATION_DATA_DIR): the
// probe matrix, the BMP sort order, the function equality table and the
// trailing-ignorable table, plus the Windows host's existing byte model.

#include "eval/collation.h"

#include <algorithm>
#include <array>
#include <cstddef>
#include <cstdint>
#include <fstream>
#include <iterator>
#include <map>
#include <random>
#include <set>
#include <string>
#include <string_view>
#include <vector>

#include "eval/collation_table.h"
#include "eval/dbcs_table.h"
#include "eval/eval_profile_scope.h"
#include "eval/jp_fold.h"
#include "excel_locale.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "utils/text_ops.h"

#ifndef FORMULON_COLLATION_DATA_DIR
#error "FORMULON_COLLATION_DATA_DIR must name tools/collation/data"
#endif

namespace formulon {
namespace eval {
namespace {

constexpr const char* kTags[14] = {"ja_JP", "en_US", "de_DE", "fr_FR", "zh_CN", "ko_KR", "th_TH",
                                   "es_ES", "es_MX", "pt_BR", "ru_RU", "zh_TW", "it_IT", "nl_NL"};

using Rows = std::vector<std::vector<std::string>>;

std::string tag_of(ExcelLocale locale) {
  return kTags[static_cast<std::size_t>(locale)];
}

bool is_zh(ExcelLocale locale) {
  return locale == ExcelLocale::kZhCN || locale == ExcelLocale::kZhTW;
}

ExcelProfile mac(ExcelLocale locale) {
  return ExcelProfile{ExcelHost::kMac365, locale};
}

ExcelProfile win(ExcelLocale locale) {
  return ExcelProfile{ExcelHost::kWin365, locale};
}

// Rows of a TSV under the data directory, header skipped, empty fields kept.
Rows read_tsv(const std::string& relative) {
  std::ifstream in(std::string(FORMULON_COLLATION_DATA_DIR) + "/" + relative);
  Rows rows;
  std::string line;
  bool header = true;
  while (std::getline(in, line)) {
    if (header) {
      header = false;
      continue;
    }
    if (line.empty()) {
      continue;
    }
    std::vector<std::string> fields;
    std::size_t start = 0;
    for (std::size_t tab = line.find('\t'); tab != std::string::npos; tab = line.find('\t', start)) {
      fields.push_back(line.substr(start, tab - start));
      start = tab + 1;
    }
    fields.push_back(line.substr(start));
    rows.push_back(std::move(fields));
  }
  return rows;
}

std::uint32_t parse_cp(std::string_view token) {
  return static_cast<std::uint32_t>(std::stoul(std::string(token.substr(2)), nullptr, 16));
}

// "U+0061 U+0301" -> UTF-8.
std::string decode_field(std::string_view field) {
  std::string out;
  std::size_t i = 0;
  while (i < field.size()) {
    const std::size_t space = field.find(' ', i);
    const std::size_t end = space == std::string_view::npos ? field.size() : space;
    out += encode_utf8_codepoint(parse_cp(field.substr(i, end - i)));
    i = end + 1;
  }
  return out;
}

std::string u8(std::uint32_t cp) {
  return encode_utf8_codepoint(cp);
}

char sign_char(int c) {
  return c < 0 ? '-' : c == 0 ? '0' : '+';
}

int key_sign(std::string_view a, std::string_view b, ExcelProfile p, CollationStrength s) {
  return collate_key_compare(collation_key(a, p, s), collation_key(b, p, s));
}

// Applies `check` to every Mac locale as its own test case.
#define FORMULON_PER_LOCALE(Suite, check) \
  TEST(Suite, ja_JP) {                    \
    check(ExcelLocale::kJaJP);            \
  }                                       \
  TEST(Suite, en_US) {                    \
    check(ExcelLocale::kEnUS);            \
  }                                       \
  TEST(Suite, de_DE) {                    \
    check(ExcelLocale::kDeDE);            \
  }                                       \
  TEST(Suite, fr_FR) {                    \
    check(ExcelLocale::kFrFR);            \
  }                                       \
  TEST(Suite, zh_CN) {                    \
    check(ExcelLocale::kZhCN);            \
  }                                       \
  TEST(Suite, ko_KR) {                    \
    check(ExcelLocale::kKoKR);            \
  }                                       \
  TEST(Suite, th_TH) {                    \
    check(ExcelLocale::kThTH);            \
  }                                       \
  TEST(Suite, es_ES) {                    \
    check(ExcelLocale::kEsES);            \
  }                                       \
  TEST(Suite, es_MX) {                    \
    check(ExcelLocale::kEsMX);            \
  }                                       \
  TEST(Suite, pt_BR) {                    \
    check(ExcelLocale::kPtBR);            \
  }                                       \
  TEST(Suite, ru_RU) {                    \
    check(ExcelLocale::kRuRU);            \
  }                                       \
  TEST(Suite, zh_TW) {                    \
    check(ExcelLocale::kZhTW);            \
  }                                       \
  TEST(Suite, it_IT) {                    \
    check(ExcelLocale::kItIT);            \
  }                                       \
  TEST(Suite, nl_NL) {                    \
    check(ExcelLocale::kNlNL);            \
  }

// ---------------------------------------------------------------------------
// Probe matrix: every cell matches Excel except exactly the excused pairs.
// ---------------------------------------------------------------------------

void check_matrix(ExcelLocale locale) {
  const ExcelProfile profile = mac(locale);
  ScopedEvalProfile scope(profile);
  const std::string tag = tag_of(locale);
  const Rows probes = read_tsv("fixtures/probe_strings.tsv");
  const Rows matrix = read_tsv("fixtures/matrix_" + tag + ".tsv");
  ASSERT_FALSE(probes.empty());
  ASSERT_EQ(matrix.size(), probes.size());
  std::vector<std::string> strings;
  for (const auto& row : probes) {
    strings.push_back(decode_field(row[1]));
  }
  const std::size_t n = strings.size();
  std::set<std::string> misses;
  const struct {
    CollationStrength strength;
    const char* name;
    std::size_t column;
  } kModes[] = {{CollationStrength::kCompare, "compare", 1}, {CollationStrength::kDynamic, "dynamic", 2}};
  for (const auto& mode : kModes) {
    std::vector<std::string> keys;
    for (const std::string& s : strings) {
      keys.push_back(collation_key(s, profile, mode.strength));
    }
    for (std::size_t i = 0; i < n; ++i) {
      ASSERT_EQ(matrix[i][mode.column].size(), n) << "row " << i;
      for (std::size_t j = 0; j < n; ++j) {
        const char got = sign_char(collate_key_compare(keys[i], keys[j]));
        ASSERT_EQ(sign_char(collate(strings[i], strings[j], profile, mode.strength)), got)
            << "collate disagrees with its keys: " << probes[i][1] << " | " << probes[j][1];
        const char want = matrix[i][mode.column][j];
        if (want != got) {
          misses.insert(probes[i][1] + "\t" + probes[j][1] + "\t" + mode.name + "\t" + want + "\t" + got);
        }
      }
    }
  }
  std::set<std::string> excused;
  for (const auto& row : read_tsv("fixtures/excused_" + tag + ".tsv")) {
    ASSERT_EQ(row.size(), 5u);
    excused.insert(row[0] + "\t" + row[1] + "\t" + row[2] + "\t" + row[3] + "\t" + row[4]);
  }
  ASSERT_FALSE(excused.empty());
  std::vector<std::string> unexpected;
  std::vector<std::string> missing;
  std::set_difference(misses.begin(), misses.end(), excused.begin(), excused.end(), std::back_inserter(unexpected));
  std::set_difference(excused.begin(), excused.end(), misses.begin(), misses.end(), std::back_inserter(missing));
  for (const std::string& m : unexpected) {
    ADD_FAILURE() << "unexcused mismatch: " << m;
  }
  for (const std::string& m : missing) {
    ADD_FAILURE() << "excused pair not reproduced: " << m;
  }
}

FORMULON_PER_LOCALE(CollationMatrix, check_matrix)

// ---------------------------------------------------------------------------
// BMP sort order round trip.
// ---------------------------------------------------------------------------

constexpr std::uint32_t kUroFirst = 0x4E00;
constexpr std::uint32_t kUroLast = 0x9FFF;

bool in_exact_range(std::uint32_t cp) {
  for (const auto& r : collation_detail::kCollationExactRanges) {
    if (cp >= r.first && cp <= r.last) {
      return true;
    }
  }
  return false;
}

struct Measured {
  std::uint32_t cp;
  int cls;  // 0 full, 1 mid, 2 word, 3 sec, 4 base
  bool expanding;
  std::array<int, 4> w;  // P, S, D, C ranks
};

int class_index(const std::string& name) {
  static const char* kNames[] = {"full", "mid", "word", "sec", "base"};
  for (int i = 0; i < 5; ++i) {
    if (name == kNames[i]) {
      return i;
    }
  }
  return -1;
}

// Measured ranks per class, as the generator derives them from the `rel` column.
std::vector<Measured> load_order(const std::string& tag) {
  std::vector<Measured> out;
  std::map<int, std::array<int, 4>> counters;
  for (const auto& row : read_tsv("order_" + tag + ".tsv")) {
    Measured m{parse_cp(row[0]), class_index(row[1]), row[2] == "X", {0, 0, 0, 0}};
    if (!m.expanding && m.cls >= 2) {
      std::array<int, 4>& c = counters[m.cls];
      const std::string& rel = row[2];
      const std::size_t level = rel == "P" ? 0 : rel == "S" ? 1 : rel == "D" ? 2 : rel == "C" ? 3 : 4;
      if (level < 4) {
        ++c[level];
        for (std::size_t k = level + 1; k < 4; ++k) {
          c[k] = 0;
        }
      }
      m.w = c;
    }
    out.push_back(m);
  }
  return out;
}

// Class a lone character shows in its key: which parts carry weights.
int key_class(const std::string& key) {
  std::vector<std::string> parts(1);
  for (const char ch : key) {
    if (ch == '\x01') {
      parts.emplace_back();
    } else {
      parts.back().push_back(ch);
    }
  }
  if (!parts[0].empty()) {
    return 4;
  }
  if (!parts[1].empty()) {
    return 3;
  }
  if (!parts[parts.size() - 2].empty()) {
    return 2;
  }
  return parts.back().empty() ? 0 : 1;
}

void check_order(ExcelLocale locale) {
  const ExcelProfile profile = mac(locale);
  ScopedEvalProfile scope(profile);
  const std::string tag = tag_of(locale);
  const std::vector<Measured> order = load_order(tag);
  ASSERT_GT(order.size(), 60000u);

  std::set<std::uint32_t> excluded;
  for (const Measured& m : order) {
    if (!in_exact_range(m.cp)) {
      excluded.insert(m.cp);
    }
  }
  if (locale == ExcelLocale::kZhCN) {
    // Ideographs outside GB 2312 level 1 are placed approximately, and so are the
    // compatibility ideographs measured equal to one of them.
    std::set<std::uint32_t> level1;
    for (unsigned row = 16; row <= 55; ++row) {
      for (unsigned cell = 1; cell <= kDbcsGridSize; ++cell) {
        level1.insert(lookup_dbcs_to_unicode(DbcsCodepage::kGb2312, static_cast<std::uint8_t>(row),
                                             static_cast<std::uint8_t>(cell)));
      }
    }
    std::set<std::array<int, 5>> approximated_weights;
    for (const Measured& m : order) {
      if (m.cp >= kUroFirst && m.cp <= kUroLast && m.cls == 4 && !m.expanding && level1.count(m.cp) == 0) {
        excluded.insert(m.cp);
        approximated_weights.insert({m.cls, m.w[0], m.w[1], m.w[2], m.w[3]});
      }
    }
    for (const Measured& m : order) {
      if (m.cp >= 0xF900 && m.cp <= 0xFAFF && !m.expanding &&
          approximated_weights.count({m.cls, m.w[0], m.w[1], m.w[2], m.w[3]}) != 0) {
        excluded.insert(m.cp);
      }
    }
  }
  // Out-of-range counts recorded when the tables were first generated.
  EXPECT_EQ(excluded.size(), locale == ExcelLocale::kZhCN ? 44301u : 26986u);

  // ko ideographs placed by their reading compare by syllable, not by the measured rank.
  std::set<std::uint32_t> by_reading;
  if (locale == ExcelLocale::kKoKR) {
    bool in_syllable = false;
    for (const Measured& m : order) {
      if (m.cp >= 0xAC00 && m.cp <= 0xD7A3) {
        in_syllable = true;
      } else if (m.cp >= kUroFirst && m.cp <= kUroLast && in_syllable) {
        by_reading.insert(m.cp);
      } else if (!(m.cp >= 0xF900 && m.cp <= 0xFAFF)) {
        in_syllable = false;
      }
    }
  }

  for (const CollationStrength strength : {CollationStrength::kCompare, CollationStrength::kDynamic}) {
    std::map<int, std::pair<const Measured*, std::string>> prev;
    int failures = 0;
    for (const Measured& m : order) {
      if (m.expanding || excluded.count(m.cp) != 0) {
        continue;
      }
      std::string key = collation_key(u8(m.cp), profile, strength);
      const bool exempt = (m.cp >= 0xAC00 && m.cp <= 0xD7A3) || by_reading.count(m.cp) != 0;
      if (key_class(key) != m.cls && !exempt) {
        ADD_FAILURE() << tag << " U+" << std::hex << m.cp << " class " << key_class(key) << " want " << m.cls;
      }
      auto it = prev.find(m.cls);
      if (it != prev.end()) {
        const Measured& p = *it->second.first;
        std::size_t level = 0;
        while (level < 4 && p.w[level] == m.w[level]) {
          ++level;
        }
        char want = (level == 4 || (level == 3 && strength == CollationStrength::kDynamic)) ? '0' : '-';
        if (m.cls <= 1) {
          want = '0';
        }
        const char got = sign_char(collate_key_compare(it->second.second, key));
        if (got != want && !exempt) {
          ++failures;
          if (failures <= 10) {
            ADD_FAILURE() << tag << " U+" << std::hex << p.cp << " vs U+" << m.cp << " got " << got << " want " << want;
          }
        }
      }
      prev[m.cls] = {&m, std::move(key)};
    }
    EXPECT_EQ(failures, 0);
  }
}

FORMULON_PER_LOCALE(CollationOrder, check_order)

// ---------------------------------------------------------------------------
// Function equality table: each measured function column through its strength.
// ---------------------------------------------------------------------------

// Criteria pre-fold the measured cells imply: full-width digits fold everywhere,
// and on Mac ja_JP a half-width voicing mark also composes onto its base. A wider
// kana fold breaks iteration marks (`かゝ` matches `かか` in COUNTIF), and composing
// outside ja contradicts the prefix rule (`ガ` against a `ｶﾞ` cell is 0).
std::string criteria_prefold(std::string_view s, ExcelProfile profile) {
  std::string out;
  for (std::size_t i = 0; i < s.size();) {
    std::size_t n = 0;
    const std::uint32_t cp = decode_utf8_step(s, i, &n);
    if (cp >= 0xFF10 && cp <= 0xFF19) {
      out.push_back(static_cast<char>('0' + (cp - 0xFF10)));
    } else {
      out.append(s.substr(i, n));
    }
    i += n;
  }
  if (profile.host == ExcelHost::kMac365 && profile.locale == ExcelLocale::kJaJP) {
    return compose_jp_halfwidth_voicing(out);
  }
  return out;
}

void check_eq(ExcelLocale locale) {
  const ExcelProfile profile = mac(locale);
  ScopedEvalProfile scope(profile);
  const std::string tag = tag_of(locale);
  std::ifstream in(std::string(FORMULON_COLLATION_DATA_DIR) + "/fixtures/eq_" + tag + ".tsv");
  std::string header_line;
  std::getline(in, header_line);
  std::vector<std::string> header;
  std::size_t start = 0;
  for (std::size_t tab = header_line.find('\t'); tab != std::string::npos; tab = header_line.find('\t', start)) {
    header.push_back(header_line.substr(start, tab - start));
    start = tab + 1;
  }
  header.push_back(header_line.substr(start));
  const Rows rows = read_tsv("fixtures/eq_" + tag + ".tsv");
  ASSERT_FALSE(rows.empty());
  int checked = 0;
  for (const auto& row : rows) {
    ASSERT_EQ(row.size(), header.size());
    const std::string a = decode_field(row[0]);
    const std::string b = decode_field(row[1]);
    for (std::size_t col = 2; col < header.size(); ++col) {
      const std::string& fn = header[col];
      bool match = false;
      std::string hit;
      if (fn == "eq" || fn == "switch" || fn == "filter") {
        match = collate(a, b, profile, CollationStrength::kCompare) == 0;
        hit = "1";
      } else if (fn == "countif" || fn == "countifeq" || fn == "sumif") {
        match = collate(criteria_prefold(a, profile), criteria_prefold(b, profile), profile,
                        CollationStrength::kCompare) == 0;
        hit = fn == "sumif" ? "2" : "1";
      } else if (fn == "match0" || fn == "vlookupF" || fn == "hlookupF") {
        match = collate(a, b, profile, CollationStrength::kClassicExact) == 0;
        hit = "2";
      } else if (fn == "xlookup" || fn == "xmatch0") {
        match = collate(a, b, profile, CollationStrength::kDynamic) == 0;
        hit = "2";
      } else if (fn == "unique" || fn == "groupby") {
        match = collate(a, b, profile, CollationStrength::kDynamic) == 0;
        hit = "1";
      } else if (fn == "dcount") {
        match = collate_prefix_equal(criteria_prefold(b, profile), criteria_prefold(a, profile), profile);
        hit = "1";
      } else {
        continue;  // exact (byte equality) and xmatchbin (binary search side effect)
      }
      ++checked;
      EXPECT_EQ(match, row[col] == hit) << tag << " " << fn << " " << row[0] << " | " << row[1] << " = " << row[col];
    }
  }
  EXPECT_EQ(checked, static_cast<int>(rows.size()) * 14);
}

FORMULON_PER_LOCALE(CollationEq, check_eq)

// ---------------------------------------------------------------------------
// Contextual rules and trailing ignorables.
// ---------------------------------------------------------------------------

// Trailing-ignorable probes: every row matches Excel except the pairs whose two
// sides differ only by a combining sequence, a kana type, an expansion (where a
// trailing ignorable on one side is not ordered as Excel orders it), or, on zh,
// two empty strings each followed by an ignorable.
void check_tail(ExcelLocale locale) {
  const ExcelProfile profile = mac(locale);
  ScopedEvalProfile scope(profile);
  const std::string tag = tag_of(locale);
  const Rows rows = read_tsv("fixtures/tail_" + tag + ".tsv");
  ASSERT_EQ(rows.size(), 6144u);
  std::set<std::pair<std::string, std::string>> known;
  const auto both = [&known](const char* x, const char* y) {
    known.emplace(x, y);
    known.emplace(y, x);
  };
  both("U+00E9", "U+0065 U+0301");
  if (is_zh(locale)) {
    both("", "");
  } else {
    both("U+30A2", "U+3042");
    both("U+0073 U+0073", "U+00DF");
  }
  int known_misses = 0;
  for (const auto& row : rows) {
    ASSERT_EQ(row.size(), 6u);
    const std::string c = decode_field(row[1]);
    const std::string c2 = decode_field(row[2]);
    const std::string x = decode_field(row[3]);
    const std::string y = decode_field(row[4]);
    std::string a = x;
    std::string b = y;
    if (row[0] == "x|y+c") {
      b += c;
    } else if (row[0] == "x+c|y") {
      a += c;
    } else {
      a += c;
      b += c2;
    }
    const char got = sign_char(collate(a, b, profile, CollationStrength::kCompare));
    if (got == row[5][0]) {
      continue;
    }
    if (known.count({row[3], row[4]}) != 0) {
      ++known_misses;
      continue;
    }
    ADD_FAILURE() << tag << " " << row[0] << " c=" << row[1] << " x=" << row[3] << " y=" << row[4] << " got " << got
                  << " want " << row[5];
  }
  // The reference model misses the same number of rows within those pairs.
  EXPECT_EQ(known_misses, is_zh(locale) ? 64 : 82);
}

FORMULON_PER_LOCALE(CollationRules, check_tail)

int cmp(std::string_view a, std::string_view b, ExcelLocale locale,
        CollationStrength strength = CollationStrength::kCompare) {
  return collate(a, b, mac(locale), strength);
}

TEST(CollationRules, JaProlongedAndIterationMarks) {
  ScopedEvalProfile scope(mac(ExcelLocale::kJaJP));
  const ExcelLocale ja = ExcelLocale::kJaJP;
  EXPECT_EQ(cmp("かー", "かあ", ja), 0);
  EXPECT_EQ(cmp("かゝ", "かか", ja), 0);
  EXPECT_GT(cmp("がゝ", "がか", ja), 0);
  EXPECT_LT(cmp("かヽ", "かか", ja), 0);
  EXPECT_GT(cmp("aー", "a", ja), 0);
  EXPECT_GT(cmp("んー", "ん", ja), 0);
}

TEST(CollationRules, LatinExpansionsAndAccents) {
  ScopedEvalProfile scope(mac(ExcelLocale::kEnUS));
  EXPECT_EQ(cmp("straße", "strasse", ExcelLocale::kEnUS), 0);
  EXPECT_LT(cmp("straße", "strasze", ExcelLocale::kEnUS), 0);
  EXPECT_NE(cmp("ß", "ss", ExcelLocale::kEnUS, CollationStrength::kClassicExact), 0);
  const ExcelLocale fr = ExcelLocale::kFrFR;
  EXPECT_LT(cmp("cote", "coté", fr), 0);
  EXPECT_LT(cmp("coté", "côte", fr), 0);
  EXPECT_LT(cmp("côte", "côté", fr), 0);
  const ExcelLocale de = ExcelLocale::kDeDE;
  EXPECT_LT(cmp("ä", "ae", de), 0);
  EXPECT_GT(cmp("äz", "aea", de), 0);
  EXPECT_GT(cmp("ñ", "nz", ExcelLocale::kEsES), 0);
  EXPECT_LT(cmp("ñ", "nz", ExcelLocale::kEnUS), 0);
}

TEST(CollationRules, HangulDecomposesOutsideZh) {
  ScopedEvalProfile scope(mac(ExcelLocale::kKoKR));
  const std::string ga_jamo = "가" + u8(0x11A8);
  for (const ExcelLocale w : {ExcelLocale::kJaJP, ExcelLocale::kEnUS, ExcelLocale::kKoKR, ExcelLocale::kThTH}) {
    EXPECT_EQ(cmp("각", ga_jamo, w), 0) << tag_of(w);
    EXPECT_GT(cmp("간", "가나", w), 0) << tag_of(w);
  }
  EXPECT_GT(cmp("가z", "갂a", ExcelLocale::kKoKR), 0);
  EXPECT_NE(cmp("각", ga_jamo, ExcelLocale::kZhCN), 0);
}

TEST(CollationRules, ZhWordSortAndWidth) {
  ScopedEvalProfile scope(mac(ExcelLocale::kZhCN));
  for (const ExcelLocale zh : {ExcelLocale::kZhCN, ExcelLocale::kZhTW}) {
    EXPECT_LT(cmp("ab", "a'b", zh), 0);
    EXPECT_LT(cmp("a'b", "a-b", zh), 0);
    EXPECT_NE(cmp("a'", "a", zh), 0);
    EXPECT_NE(cmp("A", "Ａ", zh), 0);
    EXPECT_EQ(cmp("A", "Ａ", zh, CollationStrength::kDynamic), 0);
    EXPECT_NE(cmp("ｶﾞ", "ガ", zh), 0);
    EXPECT_EQ(cmp("ｶﾞ", "ガ", zh, CollationStrength::kDynamic), 0);
    EXPECT_LT(cmp("เก", "ก", zh), 0);
  }
}

TEST(CollationRules, TrailingIgnorableSortsAfter) {
  ScopedEvalProfile scope(mac(ExcelLocale::kEnUS));
  EXPECT_GT(cmp("a" + u8(0x01), "a", ExcelLocale::kEnUS), 0);
  EXPECT_EQ(cmp("a" + u8(0x01) + "b", "ab", ExcelLocale::kEnUS), 0);
}

TEST(CollationRules, LongStringKeyStaysLinear) {
  std::string s;
  const char* kPieces[] = {"かー", "ｶﾞ", "漢字", "a", "ß", "각", "เก", "ゝ"};
  while (s.size() < 3 * 32767) {
    s += kPieces[s.size() % 8];
  }
  const std::size_t cps = 32767;
  std::string text;
  std::size_t count = 0;
  for (std::size_t i = 0; i < s.size() && count < cps; ++count) {
    std::size_t n = 0;
    decode_utf8_step(s, i, &n);
    text.append(s, i, n);
    i += n;
  }
  for (const ExcelLocale locale : {ExcelLocale::kJaJP, ExcelLocale::kZhCN, ExcelLocale::kKoKR}) {
    for (const CollationStrength strength :
         {CollationStrength::kCompare, CollationStrength::kDynamic, CollationStrength::kClassicExact}) {
      const std::string key = collation_key(text, mac(locale), strength);
      EXPECT_GT(key.size(), cps);
      EXPECT_LT(key.size(), 40 * cps);
    }
  }
}

// ---------------------------------------------------------------------------
// ASCII fast path versus the general key path.
// ---------------------------------------------------------------------------

std::vector<std::string> random_ascii(std::size_t count) {
  std::mt19937 rng(20261011u);
  const std::string alphabet = " !\"#$%&'()*+,-./0123456789:;<=>?@AZaz[\\]^_`{|}~bBmMyY";
  std::vector<std::string> out = {"", "a", "A", "ab", "a b", "a-b", "a'b", "co-op", "coop", "10", "9"};
  while (out.size() < count) {
    std::string s;
    const std::size_t len = 1 + rng() % 6;
    for (std::size_t i = 0; i < len; ++i) {
      s.push_back(alphabet[rng() % alphabet.size()]);
    }
    out.push_back(s);
  }
  return out;
}

TEST(CollationFastPath, AsciiMatchesGeneralPath) {
  const std::vector<std::string> strings = random_ascii(160);
  // Outside zh, U+0001 is ignorable before other characters, so a leading one
  // forces the general path without changing the key.
  const std::string lead = u8(0x01);
  for (std::size_t li = 0; li < 14; ++li) {
    const auto locale = static_cast<ExcelLocale>(li);
    const ExcelProfile profile = mac(locale);
    ScopedEvalProfile scope(profile);
    for (const CollationStrength strength :
         {CollationStrength::kCompare, CollationStrength::kDynamic, CollationStrength::kClassicExact}) {
      for (const std::string& s : strings) {
        if (!s.empty() && !is_zh(locale)) {
          ASSERT_EQ(collation_key(s, profile, strength), collation_key(lead + s, profile, strength))
              << tag_of(locale) << " '" << s << "'";
        }
      }
      for (const std::string& a : strings) {
        for (const std::string& b : strings) {
          const int fast = collate(a, b, profile, strength);
          ASSERT_EQ(fast, key_sign(a, b, profile, strength)) << tag_of(locale) << " '" << a << "' '" << b << "'";
          if (!a.empty() && !b.empty() && !is_zh(locale)) {
            ASSERT_EQ(fast, collate(lead + a, lead + b, profile, strength))
                << tag_of(locale) << " '" << a << "' '" << b << "'";
          }
        }
      }
    }
  }
}

TEST(CollationFastPath, CaseInsensitiveAndPrefixOrder) {
  for (const ExcelLocale locale : {ExcelLocale::kJaJP, ExcelLocale::kEnUS, ExcelLocale::kRuRU}) {
    ScopedEvalProfile scope(mac(locale));
    EXPECT_EQ(cmp("ABC", "abc", locale), 0);
    EXPECT_LT(cmp("ab", "abc", locale), 0);
    EXPECT_LT(cmp("", "a", locale), 0);
    EXPECT_GT(cmp("9", "10", locale), 0) << "digits order by character, not value";
  }
}

// ---------------------------------------------------------------------------
// Windows host: the engine's byte model.
// ---------------------------------------------------------------------------

int wcmp(std::string_view a, std::string_view b, CollationStrength strength = CollationStrength::kCompare) {
  return collate(a, b, win(ExcelLocale::kJaJP), strength);
}

TEST(CollationWinLegacy, AsciiCaseInsensitiveUnsignedBytes) {
  ScopedEvalProfile scope(win(ExcelLocale::kJaJP));
  EXPECT_EQ(wcmp("abc", "ABC"), 0);
  EXPECT_LT(wcmp("_", "A"), 0) << "letters fold to lower case before the byte compare";
  EXPECT_GT(wcmp("é", "z"), 0) << "bytes compare unsigned";
  EXPECT_LT(wcmp("ab", "abc"), 0);
  EXPECT_NE(wcmp("ß", "ss"), 0);
  EXPECT_NE(wcmp("Ａ", "A"), 0);
}

TEST(CollationWinLegacy, HalfwidthVoicingComposesOutsideCompare) {
  ScopedEvalProfile scope(win(ExcelLocale::kJaJP));
  EXPECT_NE(wcmp("ｶﾞ", "ガ"), 0);
  EXPECT_EQ(wcmp("ｶﾞ", "ガ", CollationStrength::kDynamic), 0);
  EXPECT_EQ(wcmp("ｶﾞ", "ガ", CollationStrength::kClassicExact), 0);
  EXPECT_EQ(wcmp("ﾊﾟ", "パ", CollationStrength::kDynamic), 0);
  EXPECT_NE(wcmp("ｱ", "ア", CollationStrength::kDynamic), 0);
}

TEST(CollationWinLegacy, KeysAgreeWithCollate) {
  const std::vector<std::string> strings = {"", "a", "A", "_", "é", "z", "ｶﾞ", "ガ", "あ", "ア", "ｱ", "abc", "ABD"};
  for (const ExcelLocale locale : {ExcelLocale::kJaJP, ExcelLocale::kEnUS, ExcelLocale::kZhCN}) {
    const ExcelProfile profile = win(locale);
    ScopedEvalProfile scope(profile);
    for (const CollationStrength strength :
         {CollationStrength::kCompare, CollationStrength::kDynamic, CollationStrength::kClassicExact}) {
      for (const std::string& a : strings) {
        for (const std::string& b : strings) {
          EXPECT_EQ(collate(a, b, profile, strength), key_sign(a, b, profile, strength)) << a << " | " << b;
        }
      }
    }
  }
}

// Criteria count over `cells` with the Windows digit pre-fold, as COUNTIF does.
int win_count(std::string_view criterion, std::vector<std::string> cells) {
  const ExcelProfile profile = win(ExcelLocale::kJaJP);
  int count = 0;
  for (const std::string& cell : cells) {
    count += collate(criteria_prefold(cell, profile), criteria_prefold(criterion, profile), profile,
                     CollationStrength::kCompare) == 0;
  }
  return count;
}

// First 1-based position equal to `value` at `strength`, 0 when none.
int win_find(std::string_view value, std::vector<std::string> cells, CollationStrength strength) {
  for (std::size_t i = 0; i < cells.size(); ++i) {
    if (collate(cells[i], value, win(ExcelLocale::kJaJP), strength) == 0) {
      return static_cast<int>(i) + 1;
    }
  }
  return 0;
}

// win-365-ja_JP goldens (countif / lookup / dfunc kana folding probes). The
// wildcard cases and the D-function header cases go through other paths.
TEST(CollationWinLegacy, WindowsKanaFoldingGoldens) {
  ScopedEvalProfile scope(win(ExcelLocale::kJaJP));
  EXPECT_EQ(win_count("あ", {"あ", "ア", "い"}), 1);
  EXPECT_EQ(win_count("ア", {"あ", "ア", "い"}), 1);
  EXPECT_EQ(win_count("アップル", {"アップル", "あっぷる", "アップル"}), 2);
  EXPECT_EQ(win_count("が", {"が", "ガ", "か"}), 1);
  EXPECT_EQ(win_count("ぱ", {"ぱ", "パ", "ば"}), 1);
  EXPECT_EQ(win_count("ベー", {"ベー", "べー", "ベ"}), 1);
  EXPECT_EQ(win_count("ｱ", {"ｱ", "ア", "イ"}), 1);
  EXPECT_EQ(win_count("ｱ", {"ｱ", "あ", "い"}), 1);
  EXPECT_EQ(win_count("ｶﾞ", {"ｶﾞ", "ガ", "カ"}), 1);
  EXPECT_EQ(win_count("＠", {"＠", "@", "x"}), 1);
  EXPECT_EQ(win_count("１２３", {"１２３", "123", "abc"}), 2);
  EXPECT_EQ(win_count("ＡＢＣ", {"ＡＢＣ", "ABC", "xyz"}), 1);
  EXPECT_EQ(win_count("a　b", {"a　b", "a b", "ab"}), 1);
  EXPECT_EQ(win_count("あ", {"あ", "ア", "い"}), 1);  // SUMIF

  const CollationStrength exact = CollationStrength::kClassicExact;
  const CollationStrength dynamic = CollationStrength::kDynamic;
  EXPECT_EQ(win_find("あ", {"ア", "い", "う"}, exact), 0);  // MATCH
  EXPECT_EQ(win_find("ア", {"あ", "い", "う"}, exact), 0);  // MATCH
  EXPECT_EQ(win_find("ｱ", {"ア", "イ", "ウ"}, exact), 0);   // MATCH
  EXPECT_EQ(win_find("ＡＢＣ", {"ABC", "DEF", "GHI"}, exact), 0);
  EXPECT_EQ(win_find("１２３", {"123", "456", "789"}, exact), 0);
  EXPECT_EQ(win_find("あ", {"ア", "い", "う"}, exact), 0);    // VLOOKUP
  EXPECT_EQ(win_find("ｱ", {"ア", "イ", "ウ"}, exact), 0);     // VLOOKUP
  EXPECT_EQ(win_find("あ", {"ア", "い", "う"}, dynamic), 0);  // XLOOKUP
  EXPECT_EQ(win_find("ｶﾞ", {"ガ", "ギ", "グ"}, dynamic), 1);  // XLOOKUP
  EXPECT_EQ(win_find("あ", {"ア", "い", "う"}, dynamic), 0);  // XMATCH
  EXPECT_EQ(win_find("あ", {"ア", "い", "う"}, exact), 0);    // HLOOKUP

  const ExcelProfile profile = win(ExcelLocale::kJaJP);
  for (const char* cell : {"アップル", "アップル", "ばなな"}) {
    EXPECT_FALSE(collate_prefix_equal(cell, "あっぷる", profile));
    EXPECT_FALSE(collate_prefix_equal(cell, "ｱｯﾌﾟﾙ", profile));
  }
  EXPECT_FALSE(collate_prefix_equal("ガ", "ｶﾞ", profile));
  EXPECT_FALSE(collate_prefix_equal("ギ", "ｶﾞ", profile));
  EXPECT_TRUE(collate_prefix_equal("Smithfield", "sm", profile));
}

}  // namespace
}  // namespace eval
}  // namespace formulon
