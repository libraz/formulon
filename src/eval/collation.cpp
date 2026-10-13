// Text collation keys. The Mac model mirrors `Collator` in
// tools/collation/generate_tables.py key for key: the generated tables are
// decoded lazily per locale into sorted run arrays plus dense per-ideograph
// arrays for the locale's Han rule, a string is tokenized with the contextual
// rules (half-width voicing, ja prolonged-sound and iteration marks, zh repeat
// marks, Thai prefix vowels), and each token's elements feed the key parts.
//
// Windows profiles reuse the locale's tables under the host rules measured on
// win-365-ja_JP, an NLS-style word sort: script blocks keep the base order with
// the locale's Han rule; hyphen, apostrophe and C0/C1 controls weigh only after
// every other level; width and enclosed or super/subscript forms differ at the
// compare-only level, kana type and smallness at the dynamic tertiary level;
// ligatures expand exactly; prolonged-sound and iteration marks follow any kana
// and sort above their substitute; number forms take their own primaries before
// the digits; Thai prefix vowels stay in place and sort before the consonants.

#include "eval/collation.h"

#include <algorithm>
#include <array>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <iterator>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/collation_table.h"
#include "eval/dbcs_table.h"
#include "eval/jp_kana_table.h"
#include "utils/text_ops.h"

namespace formulon {
namespace eval {
namespace {

namespace cd = collation_detail;

constexpr std::uint32_t kUroFirst = 0x4E00u;
constexpr std::uint32_t kUroLast = 0x9FFFu;
constexpr std::size_t kUroSize = kUroLast - kUroFirst + 1u;
constexpr std::uint32_t kSyllableFirst = 0xAC00u;
constexpr std::uint32_t kSyllableLast = 0xD7A3u;
constexpr std::int32_t kWeightRadix = 254;
constexpr std::int32_t kMaxWeight = cd::kMaxCollationWeight;

// Element classes; kMark is an expansion or repeat level marker whose level sits in `s`.
constexpr std::uint8_t kFull = static_cast<std::uint8_t>(cd::CollationClass::kFull);
constexpr std::uint8_t kMid = static_cast<std::uint8_t>(cd::CollationClass::kMid);
constexpr std::uint8_t kWord = static_cast<std::uint8_t>(cd::CollationClass::kWord);
constexpr std::uint8_t kSec = static_cast<std::uint8_t>(cd::CollationClass::kSec);
constexpr std::uint8_t kBase = static_cast<std::uint8_t>(cd::CollationClass::kBase);
constexpr std::uint8_t kMark = 5;

bool in_uro(std::uint32_t cp) noexcept {
  return cp >= kUroFirst && cp <= kUroLast;
}

struct Weight {
  std::uint8_t cls;
  std::int32_t p, s, d, c;
};

// Piecewise shift of primaries: (first source P, its target P), sorted by source.
using Segments = std::vector<std::pair<std::int32_t, std::int32_t>>;

std::int32_t apply_segments(const Segments& segments, std::int32_t p) noexcept {
  auto it =
      std::upper_bound(segments.begin(), segments.end(), p,
                       [](std::int32_t v, const std::pair<std::int32_t, std::int32_t>& s) { return v < s.first; });
  if (it == segments.begin()) {
    return p;
  }
  --it;
  return it->second + (p - it->first);
}

// ---------------------------------------------------------------- table decoding

class Varints {
 public:
  explicit Varints(const cd::CollationColumn& column) noexcept : cur_(column.data), end_(column.data + column.size) {}

  std::uint32_t next() noexcept {
    std::uint32_t value = 0;
    int shift = 0;
    while (cur_ < end_) {
      const std::uint8_t byte = *cur_++;
      value |= static_cast<std::uint32_t>(byte & 0x7Fu) << shift;
      if ((byte & 0x80u) == 0) {
        break;
      }
      shift += 7;
    }
    return value;
  }

  std::int32_t next_signed() noexcept { return unzigzag(next()); }

  static std::int32_t unzigzag(std::uint32_t v) noexcept {
    return (v & 1u) != 0 ? -static_cast<std::int32_t>((v + 1u) >> 1) : static_cast<std::int32_t>(v >> 1);
  }

 private:
  const std::uint8_t* cur_;
  const std::uint8_t* end_;
};

struct Run {
  std::uint32_t first;
  std::uint32_t count;
  std::int32_t p;
  std::uint16_t s;
  std::uint8_t cls;
  std::uint8_t step;
};

struct Tertiary {
  std::uint32_t cp;
  std::uint16_t d, c;
};

struct RunTable {
  std::vector<Run> runs;
  std::vector<Tertiary> tertiary;

  RunTable(const cd::CollationRunTable& table, const cd::CollationTertiaryTable& tert) {
    Varints gap(table.gap), length(table.length), p(table.p), s(table.s), cls_step(table.cls_step);
    runs.reserve(table.count);
    std::uint32_t prev_end = 0;
    std::int32_t prev_p = 0;
    for (std::uint32_t i = 0; i < table.count; ++i) {
      Run r{};
      r.first = prev_end + gap.next();
      r.count = length.next() + 1u;
      r.p = prev_p + p.next_signed();
      r.s = static_cast<std::uint16_t>(s.next());
      const std::uint32_t cs = cls_step.next();
      r.cls = static_cast<std::uint8_t>(cs & 7u);
      r.step = static_cast<std::uint8_t>(cs >> 3);
      prev_end = r.first + r.count;
      prev_p = r.p + static_cast<std::int32_t>((r.count - 1u) * r.step);
      runs.push_back(r);
    }
    Varints tgap(tert.gap), td(tert.d), tc(tert.c);
    tertiary.reserve(tert.count);
    std::uint32_t cp = 0;
    for (std::uint32_t i = 0; i < tert.count; ++i) {
      cp += tgap.next();
      const auto d = static_cast<std::uint16_t>(td.next());
      tertiary.push_back(Tertiary{cp, d, static_cast<std::uint16_t>(tc.next())});
    }
  }

  bool find(std::uint32_t cp, Weight* out) const noexcept {
    auto it = std::upper_bound(runs.begin(), runs.end(), cp, [](std::uint32_t v, const Run& r) { return v < r.first; });
    if (it == runs.begin()) {
      return false;
    }
    --it;
    if (cp - it->first >= it->count) {
      return false;
    }
    *out = Weight{it->cls, it->p + static_cast<std::int32_t>((cp - it->first) * it->step), it->s, 0, 0};
    auto t = std::lower_bound(tertiary.begin(), tertiary.end(), cp,
                              [](const Tertiary& e, std::uint32_t v) { return e.cp < v; });
    if (t != tertiary.end() && t->cp == cp) {
      out->d = t->d;
      out->c = t->c;
    }
    return true;
  }

  bool contains(std::uint32_t cp) const noexcept {
    Weight unused{};
    return find(cp, &unused);
  }
};

struct Expansion {
  std::uint32_t cp;
  std::uint32_t offset;
  std::uint8_t count;
  std::uint8_t level;
};

struct ExpansionTable {
  std::vector<Expansion> entries;
  std::vector<std::uint32_t> targets;

  explicit ExpansionTable(const cd::CollationExpansionTable& table) {
    Varints gap(table.gap), head(table.head), tgt(table.targets);
    entries.reserve(table.count);
    std::uint32_t cp = 0;
    std::int32_t first = 0;
    for (std::uint32_t i = 0; i < table.count; ++i) {
      cp += gap.next();
      const std::uint32_t h = head.next();
      const std::uint32_t n = (h >> 2) + 1u;
      Expansion e{cp, static_cast<std::uint32_t>(targets.size()), static_cast<std::uint8_t>(n),
                  static_cast<std::uint8_t>(h & 3u)};
      std::int32_t last = first;
      for (std::uint32_t k = 0; k < n; ++k) {
        last += tgt.next_signed();
        targets.push_back(static_cast<std::uint32_t>(last));
      }
      first = static_cast<std::int32_t>(targets[e.offset]);
      entries.push_back(e);
    }
  }

  const Expansion* find(std::uint32_t cp) const noexcept {
    auto it = std::lower_bound(entries.begin(), entries.end(), cp,
                               [](const Expansion& e, std::uint32_t v) { return e.cp < v; });
    return it != entries.end() && it->cp == cp ? &*it : nullptr;
  }
};

const std::vector<std::uint16_t>& han_rank() {
  static const std::vector<std::uint16_t> ranks = [] {
    std::vector<std::uint16_t> out(kUroSize, 0);
    const cd::CollationRankTable& t = cd::kCollationHanRank;
    Varints gap(t.gap), length(t.length), rank_col(t.rank), step_col(t.step);
    std::uint32_t prev_end = kUroFirst;
    std::int32_t rank = 0;
    for (std::uint32_t i = 0; i < t.count; ++i) {
      const std::uint32_t first = prev_end + gap.next();
      rank += rank_col.next_signed();
      const std::uint32_t n = length.next() + 1u;
      const auto step = static_cast<std::int32_t>(step_col.next());
      for (std::uint32_t k = 0; k < n; ++k) {
        out[first + k - kUroFirst] = static_cast<std::uint16_t>(rank + static_cast<std::int32_t>(k) * step);
      }
      prev_end = first + n;
      rank += static_cast<std::int32_t>(n - 1u) * step;
    }
    return out;
  }();
  return ranks;
}

struct BaseData {
  RunTable elements;
  ExpansionTable expansions;
  std::vector<std::pair<std::uint32_t, std::uint8_t>> repeat_marks;
  std::int32_t han_rest_p;
  std::int32_t out_of_range_p;

  explicit BaseData(const cd::CollationBaseTable& t)
      : elements(t.elements, t.tertiary),
        expansions(t.expansions),
        han_rest_p(t.han_rest_p),
        out_of_range_p(t.out_of_range_p) {
    Varints cp_col(t.repeat_marks.cp), kind_col(t.repeat_marks.kind);
    std::uint32_t cp = 0;
    for (std::uint32_t i = 0; i < t.repeat_marks.count; ++i) {
      cp += cp_col.next();
      repeat_marks.emplace_back(cp, static_cast<std::uint8_t>(kind_col.next()));
    }
  }

  int repeat_kind(std::uint32_t cp) const noexcept {
    for (const auto& [mark, kind] : repeat_marks) {
      if (mark == cp) {
        return kind;
      }
    }
    return -1;
  }
};

const BaseData& base_w() {
  static const BaseData data(*cd::kCollationLocales[static_cast<std::size_t>(ExcelLocale::kEnUS)].base);
  return data;
}

const BaseData& base_z() {
  static const BaseData data(*cd::kCollationLocales[static_cast<std::size_t>(ExcelLocale::kZhCN)].base);
  return data;
}

// Unified ideographs of the rule's code page rows, in code order.
std::vector<std::uint32_t> codepage_ideographs(const cd::CollationHanRule& rule) {
  std::vector<std::uint32_t> out;
  for (unsigned row = rule.first_row; row <= rule.last_row; ++row) {
    for (unsigned cell = 1; cell <= kDbcsGridSize; ++cell) {
      const std::uint32_t cp =
          lookup_dbcs_to_unicode(rule.codepage, static_cast<std::uint8_t>(row), static_cast<std::uint8_t>(cell));
      if (in_uro(cp)) {
        out.push_back(cp);
      }
    }
  }
  return out;
}

struct LocaleData {
  const BaseData* base;
  std::uint8_t rules;
  Segments remap;
  RunTable overrides;
  ExpansionTable expansions;
  std::vector<std::uint16_t> han_p;         // per unified ideograph: locale P, 0 when the rule does not place it
  std::vector<std::uint32_t> readings;      // per unified ideograph: syllable << 16 | rank, 0 when none
  std::array<std::int32_t, 128> ascii_p{};  // printable ASCII primary, -1 when not a plain primary
  bool ascii_fast = false;
  bool win = false;                 // Windows host rules
  std::vector<Segments> win_order;  // Windows: applied in turn to every table primary
  std::int32_t number_form_p = 0;   // Windows: primary of the first number form

  explicit LocaleData(const cd::CollationLocaleTable& t)
      : base(t.base == cd::kCollationLocales[static_cast<std::size_t>(ExcelLocale::kEnUS)].base ? &base_w()
                                                                                                : &base_z()),
        rules(t.rules),
        overrides(t.overrides, t.override_tertiary),
        expansions(t.expansions) {
    Varints base_p(t.remap.base_p), offset(t.remap.offset);
    std::int32_t b = 0;
    std::int32_t off = 0;
    for (std::uint32_t i = 0; i < t.remap.count; ++i) {
      b += static_cast<std::int32_t>(base_p.next());
      off += offset.next_signed();
      remap.emplace_back(b, b + off);
    }
    switch (t.han.kind) {
      case cd::CollationHanKind::kCodepage:
        load_codepage(t.han);
        break;
      case cd::CollationHanKind::kStrokes:
        load_strokes(t.han.strokes);
        break;
      case cd::CollationHanKind::kReadings:
        load_readings(t.han);
        break;
      case cd::CollationHanKind::kBase:
        break;
    }
    ascii_fast = (rules & cd::kRuleIgnoreUnlisted) == 0;
    ascii_p.fill(-1);
  }

  std::int32_t remap_p(std::int32_t p) const noexcept { return apply_segments(remap, p); }

  std::int32_t order_p(std::int32_t p) const noexcept {
    for (const Segments& s : win_order) {
      p = apply_segments(s, p);
    }
    return p;
  }

  void load_codepage(const cd::CollationHanRule& rule) {
    han_p.assign(kUroSize, 0);
    Varints slot_col(rule.list.slot), p_col(rule.list.p), length(rule.list.length);
    std::int32_t last_slot = -1;
    std::int32_t last_p = 0;
    for (std::uint32_t i = 0; i < rule.list.count; ++i) {
      const std::int32_t slot = last_slot + 1 + slot_col.next_signed();
      const std::int32_t p = last_p + static_cast<std::int32_t>(p_col.next());
      const auto n = static_cast<std::int32_t>(length.next()) + 1;
      for (std::int32_t k = 0; k < n; ++k) {
        const std::int32_t cell = slot + k;
        const std::uint32_t cp =
            lookup_dbcs_to_unicode(rule.codepage, static_cast<std::uint8_t>(rule.first_row + cell / 94),
                                   static_cast<std::uint8_t>(cell % 94 + 1));
        if (in_uro(cp)) {
          han_p[cp - kUroFirst] = static_cast<std::uint16_t>(p + k);
        }
      }
      last_slot = slot + n - 1;
      last_p = p + n - 1;
    }
  }

  void load_strokes(const cd::CollationStrokeTable& t) {
    std::vector<std::pair<std::int32_t, std::uint32_t>> order;  // (group, cp)
    std::vector<std::uint32_t> last_group;
    Varints gap(t.gap), length(t.length), group(t.group);
    std::uint32_t prev_end = kUroFirst;
    std::int32_t chained = 0;
    std::int32_t top = 0;
    for (std::uint32_t i = 0; i < t.run_count; ++i) {
      const std::uint32_t first = prev_end + gap.next();
      const std::uint32_t n = length.next() + 1u;
      const std::uint32_t g = group.next();
      if (g != 0) {
        chained += Varints::unzigzag(g - 1u);
        top = std::max(top, chained);
      }
      for (std::uint32_t k = 0; k < n; ++k) {
        if (g != 0) {
          order.emplace_back(chained, first + k);
        } else {
          last_group.push_back(first + k);
        }
      }
      prev_end = first + n;
    }
    for (std::uint32_t cp : last_group) {
      order.emplace_back(top + 1, cp);
    }
    std::sort(order.begin(), order.end());
    han_p.assign(kUroSize, 0);
    Varints p_col(t.p), p_len(t.p_length);
    std::size_t idx = 0;
    std::int32_t last = 0;
    for (std::uint32_t i = 0; i < t.p_count; ++i) {
      const std::int32_t first = last + static_cast<std::int32_t>(p_col.next());
      const std::uint32_t n = p_len.next() + 1u;
      for (std::uint32_t k = 0; k < n && idx < order.size(); ++k, ++idx) {
        han_p[order[idx].second - kUroFirst] = static_cast<std::uint16_t>(first + static_cast<std::int32_t>(k));
      }
      last = first + static_cast<std::int32_t>(n - 1u);
    }
  }

  void load_readings(const cd::CollationHanRule& rule) {
    readings.assign(kUroSize, 0);
    const cd::CollationReadingTable& t = rule.readings;
    const std::vector<std::uint32_t> cells = codepage_ideographs(rule);
    std::vector<std::uint16_t> counts(kSyllableLast - kSyllableFirst + 1u, 0);
    Varints syl_col(t.syllable), length(t.length);
    std::int32_t syl = 0;
    std::size_t at = 0;
    for (std::uint32_t i = 0; i < t.run_count; ++i) {
      syl += syl_col.next_signed();
      const std::uint32_t n = length.next() + 1u;
      for (std::uint32_t k = 0; k < n && at < cells.size(); ++k, ++at) {
        if (syl != 0) {
          const std::uint16_t rank = ++counts[static_cast<std::uint32_t>(syl) - kSyllableFirst];
          readings[cells[at] - kUroFirst] = static_cast<std::uint32_t>(syl) << 16 | rank;
        }
      }
    }
    Varints group_syl(t.group_syllable), group_size(t.group_size), unlisted(t.unlisted_cp);
    std::uint32_t gsyl = 0;
    for (std::uint32_t i = 0; i < t.group_count; ++i) {
      gsyl += group_syl.next();
      const std::uint32_t n = group_size.next() + 1u;
      std::int32_t cp = static_cast<std::int32_t>(kUroFirst);
      for (std::uint32_t rank = 1; rank <= n; ++rank) {
        cp += unlisted.next_signed();
        readings[static_cast<std::uint32_t>(cp) - kUroFirst] = gsyl << 16 | (counts[gsyl - kSyllableFirst] + rank);
      }
    }
  }
};

// ---------------------------------------------------------------- kana helpers

// Vowel of each hiragana letter U+3041..U+3096 ('-' for none); katakana
// U+30A1..U+30F6 follows the same layout.
constexpr char kKanaVowels[] =
    "aaiiuueeoo"
    "aaiiuueeoo"
    "aaiiuueeoo"
    "aaiiuuueeoo"
    "aiueo"
    "aaaiiiuuueeeooo"
    "aiueo"
    "aauuoo"
    "aiueo"
    "aaieo-uae";
constexpr char kSmallKatakanaVowels[] = "uiuouaiueouaiueo";  // U+31F0..U+31FF

struct SupplementKana {
  std::uint32_t cp;
  char vowel;
  bool katakana;
};

// Kana letters outside the BMP blocks above.
constexpr SupplementKana kSupplementKana[] = {
    {0x1B000u, 'e', true},  {0x1B001u, 'e', false}, {0x1B11Fu, 'u', false}, {0x1B120u, 'i', true},
    {0x1B121u, 'e', true},  {0x1B122u, 'u', true},  {0x1B132u, 'o', false}, {0x1B150u, 'i', false},
    {0x1B151u, 'e', false}, {0x1B152u, 'o', false}, {0x1B155u, 'o', true},  {0x1B164u, 'i', true},
    {0x1B165u, 'e', true},  {0x1B166u, 'o', true},  {0x1B167u, '-', true},
};

const SupplementKana* supplement_kana(std::uint32_t cp) noexcept {
  for (const SupplementKana& k : kSupplementKana) {
    if (k.cp == cp) {
      return &k;
    }
  }
  return nullptr;
}

std::uint32_t vowel_letter(char vowel, bool katakana) noexcept {
  static constexpr char kOrder[] = "aiueo";
  for (std::uint32_t i = 0; i < 5; ++i) {
    if (kOrder[i] == vowel) {
      return 0x3042u + 2u * i + (katakana ? 0x60u : 0u);
    }
  }
  return 0;
}

bool is_hiragana_letter(std::uint32_t cp) noexcept {
  if (cp >= 0x3041u && cp <= 0x3096u) {
    return true;
  }
  const SupplementKana* k = supplement_kana(cp);
  return k != nullptr && !k->katakana;
}

bool is_katakana_letter(std::uint32_t cp) noexcept {
  if ((cp >= 0x30A1u && cp <= 0x30FAu) || (cp >= 0x31F0u && cp <= 0x31FFu) || (cp >= 0xFF66u && cp <= 0xFF6Fu) ||
      (cp >= 0xFF71u && cp <= 0xFF9Du)) {
    return true;
  }
  if (cp >= 0x1AFF0u && cp <= 0x1AFFEu) {
    return cp != 0x1AFF4u && cp != 0x1AFFCu;
  }
  const SupplementKana* k = supplement_kana(cp);
  return k != nullptr && k->katakana;
}

// The vowel kana a following prolonged-sound mark stands for, or 0.
std::uint32_t kana_vowel(std::uint32_t cp) noexcept {
  if (cp >= 0x3041u && cp <= 0x3096u) {
    return vowel_letter(kKanaVowels[cp - 0x3041u], false);
  }
  if (cp >= 0x30A1u && cp <= 0x30F6u) {
    return vowel_letter(kKanaVowels[cp - 0x30A1u], true);
  }
  if (cp >= 0x30F7u && cp <= 0x30FAu) {
    return vowel_letter("aieo"[cp - 0x30F7u], true);
  }
  if (cp >= 0x31F0u && cp <= 0x31FFu) {
    return vowel_letter(kSmallKatakanaVowels[cp - 0x31F0u], true);
  }
  if ((cp >= 0xFF66u && cp <= 0xFF6Fu) || (cp >= 0xFF71u && cp <= 0xFF9Du)) {
    return kana_vowel(half_to_full_katakana_base(cp));
  }
  const SupplementKana* k = supplement_kana(cp);
  return k != nullptr ? vowel_letter(k->vowel, k->katakana) : 0;
}

// Voiced hiragana for an unvoiced or voiced hiragana `k`, or 0.
std::uint32_t voiced_hiragana(std::uint32_t k) noexcept {
  if (k >= 0x304Bu && k <= 0x3062u) {
    return k - ((k - 0x304Bu) & 1u) + 1u;
  }
  if (k >= 0x3064u && k <= 0x3069u) {
    return k - ((k - 0x3064u) & 1u) + 1u;
  }
  if (k >= 0x306Fu && k <= 0x307Du) {
    return k - (k - 0x306Fu) % 3u + 1u;
  }
  if (k == 0x3046u || k == 0x3094u) {
    return 0x3094u;
  }
  if (k == 0x309Du || k == 0x309Eu) {
    return 0x309Eu;
  }
  return 0;
}

std::uint32_t voiced_full_kana(std::uint32_t k) noexcept {
  if (k >= 0x30EFu && k <= 0x30F2u) {
    return k + 8u;  // ワ ヰ ヱ ヲ
  }
  if (k >= 0x30F7u && k <= 0x30FAu) {
    return k;
  }
  if (k >= 0x30A1u && k <= 0x30FEu) {
    const std::uint32_t h = voiced_hiragana(k - 0x60u);
    return h != 0 ? h + 0x60u : 0;
  }
  return voiced_hiragana(k);
}

struct VoicedSquare {
  std::uint32_t first, last, voiced;
};

// Squared words whose compatibility form starts with a voiceable kana.
constexpr VoicedSquare kVoicedSquares[] = {
    {0x3306u, 0x3306u, 0x30F4u},   {0x330Bu, 0x330Fu, 0x30ACu},   {0x3310u, 0x3317u, 0x30AEu},
    {0x3318u, 0x331Bu, 0x30B0u},   {0x331Cu, 0x331Cu, 0x30B2u},   {0x331Du, 0x331Eu, 0x30B4u},
    {0x331Fu, 0x3320u, 0x30B6u},   {0x3321u, 0x3321u, 0x30B8u},   {0x3322u, 0x3323u, 0x30BCu},
    {0x3324u, 0x3324u, 0x30C0u},   {0x3325u, 0x3325u, 0x30C7u},   {0x3326u, 0x3327u, 0x30C9u},
    {0x332Au, 0x332Du, 0x30D0u},   {0x332Eu, 0x3331u, 0x30D3u},   {0x3332u, 0x3335u, 0x30D6u},
    {0x3336u, 0x333Cu, 0x30D9u},   {0x333Du, 0x3342u, 0x30DCu},   {0x3357u, 0x3357u, 0x30F7u},
    {0x1F200u, 0x1F200u, 0x307Cu}, {0x1F201u, 0x1F201u, 0x30B4u}, {0x1F202u, 0x1F202u, 0x30B6u},
    {0x1F213u, 0x1F213u, 0x30C7u},
};

// The character composed from the first character of `cp`'s compatibility form
// and a voiced sound mark, or 0 when that does not compose.
std::uint32_t voiced_form(std::uint32_t cp) noexcept {
  if (cp >= 0xFF66u && cp <= 0xFF9Du) {
    return voiced_full_kana(half_to_full_katakana_base(cp));
  }
  if (cp >= 0x32D0u && cp <= 0x32FEu) {  // circled katakana, ア..ワ then ヰ ヱ ヲ
    const std::uint32_t i = cp - 0x32D0u;
    return voiced_full_kana(i < 44u ? half_to_full_katakana_base(0xFF71u + i) : 0x30F0u + (i - 44u));
  }
  if (cp == 0x30FFu) {
    return 0x30B4u;  // ヿ reads コト
  }
  if (cp > 0x30FFu) {
    for (const VoicedSquare& sq : kVoicedSquares) {
      if (cp >= sq.first && cp <= sq.last) {
        return sq.voiced;
      }
    }
    return 0;
  }
  return voiced_full_kana(cp);
}

// Half-width kana followed by U+FF9E or U+FF9F composed into one full-width kana, or 0.
std::uint32_t halfwidth_voiced(std::uint32_t half, std::uint32_t mark) noexcept {
  if (mark == 0xFF9Eu) {
    return voiced_form(half);
  }
  const std::uint32_t full = half_to_full_katakana_base(half);
  return full >= 0x30CFu && full <= 0x30DDu && (full - 0x30CFu) % 3u == 0 ? full + 2u : 0;
}

// ---------------------------------------------------------------- Windows host

constexpr std::int32_t kNumberFormCount = 3 + 0x40;  // U+00BC..U+00BE, U+2150..U+218F
constexpr std::int32_t kWinMarkD = 4;                // kana mark above its substitute

std::int32_t number_form_index(std::uint32_t cp) noexcept {
  if (cp >= 0x00BCu && cp <= 0x00BEu) {
    return static_cast<std::int32_t>(cp - 0x00BCu);
  }
  return cp >= 0x2150u && cp <= 0x218Fu ? static_cast<std::int32_t>(cp - 0x2150u) + 3 : -1;
}

// Word-sort weight of a hyphen or apostrophe, compared after every other level; -1 otherwise.
std::int32_t win_word_weight(std::uint32_t cp) noexcept {
  switch (cp) {
    case 0x0027u:
    case 0xFF07u:
      return 1;
    case 0x002Du:
    case 0xFF0Du:
      return 2;
    default:
      return -1;
  }
}

// Compare-only width weight: -1 for narrow forms, 1 for wide, enclosed and super/subscript forms.
std::int32_t win_width(std::uint32_t cp) noexcept {
  if (cp == 0xFF9Eu || cp == 0xFF9Fu) {
    return 0;
  }
  if ((cp >= 0xFF61u && cp <= 0xFFDCu) || (cp >= 0xFFE8u && cp <= 0xFFEEu)) {
    return -1;
  }
  if (cp == 0x3000u || (cp >= 0xFF01u && cp <= 0xFF60u) || (cp >= 0xFFE0u && cp <= 0xFFE6u)) {
    return 1;
  }
  if (cp == 0x00B2u || cp == 0x00B3u || cp == 0x00B9u || (cp >= 0x2070u && cp <= 0x209Fu) ||
      (cp >= 0x2460u && cp <= 0x24FFu)) {
    return 1;
  }
  return 0;
}

bool is_small_kana(std::uint32_t cp) noexcept {
  static constexpr std::uint16_t kSmall[] = {0x3041u, 0x3043u, 0x3045u, 0x3047u, 0x3049u, 0x3063u, 0x3083u, 0x3085u,
                                             0x3087u, 0x308Eu, 0x3095u, 0x3096u, 0x30A1u, 0x30A3u, 0x30A5u, 0x30A7u,
                                             0x30A9u, 0x30C3u, 0x30E3u, 0x30E5u, 0x30E7u, 0x30EEu, 0x30F5u, 0x30F6u};
  if ((cp >= 0x31F0u && cp <= 0x31FFu) || (cp >= 0xFF67u && cp <= 0xFF6Fu)) {
    return true;
  }
  return std::find(std::begin(kSmall), std::end(kSmall), cp) != std::end(kSmall);
}

// Kana-type weight at the dynamic tertiary level: katakana below hiragana, small below normal.
std::int32_t win_kana_d(std::uint32_t cp) noexcept {
  const bool hiragana = is_hiragana_letter(cp);
  if (!hiragana && !is_katakana_letter(cp)) {
    return 0;
  }
  return (hiragana ? 2 : 0) + (is_small_kana(cp) ? 0 : 1);
}

// ---------------------------------------------------------------- elements

std::int32_t out_of_range_index(std::uint32_t cp) noexcept {
  std::int64_t index = cp;
  for (const cd::CollationRange& r : cd::kCollationExactRanges) {
    const std::int64_t below = std::min<std::int64_t>(static_cast<std::int64_t>(r.last) + 1, cp) - r.first;
    index -= std::max<std::int64_t>(0, below);
  }
  return static_cast<std::int32_t>(index);
}

// Weight of a code point that is a single element, before expansions.
bool single_weight(const LocaleData& L, std::uint32_t cp, Weight* out) noexcept {
  if (L.overrides.find(cp, out)) {
    return true;
  }
  if (L.base->elements.find(cp, out)) {
    if (out->cls == kBase) {
      out->p = L.remap_p(out->p);
    }
    return true;
  }
  if (in_uro(cp)) {
    const std::size_t i = cp - kUroFirst;
    if (!L.han_p.empty() && L.han_p[i] != 0) {
      *out = Weight{kBase, L.han_p[i], 0, 0, 0};
      return true;
    }
    if (L.readings.empty() || L.readings[i] == 0) {
      *out = Weight{kBase, L.remap_p(L.base->han_rest_p + han_rank()[i]), 0, 0, 0};
      return true;
    }
  }
  return false;
}

void append_elements(const LocaleData& L, std::uint32_t cp, bool expand, std::vector<Weight>& out, int depth);

// Elements from the locale tables; on Windows, primaries pass through `order_p`.
void append_table_elements(const LocaleData& L, std::uint32_t cp, bool expand, std::vector<Weight>& out, int depth) {
  const Expansion* e = L.expansions.find(cp);
  const ExpansionTable* pool = &L.expansions;
  if (e == nullptr && !L.overrides.contains(cp)) {
    e = L.base->expansions.find(cp);
    pool = &L.base->expansions;
  }
  if (e != nullptr && depth < 8) {
    std::int32_t level = L.win ? 0 : e->level;  // Windows expansions are exact
    if (!expand && e->count > 1 && level == 0) {
      level = 3;  // unexpanded, it still differs from its targets at the compare-only level
    }
    for (std::uint32_t k = 0; k < e->count; ++k) {
      append_elements(L, pool->targets[e->offset + k], expand, out, depth + 1);
    }
    if (level != 0) {
      out.push_back(Weight{kMark, 0, level, 0, 0});
    }
    return;
  }
  Weight w{};
  if (single_weight(L, cp, &w)) {
    if (w.cls == kBase) {
      w.p = L.order_p(w.p);
    }
    out.push_back(w);
    return;
  }
  if (in_uro(cp) && !L.readings.empty() && L.readings[cp - kUroFirst] != 0) {
    // A ko ideograph is its reading syllable plus its rank on the last element's S.
    const std::uint32_t r = L.readings[cp - kUroFirst];
    append_elements(L, r >> 16, expand, out, depth + 1);
    out.back().s += static_cast<std::int32_t>(r & 0xFFFFu);
    return;
  }
  if (cp >= kSyllableFirst && cp <= kSyllableLast && (L.rules & cd::kRuleHangulJamo) != 0) {
    const std::uint32_t s = cp - kSyllableFirst;
    append_elements(L, 0x1100u + s / 588u, expand, out, depth + 1);
    append_elements(L, 0x1161u + (s % 588u) / 28u, expand, out, depth + 1);
    if (s % 28u != 0) {
      append_elements(L, 0x11A7u + s % 28u, expand, out, depth + 1);
    }
    return;
  }
  if ((L.rules & cd::kRuleIgnoreUnlisted) != 0) {
    out.push_back(Weight{kFull, 0, 0, 0, 0});
    return;
  }
  if (cp <= 0xFFFFu) {
    out.push_back(Weight{kBase, L.order_p(L.remap_p(L.base->out_of_range_p + out_of_range_index(cp))), 0, 0, 0});
    return;
  }
  // Supplementary planes follow every BMP weight in code-point order, as two primaries.
  const std::int32_t top = L.remap_p(L.base->out_of_range_p + out_of_range_index(0xFFFFu));
  const auto index = static_cast<std::int32_t>(cp - 0x10000u);
  out.push_back(Weight{kBase, L.order_p(top + 1 + index / (kMaxWeight + 1)), 0, 0, 0});
  out.push_back(Weight{kBase, index % (kMaxWeight + 1), 0, 0, 0});
}

void append_elements(const LocaleData& L, std::uint32_t cp, bool expand, std::vector<Weight>& out, int depth) {
  if (!L.win) {
    append_table_elements(L, cp, expand, out, depth);
    return;
  }
  if (const std::int32_t i = number_form_index(cp); i >= 0) {
    out.push_back(Weight{kBase, L.number_form_p + i, 0, 0, 0});
    return;
  }
  if (cp == 0xFF9Eu || cp == 0xFF9Fu) {
    append_elements(L, cp - 0xFF9Eu + 0x309Bu, expand, out, depth + 1);  // the full-width voicing marks
    return;
  }
  if (cp == 0x200Bu) {
    append_elements(L, 0x20u, expand, out, depth + 1);
    return;
  }
  if (const std::int32_t word = win_word_weight(cp); word >= 0) {
    out.push_back(Weight{kWord, word, 0, 0, win_width(cp)});
    return;
  }
  const std::size_t start = out.size();
  append_table_elements(L, cp, expand, out, depth);
  // A precomposed syllable sorts after the same jamo typed separately.
  const std::int32_t d = win_kana_d(cp) + (cp >= kSyllableFirst && cp <= kSyllableLast ? 1 : 0);
  const std::int32_t c = win_width(cp);
  for (std::size_t k = start; k < out.size(); ++k) {
    Weight& w = out[k];
    if (w.cls == kMid && cp < 0xA0u) {
      w = Weight{kWord, 0, 0, 0, 0};  // a control character is a word-sort symbol
    } else if (w.cls != kMark) {
      w.d += d;
      w.c = w.c == 0 ? c : w.c;
    }
  }
}

// Windows keeps the base script order: the locale's remap blocks, each with the
// locale-only primaries that follow it, are laid out again in base order.
Segments base_block_order(const LocaleData& L) {
  const std::int32_t top = L.remap_p(L.base->out_of_range_p + out_of_range_index(0xFFFFu));
  const std::int32_t end = top + 2 + static_cast<std::int32_t>(0x10FFFFu - 0x10000u) / (kMaxWeight + 1);
  Segments by_locale = L.remap;  // (base P, locale P)
  std::sort(by_locale.begin(), by_locale.end(), [](const auto& a, const auto& b) { return a.second < b.second; });
  struct Block {
    std::int32_t base, first, end;  // locale P range [first, end)
  };
  std::vector<Block> blocks;
  for (std::size_t i = 0; i < by_locale.size(); ++i) {
    const std::int32_t next = i + 1 < by_locale.size() ? by_locale[i + 1].second : std::max(end, by_locale[i].second);
    blocks.push_back(Block{by_locale[i].first, by_locale[i].second, next});
  }
  std::sort(blocks.begin(), blocks.end(), [](const Block& a, const Block& b) { return a.base < b.base; });
  Segments out;
  std::int32_t at = 0;
  for (const Block& b : blocks) {
    out.emplace_back(b.first, at);
    at += b.end - b.first;
  }
  std::sort(out.begin(), out.end());
  return out;
}

void set_win_order(LocaleData& L) {
  L.win = true;
  L.rules = static_cast<std::uint8_t>(L.rules & ~cd::kRuleThaiReorder);
  if (L.base == &base_w()) {
    L.win_order.push_back(base_block_order(L));
  }
  // Thai prefix vowels move ahead of the consonants.
  Weight ko_kai{};
  Weight first{};
  if (single_weight(L, 0x0E01u, &ko_kai) && single_weight(L, 0x0E40u, &first)) {
    const std::int32_t pk = L.order_p(ko_kai.p);
    const std::int32_t pv = L.order_p(first.p);
    bool run = pv > pk;
    for (std::uint32_t k = 1; k < 5 && run; ++k) {
      Weight w{};
      run = single_weight(L, 0x0E40u + k, &w) && L.order_p(w.p) == pv + static_cast<std::int32_t>(k);
    }
    if (run) {
      L.win_order.push_back(Segments{{0, 0}, {pk, pk + 5}, {pv, pk}, {pv + 5, pv + 5}});
    }
  }
  // Number forms take their own primaries just before the digits.
  Weight zero{};
  if (single_weight(L, 0x30u, &zero)) {
    L.number_form_p = L.order_p(zero.p);
    L.win_order.push_back(Segments{{0, 0}, {L.number_form_p, L.number_form_p + kNumberFormCount}});
  }
}

LocaleData make_locale(std::size_t index, bool win) {
  LocaleData data(cd::kCollationLocales[index]);
  if (win) {
    set_win_order(data);
  }
  if (data.ascii_fast) {
    std::vector<Weight> expanded;
    std::vector<Weight> unexpanded;
    for (std::uint32_t ch = 0x20; ch < 0x7F; ++ch) {
      expanded.clear();
      unexpanded.clear();
      append_elements(data, ch, true, expanded, 0);
      append_elements(data, ch, false, unexpanded, 0);
      const bool plain = expanded.size() == 1 && unexpanded.size() == 1 && expanded[0].cls == kBase &&
                         expanded[0].s == 0 && expanded[0].d == 0 && expanded[0].c == 0 && unexpanded[0].cls == kBase &&
                         unexpanded[0].p == expanded[0].p && unexpanded[0].s == 0 && unexpanded[0].d == 0 &&
                         unexpanded[0].c == 0;
      data.ascii_p[ch] = plain ? expanded[0].p : -1;
    }
  }
  return data;
}

constexpr std::size_t kLocaleCount = std::size(cd::kCollationLocales);

// Index I below kLocaleCount is the Mac model of locale I, above it the Windows one.
template <std::size_t I>
const LocaleData& locale_data_at() {
  static const LocaleData data = make_locale(I % kLocaleCount, I >= kLocaleCount);
  return data;
}

using LocaleGetter = const LocaleData& (*)();

template <std::size_t... I>
constexpr std::array<LocaleGetter, sizeof...(I)> make_getters(std::index_sequence<I...>) {
  return {{&locale_data_at<I>...}};
}

const LocaleData& locale_data(ExcelProfile profile) {
  static constexpr auto kGetters = make_getters(std::make_index_sequence<2 * kLocaleCount>{});
  const std::size_t host = profile.host == ExcelHost::kWin365 ? kLocaleCount : 0;
  return kGetters[host + static_cast<std::size_t>(profile.locale)]();
}

// ---------------------------------------------------------------- keys

struct Token {
  std::uint32_t cp;
  std::uint32_t pos;     // code-point index in the original string
  std::uint32_t c_from;  // half-width base whose compare-only offset a composed kana keeps, or 0
  bool repeat;           // zh repeat mark: a secondary marker follows
  std::uint32_t mark;    // Windows: the kana mark this token substitutes, or 0
};

void tokenize(const LocaleData& L, std::string_view s, CollationStrength strength, std::vector<Token>& toks) {
  toks.clear();
  toks.reserve(s.size());
  std::uint32_t pos = 0;
  for (std::size_t i = 0; i < s.size();) {
    std::size_t n = 0;
    const std::uint32_t cp = decode_utf8_step(s, i, &n);
    i += n;
    if (strength != CollationStrength::kClassicExact && (cp == 0xFF9Eu || cp == 0xFF9Fu) && !toks.empty()) {
      Token& prev = toks.back();
      if (prev.c_from == 0 && prev.cp >= 0xFF66u && prev.cp <= 0xFF9Du) {
        if (const std::uint32_t v = halfwidth_voiced(prev.cp, cp); v != 0) {
          prev.c_from = prev.cp;
          prev.cp = v;
          ++pos;
          continue;
        }
      }
    }
    toks.push_back(Token{cp, pos++, 0, false, 0});
  }
  if (L.win || (L.rules & cd::kRuleKanaMarks) != 0) {
    for (std::size_t i = 1; i < toks.size(); ++i) {
      const std::uint32_t prev = toks[i - 1].cp;
      const std::uint32_t cp = toks[i].cp;
      // Windows repeats a kana of either script with either iteration mark.
      const bool hiragana = is_hiragana_letter(prev) || (L.win && is_katakana_letter(prev));
      const bool katakana = is_katakana_letter(prev) || (L.win && is_hiragana_letter(prev));
      std::uint32_t sub = 0;
      if (cp == 0x30FCu || cp == 0xFF70u) {
        sub = kana_vowel(prev);
      } else if ((cp == 0x309Du || cp == 0x309Eu) && hiragana) {
        sub = cp == 0x309Du ? prev : voiced_form(prev);
      } else if ((cp == 0x30FDu || cp == 0x30FEu) && katakana) {
        sub = cp == 0x30FDu ? prev : voiced_form(prev);
      }
      if (sub != 0) {
        toks[i] = Token{sub, toks[i].pos, 0, false, L.win ? cp : 0u};
      }
    }
  }
  if ((L.rules & cd::kRuleRepeatMarks) != 0) {
    for (std::size_t i = 1; i < toks.size(); ++i) {
      const int kind = L.base->repeat_kind(toks[i].cp);
      if (kind < 0) {
        continue;
      }
      const std::uint32_t prev = toks[i - 1].cp;
      std::uint32_t sub = 0;
      if (kind == static_cast<int>(cd::CollationRepeatKind::kVowel)) {
        sub = kana_vowel(prev);
      } else if (kind == static_cast<int>(cd::CollationRepeatKind::kVoiced)) {
        sub = voiced_form(prev);
      }
      toks[i] = Token{sub != 0 ? sub : prev, toks[i].pos, 0, true, 0};
    }
  }
  if ((L.rules & cd::kRuleThaiReorder) != 0) {
    for (std::size_t i = 0; i + 1 < toks.size();) {
      if (toks[i].cp >= 0x0E40u && toks[i].cp <= 0x0E44u && toks[i + 1].cp >= 0x0E01u && toks[i + 1].cp <= 0x0E2Eu) {
        std::swap(toks[i], toks[i + 1]);
        i += 2;
      } else {
        ++i;
      }
    }
  }
}

void put_weight(std::string& out, std::int32_t v) {
  v = std::clamp(v, 0, kMaxWeight);
  out.push_back(static_cast<char>(v / kWeightRadix + 2));
  out.push_back(static_cast<char>(v % kWeightRadix + 2));
}

std::int32_t first_c(const LocaleData& L, std::uint32_t cp, std::vector<Weight>& tmp) {
  tmp.clear();
  append_elements(L, cp, true, tmp, 0);
  return tmp.empty() ? 0 : tmp.front().c;
}

void table_key(const LocaleData& L, std::string_view s, CollationStrength strength, std::string& out) {
  std::vector<Token> toks;
  tokenize(L, s, strength, toks);
  std::vector<Weight> els;
  els.reserve(toks.size() + 4);
  std::vector<std::uint32_t> bounds;
  bounds.reserve(toks.size() + 1);
  std::vector<Weight> tmp;
  const bool expand = strength != CollationStrength::kClassicExact;
  for (const Token& t : toks) {
    const std::size_t start = els.size();
    bounds.push_back(static_cast<std::uint32_t>(start));
    append_elements(L, t.cp, expand, els, 0);
    if (t.c_from != 0 && els.size() > start) {
      // Composed half-width kana: the voiced kana's weights, C moved by the half-width offset.
      Weight w = els[start];
      w.c += first_c(L, t.c_from, tmp) - first_c(L, half_to_full_katakana_base(t.c_from), tmp);
      els.resize(start);
      els.push_back(w);
    }
    if (t.repeat) {
      els.push_back(Weight{kMark, 0, 1, 0, 0});
    }
    // Windows kana mark: its substitute's weights, above them at the kana-type level.
    for (std::size_t k = start; t.mark != 0 && k < els.size(); ++k) {
      els[k].d += kWinMarkD;
      els[k].c = win_width(t.mark) != 0 ? win_width(t.mark) : els[k].c;
    }
  }
  bounds.push_back(static_cast<std::uint32_t>(els.size()));
  const std::size_t n = toks.size();

  // Word symbols after the last token with a primary, secondary or marker also weigh primary 0.
  std::size_t tail_from = n;
  bool found = false;
  for (std::size_t i = n; i-- > 0 && !found;) {
    for (std::uint32_t k = bounds[i]; k < bounds[i + 1]; ++k) {
      const std::uint8_t cls = els[k].cls;
      if (cls == kBase || cls == kSec || cls == kMark) {
        tail_from = i + 1;
        found = true;
        break;
      }
    }
  }

  out.reserve(out.size() + els.size() * 8 + 8);
  for (std::size_t i = 0; i < n; ++i) {
    for (std::uint32_t k = bounds[i]; k < bounds[i + 1]; ++k) {
      if (els[k].cls == kBase) {
        put_weight(out, els[k].p);
      } else if (els[k].cls == kWord && i >= tail_from) {
        put_weight(out, 0);
      }
    }
  }
  const bool with_c = strength != CollationStrength::kDynamic;
  for (std::int32_t level = 1; level <= 3; ++level) {
    if (level == 3 && !with_c) {
      continue;
    }
    out.push_back('\x01');
    for (const Weight& w : els) {
      if (w.cls == kBase || w.cls == kSec) {
        put_weight(out, level == 1 ? w.s : level == 2 ? w.d : w.c + 1);
      } else if (w.cls == kMark && w.s == level) {
        put_weight(out, 0);
      }
    }
  }
  out.push_back('\x01');
  for (std::size_t i = 0; i < n; ++i) {
    for (std::uint32_t k = bounds[i]; k < bounds[i + 1]; ++k) {
      const Weight& w = els[k];
      if (w.cls == kWord) {
        put_weight(out, static_cast<std::int32_t>(toks[i].pos));
        put_weight(out, w.p);
        put_weight(out, w.s);
        put_weight(out, w.d);
        if (with_c) {
          put_weight(out, w.c);
        }
      }
    }
  }
  out.push_back('\x01');
  if (n > 0 && bounds[n] - bounds[n - 1] == 1 && els[bounds[n - 1]].cls == kMid) {
    put_weight(out, 0);
  }
}

// Printable ASCII whose every character is a plain primary compares by primaries alone.
bool ascii_fast(const LocaleData& L, std::string_view s) noexcept {
  if (!L.ascii_fast) {
    return false;
  }
  for (const char ch : s) {
    const auto b = static_cast<unsigned char>(ch);
    if (b < 0x20u || b > 0x7Eu || L.ascii_p[b] < 0) {
      return false;
    }
  }
  return true;
}

void ascii_key(const LocaleData& L, std::string_view s, CollationStrength strength, std::string& out) {
  out.reserve(s.size() * 8 + 6);
  for (const char ch : s) {
    put_weight(out, L.ascii_p[static_cast<unsigned char>(ch)]);
  }
  const int levels = strength == CollationStrength::kDynamic ? 2 : 3;
  for (int level = 1; level <= levels; ++level) {
    out.push_back('\x01');
    for (std::size_t i = 0; i < s.size(); ++i) {
      out.push_back('\x02');
      out.push_back(level == 3 ? '\x03' : '\x02');
    }
  }
  out.append("\x01\x01", 2);
}

int ascii_compare(const LocaleData& L, std::string_view a, std::string_view b) noexcept {
  const std::size_t n = std::min(a.size(), b.size());
  for (std::size_t i = 0; i < n; ++i) {
    const std::int32_t pa = L.ascii_p[static_cast<unsigned char>(a[i])];
    const std::int32_t pb = L.ascii_p[static_cast<unsigned char>(b[i])];
    if (pa != pb) {
      return pa < pb ? -1 : 1;
    }
  }
  return a.size() == b.size() ? 0 : (a.size() < b.size() ? -1 : 1);
}

}  // namespace

std::string collation_key(std::string_view s, ExcelProfile profile, CollationStrength strength) {
  const LocaleData& L = locale_data(profile);
  std::string out;
  if (ascii_fast(L, s)) {
    ascii_key(L, s, strength, out);
  } else {
    table_key(L, s, strength, out);
  }
  return out;
}

int collate(std::string_view a, std::string_view b, ExcelProfile profile, CollationStrength strength) {
  const LocaleData& L = locale_data(profile);
  if (ascii_fast(L, a) && ascii_fast(L, b)) {
    return ascii_compare(L, a, b);
  }
  return collate_key_compare(collation_key(a, profile, strength), collation_key(b, profile, strength));
}

int collate_key_compare(std::string_view key_a, std::string_view key_b) noexcept {
  const std::size_t n = std::min(key_a.size(), key_b.size());
  const int c = n == 0 ? 0 : std::memcmp(key_a.data(), key_b.data(), n);
  if (c != 0) {
    return c < 0 ? -1 : 1;
  }
  return key_a.size() == key_b.size() ? 0 : (key_a.size() < key_b.size() ? -1 : 1);
}

bool collate_prefix_equal(std::string_view cell, std::string_view criterion, ExcelProfile profile) {
  std::size_t want = 0;
  for (std::size_t i = 0; i < criterion.size(); ++want) {
    std::size_t n = 0;
    decode_utf8_step(criterion, i, &n);
    i += n;
  }
  std::size_t end = 0;
  for (std::size_t k = 0; k < want && end < cell.size(); ++k) {
    std::size_t n = 0;
    decode_utf8_step(cell, end, &n);
    end += n;
  }
  return collate(cell.substr(0, end), criterion, profile, CollationStrength::kCompare) == 0;
}

}  // namespace eval
}  // namespace formulon
