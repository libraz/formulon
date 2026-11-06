// Kana classification for the collation tokenizer: hiragana / katakana
// predicates, the vowel a prolonged-sound mark stands for, voiced-sound
// composition (including half-width and squared forms), and the Windows
// kana-type tertiary weight.

#include <algorithm>
#include <cstdint>
#include <iterator>

#include "eval/collation_internal.h"
#include "eval/jp_kana_table.h"

namespace formulon {
namespace eval {
namespace detail {
namespace {

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

}  // namespace

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

namespace {

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

}  // namespace

// The character composed from the first character of `cp`'s compatibility form
// and a voiced sound mark, or 0 when that does not compose.
std::uint32_t voiced_form(std::uint32_t cp) noexcept {
  if (is_halfwidth_katakana(cp)) {
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

namespace {

bool is_small_kana(std::uint32_t cp) noexcept {
  static constexpr std::uint16_t kSmall[] = {0x3041u, 0x3043u, 0x3045u, 0x3047u, 0x3049u, 0x3063u, 0x3083u, 0x3085u,
                                             0x3087u, 0x308Eu, 0x3095u, 0x3096u, 0x30A1u, 0x30A3u, 0x30A5u, 0x30A7u,
                                             0x30A9u, 0x30C3u, 0x30E3u, 0x30E5u, 0x30E7u, 0x30EEu, 0x30F5u, 0x30F6u};
  if ((cp >= 0x31F0u && cp <= 0x31FFu) || (cp >= 0xFF67u && cp <= 0xFF6Fu)) {
    return true;
  }
  return std::find(std::begin(kSmall), std::end(kSmall), cp) != std::end(kSmall);
}

}  // namespace

// Kana-type weight at the dynamic tertiary level: katakana below hiragana, small below normal.
std::int32_t win_kana_d(std::uint32_t cp) noexcept {
  const bool hiragana = is_hiragana_letter(cp);
  if (!hiragana && !is_katakana_letter(cp)) {
    return 0;
  }
  return (hiragana ? 2 : 0) + (is_small_kana(cp) ? 0 : 1);
}

}  // namespace detail
}  // namespace eval
}  // namespace formulon
