//
// Implementation of the shared half-to-full katakana base lookup. See
// `jp_kana_table.h` for the contract.

#include "eval/jp_kana_table.h"

#include <cstdint>

namespace formulon {
namespace eval {
namespace {

// Half-width katakana base block U+FF66..U+FF9D mapped to its full-width
// counterpart. Indexed by `cp - 0xFF66`. Verified bit-for-bit against the
// previous duplicate tables in `jp_fold.cpp` (entries 5..60 of the
// original `kHalfToFullKana`) and `text_width.cpp` (`kKatakanaHalfToFull`).
constexpr std::uint32_t kHalfToFullBase[] = {
    0x30F2,  // U+FF66 ｦ -> ヲ
    0x30A1,  // U+FF67 ｧ -> ァ
    0x30A3,  // U+FF68 ｨ -> ィ
    0x30A5,  // U+FF69 ｩ -> ゥ
    0x30A7,  // U+FF6A ｪ -> ェ
    0x30A9,  // U+FF6B ｫ -> ォ
    0x30E3,  // U+FF6C ｬ -> ャ
    0x30E5,  // U+FF6D ｭ -> ュ
    0x30E7,  // U+FF6E ｮ -> ョ
    0x30C3,  // U+FF6F ｯ -> ッ
    0x30FC,  // U+FF70 ｰ -> ー
    0x30A2,  // U+FF71 ｱ -> ア
    0x30A4,  // U+FF72 ｲ -> イ
    0x30A6,  // U+FF73 ｳ -> ウ
    0x30A8,  // U+FF74 ｴ -> エ
    0x30AA,  // U+FF75 ｵ -> オ
    0x30AB,  // U+FF76 ｶ -> カ
    0x30AD,  // U+FF77 ｷ -> キ
    0x30AF,  // U+FF78 ｸ -> ク
    0x30B1,  // U+FF79 ｹ -> ケ
    0x30B3,  // U+FF7A ｺ -> コ
    0x30B5,  // U+FF7B ｻ -> サ
    0x30B7,  // U+FF7C ｼ -> シ
    0x30B9,  // U+FF7D ｽ -> ス
    0x30BB,  // U+FF7E ｾ -> セ
    0x30BD,  // U+FF7F ｿ -> ソ
    0x30BF,  // U+FF80 ﾀ -> タ
    0x30C1,  // U+FF81 ﾁ -> チ
    0x30C4,  // U+FF82 ﾂ -> ツ
    0x30C6,  // U+FF83 ﾃ -> テ
    0x30C8,  // U+FF84 ﾄ -> ト
    0x30CA,  // U+FF85 ﾅ -> ナ
    0x30CB,  // U+FF86 ﾆ -> ニ
    0x30CC,  // U+FF87 ﾇ -> ヌ
    0x30CD,  // U+FF88 ﾈ -> ネ
    0x30CE,  // U+FF89 ﾉ -> ノ
    0x30CF,  // U+FF8A ﾊ -> ハ
    0x30D2,  // U+FF8B ﾋ -> ヒ
    0x30D5,  // U+FF8C ﾌ -> フ
    0x30D8,  // U+FF8D ﾍ -> ヘ
    0x30DB,  // U+FF8E ﾎ -> ホ
    0x30DE,  // U+FF8F ﾏ -> マ
    0x30DF,  // U+FF90 ﾐ -> ミ
    0x30E0,  // U+FF91 ﾑ -> ム
    0x30E1,  // U+FF92 ﾒ -> メ
    0x30E2,  // U+FF93 ﾓ -> モ
    0x30E4,  // U+FF94 ﾔ -> ヤ
    0x30E6,  // U+FF95 ﾕ -> ユ
    0x30E8,  // U+FF96 ﾖ -> ヨ
    0x30E9,  // U+FF97 ﾗ -> ラ
    0x30EA,  // U+FF98 ﾘ -> リ
    0x30EB,  // U+FF99 ﾙ -> ル
    0x30EC,  // U+FF9A ﾚ -> レ
    0x30ED,  // U+FF9B ﾛ -> ロ
    0x30EF,  // U+FF9C ﾜ -> ワ
    0x30F3,  // U+FF9D ﾝ -> ン
};

// Half-width punctuation U+FF61..U+FF65 mapped to its full-width
// counterpart. The original `jp_fold.cpp` table covered this range
// inline; `text_width.cpp` handled the same five codepoints with a
// dedicated switch. Centralised here so both paths share one source.
constexpr std::uint32_t kHalfToFullPunctuation[] = {
    0x3002,  // U+FF61 ｡ -> 。
    0x300C,  // U+FF62 ｢ -> 「
    0x300D,  // U+FF63 ｣ -> 」
    0x3001,  // U+FF64 ､ -> 、
    0x30FB,  // U+FF65 ･ -> ・
};

}  // namespace

std::uint32_t half_to_full_katakana_base(std::uint32_t cp) noexcept {
  if (cp < 0xFF66u || cp > 0xFF9Du) {
    return 0u;
  }
  return kHalfToFullBase[cp - 0xFF66u];
}

std::uint32_t half_to_full_punctuation(std::uint32_t cp) noexcept {
  if (cp < 0xFF61u || cp > 0xFF65u) {
    return 0u;
  }
  return kHalfToFullPunctuation[cp - 0xFF61u];
}

std::uint32_t voiced_full_from_base(std::uint32_t full_base) noexcept {
  switch (full_base) {
    case 0x30A6u:
      return 0x30F4u;  // ウ -> ヴ
    case 0x30ABu:
    case 0x30ADu:
    case 0x30AFu:
    case 0x30B1u:
    case 0x30B3u:
    case 0x30B5u:
    case 0x30B7u:
    case 0x30B9u:
    case 0x30BBu:
    case 0x30BDu:
    case 0x30BFu:
    case 0x30C1u:
    case 0x30C4u:
    case 0x30C6u:
    case 0x30C8u:
    case 0x30CFu:
    case 0x30D2u:
    case 0x30D5u:
    case 0x30D8u:
    case 0x30DBu:
      return full_base + 1u;
    default:
      return 0u;
  }
}

std::uint32_t semi_voiced_full_from_base(std::uint32_t full_base) noexcept {
  switch (full_base) {
    case 0x30CFu:
    case 0x30D2u:
    case 0x30D5u:
    case 0x30D8u:
    case 0x30DBu:
      return full_base + 2u;
    default:
      return 0u;
  }
}

}  // namespace eval
}  // namespace formulon
