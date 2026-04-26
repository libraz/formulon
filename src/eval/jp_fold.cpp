// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of `fold_jp_text`. See `jp_fold.h` for the full mapping
// contract; this file performs a single forward pass over `input`, decoding
// UTF-8 codepoint-by-codepoint and re-encoding each folded result.

#include "eval/jp_fold.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>

namespace formulon {
namespace eval {
namespace {

// Decodes a single UTF-8 codepoint at `text[i]`. Mirrors the lenient
// broken-UTF-8 handling used elsewhere in the engine: a malformed leading
// byte returns U+FFFD with `*out_bytes == 1`, and a bad continuation likewise
// returns U+FFFD over a single-byte advance.
std::uint32_t decode_utf8(std::string_view text, std::size_t i, std::size_t* out_bytes) {
  const auto lead = static_cast<unsigned char>(text[i]);
  if ((lead & 0x80u) == 0x00u) {
    *out_bytes = 1;
    return lead;
  }
  std::size_t need = 0;
  std::uint32_t value = 0;
  if ((lead & 0xE0u) == 0xC0u) {
    need = 1;
    value = lead & 0x1Fu;
  } else if ((lead & 0xF0u) == 0xE0u) {
    need = 2;
    value = lead & 0x0Fu;
  } else if ((lead & 0xF8u) == 0xF0u) {
    need = 3;
    value = lead & 0x07u;
  } else {
    *out_bytes = 1;
    return 0xFFFDu;
  }
  if (i + need >= text.size()) {
    *out_bytes = 1;
    return 0xFFFDu;
  }
  for (std::size_t k = 0; k < need; ++k) {
    const auto ck = static_cast<unsigned char>(text[i + 1 + k]);
    if ((ck & 0xC0u) != 0x80u) {
      *out_bytes = 1;
      return 0xFFFDu;
    }
    value = (value << 6) | (ck & 0x3Fu);
  }
  *out_bytes = need + 1;
  return value;
}

// Appends `cp` to `out` as a UTF-8 byte sequence (1..4 bytes). Codepoints
// above U+10FFFF are silently clamped to U+FFFD.
void encode_utf8(std::uint32_t cp, std::string* out) {
  if (cp <= 0x7Fu) {
    out->push_back(static_cast<char>(cp));
    return;
  }
  if (cp <= 0x7FFu) {
    out->push_back(static_cast<char>(0xC0u | (cp >> 6)));
    out->push_back(static_cast<char>(0x80u | (cp & 0x3Fu)));
    return;
  }
  if (cp <= 0xFFFFu) {
    out->push_back(static_cast<char>(0xE0u | (cp >> 12)));
    out->push_back(static_cast<char>(0x80u | ((cp >> 6) & 0x3Fu)));
    out->push_back(static_cast<char>(0x80u | (cp & 0x3Fu)));
    return;
  }
  if (cp <= 0x10FFFFu) {
    out->push_back(static_cast<char>(0xF0u | (cp >> 18)));
    out->push_back(static_cast<char>(0x80u | ((cp >> 12) & 0x3Fu)));
    out->push_back(static_cast<char>(0x80u | ((cp >> 6) & 0x3Fu)));
    out->push_back(static_cast<char>(0x80u | (cp & 0x3Fu)));
    return;
  }
  // Out-of-range: emit replacement character.
  out->push_back(static_cast<char>(0xEFu));
  out->push_back(static_cast<char>(0xBFu));
  out->push_back(static_cast<char>(0xBDu));
}

// Maps half-width katakana U+FF61..U+FF9D to their full-width equivalents.
// Index = cp - 0xFF61u; entries cover the 61 half-width characters in this
// block (excluding the trailing voicing marks U+FF9E / U+FF9F which require
// composition with a preceding base — handled separately).
constexpr std::uint32_t kHalfToFullKana[] = {
    0x3002u,  // FF61 ｡  -> 。
    0x300Cu,  // FF62 ｢  -> 「
    0x300Du,  // FF63 ｣  -> 」
    0x3001u,  // FF64 ､  -> 、
    0x30FBu,  // FF65 ･  -> ・
    0x30F2u,  // FF66 ｦ  -> ヲ
    0x30A1u,  // FF67 ｧ  -> ァ
    0x30A3u,  // FF68 ｨ  -> ィ
    0x30A5u,  // FF69 ｩ  -> ゥ
    0x30A7u,  // FF6A ｪ  -> ェ
    0x30A9u,  // FF6B ｫ  -> ォ
    0x30E3u,  // FF6C ｬ  -> ャ
    0x30E5u,  // FF6D ｭ  -> ュ
    0x30E7u,  // FF6E ｮ  -> ョ
    0x30C3u,  // FF6F ｯ  -> ッ
    0x30FCu,  // FF70 ｰ  -> ー
    0x30A2u,  // FF71 ｱ  -> ア
    0x30A4u,  // FF72 ｲ  -> イ
    0x30A6u,  // FF73 ｳ  -> ウ
    0x30A8u,  // FF74 ｴ  -> エ
    0x30AAu,  // FF75 ｵ  -> オ
    0x30ABu,  // FF76 ｶ  -> カ
    0x30ADu,  // FF77 ｷ  -> キ
    0x30AFu,  // FF78 ｸ  -> ク
    0x30B1u,  // FF79 ｹ  -> ケ
    0x30B3u,  // FF7A ｺ  -> コ
    0x30B5u,  // FF7B ｻ  -> サ
    0x30B7u,  // FF7C ｼ  -> シ
    0x30B9u,  // FF7D ｽ  -> ス
    0x30BBu,  // FF7E ｾ  -> セ
    0x30BDu,  // FF7F ｿ  -> ソ
    0x30BFu,  // FF80 ﾀ  -> タ
    0x30C1u,  // FF81 ﾁ  -> チ
    0x30C4u,  // FF82 ﾂ  -> ツ
    0x30C6u,  // FF83 ﾃ  -> テ
    0x30C8u,  // FF84 ﾄ  -> ト
    0x30CAu,  // FF85 ﾅ  -> ナ
    0x30CBu,  // FF86 ﾆ  -> ニ
    0x30CCu,  // FF87 ﾇ  -> ヌ
    0x30CDu,  // FF88 ﾈ  -> ネ
    0x30CEu,  // FF89 ﾉ  -> ノ
    0x30CFu,  // FF8A ﾊ  -> ハ
    0x30D2u,  // FF8B ﾋ  -> ヒ
    0x30D5u,  // FF8C ﾌ  -> フ
    0x30D8u,  // FF8D ﾍ  -> ヘ
    0x30DBu,  // FF8E ﾎ  -> ホ
    0x30DEu,  // FF8F ﾏ  -> マ
    0x30DFu,  // FF90 ﾐ  -> ミ
    0x30E0u,  // FF91 ﾑ  -> ム
    0x30E1u,  // FF92 ﾒ  -> メ
    0x30E2u,  // FF93 ﾓ  -> モ
    0x30E4u,  // FF94 ﾔ  -> ヤ
    0x30E6u,  // FF95 ﾕ  -> ユ
    0x30E8u,  // FF96 ﾖ  -> ヨ
    0x30E9u,  // FF97 ﾗ  -> ラ
    0x30EAu,  // FF98 ﾘ  -> リ
    0x30EBu,  // FF99 ﾙ  -> ル
    0x30ECu,  // FF9A ﾚ  -> レ
    0x30EDu,  // FF9B ﾛ  -> ロ
    0x30EFu,  // FF9C ﾜ  -> ワ
    0x30F3u,  // FF9D ﾝ  -> ン
};

// Returns the voiced (dakuten) form of full-width katakana `full_cp`, or 0
// if no voiced form exists. The Unicode katakana block places voiced forms
// at base+1 across the ka/sa/ta and ha rows (e.g. カ U+30AB + 1 -> ガ
// U+30AC, ハ U+30CF + 1 -> バ U+30D0). Special-case: ウ U+30A6 voices to
// ヴ U+30F4 — Unicode places this outside the +1 pattern.
std::uint32_t voiced_form(std::uint32_t full_cp) {
  if (full_cp == 0x30A6u) {
    // ウ + ﾞ -> ヴ
    return 0x30F4u;
  }
  // ka/sa/ta rows: U+30AB..U+30C9 (カ..ド). Only the unvoiced base codepoints
  // (even offsets within ka/sa/ta) accept voicing.
  if (full_cp >= 0x30ABu && full_cp <= 0x30C9u && ((full_cp - 0x30ABu) % 2u) == 0u) {
    return full_cp + 1u;
  }
  // ha row: U+30CF..U+30DD (ハ..ポ). Voiced/semi-voiced forms occupy the
  // +1/+2 offsets, so accept only offsets divisible by 3 from the base.
  if (full_cp >= 0x30CFu && full_cp <= 0x30DDu && ((full_cp - 0x30CFu) % 3u) == 0u) {
    return full_cp + 1u;
  }
  return 0u;
}

// Returns the semi-voiced (handakuten) form of full-width katakana
// `full_cp`, or 0 if no semi-voiced form exists. Only the ハ row supports
// semi-voicing; the codepoint sits at base+2 within the same row.
std::uint32_t semi_voiced_form(std::uint32_t full_cp) {
  if (full_cp >= 0x30CFu && full_cp <= 0x30DDu && ((full_cp - 0x30CFu) % 3u) == 0u) {
    return full_cp + 2u;
  }
  return 0u;
}

// Returns true when `full_cp` is the full-width katakana mapping of a
// half-width base (i.e. it appeared in `kHalfToFullKana`). Only those bases
// can absorb a following ﾞ / ﾟ; punctuation entries (。「」、・) and the
// long-mark ー must not absorb voicing marks.
bool is_voicing_base(std::uint32_t full_cp) {
  // Voicing eligibility = `voiced_form` or `semi_voiced_form` returns nonzero.
  return voiced_form(full_cp) != 0u || semi_voiced_form(full_cp) != 0u;
}

}  // namespace

std::string fold_jp_text(std::string_view input) {
  std::string out;
  out.reserve(input.size());
  std::size_t i = 0;
  while (i < input.size()) {
    std::size_t n = 0;
    std::uint32_t cp = decode_utf8(input, i, &n);

    // Hiragana U+3041..U+3096 -> Katakana (+0x60).
    if (cp >= 0x3041u && cp <= 0x3096u) {
      cp += 0x60u;
      encode_utf8(cp, &out);
      i += n;
      continue;
    }

    // Full-width ASCII U+FF01..U+FF5E -> half-width ASCII (-0xFEE0).
    if (cp >= 0xFF01u && cp <= 0xFF5Eu) {
      encode_utf8(cp - 0xFEE0u, &out);
      i += n;
      continue;
    }

    // Ideographic space U+3000 -> ASCII space.
    if (cp == 0x3000u) {
      out.push_back(' ');
      i += n;
      continue;
    }

    // Half-width katakana U+FF61..U+FF9D: map to full-width, optionally
    // composing a trailing ﾞ / ﾟ from the next codepoint.
    if (cp >= 0xFF61u && cp <= 0xFF9Du) {
      std::uint32_t base = kHalfToFullKana[cp - 0xFF61u];
      if (is_voicing_base(base) && i + n < input.size()) {
        std::size_t n2 = 0;
        const std::uint32_t next = decode_utf8(input, i + n, &n2);
        if (next == 0xFF9Eu) {
          if (const std::uint32_t v = voiced_form(base); v != 0u) {
            encode_utf8(v, &out);
            i += n + n2;
            continue;
          }
        } else if (next == 0xFF9Fu) {
          if (const std::uint32_t s = semi_voiced_form(base); s != 0u) {
            encode_utf8(s, &out);
            i += n + n2;
            continue;
          }
        }
      }
      encode_utf8(base, &out);
      i += n;
      continue;
    }

    // Standalone half-width voicing marks (no composable base preceded
    // them). Mac Excel ja-JP normalises these to the spacing combining
    // marks U+309B (゛) and U+309C (゜).
    if (cp == 0xFF9Eu) {
      encode_utf8(0x309Bu, &out);
      i += n;
      continue;
    }
    if (cp == 0xFF9Fu) {
      encode_utf8(0x309Cu, &out);
      i += n;
      continue;
    }

    // Pass-through for every other codepoint.
    encode_utf8(cp, &out);
    i += n;
  }
  return out;
}

}  // namespace eval
}  // namespace formulon
