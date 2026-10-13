//
// Implementation of `fold_jp_text`. See `jp_fold.h` for the full mapping
// contract; this file performs a single forward pass over `input`, decoding
// UTF-8 codepoint-by-codepoint and re-encoding each folded result.

#include "eval/jp_fold.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>

#include "eval/jp_kana_table.h"
#include "utils/strings.h"
#include "utils/text_ops.h"

namespace formulon {
namespace eval {
namespace {

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

// Maps half-width katakana / punctuation U+FF61..U+FF9D to their
// full-width equivalents using the shared `jp_kana_table` lookups.
// Returns 0 for codepoints outside the supported range.
std::uint32_t half_to_full_kana_or_punct(std::uint32_t cp) noexcept {
  if (cp >= 0xFF66u && cp <= 0xFF9Du) {
    return half_to_full_katakana_base(cp);
  }
  return half_to_full_punctuation(cp);
}

// Returns true when `full_cp` is the full-width katakana mapping of a
// half-width base (i.e. it would emerge from `half_to_full_katakana_base`).
// Only those bases can absorb a following ﾞ / ﾟ; punctuation entries
// (。「」、・) and the long-mark ー must not absorb voicing marks.
bool is_voicing_base(std::uint32_t full_cp) {
  return voiced_full_from_base(full_cp) != 0u || semi_voiced_full_from_base(full_cp) != 0u;
}

}  // namespace

std::string fold_jp_text(std::string_view input, bool fold_fullwidth_digits, bool fold_halfwidth_kana) {
  std::string out;
  out.reserve(input.size());
  std::size_t i = 0;
  while (i < input.size()) {
    std::size_t n = 0;
    std::uint32_t cp = decode_utf8_step(input, i, &n);

    // Hiragana U+3041..U+3096 -> Katakana (+0x60).
    if (cp >= 0x3041u && cp <= 0x3096u) {
      cp += 0x60u;
      encode_utf8(cp, &out);
      i += n;
      continue;
    }

    // Full-width ASCII U+FF01..U+FF5E -> half-width ASCII (-0xFEE0).
    // `fold_fullwidth_digits = false` keeps U+FF10..U+FF19 unfolded.
    if (cp >= 0xFF01u && cp <= 0xFF5Eu) {
      const bool is_fullwidth_digit = cp >= 0xFF10u && cp <= 0xFF19u;
      if (is_fullwidth_digit && !fold_fullwidth_digits) {
        encode_utf8(cp, &out);
      } else {
        encode_utf8(cp - 0xFEE0u, &out);
      }
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
    // composing a trailing ﾞ / ﾟ from the next codepoint. With
    // `fold_halfwidth_kana = false` the whole branch is skipped, so `ｶﾞ`
    // stays FF76 FF9E (see the header for the caller that needs it).
    if (fold_halfwidth_kana && cp >= 0xFF61u && cp <= 0xFF9Du) {
      std::uint32_t base = half_to_full_kana_or_punct(cp);
      if (is_voicing_base(base) && i + n < input.size()) {
        std::size_t n2 = 0;
        const std::uint32_t next = decode_utf8_step(input, i + n, &n2);
        if (next == 0xFF9Eu) {
          if (const std::uint32_t v = voiced_full_from_base(base); v != 0u) {
            encode_utf8(v, &out);
            i += n + n2;
            continue;
          }
        } else if (next == 0xFF9Fu) {
          if (const std::uint32_t s = semi_voiced_full_from_base(base); s != 0u) {
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
    // marks U+309B (゛) and U+309C (゜) — but only when half-width kana
    // folding is enabled. The D-function header path passes
    // `fold_halfwidth_kana = false`, in which case the marks pass through
    // unchanged so the comparison sees the raw FF9E / FF9F bytes.
    if (fold_halfwidth_kana && cp == 0xFF9Eu) {
      encode_utf8(0x309Bu, &out);
      i += n;
      continue;
    }
    if (fold_halfwidth_kana && cp == 0xFF9Fu) {
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

std::string compose_jp_halfwidth_voicing(std::string_view input) {
  std::string out;
  out.reserve(input.size());
  std::size_t i = 0;
  while (i < input.size()) {
    std::size_t n = 0;
    const std::uint32_t cp = decode_utf8_step(input, i, &n);

    // Only a half-width katakana base can absorb a following ﾞ / ﾟ.
    if (cp >= 0xFF66u && cp <= 0xFF9Du && i + n < input.size()) {
      const std::uint32_t base = half_to_full_kana_or_punct(cp);
      if (is_voicing_base(base)) {
        std::size_t n2 = 0;
        const std::uint32_t next = decode_utf8_step(input, i + n, &n2);
        if (next == 0xFF9Eu) {
          if (const std::uint32_t v = voiced_full_from_base(base); v != 0u) {
            encode_utf8(v, &out);
            i += n + n2;
            continue;
          }
        } else if (next == 0xFF9Fu) {
          if (const std::uint32_t s = semi_voiced_full_from_base(base); s != 0u) {
            encode_utf8(s, &out);
            i += n + n2;
            continue;
          }
        }
      }
    }

    // Every other codepoint (including a plain half-width base with no
    // trailing voicing mark) passes through unchanged.
    encode_utf8(cp, &out);
    i += n;
  }
  return out;
}

std::string fold_and_lower(std::string_view input, bool fold_fullwidth_digits) {
  return strings::to_ascii_lower_owned(fold_jp_text(input, fold_fullwidth_digits, /*fold_halfwidth_kana=*/true));
}

}  // namespace eval
}  // namespace formulon
