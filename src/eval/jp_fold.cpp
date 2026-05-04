// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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
#include "eval/text_ops.h"
#include "utils/strings.h"

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
// half-width base (i.e. it would emerge from `half_to_full_katakana_base`).
// Only those bases can absorb a following ﾞ / ﾟ; punctuation entries
// (。「」、・) and the long-mark ー must not absorb voicing marks.
bool is_voicing_base(std::uint32_t full_cp) {
  // Voicing eligibility = `voiced_form` or `semi_voiced_form` returns nonzero.
  return voiced_form(full_cp) != 0u || semi_voiced_form(full_cp) != 0u;
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
    // Lookup callers (MATCH / VLOOKUP / HLOOKUP / XLOOKUP / XMATCH) pass
    // `fold_fullwidth_digits = false` to keep U+FF10..U+FF19 unfolded,
    // matching the Mac Excel asymmetry documented in the header.
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
    // composing a trailing ﾞ / ﾟ from the next codepoint. D-function
    // header callers pass `fold_halfwidth_kana = false`, in which case
    // we suppress the entire branch — including the voicing-mark
    // composition — so that `ｶﾞ` stays as the two-codepoint sequence
    // FF76 FF9E and does not compose to ガ. See header for the Mac Excel
    // empirical asymmetry that motivates this.
    if (fold_halfwidth_kana && cp >= 0xFF61u && cp <= 0xFF9Du) {
      std::uint32_t base = half_to_full_kana_or_punct(cp);
      if (is_voicing_base(base) && i + n < input.size()) {
        std::size_t n2 = 0;
        const std::uint32_t next = decode_utf8_step(input, i + n, &n2);
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

std::string fold_and_lower(std::string_view input, bool fold_fullwidth_digits) {
  // Two-pass: kana fold first (so ｶﾞ -> ガ, Ａ -> a, etc.), then
  // ASCII lowercase. The fold pass already yields half-width ASCII
  // letters where applicable, so the lowercase pass is purely a byte-
  // wise A..Z -> a..z rewrite. Composing the two helpers in one
  // function lets every lookup-family caller share a single allocation
  // path and keeps the kana-fold + lowercase invariant trivially in
  // sync across `lookups/classic.cpp` and `lookups/xlookup.cpp`.
  return strings::to_ascii_lower(fold_jp_text(input, fold_fullwidth_digits, /*fold_halfwidth_kana=*/true));
}

}  // namespace eval
}  // namespace formulon
