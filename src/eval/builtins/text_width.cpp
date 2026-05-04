// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of the width-conversion text built-ins: ASC (full-width
// to half-width) and JIS / DBCS (half-width to full-width). Both functions
// coerce their single argument through `coerce_to_text`, walk the UTF-8
// byte sequence codepoint by codepoint, and emit the translated text into
// the caller's arena.
//
// ASC maps:
//   * U+FF01..U+FF5E (full-width ASCII) -> U+0021..U+007E via `cp - 0xFEE0`.
//   * U+3000 (ideographic space) -> U+0020.
//   * U+30A1..U+30FC (full-width katakana) -> half-width equivalents from
//     `kKatakanaFullToHalf`. Voiced / semi-voiced forms emit two codepoints
//     (a base katakana plus U+FF9E dakuten or U+FF9F handakuten); the
//     remaining katakana emit a single codepoint. Archaic katakana that
//     have no half-width mapping (ヮ, ヰ, ヱ, ヵ, ヶ, ヷ, ヸ, ヹ, ヺ) pass
//     through unchanged.
//   * All other codepoints (hiragana, kanji, Latin letters with accents,
//     emoji, ...) pass through verbatim.
//
// JIS (aliased as DBCS) is the inverse:
//   * U+0020 (ASCII space) -> U+3000.
//   * U+0021..U+007E (ASCII printable) -> U+FF01..U+FF5E via `cp + 0xFEE0`.
//   * U+FF66..U+FF9D (half-width katakana) -> the corresponding full-width
//     base from `eval/jp_kana_table.h`. When the next codepoint is U+FF9E or
//     U+FF9F and the base accepts a dakuten / handakuten, the two are
//     combined into a single voiced / semi-voiced full-width katakana.
//   * Specific half-width punctuation (U+FF61..U+FF65) -> full-width
//     equivalents (U+3002, U+300C, U+300D, U+3001, U+30FB).
//   * Lone U+FF9E / U+FF9F (not preceded by a compatible base) -> the
//     spacing voiced / semi-voiced sound marks (U+309B / U+309C).
//   * All other codepoints pass through verbatim.
//
// JIS and DBCS share one implementation - they differ only in registered
// name.

#include "eval/builtins/text_width.h"

#include <array>
#include <cstdint>
#include <string>
#include <string_view>

#include "eval/coerce.h"
#include "eval/function_registry.h"
#include "eval/jp_kana_table.h"
#include "eval/text_ops.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// --- Full-width katakana -> half-width mapping --------------------------
//
// Indexed by `codepoint - 0x30A1`. `second == 0` means "emit `first`
// alone"; a non-zero `second` means "emit `first` then `second`" (used for
// voiced / semi-voiced katakana that decompose into a base + combining
// mark). A zero `first` means "no mapping; pass the original codepoint
// through unchanged" (used for archaic katakana: ヮ, ヰ, ヱ, ヵ, ヶ, ヷ,
// ヸ, ヹ, ヺ). Table is 92 entries covering U+30A1..U+30FC.
struct KatakanaMap {
  std::uint16_t first;
  std::uint16_t second;
};

constexpr std::uint16_t kDakuten = 0xFF9E;
constexpr std::uint16_t kHandakuten = 0xFF9F;

constexpr std::array<KatakanaMap, 92> kKatakanaFullToHalf = {{
    {0xFF67, 0},            // U+30A1 ァ
    {0xFF71, 0},            // U+30A2 ア
    {0xFF68, 0},            // U+30A3 ィ
    {0xFF72, 0},            // U+30A4 イ
    {0xFF69, 0},            // U+30A5 ゥ
    {0xFF73, 0},            // U+30A6 ウ
    {0xFF6A, 0},            // U+30A7 ェ
    {0xFF74, 0},            // U+30A8 エ
    {0xFF6B, 0},            // U+30A9 ォ
    {0xFF75, 0},            // U+30AA オ
    {0xFF76, 0},            // U+30AB カ
    {0xFF76, kDakuten},     // U+30AC ガ
    {0xFF77, 0},            // U+30AD キ
    {0xFF77, kDakuten},     // U+30AE ギ
    {0xFF78, 0},            // U+30AF ク
    {0xFF78, kDakuten},     // U+30B0 グ
    {0xFF79, 0},            // U+30B1 ケ
    {0xFF79, kDakuten},     // U+30B2 ゲ
    {0xFF7A, 0},            // U+30B3 コ
    {0xFF7A, kDakuten},     // U+30B4 ゴ
    {0xFF7B, 0},            // U+30B5 サ
    {0xFF7B, kDakuten},     // U+30B6 ザ
    {0xFF7C, 0},            // U+30B7 シ
    {0xFF7C, kDakuten},     // U+30B8 ジ
    {0xFF7D, 0},            // U+30B9 ス
    {0xFF7D, kDakuten},     // U+30BA ズ
    {0xFF7E, 0},            // U+30BB セ
    {0xFF7E, kDakuten},     // U+30BC ゼ
    {0xFF7F, 0},            // U+30BD ソ
    {0xFF7F, kDakuten},     // U+30BE ゾ
    {0xFF80, 0},            // U+30BF タ
    {0xFF80, kDakuten},     // U+30C0 ダ
    {0xFF81, 0},            // U+30C1 チ
    {0xFF81, kDakuten},     // U+30C2 ヂ
    {0xFF6F, 0},            // U+30C3 ッ (small tsu)
    {0xFF82, 0},            // U+30C4 ツ
    {0xFF82, kDakuten},     // U+30C5 ヅ
    {0xFF83, 0},            // U+30C6 テ
    {0xFF83, kDakuten},     // U+30C7 デ
    {0xFF84, 0},            // U+30C8 ト
    {0xFF84, kDakuten},     // U+30C9 ド
    {0xFF85, 0},            // U+30CA ナ
    {0xFF86, 0},            // U+30CB ニ
    {0xFF87, 0},            // U+30CC ヌ
    {0xFF88, 0},            // U+30CD ネ
    {0xFF89, 0},            // U+30CE ノ
    {0xFF8A, 0},            // U+30CF ハ
    {0xFF8A, kDakuten},     // U+30D0 バ
    {0xFF8A, kHandakuten},  // U+30D1 パ
    {0xFF8B, 0},            // U+30D2 ヒ
    {0xFF8B, kDakuten},     // U+30D3 ビ
    {0xFF8B, kHandakuten},  // U+30D4 ピ
    {0xFF8C, 0},            // U+30D5 フ
    {0xFF8C, kDakuten},     // U+30D6 ブ
    {0xFF8C, kHandakuten},  // U+30D7 プ
    {0xFF8D, 0},            // U+30D8 ヘ
    {0xFF8D, kDakuten},     // U+30D9 ベ
    {0xFF8D, kHandakuten},  // U+30DA ペ
    {0xFF8E, 0},            // U+30DB ホ
    {0xFF8E, kDakuten},     // U+30DC ボ
    {0xFF8E, kHandakuten},  // U+30DD ポ
    {0xFF8F, 0},            // U+30DE マ
    {0xFF90, 0},            // U+30DF ミ
    {0xFF91, 0},            // U+30E0 ム
    {0xFF92, 0},            // U+30E1 メ
    {0xFF93, 0},            // U+30E2 モ
    {0xFF6C, 0},            // U+30E3 ャ
    {0xFF94, 0},            // U+30E4 ヤ
    {0xFF6D, 0},            // U+30E5 ュ
    {0xFF95, 0},            // U+30E6 ユ
    {0xFF6E, 0},            // U+30E7 ョ
    {0xFF96, 0},            // U+30E8 ヨ
    {0xFF97, 0},            // U+30E9 ラ
    {0xFF98, 0},            // U+30EA リ
    {0xFF99, 0},            // U+30EB ル
    {0xFF9A, 0},            // U+30EC レ
    {0xFF9B, 0},            // U+30ED ロ
    {0, 0},                 // U+30EE ヮ (archaic, no half-width)
    {0xFF9C, 0},            // U+30EF ワ
    {0, 0},                 // U+30F0 ヰ (archaic)
    {0, 0},                 // U+30F1 ヱ (archaic)
    {0xFF66, 0},            // U+30F2 ヲ
    {0xFF9D, 0},            // U+30F3 ン
    {0, 0},                 // U+30F4 ヴ (Mac Excel passes through unchanged)
    {0, 0},                 // U+30F5 ヵ (archaic)
    {0, 0},                 // U+30F6 ヶ (archaic)
    {0, 0},                 // U+30F7 ヷ (archaic)
    {0, 0},                 // U+30F8 ヸ (archaic)
    {0, 0},                 // U+30F9 ヹ (archaic)
    {0, 0},                 // U+30FA ヺ (archaic)
    {0xFF65, 0},            // U+30FB ・
    {0xFF70, 0},            // U+30FC ー
}};

// Half-width katakana -> full-width base mapping lives in the shared
// `eval/jp_kana_table.h` so `text_width` (ASC / JIS / DBCS) and
// `jp_fold` (lookup-key folding) draw from the same source. Callers
// here still merge a trailing U+FF9E / U+FF9F into the voiced /
// semi-voiced form via `voiced_full_from_base` /
// `semi_voiced_full_from_base` defined below.

// Returns the voiced full-width katakana for `full_base`, or 0 if no
// such voiced form exists (callers should then emit the base codepoint
// plus U+309B spacing dakuten).
std::uint16_t voiced_full_from_base(std::uint16_t full_base) noexcept {
  switch (full_base) {
    case 0x30A6:  // ウ -> ヴ
      return 0x30F4;
    case 0x30AB:  // カ -> ガ
      return 0x30AC;
    case 0x30AD:  // キ -> ギ
      return 0x30AE;
    case 0x30AF:  // ク -> グ
      return 0x30B0;
    case 0x30B1:  // ケ -> ゲ
      return 0x30B2;
    case 0x30B3:  // コ -> ゴ
      return 0x30B4;
    case 0x30B5:  // サ -> ザ
      return 0x30B6;
    case 0x30B7:  // シ -> ジ
      return 0x30B8;
    case 0x30B9:  // ス -> ズ
      return 0x30BA;
    case 0x30BB:  // セ -> ゼ
      return 0x30BC;
    case 0x30BD:  // ソ -> ゾ
      return 0x30BE;
    case 0x30BF:  // タ -> ダ
      return 0x30C0;
    case 0x30C1:  // チ -> ヂ
      return 0x30C2;
    case 0x30C4:  // ツ -> ヅ
      return 0x30C5;
    case 0x30C6:  // テ -> デ
      return 0x30C7;
    case 0x30C8:  // ト -> ド
      return 0x30C9;
    case 0x30CF:  // ハ -> バ
      return 0x30D0;
    case 0x30D2:  // ヒ -> ビ
      return 0x30D3;
    case 0x30D5:  // フ -> ブ
      return 0x30D6;
    case 0x30D8:  // ヘ -> ベ
      return 0x30D9;
    case 0x30DB:  // ホ -> ボ
      return 0x30DC;
    default:
      return 0;
  }
}

// Returns the semi-voiced full-width katakana for `full_base`, or 0 if no
// such form exists (callers should then emit the base plus U+309C spacing
// handakuten). Only the ハ-row bases accept handakuten.
std::uint16_t semi_voiced_full_from_base(std::uint16_t full_base) noexcept {
  switch (full_base) {
    case 0x30CF:  // ハ -> パ
      return 0x30D1;
    case 0x30D2:  // ヒ -> ピ
      return 0x30D4;
    case 0x30D5:  // フ -> プ
      return 0x30D7;
    case 0x30D8:  // ヘ -> ペ
      return 0x30DA;
    case 0x30DB:  // ホ -> ポ
      return 0x30DD;
    default:
      return 0;
  }
}

// Appends the UTF-8 encoding of `cp` to `out`. Thin wrapper over
// `encode_utf8_codepoint`; a zero codepoint or an invalid codepoint
// yields no bytes.
void append_codepoint(std::string& out, std::uint32_t cp) {
  if (cp == 0) {
    return;
  }
  out += encode_utf8_codepoint(cp);
}

// ASC(text) - full-width to half-width conversion.
Value Asc(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  const std::string& src = text.value();
  std::string out;
  out.reserve(src.size());
  std::string_view view = src;
  while (!view.empty()) {
    const Utf8DecodeResult d = decode_first_utf8_codepoint(view);
    if (!d.valid) {
      // Malformed leading byte: emit one byte verbatim and advance.
      out.push_back(view.front());
      view.remove_prefix(1);
      continue;
    }
    const std::uint32_t cp = d.codepoint;
    if (cp >= 0xFF01 && cp <= 0xFF5E) {
      append_codepoint(out, cp - 0xFEE0);
    } else if (cp == 0x3000) {
      out.push_back(' ');
    } else if (cp == 0xFFE5) {
      // Full-width yen sign -> ASCII backslash. Mac Excel ja-JP follows the
      // CP932 mapping where 0x5C is the yen position; the full-width yen
      // collapses to U+005C even though a half-width yen U+00A5 also exists.
      out.push_back('\\');
    } else if (cp >= 0x30A1 && cp <= 0x30FC) {
      const auto& m = kKatakanaFullToHalf[cp - 0x30A1];
      if (m.first == 0) {
        // Archaic katakana without a half-width mapping: pass through.
        out.append(view.data(), d.byte_len);
      } else {
        append_codepoint(out, m.first);
        if (m.second != 0) {
          append_codepoint(out, m.second);
        }
      }
    } else {
      out.append(view.data(), d.byte_len);
    }
    view.remove_prefix(d.byte_len);
  }
  return Value::text(arena.intern(out));
}

// JIS(text) / DBCS(text) - half-width to full-width conversion. Shared
// implementation; the two names differ only in registration.
Value Jis(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  const std::string& src = text.value();
  std::string out;
  // Full-width codepoints are 3 UTF-8 bytes; a single ASCII input grows
  // to 3 bytes. Reserve generously to avoid re-allocations.
  out.reserve(src.size() * 3);
  std::string_view view = src;
  while (!view.empty()) {
    const Utf8DecodeResult d = decode_first_utf8_codepoint(view);
    if (!d.valid) {
      out.push_back(view.front());
      view.remove_prefix(1);
      continue;
    }
    const std::uint32_t cp = d.codepoint;
    if (cp == 0x20) {
      append_codepoint(out, 0x3000);
      view.remove_prefix(d.byte_len);
      continue;
    }
    if (cp >= 0x21 && cp <= 0x7E) {
      append_codepoint(out, cp + 0xFEE0);
      view.remove_prefix(d.byte_len);
      continue;
    }
    if (cp >= 0xFF66 && cp <= 0xFF9D) {
      const std::uint16_t full_base = static_cast<std::uint16_t>(half_to_full_katakana_base(cp));
      view.remove_prefix(d.byte_len);
      // Peek the next codepoint for dakuten / handakuten merging.
      std::uint16_t emitted = full_base;
      if (!view.empty()) {
        const Utf8DecodeResult next = decode_first_utf8_codepoint(view);
        if (next.valid && next.codepoint == 0xFF9E) {
          const std::uint16_t voiced = voiced_full_from_base(full_base);
          if (voiced != 0) {
            emitted = voiced;
            view.remove_prefix(next.byte_len);
          }
        } else if (next.valid && next.codepoint == 0xFF9F) {
          const std::uint16_t semi = semi_voiced_full_from_base(full_base);
          if (semi != 0) {
            emitted = semi;
            view.remove_prefix(next.byte_len);
          }
        }
      }
      append_codepoint(out, emitted);
      continue;
    }
    if (cp >= 0xFF61 && cp <= 0xFF65) {
      // Half-width punctuation -> full-width CJK punctuation.
      std::uint16_t mapped = 0;
      switch (cp) {
        case 0xFF61:
          mapped = 0x3002;  // 。
          break;
        case 0xFF62:
          mapped = 0x300C;  // 「
          break;
        case 0xFF63:
          mapped = 0x300D;  // 」
          break;
        case 0xFF64:
          mapped = 0x3001;  // 、
          break;
        case 0xFF65:
          mapped = 0x30FB;  // ・
          break;
        default:
          break;
      }
      append_codepoint(out, mapped);
      view.remove_prefix(d.byte_len);
      continue;
    }
    if (cp == 0xFF9E) {
      // Lone dakuten -> spacing voiced sound mark.
      append_codepoint(out, 0x309B);
      view.remove_prefix(d.byte_len);
      continue;
    }
    if (cp == 0xFF9F) {
      // Lone handakuten -> spacing semi-voiced sound mark.
      append_codepoint(out, 0x309C);
      view.remove_prefix(d.byte_len);
      continue;
    }
    // Everything else passes through unchanged.
    out.append(view.data(), d.byte_len);
    view.remove_prefix(d.byte_len);
  }
  return Value::text(arena.intern(out));
}

}  // namespace

void register_text_width_builtins(FunctionRegistry& registry) {
  registry.register_function(FunctionDef{"ASC", 1u, 1u, &Asc});
  // DBCS is a documented alias of JIS in Excel; we share the implementation
  // rather than maintain two parallel code paths.
  registry.register_function(FunctionDef{"JIS", 1u, 1u, &Jis});
  registry.register_function(FunctionDef{"DBCS", 1u, 1u, &Jis});
}

}  // namespace eval
}  // namespace formulon
