//
// Excel's `*` / `?` / `~` wildcard dialect used by criteria matching and the
// text and lookup builtins. See `wildcard.h` for the authoritative contract.

#include "eval/wildcard.h"

#include <cstddef>
#include <cstdint>
#include <string_view>

#include "utils/strings.h"
#include "utils/text_ops.h"
#include "utils/utf8_length.h"

namespace formulon {
namespace eval {

bool scan_has_wildcard(std::string_view rhs) {
  for (std::size_t i = 0; i < rhs.size(); ++i) {
    const char c = rhs[i];
    if (c == '~' && i + 1 < rhs.size()) {
      ++i;  // Skip the escaped byte.
      continue;
    }
    if (c == '*' || c == '?') {
      return true;
    }
  }
  return false;
}

namespace {

// Core two-pointer wildcard matcher. When `prefix_match_ok` is false the
// pattern must consume `text` exactly (SEARCH whole-string / criteria
// equality semantics). When `prefix_match_ok` is true the match succeeds
// as soon as the pattern is exhausted, even if `text` has unmatched bytes
// remaining — used by `wildcard_find` to answer "does the pattern match a
// prefix of this suffix?". When the match succeeds in prefix mode,
// `*out_consumed` receives the number of `text` bytes the pattern
// consumed.
//
// Character comparisons are ASCII case-insensitive to match Excel's
// criteria semantics (e.g. `*blue` matches `BLUE`). Non-ASCII bytes are
// compared byte-for-byte, consistent with `strings::case_insensitive_eq`.
// UTF-8 continuation: returns the byte length of the code point that
// starts at `text[i]`, capped at `text.size() - i`. A leading byte with
// an invalid continuation pattern falls back to 1 byte (mirrors Excel's
// lenient "one display unit" heuristic for broken UTF-8).
std::size_t utf8_codepoint_bytes(std::string_view text, std::size_t i) {
  if (i >= text.size()) {
    return 0;
  }
  const auto lead = static_cast<unsigned char>(text[i]);
  std::size_t len = 1;
  if ((lead & 0x80u) == 0x00u) {
    len = 1;
  } else if ((lead & 0xE0u) == 0xC0u) {
    len = 2;
  } else if ((lead & 0xF0u) == 0xE0u) {
    len = 3;
  } else if ((lead & 0xF8u) == 0xF0u) {
    len = 4;
  }
  if (i + len > text.size()) {
    len = text.size() - i;
  }
  return len;
}

bool wildcard_match_impl(std::string_view pattern, std::string_view text, bool prefix_match_ok,
                         std::size_t* out_consumed, bool question_is_byte = false) {
  std::size_t pi = 0;
  std::size_t ti = 0;
  std::size_t star_pi = std::string_view::npos;
  std::size_t star_ti = 0;
  while (ti < text.size()) {
    // Prefix mode: succeed as soon as the pattern is fully consumed, even
    // if there is unmatched text remaining.
    if (prefix_match_ok && pi == pattern.size()) {
      if (out_consumed != nullptr) {
        *out_consumed = ti;
      }
      return true;
    }
    if (pi < pattern.size()) {
      const char pc = pattern[pi];
      if (pc == '~' && pi + 1 < pattern.size()) {
        // Escaped literal: match the next pattern byte case-insensitively
        // (Excel preserves the overall case-insensitive rule even for
        // escape-protected characters).
        if (strings::ascii_to_lower(pattern[pi + 1]) == strings::ascii_to_lower(text[ti])) {
          pi += 2;
          ++ti;
          continue;
        }
      } else if (pc == '~' && pi + 1 == pattern.size()) {
        // Trailing unpaired `~`: Excel treats it as a no-op (the escape
        // consumes nothing of the text), so the remaining pattern is
        // effectively empty — succeed if we're in prefix mode or if the
        // text is also exhausted at this point.
        ++pi;
        continue;
      } else if (pc == '*') {
        star_pi = pi;
        star_ti = ti;
        ++pi;
        continue;
      } else if (pc == '?') {
        if (question_is_byte) {
          // SEARCHB byte-mode `?`: the pattern character matches only a
          // codepoint whose ja-JP DBCS cost is 1 (ASCII or half-width
          // katakana). Anything wider (kanji, hiragana, full-width, emoji)
          // refuses to match here, falling through to the `*`-backtrack /
          // failure path below.
          std::size_t cp_bytes = 0;
          const std::uint32_t cp = decode_utf8_step(text, ti, &cp_bytes);
          if (byte_count_jajp(cp) == 1) {
            ++pi;
            ti += cp_bytes;
            continue;
          }
        } else {
          // Default mode: `?` matches exactly one UTF-8 code point (1-4
          // bytes) rather than a single byte, so multibyte scripts (CJK,
          // emoji, combining marks treated as atomic) match as Excel users
          // would expect.
          const std::size_t cp_bytes = utf8_codepoint_bytes(text, ti);
          ++pi;
          ti += cp_bytes;
          continue;
        }
      } else if (strings::ascii_to_lower(pc) == strings::ascii_to_lower(text[ti])) {
        ++pi;
        ++ti;
        continue;
      }
    }
    if (star_pi != std::string_view::npos) {
      // Backtrack: extend the most-recent `*` by one CODEPOINT of text. This
      // must advance a whole UTF-8 codepoint, not a single byte — otherwise
      // `star_ti` can land mid-sequence and a following `?` or literal would
      // match a continuation byte (e.g. `*??c*` against `あcう` split the
      // 3-byte `あ`). Codepoint-stepping keeps every subsequent match aligned.
      pi = star_pi + 1;
      const std::size_t step = utf8_codepoint_bytes(text, star_ti);
      star_ti += (step == 0 ? 1 : step);
      ti = star_ti;
      continue;
    }
    return false;
  }
  // Trailing `*`s consume nothing.
  while (pi < pattern.size() && pattern[pi] == '*') {
    ++pi;
  }
  // A single unpaired trailing `~` also consumes nothing (Excel leniency).
  // This must be the very last pattern char, otherwise it is part of an
  // escape pair `~X` and that pair requires a text char to match.
  if (pi + 1 == pattern.size() && pattern[pi] == '~') {
    ++pi;
  }
  const bool ok = (pi == pattern.size());
  if (ok && out_consumed != nullptr) {
    *out_consumed = ti;
  }
  return ok;
}

}  // namespace

bool wildcard_match(std::string_view pattern, std::string_view text) {
  return wildcard_match_impl(pattern, text, /*prefix_match_ok=*/false, /*out_consumed=*/nullptr,
                             /*question_is_byte=*/false);
}

std::size_t wildcard_find(std::string_view pattern, std::string_view text) {
  // Try each start offset in `text` and attempt a prefix match. Because `*`
  // already permits zero-or-more bytes, a leading `*` in the pattern makes
  // offset 0 always match — but the loop below is still correct in that
  // case (it would find offset 0 on the first iteration).
  for (std::size_t start = 0; start <= text.size(); ++start) {
    if (wildcard_match_impl(pattern, text.substr(start), /*prefix_match_ok=*/true,
                            /*out_consumed=*/nullptr, /*question_is_byte=*/false)) {
      return start;
    }
  }
  return std::string_view::npos;
}

std::size_t wildcard_find_dbcs(std::string_view pattern, std::string_view text) {
  // Iterate start positions over codepoint boundaries (not byte-by-byte) so
  // the returned offset is always codepoint-aligned. SEARCHB looks the
  // offset up against `build_dbcs_char_map`'s `byte_offset` records via
  // direct equality; a misaligned offset would silently route into the
  // defensive `total_dbcs + 1` fallback.
  std::size_t start = 0;
  while (start <= text.size()) {
    if (wildcard_match_impl(pattern, text.substr(start), /*prefix_match_ok=*/true,
                            /*out_consumed=*/nullptr, /*question_is_byte=*/true)) {
      return start;
    }
    if (start == text.size()) {
      break;
    }
    const std::size_t step = utf8_codepoint_bytes(text, start);
    start += (step == 0 ? 1 : step);
  }
  return std::string_view::npos;
}

}  // namespace eval
}  // namespace formulon
