//
// Excel's `*` / `?` / `~` wildcard dialect used by criteria matching and the
// text and lookup builtins. See `wildcard.h` for the authoritative contract.

#include "eval/wildcard.h"

#include <cstddef>
#include <string>
#include <string_view>

#include "utils/strings.h"

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

std::string unescape_wildcard_literal(std::string_view pattern) {
  std::string out;
  out.reserve(pattern.size());
  for (std::size_t i = 0; i < pattern.size(); ++i) {
    if (pattern[i] == '~' && i + 1 < pattern.size()) {
      out.push_back(pattern[i + 1]);
      ++i;
      continue;
    }
    out.push_back(pattern[i]);
  }
  return out;
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

// A text position: a byte offset, plus `half` when only the first of the two
// UTF-16 units of the supplementary-plane character at `ti` is consumed.
struct TextPos {
  std::size_t ti = 0;
  bool half = false;
};

// One UTF-16 unit further.
TextPos advance_unit(std::string_view text, TextPos pos) {
  const std::size_t bytes = utf8_codepoint_bytes(text, pos.ti);
  if (pos.half) {
    return {pos.ti + bytes, false};
  }
  if (bytes == 4) {
    return {pos.ti, true};
  }
  return {pos.ti + (bytes == 0 ? 1 : bytes), false};
}

// `?` matches one UTF-16 unit, so a supplementary-plane character takes `??`
// and a literal never matches half of one.
bool wildcard_match_impl(std::string_view pattern, std::string_view text) {
  std::size_t pi = 0;
  TextPos pos;
  std::size_t star_pi = std::string_view::npos;
  TextPos star_pos;
  while (pos.ti < text.size()) {
    if (pi < pattern.size()) {
      const char pc = pattern[pi];
      if (pc == '~' && pi + 1 < pattern.size()) {
        // Escaped literal: match the next pattern byte case-insensitively
        // (Excel preserves the overall case-insensitive rule even for
        // escape-protected characters).
        if (!pos.half && strings::ascii_to_lower(pattern[pi + 1]) == strings::ascii_to_lower(text[pos.ti])) {
          pi += 2;
          ++pos.ti;
          continue;
        }
      } else if (pc == '~' && pi + 1 == pattern.size()) {
        // Trailing unpaired `~`: Excel treats it as a no-op (the escape
        // consumes nothing of the text).
        ++pi;
        continue;
      } else if (pc == '*') {
        star_pi = pi;
        star_pos = pos;
        ++pi;
        continue;
      } else if (pc == '?') {
        ++pi;
        pos = advance_unit(text, pos);
        continue;
      } else if (!pos.half && strings::ascii_to_lower(pc) == strings::ascii_to_lower(text[pos.ti])) {
        ++pi;
        ++pos.ti;
        continue;
      }
    }
    if (star_pi != std::string_view::npos) {
      // Backtrack: extend the most-recent `*` by one UTF-16 unit of text.
      pi = star_pi + 1;
      star_pos = advance_unit(text, star_pos);
      pos = star_pos;
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
  return pi == pattern.size();
}

}  // namespace

bool wildcard_match(std::string_view pattern, std::string_view text) {
  return wildcard_match_impl(pattern, text);
}

}  // namespace eval
}  // namespace formulon
