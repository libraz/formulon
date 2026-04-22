// Copyright 2026 libraz. Licensed under the MIT License.
//
// ASCII-only string utilities shared across the Formulon engine.
//
// The helpers in this header deliberately avoid locale-aware case folding and
// Unicode normalisation. Those concerns belong to the text layer (see the
// `text/` module and `backup/plans/11-text-layer.md`); the primitives below
// are a small, dependency-free toolkit suitable for parser, tokenizer and
// internal diagnostics work where all inputs are guaranteed to be 7-bit
// ASCII or treated as opaque bytes.

#ifndef FORMULON_UTILS_STRINGS_H_
#define FORMULON_UTILS_STRINGS_H_

#include <cstddef>
#include <string>
#include <string_view>
#include <vector>

namespace formulon {
namespace strings {

/// Returns true if `c` is one of the six ASCII whitespace characters
/// recognised by the trim helpers: space, tab, CR, LF, vertical tab, form
/// feed.
constexpr bool is_ascii_space(char c) noexcept {
  return c == ' ' || c == '\t' || c == '\r' || c == '\n' || c == '\v' || c == '\f';
}

/// Returns `input` with leading ASCII whitespace removed.
inline std::string_view ltrim(std::string_view input) noexcept {
  std::size_t start = 0;
  while (start < input.size() && is_ascii_space(input[start])) {
    ++start;
  }
  return input.substr(start);
}

/// Returns `input` with trailing ASCII whitespace removed.
inline std::string_view rtrim(std::string_view input) noexcept {
  std::size_t end = input.size();
  while (end > 0 && is_ascii_space(input[end - 1])) {
    --end;
  }
  return input.substr(0, end);
}

/// Returns `input` with leading and trailing ASCII whitespace removed.
inline std::string_view trim(std::string_view input) noexcept {
  return rtrim(ltrim(input));
}

/// Splits `input` on every occurrence of the single-byte `delimiter`.
///
/// Empty fields are preserved. Examples:
///   split(",,a,", ',') -> {"", "", "a", ""}
///   split("",    ',') -> {""}
inline std::vector<std::string_view> split(std::string_view input, char delimiter) {
  std::vector<std::string_view> parts;
  std::size_t begin = 0;
  for (std::size_t i = 0; i < input.size(); ++i) {
    if (input[i] == delimiter) {
      parts.emplace_back(input.substr(begin, i - begin));
      begin = i + 1;
    }
  }
  parts.emplace_back(input.substr(begin));
  return parts;
}

/// Splits `input` on every occurrence of the non-empty byte sequence
/// `delimiter`. Passing an empty delimiter is a programmer error and is
/// treated as "no splits" (returns `{input}`).
///
/// Empty `input` yields a single empty element, matching `split(char)`.
inline std::vector<std::string_view> split(std::string_view input, std::string_view delimiter) {
  std::vector<std::string_view> parts;
  if (delimiter.empty()) {
    parts.emplace_back(input);
    return parts;
  }
  std::size_t begin = 0;
  while (begin <= input.size()) {
    const std::size_t pos = input.find(delimiter, begin);
    if (pos == std::string_view::npos) {
      parts.emplace_back(input.substr(begin));
      break;
    }
    parts.emplace_back(input.substr(begin, pos - begin));
    begin = pos + delimiter.size();
  }
  return parts;
}

/// Concatenates `parts` with `sep` between each consecutive pair.
inline std::string join(const std::vector<std::string>& parts, std::string_view sep) {
  if (parts.empty()) {
    return {};
  }
  std::size_t total = 0;
  for (const auto& p : parts) {
    total += p.size();
  }
  total += sep.size() * (parts.size() - 1);
  std::string out;
  out.reserve(total);
  for (std::size_t i = 0; i < parts.size(); ++i) {
    if (i != 0) {
      out.append(sep.data(), sep.size());
    }
    out.append(parts[i]);
  }
  return out;
}

/// Concatenates `parts` with `sep` between each consecutive pair.
inline std::string join(const std::vector<std::string_view>& parts, std::string_view sep) {
  if (parts.empty()) {
    return {};
  }
  std::size_t total = 0;
  for (const auto& p : parts) {
    total += p.size();
  }
  total += sep.size() * (parts.size() - 1);
  std::string out;
  out.reserve(total);
  for (std::size_t i = 0; i < parts.size(); ++i) {
    if (i != 0) {
      out.append(sep.data(), sep.size());
    }
    out.append(parts[i].data(), parts[i].size());
  }
  return out;
}

/// Returns the ASCII lowercase form of `c`, or `c` unchanged for non-letters.
constexpr char ascii_to_lower(char c) noexcept {
  return (c >= 'A' && c <= 'Z') ? static_cast<char>(c + ('a' - 'A')) : c;
}

/// Returns the ASCII uppercase form of `c`, or `c` unchanged for non-letters.
constexpr char ascii_to_upper(char c) noexcept {
  return (c >= 'a' && c <= 'z') ? static_cast<char>(c - ('a' - 'A')) : c;
}

/// Case-insensitive equality over the ASCII letters only. Any byte with the
/// high bit set (for example, the leading byte of a UTF-8 multibyte sequence)
/// is compared verbatim, so `"straße" != "STRASSE"` and
/// `"\xe3\x81\x82" == "\xe3\x81\x82"`.
inline bool case_insensitive_eq(std::string_view a, std::string_view b) noexcept {
  if (a.size() != b.size()) {
    return false;
  }
  for (std::size_t i = 0; i < a.size(); ++i) {
    if (ascii_to_lower(a[i]) != ascii_to_lower(b[i])) {
      return false;
    }
  }
  return true;
}

/// Returns a lowercase copy of `input` using `ascii_to_lower` byte by byte.
/// Non-ASCII bytes are preserved exactly.
inline std::string to_ascii_lower(std::string_view input) {
  std::string out;
  out.resize(input.size());
  for (std::size_t i = 0; i < input.size(); ++i) {
    out[i] = ascii_to_lower(input[i]);
  }
  return out;
}

/// Returns an uppercase copy of `input` using `ascii_to_upper` byte by byte.
/// Non-ASCII bytes are preserved exactly.
inline std::string to_ascii_upper(std::string_view input) {
  std::string out;
  out.resize(input.size());
  for (std::size_t i = 0; i < input.size(); ++i) {
    out[i] = ascii_to_upper(input[i]);
  }
  return out;
}

/// Returns true iff `haystack` begins with `prefix`.
inline bool starts_with(std::string_view haystack, std::string_view prefix) noexcept {
  if (prefix.size() > haystack.size()) {
    return false;
  }
  return haystack.compare(0, prefix.size(), prefix) == 0;
}

/// Returns true iff `haystack` ends with `suffix`.
inline bool ends_with(std::string_view haystack, std::string_view suffix) noexcept {
  if (suffix.size() > haystack.size()) {
    return false;
  }
  return haystack.compare(haystack.size() - suffix.size(), suffix.size(), suffix) == 0;
}

}  // namespace strings
}  // namespace formulon

#endif  // FORMULON_UTILS_STRINGS_H_
