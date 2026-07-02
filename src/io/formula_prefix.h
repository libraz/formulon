// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Storage-prefix normalization for formula text.
//
// Excel stores post-2007 / dynamic-array function names and LET / LAMBDA
// parameter names with hidden storage prefixes so that older Excel
// versions do not mis-evaluate them:
//
//   * `_xlfn.`        — post-2007 functions (IFS, TEXTJOIN, CONCAT, LET,
//                       XLOOKUP, SEQUENCE, ...).
//   * `_xlfn._xlws.`  — the subset of the above that are worksheet-only
//                       dynamic-array functions (FILTER, SORT, ...).
//   * `_xlws.`        — a legacy standalone worksheet marker.
//   * `_xlpm.`        — LET / LAMBDA parameter names.
//
// Excel's formula bar shows the names WITHOUT these prefixes. Formulon's
// public `formula_text` (surfaced through every binding) must match the
// formula bar, and the evaluator / parser reason over the canonical
// names, so the prefixes are stripped at ingestion. `strip_storage_prefixes`
// removes them from a stored formula string, leaving string literals
// untouched.

#ifndef FORMULON_IO_FORMULA_PREFIX_H_
#define FORMULON_IO_FORMULA_PREFIX_H_

#include <array>
#include <cstddef>
#include <string>
#include <string_view>

namespace formulon {
namespace io {

/// True when `c` can continue an identifier (letters, digits, `_`, `.`,
/// plus any non-ASCII UTF-8 byte, matching the tokenizer's identifier
/// rule closely enough for boundary detection).
inline bool formula_prefix_is_ident_byte(char c) noexcept {
  const auto u = static_cast<unsigned char>(c);
  return (u >= 'A' && u <= 'Z') || (u >= 'a' && u <= 'z') || (u >= '0' && u <= '9') || u == '_' || u == '.' ||
         u >= 0x80;
}

/// Case-insensitive ASCII compare of `s` against the lowercase `lit`.
inline bool formula_prefix_ci_eq(std::string_view s, std::string_view lit) noexcept {
  if (s.size() != lit.size()) {
    return false;
  }
  for (std::size_t i = 0; i < s.size(); ++i) {
    char c = s[i];
    if (c >= 'A' && c <= 'Z') {
      c = static_cast<char>(c - 'A' + 'a');
    }
    if (c != lit[i]) {
      return false;
    }
  }
  return true;
}

/// Removes Excel's `_xlfn._xlws.` / `_xlfn.` / `_xlws.` / `_xlpm.` storage
/// prefixes from `formula`, returning the canonical (formula-bar) form.
///
/// The removal is applied only OUTSIDE double-quoted string literals (so a
/// literal such as `"cost _xlfn. note"` is preserved verbatim), and only
/// when the prefix starts at an identifier boundary — i.e. the byte before
/// it does not itself continue an identifier — so a defined name that
/// merely embeds the letters (`foo_xlfn.bar`) is not disturbed. A leading
/// `=` and every other character pass through unchanged, so the transform
/// is safe to run on an already-canonical formula (it is a no-op there).
inline std::string strip_storage_prefixes(std::string_view formula) {
  // Longest first so `_xlfn._xlws.` is consumed as a unit before `_xlfn.`.
  static constexpr std::array<std::string_view, 4> kPrefixes = {"_xlfn._xlws.", "_xlfn.", "_xlws.", "_xlpm."};
  std::string out;
  out.reserve(formula.size());
  bool in_string = false;
  for (std::size_t i = 0; i < formula.size();) {
    const char c = formula[i];
    if (in_string) {
      out.push_back(c);
      if (c == '"') {
        // A doubled `""` is an escaped quote, not a string terminator.
        if (i + 1 < formula.size() && formula[i + 1] == '"') {
          out.push_back('"');
          i += 2;
          continue;
        }
        in_string = false;
      }
      ++i;
      continue;
    }
    if (c == '"') {
      in_string = true;
      out.push_back(c);
      ++i;
      continue;
    }
    if (c == '_') {
      const bool at_boundary = out.empty() || !formula_prefix_is_ident_byte(out.back());
      if (at_boundary) {
        bool matched = false;
        for (const std::string_view prefix : kPrefixes) {
          if (i + prefix.size() <= formula.size() && formula_prefix_ci_eq(formula.substr(i, prefix.size()), prefix)) {
            i += prefix.size();
            matched = true;
            break;
          }
        }
        if (matched) {
          continue;
        }
      }
    }
    out.push_back(c);
    ++i;
  }
  return out;
}

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_FORMULA_PREFIX_H_
