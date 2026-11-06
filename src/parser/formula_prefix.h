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
//   * `_xleta.`       — a built-in function named as a value (`TYPE(SUM)`).
//
// Excel's formula bar shows the names WITHOUT these prefixes. Formulon's
// public `formula_text` (surfaced through every binding) must match the
// formula bar, and the evaluator / parser reason over the canonical
// names, so the prefixes are stripped at ingestion. `strip_storage_prefixes`
// removes them from a stored formula string, leaving string literals
// untouched.

#ifndef FORMULON_PARSER_FORMULA_PREFIX_H_
#define FORMULON_PARSER_FORMULA_PREFIX_H_

#include <array>
#include <cstddef>
#include <string>
#include <string_view>

#include "parser/reference.h"

namespace formulon {
namespace parser {

/// Storage prefix of a LET / LAMBDA parameter name.
inline constexpr std::string_view kParamPrefix = "_xlpm.";
/// Storage prefix of a cell-shaped defined name.
inline constexpr std::string_view kNamePrefix = "_xlnm.";

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

/// Removes Excel's `_xlfn._xlws.` / `_xlfn.` / `_xlws.` / `_xlpm.` / `_xleta.`
/// storage prefixes from `formula`, returning the canonical (formula-bar) form.
///
/// The removal is applied only OUTSIDE double-quoted string literals (so a
/// literal such as `"cost _xlfn. note"` is preserved verbatim), and only
/// when the prefix starts at an identifier boundary — i.e. the byte before
/// it does not itself continue an identifier — so a defined name that
/// merely embeds the letters (`foo_xlfn.bar`) is not disturbed. A leading
/// `=` and every other character pass through unchanged, so the transform
/// is safe to run on an already-canonical formula (it is a no-op there).
///
/// With `known`, a function prefix (`_xlfn.`, `_xlfn._xlws.`, `_xlws.`) is
/// removed only before a name `known` accepts: Excel keeps it on a name it
/// does not recognise (`_xlfn.FOOBAR(1)` stays so in the formula bar and the
/// file), and the writers re-apply it only to the names they know. `_xleta.`
/// always goes, and so does `_xlpm.` unless the identifier run after it is
/// A1 cell-shaped (`_xlpm.B1`), which keeps it so it cannot read as a cell.
inline std::string strip_storage_prefixes(std::string_view formula, bool (*known)(std::string_view) = nullptr) {
  // Longest first so `_xlfn._xlws.` is consumed as a unit before `_xlfn.`.
  static constexpr std::array<std::string_view, 5> kPrefixes = {"_xlfn._xlws.", "_xlfn.", "_xlws.", kParamPrefix,
                                                                "_xleta."};
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
            std::size_t end = i + prefix.size();
            while (end < formula.size() && formula_prefix_is_ident_byte(formula[end])) {
              ++end;
            }
            const std::string_view name = formula.substr(i + prefix.size(), end - i - prefix.size());
            if (prefix == kParamPrefix && is_a1_cell_shaped(name)) {
              break;
            }
            if (prefix != kParamPrefix && prefix != "_xleta." && known != nullptr && !known(name)) {
              break;
            }
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

/// `name` without its `_xlpm.` / `_xlnm.` prefix (ASCII case-insensitive)
/// when the rest is A1 cell-shaped (`_xlpm.B1` -> `B1`); `name` otherwise.
std::string_view cell_shaped_name_bare(std::string_view name) noexcept;

/// The prefix `cell_shaped_name_bare` drops, spelled `_xlpm.` or `_xlnm.`
/// in lower case; empty when it drops none.
std::string_view cell_shaped_name_prefix(std::string_view name) noexcept;

/// The formula-bar spelling of the cell-shaped names in stored text: outside
/// string literals and quoted sheet names, at an identifier boundary,
/// `_xlpm.X` becomes `X` and `_xlnm.X` becomes `'X'` for a cell-shaped `X`
/// (`Sheet1!_xlnm.B1` -> `Sheet1!'B1'`). Every other byte is copied.
std::string display_cell_shaped_names(std::string_view formula);

/// Respells each one-argument call to `SINGLE` / `ANCHORARRAY` / `_TRO_*` in
/// `formula` (a leading `=` optional) as the `@x` / `x#` / `A1:.A10` operator
/// Excel's formula bar shows for it, copying every other character verbatim. Returns `formula`
/// unchanged when it holds none or does not parse.
std::string spell_storage_operators(std::string_view formula);

}  // namespace parser
}  // namespace formulon

#endif  // FORMULON_PARSER_FORMULA_PREFIX_H_
