//
// Canonical A1 text for a `Reference`; the contract is declared in
// `parser/reference.h`.

#include "parser/reference.h"

#include <cstddef>
#include <string>
#include <string_view>

#include "utils/a1_column.h"
#include "utils/expected.h"  // FM_CHECK
#include "utils/strings.h"

namespace formulon {
namespace parser {

// ---------------------------------------------------------------------------
// Reference helper
// ---------------------------------------------------------------------------

bool sheet_name_needs_quoting(std::string_view name) noexcept {
  if (name.empty()) {
    return true;
  }
  // A bare sheet name cannot open with a digit: the tokenizer reads the run
  // as a numeric literal and never reaches the `!`, so `3Q!A1` is not a
  // reference at all. Quoting is the only way to write such a name.
  if (name.front() >= '0' && name.front() <= '9') {
    return true;
  }
  // A sheet named TRUE/FALSE (any ASCII case) tokenizes as a Bool literal
  // before the `!` is ever consulted, so `TRUE!A1` re-parses as `(bool
  // true)` followed by a dangling `!`. Quoting is the only unambiguous form.
  if (strings::case_insensitive_eq(name, "TRUE") || strings::case_insensitive_eq(name, "FALSE")) {
    return true;
  }

  std::size_t i = 0;
  while (i < name.size() && ((name[i] >= 'A' && name[i] <= 'Z') || (name[i] >= 'a' && name[i] <= 'z'))) {
    ++i;
  }
  const std::size_t letters = i;
  const bool cell_ref_prefix = letters >= 1 && letters <= 3 && i < name.size() && name[i] >= '1' && name[i] <= '9';
  for (char c : name) {
    const bool bare_name_char =
        (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || (c >= '0' && c <= '9') || c == '_' || c == '.';
    if (!bare_name_char) {
      return true;
    }
  }
  if (!cell_ref_prefix) {
    return false;
  }
  ++i;
  while (i < name.size() && name[i] >= '0' && name[i] <= '9') {
    ++i;
  }
  return i == name.size();
}

std::string format_a1(const Reference& r) {
  FM_CHECK(!(r.is_full_col && r.is_full_row), "format_a1: is_full_col and is_full_row must not both be set");
  std::string out;
  if (!r.sheet.empty()) {
    if (r.sheet_quoted || sheet_name_needs_quoting(r.sheet)) {
      out.push_back('\'');
      // Escape any embedded single quotes by doubling them.
      for (char c : r.sheet) {
        if (c == '\'') {
          out.push_back('\'');
        }
        out.push_back(c);
      }
      out.push_back('\'');
    } else {
      out.append(r.sheet);
    }
    out.push_back('!');
  }
  if (r.is_full_col) {
    if (r.col_abs) {
      out.push_back('$');
    }
    FM_CHECK(a1::append_column_letters(out, r.col), "reference column is outside Excel's grid");
    out.push_back(':');
    if (r.col_abs) {
      out.push_back('$');
    }
    FM_CHECK(a1::append_column_letters(out, r.col), "reference column is outside Excel's grid");
    return out;
  }
  if (r.is_full_row) {
    if (r.row_abs) {
      out.push_back('$');
    }
    out.append(std::to_string(r.row + 1));
    out.push_back(':');
    if (r.row_abs) {
      out.push_back('$');
    }
    out.append(std::to_string(r.row + 1));
    return out;
  }
  if (r.col_abs) {
    out.push_back('$');
  }
  FM_CHECK(a1::append_column_letters(out, r.col), "reference column is outside Excel's grid");
  if (r.row_abs) {
    out.push_back('$');
  }
  out.append(std::to_string(r.row + 1));
  return out;
}

}  // namespace parser
}  // namespace formulon
