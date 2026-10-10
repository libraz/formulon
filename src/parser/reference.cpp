//
// Canonical A1 text for a `Reference`; the contract is declared in
// `parser/reference.h`.

#include "parser/reference.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>

#include "parser/parser_detail.h"
#include "utils/a1_column.h"
#include "utils/expected.h"  // FM_CHECK
#include "utils/strings.h"

namespace formulon {
namespace parser {

// ---------------------------------------------------------------------------
// Reference helper
// ---------------------------------------------------------------------------

namespace {

// True when every byte of `name` belongs to the bare qualifier run.
bool IsBareRun(std::string_view name) noexcept {
  return detail::BareQualifierRunLength(name, 0) == name.size();
}

// True when `name` contains a byte outside `[A-Za-z0-9_.]`.
bool HasNonBareAsciiByte(std::string_view name) noexcept {
  for (char c : name) {
    const bool bare = detail::IsAsciiLetter(c) || detail::IsAsciiDigit(c) || c == '_' || c == '.';
    if (!bare) {
      return true;
    }
  }
  return false;
}

// Advances `*i` over a run of ASCII digits and returns how many it skipped.
std::size_t SkipDigits(std::string_view name, std::size_t* i) noexcept {
  const std::size_t start = *i;
  while (*i < name.size() && detail::IsAsciiDigit(name[*i])) {
    ++*i;
  }
  return *i - start;
}

// A cell inside the A1 grid (`A1`..`XFD1048576`, any letter case): the
// tokenizer reads such a run as a cell reference. `XFE1` lies past the last
// column and is an ordinary name.
bool IsA1CellShaped(std::string_view name) noexcept {
  std::size_t i = 0;
  std::uint32_t col = 0;
  while (i < name.size() && detail::IsAsciiLetter(name[i]) && i < 3) {
    const char upper = strings::ascii_to_upper(name[i]);
    col = col * 26U + static_cast<std::uint32_t>(upper - 'A') + 1U;
    ++i;
  }
  if (i == 0 || col > detail::kMaxColumn) {
    return false;
  }
  const std::size_t digits_begin = i;
  if (SkipDigits(name, &i) == 0 || i != name.size() || name[digits_begin] == '0') {
    return false;
  }
  return i - digits_begin <= 7 &&
         detail::DecodeDigitRunClamped(name.substr(digits_begin), detail::kMaxRow) <= detail::kMaxRow;
}

// `R`, `C`, `R<n>`, `C<n>`, `R<n>C<n>`, `RC<n>`, `R<n>C`, `RC`, ASCII
// case-insensitively: the absolute R1C1 shapes a sheet name could be read as.
bool IsR1C1Shaped(std::string_view name) noexcept {
  std::size_t i = 0;
  const auto at = [&](char upper) { return i < name.size() && strings::ascii_to_upper(name[i]) == upper; };
  if (at('R')) {
    ++i;
    SkipDigits(name, &i);
    if (at('C')) {
      ++i;
      SkipDigits(name, &i);
    }
    return i == name.size();
  }
  if (at('C')) {
    ++i;
    SkipDigits(name, &i);
    return i == name.size();
  }
  return false;
}

// The rules both notations share: an empty name, a byte outside the bare
// ASCII run, a leading digit (read as a number), TRUE / FALSE (read as a
// bool) and an R1C1 shape.
bool LocalSheetNeedsQuotingInEitherNotation(std::string_view name) noexcept {
  return name.empty() || HasNonBareAsciiByte(name) || detail::IsAsciiDigit(name.front()) ||
         strings::case_insensitive_eq(name, "TRUE") || strings::case_insensitive_eq(name, "FALSE") ||
         IsR1C1Shaped(name);
}

}  // namespace

bool local_sheet_needs_quoting_a1(std::string_view name) noexcept {
  return LocalSheetNeedsQuotingInEitherNotation(name) || IsA1CellShaped(name);
}

bool local_sheet_needs_quoting_r1c1(std::string_view name) noexcept {
  return LocalSheetNeedsQuotingInEitherNotation(name);
}

bool external_book_needs_quoting(std::string_view book) noexcept {
  return book.empty() || !IsBareRun(book);
}

bool external_sheet_needs_quoting(std::string_view sheet) noexcept {
  return !sheet.empty() && (detail::IsAsciiDigit(sheet.front()) || !IsBareRun(sheet));
}

std::string format_a1(const Reference& r) {
  FM_CHECK(!(r.is_full_col && r.is_full_row), "format_a1: is_full_col and is_full_row must not both be set");
  std::string out;
  if (!r.sheet.empty()) {
    if (r.sheet_quoted || local_sheet_needs_quoting_a1(r.sheet)) {
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
