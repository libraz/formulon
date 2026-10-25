//
// Canonical A1 text for a `Reference`; the contract is declared in
// `parser/reference.h`.

#include "parser/reference.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>

#include "excel_locale.h"
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

// Advances `*i` over a run of ASCII digits and returns how many it skipped.
std::size_t SkipDigits(std::string_view name, std::size_t* i) noexcept {
  const std::size_t start = *i;
  while (*i < name.size() && detail::IsAsciiDigit(name[*i])) {
    ++*i;
  }
  return *i - start;
}

// A cell inside the A1 grid (`A1`..`XFD1048576`, any letter case): the
// tokenizer reads such a run as a cell reference. Zero padding counts
// (`A01` is row 1); `XFE1` and `A0` lie off the grid and are ordinary names.
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
  if (SkipDigits(name, &i) == 0 || i != name.size()) {
    return false;
  }
  const std::uint64_t row = detail::DecodeDigitRunClamped(name.substr(digits_begin), detail::kMaxRow);
  return row >= 1U && row <= detail::kMaxRow;
}

// `<row>`, `<col>`, `<row><n>`, `<col><n>`, `<row><n><col><n>`, `<row><col><n>`,
// `<row><n><col>`, `<row><col>` for the given R1C1 letters, ASCII
// case-insensitively: the absolute R1C1 shapes a sheet name could be read as.
bool IsR1C1ShapedWith(std::string_view name, char row_letter, char col_letter) noexcept {
  std::size_t i = 0;
  const auto at = [&](char upper) { return i < name.size() && strings::ascii_to_upper(name[i]) == upper; };
  if (at(row_letter)) {
    ++i;
    SkipDigits(name, &i);
    if (at(col_letter)) {
      ++i;
      SkipDigits(name, &i);
    }
    return i == name.size();
  }
  if (at(col_letter)) {
    ++i;
    SkipDigits(name, &i);
    return i == name.size();
  }
  return false;
}

bool IsR1C1Shaped(std::string_view name) noexcept {
  return IsR1C1ShapedWith(name, 'R', 'C');
}

// `R<n>K<n>`: read as a reference in every locale, not only where `K` is the
// column letter.
bool IsRowKColShaped(std::string_view name) noexcept {
  std::size_t i = 0;
  const auto letter = [&](char upper) {
    if (i >= name.size() || strings::ascii_to_upper(name[i]) != upper) {
      return false;
    }
    ++i;
    return true;
  };
  return letter('R') && SkipDigits(name, &i) != 0 && letter('K') && SkipDigits(name, &i) != 0 && i == name.size();
}

// The bare qualifier run, with U+3000 allowed where the locale is not
// double-byte.
bool IsBareRunIn(std::string_view name, bool dbcs) noexcept {
  if (dbcs) {
    return IsBareRun(name);
  }
  constexpr std::string_view kIdeographicSpace = "\xE3\x80\x80";
  std::size_t begin = 0;
  while (true) {
    const std::size_t end = name.find(kIdeographicSpace, begin);
    if (!IsBareRun(name.substr(begin, end == std::string_view::npos ? end : end - begin))) {
      return false;
    }
    if (end == std::string_view::npos) {
      return true;
    }
    begin = end + kIdeographicSpace.size();
  }
}

// The rules both notations share: an empty name, a byte outside the bare
// qualifier run, a leading digit (read as a number), TRUE / FALSE (read as a
// bool) and an R1C1 shape.
bool LocalSheetNeedsQuotingInEitherNotation(std::string_view name) noexcept {
  return name.empty() || !IsBareRun(name) || detail::IsAsciiDigit(name.front()) ||
         strings::case_insensitive_eq(name, "TRUE") || strings::case_insensitive_eq(name, "FALSE") ||
         IsR1C1Shaped(name);
}

}  // namespace

bool local_sheet_needs_quoting_a1(std::string_view name) noexcept {
  return LocalSheetNeedsQuotingInEitherNotation(name) || IsA1CellShaped(name);
}

bool local_sheet_needs_quoting_display(std::string_view name, const LocaleFacts& facts) noexcept {
  if (name.empty() || !IsBareRunIn(name, facts.dbcs_codepage != DbcsCodepage::kNone) ||
      detail::IsAsciiDigit(name.front()) || IsA1CellShaped(name) || IsRowKColShaped(name)) {
    return true;
  }
  const bool english_bool = strings::case_insensitive_eq(facts.true_name, "TRUE");
  if ((english_bool && (strings::case_insensitive_eq(name, "TRUE") || strings::case_insensitive_eq(name, "FALSE"))) ||
      strings::case_insensitive_eq(name, facts.true_name) || strings::case_insensitive_eq(name, facts.false_name)) {
    return true;
  }
  // Where the row letter stays `R` and the column letter moves (nl), `RC` is no reference.
  const bool rc_free = facts.r1c1_row == 'R' && facts.r1c1_col != 'C' && strings::case_insensitive_eq(name, "RC");
  return (IsR1C1Shaped(name) && !rc_free) || IsR1C1ShapedWith(name, facts.r1c1_row, facts.r1c1_col);
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

bool whole_axis_pair_is_single(const Reference& lhs, const Reference& rhs) noexcept {
  if (lhs.is_full_col && rhs.is_full_col) {
    return lhs.col == rhs.col && lhs.col_abs == rhs.col_abs;
  }
  return lhs.is_full_row && rhs.is_full_row && lhs.row == rhs.row && lhs.row_abs == rhs.row_abs;
}

}  // namespace parser
}  // namespace formulon
