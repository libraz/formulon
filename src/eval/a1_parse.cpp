//
// A1 and R1C1 textual reference parsers declared in `eval/a1_parse.h`:
// `parse_a1_ref`, `parse_r1c1_ref` and `column_letters`, with the sheet
// qualifier and endpoint lexers they share.

#include "eval/a1_parse.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string_view>

#include "eval/eval_profile_scope.h"
#include "excel_locale.h"
#include "parser/parser_detail.h"
#include "sheet.h"
#include "utils/a1_column.h"
#include "utils/strings.h"

namespace formulon {
namespace eval {
namespace refs_internal {

namespace {

// ASCII letter / digit predicates — intentionally local to avoid pulling
// in `<cctype>` which is locale-sensitive on some platforms.
constexpr bool is_letter(char ch) noexcept {
  return (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z');
}

constexpr bool is_digit(char ch) noexcept {
  return ch >= '0' && ch <= '9';
}

// Parses letters (A..XFD) at `text[*i]` into a 1-based column index, then
// advances `*i` past the consumed bytes. Returns 0 on malformed input
// (no letters, too many letters, or overflow past XFD = 16384).
std::uint32_t parse_column_letters(std::string_view text, std::size_t* i) {
  std::uint32_t col = 0;
  std::size_t letters_seen = 0;
  while (*i < text.size() && is_letter(text[*i])) {
    const char ch = strings::ascii_to_upper(text[*i]);
    col = col * 26u + static_cast<std::uint32_t>(ch - 'A' + 1);
    ++(*i);
    ++letters_seen;
    if (letters_seen > 3 || col > Sheet::kMaxCols) {
      return 0;
    }
  }
  if (letters_seen == 0) {
    return 0;
  }
  return col;
}

// Parses digits at `text[*i]` into a 1-based row index, advancing `*i`
// past them. Returns 0 on no digits, too many digits, or row > kMaxRows.
std::uint32_t parse_row_digits(std::string_view text, std::size_t* i) {
  std::uint64_t row = 0;
  std::size_t digits_seen = 0;
  while (*i < text.size() && is_digit(text[*i])) {
    row = row * 10u + static_cast<std::uint32_t>(text[*i] - '0');
    ++(*i);
    ++digits_seen;
    if (digits_seen > 7 || row > Sheet::kMaxRows) {
      return 0;
    }
  }
  if (digits_seen == 0 || row == 0) {
    return 0;
  }
  return static_cast<std::uint32_t>(row);
}

// Parses `[$]?<unit>:[$]?<unit>` consuming the entire remainder of
// `text[start..]`, where `parse_unit` reads one 1-based unit and returns 0
// on malformed input. On success writes the 0-based min / max of the two
// units and returns true.
bool parse_full_axis_span(std::string_view text, std::size_t start,
                          std::uint32_t (*parse_unit)(std::string_view, std::size_t*), std::uint32_t* out_lo,
                          std::uint32_t* out_hi) {
  std::size_t i = start;
  if (i < text.size() && text[i] == '$') {
    ++i;
  }
  const std::uint32_t u1 = parse_unit(text, &i);
  if (u1 == 0) {
    return false;
  }
  if (i >= text.size() || text[i] != ':') {
    return false;
  }
  ++i;
  if (i < text.size() && text[i] == '$') {
    ++i;
  }
  const std::uint32_t u2 = parse_unit(text, &i);
  if (u2 == 0) {
    return false;
  }
  if (i != text.size()) {
    return false;
  }
  *out_lo = std::min(u1, u2) - 1U;
  *out_hi = std::max(u1, u2) - 1U;
  return true;
}

// Attempts to parse `text[start..]` as a full-column shape
// `[$]?<letters>:[$]?<letters>` (`full_col`) or a full-row shape
// `[$]?<digits>:[$]?<digits>`, consuming the entire remainder. On success
// sets `out->is_full_col` / `out->is_full_row` and `out->is_range`,
// populates the spanned axis with the min/max 0-based indices, and fills the
// other axis with its full extent, then returns true. On failure leaves
// `*out` untouched and returns false.
bool try_parse_full_axis(std::string_view text, std::size_t start, bool full_col, A1Parse* out) {
  std::uint32_t lo = 0;
  std::uint32_t hi = 0;
  if (!parse_full_axis_span(text, start, full_col ? parse_column_letters : parse_row_digits, &lo, &hi)) {
    return false;
  }
  if (full_col) {
    out->col = lo;
    out->col2 = hi;
    out->row = 0;
    out->row2 = Sheet::kMaxRows - 1U;
    out->is_full_col = true;
  } else {
    out->row = lo;
    out->row2 = hi;
    out->col = 0;
    out->col2 = Sheet::kMaxCols - 1U;
    out->is_full_row = true;
  }
  out->is_range = true;
  return true;
}

// Parses a single A1 endpoint (optional `$` markers, letters, digits).
// Returns `false` on any malformed shape; on success writes 0-based
// row/col to `*out_row` / `*out_col` and advances `*i`.
bool parse_a1_endpoint(std::string_view text, std::size_t* i, std::uint32_t* out_row, std::uint32_t* out_col) {
  // Optional leading `$` on the column.
  if (*i < text.size() && text[*i] == '$') {
    ++(*i);
  }
  const std::uint32_t col_1based = parse_column_letters(text, i);
  if (col_1based == 0) {
    return false;
  }
  // Optional `$` between column and row.
  if (*i < text.size() && text[*i] == '$') {
    ++(*i);
  }
  const std::uint32_t row_1based = parse_row_digits(text, i);
  if (row_1based == 0) {
    return false;
  }
  *out_col = col_1based - 1U;
  *out_row = row_1based - 1U;
  return true;
}

// Copies `quoted` (everything between the surrounding single quotes)
// into `out`, collapsing each `''` pair into a single `'`. Returns the
// new size. `quoted` already excludes the outer quotes.
std::size_t unescape_quoted_sheet(std::string_view quoted, char* out, std::size_t cap) {
  std::size_t w = 0;
  for (std::size_t r = 0; r < quoted.size() && w < cap; ++r) {
    if (quoted[r] == '\'' && r + 1 < quoted.size() && quoted[r + 1] == '\'') {
      out[w++] = '\'';
      ++r;
      continue;
    }
    out[w++] = quoted[r];
  }
  return w;
}

}  // namespace

// Outcome of reading an optional leading `Sheet!` qualifier.
enum class SheetQualifier : std::uint8_t {
  kNone,       ///< No qualifier; `*i` is unmoved and the whole text is the reference.
  kParsed,     ///< Qualifier consumed; `*sheet` is set and `*i` points past the `!`.
  kMalformed,  ///< Unterminated quote or a quoted name with no `!`; the whole reference is invalid.
};

// Reads the optional sheet qualifier at the head of `text`, advancing `*i`
// past it. Shared by the A1 and R1C1 parsers, which differ only in how
// they read the reference body that follows.
SheetQualifier parse_sheet_qualifier(std::string_view text, std::size_t* i, std::string_view* sheet) {
  if (text.empty()) {
    return SheetQualifier::kNone;
  }
  if (text[0] == '\'') {
    // Scan for the closing `'` that is NOT followed by another `'` (the
    // doubled form is an escaped apostrophe and stays inside the name).
    std::size_t j = 1;
    while (j < text.size()) {
      if (text[j] == '\'') {
        if (j + 1 < text.size() && text[j + 1] == '\'') {
          j += 2;
          continue;
        }
        break;
      }
      ++j;
    }
    if (j >= text.size() || text[j] != '\'') {
      return SheetQualifier::kMalformed;  // unterminated
    }
    // Inside content is `text.substr(1, j - 1)`; unescape `''` -> `'`.
    // We don't own backing storage here, so write into a static-sized
    // local buffer; Excel sheet-name limit is 31 chars (we allow up to
    // 255 defensively, capped by the source view size).
    static thread_local char scratch[256];
    const std::size_t content_len = j - 1;
    const std::size_t used = unescape_quoted_sheet(text.substr(1, content_len), scratch, sizeof(scratch));
    // This view points into thread_local storage — callers keep it alive
    // only until the next parse on the same thread. The consuming code
    // (`indirect`) copies before storing.
    *sheet = std::string_view(scratch, used);
    std::size_t after = j + 1;
    if (after >= text.size() || text[after] != '!') {
      return SheetQualifier::kMalformed;  // missing `!` after quoted sheet
    }
    *i = after + 1;
    return SheetQualifier::kParsed;
  }
  // Bare sheet name: the formula parser's bare qualifier run followed by `!`.
  // We only commit to treating it as a sheet qualifier if we find the
  // `!` — otherwise the run is part of the reference itself.
  const std::size_t j = parser::detail::BareQualifierRunLength(text, 0);
  if (j > 0 && j < text.size() && text[j] == '!') {
    *sheet = text.substr(0, j);
    *i = j + 1;
    return SheetQualifier::kParsed;
  }
  return SheetQualifier::kNone;
}

A1Parse parse_a1_ref(std::string_view text) {
  A1Parse out;
  if (text.empty()) {
    return out;
  }
  std::size_t i = 0;
  if (parse_sheet_qualifier(text, &i, &out.sheet) == SheetQualifier::kMalformed) {
    return out;
  }

  // Full-column / full-row shapes (`D:D`, `$FF:FG`, `5:5`, `$12:$23`)
  // are tried before the single-endpoint path because they never share a
  // prefix with a valid single-cell reference (the latter always has a
  // digit immediately after the letter run, never a `:`).
  if (try_parse_full_axis(text, i, /*full_col=*/true, &out)) {
    out.valid = true;
    return out;
  }
  if (try_parse_full_axis(text, i, /*full_col=*/false, &out)) {
    out.valid = true;
    return out;
  }

  // Parse the first endpoint.
  if (!parse_a1_endpoint(text, &i, &out.row, &out.col)) {
    return out;
  }

  // Optional `:` + second endpoint for ranges.
  if (i < text.size() && text[i] == ':') {
    ++i;
    if (!parse_a1_endpoint(text, &i, &out.row2, &out.col2)) {
      return out;
    }
    out.is_range = true;
  }

  // Trailing garbage -> invalid.
  if (i != text.size()) {
    return out;
  }
  out.valid = true;
  return out;
}

namespace {

// Reads one `R`/`C` axis at `text[*i]` when the letter matches `marker`.
// Returns false only on a malformed bracket; an absent axis leaves `*out`
// unset and still returns true.
bool read_r1c1_axis(std::string_view text, std::size_t* i, const R1C1Letters& letters, char marker,
                    R1C1EndpointAxis* out, bool* bracketed) {
  if (*i >= text.size() || strings::ascii_to_upper(text[*i]) != marker) {
    return true;
  }
  ++*i;
  out->present = true;
  if (*i < text.size() && text[*i] == letters.open) {
    *bracketed = true;
    ++*i;
    bool negative = false;
    if (*i < text.size() && (text[*i] == '-' || text[*i] == '+')) {
      negative = text[*i] == '-';
      ++*i;
    }
    const std::size_t digits_start = *i;
    long long magnitude = 0;
    while (*i < text.size() && is_digit(text[*i])) {
      magnitude = magnitude * 10 + (text[*i] - '0');
      if (magnitude > Sheet::kMaxRows) {
        return false;  // saturate well before overflow; out of grid either way
      }
      ++*i;
    }
    if (*i == digits_start || *i >= text.size() || text[*i] != letters.close) {
      return false;
    }
    ++*i;
    out->relative = true;
    out->value = negative ? -magnitude : magnitude;
    return true;
  }
  // A bare `R` / `C` is the current row / column: a zero relative offset.
  const std::size_t digits_start = *i;
  long long absolute = 0;
  while (*i < text.size() && is_digit(text[*i])) {
    // Saturated past the grid: any longer run is a name all the same.
    absolute = std::min<long long>(absolute * 10 + (text[*i] - '0'), Sheet::kMaxRows + 1LL);
    ++*i;
  }
  if (*i == digits_start) {
    out->relative = true;
    out->value = 0;
    return true;
  }
  out->value = absolute;
  return true;
}

bool absolute_in_grid(const R1C1EndpointAxis& axis, std::uint32_t max) {
  return !axis.present || axis.relative || (axis.value >= 1 && axis.value <= static_cast<long long>(max));
}

}  // namespace

R1C1Endpoint parse_r1c1_endpoint_axes(std::string_view text, std::size_t start, const R1C1Letters& letters) {
  R1C1Endpoint out;
  std::size_t i = start;
  const bool read = read_r1c1_axis(text, &i, letters, letters.row, &out.row, &out.bracketed) &&
                    read_r1c1_axis(text, &i, letters, letters.col, &out.col, &out.bracketed);
  out.end = i;
  if (!read || (!out.row.present && !out.col.present)) {
    out.kind = R1C1EndpointKind::kMalformed;
    return out;
  }
  if (absolute_in_grid(out.row, Sheet::kMaxRows) && absolute_in_grid(out.col, Sheet::kMaxCols)) {
    out.kind = R1C1EndpointKind::kReference;
  } else {
    out.kind = out.bracketed ? R1C1EndpointKind::kMalformed : R1C1EndpointKind::kName;
  }
  return out;
}

bool resolve_r1c1_axis(const R1C1EndpointAxis& axis, bool base_present, std::uint32_t base, std::uint32_t max,
                       std::uint32_t* out) {
  if (axis.relative && !base_present) {
    return false;
  }
  if (axis.relative) {
    const long long span = static_cast<long long>(max);
    *out = static_cast<std::uint32_t>(((static_cast<long long>(base) + axis.value) % span + span) % span);
    return true;
  }
  const long long resolved = axis.value - 1;
  if (resolved < 0 || resolved >= static_cast<long long>(max)) {
    return false;
  }
  *out = static_cast<std::uint32_t>(resolved);
  return true;
}

A1Parse parse_r1c1_ref(std::string_view text, const R1C1Base& base) {
  const LocaleFacts& facts = locale_facts(current_eval_profile());
  A1Parse out;
  if (text.empty()) {
    return out;
  }
  std::size_t i = 0;
  if (parse_sheet_qualifier(text, &i, &out.sheet) == SheetQualifier::kMalformed) {
    return out;
  }

  // An absolute axis outside the grid reads as a name, which is no reference here.
  const R1C1Letters letters{facts.r1c1_row, facts.r1c1_col, facts.r1c1_open, facts.r1c1_close};
  const R1C1Endpoint first = parse_r1c1_endpoint_axes(text, i, letters);
  if (first.kind != R1C1EndpointKind::kReference) {
    return out;
  }
  i = first.end;
  const R1C1EndpointAxis& row1 = first.row;
  const R1C1EndpointAxis& col1 = first.col;
  R1C1EndpointAxis row2 = row1;
  R1C1EndpointAxis col2 = col1;
  bool is_range = false;
  if (i < text.size() && text[i] == ':') {
    const R1C1Endpoint second = parse_r1c1_endpoint_axes(text, i + 1, letters);
    if (second.kind != R1C1EndpointKind::kReference) {
      return out;
    }
    i = second.end;
    row2 = second.row;
    col2 = second.col;
    // A range whose endpoints name different axes (`R2:R3C4`) has no
    // rectangle; Excel rejects it rather than guessing the missing bound.
    if (row1.present != row2.present || col1.present != col2.present) {
      return out;
    }
    is_range = true;
  }
  if (i != text.size()) {
    return out;  // trailing garbage
  }

  // An endpoint that names only one axis is unbounded along the other, the
  // same shape `5:5` and `D:D` take on the A1 side.
  if (!col1.present) {
    if (!resolve_r1c1_axis(row1, base.present, base.row, Sheet::kMaxRows, &out.row) ||
        !resolve_r1c1_axis(row2, base.present, base.row, Sheet::kMaxRows, &out.row2)) {
      return out;
    }
    out.col = 0;
    out.col2 = Sheet::kMaxCols - 1U;
    out.is_full_row = true;
    out.is_range = true;
    out.valid = true;
    return out;
  }
  if (!row1.present) {
    if (!resolve_r1c1_axis(col1, base.present, base.col, Sheet::kMaxCols, &out.col) ||
        !resolve_r1c1_axis(col2, base.present, base.col, Sheet::kMaxCols, &out.col2)) {
      return out;
    }
    out.row = 0;
    out.row2 = Sheet::kMaxRows - 1U;
    out.is_full_col = true;
    out.is_range = true;
    out.valid = true;
    return out;
  }

  if (!resolve_r1c1_axis(row1, base.present, base.row, Sheet::kMaxRows, &out.row) ||
      !resolve_r1c1_axis(col1, base.present, base.col, Sheet::kMaxCols, &out.col) ||
      !resolve_r1c1_axis(row2, base.present, base.row, Sheet::kMaxRows, &out.row2) ||
      !resolve_r1c1_axis(col2, base.present, base.col, Sheet::kMaxCols, &out.col2)) {
    return out;
  }
  out.is_range = is_range;
  out.valid = true;
  return out;
}

}  // namespace refs_internal
}  // namespace eval
}  // namespace formulon
