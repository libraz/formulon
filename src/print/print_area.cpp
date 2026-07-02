// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.

#include "print/print_area.h"

#include <algorithm>
#include <cstddef>
#include <string>
#include <string_view>

#include "cell.h"
#include "io/a1_ref.h"
#include "io/defined_names.h"
#include "sheet.h"
#include "workbook.h"

namespace formulon {
namespace print {
namespace {

// OOXML built-in defined-name identifiers for the print area and titles.
constexpr std::string_view kPrintAreaName = "_xlnm.Print_Area";
constexpr std::string_view kPrintTitlesName = "_xlnm.Print_Titles";

constexpr char kAreaSeparator = ',';
constexpr char kRangeSeparator = ':';
constexpr char kSheetSeparator = '!';
constexpr char kAnchorChar = '$';

/// Removes every `$` anchor from `text`.
std::string StripAnchors(std::string_view text) {
  std::string out;
  out.reserve(text.size());
  for (const char c : text) {
    if (c != kAnchorChar) {
      out.push_back(c);
    }
  }
  return out;
}

/// Drops a leading `Sheet!` qualifier, if present.
///
/// Excel stores print-area formulas fully qualified (`Sheet1!$A$1:$H$80`)
/// and may quote sheet names containing spaces (`'My Sheet'!$A$1`). The
/// last unquoted `!` is the qualifier boundary; anything before it is the
/// sheet name, which this module does not need.
std::string_view StripSheetQualifier(std::string_view token) {
  bool in_quote = false;
  std::size_t boundary = std::string_view::npos;
  for (std::size_t i = 0; i < token.size(); ++i) {
    const char c = token[i];
    if (c == '\'') {
      in_quote = !in_quote;
    } else if (c == kSheetSeparator && !in_quote) {
      boundary = i;
    }
  }
  if (boundary == std::string_view::npos) {
    return token;
  }
  return token.substr(boundary + 1);
}

/// Finds the next unquoted `,` area separator in `body` at or after
/// `start`, skipping over any `'...'`-quoted sheet-name segment that may
/// itself contain a literal comma (e.g. `'Sheet,1'!A1:B2,C3:D4` has two
/// print areas, not three). Mirrors `StripSheetQualifier`'s quote
/// tracking: each `'` toggles the quote state, so Excel's `''`-escaped
/// literal apostrophe inside a quoted name round-trips back to the same
/// state instead of prematurely closing it. Returns
/// `std::string_view::npos` when no unquoted separator remains.
std::size_t FindAreaSeparator(std::string_view body, std::size_t start) {
  bool in_quote = false;
  for (std::size_t i = start; i < body.size(); ++i) {
    const char c = body[i];
    if (c == '\'') {
      in_quote = !in_quote;
    } else if (c == kAreaSeparator && !in_quote) {
      return i;
    }
  }
  return std::string_view::npos;
}

/// Trims ASCII whitespace from both ends of `text`.
std::string_view Trim(std::string_view text) {
  std::size_t begin = 0;
  std::size_t end = text.size();
  while (begin < end && (text[begin] == ' ' || text[begin] == '\t')) {
    ++begin;
  }
  while (end > begin && (text[end - 1] == ' ' || text[end - 1] == '\t')) {
    --end;
  }
  return text.substr(begin, end - begin);
}

/// Parses one anchor-free A1 reference (e.g. "A1", "BC42") into 0-based
/// row/col. Returns false on any malformed input or trailing characters.
bool ParseCellRef(std::string_view text, std::uint32_t* out_row, std::uint32_t* out_col) {
  return io::parse_a1_ref(text, out_row, out_col);
}

/// Largest valid 0-based row / column index (Excel's grid ceiling). Used
/// to clamp synthesized whole-axis spans and to cap over-large endpoints
/// so a malformed area cannot reserve a runaway track vector.
constexpr std::uint32_t kMaxRowIndex = Sheet::kMaxRows - 1U;
constexpr std::uint32_t kMaxColIndex = Sheet::kMaxCols - 1U;

/// Recognises a whole-column (`A:D`) or whole-row (`1:50`) print-area
/// span and synthesizes the clamped rectangle: a column span covers every
/// row up to the sheet ceiling, a row span every column. Mirrors how
/// `ParseTitleToken` accepts the same two shapes for `Print_Titles`.
/// Returns false when the token is neither shape.
bool ParseWholeAxisToken(std::string_view lhs, std::string_view rhs, CellRange* out_range) {
  // Whole-column span: both endpoints are pure column letters.
  std::size_t pi = 0;
  std::size_t pj = 0;
  std::uint32_t c1 = 0;
  std::uint32_t c2 = 0;
  if (io::parse_column_letters(lhs, &pi, &c1) && pi == lhs.size() && io::parse_column_letters(rhs, &pj, &c2) &&
      pj == rhs.size()) {
    out_range->first_col = std::min(std::min(c1, c2) - 1U, kMaxColIndex);
    out_range->last_col = std::min(std::max(c1, c2) - 1U, kMaxColIndex);
    out_range->first_row = 0U;
    out_range->last_row = kMaxRowIndex;
    return true;
  }

  // Whole-row span: both endpoints are pure decimal row indices.
  pi = 0;
  pj = 0;
  std::uint32_t r1 = 0;
  std::uint32_t r2 = 0;
  if (io::parse_uint(lhs, &pi, &r1) && pi == lhs.size() && r1 != 0U && io::parse_uint(rhs, &pj, &r2) &&
      pj == rhs.size() && r2 != 0U) {
    out_range->first_row = std::min(std::min(r1, r2) - 1U, kMaxRowIndex);
    out_range->last_row = std::min(std::max(r1, r2) - 1U, kMaxRowIndex);
    out_range->first_col = 0U;
    out_range->last_col = kMaxColIndex;
    return true;
  }
  return false;
}

/// Parses one anchor-free A1 range token into a normalised `CellRange`.
///
/// Accepts a full `A1:H80` range, a degenerate single-cell `A1`, or a
/// whole-column (`A:D`) / whole-row (`1:50`) span. The token must already
/// have its sheet qualifier and `$` anchors removed. Endpoints are
/// clamped to Excel's grid ceiling so an over-large reference does not
/// reserve a runaway track vector downstream. Returns false when neither
/// endpoint parses.
bool ParseRangeToken(std::string_view token, CellRange* out_range) {
  const std::size_t colon = token.find(kRangeSeparator);
  if (colon == std::string_view::npos) {
    std::uint32_t row = 0;
    std::uint32_t col = 0;
    if (!ParseCellRef(token, &row, &col)) {
      return false;
    }
    *out_range = CellRange{std::min(row, kMaxRowIndex), std::min(col, kMaxColIndex), std::min(row, kMaxRowIndex),
                           std::min(col, kMaxColIndex)};
    return true;
  }

  const std::string_view lhs = token.substr(0, colon);
  const std::string_view rhs = token.substr(colon + 1);
  if (lhs.empty() || rhs.empty()) {
    return false;
  }

  // Whole-column / whole-row forms (`A:D`, `1:50`) before the A1 path:
  // `parse_a1_ref` rejects them for lacking a row/column component.
  if (ParseWholeAxisToken(lhs, rhs, out_range)) {
    return true;
  }

  std::uint32_t r1 = 0;
  std::uint32_t c1 = 0;
  std::uint32_t r2 = 0;
  std::uint32_t c2 = 0;
  if (!ParseCellRef(lhs, &r1, &c1) || !ParseCellRef(rhs, &r2, &c2)) {
    return false;
  }
  out_range->first_row = std::min(std::min(r1, r2), kMaxRowIndex);
  out_range->first_col = std::min(std::min(c1, c2), kMaxColIndex);
  out_range->last_row = std::min(std::max(r1, r2), kMaxRowIndex);
  out_range->last_col = std::min(std::max(c1, c2), kMaxColIndex);
  return true;
}

/// Locates a sheet-scoped built-in defined name.
///
/// Returns the matching `formula`, or `nullptr` when no defined name
/// matches `name` for `sheet_index`. Names are matched case-insensitively
/// only on the comparison Excel performs; the built-in identifiers are
/// stored verbatim, so a plain comparison suffices here.
const std::string* FindSheetScopedFormula(const Workbook& wb, std::uint32_t sheet_index, std::string_view name) {
  for (const io::DefinedName& dn : wb.defined_names()) {
    if (dn.local_sheet_id < 0) {
      continue;
    }
    if (static_cast<std::uint32_t>(dn.local_sheet_id) != sheet_index) {
      continue;
    }
    if (dn.name == name) {
      return &dn.formula;
    }
  }
  return nullptr;
}

/// Parses a whole-row or whole-column title span from one cleaned token.
///
/// `Print_Titles` encodes repeat rows as `1:1` (digits only) and repeat
/// columns as `A:A` (letters only) after sheet/anchor stripping. Writes
/// the parsed span into `out_rows` or `out_cols` and returns true; returns
/// false when the token is neither shape.
bool ParseTitleToken(std::string_view token, PrintTitles* out_titles) {
  const std::size_t colon = token.find(kRangeSeparator);
  if (colon == std::string_view::npos) {
    return false;
  }
  const std::string_view lhs = token.substr(0, colon);
  const std::string_view rhs = token.substr(colon + 1);
  if (lhs.empty() || rhs.empty()) {
    return false;
  }

  // Whole-row span: both endpoints are pure decimal row indices.
  if (lhs.front() >= '0' && lhs.front() <= '9') {
    std::size_t pi = 0;
    std::size_t pj = 0;
    std::uint32_t r1 = 0;
    std::uint32_t r2 = 0;
    if (!io::parse_uint(lhs, &pi, &r1) || pi != lhs.size()) {
      return false;
    }
    if (!io::parse_uint(rhs, &pj, &r2) || pj != rhs.size()) {
      return false;
    }
    if (r1 == 0U || r2 == 0U) {
      return false;
    }
    const std::uint32_t lo = std::min(r1, r2) - 1U;
    const std::uint32_t hi = std::max(r1, r2) - 1U;
    out_titles->repeat_rows = std::make_pair(lo, hi);
    return true;
  }

  // Whole-column span: both endpoints are pure column letters.
  std::size_t pi = 0;
  std::size_t pj = 0;
  std::uint32_t c1 = 0;
  std::uint32_t c2 = 0;
  if (!io::parse_column_letters(lhs, &pi, &c1) || pi != lhs.size()) {
    return false;
  }
  if (!io::parse_column_letters(rhs, &pj, &c2) || pj != rhs.size()) {
    return false;
  }
  const std::uint32_t lo = std::min(c1, c2) - 1U;
  const std::uint32_t hi = std::max(c1, c2) - 1U;
  out_titles->repeat_cols = std::make_pair(lo, hi);
  return true;
}

}  // namespace

Expected<std::vector<CellRange>, Error> resolve_print_area(const Workbook& wb, std::uint32_t sheet_index) {
  std::vector<CellRange> ranges;
  const std::string* formula = FindSheetScopedFormula(wb, sheet_index, kPrintAreaName);
  if (formula == nullptr) {
    // An absent print area is not an error: callers fall back to the
    // sheet's used range.
    return ranges;
  }

  const std::string_view body = Trim(*formula);
  if (body.empty()) {
    return make_error(FormulonErrorCode::kPrintInvalidArea, "Empty Print_Area formula",
                      "sheet_index=" + std::to_string(sheet_index));
  }

  std::size_t start = 0;
  while (start <= body.size()) {
    std::size_t comma = FindAreaSeparator(body, start);
    if (comma == std::string_view::npos) {
      comma = body.size();
    }
    const std::string_view raw = Trim(body.substr(start, comma - start));
    start = comma + 1;
    if (raw.empty()) {
      return make_error(FormulonErrorCode::kPrintInvalidArea, "Empty area in Print_Area formula",
                        "sheet_index=" + std::to_string(sheet_index));
    }
    const std::string cleaned = StripAnchors(StripSheetQualifier(raw));
    CellRange range;
    if (!ParseRangeToken(cleaned, &range)) {
      return make_error(FormulonErrorCode::kPrintInvalidArea, "Malformed Print_Area range",
                        "sheet_index=" + std::to_string(sheet_index) + " token=" + cleaned);
    }
    ranges.push_back(range);
  }
  return ranges;
}

Expected<PrintTitles, Error> resolve_print_titles(const Workbook& wb, std::uint32_t sheet_index) {
  PrintTitles titles;
  const std::string* formula = FindSheetScopedFormula(wb, sheet_index, kPrintTitlesName);
  if (formula == nullptr) {
    return titles;
  }

  const std::string_view body = Trim(*formula);
  if (body.empty()) {
    return make_error(FormulonErrorCode::kPrintInvalidArea, "Empty Print_Titles formula",
                      "sheet_index=" + std::to_string(sheet_index));
  }

  std::size_t start = 0;
  while (start <= body.size()) {
    std::size_t comma = FindAreaSeparator(body, start);
    if (comma == std::string_view::npos) {
      comma = body.size();
    }
    const std::string_view raw = Trim(body.substr(start, comma - start));
    start = comma + 1;
    if (raw.empty()) {
      return make_error(FormulonErrorCode::kPrintInvalidArea, "Empty token in Print_Titles formula",
                        "sheet_index=" + std::to_string(sheet_index));
    }
    const std::string cleaned = StripAnchors(StripSheetQualifier(raw));
    if (!ParseTitleToken(cleaned, &titles)) {
      return make_error(FormulonErrorCode::kPrintInvalidArea, "Malformed Print_Titles token",
                        "sheet_index=" + std::to_string(sheet_index) + " token=" + cleaned);
    }
  }
  return titles;
}

}  // namespace print
}  // namespace formulon
