// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of the structured-reference resolver. See
// `eval/structured_ref.h` for the public contract.

#include "eval/structured_ref.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/cell_parser.h"
#include "io/tables_reader.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/strings.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {

namespace {

/// Trims ASCII spaces / tabs from both ends of `s`. Excel structured
/// references frequently emit a single space between specifier
/// fragments (`Table[#This Row]`) but never line breaks; we strip
/// ASCII spaces only and leave UTF-8 bytes in column names intact.
std::string_view ascii_trim(std::string_view s) noexcept {
  std::size_t start = 0;
  while (start < s.size() && (s[start] == ' ' || s[start] == '\t')) {
    ++start;
  }
  std::size_t end = s.size();
  while (end > start && (s[end - 1] == ' ' || s[end - 1] == '\t')) {
    --end;
  }
  return s.substr(start, end - start);
}

/// Maps a `#`-prefixed specifier name (case-insensitive, ASCII-fold) to
/// its bitmask. Unknown specifiers return 0 so the caller can flag them as
/// `#NAME?`. `#This Row` is accepted with one space between `This` and
/// `Row`; arbitrary whitespace inside the specifier name would imply a
/// bracket payload not produced by Excel and is rejected.
std::uint8_t classify_hash_specifier(std::string_view name) noexcept {
  if (strings::case_insensitive_eq(name, "#All")) {
    return StructuredRefSpecifiers::kAll;
  }
  if (strings::case_insensitive_eq(name, "#Data")) {
    return StructuredRefSpecifiers::kData;
  }
  if (strings::case_insensitive_eq(name, "#Headers")) {
    return StructuredRefSpecifiers::kHeaders;
  }
  if (strings::case_insensitive_eq(name, "#Totals")) {
    return StructuredRefSpecifiers::kTotals;
  }
  if (strings::case_insensitive_eq(name, "#This Row")) {
    return StructuredRefSpecifiers::kThisRow;
  }
  return 0u;
}

/// Splits a bracket payload at top-level commas. Bracketed sub-fragments
/// (`[#Headers]`, `[Col Name]`) are recognised so commas inside them stay
/// attached to their parent fragment. Returns `false` on unmatched bracket.
bool split_top_level_commas(std::string_view payload, std::vector<std::string_view>* out) {
  out->clear();
  std::size_t depth = 0;
  std::size_t start = 0;
  for (std::size_t i = 0; i < payload.size(); ++i) {
    const char c = payload[i];
    if (c == '[') {
      ++depth;
    } else if (c == ']') {
      if (depth == 0) {
        return false;
      }
      --depth;
    } else if (c == ',' && depth == 0) {
      out->push_back(payload.substr(start, i - start));
      start = i + 1;
    }
  }
  if (depth != 0) {
    return false;
  }
  out->push_back(payload.substr(start));
  return true;
}

/// Strips a single leading `[` and trailing `]` from `s`. A non-bracketed
/// input is returned unchanged. Single-sided brackets are malformed.
bool strip_outer_brackets(std::string_view s, std::string_view* out) {
  s = ascii_trim(s);
  if (s.size() >= 2 && s.front() == '[' && s.back() == ']') {
    *out = s.substr(1, s.size() - 2);
    return true;
  }
  if (!s.empty() && (s.front() == '[' || s.back() == ']')) {
    return false;
  }
  *out = s;
  return true;
}

/// Splits a column-list fragment (already de-bracketed of any outer `[ ]`)
/// at the top-level `:` if present. Bracketed inner names are skipped so a
/// `[Col:Name]` quirk is not split by the operator. Writes the first (and
/// last, when present) column-name view into the selector slots; returns
/// `false` on malformed bracketing.
bool split_column_range(std::string_view text, StructuredRefSelector* sel) {
  text = ascii_trim(text);
  if (text.empty()) {
    return true;
  }
  std::size_t depth = 0;
  std::size_t colon_at = std::string_view::npos;
  for (std::size_t i = 0; i < text.size(); ++i) {
    const char c = text[i];
    if (c == '[') {
      ++depth;
    } else if (c == ']') {
      if (depth == 0) {
        return false;
      }
      --depth;
    } else if (c == ':' && depth == 0) {
      colon_at = i;
      break;
    }
  }
  std::string_view first;
  std::string_view last;
  if (colon_at == std::string_view::npos) {
    first = text;
  } else {
    first = ascii_trim(text.substr(0, colon_at));
    last = ascii_trim(text.substr(colon_at + 1));
  }
  std::string_view first_inner;
  if (!strip_outer_brackets(first, &first_inner)) {
    return false;
  }
  sel->column_first = first_inner;
  sel->column_first_set = true;
  if (colon_at != std::string_view::npos) {
    std::string_view last_inner;
    if (!strip_outer_brackets(last, &last_inner)) {
      return false;
    }
    sel->column_last = last_inner;
    sel->column_last_set = true;
  }
  return true;
}

/// Looks up a column by case-insensitive name. Returns a 0-based index into
/// `t.columns` on success, `-1u` on failure.
std::uint32_t find_column_index(const io::TableMetadata& t, std::string_view name) noexcept {
  for (std::size_t i = 0; i < t.columns.size(); ++i) {
    if (strings::case_insensitive_eq(t.columns[i].name, name)) {
      return static_cast<std::uint32_t>(i);
    }
  }
  return static_cast<std::uint32_t>(-1);
}

/// Parses the table's `ref` attribute (e.g. `"A1:C10"`) into a 0-based
/// rectangle. Returns `false` if `ref` is not a valid two-endpoint A1
/// range.
bool parse_table_ref(std::string_view ref, std::uint32_t* row_top, std::uint32_t* col_left, std::uint32_t* row_bot,
                     std::uint32_t* col_right) {
  const std::size_t colon = ref.find(':');
  if (colon == std::string_view::npos) {
    return false;
  }
  const std::string_view a = ref.substr(0, colon);
  const std::string_view b = ref.substr(colon + 1);
  auto a_rc = io::parse_a1(a);
  auto b_rc = io::parse_a1(b);
  if (!a_rc || !b_rc) {
    return false;
  }
  *row_top = std::min(a_rc.value().first, b_rc.value().first);
  *row_bot = std::max(a_rc.value().first, b_rc.value().first);
  *col_left = std::min(a_rc.value().second, b_rc.value().second);
  *col_right = std::max(a_rc.value().second, b_rc.value().second);
  return true;
}

}  // namespace

Expected<StructuredRefSelector, ErrorCode> parse_structured_ref_payload(std::string_view payload) {
  payload = ascii_trim(payload);
  StructuredRefSelector sel;
  sel.specifiers = 0u;

  if (payload.empty()) {
    sel.specifiers = StructuredRefSpecifiers::kData;
    return sel;
  }

  // Detect the row-implicit shorthand (`@...`). Excel emits `@` only as the
  // first character of the bracket payload; whitespace between `@` and the
  // column name is allowed.
  if (payload.front() == '@') {
    sel.specifiers |= StructuredRefSpecifiers::kThisRow;
    payload = ascii_trim(payload.substr(1));
    if (payload.empty()) {
      // Bare `Table[@]` is not a useful Excel construct; surface #NAME?.
      return ErrorCode::Name;
    }
  }

  std::vector<std::string_view> fragments;
  if (!split_top_level_commas(payload, &fragments)) {
    return ErrorCode::Name;
  }

  bool column_seen = false;
  for (std::string_view frag_raw : fragments) {
    std::string_view frag = ascii_trim(frag_raw);
    if (frag.empty()) {
      return ErrorCode::Name;
    }
    std::string_view inner;
    if (!strip_outer_brackets(frag, &inner)) {
      return ErrorCode::Name;
    }
    inner = ascii_trim(inner);
    if (!inner.empty() && inner.front() == '#') {
      const std::uint8_t bit = classify_hash_specifier(inner);
      if (bit == 0u) {
        return ErrorCode::Name;
      }
      sel.specifiers |= bit;
      continue;
    }
    if (column_seen) {
      return ErrorCode::Name;
    }
    column_seen = true;
    if (!split_column_range(frag, &sel)) {
      return ErrorCode::Name;
    }
  }

  // Default the area selector to `kData` when only a column was given.
  if ((sel.specifiers & ~StructuredRefSpecifiers::kThisRow) == 0u) {
    sel.specifiers |= StructuredRefSpecifiers::kData;
  }

  return sel;
}

Expected<StructuredRefRange, ErrorCode> resolve_structured_ref(const StructuredRefSelector& selector,
                                                               const Workbook& wb, std::uint32_t current_sheet_index,
                                                               std::uint32_t current_row) {
  (void)current_sheet_index;
  // Locate the table by case-insensitive name. OOXML's `name` is the
  // workbook-unique programmatic name; `display_name` is what users type
  // in formulas. Excel emits the two equal in practice but the spec
  // allows them to differ; accept either to match Excel.
  const io::TableMetadata* table = nullptr;
  for (const io::TableMetadata& t : wb.tables()) {
    if (strings::case_insensitive_eq(t.display_name, selector.table_name) ||
        strings::case_insensitive_eq(t.name, selector.table_name)) {
      table = &t;
      break;
    }
  }
  if (table == nullptr) {
    return ErrorCode::Name;
  }
  if (table->sheet_index >= wb.sheet_count()) {
    return ErrorCode::Ref;
  }

  std::uint32_t r_top = 0;
  std::uint32_t c_left = 0;
  std::uint32_t r_bot = 0;
  std::uint32_t c_right = 0;
  if (!parse_table_ref(table->ref, &r_top, &c_left, &r_bot, &c_right)) {
    return ErrorCode::Ref;
  }

  std::uint8_t spec = selector.specifiers;
  if (spec == 0u) {
    spec = StructuredRefSpecifiers::kData;
  }

  // Header row (when present) sits at `r_top`; data rows are below it; the
  // totals row (when present) sits at `r_bot`.
  const bool has_header = table->header_row;
  const bool has_totals = table->totals_row;
  const std::uint32_t header_row = r_top;
  const std::uint32_t totals_row = r_bot;
  const std::uint32_t data_top = has_header ? (r_top + 1u) : r_top;
  const std::uint32_t data_bot = has_totals ? (r_bot - 1u) : r_bot;
  const bool has_data = data_top <= data_bot;

  std::uint32_t out_row_top = 0;
  std::uint32_t out_row_bot = 0;
  bool row_set = false;

  auto extend = [&](std::uint32_t lo, std::uint32_t hi) {
    if (!row_set) {
      out_row_top = lo;
      out_row_bot = hi;
      row_set = true;
    } else {
      out_row_top = std::min(out_row_top, lo);
      out_row_bot = std::max(out_row_bot, hi);
    }
  };

  // `kThisRow` (the `@` shorthand) is mutually exclusive with the area
  // selectors at the row layer: when it is set the resolved row is the
  // single formula-cell row, regardless of any `kData` default the parser
  // added to record "no explicit area was specified".
  if ((spec & StructuredRefSpecifiers::kThisRow) != 0u) {
    if (!has_data || current_row < data_top || current_row > data_bot) {
      return ErrorCode::Value;
    }
    extend(current_row, current_row);
  } else {
    if ((spec & StructuredRefSpecifiers::kAll) != 0u) {
      extend(r_top, r_bot);
    }
    if ((spec & StructuredRefSpecifiers::kHeaders) != 0u) {
      if (!has_header) {
        return ErrorCode::Ref;
      }
      extend(header_row, header_row);
    }
    if ((spec & StructuredRefSpecifiers::kTotals) != 0u) {
      if (!has_totals) {
        return ErrorCode::Ref;
      }
      extend(totals_row, totals_row);
    }
    if ((spec & StructuredRefSpecifiers::kData) != 0u) {
      if (!has_data) {
        return ErrorCode::Ref;
      }
      extend(data_top, data_bot);
    }
  }
  if (!row_set) {
    return ErrorCode::Ref;
  }

  // Compute the column span.
  std::uint32_t out_col_left = c_left;
  std::uint32_t out_col_right = c_right;
  if (selector.column_first_set) {
    const std::uint32_t idx_first = find_column_index(*table, selector.column_first);
    if (idx_first == static_cast<std::uint32_t>(-1)) {
      return ErrorCode::Ref;
    }
    std::uint32_t idx_last = idx_first;
    if (selector.column_last_set) {
      const std::uint32_t i = find_column_index(*table, selector.column_last);
      if (i == static_cast<std::uint32_t>(-1)) {
        return ErrorCode::Ref;
      }
      idx_last = i;
    }
    const std::uint32_t lo = std::min(idx_first, idx_last);
    const std::uint32_t hi = std::max(idx_first, idx_last);
    out_col_left = c_left + lo;
    out_col_right = c_left + hi;
  }

  StructuredRefRange out;
  out.sheet_index = static_cast<std::uint32_t>(table->sheet_index);
  out.sheet_name = wb.sheet(table->sheet_index).name();
  out.row_first = out_row_top;
  out.col_first = out_col_left;
  out.row_last = out_row_bot;
  out.col_last = out_col_right;
  return out;
}

}  // namespace eval
}  // namespace formulon
