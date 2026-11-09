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

#include "sheet.h"
#include "sheet_name.h"
#include "table.h"
#include "utils/a1_ref.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {

namespace {

/// Looks up a column by case-insensitive name. Returns a 0-based index into
/// `t.columns` on success, `-1u` on failure.
std::uint32_t find_column_index(const TableMetadata& t, std::string_view name) noexcept {
  for (std::size_t i = 0; i < t.columns.size(); ++i) {
    if (collection_names::equal(t.columns[i].name, name)) {
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
  std::uint32_t a_row = 0;
  std::uint32_t a_col = 0;
  std::uint32_t b_row = 0;
  std::uint32_t b_col = 0;
  if (!a1::parse_a1_ref(a, &a_row, &a_col) || !a1::parse_a1_ref(b, &b_row, &b_col)) {
    return false;
  }
  *row_top = std::min(a_row, b_row);
  *row_bot = std::max(a_row, b_row);
  *col_left = std::min(a_col, b_col);
  *col_right = std::max(a_col, b_col);
  return true;
}

}  // namespace

Expected<StructuredRefSelector, ErrorCode> parse_structured_ref_payload(std::string_view payload) {
  auto parsed = parser::parse_structured_ref_payload_text(payload);
  if (!parsed) {
    return ErrorCode::Name;
  }
  StructuredRefSelector sel;
  sel.payload = std::move(*parsed);
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
  const TableMetadata* table = nullptr;
  for (const TableMetadata& t : wb.tables()) {
    if (collection_names::equal(t.display_name, selector.table_name) ||
        collection_names::equal(t.name, selector.table_name)) {
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
  // A table needs at least distinct top and bottom rows. In particular, a
  // crafted one-row table marked as both header and totals would otherwise
  // compute `r_bot - 1` below and wrap to UINT32_MAX.
  if (r_bot <= r_top) {
    return ErrorCode::Ref;
  }

  std::uint8_t spec = selector.payload.specifiers;
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

  // `kThisRow` (the `@` shorthand) wins over the area selectors: the
  // resolved row is the single formula-cell row.
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
  const std::vector<std::string>& names = selector.payload.columns;
  if (!names.empty()) {
    const std::uint32_t idx_first = find_column_index(*table, names.front());
    if (idx_first == static_cast<std::uint32_t>(-1)) {
      return ErrorCode::Ref;
    }
    std::uint32_t idx_last = idx_first;
    if (names.size() > 1U) {
      const std::uint32_t i = find_column_index(*table, names.back());
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
