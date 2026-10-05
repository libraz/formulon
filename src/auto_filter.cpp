#include "auto_filter.h"

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>

#include "sheet.h"
#include "table.h"
#include "utils/a1_column.h"
#include "utils/a1_ref.h"
#include "workbook.h"

namespace formulon {

namespace {

Error invalid(std::string message) {
  return make_error(FormulonErrorCode::kAutoFilterInvalid, std::move(message));
}

}  // namespace

bool AutoFilter::has_criteria() const noexcept {
  for (const FilterColumn& col : columns) {
    if (col.kind != FilterKind::kNone) {
      return true;
    }
  }
  return false;
}

std::optional<MergeRange> parse_a1_rectangle(std::string_view text) noexcept {
  const std::size_t colon = text.find(':');
  MergeRange rect;
  if (!a1::parse_a1_ref(text.substr(0, colon), &rect.first_row, &rect.first_col)) {
    return std::nullopt;
  }
  if (colon == std::string_view::npos) {
    rect.last_row = rect.first_row;
    rect.last_col = rect.first_col;
  } else if (!a1::parse_a1_ref(text.substr(colon + 1U), &rect.last_row, &rect.last_col)) {
    return std::nullopt;
  }
  if (rect.first_row > rect.last_row || rect.first_col > rect.last_col) {
    return std::nullopt;
  }
  return rect;
}

bool append_a1_rectangle(std::string& out, const MergeRange& rect) {
  const auto append_cell = [&out](std::uint32_t row, std::uint32_t col) {
    if (!a1::append_column_letters(out, col)) {
      return false;
    }
    out += std::to_string(static_cast<std::uint64_t>(row) + 1U);
    return true;
  };
  if (!append_cell(rect.first_row, rect.first_col)) {
    return false;
  }
  if (rect.first_row == rect.last_row && rect.first_col == rect.last_col) {
    return true;
  }
  out += ':';
  return append_cell(rect.last_row, rect.last_col);
}

Expected<void, Error> validate_auto_filter(const AutoFilter& filter) {
  if (filter.is_opaque()) {
    return invalid("autoFilter: model holds an unparsed fragment");
  }
  std::string probe;
  if (filter.range.first_row > filter.range.last_row || filter.range.first_col > filter.range.last_col ||
      !append_a1_rectangle(probe, filter.range) || !parse_a1_rectangle(probe)) {
    return invalid("autoFilter: range is not an in-grid rectangle");
  }
  const std::uint64_t width = static_cast<std::uint64_t>(filter.range.last_col) - filter.range.first_col + 1U;
  for (std::size_t i = 0; i < filter.columns.size(); ++i) {
    const FilterColumn& col = filter.columns[i];
    if (col.col_id >= width || (i > 0U && col.col_id <= filter.columns[i - 1U].col_id)) {
      return invalid("autoFilter: column ids must be ascending and inside the range");
    }
    if (col.kind == FilterKind::kCustom && (col.custom.filters.empty() || col.custom.filters.size() > 2U)) {
      return invalid("autoFilter: a custom filter holds one or two conditions");
    }
    for (const DateGroupItem& item : col.values.date_groups) {
      const auto level = static_cast<std::size_t>(item.grouping);
      if (level > static_cast<std::size_t>(DateTimeGrouping::kSecond) || item.year > 9999U ||
          (level >= 1U && (item.month < 1U || item.month > 12U)) ||
          (level >= 2U && (item.day < 1U || item.day > 31U)) || item.hour > 23U || item.minute > 59U ||
          item.second > 59U) {
        return invalid("autoFilter: date group item fields do not fit its grouping");
      }
    }
  }
  return Expected<void, Error>::Ok();
}

bool sheet_has_filter_criteria(const Workbook* wb, const Sheet& sheet) noexcept {
  if (const AutoFilter* own = sheet.auto_filter(); own != nullptr && own->has_criteria()) {
    return true;
  }
  if (wb == nullptr) {
    return false;
  }
  for (const TableMetadata& table : wb->tables()) {
    const AutoFilter* f = table.auto_filter_xml.get();
    if (f == nullptr || !f->has_criteria() || table.sheet_index >= wb->sheet_count()) {
      continue;
    }
    if (&wb->sheet(table.sheet_index) == &sheet) {
      return true;
    }
  }
  return false;
}

}  // namespace formulon
