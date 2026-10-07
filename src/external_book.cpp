
#include "external_book.h"

#include <cstddef>
#include <cstdint>
#include <string_view>

#include "utils/index_sort.h"
#include "utils/strings.h"
#include "value.h"

namespace formulon {

std::vector<std::uint64_t> ExternalBook::sorted_cell_keys() const {
  std::vector<std::uint64_t> keys;
  keys.reserve(cells.size());
  for (const auto& entry : cells) {
    keys.push_back(entry.first);
  }
  sort_by_index(keys, [](std::uint64_t lhs, std::uint64_t rhs) { return lhs < rhs; });
  return keys;
}

std::uint32_t ExternalBook::sheet_index(std::string_view sheet) const noexcept {
  for (std::size_t i = 0; i < sheet_names.size(); ++i) {
    if (strings::case_insensitive_eq(sheet_names[i], sheet)) {
      return static_cast<std::uint32_t>(i);
    }
  }
  return kNoSheet;
}

const ExternalBookName* ExternalBook::find_name(std::string_view name, std::uint32_t scope_sheet) const noexcept {
  for (const ExternalBookName& entry : names) {
    if (entry.scope_sheet == scope_sheet && strings::case_insensitive_eq(entry.name, name)) {
      return &entry;
    }
  }
  return nullptr;
}

bool ExternalBook::cached_extent(std::uint32_t sheet, std::uint32_t* last_row, std::uint32_t* last_col) const noexcept {
  bool any = false;
  for (const auto& entry : cells) {
    if ((entry.first >> kSheetShift) != sheet) {
      continue;
    }
    const auto row = static_cast<std::uint32_t>((entry.first >> kRowShift) & kRowMask);
    const auto col = static_cast<std::uint32_t>(entry.first & kColMask);
    *last_row = any && *last_row > row ? *last_row : row;
    *last_col = any && *last_col > col ? *last_col : col;
    any = true;
  }
  return any;
}

Value ExternalBook::cached_cell(std::uint32_t sheet, std::uint32_t row, std::uint32_t col) const noexcept {
  const auto found = cells.find(cell_key(sheet, row, col));
  if (found == cells.end()) {
    return Value::number(0.0);
  }
  return found->second.resolved();
}

}  // namespace formulon
