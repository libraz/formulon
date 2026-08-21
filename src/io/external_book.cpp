
#include "io/external_book.h"

#include <cstddef>
#include <cstdint>
#include <string_view>

#include "utils/strings.h"
#include "value.h"

namespace formulon {
namespace io {

std::uint32_t ExternalBook::sheet_index(std::string_view sheet) const noexcept {
  for (std::size_t i = 0; i < sheet_names.size(); ++i) {
    if (sheet_names[i] == sheet) {
      return static_cast<std::uint32_t>(i);
    }
  }
  return kNoSheet;
}

const ExternalBookName* ExternalBook::find_name(std::string_view name) const noexcept {
  for (const ExternalBookName& entry : names) {
    if (strings::case_insensitive_eq(entry.name, name)) {
      return &entry;
    }
  }
  return nullptr;
}

Value ExternalBook::cached_cell(std::uint32_t sheet, std::uint32_t row, std::uint32_t col) const noexcept {
  const auto found = cells.find(cell_key(sheet, row, col));
  if (found == cells.end()) {
    return Value::number(0.0);
  }
  return found->second.resolved();
}

}  // namespace io
}  // namespace formulon
