//
// Owned page handle behind `fm_cell_range_t`, shared by the entry points that
// enumerate cells with a resumable `row * 16384 + col` cursor
// (`fm_sheet_cells_in_range`, `fm_sheet_list_invalid_cells`).

#ifndef FORMULON_C_API_PARTS_CELL_RANGE_H_
#define FORMULON_C_API_PARTS_CELL_RANGE_H_

#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "sheet.h"
#include "utils/error.h"

namespace formulon {
namespace c_api {
namespace parts {

// Page size cap for a cell-range page, and the size a zero limit asks for.
constexpr std::uint32_t kMaxCellRangeLimit = 65536U;

// One past the largest cursor the grid can encode.
constexpr std::uint64_t kCellCursorLimit = static_cast<std::uint64_t>(Sheet::kMaxRows) * Sheet::kMaxCols;

inline std::uint32_t cell_range_page_size(std::uint32_t limit) noexcept {
  return (limit == 0U || limit > kMaxCellRangeLimit) ? kMaxCellRangeLimit : limit;
}

// Rejects a cursor past the last cell of the grid.
inline fm_status_t check_cell_cursor(std::uint64_t cursor, const char* api) {
  if (cursor >= kCellCursorLimit) {
    return set_binding_error(FormulonErrorCode::kInvalidArgument, (std::string(api) + ": cursor out of range").c_str(),
                             "cursor=" + std::to_string(cursor));
  }
  return 0;
}

}  // namespace parts
}  // namespace c_api
}  // namespace formulon

struct fm_cell_range {
  struct Entry {
    std::uint32_t row = 0;
    std::uint32_t col = 0;
    const char* formula = nullptr;
    fm_value_t value{};
  };
  std::vector<Entry> entries;
  // Owns every formula string and Text payload `entries` points at.
  formulon::c_api::parts::TextStore text;
  std::uint64_t next_cursor = formulon::Sheet::kCellCursorEnd;

  // Appends `cell`, copying its formula text and Text payload into the handle.
  void append(const formulon::Sheet::RangeCell& cell) {
    Entry entry;
    entry.row = cell.row;
    entry.col = cell.col;
    if (!cell.formula_text.empty()) {
      text.emplace_back(cell.formula_text);
      entry.formula = text.back().c_str();
    }
    formulon::c_api::parts::value_to_fm(cell.value, text, &entry.value);
    entries.push_back(entry);
  }
};

#endif  // FORMULON_C_API_PARTS_CELL_RANGE_H_
