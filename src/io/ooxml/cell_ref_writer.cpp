//
// Implementation of the OOXML writer's cell-reference formatters.

#include "io/ooxml/cell_ref_writer.h"

#include <cstdint>
#include <string>

#include "sheet.h"  // for MergeRange
#include "utils/a1_column.h"
#include "utils/expected.h"

namespace formulon {
namespace io {

void AppendColumnLettersForRef(std::string& out, std::uint32_t col) {
  FM_CHECK(a1::append_column_letters(out, col), "column is outside Excel's grid");
}

void AppendCellRefForRef(std::string& out, std::uint32_t row, std::uint32_t col) {
  AppendColumnLettersForRef(out, col);
  out.append(std::to_string(row + 1));
}

void AppendRangeRef(std::string& out, const MergeRange& r) {
  AppendCellRefForRef(out, r.first_row, r.first_col);
  if (r.first_row != r.last_row || r.first_col != r.last_col) {
    out.push_back(':');
    AppendCellRefForRef(out, r.last_row, r.last_col);
  }
}

}  // namespace io
}  // namespace formulon
