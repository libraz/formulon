//
// Implementation of the OOXML writer's cell-reference formatters.

#include "io/ooxml/cell_ref_writer.h"

#include <cstdint>
#include <string>

#include "sheet.h"  // for MergeRange
#include "utils/a1_ref.h"
#include "utils/expected.h"

namespace formulon {
namespace io {

void AppendCellRefForRef(std::string& out, std::uint32_t row, std::uint32_t col) {
  FM_CHECK(a1::append_a1(out, row, col), "column is outside Excel's grid");
}

void AppendRangeRef(std::string& out, std::uint32_t first_row, std::uint32_t first_col, std::uint32_t last_row,
                    std::uint32_t last_col) {
  AppendCellRefForRef(out, first_row, first_col);
  if (first_row != last_row || first_col != last_col) {
    out.push_back(':');
    AppendCellRefForRef(out, last_row, last_col);
  }
}

void AppendRangeRef(std::string& out, const MergeRange& r) {
  AppendRangeRef(out, r.first_row, r.first_col, r.last_row, r.last_col);
}

}  // namespace io
}  // namespace formulon
