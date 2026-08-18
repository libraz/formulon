//
// Implementation of the shared spill-anchor projection. See
// `spill_anchor.h` for the error contract.

#include "eval/spill_anchor.h"

#include <cstddef>

#include "eval/array_alloc.h"
#include "eval/eval_context.h"
#include "sheet.h"
#include "workbook.h"

namespace formulon {
namespace eval {

ArrayValue* project_spill_at_anchor(std::string_view sheet, std::uint32_t row, std::uint32_t col, Arena& arena,
                                    const EvalContext& ctx, ErrorCode* out_err) {
  const Sheet* current = ctx.current_sheet();
  if (current == nullptr) {
    *out_err = ErrorCode::Name;
    return nullptr;
  }
  const Sheet* target = current;
  if (!sheet.empty()) {
    const Workbook* wb = ctx.workbook();
    if (wb == nullptr) {
      *out_err = ErrorCode::Ref;
      return nullptr;
    }
    target = wb->sheet_by_name(sheet);
    if (target == nullptr) {
      *out_err = ErrorCode::Ref;
      return nullptr;
    }
  }
  if (row >= Sheet::kMaxRows || col >= Sheet::kMaxCols) {
    *out_err = ErrorCode::Ref;
    return nullptr;
  }
  const SpillRegion* region = target->spill_region_at_anchor(row, col);
  if (region == nullptr) {
    *out_err = ErrorCode::Ref;
    return nullptr;
  }
  Value* buffer = nullptr;
  ArrayValue* arr = allocate_array_value(region->rows, region->cols, arena, buffer, kMaxDerivedArrayCells);
  if (arr == nullptr) {
    *out_err = ErrorCode::Num;
    return nullptr;
  }
  const std::size_t cells = static_cast<std::size_t>(region->rows) * static_cast<std::size_t>(region->cols);
  for (std::size_t i = 0; i < cells; ++i) {
    buffer[i] = region->cells[i];
  }
  return arr;
}

}  // namespace eval
}  // namespace formulon
