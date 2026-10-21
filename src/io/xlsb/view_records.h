// BrtPane / BrtSel codec: a worksheet view's frozen panes and per-pane
// selections to and from `SheetView`.

#ifndef FORMULON_IO_XLSB_VIEW_RECORDS_H_
#define FORMULON_IO_XLSB_VIEW_RECORDS_H_

#include <cstdint>
#include <vector>

#include "io/xlsb/record.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

struct SheetView;

namespace io {
namespace xlsb {

/// Decodes one BrtPane into `view`'s frozen-pane fields. False for a pane
/// that is not frozen (a split, which the model does not carry); an error
/// when a frozen pane's counts are not whole in-grid row/column numbers.
Expected<bool, Error> decode_pane(const XlsbRecord& rec, SheetView& view);

/// Decodes one BrtSel and appends it to `view.selections`; the top-left
/// pane's default A1 selection appends nothing, as Excel's XLSX omits it.
/// False, appending nothing, when the record is truncated, inconsistent or
/// leaves the grid.
bool decode_selection(const XlsbRecord& rec, SheetView& view);

/// Emits a frozen view's BrtPane, then the top-left pane's BrtSel (A1 when
/// none is stored) and, in a frozen view, every other pane's.
void emit_pane_and_selections(std::vector<std::uint8_t>& dst, const SheetView& view);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_VIEW_RECORDS_H_
