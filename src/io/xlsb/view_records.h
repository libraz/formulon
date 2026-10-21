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

/// Decodes one BrtPane into `view`'s freeze or split fields. False for an
/// unfrozen pane with no valid split; an error when a frozen pane's counts
/// are not in-grid whole numbers.
Expected<bool, Error> decode_pane(const XlsbRecord& rec, SheetView& view);

/// Decodes one BrtSel and appends it to `view.selections`, an A1 active cell
/// or sqref left empty and the top-left pane's A1 selection not appended,
/// as Excel's XLSX spells them. False, appending nothing, when the record is
/// truncated, inconsistent or leaves the grid.
bool decode_selection(const XlsbRecord& rec, SheetView& view);

/// Emits the BrtPane of a frozen or split view, then one BrtSel per pane
/// the view has, as `resolve_view_selections` orders and fills them.
void emit_pane_and_selections(std::vector<std::uint8_t>& dst, const SheetView& view);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_VIEW_RECORDS_H_
